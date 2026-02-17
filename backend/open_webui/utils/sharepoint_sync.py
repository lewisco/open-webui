import asyncio
import io
import logging
import mimetypes
import time
import uuid
from dataclasses import dataclass
from typing import Optional

from open_webui.models.files import FileForm, Files
from open_webui.models.knowledge import Knowledges
from open_webui.models.sharepoint import SharePoints
from open_webui.retrieval.loaders.sharepoint import SharePointGraphClient
from open_webui.storage.provider import Storage

log = logging.getLogger(__name__)

SYNC_USER_ID = "sharepoint-sync"


@dataclass
class _SyncUser:
    """Lightweight user object for internal process_file calls.
    Uses 'user' role (not 'admin') — it only needs access to files
    it created itself (matched via SYNC_USER_ID).
    """
    id: str = SYNC_USER_ID
    role: str = "user"
    name: str = "SharePoint Sync"
    email: str = "sharepoint-sync@system"


@dataclass
class _SyncRequest:
    """Lightweight request wrapper for internal process_file calls."""
    app: object = None


def _cleanup_vector_entries(kb_id: str, owui_file_id: str):
    """Remove embeddings for a file from the KB vector collection."""
    try:
        from open_webui.retrieval.vector.factory import VECTOR_DB_CLIENT

        VECTOR_DB_CLIENT.delete(
            collection_name=kb_id,
            filter={"file_id": owui_file_id},
        )
    except Exception as e:
        log.debug(f"Vector cleanup for file {owui_file_id}: {e}")


def _is_within_selection(item_path: str, selected_items: list) -> bool:
    """
    Check if an item path falls within the selected scope.
    If selected_items is empty/None, everything is in scope.
    """
    if not selected_items:
        return True

    item_path_lower = item_path.lower().rstrip("/")

    for sel in selected_items:
        sel_path = sel.get("path", "").lower().rstrip("/")
        sel_type = sel.get("type", "")

        if sel_type == "folder":
            # Item is within this folder if path starts with folder path
            if item_path_lower == sel_path or item_path_lower.startswith(
                sel_path + "/"
            ):
                return True
        elif sel_type == "file":
            if item_path_lower == sel_path:
                return True

    return False


def _upload_to_storage(content: bytes, filename: str, file_id: str) -> str:
    """Upload file content to storage and return the file path."""
    storage_filename = f"{file_id}_{filename}"
    _contents, file_path = Storage.upload_file(
        io.BytesIO(content),
        storage_filename,
        {
            "OpenWebUI-User-Id": SYNC_USER_ID,
            "OpenWebUI-File-Id": file_id,
        },
    )
    return file_path


def _process_and_embed_file(app, file_id: str, kb_id: str):
    """
    Process a file for embedding — two-step:
    1. Extract content into per-file collection (collection_name=None)
    2. Copy embeddings to KB collection (collection_name=kb_id)
    """
    from open_webui.internal.db import get_db
    from open_webui.routers.retrieval import process_file, ProcessFileForm

    sync_request = _SyncRequest(app=app)
    sync_user = _SyncUser()

    # Step 1: Extract content into per-file collection
    with get_db() as db:
        try:
            process_file(
                request=sync_request,
                form_data=ProcessFileForm(file_id=file_id, collection_name=None),
                user=sync_user,
                db=db,
            )
        except Exception as e:
            log.warning(f"Step 1 (extract) failed for file {file_id}: {e}")
            raise

    # Step 2: Copy embeddings to KB collection
    with get_db() as db:
        try:
            process_file(
                request=sync_request,
                form_data=ProcessFileForm(
                    file_id=file_id, collection_name=kb_id
                ),
                user=sync_user,
                db=db,
            )
        except Exception as e:
            log.warning(f"Step 2 (KB embed) failed for file {file_id}: {e}")
            raise


def sync_site(app, site_config) -> dict:
    """
    Sync a single SharePoint site.
    Returns stats dict: {added, updated, skipped, deleted, errors}
    """
    stats = {"added": 0, "updated": 0, "skipped": 0, "deleted": 0, "errors": 0}

    config = app.state.config
    allowed_extensions = getattr(config, "ALLOWED_FILE_EXTENSIONS", None)
    max_size = getattr(config, "FILE_MAX_SIZE", None)

    try:
        # Mark as syncing
        SharePoints.update_site(site_config.id, {"sync_status": "syncing", "sync_error": None})

        client = SharePointGraphClient(
            tenant_id=config.SHAREPOINT_TENANT_ID,
            client_id=config.SHAREPOINT_CLIENT_ID,
            client_secret=config.SHAREPOINT_CLIENT_SECRET,
        )

        # Get incremental changes
        delta_link = site_config.delta_link
        items, new_delta_link = client.get_folder_delta(
            site_config.drive_id, delta_link
        )

        log.info(
            f"SharePoint sync for '{site_config.site_name}': "
            f"{len(items)} delta items"
        )

        for item in items:
            try:
                item_id = item.get("id", "")
                item_name = item.get("name", "")
                item_path = SharePointGraphClient.get_item_path(item)
                file_id = None

                # Handle deleted items
                if item.get("deleted"):
                    tracked = SharePoints.get_file_by_sp_item_id(
                        site_config.id, item_id
                    )
                    if tracked and tracked.owui_file_id:
                        _cleanup_vector_entries(site_config.kb_id, tracked.owui_file_id)

                        # Delete OWUI file
                        Files.delete_file_by_id(tracked.owui_file_id)

                        # Remove tracking record
                        SharePoints.delete_file_by_sp_item_id(
                            site_config.id, item_id
                        )
                        stats["deleted"] += 1
                    continue

                # Skip folders
                if "folder" in item:
                    continue

                # Skip non-file items (packages, etc.)
                if "file" not in item:
                    continue

                # Selection filter
                if not _is_within_selection(
                    item_path, site_config.selected_items
                ):
                    continue

                # File type/size check
                item_size = item.get("size", 0)
                if not SharePointGraphClient.is_supported_file(
                    item_name, item_size, allowed_extensions, max_size
                ):
                    stats["skipped"] += 1
                    continue

                # Exclusion check
                tracked = SharePoints.get_file_by_sp_item_id(
                    site_config.id, item_id
                )
                if tracked and tracked.excluded:
                    stats["skipped"] += 1
                    continue

                # Check if file changed (via eTag)
                item_etag = item.get("eTag", "")
                if tracked and tracked.sp_etag == item_etag and tracked.sync_status == "synced":
                    # No change
                    continue

                # Download file
                content = client.download_file(site_config.drive_id, item_id)

                # Upload to OWUI storage
                file_id = str(uuid.uuid4())
                file_path = _upload_to_storage(content, item_name, file_id)

                # Guess content type
                content_type = (
                    mimetypes.guess_type(item_name)[0] or "application/octet-stream"
                )

                # Create OWUI file record
                file_record = Files.insert_new_file(
                    user_id=SYNC_USER_ID,
                    form_data=FileForm(
                        id=file_id,
                        filename=item_name,
                        path=file_path,
                        data={"status": "pending"},
                        meta={
                            "name": item_name,
                            "content_type": content_type,
                            "size": len(content),
                            "data": {
                                "sharepoint_site_id": site_config.id,
                                "sharepoint_item_id": item_id,
                            },
                        },
                    ),
                )

                if not file_record:
                    log.error(f"Failed to create file record for {item_name}")
                    stats["errors"] += 1
                    continue

                # Add file to KB
                Knowledges.add_file_to_knowledge_by_id(
                    knowledge_id=site_config.kb_id,
                    file_id=file_id,
                    user_id=SYNC_USER_ID,
                )

                # Process and embed
                try:
                    _process_and_embed_file(app, file_id, site_config.kb_id)
                except Exception as e:
                    log.error(f"Embedding failed for {item_name}: {e}")
                    # File is created but embedding failed — mark as error
                    SharePoints.upsert_file(
                        site_config.id,
                        item_id,
                        {
                            "owui_file_id": file_id,
                            "filename": item_name,
                            "sp_item_path": item_path,
                            "sp_etag": item_etag,
                            "sp_last_modified": item.get("lastModifiedDateTime", ""),
                            "sync_status": "error",
                            "sync_error": str(e)[:500],
                        },
                    )
                    stats["errors"] += 1
                    continue

                # Get permissions if site uses filter mode
                allowed_users = None
                allowed_groups = None
                if site_config.sync_mode == "filter":
                    try:
                        perms = client.get_item_permissions(
                            site_config.drive_id, item_id
                        )
                        allowed_users = perms.get("users", [])
                        allowed_groups = perms.get("groups", [])
                    except Exception as e:
                        log.warning(
                            f"Failed to get permissions for {item_name}: {e}"
                        )
                        # In filter mode, never store a file without permissions —
                        # it would be readable by everyone.
                        SharePoints.upsert_file(
                            site_config.id,
                            item_id,
                            {
                                "owui_file_id": file_id,
                                "filename": item_name,
                                "sp_item_path": item_path,
                                "sp_etag": item_etag,
                                "sp_last_modified": item.get("lastModifiedDateTime", ""),
                                "sync_status": "error",
                                "sync_error": f"Permission fetch failed: {str(e)[:400]}",
                            },
                        )
                        stats["errors"] += 1
                        continue

                # Upsert tracking record
                SharePoints.upsert_file(
                    site_config.id,
                    item_id,
                    {
                        "owui_file_id": file_id,
                        "filename": item_name,
                        "sp_item_path": item_path,
                        "sp_etag": item_etag,
                        "sp_last_modified": item.get("lastModifiedDateTime", ""),
                        "allowed_users": allowed_users,
                        "allowed_groups": allowed_groups,
                        "sync_status": "synced",
                        "sync_error": None,
                    },
                )

                if tracked:
                    # Updated existing file — clean up old OWUI file
                    if tracked.owui_file_id and tracked.owui_file_id != file_id:
                        _cleanup_vector_entries(site_config.kb_id, tracked.owui_file_id)
                        Files.delete_file_by_id(tracked.owui_file_id)
                    stats["updated"] += 1
                else:
                    stats["added"] += 1

            except Exception as e:
                log.exception(f"Error processing item {item.get('name', '?')}: {e}")
                # Track the error so it's visible in the UI — same upsert
                # pattern as the embedding and permission-fetch error handlers.
                if item_id and "file" in item:
                    SharePoints.upsert_file(
                        site_config.id,
                        item_id,
                        {
                            "owui_file_id": file_id,
                            "filename": item_name,
                            "sp_item_path": item_path,
                            "sp_etag": item.get("eTag", ""),
                            "sp_last_modified": item.get("lastModifiedDateTime", ""),
                            "sync_status": "error",
                            "sync_error": str(e)[:500],
                        },
                    )
                stats["errors"] += 1

        # Update site metadata
        SharePoints.update_site(
            site_config.id,
            {
                "delta_link": new_delta_link,
                "last_sync_at": int(time.time()),
                "sync_status": "idle",
                "sync_error": None,
            },
        )

    except Exception as e:
        log.exception(f"Sync failed for site {site_config.id}: {e}")
        SharePoints.update_site(
            site_config.id,
            {
                "sync_status": "error",
                "sync_error": str(e)[:500],
            },
        )
        raise

    return stats


def retry_error_files(app, site_config) -> dict:
    """
    Retry embedding for files that previously failed.
    Returns stats dict: {retried, succeeded, failed}
    """
    stats = {"retried": 0, "succeeded": 0, "failed": 0}

    error_files = SharePoints.get_error_files_by_site(site_config.id)
    if not error_files:
        return stats

    log.info(
        f"Retrying {len(error_files)} error files for '{site_config.site_name}'"
    )

    for sp_file in error_files:
        stats["retried"] += 1

        if not sp_file.owui_file_id:
            log.warning(f"Error file {sp_file.filename} has no OWUI file ID, skipping")
            stats["failed"] += 1
            continue

        try:
            _process_and_embed_file(app, sp_file.owui_file_id, site_config.kb_id)

            # Success — update tracking record
            SharePoints.upsert_file(
                site_config.id,
                sp_file.sp_item_id,
                {
                    "sync_status": "synced",
                    "sync_error": None,
                },
            )
            stats["succeeded"] += 1
        except Exception as e:
            log.error(f"Retry failed for {sp_file.filename}: {e}")
            SharePoints.upsert_file(
                site_config.id,
                sp_file.sp_item_id,
                {
                    "sync_error": str(e)[:500],
                },
            )
            stats["failed"] += 1

    return stats


async def trigger_retry_errors(app, site_id: str) -> dict:
    """Retry failed files for a specific site."""
    site = SharePoints.get_site_by_id(site_id)
    if not site:
        raise ValueError(f"Site {site_id} not found")

    stats = await asyncio.to_thread(retry_error_files, app, site)
    return {"id": site.id, "name": site.site_name, "stats": stats}


async def trigger_sync(
    app,
    site_id: Optional[str] = None,
    force: bool = False,
    clear_exclusions: bool = False,
) -> dict:
    """
    Trigger SharePoint sync.
    Returns: {sites: [{id, name, stats}]}
    """
    results = []

    if site_id:
        sites = [SharePoints.get_site_by_id(site_id)]
        if not sites[0]:
            raise ValueError(f"Site {site_id} not found")
    else:
        sites = SharePoints.get_sites()

    for site in sites:
        if clear_exclusions:
            SharePoints.clear_exclusions_by_site(site.id)

        if force:
            SharePoints.update_site(site.id, {"delta_link": None})
            # Refresh after clearing delta_link
            site = SharePoints.get_site_by_id(site.id)

        try:
            stats = await asyncio.to_thread(sync_site, app, site)
            results.append(
                {"id": site.id, "name": site.site_name, "stats": stats}
            )
        except Exception as e:
            results.append(
                {
                    "id": site.id,
                    "name": site.site_name,
                    "error": str(e),
                }
            )

    return {"sites": results}


async def sharepoint_sync_periodic(app):
    """Periodic sync task — runs in the background."""
    log.info("SharePoint periodic sync task started")

    while True:
        try:
            config = app.state.config
            interval = getattr(config, "SHAREPOINT_SYNC_INTERVAL", 900)
            enabled = getattr(config, "ENABLE_SHAREPOINT_SYNC", False)

            if not enabled:
                await asyncio.sleep(60)  # Check again in a minute
                continue

            await asyncio.sleep(interval)

            # Check again after sleep in case config changed
            if not getattr(config, "ENABLE_SHAREPOINT_SYNC", False):
                continue

            log.info("Running periodic SharePoint sync")
            try:
                await trigger_sync(app)
            except Exception as e:
                log.error(f"Periodic sync error: {e}")

        except asyncio.CancelledError:
            log.info("SharePoint periodic sync task cancelled")
            break
        except Exception as e:
            log.exception(f"Unexpected error in periodic sync: {e}")
            await asyncio.sleep(60)
