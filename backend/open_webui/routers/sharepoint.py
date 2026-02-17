import logging
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, Request, status
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from sqlalchemy.orm import Session

from open_webui.constants import ERROR_MESSAGES
from open_webui.internal.db import get_session
from open_webui.models.files import Files
from open_webui.models.knowledge import KnowledgeForm, Knowledges
from open_webui.models.sharepoint import (
    SharePointConfigForm,
    SharePointFileModel,
    SharePointSiteForm,
    SharePointSiteModel,
    SharePointSiteUpdateForm,
    SharePointSyncForm,
    SharePoints,
)
from open_webui.retrieval.loaders.sharepoint import SharePointGraphClient
from open_webui.retrieval.vector.factory import VECTOR_DB_CLIENT
from open_webui.utils.auth import get_admin_user

log = logging.getLogger(__name__)
router = APIRouter()


############################
# Config
############################


class SharePointConfigResponse(BaseModel):
    ENABLE_SHAREPOINT_SYNC: bool
    SHAREPOINT_TENANT_ID: str
    SHAREPOINT_CLIENT_ID: str
    SHAREPOINT_CLIENT_SECRET: str
    SHAREPOINT_SYNC_INTERVAL: int


@router.get("/config", response_model=SharePointConfigResponse)
async def get_sharepoint_config(
    request: Request,
    user=Depends(get_admin_user),
):
    secret = request.app.state.config.SHAREPOINT_CLIENT_SECRET
    return SharePointConfigResponse(
        ENABLE_SHAREPOINT_SYNC=request.app.state.config.ENABLE_SHAREPOINT_SYNC,
        SHAREPOINT_TENANT_ID=request.app.state.config.SHAREPOINT_TENANT_ID,
        SHAREPOINT_CLIENT_ID=request.app.state.config.SHAREPOINT_CLIENT_ID,
        SHAREPOINT_CLIENT_SECRET="**********" if secret else "",
        SHAREPOINT_SYNC_INTERVAL=request.app.state.config.SHAREPOINT_SYNC_INTERVAL,
    )


@router.post("/config", response_model=SharePointConfigResponse)
async def update_sharepoint_config(
    request: Request,
    form_data: SharePointConfigForm,
    user=Depends(get_admin_user),
):
    config = request.app.state.config

    if form_data.ENABLE_SHAREPOINT_SYNC is not None:
        config.ENABLE_SHAREPOINT_SYNC = form_data.ENABLE_SHAREPOINT_SYNC
    if form_data.SHAREPOINT_TENANT_ID is not None:
        config.SHAREPOINT_TENANT_ID = form_data.SHAREPOINT_TENANT_ID
    if form_data.SHAREPOINT_CLIENT_ID is not None:
        config.SHAREPOINT_CLIENT_ID = form_data.SHAREPOINT_CLIENT_ID
    if form_data.SHAREPOINT_CLIENT_SECRET is not None:
        # Only update if not the masked placeholder
        if form_data.SHAREPOINT_CLIENT_SECRET != "**********":
            config.SHAREPOINT_CLIENT_SECRET = form_data.SHAREPOINT_CLIENT_SECRET
    if form_data.SHAREPOINT_SYNC_INTERVAL is not None:
        config.SHAREPOINT_SYNC_INTERVAL = form_data.SHAREPOINT_SYNC_INTERVAL

    secret = config.SHAREPOINT_CLIENT_SECRET
    return SharePointConfigResponse(
        ENABLE_SHAREPOINT_SYNC=config.ENABLE_SHAREPOINT_SYNC,
        SHAREPOINT_TENANT_ID=config.SHAREPOINT_TENANT_ID,
        SHAREPOINT_CLIENT_ID=config.SHAREPOINT_CLIENT_ID,
        SHAREPOINT_CLIENT_SECRET="**********" if secret else "",
        SHAREPOINT_SYNC_INTERVAL=config.SHAREPOINT_SYNC_INTERVAL,
    )


############################
# Browse
############################


class ResolveUrlForm(BaseModel):
    url: str


class ResolvedSiteResponse(BaseModel):
    site_id: str
    site_name: str
    web_url: str


class DriveResponse(BaseModel):
    id: str
    name: str
    driveType: str


class DriveItemResponse(BaseModel):
    id: str
    name: str
    size: int
    isFolder: bool
    childCount: int = 0
    mimeType: str = ""
    lastModifiedDateTime: str = ""


def _get_graph_client(request: Request) -> SharePointGraphClient:
    """Create a Graph client from current config."""
    config = request.app.state.config
    tenant_id = config.SHAREPOINT_TENANT_ID
    client_id = config.SHAREPOINT_CLIENT_ID
    client_secret = config.SHAREPOINT_CLIENT_SECRET

    if not all([tenant_id, client_id, client_secret]):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="SharePoint credentials not configured. Set Tenant ID, Client ID, and Client Secret.",
        )

    return SharePointGraphClient(tenant_id, client_id, client_secret)


@router.post("/browse/resolve", response_model=ResolvedSiteResponse)
async def resolve_sharepoint_site(
    request: Request,
    form_data: ResolveUrlForm,
    user=Depends(get_admin_user),
):
    """Resolve a SharePoint URL to site information."""
    client = _get_graph_client(request)

    try:
        site = client.resolve_site_from_url(form_data.url)
    except ValueError as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(e),
        )
    except Exception:
        log.exception("Failed to resolve SharePoint site")
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Failed to resolve site. Check the URL and credentials.",
        )

    return ResolvedSiteResponse(
        site_id=site["id"],
        site_name=site["displayName"],
        web_url=site["webUrl"],
    )


@router.get("/browse/{site_id}/drives", response_model=list[DriveResponse])
async def list_sharepoint_drives(
    request: Request,
    site_id: str,
    user=Depends(get_admin_user),
):
    """List document libraries (drives) for a resolved site."""
    client = _get_graph_client(request)

    try:
        drives = client.list_drives(site_id)
    except Exception:
        log.exception("Failed to list drives")
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Failed to list document libraries. Check credentials and site access.",
        )

    return [DriveResponse(**d) for d in drives]


@router.get(
    "/browse/drives/{drive_id}/items", response_model=list[DriveItemResponse]
)
async def list_sharepoint_items(
    request: Request,
    drive_id: str,
    parent_id: Optional[str] = None,
    user=Depends(get_admin_user),
):
    """List files and folders in a drive or folder."""
    client = _get_graph_client(request)

    try:
        items = client.list_items(drive_id, parent_id)
    except Exception:
        log.exception("Failed to list items")
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Failed to list items. Check credentials and drive access.",
        )

    return [DriveItemResponse(**i) for i in items]


############################
# Sites
############################

@router.get("/sites", response_model=list[SharePointSiteModel])
async def get_sharepoint_sites(
    request: Request,
    user=Depends(get_admin_user),
    db: Session = Depends(get_session),
):
    """List all configured SharePoint sites."""
    sites = SharePoints.get_sites(db=db)
    all_counts = SharePoints.get_all_file_counts(db=db)
    for site in sites:
        counts = all_counts.get(site.id, {})
        site.file_count = counts.get("file_count", 0)
        site.error_count = counts.get("error_count", 0)
        site.excluded_count = counts.get("excluded_count", 0)
    return sites


@router.post("/sites", response_model=Optional[SharePointSiteModel])
async def add_sharepoint_site(
    request: Request,
    form_data: SharePointSiteForm,
    user=Depends(get_admin_user),
    db: Session = Depends(get_session),
):
    """Add a new SharePoint site configuration and create a linked KB."""
    kb_name = form_data.kb_name or f"SharePoint: {form_data.site_name or form_data.site_id}"

    # Create a Knowledge Base for this site
    kb = Knowledges.insert_new_knowledge(
        user_id=user.id,
        form_data=KnowledgeForm(
            name=kb_name,
            description=f"Auto-synced from SharePoint: {form_data.site_url}",
        ),
        db=db,
    )

    if not kb:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=ERROR_MESSAGES.DEFAULT("Failed to create Knowledge Base for SharePoint site."),
        )

    site = SharePoints.insert_new_site(form_data, kb_id=kb.id, db=db)
    if not site:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=ERROR_MESSAGES.DEFAULT("Failed to save SharePoint site configuration."),
        )

    return site


@router.delete("/sites/{id}")
async def delete_sharepoint_site(
    request: Request,
    id: str,
    user=Depends(get_admin_user),
    db: Session = Depends(get_session),
):
    """Remove a SharePoint site configuration and all associated resources."""
    site = SharePoints.get_site_by_id(id, db=db)
    if not site:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=ERROR_MESSAGES.NOT_FOUND,
        )

    kb_id = site.kb_id

    # 1. Clean up OWUI files and their vector entries
    sp_files = SharePoints.get_files_by_site(id, limit=1000, db=db)
    for sp_file in sp_files:
        if sp_file.owui_file_id:
            try:
                VECTOR_DB_CLIENT.delete(
                    collection_name=kb_id,
                    filter={"file_id": sp_file.owui_file_id},
                )
            except Exception as e:
                log.debug(f"Vector cleanup for {sp_file.owui_file_id}: {e}")
            Files.delete_file_by_id(sp_file.owui_file_id)

    # 2. Delete the KB vector collection and metadata embedding
    if kb_id:
        try:
            VECTOR_DB_CLIENT.delete_collection(collection_name=kb_id)
        except Exception as e:
            log.debug(f"KB collection cleanup: {e}")

        try:
            from open_webui.routers.knowledge import (
                remove_knowledge_base_metadata_embedding,
            )

            remove_knowledge_base_metadata_embedding(kb_id)
        except Exception as e:
            log.debug(f"KB metadata embedding cleanup: {e}")

        Knowledges.delete_knowledge_by_id(id=kb_id, db=db)

    # 3. Delete the SharePoint site records (files cascade via FK)
    success = SharePoints.delete_site_by_id(id, db=db)
    if not success:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=ERROR_MESSAGES.DEFAULT("Failed to delete site configuration."),
        )

    return {"status": True, "id": id}


@router.put("/sites/{id}", response_model=Optional[SharePointSiteModel])
async def update_sharepoint_site(
    request: Request,
    id: str,
    form_data: SharePointSiteUpdateForm,
    user=Depends(get_admin_user),
    db: Session = Depends(get_session),
):
    """Update a SharePoint site's sync scope or mode."""
    site = SharePoints.get_site_by_id(id, db=db)
    if not site:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=ERROR_MESSAGES.NOT_FOUND,
        )

    update_data = {}
    if form_data.selected_items is not None:
        update_data["selected_items"] = form_data.selected_items
        # Clear delta_link so next sync picks up the new scope
        update_data["delta_link"] = None
    if form_data.sync_mode is not None:
        update_data["sync_mode"] = form_data.sync_mode
    if form_data.kb_name is not None:
        update_data["kb_name"] = form_data.kb_name

    if not update_data:
        return site

    updated = SharePoints.update_site(id, update_data, db=db)
    if not updated:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=ERROR_MESSAGES.DEFAULT("Failed to update site configuration."),
        )

    return updated


############################
# Site Files
############################


@router.get(
    "/sites/{site_id}/files", response_model=list[SharePointFileModel]
)
async def get_sharepoint_site_files(
    request: Request,
    site_id: str,
    skip: int = 0,
    limit: int = 200,
    user=Depends(get_admin_user),
    db: Session = Depends(get_session),
):
    """List tracked files for a site with pagination."""
    site = SharePoints.get_site_by_id(site_id, db=db)
    if not site:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=ERROR_MESSAGES.NOT_FOUND,
        )
    return SharePoints.get_files_by_site(site_id, skip=skip, limit=limit, db=db)


############################
# Sync
############################


@router.post("/retry/{site_id}")
async def retry_sharepoint_errors(
    request: Request,
    site_id: str,
    user=Depends(get_admin_user),
):
    """Retry embedding for files that previously failed."""
    from open_webui.utils.sharepoint_sync import trigger_retry_errors

    try:
        result = await trigger_retry_errors(app=request.app, site_id=site_id)
        return result
    except ValueError as e:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=str(e),
        )
    except Exception:
        log.exception("Retry failed")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=ERROR_MESSAGES.DEFAULT("Retry operation failed."),
        )


@router.post("/sync")
async def trigger_sharepoint_sync(
    request: Request,
    form_data: SharePointSyncForm,
    user=Depends(get_admin_user),
):
    """Trigger SharePoint sync (manual or forced)."""
    from open_webui.utils.sharepoint_sync import trigger_sync

    try:
        result = await trigger_sync(
            app=request.app,
            site_id=form_data.site_id,
            force=form_data.force,
            clear_exclusions=form_data.clear_exclusions,
        )
        return result
    except Exception:
        log.exception("Sync failed")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=ERROR_MESSAGES.DEFAULT("Sync operation failed."),
        )


@router.post("/sync/stream")
async def trigger_sharepoint_sync_stream(
    request: Request,
    form_data: SharePointSyncForm,
    user=Depends(get_admin_user),
):
    """Trigger SharePoint sync with SSE progress streaming."""
    from open_webui.utils.sharepoint_sync import trigger_sync_stream

    return StreamingResponse(
        trigger_sync_stream(
            app=request.app,
            site_id=form_data.site_id,
            force=form_data.force,
            clear_exclusions=form_data.clear_exclusions,
        ),
        media_type="text/event-stream",
    )
