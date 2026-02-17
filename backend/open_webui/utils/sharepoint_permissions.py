"""
Built-in SharePoint permission filtering for RAG retrieval.

When a SharePoint site is configured with sync_mode="filter", retrieved
chunks are filtered based on the user's SharePoint permissions. This
prevents unauthorized content from reaching the LLM.

Permissions are stored during sync (allowed_users / allowed_groups on
each SharePointFile) and checked at retrieval time against the current
user's email and group memberships (resolved via Microsoft Graph API).
"""

import logging
import os
import threading
import time
from typing import Optional

log = logging.getLogger(__name__)

# In-memory cache for user group memberships: {email: (groups, timestamp)}
_group_cache: dict[str, tuple[list[str], float]] = {}
_group_cache_lock = threading.Lock()
_GROUP_CACHE_TTL = int(os.environ.get("SHAREPOINT_GROUP_CACHE_TTL", "300"))  # seconds


def _get_user_groups(app, user_email: str) -> list[str]:
    """
    Get a user's group memberships via Microsoft Graph API.
    Results are cached for _GROUP_CACHE_TTL seconds.
    """
    now = time.time()
    with _group_cache_lock:
        cached = _group_cache.get(user_email)
        if cached and (now - cached[1]) < _GROUP_CACHE_TTL:
            log.debug(
                f"SharePoint filter: group cache hit for {user_email}, "
                f"{len(cached[0])} groups"
            )
            return cached[0]

    try:
        from open_webui.retrieval.loaders.sharepoint import SharePointGraphClient

        config = app.state.config
        tenant_id = getattr(config, "SHAREPOINT_TENANT_ID", "")
        client_id = getattr(config, "SHAREPOINT_CLIENT_ID", "")
        client_secret = getattr(config, "SHAREPOINT_CLIENT_SECRET", "")

        if not all([tenant_id, client_id, client_secret]):
            return []

        client = SharePointGraphClient(tenant_id, client_id, client_secret)
        groups = client.get_user_group_memberships(user_email)
        log.debug(
            f"SharePoint filter: resolved {len(groups)} groups for "
            f"{user_email}: {groups}"
        )
        with _group_cache_lock:
            _group_cache[user_email] = (groups, now)
        return groups
    except Exception as e:
        log.warning(
            f"SharePoint filter: failed to get group memberships for "
            f"{user_email}: {e} — user will have no group-based access"
        )
        return []


def _check_user_access(
    allowed_users: list,
    allowed_groups: list,
    user_email: str,
    user_groups: list[str],
) -> bool:
    """Check if a user has access based on stored permissions."""
    # No permissions recorded = open access
    if not allowed_users and not allowed_groups:
        return True

    # Org-wide access
    if "*" in allowed_users:
        return True

    # Direct user email match
    email_lower = user_email.lower()
    for allowed in allowed_users:
        if allowed.lower() == email_lower:
            return True

    # Group membership match
    if allowed_groups and user_groups:
        user_groups_lower = {g.lower() for g in user_groups}
        for group in allowed_groups:
            if group.lower() in user_groups_lower:
                return True

    return False


def filter_by_sharepoint_permissions(
    chunks: list[dict],
    user: dict,
    app,
) -> list[dict]:
    """
    Filter retrieval chunks based on SharePoint permissions.

    Only affects files from SharePoint sites with sync_mode="filter".
    Non-SharePoint files and files from "none" mode sites pass through.

    Args:
        chunks: List of retrieval result dicts with "file_id" keys
        user: User dict with "email", "id", etc.
        app: FastAPI app for accessing config

    Returns:
        Filtered list of chunks the user has access to.
    """
    if not chunks:
        log.debug("SharePoint filter: no chunks to filter")
        return chunks

    # Collect unique file_ids from chunks
    file_ids = {c.get("file_id") for c in chunks if c.get("file_id")}
    if not file_ids:
        log.debug(
            f"SharePoint filter: {len(chunks)} chunks but none have file_id"
        )
        return chunks

    try:
        from open_webui.models.sharepoint import SharePoints

        log.debug(
            f"SharePoint filter: {len(chunks)} chunks, "
            f"{len(file_ids)} unique file_ids: {file_ids}"
        )

        # Batch lookup: get SharePoint file records for these file_ids
        sp_file_list = SharePoints.get_files_by_owui_ids(list(file_ids))
        if not sp_file_list:
            log.debug("SharePoint filter: no SP file records found — passing all chunks through")
            return chunks

        sp_files = {f.owui_file_id: f for f in sp_file_list}

        # Batch lookup site configs to check sync_mode
        site_ids = list({f.site_config_id for f in sp_file_list})
        site_list = SharePoints.get_sites_by_ids(site_ids)
        site_configs = {s.id: s for s in site_list}

        # Check if any sites use filter mode
        filter_sites = {
            sid
            for sid, site in site_configs.items()
            if site and site.sync_mode == "filter"
        }

        if not filter_sites:
            log.debug("SharePoint filter: no filter-mode sites — passing all chunks through")
            return chunks

        # We need to check permissions — get user info
        user_email = user.get("email", "")
        log.debug(
            f"SharePoint filter: checking permissions for user_email='{user_email}', "
            f"{len(chunks)} chunks, {len(sp_files)} SP files in filter-mode sites"
        )
        if not user_email:
            # No email = can't verify identity, block filter-mode files
            log.warning(
                "SharePoint filter: user has no email, blocking filter-mode files"
            )
            return [
                c
                for c in chunks
                if c.get("file_id") not in sp_files
                or sp_files.get(c.get("file_id"), None) is None
                or sp_files[c["file_id"]].site_config_id not in filter_sites
            ]

        # Get user's group memberships (cached)
        user_groups = _get_user_groups(app, user_email)

        # Filter chunks
        filtered = []
        for chunk in chunks:
            file_id = chunk.get("file_id")

            # Not a SharePoint file — keep
            if not file_id or file_id not in sp_files:
                filtered.append(chunk)
                continue

            sp_file = sp_files[file_id]

            # Site doesn't use filter mode — keep
            if sp_file.site_config_id not in filter_sites:
                filtered.append(chunk)
                continue

            # Check permissions
            au = sp_file.allowed_users or []
            ag = sp_file.allowed_groups or []
            has_access = _check_user_access(
                allowed_users=au,
                allowed_groups=ag,
                user_email=user_email,
                user_groups=user_groups,
            )
            log.debug(
                f"SharePoint filter: {'GRANTED' if has_access else 'DENIED'} "
                f"file={sp_file.filename} user={user_email} "
                f"allowed_users={au} allowed_groups={ag} "
                f"user_groups={user_groups}"
            )
            if has_access:
                filtered.append(chunk)

        return filtered
    except Exception as e:
        log.error(f"SharePoint permission filter error: {e}")
        # Fail closed: remove SharePoint-sourced chunks rather than
        # risking exposure of permission-restricted content.
        return [c for c in chunks if c.get("file_id") not in file_ids]


def filter_sources_by_sharepoint_permissions(
    sources: list[dict],
    user,
    app,
) -> list[dict]:
    """
    Filter middleware RAG sources based on SharePoint permissions.

    Works with the nested sources format from get_sources_from_items():
    each source has 'document', 'metadata', and 'distances' parallel arrays.

    Only affects files from SharePoint sites with sync_mode="filter".
    Non-SharePoint files pass through unchanged.

    Args:
        sources: List of source dicts with document/metadata/distances arrays
        user: UserModel object or dict with "email" key
        app: FastAPI app for accessing config

    Returns:
        Filtered sources with denied entries removed.
    """
    if not sources:
        log.debug("SharePoint filter (sources): no sources to filter")
        return sources

    # Collect all file_ids across all sources
    all_file_ids = set()
    for source in sources:
        for meta in source.get("metadata") or []:
            fid = (meta or {}).get("file_id")
            if fid:
                all_file_ids.add(fid)

    if not all_file_ids:
        log.debug("SharePoint filter (sources): no file_ids in metadata")
        return sources

    try:
        from open_webui.models.sharepoint import SharePoints

        log.debug(
            f"SharePoint filter (sources): {len(all_file_ids)} unique "
            f"file_ids: {all_file_ids}"
        )

        # Batch lookup SharePoint file records
        sp_file_list = SharePoints.get_files_by_owui_ids(list(all_file_ids))
        if not sp_file_list:
            log.debug(
                "SharePoint filter (sources): no SP file records — "
                "passing all through"
            )
            return sources

        sp_files = {f.owui_file_id: f for f in sp_file_list}

        # Batch lookup site configs, check for filter mode
        site_ids = list({f.site_config_id for f in sp_file_list})
        site_list = SharePoints.get_sites_by_ids(site_ids)
        site_configs = {s.id: s for s in site_list}

        filter_sites = {
            sid
            for sid, site in site_configs.items()
            if site and site.sync_mode == "filter"
        }

        if not filter_sites:
            log.debug(
                "SharePoint filter (sources): no filter-mode sites — "
                "passing all through"
            )
            return sources

        # Get user email (UserModel or dict)
        if hasattr(user, "email"):
            user_email = user.email or ""
        elif isinstance(user, dict):
            user_email = user.get("email", "")
        else:
            user_email = ""

        log.debug(
            f"SharePoint filter (sources): checking permissions for "
            f"user_email='{user_email}'"
        )

        if not user_email:
            log.warning(
                "SharePoint filter (sources): user has no email, "
                "blocking filter-mode files"
            )
            return _filter_source_entries(
                sources, sp_files, filter_sites, deny_all=True
            )

        # Get user's group memberships (cached)
        user_groups = _get_user_groups(app, user_email)

        return _filter_source_entries(
            sources, sp_files, filter_sites,
            user_email=user_email, user_groups=user_groups,
        )
    except Exception as e:
        log.error(f"SharePoint permission filter (sources) error: {e}")
        # Fail closed: remove all entries whose file_id is in the SP set
        filtered = []
        for source in sources:
            metadatas = source.get("metadata") or []
            has_sp = any(
                (m or {}).get("file_id") in all_file_ids for m in metadatas
            )
            if not has_sp:
                filtered.append(source)
        return filtered


def _filter_source_entries(
    sources: list[dict],
    sp_files: dict,
    filter_sites: set,
    user_email: str = "",
    user_groups: list[str] = None,
    deny_all: bool = False,
) -> list[dict]:
    """
    Remove denied entries from the nested sources structure.

    Preserves parallel alignment of document/metadata/distances arrays.
    Drops entire source dicts that become empty after filtering.
    """
    filtered_sources = []
    for source in sources:
        documents = source.get("document") or []
        metadatas = source.get("metadata") or []
        distances = source.get("distances")

        keep = []
        for i, meta in enumerate(metadatas):
            file_id = (meta or {}).get("file_id")

            # Not a SharePoint file — keep
            if not file_id or file_id not in sp_files:
                keep.append(i)
                continue

            sp_file = sp_files[file_id]

            # Site not in filter mode — keep
            if sp_file.site_config_id not in filter_sites:
                keep.append(i)
                continue

            if deny_all:
                log.debug(
                    f"SharePoint filter (sources): DENIED (no email) "
                    f"file={sp_file.filename}"
                )
                continue

            has_access = _check_user_access(
                allowed_users=sp_file.allowed_users or [],
                allowed_groups=sp_file.allowed_groups or [],
                user_email=user_email,
                user_groups=user_groups or [],
            )
            log.debug(
                f"SharePoint filter (sources): "
                f"{'GRANTED' if has_access else 'DENIED'} "
                f"file={sp_file.filename} user={user_email} "
                f"allowed_users={sp_file.allowed_users or []} "
                f"allowed_groups={sp_file.allowed_groups or []} "
                f"user_groups={user_groups}"
            )
            if has_access:
                keep.append(i)

        if keep:
            new_source = dict(source)
            new_source["document"] = [
                documents[i] for i in keep if i < len(documents)
            ]
            new_source["metadata"] = [
                metadatas[i] for i in keep if i < len(metadatas)
            ]
            if distances is not None:
                new_source["distances"] = [
                    distances[i] for i in keep if i < len(distances)
                ]
            filtered_sources.append(new_source)

    return filtered_sources
