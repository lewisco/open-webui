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
import threading
import time
from typing import Optional

log = logging.getLogger(__name__)

# In-memory cache for user group memberships: {email: (groups, timestamp)}
_group_cache: dict[str, tuple[list[str], float]] = {}
_group_cache_lock = threading.Lock()
_GROUP_CACHE_TTL = 300  # 5 minutes


def _get_user_groups(app, user_email: str) -> list[str]:
    """
    Get a user's group memberships via Microsoft Graph API.
    Results are cached for _GROUP_CACHE_TTL seconds.
    """
    now = time.time()
    with _group_cache_lock:
        cached = _group_cache.get(user_email)
        if cached and (now - cached[1]) < _GROUP_CACHE_TTL:
            log.info(
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
        log.info(
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
        log.info("SharePoint filter: no chunks to filter")
        return chunks

    # Collect unique file_ids from chunks
    file_ids = {c.get("file_id") for c in chunks if c.get("file_id")}
    if not file_ids:
        log.info(
            f"SharePoint filter: {len(chunks)} chunks but none have file_id"
        )
        return chunks

    try:
        from open_webui.models.sharepoint import SharePoints

        log.info(
            f"SharePoint filter: {len(chunks)} chunks, "
            f"{len(file_ids)} unique file_ids: {file_ids}"
        )

        # Batch lookup: get SharePoint file records for these file_ids
        sp_file_list = SharePoints.get_files_by_owui_ids(list(file_ids))
        if not sp_file_list:
            log.info("SharePoint filter: no SP file records found — passing all chunks through")
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
            log.info("SharePoint filter: no filter-mode sites — passing all chunks through")
            return chunks

        # We need to check permissions — get user info
        user_email = user.get("email", "")
        log.info(
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
            log.info(
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
