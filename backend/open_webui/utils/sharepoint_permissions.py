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
_GROUP_CACHE_MAX_SIZE = 10_000


def _get_user_groups(app, user_email: str) -> list[str]:
    """
    Get a user's group memberships via Microsoft Graph API.
    Results are cached for _GROUP_CACHE_TTL seconds.
    """
    now = time.time()
    with _group_cache_lock:
        cached = _group_cache.get(user_email)
        if cached and (now - cached[1]) < _GROUP_CACHE_TTL:
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
            # Evict oldest entries if cache exceeds max size
            if len(_group_cache) >= _GROUP_CACHE_MAX_SIZE:
                sorted_keys = sorted(
                    _group_cache, key=lambda k: _group_cache[k][1]
                )
                for k in sorted_keys[: len(_group_cache) - _GROUP_CACHE_MAX_SIZE + 1]:
                    del _group_cache[k]
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
    if not allowed_users and not allowed_groups:
        return True

    if "*" in allowed_users:
        return True

    email_lower = user_email.lower()
    for allowed in allowed_users:
        if allowed.lower() == email_lower:
            return True

    if allowed_groups and user_groups:
        user_groups_lower = {g.lower() for g in user_groups}
        for group in allowed_groups:
            if group.lower() in user_groups_lower:
                return True

    return False


def _resolve_sp_context(
    file_ids: set[str], user, app
) -> Optional[tuple[dict, set, str, list[str]]]:
    """
    Batch-resolve SharePoint file records, filter-mode sites, user email,
    and group memberships. Returns None if no filtering is needed.
    """
    from open_webui.models.sharepoint import SharePoints

    sp_file_list = SharePoints.get_files_by_owui_ids(list(file_ids))
    if not sp_file_list:
        return None

    sp_files = {f.owui_file_id: f for f in sp_file_list}

    site_ids = list({f.site_config_id for f in sp_file_list})
    site_list = SharePoints.get_sites_by_ids(site_ids)
    filter_sites = {
        s.id for s in site_list if s and s.sync_mode == "filter"
    }
    if not filter_sites:
        return None

    if hasattr(user, "email"):
        user_email = user.email or ""
    elif isinstance(user, dict):
        user_email = user.get("email", "")
    else:
        user_email = ""

    log.debug(
        f"SharePoint filter: checking permissions for "
        f"user_email='{user_email}'"
    )

    if not user_email:
        log.warning("SharePoint filter: user has no email, blocking filter-mode files")

    user_groups = _get_user_groups(app, user_email) if user_email else []

    return sp_files, filter_sites, user_email, user_groups


def filter_by_sharepoint_permissions(
    chunks: list[dict],
    user: dict,
    app,
) -> list[dict]:
    """
    Filter retrieval chunks based on SharePoint permissions (builtin tool path).

    Only affects files from SharePoint sites with sync_mode="filter".
    Non-SharePoint files and files from "none" mode sites pass through.
    """
    if not chunks:
        return chunks

    file_ids = {c.get("file_id") for c in chunks if c.get("file_id")}
    if not file_ids:
        return chunks

    try:
        ctx = _resolve_sp_context(file_ids, user, app)
        if ctx is None:
            return chunks

        sp_files, filter_sites, user_email, user_groups = ctx

        filtered = []
        for chunk in chunks:
            file_id = chunk.get("file_id")

            if not file_id or file_id not in sp_files:
                filtered.append(chunk)
                continue

            sp_file = sp_files[file_id]

            if sp_file.site_config_id not in filter_sites:
                filtered.append(chunk)
                continue

            if not user_email:
                continue

            has_access = _check_user_access(
                sp_file.allowed_users or [],
                sp_file.allowed_groups or [],
                user_email,
                user_groups,
            )
            log.debug(
                f"SharePoint filter: {'GRANTED' if has_access else 'DENIED'} "
                f"file={sp_file.filename} user={user_email}"
            )
            if has_access:
                filtered.append(chunk)

        return filtered
    except Exception as e:
        log.error(f"SharePoint permission filter error: {e}")
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
    """
    if not sources:
        return sources

    all_file_ids = set()
    for source in sources:
        for meta in source.get("metadata") or []:
            fid = (meta or {}).get("file_id")
            if fid:
                all_file_ids.add(fid)

    if not all_file_ids:
        return sources

    try:
        ctx = _resolve_sp_context(all_file_ids, user, app)
        if ctx is None:
            return sources

        sp_files, filter_sites, user_email, user_groups = ctx

        filtered_sources = []
        for source in sources:
            documents = source.get("document") or []
            metadatas = source.get("metadata") or []
            distances = source.get("distances")

            keep = []
            for i, meta in enumerate(metadatas):
                file_id = (meta or {}).get("file_id")

                if not file_id or file_id not in sp_files:
                    keep.append(i)
                    continue

                sp_file = sp_files[file_id]

                if sp_file.site_config_id not in filter_sites:
                    keep.append(i)
                    continue

                if not user_email:
                    continue

                has_access = _check_user_access(
                    sp_file.allowed_users or [],
                    sp_file.allowed_groups or [],
                    user_email,
                    user_groups,
                )
                log.debug(
                    f"SharePoint filter: {'GRANTED' if has_access else 'DENIED'} "
                    f"file={sp_file.filename} user={user_email}"
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
    except Exception as e:
        log.error(f"SharePoint permission filter (sources) error: {e}")
        # Fail closed: remove all entries whose file_id is in the SP set
        return [
            s for s in sources
            if not any(
                (m or {}).get("file_id") in all_file_ids
                for m in (s.get("metadata") or [])
            )
        ]
