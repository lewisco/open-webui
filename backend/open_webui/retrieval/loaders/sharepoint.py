import logging
import os
from typing import Optional
from urllib.parse import urlparse

import requests

log = logging.getLogger(__name__)

# Extensions to always skip during sync
SKIPPED_EXTENSIONS = {
    ".tmp",
    ".lock",
    ".bak",
    ".swp",
    ".ds_store",
    ".thumbs.db",
    ".gitignore",
    ".gitkeep",
}


class SharePointGraphClient:
    """Microsoft Graph API client for SharePoint operations using MSAL client credentials."""

    GRAPH_BASE = "https://graph.microsoft.com/v1.0"
    METADATA_TIMEOUT = 30  # seconds, for API metadata calls
    DOWNLOAD_TIMEOUT = 120  # seconds, for file content downloads

    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self._msal_app = None

    def _get_msal_app(self):
        """Lazy-initialize and return the MSAL app (reused for token caching)."""
        if self._msal_app is None:
            import msal

            self._msal_app = msal.ConfidentialClientApplication(
                self.client_id,
                authority=f"https://login.microsoftonline.com/{self.tenant_id}",
                client_credential=self.client_secret,
            )
        return self._msal_app

    def _get_token(self) -> str:
        """Acquire token via MSAL client credentials flow."""
        result = self._get_msal_app().acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        if "access_token" not in result:
            error_desc = result.get("error_description", result.get("error", "Unknown error"))
            raise Exception(f"Failed to acquire token: {error_desc}")
        return result["access_token"]

    def _headers(self) -> dict:
        """Return authorization headers."""
        token = self._get_token()
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

    # ---- Browse Methods ----

    def resolve_site(self, hostname: str, site_path: str) -> dict:
        """
        Resolve a SharePoint site from hostname and path.
        GET /sites/{hostname}:/{site_path}
        Returns: {id, displayName, webUrl}
        """
        # Clean up the site path
        site_path = site_path.strip("/")
        url = f"{self.GRAPH_BASE}/sites/{hostname}:/{site_path}"
        resp = requests.get(url, headers=self._headers(), timeout=self.METADATA_TIMEOUT)
        resp.raise_for_status()
        data = resp.json()
        return {
            "id": data["id"],
            "displayName": data.get("displayName", ""),
            "webUrl": data.get("webUrl", ""),
        }

    def resolve_site_from_url(self, sharepoint_url: str) -> dict:
        """
        Parse a SharePoint URL and resolve the site.
        Example: https://contoso.sharepoint.com/sites/MySite
        """
        parsed = urlparse(sharepoint_url)
        hostname = parsed.hostname
        if not hostname or not hostname.endswith(".sharepoint.com"):
            raise ValueError(
                "URL must be a SharePoint site (*.sharepoint.com). "
                f"Got: {sharepoint_url}"
            )

        # Extract site path from URL path
        path = parsed.path.strip("/")
        if not path:
            raise ValueError(
                f"No site path found in URL: {sharepoint_url}. "
                "Expected format: https://tenant.sharepoint.com/sites/SiteName"
            )

        return self.resolve_site(hostname, path)

    def list_drives(self, site_id: str) -> list[dict]:
        """
        List document libraries (drives) for a site.
        GET /sites/{site_id}/drives
        Returns: [{id, name, driveType}]
        """
        url = f"{self.GRAPH_BASE}/sites/{site_id}/drives"
        resp = requests.get(url, headers=self._headers(), timeout=self.METADATA_TIMEOUT)
        resp.raise_for_status()
        data = resp.json()
        return [
            {
                "id": d["id"],
                "name": d.get("name", ""),
                "driveType": d.get("driveType", ""),
            }
            for d in data.get("value", [])
        ]

    def list_items(
        self, drive_id: str, item_id: Optional[str] = None
    ) -> list[dict]:
        """
        List files and folders in a drive or folder.
        GET /drives/{drive_id}/root/children (root)
        GET /drives/{drive_id}/items/{item_id}/children (subfolder)
        Returns: [{id, name, size, folder, file, parentReference}]
        """
        if item_id:
            url = f"{self.GRAPH_BASE}/drives/{drive_id}/items/{item_id}/children"
        else:
            url = f"{self.GRAPH_BASE}/drives/{drive_id}/root/children"

        items = []
        while url:
            resp = requests.get(url, headers=self._headers(), timeout=self.METADATA_TIMEOUT)
            resp.raise_for_status()
            data = resp.json()

            for item in data.get("value", []):
                items.append(
                    {
                        "id": item["id"],
                        "name": item.get("name", ""),
                        "size": item.get("size", 0),
                        "isFolder": "folder" in item,
                        "childCount": item.get("folder", {}).get("childCount", 0)
                        if "folder" in item
                        else 0,
                        "mimeType": item.get("file", {}).get("mimeType", "")
                        if "file" in item
                        else "",
                        "lastModifiedDateTime": item.get("lastModifiedDateTime", ""),
                        "parentReference": {
                            "path": item.get("parentReference", {}).get("path", ""),
                            "driveId": item.get("parentReference", {}).get(
                                "driveId", ""
                            ),
                        },
                    }
                )

            url = data.get("@odata.nextLink")

        return items

    # ---- Sync Methods ----

    def get_folder_delta(
        self, drive_id: str, delta_link: Optional[str] = None
    ) -> tuple[list[dict], Optional[str]]:
        """
        Get incremental changes using delta query.
        Returns: (items, new_delta_link)
        """
        if delta_link:
            if not delta_link.startswith(self.GRAPH_BASE):
                raise ValueError(f"Invalid delta link origin: {delta_link[:100]}")
            url = delta_link
        else:
            url = f"{self.GRAPH_BASE}/drives/{drive_id}/root/delta"

        all_items = []
        new_delta_link = None

        while url:
            resp = requests.get(url, headers=self._headers(), timeout=self.METADATA_TIMEOUT)
            resp.raise_for_status()
            data = resp.json()

            for item in data.get("value", []):
                all_items.append(item)

            # Check for next page or delta link
            if "@odata.nextLink" in data:
                url = data["@odata.nextLink"]
            else:
                new_delta_link = data.get("@odata.deltaLink")
                url = None

        return all_items, new_delta_link

    def download_file(self, drive_id: str, item_id: str) -> bytes:
        """Download file content from a drive item."""
        url = f"{self.GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"
        resp = requests.get(url, headers=self._headers(), allow_redirects=True, timeout=self.DOWNLOAD_TIMEOUT)
        resp.raise_for_status()
        return resp.content

    def get_item_permissions(self, drive_id: str, item_id: str) -> dict:
        """
        Get permissions for a drive item.
        Returns: {users: [...], groups: [...]}
        """
        url = f"{self.GRAPH_BASE}/drives/{drive_id}/items/{item_id}/permissions"
        resp = requests.get(url, headers=self._headers(), timeout=self.METADATA_TIMEOUT)
        resp.raise_for_status()
        data = resp.json()

        users = []
        groups = []

        for perm in data.get("value", []):
            granted = perm.get("grantedToV2") or perm.get("grantedTo") or {}

            if "user" in granted:
                user_info = granted["user"]
                email = user_info.get("email") or user_info.get(
                    "displayName", ""
                )
                if email:
                    users.append(email)

            if "group" in granted:
                group_info = granted["group"]
                group_name = group_info.get("displayName") or group_info.get(
                    "email", ""
                )
                if group_name:
                    groups.append(group_name)

            # Handle shared links with scope
            if "link" in perm:
                link = perm["link"]
                if link.get("scope") == "organization":
                    # Org-wide access
                    users.append("*")

        return {"users": users, "groups": groups}

    def get_user_group_memberships(self, user_email: str) -> list[str]:
        """
        Get group display names for a user via Microsoft Graph.
        Requires User.Read.All or GroupMember.Read.All app permission.
        Returns list of group displayName strings.
        """
        url = f"{self.GRAPH_BASE}/users/{user_email}/memberOf"
        resp = requests.get(url, headers=self._headers(), timeout=self.METADATA_TIMEOUT)
        resp.raise_for_status()
        data = resp.json()

        groups = []
        for member in data.get("value", []):
            # Only include groups (not roles or other directory objects)
            odata_type = member.get("@odata.type", "")
            if odata_type in (
                "#microsoft.graph.group",
                "#microsoft.graph.unifiedGroup",
            ):
                display_name = member.get("displayName", "")
                if display_name:
                    groups.append(display_name)

        return groups

    # ---- Helpers ----

    @staticmethod
    def is_supported_file(
        filename: str,
        size: int,
        allowed_extensions: Optional[list[str]] = None,
        max_size: Optional[int] = None,
    ) -> bool:
        """
        Check if a file should be synced based on admin-configured limits.
        Uses the same extension/size rules as the Documents settings tab.
        """
        if not filename:
            return False

        ext = os.path.splitext(filename)[1].lower()

        # Skip known non-document extensions
        if ext in SKIPPED_EXTENSIONS:
            return False

        # Check against admin-configured allowed extensions
        if allowed_extensions:
            # Filter out empty strings
            allowed = [e for e in allowed_extensions if e]
            if allowed:
                ext_without_dot = ext[1:] if ext.startswith(".") else ext
                if ext_without_dot not in allowed:
                    return False

        # Check file size limit (in MB, matching admin config)
        if max_size and max_size > 0:
            max_bytes = max_size * 1024 * 1024
            if size > max_bytes:
                return False

        return True

    @staticmethod
    def get_item_path(item: dict) -> str:
        """Extract the full path for a delta item."""
        parent_path = item.get("parentReference", {}).get("path", "")
        name = item.get("name", "")

        # parentReference.path format: /drives/{id}/root:/folder/subfolder
        # Extract the portion after "root:" (the user-meaningful path)
        _, _, after_root = parent_path.partition("root:")

        if after_root:
            return f"{after_root}/{name}"
        return f"/{name}"
