"""Microsoft Graph API client for querying Intune assignments."""

import msal
import requests


class GraphClient:
    """Handles authentication and API calls to Microsoft Graph."""

    GRAPH_BASE = "https://graph.microsoft.com/beta"
    SCOPES = ["https://graph.microsoft.com/.default"]

    def __init__(self, tenant_id, client_id, client_secret):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self._token_cache = None
        self._app = msal.ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret,
        )

    def _get_token(self):
        """Acquire an access token using client credentials flow."""
        result = self._app.acquire_token_for_client(scopes=self.SCOPES)
        if "access_token" not in result:
            error = result.get("error_description", result.get("error", "Unknown"))
            raise RuntimeError(f"Failed to acquire token: {error}")
        return result["access_token"]

    def _headers(self):
        return {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json",
        }

    def _get_paginated(self, url, params=None):
        """Fetch all pages of a paginated Graph API response."""
        results = []
        while url:
            resp = requests.get(url, headers=self._headers(), params=params, timeout=30)
            resp.raise_for_status()
            data = resp.json()
            results.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            params = None  # nextLink already includes query params
        return results

    # ── Groups ───────────────────────────────────────────────────────────

    def get_groups(self):
        """Return all Entra ID (Azure AD) groups."""
        url = f"{self.GRAPH_BASE}/groups"
        params = {
            "$select": "id,displayName,description,groupTypes,membershipRule",
            "$orderby": "displayName",
            "$top": "999",
        }
        return self._get_paginated(url, params)

    # ── Assignment helpers ───────────────────────────────────────────────

    def _extract_assignments_for_group(self, items, group_id):
        """Filter items that have an assignment targeting the given group."""
        matched = []
        for item in items:
            assignments = item.get("assignments", [])
            for assignment in assignments:
                target = assignment.get("target", {})
                if target.get("groupId") == group_id:
                    intent = assignment.get("intent", "")
                    filter_id = target.get("deviceAndAppManagementAssignmentFilterId", "")
                    filter_type = target.get("deviceAndAppManagementAssignmentFilterType", "")
                    target_type = target.get("@odata.type", "")
                    matched.append({
                        "id": item.get("id", ""),
                        "displayName": item.get("displayName", item.get("name", "N/A")),
                        "description": item.get("description", ""),
                        "assignmentType": _friendly_target_type(target_type),
                        "intent": intent,
                        "filterId": filter_id,
                        "filterType": filter_type,
                    })
                    break
        return matched

    # ── Device Configurations ────────────────────────────────────────────

    def get_device_configurations(self, group_id):
        """Get device configuration profiles assigned to a group."""
        url = f"{self.GRAPH_BASE}/deviceManagement/deviceConfigurations"
        params = {"$expand": "assignments"}
        items = self._get_paginated(url, params)
        return self._extract_assignments_for_group(items, group_id)

    # ── Settings Catalog ─────────────────────────────────────────────────

    def get_settings_catalog(self, group_id):
        """Get Settings Catalog policies assigned to a group."""
        url = f"{self.GRAPH_BASE}/deviceManagement/configurationPolicies"
        params = {"$expand": "assignments"}
        items = self._get_paginated(url, params)
        return self._extract_assignments_for_group(items, group_id)

    # ── Applications ─────────────────────────────────────────────────────

    def get_applications(self, group_id):
        """Get applications assigned to a group."""
        url = f"{self.GRAPH_BASE}/deviceAppManagement/mobileApps"
        params = {
            "$expand": "assignments",
            "$filter": "isAssigned eq true",
        }
        items = self._get_paginated(url, params)
        return self._extract_assignments_for_group(items, group_id)

    # ── Scripts ──────────────────────────────────────────────────────────

    def get_scripts(self, group_id):
        """Get PowerShell scripts assigned to a group."""
        url = f"{self.GRAPH_BASE}/deviceManagement/deviceManagementScripts"
        params = {"$expand": "assignments"}
        items = self._get_paginated(url, params)
        return self._extract_assignments_for_group(items, group_id)

    # ── Remediation Scripts (Proactive Remediations) ─────────────────────

    def get_remediations(self, group_id):
        """Get proactive remediation scripts assigned to a group."""
        url = f"{self.GRAPH_BASE}/deviceManagement/deviceHealthScripts"
        params = {"$expand": "assignments"}
        items = self._get_paginated(url, params)
        return self._extract_assignments_for_group(items, group_id)


def _friendly_target_type(odata_type):
    """Convert OData target type to a human-readable label."""
    mapping = {
        "#microsoft.graph.groupAssignmentTarget": "Include",
        "#microsoft.graph.exclusionGroupAssignmentTarget": "Exclude",
        "#microsoft.graph.allDevicesAssignmentTarget": "All Devices",
        "#microsoft.graph.allLicensedUsersAssignmentTarget": "All Users",
    }
    return mapping.get(odata_type, odata_type)
