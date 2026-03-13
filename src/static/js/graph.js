/* ═══════════════════════════════════════════════════════════════════════════
   Intune Assignment Checker — SPA Graph Client (MSAL.js)
   Handles authentication and direct Graph API calls when running without
   the PowerShell backend (e.g. GitHub Pages).
   ═══════════════════════════════════════════════════════════════════════════ */

// eslint-disable-next-line no-unused-vars
var GraphClient = (function () {
    "use strict";

    var REDIRECT_URI = window.location.origin + window.location.pathname;
    var GRAPH_BASE = "https://graph.microsoft.com";
    var SCOPES = [
        "DeviceManagementConfiguration.Read.All",
        "DeviceManagementApps.Read.All",
        "DeviceManagementManagedDevices.Read.All",
        "Group.Read.All",
        "User.Read.All"
    ];

    var msalInstance = null;
    var activeAccount = null;

    // ── Initialise MSAL ──────────────────────────────────────────────────

    async function init(clientId, tenantId) {
        var authority = "https://login.microsoftonline.com/" + tenantId;

        var msalConfig = {
            auth: {
                clientId: clientId,
                authority: authority,
                redirectUri: REDIRECT_URI
            },
            cache: {
                cacheLocation: "localStorage",
                storeAuthStateInCookie: false
            }
        };

        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();
    }

    // ── Handle redirect callback (for redirect flow) ─────────────────────

    async function handleRedirect() {
        if (!msalInstance) return null;
        try {
            var response = await msalInstance.handleRedirectPromise();
            if (response) {
                activeAccount = response.account;
                return activeAccount;
            }
            var accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                activeAccount = accounts[0];
                return activeAccount;
            }
        } catch (err) {
            console.error("MSAL redirect error:", err);
        }
        return null;
    }

    // ── Sign in ──────────────────────────────────────────────────────────

    async function signIn() {
        if (!msalInstance) throw new Error("MSAL not initialised. Call init() first.");

        // Use redirect flow — popups are unreliable when async work
        // happens between the user click and the login call.
        await msalInstance.loginRedirect({ scopes: SCOPES });
        // Page will redirect to Microsoft login, then back.
        // handleRedirect() picks up the response on reload.
        return null;
    }

    // ── Sign out ─────────────────────────────────────────────────────────

    async function signOut() {
        if (!msalInstance) return;
        var account = activeAccount || (msalInstance.getAllAccounts()[0] || null);
        activeAccount = null;
        if (account) {
            try {
                await msalInstance.logoutRedirect({ account: account });
            } catch (e) {
                // If redirect fails, clear local cache
                msalInstance.clearCache();
            }
        }
    }

    // ── Acquire token silently (with fallback to redirect) ─────────────

    async function getToken() {
        if (!msalInstance) throw new Error("MSAL not initialised.");
        var account = activeAccount || (msalInstance.getAllAccounts()[0] || null);
        if (!account) throw new Error("No signed-in account. Please sign in first.");

        var request = { scopes: SCOPES, account: account };
        try {
            var response = await msalInstance.acquireTokenSilent(request);
            return response.accessToken;
        } catch (err) {
            // If silent fails (e.g. token expired & interaction required), redirect
            if (err instanceof msal.InteractionRequiredAuthError) {
                await msalInstance.acquireTokenRedirect(request);
                return null; // page will redirect
            }
            throw err;
        }
    }

    // ── Graph API fetch with auth header ─────────────────────────────────

    async function graphFetch(url) {
        var token = await getToken();
        var resp = await fetch(url, {
            headers: { "Authorization": "Bearer " + token }
        });
        if (!resp.ok) {
            var body = await resp.json().catch(function () { return {}; });
            var msg = (body.error && body.error.message) || "HTTP " + resp.status;
            throw new Error(msg);
        }
        return resp.json();
    }

    // ── Paginated fetch (follows @odata.nextLink) ────────────────────────

    async function graphFetchAll(url) {
        var results = [];
        var nextUrl = url;
        while (nextUrl) {
            var data = await graphFetch(nextUrl);
            if (data.value) {
                results = results.concat(data.value);
            }
            nextUrl = data["@odata.nextLink"] || null;
        }
        return results;
    }

    // ── Public API methods (mirror the PowerShell backend endpoints) ─────

    async function getAllGroups() {
        var url = GRAPH_BASE + "/v1.0/groups?$select=id,displayName,description,groupTypes,membershipRule&$top=999";
        var groups = await graphFetchAll(url);
        groups.sort(function (a, b) {
            return (a.displayName || "").localeCompare(b.displayName || "");
        });
        return groups;
    }

    // Helper to extract assignment info
    function extractAssignment(assignment, groupId) {
        var target = assignment.target || {};
        var targetGroupId = target.groupId || null;
        var odataType = target["@odata.type"] || "";

        if (odataType === "#microsoft.graph.allDevicesAssignmentTarget") {
            return { match: true, assignmentType: "All Devices" };
        }
        if (odataType === "#microsoft.graph.allLicensedUsersAssignmentTarget") {
            return { match: true, assignmentType: "All Users" };
        }
        if (targetGroupId === groupId) {
            if (odataType.indexOf("exclusion") !== -1) {
                return { match: true, assignmentType: "Exclude" };
            }
            return { match: true, assignmentType: "Include" };
        }
        return { match: false };
    }

    function buildItem(policy, assignmentInfo) {
        var a = assignmentInfo.source || {};
        return {
            id: policy.id,
            displayName: policy.displayName || policy.name || "Unnamed",
            description: policy.description || "",
            assignmentType: assignmentInfo.assignmentType,
            intent: a.intent || "",
            filterId: (a.target && a.target.deviceAndAppManagementAssignmentFilterId) || "",
            filterType: (a.target && a.target.deviceAndAppManagementAssignmentFilterType) || "none"
        };
    }

    async function getAssignmentsForCategory(url, groupId) {
        var items = await graphFetchAll(url);
        var matched = [];
        items.forEach(function (item) {
            var assignments = item.assignments || [];
            assignments.forEach(function (a) {
                var info = extractAssignment(a, groupId);
                if (info.match) {
                    matched.push(buildItem(item, { assignmentType: info.assignmentType, source: a }));
                }
            });
        });
        return matched;
    }

    async function getAssignmentsForGroup(groupId) {
        var endpoints = {
            configurations: GRAPH_BASE + "/beta/deviceManagement/deviceConfigurations?$expand=assignments&$select=id,displayName,description,assignments",
            settingsCatalog: GRAPH_BASE + "/beta/deviceManagement/configurationPolicies?$expand=assignments&$select=id,name,description,assignments",
            applications: GRAPH_BASE + "/beta/deviceAppManagement/mobileApps?$expand=assignments&$filter=isAssigned eq true&$select=id,displayName,description,assignments",
            scripts: GRAPH_BASE + "/beta/deviceManagement/deviceManagementScripts?$expand=assignments&$select=id,displayName,description,assignments",
            remediations: GRAPH_BASE + "/beta/deviceManagement/deviceHealthScripts?$expand=assignments&$select=id,displayName,description,assignments"
        };

        var keys = Object.keys(endpoints);
        var promises = keys.map(function (key) {
            return getAssignmentsForCategory(endpoints[key], groupId);
        });
        var results = await Promise.all(promises);

        var data = {};
        keys.forEach(function (key, i) {
            data[key] = results[i];
            // Settings catalog uses "name" not "displayName"
            if (key === "settingsCatalog") {
                data[key].forEach(function (item) {
                    if (!item.displayName && item.name) {
                        item.displayName = item.name;
                    }
                });
            }
        });
        return data;
    }

    async function getAssignedGroupIds() {
        var endpoints = [
            GRAPH_BASE + "/beta/deviceManagement/deviceConfigurations?$expand=assignments&$select=id,assignments",
            GRAPH_BASE + "/beta/deviceManagement/configurationPolicies?$expand=assignments&$select=id,assignments",
            GRAPH_BASE + "/beta/deviceAppManagement/mobileApps?$expand=assignments&$filter=isAssigned eq true&$select=id,assignments",
            GRAPH_BASE + "/beta/deviceManagement/deviceManagementScripts?$expand=assignments&$select=id,assignments",
            GRAPH_BASE + "/beta/deviceManagement/deviceHealthScripts?$expand=assignments&$select=id,assignments"
        ];

        var promises = endpoints.map(function (url) {
            return graphFetchAll(url).catch(function () { return []; });
        });
        var allResults = await Promise.all(promises);

        var ids = new Set();
        allResults.forEach(function (items) {
            items.forEach(function (item) {
                (item.assignments || []).forEach(function (a) {
                    var gid = a.target && a.target.groupId;
                    if (gid) ids.add(gid);
                });
            });
        });
        return Array.from(ids);
    }

    async function getScriptContent(scriptId) {
        var data = await graphFetch(
            GRAPH_BASE + "/beta/deviceManagement/deviceManagementScripts/" + scriptId
        );
        var content = "";
        if (data.scriptContent) {
            try {
                content = atob(data.scriptContent);
            } catch (e) {
                content = data.scriptContent;
            }
        }
        return {
            id: data.id,
            fileName: data.fileName || "",
            content: content
        };
    }

    function isInitialised() {
        return msalInstance !== null;
    }

    function getAccount() {
        if (activeAccount) return activeAccount;
        if (msalInstance) {
            var accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) return accounts[0];
        }
        return null;
    }

    // ── Public interface ─────────────────────────────────────────────────

    return {
        init: init,
        handleRedirect: handleRedirect,
        signIn: signIn,
        signOut: signOut,
        isInitialised: isInitialised,
        getAccount: getAccount,
        getAllGroups: getAllGroups,
        getAssignedGroupIds: getAssignedGroupIds,
        getAssignmentsForGroup: getAssignmentsForGroup,
        getScriptContent: getScriptContent
    };
})();
