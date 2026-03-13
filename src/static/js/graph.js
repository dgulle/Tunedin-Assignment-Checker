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
        "DeviceManagementScripts.Read.All",
        "Group.Read.All",
        "User.Read.All"
    ];

    var msalInstance = null;
    var activeAccount = null;
    var initPromise = null; // tracks the full init+handleRedirect sequence

    // ── Initialise MSAL and handle any pending redirect ──────────────────

    function init(clientId, tenantId) {
        // If already initialising/initialised, return the same promise
        if (initPromise) return initPromise;

        initPromise = _doInit(clientId, tenantId);
        return initPromise;
    }

    async function _doInit(clientId, tenantId) {
        var authority = "https://login.microsoftonline.com/" + tenantId;

        var msalConfig = {
            auth: {
                clientId: clientId,
                authority: authority,
                redirectUri: REDIRECT_URI
            },
            cache: {
                cacheLocation: "sessionStorage",
                storeAuthStateInCookie: false
            }
        };

        msalInstance = new msal.PublicClientApplication(msalConfig);

        // initialize() is required in MSAL 2.x before any operations
        if (typeof msalInstance.initialize === "function") {
            await msalInstance.initialize();
        }

        // handleRedirectPromise MUST be called once after init to clear
        // any pending redirect state — otherwise loginRedirect will refuse
        // to start a new interaction ("interaction_in_progress").
        try {
            var response = await msalInstance.handleRedirectPromise();
            if (response && response.account) {
                activeAccount = response.account;
            }
        } catch (err) {
            console.warn("handleRedirectPromise error (clearing state):", err);
            // Clear stale interaction state so future logins work
            _clearInteractionState();
        }

        // Check for existing accounts from cache
        if (!activeAccount) {
            var accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                activeAccount = accounts[0];
            }
        }

        return activeAccount;
    }

    // Clear MSAL's interaction-in-progress flag from sessionStorage
    function _clearInteractionState() {
        try {
            var keys = Object.keys(sessionStorage);
            for (var i = 0; i < keys.length; i++) {
                if (keys[i].indexOf("msal.") === 0 && keys[i].indexOf("interaction") !== -1) {
                    sessionStorage.removeItem(keys[i]);
                }
            }
        } catch (e) {
            // sessionStorage not available
        }
    }

    // ── Sign in ──────────────────────────────────────────────────────────

    async function signIn() {
        if (!msalInstance) throw new Error("MSAL not initialised. Call init() first.");

        // Clear any leftover interaction state before starting a new login
        _clearInteractionState();

        // Use redirect flow — page navigates to Microsoft login, then back
        await msalInstance.loginRedirect({ scopes: SCOPES });
        return null;
    }

    // ── Sign out ─────────────────────────────────────────────────────────

    async function signOut() {
        if (!msalInstance) return;
        var account = activeAccount || (msalInstance.getAllAccounts()[0] || null);
        activeAccount = null;
        initPromise = null;
        if (account) {
            try {
                await msalInstance.logoutRedirect({ account: account });
            } catch (e) {
                msalInstance.clearCache();
            }
        }
    }

    // ── Acquire token silently (with fallback to redirect) ───────────────

    async function getToken() {
        if (!msalInstance) throw new Error("MSAL not initialised.");
        var account = activeAccount || (msalInstance.getAllAccounts()[0] || null);
        if (!account) throw new Error("No signed-in account. Please sign in first.");

        var request = { scopes: SCOPES, account: account };
        try {
            var response = await msalInstance.acquireTokenSilent(request);
            return response.accessToken;
        } catch (err) {
            if (err instanceof msal.InteractionRequiredAuthError) {
                _clearInteractionState();
                await msalInstance.acquireTokenRedirect(request);
                return null;
            }
            throw err;
        }
    }

    // ── Graph API fetch with auth header and retry logic ────────────────

    var MAX_RETRIES = 3;

    async function graphFetch(url) {
        var token = await getToken();
        var attempt = 0;

        while (true) {
            var resp = await fetch(url, {
                headers: { "Authorization": "Bearer " + token }
            });

            // Retry on 429 (throttled) or 5xx (server error)
            if ((resp.status === 429 || resp.status >= 500) && attempt < MAX_RETRIES) {
                attempt++;
                // Use Retry-After header if present, otherwise exponential backoff
                var retryAfter = resp.headers.get("Retry-After");
                var delay = retryAfter ? parseInt(retryAfter, 10) * 1000 : Math.pow(2, attempt) * 1000;
                console.warn("Graph API " + resp.status + " on attempt " + attempt + ", retrying in " + delay + "ms: " + url);
                await new Promise(function (resolve) { setTimeout(resolve, delay); });
                // Refresh token in case it expired during the wait
                token = await getToken();
                continue;
            }

            if (!resp.ok) {
                var body = await resp.json().catch(function () { return {}; });
                var msg = (body.error && body.error.message) || "HTTP " + resp.status;
                throw new Error(msg);
            }
            return resp.json();
        }
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
            // Only follow nextLink if it points to the Graph API (prevent token leakage)
            var link = data["@odata.nextLink"] || null;
            if (link && link.indexOf(GRAPH_BASE) !== 0) {
                console.warn("Ignoring untrusted @odata.nextLink:", link);
                link = null;
            }
            nextUrl = link;
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
            return getAssignmentsForCategory(endpoints[key], groupId)
                .catch(function (err) {
                    console.error("Failed to fetch " + key + ":", err);
                    return { _error: err.message || "Failed to load" };
                });
        });
        var results = await Promise.all(promises);

        var data = { _errors: {} };
        keys.forEach(function (key, i) {
            if (results[i] && results[i]._error) {
                data[key] = [];
                data._errors[key] = results[i]._error;
            } else {
                data[key] = results[i];
            }
            if (key === "settingsCatalog" && Array.isArray(data[key])) {
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

        var counts = {};  // groupId -> assignment count
        allResults.forEach(function (items) {
            items.forEach(function (item) {
                (item.assignments || []).forEach(function (a) {
                    var gid = a.target && a.target.groupId;
                    if (gid) {
                        counts[gid] = (counts[gid] || 0) + 1;
                    }
                });
            });
        });
        return { ids: Object.keys(counts), counts: counts };
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

    // Reset MSAL state so a new tenant/client can be used without a page reload
    function reset() {
        if (msalInstance) {
            try { msalInstance.clearCache(); } catch (e) { /* ignore */ }
        }
        msalInstance = null;
        activeAccount = null;
        initPromise = null;
        // Clear all MSAL keys from sessionStorage
        try {
            var keys = Object.keys(sessionStorage);
            for (var i = 0; i < keys.length; i++) {
                if (keys[i].indexOf("msal.") === 0 || keys[i].indexOf("msal:") === 0) {
                    sessionStorage.removeItem(keys[i]);
                }
            }
        } catch (e) { /* sessionStorage not available */ }
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

    return {
        init: init,
        reset: reset,
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
