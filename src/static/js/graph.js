/* ═══════════════════════════════════════════════════════════════════════════
   Tunedin Assignment Checker — SPA Graph Client (MSAL.js)
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

    var GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

    function validateGuid(id, label) {
        if (!id || !GUID_RE.test(id)) {
            throw new Error("Invalid " + (label || "ID") + " format. Expected a valid GUID.");
        }
    }

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
    var FETCH_TIMEOUT_MS = 30000;

    async function graphFetch(url) {
        var token = await getToken();
        var attempt = 0;

        while (true) {
            var controller = new AbortController();
            var timeoutId = setTimeout(function () { controller.abort(); }, FETCH_TIMEOUT_MS);

            var resp;
            try {
                resp = await fetch(url, {
                    headers: { "Authorization": "Bearer " + token },
                    signal: controller.signal
                });
            } catch (err) {
                clearTimeout(timeoutId);
                if (err.name === "AbortError") {
                    throw new Error("Request timed out after " + (FETCH_TIMEOUT_MS / 1000) + "s");
                }
                throw err;
            }
            clearTimeout(timeoutId);

            // Retry on 429 (throttled) or 5xx (server error)
            if ((resp.status === 429 || resp.status >= 500) && attempt < MAX_RETRIES) {
                attempt++;
                // Use Retry-After header if present, otherwise exponential backoff
                var retryAfter = resp.headers.get("Retry-After");
                var delay = retryAfter ? parseInt(retryAfter, 10) * 1000 : Math.pow(2, attempt) * 1000;
                console.warn("Graph API " + resp.status + " on attempt " + attempt + ", retrying in " + delay + "ms");
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

            // Validate response Content-Type before parsing
            var contentType = (resp.headers.get("Content-Type") || "").split(";")[0].trim();
            if (contentType !== "application/json") {
                throw new Error("Unexpected response type: " + contentType);
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
            if (link) {
                try {
                    var parsed = new URL(link);
                    if (parsed.protocol !== "https:" || parsed.hostname !== "graph.microsoft.com") {
                        console.warn("Ignoring untrusted @odata.nextLink");
                        link = null;
                    }
                } catch (e) {
                    console.warn("Ignoring malformed @odata.nextLink");
                    link = null;
                }
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

    async function getAssignmentsByTargetType(targetOdataType, label) {
        var endpoints = {
            configurations: GRAPH_BASE + "/beta/deviceManagement/deviceConfigurations?$expand=assignments&$select=id,displayName,description,assignments",
            settingsCatalog: GRAPH_BASE + "/beta/deviceManagement/configurationPolicies?$expand=assignments&$select=id,name,description,assignments",
            applications: GRAPH_BASE + "/beta/deviceAppManagement/mobileApps?$expand=assignments&$filter=isAssigned eq true&$select=id,displayName,description,assignments",
            scripts: GRAPH_BASE + "/beta/deviceManagement/deviceManagementScripts?$expand=assignments&$select=id,displayName,description,assignments",
            remediations: GRAPH_BASE + "/beta/deviceManagement/deviceHealthScripts?$expand=assignments&$select=id,displayName,description,assignments"
        };

        var keys = Object.keys(endpoints);
        var promises = keys.map(function (key) {
            return graphFetchAll(endpoints[key]).then(function (items) {
                var matched = [];
                items.forEach(function (item) {
                    (item.assignments || []).forEach(function (a) {
                        var t = (a.target && a.target["@odata.type"]) || "";
                        if (t === targetOdataType) {
                            matched.push(buildItem(item, { assignmentType: label, source: a }));
                        }
                    });
                });
                return matched;
            }).catch(function (err) {
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

    async function getAssignmentsForGroup(groupId) {
        validateGuid(groupId, "group ID");
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

    async function getGroupParents(groupId) {
        validateGuid(groupId, "group ID");
        var url = GRAPH_BASE + "/v1.0/groups/" + groupId + "/transitiveMemberOf/microsoft.graph.group?$select=id,displayName&$top=999";
        return graphFetchAll(url);
    }

    async function getNestedAssignments(groupId) {
        validateGuid(groupId, "group ID");
        var parents = await getGroupParents(groupId);
        if (!parents || parents.length === 0) {
            return {
                configurations: [], settingsCatalog: [], applications: [],
                scripts: [], remediations: [], _errors: {}
            };
        }

        // Build lookup of parent group IDs to names
        var parentLookup = {};
        parents.forEach(function (p) {
            if (p.id) parentLookup[p.id] = p.displayName || p.id;
        });

        var endpoints = {
            configurations: GRAPH_BASE + "/beta/deviceManagement/deviceConfigurations?$expand=assignments&$select=id,displayName,description,assignments",
            settingsCatalog: GRAPH_BASE + "/beta/deviceManagement/configurationPolicies?$expand=assignments&$select=id,name,description,assignments",
            applications: GRAPH_BASE + "/beta/deviceAppManagement/mobileApps?$expand=assignments&$filter=isAssigned eq true&$select=id,displayName,description,assignments",
            scripts: GRAPH_BASE + "/beta/deviceManagement/deviceManagementScripts?$expand=assignments&$select=id,displayName,description,assignments",
            remediations: GRAPH_BASE + "/beta/deviceManagement/deviceHealthScripts?$expand=assignments&$select=id,displayName,description,assignments"
        };

        var keys = Object.keys(endpoints);
        var promises = keys.map(function (key) {
            return graphFetchAll(endpoints[key]).then(function (items) {
                var matched = [];
                items.forEach(function (item) {
                    (item.assignments || []).forEach(function (a) {
                        var t = a.target || {};
                        var tGroupId = t.groupId || null;
                        if (tGroupId && parentLookup[tGroupId]) {
                            var odataType = t["@odata.type"] || "";
                            var assignmentType = odataType.indexOf("exclusion") !== -1 ? "Exclude" : "Include";
                            var built = buildItem(item, { assignmentType: assignmentType, source: a });
                            built.inheritedFrom = parentLookup[tGroupId];
                            built.inheritedFromId = tGroupId;
                            matched.push(built);
                        }
                    });
                });
                return matched;
            }).catch(function (err) {
                console.error("Failed to fetch nested " + key + ":", err);
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
                    if (!item.displayName && item.name) item.displayName = item.name;
                });
            }
        });
        return data;
    }

    async function getOrphanedItems() {
        var endpoints = {
            configurations: GRAPH_BASE + "/beta/deviceManagement/deviceConfigurations?$expand=assignments&$select=id,displayName,description,assignments",
            settingsCatalog: GRAPH_BASE + "/beta/deviceManagement/configurationPolicies?$expand=assignments&$select=id,name,description,assignments",
            applications: GRAPH_BASE + "/beta/deviceAppManagement/mobileApps?$expand=assignments&$select=id,displayName,description,assignments",
            scripts: GRAPH_BASE + "/beta/deviceManagement/deviceManagementScripts?$expand=assignments&$select=id,displayName,description,assignments",
            remediations: GRAPH_BASE + "/beta/deviceManagement/deviceHealthScripts?$expand=assignments&$select=id,displayName,description,assignments"
        };

        var keys = Object.keys(endpoints);
        var promises = keys.map(function (key) {
            return graphFetchAll(endpoints[key]).then(function (items) {
                var orphaned = [];
                items.forEach(function (item) {
                    var assignments = item.assignments || [];
                    if (assignments.length === 0) {
                        orphaned.push({
                            id: item.id,
                            displayName: item.displayName || item.name || "Unnamed",
                            description: item.description || ""
                        });
                    }
                });
                return orphaned;
            }).catch(function (err) {
                console.error("Failed to fetch orphaned " + key + ":", err);
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
                    if (!item.displayName && item.name) item.displayName = item.name;
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
        validateGuid(scriptId, "script ID");
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
        getAssignmentsByTargetType: getAssignmentsByTargetType,
        getScriptContent: getScriptContent,
        getGroupParents: getGroupParents,
        getNestedAssignments: getNestedAssignments,
        getOrphanedItems: getOrphanedItems
    };
})();
