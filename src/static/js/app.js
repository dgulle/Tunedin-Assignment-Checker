/* ═══════════════════════════════════════════════════════════════════════════
   Tunedin Assignment Checker — Frontend (ZeroToTrust Edition)
   Supports dual mode:
     - Backend mode: PowerShell HTTP server at /api/*
     - SPA mode: MSAL.js + direct Graph API calls (for GitHub Pages)
   ═══════════════════════════════════════════════════════════════════════════ */

(function () {
    "use strict";

    // ── DOM references ──────────────────────────────────────────────────
    var groupList       = document.getElementById("groupList");
    var groupSearch     = document.getElementById("groupSearch");
    var groupCount      = document.getElementById("groupCount");
    var sidebarLoading  = document.getElementById("sidebarLoading");
    var sidebarError    = document.getElementById("sidebarError");
    var sidebarErrorMsg = document.getElementById("sidebarErrorMsg");

    var emptyState      = document.getElementById("emptyState");
    var contentLoading  = document.getElementById("contentLoading");
    var contentError    = document.getElementById("contentError");
    var contentErrorMsg = document.getElementById("contentErrorMsg");
    var assignments     = document.getElementById("assignments");

    var selectedGroupName    = document.getElementById("selectedGroupName");
    var selectedGroupDesc    = document.getElementById("selectedGroupDesc");
    var membershipRuleEl     = document.getElementById("membershipRule");
    var membershipRuleQuery  = document.getElementById("membershipRuleQuery");
    var categoryTabs      = document.getElementById("categoryTabs");
    var cardGrid          = document.getElementById("cardGrid");
    var categoryEmpty     = document.getElementById("categoryEmpty");
    var platformFilter    = document.getElementById("platformFilter");

    var connectionBadge = document.getElementById("connectionBadge");
    var badgeDot        = connectionBadge.querySelector(".badge-dot");
    var badgeText       = connectionBadge.querySelector(".badge-text");

    var scriptModal      = document.getElementById("scriptModal");
    var scriptModalTitle = document.getElementById("scriptModalTitle");
    var scriptModalFile  = document.getElementById("scriptModalFile");
    var scriptModalBody  = document.getElementById("scriptModalBody");
    var btnCopyScript    = document.getElementById("btnCopyScript");
    var copyBtnLabel     = document.getElementById("copyBtnLabel");

    var btnGroupFilter = document.getElementById("btnGroupFilter");

    // Setup screen elements
    var setupScreen   = document.getElementById("setupScreen");
    var appHeader     = document.querySelector(".app-header");
    var appLayout     = document.querySelector(".app-layout");

    // ── State ───────────────────────────────────────────────────────────
    var allGroups          = [];
    var assignedGroupIds   = new Set();
    var groupAssignCounts  = {};   // groupId -> total assignment count
    var groupMemberCounts  = {};   // groupId -> member count (populated lazily)
    var memberCountsToken  = 0;    // invalidates in-flight fetches when groups reload
    var filterAssigned     = true;
    var filterMinCount     = 0;    // min assignment count filter (0 = no filter)
    var filterMaxCount     = 0;    // max assignment count filter (0 = no filter)
    var showAllDevices     = true; // toggle for All Devices assignments
    var showAllUsers       = true; // toggle for All Users assignments
    var activeGroupId      = null;
    var assignmentData     = null;
    var nestedData         = null;  // nested/inherited assignments via parent groups
    var orphanedData       = null;  // orphaned items (no assignments)
    var showNested         = true;  // toggle for nested group assignments
    var activeCategory     = "configurations";
    var activePlatformFilter = null; // null = All, or "Android"/"iOS"/"Windows"/"macOS"

    // Mode: "backend" or "spa"
    var appMode = "backend";

    // Synthetic groups for All Devices / All Users / Orphaned. Defined here
    // (not later in the file) so renderGroupList can never read it before
    // its initializer has run — an earlier definition lower in the IIFE
    // manifested as "Cannot read properties of undefined (reading 'filter')"
    // when renderGroupList fired from an async path.
    var SYNTHETIC_GROUPS = [
        { id: "__allDevices__", displayName: "All Devices", description: "Policies and apps assigned to all devices", _synthetic: true },
        { id: "__allUsers__",   displayName: "All Users",   description: "Policies and apps assigned to all licensed users", _synthetic: true },
        { id: "__orphaned__",   displayName: "Orphaned Items", description: "Items with no assignments — review for deletion", _synthetic: true }
    ];

    // Per-launch backend API secret. The PowerShell script opens the
    // browser at http://localhost:PORT/#k=<secret>. We read the fragment
    // once at startup, scrub it from the address bar, and attach it as
    // X-Backend-Key on every /api/* request. Without this key the local
    // HTTP listener returns 401 for all API routes — so other processes
    // running as the same user cannot piggyback on the Graph session.
    var _backendKey = null;
    (function extractBackendKey() {
        try {
            var m = /^#k=([A-Za-z0-9_\-]+)/.exec(window.location.hash || "");
            if (m) {
                _backendKey = m[1];
                window.history.replaceState(null, "",
                    window.location.pathname + window.location.search);
            }
        } catch (e) { /* ignore */ }
    })();

    // ── SPA session inactivity timeout ──────────────────────────────────
    var SPA_IDLE_TIMEOUT_MS = 30 * 60 * 1000; // 30 minutes
    var _idleTimer = null;

    function resetIdleTimer() {
        if (appMode !== "spa" || !_idleTimer) return;
        clearTimeout(_idleTimer);
        _idleTimer = setTimeout(onIdleTimeout, SPA_IDLE_TIMEOUT_MS);
    }

    function startIdleTimer() {
        if (appMode !== "spa") return;
        _idleTimer = setTimeout(onIdleTimeout, SPA_IDLE_TIMEOUT_MS);
        ["mousemove", "keydown", "click", "scroll", "touchstart"].forEach(function (evt) {
            document.addEventListener(evt, resetIdleTimer, { passive: true });
        });
    }

    function stopIdleTimer() {
        if (_idleTimer) { clearTimeout(_idleTimer); _idleTimer = null; }
        ["mousemove", "keydown", "click", "scroll", "touchstart"].forEach(function (evt) {
            document.removeEventListener(evt, resetIdleTimer);
        });
    }

    function onIdleTimeout() {
        stopIdleTimer();
        alert("Your session has expired due to inactivity. Please sign in again.");
        logout();
    }

    // ── Boot ────────────────────────────────────────────────────────────
    // Hide main app immediately to prevent flash while detectMode runs
    appHeader.style.display = "none";
    appLayout.style.display = "none";

    initTheme();
    detectMode();

    groupSearch.addEventListener("input", function () { renderGroupList(); });
    document.getElementById("btnRetry").addEventListener("click", function () { loadGroups(); });
    document.getElementById("btnLogout").addEventListener("click", logout);
    document.getElementById("btnTheme").addEventListener("click", toggleTheme);
    document.getElementById("btnModalClose").addEventListener("click", closeScriptModal);
    if (btnCopyScript) btnCopyScript.addEventListener("click", copyScriptContent);
    btnGroupFilter.addEventListener("click", toggleGroupFilter);

    // Count filter controls
    document.getElementById("btnCountFilter").addEventListener("click", function () {
        var panel = document.getElementById("countFilterPanel");
        panel.style.display = panel.style.display === "none" ? "block" : "none";
    });
    document.getElementById("btnCountApply").addEventListener("click", function () {
        filterMinCount = parseInt(document.getElementById("filterMinCount").value, 10) || 0;
        filterMaxCount = parseInt(document.getElementById("filterMaxCount").value, 10) || 0;
        document.getElementById("countFilterPanel").style.display = "none";
        var btn = document.getElementById("btnCountFilter");
        btn.classList.toggle("active", filterMinCount > 0 || filterMaxCount > 0);
        renderGroupList();
    });
    document.getElementById("btnCountClear").addEventListener("click", function () {
        filterMinCount = 0;
        filterMaxCount = 0;
        document.getElementById("filterMinCount").value = "";
        document.getElementById("filterMaxCount").value = "";
        document.getElementById("countFilterPanel").style.display = "none";
        document.getElementById("btnCountFilter").classList.remove("active");
        renderGroupList();
    });
    // Null-guarded: if this element is missing (e.g. cached old index.html)
    // don't let a null throw abort the rest of init — that would break the
    // setup screen's Sign In button wiring further down.
    var _btnCountExport = document.getElementById("btnCountExport");
    if (_btnCountExport) _btnCountExport.addEventListener("click", exportFilteredGroupsCsv);

    // All Devices / All Users toggles
    document.getElementById("btnShowAllDevices").addEventListener("click", function () {
        showAllDevices = !showAllDevices;
        this.classList.toggle("active", showAllDevices);
        updateCounts();
        renderCards();
    });
    document.getElementById("btnShowAllUsers").addEventListener("click", function () {
        showAllUsers = !showAllUsers;
        this.classList.toggle("active", showAllUsers);
        updateCounts();
        renderCards();
    });

    document.getElementById("btnShowNested").addEventListener("click", function () {
        showNested = !showNested;
        this.classList.toggle("active", showNested);
        updateCounts();
        renderCards();
    });

    // Export CSV
    document.getElementById("btnExportCsv").addEventListener("click", exportCsv);

    scriptModal.addEventListener("click", function (e) {
        if (e.target === scriptModal) closeScriptModal();
    });

    categoryTabs.addEventListener("click", function (e) {
        var tab = e.target.closest(".tab");
        if (!tab) return;
        activeCategory = tab.dataset.category;
        highlightTab();
        if (activeGroupId === "__orphaned__") {
            updateOrphanedCounts();
            renderOrphanedCards();
        } else {
            renderCards();
        }
    });

    platformFilter.addEventListener("click", function (e) {
        var btn = e.target.closest(".platform-btn");
        if (!btn) return;
        activePlatformFilter = btn.dataset.platform || null;
        platformFilter.querySelectorAll(".platform-btn").forEach(function (b) {
            b.classList.toggle("active", b === btn);
        });
        updateOrphanedCounts();
        renderOrphanedCards();
    });

    // Setup screen connect button
    document.getElementById("btnSetupConnect").addEventListener("click", setupConnect);

    // ── Mode Detection ──────────────────────────────────────────────────

    async function detectMode() {
        // Try to reach the PowerShell backend via the /api/status endpoint.
        // If no backend key is present in the URL fragment, skip the probe
        // entirely — the backend will reject us anyway, and the request
        // might come from a misrouted SPA-mode load.
        if (_backendKey) {
            try {
                var resp = await fetch("/api/status", {
                    headers: { "X-Backend-Key": _backendKey }
                });
                if (resp.ok) {
                    appMode = "backend";
                    showApp();
                    loadGroups();
                    return;
                }
            } catch (e) {
                // Backend not available — fall through to SPA mode
            }
        }

        // SPA mode
        appMode = "spa";

        // Security: warn if SPA mode is running over HTTP (excluding localhost)
        if (window.location.protocol === "http:" &&
            window.location.hostname !== "localhost" &&
            window.location.hostname !== "127.0.0.1") {
            console.warn("Security warning: running over HTTP. OAuth tokens may be intercepted. Use HTTPS.");
            alert("Security warning: this page is served over HTTP. Your authentication tokens could be intercepted by an attacker. Please use HTTPS.");
        }

        var savedClientId = localStorage.getItem("iac_clientId");
        var savedTenantId = localStorage.getItem("iac_tenantId");

        if (savedClientId && savedTenantId) {
            try {
                // Init MSAL, handle any pending redirect, check for cached account
                var account = await GraphClient.init(savedClientId, savedTenantId);
                if (account) {
                    showApp();
                    setConnection("connected", account.username || "Connected");
                    startIdleTimer();
                    loadGroups();
                    return;
                }
            } catch (err) {
                console.error("MSAL init error:", err);
            }
        }

        // Show setup screen
        showSetup(savedClientId, savedTenantId);
    }

    // ── Setup Screen ────────────────────────────────────────────────────

    function showSetup(savedClientId, savedTenantId) {
        setupScreen.style.display = "flex";
        appHeader.style.display   = "none";
        appLayout.style.display   = "none";

        // Pre-fill saved values
        var tenantInput  = document.getElementById("setupTenantId");
        var clientInput  = document.getElementById("setupClientId");
        if (savedTenantId) tenantInput.value = savedTenantId;
        if (savedClientId) clientInput.value = savedClientId;
    }

    function hideSetup() {
        setupScreen.style.display = "none";
        appHeader.style.display   = "";
        appLayout.style.display   = "";
    }

    function showApp() {
        hideSetup();
        appHeader.style.display = "";
        appLayout.style.display = "";
    }

    var GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    var DOMAIN_RE = /^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$/i;

    async function setupConnect() {
        var tenantId = document.getElementById("setupTenantId").value.trim();
        var clientId = document.getElementById("setupClientId").value.trim();

        if (!tenantId || !clientId) {
            alert("Please enter both Tenant ID and Client ID.");
            return;
        }

        // Length limits to prevent abuse
        if (tenantId.length > 253 || clientId.length > 36) {
            alert("Tenant ID must be at most 253 characters and Client ID must be at most 36 characters.");
            return;
        }

        // Validate Client ID is a GUID
        if (!GUID_RE.test(clientId)) {
            alert("Client ID must be a valid GUID (e.g. 12345678-abcd-1234-abcd-123456789abc).");
            return;
        }

        // Validate Tenant ID is a GUID or a domain name
        if (!GUID_RE.test(tenantId) && !DOMAIN_RE.test(tenantId)) {
            alert("Tenant ID must be a valid GUID or domain (e.g. contoso.onmicrosoft.com).");
            return;
        }

        // If tenant or client changed, reset MSAL so a fresh instance is created
        var prevTenant = localStorage.getItem("iac_tenantId");
        var prevClient = localStorage.getItem("iac_clientId");
        if (GraphClient.isInitialised() && (prevTenant !== tenantId || prevClient !== clientId)) {
            GraphClient.reset();
        }

        // Save to localStorage so we can pick up after redirect
        localStorage.setItem("iac_tenantId", tenantId);
        localStorage.setItem("iac_clientId", clientId);

        try {
            // Init MSAL (idempotent — reuses if already initialised)
            await GraphClient.init(clientId, tenantId);

            // Redirect to Microsoft login. On return, detectMode() picks up the token.
            await GraphClient.signIn();
        } catch (err) {
            console.error("Sign-in failed:", err);
            alert("Sign-in failed: " + (err.message || err));
        }
    }

    // ── Dark mode ─────────────────────────────────────────────────────

    function initTheme() {
        var saved = localStorage.getItem("theme");
        if (saved === "dark" || (!saved && window.matchMedia("(prefers-color-scheme: dark)").matches)) {
            document.documentElement.setAttribute("data-theme", "dark");
        }
        updateThemeIcon();
    }

    function toggleTheme() {
        var isDark = document.documentElement.getAttribute("data-theme") === "dark";
        if (isDark) {
            document.documentElement.removeAttribute("data-theme");
            localStorage.setItem("theme", "light");
        } else {
            document.documentElement.setAttribute("data-theme", "dark");
            localStorage.setItem("theme", "dark");
        }
        updateThemeIcon();
    }

    function updateThemeIcon() {
        var isDark = document.documentElement.getAttribute("data-theme") === "dark";
        document.getElementById("iconSun").style.display  = isDark ? "block" : "none";
        document.getElementById("iconMoon").style.display = isDark ? "none"  : "block";
    }

    // ── API helpers ─────────────────────────────────────────────────────

    async function apiFetch(url, options) {
        options = options || {};
        var headers = {};
        // Merge caller-provided headers without mutating their object
        if (options.headers) {
            Object.keys(options.headers).forEach(function (k) { headers[k] = options.headers[k]; });
        }
        if (_backendKey) headers["X-Backend-Key"] = _backendKey;
        var fetchInit = {
            method: options.method || "GET",
            headers: headers
        };
        if (options.body !== undefined) fetchInit.body = options.body;
        var resp = await fetch(url, fetchInit);
        if (!resp.ok) {
            var body = await resp.json().catch(function () { return {}; });
            if (resp.status === 401 && body && body.expired) {
                onBackendSessionExpired();
            }
            throw new Error(body.error || "HTTP " + resp.status);
        }
        return resp.json();
    }

    var _backendSessionExpiredShown = false;
    function onBackendSessionExpired() {
        if (_backendSessionExpiredShown) return;
        _backendSessionExpiredShown = true;
        stopIdleTimer();
        setConnection("error", "Session expired");
        alert("Your backend session expired due to inactivity. " +
              "Please close this tab and restart the script to sign in again.");
    }

    // ── Load groups ─────────────────────────────────────────────────────

    async function loadGroups() {
        sidebarLoading.style.display = "flex";
        sidebarError.style.display   = "none";
        groupList.innerHTML          = "";

        try {
            var groups, assignedIds;

            if (appMode === "spa") {
                var results = await Promise.all([
                    GraphClient.getAllGroups(),
                    GraphClient.getAssignedGroupIds().catch(function () { return { ids: [], counts: {} }; })
                ]);
                groups = results[0];
                assignedIds = results[1].ids || results[1];
                groupAssignCounts = results[1].counts || {};
            } else {
                var backendResults = await Promise.all([
                    apiFetch("/api/groups"),
                    apiFetch("/api/assigned-group-ids").catch(function () { return []; })
                ]);
                groups = backendResults[0];
                var backendIdData = backendResults[1];
                assignedIds = backendIdData.ids || backendIdData;
                groupAssignCounts = backendIdData.counts || {};
            }

            if (!Array.isArray(groups)) {
                console.error("Unexpected /api/groups response (not an array):", groups);
                throw new Error("Server returned an unexpected groups response. Check the PowerShell console for errors.");
            }
            allGroups = groups;
            assignedGroupIds = new Set(Array.isArray(assignedIds) ? assignedIds : Object.keys(groupAssignCounts));
            // Reset any previously-fetched counts so a reload doesn't show
            // stale data while the new batch is in flight.
            groupMemberCounts = {};
            renderGroupList();
            setConnection("connected", appMode === "spa" ? (GraphClient.getAccount()?.username || "Connected") : "Connected");
            // Fetch member counts in the background — the list is already
            // usable and counts stream in as batches complete.
            loadGroupMemberCounts();
        } catch (err) {
            console.error("Failed to load groups:", err);
            sidebarErrorMsg.textContent = "Failed to load groups. " + (err.message || "Please check your connection and try again.");
            sidebarError.style.display  = "flex";
            setConnection("error", "Disconnected");
        } finally {
            sidebarLoading.style.display = "none";
        }
    }

    // ── Filter logic shared by renderGroupList and the CSV exporter ─────
    //
    // Kept in one place so the sidebar list and the export always agree on
    // which groups are "in view" — search box, assigned-only toggle, and
    // the min/max assignment-count range are all applied here.

    function getFilteredGroups() {
        var query = groupSearch.value.trim().toLowerCase();
        var filtered = allGroups;

        if (filterAssigned && assignedGroupIds.size > 0) {
            filtered = filtered.filter(function (g) { return assignedGroupIds.has(g.id); });
        }

        if (filterMinCount > 0 || filterMaxCount > 0) {
            filtered = filtered.filter(function (g) {
                var cnt = groupAssignCounts[g.id] || 0;
                if (filterMinCount > 0 && cnt < filterMinCount) return false;
                if (filterMaxCount > 0 && cnt > filterMaxCount) return false;
                return true;
            });
        }

        if (query) {
            filtered = filtered.filter(function (g) {
                return (g.displayName || "").toLowerCase().indexOf(query) !== -1 ||
                       (g.description || "").toLowerCase().indexOf(query) !== -1;
            });
        }

        return filtered;
    }

    // ── Render group list (with search filter) ──────────────────────────

    function renderGroupList() {
        var query = groupSearch.value.trim().toLowerCase();
        var filtered = getFilteredGroups();

        groupList.innerHTML = "";

        filtered.forEach(function (g) {
            var li = document.createElement("li");
            li.className = "group-item" + (g.id === activeGroupId ? " active" : "");
            li.dataset.id = g.id;

            var groupType = getGroupType(g);
            var assignCount = groupAssignCounts[g.id] || 0;
            var gName = g.displayName || "Unnamed Group";
            var memberBadge = buildMemberBadge(g.id);

            li.innerHTML =
                '<div class="group-item-header">' +
                    '<div class="group-item-name" title="' + escapeHtml(gName) + '">' + escapeHtml(gName) + '</div>' +
                    '<div class="group-item-copy"></div>' +
                '</div>' +
                (g.description ? '<div class="group-item-desc" title="' + escapeHtml(g.description) + '">' + escapeHtml(g.description) + '</div>' : '') +
                '<div class="group-item-badges">' +
                    '<span class="group-item-type">' + escapeHtml(groupType) + '</span>' +
                    (assignCount > 0 ? '<span class="group-item-count" title="Total assignments">' + assignCount + ' assignment' + (assignCount !== 1 ? 's' : '') + '</span>' : '') +
                    memberBadge +
                '</div>';

            li.querySelector(".group-item-copy").appendChild(createCopyButton(function () { return gName; }));
            li.addEventListener("click", function () { selectGroup(g); });
            groupList.appendChild(li);
        });

        // Render synthetic groups in the sticky bottom section
        var stickyList = document.getElementById("groupListSticky");
        if (stickyList) {
            stickyList.innerHTML = "";
            var syntheticFiltered = SYNTHETIC_GROUPS.filter(function (g) {
                return !query || (g.displayName || "").toLowerCase().indexOf(query) !== -1;
            });

            syntheticFiltered.forEach(function (g) {
                var li = document.createElement("li");
                li.className = "group-item group-item-synthetic" + (g.id === activeGroupId ? " active" : "");
                li.dataset.id = g.id;

                var icon;
                if (g.id === "__allDevices__") {
                    icon = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="3" width="20" height="14" rx="2" ry="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></svg>';
                } else if (g.id === "__orphaned__") {
                    icon = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>';
                } else {
                    icon = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>';
                }

                li.innerHTML =
                    '<div class="group-item-name">' + icon + ' ' + escapeHtml(g.displayName) + '</div>';

                li.addEventListener("click", function () { selectGroup(g); });
                stickyList.appendChild(li);
            });
        }

        groupCount.textContent = filtered.length;
    }

    function getGroupType(group) {
        var types = group.groupTypes || [];
        if (types.indexOf("DynamicMembership") !== -1) return "Dynamic";
        if (group.membershipRule) return "Dynamic";
        return "Assigned";
    }

    // ── Member count badge ──────────────────────────────────────────────
    //
    // Returns the HTML for a group's member-count badge. When the count
    // hasn't been resolved yet we render a dimmed placeholder so users can
    // see that counts are still loading instead of a blank space.

    function buildMemberBadge(groupId) {
        if (Object.prototype.hasOwnProperty.call(groupMemberCounts, groupId)) {
            var n = groupMemberCounts[groupId];
            var cls = "group-item-members" + (n === 0 ? " group-item-members-empty" : "");
            var label = n + " member" + (n !== 1 ? "s" : "");
            return '<span class="' + cls + '" title="Group members">' + label + '</span>';
        }
        return '<span class="group-item-members group-item-members-loading" title="Loading member count…">…</span>';
    }

    // Kick off background fetching of member counts for every loaded group.
    // Uses a token to cancel stale fetches if the user reloads before this
    // completes. We fetch for ALL groups so users can scroll the unfiltered
    // list and still see counts; the Graph $batch endpoint keeps this
    // reasonably cheap.

    async function loadGroupMemberCounts() {
        memberCountsToken += 1;
        var myToken = memberCountsToken;
        var ids = allGroups.map(function (g) { return g.id; }).filter(Boolean);
        if (ids.length === 0) return;

        function applyPartial(partial) {
            if (myToken !== memberCountsToken) return;
            Object.keys(partial).forEach(function (k) {
                groupMemberCounts[k] = partial[k];
            });
            renderGroupList();
        }

        try {
            if (appMode === "spa") {
                await GraphClient.getGroupMemberCounts(ids, applyPartial);
            } else {
                // Backend fetches server-side in 20-group batches.
                var data = await apiFetch("/api/group-member-counts", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ ids: ids })
                });
                if (myToken !== memberCountsToken) return;
                applyPartial((data && data.counts) || {});
            }
        } catch (err) {
            console.warn("Failed to load member counts:", err);
        }
    }

    // ── Group filter toggle ─────────────────────────────────────────────

    function toggleGroupFilter() {
        filterAssigned = !filterAssigned;
        btnGroupFilter.classList.toggle("active", filterAssigned);
        btnGroupFilter.title = filterAssigned
            ? "Showing groups with assignments \u2014 click to show all"
            : "Showing all groups \u2014 click to filter to assigned only";
        renderGroupList();
    }

    // ── Populate group header with copy buttons and dynamic rule ───────

    function populateGroupHeader(group) {
        var name = group.displayName || "Unnamed Group";
        var desc = group.description || "";
        var rule = group.membershipRule || "";

        // Group name with copy button
        selectedGroupName.textContent = name;
        // Remove any previous copy button
        var existingCopy = selectedGroupName.parentNode.querySelector(".btn-copy-name");
        if (existingCopy) existingCopy.remove();
        var nameCopy = createCopyButton(function () { return name; });
        nameCopy.classList.add("btn-copy-name");
        selectedGroupName.parentNode.insertBefore(nameCopy, selectedGroupName.nextSibling);

        // Description with copy button
        selectedGroupDesc.textContent = desc;
        var existingDescCopy = selectedGroupDesc.parentNode.querySelector(".btn-copy-desc");
        if (existingDescCopy) existingDescCopy.remove();
        if (desc) {
            var descCopy = createCopyButton(function () { return desc; });
            descCopy.classList.add("btn-copy-desc");
            selectedGroupDesc.parentNode.insertBefore(descCopy, selectedGroupDesc.nextSibling);
        }

        // Dynamic membership rule
        if (membershipRuleEl) {
            if (rule) {
                membershipRuleQuery.textContent = rule;
                membershipRuleEl.style.display = "";
                // Copy button for rule
                var existingRuleCopy = membershipRuleEl.querySelector(".btn-copy");
                if (existingRuleCopy) existingRuleCopy.remove();
                membershipRuleEl.appendChild(createCopyButton(function () { return rule; }));
            } else {
                membershipRuleEl.style.display = "none";
            }
        }
    }

    // ── Select a group ──────────────────────────────────────────────────

    async function selectGroup(group) {
        activeGroupId = group.id;
        nestedData = null;
        orphanedData = null;
        // Show/hide the Groups tab (only relevant in Orphaned Items view)
        var groupsTab = categoryTabs.querySelector('[data-category="groups"]');
        if (groupsTab) groupsTab.style.display = (group.id === "__orphaned__") ? "" : "none";

        if (group.id !== "__orphaned__") {
            activePlatformFilter = null;
            platformFilter.style.display = "none";
            platformFilter.querySelectorAll(".platform-btn").forEach(function (b) {
                b.classList.toggle("active", !b.dataset.platform);
            });
        }
        renderGroupList();
        showPanel("loading");

        try {
            if (group._synthetic && group.id === "__orphaned__") {
                // Orphaned items view
                if (appMode === "spa") {
                    orphanedData = await GraphClient.getOrphanedItems(allGroups, assignedGroupIds);
                } else {
                    orphanedData = await apiFetch("/api/orphaned-items");
                }
                assignmentData = null;
                populateGroupHeader(group);
                updateOrphanedCounts();
                activeCategory = getFirstNonEmptyOrphanedCategory() || "configurations";
                highlightTab();
                platformFilter.style.display = "flex";
                renderOrphanedCards();
                showPanel("assignments");
                return;
            } else if (group._synthetic) {
                // Synthetic group: fetch by target type
                var targetType = group.id === "__allDevices__"
                    ? "#microsoft.graph.allDevicesAssignmentTarget"
                    : "#microsoft.graph.allLicensedUsersAssignmentTarget";
                var label = group.id === "__allDevices__" ? "All Devices" : "All Users";

                if (appMode === "spa") {
                    assignmentData = await GraphClient.getAssignmentsByTargetType(targetType, label);
                } else {
                    assignmentData = await apiFetch("/api/assignments-by-target?type=" + encodeURIComponent(targetType));
                }
            } else if (appMode === "spa") {
                var groupResults = await Promise.all([
                    GraphClient.getAssignmentsForGroup(group.id),
                    GraphClient.getNestedAssignments(group.id).catch(function () { return null; })
                ]);
                assignmentData = groupResults[0];
                nestedData = groupResults[1];
            } else {
                var backendGroupResults = await Promise.all([
                    apiFetch("/api/groups/" + group.id + "/assignments"),
                    apiFetch("/api/groups/" + group.id + "/nested-assignments").catch(function () { return null; })
                ]);
                assignmentData = backendGroupResults[0];
                nestedData = backendGroupResults[1];
            }

            populateGroupHeader(group);
            updateCounts();
            activeCategory = getFirstNonEmptyCategory() || "configurations";
            highlightTab();
            renderCards();
            showPanel("assignments");
        } catch (err) {
            console.error("Failed to load assignments:", err);
            // Try to show partial results if we got any data at all
            if (assignmentData) {
                populateGroupHeader(group);
                updateCounts();
                activeCategory = getFirstNonEmptyCategory() || "configurations";
                highlightTab();
                renderCards();
                showPanel("assignments");
            } else {
                contentErrorMsg.textContent = err.message || "Failed to load assignments.";
                showPanel("error");
            }
        }
    }

    // ── Panel visibility ────────────────────────────────────────────────

    function showPanel(panel) {
        emptyState.style.display     = panel === "empty"       ? "flex" : "none";
        contentLoading.style.display = panel === "loading"     ? "flex" : "none";
        contentError.style.display   = panel === "error"       ? "flex" : "none";
        assignments.style.display    = panel === "assignments" ? "block" : "none";
    }

    // ── Tabs & counts ───────────────────────────────────────────────────

    var CATEGORIES = [
        { key: "configurations",  countId: "countConfigurations"  },
        { key: "settingsCatalog", countId: "countSettingsCatalog" },
        { key: "applications",    countId: "countApplications"    },
        { key: "scripts",         countId: "countScripts"         },
        { key: "remediations",    countId: "countRemediations"    },
        { key: "groups",           countId: "countGroups"          }
    ];

    function getFilteredItems(key) {
        var items = assignmentData[key] || [];
        var filtered = items.filter(function (item) {
            if (!showAllDevices && item.assignmentType === "All Devices") return false;
            if (!showAllUsers && item.assignmentType === "All Users") return false;
            return true;
        });

        // Merge nested/inherited assignments if available and enabled
        if (showNested && nestedData && nestedData[key]) {
            var nestedItems = nestedData[key] || [];
            // Avoid duplicates: only add nested items not already in the direct list
            var directIds = {};
            filtered.forEach(function (item) {
                directIds[item.id + "|" + (item.assignmentType || "")] = true;
            });
            nestedItems.forEach(function (item) {
                var dedupKey = item.id + "|" + (item.assignmentType || "");
                if (!directIds[dedupKey]) {
                    filtered.push(item);
                }
            });
        }

        return filtered;
    }

    function updateCounts() {
        if (!assignmentData) return;
        var errors = assignmentData._errors || {};
        CATEGORIES.forEach(function (c) {
            var el = document.getElementById(c.countId);
            if (el) {
                if (errors[c.key]) {
                    el.textContent = "!";
                    el.title = "Failed to load — click to retry";
                } else {
                    el.textContent = getFilteredItems(c.key).length;
                    el.title = "";
                }
            }
        });
    }

    function highlightTab() {
        categoryTabs.querySelectorAll(".tab").forEach(function (t) {
            t.classList.toggle("active", t.dataset.category === activeCategory);
        });
    }

    function getFirstNonEmptyCategory() {
        if (!assignmentData) return null;
        for (var i = 0; i < CATEGORIES.length; i++) {
            if (getFilteredItems(CATEGORIES[i].key).length > 0) return CATEGORIES[i].key;
        }
        return null;
    }

    // ── Intune deep links ───────────────────────────────────────────────

    var INTUNE_BASE = "https://intune.microsoft.com/";

    function getIntuneUrl(category, itemId) {
        var encodedId = itemId ? encodeURIComponent(itemId) : "";
        switch (category) {
            case "configurations":
                return INTUNE_BASE + "#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/configuration";
            case "settingsCatalog":
                return INTUNE_BASE + "#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/configuration";
            case "applications":
                return INTUNE_BASE + "#view/Microsoft_Intune_Apps/SettingsMenu/appId/" + encodedId;
            case "scripts":
                return INTUNE_BASE + "#view/Microsoft_Intune_DeviceSettings/ConfigureWMPolicyMenuBlade/policyId/" + encodedId + "/policyType~/0";
            case "remediations":
                return INTUNE_BASE + "#view/Microsoft_Intune_Enrollment/UNTRemediations";
            case "groups":
                return "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/" + encodedId;
        }
        return null;
    }

    // ── Render assignment cards ─────────────────────────────────────────

    function renderCards() {
        if (!assignmentData) return;

        var errors = assignmentData._errors || {};
        var categoryError = errors[activeCategory];
        var items = getFilteredItems(activeCategory);
        cardGrid.innerHTML = "";

        // Show error banner if this category failed
        var existingBanner = document.getElementById("categoryErrorBanner");
        if (existingBanner) existingBanner.remove();

        if (categoryError) {
            var banner = document.createElement("div");
            banner.id = "categoryErrorBanner";
            banner.className = "category-error-banner";
            banner.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>' +
                '<span>Failed to load this category. The Graph API returned an error — this is usually temporary. Try selecting the group again.</span>';
            cardGrid.parentNode.insertBefore(banner, cardGrid);
        }

        if (items.length === 0) {
            cardGrid.style.display     = "none";
            categoryEmpty.style.display = categoryError ? "none" : "flex";
            return;
        }

        cardGrid.style.display      = "grid";
        categoryEmpty.style.display = "none";

        items.forEach(function (item) {
            var card = document.createElement("div");
            card.className = "assignment-card";

            var badges = [];
            if (item.assignmentType) {
                var isExclude = item.assignmentType === "Exclude";
                badges.push(
                    '<span class="badge ' + (isExclude ? "badge-exclude" : "badge-include") + '">' + escapeHtml(item.assignmentType) + '</span>'
                );
            }
            if (item.intent) {
                badges.push('<span class="badge badge-intent">' + escapeHtml(item.intent) + '</span>');
            }
            if (item.filterType && item.filterType !== "none") {
                badges.push('<span class="badge badge-filter">Filter: ' + escapeHtml(item.filterType) + '</span>');
            }
            if (item.inheritedFrom) {
                badges.push('<span class="badge badge-inherited" title="Inherited via nested group membership">Inherited: ' + escapeHtml(item.inheritedFrom) + '</span>');
            }
            if (item.platform) {
                badges.push('<span class="badge badge-platform">' + escapeHtml(item.platform) + '</span>');
            }

            var url = getIntuneUrl(activeCategory, item.id);
            var linkIcon = '<svg class="link-icon" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>';
            var nameHtml = url
                ? '<a href="' + escapeHtml(url) + '" target="_blank" rel="noopener noreferrer" title="Open in Intune">' + escapeHtml(item.displayName || "Unnamed") + linkIcon + '</a>'
                : escapeHtml(item.displayName || "Unnamed");

            var showPreview = activeCategory === "scripts" && item.id;
            var previewBtn = showPreview
                ? '<button class="btn-preview" data-script-id="' + escapeHtml(item.id) + '" data-script-name="' + escapeHtml(item.displayName || "Script") + '" title="View script content"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg></button>'
                : "";

            card.innerHTML =
                '<div class="card-header">' +
                    '<div class="card-name">' + nameHtml + '</div>' +
                    '<div class="card-actions">' + previewBtn + '</div>' +
                '</div>' +
                (item.description ? '<div class="card-desc">' + escapeHtml(item.description) + '</div>' : '') +
                '<div class="card-meta">' + badges.join("") + '</div>';

            // Add copy button to card actions
            var cardActions = card.querySelector(".card-actions");
            var itemName = item.displayName || "Unnamed";
            cardActions.appendChild(createCopyButton(function () { return itemName; }));

            if (showPreview) {
                var btn = card.querySelector(".btn-preview");
                if (btn) {
                    btn.addEventListener("click", function (e) {
                        e.stopPropagation();
                        openScriptModal(item.id, item.displayName || "Script");
                    });
                }
            }

            cardGrid.appendChild(card);
        });
    }

    // ── Orphaned items helpers ─────────────────────────────────────────

    function getFilteredOrphanedItems(key) {
        var items = orphanedData[key] || [];
        if (key === "groups") return items; // groups don't have platforms
        if (!activePlatformFilter) return items;
        return items.filter(function (item) {
            return item.platform === activePlatformFilter;
        });
    }

    function updateOrphanedCounts() {
        if (!orphanedData) return;
        var errors = orphanedData._errors || {};
        CATEGORIES.forEach(function (c) {
            var el = document.getElementById(c.countId);
            if (el) {
                if (errors[c.key]) {
                    el.textContent = "!";
                    el.title = "Failed to load — click to retry";
                } else {
                    el.textContent = getFilteredOrphanedItems(c.key).length;
                    el.title = "";
                }
            }
        });
    }

    function getFirstNonEmptyOrphanedCategory() {
        if (!orphanedData) return null;
        for (var i = 0; i < CATEGORIES.length; i++) {
            if ((orphanedData[CATEGORIES[i].key] || []).length > 0) return CATEGORIES[i].key;
        }
        return null;
    }

    function renderOrphanedCards() {
        if (!orphanedData) return;

        var errors = orphanedData._errors || {};
        var categoryError = errors[activeCategory];
        var items = getFilteredOrphanedItems(activeCategory);
        cardGrid.innerHTML = "";

        // Show error banner if this category failed
        var existingBanner = document.getElementById("categoryErrorBanner");
        if (existingBanner) existingBanner.remove();

        if (categoryError) {
            var banner = document.createElement("div");
            banner.id = "categoryErrorBanner";
            banner.className = "category-error-banner";
            banner.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>' +
                '<span>Failed to load this category.</span>';
            cardGrid.parentNode.insertBefore(banner, cardGrid);
        }

        // "Orphaned" Groups are only orphaned from an Intune-assignment perspective.
        // Warn before the user treats this list as safe-to-delete.
        var existingGroupsNotice = document.getElementById("groupsScopeNotice");
        if (existingGroupsNotice) existingGroupsNotice.remove();

        if (activeCategory === "groups") {
            var notice = document.createElement("div");
            notice.id = "groupsScopeNotice";
            notice.className = "category-info-banner";
            notice.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>' +
                '<span><strong>Heads up:</strong> &ldquo;Orphaned&rdquo; here means these groups have no Intune assignments. They may still be used elsewhere in Microsoft 365 (Exchange, SharePoint, Teams, licensing, Conditional Access, etc.). Review in the Entra admin center before deleting.</span>';
            cardGrid.parentNode.insertBefore(notice, cardGrid);
        }

        if (items.length === 0) {
            cardGrid.style.display     = "none";
            categoryEmpty.style.display = categoryError ? "none" : "flex";
            return;
        }

        cardGrid.style.display      = "grid";
        categoryEmpty.style.display = "none";

        items.forEach(function (item) {
            var card = document.createElement("div");
            card.className = "assignment-card orphaned-card";

            var url = getIntuneUrl(activeCategory, item.id);
            var linkIcon = '<svg class="link-icon" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>';
            var nameHtml = url
                ? '<a href="' + escapeHtml(url) + '" target="_blank" rel="noopener noreferrer" title="Open in Intune">' + escapeHtml(item.displayName || "Unnamed") + linkIcon + '</a>'
                : escapeHtml(item.displayName || "Unnamed");

            card.innerHTML =
                '<div class="card-header">' +
                    '<div class="card-name">' + nameHtml + '</div>' +
                    '<div class="card-actions"></div>' +
                '</div>' +
                (item.description ? '<div class="card-desc">' + escapeHtml(item.description) + '</div>' : '') +
                '<div class="card-meta">' +
                    (activeCategory === "groups" && item.groupType
                        ? '<span class="badge badge-platform">' + escapeHtml(item.groupType) + '</span>'
                        : (item.platform ? '<span class="badge badge-platform">' + escapeHtml(item.platform) + '</span>' : '')) +
                    '<span class="badge badge-orphaned">No Assignments</span>' +
                '</div>';

            var cardActions = card.querySelector(".card-actions");
            var itemName = item.displayName || "Unnamed";
            cardActions.appendChild(createCopyButton(function () { return itemName; }));

            cardGrid.appendChild(card);
        });
    }

    function exportOrphanedCsv() {
        if (!orphanedData) return;

        // Group ID column is populated for the "groups" category (orphaned
        // groups have a real Graph ID worth keeping); blank for orphaned
        // policies/apps/scripts since those IDs aren't group IDs.
        var rows = [["Category", "Group ID", "Name", "Description", "Platform / Group Type"]];

        CATEGORIES.forEach(function (c) {
            var items = getFilteredOrphanedItems(c.key);
            var label = c.key.replace(/([A-Z])/g, " $1").replace(/^./, function (s) { return s.toUpperCase(); });
            items.forEach(function (item) {
                rows.push([
                    label,
                    c.key === "groups" ? exportableGroupId(item.id) : "",
                    item.displayName || "",
                    item.description || "",
                    c.key === "groups" ? (item.groupType || "") : (item.platform || "")
                ]);
            });
        });

        var csv = rows.map(function (row) {
            return row.map(function (cell) {
                var s = String(cell).trim();
                if (/^[=+\-@\t\r]/.test(s)) { s = "'" + s; }
                s = s.replace(/"/g, '""');
                return '"' + s + '"';
            }).join(",");
        }).join("\r\n");

        var blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;
        a.download = "orphaned_items.csv";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // ── Script preview modal ────────────────────────────────────────────

    async function openScriptModal(scriptId, scriptName) {
        scriptModalTitle.textContent = scriptName;
        scriptModalFile.textContent  = "";
        scriptModalBody.innerHTML    = '<div class="modal-loading"><div class="spinner"></div><p>Loading script content...</p></div>';
        copyBtnLabel.textContent = "Copy Script";
        btnCopyScript.classList.remove("copied");
        scriptModal.classList.add("active");

        try {
            var data;
            if (appMode === "spa") {
                data = await GraphClient.getScriptContent(scriptId);
            } else {
                data = await apiFetch("/api/scripts/" + scriptId + "/content");
            }
            scriptModalFile.textContent = data.fileName ? "(" + data.fileName + ")" : "";

            if (data.content) {
                var pre = document.createElement("pre");
                pre.textContent = data.content;
                scriptModalBody.innerHTML = "";
                scriptModalBody.appendChild(pre);
            } else {
                scriptModalBody.innerHTML = '<p style="color:var(--text-muted);text-align:center;padding:32px;">No script content available.</p>';
            }
        } catch (err) {
            console.error("Failed to load script content:", err);
            scriptModalBody.innerHTML = '<p style="color:#f87171;text-align:center;padding:32px;">' + escapeHtml(err.message || "Failed to load script content.") + '</p>';
        }
    }

    function closeScriptModal() {
        scriptModal.classList.remove("active");
    }

    function copyScriptContent() {
        var pre = scriptModalBody.querySelector("pre");
        if (!pre) return;
        navigator.clipboard.writeText(pre.textContent).then(function () {
            copyBtnLabel.textContent = "Copied!";
            btnCopyScript.classList.add("copied");
            setTimeout(function () {
                copyBtnLabel.textContent = "Copy Script";
                btnCopyScript.classList.remove("copied");
            }, 2000);
        });
    }

    // ── Connection badge ────────────────────────────────────────────────

    function setConnection(state, text) {
        badgeDot.className = "badge-dot " + state;
        badgeText.textContent = text;
    }

    // ── CSV helpers ─────────────────────────────────────────────────────
    //
    // Synthetic group IDs (__allDevices__, __allUsers__, __orphaned__) are
    // internal placeholders with no Graph-side equivalent — never write them
    // to an export, leave the column blank instead.

    function exportableGroupId(id) {
        if (!id || typeof id !== "string") return "";
        if (id.indexOf("__") === 0) return "";
        return id;
    }

    // ── Export the current filtered group list to CSV ──────────────────
    //
    // Dumps whatever the sidebar is currently showing — honours the search
    // box, assigned-only toggle, and the min/max assignment-count filter.
    // Columns: Group Name, Group Type, Number of Assignments, Member Count.
    // Member count is left blank for groups whose count hasn't loaded yet
    // (rather than writing a misleading 0).

    function exportFilteredGroupsCsv() {
        var filtered = getFilteredGroups();
        if (!filtered.length) {
            alert("No groups match the current filter — nothing to export.");
            return;
        }

        var rows = [["Group ID", "Group Name", "Group Type", "Number of Assignments", "Member Count"]];
        filtered.forEach(function (g) {
            var assignCount = groupAssignCounts[g.id] || 0;
            var memberCount = Object.prototype.hasOwnProperty.call(groupMemberCounts, g.id)
                ? groupMemberCounts[g.id]
                : "";
            rows.push([
                exportableGroupId(g.id),
                g.displayName || "Unnamed Group",
                getGroupType(g),
                assignCount,
                memberCount
            ]);
        });

        var csv = rows.map(function (row) {
            return row.map(function (cell) {
                var s = String(cell).trim();
                // Prevent CSV formula injection in Excel
                if (/^[=+\-@\t\r]/.test(s)) {
                    s = "'" + s;
                }
                s = s.replace(/"/g, '""');
                return '"' + s + '"';
            }).join(",");
        }).join("\r\n");

        var blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;

        // Tag filename with filter state so exports are self-describing.
        var nameParts = ["groups"];
        if (filterAssigned) nameParts.push("assigned");
        if (filterMinCount > 0 || filterMaxCount > 0) {
            nameParts.push("count-" + (filterMinCount || 0) + "-to-" + (filterMaxCount || "any"));
        }
        var stamp = new Date().toISOString().slice(0, 10);
        a.download = nameParts.join("_") + "_" + stamp + ".csv";

        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // ── Export CSV ─────────────────────────────────────────────────────

    function exportCsv() {
        // Handle orphaned items export
        if (activeGroupId === "__orphaned__") {
            exportOrphanedCsv();
            return;
        }

        if (!assignmentData) return;

        var groupName = selectedGroupName.textContent || "Group";
        var groupIdCell = exportableGroupId(activeGroupId);
        var rows = [["Group ID", "Group Name", "Category", "Name", "Description", "Platform", "Assignment Type", "Intent", "Filter Type", "Inherited From"]];

        CATEGORIES.forEach(function (c) {
            var items = getFilteredItems(c.key);
            var label = c.key.replace(/([A-Z])/g, " $1").replace(/^./, function (s) { return s.toUpperCase(); });
            items.forEach(function (item) {
                rows.push([
                    groupIdCell,
                    groupName,
                    label,
                    item.displayName || "",
                    item.description || "",
                    item.platform || "",
                    item.assignmentType || "",
                    item.intent || "",
                    item.filterType && item.filterType !== "none" ? item.filterType : "",
                    item.inheritedFrom || ""
                ]);
            });
        });

        var csv = rows.map(function (row) {
            return row.map(function (cell) {
                var s = String(cell).trim();
                // Prevent CSV formula injection in Excel
                if (/^[=+\-@\t\r]/.test(s)) {
                    s = "'" + s;
                }
                s = s.replace(/"/g, '""');
                return '"' + s + '"';
            }).join(",");
        }).join("\r\n");

        var blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;
        a.download = groupName.replace(/[^a-z0-9]/gi, "_") + "_assignments.csv";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // ── Logout ────────────────────────────────────────────────────────

    async function logout() {
        stopIdleTimer();

        if (appMode === "spa") {
            if (!confirm("Sign out from Microsoft Graph?")) return;

            try {
                await GraphClient.signOut();
            } catch (err) {
                console.error("Logout error:", err);
            }

            // Reset UI
            allGroups         = [];
            assignedGroupIds  = new Set();
            groupAssignCounts = {};
            groupMemberCounts = {};
            memberCountsToken += 1;
            activeGroupId     = null;
            assignmentData    = null;
            nestedData        = null;
            orphanedData      = null;
            groupList.innerHTML    = "";
            groupCount.textContent = "0";
            groupSearch.value      = "";
            showPanel("empty");
            setConnection("error", "Signed out");

            // Show setup screen again
            showSetup(
                localStorage.getItem("iac_clientId"),
                localStorage.getItem("iac_tenantId")
            );
        } else {
            if (!confirm("Sign out from Microsoft Graph? You will need to restart the script to sign in again.")) {
                return;
            }

            try {
                var logoutHeaders = {};
                if (_backendKey) logoutHeaders["X-Backend-Key"] = _backendKey;
                var resp = await fetch("/api/logout", { method: "POST", headers: logoutHeaders });
                var data = await resp.json().catch(function () { return {}; });

                allGroups         = [];
                assignedGroupIds  = new Set();
                groupAssignCounts = {};
                groupMemberCounts = {};
                memberCountsToken += 1;
                activeGroupId     = null;
                assignmentData    = null;
                groupList.innerHTML    = "";
                groupCount.textContent = "0";
                groupSearch.value      = "";
                showPanel("empty");

                setConnection("error", "Signed out");
                contentErrorMsg.textContent = data.message || "Signed out. Restart the script to sign in again.";
            } catch (err) {
                console.error("Logout failed:", err);
                alert("Failed to sign out. Please try again.");
            }
        }
    }

    // ── Utilities ───────────────────────────────────────────────────────

    function escapeHtml(str) {
        var div = document.createElement("div");
        div.textContent = str;
        return div.innerHTML;
    }

    // ── Copy-to-clipboard helper ─────────────────────────────────────────

    var COPY_ICON_SVG = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>';
    var CHECK_ICON_SVG = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>';

    function createCopyButton(textFn) {
        var btn = document.createElement("button");
        btn.className = "btn-copy";
        btn.title = "Copy to clipboard";
        btn.innerHTML = COPY_ICON_SVG;
        btn.addEventListener("click", function (e) {
            e.stopPropagation();
            var text = typeof textFn === "function" ? textFn() : textFn;
            navigator.clipboard.writeText(text).then(function () {
                btn.innerHTML = CHECK_ICON_SVG;
                btn.classList.add("copied");
                setTimeout(function () {
                    btn.innerHTML = COPY_ICON_SVG;
                    btn.classList.remove("copied");
                }, 1500);
            });
        });
        return btn;
    }
})();
