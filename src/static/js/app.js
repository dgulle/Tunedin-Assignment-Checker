/* ═══════════════════════════════════════════════════════════════════════════
   Intune Assignment Checker — Frontend (ZeroToTrust Edition)
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

    var selectedGroupName = document.getElementById("selectedGroupName");
    var selectedGroupDesc = document.getElementById("selectedGroupDesc");
    var categoryTabs      = document.getElementById("categoryTabs");
    var cardGrid          = document.getElementById("cardGrid");
    var categoryEmpty     = document.getElementById("categoryEmpty");

    var connectionBadge = document.getElementById("connectionBadge");
    var badgeDot        = connectionBadge.querySelector(".badge-dot");
    var badgeText       = connectionBadge.querySelector(".badge-text");

    var scriptModal      = document.getElementById("scriptModal");
    var scriptModalTitle = document.getElementById("scriptModalTitle");
    var scriptModalFile  = document.getElementById("scriptModalFile");
    var scriptModalBody  = document.getElementById("scriptModalBody");

    var btnGroupFilter = document.getElementById("btnGroupFilter");

    // Setup screen elements
    var setupScreen   = document.getElementById("setupScreen");
    var appHeader     = document.querySelector(".app-header");
    var appLayout     = document.querySelector(".app-layout");

    // ── State ───────────────────────────────────────────────────────────
    var allGroups        = [];
    var assignedGroupIds = new Set();
    var filterAssigned   = true;
    var activeGroupId    = null;
    var assignmentData   = null;
    var activeCategory   = "configurations";

    // Mode: "backend" or "spa"
    var appMode = "backend";

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
    btnGroupFilter.addEventListener("click", toggleGroupFilter);

    scriptModal.addEventListener("click", function (e) {
        if (e.target === scriptModal) closeScriptModal();
    });

    categoryTabs.addEventListener("click", function (e) {
        var tab = e.target.closest(".tab");
        if (!tab) return;
        activeCategory = tab.dataset.category;
        highlightTab();
        renderCards();
    });

    // Setup screen connect button
    document.getElementById("btnSetupConnect").addEventListener("click", setupConnect);

    // ── Mode Detection ──────────────────────────────────────────────────

    async function detectMode() {
        // Try to reach the PowerShell backend
        try {
            var resp = await fetch("/api/groups", { method: "HEAD" });
            if (resp.ok || resp.status === 200) {
                appMode = "backend";
                showApp();
                loadGroups();
                return;
            }
        } catch (e) {
            // Backend not available
        }

        // SPA mode — check if we have saved config
        appMode = "spa";
        var savedClientId = localStorage.getItem("iac_clientId");
        var savedTenantId = localStorage.getItem("iac_tenantId");

        if (savedClientId && savedTenantId) {
            // Init MSAL and try to handle redirect / silent auth
            await GraphClient.init(savedClientId, savedTenantId);
            var account = await GraphClient.handleRedirect();
            if (account) {
                showApp();
                setConnection("connected", account.username || "Connected");
                loadGroups();
                return;
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

        // Pre-initialise MSAL if we have both values, so it's ready when
        // the user clicks "Sign in" (avoids async work in click handler)
        if (savedClientId && savedTenantId && !GraphClient.isInitialised()) {
            GraphClient.init(savedClientId, savedTenantId);
        }
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

    async function setupConnect() {
        var tenantId = document.getElementById("setupTenantId").value.trim();
        var clientId = document.getElementById("setupClientId").value.trim();

        if (!tenantId || !clientId) {
            alert("Please enter both Tenant ID and Client ID.");
            return;
        }

        // Save to localStorage so we can pick up after redirect
        localStorage.setItem("iac_tenantId", tenantId);
        localStorage.setItem("iac_clientId", clientId);

        // Init MSAL if not already done (pre-init in showSetup may have done it)
        if (!GraphClient.isInitialised()) {
            await GraphClient.init(clientId, tenantId);
        }

        try {
            // This will redirect the page to Microsoft login.
            // On return, detectMode() → handleRedirect() picks up the token.
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

    async function apiFetch(url) {
        var resp = await fetch(url);
        if (!resp.ok) {
            var body = await resp.json().catch(function () { return {}; });
            throw new Error(body.error || "HTTP " + resp.status);
        }
        return resp.json();
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
                    GraphClient.getAssignedGroupIds().catch(function () { return []; })
                ]);
                groups = results[0];
                assignedIds = results[1];
            } else {
                var backendResults = await Promise.all([
                    apiFetch("/api/groups"),
                    apiFetch("/api/assigned-group-ids").catch(function () { return []; })
                ]);
                groups = backendResults[0];
                assignedIds = backendResults[1];
            }

            allGroups = groups;
            assignedGroupIds = new Set(assignedIds);
            renderGroupList();
            setConnection("connected", appMode === "spa" ? (GraphClient.getAccount()?.username || "Connected") : "Connected");
        } catch (err) {
            console.error("Failed to load groups:", err);
            sidebarErrorMsg.textContent = "Failed to load groups. " + (err.message || "Please check your connection and try again.");
            sidebarError.style.display  = "flex";
            setConnection("error", "Disconnected");
        } finally {
            sidebarLoading.style.display = "none";
        }
    }

    // ── Render group list (with search filter) ──────────────────────────

    function renderGroupList() {
        var query = groupSearch.value.trim().toLowerCase();
        var filtered = allGroups;

        if (filterAssigned && assignedGroupIds.size > 0) {
            filtered = filtered.filter(function (g) { return assignedGroupIds.has(g.id); });
        }

        if (query) {
            filtered = filtered.filter(function (g) {
                return (g.displayName || "").toLowerCase().indexOf(query) !== -1 ||
                       (g.description || "").toLowerCase().indexOf(query) !== -1;
            });
        }

        groupList.innerHTML = "";

        filtered.forEach(function (g) {
            var li = document.createElement("li");
            li.className = "group-item" + (g.id === activeGroupId ? " active" : "");
            li.dataset.id = g.id;

            var groupType = getGroupType(g);

            li.innerHTML =
                '<div class="group-item-name" title="' + escapeHtml(g.displayName || "") + '">' + escapeHtml(g.displayName || "Unnamed Group") + '</div>' +
                (g.description ? '<div class="group-item-desc" title="' + escapeHtml(g.description) + '">' + escapeHtml(g.description) + '</div>' : '') +
                '<span class="group-item-type">' + groupType + '</span>';

            li.addEventListener("click", function () { selectGroup(g); });
            groupList.appendChild(li);
        });

        groupCount.textContent = filtered.length;
    }

    function getGroupType(group) {
        var types = group.groupTypes || [];
        if (types.indexOf("DynamicMembership") !== -1) return "Dynamic";
        if (group.membershipRule) return "Dynamic";
        return "Assigned";
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

    // ── Select a group ──────────────────────────────────────────────────

    async function selectGroup(group) {
        activeGroupId = group.id;
        renderGroupList();
        showPanel("loading");

        try {
            if (appMode === "spa") {
                assignmentData = await GraphClient.getAssignmentsForGroup(group.id);
            } else {
                assignmentData = await apiFetch("/api/groups/" + group.id + "/assignments");
            }

            selectedGroupName.textContent = group.displayName || "Unnamed Group";
            selectedGroupDesc.textContent = group.description || "";
            updateCounts();
            activeCategory = getFirstNonEmptyCategory() || "configurations";
            highlightTab();
            renderCards();
            showPanel("assignments");
        } catch (err) {
            console.error("Failed to load assignments:", err);
            contentErrorMsg.textContent = err.message || "Failed to load assignments.";
            showPanel("error");
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
        { key: "remediations",    countId: "countRemediations"    }
    ];

    function updateCounts() {
        if (!assignmentData) return;
        CATEGORIES.forEach(function (c) {
            var el = document.getElementById(c.countId);
            if (el) el.textContent = (assignmentData[c.key] || []).length;
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
            if ((assignmentData[CATEGORIES[i].key] || []).length > 0) return CATEGORIES[i].key;
        }
        return null;
    }

    // ── Intune deep links ───────────────────────────────────────────────

    var INTUNE_BASE = "https://intune.microsoft.com/";

    function getIntuneUrl(category, itemId) {
        switch (category) {
            case "configurations":
                return INTUNE_BASE + "#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/configuration";
            case "settingsCatalog":
                return INTUNE_BASE + "#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/configuration";
            case "applications":
                return INTUNE_BASE + "#view/Microsoft_Intune_Apps/SettingsMenu/appId/" + itemId;
            case "scripts":
                return INTUNE_BASE + "#view/Microsoft_Intune_DeviceSettings/ConfigureWMPolicyMenuBlade/policyId/" + itemId + "/policyType~/0";
            case "remediations":
                return INTUNE_BASE + "#view/Microsoft_Intune_Enrollment/UNTHealthScriptPolicy/healthScriptId/" + itemId;
        }
        return null;
    }

    // ── Render assignment cards ─────────────────────────────────────────

    function renderCards() {
        if (!assignmentData) return;

        var items = assignmentData[activeCategory] || [];
        cardGrid.innerHTML = "";

        if (items.length === 0) {
            cardGrid.style.display     = "none";
            categoryEmpty.style.display = "flex";
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
                    (previewBtn ? '<div class="card-actions">' + previewBtn + '</div>' : '') +
                '</div>' +
                (item.description ? '<div class="card-desc">' + escapeHtml(item.description) + '</div>' : '') +
                '<div class="card-meta">' + badges.join("") + '</div>';

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

    // ── Script preview modal ────────────────────────────────────────────

    async function openScriptModal(scriptId, scriptName) {
        scriptModalTitle.textContent = scriptName;
        scriptModalFile.textContent  = "";
        scriptModalBody.innerHTML    = '<div class="modal-loading"><div class="spinner"></div><p>Loading script content...</p></div>';
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

    // ── Connection badge ────────────────────────────────────────────────

    function setConnection(state, text) {
        badgeDot.className = "badge-dot " + state;
        badgeText.textContent = text;
    }

    // ── Logout ────────────────────────────────────────────────────────

    async function logout() {
        if (appMode === "spa") {
            if (!confirm("Sign out from Microsoft Graph?")) return;

            try {
                await GraphClient.signOut();
            } catch (err) {
                console.error("Logout error:", err);
            }

            // Reset UI
            allGroups        = [];
            assignedGroupIds = new Set();
            activeGroupId    = null;
            assignmentData   = null;
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
                var resp = await fetch("/api/logout", { method: "POST" });
                var data = await resp.json().catch(function () { return {}; });

                allGroups        = [];
                assignedGroupIds = new Set();
                activeGroupId    = null;
                assignmentData   = null;
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
})();
