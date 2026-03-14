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
    var filterAssigned     = true;
    var filterMinCount     = 0;    // min assignment count filter (0 = no filter)
    var filterMaxCount     = 0;    // max assignment count filter (0 = no filter)
    var showAllDevices     = true; // toggle for All Devices assignments
    var showAllUsers       = true; // toggle for All Users assignments
    var activeGroupId      = null;
    var assignmentData     = null;
    var activeCategory     = "configurations";

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

        // SPA mode
        appMode = "spa";
        var savedClientId = localStorage.getItem("iac_clientId");
        var savedTenantId = localStorage.getItem("iac_tenantId");

        if (savedClientId && savedTenantId) {
            try {
                // Init MSAL, handle any pending redirect, check for cached account
                var account = await GraphClient.init(savedClientId, savedTenantId);
                if (account) {
                    showApp();
                    setConnection("connected", account.username || "Connected");
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

            allGroups = groups;
            assignedGroupIds = new Set(Array.isArray(assignedIds) ? assignedIds : Object.keys(groupAssignCounts));
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

    // ── Synthetic groups for All Devices / All Users ───────────────────

    var SYNTHETIC_GROUPS = [
        { id: "__allDevices__", displayName: "All Devices", description: "Policies and apps assigned to all devices", _synthetic: true },
        { id: "__allUsers__",   displayName: "All Users",   description: "Policies and apps assigned to all licensed users", _synthetic: true }
    ];

    // ── Render group list (with search filter) ──────────────────────────

    function renderGroupList() {
        var query = groupSearch.value.trim().toLowerCase();
        var filtered = allGroups;

        if (filterAssigned && assignedGroupIds.size > 0) {
            filtered = filtered.filter(function (g) { return assignedGroupIds.has(g.id); });
        }

        // Assignment count range filter
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

        groupList.innerHTML = "";

        filtered.forEach(function (g) {
            var li = document.createElement("li");
            li.className = "group-item" + (g.id === activeGroupId ? " active" : "");
            li.dataset.id = g.id;

            var groupType = getGroupType(g);
            var assignCount = groupAssignCounts[g.id] || 0;
            var gName = g.displayName || "Unnamed Group";

            li.innerHTML =
                '<div class="group-item-header">' +
                    '<div class="group-item-name" title="' + escapeHtml(gName) + '">' + escapeHtml(gName) + '</div>' +
                    '<div class="group-item-copy"></div>' +
                '</div>' +
                (g.description ? '<div class="group-item-desc" title="' + escapeHtml(g.description) + '">' + escapeHtml(g.description) + '</div>' : '') +
                '<div class="group-item-badges">' +
                    '<span class="group-item-type">' + escapeHtml(groupType) + '</span>' +
                    (assignCount > 0 ? '<span class="group-item-count" title="Total assignments">' + assignCount + ' assignment' + (assignCount !== 1 ? 's' : '') + '</span>' : '') +
                '</div>';

            li.querySelector(".group-item-copy").appendChild(createCopyButton(function () { return gName; }));
            li.addEventListener("click", function () { selectGroup(g); });
            groupList.appendChild(li);
        });

        // Separator before synthetic groups at the bottom
        var syntheticFiltered = SYNTHETIC_GROUPS.filter(function (g) {
            return !query || (g.displayName || "").toLowerCase().indexOf(query) !== -1;
        });
        if (syntheticFiltered.length > 0 && filtered.length > 0) {
            var sep = document.createElement("li");
            sep.className = "group-list-separator";
            groupList.appendChild(sep);
        }

        syntheticFiltered.forEach(function (g) {
            var li = document.createElement("li");
            li.className = "group-item group-item-synthetic" + (g.id === activeGroupId ? " active" : "");
            li.dataset.id = g.id;

            var icon = g.id === "__allDevices__"
                ? '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="3" width="20" height="14" rx="2" ry="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></svg>'
                : '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>';

            li.innerHTML =
                '<div class="group-item-name">' + icon + ' ' + escapeHtml(g.displayName) + '</div>' +
                '<div class="group-item-desc">' + escapeHtml(g.description) + '</div>';

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
        renderGroupList();
        showPanel("loading");

        try {
            if (group._synthetic) {
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
                assignmentData = await GraphClient.getAssignmentsForGroup(group.id);
            } else {
                assignmentData = await apiFetch("/api/groups/" + group.id + "/assignments");
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
        { key: "remediations",    countId: "countRemediations"    }
    ];

    function getFilteredItems(key) {
        var items = assignmentData[key] || [];
        return items.filter(function (item) {
            if (!showAllDevices && item.assignmentType === "All Devices") return false;
            if (!showAllUsers && item.assignmentType === "All Users") return false;
            return true;
        });
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
                return INTUNE_BASE + "#view/Microsoft_Intune_Enrollment/UNTRemediations";
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

    // ── Export CSV ─────────────────────────────────────────────────────

    function exportCsv() {
        if (!assignmentData) return;

        var groupName = selectedGroupName.textContent || "Group";
        var rows = [["Category", "Name", "Description", "Assignment Type", "Intent", "Filter Type"]];

        CATEGORIES.forEach(function (c) {
            var items = getFilteredItems(c.key);
            var label = c.key.replace(/([A-Z])/g, " $1").replace(/^./, function (s) { return s.toUpperCase(); });
            items.forEach(function (item) {
                rows.push([
                    label,
                    item.displayName || "",
                    item.description || "",
                    item.assignmentType || "",
                    item.intent || "",
                    item.filterType && item.filterType !== "none" ? item.filterType : ""
                ]);
            });
        });

        var csv = rows.map(function (row) {
            return row.map(function (cell) {
                var s = String(cell);
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
            activeGroupId     = null;
            assignmentData    = null;
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

                allGroups         = [];
                assignedGroupIds  = new Set();
                groupAssignCounts = {};
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
