/* ═══════════════════════════════════════════════════════════════════════════
   Intune Assignment Checker — Frontend (ZeroToTrust Edition)
   ═══════════════════════════════════════════════════════════════════════════ */

(function () {
    "use strict";

    // ── DOM references ──────────────────────────────────────────────────
    const groupList       = document.getElementById("groupList");
    const groupSearch     = document.getElementById("groupSearch");
    const groupCount      = document.getElementById("groupCount");
    const sidebarLoading  = document.getElementById("sidebarLoading");
    const sidebarError    = document.getElementById("sidebarError");
    const sidebarErrorMsg = document.getElementById("sidebarErrorMsg");

    const emptyState      = document.getElementById("emptyState");
    const contentLoading  = document.getElementById("contentLoading");
    const contentError    = document.getElementById("contentError");
    const contentErrorMsg = document.getElementById("contentErrorMsg");
    const assignments     = document.getElementById("assignments");

    const selectedGroupName = document.getElementById("selectedGroupName");
    const selectedGroupDesc = document.getElementById("selectedGroupDesc");
    const categoryTabs      = document.getElementById("categoryTabs");
    const cardGrid          = document.getElementById("cardGrid");
    const categoryEmpty     = document.getElementById("categoryEmpty");

    const connectionBadge = document.getElementById("connectionBadge");
    const badgeDot        = connectionBadge.querySelector(".badge-dot");
    const badgeText       = connectionBadge.querySelector(".badge-text");

    const scriptModal     = document.getElementById("scriptModal");
    const scriptModalTitle = document.getElementById("scriptModalTitle");
    const scriptModalFile  = document.getElementById("scriptModalFile");
    const scriptModalBody  = document.getElementById("scriptModalBody");

    // ── State ───────────────────────────────────────────────────────────
    let allGroups      = [];
    let activeGroupId  = null;
    let assignmentData = null;
    let activeCategory = "configurations";

    // ── Boot ────────────────────────────────────────────────────────────
    initTheme();
    loadGroups();

    groupSearch.addEventListener("input", () => renderGroupList());
    document.getElementById("btnRetry").addEventListener("click", loadGroups);
    document.getElementById("btnLogout").addEventListener("click", logout);
    document.getElementById("btnTheme").addEventListener("click", toggleTheme);
    document.getElementById("btnModalClose").addEventListener("click", closeScriptModal);

    scriptModal.addEventListener("click", (e) => {
        if (e.target === scriptModal) closeScriptModal();
    });

    categoryTabs.addEventListener("click", (e) => {
        const tab = e.target.closest(".tab");
        if (!tab) return;
        activeCategory = tab.dataset.category;
        highlightTab();
        renderCards();
    });

    // ── Dark mode ─────────────────────────────────────────────────────

    function initTheme() {
        const saved = localStorage.getItem("theme");
        if (saved === "dark" || (!saved && window.matchMedia("(prefers-color-scheme: dark)").matches)) {
            document.documentElement.setAttribute("data-theme", "dark");
        }
        updateThemeIcon();
    }

    function toggleTheme() {
        const isDark = document.documentElement.getAttribute("data-theme") === "dark";
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
        const isDark = document.documentElement.getAttribute("data-theme") === "dark";
        document.getElementById("iconSun").style.display  = isDark ? "block" : "none";
        document.getElementById("iconMoon").style.display  = isDark ? "none"  : "block";
    }

    // ── API helpers ─────────────────────────────────────────────────────

    async function apiFetch(url) {
        const resp = await fetch(url);
        if (!resp.ok) {
            const body = await resp.json().catch(() => ({}));
            throw new Error(body.error || `HTTP ${resp.status}`);
        }
        return resp.json();
    }

    // ── Load groups ─────────────────────────────────────────────────────

    async function loadGroups() {
        sidebarLoading.style.display = "flex";
        sidebarError.style.display   = "none";
        groupList.innerHTML           = "";

        try {
            allGroups = await apiFetch("/api/groups");
            groupCount.textContent = allGroups.length;
            renderGroupList();
            setConnection("connected", "Connected");
        } catch (err) {
            console.error("Failed to load groups:", err);
            sidebarErrorMsg.textContent = "Failed to load groups. Please check your connection and try again.";
            sidebarError.style.display  = "flex";
            setConnection("error", "Disconnected");
        } finally {
            sidebarLoading.style.display = "none";
        }
    }

    // ── Render group list (with search filter) ──────────────────────────

    function renderGroupList() {
        const query = groupSearch.value.trim().toLowerCase();
        const filtered = query
            ? allGroups.filter(g =>
                (g.displayName || "").toLowerCase().includes(query) ||
                (g.description || "").toLowerCase().includes(query)
            )
            : allGroups;

        groupList.innerHTML = "";

        filtered.forEach(g => {
            const li = document.createElement("li");
            li.className = "group-item" + (g.id === activeGroupId ? " active" : "");
            li.dataset.id = g.id;

            const groupType = getGroupType(g);

            li.innerHTML = `
                <div class="group-item-name" title="${escapeHtml(g.displayName || "")}">${escapeHtml(g.displayName || "Unnamed Group")}</div>
                ${g.description ? `<div class="group-item-desc" title="${escapeHtml(g.description)}">${escapeHtml(g.description)}</div>` : ""}
                <span class="group-item-type">${groupType}</span>
            `;

            li.addEventListener("click", () => selectGroup(g));
            groupList.appendChild(li);
        });

        groupCount.textContent = filtered.length;
    }

    function getGroupType(group) {
        const types = group.groupTypes || [];
        if (types.includes("DynamicMembership")) return "Dynamic";
        if (group.membershipRule) return "Dynamic";
        return "Assigned";
    }

    // ── Select a group ──────────────────────────────────────────────────

    async function selectGroup(group) {
        activeGroupId = group.id;
        renderGroupList(); // update active highlight

        showPanel("loading");

        try {
            assignmentData = await apiFetch(`/api/groups/${group.id}/assignments`);
            selectedGroupName.textContent = group.displayName || "Unnamed Group";
            selectedGroupDesc.textContent = group.description || "";
            updateCounts();
            activeCategory = getFirstNonEmptyCategory() || "configurations";
            highlightTab();
            renderCards();
            showPanel("assignments");
        } catch (err) {
            console.error("Failed to load assignments:", err);
            contentErrorMsg.textContent = err.message || "Failed to load assignments. Please check your connection and try again.";
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

    const CATEGORIES = [
        { key: "configurations",  countId: "countConfigurations"  },
        { key: "settingsCatalog", countId: "countSettingsCatalog" },
        { key: "applications",    countId: "countApplications"    },
        { key: "scripts",         countId: "countScripts"         },
        { key: "remediations",    countId: "countRemediations"    },
    ];

    function updateCounts() {
        if (!assignmentData) return;
        CATEGORIES.forEach(c => {
            const el = document.getElementById(c.countId);
            if (el) el.textContent = (assignmentData[c.key] || []).length;
        });
    }

    function highlightTab() {
        categoryTabs.querySelectorAll(".tab").forEach(t => {
            t.classList.toggle("active", t.dataset.category === activeCategory);
        });
    }

    function getFirstNonEmptyCategory() {
        if (!assignmentData) return null;
        for (const c of CATEGORIES) {
            if ((assignmentData[c.key] || []).length > 0) return c.key;
        }
        return null;
    }

    // ── Intune deep links ───────────────────────────────────────────────

    const INTUNE_BASE = "https://intune.microsoft.com/";

    function getIntuneUrl(category, itemId) {
        switch (category) {
            case "configurations":
                return `${INTUNE_BASE}#view/Microsoft_Intune_DeviceSettings/DevicesConfigProfileMenu/configurationId/${itemId}`;
            case "settingsCatalog":
                return `${INTUNE_BASE}#view/Microsoft_Intune_DeviceSettings/DevicesConfigProfileMenu/configurationId/${itemId}`;
            case "applications":
                return `${INTUNE_BASE}#view/Microsoft_Intune_Apps/SettingsMenu/appId/${itemId}`;
            case "scripts":
                return `${INTUNE_BASE}#view/Microsoft_Intune_DeviceSettings/ConfigureWMPolicyMenuBlade/policyId/${itemId}/policyType~/0`;
            case "remediations":
                return `${INTUNE_BASE}#view/Microsoft_Intune_Enrollment/UNTHealthScriptPolicy/healthScriptId/${itemId}`;
        }
        return null;
    }

    // ── Render assignment cards ─────────────────────────────────────────

    function renderCards() {
        if (!assignmentData) return;

        const items = assignmentData[activeCategory] || [];
        cardGrid.innerHTML = "";

        if (items.length === 0) {
            cardGrid.style.display     = "none";
            categoryEmpty.style.display = "flex";
            return;
        }

        cardGrid.style.display      = "grid";
        categoryEmpty.style.display = "none";

        items.forEach(item => {
            const card = document.createElement("div");
            card.className = "assignment-card";

            const badges = [];
            if (item.assignmentType) {
                const isExclude = item.assignmentType === "Exclude";
                badges.push(
                    `<span class="badge ${isExclude ? "badge-exclude" : "badge-include"}">${escapeHtml(item.assignmentType)}</span>`
                );
            }
            if (item.intent) {
                badges.push(`<span class="badge badge-intent">${escapeHtml(item.intent)}</span>`);
            }
            if (item.filterType && item.filterType !== "none") {
                badges.push(`<span class="badge badge-filter">Filter: ${escapeHtml(item.filterType)}</span>`);
            }

            // Intune deep link
            const url = getIntuneUrl(activeCategory, item.id);
            const linkIcon = '<svg class="link-icon" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>';
            const nameHtml = url
                ? `<a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer" title="Open in Intune">${escapeHtml(item.displayName || "Unnamed")}${linkIcon}</a>`
                : escapeHtml(item.displayName || "Unnamed");

            // Script preview button (eye icon) — only for scripts category
            const showPreview = activeCategory === "scripts" && item.id;
            const previewBtn = showPreview
                ? `<button class="btn-preview" data-script-id="${escapeHtml(item.id)}" data-script-name="${escapeHtml(item.displayName || "Script")}" title="View script content"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg></button>`
                : "";

            card.innerHTML = `
                <div class="card-header">
                    <div class="card-name">${nameHtml}</div>
                    ${previewBtn ? `<div class="card-actions">${previewBtn}</div>` : ""}
                </div>
                ${item.description ? `<div class="card-desc">${escapeHtml(item.description)}</div>` : ""}
                <div class="card-meta">${badges.join("")}</div>
            `;

            // Bind preview button click
            if (showPreview) {
                const btn = card.querySelector(".btn-preview");
                if (btn) {
                    btn.addEventListener("click", (e) => {
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
            const data = await apiFetch(`/api/scripts/${scriptId}/content`);
            scriptModalFile.textContent = data.fileName ? `(${data.fileName})` : "";

            if (data.content) {
                const pre = document.createElement("pre");
                pre.textContent = data.content;
                scriptModalBody.innerHTML = "";
                scriptModalBody.appendChild(pre);
            } else {
                scriptModalBody.innerHTML = '<p style="color:var(--text-muted);text-align:center;padding:32px;">No script content available.</p>';
            }
        } catch (err) {
            console.error("Failed to load script content:", err);
            scriptModalBody.innerHTML = `<p style="color:#f87171;text-align:center;padding:32px;">${escapeHtml(err.message || "Failed to load script content.")}</p>`;
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
        if (!confirm("Sign out from Microsoft Graph? You will need to restart the script to sign in again.")) {
            return;
        }

        try {
            const resp = await fetch("/api/logout", { method: "POST" });
            const data = await resp.json().catch(() => ({}));

            // Reset UI state
            allGroups      = [];
            activeGroupId  = null;
            assignmentData = null;
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

    // ── Utilities ───────────────────────────────────────────────────────

    function escapeHtml(str) {
        const div = document.createElement("div");
        div.textContent = str;
        return div.innerHTML;
    }
})();
