/* ═══════════════════════════════════════════════════════════════════════════
   Intune Assignment Checker — Frontend
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

    // ── State ───────────────────────────────────────────────────────────
    let allGroups      = [];
    let activeGroupId  = null;
    let assignmentData = null;
    let activeCategory = "configurations";

    // ── Boot ────────────────────────────────────────────────────────────
    loadGroups();

    groupSearch.addEventListener("input", () => renderGroupList());
    document.getElementById("btnRetry").addEventListener("click", loadGroups);

    categoryTabs.addEventListener("click", (e) => {
        const tab = e.target.closest(".tab");
        if (!tab) return;
        activeCategory = tab.dataset.category;
        highlightTab();
        renderCards();
    });

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
            contentErrorMsg.textContent = "Failed to load assignments. Please check your connection and try again.";
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

            card.innerHTML = `
                <div class="card-name">${escapeHtml(item.displayName || "Unnamed")}</div>
                ${item.description ? `<div class="card-desc">${escapeHtml(item.description)}</div>` : ""}
                <div class="card-meta">${badges.join("")}</div>
            `;

            cardGrid.appendChild(card);
        });
    }

    // ── Connection badge ────────────────────────────────────────────────

    function setConnection(state, text) {
        badgeDot.className = "badge-dot " + state;
        badgeText.textContent = text;
    }

    // ── Utilities ───────────────────────────────────────────────────────

    function escapeHtml(str) {
        const div = document.createElement("div");
        div.textContent = str;
        return div.innerHTML;
    }
})();
