# Tunedin Assignment Checker

**Live App:** https://tunedin.zerototrust.tech/

A web dashboard that connects to Microsoft Intune via the Microsoft Graph API and displays policy and application assignments for Entra ID groups. Run it locally with the PowerShell backend, or use it directly from GitHub Pages as a standalone single-page app, your choice.

## Features

- Browse all Entra ID groups in a searchable sidebar — each group shows a **Dynamic** or **Assigned** badge so you can quickly identify its membership type
- Filter to show only groups that have Intune assignments
- **Assignment count per group** — each group displays the total number of directly-targeted assignments (excludes All Devices/All Users)
- **Assignment count range filter** — narrow down the group list by assignment count (e.g. show only groups with 1–10 assignments)
- **Live connection status** — a badge in the header shows who you're signed in as

- View Intune assignments per group across five categories:
  - **Device Configurations** — Configuration profiles
  - **Settings Catalog** — Settings Catalog policies
  - **Applications** — Assigned apps (required, available, uninstall)
  - **Scripts** — PowerShell device management scripts (with content preview)
  - **Remediations** — Proactive remediation (health) scripts
- See assignment type (Include / Exclude / All Users / All Devices), intent, and filter information at a glance with colour-coded badges
- **Nested group assignments** — when a group is nested inside another group, inherited assignments from parent groups are automatically discovered and shown with an "Inherited: Parent Group Name" badge, so you can see exactly where each assignment originates. A **Nested Groups** toggle in the header lets you show or hide inherited assignments.

- **All Devices & All Users groups** — dedicated entries at the bottom of the group list let you see every policy, app, script, and remediation assigned to All Devices or All Users across all categories in one click
- **Show/Hide All Devices & All Users** — global toggle buttons in the header to show or hide those assignments across all groups, reducing clutter in large tenants
- **Dynamic membership rule display** — when selecting a dynamic group, the membership rule query is shown below the group name for quick reference
- **Copy to clipboard** — hover-to-reveal copy buttons on group names, descriptions, dynamic membership rules, and assignment card names for fast copy/paste

- **Orphaned items detection** — a dedicated **Orphaned Items** view lists all Intune items (configurations, settings catalog policies, applications, scripts, and remediations) that have zero assignments, making it easy to identify stale items for cleanup. Includes a **CSV export** for quick reporting.

- **Export to CSV** — download all assignments for the selected group as a CSV file for offline review or reporting
- Direct deep links to policies and apps in the Intune portal
- **Tenant switching** — change tenants without needing to clear your browser history
- **Reliable data loading** — the app handles temporary Microsoft Graph issues automatically and still shows results even if one category encounters an error
- Dark mode with system preference detection
- Responsive design for desktop, tablet, and mobile

## Two Ways to Run

### Option 1: PowerShell Backend (Local)

Run the PowerShell script locally. **No app registration required** — permissions are requested automatically through the **Microsoft Graph Command Line Tools** enterprise application.

**Prerequisites:**

- **PowerShell 5.1+** or **PowerShell 7+** (preferred)
- An Entra ID account with sufficient privileges to read Intune configuration and group data

> The script installs the `Microsoft.Graph.Authentication` module automatically if it is not already present.

**Quick Start:**

```powershell
# Clone the repository
git clone https://github.com/dgulle/Tunedin-Assignment-Checker.git
cd Tunedin-Assignment-Checker/src

# Run the script
.\TunedinAssignmentChecker.ps1
```

The script will:

1. Install the `Microsoft.Graph.Authentication` module (first run only).
2. Connect to Microsoft Graph via the **Microsoft Graph Command Line Tools** app — no app registration needed.
3. Open a browser window for sign-in. On first use, you may be prompted to consent to the required permissions.
4. Start a local web server on **http://localhost:8080** and open it in your default browser.

**Custom Port:**

```powershell
.\TunedinAssignmentChecker.ps1 -Port 9090
```

### Option 2: GitHub Pages / Static Hosting (SPA Mode)

Use the app directly from [http://tunedin.zerototrust.tech/](http://tunedin.zerototrust.tech/) - no PowerShell or local install needed. The app runs entirely in your browser using MSAL.js to authenticate directly with Microsoft Graph.

**This option requires you to create your own Entra ID app registration** in your tenant.

#### Setting Up Your App Registration

1. Go to the **Entra Portal** ([entra.microsoft.com](https://entra.microsoft.com)) > **App registrations**
2. Click **New registration**
3. Fill in the details:
   - **Name:** `Tunedin Assignment Checker` (or any name you prefer)
   - **Supported account types:** Accounts in this organizational directory only (single tenant)
   - **Redirect URI:** Select **Single-page application (SPA)** and enter:
     ```
     http://tunedin.zerototrust.tech/
     ```
4. Click **Register**
5. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions** and add:
   - `DeviceManagementApps.Read.All`
   - `DeviceManagementConfiguration.Read.All`
   - `DeviceManagementManagedDevices.Read.All`
   - `DeviceManagementScripts.Read.All`
   - `Group.Read.All`
   - `User.Read`
   - `User.Read.All`
6. (Recommended) Click **Grant admin consent** for your organisation so users don't have to consent individually
7. Copy the **Application (client) ID** and your **Tenant ID** from the app registration overview page

#### Connecting

1. Open the app at [http://tunedin.zerototrust.tech/](http://tunedin.zerototrust.tech/)
2. Enter your **Tenant ID** and **Client ID** on the setup screen
3. Click **Sign in with Microsoft**
4. Sign in with your Entra ID account and consent to the permissions if prompted
5. Your Client ID and Tenant ID are saved in your browser's local storage, so you won't need to re-enter them next time

> **Switching tenants:** To connect to a different tenant, click **Sign out**, update the Tenant ID and Client ID on the setup screen, and sign in again. The app automatically resets the authentication session when it detects changed credentials - no need to clear browser history.

> If you self-host the app on a different domain, update the Redirect URI in your app registration to match.



## Project Structure

```
.github/
└── workflows/
    └── deploy-pages.yml              # GitHub Actions workflow (deploys to GitHub Pages on push to main)

src/
├── TunedinAssignmentChecker.ps1   # PowerShell backend (auth, HTTP server, Graph queries)
├── static/
│   ├── css/
│   │   └── style.css             # UI theme and layout styles
│   ├── img/
│   │   └── logo.svg              # Shield logo
│   └── js/
│       ├── app.js                # Frontend logic (dual-mode: backend + SPA)
│       └── graph.js              # MSAL.js auth and Graph API client (SPA mode)
└── templates/
    └── index.html                # Single-page application template (backend mode)

index.html                        # Root HTML entry point (GitHub Pages / SPA mode)
```

## Permissions

Both modes require the same Microsoft Graph **delegated** permissions:

| Permission | Type | Description | Admin Consent Required |
|---|---|---|---|
| `DeviceManagementApps.Read.All` | Delegated | Read Microsoft Intune apps | Yes |
| `DeviceManagementConfiguration.Read.All` | Delegated | Read Microsoft Intune Device Configuration | Yes |
| `DeviceManagementManagedDevices.Read.All` | Delegated | Read Microsoft Intune devices | Yes |
| `DeviceManagementScripts.Read.All` | Delegated | Read Microsoft Intune Scripts | Yes |
| `Group.Read.All` | Delegated | Read all groups | Yes |
| `User.Read` | Delegated | Sign in and read user profile | No |
| `User.Read.All` | Delegated | Read all users' full profiles | Yes |

All permissions are **read-only**. The app cannot modify your Intune environment.

## Large Tenant Support

For tenants with thousands of groups, the app is built to stay responsive:

- **Stays fast in large tenants** — the group list handles thousands of groups without slowing down, and the assignment count filter helps you quickly find what you're looking for
- **Handles temporary issues gracefully** — if Microsoft Graph is temporarily slow or unavailable, the app retries automatically and still displays results for any categories that loaded successfully
- **Partial results over no results** — if one category (e.g. Applications) can't be loaded, the other four still display. A `!` on the affected tab lets you know, and re-selecting the group will retry.

## Security Notes

- **Read-only** — the app only requests read permissions and cannot make any changes to your Intune environment
- **No secrets stored** — in PowerShell mode, no credentials are saved anywhere; in SPA mode, only your Client ID and Tenant ID are remembered in your browser (these are not secrets)
- **Auto sign-out** — in SPA mode, the app automatically signs you out after 30 minutes of inactivity
- Treat all Intune and Entra data as sensitive — avoid using the app on shared or public computers
- Press **Ctrl+C** to stop the PowerShell server; it disconnects from Microsoft Graph automatically
- In SPA mode, click **Sign out** to end your session
