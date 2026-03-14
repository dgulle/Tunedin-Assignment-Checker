# Tunedin Assignment Checker

A web dashboard that connects to Microsoft Intune via the Microsoft Graph API and displays policy and application assignments for Entra ID groups. Run it locally with the PowerShell backend, or use it directly from GitHub Pages as a standalone single-page app, your choice.

## Features

- Browse all Entra ID groups in a searchable sidebar
- Filter to show only groups that have Intune assignments
- **Assignment count per group** - each group displays the total number of directly-targeted assignments (excludes All Devices/All Users)
- **Assignment count range filter** - filter the group list by assignment count (e.g. show only groups with 1–10 assignments)
  
- View Intune assignments per group across five categories:
  - **Device Configurations** - Configuration profiles
  - **Settings Catalog** - Settings Catalog policies
  - **Applications** - Assigned apps (required, available, uninstall)
  - **Scripts** - PowerShell device management scripts (with content preview)
  - **Remediations** - Proactive remediation (health) scripts
- See assignment type (Include / Exclude / All Users / All Devices) and filter information

- **All Devices & All Users groups** - dedicated entries at the bottom of the group list let you see every policy, app, script, and remediation assigned to All Devices or All Users across all categories in one click
- **Show/Hide All Devices & All Users** - global toggle buttons in the header to show or hide All Devices and All Users assignments across all groups, reducing clutter in large tenants
- **Dynamic membership rule display** - when selecting a dynamic group, the membership rule query is shown below the group name for quick reference
- **Copy to clipboard** - hover-to-reveal copy buttons on group names, descriptions, dynamic membership rules, and assignment card names for fast copy/paste

- **Export to CSV** - download all assignments for the selected group as a CSV file for offline review or troubleshooting
- Direct deep links to policies and apps in the Intune portal
- **Tenant switching** - change Tenant ID and Client ID without needing to clear your browser history
- **Resilient Graph API handling** - automatic retry with exponential backoff on throttling (HTTP 429) and transient server errors (HTTP 5xx); partial results are shown if individual categories fail
- Dark mode with system preference detection
- Responsive design for desktop, tablet, and mobile

## Two Ways to Run

### Option 1: PowerShell Backend (Local)

Run the PowerShell script locally. No app registration required - permissions are requested automatically at sign-in.

**Prerequisites:**

- **PowerShell 5.1+** (Windows PowerShell) or **PowerShell 7+** (cross-platform)
- An Entra ID account with sufficient privileges to read Intune configuration and group data
- The ability to consent to the required Microsoft Graph scopes (see [Permissions](#permissions))

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
2. Open a browser window for interactive Entra ID sign-in.
3. Request the required Graph permissions (consent prompt).
4. Start a local web server on **http://localhost:8080** and open it in your default browser.

**Custom Port:**

```powershell
.\TunedinAssignmentChecker.ps1 -Port 9090
```

### Option 2: GitHub Pages / Static Hosting (SPA Mode)

Use the app directly from [https://dgulle.github.io/Tunedin-Assignment-Checker/](https://dgulle.github.io/Tunedin-Assignment-Checker/) - no PowerShell or local install needed. The app runs entirely in your browser using MSAL.js to authenticate directly with Microsoft Graph.

**This option requires you to create your own Entra ID app registration** in your tenant.

#### Setting Up Your App Registration

1. Go to the **Entra Portal** ([entra.microsoft.com](https://entra.microsoft.com)) > **App registrations**
2. Click **New registration**
3. Fill in the details:
   - **Name:** `Tunedin Assignment Checker` (or any name you prefer)
   - **Supported account types:** Accounts in this organizational directory only (single tenant)
   - **Redirect URI:** Select **Single-page application (SPA)** and enter:
     ```
     https://dgulle.github.io/Tunedin-Assignment-Checker/
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

1. Open the app at [https://dgulle.github.io/Tunedin-Assignment-Checker/](https://dgulle.github.io/Tunedin-Assignment-Checker/)
2. Enter your **Tenant ID** and **Client ID** on the setup screen
3. Click **Sign in with Microsoft**
4. Sign in with your Entra ID account and consent to the permissions if prompted
5. Your Client ID and Tenant ID are saved in your browser's local storage, so you won't need to re-enter them next time

> **Switching tenants:** To connect to a different tenant, click **Sign out**, update the Tenant ID and Client ID on the setup screen, and sign in again. The app automatically resets the authentication session when it detects changed credentials - no need to clear browser history.

> If you self-host the app on a different domain, update the Redirect URI in your app registration to match.



## Project Structure

```
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

For tenants with thousands of groups (6,000+), the app includes:

- **Automatic retry** - Graph API requests that fail with HTTP 429 (throttled) or 5xx (server error) are automatically retried up to 3 times with exponential backoff (2s, 4s, 8s). The `Retry-After` header is respected when present.
- **Partial results** - if one category (e.g. Applications) fails after retries, the other four categories still display their results. The failed category's tab shows `!` and an error banner explains the issue. Re-selecting the group retries the fetch.
- **Assignment count filter** - quickly narrow down the group list using the min/max assignment count filter to find the groups you care about.

## Security Notes

- **PowerShell mode:** Authentication uses interactive delegated flow - no secrets are stored anywhere
- **SPA mode:** Authentication uses MSAL.js with PKCE (auth code flow) - no client secret needed. Only your Client ID and Tenant ID are stored in localStorage (these are not secrets)
- The app only requests **read** permissions; it cannot modify your Intune environment
- All Intune and Entra data should be treated as sensitive - avoid using the app on shared/public computers
- Press **Ctrl+C** to stop the PowerShell server; the script disconnects from Microsoft Graph automatically
- In SPA mode, click **Sign out** to clear your session
