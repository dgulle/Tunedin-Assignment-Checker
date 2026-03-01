# Intune Assignment Checker

A PowerShell-based web dashboard that connects to Microsoft Intune via the Microsoft Graph API and displays policy and application assignments for Entra ID groups. No app registrations, client secrets, or manual Azure portal setup required — just run the script and sign in.

## Features

- **Zero configuration** — no app registrations or secrets to manage; permissions are requested automatically at sign-in
- Browse all Entra ID groups in a searchable sidebar
- View Intune assignments per group across five categories:
  - **Device Configurations** — Configuration profiles
  - **Settings Catalog** — Settings Catalog policies
  - **Applications** — Assigned apps (required, available, uninstall)
  - **Scripts** — PowerShell device management scripts
  - **Remediations** — Proactive remediation (health) scripts
- See assignment type (Include / Exclude) and assignment filter information
- Responsive design that works on desktop and tablet

## Prerequisites

- **PowerShell 5.1+** (Windows PowerShell) or **PowerShell 7+** (cross-platform)
- An Entra ID account with sufficient privileges to read Intune configuration and group data
- The ability to consent to (or have an admin pre-consent) the following Microsoft Graph scopes:
  - `DeviceManagementApps.Read.All`
  - `DeviceManagementConfiguration.Read.All`
  - `DeviceManagementManagedDevices.Read.All`
  - `Group.Read.All`
  - `User.Read.All`

> The script installs the `Microsoft.Graph.Authentication` module automatically if it is not already present.

## Quick Start

```powershell
# Clone the repository
git clone https://github.com/<your-org>/Intune-Assignment-Checker.git
cd Intune-Assignment-Checker/src

# Run the script
.\IntuneAssignmentChecker.ps1
```

The script will:

1. Install the `Microsoft.Graph.Authentication` module (first run only).
2. Open a browser window for interactive Entra ID sign-in.
3. Request the required Graph permissions (consent prompt).
4. Start a local web server on **http://localhost:8080** and open it in your default browser.

### Custom Port

```powershell
.\IntuneAssignmentChecker.ps1 -Port 9090
```

## How It Works

```
┌─────────────────────────────────────────────────────────────────┐
│  Browser (http://localhost:8080)                                │
│  ┌──────────────┐  ┌────────────────────────────────────────┐  │
│  │ Entra Groups  │  │  Configurations | Settings Catalog |  │  │
│  │ (sidebar)     │  │  Applications   | Scripts          |  │  │
│  │               │  │  Remediations                       |  │  │
│  └──────────────┘  └────────────────────────────────────────┘  │
└───────────────────────────┬─────────────────────────────────────┘
                            │ /api/groups
                            │ /api/groups/{id}/assignments
                            ▼
┌───────────────────────────────────────┐
│  PowerShell HTTP Listener             │
│  IntuneAssignmentChecker.ps1          │
│  (serves UI + proxies Graph calls)    │
└───────────────────────┬───────────────┘
                        │ Invoke-MgGraphRequest
                        ▼
┌───────────────────────────────────────┐
│  Microsoft Graph API (beta)           │
│  graph.microsoft.com                  │
└───────────────────────────────────────┘
```

## Project Structure

```
src/
├── IntuneAssignmentChecker.ps1   # Main script (auth, HTTP server, Graph queries)
├── static/
│   ├── css/
│   │   └── style.css             # UI theme (base color #6969e9)
│   └── js/
│       └── app.js                # Frontend logic
└── templates/
    └── index.html                # Single-page application template
```

## Security Notes

- Authentication uses **interactive delegated flow** — no secrets are stored anywhere
- The script only requests **read** permissions; it cannot modify your Intune environment
- All Intune and Entra data should be treated as sensitive
- Press **Ctrl+C** to stop the server; the script disconnects from Microsoft Graph automatically
