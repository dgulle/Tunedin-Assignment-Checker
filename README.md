# Intune Assignment Checker

A web-based tool that connects to Microsoft Intune via the Microsoft Graph API and displays policy and application assignments for Entra ID (Azure AD) groups.

## Features

- Browse all Entra ID groups in a searchable sidebar
- View Intune assignments per group across five categories:
  - **Device Configurations** — Configuration profiles
  - **Settings Catalog** — Settings Catalog policies
  - **Applications** — Assigned apps (required, available, uninstall)
  - **Scripts** — PowerShell device management scripts
  - **Remediations** — Proactive remediation (health) scripts
- See assignment type (Include / Exclude) and filter information
- Responsive design that works on desktop and tablet

## Prerequisites

1. **Python 3.9+**
2. **Azure AD App Registration** with the following **Application** permissions:
   - `DeviceManagementApps.Read.All`
   - `DeviceManagementConfiguration.Read.All`
   - `DeviceManagementManagedDevices.Read.All`
   - `Group.Read.All`
   - `User.Read.All`
3. Admin consent granted for the above permissions

## Quick Start

```bash
# Clone the repository
git clone https://github.com/<your-org>/Intune-Assignment-Checker.git
cd Intune-Assignment-Checker/src

# Create a virtual environment & install dependencies
python -m venv venv
source venv/bin/activate      # Linux / macOS
# venv\Scripts\activate       # Windows
pip install -r requirements.txt

# Configure credentials
cp .env.example .env
# Edit .env and fill in your Azure AD tenant ID, client ID, and client secret

# Run the app
python app.py
```

Open your browser to **http://localhost:5000**.

## Configuration

| Variable | Description |
|---|---|
| `AZURE_TENANT_ID` | Azure AD tenant ID |
| `AZURE_CLIENT_ID` | App registration client / application ID |
| `AZURE_CLIENT_SECRET` | App registration client secret |
| `PORT` | Web server port (default `5000`) |
| `FLASK_DEBUG` | Set to `true` for development hot-reload |

## Project Structure

```
src/
├── app.py              # Flask web server & API routes
├── graph_client.py     # Microsoft Graph API client
├── requirements.txt    # Python dependencies
├── .env.example        # Environment variable template
├── static/
│   ├── css/
│   │   └── style.css   # UI theme (base color #6969e9)
│   └── js/
│       └── app.js      # Frontend logic
└── templates/
    └── index.html      # Single-page application template
```

## Security Notes

- Never commit your `.env` file or any client secrets
- The app uses **client credentials flow** (application permissions) — keep the secret secure
- All Intune and Entra data should be treated as sensitive
