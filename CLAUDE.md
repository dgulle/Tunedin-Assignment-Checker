# CLAUDE.md

This file provides guidance to Claude Code when working with this repository.

## Project Overview

**Tunedin Assignment Checker** is a tool for querying and reporting Microsoft Intune policy/app assignments. It helps administrators discover what policies, applications, and configurations are assigned to users, devices, or groups within an Intune tenant.

## Repository Structure

```
Intune-Assignment-Checker/
├── CLAUDE.md          # This file
├── README.md          # Project documentation
└── src/               # Source code (to be added)
```

## Development Workflow

### Branch Strategy

- `master` / `main` — stable, production-ready code
- Feature branches — `feature/<description>`
- Fix branches — `fix/<description>`
- Always develop on a dedicated branch; never commit directly to `master`/`main`

### Commit Messages

Use clear, descriptive commit messages:

```
<type>: <short summary>

<optional body explaining why, not what>
```

Types: `feat`, `fix`, `docs`, `refactor`, `test`, `chore`

Example: `feat: add group-based assignment lookup`

### Git Operations

```bash
git fetch origin <branch-name>
git pull origin <branch-name>
git push -u origin <branch-name>
```

## Microsoft Graph API / Intune

This project likely integrates with the **Microsoft Graph API** to query Intune data. Key points:

- Authentication uses **Entra ID app registration** with appropriate Intune/Graph permissions
- Required API permissions (application or delegated):
  - `DeviceManagementApps.Read.All`
  - `DeviceManagementConfiguration.Read.All`
  - `DeviceManagementManagedDevices.Read.All`
  - `Group.Read.All`
  - `User.Read.All`
- Base URL: `https://graph.microsoft.com/v1.0/` or `/beta/`
- Avoid storing credentials or secrets in source code — use environment variables or a secrets manager

## Code Conventions

- Keep secrets out of source control (`.env`, key files, certificates)
- Validate and sanitize all inputs, especially IDs passed to API calls
- Handle API pagination (`@odata.nextLink`) when fetching large result sets
- Respect API throttling — implement retry logic with exponential backoff on HTTP 429 responses
- Log errors with enough context to diagnose issues without exposing sensitive data

## Running / Testing

> Build, run, and test instructions will be added once the project stack is established.

Typical steps will follow the pattern:

```bash
# Install dependencies
<package-manager> install

# Run the tool
<entry-point> --help

# Run tests
<test-runner>
```

## Security Notes

- Never commit `.env` files, client secrets, certificates, or access tokens
- Add sensitive file patterns to `.gitignore` immediately
- Treat all Intune/AAD data as sensitive — do not log full API responses in production
