<#
.SYNOPSIS
    Intune Assignment Checker — interactive web dashboard for Intune group assignments.

.DESCRIPTION
    This script:
      1. Installs the Microsoft.Graph.Authentication module if missing.
      2. Prompts the user to sign in with their Entra ID credentials (interactive browser flow).
      3. Automatically requests the required Microsoft Graph permissions (consent prompt).
      4. Starts a local web server and opens the dashboard in the default browser.

    No manual app registrations, client secrets, or portal configuration required.

.NOTES
    Requires PowerShell 5.1+ (Windows PowerShell) or PowerShell 7+ (cross-platform).
    The signed-in user must be able to consent to (or have an admin pre-consent)
    the listed Graph scopes.
#>

[CmdletBinding()]
param(
    [int]$Port = 8080
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ─────────────────────────────────────────────────────────────────────────────
# 1. Ensure the Microsoft Graph Authentication module is available
# ─────────────────────────────────────────────────────────────────────────────

$moduleName = "Microsoft.Graph.Authentication"

if (-not (Get-Module -ListAvailable -Name $moduleName)) {
    Write-Host ""
    Write-Host "  Installing $moduleName module (one-time setup)..." -ForegroundColor Cyan
    Write-Host ""
    try {
        Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
        Write-Host "  Module installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install $moduleName. Please run: Install-Module $moduleName -Scope CurrentUser"
        exit 1
    }
}

Import-Module $moduleName -ErrorAction Stop

# ─────────────────────────────────────────────────────────────────────────────
# 2. Connect to Microsoft Graph with required scopes (interactive sign-in)
# ─────────────────────────────────────────────────────────────────────────────

$requiredScopes = @(
    "DeviceManagementConfiguration.Read.All"
    "DeviceManagementApps.Read.All"
    "DeviceManagementManagedDevices.Read.All"
    "Group.Read.All"
    "User.Read.All"
)

Write-Host ""
Write-Host "  ┌──────────────────────────────────────────────────┐" -ForegroundColor Magenta
Write-Host "  │       Intune Assignment Checker                  │" -ForegroundColor Magenta
Write-Host "  │                                                  │" -ForegroundColor Magenta
Write-Host "  │  A browser window will open for sign-in.         │" -ForegroundColor Magenta
Write-Host "  │  Sign in with your Entra ID credentials and      │" -ForegroundColor Magenta
Write-Host "  │  accept the requested permissions.               │" -ForegroundColor Magenta
Write-Host "  └──────────────────────────────────────────────────┘" -ForegroundColor Magenta
Write-Host ""

try {
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    $context = Get-MgContext
    Write-Host "  Signed in as: $($context.Account)" -ForegroundColor Green
    Write-Host "  Tenant:       $($context.TenantId)" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Error "Authentication failed: $_"
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
# 3. Graph API helper functions
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-GraphPaginated {
    <#
    .SYNOPSIS
        Fetches all pages from a Microsoft Graph endpoint.
    #>
    param(
        [Parameter(Mandatory)][string]$Uri
    )

    $all = @()
    $nextUri = $Uri

    while ($nextUri) {
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject
        }
        catch {
            Write-Warning "Graph request failed for $nextUri : $_"
            break
        }

        if ($response.value) {
            $all += $response.value
        }

        $nextUri = $response.'@odata.nextLink'
    }

    return $all
}

function Get-AllGroups {
    $uri = "/v1.0/groups?`$select=id,displayName,description,groupTypes,membershipRule&`$orderby=displayName&`$top=999"
    return Invoke-GraphPaginated -Uri $uri
}

function Get-AssignmentsForGroup {
    param([string]$GroupId)

    $categories = @{
        configurations  = "/beta/deviceManagement/deviceConfigurations?`$expand=assignments"
        settingsCatalog = "/beta/deviceManagement/configurationPolicies?`$expand=assignments"
        applications    = "/beta/deviceAppManagement/mobileApps?`$expand=assignments&`$filter=isAssigned eq true"
        scripts         = "/beta/deviceManagement/deviceManagementScripts?`$expand=assignments"
        remediations    = "/beta/deviceManagement/deviceHealthScripts?`$expand=assignments"
    }

    $result = @{}

    foreach ($cat in $categories.GetEnumerator()) {
        $items = Invoke-GraphPaginated -Uri $cat.Value
        $matched = @()

        foreach ($item in $items) {
            if (-not $item.assignments) { continue }

            foreach ($assignment in $item.assignments) {
                $target = $assignment.target
                if (-not $target) { continue }

                if ($target.groupId -eq $GroupId) {
                    $targetType = $target.'@odata.type'
                    $friendly = switch ($targetType) {
                        "#microsoft.graph.groupAssignmentTarget"            { "Include" }
                        "#microsoft.graph.exclusionGroupAssignmentTarget"   { "Exclude" }
                        "#microsoft.graph.allDevicesAssignmentTarget"       { "All Devices" }
                        "#microsoft.graph.allLicensedUsersAssignmentTarget" { "All Users" }
                        default { $targetType }
                    }

                    $displayName = if ($item.displayName) { $item.displayName } elseif ($item.name) { $item.name } else { "N/A" }

                    $matched += @{
                        id              = $item.id
                        displayName     = $displayName
                        description     = if ($item.description) { $item.description } else { "" }
                        assignmentType  = $friendly
                        intent          = if ($assignment.intent) { $assignment.intent } else { "" }
                        filterId        = if ($target.deviceAndAppManagementAssignmentFilterId) { $target.deviceAndAppManagementAssignmentFilterId } else { "" }
                        filterType      = if ($target.deviceAndAppManagementAssignmentFilterType) { $target.deviceAndAppManagementAssignmentFilterType } else { "" }
                    }
                    break   # one match per item is enough
                }
            }
        }

        $result[$cat.Key] = $matched
    }

    return $result
}

# ─────────────────────────────────────────────────────────────────────────────
# 4. JSON serialization helper
# ─────────────────────────────────────────────────────────────────────────────

function ConvertTo-SafeJson {
    param($InputObject)
    return ($InputObject | ConvertTo-Json -Depth 10 -Compress)
}

# ─────────────────────────────────────────────────────────────────────────────
# 5. Resolve static file paths
# ─────────────────────────────────────────────────────────────────────────────

$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }

$staticRoot   = Join-Path $scriptDir "static"
$templateRoot = Join-Path $scriptDir "templates"

$mimeTypes = @{
    ".html" = "text/html; charset=utf-8"
    ".css"  = "text/css; charset=utf-8"
    ".js"   = "application/javascript; charset=utf-8"
    ".json" = "application/json; charset=utf-8"
    ".png"  = "image/png"
    ".svg"  = "image/svg+xml"
    ".ico"  = "image/x-icon"
}

# ─────────────────────────────────────────────────────────────────────────────
# 6. Start the HTTP listener
# ─────────────────────────────────────────────────────────────────────────────

$listener = New-Object System.Net.HttpListener
$prefix   = "http://localhost:$Port/"
$listener.Prefixes.Add($prefix)

try {
    $listener.Start()
}
catch {
    Write-Error "Could not start HTTP listener on port $Port. Is it already in use? Error: $_"
    exit 1
}

Write-Host "  ┌──────────────────────────────────────────────────┐" -ForegroundColor Cyan
Write-Host "  │  Web server running at http://localhost:$Port     │" -ForegroundColor Cyan
Write-Host "  │  Press Ctrl+C to stop.                           │" -ForegroundColor Cyan
Write-Host "  └──────────────────────────────────────────────────┘" -ForegroundColor Cyan
Write-Host ""

# Open default browser
try {
    Start-Process "http://localhost:$Port"
}
catch {
    Write-Host "  Open http://localhost:$Port in your browser." -ForegroundColor Yellow
}

# ─────────────────────────────────────────────────────────────────────────────
# 7. Request loop
# ─────────────────────────────────────────────────────────────────────────────

try {
    while ($listener.IsListening) {
        $contextTask = $listener.GetContextAsync()
        # Allow Ctrl+C to interrupt
        while (-not $contextTask.AsyncWaitHandle.WaitOne(500)) { }
        $ctx = $contextTask.GetAwaiter().GetResult()

        $req  = $ctx.Request
        $resp = $ctx.Response
        $path = $req.Url.AbsolutePath

        try {
            # ── API: list groups ─────────────────────────────────
            if ($path -eq "/api/groups" -and $req.HttpMethod -eq "GET") {
                $groups = Get-AllGroups
                $json   = ConvertTo-SafeJson -InputObject $groups
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                $resp.ContentType     = "application/json; charset=utf-8"
                $resp.ContentLength64 = $buffer.Length
                $resp.StatusCode      = 200
                $resp.OutputStream.Write($buffer, 0, $buffer.Length)
            }
            # ── API: group assignments ───────────────────────────
            elseif ($path -match "^/api/groups/([^/]+)/assignments$" -and $req.HttpMethod -eq "GET") {
                $groupId     = $Matches[1]
                $assignments = Get-AssignmentsForGroup -GroupId $groupId
                $json        = ConvertTo-SafeJson -InputObject $assignments
                $buffer      = [System.Text.Encoding]::UTF8.GetBytes($json)
                $resp.ContentType     = "application/json; charset=utf-8"
                $resp.ContentLength64 = $buffer.Length
                $resp.StatusCode      = 200
                $resp.OutputStream.Write($buffer, 0, $buffer.Length)
            }
            # ── Serve index.html ─────────────────────────────────
            elseif ($path -eq "/" -or $path -eq "/index.html") {
                $filePath = Join-Path $templateRoot "index.html"
                if (Test-Path $filePath) {
                    $bytes = [System.IO.File]::ReadAllBytes($filePath)
                    $resp.ContentType     = "text/html; charset=utf-8"
                    $resp.ContentLength64 = $bytes.Length
                    $resp.StatusCode      = 200
                    $resp.OutputStream.Write($bytes, 0, $bytes.Length)
                }
                else {
                    $resp.StatusCode = 404
                }
            }
            # ── Serve static files ───────────────────────────────
            elseif ($path.StartsWith("/static/")) {
                $relativePath = $path.Substring("/static/".Length).Replace("/", [System.IO.Path]::DirectorySeparatorChar)
                $filePath     = Join-Path $staticRoot $relativePath

                # Prevent directory traversal
                $resolvedPath = [System.IO.Path]::GetFullPath($filePath)
                $resolvedRoot = [System.IO.Path]::GetFullPath($staticRoot)

                if ($resolvedPath.StartsWith($resolvedRoot) -and (Test-Path $resolvedPath -PathType Leaf)) {
                    $ext  = [System.IO.Path]::GetExtension($resolvedPath).ToLower()
                    $mime = if ($mimeTypes.ContainsKey($ext)) { $mimeTypes[$ext] } else { "application/octet-stream" }
                    $bytes = [System.IO.File]::ReadAllBytes($resolvedPath)
                    $resp.ContentType     = $mime
                    $resp.ContentLength64 = $bytes.Length
                    $resp.StatusCode      = 200
                    $resp.OutputStream.Write($bytes, 0, $bytes.Length)
                }
                else {
                    $resp.StatusCode = 404
                }
            }
            # ── 404 ──────────────────────────────────────────────
            else {
                $resp.StatusCode = 404
                $body   = '{"error":"Not found"}'
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($body)
                $resp.ContentType     = "application/json"
                $resp.ContentLength64 = $buffer.Length
                $resp.OutputStream.Write($buffer, 0, $buffer.Length)
            }
        }
        catch {
            Write-Warning "Request error ($path): $_"
            try {
                $resp.StatusCode = 500
                $errBody = "{`"error`":`"$($_.Exception.Message -replace '"','\"')`"}"
                $buffer  = [System.Text.Encoding]::UTF8.GetBytes($errBody)
                $resp.ContentType     = "application/json"
                $resp.ContentLength64 = $buffer.Length
                $resp.OutputStream.Write($buffer, 0, $buffer.Length)
            }
            catch { }
        }
        finally {
            $resp.OutputStream.Close()
        }
    }
}
finally {
    Write-Host ""
    Write-Host "  Shutting down..." -ForegroundColor Yellow
    $listener.Stop()
    $listener.Close()
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Host "  Disconnected from Microsoft Graph. Goodbye!" -ForegroundColor Green
}
