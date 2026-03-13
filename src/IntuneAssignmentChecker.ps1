<#
.SYNOPSIS
    Intune Assignment Checker - interactive web dashboard for Intune group assignments.

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

# -----------------------------------------------------------------------------
# 1. Ensure the Microsoft Graph Authentication module is available
# -----------------------------------------------------------------------------

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

# -----------------------------------------------------------------------------
# 2. Connect to Microsoft Graph with required scopes (interactive sign-in)
# -----------------------------------------------------------------------------

$requiredScopes = @(
    "DeviceManagementConfiguration.Read.All"
    "DeviceManagementApps.Read.All"
    "DeviceManagementManagedDevices.Read.All"
    "Group.Read.All"
    "User.Read.All"
)

Write-Host ""
Write-Host "  ======================================================" -ForegroundColor Magenta
Write-Host "         Intune Assignment Checker                      " -ForegroundColor Magenta
Write-Host "  ======================================================" -ForegroundColor Magenta
Write-Host "  A browser window will open for sign-in.               " -ForegroundColor Magenta
Write-Host "  Sign in with your Entra ID credentials and            " -ForegroundColor Magenta
Write-Host "  accept the requested permissions.                     " -ForegroundColor Magenta
Write-Host "  ======================================================" -ForegroundColor Magenta
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

# -----------------------------------------------------------------------------
# 3. Graph API helper functions
# -----------------------------------------------------------------------------

function Invoke-GraphPaginated {
    <#
    .SYNOPSIS
        Fetches all pages from a Microsoft Graph endpoint.
    #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [switch]$SilentErrors
    )

    $all = [System.Collections.ArrayList]::new()
    $nextUri = $Uri

    while ($nextUri) {
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject
        }
        catch {
            $safeUri = ($nextUri -split '\?')[0]
            Write-Warning "Graph request failed for $safeUri : $($_.Exception.Message)"
            if ($SilentErrors) { break }
            throw
        }

        if ($response.value) {
            $response.value | ForEach-Object { [void]$all.Add($_) }
        }

        # Safely check for next page link — the property may not exist
        # on the response object depending on output type and PS version.
        if ($response -is [hashtable]) {
            $nextUri = $response['@odata.nextLink']
        } elseif ($response.PSObject.Properties.Match('@odata.nextLink').Count) {
            $nextUri = $response.'@odata.nextLink'
        } else {
            $nextUri = $null
        }
    }

    # Return items through the pipeline individually.
    # Callers should use @() to collect results into an array.
    $all.ToArray()
}

function Get-AllGroups {
    # Note: $orderby on /groups requires ConsistencyLevel:eventual + $count=true
    # which complicates pagination. Sort client-side instead.
    $uri = "/v1.0/groups?`$select=id,displayName,description,groupTypes,membershipRule&`$top=999"
    $groups = @(Invoke-GraphPaginated -Uri $uri)
    if ($groups.Count -gt 0) {
        $groups = @($groups | Sort-Object { $_.displayName })
    }
    $groups
}

function Get-SafeValue {
    <#
    .SYNOPSIS
        Safely reads a property from an object that may be a hashtable or PSObject.
        Returns $null when the key/property does not exist, avoiding StrictMode errors.
    #>
    param($Object, [string]$Key)

    if ($null -eq $Object) { return $null }
    if ($Object -is [hashtable]) { return $Object[$Key] }
    $prop = $Object.PSObject.Properties.Match($Key)
    if ($prop.Count) { return $prop[0].Value }
    return $null
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
        $items = Invoke-GraphPaginated -Uri $cat.Value -SilentErrors
        $matched = [System.Collections.ArrayList]::new()

        foreach ($item in $items) {
            $itemAssignments = Get-SafeValue $item 'assignments'
            if (-not $itemAssignments) { continue }

            foreach ($assignment in $itemAssignments) {
                $target = Get-SafeValue $assignment 'target'
                if (-not $target) { continue }

                $targetGroupId = Get-SafeValue $target 'groupId'
                if ($targetGroupId -eq $GroupId) {
                    $targetType = Get-SafeValue $target '@odata.type'
                    $friendly = switch ($targetType) {
                        "#microsoft.graph.groupAssignmentTarget"            { "Include" }
                        "#microsoft.graph.exclusionGroupAssignmentTarget"   { "Exclude" }
                        "#microsoft.graph.allDevicesAssignmentTarget"       { "All Devices" }
                        "#microsoft.graph.allLicensedUsersAssignmentTarget" { "All Users" }
                        default { $targetType }
                    }

                    $itemDisplayName = Get-SafeValue $item 'displayName'
                    $itemName        = Get-SafeValue $item 'name'
                    $displayName     = if ($itemDisplayName) { $itemDisplayName } elseif ($itemName) { $itemName } else { "N/A" }
                    $itemDesc        = Get-SafeValue $item 'description'
                    $assignIntent    = Get-SafeValue $assignment 'intent'
                    $filterId        = Get-SafeValue $target 'deviceAndAppManagementAssignmentFilterId'
                    $filterType      = Get-SafeValue $target 'deviceAndAppManagementAssignmentFilterType'

                    [void]$matched.Add(@{
                        id              = Get-SafeValue $item 'id'
                        displayName     = $displayName
                        description     = if ($itemDesc) { $itemDesc } else { "" }
                        assignmentType  = $friendly
                        intent          = if ($assignIntent) { $assignIntent } else { "" }
                        filterId        = if ($filterId) { $filterId } else { "" }
                        filterType      = if ($filterType) { $filterType } else { "" }
                    })
                    break   # one match per item is enough
                }
            }
        }

        $result[$cat.Key] = @($matched.ToArray())
    }

    return $result
}

function Get-AssignedGroupIds {
    <#
    .SYNOPSIS
        Returns a list of unique group IDs that appear as assignment targets
        across all Intune policy categories.
    #>
    $endpoints = @(
        "/beta/deviceManagement/deviceConfigurations?`$expand=assignments&`$select=id,assignments"
        "/beta/deviceManagement/configurationPolicies?`$expand=assignments&`$select=id,assignments"
        "/beta/deviceAppManagement/mobileApps?`$expand=assignments&`$filter=isAssigned eq true&`$select=id,assignments"
        "/beta/deviceManagement/deviceManagementScripts?`$expand=assignments&`$select=id,assignments"
        "/beta/deviceManagement/deviceHealthScripts?`$expand=assignments&`$select=id,assignments"
    )

    $ids = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($uri in $endpoints) {
        $items = Invoke-GraphPaginated -Uri $uri -SilentErrors
        foreach ($item in $items) {
            $itemAssignments = Get-SafeValue $item 'assignments'
            if (-not $itemAssignments) { continue }
            foreach ($assignment in $itemAssignments) {
                $target  = Get-SafeValue $assignment 'target'
                if (-not $target) { continue }
                $gid = Get-SafeValue $target 'groupId'
                if ($gid) { [void]$ids.Add($gid) }
            }
        }
    }

    @($ids)
}

# -----------------------------------------------------------------------------
# 4. JSON serialization helper
# -----------------------------------------------------------------------------

function ConvertTo-SafeJson {
    param($InputObject, [switch]$AsArray)

    # -AsArray: guarantee the output is always a JSON array, regardless
    # of PowerShell's array-unwrapping quirks (0 items → "[]",
    # 1 item → "[{…}]", N items → "[{…},{…},…]").
    if ($AsArray) {
        if ($null -eq $InputObject) { return "[]" }
        $arr = @($InputObject)
        if ($arr.Count -eq 0) { return "[]" }
        $json = ConvertTo-Json -InputObject $arr -Depth 10 -Compress
        # Guard: some PS versions still unwrap single-element arrays
        if ($json[0] -ne '[') { $json = "[$json]" }
        return $json
    }

    if ($null -eq $InputObject) { return "null" }
    return (ConvertTo-Json -InputObject $InputObject -Depth 10 -Compress)
}

function Set-SecurityHeaders {
    <#
    .SYNOPSIS
        Adds standard security headers to an HTTP response.
    #>
    param(
        [Parameter(Mandatory)]
        [System.Net.HttpListenerResponse]$Response
    )

    $Response.Headers.Set("X-Content-Type-Options", "nosniff")
    $Response.Headers.Set("X-Frame-Options", "DENY")
    $Response.Headers.Set("Referrer-Policy", "strict-origin-when-cross-origin")
    $csp = "default-src 'none'; script-src 'self' https://alcdn.msauth.net; style-src 'self'; img-src 'self'; " +
           "connect-src 'self' https://graph.microsoft.com https://login.microsoftonline.com; font-src 'self'; frame-ancestors 'none'; base-uri 'self'; form-action 'self'"
    $Response.Headers.Set("Content-Security-Policy", $csp)
}

# -----------------------------------------------------------------------------
# 5. Resolve static file paths
# -----------------------------------------------------------------------------

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

# -----------------------------------------------------------------------------
# 6. Start the HTTP listener
# -----------------------------------------------------------------------------

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

Write-Host "  ======================================================" -ForegroundColor Cyan
Write-Host "  Web server running at http://localhost:$Port          " -ForegroundColor Cyan
Write-Host "  Press Ctrl+C to stop.                                " -ForegroundColor Cyan
Write-Host "  ======================================================" -ForegroundColor Cyan
Write-Host ""

# Open default browser
try {
    Start-Process "http://localhost:$Port"
}
catch {
    Write-Host "  Open http://localhost:$Port in your browser." -ForegroundColor Yellow
}

# -----------------------------------------------------------------------------
# 7. Request loop
# -----------------------------------------------------------------------------

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
            Set-SecurityHeaders -Response $resp

            # -- API: list groups ------------------------------------
            if ($path -eq "/api/groups" -and $req.HttpMethod -eq "GET") {
                $groups = @(Get-AllGroups)
                $json   = ConvertTo-SafeJson -InputObject $groups -AsArray
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                $resp.ContentType     = "application/json; charset=utf-8"
                $resp.ContentLength64 = $buffer.Length
                $resp.StatusCode      = 200
                $resp.OutputStream.Write($buffer, 0, $buffer.Length)
            }
            # -- API: assigned group IDs -----------------------------
            elseif ($path -eq "/api/assigned-group-ids" -and $req.HttpMethod -eq "GET") {
                try {
                    $groupIds = @(Get-AssignedGroupIds)
                    $json   = ConvertTo-SafeJson -InputObject $groupIds -AsArray
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 200
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
                catch {
                    Write-Warning "Failed to fetch assigned group IDs: $($_.Exception.Message)"
                    $errMsg  = $_.Exception.Message -replace '"', '\"'
                    $errBody = "{`"error`":`"Failed to fetch assigned group IDs: $errMsg`"}"
                    $buffer  = [System.Text.Encoding]::UTF8.GetBytes($errBody)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 502
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
            }
            # -- API: group assignments ------------------------------
            elseif ($path -match "^/api/groups/([^/]+)/assignments$" -and $req.HttpMethod -eq "GET") {
                $groupId    = $Matches[1]
                $parsedGuid = [System.Guid]::Empty
                if (-not [System.Guid]::TryParse($groupId, [ref]$parsedGuid)) {
                    $resp.StatusCode = 400
                    $body   = '{"error":"Invalid group ID format. Expected a valid GUID."}'
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($body)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                    continue
                }
                $groupId = $parsedGuid.ToString()
                try {
                    $assignmentsResult = Get-AssignmentsForGroup -GroupId $groupId
                    $json   = ConvertTo-SafeJson -InputObject $assignmentsResult
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 200
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
                catch {
                    Write-Warning "Assignment fetch failed for group $groupId : $($_.Exception.Message)"
                    $errMsg  = $_.Exception.Message -replace '"', '\"'
                    $errBody = "{`"error`":`"Failed to fetch assignments: $errMsg`"}"
                    $buffer  = [System.Text.Encoding]::UTF8.GetBytes($errBody)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 502
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
            }
            # -- API: script content ---------------------------------
            elseif ($path -match "^/api/scripts/([^/]+)/content$" -and $req.HttpMethod -eq "GET") {
                $scriptId   = $Matches[1]
                $parsedGuid = [System.Guid]::Empty
                if (-not [System.Guid]::TryParse($scriptId, [ref]$parsedGuid)) {
                    $resp.StatusCode = 400
                    $body   = '{"error":"Invalid script ID format. Expected a valid GUID."}'
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($body)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                    continue
                }
                $scriptId = $parsedGuid.ToString()
                try {
                    $scriptObj  = Invoke-MgGraphRequest -Method GET -Uri "/beta/deviceManagement/deviceManagementScripts/$scriptId" -OutputType PSObject
                    $b64Content = Get-SafeValue $scriptObj 'scriptContent'
                    $fileName   = Get-SafeValue $scriptObj 'fileName'
                    $decoded    = ""
                    if ($b64Content) {
                        $decoded = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($b64Content))
                    }
                    $result = @{
                        id          = $scriptId
                        fileName    = if ($fileName) { $fileName } else { "" }
                        content     = $decoded
                    }
                    $json   = ConvertTo-SafeJson -InputObject $result
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 200
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
                catch {
                    Write-Warning "Script content fetch failed for $scriptId : $($_.Exception.Message)"
                    $errMsg  = $_.Exception.Message -replace '"', '\"'
                    $errBody = "{`"error`":`"Failed to fetch script content: $errMsg`"}"
                    $buffer  = [System.Text.Encoding]::UTF8.GetBytes($errBody)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 502
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
            }
            # -- API: logout -----------------------------------------
            elseif ($path -eq "/api/logout" -and $req.HttpMethod -eq "POST") {
                try {
                    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                    Write-Host "  User logged out via web UI." -ForegroundColor Yellow
                    $body   = '{"success":true,"message":"Disconnected from Microsoft Graph. Restart the script to sign in again."}'
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($body)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 200
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
                catch {
                    $resp.StatusCode = 500
                    $body   = '{"error":"Failed to disconnect."}'
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($body)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
            }
            # -- Serve index.html ------------------------------------
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
            # -- Serve static files ----------------------------------
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
            # -- 404 ------------------------------------------------
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
                Set-SecurityHeaders -Response $resp
                $resp.StatusCode = 500
                $errBody = '{"error":"An internal error occurred. Please try again later."}'
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
    if ($null -ne (Get-Variable -Name listener -ValueOnly -ErrorAction SilentlyContinue)) {
        try { $listener.Stop() } catch { }
        try { $listener.Close() } catch { }
    }
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Host "  Disconnected from Microsoft Graph. Goodbye!" -ForegroundColor Green
}
