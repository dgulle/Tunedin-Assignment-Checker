<#
.SYNOPSIS
    Tunedin Assignment Checker - interactive web dashboard for Intune group assignments.

.DESCRIPTION
    This script:
      1. Installs the Microsoft.Graph.Authentication module if missing.
      2. Connects to Microsoft Graph via the "Microsoft Graph Command Line Tools"
         enterprise application (no custom app registration required).
      3. Opens a browser for interactive sign-in and consent to the required
         permissions (only on first use or when new scopes are added).
      4. Starts a local web server and opens the dashboard in the default browser.

    No app registrations, client secrets, or portal configuration required.

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
        Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -MinimumVersion 2.0.0
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
    "DeviceManagementApps.Read.All"
    "DeviceManagementConfiguration.Read.All"
    "DeviceManagementManagedDevices.Read.All"
    "DeviceManagementScripts.Read.All"
    "Group.Read.All"
    "User.Read"
    "User.Read.All"
)

Write-Host ""
Write-Host "  ======================================================" -ForegroundColor Magenta
Write-Host "         Tunedin Assignment Checker                     " -ForegroundColor Magenta
Write-Host "  ======================================================" -ForegroundColor Magenta
Write-Host "  No app registration required.                         " -ForegroundColor Magenta
Write-Host "  Permissions are requested via Microsoft Graph          " -ForegroundColor Magenta
Write-Host "  Command Line Tools.                                   " -ForegroundColor Magenta
Write-Host "                                                        " -ForegroundColor Magenta
Write-Host "  A browser window will open for sign-in.               " -ForegroundColor Magenta
Write-Host "  You may be prompted to consent to permissions         " -ForegroundColor Magenta
Write-Host "  on first use.                                         " -ForegroundColor Magenta
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

    $maxRetries = 3

    while ($nextUri) {
        $response = $null
        for ($attempt = 0; $attempt -le $maxRetries; $attempt++) {
            try {
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject
                break
            }
            catch {
                $statusCode = 0
                if ($_.Exception.Response) {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
                # Retry on 429 (throttled) or 5xx (server error)
                if (($statusCode -eq 429 -or $statusCode -ge 500) -and $attempt -lt $maxRetries) {
                    $delay = [Math]::Pow(2, $attempt + 1)
                    Write-Warning "Graph API $statusCode on attempt $($attempt + 1), retrying in ${delay}s"
                    Start-Sleep -Seconds $delay
                    continue
                }
                $safeUri = ($nextUri -split '\?')[0]
                Write-Warning "Graph request failed for $safeUri (HTTP $statusCode)"
                if ($SilentErrors) { $nextUri = $null; break }
                throw
            }
        }
        if (-not $response) { break }

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

    $result  = @{}
    $_errors = @{}

    foreach ($cat in $categories.GetEnumerator()) {
        $matched = [System.Collections.ArrayList]::new()
        try {
            $items = @(Invoke-GraphPaginated -Uri $cat.Value)
        }
        catch {
            Write-Warning "Category $($cat.Key) failed: $($_.Exception.Message)"
            $_errors[$cat.Key] = $_.Exception.Message
            $result[$cat.Key]  = @()
            continue
        }

        foreach ($item in $items) {
            $itemAssignments = Get-SafeValue $item 'assignments'
            if (-not $itemAssignments) { continue }

            foreach ($assignment in $itemAssignments) {
                $target = Get-SafeValue $assignment 'target'
                if (-not $target) { continue }

                $targetGroupId = Get-SafeValue $target 'groupId'
                $targetType    = Get-SafeValue $target '@odata.type'

                # Match: group assignment to this group, OR All Devices, OR All Users
                $isGroupMatch     = ($targetGroupId -eq $GroupId)
                $isAllDevices     = ($targetType -eq '#microsoft.graph.allDevicesAssignmentTarget')
                $isAllUsers       = ($targetType -eq '#microsoft.graph.allLicensedUsersAssignmentTarget')

                if ($isGroupMatch -or $isAllDevices -or $isAllUsers) {
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
                    # Don't break — same item may match as both group + All Devices/Users
                }
            }
        }

        $result[$cat.Key] = @($matched.ToArray())
    }

    $result['_errors'] = $_errors
    return $result
}

function Get-GroupParentGroups {
    <#
    .SYNOPSIS
        Returns the transitive group memberships for a given group.
        This reveals which parent groups this group is nested within.
    #>
    param([Parameter(Mandatory)][string]$GroupId)

    $uri = "/v1.0/groups/$GroupId/transitiveMemberOf/microsoft.graph.group?`$select=id,displayName&`$top=999"
    $parents = @(Invoke-GraphPaginated -Uri $uri -SilentErrors)
    return $parents
}

function Get-NestedGroupAssignments {
    <#
    .SYNOPSIS
        For a given group, finds all assignments that come through parent group
        memberships (nested/inherited assignments).
    #>
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][array]$ParentGroups
    )

    $categories = @{
        configurations  = "/beta/deviceManagement/deviceConfigurations?`$expand=assignments"
        settingsCatalog = "/beta/deviceManagement/configurationPolicies?`$expand=assignments"
        applications    = "/beta/deviceAppManagement/mobileApps?`$expand=assignments&`$filter=isAssigned eq true"
        scripts         = "/beta/deviceManagement/deviceManagementScripts?`$expand=assignments"
        remediations    = "/beta/deviceManagement/deviceHealthScripts?`$expand=assignments"
    }

    # Build a lookup of parent group IDs to names
    $parentLookup = @{}
    foreach ($pg in $ParentGroups) {
        $pgId = Get-SafeValue $pg 'id'
        $pgName = Get-SafeValue $pg 'displayName'
        if ($pgId) { $parentLookup[$pgId] = if ($pgName) { $pgName } else { $pgId } }
    }

    $result  = @{}
    $_errors = @{}

    foreach ($cat in $categories.GetEnumerator()) {
        $matched = [System.Collections.ArrayList]::new()
        try {
            $items = @(Invoke-GraphPaginated -Uri $cat.Value)
        }
        catch {
            $_errors[$cat.Key] = $_.Exception.Message
            $result[$cat.Key]  = @()
            continue
        }

        foreach ($item in $items) {
            $itemAssignments = Get-SafeValue $item 'assignments'
            if (-not $itemAssignments) { continue }

            foreach ($assignment in $itemAssignments) {
                $target = Get-SafeValue $assignment 'target'
                if (-not $target) { continue }

                $targetGroupId = Get-SafeValue $target 'groupId'
                $targetType    = Get-SafeValue $target '@odata.type'

                # Check if assignment targets a parent group
                if ($targetGroupId -and $parentLookup.ContainsKey($targetGroupId)) {
                    $friendly = switch ($targetType) {
                        "#microsoft.graph.groupAssignmentTarget"          { "Include" }
                        "#microsoft.graph.exclusionGroupAssignmentTarget" { "Exclude" }
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
                        inheritedFrom   = $parentLookup[$targetGroupId]
                        inheritedFromId = $targetGroupId
                    })
                }
            }
        }

        $result[$cat.Key] = @($matched.ToArray())
    }

    $result['_errors'] = $_errors
    return $result
}

function Get-OrphanedItems {
    <#
    .SYNOPSIS
        Returns all Intune items (policies, apps, scripts, remediations) that have
        zero assignments — i.e. orphaned items that may be candidates for cleanup.
    #>
    $categories = @{
        configurations  = "/beta/deviceManagement/deviceConfigurations?`$expand=assignments"
        settingsCatalog = "/beta/deviceManagement/configurationPolicies?`$expand=assignments"
        applications    = "/beta/deviceAppManagement/mobileApps?`$expand=assignments&`$select=id,displayName,description,assignments"
        scripts         = "/beta/deviceManagement/deviceManagementScripts?`$expand=assignments"
        remediations    = "/beta/deviceManagement/deviceHealthScripts?`$expand=assignments"
    }

    $result  = @{}
    $_errors = @{}

    foreach ($cat in $categories.GetEnumerator()) {
        $orphaned = [System.Collections.ArrayList]::new()
        try {
            $items = @(Invoke-GraphPaginated -Uri $cat.Value)
        }
        catch {
            Write-Warning "Orphaned check for $($cat.Key) failed: $($_.Exception.Message)"
            $_errors[$cat.Key] = $_.Exception.Message
            $result[$cat.Key]  = @()
            continue
        }

        foreach ($item in $items) {
            $itemAssignments = Get-SafeValue $item 'assignments'
            $assignCount = 0
            if ($itemAssignments) {
                $assignCount = @($itemAssignments).Count
            }

            if ($assignCount -eq 0) {
                $itemDisplayName = Get-SafeValue $item 'displayName'
                $itemName        = Get-SafeValue $item 'name'
                $displayName     = if ($itemDisplayName) { $itemDisplayName } elseif ($itemName) { $itemName } else { "N/A" }
                $itemDesc        = Get-SafeValue $item 'description'

                [void]$orphaned.Add(@{
                    id          = Get-SafeValue $item 'id'
                    displayName = $displayName
                    description = if ($itemDesc) { $itemDesc } else { "" }
                })
            }
        }

        $result[$cat.Key] = @($orphaned.ToArray())
    }

    $result['_errors'] = $_errors
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

    $counts = @{}

    foreach ($uri in $endpoints) {
        $items = Invoke-GraphPaginated -Uri $uri -SilentErrors
        foreach ($item in $items) {
            $itemAssignments = Get-SafeValue $item 'assignments'
            if (-not $itemAssignments) { continue }
            foreach ($assignment in $itemAssignments) {
                $target  = Get-SafeValue $assignment 'target'
                if (-not $target) { continue }
                $gid = Get-SafeValue $target 'groupId'
                if ($gid) {
                    if ($counts.ContainsKey($gid)) {
                        $counts[$gid] = $counts[$gid] + 1
                    } else {
                        $counts[$gid] = 1
                    }
                }
            }
        }
    }

    @{
        ids    = @($counts.Keys)
        counts = $counts
    }
}

function Get-AssignmentsByTargetType {
    param([string]$TargetOdataType)

    $categories = @{
        configurations  = "/beta/deviceManagement/deviceConfigurations?`$expand=assignments"
        settingsCatalog = "/beta/deviceManagement/configurationPolicies?`$expand=assignments"
        applications    = "/beta/deviceAppManagement/mobileApps?`$expand=assignments&`$filter=isAssigned eq true"
        scripts         = "/beta/deviceManagement/deviceManagementScripts?`$expand=assignments"
        remediations    = "/beta/deviceManagement/deviceHealthScripts?`$expand=assignments"
    }

    $friendly = switch ($TargetOdataType) {
        "#microsoft.graph.allDevicesAssignmentTarget"       { "All Devices" }
        "#microsoft.graph.allLicensedUsersAssignmentTarget" { "All Users" }
        default { $TargetOdataType }
    }

    $result  = @{}
    $_errors = @{}

    foreach ($cat in $categories.GetEnumerator()) {
        $matched = [System.Collections.ArrayList]::new()
        try {
            $items = @(Invoke-GraphPaginated -Uri $cat.Value)
        }
        catch {
            Write-Warning "Category $($cat.Key) failed: $($_.Exception.Message)"
            $_errors[$cat.Key] = $_.Exception.Message
            $result[$cat.Key]  = @()
            continue
        }

        foreach ($item in $items) {
            $itemAssignments = Get-SafeValue $item 'assignments'
            if (-not $itemAssignments) { continue }

            foreach ($assignment in $itemAssignments) {
                $target = Get-SafeValue $assignment 'target'
                if (-not $target) { continue }

                $targetType = Get-SafeValue $target '@odata.type'
                if ($targetType -ne $TargetOdataType) { continue }

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
            }
        }

        $result[$cat.Key] = @($matched.ToArray())
    }

    $result['_errors'] = $_errors
    return $result
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
    $Response.Headers.Set("Permissions-Policy", "geolocation=(), microphone=(), camera=(), usb=(), magnetometer=(), gyroscope=(), accelerometer=()")
    $Response.Headers.Set("X-Permitted-Cross-Domain-Policies", "none")
    $csp = "default-src 'none'; script-src 'self' https://cdn.jsdelivr.net; style-src 'self'; img-src 'self'; " +
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

            # -- API: status (backend mode detection) ------------------
            if ($path -eq "/api/status" -and ($req.HttpMethod -eq "GET" -or $req.HttpMethod -eq "HEAD")) {
                $statusBody = '{"mode":"backend"}'
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($statusBody)
                $resp.ContentType     = "application/json; charset=utf-8"
                $resp.ContentLength64 = $buffer.Length
                $resp.StatusCode      = 200
                $resp.OutputStream.Write($buffer, 0, $buffer.Length)
            }
            # -- API: list groups ------------------------------------
            elseif ($path -eq "/api/groups" -and $req.HttpMethod -eq "GET") {
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
                    $result = Get-AssignedGroupIds
                    $json   = ConvertTo-Json -InputObject $result -Depth 10 -Compress
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 200
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
                catch {
                    Write-Warning "Failed to fetch assigned group IDs: $($_.Exception.Message)"
                    $errBody = ConvertTo-Json -InputObject @{ error = "Failed to fetch assigned group IDs. Please try again." } -Compress
                    $buffer  = [System.Text.Encoding]::UTF8.GetBytes($errBody)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 502
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
            }
            # -- API: group parent groups (nested membership) ----------
            elseif ($path -match "^/api/groups/([^/]+)/parents$" -and $req.HttpMethod -eq "GET") {
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
                    $parents = @(Get-GroupParentGroups -GroupId $groupId)
                    $json    = ConvertTo-SafeJson -InputObject $parents -AsArray
                    $buffer  = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 200
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
                catch {
                    Write-Warning "Parent group fetch failed for $groupId : $($_.Exception.Message)"
                    $errBody = ConvertTo-Json -InputObject @{ error = "Failed to fetch parent groups." } -Compress
                    $buffer  = [System.Text.Encoding]::UTF8.GetBytes($errBody)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 502
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
            }
            # -- API: nested group assignments -------------------------
            elseif ($path -match "^/api/groups/([^/]+)/nested-assignments$" -and $req.HttpMethod -eq "GET") {
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
                    $parents = @(Get-GroupParentGroups -GroupId $groupId)
                    if ($parents.Count -gt 0) {
                        $nestedResult = Get-NestedGroupAssignments -GroupId $groupId -ParentGroups $parents
                    } else {
                        $nestedResult = @{
                            configurations  = @()
                            settingsCatalog = @()
                            applications    = @()
                            scripts         = @()
                            remediations    = @()
                            _errors         = @{}
                        }
                    }
                    $json   = ConvertTo-SafeJson -InputObject $nestedResult
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 200
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
                catch {
                    Write-Warning "Nested assignment fetch failed for $groupId : $($_.Exception.Message)"
                    $errBody = ConvertTo-Json -InputObject @{ error = "Failed to fetch nested assignments." } -Compress
                    $buffer  = [System.Text.Encoding]::UTF8.GetBytes($errBody)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 502
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
            }
            # -- API: orphaned items -----------------------------------
            elseif ($path -eq "/api/orphaned-items" -and $req.HttpMethod -eq "GET") {
                try {
                    $orphanedResult = Get-OrphanedItems
                    $json   = ConvertTo-SafeJson -InputObject $orphanedResult
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 200
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
                catch {
                    Write-Warning "Orphaned items fetch failed: $($_.Exception.Message)"
                    $errBody = ConvertTo-Json -InputObject @{ error = "Failed to fetch orphaned items." } -Compress
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
                    $errBody = ConvertTo-Json -InputObject @{ error = "Failed to fetch assignments. Please try again." } -Compress
                    $buffer  = [System.Text.Encoding]::UTF8.GetBytes($errBody)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 502
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
            }
            # -- API: assignments by target type -----------------------
            elseif ($path -eq "/api/assignments-by-target" -and $req.HttpMethod -eq "GET") {
                $targetType = $req.QueryString["type"]
                $validTypes = @(
                    "#microsoft.graph.allDevicesAssignmentTarget",
                    "#microsoft.graph.allLicensedUsersAssignmentTarget"
                )
                if (-not $targetType -or $targetType -notin $validTypes) {
                    $resp.StatusCode = 400
                    $body   = '{"error":"Invalid or missing target type parameter."}'
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($body)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                    continue
                }
                try {
                    $assignmentsResult = Get-AssignmentsByTargetType -TargetOdataType $targetType
                    $json   = ConvertTo-Json -InputObject $assignmentsResult -Depth 10 -Compress
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 200
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
                catch {
                    Write-Warning "Assignment fetch by target type failed: $($_.Exception.Message)"
                    $errBody = ConvertTo-Json -InputObject @{ error = "Failed to fetch assignments. Please try again." } -Compress
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
                    $errBody = ConvertTo-Json -InputObject @{ error = "Failed to fetch script content. Please try again." } -Compress
                    $buffer  = [System.Text.Encoding]::UTF8.GetBytes($errBody)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.StatusCode      = 502
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                }
            }
            # -- API: logout -----------------------------------------
            elseif ($path -eq "/api/logout" -and $req.HttpMethod -eq "POST") {
                # CSRF protection: validate Origin header, fall back to Referer
                $origin         = $req.Headers["Origin"]
                $referer        = $req.Headers["Referer"]
                $expectedOrigin = "http://localhost:$Port"
                $csrfValid      = $false

                if ($origin) {
                    # Origin header present — must match exactly
                    $csrfValid = ($origin -eq $expectedOrigin)
                } elseif ($referer) {
                    # No Origin header — fall back to Referer prefix check
                    $csrfValid = $referer.StartsWith("$expectedOrigin/")
                }
                # If neither header is present, reject the request

                if (-not $csrfValid) {
                    $resp.StatusCode = 403
                    $body   = '{"error":"Forbidden: invalid origin."}'
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($body)
                    $resp.ContentType     = "application/json; charset=utf-8"
                    $resp.ContentLength64 = $buffer.Length
                    $resp.OutputStream.Write($buffer, 0, $buffer.Length)
                    $resp.OutputStream.Close()
                    continue
                }
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
