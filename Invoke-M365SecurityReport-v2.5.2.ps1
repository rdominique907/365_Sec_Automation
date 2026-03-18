<#
.SYNOPSIS
    Consolidated M365 Security Report with weekly→monthly rollup, deduplication, MFA status tracking, and AI executive summary.

.DESCRIPTION
    Single script replacing Measure-M365Security + New-SecurityTemplate.
    Creates or updates monthly Excel workbooks with deduplication.
    
    Features:
    - Auto-create workbook if missing, append if exists
    - Deterministic deduplication (no duplicates on re-run)
    - Weekly default (last 7 days) with monthly accumulation
    - 14 sheets: 11 data domains + MFA-Status + AI-Summary + Config + Runs
    - Optional JSON export per domain
    - AI executive summary via Azure OpenAI (gpt-5-chat) with monthly create-or-update lifecycle
    - Self-test mode for validation

    Data Sources (11 Domains):
    1. IAM-Entra (Sign-ins, Risky Users)
    2. Hybrid-ADConnect (Security-relevant Directory Audits)
    3. M365-Exchange (Admin operations)
    4. M365-SharePoint-Teams (Sharing events)
    5. Endpoint-MDE (Security recommendations)
    6. Data-Purview (DLP events)
    7. Alerts-Response (Security Alerts)
    8. App-Consents (OAuth, Service Principals)
    9. Privileged-Access (Role assignments, PIM activations)
    10. Conditional-Access (CA policy changes)
    11. AI-Summary (Computed)

    Required Permissions:
    - Graph: AuditLog.Read.All, IdentityRiskyUser.Read.All, SecurityAlert.Read.All,
             Application.Read.All, Directory.Read.All, RoleManagement.Read.Directory
    - O365 Management: ActivityFeed.Read, ActivityFeed.ReadDlp
    - WindowsDefenderATP: SecurityRecommendation.Read.All
    - Azure: Reader role on subscriptions

.PARAMETER TenantId
    M365 Tenant ID (required)

.PARAMETER ClientId
    App Registration Client ID (required)

.PARAMETER ClientSecret
    Client Secret. Can also use $env:M365_CLIENT_SECRET

.PARAMETER Month
    Target month in YYYY-MM format. Default: current month

.PARAMETER StartDate
    Override start date. Default: 7 days ago

.PARAMETER EndDate
    Override end date. Default: now

.PARAMETER SkipAzureResources
    Skip Azure Resource Graph queries

.PARAMETER SkipDefenderEndpoint
    Skip Defender for Endpoint queries

.PARAMETER DefenderRegion
    Override MDE regional endpoint

.PARAMETER ThrottleLimit
    Parallel download threads (default: 10)

.PARAMETER OpenAIKey
    Optional OpenAI API key for AI summary (public api.openai.com)

.PARAMETER OpenAIModel
    OpenAI model (default: gpt-5.2)

.PARAMETER AzureOpenAIEndpoint
    Azure OpenAI endpoint URL (e.g., https://myresource.openai.azure.com)
    Use this instead of OpenAIKey for enterprise Azure OpenAI deployments.

.PARAMETER AzureOpenAIDeployment
    Azure OpenAI model deployment name (e.g., gpt-5-chat)

.PARAMETER AzureOpenAIApiVersion
    Azure OpenAI API version (default: 2024-10-21)

.PARAMETER NoAI
    Skip AI artifacts even if key provided

.PARAMETER RedactForAI
    Strip PII from AI context

.PARAMETER ExecutiveSummaryPath
    Custom path for executive summary Markdown file

.PARAMETER SkipExecutiveSummary
    Skip executive summary generation even if AI keys are provided

.PARAMETER ReportPath
    Custom report path

.PARAMETER ExportJson
    Export JSON files per domain

.PARAMETER SharePointSiteUrl
    SharePoint site URL to upload report artifacts (e.g., https://tenant.sharepoint.com/sites/SecurityOps)

.PARAMETER SharePointFolder
    Folder path in SharePoint (default: "Shared Documents/Security-Reports")

.PARAMETER WhatIf
    Show planned actions without writing

.PARAMETER SelfTest
    Run built-in validation tests

.EXAMPLE
    ./Invoke-M365SecurityReport.ps1 -TenantId "xxx" -ClientId "xxx" -ClientSecret "secret"

.EXAMPLE
    # Weekly scheduled job
    ./Invoke-M365SecurityReport.ps1 -TenantId "xxx" -ClientId "xxx" -ThrottleLimit 20

.NOTES
    Version: 2.5.2
    Author: Rolando Dominique
    Requires: PowerShell 7, ImportExcel module
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $true)][string]$TenantId,
    [Parameter(Mandatory = $true)][string]$ClientId,
    [Parameter(Mandatory = $false)][string]$ClientSecret,
    
    [string]$Month,
    [DateTime]$StartDate,
    [DateTime]$EndDate,
    
    [switch]$SkipAzureResources,
    [switch]$SkipDefenderEndpoint,
    [string]$DefenderRegion,
    [int]$ThrottleLimit = 10,
    
    [string]$OpenAIKey,
    [string]$OpenAIModel = "gpt-5.2",
    [switch]$NoAI,
    [switch]$RedactForAI,
    
    # Azure OpenAI (v2.4.0)
    [string]$AzureOpenAIEndpoint,
    [string]$AzureOpenAIDeployment,
    [string]$AzureOpenAIApiVersion = "2024-10-21",
    
    # Executive Summary (v2.4.0)
    [string]$ExecutiveSummaryPath,
    [switch]$SkipExecutiveSummary,
    
    [string]$ReportPath,
    [switch]$ExportJson,
    [switch]$SelfTest,
    
    # SharePoint storage (v2.2.0)
    [string]$SharePointSiteUrl,
    [string]$SharePointFolder = "Shared Documents/Security-Reports"
)

$ErrorActionPreference = "Stop"
$script:Version = "2.5.2"
$script:SchemaVersion = 1
$script:UseSharePoint = $false
$script:SharePointSiteId = $null
$script:LocalTempPath = $null

# ==================================================================================
# CONSTANTS
# ==================================================================================
$SheetNames = @{
    Identity    = "IAM-Entra"
    Hybrid      = "Hybrid-ADConnect"
    Exchange    = "M365-Exchange"
    SharePoint  = "M365-SharePoint-Teams"
    Endpoint    = "Endpoint-MDE"
    DataProtect = "Data-Purview"
    Alerts      = "Alerts-Response"
    AppConsents = "App-Consents"
    PrivAccess  = "Privileged-Access"
    CondAccess  = "Conditional-Access"
    MFAStatus   = "MFA-Status"
    Summary     = "Summary"
    Config      = "Config"
    Runs        = "Runs"
}

$SheetHeaders = @{
    Identity    = @("FindingId", "Time", "Source", "Operation", "Actor", "Target", "Result", "RiskLevel(1-10)", "Category", "DedupKey", "RawJson")
    Hybrid      = @("FindingId", "Time", "Source", "Operation", "Actor", "Target", "Result", "RiskLevel(1-10)", "Details", "DedupKey", "RawJson")
    Exchange    = @("FindingId", "Time", "Workload", "Operation", "User", "Item", "ClientApp", "IP", "Result", "RiskLevel(1-10)", "DedupKey", "RawJson")
    SharePoint  = @("FindingId", "Time", "Workload", "Operation", "User", "SiteOrTeam", "TargetItem", "SharingType", "Result", "RiskLevel(1-10)", "DedupKey")
    Endpoint    = @("FindingId", "Time", "Product", "RecommendationId", "RecommendationName", "SeverityScore", "CVE", "CVSS", "ExposedMachines", "RemediationType", "Status", "DedupKey")
    DataProtect = @("FindingId", "Time", "Workload", "Operation", "User", "Policy", "Target", "Action", "Result", "Severity", "DedupKey")
    Alerts      = @("FindingId", "Time", "Provider", "AlertId", "Title", "Severity", "Category", "Entity", "Status", "DedupKey", "RawJson")
    AppConsents = @("FindingId", "Time", "Operation", "AppName", "AppId", "Permissions", "InitiatedBy", "ConsentType", "Result", "RiskLevel(1-10)", "DedupKey", "RawJson")
    PrivAccess  = @("FindingId", "Time", "Operation", "Actor", "RoleOrScope", "Target", "Justification", "Result", "RiskLevel(1-10)", "DedupKey", "RawJson")
    CondAccess  = @("FindingId", "Time", "Operation", "PolicyName", "PolicyId", "ModifiedBy", "ChangeDetails", "Result", "RiskLevel(1-10)", "DedupKey", "RawJson")
    MFAStatus   = @("UserPrincipalName", "DisplayName", "IsMfaRegistered", "IsMfaCapable", "DefaultMfaMethod", "MethodsRegistered", "AccountEnabled", "UserType", "DedupKey")
    Runs        = @("RunId", "StartTime", "EndTime", "DateRangeStart", "DateRangeEnd", "Version", "TotalInserted", "TotalSkipped", "Status", "Domains")
}

$script:ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }

# ==================================================================================
# INITIALIZATION
# ==================================================================================
$scriptStart = Get-Date

# Resolve client secret from env if not provided
if ([string]::IsNullOrEmpty($ClientSecret)) {
    $ClientSecret = $env:M365_CLIENT_SECRET
    if ([string]::IsNullOrEmpty($ClientSecret)) {
        throw "ClientSecret required. Provide via parameter or M365_CLIENT_SECRET environment variable."
    }
}
$ClientSecretPlain = $ClientSecret

# Resolve dates
if (-not $EndDate) { $EndDate = Get-Date }
if (-not $StartDate) { $StartDate = $EndDate.AddDays(-7) }

# Resolve target month
$targetMonth = if ($Month) { $Month } else { $EndDate.ToString("yyyy-MM") }

# Finding ID infrastructure (v2.5.1)
$script:ReportMonth = $targetMonth.Replace("-", "")
$script:FindingCounter = @{
    ID = 0; HY = 0; EX = 0; SP = 0; EP = 0; DP = 0; AL = 0; AC = 0; PA = 0; CA = 0
}

# Resolve report path
if (-not $ReportPath) {
    $ReportPath = Join-Path $script:ScriptRoot "Security-Monthly-Report-$targetMonth.xlsx"
}

# Setup logging - use monthly naming convention for SharePoint upload
$script:LogFileName = "Security-Log-$targetMonth.txt"
$LogPath = Join-Path $script:ScriptRoot $script:LogFileName
try { Start-Transcript -Path $LogPath -Append | Out-Null } catch { }

# ==================================================================================
# HELPER FUNCTIONS
# ==================================================================================
function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')][string]$Level = 'INFO')
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Message"
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN' { Write-Host $line -ForegroundColor Yellow }
        'DEBUG' { Write-Host $line -ForegroundColor DarkGray }
        default { Write-Host $line -ForegroundColor Gray }
    }
}

function Get-HashString {
    param([string]$InputString)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($InputString)
    $hash = [System.Security.Cryptography.SHA256]::HashData($bytes)
    return [System.BitConverter]::ToString($hash).Replace("-", "").ToLower().Substring(0, 16)
}

# Finding ID generator (v2.5.1)
function Get-FindingId {
    param([string]$Prefix)
    $script:FindingCounter[$Prefix]++
    return "$Prefix-$($script:ReportMonth)-$($script:FindingCounter[$Prefix].ToString('D4'))"
}

function Invoke-RetryableRequest {
    param(
        [string]$Uri,
        [hashtable]$Headers,
        [string]$Method = "Get",
        [object]$Body,
        [int]$MaxRetries = 3,
        [int]$TimeoutSec = 120
    )
    
    $retryCount = 0
    $baseDelay = 5
    
    while ($retryCount -lt $MaxRetries) {
        try {
            $params = @{ Uri = $Uri; Headers = $Headers; Method = $Method; TimeoutSec = $TimeoutSec }
            if ($Body) {
                $params.Body = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 }
                $params.ContentType = "application/json"
            }
            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $_.Exception.Response.StatusCode.value__
            if ($statusCode -eq 429 -or $statusCode -eq 503) {
                $delay = $baseDelay * [math]::Pow(2, $retryCount)
                Write-Log "Throttled ($statusCode). Waiting ${delay}s..." -Level WARN
                Start-Sleep -Seconds $delay
                $retryCount++
            }
            elseif ($statusCode -in @(401, 403)) {
                Write-Log "Auth failed for $Uri" -Level ERROR
                throw $_
            }
            else { throw $_ }
        }
    }
    throw "Max retries exceeded for $Uri"
}

# ==================================================================================
# AUTHENTICATION
# ==================================================================================
$script:TokenAcquiredTime = $null
$script:TokenMaxAgeMinutes = 45
$script:GraphToken = $null
$script:O365Token = $null
$script:AzureToken = $null
$script:DefenderToken = $null

function Get-GraphToken {
    $body = @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecretPlain
        grant_type    = "client_credentials"
    }
    $r = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -Body $body
    return $r.access_token
}

function Get-O365Token {
    $body = @{
        client_id     = $ClientId
        resource      = "https://manage.office.com"
        client_secret = $ClientSecretPlain
        grant_type    = "client_credentials"
    }
    $r = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/token" -Method Post -Body $body
    return $r.access_token
}

function Get-AzureToken {
    $body = @{
        client_id     = $ClientId
        scope         = "https://management.azure.com/.default"
        client_secret = $ClientSecretPlain
        grant_type    = "client_credentials"
    }
    $r = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -Body $body
    return $r.access_token
}

function Get-DefenderToken {
    $body = @{
        client_id     = $ClientId
        scope         = "https://api.securitycenter.microsoft.com/.default"
        client_secret = $ClientSecretPlain
        grant_type    = "client_credentials"
    }
    $r = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -Body $body
    return $r.access_token
}

function Refresh-TokensIfNeeded {
    $age = if ($script:TokenAcquiredTime) { ((Get-Date) - $script:TokenAcquiredTime).TotalMinutes } else { 999 }
    if ($age -gt $script:TokenMaxAgeMinutes) {
        Write-Log "Refreshing tokens (age: $([math]::Round($age))m)" -Level DEBUG
        $script:GraphToken = Get-GraphToken
        $script:O365Token = Get-O365Token
        try { $script:DefenderToken = Get-DefenderToken } catch { Write-Log "Defender token refresh skipped: $_" -Level DEBUG }
        $script:TokenAcquiredTime = Get-Date
    }
}

# Continue in next part...

# ==================================================================================
# SHAREPOINT STORAGE (v2.2.0)
# ==================================================================================
function Get-SharePointSiteId {
    param([string]$SiteUrl, [string]$Token)
    
    # Parse site URL to extract hostname and site path
    # Example: https://contoso.sharepoint.com/sites/SecurityOps
    if ($SiteUrl -match "https://([^/]+)/sites/([^/]+)") {
        $hostname = $Matches[1]
        $siteName = $Matches[2]
        $uri = "https://graph.microsoft.com/v1.0/sites/${hostname}:/sites/${siteName}"
    }
    elseif ($SiteUrl -match "https://([^/]+)/?$") {
        # Root site
        $hostname = $Matches[1]
        $uri = "https://graph.microsoft.com/v1.0/sites/${hostname}"
    }
    else {
        throw "Invalid SharePoint site URL format: $SiteUrl"
    }
    
    try {
        $headers = @{ Authorization = "Bearer $Token" }
        $site = Invoke-RestMethod -Uri $uri -Headers $headers -ErrorAction Stop
        return $site.id
    }
    catch {
        Write-Log "Failed to resolve SharePoint site ID: $_" -Level ERROR
        throw $_
    }
}

function Get-SharePointFile {
    param(
        [string]$SiteId,
        [string]$FolderPath,
        [string]$FileName,
        [string]$LocalPath,
        [string]$Token
    )
    
    $headers = @{ Authorization = "Bearer $Token" }
    $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$FolderPath/${FileName}"
    
    try {
        # First check if file exists
        $fileInfo = Invoke-RestMethod -Uri $uri -Headers $headers -ErrorAction Stop
        
        # Download the file content
        $downloadUri = "$uri`:/content"
        Invoke-RestMethod -Uri $downloadUri -Headers $headers -OutFile $LocalPath -ErrorAction Stop
        
        Write-Log "Downloaded report from SharePoint: $FileName" -Level INFO
        return $true
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 404) {
            Write-Log "Report not found in SharePoint - will create new" -Level DEBUG
            return $false
        }
        throw $_
    }
}

function Set-SharePointFile {
    param(
        [string]$SiteId,
        [string]$FolderPath,
        [string]$FileName,
        [string]$LocalPath,
        [string]$Token
    )
    
    $headers = @{ 
        Authorization  = "Bearer $Token"
        "Content-Type" = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
    
    # Ensure folder exists by creating it (idempotent)
    try {
        $folderUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$FolderPath"
        $folderCheck = Invoke-RestMethod -Uri $folderUri -Headers @{ Authorization = "Bearer $Token" } -ErrorAction SilentlyContinue
    }
    catch {
        # Folder doesn't exist - create it
        try {
            $parts = $FolderPath -split "/"
            $currentPath = ""
            foreach ($part in $parts) {
                if ([string]::IsNullOrEmpty($part)) { continue }
                $parentPath = if ($currentPath) { $currentPath } else { "root" }
                $createUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/$parentPath/children"
                $body = @{ name = $part; folder = @{}; "@microsoft.graph.conflictBehavior" = "fail" } | ConvertTo-Json
                try {
                    Invoke-RestMethod -Uri $createUri -Headers @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json" } -Method Post -Body $body -ErrorAction SilentlyContinue | Out-Null
                }
                catch { }
                $currentPath = if ($currentPath) { "root:/${currentPath}/${part}:" } else { "root:/${part}:" }
            }
        }
        catch { Write-Log "Note: Folder creation may have partially succeeded" -Level DEBUG }
    }
    
    # Upload file - try simple PUT first, then upload session for locked files
    $uploadUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/${FolderPath}/${FileName}:/content"
    
    try {
        $fileBytes = [System.IO.File]::ReadAllBytes($LocalPath)
        Invoke-RestMethod -Uri $uploadUri -Headers $headers -Method Put -Body $fileBytes -ErrorAction Stop | Out-Null
        Write-Log "Uploaded report to SharePoint: $FolderPath/$FileName" -Level INFO
        return $true
    }
    catch {
        $errBody = $_.ErrorDetails.Message
        if ($errBody -match "resourceLocked" -or $errBody -match "locked") {
            # Co-authoring lock detected — use createUploadSession which can bypass it
            Write-Log "File locked, attempting upload session bypass..." -Level DEBUG
            try {
                $sessionUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/${FolderPath}/${FileName}:/createUploadSession"
                $sessionBody = @{
                    item = @{
                        "@microsoft.graph.conflictBehavior" = "replace"
                    }
                } | ConvertTo-Json -Depth 5
                $session = Invoke-RestMethod -Uri $sessionUri -Headers @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json" } -Method Post -Body $sessionBody -ErrorAction Stop
                
                # Upload the file content to the session URL
                $fileBytes = [System.IO.File]::ReadAllBytes($LocalPath)
                $fileSize = $fileBytes.Length
                $uploadHeaders = @{
                    "Content-Length" = $fileSize
                    "Content-Range"  = "bytes 0-$($fileSize - 1)/$fileSize"
                }
                Invoke-RestMethod -Uri $session.uploadUrl -Headers $uploadHeaders -Method Put -Body $fileBytes -ErrorAction Stop | Out-Null
                Write-Log "Uploaded via session (bypassed lock): $FolderPath/$FileName" -Level INFO
                return $true
            }
            catch {
                Write-Log "Upload session also failed: $_" -Level ERROR
                throw $_
            }
        }
        else {
            Write-Log "Failed to upload to SharePoint: $_" -Level ERROR
            throw $_
        }
    }
}

function Get-SharePointItemId {
    param(
        [string]$SiteId,
        [string]$FolderPath,
        [string]$FileName,
        [string]$Token
    )
    
    $headers = @{ Authorization = "Bearer $Token" }
    $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$FolderPath/${FileName}"
    
    try {
        $item = Invoke-RestMethod -Uri $uri -Headers $headers -ErrorAction Stop
        return $item.id
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 404) {
            return $null  # File doesn't exist yet
        }
        throw $_
    }
}

function Unlock-SharePointFile {
    param(
        [string]$SiteId,
        [string]$ItemId,
        [string]$Token
    )
    
    $headers = @{ 
        Authorization  = "Bearer $Token"
        "Content-Type" = "application/json"
    }
    
    # Try force check-in first (works for checked-out files)
    $checkinUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/items/$ItemId/checkin"
    $checkinBody = @{ comment = "Forced check-in by M365 Security Report script" } | ConvertTo-Json
    
    try {
        Invoke-RestMethod -Uri $checkinUri -Headers $headers -Method Post -Body $checkinBody -ErrorAction Stop | Out-Null
        Write-Log "Forced check-in successful for item $ItemId" -Level INFO
        return $true
    }
    catch {
        $checkinError = $_.ErrorDetails.Message
        Write-Log "Check-in failed (co-authoring lock, not checkout): $checkinError" -Level DEBUG
        # Co-authoring locks cannot be released via API — the upload session approach in Set-SharePointFile handles this
        return $false
    }
}

# ==================================================================================
# GRAPH HELPERS
# ==================================================================================
function Invoke-GraphWithPagination {
    param([string]$Uri, [string]$Token, [int]$MaxPages = 100, [string]$Label = "Data")
    $headers = @{ Authorization = "Bearer $Token" }
    $all = @()
    $page = 0
    do {
        $page++
        Write-Host "`r    Fetching $Label... page $page ($($all.Count) items)" -NoNewline -ForegroundColor DarkGray
        $r = Invoke-RetryableRequest -Uri $Uri -Headers $headers
        if ($r.value) { $all += $r.value } elseif ($r -is [array]) { $all += $r }
        $Uri = $r.'@odata.nextLink'
    } while ($Uri -and $page -lt $MaxPages)
    Write-Host "`r    Fetched ${Label}: $($all.Count) items                    " -ForegroundColor DarkGray
    return $all
}

$script:DefenderRegions = @(
    "api.securitycenter.microsoft.com", "api-eu.securitycenter.microsoft.com",
    "api-uk.securitycenter.microsoft.com", "api-us.securitycenter.microsoft.com"
)

function Find-DefenderEndpoint {
    param([string]$Token)
    $headers = @{ Authorization = "Bearer $Token" }
    foreach ($region in $script:DefenderRegions) {
        try {
            $null = Invoke-RestMethod -Uri "https://$region/api/machines?`$top=1" -Headers $headers -TimeoutSec 10 -ErrorAction Stop
            return $region
        }
        catch {
            $sc = $_.Exception.Response.StatusCode.value__
            if ($sc -in @(401, 403)) { return $region }
        }
    }
    return $null
}

# ==================================================================================
# DEDUPLICATION ENGINE
# ==================================================================================
function Get-DedupKey {
    param([string]$Domain, [object]$Item)
    
    try {
        switch ($Domain) {
            "Identity" {
                if ($Item.riskLevel) {
                    # Risky user
                    return Get-HashString "$($Item.id)|$($Item.riskLastUpdatedDateTime)|$($Item.riskLevel)"
                }
                else {
                    # Sign-in - handle null status safely
                    $errorCode = if ($Item.status) { $Item.status.errorCode } else { "" }
                    return Get-HashString "$($Item.userId)|$($Item.createdDateTime)|$($Item.ipAddress)|$($Item.appId)|$errorCode"
                }
            }
            "Hybrid" {
                if ($Item.id) { return Get-HashString $Item.id }
                # Safe access to initiatedBy
                $init = ""
                if ($Item.initiatedBy) {
                    if ($Item.initiatedBy.user) { $init = $Item.initiatedBy.user.id }
                    elseif ($Item.initiatedBy.app) { $init = $Item.initiatedBy.app.id }
                }
                return Get-HashString "$($Item.activityDateTime)|$($Item.activityDisplayName)|$init"
            }
            "Exchange" {
                if ($Item.Id) { return Get-HashString $Item.Id }
                return Get-HashString "$($Item.CreationTime)|$($Item.Operation)|$($Item.UserId)|$($Item.ClientIP)|$($Item.ObjectId)"
            }
            "SharePoint" {
                if ($Item.Id) { return Get-HashString $Item.Id }
                return Get-HashString "$($Item.CreationTime)|$($Item.Operation)|$($Item.UserId)|$($Item.SiteUrl)|$($Item.ObjectId)"
            }
            "Endpoint" {
                $id = if ($Item.id) { $Item.id } elseif ($Item.recommendationId) { $Item.recommendationId } else { "$($Item.productName)|$($Item.recommendationName)" }
                return Get-HashString $id
            }
            "DataProtect" {
                if ($Item.Id) { return Get-HashString $Item.Id }
                return Get-HashString "$($Item.CreationTime)|$($Item.Operation)|$($Item.UserId)|$($Item.ObjectId)"
            }
            "Alerts" {
                return Get-HashString "$($Item.id)"
            }
            "AppConsents" {
                if ($Item.id) { return Get-HashString $Item.id }
                $target = ""
                if ($Item.targetResources -and $Item.targetResources.Count -gt 0) { 
                    $target = $Item.targetResources[0].id 
                }
                $init = ""
                if ($Item.initiatedBy -and $Item.initiatedBy.user) { 
                    $init = $Item.initiatedBy.user.id 
                }
                return Get-HashString "$($Item.activityDateTime)|$target|$init"
            }
            "PrivAccess" {
                if ($Item.id) { return Get-HashString $Item.id }
                $init = ""
                if ($Item.initiatedBy) {
                    if ($Item.initiatedBy.user) { $init = $Item.initiatedBy.user.id }
                    elseif ($Item.initiatedBy.app) { $init = $Item.initiatedBy.app.id }
                }
                return Get-HashString "$($Item.activityDateTime)|$($Item.activityDisplayName)|$init"
            }
            "CondAccess" {
                if ($Item.id) { return Get-HashString $Item.id }
                $target = ""
                if ($Item.targetResources -and $Item.targetResources.Count -gt 0) {
                    $target = $Item.targetResources[0].id
                }
                return Get-HashString "$($Item.activityDateTime)|$($Item.activityDisplayName)|$target"
            }
            "MFAStatus" {
                return Get-HashString "$($Item.userPrincipalName)|$($Item.isMfaRegistered)"
            }
            default { return Get-HashString ($Item | ConvertTo-Json -Compress -Depth 10 -WarningAction SilentlyContinue) }
        }
    }
    catch {
        # Fallback: generate key from JSON representation
        return Get-HashString ($Item | ConvertTo-Json -Compress -Depth 10 -WarningAction SilentlyContinue)
    }
}

function Build-ExistingKeyIndex {
    param([string]$Path, [string]$SheetName)
    $keys = [System.Collections.Generic.HashSet[string]]::new()
    if (-not (Test-Path $Path)) { return $keys }
    
    try {
        $data = Import-Excel -Path $Path -WorksheetName $SheetName -ErrorAction SilentlyContinue
        # Handle empty sheet (Import-Excel returns null or empty when only headers exist)
        if ($null -eq $data -or $data.Count -eq 0) { 
            return $keys 
        }
        # Handle single row (not wrapped in array)
        if ($data -isnot [array]) { $data = @($data) }
        foreach ($row in $data) {
            if ($null -ne $row -and $row.DedupKey) { 
                [void]$keys.Add($row.DedupKey) 
            }
        }
    }
    catch {
        # Sheet doesn't exist or other error - return empty keys (will insert all)
        Write-Log "Note: Could not read existing keys from $SheetName - will insert all new rows" -Level DEBUG
    }
    return $keys
}

# ==================================================================================
# EXCEL WORKBOOK MANAGEMENT
# ==================================================================================
function New-ReportWorkbook {
    param([string]$Path)
    
    Write-Log "Creating new workbook: $Path" -Level INFO
    
    # Create Config sheet first
    $initData = [PSCustomObject]@{ Key = "Initialized"; Value = (Get-Date).ToString("o") }
    $initData | Export-Excel -Path $Path -WorksheetName $SheetNames.Config -AutoSize
    
    # Create each data sheet by exporting an empty header row
    foreach ($domain in @("Identity", "Hybrid", "Exchange", "SharePoint", "Endpoint", "DataProtect", "Alerts", "AppConsents", "PrivAccess", "CondAccess", "MFAStatus")) {
        $headers = $SheetHeaders[$domain]
        # Create a PSCustomObject with all headers as properties
        $headerRow = [ordered]@{}
        foreach ($h in $headers) { $headerRow[$h] = $null }
        [PSCustomObject]$headerRow | Export-Excel -Path $Path -WorksheetName $SheetNames[$domain] -AutoSize
        
        # Remove the null data row, keeping only the header
        $excel = Open-ExcelPackage -Path $Path
        $sheet = $excel.Workbook.Worksheets[$SheetNames[$domain]]
        if ($sheet.Dimension.Rows -gt 1) {
            $sheet.DeleteRow(2)
        }
        # Bold the header row
        $sheet.Row(1).Style.Font.Bold = $true
        Close-ExcelPackage $excel
    }
    
    # Create Summary sheet (will be populated at end by Update-SummarySheet)
    $summaryInit = [PSCustomObject]@{ A = "Summary will be generated after data collection"; B = "" }
    $summaryInit | Export-Excel -Path $Path -WorksheetName $SheetNames.Summary -NoHeader
    
    # Create Runs sheet
    $runsHeaders = $SheetHeaders.Runs
    $runsHeaderRow = [ordered]@{}
    foreach ($h in $runsHeaders) { $runsHeaderRow[$h] = $null }
    [PSCustomObject]$runsHeaderRow | Export-Excel -Path $Path -WorksheetName $SheetNames.Runs -AutoSize
    
    # Clean up Runs sheet (remove null row)
    $excel = Open-ExcelPackage -Path $Path
    $runsSheet = $excel.Workbook.Worksheets[$SheetNames.Runs]
    if ($runsSheet.Dimension.Rows -gt 1) {
        $runsSheet.DeleteRow(2)
    }
    $runsSheet.Row(1).Style.Font.Bold = $true
    Close-ExcelPackage $excel
    
    Write-Log "Workbook created with 14 sheets" -Level INFO
}

function Update-ConfigSheet {
    param([string]$Path, [hashtable]$Stats, [string]$RunId)
    
    $excel = Open-ExcelPackage -Path $Path
    $cfg = $excel.Workbook.Worksheets[$SheetNames.Config]
    
    # Clear and rewrite config
    $cfg.Cells.Clear()
    $row = 1
    $configData = @{
        "TenantId"      = $TenantId.Substring(0, [Math]::Min(8, $TenantId.Length)) + "..."
        "ScriptVersion" = $script:Version
        "SchemaVersion" = $script:SchemaVersion
        "LastRunTime"   = (Get-Date).ToString("o")
        "LastRunRange"  = "$($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))"
        "LastRunId"     = $RunId
        "TotalInserted" = ($Stats.Values | ForEach-Object { $_.Inserted } | Measure-Object -Sum).Sum
        "TotalSkipped"  = ($Stats.Values | ForEach-Object { $_.Skipped } | Measure-Object -Sum).Sum
    }
    
    foreach ($key in $configData.Keys) {
        $cfg.Cells[$row, 1].Value = $key
        $cfg.Cells[$row, 2].Value = $configData[$key]
        $row++
    }
    
    # Per-domain stats
    foreach ($domain in $Stats.Keys) {
        $cfg.Cells[$row, 1].Value = "${domain}_Inserted"
        $cfg.Cells[$row, 2].Value = $Stats[$domain].Inserted
        $row++
        $cfg.Cells[$row, 1].Value = "${domain}_Skipped"
        $cfg.Cells[$row, 2].Value = $Stats[$domain].Skipped
        $row++
    }
    
    Close-ExcelPackage $excel
}

function Add-RunRecord {
    param([string]$Path, [string]$RunId, [DateTime]$Start, [DateTime]$RangeStart, [DateTime]$RangeEnd, [int]$Inserted, [int]$Skipped, [string]$Status, [string]$Domains)
    
    $runRow = [PSCustomObject]@{
        RunId          = $RunId
        StartTime      = $Start.ToString("o")
        EndTime        = (Get-Date).ToString("o")
        DateRangeStart = $RangeStart.ToString("yyyy-MM-dd")
        DateRangeEnd   = $RangeEnd.ToString("yyyy-MM-dd")
        Version        = $script:Version
        TotalInserted  = $Inserted
        TotalSkipped   = $Skipped
        Status         = $Status
        Domains        = $Domains
    }
    $runRow | Export-Excel -Path $Path -WorksheetName $SheetNames.Runs -Append
}

function Update-SummarySheet {
    param(
        [string]$Path,
        [hashtable]$Stats,
        [DateTime]$RangeStart,
        [DateTime]$RangeEnd,
        [string]$RunId
    )
    
    Write-Host "`n[Summary] Generating executive summary..." -ForegroundColor Cyan
    
    # Domain config for reading data
    $domainConfig = @(
        @{ Key = "Identity"; Name = "Identity & Access"; Sheet = $SheetNames.Identity; RiskCol = "RiskLevel(1-10)" }
        @{ Key = "Hybrid"; Name = "Hybrid Identity"; Sheet = $SheetNames.Hybrid; RiskCol = "RiskLevel(1-10)" }
        @{ Key = "Exchange"; Name = "Exchange Online"; Sheet = $SheetNames.Exchange; RiskCol = "RiskLevel(1-10)" }
        @{ Key = "SharePoint"; Name = "SharePoint/Teams"; Sheet = $SheetNames.SharePoint; RiskCol = "RiskLevel(1-10)" }
        @{ Key = "Endpoint"; Name = "Endpoint (MDE)"; Sheet = $SheetNames.Endpoint; RiskCol = "SeverityScore" }
        @{ Key = "DataProtect"; Name = "Data Protection"; Sheet = $SheetNames.DataProtect; RiskCol = "Severity" }
        @{ Key = "Alerts"; Name = "Security Alerts"; Sheet = $SheetNames.Alerts; RiskCol = "Severity" }
        @{ Key = "AppConsents"; Name = "App Consents"; Sheet = $SheetNames.AppConsents; RiskCol = "RiskLevel(1-10)" }
        @{ Key = "PrivAccess"; Name = "Privileged Access"; Sheet = $SheetNames.PrivAccess; RiskCol = "RiskLevel(1-10)" }
        @{ Key = "CondAccess"; Name = "Conditional Access"; Sheet = $SheetNames.CondAccess; RiskCol = "RiskLevel(1-10)" }
    )
    
    $totalInserted = ($Stats.Values | ForEach-Object { $_.Inserted } | Measure-Object -Sum).Sum
    $totalSkipped = ($Stats.Values | ForEach-Object { $_.Skipped } | Measure-Object -Sum).Sum
    
    # Build summary data as objects
    $summaryRows = @()
    
    # Section 1: Header
    $summaryRows += [PSCustomObject]@{ A = "SECURITY REPORT SUMMARY"; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
    $summaryRows += [PSCustomObject]@{ A = ""; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
    $summaryRows += [PSCustomObject]@{ A = "Report Period:"; B = "$($RangeStart.ToString('MMM d')) - $($RangeEnd.ToString('MMM d, yyyy'))"; C = ""; D = ""; E = ""; F = ""; G = "" }
    $summaryRows += [PSCustomObject]@{ A = "Generated:"; B = (Get-Date).ToString("MMM d, yyyy h:mm tt"); C = ""; D = ""; E = ""; F = ""; G = "" }
    $summaryRows += [PSCustomObject]@{ A = "Run ID:"; B = $RunId; C = ""; D = ""; E = ""; F = ""; G = "" }
    $summaryRows += [PSCustomObject]@{ A = "Total New Findings:"; B = $totalInserted; C = ""; D = ""; E = ""; F = ""; G = "" }
    $summaryRows += [PSCustomObject]@{ A = "Deduplicated:"; B = $totalSkipped; C = ""; D = ""; E = ""; F = ""; G = "" }
    $summaryRows += [PSCustomObject]@{ A = ""; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
    
    # Section 2: Domain table header
    $summaryRows += [PSCustomObject]@{ A = "FINDINGS BY DOMAIN"; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
    $summaryRows += [PSCustomObject]@{ A = "Domain"; B = "Events"; C = "Critical"; D = "High"; E = "Medium"; F = "Low"; G = "Sheet" }
    
    # Collect all high-risk findings for top 10
    $allFindings = @()
    
    foreach ($dc in $domainConfig) {
        $sheetData = @()
        try {
            $sheetData = Import-Excel -Path $Path -WorksheetName $dc.Sheet -ErrorAction SilentlyContinue
            if ($null -eq $sheetData) { $sheetData = @() }
            if ($sheetData -isnot [array]) { $sheetData = @($sheetData) }
        }
        catch { }
        
        $eventCount = $sheetData.Count
        $critical = 0; $high = 0; $medium = 0; $low = 0
        
        foreach ($item in $sheetData) {
            $risk = 0
            try {
                $riskProp = $dc.RiskCol
                if ($item.PSObject.Properties[$riskProp]) { $risk = [int]($item.$riskProp) }
            }
            catch { }
            if ($risk -ge 9) { $critical++; $allFindings += @{ Domain = $dc.Name; Item = $item; Risk = $risk } }
            elseif ($risk -ge 7) { $high++; $allFindings += @{ Domain = $dc.Name; Item = $item; Risk = $risk } }
            elseif ($risk -ge 4) { $medium++ }
            else { $low++ }
        }
        
        $summaryRows += [PSCustomObject]@{ A = $dc.Name; B = $eventCount; C = $critical; D = $high; E = $medium; F = $low; G = $dc.Sheet }
    }
    
    $summaryRows += [PSCustomObject]@{ A = ""; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
    
    # Section 3: Top 10
    $summaryRows += [PSCustomObject]@{ A = "TOP 10 CRITICAL/HIGH FINDINGS"; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
    $summaryRows += [PSCustomObject]@{ A = "#"; B = "Domain"; C = "Severity"; D = "Time"; E = "Description"; F = "Actor"; G = "" }
    
    $top10 = $allFindings | Sort-Object { $_.Risk } -Descending | Select-Object -First 10
    $rank = 1
    foreach ($f in $top10) {
        $item = $f.Item
        $time = if ($item.Time) { $item.Time } elseif ($item.CreationTime) { $item.CreationTime } else { "" }
        $desc = if ($item.Operation) { $item.Operation } elseif ($item.Title) { $item.Title } else { "" }
        $actor = if ($item.Actor) { $item.Actor } elseif ($item.User) { $item.User } elseif ($item.Account) { $item.Account } else { "" }
        $summaryRows += [PSCustomObject]@{ A = $rank; B = $f.Domain; C = $f.Risk; D = $time; E = $desc; F = $actor; G = "" }
        $rank++
    }
    if ($top10.Count -eq 0) {
        $summaryRows += [PSCustomObject]@{ A = "No critical or high severity findings"; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
    }
    
    $summaryRows += [PSCustomObject]@{ A = ""; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
    
    # Section 4: Actions
    $summaryRows += [PSCustomObject]@{ A = "RECOMMENDED ACTIONS"; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
    
    # Count specific findings for recommendations
    $exData = @(); try { $exData = Import-Excel -Path $Path -WorksheetName $SheetNames.Exchange -ErrorAction SilentlyContinue } catch { }
    $inboxRules = @($exData | Where-Object { $_.Operation -match "InboxRule" })
    if ($inboxRules.Count -gt 0) {
        $summaryRows += [PSCustomObject]@{ A = "Review $($inboxRules.Count) inbox rule change(s) - BEC attack vector"; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
    }
    
    $spData = @(); try { $spData = Import-Excel -Path $Path -WorksheetName $SheetNames.SharePoint -ErrorAction SilentlyContinue } catch { }
    $anonLinks = @($spData | Where-Object { $_.Operation -match "Anonymous" })
    $extUsers = @($spData | Where-Object { $_.Operation -match "External" })
    if ($anonLinks.Count -gt 0) { $summaryRows += [PSCustomObject]@{ A = "Review $($anonLinks.Count) anonymous link(s) created"; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" } }
    if ($extUsers.Count -gt 0) { $summaryRows += [PSCustomObject]@{ A = "Review $($extUsers.Count) external user(s) added"; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" } }
    
    $idData = @(); try { $idData = Import-Excel -Path $Path -WorksheetName $SheetNames.Identity -ErrorAction SilentlyContinue } catch { }
    $highRisk = @($idData | Where-Object { $_.'RiskLevel(1-10)' -ge 7 })
    if ($highRisk.Count -gt 0) { $summaryRows += [PSCustomObject]@{ A = "Investigate $($highRisk.Count) high-risk sign-in failure(s)"; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" } }
    
    if ($inboxRules.Count -eq 0 -and $anonLinks.Count -eq 0 -and $extUsers.Count -eq 0 -and $highRisk.Count -eq 0) {
        $summaryRows += [PSCustomObject]@{ A = "No critical action items identified"; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
    }
    
    $summaryRows += [PSCustomObject]@{ A = ""; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
    
    # Section 5: MFA Registration Status
    $summaryRows += [PSCustomObject]@{ A = "MFA REGISTRATION STATUS"; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
    
    # Read MFA stats from script variable or from sheet
    $mfaStatsAvailable = $false
    $mfaTotal = 0; $mfaRegistered = 0; $mfaNotRegistered = 0
    
    if ($null -ne $script:MfaStats -and $script:MfaStats.TotalUsers -gt 0) {
        $mfaTotal = $script:MfaStats.TotalUsers
        $mfaRegistered = $script:MfaStats.Registered
        $mfaNotRegistered = $script:MfaStats.NotRegistered
        $mfaStatsAvailable = $true
    }
    else {
        # Try to read from MFA-Status sheet
        try {
            $mfaData = Import-Excel -Path $Path -WorksheetName $SheetNames.MFAStatus -ErrorAction SilentlyContinue
            if ($null -ne $mfaData) {
                if ($mfaData -isnot [array]) { $mfaData = @($mfaData) }
                $mfaNotRegistered = $mfaData.Count
                $mfaStatsAvailable = $true
            }
        }
        catch { }
    }
    
    if ($mfaStatsAvailable) {
        $pct = if ($mfaTotal -gt 0) { [math]::Round([double](($mfaRegistered / $mfaTotal) * 100), 1) } else { 0 }
        $summaryRows += [PSCustomObject]@{ A = "Total Users Scanned:"; B = $mfaTotal; C = ""; D = ""; E = ""; F = ""; G = "" }
        $summaryRows += [PSCustomObject]@{ A = "MFA Registered:"; B = $mfaRegistered; C = "($pct%)"; D = ""; E = ""; F = ""; G = "" }
        $summaryRows += [PSCustomObject]@{ A = "NOT Registered for MFA:"; B = $mfaNotRegistered; C = ""; D = "(See MFA-Status sheet)"; E = ""; F = ""; G = "" }
        if ($mfaNotRegistered -gt 0) {
            $summaryRows += [PSCustomObject]@{ A = "[!] ACTION: Review $mfaNotRegistered account(s) without MFA registration"; B = ""; C = ""; D = ""; E = ""; F = ""; G = "" }
        }
    }
    else {
        $summaryRows += [PSCustomObject]@{ A = "MFA data not available"; B = "(Requires UserAuthenticationMethod.Read.All permission)"; C = ""; D = ""; E = ""; F = ""; G = "" }
    }
    
    # Delete old Summary sheet if exists and write new
    try {
        $excel = Open-ExcelPackage -Path $Path
        $oldSummary = $excel.Workbook.Worksheets[$SheetNames.Summary]
        if ($oldSummary) { $excel.Workbook.Worksheets.Delete($SheetNames.Summary) }
        # Also remove old AI-Summary if present
        $oldAI = $excel.Workbook.Worksheets["AI-Summary"]
        if ($oldAI) { $excel.Workbook.Worksheets.Delete("AI-Summary") }
        Close-ExcelPackage $excel
    }
    catch { }
    
    # Write using Export-Excel (more reliable)
    $summaryRows | Export-Excel -Path $Path -WorksheetName $SheetNames.Summary -NoHeader -AutoSize
    
    # Move Summary to first position and hide Config/Runs tabs
    try {
        $excel = Open-ExcelPackage -Path $Path
        $excel.Workbook.Worksheets.MoveToStart($SheetNames.Summary)
        # Hide administrative tabs from human view (v2.5.1)
        $configSheet = $excel.Workbook.Worksheets[$SheetNames.Config]
        if ($null -ne $configSheet) { $configSheet.Hidden = [OfficeOpenXml.eWorkSheetHidden]::Hidden }
        $runsSheet = $excel.Workbook.Worksheets[$SheetNames.Runs]
        if ($null -ne $runsSheet) { $runsSheet.Hidden = [OfficeOpenXml.eWorkSheetHidden]::Hidden }
        Close-ExcelPackage $excel
    }
    catch { }
    
    # ==================================================================================
    # CHART GENERATION (v2.3.1 - Uses hidden data sheets for reliable rendering)
    # ==================================================================================
    Write-Host "  Generating charts..." -ForegroundColor DarkGray
    
    $mfaChartCreated = $false
    $domainChartCreated = $false
    $driftChartCreated = $false
    
    try {
        # --------------------------------------------------------------------------
        # Step 1: Create hidden data sheets with chart data
        # EPPlus 4.x renders charts correctly when referencing data from separate sheets
        # --------------------------------------------------------------------------
        
        $excel = Open-ExcelPackage -Path $Path
        
        # Remove old chart data sheets if they exist
        foreach ($oldSheet in @("ChartData_MFA", "ChartData_Domain", "ChartData_Drift")) {
            $existing = $excel.Workbook.Worksheets[$oldSheet]
            if ($null -ne $existing) {
                $excel.Workbook.Worksheets.Delete($oldSheet)
            }
        }
        
        # ----- MFA Chart Data Sheet -----
        if ($mfaStatsAvailable -and ($mfaRegistered -gt 0 -or $mfaNotRegistered -gt 0)) {
            $mfaDataSheet = $excel.Workbook.Worksheets.Add("ChartData_MFA")
            $mfaDataSheet.Cells["A1"].Value = "Status"
            $mfaDataSheet.Cells["B1"].Value = "Users"
            $mfaDataSheet.Cells["A2"].Value = "MFA Registered"
            $mfaDataSheet.Cells["B2"].Value = [int]$mfaRegistered
            $mfaDataSheet.Cells["A3"].Value = "Not Registered"
            $mfaDataSheet.Cells["B3"].Value = [int]$mfaNotRegistered
            $mfaDataSheet.Hidden = [OfficeOpenXml.eWorkSheetHidden]::Hidden
            Write-Host "    Created ChartData_MFA (Registered: $mfaRegistered, Not Registered: $mfaNotRegistered)" -ForegroundColor DarkGray
        }
        
        # ----- Domain Chart Data Sheet -----
        # Collect domain stats from the Summary sheet or from the $Stats hashtable
        $domainDataSheet = $excel.Workbook.Worksheets.Add("ChartData_Domain")
        $domainDataSheet.Cells["A1"].Value = "Domain"
        $domainDataSheet.Cells["B1"].Value = "Events"
        
        $domainConfig = @(
            @{ Key = "Identity"; Name = "Identity & Access" }
            @{ Key = "Hybrid"; Name = "Hybrid Identity" }
            @{ Key = "Exchange"; Name = "Exchange Online" }
            @{ Key = "SharePoint"; Name = "SharePoint/Teams" }
            @{ Key = "Endpoint"; Name = "Endpoint (MDE)" }
            @{ Key = "DataProtect"; Name = "Data Protection" }
            @{ Key = "Alerts"; Name = "Security Alerts" }
            @{ Key = "AppConsents"; Name = "App Consents" }
            @{ Key = "PrivAccess"; Name = "Privileged Access" }
            @{ Key = "CondAccess"; Name = "Conditional Access" }
        )
        
        $domainRow = 2
        foreach ($dc in $domainConfig) {
            $eventCount = 0
            if ($Stats.ContainsKey($dc.Key)) {
                $eventCount = $Stats[$dc.Key].Inserted + $Stats[$dc.Key].Skipped
            }
            $domainDataSheet.Cells["A$domainRow"].Value = $dc.Name
            $domainDataSheet.Cells["B$domainRow"].Value = [int]$eventCount
            $domainRow++
        }
        $domainLastRow = $domainRow - 1
        $domainDataSheet.Hidden = [OfficeOpenXml.eWorkSheetHidden]::Hidden
        Write-Host "    Created ChartData_Domain ($($domainLastRow - 1) domains)" -ForegroundColor DarkGray
        
        # ----- Drift Chart Data Sheet -----
        $runsWs = $excel.Workbook.Worksheets[$SheetNames.Runs]
        $driftLastRow = 1  # Just header
        
        if ($null -ne $runsWs -and $runsWs.Dimension -and $runsWs.Dimension.Rows -ge 2) {
            $driftDataSheet = $excel.Workbook.Worksheets.Add("ChartData_Drift")
            $driftDataSheet.Cells["A1"].Value = "Date"
            $driftDataSheet.Cells["B1"].Value = "Findings"
            
            $driftRow = 2
            for ($r = 2; $r -le $runsWs.Dimension.Rows; $r++) {
                $dateVal = $runsWs.Cells[$r, 4].Text    # Column D = DateRangeStart
                $findingsVal = $runsWs.Cells[$r, 7].Value  # Column G = TotalInserted
                if (-not [string]::IsNullOrEmpty($dateVal)) {
                    $driftDataSheet.Cells["A$driftRow"].Value = $dateVal
                    $driftDataSheet.Cells["B$driftRow"].Value = [int]$findingsVal
                    $driftRow++
                }
            }
            $driftLastRow = $driftRow - 1
            $driftDataSheet.Hidden = [OfficeOpenXml.eWorkSheetHidden]::Hidden
            Write-Host "    Created ChartData_Drift ($($driftLastRow - 1) data points)" -ForegroundColor DarkGray
        }
        
        Close-ExcelPackage $excel
        
        # --------------------------------------------------------------------------
        # Step 2: Create charts on Summary sheet referencing the data sheets
        # --------------------------------------------------------------------------
        
        $excel = Open-ExcelPackage -Path $Path
        $summaryWs = $excel.Workbook.Worksheets[$SheetNames.Summary]
        
        # Remove any existing charts on Summary
        $chartNames = @($summaryWs.Drawings | ForEach-Object { $_.Name })
        foreach ($name in $chartNames) {
            $summaryWs.Drawings.Remove($name)
        }
        
        # ----- Chart 1: MFA Pie Chart -----
        $mfaDataSheet = $excel.Workbook.Worksheets["ChartData_MFA"]
        if ($null -ne $mfaDataSheet) {
            $mfaChart = $summaryWs.Drawings.AddChart("MFA_Compliance", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
            $mfaChart.Title.Text = "MFA Registration Status"
            $mfaChart.Title.Font.Size = 12
            $mfaChart.Title.Font.Bold = $true
            $mfaChart.SetPosition(0, 5, 7, 5)
            $mfaChart.SetSize(300, 220)
            
            $mfaSeries = $mfaChart.Series.Add(
                $mfaDataSheet.Cells["B2:B3"],
                $mfaDataSheet.Cells["A2:A3"]
            )
            $mfaSeries.Header = "Users"
            $mfaChart.Legend.Position = [OfficeOpenXml.Drawing.Chart.eLegendPosition]::Bottom
            
            $mfaChartCreated = $true
            Write-Host "    + MFA Pie chart created" -ForegroundColor Green
        }
        
        # ----- Chart 2: Domain Bar Chart -----
        $domainDataSheet = $excel.Workbook.Worksheets["ChartData_Domain"]
        if ($null -ne $domainDataSheet -and $domainLastRow -gt 1) {
            $domainChart = $summaryWs.Drawings.AddChart("Domain_Findings", [OfficeOpenXml.Drawing.Chart.eChartType]::BarClustered)
            $domainChart.Title.Text = "Findings by Domain"
            $domainChart.Title.Font.Size = 11
            $domainChart.Title.Font.Bold = $true
            $domainChart.SetPosition(12, 0, 7, 5)
            $domainChart.SetSize(340, 280)
            
            $domainSeries = $domainChart.Series.Add(
                $domainDataSheet.Cells["B2:B$domainLastRow"],
                $domainDataSheet.Cells["A2:A$domainLastRow"]
            )
            $domainSeries.Header = "Events"
            
            try { $domainChart.Legend.Remove() } catch { }
            
            $domainChartCreated = $true
            Write-Host "    + Domain Bar chart created" -ForegroundColor Green
        }
        
        # ----- Chart 3: Drift Column Chart -----
        $driftDataSheet = $excel.Workbook.Worksheets["ChartData_Drift"]
        if ($null -ne $driftDataSheet -and $driftLastRow -gt 1) {
            $driftChart = $summaryWs.Drawings.AddChart("Security_Drift", [OfficeOpenXml.Drawing.Chart.eChartType]::ColumnClustered)
            $driftChart.Title.Text = "New Findings Trend"
            $driftChart.Title.Font.Size = 11
            $driftChart.Title.Font.Bold = $true
            $driftChart.SetPosition(27, 0, 7, 5)
            $driftChart.SetSize(340, 180)
            
            $driftSeries = $driftChart.Series.Add(
                $driftDataSheet.Cells["B2:B$driftLastRow"],
                $driftDataSheet.Cells["A2:A$driftLastRow"]
            )
            $driftSeries.Header = "Findings"
            
            try { $driftChart.Legend.Remove() } catch { }
            
            $driftChartCreated = $true
            Write-Host "    + Drift Column chart created" -ForegroundColor Green
        }
        
        Close-ExcelPackage $excel
        
        # Summary of chart creation
        $chartCount = @($mfaChartCreated, $domainChartCreated, $driftChartCreated) | Where-Object { $_ } | Measure-Object | Select-Object -ExpandProperty Count
        Write-Host "  Charts created: $chartCount of 3 (MFA: $(if($mfaChartCreated){'OK'}else{'--'}), Domain: $(if($domainChartCreated){'OK'}else{'--'}), Drift: $(if($driftChartCreated){'OK'}else{'--'}))" -ForegroundColor $(if ($chartCount -eq 3) { 'Green' } else { 'Yellow' })
    }
    catch {
        Write-Log "Chart generation failed: $_" -Level WARN
        Write-Host "  ! Chart generation error: $_" -ForegroundColor Yellow
        Write-Host "    (Charts require EPPlus with charting support)" -ForegroundColor DarkGray
    }
    
    Write-Host "  Summary sheet updated with executive overview" -ForegroundColor Green
    
    # ==================================================================================
    # STRUCTURED DATA TABLES (v2.5.1 - Named Excel Tables for Power Automate)
    # ==================================================================================
    Write-Host "  Creating structured data tables..." -ForegroundColor DarkGray
    
    # Store collected data in script scope for AI Context reuse
    $script:SummaryDomainRows = @()
    $script:SummaryFindingsRows = @()
    $script:SummaryActionsList = @()
    $script:SummaryMFAData = @{
        Available     = $mfaStatsAvailable
        Total         = $mfaTotal
        Registered    = $mfaRegistered
        NotRegistered = $mfaNotRegistered
    }
    
    try {
        # PRE-READ all domain sheet data
        $domainReadData = @{}
        foreach ($dc in $domainConfig) {
            $sheetData = @()
            try {
                $sheetData = Import-Excel -Path $Path -WorksheetName $dc.Sheet -ErrorAction SilentlyContinue
                if ($null -eq $sheetData) { $sheetData = @() }
                if ($sheetData -isnot [array]) { $sheetData = @($sheetData) }
            }
            catch { }
            $domainReadData[$dc.Key] = @{ SheetData = $sheetData; RiskCol = $dc.RiskCol; Name = $dc.Name }
        }
        
        # ----- Table 1: Domain Statistics (build as PSCustomObject array, write with Export-Excel) -----
        foreach ($dc in $domainConfig) {
            $drd = $domainReadData[$dc.Key]
            $sheetData = $drd.SheetData
            
            $crit = 0; $hi = 0; $med = 0; $lo = 0
            foreach ($item in $sheetData) {
                $risk = 0
                try { if ($item.PSObject.Properties[$dc.RiskCol]) { $risk = [int]($item.($dc.RiskCol)) } } catch { }
                if ($risk -ge 9) { $crit++ } elseif ($risk -ge 7) { $hi++ } elseif ($risk -ge 4) { $med++ } else { $lo++ }
            }
            
            $ins = if ($Stats.ContainsKey($dc.Key)) { $Stats[$dc.Key].Inserted } else { 0 }
            $skp = if ($Stats.ContainsKey($dc.Key)) { $Stats[$dc.Key].Skipped } else { 0 }
            
            $script:SummaryDomainRows += [PSCustomObject]@{
                Domain = $dc.Name; TotalEvents = [int]$sheetData.Count; NewThisRun = [int]$ins; Deduplicated = [int]$skp
                Critical = [int]$crit; High = [int]$hi; Medium = [int]$med; Low = [int]$lo
            }
        }
        $script:SummaryDomainRows | Export-Excel -Path $Path -WorksheetName "SummaryData_Domain" -TableName "DomainStats" -TableStyle Medium2 -AutoSize
        Write-Host "    + DomainStats table ($($script:SummaryDomainRows.Count) rows)" -ForegroundColor Green
        
        # ----- Table 2: Top Findings -----
        $rank = 1
        foreach ($f in $top10) {
            $item = $f.Item
            $time = if ($item.Time) { "$($item.Time)" } elseif ($item.CreationTime) { "$($item.CreationTime)" } else { "" }
            $desc = if ($item.Operation) { "$($item.Operation)" } elseif ($item.Title) { "$($item.Title)" } else { "" }
            $actor = if ($item.Actor) { "$($item.Actor)" } elseif ($item.User) { "$($item.User)" } elseif ($item.Account) { "$($item.Account)" } else { "" }
            
            $script:SummaryFindingsRows += [PSCustomObject]@{
                Rank = $rank; Domain = $f.Domain; Severity = [int]$f.Risk
                Time = $time; Description = $desc; Actor = $actor
            }
            $rank++
        }
        if ($script:SummaryFindingsRows.Count -eq 0) {
            $script:SummaryFindingsRows += [PSCustomObject]@{
                Rank = 0; Domain = "None"; Severity = 0
                Time = ""; Description = "No critical or high findings"; Actor = ""
            }
        }
        $script:SummaryFindingsRows | Export-Excel -Path $Path -WorksheetName "SummaryData_Findings" -TableName "TopFindings" -TableStyle Medium2 -AutoSize
        Write-Host "    + TopFindings table ($($script:SummaryFindingsRows.Count) rows)" -ForegroundColor Green
        
        # ----- Table 3: Recommended Actions -----
        $priority = 1
        if ($inboxRules.Count -gt 0) {
            $script:SummaryActionsList += [PSCustomObject]@{ Priority = $priority; Category = "Exchange"; Action = "Review inbox rule changes - BEC attack vector"; Count = [int]$inboxRules.Count }
            $priority++
        }
        if ($anonLinks.Count -gt 0) {
            $script:SummaryActionsList += [PSCustomObject]@{ Priority = $priority; Category = "SharePoint"; Action = "Review anonymous sharing links created"; Count = [int]$anonLinks.Count }
            $priority++
        }
        if ($extUsers.Count -gt 0) {
            $script:SummaryActionsList += [PSCustomObject]@{ Priority = $priority; Category = "SharePoint"; Action = "Review external users added to organization"; Count = [int]$extUsers.Count }
            $priority++
        }
        if ($highRisk.Count -gt 0) {
            $script:SummaryActionsList += [PSCustomObject]@{ Priority = $priority; Category = "Identity"; Action = "Investigate high-risk sign-in failures"; Count = [int]$highRisk.Count }
            $priority++
        }
        if ($mfaNotRegistered -gt 0) {
            $script:SummaryActionsList += [PSCustomObject]@{ Priority = $priority; Category = "MFA"; Action = "Review accounts without MFA registration"; Count = [int]$mfaNotRegistered }
            $priority++
        }
        if ($script:SummaryActionsList.Count -eq 0) {
            $script:SummaryActionsList += [PSCustomObject]@{ Priority = 0; Category = "None"; Action = "No critical action items identified"; Count = 0 }
        }
        $script:SummaryActionsList | Export-Excel -Path $Path -WorksheetName "SummaryData_Actions" -TableName "RecommendedActions" -TableStyle Medium2 -AutoSize
        Write-Host "    + RecommendedActions table ($($script:SummaryActionsList.Count) rows)" -ForegroundColor Green
        
        # ----- Table 4: MFA Status -----
        $mfaRows = @()
        if ($mfaStatsAvailable -and $mfaTotal -gt 0) {
            $mfRegPct = [math]::Round([double](($mfaRegistered / $mfaTotal) * 100), 1)
            $mfNotPct = [math]::Round([double](($mfaNotRegistered / $mfaTotal) * 100), 1)
            $mfaRows += [PSCustomObject]@{ Metric = "Total Users Scanned"; Value = [int]$mfaTotal; Percentage = "100%" }
            $mfaRows += [PSCustomObject]@{ Metric = "MFA Registered"; Value = [int]$mfaRegistered; Percentage = "$mfRegPct%" }
            $mfaRows += [PSCustomObject]@{ Metric = "NOT Registered"; Value = [int]$mfaNotRegistered; Percentage = "$mfNotPct%" }
        }
        else {
            $mfaRows += [PSCustomObject]@{ Metric = "MFA Data"; Value = 0; Percentage = "Not Available" }
        }
        $mfaRows | Export-Excel -Path $Path -WorksheetName "SummaryData_MFA" -TableName "MFAStatus" -TableStyle Medium2 -AutoSize
        Write-Host "    + MFAStatus table created" -ForegroundColor Green
        
        # Hide data table sheets
        try {
            $excel = Open-ExcelPackage -Path $Path
            foreach ($dtName in @("SummaryData_Domain", "SummaryData_Findings", "SummaryData_Actions", "SummaryData_MFA")) {
                $dtSheet = $excel.Workbook.Worksheets[$dtName]
                if ($null -ne $dtSheet) { $dtSheet.Hidden = [OfficeOpenXml.eWorkSheetHidden]::Hidden }
            }
            Close-ExcelPackage $excel
        }
        catch { Write-Log "Could not hide data table sheets: $_" -Level DEBUG }
        
        Write-Host "  Structured data tables ready for Power Automate" -ForegroundColor Green
    }
    catch {
        Write-Log "Data table creation failed: $_" -Level WARN
        Write-Host "  ! Data table creation error (non-fatal): $_" -ForegroundColor Yellow
    }
}

# ==================================================================================
# MONTH-OVER-MONTH COMPARISON (v2.5.1)
# ==================================================================================
function Get-MonthOverMonthDeltas {
    param(
        [string]$Path,
        [int]$CurrentInserted,
        [int]$CurrentSkipped,
        [int]$CurrentCritical,
        [int]$CurrentHigh,
        [int]$CurrentMFAGap,
        [int]$CurrentExternalSharing
    )
    
    $result = @{
        HasPreviousMonth = $false
        PrevInserted = 0; CurrInserted = $CurrentInserted
        PrevSkipped = 0; CurrSkipped = $CurrentSkipped
        PrevCritical = 0; CurrCritical = $CurrentCritical
        PrevHigh = 0; CurrHigh = $CurrentHigh
        PrevMFAGap = 0; CurrMFAGap = $CurrentMFAGap
        PrevExternalSharing = 0; CurrExternalSharing = $CurrentExternalSharing
        DeltaInserted = 0; DeltaPct = "N/A"
        Trend = "Unknown"
    }
    
    try {
        $runsData = Import-Excel -Path $Path -WorksheetName $SheetNames.Runs -ErrorAction SilentlyContinue
        if ($null -eq $runsData) { return $result }
        if ($runsData -isnot [array]) { $runsData = @($runsData) }
        
        # Find the most recent completed run (not the current one being recorded)
        if ($runsData.Count -ge 1) {
            $prevRun = $runsData | Select-Object -Last 1
            $result.HasPreviousMonth = $true
            $result.PrevInserted = [int]$prevRun.TotalInserted
            $result.PrevSkipped = [int]$prevRun.TotalSkipped
            $result.DeltaInserted = $CurrentInserted - $result.PrevInserted
            
            if ($result.PrevInserted -gt 0) {
                $pct = [math]::Round([double](($result.DeltaInserted / $result.PrevInserted) * 100), 0)
                $result.DeltaPct = if ($pct -ge 0) { "+$pct%" } else { "$pct%" }
            }
            
            # Determine trend based on new findings volume
            if ($result.DeltaInserted -gt ($result.PrevInserted * 0.2)) { $result.Trend = "Worsening" }
            elseif ($result.DeltaInserted -lt - ($result.PrevInserted * 0.1)) { $result.Trend = "Improving" }
            else { $result.Trend = "Stable" }
        }
    }
    catch { }
    
    return $result
}

# ==================================================================================
# SECURITY POSTURE SCORE (v2.5.1)
# ==================================================================================
function Get-SecurityPostureScore {
    param(
        [hashtable]$MFAData,
        [array]$TopFindings,
        [array]$ActionsList,
        [hashtable]$MoMDeltas,
        [int]$InboxRuleCount,
        [int]$SendAsCount,
        [int]$AnonLinkCount,
        [int]$ExtUserCount
    )
    
    # Factor 1: MFA Coverage (25%)
    $mfaScore = 100
    if ($MFAData.Available -and $MFAData.Total -gt 0) {
        $mfaPct = ($MFAData.Registered / $MFAData.Total) * 100
        $mfaScore = [math]::Min(100, [math]::Round($mfaPct))
    }
    
    # Factor 2: Critical/High Findings (30%)
    $critHighCount = 0
    if ($TopFindings) {
        $critHighCount = ($TopFindings | Where-Object { $_.Severity -ge 7 }).Count
    }
    $findingsScore = switch ($critHighCount) {
        0 { 100; break }
        { $_ -le 2 } { 80; break }
        { $_ -le 5 } { 60; break }
        { $_ -le 10 } { 40; break }
        default { 20 }
    }
    
    # Factor 3: BEC Indicators (10%) - inbox rules + SendAs
    $becCount = $InboxRuleCount + $SendAsCount
    $becScore = switch ($becCount) {
        0 { 100; break }
        { $_ -le 2 } { 70; break }
        { $_ -le 5 } { 40; break }
        default { 10 }
    }
    
    # Factor 4: External Exposure (10%) - anon links + external users
    $extCount = $AnonLinkCount + $ExtUserCount
    $extScore = switch ($extCount) {
        0 { 100; break }
        { $_ -le 3 } { 80; break }
        { $_ -le 10 } { 50; break }
        default { 20 }
    }
    
    # Factor 5: Month-over-Month Trend (15%)
    $trendScore = 70  # Default neutral
    if ($MoMDeltas.HasPreviousMonth) {
        $trendScore = switch ($MoMDeltas.Trend) {
            "Improving" { 90 }
            "Stable" { 70 }
            "Worsening" { 40 }
            default { 70 }
        }
    }
    
    # Factor 6: Unresolved Actions (10%)
    $actionCount = if ($ActionsList) { $ActionsList.Count } else { 0 }
    $actionsScore = switch ($actionCount) {
        0 { 100; break }
        { $_ -le 2 } { 75; break }
        { $_ -le 5 } { 50; break }
        default { 25 }
    }
    
    # Weighted total
    $factors = @(
        @{ Name = "MFA Coverage"; Score = $mfaScore; Weight = 0.25; Detail = "$(if ($MFAData.Available) { "$([math]::Round([double](($MFAData.Registered / [math]::Max($MFAData.Total, 1)) * 100)))%" } else { 'N/A' })"; Contribution = 0 }
        @{ Name = "Critical/High Findings"; Score = $findingsScore; Weight = 0.30; Detail = "$critHighCount findings"; Contribution = 0 }
        @{ Name = "BEC Indicators"; Score = $becScore; Weight = 0.10; Detail = "$becCount events"; Contribution = 0 }
        @{ Name = "External Exposure"; Score = $extScore; Weight = 0.10; Detail = "$extCount events"; Contribution = 0 }
        @{ Name = "Month-over-Month Trend"; Score = $trendScore; Weight = 0.15; Detail = $MoMDeltas.Trend; Contribution = 0 }
        @{ Name = "Unresolved Actions"; Score = $actionsScore; Weight = 0.10; Detail = "$actionCount items"; Contribution = 0 }
    )
    
    $totalScore = 0
    foreach ($f in $factors) {
        $f.Contribution = [math]::Round([double]($f.Score * $f.Weight), 1)
        $totalScore += $f.Contribution
    }
    $totalScore = [math]::Round([double]$totalScore)
    
    # Rating bands
    $rating = switch ($totalScore) {
        { $_ -ge 85 } { "Good" }
        { $_ -ge 70 } { "Moderate" }
        { $_ -ge 50 } { "Elevated" }
        default { "Critical" }
    }
    $emoji = switch ($rating) {
        "Good" { "✅" }
        "Moderate" { "⚠️" }
        "Elevated" { "🟠" }
        "Critical" { "🔴" }
    }
    
    return @{
        Score = $totalScore; Rating = $rating; Emoji = $emoji; Factors = $factors
    }
}

# ==================================================================================
# DOMAIN RISK NARRATIVES (v2.5.1)
# ==================================================================================
function Get-DomainNarrative {
    param(
        [string]$DomainName,
        [array]$Events,
        [int]$Critical,
        [int]$High,
        [int]$Medium,
        [int]$Low,
        [string]$TopActor
    )
    
    $total = $Critical + $High + $Medium + $Low
    
    # Determine risk level
    $riskLevel = if ($Critical -gt 0) { "CRITICAL" }
    elseif ($High -gt 0) { "ELEVATED" }
    elseif ($Medium -gt 0 -and $total -gt 20) { "MODERATE" }
    else { "NORMAL" }
    
    $riskEmoji = switch ($riskLevel) {
        "CRITICAL" { "🔴" }
        "ELEVATED" { "⚠️" }
        "MODERATE" { "🟡" }
        "NORMAL" { "✅" }
    }
    
    # Generate narrative based on domain and event types
    $narrativeParts = @()
    
    switch ($DomainName) {
        "Exchange" {
            $inboxOps = @($Events | Where-Object { $_.Operation -match "InboxRule" }).Count
            $sendAsOps = @($Events | Where-Object { $_.Operation -match "SendAs" }).Count
            $mailAccess = @($Events | Where-Object { $_.Operation -match "MailItemsAccessed" }).Count
            $permOps = @($Events | Where-Object { $_.Operation -match "MailboxPermission" }).Count
            
            if ($inboxOps -gt 0) { $narrativeParts += "$inboxOps inbox rule change(s) detected — classic BEC attack vector" }
            if ($sendAsOps -gt 0) { $narrativeParts += "$sendAsOps SendAs impersonation event(s) — verify authorization" }
            if ($mailAccess -gt 0) { $narrativeParts += "$mailAccess email access event(s) logged (E5 Audit Premium)" }
            if ($permOps -gt 0) { $narrativeParts += "$permOps mailbox permission change(s) — review delegation grants" }
        }
        "SharePoint" {
            $anonCreated = @($Events | Where-Object { $_.Operation -match "AnonymousLinkCreated" }).Count
            $anonUsed = @($Events | Where-Object { $_.Operation -match "AnonymousLinkUsed" }).Count
            $versionsDeleted = @($Events | Where-Object { $_.Operation -match "FileVersionsAllDeleted" }).Count
            $extUsers = @($Events | Where-Object { $_.Operation -match "ExternalUserAdded" }).Count
            $adminAdded = @($Events | Where-Object { $_.Operation -match "SiteCollectionAdminAdded" }).Count
            $fileDeleted = @($Events | Where-Object { $_.Operation -match "FileDeleted" }).Count
            
            if ($versionsDeleted -gt 0) { $narrativeParts += "⚠ $versionsDeleted file version history wipe(s) — potential ransomware or evidence destruction" }
            if ($anonUsed -gt 0) { $narrativeParts += "$anonUsed anonymous link access event(s) — external parties accessed data" }
            if ($anonCreated -gt 0) { $narrativeParts += "$anonCreated anonymous sharing link(s) created — review data exposure" }
            if ($extUsers -gt 0) { $narrativeParts += "$extUsers external user(s) added to tenant" }
            if ($adminAdded -gt 0) { $narrativeParts += "$adminAdded site collection admin elevation(s)" }
            if ($fileDeleted -gt 0) { $narrativeParts += "$fileDeleted file deletion(s) flagged" }
        }
        "Identity" {
            $riskyUsers = @($Events | Where-Object { $_.Category -eq "RiskyUser" }).Count
            $authFailures = @($Events | Where-Object { $_.Category -eq "AuthFailure" }).Count
            
            if ($riskyUsers -gt 0) { $narrativeParts += "$riskyUsers unresolved risky user(s) flagged by Identity Protection" }
            if ($authFailures -gt 0) { $narrativeParts += "$authFailures failed sign-in attempt(s) — review for credential attacks" }
        }
        "Endpoint" {
            $highSev = @($Events | Where-Object { [int]$_.SeverityScore -ge 7 }).Count
            $exposedTotal = ($Events | Measure-Object -Property ExposedMachines -Sum).Sum
            
            if ($highSev -gt 0) { $narrativeParts += "$highSev high-severity vulnerability recommendation(s)" }
            if ($exposedTotal -gt 0) { $narrativeParts += "$exposedTotal machine(s) with known exposure" }
        }
        "Alerts" {
            $highAlerts = @($Events | Where-Object { [int]$_.Severity -ge 7 }).Count
            $activeAlerts = @($Events | Where-Object { $_.Status -eq "new" -or $_.Status -eq "inProgress" }).Count
            
            if ($highAlerts -gt 0) { $narrativeParts += "$highAlerts high-severity security alert(s)" }
            if ($activeAlerts -gt 0) { $narrativeParts += "$activeAlerts unresolved active alert(s)" }
        }
        "AppConsents" {
            $credOps = @($Events | Where-Object { $_.Operation -match "credentials" }).Count
            $consentOps = @($Events | Where-Object { $_.Operation -match "Consent|delegated|app role" }).Count
            $newApps = @($Events | Where-Object { $_.Operation -match "Add application" }).Count
            $ownerOps = @($Events | Where-Object { $_.Operation -match "owner" }).Count
            
            if ($credOps -gt 0) { $narrativeParts += "⚠ $credOps credential addition/removal(s) — potential persistence mechanism" }
            if ($consentOps -gt 0) { $narrativeParts += "$consentOps consent grant(s) — review for illicit consent attacks" }
            if ($newApps -gt 0) { $narrativeParts += "$newApps new application registration(s)" }
            if ($ownerOps -gt 0) { $narrativeParts += "$ownerOps ownership change(s) on apps/service principals" }
        }
        "PrivAccess" {
            $roleAdds = @($Events | Where-Object { $_.Operation -match "Add.*member" }).Count
            $roleRemoves = @($Events | Where-Object { $_.Operation -match "Remove.*member" }).Count
            $pimActivations = @($Events | Where-Object { $_.Operation -eq "PIM Role Activation" }).Count
            
            if ($roleAdds -gt 0) { $narrativeParts += "⚠ $roleAdds privileged role assignment(s) — verify authorization" }
            if ($roleRemoves -gt 0) { $narrativeParts += "$roleRemoves role removal(s)" }
            if ($pimActivations -gt 0) { $narrativeParts += "$pimActivations PIM just-in-time activation(s)" }
        }
        "CondAccess" {
            $policyDeletes = @($Events | Where-Object { $_.Operation -match "Delete" }).Count
            $policyUpdates = @($Events | Where-Object { $_.Operation -match "Update" }).Count
            $policyAdds = @($Events | Where-Object { $_.Operation -match "Add" }).Count
            
            if ($policyDeletes -gt 0) { $narrativeParts += "⚠ $policyDeletes CA policy deletion(s) — security perimeter weakened" }
            if ($policyUpdates -gt 0) { $narrativeParts += "$policyUpdates CA policy modification(s) — review exclusion changes" }
            if ($policyAdds -gt 0) { $narrativeParts += "$policyAdds new CA policy/location(s) created" }
        }
        default {
            if ($total -gt 0) { $narrativeParts += "$total event(s) recorded" }
            else { $narrativeParts += "No events recorded" }
        }
    }
    
    if ($TopActor -and $TopActor -ne "") {
        $narrativeParts += "Top actor: $TopActor"
    }
    
    $narrative = if ($narrativeParts.Count -gt 0) { $narrativeParts -join ". " + "." } else { "No notable activity." }
    
    return @{
        RiskLevel = $riskLevel
        RiskEmoji = $riskEmoji
        Narrative = $narrative
    }
}


# AI EXECUTIVE SUMMARY ENGINE (v2.5.1)
# ==================================================================================
function Invoke-OpenAIChat {
    param(
        [string]$SystemPrompt,
        [string]$UserPrompt,
        [int]$MaxTokens = 4096
    )
    
    # Determine which endpoint to use
    $useAzure = -not [string]::IsNullOrEmpty($AzureOpenAIEndpoint)
    $usePublic = -not [string]::IsNullOrEmpty($OpenAIKey)
    
    if (-not $useAzure -and -not $usePublic) {
        Write-Log "No OpenAI endpoint configured (need AzureOpenAIEndpoint or OpenAIKey)" -Level WARN
        return $null
    }
    
    # Build request body
    $body = @{
        messages              = @(
            @{ role = "system"; content = $SystemPrompt }
            @{ role = "user"; content = $UserPrompt }
        )
        max_completion_tokens = $MaxTokens
    }
    
    if ($useAzure) {
        # Azure OpenAI Chat Completions endpoint
        $deployment = if ($AzureOpenAIDeployment) { $AzureOpenAIDeployment } else { $OpenAIModel }
        $endpoint = $AzureOpenAIEndpoint.TrimEnd('/')
        $uri = "$endpoint/openai/deployments/$deployment/chat/completions?api-version=$AzureOpenAIApiVersion"
        
        # Azure uses api-key header. If no separate key, try using the OpenAIKey param
        $apiKey = if ($OpenAIKey) { $OpenAIKey } else { $null }
        
        $headers = @{ "Content-Type" = "application/json" }
        if ($apiKey) {
            $headers["api-key"] = $apiKey
        }
        else {
            # Fall back to Graph token with Cognitive Services scope (requires Cognitive Services OpenAI User role)
            Write-Log "No API key for Azure OpenAI — attempting with Graph token (requires Cognitive Services OpenAI User role)" -Level DEBUG
            $headers["Authorization"] = "Bearer $($script:GraphToken)"
        }
        
        Write-Log "Calling Azure OpenAI: $deployment via $endpoint" -Level INFO
    }
    else {
        # Public OpenAI API
        $uri = "https://api.openai.com/v1/chat/completions"
        $body["model"] = $OpenAIModel
        $headers = @{
            "Content-Type"  = "application/json"
            "Authorization" = "Bearer $OpenAIKey"
        }
        Write-Log "Calling public OpenAI: $OpenAIModel" -Level INFO
    }
    
    try {
        $jsonBody = $body | ConvertTo-Json -Depth 10
        $response = Invoke-RetryableRequest -Uri $uri -Headers $headers -Method Post -Body $jsonBody -TimeoutSec 120
        
        if ($response.choices -and $response.choices.Count -gt 0) {
            $content = $response.choices[0].message.content
            $usage = $response.usage
            Write-Log "AI response received: $($usage.total_tokens) tokens (prompt: $($usage.prompt_tokens), completion: $($usage.completion_tokens))" -Level INFO
            return $content
        }
        else {
            Write-Log "AI response had no choices" -Level WARN
            return $null
        }
    }
    catch {
        $errorDetail = $_.ToString()
        # Check for content filter (Azure OpenAI specific)
        if ($errorDetail -match "content_filter" -or $errorDetail -match "ResponsibleAIPolicyViolation") {
            Write-Log "AI request blocked by content filter. Try with -RedactForAI to strip PII." -Level WARN
        }
        else {
            Write-Log "AI request failed: $errorDetail" -Level ERROR
        }
        return $null
    }
}

function Build-AIContextPayload {
    param(
        [string]$Path,
        [hashtable]$Stats,
        [DateTime]$RangeStart,
        [DateTime]$RangeEnd,
        [string]$RunId
    )
    
    $sb = [System.Text.StringBuilder]::new()
    
    # Header
    [void]$sb.AppendLine("=== M365 SECURITY REPORT DATA ===")
    [void]$sb.AppendLine("Report Period: $($RangeStart.ToString('yyyy-MM-dd')) to $($RangeEnd.ToString('yyyy-MM-dd'))")
    [void]$sb.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
    [void]$sb.AppendLine("Run ID: $RunId")
    [void]$sb.AppendLine("")
    
    # Domain statistics
    [void]$sb.AppendLine("=== DOMAIN STATISTICS ===")
    $domainConfig = @(
        @{ Key = "Identity"; Name = "Identity & Access"; Sheet = $SheetNames.Identity; RiskCol = "RiskLevel(1-10)" }
        @{ Key = "Hybrid"; Name = "Hybrid Identity"; Sheet = $SheetNames.Hybrid; RiskCol = "RiskLevel(1-10)" }
        @{ Key = "Exchange"; Name = "Exchange Online"; Sheet = $SheetNames.Exchange; RiskCol = "RiskLevel(1-10)" }
        @{ Key = "SharePoint"; Name = "SharePoint/Teams"; Sheet = $SheetNames.SharePoint; RiskCol = "RiskLevel(1-10)" }
        @{ Key = "Endpoint"; Name = "Endpoint (MDE)"; Sheet = $SheetNames.Endpoint; RiskCol = "SeverityScore" }
        @{ Key = "DataProtect"; Name = "Data Protection"; Sheet = $SheetNames.DataProtect; RiskCol = "Severity" }
        @{ Key = "Alerts"; Name = "Security Alerts"; Sheet = $SheetNames.Alerts; RiskCol = "Severity" }
        @{ Key = "AppConsents"; Name = "App Consents"; Sheet = $SheetNames.AppConsents; RiskCol = "RiskLevel(1-10)" }
        @{ Key = "PrivAccess"; Name = "Privileged Access"; Sheet = $SheetNames.PrivAccess; RiskCol = "RiskLevel(1-10)" }
        @{ Key = "CondAccess"; Name = "Conditional Access"; Sheet = $SheetNames.CondAccess; RiskCol = "RiskLevel(1-10)" }
    )
    
    $allFindings = @()
    foreach ($dc in $domainConfig) {
        $sheetData = @()
        try {
            $sheetData = Import-Excel -Path $Path -WorksheetName $dc.Sheet -ErrorAction SilentlyContinue
            if ($null -eq $sheetData) { $sheetData = @() }
            if ($sheetData -isnot [array]) { $sheetData = @($sheetData) }
        }
        catch { }
        
        $critical = 0; $high = 0; $medium = 0; $low = 0
        foreach ($item in $sheetData) {
            $risk = 0
            try {
                $riskProp = $dc.RiskCol
                if ($item.PSObject.Properties[$riskProp]) { $risk = [int]($item.$riskProp) }
            }
            catch { }
            if ($risk -ge 9) { $critical++; $allFindings += @{ Domain = $dc.Name; Item = $item; Risk = $risk } }
            elseif ($risk -ge 7) { $high++; $allFindings += @{ Domain = $dc.Name; Item = $item; Risk = $risk } }
            elseif ($risk -ge 4) { $medium++ }
            else { $low++ }
        }
        
        $ins = if ($Stats.ContainsKey($dc.Key)) { $Stats[$dc.Key].Inserted } else { 0 }
        $skp = if ($Stats.ContainsKey($dc.Key)) { $Stats[$dc.Key].Skipped } else { 0 }
        [void]$sb.AppendLine("- $($dc.Name): $($sheetData.Count) total events ($ins new this run, $skp deduplicated) | Critical: $critical, High: $high, Medium: $medium, Low: $low")
    }
    [void]$sb.AppendLine("")
    
    # Top findings
    $top10 = $allFindings | Sort-Object { $_.Risk } -Descending | Select-Object -First 10
    if ($top10.Count -gt 0) {
        [void]$sb.AppendLine("=== TOP $($top10.Count) CRITICAL/HIGH FINDINGS ===")
        $rank = 1
        $piiCounter = 0
        foreach ($f in $top10) {
            $item = $f.Item
            $time = if ($item.Time) { $item.Time } elseif ($item.CreationTime) { $item.CreationTime } else { "N/A" }
            $desc = if ($item.Operation) { $item.Operation } elseif ($item.Title) { $item.Title } elseif ($item.RecommendationName) { $item.RecommendationName } else { "N/A" }
            $actor = if ($item.Actor) { $item.Actor } elseif ($item.User) { $item.User } elseif ($item.Account) { $item.Account } else { "N/A" }
            
            if ($RedactForAI) {
                $piiCounter++
                if ($actor -ne "N/A") { $actor = "[USER-$($piiCounter.ToString('D3'))]" }
            }
            
            [void]$sb.AppendLine("  $rank. [$($f.Domain)] Severity: $($f.Risk)/10 | $desc | Actor: $actor | Time: $time")
            $rank++
        }
        [void]$sb.AppendLine("")
    }
    
    # MFA Status
    [void]$sb.AppendLine("=== MFA REGISTRATION STATUS ===")
    if ($null -ne $script:MfaStats -and $script:MfaStats.TotalUsers -gt 0) {
        $pct = [math]::Round([double](($script:MfaStats.Registered / $script:MfaStats.TotalUsers) * 100), 1)
        [void]$sb.AppendLine("Total Users: $($script:MfaStats.TotalUsers)")
        [void]$sb.AppendLine("MFA Registered: $($script:MfaStats.Registered) ($pct%)")
        [void]$sb.AppendLine("NOT Registered: $($script:MfaStats.NotRegistered)")
    }
    else {
        [void]$sb.AppendLine("MFA data not available for this run")
    }
    [void]$sb.AppendLine("")
    
    # Historical trend from Runs sheet
    try {
        $runsData = Import-Excel -Path $Path -WorksheetName $SheetNames.Runs -ErrorAction SilentlyContinue
        if ($null -ne $runsData) {
            if ($runsData -isnot [array]) { $runsData = @($runsData) }
            if ($runsData.Count -gt 0) {
                [void]$sb.AppendLine("=== HISTORICAL RUN TREND ===")
                $recentRuns = $runsData | Select-Object -Last 10
                foreach ($run in $recentRuns) {
                    [void]$sb.AppendLine("  Run $($run.RunId): $($run.DateRangeStart) to $($run.DateRangeEnd) | $($run.TotalInserted) new findings | v$($run.Version)")
                }
                [void]$sb.AppendLine("")
            }
        }
    }
    catch { }
    
    return $sb.ToString()
}

function Update-ExecutiveSummary {
    param(
        [string]$Path,
        [hashtable]$Stats,
        [DateTime]$RangeStart,
        [DateTime]$RangeEnd,
        [string]$RunId
    )
    
    Write-Host "`n[AI] Generating executive summary..." -ForegroundColor Cyan
    
    # Determine summary file path — output as .doc (HTML format Word opens natively)
    $summaryFileName = "Executive-Summary-$targetMonth.doc"
    $summaryFilePath = if ($ExecutiveSummaryPath) { $ExecutiveSummaryPath } else { Join-Path $script:ScriptRoot $summaryFileName }
    $script:ExecutiveSummaryFilePath = $summaryFilePath
    $script:ExecutiveSummaryFileName = $summaryFileName
    
    # Read the cumulative AI Context file as the data source
    $contextPayload = $null
    if ($script:AIContextPath -and (Test-Path $script:AIContextPath)) {
        try {
            $contextPayload = Get-Content $script:AIContextPath -Raw -ErrorAction Stop
            Write-Log "Read AI Context file: $script:AIContextPath ($($contextPayload.Length) chars)" -Level INFO
        }
        catch {
            Write-Log "Failed to read AI Context file: $_" -Level WARN
        }
    }
    
    # Fallback: build payload from Excel if AI Context file not available
    if ([string]::IsNullOrEmpty($contextPayload)) {
        Write-Log "AI Context file not available — falling back to Build-AIContextPayload" -Level INFO
        $contextPayload = Build-AIContextPayload -Path $Path -Stats $Stats -RangeStart $RangeStart -RangeEnd $RangeEnd -RunId $RunId
    }
    
    if ([string]::IsNullOrEmpty($contextPayload)) {
        Write-Host "  [--] No data available for executive summary — skipping" -ForegroundColor Yellow
        return
    }
    
    # System prompt — executive intelligence brief persona
    $systemPrompt = @"
You are a senior Microsoft 365 security analyst preparing a formal executive intelligence brief for C-suite leadership and board members.

Using the cumulative Microsoft 365 security report data provided below, produce a professional executive summary suitable for executive circulation.

This report reflects cumulative data from multiple runs within the same month. Analyze trends across all runs shown, not only the most recent execution.

Writing Requirements:

Use formal business language appropriate for executive leadership.

Do not use markdown formatting (no bold styling, no decorative symbols, no ASCII graphics).

Do not include time-bound remediation language (no references to "30 days," "60 days," "immediate," etc.).

Avoid unnecessary technical jargon.

Do not reference script versions or run identifiers unless strategically relevant.

Maintain a calm, controlled, and authoritative tone.

Focus on business risk, governance posture, and oversight implications.

Structure the summary using the following section headings written in plain text:

Executive Overview
Provide a concise summary of overall posture, score, rating, and trend direction.

Current Security Posture Interpretation
Interpret the meaning of the current posture score and stability across runs.

Primary Risk Concentrations
Present the 3 to 5 most significant findings requiring executive attention.

Identity and MFA Assessment
State registration coverage percentage.
Quantify the gap.
Explain risk implications in business terms.

Endpoint Risk Analysis
Summarize the concentration of critical findings.
Explain enterprise impact exposure.

Business Email Compromise Indicators
Summarize any inbox rule changes, anomaly indicators, or early warning signs.
Frame implications in terms of financial and reputational risk.

Trend Commentary
Compare posture trends, deduplication patterns, and recurring findings across all runs.
Identify whether risk is improving, stagnating, or concentrating.

Strategic Recommendations
Provide 3 to 5 prioritized leadership-level actions.
Frame these as governance and oversight priorities rather than technical tasks.
Emphasize risk reduction, accountability, and control maturity.
Do not assign deadlines or operational timelines.

Ensure the final output reads as a cohesive executive briefing document suitable for board-level review.
"@
    
    # Build the user prompt with the cumulative data
    $userPrompt = @"
Here is the cumulative Microsoft 365 security report data for $targetMonth :

$contextPayload

Please produce the executive intelligence brief based on this data.
"@
    
    # Call OpenAI
    $summaryContent = Invoke-OpenAIChat -SystemPrompt $systemPrompt -UserPrompt $userPrompt -MaxTokens 4096
    
    if ([string]::IsNullOrEmpty($summaryContent)) {
        Write-Host "  [--] AI summary generation failed or returned empty — skipping" -ForegroundColor Yellow
        return
    }
    
    # Convert plain text to styled HTML Word document
    # Detect section headings and wrap in proper HTML tags
    $headings = @(
        'Executive Overview',
        'Current Security Posture Interpretation',
        'Primary Risk Concentrations',
        'Identity and MFA Assessment',
        'Endpoint Risk Analysis',
        'Business Email Compromise Indicators',
        'Trend Commentary',
        'Strategic Recommendations'
    )
    
    $aiLines = $summaryContent -split "`n"
    $htmlBody = [System.Text.StringBuilder]::new()
    foreach ($line in $aiLines) {
        $trimmed = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }
        $isHeading = $false
        foreach ($h in $headings) {
            if ($trimmed -eq $h -or $trimmed -match "^$([regex]::Escape($h))") {
                [void]$htmlBody.AppendLine("<h2>$([System.Net.WebUtility]::HtmlEncode($trimmed))</h2>")
                $isHeading = $true
                break
            }
        }
        if (-not $isHeading) {
            [void]$htmlBody.AppendLine("<p>$([System.Net.WebUtility]::HtmlEncode($trimmed))</p>")
        }
    }
    
    $htmlContent = @"
<html>
<head>
<meta charset="UTF-8">
<style>
    body {
        font-family: Calibri, 'Segoe UI', Arial, sans-serif;
        font-size: 11pt;
        line-height: 1.6;
        color: #333333;
        max-width: 800px;
        margin: 40px auto;
        padding: 0 40px;
    }
    h1 {
        font-size: 16pt;
        color: #1a3a5c;
        border-bottom: 2px solid #1a3a5c;
        padding-bottom: 8px;
        margin-top: 30px;
    }
    h2 {
        font-size: 13pt;
        color: #2c5f8a;
        margin-top: 24px;
        margin-bottom: 8px;
    }
    p {
        margin: 8px 0;
        text-align: justify;
    }
    .header {
        text-align: center;
        border-bottom: 3px solid #1a3a5c;
        padding-bottom: 16px;
        margin-bottom: 30px;
    }
    .header h1 {
        border-bottom: none;
        margin-bottom: 4px;
    }
    .header .subtitle {
        font-size: 10pt;
        color: #666666;
    }
    .footer {
        margin-top: 40px;
        padding-top: 12px;
        border-top: 1px solid #cccccc;
        font-size: 9pt;
        color: #999999;
        text-align: center;
    }
</style>
</head>
<body>
<div class="header">
    <h1>Microsoft 365 Security Executive Intelligence Brief</h1>
    <div class="subtitle">$targetMonth | Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') | Confidential</div>
</div>
$($htmlBody.ToString())
<div class="footer">
    This document was generated by the M365 Security Report automation system v$($script:Version).<br>
    Classification: Confidential - Internal Use Only
</div>
</body>
</html>
"@
    
    # Save as .doc (HTML format that Word opens natively with full styling)
    try {
        $htmlContent | Out-File -FilePath $summaryFilePath -Encoding UTF8 -Force
        Write-Host "  [OK] Executive summary saved: $summaryFilePath" -ForegroundColor Green
    }
    catch {
        Write-Log "Failed to save executive summary file: $_" -Level ERROR
    }
    
    # Also write to AI-Summary sheet in Excel (plain text version)
    try {
        $excel = Open-ExcelPackage -Path $Path
        
        # Remove old AI-Summary sheet if exists
        $oldAISheet = $excel.Workbook.Worksheets["AI-Summary"]
        if ($null -ne $oldAISheet) {
            $excel.Workbook.Worksheets.Delete("AI-Summary")
        }
        Close-ExcelPackage $excel
        
        # Write summary as rows to AI-Summary sheet
        $aiRows = @()
        foreach ($line in $aiLines) {
            $aiRows += [PSCustomObject]@{ Content = $line }
        }
        $aiRows | Export-Excel -Path $Path -WorksheetName "AI-Summary" -NoHeader -AutoSize
        Write-Host "  [OK] AI-Summary sheet updated in workbook" -ForegroundColor Green
    }
    catch {
        Write-Log "Failed to write AI-Summary sheet: $_" -Level WARN
    }
    
    Write-Host "  Executive summary generated for $targetMonth" -ForegroundColor Green
}


# ==================================================================================
# O365 MANAGEMENT API
# ==================================================================================
function Start-O365Subscription {
    param([string]$Token, [string]$ContentType)
    $headers = @{ Authorization = "Bearer $Token" }
    $listUri = "https://manage.office.com/api/v1.0/$TenantId/activity/feed/subscriptions/list"
    try {
        $subs = Invoke-RetryableRequest -Uri $listUri -Headers $headers
        $existing = $subs | Where-Object { $_.contentType -eq $ContentType }
        if ($existing -and $existing.status -eq "enabled") { return $true }
    }
    catch { }
    $startUri = "https://manage.office.com/api/v1.0/$TenantId/activity/feed/subscriptions/start?contentType=$ContentType"
    try { Invoke-RetryableRequest -Uri $startUri -Headers $headers -Method Post | Out-Null; return $true }
    catch { return $false }
}

function Get-O365AuditRecordsParallel {
    param(
        [string]$Token, 
        [string]$ContentType, 
        [DateTime]$Start, 
        [DateTime]$End, 
        [int]$Throttle = 10,
        [string[]]$FilterOperations = $null  # NEW: Filter operations during download
    )
    
    # Clamp to 7 days (API limit)
    $sevenDaysAgo = (Get-Date).AddDays(-7)
    $effectiveStart = if ($Start -lt $sevenDaysAgo) { $sevenDaysAgo } else { $Start }
    
    $null = Start-O365Subscription -Token $Token -ContentType $ContentType
    $headers = @{ Authorization = "Bearer $Token" }
    $allBlobUris = [System.Collections.Generic.List[string]]::new()
    $currentStart = $effectiveStart
    
    # Phase 1: List blobs
    while ($currentStart -lt $End) {
        $chunkEnd = [DateTime]::MinValue + [TimeSpan]::FromTicks([Math]::Min(($currentStart.AddHours(24)).Ticks, $End.Ticks))
        if ($chunkEnd -le $currentStart) { $chunkEnd = $currentStart.AddHours(24) }
        if ($chunkEnd -gt $End) { $chunkEnd = $End }
        
        $st = $currentStart.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss")
        $et = $chunkEnd.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss")
        $listUri = "https://manage.office.com/api/v1.0/$TenantId/activity/feed/subscriptions/content?contentType=$ContentType&startTime=$st&endTime=$et"
        
        try {
            $blobs = Invoke-RetryableRequest -Uri $listUri -Headers $headers
            if ($blobs) {
                if ($blobs -isnot [array]) { $blobs = @($blobs) }
                foreach ($b in $blobs) { if ($b.contentUri) { $allBlobUris.Add($b.contentUri) } }
            }
        }
        catch { Write-Log "Failed to list $ContentType chunk" -Level WARN }
        $currentStart = $chunkEnd
    }
    
    if ($allBlobUris.Count -eq 0) { return @() }
    
    $totalBlobs = $allBlobUris.Count
    $filterMode = if ($FilterOperations) { "filtering for $($FilterOperations.Count) operations" } else { "no filter" }
    Write-Host "          Downloading $totalBlobs blobs ($Throttle threads, $filterMode)..." -ForegroundColor DarkGray
    
    # Phase 2: Parallel download in batches with STREAMING filter
    $batchSize = 50  # Smaller batches to reduce memory spikes
    $filteredEvents = [System.Collections.Generic.List[object]]::new()
    $totalDownloaded = 0
    $totalFiltered = 0
    
    for ($i = 0; $i -lt $allBlobUris.Count; $i += $batchSize) {
        $batch = $allBlobUris[$i..([Math]::Min($i + $batchSize - 1, $allBlobUris.Count - 1))]
        $batchNum = [math]::Floor($i / $batchSize) + 1
        $totalBatches = [math]::Ceiling($totalBlobs / $batchSize)
        
        # Progress indicator
        $pct = [math]::Round(($i / $totalBlobs) * 100)
        Write-Host "`r          Batch $batchNum/$totalBatches ($pct%) - found $($filteredEvents.Count) matching events..." -NoNewline -ForegroundColor DarkGray
        
        # Parallel download with in-flight filtering
        $results = $batch | ForEach-Object -Parallel {
            $uri = $_; $tok = $using:Token; $filterOps = $using:FilterOperations
            try { 
                $events = Invoke-RestMethod -Uri $uri -Headers @{Authorization = "Bearer $tok" } -TimeoutSec 15
                if ($null -eq $events) { return $null }
                
                # Filter during download if filter provided
                if ($filterOps -and $filterOps.Count -gt 0) {
                    $filtered = @()
                    foreach ($evt in $events) {
                        foreach ($op in $filterOps) {
                            if ($evt.Operation -eq $op) { $filtered += $evt; break }
                        }
                    }
                    return $filtered
                }
                return $events 
            }
            catch { $null }
        } -ThrottleLimit $Throttle
        
        # Collect results
        foreach ($r in $results) {
            if ($null -ne $r) {
                if ($r -is [array]) { 
                    foreach ($item in $r) { 
                        $filteredEvents.Add($item)
                        $totalDownloaded++
                    }
                }
                else { 
                    $filteredEvents.Add($r)
                    $totalDownloaded++
                }
            }
        }
        
        # Aggressive GC every 5 batches
        if ($batchNum % 5 -eq 0) { 
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
    
    Write-Host "`r          Retrieved $($filteredEvents.Count) events (filtered from ~$totalDownloaded raw)          " -ForegroundColor DarkGray
    return $filteredEvents.ToArray()
}

# ==================================================================================
# RISK SCORING
# ==================================================================================
function Get-RiskScore {
    param([string]$EventType, [object]$Data)
    $score = 5
    switch ($EventType) {
        "SignIn" {
            if ($Data.riskLevelDuringSignIn) {
                $score = switch ($Data.riskLevelDuringSignIn) { "high" { 9 } "medium" { 6 } "low" { 3 } default { 1 } }
            }
            if ($Data.status.errorCode -in @(50053, 50126, 50055, 50056, 50057)) { $score = [Math]::Min($score + 2, 10) }
        }
        "Exchange" {
            $op = $Data.Operation
            if ($op -match "New-InboxRule|Set-InboxRule|Remove-InboxRule") { $score = 8 }
            elseif ($op -match "Set-TransportRule|New-TransportRule|Remove-TransportRule") { $score = 7 }
            elseif ($op -match "Add-MailboxPermission|Remove-MailboxPermission|Add-RecipientPermission") { $score = 7 }
            elseif ($op -match "Set-MailboxJunkEmailConfiguration") { $score = 5 }
            elseif ($op -match "Set-OwaMailboxPolicy") { $score = 6 }
            elseif ($op -match "Set-Mailbox") { $score = 5 }
        }
        "SharePoint" {
            $op = $Data.Operation
            if ($op -match "FileVersionsAllDeleted") { $score = 10 }
            elseif ($op -match "AnonymousLinkUsed") { $score = 9 }
            elseif ($op -match "AnonymousLinkCreated") { $score = 8 }
            elseif ($op -match "SharingSet|SharingInvitation") { $score = 6 }
        }
        "Alert" {
            $score = switch ($Data.severity) { "high" { 9 } "medium" { 6 } "low" { 3 } default { 1 } }
        }
    }
    return $score
}

# ==================================================================================
# SELF-TEST MODE
# ==================================================================================
function Invoke-SelfTest {
    Write-Host "`n=== SELF-TEST MODE ===" -ForegroundColor Cyan
    $tempDir = if ($env:TEMP) { $env:TEMP } elseif ($env:TMPDIR) { $env:TMPDIR } else { "/tmp" }
    $testPath = Join-Path $tempDir "SelfTest-$(Get-Date -Format 'yyyyMMddHHmmss').xlsx"
    $passed = 0; $failed = 0
    
    # Test 1: Hash consistency
    Write-Host "  [1] Hash consistency..." -NoNewline
    $h1 = Get-HashString "test|value|123"
    $h2 = Get-HashString "test|value|123"
    if ($h1 -eq $h2 -and $h1.Length -eq 16) { Write-Host " PASS" -ForegroundColor Green; $passed++ }
    else { Write-Host " FAIL" -ForegroundColor Red; $failed++ }
    
    # Test 2: Workbook creation
    Write-Host "  [2] Workbook creation..." -NoNewline
    try {
        New-ReportWorkbook -Path $testPath
        $excel = Open-ExcelPackage -Path $testPath
        $sheets = $excel.Workbook.Worksheets | ForEach-Object { $_.Name }
        Close-ExcelPackage $excel -NoSave
        if ($sheets.Count -ge 12) { Write-Host " PASS ($($sheets.Count) sheets)" -ForegroundColor Green; $passed++ }
        else { Write-Host " FAIL" -ForegroundColor Red; $failed++ }
    }
    catch { Write-Host " FAIL: $_" -ForegroundColor Red; $failed++ }
    
    # Test 3: Dedupe logic
    Write-Host "  [3] Deduplication..." -NoNewline
    try {
        $testRow = [PSCustomObject]@{ Time = (Get-Date).ToString("o"); Source = "Test"; Operation = "Test"; Actor = "Test"; Target = "Test"; Result = "OK"; 'RiskLevel(1-10)' = 5; Category = "Test"; DedupKey = "test123"; RawJson = "{}" }
        $testRow | Export-Excel -Path $testPath -WorksheetName $SheetNames.Identity -Append
        $keys = Build-ExistingKeyIndex -Path $testPath -SheetName $SheetNames.Identity
        if ($keys.Contains("test123")) { Write-Host " PASS" -ForegroundColor Green; $passed++ }
        else { Write-Host " FAIL" -ForegroundColor Red; $failed++ }
    }
    catch { Write-Host " FAIL: $_" -ForegroundColor Red; $failed++ }
    
    # Cleanup
    Remove-Item $testPath -Force -ErrorAction SilentlyContinue
    
    Write-Host "`n  Results: $passed passed, $failed failed" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Yellow" })
    return $failed -eq 0
}

# ==================================================================================
# MAIN EXECUTION
# ==================================================================================
if ($SelfTest) {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "ImportExcel module required for self-test" -ForegroundColor Red
        exit 1
    }
    Import-Module ImportExcel -ErrorAction Stop
    $result = Invoke-SelfTest
    exit $(if ($result) { 0 } else { 1 })
}

# Normal execution
try {
    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host "  M365 Security Report v$($script:Version) - PowerShell 7" -ForegroundColor Cyan
    Write-Host "==========================================================================" -ForegroundColor Cyan
    Write-Host "  Date Range:     $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))" -ForegroundColor White
    Write-Host "  Target Month:   $targetMonth" -ForegroundColor White
    Write-Host "  Report:         $ReportPath" -ForegroundColor White
    Write-Host "  Throttle:       $ThrottleLimit threads" -ForegroundColor White
    Write-Host ""
    
    # Check ImportExcel
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "ImportExcel module required. Install: Install-Module ImportExcel -Scope CurrentUser"
    }
    Import-Module ImportExcel -ErrorAction Stop
    
    # Authenticate (needed before SharePoint operations)
    Write-Host "Authenticating..." -ForegroundColor Yellow
    try {
        $script:GraphToken = Get-GraphToken
        $script:O365Token = Get-O365Token
        $script:TokenAcquiredTime = Get-Date
        Write-Host "  [OK] Graph + O365 Management" -ForegroundColor Green
    }
    catch {
        throw "Authentication failed: $_"
    }
    
    # ==================================================================================
    # SHAREPOINT INTEGRATION (v2.2.0)
    # ==================================================================================
    if ($SharePointSiteUrl) {
        Write-Host "`n[SharePoint] Initializing cloud storage..." -ForegroundColor Cyan
        try {
            $script:SharePointSiteId = Get-SharePointSiteId -SiteUrl $SharePointSiteUrl -Token $script:GraphToken
            Write-Host "  [OK] Site resolved: $($script:SharePointSiteId.Split(',')[1])" -ForegroundColor Green
            $script:UseSharePoint = $true
            
            # Determine temp path for local operations
            $tempDir = if ($env:TEMP) { $env:TEMP } elseif ($env:TMPDIR) { $env:TMPDIR } else { "/tmp" }
            $script:LocalTempPath = Join-Path $tempDir "Security-Monthly-Report-$targetMonth.xlsx"
            $spFileName = "Security-Monthly-Report-$targetMonth.xlsx"
            
            # Try to download existing report from SharePoint
            $downloaded = Get-SharePointFile -SiteId $script:SharePointSiteId -FolderPath $SharePointFolder -FileName $spFileName -LocalPath $script:LocalTempPath -Token $script:GraphToken
            
            if ($downloaded) {
                Write-Host "  [OK] Downloaded existing report from SharePoint" -ForegroundColor Green
                $ReportPath = $script:LocalTempPath
            }
            else {
                Write-Host "  [--] No existing report in SharePoint - will create new" -ForegroundColor DarkGray
                $ReportPath = $script:LocalTempPath
            }
        }
        catch {
            Write-Log "SharePoint initialization failed: $_ - falling back to local storage" -Level WARN
            Write-Host "  [!!] SharePoint failed - using local storage" -ForegroundColor Yellow
            $script:UseSharePoint = $false
        }
    }
    
    # Update display with resolved path
    Write-Host "  Working Path:   $ReportPath" -ForegroundColor White
    
    # Create or verify workbook
    if (-not (Test-Path $ReportPath)) {
        if ($PSCmdlet.ShouldProcess($ReportPath, "Create new workbook")) {
            New-ReportWorkbook -Path $ReportPath
        }
    }
    else {
        Write-Log "Using existing workbook: $ReportPath" -Level INFO
    }
    
    $runId = [Guid]::NewGuid().ToString().Substring(0, 8)
    $domainStats = @{}
    $domainResults = @()
    
    # Azure token no longer needed (Azure-Platform removed in v2.5.1)
    
    $ActiveDefenderRegion = $null
    if (-not $SkipDefenderEndpoint) {
        try {
            $script:DefenderToken = Get-DefenderToken
            $ActiveDefenderRegion = if ($DefenderRegion) { $DefenderRegion } else { Find-DefenderEndpoint -Token $script:DefenderToken }
            if ($ActiveDefenderRegion) { Write-Host "  [OK] Defender ($ActiveDefenderRegion)" -ForegroundColor Green }
            else { $SkipDefenderEndpoint = $true }
        }
        catch { Write-Host "  [--] Defender (skipped)" -ForegroundColor DarkGray; $SkipDefenderEndpoint = $true }
    }
    
    $startIso = $StartDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    
    # ============ 1. IDENTITY ============
    Write-Host "`n[1/11] Identity & Access..." -ForegroundColor Cyan
    $existingKeys = Build-ExistingKeyIndex -Path $ReportPath -SheetName $SheetNames.Identity
    # Defensive: ensure we have a valid HashSet
    if ($null -eq $existingKeys) { $existingKeys = [System.Collections.Generic.HashSet[string]]::new() }
    $inserted = 0; $skipped = 0
    
    try {
        $signIns = Invoke-GraphWithPagination -Uri "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=createdDateTime ge $startIso and status/errorCode ne 0&`$top=500" -Token $script:GraphToken -Label "Sign-ins"
        $riskyUsers = @()
        try { $riskyUsers = Invoke-GraphWithPagination -Uri "https://graph.microsoft.com/v1.0/identityProtection/riskyUsers?`$filter=riskState ne 'remediated'" -Token $script:GraphToken -Label "Risky Users" } catch { }
        
        $rows = @()
        foreach ($si in $signIns) {
            $key = Get-DedupKey -Domain "Identity" -Item $si
            if ($existingKeys.Contains($key)) { $skipped++; continue }
            $rows += [PSCustomObject]@{
                FindingId = Get-FindingId -Prefix "ID"
                Time = $si.createdDateTime; Source = $si.appDisplayName; Operation = "Sign-in Failure"
                Actor = $si.userDisplayName; Target = $si.resourceDisplayName
                Result = "Error: $($si.status.errorCode)"; 'RiskLevel(1-10)' = Get-RiskScore -EventType "SignIn" -Data $si
                Category = "AuthFailure"; DedupKey = $key; RawJson = ($si | ConvertTo-Json -Depth 5 -Compress)
            }
            $inserted++
        }
        foreach ($ru in $riskyUsers) {
            $key = Get-DedupKey -Domain "Identity" -Item $ru
            if ($existingKeys.Contains($key)) { $skipped++; continue }
            $sev = switch ($ru.riskLevel) { "high" { 9 } "medium" { 6 } default { 3 } }
            $rows += [PSCustomObject]@{
                FindingId = Get-FindingId -Prefix "ID"
                Time = $ru.riskLastUpdatedDateTime; Source = "Identity Protection"; Operation = "Risky User"
                Actor = $ru.userDisplayName; Target = $ru.userPrincipalName; Result = "Risk: $($ru.riskLevel)"
                'RiskLevel(1-10)' = $sev; Category = "RiskyUser"; DedupKey = $key; RawJson = ($ru | ConvertTo-Json -Depth 5 -Compress)
            }
            $inserted++
        }
        if ($rows.Count -gt 0 -and $PSCmdlet.ShouldProcess("Identity", "Append $($rows.Count) rows")) {
            $rows | Export-Excel -Path $ReportPath -WorksheetName $SheetNames.Identity -Append
        }
        $domainResults += "Identity"
    }
    catch { Write-Log "Identity failed: $_" -Level ERROR }
    $domainStats["Identity"] = @{ Inserted = $inserted; Skipped = $skipped }
    Write-Host "  Inserted: $inserted | Skipped: $skipped" -ForegroundColor Gray
    
    # ============ 2-4. DIRECTORY AUDITS → Hybrid + AppConsents + CondAccess + PrivAccess ============
    # Fetch all directory audits ONCE and distribute to correct sheets
    Write-Host "`n[2/11] Fetching Directory Audits (shared source)..." -ForegroundColor Cyan
    $allDirAudits = @()
    try {
        Refresh-TokensIfNeeded
        $allDirAudits = Invoke-GraphWithPagination -Uri "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=activityDateTime ge $startIso&`$top=500" -Token $script:GraphToken -Label "Directory Audits"
        Write-Host "  Fetched $($allDirAudits.Count) total directory audit events" -ForegroundColor DarkGray
    }
    catch { Write-Log "Directory Audits fetch failed: $_" -Level ERROR }
    
    # Define operation filters for each domain
    $hybridOps = @("sync", "federation", "password hash", "connect", "provisioning", "hybrid", "adfs")
    $appConsentOps = @(
        "Add application", "Update application", "Delete application",
        "Add service principal", "Update service principal", "Delete service principal",
        "Consent to application", "Add delegated permission grant", "Remove delegated permission grant",
        "Add app role assignment to service principal", "Remove app role assignment from service principal",
        "Add service principal credentials", "Remove service principal credentials",
        "Add owner to application", "Add owner to service principal",
        "Update application – Certificates and secrets management"
    )
    $condAccessOps = @(
        "Add conditional access policy", "Update conditional access policy", "Delete conditional access policy",
        "Add named location", "Update named location", "Delete named location"
    )
    $privAccessOps = @(
        "Add member to role", "Remove member from role",
        "Add eligible member to role", "Remove eligible member from role",
        "Add scoped member to role", "Remove scoped member from role"
    )
    
    # Helper to extract initiatedBy display name
    function Get-AuditActor {
        param($Audit)
        if ($Audit.initiatedBy.app) { return $Audit.initiatedBy.app.displayName }
        elseif ($Audit.initiatedBy.user) { return $Audit.initiatedBy.user.displayName }
        else { return "System" }
    }
    
    # ---- 2. HYBRID (security-relevant sync/federation) ----
    Write-Host "  [2/11] Hybrid Identity..." -ForegroundColor Cyan
    $existingKeys = Build-ExistingKeyIndex -Path $ReportPath -SheetName $SheetNames.Hybrid
    if ($null -eq $existingKeys) { $existingKeys = [System.Collections.Generic.HashSet[string]]::new() }
    $inserted = 0; $skipped = 0
    try {
        $hybridEvents = $allDirAudits | Where-Object { 
            $actName = $_.activityDisplayName
            if (-not $actName) { return $false }
            $actLower = $actName.ToLower()
            foreach ($kw in $hybridOps) { if ($actLower -match $kw) { return $true } }
            return $false
        }
        
        $rows = @()
        foreach ($h in $hybridEvents) {
            $key = Get-DedupKey -Domain "Hybrid" -Item $h
            if ($existingKeys.Contains($key)) { $skipped++; continue }
            $init = Get-AuditActor -Audit $h
            $rows += [PSCustomObject]@{
                FindingId = Get-FindingId -Prefix "HY"
                Time = $h.activityDateTime; Source = $init; Operation = $h.activityDisplayName; Actor = $init
                Target = if ($h.targetResources.Count) { $h.targetResources[0].displayName } else { "" }
                Result = $h.result; 'RiskLevel(1-10)' = if ($h.result -eq "failure") { 7 } else { 3 }
                Details = ""; DedupKey = $key; RawJson = ($h | ConvertTo-Json -Depth 5 -Compress)
            }
            $inserted++
        }
        if ($rows.Count -gt 0 -and $PSCmdlet.ShouldProcess("Hybrid", "Append")) { $rows | Export-Excel -Path $ReportPath -WorksheetName $SheetNames.Hybrid -Append }
        $domainResults += "Hybrid"
    }
    catch { Write-Log "Hybrid failed: $_" -Level ERROR }
    $domainStats["Hybrid"] = @{ Inserted = $inserted; Skipped = $skipped }
    Write-Host "    Inserted: $inserted | Skipped: $skipped" -ForegroundColor Gray
    
    # ---- 3. APP CONSENTS (OAuth, service principals) ----
    Write-Host "  [3/11] App Consents..." -ForegroundColor Cyan
    $existingKeys = Build-ExistingKeyIndex -Path $ReportPath -SheetName $SheetNames.AppConsents
    if ($null -eq $existingKeys) { $existingKeys = [System.Collections.Generic.HashSet[string]]::new() }
    $inserted = 0; $skipped = 0
    try {
        $consentEvents = $allDirAudits | Where-Object { $_.activityDisplayName -in $appConsentOps }
        
        $rows = @()
        foreach ($c in $consentEvents) {
            $key = Get-DedupKey -Domain "AppConsents" -Item $c
            if ($existingKeys.Contains($key)) { $skipped++; continue }
            $init = Get-AuditActor -Audit $c
            $appName = if ($c.targetResources.Count) { $c.targetResources[0].displayName } else { "Unknown" }
            $appId = if ($c.targetResources.Count) { $c.targetResources[0].id } else { "" }
            
            # Extract permissions from modified properties
            $perms = ""
            if ($c.targetResources.Count -gt 0 -and $c.targetResources[0].modifiedProperties) {
                $permProp = $c.targetResources[0].modifiedProperties | Where-Object { $_.displayName -match "Scope|Permission|DelegatedPermission" } | Select-Object -First 1
                if ($permProp) { $perms = $permProp.newValue }
            }
            
            # Determine consent type
            $consentType = switch -Wildcard ($c.activityDisplayName) {
                "Consent*" { "User Consent" }
                "*delegated*" { "Delegated" }
                "*app role*" { "Application" }
                "*credentials*" { "Credential" }
                "*owner*" { "Ownership" }
                default { "Configuration" }
            }
            
            # Risk scoring: credential additions and consent grants are high risk
            $risk = switch -Wildcard ($c.activityDisplayName) {
                "*credentials*" { 8 }
                "Consent*" { 7 }
                "*delegated*" { 7 }
                "*app role*" { 7 }
                "*owner*" { 6 }
                "*Add application*" { 5 }
                default { 4 }
            }
            
            $rows += [PSCustomObject]@{
                FindingId = Get-FindingId -Prefix "AC"
                Time = $c.activityDateTime; Operation = $c.activityDisplayName
                AppName = $appName; AppId = $appId; Permissions = $perms
                InitiatedBy = $init; ConsentType = $consentType
                Result = $c.result; 'RiskLevel(1-10)' = $risk
                DedupKey = $key; RawJson = ($c | ConvertTo-Json -Depth 5 -Compress)
            }
            $inserted++
        }
        if ($rows.Count -gt 0 -and $PSCmdlet.ShouldProcess("AppConsents", "Append")) { $rows | Export-Excel -Path $ReportPath -WorksheetName $SheetNames.AppConsents -Append }
        $domainResults += "AppConsents"
    }
    catch { Write-Log "AppConsents failed: $_" -Level ERROR }
    $domainStats["AppConsents"] = @{ Inserted = $inserted; Skipped = $skipped }
    Write-Host "    Inserted: $inserted | Skipped: $skipped" -ForegroundColor Gray
    
    # ---- 4. CONDITIONAL ACCESS ----
    Write-Host "  [4/11] Conditional Access..." -ForegroundColor Cyan
    $existingKeys = Build-ExistingKeyIndex -Path $ReportPath -SheetName $SheetNames.CondAccess
    if ($null -eq $existingKeys) { $existingKeys = [System.Collections.Generic.HashSet[string]]::new() }
    $inserted = 0; $skipped = 0
    try {
        $caEvents = $allDirAudits | Where-Object { $_.activityDisplayName -in $condAccessOps }
        
        $rows = @()
        foreach ($ca in $caEvents) {
            $key = Get-DedupKey -Domain "CondAccess" -Item $ca
            if ($existingKeys.Contains($key)) { $skipped++; continue }
            $init = Get-AuditActor -Audit $ca
            $policyName = if ($ca.targetResources.Count) { $ca.targetResources[0].displayName } else { "Unknown" }
            $policyId = if ($ca.targetResources.Count) { $ca.targetResources[0].id } else { "" }
            
            # Extract change details from modified properties
            $changeDetails = ""
            if ($ca.targetResources.Count -gt 0 -and $ca.targetResources[0].modifiedProperties) {
                $changes = $ca.targetResources[0].modifiedProperties | ForEach-Object { "$($_.displayName): $($_.newValue)" }
                $changeDetails = ($changes -join "; ").Substring(0, [Math]::Min(500, ($changes -join "; ").Length))
            }
            
            # Risk scoring: deletes and updates are higher risk than adds
            $risk = switch -Wildcard ($ca.activityDisplayName) {
                "Delete*" { 9 }
                "Update*" { 7 }
                "Add*" { 5 }
                default { 4 }
            }
            
            $rows += [PSCustomObject]@{
                FindingId = Get-FindingId -Prefix "CA"
                Time = $ca.activityDateTime; Operation = $ca.activityDisplayName
                PolicyName = $policyName; PolicyId = $policyId
                ModifiedBy = $init; ChangeDetails = $changeDetails
                Result = $ca.result; 'RiskLevel(1-10)' = $risk
                DedupKey = $key; RawJson = ($ca | ConvertTo-Json -Depth 5 -Compress)
            }
            $inserted++
        }
        if ($rows.Count -gt 0 -and $PSCmdlet.ShouldProcess("CondAccess", "Append")) { $rows | Export-Excel -Path $ReportPath -WorksheetName $SheetNames.CondAccess -Append }
        $domainResults += "CondAccess"
    }
    catch { Write-Log "CondAccess failed: $_" -Level ERROR }
    $domainStats["CondAccess"] = @{ Inserted = $inserted; Skipped = $skipped }
    Write-Host "    Inserted: $inserted | Skipped: $skipped" -ForegroundColor Gray

    # ============ 5-6: Exchange + SharePoint (unchanged from v2.5.0) ============
    Write-Host "[3/11] Exchange Online..." -ForegroundColor Cyan
    $existingKeys = Build-ExistingKeyIndex -Path $ReportPath -SheetName $SheetNames.Exchange
    if ($null -eq $existingKeys) { $existingKeys = [System.Collections.Generic.HashSet[string]]::new() }
    $inserted = 0; $skipped = 0
    try {
        Refresh-TokensIfNeeded
        # Filter during download to save memory
        $exOps = @("Set-Mailbox", "New-InboxRule", "Set-InboxRule", "Remove-InboxRule", "Set-TransportRule", "Add-MailboxPermission", "Remove-MailboxPermission", "Set-MailboxJunkEmailConfiguration", "New-TransportRule", "Remove-TransportRule", "Add-RecipientPermission", "Set-OwaMailboxPolicy")
        $exEvents = Get-O365AuditRecordsParallel -Token $script:O365Token -ContentType "Audit.Exchange" -Start $StartDate -End $EndDate -Throttle $ThrottleLimit -FilterOperations $exOps
        
        $rows = @()
        foreach ($ex in $exEvents) {
            $key = Get-DedupKey -Domain "Exchange" -Item $ex
            if ($existingKeys.Contains($key)) { $skipped++; continue }
            $rows += [PSCustomObject]@{
                FindingId = Get-FindingId -Prefix "EX"
                Time = $ex.CreationTime; Workload = $ex.Workload; Operation = $ex.Operation; User = $ex.UserId
                Item = $ex.ObjectId; ClientApp = $ex.ClientInfoString; IP = $ex.ClientIP; Result = $ex.ResultStatus
                'RiskLevel(1-10)' = Get-RiskScore -EventType "Exchange" -Data $ex; DedupKey = $key
                RawJson = ($ex | ConvertTo-Json -Depth 10 -Compress -WarningAction SilentlyContinue)
            }
            $inserted++
        }
        if ($rows.Count -gt 0 -and $PSCmdlet.ShouldProcess("Exchange", "Append")) { $rows | Export-Excel -Path $ReportPath -WorksheetName $SheetNames.Exchange -Append }
        $domainResults += "Exchange"
    }
    catch { Write-Log "Exchange failed: $_" -Level ERROR }
    $domainStats["Exchange"] = @{ Inserted = $inserted; Skipped = $skipped }
    Write-Host "  Inserted: $inserted | Skipped: $skipped" -ForegroundColor Gray

    # ============ 4. SHAREPOINT ============
    Write-Host "[4/11] SharePoint/OneDrive/Teams..." -ForegroundColor Cyan
    $existingKeys = Build-ExistingKeyIndex -Path $ReportPath -SheetName $SheetNames.SharePoint
    if ($null -eq $existingKeys) { $existingKeys = [System.Collections.Generic.HashSet[string]]::new() }
    $inserted = 0; $skipped = 0
    try {
        Refresh-TokensIfNeeded
        # Filter during download to save memory
        $spOps = @("SharingSet", "AnonymousLinkCreated", "SharingInvitationCreated", "FileDeleted", "PermissionLevelAdded", "PermissionLevelModified", "SiteCollectionAdminAdded", "ExternalUserAdded", "FileVersionsAllDeleted", "AnonymousLinkUsed")
        $spEvents = Get-O365AuditRecordsParallel -Token $script:O365Token -ContentType "Audit.SharePoint" -Start $StartDate -End $EndDate -Throttle $ThrottleLimit -FilterOperations $spOps
        
        $rows = @()
        foreach ($sp in $spEvents) {
            $key = Get-DedupKey -Domain "SharePoint" -Item $sp
            if ($existingKeys.Contains($key)) { $skipped++; continue }
            $rows += [PSCustomObject]@{
                FindingId = Get-FindingId -Prefix "SP"
                Time = $sp.CreationTime; Workload = $sp.Workload; Operation = $sp.Operation; User = $sp.UserId
                SiteOrTeam = $sp.SiteUrl; TargetItem = $sp.ObjectId; SharingType = $sp.TargetUserOrGroupType
                Result = $sp.ResultStatus; 'RiskLevel(1-10)' = Get-RiskScore -EventType "SharePoint" -Data $sp; DedupKey = $key
            }
            $inserted++
        }
        if ($rows.Count -gt 0 -and $PSCmdlet.ShouldProcess("SharePoint", "Append")) { $rows | Export-Excel -Path $ReportPath -WorksheetName $SheetNames.SharePoint -Append }
        $domainResults += "SharePoint"
    }
    catch { Write-Log "SharePoint failed: $_" -Level ERROR }
    $domainStats["SharePoint"] = @{ Inserted = $inserted; Skipped = $skipped }
    Write-Host "  Inserted: $inserted | Skipped: $skipped" -ForegroundColor Gray

    # ============ 7. ENDPOINT (MDE) — graceful skip ============
    Write-Host "[7/11] Endpoint (MDE)..." -ForegroundColor Cyan
    $inserted = 0; $skipped = 0
    if (-not $SkipDefenderEndpoint -and $ActiveDefenderRegion) {
        $existingKeys = Build-ExistingKeyIndex -Path $ReportPath -SheetName $SheetNames.Endpoint
        if ($null -eq $existingKeys) { $existingKeys = [System.Collections.Generic.HashSet[string]]::new() }
        try {
            $recs = Invoke-RetryableRequest -Uri "https://$ActiveDefenderRegion/api/recommendations" -Headers @{Authorization = "Bearer $($script:DefenderToken)" }
            $rows = @()
            foreach ($r in $recs.value) {
                $key = Get-DedupKey -Domain "Endpoint" -Item $r
                if ($existingKeys.Contains($key)) { $skipped++; continue }
                $rows += [PSCustomObject]@{
                    FindingId = Get-FindingId -Prefix "EP"
                    Time = $r.publishedOn; Product = $r.productName; RecommendationId = $r.id
                    RecommendationName = $r.recommendationName; SeverityScore = $r.severityScore
                    CVE = ($r.relatedCves | Select-Object -First 1); CVSS = $r.cvssScore
                    ExposedMachines = $r.exposedMachinesCount; RemediationType = $r.remediationType
                    Status = $r.status; DedupKey = $key
                }
                $inserted++
            }
            if ($rows.Count -gt 0 -and $PSCmdlet.ShouldProcess("Endpoint", "Append")) { $rows | Export-Excel -Path $ReportPath -WorksheetName $SheetNames.Endpoint -Append }
            $domainResults += "Endpoint"
        }
        catch { Write-Log "Endpoint failed: $_" -Level ERROR }
    }
    else {
        Write-Host "  Defender for Endpoint not licensed — skipping" -ForegroundColor DarkGray
    }
    $domainStats["Endpoint"] = @{ Inserted = $inserted; Skipped = $skipped }
    Write-Host "  Inserted: $inserted | Skipped: $skipped" -ForegroundColor Gray

    # ============ 8. DATA PROTECTION (DLP) — fixed pipe ============
    Write-Host "[8/11] Data Protection (DLP)..." -ForegroundColor Cyan
    $existingKeys = Build-ExistingKeyIndex -Path $ReportPath -SheetName $SheetNames.DataProtect
    if ($null -eq $existingKeys) { $existingKeys = [System.Collections.Generic.HashSet[string]]::new() }
    $inserted = 0; $skipped = 0
    try {
        Refresh-TokensIfNeeded
        $dlpEvents = Get-O365AuditRecordsParallel -Token $script:O365Token -ContentType "DLP.All" -Start $StartDate -End $EndDate -Throttle $ThrottleLimit
        if ($null -eq $dlpEvents) { $dlpEvents = @() }
        if ($dlpEvents -isnot [array]) { $dlpEvents = @($dlpEvents) }
        
        $rows = @()
        foreach ($d in $dlpEvents) {
            $key = Get-DedupKey -Domain "DataProtect" -Item $d
            if ($existingKeys.Contains($key)) { $skipped++; continue }
            
            $policyName = ""
            if ($d.PolicyDetails -and $d.PolicyDetails.Count -gt 0) {
                $policyName = $d.PolicyDetails[0].PolicyName
            }
            elseif ($d.PolicyId) { $policyName = $d.PolicyId }
            
            $action = ""
            if ($d.Actions -and $d.Actions.Count -gt 0) {
                $action = ($d.Actions | Select-Object -First 3) -join ", "
            }
            
            $sev = switch ($d.Severity) {
                "High" { 9 }
                "Medium" { 6 }
                "Low" { 3 }
                default { 4 }
            }
            
            $rows += [PSCustomObject]@{
                FindingId = Get-FindingId -Prefix "DP"
                Time = $d.CreationTime; Workload = $d.Workload; Operation = $d.Operation
                User = $d.UserId; Policy = $policyName; Target = $d.ObjectId
                Action = $action; Result = $d.ResultStatus
                Severity = $sev; DedupKey = $key
            }
            $inserted++
        }
        if ($rows.Count -gt 0 -and $PSCmdlet.ShouldProcess("DataProtect", "Append")) { $rows | Export-Excel -Path $ReportPath -WorksheetName $SheetNames.DataProtect -Append }
        $domainResults += "DataProtect"
    }
    catch { Write-Log "DataProtect (DLP) failed: $_" -Level ERROR }
    $domainStats["DataProtect"] = @{ Inserted = $inserted; Skipped = $skipped }
    Write-Host "  Inserted: $inserted | Skipped: $skipped" -ForegroundColor Gray

    # ============ 9. SECURITY ALERTS ============
    Write-Host "[9/11] Security Alerts..." -ForegroundColor Cyan
    $existingKeys = Build-ExistingKeyIndex -Path $ReportPath -SheetName $SheetNames.Alerts
    if ($null -eq $existingKeys) { $existingKeys = [System.Collections.Generic.HashSet[string]]::new() }
    $inserted = 0; $skipped = 0
    try {
        $alerts = Invoke-GraphWithPagination -Uri "https://graph.microsoft.com/v1.0/security/alerts_v2?`$filter=createdDateTime ge $startIso&`$top=500" -Token $script:GraphToken -Label "Alerts"
        $rows = @()
        foreach ($a in $alerts) {
            $key = Get-DedupKey -Domain "Alerts" -Item $a
            if ($existingKeys.Contains($key)) { $skipped++; continue }
            $sev = switch ($a.severity) { "high" { 9 } "medium" { 6 } "low" { 3 } default { 1 } }
            $rows += [PSCustomObject]@{
                FindingId = Get-FindingId -Prefix "AL"
                Time = $a.createdDateTime; Provider = $a.serviceSource; AlertId = $a.id; Title = $a.title
                Severity = $sev; Category = $a.category; Entity = ($a.evidence | Select-Object -First 1).displayName
                Status = $a.status; DedupKey = $key; RawJson = ($a | ConvertTo-Json -Depth 5 -Compress)
            }
            $inserted++
        }
        if ($rows.Count -gt 0 -and $PSCmdlet.ShouldProcess("Alerts", "Append")) { $rows | Export-Excel -Path $ReportPath -WorksheetName $SheetNames.Alerts -Append }
        $domainResults += "Alerts"
    }
    catch { Write-Log "Alerts failed: $_" -Level ERROR }
    $domainStats["Alerts"] = @{ Inserted = $inserted; Skipped = $skipped }
    Write-Host "  Inserted: $inserted | Skipped: $skipped" -ForegroundColor Gray

    # ============ 10. PRIVILEGED ACCESS (role assignments + PIM) ============
    Write-Host "[10/11] Privileged Access..." -ForegroundColor Cyan
    $existingKeys = Build-ExistingKeyIndex -Path $ReportPath -SheetName $SheetNames.PrivAccess
    if ($null -eq $existingKeys) { $existingKeys = [System.Collections.Generic.HashSet[string]]::new() }
    $inserted = 0; $skipped = 0
    try {
        # Part A: Role assignment changes from directory audits
        $roleEvents = $allDirAudits | Where-Object { $_.activityDisplayName -in $privAccessOps }
        
        $rows = @()
        foreach ($r in $roleEvents) {
            $key = Get-DedupKey -Domain "PrivAccess" -Item $r
            if ($existingKeys.Contains($key)) { $skipped++; continue }
            $init = Get-AuditActor -Audit $r
            $target = if ($r.targetResources.Count) { $r.targetResources[0].displayName } else { "" }
            $role = ""
            if ($r.targetResources.Count -gt 0 -and $r.targetResources[0].modifiedProperties) {
                $roleProp = $r.targetResources[0].modifiedProperties | Where-Object { $_.displayName -match "Role" } | Select-Object -First 1
                if ($roleProp) { $role = $roleProp.newValue }
            }
            if (-not $role -and $r.targetResources.Count -gt 1) { $role = $r.targetResources[1].displayName }
            
            # Risk: adding members to privileged roles is high risk
            $risk = if ($r.activityDisplayName -match "Add.*member") { 8 } else { 6 }
            
            $rows += [PSCustomObject]@{
                FindingId = Get-FindingId -Prefix "PA"
                Time = $r.activityDateTime; Operation = $r.activityDisplayName
                Actor = $init; RoleOrScope = $role; Target = $target
                Justification = ""; Result = $r.result; 'RiskLevel(1-10)' = $risk
                DedupKey = $key; RawJson = ($r | ConvertTo-Json -Depth 5 -Compress)
            }
            $inserted++
        }
        
        # Part B: PIM activations (just-in-time role activations)
        try {
            Refresh-TokensIfNeeded
            $pimActivations = Invoke-GraphWithPagination -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?`$filter=startDateTime ge $startIso" -Token $script:GraphToken -Label "PIM Activations"
            foreach ($p in $pimActivations) {
                $pKey = Get-DedupKey -Domain "PrivAccess" -Item $p
                if ($existingKeys.Contains($pKey)) { $skipped++; continue }
                $rows += [PSCustomObject]@{
                    FindingId = Get-FindingId -Prefix "PA"
                    Time = $p.startDateTime; Operation = "PIM Role Activation"
                    Actor = $p.principal.displayName; RoleOrScope = $p.roleDefinition.displayName
                    Target = $p.principal.userPrincipalName
                    Justification = if ($p.assignmentType -eq "Activated") { "JIT Activation" } else { $p.assignmentType }
                    Result = "Active"; 'RiskLevel(1-10)' = 5
                    DedupKey = $pKey; RawJson = ($p | ConvertTo-Json -Depth 5 -Compress)
                }
                $inserted++
            }
        }
        catch { Write-Log "PIM data not available (may require RoleManagement.Read.Directory): $_" -Level DEBUG }
        
        if ($rows.Count -gt 0 -and $PSCmdlet.ShouldProcess("PrivAccess", "Append")) { $rows | Export-Excel -Path $ReportPath -WorksheetName $SheetNames.PrivAccess -Append }
        $domainResults += "PrivAccess"
    }
    catch { Write-Log "PrivAccess failed: $_" -Level ERROR }
    $domainStats["PrivAccess"] = @{ Inserted = $inserted; Skipped = $skipped }
    Write-Host "  Inserted: $inserted | Skipped: $skipped" -ForegroundColor Gray

    # ============ 11. MFA STATUS ============
    Write-Host "[11/11] MFA Registration Status..." -ForegroundColor Cyan
    $existingKeys = Build-ExistingKeyIndex -Path $ReportPath -SheetName $SheetNames.MFAStatus
    if ($null -eq $existingKeys) { $existingKeys = [System.Collections.Generic.HashSet[string]]::new() }
    $inserted = 0; $skipped = 0
    $script:MfaStats = @{ TotalUsers = 0; Registered = 0; NotRegistered = 0 }
    try {
        Refresh-TokensIfNeeded
        # Query user registration details for MFA status
        $mfaDetails = Invoke-GraphWithPagination -Uri "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?`$top=999" -Token $script:GraphToken -Label "MFA Registration"
        
        $rows = @()
        foreach ($user in $mfaDetails) {
            $script:MfaStats.TotalUsers++
            if ($user.isMfaRegistered -eq $true) {
                $script:MfaStats.Registered++
            }
            else {
                $script:MfaStats.NotRegistered++
            }
            
            # Only add to sheet if NOT registered for MFA (the security concern)
            if ($user.isMfaRegistered -ne $true) {
                $key = Get-DedupKey -Domain "MFAStatus" -Item $user
                if ($existingKeys.Contains($key)) { $skipped++; continue }
                
                $methods = if ($user.methodsRegistered) { ($user.methodsRegistered -join ", ") } else { "None" }
                $rows += [PSCustomObject]@{
                    UserPrincipalName = $user.userPrincipalName
                    DisplayName       = $user.userDisplayName
                    IsMfaRegistered   = $user.isMfaRegistered
                    IsMfaCapable      = $user.isMfaCapable
                    DefaultMfaMethod  = if ($user.defaultMfaMethod) { $user.defaultMfaMethod } else { "None" }
                    MethodsRegistered = $methods
                    AccountEnabled    = $user.isEnabled
                    UserType          = $user.userType
                    DedupKey          = $key
                }
                $inserted++
            }
        }
        if ($rows.Count -gt 0 -and $PSCmdlet.ShouldProcess("MFAStatus", "Append $($rows.Count) rows")) {
            $rows | Export-Excel -Path $ReportPath -WorksheetName $SheetNames.MFAStatus -Append
        }
        $domainResults += "MFAStatus"
        Write-Host "  Total Users: $($script:MfaStats.TotalUsers) | MFA Registered: $($script:MfaStats.Registered) | NOT Registered: $($script:MfaStats.NotRegistered)" -ForegroundColor Yellow
    }
    catch { 
        Write-Log "MFA Status failed: $_ (Requires UserAuthenticationMethod.Read.All permission)" -Level WARN
        $script:MfaStats = @{ TotalUsers = 0; Registered = 0; NotRegistered = 0 }
    }
    $domainStats["MFAStatus"] = @{ Inserted = $inserted; Skipped = $skipped }
    Write-Host "  Inserted: $inserted | Skipped: $skipped" -ForegroundColor Gray

    # ============ FINALIZE ============
    $totalInserted = ($domainStats.Values | ForEach-Object { $_.Inserted } | Measure-Object -Sum).Sum
    $totalSkipped = ($domainStats.Values | ForEach-Object { $_.Skipped } | Measure-Object -Sum).Sum
    
    if ($PSCmdlet.ShouldProcess("Config", "Update")) {
        Update-ConfigSheet -Path $ReportPath -Stats $domainStats -RunId $runId
        Add-RunRecord -Path $ReportPath -RunId $runId -Start $scriptStart -RangeStart $StartDate -RangeEnd $EndDate -Inserted $totalInserted -Skipped $totalSkipped -Status "Success" -Domains ($domainResults -join ",")
    }
    
    # Generate Summary sheet
    if ($PSCmdlet.ShouldProcess("Summary", "Generate")) {
        Update-SummarySheet -Path $ReportPath -Stats $domainStats -RangeStart $StartDate -RangeEnd $EndDate -RunId $runId
    }
    

    
    # JSON export
    if ($ExportJson) {
        $jsonDir = Join-Path $script:ScriptRoot "json-$targetMonth"
        if (-not (Test-Path $jsonDir)) { New-Item -ItemType Directory -Path $jsonDir | Out-Null }
        Write-Log "JSON export to $jsonDir" -Level INFO
    }
    
    # AI Context Summary (v2.5.1 - Executive Intelligence Brief)
    if (-not $NoAI) {
        $script:AIContextFileName = "AI-Context-Summary-$targetMonth.txt"
        $script:AIContextPath = Join-Path $script:ScriptRoot $script:AIContextFileName
        
        # ---- Compute aggregates for intelligence functions ----
        $grandCrit = 0; $grandHigh = 0; $grandMed = 0; $grandLow = 0; $grandTotal = 0; $grandNew = 0
        $inboxRuleCount = 0; $sendAsCount = 0; $anonLinkCount = 0; $extUserCount = 0
        $domainNarratives = @{}
        
        if ($null -ne $script:SummaryDomainRows) {
            foreach ($dr in $script:SummaryDomainRows) {
                $grandTotal += $dr.TotalEvents; $grandNew += $dr.NewThisRun
                $grandCrit += $dr.Critical; $grandHigh += $dr.High; $grandMed += $dr.Medium; $grandLow += $dr.Low
            }
        }
        
        # Count BEC and exposure events from raw domain data
        try {
            $exData = Import-Excel -Path $ReportPath -WorksheetName $SheetNames.Exchange -ErrorAction SilentlyContinue
            if ($null -ne $exData) {
                if ($exData -isnot [array]) { $exData = @($exData) }
                $inboxRuleCount = @($exData | Where-Object { $_.Operation -match "InboxRule" }).Count
                $sendAsCount = @($exData | Where-Object { $_.Operation -match "SendAs" }).Count
            }
        }
        catch { $exData = @() }
        
        try {
            $spData = Import-Excel -Path $ReportPath -WorksheetName $SheetNames.SharePoint -ErrorAction SilentlyContinue
            if ($null -ne $spData) {
                if ($spData -isnot [array]) { $spData = @($spData) }
                $anonLinkCount = @($spData | Where-Object { $_.Operation -match "AnonymousLink" }).Count
                $extUserCount = @($spData | Where-Object { $_.Operation -match "ExternalUserAdded" }).Count
            }
        }
        catch { $spData = @() }
        
        # Compute Month-over-Month deltas
        $momDeltas = Get-MonthOverMonthDeltas -Path $ReportPath `
            -CurrentInserted $totalInserted -CurrentSkipped $totalSkipped `
            -CurrentCritical $grandCrit -CurrentHigh $grandHigh `
            -CurrentMFAGap $(if ($script:SummaryMFAData) { $script:SummaryMFAData.NotRegistered } else { 0 }) `
            -CurrentExternalSharing ($anonLinkCount + $extUserCount)
        
        # Compute Security Posture Score
        $postureResult = Get-SecurityPostureScore `
            -MFAData $(if ($script:SummaryMFAData) { $script:SummaryMFAData } else { @{ Available = $false; Total = 0; Registered = 0; NotRegistered = 0 } }) `
            -TopFindings $(if ($script:SummaryFindingsRows) { $script:SummaryFindingsRows } else { @() }) `
            -ActionsList $(if ($script:SummaryActionsList) { $script:SummaryActionsList } else { @() }) `
            -MoMDeltas $momDeltas `
            -InboxRuleCount $inboxRuleCount -SendAsCount $sendAsCount `
            -AnonLinkCount $anonLinkCount -ExtUserCount $extUserCount
        
        $script:PostureScore = $postureResult  # Store for Excel tables
        
        # Generate per-domain narratives
        $domainSheetMap = @{
            "Exchange"    = @{ Sheet = $SheetNames.Exchange; Key = "Exchange" }
            "SharePoint"  = @{ Sheet = $SheetNames.SharePoint; Key = "SharePoint" }
            "Identity"    = @{ Sheet = $SheetNames.Identity; Key = "Identity" }
            "Endpoint"    = @{ Sheet = $SheetNames.Endpoint; Key = "Endpoint" }
            "Alerts"      = @{ Sheet = $SheetNames.Alerts; Key = "Alerts" }
            "AppConsents" = @{ Sheet = $SheetNames.AppConsents; Key = "AppConsents" }
            "PrivAccess"  = @{ Sheet = $SheetNames.PrivAccess; Key = "PrivAccess" }
            "CondAccess"  = @{ Sheet = $SheetNames.CondAccess; Key = "CondAccess" }
        }
        
        foreach ($dn in $domainSheetMap.Keys) {
            $dInfo = $domainSheetMap[$dn]
            $dEvents = @()
            try {
                $dEvents = Import-Excel -Path $ReportPath -WorksheetName $dInfo.Sheet -ErrorAction SilentlyContinue
                if ($null -eq $dEvents) { $dEvents = @() }
                if ($dEvents -isnot [array]) { $dEvents = @($dEvents) }
            }
            catch { }
            
            $dRow = $script:SummaryDomainRows | Where-Object { $_.Domain -eq $dn } | Select-Object -First 1
            $dCrit = if ($dRow) { $dRow.Critical } else { 0 }
            $dHigh = if ($dRow) { $dRow.High } else { 0 }
            $dMed = if ($dRow) { $dRow.Medium } else { 0 }
            $dLow = if ($dRow) { $dRow.Low } else { 0 }
            
            # Find top actor
            $topActor = ""
            if ($dEvents.Count -gt 0) {
                $actorCol = switch ($dn) {
                    "Exchange" { "User" }
                    "SharePoint" { "User" }
                    "Identity" { "Actor" }
                    "PrivAccess" { "Actor" }
                    "AppConsents" { "InitiatedBy" }
                    "CondAccess" { "ModifiedBy" }
                    default { "" }
                }
                if ($actorCol -and $dEvents[0].PSObject.Properties[$actorCol]) {
                    $topActorObj = $dEvents | Group-Object -Property $actorCol -ErrorAction SilentlyContinue | Sort-Object Count -Descending | Select-Object -First 1
                    if ($topActorObj) { $topActor = "$($topActorObj.Name) ($($topActorObj.Count) events)" }
                }
            }
            
            $domainNarratives[$dn] = Get-DomainNarrative -DomainName $dn -Events $dEvents `
                -Critical $dCrit -High $dHigh -Medium $dMed -Low $dLow -TopActor $topActor
        }
        
        # ---- Build THIS RUN's section (will be prepended to historical sections) ----
        $runSb = [System.Text.StringBuilder]::new()
        $runId = $script:RunId
        $runTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
        
        [void]$runSb.AppendLine("")
        [void]$runSb.AppendLine([string]::new("=", 80))
        [void]$runSb.AppendLine("  RUN $runId | $runTimestamp | v$($script:Version)")
        [void]$runSb.AppendLine("  Period: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))")
        [void]$runSb.AppendLine("  Posture: $($postureResult.Score)/100 $($postureResult.Emoji) $($postureResult.Rating.ToUpper())")
        [void]$runSb.AppendLine("  Inserted: $totalInserted | Deduplicated: $totalSkipped")
        [void]$runSb.AppendLine([string]::new("=", 80))
        [void]$runSb.AppendLine("")
        
        # Key Risks
        [void]$runSb.AppendLine("  KEY RISKS:")
        if ($null -ne $script:SummaryFindingsRows -and $script:SummaryFindingsRows.Count -gt 0) {
            $riskNum = 1
            foreach ($finding in $script:SummaryFindingsRows) {
                if ($finding.Severity -ge 9) { $sev = "CRITICAL" } elseif ($finding.Severity -ge 7) { $sev = "HIGH" } else { continue }
                [void]$runSb.AppendLine("    RISK-$riskNum`: [$sev] $($finding.Description) ($($finding.Domain))")
                $riskNum++
            }
            if ($script:SummaryMFAData -and $script:SummaryMFAData.Available -and $script:SummaryMFAData.NotRegistered -gt 0) {
                [void]$runSb.AppendLine("    RISK-$riskNum`: [HIGH] MFA Gap - $($script:SummaryMFAData.NotRegistered) account(s) without MFA")
            }
        }
        else {
            [void]$runSb.AppendLine("    No critical or high-severity findings.")
        }
        [void]$runSb.AppendLine("")
        
        # Changes vs Last Run
        if ($momDeltas.HasPreviousMonth) {
            [void]$runSb.AppendLine("  CHANGES vs LAST RUN:")
            $changes = @()
            if ($momDeltas.DeltaInserted -lt 0) { $changes += "Findings decreased: $($momDeltas.PrevInserted) -> $totalInserted ($($momDeltas.DeltaPct))" }
            if ($momDeltas.DeltaInserted -gt 0) { $changes += "Findings increased: $($momDeltas.PrevInserted) -> $totalInserted ($($momDeltas.DeltaPct))" }
            if ($grandCrit -eq 0) { $changes += "No critical findings detected" }
            if ($grandCrit -gt 0) { $changes += "Critical findings: $grandCrit detected" }
            if ($changes.Count -eq 0) { $changes += "No significant changes" }
            foreach ($ch in $changes) { [void]$runSb.AppendLine("    $ch") }
            [void]$runSb.AppendLine("")
        }
        
        # Posture Score Breakdown
        [void]$runSb.AppendLine("  POSTURE SCORE BREAKDOWN:")
        [void]$runSb.AppendLine([string]::Format("    {0,-26} {1,6} {2,8} {3,14} {4}", "Factor", "Score", "Weight", "Contribution", "Detail"))
        [void]$runSb.AppendLine("    " + [string]::new("-", 80))
        foreach ($f in $postureResult.Factors) {
            $wPct = "$([math]::Round($f.Weight * 100))%"
            [void]$runSb.AppendLine([string]::Format("    {0,-26} {1,6} {2,8} {3,14} {4}", $f.Name, $f.Score, $wPct, $f.Contribution, $f.Detail))
        }
        [void]$runSb.AppendLine("    " + [string]::new("-", 80))
        [void]$runSb.AppendLine([string]::Format("    {0,-26} {1,6} {2,8} {3,14}", "TOTAL", $postureResult.Score, "100%", $postureResult.Score))
        [void]$runSb.AppendLine("")
        
        # Domain Analysis with Narratives
        [void]$runSb.AppendLine("  DOMAIN ANALYSIS:")
        if ($null -ne $script:SummaryDomainRows) {
            foreach ($dr in $script:SummaryDomainRows) {
                $narr = if ($domainNarratives.ContainsKey($dr.Domain)) { $domainNarratives[$dr.Domain] } else { @{ RiskEmoji = ""; RiskLevel = "UNKNOWN"; Narrative = "" } }
                [void]$runSb.AppendLine("    $($dr.Domain) — $($narr.RiskEmoji) $($narr.RiskLevel)")
                [void]$runSb.AppendLine("    $($dr.TotalEvents) events | Crit: $($dr.Critical), High: $($dr.High), Med: $($dr.Medium), Low: $($dr.Low)")
                if ($narr.Narrative) { [void]$runSb.AppendLine("    $($narr.Narrative)") }
                [void]$runSb.AppendLine("")
            }
        }
        
        # Domain Statistics Table
        [void]$runSb.AppendLine("  DOMAIN STATISTICS:")
        [void]$runSb.AppendLine([string]::Format("    {0,-22} {1,8} {2,6} {3,6} {4,8} {5,6} {6,6} {7,6}", "Domain", "Total", "New", "Dedup", "Critical", "High", "Medium", "Low"))
        [void]$runSb.AppendLine("    " + [string]::new("-", 76))
        if ($null -ne $script:SummaryDomainRows) {
            foreach ($dr in $script:SummaryDomainRows) {
                [void]$runSb.AppendLine([string]::Format("    {0,-22} {1,8} {2,6} {3,6} {4,8} {5,6} {6,6} {7,6}", $dr.Domain, $dr.TotalEvents, $dr.NewThisRun, $dr.Deduplicated, $dr.Critical, $dr.High, $dr.Medium, $dr.Low))
            }
            [void]$runSb.AppendLine("    " + [string]::new("-", 76))
            [void]$runSb.AppendLine([string]::Format("    {0,-22} {1,8} {2,6} {3,6} {4,8} {5,6} {6,6} {7,6}", "TOTALS", $grandTotal, $grandNew, $totalSkipped, $grandCrit, $grandHigh, $grandMed, $grandLow))
        }
        [void]$runSb.AppendLine("")
        
        # Top Findings
        [void]$runSb.AppendLine("  TOP FINDINGS:")
        if ($null -ne $script:SummaryFindingsRows -and $script:SummaryFindingsRows.Count -gt 0 -and $script:SummaryFindingsRows[0].Severity -gt 0) {
            foreach ($finding in $script:SummaryFindingsRows) {
                [void]$runSb.AppendLine("    $($finding.Rank). $($finding.Domain) | Sev: $($finding.Severity)/10 | $($finding.Description)")
                [void]$runSb.AppendLine("       Actor: $($finding.Actor) | Time: $($finding.Time)")
            }
        }
        else {
            [void]$runSb.AppendLine("    No critical or high-severity findings detected.")
        }
        [void]$runSb.AppendLine("")
        
        # MFA Status (compact)
        if ($null -ne $script:SummaryMFAData -and $script:SummaryMFAData.Available) {
            $mfTotal = $script:SummaryMFAData.Total
            $mfReg = $script:SummaryMFAData.Registered
            $mfNot = $script:SummaryMFAData.NotRegistered
            $mfPct = if ($mfTotal -gt 0) { [math]::Round([double](($mfReg / $mfTotal) * 100), 1) } else { 0 }
            [void]$runSb.AppendLine("  MFA: $mfReg/$mfTotal registered ($mfPct%) | $mfNot NOT registered")
        }
        [void]$runSb.AppendLine("")
        
        # Recommended Actions (compact)
        [void]$runSb.AppendLine("  ACTIONS:")
        if ($null -ne $script:SummaryActionsList -and $script:SummaryActionsList.Count -gt 0) {
            foreach ($act in $script:SummaryActionsList) {
                $urgency = if ($act.Priority -le 2) { "URGENT" } elseif ($act.Priority -le 4) { "HIGH" } else { "MEDIUM" }
                [void]$runSb.AppendLine("    $($act.Priority). [$urgency] [$($act.Category)] $($act.Action) ($($act.Count))")
            }
        }
        else {
            [void]$runSb.AppendLine("    No critical action items.")
        }
        [void]$runSb.AppendLine("")
        
        # ---- CUMULATIVE FILE ASSEMBLY ----
        # Read existing run sections from prior file
        $previousRunSections = ""
        $runDelimiter = [string]::new("=", 80)
        if (Test-Path $script:AIContextPath) {
            $existingContent = Get-Content $script:AIContextPath -Raw -ErrorAction SilentlyContinue
            if ($existingContent) {
                # Find the first per-run section (starts with ====...==== line followed by "  RUN ")
                $pattern = "(?m)^={80}\r?\n  RUN "
                if ($existingContent -match $pattern) {
                    $firstRunMatch = [regex]::Match($existingContent, $pattern)
                    if ($firstRunMatch.Success) {
                        $previousRunSections = $existingContent.Substring($firstRunMatch.Index)
                        # Remove any trailing END OF BRIEF markers from previous content
                        $endPattern = "(?m)^=+\r?\n  END OF EXECUTIVE INTELLIGENCE BRIEF\r?\n=+\r?\n?"
                        $previousRunSections = [regex]::Replace($previousRunSections, $endPattern, "")
                    }
                }
            }
        }
        
        # Build the ROLLING HEADER (always reflects latest run)
        $headerSb = [System.Text.StringBuilder]::new()
        [void]$headerSb.AppendLine("================================================================================")
        [void]$headerSb.AppendLine("  M365 SECURITY REPORT — EXECUTIVE INTELLIGENCE BRIEF (CUMULATIVE)")
        [void]$headerSb.AppendLine("================================================================================")
        [void]$headerSb.AppendLine("")
        [void]$headerSb.AppendLine("LATEST POSTURE: $($postureResult.Score)/100 $($postureResult.Emoji) $($postureResult.Rating.ToUpper())")
        if ($momDeltas.HasPreviousMonth) {
            $trendArrow = switch ($momDeltas.Trend) { "Improving" { "^" } "Worsening" { "v" } default { "=" } }
            [void]$headerSb.AppendLine("Trend: $trendArrow $($momDeltas.Trend)")
        }
        [void]$headerSb.AppendLine("Month: $targetMonth | Last Updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
        [void]$headerSb.AppendLine("Script Version: $($script:Version)")
        [void]$headerSb.AppendLine("")
        
        # Run Trend table in header
        [void]$headerSb.AppendLine("================================================================================")
        [void]$headerSb.AppendLine("  HISTORICAL RUN TREND (Last 10 Runs)")
        [void]$headerSb.AppendLine("================================================================================")
        [void]$headerSb.AppendLine("")
        try {
            $runsData = Import-Excel -Path $ReportPath -WorksheetName $SheetNames.Runs -ErrorAction SilentlyContinue
            if ($null -ne $runsData) {
                if ($runsData -isnot [array]) { $runsData = @($runsData) }
                if ($runsData.Count -gt 0) {
                    [void]$headerSb.AppendLine([string]::Format("  {0,-12} {1,-12} {2,-12} {3,10} {4,10} {5,8}", "Run ID", "Start", "End", "Inserted", "Skipped", "Version"))
                    [void]$headerSb.AppendLine("  " + [string]::new("-", 70))
                    $recentRuns = $runsData | Select-Object -Last 10
                    foreach ($run in $recentRuns) {
                        [void]$headerSb.AppendLine([string]::Format("  {0,-12} {1,-12} {2,-12} {3,10} {4,10} {5,8}", 
                                $run.RunId, $run.DateRangeStart, $run.DateRangeEnd, $run.TotalInserted, $run.TotalSkipped, "v$($run.Version)"))
                    }
                }
            }
        }
        catch { [void]$headerSb.AppendLine("  Historical data not available.") }
        [void]$headerSb.AppendLine("")
        [void]$headerSb.AppendLine("================================================================================")
        [void]$headerSb.AppendLine("  PER-RUN DETAIL SECTIONS (newest first)")
        [void]$headerSb.AppendLine("================================================================================")
        
        # Assemble: Header + Current Run + Previous Runs + Footer
        $finalSb = [System.Text.StringBuilder]::new()
        [void]$finalSb.Append($headerSb.ToString())
        [void]$finalSb.Append($runSb.ToString())
        if ($previousRunSections) {
            [void]$finalSb.AppendLine("")
            [void]$finalSb.Append($previousRunSections.TrimEnd())
            [void]$finalSb.AppendLine("")
        }
        [void]$finalSb.AppendLine("")
        [void]$finalSb.AppendLine("================================================================================")
        [void]$finalSb.AppendLine("  END OF EXECUTIVE INTELLIGENCE BRIEF")
        [void]$finalSb.AppendLine("================================================================================")
        
        $finalSb.ToString() | Out-File $script:AIContextPath -Encoding UTF8
        Write-Host "`nAI Context: $script:AIContextPath" -ForegroundColor Gray
        
        # ==================================================================================
        # AI EXECUTIVE SUMMARY (v2.5.2)
        # ==================================================================================
        if (-not $SkipExecutiveSummary) {
            $hasAIKey = (-not [string]::IsNullOrEmpty($OpenAIKey)) -or (-not [string]::IsNullOrEmpty($AzureOpenAIEndpoint))
            if ($hasAIKey) {
                if ($PSCmdlet.ShouldProcess("ExecutiveSummary", "Generate AI executive summary")) {
                    Update-ExecutiveSummary -Path $ReportPath -Stats $domainStats `
                        -RangeStart $StartDate -RangeEnd $EndDate -RunId $runId
                }
            }
            else {
                Write-Log "AI executive summary skipped: no OpenAIKey or AzureOpenAIEndpoint provided" -Level DEBUG
            }
        }
    }
    
    # ==================================================================================
    # SHAREPOINT UPLOAD (v2.3.1 - retry logic with force check-in for locked files)
    # ==================================================================================
    if ($script:UseSharePoint -and $script:SharePointSiteId) {
        Write-Host "`n[SharePoint] Uploading artifacts to cloud storage..." -ForegroundColor Cyan
        
        # Helper function for upload with retry logic
        function Invoke-SharePointUploadWithRetry {
            param(
                [string]$FileName,
                [string]$LocalPath,
                [string]$DisplayName,
                [int]$MaxRetries = 3,
                [int]$RetryDelaySeconds = 15
            )
            
            $uploaded = $false
            $retryCount = 0
            
            while (-not $uploaded -and $retryCount -le $MaxRetries) {
                try {
                    Set-SharePointFile -SiteId $script:SharePointSiteId -FolderPath $SharePointFolder -FileName $FileName -LocalPath $LocalPath -Token $script:GraphToken
                    $uploaded = $true
                    Write-Host "  [OK] ${DisplayName}: $FileName" -ForegroundColor Green
                }
                catch {
                    $errorMsg = $_.ToString()
                    
                    # Check if it's a resourceLocked error
                    if ($errorMsg -match "resourceLocked" -or $errorMsg -match "locked") {
                        $retryCount++
                        
                        if ($retryCount -le $MaxRetries) {
                            # Phase 1: Polite retry with delay
                            Write-Host "  [--] $DisplayName locked, retrying in ${RetryDelaySeconds}s ($retryCount/$MaxRetries)..." -ForegroundColor Yellow
                            Start-Sleep -Seconds $RetryDelaySeconds
                        }
                        else {
                            # Phase 2: Force check-in
                            Write-Host "  [!!] $DisplayName still locked after $MaxRetries retries, attempting force check-in..." -ForegroundColor Yellow
                            
                            try {
                                $itemId = Get-SharePointItemId -SiteId $script:SharePointSiteId -FolderPath $SharePointFolder -FileName $FileName -Token $script:GraphToken
                                
                                if ($itemId) {
                                    $unlocked = Unlock-SharePointFile -SiteId $script:SharePointSiteId -ItemId $itemId -Token $script:GraphToken
                                    
                                    if ($unlocked) {
                                        Write-Host "  [OK] Lock released, retrying upload..." -ForegroundColor Cyan
                                        Start-Sleep -Seconds 2  # Brief pause after unlock
                                        
                                        # Final attempt after force unlock
                                        try {
                                            Set-SharePointFile -SiteId $script:SharePointSiteId -FolderPath $SharePointFolder -FileName $FileName -LocalPath $LocalPath -Token $script:GraphToken
                                            $uploaded = $true
                                            Write-Host "  [OK] ${DisplayName}: $FileName" -ForegroundColor Green
                                        }
                                        catch {
                                            Write-Log "Final upload attempt failed for $FileName`: $_" -Level ERROR
                                            Write-Host "  [!!] $DisplayName upload failed after unlock" -ForegroundColor Red
                                        }
                                    }
                                    else {
                                        # Lock could not be released - fallback: upload with timestamped name
                                        Write-Host "  [--] Lock persists, uploading with timestamped filename..." -ForegroundColor Yellow
                                        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
                                        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
                                        $extension = [System.IO.Path]::GetExtension($FileName)
                                        $fallbackName = "${baseName}-${timestamp}${extension}"
                                        
                                        try {
                                            Set-SharePointFile -SiteId $script:SharePointSiteId -FolderPath $SharePointFolder -FileName $fallbackName -LocalPath $LocalPath -Token $script:GraphToken
                                            $uploaded = $true
                                            Write-Host "  [OK] ${DisplayName}: $fallbackName (timestamped)" -ForegroundColor Green
                                            Write-Log "$DisplayName uploaded as $fallbackName due to lock on original" -Level INFO
                                        }
                                        catch {
                                            Write-Log "Fallback upload failed for $fallbackName`: $_" -Level ERROR
                                            Write-Host "  [!!] Fallback upload also failed" -ForegroundColor Red
                                        }
                                    }
                                }
                                else {
                                    # File doesn't exist in SharePoint yet but got lock error - shouldn't happen
                                    Write-Host "  [!!] Unexpected lock error for new file" -ForegroundColor Red
                                }
                            }
                            catch {
                                Write-Log "Force unlock failed for $FileName`: $_" -Level ERROR
                                Write-Host "  [!!] Force unlock failed for $DisplayName" -ForegroundColor Red
                            }
                        }
                    }
                    else {
                        # Non-lock error, log and break
                        Write-Log "Upload failed for $FileName`: $_" -Level ERROR
                        Write-Host "  [!!] $DisplayName upload failed: $_" -ForegroundColor Red
                        break
                    }
                }
            }
            
            return $uploaded
        }
        
        try {
            Refresh-TokensIfNeeded
            
            # 1. Upload Excel Report (most likely to be locked)
            $spFileName = "Security-Monthly-Report-$targetMonth.xlsx"
            $reportUploaded = Invoke-SharePointUploadWithRetry -FileName $spFileName -LocalPath $ReportPath -DisplayName "Report"
            
            # 2. Upload Log File
            if ($LogPath -and (Test-Path $LogPath)) {
                try { Stop-Transcript | Out-Null } catch { }  # Stop transcript before uploading log
                Invoke-SharePointUploadWithRetry -FileName $script:LogFileName -LocalPath $LogPath -DisplayName "Log" | Out-Null
            }
            
            # 3. Upload AI Context Summary
            if ($script:AIContextPath -and (Test-Path $script:AIContextPath)) {
                Invoke-SharePointUploadWithRetry -FileName $script:AIContextFileName -LocalPath $script:AIContextPath -DisplayName "AI Context" | Out-Null
            }
            
            # 4. Upload Executive Summary (v2.5.1)
            if ($script:ExecutiveSummaryFilePath -and (Test-Path $script:ExecutiveSummaryFilePath)) {
                Invoke-SharePointUploadWithRetry -FileName $script:ExecutiveSummaryFileName -LocalPath $script:ExecutiveSummaryFilePath -DisplayName "Executive Summary" | Out-Null
            }
            
            # Clean up temp files
            if ($script:LocalTempPath -and (Test-Path $script:LocalTempPath)) {
                Remove-Item $script:LocalTempPath -Force -ErrorAction SilentlyContinue
                Write-Host "  [OK] Cleaned up local temp file" -ForegroundColor DarkGray
            }
            
            if (-not $reportUploaded) {
                Write-Host "  [!!] Report upload incomplete - local file preserved at $ReportPath" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Log "SharePoint upload failed: $_ - local files preserved" -Level ERROR
            Write-Host "  [!!] Upload failed - local files preserved" -ForegroundColor Yellow
        }
    }

    
    Write-Host "`n==========================================================================" -ForegroundColor Green
    Write-Host "  COMPLETE! Run ID: $runId" -ForegroundColor Green
    Write-Host "==========================================================================" -ForegroundColor Green
    if ($script:UseSharePoint) {
        Write-Host "  SharePoint: $SharePointSiteUrl" -ForegroundColor Cyan
        Write-Host "  Folder:     $SharePointFolder" -ForegroundColor Cyan
    }
    Write-Host "  Report:    $ReportPath" -ForegroundColor Cyan
    Write-Host "  Inserted:  $totalInserted | Skipped: $totalSkipped" -ForegroundColor White
    Write-Host "  Duration:  $([math]::Round([double]((Get-Date) - $scriptStart).TotalMinutes, 1)) minutes" -ForegroundColor White
    
}
catch {
    Write-Host "`n[ERROR] $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    throw
}
finally {
    $ClientSecretPlain = $null
    [System.GC]::Collect()
    try { Stop-Transcript | Out-Null } catch { }
}
