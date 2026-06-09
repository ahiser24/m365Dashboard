<#
.SYNOPSIS
    Downloads Microsoft 365 usage reports (Copilot, Viva Engage, Teams, SharePoint, OneDrive, Exchange) via Microsoft Graph.

.DESCRIPTION
    - Uses app-only authentication (client secret)
    - Exports CSV (and JSON where applicable)
    - Deterministic folder + filename structure
    - Teams: runs ALL available Teams usage report cmdlets found in Microsoft.Graph.Reports
    - Copilot: runs your known beta endpoints + attempts any additional Copilot beta report cmdlets if present

.REQUIRES
    Microsoft.Graph.Authentication
    Microsoft.Graph.Reports
    Microsoft.Graph.Beta.Reports (optional but recommended for extra Copilot cmdlets)

.PERMISSIONS (Application)
    Reports.Read.All
#>

[CmdletBinding()]
param(
    # App registration details
    [string]$TenantId   = "1273caf7-13b7-4a89-b44a-3967d45ba0a9",
    [string]$ClientId   = "e458a67c-2643-4440-8ed0-01e01cb4b697",
    [string]$SecretPath = "$env:USERPROFILE\graph-client-secret.xml",

    # Report period (Graph-supported)
    [ValidateSet("D7","D30","D90","D180")]
    [string]$Period = "D30",

    # Output root
    [string]$OutputRoot = "C:\Reports",

    # Workload switches
    [switch]$Copilot,
    [switch]$VivaEngage,
    [switch]$Teams,
    [switch]$SharePoint,
    [switch]$OneDrive,
    [switch]$Exchange,

    # Granular SharePoint switches (optional)
    [switch]$SPSites,
    [switch]$SPFiles,
    [switch]$SPStorage,
    [switch]$SPPages,
    [switch]$SPActivity,

    # Granular OneDrive switches (optional)
    [switch]$ODUsage,
    [switch]$ODActivity,

    # Granular Exchange switches (optional)
    [switch]$EXEmailActivity,
    [switch]$EXEmailAppUsage,
    [switch]$EXMailboxUsage,

    # Run everything
    [switch]$All
)

#region Helper Functions

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

function Connect-GraphApp {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$SecretPath
    )

    if (-not (Test-Path $SecretPath)) {
        throw "Secret file not found: $SecretPath"
    }

    $secure = Import-Clixml $SecretPath
    $cred   = New-Object PSCredential ($ClientId, $secure)

    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $cred -NoWelcome | Out-Null
}

function Export-GraphReportCsv {
    <#
      Calls a Graph Reports endpoint that returns CSV and writes it to a file.
      Skips missing endpoints gracefully.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EndpointName,
        [Parameter(Mandatory)][string]$Period,
        [Parameter(Mandatory)][string]$OutFile
    )

    $uri = "https://graph.microsoft.com/v1.0/reports/$EndpointName(period='$Period')"

    try {
        Invoke-MgGraphRequest -Method GET -Uri $uri -OutputFilePath $OutFile
        Write-Host "  ✔ $([IO.Path]::GetFileName($OutFile))" -ForegroundColor DarkGreen
        return $true
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match "Resource not found for the segment" -or $msg -match "404" -or $msg -match "BadRequest") {
            Write-Warning "Skipped (not supported in tenant/API): $EndpointName"
            return $false
        }
        Write-Warning "Failed exporting $EndpointName : $msg"
        return $false
    }
}

function Invoke-ReportCmdletsByPattern {
    <#
      Runs *all* report cmdlets in a given module that:
        - match name pattern(s)
        - have a -Period parameter
      Output handling:
        - If cmdlet supports -OutFile => use it
        - Else capture output and write to file (best-effort)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ModuleName,
        [Parameter(Mandatory)][string[]]$NamePatterns,
        [Parameter(Mandatory)][ValidateSet("D7","D30","D90","D180")][string]$Period,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$FilePrefix
    )

    Ensure-Directory $OutputPath

    $cmds = foreach ($pat in $NamePatterns) {
        Get-Command -Module $ModuleName -Name $pat -ErrorAction SilentlyContinue | Sort-Object Name -Unique
     }
    if (-not $cmds) {
        Write-Warning "No commands matched patterns in module [$ModuleName]: $($NamePatterns -join ', ')"
        return
    }

    foreach ($c in $cmds) {
        # Only run cmdlets that accept -Period
        if (-not $c.Parameters.ContainsKey("Period")) { continue }

        $safeName = ($c.Name -replace '[^A-Za-z0-9]+','_')
        $outFile  = Join-Path $OutputPath ("{0}{1}_{2}.csv" -f $FilePrefix, $safeName, $Period)

        try {
            if ($c.Parameters.ContainsKey("OutFile")) {
                & $c.Name -Period $Period -OutFile $outFile -ErrorAction Stop | Out-Null
                Write-Host "  ✔ $([IO.Path]::GetFileName($outFile))" -ForegroundColor DarkGreen
            }
            else {
                # Best-effort: capture output and save
                $result = & $c.Name -Period $Period -ErrorAction Stop
                if ($null -ne $result) {
                    # If result is a string (CSV), write as-is; otherwise export objects to CSV
                    if ($result -is [string]) {
                        $result | Out-File -FilePath $outFile -Encoding utf8
                    }
                    else {
                        $result | Export-Csv -Path $outFile -NoTypeInformation
                    }
                    Write-Host "  ✔ $([IO.Path]::GetFileName($outFile))" -ForegroundColor DarkGreen
                }
                else {
                    Write-Warning "Cmdlet returned no output: $($c.Name)"
                }
            }
        }
        catch {
            Write-Warning "Failed running $($c.Name): $($_.Exception.Message)"
        }
    }
}

#endregion

#region Copilot Reports (Beta)

function Invoke-CopilotReports {
    param(
        [ValidateSet("D7","D30","D90","D180")]
        [string]$Period,
        [string]$OutputPath
    )

    Ensure-Directory $OutputPath
    $ProgressPreference = 'SilentlyContinue'

    # Your known Copilot beta endpoints (JSON -> CSV formatting)
    $u1 = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUserCountSummary(period='$Period')`?$format=application/json"
    $u2 = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUserCountTrend(period='$Period')`?$format=application/json"
    $u3 = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='$Period')`?$format=application/json"

    $j1 = Join-Path $OutputPath "CopilotAdoptionByProduct_$Period.json"
    $j2 = Join-Path $OutputPath "CopilotAdoptionTrend_$Period.json"
    $j3 = Join-Path $OutputPath "CopilotUsageUserDetail_$Period.json"

    Invoke-MgGraphRequest -Uri $u1 -OutputFilePath $j1
    Invoke-MgGraphRequest -Uri $u2 -OutputFilePath $j2
    Invoke-MgGraphRequest -Uri $u3 -OutputFilePath $j3

    # Format u1 -> CSV (match Script #1 fields + handle value as object OR array)
    $obj1 = Get-Content $j1 -Raw | ConvertFrom-Json

    # Some tenants return .value as an object, others as a single-element array
    $val = $obj1.value
    if ($val -is [System.Collections.IEnumerable] -and $val.GetType().Name -ne 'PSCustomObject') {
        $val = $val | Select-Object -First 1
    }

    $ad = $val.adoptionByProduct

    $formattedU1 = [PSCustomObject]@{
        "Report Refresh Date"                    = $val.reportRefreshDate
        "Report Period"                          = $ad.reportPeriod
        "Teams Enabled Users"                    = $ad.microsoftTeamsEnabledUsers
        "Teams Active Users"                     = $ad.microsoftTeamsActiveUsers
        "Word Enabled Users"                     = $ad.wordEnabledUsers
        "Word Active Users"                      = $ad.wordActiveUsers
        "PowerPoint Enabled Users"               = $ad.powerPointEnabledUsers
        "PowerPoint Active Users"                = $ad.powerPointActiveUsers
        "Outlook Enabled Users"                  = $ad.outlookEnabledUsers
        "Outlook Active Users"                   = $ad.outlookActiveUsers
        "Excel Enabled Users"                    = $ad.excelEnabledUsers
        "Excel Active Users"                     = $ad.excelActiveUsers
        "OneNote Enabled Users"                  = $ad.oneNoteEnabledUsers
        "OneNote Active Users"                   = $ad.oneNoteActiveUsers
        "Loop Enabled Users"                     = $ad.loopEnabledUsers
        "Loop Active Users"                      = $ad.loopActiveUsers
        "All Enabled Users"                      = $ad.anyAppEnabledUsers
        "All Active Users"                       = $ad.anyAppActiveUsers
        "Edge Enabled Users"                     = "" # not present yet
        "Edge Active Users"                      = "" # not present yet
        "Microsoft 365 Copilot (app) Enabled Users" = "" # not present yet
        "Microsoft 365 Copilot (app) Active Users"  = "" # not present yet
    }

    $formattedU1 | Export-Csv (Join-Path $OutputPath "CopilotAdoptionByProduct_$Period.csv") -NoTypeInformation

    # Format u2 -> CSV
    $obj2 = Get-Content $j2 -Raw | ConvertFrom-Json
    $obj2.value.adoptionByDate |
        Select-Object reportDate,
            @{n="reportPeriod";e={$obj2.value.reportPeriod}},
            microsoftTeamsEnabledUsers,
            microsoftTeamsActiveUsers,
            wordEnabledUsers,
            wordActiveUsers,
            excelEnabledUsers,
            excelActiveUsers,
            outlookEnabledUsers,
            outlookActiveUsers,
            powerPointEnabledUsers,
            powerPointActiveUsers,
            oneNoteEnabledUsers,
            oneNoteActiveUsers,
            loopEnabledUsers,
            loopActiveUsers,
            anyAppEnabledUsers,
            anyAppActiveUsers,
            copilotChatEnabledUsers,
            copilotChatActiveUsers |
        Export-Csv (Join-Path $OutputPath "CopilotAdoptionTrend_$Period.csv") -NoTypeInformation

    # Format u3 -> CSV
    (Get-Content $j3 -Raw | ConvertFrom-Json).value |
        Export-Csv (Join-Path $OutputPath "CopilotUsageUserDetail_$Period.csv") -NoTypeInformation

    # Cleanup JSON
    Remove-Item $OutputPath\Copilot*.json -Force -ErrorAction SilentlyContinue

    # NEW: attempt additional Copilot-related *beta report cmdlets* if available in your module version
    # This is the safest "add missing Copilot reports" method without guessing endpoint names.
    if (Get-Module -ListAvailable -Name Microsoft.Graph.Beta.Reports) {
        Import-Module Microsoft.Graph.Beta.Reports -ErrorAction SilentlyContinue

        Write-Host "📌 Attempting additional Copilot beta report cmdlets (module-driven)..." -ForegroundColor Cyan
        Invoke-ReportCmdletsByPattern `
            -ModuleName "Microsoft.Graph.Beta.Reports" `
            -NamePatterns @("Get-MgBetaReport*Copilot*","Get-MgBetaReport*Microsoft365Copilot*") `
            -Period $Period `
            -OutputPath $OutputPath `
            -FilePrefix "Copilot_Extra_"
    }
    else {
        Write-Host "ℹ Microsoft.Graph.Beta.Reports not installed; skipping extra Copilot cmdlets." -ForegroundColor DarkGray
    }

    Write-Host "✅ Copilot reports exported: $OutputPath" -ForegroundColor Green
}

#endregion

#region Viva Engage

function Invoke-VivaEngageReports {
    param(
        [ValidateSet("D7","D30","D90","D180")]
        [string]$Period,
        [string]$OutputPath
    )

    Ensure-Directory $OutputPath
    $ProgressPreference = 'SilentlyContinue'

    Write-Host "Exporting Viva Engage (Yammer) reports (Period=$Period) -> $OutputPath" -ForegroundColor Cyan

    # Map: cmdlet name -> output file. Each call is isolated so one failure
    # (deprecated endpoint, tenant not provisioned, de-identified data) won't
    # stop the others or kill the whole script.
    $reports = @(
        @{ Cmd = 'Get-MgReportYammerActivityCount';         File = 'VivaEngage_ActivityByDate' }
        @{ Cmd = 'Get-MgReportYammerActivityUserDetail';    File = 'VivaEngage_ActivityByUser' }
        @{ Cmd = 'Get-MgReportYammerActivityUserCount';     File = 'VivaEngage_UserCounts' }
        @{ Cmd = 'Get-MgReportYammerDeviceUsageUserDetail'; File = 'VivaEngage_DeviceUsageByUser' }
        @{ Cmd = 'Get-MgReportYammerGroupActivityDetail';   File = 'VivaEngage_GroupActivity' }
    )

    foreach ($r in $reports) {
        $outFile = Join-Path $OutputPath ("{0}_{1}.csv" -f $r.File, $Period)

        # Skip cleanly if the cmdlet isn't present in this module version
        if (-not (Get-Command $r.Cmd -ErrorAction SilentlyContinue)) {
            Write-Warning "Skipped (cmdlet not available): $($r.Cmd)"
            continue
        }

        try {
            & $r.Cmd -Period $Period -OutFile $outFile -ErrorAction Stop
            Write-Host "  OK  $([IO.Path]::GetFileName($outFile))" -ForegroundColor DarkGreen
        }
        catch {
            $msg = $_.Exception.Message
            if ($msg -match 'Resource not found|404|BadRequest|400|not supported|deprecated') {
                Write-Warning "Skipped (deprecated / not available in tenant): $($r.Cmd)"
            }
            else {
                Write-Warning "Failed running $($r.Cmd): $msg"
            }
        }
    }

    Write-Host "Viva Engage reports step complete: $OutputPath" -ForegroundColor Green
}

#endregion

#region Teams (EXPANDED - all available cmdlets)

function Invoke-TeamsReports {
    param(
        [ValidateSet("D7","D30","D90","D180")]
        [string]$Period,
        [string]$OutputPath
    )

    Ensure-Directory $OutputPath
    $ProgressPreference = 'SilentlyContinue'

    Write-Host "Exporting Teams reports (Period=$Period) -> $OutputPath" -ForegroundColor Cyan

    # Raw Graph report endpoints (same path your other workloads use).
    # This avoids the Get-MgReportTeam* progress-bar bug entirely.
    $reports = @(
        @{ Endpoint = 'getTeamsUserActivityCounts';                 File = 'Teams_UserActivity_Counts' }
        @{ Endpoint = 'getTeamsUserActivityUserCounts';             File = 'Teams_UserActivity_UserCounts' }
        @{ Endpoint = 'getTeamsUserActivityUserDetail';             File = 'Teams_UserActivity_UserDetail' }
        @{ Endpoint = 'getTeamsDeviceUsageUserCounts';              File = 'Teams_DeviceUsage_UserCounts' }
        @{ Endpoint = 'getTeamsDeviceUsageUserDetail';              File = 'Teams_DeviceUsage_UserDetail' }
        @{ Endpoint = 'getTeamsDeviceUsageDistributionUserCounts'; File = 'Teams_DeviceUsage_DistributionUserCounts' }
        @{ Endpoint = 'getTeamsTeamActivityCounts';                 File = 'Teams_TeamActivity_Counts' }
        @{ Endpoint = 'getTeamsTeamActivityDetail';                 File = 'Teams_TeamActivity_Detail' }
        @{ Endpoint = 'getTeamsTeamActivityDistributionCounts';    File = 'Teams_TeamActivity_DistributionCounts' }
        @{ Endpoint = 'getTeamsTeamCounts';                         File = 'Teams_TeamCounts' }
    )

    foreach ($r in $reports) {
        $outFile = Join-Path $OutputPath ("{0}_{1}.csv" -f $r.File, $Period)
        Export-GraphReportCsv -EndpointName $r.Endpoint -Period $Period -OutFile $outFile | Out-Null
    }

    Write-Host "Teams reports exported: $OutputPath" -ForegroundColor Green
}

#endregion

#region SharePoint

function Invoke-SharePointReports {
    param(
        [ValidateSet("D7","D30","D90","D180")]
        [string]$Period,
        [string]$OutputPath,
        [switch]$Sites,
        [switch]$Files,
        [switch]$Storage,
        [switch]$Pages,
        [switch]$Activity
    )

    Ensure-Directory $OutputPath
    $ProgressPreference = 'SilentlyContinue'

    if (-not ($Sites -or $Files -or $Storage -or $Pages -or $Activity)) {
        $Sites = $Files = $Storage = $Pages = $Activity = $true
    }

    Write-Host "📌 Exporting SharePoint reports (Period=$Period) -> $OutputPath" -ForegroundColor Cyan

    if ($Sites) {
        Export-GraphReportCsv -EndpointName "getSharePointSiteUsageDetail"       -Period $Period -OutFile "$OutputPath\SPO_SiteUsageDetail_$Period.csv" | Out-Null
    }
    if ($Files) {
        Export-GraphReportCsv -EndpointName "getSharePointSiteUsageFileCounts"   -Period $Period -OutFile "$OutputPath\SPO_SiteUsage_FileCounts_$Period.csv" | Out-Null
    }
    if ($Storage) {
        Export-GraphReportCsv -EndpointName "getSharePointSiteUsageStorage"      -Period $Period -OutFile "$OutputPath\SPO_SiteUsage_Storage_$Period.csv" | Out-Null
    }
    if ($Pages) {
        Export-GraphReportCsv -EndpointName "getSharePointSiteUsagePages"        -Period $Period -OutFile "$OutputPath\SPO_SiteUsage_Pages_$Period.csv" | Out-Null
    }
    if ($Activity) {
        Export-GraphReportCsv -EndpointName "getSharePointActivityFileCounts"    -Period $Period -OutFile "$OutputPath\SPO_Activity_FileCounts_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getSharePointActivityUserDetail"    -Period $Period -OutFile "$OutputPath\SPO_Activity_UserDetail_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getSharePointActivityPages"         -Period $Period -OutFile "$OutputPath\SPO_Activity_Pages_$Period.csv" | Out-Null
    }

    Write-Host "✅ SharePoint reports exported: $OutputPath" -ForegroundColor Green
}

#endregion

#region OneDrive (FileViews removed)

function Invoke-OneDriveReports {
    param(
        [ValidateSet("D7","D30","D90","D180")]
        [string]$Period,
        [string]$OutputPath,
        [switch]$Usage,
        [switch]$Activity
    )

    Ensure-Directory $OutputPath
    $ProgressPreference = 'SilentlyContinue'

    if (-not ($Usage -or $Activity)) { $Usage = $Activity = $true }

    Write-Host "📌 Exporting OneDrive reports (Period=$Period) -> $OutputPath" -ForegroundColor Cyan

    if ($Usage) {
        Export-GraphReportCsv -EndpointName "getOneDriveUsageAccountDetail"      -Period $Period -OutFile "$OutputPath\OD_Usage_AccountDetail_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getOneDriveUsageAccountCounts"      -Period $Period -OutFile "$OutputPath\OD_Usage_AccountCounts_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getOneDriveUsageFileCounts"         -Period $Period -OutFile "$OutputPath\OD_Usage_FileCounts_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getOneDriveUsageStorage"            -Period $Period -OutFile "$OutputPath\OD_Usage_Storage_$Period.csv" | Out-Null
    }

    if ($Activity) {
        Export-GraphReportCsv -EndpointName "getOneDriveActivityUserDetail"      -Period $Period -OutFile "$OutputPath\OD_Activity_UserDetail_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getOneDriveActivityUserCounts"      -Period $Period -OutFile "$OutputPath\OD_Activity_UserCounts_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getOneDriveActivityFileCounts"      -Period $Period -OutFile "$OutputPath\OD_Activity_FileCounts_$Period.csv" | Out-Null

        # REMOVED: getOneDriveActivityFileViews (not supported in your tenant/API)
        # Export-GraphReportCsv -EndpointName "getOneDriveActivityFileViews" -Period $Period -OutFile "$OutputPath\OD_Activity_FileViews_$Period.csv"
    }

    Write-Host "✅ OneDrive reports exported: $OutputPath" -ForegroundColor Green
}

#endregion

#region Exchange

function Invoke-ExchangeReports {
    param(
        [ValidateSet("D7","D30","D90","D180")]
        [string]$Period,
        [string]$OutputPath,
        [switch]$EmailActivity,
        [switch]$EmailAppUsage,
        [switch]$MailboxUsage
    )

    Ensure-Directory $OutputPath
    $ProgressPreference = 'SilentlyContinue'

    if (-not ($EmailActivity -or $EmailAppUsage -or $MailboxUsage)) {
        $EmailActivity = $EmailAppUsage = $MailboxUsage = $true
    }

    Write-Host "📌 Exporting Exchange reports (Period=$Period) -> $OutputPath" -ForegroundColor Cyan

    if ($EmailActivity) {
        Export-GraphReportCsv -EndpointName "getEmailActivityCounts"             -Period $Period -OutFile "$OutputPath\EX_EmailActivity_Counts_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getEmailActivityUserCounts"         -Period $Period -OutFile "$OutputPath\EX_EmailActivity_UserCounts_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getEmailActivityUserDetail"         -Period $Period -OutFile "$OutputPath\EX_EmailActivity_UserDetail_$Period.csv" | Out-Null
    }

    if ($EmailAppUsage) {
        Export-GraphReportCsv -EndpointName "getEmailAppUsageAppsUserCounts"     -Period $Period -OutFile "$OutputPath\EX_EmailAppUsage_AppsUserCounts_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getEmailAppUsageUserCounts"         -Period $Period -OutFile "$OutputPath\EX_EmailAppUsage_UserCounts_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getEmailAppUsageUserDetail"         -Period $Period -OutFile "$OutputPath\EX_EmailAppUsage_UserDetail_$Period.csv" | Out-Null
    }

    if ($MailboxUsage) {
        Export-GraphReportCsv -EndpointName "getMailboxUsageDetail"              -Period $Period -OutFile "$OutputPath\EX_MailboxUsage_Detail_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getMailboxUsageMailboxCounts"       -Period $Period -OutFile "$OutputPath\EX_MailboxUsage_MailboxCounts_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getMailboxUsageQuotaStatusMailboxCounts" -Period $Period -OutFile "$OutputPath\EX_MailboxUsage_QuotaStatus_$Period.csv" | Out-Null
        Export-GraphReportCsv -EndpointName "getMailboxUsageStorage"             -Period $Period -OutFile "$OutputPath\EX_MailboxUsage_Storage_$Period.csv" | Out-Null
    }

    Write-Host "✅ Exchange reports exported: $OutputPath" -ForegroundColor Green
}

#endregion

# -------------------------
# MAIN
# -------------------------

# Default to All if nothing specified
if (-not (
    $Copilot -or $VivaEngage -or $Teams -or $SharePoint -or $OneDrive -or $Exchange -or $All -or
    $SPSites -or $SPFiles -or $SPStorage -or $SPPages -or $SPActivity -or
    $ODUsage -or $ODActivity -or
    $EXEmailActivity -or $EXEmailAppUsage -or $EXMailboxUsage
)) {
    $All = $true
}

# Reduce risk of function count issues
Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue
$MaximumFunctionCount = 8192

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Reports -ErrorAction Stop

Connect-GraphApp -TenantId $TenantId -ClientId $ClientId -SecretPath $SecretPath

$base = Join-Path $OutputRoot "M365"

# Output paths
$copilotPath = Join-Path $base "Copilot"
$vivaPath    = Join-Path $base "VivaEngage"
$teamsPath   = Join-Path $base "Teams"
$spoPath     = Join-Path $base "SharePoint"
$odPath      = Join-Path $base "OneDrive"
$exPath      = Join-Path $base "Exchange"

try {
    if ($All -or $Copilot)    { Invoke-CopilotReports    -Period $Period -OutputPath $copilotPath }
    if ($All -or $VivaEngage) { Invoke-VivaEngageReports -Period $Period -OutputPath $vivaPath }
    if ($All -or $Teams)      { Invoke-TeamsReports      -Period $Period -OutputPath $teamsPath }

    $runGranularSP = ($SPSites -or $SPFiles -or $SPStorage -or $SPPages -or $SPActivity)
    if ($All -or $SharePoint -or $runGranularSP) {
        Invoke-SharePointReports -Period $Period -OutputPath $spoPath `
            -Sites:$SPSites -Files:$SPFiles -Storage:$SPStorage -Pages:$SPPages -Activity:$SPActivity
    }

    $runGranularOD = ($ODUsage -or $ODActivity)
    if ($All -or $OneDrive -or $runGranularOD) {
        Invoke-OneDriveReports -Period $Period -OutputPath $odPath `
            -Usage:$ODUsage -Activity:$ODActivity
    }

    $runGranularEX = ($EXEmailActivity -or $EXEmailAppUsage -or $EXMailboxUsage)
    if ($All -or $Exchange -or $runGranularEX) {
        Invoke-ExchangeReports -Period $Period -OutputPath $exPath `
            -EmailActivity:$EXEmailActivity -EmailAppUsage:$EXEmailAppUsage -MailboxUsage:$EXMailboxUsage
    }
}
finally {
    Disconnect-MgGraph | Out-Null
}

<# Examples

# Everything (D30)
.\Get-M365UsageReports.ps1 -All -Period D30

# OneDrive + Exchange only
.\Get-M365UsageReports.ps1 -OneDrive -Exchange -Period D90

# All Teams reports available in your module version
.\Get-M365UsageReports.ps1 -Teams -Period D30

# Copilot (known endpoints + any extra Copilot beta cmdlets if available)
.\Get-M365UsageReports.ps1 -Copilot -Period D30

#>
