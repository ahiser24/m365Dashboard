<#
.SYNOPSIS
    Downloads Microsoft 365 usage reports (Copilot, Viva Engage, Teams) via Microsoft Graph and exports to CSV/JSON.
    Additional Usage Reports to download manually:
    - Copilot
      - Copilot usage details
    - Copilot Chat
      -Adoption by app
      - Prompts submitted by app
      - Usage Details
    - Agents
      - Usage details for Agents
    All Agents:
      -Export Agent Registry

.DESCRIPTION
    - Connects to Microsoft Graph using app-only authentication (Client Secret Credential)
    - Runs selected report groups:
        * Copilot (beta reports -> JSON + CSV formatting)
        * Viva Engage (Yammer) reports -> CSV
        * Teams user activity detail -> CSV
    - Exports into a consistent folder structure under a root output directory.

.PREREQS
    Install-Module Microsoft.Graph -Scope CurrentUser
    Install-Module Microsoft.Graph.Reports -Scope CurrentUser
    Install-Module Microsoft.Graph.Beta.Reports -Scope CurrentUser

    Create an App Registration for Microsoft Graph access.

    .NOTES
    Secret is loaded from:
        $env:USERPROFILE\graph-client-secret.xml
    Create the secret xml if running locally:
        "YOUR_SECRET" | ConvertTo-SecureString -AsPlainText -Force | Export-Clixml -Path "$env:USERPROFILE\graph-client-secret.xml"
#>


[CmdletBinding()]
param(
    # App registration details
    [Parameter()]
    [string]$TenantId = "INSERT TENANT ID",

    [Parameter()]
    [string]$ClientId = "INSERT CLIENT ID",

    # Secret file path
    [Parameter()]
    [string]$SecretPath = "$env:USERPROFILE\graph-client-secret.xml",

    # Report period
    [Parameter()]
    [ValidateSet("D7","D30","D60","D90","D180")]
    [string]$Period = "D30",

    # Output root
    [Parameter()]
    [string]$OutputRoot = "C:\Reports",

    # Choose which report sets to run
    [Parameter()]
    [switch]$Copilot,

    [Parameter()]
    [switch]$VivaEngage,

    [Parameter()]
    [switch]$Teams,

    # If none chosen, run all
    [Parameter()]
    [switch]$All
)

#region Helper Functions

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

function Connect-GraphApp {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$SecretPath
    )

    if (-not (Test-Path $SecretPath)) {
        throw "Secret file not found at: $SecretPath"
    }

    # Load encrypted secret from file
    $secure = Import-Clixml -Path $SecretPath

    # Create PSCredential (username is ClientId for this pattern)
    $appCred = New-Object System.Management.Automation.PSCredential ($ClientId, $secure)

    # Connect once
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $appCred -NoWelcome | Out-Null
}

#endregion Helper Functions

#region Copilot Reports (Beta endpoints)

function Invoke-CopilotReports {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet("D7","D30","D60","D90","D180")][string]$Period,
        [Parameter(Mandatory)][string]$OutputPath
    )

    Ensure-Directory -Path $OutputPath

    $ProgressPreference = 'SilentlyContinue'

    # --- Beta report endpoints (JSON) ---
    $u1 = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUserCountSummary(period='$Period')`?$format=application/json"
    $u2 = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUserCountTrend(period='$Period')`?$format=application/json"
    $u3 = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='$Period')`?$format=application/json"

    $json1 = Join-Path $OutputPath "CopilotAdoptionByProduct_$Period.json"
    $json2 = Join-Path $OutputPath "CopilotAdoptionTrend_$Period.json"
    $json3 = Join-Path $OutputPath "CopilotUsageUserDetail_$Period.json"

    Invoke-MgGraphRequest -Method GET -Uri $u1 -OutputFilePath $json1
    Invoke-MgGraphRequest -Method GET -Uri $u2 -OutputFilePath $json2
    Invoke-MgGraphRequest -Method GET -Uri $u3 -OutputFilePath $json3

    # --- Format u1 -> CSV ---
    $obj1 = Get-Content $json1 -Raw | ConvertFrom-Json

    # Defensive: handle if value is missing/unexpected
    $v = $obj1.value
    $adopt = $v.adoptionByProduct

    $formatted1 = [PSCustomObject]@{
        "Report Refresh Date"                    = $v.reportRefreshDate
        "Report Period"                          = $adopt.reportPeriod
        "Teams Enabled Users"                    = $adopt.microsoftTeamsEnabledUsers
        "Teams Active Users"                     = $adopt.microsoftTeamsActiveUsers
        "Word Enabled Users"                     = $adopt.wordEnabledUsers
        "Word Active Users"                      = $adopt.wordActiveUsers
        "PowerPoint Enabled Users"               = $adopt.powerPointEnabledUsers
        "PowerPoint Active Users"                = $adopt.powerPointActiveUsers
        "Outlook Enabled Users"                  = $adopt.outlookEnabledUsers
        "Outlook Active Users"                   = $adopt.outlookActiveUsers
        "Excel Enabled Users"                    = $adopt.excelEnabledUsers
        "Excel Active Users"                     = $adopt.excelActiveUsers
        "OneNote Enabled Users"                  = $adopt.oneNoteEnabledUsers
        "OneNote Active Users"                   = $adopt.oneNoteActiveUsers
        "Loop Enabled Users"                     = $adopt.loopEnabledUsers
        "Loop Active Users"                      = $adopt.loopActiveUsers
        "All Enabled Users"                      = $adopt.anyAppEnabledUsers
        "All Active Users"                       = $adopt.anyAppActiveUsers
        "Edge Enabled Users"                     = ""
        "Edge Active Users"                      = ""
        "Microsoft 365 Copilot (app) Enabled Users" = ""
        "Microsoft 365 Copilot (app) Active Users"  = ""
    }

    $formatted1 | Export-Csv (Join-Path $OutputPath "CopilotAdoptionByProduct_$Period.csv") -NoTypeInformation

    # --- Format u2 -> CSV ---
    $obj2 = Get-Content $json2 -Raw | ConvertFrom-Json
    $records = $obj2.value.adoptionByDate

    $formatted2 = $records | Select-Object `
        reportDate,
        @{Name="reportPeriod"; Expression={ $obj2.value.reportPeriod }},
        @{Name="Teams Enabled Users"; Expression={ $_.microsoftTeamsEnabledUsers }},
        @{Name="Teams Active Users"; Expression={ $_.microsoftTeamsActiveUsers }},
        @{Name="Word Enabled Users"; Expression={ $_.wordEnabledUsers }},
        @{Name="Word Active Users"; Expression={ $_.wordActiveUsers }},
        @{Name="PowerPoint Enabled Users"; Expression={ $_.powerPointEnabledUsers }},
        @{Name="PowerPoint Active Users"; Expression={ $_.powerPointActiveUsers }},
        @{Name="Outlook Enabled Users"; Expression={ $_.outlookEnabledUsers }},
        @{Name="Outlook Active Users"; Expression={ $_.outlookActiveUsers }},
        @{Name="Excel Enabled Users"; Expression={ $_.excelEnabledUsers }},
        @{Name="Excel Active Users"; Expression={ $_.excelActiveUsers }},
        @{Name="OneNote Enabled Users"; Expression={ $_.oneNoteEnabledUsers }},
        @{Name="OneNote Active Users"; Expression={ $_.oneNoteActiveUsers }},
        @{Name="Loop Enabled Users"; Expression={ $_.loopEnabledUsers }},
        @{Name="Loop Active Users"; Expression={ $_.loopActiveUsers }},
        @{Name="All Enabled Users"; Expression={ $_.anyAppEnabledUsers }},
        @{Name="All Active Users"; Expression={ $_.anyAppActiveUsers }},
        @{Name="Edge Enabled Users"; Expression={ "" }},
        @{Name="Edge Active Users"; Expression={ "" }},
        @{Name="Microsoft 365 Copilot (app) Enabled Users"; Expression={ "" }},
        @{Name="Microsoft 365 Copilot (app) Active Users"; Expression={ "" }},
        @{Name="Copilot Chat Enabled Users"; Expression={ $_.copilotChatEnabledUsers }},
        @{Name="Copilot Chat Active Users"; Expression={ $_.copilotChatActiveUsers }}

    $formatted2 | Export-Csv (Join-Path $OutputPath "CopilotAdoptionTrend_$Period.csv") -NoTypeInformation

    # --- Format u3 -> CSV ---
    $obj3 = Get-Content $json3 -Raw | ConvertFrom-Json

    $formatted3 = $obj3.value | Select-Object `
        userPrincipalName, displayName, lastActivityDate, copilotChatLastActivityDate,
        microsoftTeamsCopilotLastActivityDate, wordCopilotLastActivityDate, excelCopilotLastActivityDate,
        powerPointCopilotLastActivityDate, outlookCopilotLastActivityDate, oneNoteCopilotLastActivityDate,
        loopCopilotLastActivityDate

    $formatted3 | Export-Csv (Join-Path $OutputPath "CopilotUsageUserDetail_$Period.csv") -NoTypeInformation

    Write-Host "✅ Copilot reports exported to: $OutputPath" -ForegroundColor Green

    # --- Cleanup JSON files after CSV export ---
    Get-ChildItem -Path $OutputPath -Filter 'Copilot*.json' -File | Remove-Item -Force -ErrorAction SilentlyContinue
    Write-Host " 🧹 Copilot JSON files removed" -ForegroundColor DarkGray
}

#endregion

#region Viva Engage Reports

function Invoke-VivaEngageReports {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet("D7","D30","D60","D90","D180")][string]$Period,
        [Parameter(Mandatory)][string]$OutputPath
    )

    Ensure-Directory -Path $OutputPath
    $ProgressPreference = 'SilentlyContinue'

    Get-MgReportYammerActivityCount -Period $Period -OutFile (Join-Path $OutputPath "VivaEngage_ActivityByDate_$Period.csv")
    Get-MgReportYammerActivityUserDetail -Period $Period -OutFile (Join-Path $OutputPath "VivaEngage_ActivityByUser_$Period.csv")
    Get-MgReportYammerActivityUserCount -Period $Period -OutFile (Join-Path $OutputPath "VivaEngage_UserCounts_$Period.csv")
    Get-MgReportYammerDeviceUsageUserDetail -Period $Period -OutFile (Join-Path $OutputPath "VivaEngage_DeviceUsageByUser_$Period.csv")
    Get-MgReportYammerGroupActivityDetail -Period $Period -OutFile (Join-Path $OutputPath "VivaEngage_GroupActivityDetail_$Period.csv")

    Write-Host "✅ Viva Engage reports exported to: $OutputPath" -ForegroundColor Green
}

#endregion

#region Teams Reports

function Invoke-TeamsReports {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet("D7","D30","D60","D90","D180")][string]$Period,
        [Parameter(Mandatory)][string]$OutputPath
    )

    Ensure-Directory -Path $OutputPath
    $ProgressPreference = 'SilentlyContinue'

    # Keep the filename deterministic (no spaces is usually nicer for automation)
    Get-MgReportTeamUserActivityUserDetail -Period $Period -OutFile (Join-Path $OutputPath "Teams_UserActivityUserDetail_$Period.csv")

    Write-Host "✅ Teams reports exported to: $OutputPath" -ForegroundColor Green
}

#endregion

# -------------------------
# Main
# -------------------------

# If user didn’t specify any switches, default to All
if (-not ($Copilot -or $VivaEngage -or $Teams -or $All)) {
    $All = $true
}

# Import modules once
Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue
# Must be set before importing large modules
$MaximumFunctionCount = 8192

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# Only load the report commands you actually use (dramatically reduces function count)
Import-Module Microsoft.Graph.Reports -Function `
    Get-MgReportYammerActivityCount,
    Get-MgReportYammerActivityUserDetail,
    Get-MgReportYammerActivityUserCount,
    Get-MgReportYammerDeviceUsageUserDetail,
    Get-MgReportYammerGroupActivityDetail,
    Get-MgReportTeamUserActivityUserDetail -ErrorAction Stop

# Connect once
Connect-GraphApp -TenantId $TenantId -ClientId $ClientId -SecretPath $SecretPath

# Output paths
$copilotPath = Join-Path $OutputRoot "M365\Copilot"
$vivaPath    = Join-Path $OutputRoot "M365\VivaEngage"
$teamsPath   = Join-Path $OutputRoot "M365\Teams"

try {
    if ($All -or $Copilot)    { Invoke-CopilotReports -Period $Period -OutputPath $copilotPath }
    if ($All -or $VivaEngage) { Invoke-VivaEngageReports -Period $Period -OutputPath $vivaPath }
    if ($All -or $Teams)      { Invoke-TeamsReports -Period $Period -OutputPath $teamsPath }
}
finally {
    Disconnect-MgGraph | Out-Null
}

# Run everything for 180 days:
# .\Get-M365UsageReports.ps1 -All -Period D180

# Run only Copilot + Teams for 30 days: 
# .\Get-M365UsageReports.ps1 -Copilot -Teams -Period D30

# Change output root:
# .\Get-M365UsageReports.ps1 -All -Period D90 -OutputRoot "D:\Reporting\M365"
