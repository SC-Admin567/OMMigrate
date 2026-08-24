#Requires -Version 5.1
<#
.SYNOPSIS
    OMMigrate-FindDuplicateRules.ps1 -- Finds duplicate Outlook rules and saves
    a report to the OMMigrate Reports folder.

.DESCRIPTION
    Scans all rules in the active Outlook profile and reports any rule whose
    name matches another rule (including (n) suffix copies created by .rwz
    imports). Read-only diagnostic -- makes NO changes to rules or Outlook.

    Output is displayed in the console and saved to:
        Reports\DuplicateRules_<Profile>.csv

    Run OMMigrate-RemoveDuplicateRules.ps1 to remove the duplicates found.

    Standalone utility -- does NOT read or write OMMigrate_Settings.json.
    Supports multiple Outlook profiles via the -ProfileName parameter, or an
    interactive prompt if omitted: real profiles are enumerated live from
    the registry via Get-OutlookProfiles (same function Script 00 uses) --
    one found is used automatically, multiple prompts you to pick one, zero
    is an error.

.PARAMETER ProfileName
    Target a specific Outlook profile by name. If omitted, you will be
    prompted -- unless exactly one Outlook profile exists on this machine,
    in which case it is used automatically.

.PARAMETER BasePath
    Override the default working directory.
    Default: $env:USERPROFILE\Documents\OutlookMigration

.PARAMETER LogLevel
    Logging verbosity. DEBUG | INFO | WARN | ERROR. Default: INFO

.EXAMPLE
    .\OMMigrate-FindDuplicateRules.ps1
    .\OMMigrate-FindDuplicateRules.ps1 -ProfileName "Outlook"
.NOTES
    -------------------------------------------------------------------------
    OutlookMailMigrator (OMMigrate)
    -------------------------------------------------------------------------
    Originator & Architect:    Kirk Shallcross - Shallcross Consulting
    Implementation Specialist: Anthropic Claude AI
    Inception Date:            May 2026
    Version:                   1.5.2
    -------------------------------------------------------------------------
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ProfileName = '',

    [Parameter(Mandatory = $false)]
    [string]$BasePath = '',

    [Parameter(Mandatory = $false)]
    [ValidateSet('DEBUG','INFO','WARN','ERROR')]
    [string]$LogLevel = 'INFO'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8


# ============================================================
#  REGION: MODULE IMPORT
# ============================================================

# Utilities\ is a sibling of Modules\ (both directly under Documents\OMMigrate\),
# not nested inside it like Scripts\ scripts might assume -- go up one level.
$Script:ModulesPath = Join-Path $PSScriptRoot '..\Modules'

if (-not (Test-Path $Script:ModulesPath)) {
    Write-Error "OMMigrate Modules folder not found at: $Script:ModulesPath"
    exit 1
}

foreach ($moduleName in @('OMMigrate-Core','OMMigrate-Registry',
                           'OMMigrate-Outlook','OMMigrate-Reporting')) {
    $modulePath = Join-Path $Script:ModulesPath "$moduleName.psm1"
    if (-not (Test-Path $modulePath)) {
        Write-Error "Required module not found: $modulePath"
        exit 1
    }
    try {
        Import-Module $modulePath -Force -DisableNameChecking -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to import $moduleName : $_"
        exit 1
    }
}


# ============================================================
#  REGION: SESSION INITIALIZATION
# ============================================================

Initialize-OMMigrate `
    -ScriptName 'OMMigrate-FindDuplicateRules' `
    -BasePath   $BasePath `
    -LogLevel   $LogLevel `
    -IsWhatIf   $false `
    -Sanitize   $false

Register-ExitHandlers -ScriptStep 0
Show-ExitBanner


# ============================================================
#  REGION: MAIN
# ============================================================

Write-Host ''
Write-Host '  OMMigrate-FindDuplicateRules' -ForegroundColor Cyan
Write-Host '  Read-only scan -- no changes made to Outlook.' -ForegroundColor Gray
Write-Host ''

# -- Resolve Outlook profile ---------------------------------------------------
# Standalone utility -- never reads or writes OMMigrate_Settings.json. Real
# profiles are enumerated live from the registry via Get-OutlookProfiles
# (same function Script 00 uses) -- never inferred from leftover files.
$foundProfiles = @(Get-OutlookProfiles)

if ([string]::IsNullOrWhiteSpace($ProfileName)) {
    if ($foundProfiles.Count -eq 1) {
        # Exactly one Outlook profile on this machine -- use it automatically.
        $ProfileName = $foundProfiles[0].Name
    }
    elseif ($foundProfiles.Count -gt 1) {
        # Multiple profiles found -- must ask, no way to safely guess.
        Write-Host '  Multiple Outlook profiles found on this machine:' -ForegroundColor Yellow
        foreach ($p in $foundProfiles) {
            $defaultTag = if ($p.IsDefault) { ' (default)' } else { '' }
            Write-Host "    - $($p.Name)$defaultTag" -ForegroundColor Gray
        }
        Write-Host ''
        $ProfileName = Read-Host '  Enter the Outlook profile name to scan'
        if ([string]::IsNullOrWhiteSpace($ProfileName)) {
            Write-Host '  ERROR: A profile name is required when multiple profiles exist.' `
                       -ForegroundColor Red
            Write-Host ''
            exit 1
        }
    }
    else {
        # No Outlook profile found in the registry at all -- real problem,
        # not something to paper over with a fallback.
        Write-Host '  ERROR: No Outlook profile found on this machine.' -ForegroundColor Red
        Write-Host ''
        exit 1
    }
}

# Validate/normalize against the real registered profiles -- applies
# whether $ProfileName came from -ProfileName, auto-detect, or the prompt
# above. Case-insensitive match, but the matched profile's ACTUAL
# registered casing is what gets used from here on (e.g. typing
# "restorerules" resolves to the real "RestoreRules"), so the CSV
# filename, Outlook COM Logon(), and all log/console output stay
# consistent with what Script 00 itself would have produced.
$matchedProfile = $foundProfiles | Where-Object { $_.Name -eq $ProfileName } | Select-Object -First 1
if (-not $matchedProfile) {
    Write-Host "  ERROR: '$ProfileName' does not match any Outlook profile on this machine." `
               -ForegroundColor Red
    if ($foundProfiles.Count -gt 0) {
        Write-Host '  Available profiles:' -ForegroundColor Yellow
        foreach ($p in $foundProfiles) {
            Write-Host "    - $($p.Name)" -ForegroundColor Gray
        }
    }
    Write-Host ''
    exit 1
}
$ProfileName = $matchedProfile.Name

$profileName = $ProfileName
$Global:OMMigrate.Settings.OutlookProfile.SelectedProfile = $ProfileName

Write-Host "  Profile          : $profileName" -ForegroundColor Gray

# -- Confirm before proceeding --------------------------------------------------
# Gives an explicit way out (including when a single profile was
# auto-selected with no prompt) before connecting to Outlook.
$proceed = Confirm-Action -Message "Scan profile '$profileName' for duplicate rules?"
if (-not $proceed) {
    Write-Host '  Cancelled -- no scan performed.' -ForegroundColor Yellow
    Write-Host ''
    exit 0
}

Write-Host '  Connecting to Outlook...' -ForegroundColor Cyan
Write-OMMigrateLog -Message "FindDuplicateRules: Connecting to Outlook profile '$profileName'." `
                   -Level INFO

# -- Connect to Outlook -------------------------------------------------------
$outlook   = $null
$namespace = $null
try {
    Add-Type -AssemblyName 'Microsoft.Office.Interop.Outlook' -ErrorAction SilentlyContinue
} catch { }

try {
    $outlook   = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace('MAPI')
    $namespace.Logon($profileName, $null, $false, $true)
}
catch {
    Write-Host "  ERROR: Could not connect to Outlook: $_" -ForegroundColor Red
    Write-OMMigrateLog -Message "FindDuplicateRules: Outlook connection failed: $_" -Level ERROR
    exit 1
}

Write-Host '  Scanning rules...' -ForegroundColor Cyan

# -- Scan all stores for rules ------------------------------------------------
$allRules   = [System.Collections.Generic.List[PSCustomObject]]::new()
$duplicates = [System.Collections.Generic.List[PSCustomObject]]::new()

try {
    $stores = $namespace.Stores
    for ($s = 1; $s -le $stores.Count; $s++) {
        $store = $null
        try {
            $store     = $stores.Item($s)
            $storeName = ''
            try { $storeName = $store.DisplayName } catch { continue }

            $rules = $null
            try { $rules = $store.GetRules() } catch { continue }

            for ($r = 1; $r -le $rules.Count; $r++) {
                try {
                    $rule     = $rules.Item($r)
                    $ruleName = $rule.Name
                    # Base name = strip trailing (n) suffix
                    $baseName = $ruleName -replace '\s*\(\d+\)$', ''

                    $hasMoveFolder = $false
                    try {
                        $mf = $rule.Actions.MoveToFolder
                        $hasMoveFolder = ($mf.Enabled -and $null -ne $mf.Folder)
                    } catch { }

                    $allRules.Add([PSCustomObject]@{
                        StoreName     = $storeName
                        RuleName      = $ruleName
                        BaseName      = $baseName
                        Enabled       = $rule.Enabled
                        Order         = $rule.ExecutionOrder
                        HasMoveFolder = $hasMoveFolder
                    })
                } catch { }
            }
        } catch { }
    }
}
finally {
    try { $namespace.Logoff() } catch { }
    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)   | Out-Null
    } catch { }
}

# -- Find duplicates by BaseName ----------------------------------------------
$grouped = $allRules | Group-Object -Property BaseName
foreach ($group in $grouped) {
    if ($group.Count -gt 1) {
        foreach ($entry in $group.Group) {
            $duplicates.Add([PSCustomObject]@{
                StoreName     = $entry.StoreName
                RuleName      = $entry.RuleName
                BaseName      = $entry.BaseName
                Enabled       = $entry.Enabled
                Order         = $entry.Order
                HasMoveFolder = $entry.HasMoveFolder
            })
        }
    }
}

# -- Report -------------------------------------------------------------------
Write-Host ''
if ($duplicates.Count -eq 0) {
    Write-Host '  No duplicate rules found.' -ForegroundColor Green
    Write-OMMigrateLog -Message 'FindDuplicateRules: No duplicates found.' -Level INFO
}
else {
    Write-Host "  Found $($duplicates.Count) duplicate rule entries across $(@($grouped | Where-Object { $_.Count -gt 1 }).Count) rule name(s):" `
               -ForegroundColor Yellow
    Write-Host ''
    $duplicates | Sort-Object StoreName, BaseName, RuleName | Format-Table StoreName, RuleName, Enabled, Order, HasMoveFolder -AutoSize

    # Profile-suffixed report filename -- reuses Get-OMMigrateCsvPath's
    # existing suffix/sanitize logic via -BasePathOverride, pointed at
    # ReportPath instead of its default ConfigPath. This is the fixed
    # filename OMMigrate-RemoveDuplicateRules.ps1 reads -- always
    # overwritten with the latest scan for this profile.
    $csvPath = Get-OMMigrateCsvPath -BaseName 'DuplicateRules.csv' `
                                    -BasePathOverride $Global:OMMigrate.ReportPath
    $sortedDuplicates = $duplicates | Sort-Object StoreName, BaseName, RuleName
    $sortedDuplicates | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "  CSV saved        : $csvPath" -ForegroundColor Green

    # Timestamped history copy -- preserves every run's findings so an
    # overwrite of the fixed-name file above never silently loses a prior
    # scan's results. Never read by any script; for the operator's own
    # record-keeping only.
    $historyStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $historyPath  = $csvPath -replace '\.csv$', "_$historyStamp.csv"
    $sortedDuplicates | Export-Csv -Path $historyPath -NoTypeInformation -Encoding UTF8
    Write-Host "  History copy     : $historyPath" -ForegroundColor Gray

    Write-Host "  Run OMMigrate-RemoveDuplicateRules.ps1 -ProfileName `"$profileName`" to remove the duplicates." `
               -ForegroundColor Cyan
    Write-OMMigrateLog -Message "FindDuplicateRules: $($duplicates.Count) duplicate entries found. CSV: $csvPath" `
                       -Level INFO
}
Write-Host ''

# ***** END OF FILE *****
