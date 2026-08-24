#Requires -Version 5.1
<#
.SYNOPSIS
    OMMigrate-CleanRulesCSV.ps1 -- Removes duplicate rows from rules_inventory.csv.

.DESCRIPTION
    For each RuleName group in rules_inventory.csv, keeps exactly one row using
    priority scoring:
        8 pts -- valid TargetFolderPath (contains @ and \)
        4 pts -- has real Actions (not [No actions])
        2 pts -- NeedsFolderUpdate = True
        1 pt  -- IsEnabled = True

    Highest score wins. All other rows in the group are removed.
    A timestamped backup of the original CSV is written before any changes.

    Standalone utility -- does NOT read or write OMMigrate_Settings.json.
    Supports multiple Outlook profiles via the -ProfileName parameter, or an
    interactive prompt if omitted:
        - Exactly one rules_inventory_<Profile>.csv found in Config -- that
          profile is used automatically.
        - Multiple found -- you are prompted to enter the profile name.
        - None found -- falls back to the unsuffixed rules_inventory.csv
          (normal "run Script 00 first" error applies if that's missing too).

.PARAMETER ProfileName
    Target a specific Outlook profile by name (matches the profile suffix
    used in rules_inventory_<Profile>.csv). If omitted, you will be prompted
    -- unless exactly one profile-suffixed CSV already exists in Config, in
    which case it is used automatically.

.PARAMETER BasePath
    Override the default working directory.
    Default: $env:USERPROFILE\Documents\OutlookMigration

.PARAMETER LogLevel
    Logging verbosity. DEBUG | INFO | WARN | ERROR. Default: INFO

.PARAMETER WhatIf
    Preview mode -- shows what would be removed without writing changes.

.EXAMPLE
    .\OMMigrate-CleanRulesCSV.ps1 -WhatIf
    .\OMMigrate-CleanRulesCSV.ps1
    .\OMMigrate-CleanRulesCSV.ps1 -ProfileName "Outlook"
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
    [string]$LogLevel = 'INFO',

    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
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
    -ScriptName 'OMMigrate-CleanRulesCSV' `
    -BasePath   $BasePath `
    -LogLevel   $LogLevel `
    -IsWhatIf   $WhatIf.IsPresent `
    -Sanitize   $false

Register-ExitHandlers -ScriptStep 0
Show-ExitBanner


# ============================================================
#  REGION: MAIN
# ============================================================

Write-Host ''
Write-Host '  OMMigrate-CleanRulesCSV' -ForegroundColor Cyan
if ($WhatIf) {
    Write-Host '  [PREVIEW MODE] No changes will be written.' -ForegroundColor Yellow
}
Write-Host ''

# -- Resolve Outlook profile ---------------------------------------------------
# Standalone utility -- never reads or writes OMMigrate_Settings.json. The
# profile (if any) is only ever held in $Global:OMMigrate.Settings in memory
# for this run, purely so Get-OMMigrateCsvPath below resolves the correct
# profile-suffixed filename (rules_inventory_<Profile>.csv). Real profiles
# are enumerated live from the registry via Get-OutlookProfiles (same
# function Script 00 uses) -- never inferred from leftover files on disk.
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
        $ProfileName = Read-Host '  Enter the Outlook profile name to clean'
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

$Global:OMMigrate.Settings.OutlookProfile.SelectedProfile = $ProfileName

# -- Confirm before proceeding --------------------------------------------------
# Gives an explicit way out (including when a single profile was
# auto-selected with no prompt) before any file changes happen.
$proceed = Confirm-Action -Message "Clean rules_inventory.csv for profile '$ProfileName'?"
if (-not $proceed) {
    Write-Host '  Cancelled -- no changes made.' -ForegroundColor Yellow
    Write-Host ''
    exit 0
}

# -- Locate rules_inventory.csv -----------------------------------------------
$csvPath = Get-OMMigrateCsvPath -BaseName 'rules_inventory.csv'
if (-not (Test-Path $csvPath)) {
    Write-Host '  ERROR: rules_inventory.csv not found.' -ForegroundColor Red
    Write-Host "  Expected: $csvPath" -ForegroundColor Yellow
    Write-Host '  Run Script 00 first.' -ForegroundColor Yellow
    Write-Host ''
    exit 1
}

Write-Host "  Profile          : $(if ($ProfileName) { $ProfileName } else { '(none -- unsuffixed file)' })" -ForegroundColor Gray
Write-Host "  CSV              : $csvPath" -ForegroundColor Gray

# -- Backup -------------------------------------------------------------------
$backupPath = $csvPath -replace '\.csv$', ("_backup_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
if (-not $WhatIf) {
    Copy-Item -Path $csvPath -Destination $backupPath
    Write-Host "  Backup written   : $backupPath" -ForegroundColor Gray
}
Write-Host ''

# -- Load CSV -----------------------------------------------------------------
$allRows      = Import-Csv -Path $csvPath -Encoding UTF8
$keptRows     = [System.Collections.Generic.List[PSCustomObject]]::new()
$deletedCount = 0
$groupsFixed  = 0

# Separate blank separator rows -- excluded from all counts below (Total,
# Groups fixed, Rows deleted, Rows remaining all refer to real data rows
# only) but re-inserted via Add-RulesCsvSeparatorRows before the final
# write, so the output CSV keeps the same RuleStoreName-group separator
# layout Export-RulesToCSV/Script 03 already produce -- this function was
# previously dropping every separator row from the written file.
$dataRows  = @($allRows | Where-Object {
    $_.RuleStoreName -and -not [string]::IsNullOrWhiteSpace($_.RuleStoreName) -and
    $_.RuleName      -and -not [string]::IsNullOrWhiteSpace($_.RuleName)
})
$totalRows = $dataRows.Count

# -- Group by RuleName and score ----------------------------------------------
$groups = $dataRows | Group-Object -Property RuleName

foreach ($group in $groups) {
    if ($group.Count -eq 1) {
        [void]$keptRows.Add($group.Group[0])
        continue
    }

    $best      = $null
    $bestScore = -1

    foreach ($row in $group.Group) {
        $score = 0

        $hasValidPath = (
            $row.TargetFolderPath -and
            $row.TargetFolderPath -like '*@*' -and
            $row.TargetFolderPath -like '*\*'
        )
        $hasActions  = (
            $row.Actions -and
            $row.Actions -ne '[No actions]'
        )
        $needsTrue   = ($row.NeedsFolderUpdate -in @('True','true','1','TRUE'))
        $enabledTrue = ($row.IsEnabled -in @('True','true','1','TRUE'))

        if ($hasValidPath)  { $score += 8 }
        if ($hasActions)    { $score += 4 }
        if ($needsTrue)     { $score += 2 }
        if ($enabledTrue)   { $score += 1 }

        if ($score -gt $bestScore) {
            $bestScore = $score
            $best      = $row
        }
    }

    [void]$keptRows.Add($best)
    $deleted = $group.Count - 1
    $deletedCount += $deleted
    $groupsFixed++

    if ($WhatIf) {
        Write-Host "  [PREVIEW] '$($group.Name)' -- keeping 1 of $($group.Count), would delete $deleted" `
                   -ForegroundColor Yellow
    }
    else {
        Write-Host "  [FIXED]   '$($group.Name)' -- kept 1 of $($group.Count), deleted $deleted" `
                   -ForegroundColor Green
    }
}

# -- Summary ------------------------------------------------------------------
Write-Host ''
Write-Host "  Total rows       : $totalRows"    -ForegroundColor Gray
Write-Host "  Groups fixed     : $groupsFixed"  -ForegroundColor Gray
Write-Host "  Rows deleted     : $deletedCount" -ForegroundColor $(if ($deletedCount -gt 0) { 'Yellow' } else { 'Gray' })
Write-Host "  Rows remaining   : $($keptRows.Count)" -ForegroundColor Gray
Write-Host ''

if ($WhatIf) {
    Write-Host '  WhatIf mode -- no changes written.' -ForegroundColor Yellow
    exit 0
}

# -- Write cleaned CSV --------------------------------------------------------
# Re-insert blank RuleStoreName-group separator rows (reuses the same
# Add-RulesCsvSeparatorRows helper Script 00/03 use) -- $keptRows itself
# stays data-rows-only so "Rows remaining" above reports the real count.
$outputRows = Add-RulesCsvSeparatorRows -Rows $keptRows.ToArray()
$outputRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "  CSV updated      : $csvPath" -ForegroundColor Green
Write-OMMigrateLog -Message "CleanRulesCSV complete: $deletedCount duplicate rows removed. $($keptRows.Count) rows remaining." `
                   -Level INFO
Write-Host ''

# ***** END OF FILE *****
