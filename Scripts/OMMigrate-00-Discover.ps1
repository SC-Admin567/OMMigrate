#Requires -Version 5.1
<#
.SYNOPSIS
    OMMigrate-00-Discover.ps1 -- Account Discovery and Inventory

.DESCRIPTION
    Step 00 of the OutlookMailMigrator (OMMigrate) toolkit.

    This is the first script to run and the ONLY script that is safe
    to run at any time without risk -- it makes NO changes to Outlook,
    the registry, any PST files, or any account settings.

    What this script does:
        1. Validates the environment (PowerShell version, OS, Outlook)
        2. Reads all Outlook profiles from the Windows Registry
        3. Discovers and classifies all email accounts (POP3/IMAP/Exchange)
        4. Locates PST and OST data files for each account
        5. Corroborates registry data with live Outlook COM session
        6. Enumerates the complete folder tree across all stores
        7. Inventories all Outlook Rules with folder target flags
        8. Generates pre-filled migration_accounts.csv control file
        9. Generates pre-filled folder_map.csv control file
       10. Generates rules_inventory_<Profile>.csv for Script 03 reference
       11. Generates the Discovery HTML report

    Output files written to Config\ :
        migration_accounts.csv    -- Operator adds passwords, confirms servers
        folder_map.csv            -- Pre-filled by Script 00. Server/Local destinations
                                     are set automatically via the folder picker in
                                     Script 01. Re-run Script 01 -RefreshFolderMap
                                     to adjust destinations at any time.
        rules_inventory_<Profile>.csv -- Outlook rules with folder update flags

    Output files written to Reports\ :
        Discovery_Report_YYYYMMDD.html

    After running this script:
        1. Open Config\migration_accounts.csv
           Add your email password for each POP3 account (marked [ENTER_PASSWORD])
           Confirm or correct server settings for any accounts that changed
        2. Run Script 01 (Backup)
           Script 01 backs up all accounts, builds the Archive PST folder
           structures, and opens the folder destination picker automatically.
           Review Server/Local assignments in the picker -- Local folders stay
           on this machine only, Server folders sync to all your devices.
        3. Complete any Prerequisites listed in the Discovery Report
           (e.g. generate Secure Mail Key for ameritech.net)
        4. Run OMMigrate-01-Backup.ps1

.PARAMETER ProfileName
    Target a specific Outlook profile by name.
    If omitted, all profiles found on this machine are scanned.

.PARAMETER BasePath
    Override the default working directory.
    Default: $env:USERPROFILE\Documents\OutlookMigration

.PARAMETER Preview
    Simulate execution -- reads registry and COM data but writes
    no output files (no CSV, no report, no manifest).
    Safe to use for testing the discovery logic.

.PARAMETER LogLevel
    Logging verbosity written to the run log file.
    DEBUG | INFO | WARN | ERROR
    Default: INFO

.PARAMETER OpenReport
    Open the Discovery HTML report in the default browser after completion.
    Default: $true (respects OMMigrate_Settings.json setting)

.EXAMPLE
    # Standard run -- scans all profiles, generates all output files
    .\OMMigrate-00-Discover.ps1

.EXAMPLE
    # Target a specific Outlook profile
    .\OMMigrate-00-Discover.ps1 -ProfileName "Outlook"

.EXAMPLE
    # Dry run -- no files written, useful for testing
    .\OMMigrate-00-Discover.ps1 -Preview

.EXAMPLE
    # Custom working directory with verbose logging
    .\OMMigrate-00-Discover.ps1 -BasePath "D:\Migration" -LogLevel DEBUG

.NOTES
    -------------------------------------------------------------------------
    OutlookMailMigrator (OMMigrate)
    -------------------------------------------------------------------------
    Originator & Architect:    Kirk Shallcross - Shallcross Consulting
    Implementation Specialist: Anthropic Claude AI
    Inception Date:            May 2026
    Version:                   1.5.2
    -------------------------------------------------------------------------

    Requirements:
        PowerShell  : 5.1 or higher
        Windows     : 10 or 11 (64-bit)
        Outlook     : Classic Outlook 2016 / 2019 / 2021 (not New Outlook)

    Run As:
        The same Windows user account that owns the Outlook profile.
        Do NOT run as Administrator.

    Safe to run multiple times:
        This script is read-only. Running it again overwrites the CSV
        and report files with fresh data -- nothing else is affected.

    Module dependencies (must be in .\Modules\ relative to this script):
        OMMigrate-Core.psm1
        OMMigrate-Registry.psm1
        OMMigrate-Outlook.psm1
        OMMigrate-Reporting.psm1
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
    [switch]$OpenReport,

    [Parameter(Mandatory = $false)]
    [switch]$Preview,

    [Parameter(Mandatory = $false)]
    [switch]$Sanitize,

    # *** UNDOCUMENTED -- FOR DEVELOPER USE ONLY ***
    # Skips COM rules enumeration and instead patches SendersDomain
    # into the existing rules_inventory_<Profile>.csv from disk.
    # Used when the active Outlook profile has an overflowed IMAP/OST
    # rules blob that prevents COM enumeration (e.g. Administrator's environment).
    # Normal users should NEVER use this parameter.
    [Parameter(Mandatory = $false)]
    [switch]$PatchRulesCSV
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Capture Preview (WhatIf) state from explicit -Preview switch parameter
$Script:IsWhatIf = $Preview.IsPresent


# ============================================================
#  REGION: MODULE IMPORT
#  All OMMigrate modules must be in .\Modules\ relative to
#  the location of this script file.
# ============================================================

$Script:ModulesPath = Join-Path $PSScriptRoot '..\Modules'

# Verify Modules folder exists before attempting imports
if (-not (Test-Path $Script:ModulesPath)) {
    Write-Error (
        "OMMigrate Modules folder not found at: $Script:ModulesPath`n" +
        "Ensure all four .psm1 files are in the Modules\ folder " +
        "relative to this script."
    )
    exit 1
}

# Import modules in dependency order
$Script:RequiredModules = @(
    'OMMigrate-Core',
    'OMMigrate-Registry',
    'OMMigrate-Outlook',
    'OMMigrate-Reporting'
)

foreach ($moduleName in $Script:RequiredModules) {
    $modulePath = Join-Path $Script:ModulesPath "$moduleName.psm1"

    if (-not (Test-Path $modulePath)) {
        Write-Error (
            "Required module not found: $modulePath`n" +
            "Ensure $moduleName.psm1 exists in the Modules\ folder."
        )
        exit 1
    }

    try {
        # Force re-import to ensure latest version is loaded
        Import-Module $modulePath -Force -DisableNameChecking -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to import module $moduleName : $_"
        exit 1
    }
}


# ============================================================
#  REGION: SESSION INITIALIZATION
# ============================================================

# Initialize the OMMigrate session context -- sets up paths,
# logging, settings, and writes the run log header
Initialize-OMMigrate `
    -ScriptName 'OMMigrate-00-Discover' `
    -BasePath   $BasePath `
    -LogLevel   $LogLevel `
    -IsWhatIf   $Script:IsWhatIf `
    -Sanitize   $Sanitize.IsPresent

# Register Ctrl+C and process-exit handlers
# Script 00 is read-only so exit handling is lighter --
# no checkpoint needed, but COM must still be released
Register-ExitHandlers -ScriptStep 0

# -- Sanitize Write-Host override ------------------------------
# When -Sanitize is active this local function intercepts every
# Write-Host call in this script scope and passes the message
# through Invoke-OMMigrateSanitize before writing to the console.
# When -Sanitize is not active it calls the real Write-Host directly
# with no overhead -- the override is fully transparent either way.
function Write-Host {
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [object]$Object = '',
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor,
        [switch]$NoNewline,
        [switch]$Separator
    )
    if ($Global:OMMigrate.Sanitize -and $Object -is [string] -and $Object -ne '') {
        $Object = Invoke-OMMigrateSanitize -Text $Object
    }
    $params = @{ Object = $Object }
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $params['ForegroundColor'] = $ForegroundColor }
    if ($PSBoundParameters.ContainsKey('BackgroundColor')) { $params['BackgroundColor'] = $BackgroundColor }
    if ($NoNewline)  { $params['NoNewline']  = $true }
    if ($Separator)  { $params['Separator']  = $true }
    Microsoft.PowerShell.Utility\Write-Host @params
}

# Track whether COM session was opened so finally block
# knows whether to release it
$Script:COMSessionOpen = $false

# Resolve OpenReport from settings if not explicitly passed on command line
# Settings file OpenReportAfterRun = $true means report opens automatically.
# Passing -OpenReport on the command line always overrides the setting.
if (-not $OpenReport.IsPresent) {
    $openReportSetting = $true  # default
    try {
        if ($Global:OMMigrate.Settings -and
            $null -ne $Global:OMMigrate.Settings.Reporting.OpenReportAfterRun) {
            $openReportSetting = $Global:OMMigrate.Settings.Reporting.OpenReportAfterRun
        }
    }
    catch { }
    $Script:OpenReport = $openReportSetting
}
else {
    $Script:OpenReport = $true
}

# Collection variables -- populated through the script
$Script:AllAccounts     = [System.Collections.Generic.List[PSCustomObject]]::new()
$Script:AllFolders      = [System.Collections.Generic.List[PSCustomObject]]::new()
$Script:AllRules        = [System.Collections.Generic.List[PSCustomObject]]::new()
$Script:AllDataFiles    = [System.Collections.Generic.List[PSCustomObject]]::new()
$Script:DiscoveredProfiles = @()
$Script:ReportFile      = ''
$Script:FinalStatus     = 'SUCCESS'
$accountsCsvPath        = ''


# ============================================================
#  MAIN EXECUTION BLOCK
#  Wrapped in try/finally to guarantee COM release and
#  session completion regardless of how the script exits
# ============================================================

try {

    # ----------------------------------------------------------
    #  RE-RUN DETECTION -- Status Summary
    #  If migration_accounts.csv already exists this is a re-run.
    #  Show the operator a snapshot of where things stand before
    #  the account picker opens so they have full context for
    #  their selections. Skipped on first run (no CSV yet).
    # ----------------------------------------------------------
    $reRunCsvPath = Get-OMMigrateCsvPath -BaseName 'migration_accounts.csv'
    if (Test-Path $reRunCsvPath) {
        try {
            $reRunRows = Import-Csv -Path $reRunCsvPath -Encoding UTF8

            $rrComplete       = @($reRunRows | Where-Object { $_.MigrationAction -eq 'COMPLETE'     }).Count
            $rrImapConverted  = @($reRunRows | Where-Object { $_.ProviderTag     -eq 'IMAP-CONVERTED' }).Count
            $rrFolderOnly     = @($reRunRows | Where-Object {
                                    $_.MigrationAction -eq 'FOLDER-ONLY' -and
                                    $_.ProviderTag     -ne 'IMAP-CONVERTED'
                                }).Count
            $rrMigrate        = @($reRunRows | Where-Object { $_.MigrationAction -eq 'MIGRATE'      }).Count
            $rrSkip           = @($reRunRows | Where-Object { $_.MigrationAction -eq 'SKIP'         }).Count
            $rrOther          = @($reRunRows | Where-Object {
                                    $_.MigrationAction -notin @('COMPLETE','FOLDER-ONLY','MIGRATE','SKIP')
                                }).Count

            Write-Host ''
            Write-Host ('-' * 60) -ForegroundColor DarkCyan
            Write-Host '  RE-RUN DETECTED -- Current Migration Status' -ForegroundColor White
            Write-Host ('-' * 60) -ForegroundColor DarkCyan
            Write-Host ''
            Write-Host ("  {0,-30}: {1,3}" -f 'Fully complete (Script 03 done)',   $rrComplete)      -ForegroundColor $(if ($rrComplete      -gt 0) { 'Green'  } else { 'DarkGray' })
            Write-Host ("  {0,-30}: {1,3}" -f 'Converted to IMAP (Script 02 done)', $rrImapConverted) -ForegroundColor $(if ($rrImapConverted -gt 0) { 'Green'  } else { 'DarkGray' })
            Write-Host ("  {0,-30}: {1,3}" -f 'Already IMAP (folder-only)',         $rrFolderOnly)    -ForegroundColor $(if ($rrFolderOnly    -gt 0) { 'Cyan'   } else { 'DarkGray' })
            Write-Host ("  {0,-30}: {1,3}" -f 'Pending migration (MIGRATE)',        $rrMigrate)       -ForegroundColor $(if ($rrMigrate       -gt 0) { 'Yellow' } else { 'DarkGray' })
            Write-Host ("  {0,-30}: {1,3}" -f 'Deferred (SKIP)',                    $rrSkip)          -ForegroundColor $(if ($rrSkip          -gt 0) { 'Yellow' } else { 'DarkGray' })
            if ($rrOther -gt 0) {
                Write-Host ("  {0,-30}: {1,3}" -f 'Other (Exchange/Manual/etc.)', $rrOther) -ForegroundColor DarkGray
            }
            Write-Host ''
            Write-Host '  Discovery will re-scan Outlook and refresh the CSV.' -ForegroundColor DarkGray
            Write-Host '  Your previous selections and credentials are preserved.' -ForegroundColor DarkGray
            Write-Host ''

            Write-OMMigrateLog -Message (
                "Re-run detected -- existing CSV summary: " +
                "Complete=$rrComplete | IMAP-CONVERTED=$rrImapConverted | " +
                "FolderOnly=$rrFolderOnly | MIGRATE=$rrMigrate | " +
                "SKIP=$rrSkip | Other=$rrOther"
            ) -Level INFO
        }
        catch {
            # Non-fatal -- if the CSV can't be read, just skip the summary
            Write-OMMigrateLog -Message "Re-run detection: could not read existing CSV (non-fatal): $_" `
                               -Level WARN
        }
    }


    # ----------------------------------------------------------
    #  EXCEL OPEN FILE CHECK
    #  If Excel is running and has any OMMigrate CSV files open,
    #  prompt the operator to save and close before continuing.
    #  CSV writes will fail silently if Excel holds a file lock.
    #  Skipped in WhatIf/Preview mode.
    # ----------------------------------------------------------
    if (-not $Script:IsWhatIf) {
        $excelProc = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
        if ($excelProc) {
            # Check if any of the three OMMigrate CSVs are locked by Excel
            $csvNames = @(
                'migration_accounts.csv',
                'folder_map.csv',
                'rules_inventory.csv'
            )
            $lockedFiles = [System.Collections.Generic.List[string]]::new()

            foreach ($csvName in $csvNames) {
                # Use Get-OMMigrateCsvPath so the profile-suffixed filename
                # is resolved correctly (e.g. folder_map_Outlook.csv).
                $csvTestPath = Get-OMMigrateCsvPath -BaseName $csvName
                if (Test-Path $csvTestPath) {
                    try {
                        $lockTest = [System.IO.File]::Open(
                            $csvTestPath,
                            [System.IO.FileMode]::Open,
                            [System.IO.FileAccess]::ReadWrite,
                            [System.IO.FileShare]::None
                        )
                        $lockTest.Close()
                        $lockTest.Dispose()
                    }
                    catch {
                        $lockedFiles.Add($csvName)
                    }
                }
            }

            if ($lockedFiles.Count -gt 0) {
                Write-Host ''
                Write-Host ('-' * 60) -ForegroundColor Yellow
                Write-Host '  EXCEL IS OPEN -- ACTION REQUIRED' -ForegroundColor Yellow
                Write-Host ('-' * 60) -ForegroundColor Yellow
                Write-Host ''
                Write-Host '  Excel has the following OMMigrate file(s) open:' -ForegroundColor Yellow
                foreach ($f in $lockedFiles) {
                    Write-Host "    * $f" -ForegroundColor White
                }
                Write-Host ''
                Write-Host '  Save any changes you want to keep, then close Excel.' -ForegroundColor Gray
                Write-Host '  The script will wait until Excel releases the file(s).' -ForegroundColor Gray
                Write-Host ''

                Write-OMMigrateLog -Message (
                    "Excel file lock detected on: $($lockedFiles -join ', ') -- waiting for operator to close Excel."
                ) -Level INFO

                # Poll until all locked files are released
                $waitedSeconds = 0
                $pollInterval  = 3
                while ($true) {
                    Start-Sleep -Seconds $pollInterval
                    $waitedSeconds += $pollInterval

                    $stillLocked = $false
                    foreach ($csvName in $lockedFiles) {
                        $csvTestPath = Get-OMMigrateCsvPath -BaseName $csvName
                        if (Test-Path $csvTestPath) {
                            try {
                                $lt = [System.IO.File]::Open(
                                    $csvTestPath,
                                    [System.IO.FileMode]::Open,
                                    [System.IO.FileAccess]::ReadWrite,
                                    [System.IO.FileShare]::None
                                )
                                $lt.Close()
                                $lt.Dispose()
                            }
                            catch {
                                $stillLocked = $true
                                break
                            }
                        }
                    }

                    if (-not $stillLocked) {
                        Write-Host '  Excel closed -- continuing.' -ForegroundColor Green
                        Write-Host ''
                        Write-OMMigrateLog -Message "Excel file lock released after ${waitedSeconds}s -- continuing." `
                                           -Level INFO
                        break
                    }

                    # Remind every 15 seconds
                    if ($waitedSeconds % 15 -eq 0) {
                        Write-Host "  Still waiting for Excel to close... (${waitedSeconds}s)" `
                                   -ForegroundColor DarkGray
                    }
                }
            }
        }
    }


    # ----------------------------------------------------------
    #  STEP 1 -- Environment Pre-Flight
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Environment Pre-Flight Check' -Step '1 of 7'

    $envResult = Test-OMMigrateEnvironment

    # Log Outlook version found
    if ($envResult.OutlookVersion) {
        Write-OMMigrateLog -Message "Outlook version: $($envResult.OutlookVersion)" `
                           -Level INFO
    }

    # Warn if running elevated -- usually wrong for Outlook profile work
    if (Test-AdminElevation) {
        Write-OMMigrateLog -Message (
            'WARNING: Script is running with Administrator elevation. ' +
            'Outlook profiles belong to the regular user account. ' +
            'If discovery finds no profiles, re-run without elevation.'
        ) -Level WARN
    }

    # Handle non-critical warnings
    foreach ($warn in $envResult.Warnings) {
        Write-OMMigrateLog -Message "Environment warning: $warn" -Level WARN
    }

    # Hard stop on critical failures
    if (-not $envResult.Passed) {
        Write-OMMigrateLog -Message 'Environment pre-flight failed. Cannot continue.' `
                           -Level ERROR
        foreach ($fail in $envResult.Failures) {
            Write-Host "  [FAIL] $fail" -ForegroundColor Red
        }
        $Script:FinalStatus = 'FAILED'
        exit 1
    }

    Write-OMMigrateLog -Message 'Environment pre-flight passed.' -Level INFO
    Write-Host '  All environment checks passed.' -ForegroundColor Green
    Write-Host ''


    # ----------------------------------------------------------
    #  STEP 2 -- Registry Profile Discovery
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Registry Profile Discovery' -Step '2 of 7'

    Write-OMMigrateLog -Message 'Scanning Windows Registry for Outlook profiles...' `
                       -Level INFO

    $Script:DiscoveredProfiles = @(Get-OutlookProfiles -ProfileName $ProfileName)

    if ($Script:DiscoveredProfiles.Count -eq 0) {
        Write-OMMigrateLog -Message (
            'No Outlook profiles found in registry. ' +
            'Verify Outlook is installed and has been configured for this user account.'
        ) -Level ERROR
        $Script:FinalStatus = 'FAILED'
        exit 1
    }

    Write-Host "  Profiles found: $($Script:DiscoveredProfiles.Count)" `
               -ForegroundColor Green

    foreach ($profile in $Script:DiscoveredProfiles) {
        $defaultTag = if ($profile.IsDefault) { ' [DEFAULT]' } else { '' }
        Write-Host "    * $($profile.Name)$defaultTag" -ForegroundColor Gray
    }
    Write-Host ''


    # ----------------------------------------------------------
    #  PROFILE SELECTION
    #  The operator must select which Outlook profile to use for
    #  all COM sessions and account discovery before any account
    #  scanning begins. This prevents secondary profiles (e.g.
    #  "TestProfile") from silently contaminating the discovery.
    #
    #  If only one profile exists, it is auto-selected and no
    #  picker is shown. In all cases the selection is persisted
    #  to OMMigrate_Settings.json so Scripts 01-04 can use it
    #  without re-prompting the operator.
    # ----------------------------------------------------------

    # Resolve selected profile -- command-line -ProfileName always wins
    $Script:SelectedProfileName = ''

    if ($ProfileName) {
        # Explicit -ProfileName parameter -- use it directly
        $Script:SelectedProfileName = $ProfileName
        Write-OMMigrateLog -Message "Profile specified via -ProfileName parameter: '$ProfileName'" `
                           -Level INFO
        Write-Host "  Profile: $ProfileName (specified via -ProfileName)" -ForegroundColor Green
        Write-Host ''
    }
    elseif ($Script:DiscoveredProfiles.Count -eq 1) {
        # Only one profile -- auto-select, no picker needed
        $Script:SelectedProfileName = $Script:DiscoveredProfiles[0].Name
        Write-OMMigrateLog -Message "Single profile found -- auto-selected: '$($Script:SelectedProfileName)'" `
                           -Level INFO
        Write-Host "  Profile auto-selected (only one found): $($Script:SelectedProfileName)" `
                   -ForegroundColor Green
        Write-Host ''
    }
    else {
        # Multiple profiles -- require operator selection via WinForms picker.
        # No default pre-selection -- operator must make an explicit choice.
        Write-OMMigrateLog -Message "Multiple profiles found -- presenting picker for operator selection." `
                           -Level INFO

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $pickerForm              = New-Object System.Windows.Forms.Form
        $pickerForm.Text         = 'OMMigrate -- Select Outlook Profile'
        $pickerForm.Size         = New-Object System.Drawing.Size(480, 280)
        $pickerForm.StartPosition = 'CenterScreen'
        $pickerForm.FormBorderStyle = 'FixedDialog'
        $pickerForm.MaximizeBox  = $false
        $pickerForm.MinimizeBox  = $false
        $pickerForm.TopMost      = $true

        $lblPrompt               = New-Object System.Windows.Forms.Label
        $lblPrompt.Text          = 'Select the Outlook profile to use for this migration session:'
        $lblPrompt.Location      = New-Object System.Drawing.Point(16, 16)
        $lblPrompt.Size          = New-Object System.Drawing.Size(440, 40)
        $pickerForm.Controls.Add($lblPrompt)

        $listBox                 = New-Object System.Windows.Forms.ListBox
        $listBox.Location        = New-Object System.Drawing.Point(16, 60)
        $listBox.Size            = New-Object System.Drawing.Size(440, 130)
        $listBox.SelectionMode   = 'One'
        foreach ($prof in $Script:DiscoveredProfiles) {
            [void]$listBox.Items.Add($prof.Name)
        }
        $pickerForm.Controls.Add($listBox)

        $btnOK                   = New-Object System.Windows.Forms.Button
        $btnOK.Text              = 'OK'
        $btnOK.Location          = New-Object System.Drawing.Point(140, 205)
        $btnOK.Size              = New-Object System.Drawing.Size(90, 28)
        $btnOK.Enabled           = $false   # disabled until a selection is made
        $btnOK.DialogResult      = [System.Windows.Forms.DialogResult]::OK
        $pickerForm.Controls.Add($btnOK)
        $pickerForm.AcceptButton = $btnOK

        # Added 2026-07-10, Administrator direction: Cancel button added for
        # consistency with the other three OMMigrate pickers
        # (Invoke-FolderMapPicker, Invoke-RulesInventoryPicker,
        # Invoke-MigrateAccountPicker), which all already have an explicit
        # Cancel button alongside OK. Previously this picker (and the
        # TargetStoreName picker) relied on the window's X button only --
        # functionally correct, since ShowDialog() already returns non-OK
        # on X and the caller below already handles a non-OK result
        # correctly, but inconsistent with the rest of the tool's pickers.
        # Same style as the other three: Segoe UI 9, 80x28,
        # DialogResult.Cancel, form.CancelButton.
        $btnCancel               = New-Object System.Windows.Forms.Button
        $btnCancel.Text          = 'Cancel'
        $btnCancel.Font          = New-Object System.Drawing.Font('Segoe UI', 9)
        $btnCancel.Location      = New-Object System.Drawing.Point(240, 205)
        $btnCancel.Size          = New-Object System.Drawing.Size(80, 28)
        $btnCancel.DialogResult  = [System.Windows.Forms.DialogResult]::Cancel
        $pickerForm.CancelButton = $btnCancel
        $pickerForm.Controls.Add($btnCancel)

        # Enable OK only when a profile is selected
        $listBox.Add_SelectedIndexChanged({
            $btnOK.Enabled = ($listBox.SelectedIndex -ge 0)
        })

        $pickerResult = $pickerForm.ShowDialog()
        $capturedProfile = $listBox.SelectedItem
        $pickerForm.Dispose()

        if ($pickerResult -eq [System.Windows.Forms.DialogResult]::OK -and $capturedProfile) {
            $Script:SelectedProfileName = $capturedProfile
            Write-OMMigrateLog -Message "Operator selected profile: '$($Script:SelectedProfileName)'" `
                               -Level INFO
            Write-Host "  Profile selected: $($Script:SelectedProfileName)" -ForegroundColor Green
            Write-Host ''
        }
        else {
            Write-OMMigrateLog -Message "Profile picker cancelled -- cannot continue without a profile selection." `
                               -Level ERROR
            Write-Host ''
            Write-Host '  Profile selection is required. Script 00 cannot continue.' `
                       -ForegroundColor Red
            Write-Host '  Re-run Script 00 and select a profile to proceed.' -ForegroundColor Red
            Write-Host ''
            $Script:FinalStatus = 'FAILED'
            exit 1
        }
    }

    # Added 2026-07-10, Administrator direction. Live-tested problem found:
    # RulesEngine settings (ArchiveStoreMappings, MasterArchiveNames) are
    # profile-specific -- switching profiles was silently applying one
    # profile's archive/account mappings to a different profile, since
    # OMMigrate_Settings.json was a single flat file shared by all
    # profiles. Switch-OMMigrateProfileSettings (OMMigrate-Core.psm1) swaps
    # in this profile's own settings file (creating it on first use) BEFORE
    # Save-OMMigrateSelectedProfile writes the profile name -- so the
    # profile name lands in the correct per-profile file as a natural side
    # effect, with no changes needed to Save-OMMigrateSelectedProfile
    # itself. Every other script/function in this codebase continues
    # reading/writing the plain OMMigrate_Settings.json filename unchanged.
    if (-not $Script:IsWhatIf) {
        Switch-OMMigrateProfileSettings -ProfileName $Script:SelectedProfileName
    }

    # Persist the selected profile to Settings.json so Scripts 01-04
    # can use it without re-prompting the operator.
    if (-not $Script:IsWhatIf) {
        Save-OMMigrateSelectedProfile -ProfileName $Script:SelectedProfileName
    }


    # ----------------------------------------------------------
    #  STEP 3 -- Account Discovery from Registry
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Account Discovery -- Registry' -Step '3 of 7'

    # Scope discovery to the selected profile only -- scanning all profiles
    # would pull in accounts from secondary profiles (e.g. "TestProfile") and
    # contaminate the CSV with duplicate and irrelevant accounts.
    $Script:ProfilesToScan = @($Script:DiscoveredProfiles |
        Where-Object { $_.Name -eq $Script:SelectedProfileName })

    if ($Script:ProfilesToScan.Count -eq 0) {
        Write-OMMigrateLog -Message (
            "Selected profile '$($Script:SelectedProfileName)' not found in discovered profiles -- " +
            "falling back to all profiles."
        ) -Level WARN
        $Script:ProfilesToScan = $Script:DiscoveredProfiles
    }

    foreach ($profile in $Script:ProfilesToScan) {

        Write-OMMigrateLog -Message "Processing profile: '$($profile.Name)'" `
                           -Level INFO

        # Discover accounts from registry
        $profileAccounts = Get-OutlookAccountsFromRegistry -Profile $profile

        # Discover data files (PST/OST) for this profile
        $dataFiles = Get-OutlookDataFiles -Profile $profile

        # Match data files to accounts
        $joinResult = Join-AccountsWithDataFiles `
                          -Accounts  $profileAccounts `
                          -DataFiles $dataFiles

        # Tag each account with its profile name
        foreach ($account in $joinResult.Accounts) {
            $account.ProfileName    = $profile.Name
            $account.DiscoveredAt  = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
            $account.RegistryPath  = $profile.RegistryPath
            [void]$Script:AllAccounts.Add($account)
        }

        # Collect unmatched data files for report warning
        foreach ($file in $joinResult.UnmatchedFiles) {
            [void]$Script:AllDataFiles.Add($file)
        }
    }

    # -- Build sanitization map from registry accounts -----------
    # Must be done BEFORE any account data is displayed so that
    # all subsequent output is filtered through the sanitize map.
    if ($Global:OMMigrate.Sanitize) {
        Initialize-SanitizeMap -Accounts @($Script:AllAccounts)
        Write-OMMigrateLog -Message '[SANITIZE] Account sanitization map built.' `
                           -Level INFO
    }

    # -- Now display accounts and log data file info (map is ready) -
    foreach ($profile in $Script:ProfilesToScan) {
        foreach ($account in $Script:AllAccounts | Where-Object { $_.ProfileName -eq $profile.Name }) {

            # Log data file matches
            if ($account.PSTPath -and (Test-Path $account.PSTPath -ErrorAction SilentlyContinue)) {
                Write-OMMigrateLog -Message (
                    "Data file (from registry): $($account.PSTPath) | " +
                    "Account: $($account.EmailAddress)"
                ) -Level INFO
            }
            elseif ($account.OSTPath -and (Test-Path $account.OSTPath -ErrorAction SilentlyContinue)) {
                Write-OMMigrateLog -Message (
                    "Data file (from registry): $($account.OSTPath) | " +
                    "Account: $($account.EmailAddress)"
                ) -Level INFO
            }

            # Display per-account status on console
            $action = switch ($account.MigrationAction) {
                'MIGRATE'     { 'Will migrate POP3 -> IMAP' }
                'FOLDER-ONLY' { 'Already IMAP -- folder assessment only' }
                'SKIP'        { 'Exchange -- no action required' }
                'MANUAL'      { 'Requires manual review' }
                default       { $account.MigrationAction }
            }

            $status = switch ($account.MigrationAction) {
                'MIGRATE'     { 'INFO' }
                'FOLDER-ONLY' { 'OK'   }
                'SKIP'        { 'SKIP' }
                'MANUAL'      { 'WARN' }
                default       { 'INFO' }
            }

            Show-AccountStatus -Email  (Invoke-OMMigrateSanitize -Text $account.EmailAddress) `
                               -Tag    $account.ProviderTag `
                               -Action $action `
                               -Status $status
        }
    }

    Write-Host ''
    Write-OMMigrateLog -Message "Total accounts discovered: $($Script:AllAccounts.Count)" `
                       -Level INFO

    if ($Script:AllAccounts.Count -eq 0) {
        Write-OMMigrateLog -Message (
            'No email accounts discovered. The Outlook profile may be empty ' +
            'or account data may be stored in a format not readable from the registry.'
        ) -Level WARN
        $Script:FinalStatus = 'WARNING'
    }


    # ----------------------------------------------------------
    #  STEP 4 -- COM Corroboration and Folder/Rules Enumeration
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Outlook COM Session -- Folder and Rules Discovery' `
                       -Step '4 of 7'

    Write-OMMigrateLog -Message 'Starting Outlook COM session for folder and rules discovery...' `
                       -Level INFO
    Write-Host '  Launching Outlook COM session...' -ForegroundColor Cyan
    Write-Host '  (Outlook will open briefly -- do not interact with it)' `
               -ForegroundColor DarkGray
    Write-Host ''

    # Use the profile selected by the operator in the picker above.
    # $Script:SelectedProfileName is always set before this point --
    # either via -ProfileName parameter, auto-selection (single profile),
    # or the WinForms picker. If somehow empty, fall back to default.
    $comProfile = if ($Script:SelectedProfileName) {
        $Script:SelectedProfileName
    }
    else {
        $defaultProf = $Script:DiscoveredProfiles |
                       Where-Object { $_.IsDefault } |
                       Select-Object -First 1
        if ($defaultProf) { $defaultProf.Name } else {
            ($Script:DiscoveredProfiles | Select-Object -First 1).Name
        }
    }

    $outlook = Connect-OutlookCOM -ProfileName $comProfile

    if ($outlook) {
        $Script:COMSessionOpen = $true

        # -- COM account corroboration ----------------------
        Write-OMMigrateLog -Message 'Corroborating registry accounts with COM accounts...' `
                           -Level INFO

        $comAccounts = Get-OutlookAccountsViaCOM

            # Cross-reference COM accounts with registry accounts
            # COM is authoritative for account type -- update registry findings
            foreach ($comAccount in $comAccounts) {
                $registryMatch = $Script:AllAccounts |
                    Where-Object {
                        $_.EmailAddress -eq $comAccount.EmailAddress
                    } | Select-Object -First 1

                if ($registryMatch) {
                    # COM is authoritative -- always update account type
                    if ($comAccount.AccountType -ne 'Other') {
                        $registryMatch.AccountType = $comAccount.AccountType
                        # Re-classify with corrected type
                        $reclassified = Set-AccountTag -Account $registryMatch
                        $registryMatch.ProviderTag     = $reclassified.ProviderTag
                        $registryMatch.MigrationAction = $reclassified.MigrationAction
                        $registryMatch.Notes           = $reclassified.Notes

                        Write-OMMigrateLog -Message (
                            "COM corroboration updated account type for " +
                            "$($registryMatch.EmailAddress): $($registryMatch.AccountType)"
                        ) -Level INFO
                    }

                    # COM is authoritative for server names -- overwrite registry values
                    if ($comAccount.IncomingServer) {
                        $registryMatch.IncomingServer = $comAccount.IncomingServer
                    }
                    if ($comAccount.OutgoingServer) {
                        $registryMatch.OutgoingServer = $comAccount.OutgoingServer
                    }
                    if ($comAccount.IncomingPort -gt 0) {
                        $registryMatch.IncomingPort = $comAccount.IncomingPort
                    }
                    if ($comAccount.OutgoingPort -gt 0) {
                        $registryMatch.OutgoingPort = $comAccount.OutgoingPort
                    }
                    $registryMatch.IncomingSSL = $comAccount.IncomingSSL
                    $registryMatch.OutgoingSSL = $comAccount.OutgoingSSL

                    # COM is authoritative for PST/OST file paths
                    if ($comAccount.FilePath) {
                        $ext = [System.IO.Path]::GetExtension($comAccount.FilePath).ToLower()
                        if ($ext -eq '.pst') {
                            $registryMatch.PSTPath = $comAccount.FilePath
                            # Update file size if not already set
                            if ($registryMatch.DataFileSizeBytes -eq 0 -and
                                (Test-Path $comAccount.FilePath)) {
                                $fileInfo = Get-Item $comAccount.FilePath
                                $registryMatch.DataFileSizeBytes = $fileInfo.Length
                                $registryMatch.DataFileSizeFormatted = Format-FileSize -Bytes $fileInfo.Length
                            }
                        }
                        elseif ($ext -eq '.ost') {
                            $registryMatch.OSTPath = $comAccount.FilePath
                            if ($registryMatch.DataFileSizeBytes -eq 0 -and
                                (Test-Path $comAccount.FilePath)) {
                                $fileInfo = Get-Item $comAccount.FilePath
                                $registryMatch.DataFileSizeBytes = $fileInfo.Length
                                $registryMatch.DataFileSizeFormatted = Format-FileSize -Bytes $fileInfo.Length
                            }
                        }
                    }

                    Write-OMMigrateLog -Message (
                        "COM enriched: $($registryMatch.EmailAddress) | " +
                        "Server=$($registryMatch.IncomingServer) | " +
                        "PST=$($registryMatch.PSTPath)"
                    ) -Level DEBUG
                }
                else {
                    # Account in COM but not found in registry -- add it
                    # Register the email in sanitize map BEFORE logging
                    if ($Global:OMMigrate.Sanitize -and $comAccount.EmailAddress) {
                        Register-SanitizeTerms -Terms @($comAccount.EmailAddress) -Category 'Email'
                        if ($comAccount.AccountName) {
                            Register-SanitizeTerms -Terms @($comAccount.AccountName) -Category 'Display'
                        }
                    }
                    Write-OMMigrateLog -Message (
                        "COM account not found in registry discovery: " +
                        "$($comAccount.EmailAddress) [$($comAccount.AccountType)] -- adding."
                    ) -Level INFO

                    $newAccount = New-AccountObject
                    $newAccount.EmailAddress    = $comAccount.EmailAddress
                    $newAccount.DisplayName     = $comAccount.AccountName
                    $newAccount.AccountType     = $comAccount.AccountType
                    $newAccount.ProfileName     = $comProfile
                    $newAccount.DiscoveredAt    = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
                    $newAccount.Notes           = 'Found via COM -- not in registry discovery'
                    $newAccount.IncomingServer  = $comAccount.IncomingServer
                    $newAccount.IncomingPort    = $comAccount.IncomingPort
                    $newAccount.OutgoingServer  = $comAccount.OutgoingServer
                    $newAccount.OutgoingPort    = $comAccount.OutgoingPort
                    $classified = Set-AccountTag -Account $newAccount
                    [void]$Script:AllAccounts.Add($classified)
                }
            }

            # -- OST path fallback for COM-only accounts -------
            # Some IMAP accounts added via the manual Add Account dialog
            # do not get registry entries written by Outlook. These accounts
            # are discovered via COM but Join-AccountsWithDataFiles cannot
            # match them because they have no registry representation.
            # Scan the standard Outlook OST folder and match by email address
            # in the filename for any IMAP account still missing an OSTPath.
            $ostFolder = [System.IO.Path]::Combine(
                $env:LOCALAPPDATA, 'Microsoft', 'Outlook'
            )
            if (Test-Path $ostFolder) {
                $ostFiles = Get-ChildItem -Path $ostFolder -Filter '*.ost' `
                                          -ErrorAction SilentlyContinue
                foreach ($account in $Script:AllAccounts) {
                    if ($account.OSTPath) { continue }                 # already matched
                    if ($account.AccountType -notin @('IMAP','Exchange')) { continue }

                    foreach ($ostFile in $ostFiles) {
                        $fileNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($ostFile.Name)
                        if ($fileNameNoExt -like "*$($account.EmailAddress)*") {
                            $account.OSTPath               = $ostFile.FullName
                            $account.DataFileSizeBytes     = $ostFile.Length
                            $account.DataFileSizeFormatted = Format-FileSize -Bytes $ostFile.Length
                            Write-OMMigrateLog -Message (
                                "OST fallback matched: $($account.EmailAddress) -> $($ostFile.FullName)"
                            ) -Level INFO
                            break
                        }
                    }
                }
            }

            # -- Standalone mounted PST discovery --------------
            # Registry account discovery only finds accounts with registry
            # entries. Standalone attached PST stores (e.g. archive PSTs,
            # OST->PST backups manually attached, user@example.com
            # Archive) have no registry account entry and are invisible to
            # Get-OutlookAccountsFromRegistry. Their folders ARE enumerated
            # by Get-FolderTree (it scans all $namespace.Stores), and their
            # rules ARE read by Get-OutlookRules -- but without an account
            # entry in $Script:AllAccounts, rule target resolution and
            # migration_accounts.csv have no row for these stores.
            #
            # Fix: scan all mounted stores, find PST stores whose DisplayName
            # does not match any existing account, and add a synthetic
            # PST-ARCHIVE account entry for each one.
            #
            # ADDENDUM (2026-06-27): the "Their folders ARE enumerated by
            # Get-FolderTree" line above no longer applies to stores named
            # 'Backup -- *' or 'ArchiveBuild -- *' -- Administrator's manually attached
            # backup/snapshot PSTs. Get-FolderTree now skips those by design
            # (see -ExcludeBackupPSTs), and this discovery loop skips adding
            # a PST-ARCHIVE row for them too, since a row with no folders
            # underneath serves no purpose. Real working archive PSTs
            # (e.g. user@example.com Archive) are unaffected --
            # they don't match the Backup --/ArchiveBuild -- prefix pattern.
            try {
                $mountedNS     = Get-OutlookNamespace
                if (-not $mountedNS) {
                    Write-OMMigrateLog -Message "Standalone PST discovery: no active COM session -- skipping." -Level DEBUG
                    return
                }
                $mountedStores = $mountedNS.Stores
                $knownStoreNames = [System.Collections.Generic.HashSet[string]]::new(
                    [System.StringComparer]::OrdinalIgnoreCase
                )
                # Build set of known store names from existing accounts
                foreach ($acct in $Script:AllAccounts) {
                    if ($acct.EmailAddress) {
                        [void]$knownStoreNames.Add($acct.EmailAddress)
                    }
                    if ($acct.DisplayName) {
                        [void]$knownStoreNames.Add($acct.DisplayName)
                    }
                }

                for ($ms = 1; $ms -le $mountedStores.Count; $ms++) {
                    $mountedStore = $null
                    try {
                        $mountedStore = $mountedStores.Item($ms)
                        $msName     = $mountedStore.DisplayName
                        $msType     = Get-StoreType -Store $mountedStore
                        $msFilePath = ''
                        try { $msFilePath = $mountedStore.FilePath } catch { }

                        # Only interested in PST stores not already known
                        if ($msType -ne 'PST') { continue }
                        if ($knownStoreNames.Contains($msName)) { continue }

                        # Skip manually attached backup/archive PSTs (e.g.
                        # 'Backup -- ameritech', 'ArchiveBuild -- ameritech') --
                        # these are disposable safety copies Administrator attaches to
                        # the nav pane, not migration targets. A synthetic
                        # PST-ARCHIVE row for them would have no folders
                        # underneath (Get-FolderTree skips them too) and would
                        # just be dead weight in migration_accounts.csv.
                        if ($msName -and
                            ($msName -like 'Backup --*' -or $msName -like 'ArchiveBuild --*')) {
                            Write-OMMigrateLog -Message (
                                "Standalone PST discovery: skipping manually attached backup/archive PST: '$msName'."
                            ) -Level DEBUG
                            continue
                        }

                        # This is a standalone mounted PST -- add synthetic account
                        $archiveAccount                    = New-AccountObject
                        $archiveAccount.DisplayName        = $msName
                        $archiveAccount.EmailAddress       = 'N/A'     # not a mail account; store name is in DisplayName
                        $archiveAccount.AccountType        = 'Unknown'
                        $archiveAccount.ProviderTag        = 'PST-ARCHIVE'
                        $archiveAccount.MigrationAction    = 'FOLDER-ONLY'
                        $archiveAccount.ProfileName        = $comProfile
                        $archiveAccount.DiscoveredAt       = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
                        $archiveAccount.Notes              = 'Standalone mounted PST -- no registry account entry. ' +
                                                             'Included for folder map and rule target resolution.'
                        $archiveAccount.PSTPath            = $msFilePath
                        if ($msFilePath -and (Test-Path $msFilePath -ErrorAction SilentlyContinue)) {
                            $msFileInfo = Get-Item $msFilePath -ErrorAction SilentlyContinue
                            if ($msFileInfo) {
                                $archiveAccount.DataFileSizeBytes     = $msFileInfo.Length
                                $archiveAccount.DataFileSizeFormatted = Format-FileSize -Bytes $msFileInfo.Length
                            }
                        }

                        [void]$Script:AllAccounts.Add($archiveAccount)
                        [void]$knownStoreNames.Add($msName)

                        Write-OMMigrateLog -Message (
                            "Standalone PST store added: '$msName' | " +
                            "Path=$msFilePath"
                        ) -Level INFO
                        Write-Host "  Standalone PST found: $msName" -ForegroundColor Cyan
                    }
                    catch {
                        Write-OMMigrateLog -Message "Error checking mounted store [$ms] for standalone PST: $_" `
                                           -Level DEBUG
                    }
                }
            }
            catch {
                Write-OMMigrateLog -Message "Standalone PST discovery failed (non-fatal): $_" -Level WARN
            }

            # EXTRACTED 2026-07-09 (Administrator direction) from OMMigrate-00-Discover_WIP.ps1,
            # a 6-day-old (2026-07-03/04) work-in-progress file -- this is the
            # TargetStoreName hardcode fix work paused mid-implementation that
            # session. NOT a completed/finished feature -- extracted as a
            # jumpstart on resuming this work, not yet live-tested.
            #
            # CROSS-FILE DEPENDENCY WARNING: this block calls
            # Save-OMMigrateArchiveStoreMappings, which currently only exists
            # in OMMigrate-Core_WIP.psm1 -- NOT in the production
            # OMMigrate-Core.psm1 that this script actually imports (see
            # $Script:RequiredModules above). Running this script as-is right
            # now would throw "command not found" the moment this block tries
            # to call that function. This is expected and correct for the
            # current WIP state -- do NOT treat this as a bug to silently
            # patch around. This picker must not be tested/run until
            # OMMigrate-Core_WIP.psm1 is reviewed, tested, and promoted to
            # replace the production OMMigrate-Core.psm1 (per the active
            # feature freeze).
            #
            # -- TargetStoreName picker --------------------------
            # Lets the operator/admin decide which attached archive PST(s)
            # each account's rules should target. Shown every run (not just
            # first-run) so the admin can maintain this mapping over time --
            # defaults to the last saved selection so a normal run only
            # requires clicking OK. Both the PST list and the account list
            # are built entirely from what was just live-detected above
            # ($Script:AllAccounts, including the synthetic PST-ARCHIVE rows
            # from standalone PST discovery) -- no store or account name is
            # ever hardcoded here or in Export-RulesToCSV's consumption of
            # the resulting mapping.
            try {
                $pstChoices = @($Script:AllAccounts |
                    Where-Object { $_.ProviderTag -eq 'PST-ARCHIVE' -and $_.DisplayName } |
                    Select-Object -ExpandProperty DisplayName -Unique)
                $acctChoices = @($Script:AllAccounts |
                    Where-Object { $_.ProviderTag -ne 'PST-ARCHIVE' -and $_.EmailAddress -and $_.EmailAddress -ne 'N/A' } |
                    Select-Object -ExpandProperty EmailAddress -Unique)

                if ($pstChoices.Count -eq 0) {
                    Write-OMMigrateLog -Message 'TargetStoreName picker: no attached archive PSTs detected -- skipping picker.' -Level INFO
                }
                else {
                    # Load prior mapping (if any) to pre-populate the picker.
                    $priorMappings = @()
                    if ($Global:OMMigrate.Settings.PSObject.Properties['RulesEngine'] -and
                        $Global:OMMigrate.Settings.RulesEngine.PSObject.Properties['ArchiveStoreMappings']) {
                        $priorMappings = @($Global:OMMigrate.Settings.RulesEngine.ArchiveStoreMappings)
                    }

                    Add-Type -AssemblyName System.Windows.Forms
                    Add-Type -AssemblyName System.Drawing

                    $tsnForm               = New-Object System.Windows.Forms.Form
                    $tsnForm.Text          = 'OMMigrate -- Select Rule Target Archive(s)'
                    $tsnForm.Size          = New-Object System.Drawing.Size(660, 600)
                    $tsnForm.StartPosition = 'CenterScreen'
                    $tsnForm.FormBorderStyle = 'FixedDialog'
                    $tsnForm.MaximizeBox   = $false
                    $tsnForm.MinimizeBox   = $false
                    $tsnForm.TopMost       = $true

                    $tsnPrompt             = New-Object System.Windows.Forms.Label
                    # Corrected 2026-07-10, Administrator direction: the prior wording ("will
                    # default to targeting its own store") was inaccurate for
                    # Local-destination rules -- confirmed via full code trace
                    # (Export-RulesToCSV / Invoke-DeployConsolidatedRules /
                    # Invoke-RulesRecreation / Strategy 1+2 remap) that an
                    # unmapped account's Local-destination rules (the majority
                    # of Administrator's own rules, per his direct confirmation) actually
                    # fall back to the single default/shared Archive PST, not to
                    # the account's own store. Only Server-destination rules
                    # (a minority) ever resolve against the account's own store,
                    # and TargetStoreName is never even consulted for those.
                    $tsnPrompt.Text        = 'For each attached archive PST, check the account(s) whose rules should target it. ' +
                                             'An account not checked under any PST will have its Local-destination rules default ' +
                                             'to the shared Archive PST (Server-destination rules are unaffected).'
                    $tsnPrompt.Location    = New-Object System.Drawing.Point(16, 12)
                    $tsnPrompt.Size        = New-Object System.Drawing.Size(620, 45)
                    $tsnForm.Controls.Add($tsnPrompt)

                    $tsnScroll             = New-Object System.Windows.Forms.Panel
                    $tsnScroll.Location    = New-Object System.Drawing.Point(16, 64)
                    # Added 2026-07-10, Administrator direction: shrunk from 420 to 396 to make
                    # room for the live unmapped-account status label below, without
                    # resizing the form or moving the OK button.
                    $tsnScroll.Size        = New-Object System.Drawing.Size(620, 396)
                    $tsnScroll.AutoScroll  = $true
                    $tsnScroll.BorderStyle = 'FixedSingle'
                    $tsnForm.Controls.Add($tsnScroll)

                    # Added 2026-07-10, Administrator direction: live status label showing how
                    # many accounts are currently unmapped (checked in no PST panel)
                    # and what will happen to their rules -- updated on every checkbox
                    # change via ItemCheck below, so the operator sees the real
                    # consequence of an unmapped account BEFORE clicking OK, rather
                    # than only discovering it later in a Script 03 WARN log line.
                    $tsnUnmappedStatus            = New-Object System.Windows.Forms.Label
                    $tsnUnmappedStatus.Location   = New-Object System.Drawing.Point(16, 464)
                    $tsnUnmappedStatus.Size       = New-Object System.Drawing.Size(620, 28)
                    $tsnUnmappedStatus.ForeColor  = [System.Drawing.Color]::DarkOrange
                    $tsnForm.Controls.Add($tsnUnmappedStatus)

                    # One panel per PST, each containing a label and a
                    # CheckedListBox of every known account's DisplayName.
                    $tsnCheckLists = @{}
                    $panelY = 8
                    foreach ($pstName in $pstChoices) {
                        $pstLabel            = New-Object System.Windows.Forms.Label
                        $pstLabel.Text       = $pstName
                        $pstLabel.Font       = New-Object System.Drawing.Font($pstLabel.Font, [System.Drawing.FontStyle]::Bold)
                        $pstLabel.Location   = New-Object System.Drawing.Point(8, $panelY)
                        $pstLabel.Size       = New-Object System.Drawing.Size(480, 20)
                        $tsnScroll.Controls.Add($pstLabel)
                        $panelY += 22

                        $rowHeight = [Math]::Min(110, 20 + ($acctChoices.Count * 16))
                        $chkList             = New-Object System.Windows.Forms.CheckedListBox
                        $chkList.Location    = New-Object System.Drawing.Point(8, $panelY)
                        $chkList.Size        = New-Object System.Drawing.Size(480, $rowHeight)
                        $chkList.CheckOnClick = $true
                        foreach ($acctName in $acctChoices) {
                            [void]$chkList.Items.Add($acctName)
                        }
                        # Pre-check accounts already mapped to this PST from the prior save.
                        $priorEntry = $priorMappings | Where-Object { $_.TargetStoreName -eq $pstName } | Select-Object -First 1
                        if ($priorEntry -and $priorEntry.RuleStoreNames) {
                            for ($i = 0; $i -lt $chkList.Items.Count; $i++) {
                                if (@($priorEntry.RuleStoreNames) -contains $chkList.Items[$i]) {
                                    $chkList.SetItemChecked($i, $true)
                                }
                            }
                        }
                        $tsnScroll.Controls.Add($chkList)
                        $tsnCheckLists[$pstName] = $chkList

                        # Added -- Select All / Unselect All buttons for this PST's
                        # CheckedListBox, per Administrator direction (Next Steps item 1: with
                        # 24 real accounts, manually checking each one individually on
                        # first use of a profile is tedious). Placed to the right of
                        # each PST's CheckedListBox rather than adding a new row, so no
                        # existing layout math ($panelY increments) needs to change.
                        # Looping SetItemChecked fires the existing ItemCheck handler
                        # once per item, which already recomputes the unmapped-account
                        # status label -- no separate status-refresh call is needed here.
                        $tsnBtnSelectAll           = New-Object System.Windows.Forms.Button
                        $tsnBtnSelectAll.Text      = 'Select All'
                        $tsnBtnSelectAll.Font      = New-Object System.Drawing.Font('Segoe UI', 7.5)
                        $tsnBtnSelectAll.Location  = New-Object System.Drawing.Point(494, $panelY)
                        $tsnBtnSelectAll.Size      = New-Object System.Drawing.Size(90, 22)
                        $tsnBtnSelectAll.Add_Click({
                            param($evtSender, $evtArgs)
                            for ($i = 0; $i -lt $chkList.Items.Count; $i++) {
                                $chkList.SetItemChecked($i, $true)
                            }
                        }.GetNewClosure())
                        $tsnScroll.Controls.Add($tsnBtnSelectAll)

                        $tsnBtnUnselectAll             = New-Object System.Windows.Forms.Button
                        $tsnBtnUnselectAll.Text        = 'Unselect All'
                        $tsnBtnUnselectAll.Font        = New-Object System.Drawing.Font('Segoe UI', 7.5)
                        $tsnBtnUnselectAll.Location    = New-Object System.Drawing.Point(494, ($panelY + 26))
                        $tsnBtnUnselectAll.Size        = New-Object System.Drawing.Size(90, 22)
                        $tsnBtnUnselectAll.Add_Click({
                            param($evtSender, $evtArgs)
                            for ($i = 0; $i -lt $chkList.Items.Count; $i++) {
                                $chkList.SetItemChecked($i, $false)
                            }
                        }.GetNewClosure())
                        $tsnScroll.Controls.Add($tsnBtnUnselectAll)

                        $panelY += $rowHeight + 16
                    }

                    # Added 2026-07-10, Administrator direction: identify which attached PST
                    # (if any) is the actual default/fallback Archive PST every
                    # unmapped account's Local-destination rules will land in, so
                    # the status message can name it specifically instead of only
                    # describing it generically. Matched by PSTPath (captured on
                    # every synthetic PST-ARCHIVE row during standalone PST
                    # discovery above) against the same hardcoded default path
                    # Script 01/03/Install.ps1 all use -- not a new hardcode, just
                    # reading the same known constant this feature already
                    # depends on elsewhere. If that specific PST is not currently
                    # attached (not in $Script:AllAccounts's PST-ARCHIVE rows this
                    # run), $tsnDefaultArchiveDisplayName stays $null and the
                    # message below falls back to the generic wording -- never
                    # guesses at a name that cannot be confirmed live.
                    $tsnDefaultArchivePSTPath = Join-Path $Global:OMMigrate.BackupPath 'OMMigrate_Archive.pst'
                    $tsnDefaultArchiveDisplayName = $null
                    $tsnDefaultArchiveMatch = $Script:AllAccounts | Where-Object {
                        $_.ProviderTag -eq 'PST-ARCHIVE' -and
                        $_.PSTPath -and
                        $_.PSTPath -eq $tsnDefaultArchivePSTPath
                    } | Select-Object -First 1
                    if ($tsnDefaultArchiveMatch) {
                        $tsnDefaultArchiveDisplayName = $tsnDefaultArchiveMatch.DisplayName
                    }

                    # Recompute and display the live unmapped-account count and
                    # consequence message every time any checkbox in any PST panel
                    # changes. ItemCheck fires BEFORE the checkbox's own state
                    # actually updates, so the box being changed right now is
                    # corrected using $_.NewValue rather than its still-stale
                    # CheckState -- every other box's current state is read
                    # normally since ItemCheck only reports on the one box that
                    # triggered the event.
                    $tsnUpdateUnmappedStatus = {
                        param($changingList, $changingIndex, $changingNewValue)

                        $checkedAnywhere = [System.Collections.Generic.HashSet[string]]::new(
                            [System.StringComparer]::OrdinalIgnoreCase
                        )
                        foreach ($pstNameForCount in $tsnCheckLists.Keys) {
                            $listForCount = $tsnCheckLists[$pstNameForCount]
                            for ($ci = 0; $ci -lt $listForCount.Items.Count; $ci++) {
                                $isCheckedNow = $listForCount.GetItemChecked($ci)
                                # Override with the pending new value for the one box
                                # currently being toggled, since GetItemChecked still
                                # reflects the PRE-click state at the moment ItemCheck fires.
                                if ($null -ne $changingList -and
                                    [object]::ReferenceEquals($listForCount, $changingList) -and
                                    $ci -eq $changingIndex) {
                                    $isCheckedNow = ($changingNewValue -eq [System.Windows.Forms.CheckState]::Checked)
                                }
                                if ($isCheckedNow) {
                                    [void]$checkedAnywhere.Add([string]$listForCount.Items[$ci])
                                }
                            }
                        }

                        $unmappedCount = 0
                        foreach ($acctNameForCount in $acctChoices) {
                            if (-not $checkedAnywhere.Contains($acctNameForCount)) { $unmappedCount++ }
                        }

                        if ($unmappedCount -eq 0) {
                            $tsnUnmappedStatus.Text = 'All accounts are mapped to an archive PST.'
                            $tsnUnmappedStatus.ForeColor = [System.Drawing.Color]::DarkGreen
                        }
                        else {
                            # Name the actual default archive when it's confirmed
                            # attached this run; otherwise describe it generically
                            # rather than guessing at a name that can't be verified.
                            if ($tsnDefaultArchiveDisplayName) {
                                $tsnUnmappedStatus.Text = (
                                    "$unmappedCount account(s) unmapped -- their Local-destination rules " +
                                    "will default to '$tsnDefaultArchiveDisplayName' when this run's rules are processed."
                                )
                            }
                            else {
                                $tsnUnmappedStatus.Text = (
                                    "$unmappedCount account(s) unmapped -- their Local-destination rules will " +
                                    "default to the shared Archive PST (not currently attached -- name unconfirmed)."
                                )
                            }
                            $tsnUnmappedStatus.ForeColor = [System.Drawing.Color]::DarkOrange
                        }
                    }

                    foreach ($pstNameForHandler in $tsnCheckLists.Keys) {
                        $tsnCheckLists[$pstNameForHandler].Add_ItemCheck({
                            param($evtSender, $evtArgs)
                            & $tsnUpdateUnmappedStatus $evtSender $evtArgs.Index $evtArgs.NewValue
                        }.GetNewClosure())
                    }

                    # Initial computation so the label is correct on first display,
                    # accounting for any pre-checked accounts from a prior saved mapping.
                    & $tsnUpdateUnmappedStatus $null 0 $null

                    # Added 2026-07-10, Administrator direction: small static note clarifying
                    # that Cancel does not abort the run -- it only skips the
                    # mapping update. Placed as a separate label from
                    # $tsnUnmappedStatus above since that label's text is
                    # dynamic (unmapped-account count) and already carries a
                    # different message; this one is a fixed clarification
                    # about what Cancel specifically does, always visible
                    # regardless of mapping state.
                    $tsnCancelNote             = New-Object System.Windows.Forms.Label
                    $tsnCancelNote.Text        = 'Cancel does not stop this script -- it only skips updating the archive mapping; the run continues normally.'
                    $tsnCancelNote.Location    = New-Object System.Drawing.Point(16, 494)
                    $tsnCancelNote.Size        = New-Object System.Drawing.Size(620, 18)
                    $tsnCancelNote.Font        = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Italic)
                    $tsnCancelNote.ForeColor   = [System.Drawing.Color]::Gray
                    $tsnForm.Controls.Add($tsnCancelNote)

                    $tsnBtnOK              = New-Object System.Windows.Forms.Button
                    $tsnBtnOK.Text         = 'OK'
                    $tsnBtnOK.Location     = New-Object System.Drawing.Point(230, 520)
                    $tsnBtnOK.Size         = New-Object System.Drawing.Size(90, 28)
                    $tsnBtnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $tsnForm.Controls.Add($tsnBtnOK)
                    $tsnForm.AcceptButton  = $tsnBtnOK

                    # Added 2026-07-10, Administrator direction: Cancel button added for
                    # consistency with the other three OMMigrate pickers
                    # (Invoke-FolderMapPicker, Invoke-RulesInventoryPicker,
                    # Invoke-MigrateAccountPicker), which all already have an
                    # explicit Cancel button alongside OK. Previously this
                    # picker (and the profile picker) relied on the window's X
                    # button only for cancel -- functionally correct, since
                    # ShowDialog() already returns non-OK on X per the existing
                    # cancel-branch logic below, but inconsistent with the rest
                    # of the tool's pickers. Same style as the other three:
                    # Segoe UI 9, 80x28, DialogResult.Cancel, form.CancelButton.
                    $tsnBtnCancel              = New-Object System.Windows.Forms.Button
                    $tsnBtnCancel.Text         = 'Cancel'
                    $tsnBtnCancel.Font         = New-Object System.Drawing.Font('Segoe UI', 9)
                    $tsnBtnCancel.Location     = New-Object System.Drawing.Point(330, 520)
                    $tsnBtnCancel.Size         = New-Object System.Drawing.Size(80, 28)
                    $tsnBtnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                    $tsnForm.CancelButton      = $tsnBtnCancel
                    $tsnForm.Controls.Add($tsnBtnCancel)

                    $tsnResult = $tsnForm.ShowDialog()

                    if ($tsnResult -eq [System.Windows.Forms.DialogResult]::OK) {
                        # An account may only be assigned to one PST -- if checked in
                        # more than one CheckedListBox, the first PST in picker order wins
                        # and it is unchecked from any subsequent list before saving.
                        $assignedAccounts = [System.Collections.Generic.HashSet[string]]::new(
                            [System.StringComparer]::OrdinalIgnoreCase
                        )
                        $newMappings = [System.Collections.Generic.List[PSCustomObject]]::new()
                        foreach ($pstName in $pstChoices) {
                            $chkList = $tsnCheckLists[$pstName]
                            $checkedAccounts = @()
                            foreach ($checkedItem in $chkList.CheckedItems) {
                                if (-not $assignedAccounts.Contains($checkedItem)) {
                                    [void]$assignedAccounts.Add($checkedItem)
                                    $checkedAccounts += $checkedItem
                                }
                            }
                            if ($checkedAccounts.Count -gt 0) {
                                $newMappings.Add([PSCustomObject]@{
                                    TargetStoreName = $pstName
                                    RuleStoreNames  = @($checkedAccounts)
                                })
                            }
                        }

                        if (-not $Global:OMMigrate.WhatIf) {
                            Save-OMMigrateArchiveStoreMappings -Mappings @($newMappings)
                        }
                        Write-OMMigrateLog -Message "TargetStoreName picker: $($newMappings.Count) archive mapping(s) confirmed." -Level INFO
                        Write-Host "  Archive target mapping confirmed: $($newMappings.Count) PST(s) mapped." -ForegroundColor Green
                    }
                    else {
                        Write-OMMigrateLog -Message 'TargetStoreName picker: cancelled -- prior settings mapping (if any) left unchanged.' -Level INFO
                        Write-Host '  Archive target mapping unchanged (picker cancelled).' -ForegroundColor Yellow
                    }
                    $tsnForm.Dispose()
                }
            }
            catch {
                Write-OMMigrateLog -Message "TargetStoreName picker failed (non-fatal): $_" -Level WARN
            }

            # -- Folder tree enumeration ------------------------
            Write-OMMigrateLog -Message 'Enumerating folder tree...' -Level INFO
            Write-Host '  Enumerating folder tree (this may take a moment for large mailboxes)...' `
                       -ForegroundColor Cyan

            $folderList = Get-FolderTree -ExcludeSystemFolders $true

            foreach ($folder in $folderList) {
                [void]$Script:AllFolders.Add($folder)
            }

            Write-Host "  Folders found: $($Script:AllFolders.Count)" `
                       -ForegroundColor Green
            Write-OMMigrateLog -Message "Total folders enumerated: $($Script:AllFolders.Count)" `
                               -Level INFO

            # -- Rules inventory --------------------------------
            if ($PatchRulesCSV) {
                # *** UNDOCUMENTED -PatchRulesCSV PATH ***
                # Skips COM enumeration. Reads existing rules_inventory CSV from
                # disk, patches SendersDomain best-guess for any row missing it,
                # and loads rules into $Script:AllRules for normal CSV write flow.
                # Used when IMAP/OST blob overflow prevents COM enumeration.
                Write-OMMigrateLog -Message '-PatchRulesCSV: Skipping COM rules enumeration -- patching existing CSV.' -Level INFO
                Write-Host '  -PatchRulesCSV: Loading existing rules_inventory CSV...' -ForegroundColor Yellow

                $existingCsvPath = Get-OMMigrateCsvPath -BaseName 'rules_inventory.csv'
                if (Test-Path $existingCsvPath) {
                    # Filter out blank separator rows on load -- they sort to the top
                    # when written directly via Export-Csv. Separator rows are regenerated
                    # fresh on the next normal Script 00 run via Export-RulesToCSV.
                    $existingRows = @(Import-Csv -Path $existingCsvPath -Encoding UTF8 |
                        Where-Object { $_.RuleStoreName -and
                                       -not [string]::IsNullOrWhiteSpace($_.RuleStoreName) -and
                                       $_.RuleName -and
                                       -not [string]::IsNullOrWhiteSpace($_.RuleName) })
                    $patchCount   = 0
                    $patchedRows  = [System.Collections.Generic.List[PSCustomObject]]::new()

                    foreach ($row in $existingRows) {
                        # Build a fresh object with all required columns in correct order.
                        # This bypasses Export-RulesToCSV merge logic which would re-read
                        # the old CSV and overwrite SendersDomain with the blank value from disk.
                        $patched = [PSCustomObject]@{
                            RuleStoreName       = if ($row.PSObject.Properties['RuleStoreName'])       { $row.RuleStoreName }       else { '' }
                            TargetStoreName     = if ($row.PSObject.Properties['TargetStoreName'])     { $row.TargetStoreName }     else { '' }
                            RuleName            = if ($row.PSObject.Properties['RuleName'])            { $row.RuleName }            else { '' }
                            LastDeployedRun     = if ($row.PSObject.Properties['LastDeployedRun'])     { $row.LastDeployedRun }     else { '' }
                            # NEW (2026-07-07, Administrator direction): LastTargetRun -- separate
                            # idempotency tracking column from LastDeployedRun, positioned
                            # immediately after it. LastDeployedRun tracks rule consolidation/
                            # condition-validation work (Invoke-DeployConsolidatedRules);
                            # LastTargetRun tracks Script 03 Phase 3's folder-target (Target-
                            # FolderPath) remap work specifically. Kept as two independently-
                            # blank-gated columns so a rule freshly consolidated this run
                            # (LastDeployedRun just stamped) can still have its folder target
                            # remapped in the SAME run (LastTargetRun still blank), instead of
                            # one phase's stamp incorrectly gating the other phase's work.
                            LastTargetRun       = if ($row.PSObject.Properties['LastTargetRun'])       { $row.LastTargetRun }       else { '' }
                            TargetFolderPath    = if ($row.PSObject.Properties['TargetFolderPath'])    { $row.TargetFolderPath }    else { '' }
                            SendersDomain       = if ($row.PSObject.Properties['SendersDomain'])       { $row.SendersDomain }       else { '' }
                            NeedsFolderUpdate   = if ($row.PSObject.Properties['NeedsFolderUpdate'])   { $row.NeedsFolderUpdate }   else { 'True' }
                            IsEnabled           = if ($row.PSObject.Properties['IsEnabled'])           { $row.IsEnabled }           else { 'True' }
                            ExecutionOrder      = if ($row.PSObject.Properties['ExecutionOrder'])      { $row.ExecutionOrder }      else { '' }
                            RuleType            = if ($row.PSObject.Properties['RuleType'])            { $row.RuleType }            else { '' }
                            StopProcessing      = if ($row.PSObject.Properties['StopProcessing'])      { $row.StopProcessing }      else { '' }
                            Conditions          = if ($row.PSObject.Properties['Conditions'])          { $row.Conditions }          else { '' }
                            Actions             = if ($row.PSObject.Properties['Actions'])             { $row.Actions }             else { '' }
                            TargetFolderEntryID = if ($row.PSObject.Properties['TargetFolderEntryID']) { $row.TargetFolderEntryID } else { '' }
                            Notes               = if ($row.PSObject.Properties['Notes'])               { $row.Notes }               else { '' }
                        }

                        # Apply best-guess SendersDomain default if blank and TargetFolderPath exists
                        if ([string]::IsNullOrWhiteSpace($patched.SendersDomain) -and
                            -not [string]::IsNullOrWhiteSpace($patched.TargetFolderPath)) {
                            $pathSegments = $patched.TargetFolderPath -split '\\'
                            $patched.SendersDomain = ($pathSegments | Where-Object { $_ -ne '' } | Select-Object -Last 1).ToLower()
                            $patchCount++
                        }

                        $patchedRows.Add($patched)
                        [void]$Script:AllRules.Add($patched)
                    }

                    # Sort data rows and insert blank separator rows between account groups
                    # before writing -- matches the output format of Export-RulesToCSV.
                    $blankPatch = [PSCustomObject]@{
                        RuleStoreName = ''; TargetStoreName = ''; RuleName = ''; LastDeployedRun = ''; LastTargetRun = ''; TargetFolderPath = ''; SendersDomain = ''
                        NeedsFolderUpdate = ''; IsEnabled = ''; ExecutionOrder = ''; RuleType = ''; StopProcessing = ''
                        Conditions = ''; Actions = ''; TargetFolderEntryID = ''; Notes = ''
                    }
                    $sortedPatch   = @($patchedRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.RuleStoreName) } |
                                       Sort-Object RuleStoreName, RuleName)
                    $finalPatch    = [System.Collections.Generic.List[PSCustomObject]]::new()
                    $lastStorePatch = $null
                    foreach ($pr in $sortedPatch) {
                        if ($null -ne $lastStorePatch -and $pr.RuleStoreName -ne $lastStorePatch) {
                            $finalPatch.Add($blankPatch)
                        }
                        $finalPatch.Add($pr)
                        $lastStorePatch = $pr.RuleStoreName
                    }

                    # Write the patched CSV directly -- bypass Export-RulesToCSV merge
                    # which would re-read the old CSV and overwrite SendersDomain with blank.
                    if (-not $Global:OMMigrate.WhatIf) {
                        try {
                            $finalPatch | Export-Csv -Path $existingCsvPath -NoTypeInformation -Encoding UTF8
                            [void](Invoke-NormalizeRulesExecutionOrder -CsvPath $existingCsvPath)
                            Write-OMMigrateLog -Message "-PatchRulesCSV: CSV written directly with SendersDomain column: $existingCsvPath" `
                                               -Level INFO
                        }
                        catch {
                            Write-OMMigrateLog -Message "-PatchRulesCSV: Failed to write patched CSV: $_" -Level ERROR
                        }
                    }

                    Write-Host "  -PatchRulesCSV: Loaded $($existingRows.Count) rules, patched $patchCount SendersDomain values." `
                               -ForegroundColor Green
                    Write-OMMigrateLog -Message "-PatchRulesCSV: Loaded $($existingRows.Count) rules from CSV. Patched SendersDomain on $patchCount rows." `
                                       -Level INFO
                }
                else {
                    Write-Host "  -PatchRulesCSV: No existing rules_inventory CSV found at: $existingCsvPath" `
                               -ForegroundColor Red
                    Write-Host '  Run Script 00 against <Your Profile> first to generate the CSV.' `
                               -ForegroundColor Yellow
                    Write-OMMigrateLog -Message "-PatchRulesCSV: rules_inventory CSV not found: $existingCsvPath" -Level ERROR
                }
            }
            else {
                # -- Normal COM rules enumeration path ----------------
                Write-OMMigrateLog -Message 'Inventorying Outlook Rules...' -Level INFO
                Write-Host '  Reading Outlook Rules...' -ForegroundColor Cyan

                $rulesList = Get-OutlookRules

                # Build a set of PST store names from the folder enumeration.
                # Every rule with a folder target gets NeedsFolderUpdate=True --
                # all rules are updated after migration, no exceptions.
                foreach ($rule in $rulesList) {
                    # Always True -- never path-dependent. Code never sets either
                    # NeedsFolderUpdate or IsEnabled to False under any circumstance.
                    $rule.NeedsFolderUpdate = $true
                    $rule.IsEnabled         = $true

                    # SendersDomain -- default = last segment of TargetFolderPath.
                    # This is a BEST GUESS only. The folder name is often not the
                    # same as the real sender domain. Admin MUST review and correct
                    # this column in rules_inventory_<Profile>.csv before running Script 03.
                    # Value is preserved on re-run -- never overwritten once set.
                    #
                    # FIXED (2026-07-09, Administrator + Claude, root-caused via a 3-round
                    # diagnostic-logging investigation on the "TestProfile" profile):
                    # when TargetFolderPath is an empty string (a legitimate,
                    # documented fallback case -- see Get-OutlookRules's GetTable
                    # fallback, module line ~2982), '' -split '\' returns a
                    # single-element array containing '', which Where-Object
                    # { $_ -ne '' } then filters down to an EMPTY collection.
                    # Select-Object -Last 1 on an empty collection returns $null,
                    # and calling .ToLower() on that $null threw "You cannot call
                    # a method on a null-valued expression" -- this crashed Script
                    # 00 outright on any rule with a blank TargetFolderPath and a
                    # blank SendersDomain, which had simply never been hit before
                    # "TestProfile" surfaced a rule in this exact state (rule: "LRS
                    # Recycling (Part 1)", TargetFolderPath=''). Confirmed via a
                    # 3-round diagnostic trace: round 1 showed the rule element
                    # itself was NOT null (ruling out a null rulesList item);
                    # round 2 confirmed both preceding property assignments
                    # succeeded; round 3 confirmed IsNullOrWhiteSpace correctly
                    # evaluated True and the fill block was entered, isolating
                    # the crash to the .ToLower() call on the $null result of
                    # Select-Object -Last 1 against an empty filtered collection.
                    # Fixed by only assigning SendersDomain when at least one
                    # non-empty path segment actually exists; otherwise
                    # SendersDomain is left blank (same as it was before this
                    # fill attempt), same-spirit guard as [string]::IsNullOr-
                    # WhiteSpace checks already used elsewhere in this block.
                    if ([string]::IsNullOrWhiteSpace($rule.SendersDomain)) {
                        $pathSegments     = $rule.TargetFolderPath -split '\\'
                        $lastPathSegment  = $pathSegments | Where-Object { $_ -ne '' } | Select-Object -Last 1
                        if (-not [string]::IsNullOrWhiteSpace($lastPathSegment)) {
                            $rule.SendersDomain = $lastPathSegment.ToLower()
                        }
                    }

                    [void]$Script:AllRules.Add($rule)
                }

                $rulesWithTargets = @($Script:AllRules |
                                      Where-Object { $_.NeedsFolderUpdate }).Count

                Write-Host "  Rules found: $($Script:AllRules.Count) ($rulesWithTargets with folder targets)" `
                           -ForegroundColor Green

                Write-OMMigrateLog -Message (
                    "Rules inventoried: $($Script:AllRules.Count) total | " +
                    "$rulesWithTargets with folder targets requiring Script 03 update"
                ) -Level INFO

                # -- Rules blob backup (.bin and .rwz) ---------------
                # Export PR_RULES_DATA from the active profile's default store
                # root folder to the Backups folder. This must happen while the
                # POP3 COM session is open -- the only window where the blob is
                # healthy. For "TestProfile": run Script 00 against
                # "TestProfile" to generate the backup, then copy/rename
                # rules_inventory_<ProfileName>.csv to rules_inventory_Outlook.csv
                # before running Script 03. See QUICKSTART for details.
                Write-Host ''
                Write-Host '  Backing up rules blob (.bin / .rwz)...' -ForegroundColor Cyan
                [void](Export-RulesBlob -ProfileName $Script:SelectedProfileName)
            }
        }
        else {
            Write-OMMigrateLog -Message (
                'Could not start Outlook COM session. ' +
                'Folder and rules discovery skipped. ' +
                'Re-run Script 00 with Outlook closed to enable COM discovery.'
            ) -Level WARN
            Write-Host '  WARNING: COM session failed -- folder and rules discovery skipped.' `
                       -ForegroundColor Yellow
            $Script:FinalStatus = 'WARNING'
        }

        Write-Host ''


    # -- Update sanitization map after COM enrichment --------------
    # Re-register accounts to pick up SMTP addresses discovered via COM
    # (Exchange accounts often have different registry vs COM email addresses).
    # Also register folder names and rule names now that COM enumeration is done.
    if ($Global:OMMigrate.Sanitize) {
        Initialize-SanitizeMap `
            -Accounts @($Script:AllAccounts) `
            -Folders  @($Script:AllFolders) `
            -Rules    @($Script:AllRules)
        Write-OMMigrateLog -Message '[SANITIZE] Sanitization map updated with COM-enriched data.' `
                           -Level INFO
    }


    # ----------------------------------------------------------
    #  STEP 5 -- Generate CSV Control Files
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Generating CSV Control Files' -Step '5 of 7'

    # Compute account summary statistics
    $accountSummary = Get-AccountSummary -Accounts @($Script:AllAccounts)

    Write-Host '  Account breakdown:' -ForegroundColor Cyan
    Write-Host "    Total accounts    : $($accountSummary.Total)"      -ForegroundColor Gray
    Write-Host "    To migrate (POP3) : $($accountSummary.ToMigrate)"  -ForegroundColor Gray
    Write-Host "    Already IMAP      : $($accountSummary.FolderOnly)" -ForegroundColor Gray
    Write-Host "    Exchange (skip)   : $($accountSummary.Skipped)"    -ForegroundColor Gray
    if ($accountSummary.Manual -gt 0) {
        Write-Host "    Manual review     : $($accountSummary.Manual)" -ForegroundColor Yellow
    }
    if ($accountSummary.Unknown -gt 0) {
        Write-Host "    Unknown type      : $($accountSummary.Unknown)" -ForegroundColor Yellow
    }
    Write-Host "    Total data size   : $($accountSummary.TotalDataSizeFormatted)" `
               -ForegroundColor Gray
    Write-Host ''

    # -- migration_accounts.csv ---------------------------------
    Write-OMMigrateLog -Message 'Writing migration_accounts.csv...' -Level INFO
    # Debug: log key fields for all accounts before CSV export
    foreach ($dbgAcct in $Script:AllAccounts) {
        Write-OMMigrateLog -Message "Pre-export: $($dbgAcct.EmailAddress) | ProviderTag='$($dbgAcct.ProviderTag)' | MigrationAction='$($dbgAcct.MigrationAction)' | IncomingServer='$($dbgAcct.IncomingServer)' | Domain='$($dbgAcct.Domain)'" -Level DEBUG
    }
    $accountsCsvPath = Export-AccountsToCSV -Accounts $Script:AllAccounts

    Write-Host "  migration_accounts.csv  -> $accountsCsvPath" `
               -ForegroundColor Green

    # -- MIGRATE account picker --------------------------------
    # After the CSV is written, show a WinForms popup so the operator
    # can choose which POP3 accounts to migrate now versus defer.
    # Checked accounts get MigrationAction=MIGRATE.
    # Unchecked accounts get MigrationAction=SKIP.
    # Exchange and already-IMAP accounts are never shown -- they are
    # not migration candidates. Skipped in WhatIf/Preview mode.
    $pop3Candidates = @($Script:AllAccounts | Where-Object {
        $_.ProviderTag -like 'POP3-*'
    } | Sort-Object EmailAddress)

    if ($pop3Candidates.Count -gt 0 -and -not $Script:IsWhatIf) {
        Write-Host ''
        Write-Host '  Opening account selection window...' -ForegroundColor Cyan
        Write-Host '  Check each POP3 account you want to migrate now.' -ForegroundColor Cyan
        Write-Host '  Unchecked accounts will be set to SKIP (can be migrated later).' `
                   -ForegroundColor Cyan
        $Script:PickerConfirmed = Invoke-MigrateAccountPicker -CsvPath $accountsCsvPath
    }
    elseif ($pop3Candidates.Count -gt 0 -and $Script:IsWhatIf) {
        Write-OMMigrateLog -Message (
            "WhatIf: Would display MIGRATE account picker for $($pop3Candidates.Count) POP3 account(s)."
        ) -Level INFO -WhatIfPrefix
        Write-Host "  [PREVIEW] Account picker would open for $($pop3Candidates.Count) POP3 account(s)." `
                   -ForegroundColor DarkGray
        $Script:PickerConfirmed = $false
    }
    else {
        $Script:PickerConfirmed = $false
    }

    # -- folder_map.csv ----------------------------------------
    # Written here before the per-account loop so the folder picker
    # can filter from the already-written CSV file.
    $folderCsvPath = ''
    if ($Script:AllFolders.Count -gt 0) {
        Write-OMMigrateLog -Message 'Writing folder_map.csv...' -Level INFO
        $folderCsvPath = Export-FolderMapCSV -Folders $Script:AllFolders

        Write-Host "  folder_map.csv          -> $folderCsvPath" `
                   -ForegroundColor Green

        # Re-read the saved folder_map.csv back into $Script:AllFolders so that
        # Destination values reflect any operator selections or merge-preserved
        # values from a prior run. Export-FolderMapCSV merges existing CSV data
        # before writing, so reading it back gives us the authoritative values.
        # Without this, the report counts Server/Local from fresh default objects
        # (all Local) rather than the actual saved selections -- showing 0 for
        # Server on every rerun even when the operator has set folders to Server.
        if (Test-Path $folderCsvPath) {
            try {
                $savedFolderRows = @(Import-Csv -Path $folderCsvPath -Encoding UTF8)
                # Build a lookup keyed by StoreName + FolderPath for fast matching
                $folderDestLookup = @{}
                foreach ($savedRow in $savedFolderRows) {
                    if ($savedRow.StoreName -and $savedRow.FolderPath) {
                        $lookupKey = "$($savedRow.StoreName)|$($savedRow.FolderPath)"
                        $folderDestLookup[$lookupKey] = $savedRow.Destination
                    }
                }
                # Update in-memory folder objects with saved Destination values
                foreach ($folder in $Script:AllFolders) {
                    if ($folder.StoreName -and $folder.FolderPath) {
                        $lookupKey = "$($folder.StoreName)|$($folder.FolderPath)"
                        if ($folderDestLookup.ContainsKey($lookupKey)) {
                            # Use Add-Member -Force because Get-FolderTree objects
                            # do not have a Destination property -- it is only added
                            # by Export-FolderMapCSV to the CSV, not to in-memory objects.
                            $folder | Add-Member -NotePropertyName Destination `
                                                 -NotePropertyValue $folderDestLookup[$lookupKey] `
                                                 -Force
                        }
                    }
                }
                Write-OMMigrateLog -Message "Folder Destination values refreshed from saved folder_map.csv ($($savedFolderRows.Count) rows)." `
                                   -Level INFO
            }
            catch {
                Write-OMMigrateLog -Message "Could not refresh Destination values from folder_map.csv (non-fatal): $_" `
                                   -Level WARN
            }
        }
    }
    else {
        Write-Host '  folder_map.csv          -> Skipped (no folder data -- COM required)' `
                   -ForegroundColor Yellow
        Write-OMMigrateLog -Message 'folder_map.csv not generated -- no folder data available.' `
                           -Level WARN
    }

    # -- rules_inventory_<Profile>.csv -----------------------------------
    # Written here before the per-account loop so the rules picker
    # can filter from the already-written CSV file.
    # -PatchRulesCSV: CSV already written directly -- skip Export-RulesToCSV
    # to avoid re-merge and duplicate banner output.
    $rulesCsvPath = ''
    if ($PatchRulesCSV -and $Script:AllRules.Count -gt 0) {
        $rulesCsvPath = Get-OMMigrateCsvPath -BaseName 'rules_inventory.csv'
        Write-Host "  rules_inventory.csv     -> $rulesCsvPath" -ForegroundColor Green
        Write-Host ''
        Write-Host '  ******************************************************************' -ForegroundColor Yellow
        Write-Host '  *  ACTION REQUESTED: Review SendersDomain in rules_inventory.csv  *' -ForegroundColor Yellow
        Write-Host '  *                                                                *' -ForegroundColor Yellow
        Write-Host '  *  The SendersDomain column has been set to a BEST GUESS         *' -ForegroundColor Yellow
        Write-Host '  *  (last folder name segment). This is often INCORRECT.          *' -ForegroundColor Yellow
        Write-Host '  *  Open rules_inventory.csv and correct SendersDomain for        *' -ForegroundColor Yellow
        Write-Host '  *  every rule before running Script 03. Rules will NOT           *' -ForegroundColor Yellow
        Write-Host '  *  fire correctly until this column is verified.                 *' -ForegroundColor Yellow
        Write-Host '  ******************************************************************' -ForegroundColor Yellow
        Write-Host ''
        Write-OMMigrateLog -Message 'ADMIN ACTION REQUESTED: Review and correct SendersDomain column in rules_inventory.csv before running Script 03.' -Level INFO
    }
    elseif (-not $PatchRulesCSV -and $Script:AllRules.Count -gt 0) {
        Write-OMMigrateLog -Message 'Writing rules_inventory.csv...' -Level INFO
        # Added 2026-07-10, Administrator direction: pass the active namespace so
        # Export-RulesToCSV can resolve the real default Archive PST display
        # name for unmapped rows instead of falling back to each row's own
        # RuleStoreName -- see Export-RulesToCSV's own parameter doc for
        # full detail. Get-OutlookNamespace safely returns the same cached
        # session object established earlier in this run; $null if no COM
        # session is active, in which case Export-RulesToCSV's own existing
        # fallback behavior is unaffected.
        $rulesCsvPath = Export-RulesToCSV -Rules $Script:AllRules -Namespace (Get-OutlookNamespace)
        [void](Invoke-NormalizeRulesExecutionOrder -CsvPath $rulesCsvPath)

        Write-Host "  rules_inventory.csv     -> $rulesCsvPath" `
                   -ForegroundColor Green

        # *** ADMIN ACTION REQUESTED ***
        # The SendersDomain column has been populated with a best-guess default
        # (last segment of each rule's TargetFolderPath). This value is often
        # INCORRECT -- folder names rarely match the real sender domain.
        # Script 03 will use whatever is in this column to set the
        # 'with [domain] in the sender's address' rule condition.
        # Rules will NOT fire correctly until this column is verified.
        Write-Host '' 
        Write-Host '  ******************************************************************' -ForegroundColor Yellow
        Write-Host '  *  ACTION REQUESTED: Review SendersDomain in rules_inventory.csv  *' -ForegroundColor Yellow
        Write-Host '  *                                                                *' -ForegroundColor Yellow
        Write-Host '  *  The SendersDomain column has been set to a BEST GUESS         *' -ForegroundColor Yellow
        Write-Host '  *  (last folder name segment). This is often INCORRECT.          *' -ForegroundColor Yellow
        Write-Host '  *  Open rules_inventory.csv and correct SendersDomain for        *' -ForegroundColor Yellow
        Write-Host '  *  every rule before running Script 03. Rules will NOT           *' -ForegroundColor Yellow
        Write-Host '  *  fire correctly until this column is verified.                 *' -ForegroundColor Yellow
        Write-Host '  ******************************************************************' -ForegroundColor Yellow
        Write-Host ''
        Write-OMMigrateLog -Message 'ADMIN ACTION REQUESTED: Review and correct SendersDomain column in rules_inventory.csv before running Script 03.' -Level INFO
    }
    else {
        # No rules returned from COM this run (e.g. post-migration all-IMAP environment
        # where GetRules() returns 0 due to COM ceiling, or COM session not available).
        # Always write the rules CSV regardless -- if an existing CSV is on disk, load
        # it and pass it through Export-RulesToCSV so the merge logic runs and preserves
        # all operator edits (SendersDomain, Notes, TargetFolderPath, LastDeployedRun).
        # This ensures Script 00 always refreshes the CSV on every run, and new columns
        # (like LastDeployedRun) are injected into existing files via the merge.
        $rulesCsvPath = Get-OMMigrateCsvPath -BaseName 'rules_inventory.csv'
        if (Test-Path $rulesCsvPath) {
            Write-OMMigrateLog -Message "rules_inventory.csv: 0 rules from COM -- re-writing from existing CSV to preserve operator edits and inject any new columns." `
                               -Level INFO
            try {
                $existingRuleRows = @(Import-Csv -Path $rulesCsvPath -Encoding UTF8 |
                    Where-Object { $_.RuleName -and -not [string]::IsNullOrWhiteSpace($_.RuleName) })
                if ($existingRuleRows.Count -gt 0) {
                    # Added 2026-07-10, Administrator direction: same Namespace pass-through
                    # as the primary Export-RulesToCSV call site above.
                    $rulesCsvPath = Export-RulesToCSV -Rules ([System.Collections.Generic.List[PSCustomObject]]$existingRuleRows) -Namespace (Get-OutlookNamespace)
                    [void](Invoke-NormalizeRulesExecutionOrder -CsvPath $rulesCsvPath)
                    Write-Host "  rules_inventory.csv     -> $rulesCsvPath (preserved from prior run -- $($existingRuleRows.Count) rule(s))" `
                               -ForegroundColor Green
                    Write-OMMigrateLog -Message "rules_inventory.csv re-written from existing CSV: $($existingRuleRows.Count) rule(s) preserved." `
                                       -Level INFO
                } else {
                    Write-Host "  rules_inventory.csv     -> Skipped (existing CSV is empty)" `
                               -ForegroundColor Yellow
                    Write-OMMigrateLog -Message "rules_inventory.csv: existing CSV has no data rows -- skipped." -Level WARN
                }
            } catch {
                Write-Host "  rules_inventory.csv     -> Could not re-write existing CSV: $_" `
                           -ForegroundColor Yellow
                Write-OMMigrateLog -Message "rules_inventory.csv: re-write from existing CSV failed: $_" -Level WARN
            }
        } else {
            Write-Host "  rules_inventory.csv     -> Skipped (no rules data and no existing CSV)" `
                       -ForegroundColor Yellow
            Write-OMMigrateLog -Message "rules_inventory.csv not generated -- no rules from COM and no existing CSV on disk." `
                               -Level WARN
        }
    }

    # -- Per-account credential + folder + rules loop ----------
    # Only runs if the operator confirmed selections in the account picker.
    # If picker was cancelled or nothing was selected, skip entirely.
    if (-not $Script:IsWhatIf -and $Script:PickerConfirmed) {

        # Re-read CSV to get the post-picker MigrationAction values
        $postPickerAccounts = @()
        if (Test-Path $accountsCsvPath) {
            $postPickerAccounts = @(
                Import-Csv -Path $accountsCsvPath -Encoding UTF8 |
                Where-Object {
                    # POP3 accounts selected for migration in the picker
                    $_.MigrationAction -eq 'MIGRATE' -or
                    # IMAP-ALREADY accounts -- already on IMAP, no backup PST needed,
                    # but still need folder destination picker and rules review so
                    # their folders are created in the Archive PST correctly.
                    $_.ProviderTag -eq 'IMAP-ALREADY'
                }
            )
        }

        if ($postPickerAccounts.Count -gt 0) {
            Write-Host ''
            Write-Host ('  ' + ('-' * 54)) -ForegroundColor DarkCyan
            Write-Host "  Per-account setup: $($postPickerAccounts.Count) account(s) to configure." `
                       -ForegroundColor Cyan
            Write-Host '  For each account you will be asked for credentials,' -ForegroundColor DarkGray
            Write-Host '  folder destinations, and rules review.' -ForegroundColor DarkGray
            Write-Host ('  ' + ('-' * 54)) -ForegroundColor DarkCyan

            foreach ($migrateAcct in $postPickerAccounts) {
                $acctEmail     = $migrateAcct.EmailAddress
                # StoreName in folder_map.csv and rules_inventory_<Profile>.csv matches the
                # Outlook store display name -- for POP3/IMAP accounts this is the
                # email address, not the DisplayName field from the accounts CSV.
                $acctStoreName = $acctEmail

                $displayAcct = Invoke-OMMigrateSanitize -Text $acctEmail

                Write-Host ''
                Write-Host ('  ' + ('=' * 54)) -ForegroundColor DarkCyan
                Write-Host "  Configuring: $displayAcct" -ForegroundColor White
                Write-Host ('  ' + ('=' * 54)) -ForegroundColor DarkCyan

                # -- 1. Credential entry for this account ------
                # IMAP-ALREADY accounts are already on IMAP -- no credentials needed.
                if ($migrateAcct.ProviderTag -ne 'IMAP-ALREADY') {
                    Write-Host ''
                    Write-Host '  Step 1 of 3 -- Credentials' -ForegroundColor Cyan
                    Invoke-CredentialEntryUI -CsvPath $accountsCsvPath -FilterEmail $acctEmail
                }
                else {
                    Write-Host ''
                    Write-Host '  Step 1 of 3 -- Credentials: skipped (IMAP-ALREADY account).' `
                               -ForegroundColor DarkGray
                }

                # -- 2. Folder picker for this account ---------
                if ($folderCsvPath -and (Test-Path $folderCsvPath)) {
                    $acctFolderCount = @(
                        Import-Csv -Path $folderCsvPath -Encoding UTF8 |
                        Where-Object { $_.StoreName -eq $acctStoreName -and
                                       -not [string]::IsNullOrWhiteSpace($_.StoreName) }
                    ).Count

                    if ($acctFolderCount -gt 0) {
                        Write-Host ''
                        Write-Host '  Step 2 of 3 -- Folder destinations' -ForegroundColor Cyan
                        Invoke-FolderMapPicker -CsvPath $folderCsvPath `
                                               -FilterStoreName $acctStoreName
                    }
                    else {
                        Write-Host '  Step 2 of 3 -- Folder destinations: no folders found for this account.' `
                                   -ForegroundColor DarkGray
                    }
                }
                else {
                    Write-Host '  Step 2 of 3 -- Folder destinations: skipped (no folder data).' `
                               -ForegroundColor DarkGray
                }

                # -- 3. Rules picker for this account ----------
                if ($rulesCsvPath -and (Test-Path $rulesCsvPath)) {
                    $acctRuleCount = @(
                        Import-Csv -Path $rulesCsvPath -Encoding UTF8 |
                        Where-Object { ($_.RuleStoreName -eq $acctStoreName -or
                                        $_.TargetStoreName -eq $acctStoreName) -and
                                       -not [string]::IsNullOrWhiteSpace($_.RuleStoreName) }
                    ).Count

                    if ($acctRuleCount -gt 0) {
                        Write-Host ''
                        Write-Host '  Step 3 of 3 -- Rules review' -ForegroundColor Cyan
                        Invoke-RulesInventoryPicker -CsvPath $rulesCsvPath `
                                                    -FilterStoreName $acctStoreName
                    }
                    else {
                        Write-Host '  Step 3 of 3 -- Rules review: no rules found for this account.' `
                                   -ForegroundColor DarkGray
                    }
                }
                else {
                    Write-Host '  Step 3 of 3 -- Rules review: skipped (no rules data).' `
                               -ForegroundColor DarkGray
                }
            }

            Write-Host ''
            Write-Host ('  ' + ('-' * 54)) -ForegroundColor DarkCyan
            Write-Host '  Per-account setup complete.' -ForegroundColor Green
            Write-Host ('  ' + ('-' * 54)) -ForegroundColor DarkCyan
        }
    }
    elseif ($Script:IsWhatIf -and $pop3Candidates.Count -gt 0) {
        Write-OMMigrateLog -Message (
            "WhatIf: Would run per-account credential/folder/rules loop for MIGRATE account(s)."
        ) -Level INFO -WhatIfPrefix
        Write-Host "  [PREVIEW] Per-account setup loop would run for MIGRATE account(s)." `
                   -ForegroundColor DarkGray
    }

    Write-Host ''


    # ----------------------------------------------------------
    #  STEP 6 -- Generate Discovery HTML Report
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Generating Discovery Report' -Step '6 of 7'

    $activeProfileName = if ($ProfileName) {
        $ProfileName
    }
    elseif ($Script:DiscoveredProfiles.Count -gt 0) {
        $defaultProf = $Script:DiscoveredProfiles |
                       Where-Object { $_.IsDefault } |
                       Select-Object -First 1
        if ($defaultProf -and $defaultProf.Name) {
            $defaultProf.Name
        }
        else {
            ($Script:DiscoveredProfiles | Select-Object -First 1).Name
        }
    }
    else { 'Unknown' }

    # Convert Lists to arrays for report functions
    $accountsArray = @($Script:AllAccounts)
    $foldersArray  = @($Script:AllFolders)
    $rulesArray    = @($Script:AllRules)

    $Script:ReportFile = New-DiscoveryReport `
        -Accounts    $accountsArray `
        -Summary     $accountSummary `
        -Folders     $foldersArray `
        -Rules       $rulesArray `
        -ProfileName $activeProfileName

    Write-Host "  Discovery Report written:" -ForegroundColor Green
    Write-Host "  $Script:ReportFile" -ForegroundColor Gray
    Write-Host ''


    # ----------------------------------------------------------
    #  STEP 7 -- Write Completion Manifest
    # ----------------------------------------------------------
    # Added 2026-07-10, Administrator direction. Write-back counterpart to the
    # Switch-OMMigrateProfileSettings call earlier in this script. Live-
    # tested gap found: without this, a fresh TargetStoreName picker
    # selection (RulesEngine.ArchiveStoreMappings) would be silently
    # reverted on the NEXT run for this same profile, since nothing was
    # copying this run's changes back to the durable per-profile settings
    # file (OMMigrate_Settings_<Profile>.json). Placed here, after every
    # Settings.json-writing step this run (TargetStoreName picker included)
    # has already completed, so the synced copy reflects the final state.
    if (-not $Script:IsWhatIf) {
        Sync-OMMigrateProfileSettings -ProfileName $Script:SelectedProfileName
    }

    Show-SectionHeader -Title 'Finalizing Discovery' -Step '7 of 7'

    # Determine final status
    # Script 00 is read-only -- warnings are acceptable, only hard failures
    # (no accounts found, or explicit FAILED set earlier) cause FAILED status
    if ($Script:AllAccounts.Count -eq 0) {
        $Script:FinalStatus = 'FAILED'
    }
    elseif ($Script:FinalStatus -eq 'FAILED') {
        $Script:FinalStatus = 'FAILED'
    }
    elseif ($Global:OMMigrate.Counters.Warnings -gt 0 -or
            $Script:FinalStatus -eq 'WARNING') {
        $Script:FinalStatus = 'WARNING'
    }
    else {
        $Script:FinalStatus = 'SUCCESS'
    }

    # Write step completion manifest
    Write-StepManifest -Step 0 -Status $Script:FinalStatus -Data @{
        ProfilesScanned    = $Script:DiscoveredProfiles.Count
        AccountsDiscovered = $Script:AllAccounts.Count
        AccountsToMigrate  = $accountSummary.ToMigrate
        AccountsIMAPAlready = $accountSummary.FolderOnly
        AccountsExchange   = $accountSummary.Skipped
        FoldersEnumerated  = $Script:AllFolders.Count
        RulesInventoried   = $Script:AllRules.Count
        TotalDataSize      = $accountSummary.TotalDataSizeFormatted
        ReportFile         = $Script:ReportFile
        AccountsCsvPath    = $accountsCsvPath
    }

    Write-OMMigrateLog -Message "Step 00 manifest written. Status: $Script:FinalStatus" `
                       -Level INFO


    # ----------------------------------------------------------
    #  OPERATOR NEXT-STEPS SUMMARY
    # ----------------------------------------------------------
    Write-Host ''
    Write-Host ('-' * 60) -ForegroundColor DarkCyan
    Write-Host '  NEXT STEPS FOR OPERATOR' -ForegroundColor White
    Write-Host ('-' * 60) -ForegroundColor DarkCyan
    Write-Host ''
    Write-Host '  1. Open and review the Discovery Report:' -ForegroundColor Cyan
    Write-Host "     $Script:ReportFile" -ForegroundColor Gray
    Write-Host ''
    Write-Host '  2. Open Config\migration_accounts.csv' -ForegroundColor Cyan
    Write-Host '     Fill in the Password column for each POP3 account marked MIGRATE.' -ForegroundColor Gray
    Write-Host '     Confirm or correct IMAP server settings (NewImapServer, NewImapPort).' `
               -ForegroundColor Gray
    Write-Host '     (MigrationAction was set by the account picker above. Re-run Script 00' `
               -ForegroundColor DarkGray
    Write-Host '      at any time to change your selection before running Script 01.)' `
               -ForegroundColor DarkGray

    $awsAccounts = @($Script:AllAccounts | Where-Object { $_.IsAWSSES -eq $true })
    if ($awsAccounts.Count -gt 0) {
        Write-Host ''
        Write-Host '     AWS SES accounts also require:' -ForegroundColor Yellow
        Write-Host '       SmtpUsername  -- AWS IAM SMTP access key ID' `
                   -ForegroundColor Yellow
        Write-Host '                        (looks like: AKIAIOSFODNN7EXAMPLE)' `
                   -ForegroundColor DarkGray
        Write-Host '       SmtpPassword  -- AWS IAM SMTP secret access key' `
                   -ForegroundColor Yellow
        Write-Host '                        (NOT your AWS console password)' `
                   -ForegroundColor DarkGray
        Write-Host '       Generate these in AWS SES console under:' `
                   -ForegroundColor DarkGray
        Write-Host '       SES > SMTP Settings > Create SMTP Credentials' `
                   -ForegroundColor DarkGray

    }

    if ($accountSummary.RequiresSecureKey -gt 0) {
        Write-Host ''
        Write-Host "  3. REQUIRED: Generate Secure Mail Key for:" -ForegroundColor Yellow
        $attAccounts = $Script:AllAccounts |
            Where-Object { $_.RequiresSecureKey -eq $true }
        foreach ($att in $attAccounts) {
            Write-Host "     * $($att.EmailAddress)" -ForegroundColor Yellow
        }
        Write-Host '     Sign in at currently.com, then go directly to:' `
                   -ForegroundColor Yellow
        Write-Host '     https://www.att.com/acctmgmt/myprofile/overview?flow=settings' `
                   -ForegroundColor Cyan
        Write-Host '     Account Security -> Secure Mail Key -> Manage -> Generate' `
                   -ForegroundColor Yellow

    }

    if ($Script:AllFolders.Count -gt 0) {
        Write-Host ''
        Write-Host '  4. Folder destinations will be set automatically during Script 01.' -ForegroundColor Cyan
        Write-Host '     Re-run Script 01 -RefreshFolderMap at any time to adjust destinations.' `
                   -ForegroundColor Gray
        Write-Host '     To add a new subfolder not in the original backup, add a row to' -ForegroundColor Gray
        Write-Host '     Config\folder_map.csv manually, then re-run Script 01 -RefreshFolderMap.' `
                   -ForegroundColor Gray
    }

    Write-Host ''
    Write-Host '  When ready, run the next script:' -ForegroundColor Cyan
    Write-Host '  .\Scripts\OMMigrate-01-Backup.ps1' -ForegroundColor White
    Write-Host ''

}
catch {
    # -- Unhandled fatal error ----------------------------------
    Write-OMMigrateLog -Message "FATAL ERROR in Script 00: $_" -Level ERROR
    Write-OMMigrateLog -Message "Stack trace: $($_.ScriptStackTrace)" -Level DEBUG
    $Script:FinalStatus = 'FAILED'

    Write-Host ''
    Write-Host '  FATAL ERROR occurred during discovery:' -ForegroundColor Red
    Write-Host "  $_" -ForegroundColor Red
    Write-Host ''
    Write-Host '  Review the run log for full details:' -ForegroundColor Gray
    Write-Host "  $($Global:OMMigrate.RunLogFile)" -ForegroundColor Gray
    Write-Host ''
}
finally {
    # -- Always runs -- COM release and session close ------------
    # This block executes whether the script succeeded, failed,
    # or was interrupted. Ensures Outlook is always released.

    if ($Script:COMSessionOpen) {
        try {
            Release-OutlookCOM
        }
        catch {
            Write-OMMigrateLog -Message "COM release error (non-fatal): $_" -Level WARN
        }
        $Script:COMSessionOpen = $false
    }

    # Mark session as completing normally to suppress emergency handler
    if ($Global:OMMigrate) {
        $Global:OMMigrate.SessionCompletedNormally = $true
    }

    # Close Outlook if it was opened by this script and script completed
    # without hard failures -- leaves machine ready for Script 01
    if ($Script:COMSessionOpen -eq $false -and
        $Script:FinalStatus -ne 'FAILED' -and
        -not $Script:IsWhatIf) {
        [void](Close-OutlookIfRunning -Reason 'after successful discovery -- ready for Script 01')
    }

    # Write session completion summary
    Complete-OMMigrateSession `
        -Status     $Script:FinalStatus `
        -ReportFile $Script:ReportFile

    # Open report in browser if enabled and report was generated
    if ($Script:OpenReport -and
        $Script:ReportFile -and
        (Test-Path $Script:ReportFile) -and
        $Script:FinalStatus -ne 'FAILED' -and
        -not $Script:IsWhatIf) {
        Open-FileInEditor -FilePath $Script:ReportFile
    }

    # Open accounts CSV in editor for operator to complete
    if ($accountsCsvPath -and
        (Test-Path $accountsCsvPath) -and
        -not $Script:IsWhatIf) {
        Write-Host '  Opening migration_accounts.csv for review...' `
                   -ForegroundColor Cyan
        Open-FileInEditor -FilePath $accountsCsvPath
    }

    # Keep PowerShell window open regardless of how it was launched
    Wait-UserKeypress
}
# ***** END OF FILE *****
