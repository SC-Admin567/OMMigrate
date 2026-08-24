#Requires -Version 5.1
<#
.SYNOPSIS
    OMMigrate-01-Backup.ps1 -- PST Backup and Verification

.DESCRIPTION
    Step 01 of the OutlookMailMigrator (OMMigrate) toolkit.

    This script exports and verifies a PST backup for every POP3
    account that will be migrated, and optionally exports an OST->PST
    backup for IMAP-ALREADY accounts. It is the safety net that makes
    the entire migration reversible.

    NOTHING IS DELETED OR MODIFIED BY THIS SCRIPT.
    It is non-destructive and safe to re-run as many times as needed.

    What this script does:
        1. Reads Step 00 manifest to confirm discovery completed
        2. Reads migration_accounts.csv for the list of accounts
        3. Displays pre-flight checklist and requires confirmation
        4. For each POP3 account requiring migration:
               a. Confirms source PST file exists and is readable
               b. Prompts operator Y/N before backing up each account
               c. Exports a full PST backup via file copy
               d. Verifies backup: file exists, size > minimum, opens cleanly
               e. Logs verified backup path and size
               f. Writes per-account audit entry
        5. For each IMAP-ALREADY account:
               a. Opens Outlook COM and exports the full OST contents
                  to a PST backup via folder-by-folder item copy
               b. Verifies backup PST exists and meets minimum size
               c. Detaches backup PST and releases COM session
        6. Generates the Backup Verification HTML report
        7. Writes Step 01 manifest with list of verified backup files
           (Script 02 will not run without this manifest)

    Accounts tagged IMAP-ALREADY use an OST file that Outlook holds
    open exclusively -- they cannot be direct-copied like POP3 PSTs.
    IMAP-ALREADY accounts use Outlook COM to copy all items from the
    OST into a new PST file. This backup is required for Script 03
    rules recreation and Script 04 personal items migration.

    Accounts tagged EXCHANGE-SKIP or FOLDER-ONLY are always skipped.

    Backup files are written to:
        $BasePath\Backups\<email_address>.pst

.PARAMETER BasePath
    Override the default working directory.
    Default: $env:USERPROFILE\Documents\OutlookMigration
    Must match the BasePath used in Script 00.

.PARAMETER Preview
    Simulate the backup process -- no PST files are written,
    no Outlook changes are made. All steps are logged with [WHATIF].

.PARAMETER LogLevel
    Logging verbosity: DEBUG | INFO | WARN | ERROR
    Default: INFO

.PARAMETER SkipVerification
    Skip the post-export PST integrity verification step.
    NOT RECOMMENDED -- verification is your confirmation that the
    backup is valid before Script 02 removes anything.

.PARAMETER MinBackupSizeMB
    Minimum acceptable backup file size in MB.
    Backups smaller than this threshold fail verification.
    Default: 0 MB -- no minimum size enforced (overridden by OMMigrate_Settings.json if present)

.PARAMETER RefreshFolderMap
    Re-run mode: skips all backup and Archive pre-build steps and
    opens the folder destination picker directly. Use this after the
    initial run to reassign Server/Local destinations without
    re-running backups. folder_map.csv must already exist.

.PARAMETER Force
    Skip the Y/N prompt for each account and back up all accounts
    automatically. Use with caution -- still requires the initial
    pre-flight confirmation.

.EXAMPLE
    # Standard run -- backs up all POP3 and IMAP-ALREADY accounts
    .\OMMigrate-01-Backup.ps1

.EXAMPLE
    # Dry run -- simulate without writing any files
    .\OMMigrate-01-Backup.ps1 -Preview

.EXAMPLE
    # Auto-confirm all accounts (no per-account prompts)
    .\OMMigrate-01-Backup.ps1 -Force

.EXAMPLE
    # Custom working directory matching Script 00 run
    .\OMMigrate-01-Backup.ps1 -BasePath "D:\Migration"

.NOTES
    -------------------------------------------------------------------------
    OutlookMailMigrator (OMMigrate)
    -------------------------------------------------------------------------
    Originator & Architect:    Kirk Shallcross - Shallcross Consulting
    Implementation Specialist: Anthropic Claude AI
    Inception Date:            May 2026
    Version:                   1.5.2
    -------------------------------------------------------------------------

    Prerequisites:
        OMMigrate-00-Discover.ps1 must have completed successfully.
        A valid Step00_Complete.json manifest must exist in Manifests\.

    Safe to re-run:
        Yes. If a backup already exists and passes verification it will
        be reported as already complete. The operator is still prompted
        Y/N for each account -- typing N skips that account.

    Outlook must be closed:
        Outlook is launched in a controlled COM session by this script.
        Ensure Outlook is fully closed before running.

    Module dependencies (must be in ..\Modules\ relative to this script):
        OMMigrate-Core.psm1
        OMMigrate-Registry.psm1
        OMMigrate-Outlook.psm1
        OMMigrate-Reporting.psm1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BasePath = '',

    [Parameter(Mandatory = $false)]
    [ValidateSet('DEBUG','INFO','WARN','ERROR')]
    [string]$LogLevel = 'INFO',

    [Parameter(Mandatory = $false)]
    [switch]$SkipVerification,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$RefreshFolderMap,

    [Parameter(Mandatory = $false)]
    [switch]$Preview,

    [Parameter(Mandatory = $false)]
    [switch]$Sanitize
)

# Default MinBackupSizeMB -- passed as int param previously, now defaulted here
# to avoid SupportsShouldProcess conflict with value-type parameters
$MinBackupSizeMB = 0

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Capture Preview (WhatIf) state from explicit -Preview switch parameter
$Script:IsWhatIf = $Preview.IsPresent

# ============================================================
#  REGION: MODULE IMPORT
#  MOVED (2026-07-07, Administrator + Claude): relocated to this earliest possible
#  point in the script -- immediately after $Script:IsWhatIf is set, before
#  EVERY other block, including the -RefreshFolderMap branch. This block
#  used to run much later in the file (after the IMAP-ALREADY PRE-CREATION
#  PASS block), which meant any module-exported function called before it
#  -- Close-OutlookIfRunning in the pre-creation pass, and also
#  Initialize-OMMigrate / Invoke-FolderMapPicker in the -RefreshFolderMap
#  branch above -- did not exist yet, causing:
#    "The term 'Close-OutlookIfRunning' is not recognized..."
#  (or the equivalent for whichever function ran first) on every real
#  (non -Preview) run. Nothing in the import logic itself was changed,
#  only its position in the script.
# ============================================================

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
#  REFRESHFOLDERMAP MODE
#  When -RefreshFolderMap is specified, skip all backup and Archive
#  pre-build steps and go straight to the folder destination picker.
#  folder_map.csv must already exist from a prior run.
# ============================================================

if ($RefreshFolderMap.IsPresent) {
    Initialize-OMMigrate -BasePath  $BasePath `
                         -LogLevel  $LogLevel `
                         -IsWhatIf  $Script:IsWhatIf `
                         -Sanitize  $Sanitize.IsPresent

    Write-Host ''
    Write-Host ('=' * 60) -ForegroundColor Cyan
    Write-Host '  OMMigrate-01-Backup -- Folder Map Refresh' -ForegroundColor White
    Write-Host ('=' * 60) -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  Mode: Folder destination picker only.' -ForegroundColor Cyan
    Write-Host '  Skipping backups and Archive pre-build.' -ForegroundColor DarkGray
    Write-Host ''

    $rfmCsvPath = Get-OMMigrateCsvPath -BaseName 'folder_map.csv'
    if (-not (Test-Path $rfmCsvPath)) {
        Write-Host '  ERROR: folder_map.csv not found -- run Script 01 normally first.' `
                   -ForegroundColor Red
        Write-OMMigrateLog -Message 'RefreshFolderMap: folder_map.csv not found -- aborting.' `
                           -Level ERROR
        exit 1
    }

    Write-Host '  Opening folder destination picker...' -ForegroundColor Cyan
    Invoke-FolderMapPicker -CsvPath $rfmCsvPath
    Write-Host ''
    Write-Host '  Folder map updated. Run Script 02 to continue migration.' `
               -ForegroundColor Green
    Write-OMMigrateLog -Message 'RefreshFolderMap: Folder picker complete.' -Level INFO
    exit 0
}

# ============================================================
#  REGION: SESSION INITIALIZATION
# ============================================================
#
# FIXED 2026-07-10, Administrator direction. Live-tested crash found: this call was
# previously positioned AFTER the IMAP-ALREADY Pre-Creation Pass below,
# because that block's own comment said it "must run before the main try{}
# block so no COM session or Outlook state exists yet." That requirement is
# about avoiding a COM session -- Initialize-OMMigrate never opens one (it
# only sets up paths/logging/settings), so moving it earlier does not
# violate the original intent, and nothing between the two blocks depended
# on the old order. Moved here specifically because the Pre-Creation Pass
# now calls Get-OMMigrateCsvPath (profile-aware CSV path fix, same session),
# which requires $Global:OMMigrate.ConfigPath and .Settings -- both only
# set by Initialize-OMMigrate. Confirmed live: running with the old order
# crashed with "Cannot bind argument to parameter 'Path' because it is
# null" the moment Get-OMMigrateCsvPath tried to use a ConfigPath that
# didn't exist yet.
Initialize-OMMigrate `
    -ScriptName 'OMMigrate-01-Backup' `
    -BasePath   $BasePath `
    -LogLevel   $LogLevel `
    -IsWhatIf   $Script:IsWhatIf `
    -Sanitize   $Sanitize.IsPresent

# Register exit handlers -- Script 01 is non-destructive so
# exit at any point is safe, but COM must still be released
Register-ExitHandlers -ScriptStep 1


# ============================================================
#  IMAP-ALREADY PRE-CREATION PASS
#  Runs automatically for all IMAP-ALREADY accounts before any
#  other work, while Outlook is guaranteed clean. Pre-creates
#  backup PST files from template and stores account list in
#  $Script:IMAPAccountsToBackup for use in Step 5b.
#  Must run before the main try{} block so no COM session or
#  Outlook state exists yet -- see Session Initialization region
#  above for why Initialize-OMMigrate itself was moved ahead of
#  this block (it does not open COM, so this ordering requirement
#  is unaffected).
# ============================================================

$Script:IMAPAccountsToBackup = [System.Collections.Generic.List[PSCustomObject]]::new()

if (-not $Script:IsWhatIf) {

    # -- Kill any running Outlook before attempting COM -----------
    [void](Close-OutlookIfRunning -Reason 'before IMAP backup pre-creation')

    # -- Load IMAP-ALREADY accounts from CSV ----------------------
    # FIXED 2026-07-10, Administrator direction. Live-tested bug found: this path was
    # hardcoded to the plain 'migration_accounts.csv' filename via a raw
    # $env:USERPROFILE join, bypassing BOTH Get-OMMigrateCsvPath's profile
    # suffix (e.g. '_TestProfile') AND $Global:OMMigrate.ConfigPath (which
    # itself may differ from the hardcoded default if -BasePath was ever
    # overridden at Initialize-OMMigrate). Every other CSV path reference in
    # this same file already correctly uses Get-OMMigrateCsvPath -- this one
    # was missed, and confirmed live to have been silently reading a CSV
    # that doesn't exist (the real data lives in the profile-suffixed file),
    # so $imapAlreadyAccounts was always empty and OST backup for
    # IMAP-ALREADY accounts (Step 5b) never actually ran, on any profile,
    # since this code was written -- not specific to today's feature work.
    $preCsvPath = Get-OMMigrateCsvPath -BaseName 'migration_accounts.csv'
    $imapAlreadyAccounts = @()
    if (Test-Path $preCsvPath) {
        $imapAlreadyAccounts = @(
            Import-Csv -Path $preCsvPath -Encoding UTF8 |
            Where-Object { $_.ProviderTag -eq 'IMAP-ALREADY' } |
            Sort-Object EmailAddress
        )
    }

    if ($imapAlreadyAccounts.Count -eq 0) {
        Write-Host '  No IMAP-ALREADY accounts found -- OST backup skipped.' `
                   -ForegroundColor DarkGray
    }
    else {
        # All IMAP-ALREADY accounts are always backed up -- no picker
        foreach ($a in $imapAlreadyAccounts) { [void]$Script:IMAPAccountsToBackup.Add($a) }
        Write-Host "  Found $($Script:IMAPAccountsToBackup.Count) IMAP-ALREADY account(s) for OST backup." `
                   -ForegroundColor Cyan
    }

    # -- Pre-create PST files via throwaway COM -------------------
    # Outlook is closed. Copy from OMMigrate_Template.pst -- no COM required.
    if ($Script:IMAPAccountsToBackup.Count -gt 0) {
        Write-Host '  Pre-creating backup PST files from template...' `
                   -ForegroundColor Cyan

        # FIXED (launch-readiness review, 2026-08-18): was hardcoded to the
        # default Documents\OutlookMigration path, bypassing -BasePath entirely --
        # a custom -BasePath run would split IMAP-ALREADY OST backups from the
        # rest of that same run's data. Now uses $Global:OMMigrate.BackupPath,
        # same as every other backup-path reference in this script.
        $preBackupDir    = $Global:OMMigrate.BackupPath
        $archiveTemplate = Join-Path $preBackupDir 'OMMigrate_Template.pst'

        # ADDED (Administrator direction, 2026-08-18 -- Fix 4 live-test stability):
        # confirmed live that a custom -BasePath pointed at a brand-new
        # location has no OMMigrate_Template.pst yet -- Install.ps1 only ever
        # seeds the DEFAULT Documents\OutlookMigration\Backups location, so
        # a fresh -BasePath folder genuinely has nothing to copy from,
        # even though the fix above correctly resolves $preBackupDir to the
        # new location. Before failing, check whether the template already
        # exists at the ORIGINAL default location and copy it over --
        # mirrors the same Copy-Item-from-template pattern this script
        # already uses to seed each account's own backup PST just below,
        # just sourced from the default location instead of $preBackupDir
        # when $preBackupDir doesn't have one yet. Only falls through to the
        # original hard error if neither location has a template to copy.
        if (-not (Test-Path $archiveTemplate)) {
            $defaultTemplateSource = Join-Path (Join-Path $env:USERPROFILE 'Documents\OutlookMigration\Backups') 'OMMigrate_Template.pst'
            if (Test-Path $defaultTemplateSource) {
                try {
                    Copy-Item -Path $defaultTemplateSource -Destination $archiveTemplate -Force -ErrorAction Stop
                    Write-Host "  OMMigrate_Template.pst not found at -BasePath -- copied from default location ($defaultTemplateSource)." `
                               -ForegroundColor Yellow
                    Write-OMMigrateLog -Message "OST pre-create: OMMigrate_Template.pst not found at '$preBackupDir' -- copied from default location '$defaultTemplateSource'." `
                                       -Level INFO
                }
                catch {
                    Write-Host "  ERROR: OMMigrate_Template.pst found at default location but could not be copied to -BasePath: $_" `
                               -ForegroundColor Red
                    $Script:IMAPAccountsToBackup.Clear()
                }
            }
        }

        if (-not (Test-Path $archiveTemplate)) {
            Write-Host '  ERROR: OMMigrate_Template.pst not found in Backups folder (checked -BasePath and default location).' `
                       -ForegroundColor Red
            Write-Host '  Re-run Install.ps1 to generate it, then retry.' `
                       -ForegroundColor Yellow
            $Script:IMAPAccountsToBackup.Clear()
        }
        else {
            $created = 0
            foreach ($acct in $Script:IMAPAccountsToBackup) {
                $safeEm  = $acct.EmailAddress -replace '[\\/:*?"<>|@]', '_'
                # FIXED 2026-07-10, Administrator direction. Live-tested bug found:
                # this backup PST filename was NOT profile-suffixed, so the
                # same account backed up under two different Outlook
                # profiles (e.g. "TestProfile" and Administrator's real Outlook profile)
                # silently overwrote one profile's backup with the other's
                # on every run -- confirmed live. Now uses
                # Get-OMMigrateCsvPath's -BasePathOverride (same profile-
                # suffix logic already proven for the three CSVs) so each
                # profile gets its own uniquely-named backup file.
                $pstPath = Get-OMMigrateCsvPath -BaseName "${safeEm}_osttoimap.pst" -BasePathOverride $preBackupDir

                if (Test-Path $pstPath) {
                    try { Remove-Item -Path $pstPath -Force -ErrorAction Stop } catch { }
                }

                try {
                    Copy-Item -Path $archiveTemplate -Destination $pstPath -Force -ErrorAction Stop
                    $created++
                    Write-Host "    Pre-created: $([System.IO.Path]::GetFileName($pstPath))" `
                               -ForegroundColor DarkGray
                }
                catch {
                    Write-Host "    WARNING: Could not pre-create PST for $($acct.EmailAddress): $_" `
                               -ForegroundColor Yellow
                }
            }

            if ($created -eq 0) {
                Write-Host '  ERROR: No PST files were created. OST backup will be skipped.' `
                           -ForegroundColor Red
                $Script:IMAPAccountsToBackup.Clear()
            }
            else {
                Write-Host "  PST pre-creation complete ($created file(s))." -ForegroundColor Green
                Write-Host ''
            }
        }
    }
}

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

# Show the safe-exit banner before any work begins
Show-ExitBanner

# Script-level state variables
$Script:COMSessionOpen       = $false
$Script:FinalStatus          = 'SUCCESS'
$Script:ReportFile           = ''
$Script:AccountResults       = [System.Collections.Generic.List[PSCustomObject]]::new()
$Script:VerifiedBackups      = [System.Collections.Generic.List[string]]::new()
$Script:TotalBackupBytes     = 0L
$Script:IMAPBackupSucceeded  = 0
$Script:IMAPBackupFailed     = 0
$Script:IMAPBackupSkipped    = 0


# ============================================================
#  HELPER: Export-PSTBackup
#  Exports a single account's PST data via Outlook COM.
#  Returns a result object with BackupPath, size, and status.
# ============================================================

function Export-PSTBackup {
    <#
    .SYNOPSIS
        Copies a POP3 account's PST file to the backup folder.

    .DESCRIPTION
        Performs a direct file copy of the source PST file to the
        backup destination. Outlook must be fully closed before this
        runs -- when Outlook is closed the PST is fully flushed to
        disk and not locked, making a direct file copy safe, fast,
        and reliable.

        No COM session is required or used by this function.

        Verifies the output file exists and meets the minimum size
        requirement before reporting success.

    .PARAMETER Account
        The account object to back up. Must have a valid PSTPath.

    .PARAMETER BackupFolder
        Destination folder for the backup PST file.

    .PARAMETER MinSizeBytes
        Minimum acceptable file size in bytes. Backups smaller
        than this threshold are marked as FAILED.

    .OUTPUTS
        PSCustomObject with BackupPath, BackupSizeBytes,
        BackupSizeFormatted, BackupStatus, BackupDetail.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Account,

        [Parameter(Mandatory = $true)]
        [string]$BackupFolder,

        [Parameter(Mandatory = $true)]
        [long]$MinSizeBytes
    )

    $email           = $Account.EmailAddress
    $safeEmail       = Get-SafeFileName -InputString $email
    # REVERTED 2026-07-11, Administrator direction. The 2026-07-10 profile-suffix fix
    # was WRONG for this specific file. Unlike the CSVs and the _osttoimap
    # OST-export backups (which genuinely differ per profile and must stay
    # suffixed), the plain POP3 backup PST is a ONE-TIME artifact created
    # the first time an account is converted from POP3 -- it is the single
    # authoritative source of that account's historical mail/folders, read
    # by Script 01's Archive pre-build, Script 03, and Script 04, regardless
    # of which Outlook profile (Outlook, "TestProfile", etc.) is currently running
    # the pipeline. Suffixing it broke that: "TestProfile" (seeded from copies of
    # the Outlook profile's data files, not a fresh conversion) was looking
    # for a file that only the original Outlook-profile conversion could
    # ever create, and "TestProfile" never runs a POP3 conversion for this
    # account. Confirmed live -- see 2026-07-11 session notes. This path is
    # intentionally NOT run through Get-OMMigrateCsvPath's profile-suffix
    # logic; it stays a bare, profile-independent filename everywhere it is
    # referenced.
    $backupFile      = Join-Path $BackupFolder "$safeEmail.pst"
    # Sanitized display path for log output -- real path used for operations
    $safeEmailMasked = Get-SafeFileName -InputString (Invoke-OMMigrateSanitize -Text $email)
    $backupFileMasked = Join-Path $BackupFolder "$safeEmailMasked.pst"

    $result = [PSCustomObject]@{
        BackupPath          = $backupFile
        BackupSizeBytes     = 0L
        BackupSizeFormatted = ''
        BackupStatus        = 'PENDING'
        BackupDetail        = ''
    }

    # -- WhatIf mode -- simulate only --------------------------
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message "WhatIf: Would copy $(Invoke-OMMigrateSanitize -Text $Account.PSTPath) -> $backupFileMasked" `
                           -Level INFO -WhatIfPrefix
        $result.BackupStatus        = 'SUCCESS'
        $result.BackupDetail        = 'WhatIf -- not written'
        $result.BackupSizeFormatted = 'N/A'
        return $result
    }

    # -- Verify source PST exists -------------------------------
    if (-not $Account.PSTPath -or -not (Test-Path $Account.PSTPath)) {
        $msg = if (-not $Account.PSTPath) {
            "No PST path known for $email -- cannot back up."
        } else {
            "PST file not found at: $($Account.PSTPath)"
        }
        Write-OMMigrateLog -Message $msg -Level ERROR
        $result.BackupStatus = 'FAILED'
        $result.BackupDetail = $msg
        return $result
    }

    # -- File copy backup ---------------------------------------
    # Outlook is closed -- PST is fully flushed and unlocked.
    # Direct file copy is fast, reliable, and byte-for-byte accurate.
    Write-OMMigrateLog -Message "Copying PST: $(Invoke-OMMigrateSanitize -Text $Account.PSTPath) -> $backupFileMasked" `
                       -Level INFO

    try {
        Copy-Item -Path        $Account.PSTPath `
                  -Destination $backupFile `
                  -Force `
                  -ErrorAction Stop

        # Update timestamp to reflect when backup was taken
        # Copy-Item preserves the source file's timestamp by default
        (Get-Item $backupFile).LastWriteTime = Get-Date

        Write-OMMigrateLog -Message "PST copy complete: $backupFileMasked" -Level INFO
    }
    catch {
        Write-OMMigrateLog -Message "PST copy failed for $email : $_" -Level ERROR
        $result.BackupStatus = 'FAILED'
        $result.BackupDetail = "Copy error: $_"
        return $result
    }

    # -- Verify the backup file ---------------------------------
    if (-not (Test-Path $backupFile)) {
        $result.BackupStatus = 'FAILED'
        $result.BackupDetail = 'Backup file not found after copy'
        return $result
    }

    $fileInfo = Get-Item $backupFile
    $result.BackupSizeBytes     = $fileInfo.Length
    $result.BackupSizeFormatted = Format-FileSize -Bytes $fileInfo.Length

    # Size check
    if ($fileInfo.Length -lt $MinSizeBytes) {
        $result.BackupStatus = 'FAILED'
        $result.BackupDetail = (
            "Backup too small: $($result.BackupSizeFormatted) " +
            "(minimum: $(Format-FileSize -Bytes $MinSizeBytes)). " +
            "File may be empty or corrupt."
        )
        Write-OMMigrateLog -Message "VERIFICATION FAILED -- $email backup too small: $($result.BackupSizeFormatted)" `
                           -Level ERROR
        return $result
    }

    $result.BackupStatus = 'SUCCESS'
    $result.BackupDetail = "Verified: $($result.BackupSizeFormatted)"
    Write-OMMigrateLog -Message "Backup verified: $backupFileMasked ($($result.BackupSizeFormatted))" `
                       -Level INFO

    return $result
}


# ============================================================
#  HELPER: Export-OSTBackup
#  Exports an IMAP-ALREADY account's OST to a PST backup via
#  Outlook COM folder-by-folder item copy.
#  Returns a result object with BackupPath, size, and status.
# ============================================================

function Export-OSTBackup {
    <#
    .SYNOPSIS
        Exports an IMAP-ALREADY account's full OST contents to a
        PST backup file via Outlook COM item copy.

    .DESCRIPTION
        OST files are held exclusively by Outlook and cannot be
        direct-copied like POP3 PSTs. This function uses an active
        Outlook COM session to:
            1. Locate the IMAP store for the account by email address
            2. Mount a new empty PST at the backup destination path
            3. Walk every folder in the IMAP store
            4. Copy all items from each folder into the backup PST
            5. Detach the backup PST from the Outlook profile

        System folders (Calendar, Contacts, Tasks, Notes, Journal,
        Junk, Sync Issues) are included -- this is a full backup
        needed for Script 03 rules recreation and Script 04 personal
        items migration.

        UPDATED 2026-07-11, Administrator direction: Calendar, Contacts, Tasks,
        Notes, and Journal folders are now SKIPPED by this backup step
        (see Copy-OSTFolderRecursive's DefaultItemType-based skip logic
        below). Root cause: Item.Copy().Move() on Calendar/Contacts
        items over Outlook COM was found to leave duplicate items
        behind in the LIVE source folder on every run (confirmed via
        Outlook COM's documented recurring-appointment copy/move
        restrictions), silently corrupting production data. Script 04
        already owns Calendar/Contacts/Tasks/Notes/Journal migration
        directly from the live account and does not depend on this
        backup PST containing them, so removing them from this backup
        step is a pure risk-reduction with no functional loss. This
        backup step is now MAIL-ONLY (Inbox, Sent, Drafts, Deleted
        Items, Junk, and custom mail subfolders).

        Requires an active Outlook COM session. The calling script is
        responsible for Connect-OutlookCOM and Release-OutlookCOM.

    .PARAMETER Account
        The IMAP-ALREADY account object from migration_accounts.csv.

    .PARAMETER BackupFolder
        Destination folder path for the backup PST file.

    .PARAMETER Namespace
        Active Outlook MAPI namespace COM object.

    .PARAMETER MinSizeBytes
        Minimum acceptable backup file size in bytes.

    .OUTPUTS
        PSCustomObject with BackupPath, BackupSizeBytes,
        BackupSizeFormatted, BackupStatus, BackupDetail,
        FoldersCopied, ItemsCopied.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Account,

        [Parameter(Mandatory = $true)]
        [string]$BackupFolder,

        [Parameter(Mandatory = $true)]
        [object]$Namespace,

        [Parameter(Mandatory = $true)]
        [long]$MinSizeBytes
    )

    $email     = $Account.EmailAddress
    $safeEmail = Get-SafeFileName -InputString $email
    # FIXED 2026-07-10, Administrator direction. Same profile-suffix fix as the
    # pre-creation pass above -- see that call site's comment for full
    # context. This function's own filename construction must match the
    # pre-creation pass exactly, or it will look for a file that was never
    # created (a different, unsuffixed name).
    $backupFile = Get-OMMigrateCsvPath -BaseName "${safeEmail}_osttoimap.pst" -BasePathOverride $BackupFolder

    $result = [PSCustomObject]@{
        BackupPath          = $backupFile
        BackupSizeBytes     = 0L
        BackupSizeFormatted = ''
        BackupStatus        = 'PENDING'
        BackupDetail        = ''
        FoldersCopied       = 0
        ItemsCopied         = 0
        OSTSourcePath       = ''   # Actual OST filepath -- populated after store found
    }

    # -- WhatIf mode -------------------------------------------
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message "WhatIf: Would export OST -> PST for $(Invoke-OMMigrateSanitize -Text $email)" `
                           -Level INFO -WhatIfPrefix
        $result.BackupStatus        = 'SUCCESS'
        $result.BackupDetail        = 'WhatIf -- not written'
        $result.BackupSizeFormatted = 'N/A'
        return $result
    }

    # -- Locate IMAP store for this account --------------------
    $imapStore = $null
    try {
        $stores = $Namespace.Stores
        Register-COMObject -ComObject $stores
        for ($s = 1; $s -le $stores.Count; $s++) {
            $st = $stores.Item($s)
            Register-COMObject -ComObject $st
            $fp = ''
            try { $fp = $st.FilePath } catch { }
            # Match by email in display name, exclude PSTs and the
            # backup PST we are about to create
            if ($st.DisplayName -like "*$email*" -and
                -not ($fp -like '*.pst')) {
                $imapStore = $st
                break
            }
        }
    }
    catch {
        $result.BackupStatus = 'FAILED'
        $result.BackupDetail = "Error scanning stores: $_"
        Write-OMMigrateLog -Message "Export-OSTBackup: Store scan failed for $email : $_" `
                           -Level ERROR
        return $result
    }

    if (-not $imapStore) {
        $result.BackupStatus = 'FAILED'
        $result.BackupDetail = "IMAP store not found in active Outlook session for: $email"
        Write-OMMigrateLog -Message "Export-OSTBackup: IMAP store not found for $email" `
                           -Level ERROR
        return $result
    }

    Write-OMMigrateLog -Message "Export-OSTBackup: Found IMAP store '$($imapStore.DisplayName)' for $email" `
                       -Level INFO

    # Capture the actual OST filepath for the backup report source column
    try { $result.OSTSourcePath = $imapStore.FilePath } catch { }

    # The backup PST file must already exist -- pre-created by the
    # Step 5b pre-creation pass before the main COM session opened.
    # AddStore on an already-existing file never hangs -- MAPI just
    # opens it. Cannot create new PSTs from within a COM session that
    # is already managing live IMAP stores (MAPI serializes and blocks).
    if (-not (Test-Path $backupFile)) {
        $result.BackupStatus = 'FAILED'
        $result.BackupDetail = "Backup PST was not pre-created. Re-run to retry."
        Write-OMMigrateLog -Message "Export-OSTBackup: $($result.BackupDetail)" -Level ERROR
        return $result
    }

    # -- Mount the pre-created PST in the main COM session -----
    $backupStore = $null
    try {
        $Namespace.AddStore($backupFile)

        # Poll up to 15 seconds for the store to register
        $storeFound  = $false
        $pollElapsed = 0
        while ($pollElapsed -lt 30) {
            Start-Sleep -Milliseconds 500
            $pollElapsed++
            $storesRefresh = $Namespace.Stores
            Register-COMObject -ComObject $storesRefresh
            for ($s = 1; $s -le $storesRefresh.Count; $s++) {
                $st = $storesRefresh.Item($s)
                Register-COMObject -ComObject $st
                $fp = ''
                try { $fp = $st.FilePath } catch { }
                if ($fp -eq $backupFile) {
                    $backupStore = $st
                    $storeFound  = $true
                    break
                }
            }
            if ($storeFound) { break }
        }

        if (-not $storeFound) {
            $result.BackupStatus = 'FAILED'
            $result.BackupDetail = "Backup PST file exists but could not be mounted in Outlook after 15s."
            Write-OMMigrateLog -Message "Export-OSTBackup: $($result.BackupDetail)" -Level ERROR
            return $result
        }
    }
    catch {
        $result.BackupStatus = 'FAILED'
        $result.BackupDetail = "Could not mount backup PST in main COM session: $_"
        Write-OMMigrateLog -Message "Export-OSTBackup: $($result.BackupDetail)" -Level ERROR
        return $result
    }

    # Rename backup PST root folder to identify it clearly
    try {
        $backupRoot = $backupStore.GetRootFolder()
        Register-COMObject -ComObject $backupRoot
        $backupRoot.Name = "Backup -- $email"
    }
    catch { }

    # -- Walk IMAP store and copy all folders ------------------
    $foldersCopied = 0
    $itemsCopied   = 0

    try {
        $imapRoot = $imapStore.GetRootFolder()
        Register-COMObject -ComObject $imapRoot
        $backupRoot = $backupStore.GetRootFolder()
        Register-COMObject -ComObject $backupRoot

        # Recursive folder copy -- mirrors the IMAP folder structure
        # into the backup PST preserving all hierarchy
        $copyStats = Copy-OSTFolderRecursive `
            -SourceFolder $imapRoot `
            -DestFolder   $backupRoot `
            -FolderPath   ''

        $foldersCopied = $copyStats.Folders
        $itemsCopied   = $copyStats.Items
    }
    catch {
        Write-OMMigrateLog -Message "Export-OSTBackup: Error during folder copy for $email : $_" `
                           -Level WARN
        # Non-fatal -- partial backup is still valuable
    }

    # -- Detach backup PST from Outlook profile ----------------
    # The backup PST must not remain mounted in the live profile.
    try {
        $detachRoot = $backupStore.GetRootFolder()
        Register-COMObject -ComObject $detachRoot
        $Namespace.RemoveStore($detachRoot)
        Start-Sleep -Milliseconds 500
        Write-OMMigrateLog -Message "Export-OSTBackup: Backup PST detached from profile: $backupFile" `
                           -Level INFO
    }
    catch {
        Write-OMMigrateLog -Message "Export-OSTBackup: Could not detach backup PST '$backupFile': $_" `
                           -Level WARN
        # Non-fatal -- PST is written, operator can close manually
    }

    # -- Verify backup file ------------------------------------
    if (-not (Test-Path $backupFile)) {
        $result.BackupStatus = 'FAILED'
        $result.BackupDetail = 'Backup PST not found on disk after export.'
        Write-OMMigrateLog -Message "Export-OSTBackup: $($result.BackupDetail)" -Level ERROR
        return $result
    }

    $fileInfo = Get-Item $backupFile
    $result.BackupSizeBytes     = $fileInfo.Length
    $result.BackupSizeFormatted = Format-FileSize -Bytes $fileInfo.Length
    $result.FoldersCopied       = $foldersCopied
    $result.ItemsCopied         = $itemsCopied

    if ($fileInfo.Length -lt $MinSizeBytes -and $MinSizeBytes -gt 0) {
        $result.BackupStatus = 'FAILED'
        $result.BackupDetail = (
            "Backup too small: $($result.BackupSizeFormatted) " +
            "(minimum: $(Format-FileSize -Bytes $MinSizeBytes)). " +
            "File may be empty or corrupt."
        )
        Write-OMMigrateLog -Message "Export-OSTBackup: VERIFICATION FAILED -- $email backup too small: $($result.BackupSizeFormatted)" `
                           -Level ERROR
        return $result
    }

    $result.BackupStatus = 'SUCCESS'
    $result.BackupDetail = "Verified: $($result.BackupSizeFormatted) | Folders=$foldersCopied | Items=$itemsCopied"
    Write-OMMigrateLog -Message "Export-OSTBackup: Backup verified: $backupFile ($($result.BackupSizeFormatted) | Folders=$foldersCopied | Items=$itemsCopied)" `
                       -Level INFO

    return $result
}


# ============================================================
#  HELPER: Copy-OSTFolderRecursive
#  Recursively copies all folders and items from a source
#  MAPIFolder into a destination MAPIFolder.
# ============================================================

function Copy-OSTFolderRecursive {
    <#
    .SYNOPSIS
        Recursively copies all subfolders and items from SourceFolder
        into DestFolder, preserving the full folder hierarchy.

    .DESCRIPTION
        Used by Export-OSTBackup to mirror an IMAP OST store into a
        backup PST. Skips the store root itself (it has no items of
        its own) and walks all subfolders recursively.

        Items are copied using Item.Copy().Move() -- same proven
        pattern used by Script 03's Copy-FolderContents.

        System infrastructure folders that contain no user data are
        skipped: Sync Issues, Conflicts, Local Failures, Server
        Failures, Recoverable Items. All personal data folders
        (Calendar, Contacts, Tasks, Notes, Journal, Inbox, Sent,
        and all custom folders) are included.

        UPDATED 2026-07-11, Administrator direction: Calendar, Contacts, Tasks,
        Notes, and Journal folders are now SKIPPED entirely (matched
        by DefaultItemType, same name-independent method Script 04
        uses -- see the $skipArtifactItemTypes check below). This
        function is now MAIL-ONLY. Root cause of the change: copying
        these item types via Item.Copy().Move() over Outlook COM was
        found to duplicate items in the LIVE source folder on repeated
        runs (Outlook COM restricts moving individual occurrences of
        recurring appointment series -- see Microsoft Learn KB on
        'Cannot move the items' errors for recurring appointments).
        Script 04 already migrates these artifact types directly from
        the live account and does not read from this backup PST, so
        no functional coverage is lost by skipping them here.

    .PARAMETER SourceFolder
        Source MAPIFolder COM object to copy from.

    .PARAMETER DestFolder
        Destination MAPIFolder COM object to copy into.

    .PARAMETER FolderPath
        Display path for logging (built up during recursion).

    .OUTPUTS
        PSCustomObject with Folders [int] and Items [int] counts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$SourceFolder,

        [Parameter(Mandatory = $true)]
        [object]$DestFolder,

        [Parameter(Mandatory = $false)]
        [string]$FolderPath = ''
    )

    $totalFolders = 0
    $totalItems   = 0

    # System infrastructure folders -- no user data, skip entirely
    $skipFolders = @(
        'Sync Issues', 'Conflicts', 'Local Failures', 'Server Failures',
        'Recoverable Items', 'Deletions', 'Purges', 'Versions',
        'Quick Step Settings', 'Conversation Action Settings'
    )

    # ADDED 2026-07-11, Administrator direction: Personal artifact folders
    # (Calendar, Contacts, Tasks, Notes, Journal) are owned exclusively
    # by Script 04 and must be skipped by this mail-only backup step.
    # Matched by DefaultItemType rather than folder name -- same
    # name-independent method Script 04 uses in Get-ArtifactFolder,
    # so a renamed folder (e.g. a Notes folder called 'Passwords')
    # is still correctly identified and skipped.
    # OlItemType constants: olAppointmentItem=1, olContactItem=2,
    # olTaskItem=3, olJournalItem=4, olNoteItem=5.
    $skipArtifactItemTypes = @(1, 2, 3, 4, 5)

    try {
        $subfolders = $SourceFolder.Folders
        Register-COMObject -ComObject $subfolders

        for ($i = 1; $i -le $subfolders.Count; $i++) {
            $srcSub = $null
            try {
                $srcSub     = $subfolders.Item($i)
                Register-COMObject -ComObject $srcSub
                $folderName = $srcSub.Name

                # Skip infrastructure folders
                if ($skipFolders -contains $folderName) {
                    Write-OMMigrateLog -Message "Copy-OSTFolderRecursive: Skipping infrastructure folder '$folderName'" `
                                       -Level DEBUG
                    continue
                }

                # ADDED 2026-07-11, Administrator direction: Skip personal artifact
                # folders (Calendar, Contacts, Tasks, Notes, Journal) --
                # owned exclusively by Script 04. Matched by DefaultItemType
                # (name-independent), not by folder display name. Root cause
                # this guards against: Item.Copy().Move() on these item
                # types was found to duplicate items in the LIVE source
                # folder on repeated runs -- see function docstring above.
                $srcSubItemType = $null
                try { $srcSubItemType = $srcSub.DefaultItemType } catch { }
                if ($null -ne $srcSubItemType -and $skipArtifactItemTypes -contains $srcSubItemType) {
                    Write-OMMigrateLog -Message "Copy-OSTFolderRecursive: Skipping artifact folder '$folderName' (DefaultItemType=$srcSubItemType -- owned by Script 04)" `
                                       -Level DEBUG
                    continue
                }

                $childPath = if ($FolderPath) { "$FolderPath\$folderName" } else { $folderName }

                # Find or create matching folder in destination
                $destSub = $null
                try {
                    $destSubs = $DestFolder.Folders
                    for ($d = 1; $d -le $destSubs.Count; $d++) {
                        $ds = $destSubs.Item($d)
                        Register-COMObject -ComObject $ds
                        if ($ds.Name -eq $folderName) {
                            $destSub = $ds
                            break
                        }
                    }
                }
                catch { }

                if (-not $destSub) {
                    try {
                        $destSub = $DestFolder.Folders.Add($folderName)
                        Register-COMObject -ComObject $destSub
                        $totalFolders++
                        Write-OMMigrateLog -Message "Copy-OSTFolderRecursive: Created folder '$childPath'" `
                                           -Level DEBUG
                    }
                    catch {
                        Write-OMMigrateLog -Message "Copy-OSTFolderRecursive: Could not create '$childPath': $_" `
                                           -Level WARN
                        continue
                    }
                }

                # Copy items from source to destination folder
                $copied = 0
                $failed = 0
                try {
                    $items = $srcSub.Items
                    $count = $items.Count
                    if ($count -gt 0) {
                        Write-OMMigrateLog -Message "Copy-OSTFolderRecursive: Copying $count item(s) from '$childPath'" `
                                           -Level DEBUG
                        for ($j = 1; $j -le $count; $j++) {
                            try {
                                $item = $items.Item($j)
                                $item.Copy().Move($destSub) | Out-Null
                                $copied++
                                $totalItems++
                            }
                            catch {
                                $failed++
                                Write-OMMigrateLog -Message "Copy-OSTFolderRecursive: Item copy failed in '$childPath' [$j]: $_" `
                                                   -Level DEBUG
                            }
                        }
                        if ($copied -gt 0 -or $failed -gt 0) {
                            Write-Host "      '$childPath': $copied copied$(if ($failed -gt 0) { ", $failed failed" })" `
                                       -ForegroundColor DarkGray
                        }
                    }
                }
                catch {
                    Write-OMMigrateLog -Message "Copy-OSTFolderRecursive: Error accessing items in '$childPath': $_" `
                                       -Level WARN
                }

                # Recurse into subfolders
                $childStats = Copy-OSTFolderRecursive `
                    -SourceFolder $srcSub `
                    -DestFolder   $destSub `
                    -FolderPath   $childPath

                $totalFolders += $childStats.Folders
                $totalItems   += $childStats.Items
            }
            catch {
                Write-OMMigrateLog -Message "Copy-OSTFolderRecursive: Error processing subfolder [$i] under '$FolderPath': $_" `
                                   -Level WARN
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Copy-OSTFolderRecursive: Error enumerating subfolders of '$FolderPath': $_" `
                           -Level WARN
    }

    return [PSCustomObject]@{ Folders = $totalFolders; Items = $totalItems }
}






try {

    # ----------------------------------------------------------
    #  STEP 1 -- Gate Check: Verify Script 00 Completed
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Pre-Flight Gate Check' -Step '1 of 7'

    # Read Script 00 manifest -- hard stop if missing or failed
    $step00Manifest = Read-StepManifest -Step 0

    Write-OMMigrateLog -Message (
        "Script 00 manifest validated. " +
        "Accounts discovered: $($step00Manifest.Data.AccountsDiscovered) | " +
        "To migrate: $($step00Manifest.Data.AccountsToMigrate)"
    ) -Level INFO

    Write-Host "  Script 00 manifest verified." -ForegroundColor Green
    Write-Host "  Accounts to migrate: $($step00Manifest.Data.AccountsToMigrate)" `
               -ForegroundColor Gray
    Write-Host "  Total data size    : $($step00Manifest.Data.TotalDataSize)" `
               -ForegroundColor Gray
    Write-Host ''

    # Verify environment
    $envResult = Test-OMMigrateEnvironment
    if (-not $envResult.Passed) {
        Write-OMMigrateLog -Message 'Environment check failed. Cannot continue.' -Level ERROR
        $Script:FinalStatus = 'FAILED'
        exit 1
    }

    Write-Host '  Environment checks passed.' -ForegroundColor Green
    Write-Host ''


    # ----------------------------------------------------------
    #  STEP 2 -- Load Account List from CSV
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Loading Account List' -Step '2 of 7'

    $accountsCsvPath = Get-OMMigrateCsvPath -BaseName 'migration_accounts.csv'

    if (-not (Test-Path $accountsCsvPath)) {
        Write-OMMigrateLog -Message (
            "migration_accounts.csv not found at: $accountsCsvPath`n" +
            "Run Script 00 first to generate this file."
        ) -Level ERROR
        $Script:FinalStatus = 'FAILED'
        exit 1
    }

    $allAccounts = Import-Csv -Path $accountsCsvPath -Encoding UTF8

    # -- Build sanitization map from loaded accounts -----------
    if ($Global:OMMigrate.Sanitize) {
        Initialize-SanitizeMap -Accounts @($allAccounts)
        Write-OMMigrateLog -Message '[SANITIZE] Sanitization active -- output masked.' `
                           -Level INFO
    }

    # Filter to accounts that actually need backing up
    $accountsToBackup = @($allAccounts | Where-Object {
        $_.MigrationAction -eq 'MIGRATE'
    })

    # Accounts to skip (report-only -- no backup needed)
    $accountsToSkip = @($allAccounts | Where-Object {
        $_.MigrationAction -ne 'MIGRATE'
    })

    Write-Host "  Total accounts in CSV : $($allAccounts.Count)" -ForegroundColor Gray
    Write-Host "  Accounts to back up   : $($accountsToBackup.Count)" -ForegroundColor Cyan
    Write-Host "  Accounts skipped      : $($accountsToSkip.Count) (marked SKIP in CSV)" -ForegroundColor DarkGray
    Write-Host ''

    # ADDED (Administrator direction, 2026-07-20): per-account MigrationAction
    # confirmation gate. Root cause of the bug this fixes: an account left set
    # to MigrationAction=MIGRATE in migration_accounts.csv that the operator
    # did NOT actually intend to migrate yet (e.g. forgot to change it to
    # SKIP) previously flowed straight into the backup process and then into
    # Archive pre-build with no checkpoint to catch the mistake -- Administrator
    # confirmed this caused a real cascade of COM failures on a POP3 account
    # that was never meant to be touched this run. Every MigrationAction=MIGRATE
    # account is now paused on here, before any backup or COM work starts.
    # Y proceeds with that account. N automatically writes MigrationAction=SKIP
    # back to migration_accounts.csv for that account (via Update-AccountMigrationAction,
    # the same function Script 02/03 already use for CSV auto-updates -- Administrator
    # confirmed the operator should not have to hand-edit the CSV) and moves on to
    # the next account rather than aborting the whole run. EXIT is handled natively
    # by Confirm-Action's existing Y/N/EXIT convention used throughout this script.
    # Skipped entirely under -Force or -WhatIf, consistent with how every other
    # per-account confirmation in this script already behaves.
    # Declared before the gate below (not inside it) so it is always defined at
    # script scope for Step 5c's Archive pre-build exclusion further down, even
    # when -Force or -WhatIf causes the gate itself to be skipped entirely.
    $migrateGateDeclinedEmails = [System.Collections.Generic.List[string]]::new()
    if (-not $Force -and -not $Global:OMMigrate.WhatIf) {
        foreach ($migrateGateAccount in $accountsToBackup) {
            $migrateGateEmail = $migrateGateAccount.EmailAddress
            Write-Host ''
            Write-Host "  MigrationAction is set to MIGRATE for: $(Invoke-OMMigrateSanitize -Text $migrateGateEmail)" `
                       -ForegroundColor Yellow
            $migrateGateProceed = Confirm-Action `
                -Message      "Proceed with backing up and migrating $(Invoke-OMMigrateSanitize -Text $migrateGateEmail) now? (N to skip this account for now)" `
                -AccountEmail $migrateGateEmail `
                -DefaultYes   $false
            if (-not $migrateGateProceed) {
                Write-Host "  Marking $(Invoke-OMMigrateSanitize -Text $migrateGateEmail) as SKIP in migration_accounts.csv..." `
                           -ForegroundColor Yellow
                Write-OMMigrateLog -Message "MigrationAction confirmation declined by operator for $migrateGateEmail -- auto-updating MigrationAction=SKIP." `
                                   -Level INFO
                Update-AccountMigrationAction -EmailAddress $migrateGateEmail -NewAction 'SKIP'
                [void]$migrateGateDeclinedEmails.Add($migrateGateEmail)
            }
        }
        # Remove declined accounts from this run's backup list -- they are now
        # SKIP in the CSV and must not proceed into backup/Archive pre-build
        # this run, matching what the CSV now says. Also add them to
        # $accountsToSkip so they still get a proper SKIPPED row in the backup
        # report below, instead of silently disappearing from reporting since
        # $accountsToSkip was already built before this gate ran.
        if ($migrateGateDeclinedEmails.Count -gt 0) {
            $migrateGateDeclinedAccounts = @($accountsToBackup | Where-Object {
                $migrateGateDeclinedEmails.Contains($_.EmailAddress)
            })
            $accountsToBackup = @($accountsToBackup | Where-Object {
                -not $migrateGateDeclinedEmails.Contains($_.EmailAddress)
            })
            $accountsToSkip = @($accountsToSkip) + $migrateGateDeclinedAccounts
        }
    }

    if ($accountsToBackup.Count -eq 0) {
        # No POP3 accounts to back up -- this is a normal expected state
        # (e.g. all accounts already migrated to IMAP, or all marked SKIP).
        # Always log as INFO and leave FinalStatus as SUCCESS.
        Write-OMMigrateLog -Message 'No accounts require backup. All are IMAP or Exchange.' `
                           -Level INFO
        Write-Host '  No POP3 accounts to back up. Nothing to do.' -ForegroundColor DarkGray
        Write-Host '  If this is unexpected, review Config\migration_accounts.csv' `
                   -ForegroundColor Gray
    }

    # Add skipped accounts to results immediately
    foreach ($skipped in $accountsToSkip) {
        $skipResult = $skipped.PSObject.Copy()

        # Check if a backup PST exists from a prior run -- if so, populate
        # BackupPath so the report shows the file rather than a blank column.
        # This covers reruns where accounts are now IMAP but were backed up
        # previously as POP3 (the PST is still in the Backups folder).
        $existingBackupPath = ''
        $existingBackupSize = 'N/A'
        if ($Global:OMMigrate.BackupPath) {
            # REVERTED 2026-07-11, Administrator direction. Same as Export-PSTBackup's
            # own backupFile above -- the plain POP3 backup PST is a
            # one-time, profile-independent artifact and must never be
            # profile-suffixed. See that function's fix comment for full
            # context.
            $skipSafeEmail     = Get-SafeFileName -InputString $skipped.EmailAddress
            $skipBackupFile    = Join-Path $Global:OMMigrate.BackupPath "$skipSafeEmail.pst"
            if (Test-Path $skipBackupFile -ErrorAction SilentlyContinue) {
                $existingBackupPath = $skipBackupFile
                try {
                    $existingBackupSize = Format-FileSize -Bytes (Get-Item $skipBackupFile).Length
                }
                catch { $existingBackupSize = 'Unknown' }
            }
        }

        $skipResult | Add-Member -NotePropertyName 'BackupStatus'        -NotePropertyValue 'SKIPPED' -Force
        $skipResult | Add-Member -NotePropertyName 'BackupPath'          -NotePropertyValue $existingBackupPath -Force
        $skipResult | Add-Member -NotePropertyName 'BackupSizeFormatted' -NotePropertyValue $existingBackupSize -Force
        $skipResult | Add-Member -NotePropertyName 'BackupDetail'        -NotePropertyValue "Tag=$($skipped.ProviderTag) -- no backup required" -Force
        [void]$Script:AccountResults.Add($skipResult)
    }


    # ----------------------------------------------------------
    #  STEP 3 -- Pre-Flight Warning and Operator Confirmation
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Pre-Flight Confirmation' -Step '3 of 7'

    # Resolve minimum backup size from settings or parameter
    $minSizeBytes = $MinBackupSizeMB * 1MB
    try {
        $settingMin = $Global:OMMigrate.Settings.BackupVerification.MinimumSizeMB
        if ($null -ne $settingMin) {
            $minSizeBytes = [long]($settingMin * 1MB)
        }
    }
    catch { }

    # RESTORED 2026-07-10, Administrator direction. This warning was briefly removed
    # earlier the same session when -VisibleLaunch was (incorrectly)
    # switched to invisible for Step 5b's Connect-OutlookCOM call -- see
    # that call site's own comment for the live-tested hang this caused
    # and why VisibleLaunch was restored. This warning is accurate again.
    if ($Script:IMAPAccountsToBackup.Count -gt 0) {
        Write-Host ''
        Write-Host '  NOTE: Outlook will open automatically during the OST export phase.' `
                   -ForegroundColor Yellow
        Write-Host '  When it opens, Alt+Tab back to this window to continue.' `
                   -ForegroundColor Yellow
        Write-Host ''
    }

    Show-PreflightWarning `
        -ScriptDescription (
            "This script will export a PST backup file for each of the " +
            "$($accountsToBackup.Count) POP3 account(s) and all IMAP-ALREADY accounts. " +
            "No accounts will be modified. Backup files will be written to: " +
            "$(Invoke-OMMigrateSanitize -Text $Global:OMMigrate.BackupPath)"
        ) `
        -Prerequisites @(
            "Outlook is fully closed (PST files must be unlocked for backup)",
            "You have reviewed Config\migration_accounts.csv",
            "There is sufficient disk space for backups (estimated: $($step00Manifest.Data.TotalDataSize))",
            "You know where to find Secure Mail Keys / app passwords if needed later",
            "IMAP-ALREADY accounts will be backed up via Outlook COM (OST->PST export)"
        ) `
        -DefaultYes $true


    # ----------------------------------------------------------
    #  STEP 4 -- Verify Outlook is Fully Closed
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Verifying Outlook is Closed' -Step '4 of 7'

    # Script 01 does not use Outlook COM -- it copies PST files directly.
    # Outlook must be fully closed so PST files are flushed and unlocked.
    # Previously this step only waited/warned and never closed Outlook
    # itself; now closes automatically via the shared routine for
    # consistency with Scripts 00/02/03/04, while still confirming the
    # closed state to the operator on the console.
    [void](Close-OutlookIfRunning -Reason 'before PST backup')
    Write-Host '  Outlook is closed -- PST files are ready for backup.' `
               -ForegroundColor Green
    Write-OMMigrateLog -Message 'Outlook not running -- safe to copy PST files.' `
                       -Level INFO
    # Brief settle time to ensure any lingering file handles are released
    Start-Sleep -Seconds 2

    Write-Host ''


    # ----------------------------------------------------------
    #  STEP 5 -- Backup Each POP3 Account
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Exporting PST Backups' -Step '5 of 7'

    # Initialize progress tracker for checkpoint support
    $pendingEmails = @($accountsToBackup | ForEach-Object { $_.EmailAddress })
    Update-OMMigrateProgress -SetPending $pendingEmails

    $backupSucceeded = 0
    $backupFailed    = 0
    $backupSkipped   = 0

    foreach ($account in $accountsToBackup) {

        $email = $account.EmailAddress

        Write-Host ''
        Write-Host ('  ' + ('-' * 54)) -ForegroundColor DarkGray
        Write-Host "  Account: $(Invoke-OMMigrateSanitize -Text $email)" -ForegroundColor White
        Write-Host "  Tag    : $($account.ProviderTag)" -ForegroundColor DarkGray
        Write-Host "  Source : $(if ($account.PSTPath) { Invoke-OMMigrateSanitize -Text $account.PSTPath } else { 'Unknown -- will use COM export' })" `
                   -ForegroundColor DarkGray

        # Check if backup already exists and is valid
        # REVERTED 2026-07-11, Administrator direction. Same as Export-PSTBackup's
        # own backupFile -- this pre-check must resolve the exact same
        # bare, profile-independent path that function will actually write
        # to. See Export-PSTBackup's fix comment for full context.
        $safeEmail        = Get-SafeFileName -InputString $email
        $backupFile       = Join-Path $Global:OMMigrate.BackupPath "$safeEmail.pst"
        $safeEmailMasked  = Get-SafeFileName -InputString (Invoke-OMMigrateSanitize -Text $email)
        $backupFileMasked = Join-Path $Global:OMMigrate.BackupPath "$safeEmailMasked.pst"
        $alreadyDone = $false

        if (Test-Path $backupFile) {
            $existingInfo = Get-Item $backupFile
            if ($existingInfo.Length -ge $minSizeBytes) {
                $existingSize = Format-FileSize -Bytes $existingInfo.Length
                Write-Host "  Backup exists: $(Invoke-OMMigrateSanitize -Text $backupFile) ($existingSize)" `
                           -ForegroundColor Green
                Write-OMMigrateLog -Message "Backup already exists and is valid: $(Invoke-OMMigrateSanitize -Text $backupFile) ($existingSize)" `
                                   -Level INFO
                $alreadyDone = $true
            }
            else {
                Write-Host "  Existing backup is too small -- will re-export." `
                           -ForegroundColor Yellow
                Write-OMMigrateLog -Message "Existing backup too small -- will overwrite: $(Invoke-OMMigrateSanitize -Text $backupFile)" `
                                   -Level WARN
            }
        }

        # Mark current account in progress tracker
        Update-OMMigrateProgress -SetCurrent $email

        # Per-account Y/N confirmation (unless -Force or already done)
        $proceed = $true
        if (-not $Force -and -not $alreadyDone) {
            $proceed = Confirm-Action `
                -Message      "Export PST backup for: $(Invoke-OMMigrateSanitize -Text $email) ?" `
                -AccountEmail $email `
                -DefaultYes   $true
        }
        elseif ($alreadyDone -and -not $Force) {
            $proceed = Confirm-Action `
                -Message      "Backup exists. Re-export PST for: $(Invoke-OMMigrateSanitize -Text $email) ? (N to keep existing)" `
                -AccountEmail $email `
                -DefaultYes   $false
        }

        if (-not $proceed) {
            Write-OMMigrateLog -Message "Backup skipped by operator: $email" -Level INFO
            Write-AuditEntry  -Action 'BACKUP_SKIPPED' `
                              -AccountEmail $email `
                              -Detail 'Operator skipped at Y/N prompt' `
                              -Outcome 'SKIPPED'

            $skipResult = $account.PSObject.Copy()
            $skipResult | Add-Member -NotePropertyName 'BackupStatus'        -NotePropertyValue 'SKIPPED' -Force
            $skipResult | Add-Member -NotePropertyName 'BackupPath'          -NotePropertyValue $backupFile -Force
            $skipResult | Add-Member -NotePropertyName 'BackupSizeFormatted' -NotePropertyValue 'N/A' -Force
            $skipResult | Add-Member -NotePropertyName 'BackupDetail'        -NotePropertyValue 'Skipped by operator' -Force
            [void]$Script:AccountResults.Add($skipResult)

            Show-AccountStatus -Email  (Invoke-OMMigrateSanitize -Text $email) `
                               -Tag    $account.ProviderTag `
                               -Action 'Backup skipped by operator' `
                               -Status 'SKIP'
            $backupSkipped++
            Update-OMMigrateProgress -MarkComplete $email
            continue
        }

        # Use existing backup if valid and operator confirmed keep
        if ($alreadyDone -and -not $proceed) {
            $existingInfo = Get-Item $backupFile
            $existResult  = $account.PSObject.Copy()
            $existResult  | Add-Member -NotePropertyName 'BackupStatus' `
                                       -NotePropertyValue 'SUCCESS' -Force
            $existResult  | Add-Member -NotePropertyName 'BackupPath' `
                                       -NotePropertyValue $backupFile -Force
            $existResult  | Add-Member -NotePropertyName 'BackupSizeFormatted' `
                                       -NotePropertyValue (Format-FileSize -Bytes $existingInfo.Length) -Force
            $existResult  | Add-Member -NotePropertyName 'BackupDetail' `
                                       -NotePropertyValue 'Existing verified backup retained' -Force
            [void]$Script:AccountResults.Add($existResult)
            [void]$Script:VerifiedBackups.Add($backupFile)
            $Script:TotalBackupBytes += $existingInfo.Length
            $backupSucceeded++
            Update-OMMigrateProgress -MarkComplete $email
            continue
        }

        # -- Execute backup -------------------------------------
        Show-AccountStatus -Email  (Invoke-OMMigrateSanitize -Text $email) `
                           -Tag    $account.ProviderTag `
                           -Action 'Exporting PST backup...' `
                           -Status 'INFO'

        Write-OMMigrateLog -Message "Starting PST backup: $(Invoke-OMMigrateSanitize -Text $email) -> $backupFileMasked" `
                           -Level INFO
        Write-AuditEntry  -Action 'BACKUP_STARTED' `
                          -AccountEmail $email `
                          -Detail "Source=$($account.PSTPath) | Destination=$backupFile"

        $backupResult = Export-PSTBackup `
            -Account      $account `
            -BackupFolder $Global:OMMigrate.BackupPath `
            -MinSizeBytes $minSizeBytes

        # Attach backup results to account object for reporting
        $accountResult = $account.PSObject.Copy()
        $accountResult | Add-Member -NotePropertyName 'BackupStatus' `
                                    -NotePropertyValue $backupResult.BackupStatus -Force
        $accountResult | Add-Member -NotePropertyName 'BackupPath' `
                                    -NotePropertyValue $backupResult.BackupPath -Force
        $accountResult | Add-Member -NotePropertyName 'BackupSizeFormatted' `
                                    -NotePropertyValue $backupResult.BackupSizeFormatted -Force
        $accountResult | Add-Member -NotePropertyName 'BackupDetail' `
                                    -NotePropertyValue $backupResult.BackupDetail -Force
        [void]$Script:AccountResults.Add($accountResult)

        # Update counters and progress
        if ($backupResult.BackupStatus -eq 'SUCCESS') {
            [void]$Script:VerifiedBackups.Add($backupResult.BackupPath)
            $Script:TotalBackupBytes += $backupResult.BackupSizeBytes
            $backupSucceeded++

            Show-AccountStatus -Email  (Invoke-OMMigrateSanitize -Text $email) `
                               -Tag    $account.ProviderTag `
                               -Action "Backup verified: $($backupResult.BackupSizeFormatted)" `
                               -Status 'OK'

            Write-AuditEntry  -Action 'BACKUP_VERIFIED' `
                              -AccountEmail $email `
                              -Detail (
                                  "File=$($backupResult.BackupPath) | " +
                                  "Size=$($backupResult.BackupSizeFormatted)"
                              )
        }
        else {
            $backupFailed++
            $Script:FinalStatus = 'WARNING'

            Show-AccountStatus -Email  (Invoke-OMMigrateSanitize -Text $email) `
                               -Tag    $account.ProviderTag `
                               -Action "FAILED: $($backupResult.BackupDetail)" `
                               -Status 'FAIL'

            Write-AuditEntry  -Action 'BACKUP_FAILED' `
                              -AccountEmail $email `
                              -Detail $backupResult.BackupDetail `
                              -Outcome 'FAILED'
        }

        Update-OMMigrateProgress -MarkComplete $email
    }

    # Determine final status based on backup results
    if ($backupFailed -gt 0 -and $backupSucceeded -eq 0) {
        $Script:FinalStatus = 'FAILED'
    }
    elseif ($backupFailed -gt 0) {
        $Script:FinalStatus = 'WARNING'
    }
    elseif ($backupSucceeded -eq 0 -and $backupSkipped -eq $accountsToBackup.Count -and
            $accountsToBackup.Count -gt 0) {
        # Only warn about zero backups when there were actual POP3 accounts to process
        $Script:FinalStatus = 'WARNING'
    }

    Write-Host ''
    Write-Host ('  ' + ('-' * 54)) -ForegroundColor DarkGray
    Write-Host "  Backup Summary:" -ForegroundColor White
    Write-Host "    Succeeded : $backupSucceeded" -ForegroundColor $(
        if ($backupSucceeded -gt 0) { 'Green' } else { 'Gray' }
    )
    Write-Host "    Failed    : $backupFailed" -ForegroundColor $(
        if ($backupFailed -gt 0) { 'Red' } else { 'Gray' }
    )
    Write-Host "    Skipped   : $backupSkipped" -ForegroundColor Gray
    Write-Host "    Total size: $(Format-FileSize -Bytes $Script:TotalBackupBytes)" `
               -ForegroundColor Gray
    Write-Host ''

    if ($backupFailed -gt 0) {
        Write-Host '  WARNING: Some backups failed.' -ForegroundColor Red
        Write-Host '  DO NOT proceed to Script 02 until all backups are verified.' `
                   -ForegroundColor Red
        Write-Host '  Review the Backup Report for details.' -ForegroundColor Yellow
        Write-Host ''
    }


    # ----------------------------------------------------------
    #  STEP 5b -- Export OST Backups for IMAP-ALREADY Accounts
    #  PST pre-creation already ran before the main try{} block.
    #  This step opens Outlook COM and copies each account's OST
    #  content into the pre-created PST.
    # ----------------------------------------------------------
    if ($Script:IMAPAccountsToBackup.Count -gt 0) {
        Show-SectionHeader -Title 'Exporting OST Backups for IMAP-ALREADY Accounts'

        Write-Host "  Exporting $($Script:IMAPAccountsToBackup.Count) account(s)..." `
                   -ForegroundColor Cyan
        Write-Host ''

        # -- Open Outlook COM for OST folder copy --------------
        Write-Host '  Launching Outlook COM for OST export...' -ForegroundColor Cyan
        # Read selected profile from Settings.json -- set by Script 00 profile picker.
        # Passed to Connect-OutlookCOM so Outlook opens the correct migration profile.
        $s01ProfileName = ''
        try {
            if ($Global:OMMigrate.Settings -and
                $Global:OMMigrate.Settings.PSObject.Properties['OutlookProfile'] -and
                $Global:OMMigrate.Settings.OutlookProfile.SelectedProfile) {
                $s01ProfileName = $Global:OMMigrate.Settings.OutlookProfile.SelectedProfile
            }
        }
        catch { }
        Write-OMMigrateLog -Message "Using Outlook profile: '$s01ProfileName'" -Level INFO
        # REVERTED 2026-07-10, Administrator direction. Earlier this session,
        # -VisibleLaunch $true was removed here (see git history / prior
        # comment) on the theory it was an accidental copy from Script 02's
        # pattern, since Script 01's OST export has no operator UI
        # interaction. Live-tested and found WRONG: with the invisible
        # Connect-OutlookCOM path, $Namespace.AddStore($backupFile) inside
        # Export-OSTBackup hung indefinitely (confirmed: no further log
        # output for several minutes past "Found IMAP store", blinking
        # cursor, no progress -- required Ctrl+C and manually ending the
        # OUTLOOK.EXE process, since the invisible session never got a
        # chance to run its own cleanup). Administrator's own recollection was that
        # OST export previously worked with Outlook visible in the
        # foreground -- restored -VisibleLaunch $true here on that basis.
        # Root cause of why AddStore() specifically needs a visible/
        # foreground Outlook session (vs. every other invisible COM step in
        # this file, e.g. Archive pre-build, which works fine invisible) is
        # NOT understood yet -- this revert restores known-working behavior
        # rather than resolving the underlying mechanism. Flagged for
        # further investigation; do not remove VisibleLaunch here again
        # without first confirming AddStore() no longer hangs invisibly.
        $outlook = Connect-OutlookCOM -VisibleLaunch $true -ProfileName $s01ProfileName
        if (-not $outlook -and -not $Script:IsWhatIf) {
            Write-Host '  ERROR: Could not start Outlook COM session.' -ForegroundColor Red
            Write-Host '  OST backup skipped -- retry by running this script again.' `
                       -ForegroundColor Yellow
            Write-OMMigrateLog -Message 'OST backup: Failed to start Outlook COM -- skipped.' `
                               -Level ERROR
            $Script:FinalStatus = 'WARNING'
        }
        else {
            if ($outlook) { $Script:COMSessionOpen = $true }
            Suspend-OutlookSendReceive | Out-Null
            $namespace = Get-OutlookNamespace
            Write-Host '  Outlook COM ready.' -ForegroundColor Green
            Write-Host ''

            foreach ($account in $Script:IMAPAccountsToBackup) {
                $email = $account.EmailAddress

                Write-Host ('  ' + ('-' * 54)) -ForegroundColor DarkCyan
                Write-Host "  Account: $(Invoke-OMMigrateSanitize -Text $email)" `
                           -ForegroundColor White
                Write-Host "  Tag    : $($account.ProviderTag)" -ForegroundColor DarkGray
                Write-Host ''

                $safeEmail  = Get-SafeFileName -InputString $email
                # FIXED 2026-07-10, Administrator direction. Same profile-suffix fix
                # as Export-OSTBackup and the pre-creation pass above --
                # this display/prompt path must match those exactly.
                $backupFile = Get-OMMigrateCsvPath -BaseName "${safeEmail}_osttoimap.pst" -BasePathOverride $Global:OMMigrate.BackupPath

                # Per-account Y/N prompt
                $proceedIMAP = $true
                if (-not $Force) {
                    $proceedIMAP = Confirm-Action `
                        -Message      "Export OST backup for: $(Invoke-OMMigrateSanitize -Text $email) ?" `
                        -AccountEmail $email `
                        -DefaultYes   $true
                }

                if (-not $proceedIMAP) {
                    Write-OMMigrateLog -Message "OST backup: Skipped: $email" `
                                       -Level INFO
                    $skipResult = $account.PSObject.Copy()
                    $skipResult | Add-Member -NotePropertyName 'BackupStatus' `
                                             -NotePropertyValue 'SKIPPED' -Force
                    $skipResult | Add-Member -NotePropertyName 'BackupPath' `
                                             -NotePropertyValue '' -Force
                    $skipResult | Add-Member -NotePropertyName 'BackupSizeFormatted' `
                                             -NotePropertyValue 'N/A' -Force
                    $skipResult | Add-Member -NotePropertyName 'BackupDetail' `
                                             -NotePropertyValue 'Skipped by operator' -Force
                    [void]$Script:AccountResults.Add($skipResult)
                    $Script:IMAPBackupSkipped++
                    continue
                }

                # -- Execute OST export ----------------
                Show-AccountStatus -Email  (Invoke-OMMigrateSanitize -Text $email) `
                                   -Tag    $account.ProviderTag `
                                   -Action 'Exporting OST -> PST backup...' `
                                   -Status 'INFO'

                Write-OMMigrateLog -Message "OST backup: Starting export: $email -> $backupFile" `
                                   -Level INFO
                Write-AuditEntry  -Action 'OST_BACKUP_STARTED' `
                                  -AccountEmail $email `
                                  -Detail "Destination=$backupFile"

                $ostResult = Export-OSTBackup `
                    -Account      $account `
                    -BackupFolder $Global:OMMigrate.BackupPath `
                    -Namespace    $namespace `
                    -MinSizeBytes $minSizeBytes

                $accountResult = $account.PSObject.Copy()
                $accountResult | Add-Member -NotePropertyName 'BackupStatus' `
                                            -NotePropertyValue $ostResult.BackupStatus -Force
                $accountResult | Add-Member -NotePropertyName 'BackupPath' `
                                            -NotePropertyValue $ostResult.BackupPath -Force
                $accountResult | Add-Member -NotePropertyName 'BackupSizeFormatted' `
                                            -NotePropertyValue $ostResult.BackupSizeFormatted -Force
                $accountResult | Add-Member -NotePropertyName 'BackupDetail' `
                                            -NotePropertyValue $ostResult.BackupDetail -Force
                # Override OSTPath with the actual OST file path captured from the store
                if ($ostResult.OSTSourcePath) {
                    $accountResult | Add-Member -NotePropertyName 'OSTPath' `
                                                -NotePropertyValue $ostResult.OSTSourcePath -Force
                }
                [void]$Script:AccountResults.Add($accountResult)

                if ($ostResult.BackupStatus -eq 'SUCCESS') {
                    [void]$Script:VerifiedBackups.Add($ostResult.BackupPath)
                    $Script:TotalBackupBytes += $ostResult.BackupSizeBytes
                    $Script:IMAPBackupSucceeded++

                    Show-AccountStatus `
                        -Email  (Invoke-OMMigrateSanitize -Text $email) `
                        -Tag    $account.ProviderTag `
                        -Action "Backup verified: $($ostResult.BackupSizeFormatted) | Folders=$($ostResult.FoldersCopied) | Items=$($ostResult.ItemsCopied)" `
                        -Status 'OK'

                    Write-AuditEntry -Action 'OST_BACKUP_VERIFIED' `
                                     -AccountEmail $email `
                                     -Detail (
                                         "File=$($ostResult.BackupPath) | " +
                                         "Size=$($ostResult.BackupSizeFormatted) | " +
                                         "Folders=$($ostResult.FoldersCopied) | " +
                                         "Items=$($ostResult.ItemsCopied)"
                                     )
                }
                else {
                    $Script:IMAPBackupFailed++
                    if ($Script:FinalStatus -eq 'SUCCESS') { $Script:FinalStatus = 'WARNING' }

                    Show-AccountStatus `
                        -Email  (Invoke-OMMigrateSanitize -Text $email) `
                        -Tag    $account.ProviderTag `
                        -Action "FAILED: $($ostResult.BackupDetail)" `
                        -Status 'FAIL'

                    Write-AuditEntry -Action 'OST_BACKUP_FAILED' `
                                     -AccountEmail $email `
                                     -Detail $ostResult.BackupDetail `
                                     -Outcome 'FAILED'
                }
            }

            # Release COM after all OST exports
            try { Resume-OutlookSendReceive | Out-Null } catch { }
            try { Release-OutlookCOM } catch { }
            $Script:COMSessionOpen = $false
            Write-Host ''
            Write-Host '  Outlook COM released.' -ForegroundColor DarkGray

            Write-Host ''
            Write-Host '  OST Backup Summary:' -ForegroundColor White
            Write-Host "    Succeeded : $Script:IMAPBackupSucceeded" -ForegroundColor $(
                if ($Script:IMAPBackupSucceeded -gt 0) { 'Green' } else { 'Gray' }
            )
            Write-Host "    Failed    : $Script:IMAPBackupFailed" -ForegroundColor $(
                if ($Script:IMAPBackupFailed -gt 0) { 'Red' } else { 'Gray' }
            )
            Write-Host "    Skipped   : $Script:IMAPBackupSkipped" -ForegroundColor Gray
            Write-Host ''
        }
    }
    elseif ($Script:IMAPAccountsToBackup.Count -eq 0) {
        Write-OMMigrateLog -Message 'IMAP-ALREADY: No accounts found or pre-creation failed -- skipping OST export.' `
                           -Level INFO
    }




    # ----------------------------------------------------------
    #  STEP 5c -- Pre-Build Archive PST Folder Structures
    #  For each successfully backed-up account, open its backup
    #  PST and build matching folder hierarchy in the Archive PST.
    #  This ensures folder_map.csv has complete Local rows on the
    #  folder_map.csv -- no back-and-forth between scripts required.
    #
    #  POP3/COMPLETE/IMAP-CONVERTED backup path: safeEmail.pst
    #  IMAP-ALREADY backup path: safeEmail_osttoimap.pst
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Pre-Building Archive PST Folder Structures' -Step '5c of 7'

    # Build Archive folder structures for all eligible accounts.
    # Source is determined per-account: first run walks backup PST,
    # subsequent runs read folder_map.csv for any admin-added rows.
    # ADDED (Administrator direction, 2026-07-20): also exclude accounts declined
    # at the Step 2 MigrationAction confirmation gate above. $allAccounts was
    # loaded into memory before that gate ran and matches on ProviderTag alone,
    # so without this exclusion a just-declined POP3 account would still reach
    # this exact step -- the same Archive pre-build step that produced the
    # original COM failure cascade this whole fix addresses.
    $archiveBuildAccounts = @(
        $allAccounts | Where-Object {
            (-not $migrateGateDeclinedEmails.Contains($_.EmailAddress)) -and
            (
                $_.ProviderTag -eq 'IMAP-ALREADY' -or
                $_.ProviderTag -eq 'IMAP-CONVERTED' -or
                $_.ProviderTag -like 'POP3-*' -or
                $_.MigrationAction -eq 'COMPLETE'
            )
        }
    )

    if ($archiveBuildAccounts.Count -eq 0) {
        Write-Host '  No eligible accounts found -- Archive pre-build skipped.' `
                   -ForegroundColor DarkGray
        Write-OMMigrateLog -Message 'Archive pre-build: No eligible accounts -- skipped.' -Level INFO
    }
    else {
        Write-Host "  Building Archive folder structures for $($archiveBuildAccounts.Count) account(s)..." `
                   -ForegroundColor Cyan
        Write-OMMigrateLog -Message "Archive pre-build: Processing $($archiveBuildAccounts.Count) account(s)." `
                           -Level INFO

        # Open Outlook COM session for Archive PST operations
        $archiveOutlook = Connect-OutlookCOM -ProfileName $Global:OMMigrate.Settings.OutlookProfile.SelectedProfile
        if (-not $archiveOutlook) {
            Write-Host '  ERROR: Could not open Outlook COM session -- Archive pre-build skipped.' `
                       -ForegroundColor Red
            Write-OMMigrateLog -Message 'Archive pre-build: Could not open Outlook COM -- skipped.' `
                               -Level WARN
        }
        else {
            $Script:COMSessionOpen = $true

            # Added 2026-07-10, Administrator (multi-archive support, TargetStoreName
            # hardcode fix). Historically this block always opened ONE archive
            # PST (OMMigrate_Archive.pst / 'OMMigrate Local Archive') and built
            # every eligible account into it. That breaks now that accounts can
            # be mapped to DIFFERENT archive PSTs via the ArchiveStoreMappings
            # picker (Script 00). This block now resolves each account's own
            # TargetStoreName from ArchiveStoreMappings and groups accounts by
            # their resolved archive store, so each distinct attached PST is
            # opened once and only the accounts mapped to it are built there.
            # An account with no mapping entry (or when ArchiveStoreMappings is
            # empty/not configured -- e.g. an older settings file, or a
            # single-archive user who never needed the picker) falls back to
            # the original single default archive PST/name, so existing
            # single-archive setups are completely unaffected by this change.
            # CHANGED (Administrator direction, 2026-08-18 -- deferred to v1.6.0): was
            # `Join-Path $Global:OMMigrate.BackupPath 'OMMigrate_Archive.pst'`,
            # which correctly followed -BasePath in principle but exposed a real
            # gap live-tested tonight -- a custom -BasePath pointed at a location
            # that has never had Install.ps1 run against it has no
            # OMMigrate_Archive.pst there yet, and safely creating one requires an
            # isolated Outlook COM session (see the now-reverted attempt just
            # below, which hung when it instead reused the already-open Script 01
            # session). Administrator's explicit direction: pin this to the fixed default
            # location for now -- OMMigrate_Archive.pst is the one persistent,
            # accumulating destination PST (unlike the per-account
            # <email>_osttoimap.pst backups, which DO correctly follow -BasePath,
            # confirmed working live tonight), and C: is a safe assumption for any
            # real workstation. v1.6.0 is expected to split -BasePath into a
            # data path and a separate, dedicated archive path so this can be
            # revisited properly rather than patched again in isolation.
            $defaultArchivePSTPath     = Join-Path (Join-Path $env:USERPROFILE 'Documents\OutlookMigration\Backups') 'OMMigrate_Archive.pst'
            $defaultArchiveStoreName   = 'OMMigrate Local Archive'

            $archiveStoreMappings = @()
            try {
                if ($Global:OMMigrate.Settings.RulesEngine.ArchiveStoreMappings) {
                    $archiveStoreMappings = @($Global:OMMigrate.Settings.RulesEngine.ArchiveStoreMappings)
                }
            } catch { }

            # Build email -> TargetStoreName lookup from the mapping entries.
            # Unmapped accounts are simply absent from this hashtable, and are
            # resolved to the default archive store name below.
            $emailToTargetStoreName = @{}
            foreach ($mappingEntry in $archiveStoreMappings) {
                $mappedStoreName = $mappingEntry.TargetStoreName
                if ([string]::IsNullOrWhiteSpace($mappedStoreName)) { continue }
                foreach ($mappedEmail in @($mappingEntry.RuleStoreNames)) {
                    if ([string]::IsNullOrWhiteSpace($mappedEmail)) { continue }
                    $emailToTargetStoreName[$mappedEmail] = $mappedStoreName
                }
            }

            # Group this run's eligible accounts by their resolved archive
            # store name (defaulting unmapped accounts to the original single
            # archive store name so behavior is unchanged when no mapping
            # applies to them).
            $accountsByStoreName = [ordered]@{}
            foreach ($acctForGrouping in $archiveBuildAccounts) {
                $resolvedStoreName = $defaultArchiveStoreName
                if ($emailToTargetStoreName.ContainsKey($acctForGrouping.EmailAddress)) {
                    $resolvedStoreName = $emailToTargetStoreName[$acctForGrouping.EmailAddress]
                }
                if (-not $accountsByStoreName.Contains($resolvedStoreName)) {
                    $accountsByStoreName[$resolvedStoreName] = [System.Collections.Generic.List[object]]::new()
                }
                $accountsByStoreName[$resolvedStoreName].Add($acctForGrouping)
            }

            $totalCreated = 0
            $totalFailed  = 0
            $totalSkipped = 0

            # Track every distinct archive store name actually opened this run
            # so the folder_map.csv update pass below can enumerate all of
            # them, not just one hardcoded name.
            $archiveStoresOpenedThisRun = [System.Collections.Generic.List[string]]::new()

            foreach ($archiveStoreGroupName in $accountsByStoreName.Keys) {

            # Resolve this store name to a PST file path. The default archive
            # store keeps its original fixed path; any other TargetStoreName
            # (a user-attached PST picked in Script 00) is resolved from the
            # live list of currently attached stores rather than assumed --
            # matches the resolution pattern already used in
            # Invoke-DeployConsolidatedRules for the same ArchiveStoreMappings
            # feature.
            if ($archiveStoreGroupName -eq $defaultArchiveStoreName) {
                $archivePSTPath  = $defaultArchivePSTPath
                $archiveDisplayName = $defaultArchiveStoreName

                # ADDED (Administrator direction, 2026-08-18 -- deferred to v1.6.0): an
                # earlier version of this fix attempted to auto-create a missing
                # OMMigrate_Archive.pst at a custom -BasePath via Outlook COM
                # AddStore(), reusing the $archiveOutlook session already open at
                # this point in Script 01. Live-tested and found to hang
                # indefinitely (confirmed live 2026-08-18 -- AddStore() against an
                # already-open, visible, foreground Outlook session behaves very
                # differently than Install.ps1's own New-ArchivePSTViaCOM, which
                # only ever runs against a brand-new, invisible, standalone
                # Outlook.Application instance with nothing else attached).
                # Per Administrator's direction, this is deferred to v1.6.0 rather than
                # rushed under that COM-concurrency risk tonight -- $defaultArchivePSTPath
                # itself is now permanently pinned to the fixed default C: drive
                # location regardless of -BasePath (see its own declaration comment
                # above), so this branch simply falls through to the existing
                # Open-PSTFile call below exactly as it always has. v1.6.0 is
                # expected to revisit -BasePath's scope entirely (splitting it into
                # a data path and a separate archive path) rather than patch this
                # one file in isolation.
            }
            else {
                $archivePSTPath     = $null
                $archiveDisplayName = $archiveStoreGroupName
                try {
                    # Corrected 2026-07-10, Administrator: verified against actual code in
                    # OMMigrate-Outlook_WIP.psm1 rather than inferred. Two things
                    # confirmed by reading the real function definitions:
                    #   1) $archiveOutlook (the Application object returned by
                    #      Connect-OutlookCOM) has no .Session property.
                    #   2) The module already exports Get-OutlookNamespace, which
                    #      returns the cached $Script:OutlookNamespace set inside
                    #      Connect-OutlookCOM itself -- this is the established,
                    #      correct way to get the namespace, and Script 01 already
                    #      calls it elsewhere in this same file (Get-OutlookNamespace,
                    #      used earlier in this script), so using it here matches
                    #      Script 01's own existing convention exactly rather than
                    #      introducing a second, different way to get the namespace.
                    $archiveNamespaceForLookup = Get-OutlookNamespace
                    $liveStoresForLookup = $archiveNamespaceForLookup.Stores
                    for ($lsi2 = 1; $lsi2 -le $liveStoresForLookup.Count; $lsi2++) {
                        $lookupStore2 = $liveStoresForLookup.Item($lsi2)
                        if ($lookupStore2.DisplayName -eq $archiveStoreGroupName) {
                            try { $archivePSTPath = $lookupStore2.FilePath } catch { }
                            break
                        }
                    }
                } catch { }

                if ([string]::IsNullOrWhiteSpace($archivePSTPath)) {
                    Write-Host "  WARNING: Mapped archive store '$archiveStoreGroupName' is not currently attached -- accounts mapped to it will be skipped this run." `
                               -ForegroundColor Yellow
                    Write-OMMigrateLog -Message "Archive pre-build: Mapped archive store '$archiveStoreGroupName' not found among attached stores -- skipping $($accountsByStoreName[$archiveStoreGroupName].Count) mapped account(s)." `
                                       -Level WARN
                    $totalSkipped += $accountsByStoreName[$archiveStoreGroupName].Count
                    continue
                }
            }

            # FIXED 2026-07-10, Administrator direction (live-tested bug: this archive
            # open/close was NOT paired with a mounted-state check, unlike the
            # backup PST below in this same function -- confirmed live when
            # 'OMMigrate Local Archive', which Administrator had manually attached in
            # Outlook, was found detached after a normal Script 01 run.
            # Two protections, either one is sufficient to skip the close:
            #   1. $archiveWasAlreadyMounted -- same Test-PSTAlreadyMounted
            #      pattern already proven for the backup PST just below,
            #      applied here identically. Only detach what THIS run's own
            #      code mounted.
            #   2. $archiveIsMasterArchive -- some archives (e.g. 'OMMigrate
            #      Local Archive', a personal auto-archive store) are meant
            #      to stay permanently attached regardless of who mounted
            #      them this run. Operator-configured via
            #      RulesEngine.MasterArchiveNames in Settings.json (see
            #      Get-DefaultSettings in OMMigrate-Core.psm1) rather than
            #      hardcoded here, so adding a protected archive is a
            #      settings edit, not a code change. Matched by exact
            #      DisplayName. Empty list (default) means this check simply
            #      never protects anything extra -- behavior unaffected for
            #      operators who haven't configured any master archives.
            $archiveIsMasterArchive = $false
            try {
                if ($Global:OMMigrate.Settings.PSObject.Properties['RulesEngine'] -and
                    $Global:OMMigrate.Settings.RulesEngine.PSObject.Properties['MasterArchiveNames']) {
                    $masterArchiveNamesForCheck = @($Global:OMMigrate.Settings.RulesEngine.MasterArchiveNames)
                    if ($masterArchiveNamesForCheck -contains $archiveDisplayName) {
                        $archiveIsMasterArchive = $true
                    }
                }
            }
            catch { }

            $archiveWasAlreadyMounted = Test-PSTAlreadyMounted -PSTPath $archivePSTPath

            # Open this archive PST
            $archiveStore = Open-PSTFile -PSTPath     $archivePSTPath `
                                         -DisplayName $archiveDisplayName

            if (-not $archiveStore) {
                Write-Host "  ERROR: Could not open Archive PST '$archiveDisplayName' -- accounts mapped to it are skipped." `
                           -ForegroundColor Red
                Write-OMMigrateLog -Message "Archive pre-build: Could not open Archive PST '$archiveDisplayName' at $archivePSTPath -- skipped." `
                                   -Level WARN
                $totalSkipped += $accountsByStoreName[$archiveStoreGroupName].Count
                continue
            }
            else {
                $archiveStoresOpenedThisRun.Add($archiveDisplayName)

                $archiveRoot    = $archiveStore.GetRootFolder()
                Register-COMObject -ComObject $archiveRoot

                foreach ($acct in $accountsByStoreName[$archiveStoreGroupName]) {
                    $email     = $acct.EmailAddress
                    $safeEmail = Get-SafeFileName -InputString $email

                    # Resolve backup PST path based on ProviderTag
                    # IMAP-ALREADY branch: FIXED 2026-07-10, Administrator direction --
                    # the OST-export backup genuinely differs per profile and
                    # stays profile-suffixed.
                    # Non-IMAP-ALREADY (POP3/legacy conversion) branch:
                    # REVERTED 2026-07-11, Administrator direction -- this is the
                    # plain, one-time POP3 backup PST, a profile-independent
                    # artifact. Suffixing it broke Archive pre-build on
                    # "TestProfile", which was seeded from copies of the Outlook
                    # profile's data files rather than a fresh conversion, so
                    # it never has (and never will) create its own
                    # profile-suffixed copy. See Export-PSTBackup's fix
                    # comment for full context.
                    $providerTag = ($allAccounts | Where-Object { $_.EmailAddress -eq $email } |
                                   Select-Object -First 1).ProviderTag
                    if ($providerTag -eq 'IMAP-ALREADY') {
                        $backupPSTPath = Get-OMMigrateCsvPath -BaseName "${safeEmail}_osttoimap.pst" -BasePathOverride $Global:OMMigrate.BackupPath
                    }
                    else {
                        $backupPSTPath = Join-Path $Global:OMMigrate.BackupPath "${safeEmail}.pst"
                    }

                    Write-Host ''
                    Write-Host ('  ' + ('=' * 54)) -ForegroundColor DarkCyan
                    Write-Host "  Account: $(Invoke-OMMigrateSanitize -Text $email)" `
                               -ForegroundColor White
                    Write-Host ('  ' + ('=' * 54)) -ForegroundColor DarkCyan

                    if (-not (Test-Path $backupPSTPath)) {
                        Write-Host "    [SKIP] Backup PST not found: $(Split-Path $backupPSTPath -Leaf)" `
                                   -ForegroundColor DarkGray

                        # Log-level fix (added 2026-07-01, Administrator): the OST export step
                        # just above (Step 5) is itself skipped entirely and logged at
                        # INFO when $Script:IMAPAccountsToBackup.Count -eq 0 -- e.g. no
                        # IMAP-ALREADY accounts were eligible this run. In that case a
                        # missing backup PST here is a fully expected consequence of
                        # that same INFO-level skip, not a new problem, and console
                        # output already treats it that way (DarkGray, not Yellow).
                        # Only log WARN when the export step actually ran for this
                        # account (was in $Script:IMAPAccountsToBackup) but the file
                        # still isn't present -- that combination is a real,
                        # actionable discrepancy worth flagging.
                        $wasExportAttempted = $false
                        try {
                            $wasExportAttempted = @($Script:IMAPAccountsToBackup | Where-Object { $_.EmailAddress -eq $email }).Count -gt 0
                        } catch { }

                        if ($wasExportAttempted) {
                            Write-OMMigrateLog -Message "Archive pre-build: Backup PST not found for $email -- skipping (export step ran but file is missing -- worth investigating)." `
                                               -Level WARN
                        }
                        else {
                            Write-OMMigrateLog -Message "Archive pre-build: Backup PST not found for $email -- skipping (export step was not attempted this run -- expected)." `
                                               -Level INFO
                        }

                        $totalSkipped++
                        continue
                    }

                    # Open backup PST as source
                    # Check first whether this PST is already mounted in the profile
                    # (e.g. Administrator manually attached it for permanent manual triage use)
                    # so it is not detached later as if it were this script's own
                    # temporary mount.
                    $backupPSTWasAlreadyMounted = Test-PSTAlreadyMounted -PSTPath $backupPSTPath
                    $backupStore = Open-PSTFile -PSTPath     $backupPSTPath `
                                               -DisplayName "ArchiveBuild -- $email"

                    if (-not $backupStore) {
                        Write-Host '    [SKIP] Could not open backup PST.' -ForegroundColor Yellow
                        Write-OMMigrateLog -Message "Archive pre-build: Could not open backup PST for $email -- skipping." `
                                           -Level WARN
                        $totalSkipped++
                        continue
                    }

                    $backupRoot = $backupStore.GetRootFolder()
                    Register-COMObject -ComObject $backupRoot

                    # Create account subfolder in Archive PST
                    $accountArchiveFolder = Get-OrCreateFolder `
                        -ParentFolder $archiveRoot `
                        -FolderName   $email

                    if (-not $accountArchiveFolder) {
                        Write-Host '    [SKIP] Could not create account folder in Archive.' `
                                   -ForegroundColor Yellow
                        Write-OMMigrateLog -Message "Archive pre-build: Could not create account folder for $email." `
                                           -Level WARN
                        # FIXED 2026-07-10, Administrator direction. Live-tested bug found:
                        # $archiveSubCount (whether the account's Archive
                        # subfolder already had content from a prior run) has
                        # nothing to do with whether this backup PST should be
                        # closed -- that decision is solely about
                        # $backupPSTWasAlreadyMounted (Test-PSTAlreadyMounted,
                        # same rule already fixed in Script 03 earlier this
                        # session: only detach what THIS run's own code
                        # mounted). Confirmed live: on a rerun where the
                        # Archive subfolder already had content,
                        # $archiveSubCount was non-zero, which incorrectly
                        # skipped this close and left ArchiveBuild -- <email>
                        # stores attached indefinitely, mirroring the exact
                        # bug already fixed in Script 03. Removed the
                        # unrelated $archiveSubCount condition entirely.
                        if ((Test-Path $backupPSTPath) -and -not $backupPSTWasAlreadyMounted) {
                            # ADDED (Administrator direction, 2026-07-20): same close-failure
                            # warning as the main close call further below -- see that fix
                            # comment for full context.
                            $backupPSTClosedEarly = Close-PSTFile -PSTPath $backupPSTPath
                            if (-not $backupPSTClosedEarly) {
                                Write-Host "    WARNING: Could not automatically close backup PST for $(Invoke-OMMigrateSanitize -Text $email)." `
                                           -ForegroundColor Red
                                Write-Host "    Please manually close/remove it from Outlook's folder pane (right-click -> Close)." `
                                           -ForegroundColor Yellow
                                Write-OMMigrateLog -Message "Archive pre-build: Could not close backup PST for $email after retries -- operator must close manually in Outlook: $backupPSTPath" `
                                                   -Level WARN
                            }
                        }
                        $totalSkipped++
                        continue
                    }

                    # Determine source based on whether Archive already has folders
                    # for this account:
                    #   First run  -- Archive subfolder is empty -> walk backup PST
                    #   Subsequent -- Archive subfolder has folders -> read folder_map.csv
                    #                 so admin-added rows are picked up on reruns
                    $acctCreated = 0
                    $acctFailed  = 0
                    $folderWalkAborted = $false

                    $archiveSubCount = 0
                    try { $archiveSubCount = $accountArchiveFolder.Folders.Count } catch { }

                    if ($archiveSubCount -eq 0) {
                        # First run -- walk backup PST and mirror full folder tree
                        Write-OMMigrateLog -Message "Archive pre-build: $email -- first run, sourcing from backup PST." `
                                           -Level DEBUG

                        # Queue of [COMFolder, PathString] pairs for full recursive depth
                        $folderQueue = [System.Collections.Generic.Queue[PSCustomObject]]::new()

                        $backupSubfolders = $backupRoot.Folders
                        for ($i = 1; $i -le $backupSubfolders.Count; $i++) {
                            $topFolder = $backupSubfolders.Item($i)
                            Register-COMObject -ComObject $topFolder
                            $folderQueue.Enqueue([PSCustomObject]@{
                                Folder = $topFolder
                                Path   = $topFolder.Name
                            })
                        }

                        # ADDED (Administrator direction, 2026-07-20): consecutive-failure guard.
                        # Root cause under investigation: a reused/stale backup PST (operator
                        # kept an existing backup rather than re-exporting) may not be fully
                        # attached/stable in the Outlook COM session at the moment this walk
                        # begins, causing Get-FolderByPath/Get-OrCreateFolder to fail on a
                        # broken COM reference repeatedly -- Administrator observed 47
                        # consecutive folder failures with "Cannot bind argument to parameter
                        # 'ComObject' because it is null" in one real run. Rather than silently
                        # grinding through the entire folder queue producing dozens of
                        # cascading errors, a run of consecutive failures now stops this
                        # account's folder walk cleanly with a clear diagnostic message --
                        # Administrator's explicit preference is a graceful exit over an error
                        # cascade. A single isolated failure (e.g. one genuinely bad folder
                        # name) does not trip this -- only a sustained run does, since that
                        # pattern is what indicates a systemic COM/PST problem rather than a
                        # one-off.
                        $consecutiveFolderFailures    = 0
                        $maxConsecutiveFolderFailures = 5

                        while ($folderQueue.Count -gt 0) {
                            $item       = $folderQueue.Dequeue()
                            $srcFolder  = $item.Folder
                            $folderPath = $item.Path

                            $destFolder = Get-FolderByPath `
                                -RootFolder      $accountArchiveFolder `
                                -FolderPath      $folderPath `
                                -CreateIfMissing $true

                            if ($destFolder) {
                                $acctCreated++
                                $consecutiveFolderFailures = 0
                            }
                            else {
                                $acctFailed++
                                $consecutiveFolderFailures++
                                if ($consecutiveFolderFailures -ge $maxConsecutiveFolderFailures) {
                                    Write-Host "    [ABORT] $consecutiveFolderFailures consecutive folder failures -- stopping Archive pre-build for $(Invoke-OMMigrateSanitize -Text $email)." `
                                               -ForegroundColor Red
                                    Write-Host '    The source backup PST may be unstable (e.g. a large or stale reused backup).' `
                                               -ForegroundColor Yellow
                                    Write-Host '    Consider re-exporting a fresh backup for this account and re-running.' `
                                               -ForegroundColor Yellow
                                    Write-OMMigrateLog -Message "Archive pre-build: $consecutiveFolderFailures consecutive folder failures for $email -- aborting this account's folder walk (possible unstable/stale source backup PST). $($folderQueue.Count) remaining folder(s) not attempted." `
                                                       -Level ERROR
                                    $folderWalkAborted = $true
                                    break
                                }
                                continue
                            }

                            # Enqueue subfolders for recursion
                            try {
                                $subs = $srcFolder.Folders
                                for ($k = 1; $k -le $subs.Count; $k++) {
                                    $subFolder = $subs.Item($k)
                                    Register-COMObject -ComObject $subFolder
                                    $folderQueue.Enqueue([PSCustomObject]@{
                                        Folder = $subFolder
                                        Path   = "$folderPath\$($subFolder.Name)"
                                    })
                                }
                            }
                            catch { }
                        }
                    }
                    else {
                        # Subsequent run -- Archive already has folders.
                        # Read folder_map.csv and create any rows not yet in Archive.
                        # This picks up any subfolders the admin added to the CSV.
                        Write-OMMigrateLog -Message "Archive pre-build: $email -- subsequent run, sourcing from folder_map.csv." `
                                           -Level DEBUG

                        $folderMapCsvPath = Get-OMMigrateCsvPath -BaseName 'folder_map.csv'
                        if (Test-Path $folderMapCsvPath) {
                            $accountFolderRows = @(
                                Import-Csv -Path $folderMapCsvPath -Encoding UTF8 |
                                Where-Object {
                                    $_.StoreName -eq $email -and
                                    $_.FolderPath -ne '' -and
                                    $_.Destination -in @('Server', 'Local')
                                }
                            )
                            foreach ($row in $accountFolderRows) {
                                $destFolder = Get-FolderByPath `
                                    -RootFolder      $accountArchiveFolder `
                                    -FolderPath      $row.FolderPath `
                                    -CreateIfMissing $true

                                if ($destFolder) { $acctCreated++ }
                                else             { $acctFailed++  }
                            }
                        }
                        else {
                            Write-OMMigrateLog -Message "Archive pre-build: folder_map.csv not found -- skipping CSV-based update for $email." `
                                               -Level WARN
                        }
                    }

                    if ($archiveSubCount -eq 0) {
                        Write-Host "    Archive folders created : $acctCreated" -ForegroundColor Green
                    }
                    else {
                        Write-Host "    Archive folders verified: $acctCreated" -ForegroundColor DarkGray
                    }
                    if ($acctFailed -gt 0) {
                        Write-Host "    Archive folders failed  : $acctFailed" -ForegroundColor Yellow
                    }
                    if ($folderWalkAborted) {
                        Write-Host "    Folder walk stopped early for $(Invoke-OMMigrateSanitize -Text $email) -- see log for details." `
                                   -ForegroundColor Red
                    }
                    Write-OMMigrateLog -Message "Archive pre-build: $email -- Created=$acctCreated Failed=$acctFailed Aborted=$folderWalkAborted" `
                                       -Level INFO

                    $totalCreated += $acctCreated
                    $totalFailed  += $acctFailed

                    # Corrected 2026-07-10, Administrator direction (twice, same session).
                    # First fix: added the missing -not $backupPSTWasAlreadyMounted
                    # guard. SECOND fix, live-tested: the $archiveSubCount -eq 0
                    # condition added alongside that guard was itself wrong --
                    # $archiveSubCount (whether the account's Archive subfolder
                    # already had content from a prior run) has nothing to do
                    # with close-eligibility, which is governed solely by
                    # $backupPSTWasAlreadyMounted -- the same rule already
                    # established and fixed in Script 03 earlier this session
                    # (only detach what THIS run's own code mounted). Confirmed
                    # live: on a rerun, $archiveSubCount was non-zero (Archive
                    # subfolder already populated), which incorrectly skipped
                    # this close on every rerun and left ArchiveBuild -- <email>
                    # stores attached indefinitely. Removed the unrelated
                    # $archiveSubCount condition entirely -- this now matches
                    # the exact guard pattern used everywhere else in this
                    # project (Script 03's archive close, this file's own
                    # default-archive close further below) with no additional
                    # conditions layered on top.
                    if ((Test-Path $backupPSTPath) -and -not $backupPSTWasAlreadyMounted) {
                        # ADDED (Administrator direction, 2026-07-20): Close-PSTFile now retries
                        # internally (see its own fix comment), but if it still fails after all
                        # attempts, Administrator hit exactly this case and had to manually detach
                        # the PST in Outlook -- previously silently discarded via Out-Null with no
                        # feedback. Now surfaces a clear console warning telling the operator to
                        # close it manually if the automated close could not.
                        $backupPSTClosed = Close-PSTFile -PSTPath $backupPSTPath
                        if (-not $backupPSTClosed) {
                            Write-Host "    WARNING: Could not automatically close backup PST for $(Invoke-OMMigrateSanitize -Text $email)." `
                                       -ForegroundColor Red
                            Write-Host "    Please manually close/remove it from Outlook's folder pane (right-click -> Close)." `
                                       -ForegroundColor Yellow
                            Write-OMMigrateLog -Message "Archive pre-build: Could not close backup PST for $email after retries -- operator must close manually in Outlook: $backupPSTPath" `
                                               -Level WARN
                        }
                    }
                } # end foreach ($acct in $accountsByStoreName[$archiveStoreGroupName])

                # Update folder_map.csv with this archive store's rows so the
                # operator does not need to rerun any script before proceeding
                # to Script 02. This store is still mounted -- enumerate its
                # folders now, before closing it.
                # Added 2026-07-10, Administrator (multi-archive support): now runs
                # once per distinct archive store opened this run (via
                # $archiveDisplayName) instead of a single hardcoded store
                # name, so folder_map.csv gets rows from every attached
                # archive PST that had eligible accounts, not just one.
                Write-Host ''
                Write-Host "  Updating folder_map.csv with '$archiveDisplayName' folder rows..." `
                           -ForegroundColor Cyan
                Write-OMMigrateLog -Message "Archive pre-build: Updating folder_map.csv with '$archiveDisplayName' rows." `
                                   -Level INFO

                try {
                    $archiveFolderRows = [System.Collections.Generic.List[PSCustomObject]]::new()
                    $archiveTreeFolders = Get-FolderTree -StoreDisplayName $archiveDisplayName `
                                                         -ExcludeSystemFolders $true
                    foreach ($f in $archiveTreeFolders) { $archiveFolderRows.Add($f) }

                    if ($archiveFolderRows.Count -gt 0) {
                        $folderMapCsvPath = Get-OMMigrateCsvPath -BaseName 'folder_map.csv'
                        Export-FolderMapCSV -Folders $archiveFolderRows `
                                            -OutputPath $folderMapCsvPath | Out-Null
                        Write-Host "    folder_map.csv updated: $($archiveFolderRows.Count) folder(s) merged from '$archiveDisplayName'." `
                                   -ForegroundColor Green
                        Write-OMMigrateLog -Message "Archive pre-build: folder_map.csv updated with $($archiveFolderRows.Count) folder(s) from '$archiveDisplayName'." `
                                           -Level INFO
                    }
                    else {
                        Write-Host "    No folders found to merge into folder_map.csv from '$archiveDisplayName'." `
                                   -ForegroundColor DarkGray
                        Write-OMMigrateLog -Message "Archive pre-build: No folders found in '$archiveDisplayName' after pre-build." `
                                           -Level WARN
                    }
                }
                catch {
                    Write-OMMigrateLog -Message "Archive pre-build: Failed to update folder_map.csv for '$archiveDisplayName': $_" `
                                       -Level WARN
                    Write-Host "  WARNING: Could not update folder_map.csv with '$archiveDisplayName' rows." `
                               -ForegroundColor Yellow
                }

                # Close this archive PST before moving to the next distinct store --
                # UNLESS it was already mounted before this run, or it is a
                # protected master archive (see checks above this PST's open,
                # near $archiveWasAlreadyMounted / $archiveIsMasterArchive).
                if (-not $archiveWasAlreadyMounted -and -not $archiveIsMasterArchive) {
                    # ADDED (Administrator direction, 2026-07-20): same close-failure warning
                    # pattern as the backup PST close calls above -- see that fix comment
                    # for full context.
                    $archivePSTClosed = Close-PSTFile -PSTPath $archivePSTPath
                    if (-not $archivePSTClosed) {
                        Write-Host "  WARNING: Could not automatically close Archive PST '$archiveDisplayName'." `
                                   -ForegroundColor Red
                        Write-Host "  Please manually close/remove it from Outlook's folder pane (right-click -> Close)." `
                                   -ForegroundColor Yellow
                        Write-OMMigrateLog -Message "Archive pre-build: Could not close Archive PST '$archiveDisplayName' after retries -- operator must close manually in Outlook: $archivePSTPath" `
                                           -Level WARN
                    }
                }
                else {
                    $skipCloseReason = if ($archiveIsMasterArchive) { 'protected master archive' } else { 'was already attached before this run' }
                    Write-OMMigrateLog -Message (
                        "Archive pre-build: Leaving '$archiveDisplayName' attached ($skipCloseReason)."
                    ) -Level INFO
                }
            }

            } # end foreach ($archiveStoreGroupName in $accountsByStoreName.Keys)

            Write-Host ''
            Write-Host "  Archive pre-build complete." -ForegroundColor Green
            Write-Host "    Archive store(s) used    : $($archiveStoresOpenedThisRun.Count)" -ForegroundColor Green
            Write-Host "    Total folders created/verified : $totalCreated" -ForegroundColor Green
            if ($totalFailed -gt 0) {
                Write-Host "    Total folders failed  : $totalFailed" -ForegroundColor Yellow
            }
            if ($totalSkipped -gt 0) {
                Write-Host "    Accounts skipped      : $totalSkipped" -ForegroundColor DarkGray
            }
            Write-OMMigrateLog -Message "Archive pre-build complete: Stores=$($archiveStoresOpenedThisRun.Count) Created=$totalCreated Failed=$totalFailed Skipped=$totalSkipped" `
                               -Level INFO

            # Release COM session opened for Archive pre-build
            try { Release-OutlookCOM } catch { }
            $Script:COMSessionOpen = $false
        }
    }

    # Open folder destination picker so operator can assign Server/Local
    # destinations. folder_map.csv now has both Server (IMAP) and Local
    # (Archive) rows -- picker shows all folders with correct defaults.
    if (-not $Script:IsWhatIf) {
        $folderPickerCsvPath = Get-OMMigrateCsvPath -BaseName 'folder_map.csv'
        if (Test-Path $folderPickerCsvPath) {
            Write-Host ''
            Write-Host '  Opening folder destination picker...' -ForegroundColor Cyan
            Write-Host '  Review Server/Local assignments -- Local rows default to Local.' `
                       -ForegroundColor DarkGray
            Invoke-FolderMapPicker -CsvPath $folderPickerCsvPath
            Write-Host '  Folder destinations saved.' -ForegroundColor Green
            Write-OMMigrateLog -Message 'Folder destination picker completed.' -Level INFO
        }
        else {
            Write-OMMigrateLog -Message 'Folder picker skipped -- folder_map.csv not found.' `
                               -Level WARN
        }
    }

    # ----------------------------------------------------------
    #  STEP 6 -- Generate Report and Manifest
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Generating Report and Manifest' -Step '6 of 7'

    # Generate Backup HTML report
    $Script:ReportFile = New-BackupReport `
        -Accounts           @($Script:AccountResults) `
        -TotalSizeFormatted (Format-FileSize -Bytes $Script:TotalBackupBytes) `
        -BackupFolder       $Global:OMMigrate.BackupPath

    Write-Host "  Backup Report: $Script:ReportFile" -ForegroundColor Green

    # Write Step 01 manifest
    # Script 02 reads this and will not run if Status = FAILED
    # or if VerifiedBackups is empty
    Write-StepManifest -Step 1 -Status $Script:FinalStatus -Data @{
        BackupSucceeded      = $backupSucceeded
        BackupFailed         = $backupFailed
        BackupSkipped        = $backupSkipped
        IMAPBackupSucceeded  = $Script:IMAPBackupSucceeded
        IMAPBackupFailed     = $Script:IMAPBackupFailed
        IMAPBackupSkipped    = $Script:IMAPBackupSkipped
        TotalBackupSizeBytes = $Script:TotalBackupBytes
        TotalBackupSize      = (Format-FileSize -Bytes $Script:TotalBackupBytes)
        BackupFolder         = $Global:OMMigrate.BackupPath
        VerifiedBackups      = @($Script:VerifiedBackups)
        MinSizeMBUsed        = $MinBackupSizeMB
        VerificationSkipped  = $SkipVerification.IsPresent
        IMAPOSTsBackedUp     = $true
        ReportFile           = $Script:ReportFile
    }

    Write-Host ''
    if ($Script:FinalStatus -eq 'SUCCESS') {
        Write-Host '  All backups verified. You may now run:' -ForegroundColor Green
        Write-Host '  .\Scripts\OMMigrate-02-Convert.ps1' -ForegroundColor White
    }
    elseif ($Script:FinalStatus -eq 'WARNING') {
        Write-Host '  Backup completed with warnings.' -ForegroundColor Yellow
        Write-Host '  Review the Backup Report before proceeding to Script 02.' `
                   -ForegroundColor Yellow
    }
    else {
        Write-Host '  Backup FAILED. DO NOT run Script 02.' -ForegroundColor Red
        Write-Host '  Resolve failures and re-run this script.' -ForegroundColor Red
    }
    Write-Host ''

}
catch {
    Write-OMMigrateLog -Message "FATAL ERROR in Script 01: $_" -Level ERROR
    Write-OMMigrateLog -Message "Stack trace: $($_.ScriptStackTrace)" -Level DEBUG
    $Script:FinalStatus = 'FAILED'

    Write-Host ''
    Write-Host '  FATAL ERROR:' -ForegroundColor Red
    Write-Host "  $_" -ForegroundColor Red
    Write-Host "  Log: $($Global:OMMigrate.RunLogFile)" -ForegroundColor Gray
    Write-Host ''
}
finally {
    # -- Always runs -- release COM, write summary ---------------
    if ($Script:COMSessionOpen) {
        Write-OMMigrateLog -Message 'Releasing Outlook COM session...' -Level INFO
        try { Resume-OutlookSendReceive } catch { }
        try { Release-OutlookCOM } catch { }
        $Script:COMSessionOpen = $false
    }

    if ($Global:OMMigrate) {
        $Global:OMMigrate.SessionCompletedNormally = $true
    }

    Complete-OMMigrateSession `
        -Status     $Script:FinalStatus `
        -ReportFile $Script:ReportFile

    # Open backup report in browser
    if ($Script:ReportFile -and (Test-Path $Script:ReportFile) -and -not $Script:IsWhatIf) {
        Open-FileInEditor -FilePath $Script:ReportFile
    }

    # Keep PowerShell window open regardless of how it was launched
    Wait-UserKeypress
}
# ***** END OF FILE *****
