#Requires -Version 5.1
<#
.SYNOPSIS
    OMMigrate-03-Restore.ps1 -- Folder Migration, Email Restore and Rules Update

.DESCRIPTION
    Step 03 of the OutlookMailMigrator (OMMigrate) toolkit.

    This is the final migration script. It migrates your existing
    folder structure and email content into the new IMAP account
    architecture, then updates Outlook Rules to point to the correct
    new folder locations.

    WHAT THIS SCRIPT DOES:
        1. Reads Step 00 and Step 02 manifests (gate checks)
        2. Reads Config\folder_map.csv (operator-defined destinations)
        3. Reads Config\rules_inventory.csv (rules needing folder updates)
        4. For each account with folders to migrate:
               a. Opens the backup PST as a read-only source store
               b. Creates server-side IMAP folders (Destination = Server)
               c. Creates local Archive PST folders (Destination = Local)
               d. Copies email items from backup PST to destinations
               e. Detaches backup PST when done
        5. Recreates Outlook Rules from each backup PST onto the new
               IMAP store, remapping folder targets per folder_map.csv
               (-RecreateRules switch -- best-effort, 685+ rules supported)
        6. Updates Outlook Rules whose target folders changed
        7. Generates the final Migration HTML report
        8. Writes Step 03 manifest -- migration complete

    FOLDER DESTINATION LOGIC (from folder_map.csv):
        Server  -- Folder created on IMAP server, synced to all devices.
                  Visible in webmail and mobile apps.
                  This is what EmailIpGeoAnalyzer reads from.

        Local   -- Folder created in a permanent local Archive PST file.
                  Never touches the mail server.
                  Stays on this machine only.

        Skip    -- Folder not migrated (no longer used -- treated as Local)

    RULES UPDATE:
        Rules with NeedsFolderUpdate = True in rules_inventory.csv
        have their MoveToFolder or CopyToFolder action targets updated
        to point to the new folder locations. Rules targeting Local
        folders point to the Archive PST. Rules targeting Server
        folders point to the IMAP store.

    RESUME SUPPORT:
        Checkpointing tracks which accounts have completed folder
        migration. Re-running resumes from where it stopped.

.PARAMETER BasePath
    Override the default working directory.
    Must match BasePath used in Scripts 00, 01, and 02.
    Default: $env:USERPROFILE\Documents\OutlookMigration

.PARAMETER Preview
    Simulate all operations. No folders created, no items copied,
    no rules updated. All actions logged with [WHATIF] prefix.

.PARAMETER LogLevel
    Logging verbosity: DEBUG | INFO | WARN | ERROR
    Default: INFO

.PARAMETER RefreshRulesOnly
    Skip item migration and run rules update only.
    Use this when you have edited rules_inventory.csv (e.g. corrected
    a folder target path) and want to rerun just the rules update
    without re-copying items from backup PSTs.

.PARAMETER ArchivePSTName
    Name for the permanent local Archive PST file that holds
    all Local-destination folders.
    Default: OMMigrate_Archive.pst

.PARAMETER RecreateRules
    Recreate Outlook Rules from each account's backup PST onto the
    new IMAP store. Rules are recreated identically -- conditions and
    non-folder actions are preserved verbatim. Only the folder target
    pointer is remapped: Server-destination folders point to the IMAP
    store, Local-destination folders point to the Archive PST.
    Rules that already exist by name on the target store are skipped
    (safe to re-run). Best-effort: rules that cannot be recreated are
    logged and counted but do not abort the run.
    Recommended: run after folder migration is complete so all target
    folders exist before rules are recreated.

.PARAMETER Force
    Skip per-account Y/N confirmation prompts.
    Pre-flight confirmation is still required.

.EXAMPLE
    # Standard run
    .\OMMigrate-03-Restore.ps1

.EXAMPLE
    # Process one account at a time (interactive picker)
    # A grid window appears -- select one or more accounts, click OK
    .\OMMigrate-03-Restore.ps1

.EXAMPLE
    # Dry run
    .\OMMigrate-03-Restore.ps1 -Preview

.EXAMPLE
    # Recreate rules from backup PSTs onto IMAP stores
    .\OMMigrate-03-Restore.ps1 -RecreateRules

.EXAMPLE
    # Rules update only -- re-edited rules_inventory.csv
    .\OMMigrate-03-Restore.ps1 -RefreshRulesOnly

.EXAMPLE
    # Custom archive PST name
    .\OMMigrate-03-Restore.ps1 -ArchivePSTName "MyArchive_2026.pst"

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
        OMMigrate-00-Discover.ps1  -- Step00 manifest required
        OMMigrate-02-Convert.ps1   -- Step02 manifest required
        Config\folder_map.csv      -- Destinations set automatically by Script 01
        Config\rules_inventory.csv -- Generated by Script 00

    Outlook must be closed before running.

    Module dependencies (in ..\Modules\ relative to this script):
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
    [switch]$RefreshRulesOnly,

    # Added 2026-07-02, Administrator. Passed through to Invoke-DeployConsolidatedRules --
    # controls how many rules are created before Save() is called against the
    # live Outlook.Rules collection, instead of saving after every single rule.
    # Performance experiment: full rebuild currently takes ~1.5 hours via this
    # module vs ~30 seconds via the equivalent VBA macro; Save() is the most
    # expensive COM operation per rule (serializes the growing
    # PR_RW_RULES_STREAM to disk every call). Default of 1 preserves exact
    # prior behavior (save every rule) -- purely additive, no existing
    # behavior changes unless this is explicitly set higher.
    [Parameter(Mandatory = $false)]
    [int]$SaveBatchSize = 1,

    [Parameter(Mandatory = $false)]
    [string]$ArchivePSTName = 'OMMigrate_Archive.pst',

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$RecreateRules,

    [Parameter(Mandatory = $false)]
    [switch]$Preview,

    [Parameter(Mandatory = $false)]
    [switch]$Sanitize
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Capture Preview (WhatIf) state from explicit -Preview switch parameter
$Script:IsWhatIf = $Preview.IsPresent


# ============================================================
#  REGION: MODULE IMPORT
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
#  REGION: SESSION INITIALIZATION
# ============================================================

Initialize-OMMigrate `
    -ScriptName 'OMMigrate-03-Restore' `
    -BasePath   $BasePath `
    -LogLevel   $LogLevel `
    -IsWhatIf   $Script:IsWhatIf `
    -Sanitize   $Sanitize.IsPresent

Register-ExitHandlers -ScriptStep 3
Show-ExitBanner

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

# Script-level state
$Script:COMSessionOpen     = $false
$Script:FinalStatus        = 'SUCCESS'
$Script:ReportFile         = ''
$Script:AccountResults     = [System.Collections.Generic.List[PSCustomObject]]::new()
$Script:ArchivePSTPath     = ''
$Script:FoldersCreated     = 0
$Script:FoldersVerified    = 0
$Script:ItemsCopied        = 0
$Script:RulesUpdated       = 0
$Script:RulesRecreated     = 0
$Script:FoldersFailed      = 0
$Script:DisabledRules      = [System.Collections.Generic.List[PSCustomObject]]::new()
$Script:SecondaryStoresProcessed = [System.Collections.Generic.List[string]]::new()

# -- Additive rules-reporting counters (added for HTML report enrichment) --
# These supplement the existing $Script:RulesUpdated / $Script:RulesRecreated
# counters above with finer-grained outcome tracking so the Migration Report
# can show Skipped / No-Change / StopProcessing detail regardless of whether
# -RefreshRulesOnly was used. Existing counters above are untouched.
$Script:RulesSkippedTotal               = 0   # already-existed / no action needed (recreation pass + remap pass)
$Script:RulesNoChangeTotal              = 0   # target folder already correct, no update required (remap pass)
$Script:RulesStopProcessingSetTotal     = 0   # StopProcessing successfully applied post-Save (secondary stores)
$Script:RulesStopProcessingFailedTotal  = 0   # StopProcessing could not be applied via COM -- needs manual fix in Outlook Rules Manager

# -- Additive: default-store "not processed this pass" rollup (added 2026-06-18) --
# NOT a failure/error count. The default store (ameritech) is loaded via .rwz
# import, which natively preserves StopProcessing -- it never goes through this
# tool's Create()-based purge/recreate path at all. The ONLY thing that can
# leave it untouched on a given run is the COM Item() enumeration ceiling at
# this rule count, which silently no-ops Phase 1/3/4a for the default store.
# This counter exists purely so the report can say "primary store rules were
# not processed this run" rather than omitting it or miscounting it as a
# secondary-store-style failure.
$Script:RulesDefaultStoreNotProcessedTotal = 0

# Per-store/account rollup for the new "Rules Processing Detail" report table.
# One entry appended per store processed by any of the three rules-update
# code paths (IMAP-ALREADY remap, recreation pass, secondary-store pass).
$Script:RulesStoreSummary = [System.Collections.Generic.List[PSCustomObject]]::new()

# Effective skip flags -- bool copies of switch parameters.
# $RefreshRulesOnly drives $Script:EffectiveSkipFolderMigration so item
# migration can be skipped while rules still run.
$Script:EffectiveSkipFolderMigration = $RefreshRulesOnly.IsPresent

# Initialize rule and archive variables at script scope so they are always
# defined when referenced, regardless of which code paths execute.
$rulesNeedingUpdate = @()
$archiveStore       = $null


# Get-OrCreateFolder and Get-FolderByPath have been moved to OMMigrate-Outlook.psm1

# ============================================================
#  HELPER: Copy-FolderContents
#  Copies all items from a source folder to a destination folder.
# ============================================================

function Copy-FolderContents {
    <#
    .SYNOPSIS
        Copies all mail items from a source MAPIFolder to a
        destination MAPIFolder.

    .DESCRIPTION
        Iterates all items in the source folder and copies each to
        the destination. Handles large folders gracefully with
        progress logging. Skips items that fail individually rather
        than aborting the entire folder.

        Items are COPIED not moved -- the source (backup PST) is
        never modified. The backup remains intact throughout.

    .PARAMETER SourceFolder
        Source MAPIFolder COM object to copy from.

    .PARAMETER DestFolder
        Destination MAPIFolder COM object to copy to.

    .PARAMETER FolderPath
        Display path for logging (e.g. 'Inbox\Vendors').

    .OUTPUTS
        [int] -- Number of items successfully copied.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$SourceFolder,

        [Parameter(Mandatory = $true)]
        [object]$DestFolder,

        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    $copied  = 0
    $failed  = 0

    try {
        $items = $SourceFolder.Items
        $total = $items.Count

        if ($total -eq 0) {
            Write-OMMigrateLog -Message "Folder '$FolderPath' is empty -- skipping copy." `
                               -Level DEBUG
            return 0
        }

        # -- Duplicate detection ------------------------------------
        # If the destination already has items this folder was copied
        # in a previous run. Skip to prevent duplicate emails.
        # The source is always the same read-only backup PST so the
        # content cannot have changed -- re-copying would only duplicate.
        $destItemCount = 0
        try { $destItemCount = $DestFolder.Items.Count } catch { }
        if ($destItemCount -gt 0) {
            Write-OMMigrateLog -Message (
                "Folder '$FolderPath' already has $destItemCount item(s) in destination -- " +
                "skipping to prevent duplicates."
            ) -Level INFO
            Write-Host "    [SKIP]  '$FolderPath' -- already migrated ($destItemCount items in destination)" `
                       -ForegroundColor DarkGray
            return 0
        }

        Write-OMMigrateLog -Message "Copying $total items from '$FolderPath'..." `
                           -Level INFO

        if ($Global:OMMigrate.WhatIf) {
            Write-OMMigrateLog -Message "WhatIf: Would copy $total items from '$FolderPath'" `
                               -Level INFO -WhatIfPrefix
            return $total
        }

        # Copy items in batches with progress logging every 100 items
        for ($i = 1; $i -le $total; $i++) {
            try {
                $item = $items.Item($i)
                $item.Copy().Move($DestFolder) | Out-Null
                $copied++
                $Script:ItemsCopied++

                if ($copied % 100 -eq 0) {
                    Write-OMMigrateLog -Message "  Progress: $copied / $total items copied from '$FolderPath'" `
                                       -Level DEBUG
                    Write-Host "    Copying '$FolderPath': $copied / $total items..." `
                               -ForegroundColor DarkGray
                }
            }
            catch {
                $failed++
                Write-OMMigrateLog -Message "Failed to copy item $i from '$FolderPath': $_" `
                                   -Level WARN
            }
        }

        Write-OMMigrateLog -Message (
            "Folder '$FolderPath' complete: $copied copied, $failed failed."
        ) -Level INFO
    }
    catch {
        Write-OMMigrateLog -Message "Error accessing items in '$FolderPath': $_" -Level WARN
    }

    return $copied
}


# ============================================================
#  HELPER: Migrate-AccountFolders
#  Orchestrates the full folder migration for one account.
# ============================================================

function Invoke-AccountFolderMigration {
    <#
    .SYNOPSIS
        Migrates all folders for a single account according to
        the folder_map.csv destination assignments.

    .DESCRIPTION
        For a given account's backup PST:
            - Opens the backup PST in Outlook as a read-only source
            - Creates Server-destination folders in the IMAP store
            - Creates Local-destination folders in the Archive PST
            - Copies items from backup to the appropriate destination
            - Detaches the backup PST when done

        Called once per account that has folders to migrate.

    .PARAMETER Account
        Account object from migration_accounts.csv.

    .PARAMETER FolderMap
        Array of folder map rows from folder_map.csv, filtered
        to this account's store.

    .PARAMETER Namespace
        Active Outlook MAPI namespace COM object.

    .PARAMETER ArchiveRootFolder
        Root folder of the Archive PST store.

    .OUTPUTS
        PSCustomObject with FoldersMigrated, ItemsCopied,
        FoldersFailed, Outcome, Detail.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Account,

        [Parameter(Mandatory = $true)]
        [array]$FolderMap,

        [Parameter(Mandatory = $true)]
        [object]$Namespace,

        [Parameter(Mandatory = $false)]
        [object]$ArchiveRootFolder = $null,

        # ADDED (fix, multi-archive folder migration gap): optional rules_inventory.csv
        # rows, used to resolve this account's own TargetStoreName to a live archive
        # store folder -- same resolution pattern already proven working in
        # Invoke-DeployConsolidatedRules and RecreateRules (DisplayName match against
        # $Namespace.Stores). When not provided, or when no TargetStoreName is found for
        # this account, or when the named store isn't currently attached, behavior is
        # completely unchanged -- $ArchiveRootFolder (the single default archive) is
        # used exactly as before this fix.
        [Parameter(Mandatory = $false)]
        [array]$RulesInventory = $null
    )

    $email          = $Account.EmailAddress
    $safeEmail      = Get-SafeFileName -InputString $email
    # REVERTED 2026-07-11, Administrator direction. The 2026-07-10 profile-suffix fix
    # was WRONG for this specific file. The plain POP3 backup PST is a
    # ONE-TIME artifact created the first time an account is converted from
    # POP3 -- it is the single authoritative source of that account's
    # historical mail/folders/artifacts, read by this function regardless of
    # which Outlook profile is currently running the pipeline. Confirmed
    # live: "TestProfile" was seeded from copies of the Outlook profile's data
    # files rather than a fresh POP3 conversion, so it never creates its own
    # profile-suffixed copy of this file -- suffixing the lookup here made
    # Script 03 unable to find real historical content that has existed on
    # disk, unsuffixed, since before "TestProfile" was even created. This path is
    # intentionally NOT run through Get-OMMigrateCsvPath's profile-suffix
    # logic; it stays a bare, profile-independent filename everywhere it is
    # referenced (see OMMigrate-01-Backup.ps1's matching fix comments for
    # full context). The _osttoimap OST-export backup path is unaffected --
    # that one genuinely differs per profile and stays suffixed.
    $backupPath     = Join-Path $Global:OMMigrate.BackupPath "$safeEmail.pst"
    $foldersMigrated  = 0
    $foldersCreated   = 0
    $foldersVerified  = 0
    $itemsCopied      = 0
    $foldersFailed    = 0

    Write-OMMigrateLog -Message "Starting folder migration for: $email" -Level INFO

    # Verify backup exists.
    # IMAP-ALREADY accounts have no backup PST -- they were always IMAP so no
    # POP3 data was ever backed up. Skip the backup PST entirely for these accounts;
    # folder structures are still created in the Archive PST, items are not copied.
    $isImapAlready = ($Account.ProviderTag -eq 'IMAP-ALREADY')

    # FIXED (bug found live 2026-07-12, this same fix): the guard below was
    # originally keyed on $isImapAlready alone, which is TRUE for every
    # IMAP-ALREADY account regardless of whether it actually has a backup
    # PST on disk (e.g. ameritech, converted long ago, genuinely has one).
    # That skipped the PST-open block for ameritech too, leaving $backupRoot
    # unset and crashing the whole run. $skipBackupPSTOpen is only set TRUE
    # inside the branch below, where Test-Path has ALREADY confirmed no file
    # exists for this specific account -- not from the account's type alone.
    $skipBackupPSTOpen = $false

    if (-not (Test-Path $backupPath)) {
        if ($isImapAlready) {
            Write-OMMigrateLog -Message (
                "IMAP-ALREADY account $email -- no backup PST expected. " +
                "Folder structures will be created; no items to copy."
            ) -Level INFO
            # ADDED (fix): this branch previously only logged, then fell through to
            # the PST-open call below, which always failed since the PST was just
            # confirmed absent. $backupStore/$backupRoot are set to $null here so the
            # rest of the function (folder-structure creation) still runs exactly as
            # for any other account, just with no items copied -- matching what the
            # log message above already promised.
            $backupStore = $null
            $backupRoot  = $null
            $backupPathWasAlreadyMounted = $false
            $skipBackupPSTOpen = $true
        }
        else {
            Write-OMMigrateLog -Message "Backup PST not found for $email : $backupPath" `
                               -Level WARN
            # ADDED 2026-07-11, Administrator direction: same missing-property fix as
            # the FAILED return below -- see that comment for full context.
            return [PSCustomObject]@{
                FoldersMigrated = 0
                FoldersCreated  = 0
                FoldersVerified = 0
                ItemsCopied     = 0
                FoldersFailed   = 0
                Outcome         = 'SKIPPED'
                Detail          = "Backup PST not found: $backupPath"
            }
        }
    }

    # Open backup PST as source store
    # Check first whether this PST is already mounted in the profile
    # (e.g. Administrator manually attached it for permanent manual triage use)
    # so it is not detached later as if it were this script's own
    # temporary mount.
    # FIXED (bug found live 2026-07-12, this same fix): guard changed from
    # -not $isImapAlready to -not $skipBackupPSTOpen. The account-type flag
    # was wrong here -- it's true for EVERY IMAP-ALREADY account, including
    # ones (like ameritech) that genuinely have a backup PST on disk and need
    # this block to run. $skipBackupPSTOpen is only true when Test-Path above
    # already confirmed no file exists for THIS account, which is the actual
    # condition that should skip this block.
    if (-not $skipBackupPSTOpen) {
    $backupPathWasAlreadyMounted = Test-PSTAlreadyMounted -PSTPath $backupPath
    $backupDisplayName = "Backup -- $email"
    $backupStore = Open-PSTFile -PSTPath      $backupPath `
                                -DisplayName  $backupDisplayName

    if (-not $backupStore -and -not $Global:OMMigrate.WhatIf) {
        # ADDED 2026-07-11, Administrator direction: this early-failure return object
        # was missing FoldersCreated/FoldersVerified, which the caller reads
        # unconditionally ($Script:FoldersCreated += $migrationResult.
        # FoldersCreated). Every other return path in this function includes
        # both properties; a genuine backup-PST-open failure was crashing
        # with "property 'FoldersCreated' cannot be found" instead of
        # reporting the real FAILED outcome cleanly. Purely defensive --
        # does not change when this branch is reached, only what it returns.
        return [PSCustomObject]@{
            FoldersMigrated = 0
            FoldersCreated  = 0
            FoldersVerified = 0
            ItemsCopied     = 0
            FoldersFailed   = 0
            Outcome         = 'FAILED'
            Detail          = "Could not open backup PST: $backupPath"
        }
    }

    # Get backup PST root folder
    $backupRoot = $null
    try {
        if ($backupStore) {
            $backupRoot = $backupStore.GetRootFolder()
            Register-COMObject -ComObject $backupRoot
        }
    }
    catch {
        Write-OMMigrateLog -Message "Could not get root folder of backup PST: $_" -Level WARN
    }
    }
    # Close of the -not $skipBackupPSTOpen guard opened above -- accounts that
    # hit the no-backup-file branch skip the entire backup-PST-open block and
    # proceed here with $backupStore/$backupRoot already set to $null.

    # Find the IMAP store for this account in the active session
    $imapStore = $null
    try {
        $stores = $Namespace.Stores
        Register-COMObject -ComObject $stores

        for ($s = 1; $s -le $stores.Count; $s++) {
            $store = $stores.Item($s)
            Register-COMObject -ComObject $store
            try {
                # Match by display name containing email address
                if ($store.DisplayName -like "*$email*" -or
                    $store.DisplayName -eq $Account.DisplayName) {
                    # Verify it's IMAP (not PST)
                    $filePath = ''
                    try { $filePath = $store.FilePath } catch { }
                    if (-not ($filePath -like '*.pst')) {
                        $imapStore = $store
                        break
                    }
                }
            }
            catch { }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Error finding IMAP store for $email : $_" -Level WARN
    }

    $imapRoot = $null
    if ($imapStore) {
        try {
            $imapRoot = $imapStore.GetRootFolder()
            Register-COMObject -ComObject $imapRoot
        }
        catch {
            Write-OMMigrateLog -Message "Could not get IMAP root folder for $email : $_" `
                               -Level WARN
        }
    }
    else {
        Write-OMMigrateLog -Message (
            "IMAP store not found for $email -- Server-destination folders " +
            "will be skipped. Local folders will still be migrated."
        ) -Level WARN
    }

    # Process each folder in the map for this account
    $serverFolders = @($FolderMap | Where-Object { $_.Destination -eq 'Server' })
    $localFolders  = @($FolderMap | Where-Object { $_.Destination -eq 'Local'  })

    Write-OMMigrateLog -Message (
        "Folder map for $email : " +
        "Server=$($serverFolders.Count) | " +
        "Local=$($localFolders.Count)"
    ) -Level INFO

    Write-Host "    Server folders : $($serverFolders.Count)" -ForegroundColor Cyan
    Write-Host "    Local folders  : $($localFolders.Count)"  -ForegroundColor Cyan
    Write-Host ''

    # -- Migrate Server-destination folders --------------------
    if ($serverFolders.Count -gt 0 -and $imapRoot) {
        Write-Host '    Verifying IMAP server-side folders...' -ForegroundColor Cyan

        foreach ($folderRow in $serverFolders) {
            $folderPath = $folderRow.FolderPath
            $folderName = $folderRow.FolderName

            # Check if folder already exists before creating
            $existingFolder = Get-FolderByPath `
                -RootFolder      $imapRoot `
                -FolderPath      $folderPath `
                -CreateIfMissing $false
            $serverLabel = if ($existingFolder) { '[SERVER]' } else { '[SERVER-NEW]' }
            Write-Host "      $serverLabel $folderPath" -ForegroundColor Gray

            # Create folder in IMAP store if it doesn't exist
            $destFolder = if ($existingFolder) { $existingFolder } else {
                Get-FolderByPath `
                    -RootFolder      $imapRoot `
                    -FolderPath      $folderPath `
                    -CreateIfMissing $true
            }

            if ($destFolder -or $Global:OMMigrate.WhatIf) {
                # Copy items from backup PST to IMAP folder
                if ($backupRoot -and $destFolder) {
                    $srcFolder = Get-FolderByPath `
                        -RootFolder      $backupRoot `
                        -FolderPath      $folderPath `
                        -CreateIfMissing $false

                    if ($srcFolder) {
                        $copied = Copy-FolderContents `
                            -SourceFolder $srcFolder `
                            -DestFolder   $destFolder `
                            -FolderPath   $folderPath
                        $itemsCopied += $copied
                    }
                }
                $foldersMigrated++
                if ($existingFolder) { $foldersVerified++ } else { $foldersCreated++ }
            }
            else {
                Write-OMMigrateLog -Message "Failed to create server folder: $folderPath" `
                                   -Level WARN
                $foldersFailed++
            }
        }
    }
    elseif ($serverFolders.Count -gt 0 -and -not $imapRoot) {
        Write-OMMigrateLog -Message (
            "Skipping $($serverFolders.Count) server folders -- " +
            "IMAP store not found for $email"
        ) -Level WARN
        $foldersFailed += $serverFolders.Count
    }

    # ADDED (fix, multi-archive folder migration gap): resolve this account's own
    # TargetStoreName (from rules_inventory.csv, if provided) to a live attached
    # store, and use it INSTEAD of the single default $ArchiveRootFolder for this
    # account's folder migration. Mirrors the exact resolution pattern already
    # proven working in Invoke-DeployConsolidatedRules and RecreateRules --
    # DisplayName match against $Namespace.Stores. Falls back to the original
    # $ArchiveRootFolder untouched if $RulesInventory is not supplied, if this
    # account has no TargetStoreName set, or if the named store cannot be found
    # among currently attached stores.
    if ($RulesInventory) {
        $acctTargetStoreName = $RulesInventory |
            Where-Object {
                $_.PSObject.Properties['RuleStoreName'] -and $_.RuleStoreName -eq $email -and
                $_.PSObject.Properties['TargetStoreName'] -and
                -not [string]::IsNullOrWhiteSpace($_.TargetStoreName)
            } | Select-Object -ExpandProperty TargetStoreName -First 1

        if ($acctTargetStoreName) {
            $foundAcctStoreFolder = $null
            try {
                $allStoresForAcctLookup = $Namespace.Stores
                for ($asi = 1; $asi -le $allStoresForAcctLookup.Count; $asi++) {
                    $lookupAcctStore = $allStoresForAcctLookup.Item($asi)
                    if ($lookupAcctStore.DisplayName -eq $acctTargetStoreName) {
                        try { $foundAcctStoreFolder = $lookupAcctStore.GetRootFolder() } catch { }
                        break
                    }
                }
            } catch { }

            if ($foundAcctStoreFolder) {
                $ArchiveRootFolder = $foundAcctStoreFolder
            }
            else {
                Write-OMMigrateLog -Message (
                    "Invoke-AccountFolderMigration: TargetStoreName '$acctTargetStoreName' for account " +
                    "'$email' not found among attached stores -- falling back to the default Archive PST."
                ) -Level WARN
            }
        }
    }

    # -- Migrate Local-destination folders ---------------------
    # Always create the account subfolder in the Archive PST regardless of how
    # many Local-destination folders exist. Every account must have a folder
    # structure in the master Archive -- this is by design. Local-destination
    # folder entries are then created under it. Server-destination entries are
    # not created here (they are created in the IMAP store above).
    if ($ArchiveRootFolder) {
        # Check if account subfolder already exists in Archive
        $existingArchiveAcct = $null
        try {
            $archiveSubs = $ArchiveRootFolder.Folders
            for ($ai = 1; $ai -le $archiveSubs.Count; $ai++) {
                $as = $archiveSubs.Item($ai)
                if ($as.Name -eq $email) { $existingArchiveAcct = $as; break }
            }
        } catch { }
        $acctLabel = if ($existingArchiveAcct) { 'Verifying' } else { 'Creating' }
        Write-Host "    $acctLabel Archive PST account folder..." -ForegroundColor Cyan

        # Find or create account subfolder in Archive PST
        $accountArchiveFolder = Get-OrCreateFolder `
            -ParentFolder $ArchiveRootFolder `
            -FolderName   $email

        if ($accountArchiveFolder -or $Global:OMMigrate.WhatIf) {
            # Create Archive folder structure for ALL folders in this account --
            # both Server and Local destination. Every account gets a complete
            # folder structure in the Archive regardless of destination setting.
            # Items are only copied for Local-destination folders (from backup PST).
            # Server-destination folders get the folder structure only, no items.
            $allAccountFolders = @($FolderMap | Where-Object { $_.Destination -in @('Server','Local') })
            if ($allAccountFolders.Count -gt 0) {
                $archiveHeader = if ($existingArchiveAcct) { 'Verifying' } else { 'Creating' }
                Write-Host "    $archiveHeader Archive PST folder structure..." -ForegroundColor Cyan
            }
            foreach ($folderRow in $allAccountFolders) {
                $folderPath = $folderRow.FolderPath
                $isLocal    = ($folderRow.Destination -eq 'Local')

                # Check if folder already exists before creating
                $existingArchiveFolder = Get-FolderByPath `
                    -RootFolder      $accountArchiveFolder `
                    -FolderPath      $folderPath `
                    -CreateIfMissing $false
                $archiveLabel = if ($existingArchiveFolder) { '[ARCHIVE]' } else { '[ARCHIVE-NEW]' }
                Write-Host "      $archiveLabel $folderPath" -ForegroundColor DarkGray

                # Create folder path in Archive PST under account subfolder
                $destFolder = if ($existingArchiveFolder) { $existingArchiveFolder } else {
                    Get-FolderByPath `
                        -RootFolder      $accountArchiveFolder `
                        -FolderPath      $folderPath `
                        -CreateIfMissing $true
                }

                if ($destFolder -or $Global:OMMigrate.WhatIf) {
                    # Copy items from backup PST only for Local-destination folders
                    if ($isLocal -and $backupRoot -and $destFolder) {
                        $srcFolder = Get-FolderByPath `
                            -RootFolder      $backupRoot `
                            -FolderPath      $folderPath `
                            -CreateIfMissing $false

                        if ($srcFolder) {
                            $copied = Copy-FolderContents `
                                -SourceFolder $srcFolder `
                                -DestFolder   $destFolder `
                                -FolderPath   $folderPath
                            $itemsCopied += $copied
                        }
                    }
                    $foldersMigrated++
                    if ($existingArchiveFolder) { $foldersVerified++ } else { $foldersCreated++ }
                }
                else {
                    $foldersFailed++
                }
            }
        }
    }
    elseif (-not $ArchiveRootFolder) {
        if ($Global:OMMigrate.WhatIf) {
            # In WhatIf mode Open-PSTFile returns $null by design -- this is expected.
            # Archive PST would be opened and account/local folders would be created in a live run.
            Write-OMMigrateLog -Message (
                "[WHATIF] Archive PST open skipped in preview mode -- " +
                "account subfolder and $($localFolders.Count) local folder(s) would be created in live run."
            ) -Level INFO -WhatIfPrefix
        }
        else {
            Write-OMMigrateLog -Message (
                "Archive PST root not available -- " +
                "account subfolder and $($localFolders.Count) local folders skipped."
            ) -Level WARN
            $foldersFailed += $localFolders.Count
        }
    }

    # Detach backup PST -- it served its purpose for this account
    # but only if this script mounted it; leave alone if it was already
    # attached to the profile before this function was called.
    if ($backupStore -and -not $Global:OMMigrate.WhatIf -and -not $backupPathWasAlreadyMounted) {
        Close-PSTFile -PSTPath $backupPath | Out-Null
        Write-OMMigrateLog -Message "Backup PST detached: $backupPath" -Level INFO
    }

    $outcome = if ($foldersFailed -gt 0 -and $foldersMigrated -eq 0) { 'FAILED' }
               elseif ($foldersFailed -gt 0) { 'WARNING' }
               else { 'SUCCESS' }

    return [PSCustomObject]@{
        FoldersMigrated  = $foldersMigrated
        FoldersCreated   = $foldersCreated
        FoldersVerified  = $foldersVerified
        ItemsCopied      = $itemsCopied
        FoldersFailed    = $foldersFailed
        Outcome         = $outcome
        Detail          = "Migrated=$foldersMigrated | Items=$itemsCopied | Failed=$foldersFailed"
    }
}


# ============================================================
#  HELPER: Set-RuleFolderAction
#  Assigns a MAPIFolder COM object to a MoveOrCopyRuleAction
#  using .NET Reflection to bypass the PowerShell CLR translation
#  bug that causes standard property assignment to silently fail.
#
#  Background: PowerShell uses the CLR runtime layer to talk to
#  Outlook COM. When you write $action.Folder = $folder in PS,
#  the CLR translates the assignment into a generic format that
#  Outlook rejects because MoveOrCopyRuleAction.Folder expects a
#  strictly typed Outlook Folder Object -- the underlying pointer
#  structure is stripped by the CLR wrapper and the assignment is
#  silently dropped. VBA bypasses this by calling put_Folder
#  directly via native COM binding. InvokeMember replicates that
#  direct native call from PowerShell.
# ============================================================

function Set-RuleFolderAction {
    <#
    .SYNOPSIS
        Assigns a target MAPIFolder to a MoveOrCopyRuleAction via
        .NET Reflection, bypassing the PowerShell COM translation bug.

    .PARAMETER Action
        The MoveOrCopyRuleAction COM object (e.g. $rule.Actions.MoveToFolder).

    .PARAMETER Folder
        The target MAPIFolder COM object to assign.

    .OUTPUTS
        [bool] -- $true if assignment succeeded, $false on error.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Action,

        [Parameter(Mandatory = $true)]
        [object]$Folder
    )

    try {
        [Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction].InvokeMember(
            "Folder",
            [System.Reflection.BindingFlags]::SetProperty,
            $null,
            $Action,
            $Folder
        )
        return $true
    }
    catch {
        Write-OMMigrateLog -Message "Set-RuleFolderAction: InvokeMember failed: $_" -Level WARN
        return $false
    }
}


# ============================================================
#  MAIN EXECUTION BLOCK
# ============================================================

try {

    # ----------------------------------------------------------
    #  STEP 1 -- Gate Checks
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Pre-Flight Gate Checks' -Step '1 of 6'

    # Gate 1: Script 00 manifest
    $step00Manifest = Read-StepManifest -Step 0
    Write-Host "  Script 00 manifest : OK" -ForegroundColor Green

    # Gate 2: Script 02 manifest
    $step02Manifest = Read-StepManifest -Step 2
    Write-Host "  Script 02 manifest : OK ($($step02Manifest.Data.MigratedCount) accounts converted)" `
               -ForegroundColor Green

    if ($step02Manifest.Data.FailedCount -gt 0) {
        Write-Host ''
        Write-Host "  WARNING: Script 02 reported $($step02Manifest.Data.FailedCount) failed account(s)." `
                   -ForegroundColor Yellow
        Write-Host '  Folder migration will proceed for successfully converted accounts.' `
                   -ForegroundColor Yellow
        Write-Host '  Failed accounts will be skipped.' -ForegroundColor Yellow
    }

    # Gate 3: folder_map.csv exists
    $folderMapPath = Get-OMMigrateCsvPath -BaseName 'folder_map.csv'
    if (-not (Test-Path $folderMapPath) -and -not $Script:EffectiveSkipFolderMigration) {
        Write-Host ''
        Write-Host '  WARNING: folder_map.csv not found.' -ForegroundColor Yellow
        Write-Host '  Folder migration will be skipped.' -ForegroundColor Yellow
        Write-Host '  Run Script 00 with COM enabled to generate folder_map.csv.' `
                   -ForegroundColor Yellow
        $Script:EffectiveSkipFolderMigration = $true
    }
    else {
        Write-Host "  folder_map.csv     : OK" -ForegroundColor Green
    }

    # Gate 4: Environment
    $envResult = Test-OMMigrateEnvironment
    if (-not $envResult.Passed) {
        $Script:FinalStatus = 'FAILED'
        exit 1
    }
    Write-Host '  Environment check  : OK' -ForegroundColor Green
    Write-Host ''


    # ----------------------------------------------------------
    #  STEP 2 -- Load Data Files and Checkpoint
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Loading Data and Checkpoint Check' -Step '2 of 6'

    # Load account list
    $accountsCsvPath = Get-OMMigrateCsvPath -BaseName 'migration_accounts.csv'
    $allAccounts     = Import-Csv -Path $accountsCsvPath -Encoding UTF8

    # Load folder map
    # Always load even when -RefreshRulesOnly is active -- the rules
    # update needs folder map data to resolve new target folder locations.
    $folderMap = @()
    if (Test-Path $folderMapPath) {
        $folderMap = Import-Csv -Path $folderMapPath -Encoding UTF8
        Write-Host "  Folder map loaded  : $($folderMap.Count) folders" -ForegroundColor Gray
    }

    # -- Build sanitization map from loaded accounts + folders -
    if ($Global:OMMigrate.Sanitize) {
        Initialize-SanitizeMap -Accounts @($allAccounts) -Folders @($folderMap)
        Write-OMMigrateLog -Message '[SANITIZE] Sanitization active -- output masked.' `
                           -Level INFO
    }

    # Load rules inventory -- will be refreshed after COM session starts in Step 4
    $rulesInventoryPath = Get-OMMigrateCsvPath -BaseName 'rules_inventory.csv'
    $rulesInventory = @()
    if (Test-Path $rulesInventoryPath) {
        # Normalize ExecutionOrder before loading -- ensures sequential 1-to-N
        # per account group even if admin added rules without setting the column.
        [void](Invoke-NormalizeRulesExecutionOrder -CsvPath $rulesInventoryPath)
        $rulesInventory = Import-Csv -Path $rulesInventoryPath -Encoding UTF8
        $rulesNeedingUpdate = @($rulesInventory |
            Where-Object { $_.NeedsFolderUpdate -eq 'True' })
        Write-Host "  Rules loaded       : $($rulesInventory.Count) total | $($rulesNeedingUpdate.Count) with folder targets" -ForegroundColor Gray
    }
    else {
        Write-Host '  Rules inventory    : Not found or update skipped' -ForegroundColor DarkGray
    }

    # Determine which accounts need folder work.
    # Source of truth is migration_accounts.csv -- not the Script 02 manifest.
    # The manifest only reflects the last Script 02 run and would miss accounts
    # converted in earlier sessions.
    #
    # IMAP-CONVERTED -- was POP3, went through Scripts 01 and 02, has a backup
    #                   PST, needs Script 03 folder migration. MigrationAction
    #                   should be FOLDER-ONLY after Script 02 completes.
    #
    # IMAP-ALREADY   -- was IMAP before this project started. No backup PST
    #                   exists. Server folders are already correct. Script 03
    #                   creates their Local-destination folder structures in the
    #                   Archive PST (no items copied -- no backup PST to copy from).
    #
    # FOLDER-ONLY    -- MigrationAction used for IMAP-CONVERTED accounts after
    #                   Script 02. Combined with ProviderTag=IMAP-CONVERTED to
    #                   identify accounts that need folder migration.
    #
    # COMPLETE        -- MigrationAction set by Script 03 after successful folder
    #                   migration. Accounts with COMPLETE are excluded from
    #                   eligibility and will not appear in the account picker.
    #                   Re-running Script 03 only processes remaining work.
    $accountsForFolderWork = @($allAccounts | Where-Object {
        ($_.ProviderTag -eq 'IMAP-CONVERTED' -and $_.MigrationAction -eq 'FOLDER-ONLY') -or
        $_.ProviderTag -eq 'IMAP-ALREADY' -or
        # COMPLETE accounts are always included so their Archive structure is maintained
        ($_.ProviderTag -eq 'IMAP-CONVERTED' -and
         $_.MigrationAction -eq 'COMPLETE')
    })

    # Log manifest data for reference only -- not used for eligibility
    $manifestCompletedCount = 0
    if ($step02Manifest.Data.CompletedAccounts) {
        $manifestCompletedCount = @($step02Manifest.Data.CompletedAccounts).Count
    }

    # Count COMPLETE accounts so the operator can see overall progress
    $completedCount = @($allAccounts | Where-Object {
        $_.ProviderTag     -eq 'IMAP-CONVERTED' -and
        $_.MigrationAction -eq 'COMPLETE'
    }).Count

    Write-OMMigrateLog -Message (
        "Script 02 manifest reference: $manifestCompletedCount account(s) in last run. " +
        "Eligibility determined from migration_accounts.csv: $($accountsForFolderWork.Count) account(s) " +
        "with ProviderTag=IMAP-CONVERTED+MigrationAction=FOLDER-ONLY or ProviderTag=IMAP-ALREADY. " +
        "$completedCount account(s) already COMPLETE."
    ) -Level INFO

    Write-Host "  Accounts for folder work: $($accountsForFolderWork.Count)" -ForegroundColor Cyan
    if ($completedCount -gt 0) {
        # When -RecreateRules is active, COMPLETE accounts are eligible and
        # must NOT show as excluded -- they need to appear in the picker.
        if ($RecreateRules) {
            Write-Host "  Already complete        : $completedCount  (included for rules recreation)" `
                       -ForegroundColor Cyan
        }
        else {
            Write-Host "  Already complete        : $completedCount  (excluded from picker)" `
                       -ForegroundColor DarkGray
        }
    }
    Write-Host "  (ProviderTag=IMAP-CONVERTED+FOLDER-ONLY or IMAP-ALREADY in migration_accounts.csv)" `
               -ForegroundColor DarkGray
    if ($RecreateRules) {
        $imapAlreadyCount = @($allAccounts | Where-Object {
            $_.ProviderTag -eq 'IMAP-ALREADY'
        }).Count
        if ($imapAlreadyCount -gt 0) {
            Write-Host "  IMAP-ALREADY accounts   : $imapAlreadyCount  (included for rules recreation)" `
                       -ForegroundColor Cyan
        }
    }

    # Checkpoint check
    $checkpoint     = Read-OMMigrateCheckpoint -Step 3
    $alreadyDone    = @()
    if ($checkpoint.HasCheckpoint) {
        $alreadyDone = $checkpoint.CompletedAccounts
        Write-Host "  Checkpoint: $($alreadyDone.Count) account(s) already have folders migrated." `
                   -ForegroundColor Yellow

        $resumeConfirmed = Confirm-Action `
            -Message  'Resume from checkpoint? (N to re-process all accounts)' `
            -DefaultYes $true
        if (-not $resumeConfirmed) { $alreadyDone = @() }
    }

    $accountsToProcess = @($accountsForFolderWork | Where-Object {
        $alreadyDone -notcontains $_.EmailAddress
    })

    # When -RefreshRulesOnly is active, item migration is skipped but rules
    # still need to be updated. Always include ALL IMAP account types in
    # $accountsToProcess so the rules filter has the correct store names.
    # COMPLETE accounts are valid for rules update regardless of folder status.
    if ($Script:EffectiveSkipFolderMigration) {
        $allImapAccounts = @($allAccounts | Where-Object {
            $_.ProviderTag -eq 'IMAP-CONVERTED' -or
            $_.ProviderTag -eq 'IMAP-ALREADY' -or
            $_.MigrationAction -eq 'COMPLETE'
        })
        # Merge with any already in $accountsToProcess to avoid duplicates
        $existingEmails = @($accountsToProcess | ForEach-Object { $_.EmailAddress })
        $additionalAccounts = @($allImapAccounts | Where-Object {
            $existingEmails -notcontains $_.EmailAddress
        })
        if ($additionalAccounts.Count -gt 0) {
            $accountsToProcess = @($accountsToProcess) + $additionalAccounts
        }
        Write-OMMigrateLog -Message (
            "RefreshRulesOnly: $($accountsToProcess.Count) IMAP account(s) " +
            "(IMAP-CONVERTED + IMAP-ALREADY + COMPLETE) in scope for rules update."
        ) -Level INFO
    }

    # When -RecreateRules is active, ALL IMAP accounts are eligible regardless
    # of MigrationAction -- both IMAP-CONVERTED (POP3 accounts that were migrated)
    # and IMAP-ALREADY (accounts that were always IMAP and have rules in their PST).
    # COMPLETE accounts never had rules created since rules recreation is new.
    # Merge with $accountsToProcess so accounts appear once regardless of status.
    if ($RecreateRules) {
        $allImapAccounts = @($allAccounts | Where-Object {
            $_.ProviderTag -eq 'IMAP-CONVERTED' -or
            $_.ProviderTag -eq 'IMAP-ALREADY' -or
            $_.MigrationAction -eq 'COMPLETE'
        })
        # Add any accounts not already in $accountsToProcess
        $existingEmails = @($accountsToProcess | ForEach-Object { $_.EmailAddress })
        $additionalAccounts = @($allImapAccounts | Where-Object {
            $existingEmails -notcontains $_.EmailAddress
        })
        if ($additionalAccounts.Count -gt 0) {
            $accountsToProcess = @($accountsToProcess) + $additionalAccounts
            Write-OMMigrateLog -Message (
                "RecreateRules: Added $($additionalAccounts.Count) additional account(s) " +
                "(IMAP-CONVERTED + IMAP-ALREADY + COMPLETE) to picker scope. " +
                "Total accounts in scope: $($accountsToProcess.Count)"
            ) -Level INFO
        }
    }

    # ----------------------------------------------------------
    #  ACCOUNT PICKER -- WinForms checkbox list
    #  Lets the operator choose which accounts to process this
    #  run. Supports one, several, or all. Useful when reviewing
    #  folder_map.csv one account at a time before committing.
    #  Skipped automatically when only one account is eligible.
    #  Always shows when multiple accounts are in scope regardless
    #  of -RefreshRulesOnly or -RecreateRules -- the operator
    #  must always be able to target a single account for testing
    #  or partial runs without processing all accounts at once.
    #  Window is fixed-size and centered -- no maximize, no
    #  toolbar overlap.
    # ----------------------------------------------------------
    $showPicker = ($accountsToProcess.Count -gt 1 -and
                   -not $Script:IsWhatIf)

    if ($showPicker) {

        # Load WinForms assembly
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            Add-Type -AssemblyName System.Drawing       -ErrorAction Stop
        }
        catch {
            Write-OMMigrateLog -Message "Could not load WinForms for account picker: $_" -Level INFO
            Write-Host '  Could not open account picker (WinForms unavailable) -- processing all eligible accounts.' `
                       -ForegroundColor Yellow
        }

        # Build display list with folder count per account
        $pickerItems = @($accountsToProcess | ForEach-Object {
            $acct        = $_
            $folderCount = @($folderMap | Where-Object {
                $_.StoreName -like "*$($acct.EmailAddress)*" -or
                $_.StoreName -eq $acct.DisplayName
            }).Count
            [PSCustomObject]@{
                EmailAddress = $acct.EmailAddress
                DisplayName  = $acct.DisplayName
                ProviderTag  = $acct.ProviderTag
                FolderCount  = $folderCount
            }
        } | Sort-Object EmailAddress)

        Write-Host ''
        Write-Host '  A selection window will open -- choose one or more accounts to process.' `
                   -ForegroundColor Cyan
        Write-Host '  Use Select All to process all accounts.' -ForegroundColor Cyan
        Write-Host '  Click Cancel to exit safely without making changes.' -ForegroundColor Cyan
        Write-Host ''

        # -- Build the form ------------------------------------
        $form03                  = New-Object System.Windows.Forms.Form
        $form03.Text             = 'OMMigrate -- Select Accounts for Folder Migration'
        $form03.Size             = New-Object System.Drawing.Size(600, 420)
        $form03.StartPosition    = 'CenterScreen'
        $form03.FormBorderStyle  = 'FixedDialog'
        $form03.MaximizeBox      = $false
        $form03.MinimizeBox      = $false
        $form03.TopMost          = $true

        $label03             = New-Object System.Windows.Forms.Label
        $label03.Location    = New-Object System.Drawing.Point(12, 12)
        $label03.Size        = New-Object System.Drawing.Size(560, 44)
        $label03.Text        = 'Check each account to migrate folders for this run. Unchecked accounts remain eligible and will appear again on the next run.'
        $label03.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
        $form03.Controls.Add($label03)

        $listBox03               = New-Object System.Windows.Forms.CheckedListBox
        $listBox03.Location      = New-Object System.Drawing.Point(12, 64)
        $listBox03.Size          = New-Object System.Drawing.Size(560, 270)
        $listBox03.Font          = New-Object System.Drawing.Font('Segoe UI', 9)
        $listBox03.CheckOnClick  = $true

        foreach ($item in $pickerItems) {
            # WinForms popups always show real data -- sanitize is console-only
            $displayLine = "$($item.EmailAddress)  [$($item.ProviderTag)]  -- $($item.FolderCount) folder(s)"
            $idx = $listBox03.Items.Add($displayLine)
            # All unchecked by default -- operator checks the ones to process
            $listBox03.SetItemChecked($idx, $false)
        }

        $form03.Controls.Add($listBox03)

        $btnSelectAll03          = New-Object System.Windows.Forms.Button
        $btnSelectAll03.Location = New-Object System.Drawing.Point(12, 344)
        $btnSelectAll03.Size     = New-Object System.Drawing.Size(90, 28)
        $btnSelectAll03.Text     = 'Select All'
        $btnSelectAll03.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
        $btnSelectAll03.Add_Click({
            for ($i = 0; $i -lt $listBox03.Items.Count; $i++) {
                $listBox03.SetItemChecked($i, $true)
            }
        })
        $form03.Controls.Add($btnSelectAll03)

        $btnClearAll03           = New-Object System.Windows.Forms.Button
        $btnClearAll03.Location  = New-Object System.Drawing.Point(110, 344)
        $btnClearAll03.Size      = New-Object System.Drawing.Size(90, 28)
        $btnClearAll03.Text      = 'Clear All'
        $btnClearAll03.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
        $btnClearAll03.Add_Click({
            for ($i = 0; $i -lt $listBox03.Items.Count; $i++) {
                $listBox03.SetItemChecked($i, $false)
            }
        })
        $form03.Controls.Add($btnClearAll03)

        $btnOK03                 = New-Object System.Windows.Forms.Button
        $btnOK03.Location        = New-Object System.Drawing.Point(398, 344)
        $btnOK03.Size            = New-Object System.Drawing.Size(80, 28)
        $btnOK03.Text            = 'OK'
        $btnOK03.Font            = New-Object System.Drawing.Font('Segoe UI', 9)
        $btnOK03.DialogResult    = [System.Windows.Forms.DialogResult]::OK
        $form03.AcceptButton     = $btnOK03
        $form03.Controls.Add($btnOK03)

        $btnCancel03             = New-Object System.Windows.Forms.Button
        $btnCancel03.Location    = New-Object System.Drawing.Point(496, 344)
        $btnCancel03.Size        = New-Object System.Drawing.Size(80, 28)
        $btnCancel03.Text        = 'Cancel'
        $btnCancel03.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
        $btnCancel03.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form03.CancelButton     = $btnCancel03
        $form03.Controls.Add($btnCancel03)

        $result03 = $form03.ShowDialog()

        # Capture checked indices BEFORE disposing -- after Dispose() the
        # CheckedListBox controls are destroyed and CheckedIndices returns empty.
        $checkedIndices03 = @($listBox03.CheckedIndices)

        $form03.Dispose()

        if ($result03 -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Host ''
            Write-Host '  No accounts selected -- exiting safely. No changes made.' `
                       -ForegroundColor Yellow
            Write-OMMigrateLog -Message 'Operator cancelled at account picker -- no accounts selected.' `
                               -Level INFO
            exit 0
        }

        if (-not $checkedIndices03 -or $checkedIndices03.Count -eq 0) {
            Write-Host ''
            Write-Host '  No accounts checked -- exiting safely. No changes made.' `
                       -ForegroundColor Yellow
            Write-OMMigrateLog -Message 'Operator clicked OK with no accounts checked -- exiting.' `
                               -Level INFO
            exit 0
        }

        $selectedEmails    = @($checkedIndices03 | ForEach-Object { $pickerItems[$_].EmailAddress })
        $accountsToProcess = @($accountsToProcess | Where-Object {
            $selectedEmails -contains $_.EmailAddress
        })

        Write-Host "  Selected $($accountsToProcess.Count) account(s) to process this run." `
                   -ForegroundColor Green
        Write-OMMigrateLog -Message "Account picker: $($accountsToProcess.Count) selected -- $($selectedEmails -join ', ')" `
                           -Level INFO
    }
    elseif ($accountsToProcess.Count -eq 1) {
        Write-Host "  One account eligible -- processing: $(Invoke-OMMigrateSanitize -Text $($accountsToProcess[0].EmailAddress))" `
                   -ForegroundColor Cyan
    }
    elseif ($Script:IsWhatIf -and $accountsToProcess.Count -gt 1) {
        # In Preview mode show a console summary instead of the grid
        Write-Host ''
        Write-Host '  [PREVIEW] Account picker skipped in WhatIf mode.' -ForegroundColor DarkGray
        Write-Host "  [PREVIEW] Would prompt selection from $($accountsToProcess.Count) eligible accounts." `
                   -ForegroundColor DarkGray
        Write-OMMigrateLog -Message "WhatIf: Would display account picker for $($accountsToProcess.Count) accounts." `
                           -Level INFO -WhatIfPrefix
    }

    # -- Filter rules to selected accounts only -----------------
    # Only update rules whose target folder belongs to a store
    # matching one of the selected accounts. Rules for accounts
    # not selected this run are left untouched.
    if ($rulesNeedingUpdate.Count -gt 0 -and $accountsToProcess.Count -gt 0) {
        $selectedStoreNames = @($accountsToProcess | ForEach-Object {
            # Capture account object before entering inner pipelines --
            # inner Where-Object and ForEach-Object rebind $_ to their
            # own current item, shadowing the outer account object.
            $acct = $_
            @($folderMap | Where-Object {
                $_.StoreName -like "*$($acct.EmailAddress)*" -or
                $_.StoreName -eq $acct.DisplayName
            } | ForEach-Object { $_.StoreName }) + @($acct.EmailAddress) + @($acct.DisplayName)
        } | Select-Object -Unique | Where-Object { $_ })

        $rulesNeedingUpdate = @($rulesNeedingUpdate | Where-Object {
            $rule = $_
            # Always match on RuleStoreName -- it identifies which account owns the rule
            # and is always populated. TargetStoreName is unreliable for filtering:
            #   - Blank for recovered IPM.RuleOrganizer rules
            #   - Set to 'OMMigrate Local Archive' for rules whose target folder is
            #     in the Archive PST -- this does not match any account email address
            #     and would cause those rules to be silently excluded.
            # The rule belongs to the account in RuleStoreName regardless of where
            # the target folder lives.
            $matchName = $rule.RuleStoreName
            if (-not $matchName) { return $false }
            $selectedStoreNames | Where-Object {
                $_ -and (
                    $matchName -like "*$_*" -or
                    $_ -like "*$matchName*"
                )
            }
        })

        Write-OMMigrateLog -Message "Rules filtered to selected accounts: $($rulesNeedingUpdate.Count) rule(s) in scope." `
                           -Level INFO
    }

    Write-Host ''
    Write-Host "  To process this run : $($accountsToProcess.Count)" -ForegroundColor Cyan
    Write-Host "  Already complete    : $($alreadyDone.Count)"       -ForegroundColor DarkGray
    Write-Host ''

    # Stat summary of folder destinations
    if ($folderMap.Count -gt 0) {
        $serverCount = @($folderMap | Where-Object { $_.Destination -eq 'Server' }).Count
        $localCount  = @($folderMap | Where-Object { $_.Destination -eq 'Local'  }).Count
        Write-Host "  Folder destinations:" -ForegroundColor White
        Write-Host "    Server  : $serverCount folders -> IMAP server" -ForegroundColor Cyan
        Write-Host "    Local   : $localCount folders -> Archive PST" -ForegroundColor Gray
        Write-Host ''
    }


    # ----------------------------------------------------------
    #  STEP 3 -- Pre-Flight Confirmation
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Pre-Flight Confirmation' -Step '3 of 6'

    $prereqs = @(
        "Script 02 completed -- $($step02Manifest.Data.MigratedCount) accounts converted to IMAP",
        "Config\folder_map.csv destinations confirmed -- set automatically by Script 01",
        "Outlook is fully closed (this script will launch it via COM)"
    )

    if ($rulesNeedingUpdate.Count -gt 0) {
        $prereqs += "Outlook Rules will be updated -- $($rulesNeedingUpdate.Count) rule(s) have folder targets"
    }

    if ($RecreateRules) {
        $prereqs += "Rules will be recreated from backup PSTs onto IMAP stores (-RecreateRules)"
    }

    Show-PreflightWarning `
        -ScriptDescription (
            "This script will migrate folder structures and email content " +
            "from backup PST files into your new IMAP accounts and local Archive PST. " +
            "It will also update Outlook Rules to point to new folder locations. " +
            "Source backup PSTs are read-only -- your backup data is never modified."
        ) `
        -Prerequisites $prereqs `
        -DeclineMessage @(
            "You chose not to proceed at this time. No folders have been migrated.",
            "Your IMAP accounts from Script 02 are intact and working.",
            "When you are ready to migrate folders, re-run this script:",
            "  .\Scripts\OMMigrate-03-Restore.ps1",
            "Tip: Review Config\folder_map.csv before running to confirm",
            "which folders are set to Server vs Local."
        )

    # Auto-close Outlook if running -- scripts launch their own COM session
    # and require exclusive access. Moved here, AFTER the operator's Y/N
    # confirmation above (Show-PreflightWarning exits the script entirely
    # on a No/decline, so anything below this point only runs after an
    # explicit Yes) -- this avoids closing Outlook out from under the admin
    # if they still have unsaved/in-progress work open (e.g. a rules import
    # not yet confirmed) at the moment the script was launched.
    # TimeoutSeconds 30 (vs the 20s default): this script previously needed a
    # longer wait after force-closing Outlook to avoid the next COM session
    # failing with 'exhausted shared resources' -- kept here as the same
    # tuned value now that the close itself goes through the shared routine.
    [void](Close-OutlookIfRunning -Reason 'before pre-flight' -TimeoutSeconds 30)

    # Initialize progress tracker
    $pendingEmails = @($accountsToProcess | ForEach-Object { $_.EmailAddress })
    Update-OMMigrateProgress -SetPending $pendingEmails


    # ----------------------------------------------------------
    #  STEP 4 -- Launch COM and Set Up Archive PST
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Starting Outlook and Preparing Archive PST' -Step '4 of 6'

    Write-Host '  Launching Outlook COM session...' -ForegroundColor Cyan
    # Read selected profile from Settings.json -- set by Script 00 profile picker.
    # Passed to Connect-OutlookCOM so Outlook opens the correct migration profile.
    $s03ProfileName = ''
    try {
        if ($Global:OMMigrate.Settings -and
            $Global:OMMigrate.Settings.PSObject.Properties['OutlookProfile'] -and
            $Global:OMMigrate.Settings.OutlookProfile.SelectedProfile) {
            $s03ProfileName = $Global:OMMigrate.Settings.OutlookProfile.SelectedProfile
        }
    }
    catch { }
    Write-OMMigrateLog -Message "Using Outlook profile: '$s03ProfileName'" -Level INFO
    $outlook = Connect-OutlookCOM -ProfileName $s03ProfileName
    if (-not $outlook -and -not $Script:IsWhatIf) {
        Write-OMMigrateLog -Message 'Failed to start Outlook COM session.' -Level ERROR
        $Script:FinalStatus = 'FAILED'
        exit 1
    }
    if ($outlook) { $Script:COMSessionOpen = $true }
    Write-Host '  Outlook COM session started.' -ForegroundColor Green

    # Suspend Send/Receive to prevent sync interference during folder migration
    Suspend-OutlookSendReceive | Out-Null

    $namespace = Get-OutlookNamespace

    # -- Scan live Outlook rules and merge into rules_inventory.csv -----------
    # COM session is now active. Runs on every Script 03 execution so new
    # rules added in Outlook since the last Script 00 run are automatically
    # picked up. Export-RulesToCSV merge logic preserves all existing rows
    # and operator edits -- only genuinely new rules are added.
    #
    # CHANGED 2026-06-26 (Administrator): Get-OutlookRules always scans the DEFAULT
    # STORE specifically (all Outlook rules architecturally live there only --
    # see Get-OutlookRules' own header comment). This means it always touches
    # whichever account is the default store (ameritech in this profile),
    # producing that account's known Item() enumeration noise in the log even
    # when the picker selection doesn't include it. Administrator wants this scan
    # skipped entirely unless the default store account is actually part of
    # this run's picker selection -- so picking only a secondary account
    # (e.g. admin@example.com) no longer touches ameritech at all, while
    # selecting the default store account or "Select All" still scans it
    # exactly as before.
    $defaultStoreDisplayName = ''
    try { $defaultStoreDisplayName = $namespace.DefaultStore.DisplayName } catch { }
    $defaultStoreInSelection = $false
    foreach ($acctCheck in $accountsToProcess) {
        if ($acctCheck.EmailAddress -and $defaultStoreDisplayName -and
            ($acctCheck.EmailAddress -eq $defaultStoreDisplayName -or
             $defaultStoreDisplayName -like "*$($acctCheck.EmailAddress)*")) {
            $defaultStoreInSelection = $true
            break
        }
    }

    if ($defaultStoreInSelection) {
        Write-Host ''
        Write-Host '  Scanning live Outlook rules...' -ForegroundColor Cyan
        Write-OMMigrateLog -Message 'Script 03: Scanning live Outlook rules to pick up any new rules.' `
                           -Level INFO

        $liveRulesScanned = @(Get-OutlookRules)
        if ($liveRulesScanned.Count -gt 0) {
            # Added 2026-07-10, Administrator direction: pass the already-in-scope
            # $namespace so Export-RulesToCSV can resolve the real default
            # Archive PST display name for unmapped rows instead of falling
            # back to each row's own RuleStoreName -- see Export-RulesToCSV's
            # own parameter doc in OMMigrate-Outlook.psm1 for full detail.
            $rulesInventoryPath = Export-RulesToCSV -Rules $liveRulesScanned `
                                                    -OutputPath $rulesInventoryPath `
                                                    -Namespace $namespace
            Write-Host "  Rules scanned      : $($liveRulesScanned.Count) rule(s) found -- rules_inventory.csv updated." `
                       -ForegroundColor Green
            Write-OMMigrateLog -Message "Script 03: Live rules scan complete -- $($liveRulesScanned.Count) rule(s) merged into rules_inventory.csv." `
                               -Level INFO

            # Reload rules inventory with any newly discovered rules
            if (Test-Path $rulesInventoryPath) {
                $rulesInventory     = Import-Csv -Path $rulesInventoryPath -Encoding UTF8
                $rulesNeedingUpdate = @($rulesInventory | Where-Object { $_.NeedsFolderUpdate -eq 'True' })
                Write-Host "  Rules reloaded     : $($rulesInventory.Count) total | $($rulesNeedingUpdate.Count) with folder targets" `
                           -ForegroundColor Gray
            }
        }
        else {
            Write-Host '  Rules scanned      : No live rules found -- rules_inventory.csv unchanged.' `
                       -ForegroundColor DarkGray
            Write-OMMigrateLog -Message 'Script 03: No live rules found.' -Level INFO
        }
    }
    else {
        Write-Host ''
        Write-Host '  Live rules scan skipped -- default store account not in this run''s selection.' `
                   -ForegroundColor DarkGray
        Write-OMMigrateLog -Message "Script 03: Live rules scan skipped -- default store '$defaultStoreDisplayName' not in picker selection." -Level INFO
    }

    # -- Build secondary store name set early (used by duplicate check and
    # liveCollections loops to skip GetRules() on secondary stores without
    # zeroing -- prevents UI bleed into ameritech on restart). June 2026.
    $earlyDefaultStoreName    = $namespace.DefaultStore.DisplayName
    $earlySecondaryStoreNames = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($esr in ($rulesNeedingUpdate | Where-Object {
        $_.RuleStoreName -ne $earlyDefaultStoreName
    } | Select-Object -ExpandProperty RuleStoreName -Unique)) {
        # FIXED (Administrator direction, 2026-07-20): .Trim() added -- matches the
        # established pattern Invoke-DeployConsolidatedRules already uses for this
        # exact kind of RuleStoreName/EmailAddress comparison (OMMigrate-Outlook.psm1,
        # $scopedAccountLookup). Untrimmed values here could cause a legitimate
        # secondary-store account to not match $earlySecondaryStoreNames.Contains()
        # further below (Administrator hit a real case: a WARN fired for a run where
        # every picker-selected account genuinely was a secondary store already
        # fully handled by consolidation -- a whitespace difference between this
        # CSV's RuleStoreName and migration_accounts.csv's EmailAddress for the
        # same account is the most likely explanation, consistent with every other
        # comparison of this kind elsewhere in the codebase already trimming).
        [void]$earlySecondaryStoreNames.Add($esr.Trim())
    }
    Write-OMMigrateLog -Message "Secondary stores with CSV rules: $($earlySecondaryStoreNames.Count) -- GetRules() will be skipped on these until secondary block runs." -Level INFO

    # -- Remove duplicate rules from Outlook before remapping ----------------
    # Duplicate rules arise from .rwz imports adding a second copy alongside
    # originals. Removing them here ensures Script 03 remaps exactly one copy
    # per rule name. Uses the same $namespace already open above.
    # Groups rules by base name (strips trailing (n) suffix) and deletes all
    # but the best-scoring copy. One Save() per store -- atomic per Gemini pattern.
    Write-Host ''
    Write-Host '  Checking for duplicate rules in Outlook...' -ForegroundColor Cyan
    Write-OMMigrateLog -Message 'Script 03: Scanning for duplicate Outlook rules.' -Level INFO

    try {
        $dupStores = $namespace.Stores
        Register-COMObject -ComObject $dupStores

        for ($ds = 1; $ds -le $dupStores.Count; $ds++) {
            try {
                $dupStore = $dupStores.Item($ds)
                Register-COMObject -ComObject $dupStore
                $dupStoreName = ''
                try { $dupStoreName = $dupStore.DisplayName } catch { }

                $dupRules = $null
                # CHANGED (2026-07-08, Administrator explicit direction): same correction as
                # the liveCollections loop further below in this file -- the skip
                # that used to be here was based on a false June 2026 assumption
                # conflating this with the unrelated devil-code investigation.
                # Secondary stores now get the same plain GetRules() duplicate-check
                # pass as the default store, no special-casing.
                try { $dupRules = $dupStore.GetRules() } catch { continue }

                # Build a map of baseName -> list of (index, rule, score)
                # Base name = rule name with trailing (n) stripped
                $ruleMap = @{}
                for ($dr = 1; $dr -le $dupRules.Count; $dr++) {
                    try {
                        $dr_rule  = $dupRules.Item($dr)
                        $dr_name  = $dr_rule.Name
                        $baseName = $dr_name -replace '\s*\(\d+\)$', ''

                        # Score: valid MoveToFolder target = 8, has actions = 4,
                        #        Enabled = 1. Highest score is kept.
                        $dr_score = 0
                        try {
                            $mf = $dr_rule.Actions.MoveToFolder
                            if ($mf.Enabled -and $null -ne $mf.Folder) { $dr_score += 8 }
                            if ($mf.Enabled)                           { $dr_score += 4 }
                        } catch { }
                        try { if ($dr_rule.Enabled) { $dr_score += 1 } } catch { }

                        if (-not $ruleMap.ContainsKey($baseName)) {
                            $ruleMap[$baseName] = [System.Collections.Generic.List[PSCustomObject]]::new()
                        }
                        $ruleMap[$baseName].Add([PSCustomObject]@{
                            Index = $dr
                            Name  = $dr_name
                            Score = $dr_score
                        })
                    } catch { }
                }

                # For each base name with more than one entry, delete all but highest scorer
                $dupNamesToDelete = [System.Collections.Generic.List[string]]::new()
                foreach ($baseName in $ruleMap.Keys) {
                    $entries = @($ruleMap[$baseName] | Sort-Object Score -Descending)
                    if ($entries.Count -gt 1) {
                        # Keep entries[0] (highest score), delete the rest
                        for ($di = 1; $di -lt $entries.Count; $di++) {
                            $dupNamesToDelete.Add($entries[$di].Name)
                        }
                    }
                }

                if ($dupNamesToDelete.Count -gt 0) {
                    $dupDeleted = 0
                    foreach ($delName in $dupNamesToDelete) {
                        try {
                            $dupRules.Remove($delName)
                            $dupDeleted++
                            Write-OMMigrateLog -Message "Duplicate rule removed: '$delName' from '$dupStoreName'" `
                                               -Level INFO
                        } catch {
                            Write-OMMigrateLog -Message "Failed to remove duplicate rule '$delName': $_" `
                                               -Level WARN
                        }
                    }

                    if ($dupDeleted -gt 0) {
                        try {
                            # Corrected 2026-07-01 (Administrator): was @($true) (ShowProgress on),
                            # inconsistent with every other Save() call site in this project,
                            # all of which use @($false) for silent, unattended saves -- no
                            # documented rationale found in project history for why this one
                            # call used $true. Corrected to @($false) to match. Also added
                            # [void] to suppress the InvokeMember return value, matching every
                            # other Save() call site -- an unsuppressed statement here would
                            # become part of this function's implicit output, same class of
                            # bug already found and fixed once today in a different function.
                            [void]$dupRules.GetType().InvokeMember(
                                'Save',
                                [System.Reflection.BindingFlags]::InvokeMethod,
                                $null, $dupRules, @($false)
                            )
                            Write-Host "  Duplicates removed  : $dupDeleted rule(s) deleted from '$dupStoreName'" `
                                       -ForegroundColor Yellow
                            Write-OMMigrateLog -Message "Duplicate rules removed from '$dupStoreName': $dupDeleted" `
                                               -Level INFO
                        } catch {
                            # Save() failure here is expected on first run -- rules still
                            # have stale folder targets that haven't been remapped yet.
                            # The remap phase below will fix them. Non-fatal, log as INFO.
                            Write-OMMigrateLog -Message "Duplicate rule Save() skipped -- rules have stale targets (expected before remap): $_" `
                                               -Level INFO
                        }
                    }
                }

                # Garbage cleanup fix, corrected (2026-07-01, Administrator): the original
                # cleanup for $dupRules was placed AFTER this entire for-loop over
                # $dupStores, checking $dupRules only once. Since $dupRules is
                # reassigned to $null at the TOP of every loop iteration (see the
                # "$dupRules = $null" line above, per store), whatever the LAST
                # store in the loop left behind is all that single post-loop check
                # ever saw -- confirmed via diagnostic logging: dupStores.Count=3,
                # but dupRules was $null at the post-loop check, because a later
                # store in the loop hit a skip/continue path without ever
                # reassigning $dupRules to a real object again. This silently
                # meant EVERY store's $dupRules reference (including ones that
                # were genuinely used) was left un-released for the rest of the
                # script's execution -- the exact leaked-reference condition this
                # cleanup was meant to prevent. Moved inside the loop, at the end
                # of each store's own try block, so every store's own $dupRules
                # reference is released immediately after that store's own work
                # is done, regardless of what later iterations do to the variable.
                try {
                    if ($dupRules) {
                        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($dupRules)
                        $dupRules = $null
                        Write-OMMigrateLog -Message "Garbage cleanup: released `$dupRules COM reference for store '$dupStoreName' after duplicate rule check." -Level DEBUG
                    }
                }
                catch {
                    Write-OMMigrateLog -Message "Garbage cleanup: failed to release `$dupRules for store '$dupStoreName': $_" -Level DEBUG
                }
            } catch { }
        }

        Write-Host '  Duplicate rule check complete.' -ForegroundColor Gray
        Write-OMMigrateLog -Message 'Script 03: Duplicate rule check complete.' -Level INFO
    }
    catch {
        Write-OMMigrateLog -Message "Script 03: Duplicate rule check failed (non-fatal): $_" `
                           -Level INFO
    }

    # Garbage cleanup (added 2026-07-01, Administrator): $dupRules above was never
    # released -- a live, un-registered COM reference to the default store's
    # Rules collection held in memory for the rest of this script's execution,
    # well past the point where the consolidation phase deletes and saves
    # against a completely separate Rules collection reference for the same
    # store. Per Gemini's confirmed working-script pattern, an earlier-phase
    # reference kept alive into a later phase can cause a later GetRules()
    # call to serve stale cached data, or cause the old reference itself to
    # flush its own stale snapshot back to disk when finally released/GC'd,
    # silently overwriting later, correct changes. $dupRules' own work
    # (duplicate detection and removal) is fully complete by this point --
    # safe to release here, well before the consolidation phase runs.
    try {
        if ($dupRules) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($dupRules)
            $dupRules = $null
            Write-OMMigrateLog -Message "Garbage cleanup: released `$dupRules COM reference after duplicate rule check." -Level DEBUG
        }
    }
    catch {
        Write-OMMigrateLog -Message "Garbage cleanup: failed to release `$dupRules: $_" -Level DEBUG
    }

    # -- Remove misrouted secondary account rules from default store -----------
    # Rules belonging to secondary accounts (gmail, a custom domain, etc.) that were
    # incorrectly created into the default store's collection by a prior buggy run of the
    # secondary store block. These rules have Conditions.Account set to a secondary
    # account SmtpAddress/DisplayName -- they do not belong in the default store's
    # rule list and must be removed before the secondary store block recreates them
    # correctly. Safe to run every time -- only removes rules whose account condition
    # does not match the default store account.
    try {
        $defaultAcctSmtp = ''
        $defaultAcctDisp = ''
        try {
            $defAccts = $namespace.Session.Accounts
            for ($dai = 1; $dai -le $defAccts.Count; $dai++) {
                try {
                    $defAcct = $defAccts.Item($dai)
                    $defSmtp = ''; try { $defSmtp = $defAcct.SmtpAddress } catch { }
                    $defDisp = ''; try { $defDisp = $defAcct.DisplayName  } catch { }
                    if ($defSmtp -like "*$($namespace.DefaultStore.DisplayName)*" -or
                        $namespace.DefaultStore.DisplayName -like "*$defSmtp*") {
                        $defaultAcctSmtp = $defSmtp
                        $defaultAcctDisp = $defDisp
                        break
                    }
                } catch { }
            }
        } catch { }

        if (-not [string]::IsNullOrWhiteSpace($defaultAcctSmtp)) {
            $misroutedRules  = $namespace.DefaultStore.GetRules()
            $misroutedNames  = [System.Collections.Generic.List[string]]::new()
            for ($mri = 1; $mri -le $misroutedRules.Count; $mri++) {
                try {
                    $mr     = $misroutedRules.Item($mri)
                    $mrName = ''; try { $mrName = $mr.Name } catch { }
                    $mrAcctSmtp = ''
                    $mrAcctDisp = ''
                    try {
                        $mrAcctCond = $mr.Conditions.Account
                        if ($mrAcctCond.Enabled) {
                            $mrAcct = $mrAcctCond.Account
                            try { $mrAcctSmtp = $mrAcct.SmtpAddress } catch { }
                            try { $mrAcctDisp = $mrAcct.DisplayName  } catch { }
                        }
                    } catch { }
                    # If account condition is set and points to a different account
                    # than the default store -- this rule was misrouted.
                    if (-not [string]::IsNullOrWhiteSpace($mrAcctSmtp) -and
                        $mrAcctSmtp -ne $defaultAcctSmtp) {
                        [void]$misroutedNames.Add($mrName)
                        Write-OMMigrateLog -Message (
                            "Misrouted rule detected: '$mrName' has Account='$mrAcctSmtp' " +
                            "but lives in default store '$defaultAcctSmtp' -- queued for removal."
                        ) -Level INFO
                    }
                } catch { }
            }

            if ($misroutedNames.Count -gt 0) {
                $misroutedDeleted = 0
                foreach ($mrDel in $misroutedNames) {
                    try {
                        [void]$misroutedRules.GetType().InvokeMember(
                            'Remove',
                            [System.Reflection.BindingFlags]::InvokeMethod,
                            $null, $misroutedRules, @($mrDel)
                        )
                        $misroutedDeleted++
                        Write-OMMigrateLog -Message "Misrouted rule removed: '$mrDel'" -Level INFO
                    } catch {
                        Write-OMMigrateLog -Message "Failed to remove misrouted rule '$mrDel': $_" -Level WARN
                    }
                }
                if ($misroutedDeleted -gt 0) {
                    try {
                        [void]$misroutedRules.GetType().InvokeMember(
                            'Save',
                            [System.Reflection.BindingFlags]::InvokeMethod,
                            $null, $misroutedRules, @($false)
                        )
                        Write-Host "  Misrouted rules     : $misroutedDeleted rule(s) removed from default store." `
                                   -ForegroundColor Yellow
                        Write-OMMigrateLog -Message "Misrouted rules removed from default store: $misroutedDeleted" `
                                           -Level INFO
                    } catch {
                        Write-OMMigrateLog -Message "Misrouted rule Save() failed: $_" -Level WARN
                    }
                }
            } else {
                Write-OMMigrateLog -Message 'Misrouted rule check: no misrouted rules found.' -Level INFO
            }
        } else {
            Write-OMMigrateLog -Message 'Misrouted rule check: could not resolve default account SmtpAddress -- skipped.' -Level WARN
        }
    } catch {
        Write-OMMigrateLog -Message "Misrouted rule check failed (non-fatal): $_" -Level INFO
    }

    # -- Strip Account condition from default-store rules with a pending row --
    # Added 2026-07-03, Administrator's explicit, permanent requirement: the Account
    # condition ("through the specified account" filter) is confirmed as the
    # root cause of the 0x800C8101 "devil code" COM read failures on Outlook
    # 2021 Classic -- even a legitimately-resolved Account condition (one that
    # correctly points at the default store's own account, not a misrouted
    # one) triggers it. Set-RuleConditions (OMMigrate-Outlook.psm1) no longer
    # sets this condition on newly-created rules as of today, but that fix
    # only prevents FUTURE occurrences -- it does nothing for rules that
    # already carry the condition from an earlier .rwz import or a Script 03
    # run predating the fix. This block is the cleanup pass: for any rule in
    # the default store with an ENABLED Account condition -- regardless of
    # which account it points to, including the default store's own account
    # -- disables that one condition via InvokeMember (mirrors the proven-safe
    # From-condition disable pattern in Set-RuleConditions, Step 1: set
    # Enabled=$false only, leave the rest of the rule/condition object
    # untouched). This is distinct from and runs independently of the
    # misrouted-rule detector above, which DELETES rules pointing at a
    # different account -- Administrator was explicit that rules must never be deleted
    # for this reason, only have the Account condition stripped. Administrator
    # confirmed via live testing that rules work correctly on both the
    # primary and secondary accounts, and correctly move mail to the target
    # folder, without any Account condition at all -- it was only ever
    # required for legacy POP3-era rule scoping and is not needed for this
    # project's IMAP-based migration.
    #
    # Scope (2026-07-03, Administrator's explicit correction): originally scanned
    # every rule in the default store on every run, unconditionally. Administrator
    # pointed out this is unnecessary re-work -- LastDeployedRun in
    # rules_inventory.csv is already the established, standard signal for
    # "this row needs (re)checking" everywhere else in the pipeline (a blank
    # date on first run, an admin blanking specific rows, or a full .rwz
    # reimport where the admin blanks every row). Once stripped, a rule
    # stays stripped (no code re-adds the condition), so there is no need to
    # re-check rules whose row already has a LastDeployedRun timestamp. Gated
    # to only strip rules whose CSV row has a BLANK LastDeployedRun, matching
    # the same pending-row convention used throughout Invoke-DeployConsolidatedRules.
    try {
        $acctStripPendingNames = [System.Collections.Generic.HashSet[string]]::new()
        try {
            $acctStripCsvRows = Import-Csv -Path $rulesInventoryPath -Encoding UTF8
            foreach ($ascRow in $acctStripCsvRows) {
                $ascTimestamp = if ($ascRow.PSObject.Properties['LastDeployedRun']) { [string]$ascRow.LastDeployedRun } else { '' }
                $ascRuleName  = if ($ascRow.PSObject.Properties['RuleName'])        { [string]$ascRow.RuleName }        else { '' }
                if ([string]::IsNullOrWhiteSpace($ascTimestamp) -and -not [string]::IsNullOrWhiteSpace($ascRuleName)) {
                    [void]$acctStripPendingNames.Add($ascRuleName)
                }
            }
        } catch {
            Write-OMMigrateLog -Message "Account condition strip pass: could not read rules_inventory.csv to scope pending rows -- skipping strip pass entirely (non-fatal): $_" -Level WARN
        }

        $acctStripRules   = $namespace.DefaultStore.GetRules()
        $acctStripCount   = 0
        $acctStripChecked = 0
        for ($asi = 1; $asi -le $acctStripRules.Count; $asi++) {
            try {
                $asRule = $acctStripRules.Item($asi)
                $asRuleNameCheck = ''
                try { $asRuleNameCheck = $asRule.Name } catch { }
                if (-not $acctStripPendingNames.Contains($asRuleNameCheck)) { continue }
                $acctStripChecked++
                $asAcctCond = $null
                try { $asAcctCond = $asRule.Conditions.Account } catch { }
                if ($asAcctCond -and $asAcctCond.Enabled) {
                    $asRuleName = ''
                    try { $asRuleName = $asRule.Name } catch { }
                    # Fix (2026-07-03, Administrator): disabling alone (Enabled=$false) was
                    # judged insufficient -- Administrator wants every possible trigger
                    # surface removed, not just deactivated, since a disabled
                    # condition still holds a live COM reference to an Account
                    # object. Clear the Account reference itself (Account=$null)
                    # in addition to disabling, so nothing is left for a future
                    # code path or manual UI edit to accidentally re-enable
                    # against a stale reference.
                    try {
                        [void]$asAcctCond.GetType().InvokeMember(
                            'Account',
                            [System.Reflection.BindingFlags]::SetProperty,
                            $null, $asAcctCond, @($null)
                        )
                    } catch {
                        Write-OMMigrateLog -Message "Could not clear Account reference on '$asRuleName' (non-fatal, will still disable): $_" -Level DEBUG
                    }
                    [void]$asAcctCond.GetType().InvokeMember(
                        'Enabled',
                        [System.Reflection.BindingFlags]::SetProperty,
                        $null, $asAcctCond, @($false)
                    )
                    $acctStripCount++
                    Write-OMMigrateLog -Message "Account condition stripped: '$asRuleName' (devil-code prevention -- not needed for IMAP)." -Level INFO
                }
            } catch {
                Write-OMMigrateLog -Message "Failed to check/strip Account condition on rule index $asi (non-fatal): $_" -Level DEBUG
            }
        }
        if ($acctStripCount -gt 0) {
            try {
                [void]$acctStripRules.GetType().InvokeMember(
                    'Save',
                    [System.Reflection.BindingFlags]::InvokeMethod,
                    $null, $acctStripRules, @($false)
                )
                Write-Host "  Account conditions stripped : $acctStripCount rule(s) (of $acctStripChecked checked)." `
                           -ForegroundColor Yellow
                Write-OMMigrateLog -Message "Account condition strip pass complete: $acctStripCount rule(s) stripped, $acctStripChecked checked." -Level INFO
            } catch {
                Write-OMMigrateLog -Message "Account condition strip Save() failed: $_" -Level WARN
            }
        } else {
            Write-OMMigrateLog -Message "Account condition strip pass: no rules had an enabled Account condition ($acctStripChecked checked)." -Level INFO
        }
    } catch {
        Write-OMMigrateLog -Message "Account condition strip pass failed (non-fatal): $_" -Level INFO
    } finally {
        if ($acctStripRules) {
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($acctStripRules) } catch { }
            $acctStripRules = $null
        }
    }

    # Garbage cleanup (added 2026-07-01, Administrator): $misroutedRules above was
    # never released either -- same leaked-reference concern as $dupRules
    # just above. Its own work (misrouted rule detection and removal) is
    # fully complete by this point -- safe to release here, well before the
    # consolidation phase runs and fetches its own separate Rules collection
    # reference for the same default store.
    try {
        if ($misroutedRules) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($misroutedRules)
            $misroutedRules = $null
            Write-OMMigrateLog -Message "Garbage cleanup: released `$misroutedRules COM reference after misrouted rule check." -Level DEBUG
        }
    }
    catch {
        Write-OMMigrateLog -Message "Garbage cleanup: failed to release `$misroutedRules: $_" -Level DEBUG
    }

    # Set up the Archive PST -- always opened regardless of -RefreshRulesOnly.
    # The Rules Update phase needs $archiveRootFolder to navigate Local-destination
    # rule targets even when folder migration itself is skipped.
    $archiveRootFolder = $null

    $Script:ArchivePSTPath = Join-Path $Global:OMMigrate.BackupPath $ArchivePSTName

    # Added 2026-07-10, Administrator direction (freeze lifted specifically for this
    # fix). BUG: this Archive PST open was never paired with a matching
    # Close-PSTFile anywhere in this script -- confirmed via full-file grep,
    # unlike every backup PST open in this same file, which already checks
    # Test-PSTAlreadyMounted first and only detaches what it itself mounted
    # (see the $backupPathWasAlreadyMounted pattern used elsewhere in this
    # file). Release-OutlookCOM's Quit() does NOT undo an AddStore -- a
    # store mounted via AddStore stays registered in the Outlook PROFILE
    # itself (confirmed via Release-OutlookCOM's own implementation, which
    # only quits the Application and releases CLR COM references, never
    # calls RemoveStore) -- so the Archive PST was staying attached
    # indefinitely across runs, exactly as Administrator observed. Fixed by applying
    # the identical, already-proven pattern used for every backup PST in
    # this file: check mounted state before opening, only detach at
    # teardown if this run's own code did the mounting.
    $archivePSTWasAlreadyMounted = Test-PSTAlreadyMounted -PSTPath $Script:ArchivePSTPath

    # FIXED 2026-07-10, Administrator direction. Same MasterArchiveNames protection
    # added to Script 01's Archive pre-build (see that file's own fix
    # comment for full context) -- some archives (e.g. 'OMMigrate Local
    # Archive', a personal auto-archive store) must never be auto-detached
    # regardless of who mounted them this run. Operator-configured via
    # RulesEngine.MasterArchiveNames in Settings.json rather than hardcoded,
    # matched by exact DisplayName. Checked here (not just in Script 01)
    # since this script has its own separate archive open/close pair for
    # the default archive. Empty list (default) means this check never
    # protects anything extra beyond the existing already-mounted guard.
    $archiveIsMasterArchive = $false
    try {
        if ($Global:OMMigrate.Settings.PSObject.Properties['RulesEngine'] -and
            $Global:OMMigrate.Settings.RulesEngine.PSObject.Properties['MasterArchiveNames']) {
            $masterArchiveNamesForCheck = @($Global:OMMigrate.Settings.RulesEngine.MasterArchiveNames)
            if ($masterArchiveNamesForCheck -contains $ArchivePSTName) {
                $archiveIsMasterArchive = $true
            }
        }
    }
    catch { }

    # Archive PST is pre-created by Install.ps1 using Outlook COM (New-ArchivePSTViaCOM).
    # Outlook builds a fully valid PST structure during install, so AddStore here
    # opens a clean file with no risk of COM session destabilization.
    # If the file is missing, direct the operator to re-run Install.ps1
    # which will regenerate it via COM.
    if (Test-Path $Script:ArchivePSTPath) {
        Write-Host "  Archive PST already exists: $(Invoke-OMMigrateSanitize -Text $Script:ArchivePSTPath)" `
                   -ForegroundColor Green
        $archiveStore = Open-PSTFile `
            -PSTPath     $Script:ArchivePSTPath `
            -DisplayName 'OMMigrate Local Archive'
    }
    else {
        Write-Host '  WARNING: Archive PST not found.' -ForegroundColor Yellow
        Write-Host '  Re-run Install.ps1 to regenerate it, then re-run this script.' `
                   -ForegroundColor Yellow
        Write-Host '  Local-destination folders and rule targets will be skipped this run.' `
                   -ForegroundColor Yellow
        Write-OMMigrateLog -Message "Archive PST not found: $Script:ArchivePSTPath -- re-run Install.ps1" `
                           -Level WARN
    }

    # Get archive root folder
    try {
        if ($archiveStore) {
            $archiveRootFolder = $archiveStore.GetRootFolder()
            Register-COMObject -ComObject $archiveRootFolder
            Write-Host "  Archive PST ready: $($archiveRootFolder.Name)" -ForegroundColor Green
        }
    }
    catch {
        Write-OMMigrateLog -Message "Could not get Archive PST root folder: $_" -Level WARN
    }

    # Added 2026-07-10, Administrator (multi-archive support, TargetStoreName hardcode
    # fix, -RecreateRules scope). $archiveRootFolder above is a single,
    # opened-once archive PST (the one this account's own pre-build pass
    # attached). But an individual rule's rules_inventory.csv row may carry
    # its own TargetStoreName pointing at a DIFFERENT attached archive PST
    # (e.g. that rule's target account was mapped to a second archive via the
    # TargetStoreName picker in Script 00). This cache holds resolved root
    # folders for any such distinct TargetStoreName encountered by the
    # Strategy 1+2 folder-target remap loop further below, keyed by store
    # DisplayName, resolved once per distinct name via the same live
    # Namespace.Stores match + GetRootFolder() pattern already proven in
    # Invoke-DeployConsolidatedRules (OMMigrate-Outlook.psm1) -- not a new
    # mechanism. Script-level scope so it is built once and reused across
    # every account processed by this run, not rebuilt per account.
    $script:resolvedArchiveStoreFoldersForRemap = @{}

    Write-Host ''


    # ----------------------------------------------------------
    #  STEP 5 -- Migrate Folders Per Account
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Migrating Folders and Email Content' -Step '5 of 6'

    if ($Script:EffectiveSkipFolderMigration) {
        Write-Host '  Item migration skipped (-RefreshRulesOnly).' -ForegroundColor Yellow
        Write-OMMigrateLog -Message 'Folder migration skipped by parameter.' -Level INFO
    }
    else {
        foreach ($account in $accountsToProcess) {
            $email = $account.EmailAddress

            Write-Host ''
            Write-Host ('  ' + ('=' * 54)) -ForegroundColor DarkCyan
            Write-Host "  Account: $(Invoke-OMMigrateSanitize -Text $email)" -ForegroundColor White
            Write-Host "  Tag    : $($account.ProviderTag)" -ForegroundColor DarkGray
            Write-Host ''

            # Get folder map rows for this account's store.
            # Filter out blank separator rows (StoreName is empty) and
            # deduplicate by FolderPath -- Script 00 may discover the same
            # folder from both the PST store and the IMAP OST store, resulting
            # in duplicate entries that would create duplicate folders on the server.
            $accountFolderMap = @($folderMap | Where-Object {
                $_.StoreName -and (
                    $_.StoreName -like "*$email*" -or
                    $_.StoreName -eq $account.DisplayName
                )
            } | Sort-Object FolderPath -Unique)

            if ($accountFolderMap.Count -eq 0) {
                Write-Host '    No folder map entries found for this account.' `
                           -ForegroundColor DarkGray
                Write-OMMigrateLog -Message "No folder map entries for $email -- skipping." `
                                   -Level INFO

                $skipResult = $account.PSObject.Copy()
                $skipResult | Add-Member -NotePropertyName 'MigrationOutcome' `
                                         -NotePropertyValue 'SKIPPED' -Force
                $skipResult | Add-Member -NotePropertyName 'MigrationDetail' `
                                         -NotePropertyValue 'No folder map entries found' -Force
                [void]$Script:AccountResults.Add($skipResult)
                Update-OMMigrateProgress -MarkComplete $email
                continue
            }

            Show-AccountStatus -Email  $email `
                               -Tag    $account.ProviderTag `
                               -Action "Migrating $($accountFolderMap.Count) folders..." `
                               -Status 'INFO'

            Update-OMMigrateProgress -SetCurrent $email

            # Per-account Y/N prompt
            $proceed = $true
            if (-not $Force) {
                $proceed = Confirm-Action `
                    -Message      "Migrate $($accountFolderMap.Count) folders for: $email ?" `
                    -AccountEmail $email `
                    -DefaultYes   $true
            }

            if (-not $proceed) {
                Write-OMMigrateLog -Message "Folder migration skipped by operator: $email" `
                                   -Level INFO
                $skipResult = $account.PSObject.Copy()
                $skipResult | Add-Member -NotePropertyName 'MigrationOutcome' `
                                         -NotePropertyValue 'SKIPPED' -Force
                $skipResult | Add-Member -NotePropertyName 'MigrationDetail' `
                                         -NotePropertyValue 'Skipped by operator' -Force
                [void]$Script:AccountResults.Add($skipResult)
                Update-OMMigrateProgress -MarkComplete $email
                continue
            }

            # Execute folder migration for this account
            $migrationResult = Invoke-AccountFolderMigration `
                -Account           $account `
                -FolderMap         $accountFolderMap `
                -Namespace         $namespace `
                -ArchiveRootFolder $archiveRootFolder `
                -RulesInventory    $rulesInventory

            $Script:FoldersCreated  += $migrationResult.FoldersCreated
            $Script:FoldersVerified += $migrationResult.FoldersVerified
            $Script:ItemsCopied    += $migrationResult.ItemsCopied
            $Script:FoldersFailed  += $migrationResult.FoldersFailed

            # Attach result to account for report
            $accountResult = $account.PSObject.Copy()
            $accountResult | Add-Member -NotePropertyName 'MigrationOutcome' `
                                        -NotePropertyValue $migrationResult.Outcome -Force
            $accountResult | Add-Member -NotePropertyName 'MigrationDetail' `
                                        -NotePropertyValue $migrationResult.Detail -Force
            [void]$Script:AccountResults.Add($accountResult)

            $statusIcon = switch ($migrationResult.Outcome) {
                'SUCCESS' { 'OK'   }
                'WARNING' { 'WARN' }
                'SKIPPED' { 'SKIP' }
                default   { 'FAIL' }
            }

            Show-AccountStatus `
                -Email  $email `
                -Tag    $account.ProviderTag `
                -Action "$($migrationResult.Detail)" `
                -Status $statusIcon

            Write-AuditEntry -Action 'FOLDERS_MIGRATED' `
                             -AccountEmail $email `
                             -Detail $migrationResult.Detail `
                             -Outcome $migrationResult.Outcome

            Save-OMMigrateCheckpoint -ExitType 'NORMAL' -Reason 'Account folder migration complete'
            Update-OMMigrateProgress -MarkComplete $email

            # Update migration_accounts.csv -- mark account COMPLETE on SUCCESS or WARNING.
            # COMPLETE accounts are excluded from eligibility on future Script 03 runs.
            # FAILED accounts are left as FOLDER-ONLY so they are retried automatically.
            # Operator-skipped accounts are also left as FOLDER-ONLY to re-appear next run.
            if ($migrationResult.Outcome -in @('SUCCESS', 'WARNING')) {
                Update-AccountMigrationAction -EmailAddress $email -NewAction 'COMPLETE'
            }

            if ($migrationResult.Outcome -eq 'FAILED' -and
                $Script:FinalStatus -eq 'SUCCESS') {
                $Script:FinalStatus = 'WARNING'
            }
        }

        Write-Host ''
        Write-Host '  Folder Migration Summary:' -ForegroundColor White
        Write-Host "    Folders created  : $Script:FoldersCreated" -ForegroundColor Gray
        Write-Host "    Folders verified : $Script:FoldersVerified" -ForegroundColor Gray
        Write-Host "    Items copied    : $Script:ItemsCopied" -ForegroundColor Gray
        Write-Host "    Folders failed  : $Script:FoldersFailed" -ForegroundColor $(
            if ($Script:FoldersFailed -gt 0) { 'Yellow' } else { 'Gray' }
        )
        Write-Host ''
    }

    # -- Rules Recreation Phase --------------------------------
    # Reads rules from each account's backup PST and recreates them
    # on the new IMAP store. Folder target pointers are remapped per
    # folder_map.csv. All other rule properties are preserved verbatim.
    # Safe to re-run -- rules that already exist by name are skipped.
    if ($RecreateRules) {
        Show-SectionHeader -Title 'Recreating Rules from Backup PSTs'

        $accountsForRules = @($accountsToProcess)

        # When -RefreshRulesOnly was active $accountsToProcess may include
        # COMPLETE accounts. Include them -- rules recreation is independent
        # of folder migration status. Also include IMAP-ALREADY accounts which
        # were always IMAP and may have rules in their backup PSTs.
        if ($accountsForRules.Count -eq 0) {
            $accountsForRules = @($allAccounts | Where-Object {
                $_.ProviderTag -eq 'IMAP-CONVERTED' -or
                $_.ProviderTag -eq 'IMAP-ALREADY' -or
                $_.MigrationAction -eq 'COMPLETE'
            })
        }

        if ($accountsForRules.Count -eq 0) {
            Write-Host '  No eligible IMAP accounts found for rules recreation.' `
                       -ForegroundColor Yellow
            Write-OMMigrateLog -Message 'RecreateRules: No eligible accounts found.' -Level WARN
        }
        else {
            Write-Host "  Scanning $($accountsForRules.Count) account(s) for rules in backup PSTs..." `
                       -ForegroundColor Cyan
            Write-Host ''

            foreach ($account in $accountsForRules) {
                $email     = $account.EmailAddress
                $safeEmail = Get-SafeFileName -InputString $email
                # REVERTED 2026-07-11, Administrator direction. Same as
                # Invoke-AccountFolderMigration's backupPath above -- the
                # plain POP3 backup PST is a one-time, profile-independent
                # artifact and must never be profile-suffixed. See that
                # function's fix comment for full context.
                $backupPath = Join-Path $Global:OMMigrate.BackupPath "$safeEmail.pst"

                Write-Host ('  ' + ('-' * 54)) -ForegroundColor DarkCyan
                Write-Host "  Account: $(Invoke-OMMigrateSanitize -Text $email)" `
                           -ForegroundColor White

                if (-not (Test-Path $backupPath)) {
                    Write-Host "    [SKIP] Backup PST not found: $safeEmail.pst" `
                               -ForegroundColor DarkGray
                    Write-OMMigrateLog -Message "RecreateRules: Backup PST not found for $email -- skipping." `
                                       -Level WARN
                    continue
                }

                # Read rules from backup PST
                Write-Host '    Reading rules from backup PST...' -ForegroundColor DarkGray

                # Check first whether this PST is already mounted in the profile
                # (e.g. Administrator manually attached it for permanent manual triage use)
                # so it is not detached later as if it were this script's own
                # temporary mount.
                $backupPathWasAlreadyMounted = Test-PSTAlreadyMounted -PSTPath $backupPath

                $sourceRules = @(Read-RulesFromPSTStore `
                    -PSTPath          $backupPath `
                    -BackupDisplayName "Backup -- $email")

                # For IMAP-ALREADY accounts the backup PST is an empty OST copy
                # with no rules -- rules live on the live IMAP store already.
                # Update their folder targets in place rather than recreating.
                if ($sourceRules.Count -eq 0 -and
                    $account.ProviderTag -eq 'IMAP-ALREADY') {
                    Write-Host '    Backup PST has no rules -- updating live IMAP store rules in place...' `
                               -ForegroundColor DarkGray
                    Write-OMMigrateLog -Message "RecreateRules: IMAP-ALREADY account $email -- remapping live rule folder targets from rules_inventory.csv." `
                                       -Level INFO

                    $updatedCount = 0
                    $skippedCount = 0
                    $notFoundCount = 0
                    # Collects rule names successfully updated this account, for the
                    # LastDeployedRun write-back after the per-rule loop below.
                    $rrUpdatedRuleNames = [System.Collections.Generic.List[string]]::new()
                    try {
                        # -- Find the live IMAP store for this account ----------
                        $liveStore = $null
                        $stores = $namespace.Stores
                        for ($si = 1; $si -le $stores.Count; $si++) {
                            $st = $stores.Item($si)
                            $stFp = ''
                            try { $stFp = $st.FilePath } catch { }
                            if ((-not ($stFp -like '*.pst')) -and
                                $st.DisplayName -like "*$email*") {
                                $liveStore = $st
                                break
                            }
                        }

                        if (-not $liveStore) {
                            Write-Host '    [SKIP] Live IMAP store not found in session.' `
                                       -ForegroundColor DarkGray
                            continue
                        }

                        $liveRules = $liveStore.GetRules()
                        Register-COMObject -ComObject $liveRules
                        Write-Host "    Found $($liveRules.Count) rule(s) on live IMAP store." `
                                   -ForegroundColor Cyan

                        # -- Build lookup of live COM rules keyed by RuleName+ExecutionOrder --
                        # ExecutionOrder is the tiebreaker for duplicate rule names.
                        # rules_inventory.csv uses the same composite key.
                        $liveRuleMap = [System.Collections.Generic.Dictionary[string,object]]::new(
                            [System.StringComparer]::OrdinalIgnoreCase
                        )
                        $liveRulesCollection = $liveRules   # keep reference for Save()
                        for ($ri = 1; $ri -le $liveRules.Count; $ri++) {
                            try {
                                $lr = $liveRules.Item($ri)
                                Register-COMObject -ComObject $lr
                                $lrName  = ''
                                $lrOrder = 0
                                try { $lrName  = $lr.Name           } catch { }
                                try { $lrOrder = $lr.ExecutionOrder } catch { }
                                $lrKey = "$lrName|$lrOrder"
                                if (-not $liveRuleMap.ContainsKey($lrKey)) {
                                    $liveRuleMap[$lrKey] = $lr
                                }
                            }
                            catch { }
                        }

                        # -- Get rules_inventory.csv rows for this account where NeedsFolderUpdate=True --
                        # TargetFolderPath in the CSV is the source of truth for the new folder target.
                        # The CSV path includes the account email as the first segment
                        # (e.g. user@example.com\Inbox\FolderName) which is correct --
                        # Get-FolderByPath navigates from the store root using the full path as-is.
                        $accountRulesInCSV = @($rulesNeedingUpdate | Where-Object {
                            $_.RuleStoreName -like "*$email*" -or
                            $_.TargetStoreName -like "*$email*"
                        })

                        # PENDING-ROW GATE (added 2026-07-07, Administrator direction): -RecreateRules
                        # calls Set-RuleConditions below (condition/consolidation-equivalent
                        # work), same as Invoke-DeployConsolidatedRules, so it is gated on the
                        # SAME column -- LastDeployedRun -- not LastTargetRun (that column is
                        # reserved for Phase 3's folder-target-only remap work). This branch
                        # previously had no idempotency gate at all -- every -RecreateRules
                        # invocation re-processed and re-wrote SenderAddress on every eligible
                        # rule regardless of whether it had already been correctly recreated in
                        # a prior run. Administrator's explicit direction: -RecreateRules should only
                        # process rules whose CSV row has a blank LastDeployedRun, matching the
                        # same pending-row convention used everywhere else in this pipeline. On
                        # a genuine first run every row is blank, so this naturally processes
                        # everything, unchanged from today's behavior.
                        $accountRulesInCSV = @($accountRulesInCSV | Where-Object {
                            $rrTimestamp = if ($_.PSObject.Properties['LastDeployedRun']) { [string]$_.LastDeployedRun } else { '' }
                            [string]::IsNullOrWhiteSpace($rrTimestamp)
                        })

                        Write-OMMigrateLog -Message (
                            "RecreateRules IMAP-ALREADY: $email -- " +
                            "$($accountRulesInCSV.Count) pending rule(s) from rules_inventory.csv (NeedsFolderUpdate=True, blank LastDeployedRun)"
                        ) -Level INFO

                        if ($accountRulesInCSV.Count -eq 0) {
                            Write-Host '    [SKIP] No pending rules (blank LastDeployedRun) found in rules_inventory.csv for this account.' `
                                       -ForegroundColor DarkGray
                            continue
                        }

                        # -- Process each CSV rule row ---------------------------
                        foreach ($csvRule in $accountRulesInCSV) {
                            $csvRuleName  = $csvRule.RuleName
                            $csvExecOrder = 0
                            try { $csvExecOrder = [int]$csvRule.ExecutionOrder } catch { }
                            $csvPath      = $csvRule.TargetFolderPath

                            # Find the live COM rule by RuleName+ExecutionOrder
                            $lrKey      = "$csvRuleName|$csvExecOrder"
                            $liveRuleCOM = $null
                            if ($liveRuleMap.ContainsKey($lrKey)) {
                                $liveRuleCOM = $liveRuleMap[$lrKey]
                            }
                            else {
                                # Fallback: match by name only (ExecutionOrder may differ after restore)
                                foreach ($key in $liveRuleMap.Keys) {
                                    if ($key -like "$csvRuleName|*") {
                                        $liveRuleCOM = $liveRuleMap[$key]
                                        break
                                    }
                                }
                            }

                            if (-not $liveRuleCOM) {
                                Write-OMMigrateLog -Message "RecreateRules IMAP-ALREADY: Live rule not found for '$csvRuleName' (ExecutionOrder=$csvExecOrder) -- skipping." `
                                                   -Level WARN
                                $notFoundCount++
                                continue
                            }

                            # -- Resolve destination folder COM object -----------
                            # Look up the folder path in folder_map.csv to determine
                            # Server vs Local destination, then navigate to the folder.
                            $folderMapEntry = $folderMap | Where-Object {
                                $_.FolderPath -eq $csvPath
                            } | Select-Object -First 1

                            $newFolderCOM = $null
                            if ($folderMapEntry) {
                                if ($folderMapEntry.Destination -eq 'Server') {
                                    # Server destination -- navigate from IMAP store root.
                                    #
                                    # FIXED (2026-07-09, Administrator direction, scoped to this branch
                                    # only): $csvPath comes from rules_inventory.csv's
                                    # TargetFolderPath, which always includes the account email
                                    # as its first segment (e.g.
                                    # 'user@example.com\Inbox\FolderName') -- correct
                                    # when navigating from a shared multi-account root like the
                                    # Archive PST (see the Local branch below, unchanged,
                                    # correct as-is), but WRONG here: $imapStoreRoot is already
                                    # that single account's own IMAP store root, which has no
                                    # subfolder literally named after the account's own email
                                    # address. Passing the full account-email-prefixed path
                                    # against an already-account-scoped root meant the first
                                    # path segment could never match, so Get-FolderByPath
                                    # always returned $null for every Server-destination rule
                                    # in this branch. Confirmed via code review 2026-07-09 --
                                    # never live-hit in practice because this requires the
                                    # narrow combination of -RecreateRules (opt-in recovery
                                    # switch, not part of normal pipeline runs) + IMAP-ALREADY
                                    # + empty backup PST + a rule specifically targeting a
                                    # Server-destination folder (folder_map.csv defaults
                                    # everything to Local; Server requires explicit operator
                                    # opt-in per folder). Fix: strip the leading
                                    # account-email segment before navigating from
                                    # $imapStoreRoot, mirroring how Remove-StorePrefix strips
                                    # a store-name prefix elsewhere in this codebase for the
                                    # same reason (root-relative path vs. already-scoped root).
                                    #
                                    # NOT YET LIVE-TESTED (2026-07-09, Administrator decision): this
                                    # path requires an artificial setup to exercise (an
                                    # IMAP-ALREADY account with an empty/unattached backup PST,
                                    # -RecreateRules, and a rule whose folder_map.csv entry is
                                    # deliberately set to Server rather than the default Local)
                                    # that doesn't arise naturally from normal pipeline use.
                                    # Administrator's call: fix is logically sound and mirrors an
                                    # already-proven pattern elsewhere in this file, not worth
                                    # manufacturing a live test for right now given this is a
                                    # rarely-hit, opt-in recovery-only branch. Test live the
                                    # next time -RecreateRules is genuinely needed for a real
                                    # recovery, or during a future dedicated test/documentation
                                    # pass -- do not assume this is verified until then.
                                    try {
                                        $imapStoreRoot = $liveStore.GetRootFolder()
                                        Register-COMObject -ComObject $imapStoreRoot

                                        $serverRelativePath = $csvPath
                                        $csvPathSegments    = $csvPath -split '\\'
                                        if ($csvPathSegments.Count -gt 1 -and
                                            $csvPathSegments[0] -eq $email) {
                                            $serverRelativePath = ($csvPathSegments[1..($csvPathSegments.Count - 1)] -join '\')
                                        }

                                        $newFolderCOM = Get-FolderByPath `
                                            -RootFolder      $imapStoreRoot `
                                            -FolderPath      $serverRelativePath `
                                            -CreateIfMissing $false
                                    }
                                    catch { }
                                }
                                elseif ($folderMapEntry.Destination -eq 'Local') {
                                    # Local destination -- navigate from Archive PST root
                                    # CSV path includes account email as first segment
                                    # (e.g. user@example.com\Inbox\FolderName)
                                    # which correctly navigates the Archive PST subfolder structure.
                                    #
                                    # Added 2026-07-10, Administrator (multi-archive support,
                                    # TargetStoreName hardcode fix). $csvRule (this loop's own
                                    # CSV row) may carry a TargetStoreName pointing at a
                                    # different attached archive PST than $archiveRootFolder --
                                    # resolve and use that store's root instead when present and
                                    # attached; otherwise fall straight back to
                                    # $archiveRootFolder exactly as before this fix, so any run
                                    # not using multi-archive mappings is unaffected.
                                    $effectiveArchiveRootForRemap = $archiveRootFolder
                                    if ($csvRule.PSObject.Properties['TargetStoreName'] -and
                                        -not [string]::IsNullOrWhiteSpace($csvRule.TargetStoreName)) {

                                        $wantedRemapStoreName = $csvRule.TargetStoreName

                                        if (-not $script:resolvedArchiveStoreFoldersForRemap.ContainsKey($wantedRemapStoreName)) {
                                            $foundRemapStoreFolder = $null
                                            try {
                                                $liveStoresForRemap = $namespace.Stores
                                                for ($rmi = 1; $rmi -le $liveStoresForRemap.Count; $rmi++) {
                                                    $lookupRemapStore = $liveStoresForRemap.Item($rmi)
                                                    if ($lookupRemapStore.DisplayName -eq $wantedRemapStoreName) {
                                                        try { $foundRemapStoreFolder = $lookupRemapStore.GetRootFolder() } catch { }
                                                        break
                                                    }
                                                }
                                            }
                                            catch { }

                                            if (-not $foundRemapStoreFolder) {
                                                Write-OMMigrateLog -Message (
                                                    "RecreateRules IMAP-ALREADY: Rule '$csvRuleName' -- TargetStoreName " +
                                                    "'$wantedRemapStoreName' not found among attached stores -- " +
                                                    "falling back to the default Archive PST for this rule."
                                                ) -Level WARN
                                            }
                                            $script:resolvedArchiveStoreFoldersForRemap[$wantedRemapStoreName] = $foundRemapStoreFolder
                                        }

                                        if ($script:resolvedArchiveStoreFoldersForRemap[$wantedRemapStoreName]) {
                                            $effectiveArchiveRootForRemap = $script:resolvedArchiveStoreFoldersForRemap[$wantedRemapStoreName]
                                        }
                                    }

                                    if ($effectiveArchiveRootForRemap) {
                                        try {
                                            $newFolderCOM = Get-FolderByPath `
                                                -RootFolder      $effectiveArchiveRootForRemap `
                                                -FolderPath      $csvPath `
                                                -CreateIfMissing $false
                                        }
                                        catch { }
                                    }
                                }
                            }
                            else {
                                Write-OMMigrateLog -Message "RecreateRules IMAP-ALREADY: No folder_map.csv entry for path '$csvPath' (Rule: '$csvRuleName') -- skipping." `
                                                   -Level WARN
                                $skippedCount++
                                continue
                            }

                            if (-not $newFolderCOM) {
                                Write-OMMigrateLog -Message "RecreateRules IMAP-ALREADY: Target folder not found in destination for '$csvRuleName' (path='$csvPath', dest='$($folderMapEntry.Destination)') -- skipping." `
                                                   -Level WARN
                                $skippedCount++
                                continue
                            }

                            # -- Update MoveToFolder action on the live COM rule -
                            # Use InvokeMember reflection -- standard PS property
                            # assignment ($action.Folder = $x) silently fails due
                            # to CLR stripping the COM pointer. InvokeMember calls
                            # put_Folder directly, replicating native VBA behavior.
                            $actionUpdated = $false

                            # -- Set rule conditions via shared module function --
                            # Set-RuleConditions handles: From clear, SenderAddress,
                            # OnLocalMachine, Account, ExecutionAccount.
                            $sendersDomain = ''
                            if ($csvRule.PSObject.Properties['SendersDomain'] -and
                                -not [string]::IsNullOrWhiteSpace($csvRule.SendersDomain)) {
                                $sendersDomain = $csvRule.SendersDomain
                            } elseif (-not [string]::IsNullOrWhiteSpace($csvRule.TargetFolderPath)) {
                                # Default: last segment of TargetFolderPath
                                $sendersDomain = ($csvRule.TargetFolderPath -split '\\' |
                                    Where-Object { $_ -ne '' } | Select-Object -Last 1)
                            }
                            [void](Set-RuleConditions `
                                -Rule          $liveRuleCOM `
                                -RuleName      $csvRuleName `
                                -RuleStoreName $csvRule.RuleStoreName `
                                -SendersDomain $sendersDomain `
                                -Namespace     $namespace)

                            try {
                                $moveAction = $liveRuleCOM.Actions.MoveToFolder
                                # No guard on Enabled -- always attempt assignment via InvokeMember
                                # regardless of current enabled state. Matches Gemini pattern.
                                $folderSet          = Set-RuleFolderAction -Action $moveAction -Folder $newFolderCOM
                                [void]$moveAction.GetType().InvokeMember('Enabled', [System.Reflection.BindingFlags]::SetProperty, $null, $moveAction, @($true))
                                [void]$liveRuleCOM.GetType().InvokeMember('Enabled', [System.Reflection.BindingFlags]::SetProperty, $null, $liveRuleCOM, @($true))
                                $actionUpdated         = $folderSet
                                Write-OMMigrateLog -Message "IMAP-ALREADY '$csvRuleName': MoveToFolder InvokeMember result=$folderSet. FolderNull=$($null -eq $moveAction.Folder)" -Level DEBUG
                            }
                            catch { }

                            # Also update CopyToFolder if enabled
                            try {
                                $copyAction = $liveRuleCOM.Actions.CopyToFolder
                                if ($copyAction.Enabled) {
                                    $folderSet          = Set-RuleFolderAction -Action $copyAction -Folder $newFolderCOM
                                    [void]$copyAction.GetType().InvokeMember('Enabled', [System.Reflection.BindingFlags]::SetProperty, $null, $copyAction, @($true))
                                    if ($folderSet) { $actionUpdated = $true }
                                }
                            }
                            catch { }

                            # Always set StopProcessing on every rule regardless of folder resolution.
                            # Item(27) used instead of Actions.StopProcessingRules named property --
                            # the named property is unreliable via COM interop and returns null
                            # even when Actions.Count confirms the action slot exists. Item(27) is
                            # the confirmed fixed index for ActionType=18 (olRuleActionStopProcessingRules)
                            # in this Outlook COM implementation (confirmed June 17, 2026).
                            try { $spr = $liveRuleCOM.Actions.Item(27); [void]$spr.GetType().InvokeMember('Enabled', [System.Reflection.BindingFlags]::SetProperty, $null, $spr, @($true)) } catch { }

                            if ($actionUpdated) {
                                $updatedCount++
                                # Track for LastDeployedRun write-back below -- InvokeMember
                                # already committed this rule's changes at COM level (no
                                # batch Save() in this branch, per the comment further down),
                                # so it's safe to consider this rule "done" immediately.
                                [void]$rrUpdatedRuleNames.Add($csvRuleName)
                                Write-OMMigrateLog -Message "RecreateRules IMAP-ALREADY: Updated '$csvRuleName' -> '$csvPath' ($($folderMapEntry.Destination))" `
                                                   -Level DEBUG
                            }
                            else {
                                $skippedCount++
                                Write-OMMigrateLog -Message "RecreateRules IMAP-ALREADY: No enabled folder action on '$csvRuleName' -- skipped." `
                                                   -Level INFO
                            }
                        }

                        # -- Summary after all rules processed ------------------
                        # No Save() needed -- InvokeMember commits at COM level.
                        if ($updatedCount -gt 0) {
                            Write-Host "    [OK]   Updated  : $updatedCount rule(s) folder targets remapped." `
                                       -ForegroundColor Green
                        }
                        else {
                            Write-Host "    [OK]   No folder targets updated (Skipped=$skippedCount | NotFound=$notFoundCount)." `
                                       -ForegroundColor DarkGray
                        }
                        if ($skippedCount -gt 0) {
                            Write-Host "    [WARN] Skipped   : $skippedCount  (no folder map entry or folder not found)" `
                                       -ForegroundColor Yellow
                        }
                        if ($notFoundCount -gt 0) {
                            Write-Host "    [WARN] Not found : $notFoundCount  (rule name not found in live Outlook)" `
                                       -ForegroundColor Yellow
                        }

                        # LastDeployedRun WRITE-BACK (added 2026-07-07, Administrator direction):
                        # stamp LastDeployedRun for every rule successfully updated this
                        # account. Committed immediately per-rule at COM level (no batch
                        # Save() in this branch), so it's safe to stamp right after the loop
                        # rather than waiting for a separate confirmation step. Re-reads
                        # rules_inventory.csv fresh from disk immediately before writing to
                        # minimize the chance of clobbering fields changed elsewhere in this
                        # same run, then writes back the FULL set (preserving every other
                        # column exactly), matching the canonical column order.
                        if ($rrUpdatedRuleNames.Count -gt 0) {
                            try {
                                $rrTimestamp = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.ffffffzzz')
                                $rrRows      = @(Import-Csv -Path $rulesInventoryPath -Encoding UTF8 |
                                    Where-Object { $_.RuleStoreName -and -not [string]::IsNullOrWhiteSpace($_.RuleStoreName) -and
                                                   $_.RuleName      -and -not [string]::IsNullOrWhiteSpace($_.RuleName) })
                                $rrNameSet   = [System.Collections.Generic.HashSet[string]]::new(
                                    [string[]]$rrUpdatedRuleNames, [System.StringComparer]::OrdinalIgnoreCase
                                )
                                $rrStampedCount = 0
                                foreach ($rrRow in $rrRows) {
                                    if (($rrRow.RuleStoreName -like "*$email*" -or $rrRow.TargetStoreName -like "*$email*") -and
                                        $rrNameSet.Contains($rrRow.RuleName)) {
                                        if (-not $rrRow.PSObject.Properties['LastDeployedRun']) {
                                            Add-Member -InputObject $rrRow -MemberType NoteProperty -Name 'LastDeployedRun' -Value '' -Force
                                        }
                                        $rrRow.LastDeployedRun = $rrTimestamp
                                        $rrStampedCount++
                                    }
                                }
                                $rrCanonicalColumns = @(
                                    'RuleStoreName','TargetStoreName','RuleName','LastDeployedRun','LastTargetRun',
                                    'TargetFolderPath','SendersDomain','NeedsFolderUpdate','IsEnabled',
                                    'ExecutionOrder','RuleType','StopProcessing','Conditions',
                                    'Actions','TargetFolderEntryID','Notes'
                                )
                                # ADDED (Administrator direction, 2026-08-18): re-insert blank separator
                                # rows between RuleStoreName groups before writing -- see
                                # Add-RulesCsvSeparatorRows header comment (OMMigrate-Outlook.psm1)
                                # for full rationale. $rrRows above is already filtered to real
                                # data rows (see the Where-Object above its Import-Csv), so no
                                # further filtering needed before handing off to the helper.
                                $rrRowsWithSeparators = Add-RulesCsvSeparatorRows -Rows (@($rrRows | Select-Object $rrCanonicalColumns))
                                $rrRowsWithSeparators | Export-Csv -Path $rulesInventoryPath -NoTypeInformation -Encoding UTF8
                                Write-OMMigrateLog -Message "LastDeployedRun written for $rrStampedCount rule(s) in RecreateRules IMAP-ALREADY ($email)." -Level INFO
                            }
                            catch {
                                Write-OMMigrateLog -Message "Failed to write LastDeployedRun timestamps for RecreateRules IMAP-ALREADY ($email) (non-fatal, rules will re-process next run): $_" -Level WARN
                            }
                        }

                        $Script:RulesRecreated += $updatedCount

                        Write-AuditEntry -Action 'RULES_RECREATED' `
                                         -AccountEmail $email `
                                         -Detail "Updated=$updatedCount | Skipped=$skippedCount | NotFound=$notFoundCount (IMAP-ALREADY CSV-driven remap)"

                        # -- Additive: roll into script-scope rules-reporting totals --
                        # NotFoundCount (rule name not found in live Outlook) is folded into
                        # the Skipped bucket for reporting purposes -- both represent "no
                        # action taken on this rule", just for different reasons.
                        $Script:RulesSkippedTotal += ($skippedCount + $notFoundCount)

                        # Per-store summary row for the new Rules Processing Detail report table.
                        $Script:RulesStoreSummary.Add([PSCustomObject]@{
                            StoreName             = $email
                            Created               = $updatedCount
                            Skipped               = ($skippedCount + $notFoundCount)
                            Failed                = 0
                            StopProcessingSet     = 0
                            StopProcessingFailed  = 0
                        })
                    }
                    catch {
                        Write-OMMigrateLog -Message "RecreateRules: Live rule remap failed for $email : $_" -Level WARN
                        Write-Host "    [WARN] Live rule remap failed: $_" -ForegroundColor Yellow
                    }
                    continue
                }

                if ($sourceRules.Count -eq 0) {
                    Write-Host '    [SKIP] No rules found in backup PST.' `
                               -ForegroundColor DarkGray
                    Write-OMMigrateLog -Message "RecreateRules: No rules found in backup PST for $email." `
                                       -Level INFO
                    continue
                }

                Write-Host "    Found $($sourceRules.Count) rule(s) in backup PST." `
                           -ForegroundColor Cyan

                # Filter folder map to this account's entries only
                $accountFolderMap = @($folderMap | Where-Object {
                    $_.StoreName -and (
                        $_.StoreName -like "*$email*" -or
                        $_.StoreName -eq $account.DisplayName
                    )
                })

                Write-OMMigrateLog -Message (
                    "RecreateRules: $email -- $($sourceRules.Count) source rule(s) | " +
                    "$($accountFolderMap.Count) folder map entries"
                ) -Level INFO

                # Confirm per account unless -Force
                $proceedRules = $true
                if (-not $Force) {
                    $proceedRules = Confirm-Action `
                        -Message      "Recreate $($sourceRules.Count) rule(s) for: $email ?" `
                        -AccountEmail $email `
                        -DefaultYes   $true
                }

                if (-not $proceedRules) {
                    Write-OMMigrateLog -Message "RecreateRules: Skipped by operator for $email." `
                                       -Level INFO
                    Write-Host '    [SKIP] Skipped by operator.' -ForegroundColor DarkGray
                    continue
                }

                # Recreate rules on IMAP store
                # Filter rules_inventory.csv to this account's rules with NeedsFolderUpdate=True.
                # Passed to Invoke-RulesRecreation as the source of truth for TargetFolderPath.
                $accountRulesInventory = @($rulesNeedingUpdate | Where-Object {
                    $_.RuleStoreName -like "*$email*" -or
                    $_.TargetStoreName -like "*$email*"
                })

                $recStats = Invoke-RulesRecreation `
                    -SourceRules       $sourceRules `
                    -TargetStoreName   $email `
                    -AccountEmail      $email `
                    -FolderMap         $accountFolderMap `
                    -Namespace         $namespace `
                    -ArchiveRootFolder $archiveRootFolder `
                    -RulesInventory    $accountRulesInventory

                $Script:RulesRecreated += $recStats.Created

                # Console summary per account
                if ($recStats.Created -gt 0) {
                    Write-Host "    [OK]   Created  : $($recStats.Created)" `
                               -ForegroundColor Green
                }
                if ($recStats.Skipped -gt 0) {
                    Write-Host "    [SKIP] Skipped  : $($recStats.Skipped)  (already existed)" `
                               -ForegroundColor DarkGray
                }
                if ($recStats.Failed -gt 0) {
                    Write-Host "    [WARN] Failed   : $($recStats.Failed)  (see log for details)" `
                               -ForegroundColor Yellow
                    if ($Script:FinalStatus -eq 'SUCCESS') { $Script:FinalStatus = 'WARNING' }
                }

                Write-AuditEntry -Action 'RULES_RECREATED' `
                                 -AccountEmail $email `
                                 -Detail (
                                     "Created=$($recStats.Created) | " +
                                     "Skipped=$($recStats.Skipped) | " +
                                     "Failed=$($recStats.Failed)"
                                 )

                # -- Additive: roll into script-scope rules-reporting totals --
                $Script:RulesSkippedTotal += $recStats.Skipped

                # Per-store summary row for the new Rules Processing Detail report table.
                $Script:RulesStoreSummary.Add([PSCustomObject]@{
                    StoreName             = $email
                    Created               = $recStats.Created
                    Skipped               = $recStats.Skipped
                    Failed                = $recStats.Failed
                    StopProcessingSet     = 0
                    StopProcessingFailed  = 0
                })

                # Detach the backup PST if it was opened by Read-RulesFromPSTStore
                # and is not also being used for folder migration this run.
                # Close-PSTFile is safe to call -- it no-ops if not mounted.
                # Also skip detach entirely if the PST was already mounted in
                # the profile before this script touched it.
                if (-not $Script:EffectiveSkipFolderMigration) {
                    # Folder migration already detached it -- skip
                    Write-OMMigrateLog -Message "RecreateRules: Backup PST detach deferred to folder migration for $email." `
                                       -Level DEBUG
                }
                elseif ($backupPathWasAlreadyMounted) {
                    Write-OMMigrateLog -Message "RecreateRules: Backup PST was already mounted before this script ran -- leaving attached: $backupPath" `
                                       -Level DEBUG
                }
                else {
                    Close-PSTFile -PSTPath $backupPath | Out-Null
                    Write-OMMigrateLog -Message "RecreateRules: Backup PST detached after rules extraction: $backupPath" `
                                       -Level DEBUG
                }
            }

            Write-Host ''
            Write-Host '  Rules Recreation Summary:' -ForegroundColor White
            Write-Host "    Total rules recreated : $Script:RulesRecreated" -ForegroundColor $(
                if ($Script:RulesRecreated -gt 0) { 'Green' } else { 'Gray' }
            )
            Write-Host ''
        }
    }
    elseif (-not $RecreateRules) {
        Write-OMMigrateLog -Message 'Rules recreation skipped (use -RecreateRules to enable).' `
                           -Level DEBUG
    }

    # -- Rules Update Phase -------------------------------------
    # Rules always run in normal operation.
    if ($rulesNeedingUpdate.Count -gt 0) {
        Show-SectionHeader -Title 'Updating Outlook Rules'

        Write-Host "  $($rulesNeedingUpdate.Count) rule(s) require folder target updates." `
                   -ForegroundColor Cyan
        Write-Host ''

        $rulesUpdateConfirmed = Confirm-Action `
            -Message "Update $($rulesNeedingUpdate.Count) Outlook Rule folder targets now?" `
            -DefaultYes $true

        if ($rulesUpdateConfirmed) {
            Write-Host '  Updating rules...' -ForegroundColor Cyan

            # -- WhatIf -- log what would be done, no COM changes ----------
            if ($Global:OMMigrate.WhatIf) {
                Write-Host "  [PREVIEW] $($rulesNeedingUpdate.Count) rule(s) would be updated." `
                           -ForegroundColor Magenta
                Write-Host '  [PREVIEW] Rule details logged -- not displayed to avoid console flood.' `
                           -ForegroundColor DarkGray
                foreach ($ruleRow in $rulesNeedingUpdate) {
                    $folderMapEntry = $folderMap | Where-Object {
                        $_.FolderPath -eq $ruleRow.TargetFolderPath
                    } | Select-Object -First 1

                    $dest = if ($folderMapEntry) { $folderMapEntry.Destination } else { 'No folder map entry' }
                    # Log rule details but do not write to console -- 800+ rules would flood output
                    Write-OMMigrateLog -Message (
                        "WhatIf: Would update rule '$($ruleRow.RuleName)' | " +
                        "Old target: $($ruleRow.TargetFolderPath) | " +
                        "Destination: $dest"
                    ) -Level INFO -WhatIfPrefix
                }
            }
            else {
                # -- Build account lookup dictionary ----------------------
                # Maps SmtpAddress -> Account COM object (primary key).
                # Also maps DisplayName -> Account COM object (fallback).
                # Used to set the 'through the specified account' condition
                # on every remapped rule so it fires for the correct IMAP
                # account inbox. Universal -- works for all account types.
                # Gemini: match SmtpAddress first, fall back to DisplayName.
                $accountLookup = [System.Collections.Generic.Dictionary[string,object]]::new(
                    [System.StringComparer]::OrdinalIgnoreCase
                )
                try {
                    foreach ($acct in $namespace.Accounts) {
                        try {
                            $acctSmtp = ''
                            $acctDisp = ''
                            try { $acctSmtp = $acct.SmtpAddress  } catch { }
                            try { $acctDisp = $acct.DisplayName  } catch { }
                            # Index by SmtpAddress (primary)
                            if (-not [string]::IsNullOrWhiteSpace($acctSmtp) -and
                                -not $accountLookup.ContainsKey($acctSmtp)) {
                                $accountLookup[$acctSmtp] = $acct
                            }
                            # Index by DisplayName (fallback for accounts where SmtpAddress is empty)
                            if (-not [string]::IsNullOrWhiteSpace($acctDisp) -and
                                -not $accountLookup.ContainsKey($acctDisp)) {
                                $accountLookup[$acctDisp] = $acct
                            }
                        } catch { }
                    }
                    Write-OMMigrateLog -Message "Account lookup built: $($accountLookup.Count) account(s) indexed." `
                                       -Level INFO
                } catch {
                    Write-OMMigrateLog -Message "Failed to build account lookup (non-fatal): $_" -Level INFO
                }

                # -- Live rules update via COM ------------------------------
                # Get the Rules collection by scanning ALL stores, not just
                # DefaultStore. DefaultStore can be reassigned (e.g. to Archive
                # PST after POP3 removal) and may have no rules at all.
                # Rules live on the store that owns the account -- typically the
                # IMAP OST store for the migrated account. Scan all stores and
                # build a combined rules map keyed by rule name so the lookup
                # below works regardless of which store owns each rule.
                $rulesCollection = $null
                $rulesSaveNeeded = $false
                $allLiveRules    = [System.Collections.Generic.Dictionary[string,object]]::new(
                    [System.StringComparer]::OrdinalIgnoreCase
                )
                # liveCollections: store display name -> rules collection object.
                # NOT registered with Register-COMObject -- must stay alive until
                # after Save() is called. Deferred release would destroy the RCW
                # before Save(), so we hold these references explicitly and release
                # them manually after saving.
                $liveCollections = [System.Collections.Generic.Dictionary[string,object]]::new(
                    [System.StringComparer]::OrdinalIgnoreCase
                )
                # ruleCollectionMap: rule name -> store display name (for Save() lookup)
                $ruleCollectionMap = [System.Collections.Generic.Dictionary[string,string]]::new(
                    [System.StringComparer]::OrdinalIgnoreCase
                )
                # collectionsNeedingSave: store display names that had rule updates
                $collectionsNeedingSave = [System.Collections.Generic.HashSet[string]]::new(
                    [System.StringComparer]::OrdinalIgnoreCase
                )

                try {
                    $allStores = $namespace.Stores
                    Register-COMObject -ComObject $allStores
                    for ($si = 1; $si -le $allStores.Count; $si++) {
                        try {
                            $st = $allStores.Item($si)
                            Register-COMObject -ComObject $st
                            $stDisplayName = ''
                            try { $stDisplayName = $st.DisplayName } catch { }
                            $stRules = $null
                            # CHANGED (2026-07-08, Administrator explicit direction): the skip below
                            # this comment used to exclude secondary stores from this same
                            # loop entirely, based on a June 2026 comment claiming plain
                            # GetRules() on secondary stores (without a byte-level zeroing
                            # step) causes "UI bleed into ameritech on restart." Administrator has
                            # confirmed that comment was a false assumption made at the time,
                            # conflating this with the unrelated devil-code (0x800C8101)
                            # investigation -- which was later root-caused (2026-06-29) to a
                            # completely different mechanism (profile-complexity-driven
                            # Item() enumeration ceiling at scale), not a rules-collection-
                            # fetch-order issue at all. Administrator's explicit direction: treat every
                            # account, primary or secondary, identically for GetRules() and
                            # the LastTargetRun folder-target remap pass -- same plain
                            # GetRules() call already proven safe elsewhere in this file (the
                            # -RecreateRules IMAP-ALREADY branch), no special-casing, no
                            # zeroing step. Secondary stores now flow through this exact same
                            # loop and populate $liveCollections identically to the default
                            # store, so Phase 3 further below processes them with the exact
                            # same code path -- no separate logic was written for them.
                            try {
                                $stRules = $st.GetRules()
                                # Do NOT Register-COMObject -- must stay alive for Save()
                                if ($stDisplayName) { $liveCollections[$stDisplayName] = $stRules }
                            }
                            catch { continue }   # store doesn't support GetRules -- skip
                            for ($ri = 1; $ri -le $stRules.Count; $ri++) {
                                try {
                                    $r = $stRules.Item($ri)
                                    Register-COMObject -ComObject $r
                                    if (-not $allLiveRules.ContainsKey($r.Name)) {
                                        $allLiveRules[$r.Name]        = $r
                                        $ruleCollectionMap[$r.Name]   = $stDisplayName
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                    Write-OMMigrateLog -Message "Rules scan complete: $($allLiveRules.Count) unique rules found across all stores." `
                                       -Level INFO

                    # -- Default-store and secondary-store rule recreation -----------------
                    # REMOVED 2026-06-26 (explicitly authorized by Administrator): this block used to
                    # contain (a) a permanent guard blocking purge/recreate on the default
                    # store (ameritech) because Outlook COM's rule .Create() could not
                    # originate the StopProcessing action's enabled state at creation time,
                    # and (b) a full inline purge/recreate implementation for secondary
                    # stores with its own Match/Additive/Full idempotency logic, bytes
                    # 44-45 zero/patch steps, and a post-Save StopProcessing pass.
                    #
                    # Both are superseded by Invoke-DeployConsolidatedRules (and its helper
                    # Invoke-BuildRulesFromMap) in OMMigrate-Outlook.psm1 -- a PowerShell
                    # port of Gemini's DeployConsolidatedRules VBA macro. That function
                    # creates rules and sets StopProcessing correctly via Actions.Item(27)
                    # after each rule's Save(), which is why the limitation that motivated
                    # the old default-store guard no longer applies. All accounts now go
                    # through the same single path -- wired in immediately below.
                    #
                    # CORRECTED 2026-07-06, Administrator: this comment previously (and
                    # incorrectly) stated the function "creates rules across ALL
                    # accounts... no exceptions" -- that was true before today, but was
                    # itself the documented version of a real bug Administrator caught live:
                    # picker selection was NOT actually authoritative for what this
                    # function did, only for what got reported. Fixed by adding
                    # -ScopedAccountNames below (see Invoke-DeployConsolidatedRules's
                    # own updated .DESCRIPTION in OMMigrate-Outlook.psm1 for full
                    # detail) -- now only the picker-selected account(s) in
                    # $accountsToProcess are eligible for rule creation, LastDeployedRun
                    # stamping, AND the resort pass, consistent with every other
                    # picker-scoped step in this pipeline.
                    #
                    # $rulesCollection is still set to the non-null sentinel here so the
                    # existing "Strategy 1 + 2" per-store remap loop below (folder-target
                    # remapping for rules that already exist) continues to run as before --
                    # that loop is unrelated to consolidated-rule creation and was not
                    # removed.
                    $rulesCollection = $allLiveRules

                    # -- Consolidated rule deployment (Gemini's DeployConsolidatedRules,
                    #    ported to PowerShell) -- scoped to picker-selected account(s)
                    #    only (added 2026-07-06, Administrator -- see comment above). Reads
                    #    pending (blank LastDeployedRun) rows from rules_inventory.csv
                    #    for those account(s) only, creates consolidated rules, and
                    #    writes timestamps back to the CSV so this function and Gemini's
                    #    macro stay in sync with each other regardless of which one runs.
                    if ($archiveRootFolder) {
                        try {
                            $consolidatedRulesCsv = Import-Csv -Path $rulesInventoryPath -Encoding UTF8

                            # ADDED (Administrator direction -- Rules Updated double-count fix):
                            # count, BEFORE Pass 1 (Invoke-DeployConsolidatedRules) touches
                            # anything, how many of THIS run's scoped rows have BOTH
                            # LastDeployedRun and LastTargetRun blank. A row in this state
                            # will be processed by Pass 1 (creates/consolidates, since
                            # LastDeployedRun is blank) AND by Pass 2 below (remaps folder
                            # target, since LastTargetRun is blank), and each pass
                            # legitimately increments $Script:RulesUpdated for its own
                            # reason -- so the same physical rule is counted twice in a
                            # single run. Administrator's requirement: an admin may blank either
                            # date alone (only one pass fires -- count once, already
                            # correct) or both dates together (both passes fire on the
                            # same rule -- must still count once, not twice). Counted here,
                            # from the pre-Pass-1 CSV snapshot, because this is the only
                            # point where a row's identity is still its pre-consolidation
                            # RuleName/RuleStoreName -- Pass 2 later loops LIVE rules by
                            # their CURRENT name, which for a freshly-consolidated rule is
                            # already the new standardized name and no longer matches this
                            # snapshot. Scoping (RuleStoreName vs $scopedAccountEmails,
                            # trimmed, case-insensitive) matches Invoke-DeployConsolidated-
                            # Rules's own $scopedAccountLookup logic exactly, so this count
                            # reflects only rows this run will actually touch.
                            $rulesUpdatedDoubleCountAdjustment = 0
                            $scopedAccountEmailsForCount = @($accountsToProcess | ForEach-Object { $_.EmailAddress.Trim() })
                            $scopedAccountLookupForCount = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                            foreach ($saCount in $scopedAccountEmailsForCount) { [void]$scopedAccountLookupForCount.Add($saCount) }
                            foreach ($bothBlankRow in $consolidatedRulesCsv) {
                                $bbStoreName = if ($bothBlankRow.PSObject.Properties['RuleStoreName']) { [string]$bothBlankRow.RuleStoreName.Trim() } else { '' }
                                if (-not $scopedAccountLookupForCount.Contains($bbStoreName)) { continue }
                                $bbDeployed = if ($bothBlankRow.PSObject.Properties['LastDeployedRun']) { [string]$bothBlankRow.LastDeployedRun } else { '' }
                                $bbTargeted = if ($bothBlankRow.PSObject.Properties['LastTargetRun'])   { [string]$bothBlankRow.LastTargetRun }   else { '' }
                                if ([string]::IsNullOrWhiteSpace($bbDeployed) -and [string]::IsNullOrWhiteSpace($bbTargeted)) {
                                    $rulesUpdatedDoubleCountAdjustment++
                                }
                            }
                            if ($rulesUpdatedDoubleCountAdjustment -gt 0) {
                                Write-OMMigrateLog -Message "Rules Updated double-count adjustment: $rulesUpdatedDoubleCountAdjustment rule(s) have both LastDeployedRun and LastTargetRun blank this run -- will be subtracted once from `$Script:RulesUpdated after both passes complete, so each such rule counts once, not twice." -Level INFO
                            }

                            # Picker-scoping fix (added 2026-07-06, Administrator): pass the
                            # picker's selected account email addresses through so
                            # Invoke-DeployConsolidatedRules only acts on these
                            # accounts -- every other account's pending rows are left
                            # untouched this run, per Administrator's explicit requirement that
                            # picker selection be authoritative for all activity, not
                            # just what gets reported.
                            $scopedAccountEmails = @($accountsToProcess | ForEach-Object { $_.EmailAddress })
                            $deployResult = Invoke-DeployConsolidatedRules `
                                -Namespace           $namespace `
                                -RulesInventory      $consolidatedRulesCsv `
                                -CsvPath             $rulesInventoryPath `
                                -ArchiveRootFolder   $archiveRootFolder `
                                -SaveBatchSize       $SaveBatchSize `
                                -ScopedAccountNames  $scopedAccountEmails
                            Write-OMMigrateLog -Message (
                                "Invoke-DeployConsolidatedRules: Created=$($deployResult.Created) " +
                                "Failed=$($deployResult.Failed) DomainsProcessed=$($deployResult.DomainsProcessed)."
                            ) -Level INFO

                            # ADDED (Administrator direction, this same fix): Invoke-DeployConsolidatedRules
                            # writes rules_inventory.csv to disk BEFORE returning here (Step 4 of that
                            # function -- LastDeployedRun stamping, and now also the standardized-name
                            # rename this same fix adds). Script 03's own $rulesInventory in memory was
                            # loaded earlier in this run (during the live rules scan, before any of
                            # THIS run's consolidation happened) and was never refreshed after -- every
                            # downstream check in Phase 3 ($phase3PendingNames, $csvRuleMap, the live-
                            # folder-path CSV lookup) was reading a stale snapshot missing this run's
                            # own renames and timestamp writes. Reloading here, immediately after the
                            # function that just wrote the current file, so Phase 3 sees the true
                            # current state instead of requiring a second Script 03 run to catch up.
                            if (Test-Path $rulesInventoryPath) {
                                try {
                                    $rulesInventory = Import-Csv -Path $rulesInventoryPath -Encoding UTF8
                                    Write-OMMigrateLog -Message "rules_inventory.csv reloaded after Invoke-DeployConsolidatedRules -- $($rulesInventory.Count) row(s), reflects this run's own renames/timestamps." -Level INFO
                                }
                                catch {
                                    Write-OMMigrateLog -Message "Failed to reload rules_inventory.csv after Invoke-DeployConsolidatedRules -- Phase 3 will use the pre-consolidation snapshot: $_" -Level WARN
                                }
                            }

                            # CHANGED 2026-06-26 (Administrator): wire deployResult into the existing
                            # report-feeding variables. Previously $deployResult was only
                            # logged to console/log file and never reached the HTML
                            # Migration Report -- $Script:RulesUpdated and
                            # $Script:RulesStoreSummary are what the report generator
                            # actually reads, and neither was being updated here.
                            if ($deployResult.Created -gt 0 -or $deployResult.Failed -gt 0) {
                                $Script:RulesUpdated += $deployResult.Created
                                $Script:RulesStoreSummary.Add([PSCustomObject]@{
                                    StoreName             = 'Invoke-DeployConsolidatedRules (all accounts)'
                                    Created               = $deployResult.Created
                                    Skipped               = 0
                                    Failed                = $deployResult.Failed
                                    StopProcessingSet     = $deployResult.Created
                                    StopProcessingFailed   = 0
                                })
                            }

                            # Stale-collection-overwrite fix (added 2026-07-01, Administrator,
                            # per Gemini review): Outlook's rules persistence is a
                            # "last Save() wins" model at the whole-collection level --
                            # confirmed live. Invoke-DeployConsolidatedRules correctly
                            # deletes and saves via its OWN rules collection reference,
                            # but Strategy 1 (below, "Processing rules across N store(s)")
                            # reads from $liveCollections, a SEPARATE reference fetched
                            # earlier in this run, before the consolidation deletes
                            # happened. Even though Strategy 1 now correctly skips
                            # remapping already-consolidated rules by name, it still
                            # calls its own Rules Save() at the end of its per-store
                            # loop -- and that Save() serializes its own stale in-memory
                            # snapshot (which still contains "TestProfile ACE") back to
                            # PR_RW_RULES_STREAM, silently reverting the earlier,
                            # correct deletion. Force-release every store reference
                            # currently cached in $liveCollections and let Strategy 1
                            # re-fetch fresh via GetRules() the next time it's read, so
                            # its own Save() operates on the current, post-consolidation
                            # state instead of an outdated one. Isolated in its own
                            # scriptblock invocation to avoid any scope interaction
                            # with $deployResult, matching the pattern used to fix an
                            # earlier scoping bug in this same block today.
                            & {
                                try {
                                    $storeNamesToPurge = @($liveCollections.Keys)
                                    foreach ($purgeStoreName in $storeNamesToPurge) {
                                        try {
                                            $staleRef = $liveCollections[$purgeStoreName]
                                            if ($staleRef) {
                                                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($staleRef)
                                            }
                                            # [void] added (2026-07-06, Administrator) -- Hashtable.Remove()
                                            # returns a bool (whether the key existed), which was
                                            # previously unsuppressed and leaked to the console as a
                                            # bare "True" line after every purge, confirmed live via
                                            # Administrator's DEBUG-level test run. Same pattern as
                                            # FinalReleaseComObject's [void] cast two lines above --
                                            # this is a return-value suppression fix, not a logging
                                            # or LogLevel issue; it would leak at any log level.
                                            [void]$liveCollections.Remove($purgeStoreName)
                                            Write-OMMigrateLog -Message "liveCollections: purged stale reference for '$purgeStoreName' after rule consolidation -- will re-fetch fresh." -Level DEBUG
                                        }
                                        catch {
                                            Write-OMMigrateLog -Message "liveCollections: failed to purge stale reference for '$purgeStoreName': $_" -Level WARN
                                        }
                                    }
                                    [GC]::Collect()
                                    [GC]::WaitForPendingFinalizers()

                                    # Re-fetch fresh references for the same store names so
                                    # Strategy 1's existing lookup ($liveCollections[$activeStoreName])
                                    # still finds a valid collection -- just a current one now,
                                    # not the purged stale one.
                                    foreach ($refetchStoreName in $storeNamesToPurge) {
                                        try {
                                            $refetchStore = $null
                                            for ($rfi = 1; $rfi -le $allStores.Count; $rfi++) {
                                                $candidateRefetchStore = $allStores.Item($rfi)
                                                $candidateRefetchName = ''
                                                try { $candidateRefetchName = $candidateRefetchStore.DisplayName } catch { }
                                                if ($candidateRefetchName -eq $refetchStoreName) {
                                                    $refetchStore = $candidateRefetchStore
                                                    break
                                                }
                                            }
                                            if ($refetchStore) {
                                                $liveCollections[$refetchStoreName] = $refetchStore.GetRules()
                                            }
                                        }
                                        catch {
                                            Write-OMMigrateLog -Message "liveCollections: failed to re-fetch fresh reference for '$refetchStoreName': $_" -Level WARN
                                        }
                                    }
                                }
                                catch {
                                    Write-OMMigrateLog -Message "liveCollections purge/refetch pass failed: $_" -Level WARN
                                }
                            }
                        }
                        catch {
                            Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules failed (non-fatal): $_" -Level WARN
                            Write-Host "  WARNING: Consolidated rules deployment failed -- see log for details." -ForegroundColor Yellow
                        }
                    }
                    else {
                        Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules skipped -- Archive PST root not available." -Level WARN
                    }
                }
                catch {
                    Write-OMMigrateLog -Message "Failed to scan Outlook Rules across stores: $_" `
                                       -Level ERROR
                    Write-Host '  ERROR: Could not access Outlook Rules.' -ForegroundColor Red
                    Write-Host '  Rules update skipped -- update rules manually in Outlook.' `
                               -ForegroundColor Yellow
                }

                if ($rulesCollection) {

                    # --------------------------------------------------------
                    #  Gemini Strategy 1 + 2 Combined -- Per-Store Loop
                    #  Process each store's rules collection independently.
                    #  Strategy 1: Remap known rules to correct folder target.
                    #  Strategy 2: Disable MoveToFolder on all other rules
                    #              with stale targets so Save() passes validation.
                    #  ONE Save() per store after all its rules are processed.
                    # --------------------------------------------------------

                    # Determine which stores to process -- ONLY the picker-selected
                    # account(s) in $accountsToProcess. CHANGED 2026-06-26 (explicitly
                    # authorized by Administrator): this loop previously also added every store
                    # in $liveCollections regardless of picker selection, so selecting
                    # one account in the picker still processed every account's rules
                    # in the entire profile (confirmed live: picker selected 1 account,
                    # log showed "Processing rules across 23 store(s)"). Administrator wants this
                    # scoped to the picker selection going forward.
                    #
                    # ADDED (Administrator direction, 2026-07-20): Exchange accounts
                    # (ProviderTag EXCHANGE-*, e.g. EXCHANGE-SKIP) are always excluded from
                    # $accountsForFolderWork/$accountsToProcess further up this file --
                    # correctly so, since they never go through Script 02 conversion and have
                    # no folder_map.csv entries (content migration already safely no-ops for
                    # them via the "No folder map entries" check in the item-migration loop
                    # above). But Administrator confirmed rules WORK still matters for these
                    # accounts -- he edits their rules directly in Outlook's Rules Manager UI
                    # and expects LastTargetRun to be tracked/updated same as any other
                    # account's rules. Sourced independently, directly from $allAccounts, so
                    # this does not touch $accountsForFolderWork, $accountsToProcess, the
                    # picker's candidate list, or content-migration eligibility -- it only
                    # widens which stores Phase 3 (this rules-remap loop) will process.
                    $exchangeAccountsForRules = @($allAccounts | Where-Object {
                        $_.ProviderTag -like 'EXCHANGE-*'
                    })

                    $storesToProcess = [System.Collections.Generic.List[string]]::new()
                    foreach ($acct in $accountsToProcess) {
                        $acctEmail = $acct.EmailAddress
                        if ($acctEmail -and $liveCollections.ContainsKey($acctEmail)) {
                            if (-not $storesToProcess.Contains($acctEmail)) {
                                $storesToProcess.Add($acctEmail)
                            }
                        }
                    }
                    foreach ($exAcct in $exchangeAccountsForRules) {
                        $exAcctEmail = $exAcct.EmailAddress
                        if ($exAcctEmail -and $liveCollections.ContainsKey($exAcctEmail)) {
                            if (-not $storesToProcess.Contains($exAcctEmail)) {
                                $storesToProcess.Add($exAcctEmail)
                            }
                        }
                    }
                    # FIXED (Administrator direction, 2026-07-21, corrected from an initial
                    # too-broad "always include the default store" attempt): mirrors the
                    # Exchange-accounts fix immediately above exactly, but scoped by
                    # ProviderTag rather than which store happens to be the default/
                    # primary store. An account still tagged POP3-* (not yet converted to
                    # IMAP, regardless of MigrationAction or whether it's selected in this
                    # run's picker) gets its rules processed the same way an Exchange
                    # account does -- rule maintenance is independent of migration state.
                    # Explicitly NOT "always include the default store" -- Administrator was
                    # clear this must stay scoped to accounts genuinely still in a POP3
                    # (or Exchange) state, not broadened to every account regardless of
                    # picker selection; the 2026-06-26 scoping decision for converted/
                    # already-IMAP secondary accounts is unaffected.
                    $pop3AccountsForRules = @($allAccounts | Where-Object {
                        $_.ProviderTag -like 'POP3-*'
                    })
                    foreach ($pAcct in $pop3AccountsForRules) {
                        $pAcctEmail = $pAcct.EmailAddress
                        if ($pAcctEmail -and $liveCollections.ContainsKey($pAcctEmail)) {
                            if (-not $storesToProcess.Contains($pAcctEmail)) {
                                $storesToProcess.Add($pAcctEmail)
                            }
                        }
                    }
                    # Original behavior (added every remaining live-collection store
                    # regardless of picker selection) -- disabled per Administrator's instruction
                    # above. Left in place, commented out, rather than deleted, in case
                    # this needs to be reverted.
                    # foreach ($key in $liveCollections.Keys) {
                    #     if (-not $storesToProcess.Contains($key)) {
                    #         $storesToProcess.Add($key)
                    #     }
                    # }

                    if ($storesToProcess.Count -eq 0) {
                        # CHANGED 2026-06-26 (Administrator): distinguish the expected, harmless
                        # case from a genuine problem, rather than always logging WARN.
                        # Expected case: every picker-selected account is a secondary
                        # store with CSV rules, already deliberately excluded from
                        # $liveCollections by the early skip (to avoid UI bleed into
                        # ameritech -- see $earlySecondaryStoreNames above). That account
                        # is correctly handled by Invoke-DeployConsolidatedRules instead,
                        # which already ran earlier in this same block. Nothing is wrong
                        # in that case -- there is simply nothing left for THIS loop
                        # (the Strategy 1+2 folder-target remap loop) to do.
                        $allSelectedAreSecondarySkipped = $true
                        foreach ($acct in $accountsToProcess) {
                            # FIXED (Administrator direction, 2026-07-20): .Trim() added on this
                            # side too, matching the build-side fix on $earlySecondaryStoreNames
                            # above -- see that fix comment for full context.
                            if ($acct.EmailAddress -and
                                -not $earlySecondaryStoreNames.Contains($acct.EmailAddress.Trim())) {
                                $allSelectedAreSecondarySkipped = $false
                                break
                            }
                        }
                        # ADDED (Administrator direction, 2026-07-20): second harmless case --
                        # the default/primary store simply was not selected in the picker this
                        # run (confirmed live: "Live rules scan skipped -- default store
                        # 'account@example-provider.com' not in picker selection" earlier in the
                        # same run's log). This is a deliberate operator choice, not a problem,
                        # and there is no admin action to take -- same spirit as the secondary-
                        # store case above, just for the primary account instead.
                        $defaultStoreNotSelected = (
                            $earlyDefaultStoreName -and
                            -not (@($accountsToProcess | ForEach-Object { $_.EmailAddress }) -contains $earlyDefaultStoreName)
                        )
                        if (($allSelectedAreSecondarySkipped -or $defaultStoreNotSelected) -and $accountsToProcess.Count -gt 0) {
                            Write-OMMigrateLog -Message "No folder-target remap needed for this run's selection -- default account was not selected in the picker, or all selected account(s) are secondary stores already handled by Invoke-DeployConsolidatedRules." -Level INFO
                        } else {
                            Write-OMMigrateLog -Message "No live rules collections found -- skipping rules update." -Level WARN
                        }
                    }
                    else {
                    Write-OMMigrateLog -Message "Processing rules across $($storesToProcess.Count) store(s)." -Level INFO

                    foreach ($activeStoreName in $storesToProcess) {
                        $activeCollection = $liveCollections[$activeStoreName]
                        if (-not $activeCollection) { continue }

                        Write-OMMigrateLog -Message "Using rules collection for store: '$activeStoreName'" -Level DEBUG
                        Write-Host "  Store: $(Invoke-OMMigrateSanitize -Text $activeStoreName)" -ForegroundColor Cyan

                    # Wrap entire store processing in try/catch -- $ErrorActionPreference
                    # is Stop so any unhandled COM exception would be fatal without this.
                    try {

                    # Local list for disabled rules in this store -- merged into
                    # $Script:DisabledRules after the store loop to avoid scope issues.
                    $storeDisabledRules = [System.Collections.Generic.List[PSCustomObject]]::new()

                    # Phase 1: Structural integrity scan -- remove truly corrupt rules.
                    $brokenRemoved = 0
                    $colRuleCount  = 0
                    try { $colRuleCount = $activeCollection.Count } catch { 
                        Write-OMMigrateLog -Message "Could not get rule count for '$activeStoreName' -- skipping store." -Level INFO
                        continue
                    }
                    for ($bri = $colRuleCount; $bri -ge 1; $bri--) {
                        # LIFTED MANDATE 2026-06-20 (one-time, this line only, explicitly
                        # authorized by Administrator): same skip-and-continue pattern as Phase 4a --
                        # Item() itself can throw 0x800C8101; skip this index and continue
                        # the scan rather than letting it abort the whole store's Phase 1.
                        try { $br = $activeCollection.Item($bri) } catch { continue }
                        $brName    = ''
                        $brBroken  = $false
                        try { $brName = $br.Name } catch { $brBroken = $true }
                        if (-not $brBroken) {
                            try { $brAct = $br.Actions; $brMove = $brAct.MoveToFolder }
                            catch { $brBroken = $true }
                        }
                        if ($brBroken) {
                            Write-OMMigrateLog -Message "Removing truly corrupt rule[$bri]: '$brName'" -Level WARN
                            try { $activeCollection.Remove($brName); $brokenRemoved++ } catch { }
                        }
                    }
                    if ($brokenRemoved -gt 0) {
                        # SAFETY GUARD added 2026-06-19, REMOVED 2026-06-19 (same night):
                        # the guard was based on a misreading of that session's own test
                        # results. A diagnostic test script covering Phase 1 + Phase 1.5 +
                        # Phase 4a + this same Save() call was run live against "TestProfile"'s
                        # 573-rule ameritech ruleset: Save() succeeded, and -- confirmed via
                        # Outlook Rules and Alerts UI screenshots across a full Outlook
                        # close/reopen -- the sorted order PERSISTED. The post-Save Item()
                        # readback failing in that same script is a separate, cosmetic
                        # COM enumeration quirk that does not reflect the actual saved
                        # state; the UI and mail-flow are the ground truth and both are
                        # correct after this Save(). See Memory #28 (closes #21/#23/#25/#26).
                        Write-OMMigrateLog -Message "Corrupt rules removed: $brokenRemoved. Saving clean collection..." -Level INFO
                        try {
                            $activeCollection.GetType().InvokeMember("Save", [System.Reflection.BindingFlags]::InvokeMethod, $null, $activeCollection, @($true))
                            Write-OMMigrateLog -Message "Clean collection saved after removing $brokenRemoved corrupt rules." -Level INFO
                        }
                        catch {
                            Write-OMMigrateLog -Message "Save() failed after corrupt rule removal: $_" -Level WARN
                        }
                    }

                    # REMOVED 2026-07-02 (Administrator): Phase 1.5 (added 2026-06-19) used to
                    # re-resolve each rule's Account condition and force-enable it on
                    # every default-store rule, every run. That was correct for its
                    # original purpose (repairing .rwz-imported rules whose Account
                    # condition Object reference didn't survive import), but it
                    # directly conflicts with the 2026-07-01 finding that the Account
                    # condition itself -- even fully valid and resolved -- triggers
                    # 0x800C8101 on Item() access at this rule count. Confirmed live
                    # 2026-07-02: rules freshly rebuilt by Module3.bas/Invoke-
                    # DeployConsolidatedRules with NO Account condition were still
                    # showing the condition in the Rules and Alerts UI after this
                    # script ran -- traced to this block silently force-enabling it
                    # again on every pass. Removed entirely per Administrator's explicit
                    # instruction rather than left disabled-in-place, since Administrator
                    # identified obsolete-but-still-called code as the actual
                    # mechanism of this regression and wants it gone, not dormant.

                    # Phase 2: Build CSV lookup from ALL rules_inventory rows.
                    # No RuleStoreName filter -- all rules live in the single master
                    # collection regardless of which account they fire for. RuleStoreName
                    # is used only to stamp the account condition, not to find the rule.
                    $csvRuleMap = [System.Collections.Generic.Dictionary[string,object]]::new(
                        [System.StringComparer]::OrdinalIgnoreCase
                    )
                    foreach ($ruleRow in $rulesNeedingUpdate) {
                        if (-not $csvRuleMap.ContainsKey($ruleRow.RuleName)) {
                            $csvRuleMap[$ruleRow.RuleName] = $ruleRow
                        }
                    }

                    # ADDED (Administrator direction, 2026-08-18 -- multi-run reprocessing bug fix):
                    # $csvRuleMap above is keyed by RuleName ALONE, but RuleName is not a
                    # unique identifier across the inventory -- the same RuleName can
                    # legitimately occur under more than one RuleStoreName (confirmed by
                    # Administrator: multiple RuleNames can also share a single TargetFolderPath,
                    # so TargetFolderPath is not a safe substitute key either). When two
                    # rows share a RuleName, the plain-RuleName dictionary above silently
                    # keeps only the FIRST one seen -- the second row's real data is never
                    # reachable via $csvRuleMap[$ruleName], even though ContainsKey($ruleName)
                    # still reports true for it. $csvRuleMapByStoreAndName below composite-
                    # keys on RuleStoreName+RuleName (matching the composite-key pattern
                    # already used successfully elsewhere in this project, e.g. Export-
                    # RulesToCSV's merge logic and Step 4's write-back in Invoke-
                    # DeployConsolidatedRules) so same-named rules in different stores no
                    # longer collide.
                    $csvRuleMapByStoreAndName = [System.Collections.Generic.Dictionary[string,object]]::new(
                        [System.StringComparer]::OrdinalIgnoreCase
                    )
                    foreach ($ruleRowForComposite in $rulesNeedingUpdate) {
                        $compositeStoreNameKey = "$($ruleRowForComposite.RuleStoreName)|$($ruleRowForComposite.RuleName)"
                        if (-not $csvRuleMapByStoreAndName.ContainsKey($compositeStoreNameKey)) {
                            $csvRuleMapByStoreAndName[$compositeStoreNameKey] = $ruleRowForComposite
                        }
                    }

                    # PENDING-ROW GATE (added 2026-07-07, Administrator direction): Phase 3 below
                    # used to process EVERY rule in the live collection whose name matched
                    # a CSV row with NeedsFolderUpdate=True, with no idempotency check at
                    # all -- meaning a rule that had ALREADY been correctly folder-remapped
                    # in a prior run still got fully re-processed (including a SendersDomain
                    # condition rewrite via Set-RuleConditions, see below) on every single
                    # subsequent run. Confirmed live as the actual mechanism behind ComEd
                    # and GlenbardWest73's SendersDomain conditions being corrupted despite
                    # never having been "pending" in any prior sense.
                    #
                    # Administrator's explicit DIRECTION: use a NEW, separate column -- LastTargetRun
                    # -- not LastDeployedRun (which is reserved for Invoke-DeployConsolidat-
                    # edRules's own consolidation/condition-validation idempotency and runs
                    # BEFORE Phase 3 in the same script execution; gating Phase 3 on that
                    # same column would cause a rule freshly consolidated this run to be
                    # incorrectly skipped by Phase 3 in the SAME run, since its
                    # LastDeployedRun would already be stamped by the time Phase 3 reads
                    # it). LastTargetRun tracks Script 03 Phase 3's folder-target
                    # (TargetFolderPath) remap work specifically, independently blank-gated.
                    # On a genuine first run every row is blank, so this naturally processes
                    # everything, same as today's behavior -- this only changes behavior on
                    # SUBSEQUENT runs, where already-remapped (non-blank LastTargetRun) rows
                    # are now correctly left alone.
                    $phase3PendingNames = [System.Collections.Generic.HashSet[string]]::new(
                        [System.StringComparer]::OrdinalIgnoreCase
                    )
                    foreach ($p3Row in $rulesInventory) {
                        $p3Timestamp = if ($p3Row.PSObject.Properties['LastTargetRun']) { [string]$p3Row.LastTargetRun } else { '' }
                        $p3RuleName  = if ($p3Row.PSObject.Properties['RuleName'])       { [string]$p3Row.RuleName }       else { '' }
                        if ([string]::IsNullOrWhiteSpace($p3Timestamp) -and -not [string]::IsNullOrWhiteSpace($p3RuleName)) {
                            [void]$phase3PendingNames.Add($p3RuleName)
                        }
                    }

                    # ADDED (Administrator direction, 2026-08-18 -- multi-run reprocessing bug fix):
                    # same collision problem as $csvRuleMap above, applied to the pending-
                    # row gate itself. $phase3PendingNames is a plain HashSet[string] keyed
                    # by RuleName alone -- .Add() on a HashSet silently no-ops on a
                    # duplicate value, so when two rows share a RuleName across different
                    # RuleStoreNames, only ONE of them actually gets registered as pending.
                    # The other row's live rule still has a blank LastTargetRun and is
                    # genuinely due for remap, but $phase3PendingNames.Contains($ruleName)
                    # cannot distinguish which store's row it belongs to, so results are
                    # inconsistent from run to run depending on CSV row order. Composite-
                    # keyed on RuleStoreName+RuleName, matching $csvRuleMapByStoreAndName
                    # above.
                    $phase3PendingNamesByStore = [System.Collections.Generic.HashSet[string]]::new(
                        [System.StringComparer]::OrdinalIgnoreCase
                    )
                    foreach ($p3RowForComposite in $rulesInventory) {
                        $p3TimestampForComposite = if ($p3RowForComposite.PSObject.Properties['LastTargetRun']) { [string]$p3RowForComposite.LastTargetRun } else { '' }
                        $p3RuleNameForComposite  = if ($p3RowForComposite.PSObject.Properties['RuleName'])       { [string]$p3RowForComposite.RuleName }       else { '' }
                        $p3StoreNameForComposite = if ($p3RowForComposite.PSObject.Properties['RuleStoreName'])  { [string]$p3RowForComposite.RuleStoreName }  else { '' }
                        if ([string]::IsNullOrWhiteSpace($p3TimestampForComposite) -and -not [string]::IsNullOrWhiteSpace($p3RuleNameForComposite)) {
                            [void]$phase3PendingNamesByStore.Add("$p3StoreNameForComposite|$p3RuleNameForComposite")
                        }
                    }

                    # Phase 3: Loop ALL rules in this store's collection, but only ACT on
                    # ones whose CSV row has a blank LastTargetRun (see gate above). Rules
                    # already folder-remapped in a prior run are left completely untouched
                    # -- no remap, no disable, no condition rewrite.
                    # Strategy 1: CSV match + folder found -> remap + enable
                    # Strategy 2: no CSV match or folder not found -> disable
                    Write-Host "  Processing $($activeCollection.Count) rules..." -ForegroundColor Cyan
                    $remapped             = 0
                    $disabledCount        = 0
                    $skippedCount         = 0
                    $alreadyTargetedCount = 0
                    # Collects rule names successfully remapped this run, so LastTargetRun
                    # can be written back to the CSV once, after Save() confirms the remap
                    # actually committed (see Phase 4 below).
                    $newlyTargetedRuleNames = [System.Collections.Generic.List[string]]::new()

                    for ($ri = 1; $ri -le $activeCollection.Count; $ri++) {
                        # LIFTED MANDATE 2026-06-20 (one-time, this line only, explicitly
                        # authorized by Administrator): same skip-and-continue pattern as Phase 1
                        # and Phase 4a -- Item() itself can throw 0x800C8101; skip this
                        # index and continue the remap rather than aborting the whole
                        # store's Phase 3.
                        try { $liveRule = $activeCollection.Item($ri) } catch { continue }
                        $ruleName   = ''
                        try { $ruleName = $liveRule.Name } catch { continue }

                        # PENDING-ROW GATE: skip entirely if this rule's CSV row already has
                        # a real LastTargetRun timestamp -- its folder target was already
                        # remapped in a prior run and must not be re-processed. Rules with NO
                        # CSV row at all still fall through to the existing "no CSV entry"
                        # handling further below, unchanged.
                        # FIXED (Administrator direction, 2026-08-18 -- multi-run reprocessing bug):
                        # was $phase3PendingNames.Contains($ruleName) -- plain-RuleName gate,
                        # see $phase3PendingNamesByStore's declaration above for why that
                        # silently mis-gated rows when a RuleName repeats across stores.
                        # Composite-keyed check below is otherwise identical in behavior.
                        if (-not $phase3PendingNamesByStore.Contains("$activeStoreName|$ruleName")) {
                            $isKnownButTargeted = $rulesInventory | Where-Object {
                                $_.RuleStoreName -eq $activeStoreName -and
                                $_.RuleName      -eq $ruleName
                            } | Select-Object -First 1
                            if ($isKnownButTargeted) {
                                $alreadyTargetedCount++
                                Write-OMMigrateLog -Message "Rule '$ruleName' already folder-targeted (LastTargetRun set) -- skipped, not re-processed." -Level DEBUG
                                continue
                            }
                        }

                        $actions    = $liveRule.Actions
                        $moveAction = $actions.MoveToFolder

                        # FIXED (bug found live 2026-07-12, this same fix): resolve this
                        # rule's CURRENT live target folder path here, once, and use it for
                        # BOTH the already-consolidated check and the CSV lookup below,
                        # instead of matching either one by $ruleName. A rule that was just
                        # renamed by consolidation earlier in THIS SAME RUN (raw UI name,
                        # e.g. 'Ecobee', standardized to 'Rule: [account] ecobee (Part 1)')
                        # has a $ruleName that no longer matches the name recorded anywhere
                        # under its old identity -- $ruleName-based lookups silently miss
                        # it and it falls through to the "unknown rule" disable path even
                        # though it is fully known and correctly targeted. Per Administrator's
                        # standing, repeated direction to replicate the VBA macro's own
                        # logic (Module3.bas BuildRulesFromMap / WriteTimestampsToCSV, which
                        # have never matched by rule name -- only by RuleStoreName + folder
                        # path), this resolves the rule's live folder path via the existing
                        # Get-FolderFullPath helper (already used elsewhere for this exact
                        # purpose) and strips the store prefix, giving a path stable across
                        # any rename. Falls back to blank on any resolution failure -- both
                        # consumers below already handle a non-matching path exactly as they
                        # handled a non-matching name before this fix, so no new failure mode
                        # is introduced.
                        $liveRuleFolderPath = ''
                        try {
                            if ($moveAction -and $moveAction.Enabled -and $moveAction.Folder) {
                                $liveFullPath = Get-FolderFullPath -Folder $moveAction.Folder
                                if ($liveFullPath) {
                                    $liveStorePrefix = "$activeStoreName\"
                                    if ($liveFullPath.StartsWith($liveStorePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                                        $liveRuleFolderPath = $liveFullPath.Substring($liveStorePrefix.Length)
                                    }
                                    else {
                                        foreach ($prefixStrip in $Namespace.Stores) {
                                            $stripPrefix = "$($prefixStrip.DisplayName)\"
                                            if ($liveFullPath.StartsWith($stripPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                                                $liveRuleFolderPath = $liveFullPath.Substring($stripPrefix.Length)
                                                break
                                            }
                                        }
                                    }
                                }
                            }
                        } catch { }

                        # FIXED (bug found live 2026-07-12, this same fix): checks
                        # $liveRuleFolderPath against ConsolidatedRuleTargets's VALUES
                        # (folder paths) instead of ConsolidatedRuleNames' entries (old
                        # rule names) against $ruleName. Same rationale as above.
                        # FIXED (Administrator direction, this same fix): no longer skips the
                        # actual folder-set logic below just because the live path matched
                        # a consolidated target. Confirmed live that a rule can report a
                        # matching path here while its MoveToFolder action is still not
                        # genuinely applied in Outlook -- skipping left it broken. Now only
                        # marks this rule as already-consolidated (for logging/tracking) and
                        # falls through to the normal remap path below, which does the real
                        # work: reads TargetFolderPath, resolves the folder, sets MoveToFolder.
                        $alreadyConsolidatedByPath = $false
                        if (-not [string]::IsNullOrWhiteSpace($liveRuleFolderPath) -and
                            $deployResult -and
                            $deployResult.PSObject.Properties['ConsolidatedRuleTargets'] -and
                            $deployResult.ConsolidatedRuleTargets.ContainsValue($liveRuleFolderPath)) {
                            $alreadyConsolidatedByPath = $true
                        }

                        # FIXED (bug found live 2026-07-12, this same fix): CSV lookup now
                        # tries $liveRuleFolderPath first (matches Module3.bas's folder-path
                        # based identity), falling back to the original $ruleName lookup
                        # for any rule whose live folder path could not be resolved above --
                        # preserves exact prior behavior for that case, changes nothing for
                        # rules unaffected by a same-run rename.
                        $csvRowByPath = $null
                        if (-not [string]::IsNullOrWhiteSpace($liveRuleFolderPath)) {
                            $csvRowByPath = $rulesInventory | Where-Object {
                                $_.RuleStoreName -eq $activeStoreName -and
                                $_.TargetFolderPath -eq $liveRuleFolderPath
                            } | Select-Object -First 1
                        }

                        # FIXED (real structural bug found by tracing, not diagnostics):
                        # this was previously its own sibling "if" block that only set
                        # $csvRow and then fell through PAST the entire elseif/else chain
                        # below -- since $csvRowByPath being truthy satisfied THIS if, the
                        # elseif/else (which contain every counter increment and the actual
                        # folder-navigation/remap code) never ran at all. A rule matched by
                        # live folder path got $csvRow set correctly but was never actually
                        # remapped, never counted, and LastTargetRun was never queued for
                        # write-back -- explains every "Remapped=0 Disabled=0 Skipped=0"
                        # result seen tonight for a rule that WAS present, WAS pending, and
                        # WAS matched. Now folded into a single unified condition so a
                        # path-match and a name-match both enter the SAME body that performs
                        # the actual remap.
                        # FIXED (Administrator direction, 2026-08-18 -- multi-run reprocessing bug):
                        # was $csvRuleMap.ContainsKey($ruleName) / $csvRuleMap[$ruleName] --
                        # plain-RuleName lookup, see $csvRuleMapByStoreAndName's declaration
                        # above for why that silently returned the wrong row (or missed
                        # entirely) whenever a RuleName repeats across stores. Composite key
                        # otherwise preserves identical fallback order: path-match first,
                        # then name-match within THIS rule's own store.
                        $csvCompositeKeyForLookup = "$activeStoreName|$ruleName"
                        if ($csvRowByPath -or $csvRuleMapByStoreAndName.ContainsKey($csvCompositeKeyForLookup)) {
                            if ($csvRowByPath) {
                                $csvRow = $csvRowByPath
                                Write-OMMigrateLog -Message "Rule '$ruleName' matched CSV row by live folder path ('$liveRuleFolderPath') instead of name -- name mismatch likely from same-run consolidation rename." -Level DEBUG
                            }
                            else {
                                $csvRow = $csvRuleMapByStoreAndName[$csvCompositeKeyForLookup]
                            }

                            # Skip already-consolidated rules (added 2026-07-01, Administrator):
                            # $deployResult.ConsolidatedRuleNames lists every pre-existing
                            # rule already deleted by Invoke-DeployConsolidatedRules's
                            # consolidation scan this run -- if $ruleName is in that list,
                            # the rule no longer exists in Outlook (it was deleted and
                            # replaced under a new standardized name), so Strategy 1's
                            # remap-by-name below has nothing valid left to act on. Without
                            # this, Strategy 1 still "sees" the rule via its own separately-
                            # fetched $activeCollection reference and remaps it again,
                            # leaving it standing alongside the newly-created consolidated
                            # rule (confirmed live, 2026-07-01: "TestProfile ACE"). $deployResult
                            # may not exist at all if the consolidation block above was
                            # skipped or failed -- guard every property access accordingly.
                            #
                            # FIXED (2026-07-09, Administrator direction): the skip above correctly
                            # stopped Strategy 1 from touching a stale/deleted COM reference,
                            # but had the side effect of leaving LastTargetRun blank for every
                            # rule consolidation touched this run -- silently deferring that
                            # work to a SECOND Script 03 run, with nothing telling the operator
                            # anything was incomplete. Confirmed live 2026-07-09 on "TestProfile":
                            # 360/361 rules consolidated in one run, only 4 got LastTargetRun
                            # stamped, 356 silently pending. Administrator's explicit design intent
                            # (2026-07-07, LastDeployedRun/LastTargetRun split) was that BOTH
                            # jobs complete in a SINGLE run when both columns start blank --
                            # not across two runs. Fix: if this rule's name is in
                            # $deployResult.ConsolidatedRuleTargets (populated by
                            # Invoke-BuildRulesFromMap at the moment it successfully creates
                            # and saves the new rule, using the TargetFolderPath it already
                            # resolved and set on that rule's MoveToFolder action), the rule
                            # is ALREADY correctly targeted -- no further COM action needed,
                            # no stale-reference risk (we never touch the live object again),
                            # just add it straight to $newlyTargetedRuleNames so it rides the
                            # existing post-Save() LastTargetRun write-back below, in this
                            # same run.
                            $alreadyConsolidated = $false
                            $consolidatedTargetPath = $null
                            # FIXED (Administrator direction, 2026-07-21): ConsolidatedRuleTargets is now
                            # checked independently, not only as a secondary lookup gated behind
                            # ConsolidatedRuleNames succeeding first. Root cause: ConsolidatedRuleNames
                            # only ever receives the name of an OLD rule absorbed/deleted during
                            # consolidation (see OMMigrate-Outlook.psm1's Invoke-BuildRulesFromMap,
                            # $Result.ConsolidatedRuleNames.Add($candidateRuleName) -- absorption path
                            # only) -- it never receives a newly-created chunk's own name.
                            # ConsolidatedRuleTargets, by contrast, IS correctly populated with every
                            # newly-created chunk's name (same function, $Result.ConsolidatedRuleTargets
                            # [$chunkName] = $TargetPath, set unconditionally on every rule creation).
                            # A genuinely new Part-N rule with no prior live rule to absorb from (the
                            # exact scenario when both LastDeployedRun and LastTargetRun start blank)
                            # never satisfies the old ConsolidatedRuleNames-first check, so this
                            # ConsolidatedRuleTargets lookup was unreachable dead code for that case --
                            # LastTargetRun stayed blank for every such rule, silently deferring to a
                            # second run that Administrator's original design (2026-07-07,
                            # LastDeployedRun/LastTargetRun split) explicitly says should never be
                            # required when both columns start blank.
                            if ($deployResult -and
                                $deployResult.PSObject.Properties['ConsolidatedRuleTargets'] -and
                                $deployResult.ConsolidatedRuleTargets.ContainsKey($ruleName)) {
                                $alreadyConsolidated    = $true
                                $consolidatedTargetPath = $deployResult.ConsolidatedRuleTargets[$ruleName]
                            }
                            elseif ($deployResult -and
                                $deployResult.PSObject.Properties['ConsolidatedRuleNames'] -and
                                $deployResult.ConsolidatedRuleNames.Contains($ruleName)) {
                                $alreadyConsolidated = $true
                                if ($deployResult.PSObject.Properties['ConsolidatedRuleTargets'] -and
                                    $deployResult.ConsolidatedRuleTargets.ContainsKey($ruleName)) {
                                    $consolidatedTargetPath = $deployResult.ConsolidatedRuleTargets[$ruleName]
                                }
                            }
                            if ($alreadyConsolidated) {
                                if (-not [string]::IsNullOrWhiteSpace($consolidatedTargetPath)) {
                                    # ADDED (Administrator direction -- Rules Updated double-count fix,
                                    # follow-up): this branch is Phase 3 recognizing a rule
                                    # that Pass 1 (Invoke-DeployConsolidatedRules) already
                                    # created and correctly targeted this run -- no live COM
                                    # remap action is taken here, but the rule's folder target
                                    # IS finalized in this run, exactly as a real Strategy 1
                                    # remap below (line ~3795) finalizes one. That branch
                                    # increments $Script:RulesUpdated; this one never did,
                                    # even though Administrator's report language ("successfully
                                    # remapped") already counted rules reaching this path.
                                    # Confirmed live (2026-07-13, Six Sigma test case) this
                                    # caused $Script:RulesUpdated to UNDERcount by exactly
                                    # $deployResult.Created every run that creates any rules,
                                    # not specific to the both-dates-blank scenario -- every
                                    # rule Pass 1 creates gets a ConsolidatedRuleTargets entry
                                    # (see OMMigrate-Outlook.psm1 Invoke-BuildRulesFromMap,
                                    # $Result.Created++ immediately followed by
                                    # $Result.ConsolidatedRuleTargets[$chunkName] = $TargetPath,
                                    # same rule, same moment) and is therefore guaranteed to
                                    # reach this exact branch in the same run. Added here so
                                    # this path counts consistently with the real remap branch;
                                    # the double-count adjustment further below (subtracting
                                    # once when both LastDeployedRun and LastTargetRun started
                                    # blank) now nets to the correct total in all three of
                                    # Administrator's stated scenarios.
                                    $Script:RulesUpdated++
                                    $remapped++
                                    [void]$newlyTargetedRuleNames.Add($ruleName)
                                    Write-OMMigrateLog -Message "Strategy 1 remap skipped for '$ruleName' -- already consolidated and correctly targeted this run (TargetFolderPath='$consolidatedTargetPath'); LastTargetRun will be stamped in this same run." -Level INFO
                                }
                                else {
                                    # Consolidated but its resolved target path wasn't captured
                                    # (defensive fallback -- should not normally happen). Same
                                    # behavior as before this fix: skip, leave LastTargetRun
                                    # blank, self-heals on a later run.
                                    Write-OMMigrateLog -Message "Strategy 1 remap skipped for '$ruleName' -- already consolidated and recreated under a new name this run (target path unavailable, will re-process next run)." -Level INFO
                                }
                                $skippedCount++
                                continue
                            }

                            $oldFolderPath = $csvRow.TargetFolderPath
                            $folderMapEntry = $folderMap | Where-Object {
                                $_.FolderPath -eq $oldFolderPath
                            } | Select-Object -First 1

                            if (-not $folderMapEntry) {
                                [void]$moveAction.GetType().InvokeMember('Enabled', [System.Reflection.BindingFlags]::SetProperty, $null, $moveAction, @($false))
                                $disabledCount++
                                Write-OMMigrateLog -Message "No folder map entry for '$ruleName' -- disabled." -Level WARN
                                $storeDisabledRules.Add([PSCustomObject]@{
                                    StoreName = $activeStoreName
                                    RuleName  = $ruleName
                                    Reason    = "No folder map entry for: $oldFolderPath"
                                })
                                continue
                            }

                            $navOK     = $true
                            $curFolder = $null
                            if ($folderMapEntry.Destination -eq 'Server') {
                                $navStore  = $null
                                $navStores = $namespace.Stores
                                for ($nsi = 1; $nsi -le $navStores.Count; $nsi++) {
                                    try {
                                        $nst = $navStores.Item($nsi)
                                        $nstFp = ''; try { $nstFp = $nst.FilePath } catch { }
                                        if ($nst.DisplayName -like "*$activeStoreName*" -and
                                            -not ($nstFp -like '*Backups*') -and
                                            -not ($nstFp -like '*Archive*')) {
                                            $navStore = $nst; break
                                        }
                                    } catch { }
                                }
                                if (-not $navStore) {
                                    $navOK = $false
                                    Write-OMMigrateLog -Message "Server nav: IMAP store not found for '$activeStoreName' -- rule '$ruleName' disabled." -Level WARN
                                }
                                else {
                                    $pathParts = $oldFolderPath.Split('\\')
                                    $curFolder = $namespace.Folders.Item($navStore.DisplayName)
                                    for ($pi = 1; $pi -lt $pathParts.Count; $pi++) {
                                        try { $curFolder = $curFolder.Folders.Item($pathParts[$pi]) }
                                        catch { $navOK = $false; break }
                                    }
                                }
                            }
                            else {
                                # Navigate from Archive PST root -- preserves PST store
                                # context natively so PR_STORE_ENTRYID is correctly set.
                                # Gemini: walk from GetRootFolder() not namespace.Folders.
                                $pathParts = $oldFolderPath.Split('\\')
                                $curFolder = $archiveStore.GetRootFolder()
                                foreach ($part in $pathParts) {
                                    if (-not [string]::IsNullOrWhiteSpace($part)) {
                                        try { $curFolder = $curFolder.Folders.Item($part) }
                                        catch { $navOK = $false; break }
                                    }
                                }
                            }

                            if ($navOK) {
                                # Get StoreID from Archive PST store -- required to set
                                # PR_STORE_ENTRYID correctly and clear red Error labels.
                                # Gemini: GetFolderFromID without StoreID causes MAPI
                                # cross-store identity mismatch and UI validation failure.
                                $archiveStoreID = ''
                                try { $archiveStoreID = $archiveStore.StoreID } catch { }
                                $entryID    = $curFolder.EntryID
                                $mapiFolder = if ($archiveStoreID) {
                                    $namespace.GetFolderFromID($entryID, $archiveStoreID)
                                } else {
                                    $namespace.GetFolderFromID($entryID)
                                }
                                # REMOVED (2026-07-07, Administrator explicit direction): the
                                # Set-RuleConditions call that used to run here has been
                                # removed entirely. Phase 3's job is folder-target (Target-
                                # FolderPath / MoveToFolder) remapping ONLY -- rewriting a
                                # rule's SenderAddress/other conditions is Invoke-
                                # DeployConsolidatedRules's job (via Invoke-BuildRulesFromMap
                                # and its own call to Set-RuleConditions during consolidated
                                # rule creation), which already runs earlier in this same
                                # script, gated independently on LastDeployedRun. Having BOTH
                                # phases write SenderAddress was scope creep in Phase 3 that
                                # caused live condition corruption on every run (confirmed:
                                # ComEd, GlenbardWest73) since Phase 3 previously had no
                                # idempotency gate of its own. Phase 3 now only ever touches
                                # the MoveToFolder action below -- conditions are left exactly
                                # as consolidation set them.
                                [Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction].InvokeMember(
                                    "Folder",
                                    [System.Reflection.BindingFlags]::SetProperty,
                                    $null, $moveAction, $mapiFolder
                                )
                                [void]$moveAction.GetType().InvokeMember('Enabled', [System.Reflection.BindingFlags]::SetProperty, $null, $moveAction, @($true))
                                [void]$liveRule.GetType().InvokeMember('Enabled', [System.Reflection.BindingFlags]::SetProperty, $null, $liveRule, @($true))
                                # Always set StopProcessing on every successfully remapped rule.
                                # Item(27) used instead of Actions.StopProcessingRules named property --
                                # the named property is unreliable via COM interop and returns null
                                # even when Actions.Count confirms the action slot exists. Item(27) is
                                # the confirmed fixed index for ActionType=18 (olRuleActionStopProcessingRules)
                                # in this Outlook COM implementation (confirmed June 17, 2026).
                                try { $spr = $liveRule.Actions.Item(27); [void]$spr.GetType().InvokeMember('Enabled', [System.Reflection.BindingFlags]::SetProperty, $null, $spr, @($true)) } catch { }
                                $Script:RulesUpdated++
                                $remapped++
                                # Track this rule name for LastTargetRun write-back after
                                # Save() confirms the remap actually committed (Phase 4 below).
                                # NOT stamped here -- only once Save() succeeds, since a rule
                                # should not be marked "done" if the store-level Save() fails.
                                [void]$newlyTargetedRuleNames.Add($ruleName)
                                Write-OMMigrateLog -Message "Remapped '$ruleName' -> $oldFolderPath" -Level INFO
                                Write-Host "    [OK]   '$(Invoke-OMMigrateSanitize -Text $ruleName)' -> $($folderMapEntry.Destination)" `
                                           -ForegroundColor Green
                                Write-AuditEntry -Action 'RULE_UPDATED' `
                                                 -Detail "Rule='$ruleName' | OldFolder='$oldFolderPath' | Destination='$($folderMapEntry.Destination)'"
                            }
                            else {
                                [void]$moveAction.GetType().InvokeMember('Enabled', [System.Reflection.BindingFlags]::SetProperty, $null, $moveAction, @($false))
                                $disabledCount++
                                Write-OMMigrateLog -Message "Folder not found for '$ruleName' ($oldFolderPath) -- disabled." -Level WARN
                                Write-Host "    [WARN] Folder not found: '$(Invoke-OMMigrateSanitize -Text $ruleName)'" -ForegroundColor Yellow
                                $storeDisabledRules.Add([PSCustomObject]@{
                                    StoreName = $activeStoreName
                                    RuleName  = $ruleName
                                    Reason    = "Folder not found: $oldFolderPath"
                                })
                            }
                        }
                        else {
                            # No CSV entry in rulesNeedingUpdate for this rule.
                            # Check if it exists in rules_inventory.csv with NeedsFolderUpdate=False
                            # (known rule that doesn't need remapping) -- skip it entirely.
                            $knownRule = $rulesInventory | Where-Object {
                                $_.RuleStoreName -eq $activeStoreName -and
                                $_.RuleName      -eq $ruleName
                            } | Select-Object -First 1

                            if ($knownRule) {
                                # Known rule with NeedsFolderUpdate=False -- no action needed
                                $skippedCount++
                                Write-OMMigrateLog -Message "Rule '$ruleName' in inventory with NeedsFolderUpdate=False -- skipped." -Level DEBUG
                            }
                            elseif ($moveAction.Enabled) {
                                # Unknown rule with active MoveToFolder action -- disable
                                [void]$moveAction.GetType().InvokeMember('Enabled', [System.Reflection.BindingFlags]::SetProperty, $null, $moveAction, @($false))
                                $disabledCount++
                                $storeDisabledRules.Add([PSCustomObject]@{
                                    StoreName = $activeStoreName
                                    RuleName  = $ruleName
                                    Reason    = 'Not in rules_inventory.csv -- update CSV and re-run Script 03'
                                })
                            }
                            else {
                                # No MoveToFolder action -- notification/condition rule, skip entirely
                                $skippedCount++
                                Write-OMMigrateLog -Message "Rule '$ruleName' has no MoveToFolder action -- skipped." -Level DEBUG
                            }
                        }
                    }

                    Write-Host ""
                    Write-OMMigrateLog -Message "Store '$activeStoreName' rules: Remapped=$remapped | Disabled=$disabledCount | Skipped=$skippedCount" -Level INFO

                    # -- Phase 4a (alphabetical sort/renumber via ExecutionOrder) --------
                    # REMOVED 2026-06-26 (explicitly authorized by Administrator): this block tried
                    # to read every rule in a store's live collection via Item(), sort the
                    # readable ones alphabetically, and renumber ExecutionOrder 1..N so
                    # rules display alphabetically in Outlook's Rules and Alerts dialog.
                    #
                    # This is the same Item() enumeration problem documented throughout
                    # this project's history (0x800C8101, confirmed unfixable at scale
                    # after extensive investigation -- see Memory #21/#28) -- this phase
                    # could only ever sort stores where every single rule happened to be
                    # readable via Item() in that session, which in practice meant it
                    # almost never succeeded on ameritech and intermittently failed on
                    # other stores too (confirmed live 2026-06-26: also failed on two
                    # small secondary stores with only 1 and 4 rules).
                    #
                    # Superseded by Invoke-DeployConsolidatedRules (and its helper
                    # Invoke-BuildRulesFromMap) in OMMigrate-Outlook.psm1 -- the PowerShell
                    # port of Gemini's DeployConsolidatedRules VBA macro. That function
                    # achieves the same end goal (rules display alphabetically) through a
                    # completely different mechanism: it controls the ORDER new rules are
                    # CREATED in (reverse iteration over group keys, relying on Outlook's
                    # natural append-order behavior) rather than trying to read back and
                    # re-sort an existing live collection after the fact. Because it never
                    # needs Item() to succeed across an entire collection, it sidesteps
                    # this problem entirely instead of working around it.
                    #
                    # The unrelated Phase 4 Save() immediately below (commits the Strategy
                    # 1+2 folder-target remap from earlier in this same per-store loop) was
                    # NOT touched -- it has nothing to do with sorting and must stay.


                    # Phase 4: ONE Save() per store -- commits all remaps atomically.
                    # SAFETY GUARD added 2026-06-19, REMOVED 2026-06-19 (same night):
                    # the guard was based on a misreading of that session's own test
                    # results. A diagnostic test script covering Phase 1, this same
                    # Phase 1.5 account-condition fix, this same Phase 4a sort, and this
                    # same Save() call was run live against "TestProfile"'s 573-rule
                    # ameritech ruleset and Save() succeeded. Confirmed via Outlook
                    # Rules and Alerts UI screenshots across a full Outlook close and
                    # reopen: the sorted order PERSISTED correctly. The post-Save
                    # Item() readback failing in that test script is a separate,
                    # cosmetic COM enumeration quirk that does not reflect the actual
                    # saved state -- the UI and mail-flow are the ground truth, and
                    # both are correct after this Save(). See Memory #28 (closes
                    # #21/#23/#25/#26 and this guard's own rationale).
                    Write-Host "  Saving rules for $(Invoke-OMMigrateSanitize -Text $activeStoreName)..." -ForegroundColor Cyan
                    try {
                        $activeCollection.GetType().InvokeMember("Save", [System.Reflection.BindingFlags]::InvokeMethod, $null, $activeCollection, @($true))
                        Write-Host "  Rules saved. Remapped: $remapped" -ForegroundColor Green
                        Write-OMMigrateLog -Message "Rules Save() succeeded for '$activeStoreName'. Remapped=$remapped Disabled=$disabledCount" -Level INFO

                        # LastTargetRun WRITE-BACK (added 2026-07-07, Administrator direction): stamp
                        # LastTargetRun for every rule successfully remapped this store, but
                        # ONLY here, after Save() has actually confirmed the remap committed
                        # -- not earlier in the loop. If Save() fails (catch block below),
                        # no rows are stamped and every rule in $newlyTargetedRuleNames
                        # remains pending (blank LastTargetRun) for the next run, which is
                        # the correct, safe behavior. Re-reads rules_inventory.csv fresh from
                        # disk immediately before writing to minimize the chance of clobbering
                        # unrelated fields changed by other stores' Phase 3 passes earlier in
                        # this same script run, then writes back the FULL set (preserving
                        # every other column exactly) via Export-Csv, matching the canonical
                        # column order already established in Export-RulesToCSV.
                        if ($newlyTargetedRuleNames.Count -gt 0) {
                            try {
                                $ltrTimestamp = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.ffffffzzz')
                                $ltrRows      = @(Import-Csv -Path $rulesInventoryPath -Encoding UTF8 |
                                    Where-Object { $_.RuleStoreName -and -not [string]::IsNullOrWhiteSpace($_.RuleStoreName) -and
                                                   $_.RuleName      -and -not [string]::IsNullOrWhiteSpace($_.RuleName) })
                                $ltrNameSet   = [System.Collections.Generic.HashSet[string]]::new(
                                    [string[]]$newlyTargetedRuleNames, [System.StringComparer]::OrdinalIgnoreCase
                                )
                                $ltrStampedCount = 0
                                foreach ($ltrRow in $ltrRows) {
                                    if ($ltrRow.RuleStoreName -eq $activeStoreName -and $ltrNameSet.Contains($ltrRow.RuleName)) {
                                        if (-not $ltrRow.PSObject.Properties['LastTargetRun']) {
                                            Add-Member -InputObject $ltrRow -MemberType NoteProperty -Name 'LastTargetRun' -Value '' -Force
                                        }
                                        $ltrRow.LastTargetRun = $ltrTimestamp
                                        $ltrStampedCount++
                                    }
                                }
                                $ltrCanonicalColumns = @(
                                    'RuleStoreName','TargetStoreName','RuleName','LastDeployedRun','LastTargetRun',
                                    'TargetFolderPath','SendersDomain','NeedsFolderUpdate','IsEnabled',
                                    'ExecutionOrder','RuleType','StopProcessing','Conditions',
                                    'Actions','TargetFolderEntryID','Notes'
                                )
                                # ADDED (Administrator direction, 2026-08-18): re-insert blank separator
                                # rows between RuleStoreName groups before writing -- see
                                # Add-RulesCsvSeparatorRows header comment (OMMigrate-Outlook.psm1)
                                # for full rationale. $ltrRows above is already filtered to real
                                # data rows (see the Where-Object above its Import-Csv), so no
                                # further filtering needed before handing off to the helper.
                                $ltrRowsWithSeparators = Add-RulesCsvSeparatorRows -Rows (@($ltrRows | Select-Object $ltrCanonicalColumns))
                                $ltrRowsWithSeparators | Export-Csv -Path $rulesInventoryPath -NoTypeInformation -Encoding UTF8
                                Write-OMMigrateLog -Message "LastTargetRun written for $ltrStampedCount rule(s) in '$activeStoreName'." -Level INFO
                            }
                            catch {
                                Write-OMMigrateLog -Message "Failed to write LastTargetRun timestamps for '$activeStoreName' (non-fatal, rules will re-process next run): $_" -Level WARN
                            }
                        }
                    }
                    catch {
                        Write-Host "  ERROR: Rules Save() failed for '$activeStoreName': $_" -ForegroundColor Red
                        Write-OMMigrateLog -Message "Rules Save() failed for '$activeStoreName': $_" -Level ERROR
                        if ($Script:FinalStatus -eq 'SUCCESS') { $Script:FinalStatus = 'WARNING' }
                    }

                    # Merge this store's disabled rules into the script-level list
                    foreach ($dr in $storeDisabledRules) {
                        $Script:DisabledRules.Add($dr)
                    }

                    Write-OMMigrateLog -Message "Store '$activeStoreName' disabled rules added to report list: $($storeDisabledRules.Count)" -Level DEBUG

                    } catch {
                        # ORIGINAL (still applies to most stores): non-default stores
                        # don't support independent rule operations via this code path --
                        # their rules were already processed by the secondary-store block
                        # above, so a failure here for a SECONDARY store is expected and
                        # benign.
                        #
                        # ADDED 2026-06-17: this same catch was found to ALSO be silently
                        # absorbing a real, permanent, store-size-inherent COM limitation
                        # for the DEFAULT store (ameritech, 575 rules): Item() indexed
                        # access fails via 0x800C8101 at this scale -- confirmed in prior
                        # sessions as an inherent characteristic of large rule collections
                        # on this store, not data corruption. When that happens, Phase 1
                        # (structural scan), Phase 3 (folder-target remap), and Phase 4a
                        # (sort/renumber) ALL silently no-op for the default store -- it
                        # is not just sort order that fails, folder-target remapping does
                        # too, with no visible error to the operator. Logging this case at
                        # INFO with "expected" framing was actively misleading, since for
                        # the default store this condition is never benign or expected --
                        # it means real per-run work silently did not happen.
                        #
                        # This block does NOT fix the underlying COM limitation -- that
                        # is a known permanent limitation requiring manual remediation via
                        # Outlook's native Rules and Alerts UI (Move Up/Down arrows for
                        # sort order; the UI's own write path for folder-target changes).
                        # This addition only ensures the operator is loudly told it
                        # happened, every run, instead of it being silently logged as INFO
                        # alongside genuinely benign secondary-store skips.
                        if ($activeStoreName -eq $namespace.DefaultStore.DisplayName) {
                            Write-Host ''
                            Write-Host "  *** DEFAULT STORE RULES UPDATE FAILED (known limitation) ***" `
                                       -ForegroundColor Yellow
                            Write-Host "  '$activeStoreName' has too many rules for Outlook's COM Item() access to enumerate." `
                                       -ForegroundColor Yellow
                            Write-Host "  Folder-target remapping and execution-order sorting did NOT run for this store this pass." `
                                       -ForegroundColor Yellow
                            Write-Host "  This is a known, permanent limitation at this rule count -- not a new bug." `
                                       -ForegroundColor Yellow
                            Write-Host "  Manual remediation (if needed): use Outlook's Rules and Alerts UI directly --" `
                                       -ForegroundColor Cyan
                            Write-Host "  Move Up/Down arrows for sort order; edit folder targets there if they're stale." `
                                       -ForegroundColor Cyan
                            Write-OMMigrateLog -Message (
                                "DEFAULT STORE RULES UPDATE FAILED (known permanent limitation, not benign): " +
                                "'$activeStoreName' -- COM Item() access failed at this rule count. " +
                                "Folder-target remap and sort/renumber did NOT run this pass. " +
                                "Manual remediation via Outlook's native Rules and Alerts UI required if changes are needed. " +
                                "Underlying error: $_"
                            ) -Level INFO

                            # -- Additive: roll into script-scope rules-reporting totals --
                            # NOT a failure count -- the default store loads rules via .rwz
                            # import (preserves StopProcessing natively, never goes through
                            # Create()-based purge/recreate). This only records that the COM
                            # Item() ceiling prevented folder-target remap/sort from running
                            # for the default store THIS PASS.
                            $Script:RulesDefaultStoreNotProcessedTotal++

                            # Per-store summary row for the new Rules Processing Detail report table.
                            # Created/Skipped/Failed/StopProcessing* are all left at 0 -- they don't
                            # apply here. NotProcessed=1 flags this row as the distinct "primary
                            # store, not touched this run" case rather than a create/StopProcessing
                            # outcome, so the report can render dedicated wording for it.
                            $Script:RulesStoreSummary.Add([PSCustomObject]@{
                                StoreName             = $activeStoreName
                                Created               = 0
                                Skipped                = 0
                                Failed                = 0
                                StopProcessingSet     = 0
                                StopProcessingFailed  = 0
                                NotProcessed           = $true
                            })
                        } else {
                            # Original behavior, unchanged, for every non-default store.
                            Write-OMMigrateLog -Message "Store '$activeStoreName' skipped (rules live on default store): $_" `
                                               -Level INFO
                        }
                    }

                    } # end foreach store

                    # ADDED (Administrator direction -- Rules Updated double-count fix): apply the
                    # adjustment computed before Pass 1 ran (see $rulesUpdatedDoubleCountAdjustment
                    # above). Applied once here, after ALL stores' Pass 2 remapping is done,
                    # not per-store, since the adjustment already reflects every scoped
                    # account's both-blank rows in one pass-1 count. Never lets the total go
                    # negative (defensive -- should not happen given the counting logic above,
                    # but $Script:RulesUpdated must never be reported as less than 0).
                    if ($rulesUpdatedDoubleCountAdjustment -gt 0) {
                        $Script:RulesUpdated = [Math]::Max(0, $Script:RulesUpdated - $rulesUpdatedDoubleCountAdjustment)
                        Write-OMMigrateLog -Message "Rules Updated double-count adjustment applied: -$rulesUpdatedDoubleCountAdjustment (rules created/consolidated AND folder-remapped in this same run now counted once each)." -Level INFO
                    }

                    Write-Host ""
                    Write-Host "  Rules updated: $Script:RulesUpdated" -ForegroundColor Green
                } # end else storesToProcess
                } # end if ($rulesCollection)
            }
        }
        else {
            Write-Host '  Rules update skipped by operator.' -ForegroundColor DarkGray
            Write-OMMigrateLog -Message 'Rules update skipped by operator.' -Level INFO
        }
    }
    elseif ($false) {
        # Rules are always run in normal operation
    }



    # ----------------------------------------------------------
    #  STEP 6 -- Generate Report and Final Manifest
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Generating Final Report and Manifest' -Step '6 of 6'

    # Determine final status
    if ($Script:FoldersFailed -gt 0 -and $Script:FoldersCreated -eq 0 -and $Script:FoldersVerified -eq 0) {
        $Script:FinalStatus = 'FAILED'
    }
    elseif ($Script:FoldersFailed -gt 0) {
        $Script:FinalStatus = 'WARNING'
    }

    # Generate final Migration report
    # Re-use Migration report type with folder stats in detail
    $allResultsForReport = @($Script:AccountResults | Sort-Object EmailAddress)

    $disabledCsvPath = ''   # Populated below if disabled rules exist -- passed to HTML report
    # Always generate report if rules were updated, disabled, or even just confirmed
    # already-correct (Match path adds a RulesStoreSummary row with Created=0 but
    # never increments RulesUpdated) -- the operator deserves a report either way.
    if ($allResultsForReport.Count -gt 0 -or $Script:RulesUpdated -gt 0 -or
        $Script:RulesStoreSummary.Count -gt 0) {
        $Script:ReportFile = New-MigrationReport `
            -Accounts             $allResultsForReport `
            -RulesUpdated         $Script:RulesUpdated `
            -RulesDisabled        $Script:DisabledRules.Count `
            -DisabledRulesCsvPath $disabledCsvPath `
            -Subtitle             'Folder Migration and Rules Restore Results -- Script 03' `
            -ReportName           'Migration' `
            -RulesSkipped                  $Script:RulesSkippedTotal `
            -RulesNoChange                 $Script:RulesNoChangeTotal `
            -RulesStopProcessingSet        $Script:RulesStopProcessingSetTotal `
            -RulesStopProcessingFailed     $Script:RulesStopProcessingFailedTotal `
            -RulesDefaultStoreNotProcessed $Script:RulesDefaultStoreNotProcessedTotal `
            -RulesStoreSummary             $Script:RulesStoreSummary
        Write-Host "  Migration Report: $Script:ReportFile" -ForegroundColor Green
    }
    else {
        Write-OMMigrateLog -Message 'No account results to report -- skipping HTML report generation.' `
                           -Level INFO
        Write-Host '  No account results to include in HTML report.' -ForegroundColor DarkGray
    }


    # Write disabled rules to CSV and auto-open in Excel if any exist
    if ($Script:DisabledRules.Count -gt 0) {
        $disabledCsvPath = Join-Path $Global:OMMigrate.ReportPath (
            'DisabledRules_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.csv'
        )
        try {
            $Script:DisabledRules | Export-Csv -Path $disabledCsvPath -NoTypeInformation -Encoding UTF8
            Write-Host '' 
            Write-Host "  Disabled rules   : $($Script:DisabledRules.Count) rule(s) need attention." `
                       -ForegroundColor Yellow
            Write-Host "  Disabled rules CSV: $disabledCsvPath" -ForegroundColor Yellow
            Write-OMMigrateLog -Message "Disabled rules CSV written: $disabledCsvPath ($($Script:DisabledRules.Count) rules)" -Level INFO
            if (-not $Script:IsWhatIf) {
                Start-Process $disabledCsvPath
            }
        }
        catch {
            Write-OMMigrateLog -Message "Failed to write disabled rules CSV: $_" -Level WARN
        }
    }

    # Write Step 03 manifest -- migration complete
    Write-StepManifest -Step 3 -Status $Script:FinalStatus -Data @{
        FoldersCreated   = $Script:FoldersCreated
        ItemsCopied      = $Script:ItemsCopied
        FoldersFailed    = $Script:FoldersFailed
        RulesRecreated   = $Script:RulesRecreated
        RulesUpdated     = $Script:RulesUpdated
        ArchivePSTPath   = $Script:ArchivePSTPath
        AccountsProcessed = $Script:AccountResults.Count
        ReportFile       = $Script:ReportFile
        FolderMigrationSkipped = $false
        RulesUpdateSkipped     = $false
        RulesRecreationRun     = $RecreateRules.IsPresent
        # -- Additive rules-reporting fields (manifest record, not used by report generation above) --
        RulesSkipped              = $Script:RulesSkippedTotal
        RulesNoChange              = $Script:RulesNoChangeTotal
        RulesStopProcessingSet     = $Script:RulesStopProcessingSetTotal
        RulesStopProcessingFailed  = $Script:RulesStopProcessingFailedTotal
        RulesDefaultStoreNotProcessed = $Script:RulesDefaultStoreNotProcessedTotal
    }

    Write-Host ''

    # Final console summary
    Write-Host ('=' * 60) -ForegroundColor DarkCyan
    Write-Host '  MIGRATION COMPLETE' -ForegroundColor White
    Write-Host ('=' * 60) -ForegroundColor DarkCyan
    Write-Host ''
    Write-Host "  Folders created  : $Script:FoldersCreated" -ForegroundColor Gray
    Write-Host "  Folders verified : $Script:FoldersVerified" -ForegroundColor Gray
    Write-Host "  Email items copied: $Script:ItemsCopied" -ForegroundColor Gray
    Write-Host "  Rules recreated  : $Script:RulesRecreated" -ForegroundColor Gray
    Write-Host "  Rules updated    : $Script:RulesUpdated" -ForegroundColor Gray
    if ($Script:DisabledRules.Count -gt 0) {
        Write-Host "  Rules disabled   : $($Script:DisabledRules.Count)  (folder not found -- see report)" `
                   -ForegroundColor Yellow
    }
    if ($Script:ArchivePSTPath) {
        Write-Host "  Archive PST      : $(Invoke-OMMigrateSanitize -Text $Script:ArchivePSTPath)" -ForegroundColor Gray
    }
    Write-Host ''

    if ($Script:FinalStatus -eq 'SUCCESS') {
        Write-Host '  All done! Your migration is complete.' -ForegroundColor Green
        Write-Host ''
        Write-Host '  RECOMMENDED NEXT STEPS:' -ForegroundColor Cyan
        Write-Host '  1. Open Outlook and verify all accounts are connected.' `
                   -ForegroundColor Gray
        Write-Host '  2. Check that email is arriving in the correct folders.' `
                   -ForegroundColor Gray
        Write-Host '  3. Verify Outlook Rules are routing new email correctly.' `
                   -ForegroundColor Gray
        Write-Host '  4. Review the Archive PST to confirm historical email is present.' `
                   -ForegroundColor Gray
        Write-Host '  5. Keep your Backups\ folder permanently -- do not delete.' `
                   -ForegroundColor Gray
        Write-Host '     These files are required if the pipeline ever needs to re-run.' `
                   -ForegroundColor Gray
        Write-Host '     If disk space is a concern, move the folder to external storage' `
                   -ForegroundColor Gray
        Write-Host '     and keep a record of its location for future recovery.' `
                   -ForegroundColor Gray
    }
    elseif ($Script:FinalStatus -eq 'WARNING') {
        Write-Host '  Migration completed with warnings.' -ForegroundColor Yellow
        Write-Host '  Review the Migration Report for details.' -ForegroundColor Yellow
    }
    else {
        Write-Host '  Migration encountered failures.' -ForegroundColor Red
        Write-Host '  Review the Migration Report and re-run if needed.' -ForegroundColor Red
        Write-Host '  Your backup PST files are intact in the Backups\ folder.' `
                   -ForegroundColor Yellow
    }
    Write-Host ''

}
catch {
    Write-OMMigrateLog -Message "FATAL ERROR in Script 03: $_" -Level ERROR
    Write-OMMigrateLog -Message "Stack trace: $($_.ScriptStackTrace)" -Level DEBUG
    $Script:FinalStatus = 'FAILED'

    Write-Host ''
    Write-Host '  FATAL ERROR:' -ForegroundColor Red
    Write-Host "  $_" -ForegroundColor Red
    Write-Host ''
    Write-Host '  Your account settings from Script 02 are intact.' `
               -ForegroundColor Yellow
    Write-Host '  Backup PST files are safe in the Backups\ folder.' `
               -ForegroundColor Yellow
    Write-Host "  Log: $($Global:OMMigrate.RunLogFile)" -ForegroundColor Gray
    Write-Host ''
}
finally {
    # -- Always runs --------------------------------------------

    # Added 2026-07-10, Administrator direction (freeze lifted specifically for this
    # fix). Detach the Archive PST here, in the always-runs finally block,
    # BEFORE Release-OutlookCOM quits Outlook -- matches the same
    # already-mounted-aware pattern used for every backup PST in this file
    # ($archivePSTWasAlreadyMounted set alongside Open-PSTFile near the top
    # of this script). Only detaches if this run's own code did the
    # mounting; if Administrator (or any admin) had it manually attached before this
    # script ran, it is left exactly as they had it. Close-PSTFile is safe
    # to call even if $Script:ArchivePSTPath was never actually opened this
    # run (e.g. the "PST not found" branch) -- it no-ops on an unmounted path.
    if ($Script:ArchivePSTPath -and
        -not $Global:OMMigrate.WhatIf -and
        -not $archivePSTWasAlreadyMounted -and
        -not $archiveIsMasterArchive) {
        try {
            Close-PSTFile -PSTPath $Script:ArchivePSTPath | Out-Null
            Write-OMMigrateLog -Message "Archive PST detached at script end: $Script:ArchivePSTPath" -Level INFO
        }
        catch {
            Write-OMMigrateLog -Message "Failed to detach Archive PST at script end: $_" -Level WARN
        }
    }
    elseif ($Script:ArchivePSTPath -and $archiveIsMasterArchive) {
        Write-OMMigrateLog -Message "Archive PST is a protected master archive -- leaving attached: $Script:ArchivePSTPath" -Level DEBUG
    }
    elseif ($Script:ArchivePSTPath -and $archivePSTWasAlreadyMounted) {
        Write-OMMigrateLog -Message "Archive PST was already mounted before this script ran -- leaving attached: $Script:ArchivePSTPath" -Level DEBUG
    }

    if ($Script:COMSessionOpen) {
        try { Resume-OutlookSendReceive | Out-Null } catch { }
        try { Release-OutlookCOM } catch { }
        $Script:COMSessionOpen = $false
    }

    # -- Post-COM session patch REMOVED (June 2026) ──────────────────────────
    # The pre-Save bytes 44-45 patch in the secondary store loop sets the
    # correct count before IRulesCollection.Save(). Save() preserves the
    # pre-patched count (proved by Test-WriteRulesStream4.ps1). A post-COM
    # session is not needed and was causing the UI bleed by opening a second
    # Outlook COM session without zeroing bytes 44-45 first.

    # Only save checkpoint if work actually started (progress was initialized)
    # Avoids writing a checkpoint when operator declined at pre-flight
    if ($Script:COMSessionOpen -or $Script:AccountResults.Count -gt 0) {
        Save-OMMigrateCheckpoint -ExitType 'NORMAL' -Reason 'Script 03 session ending'
    }

    # Clean up checkpoint file after a fully successful session.
    # A clean run means all selected accounts completed -- no resume needed.
    # Failed or warned sessions keep the checkpoint so the operator can resume.
    if ($Script:FinalStatus -eq 'SUCCESS') {
        $checkpointPath = Join-Path $Global:OMMigrate.ManifestPath 'Step03_Checkpoint.json'
        if (Test-Path $checkpointPath) {
            try {
                Remove-Item $checkpointPath -Force -ErrorAction Stop
                Write-OMMigrateLog -Message 'Step03_Checkpoint.json removed after successful session.' `
                                   -Level DEBUG
            }
            catch {
                Write-OMMigrateLog -Message "Could not remove Step03_Checkpoint.json: $_" -Level INFO
            }
        }
    }

    if ($Global:OMMigrate) {
        $Global:OMMigrate.SessionCompletedNormally = $true
    }

    Complete-OMMigrateSession `
        -Status     $Script:FinalStatus `
        -ReportFile $Script:ReportFile

    if ($Script:ReportFile -and (Test-Path $Script:ReportFile) -and -not $Script:IsWhatIf) {
        Open-FileInEditor -FilePath $Script:ReportFile
    }

    # Keep PowerShell window open regardless of how it was launched
    Wait-UserKeypress
}
# ***** END OF FILE *****
