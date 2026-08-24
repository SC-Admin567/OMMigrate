#Requires -Version 5.1
# ============================================================
#  WARNING -- SCRIPT UNDER ACTIVE DEVELOPMENT
# ============================================================
#  Script 02 is NOT ready for production use.
#  Known issues being fixed:
#    1. POP3 removal now uses guided manual process via Outlook
#       Account Settings -- registry deletion retired after causing
#       profile corruption in May 2026 session (required ntuser.dat
#       restore to recover). See Remove-POP3Account in
#       OMMigrate-Outlook.psm1 for full explanation.
#    2. Accounts.Add() COM method not supported in Outlook 2021
#       -- IMAP account add uses guided manual File > Add Account
#
#  DO NOT RUN this script until both issues are resolved and
#  WhatIf simulation runs completely clean end to end.
# ============================================================
<#
.SYNOPSIS
    OMMigrate-02-Convert.ps1 -- Account Protocol Conversion

.DESCRIPTION
    Step 02 of the OutlookMailMigrator (OMMigrate) toolkit.

    This is the most critical script in the migration. It removes
    each POP3 account from Outlook and adds it back as IMAP using
    the settings from Config\migration_accounts.csv.

    SAFETY GATES -- this script will NOT run unless:
        - Step 00 manifest exists and shows SUCCESS or WARNING
        - Step 01 manifest exists and shows SUCCESS or WARNING
        - Every verified backup path from Step 01 exists on disk
          and meets the minimum size requirement
        - The operator confirms the pre-flight checklist
        - The operator confirms each individual account before
          its POP3 entry is removed

    WHAT THIS SCRIPT DOES (per account):
        1. Checks live Outlook -- if already IMAP, auto-skips cleanly
        2. Verifies the backup PST exists and is intact
        3. Displays account details -- old POP3 settings, new IMAP settings
        4. Prompts [y/N/EXIT] -- default N -- before removing POP3
        5. Guides operator to remove the POP3 account via Outlook
           File > Account Settings > Account Settings > Remove
           (Outlook handles all internal MAPI cleanup -- no registry writes)
        6. Closes Outlook COM session between phases
        7. Guides operator to add the new IMAP account via Outlook
           File > Add Account > Manual setup > IMAP
        8. Pauses for operator to enter credentials in Outlook dialog
        9. Waits for operator to confirm connection is live in Outlook
       10. Logs and audits the result
       11. Saves checkpoint -- safe to resume if interrupted

    ALREADY-IMAP DETECTION:
        Before processing each account, the script asks Outlook directly
        whether that email address is already configured as IMAP. If it
        is -- regardless of what the CSV says -- the account is skipped
        automatically with a clear green status message. This means:
            - You can leave converted accounts as MIGRATE in the CSV
            - Re-running after partial migrations is always safe
            - Non-consecutive migrations across multiple sessions work correctly

    ACCOUNT HANDLING BY TAG:
        POP3-STANDARD     Full removal + IMAP add
        POP3-ATTAMERITECH Full removal + IMAP add (Yahoo servers, Secure Mail Key)
        POP3-AWS          Full removal + IMAP add (AWS SES SMTP re-entry required)
        POP3-GMAIL        Full removal + IMAP add (Gmail -- browser OAuth)
        IMAP-ALREADY      Auto-skipped -- confirmed IMAP in live Outlook session
        IMAP-GMAIL        Auto-skipped -- confirmed IMAP in live Outlook session
        EXCHANGE-SKIP     Auto-skipped -- Exchange, no action needed
        UNKNOWN           Auto-skipped -- flagged for manual review

    PASSWORDS:
        This script NEVER handles, stores, or logs passwords.
        When adding each IMAP account, Outlook will display its own
        password dialog. The operator enters the password there.
        The script pauses and waits for the operator to confirm.

    RESUME SUPPORT:
        If interrupted, re-running this script performs a live Outlook
        check on each account. Already-converted accounts are detected
        and skipped automatically. No manual checkpoint management needed.

.PARAMETER BasePath
    Override the default working directory.
    Must match the BasePath used in Scripts 00 and 01.
    Default: $env:USERPROFILE\Documents\OutlookMigration

.PARAMETER Preview
    Simulate all operations -- no accounts are removed or added.
    Reads manifests and CSV normally but takes no Outlook action.
    Useful for a final dry-run review before committing.

.PARAMETER LogLevel
    Logging verbosity: DEBUG | INFO | WARN | ERROR
    Default: INFO

.PARAMETER SkipBackupCheck
    Skip the backup file existence and size verification gate.
    NOT RECOMMENDED. Only use if you have independently verified
    all backups and understand the risk of proceeding without them.

.PARAMETER Force
    Skip the per-account Y/N confirmation prompts.
    The initial pre-flight confirmation is still required.
    Use with extreme caution -- this removes accounts automatically.

.EXAMPLE
    # Standard run
    .\OMMigrate-02-Convert.ps1

.EXAMPLE
    # Dry run -- see what would happen without making changes
    .\OMMigrate-02-Convert.ps1 -Preview

.EXAMPLE
    # Resume after interruption or partial migration
    # Simply re-run -- already-IMAP accounts auto-detected and skipped
    .\OMMigrate-02-Convert.ps1

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
        OMMigrate-00-Discover.ps1  -- must have completed (Step00 manifest)
        OMMigrate-01-Backup.ps1    -- must have completed (Step01 manifest)
        Config\migration_accounts.csv -- passwords filled in by operator
        Secure Mail Key generated for ameritech.net (if applicable)

    Outlook must be closed before running.
    This script launches Outlook via COM and requires exclusive access.

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
    [switch]$SkipBackupCheck,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

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
    -ScriptName 'OMMigrate-02-Convert' `
    -BasePath   $BasePath `
    -LogLevel   $LogLevel `
    -IsWhatIf   $Script:IsWhatIf `
    -Sanitize   $Sanitize.IsPresent

# Verify WhatIf state was captured correctly
Write-OMMigrateLog -Message "DEBUG: Script:IsWhatIf=$($Script:IsWhatIf) | Global.WhatIf=$($Global:OMMigrate.WhatIf) | Preview=$($Preview.IsPresent)" `
                   -Level INFO

# Register Ctrl+C and exit handlers -- CRITICAL for Script 02
Register-ExitHandlers -ScriptStep 2

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

# Show safe-exit banner prominently before any work
Show-ExitBanner

# Script-level state
$Script:COMSessionOpen   = $false
$Script:FinalStatus      = 'SUCCESS'
$Script:ReportFile       = ''
$Script:AccountResults   = [System.Collections.Generic.List[PSCustomObject]]::new()
$Script:MigratedCount      = 0
$Script:FailedCount        = 0
$Script:SkippedCount       = 0
$Script:AutoSkippedCount   = 0   # MIGRATE accounts found already IMAP in live Outlook -- resume safety net
$Script:NotMarkedCount     = 0   # All non-MIGRATE accounts combined -- used internally, not shown in summary
$Script:WarningCount       = 0   # Accounts migrated but connection not confirmed by operator
$Script:AlreadyImapCount   = 0   # Accounts with IMAP-ALREADY or IMAP-GMAIL ProviderTag -- never candidates
$Script:ExchangeCount      = 0   # Accounts with EXCHANGE-SKIP ProviderTag -- never candidates
$Script:OtherNotMarkedCount = 0  # Remaining non-MIGRATE accounts (FOLDER-ONLY, SKIP, etc.)


# ============================================================
#  HELPER: Get-LiveAccountType
#  Asks Outlook directly what protocol an email address is
#  currently using. This is the authoritative source -- not the
#  CSV, not the checkpoint file.
# ============================================================

function Get-LiveAccountType {
    <#
    .SYNOPSIS
        Returns the live Outlook account type for an email address
        by querying the active COM session directly.

    .PARAMETER Email
        The email address to look up.

    .PARAMETER ComAccounts
        Pre-fetched array of COM account objects from
        Get-OutlookAccountsViaCOM -- passed in to avoid
        re-querying Outlook for every account in the loop.

    .OUTPUTS
        [string] -- 'IMAP' | 'POP3' | 'Exchange' | 'NotFound'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email,

        [Parameter(Mandatory = $true)]
        [array]$ComAccounts
    )

    $match = $ComAccounts | Where-Object {
        $_.EmailAddress -eq $Email
    } | Select-Object -First 1

    if (-not $match) { return 'NotFound' }
    return $match.AccountType
}


# ============================================================
#  MAIN EXECUTION BLOCK
# ============================================================

try {

    # ----------------------------------------------------------
    #  STEP 1 -- Gate Checks: Manifests and Backup Verification
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Pre-Flight Gate Checks' -Step '1 of 5'

    # Gate 1: Script 00 manifest
    Write-OMMigrateLog -Message 'Verifying Script 00 manifest...' -Level INFO
    $step00Manifest = Read-StepManifest -Step 0
    Write-Host "  Script 00 manifest : OK ($($step00Manifest.Data.AccountsDiscovered) accounts discovered)" `
               -ForegroundColor Green

    # Gate 2: Script 01 manifest
    Write-OMMigrateLog -Message 'Verifying Script 01 manifest...' -Level INFO
    $step01Manifest = Read-StepManifest -Step 1
    Write-Host "  Script 01 manifest : OK ($($step01Manifest.Data.BackupSucceeded) backups verified)" `
               -ForegroundColor Green

    # -- Step02 manifest auto-clear --------------------------------
    # If a previous Step02_Complete.json exists, check whether all the
    # accounts it recorded have since been updated to IMAP-CONVERTED
    # in migration_accounts.csv. If so, clear the manifest automatically
    # so this run can proceed without a manual delete step.
    # This replaces the manual "delete Step02_Complete.json" step the
    # operator previously had to perform between Script 02 runs.
    # Skipped in WhatIf mode -- no file changes.
    $step02ManifestPath = Join-Path $Global:OMMigrate.ManifestPath 'Step02_Complete.json'
    if ((Test-Path $step02ManifestPath) -and -not $Script:IsWhatIf) {
        try {
            $prevManifest  = Read-StepManifest -Step 2
            $prevCompleted = @()
            if ($prevManifest.Data.CompletedAccounts) {
                $prevCompleted = @($prevManifest.Data.CompletedAccounts)
            }

            if ($prevCompleted.Count -gt 0) {
                # Load CSV to check current state of previously completed accounts
                $csvCheckPath = Get-OMMigrateCsvPath -BaseName 'migration_accounts.csv'
                $stillPOP3    = @()
                if (Test-Path $csvCheckPath) {
                    $csvRows   = Import-Csv -Path $csvCheckPath -Encoding UTF8
                    $stillPOP3 = @($prevCompleted | Where-Object {
                        $email = $_
                        $row   = $csvRows | Where-Object { $_.EmailAddress -eq $email } |
                                 Select-Object -First 1
                        $row -and $row.ProviderTag -notlike 'IMAP-*'
                    })
                }
                $allConverted = ($stillPOP3.Count -eq 0)

                if ($allConverted) {
                    Write-Host ''
                    Write-Host ('  ' + ('-' * 54)) -ForegroundColor Cyan
                    Write-Host '  Previous Script 02 run detected.' -ForegroundColor Cyan
                    Write-Host "  $($prevCompleted.Count) account(s) from that run are now IMAP-CONVERTED." `
                               -ForegroundColor Cyan
                    Write-Host '  Clearing Step02_Complete.json so this run can proceed.' `
                               -ForegroundColor Cyan
                    Write-Host ('  ' + ('-' * 54)) -ForegroundColor Cyan
                    Write-Host ''

                    Remove-Item $step02ManifestPath -Force -ErrorAction SilentlyContinue
                    Write-OMMigrateLog -Message (
                        "Step02_Complete.json auto-cleared -- all $($prevCompleted.Count) previously " +
                        "completed account(s) confirmed IMAP-CONVERTED in CSV."
                    ) -Level INFO
                    Write-AuditEntry -Action 'STEP02_MANIFEST_AUTO_CLEARED' `
                                     -Detail "PreviouslyCompleted=$($prevCompleted -join ', ')"
                }
                else {
                    Write-Host ''
                    Write-Host '  NOTE: Previous Step02_Complete.json found.' -ForegroundColor DarkGray
                    Write-Host '  Some accounts from the previous run have not yet been confirmed' `
                               -ForegroundColor DarkGray
                    Write-Host '  as IMAP-CONVERTED -- manifest retained.' -ForegroundColor DarkGray
                    Write-Host '  If you need to re-run for new accounts, delete' -ForegroundColor DarkGray
                    Write-Host '  Manifests\Step02_Complete.json manually before proceeding.' `
                               -ForegroundColor DarkGray
                    Write-Host ''
                    Write-OMMigrateLog -Message (
                        "Step02_Complete.json NOT auto-cleared -- $($stillPOP3.Count) previously " +
                        "completed account(s) still show non-IMAP ProviderTag in CSV."
                    ) -Level WARN
                }
            }
        }
        catch {
            # Non-fatal -- if manifest check fails just log and continue
            Write-OMMigrateLog -Message "Step02 manifest auto-clear check failed (non-fatal): $_" `
                               -Level WARN
        }
    }

    # Gate 3: Backup file integrity check
    if (-not $SkipBackupCheck) {
        Write-OMMigrateLog -Message 'Verifying backup files on disk...' -Level INFO

        $verifiedBackups   = $step01Manifest.Data.VerifiedBackups
        $backupCheckFailed = $false

        if (-not $verifiedBackups -or $verifiedBackups.Count -eq 0) {
            # If Script 01 ran with zero accounts to back up (BackupSucceeded = 0),
            # an empty VerifiedBackups list is correct -- no backups were needed.
            # Only fail the gate if Script 01 reported at least one successful backup
            # but produced no verified paths (which would indicate a manifest error).
            if ($step01Manifest.Data.BackupSucceeded -gt 0) {
                Write-OMMigrateLog -Message (
                    'No verified backup paths in Step 01 manifest. ' +
                    'Re-run Script 01 before proceeding.'
                ) -Level ERROR
                Write-Host '  GATE FAILED: No verified backups in Step 01 manifest.' `
                           -ForegroundColor Red
                $Script:FinalStatus = 'FAILED'
                exit 1
            }
            # Zero backups needed -- gate passes
            Write-OMMigrateLog -Message 'No POP3 backups in Step 01 manifest -- no accounts required backup. Gate passed.' `
                               -Level INFO
        }

        foreach ($backupPath in $verifiedBackups) {
            # Compute masked path for console output -- real path goes to log only.
            # NOTE: The sanitize map is not yet built at this point (accounts CSV is
            # loaded in Step 2, after this gate check). Use a generic placeholder
            # for the filename when -Sanitize is active rather than calling
            # Invoke-OMMigrateSanitize which would have nothing to substitute yet.
            $backupFileName   = [System.IO.Path]::GetFileName($backupPath)
            $maskedFileName   = if ($Global:OMMigrate.Sanitize) { '[backup-file].pst' } else { $backupFileName }
            $backupPathMasked = Join-Path $Global:OMMigrate.BackupPath $maskedFileName

            if (-not (Test-Path $backupPath)) {
                Write-OMMigrateLog -Message "Backup file MISSING: $backupPathMasked" -Level ERROR
                Write-Host "  GATE FAILED: Backup missing -- $backupPathMasked" -ForegroundColor Red
                $backupCheckFailed = $true
            }
            else {
                # Read minimum size from settings -- default 0 (no minimum)
                $minBackupBytes = 0L
                try {
                    $minMB = $Global:OMMigrate.Settings.BackupVerification.MinimumSizeMB
                    if ($null -ne $minMB) { $minBackupBytes = [long]($minMB * 1MB) }
                }
                catch { }

                $fileSize = (Get-Item $backupPath).Length
                if ($minBackupBytes -gt 0 -and $fileSize -lt $minBackupBytes) {
                    Write-OMMigrateLog -Message "Backup file too small: $backupPathMasked" -Level ERROR
                    Write-Host "  GATE FAILED: Backup too small -- $backupPathMasked" -ForegroundColor Red
                    $backupCheckFailed = $true
                }
            }
        }

        if ($backupCheckFailed) {
            Write-Host ''
            Write-Host '  One or more backup files are missing or too small.' -ForegroundColor Red
            Write-Host '  Re-run Script 01 to re-verify backups before proceeding.' `
                       -ForegroundColor Red
            $Script:FinalStatus = 'FAILED'
            exit 1
        }

        Write-Host "  Backup verification : OK ($($verifiedBackups.Count) files on disk)" `
                   -ForegroundColor Green
    }
    else {
        Write-OMMigrateLog -Message '-SkipBackupCheck specified -- backup verification bypassed.' `
                           -Level WARN
        Write-Host '  Backup verification : SKIPPED (-SkipBackupCheck)' -ForegroundColor Yellow
    }

    # Gate 4: Environment check
    $envResult = Test-OMMigrateEnvironment
    if (-not $envResult.Passed) {
        Write-OMMigrateLog -Message 'Environment check failed.' -Level ERROR
        $Script:FinalStatus = 'FAILED'
        exit 1
    }
    Write-Host '  Environment check   : OK' -ForegroundColor Green
    Write-Host ''


    # ----------------------------------------------------------
    #  STEP 2 -- Load Account List
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Loading Account List' -Step '2 of 5'

    $accountsCsvPath = Get-OMMigrateCsvPath -BaseName 'migration_accounts.csv'
    if (-not (Test-Path $accountsCsvPath)) {
        Write-OMMigrateLog -Message "migration_accounts.csv not found: $accountsCsvPath" `
                           -Level ERROR
        $Script:FinalStatus = 'FAILED'
        exit 1
    }

    $allAccounts = Import-Csv -Path $accountsCsvPath -Encoding UTF8

    # -- Build sanitization map from loaded accounts -----------
    if ($Global:OMMigrate.Sanitize) {
        Initialize-SanitizeMap -Accounts @($allAccounts)

        # Register SmtpUsername values that differ from the email address.
        # AWS SES accounts use an IAM access key ID as the SMTP username --
        # that value is not an email address so Initialize-SanitizeMap won't
        # pick it up automatically. Register it here so it gets masked.
        foreach ($acct in $allAccounts) {
            if ($acct.PSObject.Properties['SmtpUsername'] -and
                $acct.SmtpUsername -and
                $acct.SmtpUsername -ne $acct.EmailAddress) {
                Register-SanitizeTerms -Terms @($acct.SmtpUsername) -Category 'Generic'
            }
        }

        Write-OMMigrateLog -Message '[SANITIZE] Sanitization active -- output masked.' `
                           -Level INFO
    }

    # Accounts marked MIGRATE are candidates -- live Outlook check
    # will determine which of these actually need processing
    $accountsToMigrate = @($allAccounts | Where-Object {
        $_.MigrationAction -eq 'MIGRATE'
    })

    # Accounts not marked MIGRATE are skipped immediately
    $accountsNotMigrate = @($allAccounts | Where-Object {
        $_.MigrationAction -ne 'MIGRATE'
    })

    Write-OMMigrateLog -Message (
        "CSV accounts marked MIGRATE: $($accountsToMigrate.Count) | " +
        "Not marked MIGRATE: $($accountsNotMigrate.Count)"
    ) -Level INFO

    # Add non-MIGRATE accounts to results immediately.
    # Split into three buckets for accurate operator-facing summary:
    #   AlreadyImapCount    -- IMAP-ALREADY / IMAP-GMAIL -- already on correct protocol
    #   ExchangeCount       -- EXCHANGE-SKIP -- Exchange accounts, no migration needed
    #   OtherNotMarkedCount -- everything else (FOLDER-ONLY, SKIP, MANUAL, etc.)
    foreach ($notMigrate in $accountsNotMigrate) {
        $skipResult = $notMigrate.PSObject.Copy()
        $skipResult | Add-Member -NotePropertyName 'MigrationOutcome' `
                                 -NotePropertyValue 'SKIPPED' -Force
        $skipResult | Add-Member -NotePropertyName 'MigrationDetail' `
                                 -NotePropertyValue "MigrationAction=$($notMigrate.MigrationAction) in CSV -- not a MIGRATE candidate" `
                                 -Force
        [void]$Script:AccountResults.Add($skipResult)
        $Script:NotMarkedCount++

        # Categorize for summary display
        switch -Wildcard ($notMigrate.ProviderTag) {
            'IMAP-*'     { $Script:AlreadyImapCount++    }
            'EXCHANGE-*' { $Script:ExchangeCount++        }
            default      { $Script:OtherNotMarkedCount++  }
        }
    }

    Write-Host "  Accounts marked MIGRATE in CSV : $($accountsToMigrate.Count)" `
               -ForegroundColor Cyan
    Write-Host "  Accounts not marked MIGRATE    : $($accountsNotMigrate.Count)" `
               -ForegroundColor DarkGray
    Write-Host ''

    # -- Zero-MIGRATE gate ------------------------------------
    # If no accounts are marked MIGRATE there is nothing for this
    # script to do. Exit cleanly before launching Outlook at all --
    # no point opening a COM session, suspending Send/Receive, or
    # prompting the operator for a run that will do nothing.
    # This happens when all accounts have been migrated and the CSV
    # has been updated to FOLDER-ONLY / SKIP.
    if ($accountsToMigrate.Count -eq 0) {
        Write-Host ''
        Write-Host ('  ' + ('-' * 54)) -ForegroundColor Cyan
        Write-Host '  Nothing to migrate.' -ForegroundColor Cyan
        Write-Host ''
        Write-Host '  No accounts are marked MIGRATE in migration_accounts.csv.' `
                   -ForegroundColor Gray
        Write-Host '  All accounts have either been migrated or are set to SKIP.' `
                   -ForegroundColor Gray
        Write-Host ''
        Write-Host '  If migration is complete, proceed to Script 03:' `
                   -ForegroundColor Green
        Write-Host '  .\Scripts\OMMigrate-03-Restore.ps1' -ForegroundColor White
        Write-Host ('  ' + ('-' * 54)) -ForegroundColor Cyan
        Write-Host ''

        Write-OMMigrateLog -Message 'Zero accounts marked MIGRATE -- nothing to do. Exiting cleanly.' `
                           -Level INFO
        Write-AuditEntry  -Action 'NOTHING_TO_MIGRATE' `
                          -Detail 'Zero MIGRATE accounts in CSV -- script exited before launching Outlook.'

        # Generate report and manifest so the run is on record
        $Script:ReportFile = New-MigrationReport -Accounts @($Script:AccountResults | Sort-Object EmailAddress)
        Write-StepManifest -Step 2 -Status 'SUCCESS' -Data @{
            MigratedCount        = 0
            AlreadyImapCount     = $Script:AlreadyImapCount
            ExchangeCount        = $Script:ExchangeCount
            AutoSkippedCount     = 0
            FailedCount          = 0
            SkippedCount         = 0
            WarningCount         = 0
            OtherNotMarkedCount  = $Script:OtherNotMarkedCount
            NotMarkedCount       = $Script:NotMarkedCount
            TotalAccountCount    = $Script:NotMarkedCount
            TotalProcessed       = $Script:AccountResults.Count
            ReportFile           = $Script:ReportFile
            CompletedAccounts    = @()
        }

        $Script:FinalStatus = 'SUCCESS'
        exit 0
    }

    Write-Host '  NOTE: The script will now launch Outlook and check each' `
               -ForegroundColor Gray
    Write-Host '  MIGRATE account live. Any account already showing as IMAP' `
               -ForegroundColor Gray
    Write-Host '  in Outlook will be skipped automatically.' -ForegroundColor Gray
    Write-Host ''

    # Password readiness check -- warn but do not expose placeholder text
    $missingPasswords = @($accountsToMigrate | Where-Object {
        $_.Password -like '*ENTER_PASSWORD*' -or
        $_.Password -like '*ENTER_SECURE*'   -or
        [string]::IsNullOrWhiteSpace($_.Password)
    })

    if ($missingPasswords.Count -gt 0 -and -not $Script:IsWhatIf) {
        Write-Host ''
        Write-Host ('  ' + ('-' * 54)) -ForegroundColor Yellow
        Write-Host '  PASSWORD REMINDER' -ForegroundColor Yellow
        Write-Host ('  ' + ('-' * 54)) -ForegroundColor Yellow
        Write-Host ''
        Write-Host "  $($missingPasswords.Count) account(s) still have placeholder" `
                   -ForegroundColor Yellow
        Write-Host '  passwords in the CSV. Outlook will prompt you to enter' `
                   -ForegroundColor Yellow
        Write-Host '  the password for each of these accounts:' -ForegroundColor Yellow
        Write-Host ''
        foreach ($mp in $missingPasswords) {
            Write-Host "    * $(Invoke-OMMigrateSanitize -Text $mp.EmailAddress)" -ForegroundColor White
        }
        Write-Host ''
        Write-Host '  Have your email passwords ready before continuing.' `
                   -ForegroundColor Gray

        # Special reminder for AT&T/Yahoo accounts
        $attAccounts = @($missingPasswords | Where-Object {
            $_.ProviderTag -eq 'POP3-ATTAMERITECH'
        })
        if ($attAccounts.Count -gt 0) {
            Write-Host ''
            Write-Host '  AT&T/Yahoo accounts require a Secure Mail Key,' `
                       -ForegroundColor Yellow
            Write-Host '  not your regular email password.' -ForegroundColor Yellow
            Write-Host '  Generate one at: currently.com -> Account Security ->' `
                       -ForegroundColor Yellow
            Write-Host '  Secure Mail Key -> Manage -> Generate' -ForegroundColor Yellow
            Write-Host '  https://www.att.com/acctmgmt/myprofile/overview?flow=settings' `
                       -ForegroundColor Cyan
        }

        Write-Host ''
        $passwordConfirmed = Confirm-Action `
            -Message      'I have all required passwords and Secure Mail Keys ready' `
            -AccountEmail '' `
            -DefaultYes   $false

        if (-not $passwordConfirmed) {
            Write-Host ''
            Write-Host '  Exiting -- please gather your passwords and re-run.' `
                       -ForegroundColor Yellow
            Write-Host '  Your accounts have not been changed.' -ForegroundColor Green
            exit 0
        }
    }


    # ----------------------------------------------------------
    #  STEP 3 -- Pre-Flight Warning and Final Confirmation
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Pre-Flight Confirmation' -Step '3 of 5'

    $prereqs = [System.Collections.Generic.List[string]]::new()
    $prereqs.Add("Script 01 backups verified -- $($step01Manifest.Data.BackupSucceeded) PST files confirmed in Backups\ folder")
    $prereqs.Add("Config\migration_accounts.csv reviewed -- passwords entered for all POP3 accounts")
    $prereqs.Add("Outlook is fully closed (this script will launch it via COM)")

    $attAccounts = @($accountsToMigrate | Where-Object {
        $_.RequiresSecureKey -eq 'True'
    })
    if ($attAccounts.Count -gt 0) {
        foreach ($att in $attAccounts) {
            $prereqs.Add("Secure Mail Key generated and ready for $($att.EmailAddress)")
        }
    }

    $awsAccounts = @($accountsToMigrate | Where-Object { $_.IsAWSSES -eq 'True' })
    if ($awsAccounts.Count -gt 0) {
        $prereqs.Add("AWS SES SMTP credentials (IAM SMTP username and password) ready -- browser will open for AWS login, then Alt+Tab back to this console to continue")
    }

    # -- Auto-launch browser URLs for MIGRATE accounts requiring credentials --
    # Scoped to accounts marked MIGRATE only -- no point opening these for
    # SKIP / FOLDER-ONLY / EXCHANGE accounts.
    # Uses cmd.exe /c start routing to guarantee the default browser opens
    # rather than whatever app Windows has associated with https: (Acrobat
    # was opened instead of a browser in Script 00 during May 2026 testing).
    # Skipped in WhatIf/Preview mode.
    if (-not $Script:IsWhatIf) {

        # AWS SES SMTP credentials page -- shown when any MIGRATE account uses AWS SES
        if ($awsAccounts.Count -gt 0) {
            try {
                Start-Process 'cmd.exe' -ArgumentList '/c', 'start', 'https://console.aws.amazon.com/ses/home#/smtp'
                Write-Host ''
                Write-Host '  Opening AWS SES SMTP Settings in your browser...' `
                           -ForegroundColor DarkGray
                Write-Host '  SES > SMTP Settings > Create SMTP Credentials' `
                           -ForegroundColor DarkGray
                Write-OMMigrateLog -Message 'Opened AWS SES SMTP Settings URL in browser (MIGRATE accounts).' -Level INFO
            }
            catch {
                Write-OMMigrateLog -Message "Could not open AWS SES URL: $_" -Level INFO
            }
        }

        # AT&T Secure Mail Key page -- shown when any MIGRATE account requires a Secure Mail Key
        $attMigrateAccounts = @($accountsToMigrate | Where-Object { $_.RequiresSecureKey -eq 'True' })
        if ($attMigrateAccounts.Count -gt 0) {
            try {
                Start-Process 'cmd.exe' -ArgumentList '/c', 'start', 'https://www.att.com/acctmgmt/myprofile/overview?flow=settings'
                Write-Host ''
                Write-Host '  Opening AT&T account security settings in your browser...' `
                           -ForegroundColor DarkGray
                Write-Host '  Account Security -> Secure Mail Key -> Manage -> Generate' `
                           -ForegroundColor DarkGray
                Write-OMMigrateLog -Message 'Opened AT&T/Yahoo Secure Mail Key URL in browser (MIGRATE accounts).' -Level INFO
            }
            catch {
                Write-OMMigrateLog -Message "Could not open Secure Mail Key URL: $_" -Level INFO
            }
        }
    }
    else {
        if ($awsAccounts.Count -gt 0) {
            Write-OMMigrateLog -Message 'WhatIf: Would open AWS SES SMTP Settings URL in browser.' `
                               -Level INFO -WhatIfPrefix
        }
        $attMigrateAccounts = @($accountsToMigrate | Where-Object { $_.RequiresSecureKey -eq 'True' })
        if ($attMigrateAccounts.Count -gt 0) {
            Write-OMMigrateLog -Message 'WhatIf: Would open AT&T/Yahoo Secure Mail Key URL in browser.' `
                               -Level INFO -WhatIfPrefix
        }
    }

    Show-PreflightWarning `
        -ScriptDescription (
            "This script will check each of the $($accountsToMigrate.Count) " +
            "account(s) marked MIGRATE against live Outlook. Accounts already " +
            "converted to IMAP will be skipped automatically. For accounts " +
            "still on POP3 you will be prompted Y/N before anything is removed. " +
            "Type EXIT at any prompt to stop safely."
        ) `
        -Prerequisites $prereqs.ToArray() `
        -DeclineMessage @(
            "You chose not to proceed at this time. No accounts have been changed.",
            "When you are ready to migrate, re-run this script:",
            "  .\Scripts\OMMigrate-02-Convert.ps1",
            "IMPORTANT: This script will DELETE each POP3 account and replace it",
            "with a new IMAP account. Make sure your PST backups are verified",
            "before proceeding. You will be prompted Y/N before each account",
            "is touched -- type EXIT at any prompt to stop safely at any time."
        )

    # Auto-close Outlook if running -- scripts launch their own COM session
    # and require exclusive access. Moved here, AFTER the operator's Y/N
    # confirmation above (Show-PreflightWarning exits the script entirely
    # on a No/decline, so anything below this point only runs after an
    # explicit Yes) -- this is the interactive Outlook-dialog-driven script
    # (Account Settings Remove, Add Account wizard), so the admin may still
    # be mid-task in Outlook from a prior account. Closing automatically
    # only once the admin has explicitly said they are ready to continue.
    [void](Close-OutlookIfRunning -Reason 'before pre-flight')

    # Initialize progress tracker
    $pendingEmails = @($accountsToMigrate | ForEach-Object { $_.EmailAddress })
    Update-OMMigrateProgress -SetPending $pendingEmails


    # ----------------------------------------------------------
    #  STEP 4 -- Launch COM and Process Each Account
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Converting Accounts' -Step '4 of 5'

    Write-Host '  Launching Outlook...' -ForegroundColor Cyan
    Write-Host '  (Outlook will open in the foreground -- do not interact until prompted)' -ForegroundColor DarkGray
    Write-Host ''

    # VisibleLaunch=true -- launches Outlook as a visible foreground window
    # via Start-Process, then attaches via GetActiveObject once the COM
    # class is registered. Replaces the old New-Object -ComObject path
    # which launched Outlook as a hidden background process and required
    # the operator to manually click the taskbar icon to see it.
    # Read selected profile from Settings.json -- set by Script 00 profile picker.
    # Passed to Connect-OutlookCOM so Outlook opens the correct migration profile.
    $s02ProfileName = ''
    try {
        if ($Global:OMMigrate.Settings -and
            $Global:OMMigrate.Settings.PSObject.Properties['OutlookProfile'] -and
            $Global:OMMigrate.Settings.OutlookProfile.SelectedProfile) {
            $s02ProfileName = $Global:OMMigrate.Settings.OutlookProfile.SelectedProfile
        }
    }
    catch { }
    Write-OMMigrateLog -Message "Using Outlook profile: '$s02ProfileName'" -Level INFO
    $outlook = Connect-OutlookCOM -VisibleLaunch $true -ProfileName $s02ProfileName
    if (-not $outlook -and -not $Script:IsWhatIf) {
        Write-OMMigrateLog -Message 'Failed to start Outlook COM session.' -Level ERROR
        $Script:FinalStatus = 'FAILED'
        exit 1
    }
    if ($outlook) { $Script:COMSessionOpen = $true }

    Write-Host '  Outlook is open and ready.' -ForegroundColor Green

    # Suspend Send/Receive to prevent sync interference during conversion
    Suspend-OutlookSendReceive | Out-Null

    Write-Host ''

    # Fetch live account list ONCE -- used for every account's live check
    # This is more efficient than calling the COM API per-account
    $liveComAccounts = @()
    if (-not $Script:IsWhatIf) {
        try {
            $liveComAccounts = @(Get-OutlookAccountsViaCOM)
            Write-OMMigrateLog -Message "Live Outlook accounts loaded: $($liveComAccounts.Count)" `
                               -Level INFO
        }
        catch {
            Write-OMMigrateLog -Message "Could not load live COM accounts: $_" -Level WARN
            Write-Host '  WARNING: Could not read live Outlook account list.' `
                       -ForegroundColor Yellow
            Write-Host '  Live IMAP detection may not work correctly.' -ForegroundColor Yellow
        }
    }

    # IMPORTANT: Send/Receive is intentionally LEFT SUSPENDED here.
    # Outlook stays open through the POP3 removal phase -- the operator
    # needs File > Account Settings accessible. Send/Receive must remain
    # suspended so the POP3 account does not sync while being removed.
    # Send/Receive state is saved to Step02_SendReceiveState.json by
    # Suspend-OutlookSendReceive so Resume-OutlookSendReceive can read
    # it correctly after Outlook is closed and reopened between phases.
    # Resume happens in a second COM session AFTER IMAP is confirmed working.
    # Do NOT call Resume-OutlookSendReceive or Release-OutlookCOM here --
    # the COM session must remain active through Phase A (POP3 removal).
    # The COM release happens inside the account loop between Phase A and
    # Phase B so Outlook is fully closed before the IMAP add dialog opens.

    $accountNumber = 0

    foreach ($account in $accountsToMigrate) {

        $accountNumber++
        $email     = $account.EmailAddress
        $totalAccounts = $accountsToMigrate.Count

        Write-Host ''
        Write-Host ('  ' + ('=' * 54)) -ForegroundColor DarkCyan
        Write-Host "  Account $accountNumber of $totalAccounts : $(Invoke-OMMigrateSanitize -Text $email)" -ForegroundColor White
        Write-Host "  Tag     : $($account.ProviderTag)" -ForegroundColor DarkGray
        Write-Host ('  ' + ('=' * 54)) -ForegroundColor DarkCyan
        Write-Host ''

        # -- LIVE IMAP CHECK ------------------------------------
        # Ask Outlook directly -- is this account already IMAP?
        # This is the authoritative source. The CSV and checkpoint
        # are informational only. Outlook is the truth.
        $liveType = 'Unknown'
        if (-not $Script:IsWhatIf -and $liveComAccounts -and $liveComAccounts.Count -gt 0) {
            $liveType = Get-LiveAccountType -Email $email `
                                            -ComAccounts $liveComAccounts
        }

        if ($liveType -eq 'IMAP') {
            # Already IMAP -- skip cleanly with a positive message
            Write-Host "  [ALREADY IMAP] $(Invoke-OMMigrateSanitize -Text $email)" -ForegroundColor Green
            Write-Host "  This account is already configured as IMAP in Outlook." `
                       -ForegroundColor Green
            Write-Host '  No action needed -- skipping.' -ForegroundColor Green

            Write-OMMigrateLog -Message "Auto-skipped (already IMAP in Outlook): $email" `
                               -Level INFO
            Write-AuditEntry  -Action 'ACCOUNT_ALREADY_IMAP' `
                              -AccountEmail $email `
                              -Detail 'Live Outlook check confirmed IMAP -- auto-skipped' `
                              -Outcome 'SKIPPED'

            $alreadyResult = $account.PSObject.Copy()
            $alreadyResult | Add-Member -NotePropertyName 'MigrationOutcome' `
                                        -NotePropertyValue 'SKIPPED' -Force
            $alreadyResult | Add-Member -NotePropertyName 'MigrationDetail' `
                                        -NotePropertyValue 'Already IMAP -- confirmed by live Outlook check' `
                                        -Force
            [void]$Script:AccountResults.Add($alreadyResult)
            $Script:AutoSkippedCount++
            Update-OMMigrateProgress -MarkComplete $email
            continue
        }

        if ($liveType -eq 'Exchange') {
            # Exchange account -- skip
            Write-Host "  [EXCHANGE] $(Invoke-OMMigrateSanitize -Text $email)" -ForegroundColor DarkYellow
            Write-Host '  Exchange account -- already optimal protocol. Skipping.' `
                       -ForegroundColor DarkYellow

            Write-OMMigrateLog -Message "Auto-skipped (Exchange in Outlook): $email" `
                               -Level INFO

            $exchResult = $account.PSObject.Copy()
            $exchResult | Add-Member -NotePropertyName 'MigrationOutcome' `
                                     -NotePropertyValue 'SKIPPED' -Force
            $exchResult | Add-Member -NotePropertyName 'MigrationDetail' `
                                     -NotePropertyValue 'Exchange account -- no migration needed' `
                                     -Force
            [void]$Script:AccountResults.Add($exchResult)
            $Script:SkippedCount++
            Update-OMMigrateProgress -MarkComplete $email
            continue
        }

        if ($liveType -eq 'NotFound' -and -not $Script:IsWhatIf) {
            # Account not found in live Outlook at all -- warn operator
            Write-Host "  [NOT FOUND] $(Invoke-OMMigrateSanitize -Text $email)" -ForegroundColor Yellow
            Write-Host '  This account was not found in the live Outlook session.' `
                       -ForegroundColor Yellow
            Write-Host '  It may have been removed manually or the profile may have changed.' `
                       -ForegroundColor Yellow
            Write-Host ''

            Write-OMMigrateLog -Message "Account not found in live Outlook: $email" `
                               -Level WARN

            $notFoundResult = $account.PSObject.Copy()
            $notFoundResult | Add-Member -NotePropertyName 'MigrationOutcome' `
                                         -NotePropertyValue 'WARNING' -Force
            $notFoundResult | Add-Member -NotePropertyName 'MigrationDetail' `
                                         -NotePropertyValue 'Account not found in live Outlook -- verify manually' `
                                         -Force
            [void]$Script:AccountResults.Add($notFoundResult)
            $Script:FailedCount++
            if ($Script:FinalStatus -eq 'SUCCESS') { $Script:FinalStatus = 'WARNING' }
            Update-OMMigrateProgress -MarkComplete $email
            continue
        }

        # -- Account is POP3 (or WhatIf mode) -- show details ---
        # Encryption label is derived from port number -- matches Outlook's own terminology:
        #   Port 993 / 995 / 465 --> SSL/TLS
        #   Port 587             --> STARTTLS
        #   All others           --> None
        $pop3InEnc  = switch ([int]$account.IncomingPort)  { 993 {'SSL/TLS'} 995 {'SSL/TLS'} 465 {'SSL/TLS'} 587 {'STARTTLS'} default {'None'} }
        $pop3OutEnc = switch ([int]$account.OutgoingPort)  { 993 {'SSL/TLS'} 995 {'SSL/TLS'} 465 {'SSL/TLS'} 587 {'STARTTLS'} default {'None'} }
        $imapInEnc  = switch ([int]$account.NewImapPort)   { 993 {'SSL/TLS'} 995 {'SSL/TLS'} 465 {'SSL/TLS'} 587 {'STARTTLS'} default {'None'} }
        $imapOutEnc = switch ([int]$account.OutgoingPort)  { 993 {'SSL/TLS'} 995 {'SSL/TLS'} 465 {'SSL/TLS'} 587 {'STARTTLS'} default {'None'} }

        Write-Host '  Current POP3 settings:' -ForegroundColor DarkGray
        Write-Host "    Incoming : $($account.IncomingServer):$($account.IncomingPort) ($pop3InEnc)" `
                   -ForegroundColor DarkGray
        Write-Host "    Outgoing : $($account.OutgoingServer):$($account.OutgoingPort) ($pop3OutEnc)" `
                   -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '  New IMAP settings:' -ForegroundColor Cyan
        Write-Host "    Incoming : $(Invoke-OMMigrateSanitize -Text $account.NewImapServer):$($account.NewImapPort) ($imapInEnc)" `
                   -ForegroundColor Cyan
        Write-Host "    Outgoing : $($account.OutgoingServer):$($account.OutgoingPort) ($imapOutEnc)" `
                   -ForegroundColor Cyan

        # Special notices per provider tag
        switch ($account.ProviderTag) {
            'POP3-ATTAMERITECH' {
                Write-Host ''
                Write-Host '  IMPORTANT: When Outlook prompts for your password,' `
                           -ForegroundColor Yellow
                Write-Host '  enter your Secure Mail Key -- NOT your regular email' `
                           -ForegroundColor Yellow
                Write-Host '  password. Generate one at currently.com if needed:' `
                           -ForegroundColor Yellow
                Write-Host '  https://www.att.com/acctmgmt/myprofile/overview?flow=settings' `
                           -ForegroundColor Cyan
                Write-Host '  Account Security -> Secure Mail Key -> Manage -> Generate' `
                           -ForegroundColor Yellow
            }
            'POP3-AWS' {
                Write-Host ''
                Write-Host '  IMPORTANT: This account uses AWS SES for outbound mail.' `
                           -ForegroundColor Yellow
                Write-Host '  When prompted for SMTP credentials enter your AWS SES' `
                           -ForegroundColor Yellow
                Write-Host '  IAM SMTP username and password (not your AWS console login).' `
                           -ForegroundColor Yellow
            }
            'POP3-GMAIL' {
                Write-Host ''
                Write-Host '  IMPORTANT: Gmail will open a browser window for sign-in.' `
                           -ForegroundColor Yellow
                Write-Host '  Sign in with your Google account when the browser opens.' `
                           -ForegroundColor Yellow
                Write-Host '  Approve the Outlook permission when prompted.' `
                           -ForegroundColor Yellow
            }
        }

        Write-Host ''

        # -- Backup safety check per account -------------------
        $safeEmail        = Get-SafeFileName -InputString $email
        $backupPath       = Join-Path $Global:OMMigrate.BackupPath "$safeEmail.pst"
        $safeEmailMasked  = Get-SafeFileName -InputString (Invoke-OMMigrateSanitize -Text $email)
        $backupPathMasked = Join-Path $Global:OMMigrate.BackupPath "$safeEmailMasked.pst"

        if (-not $SkipBackupCheck) {
            if (-not (Test-Path $backupPath)) {
                Write-Host ('  ' + ('-' * 54)) -ForegroundColor Red
                Write-Host "  BLOCKED: No backup PST found for $(Invoke-OMMigrateSanitize -Text $email)" -ForegroundColor Red
                Write-Host "  Expected location: $backupPathMasked" -ForegroundColor Red
                Write-Host ''
                Write-Host '  This account will NOT be touched.' -ForegroundColor Red
                Write-Host '  Re-run Script 01 to create a backup, then re-run this script.' `
                           -ForegroundColor Yellow
                Write-Host ('  ' + ('-' * 54)) -ForegroundColor Red

                Write-OMMigrateLog -Message "BLOCKED: No backup PST for $email at $backupPathMasked" `
                                   -Level ERROR

                $blockResult = $account.PSObject.Copy()
                $blockResult | Add-Member -NotePropertyName 'MigrationOutcome' `
                                          -NotePropertyValue 'FAILED' -Force
                $blockResult | Add-Member -NotePropertyName 'MigrationDetail' `
                                          -NotePropertyValue "BLOCKED -- Backup PST not found: $backupPath" `
                                          -Force
                [void]$Script:AccountResults.Add($blockResult)
                $Script:FailedCount++
                if ($Script:FinalStatus -eq 'SUCCESS') { $Script:FinalStatus = 'WARNING' }
                Update-OMMigrateProgress -MarkComplete $email
                continue
            }
        }

        # -- Per-account Y/N confirmation -----------------------
        # Default is Y -- press Enter to proceed, N to skip this account
        Update-OMMigrateProgress -SetCurrent $email

        $proceed = $true
        if (-not $Force) {
            Write-Host "  Backup PST verified: $backupPathMasked" -ForegroundColor Green
            Write-Host "  $(Format-FileSize -Bytes (Get-Item $backupPath).Length) -- your email data is safe." `
                       -ForegroundColor Green
            Write-Host ''

            $proceed = Confirm-Action `
                -Message      "REMOVE POP3 and ADD IMAP for: $(Invoke-OMMigrateSanitize -Text $email) ?" `
                -AccountEmail $email `
                -DefaultYes   $true     # Y is default -- press Enter to confirm
        }

        if (-not $proceed) {
            Write-Host ''
            Write-Host "  $(Invoke-OMMigrateSanitize -Text $email) skipped -- no changes made to this account." `
                       -ForegroundColor DarkGray
            Write-OMMigrateLog -Message "Account skipped by operator: $email" -Level INFO
            Write-AuditEntry  -Action 'ACCOUNT_SKIPPED' `
                              -AccountEmail $email `
                              -Detail 'Operator typed N at confirmation prompt' `
                              -Outcome 'SKIPPED'

            $skipResult = $account.PSObject.Copy()
            $skipResult | Add-Member -NotePropertyName 'MigrationOutcome' `
                                     -NotePropertyValue 'SKIPPED' -Force
            $skipResult | Add-Member -NotePropertyName 'MigrationDetail' `
                                     -NotePropertyValue 'Skipped by operator at Y/N prompt' -Force
            [void]$Script:AccountResults.Add($skipResult)
            $Script:SkippedCount++
            Update-OMMigrateProgress -MarkComplete $email
            continue
        }

        # -- PHASE A: Remove POP3 Account (Outlook open -- operator uses Account Settings) --
        # POP3 must be removed FIRST via Outlook's own Account Settings dialog.
        # This matches the original manual process (Google AI research, May 2026).
        # Outlook's Remove button handles all internal MAPI cleanup:
        #   - Account registry key
        #   - MAPI store key reference
        #   - Send/Receive group entries
        #   - Search folder definitions
        #   - All internal MAPI profile state
        # Direct registry deletion bypasses this cleanup and causes Outlook to
        # crash on next startup. This was confirmed by a real corruption event
        # in the May 2026 session that required ntuser.dat recovery.
        # The COM session is still active at this point -- Outlook is open and
        # the operator can access File > Account Settings without re-launching.
        # Send/Receive is suspended so the POP3 account does not sync during removal.
        Show-AccountStatus -Email  (Invoke-OMMigrateSanitize -Text $email) `
                           -Tag    $account.ProviderTag `
                           -Action 'Removing POP3 account (guided manual)...' `
                           -Status 'INFO'

        Write-OMMigrateLog -Message "Removing POP3 account (guided manual via Outlook Account Settings): $email" `
                           -Level INFO

        $removed = Remove-POP3Account `
            -EmailAddress  $email `
            -BackupPSTPath $backupPath

        if (-not $removed) {
            Write-Host ''
            Write-Host ('  ' + ('-' * 54)) -ForegroundColor Red
            Write-Host "  BLOCKED: Cannot remove POP3 account: $(Invoke-OMMigrateSanitize -Text $email)" `
                       -ForegroundColor Red
            Write-Host ''
            Write-Host '  The backup PST could not be confirmed.' -ForegroundColor Red
            Write-Host '  Re-run Script 01 to verify backups, then re-run this script.' `
                       -ForegroundColor Yellow
            Write-Host '  Your Outlook accounts have NOT been changed.' -ForegroundColor Green
            Write-Host ('  ' + ('-' * 54)) -ForegroundColor Red

            Write-OMMigrateLog -Message "POP3 removal blocked (backup check failed): $email" `
                               -Level ERROR
            Write-AuditEntry  -Action 'POP3_REMOVE_BLOCKED' `
                              -AccountEmail $email `
                              -Detail "Backup PST safety check failed -- account not touched." `
                              -Outcome 'FAILED'

            $failResult = $account.PSObject.Copy()
            $failResult | Add-Member -NotePropertyName 'MigrationOutcome' `
                                     -NotePropertyValue 'FAILED' -Force
            $failResult | Add-Member -NotePropertyName 'MigrationDetail' `
                                     -NotePropertyValue 'POP3 removal blocked -- backup PST check failed. Re-run Script 01 then retry.' `
                                     -Force
            [void]$Script:AccountResults.Add($failResult)
            $Script:FailedCount++
            $Script:FinalStatus = 'WARNING'
            Update-OMMigrateProgress -MarkComplete $email
            continue
        }

        # -- Release COM after Phase A --------------------------
        # Remove-POP3Account has completed the PST detach and signaled
        # Outlook to close. Release the COM session here to ensure
        # Outlook fully exits before Phase B reopens it.
        # Send/Receive remains suspended (state persisted in JSON) so the
        # newly added IMAP account does not immediately start a full sync.
        # Send/Receive restore happens in a single post-loop COM session
        # after ALL accounts are processed -- not per account -- to avoid
        # TCP connection exhaustion on the mail server.
        if ($Script:COMSessionOpen -and -not $Script:IsWhatIf) {
            Write-OMMigrateLog -Message (
                'Releasing COM session after Phase A -- ' +
                'Send/Receive left suspended for safe manual IMAP add.'
            ) -Level INFO

            # Do NOT call Resume-OutlookSendReceive here -- intentionally
            # left suspended. State persisted in Step02_SendReceiveState.json.

            # Null out COM account references before releasing -- these hold
            # MAPI resources that prevent Outlook from fully closing and cause
            # the 'Outlook has exhausted all shared resources' MAPI error.
            $liveComAccounts = $null
            $outlook         = $null

            try { Release-OutlookCOM } catch { }

            # Additional GC pass after nulling local COM references
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            [GC]::Collect()

            # Give Outlook process time to fully exit before reopening
            Start-Sleep -Seconds 3

            $Script:COMSessionOpen = $false
            Write-Host '  Outlook closed.' -ForegroundColor DarkGray
            Write-Host ''
        }

        # Brief pause -- let Outlook fully settle before reopening
        if (-not $Script:IsWhatIf) { Start-Sleep -Milliseconds 1500 }

        # -- Reopen Outlook for Phase B -------------------------
        # Outlook must be fully visible for the operator to use the
        # File > Add Account wizard. Use Start-Process to open it as
        # a normal foreground window -- not a background COM session.
        if (-not $Script:IsWhatIf) {
            Write-Host '  Opening Outlook for IMAP account add...' -ForegroundColor Cyan
            try {
                Start-Process 'outlook.exe'
                # Wait for Outlook to fully load before prompting operator
                Start-Sleep -Seconds 4
                Write-Host '  Outlook is open -- proceed with Add Account below.' `
                           -ForegroundColor Green
                Write-Host ''
                Write-OMMigrateLog -Message 'Outlook reopened via Start-Process for Phase B IMAP add.' `
                                   -Level INFO
            }
            catch {
                Write-OMMigrateLog -Message "Could not reopen Outlook automatically: $_" -Level WARN
                Write-Host '  Could not open Outlook automatically.' -ForegroundColor Yellow
                Write-Host '  Please open Outlook manually before continuing.' -ForegroundColor Yellow
                Write-Host ''
            }
        }

        # -- PHASE B: Add IMAP Account (operator uses File > Add Account) --
        # IMAP is added AFTER POP3 is confirmed removed. Outlook is now
        # closed. The operator reopens Outlook and uses File > Add Account
        # to add the IMAP account using the Account Setup Reference displayed
        # by Add-IMAPAccount below.
        # Outlook handles all MAPI initialization, OST creation, and search
        # folder setup internally through its own Add Account dialog.
        # This matches the original manual process exactly.
        Show-AccountStatus -Email  (Invoke-OMMigrateSanitize -Text $email) `
                           -Tag    $account.ProviderTag `
                           -Action 'Adding IMAP account (guided manual)...' `
                           -Status 'INFO'

        Write-OMMigrateLog -Message "Adding IMAP account (guided manual via File > Add Account): $email" `
                           -Level INFO

        $accountConfig = [PSCustomObject]@{
            EmailAddress   = $email
            DisplayName    = $account.DisplayName
            NewImapServer  = $account.NewImapServer
            NewImapPort    = [int]$account.NewImapPort
            NewImapSSL     = [bool]::Parse($account.NewImapSSL)
            OutgoingServer = $account.OutgoingServer
            OutgoingPort   = [int]$account.OutgoingPort
            OutgoingSSL    = [bool]::Parse($account.OutgoingSSL)
            ProviderTag    = $account.ProviderTag
            # ImapUsername defaults to email -- override in CSV if IMAP login differs
            ImapUsername   = if ($account.PSObject.Properties['ImapUsername'] -and $account.ImapUsername) { $account.ImapUsername } else { $email }
            # SmtpUsername defaults to email -- AWS SES accounts use IAM SMTP key ID
            SmtpUsername   = if ($account.PSObject.Properties['SmtpUsername'] -and $account.SmtpUsername) { $account.SmtpUsername } else { $email }
        }

        $added = Add-IMAPAccount -AccountConfig $accountConfig

        if (-not $added) {
            Write-Host ''
            Write-Host ('  ' + ('-' * 54)) -ForegroundColor Yellow
            Write-Host "  WARNING: IMAP add not confirmed: $(Invoke-OMMigrateSanitize -Text $email)" `
                       -ForegroundColor Yellow
            Write-Host ''
            Write-Host '  POP3 account has already been removed from Outlook.' -ForegroundColor Yellow
            Write-Host '  Add the IMAP account manually via Outlook File > Add Account.' `
                       -ForegroundColor Yellow
            Write-Host '  Then re-run this script -- the account will be detected as IMAP' `
                       -ForegroundColor Yellow
            Write-Host '  and skipped automatically.' -ForegroundColor Yellow
            Write-Host ('  ' + ('-' * 54)) -ForegroundColor Yellow

            Write-OMMigrateLog -Message "IMAP add not confirmed after POP3 removal: $email" `
                               -Level WARN
            Write-AuditEntry  -Action 'IMAP_ADD_NOT_CONFIRMED_AFTER_POP3_REMOVE' `
                              -AccountEmail $email `
                              -Detail "POP3 removed. IMAP add not confirmed -- manual add required." `
                              -Outcome 'WARNING'

            $warnResult = $account.PSObject.Copy()
            $warnResult | Add-Member -NotePropertyName 'MigrationOutcome' `
                                     -NotePropertyValue 'WARNING' -Force
            $warnResult | Add-Member -NotePropertyName 'MigrationDetail' `
                                     -NotePropertyValue 'POP3 removed. IMAP add not confirmed -- add manually via Outlook, then re-run.' `
                                     -Force
            [void]$Script:AccountResults.Add($warnResult)
            $Script:WarningCount++
            if ($Script:FinalStatus -eq 'SUCCESS') { $Script:FinalStatus = 'WARNING' }
            Update-OMMigrateProgress -MarkComplete $email
            continue
        }

        Write-OMMigrateLog -Message "IMAP account add confirmed by operator: $email" -Level INFO

        # -- Auto-close Outlook after Phase B -------------------
        # Operator has confirmed the account is visible. Script now
        # closes Outlook automatically via COM attach so the operator
        # does not need to do it manually.
        if (-not $Script:IsWhatIf) {
            [void](Close-OutlookIfRunning -Reason 'after Phase B IMAP add')
        }

        # Brief pause -- let Outlook fully settle after close
        if (-not $Script:IsWhatIf) { Start-Sleep -Milliseconds 1500 }

        # -- PHASE E: Auto-correct IMAP/SMTP registry credentials ------
        # After Outlook's Add Account wizard runs, it scrambles the IMAP
        # and SMTP credential fields for accounts with separate credentials
        # (e.g. POP3-AWS using AWS SES). Repair-IMAPCredentials detects and
        # corrects this automatically from the registry while Outlook is closed.
        # A .reg backup of the subkey is saved to Backups\ before and after writes.
        # Eligibility is checked per account type -- non-eligible accounts
        # are silently skipped. The operator still needs to enter the IMAP
        # password once manually in Outlook after this script completes.
        $smtpUser = if ($accountConfig.PSObject.Properties['SmtpUsername'] -and
                        $accountConfig.SmtpUsername) {
                        $accountConfig.SmtpUsername
                    } else { $email }

        $credResult = Repair-IMAPCredentials `
            -EmailAddress $email `
            -SmtpUsername $smtpUser `
            -ProviderTag  $account.ProviderTag `
            -BackupPath   $Global:OMMigrate.BackupPath

        switch ($credResult) {
            'FIXED' {
                Write-Host ''
                Write-Host '  Credential fix applied successfully.' -ForegroundColor Green
                Write-Host '  IMAP and SMTP usernames corrected in registry.' -ForegroundColor Green
                Write-Host '  You will need to enter the IMAP password once manually in Outlook.' `
                           -ForegroundColor Yellow
                Write-OMMigrateLog -Message "Credential fix applied for $email" -Level INFO
            }
            'SKIPPED' {
                Write-OMMigrateLog -Message "Credential fix skipped for $email (not applicable for this account type)." `
                                   -Level INFO
            }
            'NOT_FOUND' {
                Write-Host ''
                Write-Host '  WARNING: Credential fix could not locate the account subkey.' `
                           -ForegroundColor Yellow
                Write-Host '  Manual credential repair may be required in Outlook.' `
                           -ForegroundColor Yellow
                Write-OMMigrateLog -Message "Credential fix skipped for $email -- subkey not found." `
                                   -Level WARN
            }
            'FAILED' {
                Write-Host ''
                Write-Host '  WARNING: Credential fix encountered an error.' -ForegroundColor Yellow
                Write-Host '  Check the log for details. Manual repair may be required.' `
                           -ForegroundColor Yellow
                Write-OMMigrateLog -Message "Credential fix failed for $email -- see log for details." `
                                   -Level WARN
            }
        }

        # -- PHASE C: Send/Receive restore is intentionally deferred -------
        # Phase C (Resume Send/Receive) has been moved OUT of the per-account
        # loop and runs ONCE after all accounts are processed.
        # Reason: opening and closing a COM session per account caused rapid
        # TCP connection bursts to the mail server -- with 25+ accounts this
        # exhausted server connection limits and required a mail server reboot
        # to recover. A single post-loop COM session eliminates this problem.
        # See: "Post-Loop Phase C" block below, just before Step 5.

        # -- PHASE D: Operator Verifies Account Visible ------------------
        # Confirms the new IMAP account was visible in the Outlook folder
        # pane before the script closed it. Outlook is now closed.
        # Credential fix has already been applied automatically above.
        # The only remaining manual step is entering the IMAP password
        # once via Outlook Send/Receive Groups after this script completes.
        Write-Host ''
        Write-Host ('  ' + ('-' * 54)) -ForegroundColor DarkCyan
        Write-Host "  VERIFY: Confirm $(Invoke-OMMigrateSanitize -Text $email) in Outlook" `
                   -ForegroundColor White
        Write-Host ('  ' + ('-' * 54)) -ForegroundColor DarkCyan
        Write-Host ''
        Write-Host '  Confirm you saw both of the following before pressing C:' -ForegroundColor Gray
        Write-Host "    1. $(Invoke-OMMigrateSanitize -Text $email) visible in the left folder pane" `
                   -ForegroundColor Gray
        Write-Host '    2. The old POP3 account is no longer listed' -ForegroundColor Gray
        Write-Host ''
        Write-Host '  NOTE: Do not attempt to send or receive mail yet.' -ForegroundColor Yellow
        Write-Host '  Reopen Outlook (if it is not already open) and enter' -ForegroundColor Yellow
        Write-Host '  your IMAP password via:' -ForegroundColor Yellow
        Write-Host '  Send/Receive Groups > Edit Group > Account Properties' -ForegroundColor White
        Write-Host '  IMAP and SMTP usernames have been corrected automatically.' `
                   -ForegroundColor Green
        Write-Host '  NOTE: Do not close Outlook -- it will be opened' `
                   -ForegroundColor DarkGray
        Write-Host '  automatically by the script when ready.' -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '  PASSWORD ENTRY NOTE:' -ForegroundColor Yellow
        Write-Host '  You may be prompted for your IMAP password twice:' -ForegroundColor Yellow
        Write-Host '    1. A popup dialog when Outlook first opens -- enter' `
                   -ForegroundColor Gray
        Write-Host '       your IMAP password and click OK.' -ForegroundColor Gray
        Write-Host '    2. Account Properties in Send/Receive Groups -- enter' `
                   -ForegroundColor Gray
        Write-Host '       your IMAP password again and click OK.' -ForegroundColor Gray
        Write-Host '  This is expected and only happens once per account.' `
                   -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '  If the account was visible in the folder pane -- type Y' -ForegroundColor Green
        Write-Host '  If the account was missing or Outlook showed an error -- type N' `
                   -ForegroundColor Yellow
        Write-Host '  Typing N flags this account for review but continues to the next' `
                   -ForegroundColor DarkGray
        Write-Host ''

        # Default Y -- operator confirmed account visible in folder pane
        $connected = Confirm-Action `
            -Message      "Was $(Invoke-OMMigrateSanitize -Text $email) visible in the Outlook folder pane?" `
            -AccountEmail $email `
            -DefaultYes   $true

        if ($connected) {
            $outcomeStatus    = 'SUCCESS'
            $connectionDetail = 'Operator confirmed account visible in Outlook folder pane'

            Show-AccountStatus -Email  (Invoke-OMMigrateSanitize -Text $email) `
                               -Tag    $account.ProviderTag `
                               -Action "Migrated to IMAP -- $(Invoke-OMMigrateSanitize -Text $accountConfig.NewImapServer)" `
                               -Status 'OK'

            $Script:MigratedCount++

            # -- Auto-update migration_accounts.csv -----------------
            # Write the post-conversion values automatically so the operator
            # does not need to edit the CSV manually after Script 02.
            # Uses the same file-lock-checked write as Script 03.
            # WhatIf: logs only, no write.
            Write-OMMigrateLog -Message (
                "Auto-updating migration_accounts.csv for $email : " +
                "MigrationAction=FOLDER-ONLY | ProviderTag=IMAP-CONVERTED | AccountType=IMAP"
            ) -Level INFO
            Update-AccountMigrationAction `
                -EmailAddress   $email `
                -NewAction      'FOLDER-ONLY' `
                -NewProviderTag 'IMAP-CONVERTED' `
                -NewAccountType 'IMAP'
        }
        else {
            $outcomeStatus    = 'WARNING'
            $connectionDetail = 'Operator indicated connection not confirmed -- flagged for review'

            Write-Host ''
            Write-Host ('  ' + ('-' * 54)) -ForegroundColor Yellow
            Write-Host "  $(Invoke-OMMigrateSanitize -Text $email) flagged -- connection not confirmed." `
                       -ForegroundColor Yellow
            Write-Host ''
            Write-Host '  What this means:' -ForegroundColor Gray
            Write-Host '  The IMAP account was added but Outlook may not have' -ForegroundColor Gray
            Write-Host '  connected successfully yet. This can happen when:' -ForegroundColor Gray
            Write-Host '    - The password dialog was dismissed without entering a password' -ForegroundColor Gray
            Write-Host '    - The server took longer than expected to respond' -ForegroundColor Gray
            Write-Host '    - Network connectivity was interrupted briefly' -ForegroundColor Gray
            Write-Host ''
            Write-Host '  What to do:' -ForegroundColor Yellow
            Write-Host '  - Check Outlook now and try entering your password if prompted' -ForegroundColor Gray
            Write-Host '  - The script will continue to the next account' -ForegroundColor Gray
            Write-Host '  - Review this account in Outlook before running Script 03' -ForegroundColor Gray
            Write-Host ('  ' + ('-' * 54)) -ForegroundColor Yellow
            Write-Host ''

            Show-AccountStatus -Email  (Invoke-OMMigrateSanitize -Text $email) `
                               -Tag    $account.ProviderTag `
                               -Action 'Added -- connection not confirmed -- review before Script 03' `
                               -Status 'WARN'

            if ($Script:FinalStatus -eq 'SUCCESS') {
                $Script:FinalStatus = 'WARNING'
            }
            $Script:WarningCount++
        }

        Write-AuditEntry -Action 'IMAP_ADDED' `
                         -AccountEmail $email `
                         -Detail (
                             "Server=$($accountConfig.NewImapServer):$($accountConfig.NewImapPort) | " +
                             "SMTP=$($accountConfig.OutgoingServer):$($accountConfig.OutgoingPort) | " +
                             "$connectionDetail"
                         ) `
                         -Outcome $outcomeStatus

        $successResult = $account.PSObject.Copy()
        $successResult | Add-Member -NotePropertyName 'MigrationOutcome' `
                                    -NotePropertyValue $outcomeStatus -Force
        $successResult | Add-Member -NotePropertyName 'MigrationDetail' `
                                    -NotePropertyValue $connectionDetail -Force
        [void]$Script:AccountResults.Add($successResult)

        # Save checkpoint and mark complete
        Save-OMMigrateCheckpoint -ExitType 'NORMAL' -Reason 'Account completed'
        Update-OMMigrateProgress -MarkComplete $email

        Write-OMMigrateLog -Message (
            "Account complete: $email | Outcome=$outcomeStatus | $connectionDetail"
        ) -Level INFO

        # Brief pause between accounts -- skip after last account
        if ($accountNumber -lt $totalAccounts) {
            Write-Host ''
            Write-Host '  Preparing next account...' -ForegroundColor DarkGray
            if (-not $Script:IsWhatIf) { Start-Sleep -Seconds 2 }
        }
    }

    # ----------------------------------------------------------
    #  POST-LOOP PHASE C -- Restore Send/Receive (single COM session)
    # ----------------------------------------------------------
    # Send/Receive was suspended at the start of Phase A and intentionally
    # left suspended through all account processing. Now that all accounts
    # are done, restore it in a single COM session.
    #
    # WHY POST-LOOP (not per-account):
    #   Opening and closing a COM session per account caused rapid TCP
    #   connection bursts to the mail server. With 25+ accounts this
    #   exhausted server-side connection limits and required a full mail
    #   server reboot to recover (confirmed in May 2026 live run).
    #   One COM session here -- regardless of account count -- eliminates
    #   that problem entirely.
    #
    # Resume-OutlookSendReceive reads Step02_SendReceiveState.json to
    # determine which groups were active before suspension and restores
    # only those. Groups already disabled before migration stay disabled.
    #
    # TWO PATHS:
    #   1. COM session still open ($Script:COMSessionOpen = $true):
    #      All accounts were auto-skipped (already IMAP). Outlook never
    #      closed between phases. Use the existing session directly --
    #      do NOT try to Connect-OutlookCOM again or it will fail with
    #      "Outlook is already running".
    #   2. COM session closed ($Script:COMSessionOpen = $false):
    #      Normal migration path -- Outlook was closed between Phase A
    #      and Phase B. Operator may have left Outlook open after manually
    #      adding the IMAP account in Phase B. Use AllowRunning=$true so
    #      Connect-OutlookCOM attaches to the existing instance if running,
    #      or launches a fresh instance if Outlook is closed.
    if (-not $Script:IsWhatIf) {
        Write-Host ''
        Write-Host '  Restoring Outlook Send/Receive groups...' -ForegroundColor Cyan

        if ($Script:COMSessionOpen) {
            # Path 1 -- existing COM session is still active (all-skip path)
            # Call Resume directly -- no new Connect-OutlookCOM needed
            Write-OMMigrateLog -Message 'Post-loop Phase C: using existing COM session to restore Send/Receive state...' `
                               -Level INFO
            try {
                Resume-OutlookSendReceive | Out-Null
                Write-OMMigrateLog -Message 'Post-loop Phase C: Send/Receive state restored via existing session.' `
                                   -Level INFO
            }
            catch {
                Write-OMMigrateLog -Message "Post-loop Phase C: Send/Receive restore failed on existing session: $_" `
                                   -Level WARN
                Write-Host '  WARNING: Send/Receive restore encountered an error.' -ForegroundColor Yellow
                Write-Host '  Verify manually: Outlook > File > Options > Advanced > Send/Receive.' `
                           -ForegroundColor Yellow
            }
            # COM session will be released by the finally block as normal
        }
        else {
            # Path 2 -- COM session was closed between phases (normal migration path)
            # Use AllowRunning=$true -- operator may have left Outlook open after
            # Phase B manual IMAP add. Attach to existing instance if running,
            # launch fresh if not.
            Write-OMMigrateLog -Message 'Post-loop Phase C: opening COM session to restore Send/Receive state...' `
                               -Level INFO
            try {
                $outlookResume = Connect-OutlookCOM -AllowRunning $true
                if ($outlookResume) {
                    $Script:COMSessionOpen = $true
                    Resume-OutlookSendReceive | Out-Null
                    $outlookResume = $null
                    [GC]::Collect()
                    [GC]::WaitForPendingFinalizers()
                    [GC]::Collect()
                    try { Release-OutlookCOM } catch { }
                    # Allow mail server connections to settle before script ends
                    Start-Sleep -Seconds 3
                    $Script:COMSessionOpen = $false
                    Write-OMMigrateLog -Message 'Post-loop Phase C: Send/Receive state restored.' -Level INFO
                }
                else {
                    Write-OMMigrateLog -Message (
                        'Post-loop Phase C: Could not open COM session for Send/Receive restore -- ' +
                        'verify Send/Receive settings in Outlook manually via ' +
                        'File > Options > Advanced > Send/Receive.'
                    ) -Level WARN
                    Write-Host '  WARNING: Could not restore Send/Receive groups automatically.' `
                               -ForegroundColor Yellow
                    Write-Host '  Verify manually: Outlook > File > Options > Advanced > Send/Receive.' `
                               -ForegroundColor Yellow
                }
            }
            catch {
                Write-OMMigrateLog -Message "Post-loop Phase C: Send/Receive restore failed: $_" `
                                   -Level WARN
                Write-Host '  WARNING: Send/Receive restore encountered an error.' -ForegroundColor Yellow
                Write-Host '  Verify manually: Outlook > File > Options > Advanced > Send/Receive.' `
                           -ForegroundColor Yellow
            }
        }
    }
    else {
        Write-OMMigrateLog -Message 'WhatIf: Would restore Send/Receive groups in post-loop Phase C.' `
                           -Level INFO -WhatIfPrefix
    }

    # Final status
    if ($Script:FailedCount -gt 0 -and $Script:MigratedCount -eq 0 -and
        $Script:AutoSkippedCount -eq 0) {
        $Script:FinalStatus = 'FAILED'
    }
    elseif ($Script:FailedCount -gt 0) {
        $Script:FinalStatus = 'WARNING'
    }

    # -- Auto-open Outlook for IMAP password entry -----------------
    # After all accounts are migrated and Send/Receive is restored,
    # open Outlook automatically so the operator can enter the IMAP
    # password immediately without having to open it manually.
    # Only runs when at least one account was successfully migrated
    # and not in WhatIf mode.
    if ($Script:MigratedCount -gt 0 -and -not $Script:IsWhatIf) {
        Write-Host ''
        Write-Host '  Opening Outlook for IMAP password entry...' -ForegroundColor Cyan
        Write-Host '  Send/Receive Groups > Edit Group > Account Properties' -ForegroundColor White
        try {
            Start-Process 'outlook.exe'
            Write-OMMigrateLog -Message 'Outlook opened automatically for IMAP password entry.' -Level INFO
        }
        catch {
            Write-OMMigrateLog -Message "Could not open Outlook automatically: $_" -Level WARN
            Write-Host '  Could not open Outlook automatically -- please open it manually.' `
                       -ForegroundColor Yellow
        }
    }


    # ----------------------------------------------------------
    #  STEP 5 -- Generate Report and Manifest
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Generating Report and Manifest' -Step '5 of 5'

    Write-Host '  Conversion Summary:' -ForegroundColor White
    Write-Host ("    {0,-28}: {1,3}" -f 'Migrated to IMAP', $Script:MigratedCount) -ForegroundColor $(
        if ($Script:MigratedCount -gt 0) { 'Green' } else { 'Cyan' }
    )
    Write-Host ("    {0,-28}: {1,3}" -f 'Already IMAP (no action)', $Script:AlreadyImapCount) -ForegroundColor Green
    Write-Host ("    {0,-28}: {1,3}" -f 'Live IMAP detected (re-run)', $Script:AutoSkippedCount) -ForegroundColor Green
    Write-Host ("    {0,-28}: {1,3}" -f 'Exchange (no action)', $Script:ExchangeCount) -ForegroundColor Cyan
    Write-Host ("    {0,-28}: {1,3}" -f 'Skipped by operator', $Script:SkippedCount) -ForegroundColor Yellow
    Write-Host ("    {0,-28}: {1,3}" -f 'Connection unconfirmed', $Script:WarningCount) -ForegroundColor $(
        if ($Script:WarningCount -gt 0) { 'Yellow' } else { 'Green' }
    )
    Write-Host ("    {0,-28}: {1,3}" -f 'Failed', $Script:FailedCount) -ForegroundColor $(
        if ($Script:FailedCount -gt 0) { 'Red' } else { 'Green' }
    )
    Write-Host ("    {0,-28}: {1,3}" -f 'Not marked (other)', $Script:OtherNotMarkedCount) -ForegroundColor Cyan
    Write-Host "    $('-' * 36)" -ForegroundColor DarkGray
    $Script:TotalAccountCount = $Script:MigratedCount + $Script:AlreadyImapCount +
                                $Script:AutoSkippedCount + $Script:ExchangeCount +
                                $Script:SkippedCount + $Script:WarningCount +
                                $Script:FailedCount + $Script:OtherNotMarkedCount
    Write-Host ("    {0,-28}: {1,3}" -f 'Total accounts identified', $Script:TotalAccountCount) -ForegroundColor White
    Write-Host ''

    $Script:ReportFile = New-MigrationReport -Accounts @($Script:AccountResults | Sort-Object EmailAddress)
    Write-Host "  Migration Report: $Script:ReportFile" -ForegroundColor Green

    if ($Script:FailedCount -gt 0) {
        $rollbackAccounts = @($Script:AccountResults | Where-Object {
            $_.MigrationOutcome -eq 'FAILED'
        }) | ForEach-Object {
            $ra = $_.PSObject.Copy()
            $ra | Add-Member -NotePropertyName 'LastCompletedStep' `
                             -NotePropertyValue 'Script 02 -- Partial' -Force
            $ra | Add-Member -NotePropertyName 'RollbackOutcome' `
                             -NotePropertyValue 'FAILED' -Force
            $ra | Add-Member -NotePropertyName 'RecoveryNote' `
                             -NotePropertyValue $_.MigrationDetail -Force
            $ra
        }

        $rollbackFile = New-RollbackReport `
            -Accounts $rollbackAccounts `
            -Reason   "$Script:FailedCount account(s) failed during Script 02"
        Write-Host "  Rollback Report : $rollbackFile" -ForegroundColor Yellow
    }

    Write-StepManifest -Step 2 -Status $Script:FinalStatus -Data @{
        MigratedCount        = $Script:MigratedCount
        AlreadyImapCount     = $Script:AlreadyImapCount
        ExchangeCount        = $Script:ExchangeCount
        AutoSkippedCount     = $Script:AutoSkippedCount
        FailedCount          = $Script:FailedCount
        SkippedCount         = $Script:SkippedCount
        WarningCount         = $Script:WarningCount
        OtherNotMarkedCount  = $Script:OtherNotMarkedCount
        NotMarkedCount       = $Script:NotMarkedCount
        TotalAccountCount    = $Script:TotalAccountCount
        TotalProcessed       = $Script:AccountResults.Count
        ReportFile        = $Script:ReportFile
        CompletedAccounts = @($Script:AccountResults |
                              Where-Object { $_.MigrationOutcome -eq 'SUCCESS' } |
                              ForEach-Object { $_.EmailAddress })
    }

    Write-Host ''
    if ($Script:FinalStatus -eq 'SUCCESS') {
        Write-Host '  All accounts processed successfully.' -ForegroundColor Green
        Write-Host '  Review the Migration Report, verify Outlook, then run:' -ForegroundColor Green
        Write-Host '  .\Scripts\OMMigrate-03-Restore.ps1' -ForegroundColor White
    }
    elseif ($Script:FinalStatus -eq 'WARNING') {
        Write-Host '  Completed with warnings.' -ForegroundColor Yellow
        Write-Host '  Review flagged accounts in Outlook before running Script 03.' `
                   -ForegroundColor Yellow
        Write-Host '  When ready: .\Scripts\OMMigrate-03-Restore.ps1' -ForegroundColor White
    }
    else {
        Write-Host '  Conversion FAILED.' -ForegroundColor Red
        Write-Host '  Review the Rollback Report. DO NOT run Script 03 yet.' -ForegroundColor Red
        Write-Host '  Your backup PST files are safe in the Backups\ folder.' `
                   -ForegroundColor Yellow
    }
    Write-Host ''

}
catch {
    Write-OMMigrateLog -Message "FATAL ERROR in Script 02: $_" -Level ERROR
    Write-OMMigrateLog -Message "Stack trace: $($_.ScriptStackTrace)" -Level DEBUG
    $Script:FinalStatus = 'FAILED'

    Write-Host ''
    Write-Host '  FATAL ERROR:' -ForegroundColor Red
    Write-Host "  $_" -ForegroundColor Red
    Write-Host ''
    Write-Host '  Open Outlook and verify your account list.' -ForegroundColor Yellow
    Write-Host '  Backup PST files are safe in the Backups\ folder.' -ForegroundColor Yellow
    Write-Host "  Log: $($Global:OMMigrate.RunLogFile)" -ForegroundColor Gray

    if ($Script:AccountResults.Count -gt 0) {
        try {
            New-RollbackReport `
                -Accounts @($Script:AccountResults) `
                -Reason   "Fatal error in Script 02: $_" | Out-Null
        }
        catch { }
    }
}
finally {
    if ($Script:COMSessionOpen -and -not $Script:IsWhatIf) {
        # Release-OutlookCOM logs its own "Releasing Outlook COM session..." message internally.
        # Do not add a duplicate log line here -- it would appear twice in the log.
        try { Resume-OutlookSendReceive | Out-Null } catch { }
        try { Release-OutlookCOM } catch { }
        $Script:COMSessionOpen = $false
    }

    # Only save checkpoint if work actually started (COM session was opened)
    # Avoids writing a checkpoint when operator declined at pre-flight
    if ($Script:COMSessionOpen -or $Script:AccountResults.Count -gt 0) {
        Save-OMMigrateCheckpoint -ExitType 'NORMAL' -Reason 'Script 02 session ending'
    }

    # Clean up checkpoint file after a fully successful session.
    # A clean run means all selected accounts completed -- no resume needed.
    # Failed or warned sessions keep the checkpoint so the operator can resume.
    if ($Script:FinalStatus -eq 'SUCCESS') {
        $checkpointPath = Join-Path $Global:OMMigrate.ManifestPath 'Step02_Checkpoint.json'
        if (Test-Path $checkpointPath) {
            try {
                Remove-Item $checkpointPath -Force -ErrorAction Stop
                Write-OMMigrateLog -Message 'Step02_Checkpoint.json removed after successful session.' `
                                   -Level DEBUG
            }
            catch {
                Write-OMMigrateLog -Message "Could not remove Step02_Checkpoint.json: $_" -Level INFO
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
