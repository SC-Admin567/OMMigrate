#Requires -Version 5.1
<#
.SYNOPSIS
    OMMigrate-Outlook.psm1 -- Outlook COM API Interaction Layer

.DESCRIPTION
    Provides all functions that interact with Microsoft Outlook through
    the Outlook COM Object Model (also known as the Outlook Object Model
    or OOM). This module is used by Scripts 00 through 04.

    Script 00 (Discovery) uses this module READ-ONLY to:
        - Enumerate all email accounts in the Outlook session
        - Build the complete folder tree for each account
        - Inventory all Outlook Rules (name, conditions, actions, targets)
        - Corroborate registry data with live Outlook account objects

    Script 01 (Backup) uses this module to:
        - Close Outlook automatically if found running, via the shared
          Close-OutlookIfRunning routine, before PST files are copied
          (Script 01 does not open its own Outlook COM session)

    Script 02 (Convert) uses this module to:
        - Guide operator through removing POP3 accounts via Outlook
          Account Settings (File > Account Settings > Remove)
        - Guide operator through adding IMAP accounts via Outlook
          Add Account dialog (File > Add Account > Manual setup)
        - Verify new IMAP account connectivity after add

    Script 03 (Restore) uses this module to:
        - Create server-side IMAP folders per folder_map.csv
        - Open and read backup PST files
        - Copy email items and folder structures between stores
        - Update Outlook Rules to point to new folder locations

    Script 04 (Artifacts) uses this module to:
        - Close Outlook automatically if found running, via the shared
          Close-OutlookIfRunning routine, before its own Outlook COM
          session is launched to copy Calendar/Contacts/Tasks/Notes/
          Journal items

    SHARED OUTLOOK-CLOSE ROUTINE:
        Close-OutlookIfRunning is the single shared routine every script
        calls when it needs Outlook closed -- whether as a pre-flight
        check before launching a COM session, mid-script before a new
        phase's session, or as post-completion cleanup. It always
        attempts a graceful Quit() (via Connect-OutlookCOM -AllowRunning
        + Release-OutlookCOM) before falling back to a forced process
        kill, so pending writes (e.g. PR_RW_RULES_STREAM after a rules
        import/Apply, or PST buffers) get a chance to flush to disk
        first. See the function's own comment-based help for details.
        Do not reintroduce script-local inline close logic -- call this
        routine instead so any future fix only needs to be made once.

    COM API SAFETY NOTES:
        - Outlook must be CLOSED before this module launches it
        - This module launches Outlook in a controlled COM session
        - The COM session is always released in finally blocks
        - No Outlook dialogs are suppressed -- security prompts must
          be answered by the operator
        - All destructive operations (account remove, folder create)
          require a preceding Confirm-Action Y/N gate

    ARCHITECTURE NOTE:
        Outlook's COM API is single-threaded. All COM calls in this
        module run synchronously on the calling thread. Do not attempt
        to parallelize COM calls -- Outlook's object model is not
        thread-safe.

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
        PowerShell      : 5.1 or higher
        Windows         : 10 or 11 (64-bit)
        Outlook         : Classic Outlook 2016 / 2019 / 2021
        OMMigrate-Core  : Must be imported before this module

    COM Interop:
        Uses [System.Runtime.InteropServices.Marshal]::ReleaseComObject
        to properly release all COM objects and prevent memory leaks.
        All COM object variables are released in finally blocks.

    Outlook Must Be Closed:
        Test-OMMigrateEnvironment (in OMMigrate-Core) verifies Outlook
        is not running before any script proceeds. This module will
        launch Outlook in a COM-controlled session. The operator must
        not manually open Outlook while a script is running.
#>

Set-StrictMode -Version Latest


# ============================================================
#  MODULE CONSTANTS -- Outlook COM Enumeration Values
#  These are official Outlook Object Model constant values.
#  Documented at:
#  https://docs.microsoft.com/en-us/office/vba/api/overview/outlook
# ============================================================

# OlAccountType enumeration
$Script:OlAccountType = @{
    Exchange      = 0
    IMAP          = 1
    POP3          = 2
    HTTP          = 3
    OtherAccount  = 5
}

# OlDefaultFolders enumeration (most common)
$Script:OlDefaultFolders = @{
    olFolderInbox          = 6
    olFolderOutbox         = 4
    olFolderSentMail       = 5
    olFolderDeletedItems   = 3
    olFolderDrafts         = 16
    olFolderJunk           = 23
    olFolderCalendar       = 9
    olFolderContacts       = 10
    olFolderTasks          = 13
    olFolderNotes          = 12
}

# OlItemType enumeration
$Script:OlItemType = @{
    olMailItem        = 0
    olAppointmentItem = 1
    olContactItem     = 2
    olTaskItem        = 3
    olJournalItem     = 4
    olNoteItem        = 5
    olPostItem        = 6
}

# OlRuleConditionType enumeration
$Script:OlRuleConditionType = @{
    olConditionFrom           = 0
    olConditionSubject        = 1
    olConditionBody           = 2
    olConditionImportance     = 3
    olConditionSensitivity    = 4
    olConditionBodyOrSubject  = 5
    olConditionSenderAddress  = 12
    olConditionRecipientAddress = 13
    olConditionAccount        = 15
}

# OlRuleActionType enumeration
$Script:OlRuleActionType = @{
    olRuleActionMoveToFolder    = 1
    olRuleActionCopyToFolder    = 2
    olRuleActionDelete          = 3
    olRuleActionDeletePermanently = 4
    olRuleActionMarkAsRead      = 6
    olRuleActionMarkImportance  = 7
    olRuleActionForwardToRecipient = 8
    olRuleActionForwardAsAttachment = 9
    olRuleActionReplyWithTemplate = 10
    olRuleActionCustomAction    = 12
    olRuleActionAssignToCategory = 14
    olRuleActionNotifyRead      = 15
    olRuleActionNotifyDelivery  = 16
    olRuleActionCcToRecipient   = 17
    olRuleActionStopProcessingRules = 18
    olRuleActionRunScript       = 20
    olRuleActionSendToOneNoteAndDelete = 21
    olRuleActionClearCategories = 22
    olRuleActionDeferred        = 23
}

# OlStoreType enumeration
$Script:OlStoreType = @{
    olExchangeMailbox  = 1
    olExchangePublicFolder = 2
    olLocalStore       = 3    # PST
    olExchangePrivateStore = 4
}

# OlSaveAsType for PST export
$Script:OlSaveAsType = @{
    olTXT   = 0
    olRTF   = 1
    olMSG   = 3
    olDoc   = 4
    olHTML  = 5
    olVCard = 6
    olVCal  = 7
    olICal  = 8
}

# Export format for PST backup
$Script:OlExportType = @{
    olExportPST = 0
}


# ============================================================
#  REGION: COM SESSION MANAGEMENT
# ============================================================

# Module-level COM object references
# Stored at module scope so Release-OutlookCOM can clean up
$Script:OutlookApp      = $null
$Script:OutlookNamespace = $null
$Script:COMObjectsToRelease = [System.Collections.Generic.List[object]]::new()
$Script:MountedStoreNames = $null   # Cached store display names for Remove-StorePrefix -- reset on COM release


function ConvertTo-NormalizedSenderDomains {
    <#
    .SYNOPSIS
        Normalizes a raw SendersDomain CSV cell into an array of one or
        more clean, individual sender-address values for use as
        SenderAddress.Address array entries.

    .DESCRIPTION
        Added 2026-07-02, Administrator. REPLACED 2026-07-06, Administrator (Issue 3 fix) --
        the original version below is what this function used to do; kept
        here for history since the new logic is a genuinely different
        validation approach, not a tweak of the old one:

          OLD (removed): unconditionally stripped every "@" and "&"
          character from the ENTIRE raw value before splitting on
          space/semicolon. This corrupted any valid full email address in
          the cell -- e.g. "joe_black@123.com" became "joe_black123.com"
          -- because the strip ran across the whole string rather than
          per-token, with no concept of "this token is a real email, leave
          it alone."

        NEW validation approach: split into tokens FIRST (space-delimited
        ONLY -- see Split rules below, semicolon support was removed later
        the same night), then validate EACH token independently as one of
        two acceptable shapes:
          (a) a full email address (local-part@domain). Local-part
              character set NARROWED (2026-07-06, Administrator, 2nd pass) from
              Administrator's initially-supplied RFC 5322-based regex down to just
              ".", "_", "+", "-" -- the lowest common denominator that
              works cleanly across every layer this value passes through:
              PowerShell (this module), Python (used elsewhere in this
              project's tooling), and the VBA macro (Module3.bas), plus
              this project's own conventions (comma-split CSV parsing in
              the macro, "|"-joined log output elsewhere in this module,
              path separators throughout). RFC 5322 technically permits a
              much wider set (!#$%&'*/=?^`{|}~), but per Administrator's direction,
              several of those are ALSO syntactically meaningful in one or
              more of PowerShell ($, `, {, }, |), Python, or VBA -- and per
              Administrator's own research, real-world providers (Gmail, Outlook,
              Yahoo) only actually permit "., _, -" (plus "+" widely
              supported for subaddressing) at signup anyway, so this
              narrower set loses essentially no real-world coverage while
              removing every character that could cause a parsing/
              interpolation surprise somewhere in this pipeline.
          (b) a bare word/domain fragment (no "@" at all) -- e.g.
              "example-provider", "amazon", "aws" -- matched separately since
              these are not email addresses and don't need local-part/
              domain structure, just ordinary word characters, dots, and
              hyphens (the same safe domain-label characters as the email
              pattern's domain half).
        A token matching NEITHER shape is dropped with a debug log line
        rather than passed through corrupted or blocked entirely -- this
        differs from the old behavior (which always produced SOME output,
        even if mangled) but is safer: an admin sees a dropped token is
        missing (and can fix the CSV) rather than silently getting a
        corrupted or misleading SenderAddress condition deployed to a live
        Outlook rule.

        Split rules -- ONLY space is a supported delimiter (semicolon
        REMOVED, 2026-07-06, Administrator, later same night -- previously treated
        identically to space, but per Administrator's explicit direction semicolon
        is rarely used in practice and dropping it simplifies the rule to
        one delimiter, consistent with comma's exclusion below):
          1. A space between words means the cell holds multiple distinct
             values -- split into separate entries (e.g. "amazon aws"
             becomes two entries: "amazon" and "aws"). Outlook's own
             SenderAddress.Address array evaluates multiple string
             elements as a logical OR automatically (confirmed via
             Microsoft's AddressRuleCondition.Address documentation) --
             this function only needs to produce the separate array
             elements, never a literal " or " string.
          2. A semicolon is NO LONGER treated as a delimiter. A semicolon
             appearing inside a cell just becomes part of whatever token
             it's in, and will typically fail both the email and bare-word
             patterns (semicolon is not in either character set) and be
             dropped as invalid -- same outcome as any other unsupported
             character, no special-casing needed.
          3. Comma is NOT a supported SendersDomain separator -- Export-Csv
             auto-quotes any CSV cell containing a literal comma, but the
             companion VBA macro (Module3.bas) reads rules_inventory.csv
             with a plain, non-quote-aware line split on comma, so a comma
             inside a SendersDomain cell would silently misalign every
             column after it. A defensive comma-to-space scrub (below)
             still runs before splitting, as a safety net against a manual
             CSV edit introducing one -- this is unrelated to the
             delimiter rule itself.

    .PARAMETER RawValue
        The raw, as-typed SendersDomain cell value from rules_inventory.csv.

    .OUTPUTS
        [string[]] One or more validated, trimmed, lowercased tokens --
        either full email addresses or bare word/domain fragments. Tokens
        that match neither valid shape are dropped (with a DEBUG log line
        naming the dropped token). Empty/whitespace-only entries are
        dropped silently (not logged, since this is normal -- e.g. a
        double space between tokens). Returns an empty array if RawValue
        is blank or normalizes to nothing usable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$RawValue
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) { return @() }

    # Defensive comma scrub (2026-07-02, Administrator): Export-RulesToCSV already
    # scrubs commas to spaces at CSV WRITE time (end of Script 00), but an
    # admin can manually edit rules_inventory.csv in the window between
    # Script 00 finishing and Script 03 running, potentially typing a comma
    # despite the documented rule against it. This is the actual
    # consumption/runtime choke point (called from both Invoke-
    # DeployConsolidatedRules's Step 1 grouping and Step 4 timestamp
    # writeback), so catching a stray comma here -- not just at Script 00's
    # write time -- protects against a manual edit made after Script 00
    # already wrote the file. Comma is still never a DOCUMENTED/supported
    # separator (see header comment) -- this is a safety net, not a new
    # feature -- so it is simply treated the same as a space here rather
    # than being split any differently.
    $RawValue = $RawValue.Replace(',', ' ')

    # Split FIRST (Issue 3 fix, 2026-07-06, Administrator) -- unlike the old version,
    # nothing is stripped from the raw value before splitting, so a valid
    # email's "@" survives into its own token intact.
    #
    # Semicolon REMOVED as a delimiter (2026-07-06, Administrator, later same night):
    # per Administrator's explicit direction, only space is a supported
    # SendersDomain delimiter going forward -- semicolon is rarely used in
    # practice and its removal simplifies the rule to one delimiter,
    # matching comma's exclusion for the same reason (CSV-safety). A
    # semicolon inside a token is no longer converted to a space at all;
    # it just becomes part of whatever token it's in, and will typically
    # fail both the email and bare-word patterns and be dropped as
    # invalid, same as any other unsupported character would be.
    $parts = $RawValue -split '\s+'

    # Validation patterns (Issue 3 fix, 2026-07-06, Administrator, narrowed 2nd pass).
    # Local-part limited to ".", "_", "+", "-" -- the lowest common
    # denominator safe across PowerShell/Python/VBA and this project's own
    # conventions (comma-split CSV parsing, "|"-joined log output, path
    # separators). See .DESCRIPTION above for the full rationale.
    $emailPattern = '^[a-zA-Z0-9_+.-]+(?:\.[a-zA-Z0-9_+.-]+)*@(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}$'
    # Single-quoted for consistency with the rest of this function's string
    # style, though the narrowed character set above no longer strictly
    # requires it (no backtick or single-quote characters remain in the
    # pattern -- both were dropped when the character set was narrowed).
    # Bare word/domain fragment: ordinary word characters, dots, hyphens,
    # and underscores, no "@" at all -- covers plain terms like
    # "example-provider", "amazon", or "smith_landscaping" that are not
    # full email addresses. Underscore added (2026-07-06, Administrator, live macro
    # test) -- originally omitted from this pattern despite being one of
    # the 4 approved local-part characters for the EMAIL pattern above,
    # which caused a real false-rejection: "smith_landscaping" (a
    # legitimate folder-name-derived SendersDomain token, consistent with
    # this project's existing naming convention) was being dropped as
    # invalid. Must not start/end with a dot or hyphen, matching the same
    # safe-label convention as the domain half of the email pattern above.
    $wordPattern = '^[a-zA-Z0-9]([a-zA-Z0-9._-]*[a-zA-Z0-9])?$'

    $result = @()
    foreach ($part in $parts) {
        $trimmedPart = $part.Trim().ToLower()
        if ([string]::IsNullOrWhiteSpace($trimmedPart)) { continue }

        if ($trimmedPart -match $emailPattern -or $trimmedPart -match $wordPattern) {
            $result += $trimmedPart
        }
        else {
            # Dropped, not corrupted or passed through (Issue 3 fix,
            # 2026-07-06, Administrator) -- see .DESCRIPTION above for rationale.
            Write-OMMigrateLog -Message "ConvertTo-NormalizedSenderDomains: dropped invalid SendersDomain token '$trimmedPart' (matched neither email nor bare-word pattern)." -Level DEBUG
        }
    }

    return @($result)
}


function Test-MeaningfulRuleSummaryText {
    <#
    .SYNOPSIS
        Checks whether a Conditions/Actions summary string is meaningful
        enough to trust as real content, versus blank, a known placeholder,
        or too short/cryptic to mean anything to an admin reading the CSV.

    .DESCRIPTION
        Added 2026-07-06, Administrator (Issues 4/5, refined 3rd pass). Conditions
        and Actions are purely informational/display columns (confirmed
        never consumed by rule-building logic -- only TargetFolderPath,
        TargetFolderEntryID, and SendersDomain actually drive what gets
        deployed). Administrator's direction: these columns should update to
        reflect genuine current state on a rerun (so a manually-added
        Outlook UI condition, e.g. a Subject check, shows up after the
        next Script 00 scan) -- but must never regress from something
        real to something blank, cryptic, or a known failure placeholder,
        since that would confuse an admin into thinking something
        happened to the live rule when it didn't; the rule is fine, the
        CSV's display of it just went stale/wrong.

        A value is considered NOT meaningful (untrustworthy to overwrite
        existing real content with) if ANY of:
          - blank / whitespace-only
          - exactly matches a known read-failure placeholder string
            ('[No actions]', '[Could not read actions]', '[No
            conditions]', '[Could not read conditions]')
          - shorter than $MinimumLength characters (default 10) -- even
            the shortest legitimate real summaries ("Delete", "Mark as
            Read") clear this easily; a value this short is far more
            likely a truncated/corrupted fragment than genuine content.

    .PARAMETER Value
        The Conditions or Actions string to check.

    .PARAMETER MinimumLength
        Minimum character length to be considered meaningful. Default 10.

    .OUTPUTS
        [bool] $true if the value is meaningful/trustworthy, $false if it
        is blank, a known placeholder, or too short to be meaningful.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Value,

        [int]$MinimumLength = 10
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }

    $knownPlaceholders = @(
        '[No actions]',
        '[Could not read actions]',
        '[No conditions]',
        '[Could not read conditions]'
    )
    if ($knownPlaceholders -contains $Value.Trim()) { return $false }

    if ($Value.Trim().Length -lt $MinimumLength) { return $false }

    return $true
}


function Connect-OutlookCOM {
    <#
    .SYNOPSIS
        Creates and returns a controlled Outlook COM application session.

    .DESCRIPTION
        Launches Outlook via COM automation and returns the Application
        object. Outlook is started minimized and without any splash screen
        to minimize operator distraction during script execution.

        The returned Application object is stored at module scope so
        Release-OutlookCOM can clean up properly regardless of how
        the calling script exits.

        IMPORTANT: Always pair this call with Release-OutlookCOM in a
        finally block to ensure proper COM cleanup.

    .PARAMETER ProfileName
        The Outlook profile to use. If empty, uses the default profile.
        Outlook will prompt for profile selection if multiple profiles
        exist and no default is set.

    .OUTPUTS
        [System.__ComObject] -- The Outlook.Application COM object,
        or $null if Outlook could not be started.

    .EXAMPLE
        $outlook = Connect-OutlookCOM
        try {
            $namespace = $outlook.GetNamespace('MAPI')
            # ... work ...
        }
        finally {
            Release-OutlookCOM
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ProfileName = '',

        [Parameter(Mandatory = $false)]
        [bool]$AllowRunning = $false,
        # When $true, attaches to an already-running Outlook instance via
        # GetActiveObject instead of failing. Used by Phase C Send/Receive
        # restore when the operator left Outlook open after Phase B IMAP add.

        [Parameter(Mandatory = $false)]
        [bool]$VisibleLaunch = $false
        # When $true, launches Outlook as a visible foreground window via
        # Start-Process outlook.exe, waits for it to fully load, then attaches
        # to the running instance via GetActiveObject. This ensures the Outlook
        # window is visible on the taskbar without requiring the operator to
        # manually click the taskbar icon. Used by Phase A (initial COM open)
        # in Script 02. Phase B uses Start-Process directly and does not use
        # this path. AllowRunning is implied when VisibleLaunch is true.
    )

    Write-OMMigrateLog -Message 'Initializing Outlook COM session...' -Level INFO

    # Check if Outlook is already running
    $existing = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
    if ($existing) {
        if ($AllowRunning) {
            # Attach to the existing running instance via GetActiveObject.
            # This is safe for Phase C -- we just need to call Resume-OutlookSendReceive
            # on whatever Outlook session is already active.
            Write-OMMigrateLog -Message 'Outlook already running -- attaching to existing instance.' `
                               -Level INFO
            try {
                $Script:OutlookApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
                Register-COMObject -ComObject $Script:OutlookApp
                $Script:OutlookNamespace = $Script:OutlookApp.GetNamespace('MAPI')
                Register-COMObject -ComObject $Script:OutlookNamespace
                Write-OMMigrateLog -Message 'Attached to existing Outlook COM session.' -Level INFO
                Write-AuditEntry -Action 'OUTLOOK_COM_CONNECTED' -Detail 'Attached to existing running instance'
                return $Script:OutlookApp
            }
            catch {
                Write-OMMigrateLog -Message "Failed to attach to existing Outlook instance: $_" `
                                   -Level ERROR
                return $null
            }
        }
        else {
            Write-OMMigrateLog -Message (
                'Outlook is already running. OMMigrate requires exclusive COM access. ' +
                'Please close Outlook and retry.'
            ) -Level ERROR
            return $null
        }
    }

    # -- VisibleLaunch path ------------------------------------
    # Launch Outlook as a visible foreground window, wait for it
    # to fully load, then attach via GetActiveObject. This guarantees
    # the Outlook window is on the taskbar and interactive without
    # any need for ActiveExplorer().Activate() on a background instance.
    if ($VisibleLaunch) {
        Write-OMMigrateLog -Message 'VisibleLaunch: Starting Outlook as visible foreground window...' `
                           -Level INFO
        try {
            # Pass /profile to ensure Outlook opens the correct profile.
            # Without this, Outlook defaults to the last-used or default profile
            # which may not be the migration profile selected in Script 00.
            if ($ProfileName) {
                Start-Process 'outlook.exe' -ArgumentList "/profile `"$ProfileName`""
                Write-OMMigrateLog -Message "VisibleLaunch: Launching Outlook with profile '$ProfileName'." `
                                   -Level INFO
            }
            else {
                Start-Process 'outlook.exe'
                Write-OMMigrateLog -Message 'VisibleLaunch: Launching Outlook with default profile (no profile specified).' `
                                   -Level INFO
            }
        }
        catch {
            Write-OMMigrateLog -Message "VisibleLaunch: Could not launch outlook.exe: $_" -Level ERROR
            return $null
        }

        # Poll for the process and a loaded window. GetActiveObject will
        # fail until Outlook has registered its COM class, which happens
        # after the profile is loaded, not at process start. Ceiling is
        # read from settings (RulesEngine.OutlookLaunchTimeoutSeconds),
        # falling back to 30 if that setting is missing or unavailable --
        # slower machines or profiles with many PST/OST files may need more.
        $attached    = $false
        if ($Global:OMMigrate.Settings -and
            $Global:OMMigrate.Settings.PSObject.Properties['RulesEngine'] -and
            $Global:OMMigrate.Settings.RulesEngine.PSObject.Properties['OutlookLaunchTimeoutSeconds'] -and
            $Global:OMMigrate.Settings.RulesEngine.OutlookLaunchTimeoutSeconds -gt 0) {
            $maxWait = $Global:OMMigrate.Settings.RulesEngine.OutlookLaunchTimeoutSeconds
        }
        else {
            $maxWait = 30   # Fallback: settings unavailable or predates this option
        }
        $pollSeconds = 2
        $waited      = 0

        while ($waited -lt $maxWait) {
            Start-Sleep -Seconds $pollSeconds
            $waited += $pollSeconds
            $proc = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
            if (-not $proc) { continue }

            try {
                $Script:OutlookApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
                Register-COMObject -ComObject $Script:OutlookApp
                $Script:OutlookNamespace = $Script:OutlookApp.GetNamespace('MAPI')
                Register-COMObject -ComObject $Script:OutlookNamespace
                $attached = $true
                Write-OMMigrateLog -Message "VisibleLaunch: Attached to Outlook COM after ${waited}s." `
                                   -Level INFO
                Write-AuditEntry -Action 'OUTLOOK_COM_CONNECTED' `
                                 -Detail "VisibleLaunch=true | Profile='$ProfileName'"
                break
            }
            catch {
                # COM class not registered yet -- keep polling
                Write-OMMigrateLog -Message "VisibleLaunch: Outlook not ready yet (${waited}s) -- retrying..." `
                                   -Level DEBUG
            }
        }

        if (-not $attached) {
            Write-OMMigrateLog -Message "VisibleLaunch: Outlook did not become ready within ${maxWait}s." `
                               -Level ERROR
            return $null
        }

        return $Script:OutlookApp
    }

    try {
        # Create Outlook Application COM object
        $Script:OutlookApp = New-Object -ComObject 'Outlook.Application' -ErrorAction Stop
        Register-COMObject -ComObject $Script:OutlookApp

        # Get MAPI namespace
        $Script:OutlookNamespace = $Script:OutlookApp.GetNamespace('MAPI')
        Register-COMObject -ComObject $Script:OutlookNamespace

        # Log on to MAPI with specified profile (empty = default)
        # ShowDialog = false (don't show Choose Profile dialog)
        # NewSession = false (use existing profile data)
        if ($ProfileName) {
            $Script:OutlookNamespace.Logon($ProfileName, $null, $false, $false)
            Write-OMMigrateLog -Message "Outlook COM logged on to profile: '$ProfileName'" `
                               -Level INFO
        }
        else {
            $Script:OutlookNamespace.Logon('', $null, $false, $false)
            Write-OMMigrateLog -Message 'Outlook COM logged on to default profile.' `
                               -Level INFO
        }

        Write-AuditEntry -Action 'OUTLOOK_COM_CONNECTED' `
                         -Detail "Profile='$ProfileName'"

        # DISABLED 2026-07-01 (Administrator, confirmed via isolated Gemini-provided
        # minimal reproduction script, Test-MinimalHeadlessRuleDelete.ps1):
        # ActiveExplorer().Activate() initializes Outlook's Explorer/UI layer,
        # which caches its own independent snapshot of the rules table at that
        # moment. When the COM session later calls Quit(), that UI layer's
        # shutdown sequence flushes its own (now-stale) cached rules snapshot
        # back to PR_RW_RULES_STREAM, silently overwriting any rules changes
        # (specifically Remove()) made via the script's own COM-layer Rules
        # collection reference after the Explorer was activated -- even when
        # Save() was called correctly on that reference. Confirmed via direct
        # A/B test: the identical Remove()+InvokeMember-Save() sequence, run
        # standalone with no ActiveExplorer()/Activate() call anywhere, left
        # the deleted rule permanently gone; the same sequence inside the full
        # pipeline (which does call ActiveExplorer().Activate() here) did not.
        # Block left in place, not deleted, in case foreground-window behavior
        # needs to be restored for non-rules-related scripts in the future --
        # short-circuited via "-and $false" so it never executes.
        if ($true -and $false) {
            # Bring Outlook window to foreground so it appears on the taskbar
            # and the operator can interact with it without manually clicking
            # the icon. ActiveExplorer() returns the main Outlook window;
            # Activate() brings it to the foreground and gives it focus.
            # Non-fatal if this fails -- Outlook is still running correctly,
            # the operator can Alt+Tab to it if needed.
            try {
                $explorer = $Script:OutlookApp.ActiveExplorer()
                if ($explorer) {
                    $explorer.Activate()
                    Write-OMMigrateLog -Message 'Outlook window brought to foreground.' -Level DEBUG
                }
            }
            catch {
                Write-OMMigrateLog -Message "Could not bring Outlook window to foreground: $_" `
                                   -Level DEBUG
            }
        }

        return $Script:OutlookApp
    }
    catch {
        Write-OMMigrateLog -Message "Failed to initialize Outlook COM session: $_" `
                           -Level ERROR
        Release-OutlookCOM
        return $null
    }
}


function Close-OutlookIfRunning {
    <#
    .SYNOPSIS
        Closes Outlook if it is currently running, attempting a graceful
        shutdown first and only force-killing the process as a last resort.

    .DESCRIPTION
        Every OMMigrate script needs the same underlying guarantee before
        it can safely open its own Outlook COM session: Outlook must not
        already be running. This function is the single shared routine for
        satisfying that guarantee, regardless of which script or which
        moment in a script's execution is calling it (pre-flight before
        launch, mid-script before a new phase's COM session, or
        post-completion cleanup on the way out).

        Mechanically this always does the same thing, because the
        underlying COM/process reality is always the same thing:
          1. Check whether OUTLOOK.EXE is running.
          2. If so, attach to it via Connect-OutlookCOM -AllowRunning $true
             (GetActiveObject) so Release-OutlookCOM can drive a proper
             Logoff() + Quit() against the real session.
          3. Release-OutlookCOM then waits (default up to 20s) for Outlook
             to exit naturally so pending writes (e.g. PR_RW_RULES_STREAM
             after a rules import/Apply, or PST buffers) are flushed to
             disk before the process exits.
          4. Only if Outlook still has not exited after that wait does
             Release-OutlookCOM fall back to Stop-Process -Force. A
             premature force-kill is what risks silently discarding
             unflushed writes -- this is why graceful-first matters here.

        What legitimately differs per call site is context, not mechanics:
        what to tell the operator, and what to log. That context is
        supplied via -Reason. Some call sites may also want a different
        tolerance for how long to wait before giving up on a graceful
        exit -- that is supplied via -TimeoutSeconds, which is passed
        through to the underlying wait.

    .PARAMETER Reason
        Short human-readable phrase describing why Outlook is being closed
        at this point in the script, used only for operator-facing
        Write-Host messages and the log entry (e.g. 'before pre-flight',
        'after Phase B IMAP add', 'ready for Script 01'). Optional --
        if omitted, a generic message is used.

    .PARAMETER TimeoutSeconds
        Maximum seconds to allow for a graceful exit before falling back
        to a forced process kill. Passed through to the wait performed
        inside Release-OutlookCOM. If omitted, Release-OutlookCOM's own
        default (20s) is used.

    .OUTPUTS
        [bool] $true if Outlook was not running, or was running and is now
        confirmed closed. $false if Outlook was running and could not be
        confirmed closed (caller should warn the operator and/or abort).

    .EXAMPLE
        if (-not (Close-OutlookIfRunning -Reason 'before pre-flight')) {
            Write-Host '  Could not close Outlook automatically -- please close it manually.' -ForegroundColor Yellow
        }

    .EXAMPLE
        Close-OutlookIfRunning -Reason 'after Phase B IMAP add' -TimeoutSeconds 30
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Reason = '',

        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 0
        # 0 means "not specified" -- Release-OutlookCOM's own default is used.
    )

    $reasonSuffix = if ($Reason) { " $Reason" } else { '' }

    $outlookRunning = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
    if (-not $outlookRunning) {
        Write-OMMigrateLog -Message "Outlook is not running$reasonSuffix -- nothing to close." -Level DEBUG
        return $true
    }

    Write-OMMigrateLog -Message "Outlook is running -- closing automatically$reasonSuffix." -Level INFO
    Write-Host ''
    Write-Host "  Outlook is open -- closing automatically$reasonSuffix..." -ForegroundColor Cyan

    try {
        # Attach to the already-running instance so Release-OutlookCOM has a
        # real session to Logoff()/Quit() against, rather than going straight
        # to a force-kill with no chance for Outlook to flush pending writes.
        $attached = Connect-OutlookCOM -AllowRunning $true
        if (-not $attached) {
            Write-OMMigrateLog -Message "Could not attach to running Outlook instance$reasonSuffix -- falling back to direct force-close." `
                               -Level WARN
            $outlookRunning | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 3
        }
        else {
            if ($TimeoutSeconds -gt 0) {
                Release-OutlookCOM -TimeoutSeconds $TimeoutSeconds
            }
            else {
                Release-OutlookCOM
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Error while closing Outlook automatically${reasonSuffix}: $_" -Level WARN
    }

    if (Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue) {
        Write-OMMigrateLog -Message "Outlook still running after close attempt$reasonSuffix." -Level WARN
        Write-Host '  Could not close Outlook automatically -- please close it manually.' `
                   -ForegroundColor Yellow
        return $false
    }

    Write-Host '  Outlook closed.' -ForegroundColor Green
    Write-OMMigrateLog -Message "Outlook closed automatically$reasonSuffix." -Level INFO
    return $true
}


function Get-OutlookNamespace {
    <#
    .SYNOPSIS
        Returns the current MAPI namespace object from the active
        COM session.

    .DESCRIPTION
        Returns the cached namespace object created by Connect-OutlookCOM.
        Validates that a COM session is active before returning.

    .OUTPUTS
        [System.__ComObject] -- The MAPI namespace object, or $null.
    #>
    if (-not $Script:OutlookNamespace) {
        Write-OMMigrateLog -Message 'No active Outlook COM session. Call Connect-OutlookCOM first.' `
                           -Level ERROR
        return $null
    }
    return $Script:OutlookNamespace
}


function Register-COMObject {
    <#
    .SYNOPSIS
        Registers a COM object for deferred release by Release-OutlookCOM.

    .DESCRIPTION
        All COM objects created during a session should be registered
        here so Release-OutlookCOM can systematically release them.
        This prevents COM memory leaks and ensures Outlook terminates
        cleanly after the script finishes.

        Internal function -- not exported.

    .PARAMETER ComObject
        The COM object to register.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$ComObject
    )

    if ($ComObject -and $ComObject -is [System.__ComObject]) {
        [void]$Script:COMObjectsToRelease.Add($ComObject)
    }
}


function Release-OutlookCOM {
    <#
    .SYNOPSIS
        Releases all COM objects and terminates the Outlook session.

    .DESCRIPTION
        Releases all registered COM objects in reverse order (most
        recently created first), then quits the Outlook Application.

        Must be called in a finally block in every script that uses
        Connect-OutlookCOM to ensure proper cleanup regardless of
        whether the script succeeded or failed.

        Safe to call even if Connect-OutlookCOM was never successfully
        called -- handles null references gracefully.

    .PARAMETER TimeoutSeconds
        Maximum seconds to wait for Outlook to exit naturally after Quit()
        before falling back to a forced process kill. Default: read from
        OMMigrate_Settings.json (RulesEngine.OutlookQuitTimeoutSeconds),
        falling back to 20 if that setting is missing (e.g. an older
        settings file predating this option) or unavailable. Passing this
        parameter explicitly overrides the settings value for this one
        call only -- most callers should leave it unset and control the
        timeout via settings instead, so the value can be tuned per
        environment (larger rule sets need more time to serialize on
        Quit) without editing code.

    .EXAMPLE
        $outlook = Connect-OutlookCOM
        try {
            # ... script work ...
        }
        finally {
            Release-OutlookCOM
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 0   # 0 = "not explicitly specified" -- resolved below from settings
    )

    # Resolve the actual timeout: explicit parameter wins if the caller
    # passed one; otherwise read from settings; otherwise fall back to the
    # long-standing default of 20. This mirrors the same "0 means use the
    # settings/default value" convention already used by Connect-OutlookCOM's
    # own TimeoutSeconds parameter elsewhere in this module, for consistency.
    if ($TimeoutSeconds -gt 0) {
        $resolvedTimeoutSeconds = $TimeoutSeconds
    }
    elseif ($Global:OMMigrate.Settings -and
            $Global:OMMigrate.Settings.PSObject.Properties['RulesEngine'] -and
            $Global:OMMigrate.Settings.RulesEngine.PSObject.Properties['OutlookQuitTimeoutSeconds'] -and
            $Global:OMMigrate.Settings.RulesEngine.OutlookQuitTimeoutSeconds -gt 0) {
        $resolvedTimeoutSeconds = $Global:OMMigrate.Settings.RulesEngine.OutlookQuitTimeoutSeconds
    }
    else {
        $resolvedTimeoutSeconds = 20   # Fallback: settings unavailable or predates this option
    }

    Write-OMMigrateLog -Message 'Releasing Outlook COM session...' -Level INFO

    # Release registered COM objects in reverse order
    for ($i = $Script:COMObjectsToRelease.Count - 1; $i -ge 0; $i--) {
        try {
            $obj = $Script:COMObjectsToRelease[$i]
            if ($obj) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
            }
        }
        catch { }
    }
    $Script:COMObjectsToRelease.Clear()

    # Quit Outlook application
    # Logoff() before Quit() is required to prevent secondary IMAP store
    # rules from bleeding into the default store's Rules & Alerts UI on restart.
    # Proved June 2026: Restore-GmailRules.ps1 calls Logoff() and never
    # bleeds; scripts that skip Logoff() cause the bleed.
    try {
        if ($Script:OutlookNamespace) {
            try { $Script:OutlookNamespace.Logoff() } catch { }
        }
    } catch { }

    try {
        if ($Script:OutlookApp) {
            $Script:OutlookApp.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Script:OutlookApp
            )
            $Script:OutlookApp      = $null
            $Script:OutlookNamespace = $null
            $Script:MountedStoreNames = $null
        }
    }
    catch { }

    # Force garbage collection to release COM references
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()

    # Wait for Outlook to exit naturally after Quit().
    # Outlook must flush all pending PST write buffers before the process exits.
    # Stop-Process -Force must NOT be used until Outlook has fully written to disk --
    # a premature kill will silently discard unflushed PST changes (folder creates,
    # item copies) leaving the PST file in a stale state with no error reported.
    # Wait up to $resolvedTimeoutSeconds (settings-driven, default 20) for a
    # clean exit before resorting to force kill. The 3s RemoveStore settle
    # pause in Remove-POP3Account now handles the main MAPI flush delay --
    # the full timeout here exists mainly to let large rule sets fully
    # serialize on larger environments; smaller ones can lower it in settings.
    $waitSeconds  = $resolvedTimeoutSeconds
    $pollInterval = 500   # ms
    $waited       = 0
    $exited       = $false

    while ($waited -lt ($waitSeconds * 1000)) {
        Start-Sleep -Milliseconds $pollInterval
        $waited += $pollInterval
        $stillRunning = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
        if (-not $stillRunning) {
            $exited = $true
            Write-OMMigrateLog -Message "Outlook exited cleanly after $($waited)ms." -Level DEBUG
            break
        }
    }

    # Only force-kill if Outlook refused to exit within the wait window.
    # At this point all COM references are released so any remaining process
    # is a true zombie -- safe to terminate.
    if (-not $exited) {
        try {
            Stop-Process -Name 'OUTLOOK' -Force -ErrorAction Stop
            Start-Sleep -Milliseconds 1000
            Write-OMMigrateLog -Message "Outlook force-terminated after $($waitSeconds)s wait." -Level INFO
        }
        catch {
            Write-OMMigrateLog -Message "Could not terminate Outlook process: $_" -Level WARN
        }
    }

    Write-OMMigrateLog -Message 'Outlook COM session released.' -Level INFO
    Write-AuditEntry  -Action 'OUTLOOK_COM_RELEASED' -Detail 'COM session closed cleanly'
}


# ============================================================
#  REGION: SEND/RECEIVE MANAGEMENT
# ============================================================

function Suspend-OutlookSendReceive {
    <#
    .SYNOPSIS
        Suspends all Outlook Send/Receive groups and stops any
        active synchronization before migration begins.

    .DESCRIPTION
        Iterates all Send/Receive groups (SyncObjects) in the active
        Outlook COM session and stops each one. Also sets the
        ScheduledSendReceive interval to 0 to prevent automatic
        Send/Receive from firing during the migration window.

        The original schedule interval is saved to the global session
        context so Resume-OutlookSendReceive can restore it exactly.

        Called automatically by Scripts 01, 02, and 03 immediately
        after Connect-OutlookCOM succeeds.

        Why this matters:
            Active Send/Receive during migration can cause:
            - PST export conflicts (Script 01)
            - Timing issues when removing/adding accounts (Script 02)
            - Folder item duplication during content copy (Script 03)

        Safe for all configurations including:
            - Standard automatic Send/Receive schedules
            - Custom Send/Receive groups (e.g. split groups for
              installations with more than 20 accounts)
            - Already-disabled Send/Receive (no-op -- restores same state)

    .OUTPUTS
        [bool] -- $true if suspension succeeded, $false if failed.
                 Failure is non-fatal -- migration continues with a warning.

    .EXAMPLE
        $outlook = Connect-OutlookCOM
        Suspend-OutlookSendReceive
        try {
            # ... migration work ...
        }
        finally {
            Resume-OutlookSendReceive
            Release-OutlookCOM
        }
    #>
    [CmdletBinding()]
    param()

    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message 'WhatIf: Would suspend Outlook Send/Receive groups.' `
                           -Level INFO -WhatIfPrefix
        return $true
    }

    $namespace = Get-OutlookNamespace
    if (-not $namespace) { return $false }

    Write-OMMigrateLog -Message 'Suspending Outlook Send/Receive groups...' -Level INFO

    # Store original state for restore by Resume-OutlookSendReceive.
    # State is saved to BOTH the global session context (for use in the same
    # PowerShell session) AND a JSON file in Manifests (for use after Outlook
    # is closed and reopened, or if the script is interrupted and restarted).
    # Resume-OutlookSendReceive reads the JSON file if the global state is gone.
    # This is critical for Script 02 where Outlook is closed between suspend and
    # resume -- the COM session state is lost when Outlook exits.
    if (-not $Global:OMMigrate.Contains('SendReceiveState')) {
        $Global:OMMigrate['SendReceiveState'] = [ordered]@{
            GroupStates = [System.Collections.Generic.List[PSCustomObject]]::new()
            Suspended   = $false
        }
    }

    $succeeded   = $true
    $groupsStopped = 0

    try {
        $syncObjects = $namespace.SyncObjects
        Register-COMObject -ComObject $syncObjects

        $groupCount = $syncObjects.Count
        Write-OMMigrateLog -Message "Send/Receive groups found: $groupCount" -Level DEBUG

        for ($i = 1; $i -le $groupCount; $i++) {
            try {
                $group = $syncObjects.Item($i)
                Register-COMObject -ComObject $group

                # Save group name and current OnDemandOnly state
                # OnDemandOnly = $true means automatic sync is already off
                $groupState = [PSCustomObject]@{
                    Name        = $group.Name
                    OnDemandOnly = $true   # Default assumption -- safe fallback
                }

                try {
                    $groupState.OnDemandOnly = $group.ScheduledSendReceive
                }
                catch {
                    # ScheduledSendReceive is not available on all Send/Receive group types
                    # in Outlook 2021 and later. This is expected -- the safe fallback
                    # (OnDemandOnly = $true) is already set above so sync is correctly
                    # treated as already disabled. No action needed.
                    Write-OMMigrateLog -Message (
                        "ScheduledSendReceive not available for group '$($group.Name)' " +
                        "(expected on some group types) -- using safe fallback."
                    ) -Level DEBUG
                }

                [void]$Global:OMMigrate['SendReceiveState'].GroupStates.Add($groupState)

                # Stop any active sync on this group
                try {
                    $group.Stop()
                    $groupsStopped++
                    Write-OMMigrateLog -Message "Send/Receive stopped: '$($group.Name)'" `
                                       -Level DEBUG
                }
                catch {
                    Write-OMMigrateLog -Message "Could not stop group '$($group.Name)': $_" `
                                       -Level DEBUG
                }
            }
            catch {
                Write-OMMigrateLog -Message "Error processing Send/Receive group [$i]: $_" `
                                   -Level WARN
                $succeeded = $false
            }
        }

        $Global:OMMigrate['SendReceiveState'].Suspended = $true

        # Save state to persistent JSON file in Manifests folder.
        # This survives Outlook closing and reopening between suspend and resume
        # (Script 02 closes Outlook before the manual IMAP add step).
        # Resume-OutlookSendReceive reads this file if global state is gone.
        try {
            $stateFile = Join-Path $Global:OMMigrate.ManifestPath 'Step02_SendReceiveState.json'
            $stateData = [ordered]@{
                CapturedAt = (Get-Date -Format 'o')
                Suspended  = $true
                Groups     = @(
                    $Global:OMMigrate['SendReceiveState'].GroupStates | ForEach-Object {
                        [ordered]@{
                            Name         = $_.Name
                            WasEnabled   = -not $_.OnDemandOnly
                        }
                    }
                )
            }
            $stateData | ConvertTo-Json -Depth 3 | Set-Content -Path $stateFile -Encoding UTF8
            Write-OMMigrateLog -Message "Send/Receive state saved to: $stateFile" -Level DEBUG
        }
        catch {
            # Non-fatal -- global state is still available for same-session resume
            Write-OMMigrateLog -Message "Could not save Send/Receive state file: $_" -Level WARN
        }

        Write-OMMigrateLog -Message (
            "Send/Receive suspended. Groups stopped: $groupsStopped of $groupCount"
        ) -Level INFO

        Write-Host "  Send/Receive suspended ($groupsStopped group(s) stopped)." `
                   -ForegroundColor Green

        Write-AuditEntry -Action 'SENDRECEIVE_SUSPENDED' `
                         -Detail "Groups stopped: $groupsStopped of $groupCount"
    }
    catch {
        Write-OMMigrateLog -Message "Failed to suspend Send/Receive: $_" -Level WARN
        Write-Host '  WARNING: Could not suspend Send/Receive groups.' -ForegroundColor Yellow
        Write-Host '  Migration will continue -- monitor Outlook for sync activity.' `
                   -ForegroundColor Yellow
        $succeeded = $false
    }

    return $succeeded
}


function Resume-OutlookSendReceive {
    <#
    .SYNOPSIS
        Restores Outlook Send/Receive groups to their original state
        after migration is complete.

    .DESCRIPTION
        Reads the group state saved by Suspend-OutlookSendReceive and
        restores each group to its original configuration. If Send/Receive
        was already disabled before the migration, it is left disabled.
        If it was enabled on a schedule, the schedule is restored.

        Called automatically by Scripts 01, 02, and 03 in the finally
        block immediately before Release-OutlookCOM.

        Safe to call even if Suspend-OutlookSendReceive was never called
        or failed -- handles missing state gracefully.

    .OUTPUTS
        [bool] -- $true if resume succeeded or nothing to restore.

    .EXAMPLE
        finally {
            Resume-OutlookSendReceive
            Release-OutlookCOM
        }
    #>
    [CmdletBinding()]
    param()

    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message 'WhatIf: Would resume Outlook Send/Receive groups.' `
                           -Level INFO -WhatIfPrefix
        return $true
    }

    # -- Load saved state ------------------------------------------
    # Try global session context first (same PowerShell session).
    # Fall back to the persistent JSON file written by Suspend-OutlookSendReceive
    # if global state is gone -- this happens when Outlook was closed and reopened
    # between suspend and resume (the normal Script 02 flow).
    $savedGroupStates = $null
    $stateFile = Join-Path $Global:OMMigrate.ManifestPath 'Step02_SendReceiveState.json'

    if ($Global:OMMigrate.Contains('SendReceiveState') -and
        $Global:OMMigrate['SendReceiveState'].Suspended) {
        # Use global state -- same PowerShell session
        $savedGroupStates = $Global:OMMigrate['SendReceiveState'].GroupStates
        Write-OMMigrateLog -Message 'Restoring Send/Receive from global session state.' -Level DEBUG
    }
    elseif (Test-Path $stateFile) {
        # Use persistent JSON file -- Outlook was closed/reopened between sessions
        try {
            $stateData = Get-Content -Path $stateFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($stateData.Suspended) {
                $savedGroupStates = $stateData.Groups
                Write-OMMigrateLog -Message "Restoring Send/Receive from state file: $stateFile" `
                                   -Level DEBUG
            }
        }
        catch {
            Write-OMMigrateLog -Message "Could not read Send/Receive state file: $_" -Level WARN
        }
    }

    if (-not $savedGroupStates) {
        Write-OMMigrateLog -Message 'No Send/Receive state to restore.' -Level DEBUG
        return $true
    }

    $namespace = Get-OutlookNamespace
    if (-not $namespace) { return $false }

    Write-OMMigrateLog -Message 'Restoring Outlook Send/Receive groups...' -Level INFO

    $restored  = 0
    $leftAlone = 0

    try {
        $syncObjects = $namespace.SyncObjects
        Register-COMObject -ComObject $syncObjects

        for ($i = 1; $i -le $syncObjects.Count; $i++) {
            try {
                $group = $syncObjects.Item($i)
                Register-COMObject -ComObject $group

                # Find the saved state for this group by name
                $savedState = $savedGroupStates | Where-Object {
                    $_.Name -eq $group.Name
                } | Select-Object -First 1

                if ($savedState) {
                    # WasEnabled = $true means the group was active before we suspended it
                    # WasEnabled = $false means it was already disabled -- leave it that way
                    # This correctly handles Administrator's setup where both groups are already off
                    $wasEnabled = if ($savedState.PSObject.Properties['WasEnabled']) {
                        [bool]$savedState.WasEnabled
                    } else {
                        # Legacy global state uses OnDemandOnly (inverted) -- convert
                        -not [bool]$savedState.OnDemandOnly
                    }

                    if ($wasEnabled) {
                        # Group was active before -- restore it
                        try {
                            $group.ScheduledSendReceive = $true
                            $restored++
                            Write-OMMigrateLog -Message "Send/Receive restored (was enabled): '$($group.Name)'" `
                                               -Level DEBUG
                        }
                        catch {
                            Write-OMMigrateLog -Message "Could not restore group '$($group.Name)': $_" `
                                               -Level DEBUG
                        }
                    }
                    else {
                        # Group was already disabled before we suspended it -- leave it off
                        $leftAlone++
                        Write-OMMigrateLog -Message "Send/Receive left disabled (was already off): '$($group.Name)'" `
                                           -Level DEBUG
                    }
                }
            }
            catch {
                Write-OMMigrateLog -Message "Error restoring Send/Receive group [$i]: $_" `
                                   -Level WARN
            }
        }

        Write-OMMigrateLog -Message (
            "Send/Receive restore complete. Restored: $restored | Left disabled: $leftAlone"
        ) -Level INFO
        Write-Host "  Send/Receive groups restored ($restored enabled, $leftAlone left disabled as before)." `
                   -ForegroundColor Green
        Write-AuditEntry -Action 'SENDRECEIVE_RESUMED' `
                         -Detail "Restored: $restored | Left disabled (already off): $leftAlone"
    }
    catch {
        Write-OMMigrateLog -Message "Failed to restore Send/Receive: $_" -Level WARN
        Write-Host '  WARNING: Could not restore Send/Receive groups.' -ForegroundColor Yellow
        Write-Host '  Verify your Send/Receive settings in Outlook after migration.' `
                   -ForegroundColor Yellow
    }

    # Clean up -- remove both global state and the JSON file
    if ($Global:OMMigrate.Contains('SendReceiveState')) {
        $Global:OMMigrate.Remove('SendReceiveState')
    }
    try {
        if (Test-Path $stateFile) {
            Remove-Item -Path $stateFile -Force -ErrorAction SilentlyContinue
            Write-OMMigrateLog -Message "Send/Receive state file cleaned up." -Level DEBUG
        }
    }
    catch { }

    return $true
}


# ============================================================
#  REGION: ACCOUNT ENUMERATION VIA COM
# ============================================================

function Get-OutlookAccountsViaCOM {
    <#
    .SYNOPSIS
        Returns all accounts from the active Outlook COM session.

    .DESCRIPTION
        Uses the Outlook.Application.Session.Accounts collection to
        enumerate all configured email accounts. This is the authoritative
        source for account type information -- more reliable than registry
        parsing alone for edge cases.

        Results are used by Script 00 to corroborate and supplement
        the registry-based discovery, and by Scripts 02/03 for
        account-level operations.

        Requires an active COM session from Connect-OutlookCOM.

    .OUTPUTS
        [PSCustomObject[]] -- Array of COM account info objects:
            AccountName       [string]  Display name
            EmailAddress      [string]  Email address
            AccountType       [string]  POP3 | IMAP | Exchange | HTTP | Other
            AccountTypeInt    [int]     Raw OlAccountType value
            DeliveryStore     [string]  Store name for this account
            UserName          [string]  Login username

    .EXAMPLE
        $outlook  = Connect-OutlookCOM
        $accounts = Get-OutlookAccountsViaCOM
    #>
    [CmdletBinding()]
    param()

    $namespace = Get-OutlookNamespace
    if (-not $namespace) { return @() }

    Write-OMMigrateLog -Message 'Enumerating accounts via Outlook COM...' -Level INFO

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        $accounts = $namespace.Accounts
        Register-COMObject -ComObject $accounts

        $count = $accounts.Count
        Write-OMMigrateLog -Message "Outlook reports $count account(s) in this profile." `
                           -Level INFO

        for ($i = 1; $i -le $count; $i++) {
            $account = $null
            try {
                $account = $accounts.Item($i)
                Register-COMObject -ComObject $account

                # Map OlAccountType integer to string
                $typeInt    = $account.AccountType
                $typeString = switch ($typeInt) {
                    0 { 'Exchange' }
                    1 { 'IMAP'     }
                    2 { 'POP3'     }
                    3 { 'HTTP'     }
                    5 { 'Other'    }
                    default { "Unknown($typeInt)" }
                }

                # Get delivery store name safely
                $storeName = ''
                try {
                    $store     = $account.DeliveryStore
                    $storeName = $store.DisplayName
                    Register-COMObject -ComObject $store
                }
                catch { }

                # Get server settings from COM Account object
                $incomingServer = ''
                $outgoingServer = ''
                $incomingPort   = 0
                $outgoingPort   = 0
                $incomingSSL    = $true
                $outgoingSSL    = $true
                try {
                    # These properties exist on POP3/IMAP accounts
                    $incomingServer = $account.POPServerName
                    if (-not $incomingServer) { $incomingServer = $account.IMAPServerName }
                }
                catch { }
                try { $outgoingServer = $account.SMTPServerName }   catch { }
                try { $incomingPort   = $account.POPServerPort  }   catch { }
                try { if (-not $incomingPort) { $incomingPort = $account.IMAPServerPort } } catch { }
                try { $outgoingPort   = $account.SMTPServerPort }   catch { }
                try { $incomingSSL    = $account.POPServerRequiresSSL  } catch { }
                try { if (-not $incomingSSL) { $incomingSSL = $account.IMAPServerRequiresSSL } } catch { }
                try { $outgoingSSL    = $account.SMTPServerRequiresSSL } catch { }

                # Get PST/OST file path from delivery store
                $filePath = ''
                try {
                    $deliveryStore = $account.DeliveryStore
                    if ($deliveryStore) {
                        $filePath  = $deliveryStore.FilePath
                        $storeName = $deliveryStore.DisplayName
                        Register-COMObject -ComObject $deliveryStore
                    }
                }
                catch { }

                # Validate the delivery store path belongs to this account.
                # DeliveryStore.FilePath can return a different account's OST
                # when two accounts share a delivery store or Outlook internally
                # assigns the wrong store. Check that the filename contains the
                # account's email address. If not, scan mounted stores then the
                # Outlook data folder on disk to find the correct OST file.
                $emailForMatch = $account.SmtpAddress
                if ($filePath -and $emailForMatch -and
                    -not ([System.IO.Path]::GetFileName($filePath) -like "*$emailForMatch*")) {

                    $correctedPath = ''

                    # First try mounted stores
                    try {
                        $allStores = $Script:OutlookNamespace.Stores
                        for ($si = 1; $si -le $allStores.Count; $si++) {
                            $st = $allStores.Item($si)
                            $stFp = ''
                            try { $stFp = $st.FilePath } catch { }
                            if ($stFp -like '*.ost' -and
                                [System.IO.Path]::GetFileName($stFp) -like "*$emailForMatch*") {
                                $correctedPath = $stFp
                                break
                            }
                        }
                    }
                    catch { }

                    # Fall back to scanning Outlook data folder on disk
                    if (-not $correctedPath) {
                        try {
                            $ostFolder = [System.IO.Path]::Combine(
                                $env:LOCALAPPDATA, 'Microsoft', 'Outlook'
                            )
                            $ostFiles = Get-ChildItem -Path $ostFolder -Filter '*.ost' `
                                                      -ErrorAction SilentlyContinue
                            foreach ($ostFile in $ostFiles) {
                                if ($ostFile.Name -like "*$emailForMatch*") {
                                    $correctedPath = $ostFile.FullName
                                    break
                                }
                            }
                        }
                        catch { }
                    }

                    if ($correctedPath) {
                        Write-OMMigrateLog -Message (
                            "Get-OutlookAccountsViaCOM: Corrected OST for '$emailForMatch' " +
                            "from '$([System.IO.Path]::GetFileName($filePath))' " +
                            "to '$([System.IO.Path]::GetFileName($correctedPath))'"
                        ) -Level INFO
                        $filePath = $correctedPath
                    }
                }

                $result = [PSCustomObject]@{
                    AccountName     = $account.DisplayName
                    EmailAddress    = $account.SmtpAddress
                    AccountType     = $typeString
                    AccountTypeInt  = $typeInt
                    DeliveryStore   = $storeName
                    UserName        = $account.UserName
                    IncomingServer  = $incomingServer
                    OutgoingServer  = $outgoingServer
                    IncomingPort    = $incomingPort
                    OutgoingPort    = $outgoingPort
                    IncomingSSL     = $incomingSSL
                    OutgoingSSL     = $outgoingSSL
                    FilePath        = $filePath
                }

                $results.Add($result)

                # Suppress COM account detail when Sanitize is active --
                # email addresses are not yet in the sanitize map at this point.
                if (-not $Global:OMMigrate.Sanitize) {
                    Write-OMMigrateLog -Message (
                        "COM Account [$i]: $($result.EmailAddress) | " +
                        "Type=$($result.AccountType) | Store=$($result.DeliveryStore)"
                    ) -Level INFO
                }
            }
            catch {
                Write-OMMigrateLog -Message "Error reading COM account [$i]: $_" `
                                   -Level WARN
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Failed to enumerate accounts via COM: $_" -Level ERROR
    }

    return $results
}


# ============================================================
#  REGION: FOLDER TREE ENUMERATION
# ============================================================

function Get-FolderTree {
    <#
    .SYNOPSIS
        Recursively enumerates the complete folder tree for all
        mail stores in the active Outlook COM session.

    .DESCRIPTION
        Walks the entire Outlook folder hierarchy -- all stores,
        all folders, all subfolders -- and returns a flat list of
        folder objects. Each object contains the folder's full path,
        item count, and the store it belongs to.

        Used by Script 00 to build folder_map.csv -- the operator's
        control file for deciding which folders go to the IMAP server
        vs stay in local PST storage.

        Requires an active COM session from Connect-OutlookCOM.

        Performance note: For large mailboxes (hundreds of folders,
        many thousands of items) this enumeration can take 30-60
        seconds. Progress is logged per store.

    .PARAMETER StoreDisplayName
        Optional. Limit enumeration to a specific store by display name.
        If omitted, all stores are enumerated.

    .PARAMETER ExcludeSystemFolders
        When $true, excludes Outlook system folders (Calendar, Contacts,
        Tasks, Notes, Journal) from the results -- returns mail folders only.
        Default: $true

    .PARAMETER ExcludeBackupPSTs
        When $true, skips folder enumeration entirely for any store whose
        DisplayName starts with 'Backup --' or 'ArchiveBuild --' (Administrator's
        manually attached backup/snapshot PSTs in the navigation pane). These
        are disposable safety copies, not migration targets, and walking
        their full folder trees on large PSTs (e.g. an ameritech backup with
        hundreds of folders) adds significant runtime with no migration
        value. Default: $true

    .OUTPUTS
        [PSCustomObject[]] -- Array of folder objects:
            FolderPath        [string]  Full path (e.g. 'Inbox\Vendors\Microsoft')
            FolderName        [string]  Folder name only
            StoreName         [string]  Parent store display name
            StoreType         [string]  'PST' | 'OST' | 'Exchange' | 'IMAP'
            ItemCount         [int]     Number of items in folder
            UnreadCount       [int]     Number of unread items
            HasSubfolders     [bool]    Whether folder has child folders
            IsSystemFolder    [bool]    Whether this is an Outlook system folder
            EntryID           [string]  Outlook Entry ID (for rule target matching)
            FolderDepth       [int]     Depth in hierarchy (0 = root store)

    .EXAMPLE
        $folders = Get-FolderTree
        $mailFolders = $folders | Where-Object { -not $_.IsSystemFolder }
        Write-Host "Total mail folders: $($mailFolders.Count)"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$StoreDisplayName = '',

        [Parameter(Mandatory = $false)]
        [bool]$ExcludeSystemFolders = $true,

        [Parameter(Mandatory = $false)]
        [bool]$ExcludeBackupPSTs = $true
    )

    $namespace = Get-OutlookNamespace
    if (-not $namespace) { return @() }

    Write-OMMigrateLog -Message 'Enumerating Outlook folder tree...' -Level INFO

    $allFolders = [System.Collections.Generic.List[PSCustomObject]]::new()

    # System folder names to identify and optionally exclude
    $systemFolderNames = @(
        'Calendar', 'Contacts', 'Tasks', 'Notes', 'Journal',
        'Suggested Contacts', 'Quick Step Settings', 'Conversation History',
        'Conversation Action Settings',
        'RSS Feeds', 'Sync Issues', 'Conflicts', 'Local Failures',
        'Server Failures', 'Search Folders', 'Recoverable Items',
        'Deletions', 'Purges', 'Versions'
    )

    try {
        $stores = $namespace.Stores
        Register-COMObject -ComObject $stores

        $storeCount = $stores.Count
        Write-OMMigrateLog -Message "Found $storeCount store(s) in Outlook profile." `
                           -Level INFO

        for ($s = 1; $s -le $storeCount; $s++) {
            $store = $null
            try {
                $store = $stores.Item($s)
                Register-COMObject -ComObject $store

                $storeName = $store.DisplayName

                # Skip manually attached backup/archive PSTs entirely --
                # these are disposable safety copies (e.g. 'Backup -- ameritech',
                # 'ArchiveBuild -- ameritech') Administrator attaches to the nav pane for
                # safekeeping, not migration targets. Walking their full folder
                # trees (large PSTs can have hundreds of folders) adds significant
                # runtime with zero migration value. Checked before Get-StoreType
                # and the recursive walk so the cost is avoided entirely, not just
                # the result discarded after the fact.
                if ($ExcludeBackupPSTs -and $storeName -and
                    ($storeName -like 'Backup --*' -or $storeName -like 'ArchiveBuild --*')) {
                    Write-OMMigrateLog -Message (
                        "Skipping folder enumeration for manually attached backup/archive PST: '$storeName'."
                    ) -Level INFO
                    continue
                }

                # Filter by store name if requested
                if ($StoreDisplayName -and
                    $storeName -ne $StoreDisplayName) {
                    continue
                }

                # Determine store type
                $storeType = Get-StoreType -Store $store

                # Register non-email store names in sanitize map before logging
                if ($Global:OMMigrate.Sanitize -and $storeName -and
                    $storeName -notmatch '@') {
                    Register-SanitizeTerms -Terms @($storeName) -Category 'Store'
                }
                Write-OMMigrateLog -Message "Enumerating store: '$storeName' (Type=$storeType)" `
                                   -Level INFO

                # Get root folder for this store
                $rootFolder = $null
                try {
                    $rootFolder = $store.GetRootFolder()
                    Register-COMObject -ComObject $rootFolder
                }
                catch {
                    Write-OMMigrateLog -Message "Cannot access root folder for store '$storeName': $_" `
                                       -Level WARN
                    continue
                }

                # Recursively enumerate all folders in this store
                $storeFolders = Get-FolderTreeRecursive `
                    -Folder      $rootFolder `
                    -StoreName   $storeName `
                    -StoreType   $storeType `
                    -ParentPath  '' `
                    -Depth       0 `
                    -SystemFolderNames $systemFolderNames `
                    -ExcludeSystem $ExcludeSystemFolders

                foreach ($f in $storeFolders) {
                    $allFolders.Add($f)
                }

                # For OST (IMAP) stores, guarantee all default mail folders are present.
                # COM .Folders enumeration can silently miss default folders like Inbox
                # and Outbox on IMAP accounts. Rather than relying on COM, generate rows
                # directly for any missing default folder using known names and defaults.
                if ($storeType -eq 'OST') {
                    $defaultFolderNames = @(
                        'Inbox', 'Outbox', 'Sent Items', 'Sent',
                        'Deleted Items', 'Drafts', 'Junk Email', 'Junk', 'Trash'
                    )
                    # Build a set of folder names already enumerated for this store
                    # Key includes StoreType so a PST Inbox does not suppress the OST Inbox
                    $existingNames = [System.Collections.Generic.HashSet[string]]::new(
                        [System.StringComparer]::OrdinalIgnoreCase
                    )
                    foreach ($f in $storeFolders) { [void]$existingNames.Add($f.FolderName) }

                    foreach ($dfName in $defaultFolderNames) {
                        if (-not $existingNames.Contains($dfName)) {
                            $dfObj = [PSCustomObject]@{
                                FolderPath      = $dfName
                                FolderName      = $dfName
                                StoreName       = $storeName
                                StoreType       = $storeType
                                ItemCount       = 0
                                UnreadCount     = 0
                                HasSubfolders   = $false
                                IsSystemFolder  = $false
                                EntryID         = ''
                                FolderDepth     = 0
                            }
                            $allFolders.Add($dfObj)
                            Write-OMMigrateLog -Message (
                                "OST default folder guarantee: Added missing '$dfName' for store '$storeName'."
                            ) -Level DEBUG
                        }
                    }
                }

                # Suppress store name when Sanitize is active -- map not built yet
                #
                # FIXED 2026-07-10, Administrator direction (live-tested on "TestProfile": a
                # brand-new, near-empty archive PST with only 1 folder --
                # Deleted Items -- triggered "The property 'Count' cannot be
                # found on this object", because Get-FolderTreeRecursive can
                # return a single bare PSCustomObject instead of an array when
                # exactly one folder is found, and PowerShell unwraps a
                # single-element pipeline result rather than keeping it as a
                # collection. $storeFolders.Count then fails since a bare
                # object has no .Count property. Folder DATA was never lost --
                # $allFolders.Add($f) above already added the folder correctly
                # via the foreach loop (which iterates a bare object fine) --
                # only this diagnostic count line broke. Force-cast to array
                # with @(...) so .Count always works regardless of how many
                # folders were found, 0/1/many, fixing the root cause instead
                # of just silencing the resulting WARN.
                if (-not $Global:OMMigrate.Sanitize) {
                    Write-OMMigrateLog -Message "Store '$storeName': $(@($storeFolders).Count) folders enumerated." `
                                       -Level INFO
                }
            }
            catch {
                Write-OMMigrateLog -Message "Error processing store [$s]: $_" -Level WARN
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Failed to enumerate folder tree: $_" -Level ERROR
    }

    Write-OMMigrateLog -Message "Total folders enumerated: $($allFolders.Count)" -Level INFO
    return $allFolders
}


function Get-FolderTreeRecursive {
    <#
    .SYNOPSIS
        Internal recursive helper for Get-FolderTree.

    .DESCRIPTION
        Walks a folder and all its subfolders recursively, building
        the flat folder list. Called by Get-FolderTree -- not intended
        for direct use.

        Error handling at each level ensures one bad folder does not
        stop enumeration of the rest of the tree.

    .PARAMETER Folder
        The current Outlook MAPIFolder COM object to process.

    .PARAMETER StoreName
        Display name of the store this folder belongs to.

    .PARAMETER StoreType
        Type string of the store (PST/OST/IMAP/Exchange).

    .PARAMETER ParentPath
        Path of the parent folder (used to build full path).

    .PARAMETER Depth
        Current recursion depth.

    .PARAMETER SystemFolderNames
        Array of known system folder names.

    .PARAMETER ExcludeSystem
        Whether to exclude system folders from results.

    .OUTPUTS
        [PSCustomObject[]] -- Folder objects for this folder and all descendants.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Folder,

        [Parameter(Mandatory = $true)]
        [string]$StoreName,

        [Parameter(Mandatory = $true)]
        [string]$StoreType,

        [Parameter(Mandatory = $false)]
        [string]$ParentPath = '',

        [Parameter(Mandatory = $false)]
        [int]$Depth = 0,

        [Parameter(Mandatory = $false)]
        [string[]]$SystemFolderNames = @(),

        [Parameter(Mandatory = $false)]
        [bool]$ExcludeSystem = $true
    )

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        $subfolders = $Folder.Folders
        Register-COMObject -ComObject $subfolders
        $subCount = $subfolders.Count

        for ($i = 1; $i -le $subCount; $i++) {
            $subfolder = $null
            try {
                $subfolder = $subfolders.Item($i)
                Register-COMObject -ComObject $subfolder

                $folderName = $subfolder.Name
                $isSystem   = $SystemFolderNames -contains $folderName

                # Build full path
                $fullPath = if ($ParentPath) {
                    "$ParentPath\$folderName"
                } else {
                    $folderName
                }

                # Skip system folders if requested -- applies at any depth,
                # not just root level. Hidden system folders like
                # 'Conversation Action Settings' can appear at any level.
                if ($ExcludeSystem -and $isSystem) {
                    continue
                }

                # Get item and unread counts safely
                $itemCount   = 0
                $unreadCount = 0
                try {
                    $itemCount   = $subfolder.Items.Count
                    $unreadCount = $subfolder.UnReadItemCount
                }
                catch { }

                # Get EntryID for rule target matching
                $entryID = ''
                try { $entryID = $subfolder.EntryID } catch { }

                $folderObj = [PSCustomObject]@{
                    FolderPath      = $fullPath
                    FolderName      = $folderName
                    StoreName       = $StoreName
                    StoreType       = $StoreType
                    ItemCount       = $itemCount
                    UnreadCount     = $unreadCount
                    HasSubfolders   = ($subfolder.Folders.Count -gt 0)
                    IsSystemFolder  = $isSystem
                    EntryID         = $entryID
                    FolderDepth     = $Depth
                }

                $results.Add($folderObj)

                # Recurse into subfolders
                if ($subfolder.Folders.Count -gt 0) {
                    $childFolders = Get-FolderTreeRecursive `
                        -Folder            $subfolder `
                        -StoreName         $StoreName `
                        -StoreType         $StoreType `
                        -ParentPath        $fullPath `
                        -Depth             ($Depth + 1) `
                        -SystemFolderNames $SystemFolderNames `
                        -ExcludeSystem     $ExcludeSystem

                    foreach ($child in $childFolders) {
                        $results.Add($child)
                    }
                }
            }
            catch {
                Write-OMMigrateLog -Message "Error reading subfolder [$i] in '$ParentPath': $_" `
                                   -Level DEBUG
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Error in folder recursion at '$ParentPath': $_" `
                           -Level WARN
    }

    return $results
}


function Get-StoreType {
    <#
    .SYNOPSIS
        Determines the storage type of an Outlook store object.

    .DESCRIPTION
        Uses the store's file path extension and exchange store type
        to classify it as PST, OST, IMAP, or Exchange.

    .PARAMETER Store
        An Outlook Store COM object.

    .OUTPUTS
        [string] -- 'PST' | 'OST' | 'IMAP' | 'Exchange' | 'Unknown'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Store
    )

    # Check ExchangeStoreType FIRST -- Exchange accounts cache locally to OST
    # files but must be classified as Exchange, not OST
    try {
        $storeType = $Store.ExchangeStoreType
        switch ($storeType) {
            0 { return 'Exchange'  }   # olExchangeMailbox (default/unset)
            1 { return 'Exchange'  }   # olExchangeMailbox
            2 { return 'Exchange'  }   # olExchangePublicFolder
            3 { }                      # IMAP -- fall through to file path check
            4 { return 'Exchange'  }   # olExchangePrivateStore
        }
    }
    catch { }

    # Check file path extension -- PST (POP3) or OST (IMAP)
    try {
        $filePath = $Store.FilePath
        if ($filePath -like '*.pst') { return 'PST' }
        if ($filePath -like '*.ost') { return 'OST' }
    }
    catch { }

    # ExchangeStoreType 3 with no file path = IMAP OST
    try {
        if ($Store.ExchangeStoreType -eq 3) { return 'OST' }
    }
    catch { }

    return 'Unknown'
}


function Export-FolderMapCSV {
    <#
    .SYNOPSIS
        Exports the folder tree to folder_map.csv for operator review.

    .DESCRIPTION
        Generates the folder mapping control file that the operator
        uses to decide which folders should be:

            Server   -- Created as IMAP server-side folders
                        (visible on all devices, synced to server)

            Local    -- Kept in local PST storage
                        (desktop only, never touches server)

            Skip     -- Do not migrate (empty folders, system artifacts)

        The CSV is pre-populated with intelligent defaults:
            - Inbox and standard mail folders -> Server (suggested)
            - All other folders -> Local (default)
            - Folders from POP3 accounts -> Local (default -- safest)
            NOTE: All folders default to Local. The operator changes
            any folder to Server if it should be visible on all devices.

        The operator changes the 'Destination' column values as needed
        before running Script 03.

        SECURITY: No email content is written to this file.
        Only folder names, paths, item counts, and migration decisions.

    .PARAMETER Folders
        Array of folder objects from Get-FolderTree.

    .PARAMETER OutputPath
        Full path for the output CSV file.
        Default: Config\folder_map.csv

    .OUTPUTS
        [string] -- Path to the written CSV file.

    .EXAMPLE
        $folders = Get-FolderTree
        $csvPath = Export-FolderMapCSV -Folders $folders
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[PSCustomObject]]$Folders,

        [Parameter(Mandatory = $false)]
        [string]$OutputPath = ''
    )

    if ([string]::IsNullOrEmpty($OutputPath)) {
        $OutputPath = Get-OMMigrateCsvPath -BaseName 'folder_map.csv'
    }

    Write-OMMigrateLog -Message "Generating folder map CSV: $OutputPath" -Level INFO

    # ADDED (fix, folder_map.csv StoreName staleness): build the same
    # account-email -> TargetStoreName lookup already used by Export-RulesToCSV
    # (RulesEngine.ArchiveStoreMappings from the Script 00 picker), so this
    # function can also reflect the operator's current archive-remap decision
    # in folder_map.csv's StoreName column, instead of always writing whatever
    # store a folder was physically found in by the live COM enumeration this
    # run. Read-only lookup here -- does not change Destination/Notes merge
    # behavior at all. Empty/no mappings configured leaves StoreName exactly
    # as before this fix (the live-enumeration value).
    $folderMapArchiveMappings = @()
    try {
        if ($Global:OMMigrate.Settings.PSObject.Properties['RulesEngine'] -and
            $Global:OMMigrate.Settings.RulesEngine.PSObject.Properties['ArchiveStoreMappings']) {
            $folderMapArchiveMappings = @($Global:OMMigrate.Settings.RulesEngine.ArchiveStoreMappings)
        }
    }
    catch { }

    $folderMapStoreToTarget = @{}
    foreach ($fmMapping in $folderMapArchiveMappings) {
        if (-not $fmMapping.PSObject.Properties['TargetStoreName'] -or
            [string]::IsNullOrWhiteSpace($fmMapping.TargetStoreName)) { continue }
        foreach ($fmMappedAccount in @($fmMapping.RuleStoreNames)) {
            if (-not [string]::IsNullOrWhiteSpace($fmMappedAccount)) {
                $folderMapStoreToTarget[$fmMappedAccount] = $fmMapping.TargetStoreName
            }
        }
    }

    # Standard mail folder names that are good IMAP server candidates.
    # These match what IMAP servers and Outlook create by default on new accounts.
    # Depth check (FolderDepth -le 1) ensures only top-level folders get Server --
    # subfolders of Inbox etc. stay Local unless the operator changes them.
    $serverSuggested = @(
        'Inbox', 'Sent Items', 'Sent Mail', 'Sent', 'Drafts', 'Draft',
        'Junk Email', 'Junk', 'Spam', 'Deleted Items', 'Trash', 'Deleted',
        'Archive', 'Outbox'
    )

    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()

    # ADDED (fix, folder_map.csv duplicate rows): tracks which effective
    # StoreType|FolderPath (PST) or StoreName|StoreType|FolderPath (non-PST)
    # combinations have already been added to $rows this run. Needed because
    # the StoreName-override fix above can cause two DIFFERENT live folders
    # (e.g. the same account's folder tree present in both its old and newly
    # remapped archive PST, both still attached in Outlook) to resolve to the
    # SAME effective StoreName -- without this check both would be added as
    # separate rows with identical StoreName+FolderPath, producing visible
    # duplicates in the CSV. First occurrence wins; later ones are skipped.
    $seenRowKeysThisRun = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($folder in $Folders) {

        # Determine suggested destination
        $suggested = 'Local'   # Default -- safest for POP3 migrations

        # Standard top-level mail folders suggest Server
        if ($serverSuggested -contains $folder.FolderName -and
            $folder.FolderDepth -le 1) {
            $suggested = 'Server'
        }

        # OST folders (IMAP-ALREADY, IMAP-CONVERTED, and COMPLETE accounts) all
        # default to Server -- they already live on the IMAP server and should stay
        # there. Account type eligibility is driven by the calling script (00, 03).
        # The force-Local post-merge correction handles exceptions (Sync Issues etc.)
        if ($folder.StoreType -eq 'OST') {
            $suggested = 'Server'
        }

        # ADDED (fix, folder_map.csv StoreName staleness): if this folder's
        # account (the first path segment of FolderPath, e.g.
        # 'user@example.com' from 'user@example.com\Inbox\AAA')
        # has a current ArchiveStoreMappings entry, and this is a PST-type folder
        # (an archive folder, not a live IMAP/OST folder), write the mapped
        # TargetStoreName here instead of the store this folder was physically
        # found in during this run's live enumeration. Non-PST (OST/Server)
        # folders are never touched -- archive mappings only apply to Local/
        # Archive-destined folders. No mapping found -- falls back to
        # $folder.StoreName exactly as before this fix.
        # FIXED (bug found live 2026-07-12, this same fix): the original
        # condition required FolderPath to contain a backslash before applying
        # the override -- this correctly handles subfolders like
        # 'user@example.com\Inbox\AAA', but silently skipped the
        # bare account-root row itself ('user@example.com', no
        # backslash, FolderDepth=0), which is a real row Get-FolderTreeRecursive
        # emits for every PST account folder. That row kept whatever store it
        # was physically enumerated from and then got lost on later merges
        # (confirmed live: row present under the old store, absent entirely
        # after a remap). Now checks FolderPath directly against the mapping
        # when there is no backslash (the whole path IS the account name), in
        # addition to the existing split-based check for subfolder paths.
        $effectiveFolderStoreName = $folder.StoreName
        if ($folder.StoreType -eq 'PST') {
            $fmFolderAccount = if ($folder.FolderPath -like '*\*') {
                ($folder.FolderPath -split '\\')[0]
            } else {
                $folder.FolderPath
            }
            if ($folderMapStoreToTarget.ContainsKey($fmFolderAccount)) {
                $effectiveFolderStoreName = $folderMapStoreToTarget[$fmFolderAccount]
            }
        }

        $row = [PSCustomObject]@{
            # -- Account grouping column first -- makes Excel review easy ----
            'StoreName'         = $effectiveFolderStoreName

            # -- Operator sets this -----------------------------------------
            'Destination'       = $suggested    # Server | Local
                                                # Operator reviews and confirms

            # -- Auto-populated ---------------------------------------------
            'FolderPath'        = $folder.FolderPath
            'FolderName'        = $folder.FolderName
            'StoreType'         = $folder.StoreType
            'ItemCount'         = $folder.ItemCount
            'UnreadCount'       = $folder.UnreadCount
            'HasSubfolders'     = $folder.HasSubfolders
            'FolderDepth'       = $folder.FolderDepth
            'IsSystemFolder'    = $folder.IsSystemFolder
            'EntryID'           = $folder.EntryID

            # -- Notes for operator -----------------------------------------
            'Notes'             = $(
                if ($suggested -eq 'Server') {
                    'Suggested: Create on IMAP server -- visible on all devices'
                }
                elseif ($folder.ItemCount -eq 0) {
                    'Empty now -- may receive mail via rules. Change to Server if needed on all devices.'
                }
                else {
                    'Suggested: Keep local -- change to Server if needed on all devices'
                }
            )
        }

        # ADDED (fix, folder_map.csv duplicate rows): skip this folder if its
        # effective StoreName+FolderPath key was already added this run --
        # see $seenRowKeysThisRun declaration above for why this can happen.
        $rowDedupKey = if ($row.StoreType -eq 'PST') {
            "$($row.StoreType)|$($row.FolderPath)"
        } else {
            "$($row.StoreName)|$($row.StoreType)|$($row.FolderPath)"
        }
        if ($seenRowKeysThisRun.Contains($rowDedupKey)) {
            continue
        }
        [void]$seenRowKeysThisRun.Add($rowDedupKey)

        $rows.Add($row)
    }

    if (-not $Global:OMMigrate.WhatIf) {
        try {
            # -- Merge with existing CSV if it exists ----------
            # Define force-Local folder names here so they are in scope for both
            # the merge retain block and the post-merge correction block.
            $forceLocalFolderNames = @(
                'Sync Issues', 'Sync Issues (This computer only)',
                'Local Failures', 'Local Failures (This computer only)',
                'Server Failures', 'Server Failures (This computer only)',
                'Conflicts', 'Recoverable Items'
            )
            $mergeStats = @{ Retained = 0; Added = 0 }

            if (Test-Path $OutputPath) {
                Write-OMMigrateLog -Message "Existing folder_map.csv found -- merging to preserve operator edits." `
                                   -Level INFO

                # Load existing CSV keyed by StoreName+FolderPath.
                # Using both columns prevents cross-account collisions when two
                # accounts have identically named folders (e.g. both have Inbox\SP).
                $existingRows = @{}
                try {
                    Import-Csv -Path $OutputPath -Encoding UTF8 |
                        ForEach-Object {
                            # Skip blank separator rows -- they are regenerated fresh
                            # on every write and must never be used as merge keys.
                            # A blank StoreName or blank FolderPath means separator row.
                            if ($_.StoreName -and $_.FolderPath -and
                                -not [string]::IsNullOrWhiteSpace($_.StoreName) -and
                                -not [string]::IsNullOrWhiteSpace($_.FolderPath)) {
                                $storeTypeVal = if ($_.PSObject.Properties['StoreType'] -and $_.StoreType) { $_.StoreType } else { '' }
                                # ADDED (fix, folder_map.csv StoreName staleness): PST rows are
                                # now keyed on StoreType+FolderPath only, not StoreName. A PST
                                # folder's StoreName can legitimately change between runs when
                                # the operator remaps the account's archive via the Script 00
                                # picker -- keying on the old StoreName would make the merge
                                # treat the same logical folder as a brand-new row and leave the
                                # old row behind as an orphan, producing a duplicate. FolderPath
                                # already begins with the account name for PST archive folders
                                # (e.g. 'user@example.com\Inbox\AAA'), so StoreType+
                                # FolderPath alone is still collision-safe. OST/Server rows keep
                                # the original StoreName-inclusive key -- unaffected by this fix.
                                $key = if ($storeTypeVal -eq 'PST') {
                                    "$storeTypeVal|$($_.FolderPath)"
                                } else {
                                    "$($_.StoreName)|$storeTypeVal|$($_.FolderPath)"
                                }
                                $existingRows[$key] = $_
                            }
                        }
                }
                catch {
                    Write-OMMigrateLog -Message "Could not read existing folder_map.csv for merge -- will overwrite: $_" `
                                       -Level WARN
                }

                if ($existingRows.Count -gt 0) {
                    $mergedRows = [System.Collections.Generic.List[PSCustomObject]]::new()

                    foreach ($newRow in $rows) {
                        # ADDED (fix, folder_map.csv StoreName staleness): lookup key must
                        # match the same PST-vs-other logic used to build $existingRows above,
                        # or a remapped PST row would never find its prior-run match.
                        $folderPath = if ($newRow.StoreType -eq 'PST') {
                            "$($newRow.StoreType)|$($newRow.FolderPath)"
                        } else {
                            "$($newRow.StoreName)|$($newRow.StoreType)|$($newRow.FolderPath)"
                        }
                        if ($existingRows.ContainsKey($folderPath)) {
                            # Folder exists -- keep operator Destination edit.
                            # OST=Server is the default for new rows; existing operator
                            # choices (including intentional Local assignments) are
                            # always preserved as-is.
                            $existing = $existingRows[$folderPath]
                            $mergedRow = $newRow.PSObject.Copy()
                            if ($existing.Destination -and
                                $existing.Destination -in @('Server','Local')) {
                                $mergedRow.Destination = $existing.Destination
                            }
                            if ($existing.Notes -and $existing.Notes -ne $newRow.Notes) {
                                $mergedRow.Notes = $existing.Notes
                            }
                            $mergedRows.Add($mergedRow)
                            $mergeStats.Retained++
                            $existingRows.Remove($folderPath)
                        }
                        else {
                            # New folder -- add with suggested default
                            $mergedRows.Add($newRow)
                            $mergeStats.Added++
                            Write-OMMigrateLog -Message "New folder added to CSV: $folderPath" -Level INFO
                        }
                    }

                    # Preserve folders from prior runs that are not in the current
                    # live session. Folders from detached stores must never be
                    # silently dropped -- they are needed for Script 03 folder
                    # migration and rule target resolution.
                    $mergeStats.PreservedPrior = 0
                    foreach ($orphanKey in @($existingRows.Keys)) {
                        $orphanFolder = $existingRows[$orphanKey]
                        # Skip blank separator rows
                        if (-not $orphanFolder.StoreName -or
                            [string]::IsNullOrWhiteSpace($orphanFolder.StoreName)) { continue }
                        # Skip old-format rows with no StoreType -- superseded by new
                        # StoreType-keyed rows from the current run. Preserving them
                        # would create duplicates and mask the correct OST/PST rows.
                        $orphanStoreType = if ($orphanFolder.PSObject.Properties['StoreType']) { $orphanFolder.StoreType } else { '' }
                        if ([string]::IsNullOrWhiteSpace($orphanStoreType)) { continue }
                        $mergedRows.Add($orphanFolder)
                        $mergeStats.PreservedPrior++
                        Write-OMMigrateLog -Message (
                            "Prior-run folder preserved (store not in live session): " +
                            "$($orphanFolder.StoreName) | $($orphanFolder.FolderPath)"
                        ) -Level DEBUG
                    }

                    $rows = $mergedRows

                    # Post-merge correction: force Local for any PST store and for
                    # known local-only Outlook system folders.
                    #
                    # PST stores are local files by definition -- they cannot sync to
                    # an IMAP server and there is no valid reason to assign any PST
                    # folder (including Inbox, Sent, Trash etc.) to Server. This covers
                    # any attached local archive PST, per-account archive PSTs,
                    # AutoArchive PSTs, and any other attached PST store.
                    #
                    # System folders (Sync Issues, Local Failures etc.) are Outlook
                    # diagnostic artifacts that should never be created on the server,
                    # regardless of store type.
                    #
                    # Only Destination is touched -- all other columns retain their
                    # merged values.
                    foreach ($row in $rows) {
                        $forceLocal = $false
                        $forceReason = ''

                        # Rule 1: PST store -- always Local
                        $rowStoreType = if ($row.PSObject.Properties['StoreType'] -and $row.StoreType) { $row.StoreType } else { '' }
                        if ($rowStoreType -eq 'PST') {
                            $forceLocal  = $true
                            $forceReason = 'PST store'
                        }
                        # Rule 2: known local-only system folder names
                        elseif ($forceLocalFolderNames -contains $row.FolderName) {
                            $forceLocal  = $true
                            $forceReason = 'system folder'
                        }

                        if ($forceLocal -and $row.Destination -eq 'Server') {
                            $row.Destination = 'Local'
                            Write-OMMigrateLog -Message (
                                "Force-Local ($forceReason): " +
                                "$($row.StoreName) | $($row.FolderPath)"
                            ) -Level DEBUG
                        }
                    }

                    Write-OMMigrateLog -Message (
                        "Folder map merge complete: $($mergeStats.Retained) retained, " +
                        "$($mergeStats.Added) new, " +
                        "$($mergeStats.PreservedPrior) preserved from prior runs"
                    ) -Level INFO
                }
            }

            # Sort by StoreName then FolderPath and insert blank separator rows
            # between account groups so the CSV is easy to review in Excel.
            $sortedRows    = @($rows | Sort-Object StoreName, FolderPath)
            $finalRows     = [System.Collections.Generic.List[PSCustomObject]]::new()
            $lastStore     = $null
            $blankTemplate = [PSCustomObject]@{
                StoreName = ''; Destination = ''; FolderPath = ''; FolderName = ''
                StoreType = ''; ItemCount = ''; UnreadCount = ''; HasSubfolders = ''
                FolderDepth = ''; IsSystemFolder = ''; EntryID = ''; Notes = ''
            }

            foreach ($row in $sortedRows) {
                if ($null -ne $lastStore -and $row.StoreName -ne $lastStore) {
                    # Insert blank separator row between account groups
                    $finalRows.Add($blankTemplate)
                }
                $finalRows.Add($row)
                $lastStore = $row.StoreName
            }

            # Check for Excel file lock before writing -- poll until released
            if (Test-Path $OutputPath) {
                $csvLocked = $true
                $lockWaited = 0
                while ($csvLocked) {
                    try {
                        $lt = [System.IO.File]::Open($OutputPath, 'Open', 'ReadWrite', 'None')
                        $lt.Close()
                        $lt.Dispose()
                        $csvLocked = $false
                    }
                    catch {
                        if ($lockWaited -eq 0) {
                            Write-OMMigrateLog -Message 'folder_map.csv is locked by Excel -- waiting for Excel to close.' `
                                               -Level WARN
                            Write-Host ''
                            Write-Host '  folder_map.csv is open in Excel.' -ForegroundColor Yellow
                            Write-Host '  Save any changes and close Excel -- script will continue automatically.' `
                                       -ForegroundColor Yellow
                        }
                        Start-Sleep -Seconds 3
                        $lockWaited += 3
                        if ($lockWaited % 15 -eq 0) {
                            Write-Host "  Still waiting for Excel to close... ($($lockWaited)s)" `
                                       -ForegroundColor DarkGray
                        }
                    }
                }
                if ($lockWaited -gt 0) {
                    Write-Host '  Excel closed -- writing folder_map.csv.' -ForegroundColor Green
                    Write-OMMigrateLog -Message "folder_map.csv lock released after $($lockWaited)s -- continuing." `
                                       -Level INFO
                }
            }

            $finalRows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            Write-OMMigrateLog -Message "Folder map CSV written: $OutputPath ($($rows.Count) folders)" `
                               -Level INFO
            Write-AuditEntry  -Action 'CSV_FOLDERMAP_WRITTEN' `
                              -Detail "Path=$OutputPath | Folders=$($rows.Count) | Retained=$($mergeStats.Retained) | Added=$($mergeStats.Added)"
        }
        catch {
            Write-OMMigrateLog -Message "Failed to write folder map CSV: $_" -Level ERROR
            throw
        }
    }
    else {
        Write-OMMigrateLog -Message "WhatIf: Would write folder map to $OutputPath ($($rows.Count) folders)" `
                           -Level INFO -WhatIfPrefix
    }

    return $OutputPath
}


# ============================================================
#  REGION: RULES INVENTORY
# ============================================================

function Get-OutlookRules {
    <#
    .SYNOPSIS
        Inventories all Outlook Rules defined in the active profile.

    .DESCRIPTION
        Reads the Outlook Rules collection via COM and returns a
        structured list of all rules with their conditions, actions,
        and target folder paths.

        Loops every account in the profile (changed 2026-06-27 -- see
        Account.DeliveryStore note above the account loop for the access
        path and rationale). Each account's rules are read using the same
        foreach -> Item(i) -> GetTable fallback chain originally built and
        proven for the default store, now simply running once per account
        instead of once total.

        Used by Script 00 to document the current rule set before
        migration, and by Script 03 to update rules whose target
        folders change during migration -- including picking up rules
        added manually in Outlook's own UI on ANY account, not just the
        default store, so they merge into rules_inventory.csv via
        Export-RulesToCSV's existing new-rule-detection logic.

        Rules that move or copy email to specific folders are flagged
        as 'NeedsFolderUpdate' -- these are the rules that Script 03
        must update after folder migration.

        Requires an active COM session from Connect-OutlookCOM.

    .OUTPUTS
        [PSCustomObject[]] -- Array of rule objects:
            RuleStoreName       [string]  Display name of the store this rule runs against
                                          (the account whose inbox this rule monitors)
            TargetStoreName     [string]  Display name of the store containing the
                                          target folder -- used by Script 03 to filter
                                          rules by selected account
            RuleName            [string]  Rule name as shown in Outlook
            IsEnabled           [bool]    Whether rule is active
            ExecutionOrder      [int]     Rule execution sequence number
            StopProcessing      [bool]    Stop processing more rules after this
            Conditions          [string]  Human-readable condition summary
            Actions             [string]  Human-readable action summary
            TargetFolderPath    [string]  Path of move/copy target folder (if any)
            TargetFolderEntryID [string]  EntryID of target folder
            NeedsFolderUpdate   [bool]    True if folder target may change in migration
            RuleType            [string]  'Incoming' | 'Outgoing' | 'Unknown'

    .EXAMPLE
        $rules = Get-OutlookRules
        $rulesNeedingUpdate = $rules | Where-Object { $_.NeedsFolderUpdate }
        Write-Host "Rules referencing folders: $($rulesNeedingUpdate.Count)"
    #>
    [CmdletBinding()]
    param()

    $namespace = Get-OutlookNamespace
    if (-not $namespace) { return @() }

    Write-OMMigrateLog -Message 'Inventorying Outlook Rules across all stores...' -Level INFO

    $ruleObjects = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        # Loop every account in the profile, reading each one's rules via
        # Account.DeliveryStore.GetRules() -- the same access path proven
        # safe and working in yesterday's Invoke-DeployConsolidatedRules
        # engine (2026-06-26/27). That engine never reads/enumerates the
        # collection, only uses it as a Create() target -- but the access
        # path itself (DeliveryStore.GetRules(), not $namespace.Stores
        # enumeration) is the same one used here. Per-store body below
        # (foreach -> Item(i) -> GetTable fallback chain) is UNCHANGED from
        # the original default-store-only version -- this only changes how
        # the store/account is selected, not how its rules are read once
        # selected. Administrator confirmed (2026-06-27) this same fallback chain
        # already succeeds often enough in production on the default store
        # (ameritech) over weeks of real use to keep rules_inventory.csv
        # current -- secondary accounts (all confirmed under 100 rules,
        # smaller than ameritech) get the identical treatment here.
        $accounts = $namespace.Session.Accounts
        $accountCount = 0
        try { $accountCount = $accounts.Count } catch { }
        Write-OMMigrateLog -Message "Inventorying rules across $accountCount account(s)..." -Level INFO

        for ($acctIdx = 1; $acctIdx -le $accountCount; $acctIdx++) {
        $acct = $null
        try { $acct = $accounts.Item($acctIdx) } catch { }
        if (-not $acct) { continue }

        $store = $null
        try {
            if ($acct.DeliveryStore -and $acct.DeliveryStore.GetRootFolder()) {
                $store = $acct.DeliveryStore
                Register-COMObject -ComObject $store
            }
        }
        catch {
            # Bumped DEBUG -> INFO (2026-07-01, Administrator): this catch was silently
            # swallowing DeliveryStore resolution failures at DEBUG level, which
            # is why ameritech dropping out of the rules scan produced no visible
            # log line at all in the 2026-07-01 "TestProfile" run. INFO makes a skip
            # actually visible in a normal run without changing any behavior.
            Write-OMMigrateLog -Message "Account [$acctIdx]: failed to resolve DeliveryStore for rules scan: $_" -Level INFO
        }

        # Retry loop REMOVED (2026-07-06, Administrator, 4th pass): confirmed live
        # (memory #28) that for the known group of 7 broken accounts,
        # Account.DeliveryStore is genuinely $null -- not a transient timing
        # race -- so two retries at 750ms each never helped and only added
        # ~1.5s of dead time per broken account. Going straight from the
        # single direct attempt above to the DefaultStore/DisplayName
        # fallbacks below. If a genuine timing race ever resurfaces for a
        # DIFFERENT account (the original motivation was admin@example-provider.com
        # appearing to resolve intermittently), the DisplayName fallback
        # below should catch it immediately on the same run rather than
        # needing a delayed retry to help.
        if (-not $store) {
            # Bumped DEBUG -> INFO (2026-07-01, Administrator): same rationale as the
            # catch block above -- this is the other silent-skip path for this
            # account resolution step.
            Write-OMMigrateLog -Message "Account [$acctIdx]: Reverting to DeliveryStore fallback." -Level INFO

            # Fallback (added 2026-07-01, Administrator -- re-added same day after being
            # missed in an earlier edit pass): Account.DeliveryStore does not
            # reliably resolve for every account in every profile -- confirmed
            # in "TestProfile", ameritech account. If this account's SmtpAddress
            # matches namespace.DefaultStore's owning account, retry via
            # namespace.DefaultStore instead of skipping the account's rules
            # scan entirely. Same fallback pattern already applied in
            # Invoke-DeployConsolidatedRules's primary-store resolution.
            try {
                $defaultStoreDisplay = ''
                try { $defaultStoreDisplay = ([string]$namespace.DefaultStore.DisplayName).ToLower() } catch { }
                $thisAcctSmtp = ''
                try { $thisAcctSmtp = ([string]$acct.SmtpAddress).ToLower() } catch { }

                if ($namespace.DefaultStore -and
                    (-not [string]::IsNullOrWhiteSpace($thisAcctSmtp)) -and
                    $defaultStoreDisplay -and
                    ($defaultStoreDisplay.Contains($thisAcctSmtp) -or $thisAcctSmtp.Contains($defaultStoreDisplay))) {
                    $store = $namespace.DefaultStore
                    Register-COMObject -ComObject $store
                    Write-OMMigrateLog -Message "Account [$acctIdx]: DeliveryStore fallback -- resolved via namespace.DefaultStore instead." -Level INFO
                }
            }
            catch {
                Write-OMMigrateLog -Message "Account [$acctIdx]: DefaultStore fallback also failed: $_" -Level DEBUG
            }

            # Fallback (added 2026-07-06, Administrator): DisplayName match against
            # namespace.Stores. Confirmed live (2026-07-06 session, memory #28)
            # that Account.DeliveryStore is genuinely $null -- not a transient
            # timing issue -- for a known group of 7 accounts (registry
            # confirmed missing "Delivery Store EntryID"/"Delivery Folder
            # EntryID" for these same accounts). Neither retrying above nor
            # the DefaultStore fallback above can help these accounts, since
            # they are secondary stores, not the default store. This fallback
            # scans every mounted store in the profile (namespace.Stores --
            # a broader collection than Accounts, includes backup/archive
            # PSTs) and matches by DisplayName against this account's own
            # DisplayName/SmtpAddress, as a last resort before giving up.
            # DisplayName matching is a string match, not a guaranteed-unique
            # key -- acceptable here given this project's Backup/ArchiveBuild
            # naming convention is very unlikely to collide with a plain
            # account DisplayName, but worth remembering if a false-positive
            # match is ever suspected. Only runs if the two attempts above
            # (direct + DefaultStore) both failed.
            if (-not $store) {
                try {
                    $thisAcctDisplay = ''
                    try { $thisAcctDisplay = ([string]$acct.DisplayName).ToLower() } catch { }
                    $thisAcctSmtpForStores = ''
                    try { $thisAcctSmtpForStores = ([string]$acct.SmtpAddress).ToLower() } catch { }

                    foreach ($candidateStore in $namespace.Stores) {
                        $candidateDisplay = ''
                        try { $candidateDisplay = ([string]$candidateStore.DisplayName).ToLower() } catch { }
                        if (-not $candidateDisplay) { continue }

                        $isMatch = $false
                        if ($thisAcctDisplay -and $candidateDisplay -eq $thisAcctDisplay) { $isMatch = $true }
                        if (-not $isMatch -and $thisAcctSmtpForStores -and $candidateDisplay -eq $thisAcctSmtpForStores) { $isMatch = $true }

                        if ($isMatch) {
                            $store = $candidateStore
                            Register-COMObject -ComObject $store
                            Write-OMMigrateLog -Message "Account [$acctIdx]: DeliveryStore fallback -- namespace.Stores DisplayName match found. Looked for DisplayName='$thisAcctDisplay' or SmtpAddress='$thisAcctSmtpForStores'. Matched store: '$candidateDisplay'." -Level INFO
                            break
                        }
                    }
                    if (-not $store) {
                        # Debug line (added 2026-07-06, Administrator, 3rd pass): the
                        # fallback loop ran to completion but found no matching
                        # store -- distinct from an exception, and distinct from
                        # never attempting the fallback at all. Logs the account's
                        # own comparison keys plus every store DisplayName seen,
                        # so a no-match result is fully diagnosable from the log
                        # alone without needing a live debugging session.
                        $allStoreDisplaysForLog = @()
                        foreach ($s in $namespace.Stores) {
                            try { $allStoreDisplaysForLog += [string]$s.DisplayName } catch { }
                        }
                        Write-OMMigrateLog -Message "Account [$acctIdx]: DeliveryStore fallback -- namespace.Stores DisplayName match NOT found. Looked for DisplayName='$thisAcctDisplay' or SmtpAddress='$thisAcctSmtpForStores'. Stores seen: $($allStoreDisplaysForLog -join ' | ')." -Level INFO
                    }
                }
                catch {
                    Write-OMMigrateLog -Message "Account [$acctIdx]: namespace.Stores DisplayName fallback also failed: $_" -Level INFO
                }
            }

            if (-not $store) {
                continue
            }
        }

        if ($store) {
            # Get store display name for tagging rules
            $storeName = ''
            try { $storeName = $store.DisplayName } catch { }

            # Get rules collection from this account's store
            $rules = $null
            try {
                $rules = $store.GetRules()
                Register-COMObject -ComObject $rules
            }
            catch {
                # INFO, not ERROR/WARN (2026-06-27): this store genuinely does
                # not support rules at all -- e.g. a PST store ("This store
                # does not support rules. Could not complete the operation.",
                # confirmed via direct MAPI testing 2026-06-27). This is a
                # known, expected condition for non-mail-server stores, not
                # something an admin needs to act on. Per Administrator's standing
                # instruction: known permanent/expected conditions use INFO
                # regardless of console color, so the run's error/warning
                # totals stay meaningful (0/0 = "did everything it could").
                Write-OMMigrateLog -Message "Account store '$storeName' does not support GetRules() (expected for non-mail stores): $_" `
                                   -Level INFO
            }

            if ($rules) {
                $ruleCount = $rules.Count
                Write-OMMigrateLog -Message "Account store '$storeName': $ruleCount rule(s) found." -Level INFO

                # Try foreach enumeration -- uses IEnumVARIANT interface.
                # Wrapped in try/catch: if foreach throws (e.g. 0x800C8101 on
                # IMAP stores with 500+ programmatically-created rules that exceed
                # the PR_RULES_DATA buffer limit), fall through to GetTable fallback.
                $foreachFailed = $false
                try {
                foreach ($rule in $rules) {
                    if (-not $rule) { continue }
                    try {
                        Register-COMObject -ComObject $rule

                        $ruleName      = $rule.Name
                        # IsEnabled is always True -- never read from COM.
                        # COM can return False for recovered or non-default-store rules.
                        # All rules are active by design; False is never a valid CSV value.
                        $isEnabled     = $true
                        $execOrder     = $rule.ExecutionOrder


                        # Determine rule type
                        $ruleType = 'Unknown'
                        try {
                            $ruleType = if ($rule.RuleType -eq 0) { 'Incoming' }
                                        elseif ($rule.RuleType -eq 1) { 'Outgoing' }
                                        else { 'Unknown' }
                        }
                        catch { }

                        # Build conditions summary
                        $conditionsSummary = Get-RuleConditionsSummary -Rule $rule

                        # NEW (2026-07-23, Administrator direction): read the rule's own
                        # SenderAddress condition words separately from the
                        # human-readable Conditions summary above, so Script 00
                        # can populate SendersDomain from real rule data when
                        # present (see Get-RuleSenderAddressWords header).
                        $senderAddressWords = Get-RuleSenderAddressWords -Rule $rule

                        # Build actions summary and extract folder targets
                        $actionInfo = Get-RuleActionsSummary -Rule $rule

                        # StopProcessing flag
                        # Default true -- all rules were created as client-side POP3 rules
                        # with stop-processing enabled. COM may return False for recovered
                        # IPM.RuleOrganizer rules where the flag was not preserved correctly.
                        # Default is preserved deliberately (Administrator's explicit instruction,
                        # 2026-06-29) -- a failed/null reflection read must NOT flip this
                        # to false; it should only be overwritten when a real value is
                        # actually confirmed via reflection below.
                        #
                        # Updated 2026-06-29 (Gemini consult + Administrator review): replaced the
                        # June 17-era Actions.Item(27) fixed-index lookup with the named
                        # property 'Stop', resolved via InvokeMember reflection -- the
                        # same mechanism already proven and in production use on the
                        # write path (Invoke-BuildRulesFromMap, 2026-06-26). Item(27)
                        # relied on a hardcoded action-slot offset that is fragile to
                        # rule-structure or profile-context shifts; the named property
                        # is the safer, more accurate accessor and is no longer null
                        # via COM interop once the rule has actually been Saved.
                        $stopProcessing = $true
                        try {
                            $stopAction = $rule.Actions.GetType().InvokeMember(
                                'Stop',
                                [System.Reflection.BindingFlags]::GetProperty,
                                $null, $rule.Actions, $null
                            )
                            if ($null -ne $stopAction) {
                                $stopProcessing = [bool]$stopAction.GetType().InvokeMember(
                                    'Enabled',
                                    [System.Reflection.BindingFlags]::GetProperty,
                                    $null, $stopAction, $null
                                )
                            }
                        }
                        catch {
                            Write-OMMigrateLog -Message "Failed to read StopProcessing action via reflection for rule '$ruleName': $_" -Level DEBUG
                        }

                        # -- Qualify partial TargetFolderPath with TargetStoreName prefix --
                        # When the COM folder object is inaccessible (e.g. rules targeting
                        # non-default IMAP store folders), Get-FolderFullPath returns a partial
                        # path with no account prefix (e.g. 'Inbox\FolderName'). If TargetStoreName
                        # is a valid email address and the first segment of TargetFolderPath
                        # contains no '@', prepend TargetStoreName to produce a fully-qualified
                        # store-relative path ('account@domain.com\Inbox\FolderName').
                        # Universal -- no account addresses are hardcoded.
                        $qualifiedTargetFolderPath = $actionInfo.TargetFolderPath
                        if (-not [string]::IsNullOrWhiteSpace($qualifiedTargetFolderPath) -and
                            $actionInfo.TargetStoreName -like '*@*') {
                            $firstSegment = ($qualifiedTargetFolderPath -split '\\')[0]
                            if ($firstSegment -notlike '*@*') {
                                $qualifiedTargetFolderPath = "$($actionInfo.TargetStoreName)\$qualifiedTargetFolderPath"
                            }
                        }

                        # -- Suggested path for inaccessible folder targets --
                        # When COM cannot return the folder object for a move/copy action
                        # (e.g. rules targeting non-default IMAP store folders), the
                        # TargetFolderPath will still be blank after qualification above.
                        # In this case, attempt to build a best-guess suggested path from:
                        #   - $storeName  : the account the rule runs against (RuleStoreName)
                        #   - 'Inbox'     : the most common parent for rule target folders
                        #   - folder name : extracted from the action summary text
                        # The suggested path is written to TargetFolderPath so Script 03
                        # has something to work with. The Notes column flags it for
                        # operator verification before Script 03 runs.
                        # Universal -- no account addresses or folder names are hardcoded.
                        $ruleNotes = ''
                        # $needsFolderUpdate is set unconditionally to $true below after path
                        # resolution. Initialized here only to satisfy variable scoping.
                        $needsFolderUpdate = $true
                        if ((-not ($qualifiedTargetFolderPath -like '*@*' -and $qualifiedTargetFolderPath -like '*\*')) -and
                            $storeName -like '*@*') {
                            # Try to extract folder name from action summary
                            # Summary format: "Move to: FolderName" or "Copy to: FolderName"
                            $extractedFolderName = ''
                            if ($actionInfo.Summary -match '(?:Move|Copy) to: (.+?)(?:;|$)') {
                                $extractedFolderName = $Matches[1].Trim()
                            }
                            # If regex found nothing, fall back to the leaf segment of the
                            # partial path already in $qualifiedTargetFolderPath (e.g. the
                            # folder name from the catch block fallback in Get-FolderFullPath).
                            if ([string]::IsNullOrWhiteSpace($extractedFolderName) -and
                                -not [string]::IsNullOrWhiteSpace($qualifiedTargetFolderPath)) {
                                $extractedFolderName = ($qualifiedTargetFolderPath -split '\\')[-1]
                            }
                            if (-not [string]::IsNullOrWhiteSpace($extractedFolderName)) {
                                $qualifiedTargetFolderPath = "$storeName\Inbox\$extractedFolderName"
                                $needsFolderUpdate         = $true
                                $ruleNotes                 = "VERIFY REQUIRED: Target folder inaccessible via COM. " +
                                                             "Suggested path auto-generated from rule store and action text. " +
                                                             "Confirm destination in Outlook Rules and Alerts and correct " +
                                                             "TargetFolderPath if wrong before running Script 03."
                                Write-OMMigrateLog -Message (
                                    "Suggested path generated for inaccessible folder target: " +
                                    "Rule='$ruleName' | Suggested='$qualifiedTargetFolderPath'"
                                ) -Level WARN
                            }
                        }

                        # Every rule always gets NeedsFolderUpdate=True -- no exceptions.
                        # Rules with empty paths still need Script 03 attention; the empty
                        # path is the signal, not this flag. Never set False from code.
                        $needsFolderUpdate = $true

                        $ruleObj = [PSCustomObject]@{
                            # -- Account grouping columns first -- makes Excel review easy --
                            # RuleStoreName  = the account this rule fires for (from account condition)
                            # TargetStoreName = the store containing the target folder (where mail goes)
                            # TargetFolderPath placed early so duplicates are easy to spot and fix
                            RuleStoreName         = $storeName
                            TargetStoreName       = $actionInfo.TargetStoreName
                            RuleName              = $ruleName
                            LastDeployedRun       = ''             # Timestamp set by Script 03 after successful secondary store Save()
                            # NEW (2026-07-07, Administrator direction): LastTargetRun -- separate
                            # idempotency column from LastDeployedRun, tracking Script 03
                            # Phase 3's folder-target (TargetFolderPath) remap work
                            # specifically, independent of LastDeployedRun's consolidation/
                            # condition-validation tracking. See Set-RuleConditions header
                            # comment for the full rationale.
                            LastTargetRun         = ''             # Timestamp set by Script 03 Phase 3 after successful folder-target remap
                            TargetFolderPath      = $qualifiedTargetFolderPath
                            # NEW (2026-07-23, Administrator direction): populated here from the rule's own
                            # SenderAddress condition words when present; Script 00 still applies
                            # its TargetFolderPath-last-segment fallback when this is blank.
                            SendersDomain         = $senderAddressWords   # Populated by Script 00 -- default = last segment of TargetFolderPath
                            NeedsFolderUpdate     = $needsFolderUpdate
                            IsEnabled             = $isEnabled
                            ExecutionOrder        = $execOrder
                            RuleType              = $ruleType
                            StopProcessing        = $stopProcessing
                            Conditions            = $conditionsSummary
                            Actions               = $actionInfo.Summary
                            TargetFolderEntryID   = $actionInfo.TargetFolderEntryID
                            Notes                 = $ruleNotes   # Operator annotation -- preserved on re-run
                        }

                        $ruleObjects.Add($ruleObj)

                        # Suppress per-rule detail when Sanitize is active --
                        # rule names are not yet in the sanitize map at this point.
                        if (-not $Global:OMMigrate.Sanitize) {
                            Write-OMMigrateLog -Message (
                                "Rule Store='$storeName': '$ruleName' | Enabled=$isEnabled | " +
                                "Type=$ruleType | FolderTarget=$($actionInfo.TargetFolderPath)"
                            ) -Level DEBUG
                        }
                    }
                    catch {
                        Write-OMMigrateLog -Message "Error reading rule from store '$storeName': $_" `
                                           -Level INFO
                    }
                    finally {
                        # Release rule COM object after each iteration.
                        # Gemini: 573 rules x nested condition/action refs = thousands of
                        # pinned MAPI pointers. Without explicit release, subsequent COM
                        # sessions hit MAPI_E_BUSY on Item() access.
                        if ($rule) {
                            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rule) } catch { }
                        }
                    }
                }
                # Release rules collection and force GC to purge all remaining
                # MAPI pointers before any subsequent COM session accesses rules.
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rules) } catch { }
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                Write-OMMigrateLog -Message 'Rules COM objects released and GC completed.' -Level DEBUG
                } catch {
                    # foreach enumeration failed -- try Item($i) loop as second attempt
                    # before falling through to GetTable. The diagnostic script confirmed
                    # that GetRules().Item(1) works even when foreach fails -- the IEnumVARIANT
                    # interface breaks at high rule counts but indexed access still works.
                    $foreachFailed = $true
                    Write-OMMigrateLog -Message "foreach enumeration failed ($_) -- trying Item(i) indexed loop as second attempt." -Level INFO
                    # Clear any partial results from the failed foreach before Item(i) loop
                    $ruleObjects.Clear()
                    try {
                        for ($ri = 1; $ri -le $ruleCount; $ri++) {
                            $rule = $null
                            try {
                                $rule = $rules.Item($ri)
                                if (-not $rule) { continue }
                                Register-COMObject -ComObject $rule

                                $ruleName      = $rule.Name
                                $executionOrder = 0
                                try { $executionOrder = $rule.ExecutionOrder } catch { }
                                $isEnabled     = $true
                                # $rule.StopProcessing is NOT a real Outlook Object Model
                                # property (confirmed via Gemini consult, 2026-06-29) --
                                # Microsoft.Office.Interop.Outlook.Rule has no such property.
                                # StopProcessing lives only inside Rule.Actions as the Stop
                                # action (ActionType = olRuleActionStopProcessingRules).
                                # The old line below was a silent failure masquerading as
                                # a default value: $rule.StopProcessing always threw a
                                # missing-property error, the empty catch{} swallowed it,
                                # and $stopProcessing simply kept its hardcoded $true
                                # default on every single rule, every time, with no actual
                                # read ever happening. Replaced with the same named-property
                                # reflection approach used in the main foreach branch above
                                # and in the write path (Invoke-BuildRulesFromMap,
                                # 2026-06-26). Default remains $true deliberately -- a
                                # failed/null reflection read must not flip this to false.
                                $stopProcessing = $true
                                try {
                                    $stopActionFb = $rule.Actions.GetType().InvokeMember(
                                        'Stop',
                                        [System.Reflection.BindingFlags]::GetProperty,
                                        $null, $rule.Actions, $null
                                    )
                                    if ($null -ne $stopActionFb) {
                                        $stopProcessing = [bool]$stopActionFb.GetType().InvokeMember(
                                            'Enabled',
                                            [System.Reflection.BindingFlags]::GetProperty,
                                            $null, $stopActionFb, $null
                                        )
                                    }
                                }
                                catch {
                                    Write-OMMigrateLog -Message "Failed to read StopProcessing action via reflection (Item(i) fallback) for rule '$ruleName': $_" -Level DEBUG
                                }
                                $ruleType = 'Incoming'
                                try { if ($rule.RuleType -eq 1) { $ruleType = 'Outgoing' } } catch { }

                                # MoveToFolder action
                                $targetPath      = ''
                                $targetEntryID   = ''
                                $targetStoreName = ''
                                $actions = ''
                                $conditions = ''
                                try {
                                    $moveAction = $rule.Actions.MoveToFolder
                                    if ($moveAction -and $moveAction.Enabled) {
                                        $tgtFolder = $moveAction.Folder
                                        if ($tgtFolder) {
                                            $targetPath    = Get-FolderFullPath -Folder $tgtFolder
                                            $targetPath    = Remove-StorePrefix  -Path $targetPath
                                            try { $targetEntryID   = $tgtFolder.EntryID }    catch { }
                                            try { $targetStoreName = $tgtFolder.Store.DisplayName } catch { }
                                        }
                                    }
                                } catch {
                                    Write-OMMigrateLog -Message "DIAGNOSTIC: Item(i) fallback folder-path read failed for rule '$ruleName': $_" -Level ERROR
                                }

                                $ruleObj = [PSCustomObject]@{
                                    RuleStoreName         = $storeName
                                    TargetStoreName       = $targetStoreName
                                    RuleName              = $ruleName
                                    LastDeployedRun       = ''
                                    LastTargetRun         = ''      # See fresh-scan template above for full rationale
                                    TargetFolderPath      = $targetPath
                                    SendersDomain         = ''
                                    NeedsFolderUpdate     = $true
                                    IsEnabled             = $true
                                    ExecutionOrder        = $executionOrder
                                    RuleType              = $ruleType
                                    StopProcessing        = $stopProcessing
                                    Conditions            = $conditions
                                    Actions               = $actions
                                    TargetFolderEntryID   = $targetEntryID
                                    Notes                 = 'Item(i) fallback: foreach enumeration failed.'
                                }
                                $ruleObjects.Add($ruleObj)
                                $foreachFailed = $false   # Mark as recovered
                            }
                            catch {
                                Write-OMMigrateLog -Message "Item($ri) failed: $_" -Level DEBUG
                            }
                        }
                        if (-not $foreachFailed) {
                            Write-OMMigrateLog -Message "Item(i) indexed loop succeeded: $($ruleObjects.Count) rules recovered." -Level INFO
                        }
                        else {
                            Write-OMMigrateLog -Message 'Item(i) indexed loop also failed -- falling through to GetTable.' -Level INFO
                        }
                    }
                    catch {
                        Write-OMMigrateLog -Message "Item(i) loop failed: $_ -- falling through to GetTable." -Level INFO
                        $foreachFailed = $true
                    }
                }

                # -- GetTable fallback -- reads rule names from hidden MAPI messages --
                # Gemini: rules are stored as hidden messages (IPM.Rule.Version2.Message)
                # in the Inbox Associated Contents table. GetTable(olHiddenItems) reads
                # them when the Rules COM collection is inaccessible due to buffer overflow.
                # Limitation: only rule names available this way -- no conditions/actions.
                # TargetFolderPath will be blank and must be corrected via rules_inventory.csv.
                if ($foreachFailed -and $ruleObjects.Count -eq 0) {
                    Write-OMMigrateLog -Message "GetTable fallback: reading rule names from Inbox hidden items..." -Level INFO
                    try {
                        # Use this account's own store Inbox, not the namespace default --
                        # critical now that this runs per-account rather than default-store-only.
                        # $namespace.GetDefaultFolder(6) would silently read the wrong account's
                        # hidden rule items for every non-default account in the loop.
                        $inbox = $store.GetDefaultFolder(6)  # 6 = olFolderInbox
                        $table = $inbox.GetTable('', 2)          # 2 = olHiddenItems
                        $fallbackCount = 0
                        while (-not $table.EndOfTable) {
                            $row = $table.GetNextRow()
                            $msgClass = ''
                            try { $msgClass = $row.Item('MessageClass') } catch { }
                            # Only process rule messages
                            if ($msgClass -notlike 'IPM.Rule*') { continue }
                            $ruleName = ''
                            try { $ruleName = $row.Item('Subject') } catch { }
                            if ([string]::IsNullOrWhiteSpace($ruleName)) { continue }
                            $ruleObj = [PSCustomObject]@{
                                RuleStoreName         = $storeName
                                TargetStoreName       = ''
                                RuleName              = $ruleName
                                LastDeployedRun       = ''
                                LastTargetRun         = ''      # See fresh-scan template above for full rationale
                                TargetFolderPath      = ''
                                SendersDomain         = ''
                                NeedsFolderUpdate     = $true
                                IsEnabled             = $true
                                ExecutionOrder        = 0
                                RuleType              = 'Incoming'
                                StopProcessing        = $true
                                Conditions            = ''
                                Actions               = ''
                                TargetFolderEntryID   = ''
                                Notes                 = 'GetTable fallback: TargetFolderPath unavailable. Correct in rules_inventory.csv before running Script 03.'
                            }
                            $ruleObjects.Add($ruleObj)
                            $fallbackCount++
                        }
                        Write-OMMigrateLog -Message "GetTable fallback: read $fallbackCount rule name(s) from hidden items." -Level INFO
                    } catch {
                        # INFO, not ERROR (2026-06-27): this is the final step of
                        # the foreach -> Item(i) -> GetTable fallback chain for
                        # the 0x800C8101 "devil code" -- a known, permanent,
                        # intermittent COM enumeration condition on this Outlook
                        # installation (confirmed across many independent tests
                        # 2026-06-27: occurs via every connection method, every
                        # process model, and even via GetTable -- not specific
                        # to any one account or access path). Per Administrator's explicit,
                        # repeated standing instruction: this error code must be
                        # ignored and the logic allowed to continue -- it is not
                        # an admin-actionable failure. The loop already continues
                        # to the next account regardless; this only fixes the
                        # log severity so the run's error/warning totals don't
                        # misrepresent a known, ignorable condition as a real
                        # problem.
                        Write-OMMigrateLog -Message "GetTable fallback also unsuccessful for '$storeName' (0x800C8101 devil code -- known, ignorable, continuing): $_" -Level INFO
                    }
                }

                # -- Rules cross-check verification gate (added 2026-06-29, Gemini consult) --
                # Read-only sanity check, runs whenever the foreach/Item(i) enumeration
                # chain above failed for this store (whether or not the GetTable fallback
                # itself succeeded). NEVER calls SetProperty() or Save() -- pure read,
                # cannot mutate the file. Compares two independent views of "how many
                # rules does this store have":
                #   1. IRulesCollection.Count -- proven to work even when .Item() throws
                #      0x800C8101 (confirmed across this project's history).
                #   2. The rule-count field at bytes 44-45 of the raw PR_RW_RULES_STREAM,
                #      read via PropertyAccessor.GetProperty() on the IPM.RuleOrganizer
                #      item itself -- NOT via $store.PropertyAccessor, which does not
                #      exist; PropertyAccessor belongs to the item, not the Store object.
                #      Empirically confirmed working 2026-06-29 (Test-PropertyAccessorRead
                #      RulesStream.ps1, live run: 198109 bytes read in 14ms, bytes 44-45 =
                #      353, matching the account's actual rule count) -- this directly
                #      disproved an earlier Gemini claim that PropertyAccessor would fail
                #      identically to MFCMAPI's flat property-grid read on the same stream.
                # A count match is a mild positive signal that the two independent views
                # of the rules data agree -- it is NOT proof that any other operation
                # (e.g. .rwz export/import) will succeed; it only confirms count parity.
                if ($foreachFailed) {
                    # Explicit COM handle tracking (added 2026-06-29, Gemini review pass 2) --
                    # the existing GetTable fallback immediately above this block, and the
                    # identical lookup pattern in Invoke-PurgeAndRecreateRules, both rely on
                    # this function's existing end-of-loop GC.Collect()/WaitForPendingFinalizers()
                    # calls rather than explicit per-object release -- there is no evidence in
                    # this project's history that this has caused lock/hang issues in practice.
                    # Still, explicit release is better hygiene than relying on GC timing inside
                    # a loop, so it's added here for this new code without changing the
                    # established pattern elsewhere in this file.
                    $crossCheckTbl  = $null
                    $ccRow          = $null
                    $crossCheckItem = $null
                    try {
                        $crossCheckComCount = $null
                        try { $crossCheckComCount = $rules.Count } catch { }

                        $ruleOrgEntryID = ''
                        $crossCheckTbl  = $inbox.GetTable("[MessageClass] = 'IPM.RuleOrganizer'", 1)  # 1 = olFormItem
                        while (-not $crossCheckTbl.EndOfTable) {
                            # Release the previous row before pulling the next one.
                            if ($null -ne $ccRow) {
                                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ccRow) } catch { }
                                $ccRow = $null
                            }
                            $ccRow = $crossCheckTbl.GetNextRow()
                            $ccmc  = ''; try { $ccmc = $ccRow['MessageClass'] } catch { }
                            $ccei  = ''; try { $ccei = $ccRow['EntryID']      } catch { }
                            if ($ccmc -eq 'IPM.RuleOrganizer') { $ruleOrgEntryID = $ccei; break }
                        }

                        if (-not [string]::IsNullOrWhiteSpace($ruleOrgEntryID)) {
                            $crossCheckItem   = $namespace.GetItemFromID($ruleOrgEntryID, $store.StoreID)
                            $crossCheckStream = $crossCheckItem.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x68020102')

                            if ($null -ne $crossCheckStream -and $crossCheckStream.Length -ge 46) {
                                # Bytes 44-45, little-endian uint16 -- same field this
                                # project has used for the persistence-patch work all
                                # along. No byte reversal -- straight little-endian read.
                                $crossCheckBinaryCount = [System.BitConverter]::ToUInt16($crossCheckStream, 44)

                                Write-OMMigrateLog -Message "Rules cross-check for '$storeName': COM Count=$crossCheckComCount | Stream byte-count=$crossCheckBinaryCount" -Level DEBUG

                                if ($crossCheckComCount -eq $crossCheckBinaryCount) {
                                    Write-OMMigrateLog -Message "Rules data stream structurally consistent for '$storeName' ($crossCheckBinaryCount rules) -- COM count and stream byte-count agree." -Level INFO
                                } else {
                                    Write-OMMigrateLog -Message "Rule count mismatch for '$storeName': COM Count=$crossCheckComCount vs Stream byte-count=$crossCheckBinaryCount -- potential drift, worth investigating." -Level WARN
                                }
                            }
                        }
                    }
                    catch {
                        Write-OMMigrateLog -Message "Rules cross-check failed for '$storeName' (non-fatal, read-only, no impact on rules data): $_" -Level DEBUG
                    }
                    finally {
                        # Explicit teardown for this new block's COM handles -- does not
                        # change the established (GC-reliant) pattern used elsewhere in
                        # this file's other GetTable call sites.
                        if ($null -ne $ccRow)          { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ccRow) }          catch { } }
                        if ($null -ne $crossCheckTbl)   { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($crossCheckTbl) }   catch { } }
                        if ($null -ne $crossCheckItem)  { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($crossCheckItem) }  catch { } }
                    }
                }

            }
        }

        }   # end for ($acctIdx ...) -- per-account loop

        # Release remaining MAPI pointers and force GC once, after the
        # entire account loop completes (not per-account).
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-OMMigrateLog -Message 'Rules COM objects released and GC completed.' -Level DEBUG
    }
    catch {
        Write-OMMigrateLog -Message "Failed to read Outlook Rules: $_" -Level ERROR
    }

    Write-OMMigrateLog -Message "Rules inventoried (raw): $($ruleObjects.Count) total | Rules with folder targets: $(@($ruleObjects | Where-Object {$_.NeedsFolderUpdate}).Count)" -Level INFO

    # -- Deduplicate by RuleStoreName+RuleName -- Outlook can contain duplicate
    # rule entries within the same account (e.g. from .rwz import adding a
    # second copy alongside originals). Keep exactly one row per
    # RuleStoreName+RuleName using priority scoring:
    #   8 pts -- valid TargetFolderPath (contains @ and \)
    #   4 pts -- has real actions (not '[No actions]')
    #   2 pts -- NeedsFolderUpdate = True
    #   1 pt  -- IsEnabled = True
    # Highest score wins. Ties go to the first encountered.
    #
    # Key changed 2026-06-27 from RuleName-only to RuleStoreName+RuleName --
    # required now that Get-OutlookRules scans every account, not just the
    # default store. Two different accounts can legitimately have identically
    # named rules (e.g. "Move to Archive"); a RuleName-only key would have
    # silently collapsed them into a single row and dropped one account's
    # rule entirely.
    $dedupedRules = [System.Collections.Generic.List[PSCustomObject]]::new()
    $bestByName   = @{}
    $bestScore    = @{}

    foreach ($r in $ruleObjects) {
        $score = 0
        $hasValidPath = (
            -not [string]::IsNullOrWhiteSpace($r.TargetFolderPath) -and
            $r.TargetFolderPath -like '*@*' -and
            $r.TargetFolderPath -like '*\*'
        )
        $hasActions = (
            -not [string]::IsNullOrWhiteSpace($r.Actions) -and
            $r.Actions -ne '[No actions]'
        )
        if ($hasValidPath)                              { $score += 8 }
        if ($hasActions)                               { $score += 4 }
        if ($r.NeedsFolderUpdate -eq $true -or
            $r.NeedsFolderUpdate -eq 'True')            { $score += 2 }
        if ($r.IsEnabled -eq $true -or
            $r.IsEnabled -eq 'True')                    { $score += 1 }

        $dedupKey = "$($r.RuleStoreName)|$($r.RuleName)"
        if (-not $bestByName.ContainsKey($dedupKey) -or $score -gt $bestScore[$dedupKey]) {
            $bestByName[$dedupKey]  = $r
            $bestScore[$dedupKey]   = $score
        }
    }

    foreach ($dedupKey in $bestByName.Keys) {
        $dedupedRules.Add($bestByName[$dedupKey])
    }

    $dupsRemoved = $ruleObjects.Count - $dedupedRules.Count
    if ($dupsRemoved -gt 0) {
        Write-OMMigrateLog -Message "Rules deduplication: $dupsRemoved duplicate(s) removed | $($dedupedRules.Count) unique rules remaining." -Level INFO
    }

    Write-OMMigrateLog -Message "Rules inventoried (final): $($dedupedRules.Count) total | Rules with folder targets: $(@($dedupedRules | Where-Object {$_.NeedsFolderUpdate}).Count)" -Level INFO

    return $dedupedRules
}


function Get-RuleConditionsSummary {
    <#
    .SYNOPSIS
        Returns a human-readable summary of a rule's conditions.

    .DESCRIPTION
        Iterates the rule's Conditions collection and builds a
        plain-language summary string for the discovery report.
        Used internally by Get-OutlookRules.

    .PARAMETER Rule
        An Outlook Rule COM object.

    .OUTPUTS
        [string] -- Human-readable conditions summary.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Rule
    )

    $parts = [System.Collections.Generic.List[string]]::new()

    try {
        $conditions = $Rule.Conditions

        # From condition
        $fromCond = $null
        try {
            $fromCond = $conditions.From
            if ($fromCond.Enabled) {
                $recipients = $fromCond.Recipients
                $senders = @($recipients) -join ', '
                $parts.Add("From: $senders")
                # Release recipients collection -- Gemini: nested COM refs pin MAPI objects
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($recipients) } catch { }
            }
        }
        catch { }
        finally { if ($fromCond) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($fromCond) } catch { } } }

        # Subject condition
        $subjCond = $null
        try {
            $subjCond = $conditions.Subject
            if ($subjCond.Enabled) {
                $subjects = @($subjCond.Text) -join ', '
                $parts.Add("Subject contains: $subjects")
            }
        }
        catch { }
        finally { if ($subjCond) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($subjCond) } catch { } } }

        # Sender address condition
        # FIXED (2026-07-07, Administrator debug session): converted from plain PS
        # dot-notation ($conditions.SenderAddress / .Enabled / .Address) to
        # InvokeMember GetProperty reflection, per this module's standing
        # COM rule (applies to ALL Outlook COM rule object property access,
        # not just Invoke-BuildRulesFromMap where the same issue was first
        # found and fixed this session). This function feeds the CSV
        # inventory/display summary (rules_inventory.csv Conditions column)
        # -- not yet confirmed as part of the live SendersDomain regression,
        # but converted proactively since it reads the exact same COM
        # property via the exact same unreliable access pattern. Debug
        # logging added so the next live run can confirm whether this read
        # site returns different/correct data via InvokeMember vs. what the
        # old dot-notation read was producing.
        $senderAddr = $null
        try {
            $senderAddr = $conditions.GetType().InvokeMember(
                'SenderAddress',
                [System.Reflection.BindingFlags]::GetProperty,
                $null, $conditions, $null
            )

            $senderAddrEnabled = $false
            $senderAddrAddressRaw = $null
            if ($null -ne $senderAddr) {
                $senderAddrEnabled = $senderAddr.GetType().InvokeMember(
                    'Enabled',
                    [System.Reflection.BindingFlags]::GetProperty,
                    $null, $senderAddr, $null
                )
                $senderAddrAddressRaw = $senderAddr.GetType().InvokeMember(
                    'Address',
                    [System.Reflection.BindingFlags]::GetProperty,
                    $null, $senderAddr, $null
                )
            }

            # DEBUG (2026-07-07, Administrator debug session): compare InvokeMember read
            # results against what plain dot-notation was previously returning.
            try {
                $debugSaDump = if ($null -eq $senderAddrAddressRaw) { '<null>' } else { (@($senderAddrAddressRaw) -join ' | ') }
                $debugSaType = if ($null -eq $senderAddrAddressRaw) { '<null>' } else { $senderAddrAddressRaw.GetType().FullName }
                Write-OMMigrateLog -Message "DEBUG Get-RuleConditionsSummary SenderAddress (InvokeMember): Enabled=$senderAddrEnabled AddressType=$debugSaType AddressRaw=[$debugSaDump]" -Level DEBUG
            } catch { }

            if ($senderAddrEnabled) {
                $addrs = @($senderAddrAddressRaw) -join ', '
                $parts.Add("Sender address: $addrs")
            }
        }
        catch { }
        finally { if ($senderAddr) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($senderAddr) } catch { } } }

        # Body contains condition
        $bodyCond = $null
        try {
            $bodyCond = $conditions.Body
            if ($bodyCond.Enabled) {
                $bodyText = @($bodyCond.Text) -join ', '
                $parts.Add("Body contains: $bodyText")
            }
        }
        catch { }
        finally { if ($bodyCond) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($bodyCond) } catch { } } }

        # Account condition
        $acctCond = $null
        try {
            $acctCond = $conditions.Account
            if ($acctCond.Enabled) {
                # Break nested dot chain -- Gemini: each dot creates a pinned COM ref
                $rawAccount = $acctCond.Account
                $parts.Add("Account: $($rawAccount.SmtpAddress)")
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rawAccount) } catch { }
            }
        }
        catch { }
        finally { if ($acctCond) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($acctCond) } catch { } } }
    }
    catch {
        $parts.Add('[Could not read conditions]')
    }
    finally {
        # Release the conditions object itself
        if ($conditions) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($conditions) } catch { } }
    }

    return $(if ($parts.Count -gt 0) { $parts -join ' AND ' } else { '[No conditions]' })
}


function Get-RuleSenderAddressWords {
    <#
    .SYNOPSIS
        Returns the rule's SenderAddress condition values as a single
        space-separated string, or an empty string if not present.

    .DESCRIPTION
        NEW (2026-07-23, Administrator direction): Script 00 previously only ever
        defaulted SendersDomain to the last segment of TargetFolderPath
        (a best guess), since manually-added Outlook rules (via the UI
        Rules Manager) commonly have a real SenderAddress condition that
        is a far more reliable source for SendersDomain. This helper
        reads that condition directly and independently of
        Get-RuleConditionsSummary's human-readable Conditions string, so
        the existing, already-tested Conditions summary output is left
        completely untouched.

        Uses the same InvokeMember reflection access pattern as the
        SenderAddress block in Get-RuleConditionsSummary (module standing
        COM rule -- applies to all Outlook COM rule property access).

    .PARAMETER Rule
        An Outlook Rule COM object.

    .OUTPUTS
        [string] -- space-separated SenderAddress words, or '' if the
        condition is absent/disabled/unreadable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Rule
    )

    $result = ''

    try {
        $conditions = $Rule.Conditions

        $senderAddr = $null
        try {
            $senderAddr = $conditions.GetType().InvokeMember(
                'SenderAddress',
                [System.Reflection.BindingFlags]::GetProperty,
                $null, $conditions, $null
            )

            $senderAddrEnabled = $false
            $senderAddrAddressRaw = $null
            if ($null -ne $senderAddr) {
                $senderAddrEnabled = $senderAddr.GetType().InvokeMember(
                    'Enabled',
                    [System.Reflection.BindingFlags]::GetProperty,
                    $null, $senderAddr, $null
                )
                $senderAddrAddressRaw = $senderAddr.GetType().InvokeMember(
                    'Address',
                    [System.Reflection.BindingFlags]::GetProperty,
                    $null, $senderAddr, $null
                )
            }

            if ($senderAddrEnabled -and $null -ne $senderAddrAddressRaw) {
                $result = (@($senderAddrAddressRaw) -join ' ').Trim()
            }
        }
        catch { }
        finally { if ($senderAddr) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($senderAddr) } catch { } } }
    }
    catch { }
    finally {
        if ($conditions) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($conditions) } catch { } }
    }

    return $result
}


function Get-RuleActionsSummary {
    <#
    .SYNOPSIS
        Returns a human-readable summary of a rule's actions and
        extracts any folder target information.

    .DESCRIPTION
        Iterates the rule's Actions collection and builds a summary
        string. Also extracts the target folder path and EntryID for
        MoveToFolder and CopyToFolder actions -- these are the rules
        that Script 03 must update when folder locations change.

        Used internally by Get-OutlookRules.

    .PARAMETER Rule
        An Outlook Rule COM object.

    .OUTPUTS
        PSCustomObject with:
            Summary           [string]  Human-readable actions summary
            TargetFolderPath  [string]  Target folder full path (if any)
            TargetFolderEntryID [string] Target folder EntryID (if any)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Rule
    )

    $parts               = [System.Collections.Generic.List[string]]::new()
    $targetFolderPath    = ''
    $targetFolderEntryID = ''
    $targetStoreName     = ''

    try {
        $actions = $Rule.Actions

        # MoveToFolder action
        $moveAction = $null
        try {
            $moveAction = $actions.MoveToFolder
            if ($moveAction.Enabled) {
                $folder = $moveAction.Folder
                Register-COMObject -ComObject $folder
                $folderPath          = Get-FolderFullPath -Folder $folder
                $folderPath          = Remove-StorePrefix -FolderPath $folderPath
                $targetFolderPath    = $folderPath
                $targetFolderEntryID = $folder.EntryID
                # Resolve store name -- walk up to root store display name
                try {
                    $storeRoot = $folder.Store.GetRootFolder()
                    $targetStoreName = $storeRoot.Name
                    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($storeRoot) } catch { }
                }
                catch { }
                $parts.Add("Move to: $folderPath")
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) } catch { }
            }
        }
        catch { }
        finally { if ($moveAction) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($moveAction) } catch { } } }

        # CopyToFolder action
        $copyAction = $null
        try {
            $copyAction = $actions.CopyToFolder
            if ($copyAction.Enabled) {
                $folder = $copyAction.Folder
                Register-COMObject -ComObject $folder
                $folderPath = Get-FolderFullPath -Folder $folder
                $folderPath = Remove-StorePrefix -FolderPath $folderPath
                if (-not $targetFolderPath) {
                    $targetFolderPath    = $folderPath
                    $targetFolderEntryID = $folder.EntryID
                    # Resolve store name if not already set by MoveToFolder
                    try {
                        $storeRoot = $folder.Store.GetRootFolder()
                        $targetStoreName = $storeRoot.Name
                        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($storeRoot) } catch { }
                    }
                    catch { }
                }
                $parts.Add("Copy to: $folderPath")
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) } catch { }
            }
        }
        catch { }
        finally { if ($copyAction) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($copyAction) } catch { } } }

        # Delete action
        $deleteAction = $null
        try {
            $deleteAction = $actions.Delete
            if ($deleteAction.Enabled) { $parts.Add('Delete') }
        }
        catch { }
        finally { if ($deleteAction) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($deleteAction) } catch { } } }

        # Mark as read action
        $markReadAction = $null
        try {
            $markReadAction = $actions.MarkAsRead
            if ($markReadAction.Enabled) { $parts.Add('Mark as Read') }
        }
        catch { }
        finally { if ($markReadAction) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($markReadAction) } catch { } } }

        # Category assignment
        $categoryAction = $null
        try {
            $categoryAction = $actions.AssignToCategory
            if ($categoryAction.Enabled) {
                $cats = @($categoryAction.Categories) -join ', '
                $parts.Add("Assign Category: $cats")
            }
        }
        catch { }
        finally { if ($categoryAction) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($categoryAction) } catch { } } }

        # Forward to
        $fwdAction = $null
        try {
            $fwdAction = $actions.Forward
            if ($fwdAction.Enabled) { $parts.Add('Forward') }
        }
        catch { }
        finally { if ($fwdAction) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($fwdAction) } catch { } } }
    }
    catch {
        $parts.Add('[Could not read actions]')
    }
    finally {
        # Release the actions object itself
        if ($actions) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($actions) } catch { } }
    }

    return [PSCustomObject]@{
        Summary              = $(if ($parts.Count -gt 0) { $parts -join '; ' } else { '[No actions]' })
        TargetFolderPath     = $targetFolderPath
        TargetFolderEntryID  = $targetFolderEntryID
        TargetStoreName      = $targetStoreName
    }
}


function Remove-StorePrefix {
    <#
    .SYNOPSIS
        Strips the leading store name segment from a fully-qualified
        Outlook folder path returned by Get-FolderFullPath.

    .DESCRIPTION
        Get-FolderFullPath walks the COM parent chain to the root,
        producing paths like 'StoreName\account@domain.com\Inbox\Folder'.
        Script 03 and FixRulePaths work with store-relative paths
        ('account@domain.com\Inbox\Folder') -- the store name prefix
        is implicit from the migration context and must not be included.

        This function builds a case-insensitive set of all mounted store
        display names from the active COM session (scoped to the selected
        profile) on first call, caches it for the lifetime of the COM
        session, and strips the leading segment if it matches any known
        store name.

        Universal -- no store names or account addresses are hardcoded.
        Works for any user's Outlook profile and store configuration.

    .PARAMETER FolderPath
        Fully-qualified folder path from Get-FolderFullPath.

    .OUTPUTS
        [string] -- Path with leading store name stripped if matched,
                    or the original path unchanged if no match found.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return $FolderPath }

    # Build store name cache on first call -- scoped to active COM session
    if ($null -eq $Script:MountedStoreNames) {
        $Script:MountedStoreNames = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )
        try {
            $ns = Get-OutlookNamespace
            if ($ns) {
                $stores = $ns.Stores
                for ($si = 1; $si -le $stores.Count; $si++) {
                    try {
                        $storeName = $stores.Item($si).DisplayName
                        if (-not [string]::IsNullOrWhiteSpace($storeName)) {
                            [void]$Script:MountedStoreNames.Add($storeName)
                        }
                    }
                    catch { }
                }
            }
        }
        catch {
            Write-OMMigrateLog -Message "Remove-StorePrefix: could not build store name cache -- paths returned as-is: $_" `
                               -Level DEBUG
        }
    }

    # Strip leading segment if it matches a known store display name.
    # Only strip if the remaining path still contains an @ -- this ensures
    # account email addresses that are also store names (e.g. a mounted PST
    # named user@example.com) are never incorrectly stripped from
    # paths where they are the account prefix, not the store wrapper.
    $segments = $FolderPath -split '\\' 
    if ($segments.Count -gt 1 -and $Script:MountedStoreNames.Contains($segments[0])) {
        $remainder = $segments[1..($segments.Count - 1)] -join '\'
        if ($remainder -like '*@*') {
            return $remainder
        }
    }

    return $FolderPath
}


function Get-FolderFullPath {
    <#
    .SYNOPSIS
        Returns the full path of an Outlook folder by walking up
        its parent chain.

    .DESCRIPTION
        Builds a backslash-delimited path string from a folder object
        by traversing its Parent property until the root store is reached.

        Used when documenting rule target folders in the discovery report.

    .PARAMETER Folder
        An Outlook MAPIFolder COM object.

    .OUTPUTS
        [string] -- Full folder path (e.g. 'user@domain.com\Inbox\Vendors')
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Folder
    )

    $parts = [System.Collections.Generic.List[string]]::new()

    try {
        $current = $Folder
        $maxDepth = 20   # Safety limit to prevent infinite loops on corrupt profiles
        $depth = 0

        while ($current -and $depth -lt $maxDepth) {
            $parts.Insert(0, $current.Name)
            $parent = $null
            try {
                $parent = $current.Parent
                if ($parent -is [string]) { break }   # Reached root (string path)
                # Stop when parent is no longer a MAPIFolder (olFolder = 2).
                # Walking past the store root into NameSpace or Application objects
                # produces empty-string or profile-name segments that corrupt the path.
                try {
                    if ($parent.Class -ne 2) { break }
                }
                catch { break }   # .Class not available -- not a folder object
                $current = $parent
            }
            catch { break }
            $depth++
        }
    }
    catch {
        # Parent walk failed -- try to build storeName\folderName before giving up.
        # A bare folder name with no store prefix is unusable by Script 03.
        try {
            $storeRootName = $Folder.Store.GetRootFolder().Name
            if ($storeRootName) {
                return "$storeRootName\$($Folder.Name)"
            }
        }
        catch { }
        # Store lookup also failed -- return empty string so the caller
        # knows there is no valid path rather than a misleading bare name.
        return ''
    }

    return $parts -join '\'
}


# ============================================================
#  HELPER: Invoke-NormalizeRulesExecutionOrder
#  Reads rules_inventory.csv, sorts ALL rules globally by RuleName
#  alphabetically across all accounts, assigns a single sequential
#  ExecutionOrder 1-to-N across all rules regardless of account.
#
#  This global numbering ensures that when Outlook displays all
#  secondary store rules alongside ameritech rules in the Rules &
#  Alerts UI, they appear in a single alphabetical list rather than
#  grouped by account with conflicting per-account order numbers.
#
#  CSV structure is preserved: separator rows (blank RuleName) and
#  account groupings remain intact -- only ExecutionOrder values change.
#
#  Called after every CSV write in Scripts 00 and 03 so that:
#    - The CSV always reflects correct global sequential order
#    - Admins adding rules manually never need to set ExecutionOrder
#    - Script 03 can trust ExecutionOrder values as authoritative
#    - Invoke-PurgeAndRecreateRules reads order directly from CSV
# ============================================================

function Invoke-NormalizeRulesExecutionOrder {
    <#
    .SYNOPSIS
        Normalizes ExecutionOrder in rules_inventory.csv to a single
        global sequential 1-to-N across ALL rules sorted alphabetically
        by RuleName regardless of RuleStoreName group.

    .PARAMETER CsvPath
        Full path to rules_inventory.csv.

    .OUTPUTS
        [int] Number of rows updated, or -1 on failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvPath
    )

    if (-not (Test-Path $CsvPath)) {
        Write-OMMigrateLog -Message "Invoke-NormalizeRulesExecutionOrder: CSV not found: $CsvPath" -Level WARN
        return -1
    }

    try {
        $allRows  = Import-Csv -Path $CsvPath -Encoding UTF8
        $dataRows = @($allRows | Where-Object { $_.RuleName -ne '' })

        # Sort fix (2026-07-02, Administrator): previously sorted by RuleStoreName
        # then TargetFolderPath, which groups rules by folder TREE structure
        # (parent folder immediately followed by its own subfolders) rather
        # than true flat alphabetical order by rule name -- confirmed live:
        # rules targeting nested subfolders (e.g. "ExampleCo
        # Accounting\aws") clustered under their parent instead of sorting
        # by their own label. This docstring always described the intended
        # behavior as "sorted alphabetically by RuleName" -- this now
        # matches that description.
        #
        # Sort key extraction: for standardized consolidated rule names
        # ("Rule: [account] Label (Part N)"), pull out just "Label" so it
        # sorts on the same leaf-folder-name basis Invoke-BuildRulesFromMap
        # uses to build $baseFolderName/$chunkName at creation time. Any
        # RuleName that does NOT match the standardized pattern -- including
        # every manually-added, end-user-typed rule name -- falls through
        # to sorting on the raw RuleName as-is, so manual rules interleave
        # naturally at their own alphabetical position alongside
        # standardized ones instead of being segregated by name format.
        $globalSorted = @($dataRows | Sort-Object {
            $sortRuleName = [string]$_.RuleName
            $sortLabel    = $sortRuleName
            if ($sortRuleName -match '^Rule: \[[^\]]*\] (.+?) \(Part \d+\)$') {
                $sortLabel = $Matches[1]
            }
            $sortLabel.Trim().ToLower()
        })
        $updatedCount = 0
        for ($i = 0; $i -lt $globalSorted.Count; $i++) {
            $newOrder = $i + 1
            if ([string]$globalSorted[$i].ExecutionOrder -ne [string]$newOrder) {
                $globalSorted[$i].ExecutionOrder = $newOrder
                $updatedCount++
            }
        }

        # Rebuild output preserving account group structure and separators.
        # Data rows are sorted by RuleStoreName then by their new global
        # ExecutionOrder so each account group stays together in the CSV
        # while ExecutionOrder values reflect the global alphabetical position.
        $sortedData = $dataRows | Sort-Object RuleStoreName, { [int]$_.ExecutionOrder }
        $output     = [System.Collections.Generic.List[object]]::new()
        $lastStore  = ''
        foreach ($row in $sortedData) {
            if ($lastStore -ne '' -and $row.RuleStoreName -ne $lastStore) {
                # Insert separator row between account groups
                $sep = '' | Select-Object $allRows[0].PSObject.Properties.Name
                $output.Add($sep)
            }
            $output.Add($row)
            $lastStore = $row.RuleStoreName
        }

        $output | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
        Write-OMMigrateLog -Message "Invoke-NormalizeRulesExecutionOrder: $updatedCount row(s) renumbered globally across $($dataRows.Count) rules in $CsvPath" -Level INFO
        return $updatedCount
    }
    catch {
        Write-OMMigrateLog -Message "Invoke-NormalizeRulesExecutionOrder: Failed: $_" -Level WARN
        return -1
    }
}

# ADDED (Administrator direction, 2026-08-18): shared helper so every Script 03 write
# site that saves rules_inventory.csv can re-insert the blank separator rows
# between RuleStoreName groups, matching Export-RulesToCSV's own behavior
# (Script 00) instead of writing a flat, unseparated row list. Script 03's
# several partial write-backs (LastDeployedRun/LastTargetRun stamping,
# NeedsFolderUpdate updates, etc.) each Import-Csv the file with separator
# rows already filtered OUT (by design -- see each call site's own "Skip
# blank separator rows" comment), update specific rows, then Export-Csv the
# result directly with no separators re-added, leaving the operator's Excel
# copy without the visual account-group breaks Script 00 originally wrote.
# This helper re-sorts by RuleStoreName/TargetFolderPath (same sort
# Export-RulesToCSV itself uses) and inserts one blank row every time
# RuleStoreName changes, so every write site can call this immediately
# before Export-Csv instead of writing $rows directly.
function Add-RulesCsvSeparatorRows {
    <#
    .SYNOPSIS
        Sorts a rules_inventory.csv data-row set by RuleStoreName/
        TargetFolderPath and inserts a blank separator row between each
        RuleStoreName group, mirroring Export-RulesToCSV's own layout.

    .PARAMETER Rows
        Data rows only -- must already exclude any blank separator rows
        (every call site's existing Import-Csv filter already does this
        before rows reach this function).

    .OUTPUTS
        [object[]] -- the sorted rows with blank separator rows inserted,
        ready to pipe directly into Export-Csv.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [object[]] $Rows
    )

    if (-not $Rows -or $Rows.Count -eq 0) { return $Rows }

    $sortedForSeparator = @($Rows | Sort-Object RuleStoreName, TargetFolderPath)
    $withSeparators     = [System.Collections.Generic.List[object]]::new()
    $lastStoreForSep    = $null
    $blankSepTemplate   = '' | Select-Object $Rows[0].PSObject.Properties.Name

    foreach ($sepRow in $sortedForSeparator) {
        if ($null -ne $lastStoreForSep -and $sepRow.RuleStoreName -ne $lastStoreForSep) {
            $withSeparators.Add($blankSepTemplate)
        }
        $withSeparators.Add($sepRow)
        $lastStoreForSep = $sepRow.RuleStoreName
    }

    return $withSeparators
}

function Export-RulesToCSV {
    <#
    .SYNOPSIS
        Exports the Outlook rules inventory to a CSV file for the
        discovery report and Script 03 reference.

    .DESCRIPTION
        Writes rules_inventory.csv to the Config directory.
        This file serves two purposes:
            1. Operator review -- see all rules before migration
            2. Script 03 input -- rules with NeedsFolderUpdate=True
               will have their TargetFolderPath updated to the new
               post-migration folder location

        SECURITY: Rule conditions may reference email addresses.
        This file stays entirely on the local machine.

    .PARAMETER Rules
        Array of rule objects from Get-OutlookRules.

    .PARAMETER OutputPath
        Full path for the output CSV file.

    .OUTPUTS
        [string] -- Path to the written CSV file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[PSCustomObject]]$Rules,

        [Parameter(Mandatory = $false)]
        [string]$OutputPath = '',

        # Added 2026-07-10, Administrator direction. Optional active Outlook MAPI
        # namespace COM object, used ONLY to resolve the real display name
        # of the default/fallback Archive PST (matched by its known,
        # hardcoded PSTPath -- same constant Script 01/03/Install.ps1 already
        # use -- via EXACT PSTPath equality, never a display-name substring
        # match) so an unmapped row's blank TargetStoreName can be filled
        # with the actual archive name instead of the account's own
        # RuleStoreName. $null (default) preserves the exact prior fallback
        # behavior (TargetStoreName = RuleStoreName) for any caller that
        # doesn't pass this -- fully backward compatible, no existing caller
        # is required to change. EXACT PSTPath match only -- deliberately
        # NOT a display-name-substring match, since Administrator's own profile has
        # 24 'ArchiveBuild -- <email>' backup stores whose names contain
        # each account's own identifier and would be unsafe to match against
        # loosely.
        [Parameter(Mandatory = $false)]
        [object]$Namespace = $null
    )

    if ([string]::IsNullOrEmpty($OutputPath)) {
        $OutputPath = Get-OMMigrateCsvPath -BaseName 'rules_inventory.csv'
    }

    Write-OMMigrateLog -Message "Exporting rules inventory: $OutputPath" -Level INFO

    # Added 2026-07-10, Administrator direction. Resolve the default/fallback Archive
    # PST's live display name ONCE per call, via exact PSTPath match against
    # the same hardcoded default path constant used elsewhere in this
    # project -- not a new hardcode, reading an existing known value. $null
    # if that specific PST is not currently attached (Namespace not
    # supplied, or no live store's FilePath matches) -- the row-filling loop
    # below falls back to the exact prior RuleStoreName behavior in that case.
    $exportDefaultArchiveDisplayName = $null
    if ($Namespace) {
        try {
            $exportDefaultArchivePSTPath = Join-Path $Global:OMMigrate.BackupPath 'OMMigrate_Archive.pst'
            $exportLiveStores = $Namespace.Stores
            for ($edi = 1; $edi -le $exportLiveStores.Count; $edi++) {
                $exportLookupStore = $exportLiveStores.Item($edi)
                $exportLookupStoreFilePath = $null
                try { $exportLookupStoreFilePath = $exportLookupStore.FilePath } catch { }
                if ($exportLookupStoreFilePath -and $exportLookupStoreFilePath -eq $exportDefaultArchivePSTPath) {
                    $exportDefaultArchiveDisplayName = $exportLookupStore.DisplayName
                    break
                }
            }
        }
        catch { }
    }

    # FIXED 2026-07-10, Administrator direction (live-tested on "TestProfile", bug found:
    # explicit picker mappings never took effect on a rerun because the old
    # merge-preserve rule below -- added 2026-07-06, before this picker
    # existed -- protects ANY non-blank TargetStoreName unconditionally,
    # including a stale value from a prior run/prior mapping). Moved the
    # $ruleStoreToTarget lookup build earlier (was previously built further
    # below, right before the blank-row-fill loop) so the merge block can
    # also consult it: a row is now only preserved as a manual admin edit
    # when its existing CSV value does NOT match what the CURRENT picker
    # mapping would produce for that RuleStoreName. This lets the operator
    # still hand-edit an individual rule's TargetStoreName in the CSV and
    # have that override the picker for that one specific rule -- but a
    # genuine account-level remapping via the picker (the more recent,
    # broader operator decision) now correctly takes effect on the next
    # export instead of being silently blocked forever by an old value.
    $archiveStoreMappingsForMerge = @()
    try {
        if ($Global:OMMigrate.Settings.PSObject.Properties['RulesEngine'] -and
            $Global:OMMigrate.Settings.RulesEngine.PSObject.Properties['ArchiveStoreMappings']) {
            $archiveStoreMappingsForMerge = @($Global:OMMigrate.Settings.RulesEngine.ArchiveStoreMappings)
        }
    }
    catch { }

    $ruleStoreToTarget = @{}
    foreach ($mappingForMerge in $archiveStoreMappingsForMerge) {
        if (-not $mappingForMerge.PSObject.Properties['TargetStoreName'] -or
            [string]::IsNullOrWhiteSpace($mappingForMerge.TargetStoreName)) { continue }
        foreach ($mappedAccountForMerge in @($mappingForMerge.RuleStoreNames)) {
            if (-not [string]::IsNullOrWhiteSpace($mappedAccountForMerge)) {
                $ruleStoreToTarget[$mappedAccountForMerge] = $mappingForMerge.TargetStoreName
            }
        }
    }

    if (-not $Global:OMMigrate.WhatIf) {
        try {
            # -- Merge with existing CSV if it exists to preserve operator edits --
            # Keyed by RuleStoreName+RuleName+ExecutionOrder -- the unique identifier for a rule.
            # ExecutionOrder is included to handle duplicate rule names on the same store
            # (e.g. rules recreated with the same name after a prior rule failed to fire).
            # Preserves NeedsFolderUpdate edits the operator may have set to False.
            $mergeStats = @{ Retained = 0; Added = 0 }

            if (Test-Path $OutputPath) {
                Write-OMMigrateLog -Message "Existing rules_inventory.csv found -- merging to preserve operator edits." `
                                   -Level INFO

                $existingRows = @{}
                # ADDED (Administrator direction -- SendersDomain/preservation-block clobber fix,
                # live-tested 2026-07-13 on "TestProfile"/testaccount@example.com): fallback
                # dictionary keyed by RuleStoreName+RuleName, consulted only when the
                # primary RuleStoreName+TargetFolderPath key (below) fails to find a
                # match. Root cause traced live: a rule whose folder target COM
                # resolution fails gets a DIFFERENT best-guess TargetFolderPath on
                # every Script 00 rerun (confirmed live -- the guess changed between
                # runs for the same rule). Since the primary merge key is built from
                # TargetFolderPath, an unstable guess means the key itself changes
                # every run, so $existingRows never finds the prior row -- silently
                # defeating every preservation block below (SendersDomain, Conditions,
                # Actions, LastDeployedRun, LastTargetRun) for exactly this class of
                # rule, even though the row objectively already existed. RuleName
                # (the raw, pre-consolidation name) does not change just because a
                # folder failed to resolve, so it survives this specific instability
                # -- unlike TargetFolderPath here, and unlike RuleName itself across a
                # CONSOLIDATION rename (the reason the primary key was moved off
                # RuleName in the first place, per the comment below). Trying the
                # primary key first and RuleName only as a fallback preserves both
                # fixes at once: consolidation renames still match correctly via
                # TargetFolderPath, and COM-inaccessible-folder rules now also survive
                # a rerun via this fallback instead of losing their preserved data.
                $existingRowsByName = @{}
                $existingRowsByLeaf = @{}
                try {
                    Import-Csv -Path $OutputPath -Encoding UTF8 |
                        ForEach-Object {
                            # Skip blank separator rows -- they are regenerated fresh
                            # on every write and must never be used as merge keys.
                            # A blank RuleStoreName or blank RuleName means separator row.
                            if ($_.RuleStoreName -and $_.RuleName -and
                                -not [string]::IsNullOrWhiteSpace($_.RuleStoreName) -and
                                -not [string]::IsNullOrWhiteSpace($_.RuleName)) {
                                # Populate the RuleName-based fallback dictionary alongside
                                # the primary one below -- same first-wins-on-collision
                                # semantics are not needed here since this is only ever
                                # consulted when the primary key already missed.
                                $existingKeyByName = "$($_.RuleStoreName)|$($_.RuleName)"
                                if (-not $existingRowsByName.ContainsKey($existingKeyByName)) {
                                    $existingRowsByName[$existingKeyByName] = $_
                                }
                                # SECOND fallback dictionary, keyed by RuleStoreName + the
                                # LEAF segment of TargetFolderPath (e.g. "testaccount@example.com"
                                # from either the raw path "testaccount@example.com" or from
                                # the standardized label inside "Rule: [account] Label (Part N)").
                                # Live-tested gap found: the RuleName fallback above fails when a
                                # row has BOTH an unstable TargetFolderPath guess AND a
                                # consolidation rename between two runs -- neither the raw old
                                # name nor the new standardized name match each other, so the
                                # RuleName-based fallback also misses, producing a duplicate row
                                # (confirmed live: 'Ommtest' vs 'Rule: [testaccount@example.com]
                                # testaccount@example.com (Part 1)'). The leaf folder segment is
                                # the one thing that stays the same across a rename, since the
                                # standardized name's Label portion IS the leaf folder name the
                                # rule targets -- same value whether the row is still raw or has
                                # already been consolidated.
                                $existingLeafSource = if ($_.PSObject.Properties['TargetFolderPath']) { [string]$_.TargetFolderPath } else { '' }
                                $existingLeaf = ($existingLeafSource -split '\\')[-1]
                                # FIXED (Administrator direction, this same fix -- consistency pass):
                                # same standardized-name regex fallback as the lookup-time
                                # leaf derivation further below must also apply HERE at
                                # build time, or a row whose TargetFolderPath is blank (not
                                # just a bad guess, genuinely empty) would never be inserted
                                # into $existingRowsByLeaf at all, while a fresh-scan row in
                                # the same blank-path state WOULD derive a leaf via this same
                                # regex fallback at lookup time -- searching for a key that
                                # was never actually built, silently failing to match.
                                if ([string]::IsNullOrWhiteSpace($existingLeaf) -and $_.RuleName -match '^Rule: \[[^\]]*\] (.+?) \(Part \d+\)$') {
                                    $existingLeaf = $Matches[1]
                                }
                                if (-not [string]::IsNullOrWhiteSpace($existingLeaf)) {
                                    $existingKeyByLeaf = "$($_.RuleStoreName)|$($existingLeaf.Trim().ToLower())"
                                    if (-not $existingRowsByLeaf.ContainsKey($existingKeyByLeaf)) {
                                        $existingRowsByLeaf[$existingKeyByLeaf] = $_
                                    }
                                }
                                # Key by RuleStoreName+RuleName -- changed 2026-06-27.
                                # RuleName-only was safe when only the default store was
                                # ever scanned (Get-OutlookRules pre-2026-06-27). Now that
                                # every account is scanned, two different accounts can
                                # legitimately have identically-named rules (e.g. "Move to
                                # Archive"); a RuleName-only key would silently collapse
                                # them into one CSV row and lose one account's rule data.
                                # ExecutionOrder is still not part of the key -- it is not
                                # stable across scans when Outlook has duplicate rules.
                                #
                                # FIXED (Administrator direction, this same fix -- root cause of the
                                # ecobee LastTargetRun-never-stamps bug traced live 2026-07-12):
                                # RuleName is NOT stable across runs -- consolidation renames a
                                # rule from its raw UI name to the standardized
                                # "Rule: [account] Folder (Part N)" form, so a RuleName-based
                                # key treats the renamed rule as an entirely new row on the
                                # very next scan, orphaning the row that had LastDeployedRun/
                                # LastTargetRun already stamped and creating a fresh blank one
                                # -- an infinite recreate-and-never-complete loop. Key on
                                # RuleStoreName+TargetFolderPath instead -- the same stable
                                # identity already used throughout Invoke-DeployConsolidatedRules,
                                # Invoke-BuildRulesFromMap, and Script 03's Phase 3 loop. Two
                                # different accounts can still share the same TargetFolderPath
                                # segment text in principle, but RuleStoreName is still the
                                # first key component, so that collision this key was originally
                                # protecting against is still fully prevented.
                                $existingKeyPath = if ($_.PSObject.Properties['TargetFolderPath']) { [string]$_.TargetFolderPath } else { '' }
                                # FIXED (Administrator direction, 2026-07-20, Part 1/Part 2 rules_inventory.csv
                                # mislabeling bug): RuleName added into the primary key alongside
                                # TargetFolderPath. Root cause: when a folder's sender list exceeds 5
                                # addresses, consolidation splits it into multiple rules (e.g. "...
                                # (Part 1)" and "... (Part 2)") that all target the SAME TargetFolderPath.
                                # A TargetFolderPath-only key collapsed all of them onto one dictionary
                                # slot, so only one Part-N row survived the build step -- and on the
                                # lookup side (below), EVERY live Part-N rule matched that single slot,
                                # so whichever rule was processed last in $Rules ended up overwriting the
                                # merged row's identity, visually mislabeling Part 1's row as Part 2 in
                                # the CSV while the actual Part 2 rule (and both live Outlook rules) were
                                # unaffected -- confirmed live by Administrator: Outlook's own Rules and
                                # Alerts UI correctly showed both Part 1 and Part 2, only the CSV was
                                # wrong. Including RuleName restores a unique key per Part-N rule while
                                # keeping TargetFolderPath as the primary component (still handles the
                                # single-rule-per-folder consolidation-rename case exactly as before).
                                $existingKey = "$($_.RuleStoreName)|$existingKeyPath|$($_.RuleName)"
                                if (-not $existingRows.ContainsKey($existingKey)) {
                                    $existingRows[$existingKey] = $_
                                }
                                else {
                                    # Already have a row for this name -- keep the one with valid path
                                    $current = $existingRows[$existingKey]
                                    $currentHasPath = ($current.TargetFolderPath -and
                                                       $current.TargetFolderPath -like '*@*' -and
                                                       $current.TargetFolderPath -like '*\*')
                                    $newHasPath     = ($_.TargetFolderPath -and
                                                       $_.TargetFolderPath -like '*@*' -and
                                                       $_.TargetFolderPath -like '*\*')
                                    $newHasActions  = ($_.Actions -and $_.Actions -ne '[No actions]')
                                    if ((-not $currentHasPath -and $newHasPath) -or
                                        (-not $currentHasPath -and $newHasActions)) {
                                        $existingRows[$existingKey] = $_
                                    }
                                }
                            }
                        }
                }
                catch {
                    Write-OMMigrateLog -Message "Could not read existing rules_inventory.csv for merge -- will overwrite: $_" `
                                       -Level WARN
                }

                if ($existingRows.Count -gt 0) {
                    $mergedRules = [System.Collections.Generic.List[PSCustomObject]]::new()

                    foreach ($newRule in $Rules) {
                        # FIXED (Administrator direction, this same fix): lookup key must match the
                        # RuleStoreName+TargetFolderPath identity used to build $existingRows
                        # above, or a renamed rule would never find its prior-run row.
                        $newKeyPath = if ($newRule.PSObject.Properties['TargetFolderPath']) { [string]$newRule.TargetFolderPath } else { '' }
                        # FIXED (Administrator direction, 2026-07-20, Part 1/Part 2 rules_inventory.csv
                        # mislabeling bug): lookup key must match the build-side key above exactly
                        # (RuleStoreName+TargetFolderPath+RuleName), or a Part-N rule would never
                        # find its own dedicated row and would incorrectly fall through to the
                        # Name/Leaf fallback dictionaries instead.
                        $key = "$($newRule.RuleStoreName)|$newKeyPath|$($newRule.RuleName)"
                        # ADDED (Administrator direction -- SendersDomain/preservation-block clobber
                        # fix, see $existingRowsByName build comment above for full root-
                        # cause explanation): when the primary TargetFolderPath-based key
                        # misses (e.g. this rule's folder target is COM-inaccessible and its
                        # best-guess path changed since the prior run), fall back to the
                        # RuleStoreName+RuleName key -- stable for exactly this case. Primary
                        # key is always tried first, so a genuine consolidation rename (the
                        # scenario the primary key exists to handle) is unaffected; this
                        # fallback only engages when the primary key already failed to match.
                        # ADDED (Administrator direction -- SendersDomain/preservation-block clobber
                        # fix, THIRD pass, live-tested 2026-07-13): $matchedExistingSource
                        # tracks WHICH dictionary actually matched, not just a key string --
                        # two different dictionaries could coincidentally produce the same
                        # key text, so the source must be tracked explicitly to read from
                        # (and later remove from) the correct one.
                        $matchedExistingKey    = $null
                        $matchedExistingSource = $null
                        if ($newRule.RuleName -and $existingRows.ContainsKey($key)) {
                            $matchedExistingKey    = $key
                            $matchedExistingSource = 'Path'
                        }
                        elseif ($newRule.RuleName -and $newRule.RuleStoreName) {
                            $fallbackKey = "$($newRule.RuleStoreName)|$($newRule.RuleName)"
                            if ($existingRowsByName.ContainsKey($fallbackKey)) {
                                $matchedExistingKey    = $fallbackKey
                                $matchedExistingSource = 'Name'
                            }
                        }
                        # THIRD fallback -- leaf folder segment (see $existingRowsByLeaf build
                        # comment above for full root-cause explanation). Tried last, only when
                        # BOTH the path-based and name-based keys have already missed -- covers
                        # the combined failure mode where a row's TargetFolderPath guess changed
                        # AND the rule was consolidation-renamed between two runs, so neither of
                        # the first two keys still matches, but the leaf folder name (present in
                        # both the raw path and the standardized name's Label portion) does.
                        if ($null -eq $matchedExistingKey -and $newRule.RuleStoreName) {
                            $newLeafSource = if ($newRule.PSObject.Properties['TargetFolderPath']) { [string]$newRule.TargetFolderPath } else { '' }
                            $newLeaf = ($newLeafSource -split '\\')[-1]
                            if ([string]::IsNullOrWhiteSpace($newLeaf) -and $newRule.RuleName -match '^Rule: \[[^\]]*\] (.+?) \(Part \d+\)$') {
                                $newLeaf = $Matches[1]
                            }
                            if (-not [string]::IsNullOrWhiteSpace($newLeaf)) {
                                $leafFallbackKey = "$($newRule.RuleStoreName)|$($newLeaf.Trim().ToLower())"
                                if ($existingRowsByLeaf.ContainsKey($leafFallbackKey)) {
                                    # FIXED (Administrator direction, 2026-07-2X, same Part 1/Part 2
                                    # collision as the altExisting block above, narrower exposure --
                                    # this path only runs when BOTH the Path and Name keys already
                                    # missed for this rule. Same guard: if the Leaf match is itself a
                                    # valid Part-N sibling for the same base folder, it is not a genuine
                                    # match for THIS rule and must not be used.
                                    $leafCandidate = $existingRowsByLeaf[$leafFallbackKey]
                                    $leafCandidateIsPartNSibling = $false
                                    if ($newRule.RuleName -match '^Rule: \[[^\]]*\] (.+?) \(Part \d+\)$') {
                                        $newRuleBaseLabelForLeaf = $Matches[1]
                                        if ($leafCandidate.RuleName -match '^Rule: \[[^\]]*\] (.+?) \(Part \d+\)$' -and
                                            $Matches[1] -eq $newRuleBaseLabelForLeaf -and
                                            $leafCandidate.RuleName -ne $newRule.RuleName) {
                                            $leafCandidateIsPartNSibling = $true
                                        }
                                    }
                                    if (-not $leafCandidateIsPartNSibling) {
                                        $matchedExistingKey    = $leafFallbackKey
                                        $matchedExistingSource = 'Leaf'
                                    }
                                }
                            }
                        }
                        if ($null -ne $matchedExistingKey) {
                            # Rule exists -- preserve valid TargetFolderPath from CSV
                            $existing    = switch ($matchedExistingSource) {
                                'Path' { $existingRows[$matchedExistingKey] }
                                'Name' { $existingRowsByName[$matchedExistingKey] }
                                'Leaf' { $existingRowsByLeaf[$matchedExistingKey] }
                            }
                            $mergedRule  = $newRule.PSObject.Copy()
                            # Inject LastDeployedRun if the copied object doesn't have it --
                            # happens when the source CSV predates this column. Add-Member
                            # ensures Export-Csv includes it in the header on this write.
                            if (-not $mergedRule.PSObject.Properties['LastDeployedRun']) {
                                Add-Member -InputObject $mergedRule -MemberType NoteProperty -Name 'LastDeployedRun' -Value '' -Force
                            }
                            # NEW (2026-07-07, Administrator direction): same injection for LastTargetRun
                            # -- ensures any pre-existing CSV that predates this column gets it
                            # added on the next Script 00 run, matching LastDeployedRun's pattern.
                            if (-not $mergedRule.PSObject.Properties['LastTargetRun']) {
                                Add-Member -InputObject $mergedRule -MemberType NoteProperty -Name 'LastTargetRun' -Value '' -Force
                            }
                            # Preserve TargetFolderPath if existing row has a valid store-relative
                            # path -- first segment must contain @ (an email address) to confirm it
                            # is a proper relative path and not a stale store-prefixed path from a
                            # prior run before Remove-StorePrefix was applied. A bare leaf name,
                            # blank, or store-prefixed path is overwritten with fresh COM data.
                            # FIXED (Administrator direction, live-tested 2026-07-13 on "TestProfile"/
                            # testaccount@example.com): the original check REQUIRED a
                            # backslash (-like '*\*'), which incorrectly rejected a
                            # legitimate single-segment path -- an account's rules targeting
                            # its OWN store root (RuleStoreName with no subfolder, e.g.
                            # "testaccount@example.com" alone, no "\Inbox\..." suffix) is
                            # a valid, real TargetFolderPath, not a stale/bad value. Confirmed
                            # live: an operator-corrected single-segment path was being
                            # discarded and overwritten with Script 00's COM-inaccessible-
                            # folder best-guess on every rerun because it failed this
                            # backslash requirement alone. Now accepts EITHER a genuine
                            # multi-segment path with an @ in its first segment (original
                            # case, unchanged) OR a single-segment value that IS itself an
                            # @ containing store name with no backslash at all (the new
                            # case) -- still rejects blank, and still rejects a bare leaf
                            # name with no @ anywhere (the original concern this check
                            # existed for).
                            $existingPathValid = (
                                $existing.PSObject.Properties['TargetFolderPath'] -and
                                $existing.TargetFolderPath -and
                                (
                                    ($existing.TargetFolderPath -like '*\*' -and
                                     ($existing.TargetFolderPath -split '\\')[0] -like '*@*') -or
                                    ($existing.TargetFolderPath -notlike '*\*' -and
                                     $existing.TargetFolderPath -like '*@*')
                                )
                            )
                            if ($existingPathValid) {
                                $mergedRule.TargetFolderPath = $existing.TargetFolderPath
                            }
                            # Preserve TargetFolderEntryID alongside the path
                            if ($existing.PSObject.Properties['TargetFolderEntryID'] -and
                                $existing.TargetFolderEntryID -and
                                -not [string]::IsNullOrWhiteSpace($existing.TargetFolderEntryID) -and
                                $existingPathValid) {
                                $mergedRule.TargetFolderEntryID = $existing.TargetFolderEntryID
                            }
                            # Preserve Notes if operator has annotated this rule
                            if ($existing.PSObject.Properties['Notes'] -and
                                $existing.Notes -and
                                -not [string]::IsNullOrWhiteSpace($existing.Notes)) {
                                $mergedRule.Notes = $existing.Notes
                            }
                            # Preserve TargetStoreName (Issue 2 fix, added 2026-07-06, Administrator;
                            # REVISED 2026-07-10, Administrator direction, live-tested bug fix on "TestProfile").
                            # Original intent: every merged row previously started as a full copy
                            # of the fresh COM scan, so TargetStoreName always reverted to whatever
                            # store the rule's live MoveToFolder action currently resolves to --
                            # overwriting any operator edit on every single rerun. That blanket
                            # non-blank preserve rule predates the Script 00 TargetStoreName
                            # picker and, once the picker existed, silently blocked it from ever
                            # taking effect on a rerun -- confirmed live: mapping ameritech to a
                            # second archive PST in the picker had zero effect because the row's
                            # existing (stale, pre-picker-selection) value was non-blank and so
                            # was preserved unconditionally, every time.
                            #
                            # Fixed to a two-tier rule instead of a blanket preserve:
                            #   1. If this row's RuleStoreName has an active entry in the CURRENT
                            #      ArchiveStoreMappings (picker selection), and the existing CSV
                            #      value differs from what that mapping would produce, this is a
                            #      genuine account-level remapping -- let it through (do NOT
                            #      preserve) so the picker's latest selection wins, exactly as the
                            #      picker's own UI text promises ("shown every run").
                            #   2. Otherwise (no active mapping for this account, OR the existing
                            #      value already matches what the mapping would produce anyway)
                            #      preserve the existing value -- this is either a deliberate
                            #      single-rule manual CSV edit the operator made that diverges from
                            #      the account-wide picker mapping (their more specific, more
                            #      recent intent for THIS rule), or there's nothing to change.
                            $existingMatchesCurrentMapping = $true
                            if ($ruleStoreToTarget.ContainsKey($mergedRule.RuleStoreName)) {
                                $existingMatchesCurrentMapping = (
                                    $existing.TargetStoreName -eq $ruleStoreToTarget[$mergedRule.RuleStoreName]
                                )
                            }
                            if ($existing.PSObject.Properties['TargetStoreName'] -and
                                $existing.TargetStoreName -and
                                -not [string]::IsNullOrWhiteSpace($existing.TargetStoreName) -and
                                $existingMatchesCurrentMapping) {
                                $mergedRule.TargetStoreName = $existing.TargetStoreName
                            }
                            elseif ($ruleStoreToTarget.ContainsKey($mergedRule.RuleStoreName)) {
                                # A genuine account-level remapping (existing value did not match
                                # the current picker mapping, per $existingMatchesCurrentMapping
                                # above) -- apply the new mapping directly here rather than
                                # leaving it to the later blank-only fill loop further below,
                                # since $mergedRule.TargetStoreName is very likely already
                                # non-blank at this point (fresh COM-scan value from the rule
                                # object built earlier this run), which would cause that later
                                # loop to also skip it, silently reproducing the exact bug this
                                # fix is for.
                                $mergedRule.TargetStoreName = $ruleStoreToTarget[$mergedRule.RuleStoreName]
                            }
                            # Preserve SendersDomain if operator has reviewed/corrected it.
                            # This is a user-editable column -- never overwrite on re-run.
                            # Script 00 sets a default (last segment of TargetFolderPath)
                            # on first write; after that the operator owns it.
                            # Validity check REPLACED (Issue 3 fix, 2026-07-06, Administrator): the
                            # old check rejected ANY value containing "@" outright, on the
                            # theory that "@" could only mean a stale full folder path had
                            # leaked into this column. That was true when this field could
                            # only ever be a bare domain word, but now that
                            # ConvertTo-NormalizedSenderDomains validates and accepts real
                            # full email addresses (which legitimately contain "@"), that
                            # blanket rejection would discard a valid operator-entered email
                            # on every rerun -- the second half of Issue 3. Reusing the same
                            # validation function here (rather than a separate ad hoc check)
                            # keeps both call sites in agreement about what counts as valid:
                            # a backslash still means a stale full folder path leaked in
                            # (never valid), but "@" alone no longer disqualifies a value.
                            #
                            # ADDED (Administrator direction -- SendersDomain/preservation-block clobber
                            # fix, FIFTH pass, live-tested 2026-07-13 on "TestProfile"/
                            # testaccount@example.com): when this row matched via the PRIMARY
                            # (Path) key, $existing may not be the row carrying the operator's
                            # correction -- confirmed live: a rule stuck in the COM-inaccessible-
                            # folder state gets an IDENTICAL best-guess TargetFolderPath on
                            # every rerun, so the primary key matches the row Script 00 itself
                            # wrote on a PRIOR run (still carrying that run's own best-guess
                            # SendersDomain), while a SEPARATE row under this rule's OLD
                            # (pre-rename) name -- the one the operator actually hand-corrected
                            # -- gets silently discarded by the orphan-skip leaf/name check
                            # instead, taking its correct SendersDomain down with it. When the
                            # match source is Path (not already Name/Leaf, which by definition
                            # only ever had ONE candidate row to begin with), also check the
                            # Name and Leaf dictionaries for another row for this same rule --
                            # if one exists and its SendersDomain passes validation while the
                            # primary-matched row's does not (or is blank), prefer the other
                            # row's value. This does not change behavior for the common case
                            # (no duplicate row exists) -- $altExisting stays $null and this
                            # block is a no-op.
                            $altExisting = $null
                            if ($matchedExistingSource -eq 'Path') {
                                if ($newRule.RuleName -and $newRule.RuleStoreName) {
                                    $altNameKey = "$($newRule.RuleStoreName)|$($newRule.RuleName)"
                                    if ($existingRowsByName.ContainsKey($altNameKey) -and
                                        $existingRowsByName[$altNameKey] -ne $existing) {
                                        $altExisting = $existingRowsByName[$altNameKey]
                                    }
                                }
                                if (-not $altExisting -and $newRule.RuleStoreName) {
                                    $altLeafSource = if ($newRule.PSObject.Properties['TargetFolderPath']) { [string]$newRule.TargetFolderPath } else { '' }
                                    $altLeaf = ($altLeafSource -split '\\')[-1]
                                    if ([string]::IsNullOrWhiteSpace($altLeaf) -and $newRule.RuleName -match '^Rule: \[[^\]]*\] (.+?) \(Part \d+\)$') {
                                        $altLeaf = $Matches[1]
                                    }
                                    if (-not [string]::IsNullOrWhiteSpace($altLeaf)) {
                                        $altLeafKey = "$($newRule.RuleStoreName)|$($altLeaf.Trim().ToLower())"
                                        if ($existingRowsByLeaf.ContainsKey($altLeafKey) -and
                                            $existingRowsByLeaf[$altLeafKey] -ne $existing) {
                                            $altCandidate = $existingRowsByLeaf[$altLeafKey]
                                            # FIXED (Administrator direction, 2026-07-2X, Part 1/Part 2
                                            # SendersDomain cross-contamination bug -- second occurrence of
                                            # the Part-N collision, this time on the altExisting duplicate-
                                            # recovery lookup rather than the primary key fixed last
                                            # session): the Leaf key above has no Part-N distinguisher, so
                                            # when newRule is itself a Part-N rule (e.g. Part 1), the same-
                                            # leaf row this finds is very often its own sibling (Part 2) --
                                            # a legitimate, different, valid row, NOT the orphaned-rename-
                                            # duplicate scenario this whole block exists to recover from.
                                            # The prior "-ne $existing" check alone cannot tell the two
                                            # cases apart, since a sibling Part-N row genuinely IS a
                                            # different object. Confirmed live: this caused Part 1 and
                                            # Part 2's SendersDomain values to swap on rerun even though
                                            # RuleName itself was already correct. Guard: only accept this
                                            # candidate as a genuine duplicate if it does NOT itself look
                                            # like a valid Part-N rule for this same base folder -- a real
                                            # sibling is skipped; a genuine pre-rename orphan (which does
                                            # not match the Part-N pattern) still passes through exactly as
                                            # before.
                                            $altCandidateIsPartNSibling = $false
                                            if ($newRule.RuleName -match '^Rule: \[[^\]]*\] (.+?) \(Part \d+\)$') {
                                                $newRuleBaseLabel = $Matches[1]
                                                if ($altCandidate.RuleName -match '^Rule: \[[^\]]*\] (.+?) \(Part \d+\)$' -and
                                                    $Matches[1] -eq $newRuleBaseLabel) {
                                                    $altCandidateIsPartNSibling = $true
                                                }
                                            }
                                            if (-not $altCandidateIsPartNSibling) {
                                                $altExisting = $altCandidate
                                            }
                                        }
                                    }
                                }
                            }
                            $existingSendersDomain = ''
                            # ADDED (Administrator direction -- corrected priority, live-tested
                            # 2026-07-13): when a duplicate row was found ($altExisting is
                            # set), ALWAYS use its SendersDomain, not just as a fallback for
                            # a blank/invalid primary-matched value. The primary-matched row
                            # ($existing, via the Path key) is, in this exact duplicate
                            # scenario, the row Script 00's own best-guess logic wrote on a
                            # PRIOR run -- its SendersDomain PASSES validation (it's a
                            # well-formed value, just the wrong one), so a blank-or-invalid-
                            # only fallback never engaged. $altExisting is the row that still
                            # carries whatever SendersDomain was set when the rule was
                            # originally added/corrected -- that is the value that must win,
                            # every time a duplicate is found, regardless of whether the
                            # primary row's own value happens to also be well-formed.
                            $sendersDomainSource = if ($altExisting) { $altExisting } else { $existing }
                            if ($sendersDomainSource.PSObject.Properties['SendersDomain'] -and
                                $sendersDomainSource.SendersDomain -and
                                -not [string]::IsNullOrWhiteSpace($sendersDomainSource.SendersDomain) -and
                                $sendersDomainSource.SendersDomain -notmatch '\\') {
                                $existingSendersDomainTokens = @(ConvertTo-NormalizedSenderDomains -RawValue $sendersDomainSource.SendersDomain)
                                if ($existingSendersDomainTokens.Count -gt 0) {
                                    $existingSendersDomain = ($existingSendersDomainTokens -join ' ')
                                }
                            }
                            # If the preferred source (alt row, when a duplicate exists) had
                            # no valid SendersDomain of its own, fall back to the primary-
                            # matched row's value instead of leaving this blank.
                            if (-not $existingSendersDomain -and $altExisting -and
                                $existing.PSObject.Properties['SendersDomain'] -and
                                $existing.SendersDomain -and
                                -not [string]::IsNullOrWhiteSpace($existing.SendersDomain) -and
                                $existing.SendersDomain -notmatch '\\') {
                                $fallbackSendersDomainTokens = @(ConvertTo-NormalizedSenderDomains -RawValue $existing.SendersDomain)
                                if ($fallbackSendersDomainTokens.Count -gt 0) {
                                    $existingSendersDomain = ($fallbackSendersDomainTokens -join ' ')
                                }
                            }
                            if ($existingSendersDomain) {
                                $mergedRule.SendersDomain = $existingSendersDomain
                            }
                            # NEW (2026-07-23, Administrator direction): if the LIVE Outlook rule
                            # currently has its own SenderAddress condition set (captured
                            # into $newRule.SendersDomain by Get-OutlookRules via the new
                            # Get-RuleSenderAddressWords helper), that value wins over
                            # whatever was just preserved from the CSV above -- including
                            # a prior manual operator edit. Administrator's explicit direction:
                            # the rule's SenderAddress words in the Outlook UI Rules
                            # Manager are more likely to be kept current than a manually-
                            # edited CSV value, so the CSV should always stay in sync
                            # with the live rule when that condition is present. This is
                            # intentionally the LAST word on SendersDomain in this merge
                            # block -- it does not alter the preservation logic above,
                            # which still applies unchanged whenever the live rule has no
                            # SenderAddress condition of its own.
                            if ($newRule.PSObject.Properties['SendersDomain'] -and
                                -not [string]::IsNullOrWhiteSpace($newRule.SendersDomain)) {
                                $liveSenderAddressTokens = @(ConvertTo-NormalizedSenderDomains -RawValue $newRule.SendersDomain)
                                if ($liveSenderAddressTokens.Count -gt 0) {
                                    $mergedRule.SendersDomain = ($liveSenderAddressTokens -join ' ')
                                }
                            }
                            # Preserve Conditions (Issue 4 fix, added 2026-07-06, Administrator; REFINED
                            # 3rd pass): Conditions is purely informational/display -- confirmed
                            # never consumed by any deploy logic (only TargetFolderPath,
                            # TargetFolderEntryID, and SendersDomain actually drive what gets
                            # deployed). Administrator's direction: this column SHOULD update to reflect
                            # genuine current state on a rerun -- e.g. if Administrator manually adds a
                            # Subject condition via the Outlook Rules Manager UI, the next
                            # Script 00 scan should show it -- but must never regress from real
                            # content to blank/cryptic/a known failure placeholder, which would
                            # wrongly suggest something happened to the live rule. See
                            # Test-MeaningfulRuleSummaryText for the exact blank/placeholder/
                            # too-short criteria. Only blocks the fresh value when EXISTING is
                            # meaningful AND fresh is NOT -- any other combination lets the fresh
                            # scan result through normally.
                            if ($existing.PSObject.Properties['Conditions'] -and
                                (Test-MeaningfulRuleSummaryText -Value $existing.Conditions) -and
                                -not (Test-MeaningfulRuleSummaryText -Value $mergedRule.Conditions)) {
                                $mergedRule.Conditions = $existing.Conditions
                            }
                            # Preserve Actions (Issue 5 fix, added 2026-07-06, Administrator; REFINED 3rd
                            # pass): same rationale and mechanism as Conditions immediately above
                            # -- updates freely to reflect genuine current state, only blocked
                            # from regressing to blank/cryptic/a known failure placeholder when
                            # the existing value was real. This replaces the narrower 2nd-pass
                            # version (which only blocked regression to the literal string
                            # '[No actions]') and the even-narrower unconditional-freeze version
                            # before that -- both were superseded once Administrator clarified the actual
                            # goal is accuracy/no-confusion, not a blanket freeze.
                            if ($existing.PSObject.Properties['Actions'] -and
                                (Test-MeaningfulRuleSummaryText -Value $existing.Actions) -and
                                -not (Test-MeaningfulRuleSummaryText -Value $mergedRule.Actions)) {
                                $mergedRule.Actions = $existing.Actions
                            }
                            # Preserve LastDeployedRun -- Script 03 idempotency sidecar (Memory #16).
                            # This timestamp records when each secondary store rule was last
                            # successfully deployed. Must survive Script 00 reruns so Script 03
                            # does not incorrectly treat already-deployed rules as needing Full
                            # purge/recreate. Only preserved if the column exists on the existing
                            # row (older CSVs without the column will naturally have blank/missing).
                            if ($existing.PSObject.Properties['LastDeployedRun'] -and
                                -not [string]::IsNullOrWhiteSpace($existing.LastDeployedRun)) {
                                $mergedRule.LastDeployedRun = $existing.LastDeployedRun
                            }
                            # Preserve LastTargetRun -- SAME preservation treatment as
                            # LastDeployedRun immediately above, but for the independent
                            # folder-target-remap idempotency column (added 2026-07-07, Administrator
                            # direction). Must survive Script 00 reruns for the exact same
                            # reason -- so Script 03 Phase 3 does not incorrectly re-process
                            # a rule whose folder target was already successfully remapped.
                            if ($existing.PSObject.Properties['LastTargetRun'] -and
                                -not [string]::IsNullOrWhiteSpace($existing.LastTargetRun)) {
                                $mergedRule.LastTargetRun = $existing.LastTargetRun
                            }
                            # Always True -- never preserve False from a prior scan.
                            # NeedsFolderUpdate and IsEnabled must always be True in the CSV.
                            # Code never sets either column to False under any circumstance.
                            $mergedRule.NeedsFolderUpdate = 'True'
                            $mergedRule.IsEnabled         = 'True'
                            $mergedRules.Add($mergedRule)
                            $mergeStats.Retained++
                            # FIXED (Administrator direction, this same fallback fix, THIRD pass):
                            # remove the matched row from ALL THREE dictionaries, always keyed
                            # by $existing's OWN values (not $newRule's) -- the row being
                            # removed is $existing, so its removal key must be derived from its
                            # own RuleName/TargetFolderPath, which can differ from $newRule's
                            # (that mismatch is exactly why the fallback had to engage in the
                            # first place). The same source row was inserted into all three
                            # dictionaries during the build loop above under its own three
                            # keys -- leaving any of them un-removed lets that entry survive to
                            # the orphan-detection pass further below and be wrongly treated as
                            # orphaned/deleted, even though it was already correctly merged.
                            $existingRemoveKeyPath = if ($existing.PSObject.Properties['TargetFolderPath']) { [string]$existing.TargetFolderPath } else { '' }
                            $existingRemoveKey = "$($existing.RuleStoreName)|$existingRemoveKeyPath"
                            $existingRows.Remove($existingRemoveKey)
                            if ($existing.RuleStoreName -and $existing.RuleName) {
                                $existingRowsByName.Remove("$($existing.RuleStoreName)|$($existing.RuleName)")
                            }
                            $existingRemoveLeaf = ($existingRemoveKeyPath -split '\\')[-1]
                            if ([string]::IsNullOrWhiteSpace($existingRemoveLeaf) -and $existing.RuleName -match '^Rule: \[[^\]]*\] (.+?) \(Part \d+\)$') {
                                $existingRemoveLeaf = $Matches[1]
                            }
                            if ($existing.RuleStoreName -and -not [string]::IsNullOrWhiteSpace($existingRemoveLeaf)) {
                                $existingRowsByLeaf.Remove("$($existing.RuleStoreName)|$($existingRemoveLeaf.Trim().ToLower())")
                            }
                        }
                        else {
                            # New rule not in existing CSV -- add with defaults
                            # Inject LastDeployedRun if missing (predates this column)
                            if (-not $newRule.PSObject.Properties['LastDeployedRun']) {
                                Add-Member -InputObject $newRule -MemberType NoteProperty -Name 'LastDeployedRun' -Value '' -Force
                            }
                            # Inject LastTargetRun if missing (predates this column) --
                            # same pattern as LastDeployedRun immediately above.
                            if (-not $newRule.PSObject.Properties['LastTargetRun']) {
                                Add-Member -InputObject $newRule -MemberType NoteProperty -Name 'LastTargetRun' -Value '' -Force
                            }
                            $mergedRules.Add($newRule)
                            $mergeStats.Added++
                            Write-OMMigrateLog -Message "New rule added to CSV: $($newRule.RuleName)" -Level INFO
                        }
                    }

                    # Preserve rules from prior runs that are not in the current
                    # live session. Rules from stores no longer mounted (e.g. after
                    # POP3->IMAP conversion) must never be silently dropped -- they
                    # are needed for Script 03 rules recreation and reporting.
                    # Build a set of RuleStoreName+RuleName keys from the LIVE SCAN
                    # only -- used to block orphan rows that are stale duplicates of
                    # live rules. Operator-added rows with names NOT in the live scan
                    # are always preserved regardless -- admin may manually add rules
                    # to the CSV.
                    #
                    # Key changed 2026-06-27 from RuleName-only to RuleStoreName+
                    # RuleName -- the live scan now covers every account, not just
                    # the default store. A RuleName-only check would incorrectly
                    # treat an orphan row from account A as a stale duplicate of a
                    # currently-live rule from account B sharing the same name, and
                    # silently drop a still-needed row.
                    $liveRuleNames = [System.Collections.Generic.HashSet[string]]::new(
                        [System.StringComparer]::OrdinalIgnoreCase
                    )
                    # ADDED (Administrator direction -- SendersDomain/preservation-block clobber fix,
                    # FOURTH pass, live-tested 2026-07-13 on "TestProfile"/testaccount@example.com):
                    # $liveRuleLeaves -- same leaf-folder-segment identity as
                    # $existingRowsByLeaf earlier in this function, built from the LIVE SCAN
                    # this time instead of the CSV. Root cause traced live via DEBUG log:
                    # a rule renamed by consolidation BETWEEN two Script 00 runs (raw name
                    # 'Ommtest' -> live name 'Rule: [account] Label (Part N)') is correctly
                    # matched and merged under its NEW name via the leaf fallback in the
                    # merge loop above -- but the OLD CSV row (still under the raw 'Ommtest'
                    # name) is left behind in $existingRows, unconsumed, and falls through to
                    # THIS orphan pass. The orphan-skip check here only ever compared
                    # RuleName+RuleStoreName against the live scan's CURRENT rule names, so a
                    # stale row under an OLD name was never recognized as a duplicate of the
                    # SAME live rule under its NEW name -- confirmed live via DEBUG log line
                    # "Prior-run rule preserved... | Ommtest" immediately after the merge loop
                    # had already (correctly) consumed the live rule under its new name.
                    # Adding the same leaf check here closes this gap: an orphan row whose
                    # leaf folder name matches a live rule's leaf folder name is a stale
                    # duplicate of that live rule under a prior name, regardless of what its
                    # own RuleName says.
                    $liveRuleLeaves = [System.Collections.Generic.HashSet[string]]::new(
                        [System.StringComparer]::OrdinalIgnoreCase
                    )
                    foreach ($lr in $Rules) {
                        if ($lr.RuleName -and $lr.RuleStoreName) {
                            [void]$liveRuleNames.Add("$($lr.RuleStoreName)|$($lr.RuleName)")
                        }
                        if ($lr.RuleStoreName) {
                            $lrLeafSource = if ($lr.PSObject.Properties['TargetFolderPath']) { [string]$lr.TargetFolderPath } else { '' }
                            $lrLeaf = ($lrLeafSource -split '\\')[-1]
                            if ([string]::IsNullOrWhiteSpace($lrLeaf) -and $lr.RuleName -match '^Rule: \[[^\]]*\] (.+?) \(Part \d+\)$') {
                                $lrLeaf = $Matches[1]
                            }
                            if (-not [string]::IsNullOrWhiteSpace($lrLeaf)) {
                                [void]$liveRuleLeaves.Add("$($lr.RuleStoreName)|$($lrLeaf.Trim().ToLower())")
                            }
                        }
                    }

                    $mergeStats.PreservedPrior = 0
                    foreach ($orphanKey in @($existingRows.Keys)) {
                        $orphanRule = $existingRows[$orphanKey]
                        # Skip blank separator rows
                        if (-not $orphanRule.RuleStoreName -or
                            [string]::IsNullOrWhiteSpace($orphanRule.RuleStoreName)) { continue }
                        # Skip if a rule with this RuleStoreName+RuleName exists in the
                        # live scan -- this orphan is a stale duplicate from a prior
                        # all-stores scan. Operator-added rules with unique
                        # RuleStoreName+RuleName combinations (not in live scan) are
                        # always preserved.
                        # Also check quote-stripped name -- quoted variants (e.g. 'albert drent')
                        # are stale duplicates of unquoted live rules and should not be preserved.
                        $orphanNameStripped = $orphanRule.RuleName.Trim("'")
                        $orphanLiveKey         = "$($orphanRule.RuleStoreName)|$($orphanRule.RuleName)"
                        $orphanLiveKeyStripped = "$($orphanRule.RuleStoreName)|$orphanNameStripped"
                        # ADDED (Administrator direction, this same fix, FOURTH pass): leaf-based
                        # duplicate check, tried when the name-based checks above miss --
                        # catches a stale row under an OLD (pre-rename) name whose target
                        # folder leaf matches a live rule's leaf under its NEW name.
                        $orphanLeafSource = if ($orphanRule.PSObject.Properties['TargetFolderPath']) { [string]$orphanRule.TargetFolderPath } else { '' }
                        $orphanLeaf = ($orphanLeafSource -split '\\')[-1]
                        if ([string]::IsNullOrWhiteSpace($orphanLeaf) -and $orphanRule.RuleName -match '^Rule: \[[^\]]*\] (.+?) \(Part \d+\)$') {
                            $orphanLeaf = $Matches[1]
                        }
                        $orphanLiveLeafKey = if (-not [string]::IsNullOrWhiteSpace($orphanLeaf)) { "$($orphanRule.RuleStoreName)|$($orphanLeaf.Trim().ToLower())" } else { $null }
                        if ($orphanRule.RuleName -and (
                                $liveRuleNames.Contains($orphanLiveKey) -or
                                $liveRuleNames.Contains($orphanLiveKeyStripped) -or
                                ($null -ne $orphanLiveLeafKey -and $liveRuleLeaves.Contains($orphanLiveLeafKey)))) {
                            Write-OMMigrateLog -Message (
                                "Orphan rule skipped -- name or leaf folder exists in live scan (stale duplicate): " +
                                "$($orphanRule.RuleStoreName) | $($orphanRule.RuleName)"
                            ) -Level DEBUG
                            continue
                        }
                        # Always force True -- never preserve False from prior runs.
                        # NeedsFolderUpdate must always be True regardless of what was stored.
                        $orphanRule.NeedsFolderUpdate = 'True'
                        $orphanRule.IsEnabled         = 'True'
                        # Apply SendersDomain best-guess ONLY if genuinely blank or fails
                        # validation. FIXED 2026-07-06, Administrator, found during macro test --
                        # this was the exact same stale "@ or backslash means invalid"
                        # blanket rejection already fixed at the merge-preserve site
                        # above, but never applied here in the separate orphan-row path.
                        # A real, validly-preserved full email address (which
                        # necessarily contains "@") would have been wrongly judged
                        # invalid and overwritten with the folder-name guess by this
                        # exact site. Administrator's explicit requirement: SendersDomain must
                        # ONLY ever be touched on first creation (best-guess) or when
                        # the existing value fails validation -- once an admin has set
                        # a real value that passes validation, it must NEVER be touched
                        # again, period. Reusing ConvertTo-NormalizedSenderDomains here
                        # too, same as the merge-preserve site, so both places agree on
                        # what counts as valid.
                        $orphanNeedsDefault = $true
                        if ($orphanRule.PSObject.Properties['SendersDomain'] -and
                            -not [string]::IsNullOrWhiteSpace($orphanRule.SendersDomain)) {
                            $orphanSendersDomainTokens = @(ConvertTo-NormalizedSenderDomains -RawValue $orphanRule.SendersDomain)
                            if ($orphanSendersDomainTokens.Count -gt 0) {
                                $orphanNeedsDefault = $false
                            }
                        }
                        if ($orphanNeedsDefault -and
                            $orphanRule.PSObject.Properties['TargetFolderPath'] -and
                            -not [string]::IsNullOrWhiteSpace($orphanRule.TargetFolderPath)) {
                            $orphanSegments = $orphanRule.TargetFolderPath -split '\\'
                            $orphanRule.SendersDomain = ($orphanSegments | Where-Object { $_ -ne '' } | Select-Object -Last 1).ToLower()
                        }
                        # Inject LastDeployedRun if missing (predates this column)
                        if (-not $orphanRule.PSObject.Properties['LastDeployedRun']) {
                            Add-Member -InputObject $orphanRule -MemberType NoteProperty -Name 'LastDeployedRun' -Value '' -Force
                        }
                        # Inject LastTargetRun if missing (predates this column) --
                        # same pattern as LastDeployedRun immediately above.
                        if (-not $orphanRule.PSObject.Properties['LastTargetRun']) {
                            Add-Member -InputObject $orphanRule -MemberType NoteProperty -Name 'LastTargetRun' -Value '' -Force
                        }
                        $mergedRules.Add($orphanRule)
                        $mergeStats.PreservedPrior++
                        Write-OMMigrateLog -Message (
                            "Prior-run rule preserved (store not in live session): " +
                            "$($orphanRule.RuleStoreName) | $($orphanRule.RuleName)"
                        ) -Level DEBUG
                    }

                    $Rules = $mergedRules
                    Write-OMMigrateLog -Message (
                        "Rules CSV merge complete: $($mergeStats.Retained) retained, " +
                        "$($mergeStats.Added) new, " +
                        "$($mergeStats.PreservedPrior) preserved from prior runs"
                    ) -Level INFO
                }
            }

            # EXTRACTED 2026-07-09 (Administrator direction) from OMMigrate-Outlook_WIP.psm1,
            # a 6-day-old (2026-07-03/04) work-in-progress file -- this is the
            # TargetStoreName hardcode fix work paused mid-implementation that
            # session, extended here to fill blank TargetStoreName values using
            # the operator-selected RulesEngine.ArchiveStoreMappings (see
            # OMMigrate-Core_WIP.psm1's Save-OMMigrateArchiveStoreMappings for
            # the settings-persistence half of this feature). NOT a completed/
            # finished feature -- extracted as a jumpstart on resuming this
            # work, not yet wired to a picker UI or live-tested.
            #
            # Assign TargetStoreName for any row still blank at this point --
            # Item(i) fallback, orphan rows, or a genuinely new rule where the
            # live COM read of the target folder's store failed. This is the
            # single choke point every row passes through before export,
            # matching the SendersDomain comma-scrub pattern elsewhere in
            # this function.
            #
            # Priority order:
            #   1. If already non-blank (live COM value, merge-preserved
            #      value, or a manual admin edit) -- leave untouched. This
            #      block only ever fills in rows that are still blank.
            #   2. Look up RuleStoreName in the operator-selected
            #      RulesEngine.ArchiveStoreMappings (Script 00's
            #      TargetStoreName picker) and assign the matching PST's
            #      TargetStoreName. Supports Administrator's use case where many/all
            #      accounts route to one shared archive PST, as well as a
            #      one-PST-per-account or any other mix the operator picks.
            #      No store or account name is hardcoded here -- the
            #      mapping is entirely operator-selected, live-detected data.
            #   3. If no mapping entry matches this row's RuleStoreName
            #      (operator has not run the picker yet, or this account
            #      was left unmapped on purpose), fall back to the ORIGINAL
            #      behavior (Default TargetStoreName to RuleStoreName if
            #      blank -- rules always target folders in the same store
            #      they run against, so RuleStoreName is the correct
            #      fallback default). This preserves exact prior behavior
            #      for every row/account that hasn't opted into the new
            #      mapping feature.
            # NOTE: $archiveStoreMappings / $ruleStoreToTarget were previously
            # (re)built here on every export. FIXED 2026-07-10, Administrator direction:
            # moved to the top of this function (see $ruleStoreToTarget /
            # $archiveStoreMappingsForMerge above) so the merge-preserve block
            # earlier in this function can also consult the same lookup --
            # both blocks now share one single build, not two separate ones
            # that could theoretically drift out of sync with each other.

            foreach ($r in $Rules) {
                if (-not $r.PSObject.Properties['TargetStoreName'] -or
                    [string]::IsNullOrWhiteSpace($r.TargetStoreName)) {
                    if ($ruleStoreToTarget.ContainsKey($r.RuleStoreName)) {
                        $r.TargetStoreName = $ruleStoreToTarget[$r.RuleStoreName]
                    }
                    elseif ($exportDefaultArchiveDisplayName) {
                        # CORRECTED 2026-07-10, Administrator direction. Prior fallback here
                        # (TargetStoreName = RuleStoreName) wrote the account's own
                        # email address into a column named TargetStoreName, which
                        # is misleading -- the column no longer describes an actual
                        # store, and downstream resolution (Invoke-DeployConsolidatedRules,
                        # Invoke-RulesRecreation, Strategy 1+2) never matches an email
                        # against a live store anyway, so it silently falls back to
                        # the default archive regardless. This now writes the REAL
                        # resolved default archive's display name directly, so the
                        # CSV states the actual outcome instead of a value that
                        # doesn't reflect where the rule's folder target truly lives.
                        $r.TargetStoreName = $exportDefaultArchiveDisplayName
                    }
                    else {
                        # ORIGINAL behavior (pre-extraction), preserved as the
                        # fallback default when no ArchiveStoreMappings entry
                        # matches this row's RuleStoreName AND the default archive
                        # could not be resolved (Namespace not supplied to this
                        # call, or that specific PST is not currently attached).
                        $r.TargetStoreName = $r.RuleStoreName
                    }
                }
            }

            # Sort by RuleStoreName then TargetFolderPath and insert
            # blank separator rows between rule store groups for Excel review.
            $sortedRules   = @($Rules | Sort-Object RuleStoreName, TargetFolderPath)
            $finalRules    = [System.Collections.Generic.List[PSCustomObject]]::new()
            $lastStore     = $null
            $blankTemplate = [PSCustomObject]@{
                RuleStoreName = ''; TargetStoreName = ''; RuleName = ''; LastDeployedRun = ''; LastTargetRun = ''; TargetFolderPath = ''; SendersDomain = ''
                NeedsFolderUpdate = ''; IsEnabled = ''; ExecutionOrder = ''; RuleType = ''; StopProcessing = ''
                Conditions = ''; Actions = ''; TargetFolderEntryID = ''; Notes = ''
            }

            foreach ($rule in $sortedRules) {
                if ($null -ne $lastStore -and $rule.RuleStoreName -ne $lastStore) {
                    # Insert blank separator row between rule store groups
                    $finalRules.Add($blankTemplate)
                }
                $finalRules.Add($rule)
                $lastStore = $rule.RuleStoreName
            }

            # Check for Excel file lock before writing -- poll until released
            if (Test-Path $OutputPath) {
                $csvLocked = $true
                $lockWaited = 0
                while ($csvLocked) {
                    try {
                        $lt = [System.IO.File]::Open($OutputPath, 'Open', 'ReadWrite', 'None')
                        $lt.Close()
                        $lt.Dispose()
                        $csvLocked = $false
                    }
                    catch {
                        if ($lockWaited -eq 0) {
                            Write-OMMigrateLog -Message 'rules_inventory.csv is locked by Excel -- waiting for Excel to close.' `
                                               -Level WARN
                            Write-Host ''
                            Write-Host '  rules_inventory.csv is open in Excel.' -ForegroundColor Yellow
                            Write-Host '  Save any changes and close Excel -- script will continue automatically.' `
                                       -ForegroundColor Yellow
                        }
                        Start-Sleep -Seconds 3
                        $lockWaited += 3
                        if ($lockWaited % 15 -eq 0) {
                            Write-Host "  Still waiting for Excel to close... ($($lockWaited)s)" `
                                       -ForegroundColor DarkGray
                        }
                    }
                }
                if ($lockWaited -gt 0) {
                    Write-Host '  Excel closed -- writing rules_inventory.csv.' -ForegroundColor Green
                    Write-OMMigrateLog -Message "rules_inventory.csv lock released after $($lockWaited)s -- continuing." `
                                       -Level INFO
                }
            }

            # Final safety sort of data rows only -- guarantees correct order
            # regardless of merge or orphan processing sequence.
            # Blank separator rows are excluded from sort and re-inserted between groups.
            $dataRowsSorted = @($finalRules |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_.RuleStoreName) } |
                Sort-Object RuleStoreName, TargetFolderPath)
            $finalRulesSorted = [System.Collections.Generic.List[PSCustomObject]]::new()
            $lastStoreSorted  = $null
            foreach ($r in $dataRowsSorted) {
                if ($null -ne $lastStoreSorted -and $r.RuleStoreName -ne $lastStoreSorted) {
                    $finalRulesSorted.Add($blankTemplate)
                }
                $finalRulesSorted.Add($r)
                $lastStoreSorted = $r.RuleStoreName
            }
            # Final SendersDomain comma scrub (2026-07-02, Administrator): a comma in
            # SendersDomain is not a supported separator -- Export-Csv below
            # auto-quotes any cell containing a literal comma, but the
            # companion VBA macro (Module3.bas) reads rules_inventory.csv
            # with a plain, non-quote-aware line split on comma, so a comma
            # inside this column would silently misalign every column after
            # it there. This is the single choke point every row passes
            # through before export, regardless of which code path produced
            # it (merge-preserved, orphan best-guess, or newly live-scanned)
            # -- replacing any comma with a space here guarantees no comma
            # ever reaches the CSV file. Space is already a valid
            # SendersDomain separator (see ConvertTo-NormalizedSenderDomains),
            # so this is a lossless normalization, not a data loss.
            foreach ($r in $finalRulesSorted) {
                if ($r.PSObject.Properties['SendersDomain'] -and
                    -not [string]::IsNullOrWhiteSpace($r.SendersDomain) -and
                    $r.SendersDomain.Contains(',')) {
                    $r.SendersDomain = $r.SendersDomain.Replace(',', ' ')
                }
            }

            # Select properties in canonical column order before writing.
            # Export-Csv writes columns in property-declaration order -- without
            # an explicit Select-Object, Add-Member'd properties (like
            # LastDeployedRun injected into pre-existing rows) appear at the end
            # rather than in their defined position after RuleName.
            $canonicalColumns = @(
                'RuleStoreName','TargetStoreName','RuleName','LastDeployedRun','LastTargetRun',
                'TargetFolderPath','SendersDomain','NeedsFolderUpdate','IsEnabled',
                'ExecutionOrder','RuleType','StopProcessing','Conditions',
                'Actions','TargetFolderEntryID','Notes'
            )
            $finalRulesSorted | Select-Object $canonicalColumns | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            Write-OMMigrateLog -Message "Rules CSV written: $OutputPath ($($Rules.Count) rules, Retained=$($mergeStats.Retained), Added=$($mergeStats.Added))" `
                               -Level INFO
            Write-AuditEntry  -Action 'CSV_RULES_WRITTEN' `
                              -Detail "Path=$OutputPath | Rules=$($Rules.Count) | Retained=$($mergeStats.Retained) | Added=$($mergeStats.Added)"
        }
        catch {
            Write-OMMigrateLog -Message "Failed to write rules CSV: $_" -Level ERROR
        }
    }
    else {
        Write-OMMigrateLog -Message "WhatIf: Would write rules CSV to $OutputPath" `
                           -Level INFO -WhatIfPrefix
    }

    return $OutputPath
}


# ============================================================
#  REGION: ACCOUNT OPERATIONS (Scripts 02 and 03)
# ============================================================

function Remove-POP3Account {
    <#
    .SYNOPSIS
        Guides the operator through manually removing a POP3 account
        from Outlook using Outlook's own Account Settings dialog.

    .DESCRIPTION
        POP3 account removal must be performed through Outlook's own
        File > Account Settings > Account Settings > Remove button.
        Direct registry deletion was attempted and abandoned -- it
        bypasses Outlook's internal MAPI cleanup (store key removal,
        Send/Receive group references, search folder definitions) and
        leaves the profile in a corrupted state that crashes Outlook
        on next startup. Outlook's own Remove button handles all of
        this cleanup correctly and safely.

        CRITICAL LESSON LEARNED (May 2026 session):
        Registry deletion of the POP3 account key without the full
        MAPI cleanup that Outlook's Remove button performs caused a
        profile corruption that required ntuser.dat restore from a
        Macrium Reflect backup to recover from. Never use registry
        deletion for account removal.

        This function:
          1. Verifies the backup PST exists and meets minimum size
             (safety gate -- will not proceed without a confirmed backup)
          2. Displays clear step-by-step instructions for the operator
             to remove the POP3 account via Outlook's Account Settings
          3. Pauses and waits for the operator to confirm removal
             using a C-to-continue prompt (prevents accidental Enter
             from advancing the script prematurely)
          4. Returns $true when the operator confirms removal is done

        Outlook must be OPEN when this function is called -- the operator
        needs access to File > Account Settings to perform the removal.
        The calling script (Script 02) keeps the COM session active
        through this phase, then releases it before the IMAP add phase.

        Passwords are never handled, stored, or logged by this function.

    .PARAMETER EmailAddress
        Email address of the POP3 account to remove.

    .PARAMETER BackupPSTPath
        Full path to the verified PST backup file for this account.
        Must exist before this function will proceed.

    .OUTPUTS
        [bool] -- $true if operator confirmed removal, $false if blocked.

    .EXAMPLE
        $removed = Remove-POP3Account `
            -EmailAddress   'user@domain.com' `
            -BackupPSTPath  'C:\Migration\Backups\user_domain_com.pst'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,

        [Parameter(Mandatory = $true)]
        [string]$BackupPSTPath
    )

    # -- Safety gate: verify backup exists --------------------
    # This is the last line of defense before the operator removes
    # the POP3 account. If the backup is missing or too small, refuse
    # to proceed -- the operator must re-run Script 01 first.
    if (-not (Test-Path $BackupPSTPath)) {
        Write-OMMigrateLog -Message (
            "SAFETY BLOCK: Cannot remove account '$EmailAddress'. " +
            "Backup PST not found at: $BackupPSTPath"
        ) -Level ERROR
        Write-AuditEntry -Action 'ACCOUNT_REMOVE_BLOCKED' `
                         -AccountEmail $EmailAddress `
                         -Detail "Backup PST not found: $BackupPSTPath" `
                         -Outcome 'FAILED'
        return $false
    }

    $backupInfo = Get-Item $BackupPSTPath

    # Read minimum size from settings -- default 0 (no minimum enforced)
    $minBackupBytes = 0L
    try {
        $minMB = $Global:OMMigrate.Settings.BackupVerification.MinimumSizeMB
        if ($null -ne $minMB) { $minBackupBytes = [long]($minMB * 1MB) }
    }
    catch { }

    if ($minBackupBytes -gt 0 -and $backupInfo.Length -lt $minBackupBytes) {
        Write-OMMigrateLog -Message (
            "SAFETY BLOCK: Backup PST for '$EmailAddress' is suspiciously small " +
            "($([Math]::Round($backupInfo.Length/1KB))KB). Removal blocked."
        ) -Level ERROR
        return $false
    }

    # -- WhatIf mode -- log and return -------------------------
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message "WhatIf: Would guide operator to remove POP3 account via Outlook Account Settings: $EmailAddress" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Operator would use File > Account Settings > Account Settings > select account > Remove" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Outlook handles all internal MAPI cleanup -- no registry writes" `
                           -Level INFO -WhatIfPrefix
        return $true
    }

    Write-OMMigrateLog -Message "Guiding operator through POP3 account removal: $EmailAddress" `
                       -Level INFO

    # -- Display guided removal instructions ------------------
    # Outlook must be open at this point -- the COM session in Script 02
    # is kept active through this phase specifically so the operator
    # can access File > Account Settings without re-opening Outlook.
    $safeEmail = Invoke-OMMigrateSanitize -Text $EmailAddress

    Write-Host ''
    Write-Host ('  ' + ('=' * 54)) -ForegroundColor Yellow
    Write-Host "  ACTION REQUIRED -- Remove POP3 account from Outlook" -ForegroundColor Yellow
    Write-Host ('  ' + ('=' * 54)) -ForegroundColor Yellow
    Write-Host ''
    Write-Host "  Account to remove:  $safeEmail" -ForegroundColor White
    Write-Host ''
    Write-Host '  Follow these steps in Outlook now:' -ForegroundColor White
    Write-Host ''
    Write-Host '  1. Click  File  (top left)' -ForegroundColor Gray
    Write-Host '  2. Click  Account Settings  then  Account Settings  again' -ForegroundColor Gray
    Write-Host '  3. On the  Email  tab -- click to select the POP3 account:' -ForegroundColor Gray
    Write-Host "       $safeEmail" -ForegroundColor Cyan
    Write-Host '  4. Click  Remove' -ForegroundColor Gray
    Write-Host '  5. Click  Yes  when Outlook asks you to confirm' -ForegroundColor Gray
    Write-Host '  6. Verify the account is gone from the Account Settings list' -ForegroundColor Gray
    Write-Host '  7. Click  Close  to close Account Settings' -ForegroundColor Gray
    Write-Host ''
    Write-Host '  NOTE: The old POP3 folder may still appear in the left folder' -ForegroundColor DarkGray
    Write-Host '  pane after removal -- this is expected. The script will detach' -ForegroundColor DarkGray
    Write-Host '  it automatically. Do not close Outlook.' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  NOTE: Your email data is safe.' -ForegroundColor Green
    Write-Host "  The backup PST ($([Math]::Round($backupInfo.Length/1KB))KB) is confirmed in your Backups folder." `
               -ForegroundColor Green
    Write-Host '  Removing the account from Outlook does NOT delete your emails.' `
               -ForegroundColor Green
    Write-Host ''
    Write-Host ('  ' + ('=' * 54)) -ForegroundColor Yellow
    Write-Host ''

    # -- C-to-continue prompt ----------------------------------
    # Outlook must remain OPEN after this prompt -- the auto-detach
    # block needs the active COM session. Script closes Outlook
    # automatically after detach completes.
    # Operator must type C deliberately.
    do {
        $response = Read-Host '  Type C and press Enter when the POP3 account has been removed'
    } while ($response.Trim().ToUpper() -ne 'C')

    Write-OMMigrateLog -Message "POP3 account removal confirmed by operator: $EmailAddress" `
                       -Level INFO
    Write-AuditEntry  -Action 'ACCOUNT_REMOVED_POP3' `
                      -AccountEmail $EmailAddress `
                      -Detail ("POP3 account removed via Outlook Account Settings (guided manual). " +
                               "Outlook handled all internal MAPI cleanup. " +
                               "Backup confirmed: $BackupPSTPath")

    # -- Auto-detach orphaned PST store from nav pane ----------
    # Runs while Outlook is still open -- COM session namespace is
    # active. Script closes Outlook automatically after this block.
    # The PST file itself is NOT deleted -- stays on disk for Scripts 02/03.
    # Expected PST filename: emailaddress.pst (Outlook uses the literal
    # email address including @ in the filename).
    # If the store cannot be found or detached, the operator is prompted
    # to right-click close it manually before the script closes Outlook.
    Write-Host ''
    Write-Host '  Attempting to detach old POP3 store from folder pane...' `
               -ForegroundColor DarkGray

    $pstFileName  = $EmailAddress + '.pst'
    $detachResult = 'NOT_FOUND'
    $detachStore  = $null

    try {
        $namespace = Get-OutlookNamespace
        if ($namespace) {
            $stores = $namespace.Stores
            Register-COMObject -ComObject $stores

            for ($i = 1; $i -le $stores.Count; $i++) {
                try {
                    $s = $stores.Item($i)
                    Register-COMObject -ComObject $s
                    $fp = ''
                    try { $fp = $s.FilePath } catch { }

                    # Safety exclusion: never match the known backup PST path,
                    # even if its bare filename happens to coincide with the
                    # old POP3 store's filename. This avoids detaching a
                    # manually re-attached backup copy by mistake -- only the
                    # true orphaned native POP3 store (a different full path)
                    # should be detached here.
                    if ($fp -and $BackupPSTPath -and ($fp -ieq $BackupPSTPath)) {
                        continue
                    }

                    if ($fp -and [System.IO.Path]::GetFileName($fp) -ieq $pstFileName) {
                        $detachStore  = $s
                        $detachResult = 'FOUND'
                        break
                    }
                }
                catch {
                    Write-OMMigrateLog -Message "Error reading store [$i] during PST detach scan: $_" `
                                       -Level DEBUG
                }
            }

            if ($detachResult -eq 'FOUND' -and $detachStore) {
                try {
                    $rootFolder = $detachStore.GetRootFolder()
                    Register-COMObject -ComObject $rootFolder
                    $namespace.RemoveStore($rootFolder)
                    # Allow Outlook to finish internal MAPI store cleanup before
                    # the caller releases the COM session and calls Quit().
                    # Without this pause, Quit() races the store removal and
                    # causes Outlook to hang until the 15s force-kill fires.
                    Start-Sleep -Seconds 3
                    $detachResult = 'DETACHED'
                    Write-OMMigrateLog -Message "Orphaned PST store detached automatically: $pstFileName" `
                                       -Level INFO
                    Write-AuditEntry  -Action 'PST_STORE_DETACHED' `
                                      -AccountEmail $EmailAddress `
                                      -Detail "Orphaned POP3 PST store detached from nav pane: $pstFileName"
                    Write-Host '  Old POP3 folder detached from folder pane successfully.' `
                               -ForegroundColor Green
                }
                catch {
                    $detachResult = 'FAILED'
                    Write-OMMigrateLog -Message "Failed to detach PST store '$pstFileName': $_" `
                                       -Level WARN
                }
            }
        }
        else {
            $detachResult = 'NO_SESSION'
        }
    }
    catch {
        $detachResult = 'FAILED'
        Write-OMMigrateLog -Message "Exception during PST store detach for '$pstFileName': $_" `
                           -Level WARN
    }

    # If auto-detach did not succeed -- prompt operator to right-click
    # close manually while Outlook is still open
    if ($detachResult -ne 'DETACHED') {
        Write-Host ''
        switch ($detachResult) {
            'NOT_FOUND' {
                Write-Host '  The old POP3 folder could not be located automatically.' `
                           -ForegroundColor Yellow
                Write-Host "  A PST named '$pstFileName' was not found in the mounted stores." `
                           -ForegroundColor Yellow
                Write-Host '  The filename may differ from the expected pattern.' `
                           -ForegroundColor Yellow
                Write-OMMigrateLog -Message "PST store not found for auto-detach -- manual close required: $pstFileName" `
                                   -Level WARN
            }
            'FAILED' {
                Write-Host '  The old POP3 folder could not be detached automatically.' `
                           -ForegroundColor Yellow
                Write-Host '  An error occurred during the detach attempt.' -ForegroundColor Yellow
                Write-OMMigrateLog -Message "PST store detach failed -- manual close required: $pstFileName" `
                                   -Level WARN
            }
            default {
                Write-Host '  The old POP3 folder could not be detached automatically.' `
                           -ForegroundColor Yellow
                Write-OMMigrateLog -Message "PST store detach skipped ($detachResult) -- manual close required: $pstFileName" `
                                   -Level WARN
            }
        }
        Write-Host ''
        Write-Host '  Please close it manually in Outlook now:' -ForegroundColor White
        Write-Host '  1. In the left folder pane -- right-click the old POP3 folder' `
                   -ForegroundColor Gray
        Write-Host "       $safeEmail" -ForegroundColor Cyan
        Write-Host '  2. Select  Close  from the context menu' -ForegroundColor Gray
        Write-Host ''

        Write-AuditEntry -Action 'PST_STORE_DETACH_MANUAL_REQUIRED' `
                         -AccountEmail $EmailAddress `
                         -Detail "Auto-detach result: $detachResult -- operator prompted to close manually."

        do {
            $response = Read-Host '  Type C and press Enter when the old folder has been closed in Outlook'
        } while ($response.Trim().ToUpper() -ne 'C')
    }

    # -- Auto-close Outlook ------------------------------------
    # PST store is now detached (automatically or manually).
    # Script closes Outlook via COM so the operator does not need
    # to do it manually. Script 02 will reopen Outlook automatically
    # before the IMAP add phase begins.
    Write-Host ''
    Write-Host '  Closing Outlook automatically...' -ForegroundColor DarkGray
    Write-OMMigrateLog -Message 'Signaling Script 02 to close Outlook after POP3 removal and PST detach.' `
                       -Level INFO

    return $true
}


function Show-IMAPAccountSetupReference {
    <#
    .SYNOPSIS
        Displays a formatted IMAP Account Setup Reference for
        the operator to use when manually adding the account in Outlook.

    .DESCRIPTION
        IMAP account creation cannot be automated via registry write in
        Outlook 2016/2019/2021 -- direct registry writes bypass MAPI
        store initialization and cause Outlook to crash. The operator
        must add the account through Outlook's own Add Account dialog.

        This function outputs all account details the operator needs to
        fill in that dialog, formatted clearly for non-technical users.
        Passwords are intentionally omitted -- the operator enters them
        directly in Outlook. This script never handles passwords.

        Provider-specific notes are included for AWS SES (IAM key ID
        vs secret key distinction) and AT&T/Ameritech (Secure Mail Key).

        All values are passed through Invoke-OMMigrateSanitize so the
        output is safe when -Sanitize is active.

    .PARAMETER AccountConfig
        PSCustomObject from migration_accounts.csv with these properties:
            EmailAddress, DisplayName, NewImapServer, NewImapPort,
            NewImapSSL, OutgoingServer, OutgoingPort, OutgoingSSL,
            ImapUsername, SmtpUsername, ProviderTag

    .EXAMPLE
        Show-IMAPAccountSetupReference -AccountConfig $accountConfig
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$AccountConfig
    )

    # -- Resolve display values --------------------------------
    # All values sanitized for safe console output when -Sanitize active
    $email       = Invoke-OMMigrateSanitize -Text $AccountConfig.EmailAddress
    $displayName = Invoke-OMMigrateSanitize -Text $AccountConfig.DisplayName
    $imapServer  = Invoke-OMMigrateSanitize -Text $AccountConfig.NewImapServer
    $imapPort    = $AccountConfig.NewImapPort
    $imapSSL     = $AccountConfig.NewImapSSL
    $smtpServer  = Invoke-OMMigrateSanitize -Text $AccountConfig.OutgoingServer
    $smtpPort    = $AccountConfig.OutgoingPort
    $smtpSSL     = $AccountConfig.OutgoingSSL

    # Resolve IMAP/SMTP usernames -- default to email if not set
    # AWS SES accounts use an IAM SMTP key ID as the SMTP username
    # Pre-resolve to variables first -- PS5 does not allow inline if
    # as a parameter value inside a function call parentheses.
    $imapUserRaw = if ($AccountConfig.PSObject.Properties['ImapUsername'] -and $AccountConfig.ImapUsername) { $AccountConfig.ImapUsername } else { $AccountConfig.EmailAddress }
    $smtpUserRaw = if ($AccountConfig.PSObject.Properties['SmtpUsername'] -and $AccountConfig.SmtpUsername) { $AccountConfig.SmtpUsername } else { $AccountConfig.EmailAddress }
    $imapUser = Invoke-OMMigrateSanitize -Text $imapUserRaw
    $smtpUser = Invoke-OMMigrateSanitize -Text $smtpUserRaw

    # -- Compute human-readable encryption labels --------------
    # Outlook's Add Account dialog uses these exact labels
    $imapEncryption = if ($imapSSL) { 'SSL/TLS' } else { 'None' }
    $smtpEncryption = if ($smtpSSL) {
        if ($smtpPort -eq 465) { 'SSL/TLS' } else { 'STARTTLS' }
    } else { 'None' }

    # -- Provider tag for provider-specific notes --------------
    $providerTag = ''
    try { $providerTag = $AccountConfig.ProviderTag } catch { }

    # -- Output Account Setup Reference ------------------------------------
    Write-Host ''
    Write-Host ('  ' + ('=' * 54)) -ForegroundColor Cyan
    Write-Host '  IMAP ACCOUNT SETUP -- USE THESE DETAILS IN OUTLOOK' -ForegroundColor Cyan
    Write-Host ('  ' + ('=' * 54)) -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  Keep this window visible while adding the account.' -ForegroundColor White
    Write-Host ''

    # -- Account identity --------------------------------------
    Write-Host '  Account Details:' -ForegroundColor White
    Write-Host "    Your Name     :  $displayName" -ForegroundColor Green
    Write-Host "    Email Address :  $email" -ForegroundColor Green
    Write-Host ''

    # -- Incoming mail (IMAP) ----------------------------------
    Write-Host '  Incoming Mail Server (IMAP):' -ForegroundColor White
    Write-Host "    Server        :  $imapServer" -ForegroundColor Green
    Write-Host "    Port          :  $imapPort" -ForegroundColor Green
    Write-Host "    Encryption    :  $imapEncryption" -ForegroundColor Green
    Write-Host "    Username      :  $imapUser" -ForegroundColor Green
    Write-Host "    Password      :  [enter your email password in Outlook]" -ForegroundColor Yellow
    Write-Host ''

    # -- Outgoing mail (SMTP) ----------------------------------
    Write-Host '  Outgoing Mail Server (SMTP):' -ForegroundColor White
    Write-Host "    Server        :  $smtpServer" -ForegroundColor Green
    Write-Host "    Port          :  $smtpPort" -ForegroundColor Green
    Write-Host "    Encryption    :  $smtpEncryption" -ForegroundColor Green
    Write-Host "    Requires Auth :  Yes" -ForegroundColor Green
    Write-Host "    Username      :  $smtpUser" -ForegroundColor Green

    # Provider-specific SMTP password notes
    if ($providerTag -eq 'POP3-AWS') {
        # AWS SES uses IAM SMTP credentials -- not the email password or AWS console password.
        # The IAM SMTP key ID is shown above as the SMTP username.
        # The IAM SMTP secret access key is the SMTP password -- it must be entered in Outlook.
        Write-Host "    Password      :  [enter your AWS IAM SMTP secret access key]" `
                   -ForegroundColor Yellow
    }
    else {
        Write-Host "    Password      :  [enter your email password in Outlook]" -ForegroundColor Yellow
    }

    Write-Host ''

    # -- Provider-specific notes -------------------------------
    if ($providerTag -eq 'POP3-AWS') {
        Write-Host ('  ' + ('-' * 54)) -ForegroundColor Yellow
        Write-Host '  AWS SES OUTBOUND NOTE:' -ForegroundColor Yellow
        Write-Host '  The SMTP Username shown above is your AWS IAM SMTP' -ForegroundColor Gray
        Write-Host '  access key ID -- it looks like: AKIAIOSFODNN7EXAMPLE' -ForegroundColor Gray
        Write-Host '  The SMTP Password is your IAM SMTP secret access key.' -ForegroundColor Gray
        Write-Host '  These are NOT your AWS console login credentials.' -ForegroundColor Yellow
        Write-Host '  Find both values in migration_accounts.csv.' -ForegroundColor Gray
        Write-Host ('  ' + ('-' * 54)) -ForegroundColor Yellow
        Write-Host ''
    }
    elseif ($providerTag -eq 'POP3-ATTAMERITECH') {
        Write-Host ('  ' + ('-' * 54)) -ForegroundColor Yellow
        Write-Host '  AT&T / AMERITECH NOTE:' -ForegroundColor Yellow
        Write-Host '  Your password is a Secure Mail Key -- NOT your regular' -ForegroundColor Gray
        Write-Host '  AT&T email password. Generate one at: currently.com' -ForegroundColor Cyan
        Write-Host '  Sign in > My Profile > Secure Mail Key > Manage' -ForegroundColor Gray
        Write-Host ('  ' + ('-' * 54)) -ForegroundColor Yellow
        Write-Host ''
    }

    Write-Host ('  ' + ('=' * 54)) -ForegroundColor Cyan
    Write-Host ''
}


function Add-IMAPAccount {
    <#
    .SYNOPSIS
        Guides the operator through manually adding a new IMAP account
        in Outlook using Outlook's own Add Account dialog.

    .DESCRIPTION
        IMAP account creation cannot be automated via registry write in
        Outlook 2016/2019/2021. Direct registry writes bypass MAPI store
        initialization, OST file creation, and search folder setup --
        Outlook crashes on load when the account is added this way.

        This function:
          1. Calls Add-IMAPAccountViaRegistry to log the operation and
             confirm the account parameters are valid (no registry write
             occurs -- see that function for the full explanation).
          2. Displays a formatted Account Setup Reference via Show-IMAPAccountSetupReference
             with all server details the operator needs.
          3. Provides clear step-by-step instructions for Outlook's
             File > Add Account > Manual setup > IMAP flow.
          4. Pauses and waits for the operator to confirm the account
             is visible and connected in Outlook before returning.

        Passwords are NEVER handled by this script. The operator enters
        them directly in Outlook's credential dialogs.

    .PARAMETER AccountConfig
        PSCustomObject from migration_accounts.csv with these properties:
            EmailAddress, DisplayName, NewImapServer, NewImapPort,
            NewImapSSL, OutgoingServer, OutgoingPort, OutgoingSSL,
            ImapUsername, SmtpUsername, ProviderTag

    .OUTPUTS
        [bool] -- $true if operator confirmed account added successfully,
                  $false if the delegation call failed unexpectedly.

    .EXAMPLE
        $config = [PSCustomObject]@{
            EmailAddress   = 'user@domain.com'
            DisplayName    = 'User Name'
            NewImapServer  = 'mail.domain.com'
            NewImapPort    = 993
            NewImapSSL     = $true
            OutgoingServer = 'mail.domain.com'
            OutgoingPort   = 587
            OutgoingSSL    = $true
            ProviderTag    = 'POP3-STANDARD'
        }
        $added = Add-IMAPAccount -AccountConfig $config
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$AccountConfig
    )

    $email = $AccountConfig.EmailAddress

    # Log the operation at the Outlook module level
    Write-OMMigrateLog -Message "Adding IMAP account (guided manual): $email" -Level INFO
    Write-OMMigrateLog -Message "  IMAP Server : $($AccountConfig.NewImapServer):$($AccountConfig.NewImapPort)" `
                       -Level INFO
    Write-OMMigrateLog -Message "  SMTP Server : $($AccountConfig.OutgoingServer):$($AccountConfig.OutgoingPort)" `
                       -Level INFO

    # -- Call Add-IMAPAccountViaRegistry -----------------------
    # No registry write occurs -- the function logs the WhatIf/audit
    # entry and returns $true immediately. The actual account creation
    # is delegated to Outlook's Add Account dialog below.
    # Resolve IMAP/SMTP usernames before the call -- inline if is not
    # valid as a parameter value in PS5. Default to email if not set.
    $imapUser = if ($AccountConfig.PSObject.Properties['ImapUsername'] -and $AccountConfig.ImapUsername) { $AccountConfig.ImapUsername } else { $email }
    $smtpUser = if ($AccountConfig.PSObject.Properties['SmtpUsername'] -and $AccountConfig.SmtpUsername) { $AccountConfig.SmtpUsername } else { $email }

    $delegated = Add-IMAPAccountViaRegistry `
        -EmailAddress  $email `
        -DisplayName   $AccountConfig.DisplayName `
        -ImapServer    $AccountConfig.NewImapServer `
        -ImapPort      $AccountConfig.NewImapPort `
        -ImapSSL       $AccountConfig.NewImapSSL `
        -ImapUsername  $imapUser `
        -SmtpServer    $AccountConfig.OutgoingServer `
        -SmtpPort      $AccountConfig.OutgoingPort `
        -SmtpSSL       $AccountConfig.OutgoingSSL `
        -SmtpUsername  $smtpUser

    if (-not $delegated) {
        # Should not happen -- Add-IMAPAccountViaRegistry returns $true
        # unconditionally in the live path. Guard defensively.
        Write-OMMigrateLog -Message "Unexpected failure in Add-IMAPAccountViaRegistry for: $email" `
                           -Level ERROR
        Write-AuditEntry  -Action 'ACCOUNT_ADD_FAILED' `
                          -AccountEmail $email `
                          -Detail "Add-IMAPAccountViaRegistry returned false unexpectedly." `
                          -Outcome 'FAILED'
        return $false
    }

    # WhatIf -- Account Setup Reference and operator prompt are skipped in preview mode.
    # Add-IMAPAccountViaRegistry already logged the WhatIf details.
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message "WhatIf: Would display IMAP Account Setup Reference and guide operator through manual add" `
                           -Level INFO -WhatIfPrefix
        return $true
    }

    # -- Display Account Setup Reference -----------------------------------
    # Show all account details the operator needs for the manual add.
    # Passwords are intentionally omitted -- operator enters in Outlook.
    Show-IMAPAccountSetupReference -AccountConfig $AccountConfig

    # -- Step-by-step operator instructions --------------------
    Write-Host ('  ' + ('-' * 54)) -ForegroundColor Yellow
    Write-Host "  ACTION REQUIRED -- Add $(Invoke-OMMigrateSanitize -Text $email) to Outlook" `
               -ForegroundColor Yellow
    Write-Host ('  ' + ('-' * 54)) -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  Follow these steps in Outlook now:' -ForegroundColor White
    Write-Host ''
    Write-Host '  1. Click  File  (top left)' -ForegroundColor Gray
    Write-Host '  2. Click  Add Account' -ForegroundColor Gray
    Write-Host '  3. Enter your email address and click  Connect' -ForegroundColor Gray
    Write-Host '  4. If Outlook auto-configures -- great, skip to step 7' -ForegroundColor Gray
    Write-Host '  5. If prompted for account type -- choose  IMAP' -ForegroundColor Cyan
    Write-Host '  6. Enter the server details from the Account Setup Reference above' -ForegroundColor Gray
    Write-Host '  7. Enter your password when Outlook prompts you' -ForegroundColor Gray
    Write-Host '  8. Click  Done  or  Finish  when the account is added' -ForegroundColor Gray
    Write-Host ''
    Write-Host '  When you see the new account appear in the left folder pane:' -ForegroundColor White
    Write-Host '  Do NOT attempt to send or receive mail yet.' -ForegroundColor Yellow
    Write-Host '  Credential corrections are required after this script completes.' `
               -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  Refer to the Account Setup Reference above for all server details.' -ForegroundColor White
    Write-Host ('  ' + ('-' * 54)) -ForegroundColor Yellow
    Write-Host ''

    # Require explicit 'C' to continue -- prevents accidental Enter
    # from advancing the script while the operator is still working in Outlook.
    # Script will close Outlook automatically after this prompt.
    # The operator must deliberately type C and press Enter to proceed.
    do {
        $response = Read-Host '  Type C and press Enter when the new account is visible in the folder pane'
    } while ($response.Trim().ToUpper() -ne 'C')

    Write-AuditEntry -Action 'ACCOUNT_ADDED_IMAP' `
                     -AccountEmail $email `
                     -Detail ("IMAP account added via Outlook Add Account dialog (guided manual). " +
                              "Server=$($AccountConfig.NewImapServer):$($AccountConfig.NewImapPort) | " +
                              "SMTP=$($AccountConfig.OutgoingServer):$($AccountConfig.OutgoingPort)")

    return $true
}


# ============================================================
#  HELPER: Get-OrCreateFolder
#  Returns an Outlook MAPIFolder by name under the given parent,
#  creating it if it does not already exist.
# ============================================================

function Get-OrCreateFolder {
    <#
    .SYNOPSIS
        Returns an Outlook MAPIFolder by name under the given parent,
        creating it if it does not already exist.

    .PARAMETER ParentFolder
        The Outlook MAPIFolder COM object to search/create under.

    .PARAMETER FolderName
        Name of the folder to find or create.

    .OUTPUTS
        [System.__ComObject] -- MAPIFolder object, or $null on failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$ParentFolder,

        [Parameter(Mandatory = $true)]
        [string]$FolderName
    )

    # Search existing subfolders first
    try {
        $subfolders = $ParentFolder.Folders
        for ($i = 1; $i -le $subfolders.Count; $i++) {
            $sub = $subfolders.Item($i)
            if ($sub.Name -eq $FolderName) {
                Register-COMObject -ComObject $sub
                return $sub
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Error searching subfolders of '$($ParentFolder.Name)': $_" `
                           -Level WARN
    }

    # Not found -- create it
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message "WhatIf: Would create folder '$FolderName' under '$($ParentFolder.Name)'" `
                           -Level INFO -WhatIfPrefix
        return $null
    }

    try {
        $newFolder = $ParentFolder.Folders.Add($FolderName)
        Register-COMObject -ComObject $newFolder
        Write-OMMigrateLog -Message "Created folder: '$FolderName' under '$($ParentFolder.Name)'" `
                           -Level DEBUG
        return $newFolder
    }
    catch {
        Write-OMMigrateLog -Message "Failed to create folder '$FolderName': $_" -Level WARN
        return $null
    }
}


# ============================================================
#  HELPER: Get-FolderByPath
#  Navigates a folder path string to return the COM folder object.
# ============================================================

function Get-FolderByPath {
    <#
    .SYNOPSIS
        Navigates a backslash-delimited folder path from a root store
        and returns the target MAPIFolder COM object.

    .PARAMETER RootFolder
        The root MAPIFolder of the store to navigate from.

    .PARAMETER FolderPath
        Backslash-delimited path relative to the store root.
        Example: 'Inbox\Vendors\Microsoft'

    .PARAMETER CreateIfMissing
        When $true, creates any missing intermediate folders.
        Default: $false

    .OUTPUTS
        [System.__ComObject] -- The target MAPIFolder, or $null if not found.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$RootFolder,

        [Parameter(Mandatory = $true)]
        [string]$FolderPath,

        [Parameter(Mandatory = $false)]
        [bool]$CreateIfMissing = $false
    )

    $parts   = $FolderPath.Split('\') | Where-Object { $_ -ne '' }
    $current = $RootFolder

    foreach ($part in $parts) {
        $found = $false
        try {
            $subs = $current.Folders
            for ($i = 1; $i -le $subs.Count; $i++) {
                $sub = $subs.Item($i)
                Register-COMObject -ComObject $sub
                if ($sub.Name -eq $part) {
                    $current = $sub
                    $found   = $true
                    break
                }
            }
        }
        catch { }

        if (-not $found) {
            if ($CreateIfMissing) {
                $current = Get-OrCreateFolder -ParentFolder $current -FolderName $part
                if (-not $current) { return $null }
            }
            else {
                return $null
            }
        }
    }

    return $current
}


function Test-PSTAlreadyMounted {
    <#
    .SYNOPSIS
        Checks whether a PST file is already mounted as a store in the
        active Outlook profile, without mounting or unmounting anything.

    .DESCRIPTION
        Read-only check used to distinguish "this script just mounted this
        PST for its own temporary use" from "this PST was already attached
        to the profile before this script touched it" (e.g. Administrator manually
        attaching a backup/osttoimap PST as a permanent store for manual
        email triage). Callers use this BEFORE calling Open-PSTFile so they
        can decide later whether it is safe to detach the store when done.

        Performs the same FilePath scan Open-PSTFile already does
        internally, but does not call AddStore and does not modify the
        profile in any way.

    .PARAMETER PSTPath
        Full path to the PST file to check.

    .OUTPUTS
        [bool] -- $true if a store with this FilePath is currently mounted,
        $false if not mounted or if the check could not be completed.

    .EXAMPLE
        $alreadyMounted = Test-PSTAlreadyMounted -PSTPath $backupPath
        $backupStore    = Open-PSTFile -PSTPath $backupPath -DisplayName '...'
        # ... later ...
        if ($backupStore -and -not $alreadyMounted) {
            Close-PSTFile -PSTPath $backupPath | Out-Null
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PSTPath
    )

    $namespace = Get-OutlookNamespace
    if (-not $namespace) { return $false }

    try {
        $stores = $namespace.Stores
        Register-COMObject -ComObject $stores

        for ($i = 1; $i -le $stores.Count; $i++) {
            $s = $stores.Item($i)
            Register-COMObject -ComObject $s
            try {
                if ($s.FilePath -eq $PSTPath) {
                    return $true
                }
            }
            catch { }
        }

        return $false
    }
    catch {
        Write-OMMigrateLog -Message "Test-PSTAlreadyMounted: Error checking '$PSTPath': $_" -Level DEBUG
        return $false
    }
}


function Open-PSTFile {
    <#
    .SYNOPSIS
        Opens a PST file in the active Outlook COM session.

    .DESCRIPTION
        Attaches a PST file to the current Outlook profile as a
        data store. Used by Script 03 to attach backup PST files
        for folder content migration.

        Returns the Store object for the opened PST.

        Requires an active COM session from Connect-OutlookCOM.

    .PARAMETER PSTPath
        Full path to the PST file to open.

    .PARAMETER DisplayName
        Optional display name for the PST in the Outlook folder pane.

    .OUTPUTS
        [System.__ComObject] -- The opened Store object, or $null on failure.

    .EXAMPLE
        $store = Open-PSTFile -PSTPath 'C:\Backups\user.pst' `
                              -DisplayName 'Backup - user@domain.com'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PSTPath,

        [Parameter(Mandatory = $false)]
        [string]$DisplayName = ''
    )

    if (-not (Test-Path $PSTPath)) {
        Write-OMMigrateLog -Message "PST file not found: $PSTPath" -Level ERROR
        return $null
    }

    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message "WhatIf: Would open PST: $PSTPath" `
                           -Level INFO -WhatIfPrefix
        return $null
    }

    $namespace = Get-OutlookNamespace
    if (-not $namespace) { return $null }

    Write-OMMigrateLog -Message "Opening PST file: $PSTPath" -Level INFO

    try {
        # Check if this PST is already mounted in the active Outlook profile.
        # This happens when the Archive PST was left mounted from a previous run
        # (by design -- so it appears in Outlook automatically). Calling AddStore
        # on an already-mounted PST opens a second handle and changes made through
        # it are not visible through the existing profile handle until Outlook
        # restarts. Instead, return the existing store reference directly.
        $stores = $namespace.Stores
        Register-COMObject -ComObject $stores

        $existingStore = $null
        for ($i = 1; $i -le $stores.Count; $i++) {
            $s = $stores.Item($i)
            Register-COMObject -ComObject $s
            try {
                if ($s.FilePath -eq $PSTPath) {
                    $existingStore = $s
                    break
                }
            }
            catch { }
        }

        if ($existingStore) {
            Write-OMMigrateLog -Message "PST already mounted -- using existing store: $PSTPath" `
                               -Level DEBUG
            Write-AuditEntry  -Action 'PST_OPENED' `
                              -Detail "Path=$PSTPath | DisplayName=$DisplayName | AlreadyMounted=True"
            return $existingStore
        }

        # PST not yet mounted -- add it now
        $namespace.AddStore($PSTPath)

        # Find the newly added store
        Start-Sleep -Milliseconds 500   # Allow Outlook to process the AddStore

        $stores = $namespace.Stores
        Register-COMObject -ComObject $stores

        $newStore = $null
        for ($i = 1; $i -le $stores.Count; $i++) {
            $s = $stores.Item($i)
            Register-COMObject -ComObject $s
            try {
                if ($s.FilePath -eq $PSTPath) {
                    $newStore = $s
                    break
                }
            }
            catch { }
        }

        if ($newStore -and $DisplayName) {
            try {
                $rootFolder      = $newStore.GetRootFolder()
                $rootFolder.Name = $DisplayName
            }
            catch { }
        }

        Write-OMMigrateLog -Message "PST opened successfully: $PSTPath" -Level INFO
        Write-AuditEntry  -Action 'PST_OPENED' `
                          -Detail "Path=$PSTPath | DisplayName=$DisplayName"

        return $newStore
    }
    catch {
        Write-OMMigrateLog -Message "Failed to open PST '$PSTPath': $_" -Level ERROR
        return $null
    }
}


function Close-PSTFile {
    <#
    .SYNOPSIS
        Detaches a PST file from the active Outlook COM session.

    .DESCRIPTION
        Removes a PST store from the current Outlook profile.
        Used by Script 03 after folder migration is complete to
        detach the backup PST from the live profile.

        The PST file itself is not deleted -- it remains on disk.

    .PARAMETER PSTPath
        Full path to the PST file to close/detach.

    .OUTPUTS
        [bool] -- $true if detached successfully. Retries up to 3 total
        attempts before returning $false -- see 2026-07-20 fix comment below.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PSTPath
    )

    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message "WhatIf: Would close PST: $PSTPath" `
                           -Level INFO -WhatIfPrefix
        return $true
    }

    $namespace = Get-OutlookNamespace
    if (-not $namespace) { return $false }

    Write-OMMigrateLog -Message "Closing PST: $PSTPath" -Level INFO

    # ADDED (Administrator direction, 2026-07-20): retry loop. Administrator hit a
    # real case where this close failed once on a COM null error (source PST left
    # unstable after a folder-walk failure elsewhere -- see the Archive pre-build
    # consecutive-failure guard added the same session) and the PST was left
    # attached in Outlook, requiring a manual close. A single transient COM
    # failure does not necessarily mean the store is permanently unclosable, so
    # this now retries up to 2 additional times (3 attempts total) with a short
    # pause between attempts before giving up. The caller is still responsible
    # for telling the operator to close manually if all attempts fail -- this
    # function only reports success/failure via its existing [bool] return.
    $maxCloseAttempts = 3
    for ($closeAttempt = 1; $closeAttempt -le $maxCloseAttempts; $closeAttempt++) {
        try {
            $stores = $namespace.Stores
            Register-COMObject -ComObject $stores

            for ($i = 1; $i -le $stores.Count; $i++) {
                $store = $stores.Item($i)
                Register-COMObject -ComObject $store
                try {
                    if ($store.FilePath -eq $PSTPath) {
                        $namespace.RemoveStore($store.GetRootFolder())
                        Write-OMMigrateLog -Message "PST detached: $PSTPath (attempt $closeAttempt of $maxCloseAttempts)" -Level INFO
                        Write-AuditEntry  -Action 'PST_CLOSED' -Detail "Path=$PSTPath"
                        return $true
                    }
                }
                catch { }
            }

            Write-OMMigrateLog -Message "PST not found in active stores: $PSTPath (attempt $closeAttempt of $maxCloseAttempts)" -Level WARN
            if ($closeAttempt -ge $maxCloseAttempts) { return $false }
        }
        catch {
            Write-OMMigrateLog -Message "Failed to close PST '$PSTPath' (attempt $closeAttempt of $maxCloseAttempts): $_" -Level WARN
            if ($closeAttempt -ge $maxCloseAttempts) {
                Write-OMMigrateLog -Message "Failed to close PST '$PSTPath' after $maxCloseAttempts attempts -- giving up." -Level ERROR
                return $false
            }
        }
        Start-Sleep -Milliseconds 750
    }

    return $false
}


# ============================================================
#  REGION: MODULE EXPORTS
# ============================================================

# ============================================================
#  REGION: FOLDER MAP PICKER
# ============================================================

function Invoke-FolderMapPicker {
    <#
    .SYNOPSIS
        Displays a WinForms popup letting the operator assign Server,
        Local, or Server destinations to each folder in folder_map.csv.

    .DESCRIPTION
        Called by Script 01 after Archive pre-build and folder_map.csv update.
        Shows all folders grouped by account (StoreName) in a DataGridView.
        The Destination column is a dropdown cell with Server / Local
        options. The operator reviews and changes assignments as needed, then
        clicks OK to save. Cancel leaves the CSV untouched.

        Pre-populates each row with the Destination already written by
        Export-FolderMapCSV (intelligent defaults: standard top-level mail
        folders = Server, all others = Local).

        Skipped automatically in WhatIf/Preview mode.

        Uses System.Windows.Forms -- built into PowerShell 5.1 on Windows.
        No external dependencies.

    .PARAMETER CsvPath
        Full path to folder_map.csv. Must exist before calling.

    .EXAMPLE
        Invoke-FolderMapPicker -CsvPath 'C:\...\Config\folder_map.csv'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,

        [Parameter(Mandatory = $false)]
        [string]$FilterStoreName = ''
        # When specified, only folders belonging to this store are shown.
        # Used by Script 00 per-account loop to scope the picker to one
        # account at a time. When empty, all folders are shown.
    )

    # Skip entirely in WhatIf mode
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message 'WhatIf: Would display folder map picker -- skipped.' `
                           -Level INFO -WhatIfPrefix
        return
    }

    if (-not (Test-Path $CsvPath)) {
        Write-OMMigrateLog -Message "Invoke-FolderMapPicker: folder_map.csv not found at $CsvPath -- skipped." `
                           -Level WARN
        return
    }

    # Load CSV rows
    # When filtering to a specific account, show all that account's folders
    # regardless of StoreType -- the account being configured may still be
    # a PST (POP3) store that hasn't been converted to IMAP yet.
    # When showing all accounts (no filter), only show OST stores since
    # PST stores are not migration candidates in the global view.
    if ($FilterStoreName -and -not [string]::IsNullOrWhiteSpace($FilterStoreName)) {
        $rows = @(
            Import-Csv -Path $CsvPath -Encoding UTF8 |
            Where-Object {
                $_.StoreName -eq $FilterStoreName -and
                -not [string]::IsNullOrWhiteSpace($_.StoreName) -and
                -not [string]::IsNullOrWhiteSpace($_.FolderPath) -and
                $_.IsSystemFolder -ne 'True'
            }
        )
        Write-OMMigrateLog -Message "Invoke-FolderMapPicker: Filtered to store '$FilterStoreName' -- $($rows.Count) folder(s)." `
                           -Level INFO
    }
    else {
        # No filter -- show OST stores and configured archive PST(s).
        # OST = live IMAP accounts, Archive PST = Local destination folders.
        #
        # FIXED (2026-07-09, Administrator direction): previously hardcoded to a
        # single specific archive store display name, which (a) baked a
        # user's personal store name into shared code, and (b) as of this
        # session's TargetStoreName feature, would silently exclude any
        # OTHER archive PST an operator has mapped accounts to via
        # RulesEngine.ArchiveStoreMappings -- the tool now genuinely
        # supports multiple archive destinations, so a single-name check
        # is not just a data-hygiene issue but a real functional gap.
        # Fixed to check membership against the full set of configured
        # archive store names (every distinct TargetStoreName across all
        # ArchiveStoreMappings entries) instead of one literal string.
        # $configuredArchiveStoreNames is empty when no mapping has been
        # set up yet, in which case this condition simply never matches
        # any row via this branch -- OST rows still show normally, and
        # StoreType-only PST rows (e.g. a not-yet-mapped archive) are
        # left out of the unfiltered view exactly as before this fix,
        # same as the original hardcoded-name behavior would do for any
        # PST that didn't match the one literal name.
        $configuredArchiveStoreNames = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )
        try {
            if ($Global:OMMigrate.Settings.PSObject.Properties['RulesEngine'] -and
                $Global:OMMigrate.Settings.RulesEngine.PSObject.Properties['ArchiveStoreMappings']) {
                foreach ($archMapping in @($Global:OMMigrate.Settings.RulesEngine.ArchiveStoreMappings)) {
                    if ($archMapping.PSObject.Properties['TargetStoreName'] -and
                        -not [string]::IsNullOrWhiteSpace($archMapping.TargetStoreName)) {
                        [void]$configuredArchiveStoreNames.Add($archMapping.TargetStoreName)
                    }
                }
            }
        }
        catch { }

        $rows = @(
            Import-Csv -Path $CsvPath -Encoding UTF8 |
            Where-Object {
                $_.StoreName -and $_.FolderPath -and
                -not [string]::IsNullOrWhiteSpace($_.StoreName) -and
                -not [string]::IsNullOrWhiteSpace($_.FolderPath) -and
                (
                    ($_.PSObject.Properties['StoreType'] -and $_.StoreType -eq 'OST') -or
                    $configuredArchiveStoreNames.Contains($_.StoreName)
                ) -and
                $_.IsSystemFolder -ne 'True'
            }
        )
    }

    if ($rows.Count -eq 0) {
        Write-OMMigrateLog -Message 'Invoke-FolderMapPicker: No folders in CSV -- picker skipped.' `
                           -Level INFO
        return
    }

    Write-OMMigrateLog -Message "Invoke-FolderMapPicker: Displaying picker for $($rows.Count) folder(s)." `
                       -Level INFO

    # Load WinForms assembly
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing       -ErrorAction Stop
    }
    catch {
        Write-OMMigrateLog -Message "Could not load WinForms -- folder map picker unavailable: $_" -Level WARN
        Write-Host '  WARNING: Could not open folder picker (WinForms unavailable).' -ForegroundColor Yellow
        Write-Host '  Open Config\folder_map.csv and set the Destination column manually.' `
                   -ForegroundColor Yellow
        return
    }

    # -- Build the form ----------------------------------------
    $form                  = New-Object System.Windows.Forms.Form
    $form.Text             = if ($FilterStoreName) {
        "OMMigrate -- Folder Destinations: $FilterStoreName"
    } else {
        'OMMigrate -- Folder Destination Assignments'
    }
    $form.Size             = New-Object System.Drawing.Size(780, 560)
    $form.StartPosition    = 'CenterScreen'
    $form.FormBorderStyle  = 'Sizable'
    $form.MaximizeBox      = $true
    $form.MinimizeBox      = $false
    $form.MinimumSize      = New-Object System.Drawing.Size(620, 420)
    $form.TopMost          = $true

    # Instruction label
    $label             = New-Object System.Windows.Forms.Label
    $label.Location    = New-Object System.Drawing.Point(12, 10)
    $label.Size        = New-Object System.Drawing.Size(740, 36)
    $label.Text        = 'Review folder destinations. Click the Destination cell to change: Server = all devices & webmail  |  Local = this desktop only'
    $label.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.Controls.Add($label)

    # DataGridView
    $grid                          = New-Object System.Windows.Forms.DataGridView
    $grid.Location                 = New-Object System.Drawing.Point(12, 52)
    $grid.Size                     = New-Object System.Drawing.Size(740, 420)
    $grid.Anchor                   = 'Top,Bottom,Left,Right'
    $grid.AllowUserToAddRows       = $false
    $grid.AllowUserToDeleteRows    = $false
    $grid.AllowUserToResizeRows    = $false
    $grid.RowHeadersVisible        = $false
    $grid.SelectionMode            = 'FullRowSelect'
    $grid.MultiSelect              = $false
    $grid.AutoSizeColumnsMode      = 'None'
    $grid.Font                     = New-Object System.Drawing.Font('Segoe UI', 9)
    $grid.BackgroundColor          = [System.Drawing.SystemColors]::Window
    $grid.BorderStyle              = 'Fixed3D'
    $grid.EditMode                 = 'EditOnEnter'

    # Define columns
    # StoreName column
    $colStore                = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStore.HeaderText     = 'Account'
    $colStore.Name           = 'StoreName'
    $colStore.Width          = 180
    $colStore.ReadOnly       = $true
    $grid.Columns.Add($colStore) | Out-Null

    # FolderPath column
    $colPath                 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPath.HeaderText      = 'Folder Path'
    $colPath.Name            = 'FolderPath'
    $colPath.Width           = 260
    $colPath.ReadOnly        = $true
    $grid.Columns.Add($colPath) | Out-Null

    # Destination column -- ComboBox dropdown
    $colDest                 = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $colDest.HeaderText      = 'Destination'
    $colDest.Name            = 'Destination'
    $colDest.Width           = 90
    $colDest.Items.AddRange(@('Server','Local'))
    $colDest.ReadOnly        = $false
    $grid.Columns.Add($colDest) | Out-Null

    # ItemCount column
    $colItems                = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colItems.HeaderText     = 'Items'
    $colItems.Name           = 'ItemCount'
    $colItems.Width          = 60
    $colItems.ReadOnly       = $true
    $colItems.DefaultCellStyle.Alignment = 'MiddleRight'
    $grid.Columns.Add($colItems) | Out-Null

    # Notes column
    $colNotes                = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colNotes.HeaderText     = 'Notes'
    $colNotes.Name           = 'Notes'
    $colNotes.Width          = 120
    $colNotes.ReadOnly       = $true
    $grid.Columns.Add($colNotes) | Out-Null

    # Populate rows -- group by StoreName with alternating background per account
    $storeColors  = @('#F0F4FF','#FFFFFF')   # Light blue / white alternating per account
    $storeIndex   = -1
    $lastStore    = ''

    foreach ($row in $rows) {
        # WinForms popups always show real values -- sanitize is console-only
        $storeName   = $row.StoreName
        $folderPath  = $row.FolderPath

        $destination = $row.Destination
        # Validate destination value -- default Local if unexpected
        if ($destination -notin @('Server','Local')) { $destination = 'Local' }

        $itemCount   = $row.ItemCount
        $notes       = $row.Notes

        # Alternate background color per store group
        if ($storeName -ne $lastStore) {
            $storeIndex++
            $lastStore = $storeName
        }
        $bgHex   = $storeColors[$storeIndex % 2]
        $bgColor = [System.Drawing.ColorTranslator]::FromHtml($bgHex)

        $rowIndex = $grid.Rows.Add($storeName, $folderPath, $destination, $itemCount, $notes)
        $gridRow  = $grid.Rows[$rowIndex]
        $gridRow.DefaultCellStyle.BackColor = $bgColor

        # Lock Destination cell for system folders -- they should never be
        # changed by the operator. Gray background signals read-only.
        $isSystem = ($row.IsSystemFolder -eq 'True' -or $row.IsSystemFolder -eq $true)
        if ($isSystem) {
            $gridRow.Cells['Destination'].ReadOnly = $true
            $gridRow.DefaultCellStyle.BackColor    = [System.Drawing.Color]::FromArgb(235, 235, 235)
            $gridRow.DefaultCellStyle.ForeColor    = [System.Drawing.Color]::Gray
        }
        elseif ($destination -eq 'Local') {
            # Green Destination cell highlights Local rows for easy scanning
            $gridRow.Cells['Destination'].Style.BackColor = [System.Drawing.Color]::FromArgb(198, 239, 206)
            $gridRow.Cells['Destination'].Style.ForeColor = [System.Drawing.Color]::FromArgb(0, 97, 0)
        }
    }

    $form.Controls.Add($grid)

    # Update Destination cell color when operator changes the value
    $grid.Add_CellValueChanged({
        param($s, $e)
        if ($e.ColumnIndex -ge 0 -and $grid.Columns[$e.ColumnIndex].Name -eq 'Destination') {
            $cell = $grid.Rows[$e.RowIndex].Cells['Destination']
            if ($cell.ReadOnly) { return }
            if ($cell.Value -eq 'Local') {
                $cell.Style.BackColor = [System.Drawing.Color]::FromArgb(198, 239, 206)
                $cell.Style.ForeColor = [System.Drawing.Color]::FromArgb(0, 97, 0)
            }
            else {
                $cell.Style.BackColor = [System.Drawing.Color]::Empty
                $cell.Style.ForeColor = [System.Drawing.Color]::Empty
            }
        }
    })

    # Select All Server / Select All Local buttons
    $btnAllServer          = New-Object System.Windows.Forms.Button
    $btnAllServer.Location = New-Object System.Drawing.Point(12, 480)
    $btnAllServer.Size     = New-Object System.Drawing.Size(110, 28)
    $btnAllServer.Text     = 'All Server'
    $btnAllServer.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnAllServer.Anchor   = 'Bottom,Left'
    $btnAllServer.Add_Click({
        for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
            $cell = $grid.Rows[$i].Cells['Destination']
            if (-not $cell.ReadOnly) {
                $cell.Value            = 'Server'
                $cell.Style.BackColor  = [System.Drawing.Color]::Empty
                $cell.Style.ForeColor  = [System.Drawing.Color]::Empty
            }
        }
    })
    $form.Controls.Add($btnAllServer)

    $btnAllLocal           = New-Object System.Windows.Forms.Button
    $btnAllLocal.Location  = New-Object System.Drawing.Point(130, 480)
    $btnAllLocal.Size      = New-Object System.Drawing.Size(110, 28)
    $btnAllLocal.Text      = 'All Local'
    $btnAllLocal.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnAllLocal.Anchor    = 'Bottom,Left'
    $btnAllLocal.Add_Click({
        for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
            $cell = $grid.Rows[$i].Cells['Destination']
            if (-not $cell.ReadOnly) {
                $cell.Value            = 'Local'
                $cell.Style.BackColor  = [System.Drawing.Color]::FromArgb(198, 239, 206)
                $cell.Style.ForeColor  = [System.Drawing.Color]::FromArgb(0, 97, 0)
            }
        }
    })
    $form.Controls.Add($btnAllLocal)

    # OK and Cancel buttons
    $btnOK                 = New-Object System.Windows.Forms.Button
    $btnOK.Size            = New-Object System.Drawing.Size(80, 28)
    $btnOK.Text            = 'OK'
    $btnOK.Font            = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnOK.Anchor          = 'Bottom,Right'
    $btnOK.DialogResult    = [System.Windows.Forms.DialogResult]::None
    # Do NOT set AcceptButton -- Enter key would close before ComboBox commits
    $form.Controls.Add($btnOK)

    # Force commit any active ComboBox cell edit before closing.
    # Without EndEdit(), a dropdown selection that hasn't lost focus
    # is not captured -- the cell retains the previous value.
    $btnOK.Add_Click({
        $grid.EndEdit()
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $btnCancel             = New-Object System.Windows.Forms.Button
    $btnCancel.Size        = New-Object System.Drawing.Size(80, 28)
    $btnCancel.Text        = 'Cancel'
    $btnCancel.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnCancel.Anchor      = 'Bottom,Right'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton     = $btnCancel
    $form.Controls.Add($btnCancel)

    # Position OK/Cancel relative to form size on resize -- Y tracks bottom of form
    $form.Add_Resize({
        $btnY = $form.ClientSize.Height - 40
        $btnOK.Location     = New-Object System.Drawing.Point(($form.ClientSize.Width - 176), $btnY)
        $btnCancel.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 92),  $btnY)
        $grid.Size          = New-Object System.Drawing.Size(($form.ClientSize.Width - 24), ($form.ClientSize.Height - 108))
    })
    # Set initial positions
    $btnY = $form.ClientSize.Height - 40
    $btnOK.Location     = New-Object System.Drawing.Point(($form.ClientSize.Width - 176), $btnY)
    $btnCancel.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 92),  $btnY)

    # -- Show form and process result --------------------------
    $result = $form.ShowDialog()

    # Capture grid values BEFORE disposing the form.
    # After Dispose() the grid controls are destroyed and
    # Rows.Count returns 0 -- all values would be lost.
    $capturedRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $destVal = $grid.Rows[$i].Cells['Destination'].FormattedValue
        if ($destVal -notin @('Server','Local')) { $destVal = 'Local' }
        $capturedRows.Add([PSCustomObject]@{
            StoreName   = $rows[$i].StoreName
            FolderPath  = $rows[$i].FolderPath
            Destination = $destVal
        })
    }

    $form.Dispose()

    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-OMMigrateLog -Message 'Folder map picker cancelled -- folder_map.csv not changed.' `
                           -Level INFO
        Write-Host '  Folder map picker cancelled -- folder_map.csv not changed.' `
                   -ForegroundColor Yellow
        Write-Host '  Open Config\folder_map.csv to review and edit destinations manually.' `
                   -ForegroundColor Yellow
        return
    }

    # Check for Excel file lock before writing
    try {
        $lockTest = [System.IO.File]::Open(
            $CsvPath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
        $lockTest.Close()
        $lockTest.Dispose()
    }
    catch {
        Write-OMMigrateLog -Message (
            "folder_map.csv is locked (Excel may have it open). " +
            "Folder picker selections were NOT saved. " +
            "Close Excel and re-run Script 00 to save your selections."
        ) -Level WARN
        Write-Host '  WARNING: folder_map.csv is open in Excel.' -ForegroundColor Yellow
        Write-Host '  Close Excel and re-run Script 00 to save your folder selections.' -ForegroundColor Yellow
        return
    }

    # Apply captured grid values and write
    $serverCount = 0
    $localCount  = 0

    foreach ($cr in $capturedRows) {
        Write-OMMigrateLog -Message "Folder picker row: '$($cr.FolderPath)' = $($cr.Destination)" -Level DEBUG
        switch ($cr.Destination) {
            'Server' { $serverCount++ }
            'Local'  { $localCount++  }
        }
    }

    Write-OMMigrateLog -Message "Folder picker grid read complete: Server=$serverCount | Local=$localCount" -Level DEBUG

    # When filtered, merge updated rows back into the full CSV so other
    # accounts' folder assignments are not overwritten.
    $rowsToWrite = $rows
    if ($FilterStoreName -and -not [string]::IsNullOrWhiteSpace($FilterStoreName)) {
        $updatedDest = @{}
        foreach ($cr in $capturedRows) {
            if ($cr.StoreName -and $cr.FolderPath) {
                $crStoreType = if ($cr.PSObject.Properties['StoreType'] -and $cr.StoreType) { $cr.StoreType } else { '' }
                $updatedDest["$($cr.StoreName)|$crStoreType|$($cr.FolderPath)"] = $cr.Destination
            }
        }
        $allCsvRows = Import-Csv -Path $CsvPath -Encoding UTF8
        foreach ($r in $allCsvRows) {
            if ($r.StoreName -and $r.FolderPath) {
                $rStoreType = if ($r.PSObject.Properties['StoreType'] -and $r.StoreType) { $r.StoreType } else { '' }
                $key = "$($r.StoreName)|$rStoreType|$($r.FolderPath)"
                if ($updatedDest.ContainsKey($key)) {
                    Write-OMMigrateLog -Message "Folder picker key match: '$key' -> $($updatedDest[$key])" -Level DEBUG
                    $r.Destination = $updatedDest[$key]
                }
            }
        }
        $rowsToWrite = $allCsvRows
    }
    else {
        # No filter -- update $rows directly from capturedRows by index
        for ($i = 0; $i -lt $capturedRows.Count; $i++) {
            $rows[$i].Destination = $capturedRows[$i].Destination
        }
    }

    try {
        $rowsToWrite | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

        Write-OMMigrateLog -Message (
            "Folder map picker saved: Server=$serverCount | Local=$localCount"
        ) -Level INFO

        Write-AuditEntry -Action 'PICKER_FOLDER_MAP_SAVED' `
                         -Detail "Server=$serverCount | Local=$localCount"

        Write-Host ''
        Write-Host '  Folder map saved:' -ForegroundColor Green
        Write-Host "    $serverCount folder(s) set to Server -- will sync to all devices." `
                   -ForegroundColor Green
        Write-Host "    $localCount folder(s) set to Local  -- stays on this desktop only." `
                   -ForegroundColor DarkGray
    }
    catch {
        Write-OMMigrateLog -Message "Failed to save folder map picker selections: $_" -Level WARN
        Write-Host '  WARNING: Could not save folder map selections.' -ForegroundColor Yellow
        Write-Host "  $_" -ForegroundColor Yellow
    }
}


# ============================================================
#  REGION: RULES INVENTORY PICKER
# ============================================================

function Invoke-RulesInventoryPicker {
    <#
    .SYNOPSIS
        Displays a WinForms popup letting the operator review all Outlook
        rules and toggle the NeedsFolderUpdate flag per rule.

    .DESCRIPTION
        Called by Script 00 immediately after rules_inventory.csv is written.
        Shows all rules grouped by account (RuleStoreName) in a DataGridView.
        The NeedsFolderUpdate column is a dropdown cell with True / False
        options. The operator reviews and changes flags as needed, then
        clicks OK to save. Cancel leaves the CSV untouched.

        Pre-populates each row with the NeedsFolderUpdate value already
        written by Export-RulesToCSV (True for rules with folder targets,
        False for all others). The operator sets a rule to False to tell
        Script 03 to skip updating that rule's folder target even though
        one was detected.

        Skipped automatically in WhatIf/Preview mode.

        Uses System.Windows.Forms -- built into PowerShell 5.1 on Windows.
        No external dependencies.

    .PARAMETER CsvPath
        Full path to rules_inventory.csv. Must exist before calling.

    .EXAMPLE
        Invoke-RulesInventoryPicker -CsvPath 'C:\...\Config\rules_inventory.csv'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,

        [Parameter(Mandatory = $false)]
        [string]$FilterStoreName = ''
        # When specified, only rules for this store are shown.
        # Used by Script 00 per-account loop to scope the picker to one
        # account at a time. When empty, all rules are shown.
    )

    # Skip entirely in WhatIf mode
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message 'WhatIf: Would display rules inventory picker -- skipped.' `
                           -Level INFO -WhatIfPrefix
        return
    }

    if (-not (Test-Path $CsvPath)) {
        Write-OMMigrateLog -Message "Invoke-RulesInventoryPicker: rules_inventory.csv not found at $CsvPath -- skipped." `
                           -Level WARN
        return
    }

    # Load CSV rows -- skip blank separator rows
    $allRows = Import-Csv -Path $CsvPath -Encoding UTF8
    $rows    = @($allRows | Where-Object {
        $_.RuleStoreName -and $_.RuleName -and
        -not [string]::IsNullOrWhiteSpace($_.RuleStoreName) -and
        -not [string]::IsNullOrWhiteSpace($_.RuleName)
    })

    # Filter to a specific store if requested
    if ($FilterStoreName -and -not [string]::IsNullOrWhiteSpace($FilterStoreName)) {
        $rows = @($rows | Where-Object {
            $_.RuleStoreName -eq $FilterStoreName -or
            $_.TargetStoreName -eq $FilterStoreName
        })
        Write-OMMigrateLog -Message "Invoke-RulesInventoryPicker: Filtered to store '$FilterStoreName' -- $($rows.Count) rule(s)." `
                           -Level INFO
    }

    if ($rows.Count -eq 0) {
        Write-OMMigrateLog -Message 'Invoke-RulesInventoryPicker: No rules in CSV -- picker skipped.' `
                           -Level INFO
        return
    }

    Write-OMMigrateLog -Message "Invoke-RulesInventoryPicker: Displaying picker for $($rows.Count) rule(s)." `
                       -Level INFO

    # Load WinForms assembly
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing       -ErrorAction Stop
    }
    catch {
        Write-OMMigrateLog -Message "Could not load WinForms -- rules picker unavailable: $_" -Level WARN
        Write-Host '  WARNING: Could not open rules picker (WinForms unavailable).' -ForegroundColor Yellow
        Write-Host '  Open Config\rules_inventory.csv and set NeedsFolderUpdate manually.' `
                   -ForegroundColor Yellow
        return
    }

    # -- Build the form ----------------------------------------
    $form                  = New-Object System.Windows.Forms.Form
    $form.Text             = if ($FilterStoreName) {
        "OMMigrate -- Rules Review: $FilterStoreName"
    } else {
        'OMMigrate -- Rules Inventory Review'
    }
    $form.Size             = New-Object System.Drawing.Size(860, 560)
    $form.StartPosition    = 'CenterScreen'
    $form.FormBorderStyle  = 'Sizable'
    $form.MaximizeBox      = $true
    $form.MinimizeBox      = $false
    $form.MinimumSize      = New-Object System.Drawing.Size(700, 420)
    $form.TopMost          = $true

    # Instruction label
    $label             = New-Object System.Windows.Forms.Label
    $label.Location    = New-Object System.Drawing.Point(12, 10)
    $label.Size        = New-Object System.Drawing.Size(820, 36)
    $label.Text        = 'Review Outlook rules. NeedsFolderUpdate = True means Script 03 will update that rule''s folder target after migration. Set to False to skip a rule.'
    $label.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.Controls.Add($label)

    # DataGridView
    $grid                          = New-Object System.Windows.Forms.DataGridView
    $grid.Location                 = New-Object System.Drawing.Point(12, 52)
    $grid.Size                     = New-Object System.Drawing.Size(820, 420)
    $grid.Anchor                   = 'Top,Bottom,Left,Right'
    $grid.AllowUserToAddRows       = $false
    $grid.AllowUserToDeleteRows    = $false
    $grid.AllowUserToResizeRows    = $false
    $grid.RowHeadersVisible        = $false
    $grid.SelectionMode            = 'FullRowSelect'
    $grid.MultiSelect              = $false
    $grid.AutoSizeColumnsMode      = 'None'
    $grid.Font                     = New-Object System.Drawing.Font('Segoe UI', 9)
    $grid.BackgroundColor          = [System.Drawing.SystemColors]::Window
    $grid.BorderStyle              = 'Fixed3D'
    $grid.EditMode                 = 'EditOnEnter'

    # Define columns

    # RuleStoreName column
    $colStore                = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStore.HeaderText     = 'Account'
    $colStore.Name           = 'RuleStoreName'
    $colStore.Width          = 160
    $colStore.ReadOnly       = $true
    # ADDED (Administrator direction -- row-index alignment fix, 2026-07-13): disables
    # click-to-sort on every column in this grid. Without this, the operator
    # could click a column header to sort the grid, which would silently
    # desync $grid.Rows[$i] from $rows[$i] at capture time further below --
    # both the RuleStoreName/RuleName capture and the TargetFolderPath capture
    # added this session depend on that index staying aligned with $rows'
    # original order.
    $colStore.SortMode       = 'NotSortable'
    $grid.Columns.Add($colStore) | Out-Null

    # RuleName column
    $colName                 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.HeaderText      = 'Rule Name'
    $colName.Name            = 'RuleName'
    $colName.Width           = 200
    $colName.ReadOnly        = $true
    $colName.SortMode        = 'NotSortable'
    $grid.Columns.Add($colName) | Out-Null

    # NeedsFolderUpdate column -- ComboBox dropdown (the only editable column)
    $colNeeds                = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $colNeeds.HeaderText     = 'NeedsFolderUpdate'
    $colNeeds.Name           = 'NeedsFolderUpdate'
    $colNeeds.Width          = 120
    $colNeeds.Items.AddRange(@('True','False'))
    $colNeeds.ReadOnly       = $false
    $colNeeds.SortMode       = 'NotSortable'
    $grid.Columns.Add($colNeeds) | Out-Null

    # IsEnabled column
    $colEnabled              = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colEnabled.HeaderText   = 'Enabled'
    $colEnabled.Name         = 'IsEnabled'
    $colEnabled.Width        = 60
    $colEnabled.ReadOnly     = $true
    $colEnabled.DefaultCellStyle.Alignment = 'MiddleCenter'
    $colEnabled.SortMode     = 'NotSortable'
    $grid.Columns.Add($colEnabled) | Out-Null

    # RuleType column
    $colType                 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colType.HeaderText      = 'Type'
    $colType.Name            = 'RuleType'
    $colType.Width           = 80
    $colType.ReadOnly        = $true
    $colType.SortMode        = 'NotSortable'
    $grid.Columns.Add($colType) | Out-Null

    # TargetFolderPath column
    $colTarget               = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colTarget.HeaderText    = 'Target Folder'
    $colTarget.Name          = 'TargetFolderPath'
    $colTarget.Width         = 180
    $colTarget.ReadOnly      = $true
    $colTarget.SortMode      = 'NotSortable'
    $grid.Columns.Add($colTarget) | Out-Null

    # Populate rows -- group by RuleStoreName with alternating background per account
    $storeColors = @('#F0F4FF','#FFFFFF')   # Light blue / white alternating per account
    $storeIndex  = -1
    $lastStore   = ''

    foreach ($row in $rows) {
        # WinForms popups always show real values -- sanitize is console-only
        $storeName  = $row.RuleStoreName
        $ruleName   = $row.RuleName
        $targetPath = $row.TargetFolderPath

        # Normalize NeedsFolderUpdate to True/False string for ComboBox
        # Default is ALWAYS True -- every rule must be migrated. Never default to False.
        $needsUpdate = if ($row.NeedsFolderUpdate -in @('False','false','0')) { 'False' } else { 'True' }

        $isEnabled = $row.IsEnabled
        $ruleType  = $row.RuleType

        # Alternate background color per store group
        if ($storeName -ne $lastStore) {
            $storeIndex++
            $lastStore = $storeName
        }
        $bgHex   = $storeColors[$storeIndex % 2]
        $bgColor = [System.Drawing.ColorTranslator]::FromHtml($bgHex)

        $rowIndex = $grid.Rows.Add($storeName, $ruleName, $needsUpdate, $isEnabled, $ruleType, $targetPath)
        $gridRow  = $grid.Rows[$rowIndex]
        $gridRow.DefaultCellStyle.BackColor = $bgColor

        # Highlight rows with NeedsFolderUpdate=True in a light yellow
        # so the operator can spot the rules that Script 03 will act on
        if ($needsUpdate -eq 'True') {
            $gridRow.DefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#FFFBE6')
        }
    }

    $form.Controls.Add($grid)

    # OK and Cancel buttons
    $btnOK                 = New-Object System.Windows.Forms.Button
    $btnOK.Size            = New-Object System.Drawing.Size(80, 28)
    $btnOK.Text            = 'OK'
    $btnOK.Font            = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnOK.Anchor          = 'Bottom,Right'
    $btnOK.DialogResult    = [System.Windows.Forms.DialogResult]::None
    # Do NOT set AcceptButton -- Enter key would close before ComboBox commits
    $form.Controls.Add($btnOK)

    # Force commit any active ComboBox cell edit before closing.
    # Without EndEdit(), a dropdown selection that hasn't lost focus
    # is not captured -- the cell retains the previous value.
    $btnOK.Add_Click({
        $grid.EndEdit()
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $btnCancel             = New-Object System.Windows.Forms.Button
    $btnCancel.Size        = New-Object System.Drawing.Size(80, 28)
    $btnCancel.Text        = 'Cancel'
    $btnCancel.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnCancel.Anchor      = 'Bottom,Right'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton     = $btnCancel
    $form.Controls.Add($btnCancel)

    # Position OK/Cancel relative to form width on resize
    $form.Add_Resize({
        $btnOK.Location     = New-Object System.Drawing.Point(($form.ClientSize.Width - 176), ($form.ClientSize.Height - 44))
        $btnCancel.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 92),  ($form.ClientSize.Height - 44))
        $grid.Size          = New-Object System.Drawing.Size(($form.ClientSize.Width - 24), ($form.ClientSize.Height - 108))
    })
    # Set initial positions
    $btnOK.Location     = New-Object System.Drawing.Point(($form.ClientSize.Width - 176), ($form.ClientSize.Height - 44))
    $btnCancel.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 92),  ($form.ClientSize.Height - 44))

    # -- Show form and process result --------------------------
    $result = $form.ShowDialog()

    # Capture grid values BEFORE disposing the form.
    # After Dispose() the grid controls are destroyed and
    # Rows.Count returns 0 -- all values would be lost.
    $capturedRules = [System.Collections.Generic.List[PSCustomObject]]::new()
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $needsVal = $grid.Rows[$i].Cells['NeedsFolderUpdate'].FormattedValue
        # Fix (2026-07-03, Administrator): fallback for an unexpected/unreadable grid cell
        # value previously defaulted to 'False' -- a fail-UNSAFE default that
        # could silently skip a rule's folder update with no operator action.
        # Administrator's permanent rule: NeedsFolderUpdate/IsEnabled always default to
        # True; only an explicit, deliberate operator edit may set False, and
        # that edit is preserved as-is (never overwritten back to True). An
        # unreadable cell is not a deliberate edit -- default to True instead.
        if ($needsVal -notin @('True','False')) { $needsVal = 'True' }
        # ADDED (Administrator direction -- merge key parity fix with Export-RulesToCSV,
        # 2026-07-13): capture TargetFolderPath alongside the existing fields so
        # the write-back below can key on RuleStoreName+TargetFolderPath instead
        # of RuleStoreName+RuleName. RuleName is not stable across a
        # consolidation rename (same root cause fixed in Export-RulesToCSV's own
        # merge key on 2026-07-12/13) -- a rule renamed by Script 03/the macro
        # between this picker's read and its write-back would silently fail to
        # match under the old RuleName-based key, losing the operator's
        # NeedsFolderUpdate edit for that rule.
        $capturedRules.Add([PSCustomObject]@{
            RuleStoreName     = $rows[$i].RuleStoreName
            RuleName          = $rows[$i].RuleName
            TargetFolderPath  = $rows[$i].TargetFolderPath
            NeedsFolderUpdate = $needsVal
        })
    }

    $form.Dispose()

    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-OMMigrateLog -Message 'Rules inventory picker cancelled -- rules_inventory.csv not changed.' `
                           -Level INFO
        Write-Host '  Rules inventory picker cancelled -- rules_inventory.csv not changed.' `
                   -ForegroundColor Yellow
        Write-Host '  Open Config\rules_inventory.csv to review and edit NeedsFolderUpdate manually.' `
                   -ForegroundColor Yellow
        return
    }

    # Check for Excel file lock before writing
    try {
        $lockTest = [System.IO.File]::Open(
            $CsvPath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
        $lockTest.Close()
        $lockTest.Dispose()
    }
    catch {
        Write-OMMigrateLog -Message (
            "rules_inventory.csv is locked (Excel may have it open). " +
            "Rules picker selections were NOT saved. " +
            "Close Excel and re-run Script 00 to save your selections."
        ) -Level WARN
        Write-Host '  WARNING: rules_inventory.csv is open in Excel.' -ForegroundColor Yellow
        Write-Host '  Close Excel and re-run Script 00 to save your rules selections.' -ForegroundColor Yellow
        return
    }

    # Apply grid NeedsFolderUpdate values back to the rows array.
    # Only real rows are in $rows (blanks were filtered on load).
    # When filtered, merge updates back into allRows so other accounts'
    # rules are not lost. Build a lookup of updated values, reload the
    # full CSV, apply changes, and write the full file back.
    $trueCount  = 0
    $falseCount = 0

    # CHANGED (Administrator direction -- merge key parity fix with Export-RulesToCSV,
    # 2026-07-13): keyed on RuleStoreName+TargetFolderPath instead of
    # RuleStoreName+RuleName -- same stable identity Export-RulesToCSV's own
    # merge block already uses, for the same reason (RuleName changes across a
    # consolidation rename; TargetFolderPath does not). Falls back to the old
    # RuleName-based key only when TargetFolderPath is blank on a row (should
    # not normally happen for a rule with NeedsFolderUpdate=True, but avoids
    # silently losing the update for a row this picker was never designed to
    # handle a blank path for).
    $updatedValues = @{}
    $updatedValuesByName = @{}
    foreach ($cr in $capturedRules) {
        if (-not [string]::IsNullOrWhiteSpace($cr.TargetFolderPath)) {
            $updatedValues["$($cr.RuleStoreName)|$($cr.TargetFolderPath)"] = $cr.NeedsFolderUpdate
        }
        $updatedValuesByName["$($cr.RuleStoreName)|$($cr.RuleName)"] = $cr.NeedsFolderUpdate
        if ($cr.NeedsFolderUpdate -eq 'True')  { $trueCount++  }
        if ($cr.NeedsFolderUpdate -eq 'False') { $falseCount++ }
    }

    # When filtered, reload the full CSV so we write all rows back,
    # not just the filtered subset. Apply updates only to matching rows.
    $writeRows = if ($FilterStoreName -and -not [string]::IsNullOrWhiteSpace($FilterStoreName)) {
        Import-Csv -Path $CsvPath -Encoding UTF8
    } else {
        $allRows
    }

    # CHANGED (Administrator direction -- merge key parity fix with Export-RulesToCSV,
    # 2026-07-13): tries the TargetFolderPath-based key first (stable across a
    # consolidation rename), falls back to the RuleName-based key only when the
    # path-based lookup misses -- same two-tier pattern as Export-RulesToCSV's
    # own merge block.
    foreach ($row in $writeRows) {
        if ($row.RuleStoreName -and $row.RuleName -and
            -not [string]::IsNullOrWhiteSpace($row.RuleStoreName) -and
            -not [string]::IsNullOrWhiteSpace($row.RuleName)) {
            $pathKey = if (-not [string]::IsNullOrWhiteSpace($row.TargetFolderPath)) {
                "$($row.RuleStoreName)|$($row.TargetFolderPath)"
            } else { $null }
            $nameKey = "$($row.RuleStoreName)|$($row.RuleName)"
            if ($pathKey -and $updatedValues.ContainsKey($pathKey)) {
                $row.NeedsFolderUpdate = $updatedValues[$pathKey]
            }
            elseif ($updatedValuesByName.ContainsKey($nameKey)) {
                $row.NeedsFolderUpdate = $updatedValuesByName[$nameKey]
            }
        }
    }

    try {
        # ADDED (Administrator direction, 2026-08-18): re-insert blank separator rows
        # between RuleStoreName groups before writing -- see Add-RulesCsvSeparatorRows
        # header comment for full rationale. $writeRows above is unfiltered (may
        # already contain old separator rows), so filter to real data rows first --
        # same filter pattern used at every other rules_inventory.csv write site --
        # before handing off to the helper, which does its own fresh sort and
        # re-insertion.
        $writeRowsForSeparator = @($writeRows | Where-Object {
            $_.RuleStoreName -and -not [string]::IsNullOrWhiteSpace($_.RuleStoreName) -and
            $_.RuleName      -and -not [string]::IsNullOrWhiteSpace($_.RuleName)
        })
        $writeRows = Add-RulesCsvSeparatorRows -Rows $writeRowsForSeparator
        $writeRows | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

        Write-OMMigrateLog -Message (
            "Rules inventory picker saved: NeedsFolderUpdate True=$trueCount | False=$falseCount"
        ) -Level INFO

        Write-AuditEntry -Action 'PICKER_RULES_INVENTORY_SAVED' `
                         -Detail "NeedsFolderUpdate: True=$trueCount | False=$falseCount"

        Write-Host ''
        Write-Host '  Rules inventory saved:' -ForegroundColor Green
        Write-Host "    $trueCount rule(s) flagged -- Script 03 will update their folder targets." `
                   -ForegroundColor Green
        Write-Host "    $falseCount rule(s) skipped -- Script 03 will leave these rules unchanged." `
                   -ForegroundColor DarkGray
    }
    catch {
        Write-OMMigrateLog -Message "Failed to save rules inventory picker selections: $_" -Level WARN
        Write-Host '  WARNING: Could not save rules inventory selections.' -ForegroundColor Yellow
        Write-Host "  $_" -ForegroundColor Yellow
    }
}


# ============================================================
#  REGION: RULES RECREATION (Script 03)
#  Shared extraction layer -- also used by Script 04 pattern.
# ============================================================
#  HELPER: Export-RulesBlob
#  Backs up the raw PR_RULES_DATA binary blob from the active
#  profile's default store root folder to disk. Called by Script 00
#  during discovery against the POP3 profile -- the only window
#  where the blob is healthy and not subject to overflow.
#
#  NOTE: After exhaustive diagnostic testing, neither PR_RULES_DATA
#  binary blob nor Rules.Export() are accessible on POP3 PST stores
#  via Outlook COM. This function logs an informational message only.
#  Rules are accessible exclusively via GetRules() enumeration during
#  a live COM session. No backup file artifact is produced.
#  Retained for future extensibility and consistent Script 00 flow.
# ============================================================

function Export-RulesBlob {
    <#
    .SYNOPSIS
        Logs rule backup status for POP3 PST stores.

    .DESCRIPTION
        Called by Script 00 during discovery. After exhaustive diagnostic
        testing, neither PR_RULES_DATA nor Rules.Export() are accessible
        on POP3 PST stores via Outlook COM. Logs an informational message
        and returns. Retained for consistent Script 00 flow and future
        extensibility if a viable backup path is discovered.

    .PARAMETER ProfileName
        The selected Outlook profile name -- used as the filename suffix.

    .OUTPUTS
        [string] Path to the .bin file written, or empty string on failure.

    .EXAMPLE
        $binPath = Export-RulesBlob -ProfileName 'Outlook'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ProfileName = ''
    )

    $binPath = ''

    $namespace = Get-OutlookNamespace
    if (-not $namespace) {
        Write-OMMigrateLog -Message 'Export-RulesBlob: No active COM session -- skipping blob backup.' -Level WARN
        return $binPath
    }

    # -- Resolve file naming ----------------------------------
    # Profile suffix sanitized for Windows filenames (same logic as Get-OMMigrateCsvPath)
    $profileSuffix = ''
    if (-not [string]::IsNullOrWhiteSpace($ProfileName)) {
        $profileSuffix = '_' + ($ProfileName -replace '[\\/:*?"<>|]', '_').Trim()
    }
    $backupDir = $Global:OMMigrate.BackupPath
    $binFile   = Join-Path $backupDir "rules_data${profileSuffix}.bin"

    # Ensure Backups directory exists
    if (-not (Test-Path $backupDir)) {
        try { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
        catch {
            Write-OMMigrateLog -Message "Export-RulesBlob: Could not create Backups directory: $_" -Level ERROR
            return $binPath
        }
    }

    # PR_RULES_DATA binary blob: not accessible on POP3 PST stores via
    # PropertyAccessor. Rules.Export() method also does not exist on the
    # Rules COM collection object in this environment. Both confirmed via
    # exhaustive diagnostic testing. No backup artifact is producible
    # via COM for POP3 PST stores -- rules are accessible only via
    # GetRules() enumeration during a live COM session.
    Write-OMMigrateLog -Message 'Export-RulesBlob: No backup artifact available for POP3 PST stores -- rules accessible via GetRules() enumeration only.' -Level INFO
    Write-Host '  rules backup        -> N/A (POP3 PST stores do not support binary rule export)' -ForegroundColor Gray

    return $binPath
}

# ============================================================


# ============================================================
#  HELPER: Set-RuleConditions
#  Sets the standard OMMigrate rule conditions on a rule object:
#  1. Clears the old From condition (POP3 era exact-match)
#  2. Sets SenderAddress condition (domain keyword)
#  3. Sets OnLocalMachine condition
#  (Account condition / ExecutionAccount permanently removed 2026-07-03 --
#  confirmed root cause of 0x800C8101 'devil code' COM read failures.)
#  All property assignments use InvokeMember -- direct assignment
#  leaves Enabled null in the MAPI stream causing UI to hide rules.
#  Shared by OMMigrate-03-Restore.ps1 and Restore-AmeriTechRules.ps1.
# ============================================================

function Set-RuleConditions {
    <#
    .SYNOPSIS
        Sets standard OMMigrate rule conditions on an Outlook rule
        COM object via InvokeMember reflection.

    .PARAMETER Rule
        The Outlook _Rule COM object to set conditions on.

    .PARAMETER RuleName
        Rule name -- used for log messages only.

    .PARAMETER RuleStoreName
        Account name (DisplayName or SmtpAddress) for the Account
        condition and ExecutionAccount. Looked up against
        $Namespace.Session.Accounts.

    .PARAMETER SendersDomain
        Domain keyword for the SenderAddress condition. Lowercased
        before assignment. If blank, SenderAddress condition is skipped.

    .PARAMETER Namespace
        Active Outlook MAPI namespace COM object -- used for account
        lookup via Session.Accounts.

    .OUTPUTS
        [bool] -- $true if all conditions set successfully, $false if
                  account lookup failed (non-fatal -- other conditions
                  are still set).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Rule,

        [Parameter(Mandatory = $true)]
        [string]$RuleName,

        [Parameter(Mandatory = $true)]
        [string]$RuleStoreName,

        [Parameter(Mandatory = $false)]
        [string]$SendersDomain = '',

        [Parameter(Mandatory = $true)]
        [object]$Namespace,

        [Parameter(Mandatory = $false)]
        [switch]$SkipExecutionAccount
    )

    $allOK = $true

    # -- 1. Clear old From condition (POP3 era exact-match) --
    # Gemini: disable via InvokeMember then remove all Recipients.
    # No global Reset() exists on the conditions object.
    try {
        $fromCondition = $Rule.Conditions.From
        [void]$fromCondition.GetType().InvokeMember(
            'Enabled',
            [System.Reflection.BindingFlags]::SetProperty,
            $null, $fromCondition, @($false)
        )
        while ($fromCondition.Recipients.Count -gt 0) {
            [void]$fromCondition.Recipients.GetType().InvokeMember(
                'Remove',
                [System.Reflection.BindingFlags]::InvokeMethod,
                $null, $fromCondition.Recipients, @(1)
            )
        }
        try { Write-OMMigrateLog -Message "From condition cleared: '$RuleName'" -Level DEBUG } catch { }
    } catch {
        try { Write-OMMigrateLog -Message "Failed to clear From condition on '$RuleName' (non-fatal): $_" -Level DEBUG } catch { }
    }

    # -- 2. Set SenderAddress condition --
    # Gemini: Address property takes a string array -- pass as
    # @(,$addressArray) to prevent PowerShell unrolling.
    # Value is lowercased -- email domains are case-insensitive per RFC 5321.
    #
    # FIXED (2026-07-07, Administrator live-evidence session): this function was the
    # REAL source of the SendersDomain "joined phrase, no or" regression --
    # confirmed via live Rules Manager evidence (ComEd -> "comed opower",
    # GlenbardWest73 -> "classcreator glenbardwest73", both as ONE address
    # array element with no OR). This is a SEPARATE code path from
    # Invoke-BuildRulesFromMap's harvesting fix earlier this session --
    # Set-RuleConditions is called from OMMigrate-03-Restore.ps1's Phase 3
    # folder-remap loop (~line 3170), which runs against EVERY rule with
    # NeedsFolderUpdate=True in the CSV, with NO LastDeployedRun/pending-row
    # gate at all -- so this bug fires on every Script 03 run, on every
    # eligible rule, not just newly-created ones. That's also why ComEd and
    # GlenbardWest73 were corrupted despite never having a blanked
    # LastDeployedRun -- they were never "skipped," Phase 3 has no concept
    # of skipping based on that column.
    #
    # ROOT CAUSE: the line below used to do
    #   $addressArray = [string[]]@($sendersDomainLower)
    # -- wrapping the ENTIRE raw $SendersDomain string (e.g. "comed opower")
    # as a SINGLE array element, with no split/validation at all. This
    # function never called ConvertTo-NormalizedSenderDomains, unlike the
    # CSV-input path (Export-RulesToCSV) and the existing-rule-harvesting
    # path (Invoke-BuildRulesFromMap) which both already use it correctly.
    #
    # FIX: reuse ConvertTo-NormalizedSenderDomains here too, the same
    # already-proven function used everywhere else in this module, so all
    # three SendersDomain-consuming code paths are finally consistent.
    # Multi-word values now correctly become multiple separate .Address
    # array elements (Outlook ORs them automatically), and full email
    # addresses are protected from being word-split around their "@".
    #
    # Also converted the Conditions.SenderAddress READ below from plain PS
    # dot-notation to InvokeMember GetProperty reflection, per this
    # module's standing COM rule (applies to ALL Outlook COM rule object
    # property access, reads and writes) -- this read site had never been
    # converted, matching the same class of gap found and fixed at four
    # other sites earlier this session.
    if (-not [string]::IsNullOrWhiteSpace($SendersDomain)) {
        try {
            $sendersDomainLower = $SendersDomain.ToLower()
            $addressArray       = @(ConvertTo-NormalizedSenderDomains -RawValue $sendersDomainLower)

            if ($addressArray.Count -eq 0) {
                try { Write-OMMigrateLog -Message "Set-RuleConditions: SendersDomain '$sendersDomainLower' normalized to zero valid tokens for '$RuleName' -- SenderAddress condition not set." -Level WARN } catch { }
            }
            else {
                $senderAddressCondition = $Rule.Conditions.GetType().InvokeMember(
                    'SenderAddress',
                    [System.Reflection.BindingFlags]::GetProperty,
                    $null, $Rule.Conditions, $null
                )
                [void]$senderAddressCondition.GetType().InvokeMember(
                    'Address',
                    [System.Reflection.BindingFlags]::SetProperty,
                    $null, $senderAddressCondition, @(,[string[]]$addressArray)
                )
                [void]$senderAddressCondition.GetType().InvokeMember(
                    'Enabled',
                    [System.Reflection.BindingFlags]::SetProperty,
                    $null, $senderAddressCondition, @($true)
                )
                try { Write-OMMigrateLog -Message "SenderAddress condition set: '$RuleName' -> [$($addressArray -join ' | ')]" -Level DEBUG } catch { }
            }
        } catch {
            try { Write-OMMigrateLog -Message "Failed to set SenderAddress condition on '$RuleName' (non-fatal): $_" -Level DEBUG } catch { }
        }
    } else {
        try { Write-OMMigrateLog -Message "SendersDomain blank for '$RuleName' -- SenderAddress condition not set." -Level WARN } catch { }
    }

    # -- 3. Set OnLocalMachine condition --
    # Gemini: purely a flag -- no value array, just enable it.
    # Required for client-side rule matching on IMAP accounts.
    try {
        $localMachineCondition = $Rule.Conditions.OnLocalMachine
        [void]$localMachineCondition.GetType().InvokeMember(
            'Enabled',
            [System.Reflection.BindingFlags]::SetProperty,
            $null, $localMachineCondition, @($true)
        )
        try { Write-OMMigrateLog -Message "OnLocalMachine condition set: '$RuleName'" -Level DEBUG } catch { }
    } catch {
        try { Write-OMMigrateLog -Message "Failed to set OnLocalMachine condition on '$RuleName' (non-fatal): $_" -Level DEBUG } catch { }
    }

    # -- 4. Account condition and ExecutionAccount: PERMANENTLY REMOVED --
    # (2026-07-03, Administrator's explicit instruction) The Account condition
    # ('through the specified account' filter) was confirmed via controlled
    # A/B testing to be the root cause of the 0x800C8101 'devil code' COM
    # read failures on Outlook 2021 Classic -- even a legitimately-resolved
    # Account condition triggers it. This block used to call
    # $Rule.Conditions.Account / .Enabled via InvokeMember, plus set
    # ExecutionAccount (cosmetic UI dropdown placement, tied to the same
    # targetAccount lookup) -- both fully deleted, not disabled, so this
    # cannot silently resurface again the way the separate Phase 1.5
    # re-enable block did (that block was deleted 2026-07-02; this is a
    # different, earlier root cause -- the original creation-time source
    # that Phase 1.5's removal did not address). No Account condition is
    # set on rules created via this function going forward.

    return $allOK
}



# ============================================================
#  HELPER: Invoke-PurgeAndRecreateRules
#  Detects PR_RW_RULES_STREAM blob overflow on the default
#  IMAP/OST store and performs a surgical reset followed by
#  full rule recreation from rules_inventory.csv.
#
#  Background: When 500+ rules with complex conditions are
#  created programmatically, the PR_RW_RULES_STREAM binary
#  blob on the IPM.RuleOrganizer hidden message in the Inbox
#  Associated Contents table grows beyond the COM deserializer
#  threshold. GetRules().Count works (reads a header int) but
#  Item(i) throws 0x800C8101 (deserializer overflow).
#
#  Fix: zero bytes 44-45 (little-endian uint16 rule count
#  field -- empirically confirmed across three accounts) to
#  reset IRulesCollection to empty, then recreate all rules
#  from CSV via GetRules().Create() with conditions set at
#  creation time to keep the stream compact.
#
#  Accessed via olFormItem (1) filter -- olHiddenItems (2)
#  filters out IPM.RuleOrganizer in non-Exchange profiles.
# ============================================================

function Invoke-PurgeAndRecreateRules {
    <#
    .SYNOPSIS
        Purges overflowed PR_RW_RULES_STREAM and recreates all
        rules from rules_inventory.csv on the default store.

    .PARAMETER Namespace
        Active Outlook MAPI namespace COM object.

    .PARAMETER RulesInventory
        Array of rule rows from rules_inventory.csv with
        NeedsFolderUpdate=True. Source of truth for all rule
        properties, conditions, and folder targets.

    .PARAMETER FolderMap
        Array of folder_map.csv rows -- used to resolve
        Server vs Local destination for each folder target.

    .PARAMETER ArchiveRootFolder
        Root MAPIFolder of the Archive PST -- used to navigate
        Local-destination folder targets.

    .PARAMETER AccountLookup
        Dictionary of SmtpAddress/DisplayName -> Account COM
        object -- used by Set-RuleConditions for Account
        condition and ExecutionAccount assignment.

    .OUTPUTS
        [PSCustomObject] with Created, Failed, Skipped counts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]  [object]   $Namespace,
        [Parameter(Mandatory = $true)]  [object[]] $RulesInventory,
        [Parameter(Mandatory = $true)]  [object[]] $FolderMap,
        [Parameter(Mandatory = $false)] [object]   $ArchiveRootFolder = $null,
        [Parameter(Mandatory = $true)]  [object]   $AccountLookup
    )

    $PR_RW_RULES_STREAM = "http://schemas.microsoft.com/mapi/proptag/0x68020102"
    $result = [PSCustomObject]@{ Created = 0; Failed = 0; Skipped = 0 }

    # ── Step 1: Locate IPM.RuleOrganizer on default Inbox ────────────────────
    Write-OMMigrateLog -Message 'Invoke-PurgeAndRecreateRules: Locating IPM.RuleOrganizer...' -Level INFO
    $inbox       = $null
    $ruleOrgItem = $null
    $entryID     = ''
    $storeID     = ''
    try {
        $inbox   = $Namespace.GetDefaultFolder(6)   # 6 = olFolderInbox
        $storeID = $inbox.Store.StoreID
        # olFormItem (1) -- olHiddenItems (2) filters out IPM.RuleOrganizer
        # in non-Exchange profiles. Confirmed empirically via MFCMAPI.
        $tbl = $inbox.GetTable("[MessageClass] = 'IPM.RuleOrganizer'", 1)
        while (-not $tbl.EndOfTable) {
            $row = $tbl.GetNextRow()
            $mc  = ''; try { $mc = $row['MessageClass'] } catch { }
            $ei  = ''; try { $ei = $row['EntryID']      } catch { }
            if ($mc -eq 'IPM.RuleOrganizer') { $entryID = $ei; break }
        }
        if ([string]::IsNullOrWhiteSpace($entryID)) {
            Write-OMMigrateLog -Message 'Invoke-PurgeAndRecreateRules: IPM.RuleOrganizer not found -- cannot purge.' -Level WARN
            return $result
        }
        $ruleOrgItem = $Namespace.GetItemFromID($entryID, $storeID)
        Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: IPM.RuleOrganizer found. Size=$($ruleOrgItem.Size) bytes." -Level INFO
    }
    catch {
        Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Failed to locate IPM.RuleOrganizer: $_" -Level WARN
        return $result
    }

    # ── Step 2: Read stream and zero bytes 44-45 (rule count field) ──────────
    # Bytes 44-45 = little-endian uint16 rule count field.
    # Empirically confirmed: 1 rule=01 00, 3 rules=03 00, 556 rules=2C 02.
    # Zeroing resets IRulesCollection to empty without touching rule data bytes.
    # Skip if count is already 0 (partial prior purge -- stream already zeroed).
    Write-OMMigrateLog -Message 'Invoke-PurgeAndRecreateRules: Reading PR_RW_RULES_STREAM and zeroing rule count...' -Level INFO
    try {
        $pa             = $ruleOrgItem.PropertyAccessor
        $streamBytes    = $pa.GetProperty($PR_RW_RULES_STREAM)
        $originalCount  = [System.BitConverter]::ToUInt16($streamBytes, 44)
        Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Stream=$($streamBytes.Length) bytes | Rule count field (bytes 44-45)=$originalCount (0x$($originalCount.ToString('X4')))" -Level INFO
        if ($originalCount -gt 0) {
            [byte[]]$modifiedStream    = $streamBytes.Clone()
            $modifiedStream[44]        = 0x00
            $modifiedStream[45]        = 0x00
            $pa.SetProperty($PR_RW_RULES_STREAM, $modifiedStream)
            $ruleOrgItem.Save()
            Write-OMMigrateLog -Message 'Invoke-PurgeAndRecreateRules: Stream zeroed and saved.' -Level INFO
        }
        else {
            Write-OMMigrateLog -Message 'Invoke-PurgeAndRecreateRules: Stream already zeroed (count=0) -- skipping purge step.' -Level INFO
        }
    }
    catch {
        Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Failed to zero stream: $_" -Level WARN
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ruleOrgItem) } catch { }
        return $result
    }

    # ── Step 3: Release COM objects and let OST I/O settle ───────────────────
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ruleOrgItem) } catch { }
    $ruleOrgItem = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2

    # ── Step 4: Verify IRulesCollection is now empty and functional ──────────
    $store      = $Namespace.DefaultStore
    $targetRules = $store.GetRules()
    $countAfter  = $targetRules.Count
    Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: GetRules().Count after purge = $countAfter" -Level INFO
    if ($countAfter -ne 0) {
        Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Count=$countAfter after purge -- expected 0. Aborting recreation." -Level WARN
        return $result
    }

    # Confirm Item(1) throws bounds error (not 0x800C8101)
    try {
        $testRule = $targetRules.Item(1)
        Write-OMMigrateLog -Message 'Invoke-PurgeAndRecreateRules: Item(1) returned object on empty collection -- unexpected.' -Level WARN
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($errMsg -like '*0x800C8101*') {
            Write-OMMigrateLog -Message 'Invoke-PurgeAndRecreateRules: Item(1) still throws 0x800C8101 -- collection still faulted. Aborting.' -Level WARN
            return $result
        }
        Write-OMMigrateLog -Message 'Invoke-PurgeAndRecreateRules: IRulesCollection confirmed empty and functional.' -Level INFO
    }

    # ── Step 5: Recreate all rules from RulesInventory ───────────────────────
    # Conditions set at creation time -- not in a separate patch pass.
    # ExecutionOrder read from CSV -- normalized to sequential 1-to-N per
    # account group by Invoke-NormalizeRulesExecutionOrder before this runs.
    # This keeps the per-rule binary footprint compact and prevents re-overflow.
    Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Recreating $($RulesInventory.Count) rules from CSV..." -Level INFO
    Write-Host "  Recreating $($RulesInventory.Count) rules from rules_inventory.csv -- this may take several minutes, please wait..." -ForegroundColor Cyan

    # Local sequential counter -- Outlook's ExecutionOrder property only
    # accepts values from 1 to the collection's current count. The global
    # CSV ExecutionOrder (1-597 across ALL accounts) can exceed this store's
    # own rule count when only a subset of global numbers fall in this
    # store's group, causing silent InvokeMember failures. $RulesInventory
    # is already in the correct relative order (filtered to this store and
    # sorted by global ExecutionOrder upstream), so a local 1-to-N counter
    # preserves alphabetical rank while staying within Outlook's valid range.
    $localExecOrder = 0

    foreach ($ruleRow in $RulesInventory) {
        $localExecOrder++
        $ruleName      = $ruleRow.RuleName
        $ruleStoreName = $ruleRow.RuleStoreName
        $sendersDomain = $ruleRow.SendersDomain
        $targetPath    = $ruleRow.TargetFolderPath
        $ruleType      = 0   # 0 = olRuleReceive

        try {
            # Create rule
            $newRule = $targetRules.GetType().InvokeMember(
                'Create',
                [System.Reflection.BindingFlags]::InvokeMethod,
                $null, $targetRules, @($ruleName, $ruleType)
            )
            Register-COMObject -ComObject $newRule

            # Enable rule via InvokeMember -- direct assignment silently fails
            [void]$newRule.GetType().InvokeMember(
                'Enabled',
                [System.Reflection.BindingFlags]::SetProperty,
                $null, $newRule, @($true)
            )

            # Set ExecutionOrder using local sequential position (1-to-N within
            # this store's collection), not the raw global CSV value, since
            # Outlook rejects ExecutionOrder values outside the collection's
            # current valid range. RulesInventory order already reflects the
            # correct global alphabetical rank.
            try {
                [void]$newRule.GetType().InvokeMember(
                    'ExecutionOrder',
                    [System.Reflection.BindingFlags]::SetProperty,
                    $null, $newRule, @($localExecOrder)
                )
            } catch {
                Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: ExecutionOrder failed for '$ruleName' (local=$localExecOrder): $_" -Level WARN
            }

            # Set conditions at creation time via Set-RuleConditions.
            # -SkipExecutionAccount keeps the per-rule binary footprint compact
            # to prevent PR_RW_RULES_STREAM re-overflow on large rule sets.
            # ExecutionAccount is cosmetic (UI dropdown) -- Account condition
            # is retained for functional account filtering.
            [void](Set-RuleConditions `
                -Rule                 $newRule `
                -RuleName             $ruleName `
                -RuleStoreName        $ruleStoreName `
                -SendersDomain        $sendersDomain `
                -Namespace            $Namespace `
                -SkipExecutionAccount)

            # Resolve folder target from FolderMap
            $newFolderCOM   = $null
            $folderResolved = $false
            if (-not [string]::IsNullOrWhiteSpace($targetPath)) {
                $folderEntry = $FolderMap | Where-Object {
                    $_.FolderPath -eq $targetPath
                } | Select-Object -First 1

                if ($folderEntry) {
                    if ($folderEntry.Destination -eq 'Server') {
                        # Navigate from default IMAP store
                        try {
                            $navStores = $Namespace.Stores
                            $navStore  = $null
                            for ($nsi = 1; $nsi -le $navStores.Count; $nsi++) {
                                try {
                                    $nst    = $navStores.Item($nsi)
                                    $nstFp  = ''; try { $nstFp = $nst.FilePath } catch { }
                                    $nstDn  = ''; try { $nstDn = $nst.DisplayName } catch { }
                                    $firstSeg = ($targetPath -split '\')[0]
                                    if ($nstDn -like "*$firstSeg*" -and
                                        -not ($nstFp -like '*Backups*') -and
                                        -not ($nstFp -like '*Archive*')) {
                                        $navStore = $nst; break
                                    }
                                } catch { }
                            }
                            if ($navStore) {
                                $pathParts = $targetPath.Split('\')
                                $curFolder = $Namespace.Folders.Item($navStore.DisplayName)
                                for ($pi = 1; $pi -lt $pathParts.Count; $pi++) {
                                    try { $curFolder = $curFolder.Folders.Item($pathParts[$pi]) }
                                    catch { $curFolder = $null; break }
                                }
                                if ($curFolder) { $newFolderCOM = $curFolder; $folderResolved = $true }
                            }
                        } catch { }
                    }
                    elseif ($folderEntry.Destination -eq 'Local' -and $ArchiveRootFolder) {
                        # Navigate from Archive PST root
                        try {
                            $pathParts = $targetPath.Split('\')
                            $curFolder = $ArchiveRootFolder
                            foreach ($part in $pathParts) {
                                if (-not [string]::IsNullOrWhiteSpace($part)) {
                                    try { $curFolder = $curFolder.Folders.Item($part) }
                                    catch { $curFolder = $null; break }
                                }
                            }
                            if ($curFolder) { $newFolderCOM = $curFolder; $folderResolved = $true }
                        } catch { }
                    }
                }
            }

            # Set MoveToFolder action via InvokeMember
            if ($folderResolved -and $newFolderCOM) {
                try {
                    $moveAction = $newRule.Actions.MoveToFolder
                    [void](Set-RuleFolderAction -Action $moveAction -Folder $newFolderCOM)
                    [void]$moveAction.GetType().InvokeMember(
                        'Enabled',
                        [System.Reflection.BindingFlags]::SetProperty,
                        $null, $moveAction, @($true)
                    )
                } catch { }
            }

            # StopProcessing is set via the named property 'Stop' (InvokeMember
            # reflection) in the post-Save pass below (Step 6c). Updated 2026-06-29
            # (Gemini consult + Administrator review): the prior Actions.Item(27) fixed-index
            # approach was replaced with the named property -- the same accessor
            # already proven and in production use on the current consolidated-
            # rules write path (Invoke-BuildRulesFromMap, 2026-06-26). The named
            # property is no longer null via COM interop once a rule has actually
            # been Saved; Item(27) relied on a hardcoded action-slot offset that is
            # fragile to rule-structure or profile-context shifts.

            $result.Created++
        }
        catch {
            Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Failed to create rule '$ruleName': $_" -Level WARN
            $result.Failed++
        }
    }

    Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Recreation complete -- Created=$($result.Created) Failed=$($result.Failed)" -Level INFO

    # ── Step 6a: Patch PR_RW_RULES_STREAM bytes 44-45 BEFORE IRulesCollection.Save() ─
    # Proved by Test-WriteRulesStream4.ps1: IRulesCollection.Save() preserves
    # bytes 44-45 when they are already set correctly before Save() is called.
    # Sequence: SetProperty(count) + item.Save() FIRST, then IRulesCollection.Save().
    # IRulesCollection.Save() writes the rule binary data and leaves bytes 44-45
    # intact because they were already committed to the OST before Save() ran.
    try {
        $PR_RW_PATCH  = 'http://schemas.microsoft.com/mapi/proptag/0x68020102'
        $patchInbox2  = $Namespace.GetDefaultFolder(6)
        $patchTbl2    = $patchInbox2.GetTable("[MessageClass] = 'IPM.RuleOrganizer'", 1)
        $patchEID2    = ''
        while (-not $patchTbl2.EndOfTable) {
            $pr2 = $patchTbl2.GetNextRow()
            $pmc2 = ''; try { $pmc2 = $pr2['MessageClass'] } catch { }
            $pei2 = ''; try { $pei2 = $pr2['EntryID']      } catch { }
            if ($pmc2 -eq 'IPM.RuleOrganizer') { $patchEID2 = $pei2; break }
        }
        if (-not [string]::IsNullOrWhiteSpace($patchEID2)) {
            $patchItem2   = $Namespace.GetItemFromID($patchEID2, $patchInbox2.Store.StoreID)
            $patchPa2     = $patchItem2.PropertyAccessor
            $patchStream2 = $patchPa2.GetProperty($PR_RW_PATCH)
            $currentCount = [System.BitConverter]::ToUInt16($patchStream2, 44)
            $targetCount  = $result.Created
            Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Pre-Save stream patch -- bytes 44-45 currently=$currentCount, setting to $targetCount." -Level INFO
            [byte[]]$patchBytes2  = $patchStream2.Clone()
            $patchBytes2[44]      = [byte]($targetCount -band 0xFF)
            $patchBytes2[45]      = [byte](($targetCount -shr 8) -band 0xFF)
            $patchPa2.SetProperty($PR_RW_PATCH, $patchBytes2)
            $patchItem2.Save()
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($patchItem2) } catch { }
            Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Pre-Save stream patch committed. Bytes 44-45 set to $targetCount." -Level INFO
        } else {
            Write-OMMigrateLog -Message 'Invoke-PurgeAndRecreateRules: IPM.RuleOrganizer not found for pre-Save patch -- proceeding to Save() anyway.' -Level WARN
        }
    } catch {
        Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Pre-Save stream patch failed (non-fatal): $_" -Level WARN
    }

    # ── Step 6b: Save rules collection ───────────────────────────────────────
    # Called AFTER the bytes 44-45 patch -- IRulesCollection.Save() writes the
    # rule binary data and preserves the pre-patched count field because it was
    # already committed to the OST before Save() ran (proved by Test4).
    # ExecutionOrder set at creation time from CSV (normalized by
    # Invoke-NormalizeRulesExecutionOrder) -- no sort/renumber pass needed.
    try {
        [void]$targetRules.GetType().InvokeMember(
            'Save',
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null, $targetRules, @($false)
        )
        Write-Host "  Rules recreated and saved: $($result.Created) rule(s)." -ForegroundColor Green
        Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: IRulesCollection.Save() succeeded. Created=$($result.Created)." -Level INFO

        # ── Verify bytes 44-45 preserved after Save() ────────────────────────
        # Confirm IRulesCollection.Save() did not zero bytes 44-45.
        # Item(1) not used -- throws on overflow-history environments.
        try {
            $PR_RW_VERIFY  = 'http://schemas.microsoft.com/mapi/proptag/0x68020102'
            $verifyInbox   = $Namespace.GetDefaultFolder(6)
            $verifyTbl     = $verifyInbox.GetTable("[MessageClass] = 'IPM.RuleOrganizer'", 1)
            $verifyEID     = ''
            while (-not $verifyTbl.EndOfTable) {
                $vr = $verifyTbl.GetNextRow()
                $vmc = ''; try { $vmc = $vr['MessageClass'] } catch { }
                $vei = ''; try { $vei = $vr['EntryID']      } catch { }
                if ($vmc -eq 'IPM.RuleOrganizer') { $verifyEID = $vei; break }
            }
            if (-not [string]::IsNullOrWhiteSpace($verifyEID)) {
                $verifyItem   = $Namespace.GetItemFromID($verifyEID, $verifyInbox.Store.StoreID)
                $verifyStream = $verifyItem.PropertyAccessor.GetProperty($PR_RW_VERIFY)
                $verifyCount  = [System.BitConverter]::ToUInt16($verifyStream, 44)
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($verifyItem) } catch { }
                if ($verifyCount -eq $result.Created) {
                    Write-Host "  [OK] Bytes 44-45 = $verifyCount after Save() -- count preserved." -ForegroundColor Green
                    Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Bytes 44-45 = $verifyCount after Save() -- count preserved." -Level INFO
                } else {
                    Write-Host "  [WARN] Bytes 44-45 = $verifyCount after Save() -- expected $($result.Created). Rules may not persist." -ForegroundColor Yellow
                    Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Bytes 44-45 = $verifyCount after Save() -- expected $($result.Created)." -Level WARN
                }
            }
        } catch {
            Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Post-Save byte verification failed (non-fatal): $_" -Level WARN
        }

        # ── Step 6c: Post-Save StopProcessing pass ────────────────────────────
        # Reuses $targetRules (the SAME collection object already Saved above)
        # rather than calling GetRules() again -- a second GetRules() call on
        # the same store within one COM session corrupts Item() enumeration
        # on that store (confirmed June 17, 2026: Count succeeds, Item(1)
        # throws 0x800C8101 afterward, even on a freshly recreated collection).
        # StopProcessing is mandatory on every rule -- no exceptions.
        # Updated 2026-06-29 (Gemini consult + Administrator review): replaced the
        # June 17-era Actions.Item(27) fixed-index write with the named
        # property 'Stop', resolved via InvokeMember GetProperty reflection --
        # the same mechanism already proven and in production use on the
        # current consolidated-rules write path (Invoke-BuildRulesFromMap,
        # 2026-06-26). Item(27) relied on a hardcoded action-slot offset
        # that is fragile to rule-structure or profile-context shifts; the
        # named property is the safer, more accurate accessor and is no
        # longer null via COM interop once the rule has actually been Saved.
        try {
            $spRules = $targetRules
            $spCount = $spRules.Count
            $spSet   = 0
            $spFailed = 0
            for ($spi = 1; $spi -le $spCount; $spi++) {
                try {
                    $spRule = $spRules.Item($spi)
                    $sprAction = $spRule.Actions.GetType().InvokeMember(
                        'Stop',
                        [System.Reflection.BindingFlags]::GetProperty,
                        $null, $spRule.Actions, $null
                    )
                    if ($null -ne $sprAction) {
                        [void]$sprAction.GetType().InvokeMember(
                            'Enabled',
                            [System.Reflection.BindingFlags]::SetProperty,
                            $null, $sprAction, @($true)
                        )
                        $spSet++
                    } else {
                        $spFailed++
                    }
                } catch {
                    $spFailed++
                }
            }
            if ($spFailed -gt 0) {
                Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Post-Save StopProcessing pass -- Set=$spSet Failed=$spFailed of $spCount rule(s)." -Level WARN
            } else {
                Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Post-Save StopProcessing pass -- Set=$spSet of $spCount rule(s)." -Level INFO
            }
            Write-Host "  StopProcessing set on $spSet of $spCount rule(s)." -ForegroundColor Green

            # Save again to commit the StopProcessing changes
            if ($spSet -gt 0) {
                try {
                    [void]$spRules.GetType().InvokeMember(
                        'Save',
                        [System.Reflection.BindingFlags]::InvokeMethod,
                        $null, $spRules, @($false)
                    )
                    Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Post-StopProcessing Save() succeeded." -Level INFO
                } catch {
                    Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Post-StopProcessing Save() failed: $_" -Level WARN
                }
            }
        } catch {
            Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: Post-Save StopProcessing pass failed (non-fatal): $_" -Level WARN
        }
    } catch {
        Write-OMMigrateLog -Message "Invoke-PurgeAndRecreateRules: IRulesCollection.Save() failed: $_" -Level WARN
    }

    return $result
}

# ============================================================
#  HELPER: Invoke-DeployConsolidatedRules
#  PowerShell port of Gemini's DeployConsolidatedRules VBA macro
#  (Module3.bas). Reads rules_inventory.csv rows where
#  LastDeployedRun is blank and TargetFolderPath/SendersDomain
#  are set, groups them by RuleStoreName|TargetFolderPath,
#  chunks each group's sender domains 5-at-a-time into newly
#  created consolidated rules, and writes a timestamp back to
#  LastDeployedRun for every domain processed -- so reruns of
#  either this function or Gemini's macro against the same CSV
#  stay in sync and never reprocess already-deployed domains.
#
#  Mirrors the macro's logic exactly, not a redesign:
#    - Reverse iteration over group keys so Outlook's natural
#      append-order creation produces clean A-to-Z display order
#      (same technique the macro uses, same reason).
#    - One rule created and Saved per chunk of up to 5 domains --
#      Save() is called immediately after each individual rule,
#      not once at the end across the whole batch.
#    - Rule naming convention: "Rule: [<account>] <folder> (Part N)"
#    - Default-store fallback for ameritech specifically, mirroring
#      the macro's "AMERITECH DEFAULT PROFILE FIX" step.
#
#  Differences from the macro are COM-quirk workarounds already
#  proven elsewhere in this module, not stylistic changes:
#    - StopProcessing set via the named property Actions.Stop,
#      resolved via InvokeMember GetProperty reflection (proven
#      2026-06-26 over 5 rounds of live testing). Note: this
#      doc-block previously and incorrectly described this as
#      Actions.Item(27) -- corrected 2026-06-29 (Gemini consult +
#      Administrator review) to match what the code actually does; the
#      named property is also now the approach used in
#      Invoke-PurgeAndRecreateRules Step 6c and Invoke-RulesRecreation.
#    - MoveToFolder assigned via Set-RuleFolderAction (.NET
#      Reflection put_Folder) rather than direct property
#      assignment, which the PowerShell CLR translation layer
#      silently drops.
#    - Folder navigation via Get-OrCreateFolder/Get-FolderByPath
#      (existing proven helpers) instead of a new GetOrCreateArchiveFolder.
# ============================================================

# ============================================================
#  HELPER: Invoke-ResortRulesByLabel
#  Added 2026-07-02, Administrator. Repositions EXISTING rules in a live
#  Outlook.Rules collection into flat alphabetical order by their
#  final display label, WITHOUT deleting or recreating any rule.
#
#  Design basis: adapted from a public, freely-shared community VBA
#  macro (SortRulesbyAlpha, Slipstick Systems forums, 2014) that
#  proved this exact technique -- snapshot rules, sort by name,
#  write ExecutionOrder sequentially, single Save() at the end.
#  Rebuilt here using this project's proven current standards:
#  InvokeMember reflection for all writes (Memory #10, not the
#  original macro's direct dot-notation), leaf-label sort key via
#  the same regex Invoke-NormalizeRulesExecutionOrder uses (so
#  standardized "Rule: [account] Label (Part N)" rules sort on
#  Label, manual rules sort on their raw Name), and object-reference
#  snapshot (not name-string re-lookup via Item(j), which carries
#  unverified/duplicate-name risk at this rule count).
#
#  CRITICAL DESIGN CONSTRAINT (Administrator, 2026-07-02): this function must
#  NEVER call Rules.Create() or Rules.Remove() -- ExecutionOrder is
#  documented by Microsoft to be directly mapped to collection Index
#  (Rule.ExecutionOrder property docs), so writing it repositions a
#  rule WITHOUT touching its Conditions/Actions/Enabled/Name. This
#  is required specifically so a rule an end user has manually
#  edited via the Outlook UI (added a condition, disabled it, etc.)
#  is NEVER silently reverted just because the list needed re-sorting
#  -- unlike a full delete+recreate pass, which would discard any
#  such manual customization.
# ============================================================

function Invoke-ResortRulesByLabel {
    <#
    .SYNOPSIS
        Repositions all rules in a live Outlook.Rules collection into
        flat alphabetical order by label, via ExecutionOrder writes
        only -- never deletes or recreates any rule.

    .DESCRIPTION
        Snapshots every rule object reference from the collection
        (no Create/Remove), computes each rule's sort label using
        the same standardized-name regex as
        Invoke-NormalizeRulesExecutionOrder (falls back to the raw
        Name for non-standardized/manually-added rules), sorts the
        snapshot array by that label, then writes ExecutionOrder
        sequentially (1-to-N) onto each rule via InvokeMember,
        finishing with a single Rules.Save().

        Because this only ever writes the ExecutionOrder property --
        never Conditions, Actions, Enabled, or Name -- any manual
        customization an end user has made to a rule via the Outlook
        UI survives this call untouched.

    .PARAMETER TargetRules
        Live Outlook.Rules COM collection to resort (already scoped
        to one account/store, same as every other function in this
        module's rules pipeline).

    .OUTPUTS
        [int] Number of rules resorted, or -1 on failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [object] $TargetRules
    )

    try {
        $ruleCount = $TargetRules.Count
        if ($ruleCount -le 1) { return $ruleCount }

        # Snapshot every rule OBJECT REFERENCE (not name/index) so the
        # later write step never needs to re-look-up a rule by name --
        # avoids any duplicate-name Item(string) ambiguity entirely.
        $ruleSnapshot = @()
        for ($ri = 1; $ri -le $ruleCount; $ri++) {
            $ruleSnapshot += $TargetRules.Item($ri)
        }

        # Sort key: same regex Invoke-NormalizeRulesExecutionOrder uses --
        # extract "Label" from "Rule: [account] Label (Part N)"; fall
        # back to the raw Name for manual/non-standardized rules, so
        # they interleave naturally at their own alphabetical position.
        $sortedRules = @($ruleSnapshot | Sort-Object {
            $sortRuleName = [string]$_.Name
            $sortLabel    = $sortRuleName
            if ($sortRuleName -match '^Rule: \[[^\]]*\] (.+?) \(Part \d+\)$') {
                $sortLabel = $Matches[1]
            }
            $sortLabel.Trim().ToLower()
        })

        # Write ExecutionOrder sequentially via InvokeMember (Memory #10 --
        # writes are proven safe via SetProperty alone, no BindingFlags
        # combination needed the way reads require).
        $newOrder = 1
        foreach ($sortedRule in $sortedRules) {
            try {
                $sortedRule.GetType().InvokeMember(
                    'ExecutionOrder',
                    [System.Reflection.BindingFlags]::SetProperty,
                    $null, $sortedRule, @($newOrder)
                )
            }
            catch {
                Write-OMMigrateLog -Message "Invoke-ResortRulesByLabel: ExecutionOrder write failed for '$($sortedRule.Name)' (target=$newOrder): $_" -Level WARN
            }
            $newOrder++
        }

        $TargetRules.Save()
        Write-OMMigrateLog -Message "Invoke-ResortRulesByLabel: $($sortedRules.Count) rule(s) resorted." -Level INFO
        return $sortedRules.Count
    }
    catch {
        Write-OMMigrateLog -Message "Invoke-ResortRulesByLabel: Failed: $_" -Level WARN
        return -1
    }
}

function Invoke-DeployConsolidatedRules {
    <#
    .SYNOPSIS
        Builds consolidated, alphabetically-ordered Outlook rules from
        rules_inventory.csv -- PowerShell equivalent of Gemini's
        DeployConsolidatedRules VBA macro.

    .DESCRIPTION
        For each RuleStoreName|TargetFolderPath group in the CSV with
        a blank LastDeployedRun, creates one or more new rules (each
        covering up to 5 sender domains via the SenderAddress
        condition), sets the Account condition, StopProcessing action,
        and MoveToFolder action, and saves each rule immediately after
        creation. Writes a shared timestamp to LastDeployedRun for
        every domain successfully processed in this run.

        Existing rules are never read, reordered, or modified --
        this function only ever creates new rules, exactly like the
        macro. Re-running this function or Gemini's macro against the
        same CSV will only process rows still showing a blank
        LastDeployedRun, so the two tools stay in sync with each other.

        PICKER SCOPING (added 2026-07-06, Administrator): when $ScopedAccountNames
        is provided, BOTH the pending-row grouping (Step 1) AND the final
        resort pass are limited to only those account(s) -- an account not
        in the list is left completely untouched this run: no rule
        creation, no LastDeployedRun stamp, no resort. This is a hard,
        explicit requirement, not a convenience default: "my requirement
        for using a picker is to only process activity for those accounts
        I pick. If I select all then I expect it to process all the
        accounts, not to do what it wants." (Administrator, 2026-07-06). Confirmed
        via live testing that prior to this fix, picking ONE account in
        Script 03's picker still silently deployed and resorted SIX other
        accounts' pending rows, because neither Step 1 nor the resort pass
        had any awareness of picker selection at all -- caught by Administrator
        directly comparing the Script 03 log and the resulting CSV. When
        $ScopedAccountNames is NOT provided (default, $null), behavior is
        completely unchanged from before this fix -- every pending row and
        every account is processed, matching every existing caller's
        current contract.

    .PARAMETER Namespace
        Active Outlook MAPI namespace COM object.

    .PARAMETER RulesInventory
        Array of rule rows from rules_inventory.csv (all rows, not
        pre-filtered) -- must include RuleStoreName, TargetFolderPath,
        SendersDomain, LastDeployedRun columns.

    .PARAMETER CsvPath
        Full path to rules_inventory.csv -- LastDeployedRun timestamps
        are written back to this file for processed rows.

    .PARAMETER ArchiveRootFolder
        Root MAPIFolder of the default/fallback Archive PST -- used for
        any group whose own TargetStoreName is blank or cannot be
        resolved among currently attached stores (see
        groupKeyToArchiveFolder/OverrideArchiveFolder). Not tied to any
        specific PST name; whichever archive store is passed in here is
        the fallback destination when no per-group mapping applies.

    .PARAMETER ScopedAccountNames
        Optional. When provided, limits BOTH rule deployment (Step 1) and
        the final resort pass to only these account name(s) -- every other
        account's rows/rules are left completely untouched this run. Pass
        the caller's account picker selection here to make picker
        selection authoritative for this function, per Administrator's explicit
        requirement (2026-07-06). When omitted or $null (default), every
        pending row and every account is processed, unchanged from prior
        behavior. Comparison is case-insensitive.

    .OUTPUTS
        [PSCustomObject] with Created, Failed, DomainsProcessed counts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]  [object]   $Namespace,
        [Parameter(Mandatory = $true)]  [object[]] $RulesInventory,
        [Parameter(Mandatory = $true)]  [string]   $CsvPath,
        [Parameter(Mandatory = $true)]  [object]   $ArchiveRootFolder,

        # Added 2026-07-02, Administrator. Mirrors the existing -RefreshRulesOnly
        # convention (Script 03). When present, skips Steps 1-3 (build/
        # consolidation) entirely and only runs the final Invoke-ResortRulesByLabel
        # pass across every account's live rules collection -- use this to
        # re-sort the full Rules and Alerts list into flat alphabetical order
        # without creating, deleting, or modifying any rule's conditions/actions.
        # When NOT present (default), Steps 1-3 run as normal, and the resort
        # pass STILL always runs afterward -- this satisfies Administrator's requirement
        # that the full rule collection is correctly sorted every run, whether
        # 0, 1, or all CSV rows were pending, without needing to delete/recreate
        # unrelated rules just to fix the display position of a few new ones.
        [Parameter(Mandatory = $false)] [switch]   $RefreshRulesSortOnly,

        # Added 2026-07-02, Administrator. Passed through to Invoke-BuildRulesFromMap --
        # see that function's SaveBatchSize parameter doc for full rationale.
        # Default of 1 preserves exact prior behavior (save every rule).
        [Parameter(Mandatory = $false)] [int]      $SaveBatchSize = 1,

        # Added 2026-07-06, Administrator (picker-scoping fix). When provided, ONLY
        # rows whose RuleStoreName matches one of these account names are
        # eligible for deployment this run -- every other row is left
        # untouched (not created, not stamped with LastDeployedRun) even if
        # it has a blank LastDeployedRun and would otherwise be pending.
        # This makes picker selection authoritative for what this function
        # actually does, not just what the caller reports doing -- per
        # Administrator's explicit requirement (2026-07-06): "my requirement for
        # using a picker is to only process activity for those accounts I
        # pick. If I select all then I expect it to process all the
        # accounts, not to do what it wants." Confirmed via live testing
        # (2026-07-06 Script 03 run) that without this parameter, picking
        # ONE account still deployed 6 OTHER accounts' pending rows,
        # because Step 1's grouping loop below considered every row in
        # $RulesInventory regardless of which account the operator actually
        # selected. When this parameter is NOT provided (default, $null),
        # behavior is unchanged from before -- every pending row is
        # eligible, matching the macro's/older callers' existing contract
        # (no accidental behavior change for any call site that doesn't
        # pass this new parameter). Comparison is case-insensitive.
        [Parameter(Mandatory = $false)] [string[]] $ScopedAccountNames = $null
    )

    # ConsolidatedRuleNames added (2026-07-01, Administrator): tracks the exact Name of
    # every pre-existing rule deleted by the consolidation scan in
    # Invoke-BuildRulesFromMap this run. Script 03's separate Strategy 1 remap
    # phase (OMMigrate-03-Restore.ps1) uses this list to skip rules that no
    # longer exist -- without it, Strategy 1 reads from its own pre-fetched
    # $liveCollections snapshot (taken before this function ran) and still
    # "sees" and remaps rules already deleted here, leaving them standing
    # alongside the newly-created consolidated rule (confirmed live,
    # 2026-07-01: "TestProfile ACE" deleted here, then remapped again by Strategy 1).
    #
    # ConsolidatedRuleTargets added (2026-07-09, Administrator direction): Administrator's
    # explicit design intent for the two-column LastDeployedRun/LastTargetRun
    # split (2026-07-07) was that BOTH consolidation work and folder-target
    # work complete in a SINGLE Script 03 run, on a genuine first run (or any
    # run where both columns are blank) -- not across two separate runs. The
    # ConsolidatedRuleNames skip above was correct to stop Strategy 1 from
    # remapping a stale/deleted rule reference, but it had the unintended side
    # effect of leaving LastTargetRun blank for every rule consolidation
    # touched this run, silently deferring that work to a second run with no
    # indication to the operator that anything was incomplete -- confirmed
    # live 2026-07-09 on "TestProfile": 360 of 361 rules consolidated in one run,
    # only 4 got LastTargetRun stamped, 356 silently left pending. A tool
    # requiring two runs to reach a stable state, with no prompt telling the
    # operator to run again, is not acceptable end-user behavior.
    # ConsolidatedRuleTargets is a name -> TargetFolderPath lookup (NOT just a
    # name list) populated at the exact moment a new consolidated rule is
    # successfully created and saved below -- Invoke-BuildRulesFromMap already
    # knows the correct resolved TargetFolderPath for that rule at creation
    # time (it just set the MoveToFolder action to it), so Phase 3 can use
    # this to stamp LastTargetRun directly for a same-run-consolidated rule,
    # without needing to re-fetch, re-navigate, or touch the live COM object
    # again -- avoiding the exact stale-reference problem ConsolidatedRuleNames
    # was built to prevent, while still completing the folder-target job in
    # the same pass consolidation ran in.
    $result = [PSCustomObject]@{
        Created                = 0
        Failed                 = 0
        DomainsProcessed       = 0
        RuleActionCounter      = 0
        ConsolidatedRuleNames  = [System.Collections.Generic.List[string]]::new()
        ConsolidatedRuleTargets = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        # ADDED (Administrator direction, this same fix): old raw rule name -> new
        # standardized rule name, for every rule absorbed by consolidation
        # this run. Invoke-DeployConsolidatedRules uses this after each
        # Invoke-BuildRulesFromMap call to rename the corresponding row in
        # $RulesInventory BEFORE the CSV is written, so RuleName in the CSV
        # matches the live rule's actual current name in the SAME run --
        # not the next one.
        ConsolidatedOldToNewName = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        # ADDED (Administrator direction, this same fix -- corrected approach): MapKey
        # (RuleStoreName|TargetFolderPath) -> new standardized rule name.
        # Covers every CSV row in a consolidated group, whether the rule for
        # that group already existed and was absorbed, or is being created
        # fresh this run -- see ConsolidatedMapKeyToNewName note at its write
        # site in Invoke-BuildRulesFromMap for full rationale.
        ConsolidatedMapKeyToNewName = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        # ADDED (Administrator direction, 2026-07-21, fix for the Part 1/Part 2
        # RuleName collision found via live testing): individual domain string ->
        # new standardized chunk name. Unlike ConsolidatedMapKeyToNewName (keyed by
        # RuleStoreName|TargetFolderPath, which collides when multiple chunks share
        # the same folder target) and ConsolidatedOldToNewName (only populated when
        # a rule was genuinely absorbed from an existing live rule -- empty for a
        # chunk created fresh with nothing to absorb, exactly the scenario that
        # exposed this bug), this key is inherently collision-free: the chunking
        # loop in Invoke-BuildRulesFromMap consumes each domain from the group's
        # domain list exactly once, sequentially, into exactly one chunk -- no
        # domain can ever belong to two chunks. Matching each CSV row by its own
        # SendersDomain against this dictionary correctly identifies which
        # specific chunk (Part 1 vs Part 2) that row belongs to, regardless of
        # whether anything was absorbed.
        ConsolidatedDomainToNewName = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        # ADDED (Administrator direction, 2026-07-21, second pass -- 3-column correction):
        # captures the authoritative, final SendersDomain string for each chunk --
        # the exact same $chunkArray values written to the live rule's
        # SenderAddress.Address, joined the same way this project already joins
        # SendersDomain elsewhere (space-separated). Keyed by chunk name, since
        # that is unique per chunk (unlike TargetFolderPath, which collides).
        # Administrator's explicit direction: only 3 columns need correcting per CSV row --
        # RuleName, SendersDomain, and Conditions -- all sourced from what
        # Outlook's Rules Manager actually shows for the live rule, not
        # reconstructed from the pre-run row's now-potentially-stale content.
        ConsolidatedChunkSendersDomain = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        # ADDED (Administrator direction, 2026-08-21 -- Conditions column accuracy
        # fix): captures the FULL, authoritative conditions summary for each
        # chunk, exactly as Get-RuleConditionsSummary reads it back from the
        # live $newRule COM object right after every condition (SenderAddress
        # plus any preserved Subject/Body/etc. from an absorbed rule) has been
        # set on it -- same function Get-OutlookRules already uses to build
        # this exact CSV column on a normal Script 00 rescan. Previously the
        # Step 4 CSV write-back (see that site's own comment) hardcoded a
        # "Sender address: ..." string built only from ConsolidatedChunkSendersDomain,
        # silently dropping any preserved Subject/Body/etc. condition text from
        # the CSV even though the LIVE rule correctly kept it -- confirmed live
        # 2026-08-21: a rule absorbed with an existing "Subject contains: X"
        # condition correctly showed both conditions in Outlook's Rules Manager
        # UI, but rules_inventory.csv only showed "Sender address: ..." until a
        # separate Script 00 rescan overwrote it with the correct combined text
        # -- meaning the CSV could stay wrong for days/weeks until that next
        # rescan happened to run. Keyed by chunk name, same as
        # ConsolidatedChunkSendersDomain.
        ConsolidatedChunkConditionsText = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }

    # ── Step 1: Group pending rows by RuleStoreName|TargetFolderPath ────────
    # Mirrors the macro's LocalPathLookupTable build exactly: only rows
    # with blank LastDeployedRun and non-blank TargetFolderPath/SendersDomain
    # are grouped. Composite key -> Dictionary of domain -> $true.
    # Added 2026-07-02, Administrator: when -RefreshRulesSortOnly is present, Steps 1-4
    # (grouping, build, consolidation, CSV writeback) are skipped entirely --
    # this run only exists to re-sort the existing collection. The resort
    # pass itself is unconditional and runs after this whole if-block,
    # regardless of which branch executes.
    if (-not $RefreshRulesSortOnly) {

    # Added 2026-07-02, Administrator (performance fix, Administrator's design). The
    # consolidation scan in Invoke-BuildRulesFromMap previously re-scanned
    # the ENTIRE live Outlook.Rules collection (~356 rules) for EVERY group
    # key being processed, walking each candidate rule's live folder path via
    # COM to check for a match -- an O(n^2) cost (~126,000 iterations on a
    # 356-rule run) that was the dominant contributor to a ~1.5-hour runtime
    # vs the macro's ~30 seconds (confirmed live -- batching Save() calls,
    # a much cheaper theory, made no measurable difference to per-rule pace).
    #
    # Administrator's insight: rules_inventory.csv IS the source of truth for which
    # existing rule currently targets which folder -- Script 00 built the CSV
    # FROM live Outlook rules in the first place (RuleName + TargetFolderPath
    # columns), and the documented pipeline contract requires Script 00 be
    # re-run whenever an admin manually adds/edits rules via the Outlook UI
    # BEFORE Script 03 runs -- so the CSV is guaranteed current when this
    # function runs. That means the "which existing rule(s) target this
    # folder" question can be answered directly from CSV data already in
    # memory, with zero live COM calls -- no need to blindly re-scan the live
    # collection at all. Only the ACTUAL DELETION of a known-by-name rule
    # still needs one live COM call, not a full scan.
    #
    # Built ONCE here from the FULL $RulesInventory (not just pending rows --
    # this must include every existing rule regardless of its own
    # LastDeployedRun status, since it's answering "what's already there").
    $folderPathToRuleNames = [ordered]@{}
    foreach ($invRow in $RulesInventory) {
        $invStoreName  = if ($invRow.PSObject.Properties['RuleStoreName'])    { [string]$invRow.RuleStoreName } else { '' }
        $invFolderPath = if ($invRow.PSObject.Properties['TargetFolderPath']) { [string]$invRow.TargetFolderPath } else { '' }
        $invRuleName   = if ($invRow.PSObject.Properties['RuleName'])         { [string]$invRow.RuleName } else { '' }
        if ([string]::IsNullOrWhiteSpace($invFolderPath) -or [string]::IsNullOrWhiteSpace($invRuleName)) { continue }
        $invKey = ($invStoreName.Trim().ToLower()) + '|' + $invFolderPath
        if (-not $folderPathToRuleNames.Contains($invKey)) {
            $folderPathToRuleNames[$invKey] = [System.Collections.Generic.List[string]]::new()
        }
        [void]$folderPathToRuleNames[$invKey].Add($invRuleName)
    }
    Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules: Built folder-to-rule-name lookup from CSV -- $($folderPathToRuleNames.Count) distinct folder(s) mapped, avoids live-collection rescanning in consolidation." -Level INFO

    Write-OMMigrateLog -Message 'Invoke-DeployConsolidatedRules: Grouping pending rules_inventory.csv rows...' -Level INFO
    $pathLookup = [ordered]@{}

    # EXTRACTED 2026-07-09 (Administrator direction) from OMMigrate-Outlook_WIP.psm1,
    # a 6-day-old (2026-07-03/04) work-in-progress file -- this is the
    # TargetStoreName hardcode fix work paused mid-implementation that
    # session. NOT a completed/finished feature -- extracted as a jumpstart
    # on resuming this work, not yet live-tested. Parallel lookup:
    # compositeGroupKey -> this group's TargetStoreName (from the CSV).
    # Built alongside $pathLookup so Invoke-BuildRulesFromMap can resolve
    # which attached store's root folder to walk from for this specific
    # group, instead of every group always walking from the one Archive
    # PST passed in via -ArchiveRootFolder. Not hardcoded to any store
    # name -- whatever value Export-RulesToCSV wrote for this row
    # (operator's ArchiveStoreMappings selection, or a live COM value) is
    # what gets used here.
    $groupKeyToTargetStore = [ordered]@{}

    # Picker-scoping fix (added 2026-07-06, Administrator) -- build a case-insensitive
    # lookup set from $ScopedAccountNames once, outside the loop, rather than
    # re-comparing the raw array on every row.
    $scopedAccountLookup = $null
    if ($ScopedAccountNames -and $ScopedAccountNames.Count -gt 0) {
        $scopedAccountLookup = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($sa in $ScopedAccountNames) { [void]$scopedAccountLookup.Add($sa.Trim()) }
        Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules: Scoped to $($scopedAccountLookup.Count) picker-selected account(s) -- rows for any other account are left untouched this run." -Level INFO
    }

    foreach ($row in $RulesInventory) {
        $timestamp    = if ($row.PSObject.Properties['LastDeployedRun']) { [string]$row.LastDeployedRun } else { '' }
        $relativePath = if ($row.PSObject.Properties['TargetFolderPath']) { [string]$row.TargetFolderPath } else { '' }
        $rawDomain    = if ($row.PSObject.Properties['SendersDomain'])    { [string]$row.SendersDomain } else { '' }
        $storeName    = if ($row.PSObject.Properties['RuleStoreName'])    { [string]$row.RuleStoreName } else { '' }

        # Picker-scoping fix (added 2026-07-06, Administrator): skip this row entirely
        # if it belongs to an account not in the picker's selection -- before
        # any grouping/normalization work, so an unscoped row can never reach
        # $pathLookup, never triggers rule creation, and never gets
        # LastDeployedRun stamped this run. No-op when $scopedAccountLookup
        # is $null (parameter not provided) -- see parameter doc above.
        if ($scopedAccountLookup -and -not $scopedAccountLookup.Contains($storeName.Trim())) {
            continue
        }

        # Normalization fix (2026-07-02, Administrator): a raw SendersDomain cell may
        # contain "@"/"&" (invalid in a TLD domain) or represent more than
        # one intended value separated by a space or semicolon (e.g.
        # "amazon aws" -- two distinct sender terms, not one malformed
        # multi-word value). ConvertTo-NormalizedSenderDomains splits/cleans
        # the raw cell into an array of individually valid domain strings --
        # each one becomes its own $pathLookup dictionary key below, so
        # Outlook's SenderAddress.Address array naturally evaluates them
        # with logical OR (confirmed via Microsoft's
        # AddressRuleCondition.Address docs).
        #
        # Fix (2026-07-02, Administrator): PowerShell's function-return unwrapping can
        # collapse a single-element array into a bare string when the
        # normalized value has only one word (e.g. "zoho") -- a bare string
        # has .Length but not .Count under Set-StrictMode, causing "The
        # property 'Count' cannot be found on this object" (confirmed live,
        # row 359/Zoho). Wrapping the CALL SITE in @() forces
        # $normalizedDomains to always be a true array regardless of how many
        # elements ConvertTo-NormalizedSenderDomains actually returned -- the
        # function's own internal "return @($result)" is not sufficient on
        # its own to prevent this unwrapping at the caller.
        $normalizedDomains = @(ConvertTo-NormalizedSenderDomains -RawValue $rawDomain)

        if ([string]::IsNullOrWhiteSpace($timestamp) -and
            -not [string]::IsNullOrWhiteSpace($relativePath) -and
            $normalizedDomains.Count -gt 0) {

            $compositeGroupKey = ($storeName.Trim().ToLower()) + '|' + $relativePath

            if (-not $pathLookup.Contains($compositeGroupKey)) {
                $pathLookup[$compositeGroupKey] = [ordered]@{}
            }
            foreach ($domain in $normalizedDomains) {
                $pathLookup[$compositeGroupKey][$domain] = $true
            }

            # EXTRACTED 2026-07-09 (Administrator direction) from OMMigrate-Outlook_WIP.psm1,
            # a 6-day-old (2026-07-03/04) work-in-progress file -- this is the
            # TargetStoreName hardcode fix work paused mid-implementation that
            # session. NOT a completed/finished feature -- extracted as a
            # jumpstart on resuming this work, not yet live-tested.
            #
            # Record this group's TargetStoreName the first time the key is
            # seen. All rows sharing a compositeGroupKey are the same
            # RuleStoreName|TargetFolderPath group, so they always carry the
            # same TargetStoreName -- only needs to be captured once per key.
            $rowTargetStoreName = if ($row.PSObject.Properties['TargetStoreName']) { [string]$row.TargetStoreName } else { '' }
            if (-not $groupKeyToTargetStore.Contains($compositeGroupKey) -and
                -not [string]::IsNullOrWhiteSpace($rowTargetStoreName)) {
                $groupKeyToTargetStore[$compositeGroupKey] = $rowTargetStoreName
            }
        }
    }

    Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules: $($pathLookup.Count) group(s) pending deployment." -Level INFO

    # EXTRACTED 2026-07-09 (Administrator direction) from OMMigrate-Outlook_WIP.psm1,
    # a 6-day-old (2026-07-03/04) work-in-progress file -- this is the
    # TargetStoreName hardcode fix work paused mid-implementation that
    # session. NOT a completed/finished feature -- extracted as a jumpstart
    # on resuming this work, not yet live-tested. Resolve each DISTINCT
    # TargetStoreName seen above to its live root folder, once, via
    # $Namespace.Stores -- not hardcoded to any single store. A group whose
    # TargetStoreName cannot be resolved (store not currently attached, or
    # blank) simply has no entry here and Invoke-BuildRulesFromMap falls
    # back to the single $ArchiveRootFolder passed into this function,
    # exactly as it did before this feature existed.
    $groupKeyToArchiveFolder = [ordered]@{}
    if ($groupKeyToTargetStore.Count -gt 0) {
        $resolvedStoreFolders = @{}
        try {
            $allStoresForLookup = $Namespace.Stores
            foreach ($groupKeyForLookup in @($groupKeyToTargetStore.Keys)) {
                $wantedStoreName = $groupKeyToTargetStore[$groupKeyForLookup]
                if ([string]::IsNullOrWhiteSpace($wantedStoreName)) { continue }

                if (-not $resolvedStoreFolders.ContainsKey($wantedStoreName)) {
                    $foundStoreFolder = $null
                    for ($lsi = 1; $lsi -le $allStoresForLookup.Count; $lsi++) {
                        $lookupStore = $allStoresForLookup.Item($lsi)
                        if ($lookupStore.DisplayName -eq $wantedStoreName) {
                            try { $foundStoreFolder = $lookupStore.GetRootFolder() } catch { }
                            break
                        }
                    }
                    if (-not $foundStoreFolder) {
                        Write-OMMigrateLog -Message (
                            "Invoke-DeployConsolidatedRules: TargetStoreName '$wantedStoreName' not found among " +
                            "attached stores -- groups mapped to it will fall back to the default Archive PST."
                        ) -Level WARN
                    }
                    $resolvedStoreFolders[$wantedStoreName] = $foundStoreFolder
                }

                if ($resolvedStoreFolders[$wantedStoreName]) {
                    $groupKeyToArchiveFolder[$groupKeyForLookup] = $resolvedStoreFolders[$wantedStoreName]
                }
            }
        }
        catch {
            Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules: Archive store resolution failed (non-fatal): $_ -- all groups will use the default Archive PST." -Level WARN
        }
    }

    # Fix (2026-07-02, Administrator): this was previously an early 'return $result',
    # which would exit the ENTIRE function -- including skipping the resort
    # pass below, which must always run regardless of whether there was any
    # pending build work. Changed to an if/else wrapping the rest of Steps
    # 2-4 so execution always falls through to the resort pass afterward.
    if ($pathLookup.Count -eq 0) {
        Write-Host '  No pending rules to deploy -- all rows already have LastDeployedRun set.' -ForegroundColor DarkGray
    }
    else {

    $domainsInThisBatch = [ordered]@{}

    # Fix (2026-07-02, Administrator): $domainsInThisBatch is keyed by individual SENDER
    # DOMAIN values, not by rule/folder identity. Step 4 below was using it to
    # decide whether a CSV ROW should be stamped with LastDeployedRun -- but two
    # unrelated rows can share an overlapping domain word (e.g. ExampleCo's
    # domain "example-provider.com" is also one of the two space-separated values in
    # ExampleCo Accounting's SendersDomain "example-provider.com accounting"). This
    # caused ExampleCo's row to be falsely stamped as deployed on a run where
    # its own rule was never actually touched (never went through
    # Invoke-BuildRulesFromMap at all), because a SIBLING row's domain happened to
    # match. $keysProcessedThisBatch tracks the actual MapKey
    # (RuleStoreName|TargetFolderPath) of every row genuinely processed this run --
    # Step 4 now checks THIS, not domain overlap, to decide whether to stamp a row.
    $keysProcessedThisBatch = [ordered]@{}

    $timestampString    = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.ffffffzzz')

    # ── Step 2: Resolve against standard account wrappers ───────────────────
    # Mirrors the macro's Step 1: iterate Application.Session.Accounts
    # equivalent, match keys whose account-binding segment matches this
    # account's SmtpAddress, DisplayName, or substring-of-DisplayName.
    $accounts = $Namespace.Session.Accounts
    for ($ai = 1; $ai -le $accounts.Count; $ai++) {
        $acct = $accounts.Item($ai)

        $hasStore = $false
        $acctDisplayForLog = ''
        try { $acctDisplayForLog = [string]$acct.DisplayName } catch { }
        try {
            if ($acct.DeliveryStore -and $acct.DeliveryStore.GetRootFolder()) { $hasStore = $true }
        } catch {
            # Logged (added 2026-07-06, Administrator): this catch previously swallowed
            # DeliveryStore resolution failures with no log line at all -- the
            # only silent DeliveryStore site left in this function without any
            # visibility, confirmed via live log analysis (2026-07-05 run) as
            # the root cause of pending CSV rows for some secondary accounts
            # never getting deployed by this function in some runs.
            Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules Step 2: failed to resolve DeliveryStore for account '$acctDisplayForLog': $_" -Level INFO
        }

        # Retry loop REMOVED (2026-07-06, Administrator, 4th pass): same rationale as
        # Get-OutlookRules -- confirmed live (memory #28) that for the known
        # group of 7 broken accounts, Account.DeliveryStore is genuinely
        # $null, not a transient timing race, so retrying never helped.
        # Going straight from the direct attempt above to the DisplayName
        # fallback below.

        # Fallback (added 2026-07-06, Administrator): DisplayName match against
        # namespace.Stores, same rationale and mechanism as the matching
        # fallback added to Get-OutlookRules (see memory #28). Only runs if
        # the direct attempt and both retries above all failed. Populates
        # $resolvedStoreForFallback so the GetRules() call below can use it
        # instead of re-deriving from $acct.DeliveryStore (which is still
        # null for these accounts) -- $acct.DeliveryStore itself cannot be
        # assigned to, so callers below must check $resolvedStoreForFallback
        # first and fall back to $acct.DeliveryStore only when it's empty.
        $resolvedStoreForFallback = $null
        if (-not $hasStore) {
            try {
                $thisAcctDisplay = ''
                try { $thisAcctDisplay = ([string]$acct.DisplayName).ToLower() } catch { }
                $thisAcctSmtpForStores = ''
                try { $thisAcctSmtpForStores = ([string]$acct.SmtpAddress).ToLower() } catch { }

                foreach ($candidateStore in $Namespace.Stores) {
                    $candidateDisplay = ''
                    try { $candidateDisplay = ([string]$candidateStore.DisplayName).ToLower() } catch { }
                    if (-not $candidateDisplay) { continue }

                    $isMatch = $false
                    if ($thisAcctDisplay -and $candidateDisplay -eq $thisAcctDisplay) { $isMatch = $true }
                    if (-not $isMatch -and $thisAcctSmtpForStores -and $candidateDisplay -eq $thisAcctSmtpForStores) { $isMatch = $true }

                    if ($isMatch) {
                        $resolvedStoreForFallback = $candidateStore
                        $hasStore = $true
                        Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules Step 2: DeliveryStore fallback -- namespace.Stores DisplayName match found for account '$acctDisplayForLog'. Looked for DisplayName='$thisAcctDisplay' or SmtpAddress='$thisAcctSmtpForStores'. Matched store: '$candidateDisplay'." -Level INFO
                        break
                    }
                }
                if (-not $hasStore) {
                    # Debug line (added 2026-07-06, Administrator, 3rd pass): matches the
                    # same no-match diagnostic added to Get-OutlookRules -- the
                    # fallback loop ran to completion but found no matching store.
                    $allStoreDisplaysForLog = @()
                    foreach ($s in $Namespace.Stores) {
                        try { $allStoreDisplaysForLog += [string]$s.DisplayName } catch { }
                    }
                    Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules Step 2: DeliveryStore fallback -- namespace.Stores DisplayName match NOT found for account '$acctDisplayForLog'. Looked for DisplayName='$thisAcctDisplay' or SmtpAddress='$thisAcctSmtpForStores'. Stores seen: $($allStoreDisplaysForLog -join ' | ')." -Level INFO
                }
            }
            catch {
                Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules Step 2: namespace.Stores DisplayName fallback also failed for account '$acctDisplayForLog': $_" -Level INFO
            }
        }

        if (-not $hasStore) {
            # Logged (added 2026-07-06, Administrator): this account is genuinely being
            # skipped by Step 2 -- any of its pending rows will only be caught
            # by Step 3's primary-store fallback if it happens to be the
            # default/primary account. Making this visible so a skip is never
            # silent again for this function.
            Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules Step 2: account '$acctDisplayForLog' has no usable DeliveryStore after retries -- skipping (any pending rows for this account will not be deployed this run unless caught by Step 3)." -Level WARN
            continue
        }

        $targetAccAddress     = ([string]$acct.SmtpAddress).Trim().ToLower()
        $targetAccDisplayName = ([string]$acct.DisplayName).Trim().ToLower()

        $accountRules = $null
        try {
            if ($resolvedStoreForFallback) {
                # Using the namespace.Stores fallback result (added 2026-07-06,
                # Administrator) -- $acct.DeliveryStore is still null for this account,
                # so GetRules() must be called on the resolved store directly.
                $accountRules = $resolvedStoreForFallback.GetRules()
            }
            else {
                $accountRules = $acct.DeliveryStore.GetRules()
            }
            # ADDED (bug found live 2026-07-12): $accountRules was never
            # registered via Register-COMObject anywhere in this function --
            # every other live COM reference in this codebase is. An
            # unregistered RCW can be released by .NET garbage collection
            # between the point it is obtained here and the point Save() is
            # called on it later in this same loop iteration (many other COM
            # calls happen in between, for other accounts' folder/store
            # resolution). Confirmed live 2026-07-12: rule creation and
            # Save() both logged success for a real rule, but the rule was
            # provably absent from the live Outlook Rules and Alerts dialog
            # immediately after, with Outlook confirmed closed during the
            # run (rules out the separate UI-cache-overwrite issue). Explicit
            # registration keeps the RCW alive for the full lifetime this
            # function needs it.
            if ($accountRules) { Register-COMObject -ComObject $accountRules }
        } catch { }
        if (-not $accountRules) {
            # Logged (added 2026-07-06, Administrator): second silent-skip path in this
            # loop -- DeliveryStore resolved above, but GetRules() itself still
            # failed. Previously invisible.
            Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules Step 2: DeliveryStore resolved for account '$acctDisplayForLog' but GetRules() failed or returned nothing -- skipping." -Level WARN
            continue
        }

        # Reverse iteration over group keys -- same technique and same
        # reason as the macro: Outlook appends new rules in creation
        # order, so processing Z-to-A here makes the rules appear in
        # clean A-to-Z order top-to-bottom in the Rules and Alerts UI.
        #
        # Sort fix (2026-07-02, Administrator): $pathLookup.Keys was previously used
        # in raw insertion order, which is inherited from rules_inventory.csv's
        # RuleStoreName+TargetFolderPath sort (Invoke-NormalizeRulesExecutionOrder).
        # That sort groups rules by folder TREE structure (parent folder
        # immediately followed by its own subfolders), not by the final
        # display name -- confirmed live: rules targeting nested subfolders
        # (e.g. "ExampleCo Accounting\aws") clustered under their
        # parent instead of falling into flat alphabetical position by their
        # own label ("aws" should sort between "AWE" and "Axelon", not
        # under "ExampleCo"). This is a DIFFERENT sort than the
        # TargetFolderPath-based grouping/consolidation match key above --
        # that match key is untouched. This sort only controls the order
        # rules are CREATED in, so the reverse-iteration trick below lands
        # them in true flat A-to-Z order in the Rules and Alerts UI by the
        # final rule label (leaf folder name / standardized rule name),
        # matching how Invoke-NormalizeRulesExecutionOrder's docstring
        # always described the intended behavior.
        $allKeys = @($pathLookup.Keys | Sort-Object {
            $sortSegments = ($_ -split '\|', 2)
            if ($sortSegments.Count -ge 2) {
                $sortPathSegments = $sortSegments[1] -split '\\'
                $sortPathSegments[-1].Trim().ToLower()
            } else {
                ([string]$_).Trim().ToLower()
            }
        })
        for ($mapIdx = $allKeys.Count - 1; $mapIdx -ge 0; $mapIdx--) {
            $mapKey = $allKeys[$mapIdx]
            $keySegments         = $mapKey -split '\|', 2
            $ruleAccountBinding  = $keySegments[0].Trim().ToLower()
            $targetPath          = $keySegments[1]

            $accountMatches = $false
            if ($ruleAccountBinding -eq $targetAccAddress) { $accountMatches = $true }
            elseif ($ruleAccountBinding -eq $targetAccDisplayName) { $accountMatches = $true }
            elseif ($targetAccDisplayName -and $targetAccDisplayName.Contains($ruleAccountBinding)) { $accountMatches = $true }

            if ($accountMatches) {
                # EXTRACTED 2026-07-09 (Administrator direction) from OMMigrate-Outlook_WIP.psm1
                # -- pass this group's own resolved store folder if one exists
                # ($groupKeyToArchiveFolder is keyed identically to $mapKey/
                # $compositeGroupKey), otherwise $null so Invoke-BuildRulesFromMap
                # falls back to $ArchiveRootFolder exactly as before this feature.
                $overrideFolderForThisKey = if ($groupKeyToArchiveFolder.Contains($mapKey)) { $groupKeyToArchiveFolder[$mapKey] } else { $null }
                Invoke-BuildRulesFromMap `
                    -TargetRules         $accountRules `
                    -PathLookup          $pathLookup `
                    -MapKey              $mapKey `
                    -TargetPath          $targetPath `
                    -RuleAccountBinding  $ruleAccountBinding `
                    -ActiveAccount       $acct `
                    -ArchiveRootFolder   $ArchiveRootFolder `
                    -DomainsInThisBatch  $domainsInThisBatch `
                    -Result              $result `
                    -SaveBatchSize       $SaveBatchSize `
                    -FolderPathToRuleNames $folderPathToRuleNames `
                    -OverrideArchiveFolder $overrideFolderForThisKey
                # Fix (2026-07-02, Administrator): track this ROW's own key as genuinely
                # processed -- see $keysProcessedThisBatch declaration above.
                $keysProcessedThisBatch[$mapKey] = $true
            }
            # 1000-action watchdog -- mirrors the macro's hard limit to avoid
            # exceeding Outlook's MAPI rule storage capacity in one run.
            if ($result.RuleActionCounter -ge 1000) { break }
        }
        # Flush any remaining pending batched Save() for THIS account's
        # collection before moving to the next account -- see
        # Invoke-FlushPendingRuleSave header comment.
        Invoke-FlushPendingRuleSave -TargetRules $accountRules
        if ($result.RuleActionCounter -ge 1000) { break }
    }

    # ── Step 3: Primary store fallback (default/primary account fix) ───────
    # Mirrors the macro's Step 2 exactly: any remaining default/primary-store-
    # bound group keys not yet handled above (e.g. the default/primary store
    # has no matching Account wrapper with its own DeliveryStore) get
    # built against Session.DefaultStore.GetRules() instead.
    $defaultStoreRules = $null
    try { $defaultStoreRules = $Namespace.Session.DefaultStore.GetRules() } catch { }

    if ($defaultStoreRules -and $result.RuleActionCounter -lt 1000) {
        # Resolve the primary/default account generically -- by matching
        # against $Namespace.Session.DefaultStore itself, not by hardcoding
        # any specific email provider or domain name. Works for any user's
        # environment regardless of which provider their primary account uses.
        $primaryAccObj      = $null
        $primaryAccAddress  = ''
        $primaryAccDisplay  = ''
        try {
            $defaultStoreId = $Namespace.Session.DefaultStore.StoreID
            for ($ai = 1; $ai -le $accounts.Count; $ai++) {
                $acct = $accounts.Item($ai)
                $acctStoreId = ''
                try { $acctStoreId = $acct.DeliveryStore.StoreID } catch { }
                if ($acctStoreId -and $acctStoreId -eq $defaultStoreId) {
                    $primaryAccObj     = $acct
                    $primaryAccAddress = ([string]$acct.SmtpAddress).ToLower()
                    $primaryAccDisplay = ([string]$acct.DisplayName).ToLower()
                    break
                }
            }

            # Fallback (added 2026-07-01, Administrator): Account.DeliveryStore does not
            # reliably resolve for every account in every profile -- same root
            # cause confirmed 2026-07-01 in Get-OutlookRules ("TestProfile" profile,
            # ameritech account). When that happens here, $acctStoreId never
            # matches $defaultStoreId for ANY account, so $primaryAccObj stays
            # $null and the primary/default account's pending rule rows are
            # silently never processed in this Step 3 fallback either --
            # confirmed live: secondary accounts got LastDeployedRun written,
            # ameritech (the actual default store) did not.
            #
            # If the DeliveryStore-based match above found nothing, retry by
            # comparing each account's SmtpAddress/DisplayName against
            # $Namespace.Session.DefaultStore's own DisplayName -- the same
            # store handle already proven to resolve correctly elsewhere
            # (Get-OutlookRules fallback, a diagnostic post-import verification test script).
            # This avoids DeliveryStore entirely for identifying the primary
            # account, instead of retrying the same broken property.
            if (-not $primaryAccObj) {
                $defaultStoreDisplay = ''
                try { $defaultStoreDisplay = ([string]$Namespace.Session.DefaultStore.DisplayName).ToLower() } catch { }
                if ($defaultStoreDisplay) {
                    for ($ai2 = 1; $ai2 -le $accounts.Count; $ai2++) {
                        $acct2 = $accounts.Item($ai2)
                        $acct2Address = ([string]$acct2.SmtpAddress).ToLower()
                        $acct2Display = ([string]$acct2.DisplayName).ToLower()
                        if (($acct2Address -and $defaultStoreDisplay.Contains($acct2Address)) -or
                            ($acct2Display -and $defaultStoreDisplay.Contains($acct2Display)) -or
                            ($acct2Display -and $acct2Display -eq $defaultStoreDisplay)) {
                            $primaryAccObj     = $acct2
                            $primaryAccAddress = $acct2Address
                            $primaryAccDisplay = $acct2Display
                            Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules: primary account resolved via DefaultStore.DisplayName fallback (DeliveryStore-based match failed)." -Level INFO
                            break
                        }
                    }
                }
            }
        } catch { }

        # Sort fix (2026-07-02, Administrator): same fix as Step 2 above -- sort by
        # leaf folder label so reverse iteration lands rules in true flat
        # A-to-Z order by their final display name, not by folder-tree
        # position inherited from $pathLookup's raw key order.
        $remainingKeys = @($pathLookup.Keys | Sort-Object {
            $remSortSegments = ($_ -split '\|', 2)
            if ($remSortSegments.Count -ge 2) {
                $remSortPathSegments = $remSortSegments[1] -split '\\'
                $remSortPathSegments[-1].Trim().ToLower()
            } else {
                ([string]$_).Trim().ToLower()
            }
        })
        for ($remIdx = $remainingKeys.Count - 1; $remIdx -ge 0; $remIdx--) {
            $remainingKey  = $remainingKeys[$remIdx]
            $remSegments   = $remainingKey -split '\|', 2
            $remAccountBinding = $remSegments[0].Trim().ToLower()
            $remTargetPath     = $remSegments[1]

            # Domain-level dedup -- mirrors the macro's INTENT (skip a group
            # already fully handled in Step 2), but checks the individual
            # domain keys belonging to THIS group rather than the account
            # binding string. $domainsInThisBatch is keyed by domain (e.g.
            # "example.com"), never by account binding, so comparing it
            # against $remAccountBinding directly (as the macro does) can
            # never match. Checking whether every domain in this group is
            # already in $domainsInThisBatch correctly detects "Step 2
            # already created rules for all of these domains" regardless of
            # whether the primary account ever actually matches an Account
            # wrapper in Step 2 in this environment.
            $remGroupDomains = @($pathLookup[$remainingKey].Keys)
            $remAlreadyHandled = $remGroupDomains.Count -gt 0
            foreach ($remDomain in $remGroupDomains) {
                if (-not $domainsInThisBatch.Contains($remDomain)) {
                    $remAlreadyHandled = $false
                    break
                }
            }

            # Generic match against the resolved primary account's address or
            # display name -- no hardcoded provider/domain name. Matches the
            # same logic Step 2 uses for every other account (exact address,
            # exact display name, or substring-of-display-name).
            $remIsPrimaryAccount = $false
            if ($primaryAccAddress -and $remAccountBinding -eq $primaryAccAddress) { $remIsPrimaryAccount = $true }
            elseif ($primaryAccDisplay -and $remAccountBinding -eq $primaryAccDisplay) { $remIsPrimaryAccount = $true }
            elseif ($primaryAccDisplay -and $primaryAccDisplay.Contains($remAccountBinding)) { $remIsPrimaryAccount = $true }

            if ($remIsPrimaryAccount -and -not $remAlreadyHandled) {
                # EXTRACTED 2026-07-09 (Administrator direction) from OMMigrate-Outlook_WIP.psm1
                # -- same override-folder lookup as Step 2's call site above,
                # keyed by $remainingKey (this loop's equivalent of $mapKey).
                $overrideFolderForThisKey = if ($groupKeyToArchiveFolder.Contains($remainingKey)) { $groupKeyToArchiveFolder[$remainingKey] } else { $null }
                Invoke-BuildRulesFromMap `
                    -TargetRules         $defaultStoreRules `
                    -PathLookup          $pathLookup `
                    -MapKey              $remainingKey `
                    -TargetPath          $remTargetPath `
                    -RuleAccountBinding  $remAccountBinding `
                    -ActiveAccount       $primaryAccObj `
                    -ArchiveRootFolder   $ArchiveRootFolder `
                    -DomainsInThisBatch  $domainsInThisBatch `
                    -Result              $result `
                    -SaveBatchSize       $SaveBatchSize `
                    -FolderPathToRuleNames $folderPathToRuleNames `
                    -OverrideArchiveFolder $overrideFolderForThisKey
                # Fix (2026-07-02, Administrator): track this ROW's own key as genuinely
                # processed -- see $keysProcessedThisBatch declaration above.
                $keysProcessedThisBatch[$remainingKey] = $true
            }
            # 1000-action watchdog -- same hard limit as Step 2 above.
            if ($result.RuleActionCounter -ge 1000) { break }
        }
        # Flush any remaining pending batched Save() for the default store's
        # collection -- see Invoke-FlushPendingRuleSave header comment.
        Invoke-FlushPendingRuleSave -TargetRules $defaultStoreRules
    }

    # ── Step 4: Write LastDeployedRun timestamps back to CSV ────────────────
    # Mirrors the macro's WriteTimestampsToCSV: every domain actually
    # processed in this run gets the same shared timestamp written to
    # its row's LastDeployedRun column. Rows whose domain was not
    # processed (no matching account/store found) are left blank so a
    # future run -- by either this function or Gemini's macro -- retries them.
    if ($domainsInThisBatch.Count -gt 0) {
        try {
            $allCsvRows = Import-Csv -Path $CsvPath -Encoding UTF8
            # ADDED (Administrator direction, 2026-07-21, THIRD pass): groups CSV rows by
            # their MapKey (RuleStoreName|TargetFolderPath) as this loop runs, so the
            # chunk-assignment pass immediately after can distribute this run's
            # actual chunks across them -- see that pass's own comment below for
            # full rationale.
            $rowGroupsForChunkAssignment = [ordered]@{}
            foreach ($csvRow in $allCsvRows) {
                # Fix (2026-07-02, Administrator): $rowDomainMatched was previously based on
                # whether ANY of this row's normalized SendersDomain words appeared
                # anywhere in the whole batch's domain pool ($domainsInThisBatch).
                # This is the WRONG identity check -- SendersDomain is a search
                # condition, not a row identifier, and two unrelated rows can share
                # an overlapping domain word (e.g. "ExampleCo" vs "ExampleCo
                # Accounting", whose SendersDomain "example-provider.com accounting"
                # normalizes to TWO words, one of which collides with ExampleCo's
                # own domain). This caused ExampleCo's row to be falsely stamped
                # as deployed on a run where its own rule was never actually
                # processed. The correct check is whether THIS ROW's own identity key
                # (RuleStoreName|TargetFolderPath, the same MapKey format used
                # throughout Step 1-3) was genuinely processed this run -- rebuild
                # that same key here and check $keysProcessedThisBatch instead.
                $rowStoreName      = if ($csvRow.PSObject.Properties['RuleStoreName']) { [string]$csvRow.RuleStoreName } else { '' }
                $rowFolderPath     = if ($csvRow.PSObject.Properties['TargetFolderPath']) { [string]$csvRow.TargetFolderPath } else { '' }
                $rowCompositeKey   = $rowStoreName.Trim().ToLower() + '|' + $rowFolderPath
                $rowDomainMatched  = $keysProcessedThisBatch.Contains($rowCompositeKey)
                $rowTimestamp = if ($csvRow.PSObject.Properties['LastDeployedRun']) {
                    [string]$csvRow.LastDeployedRun
                } else { '' }

                if ($rowDomainMatched -and [string]::IsNullOrWhiteSpace($rowTimestamp)) {
                    $csvRow.LastDeployedRun = $timestampString
                    $result.DomainsProcessed++
                }

                $rowOldNameForRename = if ($csvRow.PSObject.Properties['RuleName']) { [string]$csvRow.RuleName } else { '' }
                $rowMapKeyForRename  = $rowStoreName.Trim().ToLower() + '|' + $rowFolderPath
                if (-not $rowGroupsForChunkAssignment.Contains($rowMapKeyForRename)) {
                    $rowGroupsForChunkAssignment[$rowMapKeyForRename] = [System.Collections.Generic.List[object]]::new()
                }
                [void]$rowGroupsForChunkAssignment[$rowMapKeyForRename].Add($csvRow)
            }

            # FIXED (Administrator direction, 2026-08-18 -- single-account-per-folder chunk
            # assignment bug): $rowGroupsForChunkAssignment above (and the chunk
            # filter immediately below it) grouped/matched purely by
            # RuleStoreName|TargetFolderPath -- but Invoke-BuildRulesFromMap's own
            # MapKey (see that function's own parameter doc) is
            # RuleAccountBinding|TargetPath, i.e. scoped to the ACCOUNT the rule's
            # sender-domain group belongs to, not just the folder it targets. It is
            # entirely normal for several different accounts (e.g.
            # admin@example-provider.com, admin2@example.com,
            # user@ameritech.net) to each have their own consolidated rule
            # group targeting the SAME folder under one RuleStoreName (e.g. every
            # inbound "Github" mail across several accounts filed into the same
            # ameritech\Inbox\Github archive folder). The folder-only grouping
            # above pools ALL of those accounts' CSV rows and ALL of their chunks
            # together, so rows and chunks silently swap accounts or get left
            # unassigned whenever an account's row-count and chunk-count don't
            # individually line up -- confirmed live 2026-08-18: 3 of several dozen
            # rows across the ameritech store stayed permanently unstamped despite
            # rerun after rerun, all three sharing a folder with sibling accounts.
            #
            # $rowGroupsForChunkAssignmentByAccount below re-groups the SAME rows
            # by a true RuleStoreName|RuleAccountBinding|TargetFolderPath key so
            # each account's chunks only ever compete with that SAME account's own
            # rows. The account segment is recovered from the row's own RuleName --
            # every standardized name already carries it in the leading
            # "Rule: [account] ..." bracket (same bracket already parsed elsewhere
            # in this file, e.g. the leaf-folder regex above, just capturing the
            # other group). A row whose RuleName does not yet follow the
            # standardized pattern (never consolidated before) has no bracket to
            # recover -- it falls into an account-less bucket keyed only by
            # RuleStoreName|TargetFolderPath, identical to today's behavior, so
            # first-time consolidation is unaffected by this fix.
            $rowGroupsForChunkAssignmentByAccount = [ordered]@{}
            foreach ($acctCsvRow in $allCsvRows) {
                $acctRowStoreName  = if ($acctCsvRow.PSObject.Properties['RuleStoreName'])    { [string]$acctCsvRow.RuleStoreName }    else { '' }
                $acctRowFolderPath = if ($acctCsvRow.PSObject.Properties['TargetFolderPath'])  { [string]$acctCsvRow.TargetFolderPath } else { '' }
                $acctRowRuleName   = if ($acctCsvRow.PSObject.Properties['RuleName'])          { [string]$acctCsvRow.RuleName }         else { '' }
                $acctRowBinding    = ''
                if ($acctRowRuleName -match '^Rule: \[([^\]]*)\]') {
                    $acctRowBinding = $Matches[1].Trim().ToLower()
                }
                else {
                    # FIXED (Administrator direction, 2026-08-21 -- first-time-consolidation
                    # rename bug): a row that has NEVER been consolidated before (raw
                    # Outlook UI name, no "Rule: [account] ..." bracket to parse)
                    # previously fell into an ACCOUNT-LESS bucket here (blank
                    # $acctRowBinding), but the brand-new chunk name
                    # Invoke-BuildRulesFromMap just created for it ALWAYS has a bracket
                    # (every chunk this function creates is "Rule: [account] ...", see
                    # $ruleAccountBinding's own derivation above -- $RuleStoreName,
                    # lowercased -- which is the SAME source available on this row right
                    # here). An account-less old row could therefore never match a
                    # bracketed new chunk on its first consolidation pass -- confirmed
                    # live (2026-08-21): RuleName stayed at the pre-consolidation raw
                    # name ('GitHub') forever despite the live Outlook rule being
                    # correctly renamed, because $chunkNamesForThisGroup.Count was
                    # always 0 for these rows and the loop below silently 'continue'd
                    # past them every run. Falling back to this row's own
                    # RuleStoreName (same value, same lowercasing, as
                    # $ruleAccountBinding uses) makes an account-less row's group key
                    # carry the real account identity, so it can match the new chunk
                    # exactly like an already-consolidated row does. NOTE: this fixes
                    # the CSV RuleName write-back only -- it does not touch the live
                    # Outlook rule create/delete/save sequence, which is a separate,
                    # still-open investigation as of 2026-08-21.
                    $acctRowBinding = $acctRowStoreName.Trim().ToLower()
                }
                $acctRowKey = $acctRowStoreName.Trim().ToLower() + '|' + $acctRowBinding + '|' + $acctRowFolderPath
                if (-not $rowGroupsForChunkAssignmentByAccount.Contains($acctRowKey)) {
                    $rowGroupsForChunkAssignmentByAccount[$acctRowKey] = [System.Collections.Generic.List[object]]::new()
                }
                [void]$rowGroupsForChunkAssignmentByAccount[$acctRowKey].Add($acctCsvRow)
            }
            # FIXED AGAIN (Administrator direction, 2026-07-21, THIRD pass -- Part 1/Part 2
            # RuleName/SendersDomain/Conditions correction, corrected approach): the
            # two earlier per-row matching attempts (old-name, then domain-token)
            # both assumed a CSV row's own content could identify which NEW chunk
            # it belongs to. Live testing proved this false: consolidation merges
            # every row in a folder group into one flat domain pool BEFORE
            # re-splitting into chunks of 5, so a single row's original domains can
            # end up split across two different new chunks -- there is no longer a
            # reliable 1:1 row-to-chunk mapping to detect from a row's own content.
            # Per Administrator's explicit direction (row/domain order does not matter,
            # since SendersDomain is consumed as a set, not tied to any specific
            # row's identity) and Administrator's explicit scope (only RuleName,
            # SendersDomain, and Conditions need correcting, sourced from the live
            # rule truth): this second pass distributes each of THIS RUN's actual
            # final chunks (from ConsolidatedDomainToNewName's distinct chunk names,
            # scoped to this same folder group) across the CSV rows already grouped
            # above by MapKey, one chunk per row, in whatever order -- and writes
            # that chunk's authoritative RuleName/SendersDomain/Conditions onto it.
            # A group with only 1 chunk (the common, non-split case) simply
            # assigns its one row unchanged in effect. A group with N rows and M
            # chunks (M != N is possible after a re-split) assigns min(N, M) rows;
            # any leftover chunks or rows are logged, never silently dropped.
            # FIXED (Administrator direction, 2026-08-18 -- single-account-per-folder chunk
            # assignment bug): was `foreach ($chunkGroupKey in @($rowGroupsForChunkAssignment.Keys))`
            # with a path-only chunk filter -- see $rowGroupsForChunkAssignmentByAccount's
            # declaration above for full rationale. Iterates the account-scoped
            # grouping instead, and the chunk filter now also requires the chunk's
            # own bracketed account (parsed the same way $acctRowBinding was above)
            # to match this group's account segment -- an account-less group
            # (never-consolidated row, no bracket to parse) matches only chunks
            # that are themselves account-less, preserving prior behavior for that
            # case exactly.
            foreach ($chunkGroupKey in @($rowGroupsForChunkAssignmentByAccount.Keys)) {
                $chunkGroupKeySegments  = $chunkGroupKey -split '\|', 3
                $chunkGroupKeyAccount   = if ($chunkGroupKeySegments.Count -ge 2) { $chunkGroupKeySegments[1] } else { '' }
                $chunkGroupKeyPath      = if ($chunkGroupKeySegments.Count -ge 3) { $chunkGroupKeySegments[2] } else { '' }
                $chunkNamesForThisGroup = @(
                    $Result.ConsolidatedDomainToNewName.Values |
                    Where-Object {
                        $chunkCandidateName = $_
                        $chunkCandidateAccount = ''
                        if ($chunkCandidateName -match '^Rule: \[([^\]]*)\]') {
                            $chunkCandidateAccount = $Matches[1].Trim().ToLower()
                        }
                        $Result.ConsolidatedRuleTargets.ContainsKey($chunkCandidateName) -and
                        ($chunkGroupKeyPath -eq $Result.ConsolidatedRuleTargets[$chunkCandidateName]) -and
                        ($chunkCandidateAccount -eq $chunkGroupKeyAccount)
                    } |
                    Select-Object -Unique
                )
                if ($chunkNamesForThisGroup.Count -eq 0) { continue }
                $rowsForThisGroup = $rowGroupsForChunkAssignmentByAccount[$chunkGroupKey]
                $assignCount = [Math]::Min($rowsForThisGroup.Count, $chunkNamesForThisGroup.Count)
                for ($assignIdx = 0; $assignIdx -lt $assignCount; $assignIdx++) {
                    $assignChunkName = $chunkNamesForThisGroup[$assignIdx]
                    $assignRow       = $rowsForThisGroup[$assignIdx]
                    $assignRow.RuleName = $assignChunkName
                    if ($Result.ConsolidatedChunkSendersDomain.ContainsKey($assignChunkName)) {
                        $assignRow.SendersDomain = $Result.ConsolidatedChunkSendersDomain[$assignChunkName]
                        if ($assignRow.PSObject.Properties['Conditions']) {
                            # FIXED (Administrator direction, 2026-08-21 -- Conditions column
                            # accuracy fix): previously hardcoded "Sender address: ..." here,
                            # unconditionally discarding any preserved Subject/Body/etc.
                            # condition text even when the LIVE rule correctly kept it --
                            # see ConsolidatedChunkConditionsText's declaration at $result's
                            # initialization for full rationale. Now uses the true, complete
                            # summary captured directly from the live rule right after
                            # creation, same as a normal Script 00 rescan would produce --
                            # falls back to the old sender-address-only string only in the
                            # unexpected case that capture didn't happen for this chunk
                            # (e.g. Get-RuleConditionsSummary itself failed), so this can
                            # never regress to blank/missing Conditions text.
                            if ($Result.ConsolidatedChunkConditionsText.ContainsKey($assignChunkName)) {
                                $assignRow.Conditions = $Result.ConsolidatedChunkConditionsText[$assignChunkName]
                            } else {
                                $assignRow.Conditions = "Sender address: $($Result.ConsolidatedChunkSendersDomain[$assignChunkName] -replace ' ', ', ')"
                            }
                        }
                    }
                }
                if ($rowsForThisGroup.Count -ne $chunkNamesForThisGroup.Count) {
                    Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules: folder group '$chunkGroupKey' had $($rowsForThisGroup.Count) CSV row(s) but $($chunkNamesForThisGroup.Count) chunk(s) this run -- assigned $assignCount, remainder will self-correct on the next Script 00/03 pass." -Level INFO
                }
            }
            # ADDED (Administrator direction, 2026-08-18): re-insert blank separator rows
            # between RuleStoreName groups before writing -- see Add-RulesCsvSeparatorRows
            # header comment for full rationale. $allCsvRows above is unfiltered (may
            # already contain old separator rows from the last write), so filter to
            # real data rows first -- same filter pattern used at every other rules_
            # inventory.csv write site in this project -- before handing off to the
            # helper, which does its own fresh sort and re-insertion.
            $allCsvRowsForWrite = @($allCsvRows | Where-Object {
                $_.RuleStoreName -and -not [string]::IsNullOrWhiteSpace($_.RuleStoreName) -and
                $_.RuleName      -and -not [string]::IsNullOrWhiteSpace($_.RuleName)
            })
            $allCsvRowsForWrite = Add-RulesCsvSeparatorRows -Rows $allCsvRowsForWrite
            $allCsvRowsForWrite | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
            Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules: LastDeployedRun written for $($result.DomainsProcessed) domain(s)." -Level INFO
        }
        catch {
            Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules: Failed to write LastDeployedRun timestamps: $_" -Level WARN
        }
    }

    Write-Host "  Consolidated rules deployed: $($result.Created) created, $($result.Failed) failed, $($result.DomainsProcessed) domain(s) timestamped." -ForegroundColor Green
    Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules: Complete -- Created=$($result.Created) Failed=$($result.Failed) DomainsProcessed=$($result.DomainsProcessed)." -Level INFO

    } # end else (pathLookup.Count -gt 0) -- closes the Fix (2026-07-02) block above
    } # end if (-not $RefreshRulesSortOnly) -- closes the Steps 1-4 skip block

    # ── Step 5: Resort every account's full rules collection by label ───────
    # Added 2026-07-02, Administrator. Runs UNCONDITIONALLY -- both in normal mode
    # (after Steps 1-4 build/consolidate any pending rows) and in
    # -RefreshRulesSortOnly mode (Steps 1-4 skipped entirely above). This
    # satisfies Administrator's requirement that the full rule collection is always
    # correctly sorted every run, whether 0, 1, or all CSV rows were
    # pending -- via Invoke-ResortRulesByLabel, which only ever writes
    # ExecutionOrder on EXISTING rules (never Create/Remove), so any
    # manual customization an end user has made to a rule's conditions or
    # actions via the Outlook UI is never touched or discarded by this step.
    Write-OMMigrateLog -Message 'Invoke-DeployConsolidatedRules: Resorting all rules by label...' -Level INFO
    $resortAccounts = $Namespace.Session.Accounts
    for ($sai = 1; $sai -le $resortAccounts.Count; $sai++) {
        $sortAcct = $resortAccounts.Item($sai)
        $sortAcctRules = $null
        $sortAcctDisplayForLog = ''
        try { $sortAcctDisplayForLog = [string]$sortAcct.DisplayName } catch { }

        # Picker-scoping fix (added 2026-07-06, Administrator): for pipeline
        # consistency with Step 1's scoping above -- per Administrator's explicit
        # direction, picker selection must be authoritative for EVERY piece
        # of activity this function performs, not just rule creation. This
        # resort pass was previously the one remaining place in this
        # function that touched every account in the profile regardless of
        # $ScopedAccountNames. Skipping an unscoped account here means its
        # rules may remain out of alphabetical order until it is explicitly
        # picked in a future run -- an acceptable, expected consequence of
        # scoping, not a bug, consistent with how Step 1 already treats an
        # unscoped account's pending CSV rows. No-op when
        # $scopedAccountLookup is $null (parameter not provided).
        if ($scopedAccountLookup -and -not $scopedAccountLookup.Contains($sortAcctDisplayForLog.Trim())) {
            continue
        }

        try {
            if ($sortAcct.DeliveryStore -and $sortAcct.DeliveryStore.GetRootFolder()) {
                $sortAcctRules = $sortAcct.DeliveryStore.GetRules()
            }
        } catch {
            # Logged (added 2026-07-06, Administrator): same silent-skip pattern fixed
            # elsewhere in this function -- this resort pass had zero
            # visibility into DeliveryStore failures until now.
            Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules resort pass: failed to resolve DeliveryStore for account '$sortAcctDisplayForLog': $_" -Level INFO
        }

        # Retry loop REMOVED (2026-07-06, Administrator, 4th pass): same rationale as
        # Get-OutlookRules and Step 2 -- confirmed live (memory #28) that for
        # the known group of 7 broken accounts, Account.DeliveryStore is
        # genuinely $null, not a transient timing race, so retrying never
        # helped. Going straight from the direct attempt above to the
        # DisplayName fallback below.

        # Fallback (added 2026-07-06, Administrator): DisplayName match against
        # namespace.Stores, same rationale and mechanism as the matching
        # fallbacks added to Get-OutlookRules and Step 2 above (see memory
        # #28). Only runs if the direct attempt and both retries above all
        # failed. Populates $sortAcctRules directly, same as the successful
        # paths above, so the existing "if ($sortAcctRules)" check below
        # picks it up with no further changes needed.
        if (-not $sortAcctRules) {
            try {
                $thisSortAcctDisplay = ''
                try { $thisSortAcctDisplay = ([string]$sortAcct.DisplayName).ToLower() } catch { }
                $thisSortAcctSmtpForStores = ''
                try { $thisSortAcctSmtpForStores = ([string]$sortAcct.SmtpAddress).ToLower() } catch { }

                foreach ($candidateStore in $Namespace.Stores) {
                    $candidateDisplay = ''
                    try { $candidateDisplay = ([string]$candidateStore.DisplayName).ToLower() } catch { }
                    if (-not $candidateDisplay) { continue }

                    $isMatch = $false
                    if ($thisSortAcctDisplay -and $candidateDisplay -eq $thisSortAcctDisplay) { $isMatch = $true }
                    if (-not $isMatch -and $thisSortAcctSmtpForStores -and $candidateDisplay -eq $thisSortAcctSmtpForStores) { $isMatch = $true }

                    if ($isMatch) {
                        try {
                            $sortAcctRules = $candidateStore.GetRules()
                        } catch { }
                        if ($sortAcctRules) {
                            Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules resort pass: DeliveryStore fallback -- namespace.Stores DisplayName match found for account '$sortAcctDisplayForLog'. Looked for DisplayName='$thisSortAcctDisplay' or SmtpAddress='$thisSortAcctSmtpForStores'. Matched store: '$candidateDisplay'." -Level INFO
                            break
                        }
                    }
                }
                if (-not $sortAcctRules) {
                    # Debug line (added 2026-07-06, Administrator, 3rd pass): matches the
                    # same no-match diagnostic added to Get-OutlookRules and
                    # Step 2 -- the fallback loop ran to completion but found no
                    # matching store (or a match was found but GetRules() on it
                    # still returned nothing).
                    $allStoreDisplaysForLog = @()
                    foreach ($s in $Namespace.Stores) {
                        try { $allStoreDisplaysForLog += [string]$s.DisplayName } catch { }
                    }
                    Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules resort pass: DeliveryStore fallback -- namespace.Stores DisplayName match NOT found for account '$sortAcctDisplayForLog'. Looked for DisplayName='$thisSortAcctDisplay' or SmtpAddress='$thisSortAcctSmtpForStores'. Stores seen: $($allStoreDisplaysForLog -join ' | ')." -Level INFO
                }
            }
            catch {
                Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules resort pass: namespace.Stores DisplayName fallback also failed for account '$sortAcctDisplayForLog': $_" -Level INFO
            }
        }

        if (-not $sortAcctRules) {
            # Logged (added 2026-07-06, Administrator): previously a completely silent
            # skip -- this account's rules simply never got resorted, with no
            # trace in the log at all.
            Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules resort pass: account '$sortAcctDisplayForLog' has no usable DeliveryStore after retries -- resort skipped for this account." -Level INFO
        }
        if ($sortAcctRules) {
            $sortCount = Invoke-ResortRulesByLabel -TargetRules $sortAcctRules
            if ($sortCount -ge 0) {
                Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules: Resorted $sortCount rule(s) for account '$($sortAcct.DisplayName)'." -Level INFO
            }
        }
    }
    # Also resort the default store directly, same as Step 3's primary-store
    # fallback pattern elsewhere in this function -- covers the case where
    # the primary account's DeliveryStore never resolves via the loop above
    # (same known fallback need documented throughout this module).
    try {
        $sortDefaultStoreRules = $Namespace.Session.DefaultStore.GetRules()
        if ($sortDefaultStoreRules) {
            $sortCount = Invoke-ResortRulesByLabel -TargetRules $sortDefaultStoreRules
            if ($sortCount -ge 0) {
                Write-OMMigrateLog -Message "Invoke-DeployConsolidatedRules: Resorted $sortCount rule(s) for default store." -Level INFO
            }
        }
    } catch { }

    return $result
}

# ============================================================
#  HELPER: Invoke-BuildRulesFromMap
#  PowerShell port of Gemini's BuildRulesFromMap VBA sub.
#  Builds one or more chunked rules (5 domains each) for a single
#  RuleStoreName|TargetFolderPath group, against the given rules
#  collection, and saves each rule immediately after creation --
#  same incremental Save() timing as the macro.
# ============================================================

# ============================================================
#  HELPER: Invoke-FlushPendingRuleSave
#  Added 2026-07-02, Administrator. Companion to Invoke-BuildRulesFromMap's
#  SaveBatchSize batching -- flushes any remainder rules that were
#  created but not yet saved because they didn't land exactly on a
#  batch boundary (e.g. SaveBatchSize=10 but only 7 rules were
#  created against this collection). Must be called once after each
#  loop that owns a given $TargetRules reference finishes (Step 2's
#  per-account loop, Step 3's fallback loop) -- otherwise the final
#  few rules of an odd-sized batch would never be committed to disk.
#  Safe/no-op if nothing is pending for the given collection.
# ============================================================

function Invoke-FlushPendingRuleSave {
    <#
    .SYNOPSIS
        Flushes any pending unsaved rules for a given Outlook.Rules
        collection, per the batching started in Invoke-BuildRulesFromMap.

    .PARAMETER TargetRules
        The same Outlook.Rules collection reference passed to
        Invoke-BuildRulesFromMap as -TargetRules.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [object] $TargetRules
    )

    # Fix (2026-07-02, Administrator): under Set-StrictMode, merely REFERENCING
    # $script:PendingSaveCounts before it has ever been assigned throws
    # ("variable cannot be retrieved because it has not been set"), not just
    # returning $null/falsy -- confirmed live. This can happen legitimately
    # if this flush is called for an account that never matched any pending
    # keys, so Invoke-BuildRulesFromMap's own lazy-init never ran. Guard with
    # Get-Variable -Scope Script -ErrorAction SilentlyContinue instead of a
    # direct $script:PendingSaveCounts reference.
    $pendingVar = Get-Variable -Name 'PendingSaveCounts' -Scope Script -ErrorAction SilentlyContinue
    if (-not $pendingVar -or -not $pendingVar.Value) { return }

    $rulesKey = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($TargetRules)
    if ($script:PendingSaveCounts.ContainsKey($rulesKey) -and $script:PendingSaveCounts[$rulesKey] -gt 0) {
        try {
            [void]$TargetRules.GetType().InvokeMember(
                'Save',
                [System.Reflection.BindingFlags]::InvokeMethod,
                $null, $TargetRules, @($false)
            )
            Write-OMMigrateLog -Message "Batch flush: saved $($script:PendingSaveCounts[$rulesKey]) remaining pending rule(s)." -Level INFO
            $script:PendingSaveCounts[$rulesKey] = 0
        }
        catch {
            Write-OMMigrateLog -Message "Invoke-FlushPendingRuleSave: Save() failed: $_" -Level WARN
        }
    }
}

function Invoke-BuildRulesFromMap {
    <#
    .SYNOPSIS
        Creates chunked consolidated rules for one group key against
        a target Outlook.Rules collection, saving each rule as it is
        created. Mirrors Gemini's BuildRulesFromMap VBA sub.

    .PARAMETER TargetRules
        Outlook Rules COM collection to create new rules on.

    .PARAMETER PathLookup
        The full group-key -> domain-dictionary lookup table.

    .PARAMETER MapKey
        The specific RuleStoreName|TargetFolderPath key being built.

    .PARAMETER TargetPath
        Folder path portion of MapKey (relative to the Archive PST root).

    .PARAMETER RuleAccountBinding
        Account portion of MapKey -- used in the rule name.

    .PARAMETER ActiveAccount
        Outlook Account COM object for the Account condition, or
        $null if no matching account was found (Account condition
        is skipped in that case, mirroring the macro's behavior).

    .PARAMETER ArchiveRootFolder
        Root MAPIFolder of the Archive PST.

    .PARAMETER DomainsInThisBatch
        [ordered hashtable] shared across the whole run -- every
        domain successfully assigned to a rule is recorded here so
        Step 4 of Invoke-DeployConsolidatedRules knows which CSV
        rows to timestamp.

    .PARAMETER Result
        Shared result object -- Created/Failed counters are
        incremented directly.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [object] $TargetRules,
        [Parameter(Mandatory = $true)] [object] $PathLookup,
        [Parameter(Mandatory = $true)] [string] $MapKey,
        [Parameter(Mandatory = $true)] [string] $TargetPath,
        [Parameter(Mandatory = $true)] [string] $RuleAccountBinding,
        [Parameter(Mandatory = $false)][object] $ActiveAccount,
        [Parameter(Mandatory = $true)] [object] $ArchiveRootFolder,
        [Parameter(Mandatory = $true)] [object] $DomainsInThisBatch,
        [Parameter(Mandatory = $true)] [object] $Result,

        # Added 2026-07-02, Administrator. Performance experiment: Save() was previously
        # called once per rule created, which is the single most expensive COM
        # operation in this loop (serializes the growing PR_RW_RULES_STREAM to
        # disk every time). Batching Save() to fire every N rules instead of
        # every 1 should meaningfully cut runtime -- confirmed live that a full
        # rebuild takes ~1.5 hours via this module vs ~30 seconds via the VBA
        # macro, despite both calling Save() at the same frequency previously
        # and using the same O(n^2) consolidation scan; PowerShell's InvokeMember
        # reflection overhead on writes is the other known contributor but is
        # NOT safely reducible (Memory #10 -- direct dot-notation writes were
        # already tried and found to silently fail in PowerShell's CLR/COM
        # interop layer). Default of 1 preserves exact prior behavior (save
        # every rule) -- purely additive, no existing caller needs to change.
        # NOTE: this is unrelated to the June 2026 batch-save experiments
        # (Memory context) -- those targeted PR_RULES_DATA stream-overflow
        # prevention and a StopProcessing bug later root-caused to the Account
        # condition (now fixed, see Memory re: 0x800C8101); this batching is
        # purely a runtime-performance change, tried explicitly because that
        # root cause is resolved and Administrator approved testing it as a new angle.
        [Parameter(Mandatory = $false)] [int] $SaveBatchSize = 1,

        # Added 2026-07-02, Administrator (performance fix). Pre-built lookup from
        # Invoke-DeployConsolidatedRules: CompositeKey -> List of RuleNames
        # currently targeting that folder, per the CSV (source of truth --
        # see Invoke-DeployConsolidatedRules's build-site comment). Consumed
        # by the consolidation scan below in place of a live O(n) rescan of
        # $TargetRules per group key. $null/empty is safe -- consolidation
        # simply finds nothing to absorb, same as before this optimization.
        [Parameter(Mandatory = $false)] [object] $FolderPathToRuleNames = $null,

        # EXTRACTED 2026-07-09 (Administrator direction) from OMMigrate-Outlook_WIP.psm1,
        # a 6-day-old (2026-07-03/04) work-in-progress file -- this is the
        # TargetStoreName hardcode fix work paused mid-implementation that
        # session. NOT a completed/finished feature -- extracted as a
        # jumpstart on resuming this work, not yet live-tested. Optional
        # per-group root folder resolved by Invoke-DeployConsolidatedRules
        # from this group's TargetStoreName (see groupKeyToArchiveFolder).
        # When provided (non-null), used INSTEAD of $ArchiveRootFolder for
        # this one call -- lets different groups in the same run target
        # different attached PSTs/stores instead of always the single
        # Archive PST. $null (default) preserves exact prior behavior --
        # every existing caller that doesn't pass this parameter is
        # completely unaffected.
        [Parameter(Mandatory = $false)] [object] $OverrideArchiveFolder = $null
    )

    $subDomainDict = $PathLookup[$MapKey]
    $domainList    = @($subDomainDict.Keys)
    $dCount        = $domainList.Count

    if ($dCount -eq 0 -or [string]::IsNullOrWhiteSpace([string]$domainList[0])) { return }

    # Resolve the archive destination folder once per group key, same as
    # the macro -- via the existing proven Get-FolderByPath helper
    # (creates intermediate folders if missing), not a new path-walker.
    #
    # EXTRACTED 2026-07-09 (Administrator direction): use $OverrideArchiveFolder when
    # the caller supplied one (this group has its own resolved TargetStoreName
    # folder), otherwise fall back to $ArchiveRootFolder exactly as before --
    # see $OverrideArchiveFolder parameter doc above for full context.
    $effectiveArchiveRoot = if ($null -ne $OverrideArchiveFolder) { $OverrideArchiveFolder } else { $ArchiveRootFolder }

    $targetArchiveFolderObj = Get-FolderByPath `
        -RootFolder      $effectiveArchiveRoot `
        -FolderPath      $TargetPath `
        -CreateIfMissing $true

    $pathSegments   = $TargetPath -split '\\'
    $baseFolderName = $pathSegments[-1].Trim()

    # ADDED 2026-07-09 (Administrator + Gemini, ported to PowerShell by Claude): UI Rule
    # Manager Condition Preservation Engine. Mirrors Gemini's Module3.bas
    # BuildRulesFromMap changes of the same date -- when an existing rule
    # targeting this same folder is consolidated/absorbed below, any OTHER
    # manually-added UI condition on that rule (Subject, Body, BodyOrSubject,
    # MessageHeader, RecipientAddress, AnyCategory, Importance, Sensitivity,
    # HasAttachment, Cc, OnlyToMe, ToOrCc) is harvested here and re-applied to
    # the new standardized rule further below, instead of being silently
    # discarded on rebuild. Per Administrator's explicit spec: the "from people or
    # public group" condition is intentionally NOT preserved (SenderAddress
    # already replaces it), and the "through the specified account" condition
    # is intentionally NOT preserved (destabilizes the COM API -- see the
    # Account condition block further below, already disabled for the same
    # reason). Each condition gets its own has-flag + value pair, exactly
    # matching the macro's variable layout, so behavior stays in lockstep
    # between the two implementations per Administrator's sync requirement.
    # ADDED (Administrator direction, this same fix -- the standing, repeated
    # instruction to write the standardized rule name back to the CSV
    # BEFORE consolidation, so a later name-based lookup finds a match
    # instead of comparing an old raw name against a new standardized one
    # in the same run): tracks every raw rule name absorbed during THIS
    # call to Invoke-BuildRulesFromMap (one call = one chunk = one
    # eventual $chunkName), separate from $Result.ConsolidatedRuleNames
    # which accumulates across the ENTIRE run and cannot be used to know
    # which names belong to which chunk.
    $absorbedOldNamesThisChunk    = [System.Collections.Generic.List[string]]::new()
    $hasPreservedSubject         = $false
    $preservedSubjectText        = $null
    $hasPreservedBody            = $false
    $preservedBodyText           = $null
    $hasPreservedBodyOrSubject   = $false
    $preservedBodyOrSubjectText  = $null
    $hasPreservedHeader          = $false
    $preservedHeaderText         = $null
    $hasPreservedRecipient       = $false
    $preservedRecipientText      = $null
    $hasPreservedCategory        = $false
    $preservedCategoryText       = $null
    $hasPreservedImportance      = $false
    $preservedImportanceVal      = $null
    $hasPreservedSensitivity     = $false
    $preservedSensitivityVal     = $null
    $hasPreservedAttachment      = $false
    $hasPreservedCc              = $false
    $hasPreservedOnlyToMe        = $false
    $hasPreservedToOrCc          = $false

    # Rule consolidation (REWRITTEN 2026-07-02, Administrator's design -- performance
    # fix, replaces the O(n^2) live-collection rescan documented below).
    #
    # ORIGINAL DESIGN (added 2026-07-01, still the correct REQUIREMENT, only
    # the MECHANISM changed): ANY existing rule -- manually created or
    # previously auto-generated, regardless of its current name -- that
    # targets this same folder ($TargetPath) gets folded into the new
    # standardized "Rule: [account] Folder (Part N)" rule, not left standing
    # separately. RuleName is explicitly NOT part of the match -- it's the
    # very thing being standardized, so it's expected to differ.
    #
    # OLD MECHANISM (removed): iterated the ENTIRE live $TargetRules
    # collection (~356 rules) for EVERY group key, calling Actions.MoveToFolder
    # + a full parent-chain folder-path walk (Get-FolderFullPath) per
    # candidate, just to compare against $TargetPath. O(n^2) overall --
    # confirmed live as the dominant cost behind a ~1.5-hour full-rebuild
    # runtime (~126,000 iterations on a 356-rule run), vs the VBA macro's
    # ~30 seconds despite identical Save()-per-rule frequency.
    #
    # NEW MECHANISM: $FolderPathToRuleNames (built ONCE in
    # Invoke-DeployConsolidatedRules, directly from rules_inventory.csv's
    # RuleName+TargetFolderPath columns -- the CSV IS the source of truth
    # for existing rules' folder targets, per the documented pipeline
    # contract that Script 00 is re-run before Script 03 whenever rules
    # change) already tells us exactly which rule NAME(s), if any, target
    # this folder -- zero live COM calls needed to answer that question.
    # Only the ACTUAL deletion still needs a live call: a short, targeted
    # name-match scan (Name is a cheap property read, unlike
    # Actions.MoveToFolder + a full folder-path walk) limited to the small
    # set of already-known candidate names, not all 356 live rules.
    try {
        $candidateNames = $null
        if ($FolderPathToRuleNames -and $FolderPathToRuleNames.Contains($MapKey)) {
            $candidateNames = $FolderPathToRuleNames[$MapKey]
        }

        if ($candidateNames -and $candidateNames.Count -gt 0) {
            $candidateNameSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$candidateNames)
            $existingRuleCountForFolder = $TargetRules.Count
            for ($folderIdx = $existingRuleCountForFolder; $folderIdx -ge 1; $folderIdx--) {
                if ($candidateNameSet.Count -eq 0) { break }  # all known candidates already found/deleted
                $candidateRule = $null
                try { $candidateRule = $TargetRules.Item($folderIdx) } catch { }
                if ($null -eq $candidateRule) { continue }

                $candidateRuleName = ''
                try { $candidateRuleName = [string]$candidateRule.Name } catch { }
                if (-not $candidateNameSet.Contains($candidateRuleName)) { continue }
                [void]$candidateNameSet.Remove($candidateRuleName)

                # Confirmed match by name (from CSV source of truth) -- pull its
                # SenderAddress domains (if any) before deleting it.
                #
                # FIXED (2026-07-07, Administrator + Gemini): this loop previously took
                # each existing $existingAddr array element and dropped it
                # straight into $subDomainDict as a single dictionary key with
                # no further processing -- the exact same blind spot Gemini
                # identified and fixed in the VBA macro's BuildRulesFromMap
                # (see Module3.bas). If a PRIOR run had ever deployed a rule
                # with a mangled multi-word single array element (e.g.
                # "chatgpt openai" as ONE .Address entry instead of two
                # separate entries "chatgpt" and "openai"), this consolidation
                # step would silently inherit and re-deploy that same mistake
                # forever, since nothing here ever re-validated or re-split it.
                # Separately, a full email address read back from an existing
                # rule must NEVER be word-split around its "@" -- e.g.
                # "joe_black@123.com" must stay one token, not become
                # "joe_black" and "123.com" as two unrelated words.
                #
                # Fix: reuse ConvertTo-NormalizedSenderDomains (already proven
                # correct for the CSV-input case, see Issue 3) for each
                # existing array element too. That function already (a)
                # splits on embedded spaces if present, (b) validates each
                # resulting token as either a full email or a bare word/
                # domain fragment using the same project-wide character set,
                # and (c) checks the EMAIL pattern first, so a genuine full
                # email is protected as one unit and never split around its
                # "@" -- both fixes Gemini applied in the macro, achieved
                # here via reuse of existing, already-approved logic rather
                # than new code.
                try {
                    # FIXED (2026-07-07, Administrator debug session): the line below this
                    # comment previously read Conditions.SenderAddress via plain PS
                    # dot-notation ($candidateRule.Conditions.SenderAddress), NOT via
                    # InvokeMember reflection -- a direct violation of this module's
                    # own standing COM rule (see header .NOTES / project session
                    # rules: "All Outlook COM rule object property access... MUST use
                    # InvokeMember reflection -- both reads and writes"). Every other
                    # COM property read in this same function (OnLocalMachine, Stop,
                    # SenderAddress on the NEW rule further below, etc.) already goes
                    # through GetType().InvokeMember(..., GetProperty, ...) -- this
                    # was the one read site that never got converted. Per Administrator: this
                    # is the leading suspect for the SendersDomain regression, since
                    # the harvesting/validation logic itself (ConvertTo-
                    # NormalizedSenderDomains call below) is structurally identical
                    # to Gemini's proven-correct VBA fix and was never found to be
                    # wrong on code review alone.
                    #
                    # Switched to InvokeMember GetProperty for both the Conditions.
                    # SenderAddress read and the .Address array read on it, matching
                    # the proven pattern used elsewhere in this function (e.g. the
                    # OnLocalMachine / Stop reads above).
                    $candidateSenderCond = $candidateRule.Conditions.GetType().InvokeMember(
                        'SenderAddress',
                        [System.Reflection.BindingFlags]::GetProperty,
                        $null, $candidateRule.Conditions, $null
                    )

                    $candidateSenderEnabled = $false
                    $candidateSenderAddressRaw = $null
                    if ($null -ne $candidateSenderCond) {
                        $candidateSenderEnabled = $candidateSenderCond.GetType().InvokeMember(
                            'Enabled',
                            [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateSenderCond, $null
                        )
                        $candidateSenderAddressRaw = $candidateSenderCond.GetType().InvokeMember(
                            'Address',
                            [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateSenderCond, $null
                        )
                    }

                    # DEBUG (2026-07-07, Administrator debug session): dump exactly what the
                    # InvokeMember read returns for this rule, so next live run can
                    # confirm whether this was actually the source of the mangled
                    # multi-word-no-"or" regression, or whether it points elsewhere.
                    try {
                        $debugAddrDump = if ($null -eq $candidateSenderAddressRaw) { '<null>' } else { ($candidateSenderAddressRaw -join ' | ') }
                        $debugAddrType = if ($null -eq $candidateSenderAddressRaw) { '<null>' } else { $candidateSenderAddressRaw.GetType().FullName }
                        Write-OMMigrateLog -Message "DEBUG SendersDomain harvest (InvokeMember): Rule='$candidateRuleName' Enabled=$candidateSenderEnabled AddressType=$debugAddrType AddressRaw=[$debugAddrDump]" -Level DEBUG
                    } catch { }

                    if ($candidateSenderCond -and $candidateSenderEnabled -and $candidateSenderAddressRaw) {
                        foreach ($existingAddr in @($candidateSenderAddressRaw)) {
                            $existingAddrTrimmed = ([string]$existingAddr).Trim()
                            if (-not [string]::IsNullOrWhiteSpace($existingAddrTrimmed)) {
                                $existingAddrTokens = @(ConvertTo-NormalizedSenderDomains -RawValue $existingAddrTrimmed)
                                foreach ($existingToken in $existingAddrTokens) {
                                    $subDomainDict[$existingToken] = $true
                                }
                            }
                        }
                    }
                } catch { }

                # ADDED 2026-07-09 (Administrator + Gemini, ported to PowerShell by Claude):
                # UI Rule Manager Condition Preservation Engine -- harvest phase.
                # Mirrors Module3.bas BuildRulesFromMap's condition-preservation
                # block of the same date, immediately following the SenderAddress
                # harvest above. Each read goes through InvokeMember GetProperty
                # per the standing project COM rule. Intentionally does NOT read
                # Conditions.From or Conditions.Account -- per Administrator's spec, From
                # is superseded by SenderAddress and Account is never preserved
                # (destabilizes the COM API, see the disabled Account condition
                # block further below in this function).
                try {
                    $candidateSubjectCond = $candidateRule.Conditions.GetType().InvokeMember(
                        'Subject', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $candidateRule.Conditions, $null
                    )
                    if ($null -ne $candidateSubjectCond) {
                        $subjEnabled = $candidateSubjectCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateSubjectCond, $null
                        )
                        if ($subjEnabled) {
                            $hasPreservedSubject = $true
                            $preservedSubjectText = $candidateSubjectCond.GetType().InvokeMember(
                                'Text', [System.Reflection.BindingFlags]::GetProperty,
                                $null, $candidateSubjectCond, $null
                            )
                        }
                    }
                } catch { }

                try {
                    $candidateBodyCond = $candidateRule.Conditions.GetType().InvokeMember(
                        'Body', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $candidateRule.Conditions, $null
                    )
                    if ($null -ne $candidateBodyCond) {
                        $bodyEnabled = $candidateBodyCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateBodyCond, $null
                        )
                        if ($bodyEnabled) {
                            $hasPreservedBody = $true
                            $preservedBodyText = $candidateBodyCond.GetType().InvokeMember(
                                'Text', [System.Reflection.BindingFlags]::GetProperty,
                                $null, $candidateBodyCond, $null
                            )
                        }
                    }
                } catch { }

                try {
                    $candidateBodyOrSubjectCond = $candidateRule.Conditions.GetType().InvokeMember(
                        'BodyOrSubject', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $candidateRule.Conditions, $null
                    )
                    if ($null -ne $candidateBodyOrSubjectCond) {
                        $bodyOrSubjEnabled = $candidateBodyOrSubjectCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateBodyOrSubjectCond, $null
                        )
                        if ($bodyOrSubjEnabled) {
                            $hasPreservedBodyOrSubject = $true
                            $preservedBodyOrSubjectText = $candidateBodyOrSubjectCond.GetType().InvokeMember(
                                'Text', [System.Reflection.BindingFlags]::GetProperty,
                                $null, $candidateBodyOrSubjectCond, $null
                            )
                        }
                    }
                } catch { }

                try {
                    $candidateHeaderCond = $candidateRule.Conditions.GetType().InvokeMember(
                        'MessageHeader', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $candidateRule.Conditions, $null
                    )
                    if ($null -ne $candidateHeaderCond) {
                        $headerEnabled = $candidateHeaderCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateHeaderCond, $null
                        )
                        if ($headerEnabled) {
                            $hasPreservedHeader = $true
                            $preservedHeaderText = $candidateHeaderCond.GetType().InvokeMember(
                                'Text', [System.Reflection.BindingFlags]::GetProperty,
                                $null, $candidateHeaderCond, $null
                            )
                        }
                    }
                } catch { }

                try {
                    $candidateRecipientCond = $candidateRule.Conditions.GetType().InvokeMember(
                        'RecipientAddress', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $candidateRule.Conditions, $null
                    )
                    if ($null -ne $candidateRecipientCond) {
                        $recipEnabled = $candidateRecipientCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateRecipientCond, $null
                        )
                        if ($recipEnabled) {
                            $hasPreservedRecipient = $true
                            $preservedRecipientText = $candidateRecipientCond.GetType().InvokeMember(
                                'Address', [System.Reflection.BindingFlags]::GetProperty,
                                $null, $candidateRecipientCond, $null
                            )
                        }
                    }
                } catch { }

                # FIXED (2026-07-09, live test on admin@example-provider.com: DISP_E_UNKNOWNNAME):
                # AnyCategory (ConditionType 29, "is assigned any category") does NOT
                # have a .Categories property -- it is a boolean-only condition (has
                # any category vs. none), same shape as HasAttachment/Cc/OnlyToMe/
                # ToOrCc below. The condition that actually carries a specific-category
                # .Categories string array is Conditions.Category (singular),
                # returning a CategoryRuleCondition object -- confirmed against
                # Microsoft's own CreateTextAndCategoryRule sample, which sets
                # .Categories on Actions.AssignToCategory (same object family) the
                # same way. The macro (Module3.bas) also calls .Categories on
                # AnyCategory -- same latent bug, just never live-tested there (Administrator
                # confirmed only Subject was tested on the macro), and VBA's
                # late-bound dispatch may be masking the error rather than avoiding
                # it. Switched to 'Category' here; the macro should be checked/fixed
                # to match if/when Category preservation is actually tested there.
                try {
                    $candidateCategoryCond = $candidateRule.Conditions.GetType().InvokeMember(
                        'Category', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $candidateRule.Conditions, $null
                    )
                    if ($null -ne $candidateCategoryCond) {
                        $catEnabled = $candidateCategoryCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateCategoryCond, $null
                        )
                        if ($catEnabled) {
                            $hasPreservedCategory = $true
                            $preservedCategoryText = $candidateCategoryCond.GetType().InvokeMember(
                                'Categories', [System.Reflection.BindingFlags]::GetProperty,
                                $null, $candidateCategoryCond, $null
                            )
                        }
                    }
                } catch { }

                try {
                    $candidateImportanceCond = $candidateRule.Conditions.GetType().InvokeMember(
                        'Importance', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $candidateRule.Conditions, $null
                    )
                    if ($null -ne $candidateImportanceCond) {
                        $importanceEnabled = $candidateImportanceCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateImportanceCond, $null
                        )
                        if ($importanceEnabled) {
                            $hasPreservedImportance = $true
                            $preservedImportanceVal = $candidateImportanceCond.GetType().InvokeMember(
                                'Importance', [System.Reflection.BindingFlags]::GetProperty,
                                $null, $candidateImportanceCond, $null
                            )
                        }
                    }
                } catch { }

                try {
                    $candidateSensitivityCond = $candidateRule.Conditions.GetType().InvokeMember(
                        'Sensitivity', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $candidateRule.Conditions, $null
                    )
                    if ($null -ne $candidateSensitivityCond) {
                        $sensitivityEnabled = $candidateSensitivityCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateSensitivityCond, $null
                        )
                        if ($sensitivityEnabled) {
                            $hasPreservedSensitivity = $true
                            $preservedSensitivityVal = $candidateSensitivityCond.GetType().InvokeMember(
                                'Sensitivity', [System.Reflection.BindingFlags]::GetProperty,
                                $null, $candidateSensitivityCond, $null
                            )
                        }
                    }
                } catch { }

                try {
                    $candidateAttachmentCond = $candidateRule.Conditions.GetType().InvokeMember(
                        'HasAttachment', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $candidateRule.Conditions, $null
                    )
                    if ($null -ne $candidateAttachmentCond) {
                        $attachEnabled = $candidateAttachmentCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateAttachmentCond, $null
                        )
                        if ($attachEnabled) { $hasPreservedAttachment = $true }
                    }
                } catch { }

                try {
                    $candidateCcCond = $candidateRule.Conditions.GetType().InvokeMember(
                        'Cc', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $candidateRule.Conditions, $null
                    )
                    if ($null -ne $candidateCcCond) {
                        $ccEnabled = $candidateCcCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateCcCond, $null
                        )
                        if ($ccEnabled) { $hasPreservedCc = $true }
                    }
                } catch { }

                try {
                    $candidateOnlyToMeCond = $candidateRule.Conditions.GetType().InvokeMember(
                        'OnlyToMe', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $candidateRule.Conditions, $null
                    )
                    if ($null -ne $candidateOnlyToMeCond) {
                        $onlyToMeEnabled = $candidateOnlyToMeCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateOnlyToMeCond, $null
                        )
                        if ($onlyToMeEnabled) { $hasPreservedOnlyToMe = $true }
                    }
                } catch { }

                try {
                    $candidateToOrCcCond = $candidateRule.Conditions.GetType().InvokeMember(
                        'ToOrCc', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $candidateRule.Conditions, $null
                    )
                    if ($null -ne $candidateToOrCcCond) {
                        $toOrCcEnabled = $candidateToOrCcCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::GetProperty,
                            $null, $candidateToOrCcCond, $null
                        )
                        if ($toOrCcEnabled) { $hasPreservedToOrCc = $true }
                    }
                } catch { }

                Write-OMMigrateLog -Message "Rule consolidation: absorbing existing rule '$candidateRuleName' (same target folder '$TargetPath', matched via CSV lookup) into standardized rule -- deleting old rule." -Level INFO
                try {
                    # Record name before delete (2026-07-01, Administrator) -- see
                    # ConsolidatedRuleNames note on $result's initialization above.
                    try { $Result.ConsolidatedRuleNames.Add($candidateRuleName) } catch { }
                    # ADDED (Administrator direction, this same fix): also record it locally,
                    # scoped to this chunk -- see $absorbedOldNamesThisChunk declaration.
                    try { $absorbedOldNamesThisChunk.Add($candidateRuleName) } catch { }

                    $TargetRules.Remove($folderIdx)

                    # Save() restored, corrected (2026-07-01, Administrator -- confirmed via
                    # Gemini review of the actual Module3.bas source): the macro DOES
                    # save -- via VBA.Interaction.CallByName targetRules, "Save", VbMethod
                    # at the end of its chunking loop. Per Gemini: must use InvokeMember
                    # with an explicit @($false) ShowProgress argument.
                    [void]$TargetRules.GetType().InvokeMember(
                        'Save',
                        [System.Reflection.BindingFlags]::InvokeMethod,
                        $null, $TargetRules, @($false)
                    )
                }
                catch { Write-OMMigrateLog -Message "Rule consolidation: failed to delete existing rule '$candidateRuleName': $_" -Level WARN }
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Rule consolidation: failed to process known candidates for '$TargetPath': $_" -Level WARN
    }

    # Re-read domain list -- may have grown from the consolidation scan above.
    $domainList = @($subDomainDict.Keys)
    $dCount     = $domainList.Count

    $dCounter    = 0
    $partCounter = 1

    while ($dCounter -lt $dCount) {
        # 1000-action watchdog -- mirrors the macro's hard limit (Do While
        # loop check at the top of BuildRulesFromMap) to avoid exceeding
        # Outlook's MAPI rule storage capacity in one run.
        if ($Result.RuleActionCounter -ge 1000) { break }

        $itemsInChunk = if (($dCount - $dCounter) -lt 5) { $dCount - $dCounter } else { 5 }
        $chunkArray   = New-Object string[] $itemsInChunk

        for ($cIdx = 0; $cIdx -lt $itemsInChunk; $cIdx++) {
            $chunkArray[$cIdx] = [string]$domainList[$dCounter]
            $DomainsInThisBatch[[string]$domainList[$dCounter]] = $true
            $dCounter++
        }

        # Rule naming convention -- exact match to Gemini's macro:
        # "Rule: [<account>] <folder> (Part N)"
        $chunkName = "Rule: [$RuleAccountBinding] $baseFolderName (Part $partCounter)"

        # DISABLED 2026-07-02 (Administrator, performance): confirmed redundant now that
        # the consolidation scan above (rewritten same session) already deletes
        # any existing rule matching this folder path -- via the CSV-based
        # $FolderPathToRuleNames lookup, which runs BEFORE this while loop starts
        # and would already have deleted a same-name/same-folder rule found in
        # rules_inventory.csv (the source of truth, per the documented pipeline
        # contract that Script 00 is re-run before Script 03 whenever rules
        # change). This guard added an extra O(n) live-collection scan per CHUNK
        # (not just per group key), contributing to the same class of COM-call-
        # volume overhead the consolidation-scan rewrite targeted. Short-
        # circuited via "if ($false)" rather than deleted, per Administrator's
        # instruction, in case the CSV/consolidation lookup ever misses a case
        # this defensive guard would have caught -- original logic preserved
        # below, unreachable.
        #
        # Delete-before-create dedup guard (added 2026-07-01, Administrator): confirmed
        # in testing that manually blanking LastDeployedRun on an already-
        # deployed row and rerunning (Script 03 or the macro) recreates the
        # rule as a duplicate alongside the existing one of the same name,
        # rather than replacing it. RuleName exact match, scoped to this
        # account's own rules collection ($TargetRules) only -- never touches
        # rules belonging to other accounts. Delete only, no other property
        # is read or modified on the existing rule; if delete fails for any
        # reason, log and continue to Create() as before (defensive -- a
        # failed delete should not block the rebuild).
        if ($false) {
        try {
            $existingRuleCount = $TargetRules.Count
            for ($existIdx = $existingRuleCount; $existIdx -ge 1; $existIdx--) {
                $existingRule = $null
                try { $existingRule = $TargetRules.Item($existIdx) } catch { }
                if ($null -ne $existingRule -and $existingRule.Name -eq $chunkName) {
                    Write-OMMigrateLog -Message "Dedup guard: deleting existing rule '$chunkName' before recreating (LastDeployedRun was blanked)." -Level INFO
                    try { $TargetRules.Remove($existIdx) }
                    catch { Write-OMMigrateLog -Message "Dedup guard: failed to delete existing rule '$chunkName': $_" -Level WARN }
                    break
                }
            }
        }
        catch {
            Write-OMMigrateLog -Message "Dedup guard: failed to scan for existing rule '$chunkName' before create: $_" -Level WARN
        }
        } # end if ($false) -- dedup guard disabled, see comment above

        try {
            # =========================================================================
            # PRODUCTION RULE CREATION ENGINE (REPLACEMENT BLOCK WITH MARSHALING FIX)
            # Gemini round 5 (2026-06-26). Implemented verbatim per Administrator's explicit
            # instruction: no further deviation on our side. Order: Stop ->
            # OnLocalMachine -> Account -> SenderAddress (with restored explicit
            # [string[]] cast to fix DISP_E_TYPEMISMATCH from round 3/4's
            # [object](,...) wrapper) -> MoveToFolder -> Save.
            # =========================================================================

            # 1. Create the brand-new, clean-slate Rule object
            Write-OMMigrateLog -Message "Creating rule: '$chunkName'" -Level INFO

            $newRule = $TargetRules.GetType().InvokeMember(
                'Create',
                [System.Reflection.BindingFlags]::InvokeMethod,
                $null, $TargetRules, @($chunkName, 0) # 0 = olRuleReceive
            )
            Register-COMObject -ComObject $newRule

            [void]$newRule.GetType().InvokeMember(
                'Enabled',
                [System.Reflection.BindingFlags]::SetProperty,
                $null, $newRule, @($true)
            )

            # 2. ACTION 1: Enable 'Stop Processing More Rules' Natively (MUST BE FIRST)
            try {
                # Resolve the named property 'Stop' directly via Reflection
                $stopAction = $newRule.Actions.GetType().InvokeMember(
                    'Stop',
                    [System.Reflection.BindingFlags]::GetProperty,
                    $null, $newRule.Actions, $null
                )

                if ($null -ne $stopAction) {
                    [void]$stopAction.GetType().InvokeMember(
                        'Enabled',
                        [System.Reflection.BindingFlags]::SetProperty,
                        $null, $stopAction, @($true)
                    )
                    Write-OMMigrateLog -Message "Successfully enabled named Stop action natively via reflection." -Level DEBUG
                } else {
                    Write-OMMigrateLog -Message "Critical: Actions.Stop property returned null." -Level WARN
                }
            }
            catch {
                Write-OMMigrateLog -Message "Failed to apply native Stop action: $_" -Level WARN
            }

            # 3. CONDITION 1: Force Local Machine Isolation (REQUIRED FOR IMAP)
            try {
                $localMachineCondition = $newRule.Conditions.GetType().InvokeMember(
                    'OnLocalMachine',
                    [System.Reflection.BindingFlags]::GetProperty,
                    $null, $newRule.Conditions, $null
                )
                if ($null -ne $localMachineCondition) {
                    [void]$localMachineCondition.GetType().InvokeMember(
                        'Enabled',
                        [System.Reflection.BindingFlags]::SetProperty,
                        $null, $localMachineCondition, @($true)
                    )
                    Write-OMMigrateLog -Message "Successfully applied OnLocalMachine (on this computer only) condition." -Level DEBUG
                }
            }
            catch {
                Write-OMMigrateLog -Message "Failed configuring OnLocalMachine condition: $_" -Level WARN
            }

            # 4. CONDITION 2: Account Linkage (If active account wrapper exists)
            # DISABLED 2026-07-01 (Administrator, "TestProfile" controlled A/B test): confirmed
            # that a real, manually-selected, correctly-resolving Account
            # condition -- not just an unresolved "the specified account"
            # placeholder -- still triggers 0x800C8101 on that rule's Item()
            # read. Test sequence on a single rule ("aaa"), same rule/position,
            # only the Account condition toggled:
            #   1. No Account condition               -> 5/5 Item() clean
            #   2. Account condition added (real,      -> 4/5, Item(1) failed
            #      manually-selected account)             with 0x800C8101
            #   3. Account condition removed again     -> 5/5 Item() clean
            # New requirement: no Account condition on any rule, primary or
            # secondary, going forward. Block intentionally left in place
            # (not deleted) rather than removed, in case this needs to be
            # re-enabled -- short-circuited via "-and $false" so it never
            # executes. If this fix proves out at full production scale,
            # Administrator's direction is to delete this block permanently.
            if ($null -ne $ActiveAccount -and $false) {
                try {
                    $accountCondition = $newRule.Conditions.GetType().InvokeMember(
                        'Account',
                        [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $accountCondition) {
                        [void]$accountCondition.GetType().InvokeMember(
                            'Enabled',
                            [System.Reflection.BindingFlags]::SetProperty,
                            $null, $accountCondition, @($true)
                        )
                        [void]$accountCondition.GetType().InvokeMember(
                            'Account',
                            [System.Reflection.BindingFlags]::SetProperty,
                            $null, $accountCondition, @($ActiveAccount)
                        )
                    }
                }
                catch {
                    Write-OMMigrateLog -Message "Failed configuring Account condition: $_" -Level WARN
                }
            }

            # 5. CONDITION 3: Sender Address Domain Array (WITH FIX FOR TYPEMISMATCH)
            try {
                $senderAddressCondition = $newRule.Conditions.GetType().InvokeMember(
                    'SenderAddress',
                    [System.Reflection.BindingFlags]::GetProperty,
                    $null, $newRule.Conditions, $null
                )
                if ($null -ne $senderAddressCondition) {
                    [void]$senderAddressCondition.GetType().InvokeMember(
                        'Enabled',
                        [System.Reflection.BindingFlags]::SetProperty,
                        $null, $senderAddressCondition, @($true)
                    )
                    # RESTORED: Explicitly typed array wrapper prevents DISP_E_TYPEMISMATCH
                    [void]$senderAddressCondition.GetType().InvokeMember(
                        'Address',
                        [System.Reflection.BindingFlags]::SetProperty,
                        $null, $senderAddressCondition, @(,[string[]]$chunkArray)
                    )
                }
            }
            catch {
                Write-OMMigrateLog -Message "Failed configuring SenderAddress domains: $_" -Level WARN
            }

            # 5b. ADDED 2026-07-09 (Administrator + Gemini, ported to PowerShell by Claude):
            # UI Rule Manager Condition Preservation Engine -- apply phase. Mirrors
            # Module3.bas BuildRulesFromMap's re-application block of the same
            # date. Each condition harvested above (if any) is written back onto
            # the new standardized rule here, via InvokeMember SetProperty per the
            # standing project COM rule. Order matches the macro: Subject, Body,
            # BodyOrSubject, MessageHeader, RecipientAddress, AnyCategory,
            # Importance, Sensitivity, HasAttachment, Cc, OnlyToMe, ToOrCc.
            #
            # FIXED (2026-07-09, live test on admin@example.com and
            # admin@example-provider.com, Subject condition): Subject.Text,
            # Body.Text, BodyOrSubject.Text, MessageHeader.Text, and
            # RecipientAddress.Address are all array-typed COM properties
            # (Variant string arrays), same shape as SenderAddress.Address --
            # NOT plain scalar strings. Writing them via plain @($value)
            # (correct for a true scalar like Importance/Sensitivity) throws
            # DISP_E_TYPEMISMATCH and fails the rule's Save(), which in turn
            # deletes the OLD rule (already removed during consolidation)
            # without a working replacement -- net result: rule content lost.
            # Fixed by wrapping each of these five writes in the same explicit
            # @(,[string[]]$value) cast already proven correct for
            # SenderAddress.Address a few lines above. Confirmed live after
            # fix: both test rules rebuilt successfully with Subject condition
            # intact.
            if ($hasPreservedSubject) {
                try {
                    $newSubjectCond = $newRule.Conditions.GetType().InvokeMember(
                        'Subject', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $newSubjectCond) {
                        [void]$newSubjectCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newSubjectCond, @($true)
                        )
                        [void]$newSubjectCond.GetType().InvokeMember(
                            'Text', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newSubjectCond, @(,[string[]]$preservedSubjectText)
                        )
                    }
                }
                catch { Write-OMMigrateLog -Message "Failed re-applying preserved Subject condition: $_" -Level WARN }
            }

            if ($hasPreservedBody) {
                try {
                    $newBodyCond = $newRule.Conditions.GetType().InvokeMember(
                        'Body', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $newBodyCond) {
                        [void]$newBodyCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newBodyCond, @($true)
                        )
                        [void]$newBodyCond.GetType().InvokeMember(
                            'Text', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newBodyCond, @(,[string[]]$preservedBodyText)
                        )
                    }
                }
                catch { Write-OMMigrateLog -Message "Failed re-applying preserved Body condition: $_" -Level WARN }
            }

            if ($hasPreservedBodyOrSubject) {
                try {
                    $newBodyOrSubjectCond = $newRule.Conditions.GetType().InvokeMember(
                        'BodyOrSubject', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $newBodyOrSubjectCond) {
                        [void]$newBodyOrSubjectCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newBodyOrSubjectCond, @($true)
                        )
                        [void]$newBodyOrSubjectCond.GetType().InvokeMember(
                            'Text', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newBodyOrSubjectCond, @(,[string[]]$preservedBodyOrSubjectText)
                        )
                    }
                }
                catch { Write-OMMigrateLog -Message "Failed re-applying preserved BodyOrSubject condition: $_" -Level WARN }
            }

            if ($hasPreservedHeader) {
                try {
                    $newHeaderCond = $newRule.Conditions.GetType().InvokeMember(
                        'MessageHeader', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $newHeaderCond) {
                        [void]$newHeaderCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newHeaderCond, @($true)
                        )
                        [void]$newHeaderCond.GetType().InvokeMember(
                            'Text', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newHeaderCond, @(,[string[]]$preservedHeaderText)
                        )
                    }
                }
                catch { Write-OMMigrateLog -Message "Failed re-applying preserved MessageHeader condition: $_" -Level WARN }
            }

            if ($hasPreservedRecipient) {
                try {
                    $newRecipientCond = $newRule.Conditions.GetType().InvokeMember(
                        'RecipientAddress', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $newRecipientCond) {
                        [void]$newRecipientCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newRecipientCond, @($true)
                        )
                        [void]$newRecipientCond.GetType().InvokeMember(
                            'Address', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newRecipientCond, @(,[string[]]$preservedRecipientText)
                        )
                    }
                }
                catch { Write-OMMigrateLog -Message "Failed re-applying preserved RecipientAddress condition: $_" -Level WARN }
            }

            # FIXED (2026-07-09, live test on admin@example-provider.com: DISP_E_UNKNOWNNAME):
            # AnyCategory has no .Categories property -- see matching harvest-phase
            # comment above for full root-cause explanation. Switched to the
            # correct 'Category' property (CategoryRuleCondition object). The
            # @(,[string[]]$value) array-cast fix (same class as Subject/Body/
            # BodyOrSubject/MessageHeader/RecipientAddress above) is still
            # required and kept below -- Categories is genuinely array-typed on
            # the correct object too, this was two separate bugs stacked on the
            # same condition (wrong property name AND wrong write-cast).
            if ($hasPreservedCategory) {
                try {
                    $newCategoryCond = $newRule.Conditions.GetType().InvokeMember(
                        'Category', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $newCategoryCond) {
                        [void]$newCategoryCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newCategoryCond, @($true)
                        )
                        [void]$newCategoryCond.GetType().InvokeMember(
                            'Categories', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newCategoryCond, @(,[string[]]$preservedCategoryText)
                        )
                    }
                }
                catch { Write-OMMigrateLog -Message "Failed re-applying preserved AnyCategory condition: $_" -Level WARN }
            }

            if ($hasPreservedImportance) {
                try {
                    $newImportanceCond = $newRule.Conditions.GetType().InvokeMember(
                        'Importance', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $newImportanceCond) {
                        [void]$newImportanceCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newImportanceCond, @($true)
                        )
                        # Defensive [int] cast, same spirit as the SenderAddress
                        # [string[]] cast fix above -- COM enum values marshaled
                        # through InvokeMember can arrive boxed as a non-int
                        # numeric type; SetProperty on Importance expects an
                        # OlImportance enum-compatible int.
                        [void]$newImportanceCond.GetType().InvokeMember(
                            'Importance', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newImportanceCond, @([int]$preservedImportanceVal)
                        )
                    }
                }
                catch { Write-OMMigrateLog -Message "Failed re-applying preserved Importance condition: $_" -Level WARN }
            }

            if ($hasPreservedSensitivity) {
                try {
                    $newSensitivityCond = $newRule.Conditions.GetType().InvokeMember(
                        'Sensitivity', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $newSensitivityCond) {
                        [void]$newSensitivityCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newSensitivityCond, @($true)
                        )
                        # Defensive [int] cast -- same reasoning as Importance above.
                        [void]$newSensitivityCond.GetType().InvokeMember(
                            'Sensitivity', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newSensitivityCond, @([int]$preservedSensitivityVal)
                        )
                    }
                }
                catch { Write-OMMigrateLog -Message "Failed re-applying preserved Sensitivity condition: $_" -Level WARN }
            }

            if ($hasPreservedAttachment) {
                try {
                    $newAttachmentCond = $newRule.Conditions.GetType().InvokeMember(
                        'HasAttachment', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $newAttachmentCond) {
                        [void]$newAttachmentCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newAttachmentCond, @($true)
                        )
                    }
                }
                catch { Write-OMMigrateLog -Message "Failed re-applying preserved HasAttachment condition: $_" -Level WARN }
            }

            if ($hasPreservedCc) {
                try {
                    $newCcCond = $newRule.Conditions.GetType().InvokeMember(
                        'Cc', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $newCcCond) {
                        [void]$newCcCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newCcCond, @($true)
                        )
                    }
                }
                catch { Write-OMMigrateLog -Message "Failed re-applying preserved Cc condition: $_" -Level WARN }
            }

            if ($hasPreservedOnlyToMe) {
                try {
                    $newOnlyToMeCond = $newRule.Conditions.GetType().InvokeMember(
                        'OnlyToMe', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $newOnlyToMeCond) {
                        [void]$newOnlyToMeCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newOnlyToMeCond, @($true)
                        )
                    }
                }
                catch { Write-OMMigrateLog -Message "Failed re-applying preserved OnlyToMe condition: $_" -Level WARN }
            }

            if ($hasPreservedToOrCc) {
                try {
                    $newToOrCcCond = $newRule.Conditions.GetType().InvokeMember(
                        'ToOrCc', [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $newToOrCcCond) {
                        [void]$newToOrCcCond.GetType().InvokeMember(
                            'Enabled', [System.Reflection.BindingFlags]::SetProperty,
                            $null, $newToOrCcCond, @($true)
                        )
                    }
                }
                catch { Write-OMMigrateLog -Message "Failed re-applying preserved ToOrCc condition: $_" -Level WARN }
            }

            # 6. ACTION 2: Cross-Store Move To Archive Folder (MUST BE LAST)
            if ($null -ne $targetArchiveFolderObj) {
                try {
                    $moveAction = $newRule.Actions.GetType().InvokeMember(
                        'MoveToFolder',
                        [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Actions, $null
                    )
                    if ($null -ne $moveAction) {
                        [void]$moveAction.GetType().InvokeMember(
                            'Folder',
                            [System.Reflection.BindingFlags]::SetProperty,
                            $null, $moveAction, @($targetArchiveFolderObj)
                        )
                        [void]$moveAction.GetType().InvokeMember(
                            'Enabled',
                            [System.Reflection.BindingFlags]::SetProperty,
                            $null, $moveAction, @($true)
                        )
                    }
                }
                catch {
                    Write-OMMigrateLog -Message "Failed configuring MoveToFolder action: $_" -Level WARN
                }
            }

            # 7. TRANSACTION COMMIT: Save the finished Rule object to the store
            # Batching (2026-07-02, Administrator, performance experiment): only call
            # Save() every SaveBatchSize rules, not after every single one.
            # $Result.RuleActionCounter is incremented BEFORE this check so the
            # very first rule in a run (counter going 0->1) correctly saves when
            # SaveBatchSize=1 (default, exact prior behavior). A final flush
            # after this while loop (see below) catches any remainder chunk that
            # doesn't land exactly on a batch boundary.
            $Result.Created++
            $Result.RuleActionCounter++
            # ADDED 2026-07-09 (Administrator direction): record this rule's name and its
            # already-correct resolved TargetFolderPath ($TargetPath -- this
            # function's own parameter, already used to set the MoveToFolder
            # action above) so Script 03's Phase 3 can stamp LastTargetRun for
            # this rule in THIS SAME RUN, instead of skipping it and deferring
            # to a second run. See ConsolidatedRuleTargets note on $result's
            # initialization above for full rationale. Populated here (not
            # earlier) so a rule that fails mid-creation and never reaches this
            # line correctly does NOT appear in the lookup -- Phase 3 must never
            # stamp LastTargetRun for a rule that didn't actually get created.
            try { $Result.ConsolidatedRuleTargets[$chunkName] = $TargetPath } catch { }
            # ADDED (Administrator direction, 2026-08-21 -- Conditions column accuracy
            # fix): capture the true, complete conditions summary directly from
            # $newRule via the same Get-RuleConditionsSummary function
            # Get-OutlookRules already uses -- see ConsolidatedChunkConditionsText's
            # declaration at $result's initialization above for full rationale.
            # Called here (not earlier) because every condition write for this
            # chunk -- SenderAddress at "5. CONDITION 3" above, plus any
            # preserved Subject/Body/etc. at "5b." above -- has already been
            # applied to $newRule's live COM object by this point, so the
            # summary reflects the rule's true final state, not a partial one.
            try { $Result.ConsolidatedChunkConditionsText[$chunkName] = Get-RuleConditionsSummary -Rule $newRule } catch { }
            # ADDED (Administrator direction, 2026-07-21): populate the new,
            # inherently collision-free domain-based mapping -- see
            # ConsolidatedDomainToNewName's declaration at $result's
            # initialization above for full rationale. $chunkArray is still in
            # scope from this iteration's chunk-building step earlier in this
            # while loop.
            try {
                foreach ($chunkDomain in $chunkArray) {
                    if ($chunkDomain) {
                        $Result.ConsolidatedDomainToNewName[$chunkDomain] = $chunkName
                    }
                }
            } catch { }
            # ADDED (Administrator direction, 2026-07-21, second pass -- 3-column
            # correction): capture this chunk's authoritative SendersDomain string,
            # keyed by chunk name -- see ConsolidatedChunkSendersDomain's
            # declaration at $result's initialization for full rationale.
            try {
                $Result.ConsolidatedChunkSendersDomain[$chunkName] = ($chunkArray -join ' ')
            } catch { }
            # FIXED (Administrator direction, this same fix -- corrected approach): the
            # ConsolidatedOldToNewName dictionary added earlier only covers rules
            # ABSORBED from an existing live rule with a matching folder target.
            # A rule created from a CSV row that had NO existing live rule to
            # absorb (a genuinely new row, e.g. freshly discovered by Script 00
            # under its raw UI name) never goes through the absorption path at
            # all -- its RuleName in the CSV stays at the raw name forever, even
            # though the rule Outlook actually has is the new standardized name.
            # Recording $MapKey -> $chunkName here instead covers BOTH cases
            # identically: $MapKey is built from RuleStoreName|TargetFolderPath,
            # the exact same composite key every CSV row for this group was
            # grouped by -- Invoke-DeployConsolidatedRules can now rename every
            # row whose own RuleStoreName|TargetFolderPath matches $MapKey,
            # regardless of whether that row's rule was absorbed or freshly
            # created this run.
            try { $Result.ConsolidatedMapKeyToNewName[$MapKey] = $chunkName } catch { }
            # ADDED (Administrator direction, this same fix): map every old raw name
            # absorbed into THIS chunk to the chunk's final standardized name.
            try {
                foreach ($absorbedOldName in $absorbedOldNamesThisChunk) {
                    $Result.ConsolidatedOldToNewName[$absorbedOldName] = $chunkName
                }
            } catch { }
            # Fix (2026-07-02, Administrator): $script:PendingSaveCounts is keyed by the
            # $TargetRules object itself (RuntimeHelpers.GetHashCode identity) so
            # pending-rule tracking is correctly PER COLLECTION, not global --
            # necessary because Step 2 calls this function once per account
            # (different $TargetRules each time) and Step 3 calls it again
            # against the default store's collection.
            # Fix (2026-07-02, Administrator): same StrictMode unsafe-reference issue as
            # Invoke-FlushPendingRuleSave -- use Get-Variable to check existence
            # rather than referencing $script:PendingSaveCounts directly before
            # it's ever been assigned in this scope (confirmed live: this exact
            # line threw "variable cannot be retrieved because it has not been
            # set" on the first rule of the first run this session).
            if (-not (Get-Variable -Name 'PendingSaveCounts' -Scope Script -ErrorAction SilentlyContinue)) {
                $script:PendingSaveCounts = @{}
            }
            $rulesKey = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($TargetRules)
            $script:PendingSaveCounts[$rulesKey] = if ($script:PendingSaveCounts.ContainsKey($rulesKey)) { $script:PendingSaveCounts[$rulesKey] + 1 } else { 1 }

            if ($script:PendingSaveCounts[$rulesKey] -ge $SaveBatchSize) {
                [void]$TargetRules.GetType().InvokeMember(
                    'Save',
                    [System.Reflection.BindingFlags]::InvokeMethod,
                    $null, $TargetRules, @($false)
                )
                Write-OMMigrateLog -Message "Batch commit: saved $($script:PendingSaveCounts[$rulesKey]) rule(s) (up to and including '$chunkName')." -Level INFO
                $script:PendingSaveCounts[$rulesKey] = 0
            }

            Write-OMMigrateLog -Message "Invoke-BuildRulesFromMap: Created '$chunkName' ($itemsInChunk domain(s)) -- pending batch save." -Level DEBUG
        }
        catch {
            Write-OMMigrateLog -Message "Invoke-BuildRulesFromMap: Failed to create/save '$chunkName': $_" -Level WARN
            $Result.Failed++
        }

        $partCounter++
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


function Read-RulesFromPSTStore {
    <#
    .SYNOPSIS
        Reads all rules from a mounted PST store and returns them
        as structured objects ready for recreation on another store.

    .DESCRIPTION
        Opens (or uses an already-mounted) backup PST store and calls
        GetRules() on it to retrieve the full rules collection.

        Returns a flat list of rule descriptor objects -- one per rule --
        containing all condition values and action values captured
        directly from the source rule COM objects. No transformation
        is applied. The caller (Invoke-RulesRecreation) is responsible
        for swapping the folder target pointer based on folder_map.csv.

        This is the shared extraction layer used by Script 03 for rules
        and will be used by Script 04 for Calendar/Contacts/Tasks/Notes/
        Journal extraction following the same open-PST -> read -> return
        pattern.

        Requires an active COM session from Connect-OutlookCOM.

    .PARAMETER PSTPath
        Full path to the backup PST file to read rules from.

    .PARAMETER BackupDisplayName
        Display name used when mounting the PST (for logging).

    .OUTPUTS
        [PSCustomObject[]] -- Array of rule descriptor objects.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PSTPath,

        [Parameter(Mandatory = $false)]
        [string]$BackupDisplayName = ''
    )

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    if (-not (Test-Path $PSTPath)) {
        Write-OMMigrateLog -Message "Read-RulesFromPSTStore: PST not found: $PSTPath" `
                           -Level WARN
        return $results
    }

    Write-OMMigrateLog -Message "Read-RulesFromPSTStore: Opening PST for rules extraction: $PSTPath" `
                       -Level INFO

    $pstStore = Open-PSTFile -PSTPath $PSTPath -DisplayName $BackupDisplayName
    if (-not $pstStore) {
        Write-OMMigrateLog -Message "Read-RulesFromPSTStore: Could not open PST: $PSTPath" `
                           -Level WARN
        return $results
    }

    $rulesCollection = $null
    try {
        $rulesCollection = $pstStore.GetRules()
        Register-COMObject -ComObject $rulesCollection
    }
    catch {
        Write-OMMigrateLog -Message "Read-RulesFromPSTStore: GetRules() failed on PST '$PSTPath': $_" `
                           -Level WARN
        return $results
    }

    $ruleCount = $rulesCollection.Count
    Write-OMMigrateLog -Message "Read-RulesFromPSTStore: Found $ruleCount rule(s) in PST store." `
                       -Level INFO

    for ($i = 1; $i -le $ruleCount; $i++) {
        $rule = $null
        try {
            $rule = $rulesCollection.Item($i)
            Register-COMObject -ComObject $rule

            $ruleName  = ''
            $isEnabled = $true
            $execOrder = $i
            $ruleType  = 0
            $stopProc  = $false

            try { $ruleName   = $rule.Name           } catch { }
            try { $isEnabled  = $rule.Enabled        } catch { }
            try { $execOrder  = $rule.ExecutionOrder } catch { }
            try { $ruleType   = $rule.RuleType       } catch { }

            if ([string]::IsNullOrWhiteSpace($ruleName)) { continue }

            $condFrom          = @()
            $condSubject       = @()
            $condSenderAddress = @()
            $condBody          = @()

            try {
                $conds = $rule.Conditions
                try {
                    $fromCond = $conds.From
                    if ($fromCond.Enabled) {
                        $recips   = $fromCond.Recipients
                        $fromList = [System.Collections.Generic.List[string]]::new()
                        for ($r = 1; $r -le $recips.Count; $r++) {
                            try {
                                $recip = $recips.Item($r)
                                Register-COMObject -ComObject $recip
                                $fromList.Add($recip.Address)
                            }
                            catch { }
                        }
                        $condFrom = $fromList.ToArray()
                    }
                }
                catch { }
                try {
                    $subjCond = $conds.Subject
                    if ($subjCond.Enabled) { $condSubject = @($subjCond.Text) }
                }
                catch { }
                try {
                    # FIXED (2026-07-07, Administrator debug session): converted from plain
                    # PS dot-notation ($conds.SenderAddress / .Enabled / .Address) to
                    # InvokeMember GetProperty reflection, per this module's standing
                    # COM rule (applies to ALL Outlook COM rule object property
                    # access). Export-RulesBlob feeds Script 03/04 rule extraction/
                    # recreation -- same COM property, same unreliable dot-notation
                    # access pattern as the site already confirmed buggy in
                    # Invoke-BuildRulesFromMap this session. Debug logging added so
                    # the next live run can confirm whether this read site was also
                    # contributing to the SendersDomain regression.
                    $saCond = $conds.GetType().InvokeMember(
                        'SenderAddress',
                        [System.Reflection.BindingFlags]::GetProperty,
                        $null, $conds, $null
                    )

                    $saCondEnabled = $false
                    $saCondAddressRaw = $null
                    if ($null -ne $saCond) {
                        $saCondEnabled = $saCond.GetType().InvokeMember(
                            'Enabled',
                            [System.Reflection.BindingFlags]::GetProperty,
                            $null, $saCond, $null
                        )
                        $saCondAddressRaw = $saCond.GetType().InvokeMember(
                            'Address',
                            [System.Reflection.BindingFlags]::GetProperty,
                            $null, $saCond, $null
                        )
                    }

                    # DEBUG (2026-07-07, Administrator debug session): compare InvokeMember
                    # read results against what plain dot-notation was previously
                    # returning for this same COM property.
                    try {
                        $debugBlobSaDump = if ($null -eq $saCondAddressRaw) { '<null>' } else { (@($saCondAddressRaw) -join ' | ') }
                        $debugBlobSaType = if ($null -eq $saCondAddressRaw) { '<null>' } else { $saCondAddressRaw.GetType().FullName }
                        Write-OMMigrateLog -Message "DEBUG Export-RulesBlob SenderAddress (InvokeMember): RuleName='$ruleName' Enabled=$saCondEnabled AddressType=$debugBlobSaType AddressRaw=[$debugBlobSaDump]" -Level DEBUG
                    } catch { }

                    if ($saCondEnabled) { $condSenderAddress = @($saCondAddressRaw) }
                }
                catch { }
                try {
                    $bodyCond = $conds.Body
                    if ($bodyCond.Enabled) { $condBody = @($bodyCond.Text) }
                }
                catch { }
            }
            catch { }

            $hasMoveAction      = $false
            $hasCopyAction      = $false
            $targetFolderPath   = ''
            $actionDelete       = $false
            $actionMarkRead     = $false
            $actionCategories   = @()
            $actionForward      = $false
            $actionStopProc     = $false

            try {
                $actions = $rule.Actions
                try {
                    $moveAction = $actions.MoveToFolder
                    if ($moveAction.Enabled) {
                        $hasMoveAction = $true
                        $folder = $moveAction.Folder
                        Register-COMObject -ComObject $folder
                        $targetFolderPath = Get-FolderFullPath -Folder $folder
                    }
                }
                catch { }
                try {
                    $copyAction = $actions.CopyToFolder
                    if ($copyAction.Enabled) {
                        $hasCopyAction = $true
                        if (-not $targetFolderPath) {
                            $folder = $copyAction.Folder
                            Register-COMObject -ComObject $folder
                            $targetFolderPath = Get-FolderFullPath -Folder $folder
                        }
                    }
                }
                catch { }
                try { $delAction = $actions.Delete;             if ($delAction.Enabled)    { $actionDelete   = $true } } catch { }
                try { $mrAction  = $actions.MarkAsRead;         if ($mrAction.Enabled)     { $actionMarkRead = $true } } catch { }
                try { $catAction = $actions.AssignToCategory;   if ($catAction.Enabled)    { $actionCategories = @($catAction.Categories) } } catch { }
                try { $fwdAction = $actions.Forward;            if ($fwdAction.Enabled)    { $actionForward  = $true } } catch { }
                try { $stopAction = $actions.StopProcessingRules; if ($stopAction.Enabled) { $actionStopProc = $true } } catch { }
            }
            catch { }

            # Strip store root prefix from folder path
            if ($targetFolderPath -and $targetFolderPath -like '*\*') {
                $pathParts = $targetFolderPath -split '\\'
                $targetFolderPath = ($pathParts[1..($pathParts.Length - 1)]) -join '\'
            }

            $results.Add([PSCustomObject]@{
                RuleName             = $ruleName
                IsEnabled            = $isEnabled
                ExecutionOrder       = $execOrder
                RuleType             = $ruleType
                StopProcessing       = $stopProc
                CondFrom             = $condFrom
                CondSubject          = $condSubject
                CondSenderAddress    = $condSenderAddress
                CondBody             = $condBody
                HasMoveAction        = $hasMoveAction
                HasCopyAction        = $hasCopyAction
                TargetFolderPath     = $targetFolderPath
                ActionDelete         = $actionDelete
                ActionMarkRead       = $actionMarkRead
                ActionCategories     = $actionCategories
                ActionForward        = $actionForward
                ActionStopProcessing = $actionStopProc
            })
        }
        catch {
            Write-OMMigrateLog -Message "Read-RulesFromPSTStore: Error reading rule [$i]: $_" `
                               -Level WARN
        }
    }

    Write-OMMigrateLog -Message "Read-RulesFromPSTStore: Extracted $($results.Count) rule(s) from PST." `
                       -Level INFO

    return $results
}


function Invoke-RulesRecreation {
    <#
    .SYNOPSIS
        Recreates rules from a backup PST onto a target IMAP store,
        remapping folder targets according to folder_map.csv.

    .DESCRIPTION
        Takes rule descriptor objects from Read-RulesFromPSTStore and
        recreates each rule on the target IMAP store's rules collection.
        Rules are recreated identically -- only the folder target pointer
        is remapped: Server -> IMAP store, Local -> Archive PST.
        Rules that already exist by name are skipped (safe to re-run).

        Requires an active COM session from Connect-OutlookCOM.

        NOTE (added 2026-07-09, Administrator + Claude): this function only runs when
        Script 03 is invoked with the -RecreateRules switch, and its output
        is NOT the final state of any rule. Every Script 03 run -- with or
        without -RecreateRules -- unconditionally executes the 'Updating
        Outlook Rules' phase afterward, which calls
        Invoke-DeployConsolidatedRules -> Invoke-BuildRulesFromMap. That CSV-
        driven phase consolidates/rebuilds every rule from rules_inventory.csv
        (the single source of truth) regardless of how the rule came to
        exist, so anything -RecreateRules creates here is bootstrap raw
        material, not a persisted end state. For that reason, the 2026-07-09
        UI Rule Manager Condition Preservation Engine (Subject, Body,
        BodyOrSubject, MessageHeader, RecipientAddress, AnyCategory,
        Importance, Sensitivity, HasAttachment, Cc, OnlyToMe, ToOrCc --
        see Invoke-BuildRulesFromMap and Module3.bas's BuildRulesFromMap)
        was deliberately NOT ported to this function. Any manually-added UI
        condition this function fails to carry over from the backup PST would
        be re-harvested and re-applied correctly by Invoke-BuildRulesFromMap's
        own consolidation pass moments later in the same run (or worst case,
        the next run) once the recreated rule is picked up by the CSV-driven
        phase -- porting the preservation logic here as well would be dead
        weight, not a gap. Administrator confirmed this reasoning 2026-07-09; do not
        re-open this as an inconsistency without re-confirming the phase
        ordering in OMMigrate-03-Restore.ps1 still holds.

    .PARAMETER SourceRules
        Array of rule descriptor objects from Read-RulesFromPSTStore.

    .PARAMETER TargetStoreName
        Display name of the IMAP store to create rules on.

    .PARAMETER AccountEmail
        Email address of the account being processed.

    .PARAMETER FolderMap
        Array of folder map rows from folder_map.csv.

    .PARAMETER Namespace
        Active Outlook MAPI namespace COM object.

    .PARAMETER ArchiveRootFolder
        Root MAPIFolder of the Archive PST store.

    .PARAMETER RulesInventory
        Optional. Rows from rules_inventory.csv where NeedsFolderUpdate=True for
        this account. When supplied, TargetFolderPath is taken from the CSV row
        (source of truth) rather than the source PST COM object. Matched to source
        rules by RuleName+ExecutionOrder with name-only fallback.

    .OUTPUTS
        PSCustomObject with Created, Skipped, Failed counts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[PSCustomObject]]$SourceRules,

        [Parameter(Mandatory = $true)]
        [string]$TargetStoreName,

        [Parameter(Mandatory = $true)]
        [string]$AccountEmail,

        [Parameter(Mandatory = $true)]
        [array]$FolderMap,

        [Parameter(Mandatory = $true)]
        [object]$Namespace,

        [Parameter(Mandatory = $false)]
        [object]$ArchiveRootFolder = $null,

        [Parameter(Mandatory = $false)]
        [array]$RulesInventory = @()
        # Rows from rules_inventory.csv where NeedsFolderUpdate=True for this account.
        # When supplied, TargetFolderPath is taken from the CSV (source of truth)
        # rather than from the source PST COM folder object. This ensures recovered
        # rules (e.g. IPM.RuleOrganizer items) use the verified CSV paths.
    )

    $created = 0
    $skipped = 0
    $failed  = 0

    if ($SourceRules.Count -eq 0) {
        Write-OMMigrateLog -Message "Invoke-RulesRecreation: No source rules to recreate." -Level INFO
        return [PSCustomObject]@{ Created = 0; Skipped = 0; Failed = 0 }
    }

    # Find target IMAP store
    $targetStore = $null
    try {
        $stores = $Namespace.Stores
        Register-COMObject -ComObject $stores
        for ($s = 1; $s -le $stores.Count; $s++) {
            $st = $stores.Item($s)
            Register-COMObject -ComObject $st
            $fp = ''
            try { $fp = $st.FilePath } catch { }
            if ((-not ($fp -like '*.pst')) -and
                ($st.DisplayName -eq $TargetStoreName -or
                 $st.DisplayName -like "*$AccountEmail*")) {
                $targetStore = $st
                break
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Invoke-RulesRecreation: Error finding target store: $_" -Level ERROR
        return [PSCustomObject]@{ Created = 0; Skipped = 0; Failed = $SourceRules.Count }
    }

    if (-not $targetStore) {
        Write-OMMigrateLog -Message "Invoke-RulesRecreation: Target IMAP store not found: '$TargetStoreName'" -Level ERROR
        return [PSCustomObject]@{ Created = 0; Skipped = 0; Failed = $SourceRules.Count }
    }

    $targetRules = $null
    try {
        $targetRules = $targetStore.GetRules()
        Register-COMObject -ComObject $targetRules
    }
    catch {
        Write-OMMigrateLog -Message "Invoke-RulesRecreation: Cannot get rules collection on target store: $_" -Level ERROR
        return [PSCustomObject]@{ Created = 0; Skipped = 0; Failed = $SourceRules.Count }
    }

    # Build duplicate detection set
    $existingNames = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    try {
        for ($e = 1; $e -le $targetRules.Count; $e++) {
            try {
                $er = $targetRules.Item($e)
                Register-COMObject -ComObject $er
                [void]$existingNames.Add($er.Name)
            }
            catch { }
        }
    }
    catch { }

    Write-OMMigrateLog -Message (
        "Invoke-RulesRecreation: Target '$TargetStoreName' has $($existingNames.Count) existing rule(s). " +
        "Recreating $($SourceRules.Count) source rule(s)..."
    ) -Level INFO

    # Build a set of source rule names for fast lookup in the pre-deletion pass
    $sourceRuleNames = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($sr in $SourceRules) { [void]$sourceRuleNames.Add($sr.RuleName) }

    # Pre-deletion pass: remove broken rules that we are about to recreate.
    # A rule is considered broken if its MoveToFolder action is disabled
    # (Enabled=$false) -- this is the state Strategy 2 leaves rules in when
    # the folder target is null/stale and cannot be set via InvokeMember.
    # Only rules whose name is in $sourceRuleNames are candidates -- rules
    # with valid folder targets (Enabled=$true) are left completely untouched.
    # Iterate backwards so Remove() does not shift indices.
    $deletedCount = 0
    $delCount = $targetRules.Count
    for ($d = $delCount; $d -ge 1; $d--) {
        try {
            $dr     = $targetRules.Item($d)
            $drName = ''
            try { $drName = $dr.Name } catch { continue }

            # Only touch rules we intend to recreate
            if (-not $sourceRuleNames.Contains($drName)) { continue }

            # Check if MoveToFolder is disabled (broken null-target state)
            $drEnabled = $true
            try {
                $drMove    = $dr.Actions.MoveToFolder
                $drEnabled = $drMove.Enabled
            } catch { }

            if (-not $drEnabled) {
                try {
                    $targetRules.Remove($drName)
                    [void]$existingNames.Remove($drName)
                    $deletedCount++
                    Write-OMMigrateLog -Message "Invoke-RulesRecreation: Removed broken rule '$drName' for recreation." -Level DEBUG
                } catch {
                    Write-OMMigrateLog -Message "Invoke-RulesRecreation: Could not remove broken rule '$drName': $_" -Level WARN
                }
            }
        } catch { }
    }
    Write-OMMigrateLog -Message "Invoke-RulesRecreation: Pre-deletion pass complete. Removed=$deletedCount broken rule(s)." -Level INFO

    # Find account archive subfolder for Local targets
    $accountArchiveFolder = $null
    if ($ArchiveRootFolder) {
        try {
            $archSubs = $ArchiveRootFolder.Folders
            for ($a = 1; $a -le $archSubs.Count; $a++) {
                $sub = $archSubs.Item($a)
                Register-COMObject -ComObject $sub
                if ($sub.Name -eq $AccountEmail -or $AccountEmail -like "*$($sub.Name)*") {
                    $accountArchiveFolder = $sub
                    break
                }
            }
        }
        catch { }
    }

    # Find IMAP store root for Server targets
    $imapRoot = $null
    try {
        $imapRoot = $targetStore.GetRootFolder()
        Register-COMObject -ComObject $imapRoot
    }
    catch { }

    # Added 2026-07-10, Administrator (multi-archive support, TargetStoreName hardcode
    # fix). Verified against Invoke-DeployConsolidatedRules's proven
    # groupKeyToArchiveFolder resolution pattern (same DisplayName match +
    # GetRootFolder() approach, same Namespace.Stores enumeration) rather than
    # inventing a new mechanism. Unlike that function, this one is already
    # inside a per-rule loop with each rule's own rules_inventory.csv row
    # (RulesInventory/$csvRow) available further below, and TargetStoreName is
    # a guaranteed base column on every row (confirmed in Export-RulesToCSV) --
    # so no pre-built per-group lookup needs to be threaded in from the
    # caller here. Instead, resolved store root folders are cached in this
    # hashtable as distinct TargetStoreName values are encountered while
    # iterating $SourceRules below, so a given store is never resolved twice
    # in the same call even if many rules share the same TargetStoreName.
    $resolvedArchiveStoreFolders = @{}

    # WhatIf path
    if ($Global:OMMigrate.WhatIf) {
        foreach ($srcRule in $SourceRules) {
            Write-OMMigrateLog -Message "WhatIf: Would recreate rule '$($srcRule.RuleName)'" `
                               -Level INFO -WhatIfPrefix
        }
        return [PSCustomObject]@{ Created = $SourceRules.Count; Skipped = 0; Failed = 0 }
    }

    # Live recreation loop
    foreach ($srcRule in $SourceRules) {
        $ruleName = $srcRule.RuleName

        if ($existingNames.Contains($ruleName)) {
            Write-OMMigrateLog -Message "Invoke-RulesRecreation: Rule '$ruleName' already exists -- skipping." -Level DEBUG
            $skipped++
            continue
        }

        try {
            $newRule = $targetRules.Create($ruleName, $srcRule.RuleType)
            Register-COMObject -ComObject $newRule
            $newRule.Enabled = $srcRule.IsEnabled

            # Apply conditions
            if ($srcRule.CondFrom -and $srcRule.CondFrom.Count -gt 0) {
                try {
                    $fromCond = $newRule.Conditions.From
                    $recips   = $fromCond.Recipients
                    foreach ($addr in $srcRule.CondFrom) {
                        if ($addr) {
                            $r = $recips.Add($addr)
                            Register-COMObject -ComObject $r
                            $r.Resolve() | Out-Null
                        }
                    }
                    $fromCond.Enabled = $true
                }
                catch { Write-OMMigrateLog -Message "Rule '$ruleName': Could not set From condition: $_" -Level DEBUG }
            }
            if ($srcRule.CondSubject -and $srcRule.CondSubject.Count -gt 0) {
                try { $sc = $newRule.Conditions.Subject; $sc.Text = $srcRule.CondSubject; $sc.Enabled = $true } catch { }
            }
            if ($srcRule.CondSenderAddress -and $srcRule.CondSenderAddress.Count -gt 0) {
                # FIXED (2026-07-07, Administrator debug session): converted from plain PS
                # dot-notation (read of Conditions.SenderAddress, then plain
                # .Address / .Enabled writes on that COM object) to InvokeMember
                # reflection for both the read and the writes, per this module's
                # standing COM rule (applies to ALL Outlook COM rule object
                # property access -- reads AND writes). This is a write site,
                # not a read site like the other three fixed this session, but
                # the same unreliable dot-notation access pattern applies to
                # writes on a live COM object just as much as reads. Debug
                # logging added so the next live run can confirm the write
                # actually lands correctly (re-read the value back after
                # writing, via InvokeMember, to verify).
                try {
                    $saCondWrite = $newRule.Conditions.GetType().InvokeMember(
                        'SenderAddress',
                        [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Conditions, $null
                    )
                    if ($null -ne $saCondWrite) {
                        [void]$saCondWrite.GetType().InvokeMember(
                            'Address',
                            [System.Reflection.BindingFlags]::SetProperty,
                            $null, $saCondWrite, @(,[string[]]$srcRule.CondSenderAddress)
                        )
                        [void]$saCondWrite.GetType().InvokeMember(
                            'Enabled',
                            [System.Reflection.BindingFlags]::SetProperty,
                            $null, $saCondWrite, @($true)
                        )

                        # DEBUG (2026-07-07, Administrator debug session): re-read the
                        # value back immediately after writing (via InvokeMember)
                        # to confirm what actually landed on the live rule.
                        try {
                            $debugRecreateAddrRaw = $saCondWrite.GetType().InvokeMember(
                                'Address',
                                [System.Reflection.BindingFlags]::GetProperty,
                                $null, $saCondWrite, $null
                            )
                            $debugRecreateDump = if ($null -eq $debugRecreateAddrRaw) { '<null>' } else { (@($debugRecreateAddrRaw) -join ' | ') }
                            Write-OMMigrateLog -Message "DEBUG Invoke-RulesRecreation SenderAddress write-back (InvokeMember): Rule='$ruleName' Intended=[$($srcRule.CondSenderAddress -join ' | ')] ReadBack=[$debugRecreateDump]" -Level DEBUG
                        } catch { }
                    }
                } catch { }
            }
            if ($srcRule.CondBody -and $srcRule.CondBody.Count -gt 0) {
                try { $bc = $newRule.Conditions.Body; $bc.Text = $srcRule.CondBody; $bc.Enabled = $true } catch { }
            }

            # Resolve folder target
            # Source of truth for TargetFolderPath is rules_inventory.csv ($RulesInventory).
            # When a CSV row exists for this rule (matched by RuleName+ExecutionOrder),
            # use its TargetFolderPath. Fall back to the PST COM path ($srcRule.TargetFolderPath)
            # only when no CSV row is found -- covers rules not in the inventory.
            $newFolderCOM   = $null
            $folderResolved = $false

            # Look up CSV row for this rule by RuleName+ExecutionOrder
            $csvRow = $null
            if ($RulesInventory.Count -gt 0) {
                $csvRow = $RulesInventory | Where-Object {
                    $_.RuleName -eq $ruleName -and
                    [string]$_.ExecutionOrder -eq [string]$srcRule.ExecutionOrder
                } | Select-Object -First 1
                # Fallback: match by name only if ExecutionOrder differs
                if (-not $csvRow) {
                    $csvRow = $RulesInventory | Where-Object {
                        $_.RuleName -eq $ruleName
                    } | Select-Object -First 1
                }
            }

            # Determine the target folder path -- CSV wins over PST COM
            $targetPath = if ($csvRow -and -not [string]::IsNullOrWhiteSpace($csvRow.TargetFolderPath)) {
                $csvRow.TargetFolderPath
            } else {
                $srcRule.TargetFolderPath
            }

            if ($targetPath -and ($srcRule.HasMoveAction -or $srcRule.HasCopyAction)) {
                $folderEntry = $FolderMap | Where-Object {
                    $_.FolderPath -eq $targetPath
                } | Select-Object -First 1

                if ($folderEntry) {
                    if ($folderEntry.Destination -eq 'Server' -and $imapRoot) {
                        # Server destination -- navigate from IMAP store root.
                        # CSV path includes account email as first segment which
                        # correctly navigates the IMAP store subfolder structure.
                        try {
                            $newFolderCOM = Get-FolderByPath -RootFolder $imapRoot `
                                                             -FolderPath $targetPath `
                                                             -CreateIfMissing $false
                        }
                        catch { }
                    }
                    elseif ($folderEntry.Destination -eq 'Local' -and $ArchiveRootFolder) {
                        # Local destination -- navigate from Archive PST ROOT, not the
                        # account subfolder. The CSV TargetFolderPath already includes the
                        # correct first segment (e.g. user@example.com\Inbox\Folder)
                        # which may belong to a DIFFERENT account than the one being migrated
                        # (rules from admin@example-provider.com can target ameritech folders).
                        # Starting from $ArchiveRootFolder lets the full path navigate correctly.
                        #
                        # Added 2026-07-10, Administrator (multi-archive support, TargetStoreName
                        # hardcode fix). This rule's own CSV row ($csvRow, already resolved
                        # above by RuleName+ExecutionOrder) may carry a TargetStoreName that
                        # points at a DIFFERENT attached archive PST than the single
                        # $ArchiveRootFolder passed into this whole function call -- e.g. this
                        # rule's account was mapped to a second archive PST via the
                        # TargetStoreName picker in Script 00, while $ArchiveRootFolder is
                        # whichever archive the caller (Script 03) opened for THIS account's
                        # own pre-build pass. When $csvRow.TargetStoreName is present and
                        # resolves to a live attached store, that store's root is used
                        # INSTEAD of $ArchiveRootFolder for this one rule -- otherwise
                        # (blank TargetStoreName, store not currently attached, or no CSV
                        # row at all) behavior is completely unchanged: falls straight back
                        # to $ArchiveRootFolder exactly as before this fix, so any caller
                        # not using multi-archive mappings is unaffected.
                        $effectiveArchiveRootForThisRule = $ArchiveRootFolder
                        if ($csvRow -and
                            $csvRow.PSObject.Properties['TargetStoreName'] -and
                            -not [string]::IsNullOrWhiteSpace($csvRow.TargetStoreName)) {

                            $wantedArchiveStoreName = $csvRow.TargetStoreName

                            if (-not $resolvedArchiveStoreFolders.ContainsKey($wantedArchiveStoreName)) {
                                $foundArchiveStoreFolder = $null
                                try {
                                    $liveStoresForRecreation = $Namespace.Stores
                                    for ($rsi = 1; $rsi -le $liveStoresForRecreation.Count; $rsi++) {
                                        $lookupArchiveStore = $liveStoresForRecreation.Item($rsi)
                                        if ($lookupArchiveStore.DisplayName -eq $wantedArchiveStoreName) {
                                            try { $foundArchiveStoreFolder = $lookupArchiveStore.GetRootFolder() } catch { }
                                            break
                                        }
                                    }
                                }
                                catch { }

                                if (-not $foundArchiveStoreFolder) {
                                    Write-OMMigrateLog -Message (
                                        "Invoke-RulesRecreation: Rule '$ruleName' -- TargetStoreName " +
                                        "'$wantedArchiveStoreName' not found among attached stores -- " +
                                        "falling back to the default Archive PST for this rule."
                                    ) -Level WARN
                                }
                                $resolvedArchiveStoreFolders[$wantedArchiveStoreName] = $foundArchiveStoreFolder
                            }

                            if ($resolvedArchiveStoreFolders[$wantedArchiveStoreName]) {
                                $effectiveArchiveRootForThisRule = $resolvedArchiveStoreFolders[$wantedArchiveStoreName]
                            }
                        }

                        try {
                            $newFolderCOM = Get-FolderByPath -RootFolder $effectiveArchiveRootForThisRule `
                                                             -FolderPath $targetPath `
                                                             -CreateIfMissing $false
                        }
                        catch { }
                    }
                    if ($newFolderCOM) { $folderResolved = $true }
                    else {
                        Write-OMMigrateLog -Message "Rule '$ruleName': Folder not found in destination ('$targetPath' -> $($folderEntry.Destination)). Rule created without folder action." -Level WARN
                    }
                }
                else {
                    Write-OMMigrateLog -Message "Rule '$ruleName': No folder_map.csv entry for '$targetPath'. Rule created without folder action." -Level WARN
                }
            }

            # Apply folder actions
            if ($srcRule.HasMoveAction) {
                try {
                    $moveAction = $newRule.Actions.MoveToFolder
                    if ($folderResolved -and $newFolderCOM) {
                        # Use InvokeMember reflection -- standard PS assignment silently
                        # fails due to CLR stripping the COM pointer structure.
                        $folderSet          = Set-RuleFolderAction -Action $moveAction -Folder $newFolderCOM
                        $moveAction.Enabled = $folderSet
                        Write-OMMigrateLog -Message "Invoke-RulesRecreation: MoveToFolder InvokeMember result=$folderSet for '$ruleName'." -Level DEBUG
                    }
                }
                catch { }
            }
            if ($srcRule.HasCopyAction) {
                try {
                    $copyAction = $newRule.Actions.CopyToFolder
                    if ($folderResolved -and $newFolderCOM) {
                        # Use InvokeMember reflection -- same CLR bypass as MoveToFolder.
                        $folderSet          = Set-RuleFolderAction -Action $copyAction -Folder $newFolderCOM
                        $copyAction.Enabled = $folderSet
                        Write-OMMigrateLog -Message "Invoke-RulesRecreation: CopyToFolder InvokeMember result=$folderSet for '$ruleName'." -Level DEBUG
                    }
                }
                catch { }
            }

            # Apply non-folder actions
            if ($srcRule.ActionDelete)       { try { $newRule.Actions.Delete.Enabled              = $true } catch { } }
            if ($srcRule.ActionMarkRead)     { try { $newRule.Actions.MarkAsRead.Enabled           = $true } catch { } }
            if ($srcRule.ActionCategories -and $srcRule.ActionCategories.Count -gt 0) {
                try { $ca = $newRule.Actions.AssignToCategory; $ca.Categories = $srcRule.ActionCategories; $ca.Enabled = $true } catch { }
            }
            # Updated 2026-06-29 (Gemini consult + Administrator review): replaced the
            # June 17-era Actions.Item(27) fixed-index write with the named
            # property 'Stop', resolved via InvokeMember GetProperty reflection --
            # the same mechanism already proven and in production use on the
            # current consolidated-rules write path (Invoke-BuildRulesFromMap,
            # 2026-06-26). Item(27) relied on a hardcoded action-slot offset
            # that is fragile to rule-structure or profile-context shifts; the
            # named property is the safer, more accurate accessor and is no
            # longer null via COM interop once the rule has actually been Saved.
            if ($srcRule.ActionStopProcessing) {
                try {
                    $spAction = $newRule.Actions.GetType().InvokeMember(
                        'Stop',
                        [System.Reflection.BindingFlags]::GetProperty,
                        $null, $newRule.Actions, $null
                    )
                    if ($null -ne $spAction) {
                        [void]$spAction.GetType().InvokeMember(
                            'Enabled',
                            [System.Reflection.BindingFlags]::SetProperty,
                            $null, $spAction, @($true)
                        )
                    }
                } catch { }
            }

            $created++
            Write-OMMigrateLog -Message "Rule created: '$ruleName' | FolderResolved=$folderResolved" -Level DEBUG
        }
        catch {
            $failed++
            Write-OMMigrateLog -Message "Invoke-RulesRecreation: Failed to create rule '$ruleName': $_" -Level WARN
        }
    }

    # Save rules collection -- use InvokeMember reflection, NOT direct COM call.
    # Direct $targetRules.Save() goes through PS COM dispatch which silently
    # succeeds but does not persist changes. InvokeMember calls Save() directly
    # via native COM binding, replicating the pattern confirmed working in
    # Test-RulesUpdate.ps1.
    if ($created -gt 0) {
        try {
            $targetRules.GetType().InvokeMember("Save", [System.Reflection.BindingFlags]::InvokeMethod, $null, $targetRules, @($true))
            Write-OMMigrateLog -Message "Invoke-RulesRecreation: Rules saved. Created=$created | Skipped=$skipped | Failed=$failed" -Level INFO
        }
        catch {
            Write-OMMigrateLog -Message "Invoke-RulesRecreation: Rules Save() failed: $_" -Level ERROR
            return [PSCustomObject]@{ Created = 0; Skipped = $skipped; Failed = ($created + $failed) }
        }
    }
    else {
        Write-OMMigrateLog -Message "Invoke-RulesRecreation: No new rules created. Skipped=$skipped | Failed=$failed" -Level INFO
    }

    return [PSCustomObject]@{ Created = $created; Skipped = $skipped; Failed = $failed }
}


Export-ModuleMember -Function @(

    # COM session management
    'Connect-OutlookCOM'
    'Close-OutlookIfRunning'
    'Get-OutlookNamespace'
    'Release-OutlookCOM'
    'Register-COMObject'
    'Suspend-OutlookSendReceive'
    'Resume-OutlookSendReceive'

    # Account enumeration
    'Get-OutlookAccountsViaCOM'

    # Folder tree
    'Get-FolderTree'
    'Get-FolderTreeRecursive'
    'Get-StoreType'
    'Export-FolderMapCSV'
    'Invoke-FolderMapPicker'
    'Get-FolderFullPath'
    'Remove-StorePrefix'

    # Rules inventory
    'Get-OutlookRules'
    'Get-RuleConditionsSummary'
    'Get-RuleActionsSummary'
    'Invoke-NormalizeRulesExecutionOrder'
    'Export-RulesToCSV'
    'Invoke-RulesInventoryPicker'
    'Add-RulesCsvSeparatorRows'

    # Account operations (Scripts 02/03)
    'Remove-POP3Account'
    'Show-IMAPAccountSetupReference'
    'Add-IMAPAccount'

    # Folder navigation helpers (Scripts 01, 03)
    'Get-OrCreateFolder'
    'Get-FolderByPath'

    # PST operations (Scripts 01, 03)
    'Open-PSTFile'
    'Close-PSTFile'
    'Test-PSTAlreadyMounted'

    # Rules recreation (Script 03) / shared extraction layer (Script 04)
    'Export-RulesBlob'
    'Invoke-PurgeAndRecreateRules'
    'Invoke-DeployConsolidatedRules'
    'Invoke-ResortRulesByLabel'
    'Invoke-BuildRulesFromMap'
    'Invoke-FlushPendingRuleSave'
    'Set-RuleConditions'
    'Set-RuleFolderAction'
    'Read-RulesFromPSTStore'
    'Invoke-RulesRecreation'

    # Script 04 artifact extraction stubs -- implemented in Script 04 session:
    # 'Read-CalendarFromPSTStore'
    # 'Read-ContactsFromPSTStore'
    # 'Read-TasksFromPSTStore'
    # 'Read-NotesFromPSTStore'
    # 'Read-JournalFromPSTStore'
)
# ***** END OF FILE *****
