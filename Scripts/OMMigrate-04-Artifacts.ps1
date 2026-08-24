#Requires -Version 5.1
<#
.SYNOPSIS
    OMMigrate-04-Artifacts.ps1 -- Personal Artifacts Migration

.DESCRIPTION
    Step 04 of the OutlookMailMigrator (OMMigrate) toolkit.

    Migrates personal artifacts (Calendar, Contacts, Tasks, Notes, Journal)
    from each account's backup PST into the live IMAP store so they are
    accessible on all devices via the mail server.

    WHAT THIS SCRIPT DOES:
        1. Reads Step 00 and Step 02 manifests (gate checks)
        2. For each selected account:
               a. Opens the backup PST as a read-only source store
               b. For each artifact type (Calendar, Contacts, Tasks,
                  Notes, Journal):
                  - Builds a deduplication key set from the live IMAP store
                  - Copies only items NOT already present in the IMAP store
                  - Logs counts: found / already exists / copied / failed
               c. Detaches backup PST when done
        3. Generates the final Artifacts HTML report
        4. Writes Step 04 manifest -- artifacts migration complete

    ARTIFACT TYPES:
        Calendar  -- Appointments, meetings, recurring events
        Contacts  -- Address book entries
        Tasks     -- To-do items and task requests
        Notes     -- Sticky notes
        Journal   -- Journal entries (calls, emails, meetings logged)

    DEDUPLICATION:
        Each artifact type uses a content-based key to detect duplicates:
            Calendar  : Subject + Start + Duration
                       (GlobalAppointmentID changes between store types)
            Contacts  : FullName + Email1Address
            Tasks     : Subject + DueDate
            Notes     : Subject + CreationTime
            Journal   : Subject + Start + Type

        Items whose key already exists in the live IMAP store are skipped.
        Items with a new key are copied from the backup PST.
        This makes the script safe to re-run -- existing items are never
        duplicated regardless of how many times the script runs.

    ITEM COPY MECHANISM:
        Uses the confirmed .Copy().Move() pattern (same as Script 03 folder
        migration). For artifact item property reads that may be affected
        by the PowerShell COM CLR translation layer, InvokeMember reflection
        is used to bypass silent failures.

    SOURCE OF TRUTH:
        The live IMAP store is the source of truth after migration.
        Items flow one direction only: backup PST -> IMAP store.
        The backup PST is never modified.

.PARAMETER BasePath
    Override the default working directory.
    Must match BasePath used in Scripts 00, 01, 02, and 03.
    Default: $env:USERPROFILE\Documents\OutlookMigration

.PARAMETER LogLevel
    Logging verbosity: DEBUG | INFO | WARN | ERROR
    Default: INFO

.PARAMETER Preview
    Simulate all operations. No items copied.
    All actions logged with [WHATIF] prefix.

.PARAMETER Force
    Skip per-account Y/N confirmation prompts.
    Pre-flight confirmation is still required.

.PARAMETER Sanitize
    Mask personal data (email addresses, names) in console output.
    Log file always contains full data.

.EXAMPLE
    # Standard run
    .\OMMigrate-04-Artifacts.ps1

.EXAMPLE
    # Dry run -- see what would be copied without making changes
    .\OMMigrate-04-Artifacts.ps1 -Preview

.EXAMPLE
    # Skip per-account confirmations
    .\OMMigrate-04-Artifacts.ps1 -Force

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
        Backups\<email>.pst        -- Backup PST per account

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
    -ScriptName 'OMMigrate-04-Artifacts' `
    -BasePath   $BasePath `
    -LogLevel   $LogLevel `
    -IsWhatIf   $Script:IsWhatIf `
    -Sanitize   $Sanitize.IsPresent

Register-ExitHandlers -ScriptStep 4
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
$Script:COMSessionOpen  = $false
$Script:FinalStatus     = 'SUCCESS'
$Script:ReportFile      = ''
$Script:AccountResults  = [System.Collections.Generic.List[PSCustomObject]]::new()

# Artifact totals across all accounts
$Script:TotalCalendarCopied  = 0
$Script:TotalContactsCopied  = 0
$Script:TotalTasksCopied     = 0
$Script:TotalNotesCopied     = 0
$Script:TotalJournalCopied            = 0
$Script:TotalContactSubfoldersCopied  = 0
$Script:TotalItemsFailed              = 0


# ============================================================
#  REGION: OlDefaultFolders CONSTANTS
#  Numeric values for GetDefaultFolder() -- PS 5.1 cannot use
#  the named enum constants directly in all COM contexts.
# ============================================================

$Script:olFolderCalendar = 9
$Script:olFolderContacts = 10
$Script:olFolderJournal  = 11
$Script:olFolderNotes    = 12
$Script:olFolderTasks    = 13


# ============================================================
#  HELPER: Get-ArtifactFolder
#  Returns the MAPI folder for a given artifact type from a
#  store, using the olDefaultFolder constant as a fallback
#  when the store is the default delivery store, or by name
#  search when it is a secondary store (backup PST).
# ============================================================

function Get-ArtifactFolder {
    <#
    .SYNOPSIS
        Returns the MAPI folder for a given artifact type from
        either a live IMAP store or a backup PST store.

    .DESCRIPTION
        For the default delivery store, uses GetDefaultFolder() which
        correctly returns the 'This computer only' local artifact folders
        on standard IMAP servers (Yahoo, Dovecot, Postfix).

        For secondary stores (backup PSTs), searches by DefaultItemType
        to find folders regardless of display name -- critical for
        custom-named folders (e.g. a Notes folder named anything).

    .PARAMETER Store
        The Outlook Store COM object to search.

    .PARAMETER ArtifactType
        Artifact type name: Calendar | Contacts | Tasks | Notes | Journal

    .PARAMETER Namespace
        The active MAPI namespace (needed for GetDefaultFolder).

    .PARAMETER IsDefaultStore
        Retained for future use. Currently not used to change
        search strategy -- all stores use DefaultItemType search.

    .OUTPUTS
        [System.__ComObject] -- MAPIFolder, or $null if not found.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Store,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Calendar','Contacts','Tasks','Notes','Journal')]
        [string]$ArtifactType,

        [Parameter(Mandatory = $true)]
        [object]$Namespace,

        [Parameter(Mandatory = $false)]
        [bool]$IsDefaultStore = $false
    )

    # Map artifact type to OlDefaultFolders constant and OlItemType constant.
    # OlItemType is used to find folders by their DefaultItemType property --
    # this works regardless of what the folder is named (e.g. a Notes folder
    # named 'Passwords' still has DefaultItemType = olNoteItem = 5).
    $folderConstant = switch ($ArtifactType) {
        'Calendar' { $Script:olFolderCalendar }
        'Contacts' { $Script:olFolderContacts }
        'Tasks'    { $Script:olFolderTasks    }
        'Notes'    { $Script:olFolderNotes    }
        'Journal'  { $Script:olFolderJournal  }
    }

    # OlItemType constants for DefaultItemType matching
    # olAppointmentItem=1, olContactItem=2, olJournalItem=4,
    # olNoteItem=5, olTaskItem=3
    $itemType = switch ($ArtifactType) {
        'Calendar' { 1 }
        'Contacts' { 2 }
        'Tasks'    { 3 }
        'Notes'    { 5 }
        'Journal'  { 4 }
    }

    $standardName = switch ($ArtifactType) {
        'Calendar' { 'Calendar' }
        'Contacts' { 'Contacts' }
        'Tasks'    { 'Tasks'    }
        'Notes'    { 'Notes'    }
        'Journal'  { 'Journal'  }
    }

    # For the default delivery store, use GetDefaultFolder() -- this correctly
    # returns the 'This computer only' local artifact folders which are the
    # only artifact folders available on standard IMAP servers (Yahoo, Dovecot,
    # Postfix). These servers support email only -- Calendar/Contacts/Notes/
    # Tasks/Journal are stored locally regardless of IMAP connection.
    if ($IsDefaultStore) {
        try {
            $folder = $Namespace.GetDefaultFolder($folderConstant)
            Register-COMObject -ComObject $folder
            return $folder
        }
        catch {
            Write-OMMigrateLog -Message "GetDefaultFolder($ArtifactType) failed: $_ -- falling back to type search" `
                               -Level WARN
        }
    }

    # For secondary stores (backup PSTs), search by DefaultItemType first.
    # This finds folders regardless of their display name -- critical for
    # custom-named folders (e.g. a Notes folder named 'Passwords').
    # Falls back to standard English name match if type search finds nothing.
    try {
        $root = $Store.GetRootFolder()
        Register-COMObject -ComObject $root
        $subFolders = $root.Folders

        # Pass 1: match by DefaultItemType -- name-independent
        $typeMatch = $null
        for ($i = 1; $i -le $subFolders.Count; $i++) {
            $sub = $subFolders.Item($i)
            Register-COMObject -ComObject $sub
            $subItemType = 0
            try { $subItemType = $sub.DefaultItemType } catch { }
            if ($subItemType -eq $itemType) {
                # For types that may have multiple folders (e.g. multiple
                # calendar folders), prefer one with items over empty ones.
                # For Notes/Tasks/Journal take the first match with items,
                # fall back to first match if all are empty.
                if (-not $typeMatch) { $typeMatch = $sub }
                if ($sub.Items.Count -gt 0) {
                    $typeMatch = $sub
                    break
                }
            }
        }
        if ($typeMatch) {
            Write-OMMigrateLog -Message "Found $ArtifactType folder by type: '$($typeMatch.Name)' ($($typeMatch.Items.Count) items)" `
                               -Level DEBUG
            return $typeMatch
        }

        # Pass 2: fallback to standard English name match
        for ($i = 1; $i -le $subFolders.Count; $i++) {
            $sub = $subFolders.Item($i)
            Register-COMObject -ComObject $sub
            if ($sub.Name -eq $standardName) {
                return $sub
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Folder search failed for '$ArtifactType' in store '$($Store.DisplayName)': $_" `
                           -Level WARN
    }

    return $null
}


# ============================================================
#  HELPER: Get-DeduplicationKey
#  Builds a string key for an Outlook item used to detect
#  duplicates across stores. Uses InvokeMember reflection
#  for property reads that may silently fail via direct PS
#  COM access due to CLR translation layer issues.
# ============================================================

function Get-DeduplicationKey {
    <#
    .SYNOPSIS
        Returns a string deduplication key for an Outlook item.

    .DESCRIPTION
        Each artifact type uses a content-based key that is stable
        across store moves and re-imports:

            Calendar : Subject + Start + Duration
                       GlobalAppointmentID changes when copied between
                       store types (backup PST -> IMAP) so cannot be
                       used for cross-store deduplication. Subject +
                       Start + Duration is stable across store copies.

            Contacts : FullName + '|' + Email1Address
                       EntryID changes on store move, so content-based
                       key is used. FullName alone risks collision on
                       common names; Email1Address makes it unique.

            Tasks    : Subject + '|' + DueDate (formatted)
                       No global ID available. Subject + date is tight
                       enough for practical deduplication.

            Notes    : Subject + '|' + CreationTime (formatted)
                       Notes have few fields. Subject is auto-set to
                       first line of body. CreationTime is the tiebreaker.

            Journal  : Subject + '|' + Start (formatted) + '|' + Type
                       Entry type (phone call, email, meeting etc.) added
                       to handle multiple journal entries on same subject
                       and date.

        Uses InvokeMember reflection for property reads to bypass
        PowerShell COM CLR translation layer issues that can cause
        silent failures on certain Outlook object types.

    .PARAMETER Item
        The Outlook COM item object.

    .PARAMETER ArtifactType
        Artifact type: Calendar | Contacts | Tasks | Notes | Journal

    .OUTPUTS
        [string] -- Deduplication key, or empty string on failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Calendar','Contacts','Tasks','Notes','Journal')]
        [string]$ArtifactType
    )

    # Helper: read a property via InvokeMember to bypass CLR translation.
    # Falls back to direct PS access if reflection fails.
    # Returns empty string on any error -- caller handles missing keys.
    $readProp = {
        param($obj, $propName)
        try {
            return $obj.GetType().InvokeMember(
                $propName,
                [System.Reflection.BindingFlags]::GetProperty,
                $null, $obj, $null
            )
        }
        catch {
            # Fallback to direct PS property access
            try { return $obj.$propName }
            catch { return '' }
        }
    }

    try {
        switch ($ArtifactType) {

            'Calendar' {
                # GlobalAppointmentID cannot be used for cross-store dedup --
                # it changes when an appointment is copied between store types
                # (confirmed: backup PST -> IMAP store produces different IDs).
                # Use Subject + Start + Duration as a stable composite key.
                # Together these three fields uniquely identify an appointment
                # and survive store copies intact.
                $subj     = & $readProp $Item 'Subject'
                $start    = & $readProp $Item 'Start'
                $duration = & $readProp $Item 'Duration'
                $startFmt = if ($start) {
                    try { ([datetime]$start).ToString('yyyyMMddHHmm') } catch { 'nostart' }
                } else { 'nostart' }
                # Strip 'Copy: ' prefix that Outlook adds when copying appointments
                # via COM .Copy().Move() -- ensures dedup key matches between
                # source (no prefix) and destination (may have prefix).
                $subjClean = if ($subj) {
                    $subj.Trim() -replace '^Copy:\s*', ''
                } else { '' }
                $subjKey  = $subjClean.ToLower()
                $durKey   = if ($duration) { $duration.ToString() } else { '0' }
                return "CAL|$subjKey|$startFmt|$durKey"
            }

            'Contacts' {
                $name  = & $readProp $Item 'FullName'
                $email = & $readProp $Item 'Email1Address'
                # Normalize: trim and lower so case differences don't cause
                # false "not a duplicate" results
                $name  = if ($name)  { $name.Trim().ToLower()  } else { '' }
                $email = if ($email) { $email.Trim().ToLower() } else { '' }
                return "CON|$name|$email"
            }

            'Tasks' {
                $subj    = & $readProp $Item 'Subject'
                $dueDate = & $readProp $Item 'DueDate'
                $dueFmt  = if ($dueDate -and $dueDate -ne [datetime]::MinValue) {
                    try { ([datetime]$dueDate).ToString('yyyyMMdd') } catch { 'nodate' }
                } else { 'nodate' }
                return "TSK|$subj|$dueFmt"
            }

            'Notes' {
                # Use Subject + Body for dedup -- two notes with identical content
                # are duplicates regardless of creation time. Body is truncated to
                # first 100 chars to handle minor whitespace differences at the end
                # while still being specific enough to distinguish different notes.
                $subj = & $readProp $Item 'Subject'
                $body = & $readProp $Item 'Body'
                $bodyKey = if ($body) {
                    $body.Trim().Substring(0, [Math]::Min(100, $body.Trim().Length)).ToLower()
                } else { '' }
                return "NOT|$subj|$bodyKey"
            }

            'Journal' {
                $subj  = & $readProp $Item 'Subject'
                $start = & $readProp $Item 'Start'
                $type  = & $readProp $Item 'Type'
                $startFmt = if ($start) {
                    try { ([datetime]$start).ToString('yyyyMMddHHmm') } catch { 'notime' }
                } else { 'notime' }
                return "JRN|$subj|$startFmt|$type"
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Get-DeduplicationKey failed for $ArtifactType item: $_" `
                           -Level DEBUG
        return ''
    }

    return ''
}


# ============================================================
#  HELPER: Build-DeduplicationSet
#  Scans all items in a destination folder and builds a
#  HashSet of deduplication keys for fast O(1) lookup.
# ============================================================

function Build-DeduplicationSet {
    <#
    .SYNOPSIS
        Builds a HashSet of deduplication keys from all items
        in an Outlook folder.

    .DESCRIPTION
        Scans every item in the destination folder and computes
        its deduplication key. Returns a HashSet for O(1) lookup
        when checking whether a source item already exists.

        Called once per artifact type per account before the copy
        loop begins. This is more efficient than checking each
        source item individually against the destination.

    .PARAMETER Folder
        The destination MAPIFolder COM object to scan.

    .PARAMETER ArtifactType
        Artifact type for key generation.

    .OUTPUTS
        [System.Collections.Generic.HashSet[string]]
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Folder,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Calendar','Contacts','Tasks','Notes','Journal')]
        [string]$ArtifactType
    )

    $keySet = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    try {
        $items = $Folder.Items
        $count = $items.Count

        if ($count -eq 0) { return ,$keySet }

        Write-OMMigrateLog -Message "Building dedup set for $ArtifactType : $count item(s) in destination" `
                           -Level DEBUG

        for ($i = 1; $i -le $count; $i++) {
            try {
                $item = $items.Item($i)
                Register-COMObject -ComObject $item
                $key = Get-DeduplicationKey -Item $item -ArtifactType $ArtifactType
                if ($key) { [void]$keySet.Add($key) }
            }
            catch {
                Write-OMMigrateLog -Message "Build-DeduplicationSet: error reading item $i of $count ($ArtifactType): $_" `
                                   -Level DEBUG
            }
        }

        Write-OMMigrateLog -Message "Dedup set built for $ArtifactType : $($keySet.Count) unique keys" `
                           -Level DEBUG
    }
    catch {
        Write-OMMigrateLog -Message "Build-DeduplicationSet failed for $ArtifactType : $_" -Level WARN
    }

    # Comma operator forces PS to return the HashSet as a single object
    # rather than enumerating its contents into the pipeline.
    return ,$keySet
}


# ============================================================
#  HELPER: Copy-ArtifactItems
#  Copies items from a source folder to a destination folder,
#  skipping items whose deduplication key already exists.
# ============================================================

function Copy-ArtifactItems {
    <#
    .SYNOPSIS
        Copies artifact items from source to destination,
        skipping duplicates.

    .DESCRIPTION
        Iterates all items in the source folder. For each item:
            1. Computes its deduplication key
            2. Checks the key against the pre-built HashSet
            3. If key not found: copies item to destination
            4. If key found: skips (already exists in IMAP)

        Uses .Copy().Move() -- the confirmed reliable pattern
        for moving Outlook items between stores via COM.

        The source folder (backup PST) is never modified.

    .PARAMETER SourceFolder
        The source MAPIFolder COM object (backup PST).

    .PARAMETER DestFolder
        The destination MAPIFolder COM object (live IMAP store).

    .PARAMETER ArtifactType
        Artifact type for deduplication key generation.

    .PARAMETER ExistingKeys
        HashSet of keys already in the destination folder.

    .OUTPUTS
        PSCustomObject with Found, AlreadyExists, Copied, Failed counts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$SourceFolder,

        [Parameter(Mandatory = $true)]
        [object]$DestFolder,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Calendar','Contacts','Tasks','Notes','Journal')]
        [string]$ArtifactType,

        [Parameter(Mandatory = $false)]
        [System.Collections.Generic.HashSet[string]]$ExistingKeys = $null
    )

    # Ensure ExistingKeys is always a valid HashSet -- never null or empty collection
    if (-not $ExistingKeys) {
        $ExistingKeys = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )
    }

    $found         = 0
    $alreadyExists = 0
    $copied        = 0
    $failed        = 0

    try {
        $items = $SourceFolder.Items
        $found = $items.Count

        if ($found -eq 0) {
            Write-OMMigrateLog -Message "$ArtifactType source folder is empty -- nothing to copy." `
                               -Level DEBUG
            return [PSCustomObject]@{
                Found         = 0
                AlreadyExists = 0
                Copied        = 0
                Failed        = 0
            }
        }

        Write-OMMigrateLog -Message "$ArtifactType : $found item(s) in source. Checking for duplicates..." `
                           -Level INFO

        if ($Global:OMMigrate.WhatIf) {
            Write-OMMigrateLog -Message "[WHATIF] $ArtifactType : Would process $found item(s)." `
                               -Level INFO -WhatIfPrefix
            return [PSCustomObject]@{
                Found         = $found
                AlreadyExists = 0
                Copied        = $found
                Failed        = 0
            }
        }

        # Iterate source items. Note: when items are moved OUT of a collection
        # during iteration, indices shift. Since we use .Copy() first (which
        # leaves the source intact) and then .Move() the COPY, the source
        # collection is never modified -- index stability is maintained.
        for ($i = 1; $i -le $found; $i++) {
            try {
                $srcItem = $items.Item($i)
                Register-COMObject -ComObject $srcItem

                # Compute deduplication key for this source item
                $key = Get-DeduplicationKey -Item $srcItem -ArtifactType $ArtifactType

                if ($key -and $ExistingKeys.Contains($key)) {
                    # Item already exists in destination -- skip
                    $alreadyExists++
                    Write-OMMigrateLog -Message "$ArtifactType item $i already exists (key=$key) -- skipping." `
                                       -Level DEBUG
                    continue
                }

                # Item is new -- copy it to destination.
                # .Copy() creates a copy in the same folder, then .Move()
                # relocates it to the destination. This is the confirmed
                # reliable pattern for cross-store item migration via COM.
                $copiedItem = $srcItem.Copy()
                Register-COMObject -ComObject $copiedItem

                # Outlook prepends 'Copy: ' to the Subject of copied Calendar
                # and some other item types. Strip it back to the original
                # subject before moving so the destination item has the correct
                # subject and dedup keys match on subsequent re-runs.
                if ($ArtifactType -eq 'Calendar') {
                    try {
                        $origSubj = ''
                        try { $origSubj = $srcItem.Subject } catch { }
                        if ($origSubj) { $copiedItem.Subject = $origSubj }
                        $copiedItem.Save()
                    }
                    catch { }
                }

                $copiedItem.Move($DestFolder) | Out-Null

                # Add new key to set so re-processed items within this run
                # are also deduplicated (defensive -- source should be clean)
                if ($key) { [void]$ExistingKeys.Add($key) }

                $copied++

                if ($copied % 50 -eq 0) {
                    Write-OMMigrateLog -Message "$ArtifactType progress: $copied copied so far..." `
                                       -Level DEBUG
                    Write-Host "      Copying $ArtifactType : $copied / $found items..." `
                               -ForegroundColor DarkGray
                }
            }
            catch {
                $failed++
                Write-OMMigrateLog -Message "$ArtifactType item $i copy failed: $_" -Level WARN
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Copy-ArtifactItems failed for $ArtifactType : $_" -Level WARN
    }

    return [PSCustomObject]@{
        Found         = $found
        AlreadyExists = $alreadyExists
        Copied        = $copied
        Failed        = $failed
    }
}


# ============================================================
#  HELPER: Invoke-ContactSubfolderMigration
#  Recursively migrates contact subfolders from a backup PST
#  Contacts folder into the live IMAP store, preserving the
#  full folder hierarchy to unlimited depth.
#
#  FOLDER IDENTIFICATION (per Gemini MAPI research):
#    PR_SPECIAL_FOLDER_TAG (0x36230003) = 30
#      -> Suggested Contacts system folder. Items migrated into
#         the destination auto-created local contacts folder.
#         No new named folder created.
#    Property missing/throws (user-created GUI folder):
#      -> All contact subfolders in a PST created via the GUI
#         have empty ContainerClass by design (legacy inherited
#         context). Create matching named subfolder in destination
#         using Folders.Add(Name, 2) which sets IPF.Contact on
#         the IMAP side. Recurse into child subfolders.
#    Fallback name check (old PST format / corrupt store):
#      -> If PR_SPECIAL_FOLDER_TAG is missing on ALL folders,
#         fall back to name match 'Suggested Contacts'.
#
#  PATH LENGTH GUARD:
#    IMAP servers enforce a ~255 character absolute path limit.
#    Folder creation is skipped with a WARN if the path would
#    exceed 240 characters (safety margin).
#
#  COM CLEANUP:
#    All created folder COM references are released on unwind
#    to prevent RPC_E_SERVERFAULT crashes at depth.
# ============================================================

function Invoke-ContactSubfolderMigration {
    <#
    .SYNOPSIS
        Recursively migrates contact subfolders from a backup PST
        to a live IMAP store, preserving folder hierarchy.

    .PARAMETER SourceFolder
        The source MAPIFolder (Contacts root from backup PST).

    .PARAMETER DestLocalFolder
        The destination auto-created local contacts folder.
        Target for Suggested Contacts items.

    .PARAMETER DestParentFolder
        The destination parent folder for named subfolder creation.
        On first call this is the IMAP Contacts root.

    .PARAMETER Email
        Account email address for log messages.

    .PARAMETER CurrentPath
        Current IMAP folder path string for path length guard.

    .OUTPUTS
        [hashtable] with Found, AlreadyExists, Copied, Failed counts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]  [object]$SourceFolder,
        [Parameter(Mandatory = $true)]  [object]$DestLocalFolder,
        [Parameter(Mandatory = $true)]  [object]$DestParentFolder,
        [Parameter(Mandatory = $true)]  [string]$Email,
        [Parameter(Mandatory = $false)] [string]$CurrentPath = ''
    )

    # MAPI property tag for PR_SPECIAL_FOLDER_TAG
    # Value 30 = olFolderSuggestedContacts -- stamped directly on the folder
    # object inside the PST, travels with the PST regardless of profile.
    # User-created folders do not have this property -- reading it throws
    # MAPI_E_NOT_FOUND (0x8004010F), which is the positive signal for
    # user-created folders.
    $PR_SPECIAL_FOLDER_TAG    = 'http://schemas.microsoft.com/mapi/proptag/0x36230003'
    $SUGGESTED_CONTACTS_TAG   = 30

    # MAPI property tag for PR_CONTAINER_CLASS -- used to stamp destination
    # folders created via Folders.Add() on IMAP stores
    $PR_CONTAINER_CLASS       = 'http://schemas.microsoft.com/mapi/proptag/0x3613001E'

    $totals = @{ Found = 0; AlreadyExists = 0; Copied = 0; Failed = 0 }

    $srcSubs = $null
    try { $srcSubs = $SourceFolder.Folders } catch { return $totals }
    if (-not $srcSubs -or $srcSubs.Count -eq 0) { return $totals }

    for ($si = 1; $si -le $srcSubs.Count; $si++) {
        $srcSub = $null
        try { $srcSub = $srcSubs.Item($si) } catch { continue }
        Register-COMObject -ComObject $srcSub

        $subName = ''
        try { $subName = $srcSub.Name } catch { }

        # --------------------------------------------------------
        # Identify folder type via PR_SPECIAL_FOLDER_TAG
        # --------------------------------------------------------
        $isSuggestedContacts = $false
        try {
            $tagValue = $srcSub.PropertyAccessor.GetProperty($PR_SPECIAL_FOLDER_TAG)
            if ($tagValue -eq $SUGGESTED_CONTACTS_TAG) {
                $isSuggestedContacts = $true
            }
        }
        catch {
            # MAPI_E_NOT_FOUND -- property absent -- user-created folder
            # Check fallback: name match for old PST formats
            if ($subName -eq 'Suggested Contacts') {
                $isSuggestedContacts = $true
                Write-OMMigrateLog -Message "Contact Subfolders: '$subName' identified as Suggested Contacts via name fallback." `
                                   -Level DEBUG
            }
        }

        # --------------------------------------------------------
        # Route by folder type
        # --------------------------------------------------------
        if ($isSuggestedContacts) {
            # Suggested Contacts system folder -- migrate items into
            # destination local folder, no new folder created
            Write-OMMigrateLog -Message "Contact Subfolders: '$subName' is Suggested Contacts -- migrating items into local destination folder." `
                               -Level DEBUG

            if ($srcSub.Items.Count -gt 0) {
                $keys = Build-DeduplicationSet -Folder $DestLocalFolder -ArtifactType 'Contacts'
                if (-not $keys) {
                    $keys = [System.Collections.Generic.HashSet[string]]::new(
                        [System.StringComparer]::OrdinalIgnoreCase
                    )
                }
                $r = Copy-ArtifactItems -SourceFolder $srcSub -DestFolder $DestLocalFolder `
                                        -ArtifactType 'Contacts' -ExistingKeys $keys
                $totals.Found         += $r.Found
                $totals.AlreadyExists += $r.AlreadyExists
                $totals.Copied        += $r.Copied
                $totals.Failed        += $r.Failed
                Write-OMMigrateLog -Message "Contact Subfolders: '$subName' (Suggested Contacts): Found=$($r.Found) AlreadyExists=$($r.AlreadyExists) Copied=$($r.Copied) Failed=$($r.Failed)" `
                                   -Level INFO
            }

        } else {
            # User-created contact subfolder -- find or create named
            # subfolder in destination and migrate items

            # Path length guard
            $childPath = if ($CurrentPath) { "$CurrentPath/$subName" } else { $subName }
            if ($childPath.Length -gt 240) {
                Write-OMMigrateLog -Message "Contact Subfolders: '$subName' skipped -- IMAP path would exceed 240 chars ('$childPath')." `
                                   -Level WARN
                continue
            }

            # Find or create destination subfolder
            $destSub = $null
            $created = $false
            try {
                $destSubs = $DestParentFolder.Folders
                for ($di = 1; $di -le $destSubs.Count; $di++) {
                    $d = $destSubs.Item($di)
                    if ($d.Name -eq $subName -or $d.Name -like "$subName (*)") {
                        $destSub = $d
                        Register-COMObject -ComObject $destSub
                        break
                    }
                }
            } catch { }

            if (-not $destSub) {
                try {
                    # Folders.Add(Name) with no type -- IMAP provider rejects type parameter
                    # (E_INVALIDARG); IPF.Contact is stamped explicitly via PropertyAccessor below
                    $destSub = $DestParentFolder.Folders.Add($subName)
                    Register-COMObject -ComObject $destSub
                    $created = $true
                    # Verify ContainerClass set correctly; stamp if not
                    try {
                        $cc = $destSub.PropertyAccessor.GetProperty($PR_CONTAINER_CLASS)
                        if ($cc -ne 'IPF.Contact') {
                            $destSub.PropertyAccessor.SetProperty($PR_CONTAINER_CLASS, 'IPF.Contact')
                        }
                    } catch { }
                    Write-OMMigrateLog -Message "Contact Subfolders: created destination subfolder '$subName' at '$childPath'." `
                                       -Level INFO
                }
                catch {
                    Write-OMMigrateLog -Message "Contact Subfolders: could not create '$subName' at '$childPath': $_ -- skipping." `
                                       -Level WARN
                    continue
                }
            }

            # Migrate items into this subfolder
            if ($srcSub.Items.Count -gt 0) {
                $keys = Build-DeduplicationSet -Folder $destSub -ArtifactType 'Contacts'
                if (-not $keys) {
                    $keys = [System.Collections.Generic.HashSet[string]]::new(
                        [System.StringComparer]::OrdinalIgnoreCase
                    )
                }
                $r = Copy-ArtifactItems -SourceFolder $srcSub -DestFolder $destSub `
                                        -ArtifactType 'Contacts' -ExistingKeys $keys
                $totals.Found         += $r.Found
                $totals.AlreadyExists += $r.AlreadyExists
                $totals.Copied        += $r.Copied
                $totals.Failed        += $r.Failed
                Write-OMMigrateLog -Message "Contact Subfolders: '$subName': Found=$($r.Found) AlreadyExists=$($r.AlreadyExists) Copied=$($r.Copied) Failed=$($r.Failed)" `
                                   -Level INFO
            }

            # Recurse into child subfolders
            $childTotals = Invoke-ContactSubfolderMigration `
                -SourceFolder     $srcSub `
                -DestLocalFolder  $DestLocalFolder `
                -DestParentFolder $destSub `
                -Email            $Email `
                -CurrentPath      $childPath
            $totals.Found         += $childTotals.Found
            $totals.AlreadyExists += $childTotals.AlreadyExists
            $totals.Copied        += $childTotals.Copied
            $totals.Failed        += $childTotals.Failed

            # Release created folder COM reference on unwind
            if ($created) {
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($destSub) | Out-Null } catch { }
            }
        }
    }

    return $totals
}


# ============================================================
#  HELPER: Invoke-AccountArtifactMigration
#  Orchestrates the full artifact migration for one account.
# ============================================================

function Invoke-AccountArtifactMigration {
    <#
    .SYNOPSIS
        Migrates all artifact types for a single account from
        its backup PST into the live IMAP store.

    .DESCRIPTION
        For a given account:
            1. Locates the backup PST in the Backups folder
            2. Opens the backup PST as a source store
            3. Finds the live IMAP store in the active COM session
            4. For each artifact type:
               - Gets source folder from backup PST
               - Gets destination folder from live IMAP store
               - Builds deduplication key set from destination
               - Copies new items only
            5. Detaches the backup PST

    .PARAMETER Account
        Account object from migration_accounts.csv.

    .PARAMETER Namespace
        Active Outlook MAPI namespace COM object.

    .OUTPUTS
        PSCustomObject with per-artifact counts and overall Outcome.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Account,

        [Parameter(Mandatory = $true)]
        [object]$Namespace
    )

    $email     = $Account.EmailAddress
    $safeEmail = Get-SafeFileName -InputString $email
    $backupPath = Join-Path $Global:OMMigrate.BackupPath "$safeEmail.pst"

    # Result object -- tracks counts per artifact type
    $result = [PSCustomObject]@{
        EmailAddress     = $email
        CalendarFound    = 0; CalendarExists = 0; CalendarCopied = 0; CalendarFailed = 0
        ContactsFound    = 0; ContactsExists = 0; ContactsCopied = 0; ContactsFailed = 0
        SubfolderFound   = 0; SubfolderExists = 0; SubfolderCopied = 0; SubfolderFailed = 0
        TasksFound       = 0; TasksExists    = 0; TasksCopied    = 0; TasksFailed    = 0
        NotesFound       = 0; NotesExists    = 0; NotesCopied    = 0; NotesFailed    = 0
        JournalFound     = 0; JournalExists  = 0; JournalCopied  = 0; JournalFailed  = 0
        TotalCopied      = 0
        TotalFailed      = 0
        Outcome          = 'SUCCESS'
        Detail           = ''
    }

    Write-OMMigrateLog -Message "Starting artifact migration for: $email" -Level INFO

    # ADDED 2026-07-11, Administrator direction: accounts that were ALWAYS IMAP (e.g.
    # a mailbox created directly as IMAP on Administrator's own Dovecot/Postfix server,
    # never converted from POP3) structurally never have a backup PST -- there
    # was never a POP3 PST to back up in the first place. This is NOT an error
    # condition and must never surface as WARN: per Administrator's direction, WARN/
    # ERROR in these scripts means "admin must take action," and there is
    # nothing actionable here -- the account's Calendar/Contacts/Tasks/Notes/
    # Journal (if any) already live in their one and only home, the account's
    # own local "This computer only" folders (same local-only folders every
    # IMAP account gets in Outlook, since IMAP itself has no protocol concept
    # of these item types -- confirmed live in this account's own Script 03
    # SERVER-folder listing earlier this session, e.g. "Calendar (This
    # computer only)"). There is nothing to copy FROM and nothing to copy TO
    # that isn't the same place, so instead of the old backup-PST-required
    # path, this reports what (if anything) already exists there and exits
    # informationally at INFO level, never WARN.
    # SCOPE CORRECTION 2026-07-11, Administrator direction: the informational (no-WARN)
    # path below applies ONLY to accounts that were ALWAYS IMAP
    # (ProviderTag='IMAP-ALREADY') and therefore structurally never had a
    # POP3 PST to back up. A POP3-derived account that has gone through
    # Scripts 01/02 (ProviderTag='IMAP-CONVERTED', or any POP3-* tag still
    # pending conversion) SHOULD have a backup PST -- if one is missing for
    # those accounts, that IS an actionable problem (Script 01 failed, wrong
    # path, disk issue, etc.) and must remain a real WARN so the admin
    # catches it. Gating strictly on ProviderTag, not just "file missing."
    $isAlwaysImap = ($Account.ProviderTag -eq 'IMAP-ALREADY')

    if (-not (Test-Path $backupPath) -and $isAlwaysImap) {
        Write-OMMigrateLog -Message "No backup PST for $email (always-IMAP account, never converted from POP3) -- checking live account for existing local artifacts instead." `
                           -Level INFO

        $liveOnlyImapStore = $null
        try {
            $stores = $Namespace.Stores
            Register-COMObject -ComObject $stores
            for ($s = 1; $s -le $stores.Count; $s++) {
                $store = $stores.Item($s)
                Register-COMObject -ComObject $store
                $storeDisplayName = ''
                try { $storeDisplayName = $store.DisplayName } catch { }
                if ($storeDisplayName -like "*$email*") {
                    $storeFilePath = ''
                    try { $storeFilePath = $store.FilePath } catch { }
                    if ($storeFilePath -like '*Backups*' -or $storeFilePath -like '*Archive*') {
                        continue
                    }
                    $liveOnlyImapStore = $store
                    break
                }
            }
        }
        catch {
            Write-OMMigrateLog -Message "Error finding live IMAP store for $email (always-IMAP path): $_" -Level INFO
        }

        if (-not $liveOnlyImapStore) {
            # Genuinely nothing to report against -- no backup PST AND no
            # live store found in this COM session. Still informational, not
            # a WARN -- there is no action for the admin to take; the account
            # simply isn't visible to this run for artifact purposes.
            $result.Outcome = 'SUCCESS'
            $result.Detail  = 'No backup PST and no live IMAP store found -- this account has no accessible Calendar, Contacts, Tasks, Notes, or Journal for this run.'
            Write-OMMigrateLog -Message "$email : $($result.Detail)" -Level INFO
            return $result
        }

        $liveIsDefaultStore = $false
        try {
            $defaultStore = $Namespace.DefaultStore
            if ($defaultStore.StoreID -eq $liveOnlyImapStore.StoreID) { $liveIsDefaultStore = $true }
        }
        catch { }

        $anyFound = $false
        foreach ($artifactType in @('Calendar', 'Contacts', 'Tasks', 'Notes', 'Journal')) {
            $liveFolder = $null
            try {
                $liveFolder = Get-ArtifactFolder `
                    -Store          $liveOnlyImapStore `
                    -ArtifactType   $artifactType `
                    -Namespace      $Namespace `
                    -IsDefaultStore $liveIsDefaultStore
            }
            catch { }

            $liveCount = 0
            if ($liveFolder) {
                try { $liveCount = $liveFolder.Items.Count } catch { }
            }

            switch ($artifactType) {
                'Calendar' { $result.CalendarFound = $liveCount; $result.CalendarExists = $liveCount }
                'Contacts' { $result.ContactsFound = $liveCount; $result.ContactsExists = $liveCount }
                'Tasks'    { $result.TasksFound    = $liveCount; $result.TasksExists    = $liveCount }
                'Notes'    { $result.NotesFound    = $liveCount; $result.NotesExists    = $liveCount }
                'Journal'  { $result.JournalFound  = $liveCount; $result.JournalExists  = $liveCount }
            }

            if ($liveCount -gt 0) {
                $anyFound = $true
                Write-Host "    [$artifactType] $liveCount item(s) already local to this account." -ForegroundColor Gray
            }
        }

        $result.Outcome = 'SUCCESS'
        if ($anyFound) {
            $result.Detail = "Always-IMAP account -- artifacts already local: Cal=$($result.CalendarFound) Con=$($result.ContactsFound) Tsk=$($result.TasksFound) Not=$($result.NotesFound) Jrn=$($result.JournalFound)"
        } else {
            $result.Detail = 'Always-IMAP account -- no Calendar, Contacts, Tasks, Notes, or Journal items found for this account.'
        }
        Write-OMMigrateLog -Message "$email : $($result.Detail)" -Level INFO
        return $result
    }
    elseif (-not (Test-Path $backupPath)) {
        # RESTORED 2026-07-11, Administrator direction: this is the original WARN
        # behavior, now correctly scoped to run only when the account is
        # NOT always-IMAP (i.e. it's POP3-derived and SHOULD have a backup
        # PST from Scripts 01/02). A missing backup here is a real,
        # actionable problem and must stay a WARN.
        Write-OMMigrateLog -Message "Backup PST not found for $email : $backupPath -- skipping." `
                           -Level WARN
        $result.Outcome = 'SKIPPED'
        $result.Detail  = "Backup PST not found: $backupPath"
        return $result
    }

    # Open backup PST as source store
    # Check first whether this PST is already mounted in the profile
    # (e.g. Administrator manually attached it for permanent manual triage use)
    # so it is not detached later as if it were this script's own
    # temporary mount.
    $backupPathWasAlreadyMounted = Test-PSTAlreadyMounted -PSTPath $backupPath
    $backupDisplayName = "Backup -- $email"
    $backupStore = Open-PSTFile -PSTPath     $backupPath `
                                -DisplayName $backupDisplayName

    if (-not $backupStore -and -not $Global:OMMigrate.WhatIf) {
        $result.Outcome = 'FAILED'
        $result.Detail  = "Could not open backup PST: $backupPath"
        return $result
    }

    # Find the live IMAP store for this account
    $imapStore      = $null
    $isDefaultStore = $false
    try {
        $stores = $Namespace.Stores
        Register-COMObject -ComObject $stores
        for ($s = 1; $s -le $stores.Count; $s++) {
            $store = $stores.Item($s)
            Register-COMObject -ComObject $store

            $storeDisplayName = ''
            try { $storeDisplayName = $store.DisplayName } catch { }

            # Match by display name containing email address.
            # Exclude backup PSTs and Archive PST -- we want the live IMAP store.
            if ($storeDisplayName -like "*$email*") {
                $storeFilePath = ''
                try { $storeFilePath = $store.FilePath } catch { }
                # Skip if it's a backup or archive PST
                if ($storeFilePath -like '*Backups*' -or
                    $storeFilePath -like '*Archive*') {
                    continue
                }
                $imapStore = $store

                # Check if this is the default delivery store
                try {
                    $defaultStore = $Namespace.DefaultStore
                    if ($defaultStore.StoreID -eq $store.StoreID) {
                        $isDefaultStore = $true
                    }
                }
                catch { }

                break
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Error finding IMAP store for $email : $_" -Level WARN
    }

    if (-not $imapStore) {
        Write-OMMigrateLog -Message "Live IMAP store not found for $email -- skipping artifact migration." `
                           -Level WARN
        if ($backupStore -and -not $backupPathWasAlreadyMounted) { Close-PSTFile -PSTPath $backupPath | Out-Null }
        $result.Outcome = 'SKIPPED'
        $result.Detail  = 'Live IMAP store not found in active COM session'
        return $result
    }

    Write-OMMigrateLog -Message "IMAP store found for $email (IsDefault=$isDefaultStore)" -Level DEBUG

    # Cache the destination Contacts folder so Suggested Contacts migration
    # uses the same parent folder as the main Contacts copy.
    $cachedContactsDestFolder = $null

    # Process each artifact type
    $artifactTypes = @('Calendar', 'Contacts', 'Tasks', 'Notes', 'Journal')

    foreach ($artifactType in $artifactTypes) {

        Write-Host "    [$artifactType]" -ForegroundColor Cyan -NoNewline
        Write-Host " Scanning..." -ForegroundColor DarkGray

        # Get source folder from backup PST
        $srcFolder = $null
        try {
            $srcFolder = Get-ArtifactFolder `
                -Store          $backupStore `
                -ArtifactType   $artifactType `
                -Namespace      $Namespace `
                -IsDefaultStore $false
        }
        catch {
            Write-OMMigrateLog -Message "$artifactType source folder error for $email : $_" -Level WARN
        }

        if (-not $srcFolder) {
            Write-Host "      [SKIP] No $artifactType folder found in backup PST." `
                       -ForegroundColor DarkGray
            Write-OMMigrateLog -Message "$artifactType : No source folder in backup PST for $email -- skipping type." `
                               -Level INFO
            continue
        }

        $srcCount = 0
        try { $srcCount = $srcFolder.Items.Count } catch { }

        if ($srcCount -eq 0) {
            Write-Host "      [SKIP] $artifactType backup folder is empty." -ForegroundColor DarkGray
            Write-OMMigrateLog -Message "$artifactType : Backup folder empty for $email -- skipping type." `
                               -Level INFO
            continue
        }

        # Get destination folder from live IMAP store
        $destFolder = $null
        try {
            $destFolder = Get-ArtifactFolder `
                -Store          $imapStore `
                -ArtifactType   $artifactType `
                -Namespace      $Namespace `
                -IsDefaultStore $isDefaultStore
        }
        catch {
            Write-OMMigrateLog -Message "$artifactType destination folder error for $email : $_" -Level WARN
        }

        if (-not $destFolder) {
            Write-Host "      [WARN] No $artifactType folder found in live IMAP store -- skipping type." `
                       -ForegroundColor Yellow
            Write-OMMigrateLog -Message "$artifactType : No destination folder in IMAP store for $email -- skipping type." `
                               -Level WARN
            continue
        }

        $destCount = 0
        try { $destCount = $destFolder.Items.Count } catch { }

        Write-Host "      Source: $srcCount item(s) | Destination: $destCount item(s)" `
                   -ForegroundColor Gray

        # Build deduplication set from IMAP destination (source of truth).
        # Comma operator in Build-DeduplicationSet ensures a HashSet is
        # returned -- guard here as well in case of unexpected $null.
        $existingKeys = Build-DeduplicationSet `
            -Folder       $destFolder `
            -ArtifactType $artifactType
        if (-not $existingKeys) {
            $existingKeys = [System.Collections.Generic.HashSet[string]]::new(
                [System.StringComparer]::OrdinalIgnoreCase
            )
        }

        # Copy new items to IMAP destination
        $copyResult = Copy-ArtifactItems `
            -SourceFolder $srcFolder `
            -DestFolder   $destFolder `
            -ArtifactType $artifactType `
            -ExistingKeys $existingKeys



        # Store results
        switch ($artifactType) {
            'Calendar' {
                $result.CalendarFound  = $copyResult.Found
                $result.CalendarExists = $copyResult.AlreadyExists
                $result.CalendarCopied = $copyResult.Copied
                $result.CalendarFailed = $copyResult.Failed
                $Script:TotalCalendarCopied += $copyResult.Copied
            }
            'Contacts' {
                $result.ContactsFound  = $copyResult.Found
                $result.ContactsExists = $copyResult.AlreadyExists
                $result.ContactsCopied = $copyResult.Copied
                $result.ContactsFailed = $copyResult.Failed
                $Script:TotalContactsCopied += $copyResult.Copied
                # Cache destination folder for Suggested Contacts subfolder migration
                $cachedContactsDestFolder = $destFolder
            }
            'Tasks' {
                $result.TasksFound  = $copyResult.Found
                $result.TasksExists = $copyResult.AlreadyExists
                $result.TasksCopied = $copyResult.Copied
                $result.TasksFailed = $copyResult.Failed
                $Script:TotalTasksCopied += $copyResult.Copied
            }
            'Notes' {
                $result.NotesFound  = $copyResult.Found
                $result.NotesExists = $copyResult.AlreadyExists
                $result.NotesCopied = $copyResult.Copied
                $result.NotesFailed = $copyResult.Failed
                $Script:TotalNotesCopied += $copyResult.Copied
            }
            'Journal' {
                $result.JournalFound  = $copyResult.Found
                $result.JournalExists = $copyResult.AlreadyExists
                $result.JournalCopied = $copyResult.Copied
                $result.JournalFailed = $copyResult.Failed
                $Script:TotalJournalCopied += $copyResult.Copied
            }
        }

        $result.TotalCopied += $copyResult.Copied
        $result.TotalFailed += $copyResult.Failed
        $Script:TotalItemsFailed += $copyResult.Failed

        # Console summary for this artifact type
        $color = if ($copyResult.Failed -gt 0) { 'Yellow' }
                 elseif ($copyResult.Copied -gt 0) { 'Green' }
                 else { 'DarkGray' }

        Write-Host ("      Copied: {0} | Skipped: {1} | Failed: {2}" -f `
            $copyResult.Copied, $copyResult.AlreadyExists, $copyResult.Failed) `
            -ForegroundColor $color

        Write-OMMigrateLog -Message (
            "$artifactType for $email : " +
            "Found=$($copyResult.Found) | " +
            "AlreadyExists=$($copyResult.AlreadyExists) | " +
            "Copied=$($copyResult.Copied) | " +
            "Failed=$($copyResult.Failed)"
        ) -Level INFO

        Write-AuditEntry -Action "ARTIFACTS_${artifactType}_MIGRATED" `
                         -AccountEmail $email `
                         -Detail (
                             "Found=$($copyResult.Found) | " +
                             "AlreadyExists=$($copyResult.AlreadyExists) | " +
                             "Copied=$($copyResult.Copied) | " +
                             "Failed=$($copyResult.Failed)"
                         )
    }

    # ----------------------------------------------------------------
    # Contact Subfolders migration
    # Migrates all contact subfolders from the backup PST Contacts
    # folder into the live IMAP store, preserving full folder hierarchy.
    # Handled by Invoke-ContactSubfolderMigration (recursive).
    #
    # SOURCE: backup PST GetDefaultFolder(10) -> Contacts root
    # DESTINATION: IMAP store GetDefaultFolder(10) -> Contacts root
    #   Local destination folder (empty ContainerClass) used as
    #   target for auto-created subfolders.
    # ----------------------------------------------------------------
    try {
        $srcContactsRoot  = $null
        $destContactsRoot = $null
        $destLocalFolder  = $null

        # Get source Contacts root from backup PST
        try {
            $srcContactsRoot = $backupStore.GetDefaultFolder(10)
            Register-COMObject -ComObject $srcContactsRoot
        }
        catch {
            Write-OMMigrateLog -Message "Contact Subfolders: error accessing backup PST Contacts root for $email : $_" `
                               -Level WARN
        }

        # Get destination Contacts root from IMAP store
        try {
            $destContactsRoot = $imapStore.GetDefaultFolder(10)
            Register-COMObject -ComObject $destContactsRoot
        }
        catch {
            Write-OMMigrateLog -Message "Contact Subfolders: error accessing IMAP Contacts root for $email : $_" `
                               -Level WARN
        }

        # Find destination local contacts folder (empty ContainerClass)
        if ($destContactsRoot) {
            try {
                $dSubs = $destContactsRoot.Folders
                for ($di = 1; $di -le $dSubs.Count; $di++) {
                    $dsub = $dSubs.Item($di)
                    Register-COMObject -ComObject $dsub
                    $dsubClass = ''
                    try { $dsubClass = $dsub.ContainerClass } catch { }
                    if ($dsub.DefaultItemType -eq 2 -and [string]::IsNullOrEmpty($dsubClass)) {
                        $destLocalFolder = $dsub
                        break
                    }
                }
            }
            catch { }
        }

        if ($srcContactsRoot -and $destContactsRoot -and $destLocalFolder) {
            # Check if there are any subfolders to process
            $srcSubCount = 0
            try { $srcSubCount = $srcContactsRoot.Folders.Count } catch { }

            if ($srcSubCount -gt 0) {
                Write-Host "    [Contact Subfolders]" -ForegroundColor Cyan -NoNewline
                Write-Host " Scanning..." -ForegroundColor DarkGray
                Write-OMMigrateLog -Message "Contact Subfolders: $srcSubCount subfolder(s) found under backup PST Contacts for $email." `
                                   -Level INFO

                $subResult = Invoke-ContactSubfolderMigration `
                    -SourceFolder     $srcContactsRoot `
                    -DestLocalFolder  $destLocalFolder `
                    -DestParentFolder $destContactsRoot `
                    -Email            $email `
                    -CurrentPath      $destContactsRoot.Name

                $result.SubfolderFound   += $subResult.Found
                $result.SubfolderExists  += $subResult.AlreadyExists
                $result.SubfolderCopied  += $subResult.Copied
                $result.SubfolderFailed  += $subResult.Failed
                $result.TotalCopied      += $subResult.Copied
                $result.TotalFailed      += $subResult.Failed
                $Script:TotalContactSubfoldersCopied += $subResult.Copied
                $Script:TotalItemsFailed             += $subResult.Failed

                $color = if ($subResult.Failed -gt 0) { 'Yellow' }
                         elseif ($subResult.Copied -gt 0) { 'Green' }
                         else { 'DarkGray' }

                Write-Host ("      Copied: {0} | Skipped: {1} | Failed: {2}" -f `
                    $subResult.Copied, $subResult.AlreadyExists, $subResult.Failed) `
                    -ForegroundColor $color

                Write-OMMigrateLog -Message (
                    "Contact Subfolders for $email : " +
                    "Found=$($subResult.Found) | " +
                    "AlreadyExists=$($subResult.AlreadyExists) | " +
                    "Copied=$($subResult.Copied) | " +
                    "Failed=$($subResult.Failed)"
                ) -Level INFO
            }
            else {
                Write-OMMigrateLog -Message "Contact Subfolders: no subfolders found under backup PST Contacts for $email -- skipping." `
                                   -Level DEBUG
            }
        }
        else {
            # Only warn if the source has subfolders but destination is not ready.
            # If source has no subfolders, skip silently at DEBUG -- nothing to do.
            $srcSubCount = 0
            if ($srcContactsRoot) {
                try { $srcSubCount = $srcContactsRoot.Folders.Count } catch { }
            }
            if ($srcSubCount -gt 0) {
                Write-OMMigrateLog -Message "Contact Subfolders: missing destination root or local destination folder for $email -- skipping." `
                                   -Level WARN
            } else {
                Write-OMMigrateLog -Message "Contact Subfolders: no subfolders in backup PST for $email -- skipping." `
                                   -Level DEBUG
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Contact Subfolders migration error for $email : $_" -Level WARN
    }

    # Detach backup PST
    # but only if this script mounted it; leave alone if it was already
    # attached to the profile before this function was called.
    if ($backupStore -and -not $Global:OMMigrate.WhatIf -and -not $backupPathWasAlreadyMounted) {
        Close-PSTFile -PSTPath $backupPath | Out-Null
        Write-OMMigrateLog -Message "Backup PST detached: $backupPath" -Level DEBUG
    }

    # Set overall outcome
    if ($result.TotalFailed -gt 0 -and $result.TotalCopied -eq 0) {
        $result.Outcome = 'FAILED'
    }
    elseif ($result.TotalFailed -gt 0) {
        $result.Outcome = 'WARNING'
    }
    else {
        $result.Outcome = 'SUCCESS'
    }

    $result.Detail = (
        "Cal={0} Con={1} Sug={2} Tsk={3} Not={4} Jrn={5} | Failed={6}" -f
        $result.CalendarCopied,
        $result.ContactsCopied,
        $result.SubfolderCopied,
        $result.TasksCopied,
        $result.NotesCopied,
        $result.JournalCopied,
        $result.TotalFailed
    )

    return $result
}


# ============================================================
#  MAIN EXECUTION BLOCK
# ============================================================

try {

    # ----------------------------------------------------------
    #  STEP 1 -- Gate Checks
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Pre-Flight Gate Checks' -Step '1 of 5'

    # Gate 1: Script 00 manifest
    $step00Manifest = Read-StepManifest -Step 0
    Write-Host "  Script 00 manifest : OK" -ForegroundColor Green

    # Gate 2: Script 02 manifest
    $step02Manifest = Read-StepManifest -Step 2
    Write-Host "  Script 02 manifest : OK ($($step02Manifest.Data.MigratedCount) accounts converted)" `
               -ForegroundColor Green

    # Gate 3: Environment
    $envResult = Test-OMMigrateEnvironment
    if (-not $envResult.Passed) {
        $Script:FinalStatus = 'FAILED'
        exit 1
    }
    Write-Host '  Environment check  : OK' -ForegroundColor Green
    Write-Host ''


    # ----------------------------------------------------------
    #  STEP 2 -- Load Data and Build Account List
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Loading Account Data' -Step '2 of 5'

    $accountsCsvPath = Get-OMMigrateCsvPath -BaseName 'migration_accounts.csv'
    $allAccounts     = Import-Csv -Path $accountsCsvPath -Encoding UTF8

    # -- Build sanitization map --------------------------------
    if ($Global:OMMigrate.Sanitize) {
        Initialize-SanitizeMap -Accounts @($allAccounts)
        Write-OMMigrateLog -Message '[SANITIZE] Sanitization active -- output masked.' `
                           -Level INFO
    }

    # All IMAP accounts are eligible for artifact migration.
    # This includes:
    #   IMAP-CONVERTED -- was POP3, migrated via Scripts 01/02, has backup PST
    #   IMAP-ALREADY   -- was always IMAP, has OST-exported backup PST
    #   COMPLETE       -- folder migration done, artifacts may not be migrated yet
    # EXCHANGE-SKIP accounts are excluded -- Exchange manages its own artifacts.
    $accountsForArtifacts = @($allAccounts | Where-Object {
        $_.ProviderTag -ne 'EXCHANGE-SKIP' -and
        $_.ProviderTag -ne 'PST-ARCHIVE'   -and
        $_.AccountType -eq 'IMAP'
    })

    Write-Host "  Total accounts     : $($allAccounts.Count)" -ForegroundColor Gray
    Write-Host "  Eligible for artifacts: $($accountsForArtifacts.Count)" -ForegroundColor Cyan
    Write-Host "  (All IMAP accounts -- Exchange and PST-ARCHIVE excluded)" `
               -ForegroundColor DarkGray
    Write-Host ''

    # Checkpoint check
    $checkpoint  = Read-OMMigrateCheckpoint -Step 4
    $alreadyDone = @()
    if ($checkpoint.HasCheckpoint) {
        $alreadyDone = $checkpoint.CompletedAccounts
        Write-Host "  Checkpoint: $($alreadyDone.Count) account(s) already have artifacts migrated." `
                   -ForegroundColor Yellow

        $resumeConfirmed = Confirm-Action `
            -Message    'Resume from checkpoint? (N to re-process all accounts)' `
            -DefaultYes $true
        if (-not $resumeConfirmed) { $alreadyDone = @() }
    }

    $accountsToProcess = @($accountsForArtifacts | Where-Object {
        $alreadyDone -notcontains $_.EmailAddress
    })

    Write-Host "  To process this run: $($accountsToProcess.Count)" -ForegroundColor Cyan
    Write-Host ''


    # ----------------------------------------------------------
    #  ACCOUNT PICKER -- WinForms checkbox list
    #  Same pattern as Script 03. Lets the operator choose which
    #  accounts to process. Skipped when only one account eligible.
    # ----------------------------------------------------------
    $showPicker = ($accountsToProcess.Count -gt 1 -and -not $Script:IsWhatIf)

    if ($showPicker) {
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            Add-Type -AssemblyName System.Drawing       -ErrorAction Stop
        }
        catch {
            Write-OMMigrateLog -Message "Could not load WinForms for account picker: $_" -Level WARN
            Write-Host '  WARNING: Could not open account picker (WinForms unavailable).' `
                       -ForegroundColor Yellow
            Write-Host '  Processing all eligible accounts.' -ForegroundColor Yellow
        }

        # Build picker items with backup PST size info
        $pickerItems = @($accountsToProcess | ForEach-Object {
            $acct       = $_
            $safeEmail  = Get-SafeFileName -InputString $acct.EmailAddress
            $bkpPath    = Join-Path $Global:OMMigrate.BackupPath "$safeEmail.pst"
            $bkpSize    = if (Test-Path $bkpPath) {
                $bytes = (Get-Item $bkpPath).Length
                if     ($bytes -gt 1GB) { '{0:N1} GB' -f ($bytes / 1GB) }
                elseif ($bytes -gt 1MB) { '{0:N1} MB' -f ($bytes / 1MB) }
                else                    { '{0:N0} KB' -f ($bytes / 1KB) }
            } else { 'No backup' }

            [PSCustomObject]@{
                EmailAddress = $acct.EmailAddress
                DisplayName  = $acct.DisplayName
                ProviderTag  = $acct.ProviderTag
                BackupSize   = $bkpSize
            }
        } | Sort-Object EmailAddress)

        Write-Host ''
        Write-Host '  A selection window will open -- choose accounts to migrate artifacts for.' `
                   -ForegroundColor Cyan
        Write-Host '  Use Select All to process all accounts.' -ForegroundColor Cyan
        Write-Host '  Click Cancel to exit safely without making changes.' -ForegroundColor Cyan
        Write-Host ''

        $form04               = New-Object System.Windows.Forms.Form
        $form04.Text          = 'OMMigrate -- Select Accounts for Artifact Migration'
        $form04.Size          = New-Object System.Drawing.Size(640, 440)
        $form04.StartPosition = 'CenterScreen'
        $form04.FormBorderStyle = 'FixedDialog'
        $form04.MaximizeBox   = $false
        $form04.MinimizeBox   = $false
        $form04.TopMost       = $true

        $label04           = New-Object System.Windows.Forms.Label
        $label04.Location  = New-Object System.Drawing.Point(12, 12)
        $label04.Size      = New-Object System.Drawing.Size(600, 44)
        $label04.Text      = 'Check each account to migrate Calendar, Contacts, Tasks, Notes and Journal items. Only items not already in the IMAP store will be copied.'
        $label04.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
        $form04.Controls.Add($label04)

        $listBox04              = New-Object System.Windows.Forms.CheckedListBox
        $listBox04.Location     = New-Object System.Drawing.Point(12, 64)
        $listBox04.Size         = New-Object System.Drawing.Size(600, 295)
        $listBox04.Font         = New-Object System.Drawing.Font('Segoe UI', 9)
        $listBox04.CheckOnClick = $true

        foreach ($item in $pickerItems) {
            $displayLine = "$($item.EmailAddress)  [$($item.ProviderTag)]  -- Backup: $($item.BackupSize)"
            $idx = $listBox04.Items.Add($displayLine)
            $listBox04.SetItemChecked($idx, $false)
        }
        $form04.Controls.Add($listBox04)

        $btnSelectAll04          = New-Object System.Windows.Forms.Button
        $btnSelectAll04.Location = New-Object System.Drawing.Point(12, 370)
        $btnSelectAll04.Size     = New-Object System.Drawing.Size(90, 28)
        $btnSelectAll04.Text     = 'Select All'
        $btnSelectAll04.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
        $btnSelectAll04.Add_Click({
            for ($i = 0; $i -lt $listBox04.Items.Count; $i++) {
                $listBox04.SetItemChecked($i, $true)
            }
        })
        $form04.Controls.Add($btnSelectAll04)

        $btnClearAll04          = New-Object System.Windows.Forms.Button
        $btnClearAll04.Location = New-Object System.Drawing.Point(110, 370)
        $btnClearAll04.Size     = New-Object System.Drawing.Size(90, 28)
        $btnClearAll04.Text     = 'Clear All'
        $btnClearAll04.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
        $btnClearAll04.Add_Click({
            for ($i = 0; $i -lt $listBox04.Items.Count; $i++) {
                $listBox04.SetItemChecked($i, $false)
            }
        })
        $form04.Controls.Add($btnClearAll04)

        $btnOK04              = New-Object System.Windows.Forms.Button
        $btnOK04.Location     = New-Object System.Drawing.Point(438, 370)
        $btnOK04.Size         = New-Object System.Drawing.Size(80, 28)
        $btnOK04.Text         = 'OK'
        $btnOK04.Font         = New-Object System.Drawing.Font('Segoe UI', 9)
        $btnOK04.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form04.AcceptButton  = $btnOK04
        $form04.Controls.Add($btnOK04)

        $btnCancel04               = New-Object System.Windows.Forms.Button
        $btnCancel04.Location      = New-Object System.Drawing.Point(536, 370)
        $btnCancel04.Size          = New-Object System.Drawing.Size(80, 28)
        $btnCancel04.Text          = 'Cancel'
        $btnCancel04.Font          = New-Object System.Drawing.Font('Segoe UI', 9)
        $btnCancel04.DialogResult  = [System.Windows.Forms.DialogResult]::Cancel
        $form04.CancelButton       = $btnCancel04
        $form04.Controls.Add($btnCancel04)

        $result04 = $form04.ShowDialog()

        # Capture checked indices BEFORE disposing
        # Capture checked item strings BEFORE disposing -- bypass index
        # arithmetic entirely. CheckedItems returns the display strings of
        # checked items, which contain the email address as the first token.
        # This avoids PS 5.1 CheckedIndexCollection enumeration edge cases.
        $checkedItemStrings = [System.Collections.Generic.List[string]]::new()
        foreach ($checkedItem in $listBox04.CheckedItems) {
            $checkedItemStrings.Add($checkedItem.ToString())
        }
        $form04.Dispose()

        if ($result04 -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Host ''
            Write-Host '  No accounts selected -- exiting safely. No changes made.' `
                       -ForegroundColor Yellow
            Write-OMMigrateLog -Message 'Operator cancelled at account picker.' -Level INFO
            exit 0
        }

        if (-not $checkedItemStrings -or $checkedItemStrings.Count -eq 0) {
            Write-Host ''
            Write-Host '  No accounts checked -- exiting safely. No changes made.' `
                       -ForegroundColor Yellow
            Write-OMMigrateLog -Message 'Operator clicked OK with no accounts checked -- exiting.' `
                               -Level INFO
            exit 0
        }

        # Match checked display strings back to email addresses.
        # Display format: "$email  [$tag]  -- Backup: $size"
        # Email is always the first whitespace-delimited token.
        $selectedEmails = @($checkedItemStrings | ForEach-Object {
            ($_ -split '\s+')[0]
        })
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
        Write-Host ''
        Write-Host '  [PREVIEW] Account picker skipped in WhatIf mode.' -ForegroundColor DarkGray
        Write-Host "  [PREVIEW] Would prompt selection from $($accountsToProcess.Count) eligible accounts." `
                   -ForegroundColor DarkGray
    }


    # ----------------------------------------------------------
    #  STEP 3 -- Pre-Flight Confirmation
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Pre-Flight Confirmation' -Step '3 of 5'

    Show-PreflightWarning `
        -ScriptDescription (
            "This script will copy Calendar, Contacts, Tasks, Notes and Journal " +
            "items from each account's backup PST into the live IMAP account. " +
            "Only items not already present in the IMAP store will be copied. " +
            "Backup PSTs are read-only -- your backup data is never modified. " +
            "Items are copied to the IMAP server so they sync to all your devices."
        ) `
        -Prerequisites @(
            "Script 02 completed -- all accounts converted to IMAP",
            "Outlook is fully closed (this script will launch it via COM)",
            "Backup PSTs exist in the Backups folder for accounts to process"
        ) `
        -DeclineMessage @(
            "You chose not to proceed at this time. No artifacts have been migrated.",
            "Your IMAP accounts from Script 02 are intact and working.",
            "When you are ready, re-run this script:",
            "  .\Scripts\OMMigrate-04-Artifacts.ps1"
        )

    # Auto-close Outlook if running -- scripts launch their own COM session
    # and require exclusive access. Moved here, AFTER the operator's Y/N
    # confirmation above (Show-PreflightWarning exits the script entirely
    # on a No/decline, so anything below this point only runs after an
    # explicit Yes) -- avoids closing Outlook out from under the admin if
    # they still have unsaved/in-progress work open at the moment the
    # script was launched.
    [void](Close-OutlookIfRunning -Reason 'before pre-flight')

    # Initialize progress tracker
    $pendingEmails = @($accountsToProcess | ForEach-Object { $_.EmailAddress })
    Update-OMMigrateProgress -SetPending $pendingEmails


    # ----------------------------------------------------------
    #  STEP 4 -- Launch COM and Process Accounts
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Starting Outlook and Migrating Artifacts' -Step '4 of 5'

    Write-Host '  Launching Outlook COM session...' -ForegroundColor Cyan

    # Read selected profile from Settings.json
    $s04ProfileName = ''
    try {
        if ($Global:OMMigrate.Settings -and
            $Global:OMMigrate.Settings.PSObject.Properties['OutlookProfile'] -and
            $Global:OMMigrate.Settings.OutlookProfile.SelectedProfile) {
            $s04ProfileName = $Global:OMMigrate.Settings.OutlookProfile.SelectedProfile
        }
    }
    catch { }

    Write-OMMigrateLog -Message "Using Outlook profile: '$s04ProfileName'" -Level INFO
    $outlook = Connect-OutlookCOM -ProfileName $s04ProfileName

    if (-not $outlook -and -not $Script:IsWhatIf) {
        Write-OMMigrateLog -Message 'Failed to start Outlook COM session.' -Level ERROR
        $Script:FinalStatus = 'FAILED'
        exit 1
    }
    if ($outlook) { $Script:COMSessionOpen = $true }
    Write-Host '  Outlook COM session started.' -ForegroundColor Green

    # Suspend Send/Receive to prevent sync interference during migration
    Suspend-OutlookSendReceive | Out-Null

    $namespace = Get-OutlookNamespace

    Write-Host ''

    # Process each account
    foreach ($account in $accountsToProcess) {
        $email = $account.EmailAddress

        Write-Host ''
        Write-Host ('  ' + ('=' * 54)) -ForegroundColor DarkCyan
        Write-Host "  Account : $(Invoke-OMMigrateSanitize -Text $email)" -ForegroundColor White
        Write-Host "  Tag     : $($account.ProviderTag)" -ForegroundColor DarkGray
        Write-Host ''

        # Per-account Y/N prompt unless -Force
        $proceed = $true
        if (-not $Force) {
            $proceed = Confirm-Action `
                -Message      "Migrate artifacts for: $email ?" `
                -AccountEmail $email `
                -DefaultYes   $true
        }

        if (-not $proceed) {
            Write-OMMigrateLog -Message "Artifact migration skipped by operator: $email" -Level INFO

            $skipResult = [PSCustomObject]@{
                EmailAddress  = $email
                MigrationOutcome = 'SKIPPED'
                MigrationDetail  = 'Skipped by operator'
                ProviderTag   = $account.ProviderTag
            }
            [void]$Script:AccountResults.Add($skipResult)
            Update-OMMigrateProgress -MarkComplete $email
            continue
        }

        Update-OMMigrateProgress -SetCurrent $email

        # Run artifact migration for this account
        $migrationResult = Invoke-AccountArtifactMigration `
            -Account   $account `
            -Namespace $namespace

        # Attach outcome fields for the report (uses Migration report format)
        $accountResult = $account.PSObject.Copy()
        $accountResult | Add-Member -NotePropertyName 'MigrationOutcome' `
                                    -NotePropertyValue $migrationResult.Outcome -Force
        $accountResult | Add-Member -NotePropertyName 'MigrationDetail' `
                                    -NotePropertyValue $migrationResult.Detail -Force
        [void]$Script:AccountResults.Add($accountResult)

        # Console summary
        $statusIcon = switch ($migrationResult.Outcome) {
            'SUCCESS' { 'OK'   }
            'WARNING' { 'WARN' }
            'SKIPPED' { 'SKIP' }
            default   { 'FAIL' }
        }

        Show-AccountStatus `
            -Email  $email `
            -Tag    $account.ProviderTag `
            -Action $migrationResult.Detail `
            -Status $statusIcon

        if ($migrationResult.Outcome -eq 'FAILED' -and
            $Script:FinalStatus -eq 'SUCCESS') {
            $Script:FinalStatus = 'WARNING'
        }

        Save-OMMigrateCheckpoint -ExitType 'NORMAL' -Reason 'Account artifact migration complete'
        Update-OMMigrateProgress -MarkComplete $email
    }


    # ----------------------------------------------------------
    #  STEP 5 -- Generate Report and Final Manifest
    # ----------------------------------------------------------
    Show-SectionHeader -Title 'Generating Final Report and Manifest' -Step '5 of 5'

    # Determine final status
    if ($Script:TotalItemsFailed -gt 0 -and
        ($Script:TotalCalendarCopied + $Script:TotalContactsCopied +
         $Script:TotalTasksCopied + $Script:TotalNotesCopied +
         $Script:TotalJournalCopied) -eq 0) {
        $Script:FinalStatus = 'FAILED'
    }
    elseif ($Script:TotalItemsFailed -gt 0) {
        $Script:FinalStatus = 'WARNING'
    }

    # Generate report
    $allResultsForReport = @($Script:AccountResults | Sort-Object EmailAddress)
    if ($allResultsForReport.Count -gt 0) {
        $Script:ReportFile = New-MigrationReport `
            -Accounts       $allResultsForReport `
            -Subtitle       'Personal Artifacts Migration Results -- Script 04' `
            -ReportName     'Artifacts' `
            -ArtifactTotals ([PSCustomObject]@{
                Calendar          = $Script:TotalCalendarCopied
                Contacts          = $Script:TotalContactsCopied
                ContactSubfolders = $Script:TotalContactSubfoldersCopied
                Tasks             = $Script:TotalTasksCopied
                Notes             = $Script:TotalNotesCopied
                Journal           = $Script:TotalJournalCopied
            })
        Write-Host "  Artifacts Report: $Script:ReportFile" -ForegroundColor Green
    }

    # Write Step 04 manifest
    $totalCopied = $Script:TotalCalendarCopied + $Script:TotalContactsCopied +
                   $Script:TotalContactSubfoldersCopied + $Script:TotalTasksCopied +
                   $Script:TotalNotesCopied + $Script:TotalJournalCopied

    Write-StepManifest -Step 4 -Status $Script:FinalStatus -Data @{
        CalendarCopied            = $Script:TotalCalendarCopied
        ContactsCopied            = $Script:TotalContactsCopied
        ContactSubfoldersCopied   = $Script:TotalContactSubfoldersCopied
        TasksCopied               = $Script:TotalTasksCopied
        NotesCopied               = $Script:TotalNotesCopied
        JournalCopied             = $Script:TotalJournalCopied
        TotalCopied               = $totalCopied
        TotalFailed      = $Script:TotalItemsFailed
        AccountsProcessed = $Script:AccountResults.Count
        ReportFile       = $Script:ReportFile
    }

    Write-Host ''

    # Final console summary
    Write-Host ('=' * 60) -ForegroundColor DarkCyan
    Write-Host '  ARTIFACT MIGRATION COMPLETE' -ForegroundColor White
    Write-Host ('=' * 60) -ForegroundColor DarkCyan
    Write-Host ''
    Write-Host "  Calendar copied         : $Script:TotalCalendarCopied"          -ForegroundColor Gray
    Write-Host "  Contacts copied         : $Script:TotalContactsCopied"          -ForegroundColor Gray
    Write-Host "  Contact Subfolders copied: $Script:TotalContactSubfoldersCopied" -ForegroundColor Gray
    Write-Host "  Tasks copied            : $Script:TotalTasksCopied"             -ForegroundColor Gray
    Write-Host "  Notes copied            : $Script:TotalNotesCopied"             -ForegroundColor Gray
    Write-Host "  Journal copied          : $Script:TotalJournalCopied"           -ForegroundColor Gray
    Write-Host "  Total copied     : $totalCopied" -ForegroundColor $(
        if ($totalCopied -gt 0) { 'Green' } else { 'Gray' }
    )
    if ($Script:TotalItemsFailed -gt 0) {
        Write-Host "  Failed           : $Script:TotalItemsFailed  (see log for details)" `
                   -ForegroundColor Yellow
    }
    Write-Host ''

    if ($Script:FinalStatus -eq 'SUCCESS') {
        Write-Host '  All done! Artifact migration complete.' -ForegroundColor Green
        Write-Host ''
        Write-Host '  RECOMMENDED NEXT STEPS:' -ForegroundColor Cyan
        Write-Host '  1. Open Outlook and verify Calendar, Contacts and Tasks are present.' `
                   -ForegroundColor Gray
        Write-Host '  2. Check your phone mail client -- artifacts should sync within minutes.' `
                   -ForegroundColor Gray
        Write-Host '  3. Verify no duplicate entries exist in Calendar and Contacts.' `
                   -ForegroundColor Gray
    }
    elseif ($Script:FinalStatus -eq 'WARNING') {
        Write-Host '  Artifact migration completed with warnings.' -ForegroundColor Yellow
        Write-Host '  Review the Artifacts Report and log for details.' -ForegroundColor Yellow
        Write-Host '  Re-running is safe -- only missing items will be copied.' -ForegroundColor Yellow
    }
    else {
        Write-Host '  Artifact migration encountered failures.' -ForegroundColor Red
        Write-Host '  Review the log for details. Re-running is safe.' -ForegroundColor Red
    }

    Write-Host ''

}
catch {
    Write-OMMigrateLog -Message "FATAL ERROR in Script 04: $_" -Level ERROR
    Write-OMMigrateLog -Message "Stack trace: $($_.ScriptStackTrace)" -Level DEBUG
    $Script:FinalStatus = 'FAILED'

    Write-Host ''
    Write-Host '  FATAL ERROR:' -ForegroundColor Red
    Write-Host "  $_" -ForegroundColor Red
    Write-Host ''
    Write-Host '  Your IMAP account settings are intact.' -ForegroundColor Yellow
    Write-Host '  Artifact migration can be re-run safely at any time.' -ForegroundColor Yellow
    Write-Host "  Log: $($Global:OMMigrate.RunLogFile)" -ForegroundColor Gray
    Write-Host ''
}
finally {
    # -- Always runs --------------------------------------------
    if ($Script:COMSessionOpen) {
        try { Resume-OutlookSendReceive | Out-Null } catch { }
        try { Release-OutlookCOM } catch { }
        $Script:COMSessionOpen = $false
    }

    # Save checkpoint if work started
    if ($Script:COMSessionOpen -or $Script:AccountResults.Count -gt 0) {
        Save-OMMigrateCheckpoint -ExitType 'NORMAL' -Reason 'Script 04 session ending'
    }

    # Clean up checkpoint after fully successful session
    if ($Script:FinalStatus -eq 'SUCCESS') {
        $checkpointPath = Join-Path $Global:OMMigrate.ManifestPath 'Step04_Checkpoint.json'
        if (Test-Path $checkpointPath) {
            try {
                Remove-Item $checkpointPath -Force -ErrorAction Stop
                Write-OMMigrateLog -Message 'Step04_Checkpoint.json removed after successful session.' `
                                   -Level DEBUG
            }
            catch {
                Write-OMMigrateLog -Message "Could not remove Step04_Checkpoint.json: $_" -Level INFO
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

    Wait-UserKeypress
}

# ***** END OF FILE *****
