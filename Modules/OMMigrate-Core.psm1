#Requires -Version 5.1
<#
.SYNOPSIS
    OMMigrate-Core.psm1 -- Core Shared Module for OutlookMailMigrator

.DESCRIPTION
    Provides the foundational shared functions used by all OMMigrate scripts:

        - Logging engine (run log + cumulative audit log)
        - Console output formatting with color-coded severity levels
        - Settings management (load/save OMMigrate_Settings.json)
        - Manifest read/write (inter-script gate control)
        - Input validation utilities
        - Environment pre-flight checks (OS, PowerShell version, Outlook install)
        - User prompt / confirmation helpers (Y/N gates)
        - HTML report scaffolding (header, footer, section builders)
        - Version and credit banner

    This module must be imported by all OMMigrate scripts before any other
    OMMigrate module is loaded. It initializes the global session context
    ($Global:OMMigrate) that all modules share.

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
        Outlook     : Classic Outlook 2016 / 2019 / 2021
                      (NOT New Outlook / web-based Outlook)

    Run Context:
        Must be run as the Windows user who owns the Outlook profile.
        Do NOT run as Administrator unless the Outlook profile itself
        was created under the Administrator account.

    Compatibility Notes:
        - Outlook COM automation requires classic Outlook to be installed.
        - New Outlook (Microsoft Store version) does NOT expose a COM API
          and is not supported by this tool.
        - Office 365 / Microsoft 365 subscription Outlook is supported
          provided it is the classic desktop client (not web-based).

.LINK
    https://github.com/[repository-pending]
#>

Set-StrictMode -Version Latest

# ============================================================
#  MODULE VERSION & IDENTITY
# ============================================================

# -- Read version from version.txt (single source of truth) -------------------
# version.txt lives in the project root -- one level above the Modules folder.
# If the file is not found (e.g. running from an unusual path), fall back to
# the inline string so the module always loads without error.
$Script:OMMigrateVersion = '1.5.2'   # fallback -- overwritten below if version.txt found
try {
    $versionFilePath = Join-Path $PSScriptRoot '..\version.txt'
    if (Test-Path $versionFilePath) {
        $readVersion = (Get-Content $versionFilePath -Raw -ErrorAction Stop).Trim()
        if ($readVersion -match '^\d+\.\d+\.\d+$') {
            $Script:OMMigrateVersion = $readVersion
        }
    }
}
catch { }
# -----------------------------------------------------------------------------

$Script:OMMigrateProduct  = 'OutlookMailMigrator'
$Script:OMMigrateShort    = 'OMMigrate'
$Script:OMMigrateAuthor   = 'Kirk Shallcross - Shallcross Consulting'
$Script:OMMigrateAI       = 'Anthropic Claude AI'
$Script:OMMigrateInception = 'May 2026'


# ============================================================
#  GLOBAL SESSION CONTEXT
#  Shared across all OMMigrate modules and scripts in the
#  same PowerShell session. Initialized by Initialize-OMMigrate.
# ============================================================

if (-not (Get-Variable -Name OMMigrate -Scope Global -ErrorAction SilentlyContinue)) {
    $Global:OMMigrate = [ordered]@{

        # Tool identity
        Version         = $Script:OMMigrateVersion
        Product         = $Script:OMMigrateProduct
        ShortName       = $Script:OMMigrateShort

        # Session runtime
        SessionID       = $null          # GUID assigned at Initialize-OMMigrate
        StartTime       = $null          # DateTime of session start
        ScriptName      = $null          # Name of the calling script
        RunUser         = $null          # Windows username running the tool
        MachineName     = $null          # Computer name
        WhatIf          = $false         # Dry-run mode flag

        # Paths (set by Initialize-OMMigrate from settings + parameters)
        BasePath        = $null          # Root working directory
        ConfigPath      = $null          # .\Config\
        LogPath         = $null          # .\Logs\
        ReportPath      = $null          # .\Reports\
        BackupPath      = $null          # .\Backups\
        ManifestPath    = $null          # .\Manifests\

        # Active log file paths for this session
        RunLogFile      = $null          # .\Logs\OMMigrate_YYYYMMDD_HHMMSS.log
        AuditLogFile    = $null          # .\Logs\OMMigrate_Audit.log

        # Settings loaded from OMMigrate_Settings.json
        Settings        = $null

        # Logging
        LogLevel        = 'INFO'         # DEBUG | INFO | WARN | ERROR
        LogLevelMap     = @{
            DEBUG = 0
            INFO  = 1
            WARN  = 2
            ERROR = 3
        }

        # Counters for summary reporting
        Counters        = [ordered]@{
            LogEntries   = 0
            Warnings     = 0
            Errors       = 0
            Prompts      = 0
            Confirmed    = 0
            Skipped      = 0
        }

        # Sanitization -- active when -Sanitize switch is passed
        # Masks emails, servers, display names, paths, and custom
        # folder names in all console output and log file writes.
        Sanitize        = $false         # $true when -Sanitize is active
        SanitizeMap     = $null          # ordered hashtable: real -> alias
    }
}


# ============================================================
#  REGION: INITIALIZATION
# ============================================================

function Initialize-OMMigrate {
    <#
    .SYNOPSIS
        Initializes the OMMigrate session context.

    .DESCRIPTION
        Must be called at the start of every OMMigrate script before any
        other OMMigrate function is used. Sets up paths, logging, session
        identity, and loads settings from OMMigrate_Settings.json.

        Safe to call multiple times -- re-entrant calls update the
        ScriptName and refresh the session timestamp without resetting
        counters or log files already open.

    .PARAMETER ScriptName
        Name of the calling script (e.g. 'OMMigrate-00-Discover').
        Used in log headers and audit entries.

    .PARAMETER BasePath
        Override the default working directory.
        Default: $env:USERPROFILE\Documents\OutlookMigration

    .PARAMETER LogLevel
        Logging verbosity: DEBUG | INFO | WARN | ERROR
        Default: INFO (or value from OMMigrate_Settings.json if present)

    .PARAMETER WhatIf
        When $true, no files are written, no Outlook changes are made.
        All actions are simulated and logged with [WHATIF] prefix.

    .EXAMPLE
        Initialize-OMMigrate -ScriptName 'OMMigrate-00-Discover'

    .EXAMPLE
        Initialize-OMMigrate -ScriptName 'OMMigrate-01-Backup' `
                             -BasePath 'D:\Migration' `
                             -LogLevel 'DEBUG' `
                             -WhatIf $true
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptName,

        [Parameter(Mandatory = $false)]
        [string]$BasePath = '',

        [Parameter(Mandatory = $false)]
        [ValidateSet('DEBUG','INFO','WARN','ERROR')]
        [string]$LogLevel = 'INFO',

        [Parameter(Mandatory = $false)]
        [bool]$IsWhatIf = $false,

        [Parameter(Mandatory = $false)]
        [bool]$Sanitize = $false
    )

    # -- Reset counters for fresh run -------------------------
    # Always reset counters at start of each script run to prevent
    # accumulation across multiple runs in the same PowerShell session
    $Global:OMMigrate.Counters.LogEntries = 0
    $Global:OMMigrate.Counters.Warnings   = 0
    $Global:OMMigrate.Counters.Errors     = 0
    $Global:OMMigrate.Counters.Prompts    = 0
    $Global:OMMigrate.Counters.Confirmed  = 0
    $Global:OMMigrate.SessionCompletedNormally = $false

    # -- Session identity --------------------------------------
    $Global:OMMigrate.SessionID   = [System.Guid]::NewGuid().ToString()
    $Global:OMMigrate.StartTime   = Get-Date
    $Global:OMMigrate.ScriptName  = $ScriptName
    $Global:OMMigrate.RunUser     = "$env:USERDOMAIN\$env:USERNAME"
    $Global:OMMigrate.MachineName = $env:COMPUTERNAME
    $Global:OMMigrate.WhatIf      = $IsWhatIf
    $Global:OMMigrate.LogLevel    = $LogLevel

    # -- Sanitize mode -----------------------------------------
    # Reset map each run so a fresh session never inherits stale aliases
    $Global:OMMigrate.Sanitize    = $Sanitize
    $Global:OMMigrate.SanitizeMap = [ordered]@{}

    # -- Resolve base path -------------------------------------
    if ([string]::IsNullOrWhiteSpace($BasePath)) {
        $BasePath = Join-Path $env:USERPROFILE 'Documents\OutlookMigration'
    }
    $Global:OMMigrate.BasePath     = $BasePath
    $Global:OMMigrate.ConfigPath   = Join-Path $BasePath 'Config'
    $Global:OMMigrate.LogPath      = Join-Path $BasePath 'Logs'
    $Global:OMMigrate.ReportPath   = Join-Path $BasePath 'Reports'
    $Global:OMMigrate.BackupPath   = Join-Path $BasePath 'Backups'
    $Global:OMMigrate.ManifestPath = Join-Path $BasePath 'Manifests'

    # -- Create directory structure if needed ------------------
    $dirs = @(
        $Global:OMMigrate.BasePath,
        $Global:OMMigrate.ConfigPath,
        $Global:OMMigrate.LogPath,
        $Global:OMMigrate.ReportPath,
        $Global:OMMigrate.BackupPath,
        $Global:OMMigrate.ManifestPath
    )
    foreach ($dir in $dirs) {
        if (-not (Test-Path $dir)) {
            if (-not $IsWhatIf) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
        }
    }

    # -- Initialize log file paths -----------------------------
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $Global:OMMigrate.RunLogFile  = Join-Path $Global:OMMigrate.LogPath `
                                              "OMMigrate_${timestamp}.log"
    $Global:OMMigrate.AuditLogFile = Join-Path $Global:OMMigrate.LogPath `
                                              'OMMigrate_Audit.log'

    # -- Load settings -----------------------------------------
    $settingsFile = Join-Path $Global:OMMigrate.ConfigPath 'OMMigrate_Settings.json'
    if (Test-Path $settingsFile) {
        try {
            $Global:OMMigrate.Settings = Get-Content $settingsFile -Raw |
                                         ConvertFrom-Json
            # Settings file can override log level if not explicitly passed
            if ($Global:OMMigrate.Settings.Logging.LogLevel -and
                $LogLevel -eq 'INFO') {
                $Global:OMMigrate.LogLevel = $Global:OMMigrate.Settings.Logging.LogLevel
            }
        }
        catch {
            # Settings file corrupt -- use defaults, will warn after logging starts
            $Global:OMMigrate.Settings = Get-DefaultSettings
        }
    }
    else {
        $Global:OMMigrate.Settings = Get-DefaultSettings
        # Write default settings file for future runs
        if (-not $IsWhatIf) {
            $Global:OMMigrate.Settings |
                ConvertTo-Json -Depth 10 |
                Set-Content -Path $settingsFile -Encoding UTF8
        }
    }

    # -- Write session header to run log -----------------------
    Write-RunLogHeader

    # -- Write audit session open entry ------------------------
    Write-AuditEntry -Action 'SESSION_START' `
                     -Detail "Script=$ScriptName | User=$($Global:OMMigrate.RunUser) | Machine=$($Global:OMMigrate.MachineName) | WhatIf=$IsWhatIf"

    # -- Pre-register machine and operator for banner sanitization -
    # These are known at this point and must be registered before
    # Show-Banner displays them, since the full account map isn't
    # built until after Initialize-OMMigrate returns.
    if ($Sanitize) {
        if ($env:COMPUTERNAME) {
            Register-SanitizeTerms -Terms @($env:COMPUTERNAME) -Category 'Path'
        }
        if ($env:USERDOMAIN -and $env:USERNAME) {
            Register-SanitizeTerms -Terms @("$env:USERDOMAIN\$env:USERNAME") -Category 'Path'
            Register-SanitizeTerms -Terms @($env:USERNAME) -Category 'Path'
        }
    }

    # -- Display banner ----------------------------------------
    Show-Banner
}


# ============================================================
#  REGION: SETTINGS
# ============================================================

function Get-DefaultSettings {
    <#
    .SYNOPSIS
        Returns the default settings object used when no settings file exists.

    .DESCRIPTION
        Called internally by Initialize-OMMigrate when OMMigrate_Settings.json
        is not found. The returned object is also written to disk as the
        default settings file for the user to customize.

    .OUTPUTS
        PSCustomObject -- default settings structure.
    #>
    return [PSCustomObject]@{
        ToolVersion   = $Script:OMMigrateVersion
        DefaultPaths  = [PSCustomObject]@{
            BasePath  = "$env:USERPROFILE\Documents\OutlookMigration"
        }
        BackupVerification = [PSCustomObject]@{
            CheckFileSize   = $true
            MinimumSizeMB   = 0
            CheckCanOpen    = $true
        }
        Migration = [PSCustomObject]@{
            DefaultImapPort          = 993
            DefaultSmtpPort          = 587
            DefaultSSL               = $true
            PromptBeforeEachAccount  = $true
            WhatIfByDefault          = $false
        }
        Reporting = [PSCustomObject]@{
            GenerateHTML         = $true
            GenerateCSV          = $true
            OpenReportAfterRun   = $true
            # Optional: full path to a preferred text editor for opening logs,
            # CSV files, and plain-text output files.
            # Leave empty ("") to use the Windows default file association --
            # this is the recommended default for broad compatibility.
            #
            # Examples:
            #   "C:\Program Files\Sublime Text\subl.exe"
            #   "C:\Program Files\Microsoft VS Code\Code.exe"
            #   "C:\Program Files\Notepad++\notepad++.exe"
            #
            # If empty, OMMigrate will still auto-detect Sublime Text if
            # installed and prefer it automatically.
            PreferredEditor      = ''
        }
        Logging = [PSCustomObject]@{
            LogLevel              = 'INFO'
            RetainRunLogsForDays  = 90
            AuditLogMaxSizeMB     = 50
        }
        OutlookProfile = [PSCustomObject]@{
            # Name of the Outlook profile selected by the operator in Script 00.
            # Written by Script 00 at startup and read by all subsequent scripts
            # to ensure the correct profile is used for every COM session.
            # Empty string means no profile has been selected yet.
            SelectedProfile       = ''
        }
        # EXTRACTED 2026-07-09 (Administrator direction) from OMMigrate-Core_WIP.psm1,
        # a 6-day-old (2026-07-03) work-in-progress file -- this is the
        # TargetStoreName hardcode fix work paused mid-implementation that
        # session. Extracted here as a jumpstart on resuming that work, NOT
        # a completed/finished feature -- see master task list Tier
        # (TargetStoreName hardcode) for full context. Do not treat this as
        # done; it is the settings-schema half of a multi-file change that
        # also touches OMMigrate-Outlook.psm1 and OMMigrate-00-Discover.ps1.
        RulesEngine = [PSCustomObject]@{
            # Archive-store-to-account mapping selected by the operator in
            # Script 00's TargetStoreName picker (shown every run, after
            # standalone PST discovery). Each entry maps one attached PST
            # (by its live-detected store DisplayName) to one or more email
            # account DisplayNames whose rules should target that PST.
            # No PST or account name is ever hardcoded -- both sides of
            # every entry come from what Script 00 actually finds on this
            # machine/profile at run time. Empty array means the operator
            # has not made a selection yet (Export-RulesToCSV falls back
            # to defaulting TargetStoreName = RuleStoreName in that case).
            #
            # Example shape (one PST, multiple accounts mapped to it):
            #   [
            #     {
            #       "TargetStoreName": "MyArchivePST",
            #       "RuleStoreNames": ["user@example.com", "admin2@example.com"]
            #     }
            #   ]
            ArchiveStoreMappings = @()

            # Added 2026-07-10, Administrator direction. Live-tested bug found: Script
            # 01's Archive pre-build (multi-archive support feature, same
            # session) unconditionally closed/detached every archive PST it
            # opened, including 'OMMigrate Local Archive' -- which had been
            # manually attached by Administrator in Outlook, not mounted by this run,
            # and should never be auto-detached. Confirmed live: after a
            # normal Script 01 run, the archive was found missing from
            # Outlook's folder pane entirely.
            #
            # This list names archive PSTs that must NEVER be auto-detached
            # by any OMMigrate script, regardless of whether that script's
            # own code did the mounting this run -- these are permanent,
            # always-attached infrastructure the operator manages manually
            # (e.g. 'OMMigrate Local Archive', a personal auto-archive store
            # like 'Personal-Archive'). Matched by exact store DisplayName, checked
            # before any Close-PSTFile call on an archive-type store. Operator-
            # editable here rather than hardcoded in script logic, so adding
            # or removing a master archive is a settings edit, not a code
            # change -- same reasoning as ArchiveStoreMappings above.
            #
            # Empty array (default) means no protected master archives are
            # configured -- existing attach/detach behavior (only detach what
            # this run's own code mounted, via Test-PSTAlreadyMounted) is
            # unaffected until the operator adds names here.
            #
            # Example: ["OMMigrate Local Archive", "Personal-Archive"]
            MasterArchiveNames = @()

            # Maximum seconds Release-OutlookCOM waits for Outlook to exit
            # cleanly after Quit() before force-killing the process. Outlook
            # must flush all pending PST write buffers before it can safely
            # be terminated -- a premature force-kill can silently discard
            # unflushed PST changes. Larger rule sets (hundreds of rules)
            # take longer to serialize on exit and may need a higher value;
            # environments with few rules can safely lower this for a
            # faster finish at the end of every script. Default: 20
            # (tuned for a large, multi-hundred-rule environment -- most
            # installations with fewer rules can likely use a lower value).
            OutlookQuitTimeoutSeconds = 20

            # Maximum seconds Connect-OutlookCOM's VisibleLaunch path waits
            # for Outlook to finish starting up and register its COM class
            # after Start-Process launches outlook.exe. Slower machines, or
            # profiles with many PST/OST files to load at startup, may need
            # more time; most environments can likely use less. Default: 30.
            OutlookLaunchTimeoutSeconds = 30
        }
    }
}


function Save-OMMigrateSelectedProfile {
    <#
    .SYNOPSIS
        Persists the operator-selected Outlook profile name to
        OMMigrate_Settings.json.

    .DESCRIPTION
        Called by Script 00 after the operator selects a profile
        in the WinForms picker (or after auto-selection when only
        one profile exists). Writes the profile name to the
        OutlookProfile.SelectedProfile key in Settings.json so
        all subsequent scripts (01-04) can read it without
        re-prompting the operator.

    .PARAMETER ProfileName
        The Outlook profile name to persist.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProfileName
    )

    $settingsFile = Join-Path $Global:OMMigrate.ConfigPath 'OMMigrate_Settings.json'

    try {
        # Load current settings from disk to avoid overwriting any keys
        # that may have been changed since Initialize-OMMigrate loaded them.
        $settingsObj = $null
        if (Test-Path $settingsFile) {
            $settingsObj = Get-Content $settingsFile -Raw | ConvertFrom-Json
        }
        else {
            $settingsObj = Get-DefaultSettings
        }

        # Add OutlookProfile section if missing (older settings files)
        if (-not $settingsObj.PSObject.Properties['OutlookProfile']) {
            $settingsObj | Add-Member -MemberType NoteProperty `
                                      -Name 'OutlookProfile' `
                                      -Value ([PSCustomObject]@{ SelectedProfile = '' })
        }

        $settingsObj.OutlookProfile.SelectedProfile = $ProfileName

        $settingsObj | ConvertTo-Json -Depth 10 |
            Set-Content -Path $settingsFile -Encoding UTF8

        # Also update the in-memory settings so the current session
        # can read the value without reloading from disk.
        if ($Global:OMMigrate.Settings) {
            if (-not $Global:OMMigrate.Settings.PSObject.Properties['OutlookProfile']) {
                $Global:OMMigrate.Settings | Add-Member -MemberType NoteProperty `
                                                        -Name 'OutlookProfile' `
                                                        -Value ([PSCustomObject]@{ SelectedProfile = '' })
            }
            $Global:OMMigrate.Settings.OutlookProfile.SelectedProfile = $ProfileName
        }

        Write-OMMigrateLog -Message "Selected Outlook profile saved to settings: '$ProfileName'" `
                           -Level INFO
    }
    catch {
        Write-OMMigrateLog -Message "Could not save selected profile to settings file: $_" `
                           -Level WARN
    }
}


# Added 2026-07-10, Administrator direction. Live-tested problem found: RulesEngine
# settings (ArchiveStoreMappings, MasterArchiveNames) are profile-specific --
# archive PST names, attached accounts, and which archives should be
# protected from auto-detach are all genuinely different per Outlook
# profile (confirmed live: "TestProfile" has 'OMMigrate Local Archive' +
# 'Test Archive 2' with test accounts; Administrator's real Outlook profile has a
# different 'OMMigrate Local Archive' plus 'Personal-Archive', with real accounts).
# A single flat OMMigrate_Settings.json shared across every profile means
# switching profiles would silently apply one profile's stale
# archive/account mappings to a completely different profile.
#
# Fix: one physical settings file PER profile on disk
# (OMMigrate_Settings_<ProfileName>.json), with OMMigrate_Settings.json
# itself continuing to act as the single "currently active" file every
# other function in this codebase already reads/writes -- so NOTHING else
# in Core.psm1, Outlook.psm1, or Scripts 00/01/03 needs to change. This
# function is called once, immediately after a profile is selected (all
# three selection paths -- -ProfileName param, single-profile auto-select,
# and the WinForms picker -- converge on one call site in Script 00), and
# BEFORE Save-OMMigrateSelectedProfile is called for that same profile.
function Switch-OMMigrateProfileSettings {
    <#
    .SYNOPSIS
        Swaps OMMigrate_Settings.json to the per-profile settings file for
        the given Outlook profile, creating it on first use.

    .DESCRIPTION
        Called by Script 00 immediately after a profile is selected (via
        -ProfileName, single-profile auto-select, or the WinForms picker),
        and BEFORE Save-OMMigrateSelectedProfile. Every other read/write of
        settings in this codebase continues to target the plain
        OMMigrate_Settings.json filename unchanged -- this function's only
        job is to make sure that file contains the correct profile's data
        before anything else touches it this run.

        Behavior:
          - If OMMigrate_Settings_<ProfileName>.json already exists, it is
            copied over OMMigrate_Settings.json (this profile's own saved
            settings become active).
          - If it does not exist yet (first time this profile has been
            used), the CURRENT OMMigrate_Settings.json is copied to
            OMMigrate_Settings_<ProfileName>.json first, so the new
            per-profile file starts from whatever settings already exist
            rather than silently discarding them. It is then treated as
            step 1 above (copied back over OMMigrate_Settings.json,
            functionally a no-op copy in this case, kept for a single
            consistent code path).
          - $Global:OMMigrate.Settings is reloaded from the now-current
            OMMigrate_Settings.json so the rest of THIS run sees the
            correct profile's settings immediately, not just future runs.

        Profile name is sanitized for filesystem safety using the same
        approach as Get-OMMigrateCsvPath's profile-suffix handling, for
        consistency with the existing per-profile CSV naming convention
        this project already uses (rules_inventory_<Profile>.csv, etc).

    .PARAMETER ProfileName
        The Outlook profile name that was just selected.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProfileName
    )

    $activeSettingsFile = Join-Path $Global:OMMigrate.ConfigPath 'OMMigrate_Settings.json'

    # Sanitize profile name for filesystem safety -- strip characters not
    # valid in Windows filenames, same approach as Get-OMMigrateCsvPath.
    $safeProfileName = ($ProfileName -replace '[\/:*?"<>|]', '_')
    $perProfileSettingsFile = Join-Path $Global:OMMigrate.ConfigPath "OMMigrate_Settings_$safeProfileName.json"

    try {
        if (Test-Path $perProfileSettingsFile) {
            # This profile has been used before -- activate its own settings.
            Copy-Item -Path $perProfileSettingsFile -Destination $activeSettingsFile -Force
            Write-OMMigrateLog -Message "Activated per-profile settings for '$ProfileName'." -Level INFO
        }
        else {
            # First time this profile has been selected -- seed its
            # per-profile file from whatever is currently active, rather
            # than starting from a blank Get-DefaultSettings and losing
            # any settings already on disk (e.g. Logging, Reporting
            # preferences the operator already configured).
            if (Test-Path $activeSettingsFile) {
                Copy-Item -Path $activeSettingsFile -Destination $perProfileSettingsFile -Force
            }
            else {
                Get-DefaultSettings | ConvertTo-Json -Depth 10 |
                    Set-Content -Path $perProfileSettingsFile -Encoding UTF8
                Copy-Item -Path $perProfileSettingsFile -Destination $activeSettingsFile -Force
            }
            Write-OMMigrateLog -Message "First use of profile '$ProfileName' -- created new per-profile settings file." -Level INFO
        }

        # Reload in-memory settings so the REST of this run sees the
        # correct profile's data immediately -- without this, Initialize-
        # OMMigrate's earlier load (before any profile was known) would
        # remain active in memory for the rest of the run despite the file
        # on disk now being correct for future runs.
        if (Test-Path $activeSettingsFile) {
            $Global:OMMigrate.Settings = Get-Content $activeSettingsFile -Raw | ConvertFrom-Json
        }
    }
    catch {
        Write-OMMigrateLog -Message "Could not switch to per-profile settings for '$ProfileName': $_ -- continuing with currently loaded settings." `
                           -Level WARN
    }
}


# Added 2026-07-10, Administrator direction. Write-back counterpart to
# Switch-OMMigrateProfileSettings above. Live-tested gap found: that
# function copies a profile's OWN settings file INTO the active
# OMMigrate_Settings.json at the START of a run, but nothing copies changes
# made DURING the rest of that same run (e.g. a new ArchiveStoreMappings
# selection from the Script 00 TargetStoreName picker) back OUT to the
# per-profile file -- so on the very next run for that same profile,
# Switch-OMMigrateProfileSettings would silently overwrite the fresh
# mapping with the stale one from the previous seed, reverting the
# operator's picker selection with no warning.
#
# Called once, near the end of Script 00 (after every picker that can
# write to Settings.json this run has already run), so the per-profile
# file captures the final state of everything from this run, not a
# snapshot from partway through.
function Sync-OMMigrateProfileSettings {
    <#
    .SYNOPSIS
        Copies the currently active OMMigrate_Settings.json back onto the
        per-profile settings file for the given Outlook profile.

    .DESCRIPTION
        Called by Script 00 near the end of its run (after the TargetStoreName
        picker and any other Settings.json-writing step has already
        completed), so that changes made during this run -- most notably
        RulesEngine.ArchiveStoreMappings from the TargetStoreName picker --
        are actually retained in that profile's own settings file
        (OMMigrate_Settings_<ProfileName>.json) rather than being silently
        lost the next time Switch-OMMigrateProfileSettings activates this
        profile again.

        Every other script/function in this codebase continues to
        read/write the plain OMMigrate_Settings.json filename unchanged --
        this function's only job is to make sure that file's current
        content also gets preserved in the durable per-profile copy before
        the run ends.

    .PARAMETER ProfileName
        The Outlook profile whose per-profile settings file should be
        updated from the currently active OMMigrate_Settings.json.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProfileName
    )

    $activeSettingsFile = Join-Path $Global:OMMigrate.ConfigPath 'OMMigrate_Settings.json'

    $safeProfileName = ($ProfileName -replace '[\/:*?"<>|]', '_')
    $perProfileSettingsFile = Join-Path $Global:OMMigrate.ConfigPath "OMMigrate_Settings_$safeProfileName.json"

    try {
        if (Test-Path $activeSettingsFile) {
            Copy-Item -Path $activeSettingsFile -Destination $perProfileSettingsFile -Force
            Write-OMMigrateLog -Message "Per-profile settings synced for '$ProfileName'." -Level INFO
        }
        else {
            Write-OMMigrateLog -Message "Could not sync per-profile settings for '$ProfileName' -- active settings file not found." `
                               -Level WARN
        }
    }
    catch {
        Write-OMMigrateLog -Message "Could not sync per-profile settings for '$ProfileName': $_" `
                           -Level WARN
    }
}


# EXTRACTED 2026-07-09 (Administrator direction) from OMMigrate-Core_WIP.psm1, a
# 6-day-old (2026-07-03) work-in-progress file -- see the RulesEngine
# settings-schema comment above for full context. NOT a completed feature;
# this is the settings-persistence half of the paused TargetStoreName
# hardcode fix, extracted as a jumpstart on resuming that work.
function Save-OMMigrateArchiveStoreMappings {
    <#
    .SYNOPSIS
        Persists the operator-selected TargetStoreName archive mappings to
        OMMigrate_Settings.json.

    .DESCRIPTION
        Called by Script 00 after the operator confirms selections in the
        TargetStoreName picker (shown every run, after standalone PST
        discovery -- defaults to the last saved mapping so the operator can
        just click OK to keep it, or change it before clicking OK).
        Writes the mapping array to the RulesEngine.ArchiveStoreMappings
        key in Settings.json so Export-RulesToCSV can read it on this and
        future runs without re-prompting.

        No PST or account name is hardcoded here or anywhere downstream --
        this function only persists whatever the picker collected from the
        live-detected stores/accounts on this machine/profile.

    .PARAMETER Mappings
        Array of PSCustomObject, each with a TargetStoreName (string) and
        RuleStoreNames (string array) property. May be an empty array to
        clear all mappings.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$Mappings
    )

    $settingsFile = Join-Path $Global:OMMigrate.ConfigPath 'OMMigrate_Settings.json'

    try {
        # Load current settings from disk to avoid overwriting any keys
        # that may have been changed since Initialize-OMMigrate loaded them.
        $settingsObj = $null
        if (Test-Path $settingsFile) {
            $settingsObj = Get-Content $settingsFile -Raw | ConvertFrom-Json
        }
        else {
            $settingsObj = Get-DefaultSettings
        }

        # Add RulesEngine section if missing (older settings files)
        if (-not $settingsObj.PSObject.Properties['RulesEngine']) {
            $settingsObj | Add-Member -MemberType NoteProperty `
                                      -Name 'RulesEngine' `
                                      -Value ([PSCustomObject]@{ ArchiveStoreMappings = @() })
        }

        $settingsObj.RulesEngine.ArchiveStoreMappings = $Mappings

        $settingsObj | ConvertTo-Json -Depth 10 |
            Set-Content -Path $settingsFile -Encoding UTF8

        # Also update the in-memory settings so the current session
        # can read the value without reloading from disk.
        if ($Global:OMMigrate.Settings) {
            if (-not $Global:OMMigrate.Settings.PSObject.Properties['RulesEngine']) {
                $Global:OMMigrate.Settings | Add-Member -MemberType NoteProperty `
                                                        -Name 'RulesEngine' `
                                                        -Value ([PSCustomObject]@{ ArchiveStoreMappings = @() })
            }
            $Global:OMMigrate.Settings.RulesEngine.ArchiveStoreMappings = $Mappings
        }

        Write-OMMigrateLog -Message "Archive store mappings saved to settings: $($Mappings.Count) mapping(s)." `
                           -Level INFO
    }
    catch {
        Write-OMMigrateLog -Message "Could not save archive store mappings to settings file: $_" `
                           -Level WARN
    }
}


function Get-OMMigrateCsvPath {
    <#
    .SYNOPSIS
        Returns the profile-suffixed path for an OMMigrate CSV file (or,
        via -BasePathOverride, any other OMMigrate-managed file).

    .DESCRIPTION
        All three OMMigrate control CSVs (migration_accounts.csv,
        folder_map.csv, rules_inventory.csv) are named with the selected
        Outlook profile appended so that runs under different profiles
        (e.g. Outlook vs "TestProfile") write to separate files and never
        overwrite each other.

        Example: profile 'Outlook' + base 'rules_inventory.csv'
                 returns Config\\rules_inventory_Outlook.csv

        If no profile is set in Settings.json the base filename is
        returned unchanged -- this preserves backward compatibility
        for environments where the profile picker has never run.

        Added 2026-07-10, Administrator direction (live-tested bug found): the same
        profile-collision problem this function already solves for CSVs
        also applied to per-account backup PST filenames in
        OMMigrate-01-Backup.ps1 (e.g. user_isp-domain_osttoimap.pst)
        -- those were NOT profile-suffixed, so running the same account
        under two different Outlook profiles (e.g. "TestProfile" and Administrator's real
        Outlook profile) silently overwrote one profile's backup with the
        other's on every run, confirmed live. Rather than duplicate this
        function's profile-lookup/sanitize logic separately in Script 01,
        -BasePathOverride lets any caller reuse the same suffixing logic
        against a different base directory (e.g. BackupPath instead of
        ConfigPath) -- fully backward compatible, every existing caller
        that doesn't pass this parameter is unaffected.

    .PARAMETER BaseName
        The base CSV (or other) filename without profile suffix.
        Example: 'rules_inventory.csv'

    .PARAMETER BasePathOverride
        Optional. Directory to join the profile-suffixed filename against,
        instead of the default $Global:OMMigrate.ConfigPath. Used for
        non-CSV files (e.g. backup PSTs, which live under BackupPath).

    .OUTPUTS
        [string] -- Full path to the profile-suffixed file.

    .EXAMPLE
        $path = Get-OMMigrateCsvPath -BaseName 'rules_inventory.csv'
        # Returns e.g. C:\\...\\Config\\rules_inventory_Outlook.csv

    .EXAMPLE
        $path = Get-OMMigrateCsvPath -BaseName 'user_isp-domain_osttoimap.pst' `
                                     -BasePathOverride $Global:OMMigrate.BackupPath
        # Returns e.g. C:\\...\\Backups\\user_isp-domain_osttoimap_Outlook.pst
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseName,

        [Parameter(Mandatory = $false)]
        [string]$BasePathOverride = ''
    )

    # Read selected profile from in-memory settings (set by Script 00 picker
    # and persisted to OMMigrate_Settings.json for Scripts 01-04).
    $profile = ''
    try {
        if ($Global:OMMigrate.Settings -and
            $Global:OMMigrate.Settings.PSObject.Properties['OutlookProfile'] -and
            $Global:OMMigrate.Settings.OutlookProfile.SelectedProfile) {
            $profile = $Global:OMMigrate.Settings.OutlookProfile.SelectedProfile
        }
    }
    catch { }

    # Sanitize profile name for use as a filename suffix --
    # strip characters not valid in Windows filenames.
    if ($profile) {
        $profile = $profile -replace '[\\/:*?"<>|]', '_'
        $profile = $profile.Trim()
    }

    # Build the final filename: base_stem + '_' + profile + ext
    # Example: 'rules_inventory.csv' + 'Outlook' -> 'rules_inventory_Outlook.csv'
    if ($profile) {
        $ext      = [System.IO.Path]::GetExtension($BaseName)      # '.csv'
        $stem     = [System.IO.Path]::GetFileNameWithoutExtension($BaseName)  # 'rules_inventory'
        $fileName = "${stem}_${profile}${ext}"
    }
    else {
        # No profile set -- use base filename unchanged (backward compatible)
        $fileName = $BaseName
    }

    $targetBasePath = if ($BasePathOverride) { $BasePathOverride } else { $Global:OMMigrate.ConfigPath }
    return Join-Path $targetBasePath $fileName
}


# ============================================================
#  REGION: LOGGING ENGINE
# ============================================================

function Write-OMMigrateLog {
    <#
    .SYNOPSIS
        Writes a timestamped entry to the run log and optionally to console.

    .DESCRIPTION
        Central logging function used by all OMMigrate modules and scripts.
        Respects the current LogLevel filter -- entries below the threshold
        are suppressed from both file and console output.

        Log format:
            [YYYY-MM-DD HH:MM:SS] [LEVEL] [ScriptName] Message

    .PARAMETER Message
        The log message text.

    .PARAMETER Level
        Severity level: DEBUG | INFO | WARN | ERROR
        Default: INFO

    .PARAMETER NoConsole
        When present, suppresses console output. File logging still occurs.

    .PARAMETER WhatIfPrefix
        When present, prepends [WHATIF] to the message.

    .EXAMPLE
        Write-OMMigrateLog -Message "Account discovered: user@domain.com" -Level INFO

    .EXAMPLE
        Write-OMMigrateLog -Message "PST file not found at expected path" -Level WARN
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('DEBUG','INFO','WARN','ERROR')]
        [string]$Level = 'INFO',

        [Parameter(Mandatory = $false)]
        [switch]$NoConsole,

        [Parameter(Mandatory = $false)]
        [switch]$WhatIfPrefix
    )

    # -- Level filter ------------------------------------------
    $currentLevel  = $Global:OMMigrate.LogLevelMap[$Global:OMMigrate.LogLevel]
    $messageLevel  = $Global:OMMigrate.LogLevelMap[$Level]
    if ($messageLevel -lt $currentLevel) { return }

    # -- Format entry ------------------------------------------
    $timestamp  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $script     = $Global:OMMigrate.ScriptName
    $prefix     = if ($WhatIfPrefix -or $Global:OMMigrate.WhatIf) { '[WHATIF] ' } else { '' }
    $entry      = "[$timestamp] [$Level] [$script] ${prefix}${Message}"

    # -- Apply sanitization filter if active -------------------
    if ($Global:OMMigrate.Sanitize -and
        $Global:OMMigrate.SanitizeMap -and
        $Global:OMMigrate.SanitizeMap.Count -gt 0) {
        $entry = Invoke-OMMigrateSanitize -Text $entry
    }

    # -- Write to run log file ---------------------------------
    if ($Global:OMMigrate.RunLogFile -and -not $Global:OMMigrate.WhatIf) {
        try {
            Add-Content -Path $Global:OMMigrate.RunLogFile -Value $entry -Encoding UTF8
        }
        catch {
            # Silently continue if log file is temporarily locked
        }
    }

    # -- Update counters ---------------------------------------
    $Global:OMMigrate.Counters.LogEntries++
    if ($Level -eq 'WARN')  { $Global:OMMigrate.Counters.Warnings++ }
    if ($Level -eq 'ERROR') { $Global:OMMigrate.Counters.Errors++ }

    # -- Write to console --------------------------------------
    if (-not $NoConsole) {
        $color = switch ($Level) {
            'DEBUG' { 'DarkGray'  }
            'INFO'  { 'Cyan'      }
            'WARN'  { 'Yellow'    }
            'ERROR' { 'Red'       }
            default { 'White'     }
        }

        # WhatIf entries get a distinct color
        if ($WhatIfPrefix -or $Global:OMMigrate.WhatIf) { $color = 'Magenta' }

        Write-Host $entry -ForegroundColor $color
    }
}


function Write-RunLogHeader {
    <#
    .SYNOPSIS
        Writes a structured session header block to the run log file.

    .DESCRIPTION
        Called once by Initialize-OMMigrate. Records tool version, script
        name, operator, machine, timestamp, and WhatIf mode at the top
        of each run log file for traceability.
    #>
    $separator = '=' * 72
    $header = @"
$separator
  $Script:OMMigrateProduct v$Script:OMMigrateVersion
  $Script:OMMigrateShort Run Log
$separator
  Script     : $($Global:OMMigrate.ScriptName)
  Session ID : $($Global:OMMigrate.SessionID)
  Started    : $($Global:OMMigrate.StartTime.ToString('yyyy-MM-dd HH:mm:ss'))
  Operator   : $($Global:OMMigrate.RunUser)
  Machine    : $($Global:OMMigrate.MachineName)
  WhatIf     : $($Global:OMMigrate.WhatIf)
  Log Level  : $($Global:OMMigrate.LogLevel)
  Base Path  : $($Global:OMMigrate.BasePath)
$separator
  Originator & Architect    : $Script:OMMigrateAuthor
  Implementation Specialist : $Script:OMMigrateAI
  Inception Date            : $Script:OMMigrateInception
$separator

"@
    if (-not $Global:OMMigrate.WhatIf -and $Global:OMMigrate.RunLogFile) {
        try {
            Set-Content -Path $Global:OMMigrate.RunLogFile -Value $header -Encoding UTF8
        }
        catch {
            Write-Warning "OMMigrate: Could not write run log header. $_"
        }
    }
}


function Write-RunLogFooter {
    <#
    .SYNOPSIS
        Writes a structured session footer to the run log file.

    .DESCRIPTION
        Called by Complete-OMMigrateSession at the end of each script run.
        Records elapsed time, summary counters, and final status.

    .PARAMETER Status
        Final status of the script run: SUCCESS | WARNING | FAILED
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('SUCCESS','WARNING','FAILED')]
        [string]$Status
    )

    $elapsed   = (Get-Date) - $Global:OMMigrate.StartTime
    $separator = '=' * 72
    $footer = @"

$separator
  Session Summary
$separator
  Status      : $Status
  Elapsed     : $($elapsed.ToString('hh\:mm\:ss'))
  Log Entries : $($Global:OMMigrate.Counters.LogEntries)
  Warnings    : $($Global:OMMigrate.Counters.Warnings)
  Errors      : $($Global:OMMigrate.Counters.Errors)
  Prompts     : $($Global:OMMigrate.Counters.Prompts)
  Confirmed   : $($Global:OMMigrate.Counters.Confirmed)
  Skipped     : $($Global:OMMigrate.Counters.Skipped)
  Completed   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
$separator
"@
    if (-not $Global:OMMigrate.WhatIf -and $Global:OMMigrate.RunLogFile) {
        try {
            Add-Content -Path $Global:OMMigrate.RunLogFile -Value $footer -Encoding UTF8
        }
        catch { }
    }
}


# ============================================================
#  REGION: AUDIT LOG
# ============================================================

function Write-AuditEntry {
    <#
    .SYNOPSIS
        Appends a structured JSON entry to the cumulative audit log.

    .DESCRIPTION
        The audit log (OMMigrate_Audit.log) is a cumulative, append-only
        record of every significant action taken across all sessions and
        script runs. It is never truncated or overwritten.

        Each entry is a single-line JSON object for machine parseability.
        Useful for compliance reporting, troubleshooting, and tracking
        the full history of what was changed on a machine.

        Audit entries are written regardless of LogLevel setting.
        Audit entries are suppressed in WhatIf mode.

    .PARAMETER Action
        Short action identifier in UPPER_SNAKE_CASE.
        Examples: SESSION_START, ACCOUNT_BACKED_UP, ACCOUNT_REMOVED,
                  IMAP_ADDED, FOLDER_MIGRATED, RULE_UPDATED, SESSION_END

    .PARAMETER Detail
        Human-readable detail string for the action.

    .PARAMETER AccountEmail
        Optional. Email address of the account being acted upon.

    .PARAMETER Outcome
        Optional. Result of the action: SUCCESS | SKIPPED | FAILED
        Default: SUCCESS

    .EXAMPLE
        Write-AuditEntry -Action 'ACCOUNT_BACKED_UP' `
                         -AccountEmail 'user@domain.com' `
                         -Detail 'PST exported to Backups\user_domain_com.pst (42 MB)'

    .EXAMPLE
        Write-AuditEntry -Action 'ACCOUNT_SKIPPED' `
                         -AccountEmail 'user@gmail.com' `
                         -Detail 'Account is already IMAP - no migration required' `
                         -Outcome 'SKIPPED'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action,

        [Parameter(Mandatory = $true)]
        [string]$Detail,

        [Parameter(Mandatory = $false)]
        [string]$AccountEmail = '',

        [Parameter(Mandatory = $false)]
        [ValidateSet('SUCCESS','SKIPPED','FAILED','WARNING')]
        [string]$Outcome = 'SUCCESS'
    )

    # Suppress audit writes in WhatIf mode
    if ($Global:OMMigrate.WhatIf) { return }
    if (-not $Global:OMMigrate.AuditLogFile) { return }

    $entry = [ordered]@{
        Timestamp   = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
        SessionID   = $Global:OMMigrate.SessionID
        Script      = $Global:OMMigrate.ScriptName
        Operator    = $Global:OMMigrate.RunUser
        Machine     = $Global:OMMigrate.MachineName
        Action      = $Action
        Account     = $AccountEmail
        Outcome     = $Outcome
        Detail      = $Detail
    }

    try {
        $json = $entry | ConvertTo-Json -Compress
        Add-Content -Path $Global:OMMigrate.AuditLogFile -Value $json -Encoding UTF8
    }
    catch {
        Write-OMMigrateLog -Message "Audit log write failed: $_" -Level WARN
    }
}


# ============================================================
#  REGION: CONSOLE OUTPUT HELPERS
# ============================================================

function Show-Banner {
    <#
    .SYNOPSIS
        Displays the OMMigrate product banner on the console.

    .DESCRIPTION
        Called once by Initialize-OMMigrate. Shows tool name, version,
        credits, and WhatIf mode warning if applicable.
    #>
    $line = '-' * 60
    Write-Host ''
    Write-Host $line -ForegroundColor DarkCyan
    Write-Host "  $Script:OMMigrateProduct  v$Script:OMMigrateVersion" `
               -ForegroundColor White
    Write-Host "  `"Automating the Outlook migration Google suggested couldn't be automated.`"" `
               -ForegroundColor DarkGray
    Write-Host $line -ForegroundColor DarkCyan
    Write-Host "  Originator & Architect    : $Script:OMMigrateAuthor" `
               -ForegroundColor Gray
    Write-Host "  Implementation Specialist : $Script:OMMigrateAI" `
               -ForegroundColor Gray
    Write-Host $line -ForegroundColor DarkCyan
    Write-Host "  Script  : $($Global:OMMigrate.ScriptName)" `
               -ForegroundColor Cyan
    Write-Host "  Operator: $(Invoke-OMMigrateSanitize -Text $Global:OMMigrate.RunUser)" `
               -ForegroundColor Cyan
    Write-Host "  Machine : $(Invoke-OMMigrateSanitize -Text $Global:OMMigrate.MachineName)" `
               -ForegroundColor Cyan
    Write-Host "  Started : $($Global:OMMigrate.StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" `
               -ForegroundColor Cyan

    if ($Global:OMMigrate.WhatIf) {
        Write-Host ''
        Write-Host '  *** WHATIF MODE -- NO CHANGES WILL BE MADE ***' `
                   -ForegroundColor Magenta
    }

    Write-Host $line -ForegroundColor DarkCyan
    Write-Host ''
}


function Show-SectionHeader {
    <#
    .SYNOPSIS
        Displays a formatted section header on the console.

    .DESCRIPTION
        Used within scripts to visually separate phases of execution
        on the console, making progress easy to follow.

    .PARAMETER Title
        Section title text.

    .PARAMETER Step
        Optional step number (e.g. '1 of 4').

    .EXAMPLE
        Show-SectionHeader -Title "Discovering Outlook Accounts" -Step "1 of 4"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $false)]
        [string]$Step = ''
    )

    $stepText = if ($Step) { "  [Step $Step]" } else { '' }
    Write-Host ''
    Write-Host ('-' * 60) -ForegroundColor DarkCyan
    Write-Host "  $Title$stepText" -ForegroundColor White
    Write-Host ('-' * 60) -ForegroundColor DarkCyan
    Write-Host ''
    Write-OMMigrateLog -Message "=== $Title ===" -Level INFO -NoConsole
}


function Show-AccountStatus {
    <#
    .SYNOPSIS
        Displays a color-coded per-account status line on the console.

    .DESCRIPTION
        Used during account processing loops to show progress clearly.
        Each account gets a single formatted status line with its
        email address, detected type tag, and action being taken.

    .PARAMETER Email
        Account email address.

    .PARAMETER Tag
        Account classification tag (e.g. POP3-STANDARD, IMAP-ALREADY).

    .PARAMETER Action
        Short description of action being taken or status.

    .PARAMETER Status
        Visual status indicator: OK | SKIP | WARN | FAIL | INFO
        Default: INFO

    .EXAMPLE
        Show-AccountStatus -Email 'user@domain.com' `
                           -Tag 'POP3-STANDARD' `
                           -Action 'Backing up PST' `
                           -Status 'OK'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email,

        [Parameter(Mandatory = $true)]
        [string]$Tag,

        [Parameter(Mandatory = $true)]
        [string]$Action,

        [Parameter(Mandatory = $false)]
        [ValidateSet('OK','SKIP','WARN','FAIL','INFO')]
        [string]$Status = 'INFO'
    )

    $statusColor = switch ($Status) {
        'OK'   { 'Green'   }
        'SKIP' { 'DarkGray'}
        'WARN' { 'Yellow'  }
        'FAIL' { 'Red'     }
        'INFO' { 'Cyan'    }
    }

    $statusIcon = switch ($Status) {
        'OK'   { '[OK]   ' }
        'SKIP' { '[SKIP] ' }
        'WARN' { '[WARN] ' }
        'FAIL' { '[FAIL] ' }
        'INFO' { '[....] ' }
    }

    $tagPadded   = $Tag.PadRight(20)
    $emailPadded = (Invoke-OMMigrateSanitize -Text $Email).PadRight(35)

    Write-Host "  $statusIcon " -ForegroundColor $statusColor -NoNewline
    Write-Host "$emailPadded" -ForegroundColor White -NoNewline
    Write-Host "$tagPadded" -ForegroundColor DarkGray -NoNewline
    Write-Host "$Action" -ForegroundColor Gray
}


# ============================================================
#  REGION: USER PROMPT / CONFIRMATION
# ============================================================

function Confirm-Action {
    <#
    .SYNOPSIS
        Presents a Y/N confirmation prompt to the operator and returns
        the response as a boolean.

    .DESCRIPTION
        All destructive or significant actions in OMMigrate scripts must
        pass through this function before executing. Provides a consistent
        prompt format, logs the question and the operator's response,
        and writes an audit entry.

        In WhatIf mode, always returns $true (simulates confirmation)
        without prompting, and logs a [WHATIF] entry.

        EXIT HANDLING:
        The operator may type EXIT, QUIT, Q, or STOP at any Y/N prompt
        to trigger a controlled graceful exit. This releases the Outlook
        COM session, writes final log entries, saves a checkpoint, and
        displays recovery instructions before exiting cleanly.
        This is the recommended way to stop a script mid-run.

    .PARAMETER Message
        The confirmation question to display to the operator.

    .PARAMETER AccountEmail
        Optional. Email address for audit log context.

    .PARAMETER DefaultYes
        When $true, pressing Enter without typing defaults to Y.
        Default: $false (Enter defaults to N -- safer for destructive ops).

    .OUTPUTS
        [bool] -- $true if operator confirmed, $false if declined.
        Never returns if operator types EXIT -- calls Invoke-OMMigrateGracefulExit.

    .EXAMPLE
        $proceed = Confirm-Action -Message "Remove POP3 account user@domain.com?"
        if ($proceed) { ... }

    .EXAMPLE
        $proceed = Confirm-Action -Message "Export PST for user@domain.com?" `
                                  -AccountEmail 'user@domain.com' `
                                  -DefaultYes $true
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [string]$AccountEmail = '',

        [Parameter(Mandatory = $false)]
        [bool]$DefaultYes = $false
    )

    $Global:OMMigrate.Counters.Prompts++

    # -- WhatIf mode -- simulate confirmation -------------------
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message "PROMPT (auto-confirmed in WhatIf): $Message" `
                           -Level INFO -WhatIfPrefix
        Write-AuditEntry  -Action 'PROMPT_WHATIF' `
                          -Detail $Message `
                          -AccountEmail $AccountEmail
        $Global:OMMigrate.Counters.Confirmed++
        return $true
    }

    # -- Live prompt -------------------------------------------
    $hint    = if ($DefaultYes) { '[Y/n/EXIT]' } else { '[y/N/EXIT]' }
    $default = if ($DefaultYes) { 'Y' }          else { 'N' }

    Write-Host ''
    Write-Host "  ? $Message" -ForegroundColor Yellow
    Write-Host "    $hint : " -ForegroundColor Yellow -NoNewline

    $response = Read-Host
    if ([string]::IsNullOrWhiteSpace($response)) { $response = $default }

    # -- EXIT / QUIT / STOP detection -------------------------
    # Checked before Y/N so operator can always get out cleanly
    $exitKeywords = @('EXIT', 'QUIT', 'Q', 'STOP', 'CANCEL', 'ABORT')
    if ($exitKeywords -contains $response.Trim().ToUpper()) {
        Write-OMMigrateLog -Message "Operator requested graceful exit at prompt: '$Message'" `
                           -Level WARN
        Write-AuditEntry  -Action 'OPERATOR_EXIT_REQUESTED' `
                          -Detail "Exit keyword entered at prompt: '$Message'" `
                          -AccountEmail $AccountEmail `
                          -Outcome 'SKIPPED'
        Invoke-OMMigrateGracefulExit -Reason "Operator typed '$($response.Trim().ToUpper())' at confirmation prompt"
        # Invoke-OMMigrateGracefulExit calls exit -- execution never reaches here
        return $false
    }

    $confirmed = $response.Trim().ToUpper() -eq 'Y'

    # -- Log and audit the response ----------------------------
    $responseText = if ($confirmed) { 'CONFIRMED' } else { 'DECLINED' }
    Write-OMMigrateLog -Message "Prompt: '$Message' -> $responseText" -Level INFO
    Write-AuditEntry  -Action "PROMPT_$responseText" `
                      -Detail $Message `
                      -AccountEmail $AccountEmail `
                      -Outcome $(if ($confirmed) { 'SUCCESS' } else { 'SKIPPED' })

    if ($confirmed) {
        $Global:OMMigrate.Counters.Confirmed++
    }
    else {
        $Global:OMMigrate.Counters.Skipped++
        Write-Host "  Skipped." -ForegroundColor DarkGray
    }

    Write-Host ''
    return $confirmed
}


function Wait-UserKeypress {
    <#
    .SYNOPSIS
        Pauses script execution and waits for the operator to press
        any key before returning to the PowerShell prompt.

    .DESCRIPTION
        Called at the very end of every script's finally block to
        prevent the PowerShell window from closing before the operator
        has had a chance to read the final summary output.

        This ensures a consistent experience regardless of how
        PowerShell was launched -- taskbar shortcut, Start menu,
        File Explorer, or terminal -- without requiring the operator
        to configure -NoExit on their shortcut.

    .EXAMPLE
        # Always the last call in a script's finally block:
        Wait-UserKeypress
    #>
    [CmdletBinding()]
    param()

    Write-Host ''
    Write-Host '  Press any key to return to the PowerShell prompt...' `
               -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Write-Host ''
}


function Show-PreflightWarning {
    <#
    .SYNOPSIS
        Displays a pre-flight warning block and requires operator
        acknowledgment before the script proceeds.

    .DESCRIPTION
        Called at the start of scripts that will make changes (Scripts
        01, 02, 03). Displays a clear summary of what the script is
        about to do, reminds the operator to verify backups exist, and
        requires explicit Y confirmation before any work begins.

    .PARAMETER ScriptDescription
        One or two sentence description of what this script does.

    .PARAMETER Prerequisites
        Array of prerequisite checks the operator should have completed.

    .EXAMPLE
        Show-PreflightWarning `
            -ScriptDescription "This script will remove POP3 accounts and add IMAP accounts." `
            -Prerequisites @(
                "Script 01 (Backup) completed successfully",
                "Backup PST files verified in .\Backups\ folder",
                "Secure Mail Key generated for ameritech.net account"
            )
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptDescription,

        [Parameter(Mandatory = $false)]
        [string[]]$Prerequisites = @(),

        [Parameter(Mandatory = $false)]
        [bool]$DefaultYes = $false,

        [Parameter(Mandatory = $false)]
        [string[]]$DeclineMessage = @()
    )

    Write-Host ''
    Write-Host ('!' * 60) -ForegroundColor Yellow
    Write-Host '  PRE-FLIGHT CHECK' -ForegroundColor Yellow
    Write-Host ('!' * 60) -ForegroundColor Yellow
    Write-Host ''
    Write-Host "  $ScriptDescription" -ForegroundColor White
    Write-Host ''

    if ($Prerequisites.Count -gt 0) {
        Write-Host '  Before continuing, confirm the following:' `
                   -ForegroundColor Yellow
        Write-Host ''
        foreach ($prereq in $Prerequisites) {
            Write-Host "    [ ] $prereq" -ForegroundColor Gray
        }
        Write-Host ''
    }

    $proceed = Confirm-Action -Message 'Pre-flight checks confirmed. Proceed?' `
                              -DefaultYes $DefaultYes

    if (-not $proceed) {
        Write-OMMigrateLog -Message 'Operator declined pre-flight. Script aborted.' `
                           -Level WARN
        Write-Host ''
        Write-Host '  Script aborted -- you pressed Enter or N at the pre-flight prompt.' `
                   -ForegroundColor Yellow
        if ($DeclineMessage.Count -gt 0) {
            Write-Host ''
            foreach ($line in $DeclineMessage) {
                Write-Host "  $line" -ForegroundColor Cyan
            }
        }
        Write-Host ''
        exit 0
    }
}


# ============================================================
#  REGION: ENVIRONMENT VALIDATION
# ============================================================

function Test-OMMigrateEnvironment {
    <#
    .SYNOPSIS
        Validates that the current environment meets all requirements
        for OMMigrate to run successfully.

    .DESCRIPTION
        Performs a series of pre-flight environment checks and returns
        a result object. Scripts call this during initialization and
        abort if critical checks fail.

        Checks performed:
            - PowerShell version (5.1 minimum)
            - Windows OS (10 or 11, 64-bit)
            - Outlook installation detected
            - Outlook version compatibility (16.0 = 2016/2019/2021)
            - NOT running as SYSTEM or built-in Administrator
            - Outlook is NOT currently running (required for COM automation)

    .OUTPUTS
        PSCustomObject with properties:
            Passed      [bool]   -- All critical checks passed
            Warnings    [array]  -- Non-critical issues found
            Failures    [array]  -- Critical issues that block execution
            OutlookVersion [string] -- Detected Outlook version string

    .EXAMPLE
        $env = Test-OMMigrateEnvironment
        if (-not $env.Passed) {
            foreach ($fail in $env.Failures) {
                Write-OMMigrateLog -Message $fail -Level ERROR
            }
            exit 1
        }
    #>
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        Passed         = $true
        Warnings       = [System.Collections.Generic.List[string]]::new()
        Failures       = [System.Collections.Generic.List[string]]::new()
        OutlookVersion = ''
        OutlookPath    = ''
        PSVersion      = $PSVersionTable.PSVersion.ToString()
        OSVersion      = [System.Environment]::OSVersion.VersionString
    }

    Write-OMMigrateLog -Message 'Running environment pre-flight checks...' -Level INFO

    # -- PowerShell version ------------------------------------
    if ($PSVersionTable.PSVersion.Major -lt 5 -or
        ($PSVersionTable.PSVersion.Major -eq 5 -and
         $PSVersionTable.PSVersion.Minor -lt 1)) {
        $result.Failures.Add(
            "PowerShell 5.1 or higher required. Found: $($PSVersionTable.PSVersion)"
        )
        $result.Passed = $false
    }
    else {
        Write-OMMigrateLog -Message "PowerShell version OK: $($PSVersionTable.PSVersion)" `
                           -Level DEBUG
    }

    # -- 64-bit OS ---------------------------------------------
    if (-not [System.Environment]::Is64BitOperatingSystem) {
        $result.Failures.Add('OMMigrate requires a 64-bit Windows operating system.')
        $result.Passed = $false
    }

    # -- Windows 10 / 11 --------------------------------------
    $osVersion = [System.Environment]::OSVersion.Version
    if ($osVersion.Major -lt 10) {
        $result.Failures.Add(
            "Windows 10 or Windows 11 required. Detected OS version: $($osVersion.ToString())"
        )
        $result.Passed = $false
    }
    else {
        Write-OMMigrateLog -Message "OS version OK: $([System.Environment]::OSVersion.VersionString)" `
                           -Level DEBUG
    }

    # -- Outlook installation ----------------------------------
    $outlookPaths = @(
        "${env:ProgramFiles}\Microsoft Office\root\Office16\OUTLOOK.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OUTLOOK.EXE",
        "${env:ProgramFiles}\Microsoft Office\Office16\OUTLOOK.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OUTLOOK.EXE"
    )

    $outlookExe = $null
    foreach ($path in $outlookPaths) {
        if (Test-Path $path) {
            $outlookExe = $path
            break
        }
    }

    # Also check registry for installed Outlook path
    if (-not $outlookExe) {
        try {
            $regPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'
            if (Test-Path $regPath) {
                $outlookExe = (Get-ItemProperty -Path $regPath).'(default)'
            }
        }
        catch { }
    }

    if (-not $outlookExe -or -not (Test-Path $outlookExe)) {
        $result.Failures.Add(
            'Classic Outlook 2016/2019/2021 not found. ' +
            'New Outlook (web-based) is not supported by OMMigrate.'
        )
        $result.Passed = $false
    }
    else {
        $result.OutlookPath = $outlookExe
        try {
            $verInfo = (Get-Item $outlookExe).VersionInfo
            $result.OutlookVersion = $verInfo.FileVersion
            Write-OMMigrateLog -Message "Outlook found: $outlookExe (v$($verInfo.FileVersion))" `
                               -Level DEBUG

            # Version compatibility check (16.0 = Office 2016/2019/2021/365)
            if ($verInfo.FileMajorPart -lt 16) {
                $result.Warnings.Add(
                    "Outlook version $($verInfo.FileVersion) is older than 2016. " +
                    'OMMigrate is tested against Outlook 2016 and later. ' +
                    'Proceed with caution.'
                )
            }
        }
        catch {
            $result.Warnings.Add("Could not read Outlook version info: $_")
        }
    }

    # -- Outlook not currently running -------------------------
    $outlookProcess = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
    if ($outlookProcess) {
        Write-OMMigrateLog -Message 'Outlook is running -- prompting operator to close it.' `
                           -Level INFO
        Write-Host ''
        Write-Host '  Outlook is currently running.' -ForegroundColor Cyan
        Write-Host '  OMMigrate requires Outlook to be closed before proceeding.' -ForegroundColor Cyan
        Write-Host ''
        Write-Host '  Close Outlook automatically and continue? [Y/n] ' `
                   -ForegroundColor Cyan -NoNewline
        $closeChoice = Read-Host
        if ($closeChoice -eq '' -or $closeChoice -match '^[Yy]') {
            Write-Host '  Closing Outlook...' -ForegroundColor Cyan
            try {
                $outlookProcess | Stop-Process -Force -ErrorAction Stop
                Start-Sleep -Seconds 3
                $stillRunning = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
                if ($stillRunning) {
                    $result.Failures.Add(
                        'Could not close Outlook automatically. Please close it manually and re-run.'
                    )
                    $result.Passed = $false
                }
                else {
                    Write-Host '  Outlook closed successfully.' -ForegroundColor Green
                    Write-OMMigrateLog -Message 'Outlook closed automatically by operator request.' `
                                       -Level INFO
                }
            }
            catch {
                $result.Failures.Add(
                    "Could not close Outlook automatically: $_. Please close it manually and re-run."
                )
                $result.Passed = $false
            }
        }
        else {
            $result.Failures.Add(
                'Outlook is currently running. Please close Outlook completely ' +
                'before running OMMigrate scripts. COM automation requires ' +
                'exclusive access to the Outlook profile.'
            )
            $result.Passed = $false
        }
    }
    else {
        Write-OMMigrateLog -Message 'Outlook is not running -- OK to proceed.' `
                           -Level DEBUG
    }

    # -- Not running as SYSTEM ---------------------------------
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    if ($identity.IsSystem) {
        $result.Failures.Add(
            'OMMigrate must not be run as the SYSTEM account. ' +
            'Run as the Windows user who owns the Outlook profile.'
        )
        $result.Passed = $false
    }

    # -- Summarize ---------------------------------------------
    foreach ($warn in $result.Warnings) {
        Write-OMMigrateLog -Message "Environment warning: $warn" -Level WARN
    }
    foreach ($fail in $result.Failures) {
        Write-OMMigrateLog -Message "Environment FAILURE: $fail" -Level ERROR
    }

    if ($result.Passed) {
        Write-OMMigrateLog -Message 'Environment pre-flight: All checks passed.' `
                           -Level INFO
    }
    else {
        Write-OMMigrateLog -Message 'Environment pre-flight: Critical failures detected. Cannot continue.' `
                           -Level ERROR
    }

    return $result
}


# ============================================================
#  REGION: MANIFEST MANAGEMENT
# ============================================================

function Write-StepManifest {
    <#
    .SYNOPSIS
        Writes a completion manifest JSON file for a completed script step.

    .DESCRIPTION
        Each OMMigrate script writes a manifest upon successful completion.
        The next script in the sequence reads and validates this manifest
        before proceeding. This is the primary gate-control mechanism that
        prevents Script 02 running without Script 01, etc.

        Manifest files are stored in .\Manifests\ and named by step number.

    .PARAMETER Step
        Step number: 0 | 1 | 2 | 3

    .PARAMETER Status
        Completion status: SUCCESS | WARNING | FAILED

    .PARAMETER Data
        Optional hashtable of step-specific data to embed in the manifest.
        For example, Script 01 embeds the list of verified backup files.

    .EXAMPLE
        Write-StepManifest -Step 1 -Status 'SUCCESS' -Data @{
            BackedUpAccounts = @('user@domain.com', 'other@domain.com')
            TotalBackupSizeMB = 1240
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(0,4)]
        [int]$Step,

        [Parameter(Mandatory = $true)]
        [ValidateSet('SUCCESS','WARNING','FAILED')]
        [string]$Status,

        [Parameter(Mandatory = $false)]
        [hashtable]$Data = @{}
    )

    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message "WhatIf: Would write Step $Step manifest (Status=$Status)" `
                           -Level INFO -WhatIfPrefix
        return
    }

    $manifest = [ordered]@{
        Product       = $Script:OMMigrateProduct
        Version       = $Script:OMMigrateVersion
        Step          = $Step
        ScriptName    = $Global:OMMigrate.ScriptName
        Status        = $Status
        CompletedAt   = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
        SessionID     = $Global:OMMigrate.SessionID
        Operator      = $Global:OMMigrate.RunUser
        Machine       = $Global:OMMigrate.MachineName
        Data          = $Data
    }

    $manifestFile = Join-Path $Global:OMMigrate.ManifestPath "Step0${Step}_Complete.json"
    try {
        $manifest | ConvertTo-Json -Depth 10 |
            Set-Content -Path $manifestFile -Encoding UTF8
        Write-OMMigrateLog -Message "Step $Step manifest written: $manifestFile" `
                           -Level INFO
        Write-AuditEntry  -Action "MANIFEST_STEP${Step}_WRITTEN" `
                          -Detail "Status=$Status | File=$manifestFile"
    }
    catch {
        Write-OMMigrateLog -Message "Failed to write Step $Step manifest: $_" `
                           -Level ERROR
        throw
    }
}


function Read-StepManifest {
    <#
    .SYNOPSIS
        Reads and validates a prerequisite step manifest.

    .DESCRIPTION
        Called at the start of Scripts 01, 02, and 03 to verify the
        previous step completed successfully before proceeding.

        Returns the manifest object if valid, throws a terminating
        error if the manifest is missing, corrupt, or shows FAILED status.

    .PARAMETER Step
        The step number whose manifest must be present and valid.

    .PARAMETER RequireStatus
        The status the manifest must show to be considered valid.
        Default: SUCCESS

    .OUTPUTS
        PSCustomObject -- the parsed manifest if valid.

    .EXAMPLE
        # Script 02 requires Script 01 to have succeeded
        $manifest = Read-StepManifest -Step 1
        $backedUpAccounts = $manifest.Data.BackedUpAccounts
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(0,4)]
        [int]$Step,

        [Parameter(Mandatory = $false)]
        [ValidateSet('SUCCESS','WARNING')]
        [string]$RequireStatus = 'SUCCESS'
    )

    $manifestFile = Join-Path $Global:OMMigrate.ManifestPath "Step0${Step}_Complete.json"

    if (-not (Test-Path $manifestFile)) {
        $msg = "Step $Step manifest not found at: $manifestFile`n" +
               "Script $Step must complete successfully before this script can run.`n" +
               "Please run OMMigrate-0${Step}-*.ps1 first."
        Write-OMMigrateLog -Message $msg -Level ERROR
        throw $msg
    }

    try {
        $manifest = Get-Content $manifestFile -Raw | ConvertFrom-Json
    }
    catch {
        $msg = "Step $Step manifest is corrupt or unreadable: $manifestFile`nError: $_"
        Write-OMMigrateLog -Message $msg -Level ERROR
        throw $msg
    }

    if ($manifest.Status -eq 'FAILED') {
        $msg = "Step $Step manifest shows FAILED status. Cannot proceed.`n" +
               "Please re-run OMMigrate-0${Step}-*.ps1 and resolve errors before continuing."
        Write-OMMigrateLog -Message $msg -Level ERROR
        throw $msg
    }

    # Machine/version consistency check
    if ($manifest.Machine -ne $Global:OMMigrate.MachineName) {
        Write-OMMigrateLog -Message (
            "WARNING: Step $Step manifest was created on machine '$($manifest.Machine)' " +
            "but this script is running on '$($Global:OMMigrate.MachineName)'. " +
            "Proceeding but verify this is intentional."
        ) -Level WARN
    }

    Write-OMMigrateLog -Message (
        "Step $Step manifest validated: Status=$($manifest.Status) | " +
        "Completed=$($manifest.CompletedAt) | Operator=$($manifest.Operator)"
    ) -Level INFO

    return $manifest
}


# ============================================================
#  REGION: HTML REPORT HELPERS
# ============================================================

function Get-HtmlReportHeader {
    <#
    .SYNOPSIS
        Returns the HTML header block for OMMigrate reports.

    .DESCRIPTION
        Generates a fully styled HTML page header including embedded CSS.
        Used by all OMMigrate report generators to ensure consistent,
        professional report appearance.

        Reports are designed to be:
            - Printable (client deliverable for IT consultants)
            - Self-contained (single HTML file, no external dependencies)
            - Color-coded (green/amber/red status indicators)

    .PARAMETER Title
        Report title shown in the page header.

    .PARAMETER Subtitle
        Optional subtitle or description shown below the title.

    .OUTPUTS
        [string] -- HTML string to prepend to report content.

    .EXAMPLE
        $html  = Get-HtmlReportHeader -Title 'Discovery Report' `
                                      -Subtitle 'Account Inventory -- Pre-Migration'
        $html += Get-HtmlReportSection -Title 'Accounts Found' -Content $tableHtml
        $html += Get-HtmlReportFooter
        $html | Set-Content -Path $reportFile -Encoding UTF8
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $false)]
        [string]$Subtitle = ''
    )

    $timestamp   = Get-Date -Format 'MMMM dd, yyyy  HH:mm:ss'
    $subtitleHtml = if ($Subtitle) { "<p class='subtitle'>$Subtitle</p>" } else { '' }

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OMMigrate -- $Title</title>
<style>
  :root {
    --bg:          #0d1117;
    --surface:     #161b22;
    --surface2:    #1c2128;
    --border:      #30363d;
    --text:        #e6edf3;
    --text-muted:  #8b949e;
    --accent:      #58a6ff;
    --green:       #3fb950;
    --yellow:      #d29922;
    --red:         #f85149;
    --gray:        #6e7681;
    --tag-pop3:    #1f4a7a;
    --tag-imap:    #1a4731;
    --tag-exch:    #3d2b00;
    --tag-skip:    #2d2d2d;
    --font-mono:   'Consolas', 'Courier New', monospace;
    --font-body:   'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    background:  var(--bg);
    color:       var(--text);
    font-family: var(--font-body);
    font-size:   14px;
    line-height: 1.6;
    padding:     32px;
  }

  .report-header {
    border-bottom: 2px solid var(--accent);
    padding-bottom: 24px;
    margin-bottom: 32px;
  }

  .product-name {
    font-size:      11px;
    letter-spacing: 3px;
    text-transform: uppercase;
    color:          var(--accent);
    margin-bottom:  8px;
  }

  h1 {
    font-size:   28px;
    font-weight: 600;
    color:       var(--text);
    margin-bottom: 8px;
  }

  .subtitle {
    color:       var(--text-muted);
    font-size:   15px;
    margin-bottom: 16px;
  }

  .meta-grid {
    display:               grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap:                   12px;
    margin-top:            16px;
  }

  .meta-item {
    background:    var(--surface);
    border:        1px solid var(--border);
    border-radius: 6px;
    padding:       10px 14px;
  }

  .meta-label {
    font-size:  10px;
    letter-spacing: 1px;
    text-transform: uppercase;
    color:      var(--text-muted);
    margin-bottom: 2px;
  }

  .meta-value {
    font-family: var(--font-mono);
    font-size:   13px;
    color:       var(--text);
  }

  .section {
    background:    var(--surface);
    border:        1px solid var(--border);
    border-radius: 8px;
    margin-bottom: 24px;
    overflow:      hidden;
  }

  .section-header {
    background:  var(--surface2);
    border-bottom: 1px solid var(--border);
    padding:     12px 20px;
    font-weight: 600;
    font-size:   13px;
    letter-spacing: 0.5px;
    color:       var(--accent);
    text-transform: uppercase;
  }

  .section-body { padding: 20px; }

  table {
    width:           100%;
    border-collapse: collapse;
    font-size:       13px;
  }

  th {
    background:    var(--surface2);
    border-bottom: 1px solid var(--border);
    padding:       8px 12px;
    text-align:    left;
    font-size:     11px;
    letter-spacing: 1px;
    text-transform: uppercase;
    color:         var(--text-muted);
  }

  td {
    padding:       9px 12px;
    border-bottom: 1px solid var(--border);
    font-family:   var(--font-mono);
    font-size:     12px;
    vertical-align: middle;
  }

  tr:last-child td { border-bottom: none; }
  tr:hover td { background: var(--surface2); }

  .badge {
    display:       inline-block;
    padding:       2px 8px;
    border-radius: 4px;
    font-size:     11px;
    font-weight:   600;
    letter-spacing: 0.5px;
  }

  .badge-pop3     { background: var(--tag-pop3);  color: #79c0ff; }
  .badge-imap     { background: var(--tag-imap);  color: #56d364; }
  .badge-exchange { background: var(--tag-exch);  color: #e3b341; }
  .badge-skip     { background: var(--tag-skip);  color: #8b949e; }
  .badge-update   { background: #3d2e00;           color: #e3b341; }  /* Amber -- signals CSV update needed */

  .status-ok   { color: var(--green);  font-weight: 600; }
  .status-warn { color: var(--yellow); font-weight: 600; }
  .status-fail { color: var(--red);    font-weight: 600; }
  .status-skip { color: var(--gray);   font-weight: 600; }

  .stat-grid {
    display:               grid;
    grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
    gap:                   16px;
    margin-bottom:         24px;
  }

  .stat-card {
    background:    var(--surface);
    border:        1px solid var(--border);
    border-radius: 8px;
    padding:       16px 20px;
    text-align:    center;
  }

  .stat-number {
    font-size:   32px;
    font-weight: 700;
    color:       var(--accent);
    line-height: 1;
    margin-bottom: 4px;
  }

  .stat-label {
    font-size:  11px;
    letter-spacing: 1px;
    text-transform: uppercase;
    color:      var(--text-muted);
  }

  .alert {
    border-radius: 6px;
    padding:       12px 16px;
    margin-bottom: 16px;
    font-size:     13px;
  }

  .alert-warn {
    background:  rgba(210,153,34,0.12);
    border-left: 3px solid var(--yellow);
    color:       #e3b341;
  }

  .alert-info {
    background:  rgba(88,166,255,0.08);
    border-left: 3px solid var(--accent);
    color:       var(--accent);
  }

  .alert-ok {
    background:  rgba(63,185,80,0.08);
    border-left: 3px solid var(--green);
    color:       var(--green);
  }

  .credits {
    margin-top:  40px;
    padding-top: 20px;
    border-top:  1px solid var(--border);
    font-size:   11px;
    color:       var(--text-muted);
    text-align:  center;
    line-height: 1.8;
  }

  .credits strong { color: var(--text); }

  code {
    background:    var(--surface2);
    border:        1px solid var(--border);
    border-radius: 3px;
    padding:       1px 5px;
    font-family:   var(--font-mono);
    font-size:     12px;
  }

  @media print {
    body { background: white; color: black; padding: 20px; }
    .section { border: 1px solid #ccc; }
    .section-header { background: #f0f0f0; color: #333; }
    th { background: #f0f0f0; color: #333; }
  }
</style>
</head>
<body>
<div class="report-header">
  <div class="product-name">OutlookMailMigrator (OMMigrate) v$Script:OMMigrateVersion</div>
  <h1>$Title</h1>
  $subtitleHtml
  <div class="meta-grid">
    <div class="meta-item">
      <div class="meta-label">Generated</div>
      <div class="meta-value">$timestamp</div>
    </div>
    <div class="meta-item">
      <div class="meta-label">Operator</div>
      <div class="meta-value">$($Global:OMMigrate.RunUser)</div>
    </div>
    <div class="meta-item">
      <div class="meta-label">Machine</div>
      <div class="meta-value">$($Global:OMMigrate.MachineName)</div>
    </div>
    <div class="meta-item">
      <div class="meta-label">Session ID</div>
      <div class="meta-value">$($Global:OMMigrate.SessionID.Substring(0,8))...</div>
    </div>
  </div>
</div>
"@
}


function Get-HtmlReportSection {
    <#
    .SYNOPSIS
        Returns an HTML section block for inclusion in an OMMigrate report.

    .DESCRIPTION
        Wraps a block of HTML content in a styled section container with
        a labeled header. Multiple sections are concatenated to build
        a complete report page.

    .PARAMETER Title
        Section header text.

    .PARAMETER Content
        HTML content for the section body (tables, alerts, paragraphs).

    .OUTPUTS
        [string] -- HTML string for this section.

    .EXAMPLE
        $section = Get-HtmlReportSection -Title 'Account Inventory' -Content $tableHtml
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    return @"
<div class="section">
  <div class="section-header">$Title</div>
  <div class="section-body">
    $Content
  </div>
</div>
"@
}


function Get-HtmlReportFooter {
    <#
    .SYNOPSIS
        Returns the HTML footer block for OMMigrate reports.

    .DESCRIPTION
        Closes the HTML document with credits, product information,
        and the OMMigrate tagline. Append to the end of every report.

    .OUTPUTS
        [string] -- HTML closing string.

    .EXAMPLE
        $html += Get-HtmlReportFooter
        $html | Set-Content -Path $reportFile -Encoding UTF8
    #>
    return @"
<div class="credits">
  <strong>$Script:OMMigrateProduct</strong> v$Script:OMMigrateVersion<br>
  <em>"Automating the Outlook migration Google suggested couldn't be automated."</em><br><br>
  Originator &amp; Architect: <strong>$Script:OMMigrateAuthor</strong><br>
  Implementation Specialist: <strong>$Script:OMMigrateAI</strong><br>
  Inception: $Script:OMMigrateInception
</div>
</body>
</html>
"@
}


function Get-HtmlStatGrid {
    <#
    .SYNOPSIS
        Returns an HTML stat grid (large number cards) for report summaries.

    .DESCRIPTION
        Generates the summary statistics block shown at the top of
        migration reports. Each stat shows a large number with a label.

    .PARAMETER Stats
        Ordered hashtable of label -> value pairs.

    .OUTPUTS
        [string] -- HTML string for the stat grid.

    .EXAMPLE
        $stats = [ordered]@{
            'Total Accounts'  = 26
            'To Migrate'      = 13
            'Already IMAP'    = 11
            'Exchange (Skip)' = 2
        }
        $html = Get-HtmlStatGrid -Stats $stats
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Specialized.OrderedDictionary]$Stats
    )

    $cards = ''
    foreach ($key in $Stats.Keys) {
        $cards += @"
    <div class="stat-card">
      <div class="stat-number">$($Stats[$key])</div>
      <div class="stat-label">$key</div>
    </div>
"@
    }

    return "<div class='stat-grid'>$cards</div>"
}


# ============================================================
#  REGION: SESSION COMPLETION
# ============================================================

function Complete-OMMigrateSession {
    <#
    .SYNOPSIS
        Finalizes the OMMigrate session, writes the log footer, and
        displays a summary on the console.

    .DESCRIPTION
        Must be called at the end of every OMMigrate script, typically
        in a finally block to ensure it runs even if errors occur.

        Writes the run log footer, writes the audit SESSION_END entry,
        displays a color-coded summary on the console, and optionally
        opens the generated HTML report.

    .PARAMETER Status
        Final script status: SUCCESS | WARNING | FAILED

    .PARAMETER ReportFile
        Optional. Path to the HTML report file to open after completion.

    .EXAMPLE
        try {
            # ... script work ...
            Complete-OMMigrateSession -Status 'SUCCESS' -ReportFile $reportPath
        }
        catch {
            Write-OMMigrateLog -Message "Fatal error: $_" -Level ERROR
            Complete-OMMigrateSession -Status 'FAILED'
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('SUCCESS','WARNING','FAILED')]
        [string]$Status,

        [Parameter(Mandatory = $false)]
        [string]$ReportFile = ''
    )

    $elapsed = (Get-Date) - $Global:OMMigrate.StartTime

    # -- Write log footer --------------------------------------
    Write-RunLogFooter -Status $Status

    # -- Write audit session end -------------------------------
    Write-AuditEntry -Action 'SESSION_END' `
                     -Detail ("Status=$Status | Elapsed=$($elapsed.ToString('hh\:mm\:ss')) | " +
                              "Warnings=$($Global:OMMigrate.Counters.Warnings) | " +
                              "Errors=$($Global:OMMigrate.Counters.Errors)")

    # -- Console summary ---------------------------------------
    $statusColor = switch ($Status) {
        'SUCCESS' { 'Green'  }
        'WARNING' { 'Yellow' }
        'FAILED'  { 'Red'    }
    }

    Write-Host ''
    Write-Host ('=' * 60) -ForegroundColor $statusColor
    Write-Host "  $($Global:OMMigrate.ScriptName) -- $Status" `
               -ForegroundColor $statusColor
    Write-Host ('=' * 60) -ForegroundColor $statusColor
    Write-Host "  Elapsed  : $($elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Gray
    Write-Host "  Warnings : $($Global:OMMigrate.Counters.Warnings)" `
               -ForegroundColor $(if ($Global:OMMigrate.Counters.Warnings -gt 0) {'Yellow'} else {'Gray'})
    Write-Host "  Errors   : $($Global:OMMigrate.Counters.Errors)" `
               -ForegroundColor $(if ($Global:OMMigrate.Counters.Errors -gt 0) {'Red'} else {'Gray'})
    Write-Host "  Log File : $(Invoke-OMMigrateSanitize -Text $Global:OMMigrate.RunLogFile)" -ForegroundColor DarkGray

    if ($ReportFile -and (Test-Path $ReportFile)) {
        Write-Host "  Report   : $(Invoke-OMMigrateSanitize -Text $ReportFile)" -ForegroundColor DarkGray
    }

    Write-Host ('=' * 60) -ForegroundColor $statusColor
    Write-Host ''

    # -- Report open is handled by each calling script ---------
    # Each script opens its own report after calling this function
    # so it can apply its own OpenReport parameter and WhatIf logic.
    # Opening the report here caused duplicate browser tabs.
}


# ============================================================
#  REGION: UTILITY FUNCTIONS
# ============================================================

function Find-SublimeText {
    <#
    .SYNOPSIS
        Searches common installation locations for Sublime Text and
        returns the full path to subl.exe if found.

    .DESCRIPTION
        Checks all standard Sublime Text install paths on Windows,
        including Program Files (x64 and x86), the user AppData
        roaming folder, and the Windows PATH environment variable.
        Returns the first valid subl.exe path found, or empty string
        if Sublime Text is not installed.

        Paths checked (in order):
            Program Files\Sublime Text\subl.exe          (v4 default)
            Program Files\Sublime Text 4\subl.exe
            Program Files\Sublime Text 3\subl.exe
            Program Files (x86)\Sublime Text*\subl.exe
            AppData\Roaming\Sublime Text*\subl.exe
            Windows PATH (via where.exe)
            Wildcard search under Program Files (depth 2)

    .OUTPUTS
        [string] -- Full path to subl.exe, or empty string if not found.

    .EXAMPLE
        $sublPath = Find-SublimeText
        if ($sublPath) { Write-Host "Sublime Text found at: $sublPath" }
    #>
    [CmdletBinding()]
    param()

    # -- Common install paths ----------------------------------
    $candidates = @(
        "$env:ProgramFiles\Sublime Text\subl.exe",
        "$env:ProgramFiles\Sublime Text 4\subl.exe",
        "$env:ProgramFiles\Sublime Text 3\subl.exe",
        "$env:ProgramFiles\Sublime Text 2\subl.exe",
        "${env:ProgramFiles(x86)}\Sublime Text\subl.exe",
        "${env:ProgramFiles(x86)}\Sublime Text 4\subl.exe",
        "${env:ProgramFiles(x86)}\Sublime Text 3\subl.exe",
        "${env:ProgramFiles(x86)}\Sublime Text 2\subl.exe",
        "$env:APPDATA\Sublime Text\subl.exe",
        "$env:APPDATA\Sublime Text 4\subl.exe",
        "$env:APPDATA\Sublime Text 3\subl.exe",
        "$env:LOCALAPPDATA\Sublime Text\subl.exe",
        "$env:LOCALAPPDATA\Sublime Text 4\subl.exe"
    )

    foreach ($path in $candidates) {
        if ($path -and (Test-Path $path)) {
            Write-OMMigrateLog -Message "Sublime Text found: $path" -Level DEBUG
            return $path
        }
    }

    # -- Check Windows PATH via where.exe ----------------------
    try {
        $whereResult = (& where.exe subl 2>$null) | Select-Object -First 1
        if ($whereResult -and (Test-Path $whereResult)) {
            Write-OMMigrateLog -Message "Sublime Text found via PATH: $whereResult" `
                               -Level DEBUG
            return $whereResult
        }
    }
    catch { }

    # -- Wildcard search under Program Files (last resort) -----
    try {
        $found = Get-ChildItem -Path $env:ProgramFiles `
                               -Filter 'subl.exe' `
                               -Recurse `
                               -Depth 2 `
                               -ErrorAction SilentlyContinue |
                 Select-Object -First 1
        if ($found) {
            Write-OMMigrateLog -Message "Sublime Text found via search: $($found.FullName)" `
                               -Level DEBUG
            return $found.FullName
        }
    }
    catch { }

    Write-OMMigrateLog -Message 'Sublime Text not found on this machine.' -Level DEBUG
    return ''
}


function Open-FileInEditor {
    <#
    .SYNOPSIS
        Opens a file in Sublime Text, falling back gracefully if
        Sublime Text is not installed.

    .DESCRIPTION
        The preferred file viewer for all OMMigrate output files.
        Opens logs, CSV files, JSON manifests, and RECOVERY.txt in
        the best available editor on the current machine, with no
        filename-mangling side effects.

        Fallback chain (in order):
            1. Full path in OMMigrate_Settings.json PreferredEditor
               (if set to a valid executable path)
            2. Sublime Text auto-located by Find-SublimeText
               (works even when PreferredEditor is empty)
            3. Windows default file association via cmd /c start
               (last resort -- never calls notepad.exe directly)

        HTML reports always open in the default browser regardless
        of editor setting -- browsers render the styled report
        correctly. Text editors show raw HTML markup.

        Why not Notepad:
            Windows 11 Notepad appends _1 _2 _3 suffixes to filenames
            when a file is already open. This causes confusion with
            log and report files that are referenced by exact name.
            Sublime Text opens the same file in a new tab cleanly.

    .PARAMETER FilePath
        Full path to the file to open.

    .PARAMETER ForceEditor
        When $true, opens HTML files in the text editor rather than
        the browser. Default: $false

    .EXAMPLE
        Open-FileInEditor -FilePath 'C:\Migration\Reports\Discovery_20260521.html'
        # Opens in default browser (HTML report)

    .EXAMPLE
        Open-FileInEditor -FilePath 'C:\Migration\Logs\OMMigrate_20260521.log'
        # Opens in Sublime Text

    .EXAMPLE
        Open-FileInEditor -FilePath 'C:\Migration\Config\migration_accounts.csv'
        # Opens in Sublime Text
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $false)]
        [bool]$ForceEditor = $false
    )

    if (-not (Test-Path $FilePath)) {
        Write-OMMigrateLog -Message "Open-FileInEditor: File not found: $FilePath" `
                           -Level WARN
        return
    }

    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()

    # -- HTML reports open in browser --------------------------
    if ($extension -eq '.html' -and -not $ForceEditor) {
        Write-OMMigrateLog -Message "Opening HTML report in browser: $FilePath" `
                           -Level DEBUG
        try {
            Start-Process $FilePath
        }
        catch {
            Write-OMMigrateLog -Message "Could not open browser: $_" -Level WARN
        }
        return
    }

    # -- Resolve editor path -----------------------------------
    $editorPath = ''

    # Check settings for a user-specified override path
    $settingEditor = ''
    try {
        $settingEditor = $Global:OMMigrate.Settings.Reporting.PreferredEditor
    }
    catch { }

    if (-not [string]::IsNullOrWhiteSpace($settingEditor) -and
        (Test-Path $settingEditor)) {
        # User supplied a valid full path to a preferred editor in settings
        $editorPath = $settingEditor
        Write-OMMigrateLog -Message "Using preferred editor from settings: $editorPath" `
                           -Level DEBUG
    }
    # If PreferredEditor is empty, skip Sublime auto-detection entirely
    # and fall through to Windows default file association below.
    # Sublime auto-detection only runs when the user has explicitly set
    # PreferredEditor to a Sublime path in OMMigrate_Settings.json.

    # -- Open in configured editor -----------------------------
    if ($editorPath) {
        Write-OMMigrateLog -Message "Opening in configured editor: $FilePath" -Level DEBUG
        try {
            # --add opens file as a new tab in the existing window
            # rather than launching a new window for every file
            Start-Process -FilePath  $editorPath `
                          -ArgumentList "--add `"$FilePath`"" `
                          -ErrorAction Stop
            return
        }
        catch {
            Write-OMMigrateLog -Message "Editor launch failed: $_ -- falling back." `
                               -Level WARN
        }
    }

    # -- Fallback: Windows default file association ------------
    # Uses cmd /c start to invoke whatever the operator has set
    # as the default for this file type.
    # Notepad.exe is NEVER called explicitly by this code.
    Write-OMMigrateLog -Message (
        "Opening with Windows default app: $FilePath"
    ) -Level DEBUG

    try {
        Start-Process -FilePath    'cmd.exe' `
                      -ArgumentList "/c start `"`" `"$FilePath`"" `
                      -WindowStyle  Hidden
    }
    catch {
        Write-OMMigrateLog -Message "Could not open file with default app: $_" -Level ERROR
    }
}


function Get-SafeFileName {
    <#
    .SYNOPSIS
        Converts an email address or arbitrary string into a safe
        filename by replacing invalid characters.

    .DESCRIPTION
        Used when generating PST backup filenames, report filenames,
        and log entries that include email addresses.

    .PARAMETER InputString
        The string to sanitize.

    .PARAMETER Replacement
        Character to replace invalid filename characters with.
        Default: underscore (_)

    .OUTPUTS
        [string] -- Safe filename string.

    .EXAMPLE
        Get-SafeFileName -InputString 'user@domain.com'
        # Returns: user_domain.com
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputString,

        [Parameter(Mandatory = $false)]
        [string]$Replacement = '_'
    )

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $safe    = $InputString
    foreach ($char in $invalid) {
        $safe = $safe.Replace([string]$char, $Replacement)
    }

    # Also replace @ for email addresses
    $safe = $safe.Replace('@', $Replacement)

    return $safe
}


function Format-FileSize {
    <#
    .SYNOPSIS
        Formats a byte count as a human-readable file size string.

    .PARAMETER Bytes
        File size in bytes.

    .OUTPUTS
        [string] -- Formatted size (e.g. '1.24 GB', '847 MB', '12 KB').

    .EXAMPLE
        Format-FileSize -Bytes 1304723456
        # Returns: '1.21 GB'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [long]$Bytes
    )

    if     ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { return '{0:N0} MB' -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { return '{0:N0} KB' -f ($Bytes / 1KB) }
    else                    { return "$Bytes B" }
}


function Test-AdminElevation {
    <#
    .SYNOPSIS
        Returns $true if the current PowerShell session is elevated
        (running as Administrator).

    .DESCRIPTION
        OMMigrate scripts should NOT be run as Administrator in most
        cases -- the Outlook profile belongs to the regular user account.
        This function is used to warn when elevation is detected
        unexpectedly.

    .OUTPUTS
        [bool]
    #>
    $identity  = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [System.Security.Principal.WindowsPrincipal]$identity
    return $principal.IsInRole(
        [System.Security.Principal.WindowsBuiltInRole]::Administrator
    )
}


# ============================================================
#  REGION: EXIT HANDLING
#  Safe exit infrastructure for all OMMigrate scripts.
#  Handles planned exits (operator types EXIT at a prompt),
#  Ctrl+C interruptions, and emergency exits.
#  All paths release COM, write logs, save checkpoints,
#  and display clear recovery instructions.
# ============================================================

function Show-ExitBanner {
    <#
    .SYNOPSIS
        Displays the "How to safely exit" instructions banner.

    .DESCRIPTION
        Called at the start of every OMMigrate script that makes
        changes (Scripts 01, 02, 03) so the operator knows exactly
        how to stop safely before anything is touched.

        Script 00 (discovery only) is read-only and safe to exit
        at any time -- the banner is not shown for Script 00.

    .EXAMPLE
        Show-ExitBanner
    #>
    [CmdletBinding()]
    param()

    $border = '-' * 47

    Write-Host ''
    Write-Host "  +$($border)+" -ForegroundColor DarkYellow
    Write-Host "  |  HOW TO SAFELY EXIT AT ANY TIME               |" `
               -ForegroundColor Yellow
    Write-Host "  |$($border)|" -ForegroundColor DarkYellow
    Write-Host "  |                                               |" `
               -ForegroundColor DarkYellow
    Write-Host "  |  At any Y/N prompt  -> Type EXIT + Enter      |" `
               -ForegroundColor DarkYellow
    Write-Host "  |  Skip one account   -> Type N at the prompt   |" `
               -ForegroundColor DarkYellow
    Write-Host "  |  Emergency stop     -> Press Ctrl+C           |" `
               -ForegroundColor DarkYellow
    Write-Host "  |                                               |" `
               -ForegroundColor DarkYellow
    Write-Host "  |  EXIT and Ctrl+C both save your progress      |" `
               -ForegroundColor DarkYellow
    Write-Host "  |  and release Outlook before stopping.         |" `
               -ForegroundColor DarkYellow
    Write-Host "  |                                               |" `
               -ForegroundColor DarkYellow
    Write-Host "  |  Do NOT close this window with the X button.  |" `
               -ForegroundColor Red
    Write-Host "  |  Use Ctrl+C instead for a clean stop.         |" `
               -ForegroundColor DarkYellow
    Write-Host "  +$($border)+" -ForegroundColor DarkYellow
    Write-Host ''
}


function Register-ExitHandlers {
    <#
    .SYNOPSIS
        Registers Ctrl+C and process-exit event handlers for the
        current script session.

    .DESCRIPTION
        Must be called once at the start of every OMMigrate script
        that makes changes (Scripts 01, 02, 03). Registers two handlers:

            1. PowerShell.Exiting engine event
               Fires when the PowerShell session is terminating for
               any reason -- Ctrl+C, window close, or script completion.
               Triggers Invoke-OMMigrateEmergencyExit.

            2. ConsoleCancelEventHandler (Ctrl+C)
               Intercepts Ctrl+C before PowerShell's default handler
               to ensure clean COM release before the session ends.

        Safe to call multiple times -- re-registration is idempotent.

    .PARAMETER ScriptStep
        The step number of the currently running script (1, 2, or 3).
        Used in checkpoint and recovery file naming.

    .EXAMPLE
        Register-ExitHandlers -ScriptStep 2
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(0,4)]
        [int]$ScriptStep
    )

    # Store the step number for use by the exit handler
    $Global:OMMigrate.CurrentStep = $ScriptStep

    # -- PowerShell engine exit event --------------------------
    # Unregister any existing handler first to avoid duplicates
    Get-EventSubscriber -SourceIdentifier 'OMMigrate.Exit' `
                        -ErrorAction SilentlyContinue |
        Unregister-Event -ErrorAction SilentlyContinue

    Register-EngineEvent -SourceIdentifier 'PowerShell.Exiting' -Action {
        # Only run emergency exit if session did not complete normally
        if ($Global:OMMigrate -and
            -not $Global:OMMigrate.SessionCompletedNormally) {
            Invoke-OMMigrateEmergencyExit -Reason 'Process termination or Ctrl+C detected'
        }
    } | Out-Null

    # -- Console Ctrl+C handler ---------------------------------
    # Allows Ctrl+C to be intercepted for clean COM release
    [Console]::TreatControlCAsInput = $false

    $ctrlCHandler = {
        param($sender, $e)
        $e.Cancel = $true   # Prevent immediate hard kill
        Write-Host ''
        Write-Host '  Ctrl+C detected -- initiating safe exit...' `
                   -ForegroundColor Yellow
        Invoke-OMMigrateEmergencyExit -Reason 'Ctrl+C pressed by operator'
    }

    # Remove any previous handler before adding new one
    try {
        [Console]::remove_CancelKeyPress($Script:CtrlCHandler)
    }
    catch { }

    $Script:CtrlCHandler = $ctrlCHandler
    [Console]::add_CancelKeyPress($ctrlCHandler)

    Write-OMMigrateLog -Message "Exit handlers registered for Script Step $ScriptStep." `
                       -Level DEBUG
}


function Invoke-OMMigrateGracefulExit {
    <#
    .SYNOPSIS
        Performs a clean, controlled exit when the operator deliberately
        chooses to stop (types EXIT at a prompt).

    .DESCRIPTION
        Called when the operator types EXIT, QUIT, Q, STOP, CANCEL, or
        ABORT at any Y/N confirmation prompt. This is the planned exit
        path -- the script was between account operations, so no account
        is left in a partial state.

        Actions taken:
            1. Logs the graceful exit with reason
            2. Writes audit entry
            3. Saves current checkpoint (completed accounts so far)
            4. Releases Outlook COM session
            5. Writes run log footer with INTERRUPTED status
            6. Writes RECOVERY.txt with resume instructions
            7. Displays exit summary on console
            8. Calls exit 0 (clean exit code)

        The key difference from EmergencyExit: a graceful exit means
        the script was at a safe stopping point. No account was
        mid-operation. Resume is straightforward.

    .PARAMETER Reason
        Description of why the exit was requested.
        Included in the log and recovery file.

    .EXAMPLE
        Invoke-OMMigrateGracefulExit -Reason "Operator typed EXIT at prompt"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    Write-Host ''
    Write-Host ('-' * 60) -ForegroundColor Yellow
    Write-Host '  GRACEFUL EXIT INITIATED' -ForegroundColor Yellow
    Write-Host ('-' * 60) -ForegroundColor Yellow
    Write-Host "  Reason : $Reason" -ForegroundColor Gray
    Write-Host ''

    # Mark session as intentionally stopped (suppresses emergency handler)
    if ($Global:OMMigrate) {
        $Global:OMMigrate.SessionCompletedNormally = $true
    }

    Write-OMMigrateLog -Message "GRACEFUL EXIT: $Reason" -Level WARN
    Write-AuditEntry  -Action 'GRACEFUL_EXIT' `
                      -Detail $Reason `
                      -Outcome 'SKIPPED'

    # Save checkpoint for resume
    Save-OMMigrateCheckpoint -ExitType 'GRACEFUL' -Reason $Reason

    # Release Outlook COM if active
    Invoke-OutlookCOMRelease

    # Write log footer
    Write-RunLogFooter -Status 'WARNING'

    # Write recovery file
    Write-RecoveryFile -ExitType 'GRACEFUL' -Reason $Reason

    # Console summary
    Write-Host '  Progress has been saved.' -ForegroundColor Green
    Write-Host '  Outlook has been released.' -ForegroundColor Green
    Write-Host ''
    Write-Host '  To resume from where you stopped:' -ForegroundColor Cyan
    Write-Host "  Re-run the same script -- it will skip completed accounts." `
               -ForegroundColor Cyan
    Write-Host ''
    Write-Host "  Recovery details saved to:" -ForegroundColor Gray
    if ($Global:OMMigrate -and $Global:OMMigrate.BasePath) {
        Write-Host "  $($Global:OMMigrate.BasePath)\RECOVERY.txt" -ForegroundColor Gray
    }
    Write-Host ('-' * 60) -ForegroundColor Yellow
    Write-Host ''

    exit 0
}


function Invoke-OMMigrateEmergencyExit {
    <#
    .SYNOPSIS
        Performs emergency cleanup when the script is interrupted
        unexpectedly (Ctrl+C, process termination, or fatal error).

    .DESCRIPTION
        Called by the Ctrl+C handler and PowerShell.Exiting engine
        event when the script is interrupted outside of a normal
        Y/N prompt. The script may have been mid-operation on an
        account, so this path includes additional warnings and a
        more detailed recovery file.

        Actions taken:
            1. Saves checkpoint immediately (whatever completed so far)
            2. Releases Outlook COM session (best-effort)
            3. Writes INTERRUPTED status to log
            4. Writes audit entry
            5. Writes RECOVERY.txt with detailed recovery steps
            6. Warns about possible partial account state
            7. Displays console recovery instructions
            8. Calls exit 1 (non-zero = interrupted)

        The recovery file includes a warning to verify Outlook manually
        before resuming, since the script may have been mid-account.

    .PARAMETER Reason
        Description of why the emergency exit was triggered.

    .EXAMPLE
        Invoke-OMMigrateEmergencyExit -Reason 'Ctrl+C pressed by operator'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Reason = 'Unexpected interruption'
    )

    # Guard against recursive calls
    if ($Global:OMMigrate -and $Global:OMMigrate.EmergencyExitInProgress) { return }
    if ($Global:OMMigrate) { $Global:OMMigrate.EmergencyExitInProgress = $true }

    Write-Host ''
    Write-Host ('!' * 60) -ForegroundColor Red
    Write-Host '  EMERGENCY EXIT -- SAVING STATE AND RELEASING OUTLOOK' `
               -ForegroundColor Red
    Write-Host ('!' * 60) -ForegroundColor Red
    Write-Host ''

    # Save checkpoint immediately -- capture whatever completed
    try { Save-OMMigrateCheckpoint -ExitType 'EMERGENCY' -Reason $Reason } catch { }

    # Release Outlook COM -- best effort, wrapped in try/catch
    try { Invoke-OutlookCOMRelease } catch { }

    # Write log entries -- best effort
    try {
        Write-OMMigrateLog -Message "EMERGENCY EXIT: $Reason" -Level ERROR
        Write-AuditEntry  -Action 'EMERGENCY_EXIT' `
                          -Detail $Reason `
                          -Outcome 'FAILED'
        Write-RunLogFooter -Status 'FAILED'
    }
    catch { }

    # Write recovery file
    try { Write-RecoveryFile -ExitType 'EMERGENCY' -Reason $Reason } catch { }

    # Console recovery instructions
    Write-Host '  IMMEDIATE STEPS:' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  1. Open Task Manager (Ctrl+Shift+Esc)' -ForegroundColor White
    Write-Host '     Verify OUTLOOK.EXE is NOT running.' -ForegroundColor Gray
    Write-Host '     If it is: Right-click -> End Task' -ForegroundColor Gray
    Write-Host ''
    Write-Host '  2. Open Outlook manually and check your accounts.' -ForegroundColor White
    Write-Host '     Verify no account is missing or misconfigured.' -ForegroundColor Gray
    Write-Host ''
    Write-Host '  3. Check RECOVERY.txt in your migration folder.' -ForegroundColor White
    Write-Host '     It lists which accounts completed before the exit.' -ForegroundColor Gray
    Write-Host ''
    Write-Host '  4. When ready to resume, re-run the same script.' -ForegroundColor White
    Write-Host '     Completed accounts will be skipped automatically.' -ForegroundColor Gray
    Write-Host ''

    if ($Global:OMMigrate -and $Global:OMMigrate.BasePath) {
        Write-Host "  Recovery file: $($Global:OMMigrate.BasePath)\RECOVERY.txt" `
                   -ForegroundColor DarkGray
    }

    Write-Host ''
    Write-Host ('!' * 60) -ForegroundColor Red
    Write-Host ''

    exit 1
}


function Save-OMMigrateCheckpoint {
    <#
    .SYNOPSIS
        Saves current migration progress to a checkpoint file so
        the script can resume from where it left off.

    .DESCRIPTION
        Writes a checkpoint JSON file that records which accounts
        have been successfully processed and which are still pending.
        The checkpoint is read at the start of each script run to
        skip accounts that already completed.

        Checkpoint files are named by script step:
            Manifests\Step02_Checkpoint.json
            Manifests\Step03_Checkpoint.json

        Script 00 and 01 do not use checkpoints -- Script 00 is
        read-only and Script 01 (backup) is safe to re-run in full.

    .PARAMETER ExitType
        'GRACEFUL' or 'EMERGENCY' -- recorded in the checkpoint.

    .PARAMETER Reason
        The exit reason -- recorded for operator reference.

    .EXAMPLE
        Save-OMMigrateCheckpoint -ExitType 'GRACEFUL' -Reason 'Operator exit'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('GRACEFUL','EMERGENCY','NORMAL')]
        [string]$ExitType,

        [Parameter(Mandatory = $false)]
        [string]$Reason = ''
    )

    if (-not $Global:OMMigrate -or -not $Global:OMMigrate.ManifestPath) { return }

    # Never write checkpoints in WhatIf/Preview mode -- preview runs must not
    # poison the checkpoint file with simulated progress data.
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message 'WhatIf: Skipping checkpoint write -- preview mode.' `
                           -Level DEBUG
        return
    }

    $step = if ($Global:OMMigrate.CurrentStep) {
        $Global:OMMigrate.CurrentStep
    } else { 0 }

    # Only checkpoint steps 2 and 3 -- others are safe to re-run fully
    if ($step -lt 2) { return }

    $checkpointFile = Join-Path $Global:OMMigrate.ManifestPath `
                                "Step0${step}_Checkpoint.json"

    $checkpoint = [ordered]@{
        Product            = $Script:OMMigrateProduct
        Version            = $Script:OMMigrateVersion
        Step               = $step
        ScriptName         = $Global:OMMigrate.ScriptName
        SavedAt            = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
        ExitType           = $ExitType
        Reason             = $Reason
        SessionID          = $Global:OMMigrate.SessionID
        Operator           = $Global:OMMigrate.RunUser
        Machine            = $Global:OMMigrate.MachineName
        CompletedAccounts  = @()
        PendingAccounts    = @()
        InterruptedAccount = ''
    }

    # Do not write checkpoint if no work was started this run.
    # Check both the Progress tracker and whether pending was ever set.
    # This prevents writing a checkpoint when operator declines at pre-flight.
    $progressInitialized = $Global:OMMigrate.Contains('Progress') -and
                           $Global:OMMigrate.Progress -and
                           $Global:OMMigrate.Progress.Pending.Count -gt 0

    $hasCompleted = $progressInitialized -and
                    @($Global:OMMigrate.Progress.Completed).Count -gt 0

    if (-not $progressInitialized -and -not $hasCompleted) {
        Write-OMMigrateLog -Message 'No migration progress to checkpoint -- skipping checkpoint write.' `
                           -Level INFO
        return
    }

    # Populate checkpoint from progress tracker
    if ($progressInitialized -or $hasCompleted) {
        $checkpoint.CompletedAccounts  = $Global:OMMigrate.Progress.Completed
        $checkpoint.PendingAccounts    = $Global:OMMigrate.Progress.Pending
        $checkpoint.InterruptedAccount = $Global:OMMigrate.Progress.CurrentAccount
    }

    try {
        $checkpoint | ConvertTo-Json -Depth 5 |
            Set-Content -Path $checkpointFile -Encoding UTF8
        Write-Host "  Checkpoint saved: $(Invoke-OMMigrateSanitize -Text $checkpointFile)" -ForegroundColor Green
    }
    catch {
        Write-Host "  WARNING: Could not save checkpoint: $_" -ForegroundColor Yellow
    }
}


function Read-OMMigrateCheckpoint {
    <#
    .SYNOPSIS
        Reads the checkpoint file for a given script step and returns
        the list of already-completed accounts.

    .DESCRIPTION
        Called at the start of Scripts 02 and 03 to determine if a
        previous run was interrupted. If a checkpoint exists, accounts
        listed in CompletedAccounts are skipped in the current run.

        If no checkpoint exists, returns an empty completed list --
        the script runs all accounts normally.

    .PARAMETER Step
        Script step number (2 or 3).

    .OUTPUTS
        PSCustomObject with:
            HasCheckpoint      [bool]    Whether a checkpoint was found
            CompletedAccounts  [array]   Email addresses already processed
            InterruptedAccount [string]  Account that was mid-process (if any)
            SavedAt            [string]  Timestamp of the checkpoint
            ExitType           [string]  GRACEFUL or EMERGENCY

    .EXAMPLE
        $checkpoint = Read-OMMigrateCheckpoint -Step 2
        if ($checkpoint.HasCheckpoint) {
            Write-Host "Resuming -- $($checkpoint.CompletedAccounts.Count) accounts already done"
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(2,4)]
        [int]$Step
    )

    $empty = [PSCustomObject]@{
        HasCheckpoint      = $false
        CompletedAccounts  = @()
        InterruptedAccount = ''
        SavedAt            = ''
        ExitType           = ''
    }

    if (-not $Global:OMMigrate -or -not $Global:OMMigrate.ManifestPath) {
        return $empty
    }

    $checkpointFile = Join-Path $Global:OMMigrate.ManifestPath `
                                "Step0${Step}_Checkpoint.json"

    if (-not (Test-Path $checkpointFile)) { return $empty }

    try {
        $data = Get-Content $checkpointFile -Raw | ConvertFrom-Json

        # Validate checkpoint is for this machine
        if ($data.Machine -ne $Global:OMMigrate.MachineName) {
            Write-OMMigrateLog -Message (
                "Checkpoint found but was created on a different machine " +
                "($($data.Machine)). Ignoring checkpoint."
            ) -Level WARN
            return $empty
        }

        Write-OMMigrateLog -Message (
            "Checkpoint found for Step $Step. " +
            "Saved: $($data.SavedAt) | ExitType: $($data.ExitType) | " +
            "Completed accounts: $($data.CompletedAccounts.Count)"
        ) -Level WARN

        # Display resume notice on console
        Write-Host ''
        Write-Host ('-' * 60) -ForegroundColor Yellow
        Write-Host '  PREVIOUS RUN CHECKPOINT DETECTED' -ForegroundColor Yellow
        Write-Host ('-' * 60) -ForegroundColor Yellow
        Write-Host "  Saved    : $($data.SavedAt)" -ForegroundColor Gray
        Write-Host "  Exit type: $($data.ExitType)" -ForegroundColor Gray
        Write-Host "  Accounts already completed: $($data.CompletedAccounts.Count)" `
                   -ForegroundColor Green

        if ($data.CompletedAccounts.Count -gt 0) {
            foreach ($email in $data.CompletedAccounts) {
                Write-Host "    ? $(Invoke-OMMigrateSanitize -Text $email)" -ForegroundColor DarkGreen
            }
        }

        if ($data.InterruptedAccount) {
            Write-Host "  Interrupted during: $(Invoke-OMMigrateSanitize -Text $data.InterruptedAccount)" `
                       -ForegroundColor Yellow
            Write-Host "  Verify this account manually before resuming." `
                       -ForegroundColor Yellow
        }

        Write-Host ('-' * 60) -ForegroundColor Yellow
        Write-Host ''

        return [PSCustomObject]@{
            HasCheckpoint      = $true
            CompletedAccounts  = $data.CompletedAccounts
            InterruptedAccount = $data.InterruptedAccount
            SavedAt            = $data.SavedAt
            ExitType           = $data.ExitType
        }
    }
    catch {
        Write-OMMigrateLog -Message "Could not read checkpoint file: $_" -Level WARN
        return $empty
    }
}


function Update-OMMigrateProgress {
    <#
    .SYNOPSIS
        Updates the in-memory progress tracker used by checkpointing.

    .DESCRIPTION
        Called by Scripts 02 and 03 after each account is processed
        to keep the progress state current. If the script is
        interrupted, Save-OMMigrateCheckpoint reads this state.

        Usage pattern:
            Before processing account:
                Update-OMMigrateProgress -SetCurrent 'user@domain.com'
            After account completes successfully:
                Update-OMMigrateProgress -MarkComplete 'user@domain.com'

    .PARAMETER SetCurrent
        Email address of the account now being processed.

    .PARAMETER MarkComplete
        Email address of the account that just completed successfully.

    .PARAMETER SetPending
        Array of all pending account email addresses (set once at start).

    .EXAMPLE
        # Initialize at script start
        Update-OMMigrateProgress -SetPending $accountsToMigrate

        # Before each account
        Update-OMMigrateProgress -SetCurrent $account.EmailAddress

        # After each account succeeds
        Update-OMMigrateProgress -MarkComplete $account.EmailAddress
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$SetCurrent = '',

        [Parameter(Mandatory = $false)]
        [string]$MarkComplete = '',

        [Parameter(Mandatory = $false)]
        [string[]]$SetPending = @()
    )

    # Initialize progress tracker if not present
    if (-not $Global:OMMigrate.Contains('Progress') -or -not $Global:OMMigrate.Progress) {
        $Global:OMMigrate.Progress = [ordered]@{
            Completed      = [System.Collections.Generic.List[string]]::new()
            Pending        = [System.Collections.Generic.List[string]]::new()
            CurrentAccount = ''
        }
    }

    if ($SetPending.Count -gt 0) {
        $Global:OMMigrate.Progress.Pending.Clear()
        foreach ($email in $SetPending) {
            [void]$Global:OMMigrate.Progress.Pending.Add($email)
        }
    }

    if ($SetCurrent) {
        $Global:OMMigrate.Progress.CurrentAccount = $SetCurrent
    }

    if ($MarkComplete) {
        if (-not $Global:OMMigrate.Progress.Completed.Contains($MarkComplete)) {
            [void]$Global:OMMigrate.Progress.Completed.Add($MarkComplete)
        }
        [void]$Global:OMMigrate.Progress.Pending.Remove($MarkComplete)
        $Global:OMMigrate.Progress.CurrentAccount = ''
    }
}


function Write-RecoveryFile {
    <#
    .SYNOPSIS
        Writes a plain-text RECOVERY.txt file with clear instructions
        for the operator on how to assess and resume after an exit.

    .DESCRIPTION
        Written on every exit (graceful or emergency). Contains:
            - What happened and when
            - Which accounts completed before the exit
            - Which account was interrupted (if emergency)
            - Step-by-step instructions to verify and resume
            - Log file and checkpoint file locations

        Written in plain English -- no technical jargon. Designed
        to be readable by anyone, not just the script operator.

    .PARAMETER ExitType
        'GRACEFUL' or 'EMERGENCY'

    .PARAMETER Reason
        The exit reason string.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('GRACEFUL','EMERGENCY')]
        [string]$ExitType,

        [Parameter(Mandatory = $false)]
        [string]$Reason = ''
    )

    if (-not $Global:OMMigrate -or -not $Global:OMMigrate.BasePath) { return }

    $recoveryFile = Join-Path $Global:OMMigrate.BasePath 'RECOVERY.txt'
    $timestamp    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $step         = $Global:OMMigrate.CurrentStep
    $scriptName   = $Global:OMMigrate.ScriptName

    # Build completed/pending account lists
    $completedList = ''
    $pendingList   = ''
    $interrupted   = ''

    if ($Global:OMMigrate.Contains('Progress') -and $Global:OMMigrate.Progress) {

        # Read completed and pending account lists from the Progress tracker.
        # These variables were previously referenced but never assigned --
        # fixed to read directly from Global:OMMigrate.Progress.
        $completedAccounts = $Global:OMMigrate.Progress.Completed
        $pendingAccounts   = $Global:OMMigrate.Progress.Pending

        if ($completedAccounts.Count -gt 0) {
            $completedList = ($completedAccounts | ForEach-Object { "    ? $_" }) -join "`r`n"
        }
        else {
            $completedList = '    (none -- exit occurred before any account completed)'
        }

        if ($pendingAccounts.Count -gt 0) {
            $pendingList = ($pendingAccounts | ForEach-Object { "    ? $_" }) -join "`r`n"
        }
        else {
            $pendingList = '    (none remaining)'
        }
    }

    $urgencyBlock = if ($ExitType -eq 'EMERGENCY') {
        @"
! IMPORTANT -- VERIFY OUTLOOK BEFORE RESUMING !
===============================================
Because the script was interrupted mid-run, you must verify
Outlook is in a consistent state before resuming.

IMMEDIATE CHECKS:
  1. Open Task Manager (Ctrl+Shift+Esc -> Details tab)
     Look for OUTLOOK.EXE in the process list.
     If found: right-click -> End Task. Wait for it to disappear.

  2. Open Outlook normally (double-click desktop icon).
     Check the left-hand account list.
     Verify all expected accounts are present and connected.

  3. If an account appears to be missing or broken, DO NOT PANIC.
     Your PST backup files are safe in the Backups folder.
     Contact support or re-run Script 01 to re-verify backups.

"@
    } else {
        @"
EXIT TYPE: Planned (operator requested)
The script was at a safe stopping point between account operations.
No account was left in a partial state.

"@
    }

    $resumeStep = if ($step -ge 2) {
        @"
HOW TO RESUME
=============
Re-run the same script that was interrupted:
  $scriptName.ps1

The script will automatically detect the checkpoint file and
skip accounts that already completed. You do not need to start
from scratch.

If you want to start fully fresh (ignore the checkpoint):
  Delete: $($Global:OMMigrate.ManifestPath)\Step0${step}_Checkpoint.json
  Then re-run the script.

"@
    } else {
        @"
HOW TO RESUME
=============
Script 00 (Discovery) and Script 01 (Backup) are safe to re-run
from the beginning at any time -- they do not make irreversible changes.
Simply re-run the script.

"@
    }

    $content = @"
============================================================
  OMMigrate -- RECOVERY INSTRUCTIONS
============================================================
  Product  : $Script:OMMigrateProduct v$Script:OMMigrateVersion
  Architect: $Script:OMMigrateAuthor
  Generated: $timestamp
  Script   : $scriptName
  Exit Type: $ExitType
  Reason   : $Reason
============================================================

$urgencyBlock
ACCOUNTS COMPLETED BEFORE EXIT
===============================
$completedList

ACCOUNTS STILL PENDING
=======================
$pendingList
$(if ($interrupted) {
"
ACCOUNT INTERRUPTED (verify this one manually)
===============================================
  ! $interrupted
    This account was being processed when the exit occurred.
    Open Outlook and verify this account is connected correctly.
"
})

$resumeStep
FILE LOCATIONS
==============
  Working folder : $($Global:OMMigrate.BasePath)
  Run log        : $($Global:OMMigrate.RunLogFile)
  Audit log      : $($Global:OMMigrate.AuditLogFile)
  Backups folder : $($Global:OMMigrate.BackupPath)
$(if ($step -ge 2) {
"  Checkpoint     : $($Global:OMMigrate.ManifestPath)\Step0${step}_Checkpoint.json"
})

============================================================
  "Automating the Outlook migration Google suggested couldn't be automated."
============================================================
"@

    try {
        Set-Content -Path $recoveryFile -Value $content -Encoding UTF8
    }
    catch {
        # If we can't write RECOVERY.txt the situation is already bad --
        # silently continue so the exit process completes
    }
}


function Invoke-OutlookCOMRelease {
    <#
    .SYNOPSIS
        Best-effort Outlook COM release called from exit handlers.

    .DESCRIPTION
        Attempts to call Release-OutlookCOM from OMMigrate-Outlook.psm1
        if that module is loaded. Safe to call even if the Outlook
        module was never imported -- handles the missing function gracefully.

        Used by both graceful and emergency exit paths to ensure
        Outlook doesn't remain as a ghost process.
    #>
    [CmdletBinding()]
    param()

    try {
        # Check if the Outlook module function is available
        $releaseFunc = Get-Command 'Release-OutlookCOM' -ErrorAction SilentlyContinue
        if ($releaseFunc) {
            # Only release if Outlook is actually running -- prevents double
            # release when the script's finally block already released it.
            $outlookRunning = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
            if ($outlookRunning) {
                Write-Host '  Releasing Outlook COM session...' -ForegroundColor Gray
                Release-OutlookCOM
                Write-Host '  Outlook released successfully.' -ForegroundColor Green
            }
        }
        else {
            # Module not loaded -- try to kill any Outlook process directly
            $outlookProc = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
            if ($outlookProc) {
                Write-Host '  Warning: Outlook COM module not loaded.' -ForegroundColor Yellow
                Write-Host '  Attempting to stop Outlook process...' -ForegroundColor Yellow
                Stop-Process -Name 'OUTLOOK' -Force -ErrorAction SilentlyContinue
                Start-Sleep -Milliseconds 1000
                Write-Host '  Outlook process stopped.' -ForegroundColor Gray
            }
        }
    }
    catch {
        Write-Host "  Warning: Could not release Outlook cleanly: $_" -ForegroundColor Yellow
        Write-Host '  Please verify Outlook.exe is not running in Task Manager.' `
                   -ForegroundColor Yellow
    }
}


# ============================================================
#  REGION: SANITIZATION ENGINE
#  Masks sensitive data in console output and log files when
#  -Sanitize is active. Real values are always used for all
#  actual migration operations -- sanitization is display-only.
#
#  Standard Outlook system folders are never masked:
#    Inbox, Sent Items, Drafts, Deleted Items, Outbox,
#    Junk Email, Calendar, Contacts, Tasks, Notes, Journal,
#    Conversation History, RSS Feeds, Archive, Sync Issues,
#    Clutter, Suggested Contacts, Quick Step Settings
# ============================================================

# System folders that are never masked (not sensitive)
$Script:OMMigrateSysFolders = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
@(
    'Inbox', 'Sent Items', 'Drafts', 'Deleted Items', 'Outbox',
    'Junk Email', 'Calendar', 'Contacts', 'Tasks', 'Notes',
    'Journal', 'Conversation History', 'RSS Feeds', 'Archive',
    'Sync Issues', 'Clutter', 'Suggested Contacts',
    'Quick Step Settings', 'Recoverable Items', 'Purges',
    'Versions', 'DiscoveryHolds', 'SubstrateHolds', 'Audits',
    'Calendar Logging', 'Local Failures', 'Server Failures',
    'Conflicts', 'Unread Mail', 'Sent Mail', 'Spam',
    'Important', 'Starred', 'All Mail', 'Chats'
) | ForEach-Object { [void]$Script:OMMigrateSysFolders.Add($_) }


function Invoke-OMMigrateSanitize {
    <#
    .SYNOPSIS
        Replaces all registered sensitive terms in a string with aliases.

    .DESCRIPTION
        Applies the session sanitization map to the given text, replacing
        every registered real value with its consistent alias. Used by
        Write-OMMigrateLog and the per-script Write-Host override.

        No-ops if -Sanitize was not passed or the map is empty.

    .PARAMETER Text
        The string to sanitize.

    .OUTPUTS
        [string] -- sanitized string with all registered terms replaced.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    if (-not $Global:OMMigrate.Sanitize) { return $Text }
    if (-not $Global:OMMigrate.SanitizeMap -or
        $Global:OMMigrate.SanitizeMap.Count -le 1) { return $Text }

    # Longer terms first so partial matches don't preempt full matches
    # Skip the internal __counters__ key which is not a sanitize term
    $sortedKeys = $Global:OMMigrate.SanitizeMap.Keys |
                  Where-Object { $_ -ne '__counters__' } |
                  Sort-Object { $_.Length } -Descending

    foreach ($real in $sortedKeys) {
        if ($Text -like "*$real*") {
            $alias = $Global:OMMigrate.SanitizeMap[$real]
            $Text  = $Text.Replace($real, $alias)
        }
    }

    return $Text
}


function Register-SanitizeTerms {
    <#
    .SYNOPSIS
        Registers one or more sensitive terms into the sanitization map.

    .DESCRIPTION
        Each term is assigned a consistent alias for the session.
        If the term is already registered the existing alias is kept
        (idempotent). Does nothing when -Sanitize is not active.

    .PARAMETER Terms
        Array of sensitive strings to register.

    .PARAMETER Prefix
        Alias prefix that controls the alias format:
            Email   -> account01@domain01.com
            Server  -> server01.domain01.com  (or smtpserver01.domain01.com)
            Display -> Account 01
            Path    -> [USER]  or  folder01
            Store   -> Store01
            Folder  -> Folder01
        Default: Generic (alias01, alias02, ...)

    .PARAMETER Category
        Logical grouping for counter isolation:
            Email | Server | Display | Path | Store | Folder | Generic
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Terms,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Email','Server','SmtpServer','Display','Path',
                     'Store','Folder','Generic')]
        [string]$Category = 'Generic'
    )

    if (-not $Global:OMMigrate.Sanitize) { return }
    if (-not $Global:OMMigrate.SanitizeMap) {
        $Global:OMMigrate.SanitizeMap = [ordered]@{}
    }

    # Per-category counters stored in a nested hashtable on the map
    if (-not $Global:OMMigrate.SanitizeMap.Contains('__counters__')) {
        $Global:OMMigrate.SanitizeMap['__counters__'] = @{
            Email      = 0
            Server     = 0
            SmtpServer = 0
            Display    = 0
            Path       = 0
            Store      = 0
            Folder     = 0
            Generic    = 0
        }
    }

    $counters = $Global:OMMigrate.SanitizeMap['__counters__']

    foreach ($term in $Terms) {
        if ([string]::IsNullOrWhiteSpace($term)) { continue }
        # Skip if already registered (idempotent)
        if ($Global:OMMigrate.SanitizeMap.Contains($term)) { continue }

        $counters[$Category]++
        $n = $counters[$Category].ToString('D2')

        $alias = switch ($Category) {
            'Email'      {
                # account01@domain01.com
                # Derive a domain alias from the domain portion
                $parts   = $term -split '@'
                $domAlias = "domain$n.com"
                "account$n@$domAlias"
            }
            'Server'     { "mailserver$n.domain$n.com"  }
            'SmtpServer' { "smtpserver$n.domain$n.com"  }
            'Display'    { "Account $n"                  }
            'Path'       { "[USER$n]"                    }
            'Store'      { "Store$n"                     }
            'Folder'     { "Folder$n"                    }
            default      { "alias$n"                     }
        }

        $Global:OMMigrate.SanitizeMap[$term] = $alias
    }
}


function Initialize-SanitizeMap {
    <#
    .SYNOPSIS
        Builds the full sanitization map from account and folder data.

    .DESCRIPTION
        Called once per script after account data is loaded. Registers:
          - Email addresses        (Email category)
          - Display names          (Display category)
          - Incoming server names  (Server category)
          - Outgoing server names  (SmtpServer category)
          - PST / OST file paths   (path masking)
          - Username portion of Windows paths ([USER])
          - Store names            (Store category)
          - Custom folder names    (Folder category -- system folders excluded)

        Accepts either an array of account objects (PSCustomObjects from
        Import-Csv or the registry scan) or a flat array of folder name
        strings.

        Safe to call with $null or empty arrays -- does nothing.
        No-ops when -Sanitize is not active.

    .PARAMETER Accounts
        Array of account objects. May be $null.

    .PARAMETER Folders
        Array of folder path strings or PSCustomObjects with a FolderPath
        or FolderName property. May be $null.

    .PARAMETER Rules
        Array of rule objects with RuleName property. May be $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object[]]$Accounts = @(),

        [Parameter(Mandatory = $false)]
        [object[]]$Folders = @(),

        [Parameter(Mandatory = $false)]
        [object[]]$Rules = @()
    )

    if (-not $Global:OMMigrate.Sanitize) { return }

    # -- Process accounts first (emails must be registered before username
    #    so longer email strings are substituted before shorter username substrings)
    if ($Accounts -and $Accounts.Count -gt 0) {

        # Collect all values before registering so counters are assigned
        # in a predictable order (email -> server -> display -> paths)
        $emails   = [System.Collections.Generic.List[string]]::new()
        $servers  = [System.Collections.Generic.List[string]]::new()
        $smtps    = [System.Collections.Generic.List[string]]::new()
        $displays = [System.Collections.Generic.List[string]]::new()
        $paths    = [System.Collections.Generic.List[string]]::new()
        $stores   = [System.Collections.Generic.List[string]]::new()

        foreach ($acct in $Accounts) {
            # Email address
            $email = $null
            try { $email = $acct.EmailAddress } catch { }
            if (-not $email) { try { $email = $acct.Email } catch { } }
            if ($email -and $email.Trim()) { [void]$emails.Add($email.Trim()) }

            # Incoming server
            $srv = $null
            try { $srv = $acct.IncomingServer } catch { }
            if ($srv -and $srv.Trim()) { [void]$servers.Add($srv.Trim()) }

            # Outgoing / SMTP server
            $smtp = $null
            try { $smtp = $acct.OutgoingServer } catch { }
            if ($smtp -and $smtp.Trim()) { [void]$smtps.Add($smtp.Trim()) }

            # Display name
            $disp = $null
            try { $disp = $acct.DisplayName } catch { }
            if ($disp -and $disp.Trim()) { [void]$displays.Add($disp.Trim()) }

            # PST path
            $pst = $null
            try { $pst = $acct.PSTPath } catch { }
            if ($pst -and $pst.Trim()) { [void]$paths.Add($pst.Trim()) }

            # OST path
            $ost = $null
            try { $ost = $acct.OSTPath } catch { }
            if ($ost -and $ost.Trim()) { [void]$paths.Add($ost.Trim()) }

            # Store / data file name -- often contains email address for Exchange
            $store = $null
            try { $store = $acct.DataFileName } catch { }
            if ($store -and $store.Trim()) {
                [void]$stores.Add($store.Trim())
                # If store name looks like an email address, also register as email
                if ($store -match '@') { [void]$emails.Add($store.Trim()) }
            }

            # StoreName property (alternate name used by some account types)
            $storeName = $null
            try { $storeName = $acct.StoreName } catch { }
            if ($storeName -and $storeName.Trim() -and $storeName -match '@') {
                [void]$emails.Add($storeName.Trim())
            }
        }

        if ($emails.Count   -gt 0) { Register-SanitizeTerms -Terms $emails.ToArray()   -Category 'Email'      }
        if ($servers.Count  -gt 0) { Register-SanitizeTerms -Terms $servers.ToArray()  -Category 'Server'     }
        if ($smtps.Count    -gt 0) { Register-SanitizeTerms -Terms $smtps.ToArray()    -Category 'SmtpServer' }
        if ($displays.Count -gt 0) { Register-SanitizeTerms -Terms $displays.ToArray() -Category 'Display'    }
        if ($paths.Count    -gt 0) { Register-SanitizeTerms -Terms $paths.ToArray()    -Category 'Path'       }
        if ($stores.Count   -gt 0) { Register-SanitizeTerms -Terms $stores.ToArray()   -Category 'Store'      }
    }

    # -- Always mask the current Windows username ---------------
    # Registered AFTER emails so longer email strings take priority
    # over username substrings during substitution
    if ($env:USERNAME) {
        Register-SanitizeTerms -Terms @($env:USERNAME) -Category 'Path'
        # Also mask the full DOMAIN\user form
        if ($env:USERDOMAIN) {
            Register-SanitizeTerms -Terms @("$env:USERDOMAIN\$env:USERNAME") -Category 'Path'
        }
    }

    # -- Process folder names -----------------------------------
    if ($Folders -and $Folders.Count -gt 0) {
        $customFolders = [System.Collections.Generic.List[string]]::new()

        foreach ($f in $Folders) {
            # Accept either a plain string or an object with FolderPath/FolderName
            $raw = $null
            if ($f -is [string]) {
                $raw = $f
            }
            else {
                try { $raw = $f.FolderPath } catch { }
                if (-not $raw) { try { $raw = $f.FolderName } catch { } }
            }
            if (-not $raw -or -not $raw.Trim()) { continue }

            # Split on path separators and process each segment
            $segments = $raw -split '[/\\]' | Where-Object { $_.Trim() -ne '' }
            foreach ($seg in $segments) {
                $seg = $seg.Trim()
                if (-not $Script:OMMigrateSysFolders.Contains($seg)) {
                    [void]$customFolders.Add($seg)
                }
            }
        }

        if ($customFolders.Count -gt 0) { Register-SanitizeTerms -Terms $customFolders.ToArray() -Category 'Folder' }
    }

    # -- Process rule names ------------------------------------
    if ($Rules -and $Rules.Count -gt 0) {
        $ruleNames = [System.Collections.Generic.List[string]]::new()
        foreach ($rule in $Rules) {
            $rname = $null
            try { $rname = $rule.RuleName } catch { }
            if ($rname -and $rname.Trim()) { [void]$ruleNames.Add($rname.Trim()) }
            # Also register folder target paths from rules
            $rtarget = $null
            try { $rtarget = $rule.TargetFolderPath } catch { }
            if ($rtarget -and $rtarget.Trim()) { [void]$ruleNames.Add($rtarget.Trim()) }
        }
        if ($ruleNames.Count -gt 0) { Register-SanitizeTerms -Terms $ruleNames.ToArray() -Category 'Folder' }
    }

    Write-OMMigrateLog -Message (
        "[SANITIZE] Map initialized: $($Global:OMMigrate.SanitizeMap.Count - 1) terms registered " +
        "(excluding counter block)"
    ) -Level INFO -NoConsole
}


# ============================================================
#  REGION: MODULE EXPORTS
# ============================================================

Export-ModuleMember -Function @(

    # Initialization
    'Initialize-OMMigrate'
    'Get-DefaultSettings'
    'Save-OMMigrateSelectedProfile'
    'Switch-OMMigrateProfileSettings'
    'Sync-OMMigrateProfileSettings'
    # EXTRACTED 2026-07-09 (Administrator direction) from OMMigrate-Core_WIP.psm1 --
    # see Save-OMMigrateArchiveStoreMappings function comment for context.
    'Save-OMMigrateArchiveStoreMappings'
    'Get-OMMigrateCsvPath'

    # Logging
    'Write-OMMigrateLog'
    'Write-RunLogHeader'
    'Write-RunLogFooter'

    # Audit
    'Write-AuditEntry'

    # Console output
    'Show-Banner'
    'Show-SectionHeader'
    'Show-AccountStatus'

    # User prompts
    'Confirm-Action'
    'Show-PreflightWarning'
    'Wait-UserKeypress'

    # Environment validation
    'Test-OMMigrateEnvironment'

    # Manifest management
    'Write-StepManifest'
    'Read-StepManifest'

    # HTML report helpers
    'Get-HtmlReportHeader'
    'Get-HtmlReportSection'
    'Get-HtmlReportFooter'
    'Get-HtmlStatGrid'

    # Session completion
    'Complete-OMMigrateSession'

    # Utilities
    'Get-SafeFileName'
    'Format-FileSize'
    'Test-AdminElevation'
    'Find-SublimeText'
    'Open-FileInEditor'

    # Exit handling -- NEW
    'Show-ExitBanner'
    'Register-ExitHandlers'
    'Invoke-OMMigrateGracefulExit'
    'Invoke-OMMigrateEmergencyExit'
    'Save-OMMigrateCheckpoint'
    'Read-OMMigrateCheckpoint'
    'Update-OMMigrateProgress'
    'Write-RecoveryFile'
    'Invoke-OutlookCOMRelease'

    # Sanitization engine
    'Invoke-OMMigrateSanitize'
    'Register-SanitizeTerms'
    'Initialize-SanitizeMap'
)
# ***** END OF FILE *****
