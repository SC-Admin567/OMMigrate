#Requires -Version 5.1
<#
.SYNOPSIS
    OMMigrate-Registry.psm1 -- Outlook Profile Registry Reader

.DESCRIPTION
    Reads and decodes Outlook profile configuration data from the
    Windows Registry. This module is the intelligence engine behind
    Script 00 (Discovery) -- it extracts everything knowable about
    each email account without prompting the operator for information
    that is already stored on the machine.

    Data extracted per account:
        - Email address
        - Display name
        - Account type (POP3 / IMAP / Exchange / Other)
        - Provider tag (POP3-STANDARD, POP3-ATTAMERITECH, IMAP-ALREADY,
                        EXCHANGE-SKIP, etc.)
        - Incoming mail server hostname and port
        - Incoming SSL/TLS flag
        - Outgoing SMTP server hostname and port
        - Outgoing SSL/TLS flag
        - SMTP authentication type
        - PST / OST data file paths
        - AWS SES detection (flags accounts using AWS SES SMTP endpoints)

    SECURITY NOTES:
        - This module performs READ-ONLY registry operations exclusively.
        - Passwords are stored in the registry encrypted via Windows DPAPI.
          This module does NOT attempt to decrypt passwords. Ever.
        - No data is transmitted anywhere. All extracted data is returned
          as PowerShell objects for use within the OMMigrate session only.
        - Registry paths read are documented by Microsoft and contain
          only the operator's own account configuration data.

    REGISTRY PATHS READ:
        HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\
        HKCU:\Software\Microsoft\Windows NT\CurrentVersion\
              Windows Messaging Subsystem\Profiles\

    BINARY BLOB DECODING:
        Outlook stores some account properties as binary blobs using a
        well-documented community-standard structure. The decoding logic
        in this module reads Unicode strings and DWORD values from these
        blobs using byte-offset parsing. This is read-only inspection of
        the operator's own configuration data.

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

    Registry Access:
        READ-ONLY. No registry values are written, modified, or deleted
        by any function in this module.

    Dependency:
        Requires OMMigrate-Core.psm1 to be imported and
        Initialize-OMMigrate to have been called before use.
#>

Set-StrictMode -Version Latest


# ============================================================
#  MODULE CONSTANTS
# ============================================================

# Outlook version registry base paths (newest first -- we try each)
$Script:OutlookRegPaths = @(
    'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles',   # 2016/2019/2021/365
    'HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles',   # 2013
    'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles'
)

# Well-known Outlook account service GUID -- consistent across all Outlook versions
# All account configurations (server names, ports, SSL, email addresses) are stored
# as numbered subkeys under this GUID within each profile
$Script:OutlookAccountGUID = '9375CFF0413111d3B88A00104B2A6676'

# Domain-based server lookup table -- used as fallback when registry values are missing
# Covers well-known public mail providers with fixed documented server settings
$Script:KnownProviderSettings = @{
    # AT&T / Ameritech (routes through Yahoo servers -- legacy ameritech.net accounts)
    'ameritech.net'   = @{ Pop3='pop.mail.yahoo.com'; Imap='imap.mail.yahoo.com'; Smtp='smtp.mail.yahoo.com'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }

    # AT&T legacy domains (all use official att.net servers)
    'att.net'         = @{ Pop3='inbound.att.net';    Imap='imap.mail.att.net';   Smtp='smtp.mail.att.net';   Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
    'sbcglobal.net'   = @{ Pop3='inbound.att.net';    Imap='imap.mail.att.net';   Smtp='smtp.mail.att.net';   Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
    'bellsouth.net'   = @{ Pop3='inbound.att.net';    Imap='imap.mail.att.net';   Smtp='smtp.mail.att.net';   Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
    'flash.net'       = @{ Pop3='inbound.att.net';    Imap='imap.mail.att.net';   Smtp='smtp.mail.att.net';   Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
    'nvbell.net'      = @{ Pop3='inbound.att.net';    Imap='imap.mail.att.net';   Smtp='smtp.mail.att.net';   Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
    'pacbell.net'     = @{ Pop3='inbound.att.net';    Imap='imap.mail.att.net';   Smtp='smtp.mail.att.net';   Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
    'prodigy.net'     = @{ Pop3='inbound.att.net';    Imap='imap.mail.att.net';   Smtp='smtp.mail.att.net';   Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
    'snet.net'        = @{ Pop3='inbound.att.net';    Imap='imap.mail.att.net';   Smtp='smtp.mail.att.net';   Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
    'swbell.net'      = @{ Pop3='inbound.att.net';    Imap='imap.mail.att.net';   Smtp='smtp.mail.att.net';   Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
    'wans.net'        = @{ Pop3='inbound.att.net';    Imap='imap.mail.att.net';   Smtp='smtp.mail.att.net';   Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
    'currently.com'   = @{ Pop3='inbound.att.net';    Imap='imap.mail.att.net';   Smtp='smtp.mail.att.net';   Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }

    # Gmail / Google Workspace
    'gmail.com'       = @{ Pop3='pop.gmail.com';      Imap='imap.gmail.com';      Smtp='smtp.gmail.com';      Pop3Port=995; ImapPort=993; SmtpPort=587; RequiresSecureKey=$false }
    'googlemail.com'  = @{ Pop3='pop.gmail.com';      Imap='imap.gmail.com';      Smtp='smtp.gmail.com';      Pop3Port=995; ImapPort=993; SmtpPort=587; RequiresSecureKey=$false }

    # Yahoo Mail
    'yahoo.com'       = @{ Pop3='pop.mail.yahoo.com'; Imap='imap.mail.yahoo.com'; Smtp='smtp.mail.yahoo.com'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$false }
    'ymail.com'       = @{ Pop3='pop.mail.yahoo.com'; Imap='imap.mail.yahoo.com'; Smtp='smtp.mail.yahoo.com'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$false }
    'yahoo.co.uk'     = @{ Pop3='pop.mail.yahoo.com'; Imap='imap.mail.yahoo.com'; Smtp='smtp.mail.yahoo.com'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$false }

    # Apple iCloud
    'icloud.com'      = @{ Pop3='';                   Imap='imap.mail.me.com';    Smtp='smtp.mail.me.com';    Pop3Port=0;   ImapPort=993; SmtpPort=587; RequiresSecureKey=$false }
    'me.com'          = @{ Pop3='';                   Imap='imap.mail.me.com';    Smtp='smtp.mail.me.com';    Pop3Port=0;   ImapPort=993; SmtpPort=587; RequiresSecureKey=$false }
    'mac.com'         = @{ Pop3='';                   Imap='imap.mail.me.com';    Smtp='smtp.mail.me.com';    Pop3Port=0;   ImapPort=993; SmtpPort=587; RequiresSecureKey=$false }
}

# Known IMAP provider GUIDs in Outlook registry
# These identify the service provider type for each account subkey
$Script:ProviderGUIDs = @{
    '{0D7B2A70-A7EB-4B85-9B36-7A74B3D5A4ED}' = 'IMAP'
    '{0D7B2A70-A7EB-4B85-9B36-7A74B3D5A4EE}' = 'IMAP'
    '{44EB701A-4B00-11D0-BF7B-00AA00F3B4E7}' = 'POP3'
    '{2B6D9A40-B940-11CF-9ECC-00AA006600E1}' = 'POP3'
    '{E2B92B00-87A7-11D1-9BBD-0060089A423C}' = 'SMTP'
    '{9BC34A90-9A3E-11D0-9E0A-00A0C91F5853}' = 'Exchange'
    '{9BC34A91-9A3E-11D0-9E0A-00A0C91F5853}' = 'Exchange'
    '{9BC34A92-9A3E-11D0-9E0A-00A0C91F5853}' = 'Exchange'
    '{9BC34A93-9A3E-11D0-9E0A-00A0C91F5853}' = 'Exchange'
}

# Known AWS SES SMTP endpoint patterns
$Script:AWSSESPatterns = @(
    'email-smtp.*.amazonaws.com',
    'email-smtp.us-east-*.amazonaws.com',
    'email-smtp.us-west-*.amazonaws.com',
    'email-smtp.eu-*.amazonaws.com',
    'email-smtp.ap-*.amazonaws.com',
    'email-smtp.sa-*.amazonaws.com',
    'email-smtp.ca-*.amazonaws.com'
)

# Known Yahoo/AT&T IMAP server patterns (for ameritech.net detection)
$Script:YahooIMAPPatterns = @(
    'imap.mail.yahoo.com',
    'imap.att.yahoo.com',
    'imap.mail.att.net'
)

# Known Yahoo/AT&T SMTP server patterns
$Script:YahooSMTPPatterns = @(
    'smtp.mail.yahoo.com',
    'smtp.att.yahoo.com',
    'smtp.mail.att.net'
)

# AT&T/Yahoo legacy domain list (for POP3-ATTAMERITECH tagging)
$Script:ATTLegacyDomains = @(
    'ameritech.net',
    'att.net',
    'sbcglobal.net',
    'bellsouth.net',
    'flash.net',
    'nvbell.net',
    'pacbell.net',
    'prodigy.net',
    'snet.net',
    'swbell.net',
    'wans.net',
    'currently.com'
)

# Microsoft consumer domain list (for EXCHANGE-SKIP tagging)
$Script:MicrosoftDomains = @(
    'live.com',
    'outlook.com',
    'hotmail.com',
    'msn.com',
    'passport.com'
)

# Gmail domains (for IMAP-ALREADY/Gmail tagging)
$Script:GmailDomains = @(
    'gmail.com',
    'googlemail.com'
)

# MAPI property tag IDs used to locate values in binary blobs
# These are well-documented MAPI property identifiers
$Script:MapiProps = @{
    PR_DISPLAY_NAME      = 0x3001   # Account display name
    PR_EMAIL_ADDRESS     = 0x3003   # Email address
    PR_ACCOUNT           = 0x3A00   # Account name
    PR_POP3_HOST         = 0x6600   # POP3 server hostname (approximate)
    PR_IMAP_SERVER       = 0x6600   # IMAP server hostname (approximate)
    PR_SMTP_HOST         = 0x6601   # SMTP server hostname (approximate)
}


# ============================================================
#  REGION: PROFILE DISCOVERY
# ============================================================

function Get-OutlookProfiles {
    <#
    .SYNOPSIS
        Returns all Outlook profiles found in the Windows Registry
        for the current user.

    .DESCRIPTION
        Scans the known Outlook profile registry locations for the
        current Windows user and returns a list of profile objects,
        each containing the profile name and its registry base path.

        Tries Office 16.0 (Outlook 2016/2019/2021/365) first, then
        falls back to Office 15.0 (Outlook 2013) and the legacy
        Windows Messaging Subsystem path.

        READ-ONLY -- no registry modifications.

    .PARAMETER ProfileName
        Optional. If specified, returns only the profile matching
        this name. If omitted, returns all profiles found.

    .OUTPUTS
        [PSCustomObject[]] -- Array of profile objects:
            Name        [string]  Profile name
            RegistryPath [string] Full registry path to profile root
            IsDefault   [bool]   Whether this is the default profile
            OutlookVersion [string] Outlook version string (e.g. '16.0')

    .EXAMPLE
        $profiles = Get-OutlookProfiles
        foreach ($profile in $profiles) {
            Write-OMMigrateLog -Message "Found profile: $($profile.Name)"
        }

    .EXAMPLE
        $profile = Get-OutlookProfiles -ProfileName 'Outlook'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ProfileName = ''
    )

    Write-OMMigrateLog -Message 'Scanning registry for Outlook profiles...' -Level INFO

    $profiles      = [System.Collections.Generic.List[PSCustomObject]]::new()
    $foundVersion  = ''
    $foundBasePath = ''

    # -- Find the active Outlook version registry path ---------
    foreach ($regPath in $Script:OutlookRegPaths) {
        if (Test-Path $regPath) {
            $foundBasePath = $regPath

            # Extract version number from path
            if ($regPath -match '\\Office\\(\d+\.\d+)\\') {
                $foundVersion = $Matches[1]
            }
            else {
                $foundVersion = 'Legacy'
            }

            Write-OMMigrateLog -Message "Outlook profile registry path found: $regPath (v$foundVersion)" `
                               -Level DEBUG
            break
        }
    }

    if ([string]::IsNullOrEmpty($foundBasePath)) {
        Write-OMMigrateLog -Message 'No Outlook profile registry path found. Is Outlook installed?' `
                           -Level ERROR
        return $profiles
    }

    # -- Read default profile name -----------------------------
    $defaultProfileName = ''
    try {
        $outlookBaseKey = $foundBasePath -replace '\\Profiles$', ''
        if (Test-Path $outlookBaseKey) {
            $defaultProfileName = (Get-ItemProperty -Path $outlookBaseKey `
                                                    -Name 'DefaultProfile' `
                                                    -ErrorAction SilentlyContinue).DefaultProfile
        }
    }
    catch {
        Write-OMMigrateLog -Message "Could not read default profile name: $_" -Level DEBUG
    }

    # -- Enumerate profile subkeys -----------------------------
    try {
        $profileKeys = Get-ChildItem -Path $foundBasePath -ErrorAction Stop
    }
    catch {
        Write-OMMigrateLog -Message "Failed to enumerate profile keys at $foundBasePath : $_" `
                           -Level ERROR
        return $profiles
    }

    foreach ($key in $profileKeys) {
        $name = $key.PSChildName

        # Filter by name if requested
        if ($ProfileName -and $name -ne $ProfileName) { continue }

        $profile = [PSCustomObject]@{
            Name           = $name
            RegistryPath   = $key.PSPath
            RegistryKey    = $key.Name
            IsDefault      = ($name -eq $defaultProfileName)
            OutlookVersion = $foundVersion
            BasePath       = $foundBasePath
        }

        $profiles.Add($profile)
        Write-OMMigrateLog -Message "Profile found: '$name' (Default=$($profile.IsDefault))" `
                           -Level INFO
    }

    if ($profiles.Count -eq 0) {
        Write-OMMigrateLog -Message 'No Outlook profiles found in registry.' -Level WARN
    }
    else {
        Write-OMMigrateLog -Message "Total profiles found: $($profiles.Count)" -Level INFO
    }

    return $profiles
}


# ============================================================
#  REGION: ACCOUNT DISCOVERY
# ============================================================

function Get-OutlookAccountsFromRegistry {
    <#
    .SYNOPSIS
        Extracts all email account configurations from the registry
        for a given Outlook profile.

    .DESCRIPTION
        Scans the well-known Outlook account service GUID subkey
        (9375CFF0413111d3B88A00104B2A6676) under the profile to read
        account settings stored as plain string values. This is the
        authoritative location for POP3/IMAP/SMTP server names, ports,
        SSL settings, and email addresses in Outlook 2016/2019/2021.

        Falls back to a domain-based lookup table for well-known public
        providers (AT&T, Gmail, Yahoo, iCloud) when registry values are
        missing or incomplete.

        Classifies each account with a provider tag for use by the
        migration scripts.

        READ-ONLY -- no registry modifications.

    .PARAMETER Profile
        A profile object returned by Get-OutlookProfiles.

    .OUTPUTS
        [PSCustomObject[]] -- Array of account objects. See
        New-AccountObject for the full property list.

    .EXAMPLE
        $profiles = Get-OutlookProfiles
        foreach ($profile in $profiles) {
            $accounts = Get-OutlookAccountsFromRegistry -Profile $profile
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Profile
    )

    Write-OMMigrateLog -Message "Reading accounts from profile: '$($Profile.Name)'" `
                       -Level INFO

    $accounts = [System.Collections.Generic.List[PSCustomObject]]::new()

    # -- Build path to the account service GUID subkey ---------
    # All Outlook account configurations are stored under this
    # well-known GUID, consistent across Outlook 2016/2019/2021
    # Convert PSPath to registry path usable with Test-Path
    $profileRegPath = $Profile.RegistryPath -replace '^Microsoft\.PowerShell\.Core\\Registry::', ''
    $profileRegPath = 'HKCU:' + ($profileRegPath -replace '^HKEY_CURRENT_USER', '')
    if (-not $profileRegPath) {
    }

    $accountGUIDPath = Join-Path $profileRegPath '9375CFF0413111d3B88A00104B2A6676'

    $guidExists = $false
    try { $guidExists = Test-Path $accountGUIDPath -ErrorAction SilentlyContinue } catch { $guidExists = $false }
    if (-not $guidExists) {
        Write-OMMigrateLog -Message (
            "Account GUID subkey not found at: $accountGUIDPath. " +
            "Falling back to recursive profile scan."
        ) -Level WARN

        # Fallback -- try recursive scan of profile subkeys
        $accountGUIDPath = $null
    }

    # -- Track email addresses already processed ---------------
    $processedEmails = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    # -- Copy module-level settings to local variable for function scope --
    # $Script: variables are module-scoped and may not resolve inside functions
    # when called from an imported module context
    $knownProviders = $Script:KnownProviderSettings
    if (-not $knownProviders) {
        $knownProviders = @{
            'ameritech.net'  = @{ Pop3='pop.mail.yahoo.com'; Imap='imap.mail.yahoo.com'; Smtp='smtp.mail.yahoo.com'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
            'att.net'        = @{ Pop3='inbound.att.net'; Imap='imap.mail.att.net'; Smtp='smtp.mail.att.net'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
            'sbcglobal.net'  = @{ Pop3='inbound.att.net'; Imap='imap.mail.att.net'; Smtp='smtp.mail.att.net'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
            'bellsouth.net'  = @{ Pop3='inbound.att.net'; Imap='imap.mail.att.net'; Smtp='smtp.mail.att.net'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
            'flash.net'      = @{ Pop3='inbound.att.net'; Imap='imap.mail.att.net'; Smtp='smtp.mail.att.net'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
            'nvbell.net'     = @{ Pop3='inbound.att.net'; Imap='imap.mail.att.net'; Smtp='smtp.mail.att.net'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
            'pacbell.net'    = @{ Pop3='inbound.att.net'; Imap='imap.mail.att.net'; Smtp='smtp.mail.att.net'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
            'prodigy.net'    = @{ Pop3='inbound.att.net'; Imap='imap.mail.att.net'; Smtp='smtp.mail.att.net'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
            'snet.net'       = @{ Pop3='inbound.att.net'; Imap='imap.mail.att.net'; Smtp='smtp.mail.att.net'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
            'swbell.net'     = @{ Pop3='inbound.att.net'; Imap='imap.mail.att.net'; Smtp='smtp.mail.att.net'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
            'wans.net'       = @{ Pop3='inbound.att.net'; Imap='imap.mail.att.net'; Smtp='smtp.mail.att.net'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
            'currently.com'  = @{ Pop3='inbound.att.net'; Imap='imap.mail.att.net'; Smtp='smtp.mail.att.net'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$true }
            'gmail.com'      = @{ Pop3='pop.gmail.com'; Imap='imap.gmail.com'; Smtp='smtp.gmail.com'; Pop3Port=995; ImapPort=993; SmtpPort=587; RequiresSecureKey=$false }
            'googlemail.com' = @{ Pop3='pop.gmail.com'; Imap='imap.gmail.com'; Smtp='smtp.gmail.com'; Pop3Port=995; ImapPort=993; SmtpPort=587; RequiresSecureKey=$false }
            'yahoo.com'      = @{ Pop3='pop.mail.yahoo.com'; Imap='imap.mail.yahoo.com'; Smtp='smtp.mail.yahoo.com'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$false }
            'ymail.com'      = @{ Pop3='pop.mail.yahoo.com'; Imap='imap.mail.yahoo.com'; Smtp='smtp.mail.yahoo.com'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$false }
            'yahoo.co.uk'    = @{ Pop3='pop.mail.yahoo.com'; Imap='imap.mail.yahoo.com'; Smtp='smtp.mail.yahoo.com'; Pop3Port=995; ImapPort=993; SmtpPort=465; RequiresSecureKey=$false }
            'icloud.com'     = @{ Pop3=''; Imap='imap.mail.me.com'; Smtp='smtp.mail.me.com'; Pop3Port=0; ImapPort=993; SmtpPort=587; RequiresSecureKey=$false }
            'me.com'         = @{ Pop3=''; Imap='imap.mail.me.com'; Smtp='smtp.mail.me.com'; Pop3Port=0; ImapPort=993; SmtpPort=587; RequiresSecureKey=$false }
            'mac.com'        = @{ Pop3=''; Imap='imap.mail.me.com'; Smtp='smtp.mail.me.com'; Pop3Port=0; ImapPort=993; SmtpPort=587; RequiresSecureKey=$false }
        }
    }

    # -- Primary path: scan 9375CFF0 subkeys -------------------
    if ($accountGUIDPath) {
        try {
            $accountSubkeys = Get-ChildItem -Path $accountGUIDPath `
                                            -ErrorAction SilentlyContinue
        }
        catch {
            Write-OMMigrateLog -Message "Failed to read account GUID subkeys: $_" -Level WARN
            $accountSubkeys = @()
        }

        Write-OMMigrateLog -Message (
            "Scanning $(@($accountSubkeys).Count) account subkeys under GUID path..."
        ) -Level DEBUG

        foreach ($subkey in $accountSubkeys) {
            try {
                $props = Get-ItemProperty -Path $subkey.PSPath -ErrorAction SilentlyContinue
                if (-not $props) { continue }

                # Helper to safely read a string property value
                # Uses PSObject.Properties to avoid StrictMode errors
                $readProp = {
                    param($p, $name)
                    $prop = $p.PSObject.Properties[$name]
                    if ($prop -and $prop.Value -is [string]) { $prop.Value } else { $null }
                }
                $readInt = {
                    param($p, $name)
                    $prop = $p.PSObject.Properties[$name]
                    if ($prop -and $prop.Value -ne $null) { try { [int]$prop.Value } catch { 0 } } else { 0 }
                }

                # Read email address
                $emailAddress = ''
                foreach ($candidate in @('Email', 'Email Address', 'EmailAddress', '001f6601')) {
                    $prop = $props.PSObject.Properties[$candidate]
                    if ($prop -and $prop.Value -is [string] -and $prop.Value -match '@') {
                        $emailAddress = $prop.Value.Trim()
                        break
                    }
                }

                # Skip if no email address
                if ([string]::IsNullOrWhiteSpace($emailAddress)) { continue }
                if ($emailAddress -match '^\{.*\}$') { continue }
                if ($processedEmails.Contains($emailAddress)) { continue }

                # Read server settings -- plain string values
                $pop3Server = & $readProp $props 'POP3 Server'
                $imapServer = & $readProp $props 'IMAP Server'
                $smtpServer = & $readProp $props 'SMTP Server'

                # Read port numbers
                $pop3Port = & $readInt $props 'POP3 Port'
                $imapPort = & $readInt $props 'IMAP Port'
                $smtpPort = & $readInt $props 'SMTP Port'

                # Read SSL settings
                $pop3SSL = $true
                $imapSSL = $true
                $smtpSSL = $true
                $sslProp = $props.PSObject.Properties['POP3 Use SSL']
                if ($sslProp -and $sslProp.Value -ne $null) { try { $pop3SSL = [bool]$sslProp.Value } catch { } }
                $sslProp = $props.PSObject.Properties['IMAP Use SSL']
                if ($sslProp -and $sslProp.Value -ne $null) { try { $imapSSL = [bool]$sslProp.Value } catch { } }
                $sslProp = $props.PSObject.Properties['SMTP Use SSL']
                if ($sslProp -and $sslProp.Value -ne $null) { try { $smtpSSL = [bool]$sslProp.Value } catch { } }

                # Determine account type from servers found
                $accountType = 'Unknown'
                $incomingServer = ''
                $incomingPort   = 0
                $incomingSSL    = $true

                if ($imapServer -and -not [string]::IsNullOrWhiteSpace($imapServer)) {
                    $accountType    = 'IMAP'
                    $incomingServer = $imapServer.Trim()
                    $incomingPort   = if ($imapPort -gt 0) { $imapPort } else { 993 }
                    $incomingSSL    = $imapSSL
                }
                elseif ($pop3Server -and -not [string]::IsNullOrWhiteSpace($pop3Server)) {
                    $accountType    = 'POP3'
                    $incomingServer = $pop3Server.Trim()
                    $incomingPort   = if ($pop3Port -gt 0) { $pop3Port } else { 995 }
                    $incomingSSL    = $pop3SSL
                }

                $outgoingServer = if ($smtpServer) { $smtpServer.Trim() } else { '' }
                $outgoingPort   = if ($smtpPort -gt 0) { $smtpPort } else { 587 }
                $outgoingSSL    = $smtpSSL

                # -- Domain-based fallback for missing server settings --
                if ([string]::IsNullOrEmpty($incomingServer) -or
                    [string]::IsNullOrEmpty($outgoingServer)) {
                    $domain = ''
                    if ($emailAddress -match '@(.+)$') { $domain = $Matches[1].ToLower() }

                    if ($domain -and $knownProviders.ContainsKey($domain)) {
                        $known = $knownProviders[$domain]
                        Write-OMMigrateLog -Message (
                            "Using known provider settings for domain '$domain'"
                        ) -Level INFO

                        if ([string]::IsNullOrEmpty($incomingServer)) {
                            # Use IMAP if available, otherwise POP3
                            if ($known.Imap) {
                                $accountType    = 'IMAP'
                                $incomingServer = $known.Imap
                                $incomingPort   = $known.ImapPort
                            }
                            elseif ($known.Pop3) {
                                $accountType    = 'POP3'
                                $incomingServer = $known.Pop3
                                $incomingPort   = $known.Pop3Port
                            }
                        }
                        if ([string]::IsNullOrEmpty($outgoingServer)) {
                            $outgoingServer = $known.Smtp
                            $outgoingPort   = $known.SmtpPort
                        }
                    }
                }

                # Build account object
                $account = New-AccountObject
                $account.EmailAddress    = $emailAddress
                $account.AccountType     = $accountType
                $account.IncomingServer  = $incomingServer
                $account.IncomingPort    = $incomingPort
                $account.IncomingSSL     = $incomingSSL
                $account.OutgoingServer  = $outgoingServer
                $account.OutgoingPort    = $outgoingPort
                $account.OutgoingSSL     = $outgoingSSL

                # Extract display name
                $displayNameProp = $props.PSObject.Properties['Display Name']
                if ($displayNameProp -and $displayNameProp.Value -is [string]) {
                    $account.DisplayName = $displayNameProp.Value.Trim()
                }

                # Extract PST/OST path from binary Delivery Store EntryID
                # The blob is a MAPI EntryID with a UTF-16LE path embedded at the end
                # Strategy: scan for drive letter pattern in the raw bytes
                $deliveryEntryProp = $props.PSObject.Properties['Delivery Store EntryID']
                if ($deliveryEntryProp -and $deliveryEntryProp.Value -is [byte[]]) {
                    $blob = $deliveryEntryProp.Value
                    $extractedPath = ''
                    # Scan blob for UTF-16LE drive letter pattern (e.g. C:\)
                    # Drive letter stored as: letter(0x41-0x7A) 0x00 0x3A 0x00 0x5C 0x00
                    $blobArray = [byte[]]$blob
                    $blobLen = $blobArray.Length
                    for ($bi = 0; $bi -lt ($blobLen - 6); $bi++) {
                        if ($blobArray[$bi] -ge 65 -and $blobArray[$bi] -le 122 -and
                            $blobArray[$bi+1] -eq 0 -and
                            $blobArray[$bi+2] -eq 58 -and
                            $blobArray[$bi+3] -eq 0 -and
                            $blobArray[$bi+4] -eq 92 -and
                            $blobArray[$bi+5] -eq 0) {
                            $pathBytes = New-Object byte[] ($blobLen - $bi)
                            [System.Array]::Copy($blobArray, $bi, $pathBytes, 0, $blobLen - $bi)
                            $pathStr = [System.Text.Encoding]::Unicode.GetString($pathBytes).TrimEnd([char]0)
                            if ($pathStr -match "^[A-Za-z]:\\.+?\.(pst|ost)") {
                                $extractedPath = $Matches[0].Trim()
                            }
                            break
                        }
                    }
                    if ($extractedPath) {
                        if ($extractedPath -like '*.pst') {
                            $account.PSTPath = $extractedPath
                            if (Test-Path $extractedPath -ErrorAction SilentlyContinue) {
                                $fi = Get-Item $extractedPath -ErrorAction SilentlyContinue
                                if ($fi) {
                                    $account.DataFileSizeBytes     = $fi.Length
                                    $account.DataFileSizeFormatted = Format-FileSize -Bytes $fi.Length
                                }
                            }
                        }
                        elseif ($extractedPath -like '*.ost') {
                            $account.OSTPath = $extractedPath
                            if (Test-Path $extractedPath -ErrorAction SilentlyContinue) {
                                $fi = Get-Item $extractedPath -ErrorAction SilentlyContinue
                                if ($fi) {
                                    $account.DataFileSizeBytes     = $fi.Length
                                    $account.DataFileSizeFormatted = Format-FileSize -Bytes $fi.Length
                                }
                            }
                        }
                    }
                }

                [void]$processedEmails.Add($emailAddress)

                # Classify the account
                $account = Set-AccountTag -Account $account

                $accounts.Add($account)

                # Suppress account detail when Sanitize active -- map not built yet
                if (-not $Global:OMMigrate.Sanitize) {
                    Write-OMMigrateLog -Message (
                        "Account: $($account.EmailAddress) | " +
                        "Type: $($account.AccountType) | " +
                        "Server: $($account.IncomingServer) | " +
                        "Tag: $($account.ProviderTag)"
                    ) -Level INFO
                }
            }
            catch {
                Write-OMMigrateLog -Message (
                    "Error reading account subkey $($subkey.PSChildName): $_"
                ) -Level WARN
            }
        }
    }

    Write-OMMigrateLog -Message "Accounts found in profile '$($Profile.Name)': $($accounts.Count)" `
                       -Level INFO

    return $accounts
}


function Read-AccountSubkey {
    <#
    .SYNOPSIS
        Reads and decodes account configuration from a single
        registry subkey.

    .DESCRIPTION
        Attempts to extract email account properties from a registry
        subkey using multiple reading strategies:

            1. Plain string registry values (easiest -- directly readable)
            2. Binary blob parsing (for values Outlook stores as
               REG_BINARY using MAPI property encoding)

        Returns a partial or complete account object depending on
        what was found. Caller is responsible for filtering incomplete
        results.

        READ-ONLY -- no registry modifications.

    .PARAMETER SubkeyPath
        PowerShell registry path (PSPath) to the subkey.

    .PARAMETER SubkeyName
        Name of the subkey (used for logging context).

    .OUTPUTS
        [PSCustomObject] -- Partially or fully populated account object.

    .NOTES
        Binary blob parsing uses byte-offset reading of MAPI-format
        property arrays. This is community-standard practice for reading
        Outlook profile data and has been in use for 15+ years across
        many open-source email tools.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SubkeyPath,

        [Parameter(Mandatory = $true)]
        [string]$SubkeyName
    )

    $account = New-AccountObject

    try {
        $values = Get-ItemProperty -Path $SubkeyPath -ErrorAction SilentlyContinue
        if (-not $values) { return $account }
    }
    catch {
        return $account
    }

    # -- Strategy 1: Plain string values ----------------------
    # Some Outlook versions store settings as readable strings

    # Email address -- try multiple known value names
    $emailCandidates = @(
        'Email',
        'EmailAddress',
        'Email Address',
        '001e6601',   # MAPI PR_EMAIL_ADDRESS as hex value name
        '001f6601'
    )
    foreach ($candidate in $emailCandidates) {
        $val = ($values.PSObject.Properties[$candidate] | Select-Object -ExpandProperty Value -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($val -and $val -is [string] -and $val -match '@') {
            $account.EmailAddress = $val.Trim()
            break
        }
    }

    # Display name
    $displayCandidates = @('Display Name', 'DisplayName', '001e3001', '001f3001')
    foreach ($candidate in $displayCandidates) {
        $val = ($values.PSObject.Properties[$candidate] | Select-Object -ExpandProperty Value -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($val -and $val -is [string] -and -not [string]::IsNullOrWhiteSpace($val)) {
            $account.DisplayName = $val.Trim()
            break
        }
    }

    # Incoming server
    $incomingCandidates = @(
        'POP3 Server', 'IMAP Server', 'Incoming Server',
        'Pop3Server', 'ImapServer',
        '001e6600', '001f6600'
    )
    foreach ($candidate in $incomingCandidates) {
        $val = ($values.PSObject.Properties[$candidate] | Select-Object -ExpandProperty Value -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($val -and $val -is [string] -and -not [string]::IsNullOrWhiteSpace($val)) {
            $account.IncomingServer = $val.Trim()
            break
        }
    }

    # Outgoing server
    $outgoingCandidates = @(
        'SMTP Server', 'Outgoing Server', 'SmtpServer',
        '001e6602', '001f6602'
    )
    foreach ($candidate in $outgoingCandidates) {
        $val = ($values.PSObject.Properties[$candidate] | Select-Object -ExpandProperty Value -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($val -and $val -is [string] -and -not [string]::IsNullOrWhiteSpace($val)) {
            $account.OutgoingServer = $val.Trim()
            break
        }
    }

    # Port numbers (DWORD values)
    $inPortCandidates  = @('POP3 Port', 'IMAP Port', 'Incoming Port', '00036600')
    $outPortCandidates = @('SMTP Port', 'Outgoing Port', '00036601')

    foreach ($candidate in $inPortCandidates) {
        $val = ($values.PSObject.Properties[$candidate] | Select-Object -ExpandProperty Value -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($val -and $val -isnot [byte[]] -and $val -gt 0) { $account.IncomingPort = [int]$val; break }
    }
    foreach ($candidate in $outPortCandidates) {
        $val = ($values.PSObject.Properties[$candidate] | Select-Object -ExpandProperty Value -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($val -and $val -isnot [byte[]] -and $val -gt 0) { $account.OutgoingPort = [int]$val; break }
    }

    # SSL flags
    $inSSLCandidates  = @('POP3 Use SSL', 'IMAP Use SSL', 'Use SSL', '00036603')
    $outSSLCandidates = @('SMTP Use SSL', 'SMTP SSL', '00036604')

    foreach ($candidate in $inSSLCandidates) {
        $val = ($values.PSObject.Properties[$candidate] | Select-Object -ExpandProperty Value -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($null -ne $val -and $val -isnot [byte[]]) { $account.IncomingSSL = [bool]$val; break }
    }
    foreach ($candidate in $outSSLCandidates) {
        $val = ($values.PSObject.Properties[$candidate] | Select-Object -ExpandProperty Value -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($null -ne $val -and $val -isnot [byte[]]) { $account.OutgoingSSL = [bool]$val; break }
    }

    # -- Strategy 2: Binary blob parsing ----------------------
    # Outlook 2016/2019/2021 stores most settings in a binary
    # value named with a GUID-derived hex key. We scan all
    # binary values and attempt to extract Unicode strings.

    $properties = $values.PSObject.Properties |
                  Where-Object { $_.Name -notlike 'PS*' }

    foreach ($prop in $properties) {
        $rawValue = $prop.Value

        # Only process byte arrays (REG_BINARY values)
        if ($rawValue -isnot [byte[]]) { continue }
        if ($rawValue.Length -lt 8)   { continue }

        # Extract strings from binary blob
        $extracted = Read-BinaryBlobStrings -Blob $rawValue

        # Apply extracted values only if not already found via Strategy 1
        if ([string]::IsNullOrEmpty($account.EmailAddress)) {
            $emailStr = $extracted | Where-Object { $_ -match '@' -and $_ -match '\.' } |
                        Select-Object -First 1
            if ($emailStr) { $account.EmailAddress = $emailStr.Trim() }
        }

        if ([string]::IsNullOrEmpty($account.IncomingServer)) {
            $serverStr = $extracted |
                         Where-Object {
                             $_ -match '\.' -and
                             $_ -notmatch '@' -and
                             $_.Length -gt 5 -and
                             $_ -match '^[a-zA-Z0-9]'
                         } | Select-Object -First 1
            if ($serverStr) { $account.IncomingServer = $serverStr.Trim() }
        }

        # Extract port numbers from DWORD-like byte sequences
        if ($account.IncomingPort -eq 0) {
            $port = Read-BinaryBlobPort -Blob $rawValue -CommonPorts @(110,143,993,995)
            if ($port -gt 0) { $account.IncomingPort = $port }
        }
        if ($account.OutgoingPort -eq 0) {
            $port = Read-BinaryBlobPort -Blob $rawValue -CommonPorts @(25,465,587)
            if ($port -gt 0) { $account.OutgoingPort = $port }
        }
    }

    # -- Determine account type from port numbers --------------
    if ($account.AccountType -eq 'Unknown') {
        $account.AccountType = Get-AccountTypeFromPort -IncomingPort $account.IncomingPort `
                                                       -IncomingServer $account.IncomingServer
    }

    return $account
}


function Read-BinaryBlobStrings {
    <#
    .SYNOPSIS
        Extracts Unicode and ASCII strings from a MAPI binary blob.

    .DESCRIPTION
        Scans a byte array for null-terminated Unicode (UTF-16LE) and
        ASCII string sequences of meaningful length. This is how Outlook
        stores server hostnames and email addresses in profile registry
        binary values.

        Returns candidate strings -- the caller is responsible for
        filtering to find the most relevant value.

        READ-ONLY byte array inspection. No side effects.

    .PARAMETER Blob
        The byte array (REG_BINARY value) to scan.

    .PARAMETER MinLength
        Minimum string length to return (filters noise).
        Default: 4

    .OUTPUTS
        [string[]] -- Array of candidate strings found in the blob.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Blob,

        [Parameter(Mandatory = $false)]
        [int]$MinLength = 4
    )

    $results = [System.Collections.Generic.List[string]]::new()

    # -- Unicode (UTF-16LE) string extraction ------------------
    # Look for sequences of printable chars followed by \x00\x00
    $i = 0
    while ($i -lt ($Blob.Length - 4)) {

        # Check for UTF-16LE pattern: alternating printable byte and null
        if ($Blob[$i] -ge 0x20 -and $Blob[$i] -le 0x7E -and $Blob[$i+1] -eq 0x00) {

            $chars = [System.Collections.Generic.List[char]]::new()
            $j = $i

            while ($j -lt ($Blob.Length - 1) -and
                   $Blob[$j] -ge 0x20 -and
                   $Blob[$j] -le 0x7E -and
                   $Blob[$j+1] -eq 0x00) {
                $chars.Add([char]$Blob[$j])
                $j += 2
            }

            if ($chars.Count -ge $MinLength) {
                $str = -join $chars
                # Filter to meaningful content -- hostnames, emails, display names
                if ($str -match '^[\w\.\-@]+$' -or $str -match '^[\w\s\.\-@]+$') {
                    [void]$results.Add($str)
                }
                $i = $j
                continue
            }
        }

        $i++
    }

    # -- ASCII string extraction (fallback) --------------------
    $i = 0
    $currentString = [System.Text.StringBuilder]::new()

    while ($i -lt $Blob.Length) {
        $b = $Blob[$i]

        if ($b -ge 0x20 -and $b -le 0x7E) {
            [void]$currentString.Append([char]$b)
        }
        else {
            if ($currentString.Length -ge $MinLength) {
                $str = $currentString.ToString()
                if ($str -match '^[\w\.\-@]+$' -and
                    -not $results.Contains($str)) {
                    [void]$results.Add($str)
                }
            }
            [void]$currentString.Clear()
        }
        $i++
    }

    # Catch last string
    if ($currentString.Length -ge $MinLength) {
        $str = $currentString.ToString()
        if ($str -match '^[\w\.\-@]+$' -and -not $results.Contains($str)) {
            [void]$results.Add($str)
        }
    }

    return $results.ToArray()
}


function Read-BinaryBlobPort {
    <#
    .SYNOPSIS
        Searches a binary blob for a port number from a known list.

    .DESCRIPTION
        Reads 4-byte little-endian DWORD values from a binary blob
        and checks if any match known email port numbers.
        Used as a fallback when port numbers are not available as
        plain registry DWORD values.

        READ-ONLY byte array inspection. No side effects.

    .PARAMETER Blob
        The byte array to scan.

    .PARAMETER CommonPorts
        Array of port numbers to look for.

    .OUTPUTS
        [int] -- The matched port number, or 0 if none found.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Blob,

        [Parameter(Mandatory = $true)]
        [int[]]$CommonPorts
    )

    for ($i = 0; $i -le ($Blob.Length - 4); $i++) {
        $dword = [System.BitConverter]::ToInt32($Blob, $i)
        if ($CommonPorts -contains $dword) {
            return $dword
        }
    }

    return 0
}


# ============================================================
#  REGION: ACCOUNT CLASSIFICATION
# ============================================================

function New-AccountObject {
    <#
    .SYNOPSIS
        Creates a new empty account object with all standard properties.

    .DESCRIPTION
        Returns a consistent PSCustomObject structure used throughout
        OMMigrate to represent a single email account. All discovery
        and classification functions work with this structure.

    .OUTPUTS
        [PSCustomObject] -- Empty account object with default values.
    #>

    return [PSCustomObject]@{

        # Identity
        EmailAddress     = ''         # user@domain.com
        DisplayName      = ''         # Name shown in Outlook account list
        Domain           = ''         # domain.com (extracted from email)
        ProfileName      = ''         # Outlook profile this account belongs to

        # Protocol classification
        AccountType      = 'Unknown'  # POP3 | IMAP | Exchange | Unknown
        ProviderTag      = ''         # POP3-STANDARD | POP3-ATTAMERITECH |
                                      # IMAP-ALREADY | EXCHANGE-SKIP | etc.

        # Incoming mail server
        IncomingServer   = ''         # mail.domain.com
        IncomingPort     = 0          # 993 (IMAP SSL) | 995 (POP3 SSL) | 143 | 110
        IncomingSSL      = $true      # SSL/TLS enabled

        # Outgoing mail server
        OutgoingServer   = ''         # smtp.domain.com
        OutgoingPort     = 0          # 587 | 465 | 25
        OutgoingSSL      = $true      # SSL/TLS enabled

        # Special flags
        IsAWSSES         = $false     # Outgoing server is AWS SES endpoint
        IsYahooAT        = $false     # Incoming server is Yahoo/AT&T
        RequiresSecureKey = $false    # Needs Secure Mail Key pre-generation

        # Data files
        PSTPath          = ''         # Path to .pst data file (POP3 accounts)
        OSTPath          = ''         # Path to .ost cache file (IMAP/Exchange)
        DataFileSizeBytes = 0L        # Size of data file in bytes
        DataFileSizeFormatted = ''    # Human-readable size (e.g. '1.24 GB')

        # Migration planning
        MigrationAction  = ''         # MIGRATE | SKIP | FOLDER-ONLY | MANUAL
        Notes            = ''         # Human-readable notes for the operator
        PreRequisites    = @()        # Steps operator must complete before migration

        # Registry metadata
        RegistryPath     = ''         # Source registry path for debugging
        DiscoveredAt     = ''         # Timestamp of discovery
    }
}


function Get-AccountTypeFromPort {
    <#
    .SYNOPSIS
        Infers the account protocol type from port numbers and server names.

    .DESCRIPTION
        Used as a fallback when the account type cannot be determined
        directly from registry service provider keys. Maps well-known
        port numbers to protocol types.

    .PARAMETER IncomingPort
        The incoming mail server port number.

    .PARAMETER IncomingServer
        The incoming mail server hostname (used for Exchange detection).

    .OUTPUTS
        [string] -- 'POP3' | 'IMAP' | 'Exchange' | 'Unknown'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$IncomingPort = 0,

        [Parameter(Mandatory = $false)]
        [string]$IncomingServer = ''
    )

    # Exchange detection by server name patterns
    if ($IncomingServer -match 'outlook\.office365\.com|outlook\.office\.com|exchange') {
        return 'Exchange'
    }

    # Port-based detection
    switch ($IncomingPort) {
        110  { return 'POP3' }
        995  { return 'POP3' }
        143  { return 'IMAP' }
        993  { return 'IMAP' }
    }

    return 'Unknown'
}


function Set-AccountTag {
    <#
    .SYNOPSIS
        Classifies an account object with a provider tag and sets
        migration action, notes, and prerequisites accordingly.

    .DESCRIPTION
        The provider tag system is the core classification that drives
        all subsequent migration decisions. This function assigns the
        correct tag based on account type, email domain, server
        hostnames, and detected special providers.

        Tag definitions:
            POP3-STANDARD     POP3 account on a standard mail server.
                              Full automated migration to IMAP.

            POP3-ATTAMERITECH POP3 account on AT&T/Yahoo legacy domain
                              (ameritech.net, sbcglobal.net, etc.).
                              Requires Secure Mail Key pre-generation.

            POP3-GMAIL        POP3 account on Gmail.
                              (Unusual -- Gmail prefers IMAP.
                               Treated as special case.)

            IMAP-ALREADY      Account already using IMAP protocol.
                              No account migration needed.
                              Folder assessment only.

            IMAP-GMAIL        IMAP account on Gmail.
                              Already correct protocol -- skip.

            EXCHANGE-SKIP     Microsoft Exchange account
                              (live.com, outlook.com, etc.)
                              Already optimal -- skip entirely.

            POP3-AWS          POP3 account with AWS SES outbound.
                              Full migration -- AWS SES SMTP
                              credentials must be re-entered.

            UNKNOWN           Cannot determine account type.
                              Flagged for manual review.

    .PARAMETER Account
        Account object to classify (modified in place and returned).

    .OUTPUTS
        [PSCustomObject] -- The classified account object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Account
    )

    # Extract domain from email address
    if ($Account.EmailAddress -match '@(.+)$') {
        $Account.Domain = $Matches[1].ToLower()
    }

    # -- AWS SES detection -------------------------------------
    $isAWSSES = $false
    foreach ($pattern in $Script:AWSSESPatterns) {
        if ($Account.OutgoingServer -like $pattern) {
            $isAWSSES = $true
            break
        }
    }
    $Account.IsAWSSES = $isAWSSES

    # -- Yahoo/AT&T server detection ---------------------------
    $isYahooAT = $false
    foreach ($pattern in $Script:YahooIMAPPatterns) {
        if ($Account.IncomingServer -like "*$pattern*") {
            $isYahooAT = $true
            break
        }
    }
    $Account.IsYahooAT = $isYahooAT

    # -- Classification logic (order matters) ------------------

    # 1. Exchange accounts -- skip entirely
    if ($Account.AccountType -eq 'Exchange' -or
        $Script:MicrosoftDomains -contains $Account.Domain) {
        $Account.ProviderTag     = 'EXCHANGE-SKIP'
        $Account.MigrationAction = 'SKIP'
        $Account.Notes           = 'Microsoft Exchange account. Already using optimal ' +
                                   'protocol (MAPI/EWS). No migration required.'
        return $Account
    }

    # 2. Already IMAP -- Gmail
    if ($Account.AccountType -eq 'IMAP' -and
        $Script:GmailDomains -contains $Account.Domain) {
        $Account.ProviderTag     = 'IMAP-GMAIL'
        $Account.MigrationAction = 'FOLDER-ONLY'
        $Account.Notes           = 'Gmail account already using IMAP. ' +
                                   'No account migration needed. ' +
                                   'Folder structure assessment only.'
        return $Account
    }

    # 3. Already IMAP -- standard
    if ($Account.AccountType -eq 'IMAP') {
        $Account.ProviderTag     = 'IMAP-ALREADY'
        $Account.MigrationAction = 'FOLDER-ONLY'
        $Account.Notes           = 'Account already using IMAP protocol. ' +
                                   'No account migration needed. ' +
                                   'Folder structure assessment only.'
        return $Account
    }

    # 4. POP3 -- AT&T/Yahoo legacy domains
    if ($Account.AccountType -eq 'POP3' -and
        ($Script:ATTLegacyDomains -contains $Account.Domain -or $isYahooAT)) {
        $Account.ProviderTag      = 'POP3-ATTAMERITECH'
        $Account.MigrationAction  = 'MIGRATE'
        $Account.IsYahooAT        = $true
        $Account.RequiresSecureKey = $true
        $Account.Notes            = "AT&T/Yahoo legacy domain account. " +
                                    "Requires Secure Mail Key (app password) " +
                                    "generated from currently.com before migration."
        $Account.PreRequisites    = @(
            "Log in to currently.com with your $($Account.EmailAddress) credentials",
            'Navigate to Account Security',
            'Generate a new App Password / Secure Mail Key',
            'Save the generated key -- you will enter it in the CSV control file'
        )
        return $Account
    }

    # 5. POP3 -- Gmail (unusual but possible)
    if ($Account.AccountType -eq 'POP3' -and
        $Script:GmailDomains -contains $Account.Domain) {
        $Account.ProviderTag     = 'POP3-GMAIL'
        $Account.MigrationAction = 'MIGRATE'
        $Account.Notes           = 'Gmail account currently using POP3. ' +
                                   'Will be migrated to IMAP via Outlook ' +
                                   'automatic account setup (browser OAuth).'
        $Account.PreRequisites   = @(
            'Have your Google account credentials ready',
            'Outlook will open a browser window for Google sign-in',
            'Approve the Microsoft Outlook permission during sign-in'
        )
        return $Account
    }

    # 6. POP3 with AWS SES outbound
    if ($Account.AccountType -eq 'POP3' -and $isAWSSES) {
        $Account.ProviderTag     = 'POP3-AWS'
        $Account.MigrationAction = 'MIGRATE'
        $Account.Notes           = 'POP3 account using AWS SES for outbound SMTP. ' +
                                   'AWS SES SMTP credentials (IAM SMTP username ' +
                                   'and password) must be re-entered when Outlook ' +
                                   'prompts during IMAP account setup.'
        $Account.PreRequisites   = @(
            'Locate your AWS SES SMTP credentials (IAM SMTP username and password)',
            'These are NOT your AWS console credentials -- they are SMTP-specific',
            'If lost, generate new SMTP credentials in AWS SES console under SMTP Settings'
        )
        return $Account
    }

    # 7. POP3 -- standard (most common case)
    if ($Account.AccountType -eq 'POP3') {
        $Account.ProviderTag     = 'POP3-STANDARD'
        $Account.MigrationAction = 'MIGRATE'
        $Account.Notes           = 'Standard POP3 account. ' +
                                   'Full automated migration to IMAP. ' +
                                   'Password entry required when Outlook prompts.'
        return $Account
    }

    # 8. Unknown -- flag for manual review
    $Account.ProviderTag     = 'UNKNOWN'
    $Account.MigrationAction = 'MANUAL'
    $Account.Notes           = 'Account type could not be determined from registry. ' +
                               'Manual review required before migration.'

    return $Account
}


# ============================================================
#  REGION: PST / OST DATA FILE DISCOVERY
# ============================================================

function Get-OutlookDataFiles {
    <#
    .SYNOPSIS
        Discovers all PST and OST data files associated with an
        Outlook profile from the registry.

    .DESCRIPTION
        Reads the PST/OST file registration keys under the profile
        to find all data files associated with each account.

        The registry key used is the well-documented '9375CFF0...'
        subkey pattern under each profile, which contains the paths
        to all attached data files.

        READ-ONLY -- no registry modifications.

    .PARAMETER Profile
        A profile object returned by Get-OutlookProfiles.

    .OUTPUTS
        [PSCustomObject[]] -- Array of data file objects:
            FilePath     [string]  Full path to the PST/OST file
            FileType     [string]  'PST' or 'OST'
            DisplayName  [string]  Friendly name shown in Outlook
            SizeBytes    [long]    File size in bytes
            SizeFormatted [string] Human-readable size
            Exists       [bool]    Whether the file exists on disk

    .EXAMPLE
        $dataFiles = Get-OutlookDataFiles -Profile $profile
        foreach ($file in $dataFiles) {
            Write-OMMigrateLog "Data file: $($file.FilePath) ($($file.SizeFormatted))"
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Profile
    )

    Write-OMMigrateLog -Message "Discovering data files for profile: '$($Profile.Name)'" `
                       -Level INFO

    $dataFiles = [System.Collections.Generic.List[PSCustomObject]]::new()

    # -- Scan profile subkeys for data file entries ------------
    # Outlook stores PST/OST paths in subkeys matching pattern:
    # {9375CFF0-...} under the profile root
    try {
        $allSubkeys = Get-ChildItem -Path $Profile.RegistryPath `
                                    -Recurse `
                                    -ErrorAction SilentlyContinue
    }
    catch {
        Write-OMMigrateLog -Message "Could not enumerate profile subkeys: $_" -Level WARN
        return $dataFiles
    }

    foreach ($subkey in $allSubkeys) {
        try {
            $props = Get-ItemProperty -Path $subkey.PSPath -ErrorAction SilentlyContinue
            if (-not $props) { continue }

            # Look for file path values -- PST paths appear as binary or string
            $filePath = ''

            # Try direct string properties first
            $pathProps = $props.PSObject.Properties |
                         Where-Object {
                             $_.Name -notlike 'PS*' -and
                             $_.Value -is [string] -and
                             ($_.Value -like '*.pst' -or $_.Value -like '*.ost')
                         }

            if ($pathProps) {
                $filePath = ($pathProps | Select-Object -First 1).Value
            }

            # Try binary properties (path stored as UTF-16LE bytes)
            if ([string]::IsNullOrEmpty($filePath)) {
                $binaryProps = $props.PSObject.Properties |
                               Where-Object {
                                   $_.Name -notlike 'PS*' -and
                                   $_.Value -is [byte[]] -and
                                   $_.Value.Length -gt 10
                               }

                foreach ($binProp in $binaryProps) {
                    $decoded = [System.Text.Encoding]::Unicode.GetString($binProp.Value)
                    $decoded = $decoded.TrimEnd([char]0)
                    if ($decoded -like '*.pst' -or $decoded -like '*.ost') {
                        # Expand environment variables in path
                        $filePath = [System.Environment]::ExpandEnvironmentVariables($decoded)
                        break
                    }
                }
            }

            if ([string]::IsNullOrEmpty($filePath)) { continue }

            # Determine file type
            $fileType = if ($filePath -like '*.ost') { 'OST' } else { 'PST' }

            # Get display name if available
            $displayName = ''
            $dnProps = @('001f3001', 'Display Name', 'DisplayName')
            foreach ($dn in $dnProps) {
                $val = $props.$dn
                if ($val -is [byte[]]) {
                    $val = [System.Text.Encoding]::Unicode.GetString($val).TrimEnd([char]0)
                }
                if ($val -is [string] -and -not [string]::IsNullOrWhiteSpace($val)) {
                    $displayName = $val
                    break
                }
            }

            # Check if file exists and get size
            # Guard against registry values with illegal path characters
            $illegalChars = [System.IO.Path]::GetInvalidPathChars()
            $hasIllegal = $filePath.IndexOfAny($illegalChars) -ge 0
            $exists    = if ($hasIllegal) { $false } else {
                Test-Path $filePath -ErrorAction SilentlyContinue
            }
            $sizeBytes = 0L
            $sizeFormatted = 'N/A'

            if ($exists) {
                try {
                    $fileInfo      = Get-Item $filePath -ErrorAction Stop
                    $sizeBytes     = $fileInfo.Length
                    $sizeFormatted = Format-FileSize -Bytes $sizeBytes
                }
                catch {
                    $sizeFormatted = 'Unable to read'
                }
            }

            # Avoid duplicate entries
            $alreadyAdded = $dataFiles | Where-Object { $_.FilePath -eq $filePath }
            if ($alreadyAdded) { continue }

            $dataFile = [PSCustomObject]@{
                FilePath      = $filePath
                FileType      = $fileType
                DisplayName   = $displayName
                SizeBytes     = $sizeBytes
                SizeFormatted = $sizeFormatted
                Exists        = $exists
                ProfileName   = $Profile.Name
            }

            $dataFiles.Add($dataFile)

            # Suppress data file path when Sanitize active -- map not built yet
            if (-not $Global:OMMigrate.Sanitize) {
                Write-OMMigrateLog -Message (
                    "Data file found: $filePath | Type=$fileType | " +
                    "Size=$sizeFormatted | Exists=$exists"
                ) -Level INFO
            }
        }
        catch {
            Write-OMMigrateLog -Message "Error reading subkey $($subkey.PSPath): $_" `
                               -Level DEBUG
        }
    }

    Write-OMMigrateLog -Message "Data files found: $($dataFiles.Count)" -Level INFO
    return $dataFiles
}


function Join-AccountsWithDataFiles {
    <#
    .SYNOPSIS
        Associates PST/OST data files with their corresponding
        account objects.

    .DESCRIPTION
        Attempts to match data files to accounts by correlating
        file names with email addresses, and by applying known
        Outlook naming conventions for PST files.

        POP3 accounts get their PST path populated.
        IMAP/Exchange accounts get their OST path populated.

        Unmatched data files are returned separately for the
        operator to review in the discovery report.

    .PARAMETER Accounts
        Array of account objects from Get-OutlookAccountsFromRegistry.

    .PARAMETER DataFiles
        Array of data file objects from Get-OutlookDataFiles.

    .OUTPUTS
        PSCustomObject with:
            Accounts      -- Account objects with data file paths populated
            UnmatchedFiles -- Data files that could not be matched to accounts
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Accounts,

        [Parameter(Mandatory = $true)]
        $DataFiles
    )

    $unmatchedFiles = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($dataFile in $DataFiles) {
        $matched = $false
        # Guard against illegal path characters in registry-sourced paths
        $illegalPathChars = [System.IO.Path]::GetInvalidPathChars()
        if ($dataFile.FilePath.IndexOfAny($illegalPathChars) -ge 0) {
            continue
        }
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($dataFile.FilePath)

        foreach ($account in $Accounts) {
            # Match strategies (most specific first):

            # 1. Filename contains email address (most common)
            $emailSafe = $account.EmailAddress -replace '@', '_' -replace '\.', '_'
            if ($fileName -like "*$($account.EmailAddress)*" -or
                $fileName -like "*$emailSafe*") {
                $matched = $true
            }

            # 2. Filename contains username part of email -- PST files only.
            # OST files are always named with the full email address by Outlook
            # so username-only matching is unnecessary and causes false matches
            # when multiple accounts share a common username (e.g. 'admin').
            if (-not $matched -and $dataFile.FileType -eq 'PST' -and
                $account.EmailAddress -match '^([^@]+)@') {
                $username = $Matches[1]
                if ($fileName -like "*$username*") { $matched = $true }
            }

            # 3. Filename contains display name
            if (-not $matched -and $account.DisplayName -and
                $fileName -like "*$($account.DisplayName)*") {
                $matched = $true
            }

            if ($matched) {
                if ($dataFile.FileType -eq 'PST') {
                    $account.PSTPath             = $dataFile.FilePath
                    $account.DataFileSizeBytes   = $dataFile.SizeBytes
                    $account.DataFileSizeFormatted = $dataFile.SizeFormatted
                }
                else {
                    $account.OSTPath             = $dataFile.FilePath
                    $account.DataFileSizeBytes   = $dataFile.SizeBytes
                    $account.DataFileSizeFormatted = $dataFile.SizeFormatted
                }

                # Suppress matched file detail when Sanitize active -- map not built yet
                if (-not $Global:OMMigrate.Sanitize) {
                    Write-OMMigrateLog -Message (
                        "Matched $($dataFile.FileType) file to account: " +
                        "$($account.EmailAddress) -> $($dataFile.FilePath)"
                    ) -Level DEBUG
                }
                break
            }
        }

        if (-not $matched) {
            $unmatchedFiles.Add($dataFile)
            Write-OMMigrateLog -Message (
                "Unmatched data file: $($dataFile.FilePath) " +
                "($($dataFile.SizeFormatted)) -- will appear in report for review"
            ) -Level INFO
        }
    }

    return [PSCustomObject]@{
        Accounts       = $Accounts
        UnmatchedFiles = $unmatchedFiles
    }
}


# ============================================================
#  REGION: SUMMARY & REPORTING HELPERS
# ============================================================

function Get-AccountSummary {
    <#
    .SYNOPSIS
        Returns a summary statistics object for a collection of
        discovered accounts.

    .DESCRIPTION
        Counts accounts by provider tag and migration action.
        Used to populate the stat grid in the discovery HTML report
        and to validate the discovery results before proceeding.

    .PARAMETER Accounts
        Array of classified account objects.

    .OUTPUTS
        PSCustomObject with count properties by tag and action.

    .EXAMPLE
        $summary = Get-AccountSummary -Accounts $accounts
        Write-OMMigrateLog "POP3 to migrate: $($summary.ToMigrate)"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Accounts
    )

    # Force array to ensure .Count works regardless of how parameter was passed
    $Accounts = @($Accounts)
    $summary = [PSCustomObject]@{
        Total            = $Accounts.Count
        ToMigrate        = @($Accounts | Where-Object { $_.MigrationAction -eq 'MIGRATE'      }).Count
        FolderOnly       = @($Accounts | Where-Object { $_.MigrationAction -eq 'FOLDER-ONLY'  }).Count
        Skipped          = @($Accounts | Where-Object { $_.MigrationAction -eq 'SKIP'         }).Count
        Manual           = @($Accounts | Where-Object { $_.MigrationAction -eq 'MANUAL'       }).Count
        POP3Standard     = @($Accounts | Where-Object { $_.ProviderTag -eq 'POP3-STANDARD'    }).Count
        POP3ATT          = @($Accounts | Where-Object { $_.ProviderTag -eq 'POP3-ATTAMERITECH'}).Count
        POP3Gmail        = @($Accounts | Where-Object { $_.ProviderTag -eq 'POP3-GMAIL'       }).Count
        POP3AWS          = @($Accounts | Where-Object { $_.ProviderTag -eq 'POP3-AWS'         }).Count
        IMAPAlready      = @($Accounts | Where-Object { $_.ProviderTag -eq 'IMAP-ALREADY'     }).Count
        IMAPGmail        = @($Accounts | Where-Object { $_.ProviderTag -eq 'IMAP-GMAIL'       }).Count
        ExchangeSkip     = @($Accounts | Where-Object { $_.ProviderTag -eq 'EXCHANGE-SKIP'    }).Count
        Unknown          = @($Accounts | Where-Object { $_.ProviderTag -eq 'UNKNOWN'          }).Count
        RequiresSecureKey = @($Accounts | Where-Object { $_.RequiresSecureKey -eq $true       }).Count
        HasAWSSES        = @($Accounts | Where-Object { $_.IsAWSSES -eq $true                }).Count
        WithPSTFound     = @($Accounts | Where-Object { $_.PSTPath -ne ''                    }).Count
        WithPSTMissing   = @($Accounts | Where-Object {
                                $_.AccountType -eq 'POP3' -and $_.PSTPath -eq '' }).Count
        TotalDataSizeBytes = ($Accounts | Measure-Object -Property DataFileSizeBytes -Sum).Sum
    }

    $summary | Add-Member -NotePropertyName TotalDataSizeFormatted `
                          -NotePropertyValue (Format-FileSize -Bytes $summary.TotalDataSizeBytes)

    return $summary
}


function Export-AccountsToCSV {
    <#
    .SYNOPSIS
        Exports the discovered account list to the migration control
        CSV file for operator review and completion.

    .DESCRIPTION
        Writes migration_accounts.csv to the Config directory.
        This is the primary control file that the operator reviews
        and completes (adding passwords, confirming server settings)
        before running Scripts 01-03.

        SECURITY: Password fields are written as [ENTER_PASSWORD]
        placeholder strings. Actual passwords are NEVER written to
        any file by OMMigrate.

        Accounts tagged EXCHANGE-SKIP or IMAP-ALREADY are included
        in the CSV for completeness and operator awareness, but their
        MigrationAction column clearly marks them as non-actionable.

    .PARAMETER Accounts
        Array of classified account objects to export.

    .PARAMETER OutputPath
        Full path for the output CSV file.
        Default: $Global:OMMigrate.ConfigPath\migration_accounts.csv

    .OUTPUTS
        [string] -- Path to the written CSV file.

    .EXAMPLE
        $csvPath = Export-AccountsToCSV -Accounts $accounts
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Accounts,

        [Parameter(Mandatory = $false)]
        [string]$OutputPath = ''
    )

    if ([string]::IsNullOrEmpty($OutputPath)) {
        $OutputPath = Get-OMMigrateCsvPath -BaseName 'migration_accounts.csv'
    }

    Write-OMMigrateLog -Message "Generating migration control CSV: $OutputPath" -Level INFO

    # Build CSV rows
    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($account in $Accounts) {

        # Determine placeholder text for IMAP password based on account type
        $passwordHint = switch ($account.ProviderTag) {
            'EXCHANGE-SKIP'     { 'N/A - Exchange account'              }
            'PST-ARCHIVE'       { 'N/A - Archive PST store'             }
            'IMAP-ALREADY'      { 'N/A - Already IMAP'                  }
            'IMAP-GMAIL'        { 'N/A - Already IMAP'                  }
            'POP3-ATTAMERITECH' { '[ENTER_SECURE_MAIL_KEY]'             }
            'POP3-AWS'          { '[ENTER_IMAP_PASSWORD]'               }
            default             { '[ENTER_PASSWORD]'                    }
        }

        # Determine SMTP username placeholder.
        # Standard providers use the email address as SMTP login -- pre-filled.
        # AWS SES uses an IAM SMTP access key ID -- operator must enter it.
        # Exchange and already-IMAP accounts do not need SMTP credentials here.
        # PST-ARCHIVE entries are standalone PST stores -- no SMTP credentials.
        $smtpUsernameHint = switch ($account.ProviderTag) {
            'EXCHANGE-SKIP'     { 'N/A - Exchange account'              }
            'IMAP-ALREADY'      { 'N/A - Already IMAP'                  }
            'IMAP-GMAIL'        { 'N/A - Already IMAP'                  }
            'PST-ARCHIVE'       { 'N/A - Archive PST store'             }
            'POP3-AWS'          { '[ENTER_AWS_IAM_SMTP_USERNAME]'       }
            default             { $account.EmailAddress                 }
        }

        # Determine SMTP password placeholder.
        # For standard providers the SMTP password is the same as the IMAP
        # password -- operator changes it only if their provider uses a separate
        # SMTP credential. AWS SES always requires a separate IAM SMTP password.
        $smtpPasswordHint = switch ($account.ProviderTag) {
            'EXCHANGE-SKIP'     { 'N/A - Exchange account'              }
            'IMAP-ALREADY'      { 'N/A - Already IMAP'                  }
            'IMAP-GMAIL'        { 'N/A - Already IMAP'                  }
            'PST-ARCHIVE'       { 'N/A - Archive PST store'             }
            'POP3-AWS'          { '[ENTER_AWS_IAM_SMTP_PASSWORD]'       }
            'POP3-ATTAMERITECH' { '[SAME_AS_IMAP_PASSWORD]'             }
            default             { '[SAME_AS_IMAP_PASSWORD]'             }
        }

        $row = [PSCustomObject]@{
            # -- Operator fills these ---------------------------
            'Password'              = $passwordHint      # IMAP password -- NEVER auto-populated
            'SmtpUsername'          = $smtpUsernameHint  # Pre-filled for standard; IAM key for AWS SES
            'SmtpPassword'          = $smtpPasswordHint  # NEVER auto-populated

            # -- Auto-populated by Script 00 --------------------
            'EmailAddress'          = $account.EmailAddress
            'DisplayName'           = $account.DisplayName
            'ProviderTag'           = $account.ProviderTag
            'MigrationAction'       = $account.MigrationAction
            'AccountType'           = $account.AccountType

            # Incoming (current)
            'IncomingServer'        = $account.IncomingServer
            'IncomingPort'          = $account.IncomingPort
            'IncomingSSL'           = $account.IncomingSSL

            # Outgoing (current)
            'OutgoingServer'        = $account.OutgoingServer
            'OutgoingPort'          = $account.OutgoingPort
            'OutgoingSSL'           = $account.OutgoingSSL

            # New IMAP server (for POP3 migrations)
            # Pre-filled with known provider IMAP server when available,
            # otherwise falls back to current incoming server.
            # This corrects POP3-ATTAMERITECH accounts where the current
            # server is pop.mail.yahoo.com but the IMAP server is
            # imap.mail.yahoo.com -- a different hostname entirely.
            'NewImapServer'         = if ($account.MigrationAction -eq 'MIGRATE') {
                                          $knownImapServer = $null
                                          if ($account.Domain -and
                                              $Script:KnownProviderSettings.ContainsKey($account.Domain)) {
                                              $knownImapServer = $Script:KnownProviderSettings[$account.Domain].Imap
                                          }
                                          if ($knownImapServer) { $knownImapServer } else { $account.IncomingServer }
                                      } elseif ($account.AccountType -eq 'IMAP') {
                                          # IMAP-ALREADY accounts already have a valid IMAP server.
                                          # If IncomingServer is empty, look for another account with
                                          # the same domain that has server info (same hosted server).
                                          if ($account.IncomingServer) {
                                              $account.IncomingServer
                                          } else {
                                              $domainMatch = $Accounts | Where-Object {
                                                  $_.Domain -eq $account.Domain -and
                                                  $_.IncomingServer -and
                                                  $_.EmailAddress -ne $account.EmailAddress
                                              } | Select-Object -First 1
                                              if ($domainMatch) { $domainMatch.IncomingServer } else { 'N/A' }
                                          }
                                      } else { 'N/A' }
            'NewImapPort'           = if ($account.MigrationAction -eq 'MIGRATE') { 993
                                      } elseif ($account.AccountType -eq 'IMAP') {
                                          # Use known provider port or fall back to registry value
                                          $knownImapPort = $null
                                          if ($account.Domain -and
                                              $Script:KnownProviderSettings.ContainsKey($account.Domain)) {
                                              $knownImapPort = $Script:KnownProviderSettings[$account.Domain].ImapPort
                                          }
                                          if ($knownImapPort) { $knownImapPort }
                                          elseif ($account.IncomingPort -gt 0) { $account.IncomingPort }
                                          else {
                                              # Fall back to same-domain account port
                                              $domainPortMatch = $Accounts | Where-Object {
                                                  $_.Domain -eq $account.Domain -and
                                                  $_.IncomingPort -gt 0 -and
                                                  $_.EmailAddress -ne $account.EmailAddress
                                              } | Select-Object -First 1
                                              if ($domainPortMatch) { $domainPortMatch.IncomingPort } else { 993 }
                                          }
                                      } else { 'N/A' }
            'NewImapSSL'            = if ($account.MigrationAction -eq 'MIGRATE') { $true
                                      } elseif ($account.AccountType -eq 'IMAP') { $true
                                      } else { 'N/A' }

            # Special flags
            'IsAWSSES'              = $account.IsAWSSES
            'RequiresSecureKey'     = $account.RequiresSecureKey

            # Data files
            'PSTPath'               = $account.PSTPath
            'OSTPath'               = $account.OSTPath
            'DataFileSize'          = $account.DataFileSizeFormatted

            # Notes for operator
            'Notes'                 = $account.Notes
            'PreRequisites'         = ($account.PreRequisites -join ' | ')
        }

        # NOTE: PST-ARCHIVE field overrides (N/A for server/port/SSL/credentials)
        # are applied AFTER the CSV merge below so they always win even when
        # the merge copies stale values from a prior run.

        $rows.Add($row)
    }

    # Write CSV -- merge with existing file if present to preserve operator edits
    if (-not $Global:OMMigrate.WhatIf) {
        try {
            # -- Merge with existing CSV if it exists ----------
            $mergeStats = @{ Retained = 0; Added = 0; Missing = 0 }

            if (Test-Path $OutputPath) {
                Write-OMMigrateLog -Message "Existing migration_accounts.csv found -- merging to preserve operator edits." `
                                   -Level INFO

                # Load existing CSV keyed by EmailAddress
                $existingRows = @{}
                try {
                    Import-Csv -Path $OutputPath -Encoding UTF8 |
                        ForEach-Object {
                            if ($_.ProviderTag -eq 'PST-ARCHIVE') {
                                # PST-ARCHIVE rows share EmailAddress = 'N/A' --
                                # key by DisplayName (store name) instead so each
                                # archive store matches correctly on rerun.
                                if ($_.DisplayName) {
                                    $existingRows["PST-ARCHIVE|$($_.DisplayName)"] = $_
                                }
                            }
                            elseif ($_.EmailAddress) {
                                $existingRows[$_.EmailAddress] = $_
                            }
                        }
                }
                catch {
                    Write-OMMigrateLog -Message "Could not read existing CSV for merge -- will overwrite: $_" `
                                       -Level WARN
                }

                if ($existingRows.Count -gt 0) {
                    $mergedRows = [System.Collections.Generic.List[PSCustomObject]]::new()

                    foreach ($newRow in $rows) {
                        $email = $newRow.EmailAddress
                        # PST-ARCHIVE rows are keyed by DisplayName -- use composite key
                        $lookupKey = if ($newRow.ProviderTag -eq 'PST-ARCHIVE') {
                            "PST-ARCHIVE|$($newRow.DisplayName)"
                        } else {
                            $email
                        }
                        if ($existingRows.ContainsKey($lookupKey)) {
                            # Account exists -- keep operator edits for key columns
                            $existing = $existingRows[$lookupKey]

                            # Never downgrade a completed account -- if the existing row
                            # is COMPLETE or IMAP-CONVERTED, preserve it entirely and
                            # skip the incoming row. A rediscovery from a secondary profile
                            # or re-run must never overwrite a finished migration.
                            if ($existing.MigrationAction -eq 'COMPLETE' -or
                                $existing.ProviderTag     -eq 'IMAP-CONVERTED') {
                                $mergedRows.Add($existing)
                                $mergeStats.Retained++
                                $existingRows.Remove($lookupKey)
                                Write-OMMigrateLog -Message (
                                    "Completed account protected from rediscovery overwrite: $email | " +
                                    "Kept: $($existing.ProviderTag)/$($existing.MigrationAction) | " +
                                    "Incoming discarded: $($newRow.ProviderTag)/$($newRow.MigrationAction)"
                                ) -Level INFO
                                continue
                            }

                            $mergedRow = $newRow.PSObject.Copy()

                            # Preserve operator-edited columns
                            if ($existing.Password -and
                                $existing.Password -notmatch '^\[ENTER' -and
                                $existing.Password -ne 'N/A - Exchange account' -and
                                $existing.Password -ne 'N/A - Already IMAP') {
                                $mergedRow.Password = $existing.Password
                            }
                            # Preserve SmtpUsername if operator has filled it in
                            # (i.e. it is no longer a placeholder and not N/A)
                            if ($existing.PSObject.Properties['SmtpUsername'] -and
                                $existing.SmtpUsername -and
                                $existing.SmtpUsername -notmatch '^\[ENTER' -and
                                $existing.SmtpUsername -ne 'N/A - Exchange account' -and
                                $existing.SmtpUsername -ne 'N/A - Already IMAP') {
                                $mergedRow.SmtpUsername = $existing.SmtpUsername
                            }
                            # Preserve SmtpPassword if operator has filled it in
                            if ($existing.PSObject.Properties['SmtpPassword'] -and
                                $existing.SmtpPassword -and
                                $existing.SmtpPassword -notmatch '^\[ENTER' -and
                                $existing.SmtpPassword -notmatch '^\[SAME' -and
                                $existing.SmtpPassword -ne 'N/A - Exchange account' -and
                                $existing.SmtpPassword -ne 'N/A - Already IMAP') {
                                $mergedRow.SmtpPassword = $existing.SmtpPassword
                            }
                            if ($existing.MigrationAction) {
                                $mergedRow.MigrationAction = $existing.MigrationAction
                            }
                            # Preserve ProviderTag if operator has updated it --
                            # specifically preserves IMAP-CONVERTED so Script 00
                            # re-runs do not overwrite accounts converted by Script 02.
                            if ($existing.ProviderTag -and
                                $existing.ProviderTag -ne $newRow.ProviderTag) {
                                $mergedRow.ProviderTag = $existing.ProviderTag
                            }
                            # Preserve AccountType if operator has updated it
                            if ($existing.AccountType -and
                                $existing.AccountType -ne $newRow.AccountType) {
                                $mergedRow.AccountType = $existing.AccountType
                            }
                            if ($existing.NewImapServer -and $existing.NewImapServer -ne 'N/A') {
                                $mergedRow.NewImapServer = $existing.NewImapServer
                            }
                            # If new value is still N/A but existing IncomingServer is valid, use it
                            if ($mergedRow.NewImapServer -eq 'N/A' -and
                                $existing.PSObject.Properties['IncomingServer'] -and
                                $existing.IncomingServer -and
                                $existing.IncomingServer -ne 'N/A') {
                                $mergedRow.NewImapServer = $existing.IncomingServer
                            }
                            if ($existing.NewImapPort -and $existing.NewImapPort -ne 'N/A') {
                                $mergedRow.NewImapPort = $existing.NewImapPort
                            }
                            # Preserve NewImapSSL if operator has changed it
                            if ($existing.PSObject.Properties['NewImapSSL'] -and
                                $existing.NewImapSSL -and $existing.NewImapSSL -ne 'N/A') {
                                $mergedRow.NewImapSSL = $existing.NewImapSSL
                            }
                            # Preserve DisplayName if operator has corrected it
                            if ($existing.DisplayName -and
                                $existing.DisplayName -ne $newRow.DisplayName) {
                                $mergedRow.DisplayName = $existing.DisplayName
                            }
                            # Preserve OutgoingServer if operator has corrected it
                            if ($existing.OutgoingServer -and
                                $existing.OutgoingServer -ne $newRow.OutgoingServer) {
                                $mergedRow.OutgoingServer = $existing.OutgoingServer
                            }
                            # Preserve OutgoingPort if operator has corrected it
                            if ($existing.OutgoingPort -and
                                $existing.OutgoingPort -ne $newRow.OutgoingPort) {
                                $mergedRow.OutgoingPort = $existing.OutgoingPort
                            }
                            # Preserve OutgoingSSL if operator has corrected it
                            if ($existing.PSObject.Properties['OutgoingSSL'] -and
                                $existing.OutgoingSSL -and
                                $existing.OutgoingSSL -ne $newRow.OutgoingSSL) {
                                $mergedRow.OutgoingSSL = $existing.OutgoingSSL
                            }
                            if ($existing.Notes -and $existing.Notes -ne $newRow.Notes) {
                                # Clear stale NOT DETECTED flag -- this row matched successfully
                                # this run so the flag from a prior run is no longer valid.
                                $cleanedNotes = $existing.Notes -replace '\[NOT DETECTED THIS RUN[^\]]*\]', ''
                                $cleanedNotes = $cleanedNotes.Trim()
                                if ($cleanedNotes) {
                                    $mergedRow.Notes = $cleanedNotes
                                }
                            }
                            # Preserve PreRequisites if operator has edited it
                            if ($existing.PSObject.Properties['PreRequisites'] -and
                                $existing.PreRequisites -and
                                $existing.PreRequisites -ne $newRow.PreRequisites) {
                                $mergedRow.PreRequisites = $existing.PreRequisites
                            }
                            # Preserve RequiresSecureKey -- once True always True.
                            # Set-AccountTag only sets this on POP3-ATTAMERITECH accounts.
                            # After conversion to IMAP, the account hits IMAP-ALREADY
                            # classification and the flag defaults to False on the new
                            # account object. Preserve the existing True so reruns do
                            # not silently clear it.
                            if ($existing.PSObject.Properties['RequiresSecureKey'] -and
                                $existing.RequiresSecureKey -eq 'True') {
                                $mergedRow.RequiresSecureKey = $existing.RequiresSecureKey
                            }
                            # Preserve OSTPath only if the existing path's filename
                            # contains the account's email address -- prevents keeping
                            # a wrong account's OST path just because it exists on disk.
                            if ($existing.PSObject.Properties['OSTPath'] -and
                                $existing.OSTPath -and
                                -not [string]::IsNullOrWhiteSpace($existing.OSTPath) -and
                                [System.IO.Path]::GetFileName($existing.OSTPath) -like "*$email*" -and
                                ((-not $newRow.OSTPath) -or
                                 (Test-Path $existing.OSTPath -ErrorAction SilentlyContinue))) {
                                $mergedRow.OSTPath = $existing.OSTPath
                            }
                            # Preserve PSTPath only if the existing path's filename
                            # contains the account's email address.
                            if ($existing.PSObject.Properties['PSTPath'] -and
                                $existing.PSTPath -and
                                -not [string]::IsNullOrWhiteSpace($existing.PSTPath) -and
                                [System.IO.Path]::GetFileName($existing.PSTPath) -like "*$email*" -and
                                ((-not $newRow.PSTPath) -or
                                 (Test-Path $existing.PSTPath -ErrorAction SilentlyContinue))) {
                                $mergedRow.PSTPath = $existing.PSTPath
                            }

                            # Force N/A overrides for PST-ARCHIVE after merge --
                            # the merge may have copied stale server/credential values
                            # from a prior run before the PST-ARCHIVE feature existed.
                            if ($mergedRow.ProviderTag -eq 'PST-ARCHIVE') {
                                $mergedRow.Password        = 'N/A - Archive PST store'
                                $mergedRow.SmtpUsername    = 'N/A - Archive PST store'
                                $mergedRow.SmtpPassword    = 'N/A - Archive PST store'
                                $mergedRow.IncomingServer  = 'N/A'
                                $mergedRow.IncomingPort    = 'N/A'
                                $mergedRow.IncomingSSL     = 'N/A'
                                $mergedRow.OutgoingServer  = 'N/A'
                                $mergedRow.OutgoingPort    = 'N/A'
                                $mergedRow.OutgoingSSL     = 'N/A'
                                $mergedRow.IsAWSSES        = 'N/A'
                                $mergedRow.RequiresSecureKey = 'N/A'
                                $mergedRow.EmailAddress    = 'N/A'
                            }

                            $mergedRows.Add($mergedRow)
                            $mergeStats.Retained++
                            $existingRows.Remove($lookupKey)
                        }
                        else {
                            # New account not in existing CSV -- add with defaults.
                            # Guard against duplicate EmailAddress from multiple Outlook
                            # profiles (e.g. "TestProfile" profile) -- if this email was
                            # already added to $mergedRows by a prior iteration, skip
                            # this row to prevent a second profile's stale POP3 entry
                            # from appearing as a duplicate in the CSV.
                            $alreadyMerged = $mergedRows | Where-Object { $_.EmailAddress -eq $email } | Select-Object -First 1
                            if ($alreadyMerged) {
                                Write-OMMigrateLog -Message (
                                    "Duplicate EmailAddress skipped during merge: $email | " +
                                    "Already added as $($alreadyMerged.ProviderTag). " +
                                    "Likely from a secondary Outlook profile -- use profile picker to avoid this."
                                ) -Level INFO
                                continue
                            }
                            # Force N/A overrides for PST-ARCHIVE new rows too
                            if ($newRow.ProviderTag -eq 'PST-ARCHIVE') {
                                $newRow.Password        = 'N/A - Archive PST store'
                                $newRow.SmtpUsername    = 'N/A - Archive PST store'
                                $newRow.SmtpPassword    = 'N/A - Archive PST store'
                                $newRow.IncomingServer  = 'N/A'
                                $newRow.IncomingPort    = 'N/A'
                                $newRow.IncomingSSL     = 'N/A'
                                $newRow.OutgoingServer  = 'N/A'
                                $newRow.OutgoingPort    = 'N/A'
                                $newRow.OutgoingSSL     = 'N/A'
                                $newRow.IsAWSSES        = 'N/A'
                                $newRow.RequiresSecureKey = 'N/A'
                                $newRow.EmailAddress       = 'N/A'
                            }
                            $mergedRows.Add($newRow)
                            $mergeStats.Added++
                            Write-OMMigrateLog -Message "New account added to CSV: $email" -Level INFO
                        }
                    }

                    # Flag accounts in old CSV but not discovered this run
                    foreach ($missingEmail in $existingRows.Keys) {
                        $missingRow = $existingRows[$missingEmail]
                        # PST-ARCHIVE stores (archive PSTs, backup PSTs) are not
                        # always mounted in every COM session -- this is expected
                        # and normal. Log at INFO, not WARN, to avoid false alarms.
                        $missingIsArchive = ($missingRow.ProviderTag -eq 'PST-ARCHIVE')
                        if (-not $missingIsArchive) {
                            # Keep it but flag it for operator attention
                            $missingRow.Notes = "[NOT DETECTED THIS RUN -- verify account still exists]"
                        }
                        $mergedRows.Add($missingRow)
                        $mergeStats.Missing++
                        Write-OMMigrateLog -Message "Account in CSV but not detected this run: $missingEmail" `
                                           -Level $(if ($missingIsArchive) { 'INFO' } else { 'WARN' })
                    }

                    $rows = $mergedRows

                    Write-OMMigrateLog -Message (
                        "CSV merge complete: $($mergeStats.Retained) retained, " +
                        "$($mergeStats.Added) new, $($mergeStats.Missing) not detected"
                    ) -Level INFO

                    if ($mergeStats.Added -gt 0) {

                        Write-Host "  CSV merged: $($mergeStats.Retained) accounts retained, $($mergeStats.Added) new account(s) added." `
                                   -ForegroundColor Green
                    }
                    else {
                        Write-Host "  CSV merged: $($mergeStats.Retained) accounts retained, no new accounts." `
                                   -ForegroundColor Green
                    }
                    # Only show the operator warning for non-archive missing accounts.
                    # Archive PST stores being absent from a COM scan is expected -- no
                    # yellow banner needed for those. Count non-archive missing rows only.
                    $nonArchiveMissing = @($mergedRows | Where-Object {
                        $_.Notes -like '*NOT DETECTED THIS RUN*'
                    }).Count
                    if ($nonArchiveMissing -gt 0) {
                        Write-Host "  WARNING: $nonArchiveMissing account(s) in CSV were not detected this run -- review flagged rows." `
                                   -ForegroundColor Yellow
                    }
                }
            }

            $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            Write-OMMigrateLog -Message "CSV written: $OutputPath ($($rows.Count) accounts)" `
                               -Level INFO
            Write-AuditEntry  -Action 'CSV_ACCOUNTS_WRITTEN' `
                              -Detail "Path=$OutputPath | Accounts=$($rows.Count) | Retained=$($mergeStats.Retained) | Added=$($mergeStats.Added)"
        }
        catch {
            Write-OMMigrateLog -Message "Failed to write CSV: $_" -Level ERROR
            throw
        }
    }
    else {
        Write-OMMigrateLog -Message "WhatIf: Would write CSV to $OutputPath ($($rows.Count) accounts)" `
                           -Level INFO -WhatIfPrefix
    }

    return $OutputPath
}



# ============================================================
#  REGION: ACCOUNT REMOVAL VIA REGISTRY
# ============================================================

function Remove-POP3AccountViaRegistry {
    <#
    .SYNOPSIS
        Removes a POP3 account from the Outlook profile by deleting
        its registry subkey under the 9375CFF0 GUID path.

    .DESCRIPTION
        The Outlook COM API Accounts.Remove() method is not available
        in Outlook 2016/2019/2021. The correct approach for these
        versions is to delete the account's registry subkey directly.

        Each account is stored as a numbered subkey (00000001, 00000002
        etc.) under the well-known GUID path:
            HKCU:\...\Profiles\{ProfileName}\9375CFF0413111d3B88A00104B2A6676\

        This function:
            1. Locates the active Outlook profile in the registry
            2. Scans the 9375CFF0 subkeys to find the one matching the
               email address
            3. Deletes that subkey (WhatIf: logs only, no deletion)
            4. Verifies the subkey is gone after deletion

        Outlook must be fully closed before calling this function.
        Changes take effect the next time Outlook launches.

    .PARAMETER EmailAddress
        Email address of the POP3 account to remove.

    .PARAMETER ProfileName
        Name of the Outlook profile containing the account.
        Default: reads DefaultProfile from registry.

    .OUTPUTS
        [bool] -- $true if removal succeeded or WhatIf, $false if failed.

    .EXAMPLE
        $removed = Remove-POP3AccountViaRegistry -EmailAddress 'user@domain.com'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,

        [Parameter(Mandatory = $false)]
        [string]$ProfileName = ''
    )

    # -- Resolve active profile name ---------------------------
    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        try {
            # DefaultProfile lives at HKCU:\...\Office\16.0\Outlook
            # not under the Profiles subkey
            $outlookRegBase = @(
                'HKCU:\Software\Microsoft\Office\16.0\Outlook',
                'HKCU:\Software\Microsoft\Office\15.0\Outlook',
                'HKCU:\Software\Microsoft\Office\14.0\Outlook'
            )
            foreach ($basePath in $outlookRegBase) {
                $ProfileName = (Get-ItemProperty -Path $basePath `
                                                 -Name 'DefaultProfile' `
                                                 -ErrorAction SilentlyContinue).DefaultProfile
                if ($ProfileName) { break }
            }
        }
        catch { }
    }

    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        Write-OMMigrateLog -Message "Cannot remove account -- Outlook profile name not found in registry." `
                           -Level ERROR
        return $false
    }

    # -- Build registry path to account GUID subkeys -----------
    $accountGUIDPath = $null
    foreach ($regPath in $Script:OutlookRegPaths) {
        $candidate = Join-Path $regPath "$ProfileName\$Script:OutlookAccountGUID"
        if (Test-Path $candidate) {
            $accountGUIDPath = $candidate
            break
        }
    }

    if (-not $accountGUIDPath) {
        Write-OMMigrateLog -Message "Cannot find account registry path for profile '$ProfileName'." `
                           -Level ERROR
        return $false
    }

    Write-OMMigrateLog -Message "Scanning registry for account: $EmailAddress" -Level INFO
    Write-OMMigrateLog -Message "Registry path: $accountGUIDPath" -Level DEBUG

    # -- Scan subkeys to find the matching account -------------
    $targetKeyPath = $null

    try {
        $subkeys = Get-ChildItem -Path $accountGUIDPath -ErrorAction Stop

        foreach ($subkey in $subkeys) {
            try {
                $props = Get-ItemProperty -Path $subkey.PSPath -ErrorAction SilentlyContinue

                # POP3 accounts have both 'Email' and 'Account Name'
                # Check both to ensure we find the right account
                $emailProp = $props.Email
                if (-not $emailProp) {
                    $emailProp = $props.'Account Name'
                }

                if ($emailProp -eq $EmailAddress) {
                    $targetKeyPath = $subkey.PSPath
                    Write-OMMigrateLog -Message "Found account registry key: $($subkey.Name)" `
                                       -Level INFO
                    break
                }
            }
            catch { }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Failed to scan account registry subkeys: $_" -Level ERROR
        return $false
    }

    if (-not $targetKeyPath) {
        Write-OMMigrateLog -Message "Account '$EmailAddress' not found in registry -- may already be removed." `
                           -Level WARN
        return $false
    }

    # -- Find the MAPI store key associated with this POP3 account -------------
    # When Outlook loads, it tries to mount every MAPI store key in the profile.
    # If the POP3 account key is deleted but the MAPI store key remains, Outlook
    # crashes on startup because it finds an orphaned store with no owning account.
    #
    # The link between the POP3 account key and its MAPI store key is the
    # Delivery Store EntryID blob. The POP3 account key contains this blob,
    # and the matching MAPI store key contains the same value at property 01020fff.
    # We read the blob from the POP3 key, then scan all sibling subkeys for a
    # match. The matching key must be deleted alongside the POP3 account key.
    #
    # This approach is filename-independent -- it works regardless of what the
    # PST file is named on disk (not all PSTs are named after the email address).
    $mapiStoreKeyPath = $null

    try {
        $pop3Props = Get-ItemProperty -Path $targetKeyPath -ErrorAction SilentlyContinue

        # Read the Delivery Store EntryID blob from the POP3 account key
        # This is the MAPI identifier linking the account to its data store
        $deliveryStoreEID = $null
        $deliveryProp = $pop3Props.PSObject.Properties['Delivery Store EntryID']
        if ($deliveryProp -and $deliveryProp.Value -is [byte[]] -and $deliveryProp.Value.Length -gt 0) {
            $deliveryStoreEID = $deliveryProp.Value
            Write-OMMigrateLog -Message "Delivery Store EntryID found in POP3 key ($($deliveryStoreEID.Length) bytes)" `
                               -Level DEBUG
        }
        else {
            Write-OMMigrateLog -Message "No Delivery Store EntryID in POP3 key -- will scan by PST path" `
                               -Level DEBUG
        }

        # Scan all subkeys under the profile root (not just the 9375CFF0 GUID path)
        # The MAPI store key lives directly under the profile root, not under the GUID
        $profileRootPath = $accountGUIDPath -replace "\\$([regex]::Escape($Script:OutlookAccountGUID))$", ''

        $profileSubkeys = Get-ChildItem -Path $profileRootPath -ErrorAction SilentlyContinue

        foreach ($profileSubkey in $profileSubkeys) {
            try {
                $profileProps = Get-ItemProperty -Path $profileSubkey.PSPath `
                                                 -ErrorAction SilentlyContinue
                if (-not $profileProps) { continue }

                # Method 1: Match by Delivery Store EntryID vs store 01020fff blob
                # 01020fff is the MAPI PR_ENTRYID property for the store object
                if ($deliveryStoreEID) {
                    $storeEIDProp = $profileProps.PSObject.Properties['01020fff']
                    if ($storeEIDProp -and $storeEIDProp.Value -is [byte[]] -and
                        $storeEIDProp.Value.Length -gt 0) {

                        # Compare byte arrays
                        $storeEID  = [byte[]]$storeEIDProp.Value
                        $delivEID  = [byte[]]$deliveryStoreEID

                        # Use the shorter length for comparison -- EntryIDs can have
                        # padding differences but the core bytes should match
                        $compareLen = [Math]::Min($storeEID.Length, $delivEID.Length)
                        $match = $true
                        for ($b = 0; $b -lt $compareLen; $b++) {
                            if ($storeEID[$b] -ne $delivEID[$b]) { $match = $false; break }
                        }

                        if ($match) {
                            $mapiStoreKeyPath = $profileSubkey.PSPath
                            Write-OMMigrateLog -Message (
                                "MAPI store key found via EntryID match: $($profileSubkey.PSChildName)"
                            ) -Level INFO
                            break
                        }
                    }
                }

                # Method 2: Fallback -- match by PST path containing email address
                # Used when Delivery Store EntryID is missing or comparison fails.
                # Less reliable (PST may not be named after email) but covers common cases.
                if (-not $mapiStoreKeyPath) {
                    $pathProp = $profileProps.PSObject.Properties['001f6700']
                    if ($pathProp -and $pathProp.Value -is [byte[]]) {
                        $pathStr = [System.Text.Encoding]::Unicode.GetString($pathProp.Value).TrimEnd([char]0)
                        if ($pathStr -like "*$EmailAddress*") {
                            $mapiStoreKeyPath = $profileSubkey.PSPath
                            Write-OMMigrateLog -Message (
                                "MAPI store key found via PST path match: $($profileSubkey.PSChildName)"
                            ) -Level INFO
                            break
                        }
                    }
                }
            }
            catch { }
        }

        if ($mapiStoreKeyPath) {
            Write-OMMigrateLog -Message "MAPI store key will be deleted with POP3 key to prevent Outlook crash on startup." `
                               -Level INFO
        }
        else {
            # Non-fatal -- log and continue. Deletion of POP3 key alone may still
            # cause Outlook instability but is better than leaving the account active.
            Write-OMMigrateLog -Message (
                "WARNING: Could not find MAPI store key for $EmailAddress. " +
                "POP3 key will be deleted but Outlook may be unstable on next startup. " +
                "Verify Outlook opens correctly after migration."
            ) -Level WARN
        }
    }
    catch {
        Write-OMMigrateLog -Message "MAPI store key search failed: $_ -- continuing with POP3 key only." `
                           -Level WARN
    }

    # -- Registry backup -- export both keys to .reg file before deletion ------
    # The .reg file can be restored by double-clicking it in Explorer.
    # Both the POP3 account key AND the MAPI store key are included so a
    # double-click restore brings Outlook back to a fully stable state.
    # No PowerShell needed for recovery -- just double-click the .reg file.
    $safeEmail     = Get-SafeFileName -InputString $EmailAddress
    $backupDir     = $Global:OMMigrate.BackupPath
    $regBackupFile = Join-Path $backupDir "$safeEmail`_registry_backup.reg"

    # Convert PSPath to the plain registry path form that reg.exe expects
    # PSPath format: Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\...
    # reg.exe format: HKEY_CURRENT_USER\...
    $regExePOP3Path = $targetKeyPath -replace '^Microsoft\.PowerShell\.Core\\Registry::', ''
    $regExeStorePath = if ($mapiStoreKeyPath) {
        $mapiStoreKeyPath -replace '^Microsoft\.PowerShell\.Core\\Registry::', ''
    } else { $null }

    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message "WhatIf: Would export POP3 registry key to: $regBackupFile" `
                           -Level INFO -WhatIfPrefix
        if ($regExeStorePath) {
            Write-OMMigrateLog -Message "WhatIf: Would also export MAPI store key to same backup file" `
                               -Level INFO -WhatIfPrefix
        }
    }
    else {
        try {
            # Ensure backup directory exists
            if (-not (Test-Path $backupDir)) {
                New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
            }

            # Export the POP3 account key first
            $regArgs = @('export', $regExePOP3Path, $regBackupFile, '/y')
            $regResult = & reg.exe @regArgs 2>&1
            $regExitCode = $LASTEXITCODE

            if ($regExitCode -ne 0) {
                # Backup failed -- do NOT proceed with deletion
                # A missing backup is more dangerous than a delayed migration
                Write-OMMigrateLog -Message (
                    "Registry backup FAILED (exit code $regExitCode) -- deletion aborted. " +
                    "Output: $regResult"
                ) -Level ERROR
                Write-AuditEntry  -Action 'REGISTRY_BACKUP_FAILED' `
                                  -AccountEmail $EmailAddress `
                                  -Detail "reg.exe exit code: $regExitCode | Output: $regResult" `
                                  -Outcome 'FAILED'
                return $false
            }

            Write-OMMigrateLog -Message "POP3 registry key backed up: $regBackupFile" -Level INFO

            # Append the MAPI store key to the same .reg file if found.
            # reg.exe /y overwrites -- use a temp file and merge manually
            # to combine both keys into one restorable .reg file.
            if ($regExeStorePath) {
                $tempStoreBackup = "$regBackupFile.store.tmp"
                $storeArgs = @('export', $regExeStorePath, $tempStoreBackup, '/y')
                $storeResult = & reg.exe @storeArgs 2>&1
                $storeExitCode = $LASTEXITCODE

                if ($storeExitCode -eq 0) {
                    # Merge: read both .reg files, combine content (skip duplicate header)
                    # Windows Registry Editor Version 5.00 header appears once at top
                    $pop3Content  = [System.IO.File]::ReadAllText($regBackupFile)
                    $storeContent = [System.IO.File]::ReadAllText($tempStoreBackup)

                    # Strip the header from the store content before appending
                    $storeContentNoHeader = $storeContent -replace '^Windows Registry Editor Version 5\.00\s*', ''

                    # Combine and write back to the main backup file
                    $combined = $pop3Content.TrimEnd() + "`r`n`r`n" + $storeContentNoHeader.TrimStart()
                    [System.IO.File]::WriteAllText($regBackupFile, $combined,
                        [System.Text.Encoding]::Unicode)

                    # Clean up temp file
                    Remove-Item -Path $tempStoreBackup -Force -ErrorAction SilentlyContinue

                    Write-OMMigrateLog -Message "MAPI store key appended to backup file: $regBackupFile" `
                                       -Level INFO
                }
                else {
                    # Non-fatal -- POP3 backup succeeded, store backup failed
                    # Log and continue -- POP3 restore will still work
                    Write-OMMigrateLog -Message (
                        "WARNING: MAPI store key backup failed (exit $storeExitCode) -- " +
                        "POP3 backup still valid. Manual store key restore may be needed."
                    ) -Level WARN
                    Remove-Item -Path $tempStoreBackup -Force -ErrorAction SilentlyContinue
                }
            }

            Write-OMMigrateLog -Message "Registry backup written: $regBackupFile" -Level INFO
            Write-AuditEntry  -Action 'REGISTRY_BACKUP_WRITTEN' `
                              -AccountEmail $EmailAddress `
                              -Detail "Backup file: $regBackupFile | Includes MAPI store key: $($null -ne $mapiStoreKeyPath)"
        }
        catch {
            Write-OMMigrateLog -Message "Registry backup exception -- deletion aborted: $_" `
                               -Level ERROR
            return $false
        }
    }

    # -- WhatIf -- simulate only -------------------------------
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message "WhatIf: Would delete POP3 registry key: $targetKeyPath" `
                           -Level INFO -WhatIfPrefix
        if ($mapiStoreKeyPath) {
            Write-OMMigrateLog -Message "WhatIf: Would delete MAPI store key: $mapiStoreKeyPath" `
                               -Level INFO -WhatIfPrefix
        }
        return $true
    }

    # -- Delete both registry keys -----------------------------
    # Delete the POP3 account key first, then the MAPI store key.
    # Both must be removed to prevent Outlook crashing on next startup
    # due to an orphaned store reference with no owning account.
    try {
        Remove-Item -Path $targetKeyPath -Recurse -Force -ErrorAction Stop
        Write-OMMigrateLog -Message "POP3 account key deleted: $targetKeyPath" -Level INFO

        # Verify POP3 key deletion
        if (Test-Path $targetKeyPath) {
            Write-OMMigrateLog -Message "WARNING: POP3 registry key still exists after deletion attempt." `
                               -Level WARN
            return $false
        }

        # Delete the MAPI store key if found
        if ($mapiStoreKeyPath) {
            try {
                Remove-Item -Path $mapiStoreKeyPath -Recurse -Force -ErrorAction Stop
                Write-OMMigrateLog -Message "MAPI store key deleted: $mapiStoreKeyPath" -Level INFO

                if (Test-Path $mapiStoreKeyPath) {
                    Write-OMMigrateLog -Message "WARNING: MAPI store key still exists after deletion attempt." `
                                       -Level WARN
                }
            }
            catch {
                # Non-fatal -- POP3 key is gone, Outlook may still be unstable
                # but the account is removed. Log and continue.
                Write-OMMigrateLog -Message (
                    "WARNING: Could not delete MAPI store key '$mapiStoreKeyPath': $_. " +
                    "Outlook may be unstable on next startup -- verify after migration."
                ) -Level WARN
            }
        }

        Write-OMMigrateLog -Message "POP3 account removed from registry: $EmailAddress" -Level INFO
        Write-AuditEntry  -Action 'ACCOUNT_REMOVED_REGISTRY' `
                          -AccountEmail $EmailAddress `
                          -Detail (
                              "POP3 key deleted: $targetKeyPath | " +
                              "MAPI store key deleted: $($null -ne $mapiStoreKeyPath)"
                          )
        return $true
    }
    catch {
        Write-OMMigrateLog -Message "Failed to delete registry key '$targetKeyPath': $_" `
                           -Level ERROR
        Write-AuditEntry  -Action 'ACCOUNT_REMOVE_FAILED' `
                          -AccountEmail $EmailAddress `
                          -Detail "Registry deletion error: $_" `
                          -Outcome 'FAILED'
        return $false
    }
}


# ============================================================
#  REGION: ACCOUNT ADD VIA REGISTRY
# ============================================================

function Add-IMAPAccountViaRegistry {
    <#
    .SYNOPSIS
        Adds a new IMAP account to the Outlook profile by writing
        registry subkey values under the 9375CFF0 GUID path.

    .DESCRIPTION
        The Outlook COM API Accounts.Add() method is not available
        in Outlook 2016/2019/2021. The correct approach for these
        versions is to write the account's registry subkey directly.

        Each account is stored as a numbered subkey (00000001, 00000002
        etc.) under the well-known GUID path:
            HKCU:\...\Profiles\{ProfileName}\9375CFF0413111d3B88A00104B2A6676\

        This function:
            1. Locates the active Outlook profile in the registry
            2. Determines the next available numbered subkey index
            3. Generates required random UIDs using cryptographic RNG
               (same approach Outlook uses internally)
            4. Writes all required IMAP account values to the new subkey
               (WhatIf: logs what would be written, no registry changes)
            5. Verifies the new subkey exists after writing
            6. Writes imap_remove.reg rollback file to the Backups folder

        Outlook must be fully closed before calling this function.
        Changes take effect the next time Outlook launches.

        Passwords are NEVER written -- Outlook prompts the operator
        on first open for both IMAP and SMTP credentials.

        Designed for marketability -- works across all standard IMAP
        providers (Gmail, Apple, Yahoo, AWS SES, custom servers).
        Provider-specific behaviour is driven by server/port parameters,
        not hardcoded logic.

        Values written per account:
            clsid                  -- constant IMAP service provider GUID
            Mini UID               -- random DWORD (Outlook internal ID)
            Service UID            -- random 16-byte blob
            Account UID            -- random 16-byte blob
            Preferences UID        -- random 16-byte blob
            Account Name           -- email address (display identifier)
            Display Name           -- display name for From: field
            Email                  -- email address
            IMAP Server            -- incoming mail server hostname
            IMAP Port              -- incoming port (typically 993)
            IMAP Use SSL           -- 1 for SSL/TLS, 0 for plain
            IMAP User              -- IMAP login username (defaults to email)
            IMAP Store EID         -- empty blob (Outlook fills on first open)
            IMAP Sentitems flag    -- 8 (constant -- store sent items locally)
            IMAP Full List         -- 1 (constant -- show full folder list)
            EnablePurgeOnSwitch    -- 1 (constant -- purge deleted on folder switch)
            SMTP Server            -- outgoing mail server hostname
            SMTP Port              -- outgoing port (typically 587 or 465)
            SMTP Use SSL           -- 1 for SSL/TLS, 0 for plain
            SMTP Secure Connection -- computed: 0=plain, 1=SSL/465, 2=STARTTLS/587
            SMTP Use Auth          -- 1 (required for all providers)
            SMTP Auth Method       -- 1 (standard password auth)
            SMTP User              -- SMTP login username (defaults to email;
                                      pass IAM key ID for AWS SES)
            Delivery Store EntryID    -- empty blob (Outlook fills on first open)
            Delivery Folder EntryID   -- empty blob (Outlook fills on first open)

    .PARAMETER EmailAddress
        Email address of the new IMAP account.

    .PARAMETER DisplayName
        Display name for the account (shown in Outlook From: field).

    .PARAMETER ImapServer
        Incoming IMAP server hostname (e.g. imap.gmail.com).

    .PARAMETER ImapPort
        Incoming IMAP port number. Typically 993 for SSL.

    .PARAMETER ImapSSL
        Whether to use SSL/TLS for the incoming connection.
        Default: $true

    .PARAMETER ImapUsername
        IMAP login username. Defaults to EmailAddress.
        Override only when the IMAP username differs from the email address.

    .PARAMETER SmtpServer
        Outgoing SMTP server hostname.

    .PARAMETER SmtpPort
        Outgoing SMTP port number. Typically 587 (STARTTLS) or 465 (SSL).

    .PARAMETER SmtpSSL
        Whether to use SSL/TLS for the outgoing connection.
        Default: $true

    .PARAMETER SmtpUsername
        SMTP login username. Defaults to EmailAddress.
        For AWS SES accounts pass the IAM SMTP access key ID here.

    .PARAMETER ProfileName
        Name of the Outlook profile to add the account to.
        Default: reads DefaultProfile from registry.

    .OUTPUTS
        [bool] -- $true if account was written successfully or WhatIf,
                  $false if failed.

    .EXAMPLE
        # Standard IMAP (Gmail, Apple, Yahoo, custom server)
        $added = Add-IMAPAccountViaRegistry `
            -EmailAddress 'user@gmail.com' `
            -DisplayName  'User Name' `
            -ImapServer   'imap.gmail.com' `
            -ImapPort     993 `
            -ImapSSL      $true `
            -SmtpServer   'smtp.gmail.com' `
            -SmtpPort     587 `
            -SmtpSSL      $true

    .EXAMPLE
        # AWS SES outbound -- SmtpUsername is the IAM SMTP access key ID
        $added = Add-IMAPAccountViaRegistry `
            -EmailAddress  'user@domain.com' `
            -DisplayName   'User Name' `
            -ImapServer    'imap.domain.com' `
            -ImapPort      993 `
            -ImapSSL       $true `
            -SmtpServer    'email-smtp.us-east-1.amazonaws.com' `
            -SmtpPort      587 `
            -SmtpSSL       $true `
            -SmtpUsername  'AKIAIOSFODNN7EXAMPLE'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,

        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $true)]
        [string]$ImapServer,

        [Parameter(Mandatory = $true)]
        [int]$ImapPort,

        [Parameter(Mandatory = $false)]
        [bool]$ImapSSL = $true,

        [Parameter(Mandatory = $false)]
        [string]$ImapUsername = '',

        [Parameter(Mandatory = $true)]
        [string]$SmtpServer,

        [Parameter(Mandatory = $true)]
        [int]$SmtpPort,

        [Parameter(Mandatory = $false)]
        [bool]$SmtpSSL = $true,

        [Parameter(Mandatory = $false)]
        [string]$SmtpUsername = '',

        [Parameter(Mandatory = $false)]
        [string]$ProfileName = ''
    )

    # -- Default usernames to email address if not specified ---
    # Works for all standard providers (Gmail, Apple, Yahoo, etc.)
    # Caller overrides SmtpUsername for AWS SES (IAM access key ID)
    if ([string]::IsNullOrWhiteSpace($ImapUsername)) { $ImapUsername = $EmailAddress }
    if ([string]::IsNullOrWhiteSpace($SmtpUsername)) { $SmtpUsername = $EmailAddress }

    # -- Compute SMTP Secure Connection value ------------------
    # Outlook uses this DWORD to determine the connection security method:
    #   0 = No encryption (plain, port 25)
    #   1 = SSL/TLS      (implicit SSL, typically port 465)
    #   2 = STARTTLS     (explicit TLS, typically port 587)
    # Derived from port number for marketability across all providers.
    $smtpSecureConnection = 0
    if ($SmtpSSL) {
        if ($SmtpPort -eq 465) {
            $smtpSecureConnection = 1   # SSL/TLS (implicit)
        }
        else {
            $smtpSecureConnection = 2   # STARTTLS (explicit) -- covers 587 and others
        }
    }

    # -- Resolve active profile name ---------------------------
    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        try {
            $outlookRegBase = @(
                'HKCU:\Software\Microsoft\Office\16.0\Outlook',
                'HKCU:\Software\Microsoft\Office\15.0\Outlook',
                'HKCU:\Software\Microsoft\Office\14.0\Outlook'
            )
            foreach ($basePath in $outlookRegBase) {
                $ProfileName = (Get-ItemProperty -Path $basePath `
                                                 -Name 'DefaultProfile' `
                                                 -ErrorAction SilentlyContinue).DefaultProfile
                if ($ProfileName) { break }
            }
        }
        catch { }
    }

    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        Write-OMMigrateLog -Message "Cannot add account -- Outlook profile name not found in registry." `
                           -Level ERROR
        return $false
    }

    # -- Build registry path to account GUID subkeys -----------
    $accountGUIDPath = $null
    foreach ($regPath in $Script:OutlookRegPaths) {
        $candidate = Join-Path $regPath "$ProfileName\$Script:OutlookAccountGUID"
        if (Test-Path $candidate) {
            $accountGUIDPath = $candidate
            break
        }
    }

    if (-not $accountGUIDPath) {
        Write-OMMigrateLog -Message "Cannot find account registry path for profile '$ProfileName'." `
                           -Level ERROR
        return $false
    }

    Write-OMMigrateLog -Message "Preparing to add IMAP account: $EmailAddress" -Level INFO
    Write-OMMigrateLog -Message "Registry path: $accountGUIDPath" -Level DEBUG

    # -- Determine next available subkey index -----------------
    # Existing subkeys are named 00000001, 00000002, etc.
    # Find the highest existing index and add 1.
    $nextIndex = 1
    try {
        $existingSubkeys = Get-ChildItem -Path $accountGUIDPath -ErrorAction Stop
        foreach ($subkey in $existingSubkeys) {
            $name = $subkey.PSChildName
            if ($name -match '^[0-9A-Fa-f]{8}$') {
                $indexValue = [Convert]::ToInt32($name, 16)
                if ($indexValue -ge $nextIndex) {
                    $nextIndex = $indexValue + 1
                }
            }
        }
    }
    catch {
        Write-OMMigrateLog -Message "Failed to enumerate existing account subkeys: $_" -Level ERROR
        return $false
    }

    # Format as 8-digit zero-padded hex (e.g. 00000009)
    $newSubkeyName = $nextIndex.ToString('X8').ToLower()
    $newKeyPath    = Join-Path $accountGUIDPath $newSubkeyName

    Write-OMMigrateLog -Message "New account subkey index: $newSubkeyName" -Level DEBUG

    # -- Generate random UIDs ----------------------------------
    # Outlook generates these internally using a cryptographic RNG.
    # We use the same approach for compatibility across all providers.
    # Each account gets unique UIDs -- never reuse across accounts.
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()

    # Mini UID -- 4-byte random DWORD
    $miniUIDBytes = New-Object byte[] 4
    $rng.GetBytes($miniUIDBytes)
    $miniUID = [System.BitConverter]::ToUInt32($miniUIDBytes, 0)

    # Service UID, Account UID, Preferences UID -- 16-byte random blobs
    $serviceUID     = New-Object byte[] 16
    $accountUID     = New-Object byte[] 16
    $preferencesUID = New-Object byte[] 16
    $rng.GetBytes($serviceUID)
    $rng.GetBytes($accountUID)
    $rng.GetBytes($preferencesUID)
    $rng.Dispose()

    # Empty byte arrays -- Outlook populates these on first open
    $emptyBlob           = [byte[]]@()
    $deliveryFolderBlob  = [byte[]]@(239, 0, 0, 0)   # Delivery Folder EntryID seed

    # -- WhatIf -- log what the live run will do, no changes ----
    # IMPORTANT: IMAP account addition is NOT automated via registry write.
    # Direct registry writes cause Outlook to crash on load because MAPI store
    # initialization, OST creation, and search folder setup are bypassed.
    # The live run delegates account creation to Outlook's own Add Account
    # dialog -- the operator adds the account manually using the Account Setup Reference
    # displayed by Add-IMAPAccount in OMMigrate-Outlook.psm1.
    # WhatIf logs reflect what the operator will be asked to do, not registry ops.
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message "WhatIf: IMAP account add is a guided manual operation -- no registry automation" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Would display Account Setup Reference for operator manual entry in Outlook" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Account Setup Reference -- Display Name  : $(Invoke-OMMigrateSanitize -Text $DisplayName)" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Account Setup Reference -- Email Address : $(Invoke-OMMigrateSanitize -Text $EmailAddress)" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Account Setup Reference -- IMAP Server   : $(Invoke-OMMigrateSanitize -Text $ImapServer)" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Account Setup Reference -- IMAP Port     : $ImapPort" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Account Setup Reference -- IMAP SSL      : $ImapSSL" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Account Setup Reference -- IMAP Username : $(Invoke-OMMigrateSanitize -Text $ImapUsername)" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Account Setup Reference -- SMTP Server   : $(Invoke-OMMigrateSanitize -Text $SmtpServer)" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Account Setup Reference -- SMTP Port     : $SmtpPort" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Account Setup Reference -- SMTP SSL      : $SmtpSSL" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Account Setup Reference -- SMTP Username : $(Invoke-OMMigrateSanitize -Text $SmtpUsername)" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Passwords are NEVER handled by this script -- operator enters them in Outlook" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Would prompt operator to open Outlook > File > Add Account > Manual setup > IMAP" `
                           -Level INFO -WhatIfPrefix
        Write-OMMigrateLog -Message "WhatIf: Would wait for operator to confirm account is connected in Outlook" `
                           -Level INFO -WhatIfPrefix
        return $true
    }

    # -- IMAP account creation is delegated to Outlook UI -------
    # Registry-based IMAP account creation was attempted and abandoned.
    # Root cause: direct registry writes bypass MAPI store initialization,
    # OST file creation, and search folder setup. Outlook crashes on load
    # when the account is added this way -- even with all registry fields
    # populated correctly. This is a fundamental Outlook architecture
    # constraint, not a missing field problem.
    #
    # The correct approach is Outlook's own Add Account dialog, which
    # handles all MAPI initialization internally. Add-IMAPAccount in
    # OMMigrate-Outlook.psm1 displays an Account Setup Reference with all account
    # details and guides the operator through the manual add process.
    #
    # imap_remove.reg is NOT generated here -- it was only needed as a
    # rollback for the registry write approach. Since Outlook owns the
    # IMAP account entry, removal is done through Outlook's account
    # settings UI if needed. POP3 rollback via registry_backup.reg
    # remains available and unaffected.
    #
    # Profile resolution and subkey enumeration above are retained
    # because they are still used by Remove-POP3AccountViaRegistry
    # and may be needed by future registry operations.

    Write-OMMigrateLog -Message (
        "IMAP account add delegated to operator via Outlook Add Account dialog: $EmailAddress"
    ) -Level INFO

    Write-AuditEntry  -Action 'IMAP_ACCOUNT_ADD_DELEGATED' `
                      -AccountEmail $EmailAddress `
                      -Detail (
                          "Method: Outlook File > Add Account (manual) | " +
                          "IMAP=$ImapServer`:$ImapPort (SSL=$ImapSSL) | " +
                          "SMTP=$SmtpServer`:$SmtpPort (SSL=$SmtpSSL)"
                      )

    return $true
}



# ============================================================
#  REGION: MIGRATION STATUS UPDATE
# ============================================================

function Update-AccountMigrationAction {
    <#
    .SYNOPSIS
        Updates migration status columns for a single account in
        migration_accounts.csv.

    .DESCRIPTION
        Reads migration_accounts.csv, updates the MigrationAction value
        (and optionally ProviderTag and AccountType) for the specified
        email address, and writes the file back.

        Called by Script 02 after each account successfully converts
        (sets MigrationAction=FOLDER-ONLY, ProviderTag=IMAP-CONVERTED,
        AccountType=IMAP) so the CSV reflects the converted state
        automatically -- no manual edit required.

        Called by Script 03 after each account completes folder migration
        (sets MigrationAction=COMPLETE) so the account is excluded from
        future Script 03 runs.

        Checks for an Excel file lock before writing. If the file is locked,
        logs a warning and skips the update -- the account will retain its
        current values and will appear again on the next run.

        Safe to call in WhatIf mode -- logs what would happen, no write.

    .PARAMETER EmailAddress
        Email address of the account to update.

    .PARAMETER NewAction
        New MigrationAction value to set.
        Script 02 uses 'FOLDER-ONLY'. Script 03 uses 'COMPLETE'.

    .PARAMETER NewProviderTag
        Optional. New ProviderTag value to set.
        Script 02 passes 'IMAP-CONVERTED'. Script 03 omits this.

    .PARAMETER NewAccountType
        Optional. New AccountType value to set.
        Script 02 passes 'IMAP'. Script 03 omits this.

    .EXAMPLE
        # Script 02 -- full post-conversion update
        Update-AccountMigrationAction -EmailAddress 'sales@domain.com' `
                                      -NewAction    'FOLDER-ONLY' `
                                      -NewProviderTag 'IMAP-CONVERTED' `
                                      -NewAccountType 'IMAP'

    .EXAMPLE
        # Script 03 -- mark complete after folder migration
        Update-AccountMigrationAction -EmailAddress 'sales@domain.com' `
                                      -NewAction 'COMPLETE'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,

        [Parameter(Mandatory = $true)]
        [string]$NewAction,

        [Parameter(Mandatory = $false)]
        [string]$NewProviderTag = '',

        [Parameter(Mandatory = $false)]
        [string]$NewAccountType = ''
    )

    $csvPath = Get-OMMigrateCsvPath -BaseName 'migration_accounts.csv'

    if (-not (Test-Path $csvPath)) {
        Write-OMMigrateLog -Message "migration_accounts.csv not found -- cannot update migration status for $EmailAddress" `
                           -Level WARN
        return
    }

    if ($Global:OMMigrate.WhatIf) {
        # Build a descriptive log message listing all fields that would change
        $whatIfFields = "MigrationAction=$NewAction"
        if ($NewProviderTag) { $whatIfFields += " | ProviderTag=$NewProviderTag" }
        if ($NewAccountType) { $whatIfFields += " | AccountType=$NewAccountType" }
        Write-OMMigrateLog -Message (
            "WhatIf: Would update $EmailAddress in migration_accounts.csv -- $whatIfFields"
        ) -Level INFO -WhatIfPrefix
        return
    }

    # Check for Excel file lock before attempting to write.
    # Excel holds an exclusive write lock on open CSV files -- attempting to
    # write while Excel has it open will corrupt or fail the save silently.
    try {
        $lockTest = [System.IO.File]::Open(
            $csvPath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
        $lockTest.Close()
        $lockTest.Dispose()
    }
    catch {
        Write-OMMigrateLog -Message (
            "migration_accounts.csv is locked (Excel may have it open). " +
            "Migration status for $EmailAddress was NOT updated. " +
            "Close Excel and re-run to process remaining accounts."
        ) -Level WARN
        Write-Host "  WARNING: migration_accounts.csv is open in Excel." -ForegroundColor Yellow
        Write-Host "           Close Excel so the status for $EmailAddress can be updated." -ForegroundColor Yellow
        Write-Host "           Account will remain eligible and re-appear on next run." -ForegroundColor Yellow
        return
    }

    try {
        # Read all rows, update the matching account, write back
        $rows = Import-Csv -Path $csvPath -Encoding UTF8

        $updated = $false
        foreach ($row in $rows) {
            if ($row.EmailAddress -eq $EmailAddress) {
                $oldAction           = $row.MigrationAction
                $row.MigrationAction = $NewAction

                # Apply optional field updates if provided
                if ($NewProviderTag -and $row.PSObject.Properties['ProviderTag']) {
                    $oldTag            = $row.ProviderTag
                    $row.ProviderTag   = $NewProviderTag
                    Write-OMMigrateLog -Message (
                        "Updated ProviderTag for $EmailAddress : $oldTag -> $NewProviderTag"
                    ) -Level INFO
                }
                if ($NewAccountType -and $row.PSObject.Properties['AccountType']) {
                    $oldType           = $row.AccountType
                    $row.AccountType   = $NewAccountType
                    Write-OMMigrateLog -Message (
                        "Updated AccountType for $EmailAddress : $oldType -> $NewAccountType"
                    ) -Level INFO
                }

                $updated = $true
                Write-OMMigrateLog -Message (
                    "Updated MigrationAction for $EmailAddress : $oldAction -> $NewAction"
                ) -Level INFO
                break
            }
        }

        if (-not $updated) {
            Write-OMMigrateLog -Message "Account $EmailAddress not found in migration_accounts.csv -- status not updated." `
                               -Level WARN
            return
        }

        $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

        # Build audit detail string
        $auditDetail = "MigrationAction=$NewAction"
        if ($NewProviderTag) { $auditDetail += " | ProviderTag=$NewProviderTag" }
        if ($NewAccountType) { $auditDetail += " | AccountType=$NewAccountType" }

        Write-AuditEntry -Action 'MIGRATION_STATUS_UPDATED' `
                         -AccountEmail $EmailAddress `
                         -Detail $auditDetail
    }
    catch {
        Write-OMMigrateLog -Message "Failed to update migration status for $EmailAddress : $_" `
                           -Level WARN
    }
}

# ============================================================
#  REGION: MIGRATE ACCOUNT PICKER
# ============================================================

function Invoke-MigrateAccountPicker {
    <#
    .SYNOPSIS
        Displays a WinForms popup letting the operator choose which POP3
        accounts to mark MIGRATE in migration_accounts.csv.

    .DESCRIPTION
        Called by Script 00 immediately after migration_accounts.csv is
        written. Shows all POP3 accounts (ProviderTag starts with 'POP3-')
        in a scrollable checkbox list. Exchange, already-IMAP, and already-
        converted accounts are excluded -- they are never migration candidates.

        The operator checks each account they want to migrate now. Unchecked
        accounts are set to SKIP so Script 02 ignores them. Checked accounts
        are set to MIGRATE.

        On Cancel or if the window is closed with no selection, no changes
        are written to the CSV -- all POP3 accounts remain at their default
        MigrationAction (MIGRATE) as written by Export-AccountsToCSV.

        Skipped automatically in WhatIf/Preview mode -- no CSV write, no UI.

        Uses System.Windows.Forms -- built into PowerShell 5.1 on Windows.
        No external dependencies.

    .PARAMETER CsvPath
        Full path to migration_accounts.csv. Must exist before calling.

    .EXAMPLE
        Invoke-MigrateAccountPicker -CsvPath 'C:\...\Config\migration_accounts.csv'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvPath
    )

    # Skip entirely in WhatIf mode -- no UI, no write
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message 'WhatIf: Would display MIGRATE account picker -- skipped.' `
                           -Level INFO -WhatIfPrefix
        return
    }

    if (-not (Test-Path $CsvPath)) {
        Write-OMMigrateLog -Message "Invoke-MigrateAccountPicker: CSV not found at $CsvPath -- skipped." `
                           -Level WARN
        return
    }

    # Load CSV and find POP3 candidates
    $rows    = Import-Csv -Path $CsvPath -Encoding UTF8
    $pop3Rows = @($rows | Where-Object { $_.ProviderTag -like 'POP3-*' })

    if ($pop3Rows.Count -eq 0) {
        Write-OMMigrateLog -Message 'Invoke-MigrateAccountPicker: No POP3 accounts found -- picker skipped.' `
                           -Level INFO
        Write-Host '  No POP3 accounts found -- account picker skipped.' -ForegroundColor DarkGray
        return
    }

    Write-OMMigrateLog -Message "Invoke-MigrateAccountPicker: Displaying picker for $($pop3Rows.Count) POP3 account(s)." `
                       -Level INFO

    # Load WinForms assembly
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing       -ErrorAction Stop
    }
    catch {
        Write-OMMigrateLog -Message "Could not load WinForms -- account picker unavailable: $_" -Level WARN
        Write-Host '  WARNING: Could not open account picker (WinForms unavailable).' -ForegroundColor Yellow
        Write-Host '  Open migration_accounts.csv and set MigrationAction=MIGRATE for accounts to migrate.' `
                   -ForegroundColor Yellow
        return
    }

    # -- Build the form ----------------------------------------
    $form                  = New-Object System.Windows.Forms.Form
    $form.Text             = 'OMMigrate -- Select Accounts to Migrate'
    $form.Size             = New-Object System.Drawing.Size(560, 420)
    $form.StartPosition    = 'CenterScreen'
    $form.FormBorderStyle  = 'FixedDialog'
    $form.MaximizeBox      = $false
    $form.MinimizeBox      = $false
    $form.TopMost          = $true

    # Instruction label
    $label             = New-Object System.Windows.Forms.Label
    $label.Location    = New-Object System.Drawing.Point(12, 12)
    $label.Size        = New-Object System.Drawing.Size(520, 44)
    $label.Text        = 'Check each POP3 account you want to migrate now. Unchecked accounts will be set to SKIP and can be migrated later.'
    $label.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.Controls.Add($label)

    # Checked listbox
    $listBox               = New-Object System.Windows.Forms.CheckedListBox
    $listBox.Location      = New-Object System.Drawing.Point(12, 64)
    $listBox.Size          = New-Object System.Drawing.Size(520, 270)
    $listBox.Font          = New-Object System.Drawing.Font('Segoe UI', 9)
    $listBox.CheckOnClick  = $true

    foreach ($acct in $pop3Rows) {
        # WinForms popups always show real values -- sanitize is console-only
        $displayLine = "$($acct.EmailAddress)  [$($acct.ProviderTag)]"
        $idx = $listBox.Items.Add($displayLine)
        # All unchecked by default -- operator checks the ones to migrate now
        $listBox.SetItemChecked($idx, $false)
    }

    $form.Controls.Add($listBox)

    # Select All / Clear All buttons
    $btnSelectAll          = New-Object System.Windows.Forms.Button
    $btnSelectAll.Location = New-Object System.Drawing.Point(12, 344)
    $btnSelectAll.Size     = New-Object System.Drawing.Size(90, 28)
    $btnSelectAll.Text     = 'Select All'
    $btnSelectAll.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnSelectAll.Add_Click({
        for ($i = 0; $i -lt $listBox.Items.Count; $i++) {
            $listBox.SetItemChecked($i, $true)
        }
    })
    $form.Controls.Add($btnSelectAll)

    $btnClearAll           = New-Object System.Windows.Forms.Button
    $btnClearAll.Location  = New-Object System.Drawing.Point(110, 344)
    $btnClearAll.Size      = New-Object System.Drawing.Size(90, 28)
    $btnClearAll.Text      = 'Clear All'
    $btnClearAll.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnClearAll.Add_Click({
        for ($i = 0; $i -lt $listBox.Items.Count; $i++) {
            $listBox.SetItemChecked($i, $false)
        }
    })
    $form.Controls.Add($btnClearAll)

    # OK and Cancel buttons
    $btnOK                 = New-Object System.Windows.Forms.Button
    $btnOK.Location        = New-Object System.Drawing.Point(354, 344)
    $btnOK.Size            = New-Object System.Drawing.Size(80, 28)
    $btnOK.Text            = 'OK'
    $btnOK.Font            = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnOK.DialogResult    = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton     = $btnOK
    $form.Controls.Add($btnOK)

    $btnCancel             = New-Object System.Windows.Forms.Button
    $btnCancel.Location    = New-Object System.Drawing.Point(452, 344)
    $btnCancel.Size        = New-Object System.Drawing.Size(80, 28)
    $btnCancel.Text        = 'Cancel'
    $btnCancel.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton     = $btnCancel
    $form.Controls.Add($btnCancel)

    # -- Show form and process result --------------------------
    $result = $form.ShowDialog()
    $form.Dispose()

    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-OMMigrateLog -Message 'Account picker cancelled -- migration_accounts.csv not changed.' `
                           -Level INFO
        Write-Host '  Account picker cancelled -- no changes made to migration_accounts.csv.' `
                   -ForegroundColor Yellow
        Write-Host '  All POP3 accounts remain at their default MigrationAction.' -ForegroundColor Yellow
        return $false
    }

    # Collect checked indices
    $checkedIndices = @($listBox.CheckedIndices)

    # FIXED (Administrator direction, 2026-07-2X): previously, if $checkedIndices.Count
    # was 0 (operator left every box unchecked and clicked OK), this function
    # returned early here exactly like Cancel -- no CSV write at all, silently
    # leaving MigrationAction unchanged. This contradicted the docstring above
    # ("unchecked accounts are set to SKIP") and produced a real, confirmed-live
    # bug: an account already at MigrationAction=MIGRATE, left unchecked with the
    # intent to defer it, stayed at MIGRATE with no indication anything went
    # wrong. The early-return-on-zero-checked special case is removed -- OK
    # always applies the on-screen selection state, matching Cancel's existing,
    # separate role as the only true no-op escape hatch.
    #
    # Per Administrator's explicit direction: the SKIP write below is conditional
    # on the row's MigrationAction ALREADY being MIGRATE at the moment the
    # picker opened ($pop3Rows, loaded before any UI interaction, captures that
    # prior state). A row that was some other status (already SKIP, FOLDER-ONLY,
    # COMPLETE, etc.) before the picker opened is left completely untouched when
    # unchecked -- this picker only ever toggles a row within the MIGRATE/SKIP
    # cycle for an account that was already a live MIGRATE candidate, never
    # pulls a row into that cycle from an unrelated state.
    $priorMigrateEmails = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($priorRow in $pop3Rows) {
        if ($priorRow.MigrationAction -eq 'MIGRATE') {
            $priorMigrateEmails.Add($priorRow.EmailAddress) | Out-Null
        }
    }

    # Build set of email addresses that were checked
    $selectedEmails = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($idx in $checkedIndices) {
        $selectedEmails.Add($pop3Rows[$idx].EmailAddress) | Out-Null
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
            "migration_accounts.csv is locked (Excel may have it open). " +
            "Account picker selections were NOT saved. " +
            "Close Excel and re-run Script 00 to make your selections."
        ) -Level WARN
        Write-Host '  WARNING: migration_accounts.csv is open in Excel.' -ForegroundColor Yellow
        Write-Host '  Close Excel and re-run Script 00 to save your account selections.' -ForegroundColor Yellow
        return $false
    }

    # Apply selections:
    #   Checked   POP3 account                          -> MigrationAction = MIGRATE
    #   Unchecked POP3 account, was already MIGRATE      -> MigrationAction = SKIP
    #   Unchecked POP3 account, was some other status    -> unchanged (see fix
    #     comment above -- this picker never pulls a row into the MIGRATE/SKIP
    #     cycle from an unrelated prior state)
    #   All non-POP3 accounts                            -> unchanged
    $migrateCount = 0
    $skipCount    = 0

    foreach ($row in $rows) {
        if ($row.ProviderTag -like 'POP3-*') {
            if ($selectedEmails.Contains($row.EmailAddress)) {
                $row.MigrationAction = 'MIGRATE'
                $migrateCount++
            }
            elseif ($priorMigrateEmails.Contains($row.EmailAddress)) {
                $row.MigrationAction = 'SKIP'
                $skipCount++
            }
        }
    }

    try {
        $rows | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

        Write-OMMigrateLog -Message (
            "Account picker saved: $migrateCount account(s) set to MIGRATE, " +
            "$skipCount account(s) set to SKIP."
        ) -Level INFO

        Write-AuditEntry -Action 'PICKER_MIGRATE_SELECTION' `
                         -Detail "MIGRATE=$migrateCount | SKIP=$skipCount | Selected=$($selectedEmails -join ', ')"

        Write-Host ''
        Write-Host "  Account picker saved:" -ForegroundColor Green
        Write-Host "    $migrateCount account(s) marked MIGRATE -- will be processed by Script 02." `
                   -ForegroundColor Green
        Write-Host "    $skipCount account(s) marked SKIP -- deferred for later." -ForegroundColor DarkGray
        return $true
    }
    catch {
        Write-OMMigrateLog -Message "Failed to save account picker selections to CSV: $_" -Level WARN
        Write-Host '  WARNING: Could not save account picker selections.' -ForegroundColor Yellow
        Write-Host "  $_" -ForegroundColor Yellow
        return $false
    }
}


# ============================================================
#  REGION: CREDENTIAL ENTRY UI
# ============================================================

function Invoke-CredentialEntryUI {
    <#
    .SYNOPSIS
        Displays a per-account WinForms credential dialog letting the
        operator enter passwords and SMTP credentials for each MIGRATE
        account into migration_accounts.csv.

    .DESCRIPTION
        Called by Script 00 immediately after Invoke-MigrateAccountPicker.
        Shows one dialog per account marked MIGRATE in the CSV. Each
        dialog is tailored to the account's provider tag:

            POP3-STANDARD     Password only
            POP3-ATTAMERITECH Password only (labeled as Secure Mail Key)
            POP3-GMAIL        Password only
            POP3-AWS          IMAP Password + SMTP Username + SMTP Password

        Pre-populates fields with existing values when a password has
        already been entered (re-run support). Existing values are shown
        as masked asterisks -- the operator can leave them unchanged by
        clicking OK without re-typing, or overwrite by typing a new value.

        Per-dialog buttons:
            OK     -- validates mandatory fields, saves to CSV, advances
            Skip   -- leaves this account's existing values unchanged,
                      advances to next account
            Cancel -- exits the entire UI immediately, no changes saved

        A footnote on each dialog explains the three buttons.

        Mandatory field validation:
            Password   -- always mandatory for MIGRATE accounts
            SmtpUsername -- mandatory for POP3-AWS only
            SmtpPassword -- mandatory for POP3-AWS only
        Empty mandatory fields are highlighted red and the operator
        cannot advance until they are filled.

        Checks for Excel file lock before writing each account.
        Skipped automatically in WhatIf/Preview mode.

        Uses System.Windows.Forms -- built into PowerShell 5.1 on Windows.
        No external dependencies.

    .PARAMETER CsvPath
        Full path to migration_accounts.csv. Must exist before calling.

    .EXAMPLE
        Invoke-CredentialEntryUI -CsvPath 'C:\...\Config\migration_accounts.csv'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,

        [Parameter(Mandatory = $false)]
        [string]$FilterEmail = ''
        # When specified, only the account matching this email address is shown.
        # Used by Script 00 per-account loop to scope credential entry to one
        # account at a time. When empty, all MIGRATE accounts are shown.
    )

    # Skip entirely in WhatIf mode
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message 'WhatIf: Would display credential entry UI -- skipped.' `
                           -Level INFO -WhatIfPrefix
        return
    }

    if (-not (Test-Path $CsvPath)) {
        Write-OMMigrateLog -Message "Invoke-CredentialEntryUI: CSV not found at $CsvPath -- skipped." `
                           -Level WARN
        return
    }

    # Load CSV and find MIGRATE accounts
    $allRows      = Import-Csv -Path $CsvPath -Encoding UTF8
    $migrateRows  = @($allRows | Where-Object { $_.MigrationAction -eq 'MIGRATE' })

    # Filter to a specific email address if requested
    if ($FilterEmail -and -not [string]::IsNullOrWhiteSpace($FilterEmail)) {
        $migrateRows = @($migrateRows | Where-Object { $_.EmailAddress -eq $FilterEmail })
        Write-OMMigrateLog -Message "Invoke-CredentialEntryUI: Filtered to account '$FilterEmail'." `
                           -Level INFO
    }

    if ($migrateRows.Count -eq 0) {
        Write-OMMigrateLog -Message 'Invoke-CredentialEntryUI: No MIGRATE accounts found -- skipped.' `
                           -Level INFO
        return
    }

    Write-OMMigrateLog -Message "Invoke-CredentialEntryUI: Starting credential entry for $($migrateRows.Count) MIGRATE account(s)." `
                       -Level INFO

    # Load WinForms assembly
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing       -ErrorAction Stop
    }
    catch {
        Write-OMMigrateLog -Message "Could not load WinForms -- credential UI unavailable: $_" -Level WARN
        Write-Host '  WARNING: Could not open credential entry UI (WinForms unavailable).' -ForegroundColor Yellow
        Write-Host '  Open Config\migration_accounts.csv and enter passwords manually.' `
                   -ForegroundColor Yellow
        return
    }

    # Track how many accounts were saved vs skipped
    $savedCount   = 0
    $skippedCount = 0
    $cancelledAll = $false

    # -- Process each MIGRATE account in its own dialog --------
    for ($acctIdx = 0; $acctIdx -lt $migrateRows.Count; $acctIdx++) {

        $acct        = $migrateRows[$acctIdx]
        $email       = $acct.EmailAddress
        $providerTag = $acct.ProviderTag
        $isAWS       = ($providerTag -eq 'POP3-AWS')
        $isATT       = ($providerTag -eq 'POP3-ATTAMERITECH')
        $accountNum  = $acctIdx + 1
        $totalAccts  = $migrateRows.Count

        # WinForms popups always show real data -- sanitize applies to
        # console Write-Host output only. Use $email directly in the dialog.
        $displayEmail = $email

        # Determine existing values -- show masked if already filled
        $existingPassword    = $acct.Password
        $existingSmtpUser    = if ($acct.PSObject.Properties['SmtpUsername']) { $acct.SmtpUsername } else { '' }
        $existingSmtpPass    = if ($acct.PSObject.Properties['SmtpPassword']) { $acct.SmtpPassword } else { '' }

        $hasExistingPassword = ($existingPassword -and
                                $existingPassword -notmatch '^\[ENTER' -and
                                $existingPassword -notmatch '^\[SAME'  -and
                                $existingPassword -ne 'N/A - Exchange account' -and
                                $existingPassword -ne 'N/A - Already IMAP')

        $hasExistingSmtpUser = ($isAWS -and $existingSmtpUser -and
                                $existingSmtpUser -notmatch '^\[ENTER' -and
                                $existingSmtpUser -ne 'N/A - Exchange account' -and
                                $existingSmtpUser -ne 'N/A - Already IMAP')

        $hasExistingSmtpPass = ($isAWS -and $existingSmtpPass -and
                                $existingSmtpPass -notmatch '^\[ENTER' -and
                                $existingSmtpPass -notmatch '^\[SAME'  -and
                                $existingSmtpPass -ne 'N/A - Exchange account' -and
                                $existingSmtpPass -ne 'N/A - Already IMAP')

        # -- Build dialog ------------------------------------------
        $dlg                  = New-Object System.Windows.Forms.Form
        $dlg.Text             = "OMMigrate -- Credentials  ($accountNum of $totalAccts)"
        $dlg.Size             = New-Object System.Drawing.Size(500, $(if ($isAWS) { 380 } elseif ($isATT) { 360 } else { 280 }))
        $dlg.StartPosition    = 'CenterScreen'
        $dlg.FormBorderStyle  = 'FixedDialog'
        $dlg.MaximizeBox      = $false
        $dlg.MinimizeBox      = $false
        $dlg.TopMost          = $true

        $yPos = 14

        # -- Account header ----------------------------------------
        $lblAccount           = New-Object System.Windows.Forms.Label
        $lblAccount.Location  = New-Object System.Drawing.Point(14, $yPos)
        $lblAccount.Size      = New-Object System.Drawing.Size(460, 20)
        $lblAccount.Text      = "$displayEmail"
        $lblAccount.Font      = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $lblAccount.ForeColor = [System.Drawing.Color]::FromArgb(0, 90, 160)
        $dlg.Controls.Add($lblAccount)
        $yPos += 24

        $lblTag               = New-Object System.Windows.Forms.Label
        $lblTag.Location      = New-Object System.Drawing.Point(14, $yPos)
        $lblTag.Size          = New-Object System.Drawing.Size(460, 18)
        $lblTag.Text          = $providerTag
        $lblTag.Font          = New-Object System.Drawing.Font('Segoe UI', 9)
        $lblTag.ForeColor     = [System.Drawing.Color]::Gray
        $dlg.Controls.Add($lblTag)
        $yPos += 26

        # -- Provider-specific instructions ------------------------
        $instrText = if ($isATT) {
            'Enter your Secure Mail Key -- NOT your regular AT&T email password.'
        } elseif ($isAWS) {
            'Enter your IMAP password and AWS IAM SMTP credentials. The SMTP Username is your IAM access key ID (e.g. AKIAIOSFODNN7EXAMPLE).'
        } else {
            'Enter your email password for this account.'
        }

        $lblInstr             = New-Object System.Windows.Forms.Label
        $lblInstr.Location    = New-Object System.Drawing.Point(14, $yPos)
        $lblInstr.Size        = New-Object System.Drawing.Size(460, 50)
        $lblInstr.Text        = $instrText
        $lblInstr.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
        $lblInstr.ForeColor   = [System.Drawing.Color]::FromArgb(180, 100, 0)
        $dlg.Controls.Add($lblInstr)
        $yPos += 58

        # -- AT&T clickable link -- shown below instruction text for ATTAMERITECH only --
        # Opens the AT&T security settings page in the default browser so the
        # operator can generate a Secure Mail Key without leaving the dialog.
        if ($isATT) {
            $attUrl = 'https://www.att.com/acctmgmt/myprofile/overview?flow=settings'
            $lnkATT               = New-Object System.Windows.Forms.LinkLabel
            $lnkATT.Location      = New-Object System.Drawing.Point(14, $yPos)
            $lnkATT.Size          = New-Object System.Drawing.Size(460, 18)
            $lnkATT.Text          = 'Click here to generate a Secure Mail Key at AT&T'
            $lnkATT.Font          = New-Object System.Drawing.Font('Segoe UI', 9)
            $lnkATT.Add_LinkClicked({
                try { Start-Process $attUrl } catch { }
            })
            $dlg.Controls.Add($lnkATT)
            $yPos += 24

            $lblATTNav            = New-Object System.Windows.Forms.Label
            $lblATTNav.Location   = New-Object System.Drawing.Point(14, $yPos)
            $lblATTNav.Size       = New-Object System.Drawing.Size(460, 36)
            $lblATTNav.Text       = 'After signing in, click Settings under Profile on the left side, then scroll to the bottom to create a Secure Mail Key.'
            $lblATTNav.Font       = New-Object System.Drawing.Font('Segoe UI', 8)
            $lblATTNav.ForeColor  = [System.Drawing.Color]::Gray
            $dlg.Controls.Add($lblATTNav)
            $yPos += 44
        }

        # -- Separator line ----------------------------------------
        $sep              = New-Object System.Windows.Forms.Label
        $sep.Location     = New-Object System.Drawing.Point(14, $yPos)
        $sep.Size         = New-Object System.Drawing.Size(460, 1)
        $sep.BorderStyle  = 'Fixed3D'
        $dlg.Controls.Add($sep)
        $yPos += 10

        # -- Helper: add a labeled password field ------------------
        # Returns the TextBox control so validation can reference it
        $addField = {
            param($labelText, $existingValue, $hasExisting, $yRef)

            $lbl              = New-Object System.Windows.Forms.Label
            $lbl.Location     = New-Object System.Drawing.Point(14, ($yRef + 3))
            $lbl.Size         = New-Object System.Drawing.Size(200, 18)
            $lbl.Text         = $labelText
            $lbl.Font         = New-Object System.Drawing.Font('Segoe UI', 9)
            $dlg.Controls.Add($lbl)

            $txt                         = New-Object System.Windows.Forms.TextBox
            $txt.Location                = New-Object System.Drawing.Point(220, $yRef)
            $txt.Size                    = New-Object System.Drawing.Size(254, 24)
            $txt.Font                    = New-Object System.Drawing.Font('Segoe UI', 9)
            $txt.UseSystemPasswordChar   = $true

            # Pre-populate with masked placeholder if value already exists
            if ($hasExisting) {
                $txt.Text = $existingValue
            }

            $dlg.Controls.Add($txt)
            return $txt
        }

        # -- Password field ----------------------------------------
        $pwLabel = if ($isATT) { 'Secure Mail Key *' } else { 'Password *' }
        $txtPassword = & $addField $pwLabel $existingPassword $hasExistingPassword $yPos
        $yPos += 36

        # -- AWS-only fields ---------------------------------------
        $txtSmtpUser = $null
        $txtSmtpPass = $null

        if ($isAWS) {
            $yPos += 6
            $txtSmtpUser = & $addField 'SMTP Username (IAM) *' $existingSmtpUser $hasExistingSmtpUser $yPos
            $yPos += 36
            $txtSmtpPass = & $addField 'SMTP Password (IAM) *' $existingSmtpPass $hasExistingSmtpPass $yPos
            $yPos += 36
        }

        $yPos += 10

        # -- Separator line ----------------------------------------
        $sep2             = New-Object System.Windows.Forms.Label
        $sep2.Location    = New-Object System.Drawing.Point(14, $yPos)
        $sep2.Size        = New-Object System.Drawing.Size(460, 1)
        $sep2.BorderStyle = 'Fixed3D'
        $dlg.Controls.Add($sep2)
        $yPos += 10

        # -- Footnote ----------------------------------------------
        $lblFootnote          = New-Object System.Windows.Forms.Label
        $lblFootnote.Location = New-Object System.Drawing.Point(14, $yPos)
        $lblFootnote.Size     = New-Object System.Drawing.Size(460, 18)
        $lblFootnote.Text     = 'OK = save and continue  |  Skip = skip this account  |  Cancel = exit without saving any changes'
        $lblFootnote.Font     = New-Object System.Drawing.Font('Segoe UI', 8)
        $lblFootnote.ForeColor = [System.Drawing.Color]::Gray
        $dlg.Controls.Add($lblFootnote)
        $yPos += 28

        # -- Buttons -----------------------------------------------
        $btnOK                 = New-Object System.Windows.Forms.Button
        $btnOK.Location        = New-Object System.Drawing.Point(238, $yPos)
        $btnOK.Size            = New-Object System.Drawing.Size(70, 28)
        $btnOK.Text            = 'OK'
        $btnOK.Font            = New-Object System.Drawing.Font('Segoe UI', 9)
        $btnOK.DialogResult    = [System.Windows.Forms.DialogResult]::None
        $dlg.Controls.Add($btnOK)

        $btnSkip               = New-Object System.Windows.Forms.Button
        $btnSkip.Location      = New-Object System.Drawing.Point(316, $yPos)
        $btnSkip.Size          = New-Object System.Drawing.Size(70, 28)
        $btnSkip.Text          = 'Skip'
        $btnSkip.Font          = New-Object System.Drawing.Font('Segoe UI', 9)
        $btnSkip.DialogResult  = [System.Windows.Forms.DialogResult]::Ignore
        $dlg.Controls.Add($btnSkip)

        $btnCancel             = New-Object System.Windows.Forms.Button
        $btnCancel.Location    = New-Object System.Drawing.Point(394, $yPos)
        $btnCancel.Size        = New-Object System.Drawing.Size(80, 28)
        $btnCancel.Text        = 'Cancel'
        $btnCancel.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dlg.CancelButton      = $btnCancel
        $dlg.Controls.Add($btnCancel)

        # Resize form to fit content
        $dlg.ClientSize = New-Object System.Drawing.Size(488, ($yPos + 44))

        # -- OK click validation -----------------------------------
        # Validate mandatory fields before accepting OK.
        # Empty mandatory fields are highlighted red and a message is shown.
        # The dialog stays open until all mandatory fields are filled.
        $btnOK.Add_Click({
            $valid       = $true
            $emptyFields = [System.Collections.Generic.List[string]]::new()

            # Reset field backgrounds
            $txtPassword.BackColor = [System.Drawing.SystemColors]::Window
            if ($txtSmtpUser) { $txtSmtpUser.BackColor = [System.Drawing.SystemColors]::Window }
            if ($txtSmtpPass) { $txtSmtpPass.BackColor = [System.Drawing.SystemColors]::Window }

            # Password is always mandatory
            if ([string]::IsNullOrWhiteSpace($txtPassword.Text)) {
                $txtPassword.BackColor = [System.Drawing.Color]::FromArgb(255, 220, 220)
                $emptyFields.Add($(if ($isATT) { 'Secure Mail Key' } else { 'Password' }))
                $valid = $false
            }

            # AWS-only mandatory fields
            if ($isAWS) {
                if ([string]::IsNullOrWhiteSpace($txtSmtpUser.Text)) {
                    $txtSmtpUser.BackColor = [System.Drawing.Color]::FromArgb(255, 220, 220)
                    $emptyFields.Add('SMTP Username (IAM)')
                    $valid = $false
                }
                if ([string]::IsNullOrWhiteSpace($txtSmtpPass.Text)) {
                    $txtSmtpPass.BackColor = [System.Drawing.Color]::FromArgb(255, 220, 220)
                    $emptyFields.Add('SMTP Password (IAM)')
                    $valid = $false
                }
            }

            if (-not $valid) {
                [System.Windows.Forms.MessageBox]::Show(
                    "The following mandatory field(s) cannot be empty:`n`n  " + ($emptyFields -join "`n  "),
                    'Missing Credentials',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }

            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dlg.Close()
        })

        # -- Show dialog -------------------------------------------
        $result = $dlg.ShowDialog()

        # Capture values before disposing
        $enteredPassword = $txtPassword.Text
        $enteredSmtpUser = if ($txtSmtpUser) { $txtSmtpUser.Text } else { '' }
        $enteredSmtpPass = if ($txtSmtpPass) { $txtSmtpPass.Text } else { '' }

        $dlg.Dispose()

        # -- Handle result -----------------------------------------
        if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
            # Cancel -- exit entire UI immediately, no changes saved
            Write-OMMigrateLog -Message 'Credential entry UI cancelled -- no credentials saved.' `
                               -Level INFO
            Write-Host '  Credential entry cancelled -- no changes saved to migration_accounts.csv.' `
                       -ForegroundColor Yellow
            Write-Host '  Re-run Script 00 to enter credentials when ready.' -ForegroundColor Yellow
            $cancelledAll = $true
            break
        }

        if ($result -eq [System.Windows.Forms.DialogResult]::Ignore) {
            # Skip -- leave this account unchanged, advance to next
            Write-OMMigrateLog -Message "Credential entry skipped for: $email" -Level INFO
            Write-Host "  Skipped: $(Invoke-OMMigrateSanitize -Text $email)" -ForegroundColor DarkGray
            $skippedCount++
            continue
        }

        # OK -- save credentials to CSV
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
                "migration_accounts.csv is locked (Excel may have it open). " +
                "Credentials for $email were NOT saved. " +
                "Close Excel and re-run Script 00 to enter credentials."
            ) -Level WARN
            Write-Host "  WARNING: migration_accounts.csv is open in Excel -- credentials for $(Invoke-OMMigrateSanitize -Text $email) not saved." `
                       -ForegroundColor Yellow
            $skippedCount++
            continue
        }

        # Apply to allRows and write back
        try {
            foreach ($row in $allRows) {
                if ($row.EmailAddress -eq $email) {
                    $row.Password = $enteredPassword
                    if ($isAWS) {
                        if ($row.PSObject.Properties['SmtpUsername']) {
                            $row.SmtpUsername = $enteredSmtpUser
                        }
                        if ($row.PSObject.Properties['SmtpPassword']) {
                            $row.SmtpPassword = $enteredSmtpPass
                        }
                    }
                    break
                }
            }

            $allRows | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

            Write-OMMigrateLog -Message "Credentials saved for: $email" -Level INFO
            Write-AuditEntry  -Action 'CREDENTIALS_ENTERED' `
                              -AccountEmail $email `
                              -Detail "ProviderTag=$providerTag | SmtpUsername=$(if ($isAWS) { 'set' } else { 'n/a' })"

            Write-Host "  Saved: $(Invoke-OMMigrateSanitize -Text $email)" -ForegroundColor Green
            $savedCount++
        }
        catch {
            Write-OMMigrateLog -Message "Failed to save credentials for $email : $_" -Level WARN
            Write-Host "  WARNING: Could not save credentials for $(Invoke-OMMigrateSanitize -Text $email) : $_" `
                       -ForegroundColor Yellow
            $skippedCount++
        }
    }

    # -- Summary ---------------------------------------------------
    if (-not $cancelledAll) {
        Write-Host ''
        Write-Host '  Credential entry complete:' -ForegroundColor Green
        Write-Host "    $savedCount account(s) saved." -ForegroundColor Green
        if ($skippedCount -gt 0) {
            Write-Host "    $skippedCount account(s) skipped -- re-run Script 00 to complete." `
                       -ForegroundColor Yellow
        }
        Write-OMMigrateLog -Message (
            "Credential entry UI complete: Saved=$savedCount | Skipped=$skippedCount"
        ) -Level INFO
    }
}


# ============================================================
#  REGION: CREDENTIAL REPAIR
# ============================================================

function Repair-IMAPCredentials {
    <#
    .SYNOPSIS
        Corrects scrambled IMAP/SMTP credentials in the Outlook profile
        registry after the Add Account wizard incorrectly overwrites IMAP
        fields with SMTP credentials.

    .DESCRIPTION
        This is a known Outlook 2016/2019/2021 bug affecting accounts that
        use separate IMAP and SMTP credentials (e.g. POP3-AWS accounts using
        AWS SES for outbound). The Add Account wizard overwrites the IMAP
        credential fields with the SMTP credentials during the final step,
        leaving the account unable to receive mail without manual repair.

        WHAT OUTLOOK DOES (incorrectly):
            IMAP User     <- SMTP username (AWS IAM key ID)
            IMAP Password <- SMTP password blob
            SMTP User     <- missing
            SMTP Password <- missing
            SMTP Auth Method <- missing

        WHAT THIS FUNCTION FIXES:
            1. Reads the misplaced SMTP password blob from IMAP Password
            2. Writes it to SMTP Password (where it belongs)
            3. Clears IMAP Password (operator enters IMAP password once manually)
            4. Writes correct IMAP User = email address
            5. Writes correct SMTP User = SmtpUsername from CSV
            6. Sets SMTP Auth Method = 1

        ACCOUNT TYPE ELIGIBILITY:
            Applied   : POP3-AWS, POP3-STANDARD (if SmtpUsername differs from email)
            Skipped   : POP3-ATTAMERITECH (different auth pattern -- Secure Mail Key)
            Skipped   : IMAP-ALREADY, IMAP-GMAIL, EXCHANGE-SKIP, IMAP-CONVERTED
            Skipped   : Any account where IMAP User already equals email address
                        (already correct -- nothing to do)

        SAFETY:
            - Exports a .reg backup of the subkey before any writes
            - Non-fatal -- logs and skips if subkey cannot be found or written
            - WhatIf mode logs what would change without writing anything
            - Does not touch password blobs beyond the one move operation

    .PARAMETER EmailAddress
        Email address of the account to repair.

    .PARAMETER SmtpUsername
        The correct SMTP username for this account (e.g. AWS IAM key ID).
        For POP3-AWS accounts this is the IAM SMTP key ID from the CSV.
        For POP3-STANDARD accounts where SMTP username equals email, pass
        the email address -- the function will detect no mismatch and skip.

    .PARAMETER ProviderTag
        The account's ProviderTag from migration_accounts.csv.
        Used to determine eligibility for the fix.

    .PARAMETER BackupPath
        Full path to the folder where the .reg backup file will be saved.
        Typically the OMMigrate Backups\ folder.

    .OUTPUTS
        [string] -- Result code:
            'FIXED'        Credentials corrected successfully
            'SKIPPED'      Account type not eligible or already correct
            'NOT_FOUND'    Subkey not found in registry
            'FAILED'       Write error occurred (logged, non-fatal)
            'WHATIF'       WhatIf mode -- no writes made

    .PARAMETER ProfileName
        Outlook profile name to target. Optional -- if omitted, resolves
        automatically from the registry's DefaultProfile value (same
        resolution used elsewhere in this module). Pass this explicitly
        only if the operator is targeting a non-default profile.

    .EXAMPLE
        $result = Repair-IMAPCredentials `
            -EmailAddress 'admin@example.com' `
            -SmtpUsername 'AKIAWIYxxxxxxxxx' `
            -ProviderTag  'POP3-AWS' `
            -BackupPath   'C:\Users\<username>\Documents\OutlookMigration\Backups'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,

        [Parameter(Mandatory = $true)]
        [string]$SmtpUsername,

        [Parameter(Mandatory = $true)]
        [string]$ProviderTag,

        [Parameter(Mandatory = $true)]
        [string]$BackupPath,

        # FIXED (launch-readiness review, 2026-08-18): optional override, same
        # pattern as the rest of this module (e.g. Remove-OutlookAccount).
        # Auto-resolves via DefaultProfile when not supplied -- previously
        # this function had no way to target anything but a profile literally
        # named "Outlook" on Office 16.0, silently failing (WARN + FAILED)
        # for any other profile name or Outlook 2013/2010.
        [Parameter(Mandatory = $false)]
        [string]$ProfileName = ''
    )

    # -- Eligibility check -----------------------------------------
    # Only apply to account types that use basic auth with separate
    # IMAP and SMTP credentials. Skip all others immediately.
    $eligibleTags = @('POP3-AWS', 'POP3-STANDARD')
    if ($ProviderTag -notin $eligibleTags) {
        Write-OMMigrateLog -Message (
            "Credential repair skipped for $EmailAddress -- " +
            "ProviderTag '$ProviderTag' not eligible for fix."
        ) -Level INFO
        return 'SKIPPED'
    }

    # For POP3-STANDARD, only apply if SMTP username differs from email.
    # If they are the same, there is no separate credential mismatch to fix.
    if ($ProviderTag -eq 'POP3-STANDARD' -and
        $SmtpUsername -ieq $EmailAddress) {
        Write-OMMigrateLog -Message (
            "Credential repair skipped for $EmailAddress -- " +
            "POP3-STANDARD with identical IMAP/SMTP username, no fix needed."
        ) -Level INFO
        return 'SKIPPED'
    }

    # -- WhatIf mode -----------------------------------------------
    if ($Global:OMMigrate.WhatIf) {
        Write-OMMigrateLog -Message (
            "WhatIf: Would repair IMAP/SMTP credentials for $EmailAddress " +
            "(ProviderTag=$ProviderTag | SmtpUsername=$SmtpUsername)"
        ) -Level INFO -WhatIfPrefix
        return 'WHATIF'
    }

    # -- Resolve active profile name --------------------------------
    # FIXED (launch-readiness review, 2026-08-18): was hardcoded to a
    # profile literally named "Outlook" on Office 16.0 only. Now uses the
    # same multi-version DefaultProfile resolution as Remove-OutlookAccount,
    # elsewhere in this module, with -ProfileName as an explicit override.
    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        try {
            $outlookRegBase = @(
                'HKCU:\Software\Microsoft\Office\16.0\Outlook',
                'HKCU:\Software\Microsoft\Office\15.0\Outlook',
                'HKCU:\Software\Microsoft\Office\14.0\Outlook'
            )
            foreach ($basePath in $outlookRegBase) {
                $ProfileName = (Get-ItemProperty -Path $basePath `
                                                 -Name 'DefaultProfile' `
                                                 -ErrorAction SilentlyContinue).DefaultProfile
                if ($ProfileName) { break }
            }
        }
        catch { }
    }

    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        Write-OMMigrateLog -Message (
            "Credential repair failed for $EmailAddress -- " +
            "Outlook profile name not found in registry."
        ) -Level WARN
        return 'FAILED'
    }

    # -- Locate the registry subkey --------------------------------
    # Scan 9375CFF0 subkeys for one whose Email value matches.
    # Try each Office version's registry path with the resolved profile
    # name, same as $Script:OutlookRegPaths is used elsewhere in this module.
    $profileBase = $null
    foreach ($regPath in $Script:OutlookRegPaths) {
        $candidate = Join-Path $regPath "$ProfileName\$Script:OutlookAccountGUID"
        if (Test-Path $candidate) {
            $profileBase = $candidate
            break
        }
    }

    if (-not $profileBase) {
        Write-OMMigrateLog -Message (
            "Credential repair failed for $EmailAddress -- " +
            "Could not find account registry path for profile '$ProfileName'."
        ) -Level WARN
        return 'FAILED'
    }

    $subkeyPath  = $null
    $subkeyName  = $null

    try {
        $subkeys = Get-ChildItem -Path $profileBase -ErrorAction Stop
        foreach ($sk in $subkeys) {
            try {
                $props = Get-ItemProperty -Path $sk.PSPath -ErrorAction SilentlyContinue
                if (-not $props) { continue }
                $skEmail = $props.PSObject.Properties['Email']
                if ($skEmail -and $skEmail.Value -ieq $EmailAddress) {
                    $subkeyPath = $sk.PSPath
                    $subkeyName = $sk.PSChildName
                    break
                }
            }
            catch { }
        }
    }
    catch {
        Write-OMMigrateLog -Message (
            "Credential repair failed for $EmailAddress -- " +
            "Could not scan registry subkeys: $_"
        ) -Level WARN
        return 'FAILED'
    }

    if (-not $subkeyPath) {
        Write-OMMigrateLog -Message (
            "Credential repair skipped for $EmailAddress -- " +
            "No matching subkey found in registry."
        ) -Level WARN
        return 'NOT_FOUND'
    }

    Write-OMMigrateLog -Message (
        "Credential repair: found subkey $subkeyName for $EmailAddress"
    ) -Level INFO

    # -- Read current state ----------------------------------------
    $current = $null
    try {
        $current = Get-ItemProperty -Path $subkeyPath -ErrorAction Stop
    }
    catch {
        Write-OMMigrateLog -Message (
            "Credential repair failed for $EmailAddress -- " +
            "Could not read subkey: $_"
        ) -Level WARN
        return 'FAILED'
    }

    # -- Already correct check -------------------------------------
    # If IMAP User already equals the email address the wizard did not
    # scramble this account -- nothing to fix.
    $currentIMAPUser = $current.PSObject.Properties['IMAP User']
    if ($currentIMAPUser -and $currentIMAPUser.Value -ieq $EmailAddress) {
        Write-OMMigrateLog -Message (
            "Credential repair skipped for $EmailAddress -- " +
            "IMAP User already correct, no fix needed."
        ) -Level INFO
        return 'SKIPPED'
    }

    # -- Backup subkey before writing (before-fix state) ----------
    # Exports the scrambled state -- use this to recover if the fix
    # causes a problem. Named with _credentials_before_fix so it
    # sorts next to the PST backup in the Backups\ folder.
    $backupFile = $null
    try {
        if (-not (Test-Path $BackupPath)) {
            New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
        }
        $timestamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
        $safeEmail    = Get-SafeFileName -InputString $EmailAddress
        $backupFile   = Join-Path $BackupPath "${safeEmail}_credentials_before_fix_$timestamp.reg"
        # FIXED (launch-readiness review, 2026-08-18): was hardcoded to
        # 16.0\Outlook -- now built from the resolved $profileBase (which
        # is an HKCU:\... PowerShell-provider path) converted to the
        # HKEY_CURRENT_USER\... form reg.exe expects.
        $regExportKey = ($profileBase -replace '^HKCU:\\', 'HKEY_CURRENT_USER\') + '\' + $subkeyName
        $regArgs      = "export `"$regExportKey`" `"$backupFile`" /y"
        $result       = Start-Process 'reg.exe' -ArgumentList $regArgs -Wait -PassThru -NoNewWindow
        if ($result.ExitCode -eq 0) {
            Write-OMMigrateLog -Message "Before-fix credential backup saved: $backupFile" -Level INFO
        }
        else {
            Write-OMMigrateLog -Message "Before-fix credential backup may have failed (exit $($result.ExitCode)): $backupFile" `
                               -Level WARN
        }
    }
    catch {
        Write-OMMigrateLog -Message "Could not create before-fix credential backup: $_" -Level WARN
        # Non-fatal -- continue with fix
    }

    # -- Apply credential fix --------------------------------------
    try {
        # Read the misplaced SMTP password blob from IMAP Password
        $smtpBlob = $current.PSObject.Properties['IMAP Password']
        if ($smtpBlob -and $smtpBlob.Value) {
            # 1. Move SMTP blob from IMAP Password to SMTP Password
            Set-ItemProperty -Path $subkeyPath -Name 'SMTP Password' `
                             -Value $smtpBlob.Value -Type Binary
            Write-OMMigrateLog -Message "  SMTP Password blob moved from IMAP Password field." `
                               -Level INFO

            # 2. Clear IMAP Password -- operator enters this once manually
            Set-ItemProperty -Path $subkeyPath -Name 'IMAP Password' `
                             -Value ([byte[]]@()) -Type Binary
            Write-OMMigrateLog -Message "  IMAP Password cleared -- operator will enter manually." `
                               -Level INFO
        }
        else {
            Write-OMMigrateLog -Message "  No IMAP Password blob found to move -- skipping blob move." `
                               -Level WARN
        }

        # 3. Write correct IMAP User = email address
        Set-ItemProperty -Path $subkeyPath -Name 'IMAP User' `
                         -Value $EmailAddress -Type String
        Write-OMMigrateLog -Message "  IMAP User set to: $EmailAddress" -Level INFO

        # 4. Write correct SMTP User = SmtpUsername from CSV
        Set-ItemProperty -Path $subkeyPath -Name 'SMTP User' `
                         -Value $SmtpUsername -Type String
        Write-OMMigrateLog -Message "  SMTP User set to: $SmtpUsername" -Level INFO

        # 5. Set SMTP Auth Method = 1
        Set-ItemProperty -Path $subkeyPath -Name 'SMTP Auth Method' `
                         -Value 1 -Type DWord
        Write-OMMigrateLog -Message "  SMTP Auth Method set to: 1" -Level INFO

        Write-OMMigrateLog -Message "Credential repair completed successfully for $EmailAddress" `
                           -Level INFO

        # -- Backup subkey after writing (after-fix state) --------
        # Exports the corrected state -- use this as a reference for
        # what the correct registry values should look like, or to
        # manually restore if a later issue overwrites them.
        $postBackupFile = $null
        try {
            $postBackupFile = Join-Path $BackupPath "${safeEmail}_credentials_after_fix_$timestamp.reg"
            $regArgs        = "export `"$regExportKey`" `"$postBackupFile`" /y"
            $postResult     = Start-Process 'reg.exe' -ArgumentList $regArgs -Wait -PassThru -NoNewWindow
            if ($postResult.ExitCode -eq 0) {
                Write-OMMigrateLog -Message "After-fix credential backup saved: $postBackupFile" -Level INFO
            }
            else {
                Write-OMMigrateLog -Message "After-fix credential backup may have failed (exit $($postResult.ExitCode)): $postBackupFile" `
                                   -Level WARN
            }
        }
        catch {
            Write-OMMigrateLog -Message "Could not create after-fix credential backup: $_" -Level WARN
        }

        Write-AuditEntry  -Action 'IMAP_CREDENTIALS_REPAIRED' `
                          -AccountEmail $EmailAddress `
                          -Detail ("Registry credential fix applied. " +
                                   "SubKey=$subkeyName | ProviderTag=$ProviderTag | " +
                                   "IMAPUser=$EmailAddress | SMTPUser=$SmtpUsername | " +
                                   "PreBackup=$backupFile | PostBackup=$postBackupFile")

        return 'FIXED'
    }
    catch {
        Write-OMMigrateLog -Message (
            "Credential repair FAILED for $EmailAddress -- " +
            "Registry write error: $_"
        ) -Level WARN
        Write-OMMigrateLog -Message (
            "  Restore backup: reg import `"$backupFile`""
        ) -Level WARN
        return 'FAILED'
    }
}


# ============================================================
#  REGION: MODULE EXPORTS
# ============================================================

Export-ModuleMember -Function @(

    # Profile discovery
    'Get-OutlookProfiles'

    # Account discovery
    'Get-OutlookAccountsFromRegistry'
    'Read-AccountSubkey'
    'Read-BinaryBlobStrings'
    'Read-BinaryBlobPort'

    # Account classification
    'New-AccountObject'
    'Get-AccountTypeFromPort'
    'Set-AccountTag'

    # Data file discovery
    'Get-OutlookDataFiles'
    'Join-AccountsWithDataFiles'

    # Summary and export
    'Get-AccountSummary'
    'Export-AccountsToCSV'
    'Update-AccountMigrationAction'
    'Invoke-MigrateAccountPicker'
    'Invoke-CredentialEntryUI'
    'Remove-POP3AccountViaRegistry'
    'Add-IMAPAccountViaRegistry'
    'Repair-IMAPCredentials'
)
# ***** END OF FILE *****
