#Requires -Version 5.1
<#
.SYNOPSIS
    Update-Version.ps1 -- OMMigrate Version Number Updater

.DESCRIPTION
    Updates the version number across all OMMigrate scripts, modules,
    and documentation files in one operation.

    This is the ONLY place you ever need to change a version number.
    Everything else reads from version.txt at runtime or is updated
    by this script.

    What this script does:
        1. Reads the current version from version.txt
        2. Prompts for the new version number
        3. Writes the new version to version.txt
        4. Updates all .ps1, .psm1, .md, .bas, and .cls files:
               - .NOTES Version: lines in script/module headers
               - $Script:OMMigrateVersion in OMMigrate-Core.psm1
        5. Reports exactly which files and lines were changed

    What reads version.txt at runtime (no update needed):
        OMMigrate-Core.psm1 reads version.txt when imported.
        All banners, logs, reports, manifests, and the settings file
        then show the correct version automatically.

    When to run this script:
        Any time you increment the version number. Run it once,
        confirm the changes, and all files are consistent.

.EXAMPLE
    # From the OMMigrate project root:
    .\Update-Version.ps1

.NOTES
    -------------------------------------------------------------------------
    OutlookMailMigrator (OMMigrate)
    -------------------------------------------------------------------------
    Originator & Architect:    Kirk Shallcross - Shallcross Consulting
    Implementation Specialist: Anthropic Claude AI
    Inception Date:            May 2026
    Version:                   1.5.2    
    -------------------------------------------------------------------------

    Run from the OMMigrate project root directory (where version.txt lives).
    Safe to run multiple times -- idempotent if version number is unchanged.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptRoot = $PSScriptRoot

# ── Read current version from version.txt ────────────────────────────────────
$versionFile = Join-Path $ScriptRoot 'version.txt'

if (-not (Test-Path $versionFile)) {
    Write-Host ''
    Write-Host '  ERROR: version.txt not found.' -ForegroundColor Red
    Write-Host "  Expected: $versionFile" -ForegroundColor Red
    Write-Host '  This file must exist in the OMMigrate project root.' -ForegroundColor Red
    Write-Host ''
    exit 1
}

$currentVersion = (Get-Content $versionFile -Raw).Trim()

# ── Display current state ─────────────────────────────────────────────────────
$line = '-' * 60
Write-Host ''
Write-Host $line -ForegroundColor DarkCyan
Write-Host '  OMMigrate Version Updater' -ForegroundColor White
Write-Host $line -ForegroundColor DarkCyan
Write-Host "  Current version : $currentVersion" -ForegroundColor Cyan
Write-Host ''

# ── Prompt for new version ────────────────────────────────────────────────────
Write-Host '  Enter new version number (e.g. 1.0.0)' -ForegroundColor Yellow
Write-Host '  Press Enter to cancel without making changes.' -ForegroundColor DarkGray
Write-Host ''
Write-Host '  New version : ' -ForegroundColor Yellow -NoNewline
$newVersion = Read-Host

if ([string]::IsNullOrWhiteSpace($newVersion)) {
    Write-Host ''
    Write-Host '  No version entered -- no changes made.' -ForegroundColor DarkGray
    Write-Host ''
    exit 0
}

$newVersion = $newVersion.Trim()

# Basic format validation -- must be x.y.z
if ($newVersion -notmatch '^\d+\.\d+\.\d+$') {
    Write-Host ''
    Write-Host "  ERROR: '$newVersion' is not a valid version format." -ForegroundColor Red
    Write-Host '  Version must be in x.y.z format (e.g. 1.2.0).' -ForegroundColor Red
    Write-Host ''
    exit 1
}

if ($newVersion -eq $currentVersion) {
    Write-Host ''
    Write-Host "  Version is already $currentVersion -- no changes needed." -ForegroundColor Green
    Write-Host ''
    exit 0
}

Write-Host ''
Write-Host "  Updating: $currentVersion  ->  $newVersion" -ForegroundColor White
Write-Host ''

# ── Confirm before making changes ────────────────────────────────────────────
Write-Host '  Files that will be updated:' -ForegroundColor Gray

# Collect target files -- all .ps1, .psm1, .md, .bas, .cls in project tree
$targetExtensions = @('*.ps1', '*.psm1', '*.md', '*.bas', '*.cls')
$targetFiles = @()
foreach ($ext in $targetExtensions) {
    $found = Get-ChildItem -Path $ScriptRoot -Filter $ext -Recurse -ErrorAction SilentlyContinue |
             Where-Object { $_.FullName -notlike '*\.git*' }
    $targetFiles += $found
}

foreach ($f in $targetFiles) {
    Write-Host "    $($f.Name)" -ForegroundColor DarkGray
}
Write-Host "    version.txt" -ForegroundColor DarkGray
Write-Host ''

Write-Host '  Proceed? [Y/n] : ' -ForegroundColor Yellow -NoNewline
$confirm = Read-Host
if ($confirm -ne '' -and $confirm -notmatch '^[Yy]') {
    Write-Host ''
    Write-Host '  Cancelled -- no changes made.' -ForegroundColor DarkGray
    Write-Host ''
    exit 0
}

Write-Host ''

# ── Apply version updates  Start at 1 to include myself ───────────────────────────────────
$totalReplacements = 0
$filesChanged      = 1

foreach ($file in $targetFiles) {

    $content     = Get-Content $file.FullName -Raw -Encoding UTF8
    $newContent  = $content
    $fileChanged = $false

    # Pattern 1: .NOTES Version: line in script/module headers
    # Matches:   "    Version:                   1.5.2"
    # (any amount of whitespace before Version:, any spacing, then the version)
    $newContent = $newContent -replace `
        '([ \t]+Version\s*:\s*)(\d+\.\d+\.\d+)', `
        ('${1}' + $newVersion)

    # Pattern 2a: $Script:OMMigrateVersion = 'x.y.z'  (OMMigrate-Core.psm1)
    $newContent = $newContent -replace `
        "(\`$Script:OMMigrateVersion\s*=\s*')(\d+\.\d+\.\d+)(')", `
        ('${1}' + $newVersion + '${3}')

    # Pattern 2b: $Script:Version = 'x.y.z'  (Install.ps1)
    $newContent = $newContent -replace `
        "(\`$Script:Version\s*=\s*')(\d+\.\d+\.\d+)(')", `
        ('${1}' + $newVersion + '${3}')

    # Pattern 3: Version: x.y.z in markdown files (no leading spaces required)
    $newContent = $newContent -replace `
        '(Version:\s*)(\d+\.\d+\.\d+)', `
        ('${1}' + $newVersion)

    # Pattern 4: OMMigrate) vx.y.z  in markdown footer code blocks
    $newContent = $newContent -replace `
        '(OMMigrate\) v)(\d+\.\d+\.\d+)', `
        ('${1}' + $newVersion)

    # Pattern 5: Current Version: x.y.z  in CHANGELOG.md blockquote
    $newContent = $newContent -replace `
        '(Current Version:\s*)(\d+\.\d+\.\d+)', `
        ('${1}' + $newVersion)

    # Pattern 6: **Version:** x.y.z  in README.md bold markdown header
    $newContent = $newContent -replace `
        '(\*\*Version:\*\*\s*)(\d+\.\d+\.\d+)', `
        ('${1}' + $newVersion)

    if ($newContent -ne $content) {
        # Count replacements made in this file
        $oldMatches = ([regex]::Matches($content,   '\d+\.\d+\.\d+')).Count
        $newMatches = ([regex]::Matches($newContent, $newVersion.Replace('.', '\.'))).Count

        # Write back preserving encoding
        # Check for BOM
        $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
        $hasBOM = ($bytes.Length -ge 3 -and
                   $bytes[0] -eq 0xEF -and
                   $bytes[1] -eq 0xBB -and
                   $bytes[2] -eq 0xBF)

        if ($hasBOM) {
            [System.IO.File]::WriteAllText(
                $file.FullName,
                $newContent,
                [System.Text.UTF8Encoding]::new($true)   # UTF-8 with BOM
            )
        }
        else {
            [System.IO.File]::WriteAllText(
                $file.FullName,
                $newContent,
                [System.Text.UTF8Encoding]::new($false)  # UTF-8 without BOM
            )
        }

        Write-Host "  [UPDATED] $($file.Name)" -ForegroundColor Green
        $filesChanged++
        $totalReplacements++
    }
    else {
        Write-Host "  [NO CHANGE] $($file.Name)" -ForegroundColor DarkGray
    }
}

# ── Update version.txt ────────────────────────────────────────────────────────
Set-Content -Path $versionFile -Value $newVersion -Encoding UTF8 -NoNewline
Write-Host '  [UPDATED] version.txt' -ForegroundColor Green

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host $line -ForegroundColor DarkCyan
Write-Host "  Version updated: $currentVersion  ->  $newVersion" -ForegroundColor Green
Write-Host "  Files changed  : $filesChanged" -ForegroundColor Green
Write-Host $line -ForegroundColor DarkCyan
Write-Host ''
Write-Host '  All done. Close and reopen PowerShell before running any OMMigrate script.' `
           -ForegroundColor Cyan
Write-Host ''
