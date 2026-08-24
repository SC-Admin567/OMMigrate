<#
.SYNOPSIS
    Obsidian-Addon-Opener.ps1 -- resolves the correct Obsidian vault for a
    clicked .md file and opens it there; falls back to Notepad if the file
    is not inside an Obsidian vault.

.NOTES
    -------------------------------------------------------------------------
    OutlookMailMigrator (OMMigrate) -- Obsidian Add-on
    -------------------------------------------------------------------------
    Originator & Architect:    Kirk Shallcross - Shallcross Consulting
    Implementation Specialist: Anthropic Claude AI
    Inception Date:            May 2026
    Version:                   1.5.2
    -------------------------------------------------------------------------
#>

Param([string]$FilePath)
if (-not $FilePath -or -not (Test-Path -LiteralPath $FilePath -PathType Leaf)) { exit }
$FilePath = (Get-Item -LiteralPath $FilePath).FullName
$CurrDir = Split-Path $FilePath -Parent
$VaultFound = $false

while ($CurrDir -and (Split-Path $CurrDir -Leaf) -ne "") {
    if (Test-Path -LiteralPath "$CurrDir\.obsidian") {
        $Vault   = Split-Path $CurrDir -Parent -Leaf
        $RelPath = $FilePath.Substring($CurrDir.Length + 1).Replace('\', '/')

        $EncodedVault   = [System.Uri]::EscapeDataString($Vault)
        $EncodedRelPath = [System.Uri]::EscapeDataString($RelPath)
        $ObsidianUri    = "obsidian://open?vault=$EncodedVault&file=$EncodedRelPath"

        try {
            Start-Process -FilePath $ObsidianUri -ErrorAction Stop
        }
        catch {
            Start-Process -FilePath "C:\Windows\System32\notepad.exe" -ArgumentList $FilePath -ErrorAction SilentlyContinue
        }

        $VaultFound = $true
        exit
    }
    $CurrDir = Split-Path $CurrDir -Parent
}
if (-not $VaultFound) {
    Start-Process -FilePath "C:\Windows\System32\notepad.exe" -ArgumentList $FilePath -ErrorAction SilentlyContinue
}