<#
.SYNOPSIS
    show_tree.ps1 -- generates tree_view.txt, an ASCII directory listing
    of the OMMigrate project folder.

.NOTES
    -------------------------------------------------------------------------
    OutlookMailMigrator (OMMigrate)
    -------------------------------------------------------------------------
    Originator & Architect:    Kirk Shallcross - Shallcross Consulting
    Implementation Specialist: Anthropic Claude AI
    Inception Date:            May 2026
    Version:                   1.5.2
    -------------------------------------------------------------------------
#>

function Show-CustomTree {
    param([string]$Path = ".", [string]$Indent = "")
    Get-ChildItem -Path $Path | ForEach-Object {
        "$Indent+-- " + $_.Name
        if ($_.PSIsContainer -and $_.Name -ne ".git") {
            Show-CustomTree -Path $_.FullName -Indent "$Indent|   "
        }
    }
}
Show-CustomTree | Out-File ".\tree_view.txt"