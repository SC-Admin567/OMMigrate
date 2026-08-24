#Requires -Version 5.1
<#
.SYNOPSIS
    Install.ps1 -- OMMigrate one-command installer.

.DESCRIPTION
    Downloads all OMMigrate scripts and modules from GitHub, creates the
    project folder structure in your Documents folder, and prepares everything to run.

    This is the only file you need. Run it once and you are ready to migrate.

    USAGE -- Run this single command in PowerShell:

        irm https://raw.githubusercontent.com/SC-Admin567/OMMigrate/main/Install.ps1 | iex

    Or if you downloaded this file manually:

        .\Install.ps1

.PARAMETER ProjectRoot
    Where to create the OMMigrate project folder.
    Default: your Documents\OMMigrate\

.PARAMETER Branch
    GitHub branch to download from.
    Default: main

.PARAMETER Force
    Re-download and overwrite all files even if they already exist.
    Use this to update to the latest version from GitHub.
    Default: skip files that already exist (safe re-run behavior)

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

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ProjectRoot = "$env:USERPROFILE\Documents\OMMigrate",

    [Parameter(Mandatory = $false)]
    [string]$Branch = 'main',

    [Parameter(Mandatory = $false)]
    [switch]$Force
)


# ============================================================
#  CONFIGURATION
# ============================================================

$Script:Version   = '1.5.2'
$Script:Product   = 'OutlookMailMigrator (OMMigrate)'
$Script:Architect = 'Kirk Shallcross - Shallcross Consulting'
$Script:AI        = 'Anthropic Claude AI'
$Script:GitHubUrl = 'https://github.com/SC-Admin567/OMMigrate'
$Script:RawBase   = "https://raw.githubusercontent.com/SC-Admin567/OMMigrate/$Branch"

# Runtime data location -- mirrors OMMigrate-Core.psm1
$Script:RuntimeBase = "$env:USERPROFILE\Documents\OutlookMigration"

# Files to download from GitHub -- path in repo maps to local destination
$Script:FilesToDownload = @(
    # Scripts
    @{ Repo = 'Scripts/OMMigrate-00-Discover.ps1';  Local = 'Scripts\OMMigrate-00-Discover.ps1'  }
    @{ Repo = 'Scripts/OMMigrate-01-Backup.ps1';    Local = 'Scripts\OMMigrate-01-Backup.ps1'    }
    @{ Repo = 'Scripts/OMMigrate-02-Convert.ps1';   Local = 'Scripts\OMMigrate-02-Convert.ps1'   }
    @{ Repo = 'Scripts/OMMigrate-03-Restore.ps1';   Local = 'Scripts\OMMigrate-03-Restore.ps1'   }
    @{ Repo = 'Scripts/OMMigrate-04-Artifacts.ps1'; Local = 'Scripts\OMMigrate-04-Artifacts.ps1' }

    # Modules
    @{ Repo = 'Modules/OMMigrate-Core.psm1';        Local = 'Modules\OMMigrate-Core.psm1'        }
    @{ Repo = 'Modules/OMMigrate-Registry.psm1';    Local = 'Modules\OMMigrate-Registry.psm1'    }
    @{ Repo = 'Modules/OMMigrate-Outlook.psm1';     Local = 'Modules\OMMigrate-Outlook.psm1'     }
    @{ Repo = 'Modules/OMMigrate-Reporting.psm1';   Local = 'Modules\OMMigrate-Reporting.psm1'   }

    # Utility scripts
    @{ Repo = 'Utilities/OMMigrate-FindDuplicateRules.ps1';   Local = 'Utilities\OMMigrate-FindDuplicateRules.ps1'   }
    @{ Repo = 'Utilities/OMMigrate-RemoveDuplicateRules.ps1'; Local = 'Utilities\OMMigrate-RemoveDuplicateRules.ps1' }
    @{ Repo = 'Utilities/OMMigrate-CleanRulesCSV.ps1';        Local = 'Utilities\OMMigrate-CleanRulesCSV.ps1'        }

    # Outlook VBA macros (manual import required -- see QUICKSTART.md Step 0.5)
    @{ Repo = 'OutlookVBAMacros/Module1.bas';                 Local = 'OutlookVBAMacros\Module1.bas'                 }
    @{ Repo = 'OutlookVBAMacros/Module2.bas';                 Local = 'OutlookVBAMacros\Module2.bas'                 }
    @{ Repo = 'OutlookVBAMacros/Module3.bas';                 Local = 'OutlookVBAMacros\Module3.bas'                 }
    @{ Repo = 'OutlookVBAMacros/Module7.bas';                 Local = 'OutlookVBAMacros\Module7.bas'                 }
    @{ Repo = 'OutlookVBAMacros/ThisOutlookSession.cls';      Local = 'OutlookVBAMacros\ThisOutlookSession.cls'      }
    @{ Repo = 'OutlookVBAMacros/FrmArchivePicker.frm';        Local = 'OutlookVBAMacros\FrmArchivePicker.frm'        }
    @{ Repo = 'OutlookVBAMacros/FrmArchivePicker.frx';        Local = 'OutlookVBAMacros\FrmArchivePicker.frx'        }
    @{ Repo = 'OutlookVBAMacros/FrmProgress.frm';             Local = 'OutlookVBAMacros\FrmProgress.frm'             }
    @{ Repo = 'OutlookVBAMacros/FrmProgress.frx';             Local = 'OutlookVBAMacros\FrmProgress.frx'             }

    # Documentation
    @{ Repo = 'README.md';                          Local = 'README.md'                          }
    @{ Repo = 'QUICKSTART.md';                      Local = 'QUICKSTART.md'                      }
    @{ Repo = 'CHANGELOG.md';                       Local = 'CHANGELOG.md'                       }
    @{ Repo = 'LICENSE.md';                         Local = 'LICENSE.md'                         }
    @{ Repo = 'CONTRIBUTING.md';                    Local = 'CONTRIBUTING.md'                    }

    # GitHub Issue Forms -- no effect on running the tool, included for
    # tree_view.txt / repo parity, same as .gitattributes / .gitignore
    @{ Repo = '.github/ISSUE_TEMPLATE/bug_report.yml';       Local = '.github\ISSUE_TEMPLATE\bug_report.yml'       }
    @{ Repo = '.github/ISSUE_TEMPLATE/feature_request.yml'; Local = '.github\ISSUE_TEMPLATE\feature_request.yml' }
    @{ Repo = '.github/ISSUE_TEMPLATE/config.yml';           Local = '.github\ISSUE_TEMPLATE\config.yml'           }
    @{ Repo = 'OMMigrate_CommandLine_Reference.md'; Local = 'OMMigrate_CommandLine_Reference.md' }
    @{ Repo = 'OMMigrate_Settings_Reference.md';    Local = 'OMMigrate_Settings_Reference.md'    }

    # Documentation images -- embedded in README.md / QUICKSTART.md
    @{ Repo = 'docs_images/architecture_overview.svg';  Local = 'docs_images\architecture_overview.svg'  }
    @{ Repo = 'QUICKSTART_images/snap-01.jpg';          Local = 'QUICKSTART_images\snap-01.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-02.jpg';          Local = 'QUICKSTART_images\snap-02.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-03.jpg';          Local = 'QUICKSTART_images\snap-03.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-05.jpg';          Local = 'QUICKSTART_images\snap-05.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-06.jpg';          Local = 'QUICKSTART_images\snap-06.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-07.jpg';          Local = 'QUICKSTART_images\snap-07.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-08.jpg';          Local = 'QUICKSTART_images\snap-08.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-09.jpg';          Local = 'QUICKSTART_images\snap-09.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-10.jpg';          Local = 'QUICKSTART_images\snap-10.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-11.jpg';          Local = 'QUICKSTART_images\snap-11.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-12.jpg';          Local = 'QUICKSTART_images\snap-12.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-13.jpg';          Local = 'QUICKSTART_images\snap-13.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-14.jpg';          Local = 'QUICKSTART_images\snap-14.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-15.jpg';          Local = 'QUICKSTART_images\snap-15.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-16.jpg';          Local = 'QUICKSTART_images\snap-16.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-17.jpg';          Local = 'QUICKSTART_images\snap-17.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-18.jpg';          Local = 'QUICKSTART_images\snap-18.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-19.jpg';          Local = 'QUICKSTART_images\snap-19.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-21.jpg';          Local = 'QUICKSTART_images\snap-21.jpg'          }
    @{ Repo = 'QUICKSTART_images/snap-22.png';          Local = 'QUICKSTART_images\snap-22.png'          }
    @{ Repo = 'QUICKSTART_images/snap-23.png';          Local = 'QUICKSTART_images\snap-23.png'          }

    # Donation button embed -- referenced by README.md via ![[stripe-button.html]]
    @{ Repo = 'stripe-button.html';                     Local = 'stripe-button.html'                     }

    # Repo-maintenance files -- not required to run OMMigrate, included for
    # parity with tree_view.txt / contributors who installed via Install.ps1
    # rather than git clone
    @{ Repo = '.gitattributes';                          Local = '.gitattributes'                         }
    @{ Repo = '.gitignore';                              Local = '.gitignore'                             }
    @{ Repo = 'version.txt';                             Local = 'version.txt'                            }
    @{ Repo = 'Update-Version.ps1';                      Local = 'Update-Version.ps1'                     }
    @{ Repo = 'show_tree.ps1';                           Local = 'show_tree.ps1'                          }
    @{ Repo = 'tree_view.txt';                           Local = 'tree_view.txt'                          }

    # Obsidian Add-on (optional -- free third-party .md viewer, see README.md
    # Third-Party Software Notice). Installed like the VBA macros in
    # OutlookVBAMacros -- downloaded here, but activated only if the operator
    # chooses to run Obsidian-Addon-Install.bat. See QUICKSTART.md Step 0.75.
    @{ Repo = 'Obsidian-Addon-Install.bat';          Local = 'Obsidian-Addon-Install.bat'          }
    @{ Repo = 'Obsidian-Addon-Uninstall.bat';        Local = 'Obsidian-Addon-Uninstall.bat'        }
    @{ Repo = 'Obsidian-Addon-DeployLink.reg';       Local = 'Obsidian-Addon-DeployLink.reg'       }
    @{ Repo = 'Obsidian-Addon-UninstallLink.reg';    Local = 'Obsidian-Addon-UninstallLink.reg'    }
    @{ Repo = 'Obsidian-Addon-Opener.ps1';           Local = 'Obsidian-Addon-Opener.ps1'           }
    @{ Repo = 'Obsidian-Addon-SilentRunner.vbs';     Local = 'Obsidian-Addon-SilentRunner.vbs'     }

    # Obsidian vault config -- HTML Viewer Plus plugin (renders the donate
    # button / stripe-button.html embed in README.md when opened in Obsidian)
    # NOTE: app.json / appearance.json / workspace.json are intentionally NOT
    # downloaded here -- .gitignore excludes them (per-user Obsidian state,
    # Administrator's explicit direction 2026-08-18), so they 404 on GitHub.
    @{ Repo = '.obsidian/community-plugins.json';                      Local = '.obsidian\community-plugins.json'                      }
    @{ Repo = '.obsidian/core-plugins.json';                           Local = '.obsidian\core-plugins.json'                           }
    @{ Repo = '.obsidian/plugins/html-viewer-plus/main.js';            Local = '.obsidian\plugins\html-viewer-plus\main.js'            }
    @{ Repo = '.obsidian/plugins/html-viewer-plus/manifest.json';      Local = '.obsidian\plugins\html-viewer-plus\manifest.json'      }
    @{ Repo = '.obsidian/plugins/html-viewer-plus/styles.css';         Local = '.obsidian\plugins\html-viewer-plus\styles.css'         }
    @{ Repo = '.obsidian/plugins/html-viewer-plus/data.json';          Local = '.obsidian\plugins\html-viewer-plus\data.json'          }
    @{ Repo = '.obsidian/plugins/html-viewer-plus/debug.log';          Local = '.obsidian\plugins\html-viewer-plus\debug.log'          }
)

# Folders to create in the project directory
$Script:ProjectFolders = @('Scripts', 'Modules', 'Utilities', 'OutlookVBAMacros')

# Runtime folders created at Documents\OutlookMigration\ -- not in project
$Script:RuntimeFolders = @('Config', 'Logs', 'Manifests', 'Reports', 'Backups')


# ============================================================
#  HELPER FUNCTIONS
# ============================================================

function Write-Step {
    param([string]$Message, [string]$Color = 'Cyan')
    Write-Host "  $Message" -ForegroundColor $Color
}

function Write-Banner {
    $sep = '-' * 70
    Write-Host ''
    Write-Host $sep -ForegroundColor DarkCyan
    Write-Host "  $Script:Product" -ForegroundColor White
    Write-Host "  Installer v$Script:Version" -ForegroundColor White
    Write-Host $sep -ForegroundColor DarkCyan
    Write-Host "  Project folder : $ProjectRoot" -ForegroundColor Cyan
    Write-Host "  Runtime data   : $Script:RuntimeBase" -ForegroundColor Cyan
    Write-Host "  GitHub source  : $Script:GitHubUrl" -ForegroundColor Cyan
    if ($Force) {
        Write-Host "  Mode           : Force -- all files will be overwritten" -ForegroundColor Yellow
    }
    else {
        Write-Host "  Mode           : Safe -- existing files will be skipped" -ForegroundColor Cyan
    }
    Write-Host $sep -ForegroundColor DarkCyan
    Write-Host ''
}

function New-Dir {
    param([string]$Path, [string]$Label)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-Step "Created  : $Label" 'Green'
    }
    else {
        Write-Step "Exists   : $Label" 'DarkGray'
    }
}

# Files that are intentionally 0 bytes in the repo -- Get-GitHubFile skips
# the empty-download check for these so a correct empty download isn't
# flagged as a failure. Currently just the HTML Viewer Plus debug log,
# which .gitignore's own comment notes is "Empty by default -- enable
# debug mode in the plugin to populate."
$Script:ExpectedEmptyFiles = @('debug.log')

function Get-GitHubFile {
    <#
    .SYNOPSIS
        Downloads a single file from GitHub raw content URL.
        Skips existing files unless -Force is specified.
        Returns $true on success or skip, $false on failure.
    #>
    param(
        [string]$RepoPath,
        [string]$LocalPath
    )

    $url      = "$Script:RawBase/$RepoPath"
    $fullPath = Join-Path $ProjectRoot $LocalPath
    $fileName = Split-Path $LocalPath -Leaf

    # Skip existing files unless -Force was specified
    if ((Test-Path $fullPath) -and -not $Force) {
        Write-Step "Skipped  : $fileName  (already exists -- use -Force to overwrite)" 'DarkGray'
        return $true
    }

    try {
        # Download to a temp file first -- never corrupt the existing file on failure
        $tempFile = [System.IO.Path]::GetTempFileName()

        Invoke-WebRequest -Uri $url -OutFile $tempFile -UseBasicParsing -ErrorAction Stop

        # Verify download has content -- only reject truly empty/failed downloads.
        # NOTE: some legitimately tracked repo files (.gitattributes, .gitignore,
        # version.txt, small Obsidian JSON configs) are well under 100 bytes;
        # a flat 100-byte floor was flagging those as false-positive failures.
        # Zero bytes is the real failure signal -- except for files in
        # $Script:ExpectedEmptyFiles, which are intentionally 0 bytes in the repo.
        $tempSize = (Get-Item $tempFile).Length
        if ($tempSize -eq 0 -and $fileName -notin $Script:ExpectedEmptyFiles) {
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            Write-Step "FAILED   : $fileName  (download empty -- check GitHub URL)" 'Red'
            return $false
        }

        # Move temp file to destination -- atomic replacement
        if (Test-Path $fullPath) {
            Remove-Item $fullPath -Force
        }
        Move-Item -Path $tempFile -Destination $fullPath -Force

        Write-Step "Downloaded: $fileName" 'Green'
        return $true
    }
    catch {
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        Write-Step "FAILED   : $fileName  ($_)" 'Red'
        return $false
    }
}

function Invoke-UnblockAll {
    <#
    .SYNOPSIS
        Removes the Windows Zone.Identifier security block from all
        downloaded .ps1, .psm1, .bas, .cls, .frm, .bat, .reg, and .vbs files.
        (.frx files are binary and are not included -- nothing to unblock.)
    #>
    Write-Host ''
    Write-Host '  [ Unblocking downloaded files ]' -ForegroundColor White
    Write-Host ''

    $count = 0
    Get-ChildItem -Path $ProjectRoot -Recurse -Include '*.ps1','*.psm1','*.bas','*.cls','*.frm','*.bat','*.reg','*.vbs' |
        ForEach-Object {
            try {
                Unblock-File -Path $_.FullName -ErrorAction Stop
                $count++
            }
            catch {
                Write-Step "WARNING  : Could not unblock $($_.Name) -- $_" 'Yellow'
            }
        }

    Write-Step "Unblocked : $count file(s)" 'Green'
}

function Set-ExecutionPolicyIfNeeded {
    <#
    .SYNOPSIS
        Sets RemoteSigned execution policy for current user if not already set.
        Required for PowerShell to run local unsigned scripts.
    #>
    $policy = Get-ExecutionPolicy -Scope CurrentUser
    if ($policy -in @('Restricted', 'Undefined', 'AllSigned')) {
        try {
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
            Write-Step "Execution policy set to RemoteSigned for current user" 'Green'
        }
        catch {
            Write-Step "WARNING  : Could not set execution policy -- $_" 'Yellow'
            Write-Step "           Run manually: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser" 'Yellow'
        }
    }
    else {
        Write-Step "Execution policy OK: $policy" 'DarkGray'
    }
}

function Invoke-WaitForOutlookClose {
    <#
    .SYNOPSIS
        Detects if Outlook is running and prompts the operator to close it
        before the Archive PST can be created via COM.

        Loops until Outlook is confirmed closed or the operator aborts.
        Returns $true when Outlook is confirmed not running, $false if aborted.
    #>

    while ($true) {
        $outlookProcess = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
        if (-not $outlookProcess) {
            # Outlook is not running -- clear to proceed
            return $true
        }

        # Outlook is open -- prompt operator
        Write-Host ''
        Write-Host '  ---------------------------------------------------------------' -ForegroundColor Yellow
        Write-Host '  Outlook is currently open.' -ForegroundColor Yellow
        Write-Host '  Outlook must be closed before the Archive PST can be created.' -ForegroundColor Yellow
        Write-Host '  ---------------------------------------------------------------' -ForegroundColor Yellow
        Write-Host ''
        Write-Host '  Alt+Tab to Outlook, close it completely, then return here.' -ForegroundColor Cyan
        Write-Host ''

        $response = ''
        while ($response -notin @('Y', 'A')) {
            $response = (Read-Host '  Is Outlook closed? [Y] Yes, continue   [A] Abort').Trim().ToUpper()
        }

        if ($response -eq 'A') {
            Write-Host ''
            Write-Step 'Install aborted by operator. Re-run Install.ps1 when ready.' 'Yellow'
            return $false
        }

        # Y -- re-check; loop will verify and either proceed or prompt again
        $stillRunning = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
        if ($stillRunning) {
            Write-Host ''
            Write-Step 'Outlook is still open. Close it completely and try again.' 'Yellow'
            # Loop continues -- prompt shown again
        }
        # If not still running, loop condition will exit cleanly on next iteration
    }
}

function New-ArchivePSTViaCOM {
    <#
    .SYNOPSIS
        Creates a valid empty Archive PST file using Outlook COM.

    .DESCRIPTION
        Launches Outlook invisibly via COM, calls AddStore to a non-existent
        path so Outlook creates a fully structured valid PST, then immediately
        removes that store and quits COM cleanly.

        This produces a PST that Outlook will open reliably in Script 03 --
        no binary header guesswork, no partial structures. Outlook builds it
        correctly by definition.

        Outlook must be fully closed before calling this function.
        Use Invoke-WaitForOutlookClose before calling.

    .PARAMETER Path
        Full path where the Archive PST file should be created.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $outlook  = $null
    $ns       = $null
    $store    = $null

    try {
        Write-Step 'Launching Outlook to create Archive PST (this may take a moment)...' 'Cyan'

        # Launch Outlook COM invisibly
        $outlook = New-Object -ComObject Outlook.Application
        $ns      = $outlook.GetNamespace('MAPI')

        # AddStore creates a new PST at the given path if it does not exist.
        # Outlook builds a fully valid Unicode PST structure automatically.
        $ns.AddStore($Path)

        # Locate the newly added store so we can detach it cleanly
        $store = $null
        foreach ($s in $ns.Stores) {
            try {
                if ($s.FilePath -eq $Path) {
                    $store = $s
                    break
                }
            }
            catch { }
        }

        # Detach the store -- we only needed it created, not mounted
        if ($store) {
            $ns.RemoveStore($store.GetRootFolder())
        }

        return $true
    }
    catch {
        Write-Step "ERROR: Could not create Archive PST via COM -- $_" 'Red'
        return $false
    }
    finally {
        # Release COM objects cleanly regardless of success or failure
        if ($store)   { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($store)   | Out-Null }
        if ($ns)      { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns)      | Out-Null }
        if ($outlook) {
            $outlook.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Write-Summary {
    param([int]$Downloaded, [int]$Failed)

    $sep = '-' * 70
    Write-Host ''
    Write-Host $sep -ForegroundColor DarkCyan

    if ($Failed -gt 0) {
        Write-Host '  Install completed with errors' -ForegroundColor Yellow
        Write-Host "  $Downloaded file(s) downloaded successfully" -ForegroundColor Green
        Write-Host "  $Failed file(s) failed -- see errors above" -ForegroundColor Red
        Write-Host ''
        Write-Host '  Troubleshooting:' -ForegroundColor White
        Write-Host '  - Verify the GitHub repository is accessible' -ForegroundColor Gray
        Write-Host "  - Verify the repo exists at: $Script:GitHubUrl" -ForegroundColor Gray
        Write-Host '  - Re-run this installer to retry failed downloads' -ForegroundColor Gray    }
    else {
        Write-Host '  Install Complete -- All files downloaded successfully' -ForegroundColor Green
    }

    Write-Host $sep -ForegroundColor DarkCyan
    Write-Host ''
    Write-Host '  NEXT STEPS:' -ForegroundColor White
    Write-Host ''
    Write-Host '  1. Read QUICKSTART.md before running any script.' -ForegroundColor Cyan
    Write-Host "       $ProjectRoot\QUICKSTART.md" -ForegroundColor Gray
    Write-Host ''
    Write-Host '  2. Open PowerShell in your OMMigrate folder and run Script 00:' -ForegroundColor Cyan
    Write-Host '       cd "$env:USERPROFILE\Documents\OMMigrate"; Get-ChildItem -Recurse -Include *.ps1,*.psm1 | Unblock-File' -ForegroundColor Gray
    Write-Host '       .\Scripts\OMMigrate-00-Discover.ps1' -ForegroundColor Gray
    Write-Host ''
    Write-Host '  3. Runtime data (logs, reports, backups) is stored separately at:' -ForegroundColor Cyan
    Write-Host "       $Script:RuntimeBase\" -ForegroundColor Gray
    Write-Host ''
    Write-Host '  To update OMMigrate to the latest version from GitHub:' -ForegroundColor Cyan
    Write-Host '       cd "$env:USERPROFILE\Documents\OMMigrate"; Get-ChildItem -Recurse -Include *.ps1,*.psm1 | Unblock-File' -ForegroundColor Gray
    Write-Host '       .\Install.ps1 -Force' -ForegroundColor Gray
    Write-Host '       (re-downloads all files, runtime data is never touched)' -ForegroundColor Gray
    Write-Host ''
    Write-Host $sep -ForegroundColor DarkCyan
    Write-Host "  $Script:Product v$Script:Version" -ForegroundColor DarkGray
    Write-Host '  "Automating the Outlook migration Google suggested couldn''t be automated."' -ForegroundColor DarkGray
    Write-Host "  Architect: $Script:Architect" -ForegroundColor DarkGray
    Write-Host $sep -ForegroundColor DarkCyan
    Write-Host ''
}


# ============================================================
#  MAIN
# ============================================================

Write-Banner

# ── Execution policy ──────────────────────────────────────────────────────────
Write-Host '  [ Checking PowerShell execution policy ]' -ForegroundColor White
Write-Host ''
Set-ExecutionPolicyIfNeeded
Write-Host ''

# ── Create project folder ─────────────────────────────────────────────────────
Write-Host '  [ Creating project folder structure ]' -ForegroundColor White
Write-Host ''
New-Dir -Path $ProjectRoot -Label 'Documents\OMMigrate\'

foreach ($folder in $Script:ProjectFolders) {
    New-Dir -Path (Join-Path $ProjectRoot $folder) -Label "Documents\OMMigrate\$folder\"
}

Write-Host ''

# ── Create runtime folders ────────────────────────────────────────────────────
Write-Host '  [ Creating runtime data folders ]' -ForegroundColor White
Write-Host ''
New-Dir -Path $Script:RuntimeBase -Label 'Documents\OutlookMigration\'

foreach ($folder in $Script:RuntimeFolders) {
    New-Dir -Path (Join-Path $Script:RuntimeBase $folder) `
            -Label "Documents\OutlookMigration\$folder\"
}

Write-Host ''

# ── Create Archive PST ────────────────────────────────────────────────────────
Write-Host '  [ Creating local Archive PST ]' -ForegroundColor White
Write-Host ''

$archivePSTPath = Join-Path $Script:RuntimeBase 'Backups\OMMigrate_Archive.pst'

if (Test-Path $archivePSTPath) {
    Write-Step "Exists   : OMMigrate_Archive.pst  (already present -- skipping)" 'DarkGray'
}
else {
    # Outlook must be closed -- COM will launch it invisibly to create the PST.
    # Prompt operator to close Outlook if it is currently running.
    $outlookReady = Invoke-WaitForOutlookClose
    if (-not $outlookReady) {
        # Operator chose to abort
        Write-Host ''
        Write-Host '  Install did not complete. Re-run Install.ps1 when Outlook is closed.' -ForegroundColor Yellow
        Write-Host ''
        return
    }

    # Create the Archive PST via COM -- Outlook builds a fully valid PST structure
    $pstResult = New-ArchivePSTViaCOM -Path $archivePSTPath

    if ($pstResult -and (Test-Path $archivePSTPath)) {
        # Verify the file is a credible size -- a real PST is several KB minimum
        $pstSize = (Get-Item $archivePSTPath).Length
        if ($pstSize -gt 4096) {
            Write-Step "Created  : OMMigrate_Archive.pst  ($pstSize bytes -- verified)" 'Green'

            # Added 2026-07-10, Administrator (multi-archive support, TargetStoreName
            # hardcode fix -- documentation gap discovered during that work).
            # AddStore() creates the PST file but does NOT set any custom
            # display name -- confirmed via Microsoft documentation and
            # community reports that Outlook assigns a generic default name
            # (typically "Personal Folders") to a store added this way, and
            # separately confirmed that Store.DisplayName cannot be set
            # programmatically through the standard Outlook Object Model
            # (write attempts return access denied; only third-party
            # Extended MAPI tools like Redemption or OutlookSpy can do it,
            # none of which this project uses). This means the newly created
            # Archive PST will NOT automatically display as "OMMigrate Local
            # Archive" or any other specific name in Outlook's folder pane --
            # the operator must rename it manually, once, before it will be
            # clearly identifiable in the Script 00 TargetStoreName picker.
            Write-Host ''
            Write-Step 'ACTION REQUESTED: Rename the Archive PST in Outlook.' 'Yellow'
            Write-Host '           Outlook will show this new data file under a generic' -ForegroundColor Yellow
            Write-Host '           default name (often "Personal Folders") in the folder pane --' -ForegroundColor Yellow
            Write-Host '           it does not automatically pick up a name from this tool.' -ForegroundColor Yellow
            Write-Host '           Before running Script 00, open Outlook, right-click the new' -ForegroundColor Yellow
            Write-Host '           data file in the folder pane, choose Data File Properties,' -ForegroundColor Yellow
            Write-Host '           and give it a clear name (e.g. "OMMigrate Local Archive")' -ForegroundColor Yellow
            Write-Host '           so it is easy to identify in the archive-target picker.' -ForegroundColor Yellow
            Write-Host ''
        }
        else {
            # File exists but suspiciously small -- treat as failure
            Remove-Item $archivePSTPath -Force -ErrorAction SilentlyContinue
            Write-Host ''
            Write-Step 'FAILED   : Archive PST was created but failed size verification.' 'Red'
            Write-Step '           The file has been removed. Re-run Install.ps1 to try again.' 'Red'
            Write-Host ''
            Write-Host '  Install did not complete. Archive PST is required for Script 03.' -ForegroundColor Red
            Write-Host ''
            return
        }
    }
    else {
        # COM creation failed -- hard stop, Script 03 cannot run without this file
        Write-Host ''
        Write-Step 'FAILED   : Could not create OMMigrate_Archive.pst via Outlook COM.' 'Red'
        Write-Step '           Verify Outlook is installed and re-run Install.ps1.' 'Red'
        Write-Host ''
        Write-Host '  Install did not complete. Archive PST is required for Script 03.' -ForegroundColor Red
        Write-Host ''
        return
    }
}


# ── Create Template PST ───────────────────────────────────────────────────────
# A permanent empty PST used as a copy template by Script 01 -BackupIMAPOSTs.
# Never written to directly -- always copied to a new filename before use.
# Safe to re-run -- skipped if already present.
Write-Host ''
Write-Host '  [ Creating PST Template ]' -ForegroundColor White
Write-Host ''

$templatePSTPath = Join-Path $Script:RuntimeBase 'Backups\OMMigrate_Template.pst'

if (Test-Path $templatePSTPath) {
    Write-Step "Exists   : OMMigrate_Template.pst  (already present -- skipping)" 'DarkGray'
}
else {
    # Outlook must be closed for COM PST creation.
    # If it was already closed for the Archive PST above it will still be
    # closed here. If the Archive PST was skipped (already existed) we need
    # to check again.
    $outlookReady = Invoke-WaitForOutlookClose
    if (-not $outlookReady) {
        Write-Host ''
        Write-Host '  Install did not complete. Re-run Install.ps1 when Outlook is closed.' -ForegroundColor Yellow
        Write-Host ''
        return
    }

    $templateResult = New-ArchivePSTViaCOM -Path $templatePSTPath

    if ($templateResult -and (Test-Path $templatePSTPath)) {
        $templateSize = (Get-Item $templatePSTPath).Length
        if ($templateSize -gt 4096) {
            Write-Step "Created  : OMMigrate_Template.pst  ($templateSize bytes -- verified)" 'Green'
        }
        else {
            Remove-Item $templatePSTPath -Force -ErrorAction SilentlyContinue
            Write-Host ''
            Write-Step 'FAILED   : Template PST was created but failed size verification.' 'Red'
            Write-Step '           The file has been removed. Re-run Install.ps1 to try again.' 'Red'
            Write-Host ''
            Write-Host '  Install did not complete. Template PST is required for Script 01 -BackupIMAPOSTs.' -ForegroundColor Red
            Write-Host ''
            return
        }
    }
    else {
        Write-Host ''
        Write-Step 'FAILED   : Could not create OMMigrate_Template.pst via Outlook COM.' 'Red'
        Write-Step '           Verify Outlook is installed and re-run Install.ps1.' 'Red'
        Write-Host ''
        Write-Host '  Install did not complete. Template PST is required for Script 01 -BackupIMAPOSTs.' -ForegroundColor Red
        Write-Host ''
        return
    }
}


Write-Host '  [ Downloading files from GitHub ]' -ForegroundColor White
Write-Host "  Source: $Script:RawBase" -ForegroundColor DarkGray
Write-Host ''

$downloaded = 0
$failed     = 0

foreach ($file in $Script:FilesToDownload) {
    # Ensure subfolder exists before writing file
    $localDir = Split-Path (Join-Path $ProjectRoot $file.Local) -Parent
    if (-not (Test-Path $localDir)) {
        New-Item -ItemType Directory -Path $localDir -Force | Out-Null
    }

    $result = Get-GitHubFile -RepoPath $file.Repo -LocalPath $file.Local
    if ($result) { $downloaded++ } else { $failed++ }
}

# ── Unblock all downloaded files ──────────────────────────────────────────────
Invoke-UnblockAll

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Summary -Downloaded $downloaded -Failed $failed
