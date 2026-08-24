# OutlookMailMigrator (OMMigrate)

> *"Automating the Outlook migration Google suggested couldn't be automated."*

**Originator & Architect:** Kirk Shallcross - Shallcross Consulting
**Implementation Specialist:** Anthropic Claude AI
**Inception Date:** May 2026
**Version:** 1.5.2

---

## Architecture at a Glance

![OMMigrate pipeline architecture -- five sequential scripts, each gated on the previous script's completion manifest](docs_images/architecture_overview.svg)

Five scripts run in order, each writing a completion manifest that the next
script hard-checks before it will run. Script 00 reads your Outlook profile
and writes three control CSVs; Scripts 01-04 consume those CSVs and produce
an Archive PST, IMAP server-side folders, and an HTML report per run.
Outlook VBA macros provide a parallel, Outlook-native path to the same
rules engine Script 03 uses. See **The Five Scripts**, below, for the
step-by-step detail, or **OMMigrate_CommandLine_Reference.md** for every
parameter.

---

## Install

### Option A -- One-liner (recommended)

Open PowerShell and run these two commands:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
irm https://raw.githubusercontent.com/SC-Admin567/OMMigrate/main/Install.ps1 | iex
```

The execution policy line is a one-time step on new machines. If your policy
is already set to RemoteSigned or Unrestricted you can skip it.

The installer downloads everything, creates your project folder in Documents,
sets up runtime directories, and unblocks all files automatically.

### Option B -- Manual download

Use this if you can't run the one-liner above, or want to review
Install.ps1 before running it.

Download Install.ps1 from
https://github.com/SC-Admin567/OMMigrate/blob/main/Install.ps1, save it
into a new `Documents\OMMigrate\` folder, then from that folder:

```powershell
cd "$env:USERPROFILE\Documents\OMMigrate"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Unblock-File -Path .\Install.ps1
.\Install.ps1
```

The Unblock-File step is required because Windows marks files downloaded
from the internet as blocked. The one-liner in Option A bypasses this
automatically by piping directly into PowerShell memory.

---

Then read **QUICKSTART.md** and run Script 00.

For full parameter details and settings file keys, see
**OMMigrate_CommandLine_Reference.md** and **OMMigrate_Settings_Reference.md**.

---

## What Is OMMigrate?

OMMigrate is a PowerShell toolkit that automates migration of Microsoft
Outlook 2016/2019/2021 desktop email accounts from POP3 to IMAP protocol.

It handles the complete migration lifecycle across multiple accounts:

- **Discovery** -- reads all account settings directly from the Windows Registry
- **Backup** -- exports and verifies a PST backup before anything is touched
- **Conversion** -- removes each POP3 account and adds it as IMAP
- **Restore** -- migrates folder structure, email content, and Outlook Rules
- **Artifacts** -- migrates Calendar, Contacts, Tasks, Notes, and Journal

Every step is logged, reported, and recoverable. Nothing is permanent until
you confirm it.

---

## Why Does This Exist?

Switching an Outlook account from POP3 to IMAP requires removing the old
connection and creating a new one. For a single account this is tedious.
For ten or twenty accounts it is a project.

No automated tool existed to do this for local Outlook desktop profiles.
Every guide online described the same manual process repeated per account.
OMMigrate automates those steps safely, with full backup and rollback support,
across any number of accounts in a single controlled session.

---

## Who Is It For?

| User | Use Case |
|---|---|
| Home power users | Multiple accumulated accounts needing modernization |
| Small business owners | Moving off legacy POP3 hosting to modern IMAP |
| IT consultants / MSPs | Client desktop migrations -- turns hours into minutes |
| Email hosting migrations | Company changes provider, all staff need conversion |
| Compliance requirements | Moving to server-retained IMAP for legal hold |

> ### ⚠️ Read This Before You Decide to Use OMMigrate: It Changes How Your Rules Organize Mail
>
> OMMigrate does not just move messages -- **it restructures your Outlook
> Rules to work with a new mail-organization model**, and this is not
> optional or reversible mid-run. Understand this decision before you
> commit, especially if you're evaluating this tool for a client or an
> organization.
>
> **What changes:** your existing POP3 rules file mail into folders under
> your POP3 account's own local store. After migration, each rule's
> folder target is remapped to point at either your new IMAP account
> (server-side, syncs across devices) or a local Archive PST (desktop-only,
> stays on this machine, never touches the server). Which destination each
> rule uses is a decision OMMigrate makes for you during discovery, based
> on your folder-mapping choices -- it is a genuinely different filing
> architecture than what your POP3 rules were doing, not a like-for-like
> copy of your existing setup.
>
> **Why this matters for consultants and admins specifically:** a client
> or end user who expects their mail organization to look and behave
> exactly as it did before migration may not want this. If server-retained
> visibility across every folder (not just the Inbox) is a hard
> requirement -- for compliance, legal hold, or multi-device parity -- review
> **"Archive PST -- Local Folder Storage"** and **"rules_inventory.csv"**
> (below and in QUICKSTART.md) *before* running Script 00, so you can
> make an informed decision, or adjust the folder-mapping choices
> accordingly, rather than discovering the new structure only after a
> full pipeline run.
>
> **This is a one-way decision once rules deploy.** Script 03 recreates
> and consolidates rules against the new structure; there is no
> "undo migration" command. Your only path back to the exact original
> rules is the Outlook rules export you made *before* running any script
> -- see **"Before You Begin: Back Up Your Existing Data,"** below.

---

## Requirements

| Requirement | Details |
|---|---|
| **Operating System** | Windows 10 or Windows 11 (64-bit) |
| **PowerShell** | Version 5.1 or higher (built into Windows 10/11) |
| **Outlook** | Classic Outlook 2016, 2019, or 2021 desktop client |
| **Run as** | The Windows user who owns the Outlook profile -- a normal (non-elevated) PowerShell window, not "Run as Administrator" |

> **Not supported:** New Outlook (Microsoft Store version), Outlook 2013
> or earlier, or macOS Outlook.

> **Why not Administrator?** An Outlook profile belongs to the regular
> user account, not the Administrator account -- even if you're logged in
> as an administrator, opening PowerShell via "Run as Administrator"
> starts a session that can't see your normal Outlook profile at all. If
> nothing turns up, this is almost always why. Nothing in OMMigrate needs
> elevated permissions: `Install.ps1` only writes to your own Documents
> folder, and sets the execution policy for your user account
> only (`-Scope CurrentUser`), not machine-wide. Script 00 actively checks
> for this and logs a WARNING if it detects an elevated session -- Scripts
> 01-04 do not perform this check, so if you see Script 00 succeed but a
> later script behaves unexpectedly, double-check you're still in the
> same, non-elevated PowerShell window.

---

## Before You Begin: Back Up Your Existing Data (Safety Net)

**Do this before running any OMMigrate script, including Script 00.**

OMMigrate's own Script 01 backs up your email *content* before Script 02
makes any account changes -- but Outlook **rules** are separate from
content, and no script in the pipeline exports your existing rules for
you automatically. If a rules-side issue comes up mid-migration, Outlook's
own native rule export/import is the only way back to your original
rules -- and that only works if the export already exists.

**Prerequisites (before running any scripts):**

1. **Back up your existing PST files.** Outlook -> File -> Account Settings
   -> Account Settings (yes, click "Account Settings" twice -- the first is
   a menu, the second is the actual dialog) -> Data Files tab lists every
   PST's exact file path -- this is the authoritative source, since the
   actual location varies by Outlook version and setup history. If you're browsing manually instead,
   check `Documents\Outlook Files\` (the default on newer Outlook
   installs) and `AppData\Local\Microsoft\Outlook\` (older installs, or
   accounts originally set up in an older Outlook version) -- a profile
   can have PSTs in either location, or both, and `AppData` is hidden by
   default (enable "Hidden items" in File Explorer's View tab to see it).
   If neither default location has them, an admin may have originally
   pointed Outlook at a custom folder (e.g. a shared drive, a different
   local path) when the PST was first created -- the Data Files tab will
   still show the correct path even in that case.
   Copy your PST files somewhere *outside* `Documents\OMMigrate` and
   `Documents\OutlookMigration`.
2. **Export your existing Outlook rules, for every email account.**
   Outlook -> Rules -> Manage Rules & Alerts -> Email Rules tab -> Options
   -> Export Rules. Rules export per account, so repeat this once per
   account/store. This is your fallback to restore your original,
   pre-migration rules via **Import Rules**, independent of OMMigrate.

**Postrequisite (after migration is complete):**

1. Export rules again the same way (Options -> Export Rules, every
   account) and save it alongside the pre-migration export. This "after"
   snapshot is what you'd use to recover your new IMAP rules via
   **Import Rules**, if that's ever needed.

Treat this second export as a **point-in-time recovery option**, not a
one-time step tied only to the migration itself. Your rules will keep
changing after migration -- you'll likely hand-edit `SendersDomain` and
`TargetFolderPath` values in `rules_inventory.csv` as you correct the
auto-generated mappings (see "rules_inventory.csv," below and in
QUICKSTART.md), and Script 03 reruns will redeploy those changes to the
live rules. A fresh Export Rules snapshot, taken periodically as your
rules stabilize -- not just once, right after the first migration -- gives
you a dated restore point independent of whether anything has actually
gone wrong yet.

**Keep both sets of backups outside any OMMigrate-managed folder** -- a
separate Recovery location (an external drive, cloud storage, or just a
folder elsewhere on disk) that a script run or reinstall can't touch. See
**QUICKSTART.md, "Back Up Your Existing Outlook Data First,"** for the
full walkthrough with more detail on each step.

> **Disclaimer:** OMMigrate is provided as-is, with no warranty of any
> kind (see **License**, below). The backup and rules-export steps above
> are your safety net -- OMMigrate's own PST backups only cover message
> content, not rules, and the tool cannot recover data or rules that were
> never independently backed up before a script ran. Running OMMigrate
> without completing the backup steps above is done at your own risk. The
> author and contributors are not responsible for lost, corrupted, or
> unrecoverable email, rules, or other data resulting from use of this
> tool, including cases where these backup steps were skipped or
> incomplete.

---

## Features

- **Automated account discovery** -- reads server names, ports, SSL settings,
  and data file paths directly from the Windows Registry. No manual data entry
- **Automated account selection** -- WinForms picker after Script 00 lets you
  check which POP3 accounts to migrate now versus defer. No manual CSV editing
- **Automated CSV updates** -- Script 02 writes post-conversion status to the
  CSV automatically. Script 03 marks accounts COMPLETE automatically
- **Full PST backup** before any account is touched, with size verification
- **Hybrid folder architecture** -- choose which folders live on the IMAP
  server (all devices and webmail) and which stay in a local Archive PST
- **Multi-archive support** -- attach more than one archive PST and map
  individual accounts to whichever archive they belong in, via a picker
  in Script 00
- **Rules preservation** -- Outlook Rules are inventoried and consolidated
  onto the new IMAP folder structure automatically, including manually-added
  UI conditions (Subject, Category, Importance, and 9 others)
- **Calendar, Contacts, and other artifacts** -- Script 04 migrates
  non-mail Outlook items (Calendar, Contacts, Tasks, Notes, Journal) with
  duplicate detection
- **Multi-provider support** -- standard IMAP, AT&T/Yahoo legacy domains
  (ameritech.net, sbcglobal.net), AWS SES outbound SMTP, Gmail, iCloud,
  and Microsoft Exchange accounts handled correctly
- **Safe exit at any time** -- type `EXIT` at any prompt or press `Ctrl+C`
  for a clean logged exit with recovery instructions written automatically
- **Resume from interruption** -- checkpoint files track progress so
  re-running any script skips accounts already completed
- **Professional HTML reports** -- browser-rendered reports for every step,
  suitable as client deliverables
- **Full audit trail** -- cumulative append-only audit log across all runs
- **Preview mode** -- simulate the entire process without changing anything

---

## Account Type Support

| Account Type | Action |
|---|---|
| POP3 -- standard mail server | Full automated migration to IMAP |
| POP3 -- AT&T/Yahoo legacy domain | Migration with Secure Mail Key (pre-generate required) |
| POP3 -- AWS SES outbound SMTP | Migration -- AWS SES credentials re-entered at prompt |
| POP3 -- Gmail | Migration via Outlook automatic setup |
| IMAP already | Skipped for conversion -- folder assessment only |
| Microsoft Exchange (live.com, outlook.com) | Skipped for conversion -- already optimal protocol. Rules are still tracked and folder-target-updated by Script 03 (see Special Account Notes below) |

---

## The Five Scripts

```
Script 00 -- Discover    Read-only scan. Safe to run any time. No changes made.
Script 01 -- Backup      Exports PST backups. Non-destructive. Safe to re-run.
Script 02 -- Convert     Guides operator through POP3 removal and IMAP add. Requires Y/N per account.
Script 03 -- Restore     Migrates folders and consolidates Outlook Rules.
Script 04 -- Artifacts   Migrates Calendar, Contacts, Tasks, Notes, and Journal.
```

Each script writes a completion manifest on success and the next script
reads it before proceeding -- this is a hard gate. If a prerequisite step's
manifest is missing, corrupt, or shows a FAILED status, the next script
stops with a clear error rather than proceeding on incomplete data.

> **Filenames include your Outlook profile** (e.g. `migration_accounts_Outlook.csv`
> instead of plain `migration_accounts.csv`) -- this document uses base names
> throughout for readability. See Directory Structure, below, for details.

**Script 00 account picker:**
After discovery completes, a selection window opens automatically showing
all discovered POP3 accounts. Check each account you want to migrate now
and click OK. Unchecked accounts are set to SKIP. Re-run Script 00 at any
time to change your selection before running Script 01.

**Script 00 TargetStoreName picker (multi-archive):**
If more than one archive PST is attached, a second picker opens after the
account picker -- one panel per attached archive, with a checkbox list of
accounts per panel. Map each account to the archive PST its Local-destination
folders should live in. Unmapped accounts fall back to the default Archive
PST. Re-run Script 00 at any time to change the mapping.

**Script 02 auto-update:**
After each account is successfully converted, Script 02 automatically
updates `migration_accounts.csv` -- no manual CSV edit required:
`MigrationAction=FOLDER-ONLY`, `ProviderTag=IMAP-CONVERTED`, `AccountType=IMAP`.

**Script 03 account picker:**
When Script 03 runs it displays a selection window showing all eligible
accounts (`MigrationAction=FOLDER-ONLY`). Check the accounts you want to
process this run and click OK. Use Select All to process all at once.
Already completed accounts (`COMPLETE`) are excluded automatically.
Cancel exits safely without making changes. Your selection is authoritative --
only the accounts you check are touched by folder migration, rule deployment,
or rule resorting this run.

**Run from your Documents\OMMigrate\ folder:**

```powershell
cd "$env:USERPROFILE\Documents\OMMigrate"; Get-ChildItem -Recurse -Include *.ps1,*.psm1 | Unblock-File
.\Scripts\OMMigrate-00-Discover.ps1
.\Scripts\OMMigrate-01-Backup.ps1
.\Scripts\OMMigrate-02-Convert.ps1
.\Scripts\OMMigrate-03-Restore.ps1
.\Scripts\OMMigrate-04-Artifacts.ps1
```

---

## Optional Utilities

Three standalone scripts live in `Utilities\`, separate from the five
pipeline scripts above. None of them run automatically as part of a normal
migration -- run them only if their specific situation comes up.

**Why these exist:** these three utilities were written early in
development, when initial POP3 accounts carried a large number of
duplicate rules (one real environment went from 527 rules down to 427
after cleanup -- about 100 duplicates). Script 00's rule inventory
(`Get-OutlookRules`) later gained its own internal deduplication using
this same scoring logic, so `rules_inventory.csv` today only ever shows
one row per rule even without running these utilities first -- but that
dedup happens in the CSV, not in Outlook itself, so the duplicate rules
still exist in live Outlook until something actually removes them. These
utilities remain useful as a preventive first pass on an initial pipeline
run, especially against an account with a long POP3/rules history, or any
time you suspect duplicates after a `.rwz` import or manual Rules and
Alerts editing. Run `OMMigrate-FindDuplicateRules.ps1` to check, then
`OMMigrate-RemoveDuplicateRules.ps1` to resolve what it finds.

Each utility does exactly what its name says:

**OMMigrate-CleanRulesCSV.ps1** -- Rules CSV cleanup.
If `rules_inventory.csv` ends up with duplicate rows for the same rule
(for example, after some manual CSV editing, or a partial/interrupted
Script 00 run), this scores every row in each duplicate group and keeps
only the best one -- preferring rows with a valid `TargetFolderPath`, real
`Actions`, `NeedsFolderUpdate=True`, then `IsEnabled=True`, in that order.
A timestamped backup of the CSV is written before anything changes.

**OMMigrate-FindDuplicateRules.ps1** -- Scan for duplicate Outlook Rules.
A read-only diagnostic that connects to Outlook and reports any rule whose
name matches another rule in the same store, including the `(2)`, `(3)`
style suffix copies that a `.rwz` rule-file import can create. Useful any
time rules look duplicated in Rules and Alerts and you want a report
before deciding whether to remove anything. Makes no changes to Outlook.
When duplicates are found, saves `Reports\DuplicateRules_<Profile>.csv`
(the file `OMMigrate-RemoveDuplicateRules.ps1` reads) plus a separate
timestamped copy for your own history -- if no duplicates are found, no
file is written.

**OMMigrate-RemoveDuplicateRules.ps1** -- Remove duplicate Outlook Rules.
Reads the CSV that `OMMigrate-FindDuplicateRules.ps1` produced and deletes
each listed rule from Outlook via COM, keeping the best copy per group
(the one with a valid `MoveToFolder` target). Run `OMMigrate-FindDuplicateRules.ps1`
first and review its report before running this. Supports `-WhatIf` to
preview what would be deleted.

All three prompt for `Y`/`N` (or `EXIT`) before doing anything, and accept
`-WhatIf` for a preview run where it applies. None of the three read or
write `OMMigrate_Settings.json` -- they resolve which Outlook profile to
act on independently, either from the `-ProfileName` parameter or an
interactive prompt (auto-selected automatically if only one Outlook
profile exists on the machine).

```powershell
cd "$env:USERPROFILE\Documents\OMMigrate"; Get-ChildItem -Recurse -Include *.ps1,*.psm1 | Unblock-File
.\Utilities\OMMigrate-CleanRulesCSV.ps1 -WhatIf
.\Utilities\OMMigrate-FindDuplicateRules.ps1
.\Utilities\OMMigrate-RemoveDuplicateRules.ps1 -WhatIf
```

Add `-ProfileName "YourProfileName"` to any of the three to skip the
interactive profile prompt.

---

## Directory Structure

```
Documents\OMMigrate\                      <- Project folder (scripts and modules)
|
+-- Install.ps1                         One-command installer
+-- Update-Version.ps1                  Version number updater (all files at once)
+-- version.txt                         Single source of truth for version number
+-- show_tree.ps1                       Regenerates tree_view.txt from the current folder structure
+-- tree_view.txt                       Snapshot of the full project tree (this listing's source)
+-- README.md                           This file
+-- QUICKSTART.md                       Step-by-step first run guide
+-- OMMigrate_CommandLine_Reference.md  Every script parameter, reference card
+-- OMMigrate_Settings_Reference.md     Every settings file key, reference card
+-- CHANGELOG.md                        Version history
+-- LICENSE.md                          License terms
+-- CONTRIBUTING.md                     Bug report guidelines and contribution policy
+-- .gitattributes                      Git line-ending/encoding rules
+-- .gitignore                          Runtime data, test data, and OS/editor noise excluded from the repo
|
+-- .github\                           GitHub repo configuration
|   +-- ISSUE_TEMPLATE\
|       +-- bug_report.yml              Structured bug report form
|       +-- feature_request.yml         Structured feature request form
|       +-- config.yml                  Issue template settings (restricts blank issues to maintainers)
|
+-- docs_images\                       Images referenced by this README
|   +-- architecture_overview.svg       Pipeline architecture diagram
|
+-- Scripts\                            Executable migration scripts
|   +-- OMMigrate-00-Discover.ps1       Step 00 -- Discovery (read-only)
|   +-- OMMigrate-01-Backup.ps1         Step 01 -- PST backup and verification
|   +-- OMMigrate-02-Convert.ps1        Step 02 -- POP3 removal and IMAP add
|   +-- OMMigrate-03-Restore.ps1        Step 03 -- Folder migration and rules
|   +-- OMMigrate-04-Artifacts.ps1      Step 04 -- Calendar/Contacts/Tasks/Notes/Journal
|
+-- Modules\                            Shared PowerShell modules (auto-loaded)
|   +-- OMMigrate-Core.psm1             Logging, audit, settings, exit handling
|   +-- OMMigrate-Registry.psm1         Registry reader and account classifier
|   +-- OMMigrate-Outlook.psm1          Outlook COM API layer
|   +-- OMMigrate-Reporting.psm1        HTML report generator
|
+-- Utilities\                          Standalone maintenance scripts
|   +-- OMMigrate-CleanRulesCSV.ps1     Rules CSV cleanup
|   +-- OMMigrate-FindDuplicateRules.ps1    Scan for duplicate Outlook Rules
|   +-- OMMigrate-RemoveDuplicateRules.ps1  Remove duplicate Outlook Rules
|
+-- OutlookVBAMacros\                   Outlook VBA macros (installed into Outlook manually --
|   |                                   see QUICKSTART.md Step 0.5)
|   +-- Module1.bas                     Duplicate detection/removal
|   +-- Module2.bas                     Support functions
|   +-- Module3.bas                     Consolidated rule engine (VBA port)
|   +-- Module7.bas                     Archive misroute correction
|   +-- FrmArchivePicker.frm/.frx       Archive picker form (Module7)
|   +-- FrmProgress.frm/.frx            Progress form
|   +-- ThisOutlookSession.cls          Outlook session event handlers
|
+-- QUICKSTART_images\                 Screenshots referenced by QUICKSTART.md
|
+-- Obsidian-Addon-Install.bat          Optional -- installs free Obsidian .md viewer (see QUICKSTART.md Step 0.75)
+-- Obsidian-Addon-Uninstall.bat        Optional -- uninstalls the Obsidian .md viewer add-on
+-- Obsidian-Addon-DeployLink.reg       Registry file -- associates .md files with Obsidian (used by Install.bat)
+-- Obsidian-Addon-UninstallLink.reg    Registry file -- reverts .md file association (used by Uninstall.bat)
+-- Obsidian-Addon-Opener.ps1           Helper script -- opens the clicked .md file directly in its Obsidian vault
+-- Obsidian-Addon-SilentRunner.vbs     Helper script -- launches Opener.ps1 without a visible console window
+-- stripe-button.html                  Donate button embed, rendered by the HTML Viewer Plus plugin below
|
+-- .obsidian\                         Obsidian vault config -- renders README.md's donate-button embed when opened in Obsidian
    +-- app.json                        Not downloaded by Install.ps1 -- created locally by Obsidian
    +-- appearance.json                 the first time it opens this vault (per-user app preferences,
    +-- workspace.json                  excluded from the repo by .gitignore; see .gitignore's own
    |                                   comment for why)
    +-- community-plugins.json          Downloaded by Install.ps1 -- defines which plugins are enabled
    +-- core-plugins.json               (not a per-user preference, tracked in the repo)
    +-- plugins\
        +-- html-viewer-plus\          HTML Viewer Plus plugin -- renders the stripe-button.html embed
            +-- main.js
            +-- manifest.json
            +-- styles.css
            +-- data.json
            +-- debug.log               Empty by default -- enable debug mode in the plugin to populate


Documents\OutlookMigration\             <- Runtime data (created automatically)
|
+-- RECOVERY.txt                        Plain-English recovery instructions -- rewritten on every script exit
|
+-- Config\                             Control files -- you edit these between runs
|   +-- migration_accounts_<Profile>.csv    Account list -- add passwords here
|   +-- folder_map_<Profile>.csv            Folder destinations -- Server/Local/Skip
|   +-- rules_inventory_<Profile>.csv       Outlook Rules inventory
|   +-- OMMigrate_Settings_<Profile>.json   Per-profile settings (auto-created, editable)
|
+-- Logs\                               One log file per run plus cumulative audit
+-- Reports\                            HTML reports -- open in any browser
+-- Backups\                            PST backup files -- one per POP3 account
+-- Manifests\                          Step completion and resume checkpoint files
```

> **Why `<Profile>` in the filename?** All three control CSVs and the
> settings file are named after your active Outlook profile (e.g.
> `migration_accounts_Outlook.csv`) so that runs under different profiles
> never overwrite each other's data. If you only use one Outlook profile
> you will only ever see one version of each file. The rest of this guide
> refers to these files by their base name (`migration_accounts.csv`, etc.)
> for readability -- look for the profile-suffixed version on disk.

> **About `rules_inventory.csv`:** this file tracks how each Outlook rule
> maps to its new, post-migration folder target -- your old POP3 rules
> pointed at folders under your POP3 account's own store, and after
> migration mail lives under a different store, so every rule's target
> has to be remapped. Script 00 generates and maintains it automatically,
> but **expect to hand-edit it on a first run**: the `TargetFolderPath`
> Script 00 seeds for each row is often just a best-guess folder name,
> not a verified target, and `SendersDomain` -- the actual condition
> value Script 03 deploys -- frequently needs correcting or filling in
> before you trust the CSV. See **QUICKSTART.md,
> "rules_inventory.csv -- how OMMigrate maps rules to their new folders,"**
> for the full column-by-column reference, including how `SendersDomain`
> and `TargetFolderPath` relate to each other and what values each
> column expects.

---

## Script Parameters

> A fresh PowerShell window starts at your home folder, not the OMMigrate
> project folder, and files can be blocked after a fresh download or
> manual replacement. Run this once per session before any example below:
> ```powershell
> cd "$env:USERPROFILE\Documents\OMMigrate"; Get-ChildItem -Recurse -Include *.ps1,*.psm1 | Unblock-File
> ```
> See QUICKSTART.md, "Always Verify Your Working Directory First," for
> more detail.
>
> **Why only `*.ps1,*.psm1` here?** This command is scoped to exactly what
> Scripts 00-04 load: the Scripts themselves (`.ps1`) and the Modules they
> import (`.psm1`). It intentionally does not cover the Obsidian Add-on's
> `.bat`/`.reg`/`.vbs` files -- those are unrelated to running a migration
> and are unblocked separately (automatically, by Install.ps1 itself) only
> if you choose to use that optional add-on. See QUICKSTART.md, Step 0.75.

All scripts accept these common parameters:

```powershell
# Simulate everything without making any changes
.\Scripts\OMMigrate-00-Discover.ps1 -Preview

# Skip Y/N prompts (use with caution on Script 02)
.\Scripts\OMMigrate-01-Backup.ps1 -Force

# More verbose logging for troubleshooting
.\Scripts\OMMigrate-00-Discover.ps1 -LogLevel DEBUG

# Use a custom runtime data location
.\Scripts\OMMigrate-00-Discover.ps1 -BasePath "D:\Migration"
```

Script 00 additional parameters:

```powershell
# Target a specific Outlook profile by name
.\Scripts\OMMigrate-00-Discover.ps1 -ProfileName "Outlook"

# Patch an existing rules_inventory.csv in place -- skips full COM enumeration
.\Scripts\OMMigrate-00-Discover.ps1 -PatchRulesCSV
```

Script 01 additional parameters:

```powershell
# Skip the post-backup PST verification pass
.\Scripts\OMMigrate-01-Backup.ps1 -SkipVerification

# Refresh folder_map.csv without re-running the full backup
.\Scripts\OMMigrate-01-Backup.ps1 -RefreshFolderMap
```

Script 02 additional parameters:

```powershell
# Skip the pre-flight backup-completion check
.\Scripts\OMMigrate-02-Convert.ps1 -SkipBackupCheck
```

Script 03 additional parameters:

```powershell
# Re-run rules deployment/resort only -- skip folder/item migration entirely
.\Scripts\OMMigrate-03-Restore.ps1 -RefreshRulesOnly

# Batch rule Save() calls during a full rebuild (advanced -- see CHANGELOG 1.5.0)
.\Scripts\OMMigrate-03-Restore.ps1 -SaveBatchSize 10

# Use a custom name for the default local Archive PST
.\Scripts\OMMigrate-03-Restore.ps1 -ArchivePSTName "MyArchive_2026.pst"

# Recovery-only: recreate rules from a backup PST onto the live IMAP store
.\Scripts\OMMigrate-03-Restore.ps1 -RecreateRules
```

---

## Runtime Output Files

CSV and settings filenames below include your Outlook profile name (e.g.
`migration_accounts_Outlook.csv`) -- shown here by base name for readability.

| File | Created By | Purpose |
|---|---|---|
| RECOVERY.txt | Every run (graceful or emergency exit) | Plain-English recovery instructions -- rewritten each time, always reflects the most recent exit |
| Config\migration_accounts_&lt;Profile&gt;.csv | Script 00 | Account list -- add passwords before Script 01 |
| Config\folder_map_&lt;Profile&gt;.csv | Script 00 | Folder destinations -- review before Script 03 |
| Config\rules_inventory_&lt;Profile&gt;.csv | Script 00 | Outlook Rules inventory -- drives Script 03 rule deployment |
| Config\OMMigrate_Settings_&lt;Profile&gt;.json | Script 00 (first run per profile) | Per-profile settings, including archive mappings -- edit to customize |
| Backups\\*.pst | Script 01 | PST backup per POP3 account |
| Reports\Discovery_Report_YYYYMMDD_HHMMSS.html | Script 00 | Pre-migration account inventory |
| Reports\Backup_Report_YYYYMMDD_HHMMSS.html | Script 01 | Backup verification results |
| Reports\Migration_Report_YYYYMMDD_HHMMSS.html | Script 02 | Conversion results |
| Reports\Migration_Report_YYYYMMDD_HHMMSS.html | Script 03 | Folder migration and rules results (same report name as Script 02 -- distinguished by timestamp and subtitle) |
| Reports\Artifacts_Report_YYYYMMDD_HHMMSS.html | Script 04 | Calendar/Contacts/Tasks/Notes/Journal migration results |
| Logs\OMMigrate_YYYYMMDD_HHMMSS.log | Every run | Timestamped run log |
| Logs\OMMigrate_Audit.log | Every run | Cumulative audit trail |
| Manifests\Step0N_Complete.json | Each script (00-04) | Hard gate -- the next script reads and validates this before running |
| Manifests\Step02_Complete.json | Script 02 | Auto-clears itself once all recorded accounts are confirmed converted -- manual delete only needed as a fallback |
| Manifests\Step0N_Checkpoint.json | Scripts 02/03/04 | Resume point after interruption -- auto-deleted after successful session |
| Manifests\Step02_SendReceiveState.json | Script 02 | Send/Receive group state -- auto-deleted after restore |

> **How the `SendersDomain` column stays in sync (added 1.5.1):** every
> Script 00 run -- first run or rerun -- reads each Outlook rule's own
> SenderAddress condition (the address list you see in the Rules Manager)
> and writes it into `SendersDomain`, overwriting whatever was there
> before, including a manual CSV edit. If a rule has no SenderAddress
> condition (e.g. a fresh `.rwz` import, or a manually added rule where
> that condition was never set), `SendersDomain` falls back to the best
> guess it always used: the last folder name segment of `TargetFolderPath`.
>
> **The live Outlook rule is the source of truth.** If you edit
> `SendersDomain` directly in the CSV -- for example, adding another
> space-separated address -- that edit only reaches Outlook once you run
> Script 03, which pushes the CSV's current value into the rule's
> SenderAddress condition. Timing matters here: if you edit the CSV and
> then run Script 00 again *before* running Script 03, your edit is lost
> -- Script 00 has no way to tell a "not yet pushed" CSV edit apart from
> a stale value, so it always resyncs from whatever the live rule still
> says. The safe order is: edit the CSV, run Script 03, then any later
> Script 00 rerun will read that same value back from the now-updated
> rule and simply confirm it.

---

## Troubleshooting

**Script 00 finds no accounts, or very few compared to what you expect**
Check whether you're in an elevated ("Run as Administrator") PowerShell
window -- look at the title bar for "Administrator:". An Outlook profile
belongs to your regular user account; an elevated session runs in a
different context and can't see it. Close the window and open a normal
PowerShell window instead (Start -> type PowerShell -> Enter, not
"Run as administrator"). Script 00 logs a WARNING if it detects this;
check the run log if you're not sure. See Requirements, above, for why
elevation is never needed for OMMigrate.

**"Cannot be loaded -- file is not digitally signed"**
Two possible causes:

If this happens when running Install.ps1 after a manual download:
```powershell
Unblock-File -Path .\Install.ps1
.\Install.ps1
```

If this happens when running any other script after install:
```powershell
Get-ChildItem -Recurse -Filter *.ps1 | Unblock-File
Get-ChildItem -Recurse -Filter *.psm1 | Unblock-File
```

Or re-run the installer -- it unblocks everything automatically:
```powershell
.\Install.ps1
```

**Script fails with parse errors after updating a file**
A text editor corrupted special characters during copy/paste. Re-run the
installer -- it will re-download the affected file cleanly:
```powershell
irm https://raw.githubusercontent.com/SC-Admin567/OMMigrate/main/Install.ps1 | iex
```

**Script fails with CSV read error**
Close Excel before running any script. Excel locks `migration_accounts.csv`
exclusively -- any script that reads the CSV will fail if the file is open.
Save your changes and close Excel, then re-run the script.

**Script 00 shows wrong or missing server names**
For well-known providers (AT&T, Gmail, Yahoo, iCloud) server settings are
auto-populated from a built-in lookup table. If a server name is still
missing, fill it in manually in Config\migration_accounts.csv before Script 01.

**Script 02 -- POP3 removed but IMAP add failed**
Your email data is safe in the backup PST. Script 02 displays the exact
backup path and recovery steps. Add the account manually in Outlook via
File -> Add Account, then re-run Script 02 -- it will detect the account
is now IMAP and skip it automatically.

**Script 02 FATAL ERROR -- cannot resume**
`Manifests\Step02_Checkpoint.json` is auto-deleted after a successful run.
If Script 02 crashes with a FATAL ERROR and leaves a stale checkpoint, delete
`Manifests\Step02_Checkpoint.json` and `Manifests\Step02_SendReceiveState.json`
before re-running. These files from the failed run can cause incorrect behavior
on restart.

**An account is missing after Script 02**
Open your backup PST in Outlook: File -> Open & Export -> Open Outlook
Data File. Navigate to Documents\OutlookMigration\Backups\ and open the
PST for that account. All email is there.

**Send/Receive errors after Script 02**
If Outlook shows Send/Receive errors after Script 02, the Send/Receive groups
may still be suspended from an interrupted run. Fix manually:
1. In Outlook -- File > Options > Advanced > Send/Receive
2. Verify all groups are enabled
3. Press F9 to trigger a manual Send/Receive

**Other accounts show connection errors after Script 02**
The migration process can leave stale TCP connections on the mail server.
Reboot your mail server and wait 2-3 minutes before retrying.

**Outlook not closing between Script 02 phases**
Script 02 requires Outlook to be fully closed before it launches. Always
verify in Task Manager that no OUTLOOK.EXE process is running before
starting Script 02.

**Script fails with rules_inventory.csv read/write error**
Close Excel before running Script 00 or Script 03. Excel locks
`rules_inventory.csv` exclusively, the same as `migration_accounts.csv`.
Save your changes and close Excel, then re-run the script.

**Script 03 rules deployment is taking a very long time**
On accounts with a large existing rule set (several hundred rules), a full
rule rebuild can take significantly longer than a normal run. This is
expected -- Script 03 logs progress per rule/store as it goes, so a long
run is not necessarily a stuck one. If you need to interrupt it, `Ctrl+C`
performs a clean exit; re-running Script 03 will pick up only the rules
that still need work.

**After many repeated Script 03 reruns in one sitting (e.g. iterating on
a rule while testing), a run logs success but the change doesn't show up
in Outlook**
This is a known Outlook COM anomaly, not an OMMigrate defect -- Outlook's
underlying rules-store COM interface can accumulate memory/resource
pressure after many connect-and-disconnect cycles within a single long
Windows session, most likely to surface if you're deliberately rerunning
Script 03 several times back-to-back against the same rule (for example,
while testing or refining a rule's setup before your real migration). A
one-time production migration run is very unlikely to encounter this. If
you do hit it, a full reboot clears it completely -- close Outlook,
reboot, and re-run the script. If the same symptom reproduces immediately
after a clean reboot rather than after heavy repeated testing, that's
worth reporting (see CONTRIBUTING.md).

---

## Archive PST -- Local Folder Storage

Script 03 creates a local Archive PST (`OMMigrate_Archive.pst`) as the
permanent store for all Local-destination folders across all migrated accounts.

**The Archive PST appears in Outlook automatically** after Script 03 runs --
no manual attach step required. It will be visible in the Outlook folder pane
as **"OMMigrate Local Archive"** with a subfolder for each migrated account.

**Multiple archive PSTs:** if you attach more than one archive PST before
running Script 00, the TargetStoreName picker (see The Five Scripts, above)
lets you map individual accounts to whichever archive they belong in. An
account with no explicit mapping falls back to the default Archive PST
described above.

**Folder destinations** (set in `Config\folder_map.csv` before Script 03):

| Destination | Where it lives | Visible on |
|---|---|---|
| `Server` | IMAP server | All devices, webmail, phone |
| `Local` | OMMigrate Local Archive PST | This desktop only |
| `Skip` | Not migrated | -- |

Standard IMAP folders (Inbox, Sent, Drafts, etc.) default to Server.
All other folders default to Local. You can change any destination before
running Script 03 -- or accept the defaults and adjust afterward.

**Moving a folder after migration:**
Local folders can be moved to the IMAP server at any time by dragging and
dropping them from the Archive PST into the IMAP account in Outlook. No
scripts required.

**Rules note:** If you move a folder from the Archive PST to your IMAP
account after migration, any Outlook Rules targeting that folder must be
updated manually in Outlook (File > Manage Rules & Alerts) to point to the
new location. Rules do not follow folder moves automatically.

---

## Special Account Notes

**Outlook 2021 Classic -- Credential Scrambling After Add Account**
Outlook 2021 Classic has a known limitation: during the Add Account process
it forces the IMAP username to match the email address and copies that value
to the SMTP username field, overwriting any custom credentials including
AWS SES IAM keys. This is a Microsoft design limitation with no available patch.

**For POP3-AWS and POP3-STANDARD accounts** (separate IMAP/SMTP credentials,
e.g. AWS SES outbound), Script 02 attempts to correct this scrambling
automatically -- immediately after the IMAP account add is confirmed, it
reads the scrambled registry values and rewrites the IMAP/SMTP username
fields, backing up the registry subkey first. You will still need to enter
the IMAP password once manually in Outlook afterward (the corrected fields
alone don't restore a working session). **This automated correction is
implemented but not yet verified against a live account of this type** --
the project has not had a POP3-AWS or POP3-STANDARD account available to
test it against directly. If you have one of these account types and hit
credential problems after Script 02, the manual steps below work as a
fallback regardless of what the automated attempt did: Script 02 always
proceeds to completion no matter what the automated fix reports (applied,
skipped, or failed), and the manual Repair steps overwrite the same
IMAP/SMTP fields directly, so an incomplete or unsuccessful automated
attempt cannot block or complicate a manual correction afterward. OMMigrate
is aware of this Microsoft-side bug and makes a best effort to correct it
automatically where the account type allows; the manual path exists as a
safety net for every case that isn't covered, or where the correction was
attempted.

**For every other account type** (including AT&T/Yahoo legacy domains,
covered separately below, which use a Secure Mail Key rather than a
username/password pair and are not eligible for the automated fix), the
scrambling must still be corrected by hand after Script 02 completes --
this is the path that has been used and confirmed to work:
1. File > Account Settings > Account Settings
2. Select the IMAP account and click **Repair**
3. Verify and correct the IMAP username if your server uses a different login
4. Verify and correct the SMTP username and password
5. For AWS SES accounts -- SMTP username is your IAM SMTP access key ID,
   SMTP password is your IAM SMTP secret access key

> **Do not attempt to send or receive mail until credentials are corrected.**
> The Sync Issues folder in Outlook may show Local Failures until this is done --
> this is expected and clears after the credential fix.

**AT&T / Yahoo legacy domains** (ameritech.net, sbcglobal.net, att.net, etc.)
These accounts require a Secure Mail Key (app password) instead of your
regular email password. Script 02 opens your AT&T account profile
(`att.com/acctmgmt/myprofile`) in the browser automatically when these
accounts are detected. Log in and navigate to Account Security -> Secure
Mail Key -> Manage -> Generate. Enter the key when Outlook prompts for a
password during Script 02.

**AWS SES outbound SMTP**
Enter your IAM SMTP username and password when Outlook prompts during
Script 02. These are your SES SMTP credentials, not your AWS console login.
Find both values in `Config\migration_accounts.csv`.

**Send/Receive group membership**
When a POP3 account is removed and re-added as IMAP, the new IMAP account
does not automatically inherit your Send/Receive group membership. This is
an Outlook COM limitation. After Script 02, manually re-add each migrated
account to its Send/Receive group before running Script 03.

**Microsoft Exchange accounts** (`ProviderTag=EXCHANGE-SKIP`)
Exchange accounts (live.com, outlook.com, and similar) never go through
conversion -- there is nothing to migrate protocol-wise. Script 03 still
tracks and updates their Outlook rules each run, so a rule you edit
directly in Outlook's Rules Manager for an Exchange account will have its
folder target picked up on the next Script 03 run, same as any other
account's rules. These accounts do not appear in Script 03's account
picker, since there is no folder or content migration for them to do.

---

## POP3 Manual Recovery

OMMigrate is designed as a **one-way migration** from POP3 to IMAP. It does not
provide an automated rollback to POP3. If you ever need to revert an account
back to POP3 -- whether due to a failed migration, a change of plans, or any
other reason -- use the following manual procedure.

Your email data is always safe. Script 01 creates a verified PST backup before
Script 02 touches any account. That backup is your safety net.

---

### When You Might Need This

- Script 02 deleted the POP3 account but the IMAP account add failed
- The new IMAP account is not connecting correctly and you want to revert
- You changed your mind about migrating a particular account

---

### Manual POP3 Recovery Procedure

**Step 1 -- Remove the IMAP account if it was added**

In Outlook: File → Account Settings → Account Settings
Select the IMAP account → Remove
Close the dialog

**Step 2 -- Re-add the account as POP3**

In Outlook: File → Add Account
Enter the email address and click Connect
When Outlook prompts for account type, select **POP3**
Enter the server settings from `Config\migration_accounts.csv`:
- Incoming server: the value in the `IncomingServer` column
- Incoming port: the value in the `IncomingPort` column
- Outgoing server: the value in the `OutgoingServer` column
- Outgoing port: the value in the `OutgoingPort` column
- SSL settings as shown in the CSV

Enter your password when prompted and complete the setup.

**Step 3 -- Restore your email from the backup PST**

In Outlook: File → Open & Export → Open Outlook Data File
Navigate to `Documents\OutlookMigration\Backups\`
Open the PST file for this account (named `sanitized_email_address.pst`)
Your folders and email will appear in the Outlook folder pane
Drag any mail back to the account inbox as needed

**Step 4 -- Update migration_accounts.csv**

Open `Documents\OutlookMigration\Config\migration_accounts.csv`
Set the `MigrationAction` column for this account back to `SKIP`
This prevents Script 02 from attempting to convert it again

**Step 5 -- Re-run Script 00**

```powershell
.\Scripts\OMMigrate-00-Discover.ps1
```

This refreshes the manifests and confirms the account is back to POP3.

---

### Important Notes

- **Do not delete the backup PST** until the migration is fully verified and
  you are satisfied with the result. Keep it as long as you need it.
- **AT&T/Yahoo legacy accounts** (ameritech.net, sbcglobal.net, etc.) require
  a Secure Mail Key instead of your regular password for both POP3 and IMAP.
- **AWS SES outbound accounts** -- re-enter your IAM SMTP credentials when
  Outlook prompts for outgoing server authentication.
- If you are unsure of your original server settings, they are preserved in
  `Config\migration_accounts.csv` from the Script 00 discovery run.

---

## Updating OMMigrate

To update to the latest version, re-run the installer:

```powershell
cd "$env:USERPROFILE\Documents\OMMigrate"; Get-ChildItem -Recurse -Include *.ps1,*.psm1 | Unblock-File
irm https://raw.githubusercontent.com/SC-Admin567/OMMigrate/main/Install.ps1 | iex
```

The installer downloads fresh copies of all scripts and modules.
Your runtime data in Documents\OutlookMigration\ is never touched.

---

## Further Reading

- **[QUICKSTART.md](QUICKSTART.md)** -- Guided first-run walkthrough, step by step
- **[OMMigrate_CommandLine_Reference.md](OMMigrate_CommandLine_Reference.md)** -- Every parameter on every script, verified against the actual code
- **[OMMigrate_Settings_Reference.md](OMMigrate_Settings_Reference.md)** -- Every settings file key, its default, and what it actually controls
- **[CHANGELOG.md](CHANGELOG.md)** -- Version history and release notes

---

## Credits

```
OutlookMailMigrator (OMMigrate)
----------------------------------------------------
Originator & Architect:    Kirk Shallcross - Shallcross Consulting
Implementation Specialist: Anthropic Claude AI
Inception Date:            May 2026
----------------------------------------------------
"Automating the Outlook migration Google suggested couldn't be automated."
----------------------------------------------------
```

---

## Third-Party Software Notice

This software provides compatibility options with the Obsidian Markup Viewer (https://obsidian.md). 

**OMMigrate does not install Obsidian.** The optional Obsidian Add-on
(`Obsidian-Addon-Install.bat`) and the pre-configured `.obsidian\` vault
folder only set up file associations and plugin configuration for Obsidian
once it is already installed. Getting Obsidian itself, from
https://obsidian.md, is your own separate decision.

Please note the following:
* Obsidian is a proprietary application developed by its respective owners. This project is completely independent and is not affiliated with, funded by, or endorsed by Obsidian.
* The MIT License provided with this repository applies solely to the source code of OMMigrate. It does not extend to, cover, or grant any rights to the Obsidian software.
* If you choose to download and use Obsidian alongside this software, you are solely responsible for complying with Obsidian's End User License Agreement (EULA), including purchasing a commercial license if required by their terms.
* This software is provided "as is", and the authors assume no liability for the functionality, data handling, or terms of any third-party tools you choose to connect.

### Obsidian License Information
Obsidian is free for all personal, non-profit, and commercial use. You are legally allowed to use Obsidian for your work and business without paying for a license. If your organization wishes to support Obsidian's development, you may optionally purchase a Commercial License from their official website. Note that native cloud features like Obsidian Sync and Obsidian Publish require separate paid subscriptions.

---

## License

See [LICENSE.md](LICENSE.md) for terms.

---

## ☕ Support this project
Like using this tool? Consider making a donation to provide ongoing enhancements and support of this toolkit. Thank you.

**[Donate via Stripe](https://shallcross-consulting.com/stripedonate.html)**

-Or-

*(The embed button below only renders as a live donate button when this file is
opened in Obsidian with the HTML Viewer Plus plugin installed and active -- see the
Obsidian Add-on description, above. Viewing this file without Obsidian installed, it will just show
as a plain text tag -- use the "Donate via Stripe" link above instead.)*

Enter a Donation amount below(double click) and then double click Donate button
![[stripe-button.html]]
