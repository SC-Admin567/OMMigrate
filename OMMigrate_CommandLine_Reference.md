# OMMigrate Command Line Reference

**Version:** 1.5.2

> Full parameter reference for all five OMMigrate scripts. For the guided
> first-run walkthrough, see **QUICKSTART.md**. For a general overview,
> see **README.md**.

> **Before running any command on this page:** a fresh PowerShell window
> starts at your home folder (`PS C:\Users\yourname>`), not the OMMigrate
> project folder, and files can be blocked after a fresh download or a
> manual file replacement. Run this once per PowerShell session before any
> example below -- it's safe to run even when neither problem is actually
> present:
> ```powershell
> cd "$env:USERPROFILE\Documents\OMMigrate"; Get-ChildItem -Recurse -Include *.ps1,*.psm1 | Unblock-File
> ```
> See QUICKSTART.md, "Always Verify Your Working Directory First," for
> more detail on why both parts matter.

Every script accepts `-BasePath`, `-LogLevel`, and `-Preview`. Parameters
specific to one script are listed under that script's own section below.

All parameters below are verified against each script's actual `param()`
block, not just its help text -- a few scripts' `.PARAMETER` comments
describe behavior that doesn't match a real switch (noted where relevant).

---

## Parameters common to all scripts

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-BasePath` | string | `$env:USERPROFILE\Documents\OutlookMigration` | Override the runtime data location. Must match the `-BasePath` used by every other script in the same session. Use this if your organization standardizes on a shared or non-default data drive (e.g. `-BasePath "D:\Migration"` on a machine where `C:` is space-constrained), or if you're running more than one independent OMMigrate session on the same machine and need to keep their data fully separate. |
| `-LogLevel` | string | `INFO` | Logging verbosity written to the run log file: `DEBUG`, `INFO`, `WARN`, or `ERROR`. Use `DEBUG` when troubleshooting a failure and you need to see every step the script took, not just the summary -- the log file gets noticeably larger and noisier at this level, so it's not meant to be left on for routine runs. Use `WARN` or `ERROR` only if you specifically want a quieter log and don't need INFO-level progress detail. |
| `-Preview` | switch | off | Simulate the run -- no files written, no Outlook changes made, no manifest written. Every action is logged with a `[WHATIF]` prefix. Use this before any run where you're unsure what will happen: after editing a control CSV by hand, before running against an unfamiliar Outlook profile for the first time, or any time you want to sanity-check the plan before committing. Safe to use at any time -- it never writes anything. |
| `-Sanitize` | switch | off | Mask personal data (email addresses, names) in console output only. The log file always contains full, unmasked data. Use this when the console will be visible to someone who shouldn't see account details -- screen-sharing during a support call, presenting a demo, or recording a training video. It has no effect on what's saved to disk, so it doesn't help with data retention concerns. |
| `-Force` | switch | off | Skip per-account Y/N confirmation prompts. The initial pre-flight confirmation is still required. Use this for a large batch of accounts you've already reviewed and are confident about -- it turns an hour of babysitting Y/N prompts into a walk-away run. Avoid it on your first run with a new account type (AT&T/Yahoo, AWS SES) until you've seen what the per-account prompts actually ask for. Not available on Script 00 (nothing to confirm -- it's read-only). **What exactly gets skipped differs by script -- see the per-script notes below, especially Script 01 and Script 02.** |

---

## Script 00 -- Discover

```powershell
.\Scripts\OMMigrate-00-Discover.ps1 [-ProfileName <string>] [-BasePath <string>]
    [-LogLevel <string>] [-OpenReport] [-Preview] [-Sanitize] [-PatchRulesCSV]
```

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-ProfileName` | string | (all profiles) | Target a specific Outlook profile by name. If omitted, every profile found on the machine is scanned. Use this on a machine with multiple Outlook profiles (e.g. a shared workstation, or a technician's machine with a client-testing profile alongside their own) where scanning every profile would be slow or would surface accounts you don't want to touch this run. |
| `-OpenReport` | switch | off (follows settings) | Force the Discovery HTML report to open in the default browser after the run. Passing this always overrides `Reporting.OpenReportAfterRun` in the settings file; omitting it uses whatever the settings file says. Use this on a one-off run where you want to review the report immediately regardless of your usual setting -- for example if you've set `OpenReportAfterRun` to `false` for routine unattended runs but want to eyeball this particular one. |
| `-PatchRulesCSV` | switch | off | Patch an existing `rules_inventory_<Profile>.csv` in place instead of a full COM rules enumeration. Skips the normal discovery scan for rules specifically -- useful for applying a corrected value without waiting through a full rescan. Use this after manually fixing a bad value in the rules CSV (e.g. a wrong folder path) when you don't want to wait through a full re-scan of every account just to pick up that one edit. |

**Examples:**
```powershell
# Standard first run
.\Scripts\OMMigrate-00-Discover.ps1

# Target one specific Outlook profile
.\Scripts\OMMigrate-00-Discover.ps1 -ProfileName "Outlook"

# Dry run -- see what would be discovered without writing any files
.\Scripts\OMMigrate-00-Discover.ps1 -Preview

# Patch rules_inventory_<Profile>.csv after a manual correction, skip full rescan
.\Scripts\OMMigrate-00-Discover.ps1 -PatchRulesCSV
```

---

## Script 01 -- Backup

```powershell
.\Scripts\OMMigrate-01-Backup.ps1 [-BasePath <string>] [-LogLevel <string>]
    [-SkipVerification] [-Force] [-RefreshFolderMap] [-Preview] [-Sanitize]
```

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-SkipVerification` | switch | off | Skip the post-export PST integrity verification step. **Not recommended** -- verification is your confirmation that the backup is valid before Script 02 removes anything. The only defensible reason to use this is a very large backup where you've already confirmed verification passes reliably and you're re-running Script 01 repeatedly during testing and want to save the time each verification pass costs -- never skip it on a run whose backup you intend to actually rely on. |
| `-RefreshFolderMap` | switch | off | Re-run mode: skips all backup and Archive pre-build steps and opens the folder destination picker directly, to reassign Server/Local destinations without re-running backups. `folder_map_<Profile>.csv` must already exist. Use this when you've already backed up successfully but changed your mind about which folders should be Server vs Local -- it saves you from re-exporting PSTs you already have. |

> **`-Force` on a re-run silently overwrites existing backups.** On a first
> backup for an account, the per-account prompt defaults to Yes either way,
> so `-Force` changes little. But if a backup already exists for an
> account, the normal prompt defaults to **No** ("keep existing") -- `-Force`
> flips that default and re-exports over the existing PST automatically,
> with no confirmation. Safe if you specifically want a fresher backup;
> risky if you assumed `-Force` only skips prompts without changing what
> gets kept.

> A minimum backup size threshold exists (`BackupVerification.MinimumSizeMB`
> in the settings file), but it is not exposed as a command-line parameter --
> set it in `OMMigrate_Settings_<Profile>.json` if you need to change it
> from the default of 0 (no minimum enforced).

**Examples:**
```powershell
# Standard backup run
.\Scripts\OMMigrate-01-Backup.ps1

# Reassign folder destinations without re-backing-up
.\Scripts\OMMigrate-01-Backup.ps1 -RefreshFolderMap

# Back up all accounts without per-account Y/N prompts
.\Scripts\OMMigrate-01-Backup.ps1 -Force
```

---

## Script 02 -- Convert

```powershell
.\Scripts\OMMigrate-02-Convert.ps1 [-BasePath <string>] [-LogLevel <string>]
    [-SkipBackupCheck] [-Force] [-Preview] [-Sanitize]
```

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-SkipBackupCheck` | switch | off | Skip the backup file existence and size verification gate. **Not recommended** -- only use if you have independently verified all backups and understand the risk of proceeding without them. There is essentially no routine scenario for this; it exists as an escape hatch for a specific recovery situation (e.g. you've verified backups by hand through some other means and the automated check is failing for an unrelated reason you've already root-caused). Using it removes the one safeguard standing between you and an account with no fallback if the IMAP add fails. |

> **What `-Force` actually skips here matters more than on other scripts.**
> Script 02's normal per-account prompt is the moment you see the verified
> backup size ("your email data is safe") and confirm before POP3 is
> actually removed and IMAP added -- the one irreversible step in the whole
> pipeline. `-Force` skips that entire confirmation, defaulting every
> account to Yes with no per-account review. Combined with `-SkipBackupCheck`
> it removes every safeguard on this script at once -- avoid combining the two.

**Examples:**
```powershell
# Standard conversion run
.\Scripts\OMMigrate-02-Convert.ps1

# Dry run before committing to any account changes
.\Scripts\OMMigrate-02-Convert.ps1 -Preview
```

---

## Script 03 -- Restore

```powershell
.\Scripts\OMMigrate-03-Restore.ps1 [-BasePath <string>] [-LogLevel <string>]
    [-RefreshRulesOnly] [-SaveBatchSize <int>] [-ArchivePSTName <string>]
    [-Force] [-RecreateRules] [-Preview] [-Sanitize]
```

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-RefreshRulesOnly` | switch | off | Skip item migration and run the rules update only. Use this after editing `rules_inventory_<Profile>.csv` directly (e.g. correcting a folder target path) to rerun just the rules deployment without re-copying items from the backup PSTs. |
| `-SaveBatchSize` | int | `1` | Batches rule `Save()` calls during a full rule rebuild. Advanced/rarely needed -- testing during development found batching alone does not meaningfully improve rebuild time (see CHANGELOG.md, 1.5.0, Performance), so there's little practical reason to change this from the default. Leave it at `1` unless you're specifically diagnosing rule-deployment performance yourself and want to compare batch sizes. |
| `-ArchivePSTName` | string | `OMMigrate_Archive.pst` | Custom filename for the default local Archive PST. If you've attached multiple archive PSTs via the TargetStoreName picker (Script 00), this only affects the default/fallback archive. Change this if `OMMigrate_Archive.pst` collides with an existing file name in your Backups folder, or if your organization has a file-naming convention for archive PSTs you need to follow (e.g. including a client name or ticket number). |
| `-RecreateRules` | switch | off | Recovery-only: recreate Outlook Rules from each account's backup PST onto the new IMAP store. Conditions and non-folder actions are preserved verbatim; only the folder target is remapped. Rules that already exist by name on the target store are skipped (safe to re-run). Best-effort -- rules that can't be recreated are logged and counted but don't abort the run. Use this specifically to recover from Rules Manager corruption or an accidental bulk rule deletion, not as a normal part of the pipeline. Recommended: run after folder migration is complete so target folders already exist. |

**Examples:**
```powershell
# Standard restore run
.\Scripts\OMMigrate-03-Restore.ps1

# Re-deploy rules only, after a manual rules_inventory_<Profile>.csv correction
.\Scripts\OMMigrate-03-Restore.ps1 -RefreshRulesOnly

# Custom name for the default Archive PST
.\Scripts\OMMigrate-03-Restore.ps1 -ArchivePSTName "MyArchive_2026.pst"

# Recovery: rebuild rules from backup PSTs after Rules Manager corruption
.\Scripts\OMMigrate-03-Restore.ps1 -RecreateRules
```

---

## Script 04 -- Artifacts

```powershell
.\Scripts\OMMigrate-04-Artifacts.ps1 [-BasePath <string>] [-LogLevel <string>]
    [-Force] [-Preview] [-Sanitize]
```

Script 04 has no parameters beyond the common set above. It reads directly
from each live account (not from the backup PST) and migrates Calendar,
Contacts, Tasks, Notes, and Journal items with duplicate detection.

**Example:**
```powershell
.\Scripts\OMMigrate-04-Artifacts.ps1
```

---

## Utilities

The scripts in `Utilities\` are standalone maintenance tools, not part of
the main pipeline. Each can be run independently at any time.

| Script | Purpose |
|---|---|
| `OMMigrate-CleanRulesCSV.ps1` | Cleans up `rules_inventory_<Profile>.csv` -- removes stale or malformed rows. |
| `OMMigrate-FindDuplicateRules.ps1` | Scans all Outlook stores for rules with `(n)` duplicate-suffix names, writes `DuplicateRules.csv` to the Reports folder. |
| `OMMigrate-RemoveDuplicateRules.ps1` | Reads `DuplicateRules.csv` and deletes the listed rules via COM. Supports `-WhatIf` preview mode. |

Run any Utilities script with `-WhatIf` first if you're unsure what it will change:
```powershell
.\Utilities\OMMigrate-RemoveDuplicateRules.ps1 -WhatIf
```

---

## VBA macros (Outlook)

The macros in `OutlookVBAMacros\` run inside Outlook's own VBA editor, not
from PowerShell -- see **QUICKSTART.md, Step 0.5** for installation.

| Macro | Module | Purpose |
|---|---|---|
| `DeployConsolidatedRules` | Module3.bas | VBA counterpart to Script 03's rules deployment -- the consolidated rule engine. |
| `ResortRulesByLabel` | Module3.bas | Repositions all rules in a live Rules collection into alphabetical order via `ExecutionOrder` writes only. |
| `CorrectArchiveFolders` | Module7.bas | CSV-driven archive misroute correction with picker-based store selection. |
| `DeleteDups` | Module1.bas | Duplicate email detection and removal, with a folder picker. |
| `EmptyAllTrashFolders` | Module2.bas | Empties Deleted Items across all attached stores. |

---

## Related documents

- **[README.md](README.md)** -- Full overview, features, and troubleshooting
- **[QUICKSTART.md](QUICKSTART.md)** -- Guided first-run walkthrough
- **[OMMigrate_Settings_Reference.md](OMMigrate_Settings_Reference.md)** -- Every settings file key, its default, and what it controls
- **[CHANGELOG.md](CHANGELOG.md)** -- Version history and release notes
