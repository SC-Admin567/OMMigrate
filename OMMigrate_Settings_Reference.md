# OMMigrate Settings Reference

**Version:** 1.5.2

> Every key in `OMMigrate_Settings_<Profile>.json`, its default value, and
> what it actually controls. The file is created automatically on your
> first Script 00 run for each Outlook profile you use -- you don't need
> to create it by hand, only edit it if you want to change a default.

**File location:** `Documents\OutlookMigration\Config\OMMigrate_Settings_<Profile>.json`
(one file per Outlook profile -- see README.md, Directory Structure, for why).

Each key below is marked **Active** or **Not yet wired up**. Active keys are
read and applied somewhere in the codebase, verified by direct code search --
not just assumed from the schema. A few keys exist in the default settings
file but aren't consumed by any script yet; editing those has no effect
until a future release wires them up. They're listed here so the file
matches what you'll actually see on disk, with an honest note on each.

> ⚠️ **About "Not yet wired up" keys:** any key marked this way below is
> present in the settings file schema but is not currently read or acted on
> by any script. Editing it will have **no effect** on OMMigrate's behavior
> today. These are placeholders reserved for planned functionality and will
> be addressed in a future release. Do not rely on a "Not yet wired up" key
> to control real behavior -- check the Status column before changing
> anything you expect to actually take effect.

---

## DefaultPaths

| Key | Default | Status | Description |
|---|---|---|---|
| `BasePath` | `%USERPROFILE%\Documents\OutlookMigration` | Reference only | Documents the default runtime location. To actually change where OMMigrate stores its data, use the `-BasePath` parameter on each script (see OMMigrate_CommandLine_Reference.md) -- this setting is not read back by the scripts themselves. |

---

## BackupVerification

| Key | Default | Status | Description |
|---|---|---|---|
| `CheckFileSize` | `true` | Not yet wired up | Intended to toggle the backup file-size check on/off. Currently has no effect -- the size check always runs. |
| `MinimumSizeMB` | `0` | **Active** | Minimum acceptable PST backup size in MB. Backups smaller than this fail Script 01's verification step. `0` means no minimum is enforced. Raise this if you want a hard floor that catches a suspiciously small backup automatically -- for example, if you know every real account in your environment has at least a few hundred MB of mail, setting this to `50` or `100` would flag a near-empty backup (a sign something went wrong during export) as a failure instead of letting it pass silently. |
| `CheckCanOpen` | `true` | Not yet wired up | Intended to toggle a "can the PST actually be opened" check on/off. Currently has no effect. |

---

## Migration

| Key | Default | Status | Description |
|---|---|---|---|
| `DefaultImapPort` | `993` | Not yet wired up | Intended as a fallback IMAP port. Currently has no effect -- Script 00 populates `NewImapPort` per-account from its provider lookup table or the registry. |
| `DefaultSmtpPort` | `587` | Not yet wired up | Same as above, for SMTP. Currently has no effect. |
| `DefaultSSL` | `true` | Not yet wired up | Intended as a fallback SSL/TLS setting. Currently has no effect. |
| `PromptBeforeEachAccount` | `true` | Not yet wired up | Intended to control per-account Y/N prompting during Script 02. Currently has no effect -- use the `-Force` parameter to skip prompts instead. |
| `WhatIfByDefault` | `false` | Not yet wired up | Intended to make every run default to Preview mode without needing `-Preview` on the command line. Currently has no effect -- pass `-Preview` explicitly. |

---

## Reporting

| Key | Default | Status | Description |
|---|---|---|---|
| `GenerateHTML` | `true` | Not yet wired up | Intended to toggle HTML report generation off. Currently has no effect -- every script always generates its HTML report. |
| `GenerateCSV` | `true` | Not yet wired up | Intended to toggle CSV output off. Currently has no effect. |
| `OpenReportAfterRun` | `true` | **Active** | Whether the HTML report opens in your default browser automatically at the end of a run. Script 00's `-OpenReport` command-line switch always overrides this to `true` for that run; every other script follows this setting directly. Set this to `false` if you're running OMMigrate unattended or scripted (e.g. a scheduled batch of accounts overnight) and don't want a browser window popping up on every run -- you can still review the HTML report file directly from the Reports folder afterward. |
| `PreferredEditor` | `''` (empty) | **Active** | Full path to a text editor for opening logs, CSVs, and other plain-text output (e.g. `"C:\Program Files\Microsoft VS Code\Code.exe"`). Leave empty to use the Windows default file association, which is the recommended setting for broad compatibility. If empty, OMMigrate will still auto-detect Sublime Text if it's installed and prefer it automatically. Set this if your default file association for `.csv`/`.log` opens something inconvenient -- for example if double-clicking a log file launches a slow general-purpose app when you'd rather it open instantly in a lightweight text editor, or if you prefer reviewing `rules_inventory.csv` in a code editor rather than Excel. Remember Excel locks any control CSV it has open (see README.md Troubleshooting) -- if you set Excel as your `PreferredEditor`, be sure to close it before running the next script. |

---

## Logging

| Key | Default | Status | Description |
|---|---|---|---|
| `LogLevel` | `INFO` | **Active** | Default logging verbosity (`DEBUG`, `INFO`, `WARN`, or `ERROR`) when a script's `-LogLevel` parameter isn't explicitly passed. Set this to `DEBUG` here if you want every run to log at maximum detail by default while you're actively troubleshooting an environment -- more convenient than typing `-LogLevel DEBUG` on every script call. Set it back to `INFO` (or leave it) once things are working normally; `DEBUG` produces noticeably larger, noisier log files. |
| `RetainRunLogsForDays` | `90` | Not yet wired up | Intended to control automatic cleanup of old run logs. Currently has no effect -- run logs are never deleted automatically. |
| `AuditLogMaxSizeMB` | `50` | Not yet wired up | Intended to cap the cumulative audit log's size. Currently has no effect -- the audit log grows unbounded (append-only, as documented in README's Features list). |

---

## OutlookProfile

| Key | Default | Status | Description |
|---|---|---|---|
| `SelectedProfile` | `''` (empty) | **Active** | The Outlook profile selected in Script 00's profile picker. Written by Script 00 at startup and read by every subsequent script so all five scripts operate against the same profile in a session. Empty means no profile has been selected yet. You should not normally need to edit this by hand -- re-run Script 00 to change your active profile. |

---

## RulesEngine

| Key | Default | Status | Description |
|---|---|---|---|
| `ArchiveStoreMappings` | `[]` (empty array) | **Active** | Account-to-archive-PST mapping selected in Script 00's TargetStoreName picker (see README.md, multi-archive support). Each entry maps one attached archive PST (by its live-detected store name) to one or more account names whose Local-destination rules should target that PST. An empty array means no mapping has been made yet -- rules fall back to their own account's default store. Normally set via the picker, not hand-edited. Example shape: `[{"TargetStoreName": "MyArchivePST", "RuleStoreNames": ["user@example.com", "admin2@example.com"]}]` |
| `MasterArchiveNames` | `[]` (empty array) | **Active** | Archive PST names (exact store display name) that must never be auto-detached by any OMMigrate script, regardless of which script mounted them this run. Use this for permanent, always-attached archives you manage manually in Outlook -- for example your default `OMMigrate Local Archive`, or a personal auto-archive store. An empty array (the default) means the existing safe behavior applies: each script only ever detaches what its own code mounted during that run. Example: `["OMMigrate Local Archive", "Personal-Archive"]` |
| `OutlookQuitTimeoutSeconds` | `20` | **Active** | Maximum seconds `Release-OutlookCOM` waits for Outlook to exit cleanly after `Quit()` before force-killing the process. Outlook must flush all pending PST write buffers before it's safe to terminate -- a premature force-kill can silently discard unflushed changes. Larger rule sets (hundreds of rules) take longer to serialize on exit and may need a higher value; environments with few rules can often use a lower value for a faster finish. |
| `OutlookLaunchTimeoutSeconds` | `30` | **Active** | Maximum seconds `Connect-OutlookCOM`'s visible-launch path waits for Outlook to finish starting and register its COM class. Slower machines, or profiles with many PST/OST files to load at startup, may need more time. |

---

## Editing the settings file

The file is plain JSON -- edit it in any text editor (or your configured
`PreferredEditor`, above) while OMMigrate is not running. Changes take
effect on the next script run; no script needs to be restarted mid-run.

```json
{
  "RulesEngine": {
    "MasterArchiveNames": ["OMMigrate Local Archive"],
    "OutlookQuitTimeoutSeconds": 30
  }
}
```

You don't need to include every key -- only the ones you're changing from
their default. Missing keys fall back to the defaults shown above.

---

## Related documents

- **[README.md](README.md)** -- Full overview, features, and troubleshooting
- **[QUICKSTART.md](QUICKSTART.md)** -- Guided first-run walkthrough
- **[OMMigrate_CommandLine_Reference.md](OMMigrate_CommandLine_Reference.md)** -- Every script parameter
- **[CHANGELOG.md](CHANGELOG.md)** -- Version history and release notes
