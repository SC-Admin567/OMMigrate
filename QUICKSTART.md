# OMMigrate Quick Start Guide

> New to OMMigrate? Start here. This guide walks you through your first
> migration from install to a verified working IMAP account -- safely,
> one account at a time.

![OMMigrate pipeline architecture -- five sequential scripts, each gated on the previous script's completion manifest](docs_images/architecture_overview.svg)

Five scripts run in this order: Discover, Backup, Convert, Restore,
Artifacts. Each one hard-checks that the previous script finished
successfully before it will run. This guide walks you through all five,
one account at a time. For the full picture, see **README.md, Architecture
at a Glance**.

> ⚠️ **Before you start:** OMMigrate restructures your Outlook Rules to
> work with a new mail-organization model -- it's not a like-for-like
> copy of your existing setup, and it's not reversible mid-run. If you
> haven't already, read **README.md, "Read This Before You Decide to Use
> OMMigrate,"** first -- especially if you're evaluating this for a
> client or an organization with server-retention requirements.

---

## Step 0 -- Install

### Option A -- One-liner (recommended)

Open PowerShell and run these two commands:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
irm https://raw.githubusercontent.com/SC-Admin567/OMMigrate/main/Install.ps1 | iex
```

The execution policy line is a one-time step required on new machines.
If your policy is already RemoteSigned or Unrestricted you can skip it.

> **Why two commands instead of one?** PowerShell blocks script execution
> on new machines by default. The execution policy must be set before any
> script can run -- including the installer. The one-liner (`irm ... | iex`)
> pipes Install.ps1 directly into PowerShell memory so it never writes a
> blocked file to disk, which is why Unblock-File is not needed here.

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

The `Unblock-File` step is required here because Windows marks any file
downloaded from the internet as blocked. Without it PowerShell will refuse
to run Install.ps1 even after the execution policy is set.

---

The installer:
- Creates `Documents\OMMigrate\` with all scripts and modules
- Creates `Documents\OutlookMigration\` for runtime data (logs, reports, backups)
- Sets your PowerShell execution policy automatically on future runs
- Unblocks all downloaded files automatically

When it finishes you are ready to run Script 00 -- see Step 1, below.

> **Step 0.5, next, is optional.** It installs Outlook VBA macros that
> Script 03 (Restore) can optionally use, and a couple of standalone
> maintenance tools. Nothing in Script 00, 01, or 02 depends on it. If
> you'd rather get to your first Discovery run immediately, skip ahead to
> **Step 1** and come back to Step 0.5 later -- it doesn't need to happen
> in any particular order relative to Scripts 00-02.

---

## Step 0.5 -- Install VBA Macros (Optional)

OMMigrate includes several Outlook VBA macros that support the rules engine
and provide standalone maintenance tools (duplicate cleanup, archive misroute
correction). These live in `Documents\OMMigrate\OutlookVBAMacros\` and must be
copied into Outlook's VBA project manually -- Outlook does not allow scripts
to install VBA code automatically.

**Files to install** (all in `OutlookVBAMacros\`):

| File | What it does |
|---|---|
| `Module1.bas` | Duplicate email detection and removal |
| `Module2.bas` | Support functions (e.g. empty all Trash folders) |
| `Module3.bas` | Consolidated rule engine -- VBA counterpart to Script 03's rules deployment |
| `Module7.bas` | Archive misroute correction (`CorrectArchiveFolders`) |
| `FrmArchivePicker.frm` + `.frx` | Archive picker form used by Module7 |
| `FrmProgress.frm` + `.frx` | Progress form |
| `ThisOutlookSession.cls` | Outlook session event handlers |

> **`.frm`/`.frx` files stay paired.** Each `.frm` file has a matching `.frx`
> binary resource file with the same name -- both must be imported together
> and kept in the same folder. Do not rename either one.

**Installation procedure:**

1. Open Outlook
2. Press **Alt+F11** to open the VBA editor
3. In the Project Explorer (left pane), find **Project1 (VbaProject.OTM)**
4. Right-click **Modules** -> **Import File...** and import `Module1.bas`,
   `Module2.bas`, `Module3.bas`, and `Module7.bas` one at a time
5. Right-click **Forms** -> **Import File...** and import `FrmArchivePicker.frm`
   and `FrmProgress.frm` (Outlook will automatically pick up the matching
   `.frx` file from the same folder -- do not import the `.frx` files directly)
6. In the Project Explorer, double-click **ThisOutlookSession** under
   **Microsoft Outlook Objects**
7. Open `ThisOutlookSession.cls` in a text editor, select all, copy
8. Paste into the `ThisOutlookSession` code window in the VBA editor,
   replacing anything already there (back up any existing custom code first
   if you have any)
9. Press **Ctrl+S** to save, then close the VBA editor

**Assigning macros to the Ribbon (with custom icons):**

The macros most useful for day-to-day use are `Module3.DeployConsolidatedRules`,
`Module7.CorrectArchiveFolders`, and `Module1.DeleteDups`. Outlook always adds
a macro under its raw VBA name (e.g. `Project1.DeployConsolidatedRules`) with
a generic icon -- give each one a short display name and a real icon so
they're actually usable from the Ribbon day to day.

1. In Outlook -- File > Options > Customize Ribbon
2. Under **Choose commands from**, select **Macros** -- every installed
   macro appears in the left-hand list as `Project1.<MacroName>`
3. On the right, under **Customize the Classic Ribbon**, select the tab you
   want the macros on (e.g. **Home**), then click **New Group** to create a
   dedicated group for them (shown as "Custom Tools (Custom)" in the example
   below) -- or use an existing custom group if you already have one
4. With your new group selected on the right, select a macro on the left
   and click **Add >>**. Repeat for each macro you want on the Ribbon

![Outlook Customize Ribbon showing macros added to a custom group, with the right-click context menu open on one item showing Rename](QUICKSTART_images/snap-22.png)

5. Right-click the macro you just added (in the right-hand tree) and choose
   **Rename**

![Rename dialog showing the Symbol icon grid and Display name field](QUICKSTART_images/snap-23.png)

6. In the Rename dialog: pick an icon from the **Symbol** grid, and type a
   short **Display name** (e.g. "Deploy Consolidated Rules", "Delete
   Duplicate Emails")
7. Click **OK** in the Rename dialog, then **OK** again to close Outlook
   Options

The icon you picked now shows directly on the Ribbon under your chosen
group, labeled with the display name instead of the raw macro name.

> **Quick Access Toolbar instead of the Ribbon:** the same **Choose
> commands from > Macros** picker, Add button, and right-click Rename flow
> exist under File > Options > Quick Access Toolbar -- useful if you want
> one or two macros always visible in the title bar (top-left, always
> visible) instead of grouped on a specific Ribbon tab.

---

## Step 0.75 -- Install Obsidian Add-on (Optional)

OMMigrate optionally supports Obsidian (https://obsidian.md), a free
third-party Markdown viewer/editor, as a way to view this project's `.md`
documentation files with full formatting. This is entirely optional --
nothing in Scripts 00-04 depends on it, and it does not affect any part of
the migration itself. Treat it the same as Step 0.5 above: skip it now and
come back later if you'd like, in any order relative to Scripts 00-02.

> **Install Obsidian itself first -- that's a separate, independent
> decision.** Obsidian is free third-party software, not part of OMMigrate.
> Everything below -- the add-on files and the pre-configured `.obsidian\`
> vault folder with the HTML Viewer Plus plugin -- only *configures*
> Obsidian once it's installed; none of it installs the application itself,
> and none of it works without it. Get Obsidian from https://obsidian.md
> before continuing. See README.md, "Third-Party Software Notice," for
> full licensing details.

**What it does:** the add-on associates `.md` files with Obsidian so that
double-clicking a `.md` file (like this one) opens it in Obsidian instead
of Notepad, and opens it directly in the correct vault/note context when the
file lives inside an Obsidian vault.

**Files involved** (all in `Documents\OMMigrate\`):

| File | What it does |
|---|---|
| `Obsidian-Addon-Install.bat` | Installs the add-on -- must be run as Administrator |
| `Obsidian-Addon-Uninstall.bat` | Removes the add-on and its installed files -- must be run as Administrator |
| `Obsidian-Addon-DeployLink.reg` | Registry file imported by Install.bat -- associates `.md` with Obsidian |
| `Obsidian-Addon-UninstallLink.reg` | Registry file imported by Uninstall.bat -- reverts the `.md` association |
| `Obsidian-Addon-Opener.ps1` | Helper script -- resolves the correct Obsidian vault for the clicked `.md` file |
| `Obsidian-Addon-SilentRunner.vbs` | Helper script -- runs Opener.ps1 without flashing a console window |

**Installation procedure:**

1. Install Obsidian from https://obsidian.md if you haven't already
2. Right-click `Obsidian-Addon-Install.bat` and choose **Run as administrator**
3. Confirm the UAC prompt if shown
4. Once it reports "Installation Complete," double-clicking any `.md` file
   will open it in Obsidian

**To uninstall later:** right-click `Obsidian-Addon-Uninstall.bat` and choose
**Run as administrator**. This reverts the `.md` file association back to
its previous default and removes the add-on's helper files.

> **This project also includes an Obsidian vault** (the `.obsidian\` folder)
> with the HTML Viewer Plus community plugin pre-configured. This plugin is
> what renders the donate-button embed near the bottom of README.md when the
> file is opened as part of this vault in Obsidian. It is optional and has
> no effect on the migration scripts. A `debug.log` file is included under
> `.obsidian\plugins\html-viewer-plus\` -- it stays empty unless you enable
> the plugin's debug logging option yourself.

---

## Before You Run Any Script

### Requirements

| Requirement | How to verify |
|---|---|
| Windows 10 or 11 (64-bit) | Settings -> System -> About |
| Classic Outlook 2016 / 2019 / 2021 | Open Outlook -> File -> Office Account |
| PowerShell 5.1 or higher | Run `$PSVersionTable.PSVersion` in PowerShell |
| Run as the user who owns the Outlook profile, in a normal PowerShell window | Do NOT use "Run as Administrator" |

> **Not supported:** New Outlook (Microsoft Store version), Outlook 2013
> or earlier, or macOS Outlook.

> **How to open the right kind of PowerShell window:** click Start, type
> `PowerShell`, and press **Enter** (or click the regular result) -- not
> "Run as administrator" from the right-click menu. If you're not sure
> which kind of window you have open, look at the title bar: an elevated
> window says "Administrator: Windows PowerShell"; a normal one just says
> "Windows PowerShell". Nothing in OMMigrate needs elevated permissions --
> using an Administrator window is more likely to cause problems than
> fix them, because your Outlook profile belongs to your regular user
> account and an elevated session can't see it. If Script 00 finds no
> accounts at all, this is the first thing to check.

### Back Up Your Existing Outlook Data First (Safety Net)

**Do this before running any OMMigrate script, even Script 00.** OMMigrate's
own Script 01 backs up your email *content* (the actual messages) before
Script 02 touches your accounts -- but your Outlook **rules** are a
completely separate thing, and nothing in the pipeline exports them for
you automatically. If a rules-side problem comes up mid-migration, the
only way back to your original rules is Outlook's own native export --
and that only works if you created the export *before* anything changed.
Five minutes now can save a full rules rebuild later.

**1. Back up your existing PST files.** Find them via **File -> Account
Settings -> Account Settings** (yes, click "Account Settings" twice --
the first is a menu, the second is the actual dialog) **-> Data Files**
tab in Outlook -- this lists the exact file paths for every PST attached
to your profile, and is the most reliable way to find them since the
actual location varies by Outlook version and setup history.

> **If you'd rather browse for them manually:** PST files are commonly
> found in `Documents\Outlook Files\` (the default on newer Outlook
> installs) or `AppData\Local\Microsoft\Outlook\` (older installs, or
> an account originally set up in an older Outlook version). A single
> profile can have PSTs in either location, or both. `AppData` is hidden
> by default -- turn on "Hidden items" in File Explorer's View tab to
> see it. If a PST isn't in either default location, an admin may have
> originally pointed Outlook at a custom folder (a shared drive, a
> different local path) when it was first created -- the Data Files tab
> will still show the correct path even then. When in doubt, trust the
> Data Files tab over a manual search.

Copy those files to a location *outside* your `Documents\OMMigrate` and
`Documents\OutlookMigration` folders (an external drive, cloud storage,
or another folder entirely) -- see "Where to Keep Your Backups," below,
for why.

**2. Export your existing Outlook rules, for every email account.** In
Outlook: **Rules -> Manage Rules & Alerts -> Email Rules tab -> Options ->
Export Rules**. Do this once per account/store -- rules are exported per
account, not all at once. Save each export somewhere you'll be able to
find it later; a clear filename (e.g. including the account name and
today's date) will help.

This rules export is your fallback if you ever need to restore your
*original*, pre-migration rules via Outlook's **Import Rules** button --
independent of anything OMMigrate does.

> **Where to keep your backups:** save both the PST backups and the rules
> exports somewhere *outside* any OMMigrate-managed folder
> (`Documents\OMMigrate`, `Documents\OutlookMigration`). If you ever need
> to reinstall or reset OMMigrate, or if a script run modifies files in its
> own working folders, a backup stored inside those folders isn't a safety
> net anymore. A separate "Recovery" folder somewhere else on your machine,
> or on an external/cloud location, works well.

**After migration is complete**, repeat the rules export one more time
(Rules -> Manage Rules & Alerts -> Email Rules tab -> Options -> Export
Rules, for every account) and save that alongside your pre-migration
export in the same Recovery location. This gives you both a "before" and
an "after" snapshot -- if you ever need to fall back to your new IMAP
rules via **Import Rules**, this second export is what you'd use.

**Keep doing this periodically, not just once.** Your rules will keep
changing after migration as you correct `SendersDomain` and
`TargetFolderPath` values in `rules_inventory.csv` (covered in "Step 2 --
Edit the Control Files," further down) and rerun
Script 03 to redeploy them. Each fresh Export Rules snapshot is a new
point-in-time recovery option -- take one whenever your rules have
settled into a state you'd want back, not only right after your first
migration.

> **Disclaimer:** OMMigrate is provided as-is, with no warranty of any
> kind. These backup and rules-export steps are your safety net --
> OMMigrate's own PST backups only cover message content, not rules, and
> the tool cannot recover data or rules that were never independently
> backed up before a script ran. Running OMMigrate without completing
> these steps is done at your own risk. The author and contributors are
> not responsible for lost, corrupted, or unrecoverable email, rules, or
> other data resulting from use of this tool, including cases where these
> backup steps were skipped or incomplete. See **README.md, License**,
> for the full terms.

### Always Verify Your Working Directory First

When you open a new PowerShell window, it starts at your home folder --
you'll see a prompt like `PS C:\Users\yourname>`, not the OMMigrate
folder. Before running any script, `cd` into the project folder and
unblock the files in one line:

```powershell
cd "$env:USERPROFILE\Documents\OMMigrate"; Get-ChildItem -Recurse -Include *.ps1,*.psm1 | Unblock-File
```

> **Why only `*.ps1,*.psm1`?** This is scoped to exactly what Scripts 00-04
> need -- the Scripts themselves and the Modules they import. It does not
> cover the Obsidian Add-on's `.bat`/`.reg`/`.vbs` files from Step 0.75
> above; those aren't used by the migration scripts at all, and are
> unblocked separately (automatically, by Install.ps1) only if you choose
> to use that optional add-on.

This is safe to run every time, even if the files are already unblocked --
it's a no-op in that case. Confirm it worked:

```powershell
Get-Location
```

It should show `C:\Users\yourname\Documents\OMMigrate` (with your actual
Windows username in place of `yourname`). If it doesn't, run the `cd` line
above again.

> **Why both `cd` and `Unblock-File`, every session?** Pressing `Ctrl+C`
> to stop a running script can silently reset your working directory back
> to your home folder -- any command run from the wrong folder fails with
> a path error. Separately, Windows marks any file that's been freshly
> downloaded or replaced (a manual file swap, a partial re-install) as
> blocked -- `Unblock-File` clears that. Chaining both together means you
> only need to remember one line, and it's harmless to run even when
> neither problem is actually present.

Every command example later in this guide assumes you've already run
the line above in your current PowerShell window.

### Test With One Account First

**Do not migrate all accounts on your first run.**

Pick one test account and verify everything works before migrating the rest.

| Choose an account that is... | Reason |
|---|---|
| A standard POP3 on your own mail server | Simplest path -- no special credentials |
| Small PST file | Backup and folder migration run quickly |
| Less actively used | Less disruption while you verify |
| Not an AT&T/Yahoo legacy domain | Save special-credential accounts for later |
| Not an AWS SES outbound account | Keep first run as simple as possible |

---

## The Five Scripts -- In Order

```
Script 00 -- Discover    READ ONLY -- safe to run any time, makes no changes
Script 01 -- Backup      Exports PST backups -- non-destructive, safe to re-run
Script 02 -- Convert     Removes POP3, adds IMAP -- requires Y/N per account
Script 03 -- Restore     Migrates folders and consolidates Outlook Rules
Script 04 -- Artifacts   Migrates Calendar, Contacts, Tasks, Notes, and Journal
```

Each script writes a completion manifest on success and the next script
reads it before proceeding -- this is a hard gate, not just a suggestion.
If a prerequisite step's manifest is missing or shows a failure, the next
script stops with a clear error instead of running on incomplete data.

> This guide runs every script with its default settings, no parameters.
> For advanced options -- `-Preview` mode, custom log levels, skipping a
> verification step, and more -- see **OMMigrate_CommandLine_Reference.md**.

---

## Step 1 -- Run Script 00 (Discovery)

Script 00 is read-only. It makes no changes. Run it as many times as you like.

```powershell
.\Scripts\OMMigrate-00-Discover.ps1
```

If Outlook is running, the script will offer to close it automatically.
Press Enter (defaults to Y) and it continues. You do not need to close
Outlook manually before running.

**What Script 00 does:**
- Scans your Outlook profile settings from the Windows Registry
- Classifies every account -- POP3, IMAP, or Exchange
- Reads server names, ports, SSL settings, and data file paths
- Generates control files in `Documents\OutlookMigration\Config\`
- Opens a selection window so you can choose which accounts to migrate now
- Opens the Discovery Report in your browser

**When it finishes, verify the Discovery Report:**
- Are all your email accounts listed?
- Are the account types correct (POP3 / IMAP / Exchange)?
- Are the server names and ports correct?
- Are folder counts reasonable?

If anything looks wrong -- investigate before going further. Script 00 is
the only completely risk-free step to repeat.

> **Rerunning Script 00 resyncs `SendersDomain` from Outlook (as of 1.5.1):**
> each run reads every Outlook rule's own SenderAddress condition and
> writes it into the `SendersDomain` column of `rules_inventory.csv`,
> overwriting whatever was there before -- including a manual CSV edit
> you haven't yet pushed to Outlook via Script 03. If you've hand-edited
> `SendersDomain` and want to keep that edit, run Script 03 first so it
> reaches the live rule -- then a later Script 00 rerun will just read
> that same value back. See **README.md, Runtime Output Files** for the
> full explanation.

---

## Step 2 -- Edit the Control Files

Open `Documents\OutlookMigration\Config\` and edit these files.

> There's also a settings file in the same folder
> (`OMMigrate_Settings_<Profile>.json`) for defaults like timeouts and
> multi-archive mappings -- you don't need it for a first run. See
> **OMMigrate_Settings_Reference.md** if you want to customize it later.

> **Filenames include your Outlook profile.** All three control files are
> named after your active Outlook profile, e.g. `migration_accounts_Outlook.csv`
> instead of plain `migration_accounts.csv`. This guide refers to them by
> their base name for readability -- look for the profile-suffixed version
> on disk. If you only use one Outlook profile you'll only ever see one
> version of each file.

> **Important:** Close Excel before running any script. Excel locks
> `migration_accounts.csv` exclusively -- any script that reads the CSV
> will fail if the file is open. Always save and close Excel first.

### migration_accounts.csv -- passwords and account selection

After Script 00 runs, a selection window opens automatically showing all
discovered POP3 accounts. Check each account you want to migrate now and
click OK. Unchecked accounts are set to SKIP and can be migrated later by
re-running Script 00 and checking them then.

After making your selection, open the CSV and complete these fields for
each account marked MIGRATE:

| Column | Value | Notes |
|---|---|---|
| `Password` | your email password | Never logged or stored by OMMigrate |
| `NewImapServer` | IMAP server hostname | Pre-filled -- confirm or correct |
| `NewImapPort` | IMAP port | Defaults to 993 -- confirm or correct |

Example after selection and password entry:
```
EmailAddress              MigrationAction   AccountType   Password
testaccount@domain.com    MIGRATE           POP3          yourpassword
account2@domain.com       SKIP              POP3          [ENTER_PASSWORD]
account3@domain.com       SKIP              POP3          [ENTER_PASSWORD]
```

> `SKIP` accounts don't need a password -- the `[ENTER_PASSWORD]`
> placeholder is just the CSV's default value and is safe to leave as-is
> for any account you're not migrating this run. Only fill in a real
> password for accounts marked `MIGRATE`.

> Save the file and close Excel before running the next script.
> If using Excel, save as CSV (comma delimited), not as .xlsx.

**After a successful Script 02 migration -- automatic:**

Script 02 automatically updates the CSV for each successfully converted
account. You do not need to edit the CSV manually after Script 02:

| Column | Updated to | Why |
|---|---|---|
| `MigrationAction` | `FOLDER-ONLY` | Marks account for Script 03 folder migration |
| `ProviderTag` | `IMAP-CONVERTED` | Reflects converted state -- Script 03 uses this |
| `AccountType` | `IMAP` | Reflects current account type |

**After a successful Script 03 migration -- automatic:**

Script 03 automatically updates `MigrationAction` from `FOLDER-ONLY` to
`COMPLETE` in the CSV after each account successfully migrates. You do not
need to update this manually. Accounts marked `COMPLETE` are excluded from
the Script 03 account picker on future runs.

### folder_map.csv -- where your folders go

Review the `Destination` column for each folder:

| Value | Meaning |
|---|---|
| `Server` | Folder lives on the IMAP server -- visible on all devices and webmail |
| `Local` | Folder stays in a local Archive PST -- desktop only |
| `Skip` | Folder not migrated -- typically empty folders |

Change `Local` to `Server` for any folders you want on your phone,
tablet, or webmail. Leave historical archive folders as `Local`.

> **After migration:** Local folders can be moved to the IMAP server at any
> time by dragging and dropping them in Outlook -- no scripts required.
> If you do this, update any Outlook Rules targeting that folder manually
> (File > Manage Rules & Alerts) -- rules do not follow folder moves.

### rules_inventory.csv -- how OMMigrate maps rules to their new folders

Script 00 generates this file automatically, and Script 03 reads it to
know which rules need their folder target updated -- but **expect to
review and correct it on a first run**, not just skim past it. The
console output during a run mentions it only briefly; this section
covers what you actually need to know before trusting it.

**Why this file exists:** your old POP3 rules pointed at folders under
your POP3 account's own store. After migration, mail lives under a
different store (your new IMAP account, or a local Archive PST) --
so every rule's folder target has to be remapped. This file is how
OMMigrate tracks that remapping, one row per rule.

**How `SendersDomain` and `TargetFolderPath` actually work together:**
`SendersDomain` is the real, functional piece -- it's the sender-match
value Script 03 writes into the live Outlook rule's condition. `TargetFolderPath`
is simply which folder that rule files matching mail into. They are two
independently-edited columns, not one derived from the other.

**Why the first run's `TargetFolderPath` is very likely wrong:** Script 00
needs *something* in every row's `TargetFolderPath` to give Script 03 a
starting point, so on a first pass it seeds the field with a best guess --
usually just the last folder-name segment from the rule's original action,
not a value validated against real sender data. **Expect to correct most
of these** before trusting a first-run CSV -- the seeded value exists to
prime the field, not as a reliable recommendation. Rows where Script 00
couldn't even generate a guess are flagged in the `Notes` column with
`VERIFY REQUIRED`; the same "verify before trusting it" caution applies
to every other row too, guess or not.

**How to hand-edit `SendersDomain`:** enter the sender address(es) or bare
domain fragments this rule should match on, **separated by a single
space** -- e.g. `amazon aws` matches either. Commas and semicolons are
NOT supported separators; using either will cause values to be dropped
or misread. A full email (`billing@example.com`) and a bare domain
fragment (`example`) are both valid token shapes, and can be mixed in
the same cell.

**Column reference:**

| Column | What it means | When to hand-edit it |
|---|---|---|
| `RuleStoreName` | The account (mailbox/store) this rule belongs to. | Not normally edited -- reflects Outlook's own rule ownership. |
| `TargetStoreName` | Which store (IMAP account or Archive PST) the rule's folder target lives under after migration. | Set via Script 00's TargetStoreName picker for most rules. Hand-edit one row to override the picker for just that rule -- your edit is preserved on the next Script 00 run as long as it doesn't match what the picker would currently produce. |
| `RuleName` | The rule's name, exactly as it appears in Outlook's Rules and Alerts. | Not edited here -- rename the rule in Outlook itself. |
| `LastDeployedRun` | Timestamp of the last time Script 03 successfully consolidated/deployed this rule. Blank means pending. | **Blank this cell to force Script 03 to reprocess the rule** on its next run (e.g. after you've hand-corrected `TargetFolderPath` or `SendersDomain` below). |
| `LastTargetRun` | Timestamp of the last time Script 03's folder-target remap ran for this rule. Tracked separately from `LastDeployedRun`. | Same idea as `LastDeployedRun` -- blank it to force just the folder-remap step to rerun. Don't blank both columns for the same rule in the same session; if you do, that rule is skipped for one run and picked up again on the next. |
| `TargetFolderPath` | The new, post-migration folder path this rule should move mail to. **On a first run, this is usually a best-guess seed value -- see above -- not a verified target.** | Verify before trusting it, even when populated. Hand-edit if it's blank (a `Notes` flag will say why) or if the guessed/detected folder is wrong. |
| `SendersDomain` | The sender address(es) or bare domain fragments this rule actually matches on -- the functional condition value Script 03 deploys. See "How `SendersDomain` and `TargetFolderPath` actually work together," above. | Hand-edit to correct or add match values. **Space-separate multiple values** (e.g. `amazon aws`) -- commas and semicolons are not supported and will cause values to be dropped or misread. |
| `NeedsFolderUpdate` | `True`/`False` -- whether Script 03 should update this rule's folder target at all. | Set to `False` to tell Script 03 to leave this rule's folder target alone even though a target was detected. Also editable via Script 00's rules-review popup. |
| `IsEnabled` | Whether the rule is enabled in Outlook. | Not normally edited here -- reflects Outlook's own rule state. |
| `ExecutionOrder` | The rule's position in Outlook's rule execution order. | Not normally hand-edited -- renumbered automatically to keep account groups together. |
| `RuleType` | Incoming or outgoing rule. | Not edited -- reflects the rule's actual type in Outlook. |
| `StopProcessing` | Whether the rule has "stop processing more rules" set. | Not normally edited here -- reflects the rule's own Outlook setting. |
| `Conditions` | Human-readable summary of the rule's conditions, for review. | Reference only -- not read back by any script. |
| `Actions` | Human-readable summary of the rule's actions, for review. | Reference only -- not read back by any script. |
| `TargetFolderEntryID` | Outlook's internal folder identifier for the target folder. | Never hand-edit -- this is a COM-resolved value specific to your mailbox. |
| `Notes` | Free-text notes from OMMigrate about this row (e.g. why a value couldn't be auto-resolved). | Reference only -- safe to clear or ignore. |

> **Blank rows between accounts:** the CSV includes blank separator rows
> between each `RuleStoreName` group, matching the visual grouping you'd
> see in Script 00's own output. This is intentional formatting, not
> stray data -- leave separator rows in place.

---

## Step 3 -- Run Script 01 (Backup)

```powershell
.\Scripts\OMMigrate-01-Backup.ps1
```

**New in 1.5.1:** before backup work starts, Script 01 pauses on every
account marked `MigrationAction=MIGRATE` in `migration_accounts.csv` and
asks you to confirm. Answering **N** automatically sets that account to
`SKIP` in the CSV for you and moves on -- useful if you forgot to update
the CSV for an account you are not ready to touch yet. Answering **Y**
proceeds normally.

Script 01 exports a PST backup for your test account and verifies it.
You will be prompted `[Y/n/EXIT]` before the backup runs.

**Do not proceed to Script 02 until:**
- The Backup Report shows SUCCESS
- A `.pst` file exists in `Documents\OutlookMigration\Backups\`
- The file size looks reasonable for the amount of email in that account

Script 02 will refuse to run without a verified backup. That is by design.

---

## Step 4 -- Run Script 02 (Convert)

```powershell
.\Scripts\OMMigrate-02-Convert.ps1
```

> **Before running:** Confirm Outlook is fully closed. Check Task Manager
> and end any OUTLOOK.EXE process before starting. The script launches
> Outlook via COM and requires exclusive access.

> **If any account is missing a password in the CSV**, Script 02 pauses
> before touching any account and asks you to confirm "I have all required
> passwords and Secure Mail Keys ready" -- this one defaults to **No**, so
> you must actively type `Y` to proceed. It's a deliberate speed bump, not
> a mistake -- if you're not ready, type `N` (or press Enter) to exit
> cleanly and go fill in the missing passwords first.

This is the most significant step. For each account Script 02:

1. Verifies the backup PST exists and is valid
2. Shows old POP3 settings and new IMAP settings side by side
3. Prompts `[Y/n/EXIT]` -- **the default is Y, so pressing Enter proceeds**
   -- type `N` if you want to skip this specific account
4. Guides you through removing the POP3 account in Outlook Account Settings
5. Guides you through adding the new IMAP account via File > Add Account
6. Asks you to confirm the account is visible in the Outlook folder pane

> **Read every prompt before pressing Enter.** This confirmation defaults
> to Yes -- an accidental Enter proceeds with removing POP3 and adding
> IMAP for that account. If you want to review accounts one at a time
> rather than risk moving faster than you intend, don't use `-Force` on
> this script (see OMMigrate_CommandLine_Reference.md).

**Passwords are never handled, stored, or logged by OMMigrate.**
Outlook manages its own password dialog.

The script prints exact removal instructions before you touch Outlook:

![Script instructions for removing the POP3 account](QUICKSTART_images/snap-01.jpg)

In Outlook, this is the Account Settings dialog referenced above:

![Outlook Account Settings dialog, Email tab](QUICKSTART_images/snap-02.jpg)

Outlook will warn you before removing the account -- this is expected,
click **Yes**:

![Outlook warning about removing offline cached content](QUICKSTART_images/snap-03.jpg)

---

### Adding the IMAP Account -- Step by Step

When the script displays the IMAP Account Setup Reference and prompts you
to add the account, follow these steps in Outlook:

![Script instructions for adding the IMAP account](QUICKSTART_images/snap-05.jpg)

1. Click **File** (top left in Outlook)

![Outlook backstage menu with Account Settings and Add Account](QUICKSTART_images/snap-06.jpg)

2. Click **Add Account**

![Outlook Add Account -- email address entry](QUICKSTART_images/snap-07.jpg)

3. Enter your email address
4. Expand **Advanced options** and check **Let me set up my account manually**
5. Click **Connect**
6. Choose **IMAP** from the account type list

![Outlook Advanced setup -- account type selection with IMAP highlighted](QUICKSTART_images/snap-08.jpg)

7. Enter the server details exactly as shown in the Account Setup Reference

![IMAP Account Settings with default ports before entry](QUICKSTART_images/snap-09.jpg)
![IMAP Account Settings with ports set to 993 and 587](QUICKSTART_images/snap-10.jpg)

8. Set encryption: **SSL/TLS** for incoming (port 993), **STARTTLS** for outgoing (port 587)

![IMAP Account Settings with SSL/TLS and STARTTLS selected](QUICKSTART_images/snap-11.jpg)

9. Click **Next** -- Outlook will prompt for your password

![IMAP password entry, empty field](QUICKSTART_images/snap-12.jpg)

10. Enter your password and click **Connect**

![IMAP password entered, masked](QUICKSTART_images/snap-13.jpg)
![Outlook "Adding -- We're getting things ready"](QUICKSTART_images/snap-14.jpg)

> **You may see a credential prompt during this step** -- see
> "Known Outlook 2021 Behavior" below before entering anything here.

![Internet Email credential prompt during account add](QUICKSTART_images/snap-15.jpg)

11. When "Account successfully added" appears -- click **Done**

![Outlook "Account successfully added" screen](QUICKSTART_images/snap-16.jpg)

12. Verify the new account is visible in the Outlook left folder pane

![Outlook folder pane showing Inbox, Outbox, Search Folders](QUICKSTART_images/snap-17.jpg)

13. **Close Outlook** before returning to the script window
14. Type **C** and press Enter to continue

![Script prompt confirming the account is connected and visible](QUICKSTART_images/snap-18.jpg)

> **Do not attempt to send or receive mail after adding the account.**
> Credential corrections are required first -- see below.

---

### Known Outlook 2021 Behavior -- Credential Scrambling

Outlook 2021 Classic has a known limitation: during the Add Account process,
it forces the IMAP username to match the email address and copies that same
value to the SMTP username field. This overwrites any custom credentials,
including AWS SES IAM keys.

**POP3-AWS and POP3-STANDARD accounts** (separate IMAP/SMTP credentials, e.g.
AWS SES outbound): Script 02 attempts to correct this automatically right
after the IMAP account add is confirmed, backing up the registry subkey
first. You'll still need to enter the IMAP password once manually in Outlook
afterward. **This automated path is implemented but has not yet been
verified against a live account of this type** -- if you have one and see
credential problems after Script 02, the manual steps below work as a
fallback regardless of what the automated attempt did: Script 02 always
finishes the run no matter what the automated fix reports (applied, skipped,
or failed), and the manual Repair steps overwrite the same IMAP/SMTP fields
directly, so an incomplete or unsuccessful automated attempt can't block or
complicate correcting it by hand afterward. OMMigrate is aware of this
Microsoft-side bug and makes a best effort to correct it automatically where
the account type allows; the manual path is the safety net for everything
else.

**Every other account type** (including AT&T/Yahoo legacy domains below,
which use a Secure Mail Key and are not eligible for the automated fix):
correct these settings by hand after Script 02 completes -- this is the path
that has been used and confirmed to work:

1. In Outlook -- click **File > Account Settings > Account Settings**
2. Select the migrated IMAP account and click **Repair**
3. Click **More Settings > Advanced** (or the Outgoing Server tab)
4. Correct the **IMAP username** if your server uses a username different from your email address
5. Correct the **SMTP username** -- for AWS SES accounts this is your IAM SMTP access key ID
6. Correct the **SMTP password** -- for AWS SES accounts this is your IAM SMTP secret access key
7. Click **OK** and test the connection

> **AWS SES outbound accounts:** The SMTP username is your IAM SMTP access
> key ID (looks like AKIAIOSFODNN7EXAMPLE). The SMTP password is your IAM
> SMTP secret access key. These are NOT your AWS console login credentials.
> Find both values in `migration_accounts.csv`.

> **AT&T / Yahoo legacy domains** (ameritech.net, sbcglobal.net, etc.):
> Enter your Secure Mail Key (app password) when correcting credentials.
> Not your regular email password. Script 02 opens your AT&T account
> profile automatically when these accounts are detected -- log in and
> navigate to Account Security -> Secure Mail Key -> Manage -> Generate.

---

### Manifest Files -- What to Delete Between Runs

Script 02 writes manifest files to track progress. When re-running Script 02
for additional accounts:

| File | Action |
|---|---|
| `Manifests\Step00_Complete.json` | **Never delete** -- gate check |
| `Manifests\Step01_Complete.json` | **Never delete** -- gate check |
| `Manifests\Step02_Complete.json` | Auto-clears itself once all previously-converted accounts are confirmed IMAP in the CSV. Delete manually only if it doesn't clear on its own. |
| `Manifests\Step02_Checkpoint.json` | Auto-deleted after successful run. Delete manually only if run failed. |
| `Manifests\Step02_SendReceiveState.json` | **Delete if left over from a failed run** |

---

### After Script 02 Completes

Script 02 generates an HTML Migration Report summarizing the conversion:

![Script 02 Migration Report showing migrated, failed, skipped, and warning counts](QUICKSTART_images/snap-19.jpg)

1. Correct IMAP and SMTP credentials as described above
2. Press **F9** in Outlook to trigger a manual Send/Receive and verify mail flows
3. `migration_accounts.csv` is updated automatically -- no manual edits needed

A successful test confirms the account is fully working:

![Microsoft Outlook Test Message confirming account settings](QUICKSTART_images/snap-21.jpg)

> **If other accounts show connection errors after migration:**
> Reboot your mail server and wait 2-3 minutes before retrying.
> This clears stale TCP connections left by the migration process.

---

## Step 5 -- Run Script 03 (Restore)

```powershell
.\Scripts\OMMigrate-03-Restore.ps1
```

Script 03 migrates your folder structure and email content from the
backup PST into the new IMAP account (server folders) and a local Archive
PST (local folders), then updates Outlook Rules to point to new locations.

**Account picker:** A selection window opens showing all eligible accounts
(`MigrationAction=FOLDER-ONLY`). Check each account you want to process
this run and click OK. Use Select All to process all at once. Already
completed accounts are excluded automatically. Cancel exits safely without
making changes.

> **Exchange accounts (`ProviderTag=EXCHANGE-SKIP`):** these never appear
> in the picker above, since there is no folder/content migration for
> them to do. Their Outlook rules are still tracked and updated
> automatically each run (as of 1.5.1) -- if you edit a rule for an
> Exchange account directly in Outlook's Rules Manager, the next Script 03
> run will pick up the change.

You will be prompted to confirm before folders are migrated for each account.

**After Script 03 completes:**
- The Archive PST (`OMMigrate Local Archive`) appears automatically in the
  Outlook folder pane -- no manual attach required
- Each migrated account has a subfolder under OMMigrate Local Archive
  containing its Local-destination folders
- `MigrationAction` is updated to `COMPLETE` in `migration_accounts.csv`
  automatically for each successfully migrated account

---

## Step 6 -- Run Script 04 (Artifacts)

```powershell
.\Scripts\OMMigrate-04-Artifacts.ps1
```

Script 04 migrates Calendar, Contacts, Tasks, Notes, and Journal items --
the non-mail Outlook items Script 03 does not touch. It reads directly from
the live account (not from the backup PST) and applies duplicate detection
before copying anything.

**After Script 04 completes:**
- An Artifacts Report opens in your browser summarizing items migrated and
  any duplicates skipped
- Accounts that were always IMAP (never converted from POP3) report their
  existing artifact counts at SUCCESS -- there is nothing to migrate for
  these accounts, which is expected

---

## Step 7 -- Verify Everything Worked

Before migrating any more accounts, verify your test account thoroughly:

| What to check | Where to look |
|---|---|
| Account shows as IMAP | Outlook -> File -> Account Settings |
| Email is downloading | Outlook folder pane -- new mail appearing |
| Server folders visible | Outlook left pane -- folders marked Server |
| Archive PST attached | Outlook left pane -- OMMigrate Local Archive |
| Historical email present | Browse Archive PST folders |
| Outlook Rules working | Send a test email, verify it routes correctly |
| Calendar/Contacts present | Outlook Calendar and People views |
| Backup PST intact | Documents\OutlookMigration\Backups\ -- file still present |

---

## Step 8 -- Migrate Remaining Accounts

Once your test account is verified:

Re-run Script 00 -- the account selection window will open again. Check the
next batch of accounts you want to migrate and click OK. Then fill in the
`Password` field in the CSV for each newly selected account.

For your already-migrated accounts -- Script 03 will have already set
`MigrationAction` to `COMPLETE` automatically, and Script 02 auto-detects
already-converted accounts and skips them. The Migration Report will show
them in the correct category.

Then re-run Scripts 01 through 04 in order:

```powershell
.\Scripts\OMMigrate-01-Backup.ps1
.\Scripts\OMMigrate-02-Convert.ps1
.\Scripts\OMMigrate-03-Restore.ps1
.\Scripts\OMMigrate-04-Artifacts.ps1
```

You do not need to re-run Script 00 unless you want a fresh Discovery Report.

---

## Optional Utilities

Three standalone scripts live in `Utilities\`, separate from the five
numbered pipeline scripts above -- none of them run automatically, and you
won't need any of them for a normal migration. Run one only if its
specific situation comes up.

These were originally written early in development, when initial POP3
accounts sometimes carried a large number of duplicate rules. Script 00
now deduplicates its own `rules_inventory.csv` automatically, so
duplicates won't show up as duplicate rows in the CSV either way -- but
that dedup only affects the CSV, not live Outlook, so the duplicate rules
themselves are still there until something removes them. Consider running
`OMMigrate-FindDuplicateRules.ps1` as a preventive check on your first
pipeline run, especially for an account with a long POP3/rules history,
or any time rules look doubled up in Rules and Alerts.

- **`OMMigrate-CleanRulesCSV.ps1`** -- if `rules_inventory.csv` ends up with
  duplicate rows for the same rule, this keeps the best row per group and
  removes the rest, after writing a timestamped backup first.
- **`OMMigrate-FindDuplicateRules.ps1`** -- if rules in Rules and Alerts
  look duplicated (this can happen after a `.rwz` rule-file import), this
  read-only scan reports every duplicate it finds. Makes no changes.
- **`OMMigrate-RemoveDuplicateRules.ps1`** -- after reviewing
  `OMMigrate-FindDuplicateRules.ps1`'s report, this deletes the duplicates
  it found. Run Find first.

```powershell
cd "$env:USERPROFILE\Documents\OMMigrate"; Get-ChildItem -Recurse -Include *.ps1,*.psm1 | Unblock-File
.\Utilities\OMMigrate-CleanRulesCSV.ps1 -WhatIf
.\Utilities\OMMigrate-FindDuplicateRules.ps1
.\Utilities\OMMigrate-RemoveDuplicateRules.ps1 -WhatIf
```

All three ask which Outlook profile to work on -- if you only have one
profile it's picked automatically, otherwise you'll be prompted (or pass
`-ProfileName "YourProfileName"` to skip the prompt). All three confirm
with a `Y`/`N`/`EXIT` prompt before doing anything, and `-WhatIf` previews
what would change without changing it. See **README.md, Optional
Utilities** for full details on each.

---

## Exiting Safely

During Scripts 01, 02, and 03:

| Method | Result |
|---|---|
| Type `EXIT` at any `[Y/N/EXIT]` prompt | Clean exit -- progress saved |
| Press `Ctrl+C` | Clean exit -- same as EXIT |

After any exit, re-running the same script resumes from where it left off.
Completed accounts are skipped automatically.

> **After any exit -- graceful or not -- check `RECOVERY.txt`** in
> `Documents\OutlookMigration\` (the top level, not inside a subfolder).
> Every script rewrites this file on exit with a plain-English summary:
> what happened, which accounts completed, which account (if any) was
> interrupted, and exactly what to do next. It's overwritten each time, so
> it always reflects your most recent exit -- read it before assuming
> anything about what state you're in.

---

## If Something Goes Wrong

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

Or simply re-run the installer -- it unblocks everything automatically:
```powershell
.\Install.ps1
```

**Script fails with parse errors**
A text editor corrupted the file during a manual edit. Re-run the installer
to download a clean copy.

**Script 02 -- POP3 removed but IMAP add failed**
Your data is safe. Script 02 shows the exact backup path and recovery steps.
Add the account manually in Outlook, then re-run Script 02 -- it will
detect the account is now IMAP and skip it.

**An account is missing after Script 02**
Open the backup PST: Outlook -> File -> Open & Export -> Open Outlook Data
File -> navigate to `Documents\OutlookMigration\Backups\`. All email is there.

**Other accounts show connection errors after Script 02**
The migration process can leave stale TCP connections on the mail server.
Reboot your mail server and wait 2-3 minutes before retrying. This resolved
all connection errors in testing.

**Send/Receive errors after Script 02**
If Outlook shows Send/Receive errors and groups appear suspended:
1. In Outlook -- File > Options > Advanced > Send/Receive
2. Verify all groups are enabled
3. Press F9 to trigger a manual Send/Receive

**Script 02 FATAL ERROR -- Step02_Checkpoint.json**
`Manifests\Step02_Checkpoint.json` is auto-deleted after a successful run.
If Script 02 crashes with a FATAL ERROR and leaves a stale checkpoint, delete
`Manifests\Step02_Checkpoint.json` before re-running. The checkpoint file
from a failed run can cause the script to skip accounts incorrectly on restart.

---

## Further Reading

- **[README.md](README.md)** -- Full documentation, features, special
  account types, and complete troubleshooting reference
- **[OMMigrate_CommandLine_Reference.md](OMMigrate_CommandLine_Reference.md)** -- Every parameter on every script
- **[OMMigrate_Settings_Reference.md](OMMigrate_Settings_Reference.md)** -- Every settings file key and what it controls
- **[CHANGELOG.md](CHANGELOG.md)** -- Version history and release notes

---

```
Version: 1.5.2
OutlookMailMigrator (OMMigrate) v1.5.2
----------------------------------------------------
Originator & Architect:    Kirk Shallcross - Shallcross Consulting
Implementation Specialist: Anthropic Claude AI
----------------------------------------------------
"Automating the Outlook migration Google suggested couldn't be automated."
----------------------------------------------------
```
