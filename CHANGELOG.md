# Changelog

All notable changes to OutlookMailMigrator (OMMigrate) are documented here.

This project follows [Semantic Versioning](https://semver.org/):
- **MAJOR** version -- incompatible changes requiring operator action
- **MINOR** version -- new functionality, backwards compatible
- **PATCH** version -- bug fixes, backwards compatible

---

> **Current Version: 1.5.2**

---

> **Note on filenames below:** Starting in 1.5.0, `migration_accounts.csv`,
> `folder_map.csv`, `rules_inventory.csv`, and the settings JSON are all
> named with the active Outlook profile appended (e.g.
> `rules_inventory_Outlook.csv`) -- see README.md for details. Entries in
> this changelog use the base filename as shorthand throughout. Entries
> for 1.4.0 and earlier predate this feature and describe the tool's
> actual behavior at the time (a single shared filename, no profile
> suffix).

---

## [1.5.2] -- August 2026

### Consolidated Cleanup and First Public Release

v1.5.2 is a consolidated cleanup release, combining all pre-launch and
post-1.5.0 work into one clean initial public release of OMMigrate. No
functional or behavioral changes from the previously tested code -- this
release presents that work as a single, coherent first release. See the
`[1.5.1]` section below for the full itemized history of the fixes and
changes consolidated into this release.

---

## [1.5.1] -- August 2026

### Post-Release Bugfix Pass

This patch release documents fixes found during live post-release testing
of v1.5.0 -- eight reported issues plus one related functional gap found
during investigation. All fixes are backwards compatible; no CSV schema
or settings changes required.

---

### Install.ps1

**New: Optional Obsidian Add-on support:**
- Added download entries for the free third-party Obsidian (https://obsidian.md)
  `.md` viewer/editor add-on: `Obsidian-Addon-Install.bat`,
  `Obsidian-Addon-Uninstall.bat`, `Obsidian-Addon-DeployLink.reg`,
  `Obsidian-Addon-UninstallLink.reg`, `Obsidian-Addon-Opener.ps1`,
  `Obsidian-Addon-SilentRunner.vbs`
- Added a pre-configured `.obsidian\` vault folder to the download list,
  including the HTML Viewer Plus community plugin, which renders the
  `stripe-button.html` donate embed in README.md when opened in Obsidian
- Entirely optional and additive to the download set -- activation (the
  `.md` file association) still requires the operator to separately run
  `Obsidian-Addon-Install.bat` as Administrator; nothing in this add-on
  affects Scripts 00-04 or the migration pipeline itself
- See README.md, "Third-Party Software Notice," and QUICKSTART.md, Step
  0.75, for full detail

**New: `FilesToDownload` now has full parity with `tree_view.txt`:**
- Added the previously-missing `docs_images/architecture_overview.svg`
  and all 20 `QUICKSTART_images/*.jpg`/`*.png` files -- both README.md
  and QUICKSTART.md embed these images, so a fresh install via the
  `irm | iex` one-liner was previously missing them
- Added `stripe-button.html` (the donate-button embed source referenced
  by README.md)
- Added the repo-maintenance files (`.gitattributes`, `gitignore`,
  `version.txt`, `Update-Version.ps1`, `show_tree.ps1`, `tree_view.txt`)
  for parity with a full `git clone`, even though they have no effect on
  running a migration -- `.git` itself is intentionally excluded, since
  it is git's internal object database, not a distributable file that
  `raw.githubusercontent.com` can serve
- Added `CONTRIBUTING.md` (new file -- see below) to the download list

**`Invoke-UnblockAll` include list extended:**
- Added `*.bat`, `*.reg`, and `*.vbs` alongside the existing
  `*.ps1`,`*.psm1`,`*.bas`,`*.cls`,`*.frm` extensions, so the new
  Obsidian Add-on files are unblocked automatically after download, the
  same as every other script/module file

---

### README.md / QUICKSTART.md

**Documentation updated for the Obsidian Add-on and full file parity:**
- README.md: Directory Structure tree updated to list every new file;
  new "OMMigrate does not install Obsidian" clarification added to the
  Third-Party Software Notice; new callout in Script Parameters
  explaining why the `*.ps1,*.psm1` unblock command intentionally does
  not cover the Obsidian Add-on's `.bat`/`.reg`/`.vbs` files; new
  visible note under the `stripe-button.html` donate embed explaining
  it only renders as a live button in Obsidian with HTML Viewer Plus,
  and pointing non-Obsidian readers to the plain donate link above it
- QUICKSTART.md: new Step 0.75, "Install Obsidian Add-on (Optional),"
  mirroring the existing Step 0.5 VBA-macro pattern -- covers what the
  add-on does, the files involved, the requirement to install Obsidian
  itself first (a separate decision from anything OMMigrate installs),
  and the install/uninstall procedure; matching unblock-scope callout
  added alongside the existing "Always Verify Your Working Directory
  First" section

---

### CONTRIBUTING.md (new file)

**New: contribution guidelines for the public repo:**
- Added `CONTRIBUTING.md` at repo root ahead of going public -- Issues
  (bug reports and questions) are open immediately; pull requests are
  not yet being accepted while real-world testing experience is built
  up across a wider range of environments than the original 3-profile
  test set
- Documents what to include in a bug report (Windows/Outlook/PowerShell
  versions, account type, script number, log excerpts) for faster
  diagnosis

---

### Modules/OMMigrate-Outlook.psm1

**`Export-RulesToCSV` merge-key collision (Part 1/Part 2 rules):**
- Rules split by consolidation into multiple parts (e.g. `(Part 1)`,
  `(Part 2)`, when a folder has more than 5 sender addresses) all target
  the same `TargetFolderPath` -- the prior `RuleStoreName|TargetFolderPath`
  merge key collapsed them onto a single CSV row, silently mislabeling
  Part 1's row with Part 2's `RuleName` on rerun
- Confirmed live: Outlook's own Rules and Alerts UI correctly showed both
  parts; only `rules_inventory.csv` was wrong
- Fixed by adding `RuleName` into the primary merge key on both the build
  and lookup sides

**`Export-RulesToCSV` SendersDomain cross-contamination (Part 1/Part 2 rules):**
- A second, separate issue found during follow-up testing after the fix
  above: `SendersDomain` values were swapping between a Part 1 and Part 2
  row on rerun, even after a manual correction to one of them, despite
  `RuleName` itself already being correct
- Root cause: two places in the merge logic use a leaf-folder-name
  fallback key with no Part-N distinguisher, intended to recover a
  manually-corrected value from an orphaned duplicate row after a rename
  -- since Part 1 and Part 2 share the same leaf folder name, this key
  always matched the sibling Part-N rule instead of a genuine orphan
- Fixed by adding a guard in both locations: a leaf-key match that is
  itself a valid Part-N sibling for the same base folder is skipped;
  genuine orphan-duplicate recovery for renamed, non-Part-N rules is
  unaffected

**`Invoke-DeployConsolidatedRules` RuleName/SendersDomain/Conditions mismatch on re-split (Part 1/Part 2 rules):**
- A third, distinct issue found during further live testing: after
  consolidation re-splits an existing folder group into new chunks (e.g.
  Part 1/Part 2 boundaries shifting between reruns), the CSV write-back
  could assign the wrong chunk's identity to a row -- both prior fixes in
  this same function addressed different collision points, but neither
  covered a row whose original domains ended up split across two new
  chunks after a re-split
- Root cause: consolidation merges every row in a folder group into one
  flat domain pool before re-splitting into chunks of 5 -- a row's own
  pre-run `SendersDomain` can no longer reliably identify which new
  chunk it belongs to, since domain/row order is not preserved across a
  re-split
- Fixed by distributing each run's actual chunks across the CSV rows in
  the same folder group directly, writing each chunk's authoritative
  `RuleName`, `SendersDomain`, and `Conditions` -- sourced from the live
  rule Outlook actually created, not reconstructed from stale pre-run
  row data

**New: `SendersDomain` now synced from each Outlook rule's own SenderAddress condition:**
- Previously `SendersDomain` was only ever a best guess (the last folder
  name segment of `TargetFolderPath`), left untouched once set -- but a
  manually-added Outlook rule (via the UI Rules Manager) commonly has a
  real SenderAddress condition that is a far more reliable source
- Added `Get-RuleSenderAddressWords` (new helper): reads a rule's
  SenderAddress condition words directly, independent of the existing
  `Get-RuleConditionsSummary` display-string output
- `Get-OutlookRules` now populates `SendersDomain` from this helper when
  building each rule object, instead of always leaving it blank
- `Export-RulesToCSV`'s merge logic now overwrites `SendersDomain` with
  the live rule's SenderAddress words on every run, whenever that
  condition is present -- including over a prior manual CSV edit; falls
  back to the existing `TargetFolderPath`-based best guess, unchanged,
  when the rule has no SenderAddress condition of its own
- The live Outlook rule is the source of truth going forward: a manual
  CSV edit to `SendersDomain` is only preserved once it has been pushed
  to the live rule via Script 03 -- running Script 00 again first will
  overwrite the edit with whatever the (not-yet-updated) rule still says.
  See README.md, Runtime Output Files, for the full explanation
- Script 03 was checked for the same requirement in the reverse
  direction and needs no changes -- it already only ever reads
  `SendersDomain` from the CSV to write into the live rule's condition,
  never the other way around

**`Close-PSTFile` reliability:**
- Now retries up to 3 total attempts (750ms apart) before giving up,
  instead of failing permanently on the first transient COM error
- A prior single-attempt failure could leave a PST attached in Outlook,
  requiring the operator to close it manually

---

### Scripts/OMMigrate-01-Backup.ps1

**New: MigrationAction confirmation gate (Step 2):**
- Every account with `MigrationAction=MIGRATE` now pauses for an explicit
  Y/N confirmation before backup work begins
- Declining automatically sets `MigrationAction=SKIP` for that account in
  `migration_accounts.csv` (via `Update-AccountMigrationAction`) and
  excludes it from this run's Archive pre-build -- no manual CSV edit
  required
- Fixes a real failure mode: an account left at `MIGRATE` that the
  operator did not intend to touch yet previously flowed straight into
  backup and Archive pre-build, producing a cascade of COM failures

**Archive pre-build folder-walk hardening:**
- Added a consecutive-failure guard (5 failures in a row) to the
  first-run folder walk -- stops that account's walk cleanly with a
  clear diagnostic instead of retrying through the entire remaining
  folder queue
- A single isolated failure does not trip the guard, only a sustained run

**`Close-PSTFile` failure now surfaced to the operator:**
- All three call sites in this script now check the return value instead
  of discarding it -- if the automated close fails after retries, the
  console prints a clear warning telling the operator to close the PST
  manually in Outlook

---

### Scripts/OMMigrate-03-Restore.ps1

**Phase 3 "No live rules collections found" -- WARN downgraded to INFO:**
- Added a second recognized harmless case: the default/primary store was
  simply not selected in the account picker this run (existing 2026-06-26
  fix only covered the all-selected-are-secondary-stores case)
- Added `.Trim()` to the secondary-store name comparison, matching the
  established pattern already used elsewhere in this file, to reduce
  false WARNs from incidental whitespace differences

**Exchange accounts now included in `LastTargetRun` tracking:**
- Exchange accounts (`ProviderTag` `EXCHANGE-*`) were previously fully
  excluded from Phase 3's folder-target remap loop, since they correctly
  never enter the conversion-eligibility list used elsewhere in this
  script -- but this also meant their rules' `LastTargetRun` was never
  stamped, even after a rule was edited directly in Outlook's Rules
  Manager UI
- Fixed via a second, independent eligibility list scoped only to Phase 3
  -- does not affect the account picker, content-migration eligibility,
  or any other script in the pipeline

**`LastTargetRun` never stamping for a rule created fresh by consolidation:**
- Phase 3's check for "was this rule already handled by consolidation
  this run" only recognized rules absorbed from an existing live rule --
  a rule created fresh with nothing to absorb (e.g. a brand-new Part 1/
  Part 2 split with no prior live rule) was never recognized, silently
  deferring `LastTargetRun` to a second run even though the design
  intent (2026-07-07 `LastDeployedRun`/`LastTargetRun` split) was for
  both to complete in the same run whenever both start blank
- Fixed by checking `ConsolidatedRuleTargets` (populated for every rule
  consolidation creates, absorbed or fresh) ahead of the absorption-only
  check

**Phase 3 not processing a still-POP3 account's rules outside the picker selection:**
- A POP3 account not yet converted to IMAP had its rules skipped by
  Phase 3 whenever it wasn't selected in the current run's migration
  picker, even though its rules may have just been updated by
  consolidation in the same run
- Mirrors the existing Exchange-accounts fix exactly, scoped by
  `ProviderTag` (`POP3-*`) rather than broadening to every account --
  rule maintenance is independent of migration-selection state for an
  account still in a POP3 (or Exchange) state, but does not affect
  already-converted or already-IMAP accounts, or picker scoping for any
  other account

---

### Pre-Launch Sanitization and Documentation Pass

Full sweep of the codebase and docs ahead of public launch -- no
functional/behavioral changes; documentation, in-code comments, and
console/log text only.

**Sanitization (code and comments, all scripts and modules):**
- Removed all personal email addresses, personal domains, and personal
  Outlook profile names from source code -- comments, `.NOTES` blocks,
  and any string literal that isn't user-facing console/log output
- Personal Outlook profile names (previously two different real personal
  profiles referenced inconsistently across comments) consolidated to a
  single generic placeholder, `TestProfile`, throughout
- Two example diagnostic-script filenames in comments, built from one of
  the personal profile names, replaced with plain descriptions instead
  of a fabricated filename
- Two example bare-word tokens in a regex-behavior comment (illustrating
  non-email match patterns), tied to personal business names, replaced
  with generic equivalents
- User-facing console message in Script 00's `-PatchRulesCSV` path
  (pointed at a specific personal profile name) genericized to
  `<Your Profile>`
- Confirmed as out of scope (left unchanged, correctly): generic
  AT&T/Yahoo legacy provider-domain references (`ameritech.net`,
  `sbcglobal.net`, etc.) -- these describe a real, generic account
  category, not personal identification. Real name/credit block
  references in `.NOTES` headers, `## Credits` sections, `LICENSE.md`,
  and `CONTRIBUTING.md` also confirmed out of scope, per this project's
  standing rule that real names are permitted only in credits/copyright
  contexts

**README.md / QUICKSTART.md -- new documentation:**
- New pre-install decision-point warning in README ("Read This Before
  You Decide to Use OMMigrate") -- explains that OMMigrate restructures
  Outlook Rules to a new server/local-Archive filing model rather than
  preserving the original POP3 rule structure, that this is a one-way
  decision once rules deploy, and specifically flags this for
  consultants/admins evaluating the tool for clients or compliance/
  server-retention requirements. QUICKSTART links to it up front, before
  Step 0
- New "Before You Begin: Back Up Your Existing Data" section (README)
  and matching "Back Up Your Existing Outlook Data First" section
  (QUICKSTART) -- prerequisite PST backup and per-account Outlook rules
  export steps, a postrequisite second rules export framed as an ongoing
  point-in-time recovery practice (not a one-time step), guidance to
  keep both outside any OMMigrate-managed folder, and a liability
  disclaimer covering data loss when these steps are skipped
- PST backup guidance expanded to cover the two real Windows default
  storage locations (`Documents\Outlook Files\` and
  `AppData\Local\Microsoft\Outlook\`, the latter hidden by default)
  and a fallback note for PSTs an admin originally pointed at a custom
  folder, in addition to the existing Data Files-tab lookup method
- New `rules_inventory.csv` column-by-column reference (QUICKSTART, with
  a summary cross-reference in README) -- documents how `SendersDomain`
  and `TargetFolderPath` relate (the former is the actual deployed rule
  condition; the latter is just the filing destination), that a
  first-run `TargetFolderPath` is normally an unverified best-guess seed
  value that needs manual correction, and the space-separated
  `SendersDomain` multi-value syntax
- Script 00's existing `SendersDomain` best-guess console warning
  (yellow, `ACTION REQUESTED`) updated to reference the profile-suffixed
  filename in its instructional comments; console/log output left
  referencing the plain filename, since that text already resolves and
  displays the real, profile-suffixed path at runtime

**OMMigrate_Settings_Reference.md:**
- New header disclaimer, shown before all sections, clarifying that any
  key marked "Not yet wired up" has no effect on current behavior and is
  reserved for a future release -- prevents a reader from assuming a
  listed-but-unimplemented setting is already functional

---

### Obsidian Add-on -- Antivirus False Positive Fix

Norton (Behavioral Protection) flagged `Obsidian-Addon-Opener.ps1` as
`IDP.Generic` when triggered via the `.md` file association -- a false
positive, confirmed clean via manual code review and a full-folder scan
(636 items, 0 threats). Root cause isolated through a systematic
elimination process rather than assumed:

- Ruled out the vault-found branch's `Start-Process`-based Obsidian
  launch (clean on every test)
- Ruled out file content (a plain-text `.md` with no markdown syntax
  still triggered the flag)
- Ruled out the Notepad-fallback branch's process-launch syntax --
  rewriting `& "notepad.exe" $FilePath` as `Start-Process -FilePath ...
  -ArgumentList ...` did not resolve it
- Confirmed via a diagnostic build (fallback branch replaced with an
  inert log write, no process launched at all) that the flag stopped
  firing entirely with zero process launches, proving the trigger was a
  process-launch pattern somewhere in the chain -- not this script's
  internal logic
- Traced upstream to `Obsidian-Addon-SilentRunner.vbs`, the `.md`
  file-association handler that launches this script: its
  `-WindowStyle Hidden -ExecutionPolicy Bypass` PowerShell invocation is
  a well-documented AV heuristic signature (hidden-window PowerShell
  launched by a script host, with the execution policy explicitly
  bypassed) -- independent of what the launched script actually does

**Fix:** removed `-ExecutionPolicy Bypass` from
`Obsidian-Addon-SilentRunner.vbs`'s PowerShell invocation, keeping
`-WindowStyle Hidden` (no visible console flash). This relies on the
user's execution policy already being `RemoteSigned`, which QUICKSTART.md
Step 0 already directs every user to set during install -- confirmed
live: no further Norton flag on repeated `.md` double-clicks, both inside
and outside an Obsidian vault.

**`Obsidian-Addon-Opener.ps1` retained two related improvements** made
during the investigation, kept as genuine correctness fixes independent
of the AV root cause:
- The Obsidian launch URI now properly URL-encodes the vault name and
  relative file path via `[System.Uri]::EscapeDataString()`, instead of
  raw string interpolation -- handles vault/file names containing spaces
  or special characters correctly
- Both the vault-found and non-vault fallback paths now use
  `Start-Process -FilePath ... -ArgumentList ...` with explicit error
  handling, instead of the bare `&` call operator with no fallback if
  Notepad fails to launch

---

### OMMigrate-Outlook.psm1 -- Rule Consolidation Fixes

Three issues found and fixed during live-testing of Script 03's rule
consolidation pipeline (`Invoke-DeployConsolidatedRules` /
`Invoke-BuildRulesFromMap`), plus one confirmed non-code finding about
Outlook COM behavior under sustained testing:

**Fix: `Add-RulesCsvSeparatorRows` missing from module exports:**
- The shared CSV blank-separator-row helper (added earlier this release)
  was defined in the module but never added to the `Export-ModuleMember`
  list, so any script importing the module via `Import-Module` (not
  dot-sourcing) could not call it -- caused a non-fatal
  `LastTargetRun` write failure with a "term ... is not recognized" error
  on every affected run
- Fixed by adding the function to the module's export list

**Fix: consolidated rule name never written back to CSV on first-time
consolidation:**
- A `rules_inventory.csv` row that had never been consolidated before
  (its `RuleName` was still the raw pre-migration Outlook rule name, with
  no `Rule: [account] ...` bracket) was grouped into an account-less
  bucket during the CSV rename pass, but the brand-new consolidated rule
  name Outlook actually received always has an account bracket -- so an
  account-less old row could never match a bracketed new name, and
  `RuleName` in the CSV silently stayed at its pre-consolidation value
  forever, even though the live Outlook rule was correctly renamed
- This same mismatch also prevented `LastTargetRun` from being stamped
  in the same run, since Script 03's Phase 3 matches CSV rows against
  live rules by name
- Fixed by falling back to the row's own `RuleStoreName` (the same value
  the new rule's account bracket is built from) when no bracket exists
  to parse

**Fix: `Conditions` CSV column dropped preserved rule conditions after
consolidation:**
- When an existing rule with a non-sender-address condition (e.g.
  `Subject contains: X`) was absorbed into a new standardized
  consolidated rule, Outlook correctly preserved that original condition
  alongside the new sender-address condition (both required, joined with
  AND) -- but the CSV write-back for the `Conditions` column hardcoded a
  `Sender address: ...`-only string, silently dropping the preserved
  condition from the CSV even though the live rule was correct. The CSV
  would only show the accurate combined text after a subsequent Script 00
  rescan, which could be days or weeks later for accounts not routinely
  re-discovered
- Fixed by capturing the true, complete conditions summary directly from
  the newly-created live rule object via the existing
  `Get-RuleConditionsSummary` function (the same one `Get-OutlookRules`
  already uses for a normal Script 00 rescan), immediately after every
  condition has been applied to it, and using that captured value for the
  CSV write-back instead of the hardcoded string

**Finding (not a code fix): Outlook COM/rules-store staleness under
sustained repeated testing:**
- During an extended live-testing session (many consecutive Script 03
  runs against the same Outlook profile within one Windows session),
  rule consolidation began reporting success in the run log (absorb,
  create, save, all with no errors or warnings) while the live Outlook
  Rules and Alerts UI -- confirmed via a fully closed and reopened
  Outlook -- continued showing the old, pre-consolidation rule
  unchanged
- Investigated as a possible code defect first: reverting the two CSV
  fixes above did not resolve it, ruling out this release's own changes
  as the cause. A rule-collection COM object identity/lifetime issue of
  the same general class was found and fixed once before in this
  codebase (2026-07-12, unregistered RCW garbage-collected mid-function)
  and was suspected again, but the specific object involved this time
  was already correctly registered
- Resolved entirely by rebooting the test machine -- a clean Windows
  session with the exact same test scenario worked correctly on the
  first attempt post-reboot, and continued working correctly across
  several subsequent runs. This points to Outlook/COM-level resource or
  memory-heap exhaustion accumulating over many rules-COM connect/
  disconnect cycles in one long-running Windows session, not a defect in
  OMMigrate's own code
- See README.md and QUICKSTART.md for new guidance on this -- an
  unexpected warning, an unusually slow run, or inconsistent results
  after a long testing/working session is now documented as a sign to
  reboot before investigating further as a bug

---

### Utilities/OMMigrate-CleanRulesCSV.ps1, Utilities/OMMigrate-FindDuplicateRules.ps1, Utilities/OMMigrate-RemoveDuplicateRules.ps1

**New: `-ProfileName` parameter, profile-aware by default:**
- All three standalone Utilities scripts now accept an optional
  `-ProfileName` parameter. When omitted, the real Outlook profiles on the
  machine are read live from the registry (the same lookup Script 00
  uses) -- one profile found is used automatically, multiple profiles
  prompt for a name, and zero found is a clear error
- A typed or passed-in profile name is now matched case-insensitively
  against the real registered profiles and normalized to their actual
  casing (e.g. typing `restorerules` now correctly resolves to
  `RestoreRules`) rather than being used verbatim, which could otherwise
  produce a wrong-cased CSV filename that only worked by coincidence on
  case-insensitive filesystems
- None of the three scripts read or write `OMMigrate_Settings.json` --
  the resolved profile is held in memory for that run only, which lets
  the existing profile-suffix filename logic resolve correctly without
  any settings file involved
- All three now confirm with a `Y`/`N`/`EXIT` prompt before doing
  anything, so a lone auto-selected profile still gives the operator a
  clear point to back out before any work happens

**Fixed: Utilities scripts could not run from `Utilities\`:**
- All three scripts located their `Modules\` folder relative to their own
  script path assuming a `Scripts\`-sibling layout; since these scripts
  actually live in `Utilities\` (a separate sibling of `Modules\`, not
  nested under it), running any of them threw "OMMigrate Modules folder
  not found." Corrected to the actual one-level-up relative path

**Fixed: `OMMigrate-CleanRulesCSV.ps1` dropped blank separator rows on write:**
- "Total rows" in the summary output previously counted every row in the
  CSV including blank `RuleStoreName`-group separator rows, inflating the
  count against "Rows remaining" (which only ever counted real data
  rows). More seriously, those separator rows were never written back to
  the output file at all -- a real (non-`-WhatIf`) run would have
  silently stripped every separator row from `rules_inventory_<Profile>.csv`,
  the same layout `Export-RulesToCSV`/Script 03 maintain via
  `Add-RulesCsvSeparatorRows`. Fixed on both counts: the row count now
  excludes separators, and the final write reuses the same
  `Add-RulesCsvSeparatorRows` helper to restore them before saving

**New: `OMMigrate-FindDuplicateRules.ps1` output renamed and profile-suffixed:**
- `Reports\DuplicateRules.csv` is now `Reports\DuplicateRules_<Profile>.csv`,
  matching the profile-suffix convention used elsewhere. `OMMigrate-RemoveDuplicateRules.ps1`
  reads the same profile-suffixed filename
- When duplicates are found, a second, timestamped copy of the report is
  also written (`DuplicateRules_<Profile>_<yyyyMMdd_HHmmss>.csv`) purely
  for the operator's own record-keeping -- never read by any script, and
  never overwritten by a later run
- `OMMigrate-RemoveDuplicateRules.ps1`'s "file not found" message now
  explains both real causes -- the scan hasn't been run yet for that
  profile, or it was run and genuinely found no duplicates (in which case
  no file is written) -- rather than implying a step was skipped

---

### Install.ps1 (second pass)

**Fixed: false-positive download failures on legitimately small/empty files:**
- `Get-GitHubFile`'s download-succeeded check rejected any file under 100
  bytes as "too small," which incorrectly failed several real, correctly-tracked
  repo files: `.gitattributes`, `.gitignore`, `version.txt`,
  `.obsidian\community-plugins.json`, and `.obsidian\plugins\html-viewer-plus\debug.log`
  (intentionally 0 bytes by default). The check now only rejects a genuinely
  empty (0-byte) download
- `.obsidian\plugins\html-viewer-plus\debug.log` is a real 0-byte file in the
  repo, so even the corrected 0-byte check still flagged it. Added
  `$Script:ExpectedEmptyFiles`, a small whitelist of filenames the empty-download
  check skips entirely -- currently just `debug.log`
- Removed three stale `FilesToDownload` entries -- `.obsidian\app.json`,
  `.obsidian\appearance.json`, `.obsidian\workspace.json` -- that `.gitignore`
  deliberately excludes from the repo (per-user Obsidian state, Administrator's
  direction 2026-08-18) and that were 404ing on every install
- Live-confirmed via `.\Install.ps1 -Force`: 76 file-list errors (7 failing
  downloads) reduced to 0; installer now reports "Install Complete -- All
  files downloaded successfully"

### README.md, QUICKSTART.md

**Fixed: confusing "Option B -- Manual download" install instructions:**
- Option B never stated where `Install.ps1` needed to be before running the
  three listed commands, so a user following it literally from a fresh
  machine (e.g. running from `Documents\Downloads` default location) hit
  `Unblock-File : Cannot find path` -- Install.ps1 creates
  `Documents\OMMigrate\` itself on first run, so there was no existing
  correct folder for the instructions to reference
- Reworded to frame Option B as a fallback for when Option A's `irm | iex`
  one-liner can't be used (blocked by network policy, or the operator wants
  to review the script first) rather than an equally-default alternative,
  since Option A already handles folder creation automatically end to end
- Added the missing setup step -- move the downloaded `Install.ps1` into
  `Documents\OMMigrate\` (creating the folder if needed) and `cd` into it --
  before the existing `Set-ExecutionPolicy` / `Unblock-File` / `.\Install.ps1`
  commands

---

## [1.5.0] -- July 2026

### Multi-Archive TargetStoreName Support

This release closes out the multi-archive feature (opened under a feature
freeze on 2026-07-09) -- OMMigrate no longer assumes a single Archive PST.
Operators can attach multiple archive PSTs and map individual accounts to
whichever archive they belong in.

**New: `ArchiveStoreMappings` settings schema (`Core.psm1`):**
- `Save-OMMigrateArchiveStoreMappings` persists the operator's
  account-to-archive mapping
- `MasterArchiveNames` setting lists archive PSTs that must never be
  auto-detached by any script
- Per-profile settings (`Switch-OMMigrateProfileSettings` /
  `Sync-OMMigrateProfileSettings`) -- each Outlook profile gets its own
  settings file, eliminating cross-profile mapping collisions

**New: TargetStoreName picker (Script 00):**
- WinForms picker -- one panel per attached archive PST, checkbox list of
  accounts per panel, pre-populates from any prior saved mapping
- Select All / Unselect All buttons per archive panel
- Enforces one-PST-per-account on save
- Unmapped accounts default to the shared/default Archive PST for
  Local-destination rules

**Resolution wired through the full pipeline:**
- `Export-RulesToCSV` fills blank `TargetStoreName` via the mapping,
  falling back to the original default when unmapped
- `Invoke-DeployConsolidatedRules` and `Invoke-RulesRecreation` resolve
  each rule/group's own `TargetStoreName` to the correct attached PST
  rather than a single hardcoded archive
- `Export-FolderMapCSV` resolves `StoreName` via the mapping instead of
  whichever store a folder was physically found in (including the
  bare account-root row)
- Script 01 Archive pre-build groups accounts by resolved
  `TargetStoreName` and builds into each distinct archive PST, instead
  of always building into one hardcoded archive
- Script 03's Strategy 1/2 folder-target remap and `Invoke-AccountFolderMigration`
  resolve each account's own `TargetStoreName` instead of defaulting to
  the master archive

**`Module3.bas` (VBA macro) full rewrite for parity:**
- Removes all hardcodes (archive name, CSV path, provider-name fallback
  substring)
- Reads `OMMigrate_Settings_<Profile>.json` at runtime via a hand-rolled
  JSON extractor; resolves each account's mapping with fallback to
  `MasterArchiveNames`
- Rules CSV and settings JSON both read/written as UTF-8 (`ADODB.Stream`),
  fixing prior ASCII-only encoding
- Success message now reports domains processed, new rules created, and
  duplicates removed as three separate figures instead of one conflated
  total

**Bugs found and fixed during the live-test/promotion cycle:**
- Merge-key instability: rules whose folder target was COM-inaccessible
  had an unstable best-guess `TargetFolderPath` that broke the CSV merge
  key on every rerun -- replaced with a three-tier fallback key
  (`TargetFolderPath`, then `RuleName`, then leaf folder segment), plus
  leaf-based duplicate detection for orphan rows
- Rule-rename tracking: consolidation renames a rule from its raw UI name
  to a standardized form; a `RuleName`-only merge key orphaned rows that
  already had `LastDeployedRun`/`LastTargetRun` stamped, creating a fresh
  blank row every run. Merge key changed to
  `RuleStoreName|TargetFolderPath`; renamed/absorbed rules now write their
  new name back into the CSV in the same run
- `Invoke-AccountFolderMigration` Phase 3 structural bug: the folder-path
  match branch set `$csvRow` but never entered the body containing the
  actual remap logic, so `LastTargetRun` was never stamped even on a
  correct match -- fixed, and `rules_inventory.csv` is now reloaded after
  consolidation so Phase 3 sees the current run's own writes
- Rules Updated counter undercounted the Strategy-1-skip case (a rule
  consolidated and already correctly targeted in the same run); added a
  double-count adjustment for the case where `LastDeployedRun` and
  `LastTargetRun` are both blank on the same rule
- Archive PST attach/detach leaks: Script 01 was force-closing manually
  re-attached backup PSTs on normal successful runs (missing an
  already-mounted guard); Script 03 opened the Archive PST via
  `Open-PSTFile` but never detached it anywhere -- both fixed with
  `Test-PSTAlreadyMounted` checks and matching `Close-PSTFile` calls that
  only detach what the current run itself mounted
- Backup PST filename profile-suffix regression and revert: an earlier
  fix suffixed the plain backup PST filename by profile, but this file is
  tied to the account, not the profile, and must never be
  profile-suffixed (unlike the CSVs and the OST-export backup, which
  genuinely differ per profile) -- reverted across all call sites in
  Scripts 01 and 03

---

### Rules Engine -- Consolidated Rule Engine, Condition Preservation, Performance

**New: consolidated rule engine (`Invoke-DeployConsolidatedRules` /
`Module3.bas` -- Gemini parity):**
- Replaces all prior inline purge/recreate logic (Match/Additive/Full
  sync modes, byte-level stream patching, default-store guards) with a
  single code path that runs identically for every account, default and
  secondary, with no exceptions
- Groups pending `rules_inventory.csv` rows by
  `RuleStoreName|TargetFolderPath`, chunks sender domains 5-at-a-time into
  consolidated rules, writes `LastDeployedRun` timestamps so the
  PowerShell module and the VBA macro stay idempotent and interchangeable
  against the same CSV
- Proven build order (5 rounds of live testing): StopProcessing first via
  named-property reflection, then OnLocalMachine, then Account, then
  SenderAddress (explicit array cast required), then MoveToFolder last,
  then a single `Save()` -- order matters, since assigning MoveToFolder to
  a cross-store destination locks the rule against further edits
- 1000-action watchdog and domain-level dedup guard, matched in both
  languages
- `-ScopedAccountNames` parameter makes picker selection authoritative
  for all rule-deployment and resort activity, not just what gets
  reported -- fixes a bug where picking one account still silently
  deployed and resorted six other accounts' pending rows

**New: UI Rule Manager condition preservation:**
- When an existing rule targeting the same folder is absorbed during
  consolidation, any manually-added UI condition is harvested before
  deletion and reapplied to the new standardized rule: Subject, Body,
  BodyOrSubject, MessageHeader, RecipientAddress, Category, Importance,
  Sensitivity, HasAttachment, Cc, OnlyToMe, ToOrCc -- live-verified for
  all 12 types in both the PowerShell module and the VBA macro
- Two real bugs fixed during live verification: array-typed COM
  properties (Subject, Body, BodyOrSubject, MessageHeader,
  RecipientAddress) were being written as scalars and throwing a COM type
  mismatch; the Category condition was being read from the wrong COM
  property (`AnyCategory.Categories` instead of `Category`)
- "From people or public group" and "through the specified account" are
  deliberately not preserved -- the former is already replaced by
  SenderAddress, the latter destabilizes the COM API (see Account
  condition removal, below)

**Account condition removed as root cause of `0x800C8101`:**
- Confirmed via controlled A/B testing that the Account condition
  ("through the specified account") triggers `0x800C8101` COM read
  failures on Outlook 2021 Classic, even when it correctly resolves to
  the default store's own account -- a legacy POP3-era requirement, not
  needed for IMAP
- `Set-RuleConditions`'s Account-condition write step removed entirely;
  Script 03 gained an unconditional strip pass (later gated to rows with
  a blank `LastDeployedRun`, so already-processed rows aren't re-checked
  every run) that clears the condition from any rule that still carries
  it, without deleting the rule
- Live-verified at full production scale: 361-rule rebuild, 8
  previously-stuck rules recreated cleanly, 0 warnings, 0 errors

**StopProcessing accessor hardened:**
- Extensive diagnostics (six isolated test scripts) proved
  `StopProcessingRules` cannot be enabled via COM on a rule freshly
  created via `IRulesCollection.Create()` in the same session, by any
  method or ordering -- it can only be set via the native Outlook UI or
  inherited from `.rwz` import
- The `Actions.Item(27)` fixed-index accessor used as a workaround was
  replaced project-wide with the named property (`Actions.Stop`) via
  `InvokeMember` reflection, confirmed safe post-Save; also fixed a real
  bug where `$rule.StopProcessing` (not a valid Outlook Object Model
  property) was silently swallowed by an empty catch block, always
  returning a hardcoded default

**PR_RW_RULES_STREAM persistence (large rule sets):**
- `Invoke-PurgeAndRecreateRules` patches bytes 44-45 (the rule count
  field) via `PropertyAccessor.SetProperty` + `item.Save()` before
  calling `IRulesCollection.Save()` -- resolves rules vanishing after an
  Outlook restart on large rule sets (500+)
- Secondary stores with a 100-byte stub `IPM.RuleOrganizer` stream (no
  prior rules) get a temporary seed rule to initialize the stream before
  real rules are created, deleted before the final `Save()`
- Secondary-store idempotency (Match/Additive/Full sync modes, keyed off
  a new `LastDeployedRun` CSV column) preserves any StopProcessing
  condition the operator manually set in Rules Manager between runs --
  later fully superseded by the consolidated rule engine above

**Performance -- rule rebuild time cut from ~1.5 hours to production-viable:**
- Root cause: PowerShell's `InvokeMember`/.NET reflection COM interop
  overhead per call, not a logic difference from the VBA macro's native
  COM dispatch (~30 seconds for the same rebuild)
- Eliminated the O(n^2) consolidation scan -- previously iterated the
  entire live Rules collection for every group key being processed;
  replaced with a `$folderPathToRuleNames` lookup built once from
  `rules_inventory.csv`, zero COM calls for the common case
- Batched `Save()` support added (`-SaveBatchSize`) but confirmed via live
  testing not to be the dominant cost on its own
- Redundant delete-before-create dedup guards disabled in both languages
  once confirmed the consolidation-scan rewrite already covers the same
  case

**Rules discovery and CSV integrity fixes:**
- `Get-OutlookRules` now scans every account's own store (previously
  default store only); dedup and merge keys changed to
  `RuleStoreName|RuleName` so two accounts with identically-named rules
  no longer collapse into one CSV row
- `SendersDomain` normalization rewritten to tokenize first and validate
  each token as a full email or bare domain fragment, instead of
  unconditionally stripping `@`/`&` from the raw value (was corrupting
  valid email addresses)
- `TargetStoreName`, `Conditions`, and `Actions` columns added to
  `Export-RulesToCSV`'s merge-preserve block so a rerun no longer
  silently overwrites operator edits or regresses genuine content to a
  known read-failure placeholder
- `Account.DeliveryStore` resolution failures (confirmed permanent for a
  known group of accounts, not a timing race) now fall back to a
  DisplayName match against `namespace.Stores` instead of silently
  skipping the account's rules
- Manually-added rules (rules not matching the consolidated-rule naming
  convention) are now surfaced in an info alert on the Discovery report
  for audit visibility

---

### Script 01 -- Mail-Only Scope, Multi-Archive Pre-Build

- **Calendar/Contacts duplication fixed:** `Copy-OSTFolderRecursive` now
  skips Calendar, Contacts, Tasks, Notes, and Journal folders (matched by
  `DefaultItemType`). Root cause: `Item.Copy().Move()` over COM was
  duplicating items in the *live* source folder on every run, due to
  Outlook COM's restrictions on moving individual occurrences of a
  recurring appointment series. Script 01's backup is now strictly
  mail-only; Script 04 remains sole owner of artifact migration
- New `_Dev` utility recovers accounts affected by the pre-fix
  duplication bug (live-tested: 1,755 duplicate Calendar items and 405
  duplicate Contacts found and removed, verified clean on rerun)
- Archive pre-build now groups accounts by resolved `TargetStoreName`
  (see multi-archive section above)

---

### Script 03 -- Folder Migration and Rules Update

- Rules deployment fully migrated to the consolidated rule engine (see
  Rules Engine section above) -- Phase 4a (the old alphabetical
  sort/renumber-via-ExecutionOrder block) removed as obsolete
- New `Invoke-ResortRulesByLabel` final pass repositions the full rule
  collection into alphabetical order via `ExecutionOrder` writes only,
  never `Create()`/`Remove()` -- preserves any manual customization an
  operator made via the Outlook UI
- Phase 3 (folder-target remap) extended to secondary stores, previously
  excluded based on a disproven assumption from earlier `0x800C8101`
  investigation; gained its own `LastTargetRun` idempotency column,
  independent from `LastDeployedRun`
- Two independent, compounding causes of a deleted rule reappearing after
  Outlook closes fixed: `ActiveExplorer().Activate()` was caching a stale
  rules snapshot that got flushed back on Quit(), overwriting the
  script's own deletion; two leaked COM references
  (`$dupRules`/`$misroutedRules`) were being reassigned inside a
  per-store loop and only the last store's reference was ever released
- Centralized Outlook auto-close into a single shared routine
  (`Close-OutlookIfRunning`), replacing five independently reinvented
  inline versions across Scripts 00-04; pre-flight auto-close in Scripts
  02/03/04 reordered to fire only after the operator's Y/N confirmation,
  never before
- Fixed a `$activeNamespace` typo (undefined variable, should have been
  `$namespace`) that was silently failing the default store's rules update
  on every run and getting mislabeled as an expected condition by a broad
  catch block

---

### Script 04 -- Artifacts

- Always-IMAP accounts (never converted from POP3, structurally have no
  backup PST) no longer report a false WARN -- now checks the live IMAP
  store directly and reports actual item counts at INFO/SUCCESS

---

### Directory Structure and Housekeeping

- New `Outlook VBA Macros` directory (later renamed `OutlookVBAMacros`)
  holds all 9 project VBA files: `Module1.bas`, `Module2.bas`,
  `Module3.bas`, `Module7.bas`, `FrmArchivePicker.frm`/`.frx`,
  `FrmProgress.frm`/`.frx`, `ThisOutlookSession.cls`
- New `Module7.bas` -- `CorrectArchiveFolders`, a CSV-driven archive
  misroute correction macro with picker-based store selection (see
  in-code documentation for the three-pass sender/recipient matching
  logic); replaces the earlier `Module6.bas`
  (`DeDuplicateAndRedistribute`) design, which is removed outright after
  it was found sweeping a manually-attached backup PST that was never
  meant to be touched
- Retired/deprecated macros removed; personal-environment-specific
  functions cleansed from `ThisOutlookSession.cls`
- WIP files superseded by their promoted production versions removed
- 74 files bumped to version 1.5.0; 3 stray test binary files removed

---

## [1.4.0] -- June 2026

### Script 03 -- Rules Engine Overhaul (major)

This release covers the full rules-engine development arc from the
`.rwz`-import/COM-enumeration investigation through the first working
consolidated approach, developed in close collaboration with Gemini (AI)
on the underlying COM/MAPI mechanics.

**Root cause investigation -- `0x800C8101` ("devil code") at scale:**
- Confirmed via extensive diagnostic testing (dozens of isolated `_Dev`
  scripts) that Outlook's `IRulesCollection.Save()` is an atomic
  transaction: every rule in the collection must have a valid folder
  target or `Save()` fails for the entire collection
- Confirmed the Account condition carried over by `.rwz` import (the
  condition's Enabled flag survives import but the underlying Account COM
  object reference does not -- the same portability problem as Folder
  EntryIDs) as a contributing cause; the Rules and Alerts UI showed the
  literal placeholder "the specified account" for affected rules
- Confirmed `0x800C8101` is a genuine environmental/profile-complexity
  limitation (reproduced identically on a fresh, SCANPST-clean OST
  rebuild), not file corruption -- raw `PR_RW_RULES_STREAM` reads via
  `PropertyAccessor.GetProperty()` succeed cleanly even when COM's
  `Item()` enumeration fails on the same store

**Gemini Strategy 1+2 rules engine (first working consolidated approach):**
- Strategy 1: remap known rules to their correct Archive PST folder using
  a one-pass path traversal from the namespace Folders root
- Strategy 2: disable `MoveToFolder.Enabled` on rules with stale targets
  to bypass `Save()` validation, with a single `Save()` at the end of the
  fully processed collection
- `PR_RULES_DATA` binary fallback added as a third enumeration tier
  (reads the raw blob via `PropertyAccessor`, scans for the
  `MSFT:OutlookRules` signature) for when both `foreach` and `GetTable`
  fail
- `Export-RulesBlob` backs up the raw rules stream to `.bin` (reliable at
  any rule count) and `.rwz` (may fail at high counts)

**New: `Read-RulesFromPSTStore` / `Invoke-RulesRecreation`:**
- Shared extraction layer reads rule descriptors from a backup PST store
  and recreates them onto a target IMAP store, remapping folder targets
  via `folder_map.csv`
- `-RecreateRules` switch added to Script 03; IMAP-ALREADY accounts read
  rules directly from the live store, POP3-converted accounts read from
  the backup PST

**New utility scripts:**
- `OMMigrate-FixRulePaths.ps1` -- one-time utility patching missing
  `TargetFolderPath` values via three-tier matching against a reference
  CSV
- `Fix-NeedsFolderAndRulesUpdate.ps1` -- generates a 6-tab Excel audit
  report (`RulesUpdateReport.xlsx`) and merges `TargetFolderEntryID`/
  `Actions`/`RuleType` from a reference CSV via two-tier matching
- `Find-DuplicateRules.ps1` / `Remove-DuplicateRules.ps1` /
  `Remove-DuplicateRuleRows.ps1` -- scan for, remove, and clean up
  ghost CSV rows from `(n)`-suffixed duplicate rules

**Install.ps1:**
- Auto-launch/close of Outlook via COM, replacing manual open/close
  operator instructions across several script phases
- New `Repair-IMAPCredentials`: detects and corrects scrambled IMAP/SMTP
  credentials written by Outlook's Add Account wizard, with before/after
  `.reg` backups
- `-VisibleLaunch` parameter on `Connect-OutlookCOM` fixes a background
  launch issue where the operator had to manually click the taskbar icon
- `New-ArchivePSTViaCOM` pattern extended to create `OMMigrate_Template.pst`
  for the OST-backup pipeline (Script 01 `-BackupIMAPOSTs`)

**Registry and CSV merge hardening:**
- COMPLETE/IMAP-CONVERTED accounts protected from being silently
  downgraded by rediscovery on a Script 00 rerun
- `OSTPath`/`PSTPath` preservation in the CSV merge, gated to matching
  account email (prevents wrong-account paths surviving a rerun)
- Username-only account-to-datafile matching restricted to PST files
  only -- OST files always carry the full email in their filename, so
  username-only matching was causing false matches for accounts sharing
  a common prefix (e.g. `admin@`)

**Known limitations documented at end of arc (open items carried into
1.5.0 above):** default-store purge/recreate architecturally blocked at
scale on accounts with very large rule counts; StopProcessing cannot be
set via COM on a freshly `Create()`'d rule; secondary-store idempotency
not yet implemented.

---

## [1.3.0] -- June 2026

### Automation Pass -- Account Picker and CSV Auto-Update

- **Script 00:** WinForms MIGRATE account picker opens automatically
  after discovery. Operator checks POP3 accounts to migrate now;
  unchecked accounts set to SKIP. Replaces manual CSV editing entirely
- **Script 02:** auto-updates `migration_accounts.csv` on SUCCESS for
  each converted account (`MigrationAction=FOLDER-ONLY`,
  `ProviderTag=IMAP-CONVERTED`, `AccountType=IMAP`) -- no manual CSV edit
  required after conversion
- **Script 03:** Out-GridView picker replaced with a WinForms
  CheckedListBox (fixed-size, centered, no maximize) with Select All /
  Clear All buttons -- eliminates a toolbar overlap issue caused by the
  maximized Out-GridView window
- **`OMMigrate-Registry.psm1`:** `Update-AccountMigrationAction` extended
  with optional `NewProviderTag`/`NewAccountType` parameters (backward
  compatible); new `Invoke-MigrateAccountPicker` exported
- **QUICKSTART.md:** manual CSV edit instructions removed from Steps 2,
  4, and 7 to reflect the automation above

---

## [1.2.0] -- May 2026

### Script 03 End-to-End Verified -- Full Migration Pipeline Complete

This minor release documents all fixes, improvements, and new capabilities
developed during live end-to-end testing of Script 03 (Folder Migration and
Rules Update) against a real Outlook 2021 Classic installation. All 6
IMAP-CONVERTED accounts successfully migrated. The full OMMigrate pipeline
from Script 00 through Script 03 is now proven end-to-end.

---

### Install.ps1 -- Archive PST Creation via Outlook COM

**Retired: binary header PST generation (blocker fix):**
- `New-EmptyPSTFile` retired -- 512-byte binary header was too incomplete
  for Outlook COM `AddStore`, destabilizing the COM session and causing
  backup PST open failures (RPC server unavailable)

**New: `New-ArchivePSTViaCOM`:**
- Launches Outlook invisibly via COM, calls `AddStore` to a non-existent
  path, Outlook creates a fully valid PST structure automatically
- Detaches the store and quits COM cleanly with full object release
- Produces a PST that Outlook opens reliably in Script 03

**New: `Invoke-WaitForOutlookClose`:**
- Detects if Outlook is running via `Get-Process` before COM launch
- Prompts operator: close Outlook, return here, press Y to continue or A to abort
- Loops until Outlook is confirmed closed -- no accidental Enter continues
- Prompt only appears if Outlook is actually running

**Archive PST creation is now a hard requirement:**
- Install halts with clear message on failure -- Script 03 cannot run without it
- Size verification confirms file is credible (>4096 bytes) before reporting success
- Tested all paths: Outlook closed at start, abort, Y while still open

---

### OMMigrate-03-Restore.ps1 -- Script 03 Pipeline

**Per-account COMPLETE tracking:**
- After SUCCESS or WARNING outcome, `Update-AccountMigrationAction` writes
  `COMPLETE` to `MigrationAction` in `migration_accounts.csv` immediately
- FAILED accounts stay `FOLDER-ONLY` and are retried automatically on next run
- Operator-skipped accounts stay `FOLDER-ONLY` and re-appear next run
- Account picker only shows accounts still needing work -- COMPLETE accounts
  excluded automatically. No manual manifest file management required.

**Eligibility display:**
- Console now shows `Already complete: N (excluded from picker)` so operator
  can see overall progress at a glance

**Checkpoint auto-cleanup:**
- `Step03_Checkpoint.json` deleted automatically after a fully successful session
- Failed and warned sessions keep the checkpoint for resume

**Archive PST already-mounted fix:**
- `Open-PSTFile` now checks if the PST is already mounted before calling `AddStore`
- Prevents a second handle being opened on an already-mounted store, which caused
  folder creates to be invisible through the profile handle
- Returns existing store reference directly when already mounted

**Stale comment updated:**
- Step 4 comment updated from binary header approach to reflect COM creation

**WhatIf Archive PST warning fixed:**
- `Archive PST root not available` warning was firing in preview mode because
  `Open-PSTFile` correctly returns null in WhatIf
- Now logs `[WHATIF] Archive PST open skipped in preview mode -- N local folder(s)
  would be created in live run` at INFO level
- Preview mode no longer reports false failures or inflated failed folder count

**Double COM release log fixed:**
- Redundant `Releasing Outlook COM session...` log line removed from `finally` block
- Single authoritative log entry remains in `Release-OutlookCOM`

---

### OMMigrate-Outlook.psm1

**Release-OutlookCOM -- PST write flush fix (critical):**
- `Stop-Process -Force` was killing Outlook 500ms after `Quit()`, before Outlook
  had flushed pending PST write buffers to disk
- Silently discarded folder creates and item copies -- operations logged as complete
  but never persisted to the PST file
- Fixed: polls every 500ms for up to 15 seconds waiting for Outlook to exit naturally
  after `Quit()`. Only force-kills if Outlook has not exited after 15 seconds.
- Verified: Outlook now exits cleanly within the wait window (typically 6-7 seconds)
  with no force-kill required

**Open-PSTFile -- already-mounted store detection:**
- Checks all mounted stores for matching file path before calling `AddStore`
- Returns existing store reference if already mounted, preventing double-handle issue
- Logs `PST already mounted -- using existing store` at DEBUG level

**Suspend-OutlookSendReceive -- ScheduledSendReceive message:**
- Updated catch message from `Could not read OnDemandOnly` to
  `ScheduledSendReceive not available (expected on some group types) -- using safe fallback`
- Clarifies this is expected behavior in Outlook 2021+ -- not an error

**rules_inventory.csv -- NeedsFolderUpdate column reorder:**
- `NeedsFolderUpdate` moved from last column to column D, immediately before `IsEnabled`
- Operator no longer needs to scroll to find it in Excel
- Applied in both rule object construction and blank separator template

**rules_inventory.csv -- merge preservation:**
- Added full merge logic preserving `NeedsFolderUpdate` operator edits across
  Script 00 re-runs
- Keyed by `RuleStoreName|RuleName` -- unique identifier for each rule
- New rules get defaults. Retained/Added counts in log and audit entry.

**folder_map.csv -- key collision bug fix:**
- Merge key changed from `FolderPath` alone to `StoreName|FolderPath`
- Prevents wrong `Destination` merges when two accounts have identically
  named folders (e.g. both have `Inbox\SP`)

---

### OMMigrate-Registry.psm1

**New: `Update-AccountMigrationAction`:**
- Updates `MigrationAction` for a single account in `migration_accounts.csv`
- Called by Script 03 after each successful account migration
- Checks for Excel file lock before writing -- warns operator and skips
  gracefully if file is open, account remains eligible for next run
- WhatIf-aware. Added to `Export-ModuleMember`.

---

### OMMigrate-00-Discover.ps1

**OST path fallback for COM-only accounts:**
- Modern IMAP accounts added via the Add Account dialog do not always get
  registry entries written by Outlook (timing/flush issue during multi-account runs)
- These accounts are discovered via COM but `Join-AccountsWithDataFiles` cannot
  match them -- no registry representation to join against
- Fix: after COM corroboration, scans `%LOCALAPPDATA%\Microsoft\Outlook\` for OST
  files and matches by email address in filename for any IMAP account still missing
  an OSTPath
- Logs at INFO as `OST fallback matched`
- Verified: all affected accounts now show correct OST paths in Discovery report
  and `migration_accounts.csv`

---

### OMMigrate-02-Convert.ps1

**Step02_Checkpoint.json auto-delete:**
- Deleted automatically after a fully successful session, matching Script 03 behavior
- Failed and warned sessions keep the checkpoint for resume

---

### New: Update-Version.ps1 and version.txt

**Centralized version management:**
- `version.txt` in project root is the single source of truth for version number
- `Update-Version.ps1` updates all `.ps1`, `.psm1`, and `.md` files in one operation
- Handles all version patterns: `.NOTES Version:`, `$Script:Version`,
  `$Script:OMMigrateVersion`, markdown headers, CHANGELOG blockquote
- BOM-preserving write -- files keep their original encoding
- Safe to re-run -- idempotent if version is unchanged

---

### Archive PST Architecture -- Final Design

The Archive PST (`OMMigrate_Archive.pst`) serves as the permanent local store
for all Local-destination folders across all migrated accounts:

- Created by `Install.ps1` via Outlook COM before any migration scripts run
- Remains mounted in the Outlook profile automatically after Script 03 runs --
  no manual attach step required
- Each account gets a named subfolder under the Archive PST root
- Local-destination folders and their email items live under the account subfolder

**Folder destination defaults (folder_map.csv):**
- Standard IMAP folders (Inbox, Sent, Drafts, Deleted Items, etc.) default to Server
- All other folders default to Local (Archive PST)
- Operator can change any destination before running Script 03
- After migration, Local folders can be moved to IMAP server by drag-and-drop
  in Outlook -- no scripts required
- If a folder is moved post-migration, any Outlook Rules targeting it must be
  updated manually (File > Manage Rules & Alerts) -- rules do not follow folder moves

---

### Known Remaining Items

- Script 02 -- 7 POP3 accounts still to convert (6 POP3-AWS, 1 POP3-ATTAMERITECH)
- Script 02 -- Step02_Checkpoint.json auto-delete: verify on next live run
- Script 02 -- ScheduledSendReceive message: verify on next live run
- folder_map.csv -- key collision fix: verify when operator makes Destination edits
- rules_inventory.csv -- merge preservation: verify when operator sets NeedsFolderUpdate=False

---

## [1.1.1] -- May 2026

### Script 02 Live Testing -- End-to-End Verification

This patch release documents all fixes and improvements discovered during
live end-to-end testing of Script 02 against a real Outlook 2021 Classic
installation with 26 accounts (12 IMAP, 9 POP3-AWS, 1 POP3-ATTAMERITECH,
2 Exchange). Two accounts successfully migrated from POP3 to IMAP with
full operator-guided manual process verified.

---

### Architecture Change: Registry Operations

**Registry writes and deletes permanently retired:**
- `Remove-POP3AccountViaRegistry` -- retired from active flow, code retained
- `Add-IMAPAccountViaRegistry` -- retired from active flow, code retained
- No registry writes or deletes anywhere in the toolkit -- ever
- `.reg` backup file generation removed -- dangerous post-migration (could re-inject stale POP3 keys)
- All account operations are now guided manual via Outlook UI

**Correct migration sequence established:**
1. Back up PST (Script 01)
2. Remove POP3 via Outlook Account Settings (Script 02 Phase A)
3. Add IMAP via Outlook File > Add Account (Script 02 Phase B)
4. Restore folders and rules (Script 03)

---

### OMMigrate-02-Convert.ps1

**Phase C TCP exhaustion fix (critical):**
- Moved Send/Receive restore out of per-account loop into a single post-loop COM session
- Per-account COM open/close was causing rapid TCP connection bursts to the mail server
- With 25+ accounts this exhausted server connection limits and required a mail server reboot
- One post-loop COM session eliminates the problem regardless of account count

**Phase C AllowRunning fix:**
- Phase C now uses `Connect-OutlookCOM -AllowRunning $true`
- When operator leaves Outlook open after Phase B manual IMAP add, Phase C attaches to
  the existing instance instead of failing with "Outlook is already running"

**Phase C two-path logic:**
- Path 1: COM session still open (all-skip path) -- uses existing session directly
- Path 2: COM session closed (migration path) -- opens new session with AllowRunning support

**Zero-MIGRATE gate:**
- Script exits cleanly before launching Outlook when no accounts are marked MIGRATE in CSV
- Displays clear message and directs operator to Script 03
- Still generates report and manifest for the run

**Conversion summary improvements:**
- Added `Already IMAP (no action)` counter -- accounts with IMAP-ALREADY/IMAP-GMAIL ProviderTag
- Added `Live IMAP detected (re-run)` counter -- MIGRATE accounts found already IMAP in Outlook
- Added `Exchange (no action)` counter -- EXCHANGE-SKIP accounts
- Added `Connection unconfirmed` counter -- separated from Skipped by operator
- Added `Not marked (other)` counter -- remaining non-MIGRATE accounts
- Added separator line and `Total accounts identified` -- always equals total account count
- All counter colors updated for visibility on black terminal background

**Verify prompt updated:**
- Phase D no longer asks "connected and showing mail" -- asks if account visible in folder pane
- Added note about manual credential correction required after script completes
- Operator correctly closes Outlook before typing C, then confirms folder pane visibility

**Prompt sanitization fixes:**
- REMOVE POP3 prompt now shows sanitized email address
- Backup gate check now shows `[backup-file].pst` placeholder instead of real PST filename
- Both console and log use masked path at gate check level

**Duplicate log line fix:**
- Removed duplicate "POP3 account removal confirmed" log line from Script 02
- Authoritative log line remains in `Remove-POP3Account` in OMMigrate-Outlook.psm1

**WhatIf cleanup:**
- Removed stray `WhatIf: Would resume Outlook Send/Receive groups` from finally block
- Finally block now skips Resume/Release entirely in WhatIf mode

---

### OMMigrate-Outlook.psm1

**ManifestPath typo fix (critical):**
- `ManifestsPath` corrected to `ManifestPath` in both `Suspend-OutlookSendReceive`
  and `Resume-OutlookSendReceive` -- was causing Send/Receive state file save and
  restore to fail silently, leaving groups suspended after script ended

**Connect-OutlookCOM -- AllowRunning parameter:**
- Added `$AllowRunning = $false` parameter
- When `$true`, attaches to existing running Outlook instance via
  `Marshal.GetActiveObject('Outlook.Application')` instead of failing
- Used by Phase C post-loop Send/Receive restore

**Phase B operator instructions updated:**
- Now instructs operator to close Outlook after account appears in folder pane
- Added explicit warning: do not send or receive mail until credentials corrected
- C prompt updated: "visible in folder pane and Outlook is closed"

---

### OMMigrate-01-Backup.ps1

**Prompt sanitization fix:**
- Per-account backup confirmation prompt now shows sanitized email address
- Both "Export PST backup for" prompts updated to use `Invoke-OMMigrateSanitize`

---

### OMMigrate-Reporting.psm1

**HTML report tag badge fix:**
- Migration row now shows `IMAP-AWS`, `IMAP-STANDARD` etc. instead of `POP3-*`
  when MigrationDetail confirms account is already IMAP
- Reflects current protocol state rather than original CSV tag

**HTML report action badge:**
- `MIGRATED*` badge (amber) shown when account still has MIGRATE in CSV but
  was detected as already IMAP by live Outlook check
- `COMPLETED*` badge (amber) shown on subsequent re-runs before CSV is updated
- Both badges signal operator to update CSV

**HTML report footer note:**
- Footer note appears below Migration table when any MIGRATED*/COMPLETED* accounts present
- Instructs operator to update three CSV fields: MigrationAction, ProviderTag, AccountType
- Explains FOLDER-ONLY vs SKIP distinction -- Script 03 still needs FOLDER-ONLY accounts

**Overall status fix:**
- All-skipped runs with zero failures now correctly report SUCCESS not WARNING

---

### OMMigrate-Core.psm1

**badge-update CSS class:**
- Added amber badge style for MIGRATED* and COMPLETED* action badges
- Visually distinct from all other badges -- signals operator action required

---

### Known Outlook 2021 Classic Behavior (documented)

**Credential scrambling during Add Account:**
- Outlook 2021 Classic forces IMAP username to match email address during Add Account
- SMTP username field is overwritten with the email address, replacing any custom credentials
- This is a Microsoft design limitation -- no patch available
- Workaround: correct IMAP and SMTP credentials manually via Account Settings > Change
  after each migration before attempting Send/Receive
- Documented in QUICKSTART.md operator guide

**TCP connection exhaustion:**
- Opening/closing Outlook COM sessions rapidly across multiple accounts leaves
  half-open TCP connections on the mail server
- With 25+ accounts this can require a mail server reboot to clear
- Fixed by consolidating to a single post-loop Phase C COM session

---

### One-Command Installer -- Simplified Install Experience

This minor release replaces the manual setup and file download process with
a single PowerShell one-liner that handles everything automatically.

---

### New: Install.ps1 -- Bootstrap Installer

**The entire install process is now one command:**

```powershell
irm https://raw.githubusercontent.com/SC-Admin567/OMMigrate/main/Install.ps1 | iex
```

`Install.ps1` replaces `Setup-OMMigrateProject.ps1` as the primary entry
point for new installations.

**What the installer does automatically:**
- Sets PowerShell execution policy to RemoteSigned for the current user
- Creates `Documents\OMMigrate\` project folder with Scripts\ and Modules\ subfolders
- Creates `Documents\OutlookMigration\` runtime data folder with all subfolders
- Downloads all 8 source files (4 scripts, 4 modules) from GitHub
- Downloads all documentation files (README, QUICKSTART, CHANGELOG, LICENSE)
- Runs Unblock-File on every downloaded file
- Prints next steps when complete

**Safe to re-run:**
Re-running the installer downloads fresh copies of all files. Runtime data
in `Documents\OutlookMigration\` is never touched.

**Download method:**
All files are downloaded via `Invoke-WebRequest` directly from GitHub raw
content URLs. No text editor involved -- no encoding corruption risk.

---

### Retired: Setup-OMMigrateProject.ps1

`Setup-OMMigrateProject.ps1` is retired and removed from the repository.
Its functionality is fully replaced by `Install.ps1`.

---

### Retired: Stub Files

Stub placeholder files for Scripts\ and Modules\ are eliminated entirely.
The installer downloads real files directly -- no manual drag/drop or
file-by-file replacement required.

---

### Updated: README.md

- Install section now leads with the one-liner command
- Removed all references to Setup-OMMigrateProject.ps1 and stub files
- Removed manual Unblock-File instructions (handled automatically by installer)
- Added "Updating OMMigrate" section showing how to re-run installer for updates
- Cleaned up directory structure diagram to reflect new layout

---

### Updated: QUICKSTART.md

- Step 0 is now the install one-liner -- all manual setup steps removed
- Removed execution policy and Unblock-File manual steps (installer handles both)
- Removed drag/drop and encoding warning sections (no longer relevant)
- Simplified working directory guidance
- All content now assumes installer has already run

---

### New: .gitignore

Added `.gitignore` to the repository root:
- Excludes all runtime output folders (Config, Logs, Manifests, Reports, Backups)
- Excludes all Outlook data files (PST, OST, NST)
- Excludes all generated output files (logs, CSVs, HTML reports, manifests)
- Excludes OS noise (Thumbs.db, desktop.ini)
- Excludes editor noise (.vscode, .sublime-project, etc.)
- Ensures no test data or personal account information can be accidentally committed

---

### Repository Structure Change

Runtime data folders (Config, Logs, Manifests, Reports, Backups) are no
longer present in the repository at all. They exist only at
`Documents\OutlookMigration\` on the operator's machine, created
automatically by the installer and scripts.

The repository now contains only:
- Source files (Scripts\, Modules\)
- Documentation (README.md, QUICKSTART.md, CHANGELOG.md, LICENSE.md)
- Installer (Install.ps1)
- .gitignore

---

## [1.0.1] -- May 2026

### First Testing Release -- Script 00 Verified

This patch release documents all fixes and improvements discovered during
the first end-to-end test run of OMMigrate on a live Outlook 2021 installation
with 26 accounts (13 POP3, 11 IMAP, 2 Exchange), 2124 folders, and 672 rules.

Script 00 (Discovery) now runs to full SUCCESS with 0 errors and 0 warnings
on a clean run.

---

### Environment and Setup Fixes

**PowerShell execution policy and file unblocking:**
- Documented that `RemoteSigned` execution policy must be set before any script runs
- Added mandatory `Unblock-File` step for all downloaded `.ps1` and `.psm1` files
- Windows attaches a Zone.Identifier security flag to downloaded files;
  `Get-ChildItem -Recurse -Filter *.ps1 | Unblock-File` removes it for all files at once

**File encoding:**
- Identified that copying code through text editors (Sublime Text, Notepad++) corrupts
  UTF-8 special characters (em dashes, arrows, box-drawing characters) causing parse errors
- All script and module files must be saved directly via browser download, never via copy-paste
- Applied encoding fix across all modules and scripts -- replaced all non-ASCII characters
  with plain ASCII equivalents (-- for em dash, -> for arrow, - for box drawing)

---

### OMMigrate-Registry.psm1

**Registry path correction (critical fix):**
- Discovered that Outlook 2016/2019/2021 stores account configurations under
  `HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles` NOT under the legacy
  `Windows Messaging Subsystem` path used by older Outlook versions

**Account GUID discovery (major improvement):**
- Discovered the well-known Outlook account service GUID `9375CFF0413111d3B88A00104B2A6676`
  consistent across all Outlook versions (2016/2019/2021)
- Properties Email, POP3 Server, IMAP Server, SMTP Server, ports, and SSL settings
  all stored as plain strings -- no binary blob decoding required
- Rewrote `Get-OutlookAccountsFromRegistry` to scan this GUID path directly

**PST/OST path extraction:**
- PST and OST file paths extracted via byte-level scanner locating drive letter
  pattern within the binary blob, correctly handling both PST and OST paths

**Domain-based server lookup table (new feature):**
- Added `$Script:KnownProviderSettings` hashtable covering well-known public
  mail providers as a fallback when registry values are missing

**StrictMode compatibility:**
- All registry property reads updated to use `PSObject.Properties[$name]` pattern

---

### OMMigrate-Core.psm1

**Outlook auto-close on startup:**
- Pre-flight check now prompts to close Outlook automatically instead of hard-failing

**Outlook auto-close on completion:**
- Script 00 automatically closes Outlook at end of successful run

**Counter reset:**
- Explicit reset of all counters at start of `Initialize-OMMigrate` prevents
  accumulation when scripts are run multiple times in the same session

---

### OMMigrate-Outlook.psm1

**Store type classification fix:**
- `Get-StoreType` reordered to check `ExchangeStoreType` before file path extension
- Exchange, IMAP, and POP3 accounts now correctly classified

**Rules error level:**
- Unreadable rules now log at INFO instead of WARN

---

### OMMigrate-00-Discover.ps1

**Status logic fix:**
- Script 00 now achieves SUCCESS on clean run, WARNING only for legitimate
  operator attention items. Error counter alone no longer causes FAILED status.

**WhatIf parameter fix:**
- Removed manual `[switch]$WhatIf` that conflicted with CmdletBinding automatic support

---

### OMMigrate-Reporting.psm1

**Data file display:**
- PST and OST filenames now shown for both POP3 and IMAP accounts
- Full path shown as tooltip on hover

**Array/Count compatibility:**
- All Where-Object results wrapped with `@()` for StrictMode compatibility

---

### Known Remaining Items (carry forward to v1.0.2)

- Scripts 01-03 not yet tested against a live Outlook installation
- 2 corrupt Outlook rules in test environment log at INFO level and are not fixable by OMMigrate

---

## [1.0.0] -- May 2026

### Initial Release -- Inception and Architecture

OutlookMailMigrator was conceived and built as a full collaborative project
from day one between the Tool Architect (Originator and Architect) and
Anthropic Claude AI (Implementation Specialist).

The Tool Architect identified the problem -- Google's own documentation
stated that migrating Outlook desktop profiles from POP3 to IMAP across
multiple accounts had no automated solution. The Tool Architect defined
the requirements, made every architectural decision, drove the design
through every phase, and served as the sole tester against a real
production Outlook 2021 Classic installation with 26 accounts.

Claude contributed throughout as a genuine technical partner -- not simply
executing instructions but actively shaping the solution. Claude proposed
and designed core technical approaches that the Tool Architect evaluated
and approved: the COM session architecture, PST binary structure analysis,
registry path discovery strategy, checkpoint and resume design, the
sanitization system, exit handler safety model, the ProviderTag
classification system, and the Archive PST 3-tier storage architecture.
During debugging, Claude diagnosed root causes from console output and log
files -- identifying the TCP connection exhaustion bug, the PST write
flush timing issue, the COM double-handle problem, the registry key
collision in folder map merging, and the OST path discovery gap for
COM-only accounts. When the Tool Architect pushed back on an approach,
Claude adapted. When the code didn't match the design, Claude found the
exact failing line before proposing any fix.
The community should know: this toolkit reflects what is possible when
a domain expert architect works in close collaboration with AI as a
genuine engineering partner, not just a code generator.

The project was designed from inception with marketability in mind -- built
not just for one operator but as a toolkit any experienced administrator
could deploy against any Outlook 2021 Classic installation.

---

### Core Architecture Decisions -- Established at Inception

**Six-script pipeline:**
- Install.ps1 -- Bootstrap installer, one-command setup
- Script 00 -- Discovery: registry and COM account enumeration, folder tree,
  rules inventory, CSV generation
- Script 01 -- Backup: PST export of all POP3 accounts before any changes
- Script 02 -- Convert: guided POP3 removal and IMAP add via Outlook UI
- Script 03 -- Restore: folder migration, email copy, rules update
- Script 04 -- Artifacts: Migrates Calendar, Contacts, Tasks, Notes, and Journal

**Module architecture:**
- `OMMigrate-Core.psm1` -- session management, logging, checkpointing,
  progress tracking, exit handlers, audit trail
- `OMMigrate-Registry.psm1` -- registry-based account and data file discovery,
  CSV export, account classification
- `OMMigrate-Outlook.psm1` -- COM session management, folder operations,
  PST/OST handling, rules inventory, Send/Receive control
- `OMMigrate-Reporting.psm1` -- HTML report generation for all script phases

**Registry is read-only:**
- Registry used for account discovery (Script 00) only
- No registry writes or deletes anywhere in the toolkit
- All account operations via Outlook UI or COM -- never direct registry manipulation

**ProviderTag classification system:**
- `POP3-AWS`, `POP3-ATTAMERITECH`, `POP3-STANDARD` -- accounts to migrate
- `IMAP-ALREADY`, `IMAP-GMAIL`, `IMAP-CONVERTED` -- already correct or converted
- `EXCHANGE-SKIP` -- Exchange accounts, skip entirely
- Tags drive all downstream migration decisions

**Archive PST architecture (3-tier local storage):**
- IMAP server -- live mail, all devices
- `OMMigrate_Archive.pst` -- local Archive PST, Script 03 Local-destination folders
- `C_Archive.pst` -- deep history, Outlook auto-archive

**Checkpoint and resume system:**
- Per-account checkpoint written after each account completes
- Resume from checkpoint on re-run after interruption
- Progress tracked independently of manifest files

**Sanitization system:**
- `-Sanitize` flag masks all email addresses, machine names, paths in console
  and log output for safe sharing and documentation
- Sanitization map built at script start from discovered account data

---

### Initial Design Increments

**Account discovery via registry:**
- Researched Outlook 2021 registry structure -- accounts stored under
  `HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles` using well-known
  GUID `9375CFF0413111d3B88A00104B2A6676`
- Binary blob parsing for PST/OST paths, plain string properties for server
  settings, ports, SSL flags
- Domain-based fallback lookup table for known providers

**COM corroboration:**
- Registry discovery supplemented by live Outlook COM session
- COM is authoritative for account type, server names, and data file paths
- Accounts found in COM but not registry are added automatically

**Folder tree enumeration:**
- Full recursive folder tree from all mounted stores
- System folder exclusion, deduplication, depth tracking
- `folder_map.csv` with per-folder Destination column for operator review

**Rules inventory:**
- All Outlook Rules enumerated across all stores (not just DefaultStore)
- `rules_inventory.csv` with `NeedsFolderUpdate` flag for Script 03
- 806 rules discovered vs 674 in earlier partial implementation

**Logging and audit trail:**
- Timestamped run log at configurable level (DEBUG/INFO/WARN/ERROR)
- Separate audit log for all account operations
- Session summary at end of every run
- HTML reports for operator review after each script phase

**Exit handler safety:**
- `PowerShell.Exiting` engine event registered at script start
- Ctrl+C intercepted for clean COM release before exit
- `SessionCompletedNormally` flag prevents emergency handler firing on clean exit

---

## Related Documents

- **[README.md](README.md)** -- Full overview, features, and troubleshooting
- **[QUICKSTART.md](QUICKSTART.md)** -- Guided first-run walkthrough
- **[OMMigrate_CommandLine_Reference.md](OMMigrate_CommandLine_Reference.md)** -- Every script parameter
- **[OMMigrate_Settings_Reference.md](OMMigrate_Settings_Reference.md)** -- Every settings file key

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
