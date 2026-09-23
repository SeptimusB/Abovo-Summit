# Recovery user-change tracking and prefix - 22 September 2026

Test delivery: Summit 2.71. No source workbooks are modified.

## Policy

- New recovery saves use ~<original-stem>_recovery.xlsm beside the plan.
  The tilde is a visual distinction, not a security boundary or file-open lock.
- Both current prefixed and legacy unprefixed files are recognised, but only
  when their embedded provenance identifies the original. Discovery chooses
  the newest valid copy newer than the original. It never renames or deletes
  an older recovery; unrelated filename collisions are not overwritten.
- The model now tracks a separate committed-user-change revision and normal
  save baseline. Calculation revisions still serve calculation/integrity
  invalidation; they no longer decide whether to make a recovery save.
- Typed edits, paste, Undo/Redo, native spreadsheet input/structure events,
  imports, successful structural rules, user snapshot creation, and committed
  Stress Test changes mark user work. Unchanged/rejected typed edits do not.
- ContentChanged, document metadata, calculation/rebuilds, background TDB
  synchronisation/invalidation, and safety-state dirtiness alone do not.
- A recovery requires unsaved user work. Once backed up, no further scheduled
  recovery occurs until another user revision. Normal Save clears unsaved-user
  state. Recovery preserves it, the live file path, Undo, and protection.
- Existing progress notices, idle/edit/bulk/health guards, atomic replacement,
  restored-history prompt and explicit recovered-file Save As guidance remain.
- SetDirtyFlag is the existing import/bulk user-command entry point; its new
  optional userChange parameter must be False for background metadata. New
  background callers must not call MarkUserChange or the default user path.

## Evidence

Debug and Release builds pass. Synthetic recovery tests pass, including actual
due timer ticks with a notification owner: background-only dirtiness, repeated
calculation after backup, clean save baseline, subsequent edits, provenance,
legacy filename discovery, paste/no-op paste, failed typed input, Undo/Redo,
native handler routing, imports, locked-file/serialization failure preservation,
history, protection, pending groups, notification ownership and metadata bridge.
Native user-event API modification raising is verified disabled; the native
route test invokes the handler, not a real mouse/keyboard acceptance workflow.

Synthetic evidence directory:
obj/RecoveryTests/e85a8d1494b74a82a75e48b4a624cf1b.

Populated 32-bit AGL test passed on a disposable copy: insert one Funding
record, confirm user-dirty eligibility before any cell edit, make a typed stock
edit, run the due recovery timer, independently reopen the XLSM, reopen through
Summit, recover read-only edit history, and Save As a normal XLSB. All formulas,
constant inputs, global/local names, sheet order and protection matched across
recovery export. Funding geometry and all supported TDB mirrors remained
consistent. Save As cleared both normal and user-dirty state. The original's
hash was checked unchanged by the harness.
Evidence: obj/RecoveryTests/1b2707dedb65480c8e3e60111f4e016f.

Existing Release save-state and idle-integrity regressions also passed:
obj/SavePreparationTests/1e19548c1f2142ca9e5221c61d4d0517 and
obj/IdleIntegrityTests/9db8f7390fbc43f5b3aa08bae0c6e905.
Options guidance was visually checked and shortened to keep Save As instructions
within the tested window.

The direct MATCH implementation and calculation evidence are in
PMCost_Repair_and_Profile_2026-09-22.md. Runtime fixtures save only new disposable
files under obj and verify hashes of any supplied source.

## Acceptance still required

Client test: edit/paste, wait for one recovery, leave calculation/integrity checks
running and confirm no further recovery until another edit. Test native
Spreadsheet input and any client-specific import/custom structural paths.
Open the offered recovery, review history, Save As a new XLSB, then perform the
trusted Excel/VBA reopen/recalculate/save and Summit reopen. Native engine
round trips do not certify financial parity with VBA enabled.

API reference: [DevExpress 25.2 CellValueChanged](https://docs.devexpress.com/WindowsForms/DevExpress.XtraSpreadsheet.SpreadsheetControl.CellValueChanged?v=25.2)
distinguishes committed UI inputs from API edits and formula-result changes.
