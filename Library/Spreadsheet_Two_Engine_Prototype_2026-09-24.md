# Persistent spreadsheet worker prototype — 24 September 2026

Status: isolated research; **not integrated into Summit or approved for client files**.
Production remains 2.94. Original AGL, Blank and Demo workbooks are not edited.

Later checkpoint: [Gear projection and native fallback trial](Spreadsheet_Mirror_Projection_Trial_2026-09-24.md). Current production is 2.98, unchanged by that research. The DevExpress full Funding trial now runs but fails the array-preservation gate; the earlier pending status below is historical.

The 2.94 reference above records the production version at the prototype checkpoint, not the latest UI test build. Subsequent [vendor confirmation](SpreadsheetGear_Support_Package_2026-09-24.md) establishes that grouped-sheet reference fix-up, custom XML preservation and dynamic-array semantics are unsupported in Gear. This prototype does not remove those limitations; authoritative Gear saves and an XML-only repair strategy are not approved production routes.

## Agreed workload distinction

The user describes two phases: **build/update**, dominated by population and structural changes, and **review**, assessing the consequences of small changes. Optimise them separately rather than treating a fast scalar calculation as a complete workflow benchmark.

- Build/update: the authoritative worker should own structural commands. Pause commands targeting affected coordinates, finish the command, then refresh Gear at that revision. Avoid showing speculative, known-incorrect 3-D results. Batch related work where user-visible transaction boundaries permit it.
- Review: journal typed edits, calculate the Gear working copy and queue the same edits to the worker. Ordinary edits need acknowledgements, not a workbook read-back. Periodic parity checks remain necessary.
- Save/close: wait for the requested revision, serialize the authoritative worker, validate a private candidate, and publish only if the destination has not changed. An acknowledgement for revision N must not clear a later edit at N+1.
- Neither phase changes the authority contract, protection rules, history, recovery policy or the user's Excel/VBA round-trip requirement. Users should not have to select a spreadsheet engine each time.

Excel is an optional worker on compatible, licensed client installations; DevExpress is the alternative under evaluation. Excel automation runs in an owned process in the logged-in desktop session, not a Windows service. Existing macro policy remains in force: no force-enable setting or Trust Center changes. This trial has explicit permission for isolated model VBA; production onboarding and failure handling are not implemented.

## Implemented, bounded prototype

`Tools/SpreadsheetEngineTrial/MirrorTrial.cs` adds two-process `mirror` / `mirror-worker` commands. Gear and the native worker retain separate workbooks, with one serial owner per workbook. There is no live Summit UI adapter or shared workbook access.

The journal has schema, session, immutable baseline hash, revision, prior-record hash, type and expected previous value. Each record is flushed then renamed from a pending file. Duplicate exact records are acknowledged once; gaps and conflicting records are rejected. Restart replays committed records from the immutable baseline. The prototype does not yet recover a crashed parent/UI session or compact the journal.

Save candidates, backups, targets and reports are restricted to unique ignored directories below `obj/AsposeTrial`. Only the private trial target is replaced. Existing target hashes, a write-excluding file lease, an exact revision barrier and a separate backup protect publication. This is not a claim of a complete multi-process filesystem transaction: production needs explicit model ownership and handling for external renames, disk-full, power loss and the replace/ack crash window.

The first structural read-back is intentionally coarse: save a native checkpoint and reload/recalculate Gear from it. This avoids guessing which formulas, names, ranges, locks and array anchors changed. It is not a targeted range-delta transfer and it has measurable save/reload costs. Gear output is never used as the authoritative saved workbook.

## Completed Excel-worker AGL measurement

One sequential x64 Release pass on this development machine, warm OS cache, same immutable Excel-converted AGL XLSM as the prior benchmarks. This is **not** a minimum-spec or Summit UI timing.

| Operation | Measured time |
|---|---:|
| Worker initial load and rebuild | 19.08 s |
| Gear initial load and rebuild | 4.57 s |
| Three journal + Gear edit/recalculation/input-read operations | 0.719 / 0.559 / 0.490 s |
| Corresponding Excel edit application, without calculation | 0.508 / 0.095 / 0.149 s |
| Gear edit while the earlier revision was saving | 0.523 s |
| Worker save, full calculation and package checks | 6.41 s |
| Funding VBA insertion + eleven TDB mirrors + rebuild | 18.39 s |
| Structural checkpoint save | 6.75 s |
| Gear structural reload + rebuild | 4.74 s |
| Complete structural barrier, including parallel trial Gear insertion | **29.94 s** |

The native VBA components were 11.81 s for 32-sheet insertion and 1.97 s for the eleven mirrors. The test also ran the raw Gear insertion (3.92 s including rebuild) in parallel to the worker, but **discarded that provisional state**. That duplicate operation is a comparison workload, not a recommendation for build/update mode.

Peak working sets measured separately: Gear harness 335.2 MiB, worker host 59.0 MiB, owned Excel 751.2 MiB. Their sum is not a sampled simultaneous peak or a whole-Summit memory estimate. More than one open model and 32-bit operation remain untested.

## Correctness evidence and limits

After authoritative structural read-back, all **38,108 nonempty cells** across the five established calculated probes match the independent Excel/VBA Funding oracle, within the established absolute 1e-6 / relative 1e-10 tolerance. No numeric, text or blank/type differences were found. This removes the previously observed 998 numeric discrepancies in those probes; it is not financial certification of the whole workbook.

The saved worker workbook also matches the oracle's per-sheet dynamic-array anchor counts, array ranges and cell-metadata counts. Its custom-XML parts and drawing/control/chart family counts are retained. Counts are not a visual or behavioural acceptance test.

The final independent checks found **1,126,327 formula cells in both workbooks**, identical per-sheet formula-address coverage, no defined-name differences, and no formula-text differences at any reference formula address in the native DevExpress comparison. Together these checks cover additional/missing formulas as well as changed reference formulas; they do not certify every calculation or Excel/VBA behaviour.

Excel renumbers custom XML ZIP parts while retaining item identities and payloads. Therefore comparison is by item ID and its linked payload/properties, not by `item1.xml` filename alone. Excel also regenerates the VBA binary: the separate, in-memory audit found 335 module identities, with only `Attribute VB_Base` changes in six forms and no other module-source changes. Form designer behaviour, references and compiled execution still require a dedicated round-trip gate. No raw VBA source or passwords are exported.

Synthetic tests passed for duplicate handling, out-of-order rejection, failed-save preservation, locked-target rejection, stale destination rejection, saving N while N+1 stays dirty, and worker restart/replay. DevExpress additionally received an intentional process termination before restart; Excel has so far been restarted gracefully to avoid leaving an uncontrolled COM process. Parent crash, Excel hang/crash and mid-structural failure recovery are not proven.

### Important exception: ordinary edits can require read-back

The existing synthetic dynamic-array fixture starts with three results. A **value-only** change from 3 to 5 in its input expands the result to five cells in both native workers. Gear still returns three values and two blanks. Therefore “read back only after structural insert/delete” is not universally correct.

Possible production approaches are a verified scalar-only eligibility rule, a narrow authoritative read-back for potentially spilling dependencies, or keeping those views on the native engine. Do not silently present stale results or infer non-spilling behaviour from today's populated AGL values. The earlier inventory found current scalar results, not a proof that all future inputs remain scalar.

### DevExpress preservation adapter

Native DevExpress serialization changed the synthetic custom XML's schema reference. For this prototype only, journal commands never edit Summit XML/history, so its still-current original custom XML parts can be restored to the native candidate. This is **not** a production history strategy: production must write the latest model XML, not restore old payloads. DevExpress also rewrites the VBA binary while preserving all 335 source hashes in the examined pre-insert AGL candidate.

The real-model DevExpress run stops at a compatibility gate **before its Funding insertion**: Gear rejects the native DevExpress-saved AGL XLSM with `IOException: Corrupt OpenXML document`. A separate read-only, macro-disabled Excel check opens that same file normally, with no repair mode requested. Thus this is a demonstrated DevExpress-to-Gear reader compatibility failure, not proof that Excel cannot use the file. Do not offer this backend as an approved drop-in fallback. Its cold load/rebuild was 36.16 s, three Gear edit loops 0.712 / 0.562 / 0.521 s, and the native save/check stage 13.65 s; the full structural workflow has **not** passed.

The unexercised native Funding adapter reuses the built production AST 3-D-reference guard through reflection rather than introducing a regex formula repair. It remains isolated research, not a replacement for production transaction/change services. A direct range/formula transfer could avoid this serialization boundary, but is not implemented or validated here.

Both native checkpoints opened in Excel in normal read-only mode with macros, events and link updates disabled; each had 283 sheets, 1,752 names and a VBA project. That check did not run their VBA or certify calculation/visual behaviour.

## Remaining gates before production work

Follow-up diagnosis: disabling Gear object import (`ReadObjects=false`, VBA retained) opens the same DevExpress checkpoint, and 38,108 calculated probe cells match the pre-save baseline. The failure is isolated to the object-import path; the exact XML cause and complete structural workflow remain unproven. See [DevExpress/Gear import diagnosis](SpreadsheetGear_DevExpress_Import_Diagnosis_2026-09-24.md). This is a candidate calculation-copy adapter, not a setting to use for authoritative Gear saves.

1. Resolve the DevExpress-to-Gear checkpoint compatibility gate before its structural trial. Expand the successful Excel-worker Funding test to Development, deletes and populated boundary cases; complete interactive Excel/VBA/form acceptance.
2. Establish dynamic-array eligibility/fallback and dependency-correct displayed results in review mode.
3. Integrate an engine-neutral command boundary with `ModelChangeManager`, fill-based editability, validation, history, undo/redo and source selection. No changes to those protections are included here.
4. Test actual DIT binding/refresh and edit latency; the current harness does not render Summit controls.
5. Implement production worker ownership, cancellation boundaries, rejected/busy COM calls, bounded timeouts, worker crash recovery and confirmed save/close status.
6. Persist current Summit XML/history and recovery state with the authoritative revision. Add full save/replace/reopen failure tests, including disk-full and external edits.
7. Measure both phases on representative populated models and client/minimum-spec machines. Do not infer UI performance from warm native-engine benchmarks.

No automatic engine switch, background replacement of customer files, production migration, hidden verification cells or VBA changes have been introduced.

## Evidence and reproduction

- Main run: ignored `obj/AsposeTrial/mirror-20260924-v1/excel-agl-2/report.json`.
- Independent probe/array/VBA audit, including formula-address and defined-name coverage: `excel-agl-2/independent-audit-v2.json` in that directory.
- Synthetic controls: `excel-synthetic-2/report.json` and `dx-synthetic-4/report.json`.
- DevExpress model gate: `dx-agl-2/report.json`; normal Excel reopen: `native-reopen.json`.
- Full reference-formula comparison: `excel-agl-2/full-formula-comparison.json` (no differences at reference formula addresses).
- Code: `Tools/SpreadsheetEngineTrial/MirrorTrial.cs`, `Audit-MirrorTrial.py`.
- Reference: `Library/SpreadsheetGear_Funding_Full_Trial_2026-09-24.md`.

Build the standalone harness in Debug and Release, x64. Its commands require fresh output directories and private XLSM inputs. The optional Python VBA audit uses the existing temporary oletools installation outside the repository. Do not share generated client workbooks or private reports as vendor support material.

Validation: standalone x64 Debug and Release builds passed, Python syntax validation passed, and repository whitespace checks passed (existing LF/CRLF notices only). SHA256 checks confirmed the original AGL and both authoritative masters match their pre-trial hashes. No production executable, version, package reference or live workflow was changed; no commit or push was made.

Official threading context: [Microsoft Office threading](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=visualstudio) and [SpreadsheetGear threading/locking API](https://spreadsheetgear.com/support/help/spreadsheetgear.net.9.0/Key_Concepts_SpreadsheetGear_API.html). Keeping separate serial owners avoids concurrent calls to the same workbook; it does not eliminate Excel's modal/busy-state handling requirements.
