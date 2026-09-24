# Check Sheet full-calculation trial — 2.92

Status: **Ready to test** with Jon. Debug and Release updated. No commit/push requested; masters and client files unchanged.

## Change and boundaries

The user reported a sheet calculation that did not expose an input imbalance and requested CalculateFull. The opt-in Check Sheet watcher now calls `IWorkbook.CalculateFull()` before reading the existing override-aware Check Sheet result. It temporarily selects Manual/Recursive calculation and disables `CustomCalcEngine.DontCalcTDBS`, restoring all three settings in nested Finally blocks. The deferral switch covers Transactional DB, Comparison and Check Sheet; leaving it enabled can prevent even an explicit sheet calculation from updating the result.

This is a full calculation, not CalculateFullRebuild or a new structural repair. Soft balance warnings remain separate from formula/reference integrity failures. Existing user/dirty/history/rebuild/pending-result flags are not certified or cleared by the trial. Existing same-revision suppression, interval, idle, pending-editor, save, recovery, operation and model-health gates remain. A standard progress notice covers the atomic calculation; once started it must finish safely. Subsequent work still waits for the existing idle boundary.

Options now describes the full-workbook cost. The opt-in benchmark line is `[Check Sheet Trial Benchmark] ... mode=full, calculationAndRead=...`; other tracing stays disabled. No options are automatically enabled or persisted by this change. The original trial-acceptance reminder remains applicable.

DevExpress reference: [CalculateFull](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.IWorkbook.CalculateFull) recalculates cells whether or not they are marked dirty. Behaviour was exercised against the locally installed 25.2 assemblies, not inferred solely from current online documentation.

## Evidence

- Both normal Debug and Release builds pass, version 2.92.
- `Tools/CheckSheetFullFixture.cs`: 12 assertions pass on an unsaved disposable Demo copy (`obj/ClientReportTests/93909328398a4490ae7650457965439f`). A private Stock Assumptions input feeds another formula and Check Sheet B22. With the deferral switch enabled, the old sheet-only call retains zero after the input changes. The new watcher produces 10, raises the soft warning, and clears it when the input is corrected. Calculation engine/mode/deferral, dirty/rebuild/pending flags, user/calculation revisions, history and persisted warning settings remain unchanged. Formula-error and progress-guard cleanup paths also pass.
- That measured full-watch call took 7,105 ms on this workstation/private Demo copy. This is not a representative client-machine benchmark or a timing guarantee.
- Existing client-review regression covers Yes/No/clear/Undo/Redo, unsaved state, recovery policy, quiet-open clearance, idle/save gates, same-revision suppression and persistence only on explicit save. Evidence: `obj/ClientReportTests/4c94c29b7966491b8244e79b5895efde`.
- Quiet-diagnostics gate passes. Source hashes unchanged. The test probe formulas exist only in the discarded in-memory copy; no production sheets or metadata were added.

## Stable functional test 60 — Jon first

On a disposable copy of the affected plan, repeat the input that previously failed to appear in Check Sheet. Wait for the configured Check Sheet trial interval and idle period. Confirm the full-calculation progress, the updated balance message/indicator and Check Sheet result. Correct the input and confirm the next check clears the soft warning. Note the elapsed `mode=full` trial timing if timings are enabled, then disable temporary timings after acceptance. Check ordinary editing and recovery still follow the chosen policies.

The user's exact original breakage has not been reproduced because its input/cell was not specified. The controlled regression establishes that the identified deferred/cross-sheet path is addressed. Client/accountant and trusted Excel/VBA acceptance remain manual.
