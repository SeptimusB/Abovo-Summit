# Error handling and workbook safety audit

Date: 3 September 2026

## Scope and priority

This pass concentrated on the highest-consequence paths: Transactional DB synchronisation, multi-sheet structural changes, model imports, interactive workbook edits, save/close behaviour, worksheet protection restoration, and durable user-facing reporting.

No XLSB file was changed by this pass. Both Debug and Release configurations compile successfully. The repository has no automated workbook/UI integration suite, so the manual tests below remain required before release.

## Safety contract implemented

- `ModelSafetyManager` records a model-level `RecoveryRequired` state whenever a change may have begun and rollback or cleanup cannot be verified.
- A recovery-required model cannot use the ordinary overwrite save path. Summit directs the user to save a separate recovery copy, preserving the original workbook.
- System messages contain the failed operation, source and workbook location, so the existing message view can be saved or emailed.
- Cell-level operations and supported imports use a typed cell journal. On failure, every captured cell is restored and verified before Summit reports successful rollback.
- Structural operations are treated more conservatively: because rows, columns, formulas, names, validations and formatting cannot yet be atomically journalled, any failure after structural mutation forces recovery Save As.

## Performance follow-up

Initial testing showed an unacceptable increase from roughly 20 seconds to roughly 80 seconds. Two synchronous safety-path costs were removed without weakening the failure contract:

- import rollback capture now uses native bulk, same-address range copies in a calculation-disabled in-memory workbook instead of reading value, formula and formatted display text through the managed API for every cell;
- live system-message updates append only newly added records at the top of the grid instead of copying, sorting, rebinding and best-fitting the complete accumulated history after every operation.

Debug builds now write `[Performance]` entries for rollback range capture, structure-rule execution and Transactional DB synchronisation. These timings should be retained until representative client-model tests confirm acceptable performance.

## High-risk paths hardened

1. `TransactionalDBSynchroniser.SynchroniseRules`
   - stops after the first failed rule;
   - restores calculation mode and engine state in guarded cleanup blocks;
   - attempts analyser RangeDataSource reconnection independently;
   - records cleanup failures rather than swallowing them;
   - marks the workbook dirty after structural work;
   - forces recovery Save As if a structural attempt or cleanup cannot be verified.

2. `WorkbookStructureRuleManager` and generic `WorkbookManager` row/column resize methods
   - validate target worksheets before mutation;
   - track whether physical mutation began;
   - guard worksheet protection restoration and `EndUpdate` independently;
   - require successful Transactional DB synchronisation;
   - force recovery Save As after a partial structural operation or failed synchronisation.

3. DSA and model imports
   - rent and management/service-cost imports preflight their complete named-range contract;
   - destination cell values are journalled and verified on rollback;
   - the legacy stock-condition import now returns an `AbovoTransaction` and no longer hides errors through `On Error`;
   - DSA folder import stops after any rejected scheme that leaves the model recovery-required;
   - DSA source close and target state restoration are independently attempted;
   - the seven-sheet development resize now requires a successful Transactional DB synchronisation result.

4. Interactive changes, save/close and reporting
   - failed single-cell/history rollback is verified and escalated to recovery-required when necessary;
   - DataInterfaceTemplate publishes import and grid-command transaction results to the system message manager;
   - unhandled SpreadsheetControl errors are associated with the correct open model where possible;
   - save, Save As and validation outcomes are published;
   - message copy/save/email failures are themselves recorded;
   - active-interface refresh failures are reported without preventing other interfaces from refreshing.

## Structural-operation performance protection

Live Development Details testing showed approximately 87 seconds for a five-record insertion, of which approximately 70 seconds was spent resizing the 14 `TransCopy_DevptSingle_A` to `_N` mirrors. Disposable in-memory reproductions through both a headless `Workbook` and a hidden `SpreadsheetControl` completed the same seven source-sheet column insertions and 14 mirror insertions in approximately 21 seconds. A full calculation before the insertion increased a single mirror operation but did not reproduce the live 70-second total.

`ModelSafetyManager` now provides a model-scoped, reference-counted bulk-workbook-mutation guard. `WorkbookStructureRuleManager` and `TransactionalDBSynchroniser` enter that guard around their physical workbook mutations, and `MainModelViewer` suppresses only its interactive per-cell logging and recalculation while the guard is active. A live retest without the worksheet viewer remained at approximately 87 seconds, proving that the viewer was not the principal cause; the guard remains as preventative protection against redundant interactive calculations when the viewer is open. Transaction-level success/failure reporting, dirty-state handling, Transactional DB synchronization, snapshot invalidation and recovery Save As protection remain in force.

The GroupInterfaceTemplate sidebar no longer calculates the workbook during construction, navigation, history notifications or automatic display refresh. Previously, `RefreshSummaryData` called `CalculateFull`; the resulting calculation-completed event scheduled an automatic refresh, which called `CalculateFull` again and could sustain a background calculation/refresh feedback loop. Automatic and navigation paths now render the current workbook state only, while the user-selected sidebar Refresh button alone calculates. Per-mirror diagnostics separately report structural shift and template-fill times for validation.

After removing the feedback loop, a live five-record test fell to approximately 38 seconds overall, with approximately 25 seconds in Transactional DB synchronization. Phase timings attributed approximately 21 seconds to the 14 structural shifts and approximately 3.4 seconds to template fills. Read-only inspection confirmed that Transactional DB uses A:BZ, all worksheet-backed defined names end by BV, it contains no tables, queries or pivot tables, and all 148 shapes are free-floating within A:J. A full-workload in-memory comparison rejected full-row insertion as an optimization: all 14 full-row shifts took approximately 15.4 seconds versus approximately 14.9 seconds for the existing partial used-range shifts. The partial-range implementation was therefore retained.

Manual performance validation must measure a five-record Development Details insertion before and after any explicit sidebar refresh, capture the new per-mirror phase timings, and confirm the Check Sheet, Transactional DB mirrors, analyser data sources and Excel round-trip remain valid.

## Remaining risks and recommended next stages

### Priority 1 - atomic structural rollback

Current structural protection preserves the original workbook by forcing Save As; it does not reverse an arbitrary partial multi-sheet row/column operation. A future V2 transaction journal should capture and restore:

- inserted/deleted row and column topology;
- affected global and local defined-name formulas;
- cell formulas, values, formats, validation, comments and protection;
- worksheet visibility/protection state;
- Transactional DB rule state and analyser RangeDataSource bindings.

Commit should occur only after workbook calculation and Check Sheet validation succeed. Rollback should run in reverse order and be verified before the original workbook can be overwritten.

### Priority 1 - DSA scheme-list rollback

If `UpdateList` fails after changing `ImportedSchemes`, `Appraisal1` or `Appraisal2`, the imported worksheets are removed but the list/named-range changes are not structurally reversed. The recovery-save rule now prevents damage to the original; a dedicated DSA import transaction should add true rollback.

### Priority 2 - legacy live entry points

The following active-source routines still contain legacy exception handling and should be migrated or formally retired after confirming callers:

- `FormulaGeneration.ExecuteFormulaGeneration` uses `On Error GoTo` and its handler does not populate the supplied transaction. Much of the implementation is commented, and `ActiveSheet` can remain `Nothing`.
- `WorkbookManager.DevExpressInsertRows` and `WorkbookManager.InsertColumn(IWorkbook, ...)` retain legacy structural logic. Repository search found no production call sites for these overloads; archive or remove after confirmation.
- `SpecificRowColumnEvents.HAColumnsInsertion` retains older multi-sheet code with `On Error Resume Next`. Current event dispatch uses rule-based handlers instead; archive after confirming no reflection/XML construction path depends on it.
- `AbovoBP` retains older row/column wrappers using `On Error`. Trace runtime callers and route them through `WorkbookStructureRuleManager` or the hardened `WorkbookManager` API.

### Priority 2 - source-model cleanup consistency

All import routines should share a disposable import-session abstraction that opens a source workbook read-only, tracks its model ID, disables events/macros as appropriate, and guarantees close without masking the primary exception.

### Priority 3 - global crash evidence

Add an application-level unhandled-exception reporter that writes a small emergency log outside the workbook and mirrors it to `SystemMessageManager` when available. It must not suppress process termination or allow continued workbook mutation after an unknown exception.

## Required manual validation

Use a disposable copy of `Z:\Sandbox\TestFileClean.xlsb` and verify:

1. Successful and deliberately failed insert/delete for each rule-based assumption range.
2. Capital Grant and Development Details changes, including Transactional DB formulas and Check Sheet results.
3. DSA single, folder, consolidated and template imports, including a forced mid-operation failure.
4. Rent, management/service-cost and stock-condition imports with valid and invalid source files.
5. A forced Transactional DB synchronisation failure: normal Save must redirect to a different-path recovery copy.
6. Successful save, close, Excel/VBA reopen, recalculate, save, and Summit reopen.
7. Analysis V1/V2 RangeDataSource reconnect and snapshot/comparison invalidation after a structural change.
8. System-message HTML/text save and email fallback.

The original master must remain untouched throughout these tests.
