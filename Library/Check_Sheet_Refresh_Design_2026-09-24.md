# Check Sheet publication and refresh design

Date: 24 September 2026

Follow-up: the historical proposal below is retained. Approved implementation and validation are tracked separately in [2.98 coherence report](Check_Sheet_Coherence_2.98_2026-09-24.md); the status and source line numbers below describe the original inspection, not that later implementation.

Status: **Proposed, not implemented.** This is a bounded source inspection for functional test 80. It is not a successful runtime reproduction, a change to calculation policy, or a declaration that the public-heading/Transactional DB report is resolved. No production source, workbook, application settings or executable was changed for this inspection.

## Finding

Calculation, visible Check Sheet data and the public soft-warning state have separate completion paths. Normal calculation refreshes interfaces without publishing Check Sheet warning state; the dedicated watcher publishes warning state without refreshing mapped Check Sheet values. A repair should join these responsibilities at a safe completion boundary without implicitly requesting another expensive calculation.

| Existing path | Calculation responsibility | Visible data refresh | Check Sheet warning publication |
| --- | --- | --- | --- |
| `CalcEngine.CalculateWSs` | Registered worksheets, or delegates to `CalcFile` when multiple objects are active | Refreshes registered objects, then raises `CalculationCompleted` | None |
| `CalcEngine.CalcFile` | Workbook calculation and applicable deferred-worksheet pass | Refreshes registered objects, then raises `CalculationCompleted` | None |
| `CheckSheetWatch.RunCheck` | Full workbook calculation with the existing temporary engine/mode/deferral settings | No mapped-grid refresh | Reads the cached override-aware result and calls `RecordIdleCheckSheetResult` |
| Full Integrity | Calculates when pending, then validates the current revision | Refreshes interfaces when it performs calculation | Publishes the brief Check Sheet result before later integrity stages |
| History Undo/Redo | Restored source worksheets plus the usual active calculation | Refreshes during calculation and again after history completion | None |

Source anchors, valid at this inspection:

- `Services/EngineManagement.vb:193` (`CalculateWSs`), `:387` (`CalcFile`), `:159` (`RefreshObjsData`).
- `Services/CheckSheetWatch.vb:125` (`RunCheck`).
- `Services/IdleIntegrityManager.vb:333` (calculation stage), `:398` (Check Sheet validation).
- `Services/FileManager.vb:1199` (`RecordIdleCheckSheetResult`).

## Calculation scope is not a freshness certificate

`ReadCheckSheetValidation` reads cached worksheet results; its caller must establish freshness first (`Services/FileManager.vb:1413`). Its override handling and distinction between financial imbalance and formula errors must remain authoritative.

The installed DevExpress 25.2 source inspection also establishes an important distinction. `CalculationChain.CalculateWorksheet` reaches `CalculateRangeCore` and `ChainRangeCalculator`, bypassing `ChainCustomCalculator`. A standalone `Worksheet.Calculate` therefore does not use the customised workbook-chain suppression route, and can read existing cached values outside the requested worksheet. Merely calculating Check Sheet does not establish that its upstream dependencies are current.

`CustomCalcEngine.DontCalcTDBS` suppresses Transactional DB, TDB Comparison and Check Sheet through `OnBeginCellCalculation` during customised workbook-chain calculation. It must **not** be described as suppressing a standalone `Worksheet.Calculate` call. The explicit deferred-worksheet pass temporarily clears the flag and processes Transactional DB, valid comparison data and Check Sheet (`Services/Calculation Engine and Custom Functions/CustomCalcEngine.vb:50`, `:75`). This distinction is a source finding, not a new calculation policy or a runtime parity claim.

## Edit and history completion order

Undo/Redo currently performs this sequence:

1. Apply the requested cell snapshots with `IsApplyingHistory = True`.
2. Calculate restored source worksheets and the active interface worksheets; the usual calculation path can refresh controls and raise `CalculationCompleted` here.
3. Verify the workbook still contains every intended snapshot; roll back on failure.
4. Clear `IsApplyingHistory`, update the Undo/Redo stacks and record the history log.
5. Call `MarkUserChange`.
6. Raise `HistoryChanged`; DITs then perform their post-history refresh.

See `Services/DataService/ChangeManagerV2.vb:446`, `:528`, `:629`, and `Interface/User Interface/DataInterfaceTemplate.vb:11297`.

Ordinary single-cell edits similarly write and calculate before committing their history and calling `MarkUserChange` (`ChangeManagerV2.vb:311`). `MarkUserChange` increments the user revision and assigns `IsDirty = True`; the dirty setter increments `CalculationRevision` even when the model was already dirty (`Services/FileManager.vb:310`, `:446`). Consequently, a result published synchronously inside `CalculationCompleted` would be associated with a pre-commit revision. Updating the warning there without explicit operation provenance risks accepting provisional results, reporting rolled-back work, or immediately making an accepted result appear stale.

`ChangeInProgress` alone is not a complete ordinary-edit guard: it tests `IsApplyingHistory` or an explicit `ActiveGroup`, whereas an automatic single-cell group is local to `ProcessResolvedChange` (`ChangeManagerV2.vb:42`, `:305`). An explicit successful operation-completion boundary is required.

## Proposed bounded repair

1. Introduce one soft-result publication operation, separate from any calculation request. It accepts a freshly obtained validation result only with explicit successful calculation-scope and model/operation-revision provenance. Distinguish **accepted** from **warning changed**; the current Boolean returned by `RecordIdleCheckSheetResult` cannot distinguish an unchanged result from a rejected stale result.
2. For edit/Undo/Redo, stage any candidate result until the owning operation has successfully completed and committed its final revision. Discard it on rollback, subsequent writes, closure, wrong model or mismatched operation. Do not relabel an arbitrary old cache with the latest revision, and do not alter revision semantics merely to make the check pass.
3. Emit a targeted Check Sheet values-refreshed notification for every accepted fresh result, including balanced-to-balanced and unbalanced-to-unbalanced results. Existing mapped Check Sheet controls can then update their cached values and conditional colours without rebuilding the DIT or requesting another calculation. Warning-transition events alone are insufficient.
4. Preserve existing warning behaviour: override-aware soft financial notices, distinct formula errors, normal Save availability, default continuation of recovery autosaves, and persistence only after successful Save/Save As or against a recovery copy's own path.

The exact operation token or depth mechanism remains an implementation decision requiring review and regression tests. This document does not authorise a broader dependency-calculation redesign.

## Refresh and re-entrancy safeguards

- `ReadOnlyMappedTableGrid.RefreshData` hides its active editor before refreshing cached text and colours (`Interface/Custom Controls/ReadOnlyMappedTableGrid.vb:183`). Defer a new notification-driven refresh while an override editor contains uncommitted input or is posting; do not silently commit or discard it. Preserve scroll, selection and column widths.
- Its existing history handler refreshes only if the changed worksheet list contains the mapped source sheet (`ReadOnlyMappedTableGrid.vb:438`). A change elsewhere that affects Check Sheet is not covered by that subscription alone.
- DIT refresh already defers during workbook posting and suppresses posts while refreshing grids (`DataInterfaceTemplate.vb:738`). Preserve those guards and the history snapshot verification.
- Do not use `RefreshAfterDeferredCalculation` as a generic repaint helper: it clears calculation-dirty flags and marks the navigation generation current in addition to refreshing all interfaces (`Services/EngineManagement.vb:185`). A targeted refresh must not certify unrelated calculations.
- Do not clear `ResultsPending`, rebuild/formula-preflight flags, dirty state or user revision merely because a soft Check Sheet result was published. The watcher deliberately retains these distinctions.
- Isolate presentation subscribers individually so one disposed or failing interface cannot prevent the other headings/grids receiving an accepted result. Notifications must not escape into a completed workbook transaction.
- Keep all workbook reads and native interface work on their owning thread. Do not add `DoEvents` or accept input between calculation and acceptance. Coalesce deferred presentation work and reject stale queued work after a new edit or model close.
- Public headings already consume `CheckSheetStatusChanged` (`GroupInterfaceTemplate.vb:65`; `FileInstanceInterface.vb:64`). Retain that narrow state-change responsibility; a separate values-refreshed notification covers unchanged-status data refresh.

## Prospective validation

All workbook tests must use disposable copies. The source-only inspection has not run these new cases.

1. Break a value elsewhere and warm-switch to Check Sheet: the computed row, public warning and FileInstance summary agree.
2. Correct, Undo and Redo from both the DIT and separate History window: values, rows, history state and warning reflect the final committed operation.
3. Complete the watcher with Check Sheet still visible: values and conditional colours update without switching tabs.
4. Keep the warning active but move the failure to a different row: the visible details still refresh despite an unchanged Boolean warning.
5. Keep a Yes/No override editor pending: a background completion neither commits nor discards it, and a safe refresh occurs after the edit finishes.
6. Apply an accepted override: it removes only the soft imbalance; a formula error remains distinct and is not hidden by the override.
7. Force a failed Undo/rollback or a revision-changing callback: no provisional result clears the warning or certifies stale data.
8. Open two models: only the correct model's headings and existing Check Sheet controls update; disposal and delayed notifications are safe.
9. Break, check, discard without saving and reopen: no session-only finding is persisted to the original path; successful Save retains its existing persistence behaviour.
10. Confirm publication/repaint requests no extra calculation, preserves all pending/rebuild flags and user revisions, and retains the default soft-warning/recovery policy.

The original item 80 remains open pending an agreed repair and functional acceptance. Calculation-cost policy and the separate Transactional DB stale-result question require their own evidence; this publication design does not declare either resolved.
