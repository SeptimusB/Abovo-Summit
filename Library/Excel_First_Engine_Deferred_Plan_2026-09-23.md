# Deferred proposal: Excel-first workbook engine

Historical status: deferred at the user's request on 23 September 2026. **Superseded by explicit staged-implementation approval on 24/25 September 2026.** Prefer compatible Excel automatically, respect existing macro security, retain configurable DevExpress fallback. Current scope, tested foundation and unimplemented release gates are in `Library/Excel_DevExpress_Engine_Stage1_2026-09-25.md`. The remainder preserves the earlier proposal, not the current authorization boundary.

## Proposal discussed

Abstract workbook operations behind a common service, with installed compatible Microsoft Excel automation as the preferred engine and DevExpress as the fallback. Retain Summit's DevExpress interface, charts, navigation and change-history behaviour. Many clients have licensed desktop Excel, but availability, compatibility and macro trust must be checked rather than assumed.

Keep an owned Excel instance and the workbook open for the session. Do not save, open Excel, calculate, save and reload Summit after every edit. Batch value and formatting transfers; avoid per-cell COM traffic where practical.

The service boundary would cover reads/writes, names, calculation, formatting and validation, structural commands, save/save-as and recovery. Each model must have one authoritative engine. A display cache must not independently calculate, persist stale results or replace formulas with calculated constants.

## Current coupling verified during assessment

- Services/EngineManagement.vb calculates DevExpress worksheets/workbooks and refreshes registered interface objects.
- Services/DataService/ChangeManagerV2.vb holds a DevExpress IWorkbook. Preserve typed changes, history, dirty-state management and rollback while adapting the storage boundary.
- Interface/User Interface/BPIncomeExpenditureAnalyserV2.vb binds to DevExpress RangeDataSource objects.
- Services/DataService/DataManager.vb UpdateLocks reads worksheet fill patterns and formatting directly. Preserve the established fill-based editability contract; Funding date entry must refresh adjacent-cell permissions as well as values. Do not weaken these rules.

## Suggested next experiment, subject to approval on resumption

Use a disposable copy of C:/Sandbox/Insert Comp/AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb. Prototype a persistent Excel session outside production and measure ordinary edits, Funding date entry/unlocking, Check Sheet calculation, ten-column insertion, and save/recovery. Compare complete edit-to-refreshed-interface time, including transfers and formatting, not calculation alone. Validate formulas, names, VBA compatibility, editability and financial results.

Use a Summit-owned Excel instance without disturbing the user's Excel windows. Provide a correctly managed STA automation worker, message handling, bounded busy-call retries and controlled cleanup. Do not bypass macro security. Select DevExpress at opening when Excel is unavailable; mid-operation failure requires explicit checkpointed recovery, not a silent engine swap.

Excel speed remains unproven for this proposed workflow. Native worksheet calculation can be multithreaded, but VBA UDFs can remain a bottleneck. Prefer dependency-aware calculation for ordinary edits and rebuild only when required and validated.

## Existing evidence and boundaries

See Library/Spreadsheet_Engine_Trial_Assessment_2026-09-23.md for the isolated Aspose/DevExpress I/O trial and its unresolved compatibility gates. Those raw I/O results do not prove end-to-end Excel performance. The Excel-first assessment made no production code or workbook changes. This note is a handover, not a release, commit or implementation plan approval.

Official references reviewed:

- https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions
- https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-improving-calculation-performance
- https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=visualstudio
