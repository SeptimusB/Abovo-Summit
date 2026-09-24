# Funding / Check Sheet diagnostic — 24 September 2026

Status: investigation only; no production calculation or warning policy changed.

## Scope and preservation

- Exact reported input: `Funding Assumptions!G82`, 2700 -> 3000, then History Undo.
- Source: `C:/Sandbox/Insert Comp/AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb`.
- The native Summit 2.96 Release model and actual Funding/Check Sheet interfaces were exercised against a disposable copy. The host was invisible/non-activating; no physical mouse/keyboard input was generated.
- All edits went through `ModelChangeManager`; Undo used the real history service. Nothing was saved, original input was restored, and the runner verified the original workbook SHA-256 unchanged. Working-2.96 was not touched.
- Fixture: `Tools/CheckSheetFundingDiagnostic297.cs`.
- Evidence: `obj/ClientReportTests/ee4d85b1ea634daeb4426353b5329357/result.log` (exit 0; seven setup/edit/restore assertions).

## Conclusion

The exact reported behaviour is reproduced. This is not evidence of a fixed dependency-depth limit. There are three separate layers:

1. **Workbook calculation scope/deferral:** Undo restores the input but normal sheet-level calculation leaves cached Check Sheet results stale. Navigation's staged calculation clears most results but leaves the Transactional DB Cashflow check stale. An additional deferred-sheet pass cleared that last result in this run.
2. **Warning publication:** normal calculation/navigation can change Check Sheet values without publishing `CheckSheetStatusChanged`, so the company-header warning remains unchanged.
3. **Grid refresh:** the Check Sheet watcher publishes the warning after its full calculation, but does not refresh the already-visible mapped grid; the header and workbook can therefore be current while the table still shows old values.

The UI is not inventing the first stale result: immediately after Undo its 57 rows matched the cached workbook values exactly. A display-only refresh changed nothing.

## Observed sequence

| Action | Time | Worksheet result | Header / visible grid |
|---|---:|---|---|
| Open private workbook | not timed here | Balanced | Warning false |
| G82: 2700 -> 3000 in Funding | 230 ms | Check Sheet still cached balanced | No status event |
| Navigate to Check Sheet | 2,553 ms | Four failures: rows 23, 33, 37, 39 | Grid matches; header still hidden; zero status events |
| History Undo, Check Sheet visible | 368 ms | Input 2700, same four stale failures | Grid matches stale workbook |
| Display-only RefreshData | 9 ms | Same four failures | No improvement |
| Navigate away and back | 2,253 ms | Only row 37, TDB Cashflow, remains failed | Grid matches; no status event |
| Check Sheet.Calculate, normal deferral on | 2 ms | Row 37 remains failed | No improvement |
| CalculateDeferredWorksheets | 1,164 ms | Balanced | Grid still says Check at row 37 |
| Chain-based Workbook.Calculate, deferral off | 159 ms | Remains balanced | Grid still stale |
| Chain-based CalculateFull, deferral off | 9,407 ms | Remains balanced | Grid still stale |
| Recursive CalculateFull, deferral off | 6,730 ms | Remains balanced | Grid still stale |
| Display-only RefreshData | 13 ms | Balanced | Grid catches up |
| Second G82: 2700 -> 3000; full Check Sheet watcher | 12,781 ms watcher | Three failures: 23, 33, 39 | Warning event 1; heading shows; visible grid remains OK at those three rows |
| Second Undo | 576 ms | Input 2700, three cached failures remain | Warning remains; grid now matches stale workbook |
| Full Check Sheet watcher | 11,366 ms | Balanced | Warning event 2; heading clears; visible grid still says Check at three rows |

The first navigation's TDB Cashflow failure (row 37) is transient: the later full-workbook control calculation of the same 3000 input reports only rows 23/33/39. Thus there is a stale intermediate-result problem as well as delayed warning publication. This is an internal engine comparison, not independent Excel/accountant validation.

These are single-run diagnostic timings, not a benchmark comparison. The incremental/full/recursive calls were sequential, after the additional deferred pass had already cleared row 37. They do not establish which would be the cheapest independent repair from identical stale state.

## Source trace

- `Services/Calculation Engine and Custom Functions/CustomCalcEngine.vb:50`: `DontCalcTDBS` handles/skips Transactional DB, TDB Comparison and Check Sheet cells during customised chain-based workbook calculation. It does not intercept standalone `Worksheet.Calculate`, which follows the separate range-calculator route. At line 75 the deferred pass temporarily disables skipping and calculates those worksheets in sequence, then restores the flag.
- `Services/FileManager.vb:715`: normal HA models register that service with deferral enabled.
- `Services/DataService/ChangeManagerV2.vb:315`: interactive change uses `CalculateWSs`; at line 528 history calculates the restored source worksheets plus active worksheets, not an authoritative whole-model check.
- `Services/EngineManagement.vb:193`: normal active-sheet calculation refreshes registered UI objects and raises CalculationCompleted, but does not publish Check Sheet state. At line 387 navigation's `CalcFile` executes its workbook/deferred-sheet stages.
- `Interface/User Interface/DataInterfaceTemplate.vb:662`: reactivation controls navigation calculation and refresh. At line 738 RefreshData reads current workbook data. At line 11297 Undo/Redo history notification refreshes the interface; it does not itself repair deferred workbook caches.
- `Interface/User Interface/GroupInterfaceTemplate.vb:65`: the heading subscribes to `CheckSheetStatusChanged`. The separate calculation/history callbacks refresh the sidebar, not that state.
- `Services/FileManager.vb:1199`: `RecordIdleCheckSheetResult` publishes accepted warning-state transitions. At line 1413 `ReadCheckSheetValidation` deliberately reads cached results; the caller must first establish their currency.
- `Services/CheckSheetWatch.vb:125`: the watcher temporarily uses Recursive + CalculateFull with deferral off, reads the result and publishes it, but does not call the active UI refresh path. In contrast, `Services/IdleIntegrityManager.vb:339` and `:348` explicitly call `RefreshAfterDeferredCalculation`.

## DevExpress semantics (official documentation plus installed 25.2 source)

`Worksheet.Calculate` is scoped to the specified worksheet, not a recursively fresh whole-workbook check. `Workbook.Calculate` recalculates marked cells; `CalculateFull` disregards those markings. A chain-based full rebuild also reconstructs the dependency tree. These are different scopes/dirty-state policies, not a configurable maximum dependency depth. See [Worksheet.Calculate](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Worksheet.Calculate), [IWorkbook.Calculate](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.IWorkbook.Calculate), [CalculateFull](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.IWorkbook.CalculateFull) and [CalculateFullRebuild](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.IWorkbook.CalculateFullRebuild).

The official [calculation-process documentation](https://docs.devexpress.com/OfficeFileAPI/400926/spreadsheet-document-api/formulas/calculation-process) explains that chain-based calculation orders dependencies and marks affected cells, whereas recursive calculation evaluates on demand. It also documents that the custom calculation service can cancel individual cells and applies to the chain-based engine.

Current web API pages describe 26.1. The pertinent mechanics were also inspected in the locally licensed 25.2 source under `C:/Program Files/DevExpress 25.2/Components/Sources/Win/DevExpress.XtraSpreadsheet/DevExpress.Spreadsheet.Core/`:

- `DocumentModel/Formula/Calculation/Calculators/ChainRangeCalculator.cs`, `IsCalculated`: formula cells outside the requested range are treated as already calculated. Their cached results may therefore be consumed by a sheet/range calculation.
- `DocumentModel/Formula/Calculation/Calculators/ChainCustomCalculator.cs`, `CalculateCellCore`: a handled calculation uses the service's supplied value, initially the cached value. Summit's skip callback does not supply a replacement.
- `API/Native/Workbook/IWorkbookImpl.cs`: worksheet/workbook/full/rebuild calls dispatch to different native calculation operations.
- `DocumentModel/CalculationChain.cs`: `CalculateWorksheet` calls `CalculateRangeCore` and uses `ChainRangeCalculator`, whereas workbook calculation checks for and invokes the custom calculation service. A sheet calculation's stale external precedents should not be misdescribed as the custom service skipping that sheet calculation.

No proprietary vendor source is copied into this report. There is no evidence here that the dependency tree is structurally corrupt or that it has a one-level traversal cap. The exact upstream dependency/order that makes the first staged pass leave row 37 stale was not exhaustively traced; that should not be claimed proven.

## Independent follow-up for functional test 80

The [fresh-process matrix](Check_Sheet_Calc_Matrix_2026-09-24.md) now compares five paths from the same observed broken-and-undone state, without the intervening second navigation in the first sequence above. One or two deferred passes clear only the transient TDB Cashflow failure; three other stale failures remain. Both chain-based and recursive incremental workbook calculation leave four failures. Recursive `CalculateFull` alone restores all sampled Check Sheet B/D/E and both TDB check rows exactly to their original loaded values. This result withdraws any implication that repeating a cheap deferred pass is a sufficient repair for this case.

Calculation still does not publish the warning or refresh the grid automatically. Explicit publication and display refresh bring those into agreement afterwards. The [prospective refresh design](Check_Sheet_Refresh_Design_2026-09-24.md) also identifies the revision-order hazard: normal edits and Undo calculate before their final history/dirty revision is committed, so a new `CalculationCompleted` subscriber alone would be an unsafe place to certify fresh results.

## Next decision, not implemented

Any repair should deliberately separate (a) which calculation establishes sufficiently current balances, (b) publishing their soft-warning state at a successful operation-completion boundary, and (c) refreshing the visible Check Sheet. Full calculation is the only independently tested candidate that cleared this lifecycle; it must not silently become the cost of every ordinary edit or Undo. The precise narrower safe dependency boundary remains unproven. Debug and Release 2.97 were not rebuilt or overwritten during this follow-up, and no production calculation policy changed.
