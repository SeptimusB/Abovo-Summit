# Check Sheet calculation matrix — 24 September 2026

Item 80 follow-up. Diagnosis only: no production changes, no new build and no workbook save.

## Controlled lifecycle

Each candidate ran in a separate x86 native process using the existing Release 2.97 assemblies. Each loaded a new byte-for-byte private copy of `C:/Sandbox/Insert Comp/AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb`, opened the actual Funding interface, changed G82 from 2700 to 3000 through ChangeManager, navigated to the actual Check Sheet, then used History Undo. No earlier candidate's calculation state could carry into the next process.

All five starts had the same captured values and SHA-256 digest:

`D407525EFC3E1F097C4D81B31EAC0916E6EEE11B2E57ABBF2B15B0212FF6158C`

This digest covers all 57 rows of `Outputs_CheckSheet` in columns B/D/E and all 40 annual cells Q:BD of Transactional DB check rows 2678/2746, with addresses/labels. At every start G82 was restored to 2700 but Check Sheet rows 23, 33, 37 and 39 still showed Check. Digests establish equality of these observed cells, not every internal engine cache; fresh processes and identical replay control the wider lifecycle.

The ordinary navigation path initially left the header warning hidden. Before Undo, the fixture separately published the existing failed result through `RecordIdleCheckSheetResult` to put the heading into a known red state. This was diagnostic notification plumbing, not a claim that normal navigation publishes it, and did not calculate anything. After each candidate it sampled workbook/grid/header before explicitly publishing and refreshing.

## Results

| Candidate immediately after Undo | Time including temporary engine setting/restoration | Remaining failed rows | Result |
|---|---:|---|---|
| One deferred-worksheets pass | 1,309 ms | 23, 33, 39 | Clears TDB Cashflow only |
| Two deferred-worksheets passes | 2,536 ms | 23, 33, 39 | Exactly the same captured end state as one pass |
| ChainBased Workbook.Calculate, DontCalcTDBS=False | 1,522 ms | 23, 33, 37, 39 | Does not restore the observed checks |
| Recursive Workbook.Calculate, DontCalcTDBS=False | 11,002 ms | 23, 33, 37, 39 | Same captured end state as chain incremental |
| Recursive CalculateFull, DontCalcTDBS=False | 11,434 ms | None | Exact captured result match to the original loaded workbook |

These are one bounded run per candidate, not repeated performance estimates. Different workloads or prior navigation histories may produce different results. In particular, the earlier diagnostic applied its extra deferred pass **after** navigating away and back; its success must not be generalized to a standalone Undo repair. This matrix removes that ambiguity.

The full-control result and the original loaded result share this exact digest:

`1D4CF788AA31A191A61327032E275892D16455ED4AF616E1CB9D11DAA7081E73`

Engine, calculation mode and DontCalcTDBS were restored for every case. Sixteen diagnostic assertions passed per run, 80 total. They cover setup, read errors, reproduced failures, accepted edit/Undo, option restoration, warning/result agreement after publication and restored G82; they do not incorrectly assert that every calculation candidate solved the failure.

## Warning and visible-grid transitions

- All five initial navigation results had four failed checks but a hidden company-heading warning and zero status-change events.
- Diagnostic publication raised one status event and displayed the red warning.
- The four unsuccessful candidates kept that warning red, consistent with their still-failed workbook checks.
- Recursive CalculateFull cleared the workbook checks but left the header red and four visible grid cells stale. Explicit publication raised event two and cleared the heading; explicit DIT refresh removed all four display mismatches.

Calculation, soft-warning publication and mapped-grid refresh are independent responsibilities. Even a successful full calculation alone does not perform the latter two.

## Bounded TDB formula trace

The first materially nonzero Cashflow-check cell was `Transactional DB!U2678`:

- Formula: `=U2677-U2676`.
- Before and immediately after Undo: U2678 = -74.919196337894391.
- U2676 = 74.919196337894391, formula `=SUMIF($F$164:$F$2496,"Cash",U$164:U$2496)`.
- U2677 = 0, formula `=INDEX('Cashflow detailed'!$BT$9:$BT$48,U$5)`.
- After one deferred pass: U2676 = -5.0931703299283981E-11, U2677 remains 0, U2678 = +5.0931703299283981E-11. The Check Sheet's integer-rounded check therefore clears.

That demonstrates a stale TDB subtotal being refreshed, not a repaired formula. Other checks still failed because recalculating these selected worksheets did not refresh every producer they consume. The full source-to-output dependency path and the role of volatile/custom-function dirty marking were not exhaustively traced in this bounded matrix.

## Important worksheet-scope clarification

Installed DevExpress 25.2 source distinguishes worksheet/range calculation from custom workbook-chain calculation:

- `DocumentModel/CalculationChain.cs:946` routes CalculateWorksheet through CalculateRangeCore; line 959 creates ChainRangeCalculator directly.
- `DocumentModel/Formula/Calculation/Calculators/ChainRangeCalculator.cs`, IsCalculated, treats formula precedents outside the requested range as already calculated, allowing their cached values to be used.
- The custom workbook calculation path separately instantiates ChainCustomCalculator (`DocumentModel/CalculationChain.cs:924`). Summit's DontCalcTDBS callback affects that path.

Therefore **do not describe standalone Worksheet.Calculate as directly suppressed by DontCalcTDBS**. Its key limitation here is worksheet scope and consumption of external cached values. The flag suppresses selected cells during the custom workbook-chain path. Neither mechanism proves a universal one-sheet dependency-depth limit.

Official documentation describes [Worksheet.Calculate](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Worksheet.Calculate) as worksheet-scoped, [Workbook.Calculate](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.IWorkbook.Calculate) as recalculating marked cells, and [CalculateFull](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.IWorkbook.CalculateFull) as calculating regardless of those markings. The [calculation-process guide](https://docs.devexpress.com/OfficeFileAPI/400926/spreadsheet-document-api/formulas/calculation-process) distinguishes the chain-based custom-service path and recursive engine. Current API pages display version 26.1; the dispatch details above were verified against the installed 25.2 source. No vendor implementation is copied into this report.

## Evidence and preservation

Fixture: `Tools/CheckSheetCalculationMatrix297.cs`. Each result.log contains every captured Check Sheet B/D/E cell, both TDB check rows and stage summaries.

| Mode | Evidence directory beneath obj/ClientReportTests |
|---|---|
| deferred1 | d6aa5a4b8dce42eabed847427413e0ad |
| deferred2 | 95c372c8f1b74f69a1e51e749313a4c6 |
| chain | 8362e7a3d6f74e6bbd465e679743d779 |
| recursive | bad4929a5a5946a18031e72df1569c3a |
| full | 010d0951b879446cbca7dfbbc3c0b1f4 |

All processes exited zero. The runner verified the original source hash before/after each run:

`30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47`

The existing Release executable SHA-256 was `464B3EE68633F72BBDAE91E491F3B5B7D9B404A3150A002429435DCBF223262C`. Its assembly file version logs as 1.0.0.0; Summit's test-release number is separately defined as 2.97 in AbovoApp. No Debug/Release files, user app sessions, original workbook, master or production calculation/status code were changed.

## Scoped conclusion

For this exact AGL input/Undo lifecycle, none of the tested cheaper candidates is a substitute for the successful recursive full control. A useful next design must also deliberately publish the resulting soft-warning state and refresh the visible grid. This diagnostic does not select or implement that policy, prove full-workbook financial parity, or rule out a more narrowly scoped dependency-safe calculation that has not yet been tested.
