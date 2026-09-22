# Structural repair trial 2.55 — Ready to test

User approved the repair sequence on 22 September 2026: Development, incorrect physical axes, copied new-row inputs and linked-range updates identified in `Master_VBA_and_DIT_Structural_Audit_2026-09-21.md`.

## Implementation

- Development identified and multi-year columns shift across all seven VBA worksheets before any template is copied. Identified insertion retains the VBA first-one/remainder sequence. Leading ten identified / three multi-year columns and the trailing sentinel cannot be deleted. Parsed 3-D references are maintained.
- All seven vertical ranges identified in the preceding audit now expand by worksheet rows, irrespective of interface transposition.
- Simple row rules, Journals and Stock Conversion clear new editable constants. Existing inputs remain once; formula-backed unlocked cells retain their formulas. This follows the established native insertion policy, not the VBA's blanket clearing of unlocked formulas. Locked template defaults are retained. No master formulas are regenerated.
- Service Charge uses IR_ServChg_01 as its input boundary, maintains the one-row-larger Rep_ServChg_02, and expands its five working sheets. Specific Income expands four working sheets; Other Income expands one. The VBA's additional terminal workings template is copied as part of these contracts. All linked mutation precedes TDB synchronisation.
- Joint Venture and both Intercompany Funding column families now use seven/twelve-sheet grouped rules and existing four/six TDB mirrors. Their existing last input remains in its logical position; new following input cells are blank. Legacy single-cell array formulas in JV Interest Received retain array semantics. Unsupported multi-cell/dynamic 3-D array cases fail rather than being flattened.
- Service Charge's two repeated-input column headers, Specific/Other Income category columns, four JV input-row families and three intercompany date-row families now resolve to explicit rules. Companion IR/IC/Rep names are resized from entry-state snapshots, not inferred from the on-screen axis. Interco Loans Received / Investments Made expose the semantic Add/Delete Lines route in XML.
- Known NRCI and NRRIbyCOL tokens route through the semantic service. Legacy insert errors are surfaced before refresh and cursors restored in Finally. No success is reported after a failed mutation.
- Intercompany date-row delete reproduced a DevExpress ChainBasedCalculationLogic / SharedFormula NullReferenceException. Only those three rules temporarily use the supported Recursive engine for their bulk mutation; all three add/delete tests pass with original engine/mode restored. Other operations retain their existing engine.

The 3-D array implementation uses the supported entire-array-range API: [DevExpress 25.2 GetArrayFormulaRange](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Cell.GetArrayFormulaRange?v=25.2). Tests compare array type as well as formula text; this is not a text replacement of formulas.

## Native verification

Fresh disposable copies of the authoritative Blank passed three-record add/delete for: the seven corrected-axis ranges, Service Charge records, Other Income records/adjustments, Joint Venture columns, both Intercompany column families, Cash Journals, Service Charge column variants 01/02, Specific/Other Income category columns, JV rows 01/04, Journals and Stock Conversion. Intercompany date rows 01/02/03 passed after the scoped engine repair. Final published-Debug tests for JV rows 02/03 each pass one-record add/delete, minimum-boundary rejection, all four mirrors, all 1,679 global / 82 worksheet-local names and affected-sheet formula restoration.

Each passing fixture checks physical dimensions, out-of-range deletion rejection, workbook dirty state, calculation settings, worksheet protection/visibility, preservation of a non-zero existing input where an editable constant is available, and return of all **1,679 global defined names** and formulas on the affected worksheets after deletion. These name checks enumerate the global collection, not every worksheet-local name. Associated TDB mirror sizes are explicitly checked where identified (four JV, six Interco, one Cash Journal/Other Income); earlier Journals tests cover its two mirrors. Source hashes are checked by the runners.

Development Blank: three identified and three multi-year columns pass add/delete, protected-leading and sentinel rejection; all 106,041 linked-sheet formula cells restore. Demo: one identified and one multi-year column pass; 106,044 linked-sheet formulas restore, with all fourteen mirrors checked on each Demo run.

## Independent Excel comparisons

References are transcribed from the inspected VBA operation sequence, not generated from the native rule metadata. Excel opens private reference inputs and Summit results read-only with macros/events/link updates disabled; SaveCopyAs creates distinct outputs, then every book closes without saving its input. No VBA procedures are executed.

| Blank operation | Formula cells matched against independent Excel and Excel save-copy/reopen |
| --- | ---: |
| Development identified, +3 | 124,449 |
| Development multi-year, +3 | 123,105 |
| Service Charge records, +3 | 2,734 |
| Specific Income records, +3 | 3,849 |
| Joint Venture columns, +3 | 2,049; array types also match |
| Intercompany loans, +3 | 19,595 |
| Intercompany investments, +3 | 19,595 |
| Other Income records, +3 | 2,520 |
| Specific Income category columns, +3 | 27 |
| Other Income category columns, +3 | 104 |

Demo independent Excel and round-trip comparisons also pass: identified +1, 112,180 formula cells; multi-year +1, 111,732. All 337 VBA module identities/source hashes are unchanged through native save and Excel save-copy in each of these twelve tested cases (ten Blank, two Demo). `obj/structural255-blank-vba.log` and `obj/structural255-demo-vba.log` contain hash-only evidence.

The independent row reference follows VBA's first-row/remaining-rows sequence and second-last source-row template. The native bulk insertion's resulting formulas match in the tested linked-row families. Both use the intentional native policy of retaining formula-backed editable cells. Copying a source column over itself is omitted natively to avoid DevExpress's array-overlap error; Excel-result equivalence is tested.

Funding's eight-column add/delete regression passes with all eleven mirror sizes, ordinary/revolver protections and original formulas/names restored. Both final Debug and Release builds pass; the 3-D parser/array unit fixture passes. Published-build one-row Interco date add/delete also passes, including all 1,679 global and 82 worksheet-local names, six mirrors, minimum-boundary rejection and calculation-state restoration. Structure.xml and the project index parse; diff whitespace checks pass. Final binaries: `bin/Debug/Abovo-summit.exe` and `bin/Release/Abovo-summit.exe`, built 22 September 2026 at 01:01 local, from this checkout with test version 2.55. No Summit UI executable is running at handoff.

The master SHA-256 hashes remain exactly the values recorded in the preceding audit. No client file or source master has been saved. The current branch is `main`; no commit or push has been performed for this approved repair turn.

## Tools and reproducibility

- `Tools/Test-ColumnFamily.ps1` with `ColumnFamilyFixture.cs`: Development; optional Rule, Count, Workbook, Configuration.
- Same runner with `-Fixture LinkedStructureFixture.cs`: row and linked families, populated-input sentinel, mirrors, state and reversal checks.
- `Tools/Test-DevelopmentExcelRoundtrip.ps1`: independent Development reference and saved Excel formula/array comparisons.
- `Tools/Test-LinkedExcelRoundtrip.ps1`: Service/Specific/Other Income, JV and Interco independent references.
- `Tools/Test-Structural3D.ps1`: parser tests, mixed absolute/relative references, names, whole rows, partial-span rejection and legacy single-cell arrays.
- `Tools/Verify-VbaModuleHashes.py`: source-module identity/hash preservation without source extraction to disk.
- Local evidence is in `obj/structural255-*.log` and the output directories printed at the start of each native log. Failed preliminary JV/date tests are retained as diagnostic evidence, not counted as passes; subsequent named retests supersede them. An initial single-sheet Excel harness error was corrected by retaining a PowerShell array rather than indexing a single sheet-name string.

## Acceptance boundary / remaining work

This trial addresses the approved Development/axis/input-copy/linked-range batch. It does **not** settle the separately reported OFA/Repairs anchor/template differences or certify every historical family-specific deletion minimum. Existing damaged client files are not repaired on load.

Native and Excel formula/array equivalence is engineering evidence, not financial acceptance. Still required on copies: visible DIT footer/header controls, inputs after resize, Check Sheet, detailed SOCI/Financial Position, save in Summit -> open/edit using Excel/VBA -> save -> reopen Summit, with the client accountant reviewing financial results. The macro-disabled Excel test cannot establish UDF/VBA execution correctness. No populated-client financial totals are claimed here.

No master/client originals are to be saved or migrated. Prior dirty changes are retained. This document will be updated with actual results before delivery; a build alone does not establish workbook/financial correctness. No commit or push was requested.
