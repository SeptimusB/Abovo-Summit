# Funding insertion investigation and DIT Save trial

## Scope and authority

Test release 2.54 adds DIT Save and Save As and the explicitly approved Funding insertion/deletion repair. The user separately authorised protecting the first ten ordinary-loan columns and revolvers. Source workbooks and the already-expanded client workbook are unchanged; recovery of that existing client file remains a separate step. The subsequent full VBA/DIT audit is documented in `Master_VBA_and_DIT_Structural_Audit_2026-09-21.md`.

Client evidence: `D:/Downloads/Test BP v26_0001 - FormGenRemoved - PopInSummit - funding cols insert (1).xlsb`.
Comparators: authoritative `Library/Demo BP v26_0001.xlsb` and `Library/Blank BP v26_0001.xlsb`.
Excel inspected exact disposable copies read-only, with macros/events disabled, links not updated, and no saves. Source hashes were checked. VBA was read in memory; no raw VBA or passwords were retained.

## Base-model findings

- All three files report model version 26.0001.
- All 285 worksheet names and their order match, as does the inventory of 1,761 defined names. Range addresses differ where the client has populated/expanded the model.
- The complete Version History data matches.
- The Funding_Columns and Summit_Compatibility VBA module sources match exactly.
- Relative to Blank, the client has two additional empty worksheet VBA modules and a Menu_Module change that recalculates the menu sheet before reading its CELL-based parameters. These do not establish a different Funding model.
- The evidence establishes the same base-model family, not byte-identical or comprehensively formula-identical workbooks. Client inputs, other structural edits and its enlarged Transactional DB must be retained.

## Confirmed Funding defect

The client's FacilityNames spans E45:Z45 (22 columns), compared with E45:R45 (14) in the masters. LoanDescRev1 has moved from O47 to W47. The client therefore contains eight added columns in this family.

The master VBA's Run_Insert_Funding_Columns selects the 32 Funding worksheets as a group, inserts immediately before the ordinary-loan template (one column before LoanDescRev1), and then copies the shifted template into the new columns. It subsequently synchronises the Funding Transactional DB mirrors.

Before this repair, Summit's FUNDING_RECORDS used LoanDescRev1 itself as the insertion point and copied from the preceding column. More importantly, ExecuteInsert inserted AND copied each worksheet before moving to the next. Later worksheet insertions therefore rewrote references in formulas that had just been copied.

Read-only scanning found 6,496 formulas in the client's added O:V block with cross-sheet references displaced eight columns right:

| Worksheet | Formula cells containing the displacement |
| --- | ---: |
| Funding Assumptions | 8 |
| Loan Opening Balances | 320 |
| Hidden - Loan Opening Balances | 1,536 |
| Hidden - Open Loan Facilities | 1,544 |
| Hidden - Close Loan Facilities | 1,544 |
| Hidden - Loan Interest | 1,544 |

For example, Hidden - Loan Opening Balances!O9 reads W8 on the drawdown/interest/repayment sheets instead of O8. This is a confirmed structural formula defect. The count is a targeted signature scan, not an exhaustive count of all affected cells.

Macro-disabled Excel full recalculation and targeted hidden-sheet activation did not expose a CircularReference address. This does NOT clear the client report: VBA/UDF execution and the client's precise edit/calculation state were not reproduced. Do not describe the file as financially validated or free of cycles.

## Implemented approved repair

- Match the VBA anchor/template: insert one column before LoanDescRev1 and copy the shifted ordinary-loan template after every linked sheet has been inserted.
- Opt Funding into staged insertion across all 32 targets. Other structural families retain their existing copy order pending separate review/approval.
- Preserve entry worksheet protection/visibility, restore calculation mode/engine, retain existing dirty-state and recovery-required handling, and run existing mirror synchronisation once after the structural work.
- Restrict deletion to LoanDescsOrd, excluding revolvers and the first ten ordinary-loan columns; reject inconsistent boundaries before mutation. Delete Last counts ordinary loans, not all facilities.
- Independent Excel comparison exposed an additional DevExpress 25.2 3-D-reference adjustment gap. A Funding-only helper now uses the public formula syntax tree to preserve 3-D insert/delete/copy semantics, including mixed absolute references and defined names. Unsupported partial-sheet spans and affected array references fail preflight rather than silently corrupting formulas. No string-substitution formula rewrite, hidden metadata, worksheet addition or schema migration is used.
- No attempt is made to repair the already damaged client file automatically.

## Structural verification

Both Debug and Release builds passed. Actual native eight-column insertion/deletion passed on authoritative Blank and Demo copies: 32 target sheets, eleven mirrors, ordinary/facility names, protection/visibility, restored calculation settings, dirty-state, protected-deletion rejection, add/delete name and formula restoration, and the minimum ordinary-loan boundary.

An independent macro-disabled Excel grouped insertion/copy produced a reference result for each master. Summit's expanded result matched **104,014 linked-sheet formula cells per master**. Summit save -> read-only Excel full calculation/SaveCopyAs -> native reopen matched the same formula cells and names. All 337 VBA module source hashes were unchanged in each round-trip. Excel returned no circular-reference address; UDF/VBA execution and financial acceptance remain untested by that check.

Synthetic regression covers 3-D insert/delete, ordinary local references, absolute/relative combinations, ranges, whole-row references, defined names, repeated copies and rejection of a partially selected 3-D sheet span.

Evidence:

- Blank: `obj/FundingStructureTests/ba0f49a3374e42029718b64d6f716cb6`, `obj/funding-structure-blank-3d-test2.log`, `obj/funding-excel-blank-3d-test.log`.
- Demo: `obj/FundingStructureTests/9c1e448d01354174ad5e7e0c844edabf`, `obj/funding-structure-demo-3d-test.log`, `obj/funding-excel-demo-3d-test.log`.
- Final builds: `obj/funding-final-debug-build.log`, `obj/funding-final-release-build.log`; 3-D regression `obj/funding-final-3d-fixture.log`.
- Final Debug native add/delete recheck: `obj/funding-final-debug-structure.log`, `obj/FundingStructureTests/4bf668ea84624c399bfdcb5f9e1f1d44` (passed, including all 1,761 restored names and linked-sheet formulas). Its independent Excel reference and save-copy/reopen recheck also passed: `obj/funding-final-debug-excel.log`, 104,014 matching formula cells in each comparison; all 337 VBA module hashes unchanged.
- Repeat via `Tools/Test-FundingStructure.ps1`, `Tools/Test-FundingExcelRoundtrip.ps1`, `Tools/Test-Structural3D.ps1` and `Tools/Verify-VbaModuleHashes.py`, always using private copies.

Client acceptance still needs permitted Funding add/delete via DIT, saved Excel/VBA interactive use, Check Sheet and accountant review. The supplied damaged client workbook has not been repaired or declared safe.

## DIT Save / Save As implementation

- Both buttons use the exact File Instance embedded images, icon-only with tooltips and a separator.
- Commands dispatch through the owning model's existing SaveFile / SaveFileAs services, retaining existing recovery-save restrictions and dialogs.
- Native and in-header editors are committed before save through their established posting handlers. Native validation rejection stops the save. Structure-authoring preview cannot invoke these writes.
- No new workbook save format, calculation policy, financial validation or persistence engine is introduced.

Debug and Release builds passed. Extended Tools/Test-EditorNavigation.ps1 checks native active-cell save, reopened XLSB value, dirty state, separate-path save service, rejected native-grid validation, no duplicate buttons, and header commit/closure. Only private test copies are saved. Manual Save As dialog/cancel, multiple open models and physical client DPI remain acceptance checks.

## Button sizing investigation

The shared current renderer requests a 28-unit glyph inside a 42-unit circle, scaled together. A UI-only native fixture rendered existing DIT SVGs at 1.0, 1.25, 1.5 and 2.0 presentation scales on a 96-DPI process. The normal 1.0 render keeps the document/clipboard glyphs inside their circles. Rendering without the shared sizing produces a much tighter fit.

This is evidence about the current executable, not proof of the client's physical DPI path. Client Summit version and Windows scaling are requested before changing icons or declaring that issue repaired. No icon artwork or shared sizing has been changed in this trial.
