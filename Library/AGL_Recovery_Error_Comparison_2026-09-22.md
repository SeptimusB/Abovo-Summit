# AGL recovery error comparison — 22 September 2026

## Outcome

The supplied recovery file reproduces the recent integrity report. The findings are a mixture of Summit/Excel calculation differences, intentional chart markers, workbook lookup defects/edge cases, and the user's deliberately introduced balance discrepancy. They are not evidence of 1,059 separate corruptions.

The original recovery file was not modified. No application repair was implemented.

| Original finding | Count | Controlled Excel result |
| --- | ---: | --- |
| Menu / Contents / sheet-list VALUE errors | 563 | All clear to valid worksheet names |
| Credit Rating VALUE errors | 18 | All reproduce |
| Non-literal N/A errors | 474 | All reproduce when only reported error cells are recalculated against saved inputs |
| Literal NA() markers, separately counted | 98 | All reproduce; deliberately requested by formulas |
| Check Sheet warnings | 3 | Present in the saved file; retained during targeted test; user confirmed deliberate bad input |
| Dashboard XLSB-normalization candidate | 1 | B41 is not a calculation error in this test; its selected IF branch returns empty text |

Saved error cells total 1,153 = 581 VALUE + 474 non-literal N/A + 98 literal NA().
The integrity total of 1,059 excludes the 98 literal markers and adds the three Check Sheet findings and one XLSB normalization candidate.

## Evidence and method

Source:
`C:/Sandbox/Insert Comp/AGL - BP 2627 updated Mar26 v25 v2 v26_0005_recovery.xlsm`

Frozen, byte-identical evidence copy:
`obj/RecoveryErrorAudit/cdc8649d1e3b4c70a11e1c71d4a61ad1/recovery-evidence.xlsm`

Both original and evidence-copy SHA-256 before and after:
`405BC929A6A99A77B986EB05A0279E94196B99F025388D88B099A711F5D742B5`

Repository: main, checkpoint 49a7d34. Release executable dated 22 September 2026 19:45:42; SHA-256:
`87490859CCADB81215A9FF7150627C3BC64BF4A0DA9B2A5B1C0AEE060D4562FD`

Excel 16.0 build 20326 was opened in isolated automation instances. Evidence was opened read-only with macros/events disabled, external-link updates suppressed, and manual calculation. Every successfully opened diagnostic workbook was closed without saving; diagnostic-owned Excel instances were closed. No existing user Excel instance was closed. Package inspection found VBA present and no Excel 4 macro-sheet parts.

Tests:
1. Scan the saved caches through DevExpress and Excel: identical error counts and error values.
2. Full rebuild in each engine, retained as diagnostic evidence but **not a valid full financial parity test** for the reasons below.
3. Fresh Excel open, then calculate only the originally reported error cells, leaving VBA-dependent inputs at their saved values. All 563 menu-related errors resolve; the other 590 error cells remain (474 + 98 + 18), with no NAME errors introduced.
4. Minimal, unsaved DevExpress workbook probes isolate CELL address formatting and custom-function registration.
5. Inspect selected XLSM package cells directly to corroborate saved inputs and formulas.

This follows the spreadsheet read-only investigation workflow: trace inputs, compare results, preserve originals, and report rather than repair.

Local detailed evidence and diagnostic scripts remain under the ignored obj directory:
- native-cached.json, native-rebuilt.json, native-probes.json
- excel-cached.json, excel-rebuilt.json, excel-calculation.json
- excel-targeted-before.json, excel-targeted-after.json, excel-targeted-details.json
- Audit.cs, Run.ps1, ExcelAudit.ps1

## 1. Menu formulas: confirmed calculation-engine difference

Breakdown of the 563 VALUE errors:
- Hidden - BP Menu Sheet: 234
- Contents: 232
- Hidden - Sheet Lists: 97

For example, Hidden - BP Menu Sheet!F7 extracts the worksheet name from CELL("address",Welcome!B1), assuming a closing square bracket before the sheet name.

Observed Excel result:
`'[recovery-evidence.xlsm]Welcome'!$B$1`

Isolated DevExpress result for an equivalent cross-sheet reference:
`Other!$B$1`

DevExpress omits the bracketed workbook component. SEARCH("]", ...) consequently returns VALUE, which propagates into Contents and the sheet lists. Excel recalculation resolves every member of this group.

This is not a lost worksheet or broken structural range. It is a genuine compatibility issue in workbook metadata/navigation formulas. Do not globally suppress VALUE errors: the Credit Rating findings below are different.

Recommended follow-up: support both qualified and unqualified address formats, or provide a verified native-compatible equivalent, with exact Excel/VBA round-trip tests. No formula changes were made in this investigation.

## 2. Funding / repayment year: three related N/A results

- Funding Assumptions!I4 is the text `40+`.
- Funding Assumptions!I5 uses an exact MATCH of I4 against the numeric year-number column Hidden - Year Table!G6:G106.
- There is no numeric-year match for the text `40+`.
- Repairs & Maint. Rephasing!L5 links directly to Funding Assumptions!I5.
- Error Checks!B9 also contains `40+`; C9 performs the same year-table lookup and fails.

All three errors reproduce in Excel. The year table extends well beyond year 40; its length is not the issue.

Recommended follow-up: define the intended display for repayment beyond the model horizon, retaining the meaning of `40+`. Do not invent a repayment date or reinterpret it as year 41.

## 3. Credit Rating: out-of-scale inputs

Credit Rating!N50:P50 contains approximately:
- 8.8678334898
- 8.5537222765
- 8.7517798840

The descending lookup thresholds in B137:I137 are:
`8, 6.5, 5, 4, 3, 2, 1, 0`

These three values exceed the highest threshold. The descending approximate MATCH(...,-1) used in N213:P216 has no match. That produces twelve N/A cells.

TREND in N217:P217 then returns VALUE, which propagates into the weighted rating and displayed rating/chart outputs. The saved chain contains eighteen VALUE errors.

The twelve N/A and eighteen VALUE cells reproduce in the targeted Excel calculation. A macro-disabled full rebuild additionally contaminates Credit Rating through unavailable VBA functions, so those extra findings must not be mistaken for thirty additional independent root causes.

Recommended follow-up: ask Abovo which rating should apply above the top threshold (or whether an explicit out-of-range warning is intended). Do not silently clamp, extrapolate, replace with zero, or change scoring policy.

This investigation does not establish whether the user's deliberate balance edit caused these ratios to exceed the threshold; no historical baseline attribution is claimed.

## 4. Hidden import-template lookup

Hidden - Tenure Totals Start!B6:
`INDEX(Appraisal1,MATCH(A1,ImportedSchemes,0))`

A1 is `Imports Start`, while ImportedSchemes refers to the empty placeholder List Imported!A8. Appraisal1 refers to List Imported!F8. The lookup has no matching imported scheme.

This N/A reproduces in Excel. It is a template/boundary lookup with no imported record, not evidence by itself of damaged client data. Any future expected-result rule should remain conditional on the absence of an applicable import.

## 5. Chart/covenant marker N/A results

Of the 474 non-literal N/A cells:
- 87 are on OW - Charts Source Data.
- 131 are on OW - Covenant Calculation.
- 240 are on OW - Live Covenant Calculation.
- The remaining sixteen are the twelve Credit Rating lookups, three repayment-year cells, and one import-template cell above.

Within the 458 chart/covenant cells:
- 337 explicitly contain an NA() branch.
- 120 are INDEX-based consumers of the chart/covenant marker data.
- One, OW - Covenant Calculation!AB61, searches for a first marker value of 1 and returns N/A when none exists.

These all reproduce in the controlled Excel test. They are chart/covenant presentation results, not 458 demonstrated workbook defects. The supplied workbook includes the user's deliberate test input, so this does not certify that every covenant is financially satisfactory.

The integrity scanner currently excludes only literal `=NA()`. This explains much of the alarming count. Future classification should use verified formula/range-specific rules, distinguish inherited markers, and retain unexpected errors. **Do not blanket-ignore N/A on an entire worksheet or throughout the workbook.**

## 6. Intentional Check Sheet warnings and dashboard normalization

Saved Check Sheet E23, E33 and E39 read `Check`, and D63 reads 3. These correspond to the balance issues the user deliberately triggered by editing Funding Assumptions!G82 (saved value 9,565). Targeted error-cell recalculation leaves these checks unchanged; it does not independently recalculate the entire financial model.

Multivariable Dashboard!B41 returns empty text in both saved and Excel-rebuilt observations. The existing integrity finding is an XLSB serializer-normalization candidate, not a demonstrated broken dashboard calculation. Its long CONCATENATE branch was not exercised by this workbook state.

## 7. Important separate calculation risk: PMCOST

Follow-up: the user-authorized signature repair is implemented in test release 2.70. The frozen recovery evidence then recalculates with all 570 PMCost-containing cells numeric and no NAME errors. See [repair and profiling evidence](PMCost_Repair_and_Profile_2026-09-22.md). The observations below describe the preceding 2.69 build; full Excel/VBA financial acceptance is still pending.

The current Release native PMCOST registration still advertises **seven** parameters. The model's development formulas call it with **eight**.

An isolated DevExpress test using the current registered implementation rejects a valid eight-position PMCOST formula with `ArgumentException: Formula is incorrect`. A seven-position RESPCOST control formula succeeds.

This reconfirms the previously documented mismatch in Library/Contract_XLSB_Audit_2026-08-24.md, F1. A standalone native reload/full rebuild of this recovery XLSM introduced 52,161 NAME errors, as well as propagated errors. This is a significant calculation/recovery-path risk, not one of the 1,059 original report findings and not a completed reproduction through the running Summit UI.

A macro-disabled Excel full rebuild introduced 54,625 NAME errors, including the VBA custom-function dependency chain. That is a different test limitation: VBA was intentionally not allowed to execute. It prevents claiming full financial equivalence from this run.

Recommended priority: repair and parity-test the eight-position PMCOST interface under separate implementation authority, then repeat a full calculation comparison with the trusted Excel/VBA workflow. Do not hide this issue as an expected integrity finding.

## Proposed next work, not performed

1. Repair and parity-test PMCOST registration and the recovered-XLSM calculation path.
2. Repair menu address compatibility without altering VBA expectations.
3. Classify verified chart markers separately from actionable integrity findings.
4. Resolve repayment-horizon display and Credit Rating out-of-range policy with Abovo.
5. Repeat accountant-approved full Excel/VBA versus Summit output comparisons on disposable files.

No source workbook, source master, structure XML, application code, or test version was changed. This is a diagnostic report, not a release or financial certification.

## Documentation consulted

- [DevExpress information functions](https://docs.devexpress.com/OfficeFileAPI/15497/spreadsheet-document-api/formulas/functions/information-functions): CELL is supported, but the exact address-format difference above was established by local probes, not inferred from general documentation.
- [Microsoft AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity): macro-disable setting for programmatic opening.
- [Microsoft CalculateFullRebuild](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefullrebuild): scope of a full dependency rebuild.
