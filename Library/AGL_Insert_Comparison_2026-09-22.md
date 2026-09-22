# AGL Funding / Development insertion comparison - 22 September 2026

## Decision

The saved results are structurally equivalent to their Excel counterparts across the checked formula/name/input surfaces, apart from one reproducible native XLSB export defect. They are **not an unconditional financial or release sign-off**.

- Development: the checked saved financial outputs agree with Excel and the original baseline. The insertion takes 148.187 seconds in the supplied Summit trace, versus the user's approximate 15-second Excel timing.
- Funding: insertion dimensions, names, input constants and formulas agree, except for the separate dashboard defect. Saved financial results differ materially from Excel. Both Check Sheets nevertheless retain the same OK indicators. Calculation consistency needs investigation before acceptance.
- Both Summit files: Multivariable Dashboard!B41 has lost its original formula and contains `=#VALUE!`. A no-edit DevExpress load/save/reopen reproduces this on a private copy; Excel independently confirms that the private saved file contains the error formula. This does not require an insertion or Summit UI code.
- Originals are unchanged. No production repair, test-release increment, commit or push was made by this comparison.

## Inputs and identity

All supplied files are under `C:/Sandbox/Insert Comp/`. Exact names and SHA-256 hashes:

| Role | Filename | SHA-256 |
| --- | --- | --- |
| Baseline | AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb | 30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47 |
| Funding / Summit | AGL - BP 2627 updated Mar26 v25 v2 v26_0005 10 funding records in summit.xlsb | 25FA5AB682A27685B46975E1EFE695248AA5AF59CBEA90BA9C998615488AFFE3 |
| Funding / Excel | AGL - BP 2627 updated Mar26 v25 v2 v26_0005 10 funding cols added in excel.xlsb | 5B1F2B80B97D588F87F1EC321792CC4BF42C17A4C910DD4E0D27E0A0BDA8348B |
| Development / Summit | AGL - BP 2627 updated Mar26 v25 v2 v26_0005 10 development lines added in summit.xlsb | 6D1F51A5F76E18DEAD8478C5EA620D132E68519D9AA390A5A540EE64B123FBF1 |
| Development / Excel | AGL - BP 2627 updated Mar26 v25 v2 v26_0005 10 development cols added in excel.xlsb | 5015D8BCE151FE5B44CF6A781B6ED36CB31151B8404B75873D45767A82F73E54 |

All five report internal ModelVersion 26.0005, 283 worksheets and 1,680 global names. The two Development files are independent Development-only tests, not cumulative Funding-plus-Development results. Both preserve the original Funding widths.

This establishes comparability of the supplied test set, not absence of bespoke modifications. A separately verified generic 26.0005 master is still required to establish that; the repository masters are 26.0001.

## Structural comparison

| Named range / measure | Baseline | Both Funding results | Both Development results |
| --- | --- | --- | --- |
| FacilityNames | E45:AG45, 29 columns | E45:AQ45, 39 columns | E45:AG45, 29 columns |
| LoanDescsOrd | E47:AC47, 25 columns | E47:AM47, 35 columns | E47:AC47, 25 columns |
| HouseTypeInID | G10:AU10, 41 columns | unchanged | G10:BE10, 51 columns |
| HouseTypeInMY | AV10:BB10, 7 columns | unchanged | BF10:BL10, 7 columns |
| LastIDColNum | AU9 | unchanged | BE9 |
| LastMYColNum | BB9 | unchanged | BL9 |
| Transactional_Records | A6:BV2496, 2,491 rows | A6:BV2606, 2,601 rows | A6:BV2636, 2,631 rows |

Facility/loan ranges are on Funding Assumptions; development ranges are on Development BP Assumptions. The final row/column templates are included where the workbook names include them; these dimensions are not counts of populated client records.

Pairwise native checks:

| Measure | Funding pair | Development pair |
| --- | --- | --- |
| Worksheets compared natively | 282 | 282 |
| Formula cells compared natively | 992,092 | 1,010,802 |
| Substantive formula differences | 1: dashboard B41 | 1: dashboard B41 |
| Identifier-case-only formula differences | 1,140 | 1,140 |
| Non-formula cell-value differences | 0 | 0 |
| Array/non-array classification differences, native scope | 0 | 0 |
| Non-sensitive name-definition differences | 0 | 0 |

There are 1,762 non-sensitive global/local name identities in the comparison. Credential-like names are deliberately excluded. The 1,140 formula-text differences are identifier casing in Development Expenditure (RESPCOST versus RespCost), not different references or string literals.

Transactional DB was compared independently through bulk read-only Excel Formula arrays because native per-cell introspection of that worksheet proved exceptionally slow. **All formulas and constants match exactly** between each corresponding pair: used rectangles are 2,917 x 78 for Funding and 2,947 x 78 for Development. These rectangles include surrounding cells beyond Transactional_Records. The separate bulk check does not assert complete array-group geometry or per-cell format parity.

All worksheets remain present in the same set. Original and Summit files have the same worksheet protection flags. Excel's saved files additionally leave Stress Sensitivity List unprotected; this is not a reason to change Summit's entry-state restoration.

## Saved financial results

Excel was opened separately with macros/events disabled, external-link updating off and manual calculation selected before opening each source read-only. No source was saved or deliberately recalculated. Native reads corroborate the findings. Numeric comparisons use an absolute tolerance of 0.0000001; blank and zero are distinguished.

| Worksheet | Funding Summit versus Excel: different saved cells | Development Summit versus Excel |
| --- | ---: | ---: |
| Summary Comp Inc - Trad View | 230 | 0 |
| Detailed Comp Inc - Trad View | 246 | 0 |
| Financial Position - Trad View | 323 | 0 |
| Financial Position - Alt View | 323 | 0 |
| Cashflow detailed | 298 | 0 |
| Check Sheet | 0 | 0 |

Native checks additionally find the same Funding difference counts in the alternative summary/detailed comprehensive-income views (230/246). Across the seven differing output sheets this is 1,896 cell differences, including repeated presentations of the same underlying financial changes; it is not 1,896 independent accounting defects.

All five retain 43 cached OK indicators on Check Sheet. This demonstrates that these saved checks do not detect the discrepancies; it is not evidence of a newly completed successful calculation.

Specific Funding evidence:

- Detailed Comp Inc - Trad View!J73, Interest Receivable And Other Income: Summit 60.14572207792846 versus Excel 1272.621346666452. Both have the same INDEX formula. The Summit value equals the baseline's saved value.
- Financial Position - Trad View!K26, Cash And Bank: Summit 3332.000000000008 versus Excel 360639.14630585833. Again, matching formulas but different saved results.
- Cashflow detailed!BO9: both expanded results contain `=SUM('Loan Interest Paid:Loan Repayments'!E8:AM8)+'InterCo Decreases'!I9`. Summit saves 0; Excel saves -49396.33200000001.
- Cashflow detailed!BR9: both contain `=SUM('Loan Drawdowns:Loan Repayments'!AN8:AQ8)`. Summit saves 0; Excel saves -15166.96532686803.
- Summit Funding's saved SOCI/financial-position results equal the baseline, whereas 44 detailed-cashflow saved cells differ from that baseline. The Excel Funding result changes SOCI, financial position and cashflow. Both Development results retain the baseline's checked financial values.

Values above are in the workbook's stored/display units, not independently converted into pounds.

The evidence supports investigating calculation completion, dependent caches and save ordering. It does **not** establish that every Excel Funding figure is the correct financial answer, or that a simple refresh/full rebuild alone is the safe repair. Compare controlled, fully calculated copies with the relevant VBA behaviour before choosing the repair. No financial model logic was replaced.

## Separate native XLSB formula-loss reproduction

The untouched baseline's Multivariable Dashboard!B41 contains a long IF/MAX/CONCATENATE formula. A plain native Workbook loads that formula intact. Saving that in-memory workbook as XLSB into a unique private diagnostic directory, without any insertion or application model, then reopening it produces `=#VALUE!`. Read-only Excel independently reads the same error formula in the private copy.

Both user-supplied Summit results contain this error formula; both Excel results retain the original. This is actual formula loss, not merely an error-valued calculation cache. Reproduction currently localises it to native XLSB save/reopen, not to Funding or Development insertion. The exact exporter/parser limitation and a workbook-compatible correction remain to be established. Do not patch the original workbook's formula merely to suppress the symptom.

Private evidence: `obj/AglNoEditRoundtrip/d8508f7000924257a7ff90655a876720/native-no-edit.xlsb`; log `obj/agl-no-edit-roundtrip.log`. This diagnostic file is not a client deliverable.

## VBA preservation

All 335 extracted VBA module identities and source hashes in each Summit result are exactly identical to the baseline. Both Excel results differ only in the VB_Base attribute in six UserForm modules: ColumnsCopyForm, DSARejectsForm, ProtectForm, TermsAndConditions, UnprotectForm and ValBPSwitchForm. Ignoring only those metadata attributes makes all extracted source identical. No executable procedure changes were found by this comparison.

Only identities, hashes and classifications were reported. No passwords or raw extracted VBA source were persisted in the repository. This is source-preservation evidence, not proof that every macro or UserForm behaves identically; VBA was not executed.

## Performance

Both supplied Summit traces identify version 2.57, 32-bit execution. The user reports approximately **15 seconds** for Excel in response to the combined Funding/Development timing question. Use this as indicative, not two separately timed controlled measurements.

| Development stage | Recorded time |
| --- | ---: |
| 3-D-reference capture, two batches | 37.307 s |
| Seven linked-sheet column shifts, two batches | 55.358 s |
| Template copying | 1.343 s |
| 3-D copy handling | 0.369 s |
| Unprotect / protect | 0.969 s / 0.002 s |
| Post-actions | 52.054 s |
| Structural operation | 147.698 s |
| Complete DIT Add Lines | **148.187 s** |

Post-actions include fourteen Transactional DB mirrors: 46.384 seconds resizing and 5.280 seconds restoring the calculation engine. Capture, column shifts and post-actions together occupy 144.719 seconds, approximately 98% of the DIT total. The Development log contains no ContextSwitchDeadlock warning. Approximately 148 / 15 is 9.9x, but this is not a controlled cross-application benchmark.

Funding reports 361.472 seconds through DIT completion, including a debugger-paused Transactional DB stage. Do not quote that as clean execution time or use it for a speed ratio. Its unaffected stages still show substantial work: 18.337 seconds 3-D capture, 66.325 seconds sheet shifts, and 23.267 seconds across the other ten mirrors. The pause-affected mirror consumes 236.162 seconds. Its uncontaminated duration cannot be reconstructed by subtracting an assumed pause.

Source logs:
- Funding: `C:/Users/jmwor/.codex/attachments/11429ceb-c092-4ae4-a263-cbdbb1b7db97/Pasted text.txt`, operation 5e7c97b5.
- Development: `C:/Users/jmwor/.codex/attachments/f91f3835-331b-4b55-9542-cdf9e9922e94/Pasted text.txt`, operation 7a84ee00.

Do not conflate these user-operation timings with the separate analysis-harness runtime.

## Recommended next work, not implemented

**Follow-up:** the save/export and saved-calculation items below are addressed by the 2.58 trial in `Library/XLSB_Save_Compatibility_Trial_2026-09-22.md`. This original comparison section describes the earlier read-only phase; insertion performance optimisation remains pending.

1. Reconcile Funding's saved calculation state on disposable copies, using explicit before/after dependency results and an agreed Excel/VBA calculation reference. Protect the correct insertion boundaries and formula references already demonstrated here.
2. Diagnose and repair the no-edit XLSB export loss of dashboard B41; add a regression checking formula preservation through native save and independent Excel reopen.
3. Then optimise measured costs: repeated workbook-wide 3-D capture, the linked column-shift sequence and mirror shifts. Any indexed/cached reference inventory must be invalidated correctly after structural changes. Do not simply disable 3-D handling, defer required financial calculation, parallelise access to the shared workbook, or remove synchronisation to achieve a faster timing.
4. Re-run an unpaused Funding benchmark; compare x86/x64 Release on the same populated workbook only after correctness gates pass.

## Reproduction and boundaries

Tools added for this read-only engineering audit:
- `Tools/Compare-AglStructuralResults.ps1` and `Tools/StructuralResultComparison.cs`: native identities, ranges, names, formula/input/array classification comparison outside Transactional DB.
- `Tools/Inspect-AglExcelSavedResults.ps1`: independent Excel saved-output and bulk Transactional DB formula/constant comparison.
- `Tools/Inspect-AglLoad.ps1`: immediate-load probes; optional -SaveProbe writes only a uniquely named private no-edit round-trip copy.
- Existing `Tools/Verify-VbaModuleHashes.py`: exact module hashing. Additional in-memory attribute comparison used external temporary oletools dependencies, not repository extraction.

Final native evidence: `obj/AglStructuralComparison/e0bd1d15d6ce4e919c53ef16b3cd0a4f/`, log `obj/agl-structural-comparison-final.log`.
Final Excel evidence: `obj/AglExcelSavedResults/62fd6ad17b45492594c1cb3709868f7d/excel-saved-results.json`, log `obj/agl-excel-saved-results-final.log`.

Initial native attempts were deliberately stopped at the expensive per-cell Transactional DB inspection, and an initial Excel report writer was corrected. Their partial logs are not the completed evidence. Final passes report unchanged hashes for all five originals. No running user-started Summit or Excel session was stopped.

Full formatting, all array-group geometry, charts/controls, macro execution, complete recalculation and financial/accountant acceptance are outside this comparison. No production code changed, so no new Debug/Release application build or test version was issued.
