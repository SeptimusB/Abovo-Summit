# Remaining client input and presentation review

Date: 23 September 2026. Inspected application: **2.80 Release, x86**. This is a diagnostic checkpoint, **not a new application release or a claim that the defects are fixed**. The preceding uncommitted structural repairs remain intact for client testing.

Follow-up: the approved first repair batch is implemented as [2.81 typed-input/source-mapping repairs](Client_Input_Repairs_2026-09-23.md), with new regression evidence and the additional AGL Other Fees mapping limitation. The observations below are the retained pre-repair baseline; remaining conditional-format/layout questions are not closed by that implementation.

## Scope and source

The request is to review input precision, negative percentages, locked cells, paste, dates, conditional formatting and layouts, asking about ambiguity rather than inventing behaviour.

Source: D:/Downloads/Summit v2.51 comments.docx, unchanged SHA-256 955D44B6C714ED9BA6D5F9B4560C983764B38B83161A71B88C2BF400211BD9ED. The [90-item register](Client_Report_2026-09-22.md) and [earlier checkpoints](Client_Review_Checkpoints_2026-09-22.md) preserve the full wording and colour annotations. P-numbers are document paragraph locators, not page numbers. Orange means the user thinks fixed; green means agreed, neither means independently verified. The document predates this build.

Native tests used private copies of the current Demo master, Blank master and nominated AGL workbook. They exercised the actual DIT editors, data-point mappings and grouped ChangeManager paste path. Test edits were undone. Original files were never saved or modified. No production code/XML changes, version increment, commit or push were made in this review. Tools and audit documentation were added/updated.

## Confirmed findings and proposed repair boundaries

| Area / client reference | Current evidence | Proposed boundary |
| --- | --- | --- |
| Survey Input decimal entry/paste, P054 | Stock Condition Inputs!D9 is an unlocked decimal-validated Excel input. The DIT defines it as I. Native input 1234.56 is stored as 1235; decimal paste parsing rejects 1234.56. The inventory shows the same integer type across the survey cost columns. | Correct monetary field types, not the BP Year/count fields. Test all survey archetypes and mixed Excel pastes. Preserve the workbook's current display formats unless separately authorised. |
| Development Expenditure precision, P077 | Dvpt BP Rev and Exp Assumptions!D8, D16, D25 and D34 store 1235 after entry of 1234.56. They are monetary costs using I, while Excel validation permits decimals. | Repair the monetary ranges; do not indiscriminately change years, units or percentages. Include Management Costs, Other Costs, Responsive Repairs and Unit Rebuild Cost. |
| Housing Asset precision, P120 | Housing Asset Assumptions!D14:D22 and monetary grant/depreciation examples D59, D61 and D69:D77 use I and round 1234.56 to 1235. G29 uses M and preserves 1234.56. | Fix the identifiable monetary fields. Number of Properties and useful-life/amortisation periods also use I but are **not automatically defects**; retain those pending their business rules. |
| Weekly monetary display, P034/P069 | Service Charge Assumptions!E19 and Development BP Assumptions!O169/O170 preserve 1234.56 but display £1,235 in the DIT. Excel's native formats display two decimals. | Display-format repair separate from storage-type repair. Use workbook-owned formatting; do not round the stored values. |
| Negative percentages, P084 | Twenty representative Economic Assumptions editor tests, including CPI/RPI views and real-increase sections, have a native spin-editor range 0–1. Entering -0.025 stores 0. Excel decimal validation in inspected CPI/real-rent cells permits negative values. The paste parser separately accepts -2.5% as -0.025; full percentage-paste write was not exercised. | Respect workbook/XML constraints consistently in editor and paste. Do not apply a blanket 0–100% limit. Include legitimate values above 100% where allowed; Covenant EBITDA MRI currently contains 110%. |
| Covenant input locks and year columns, P108/P109 | The whole Covenant data source has RO=TRUE, blocking the input columns even though D8:H8 are workbook-unlocked. BP Year/Year map to E8/F8, not actual A8/B8. The MergeAcross offset applies C-3/C-2 relative to the end of the five-column Rep_Cov_01 range. Reproduced in Demo and AGL mappings; Excel independently confirms A8=1, B8=2026/27 in Demo. | Repair the year sources and table-level RO together; keep A/B read-only and respect each real input's protection/validation. Simply unlocking the table is incomplete. Inspect Debt monetary typing too. |
| Funding Other Fees, P105 | Demo's populated D230 is eligible for editing. Blank D231:D233 and one-off inputs are conditionally blocked when the corresponding B-column description is blank. The XML renders the descriptions as repeating headings but does not provide the description input editor used to populate them. B231 and B238 are unlocked in Excel. | Expose the missing description inputs and refresh dependent eligibility. Do not remove all conditional locks merely to make the amount cells editable. |
| Funding one-off fee year versus date, additional finding | Funding Assumptions!C238:C240 are defined as D with heading Date in Summit. Excel's C237 heading is Year, and C238 has General format and whole-number validation 1–40: a BP year, not a calendar date. | Use an integer BP-year editor/label for these cells. This is distinct from the client's ambiguous Annuity/date-ghost request. |
| Blank Stock Survey date, P053 | Blank master Repairs & Maint. Assumptions!C50 is genuinely empty, but the native date control displays 30/12/1899. Demo displays 01/04/2024 correctly; AGL's mapped date is 01/04/2026. AbovoDEDateEdit converts NumericValue to an OA date without first testing IsEmpty. | Show empty as empty. Preserve real dates and worksheet validation. This reproduces the bogus default, not the separate claim that a populated date cannot be edited. |
| Management conditional formatting, P046 | Clearing Description D17 through guarded paste makes the actual D19 workbook rule =ISBLANK(D$17) evaluate True with DarkGray pattern. DIT's cached category colour remains blue; its column has HasRules=False and the custom draw path does not render that conditional pattern. Undo restores D17. | Implement workbook-rule-based effective formatting and dependency refresh, retaining priority/relative reference/pattern semantics. Do not invent colours or use a generic static repaint. Keep source-cell protection independent from visual treatment. |
| Housing Asset layout, P122/P123 | Housing Asset Grant Assumptions and Remaining Useful Life contain the same range sequence (DepnType, Rep_HAA_00, Rep_HAA_05/05a/05b, Rep_HAA_06, Rep_HAA_07/07a, WriteoffOpBalIn); Housing Asset Depreciation repeats part of it again. | Agree the grouping, then remove the overlap deliberately. Swapping two tab labels would not fix it. |

### Related formatting and lock work that remains to reproduce

- P045 default Management summary-category formatting; P052 Repairs category formatting, truncated descriptions and delayed three-table refresh; P063/P064 Development House Type/Include rules; P083 Real Rents formatting: inspect each exact workbook rule and before/after state. The Management result establishes a concrete problem, not proof that all these cases have one identical cause.
- P047 Management Costs Year 1: the sampled Demo D25/E25/F25 and D35/E35/F35, and AGL D25/E25/F25 and D38/E38/F38, are workbook-unlocked and pass the current DIT edit/paste eligibility gate. No new blanket unlock is justified. A client reproduction should identify workbook, row/category and preceding action. This check did not perform physical keyboard entry in every one of those cells.
- P078 Planned Maintenance From Yr: the XML requests a repeating header editor, but lacks the complete explicit editor/type configuration used in other working repeating inputs. Body percentage cells are distinct from their year-header editors. Confirm the actual header control and its source before a targeted repair.
- P104 first Funding date: several repeating date inputs still have zero protected initial lines. Confirm the exact Variable/Cash Rate table/date before changing a shared boundary; do not lock all first date cells.
- P097 Facility Name dropdown: the exact offending list/value pair remains a client clarification/reproduction item. Do not swap repository lists solely from the report's short description.
- P087 Real Development R&M: current XML uses IR_DevRepairsReal, Rep_Econ_23 and Rep_Econ_23PI, distinct from Service Charge's sources. The current mapping is not evidence of the reported service-charge substitution. Confirm the client's workbook/build and displayed table if still wrong.

## Retests and limits

| Item | Result in this review |
| --- | --- |
| P017 single dropdown value pasted to multiple cells | **Native service pass:** two Management Cost Category cells receive one valid value, then one grouped Undo restores both originals. The OS clipboard was not touched. |
| Invalid multi-cell dropdown paste | **Native service pass:** one valid value plus one invalid category rejects the entire batch; both destination values stay unchanged. Only the diagnostic runner's expected rejection dialog was dismissed. |
| P127 Other Current Assets amount | **Native editor pass:** B7 stores 1234.56 exactly; Undo restores original. The existing 2.68 M-type repair remains effective. |
| P130 Journal amount | **Native editor pass:** E7 stores 1234.56 exactly; Undo restores original. Journal Year is not changed. |
| Original file preservation | Demo, Blank, AGL and DOCX SHA-256 values unchanged. |
| Numeric-test safety | All completed actual editor probes restored their prior cell value through ChangeManager Undo and retained the source number format and worksheet protection state. Integer probes on legitimate counts/periods are observations, not an instruction to make them decimal. |
| Copy with headings / P018 | Source review confirms the earlier flattened non-period headings and explicit-selection rectangle. Only Summit-tagged clipboard data has its own header metadata stripped on re-entry. Arbitrary Excel headings are not safely identifiable. No new physical Excel copy/paste acceptance test performed here. |
| P048 width drift, P062 caption, P067 SHG mapping, P101 month-end input | Prior engineering checks remain in the earlier checkpoint. They are not automatically client-approved or newly exhaustively tested by this review. |

No screenshot-based certification of current client DPI, restored-window layouts, nested scrolling or physical keyboard/mouse selection is claimed. No workbook formula-generation, VBA or financial acceptance test is implied. Macro-disabled Excel inspections checked validation/protection/labels and cached values, not recalculated business results.

## Clarification queue

Three questions were sent during the review and remain unanswered at this checkpoint:

1. P099/P100: does Funding date “ghost” mean an empty date-format hint, suppression of blank/zero dates, or a calculated default shown in grey?
2. P122/P123: should grant inputs be grouped under Housing Asset Grant Assumptions, and depreciation/useful-life inputs together under Housing Asset Depreciation? Current definitions overlap.
3. P085: which Economic Assumptions field and condition should change appearance? A concrete before/after example or screenshot is needed.

Additional reproductions to collect in the relevant repair batch: precise P047 locked Year 1 case, P097 wrong Facility Name list, P104 first date location, and a normal/restored client-screen screenshot for each layout complaint.

P049's “locked” Description means **freeze/pin it during horizontal scrolling**, not prohibit edits. Layout tasks P025/P026/P032/P033/P037/P038/P052/P054/P057/P083/P086/P090/P092/P093/P094/P096/P098/P103/P116/P117/P131 must retain their individual wording and source location. Do not close them based on this data-path audit. For dimensions, test 100%, 150%, 200%, restored and maximised windows, and one real client display.

The separate BP Start Date decision remains in [Client_Questions.md](Client_Questions.md): month-start versus any calendar date must be resolved from the model contract, not a UI guess.

## Recommended implementation checkpoints

1. **Typed inputs and source mappings:** monetary I definitions, negative/unbounded-as-per-workbook percentages, 2dp displays, Covenant year/RO corrections, empty date and one-off fee BP-year typing. One bounded release; direct input, paste, Undo/Redo, dirty state and independent save/reopen tests. Audit existing formulas/locks before permitting any newly reachable input.
2. **Conditional formatting and dependent editors:** use the workbook's rules for Management/Repairs/Development/Economic cases; expose Funding fee descriptions; verify rule changes refresh without tab switching. Retain edit protection and test keyboard/paste together. Avoid per-paint full-workbook recalculation.
3. **Agreed layout/wording:** Housing tab grouping and date hints only after clarification; then pinned descriptions, widths, nested scroll and visible edit focus. Client-screen acceptance is required.

Every application repair batch requires a version increment, Debug and Release builds, representative native regression and client acceptance. This review itself does not bump 2.80 or commit the earlier structural work.

## Evidence and repeatability

Added diagnostic Tools/ClientInputReviewFixture.cs and Tools/Inspect-ClientInputExcel.ps1. Tools/Test-ClientReport.ps1 accepts optional Fixture and ReviewCases; its previous defaults are preserved.

Completed native runs under obj/ClientReportTests/:

- ea7a57dcfb554a3ca1cdf04107682752: full Demo inventory, 2,410 representative field/cell records, 2,122 distinct source cells, and 103 actual editor probes. These include repeated views and intentionally integer controls, not 103 independent bugs.
- 9a3f504561d6453ab5a4606dca00c5f3: Management grouped paste/rejection, Demo survey date, Other Current Assets and Journal precision retests.
- f18d9d558e194a3a97797cd74fe7fb79: AGL Management, Repairs, Funding and Covenant inventory.
- cc8bd62c77164175a5812bd6fb4a38d4: guarded description clear, evaluated Management conditional rule, grouped paste/rejection and Undo.
- 1f3a9bd6bd6c42069c9a486e7daefadd: Blank survey-date reproduction.

CSV inventory column LoadedWhileTab identifies when pre-built data sets were first observed, **not necessarily their owning tab**. Source worksheet/address and CSID are the reliable locators; the editor-tests CSV records the actual selected tab. Older incomplete diagnostic runs (placeholder-tab iteration and an optional reflection argument mismatch) are excluded from completed-run claims.

Independent Excel read-only inspection: obj/ClientInputExcel/162d48e6ec0f4d1fae0d034ff6920f56/excel-observations.json (42 cells, including the explicit one-off fee Year heading), with macros, events, link updates and calculation disabled; closed without saving. Earlier complementary runs: f5c8837dc9c1425688ca26f25bbae18a and c89728392d0a4565bf2bd006e503e2dc. Excel's cached DisplayFormat observations are not used to claim live conditional-format parity.

The DevExpress [formatting documentation](https://docs.devexpress.com/OfficeFileAPI/14915/Spreadsheet-Document-API/Cell-Basics/Formatting-Cells) was checked as background. The specific findings above come from the local native control/workbook probes and source inspection, not a generic forum workaround.

Reproduce (Windows PowerShell / .NET Framework):

~~~powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File Tools/Test-ClientReport.ps1 -Fixture ClientInputReviewFixture.cs
powershell.exe -NoProfile -ExecutionPolicy Bypass -File Tools/Test-ClientReport.ps1 -Fixture ClientInputReviewFixture.cs -ReviewCases '14,15,41,42'
powershell.exe -NoProfile -ExecutionPolicy Bypass -File Tools/Inspect-ClientInputExcel.ps1
~~~

Run UI probes serially, on disposable copies, with no client interaction in the owned hidden test forms. The harness reports diagnostic observations even when precise=False; a zero process exit means the diagnostic completed and its restoration guards passed, **not that all application cases passed**. Existing embedded browser disposal logged Chrome_WidgetWin_0 error 1412 after clean test shutdown; it did not invalidate source hashes or editor checks, but is not treated as a newly repaired issue.

Source hashes:

- Demo: 1ED79726D8D1129C699242C9E8B3E988AC920EA7108D3C0B105BF9B0BC8A5E9C
- Blank: E05274ACD3D4821AC38F013AC45A7B57CE335BD524938D9F21E1B640D0CC5E90
- AGL: 30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47
