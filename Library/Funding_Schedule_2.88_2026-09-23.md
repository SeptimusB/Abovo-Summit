# Funding date schedules — test release 2.88

Status: Ready to test with Jon before client handoff. This is not Alex acceptance.

## User workflow and presentation

Funding Assumptions has a small, native skin-painted plus before each supported date section's header. Its tooltip is **Add schedule**. The existing footer **Add rows** button remains separate. All seven Facilities and Loans date sections, four Variable and Cash Rates sections and three Investments date sections are supported; non-date fee/headings are not offered a schedule.

The dialog uses Summit's shared font, native DevExpress editors/buttons and shared grid formatter, Abovo-blue headings, a scrollable settings area and a read-only, multi-select/copy preview. Settings and preview are side by side. No embedded web UI or third-party control is introduced. Native 96-DPI screenshots have been inspected; high-DPI/mixed-monitor client acceptance remains manual.

Choose a start date; monthly, quarterly or annual interval; same day, first/last calendar day or first/last working day; and either an occurrence count or inclusive end date. Same-day schedules retain the original day number after short months/leap years. Dates before the start or after the adjusted inclusive end are excluded. Maximum preview size is 1,000 dates. Changing any option invalidates the preview and disables Apply until refreshed.

**Delay for bank holidays** is optional, with England and Wales, Scotland or Northern Ireland selection. Working days exclude weekends; enabling holidays also moves non-working dates to the next working day, except **last working day**, which stays within its month. The official [GOV.UK feed](https://www.gov.uk/bank-holidays.json) is refreshed for each holiday-aware preview. Only published years are accepted, not extrapolated future holidays. All three live downloads passed on 23 September 2026, covering 2019–2028. A cached official response up to seven days old can be used offline and is explicitly labelled with its timestamp. Older/missing data blocks that preview unless adjustment is turned off. TLS 1.2 is scoped to this request with normal certificate verification. No claims of indefinite future holiday certainty are made.

Apply fills the first safe blank date rows in order, retaining duplicate dates and preserving existing dates/amounts/formulas. Rows with orphaned amounts or formula-owned inputs are skipped. It never fills loan amounts. If short of capacity it uses Summit's established row-insertion service for the shortage plus **five spare rows**. After rules refresh, the first editable amount beside the first added date receives focus; if no loan amount is editable (e.g. the Blank template before defining loans), focus returns to that date header without overriding locks.

## Workbook and failure contract

Date writes are one typed, calculated, journalled `ChangeManager.ProcessChanges` operation. Read-only/recovery-required models are rejected. Source templates and XLSB schema are not altered merely by opening/previewing. Existing protection state, fill-pattern editability and unavailable-cell colours are unchanged.

**Undo removes the schedule dates, not the additional blank rows.** The dialog explains this. Structural insertion uses the existing failure/recovery guard. If posting dates fails after a successful expansion, the extra blank rows remain and the error explicitly says so; the implementation does not claim a whole-workbook rollback spanning these two operations.

Thirteen pre-existing literal-row Funding datasource definitions would otherwise become stale after inserting rows above them. They now resolve from existing named-range anchors plus offsets/counts through opt-in XML fields. No workbook names or worksheets are added. Unannotated CR definitions retain the previous path. Original current-master coordinates match before insertion, and all thirteen anchors track the tested expansions. This does not sign off every older/bespoke Funding layout.

## Evidence

- Debug and Release builds pass; both normal executable locations contain 2.88.
- `Tools/FundingScheduleFixture.cs`: all fourteen sections pass duplicate-date writes, protection/formula preservation and grouped undo/redo. Expansion is tested for Repayments, Interest Receivable and Investment Increases: shortage + five, preserved existing constants, correct anchors, restored protection and retained spare capacity after date Undo.
- Native header click opens the associated preview; Apply invalidation and actual date posting are exercised. Demo tests focus the first editable amount; Blank tests the safe date-header fallback. Separate XLSB save/native reload preserves all fourteen date ranges.
- Final Release Demo evidence: `obj/ClientReportTests/351a9516cff8418ea88c3b3db4e0df4d` (158 assertions, including all three tabs and Cancel non-mutation). Debug Blank evidence: `obj/ClientReportTests/d82fad8c3b994636b779d3a2b3faba22` (155 assertions, including the safe first-date focus fallback). Earlier Release Demo evidence used for Excel round-trip: `obj/ClientReportTests/766e95407fd2499baadd657851b6489d` (155 assertions).
- `Tools/Test-FundingScheduleExcel.ps1`: read-only Excel open / SaveCopyAs / native reload preserves sheet order, all 1,679 names and 6,663 Funding cells' formulas, constants/dates, locks and fill patterns. Evidence: `obj/FundingScheduleExcel/165a6440bdba45a3a20ffe92c82dc153`. Macros/events/calculation disabled; financial/VBA execution acceptance is separate.
- `Tools/Verify-VbaModuleHashes.py`: original Demo, Summit-saved and Excel-saved copies have identical identities/source hashes for all 337 VBA modules. No extracted VBA source is stored.
- Official calendar evidence: `obj/ClientReportTests/ee6ea83d403c4ae58519dcb684b1eb6d` (three regions, fresh downloads, no cached fallback).
- Rent Weeks final Release regression: 58 assertions, `obj/ClientReportTests/8f2f195c14834f1ca08600854dc50b6f`. Debug shared fill-pattern regression: 17 assertions, `obj/ClientReportTests/745b4fdab9354165a858406465af6068`. Quiet-diagnostics checks pass in both configurations.
- All original input hashes remain unchanged. Known embedded-browser teardown error 1412 is outside this change; it appears after successful fixture assertions.

## Stable functional tests

Existing items 1–50 retain their numbers. [51 — Rent Weeks positive decimals](Rent_Weeks_2.88_2026-09-23.md) remains a separate test.

**52 — Funding date scheduler (Jon):**

1. Open Funding → Facilities and Loans → Repayments. Click the small header plus. Confirm native styling, readable sizing, scrolling and tooltip on your normal and 5K screens.
2. Preview monthly/quarterly/annual schedules, month end, first/last working day, count/end-date limits, and optional regional holidays. Change an option and confirm Apply waits for a new preview; cancel and confirm no changes.
3. Apply to a section with spare rows. Confirm dates use blank rows, duplicate dates survive, amounts remain unchanged, existing fill rules unlock the expected cells, and the first available amount receives focus.
4. Request more dates than fit on a disposable workbook. Confirm automatic expansion leaves five spare rows, lower Funding sections still display the correct fields, and entering amounts uses the intended loan.
5. Undo/redo the dates. Confirm added blank capacity remains as described. Save a separate file, reopen in Summit and Excel, and exercise the client's normal Excel/VBA calculation. Repeat in Variable and Cash Rates and Investments.

Physical keyboard/DPI acceptance, all older/bespoke workbooks, every target's expansion boundary and full financial/VBA results are not claimed by the automated tests.
