# Presentation, Journals and funding-payment dates — test release 2.53

## Scope and safety

Requested changes are confined to presentation, the Journal records structural rule, and the First Interest Payment Month input. No original XLSB master or client workbook was modified. Existing uncommitted 2.51/2.52 work is retained. The client reported that the preceding Balance Sheet functional checks passed; accountant financial acceptance remains separate and pending.

## Presentation

- FileInstance uses a script-free, encoded, IE-compatible HTML summary card: model name/type, plan start, file size/path and access details. The Edit link is removed. Opened time is stable for the instance lifetime rather than changing during resize.
- Check Sheet mapped grids hide horizontal and vertical gridlines. Workbook colours, conditional errors, Yes/No editing, links, selection, copy and undo remain in place.
- Balance Sheet TreeList inherits SOCI row/header fonts and header/line appearance. Headings/totals use the analyser lavender/steel-blue hierarchy, hot tracking is blue/white, selected cells wheat/black, and negatives red. Workbook numerical values/formats, source explanations, hierarchy, Live/Snapshot/Comparison and export are unchanged.
- Sidebar defaults are applied once: BP Status and Funding Status expanded; all other containers collapsed. Subsequent refresh/resize does not reset the user's choices or the 2.52 compact/restore strips.
- Funding summary lower J3:N5 block is padded one cell to the right, placing YE Net Debt under YE Peak Debt. The native workbook is untouched. A bespoke workbook with data in the trailing cell falls back to unshifted rendering rather than dropping any data.

## Journals diagnosis and approved repair

XML GS0/CS42 correctly declares `JOURNAL_RECORDS`, `ProcessAddJournalRecords`, `ProcessDeleteJournalRecords`, `Rep_Jour_01`, and one excluded final record. The defect was in the structural rule, not the XML routing.

Reproduced on a disposable Demo copy: `Rep_Jour_01` grew from A7:E27 to A7:E32, but `IR_Journals` remained A7:A26, and both Transactional DB mirrors retained their original 21 rows. The insert lies just outside IR's last row because the presentation range includes a sentinel and IR does not.

The rule now snapshots/resizes IR_Journals alongside the presentation range, exactly once and before the existing synchroniser runs. IR remains one row shorter than Rep_Jour_01. Both debit/credit mirrors then grow through the existing synchroniser; no new sheet or schema is introduced. Delete uses the same snapshot logic, including non-contiguous selections. Dirty state, protection and existing partial-failure/recovery handling remain in force.

Previously damaged files are NOT silently repaired. A pre-existing IR/Rep geometry mismatch is rejected before mutation with an explicit message. Such files require individual review; this release prevents the defect on consistent files.

## First Interest Payment Month

`Funding Assumptions!E137:R137` holds dates displayed as `mmm`. Excel list validation references `FundMonthDesc`: twelve actual plan-year month-end dates (April 2026 to March 2027 in the tested masters), not unrestricted month names or arbitrary month ends. Existing Demo entries include older dates; merely opening the interface must not rewrite them.

XML previously declared string `S` with `Rep_FundMonthDesc`; it now declares date `D`. The native DateEdit displays `mmm`, edits a full date and disables calendar dates outside the workbook-defined list. ChangeManager independently validates the same list for single changes and paste, allowing blank only where Excel allows it. Recognised valid dates remain genuine dates even if a caller uses an older text-type definition. Invalid dates are rejected, not silently rounded or coerced to a different month/year.

DevExpress reference checked against the installed 25.2 build: [DisableCalendarDate](https://docs.devexpress.com/WindowsForms/DevExpress.XtraEditors.Repository.RepositoryItemDateEdit.DisableCalendarDate?v=25.2). Native builds and editor tests verify the API actually used.

## Validation evidence

- Debug and Release builds passed (normal bin outputs; test version 2.53).
- `Test-FileInstanceLayout.ps1`: existing button/layout checks plus native HTML render/DOM checks at 360, 800 and 1400 pixels, escaping, full metadata, no link and no horizontal overflow. The 800px native render was visually inspected.
- `Test-MappedTable.ps1 -CheckSheet`: values/formatting, links, five workbook-backed Yes/No inputs, validation, undo/redo, selection preservation and disabled gridlines passed.
- `Test-PresentationLayout.ps1 -SidebarOnly`: initial expansion, exact lower-block column alignment, stable content, refresh, width changes 1280–5000 and compact/restore regressions passed.
- `Test-AnalyserChart.ps1`: Balance Sheet font/header parity, hot-node/selection operation, 42 columns, expansion/focus/width preservation, numerical export parity, chart drill, Live/Snapshot/Comparison, edit/undo/redo and deferred safety passed. Native Balance Sheet render visually inspected.
- `Test-EditorNavigation.ps1`: actual date editor and full-date write, undo/redo, mid-month/out-of-list/invalid rejection, invalid multi-cell paste rollback, legacy text caller preserving date type, and all previous Funding/Rent navigation regressions passed.
- `Inspect-JournalFunding.ps1 -Insert`: Demo and Blank copies passed five-row insertion, both 26-row mirrors with formulas, source-name geometry, dirty/protection, XLSB save/reopen, delete back to original geometry, and non-contiguous deletion.
- `Test-JournalExcelRoundtrip.ps1`: expanded Demo copy opened read-only in separate Excel with macros/events/links disabled. IR=25 rows; Rep and both mirrors=26. Journal Check Sheet B53/D53=0, E53=OK. Excel SaveCopyAs and native reload retained the four names.
- `Verify-VbaModuleHashes.py`: all 337 VBA module identities/source hashes identical in original Demo, Summit-saved expanded copy, and Excel-saved/reopened copy. No VBA execution was attempted and no extracted source/passwords are stored.

Evidence lives in ignored `obj/*-253*.log` and native fixture output directories. Workbook SHA-256 checks assert original files unchanged.

## Client acceptance / remaining risks

1. On the client's 5k and ordinary displays, inspect FileInstance summary wrapping and Balance Sheet font/hover/selected-cell colours in restored and maximised windows.
2. Verify both summary panels initially open, all others closed, manual expansion survives refresh, and compact/restore retains both edge strips.
3. In Funding choose a valid payment date, confirm `mmm` when idle and the full year/date when editing; test Tab/arrow continuation and paste/undo. Existing out-of-list dates are retained until deliberately edited. If clients intend ANY month-end outside the plan-year list, that is a separate workbook-validation policy change, not assumed here.
4. In a fresh consistent disposable workbook add/delete Journals, enter balanced debit/credit entries in new rows, calculate and inspect Check Sheet and analyser. Save in Summit, reopen/use intended VBA in Excel, then reopen in Summit. Automated Excel checks did not execute macros or certify financial results.
5. Send already-affected client files for a separate named-range/data-integrity review rather than bypassing the new guard.
