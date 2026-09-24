# Funding schedule groups and Settings — 2.90

Status: **Ready to test** with Jon. Not committed or client accepted. Normal Debug and Release executables updated.

## Behaviour

- Settings has OK and Cancel; OK validates, applies and closes. The redundant Apply button is removed.
- The Funding scheduler automatically refreshes a valid preview after a short editing pause. Changed options invalidate the old preview immediately; an asynchronous holiday response cannot re-enable a stale configuration. OK adds the displayed dates. There is no mandatory Preview button.
- Section-heading plus retains its section-specific action, now with a facility/loan selector and colour choice. A new plus in each facility's record heading offers multiple compatible sections across the Funding tabs. Ordinary loans have eleven sections; investments have their separate three sections.
- Each section uses its first safe blank dates, retaining duplicates and existing entries. Insufficient capacity adds the shortage plus five spare rows per section. All dates across selected sections form one ChangeManager undo group; added structural capacity is not undone. The whole operation excludes background recovery saves.
- After application the appropriate tab opens and focus goes to the first new date's amount cell for the selected facility. If workbook rules keep that amount unavailable, the defining date receives focus; no input lock is overridden.
- A user-selected colour becomes a bounded very pale tint on that facility's scheduled amounts and defining dates. Selected/locked/read-only cues take precedence. Excel fills, formulas, amount values, protection and fill-based admission rules are not changed by the colour.

## Persistence and identity

Uses the installed DevExpress 25.2 native `IWorkbook.CustomXmlParts` API. A separate part has namespace `urn:abovo:summit:funding-schedules:1` and root `FundingSchedules version="1"`. Each Schedule records a GUID, original colour, creation time, configuration summary, facility anchor/header name, and Date entries with section/value-range names, expected date and anchor.

Unique hidden `_SummitSchedule_<GUID>_...` defined names anchor the selected facility header and each physical date cell. No new worksheet is created. Names follow native Excel/Summit row/column insertions rather than matching ambiguous duplicate descriptions or dates. Unrelated XML parts are retained. Opening/previewing adds nothing; explicit scheduling changes only the in-memory workbook. Existing Save/Save As and recovery export persist the XML and anchors. Discarding an unsaved session does not alter the original file.

Undo leaves dormant presentation metadata; it has no visible effect while the expected dates are absent, and becomes active on Redo. Deleted/broken or mismatched anchors are not guessed or reassigned. A deleted hidden name is exempt from financial broken-name reporting only if the schedule XML explicitly owns it; unrelated broken names remain reportable. This is not a general corruption suppression policy. Undo of later edits can naturally reactivate a group when its original dated cells return.

All sections are preflighted before expansion; addresses are resolved again after every section has expanded. Metadata is removed/restored if date posting fails. Existing structural-service failure/recovery safeguards remain; the feature does not claim whole-workbook rollback of added rows. The error message explicitly reports retained added capacity.

## Verification

- Debug and Release build successfully. Quiet diagnostics gate passes; no tracing/benchmark opt-in was enabled.
- Final Release native UI/Settings fixture: 65 assertions (`obj/ClientReportTests/a7f6981e81e5471e810d863ae6a240d3`).
- Full Debug scheduler fixture: 168 assertions, covering fourteen sections, three expansion paths, protection, shifted input anchors, duplicate dates, date Undo/Redo, native section/facility plus clicks, auto-preview, multi-section apply, date tint and safe focus (`obj/ClientReportTests/70bf675247b04d98bf34d4374bf729e0`).
- Release metadata fixture: 21 assertions. Compares physical fills against ordinary uncoloured typed date entry (Excel's own conditional patterns legitimately change when dates are added), grouped undo/redo, incompatible-section preflight rejection, native XLSB and fully prepared recovery XLSM reload, unrelated XML retention, row/column movement and deletion (`obj/ClientReportTests/22026ab09dfb42e29c569a559de570d7`).
- Release multi-expansion fixture: 16 assertions. Lower section expanded before upper section; both retain five spare rows, earlier anchors/tints survive, grouped Undo/Redo passes, recovery exclusion is released (`obj/ClientReportTests/3bd8af75105848159f937a3e7866e4ae`).
- Existing blank-model fill-based editability regression: 17 assertions pass (`obj/ClientReportTests/5ee98b3f53994c96b915fc210b3a171b`). Native editing/paste/Undo/Redo and locked-cell cues retained.
- Excel read-only open and separate SaveCopyAs, with macros/events/calculation disabled, retains schedule XML, sentinel XML, all 1,684 names including hidden state, sheet order, and 6,231 Funding cells' formulas/constants/locks/fill patterns (`obj/FundingScheduleExcel/1425ee53ddaa4f67a1c516cea1ba3302`). All 337 VBA module source hashes match the untouched Demo master, Summit copy and Excel copy. This does not certify running VBA or financial recalculation.
- Native dialog/header screenshots were inspected. Original workbook hashes unchanged. No client recovery file deleted, no master changed, no production load-time schema migration introduced.

The reported checklist InvalidCastException was reproduced in the disposable test runner and corrected: `CheckedListBoxItem.Value` supplies the section, not the wrapper item. The harness now propagates UI exceptions to its log rather than presenting an unattended JIT dialog. Existing WebView teardown error 1412 after successful assertions remains outside this work.

## Stable functional tests — Jon first

54. **Settings:** confirm Apply is gone; alter an option, press OK and reopen to verify persistence. Cancel a different change and verify it was not applied.
55. **Facility schedule:** use a facility-heading plus, select at least two sections and a colour, change dates/interval/count/end date and confirm automatic preview. Apply, check separate dated rows, selected-loan focus or safe date fallback, and one Undo/Redo across both sections. Also use a section-heading plus. Include a shortage requiring expansion.
56. **Colour persistence:** check subtle date/amount tints and unchanged unavailable-cell cues. Save As a disposable XLSB, reopen in Summit, then test trusted Excel/VBA and return to Summit. Test a recovery copy too. Check your display scaling and Alex's PC before client acceptance.

Earlier functional items 1–53 retain their numbers. The recovery-delete feature from 2.89 is now included in normal Debug as well as Release. Outstanding acceptance remains Jon's functional checks, Alex's layouts/workflow, and accountant/Excel-VBA validation. The generic structure manager is not implemented or expanded by this schedule-specific XML part.
