# Production interface changes — 2.97

Status: **Ready to test**. Jon functional acceptance and Alex's workstation tests remain separate from automated evidence. No commit or push requested.

## Scope delivered

- All centrally configured native grids: **Ctrl + wheel zoom**, **Shift + wheel horizontal**, ordinary wheel vertical. Context-menu hint updated. Existing pointer anchoring, 60% minimum, fonts/editor sizing and modal/popup guards retained.
- Funding schedules select the earliest complete consecutive run of safe empty rows. Scattered holes are not combined. If no run fits, extend the safe trailing run by the shortage plus five spare rows. A date with an existing amount or formula in any loan is occupied, including numeric zero. Protection/fill admission is unchanged.
- The schedule preview uses the same placement plan as Apply. After all selected sections expand, ranges are resolved again and actual written addresses drive focus. Dates, fixed-amount soft skips, XML tints and grouped Undo/Redo retain their established behavior. Added blank structural capacity still remains after Undo or a later write rejection; this release does not invent structural Undo.
- Model Version already had `RO=TRUE` in Structure.xml. The text-editor construction now applies that existing read-only/calculated contract before initialization. It remains enabled/selectable; Company Name stays editable. No workbook value, schema or fill rule was changed for this repair.
- Alex's double-click report remains unconfirmed. Both header and active-editor event routes pass isolated checks. A pending defining year/date is not committed by double-click itself, so locked destinations can cause a silent no-op. This is a plausible distinction, not a diagnosis of Alex's machine; production behavior is unchanged pending his retest/call.

## Calculation: investigation only

[Exact AGL investigation](Check_Sheet_Funding_Diagnostic_2026-09-24.md) reproduces G82 2700 → 3000, Check Sheet, History Undo, navigation and the residual Transactional DB Cashflow result. It separates cached deferred-sheet results, company-heading status publication, and visible-grid refresh. There is no evidence of a fixed one-sheet dependency-depth limit. An extra deferred pass cleared the residual in this specimen; this is not a general validated repair. Full watcher calculation fixes workbook results and publishes the heading but can leave the mapped grid stale.

No calculation policy, warning persistence, Undo calculation implementation, Gear integration or workbook master was changed in 2.97. Do not treat the current warm Check Sheet / Undo refresh as accepted. Integrity and independent Excel/accountant checks remain relevant; this diagnostic does not validate every other calculated UI chain.

## Validation evidence

Both normal Debug and Release build successfully (`obj/build297-debug.log`, `obj/build297-release.log`). Quiet diagnostics tests pass for both; the explicit Check Sheet timing opt-in remains the existing exception. Native desktop tests here are 96 DPI and do not replace physical 5K/client acceptance.

| Check | Result | Evidence under obj/ClientReportTests |
|---|---|---|
| Generic GridView/VGrid/Tree pointer zoom, readable geometry and new wheel mapping | 81 assertions each build | Debug `2807e8c59d544cfcb49e49fd4bb86751`; Release `60efd873e6fb4995bdfb55cabeb0b2ec` |
| Actual Global Assumptions Model Version: native typing rejected, value/revision/dirty unchanged, selectable text and editable Company Name | 9 each build | Debug `c026d28b94a8401b9a0c0550d261a6a6`; Release `02fd1343fc934e39b811c4019a3c1cbf` |
| Actual AGL Model Version, same checks on private older-client workbook | 9 Release | `7bd735caf0a9485e93e78675a22a6d1e` |
| Actual-model contiguous placement, orphan amounts/formulas, lock/fill, exact fit, expansion + five, shifted sections/focus, tints, grouped Undo/Redo, failure rollback, native Save As/reopen | 39 each build | Debug `d9bde9c0b0ff4db1a7efa910807d797c`; Release `bec4adec27974a3b981268d311eef968` |
| Existing fixed figures, percent units, locked-target soft skips, Undo/Redo and rejection rollback | 21 Release | `3887b2508b12448f8c27d1f02c0f9e70` |
| Macro-disabled read-only Excel open → SaveCopyAs → native reload of new contiguous-schedule artifact | Pass: 1,702 names, sheet order, 6,627 Funding cell values/formulas/dates/locks/fills and schedule XML preserved | `obj/FundingScheduleExcel/0bdeeff75827432bb1f92c22ee9ba273` |
| Double-click callback/forwarding/timing and read-only Rent eligibility | 18 + 20 on existing Release 2.96 | `949a3afeb9474e038227fdcdb57ecd2a`; `783967848df44540b348b7838f363c66` |
| Exact AGL calculation investigation | 7 setup/edit/restore assertions plus recorded result transitions | `ee4d85b1ea634daeb4426353b5329357` |

The pre-existing `Chrome_WidgetWin_0` teardown diagnostic appeared in the textbox tests; no claim is made to repair it. Original masters and AGL are untouched; all test arrangements/saves used disposable copies.

## Delivery and functional gate

- Test executable: `bin/Debug/Abovo-summit.exe` or `bin/Release/Abovo-summit.exe`, version **2.97**.
- Runtime version verified in both assemblies. Debug executable SHA256 `696F3741F993FF563B62986EDC1349554ED4D7AD03A1E53363605F303B20DFCD`; Release `464B3EE68633F72BBDAE91E491F3B5B7D9B404A3150A002429435DCBF223262C`.
- User's frozen `bin/Working-2.96/Abovo-summit.exe` is unchanged: SHA256 `AD4800FD6F211B9FB635E1F1CE917FFB5EE201D63789F4C926ABED10D42B493C`.
- Repository Blank SHA256 `E05274ACD3D4821AC38F013AC45A7B57CE335BD524938D9F21E1B640D0CC5E90`; Demo `1ED79726D8D1129C699242C9E8B3E988AC920EA7108D3C0B105BF9B0BC8A5E9C`, unchanged.
- [Current numbered functional ledger](Functional_Test_Ledger_2.97_2026-09-24.md): IDs **1–86**, including amended **68/82**, new **84/85**, and unconfirmed Alex retest **86**. Earlier IDs and partial handoffs are retained.
- Review Word v05 was read, not changed. No new Alex/financial acceptance is inferred from automated tests or general positive feedback.
