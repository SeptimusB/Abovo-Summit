# Mapped tables and Check Sheet - 20 September 2026

Test release: 2.36. Structure XML revision: 1750.

## Accounts Assumptions investigation

The existing Accounts definition has 64 mapped rows and creates 151 individual WinForms/DevExpress controls. Initial construction already suspends layout. However, the later font pass runs after layout resumes and assigns fonts through the whole control tree, causing repeated AutoSize table layouts.

A scoped repair suspends the mapped TablePanel during that font pass and resumes it before any grid geometry/best-fit work. It does not replace the existing Accounts editors, validations, ModelChangeManager edit routing, calculation strategy, or mapping.

Same-machine, isolated fixture measurements using the Demo master:

| Phase | Before batching | After batching |
| --- | ---: | ---: |
| Registration / calculation | 146 ms | 143-149 ms |
| 151 control creation | 181 ms | 173-176 ms |
| Initial font pass | 293 ms | 11-12 ms |
| Complete Accounts constructor | 1,110 ms | 727-747 ms |

These are fixture results, not a guarantee for the user's full desktop, 5K/DPI setup, or production process bitness. Other costs remain: workbook calculation, native control creation, initial handles and layouts. Targeted `[Mapped Table Benchmark]` traces distinguish calculation, control creation, font application and read-only range construction.

Native TablePanel batching guidance: https://supportcenter.devexpress.com/ticket/details/t964794/speed-up-adding-a-control-to-a-tablepanel

## Check Sheet definition and navigation

Global Assumptions (GSID 0, CSID 0) now has a Check Sheet tab between Imported Files and Abovo Contacts. No CSIDs or navigator identities changed. The XML definition is a MappedTable with optional compact read-only projection:

- Worksheet: Check Sheet.
- Displayed range: A7:G63; headings from row 6.
- Read-only by default. `AllowYesNoEdits` opts into native combo editors for unlocked, non-formula cells whose workbook list validation is exactly Yes/No. Both current masters define C22, C23, C29, C33 and C39; addresses are discovered, not hard-coded in production. Unsupported or absent validation fails closed.
- Workbook display text, fonts, colours and alignment are used. The Check Sheet formula-based error font colour is evaluated from the workbook rule, not inferred from the word Status.
- Section headings, intentional blank rows and the Total row remain present, including when the workbook is blank.
- Cell multiselect, Ctrl+C, Copy and Copy with headings use the shared clipboard implementation. Filtering and sorting menus are disabled.
- Widths account for workbook widths, actual data/header fonts and DPI. Columns are explicitly defined, preventing lazy handle creation from replacing captions and widths. Long messages wrap.
- The visible G markers are NOT native spreadsheet hyperlinks: H contains the actual worksheet name. Never resolve solely from the descriptive A caption.
- With `ShowLinkDestinations`, G is displayed as an explicit Worksheet destination from H. Clicking it (or pressing Enter) opens the current model's existing global spreadsheet window at that sheet. A separate, UI-only Summit interface column lists the matching group/interface/section destinations and opens a native choice menu. These captions do not write G or H. Older definitions without the option retain the combined link menu.
- Matching is based on declared default/source worksheets in the active model's structure. Multiple legitimate interfaces are offered rather than selecting an arbitrary one.
- The XML supplies an explicit Transactional DB -> Outputs / Analysis route for the class-based analyser. Visible Analysis and Funding Assumptions names no longer include V2; GSIDs/CSIDs and the current V2 implementation classes remain unchanged.
- Existing EventCoordinator routing preserves Combined-window handling and the separate group windows. A return link points back to the originating interface.
- Missing worksheet targets are displayed as unavailable; workbook content is never treated as a command, URL or file to execute.

Existing per-control MappedTableRow definitions remain compatible. This is not a workbook schema migration; no XLSB or embedded structure parts were rewritten.

## Refresh and calculation

The mapped worksheet is registered with interface dependencies and active worksheet bookkeeping. The new grid reads updated values in place on normal DIT refresh, without rebinding or clearing selection/widths. Reactivation restores its worksheet registration. Selecting an already-built Check Sheet after a same-DIT edit brings stale calculation current through the normal staged CalcFile path; no new calculation engine or formula logic was introduced.

The formula-colour projection currently covers formula-based font-colour rules used by both authoritative Check Sheets. It is not a general renderer for icon sets, colour scales or every possible future conditional-format rule. Review the definition/range if a future master extends the Check Sheet past row 63.

Yes/No changes go through `ModelChangeManagerV2.ProcessChange` with string values and meaningful override descriptions. Normal calculation, dirty marking, undo/redo and failure restoration remain in that service. The grid rejects arbitrary text, locked cells and formulas. History notifications refresh the workbook-backed view and close a stale in-place editor. No new calculation engine, direct interactive cell writer or protection toggle was introduced.

## File-instance housekeeping

Top navigation is now Assumptions, Workings, Outputs, Combined, FFR, Stress Test, Spreadsheet. The native placeholder button is labelled HA BP; its caption comes from the workbook profile (DSA is provided separately) and remains without an action. Existing designer/user layout changes are preserved. The badge and file actions share a left column, with the badge aligned to the navigation row.

## Verification

Tools/Test-MappedTable.ps1 builds an isolated executable under obj/MappedTableTests. It loads the actual application services, constructs Accounts and Global Assumptions, and builds the Check Sheet DIT section without saving the workbook.

Verified on both Library/Demo BP v26_0001.xlsb and Library/Blank BP v26_0001.xlsb:

- 57 rows x 8 display columns. All 342 A:F values match their worksheet cells; all 57 Worksheet labels match G/H targets. The extra column is a UI-only navigation projection, not a workbook column.
- Exactly five native Yes/No editors and cell multiselect. Normal GridView changes write the workbook and mark dirty; arbitrary text and formula edits are rejected. Undo/redo refresh displayed values and preserve worksheet protection. An active native combo commits through the grid event; undo closes the stale editor. Other Override rows cannot open an editor.
- All 27 G destinations resolve to existing sheets and at least one Summit interface. Primary Analysis and Funding Assumptions labels are present without V2.
- Selection survives refresh.
- A temporary in-memory B9 test changes only the corresponding E9 conditional colour, not E10; the original value/formula is restored.
- First-display captions and description width are retained; rendered grid inspected.
- Source XLSB SHA-256 unchanged.

Debug and Release builds passed after final changes. Interactive acceptance remains necessary:

1. Open Accounts and confirm responsiveness, existing combo/text editing, validation and undo.
2. Open Global Assumptions -> Check Sheet on Blank and populated files; verify headings, total, statuses and messages.
3. Change each of the five Yes/No overrides, then undo/redo, including while its editor is open. Confirm checks recalculate. Other cells must remain read-only; invalid text must not be accepted. Verify rectangle copy with and without headings into Excel.
4. Open explicit Worksheet destinations and Summit interface choices. Test Funding, Development (multiple matches) and Analysis. Test from separate and Combined windows.
5. Open worksheet twice and confirm reuse of the model's global spreadsheet window and correct target sheet.
6. Edit an assumption in a second visible group, recalculate, and verify Check Sheet updates. Repeat an edit on another Global Assumptions tab, then return to the retained Check Sheet.
7. Hide/reactivate, use the DIT full rebuild button, and repeat links/copy. Check 5K and client DPI layouts and long error messages.
8. Standard Summit Save As -> Excel/VBA reopen -> Summit reopen remains a manual smoke test, including retained Yes/No values. This change does not alter workbook formulas, names, macros, validation or schema.

`Tools/Test-FileInstanceLayout.ps1` also verifies both model captions and native button order/hit targets across nine width/font combinations (750/1100/1600 px; 9/14/18 pt), without opening a workbook. Layout render inspected. This is not a substitute for physical monitor/DPI checks.

Prior Model Manager trial work and the user's FormMain designer/resource edits were preserved. This task did not commit or push.
