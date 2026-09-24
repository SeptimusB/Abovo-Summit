# Client UI and Funding scheduler review 2.93

Status: Ready to test with Jon. 24 September 2026. No commit or push requested for this batch.

## Scope and client decisions

- Alex's supplied conversation defines start date, inclusive end date and Monthly / Quarterly / Semi-annually / Annually, repeating the original day. The dialog now exposes exactly that date configuration and a single Date preview column. Short months clamp independently, so January 31 becomes February 28/29 then March 31; September 24 annually stays September 24. The service retains its older date-rule API for compatibility, but the dialog makes no holiday-calendar request.
- Facility, multiple sections, Summit-only tinting and optional fixed figures remain. The numeric edit mask now has sufficient integer placeholders; a preview refresh no longer reads/rebuilds the amount editor on every typed digit. Actual character entry tests cover positive and negative multi-digit decimals.
- The Funding record-header action bar is replaced by Add schedule… in the shared right-click menu, using the focused cell's actual workbook column. Section header schedule actions remain. Facility Name now reads the existing Facility named list, while Funder remains on Funders.
- Removed redundant ComboBox Enter/ShowPopup handlers that competed with the dropdown button click. Shift+Down joins native F4 / Alt+Down for popup inputs, including temporary header editors. Native ShowingEditor and workbook admission/locking are unchanged.
- Company-name width is bounded at 700 logical pixels. MergeDownAndPivot now transfers MinWidthChars for ordinary columns as well as repeating columns. The native Stock grid has its 42-character description minimum and an existing 50-pixel separator after Current Stock Numbers; no extra worksheet column is added.
- Funding ordinary wheel follows rows/page vertically; Ctrl+wheel changes visible loan records horizontally. The application-wide filter must compare the actual top-level group form, not DIT.FindForm (DIT itself is an embedded XtraForm). Dialog/Options ownership and enabled-state guards remain in place.

## Deliberately open

- Jon passed save/break and persisted-file-warning tests: review v05 marks them `> Alex to test — Test response 2`. Management Costs readability and link indentation receive the same stage. This is not Alex's green sign-off and does not imply a new commit.
- The public Check Sheet heading not refreshing after a warm switch and an unspecified Transactional DB remainder are still unresolved. No validation state was forcibly cleared, no financial formula was changed and no TDB repair is claimed. The existing opt-in full-calculation watcher remains unchanged. Need changed input, Check Sheet row and stale TDB address for the next reproduction.
- Shift+wheel grid zoom is feasible, not enabled in 2.93. DevExpress recommends FontSizeDelta rather than repeatedly allocating Font objects; column appearance can override row fonts. Summit's source-cell styling/custom painters and row/header/editor heights need a separately tested presentation-only implementation. Sources: https://supportcenter.devexpress.com/ticket/details/t1253313/gridcontrol-zooming-doesn-t-work-bold-columns and https://docs.devexpress.com/WindowsForms/DevExpress.XtraVerticalGrid.VGridOptionsBehavior.RecordsMouseWheel .

## Validation

- Normal Debug and Release builds pass at C:/Repos/Abovo Summit/bin/{Debug,Release}/Abovo-summit.exe, version 2.93. Logs: obj/build293-debug.log and obj/build293-release.log.
- Tools/Client293Fixture.cs: 27 assertions pass in Debug (obj/ClientReportTests/439396dfa001467994ffa737d6f1e0df) and Release (obj/ClientReportTests/e1c8c8fe3d224d98b608e741414e9e2a). Includes actual mask-textbox characters, one-click acceptance, popup-button mouse messages, selected-column mapping, Stock native bounds and real wheel messages. Ctrl changes LeftVisibleRecord 0 to 1; ordinary wheel advances page scroll 0 to 120 while holding the loan column.
- Tools/FundingScheduleFixture.cs updated for the context menu and end-date UI: 81 Release assertions pass, obj/ClientReportTests/2249ce1122764209841a2e782a2e2fc6. Includes modal scheduler, automatic preview, multi-section Apply, dates/amounts, focus, rendered tints, Options/dialog wheel exclusion, save/reopen and all 14 date-range comparisons. This UI run does not repeat the expensive row-capacity expansion cases; existing structural services are unchanged.
- Tools/FundingGroupsFixture.cs amount mode: 17 Release assertions pass, obj/ClientReportTests/9fce8a04912d40d9a30a7b411f99a36c. Covers date-first calculation, amount and percentage entry, one Undo/Redo, other-loan preservation, rejected locked/mixed-unit rollback and XLSB save/reopen.
- PatternEditabilityFixture: 17 Release native assertions pass on a private Blank workbook (obj/ClientReportTests/0b7c26b0fd6549e785bfbf9ae1193b6f), including editor/paste admission, undo/relock, clearing and unchanged underlying fill patterns. The first wrapper misread a disposed process ExitCode; the harness now holds the process handle before waiting, and the repeat run exits zero. No lock-rule implementation changes were made.
- Native WebView teardown emits Chrome_WidgetWin_0 unregister error 1412 in some fixtures. Successful final fixture processes return zero; no unhandled application exception was observed.
- Original masters remain unchanged: Blank SHA256 E05274ACD3D4821AC38F013AC45A7B57CE335BD524938D9F21E1B640D0CC5E90; Demo SHA256 1ED79726D8D1129C699242C9E8B3E988AC920EA7108D3C0B105BF9B0BC8A5E9C.

## Word review and acceptance

Library/Client_Review_2.93_v05.docx follows v04 without overwriting it; both prior screenshots are byte-identical. Original v04 SHA256 F2361BF8B0FD03C6182DCF9551201FD7D4A9738AD3006D4BAEAC15C0950CB0E7. Final v05 SHA256 4BE3672F89AC8D83570476C88C0366BFDC6C0FF3685AAECA8C6E3C8DDFB5412B.

The canonical renderer could not find LibreOffice. The installed DevExpress RichEditDocumentServer exported the review; all 14 rasterised pages were visually inspected, including a second check of the corrected mixed-colour issue paragraph. Source screenshots and green sign-off text are retained. Blue means Jon to test; orange heading and light-blue detail mean Alex to test. New stable functional items are 61–67; earlier unconfirmed blue items remain outstanding. Client-scale mouse/keyboard use and Excel/VBA financial round-trip acceptance are still required; automated tests are not user sign-off.
