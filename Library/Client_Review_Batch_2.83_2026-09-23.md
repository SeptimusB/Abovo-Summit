# Client review first batch 2.83

23 September 2026. Alex is a director of Abovo and the client reviewer. The user clarified that "Alex to green" means Alex to confirm, not a request for a green Save button.

## Review decisions

- P004 Open/New/Compare icons: **Alex to test.**
- P005 icons within circles: **Alex to confirm on his PC.** The user repeated this item; no additional acceptance of P006 Program Information is inferred.
- P007 removed Start Date edit link, P008 Funding summary alignment and P009 unnecessary summary horizontal scrollbars: **Alex to confirm.**
- P012 DIT Save and its clean-file disabled state, P013 navigator width: **Alex to confirm.**
- P010 company-name refresh, P014 repeating year labels and P015 Return/History icons: changes below are implemented; client confirmation remains pending.
- Other source-report items remain at their previous status. In particular, P021 difficulty typing the company name is not assumed to be the same issue as a stale heading. Conditional formatting, older Other Fees mapping and other layout/definition questions remain open.

## Changes

Committed Global Assumptions changes now refresh the model metadata through its existing workbook profile. The HA company name reads the SelectTrust named input, with the previous C6 location retained as a legacy fallback. A changed company name or start date raises a per-model presentation event; unchanged metadata does not re-render. No extra workbook write, recalculation, dirty transition or interface rebuild is introduced. The existing edit/history service continues to own the transaction.

FileInstance HTML and its main-window tab, open group/Combined headings and sidebar File Details, and the open FFR company title subscribe to this event. Undo and Redo follow the same committed-change path. Current interface suffixes, the selected FFR sheet caption and Check Sheet warnings are preserved. Subscriptions detach during disposal. This covers Summit's managed edits; external Excel edits are reread on normal file open as before.

Ordinal-year header dropdowns now start with Year N and retain the existing financial-year suffix when available, such as Year 1 - 26/27. Blank, Year and Yr repeating captions use a consistent Year N representation, including the read-only first-year column. Other explicit captions and non-year repositories are preserved. Workbook values, number formats and year validation remain unchanged; selected items still post their numeric stored year.

DIT History now uses the former curved Return artwork. Conditional Return uses a padded left-arrow vector and is the last toolbar command, after its own separator. Both the command and separator are hidden on ordinary interfaces and shown for a return link. Existing command instances, click handlers, enabled states and origin tooltips are retained. The user was told that the far-right placement was interpreted as applying to conditional Return.

## Validation

- Release and normal Debug AnyCPU builds pass. After the user closed the session, `bin/Debug/Abovo-summit.exe` was updated to 2.83. No user process was terminated.
- `Tools/ClientHeaderToolbarFixture.cs` passes **31 assertions** against a private Demo copy, including production company edit/Undo/Redo, immediate FileInstance/tab/group/FFR refresh, HTML escaping, unchanged metadata, preserved warning state and FFR sheet caption, Return order/visibility/tooltip, and numeric/formatted/read-only year labels. Evidence: `obj/ClientReportTests/cc3cb796798843daa8b50d48222fdd2e/test.log`.
- Native DIT toolbar render inspected. Physical client DPI acceptance remains Alex's test, not inferred from this render.
- The normal Debug executable also passes all 31 native assertions: `obj/ClientReportTests/b5a9123e413e460caebb3fadb1697613/test.log`. Its source workbook hash is unchanged.
- Quiet diagnostics regression passes for Release, isolated Debug and the updated normal Debug build. Application tracing and passive benchmarks remain disabled; functional timers and normal messages remain. The existing third-party Chrome_WidgetWin_0 shutdown diagnostic persists and is not claimed fixed.
- Source Demo and Blank hashes, the original client report and review v01 are unchanged. No original workbook was saved.
- The pre-existing `FileInstanceInterface.Designer.vb` changes are the user's and were preserved. No commit or push was requested for this batch.

## Numbered review copy

`D:/Downloads/Summit v2.51 comments - review v02.docx`, with identical repository copy `Library/Client_Review_2.83_v02.docx`. The document workflow preserves all original issue wording and order, adds Alex's confirmation notes, marks the three new believed fixes orange, and leaves unrelated questions open. All ten pages were rendered via private read-only Word automation and visually checked; the packaged renderer was attempted but has no LibreOffice binary on this Windows host.

Regenerate from the untouched original using `Tools/Prepare-ClientReviewDocument.py --revision 02` and a new output path. The generator refuses to overwrite an existing revision. `--revision 01` retains the earlier report's content settings.

Status: **Ready to test** in Debug and Release. Client and financial acceptance remain separate.

## Checkpoint status update

The user's subsequent review supersedes the wording above: P004-P010, P012-P018, P029, P034, P042 and P048 are now **> Alex to test**. This is a test assignment, not financial or visual acceptance. P006 Program Information is explicitly included by the latest review. The remaining annotated questions and new Check Sheet policy changes belong to the next batch. The existing user-owned FileInstance designer rescale is excluded from this checkpoint.
