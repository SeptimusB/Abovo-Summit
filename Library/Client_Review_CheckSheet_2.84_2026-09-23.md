# Client review and Check Sheet trial 2.84

Status: **Ready to test** with Jon. The uncommitted 2.84 batch requires Jon's functional test before commit and client handoff. Alex's client test and accountant acceptance remain pending.

The 2.83 batch was committed as `7bd32e1` and is ready for Alex to test, not signed off. The separate user-owned FileInstance designer changes and `.codex/` were excluded. No push was requested for this batch.

## Check Sheet and integrity policy

- A Check Sheet financial imbalance is a soft notice, separate from actual formula errors, broken names and structural integrity findings. Numeric imbalances remain available in the integrity report as balance notices; error-valued Check Sheet check/status cells remain integrity errors.
- Check results update session state only. Remembered pause settings are updated after a successful manual Save/Save As, or against the recovery copy's own path after a successful recovery write. The remembered state is application configuration, not source-workbook XML. Opening, checking and discarding an unsaved test do not rewrite the source file.
- Remembered warnings are hidden pending an automatic quiet check on open. A balanced current file clears the in-memory warning. No confirmation popup or initial red heading is required. The existing revision, editor, operation and UI-thread gates remain.
- Recovery continues through soft balance findings by default. Explicit existing opt-out settings are retained. Recovery still requires unsaved user input, valid operation boundaries and the existing safety checks.
- The new Options **Check Sheet trial** tab is opt-in, default five-minute interval and two minutes idle. It calculates just the Check Sheet after a workbook revision changes, updates the balance message on state transitions, and does not certify the full workbook or clear rebuild/pending-results flags.
- Trial timings are separately opt-in and cover the worksheet calculate/read and single-cell typed-edit path. They do not enable general tracing. Bulk-command timing is not added by this trial. Once the atomic sheet calculation starts it finishes safely; the watch does not interrupt a save or pending edit.
- The user approved temporarily replacing the monthly x86/x64 reminder with a quiet trial-acceptance monitor. Once the trial is accepted it reminds the user to disable trial timings and restores the original monthly reminder.

## Input and presentation changes

- Workbook Yes/No validation now writes literal text even if an older interface declares the input Boolean; blank, Undo and Redo retain the intended values and source number format.
- The read-only grid paint path uses source colours instead of forced white-on-lavender text. Rule refresh also refreshes cached source colours and bold state. This is not proof of complete Excel conditional-format parity.
- Company-name input spans more available width; Stock Description has a 42-character minimum; a dummy separator separates the requested stock groups without inserting worksheet columns.
- Leading XML whitespace is trimmed from section descriptions; link buttons span more width and auto-size.
- Specific Income Summary Other Income Category has no Add Lines action.
- Management Costs Description and both category columns are pinned left. Year 1 headings stay read-only, amounts stay editable; no additional ROInitialLines restriction was introduced.

## Validation evidence

- Debug and Release builds pass with 2.84 and Structure XML revision 1755.
- `Tools/ClientReview284Fixture.cs` passes 35 native assertions in each build on disposable Demo copies. Release evidence: `obj/ClientReportTests/165f94cdbd20418898c64def03ce288d/test.log`. Debug evidence: `obj/ClientReportTests/b2c11b8929cf4a099ec1b3134a36ee1b/test.log`.
- Tests include literal Yes/No and clearing/Undo/Redo; no settings persistence from an unsaved check; recovery opt-in/opt-out; hidden historical state; formula-error classification; automatic-open clearing; idle/save gates; unchanged-revision suppression; preserved pending/dirty/history/calculation state; pinned descriptors and source colour cache; wider company input; separate trial tab; and successful Save persistence/clearance.
- Source workbook hashes are unchanged. Original client/master files were not saved.
- Quiet diagnostics tests pass, including explicit runtime trial output off/on while general diagnostics remain disabled. Existing third-party Chrome_WidgetWin_0 shutdown output is not claimed repaired.
- Full conditional-format refresh, other layout/scrolling requests and ambiguous client items remain open in the report. Client DPI, physical keyboard, real-file discard/reopen and Excel/VBA financial round-trip acceptance remain manual tests.

## Working report

`D:/Downloads/Summit v2.51 comments - review v04.docx` supersedes v03 as the working report; `Library/Client_Review_2.85_v04.docx` is the repository copy. It includes the subsequent 2.85 UI scaling crash repair, also blue for Jon. The original annotated v01, v03, and both source images remain unchanged. The user prefers Word for issue narratives, screenshots and one-to-many discussion; no replacement spreadsheet tracker is requested.

- Blue highlight and **> Jon to test**: latest uncommitted 2.84 changes require Jon's functional test.
- Orange issue heading, light-blue detail and **> Alex to test**: committed fixes ready for the client's testing, not signed off.
- Green font: **Agreed**, only when Alex's returned Word document has the item green. One original green issue is retained. A scope agreement, successful build, automated test or commit is not client sign-off.
- Open questions remain neutral; partial repairs have their testable part separated from unresolved work. Jon's pass and commit are required before promoting the latest batch to Alex. This replaces both previous yellow and undifferentiated test highlighting.

`Tools/Update-ClientReviewStages.py` checks the source hash, uses explicit owner mappings, refuses to overwrite a revision, and verifies image preservation and absence of superseded yellow highlighting. The later scaling crash was repaired separately as test release 2.85; see `Presentation_Scale_Race_2.85_2026-09-23.md`. The 35 client-review assertions pass again in both 2.85 builds. The final v04 was rendered using the native document engine and all twelve pages were visually checked; no original report or screenshot was overwritten.

`Tools/Update-ClientReview284.py` refuses to overwrite previous versions. The packaged renderer has no bundled LibreOffice on this host. Independent read-only Word PDF export stalled twice; only the task-created processes were stopped. The installed DevExpress document engine supplied the successful PDF render, with all eleven pages rasterised and visually inspected. PDF and PNG files are internal QA, not user deliverables.
