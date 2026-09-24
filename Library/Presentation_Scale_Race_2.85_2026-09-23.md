# Presentation scaling race - 2.85

Status: **Ready to test** with Jon. No commit or client sign-off is implied.

## Reported failure and repair

The user reported an unhandled `InvalidOperationException`, collection modified, at `PresentationScaleManager.ApplyScaleToNewForms` in FontManager.vb line 255. The original code copied `Application.OpenForms` by enumerating the live process-wide collection. Another UI thread opening or closing a form could invalidate that copy before any scaling callback ran.

- Snapshot capture retries only the collection enumeration, at most three times. It discards partial copies and defers to a later idle pass if the collection remains busy. No collection lock is held around UI callbacks and there is no unbounded spin/retry.
- Each form retains its applied scale. A deferred pass therefore catches up to the current requested scale without losing an explicit user change or reapplying the same ratio. Re-entrant work for a form is ignored while that form is already being scaled.
- Application idle coordination is restricted to the initialising UI thread; actual form updates still marshal to their own UI thread. Already-current forms do not receive redundant callbacks.
- Closing/disposed forms are checked when queued work actually runs, and again after their scale hook, before layout/invalidation.
- The application test version is 2.85. No workbook code, source workbook, user settings value, or master schema was changed by the repair.

DevExpress documents splash windows running on a separate thread: https://docs.devexpress.com/WindowsForms/DevExpress.XtraSplashScreen.SplashScreenManager.Default . This explains a plausible concurrent window lifecycle; the user's stack does not identify the exact form that changed. The regression uses real Windows Forms on two STA threads rather than assuming a particular splash window was responsible.

## Validation

Both Debug and Release build successfully. `Tools/PresentationScaleRaceFixture.cs` passes 16 assertions per build, using `Tools/Test-ClientReport.ps1 -Configuration <Debug|Release> -Fixture PresentationScaleRaceFixture.cs`.

- The fixture reproduces the original enumeration exception deterministically, then verifies retry/discard and bounded deferral.
- Initialisation, repeated idle, explicit scale changes, re-entrant idle, deferred scale catch-up, late forms, disposal, self-closing hooks and restoration to 100% are checked.
- Two STA threads complete 200 actual form open/close cycles without a crash or rescaling stable forms.
- The initial stress run exposed excess cross-thread callback traffic and timed out. Main-thread coordination and current-scale filtering were then added; the final Release and Debug runs passed.
- Final scale evidence: Release `obj/ClientReportTests/8abf95402a1d4b27a82569a5ca364ae8/test.log`; Debug `obj/ClientReportTests/ee27ea4ffad746a1bd9901576ce0282a/test.log`.
- Existing 2.84 client-review regression passes all 35 assertions per 2.85 build on private workbook copies. Debug `obj/ClientReportTests/46a3f6443e74416b931682fe93060f28/test.log`; Release `obj/ClientReportTests/3df76a10727c4a018c4b188a9d80a808/test.log`. Source hashes are unchanged. The pre-existing Chrome_WidgetWin_0 shutdown diagnostic remains outside this repair.

## Manual test and report

Jon should repeat the triggering operation, open and close interfaces/Options while normal progress windows appear, and check 100% -> 125% -> 100% scale without cumulative growth. Native tests are not client DPI/functional acceptance.

The final Word report is `Library/Client_Review_2.85_v04.docx`, delivered as `D:/Downloads/Summit v2.51 comments - review v04.docx`. Blue means Jon functional test; orange issue headings with light-blue details mean committed fixes ready for Alex. Green is reserved for items green in Alex's returned Word document. Both source screenshots are byte-identical and all twelve final rendered pages passed visual inspection. Word remains the review document; no Excel tracker was created.
