# Version 3 preview: Stage 2 acceptance - Ready to test

For Jon's functional testing first. Not an Alex/client release and not financial sign-off. Keep the pre-engine 2.98 client line separate. This integration already descends from the 2.98 source checkpoint; later client-line repairs must be carried across deliberately, not by replacing executables in a mixed folder.

## Which executable

- Preview: `C:\Repos\Abovo Summit\bin\Version3-Preview\Abovo-summit.exe` (Release, version 3.00).
- Preview Debug: `C:\Repos\Abovo Summit\bin\Version3-Preview-Debug\Abovo-summit.exe`.
- Existing client-line binaries: `bin\Release` and `bin\Debug`, still 2.98. They have not been overwritten by this work.
- The separate client task has also produced `bin\Client2.98-Release\Abovo-summit.exe` from that pre-engine checkpoint. Its checked SHA256 is `B863AD15A2AEC167C7FF042979AFBB43CF8B72C3868D1DA6E2F33DABD74FE7F3`; its own handover records the client packaging/QA. Do not copy individual preview files into that folder.
- Pre-engine source checkpoint: local `f6b0ec8`; sanitized code-only equivalent `d01d11597a0b1f8516609f22d9458588fbf6c099`. Rebuild 2.98 from that checkpoint in a separate checkout, **not from the current engine-development source**.

Use disposable copies of real models. Start with the Release preview, open one copy, and check Options > Calculation engine to see the owner actually selected. The preference applies at the next opening, not to a model already open. Selecting DevExpress retains the ordinary no-Excel route.

Preview executable SHA256: Release `93F79A574A57D8E06F760CF58B53E879809759A7DD9E427B2B2ED724FAFD4869`; Debug `1796D1F8038562465A174CC5B8AFC17258378D79716C8313FDE7BCD07B006180`. The production 2.98 binaries and authoritative Blank/Demo masters are separately checked for unchanged bytes.

## Boundaries to know before testing

- Excel is preferred only when installed, compatible and permitted by the existing security policy to use the workbook's VBA functions. Summit never changes Trust Center or force-enables VBA. Failure after editing does not silently switch engines.
- Stage 2 supports value inputs, calculation/read-back, existing history/Undo/Redo, same-format saves and recovery. Native structural commands, schedules, imports, snapshots, upgrades, Stress Test and direct spreadsheet editing are guarded. Reopen with DevExpress for those operations until Stage 3 is qualified.
- Native Save As currently requires a new filename in the same format. A recovered XLSM opens through ordinary DevExpress so its existing guided XLSB Save As remains available.
- This is not yet a speed-qualified release: initial Funding binding measured roughly 35-42 seconds on this development machine; effective formatting and repeated view refresh are still costly. Native Save verifies and reopens the owner, so calculation timings alone do not represent save or end-to-end editing speed.
- Successful native saves retain transaction/candidate material beside the workbook. Failed saves retain verified copies and report their paths; do not delete these while investigating. Retention/cleanup and interrupted-save recovery UX need further production hardening. A permanently hung COM call is quarantined with bounded caller waits; it is not forcibly interrupted.
- Integrity scans current native values at existing mapped cells. Variable-size spill tails beyond the opening map and arbitrary XLSB dynamic-array growth in DevExpress remain specific qualification limits. No general Excel/VBA or financial certification is implied.

## Numbered functional checks

All start **Jon to test**. Report results as `V3-01 pass`, `V3-02 fail ...`, etc. Passing here does not automatically mark an item green or send it to Alex. Only Jon-approved items move to a subsequent Alex Test Response.

| ID | Action and expected result | Status |
| --- | --- | --- |
| V3-01 | Open a copied AGL model with the automatic preference. Options and System Messages identify the actual Excel owner, or give the fallback reason. Opening alone does not create input history. | Jon to test |
| V3-02 | Choose DevExpress, close/reopen a copy, and check ordinary editing and calculation. Re-enable automatic preference for later tests. | Jon to test |
| V3-03 | Edit Stock, Rent and Covenant assumptions; verify typed numbers, decimals, negative percentages, dates, dropdowns and calculated figures. Existing workbook fill-based locks must still agree. | Jon to test |
| V3-04 | Change a Funding defining date, then enter its amount. Permitted cells unlock and disallowed cells remain locked. Test copy/paste/clear, keyboard navigation and double-click copy. | Jon to test |
| V3-05 | Break a Funding opening balance, keep Check Sheet visible, then Undo/Redo. Grid figures/colours and company warning update at the same accepted result without navigating away. Overrides remain Yes/No and an accepted override clears only the soft balance warning. | Jon to test |
| V3-06 | Open SOCI/Cashflow/Balance Sheet analysis, drill down and switch chart/grid. Compare displayed figures with the copied model in Excel after saving. | Jon to test |
| V3-07 | Change Funding Dashboard selections; check charts/headings, BP Dashboard, explicit sidebar refresh and unsaved model comparisons. | Jon to test |
| V3-08 | Exercise FFR inputs, paste and Undo. Confirm current calculated views and an exported return template, including its dates and totals. | Jon to test |
| V3-09 | Edit > Save > Undo > Save As to a new same-format filename > Redo. Check dirty/Save state, history, original versus new file, company/file summaries and independent Excel/VBA reopen. | Jon to test |
| V3-10 | Make an unsaved input and let recovery save. Confirm notification/snooze/Esc behavior, current history and dirty state. Open the newer recovery, then use the guided XLSB Save As through DevExpress. | Jon to test |
| V3-11 | Edit and close without saving. Reopen: discarded inputs and temporary Check Sheet warnings must not persist as saved work. Close one Summit model while keeping unrelated Excel workbooks open. | Jon to test |
| V3-12 | Try a guarded structural/schedule/import action under Excel; it must explain the limitation without modifying the model. Save, reopen with DevExpress and confirm the existing action remains available on the copy. | Jon to test |
| V3-13 | Time cold open, first Funding view, representative edits, analysis refresh, Save and recovery on the same copy in both routes. Record whole-action time, not just calculation time. | Jon to test |

## Automated qualification

Both candidate application configurations build. The current preview's native consumer fixture passes 68 checks, including a real mapped Yes/No editor and live Check Sheet updates in both date systems/engines. Selection/failed-opening cleanup passes 21. Regression fixtures separately cover bridge/history (133), FFR (61), dashboards (19), guarded structural entry (51), ranges (32), scalar expressions (32), ongoing saves (80), recovery snapshots (162), interrupted publication/recovery (109), permission/failure safety (73), atomic batches (48), presentation (33), security (12), deadlines (6) and foreign-workbook ownership (16). Counts are scoped assertions, not a count of independent financial scenarios.

On the final preview binaries, the whole private Demo Excel opening lifecycle passes 36 assertions, and the native DevExpress lifecycle passes 35. Both exercise actual editors, current values, history, recovery, normal Save, new-name Save As, Undo/Redo, independent saved-file reads and owned close. The ordinary production-path DIT regression passes all 550 checks in each configuration with matching application configuration and large-address-aware runner. Both ordinary fixtures exit successfully; Chromium logs a window-class unregister warning during teardown, retained in the logs rather than presented as a failed assertion or a proved application fix.

The first final Release dashboard run passed all 19 value/UI assertions but exceeded its ten-second Excel-process shutdown check. The tracked process subsequently exited without intervention. The fixture now logs PID/start identity and allows a bounded 30-second shutdown grace period; three isolated repeats passed, with exit waits of 3,059, 3,217 and 3,202 ms. This is an observed intermittent shutdown-timing caveat, not evidence of a repaired COM leak. Production shutdown behavior was not changed, no process was killed, and user Excel instances remained untouched. Keep the initial failing log as well as the repeats (`obj/Version3-Preview-Release-dashboard-native.log` and `dashboard-cleanup-1/2/3.log`). Microsoft documents that Office automation lifetime depends on releasing its COM references; that remains a review consideration, not an established cause for this run: https://support.microsoft.com/en-us/servicing/visual-studio/troubleshooting/office-application-does-not-exit-after-automation-from-visual-studio-net-client

Final logs use `obj/Version3-Preview-*.log`; source/master and existing production-binary hashes remain unchanged. The native regression source is under `Tools/WorkbookEngineFoundationTests`; ordinary UI regression uses `Tools/Run-NativeRegression.ps1` with `-UseApplicationConfig -MatchApplicationAddressSpace`. Build output, generated workbooks and logs are local evidence, not published customer attachments.

The remaining final Release checks also pass: structural admission 51, range binding 32, scalar expressions 32, presentation 33, atomic batches 48, safety 73 and ownership boundaries 16. The first range run passed its 32 data checks but failed an obsolete global Excel-process inventory equality assertion when unrelated process 215960 disappeared. Range fixtures now verify PID/start identity for their own sessions only, matching the other updated fixtures; the repeat passed and its tracked owner exited in 3,745 ms. This test correction does not change application behavior or authorize closing an unrelated workbook. The original failing log remains available alongside `obj/Version3-Preview-Release-final-range.log`.

## After acceptance

Keep client testing on the 2.98 line until Jon approves a Version 3 hand-off. Stage 3 remains the approved three-route Funding/Development structural comparison: existing DevExpress .NET, VB.NET driving Excel, and the matching master VBA. It is not enabled by passing Stage 2 value-edit tests.
