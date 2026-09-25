# Stages 2n-2s: native consumers and opening selection

Internal integration checkpoint, not a client test delivery. The main opening path now selects an owner in isolated builds; installed Debug/Release remain the unchanged 2.98 executables. Continue `Excel_Stage2_Work_Plan.md` through the remaining gates.

## Ownership and interface integration

- Opening prefers compatible Excel using the user's existing macro/security policy. Options has a Calculation engine tab, actual per-model status and a DevExpress-only choice. Changes apply to subsequent opens; no mid-session engine replacement occurs. Recovery copies and Stress Test-mode workbooks retain ordinary DevExpress. Automatic Excel unavailability closes the unsuccessful trial before retaining the original ordinary DevExpress model.
- FFR views, dashboards, mapped/ordinary DIT values, dropdowns, Funding captions/tints and model comparison read the selected owner. FFR edits/paste use existing typed change-manager batches, not a second journal; return-template output copies current values into a separate workbook. Existing unlocked-plus-solid-fill FFR permission is preserved.
- Explicit summary refresh calculates the selected owner. Integrity cell scanning reads native values rather than the original display workbook's cached results. Names/formula geometry still comes from the original structural map because structural changes are prohibited in this stage. The report describes existing mapped cells; variable-size spill tails outside that map remain a qualification limit, not a claim of exhaustive native error coverage.
- Legacy direct spreadsheet surfaces, old Stock/Transaction interfaces, direct value/range helpers, imports, schedule metadata, upgrades and shared structural commands are refused before mutating the presentation workbook. Native single-cell and batch editing remains available after refusal. Ordinary unbound DevExpress remains unchanged.
- Native Save As supports a new filename in the same format. Arbitrary existing-destination overwrite and format conversion are not enabled. Recovered XLSM files deliberately open through DevExpress to support the existing guided XLSB Save As workflow.

## Completed evidence

| Fixture | Result | Scope |
| --- | --- | --- |
| Ordinary DIT, production-matching LAA runner | 550 checks | Existing DevExpress input/navigation/formatting/save regression |
| Native FFR | 61 per configuration | Generated workbook, nine view/return paths, edit/paste/Undo and protection |
| Native dashboards | 19 per configuration | Native selector edit, chart/HTML refresh, Undo/Redo |
| Expanded structural guards | 51 | Both native owners, refusal before mutation, normal editing afterwards |
| Opening/settings | 19 per configuration | Excel, unavailable Excel, recovery, Stress Test and unchanged live selection |
| Current consumer checks | 32 | Both engines/date systems; stale presentation caches, formula error/Undo, comparison and typed dates |
| Existing bridge / safety / batch | 133 / 73 / 48 | Selected-owner transaction regression |
| Presentation / security / deadline | 33 / 12 / 6 | Effective conditional appearance, security markers, bounded waits/quarantine |
| Whole private Demo opening lifecycle (2q Release) | 35 | Actual opening-selection entry, Covenant/Funding DIT edits, recovery, Save, Save As, Undo/Redo, independent saved-file reads and owned close |

Native fixtures use generated files or private copies. Source hash checks pass; only tracked Excel owners are closed. Existing user Excel workbooks are not touched. Logs are beneath ignored `obj/EngineStage2*-*.log`. `Tools/WorkbookEngineFoundationTests` contains the reproducible source; licence/customer-workbook payloads are not in the checkpoint.

An independent ordinary-DIT reopen initially returned false. A fresh-process read of exactly those bytes succeeded; enabling DevExpress `ThrowExceptionOnInvalidDocument` in the fixture exposed `OutOfMemoryException` at approximately 1.13 GB private bytes in the non-LAA x86 test process. `Run-NativeRegression.ps1 -MatchApplicationAddressSpace -UseApplicationConfig` matches Summit's LAA executable and passes all 550 checks. No production save change was made to conceal that failure.

## Current performance experiment and limits

The 2q whole-model run bound Funding in 42.2 seconds. Narrowing effective-format prefetch from eight rows/four columns to eight rows/one column measured 39.5 seconds in the first 2s run, but Covenant also varied and this is not a controlled median or a proved improvement. Native fills and protection still come from Excel DisplayFormat; none are inferred from colours or stale DevExpress caches. The 2s full lifecycle subsequently passed all 36 checks, including explicit summary refresh. The final Version 3 Release preview also passed 36 checks and bound Funding in 35.2 seconds; its DevExpress-owned private Demo lifecycle passed 35 checks. These single runs do not establish a performance improvement.

Remaining: native UI/read coverage and failure UX review; appearance transport/save latency; final Debug/Release regression matrix and separately versioned delivery; Jon/Alex functional and Excel/VBA round-trip acceptance. Permanently hung COM has bounded caller/quarantine behavior, not forced cancellation. No blanket Excel termination is permitted. Stage 3 inserts/deletes/schedules/imports remain separately gated. This checkpoint does not complete Stage 2 or approve customer deployment.

## Subsequent Version 3 preview qualification

The historical checkpoint status above is superseded by `Version3_Stage2_Acceptance_2026-09-25.md` and the current Stage 2 work-plan checklist. Final selection tests now include simultaneous opening/cleanup failure (21 assertions); both errors remain available and no owner is bound. The consumer fixture now exercises a real mapped Check Sheet grid and Yes/No override editor (68 assertions), showing break/Undo/Redo and company-warning coherence against deliberately stale presentation caches in both engines and both date systems. Client 2.98 remains a separate pre-engine build; no Alex or financial sign-off is inferred.
