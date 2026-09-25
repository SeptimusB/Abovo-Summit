# Workbook engine foundation tests

Standalone .NET Framework 4.8 harness for the production-source engine boundary. It does not enable Excel in the live Summit application. Sessions are read-only by default; isolated value-edit tests explicitly opt in to in-memory changes with compensation. Candidate exports are a separate opt-in and never replace the source or mark it saved. Use private copied workbooks, never a user's active editing file: the tested source is leased read-only for the session.

Build the main VB project using VS2022 MSBuild with `/p:Configuration=Debug /p:OutputPath=bin/EngineStage2g-Debug/`, and similarly Release to `bin/EngineStage2g-Release/`. Then `dotnet build Tools/WorkbookEngineFoundationTests/WorkbookEngineFoundationTests.csproj -c Debug`. For Release, also set `/p:ApplicationBin="C:/Repos/Abovo Summit/bin/EngineStage2g-Release"`. Default test process is x86, matching current Summit; override PlatformTarget and use a separate OutputPath for x64 tests.

Run from the repository root:

```powershell
& 'Tools/WorkbookEngineFoundationTests/bin/Debug/net48/WorkbookEngineFoundationTests.exe' 'C:/Repos/Abovo Summit/bin/EngineStage2g-Debug' 'ABSOLUTE_PRIVATE_WORKBOOK_PATH.xlsm' safety
```

Modes:

- `batch`: atomic multi-value preflight, prerequisite/final calculation, explicit skips, compensation and quarantine; includes the native DevExpress conditional-fill semantic probe.
- `bridge`: generated real native models, internal existing-change-manager binding, dirty/journal/Undo/Redo, facade reads without source-cell mutation, Check Sheet/company warning, defining-date amount locks, both engines and date systems. Not full live DIT integration.
- `presentation`: generated native workbook formatting, conditional fills/fonts, dates, errors, current calculation-generation checks and native cleanup. Source workbooks never modified.

- `history`: real STA-owned change-manager typed edits/dirty/history, immutable revision-bound snapshots, explicit candidate/publication receipts, stale/mismatched model/source/engine checks, bounded read-only recovery history, malformed-history rejection and failure cleanup. Controlled engine backend, real private package/file operations. Passing history evidence deliberately does not clear live dirty state.
- `history-native`: generated XLSM/XLSB fixtures, both engines, two successive value edits and history saves, receipt-based same-engine reopen and full-calculation comparison in the other engine. Verifies one current history part and the actual edited-cell address.
- `history-agl`: reviewed private AGL baseline, Funding G82 +100 twice, current history, new-name publication/reopen, and four other-engine comparisons over 45,586 output positions. Minimal real change-manager fixture is not live DIT integration. Original supplied workbook is never saved. Run native modes serially.

- `recovery`: bounded journal parsing, state classification, vacant-name original restoration, stale/collision/marker/cancellation failures, receipt-based reopen identity/policy and failed reopen. Includes three hidden owned child test processes which exit abruptly at publication boundaries, using a controlled backend and disposable files, not Excel. Internal `crash-*` modes are child-fixture entry points only.
- `reopen-native`: generated XLSX/XLSM/XLSB in both engines, terminal publication, verified same-engine reopen, another edit/save and other-engine full-calculation comparison.
- `reopen-agl`: new private copy per engine, G82 +300, publication/reopen, G82 +100, new-name publication and other-engine 45,586-position full-calculation comparison. Supplied baseline is never saved. Run native modes serially.
- `publish`: default-off terminal-publication tests on generated copies; revision/collision/failure/cancellation/cleanup/security-marker/ACL checks. Uses real local NTFS file operations but a controlled native backend. Original input is never replaced.
- `publish-native`: both engines, all three formats, new-name Save As and replacement of additional disposable copies; other-engine full-calculation reopen and owned-process checks.
- `publish-agl`: additional AGL copy per engine, G82 +300, source-only replacement of that disposable copy, retained original and 45,586-position other-engine full-calculation comparison. Never substitutes the supplied baseline for the disposable publication target. Native modes must run serially.
- `safety`: fake-backend ownership, queue, revision, cancellation and failure tests. No Excel opens.
- `security`: private GUID-named copies beside the supplied synthetic XLSM; exercises security-marker and package screening. Deletes only those owned files. No Excel opens.
- `deadline`: held fake opening/calculation with bounded waits, quarantine and eventual owner-thread cleanup. No real Excel hangs or process termination are induced.
- `grid`: controlled accepted results bound to a real native DevExpress GridView, including typed values, readonly/stale safeguards, batched refresh and multiselect.
- `projection`: isolated unsaved synthetic DevExpress workbook reproduces scalar callback behavior and the array limitations that disqualify it as a general display adapter. This is not an accepted production route.
- `edit`: controlled-backend typed input, snapshot/permission checks, compensation, concurrent requests, cancellation and held-native timeout cases. No native workbook or Excel opens.
- `edit-native`: generates a small GUID-folder native XLSX fixture beside the private input and tests in-memory edits with both engines. Literal text, negative fractions, serial numbers, booleans, blank, formula/array/merge guards, lock/fill rules and worksheet protection are checked. Native sessions never save the edited fixture. Fixture stays local as evidence.
- `edit-agl`: opens a reviewed private AGL XLSB/XLSM read-only with each engine, edits Funding Assumptions G82 from its baseline by +300 in memory, performs full calculation/read-back, compares outputs, restores the old value and compares again. Source bytes and owned Excel cleanup are checked. This is not Summit history/UI or save/round-trip acceptance.
- `save`: controlled candidate export/reopen failures, cancellation, stale revisions, file/hash tampering, concurrent edits and held-call deadlines; small native XML fixture checks payload/relationship/metadata/security-marker preservation. No Excel opens.
- `save-native`: generated XLSX/XLSM/XLSB fixtures with custom XML, a named input/formula and a dynamic spill. Both engines create candidates, reopen them in the other engine and compare full-calculation results. Candidates remain local as test evidence; sources are not overwritten.
- `save-agl`: approved private AGL input only. Funding G82 +300, native candidate verification, then other-engine full calculation over the established 45,586 output positions. Required VBA obeys existing policy. This is bounded value-only round-trip validation, not structural, live Save/history or workbook-event acceptance. Run native modes serially so owned-process cleanup checks are meaningful.
- `synthetic`: fixture previously created by `Tools/SpreadsheetEngineTrial`, with `Data` 4096x10 and `Summary!B1:B3`. Compares native engines including blanks, then checks cleanup and required-VBA-disabled fallback.
- `agl`: reviewed private AGL XLSB or converted XLSM. Compares five established bounded financial/check rectangles; explicitly permits required VBA under existing Excel policy. It opens an owned separate Excel process with events and link updates disabled and closes without saving. Does not certify the whole workbook or structural/save behavior.

Native tests require installed compatible desktop Excel and the licensed production DevExpress assemblies. No Office PIA, optional spreadsheet package, customer workbook, licence or result payload is included here. Fake tests need the main assembly/dependencies but do not open a native workbook. Any exception or mismatch returns a nonzero exit code. Original source SHA256 must be unchanged. The test never kills pre-existing Excel processes.

Evidence/limitations: `Library/Excel_DevExpress_Engine_Stage1_2026-09-25.md` and continuations `Stage2a` through `Stage2g` under the same Library naming/date convention. Earlier native comparison modes also bind results to real grids. Candidate receipts alone are not publication transactions. Publication is a local NTFS terminal hand-off with a brief missing-name window between guarded renames, not an uninterrupted atomic swap. Explicit journal recovery, verified reopen, current display-history XML and an internal real change-manager/appearance bridge are available; complete live editor/read routing, model opening/settings, frequent Save/dirty lifecycle, native hang supervision and structural/XML coupling remain release gates. Retained generated folders are evidence, not production backup-retention behavior.
