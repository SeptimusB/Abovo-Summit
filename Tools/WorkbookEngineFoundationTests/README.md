# Workbook engine foundation tests

Standalone .NET Framework 4.8 harness for the production-source engine boundary. It does not enable Excel in the live Summit application. Sessions are read-only by default; isolated value-edit tests explicitly opt in to in-memory changes with compensation. No save API is exposed. Use private copied workbooks, never a user's active editing file: the tested source is leased read-only for the session.

Build the main VB project using VS2022 MSBuild with `/p:Configuration=Debug /p:OutputPath=bin/EngineStage2b-Debug/`, and similarly Release to `bin/EngineStage2b-Release/`. Then `dotnet build Tools/WorkbookEngineFoundationTests/WorkbookEngineFoundationTests.csproj -c Debug`. Default test process is x86, matching current Summit; override PlatformTarget and use a separate OutputPath for x64 tests.

Run from the repository root:

```powershell
& 'Tools/WorkbookEngineFoundationTests/bin/Debug/net48/WorkbookEngineFoundationTests.exe' 'C:/Repos/Abovo Summit/bin/EngineStage2b-Debug' 'ABSOLUTE_PRIVATE_WORKBOOK_PATH.xlsm' safety
```

Modes:

- `safety`: fake-backend ownership, queue, revision, cancellation and failure tests. No Excel opens.
- `security`: private GUID-named copies beside the supplied synthetic XLSM; exercises security-marker and package screening. Deletes only those owned files. No Excel opens.
- `deadline`: held fake opening/calculation with bounded waits, quarantine and eventual owner-thread cleanup. No real Excel hangs or process termination are induced.
- `grid`: controlled accepted results bound to a real native DevExpress GridView, including typed values, readonly/stale safeguards, batched refresh and multiselect.
- `projection`: isolated unsaved synthetic DevExpress workbook reproduces scalar callback behavior and the array limitations that disqualify it as a general display adapter. This is not an accepted production route.
- `edit`: controlled-backend typed input, snapshot/permission checks, compensation, concurrent requests, cancellation and held-native timeout cases. No native workbook or Excel opens.
- `edit-native`: generates a small GUID-folder native XLSX fixture beside the private input and tests in-memory edits with both engines. Literal text, negative fractions, serial numbers, booleans, blank, formula/array/merge guards, lock/fill rules and worksheet protection are checked. Native sessions never save the edited fixture. Fixture stays local as evidence.
- `edit-agl`: opens a reviewed private AGL XLSB/XLSM read-only with each engine, edits Funding Assumptions G82 from its baseline by +300 in memory, performs full calculation/read-back, compares outputs, restores the old value and compares again. Source bytes and owned Excel cleanup are checked. This is not Summit history/UI or save/round-trip acceptance.
- `synthetic`: fixture previously created by `Tools/SpreadsheetEngineTrial`, with `Data` 4096x10 and `Summary!B1:B3`. Compares native engines including blanks, then checks cleanup and required-VBA-disabled fallback.
- `agl`: reviewed private AGL XLSB or converted XLSM. Compares five established bounded financial/check rectangles; explicitly permits required VBA under existing Excel policy. It opens an owned separate Excel process with events and link updates disabled and closes without saving. Does not certify the whole workbook or structural/save behavior.

Native tests require installed compatible desktop Excel and the licensed production DevExpress assemblies. No Office PIA, optional spreadsheet package, customer workbook, licence or result payload is included here. Fake tests need the main assembly/dependencies but do not open a native workbook. Any exception or mismatch returns a nonzero exit code. Original source SHA256 must be unchanged. The test never kills pre-existing Excel processes.

Evidence/limitations: `Library/Excel_DevExpress_Engine_Stage1_2026-09-25.md`, `Library/Excel_DevExpress_Engine_Stage2a_2026-09-25.md`, `Library/Excel_DevExpress_Engine_Stage2b_2026-09-25.md`. The native modes also bind accepted engine results to real grids. Wait deadlines and isolated compensated value edits are implemented; process-level recovery for calls that never return and live edit/save integration remain release gates.
