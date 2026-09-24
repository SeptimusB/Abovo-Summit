# Spread.NET primary-engine evaluation

Isolated research only. Not referenced by the Summit solution/project and not a production adapter.
Decision and evidence: `Library/SpreadNET_Primary_Engine_Evaluation_2026-09-24.md`.

Pinned packages: `Mescius.Spreadsheet.Core` and `GrapeCity.Spread.WinForms` 19.2.0,
runtime 19.2.20266.0; .NET Framework 4.8, x64. Restore goes to ignored `packages/spread-trial`.
Aspose/Gear/DevExpress references only reuse the existing independent comparison harness.
Neither is a worker or export-repair engine in the Spread adapter.

The default native host is an undisplayed `FpSpread(LegacyBehaviors.None)` in the same STA
process. Calculation is manual, background calculation off, on-demand calculation off,
dynamic arrays enabled. PMCost/RespCost must be registered **before** loading the model.
Use documented `OpenExcel`/`SaveExcel` with DocumentCaching and XLSM flags, and a read/write
output stream. The earlier write-only stream produced a false Save return and empty output;
that was a harness mistake, not a vendor defect. All native boolean returns are checked.

The signed trial activates through the normal package build tooling. Windows control creation
can show a trial reminder; do not suppress/bypass it. Engine construction is outside timed
load/calculation/range operations. Reject any IO sample interrupted by a later prompt.
No licence files, keys, extracted VBA or financial workbooks belong in this directory.

From the repository root:

```powershell
& 'C:/Program Files/dotnet/dotnet.exe' restore Tools/SpreadNetTrial/SpreadNetTrial.csproj --source https://api.nuget.org/v3/index.json
& 'C:/Program Files/dotnet/dotnet.exe' build Tools/SpreadNetTrial/SpreadNetTrial.csproj --no-restore -c Release
```

Executable: `Tools/SpreadNetTrial/bin/x64/Release/net48/SpreadNetTrial.exe`.
Every input/output must be under `obj/AsposeTrial`; reparse paths and overwriting outputs
are rejected. Create each output directory first. Commands:

- `smoke`: fresh dynamic spill/growth and native UDF known answers (80 and 16).
- `bench spreadnet INPUT NEW_REPORT synthetic|agl`: common calculation/range benchmark.
- `batch INPUT_DIRECTORY NEW_OUTPUT_DIRECTORY`: three synthetic runs then AGL in one process.
- `diagnose INPUT NEW_REPORT`: one AGL full rebuild, custom-function counters and output probes.
- `cases NEW_DIRECTORY`: tiny grouped-insert and scalar/dynamic-array insert/save/reopen repros.
- `io INPUT NEW_OUTPUT NEW_REPORT lossless`: native save/reopen, no requested calculation/VBA.
- `verify-excel CASE_DIRECTORY NEW_REPORT`: independent macro-disabled Excel verification;
  uses only new owned Excel processes, never saves the source.
- `excel-open INPUT NEW_REPORT`: macro-disabled, read-only Excel open, no requested calculation.
- `inspect INPUT NEW_REPORT`: independent DevExpress manifest, no calculation/save.
- `edit ENGINE INPUT NEW_REPORT`: implemented edit/restore experiment; not an accepted Spread
  comparison in this checkpoint because baseline compatibility already fails.

`SPREAD_TRIAL_HEADLESS=1` selects the Core factory experiment. Simple synthetic timings pass,
but its UDF registration is unresolved; do not use it to claim real-model parity.
`SPREAD_TRIAL_OPTIMAL=1` enables the optional optimal-calculation flag. The isolated insert
failures also reproduce without it. Neither environment switch changes Summit.

Python readers: `Audit-SpreadResults.py` compares measured output probes to the established
Excel oracle; `Inspect-SpreadFormulas.py` reads small cell-formula samples (never VBA).
The existing `Tools/SpreadsheetEngineTrial/Audit-FundingPackages.py` compares package parts.

Current result: basic speed is promising, but real-model calculation and native round-trip
gates fail. No production migration, client workbook conversion or full Funding/Development
command performance claim is approved. Early v1-v4 setup pilots are not accepted evidence.
