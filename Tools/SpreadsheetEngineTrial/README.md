# Isolated spreadsheet-engine trials

## Complete Funding checkpoint (24 September)

**Research only: NOT safe for production handoff.** See
`Library/SpreadsheetGear_Funding_Full_Trial_2026-09-24.md`. The serial port needs
Excel grouped-sheet 3-D semantics; raw Gear saving loses this AGL's custom XML
and dynamic-array metadata. AutoFill does not remedy either gate.

Commands accept only absolute paths below `obj/AsposeTrial`, reject existing
outputs and never modify their input. Input/output workbook format is XLSM.

```text
funding gear-copy|gear-fill|gear-baseline|excel|excel-baseline INPUT NEW_OUTPUT NEW_REPORT
funding dx-reload INPUT NEW_REPORT
funding-3d NEW_REPORT
```

The Funding command inserts ten ordinary-loan columns on all 32 reviewed sheets,
then updates the eleven TDB mirrors. Baseline variants perform no insertion.
Excel modes require the user's explicit isolated VBA permission and respect ByUI
security policy; they call the actual embedded Funding procedures, never the
interactive outer wrapper or workbook-open event. No source VBA is extracted to disk.
`funding-3d` is synthetic, does not open a customer workbook or execute VBA, and
compares serial inserts with Excel's grouped operation. `dx-reload` measures the
native engine load only, not Summit model initialization or UI rebind.

Run independent manifests with `inspect`, then:

```text
python Tools/SpreadsheetEngineTrial/Audit-FundingPackages.py BEFORE AFTER NEW_REPORT --manifests BEFORE_MANIFEST AFTER_MANIFEST
```

The package audit reads XML/VBA hashes, never passwords or extracted VBA source.
Counts or differing fingerprints are diagnostic gates, not automatically defects.
`Compare-Benchmarks.ps1 -AllowRoundtripInputs` is only for deliberately paired
save/reopen tests; ordinary same-input comparisons must retain the hash check.

This standalone .NET Framework 4.8 console project is not part of the Summit solution.
It does not change Summit's package references, shipping version or UI. The DevExpress
real-model comparator loads the existing Release executable only to register its two UDFs;
it does not launch Summit or use its save/transaction/UI services.
Aspose.Cells 26.9.0 is pinned exactly; retain `packages.lock.json` for reproducibility.
SpreadsheetGear 9.3.85 is also pinned exactly (runtime assembly 9.3.85.102).
Packages restore to the repository's ignored `packages/engine-trial` directory.

## SpreadsheetGear activation and smoke tests

SpreadsheetGear is available only in this standalone trial, not the Summit application.
The signed trial licence supplied on 24 September 2026 expires on 24 October 2026.
Never embed its text in source, command-line arguments, reports or exported workbooks.

The harness accepts a process-local `SUMMIT_SPREADSHEETGEAR_SIGNED_LICENSE` value.
`sg-license-install` validates that value and stores it using Windows DPAPI CurrentUser
at `%LOCALAPPDATA%/Abovo/SummitEngineTrials/SpreadsheetGear.lic.dpapi`, outside the
repository. Existing licence files are not overwritten. This encrypted copy is tied to
the Windows account; it is not a portable licence backup.

`sg-smoke` uses the explicit process variable if provided, otherwise the encrypted local
copy. A rejected, expired, missing or undecryptable licence fails the run, with no free-mode
fallback. Licence activation happens before any other SpreadsheetGear Factory call.
Activation exceptions do not print the supplied key or vendor error text.

```powershell
& Tools/SpreadsheetEngineTrial/bin/x64/Release/net48/SpreadsheetEngineTrial.exe sg-smoke
```

Nineteen assertions cover process architecture, cross-sheet/named-range calculation,
incremental edits, row/column insertion and deletion reference fixes, operation beyond
all four free-edition limits, in-memory XLSM save/reopen, and independent DevExpress reads.
No client file or Excel instance is used. This is an installation check, not an Abovo
financial, VBA-preservation, fill-locking, dynamic-array or performance comparison.
The separate `bench` commands now compare native calculation and range operations across
all four engines. Existing `io` commands still accept only Aspose and DevExpress.

The user renewed permission on 24 September for workbook VBA functions in the isolated
Excel benchmark. That permission does not change Windows/Excel macro security or permit
workbook-open events, modifying originals, or running unrelated macros. No Excel/VBA
trial was executed as part of the licence activation alone. Subsequent calculation
benchmarks use a dedicated Excel process, read-only disposable inputs and ByUI macro policy.

## Calculation and range benchmarks

All paths must be inside `obj/AsposeTrial`; outputs must be new. Build x64 Release for
the recorded performance comparison. Set the external Aspose licence path for Aspose;
SpreadsheetGear uses its previously activated, external DPAPI licence.

```text
bench fixture PRIVATE_NEW_SYNTHETIC.xlsm
bench convert PRIVATE_SOURCE.xlsb PRIVATE_NEW_COPY.xlsm PRIVATE_NEW_CONVERSION.json
bench udf-excel PRIVATE_AGL.xlsm PRIVATE_NEW_UDF_REPORT.json
bench devexpress|aspose|spreadsheetgear|excel PRIVATE_INPUT.xlsm PRIVATE_NEW_RESULT.json synthetic|agl
```

`fixture` creates 24,579 known-answer formulas. `synthetic` checks all answers after
three incremental edits and after row/column insert-copy-delete cycles, checks copied
input values while inserted, verifies defined-name restoration, and transfers/checks
200,000 values. Timings exclude validation. Use fresh processes, sequential engines
and rotated engine order; do not benchmark engines concurrently.

`convert` uses Excel with macros/events disabled and no requested calculation, writes a
new XLSM and never saves its source. `agl` uses identical XLSM bytes for all four engines,
calculates five complete worksheet ranges and records their results. Aspose and
SpreadsheetGear have trial-only PMCost/RespCost callbacks; malformed input handling is
not approved as production-equivalent. Excel runs under existing macro security policy,
with events and external-link updates disabled; no security setting is weakened.

```powershell
& Tools/SpreadsheetEngineTrial/Compare-Benchmarks.ps1 -Baseline $excelJson -Candidates $otherJsonFiles -Report $newComparisonJson
```

The comparison normalizes blanks and modern Excel serial dates, with numeric tolerance
`max(0.000001, abs(Excel value)*1e-10)`. The Aspose date export adapter explicitly rejects
1904-date workbooks; the recorded fixtures use the 1900 date system. These are native API
tests, not identical implementation algorithms: Aspose full calculation disables its dirty
chain; copy tiling, bulk transfer and calculation optimizations vary by engine.

Following explicit permission on 24 September, `agl` temporarily unprotects the two
disposable source sheets using the existing Structure.xml field in memory only. It restores
entry protection in `finally`, never writes the credential and never saves the model.
The fixed AGL geometry copies 580 Funding / 688 Development rows into ten inserted columns;
larger geometry is rejected. SpreadsheetGear AutoFill FillCopy is also measured on freshly
reinserted columns, and its formulas/constants must match Copy. After deletion and full
recalculation, the same five output probes must match their pre-insertion values.

Those primitives are NOT full Funding/Development commands, and do not validate all mirror
sheets, VBA behavior, formulas/names outside the probes or a saved round-trip. `udf-excel`
separately invokes embedded PMCost/RespCost with scratch-workbook inputs (expected 80 and 16),
proving permitted VBA execution rather than relying on cached model values. No source cells
are edited. Keep all generated workbooks away from clients.
Results and remaining gates: `Library/Spreadsheet_Engine_Calculation_Range_Benchmark_2026-09-24.md`.

## Build and smoke test

From the repository root in PowerShell:

```powershell
dotnet restore Tools/SpreadsheetEngineTrial/SpreadsheetEngineTrial.csproj --source https://api.nuget.org/v3/index.json --locked-mode
dotnet build Tools/SpreadsheetEngineTrial/SpreadsheetEngineTrial.csproj --no-restore -c Release -p:PlatformTarget=x86
& Tools/SpreadsheetEngineTrial/bin/x86/Release/net48/SpreadsheetEngineTrial.exe
dotnet build Tools/SpreadsheetEngineTrial/SpreadsheetEngineTrial.csproj --no-restore -c Release -p:PlatformTarget=x64
& Tools/SpreadsheetEngineTrial/bin/x64/Release/net48/SpreadsheetEngineTrial.exe
```

With no arguments the executable tests synthetic values and a cross-sheet
formula, recalculates after changing an input, then saves/reopens XLSB and XLSM in memory.
It leaves no workbook files on disk. Thirteen assertions check installation (including
the actual process architecture), not Abovo
model compatibility, custom-function parity, VBA preservation or speed.

## Temporary licence

Request a temporary licence for **Aspose.Cells for .NET**:
https://purchase.aspose.com/temporary-license

Keep the licence outside the repository. Set this process-local variable to its real path
before running the executable (the path below is only a placeholder):

```powershell
$env:SUMMIT_ASPOSE_LICENSE_PATH = 'C:\Licences\Aspose.Cells.lic'
```

The harness reads that exact file through a stream. It never prints the contents, copies
the file or searches other directories for licences. An invalid supplied licence fails
the run; it does not silently fall back to evaluation. Without one, synthetic tests are
allowed and explicitly report `Licensed=False`. Do not use evaluation output as evidence
of a clean workbook round-trip: evaluation adds an extra worksheet and changes active-sheet state.

No purchase, vendor account registration, workbook upload, or production integration is
performed by this project.

## Private-file I/O and preservation trial

The isolated project now also references the installed DevExpress 25.2.4 assemblies from
`bin/Release` as a comparator and independent reader. This does not alter Summit's references.

Real-file commands restrict paths to **`obj/AsposeTrial`** and refuse existing outputs and
reparse-point paths. `io aspose` requires an activated licence. Neither engine is asked to
calculate, edit inputs, run VBA or refresh links. Original files are copied and hashed by
the runner; each engine receives only a private copy.

```powershell
& Tools/SpreadsheetEngineTrial/Test-Io.ps1 -LicencePath 'C:\Licences\Aspose.Cells.lic' -Cases Demo,AGL -Runs 3 -Architecture x64
```

This creates a unique ignored output directory, runs each case in a fresh process, and
alternates engine order. Runs are sequential. The OS cache is not flushed, so results are
**warm-cache**, not cold-open measurements. Both output formats are saved from the same
XLSB source; this is not an XLSM-versus-XLSB input-load comparison. JSON records include
version, bitness, source hash, source-unchanged result, load/save times, peak working set
and output size. Licence environment changes are restored after the script.

Direct executable commands (absolute paths inside that private directory):

```text
io aspose|devexpress INPUT OUTPUT NEW_RESULT.json
inspect INPUT NEW_MANIFEST.json
inspect-core INPUT NEW_MANIFEST.json
compare-styles BEFORE AFTER NEW_REPORT.json
compare-formulas BEFORE AFTER NEW_REPORT.json
```

`inspect` reads through DevExpress without calculation. It fingerprints sheet order,
visibility/protection, all populated-cell formulas and array-range definitions, constant types/values,
cached results, selected formatting/locks, scoped names and selected package parts.
`inspect-core` omits formatting for a faster first pass. `compare-styles` classifies those
formatting differences and counts changes to the Solid-versus-non-Solid fill decision.
Schema 2 fingerprints legacy/dynamic array collections once per range; schema 1 checked
flags per cell and is not directly hash-comparable. This avoids very expensive repeated
array lookups in AGL. `compare-formulas` records difference counts and bounded samples.
These checks do **not** cover all blank-cell styles, borders, validation, conditional rules,
charts, other drawing objects, VBA behaviour or financial calculation equivalence.

`Test-ExcelOpen.ps1 -Files ... -Report ... [-ProbeStyles] [-ProbeArrays]` uses its own Excel instance,
normal read-only opening, disabled macros/events/link updates and manual calculation.
It closes without saving and verifies hashes. The optional five style samples investigate
Demo differences; the three optional AGL array samples check Excel's Formula/Formula2 and
HasArray/HasSpill properties. Neither is an exhaustive visual, conditional-format or array test.

**Raw DevExpress serialization intentionally omits Summit's production CONCATENATE and
recovery-metadata guards.** These figures are not end-to-end Summit timings and raw outputs
are not approved user files. A proposed helper must also pay for staging, reload and rebind.

Keep all generated workbooks in the ignored trial directory. Never give these unvalidated
files to clients. PMCost/RespCost callbacks and structural commands remain separate gates.
See `Library/Spreadsheet_Engine_Trial_Assessment_2026-09-23.md` for the agreed trial scope.

## Persistent worker research (not production integration)

`mirror excel|devexpress model|synthetic PRIVATE_INPUT.xlsm NEW_PRIVATE_DIRECTORY`
runs a persistent Gear working copy and a separate native worker. `model` is restricted
to the reviewed AGL Funding geometry. Excel model mode executes the previously approved
embedded Funding VBA using ByUI policy; synthetic mode disables macros. It never modifies
the input or a live Summit workbook. Do not invoke model mode on an unknown workbook.

`mirror fixture PRIVATE_SYNTHETIC_SOURCE.xlsm NEW_PRIVATE_INPUT.xlsm` makes D1 editable
in a disposable copy of the bundled public support fixture, using macro-disabled Excel.

The journal and save barriers are experimental. Trial custom-XML restoration is valid
only because the prototype never edits XML/history. Native VBA binary regeneration is
recorded for the independent source/form audit, not represented as a round-trip pass.
`Audit-MirrorTrial.py TRIAL_REPORT EXCEL_REFERENCE_JSON NEW_REPORT` compares five probes,
array geometry and in-memory VBA source hashes. It requires oletools outside the repo.

See `Library/Spreadsheet_Two_Engine_Prototype_2026-09-24.md` for phase-specific results,
tested failures, the dynamic-spill limitation and the production gates.

Second checkpoint: `Library/Spreadsheet_Mirror_Projection_Trial_2026-09-24.md`.
All mirror Gear loads now skip object import and cannot serialize a reduced copy.
The complete DevExpress Funding trial runs but fails its independent array/probe
gate; do not confuse `harnessCompleted` with production compatibility. Excel is
optional and separately owned, not required by existing production Summit.

`mirror-array-copy NEW_PRIVATE_DIRECTORY` exercises a native single-row array-copy
guard; `mirror-array-state PRIVATE_XLSM NEW_REPORT` compares TDB package declarations
with native reload interpretation without changing the workbook. The attempted
row-copy guard does not resolve AGL's missing serialized array declarations.

Synthetic mirror runs also test bounded native `read-values` snapshots: full native
calculation, exact session/base/revision/rectangle, rejection of stale reads, and
no formula overwrites. Native calculation exceptions invalidate worker publication.
No automatic production consumer eligibility or engine failover is implemented.

Third checkpoint: `Library/DevExpress_Array_Insert_Diagnosis_2026-09-24.md`.
The incomplete TDB arrays are reproduced with a small native row-insert/copy
fixture, independent of AGL, VBA and intermediate inspection. The isolated
DevExpress Funding worker now uses `MirrorDynamicArrayBatch`: retain expressions,
map reviewed row shifts, then restore single-cell dynamic array registrations
before calculation. Multi-cell spills are refused. The export gate verifies
expected native TDB array declarations in the candidate before publication;
this is not a full-model integrity or financial acceptance gate.

```text
mirror-array-insert NEW_PRIVATE_DIRECTORY
mirror-array-trace PRIVATE_BASELINE.xlsm NEW_PRIVATE_DIRECTORY
mirror-array-trace-fixed PRIVATE_BASELINE.xlsm NEW_PRIVATE_DIRECTORY
mirror-excel-array-open PRIVATE_CANDIDATE.xlsm NEW_REPORT.json
```

The unqualified trace deliberately retains the failing native path and does not
enforce the new export gate, so it can record the bad candidate. The fixed trace
uses the batch correction and export check. All files remain disposable. The
Excel open command owns its process and disables macros/events/link updates;
it neither recalculates nor saves the input. Consult the checkpoint report for
AGL oracle comparisons, timing costs, incomplete runs and remaining limitations.
