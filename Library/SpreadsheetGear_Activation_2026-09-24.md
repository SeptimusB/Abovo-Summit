# SpreadsheetGear isolated trial activation

Date: 24 September 2026.

## Scope and licence

- User supplied a signed trial licence and authorised activation. It expires on
  24 October 2026. No licence contents are stored in repository files or test output.
- NuGet `SpreadsheetGear` pinned to 9.3.85 in the existing standalone
  `Tools/SpreadsheetEngineTrial` project. Runtime assembly reports 9.3.85.102.
- Targets .NET Framework 4.8 through the package's .NET Standard 2.0 assembly.
- Licence accepted by the vendor API and stored outside the repository using Windows
  DPAPI CurrentUser. Future trial processes can activate without a plaintext file.
- No Summit production reference, UI, engine selection, version or workbook changed.
  No purchase, workbook upload or vendor account action performed.

## Verified results

The isolated project built without warnings or errors in Debug/Release for both x86 and
x64. Each of the four processes passed all 19 assertions:

- Actual process architecture.
- Cross-sheet calculation using a defined name; incremental calculation after an edit.
- Formula/name adjustment on row and column insertion; restoration after deletion.
- Access beyond free-mode limits of 1,000 rows, 100 columns, 10 sheets and 3 workbooks.
- Synthetic XLSM serialization and reopen/recalculation entirely in memory.
- Date number-format preservation and an independent DevExpress formula/date read.

The pre-existing licensed Aspose smoke test also passed its 13 assertions in x64 Release.

These are installation tests, not an Abovo compatibility or speed result. They do not
test actual VBA preservation/execution, custom XML, full financial calculations,
PMCost/RespCost parity, dynamic arrays, blank-cell fills, fill-based editability,
conditional formatting or complete structural commands. No client workbook was opened.

## Next evaluation boundary

The user accepts considering client XLSM conversion, removing the native XLSB format
barrier for this candidate. Do not mass-convert or replace originals on that basis.
Use identical disposable XLSM inputs for the four-engine comparison and account for
conversion/load/save overhead separately. Production DevExpress remains unchanged.

The earlier answer allowing Excel VBA was revoked, then explicitly renewed on
24 September: "Also allow VBA in the isolated Excel benchmark". Workbook-open events
must remain disabled, originals unchanged, and existing security policies respected.
No Excel automation or VBA execution occurred during this activation.

At this activation checkpoint, larger calculation/range benchmarks and model adapters
remained pending. The later isolated results are now recorded in
`Library/Spreadsheet_Engine_Calculation_Range_Benchmark_2026-09-24.md`. Do not count the
small installation smoke tests as Funding/Development validation.

## Vendor references

- [NuGet setup and signed trial licence](https://www.spreadsheetgear.com/nuget/spreadsheetgear/project/)
- [Release notes](https://www.spreadsheetgear.com/downloads/whats-new): 9.3.85 includes an
  array-formula fix relevant to repeated copy/insert/delete operations.
- [Current limitations](https://www.spreadsheetgear.com/support/help/spreadsheetgear.net.9.0/SpreadsheetGear_2023_Limitations.html)

The separate API tests used the XML documentation shipped with package 9.3.85.
