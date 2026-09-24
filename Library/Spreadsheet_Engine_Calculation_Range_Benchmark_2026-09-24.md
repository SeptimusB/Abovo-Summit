# Spreadsheet engines: calculation and range benchmark

Date: 24 September 2026. Isolated research checkpoint; not a Summit test release.

## Decision and scope

SpreadsheetGear merits the next isolated, end-to-end trial. It combines fast calculation,
bulk transfer and source-sheet range manipulation on this AGL case. Excel remains a strong
calculation candidate where installed and compatible. Neither result authorizes replacing
DevExpress, introducing a second authoritative workbook or shipping a converted client file.

Production Summit, its version and its project references were not changed by this trial.
Existing unrelated working-tree changes were preserved. No original or master workbook was
saved, no global Excel security setting was changed, and no workbook was uploaded.

The four adapters are in `Tools/SpreadsheetEngineTrial`, a standalone .NET Framework 4.8
console project. All performance runs use x64 Release, sequential processes and the same
machine: Intel Core Ultra 9 285K, 24 cores/logical processors. Installed Excel is x64.
These are not minimum-spec client-machine results or cold-cache open measurements.

Versions: DevExpress 25.2.4; Aspose.Cells 26.9.0; SpreadsheetGear package 9.3.85, assembly
9.3.85.102; Excel 16.0 build 20326. Trial licences are external to the repository.

## Input and reproducibility

Original, unchanged:

`C:/Sandbox/Insert Comp/AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb`

Original SHA256:
`30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47`

Excel converted a private copy to a new XLSM with macros/events/link updates disabled and
no requested calculation. Conversion open/save were 3.79/6.19 seconds in that single pass,
not a comparative format benchmark. All four calculation engines read the identical XLSM:

`62B44DE2764C4D1D1C9C7029C9297EEF4D461E300BAB4AD7DB092BBE1D23193C`

Raw evidence directory (ignored, disposable; not for client use):

`C:/Repos/Abovo Summit/obj/AsposeTrial/a79b3d867f81456a9d22887b9600ee61`

Final evidence: `synthetic-ENGINE-4.json` through `-6.json`, `AGL-verified-ENGINE.json`,
`comparison-verified.json`, `summary-final.json` (synthetic medians and earlier AGL pass),
`udf-excel-known.json`, and `conversion.json`. Preliminary/pilot/diagnostic reports are not
the accepted comparison. See the trial README for command syntax and safe path restrictions.

## Calculation and structural results

### AGL calculation, seconds

| Engine | Load XLSM | Initial calculation/rebuild | Warm full median | Post-delete rebuild |
| --- | ---: | ---: | ---: | ---: |
| DevExpress | 27.152 | 6.318 | 5.989 | 6.121 |
| Aspose | 4.233 | 4.120 | 3.655 | 3.913 |
| SpreadsheetGear | 3.968 | 0.694 | 0.544 | 0.614 |
| Excel | 12.849 | 4.041 | 0.392 | 4.024 |

SpreadsheetGear's warm median is about 11 times faster than DevExpress here; Excel is
about 15 times faster. The very different initial-rebuild and warm figures matter:
do not apply the warm ratio to workbook opening or structural rebuilds. Startup,
licence activation and process construction are outside the load timer.

All three candidate engines match Excel in the five probes: zero numeric, text or
type/blank differences outside tolerance and zero error values in the compared positions.
All four also pass the post-delete output restoration check.

### AGL source-sheet primitives, milliseconds

These are one-pass native operations; protection, validation and recalculation are outside
the individual operation timers. Copy is the common operation used in the four-way comparison.

| Engine | Sheet | Insert 10 columns | Copy into 10 | Delete 10 |
| --- | --- | ---: | ---: | ---: |
| DevExpress | Funding | 3,166.5 | 159.9 | 3,232.2 |
| Aspose | Funding | 128.5 | 52.9 | 83.1 |
| SpreadsheetGear | Funding | 126.0 | 24.5 | 52.8 |
| Excel | Funding | 1,879.2 | 167.3 | 1,768.2 |
| DevExpress | Development | 12,612.1 | 62.3 | 27,666.0 |
| Aspose | Development | 77.8 | 13.5 | 68.0 |
| SpreadsheetGear | Development | 52.3 | 2.2 | 47.1 |
| Excel | Development | 2,968.0 | 228.0 | 2,943.3 |

Insertion/deletion clearly warrants further investigation as a source of DevExpress latency.
These numbers cannot replace the existing measurements of whole multi-sheet commands.

SpreadsheetGear AutoFill FillCopy took **2.687 ms Funding / 0.355 ms Development**, with
matching target formulas and constants. That is faster than Copy in this pass, but Copy
ran first and AutoFill second, with new columns inserted before AutoFill. Repeat balanced
ordering before treating the apparent speedup as established. Other engines' AutoFill paths
were not benchmarked, and formatting equivalence has not yet been exhaustively checked.

### Synthetic medians, milliseconds

Three fresh processes per engine; see methodology below. The deliberately simple formula
mix is not representative of the whole business plan.

| Operation | DevExpress | Aspose | SpreadsheetGear | Excel |
| --- | ---: | ---: | ---: | ---: |
| Load | 539.784 | 222.388 | 207.455 | 183.995 |
| Initial calculation/rebuild | 78.952 | 118.035 | 31.633 | 17.555 |
| Warm full calculation | 32.293 | 12.700 | 0.351 | 2.705 |
| Calculation after multiplier edit | 32.901 | 36.757 | 0.343 | 2.775 |
| Write 200,000 values | 416.359 | 45.616 | 10.809 | 919.517 |
| Read 200,000 values | 76.283 | 15.764 | 2.090 | 126.955 |
| Insert 10 columns | 192.608 | 49.708 | 21.811 | 14.568 |
| Copy column into 10 | 134.963 | 49.953 | 12.413 | 177.655 |
| Delete 10 columns | 173.195 | 6.598 | 7.953 | 11.286 |
| Insert 10 rows | 138.432 | 5.930 | 2.357 | 5.942 |
| Copy row into 10 | 7.832 | 6.030 | 8.821 | 185.211 |
| Delete 10 rows | 132.878 | 3.113 | 2.533 | 5.868 |

Excel's COM transfer/copy overhead is visible even where native calculation is quick.
This is one reason to measure the complete integration boundary, not calculation alone.

## What passed

The synthetic XLSM contains 4,096 data rows and 24,579 formulas. Three fresh processes per
engine, in rotated order, checked known answers at baseline, after three multiplier edits,
and after structural changes. Every data formula plus the three totals was checked in each
state. The tests also verified 200,000 transferred values, copied input values in inserted
rows/columns, and defined-name restoration after deletion. Validation time is not included
in operation timings.

The AGL comparison covers complete used rectangles in five sheets: Detailed Comp Inc - Trad
View, Financial Position - Trad View, Cashflow detailed, Check Sheet, and Development
Expenditure. There are 38,108 populated positions per candidate comparison. Blanks are
normalized; numerical tolerance is max(0.000001, abs(Excel value) * 1e-10). Matching these
probes is useful evidence, not accountant approval or proof of full-workbook equivalence.

Excel uses the embedded VBA PMCost/RespCost under the user's renewed, isolated permission.
Events remain disabled and AutomationSecurity is ByUI, respecting existing policy.
Separate non-zero probes explicitly invoked those functions: PMCost returned 80 and
RespCost returned 16, as independently expected. Thus VBA execution was demonstrated,
not assumed from saved caches. Scratch inputs were in a new, unsaved workbook.

DevExpress uses Summit's existing compiled UDF classes. Aspose and SpreadsheetGear use
trial-only .NET callbacks. Both reject malformed numeric lookup vectors instead of copying
VBA's error-suppression behavior; these ports are not production-approved replacements.
Valid blank rate/year slots are handled, and the final AGL runs report no rejected calls.

Following explicit permission, the range tests temporarily unprotect only the disposable
Funding and Development sheets using the existing structure field in memory. Protection
is restored in a finally scope; credentials are neither logged nor copied into this report.
They insert ten whole columns, copy one source column into them, then delete the ten columns.
Copy heights are identical: Funding 580 rows, Development 688 rows. The named anchors must
return to their original references, and after full recalculation all five output probes
must return to their pre-insertion values. No workbook is saved after the mutations.

SpreadsheetGear additionally tests AutoFill with FillCopy, not automatic date/number-series
inference. Fresh inserted columns are used, not already-populated Copy targets. Every
target formula and constant must match Copy. This does not yet compare all formatting,
conditional rules, validation, merged cells or the complete Funding/Development workflow.

## Interpretation and limits

- Calculation APIs are not identical algorithms. DevExpress uses Recursive calculation;
  Aspose disables its dirty chain for full calculations, then enables it for incremental
  tests; SpreadsheetGear and Excel use native CalculateFull/CalculateFullRebuild. Native
  caching, dependency and invariant-function optimizations remain enabled. Timings cannot
  establish identical numbers of internal formula evaluations.
- The AGL calculation table is a median of three warm full calls within one final process,
  not three independent full-workflow medians. Load/initial calculation and structural
  operations are single-pass observations. Synthetic results use three fresh processes;
  repeated metrics are medians of each process's median.
- DevExpress bulk writes use its range import API, but reads use public per-cell values.
  Aspose/SpreadsheetGear/Excel have bulk read paths. Aspose tiles Copy ten times; the other
  adapters use a single source/destination call. These are measured adapter paths, not
  proof that no faster vendor API exists. Excel includes COM marshalling costs.
- These source-sheet primitives are NOT the full VBA or Summit structural commands.
  Transactional DB mirrors, synchronizer sequences, rollback, end-to-end calculation,
  UI rebind, user history, save/reload and all reference/name invariants remain to test.
- No real-model save preservation has been established for SpreadsheetGear here.
  The earlier Aspose preservation findings, including fill/array interpretation concerns,
  remain open. This checkpoint does not weaken Summit's workbook-fill locking rules.
- Calculation-only offloading may lose the measured gain if each request requires
  staging, process startup, loading, saving and reloading. A persistent engine or narrower
  batch boundary needs a separate design and measured synchronization cost.

An early Aspose comparison showed 1,140 date-type differences: ExportArray returns formatted
dates as DateTime, while Excel Value2 returns serial numbers. The harness now normalizes
modern 1900-system dates before JSON serialization; those mismatches disappeared. Earlier
UDF adapter/cleanup diagnostic failures were corrected and superseded. No affected output
was treated as a validated client workbook.

## Next bounded trial

**Follow-up completed:** see `SpreadsheetGear_Funding_Full_Trial_2026-09-24.md`.
The full Funding trial is fast but fails grouped-sheet 3-D reference and saved-package
preservation gates; do not integrate the raw helper. The proposal below is retained
as the original checkpoint, not an outstanding authorisation to ship it.

Keep the DevExpress interface and production calculation path unchanged. Use the isolated
harness to port ONE complete Funding insertion command, including Transactional DB mirrors
and existing invariants, to SpreadsheetGear. Test Copy versus AutoFill in balanced repeated
order. Measure native operation AND the complete staging/save/reload/rebind cost. Compare
against the matching Excel/VBA operation on disposable copies.

Before any application integration: test XLSM save/reopen in Excel/VBA and Summit, VBA and
custom XML preservation, formulas/arrays/names, protection/fills/validation and representative
non-zero inputs across AGL, Blank, Demo and Stori. Then decide whether the narrow helper is
worth implementing; do not infer a whole-engine replacement from these numbers.

## Final harness checks

Debug and Release builds, x86 and x64: all four succeeded with zero warnings/errors.
Each passed the 19 SpreadsheetGear and 13 licensed Aspose smoke assertions. Production
Summit was not rebuilt or modified for this research delivery. The original AGL hash and
the shared XLSM input hash remained unchanged; no trial-owned Excel process remained.
The two Excel processes present before testing were left alone. No commit or push was made.
