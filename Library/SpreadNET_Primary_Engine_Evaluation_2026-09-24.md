# Spread.NET as a primary, same-process engine

24 September 2026. Research checkpoint, **not a Summit test release**.

## Decision

**Do not replace DevExpress or purchase Spread.NET for this purpose on the present evidence.**
It can run in-process on .NET Framework 4.8 without Microsoft Excel, and its basic calculation
and transfer performance is useful. However, the tested AGL calculation and Excel round-trip
paths do not meet Summit's contract. Vendor clarification or a corrected supported configuration
is required before extending the prototype.

This is a finding about the tested APIs/configuration, not a claim that MESCIUS has confirmed
a general product limitation. No customer file, master, production dependency or executable was
changed. No workbook was uploaded. No worker repaired the Spread result.

## Setup and timing controls

- Core/WinForms NuGet 19.2.0, assembly 19.2.20266.0; x64 Release .NET Framework 4.8.
- Same development machine and same private XLSM fixtures as the previous four-engine study.
  This is not a minimum-spec client or cold-cache study.
- Primary adapter: an undisplayed FpSpread, LegacyBehaviors.None, same STA process. Native
  workbook APIs do the calculations and range operations; DevExpress controls are not migrated.
- Dynamic arrays enabled; background and on-demand calculation disabled; manual calculation.
  Insert repros fail both with and without the optional Optimal flag.
- PMCost/RespCost valid-input ports reuse the earlier verified algorithms. They must be
  registered before OpenExcel. Standalone native known answers are 80/16. The corrected AGL
  diagnostic invokes each function 950 times with zero adapter-invalid calls.
- DocumentCaching on load/save; XLSM flags; read/write output streams; native return values
  checked. A write-only-stream error in early pilots was corrected and is not a product defect.
- Trial reminders occur during host construction, outside the operation timers. Reaction time
  is excluded from the measured calculation/range operations. No licence bypass was used.
  A redundant final batch was stopped while awaiting its initial reminder, before timing began.

## Benchmarks that passed

Three fresh Core-factory processes passed the common synthetic fixture: 4,096 rows, 24,579
formulas, five known-answer states, 200,000-value transfer checks, copied payload checks and
post-delete/name restoration. This fixture deliberately does not exercise model UDFs or grouped
structural commands. Core UDF registration remains unresolved; these results do not certify AGL.

Milliseconds; Spread medians measured now, other columns are the earlier same-fixture,
same-machine medians, **not a newly interleaved five-engine run**:

| Operation | Spread Core | DevExpress | Gear | Excel |
| --- | ---: | ---: | ---: | ---: |
| Load | 362.262 | 539.784 | 207.455 | 183.995 |
| First rebuild | 74.147 | 78.952 | 31.633 | 17.555 |
| Warm full calculation | 16.720 | 32.293 | 0.351 | 2.705 |
| Calculation after input edit | 8.699 | 32.901 | 0.343 | 2.775 |
| Write 200,000 values | 3.348 | 416.359 | 10.809 | 919.517 |
| Read 200,000 values | 6.902 | 76.283 | 2.090 | 126.955 |
| Insert 10 columns | 63.378 | 192.608 | 21.811 | 14.568 |
| Copy into 10 columns | 32.779 | 134.963 | 12.413 | 177.655 |
| Delete 10 columns | 31.421 | 173.195 | 7.953 | 11.286 |
| Insert 10 rows | 10.932 | 138.432 | 2.357 | 5.942 |
| Copy into 10 rows | 8.103 | 7.832 | 8.821 | 185.211 |
| Delete 10 rows | 23.045 | 132.878 | 2.533 | 5.868 |

Spread reads use a public per-cell API loop; writes use a block API. Copy explicitly tiles the
source. Peak working set was about 74 MiB in these simple runs. These numbers exclude Summit
UI refresh, validation, undo journalling, protection restoration and full business commands.

## AGL: diagnostic timings, not an accepted speed ranking

The corrected Windows-host diagnostic loaded AGL in 11.429 s and completed its requested full
rebuild in 17.186 s (Optimal enabled). UDF calls were verified, but output parity failed:
513 numeric, 2,451 type and 12 text differences across the union of five established probe
areas; 2,463 candidate error values. The union contains 38,115 nonblank positions versus
38,108 in the earlier accepted oracle. Development Expenditure's UDF-related errors clear
after correct registration, but other failures remain. Do not count earlier zero-UDF runs.

For context only, the prior **passing** AGL first-rebuild/warm-full medians were:
DevExpress 6.318/5.989 s; Aspose 4.120/3.655 s; Gear 0.694/0.544 s; Excel 4.041/0.392 s.
Spread has not demonstrated a correct real-model speed improvement. No completed whole-model
performance comparison or end-to-end Funding/Development speed is claimed.

## Compatibility findings

1. **Real-model save gate fails.** DocumentCaching save-only AGL output reopens in Spread,
   but independent DevExpress LoadDocument returns false and macro-disabled Excel Open fails
   with 0x800A03EC. The identical Excel reader opens the unchanged input successfully. No
   repair-mode open, external restamping, VBA execution or source save was used.
2. **Preservation is partly positive, not sufficient.** All three custom XML payloads and
   their item-properties parts, plus vbaProject.bin, are byte-identical. All 283 sheets are
   present. Custom XML relationships and dynamic metadata are rewritten. Transactional DB
   array anchors/metadata cells grow from 51,948 to 122,039 in the save-only package. That
   change needs explanation; changed metadata bytes alone are not proof of corruption.
   No duplicate ZIP entries; cell-format count is 1,609 before and 1,582 after, so this is
   not evidence of a simple format-count explosion.
3. **Native scalar result after insert is wrong in the small repro.** A1=3, D10=A1*2 gives 6.
   Insert two rows above row 5 and two columns before C: formula moves to F12 and remains
   A1*2, but a native full rebuild returns 0 instead of 6. Reproduced through IWorksheet,
   IRange and SheetView insertion paths, with Optimal on and off.
4. **Native spill growth after reopen fails in that repro.** SEQUENCE(A1) grows correctly in
   a fresh workbook. After insert/save/reopen and A1=5, the native B5 result is blank. The
   same three saved files open in macro-disabled Excel, return F12=6 initially, then F12=10
   and B5=5 after A1=5. Their files remain unchanged. Thus small-file Excel metadata can be
   valid while Spread's own calculated results are wrong in the tested route.
5. **Grouped-sheet semantics are not established.** Selecting First/Last then inserting via
   IRange, IWorksheet or ClipboardInsertCommand moves only the first sheet's values in these
   tests. Serial insertion moves both, but SUM(First:Last!B1) retains its original text and
   stale value 12 after rebuild. A supported grouped-operation entry point and 3-D fix-up
   behavior require vendor confirmation; do not turn this into hand-written formula repair.
6. **Fills/locks, names and objects are not signed off.** The independent full-cell manifest
   cannot be completed because DevExpress rejects the AGL output. Part counts and byte-identical
   VBA are not a substitute for workbook-owned fill/lock semantics or standalone VBA behavior.

One native AGL save-only observation: load 11.246 s, save 5.601 s, Spread reopen 9.541 s;
peak process working set about 1.06 GiB across the two loads. The save is **not usable evidence
for a safe save policy** because its Excel/DevExpress acceptance gates fail.

## Migration impact and next gate

A same-process engine avoids worker synchronization and the installed-Excel dependency. It is
still not a drop-in replacement: workbook/cell/range types, native range data sources, formatting,
fill-based locking, validation, custom functions, history/transactions, structural commands and
serialization must be adapted. A broad inventory finds 81 workbook-coupled VB files including
legacy code; that is an indicator, not a production edit estimate. Keep DevExpress grids and
look-and-feel, but expect substantial data-adapter work. Production x86 memory, larger client
models and Excel/VBA round trips would still need separate validation. Existing permission to
consider XLSM removes one format obstacle; it does not waive preservation requirements.

Recommended next step is a MESCIUS support check with the minimal non-confidential repros and
the exact APIs/settings. Do not send AGL without explicit authorization. Obtain a supported
group-insert route, native calc/reopen explanation and a successful Excel-readable AGL save
before investing in a full Funding/Development port or further large benchmarks. Continue using
production DevExpress; this evaluation does not reopen the Gear worker integration automatically.

## Reproduction and evidence

Harness: `Tools/SpreadNetTrial`; original common inputs under
`obj/AsposeTrial/a79b3d867f81456a9d22887b9600ee61`.
AGL input SHA256: `62B44DE2764C4D1D1C9C7029C9297EEF4D461E300BAB4AD7DB092BBE1D23193C`.

- `obj/AsposeTrial/spreadnet-20260924-final/synthetic-core-1.json` through `-3.json`:
  accepted simple measurements; `agl-correct-udf-comparison.json`: failed real-model parity.
- `.../spreadnet-20260924-v5/agl-registration-before-open.json`: corrected UDF diagnostic.
- `.../spreadnet-20260924-v5/agl-io.json`, `agl-package-audit.json`,
  `agl-excel-open.json`, `agl-source-excel-open.json`: save/acceptance evidence.
- `.../spreadnet-20260924-v6/cases/cases.json`, `core-cases/cases.json`, `verify-excel.json`:
  native and independent small insert/spill evidence. v5 cases show Optimal-on behavior.
- v1-v4 and v5 `agl-1`/`agl-2` are setup/registration pilots, not accepted model benchmarks.

Trial sources build in Debug and Release. Production/master/original hashes were separately
checked; no production version bump or client test-release claim is appropriate for this work.

Official references checked during the evaluation:
[packages](https://www.nuget.org/packages/GrapeCity.Spread.WinForms/19.2.0),
[native dynamic arrays](https://developer.mescius.com/spreadnet/docs/latest/online-win/overview/spwin-devguide/spwin-cell-formula/WorkingWithDynamicArrayFormulas),
[document-caching save contract](https://developer.mescius.com/spreadnet/docs/latest/online-win/overview/spwin-devguide/spwin-usefiles/spwin-savefiles/spwin-save-excelfile),
[XLSM save flags](https://developer.mescius.com/spreadnet/api/latest/online-win/FarPoint.Excel/FarPoint.Excel.ExcelSaveFlags.html),
[custom-function registration](https://developer.mescius.com/blogs/spread-dot-net-calculation-part-three).
