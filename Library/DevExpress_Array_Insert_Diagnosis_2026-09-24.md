# DevExpress row-insert array diagnosis — isolated Gear worker trial

24 September 2026; DevExpress 25.2.4, .NET Framework 4.8, x64. Production Summit **2.98 is unchanged**. This continues `Spreadsheet_Mirror_Projection_Trial_2026-09-24.md`; it is not a production repair or customer-file certification.

## Reproduced cause

The AGL stage trace retains 51,948 TDB dynamic arrays through load, initial rebuild, 32 Funding column inserts, the production 3-D reference guard and column copies. During the eleven TDB row inserts/copies, only 781 new registrations appear instead of 4,510. The resulting 52,729 count stays unchanged through full rebuild, full calculation, serialization and native reload. The missing 3,729 declarations therefore precede saving.

The first TDB row insertion leaves every in-memory array collection coordinate unchanged, despite formulas moving ten rows down. This makes `HasDynamicArrayFormula` return true at some newly blank positions and false at the actual moved formula cells. A later CopyFrom sees those stale registrations and fails to register some of its copied dynamic formulas. Saved cells retain their expression and `cm` metadata but omit `f t="array" ref="..."`.

**Independent synthetic reproduction:** two blocks of five rows, six single-cell dynamic arrays per row. Insert ten rows before the gap, then copy a template row into them. Expected: 60 original + 60 copied = **120** arrays. Native unguarded result: **90**, both in memory and saved XML. The result repeats with/without intermediate inspection and with/without a native save/reload before the insert. Thus the instrumentation and AGL-specific VBA/UDFs are not required to reproduce it. In the saved raw fixture, C7 and moved C19 are arrays, but copied C9 is not.

This also explains why the previous post-copy guard reported zero reinstatements: its native cell-level test was consulting the stale registry. Do not use that guard alone after row insertion.

## Candidate containment (research only)

`MirrorDynamicArrayBatch` records the original single-cell array anchors, clears their registrations once while retaining their formulas, lets native row operations translate formula references, maps the known whole-row insertions and copied template anchors, then restores dynamic semantics through `DynamicArrayFormulaInvariant` **before calculation, validation or save**. No old metadata XML is pasted over new coordinates; no formulas are permanently converted to scalar expressions. An exception invalidates the disposable worker. Multi-cell spills and inconsistent manifests are rejected before staging.

The guarded synthetic matrix preserves all 120 anchors and values, survives native save/reopen, responds correctly to a later moved-row input change, and gives the expected Gear read-back. Multi-cell-spill refusals occur before mutation. A per-insert variant also passes the small fixture, but its AGL trial was stopped during the first guarded insert because repeated registration work was too expensive; no candidate was saved. The batch variant avoids repeating that work eleven times.

**AGL batch audit passes its bounded checks**, both in the instrumented direct run and the complete persistent-worker run. This is not full financial, VBA/UI or production acceptance.

## Final AGL checkpoint

- **38,108 populated values** across the five established probe sheets match the independent Excel/VBA Funding oracle (absolute tolerance 1e-6 / relative 1e-10). The previous three TDB Check Sheet failures are gone.
- **56,458 TDB arrays** are saved and recognized on native reload; no remaining metadata-marked formula lacks its array declaration. Array geometry matches the oracle across the workbook. All **1,126,327 formula-cell locations** match.
- Full formula-text comparison of the direct batch result finds only the same 570 Development Expenditure spelling differences; samples are `RespCost` versus `RESPCOST`. No other sheet differs in formula text. This does not classify every function or certify formula semantics for all inputs.
- Five scoped-name string differences were rechecked: each changes a one-cell rectangle to the identical single-cell reference. All 335 VBA module source bodies match the input; all nine Custom XML package parts and the drawing/chart/control families are retained. These checks do not certify compiled VBA, form designer storage or all formatting.
- Macro-disabled, event-disabled, read-only Excel 16.0 build 20326 opens the direct corrected checkpoint normally with 283 worksheets and VBA retained; no save or recalculation is requested. The five sampled Formula/Formula2 expressions match. Their HasSpill values are false in this manual open; that is not a new test of input-dependent spill behavior.
- The complete worker passes duplicate/gap rejection, failed-save preservation, revision N save excluding N+1, stale/locked destination rejection, controlled native-worker termination and immutable-baseline replay. Debug/Release synthetic native-result freshness/failure tests pass; the optional Excel synthetic route also passes. The final source additionally faults a worker after calculation/export-validation failure; the injected pre-publication test remains recoverable.

| Complete persistent DevExpress worker observation | Time |
|---|---:|
| Warm Gear edit/journal/calculate/read | 0.55–0.76 s |
| Native worker save, calculation and package checks | 15.48 s |
| Funding command plus native rebuild | 279.94 s |
| Checkpoint save | 17.93 s |
| Gear rebase/rebuild | 4.83 s |
| Funding + save + rebase | **302.73 s** |

Array suspension/restoration accounts for **121.02 s** of that command. This is a correctness workaround, **not a structural speed improvement**. The earlier optional Excel worker took about 29 seconds for the same full barrier. The five-minute worker-reply timeout has little headroom on this development machine; a production protocol needs progress/cancellation and an appropriate long-operation policy, not an assumption that these timings generalize.

The standard isolated DevExpress worker now selects the batch correction. Before publishing its private candidate it checks expected live TDB array declarations against the saved package; the synthetic gate rejects incomplete output and accepts corrected output. This is a targeted guard, not proof that an already-invalid input was complete, nor a full formula/name/VBA integrity gate.

## Production boundary

A read-only native inspection of a byte-identical original AGL **XLSB** reports **0 dynamic / 51,948 legacy TDB arrays**. The common converted **XLSM** is read as 51,948 dynamic arrays. Thus this dynamic-array reproduction must not automatically be attributed to the production XLSB path. Production's bounded `InsertCells(...ShiftCellsDown)` alternative also reproduces the issue in the small **dynamic-array XLSM** fixture, so merely switching insertion APIs is not a repair. Actual production XLSB structural/round-trip behavior remains a separate validation task.

Production Debug/Release remain byte-identical 2.98, and AGL/Blank/Demo hashes match the preceding checkpoint. No original workbook, production source, package version, master or license was changed; no support request was submitted. The durable Blueprint and project index point to this checkpoint. No commit/push was requested.

## Evidence and commands

- Raw AGL: `obj/AsposeTrial/array-trace-20260924-1/` — per-stage manifests, final package/native state and independent failing oracle audit.
- Small reproduction: `obj/AsposeTrial/array-insert-probe-20260924-2/`; per-insert validation: `array-insert-probe-20260924-release-1/`.
- Batch matrix: `obj/AsposeTrial/array-insert-probe-20260924-batch-debug-1/` and `array-insert-probe-20260924-batch-release-1/`.
- Stopped slow variant: `obj/AsposeTrial/array-trace-20260924-fixed-1/` (incomplete; no publication).
- Full-model batch: `obj/AsposeTrial/array-trace-20260924-batch-1/`.
- Complete Release worker: `obj/AsposeTrial/mirror-20260924-v3-dx-agl/` — report plus independent oracle audit.
- Final export-gate/cell-insert matrices: `array-insert-probe-20260924-cells-debug-1/` and `array-insert-probe-20260924-cells-release-final/` (16 cases each). The earlier `gate-debug-1` is a retained failed test of A1:A1-versus-A1 normalization in the new assertion; the assertion was corrected and rerun successfully.
- Final worker safety runs: `mirror-20260924-v3-dx-synthetic-debug-final/`, `mirror-20260924-v3-dx-synthetic-release-final/`, `mirror-20260924-v3-excel-synthetic-release/`.
- Original XLSB interpretation: `obj/AsposeTrial/native-array-source-20260924.json` (private source hash `30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47`).

All paths above are ignored private diagnostics, not files to send clients. Commands in the standalone `SpreadsheetEngineTrial` executable:

```text
mirror-array-insert NEW_PRIVATE_DIRECTORY
mirror-array-trace PRIVATE_BASELINE.xlsm NEW_PRIVATE_DIRECTORY
mirror-array-trace-fixed PRIVATE_BASELINE.xlsm NEW_PRIVATE_DIRECTORY
mirror-excel-array-open PRIVATE_CANDIDATE.xlsm NEW_REPORT.json
```

The trace is instrumented, includes disk manifests, and is not a clean speed benchmark. The fixed command currently selects the batch experiment; the unqualified trace intentionally retains the failing native path. Neither is invoked by production Summit.

## Vendor context and limits

The workaround uses the documented native [dynamic-array management API](https://docs.devexpress.com/OfficeFileAPI/14942/spreadsheet-document-api/formulas/array-formulas) and [invariant dynamic formula property](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.CellRange.DynamicArrayFormulaInvariant?v=25.2). A [DevExpress worksheet-copy ticket](https://supportcenter.devexpress.com/ticket/details/t1316801/spreadsheet-document-api-dynamic-array-formulas-are-copied-as-normal-formulas-on-copying) concerns a different operation; it is not vendor confirmation of this row-insert defect. No package upgrade or support submission has been made.

Still required: reduce the workaround cost or obtain a vendor fix; extend beyond AGL Funding to Development/add-delete and varied input-dependent spill cases; classify consumers that need native calculation; evolve XML/history during replay; validate complete formatting, actual VBA execution and production-safe publication/cancellation. Gear's missing dynamic spills, grouped-sheet fix-ups and Custom XML support remain separate restrictions. The reduced Gear projection must never save the authoritative workbook.
