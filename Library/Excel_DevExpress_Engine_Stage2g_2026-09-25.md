# Stage 2g: existing change manager and authoritative presentation

## Status

Stage 2 remains in progress. This checkpoint is not a client release and does not enable Excel in normal model opening or Options. Normal Debug/Release 2.98 and original AGL/Blank/Demo bytes remain unchanged. The user requested continued work through all Stage 2 gates; an intermediate checkpoint is not a stopping point. See `Excel_Stage2_Work_Plan.md`.

## Implemented and tested

- `WorkbookValueBatches.vb`: one serialized native transaction for up to 10,000 distinct typed inputs. Checks every before-state before the first write, supports explicit prerequisite calculation and soft-skipped unavailable targets, calculates final results, and returns one revision. Failure compensates all touched cells, recalculates and verifies the complete original state; unverifiable restoration quarantines. Native calculation/read failure in an unchanged batch also quarantines.
- `ModelEngineChanges.vb`: an internal bind on a clean real model makes the existing `ModelChangeManagerV2` own native edits, dirty state, journal, superseded actions and Undo/Redo. No second history stack. Existing pure type conversion is shared with the ordinary write path. Dates respect 1900/1904. Native writes never overwrite the presentation workbook's formulas/caches. Unspecialized edits, grouped callbacks and ordinary/recovery saves are refused while this isolated binding is active; they must not silently save the stale presentation workbook.
- `WorkbookPresentation.vb` / `NativePresentation.vb`: immutable appearance and display text accompany values at the same owner/revision/calculation generation. Includes effective number format, fill-based availability, font/colour, alignment, indentation and wrapping. Excel uses `DisplayFormat`, not base `Interior`; DevExpress Cell.Fill already includes active conditional formatting. Tests reproduced the old mismatch before correcting the adapter; Summit's existing fill-based rule is unchanged.
- Every calculation invalidates older renders even without an input revision change. `ReadCurrentAsync` can obtain further bounded rectangles against the same current calculation without recalculating. Canceled/obsolete results never become a current display.
- `ModelWorkbookRead.vb`: a common read facade uses the selected owner's detached result, with small same-generation cache-miss reads; unbound models use their original DevExpress reads. DevExpress exposes distinct public Workbook and worksheet document identities: both are bound explicitly, proved in the real fixture. Shared DataManager reads, mapped-table values/colours/fonts and Check Sheet/override reads now use this facade. This is a **partial consumer migration**, not a claim that every display/direct read has been converted.
- Native committed edits and Undo publish the existing soft Check Sheet state and company-warning events from the same accepted owner result. Formula errors remain distinct from an imbalance. Existing Yes/No override semantics are retained.

## Validation and evidence

Separate application builds: `bin/EngineStage2g-Debug` and `bin/EngineStage2g-Release`; normal outputs untouched. Both compile without errors. Serial x86 harness runs in `obj/EngineStage2g/bfb96a00eccf4932b3aa8a111ef2c16f` pass both configurations: batch 48, real change-manager bridge 101, native appearance 33, edit 62, history 39, candidate save 54, terminal publication 76, recovery 109, safety 72, deadline 6, grid 16, security 12 and the deliberately rejected scalar projection probe 10. Native candidate save/reopen: 37 assertions over XLSX/XLSM/XLSB in both engines, including conditional input formatting, formulas, dynamic spill and other-engine full calculation. Existing user Excel processes remain; test-owned instances exit.

The real bridge covers grouped typed edits, before-state rejection, Undo/Redo, real dirty/history ownership, source-cell preservation, 1900/1904 date systems, dynamic spill, current Check Sheet/company warning, and defining-date-dependent amount availability (including skip, Undo and Redo). It is not a live DIT or whole-AGL financial certification.

The existing production input fixture needed its three reflective paste calls updated for the current optional vertical-record-map parameter; no production paste behavior was changed. The complete disposable-Demo regression passed **550 checks** before and after the shared-reader changes, including native inputs, paste, Undo/Redo and independent save/reopen. Current facade-run evidence: `obj/ClientReportTests/e53ef02ce4264815ae680547b312408c`; earlier run `5c6e86feea364544b281442ef102f08d`. A WebView2 shutdown class-unregistration diagnostic is present; the fixture exits successfully. This does not establish visual high-DPI/client acceptance.

## Remaining Stage 2 gates (do not hide)

1. Complete live editor admission/before-state and read-consumer routing, including in-place headers, mapped overrides, dependent validation lists, analyzers/RangeDataSource, dashboards and formula-reader helpers. Direct template reads must not silently display stale values when an engine is bound.
2. Complete actual per-model opening/preference/status and lifecycle; only expose selection once its live edit/read/save paths are genuinely owned. The isolated backend's second DevExpress workbook is not a memory-efficient production fallback design by itself.
3. Couple continued Save/Save As/recovery and exact-revision dirty acknowledgement to the selected owner. Terminal verified publication/reopen exists, but ordinary UI save/rebind does not. Current schedule/structure XML must not be guessed from a stale template. Structural actions stay unavailable until their Stage 3 owner path is qualified.
4. Appearance transport currently reads effective Excel formatting cell-by-cell: 16-cell synthetic transfer was about 100 ms. This is correctness evidence, not a performance win; large Funding views require batching/caching and realistic end-to-end measurements before release. No claim that a fast calculation alone makes the full UI faster.
5. Native hard-hang/process supervision, policy/prompt/storage qualification and real manual Summit/Excel/VBA round trips remain release gates. Native timeout quarantine queues cleanup but cannot abort permanently blocked COM.

## Primary references

- [Excel DisplayFormat includes conditional formatting](https://learn.microsoft.com/en-us/office/vba/api/excel.range.displayformat).
- [DevExpress CellValueToStringConverter](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Export.CellValueToStringConverter._members): supported formatting of a detached numeric value. A tiny formatting-only workbook contains no model formulas; it is not another financial calculation engine.
