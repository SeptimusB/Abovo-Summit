# Excel-first integration checkpoint 2a: result grids and deadlines

Continuation of `Excel_DevExpress_Engine_Stage1_2026-09-25.md`. User has also approved all three Stage 3 structural routes; the experiment is specified in `Structural_Engine_Comparison_Protocol_2026-09-25.md`.

**No live engine switch or new client test release.** Normal 2.98 Debug/Release, customer originals, masters and workbook fill/locking rules remain untouched. This checkpoint implements and tests the display/supervision building blocks. Options, live edits/history, structural commands and save ownership are not yet connected.

## Decisions demonstrated in native tests

The proposed `ICustomCalculationService` shortcut is **not eligible as a general display projection** in the installed DevExpress 25.2.4:

- Scalar cached values can be replaced while preserving their formula text, static patterns, colours, formats and locks. An existing `RangeDataSource` sees refreshed scalar values.
- A previously calculated two-cell dynamic array retains an old spill value after the anchor is replaced by the scalar callback (anchor 700, spill still 2 instead of supplied 701). A new two-cell array does not populate its second cell through this callback. Therefore do not install this service into a production workbook or assume metadata/formula preservation proves result freshness.
- Callback `SheetId` was zero where public `Sheet.Id` was one in the synthetic case. The probe's one-sheet mapping is not production identity logic. No inferred mapping is deployed.
- The callback implementation lives only in `Tools/WorkbookEngineFoundationTests/ProjectionProbe.cs`. It does not run in Summit. Ten assertions reproduce the supported scalar behavior and the disqualifying array behavior. This is a negative suitability finding, not a passed general projection.

The selected display boundary is instead `Services/WorkbookEngines/WorkbookResultGridSource.vb`: a read-only `BindingList`/`ITypedList` over accepted immutable engine rectangles. It never writes or recalculates a worksheet, so it cannot change formulas, array definitions, protection or fills. Native GridView binds to this source and refreshes in place, with one reset notification per accepted result and unchanged row-address mapping. Numbers, text, booleans, blanks and errors retain distinct types.

The source checks exact session/revision and rectangle, refuses stale/foreign results and makes stale reads unavailable even before a queued controller refresh. UI controllers still have to invalidate/refresh the displayed view on edits and show calculation state; this class does not yet wire those events into existing DITs. Property setters and row mutations are refused. All binding access is confined to the UI owner thread. Workbook formatting and fill-based editability must be obtained through the established metadata path; result values alone must not be used to infer either.

## Bounded supervision

`WorkbookEngineOptions.OperationTimeoutMilliseconds` defaults to 120,000 for this calculation/opening API; test overrides are bounded between 50 ms and 30 minutes. Structural-command deadlines will need separate measured policy. No visible setting has been enabled prematurely.

Opening/calculation deadlines return a timeout, quarantine the session, reject results/reuse and queue cleanup on its STA owner. Close also has a bounded wait. `NativeCleanupCompletion` distinguishes actual eventual cleanup from the caller's timed-out wait. No native call is aborted or disposed from another thread, and a late opening cannot start a competing fallback. Stale-result rejection is separately typed so a normal newer revision does not accidentally fault the engine.

**Remaining supervision gate:** an Excel/DevExpress call that never returns still needs explicit process-level recovery. This checkpoint neither kills Excel nor claims to have crash-isolated DevExpress. The pending native owner/lease remains until cleanup can execute. Do not enable live ownership until the working-copy, failure UI and isolated-process recovery policy are complete.

## Validation

Both source and harness Debug/Release builds pass into separate `EngineStage2-*` folders. On each build:

- Safety regression: 72 assertions (one earlier backend-read assertion no longer occurs because stale calculations are now rejected before reading).
- Deadline cases: six assertions for held opening/calculation, bounded close, no mid-session fallback and correct late cleanup.
- Native result-grid cases: 16 assertions covering typed data, readonly behavior, batched updates, current revision, row mapping, native GridView invalidation/refresh/multiselect and cross-thread rejection.
- Security screening: 12 assertions, preserving downloaded-file markers and refusing reviewed unsafe package types.
- Callback probe: ten assertions documenting why that route is not a general projection.

AGL XLSB, Debug: the two engines agree on 45,586 positions in the existing five financial/check rectangles. Both engines' accepted results bind to real native grids (45 sampled displayed positions per engine, with complete rectangle dimensions checked). Private source hash and existing Excel process list are unchanged; the owned Excel process exits. Required-VBA-disabled open selects DevExpress.

AGL XLSM, Release: the same final native comparison/grid fixture also passes all eight integration assertions, including parity across 45,586 positions, native grid binding, owned Excel cleanup, security fallback and unchanged source bytes. These are integration smoke tests, not an uncontended speed comparison: other short tests ran during part of the session. Use Stage 1's labelled timings only within their original scope and Stage 3's new protocol for route comparisons.

Ignored evidence: `obj/EngineStage2-{safety,deadline,grid,security,projection}-{debug,release}.log`, `obj/EngineStage2-AGL-xlsb-debug.log`, `obj/EngineStage2-AGL-xlsm-release.log`. No result workbooks, extracted VBA, licences or attachments are included in published code.

## Next integration slice

1. Connect authoritative input/result/format reads and expected-before typed writes to `ModelChangeManager`, retaining undo/dirty/revision and fill-driven locking. No direct writes to a display cache as a substitute for workbook edits.
2. Connect model lifecycle and preference/status only after the model genuinely has one authoritative owner and reliable hung-session recovery. Preserve DevExpress-only behavior.
3. Adapt mapped/ordinary/analyser grids to the result source or equivalent authoritative readers, with controller invalidation, formatting, pending editor and viewport tests. Avoid the scalar callback shortcut for arrays.
4. Execute the three-route Stage 3 protocol and integrate verified structural/save/recovery ownership. Full Excel/VBA round-trip and Jon/Alex/accountant acceptance remain required.
