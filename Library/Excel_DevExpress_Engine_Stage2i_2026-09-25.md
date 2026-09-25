# Stage 2i: current analyser ranges, validation and real DIT acceptance

Internal checkpoint only. Stage 2 remains in progress; normal Debug/Release 2.98 and both authoritative masters are byte-identical to the previous checkpoint. Continue `Excel_Stage2_Work_Plan.md` without treating this checkpoint as a client release.

## Implementation

- `ModelRangeDataSource` binds the two income/expenditure analysers to bounded, detached, revision-bound native values. A private value-only document lets the supported DevExpress range datasource retain column normalisation/conversion; it contains no financial formulas and is never calculated or saved. Explicit text/numeric analyser schemas, hidden rows/columns, filters and summaries are qualified. Default inferred schemas and editable range datasources are not yet supported by this adapter.
- Native scalar evaluation now supplies numeric validation limits and expected-chart-gap predicates. Excel evaluates in the worksheet context; DevExpress uses array expression mode so accidental implicit intersection cannot disguise a multi-cell result. External/structured-reference syntax, oversized formulas and non-scalar results are refused. No temporary cells, names or selection changes are used.
- Company/date profile reads use the accepted owner and actual workbook date system. Existing worksheet cells remain structural/address metadata, not a second calculation owner.
- Actual native DIT testing exposed repaint during a pending calculation. Paint-only readers show an ellipsis; strict values, validation and editability still refuse pending results. Previously bound analyser rows return null while pending/stale/closed. This does not change the workbook fill-pattern permission rules.
- The whole-DIT fixture selects Funding's opening-balance input through `Rep_Fund_08`, not an AGL-specific coordinate. The current Demo maps its first eligible input to E81. Both real editors post through the existing manager, and Undo restores native values without rewriting the presentation workbook.

## Verified evidence

- Isolated Debug and Release builds pass: `bin/EngineStage2i-Debug`, `bin/EngineStage2i-Release`.
- Actual private Demo/Covenant/Funding UI: 17 assertions per native engine, original input bytes unchanged and owned Excel exits. Final logs `obj/EngineStage2i-Debug-dit-dx.log` and `obj/EngineStage2i-Debug-dit-excel.log`. The test closes the application model and explicitly closes the trial session; production close ownership is still a separate next gate.
- Release final suites: range 32, expression 32, existing bridge 133, grid 16. Logs `obj/EngineStage2i-Release-*-final.log`.
- Actual private AGL analyser: 2,490 rows, 74 columns, all 184,260 detached values match between native engines within the fixture's numeric tolerance. Binding/read timings in that pass were 1,079 ms DevExpress and 4,314 ms Excel, excluding calculation. Evidence `obj/EngineStage2i-Debug-range-agl.log`.
- Ordinary production DIT regression: 550 checks including private save/reopen, `obj/ClientReportTests/80b824e85e9a414da21f8393b55a9cce`. Existing WebView2 shutdown diagnostic 1412 remains with successful exit.
- Whole Funding section binding was about 1.5 seconds DevExpress versus 23.2 seconds Excel in the final UI passes. Native appearance transfer is an unresolved performance cost, not a claimed speed win.
- The first UI fixture revealed a test-harness cleanup wait capturing the WinForms context; its polling now uses `ConfigureAwait(false)`. Only the verified fixture process was stopped; no user Excel process was touched. Both engines' final process inventories match the pre-test inventory.

## Remaining gates

Remaining read consumers/default schemas, production close ownership, continued Save/Save As/recovery and exact dirty acknowledgement, current non-history model XML, safe unsupported structural commands, opening/options, native hang/security/storage cases and manual Excel/VBA round trips. The route remains internal; original DevExpress production behavior is preserved.
