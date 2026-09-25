# Excel-first checkpoint 2b: guarded in-memory value edits

Continues `Excel_DevExpress_Engine_Stage2a_2026-09-25.md`. This is an isolated implementation checkpoint, not a client release. Normal Debug/Release 2.98, its calculation/save behavior, source masters and client originals remain unchanged.

## Implemented

- `WorkbookEngineOptions.EnableValueEditTrial` is false by default. Only an explicitly opted-in calculation session can apply the new value operation. The source file remains leased read-only and Excel opens it read-only. Native changes occur only in memory and closing discards them. There is still no Save API.
- `WorkbookValueEdits.vb` defines immutable input-state snapshots, typed value-edit receipts and the optional native backend extension. Snapshots carry session and revision identities. The owner rereads the actual cell before writing, and rejects stale/foreign snapshots or changed value/permission/number-format state.
- `ApplyValueAsync` invalidates previous results before writing, performs full native calculation and obtains the requested bounded output rectangles in the same serialized operation. Accepted input state and output values share one revision. Unchanged input returns no-change without calculation or a new revision.
- Inputs are finite Double, Boolean, literal String, or blank. Date serial conversion, model validation, Yes/No normalization and existing XML/editor restrictions remain the responsibility of the existing change-manager layer when connected. Empty text input means blank. The adapter does not turn text beginning `=`, `+`, numeric/date-looking strings or apostrophes into formulas/numbers.
- Backend partial files use native cell writes on their STA owner. Excel COM references are released there. There is no worksheet unprotect, background engine switch, second writable workbook, formula replacement or direct result-cache editing.

## Permission and scope boundaries

DataManager currently has distinct lock-bit and `HasRules` fill-pattern paths. The adapter retains an explicit `UnlockedCell` versus `SolidFillRule` choice rather than silently combining or replacing them. DevExpress reads `Cell.Fill.PatternType`; Excel reads the underlying `Interior.Pattern`. This reproduces the static-fill rule being used by Summit, **not** a claim that conditional-format display rules and raw fills are interchangeable. Full presentation/conditional-format transport and Funding date-driven unlock coverage are still required before UI integration.

This narrow trial conservatively refuses formula inputs, array members (including spill children), merged cells, error-valued inputs and actually protected locked cells. It does not relax production permissions or decide how to edit an authorised formula input. The real XML, row/editor restrictions, data validation and `ModelChangeManagerV2` are unchanged and are not bypassed by live controls.

Only selected input definition fields are captured here: number format, lock, solid-fill eligibility, array/merge membership and sheet protection. This is **not** a complete font/colour/validation/rich-text metadata transport or a whole-workbook transaction snapshot.

## Failure semantics

Writes, calculation and read-back are one owner operation. On failure/cancellation, the owner restores the original value, fully recalculates and verifies the original captured state. Failed attempts produce no successful receipt. The incremented revision remains invalidated so a caller must refresh, even after successful restoration. If restoration fails, the session is quarantined and cannot serve results or silently fall back.

The mutation action, not the generic read-only dispatcher, owns cancellation. A cancellation before commit compensates the edit; cancellation after a completed commit must not discard its receipt. A timeout quarantines immediately and returns no success. If the held native call subsequently returns, restoration and cleanup finish on its owner. Calls that never return still require the separate process-level recovery design; this checkpoint does not solve that release gate or kill user Excel.

The receipt is the future handoff to `ModelChangeManagerV2` for history, dirty marking and notification. No duplicate history stack was introduced, and none of those live integrations is claimed complete. Full model compensation against side-effecting VBA/UDFs is not established by restoring one cell; only the verified pure calculation/edit cases below are covered.

## Validation

Source and harness Debug/Release builds use separate `EngineStage2b-*` folders. Deterministic tests cover 62 edit assertions per build, including foreign/stale snapshots, over-limit reads, typed input guards, no-op behavior, before/after receipts, both permission paths, partial write/calculation/read failure restoration, cancellation during calculation/read-back, concurrent edits and bounded timeout/late cleanup. An unexpected native input-read failure quarantines the session; a detected untracked change invalidates cached results before rejecting the edit.

Both configurations retain the previous safety (72), deadline (6), result-grid (16), security (12), and negative projection (10) assertions. Synthetic native edit tests cover 49 assertions for literal/typed input preservation, unlocked formula/spill guards, protected/unprotected inputs and unchanged source bytes with Excel and DevExpress.

Native AGL XLSB Debug and converted XLSM Release both pass seven assertions plus the harness source-hash check. Funding Assumptions G82 changes from 2,700 to 3,000 in memory. Both engines agree across the established 45,586 output positions after full calculation. Applying the original value through the same guarded boundary restores those outputs to each engine's baseline. Sheet protection stays enabled, source bytes are unchanged, owned Excel exits, and pre-existing Excel processes remain.

Single integration-trial timings on the development machine (not a multi-run benchmark, client guarantee or live UI latency):

| Format/build | DevExpress edit/calc/read | Excel edit/calc/read |
|---|---:|---:|
| XLSB Debug, final validation | 6,453 ms | 549 ms |
| XLSM Release, final validation | 6,843 ms | 590 ms |

The corresponding calculation portions were 6,426/419 ms and 6,815/424 ms. Earlier smoke runs measured 8,958/606 ms and 9,387/666 ms total; that variation is why these are not treated as controlled benchmark medians. Timings exclude workbook opening, visible control refresh, history/dirty processing, save and close. They must not be presented as the Stage 3 insertion comparison.

Local ignored evidence: `obj/EngineStage2b-*-debug.log` and `obj/EngineStage2b-*-release.log`. No workbook payload, extracted VBA, licence or client attachment is included in the published checkpoint.

## Next work

Checkpoint 2c now implements opt-in candidate export and bounded cross-engine save/reopen validation: `Library/Excel_DevExpress_Engine_Stage2c_2026-09-25.md`. Original-file publication and live Save/dirty/history integration are still pending.

1. Connect the operation to the existing change manager's typed validation, journal, dirty/recovery revision and UI events. Keep the current DevExpress path intact. Add authoritative formatting/permission refresh and multi-cell history tests before enabling an engine option.
2. Implement security-preserving private working copies, reliable hung-session ownership/recovery, candidate saves with exact revision verification and explicit replacement. Saving remains disabled in this adapter until that boundary is ready.
3. Execute the approved three-route structural comparison (`Structural_Engine_Comparison_Protocol_2026-09-25.md`). Complete-command and round-trip correctness still gates performance recommendations.

## API references used

- [Microsoft Range.Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2) and [HasSpill](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hasspill).
- [Microsoft DisplayFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.range.displayformat) distinguishes raw cell formatting from conditional display formatting; this checkpoint does not substitute one for the other.
- [DevExpress GetDynamicArrayFormulaRange](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Cell.GetDynamicArrayFormulaRange), also checked against the installed 25.2 XML documentation and native spill-child test.
