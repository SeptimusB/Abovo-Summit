# Fast input-save policy - 22 September 2026

Test release: **2.64 - Ready to test**. User approved deferring the extra whole-model calculation at Save, and explicitly required XML `IsCalculated` sheet/grid behavior to remain intact. This supersedes 2.63's ordinary-edit full-calculation-on-save policy, not its formula compatibility or structural safeguards.

## Policy and boundaries

- Ordinary value edits still go through ChangeManager, `CalculateWSs`, and the existing grid `UpdateCalcs`/`IsCalculated` logic. No change to that immediate sheet-level calculation contract.
- After formula/structure verification, a value-only Save writes inputs and the existing cached results without switching calculation engines or requesting another whole-model calculation. The remaining native XLSB serialization/compression cost is not removed.
- Saved inputs and up-to-date whole-model results are separate states. `ResultsPending` derives from the model's calculation revision; `DeferredSaveResultsPending` identifies a successful deferred save requiring a reader/load gate.
- Actual Save persists a Boolean standard custom document property, `Abovo.Summit.ResultsPending`. No worksheet, named range, business formula, hidden template, or VBA change is introduced for this state. The installed 25.2 API and native/Excel round-trip tests validate this use of [DocumentProperties.Custom](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.DocumentProperties.Custom).
- FileInstance and sidebar summaries explicitly indicate pending full results. Save As still updates the displayed path and metadata.
- The first unverified open, structural/formula/name/unknown changes, recovery, and failed/cancelled preparation remain conservative. A first checked save can still scan formulas and rebuild; do not present it as a fast value-only save.
- A healthy saved model can close without recalculation/check inspection, even if saved caches are pending. Existing known Check Sheet failures, recovery requirements, dirty state, and active bulk mutations do not take that shortcut.
- Failure/cancellation restores the previous pending marker/state and preserves dirty/rebuild safeguards. A newer mutation cannot be certified by an older calculation completion.

## Result consumers

Full-model open recognizes the persisted marker and calculates before exposing the model as loaded. Missing markers retain the prior open policy; present unknown/malformed markers are conservative. Successful full calculation clears in-memory pending state, refreshes registered views/events, but does not secretly save the workbook. The persisted marker clears on the next successful current-results Save. An unnecessary later refresh after a crash/reopen is preferable to treating saved stale caches as current.

Pending-result gates cover the primary analyser's Live/Comparison source preparation, snapshot creation, Outputs/Combined output navigation, normal CalcFile/dependency-sensitive calculation, DIT export preparation, FFR opening/tab selection/export, Stress Test opening, and opening the global spreadsheet. Frozen Snapshot display does not need live recalculation. Retained FFR/Stress Test windows also check pending state on activation through their shared history binding, without a redraw/calculation when already current. Bulk structural calculations do not inject a reader gate into a half-finished insertion.

These gates are synchronous, on demand, and use progress messaging. The calculation cost is deferred, not eliminated or moved to a worker thread. A first result-view operation after a fast save can therefore take longer. Automatic sidebar refresh shows pending status instead of silently doing a full calculation after Save.

Export dialogs already containing a staged selection retain their existing staging semantics; the new DIT gate applies when assembling fresh export candidates. Offline comparison participants already explicitly calculate before comparison. No general promise is made about external programs consuming XLSB cached values without calculation.

## Validation performed

- Debug and Release builds of `Abovo Business Suite.vbproj`, AnyCPU: pass. Debug is updated at `bin/Debug/Abovo-summit.exe`.
- Synthetic x86 state lifecycle: first rebuild, current reuse, deferred save, persisted marker, next-reader calculation, repeated-reader no-op, next-current-save marker clearing, retained standalone activation without a history event, bulk reader rejection/internal calculation, formula/name invalidation, model isolation, close/recovery guards, cancellation/write failure, stale completion: pass.
- Synthetic final Release XLSB compatibility suite: long CONCATENATE normalization, constants/formulas, legacy array geometry, global/local names, protection/history restoration, cancellation/failure rollback, unsupported call rejection: pass.
- Real x86 Demo Debug and AGL Release disposable-copy tests: typed Rent and Stock edits, value Undo/Redo, Save As presentation, deferred save, clean pending close, output-reader calculation, failure/cancellation, and actual FileManager reopen: pass. Independent full-rebuild comparisons cover **11,900 output cells per model** including statements, Check Sheet and Transactional DB.
- The real Stock XML's first calculated integer field increases by exactly the stock input delta. Both its workbook value and the typed `DataRows` value update immediately after `UpdateCalcs`, **before Save**. Pass on Demo and AGL.
- Native FFR/Stress Test history regression on Demo: non-modal opening, edits, native VGrid editor Undo/Redo, hidden/reopen refresh, retained tabs, no phantom posts, subscription disposal: pass.
- Final Release Analyser regression on Demo: chart/grid parity, drill, Live/Snapshot/Comparison, new-snapshot zero differences, Balance Sheet edits/Undo/Redo, expansion retention, structural deferred refresh, malformed numeric data and resizing: pass. Log: `obj/fast-save-analyser.log`.
- Private deferred Demo XLSB opened read-only in Excel with macros/events/links and recalculation disabled, then SaveCopyAs to another private path: pass. Native reimport preserves the **true pending-results property**, sheet order/name count, and **12,483 selected formulas/cached cells**. All **337 VBA module identities/source hashes** match the original Demo and both copies. VBA execution is not tested by a source-hash comparison.

The Excel comparison uses the pending file itself as the reference, deliberately testing preservation rather than asserting its pending caches are financially current. Financial freshness is separately tested by the Summit load/reader calculations against a native full-rebuild oracle. Macro-disabled full Excel recalculation remains unsuitable as the financial oracle for this workbook's VBA functions; trusted Excel/VBA and accountant acceptance remain manual.

## Observed timings (single trials, not controlled architecture comparisons)

| x86 workbook/action | Save total | Native write | Restore | Save-time full calculation |
| --- | ---: | ---: | ---: | --- |
| Demo, typed Rent | 8.683 s | 8.683 s | 0 ms | Deferred |
| Demo, value Undo/Redo | 9.443 s | 9.443 s | 0 ms | Deferred |
| Demo, typed Stock | 9.178 s | 9.178 s | 0 ms | Deferred |
| AGL, typed Rent | 14.584 s | 14.584 s | 0 ms | Deferred |
| AGL, value Undo/Redo | 14.310 s | 14.310 s | 0 ms | Deferred |
| AGL, typed Stock | 15.114 s | 15.114 s | 0 ms | Deferred |

Formula preflight is skipped on these already-verified value saves. Initial checked saves remain slower (Demo 20.546 s; AGL 33.153 s). AGL's later pending-results calculation takes about 12.6-13.0 s including engine restoration. A full-model reopen of a marked workbook similarly incurs the deferred calculation before use. These observations are not a controlled before/after comparison with the user's earlier 25.494-second trace.

## Evidence and source identity

Logs under ignored `obj`: `fast-save-state-final.log`, `fast-save-compatibility-final.log`, `fast-save-demo-final.log`, `fast-save-agl.log`, `fast-save-standalone.log`, `fast-save-excel.log`, `fast-save-excel-native.log`, `fast-save-vba.log`. Tools: `Test-SavePreparation.ps1`, `SavePreparationFixture.cs`, `Test-StandaloneHistory.ps1`, `Test-SaveExcelRoundtrip.ps1`, `Inspect-SaveRoundtrip.ps1`, `Verify-VbaModuleHashes.py`. An initial synthetic activation test needed its test-only history-manager setup corrected; the final run passes.

Originals are not saved by these tests; per-run SHA checks and fixture-input comparison confirm preservation:

- Blank: `0B4C06800FE998E8733A5D8EB9CDEFA04F517750275BB0CFDD532928900F5B79`
- Demo: `1ED79726D8D1129C699242C9E8B3E988AC920EA7108D3C0B105BF9B0BC8A5E9C`
- Current AGL source and untouched fixture input: `283E0CAF7863A938E1CC877ED7E21FE9100B94A1A2D0403C378569EE7B90ED90`. This differs from the earlier audit's historical AGL hash; the earlier name alone is not evidence of identical input bytes.

## Client acceptance

Use a disposable populated file. Establish the initial checked save; make a Stock input change; verify its calculated grid total updates immediately; Save and capture the trace. Expect `mode=deferred`, skipped formula scan, and zero engine-restore time. Repeat Save As and value Undo/Redo. Open/reactivate outputs and confirm calculation occurs before current figures are shown. Close a saved pending model and reopen it in Summit; compare its figures. Test Analyser Live/Snapshot/Differences, new snapshot and exports, and trusted Excel/VBA reopen. A structural insert/formula change must still require the protected full preparation path. External readers with macros/calculation disabled can see stale cached results: do not use such caches as a final financial report.
