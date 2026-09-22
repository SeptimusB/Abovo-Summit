# Save preparation and clean-close state — 22 September 2026

Test release: **2.62**. Status: **Ready to test**. User authority: request for a rebuild-needed flag, no integrity check on clean close, and summary HTML refresh after Save As. Changes are application-side only; supplied XLSBs and both repository masters remain unchanged. No persisted metadata, new worksheets, schema migration or VBA change is introduced.

Follow-up: **2.63** supersedes the every-save formula scan described below with a separate per-model formula-preflight revision. See `Save_Formula_Preflight_Cache_Trial_2026-09-22.md`; the remaining 2.62 state rules are retained.

## Behaviour

The state belongs to each `FileManager.ExcelModel`, not one application-wide Boolean. Saving one plan cannot certify another. It is deliberately not persisted: the first actual save of a newly opened model prepares its dependencies conservatively, but closing an untouched healthy model does not force preparation.

| State at save | Preparation |
| --- | --- |
| First save, structural/formula/name change, unknown import or failed preparation | Recursive `CalculateFullRebuild`, including deferred sheets |
| Value edit, paste or value Undo/Redo after a successful preparation | Recursive `CalculateFull`, without rebuilding dependencies |
| No changes since a successful complete calculation | Reuse the calculated state |

Normal Save already returns immediately when the model is clean. Save As still serializes a new file, but now reuses current calculation results. Structural operations flow through the existing dependency invalidation/bulk-mutation boundary. Native spreadsheet formula-capable edits, sheet/row/column/name changes, snapshot formula creation, Stress Test mode-name changes and unknown import paths explicitly request a rebuild. Typed editing requests one when it replaces or restores a formula; ordinary value edits do not. Future code that writes formulas/changes structure outside these services must call `RequireFullRebuild` before mutation and retain the existing dirty/recovery contract.

Calculation revisions prevent an earlier completion from clearing a newer change. An in-progress bulk command cannot be saved or accepted via clean close. Failed/cancelled saves restore the preceding dirty state and conservatively retain rebuild-required state. An explicit complete Recursive `CalcFile(3)` can satisfy subsequent save preparation when the revision is unchanged and no bulk command is active. A partial/deferred/sheet calculation cannot do so.

`CalculateAndValidateForClose` now skips calculation and Check Sheet inspection for a clean healthy model. Changed models retain the established validation, warnings and separate-copy protection. A previously detected check failure or recovery-required state is **not** forgotten merely because a copy was saved. If close validation has already prepared the results, the following save reuses them.

The existing XLSB formula preflight still runs for every actual save. The proven DevExpress long-argument formula guard has not been removed. If it normalizes a formula, it requests a rebuild before writing. The prior engine, calculation mode, deferred-sheet policy and worksheet protection are restored. This change does not eliminate file serialization or promise faster interactive edits/inserts.

Successful Save/Save As refreshes `FileInfo`, the registered FileInstance summary HTML and queued open group sidebars, after the new path is confirmed. Previous-access history remains unchanged. Cancellation/failure does not advertise an unsaved path. Native document-property changes made by saving itself do not request another rebuild. Deferred `ContentChanged` notifications are checked against the control's `Modified` save point: delayed load/save notifications must not dirty an untouched or just-saved model. Exact event delegates are retained and removed on disposal.

## Validation and boundaries

- Debug and Release compile from the active `C:/Repos/Abovo Summit` checkout; no running user Summit executable was replaced while locked.
- Real Blank and Demo private-copy tests pass in Release; populated AGL passes in an actual **32-bit Debug** process. The source SHA-256 checks pass. The tests instantiate the real FileInstance and inspect its native browser DOM before and after two Save As paths.
- These runs cover first rebuild, unchanged second Save As, typed rent edit, value Undo/Redo, ordinary Save, unchanged open/close, saved healthy close, metadata refresh, engine/mode/protection restoration, injected write failure and cancellation without `DocumentSaved`.
- After the typed rent edit, the ordinary-save caches agree with a separate explicit Recursive full rebuild across **11,900 cells** on Detailed Comp Income, Financial Position, Detailed Cashflow and Check Sheet, for all three workbooks. This is calculation-path equivalence, not financial approval. The populated Demo follow-up also explicitly asserts deferred-sheet-policy restoration.
- The 32-bit synthetic state suite covers two-model isolation, native formula edit events, in-progress bulk save/close rejection, explicit rebuild reuse, dirty close validation, known failing checks, recovery flags, cancelled/failed writes and stale revision tokens.
- The existing 32-bit synthetic formula-export tests still pass counts 29–60, nested calls/literals, protected sheets, arrays, global/local names, cancellation, failure rollback and unsupported long-call refusal.
- The edited Demo was opened read-only in macro-disabled Excel and saved only as a separate disposable copy. Native reopen preserves worksheet order, name count and **12,483 selected cells' formulas/caches**. All **337 VBA module identities and source hashes** agree between the original Demo, Summit save and Excel copy. Macros were not executed.
- **Not passed as an independent calculation oracle:** macro-disabled Excel full recalculation produces `#NAME?` errors (COM code `-2146826259`), with 561 differing cells on the first checked Detailed Comp Income sheet, in both the untouched Demo and the edited save. This probe cannot establish Excel/VBA recalculation parity under that policy. No automatic macro enabling, formula substitution or source-master edit was used to conceal the limitation. Client functional Excel/VBA testing remains required.

Initial exploratory tests found two genuine notification problems and are superseded by the final logs: automatic save metadata re-invalidated the rebuild flag, and delayed load `ContentChanged` re-dirtied a clean model. Both are corrected and exercised by real native-browser tests. A synthetic raw API write with history disabled does not reliably raise `ContentChanged`; service writers must continue to mark the model dirty explicitly.

## Timings

These are diagnostic, not controlled architecture benchmarks. On the populated AGL x86 run the first rebuild took 7.275 seconds, while the unchanged second Save As reports `mode=current, total=0 ms` for calculation. Ordinary edited save used `mode=full` (7.880 seconds in that run). The full pass is intentionally retained to prevent stale outputs. Formula preflight and writing still take time; no comparable reduction for those stages is claimed.

Trace lines identify `[Save Calculation Benchmark] ... mode=rebuild|full|current` and `[Close Validation] ... skipped=clean`, alongside existing XLSB save timings.

## Evidence and reproduction

- `Tools/Test-SavePreparation.ps1 -StateTracking -Architecture x86`: state machine regression.
- `Tools/Test-SavePreparation.ps1 -Architecture x86`: formula compatibility regression.
- `Tools/Test-SavePreparation.ps1 -Workbook <source>`: private model save, HTML, edit/history/cache comparison and failure tests. Add `-Configuration Debug -Architecture x86` for the delivered Debug process.
- `Tools/Test-SaveExcelRoundtrip.ps1` and `Tools/Inspect-SaveRoundtrip.ps1`: macro-disabled copy and native reopen. This run used the same saved workbook for both input reads; its comparison is not an independent financial oracle. The native Excel-copy comparison and VBA hashes establish preservation.
- `Tools/Test-SaveRecalculation.ps1`: separate diagnostic oracle; macro-disabled models requiring VBA/UDFs may fail, as recorded above.
- `Tools/Verify-VbaModuleHashes.py`: read-only hashes using dependencies outside the repository; no extracted VBA source/password is stored here.

Logs: `obj/save-flags-build-debug.log`, `obj/save-flags-build-release.log`, `obj/save-flags-state-final.log`, `obj/save-flags-compatibility-final.log`, `obj/save-flags-blank-final.log`, `obj/save-flags-demo-final.log`, `obj/save-flags-demo-policy-final.log`, `obj/save-flags-agl-debug-x86.log`, `obj/save-flags-excel-roundtrip.log`, `obj/save-flags-native-roundtrip.log`, `obj/save-flags-vba-demo.log`. The intentionally non-passing recalculation probes are `obj/save-flags-excel-demo.log` and `obj/save-flags-excel-demo-baseline.log`.

Demo saved result: `obj/SavePreparationTests/bc2fb01fc0134d25818c9c4765c871f1/model-save-second.xlsb`; Excel copy: `obj/SaveExcelRoundtrip/0360b0f9f3be471788c1ab1c8fb6a099/excel-roundtrip.xlsb`. AGL run: `obj/SavePreparationTests/f5797bc0b76343c8a2d9fe7419d7ba88/`.

Original repository SHA-256 values remain:

- Blank: `0B4C06800FE998E8733A5D8EB9CDEFA04F517750275BB0CFDD532928900F5B79`
- Demo: `1ED79726D8D1129C699242C9E8B3E988AC920EA7108D3C0B105BF9B0BC8A5E9C`

## Manual acceptance

Use a disposable populated file: save, edit a rent, save again, then Save As and inspect the summary filename/size. Open/close an unchanged file and close a just-saved file. Add Funding or Development records, save and confirm `mode=rebuild`; repeat Save As and confirm `mode=current`. Exercise an open Analyser/snapshot and native spreadsheet edit too. Check the result in normal trusted Excel/VBA and obtain the client's financial approval. This trial does not re-run every structural macro, every bespoke model or all physical UI/DPI scenarios.
