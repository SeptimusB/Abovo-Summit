# Stages 2k-2m: native recovery, Excel ownership and structural gates

Internal checkpoint only. Stage 2 continues; normal Debug/Release 2.98 and both repository masters remain unchanged. Normal model opening does not yet enable the new owner.

## Recovery implementation

- Existing recovery scheduling, notices, dirty-input eligibility and last-good replacement call a native snapshot branch when a native model is bound. Recovery does not acknowledge Save, replace the current owner, clear dirty state, or discard Undo/Redo.
- Excel exports a private same-format copy, opens that copy with macros/events/links/calculation disabled, and converts only that copy to XLSM. The live workbook never receives Save As. DevExpress exports XLSM and restores its temporary recovery properties in Finally.
- Current immutable history, exact input/calculation revisions, unchanged model XML, original opaque VBA and supported dynamic-array metadata are checked before accepting a candidate. The existing verified 117-byte XLSB metadata profile is the only binary-to-XML bridge; unknown profiles fail without replacing the last recovery.
- Native recovery candidates cannot be published as normal saved models. The store retains the existing `~name_recovery.xlsm` naming and original-filename guidance. The existing provenance-marker policy is applied to native recovery output.
- Normal save following recovery is tested separately: no recovery identity leaks into the normal file, and recalculation is exercised after reopening it.

## Validation

- Debug and Release isolated application builds pass. `EngineStage2k-*` contains recovery; `EngineStage2l-*` additionally contains Excel ownership protection; `EngineStage2m-*` includes structural gates. None is a versioned client release.
- Six generated combinations (XLSX/XLSM/XLSB, Excel/DevExpress): 162 assertions. Includes two changed-input recoveries, actual spill values, current history, unchanged original bytes, last-good retention after rejected XML changes, normal save/recalculation after recovery, dirty/Undo retention and refusal to publish a recovery candidate as the original. Nine tracked Excel owners exit, plus the two fixture-only XLSB converters.
- Existing bridge: 133 assertions; safety: 73; publication interruption/recovery: 109. Recovery entry-point assertions now verify current native values instead of expecting the former unsupported-operation refusal.
- Whole private Demo recovery/Save/Save As: 34 assertions per engine in `obj/EngineStage2k-Release-dit-save-*-stream.log`. Ownership-enabled repeats use the identical private Save/Save As operation to expose exceptions instead of blocking on wrapper error dialogs; logs are `obj/EngineStage2l-Release-dit-save-*-final.log`. The first ownership-enabled repeat reached a Save As error dialog and only that exact disposable DevExpress test process was stopped. No user Summit/Excel process was stopped.
- A preliminary fixture incorrectly used path-based `LoadDocument` to inspect an already leased saved workbook and ignored its Boolean failure. The fixture now opens explicit read-only streams, asserts successful loading, and uses a separate verification workbook for each output. The failed run remains evidence, not a pass.
- One preliminary synthetic run encountered a candidate sharing violation and refused the operation. Its source/last recovery were not overwritten; no retry suppression or broad exception swallowing was added.

## Dynamic-array distinction found by the stronger test

DevExpress 25.2.4's XLSB importer represents the tested Excel-created dynamic spill as a legacy array range. Re-export preserves its existing binary metadata; Excel recalculates the spill, and DevExpress recalculates the existing range. The same recovery converted through the reviewed XLSM bridge is recognised as dynamic by DevExpress.

This differs from creating a brand-new dynamic formula through DevExpress and directly exporting XLSB: the synthetic exported file lacks working array identity and neither engine updates its spilled values correctly. The XLSB fixture is therefore now created by converting a generated XLSX with isolated Excel, not by assuming DevExpress can author that XLSB feature. No production workbook formula or metadata was changed to hide this distinction. Variable-size spill expansion through the XLSB DevExpress fallback remains an explicit compatibility limitation to qualify; fixed-size cached-value parity is not proof of arbitrary spill growth.

`NativeSpillProbe.cs` is a local generated-file diagnostic, not a production rewrite. General DevExpress dynamic-array documentation does not establish XLSB feature parity: <https://docs.devexpress.com/OfficeFileAPI/14942/spreadsheet-document-api/formulas/array-formulas>.

## Excel process ownership

- Process-wide calculation, native writes and candidate export refuse an unexpected workbook in the owned Excel instance.
- Close discards only the tracked model/seed workbooks. It calls Quit only when the remaining collection is verified empty. If foreign workbooks remain, or the collection cannot be inspected, the instance is made visible/user-controlled and an explicit cleanup error is returned; foreign books are not saved, calculated or closed.
- Emergency hand-back deliberately does not change Trust Center or enable VBA/events. The message warns that the instance may remain in manual calculation with events disabled.
- Controlled ownership tests pass 16 assertions, including unreadable inventory and idempotent disposal. Real native fixtures exercise COM identity matching and exact PID/start-time cleanup; unrelated Excel activity is not treated as test ownership.

## Remaining before enabling normal opening

Funding/Development and all shared-rule inserts/deletes, Transactional DB synchronisation, explicit snapshot creation and generic bulk-mutation admission now refuse a bound native owner before changing the presentation workbook. Ordinary unbound DevExpress admission is unchanged. The generated gate fixture passes 21 assertions across both engines, including continued normal editing after refusal and unchanged history/dirty/source state. Additional legacy direct-mutation entry points must still be inventoried before opening is enabled.

Remaining native consumers/editors, unsupported-action inventory, recovered-format Save As policy, ordinary opening/settings and lifecycle/storage/security qualification. Native XML changes outside the admitted history path remain refused. Performance, especially Excel appearance binding and repeated save/reopen, still needs improvement. No Stage 2 completion or client round-trip acceptance is claimed here.
