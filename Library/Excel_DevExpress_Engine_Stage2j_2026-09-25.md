# Stage 2j: ongoing native saves and model-owned close

Internal checkpoint; Stage 2 continues. Normal Debug/Release 2.98 are not replaced and normal model opening still does not select a native owner.

## Implemented

- Normal `ExcelModel.SaveFile` and `SaveFileAsTo`, plus the same-format Save As dialog branch, use the selected native owner when one is bound. The ordinary production path is unchanged. A unique private folder beside the destination retains verified candidates; publication retains its original backup/journal. Staging originally used AppData, which the test automation process could not write; no ACL/security policy was relaxed to address that environment restriction.
- Save captures exact existing history, verifies current XML and native values, publishes, reopens the same engine/security policy, calculates and only then acknowledges the exact user/calculation revision. Existing Undo/Redo survives both Save and Save As. A presentation subscriber failure cannot misreport an already successful file write as failed.
- Automatic preference is resolved once at opening. Subsequent save/reopen is pinned to that selected engine; it cannot silently switch calculation owners.
- Reentrant close during an admitted operation is refused before the model slot is removed. Normal close owns native cleanup. A closed native facade never falls through to stale presentation workbook values.
- Changed, unrouted non-history XML is refused before publication. Unsupported format conversion and overwriting a different existing destination are refused. Publication failure keeps dirty state and reports the retained verified latest-edits candidate; a failed resumed session cannot show obsolete caches.
- Native process identity is diagnostic PID plus UTC start time, not authority to terminate a process. New lifecycle fixtures check only their own owners, allowing other user Excel activity.

## Evidence

- Debug and Release isolated builds: `bin/EngineStage2j-Debug`, `bin/EngineStage2j-Release`.
- Generated native model-save tests: 80 assertions, XLSX/XLSB, DevExpress and Automatic selecting Excel; six owned Excel sessions exit. Includes collision/cancellation, changed XML, reentrant close and denied-rename recovery containing the latest edited value.
- Whole private Demo: 27 assertions per engine, application Save/Save As operations, real Covenant/Funding editor posts, Undo/Redo, independently inspected saved values and model-owned close. DevExpress exercises the public Save wrapper; Excel's final fixture calls the identical private save operation to expose exceptions rather than block unattended on the wrapper's error popup. Final logs: `obj/EngineStage2j-Debug-dit-save-dx.log` and `obj/EngineStage2j-Debug-dit-save-excel-final.log`.
- Release bridge 133, range 32, expression 32, safety 73, publication recovery 109. Expression rerun with exact owner tracking passed; an earlier strict all-process inventory check failed and was not treated as a pass.
- Input masters and normal binaries remain unchanged; all mutating tests use generated/disposable copies. Ordinary input/save regression: 48 checks, including independent reopen. Broader native DIT regression: 550 checks (`obj/ClientReportTests/777abebc3d094ca5a329b2e9938f5773`, `obj/EngineStage2j-Production-DIT-regression.log`). Existing WebView2 shutdown diagnostic 1412 remains with exit 0. Neither replaces remaining client round-trip acceptance.
- One unattended error-dialog fixture was stopped by its exact test executable/PID/start time. Its orphaned Excel owner was subsequently closed via COM only after verifying PID/start time, the disposable workbook path and all 80 seed capability-test cells. No user Excel workbook/process was closed.

## Continue, do not release yet

Native recovery XLSM conversion, current non-history model XML, remaining direct read/editor consumers, structural gates, actual opening/options, unexpected foreign workbooks in an owned Excel process, permanent native hangs, storage/security qualification and manual Excel/VBA round trips remain. Reopening/recalculating each save is conservative and still needs performance measurement; no speed improvement is claimed for this checkpoint.
