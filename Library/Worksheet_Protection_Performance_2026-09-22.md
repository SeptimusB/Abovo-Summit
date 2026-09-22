# Worksheet protection performance - test 2.57

## Evidence and authorised scope

During the user's Funding insert, the paused Visual Studio main-thread stack was inside DevExpress worksheet password verification, reached from WSSecurity.UNProtectWS in the deferred column-copy step. It was performing SHA512 hashing, not reporting a circular reference. The ContextSwitchDeadlock MDA describes a long non-pumping STA operation; that observation alone did not establish a permanent deadlock. The user confirmed that worksheet-password security need not be strong.

WSSecurity now uses DevExpress 25.2's UseStrongPasswordVerifier=False only while reapplying worksheet protection, restoring the previous workbook option in Finally. It retains the runtime password, supplied permissions and cell locks. Already-protected sheets still short-circuit as before, and callers still decide whether an originally unprotected sheet needs protection. This is legacy Excel-compatible worksheet edit protection, NOT file encryption or an authentication boundary. No password, raw VBA source, load-time conversion, workbook schema change or global spin-count reduction was introduced.

This policy applies to callers of the shared ProtectWS service, not just Funding. Existing strongly protected sheets are still verified normally the first time they are unlocked. It takes effect when an operation legitimately re-protects them, and persists only if the user saves that working file. Source masters and client originals are not rewritten.

## Instrumentation and performance limits

Funding and Development insertion traces now separately measure unprotect and protect, with phase=shift/copy plus worksheet and batch context. Other existing phases remain unchanged. Protection inside the 3-D apply helper remains included in apply3D; these scopes are not nested in the aggregate. No passwords or cell contents are logged, no DoEvents was added, and workbook operations were not moved to a worker thread.

An isolated eight-cycle comparison on the same synthetic worksheet measured strong protection/unprotection at 1,015 ms and the editing policy at less than one millisecond (whole-millisecond Stopwatch reporting, not literally zero work). This is NOT a Stori insertion speedup claim. The Funding correctness fixture's +8 insertion still took 61,090 ms: column shifts 33,388 ms, 3-D capture 8,280 ms and post-actions 16,561 ms dominated. The fixture prepares an Excel reference by unprotecting/re-protecting its private copy before the insertion, so this does not measure cold unlock of existing strong protection. A long operation can still trigger the debugger MDA; it has not been disabled or hidden.

## Verification

- Debug and Release compile and publish to the active checkout's bin directories; test version 2.57.
- Tools/Test-WorksheetProtection.ps1 tests the actual production protection helper using a generated in-memory credential, never the business-plan password. Correct-password unlock, incorrect-password rejection, explicit permissions, cell locks, and restoration of both True/False verifier settings and spin count pass. An injected failure also restores options.
- Native XLSB save/reopen -> hidden macro-disabled Microsoft Excel read-only open -> edit private in-memory input, re-protect and SaveCopyAs -> native reopen passes. Formulas, named-range geometry, number format, cell locks, input edit and protected/unprotected worksheet states are retained.
- Tools/Test-StructuralInsertBenchmark.ps1 passes success/failure/interruption, repeated-disposal and listener-failure isolation checks.
- Funding +8 in Release passes the actual event-service/progress route, all 32 traced shifts and eleven mirror resizes. Save As, protected-loan/revolver deletion guards and deleting the new records pass; all original global names and linked-sheet formulas return to baseline, with protection/visibility/calculation state preserved.
- Independent macro-disabled Excel grouped insertion matches all 104,014 linked-sheet formula cells in the Summit Funding result. Excel save-copy/reopen matches those formulas and all global names. Full rebuild reports no circular-reference address. All 337 VBA module identities/source hashes match across the source, Summit result and Excel result; macros were not executed.
- Identified Development +3 in Debug passes both batches, all seven linked sheets, fourteen mirrors, Save As, leading/sentinel deletion guards, deletion restoration of all 1,679 global names and 106,041 linked-sheet formulas, and entry protection/visibility/calculation state. Multi-year was validated in 2.56 but was not rerun for this verifier-only change. No Development Excel comparison was repeated in this turn.

Funding evidence: obj/FundingStructureTests/776c8a8064274420a91765d63a2dbfff and obj/structural257-funding*.log. Development evidence: obj/ColumnFamilyTests/b32b08a4b33b4c1ba5ab21f6a78b2811 and obj/structural257-development-id.log. Development and the Funding Excel comparison overlapped; their timings are correctness evidence only, not a performance baseline. The earlier Funding insertion and synthetic protection timing ran without those concurrent checks.

Disposable synthetic evidence: obj/WorksheetProtectionTests/d9bc7f141c964bf09f624377fd8051c9. No customer workbook was used by that fixture. Blank and Demo master hashes still match their recorded baselines. Financial validation, real Stori performance, interactive VBA and visible progress acceptance remain client tests, not inferred from these structural checks.

## Next client run

Restart the refreshed Debug executable (2.57), repeat the same Funding count on a fresh Stori working copy, and retain the complete Structure Insert Benchmark / TDB / Population trace. Run Development separately with its type and count noted. Avoid debugging pauses during a timing run. Keep original files untouched and Save As distinct results for the subsequent Excel comparison. If the MDA recurs, capture the new last stage and paused stack; do not assume it is still password hashing.
