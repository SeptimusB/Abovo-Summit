# Check Sheet calculation and display coherence - 2.98

24 September 2026. Scope: functional tests 80 and the related header/table refresh in 10/60. **Ready to test** with Jon. Client/financial acceptance remains pending.

## Behaviour

- Ordinary edits with Check Sheet hidden retain the existing calculation path. No new full calculation is added there.
- Showing Check Sheet after a changed model revision requests the existing recursive full calculation, then reads the override-aware soft balance result. Repeated reads of an accepted unchanged revision do not repeat that calculation.
- Initial selection, warm return to an existing document and window visibility share a coalesced request. It runs after tab/editor initialisation rather than against a provisional revision.
- Successful edit, batch, Undo and Redo completion queues a check only if Check Sheet is visible. The request follows history presentation callbacks, so a pre-commit or pre-refresh revision is not certified.
- Accepted results publish warning state and a separate values notification. The latter fires even if the warning stays red or stays clear, keeping figures and colours current.
- The mapped grid refresh preserves selection, focus, scroll and widths. A pending Yes/No editor is neither committed nor discarded by the notification. Superseded results are rejected; native handle recreation and model disposal are guarded.

## Boundaries retained

Check Sheet remains a soft balance warning. Accepted overrides, normal Save/Save As availability, default recovery continuation and Save-only persisted warning semantics are unchanged. A Check Sheet pass does not certify formula/structural integrity or clear pending/rebuild/preflight flags. No workbook formulas, names, source fills, protection rules, VBA or master files are changed. No spreadsheet-engine replacement is introduced.

## Validation record

Both staged Debug and Release builds compile. Quiet-diagnostics regression passes for both; existing opt-in Check Sheet trial timings remain opt-in. Native AGL fixture: `Tools/CheckSheetCoherence298.cs`, always using a private file and a nonactivating host.

The initial regression exposed duplicate first-display calculation and a post-Undo presentation revision change. Those findings were addressed by coalescing reads after native activation and after the complete history/refresh sequence; revision guards were not weakened.

Test-environment finding: long native runs at the actual 5120 x 2089 desktop size intermittently failed to allocate a DevExpress drawing buffer. The original CodeDom x86 fixture lacked LargeAddressAware, whereas both actual Summit executables already have it. The regression runner now matches the application's address-space setting, retaining a verified 32-bit process. No production platform/memory setting, window geometry or forced garbage-collection policy is changed. The earlier failures remain in the evidence; a nonmatching fixture failure is not proof of an equivalent production failure. Neither final matching run repeated it; this is not a claim to have repaired a production graphics leak.

The legacy Framework compiler also accepted `anycpu32bitpreferred` without producing the matching PE flag, so the opt-in runner uses the installed VS2022 Roslyn compiler and fails before launch unless its PE flag matches. Existing fixture defaults remain unchanged. The mapped colour probe reads the workbook-derived projection: row 23 is red when failed, green after Undo and red again on Redo. It does not assume `Cell.Font.Color` is an invariant unconditioned base colour.

Independent final source review found no new blocking issue after the activation and revision-order changes. Runtime acceptance remains separate from this review.

The repeatable native command uses `Tools/Run-NativeRegression.ps1 -Fixture Tools/CheckSheetCoherence298.cs -Configuration <build-folder> -UseApplicationConfig -MatchApplicationAddressSpace -Workbook <AGL-source-path>`. The runner hashes the original before/after and the fixture edits a private copy only. The address-space switch is opt-in; it verifies PE parity and the fixture separately asserts a 32-bit process. The original hard-x86 runner default is retained for other fixtures.

Reference source: `C:/Sandbox/Insert Comp/AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb`; SHA256 `30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47`.

Final matching-address-space runs: **124/124 assertions pass per build**, exit 0 and empty error logs, original source hash unchanged. Debug evidence: `obj/ClientReportTests/bcae3fed12754d37af9469ef3b1afaa4`; Release: `obj/ClientReportTests/48bc98eaa88e4bca85ccbc99649893d7`. All eleven full checks per build ran at the native 5120 x 2089 maximised host size with no suppressed painting or forced garbage collection and no repeated native allocation failure.

Coverage includes the actual Funding G82 2700 -> 3000 edit, precise Check Sheet failures 23/33/39 and no stale Cashflow failure 37, public heading, all 57 mapped rows' A:F text, Undo/Redo, warm return, unchanged-revision no-repeat calculation, unchanged-warning value publication, red/green appearance, pending override preservation, cancelled-editor refresh, selection/viewport/width preservation, wrong-revision/reentrant/closing publication rejection, isolated failing subscribers, literal Yes/blank/No overrides and their Undo/Redo, calculation-mode/deferral restoration, pending/rebuild flags and unsaved warning persistence.

After Jon confirmed Debug closed, the same source was built into both normal output folders. Both builds pass; the runtime property (not the unchanged assembly file-version label) verifies 2.98. General diagnostics are disabled in both normal executables; explicit Check Sheet trial timings still require opt-in. Git whitespace checks pass. No commit or push was requested for this focused repair.

| Installed build | Executable | SHA256 |
| --- | --- | --- |
| Debug | `C:/Repos/Abovo Summit/bin/Debug/Abovo-summit.exe` | `921B4891473C5E7FB0E37FFE22933C340E1AF18673C830DF71AA6B07A1ECF8C3` |
| Release | `C:/Repos/Abovo Summit/bin/Release/Abovo-summit.exe` | `DA7CFD8056CCC72C62D9FF833B387DC7520BAFD382502BD0829AD82083358C86` |

The native runs used separate candidate folders to leave Jon's former Debug 2.97 available during testing. Tested candidate SHA256: Debug `F6CFABCDEF9EE60380A83BC23A42206FF981C858C84B13405DDD42005DD7D148`; Release `6A45B84EB559C6DB6D2C96733C71E7E01642E9B1C9BF3F02E0B98E4F0B6CE4BB`. Installation builds use the same unchanged production source; workbook regression was not repeated against the newly emitted normal-path binaries.

## User acceptance and other reports

Retest 80 using Funding G82 -> Check Sheet -> History Undo/Redo -> away/back. The visible checks and company header should agree after the check finishes; the temporary Transactional DB Cashflow failure should not remain. Also retest watcher refresh while the table remains visible and Yes/No override entry/clear/Undo.

Performance boundary: an initial return to Funding may populate/rebind editors and advance the existing calculation revision without a newly committed user edit. Returning to Check Sheet then conservatively checks again. Truly unchanged-revision repeat navigation is the no-extra-full-calculation guarantee; this release does not redefine dirty/revision semantics to suppress that safety check.

Manual acceptance still includes a disposable Save As -> Excel/VBA reopen/recalculate -> Summit reopen check. The native regression establishes this reported application's input/history/display lifecycle, not independent financial parity for every formula, bespoke workbook or client machine.

Jon has requested **Send to Alex Test Response 3** for **1, 2, 3 and 5**. Test **4 fails in the reported mapped-table zoom/horizontal-scroll scope**; that work is recorded but deliberately not included in this focused calculation repair. No client green/Agreed status is inferred. The numbered ledger keeps its stable IDs; the illustrated Word review is unchanged.
