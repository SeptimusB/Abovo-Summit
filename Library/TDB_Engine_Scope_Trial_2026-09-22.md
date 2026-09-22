# TDB synchronisation engine-scope trial - 2.60

## Scope

This follows the verified AGL Funding and Development Test 2 results and the user's instruction to proceed with optimisation. No original client/master workbook is modified. No schema, capacity rows, hidden sheets, XML rules, formula algorithms or VBA changes are proposed. The source-column and mirror-row insertion/copy algorithms remain unchanged.

Identified Development and Multi-year Development insertion now finish their source update and invoke the existing TDB synchroniser while the structural operation still owns its temporary Recursive calculation-engine scope. TDB retains its own update, history, snapshot, analyser-disconnection and recovery guards. The outer `Finally` restores the caller's calculation engine/mode before interface dependency notifications. This removes the intermediate ChainBased rebuild immediately followed by another switch to Recursive. Funding, deletion and other rule families retain their existing sequencing. Funding was trialled but excluded from the final candidate because no repeatable total-time benefit was demonstrated.

`RunPostActions` accepts an already-completed synchronisation result and duration. It must not synchronise twice. Failed synchronisation still reports failure, marks the partially changed model recovery-required, and requires recovery Save As. Exceptions must still release the execution/bulk guards and restore settings. No failed native mutation is retried.

## Rejected native insertion alternatives

Three serial x86, fresh-process microbenchmarks used private copies of the same AGL input and expanded fourteen Development TDB mirrors by ten rows. These diagnostic copies were deliberately not source-synchronised and were NEVER saved. Manual calculation, Recursive engine and disabled workbook history were identical between probes.

| TDB-only variant | Elapsed |
| --- | ---: |
| Existing bounded used-column InsertCells | 44.602 s |
| Whole-row Rows.Insert fallback | 44.314 s |
| One multi-area union InsertCells, then fills | 44.705 s |

No meaningful gain was demonstrated, so neither alternative enters production. A separate synthetic workbook confirmed disjoint-area insertion semantics, but it is not a correctness proof for the financial model. The union shift alone took 33.569 s; filling brought the total to 44.705 s.

Local evidence: `obj/tdb260-bounded.log`, `obj/tdb260-rows.log`, `obj/tdb260-union.log`; private directories `8a2430fed50d40f5ab5538f1f20b866f`, `a339d86bfdbf437fb28c38a164f72889`, `a6cd2818e32b4eb2ad187ab4ba060fef` under `obj/InsertionOptimisation`.

## Measured timings

Input: `C:/Sandbox/Insert Comp/AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb`, SHA-256 `30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47`. The preserved 2.59 Release assembly is `obj/InsertionOptimisationBaseline259/Abovo-summit.exe`, SHA-256 `94335BDA3B1336A44339FB12ECDC26A526D71ABB919A80CFFBC1B865EBF77C5E`.

Benchmarks run serially in fresh x86 processes, ten inserted records, no analyser, and exclude opening/saving/rendering. Private-copy source hashes are checked. These are single paired trials, not guarantees for all machines/models.

Development: 2.59 AddRecords returned in 109.662 s; candidate in 103.290 s. Baseline's operation completed successfully, but a newly added post-timing fixture assertion incorrectly looked up a field as a property and failed before save. That assertion was corrected. The known validated 2.59 saved Development result remains the formula-comparison baseline; do not count the aborted harness as a full pass.

| Full operation | 2.59 baseline | Shared-engine trial | Decision |
| --- | ---: | ---: | --- |
| Identified Development +10 | 109.662 s | 103.290 s | Retain: 6.372 s / 5.8% reduction in this pair |
| Funding +10, first pair | 99.023 s | 122.112 s | Repeat: unchanged source-shift stage also slowed markedly |
| Funding +10, repeat candidate then baseline | 117.395 s | 118.215 s | Do not retain: no demonstrated total-time improvement |

The two later Funding runs opened the unchanged workbook in 15.614 / 15.693 s, versus approximately 12 s earlier. The changing unchanged-stage timings mean these are not evidence of a universal speed factor. Do not compare across those periods to claim a gain. No performance claim is made for Multi-year Development; it shares the same seven-sheet/fourteen-mirror lifecycle and is subject to functional regression below.

The isolated no-history Funding trial returned in 125.519 s (core trace 125.511 s), without a useful demonstrated gain. Its post-timing trace reader initially failed because it did not share the writer's open handle; fixed to read with FileShare.ReadWrite. No workbook was saved by that run. The later ordinary Funding repeat passed the new exact-once/ordering assertions and Save As. No production history behaviour was changed.

The Development candidate's TDB sync took 44.928 s with zero inner engine-restoration time; the outer restoration took 5.127 s. The prior operation had separate 4.618 s outer and 4.818 s TDB restorations. The retained change removes that redundant transition; it does not accelerate the fundamental native cell-shift algorithm or promise Excel-speed parity.

Full-run evidence: `obj/optim260-baseline-development.log`, `optim260-candidate-development.log`, `optim260-baseline-funding.log`, `optim260-candidate-funding.log`, `optim260-nohistory-funding.log`, `optim260-repeat-funding.log`, `optim260-repeat-baseline-funding.log`. Their private directories under `obj/InsertionOptimisation` are respectively `b8b1cbed945f45a486a9abf5c6b43758`, `e6744d7a346a4150a24abf2375673eff`, `2000b8ec9d8c4e2d8df2cddc5886a4e0`, `6fc44d21c92c45c2b374a3ca86123dda`, `ab64b60bb6184188a050fb45127f99d4`, `bd642b24570b41b79d6f8c891a4950f8`, `154ceca93294471cac839a8009dd57f8`.

## Validation

The final code narrows the shared-engine condition to the two Development families. Both final Debug and Release builds pass. The Development code path is the same as the timed candidate; the final regression fixtures use the narrowed binaries.

The full saved-result comparison with the verified 2.59 Development output passes across all 283 sheets: 1,147,061 formula cells, zero formula/input/non-sensitive name-definition/selected-output/checked-array-type differences, with worksheet order and protection agreeing. TDB formulas/constants are included; TDB array classification, exhaustive array-group geometry and exhaustive per-cell formatting are not asserted. JSON: `obj/InsertionComparison/4ed9a7a986cf49acadbf3d2faecbaab3/comparison.json`.

Excel read-only inspection, macros/events disabled and calculation manual, finds zero saved-value differences against the client's Excel Development result on Detailed Comp Inc - Trad View, Financial Position - Trad View, Cashflow detailed and Check Sheet. The existing B41 CONCATENATE compatibility rewrite is retained. This compares saved results; it does not execute VBA or independently recalculate in Excel.

The private Excel SaveCopyAs result reopens natively with worksheet order, name count and 12,483 selected cells/formulas/caches preserved. All 335 VBA module identities/source hashes match the original AGL in both the Summit result and Excel copy. Evidence: `obj/optim260-excel-development.log`, `optim260-roundtrip-development.log`, `optim260-vba-development.log`; Excel copy `obj/SaveExcelRoundtrip/8398d2cb66c14acdbc64d53b3314bd8c/excel-roundtrip.xlsb`.

Final-binary regressions:

- Debug Demo Identified Development +3/delete: all fourteen mirrors, first-ten/sentinel rejection, 1,679 captured global names and 106,044 linked-sheet formulas restored; protection/visibility/calculation state preserved.
- Release Blank Multi-year Development +3/delete: all fourteen mirrors, first-three/sentinel rejection, 1,679 names and 106,041 linked-sheet formulas restored; entry state preserved.
- Release Blank Funding +8/delete through the existing button event: all 32 source shifts, eleven mirrors, progress notification, ordinary/revolver boundaries and formula/name restoration pass. Funding retains its original sequencing.
- Release Blank native visible Analyser plus valid snapshot, Identified Development +2: stale binding disconnected, snapshot invalidated, analyser explicitly marked stale, entry engine/mode/history and nested guards restored, explicit refresh returns a current live binding. No original was saved.
- Debug x86 Blank Identified Development +1 with the sync service deliberately unavailable: failure returned, partially mutated model dirty/recovery-required, all execution/bulk/calculation/protection guards restored. Failed private model was not saved or retried.
- Release x86 Blank Multi-year +1 entered with Automatic calculation: entry mode/engine/history restored before Save As; exactly one TDB synchronisation; engine restoration precedes interface notification; production save succeeds and settings remain restored.

These correctness jobs ran concurrently and their durations are not speed benchmarks. In particular, the warm Analyser-plus-snapshot +2 run took about 178.8 s, of which 129.9 s was TDB synchronisation. This is a remaining expensive case, not a claimed improvement; isolate it in a future before/after trial before proposing changes to retained snapshot/comparison formulas.

Logs: `obj/optim260-identified-regression.log`, `optim260-multiyear-regression.log`, `optim260-funding-regression.log`, `optim260-analyser-snapshot.log`, `optim260-failure-regression.log`, `optim260-automatic-regression.log`, `optim260-build-debug.log`, `optim260-build-release.log`. All source wrappers verify unchanged hashes. Both repository master hashes remain unchanged.

## Delivery and remaining acceptance

Status: Ready to test, version 2.60, Debug and Release. No commit/push requested. The spreadsheet preservation checks use native XLSB/Excel paths rather than converting the authoritative model to another format. Existing 2.58 save/calculation/export safeguards remain in place.

Repeat Identified Development +10 on a fresh AGL working copy with the usual interface arrangement, save under a new name, and inspect Check Sheet/SOCI/financial position in Excel. Send the full insertion trace. Multi-year timings, extensively populated x86/x64 performance, interactive Excel/VBA actions and accountant financial acceptance remain manual. No Funding speed improvement, global performance guarantee or financial sign-off is claimed.

The new `StructuralEngineScopeFixture.cs` runs through `Tools/Test-ColumnFamily.ps1 -Fixture StructuralEngineScopeFixture.cs -Count 2`. `Tools/Test-InsertionOptimisation.ps1` adds test-only `tdb-bounded`, `tdb-rows`, `tdb-union`, `no-history`, `sync-failure` and `automatic` modes; none changes ordinary application settings or enables an experimental production algorithm.
