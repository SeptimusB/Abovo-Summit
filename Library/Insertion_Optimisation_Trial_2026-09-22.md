# Funding / Development insertion optimisation - test 2.59

## Outcome and scope

The first measured optimisation pass reduces AGL ten-record insertion time by approximately 16% for Funding and 23% for Identified Development. This is an incremental improvement, not Excel-speed parity or a financial sign-off. Both Debug and Release are published; status: Ready to test.

No master/client source workbook, XML structure, insertion/copy sequence, protected leading-record boundary, Transactional DB mirror algorithm, snapshot policy or save-preparation rule was changed. The 2.58 CONCATENATE export guard and full pre-save recalculation remain in force. No commit or push was requested.

## Implementation

1. `WorkbookStructural3DReferences.Capture` builds the worksheet-order dictionary and selected-sheet prefix counts once per call instead of per candidate formula. Nothing is cached across commands: edits, changed names and worksheet order are read afresh for every capture, including Development's second batch.
2. A conservative sheet-qualifier screen replaces the expensive regular expression. It handles quoted names, escaped apostrophes, Unicode names and external qualifiers. This is only a candidate screen: DevExpress's public formula syntax tree still determines and rewrites the actual references. Partial-span rejection, invalid-reference checks and array safeguards remain. Cell formula text is retrieved once per capture visit.
3. Funding, Identified Development and Multi-year Development opt into the existing `UseRecursiveEngineForMutation` guard, for insertion and deletion. The guarded operation restores the previous engine and calculation mode in `Finally`; normal interactive calculation is unchanged. This avoids maintaining the chain through each intermediate linked-sheet shift. Transactional DB synchronisation and dependency invalidation still run normally, and no calculations/safety checks were removed.

The supported calculation-engine distinction is documented by [DevExpress](https://docs.devexpress.com/WindowsForms/114462/controls-and-libraries/spreadsheet/formulas/calculation-process). This trial does not introduce background workbook mutation, shared-workbook parallelism, capacity rows, metadata sheets or persistent formula indexes.

## Controlled timings

Input: `C:/Sandbox/Insert Comp/AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb`, SHA-256 `30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47`.

Each timed insertion used a fresh 32-bit process, a fresh private copy, Release code, ten records and no Analyser. Runs were serial, not concurrent. Timings are `AddRecords` through return; file open, Save As and DIT rendering are excluded. There was no debugger pause. These are one paired full-operation trial per family, not statistical guarantees across workbooks or machines.

| Operation | Before | Candidate | Reduction |
| --- | ---: | ---: | ---: |
| Funding +10 | 122.113 s | 102.140 s | 19.973 s / 16.4% |
| Identified Development +10 | 143.484 s | 110.840 s | 32.644 s / 22.8% |

The standalone Funding reference-capture check ran twice per process: 18.302 / 18.458 seconds before, 7.075 / 7.286 seconds after. All four runs returned the same 1,650-item reference/location fingerprint, `7348ED1E09EA150E38CAD1D77A21B5CB6536300DC1E4AC8A4A616C644D9267E7`.

Stage totals explain the improvement and its limits:

| Stage | Funding before / after | Development before / after |
| --- | ---: | ---: |
| 3-D capture | 18.126 / 7.404 s | 36.357 / 15.415 s |
| Linked-sheet shifts | 65.722 / 54.155 s | 53.737 / 36.485 s |
| Engine restoration before post-actions | 0 / 4.394 s | 0 / 4.537 s |
| Post-actions including TDB | 30.382 / 28.705 s | 50.756 / 52.581 s |

TDB itself is unchanged; its run-to-run difference is not claimed as a TDB optimisation. Native shifts remain the dominant next target. Do not compare the earlier debugger-paused Funding total, or treat the user's approximate 15-second Excel estimate as a controlled paired benchmark.

The Development timing was taken with the same candidate capture implementation and the rule's existing Recursive flag selected by the fixture, before the version increment. The identical flag is now enabled by the three production rule definitions in 2.59. Final 2.59 Debug/Release add/delete tests validate the shipped settings. Funding's after timing uses the final published 2.59 Release binary and normal production dispatch, without a fixture override.

## Correctness checks

- AGL scan: all 1,077,821 formula cells checked against the previous candidate screen. Both screens identify the same 61,448 candidates; none omitted.
- Synthetic 3-D fixture: insert/delete, mixed absolute/relative references, ordinary references, whole rows, defined names, single-cell legacy arrays, template copies, partial-span rejection, escaped/Unicode/external qualifiers and changed worksheet names between captures pass.
- Full native saved-result comparison across all 283 AGL sheets, including Transactional DB formulas/constants: **1,126,411 Funding formula cells** and **1,147,061 Development formula cells** match the corresponding pre-optimisation private result. Zero formula, input, non-sensitive name-definition, selected-output or checked array-type differences; worksheet order and protection flags agree.
- Array/non-array classification was checked outside Transactional DB. Full multi-cell array-group geometry and per-cell formatting were not exhaustively compared. Transactional DB's costly per-cell array classification was deliberately excluded, not silently assumed equivalent.
- Final Release Funding +8 / delete-last on a private Blank-master copy: all 32 shifts and 11 mirrors traced, names and linked-sheet formulas restored, first ten ordinary columns/revolvers protected, progress callback exercised, entry state restored.
- Final Debug Identified Development +3 / delete-last on a private Demo-master copy: two batches, seven sheets, fourteen mirrors, all 1,679 captured global names and 106,044 linked-sheet formulas restored; first ten and final template protected.
- Final Release Multi-year Development +3 / delete-last on a private Blank-master copy: seven sheets, fourteen mirrors, all 1,679 captured global names and 106,041 linked-sheet formulas restored; first three and final template protected.
- Those add/delete regressions ran concurrently and are correctness evidence only; their elapsed times are not comparative benchmarks.
- Independent Excel, macros/events disabled, sources read-only and calculation manual: both AGL results match the client's corresponding Excel-produced result on Detailed Comp Inc - Trad View, Financial Position - Trad View, Cashflow detailed and Check Sheet, with numeric tolerance 1e-7. B41 retains the 2.58-compatible nested CONCATENATE formula.
- Each private Excel SaveCopyAs result reopened natively with worksheet order, global name count and 12,483 selected cells/formulas/caches preserved. This check does not execute VBA or explicitly recalculate in Excel.
- All 335 VBA module identities and source hashes exactly match the AGL original across both Summit results and both Excel copies. No raw VBA or passwords were persisted.
- Original AGL and both repository master hashes remain unchanged. The reference-file wrappers also verify unchanged input hashes.

## Reproduction and local evidence

`Tools/Test-InsertionOptimisation.ps1` accepts `-Workbook`, `-Rule`, `-Count`, `-Architecture x86|x64` and `-Mode capture|scan|insert|recursive`. `recursive` is a test-only rule override retained for reproducing the trial. `-ApplicationAssembly` can select an explicitly preserved baseline assembly; the 2.58 baseline executable was copied to `obj/InsertionOptimisationBaseline258/Abovo-summit.exe`, SHA-256 `0E693C1C97659A4E3053815E902DD8D76F7D7DA71624C74941390929D2F74FE1`. It is diagnostic evidence, not a distribution executable. All workbook work occurs on uniquely named copies under `obj`.

Full-operation evidence under `obj/InsertionOptimisation/`:

- Development baseline: `6b1925bc7d3a49668b6fcb6ce4851011`
- Development candidate: `07070416adfe4b22948f047a40633139`
- Funding baseline: `10b240567191435bb5107a4f8d76bacd`
- Funding candidate: `af2e131b1fd94051b554afaa10e89802`

Each directory contains `operation-trace.log` and its private `expanded.xlsb`. Corresponding top-level logs are `obj/optim259-baseline-development.log`, `optim259-recursive-development.log`, `optim259-baseline-funding.log` and `optim259-candidate-funding.log`.

`Tools/Compare-InsertionOptimisation.ps1` runs the existing structural comparator in pairwise mode, includes Transactional DB formulas/constants, and fails on differences. Completed JSON reports: `obj/InsertionComparison/2e96878686a1453e904ddcb8b0d61fdf/comparison.json` (Development) and `94fa2db704e24cfc8c6a8bf2dbc5246a/comparison.json` (Funding).

Other evidence: `obj/optim259-screen-regression.log`, `optim259-3d.log`, `optim259-funding-regression.log`, `optim259-identified-regression.log`, `optim259-multiyear-regression.log`, `optim259-excel-funding.log`, `optim259-excel-development.log`, `optim259-roundtrip-funding.log`, `optim259-roundtrip-development.log`, `optim259-vba-preservation.log`, `optim259-build-debug.log`, `optim259-build-release.log`.

## Remaining acceptance / next performance work

Repeat the ten-record Funding and Development UI tests on fresh working copies with the client's usual Analyser/window arrangement. Supply the full trace and Save As results; inspect Check Sheet, SOCI and financial position in Excel, and exercise the relevant interactive VBA actions on another copy. Accountant acceptance remains separate from regression equality. Multi-year timing and x86/x64 comparative performance are not measured by this trial.

Further optimisation should target the measured native linked-column and TDB row shifts, without bypassing reference updates, mirror synchronisation, mutation recovery, snapshot invalidation or save-time calculation. Any batching/alternate-native-API proposal needs another exact before/after and Excel round-trip gate. The former reserved-capacity model remains withdrawn.
