# Development single-capture trial - 22 September 2026

## Scope

Test release 2.61 implements the user's approved experiment: avoid the second whole-workbook 3-D reference scan in Identified Development's first-column/remainder insertion sequence. It does not remove formula preservation, protection, mutation recovery, Transactional DB synchronisation or save-time calculation. No master/client workbook, XML, VBA, schema or normal edit-calculation policy is changed. Client-file recovery is a separate request; the exact file is awaiting confirmation and no client original has been repaired or overwritten.

## Algorithm and boundaries

- Capture once before the first insertion using the existing native formula parser, span-membership/bounds checks and array protections.
- Retain operation-local affected formula locations/references, selected worksheet identities/order and the first edit. No cross-command cache is introduced.
- After the first batch is inserted, references repaired and templates copied, advance the capture for the adjacent positive second batch. Re-read previously affected formulas/names and scan only the newly copied column on each of the seven sheets.
- An original reference wholly left of the first insertion cannot be newly affected by the adjacent second insertion. Copies are inspected separately because relative translation can change their references. Copied/overwritten areas replace their old inventory entries, not duplicate them.
- Reuse is gated to Identified Development, count greater than one, the reviewed staged full-column sequence, preceding-column templates and no extra copied columns. Funding, Multi-year Development, deletion, one-record insertion and other contracts retain full capture.
- Reject reuse before Apply, for non-adjacent/deletion edits or changed sheet order. Actual sheet identity is checked against the owning collection; DevExpress's Worksheet.Workbook facade is not necessarily reference-identical to its outer Workbook.
- A single final **tracked-reference** verification follows the source batches, before named-range resizing and TDB synchronisation can relocate cells. It checks the parsed references expected by the repair. This is not a whole-workbook financial integrity check and does not replace the existing copy guards, full save calculation, Check Sheet close validation or accountant review.

## Timing

Fresh private AGL x86 processes, ten Identified Development records; timings exclude load, save and UI rendering.

| Stage | Two-scan control | One-scan candidate |
| --- | ---: | ---: |
| First capture | 7.447 s | 7.443 s |
| Second capture / inventory advance | 7.679 s | 0.083 s |
| Final tracked-reference verification | Not present | <0.001 s |
| Source column shifts | 35.428 s | 35.266 s |
| TDB synchronisation | 48.678 s | 48.327 s |
| AddRecords return | 107.154 s | 98.330 s |

The control is the same-session intermediate trial executable in which the eligibility guard left the established two-scan path selected; its trace confirms both full captures. It is not a separately released version. The prior validated 2.60 run was 103.290 seconds. These are individual workstation observations, not a repeated statistical benchmark; the clearest isolated saving is approximately 7.6 seconds in the second capture. Most elapsed time still belongs to native source/TDB structural shifts. No Funding or Multi-year speed improvement is claimed.

Control: `obj/InsertionOptimisation/3551d4b54fc74df18760dedaf0581e17/operation-trace.log`.
Candidate: `obj/InsertionOptimisation/93223380c21b4291a93615ce147a4c33/operation-trace.log`.
Prior validated result: `obj/InsertionOptimisation/e6744d7a346a4150a24abf2375673eff/expanded.xlsb`.

## Completed validation

- Debug and Release build successfully.
- Native 3-D fixture: the existing insert/delete/copy/partial-span cases and 30 full-scan versus inventory comparisons pass, covering preceding/following templates, mixed references, whole rows, names, unaffected references and single-cell legacy arrays. Unapplied/non-adjacent/renamed-sheet reuse is rejected; final verification detects a changed reference.
- AGL saved-result comparison against validated 2.60 across all 283 sheets: 1,147,061 formula cells, zero formula/input/non-sensitive name/selected saved-output differences; worksheet order and protection agree. TDB formulas/constants are included, but TDB array classification, complete array-group geometry and exhaustive per-cell style parity are not asserted.
- Debug Demo Identified +3/delete: fourteen mirror dimensions, leading/sentinel rejection, protection/visibility/calculation state, all 1,679 global names and 106,044 linked-sheet formulas restored.
- Read-only macro/event-disabled Excel comparison agrees with the user's Excel ten-Development-column file on Detailed Comp Income, Financial Position, Cashflow and Check Sheet saved outputs. Excel SaveCopyAs then native reopen preserves worksheet order, name count and 12,483 selected cells/formulas/caches. No Excel macro execution or explicit Excel recalculation is claimed.
- All 335 AGL VBA module identities/source hashes agree between original, Summit result and Excel copy. No raw source/passwords are retained.
- Original AGL, Blank and Demo SHA-256 values are unchanged.

- Debug x86 Blank +2 with a deliberately unavailable TDB service reports recovery-required after mutation, retains dirty state, restores calculation/history/protection and releases execution/bulk guards. The failed model is never saved.
- Release x86 Blank +2 starting in Automatic mode restores that entry mode before save, uses exactly one full capture/one advance/one tracked-reference verification and one TDB synchronisation, saves successfully and retains entry settings.

Status: Ready to test. These checks establish the stated technical comparisons, not financial/client acceptance.

Evidence logs: `obj/optim261-build-{debug,release}.log`, `obj/optim261-3d.log`, `obj/optim261-development-final.log`, `obj/optim261-comparison.log`, `obj/optim261-development-regression.log`, `obj/optim261-excel-roundtrip.log`, `obj/optim261-native-roundtrip.log`, `obj/optim261-vba.log`, `obj/optim261-failure.log`, `obj/optim261-automatic.log`. Excel copy: `obj/SaveExcelRoundtrip/bde27f5bc5394a319faaf21dd5bc7c72/excel-roundtrip.xlsb`.

## Client acceptance

Repeat ten Identified Development records on a fresh copy through DIT, send the timing trace, then save/reopen in Excel/VBA and Summit. Financial/accountant and interactive VBA acceptance remain separate. Earlier snapshot/analyser lifecycle evidence is documented in the 2.60 scope trial; that full UI scenario has not been rerun for this narrowly scoped scan change. Existing workbook damage is not repaired automatically by opening or saving with 2.61.
