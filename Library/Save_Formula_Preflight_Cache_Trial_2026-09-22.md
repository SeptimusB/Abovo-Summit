# Formula-preflight state after ordinary edits - 22 September 2026

Test release: **2.63**. Status: **Ready to test**. This follow-up addresses the 2.62 report of a 7.659-second formula scan with zero rewrites after one ordinary Stock Assumptions edit. It supersedes only the every-save formula-preflight policy in `Save_Rebuild_State_Trial_2026-09-22.md`; the calculation, clean-close, failure/recovery and Save As HTML rules remain.

## Policy

Each open `ExcelModel` now has a separate formula/structure revision and successfully checked revision, exposed as `NeedsFormulaPreflight`. This is not the ordinary dirty flag and not an application-wide flag shared between plans. A complete calculation does not certify export compatibility.

- The first actual save in each newly opened model still checks all formulas/names. External Excel edits cannot be presumed safe from the filename or cached calculated values.
- After a successful check/save, typed value-only DIT edits, value Undo/Redo and unchanged Save As reuse the checked formula state. Their formula-preflight trace reports `skipped=True`, normally `0 ms`.
- Existing structural/bulk boundaries, formula replacement/restoration, unknown imports and native spreadsheet formula-capable events call `RequireFullRebuild`, which now also invalidates the formula-check state. Native name addition is included alongside name editing/deletion. Native spreadsheet edits remain conservative even when a particular edit only changes a value.
- The preflight flag clears only after successful saving of the same formula revision, outside an active bulk operation. Calculation completion alone cannot clear it. Failure or cancellation conservatively invalidates preparation. Other direct compatibility-guard callers still default to checking.
- New formula-writing API paths must use the established invalidation/bulk boundary. DevExpress API calls such as `DefinedNames.Add` do not necessarily raise the interactive name-added event; the event is not a substitute for service-side invalidation.

No hidden worksheet, persisted certificate, load-time migration, source-file rewrite or VBA change is introduced. The existing long-argument CONCATENATE normalization, refusal of unsupported long calls, protected/array/name handling and rollback remain in place. Calculation engine, mode, deferred-sheet policy and protection are restored.

## Calculation experiment - deliberately not adopted

An isolated probe compared Recursive `Calculate()`, `CalculateFull()` and a subsequent `CalculateFullRebuild()` after actual typed Stock/Rent edits with active worksheet calculation. With Stock and Transactional DB both registered, the incremental candidate left **Check Sheet B37 and B38 at 1 instead of 0**. The retained full-result pass agreed with the full-rebuild oracle in all three probes. Candidate and full timings were similar, around five seconds in this diagnostic run.

Therefore an ordinary changed workbook still receives a complete result calculation before serialization; it does not receive a dependency rebuild unless required. Removing this calculation has not been demonstrated safe. `-IncrementalProbe` intentionally exits nonzero when the candidate differs. Its non-passing result is evidence against that candidate, not a passed regression or a production-path failure.

## Additional timings

`[XLSB Save Benchmark]` now separates total elapsed time, the native save action and final engine/policy restoration. The existing formula-preflight and preparation/calculation lines remain. `nativeSaveDialog=True` means save-action timing can include time the user spends in the Save As dialog. These diagnostics are not controlled architecture comparisons; some test runs overlap.

The passing Debug Demo run recorded the stock-value save with zero formula-preflight time, calculation 4.338 seconds, save action 7.504 seconds, restoration 2.386 seconds, total 14.393 seconds. The larger AGL exploratory run similarly skipped preflight but still spent substantial time calculating, writing and restoring its engine. This change removes the unnecessary scan; it does not promise an instantaneous save or eliminate serialization costs.

Final strengthened x86 AGL run: stock-value save preflight **0 ms**, full result calculation **6.633 s**, save action **11.868 s**, restoration **5.059 s**, total **23.641 s**. Final Debug x86 Demo stock-save total was **14.310 s**. These are observed totals, not a controlled percentage improvement over the user's earlier measurement.

## Evidence and limitations

- Debug and Release builds pass from `C:/Repos/Abovo Summit`; application version is 2.63.
- The 32-bit state fixture passes per-model isolation, first-save verification, value edits, formula/name invalidation, unsafe-name refusal after a prior successful save, active-bulk rejection, known check failures, recovery rules, failure/cancellation and stale revision handling. The name-added unit case explicitly invokes the retained event delegate; it is not a physical Name Manager UI test.
- The 32-bit compatibility fixture passes the existing long-formula, literal, array, global/local name, protected-sheet and rollback tests.
- Final Debug x86 Demo and Release x86 AGL private-copy tests both pass actual Save/Save As, native summary browser DOM refresh, typed rent/stock edits, Undo/Redo, formula-scan reuse, clean-close behavior, failure/cancellation and state restoration. Both rent and stock save caches agree with separate full rebuilds; the final persisted save/reopen preserves all 11,900 checked output cells in each workbook.
- The saved edited Demo was opened read-only in Excel with macros/events/links and automatic calculation disabled, saved only to another disposable copy, and reopened with DevExpress. Worksheet order, named-range count and 12,483 selected cells' formulas/caches are preserved. All 337 VBA module identities/source hashes agree with the original Demo. No VBA ran. The two Excel input reads use the same saved file; they are not an independent financial oracle.
- Earlier populated x86 fixture attempts successfully saved but failed to load another full edited workbook alongside the live model. The file reopens successfully in a fresh native process. The strengthened fixture checks every `LoadDocument` result, compares live saved caches with a separate full rebuild, then verifies persisted caches after closing the live model. A further fixture correction detaches expected text/numeric values from DevExpress shared-string storage before disposal; retaining raw CellValue text references caused a fixture-only NullReferenceException after close. Final runs pass both stages. No production workaround or forced garbage collection was added to Summit for these fixture issues.
- Full independent macro-disabled Excel recalculation is not claimed: the 2.62 trial found the same UDF-related NAME errors on the untouched Demo baseline. Interactive Excel/VBA and client financial acceptance remain necessary.

Reproduction: `Tools/Test-SavePreparation.ps1` with `-StateTracking -Architecture x86`, no mode switch for formula compatibility, or `-Workbook <source> -Architecture x86` for private-copy model tests. `-Configuration Debug` exercises the Debug binary. `-IncrementalProbe -Workbook <source>` is the intentionally rejected calculation candidate. The real-model fixture uses a dummy active object with genuine worksheet registration/calculation, not a physical DIT keyboard interaction.

Logs: `obj/save-preflight-build-debug.log`, `obj/save-preflight-build-release.log`, `obj/save-preflight-state.log`, `obj/save-preflight-compatibility.log`, `obj/save-preflight-demo-final.log`, `obj/save-preflight-probe-final.log`, `obj/save-preflight-excel-roundtrip.log`, `obj/save-preflight-native-roundtrip.log`, `obj/save-preflight-vba-demo.log`. Intermediate non-passing populated fixture logs: `obj/save-preflight-agl-final.log` and `obj/save-preflight-agl-final2.log`; fresh-file reload: `obj/save-preflight-agl-fresh-reload.log`.

Final strengthened fixture logs: `obj/save-preflight-demo-verified.log`, `obj/save-preflight-agl-verified.log`, `obj/save-preflight-state-verified.log`, `obj/save-preflight-compatibility-verified.log`. The intermediate post-disposal fixture failures are retained in `obj/save-preflight-demo-final2.log` and `obj/save-preflight-agl-final3.log`; these are superseded, not passing runs. Final private output directories: Demo `obj/SavePreparationTests/5829d0d1d6514363a3b15d0f0f27b1f9/`, AGL `obj/SavePreparationTests/9e6e5027bb63407097c8cfb25c636033/`.

Excel preservation evidence uses `obj/SavePreparationTests/e5ed10a8f6a24e2a854f2fe40ec4b415/model-save-second.xlsb` and `obj/SaveExcelRoundtrip/993fe752793b43e2b21180526b4ba01b/excel-roundtrip.xlsb`.

Source SHA-256 values remain unchanged:

- Blank: `0B4C06800FE998E8733A5D8EB9CDEFA04F517750275BB0CFDD532928900F5B79`
- Demo: `1ED79726D8D1129C699242C9E8B3E988AC920EA7108D3C0B105BF9B0BC8A5E9C`
- AGL 26.0005 original: `30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47`

## Manual acceptance

On a disposable populated model, save once to establish the checked state, make one Stock Assumptions value edit, save again, and confirm `formulaPreflight=0 ms, rewritten=0, skipped=True`. Repeat an ordinary value Undo/Redo and Save As. Then perform a structural insertion or formula change and confirm preflight/rebuild are required again. Reopen starts conservatively unverified. Check outputs in normal trusted Excel/VBA, and send the new total/save-action/restore traces for any remaining delay.

## Follow-up save-stage investigation (application remains 2.63)

The user's subsequent populated trace confirms the cache works: formula preflight 0 ms, full result calculation 7,142 ms, native save action 13,347 ms, restoration 4,766 ms, total 25,494 ms. The user also reports that the Excel opening VBA forces CalculateFullRebuild. That report motivates a possible fast-input-save policy; it is not authority to silently treat stale cached outputs as current, and no production calculation policy was changed in this investigation.

Read-only review of the installed DevExpress 25.2 source confirms that the XLSB exporter serializes and compresses workbook parts. Its reviewed public XLSB export options do not expose a compression-level shortcut. Engine restoration enables the dependency-chain engine again. Save preparation already uses Manual calculation mode; the documented recalculate-before-saving setting is not a reason to remove the existing correctness calculation. Public reference: [DevExpress calculation process](https://docs.devexpress.com/OfficeFileAPI/400926/spreadsheet-document-api/formulas/calculation-process). No vendor source was copied into the repository or modified.

Added diagnostic-only `-SaveStageProbe` alternatives to `Tools/Test-SavePreparation.ps1`, exercising a private Demo stock-value edit after identical warm preparation. These are single fresh x86 trials, not controlled performance conclusions:

| Candidate | Prepare ms | Export ms | Disk copy ms | Restore ms | Measured subtotal ms |
| --- | ---: | ---: | ---: | ---: | ---: |
| Keep ChainBased and custom service | 9,449 | 10,703 | included | 0 | 20,152 |
| Keep ChainBased, temporarily remove skip service | 26,611 | 9,476 | included | 0 | 36,087 |
| Recursive, export to memory then write buffer | 4,437 | 8,246 | 3 | 3,030 | 15,716 |

All three pass comparison with a Recursive full-rebuild oracle and native saved-cache reopen across the four statement/check sheets plus Transactional DB. These are not Excel/VBA financial acceptance tests. The measurement excludes opening, initial formula normalization/full preparation, diagnostic value capture, oracle calculation and reload. The memory-buffer case separates serialization from disk copy: the buffer takes 8.246 seconds to create, versus 3 ms to copy to disk. It offers no demonstrated improvement over the existing roughly 14-second Demo save. The other alternatives are slower, so none was promoted to production.

Logs: `obj/save-stage-chain-demo.log`, `obj/save-stage-chain-native-demo.log`, `obj/save-stage-buffer-demo.log`. The harness temporarily restores the original engine/mode/service and verifies source hashes. No original XLSB, running Summit session or application binary was changed; application version stays 2.63.

A further fast-input-save policy would require a distinct results-pending state that survives saving, an explicit Summit-open/output-consumer calculation gate, standard Excel-compatible recalculate-on-open behavior, and defined handling for close validation and externally read cached figures. Excel VBA on opening is helpful when enabled, but does not run in Summit or macro-disabled Excel. Comparison currently explicitly recalculates its participants; other consumers must be audited too. This behavior change is proposed, not implemented or validated. Structural/formula compatibility safeguards must remain.
