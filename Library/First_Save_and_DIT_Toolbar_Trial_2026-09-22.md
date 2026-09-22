# First value-save and DIT toolbar trial - 22 September 2026

Test release **2.65 - Ready to test**. User approved correcting 2.64's first-save exception and supplied an explicit DIT WindowsUI button order. The later label request supersedes Compact/Restore with Maximise working area/Restore sidebars.

## First-save distinction

`NeedsFullRebuild` remains true for a newly opened, unverified dependency graph. A separate, initially false `NeedsSaveRebuild` records an actual in-session formula/structural/unknown mutation. `RequireFullRebuild` sets both flags; a successful current-generation full rebuild clears both. Ordinary typed value edits do not set the structural save flag.

The first value-only save may therefore defer the initial full rebuild while retaining the full XLSB formula compatibility scan. It saves the existing `Abovo.Summit.ResultsPending` custom property; the existing result-reader/full-model-open gate rebuilds before consuming those outputs. Repeated saves do not silently certify the pending result caches. XML IsCalculated, worksheet calculations and immediate grid refresh are unchanged.

Exceptions remain conservative: actual formula/structural/unknown changes, failed/cancelled preparation, recovery/integrity safeguards, and a formula normalization performed by compatibility preflight. In particular, an old long CONCATENATE requiring export repair still forces calculation before its first save. Unsupported export formulas remain blocked. No source XLSB, workbook schema, VBA, custom function or validation rule is changed by this release.

The user's 2.64 trace spent 7.511 s scanning, 7.134 s rebuilding, 14.290 s writing and 4.711 s restoring the engine (33.854 s total). The 2.65 private AGL first-stock-edit save has `rewritten=0`, `mode=deferred`, `initialRebuildPending=True` and **zero rebuild/engine-restore time**. That run took 25.917 s: 6.331 s preflight and 19.583 s native write. Other UI regressions ran concurrently, so this is not a controlled before/after speed comparison. Formula scanning and XLSB serialization remain real first-save costs; the change does not promise instantaneous saves.

## DIT toolbar

Normal visible order, with five separators:

Home | Save, Save As | Copy, Paste | Excel export, PDF export | History, Rebuild, Spreadsheet, Maximise working area | Settings, Help

Existing command objects, images, visibility/enabled state, click handlers and save/editor validation are retained. Conditional Return navigation remains next to Home when a linked interface needs it; it is hidden in the normal sequence. Runtime setup is idempotent and does not duplicate buttons. The order is local to DIT, not other interface toolbars.

Native DevExpress buttons had mixed designer/runtime `IsLeft` settings. That grouping reversed the designer buttons even after collection reordering. All DIT actions/separators now share one native flow (`IsLeft=False`) with explicit `VisibleIndex`; the panel itself remains left-aligned. Final verification checks native hit positions as well as collection order and a rendered toolbar image. No vendor source is copied or changed. Public control background: [DevExpress WindowsUI Button Panel](https://docs.devexpress.com/WindowsForms/114583/controls-and-libraries/navigation-controls/windowsui-button-panel).

The shared working-area toggle keeps its existing paired icons and remains icon-only. Caption/tooltip is **Maximise working area** in the expanded state and **Restore sidebars** in the compacted state. Shared DIT/Analyser instances receive the same wording and state changes. This does not change compact/restore behavior or hide the reopening edge strips.

## Clean-state Save buttons

Save is disabled while the model is clean in both DIT and File Instance. Save As remains available. A shared UI binding observes per-model dirty-state transitions; DIT additionally checks only the focused editor branch at UI idle. A modified, unposted editor enables DIT Save without dirtying or calculating the workbook, so the first click still invokes the existing validation/commit/save path. Cancelling that editor disables Save again. Successful Save/Save As clears both windows' Save buttons; subsequent committed edits, Undo or Redo enable them. This uses the established dirty flag, not an expensive workbook-content comparison.

Bindings detach from dirty events and UI idle on disposal (and early DIT resource release); structure-authoring previews do not acquire the binding. The observer performs no workbook writes, calculations or full control-tree scans. Save failure/cancellation retains the existing dirty-state restoration behavior. The latest addition is included in the 2.65 trial.

## Validation and evidence

- Debug and Release AnyCPU builds pass; active checkout is `C:/Repos/Abovo Summit`, branch main; version 2.65. Both normal executable locations were updated after the user closed Debug.
- Synthetic x86 save lifecycle passes: first and repeated deferred saves, initial output-reader rebuild, formula preflight independence, structural/name changes, model isolation, retained output activation, bulk guards, clean/dirty close, known check failures/recovery, cancellation/write failure and stale completion.
- Populated AGL x86 Release: real stock edit before first Save As takes deferred path, initial scan retained, saved formula/caches preserved, first reader rebuilds, subsequent Rent/Stock/Undo/Redo saves, IsCalculated immediate stock total, independent full-rebuild comparison and actual Summit reopen across 11,900 output cells pass. Cancellation/failure preservation passes. Original source and private input remain byte-identical.
- Native Debug DIT test: exact command/separator sequence, repeated ordering/object identity, state labels/icons, all visible hit targets and actual left-to-right visual order pass. Save commits an active Funding description editor to a private XLSB; Save As writes a separate private file. Original Demo hash unchanged. The initial collection-only test missed the mixed native IsLeft group; the final hit-test and screenshot checks catch and verify its correction.
- Demo x86 Debug save/reopen regression also passes, including the pending-result reader/reopen comparison across 11,900 output cells. Release panel-toggle tests pass all four initially open/closed navigator/summary combinations for three cycles each, retaining strips, widths and reopening behavior.
- Latest clean-state UI checks cover clean File Instance/DIT Save disabled, Save As enabled, unchanged editor, pending editor, cancellation, single-click Save, cross-window Undo/Redo and Save As state. The existing File Instance responsive rendering fixture passes at three widths and font sizes; its reflection call now supplies the newer optional results-pending HTML parameter explicitly. Synthetic Release x86 save lifecycle passes after adding the dirty-state event.
- Private artifacts/logs: `obj/first-save-state.log`, `obj/first-save-agl.log`, `obj/first-save-demo.log`, `obj/dit-toolbar-order-final.log`, `obj/dit-working-area-panels.log`, `obj/dit-save-clean-state-final.log`, `obj/file-instance-save-state.log`, `obj/save-clean-state-lifecycle.log`; disposable copies/renders are under `obj`. The final native Save-state runner exits successfully, including clean Save As with a new name and the source hash check. Original workbook files are never saved by the fixtures.

Final source checks: repository Blank SHA256 `0B4C06800FE998E8733A5D8EB9CDEFA04F517750275BB0CFDD532928900F5B79` and Demo `1ED79726D8D1129C699242C9E8B3E988AC920EA7108D3C0B105BF9B0BC8A5E9C` retain their baseline hashes. The populated AGL source matches the tested private input (`obj/SavePreparationTests/36d6fd373b3b4bfa90628f7f46eccf4d/input.xlsb`) at SHA256 `1E3C483E3D3B8B7D250C9B3B7DD742CFEB2DAE3AA5F9AD8B578D41C317BAF9AD`; this is the input identity for this trial, not an assumption of parity with older same-named files.

This builds on 2.64's verified pending-marker Excel SaveCopyAs preservation and 337-module VBA-source-hash check; those are historical evidence, not a newly executed Excel/VBA financial test for 2.65. Trusted Excel/VBA reopen, accountant/client acceptance, and physical mixed-DPI toolbar use remain manual.

## Suggested user check

Open the populated workbook afresh, check Save is greyed and Save As available, change a Stock figure, confirm the immediate calculated grid total, then Save. Save should enable for the edit and grey again on success. If `rewritten=0` and there was no structural/unknown change, expect `mode=deferred`, `initialRebuildPending=True`, and `restore=0 ms`, rather than `mode=rebuild`. The first formula scan remains; a subsequent value-only save in that session skips it. Reopen or visit outputs and confirm current results. Check button order and both working-area tooltip states at the client's display scale.
