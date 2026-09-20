# FFR and Stress Test history - 20 September 2026

Test release: 2.37.

## Findings and boundaries

Both standalone interfaces already record visits in the model's `InterfaceHistoryService`. Their ordinary single-cell edits and rectangular pastes use `ModelChangeManagerV2` (pastes are grouped). The missing pieces were direct Change History access, model-level shortcut interception in their native editors, and notification-driven presentation refresh after undo/redo.

FFR already used `Show()` and was non-modal. Stress Test used `ShowDialog()`, blocking the other model windows. Stress Test now uses `Show()`, activates the retained instance and restores it if minimised. User Close/X continues to hide rather than dispose both interfaces. File pickers and confirmation dialogs remain modal by design.

This does **not** make every command undoable. Stress mode switching, scenario capture/copy/clear/import, dashboard generation and sensitivity-list structural operations include direct range/name changes and do not have complete reversible command snapshots. They have not been converted to cell-journal operations. FFR return export creates an external file and cannot be undone by workbook history. Existing history conflict checks still reject reversal when the recorded cell no longer matches its expected state. Bulk history needs a separate domain-level transaction design, not a claim that a button makes it reversible.

## Implementation

- Both windows expose Change History, opening the existing model-scoped, non-modal `HistoryManagerV2`.
- `ModelFormHistoryBinding` attaches to each form and its own model manager. Ctrl+Z, Ctrl+Y and Ctrl+Shift+Z are intercepted before a child native editor consumes them as local text history. The message filter accepts only controls within its owning form, not another model/window.
- Undo/redo refreshes visible forms synchronously after the journal's state transition. Ordinary edits only mark the standalone presentation stale; an activation/reopen refresh avoids rebuilding grids inside `CellValueChanged` and avoids unnecessary redraws on each typed edit.
- FFR reloads only the selected page; other pages reload on selection. Its editable views hide existing grid editors under their existing loading guards before reading workbook values.
- Stress Test reloads the live name, capture slot, first-tab data and scenario selectors, plus the visible native page. Its refresh suppresses posting, discards the temporary band editor and does not run another full calculation after undo. Planner navigation also updates the target subtab.
- Model close explicitly disposes both standalone forms while the model/workbook still exists; their keyboard filters and history subscriptions detach. This is separate from user close/hide and prevents a modeless window retaining a disposed workbook.
- No workbook schema, formula, validation, package or change-manager snapshot format was changed.

## Verification

`Tools/Test-StandaloneHistory.ps1` compiles an isolated fixture against the selected build. It loads an authoritative master into memory and uses real application services and native WinForms/DevExpress controls. Test windows are transparent and excluded from the taskbar. The fixture never saves the model and checks the source SHA-256 afterwards.

Verified against both the Demo and Blank masters:

- Both actual FileInstance launch paths return non-modal, simultaneously enabled forms, with distinct deduplicated Interface History records and Change History controls.
- A real FFR Front Sheet edit creates workbook history; undo/redo updates its visible editor.
- A native Key Definitions VGrid cell event writes through the journal. Undo refreshes the VGrid datasource and closes an active stale editor.
- A Stress Test live-name edit supports undo/redo without phantom history entries from refresh.
- Undo while Stress Test is hidden is reflected on reopen. Close/X preserves both instances; FFR retains its selected tab. Reopening does not duplicate history entries.
- Model-close standalone cleanup disposes both forms and detaches both history bindings.
- The source workbook hash remains unchanged.

Debug and Release builds passed. Physical keyboard routing, multiple simultaneously open models and Excel/VBA round-tripping remain manual acceptance checks below; they are not implied by the isolated fixture results.

## Required interactive acceptance

1. Open FFR, Stress Test and an assumptions window side by side. Verify all remain usable and the intended model's Change History opens from each.
2. Edit/paste in FFR Front Sheet, Inputs/Adjustments, Workings and Key Definitions. Undo/redo using the history buttons and Ctrl+Z/Ctrl+Y/Ctrl+Shift+Z, both with and without an editor open. Confirm workbook and displayed values agree, and each paste remains one history group.
3. Repeat for live stress assumptions/name/capture slot, planner scenario name/inputs, targets and comparative selectors. Verify calculated outputs and no duplicate history entries; switch pages after undo.
4. Hide each form, undo from another window, reopen via Interface History and confirm refreshed values and retained tabs/scenario selection.
5. Open a second model and repeat shortcuts while moving focus between windows. Confirm history never targets the wrong workbook and closed-model entries disappear.
6. Close the model while both standalone windows are open. Verify both disappear, and subsequent models have no stale callbacks or shortcuts.
7. Run scenario generation/import, quick capture, mode switching and FFR export as ordinary regression smoke tests. They remain outside the automatic undo guarantee stated above.
8. On a disposable copy, Summit Save As -> Excel/VBA open/save -> Summit reopen; confirm edits persist and workbook checks behave as before.

No commit or push was requested for this change. Existing unrelated edits were preserved.
