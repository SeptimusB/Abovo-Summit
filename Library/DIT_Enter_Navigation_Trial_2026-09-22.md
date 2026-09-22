# DIT Enter navigation — Summit 2.69 trial

## Agreed behaviour and boundary

The user confirmed the direction-memory clarification following the Teams screenshot recorded against client report P016. Enter starts horizontally; Up/Down select vertical traversal, and Left/Right/Tab (including Shift+Tab) select horizontal traversal. Enter itself retains that axis; Shift+Enter reverses the traversal without changing the axis.

- Horizontal order is left-to-right, then the first editable cell of the next row.
- Vertical order is top-to-bottom, then the first editable cell of the next column.
- At a grid boundary, continue into the next visible grid or registered standalone workbook input, in top-to-bottom/left-to-right layout order.
- At the last input, wrap to the first; reverse traversal wraps from first to last. Reveal the destination editor by scrolling.
- Traversal is confined to the current visible tab. It neither switches tabs nor includes the action toolbar, sidebar, hidden controls or other windows. This boundary was stated to the user during implementation.
- An open dropdown keeps native Enter/arrow selection. Once the popup is closed, Enter again uses the DIT traversal. Ctrl/Alt+Enter are not intercepted.

"Commit" means accepting the edited cell through its existing typed ChangeManager handler; Enter does not save the workbook file. Existing Tab/arrow movement remains unchanged. On standalone controls, native caret/spin/Tab behaviour is retained, while those keys still select the remembered axis for subsequent Enter presses.

## Implementation and safeguards

`DataInterfaceTemplate.Navigation.vb` adds a per-DIT axis and a separate Enter selector, leaving the existing Tab/arrow selector intact. Candidates are re-evaluated after the queued workbook refresh so changed editability does not use stale targets. Coordinate comparison handles an origin that becomes locked/absent following the commit. Registered standalone inputs exclude hidden, disabled, read-only, calculated and source-locked cells and are removed from the registry on disposal.

Both in-column and vertical-row header helpers now pass the original Enter/Shift+Enter key data to the same selector. A native test also reproduced repeated activation replacing an already-open normal-grid header editor and immediately closing the replacement. Keyboard activation now reuses a live header editor; both helper types use this idempotent path.

Standalone buffered edits are validated and flushed through `SingleCell_Value_Push`, not direct workbook assignment; reverted/rejected edits stay in place. Normal/vertical grids retain native validation and posting. Header helpers retain their existing commit/failure checks. No calculation, structural insertion, rollback, save policy, workbook schema or XML definition was changed. XML remains 1754.

## Regression scope

`Tools/Test-EditorNavigation.ps1` compiles the native C# fixture against the selected Debug/Release binaries and works only on a private copy of the repository Demo master. It checks the original source SHA-256 after every run.

The extended fixture covers Funding vertical-grid cells and header editors, the two CPI/RPI grids, Rent's normal/banded grid and header editors, and actual Global Assumptions standalone inputs. Expected Enter targets come from an independently built circular ordering, not from the production selector. Every enumerated point is checked in both axes and directions, with actual editor key events at the important boundaries.

Checks include default/remembered direction; open/closed editors; next row/column; adjacent grids; first/last wrap; Shift+Enter; focused and visible destinations; native dropdown ownership; a removed/locked origin; standalone native Tab dispatch; read-only/disabled standalone exclusion; validation rejection and subsequent recovery; typed commits, dirty state, one-step Undo and Redo. Existing Funding date validation, clipboard rollback, scroll/focus retention, toolbar, Save and Save As/reopen checks remain in the same fixture.

DevExpress only raises modified-value validation when appropriate; artificial rejection tests explicitly mark the editor modified. This follows the [25.2 validation contract](https://docs.devexpress.com/WindowsForms/DevExpress.XtraEditors.BaseEdit.DoValidate?v=25.2), not a production change forcing repeated validation of unchanged values.

Final execution results are recorded in the client-review checkpoint report. An automated native pass is not physical-keyboard, mixed-DPI or client-workbook acceptance. No new financial or Excel/VBA calculation equivalence is claimed by this UI change.

## Client test steps

1. In Funding, edit a description and press Enter repeatedly. Confirm movement across the row. Press Down once, then Enter repeatedly; confirm movement down the column and then into the next column.
2. Use Shift+Enter at row, column, grid and visible-tab boundaries. Confirm reverse order and that the destination is visible. Left/Right/Tab should return Enter to horizontal traversal.
3. Try Funding/Rent header dropdowns and dates: Enter should accept an open choices list, then navigate on the subsequent Enter. Reject an invalid input, correct it, and confirm editing can continue.
4. Try CPI/RPI between-grid transitions and Global standalone inputs. Confirm no unexpected tab/sidebar switch and no workbook disk save on Enter. Verify Undo and the existing Save behaviour on a disposable client workbook.
