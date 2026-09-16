# Interface History trial test plan

## Scope

The trial adds a model-scoped, most-recently-used Interface History list to every Group Interface Template (GIT) right sidebar. It records all standard DIT and special/class interfaces activated through a GIT, plus the Financial Forecast Return (FFR) and Stress Test. FormMain and transient dialogs are deliberately excluded.

History is session-only and is not written to the XLSB or user settings. Revisiting a destination moves its existing record to the top rather than creating a duplicate. Hidden interface objects are reactivated; they are not deliberately disposed or recreated.

## Principal risks and controls

| Risk | Impact | Implementation control | Required test |
| --- | --- | --- | --- |
| A history action recreates an interface and loses UI state | High | Each entry retains its hosting GIT; FFR and Stress Test use the existing instances owned by `FileInstanceInterface` | Set distinctive state in each target, hide it, return through history and confirm the state remains |
| Two live GITs route an entry to the wrong window | High | GIT entries retain the exact host instance as well as model/group/child identifiers | Open two different GITs side by side and return alternately to entries from both |
| Entries from different groups collide because CSIDs repeat | High | GIT identity is model + GSID + CSID, not CSID alone | Visit child interfaces with the same CSID in different groups and confirm both remain |
| Duplicate entries accumulate | Medium | A stable destination key is removed and reinserted at the top on every visit | Visit A, B, A and confirm the order is A, B with two rows only |
| Stress Test modal behaviour causes re-entrancy or a second form | High | History calls the existing Stress Test launcher and preserved instance | Open, alter state, hide, reopen from history, then repeat several times |
| FFR is reconstructed or loses its selected page | High | History calls the existing preserved FFR instance | Select a non-default FFR page, hide and restore it through history |
| One model displays or activates another model's history | High | Each workbook model owns a separate history service | Open two models, navigate in each and compare their sidebar lists and activation targets |
| Sidebar subscribers survive model shutdown | High | Each history view unsubscribes when its GIT is disposed; the service is disposed during model close | Close a model with several GITs open, reopen another model and watch for errors or stale rows |
| A history failure blocks ordinary navigation | High | History recording is ancillary and guarded; activation failures are reported to System Messages | Exercise navigation during rapid hide/show and model-close boundaries; ordinary navigation must continue |
| Double-clicking headers or blank grid space opens an item | Medium | Mouse hit testing accepts data rows only | Double-click each column header and empty grid area; nothing should open |
| Font scaling or narrow sidebar widths make the list unusable | Medium | Native DevExpress grid, auto-width columns and application user scaling | Test at 0.3, 1.0 and 2.0 application scale and on each available monitor/DPI |
| The history update adds noticeable navigation delay | Low | Maximum 50 in-memory entries and a small bound grid | Fill the list and confirm navigation remains immediate |
| A primary workspace is omitted from the allow-list | Medium | Trial explicitly covers GIT children/class forms, FFR and Stress Test | Inventory the production navigation buttons during acceptance and identify any additional persistent workspace to register |

## Manual acceptance sequence

1. Open the Demo master in a Debug build and open Assumptions, Workings and Outputs GITs.
2. Visit at least three DIT interfaces and two special/class interfaces, including Analysis V2.
3. Confirm all open GIT sidebars display the same newest-first model history.
4. Revisit an older entry from the navigator and confirm it moves to the top without duplication.
5. Double-click an entry hosted by each GIT and confirm the correct window and child document activate.
6. Change a visible, non-workbook UI state such as selected tab, expanded section, scroll position or filter; hide the form and restore it from history.
7. Open FFR, select a non-default page, hide it, and restore it from history without losing that page.
8. Open Stress Test, select a non-default tab or scenario, hide it, and restore it from history without creating a second form.
9. Double-click history column headers and blank space and confirm no navigation occurs. Confirm Enter activates the focused data row.
10. Open a second model and confirm its history is independent of the first model.
11. Close a model while multiple history-enabled GITs are open, then open another model and confirm no stale entries or event errors appear.
12. Repeat representative navigation under the client's font scale and monitor/DPI arrangement.

## Pass criteria

- The list is newest-first and contains no repeated logical destinations.
- Each action returns to the original hidden interface object with its UI state preserved.
- Side-by-side GITs and multiple models do not cross-route.
- FFR and Stress Test preserve their existing show/hide semantics.
- Ordinary interface navigation remains functional if history recording or restoration cannot complete.
- No unhandled exception, stale event callback, workbook mutation or material navigation delay is observed.

## Automated validation completed

- Debug configuration builds successfully.
- Release configuration builds successfully.

The workbook/UI acceptance sequence remains manual because the repository has no automated integration harness for these interface lifecycles.
