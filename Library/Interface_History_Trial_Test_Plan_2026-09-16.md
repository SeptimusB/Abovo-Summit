# Interface History trial test plan

## Scope

The trial adds a most-recently-used Interface History list to every Group Interface Template (GIT) right sidebar. It records all standard DIT and special/class interfaces activated through a GIT, plus the Financial Forecast Return (FFR) and Stress Test. FormMain and transient dialogs are deliberately excluded. When two or more full models are open, the user can switch between **This model** and **All models**.

History entries are session-only and are not written to the XLSB or user settings. Revisiting a destination moves its existing record to the top rather than creating a duplicate. Hidden interface objects are reactivated; they are not deliberately disposed or recreated. The history scope choice alone is persisted as a user preference.

## Principal risks and controls

| Risk | Impact | Implementation control | Required test |
| --- | --- | --- | --- |
| A history action recreates an interface and loses UI state | High | Each entry retains its hosting GIT; FFR and Stress Test use the existing instances owned by `FileInstanceInterface` | Set distinctive state in each target, hide it, return through history and confirm the state remains |
| Two live GITs route an entry to the wrong window | High | GIT entries retain the exact host instance as well as model/group/child identifiers | Open two different GITs side by side and return alternately to entries from both |
| Entries from different groups collide because CSIDs repeat | High | GIT identity is model + GSID + CSID, not CSID alone | Visit child interfaces with the same CSID in different groups and confirm both remain |
| Duplicate entries accumulate | Medium | A stable destination key is removed and reinserted at the top on every visit | Visit A, B, A and confirm the order is A, B with two rows only |
| Stress Test modal behaviour causes re-entrancy or a second form | High | History calls the existing Stress Test launcher and preserved instance | Open, alter state, hide, reopen from history, then repeat several times |
| FFR is reconstructed or loses its selected page | High | History calls the existing preserved FFR instance | Select a non-default FFR page, hide and restore it through history |
| This model displays or activates another model's history | High | Each workbook model owns a separate history service; the coordinator filters by model | Open two models and confirm This model lists only its own entries |
| All models activates the wrong model or a reused numeric model ID | High | Entries carry a unique open-model instance token checked before activation | Navigate between two model types, close one, reopen it, and verify stale entries cannot navigate |
| GIT windows for one model overwrite another model's windows | High | GIT registries are now instance-owned rather than shared | Open GITs in two models, close one model and activate GITs in the other |
| Scope control appears with one model or forgets the choice | Medium | The control appears only with two or more full models; choice is user-scoped | Check the one-model and two-model states, then restart Summit with two models |
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
10. Open a second model, ideally of another supported type. Confirm the switch appears; This model lists only local entries and All models lists both with model names.
11. Double-click a foreign-model entry in All models. Confirm the exact preserved GIT/document, FFR or Stress Test instance activates; repeat in the opposite direction.
12. Close FFR and Stress Test using X, then restore each through history. Confirm their prior tab/scenario state remains and no replacement instance was created.
13. Close one model while the other remains open. Confirm closed-model entries disappear, the scope control hides, and the remaining model's GITs still work.
14. Reopen a model into a potentially reused numeric slot. Confirm earlier closed-model entries cannot activate it.
15. Restart Summit, open two models and confirm the chosen scope returns.
16. Repeat representative navigation under the client's font scale and monitor/DPI arrangement.

## Pass criteria

- The list is newest-first and contains no repeated logical destinations.
- Each action returns to the original hidden interface object with its UI state preserved.
- Side-by-side GITs and multiple models do not cross-route; All models activates the selected model.
- FFR and Stress Test preserve their existing show/hide semantics.
- Ordinary interface navigation remains functional if history recording or restoration cannot complete.
- No unhandled exception, stale event callback, workbook mutation or material navigation delay is observed.

## Automated validation completed

- Debug configuration builds successfully.
- Release configuration builds successfully.

The workbook/UI acceptance sequence remains manual because the repository has no automated integration harness for these interface lifecycles.
