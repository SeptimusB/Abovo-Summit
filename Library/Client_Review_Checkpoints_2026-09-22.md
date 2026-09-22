# Client review and unattended validation checkpoints

## Authority and review rules

The client report is `D:/Downloads/Summit v2.51 comments.docx`, SHA-256 `955D44B6C714ED9BA6D5F9B4560C983764B38B83161A71B88C2BF400211BD9ED`. The complete document body was read: 137 paragraphs including five trailing blanks (last content at P132), no embedded images, comments, footnotes or tracked changes. The user confirmed green `196B24` means "agreed", but errors may remain. Their subsequent correction on 22 September supersedes the earlier orange interpretation: orange `E97132` means the user believes it is fixed, with verification still required. All six orange stakeholder entries and the generator/legend now use that meaning; engineering findings are not automatically closed. Inherited/default text is unclassified. Neither colour is engineering verification. Paragraph references below are extraction locators, not printed page numbers. The original document is not edited.

The report concerns 2.51; the starting implementation is 2.67. Historical fixes are candidates for regression testing, not grounds to mark client acceptance complete. The bundled DOCX renderer is unavailable on this Windows runtime; review uses complete OOXML text/run-colour inspection, not claimed page-layout verification.

Each item records two independent states: client/user status (including exact source colour), and engineering evidence (unreproduced, reproduced, implemented, automated check passed, or manual acceptance required). A technical pass never implies client or accountant approval.

## Checkpoint gates

1. **CP0 baseline:** preserve and commit the existing 2.67 integrity trial after Debug/Release builds and isolated safety regressions. Exclude local `.codex/`, private workbooks and generated evidence. Record the starting source hashes and leave original workbooks untouched.
2. **CP1 intake:** retain every report item and source colour in a stable issue register. Cross-reference the outstanding work list and known repairs. Gather specific ambiguity questions without blocking independent work.
3. **CP2 recovery and Funding:** automate ten-column insertion, safe due-recovery scheduling, source preservation, recovery reopen/history, normal XLSB Save As and macro-disabled Excel round-trip on disposable copies. Capture timings and exceptions. Do not mutate the user's live process or settings.
4. **CP3 bounded repairs:** reproduce one coherent defect family before editing. Preserve workbook authority; distinguish display formatting from typed numeric storage, protection from missing editors, and user layout requests from financial semantics. Test through real shared services/editors, then build both configurations and checkpoint that batch separately.
5. **CP4 review handoff:** publish pass/fail/blocked results, exact evidence paths, remaining questions and client test steps. Keep performance observations separate from controlled benchmarks and financial acceptance. No push or release publication is implied.

## Stop and defer conditions

- Missing colour legend, unspecified workbook/field or ambiguous intended behaviour: preserve the report and ask; no speculative fix.
- Excel/VBA financial interpretation, bespoke template policy, minimum CPU/OS matrix, deployment/password policy and optional Excel offload: require the corresponding decisions; they cannot be settled by a UI fixture.
- Unexplained formula/name/value differences or source-hash change: stop that branch, keep failure evidence and last good checkpoint. Do not save the original or label a partial run passed.
- Existing client Excel sessions, open Summit files and persistent user settings are not test targets. Use privately owned processes and copies. Never terminate processes by application name.
- Running calculation/serialization is atomic. User input may defer a subsequent scheduled unit; tests must not interrupt a live workbook operation.

## Baseline evidence

CP0 is committed locally as `2c01370` (2.67 baseline, idle-integrity trial and review gates). No push was requested or performed. The source report is now indexed in [the issue register](Client_Report_2026-09-22.md), with exact run colours in the companion JSON: 90 issue paragraphs, seven green, six orange and 77 inherited/default. Multi-part paragraphs remain intact; passing one sub-check does not close the whole item.

22 September: Debug and Release builds pass. Release idle-integrity synthetic safety fixture passes (`obj/IdleIntegrityTests/117d603084554c40a7459081c46ac29e`). Debug recovery synthetic fixture passes (`obj/RecoveryTests/0ab6cd9fc40c47d1b6c8f7e5c3cb4668`), including dirty/Undo preservation, failure retention, unrelated-name collision, idle/save/grouped-edit guards, notification owner and settings cancellation. These are engineering checks, not client or financial acceptance.

Latest client Funding trace: test 2.67 x86 Debug, ten records, 126.484 seconds including 4.929 seconds interface rebuild. Source column shifts 60.772 seconds, TDB sync 36.859 seconds. No recovery event appears in that supplied log. A first-chance InvalidOperationException is not diagnosed without its message/stack; capture it in the private combined test.

## CP2 — combined Funding and scheduled recovery

The automated 2.67 Release x86 test passes on a private copy of the user-supplied AGL file. Source SHA-256 `30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47`, 19,526,399 bytes. This source is different from the earlier frozen 11.7 MB format benchmark: timings are observations, not a controlled before/after performance claim.

- Ten Funding columns inserted; engine/mode, every sheet's protection and all supported Transactional DB mirror geometry checks pass.
- Scheduled recovery defers during structural mutation, saving and recent activity; runs when due and idle; emits start/completion messages; does not repeat for an unchanged revision.
- Original bytes/path, dirty state and Undo are preserved. Recovery rereads with an exact formula/input/name/sheet-order/protection digest. Full Summit reopen retains inserted columns, matching mirrors and read-only prior-session edit history; pending results are calculated before use. Structural changes themselves are not claimed to have a persisted Undo journal.
- Recovery Save As produces a separate normal XLSB and clears recovery guidance. No original file is overwritten.
- A private Excel instance opens both results without repair and saves separate copies, with macros/events/link updates/calculation disabled. All 335 VBA module identities/source hashes, existing custom XML/properties and the recovery-history part are preserved in all four outputs. VBA execution is not tested.
- Excel's recovery-XLSM save changes only six already-classified unused `=#NAME?` function-name placeholders. Excel's normal-XLSB save changes 1,140 formula expressions on Development Expenditure (implicit-intersection `SINGLE` wrappers); the exact-digest test correctly fails. Independent full Summit recalculation agrees across 44,906 SOCI, SOFP, cashflow, Check Sheet and Development Expenditure cells, tolerance 1e-7. This does not certify Excel/VBA financial equivalence or excuse the formula-text difference.

Observed durations: insertion 109.690 s; source column shifts 55.985 s; TDB post-actions 32.983 s; scheduled recovery write 13.827 s; subsequent structurally dirty recovered XLSB Save As 35.122 s. None is a UI-responsiveness or minimum-hardware acceptance result. The captured first-chance exception was the optional `TDB Snapshot` worksheet lookup; the snapshot predicate catches it and returns false. It did not fail insertion/recovery. This identifies the exception in this run, not every earlier client exception.

Evidence: `obj/client-review-funding-recovery-x86.log`, `obj/client-review-excel-roundtrip.log`, `obj/client-review-recovery-native-roundtrip.log`, `obj/client-review-normal-native-roundtrip.log` (expected exact-match failure), `obj/client-review-normal-sheet-delta.log`, `obj/client-review-normal-calculated-values.log`, `obj/client-review-vba.log`. Private workbooks: `obj/RecoveryTests/c4c01cc4cecd4ae98fc2b9b7c76f6e5a`; Excel copies: `obj/RecoveryExcelRoundtrip/d881a639edad466d96ac4c512493dac1`.

## CP3 — first bounded client-report repair

Intake, combined recovery validation and the first two repairs are checkpointed locally as `e1140cf`.

2.68 / XML 1753 contains only two master-verified Development XML repairs at this checkpoint:

- **P067:** Identified Development's SHG Calculation Basis was incorrectly bound to `SHGProfileIn` (row 122, validated by `SHGProfile`). The master labels/validation identify the separate `SHGMethodsIn` (row 125, validated by `SHGMethods`); multiyear Development already used that correct range. A failing regression was captured before changing the identified binding. Real DIT dropdown commits now alter the separate correct cells; two Undo operations restore them independently; dirty state and protection pass.
- **P062:** `Rep_DevBP_01d` is row 43, labelled `Period Units into Mgmt (to)` in the master, not row 37's works-completion input. Caption and tooltip now describe the correct existing input. Range, editor and dates are unchanged.

Debug and Release builds pass. The new fixture passes against the repository Demo master without saving it. Existing Debug native Funding focus/date/navigation/Save/Save As regression also passes: P012, the tested portion of P095, P101, and adjacent-grid/header navigation. This does not certify every DPI layout or the separate Enter-key request P016. Evidence: `obj/client-report-before-fix.log` (expected failure), `obj/client-report-inspection.log`, `obj/client-report-after-fix.log`, `obj/client-review-build-debug.log`, `obj/client-review-build-release.log`, `obj/client-review-navigation-debug.log`.

## CP3b — width stability and two monetary input types

The same 2.68 test delivery now uses XML 1754. Additional changes are limited to:

- **P029/P042/P048 shared cause:** the normal DIT post-edit handler re-ran BestFit and added 15% padding to the edited column. A real dropdown test reproduced a manually set width changing from 333 to 217 pixels after one commit. Removed that post-edit sizing only; initial build and explicit font/layout fitting remain. Repeated dropdown edits and numeric edits now preserve their chosen widths. These tests cover the shared handler, not every named client screen or physical DPI setting.
- **P127:** `Rep_OCA_01` Amount was typed as integer in XML. The real editor stored 2345.67 as 2346. Changed this monetary field to the existing decimal `M` type.
- **P130:** Journal Amount was also typed as integer. The same real-editor test failed for 1234.56; changed only Amount to `M`, retaining Year as `I`.

Actual DIT writes use ChangeManager, not direct test assignments. Monetary tests verify exact stored values, unchanged source number formats/protection, Undo/Redo and normal Summit XLSB Save As/reopen on private copies. Workbook display formats intentionally remain unchanged: accepting decimals does not authorise changing the model's whole-number display. Other monetary/count/year definitions were not changed speculatively.

Debug and Release build and both final native fixtures pass, including Save As/reopen for both monetary values. Both runnable output folders contain the same XML 1754 as the repository (SHA-256 `5F1DFDF8516CDF5983B7739729471AB819C577D993FFE54C80847C8A3A9E3FDF`). Evidence: `obj/client-report-width-before-fix.log`, `obj/client-report-decimal-before-fix.log`, `obj/client-report-oca-before-fix.log` (three expected reproductions), `obj/client-report-final-release.log`, `obj/client-report-final-debug.log`. The fixture also logs a non-fatal WebView2 class-unregistration diagnostic at private form teardown; no claim is made to have repaired unrelated browser lifecycle issues.

**P132 regression:** `obj/client-review-journals-regression.log` passes five-row add, both formula mirrors, dirty/protection preservation, save/reopen, delete back to original geometry, noncontiguous deletion, and rejection of pre-existing name inconsistency before mutation. This confirms the tested later repair rather than treating the old report as a new uninvestigated failure.

Repository Blank/Demo hashes still match the baseline; no authoritative workbook or source DOCX was edited. No passwords, raw VBA or extraction dependencies were added to the repository.

## P016 clarification — Enter traversal and structural safeguards

Additional evidence supplied by the user on 22 September: `D:/TEMP/codex-clipboard-2b08031b-9ec2-4aee-87ae-af0a1a030f02.png`, a Teams discussion. The visible client response is "Sounds ok" to the proposed keyboard behaviour. This is agreement to a requirement, not proof that the current executable implements it. The screenshot is context for the authorised client review, not instructions to execute every proposal in the conversation. The final message is cropped and is not used as evidence.

The visible proposed behaviour is:

- Enter commits an open editor and advances right; at the end of a row it moves to the first editable cell of the next row.
- If the last navigation was Down, Enter advances down the same column; at its bottom it moves to the top of the next column to the right.
- At the end of the grid it advances to the next grid or input. At the end of the interface it wraps to the top and scrolls to the destination.
- Shift+Enter reverses the traversal. Here, committing an input is distinct from saving the workbook file to disk.

Pre-2.69 source inspection: `DataInterfaceTemplate.Navigation.vb` recognised Tab/Left/Right/Up/Down but excluded Enter. It had no remembered traversal axis. Tab traversed rows and adjacent grids, but stopped at the last host; vertical arrows used the same column and then the adjacent grid, not the next column. The shared candidate list covered visible editable normal/vertical grid cells and their in-column/row header editors; standalone inputs were not included. Native controls handled Enter separately, so this finding did not mean Enter never committed a value.

The user subsequently confirmed Up/Down selects the vertical axis and Left/Right/Tab returns to the horizontal axis. Implementation was explicitly bounded to the current visible tab, including eligible standalone inputs, without switching tabs or entering sidebars/actions. Preserve locked-cell exclusion, validation failures retaining focus, native dropdown selection, single typed ChangeManager commits, Undo and existing Tab/arrow behaviour. Regression gates cover open/closed editors, row/column boundaries, multiple grids, header editors, standalone inputs, forward/reverse wrap, scrolling, failed validation and direction changes, with independently specified expected targets.

The same conversation calls a parallel Excel shadow workbook/comparison a nice-to-have at present. It does not approve removing transaction rollback, integrity checks or safe working copies. The proposed save -> Excel/VBA operation -> save -> Summit reload remains a separate optional prototype, not an implemented production route. Retain the existing safeguards and distinguish a successfully returned insertion command from independently verified workbook integrity.

The initial documentation-only checkpoint was `38c75f6`, leaving Summit at 2.68. The subsequent confirmed implementation is the 2.69 trial below; XML remains 1754.

## CP3c — directional Enter/Shift+Enter trial

See [the navigation trial and client steps](DIT_Enter_Navigation_Trial_2026-09-22.md). Summit 2.69 shares Enter traversal between normal grids, vertical grids, both header-editor helpers and registered standalone workbook inputs. It retains native dropdown ownership and existing Tab/arrow movement. A repeated-header-activation regression was reproduced and repaired by reusing the live editor instead of replacing it. There is no change to workbook schema, financial logic, rollback, save policy or integrity safeguards.

Debug and Release builds pass. The final extended native fixtures both pass: Funding, CPI/RPI, Rent headers and Global standalone inputs; independent all-point forward/reverse order checks in both axes; actual key handlers at row/column/grid/tab boundaries; focus/scroll; open/closed editors; dropdown ownership; validation rejection and recovery; read-only/disabled standalone exclusion; real typed commits, dirty state, one-step Undo/Redo, and the existing date/clipboard/save/reopen regression. Tests run sequentially to avoid competing for focus. Artificial validation-rejection cases mark the editor modified, as required by DevExpress; that test setup does not alter production validation policy.

Evidence: `obj/enter-build-debug.log`, `obj/enter-build-release.log`, `obj/enter-navigation-final-debug.log`, `obj/enter-navigation-final-release.log`. Private output folders: `obj/EditorNavigationTests/e29eec932db246718512dda687ba6fa9` (Debug), `obj/EditorNavigationTests/cee0fe1c6d5b443b9ba37c9ecd997c76` (Release). The original header re-entry failure is retained in `obj/enter-header-before-fix.log`.

Repository Blank/Demo hashes remain the recorded baseline. Both executable folders contain the unchanged XML 1754/hash recorded under CP3b. No user workbook, source report, original master or live settings were modified. Native injected-key/editor tests do not certify physical keyboard routing on every client display or workbook; that acceptance remains open. No new Excel/VBA execution or financial sign-off is claimed.

## Remaining investigations and ambiguity gates

1. **Input precision, negative percentages and locks:** reproduce through the actual editor, clipboard parser and source cell protection. Do not change every `I` XML type to decimal: counts/years must remain integers. P022 needs the exact disagreeing Check Sheet message and workbook; P085's unspecified conditional formatting needs a field/rule example; P099/P100's "ghost" needs a defined desired date/blank appearance.
2. **Column widths / scrolling / layouts:** shared post-edit width repair is covered above; verify P029/P042/P048 on the client's specific screens. Physical monitor/DPI reports require client acceptance. Housing Asset Grant/Remaining Useful Life sections contain overlapping grant/depreciation sources, so P122/P123 is not simply two labels to swap; decide the intended separation before rewriting sections. Three concise questions about report workbook/Check Sheet, asset-tab separation and date "ghost" meaning have been sent; no replies are assumed.
3. **Structural and performance work:** source column shifts and TDB mirror resizing dominate Funding. Keep optional Excel automation and capacity/schema changes as separately reviewed prototypes; do not silently switch execution or alter authoritative workbook schema. Development/multiyear warm-analyser/snapshot benchmarks and the 32 GB client-class matrix remain open.
4. **Business/financial requests:** retain accountant acceptance of the repaired balance sheet, Excel/VBA calculations, grants/component structure and bespoke upgrade contracts. A copied screenshot or green paragraph is not financial approval.
5. **Client release:** no installer publication/push or recent-files feature added. Remind the user about the deferred recent-files HTML launcher when preparing the next client-testing release.
