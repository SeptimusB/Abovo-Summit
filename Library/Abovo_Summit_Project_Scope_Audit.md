# Abovo Summit TestFileClean scope audit

Audit date: 22 August 2026

## Current authority

`Z:\Sandbox\TestFileClean.xlsb` is the authoritative workbook master. `Library/TestFileMigrated.xlsb` deliberately retains its repository filename for existing tooling, but its bytes are an exact copy of that master:

- Size: 11,841,039 bytes
- SHA-256: `B248B1C733E1E3293536FBE1DBC9576D56FD4D34BEB74A3898B7F3D7333BCFBE`
- Worksheets: 281
- Defined names: 1,758
- `Transactional_Records`: `Transactional DB!A6:BV1599` (1,594 rows by 74 columns)
- Period fields: 40, from `2025/26` to `2064/65`
- Close validation range: `Outputs_CheckSheet = Check Sheet!A9:H63`
- Embedded VBA project: present
- ChatGPT-added `__Abovo_*` worksheets/names: none

The source workbook was inspected read-only through Excel automation with macros and events disabled and was never saved. The repository copy hash was verified against the source immediately after copying. The user owns external backup management; Summit does not create or retain an internal master backup during this replacement.

## Transactional DB contract

The workbook's formula-backed `Transactional DB` sheet and its workbook-defined names remain authoritative. `TransactionalDBSynchroniser` implements the same named-range sizing rules as the workbook's `Summit_Compatibility` VBA layer; it does not create hidden template, metadata, or materialisation sheets.

The twelve canonical source rules and all 89 `TransCopy_*` targets were checked read-only. Every target exists on `Transactional DB` and has the expected row count derived from its source named range. Development Identified therefore uses the original `TransCopy_DevptSingle_A:N` ranges on `Transactional DB`, not auxiliary sheets.

Structural commands now return a failed `AbovoTransaction` if post-change Transactional DB synchronisation fails. Interface dependency invalidation still runs so an already-changed workbook cannot leave visible controls appearing current.

## Analysis interfaces

Both `BPIncomeExpenditureAnalyser` (V1) and `BPIncomeExpenditureAnalyserV2` bind the canonical `Transactional_Records` range. The Sample range provides the original 74 datasource columns, including 40 periods. V1 adds its `ItemDesc` column as the 75th unbound grid column; it is not an XLSB field. V2 discovers period columns and required grouping fields dynamically and uses the same unbound description column.

The model resource registry continues to enforce exclusive ownership of the live `Transactional_Records` RangeDataSource, allowing V1 and V2 to disconnect and reconnect safely.

## Removed schema-10 implementation

`WorkbookMigrationManager` and `TransactionDBMaterialiser` have been removed from the project and model lifecycle. Full-model open no longer mutates workbook schema. The Development-specific materialiser branch and its `__Abovo_DevMat_Test` diagnostics have also been removed from `TransactionalDBSynchroniser`.

The retired schema-10 audit, index, and workbook analysis are retained with `.Schema10` filenames for historical evidence; they are not current product contracts.

## Validation boundary

### Model shutdown ordering

Version 7.28 keeps `ExcelModels(ModelID)` and its workbook registered until registered analyser resources and all group/data interfaces have detached. DataInterfaceTemplate now removes unbound callbacks before clearing grid columns or bindings, disposes its data sources and controls idempotently, and returns no value from a late callback once the model is marked as closing. Group interfaces are disposed directly during final model shutdown rather than passing through their cancel-and-hide user-close handler.

### Merge-down Add Lines metadata

Version 7.29 preserves `ISEDatasource.RowExpandByNR` when constructing `MergeDownAndPivot` datasets. This restores the workbook-owned Add Lines targets for Development Details (`HouseTypeInID`) and Development Dets. Multi-Year (`HouseTypeInMY`) without changing `Structure.xml` or either named range.

### Close calculation and Check Sheet validation

Version 7.30 performs a recursive full-rebuild calculation before any model close, temporarily includes the Transactional DB in calculation, and then examines every populated status in the workbook-defined `Outputs_CheckSheet` range. Blank section headings are ignored; `OK` is the only passing populated status. Calculation errors, missing/malformed validation ranges, formula errors and workbook check failures all prevent normal close. The user may cancel and return to the model, or continue only by saving a macro-preserving XLSB copy to a different path; the validation path cannot overwrite the original model.

### Explicit structural Add Lines rules

Version 7.31 adds an explicit `StructureRuleID` contract from `Structure.xml` through `ISEDatasource`, `DataCellRange` and `ActionToken`. Multi-range Assumptions grids now declare their semantic workbook rule instead of depending on the displayed named range being recognised as an implicit alias. Declared rules fail closed if the rule manager is unavailable or the ID is unknown, so Summit cannot silently fall back to resizing only one named range. Legacy single-range grids retain named-range inference and the original generic expansion path.

Version 7.34 restores the required domain-event lifecycle for declared structural grids. `StructureAddCommand` and `StructureDeleteCommand` flow with the rule ID into the DIT action; the requested count is then dispatched through `EventCoordinator`, `BPGridEvent` and the matching `SpecificRowColumnEvents` sub before `WorkbookStructureRuleManager` performs the multi-range mutation. Contraction from the footer uses the matching delete event with an explicit delete-last count. Missing or unknown commands fail closed. A live disposable-workbook trace confirmed that adding five Capital Grant columns follows `RunAction` -> `ProcessAddCapGrantRecords` -> `InsertCapGrantColumns` with a requested count of five.

Version 7.35 inserts Capital Grant records immediately before the workbook's hidden/template grant column. This keeps the insertion inside the workbook-owned Capital Grant ranges, so the assumptions names, all three Capital Grant Workings multi-area names, and existing Transactional DB `INDEX` source ranges expand together before `TransactionalDBSynchroniser` adds mirror rows. A read-only comparison of `TestFileClean.xlsb` and the failed validation copy isolated the former stale `$D$109:$H$148` reference; an unsaved insertion simulation expanded it to `$D$109:$M$148` and produced valid mirror formulas for records 1-10. Grid action extenders now force their DevExpress view layout before the selected section's best-size pass, ensuring the footer `+ Add lines` surface is included on initial display rather than only after a tab round-trip.

Version 7.36 applies the same workbook-owned template-boundary treatment to every `SpecificRowColumnEvents` rule where the authoritative XLSB requires it. Capital Expenditure and Housing Components now insert immediately before their trailing structural/template columns, so linked calculation ranges and Transactional DB source formulas expand with the user records; an unsaved read-only-workbook simulation expanded the relevant source spans from `OFA Additions!E:I` to `E:J` and `Component Totals!D:H` to `D:I`. Repairs already inserted within its dependent ranges, so only its assumptions-grid width source was corrected: formulas and formats still come from the zero-width template while new visible categories inherit the preceding visible column width. OFA, Funding, both Development rules, Journals and Stock Conversion were verified against their named-range geometry and deliberately left unchanged because their existing insertion points already preserve their formula dependencies or their append semantics.

### Funding Assumptions V2

Version 7.37 adds a workbook-backed `Funding Assumptions V2` child alongside the original interface. Its four tabs cover facilities and loans, variable and cash rates, other fees, and investments through the authoritative `Funding Assumptions` names/ranges. Funding Details now includes `Rep_Fund_08` (`Funding Assumptions!E81:R81`) as the final Opening Balance line and remains fixed at the top of the internally scrolling VGrid. Commitment Fees retains its calculation-method selector and now exposes the same in-place date-header editor and named-range row action as the earlier event sections. Native date-valued DIT fields use typed DevExpress date editors and continue to write through the established workbook change path.

Version 7.38 removes the historical assumption that a child structure's XML `CSID` must equal its zero-based list position. `GroupStructure.ResolveChildStructure` first preserves the fast positional path where ID and position agree, then resolves an explicit XML ID, with a final positional fallback only for legacy blank or non-numeric IDs. Presentation, datasource and group-caption lookups all use the resolver. This allows Funding Assumptions V2 (`CSID 138`) to coexist between Funding and Interco Funding without renumbering or redirecting any established Assumptions child.

### Workbook-backed grid clipboard support

Version 7.39 completes the DIT vertical-grid clipboard contract and extends it to the specialised FFR and Stress Test views. DIT vertical grids now use DevExpress cell multiselect and header-free native copy while retaining the existing transposed paste mapping and `ModelChangeManager` write path. FFR Workings, FFR Inputs Adjustment Statement and the editable Key Definitions Column C accept rectangular paste from the focused VGrid cell; locked/non-solid workbook cells and read-only statement views remain protected. The editable Stress Test first-tab, planner and target grids use their existing workbook-address resolvers for paste, while captured/output grids remain copy-only. Clipboard values are converted using the target workbook/data format and every accepted cell is processed through `ModelChangeManager`; no direct workbook assignment was introduced.

Version 7.40 unifies DIT clipboard entry points. The existing WindowsUI Copy and Paste buttons now act on the last focused DIT GridControl or VGridControl and call the same copy/custom-paste routines as Ctrl+C/Ctrl+V. Every DIT grid also receives a shared right-click Copy/Paste menu; hit testing focuses the clicked XtraGrid or VGrid cell before the menu opens, and an active DevExpress in-place editor is assigned the same menu so it cannot create a separate direct-bound paste route. Paste is disabled for read-only/live datasets and still performs the existing cell-level protection/rule checks before `ModelChangeManager` writes. Clipboard handlers and menu references are detached during section rebuild/disposal.

### Model-scoped change history and undo

Version 7.41 replaces the application-global V1 change display with ModelChangeManagerV2 and a model-scoped HistoryManagerV2. Successful interactive cell writes capture typed before/after workbook snapshots, including invariant formulas, and enter a linear per-model undo/redo journal only after calculation succeeds. Undo and redo validate that the current workbook cell still matches the expected snapshot, apply the whole change group in reverse/forward order, recalculate, and restore the group if any operation fails. New edits supersede the redo branch. DIT, FFR and Stress Test rectangular pastes are grouped as one history action; Ctrl+Z, Ctrl+Y and Ctrl+Shift+Z route to the active model, and the history window offers single-step and undo/redo-to-selected commands. Structural row/column operations remain deliberately non-automatic because safe reversal requires a complete domain-specific inverse across workbook-owned ranges and Transactional DB mirrors. The former change manager, history form and unused change-log manager are retained only as non-compiled text under Archive/History V1.

Version 7.42 makes DevExpress VGrid cell multiselect both consistently configured and visibly apparent. The shared VGrid setup enforces CellSelect mode and system selection colours, while DIT live/standard/Joint Venture and specialised FFR custom-draw paths reapply the selected-cell appearance after workbook-owned colours, fonts and number display have been sourced. This changes only the transient selection overlay; normal cells retain their XLSB appearance and edit protection.

Version 7.43 extends ModelChangeManagerV2 history to single-cell entities. DIT spin, text, combo and date controls now include blank-to-clear actions in the same typed journal used by grids, suppress dirty state during workbook refresh, and rely on central no-op filtering so calculation or undo refresh cannot manufacture duplicate history. The reusable ModelPosting text, date and combo controls and dormant extended-control posting hooks no longer assign workbook cells directly: committed values, including Nothing for a blank, flow through a shared single-cell posting service. ModelPosting controls refresh themselves after model-scoped undo/redo and restore workbook values after failed commits. Text and date editors commit on validation rather than on every keystroke; discrete combo selections retain immediate commit behavior.

Version 7.44 makes model undo/redo reliable while a DevExpress single-cell editor has focus. The DIT message filter now intercepts Ctrl+Z, Ctrl+Y and Ctrl+Shift+Z before the embedded editor can consume them as local text undo commands, while limiting handling to controls owned by that interface; reusable ModelPosting text, date and combo editors perform the same routing when hosted outside DIT. HistoryManagerV2 also exposes a fixed right-hand action column: applied rows show an undo icon, undone rows show a redo icon, and clicking an older row deliberately unwinds or reapplies the linear journal through that action. Superseded rows remain non-interactive.

Version 7.45 corrects the Global Assumptions single-cell undo path. Its controls are AbovoDETextEdit/DateEdit derivatives rather than ModelPosting editors, so those extended controls now route Ctrl+Z/Ctrl+Y directly to model history. DIT history refresh now reloads editable single-cell text, date, spin and combo controls as well as grids and read-only controls. A scoped suppression flag prevents the resulting EditValueChanged events from reposting the restored workbook value as a new action.

Version 7.46 corrects HistoryManagerV2 row undo/redo actions. Mouse hit-testing captures the group represented by the clicked action cell before its in-place ButtonEdit opens, and the command is deferred until that editor has closed. HistoryChanged notifications are suppressed while the manager is executing the requested stack transition, followed by one authoritative grid refresh; the row therefore changes from Undo/Applied to Redo/Undone (or back again) without an active editor restoring stale display data. Non-error no-action results are now reported rather than silently ignored.

Version 7.47 makes the restored workbook value propagate directly to live DIT single-cell editors. Writable text, date, spin and combo controls, including mapped-table text/combo editors, now hold a model-history refresh binding scoped to their source worksheet. After undo/redo each affected control reloads its own workbook cell under nested posting suppression, in addition to the containing DIT's general refresh. This removes reliance on the form-level deferred refresh reaching every dynamically-created editor and prevents refresh-generated EditValueChanged events from creating replacement history entries.

Version 7.48 makes the original DIT single-cell commit and its history record atomic from the user's perspective. The editor is marked clean and posting is suppressed while calculation performs its synchronous active-interface refresh, then the editor explicitly reloads the authoritative calculated workbook cell. ModelChangeManagerV2 now captures the history after-snapshot after calculation, so the Applied value, subsequent undo validation and visible editor cannot describe different workbook states.

Version 7.49 corrects the shared DIT undo/redo presentation refresh. History completion now refreshes its owning DIT synchronously after workbook snapshots, calculation and stack state are final. XtraGrid and VGrid controls backed by DevExpress UnboundSource reset that source at its existing row count before refreshing, invalidating values previously supplied through ValueNeeded; single-cell controls are reread in the same pass. The restored workbook state is therefore visible immediately across both entity editors and grids.

Version 7.50 closes refresh-time re-entry during undo and redo. DIT suppresses UnboundSource ValuePushed callbacks while XtraGrid and VGrid datasources are being reset or refreshed, and ModelChangeManagerV2 rejects any ordinary posting attempt while a history snapshot group is being applied. After calculation, every affected workbook cell is verified against its intended snapshot before the stack and log can report success; a re-overwritten cell now fails and rolls back the action instead of presenting a false Undone or Applied state.

Version 7.51 aligns workbook refresh with the DevExpress editor commit lifecycle. Single-cell Leave and grid/vertical-grid UnboundSource ValuePushed handlers now keep a posting-depth scope while ModelChangeManager writes and calculates. Any calculation-driven or explicit DIT refresh encountered in that scope is coalesced through BeginInvoke and runs only after DevExpress has completed the current edit message, preventing its pre-refresh editor buffer from restoring the previous visible value after a successful journalled change.

Version 7.52 separates model-history notification from the workbook transaction. HistoryChanged subscribers are invoked individually under exception isolation, so a stale, disposed or otherwise failing presentation subscriber cannot escape back into ProcessResolvedChange after the change group has been committed. This prevents the former inconsistent outcome where the history row remained Applied but the transaction catch restored the workbook's old value; one subscriber failure also no longer prevents the remaining DIT, grid and single-cell refresh subscribers from receiving the notification.

Version 7.53 adds keyboard progression for DIT XtraGrid defining-row editors. Pressing Enter in a ColumnInplaceEditorHelper validates and commits the active DevExpress header editor, then moves focus to the next visible column-header editor in the same BandedGridView after any synchronous workbook calculation or layout refresh. The final editor closes normally when no later header editor exists. Ordinary grid cell editors and VGrid row-header editors are unchanged.

Version 7.54 corrects ColumnInplaceEditor Enter traversal for banded headers. Navigation now uses each BandedGridColumn's source-order AbsoluteIndex rather than the unreliable band-relative VisibleIndex, and the temporary DevExpress editor explicitly disables EnterMoveNextControl so the form cannot transfer focus to the next grid while the helper is selecting the following header editor.

Version 7.55 preserves ColumnInplaceEditor traversal across workbook calculation and rule refresh. Successful combo/date header changes now queue one deduplicated advance request from the DIT change handler after UpdateAllRules; Enter without a changed value uses the same request. The helper closes any surviving current editor, forces the post-layout header paint to reacquire authoritative bounds, and focuses the next source-order header editor. This applies to accepted keyboard and popup/mouse edits in XtraGrid only.

Version 7.56 separates XtraGrid defining-row typing from commit. ColumnInplaceEditorHelper no longer posts on each RepositoryItem EditValueChanged event, allowing multi-character values such as 13 to be entered without a transient 1 being calculated and advancing focus. Enter and Tab validate, commit through the existing DIT ModelChangeManager handler, and move to the next header editor; Shift+Tab moves to the previous header editor. Genuine focus loss commits while respecting the control selected by the user. VGrid behavior is unchanged.

Version 7.57 normalises the in-column editor display sentinel `<Blank>` to Nothing before creating its DataChangeEvent. ModelChangeManager therefore clears the authoritative target cell through its established typed clear-contents path instead of storing the literal sentinel text. Successful UI state, history and rule refreshes receive the resulting blank value consistently; the rule applies to both XtraGrid and VGrid defining-row handlers that expose the sentinel.

Version 7.58 adds a most-recent-period fill action to DIT XtraGrid defining-row editors. Double-clicking a ColumnInplaceEditor copies a pre-write snapshot of the nearest populated defining column to its left into the editable cells beneath the selected column. Read-only, calculated, dummy, locked, spacer and control cells are skipped; all accepted writes use ModelChangeManager and appear as one grouped undo operation before the normal rules, calculations and interface refresh run. VGrid defining-row editors are unchanged.

Version 7.59 preserves the legacy grouped-sheet dependency semantics when Development identified or multi-year columns are expanded. DevExpress performs the seven worksheet insertions sequentially, so `Dvpt Component Depn` is now expanded before its dependent `Dvpt NonCash` sheet. This prevents freshly copied NonCash formulas from being shifted onto existing multi-year depreciation data by the later source-sheet insertion. A read-only comparison of the reported five-column failure showed that all fourteen Development Transactional DB mirrors had expanded correctly; the sole material error was 160 new NonCash formulas referring five columns to the right. Replacing only those references in a disposable copy and running Excel full calculation restored both the Statement of Financial Position and Transactional DB Check Sheet results to `OK`.

Version 7.60 completes the Other Fixed Asset semantic event conversion. The DIT add/delete dialog remains the sole prompt and its `RequestedRecordCount` now flows through the shared rule helpers into `OFA_RECORDS`; the legacy OFA handler no longer opens a second add-record dialog with its own default. Negative DIT adjustments also use the shared delete-last rule path, while direct selected-row commands retain the established selected-record fallback.

Version 7.61 adds an editable, context-sensitive Summit HTML help library. Every DataInterfaceTemplate now exposes a tagged Help WindowsUIButton that routes the current GroupStructure ID, ChildStructure ID, active tab and worksheet to an integrated modeless help viewer. The initial searchable reference covers all 217 configured interfaces and is generated from `Structure.xml`, the workbook User Guide and 777 legacy cell comments read from `Z:\Sandbox\TestFileClean.xlsb` with Excel macros/events disabled and without saving. Generated facts remain separate from `Help/data/overrides.js`, where durable client-authored HTML can be maintained without being overwritten by regeneration. The Help folder is copied to both application configurations, and the viewer also provides a direct Open Help Folder action.

Version 7.62 routes the existing bottom-left Help button on `FormMainScreen` to the local Summit HTML help home page through the shared `HelpManager`. This replaces the obsolete external website launch while retaining the established `GetHelp` WindowsUIButton tag and shared modeless viewer lifecycle.

Test Release 1.00 resets the client-facing Summit version and configures a repeatable offline ClickOnce package for initial functional testing. The separate installed product identity is `Abovo Summit Test`, the application/deployment version is `1.0.0.0`, automatic updates are disabled, and the bootstrapper prerequisite now matches the application target of .NET Framework 4.8. `Structure.xml` is explicitly included at the installed application startup root and the complete editable `Help` tree is explicitly included beneath it. `Tools/Publish-ClientTest.ps1` refuses to overwrite an existing versioned destination, invokes the project ClickOnce target, validates the required staged files and application-manifest entries, writes client installation notes, and produces the delivery ZIP. The available legacy test certificate is not installed in the certificate store, so this initial package is deliberately unsigned and the client README explains the expected Windows `Unknown Publisher` warning.

Test Version 1.01 restores the first `Funder` tab to both Funding Assumptions V1 (CSID 33) and Funding Assumptions V2 (CSID 138). Each tab presents the workbook's vertical `Funders` and `Facility` named ranges as one editable, non-pivoted `MergeAcross` grid. The `Funders` range is the NRRI expansion master, so an Add Lines operation inserts complete worksheet rows and keeps the aligned `Facility` range synchronized. Read-only inspection confirmed both contract ranges contain seven rows on `Funding Assumptions` (`B35:B41` and `D35:D41`), and Debug/Release builds both passed with the regenerated `Structure.xml` copied exactly to their startup folders.

Test Version 1.02 assigns the visually first row of both Funding Details vertical grids a dynamic combo editor backed by the workbook `Funders` named range. The editor uses the existing `Rep_Funders` resolver, so its choices are reread from the authoritative workbook rather than duplicated in application code.

Completed automated/static validation:

- Exact source/library hash comparison
- Read-only Excel workbook/name/range checks
- All canonical Transactional DB mirror geometries
- Capital Grant assumptions/workings/Transactional DB five-column expansion geometry
- Disposable five-column Development repair recalculated in Excel with Financial Position and Transactional DB checks both returning OK
- V1/V2 datasource field and period contract review
- Debug and Release MSBuild passed with zero build errors on 24 August 2026
- Generated Summit Help data parsed with 3 groups and 217 interfaces; a contextual Global Assumptions page rendered successfully through a hidden WinForms WebBrowser with its interface and field guidance present
- Release ClickOnce publish produced `Abovo Summit Test` version `1.0.0.0` with 89 files; exact SHA-256 comparisons confirmed the staged `Structure.xml` and all six Help files match their repository sources
- Application manifest contains the startup-root `Structure.xml` and nested Help paths; the delivery ZIP contains `setup.exe`, deployment/application manifests, client test README and all staged application files
- Client delivery archive `Z:\Sandbox\Deploy\Abovo-Summit-Test-1.00-ClickOnce.zip` SHA-256 is `63691F5F0C5CEDE7FC8DF31CF7E0DD32BBF734A4034935A3E3EF3EA41C504219`; Windows PowerShell reflection confirmed the compiled UI version text is exactly `1.00`

Required manual integration validation remains:

- Open Funding Assumptions V1 and V2, confirm `Funder` is Tab #1, edit both columns, add rows, and verify `Funders` and `Facility` remain aligned and feed the Funding Details dropdowns after calculation and reopen
- Open several Assumptions, Workings and Outputs DIT instances, select different tabs, and confirm Help opens the matching interface and scrolls to the active section; edit one `Help/data/overrides.js` entry and confirm it takes precedence after reopening Help

- Open Funding Assumptions V2, verify all four tabs, confirm Funding Details and its Opening Balance line remain fixed while scrolling, and exercise a Commitment Fees date-header add/edit/remove round trip
- Open the TestFileClean master in Summit and open Analysis V1 and V2 in both orders
- Add/remove a record in each structural domain and verify `Transactional DB` mirrors and analyser refresh
- In particular, add five Capital Grant records and verify both `Capital Grant Assumptions` and `Capital Grant Workings` expand by five records
- Verify Capital Expenditure, OFA, Repairs, Housing Components, Development Details, Development Multi Year, Journals and Stock Conversion use their declared semantic rules rather than the generic single-range path
- Add five Development Details records from the clean master and confirm the Financial Position and Transactional DB Check Sheet rows remain OK; repeat for Development Multi Year
- Close a passing model and confirm the normal Save/Discard/Cancel flow follows the completed full calculation
- Trigger a Check Sheet failure, confirm Cancel returns to the open model, and confirm OK requires a differently named XLSB copy before close
- Save a copy in Summit, open/save/recalculate it in Microsoft Excel/VBA, then reopen it in Summit
- Exercise Stress Test scenario generation/import and sensitivity expansion on a disposable copy
- Paste mixed text, numeric, percentage and blank rectangles into an editable DIT vertical grid and verify visual transposition, locked-cell skipping, recalculation and Excel round-trip
- Repeat paste checks in FFR Workings, FFR Inputs Adjustment Statement, FFR Key Definitions and each editable Stress Test grid; verify the read-only FFR/Stress output grids remain copy-only
- In a DIT XtraGrid and VGrid, verify toolbar, keyboard and right-click Copy/Paste produce identical results; repeat with an in-place editor open and confirm right-click targets the clicked cell and read-only/live grids expose Copy with Paste disabled
- Edit text, numeric, date, percentage, blank and formula-backed cells, then verify Ctrl+Z/Ctrl+Y and the History V2 buttons restore exact workbook values and displayed formulas after recalculation
- Paste rectangles in DIT, FFR and Stress Test and confirm each paste is one undo group; create a new edit after undo and confirm the discarded redo branch is marked superseded
- Modify a journalled cell externally before undo and confirm conflict detection refuses to overwrite it; confirm structural add/remove commands are not presented as automatically reversible

## Contract XLSB health and calculation audit - 24 August 2026

The authoritative `Z:\Sandbox\TestFileClean.xlsb` and repository compatibility copy remain byte-for-byte identical at SHA-256 `B248B1C733E1E3293536FBE1DBC9576D56FD4D34BEB74A3898B7F3D7333BCFBE`. The audit inspected an exact disposable copy through Excel with macros and events disabled and made no workbook or code remediation.

The workbook is structurally healthy and natively fast on the audit workstation: 638,360 formula cells completed a full dependency rebuild in 2.152 seconds and warm full calculations averaged about 0.42 seconds. No broken defined names, external formula references, workbook connections, circular reference, or true full-column calculation formulas were found. Broad formula replacement is therefore not justified.

The primary confirmed integration issue is the Summit `PMCOST` parameter contract. The XLSB deliberately calls eight compatibility positions, but only seven inputs affect the result because `FinalYear` is unused. Summit accepts positions 0 through 7 while exposing seven `ParameterInfo` entries. Preserve the eight-position workbook signature, document the unused position, correct the metadata, and require Excel/VBA-versus-Summit parity across all 380 `PMCOST` and 380 `RESPCOST` calls before changing evaluation logic.

The main workbook robustness risk is positional dependency: 35,920 formulas use 3-D worksheet references. Structural commands must validate the exact membership and order of every bounded sheet block. Multi-area names, used-range cleanup candidates, volatile `CELL` formulas and the single unguarded `Hidden - Tenure Totals Start!B6` lookup are targeted review items, not evidence for a wholesale redesign.

Future workbook, calculation-engine, Transactional DB, custom-function, structural-range and validation work must first review:

- `Library/Contract_XLSB_Audit_2026-08-24.pdf`
- `Library/Contract_XLSB_Audit_2026-08-24.md`
- `Library/Contract_XLSB_Audit_Evidence_2026-08-24.json`
