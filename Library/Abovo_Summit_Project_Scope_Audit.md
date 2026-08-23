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

Completed automated/static validation:

- Exact source/library hash comparison
- Read-only Excel workbook/name/range checks
- All canonical Transactional DB mirror geometries
- Capital Grant assumptions/workings/Transactional DB five-column expansion geometry
- V1/V2 datasource field and period contract review
- Debug and Release MSBuild passed with zero build errors on 23 August 2026

Required manual integration validation remains:

- Open Funding Assumptions V2, verify all four tabs, confirm Funding Details and its Opening Balance line remain fixed while scrolling, and exercise a Commitment Fees date-header add/edit/remove round trip
- Open the TestFileClean master in Summit and open Analysis V1 and V2 in both orders
- Add/remove a record in each structural domain and verify `Transactional DB` mirrors and analyser refresh
- In particular, add five Capital Grant records and verify both `Capital Grant Assumptions` and `Capital Grant Workings` expand by five records
- Verify Capital Expenditure, OFA, Repairs, Housing Components, Development Details, Development Multi Year, Journals and Stock Conversion use their declared semantic rules rather than the generic single-range path
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
