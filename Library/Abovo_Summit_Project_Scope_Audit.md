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
