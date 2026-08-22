# Abovo Summit Sample-master scope audit

Audit date: 22 August 2026

## Current authority

`Z:\Sandbox\SampleTestFileXLSB.xlsb` is the authoritative workbook master. `Library/TestFileMigrated.xlsb` deliberately retains its repository filename for existing tooling, but its bytes are an exact copy of that master:

- Size: 11,691,134 bytes
- SHA-256: `571E52B1815F0E441A046951DECD4DCCE25C04D35BC3B0BBF05290957D62C7B4`
- Worksheets: 284
- Defined names: 1,766
- `Transactional_Records`: `Transactional DB!A6:BV1599` (1,594 rows by 74 columns)
- Period fields: 40, from `2024/25` to `2063/64`
- ChatGPT-added `__Abovo_*` worksheets/names: none

The source workbook was inspected read-only through Excel automation with macros and events disabled and was never saved. The repository copy hash was verified against the source immediately after copying.

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

Completed automated/static validation:

- Exact source/library hash comparison
- Read-only Excel workbook/name/range checks
- All canonical Transactional DB mirror geometries
- V1/V2 datasource field and period contract review
- Debug and Release MSBuild passed with zero build errors on 22 August 2026

Required manual integration validation remains:

- Open the Sample master in Summit and open Analysis V1 and V2 in both orders
- Add/remove a record in each structural domain and verify `Transactional DB` mirrors and analyser refresh
- Save a copy in Summit, open/save/recalculate it in Microsoft Excel/VBA, then reopen it in Summit
- Exercise Stress Test scenario generation/import and sensitivity expansion on a disposable copy
