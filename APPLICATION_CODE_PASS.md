# Abovo Summit current application code pass

Date: 22 August 2026

## Baseline

`Library/Blank BP v26_0001.xlsb` is the authoritative unpopulated workbook master. `Library/Demo BP v26_0001.xlsb` is the authoritative pre-populated master used by Debug auto-open. `Library/TestFileClean.xlsb` remains available only as the preceding contract baseline.

## Current integration

- Full-model open validates and loads the workbook without schema migration.
- `TransactionalDBSynchroniser` resizes the workbook's original formula-backed `TransCopy_*` named ranges.
- Structural rule and legacy structural operations propagate synchronizer failures.
- Analysis V1 and V2 share exclusive registry ownership of the canonical `Transactional_Records` RangeDataSource.
- Debug auto-open resolves `Library/Demo BP v26_0001.xlsb` from the repository root, including AnyCPU, x86 and x64 Debug output folders.
- Model close performs a full recursive calculation, examines `Outputs_CheckSheet`, and requires a separate XLSB Save As when validation does not pass.

The former schema-10 migration/materialisation assessment is retained in `APPLICATION_CODE_PASS.Schema10.md` as historical evidence only.

## Validation

See `Library/Abovo_Summit_Project_Scope_Audit.md` for the completed contract checks and required manual round-trip tests.
