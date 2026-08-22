# Abovo Summit current application code pass

Date: 22 August 2026

## Baseline

`TestFileClean.xlsb` is the authoritative workbook master. The repository compatibility copy at `Library/TestFileMigrated.xlsb` is byte-identical (SHA-256 `B248B1C733E1E3293536FBE1DBC9576D56FD4D34BEB74A3898B7F3D7333BCFBE`).

## Current integration

- Full-model open validates and loads the workbook without schema migration.
- `TransactionalDBSynchroniser` resizes the workbook's original formula-backed `TransCopy_*` named ranges.
- Structural rule and legacy structural operations propagate synchronizer failures.
- Analysis V1 and V2 share exclusive registry ownership of the canonical `Transactional_Records` RangeDataSource.
- Debug auto-open targets `Z:\Sandbox\TestFileClean.xlsb`.
- Model close performs a full recursive calculation, examines `Outputs_CheckSheet`, and requires a separate XLSB Save As when validation does not pass.

The former schema-10 migration/materialisation assessment is retained in `APPLICATION_CODE_PASS.Schema10.md` as historical evidence only.

## Validation

See `Library/Abovo_Summit_Project_Scope_Audit.md` for the completed contract checks and required manual round-trip tests.
