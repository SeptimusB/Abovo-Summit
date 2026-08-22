# Abovo Summit current application code pass

Date: 22 August 2026

## Baseline

`SampleTestFileXLSB.xlsb` is the authoritative workbook master. The repository compatibility copy at `Library/TestFileMigrated.xlsb` is byte-identical (SHA-256 `571E52B1815F0E441A046951DECD4DCCE25C04D35BC3B0BBF05290957D62C7B4`).

## Current integration

- Full-model open validates and loads the workbook without schema migration.
- `TransactionalDBSynchroniser` resizes the workbook's original formula-backed `TransCopy_*` named ranges.
- Structural rule and legacy structural operations propagate synchronizer failures.
- Analysis V1 and V2 share exclusive registry ownership of the canonical `Transactional_Records` RangeDataSource.
- Debug auto-open targets `Z:\Sandbox\SampleTestFileXLSB.xlsb`.

The former schema-10 migration/materialisation assessment is retained in `APPLICATION_CODE_PASS.Schema10.md` as historical evidence only.

## Validation

See `Library/Abovo_Summit_Project_Scope_Audit.md` for the completed contract checks and required manual round-trip tests.
