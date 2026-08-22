# TestFileClean authoritative-master analysis

Analysis date: 22 August 2026

## Identity

- Authoritative source: `Z:\Sandbox\TestFileClean.xlsb`
- Repository compatibility copy: `Library/TestFileMigrated.xlsb`
- Size: 11,841,039 bytes
- SHA-256: `B248B1C733E1E3293536FBE1DBC9576D56FD4D34BEB74A3898B7F3D7333BCFBE`
- Source and repository-copy hashes match exactly.

## Verified workbook contract

- 281 worksheets
- 1,758 defined names
- `Transactional DB!A1:BZ1909` used range
- `Transactional_Records = Transactional DB!A6:BV1599`
- 74 datasource columns and 40 period columns (`2025/26` through `2064/65`)
- `Outputs_CheckSheet = Check Sheet!A9:H63`; column E contains the calculated status, column F the user message, and column H the destination worksheet
- Embedded VBA project present
- No `__Abovo_*` worksheets or defined names
- All 89 canonical `TransCopy_*` mirror targets are on `Transactional DB` and match their source-derived row counts

The workbook was opened read-only through Excel automation with macros/events disabled and closed without saving. No production workbook was modified.

## Summit behavior

Summit treats this original formula/name layout as the contract. It no longer applies workbook schema migrations or creates materialised mirror sheets. Structural changes are mirrored through `TransactionalDBSynchroniser`, and synchronization errors are returned to the initiating structural command.

Analysis V1 and V2 both bind `Transactional_Records`. The range contains the original 74 workbook columns; the displayed `ItemDesc` field is an unbound grid column constructed by each analyser.
