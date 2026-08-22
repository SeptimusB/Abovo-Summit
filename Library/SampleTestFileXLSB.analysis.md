# SampleTestFileXLSB authoritative-master analysis

Analysis date: 22 August 2026

## Identity

- Authoritative source: `Z:\Sandbox\SampleTestFileXLSB.xlsb`
- Repository compatibility copy: `Library/TestFileMigrated.xlsb`
- Size: 11,691,134 bytes
- SHA-256: `571E52B1815F0E441A046951DECD4DCCE25C04D35BC3B0BBF05290957D62C7B4`
- Source and repository-copy hashes match exactly.

## Verified workbook contract

- 284 worksheets
- 1,766 defined names
- `Transactional DB!A1:BZ1909` used range
- `Transactional_Records = Transactional DB!A6:BV1599`
- 74 datasource columns and 40 period columns (`2024/25` through `2063/64`)
- No `__Abovo_*` worksheets or defined names
- All 89 canonical `TransCopy_*` mirror targets are on `Transactional DB` and match their source-derived row counts

The workbook was opened read-only through Excel automation with macros/events disabled and closed without saving. No production workbook was modified.

## Summit behavior

Summit treats this original formula/name layout as the contract. It no longer applies workbook schema migrations or creates materialised mirror sheets. Structural changes are mirrored through `TransactionalDBSynchroniser`, and synchronization errors are returned to the initiating structural command.

Analysis V1 and V2 both bind `Transactional_Records`. The range contains the original 74 workbook columns; the displayed `ItemDesc` field is an unbound grid column constructed by each analyser.

