# TestFileClean authoritative-master analysis

Analysis date: 22 August 2026; authority update: 4 September 2026

## Identity

- Authoritative source: `Z:\Sandbox\TestFileClean.xlsb`
- Repository copy: `Library/TestFileClean.xlsb`
- Size: 11,813,448 bytes
- SHA-256: `D2CA9A14B432C6914612917129F20446DDD33A9BD7740AFAC704A5C357790C17`
- Source and repository-copy hashes match exactly.

## Verified workbook contract

- 283 worksheets
- 1,758 defined names
- `Transactional_Records = Transactional DB!A6:BV1599`
- 74 datasource columns and 40 period columns (`2025/26` through `2064/65`)
- `Outputs_CheckSheet = Check Sheet!A9:H63`; column E contains the calculated status, column F the user message, and column H the destination worksheet
- Embedded VBA project present
- No `__Abovo_*` worksheets or defined names
- All 89 canonical `TransCopy_*` mirror targets are on `Transactional DB` and match their source-derived row counts

The workbook was opened read-only through Excel automation with macros/events disabled and closed without saving. No production workbook was modified.

## Transactional DB capacity decision - 4 September 2026

Following client review, reserved no-shift capacity is not part of the production workbook contract. The authoritative master has reverted to the pre-capacity layout: there are no `_Capacity` or `_Continuation` defined names, and Summit and Excel/VBA both use the established structural-shift behavior when mirror ranges expand or contract.

A separate proof-of-concept workbook is retained outside the repository as `Z:\Sandbox\TestFileClean - Capacity Aware VBA Candidate.xlsb`. It is not an authoritative master and its capacity-aware VBA path did not complete full runtime and round-trip validation before the feature was withdrawn.

## Summit behavior

Summit treats this original formula/name layout as the contract. It no longer applies workbook schema migrations or creates materialised mirror sheets. Structural changes are mirrored through `TransactionalDBSynchroniser`, and synchronization errors are returned to the initiating structural command.

Analysis V1 and V2 both bind `Transactional_Records`. The range contains the original 74 workbook columns; the displayed `ItemDesc` field is an unbound grid column constructed by each analyser.
