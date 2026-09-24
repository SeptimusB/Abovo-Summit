# Missing snapshot worksheets - 2.86

Status: **Ready to test** with Jon. No commit, client sign-off or engine replacement is implied.

## Narrow repair

The inspected implementation previously required `TDB Snapshot` and `TDB Comparison` to exist and threw when either was absent. User approval on 23 September explicitly authorises creating them when needed and placing them immediately after Transactional DB.

- Only explicit `TransactionalDBSnapshotManager.CreateSnapshotAndComparison` provisions sheets. Ordinary model open, HasValidSnapshot, HasPersistedSnapshot and invalidation do not create them.
- The order is `Transactional DB`, `TDB Snapshot`, `TDB Comparison`. Existing dedicated sheets are reused and, if misplaced, moved into this order during explicit capture. Other business sheets retain their relative order.
- Existing destination-content validation and source-period validation run before provisioning or clearing. Protected workbook structure prevents required creation/reordering with a clear error. Existing sheet protection is restored through the established service.
- New sheets use the existing values-only snapshot, live-minus-snapshot comparison, local range names and persistent Balance Sheet bundle. No financial formula algorithm, structural synchroniser, VBA or ordinary-load schema migration is added.
- A failed capture attempts to remove only the sheets created by that command and restore original sheet order; existing scratch outputs retain the established failed-capture invalidation policy. Successful capture marks the model dirty for normal Save. This is not an automatic write to the source file.
- Missing Balance Sheet snapshot guidance now points to Create Snapshot rather than requiring an upgraded template.

## Automated validation

Debug and Release build successfully with test version **2.86**. Diagnostics remain disabled in the Debug delivery; the quiet-build regression passes, including its explicit Check Sheet trial opt-in exception. The normal `bin/Debug` and `bin/Release` executables are updated.

`Tools/SnapshotCreationFixture.cs`, run by `Tools/Test-ClientReport.ps1`, operates only on a newly copied private workbook:

| Configuration / source | Result | Ignored evidence directory |
| --- | --- | --- |
| Release / Demo | 24 assertions passed | `obj/ClientReportTests/a7e7934129554bf1af977f983c2ad567` |
| Debug / older Stori CLEAN pass 3 | 24 assertions passed | `obj/ClientReportTests/846ce36497a349c28b7470d7c6790b97` |
| Debug / current Blank, expanded fixture | 29 assertions passed | `obj/ClientReportTests/20c02fe43a5c49588e2628ee11f9fdbc` |

Coverage: ordinary open/probes do not add sheets, either/both missing, reuse, misplaced sheets, repeated capture, exact pair order, dirty/valid state, collision rejection without adding the missing partner, unchanged business formula/constant/name/protection digest, unchanged relative business-sheet order, zero initial Balance Sheet differences, normal guarded XLSB Save and preservation of every frozen Balance Sheet value after native reopen. Expanded Blank checks also cover protected workbook structure and invalid period headers rejecting before mutation, and confirm the pair falls outside all 16 recognised business 3-D sheet spans. Faults after partial copy/EndUpdate remain handled by the existing cleanup path plus the new best-effort sheet cleanup; no injected post-copy exception test is claimed.

Independent `Tools/Test-BalanceSheetExcelRoundtrip.ps1` passes for Demo and Stori: Excel read-only open, macros/events/calculation disabled, SaveCopyAs to another disposable filename, then native reload. Complete frozen hierarchy/values, worksheet count/order and global names are preserved. Excel outputs:

- Demo: `obj/BalanceSheetExcelRoundtrip/17764d54e631484eb8604818d4ebb5a5/excel-roundtrip.xlsb`
- Stori: `obj/BalanceSheetExcelRoundtrip/f264f8e4e1a74c95aac776a600db30ac/excel-roundtrip.xlsb`

Read-only in-memory VBA source hashing confirms all 337 Demo and 333 Stori module identities/source hashes match the original through Summit and Excel outputs. Blank's 339 modules match its Summit output. Excel rewrites the VBA binary container; module-source identities, rather than whole-container byte equality, are the stated check. No macros were executed and no raw VBA/passwords were extracted to the repository. The master and supplied Stori source hashes remain unchanged.

## Manual acceptance

In an older file without snapshot sheets, choose Create Snapshot, switch Live / Snapshot / Differences and check SOCI, Cashflow and Balance Sheet figures/charts. Save a separate test copy, reopen it and repeat. Native UI interaction, trusted Excel/VBA execution and accountant financial acceptance are not replaced by the automated checks.

## Separate performance research

See `Spreadsheet_Engine_Trial_Assessment_2026-09-23.md`. Commercial Aspose.Cells is the first isolated trial candidate. No alternative engine has been installed, benchmarked or integrated, and no speed gain is claimed. DevExpress and the current interface remain unchanged.
