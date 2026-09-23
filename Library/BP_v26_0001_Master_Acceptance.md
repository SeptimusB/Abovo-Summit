# BP v26_0001 master acceptance

Initial acceptance date: 11 September 2026

## Approved Blank replacement — 23 September 2026

At the user's explicit request, `Library/Blank BP v26_0001.xlsb` is now an exact byte copy of `C:/Sandbox/BP v26_0001 - New Blank.xlsb` (11,677,521 bytes, SHA-256 `E05274ACD3D4821AC38F013AC45A7B57CE335BD524938D9F21E1B640D0CC5E90`). This supersedes the Blank fingerprint in the initial acceptance below. The Demo master and Debug auto-open selection are unchanged.

The source and repository copy were hash-verified after replacement. All 1,389 ZIP package entries were readable, including the XLSB workbook and embedded VBA. No engine was used to resave either master. The previous repository Blank remains recoverable in Git and in `obj/ApprovedMasterUpdates/2026-09-23-new-blank/Blank BP v26_0001 before replacement.xlsb`.

The earlier Stori three-way inspection established matching worksheet formula/constant contents and complete name inventories between these two blanks, but different VBA, style and chart/object parts. The New Blank includes `MenuSheet.Calculate` before `BPMenu` reads its parameters. The user confirms that this Excel menu sheet is not used by Summit. The whole approved artifact was adopted, not just its VBA. Historical audits below describe the original September 11 artifacts; they are not new financial or macro-execution certification of this replacement.

After replacement, Debug and Release builds passed (`obj/new-blank-Debug-build.log`, `obj/new-blank-Release-build.log`). `Tools/Test-IntegrityCompatibility.ps1 -Configuration Debug -Workbook 'Library/Blank BP v26_0001.xlsb' -UI` passed on a private copy: Summit load, a single linked red warning, same-plan Check Sheet tab routing, and removal of the warning on a fresh successful status. Evidence directory: `obj/IntegrityCompatibility/f41c9113959d4311bb176cbc8e211ea7`. The cached chart scan identifies 465 intentional gaps and retains the existing `OW - Covenant Calculation!AB61` lookup for review. This is not a claim of a completely error-free financial model or a full visual/client acceptance test. The master hash remained unchanged through testing.

## Authority

- Unpopulated master: `Library/Blank BP v26_0001.xlsb`
- Pre-populated master: `Library/Demo BP v26_0001.xlsb`
- Debug auto-load: repository `Library/Demo BP v26_0001.xlsb`
- Preceding baseline retained for historical comparison: `Library/TestFileClean.xlsb`

The repository files are exact byte copies of user-approved files in `C:\Sandbox`. The following fingerprints record the initial September 11 acceptance; the current Blank fingerprint is above.

| Role | Size | SHA-256 |
| --- | ---: | --- |
| Blank | 11,643,967 bytes | `0B4C06800FE998E8733A5D8EB9CDEFA04F517750275BB0CFDD532928900F5B79` |
| Demo | 12,617,838 bytes | `1ED79726D8D1129C699242C9E8B3E988AC920EA7108D3C0B105BF9B0BC8A5E9C` |

## Formula integrity

The audit used Excel read-only with macros and events disabled. Formula coverage was measured with `ISFORMULA` because `SpecialCells(xlCellTypeFormulas)` fails on protected worksheets and must not be interpreted as zero formulas.

- The generated source, FormGenRemoved copy and Demo contain 678,983 formulas with identical formula addresses and text.
- The Blank contains 678,969 formulas.
- Its fourteen removed formulas were unlocked dummy/input entries, not generated calculations: `Development BP Assumptions!S86:U86` and `Stock Condition Inputs!O9:O19`.
- All generated calculation formulas are therefore retained in both accepted masters.

## Formula Generation removal

Both accepted masters contain fourteen paired `FormGenRemoval 11/9/26` markers. `Category_Changes`, `InitialColumnCheck`, `AllColumns`, `Launch_AllColumns` and `CreateFormGenMenu` exit immediately, and normal startup and worksheet-change entry calls are inactive. The original implementation remains as reversible dormant source but cannot run through normal workbook paths.

## Summit compatibility checks

- XLSB file format and embedded VBA are present.
- Both workbooks have 285 worksheets and 1,761 defined names.
- Defined-name definitions match exactly and contain no broken or external references.
- No external workbook links were found.
- Critical Summit sheets and names are present, including `Transactional_Records` and `Outputs_CheckSheet`.
- Embedded custom XML fingerprints match the preceding authoritative workbook.

The repository has no automated workbook/UI integration suite. Before client release, manually verify Summit open, representative population, calculation, analyser outputs, save, close validation, Excel/VBA reopen and Summit reopen.
