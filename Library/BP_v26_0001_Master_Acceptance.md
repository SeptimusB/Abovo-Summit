# BP v26_0001 master acceptance

Acceptance date: 11 September 2026

## Authority

- Unpopulated master: `Library/Blank BP v26_0001.xlsb`
- Pre-populated master: `Library/Demo BP v26_0001.xlsb`
- Debug auto-load: repository `Library/Demo BP v26_0001.xlsb`
- Preceding baseline retained for historical comparison: `Library/TestFileClean.xlsb`

The repository files are exact byte copies of the user-approved files in `C:\Sandbox`.

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
