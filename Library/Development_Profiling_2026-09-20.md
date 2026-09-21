# Development Unit Profiling - 20 September 2026

Test release 2.45; Structure.xml revision 1752; Assumptions GSID 0 / CSID 28.

## Findings and correction

The three MergeDownAndPivot tabs use HouseTypeInID for the read-only house-type description, then Rep_DevBP_12a (units), Rep_DevBP_13 (build/on costs) or Rep_DevBP_16 (SHG). Each source profile has 95 rows: monthly periods 1-60 in years 1-5, followed by annual years 6-40. This deliberately non-uniform timeline is preserved.

The old XML had only 84 period fields: month 22 was missing and it stopped at year 30. Positional binding therefore labelled every source row after month 21 one period too late and omitted the final eleven source rows. All three definitions now include month 22 and years 31-40. The misleading description caption and purchase-date tooltip are replaced by House Type / Development Details. Existing types, edit rules, ranges, source addresses, final template-column exclusion and workbook values are unchanged.

The reported blank description area is not reproduced with populated source descriptions: all ten Demo descriptions render in the native DIT. Both the authoritative Blank and C:/Sandbox/BP v26_0001 - New Blank.xlsb have empty HouseTypeInID source cells (Development BP Assumptions!G10:R10). Their early monthly profile inputs are also empty. However, they retain two later unit entries, L333=7 and K334=6, belonging to years 20 and 21 respectively; those are not absent data and must remain visible in the correct annual columns. No placeholder descriptions or zero values have been manufactured. Unsaved changes in the user's running model were not inspected.

## Verification

- Debug and Release builds passed.
- Tools/Test-DevelopmentProfiling.ps1 loads a source workbook without saving and verifies the native DIT: all three grids, each of the 95 monthly/annual headings against worksheet columns C/D, every projected source address and value, and house-type descriptions. It checks source SHA-256 before/after and writes renders only under obj/DevelopmentProfilingTests.
- Native fixtures passed for the authoritative Demo (Debug), authoritative Blank (Release), and Sandbox New Blank (Release). Each grid has 11 records and 96 columns including the description. Blank cells stay blank. A native Demo render was visually inspected.
- Neither source master nor the Sandbox workbook was saved or modified. No calculation, serialization, data-binding implementation or layout policy changed.

Run with Windows PowerShell 5.1, for example:

```powershell
& 'C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe' -NoProfile -ExecutionPolicy Bypass -File './Tools/Test-DevelopmentProfiling.ps1'
```

## Manual acceptance still required

Restart Summit/reopen the model to load the revised XML. Check all three tabs around months 21/22/23, month 60/year 6, and years 30/31/40. In a disposable copy, enter a house type through Development Details, return to profiling, and confirm its description and subsequent edits refresh. Verify edit/undo at an annual boundary and year 40 lands in the matching spreadsheet cell. Retain the standard Summit-save / Excel-VBA-reopen / Summit-reopen check; this fixture does not perform workbook writes or Excel/VBA execution. Client 5k/mixed-DPI acceptance of the preceding presentation changes remains separate and outstanding.
