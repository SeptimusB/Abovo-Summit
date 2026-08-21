# Outputs Interface Technical Audit

## Delivery approach

Outputs children 0 through 20 are implemented through Business Plan Dashboard.
The authoritative contract is `Library/TestFileMigrated.xlsb`; archived
structure files were used only to recover likely ordering and grouping before
each child was checked against the current workbook.

All worksheet-backed Output grids are detached read-only projections. Horizontal
`LiveGrid` is used for stock and traditional statement layouts. `LiveVGrid` is
used where measures are clearer down the page and workbook periods become
records across the page. Both refresh from the live workbook after calculation,
preserve cell multiselect and clipboard copy, and derive visible values and
formatting from the underlying workbook cells. `BP Dashboard` continues to use
the existing native `BP_Dashboard` interface.

## Child 0: Existing Stock Numbers

- Parent: `Outputs` (`GSID 2`)
- Navigation group: `Stock Numbers`
- Worksheet button target: `Existing Stock Numbers`
- Workbook defined name: `Outputs_ExistingStockNumbers`
- Current defined-name areas: `A12:W52`, `Y12:AS52`, and `AU12:BO52`
- Workbook headings: rows 6 and 10
- Workbook inspection: isolated read-only Excel automation with macros and
  events disabled; the XLSB timestamp was unchanged

The child is split into three tabs:

1. `Existing Stock Numbers`: `A12:W52`
2. `New Lettings Existing Stock Numbers`: row headings `A12:B52` plus
   `Y12:AS52`
3. `Total Existing Stock Numbers`: row headings `A12:B52` plus `AU12:BO52`

The two row-heading columns and the stock-type headings remain workbook driven,
so year, tenure, and stock-type formula changes are reflected when the grid is
refreshed.

## Children 1-20

- Stock Numbers: Development Stock Numbers and Stock Numbers use horizontal
  `LiveGrid` tabs split by the workbook's named-range areas.
- Cashflows: Existing Cashflows through Cashflow Detailed use `LiveVGrid`, with
  worksheet measures as rows and year/period records across the page.
- Traditional accounts: Summary Comprehensive Income, Detailed Comprehensive
  Income, Financial Position, and Cashflow Statement use `LiveGrid`.
- Alternative accounts: the corresponding four alternative views use
  `LiveVGrid`.
- Dashboard: BP Dashboard (`CSID 20`) uses the existing `BP_Dashboard` special
  interface and targets worksheet `BP Dashboard`.

The generated XML is reproducible with
`Tools/GenerateOutputsStructure.ps1 -Apply`. The generator replaces only the
`Outputs` group, preserves child 0, and emits the verified child order 0-20.

Version 7.10 passed full Debug and Release rebuilds on 21 August 2026 with zero
warnings and zero errors. XML parsing confirmed 21 sequential child IDs, all
worksheet targets and all 11 named sources were present in the read-only XLSB,
and the generated `Structure.xml` matched both build outputs.

## Required manual validation

- Open Outputs and exercise every child from Existing Stock Numbers through BP
  Dashboard.
- Switch repeatedly through all stock tabs and between horizontal and vertical
  Output children.
- Recalculate while the Output is visible beside editable assumptions and
  confirm values and headings refresh.
- Rename or add a stock type in the supported model workflow and confirm the
  refreshed headings follow the workbook.
- Use the worksheet button on every child and confirm the correct underlying
  XLSB worksheet opens, including the abbreviated cashflow sheet names.
- Confirm formula-driven measure/category headings refresh in every vertical
  view.
- Confirm cell multiselect and clipboard copy on every grid.
