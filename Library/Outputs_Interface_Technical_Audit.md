# Outputs Interface Technical Audit

## Delivery approach

Outputs children 0 through 33 are implemented through Analysis V2.
Scenario Planning is intentionally omitted because those interfaces are owned
by Stress Test.
The authoritative contract is `Library/TestFileMigrated.xlsb`; archived
structure files were used only to recover likely ordering and grouping before
each child was checked against the current workbook.

All worksheet-backed Output grids are detached read-only projections. Horizontal
`LiveGrid` preserves the source worksheet orientation for stock, cashflow and
statement layouts, including every traditional and alternative account view.
The grids refresh from the live workbook after
calculation, preserve cell multiselect and clipboard copy, and derive visible
values and formatting from the underlying workbook cells. `BP Dashboard` uses
the existing native `BP_Dashboard` interface and now rebuilds its workbook-backed
content after calculation while it remains open. `Funding Dashboard` is a
dedicated native interface whose selectors and seven visualisations remain
linked to their authoritative workbook cells and source ranges. Analysis V1
is a worksheet-free special child that opens the unchanged
BPIncomeExpenditureAnalyser class. Analysis V2 opens the independent corrected
BPIncomeExpenditureAnalyserV2 class. Neither declares a worksheet button target
or a worksheet-backed interface range.

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

## Children 1-21

- Stock Numbers: Development Stock Numbers and Stock Numbers use horizontal
  `LiveGrid` tabs split by the workbook's named-range areas.
- Cashflows: Existing Cashflows through Cashflow Detailed use horizontal
  `LiveGrid` and preserve their source worksheet row/column orientation.
- Traditional accounts: Summary Comprehensive Income, Detailed Comprehensive
  Income, Financial Position, and Cashflow Statement use `LiveGrid`.
- Alternative accounts: Summary Comprehensive Income, Detailed Comprehensive
  Income, Financial Position and Cashflow Statement all use horizontal
  `LiveGrid`, matching the stock interfaces and source worksheet orientation.
- Dashboard: BP Dashboard (`CSID 20`) uses the existing `BP_Dashboard` special
  interface and targets worksheet `BP Dashboard`.
- Dashboard: Funding Dashboard (`CSID 21`) uses the new `FundingDashboard`
  special interface and targets worksheet `Funding Dashboard`. Its exact
  workbook contract is recorded in `Funding_Dashboard_Technical_Audit.md`.

The generated XML is reproducible with
`Tools/GenerateOutputsStructure.ps1 -Apply`. The generator replaces only the
`Outputs` group, preserves child 0, and emits the verified child order 0-33
without Scenario Planning.

## Children 22-31

- Covenants and Other Reports: Covenants, Value for Money Metrics, Loan Output
  Table, Summary Stock Numbers, Surplus on Sales, NSH Breakdown, and Hsg
  Properties Mvmt.
- Other Cashflow Outputs: BP Input Scheme Cashflows, 5 Yr Monthly Cashflow,
  and 5 Yr Quarterly Cashflow.
- All ten children use read-only horizontal LiveGrid, retain workbook
  DisplayText and formatting, refresh after calculation, and target their
  identically named worksheet from the worksheet button.
- Nine children resolve the workbook's existing Outputs_* multi-area defined
  names. Value for Money Metrics uses its contiguous A9:M171 report block.
- Macro-disabled read-only Excel validation confirmed every worksheet and
  source range; the XLSB remained byte-identical at SHA-256
  7039BD6E7BAE82C0269F1F9E56D4C28E17EED7599FFB263169798D214AABF044.

Version 7.10 passed full Debug and Release rebuilds on 21 August 2026 with zero
warnings and zero errors. XML parsing confirmed 21 sequential child IDs, all
worksheet targets and all 11 named sources were present in the read-only XLSB,
and the generated `Structure.xml` matched both build outputs.

Version 7.11 suppresses fully blank projection rows and hides columns only when
both their workbook heading and every projected `DisplayText` value are blank.
LiveVGrid record-header helper fields are hidden after attachment so DevExpress
cannot reset their visibility and expose fallback names such as `Col_1`.
Meaningful unlabeled workbook data remains visible. Both LiveGrid orientations
continue to source cell text from `DisplayText`; numeric, leading-minus, and
parenthesised negative values are explicitly rendered red. Full Debug and
Release rebuilds passed on 22 August 2026 with zero warnings and zero errors.

Version 7.12 returns every cashflow-family child (`CSID 3-11`) and Alternative
Cashflow Statement (`CSID 19`) to horizontal `LiveGrid`, matching the stock
interfaces and workbook orientation. Output LiveGrids use the current monitor's
usable height so Summary Comp Inc - Trad View can show its compact report rather
than being constrained by the general 70% grid cap. An open BP Dashboard is now
an active calculation consumer: completed workbook calculations rebuild its
charts and gearing table from current workbook values, and disposal unregisters
it from the calculation engine. Full Debug and Release rebuilds passed on 22
August 2026 with zero warnings and zero errors.

Version 7.13 returns the remaining Alternative account views (`CSID 16-18`) to
horizontal `LiveGrid`; Outputs now contains no `LiveVGrid` elements. It adds the
native Funding Dashboard at `CSID 21`, including workbook-validation-driven
funder, facility, covenant and first-year selectors. Accepted selections travel
through `ModelChangeManager`; completed calculations rebuild all selector lists,
funding charts, the selected covenant table/chart and both covenant-status
charts. The read-only workbook inspection retained SHA-256
`7039BD6E7BAE82C0269F1F9E56D4C28E17EED7599FFB263169798D214AABF044`.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors.

Version 7.15 restores the two lower-right covenant status charts and aligns the
top-right selected-covenant chart with the XLSB series contract. Status values
are read from `AO:AP` and `AR:AS`; their workbook value of `1.1` is no longer
clipped by a `0-1` native axis and marker colour is no longer made transparent.
The selected chart now uses forecast `AI`, target `AJ` and breach overlay `AL`
rather than treating the display table's `Q` values as the complete chart
source.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors.

Version 7.20 adds Outputs children 22-31 under Covenants and Other Reports and
Other Cashflow Outputs. Each interface is a workbook-backed horizontal
LiveGrid using the verified source name/range and workbook header rows.
Scenario Planning is deliberately excluded because Stress Test already owns
those interfaces.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors. XML validation confirmed one Outputs group, 32
unique child IDs/names, exactly ten new children, and no Scenario Planning
group.

Version 7.21 adds the worksheet-free `Analysis` child at `CSID 32`. Its
`SpecialElement` routes directly through the existing
`BPIncomeExpenditureAnalyser` branch in `GroupInterfaceTemplate`; no interface
worksheet or default worksheet is declared, and the analyser class itself was
not modified. XML validation confirmed one Outputs group, 33 unique child IDs
and names, exactly one Analysis child, and no Scenario Planning group.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors.

Version 7.22 renames the unchanged original route to Analysis V1 and adds
Analysis V2 at CSID 33. V2 is an independent source/designer/resource set with
corrected Balance Sheet grouping and selection, dynamic named-field and period
discovery, Balance Sheet export, deterministic binding/calculation cleanup,
safe rendering resources and explicit read-only cell multiselect/copy. The
detailed changes and manual checks are recorded in
BP_Income_Expenditure_Analyser_V2_Technical_Audit.md.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors.

Version 7.23 gives the Transactional_Records RangeDataSource exclusive,
model-scoped ownership. Opening Analysis V1 or V2 releases the other analyser
in order before creating the replacement binding; the displaced object is also
removed from the calculation engine and document manager. The general
ModelResourceRegistry releases all remaining owners before model shutdown.
The focused registry behavior test and full standard Debug and Release rebuilds
passed on 22 August 2026 with zero warnings and zero errors.

Version 7.24 corrects V2 period discovery after the first runtime test exposed
a stripped regular-expression escape. Periods are now recognized from grid
field/caption metadata with an authoritative Transactional_Records header-row
fallback.
The built matcher behavior test and full standard Debug and Release rebuilds
passed with zero warnings and zero errors.

Version 7.25 prevents Analysis V2 group-summary metadata from being converted
directly to Int32. All active expansion, row styling and group custom-draw
paths now share a guarded numeric parser, so blank or text-valued legacy
summaries are ignored instead of aborting the interface open. Analysis V1 and
the authoritative workbook source remain unchanged.
The focused parser test and full standard Debug and Release rebuilds passed on
22 August 2026 with zero warnings and zero errors.

Version 7.19 disables automatic column population for the selected-covenant
XtraGrid and explicitly creates only Year, Value and Forecast visual columns.
Target and Exceeded remain fields in the bound row for custom rendering but no
longer have GridColumn instances that DevExpress can display.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors.

Version 7.18 removes the selected-covenant grid's Target and Exceeded helper
fields from both its visible layout and customization surface. Dashboard chart
cards now reserve a separate layout row for their title bars so legends cannot
be covered, and the two lower status charts use a fixed centred vertical range
for the workbook's 1.1 marker values.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors.

Version 7.17 replaces the top-right selected-covenant grid/chart split with a
single read-only DevExpress XtraGrid. Each workbook year and DisplayText value
shares its row with a native custom-drawn visual cell sourced from forecast AI,
target AJ and breach overlay AL. The blue/red data bar and green target marker
therefore remain aligned and refresh together as one workbook-backed dataset.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors.

Version 7.16 replaces the Funding Dashboard's marker-only line series with
native DevExpress point series. Target, covenant-met and covenant-breached
markers retain their workbook-derived data and colours without assigning
DashStyle.Empty to an ordinary series line, which DevExpress 25.2 rejects at
runtime.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors.

Version 7.14 corrects the native selected-covenant grid initialisation order.
Its custom `GridView` is now installed before the `DataSource` is assigned,
columns are populated explicitly, and workbook column styling is applied only
after guarded field lookup. This prevents the opening `NullReferenceException`
caused by DevExpress generating columns on its temporary default view.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors.

## Required manual validation

- Open Outputs and exercise every child from Existing Stock Numbers through
  5 Yr Quarterly Cashflow.
- Switch repeatedly through all stock, cashflow, traditional account and
  alternative account children and confirm their horizontal orientation.
- Recalculate while the Output is visible beside editable assumptions and
  confirm values and headings refresh.
- Rename or add a stock type in the supported model workflow and confirm the
  refreshed headings follow the workbook.
- Use the worksheet button on every child and confirm the correct underlying
  XLSB worksheet opens, including the abbreviated cashflow sheet names.
- Confirm formula-driven headings refresh in all views and that open BP and
  Funding Dashboards refresh after a substantive assumptions change and
  calculation.
- Exercise all four Funding Dashboard selectors and confirm the workbook cells,
  chart titles, series, selected covenant years/values and status markers update.
- Confirm cell multiselect and clipboard copy on every grid.
- Confirm the two new accordion groups contain exactly seven and three
  children respectively, with no Scenario Planning group.
- Confirm the ten new worksheet buttons open their identically named XLSB
  sheets and that all multi-area reports omit only the source spacer columns.
