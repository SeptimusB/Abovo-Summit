# Funding Dashboard Technical Audit

## Authoritative workbook contract

`Library/TestFileClean.xlsb` was inspected through Microsoft Excel automation
with macros and events disabled, alerts suppressed, and the file opened
read-only. It was closed without saving. The SHA-256 before and after inspection
was `7039BD6E7BAE82C0269F1F9E56D4C28E17EED7599FFB263169798D214AABF044`.

The `Funding Dashboard` sheet is visible, unprotected, and has used range
`A1:W68` with print area `A7:X68`. It contains 31 formula cells, 10 constant
cells, 12 shapes and seven charts. The shapes are presentation objects; the four
interactive controls are unlocked validation-list cells rather than linked form
controls.

| Purpose | Cell/name | Workbook validation source |
| --- | --- | --- |
| Display by Funder | `D6` / `DisplayFunder` | `Funders2` = `Funding Assumptions!B34:B41` |
| Display Facility | `I6` / `DisplayFacility` | `Facility2` = `Funding Assumptions!D34:D41` |
| Covenant Selector | `R6` / `CovenantSelector` | `CovenantDashboard` = `Covenant Assumptions!D5:F5` |
| From Year | `V6` / `CovenantYearSelector` | literal list `1` through `21` |

Current populated list values are not treated as fixed: every native selector
reloads from the workbook validation contract when the dashboard refreshes.
Accepted changes use `ModelChangeManager`, which performs the typed write, logs
the change, marks the model dirty, calculates, and rolls back a failed write.

## Native presentation

`FundingDashboard.vb` is a native DevExpress/WinForms user control routed by
`Structure.xml` as Outputs `CSID 21`, `SpecialElement=FundingDashboard`, data
source `Funding_Dashboard`, with worksheet button target `Funding Dashboard`.

The native layout follows the workbook:

- `Funder: <selection> - Drawn vs Available` uses source columns `AD:AF`, rows
  10-49.
- `Loans by Funder` discovers live series headings and values from `I:P`, with
  Net Closing Debt from `Q`, rows 7 and 10-49.
- `Funder: <selection> - Loan Mix` uses `S:T`, rows 7 and 10-49.
- `Funder: <selection> - Rates` uses `Z:AB`, rows 7 and 10-49.
- The selected covenant table uses the calculated, formatted
  `Funding Dashboard!P10:Q24` block. Its chart follows the XLSB chart contract:
  displayed years remain aligned to that table while forecast, target and
  breach overlay are sourced from `OW - Charts Source Data!AI`, `AJ` and `AL`,
  rows 10-24.
- `Operating Margin` covenant status uses `AO:AP`, rows 10-49.
- `EBITDA MRI` covenant status uses `AR:AS`, rows 10-49.
- Status arguments remain linked to `OW - Live Covenant Calculation!A7:A46`.

The calculated source worksheets and dashboard worksheet are registered with
the established calculation engine. Completed calculations rebuild the selector
lists and all seven visualisations while the interface is open; disposal removes
the active-object registration. The covenant table is read-only, supports cell
multiselect and header-free clipboard copy, uses workbook `DisplayText`, and
draws its visible fill, font treatment and alignment from the corresponding
worksheet cells.

Version 7.19 prevents the selected-covenant helper fields from becoming visible
by disabling automatic population and constructing only Year, Value and
Forecast GridColumn objects. The renderer retrieves Target and Exceeded
directly from the underlying DataRowView.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors.

Version 7.18 restricts the selected-covenant XtraGrid to its three intended
visible fields: year, workbook DisplayText value and the custom visual. Target
and breach helper fields remain available to the renderer but have no visible
index or customization entry. Chart-card headers and content now occupy
separate layout rows, and the lower status marker series are vertically centred.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors.

Version 7.17 presents the selected-covenant output as one read-only XtraGrid
rather than a separately rendered table and chart. Its visible year and
percentage text comes from the Funding Dashboard worksheet; the aligned visual
cell draws forecast AI, target AJ and breach overlay AL from OW - Charts Source
Data. Cell multiselect, header-free copy and calculation-driven rebuilding are
retained.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors.

Version 7.16 represents the selected-covenant target and the two lower status
indicators as native DevExpress point series. This gives the required
marker-only appearance without the unsupported use of DashStyle.Empty on
ordinary series lines.
Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
warnings and zero errors.

## Required manual validation

- Open Outputs > Dashboards > Funding Dashboard and confirm the native layout
  and worksheet button target.
- Select every populated funder and facility and confirm all four left-hand
  charts and their titles update.
- Select each covenant and several first-year values and confirm the selected
  15-year table/chart updates.
- Change a relevant assumption while the dashboard remains open and confirm all
  visible charts and selector lists refresh after calculation.
- Add or rename a funder/facility through the supported model workflow and
  confirm its selector and chart series update without reopening the model.
- Verify covenant-table multiselect and clipboard copy.
- Save, reopen in Excel/VBA, then reopen in Summit and confirm formulas, names,
  validation lists, charts and selector values round-trip intact.
