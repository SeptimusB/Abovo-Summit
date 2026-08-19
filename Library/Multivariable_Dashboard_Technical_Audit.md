# Multivariable Dashboard technical audit

Audit date: 19 August 2026

## Authority and scope

The implementation authority is `TestFileMigrated.xlsb`, not the previous
contents of Stress Test Tab #4. The workbook was inspected read-only through
DevExpress Spreadsheet. Its VBA streams were extracted statically with
`oletools`; no macro was executed and the XLSB was not saved or modified.

## Worksheet structure

`Multivariable Dashboard` has a used range of `A1:X43`. Its visible
dashboard is a ten-row window over forty calculated years, rather than a
forty-year line-chart dashboard.

- `E6` is the sole worksheet data-validation editor. It is an unlocked list
  cell, is named `R_Stress`, and selects Base Case or one of the ten captured
  scenarios.
- Rows 9-11 are the header/control area. The five covenant blocks are Gearing,
  Operating Margin, EBITDA MRI, Debt / Unit, and Debt.
- Rows 12-21 show ten calculated years. The displayed window is controlled by
  `OW - Covenant Calculation!C6` (1-31). Values are in columns
  `F/I/L/O/R`; percentile-direction signals are in `G/J/M/P/S`.
- Row 22 is the legend for target state and upper/lower percentile signals.
- Rows 24-38 contain the selectable covenant-breach graph and Debt Information
  panel. `K30` is Maximum Unfunded and `K33` is its year.
- Rows 40-41 contain the additional-assumptions note/description.

The sheet contains 19 conditional-formatting rules, one list validation,
six charts, eight linked form controls, thirty shapes (including the charts
and controls), and two pictures.

## Workbook control state

The linked form controls delegate their state to
`OW - Covenant Calculation`:

| Workbook cell | Dashboard function | Values |
| --- | --- | --- |
| `C6` | First displayed year | 1-31 |
| `C8` / `R_SortCovenant` | Sort covenant | 1-5 metrics; 6 = Year |
| `C9` | Sort order | 1 = highest first; 2 = lowest first |
| `C15` / `R_GraphOffset` | Covenant shown in breach graph | 1-5 |
| `C16` / `R_SortStress` | Captured-scenario data offset | Formula from `R_Stress` |

The hidden calculation sheet sorts all forty years, applies the `C6` window,
and exposes the visible rows through `AF:AK`. The dashboard cells reference
those results. Scenario values remain authoritative in `S0Data` through
`S10Data`; targets remain authoritative in the Multivariable Planner
target/golden-rule area.

## Chart sources

Five combo charts use ten-row source blocks in
`OW - Covenant Calculation!AL:AZ`. Each metric has three columns:
within-target (black), breach (red), and target. The sixth chart uses
`AA19:AB58`, with `A19:A58` as the year argument, and changes when
`R_GraphOffset` changes.

## VBA behavior

`Stress_Testing_Module.StressTestAssumptionsCapture` is the upstream
producer. It prepares `S0Data`, iterates Tests 1-10, either imports a
separate plan or applies the stored planner assumptions to the live model,
recalculates, and writes the resulting `StressLiveInfo` values to the
matching `SxData` block. It finally activates Multivariable Dashboard and
recalculates.

`MetBreachSwitch` supplies five entry points:
`MetBreach_Asset`, `MetBreach_Interest`, `MetBreach_IE`,
`MetBreach_DPU`, and `MetBreach_Debt`. Each sets `R_GraphOffset` to
1-5, calls `ABVCalculate`, selects the dashboard, and recolours the five
worksheet selector shapes. The calculation cell and resulting chart data are
the functional contract; shape selection/recolouring is Excel presentation.

No dashboard macro writes scenario result data directly. Scenario capture and
dashboard presentation must therefore remain separate operations.

## Tab #4 implementation (version 6.65)

The version 6.64 reinterpretation was rejected because it retained the former
grid-plus-chart-wall concept. It has been removed, including the invented
toolbar, five detached charts, lower summary grid, and their dead construction
helpers.

Tab #4 now hosts the model's existing live `SpreadsheetControl` and activates
`Multivariable Dashboard` directly. This presents the workbook's actual
ten-year matrix, embedded mini-plots, legend, Debt Information panel, selector
shapes, selected covenant graph, outcome cards, validation, conditional
formatting, and layout without a second presentation model.

The live control is temporarily reparented from the standard model viewer and
is restored when Stress Test closes or returns Home. The dashboard view uses
150% zoom with worksheet headings and gridlines hidden. Unlocked dashboard
cell edits are reverted and replayed through `ModelChangeManager`, after
which the established Stress Test calculation path is invoked. This preserves
the workbook-first round trip while avoiding a second workbook copy.

The global Stress Test grid policy remains unchanged: sorting, filtering, and
grouping are disabled, and editor buttons are shown only on the focused row.

## Validation still required with client data

The demonstration workbook has predominantly zero scenario values. Client
testing should therefore confirm the imported worksheet controls behave as
expected in DevExpress, especially the linked year scrollbar, sort option
buttons, sort-order spin button, and five VBA-backed graph selector shapes.
Where DevExpress imports a form control as presentation only, a future native
overlay should write its existing linked workbook cell rather than duplicate
the workbook calculation.
