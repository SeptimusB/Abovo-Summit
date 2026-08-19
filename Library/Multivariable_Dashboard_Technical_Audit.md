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

## Native Tab #4 prototype (version 6.64)

The previous Tab #4 chart wall has been replaced by a workbook-driven
prototype:

- Scenario, first displayed year, sort metric/order, and breach-graph metric
  controls map to `E6`, `C6`, `C8`, `C9`, and `C15`.
- Control changes use the standard `ModelChangeManager` path and then invoke
  the Stress Test calculation strategy.
- The main banded grid reads the calculated ten-row dashboard cells, including
  the workbook percentile signals and resolved conditional formatting.
- Five compact plots read the workbook black/red/target source blocks.
- The lower breach timeline reads `AA:AB`; the summary reads scenario,
  displayed-year window, Maximum Unfunded, its year, and the workbook note.
- All Stress Test grids disable sorting/filtering/grouping and show editor
  buttons only on the focused row.

## Construction plan after client prototype review

1. Match the five embedded Excel combo charts precisely (bar orientation,
   series overlap, target marker geometry, axes, legend, and spacing).
2. Reproduce the row-22 legend and the five selected-metric button states,
   using `R_GraphOffset` as state rather than duplicating the VBA shape logic.
3. Refine scaling for common Summit window sizes and high-DPI displays.
4. Validate sort direction, year-window scrolling, percentile signals,
   target breaches, Maximum Unfunded, and scenario switching against a
   populated client workbook rather than the zero-value demonstration model.
5. Only after parity evidence, decide whether any remaining worksheet shapes
   add functional value that warrants a native equivalent.

The workbook calculation ranges, names, formats, and protection state remain
the source of truth throughout this plan.
