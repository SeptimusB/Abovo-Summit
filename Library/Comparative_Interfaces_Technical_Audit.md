# Comparative interfaces technical audit

Audit date: 19 August 2026

## Authority and scope

The implementation authority is `Library/TestFileMigrated.xlsb`. The visible
`Comparative` and `Comparative 2` worksheets, their linked form controls,
chart series, support formulas, validation and formatting were inspected
read-only. The source-free VBA inventory was also checked. Neither comparative
worksheet has a dedicated VBA event procedure; their behaviour is driven by
linked cells and formulas on `OW - Covenant Calculation`.

## Shared scenario and series contract

`Comparative!C12`, `C14`, `C16`, and `C18` are the four unlocked validated
scenario selectors. They use `R_StressTestNames`, feed the support calculation
ranges, and therefore drive charts and extrema on both comparison worksheets.

Six form checkboxes link to `OW - Covenant Calculation!BL17:BQ17` for Target,
Base Case, and Comparison 1-4. All six cells are unlocked Booleans. Every one
of the five chart families tests these same cells and returns `NA()` when a
series is disabled, so visibility is a single shared setting across Tabs #5
and #6.

## Comparative / Tab #5

The worksheet used range is `A1:Y49`. It contains two line charts and two
independent date-window controls:

| Metric | Offset cell | Argument | Series |
| --- | --- | --- | --- |
| Debt | `C10` | `BK19:BK38` | `BL19:BQ38` |
| EBITDA MRI | `C11` | `BS19:BS38` | `BT19:BY38` |

Both offset cells are unlocked integers from 0 to 20. The result table has
Target, Base Case and four selected scenarios. Debt extrema are exposed in
rows 8-18 (`G:L`); EBITDA MRI extrema are exposed in rows 36-46 (`N:R`).

## Comparative 2 / Tab #6

The worksheet used range is `A1:Y55`. It contains three line charts and three
independent date-window controls:

| Metric | Offset cell | Argument | Series |
| --- | --- | --- | --- |
| Gearing | `C12` | `CE19:CE38` | `CF19:CK38` |
| Debt / Unit | `C13` | `CM19:CM38` | `CN19:CS38` |
| Operating Margin | `C14` | `CU19:CU38` | `CV19:DA38` |

The combined extrema table is exposed in rows 35-45 (`P:Y`) and reports
maximum gearing, minimum operating margin, maximum debt per unit, and their
years for Target, Base Case and Comparison 1-4.

## Native implementation (versions 6.68-6.70)

Tabs #5 and #6 do not host a spreadsheet control. Each uses a native scenario
selection area and a shared set of `CheckEdit` series selectors. Every metric
is presented in a `GroupControl` containing its own 0-20 `SpinEdit` window and
`ChartControl`. Tab #5 uses the workbook's two-chart layout with a full-width
extrema grid. Tab #6 uses a two-by-two dashboard matching the worksheet's
three-chart-plus-summary arrangement.

Charts consume the workbook's calculated support ranges rather than
reimplementing their formulas. Their six colours match the workbook exactly:
Target `RGB(0,161,193)`, Base Case `RGB(0,91,130)`, followed by
`RGB(255,88,0)`, `RGB(122,184,0)`, `RGB(202,0,93)`, and
`RGB(240,171,0)` for Comparison 1-4. Summary text remains the workbook's
`DisplayText`, preserving its percentage, currency and year formats.

Version 6.69 gives each series selector an explicit caption width, removes the
duplicated legends beneath every chart, and uses the Scenario text in the first
summary column as the legend. Those labels use the same exact six-colour palette
as the chart series, so Tabs #5 and #6 retain a compact and consistent key.

Version 6.70 replaces the five start-year `SpinEdit` controls with 0-20
`TrackBarControl` sliders. Mouse changes are committed on release and keyboard
changes on key release, preventing calculation on every intermediate drag step;
the linked workbook cells and `ModelChangeManager` route remain unchanged.

Scenario, visibility, and window changes use the existing
`ModelChangeManager` route, calculate through the established Stress Test
service, and then refresh both tabs from the workbook. No chart calculation or
scenario result is maintained in a parallel UI model.

## Validation still required with client data

The demonstration workbook contains predominantly zero captured-scenario
values. Client testing should confirm line separation, extrema, series
visibility and each independent date window against populated scenario data.
