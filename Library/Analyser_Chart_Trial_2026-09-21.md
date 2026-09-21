# Analyser line-chart trial - 2.50

## Use

The analyser main WindowsUI strip has an icon-only chart/grid toggle, separated
from its neighbours. It applies to the selected statement (SOCI, Detailed
Cashflow or Statement of Financial Position). The icon and tooltip offer the
alternate view. Figures remain the default; the former in-chart toggle is removed.
The chart starts with the top-level activity categories. `Include statement
totals` optionally includes the workbook's subtotal/carry-forward groups.

Right-click a line/point or its legend entry for `Go Up` and `Drill down to
[series name]`. Choosing the latter replaces the chart with that group's
immediate children. Go Up is disabled at Overview; drill is disabled on empty
space and terminal records. This continues through the same grouping fields as
the grid, ending at individual data records. `Up` (or Backspace while the chart
has focus) returns one level; `Overview` returns to the top. The dropdown and
`Drill down` provide an alternative for overlapping lines, including an entirely
zero-valued blank model. Non-expandable dummy/total groups explain that no
further detail is available. No chart copy command was introduced.

The existing Live, Snapshot and Differences source buttons apply to both views.
The chart breadcrumb always identifies the statement, source mode and branch.
Each statement retains its own chart/figures choice and navigation path for the
lifetime of the retained analyser instance. This state is not written into XLSB.

Mouse-wheel zoom is enabled on both axes. Ctrl + left-drag pans the zoomed plot;
right-click remains reserved for the drill menu. `Reset zoom` restores the full
view without changing the selected branch or source. Native zoom reference:
https://docs.devexpress.com/CoreLibraries/DevExpress.XtraCharts.ZoomingOptions.UseMouseWheel?v=25.2

## Workbook-first implementation

- `BPIncomeExpenditureAnalyserV2.Chart.vb` is a partial UI implementation over the
  existing `CustomGridView`. It does not create a RangeDataSource, write cells,
  insert metadata, edit XML definitions, or invoke workbook calculation.
- Groups use the existing grid's sum summaries. Leaf records use its cell values.
  No independent summing of headings/dummies can double-count the statement.
- Periods use the existing analyser discovery method, in workbook column order;
  the balance sheet also includes its opening balance.
- Values retain signs and grid units. No absolute-value, currency conversion or
  implicit scaling is applied. Non-numeric/non-finite values are chart gaps,
  with an unavailable-value count, not invented zero values.
- TitleLevel and AmDummy retain the grid's drill restrictions. Zero-valued named
  groups are not filtered out. Blank grouping captions are explicitly labelled
  `(Unclassified)`, without inventing mappings absent from the grid/workbook.
- Initial chart rendering is lazy. Reconnect invalidates all chart views but
  only rebuilds the visible chart. A hidden chart rebuilds when selected.
- Group value paths, not row handles, survive rebinding. If a branch vanishes,
  the chart returns to its nearest surviving expandable ancestor or Overview.
- Disconnect immediately clears the plotted series. The existing structural
  deferral notice disables the chart as well as figures until explicit refresh.
- Grid selection/expansion is not used to navigate the chart and is not changed
  by drilling. Existing datasource lifecycle and grid-state restoration remain
  responsible for calculation/edit/snapshot refresh.
- Native DevExpress 25.2 ChartControl hit testing is used, with visible line
  markers and line thickness increased from 2 to 3. Reference: https://docs.devexpress.com/WindowsForms/120444/controls-and-libraries/chart-control/end-user-features/hit-information
- The context menu uses native DXPopupMenu; source changes, tab hiding and chart
  disposal dismiss the menu. Reference: https://docs.devexpress.com/WindowsForms/DevExpress.Utils.Menu.DXPopupMenu?v=25.2
- Sidebar content-hosting accordion headers use navigator blue with white text
  in all interaction states, while retaining the compact 9.5pt user-scaled
  header font, content containers, compact body fonts and expansion states.

## Native regression evidence

Debug and Release compile. `Tools/Test-AnalyserChart.ps1` runs transparent native
controls on a private unsaved copy of the selected workbook, with SHA-256 checks
on the untouched original before and after. The fixture covers:

- Demo: 341 expandable group series and 505 terminal series, with point-by-point
  parity against existing group summaries/data cells across all three statements.
- Blank: 23 expandable groups and 107 terminal series; required Rental Income
  headings remain accessible with zero values.
- Actual right-click legend hit testing, Figures/chart switching, grid expansion
  preservation, and identical RangeDataSource identity through chart navigation.
- 2.48 rerun: main icon/tooltip and separator, thicker series, per-tab view state,
  real native popup opening/dismissal, enabled/disabled menu actions and menu
  navigation passed on Release Demo and Debug Blank.
- All three sources on all statements, source labels and retained branch paths.
- Snapshot creation; initially zero differences; a typed rent edit through
  ChangeManager with a second active assumptions registration; frozen Snapshot,
  changed populated Live figures, and Live-minus-Snapshot agreement in all 40 years.
- Reconnect, deferred structural refresh clearing/disabling the old chart, and
  explicit refresh restoring it. Layout bounds at 1100, 1900 and 3000 pixels.
- Native chart exports and analyser control renders under the ignored
  `obj/AnalyserChartTests/` fixture directory.

The masters currently expose an unnamed top-level balance-sheet grouping through
the existing analyser grid. The chart shows `(Unclassified)` and an explicit
source-label warning; this trial does not repair or invent the BS hierarchy.

The focused native sidebar fixture (`Test-PresentationLayout.ps1 -SidebarOnly`)
passed header colours, original content containers, expansion retention across
resizes and independent panel hide/restore. Header renders were inspected;
WinForms DrawToBitmap does not render the embedded browser bodies reliably.
The broader `-GroupOnly` fixture also flagged a restored DIT toolbar height of
285px at 1280px client width against its <200px expectation. This is outside
the chart/sidebar-header change and remains an explicit layout follow-up;
do not report the full presentation suite as passing. Physical client/5k layout
acceptance is still required.

Balance Sheet repair remains a separate plan, not part of the 2.48 implementation:
see `Analyser_Balance_Sheet_Repair_Plan_2026-09-21.md`.

## Client acceptance / risks

1. On the client's normal screen and the 5k display, maximise/restore and resize
   with chart visible. Check toolbar wrapping, axis labels, legend and breadcrumb.
2. In SOCI drill Rental Income -> Rents Receivable -> available stock/group detail
   -> record level. Try both lines and legend, and the dropdown for zero overlaps.
   Confirm Up/Overview and that figures retain widths, selections and expansions.
3. Create a snapshot, edit assumptions in a side-by-side DIT, and check each source
   while deep in the same branch. Undo/redo and verify matching grid/chart figures.
4. Add/remove supported structural lines while the chart is open. Confirm stale
   plots disappear, the existing refresh notice is respected, and removed branches
   fall back safely. Actual structural mutation remains a manual acceptance test.
5. Switch tabs, source modes, models and retained hidden/visible analyser instances.
   Check that state does not leak between statements or model instances.
6. Many coincident lines are inherently difficult to select; use the legend/list.
   Very wide detail levels can be visually busy; no arbitrary data truncation is
   applied. A chart cannot supply hierarchy missing from the underlying grid.
7. Normal Summit save -> Excel/VBA -> Summit remains a release acceptance check;
   this presentation-only trial introduces no workbook/package writes.

No production workbook, source master, Structure.xml or calculation engine was
modified. The pre-existing user project-file designer subtype edit was preserved.

## 2.49 sidebar correction and wheel-zoom follow-up

The user's blank-sidebar report was reproduced in native browser DOM reads:
BP Status, Funding Status and File Details could finish startup as empty HTML
documents. The preceding header-only appearance fixture missed this; blank
renders were not sufficient evidence to dismiss a content failure. Repeated
asynchronous DocumentText assignments during construction/presentation setup
could leave the browser at about:blank. A refresh after startup restored content.

Each sidebar browser now retains the latest generated HTML until its handle and
document are ready. The initial about:blank load is serialised; subsequent writes
use a synchronous MSHTML open/write/close stream, preserving scroll position and
skipping unchanged content. Handle recreation schedules restoration. Header
styling is scoped to headers, not copied control-wide LookAndFeel. The DOM tests
now require nonempty BP/Funding/File content on initial open, Complete ready state
after repeated refresh and resize, latest-content-wins updates, normal restoration,
panel hide/restore and a delayed second group window.

Sidebar accordion animation/smooth scrolling is disabled, header appearance changes
are batched, and System Messages/Interface History updates suspend layout and grid
painting. Message refreshes skip unchanged lists. Neither grid repeatedly runs
BestFitColumns or resets configured widths while filling records. The focused
fixture exercises new long message/history records as well as repeat refresh,
checking unchanged column widths and accordion container heights.

Release native chart tests exercised actual wheel events on all three statement
plots, narrowing both visual ranges, plus Reset zoom and existing drill/menu,
source switching, edit/snapshot, deferred-refresh and grid parity regressions.
The original master hash stayed unchanged. Debug and Release builds passed.
Physical wheel routing over the live window, pixel-level smoothness on the 5k
monitor, and user-reported layout acceptance remain manual tests. The previously
documented narrow DIT toolbar issue and Balance Sheet repair plan are unchanged.

## 2.50 easier chart selection

Line thickness increases from 3 to 5 and point marker size from 4 to 8, using
native DevExpress line/marker rendering and hit testing. Right-click on a line,
point or legend entry opens the existing Go Up / Drill down to [series] menu;
left-click still does not drill. Wheel zoom and Ctrl-drag panning are unchanged.
Overlapping lines can still be selected through their legend entries or list.

Release Demo regression passed: thicker lines/larger markers on all statements,
a native series hit two pixels away from a populated point centre, left-click
leaving the branch unchanged, right-click menu navigation, zoom/reset, grid parity,
all source modes, edit/snapshot differences, deferred-refresh safety and sizing.
The rendered Rental Income chart was inspected. The original workbook hash was
unchanged. Debug and Release compilation passed; Debug was initially built to
obj/TestDelivery250Debug because the normal Debug executable was still running.
After the user closed Summit, the normal bin/Debug build passed and reflection
verified DecVersionNumber = 2.50 in the delivered executable.
Actual 5k mouse targeting and client visual acceptance remain manual checks.
