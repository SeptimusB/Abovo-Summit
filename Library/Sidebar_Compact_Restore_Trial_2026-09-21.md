# Sidebar compact/restore trial — 2.52

## Agreed behaviour

Compact closes the navigator and summary contents but retains both edge strips. Click a strip to reopen its panel; the panel stays open until explicitly closed. Restore returns each panel to its pre-compact open/closed state and dimensions. Reopening a panel while compact does not overwrite that saved state. The icon-only button and approved images are unchanged, and state is independent per group window (including DIT and analyser hosts).

## Cause and repair

The earlier implementation set the legacy navigator strip panel to `DockVisibility.Hidden`. DevExpress disposed its `AutoHideContainer`; restore created a different container while the `WithEvents` click handler remained attached to the old one. A closed navigator therefore had no working reopen strip. Native reproduction confirmed this for both initially open and initially closed navigators.

The new path hides only the navigator contents, retains its strip in AutoHide, and collapses the summary to AutoHide. It captures both current and original panel dimensions. `ReconnectPanelStrips` rebinds the actual current native containers, including the summary strip recreated after unpinning. Explicit clicks reopen/pin the summary just as they reopen the navigator; native hover expansions are cancelled so a pending hover request cannot undo a newer compact/restore action. Closing uses the native immediate-hide API; there is no custom animation or delayed reopen callback.

Supported native API references: [HideImmediately, 25.2](https://docs.devexpress.com/WindowsForms/DevExpress.XtraBars.Docking.DockPanel.HideImmediately?v=25.2), [Expanding](https://docs.devexpress.com/WindowsForms/DevExpress.XtraBars.Docking.DockPanel.Expanding), also checked against installed 25.2.4 XML documentation.

## Validation

- Debug and Release build successfully; normal executables contain test version 2.52.
- Native regression exercises all four initially open/closed panel combinations, three cycles each; both strips remain visible and clickable, contents close, previous visibility/width is restored, and a second group window is unchanged.
- Both individual close arrows, replacement-container event reattachment, queued hover versus restore, icon-only state sharing with newly attached buttons, and unchanged approved SVG choices pass.
- Debug `Tools/Test-PresentationLayout.ps1 -Configuration Debug -SidebarOnly` passes the panel regression plus sidebar HTML content, accordion expansion, grid layout stability, repeated refresh and resize checks (1280–5000px widths). Evidence: `obj/PresentationLayoutTests/3c0c8b08e87641a587be44df85ac93bb` and `obj/sidebar-252-sidebar-test.log`.
- Release panel fixture passed at `obj/PresentationLayoutTests/0b07e9b474e845e1823fd43f650a3bf5`; final Debug checks include the subsequent disposed-strip font guard and initial event hookup. Both final configurations were rebuilt.
- Demo workbook SHA256 remained unchanged. This repair does not change workbook, calculation, Balance Sheet, or Funding editing code.

## Client check

In DIT and the analyser: close either panel with its own arrow, compact, reopen each by clicking its edge strip, then restore. Repeat with both panels open and both closed, and with Assumptions and Outputs side by side. Confirm physical 5k/DPI readability and mouse targeting. Hover alone should not reopen the summary; clicks should reopen it immediately and keep it pinned. These desktop acceptance checks remain manual.
