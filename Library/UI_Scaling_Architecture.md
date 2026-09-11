# Summit UI scaling architecture

## Runtime contract

- DevExpress application settings are loaded from `App.config` during the VB
  application `Startup` event, before the first form is constructed.
- The process uses Per-Monitor V2 DPI awareness. Windows and DevExpress own the
  physical scaling that occurs when a window moves between monitors.
- The default UI font is established before the first form is constructed.
  A user-requested interface-scale change may replace that default and broadcasts
  a presentation-scale event so already-open forms and later child controls stay
  consistent.
- Window resizing may alter layout and data density, but must not alter font
  size. Screen pixel width is not a font-size input.
- Workbook-backed cell fonts, colours, number formats and alignment remain
  workbook-owned. The application default applies only to interface chrome.

## First implementation stage (versions 1.51-1.52)

- `FontManager.InitialiseApplication` is the single startup entry point for
  DevExpress DPI and font settings.
- The former primary-screen resolution bands have been removed.
- `AbovoAppCls.GetFont` retains its existing signature for compatibility, but
  its legacy window-width scale argument no longer changes text size.
- FormMain and GroupInterfaceTemplate choose their initial bounds from the
  working area of the monitor containing the pointer rather than the primary
  monitor.
- FormMain, GroupInterfaceTemplate and the Stress Test no longer resize their
  fonts when their window size changes.
- Data Interface Template controls and grids use the central normal UI font rather
  than forcing the application small-font role.
- Live grid custom drawing retains the DPI-aware base font and applies only the
  workbook-owned bold, italic and underline style as an appearance delta.
- Version 1.52 adds a bounded large-workspace scale after removing the active
  monitor's native DPI factor. This improves legibility on very large displays
  running at low Windows scaling without double-scaling high-DPI displays.
- FormMain, GroupInterfaceTemplate and DIT controls use that stable per-display
  scale when they are constructed; font size does not track ordinary resizing.
- Fixed live-grid row, header and padding dimensions follow the same scale.
- Version 1.53 separates monitor-transition layout from the legacy DIT font
  routine. After a destination DPI change, the active section is fitted once
  from its current content and destination viewport using absolute, non-growing
  measurements; dimensions from an intermediate monitor are not retained.

## User interface scale (version 1.54)

- PresentationScaleManager owns a user-scoped 75%-200% interface multiplier,
  stored through ApplicationSettingsBase; 100% is the default.
- The effective scale is Windows/DevExpress native DPI scaling, followed by
  Summit's bounded logical-workspace scale, followed by the user multiplier.
  Summit must not manually apply the Windows DPI factor a second time.
- The options form is available from FormMain and DIT action bars. Apply updates
  open forms; newly opened forms and lazily created child controls are detected
  and receive the current scale.
- The shared pass covers ordinary controls, DevExpress appearances, XtraGrid and
  band metrics, VerticalGrid appearances, tab headers, Windows UI button panels,
  chart titles, accordion elements and embedded HTML zoom.
- DIT column and VerticalGrid in-place editors listen for scale changes, rebuild
  their cached natural height and scale fixed padding and header indent values.
- GroupInterfaceTemplate explicitly reapplies font roles to every nested
  accordion element, including popup menu items. Compact sidebar HTML is
  generated using the user scale.
- Workbook-rendered cell styling remains workbook-owned. Interface scale changes
  presentation size only and must not write formatting into the XLSB.

## Remaining staged work

1. Validate 100%, 125%, 150%, 175% and 200% display scaling, including moving
   live windows between mixed-DPI monitors.
2. Migrate specialist forms away from proportional `Resizer` behavior and
   primary-screen assumptions as each interface is tested.
3. Replace fixed layout coordinates with DevExpress `TablePanel`, `StackPanel`
   or `LayoutControl` where practical.
4. Scale custom-draw padding and line widths through `GraphicsCache.ScaleDPI`
   and retain `GraphicsCache` drawing APIs.
5. Prefer SVG icons or DPI-aware image collections over fixed raster assets.
6. Consider a separate Compact UI/data-density preference only after the
   interface-scale behaviour has been validated across all production displays.
