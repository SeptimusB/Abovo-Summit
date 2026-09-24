# Summit 2.95 - grid zoom, fixed boundaries and Check Sheet links

Status: **Ready to test** with Jon. No Alex sign-off, commit or push is implied. This is the production DevExpress route; the separate spreadsheet-engine research is unchanged.

## Behaviour and causes

- Repeating in-place header/date editors now participate in one batched zoom/layout pass. Equivalent native editor heights are measured once per pass. Normal, focused, disabled and read-only editor fonts scale together. Zoom no longer commits an unrelated pending header edit or closes its editor one helper at a time. The worksheet's masks, fills and editability rules are unchanged.
- Grid zoom uses stable size baselines rather than accumulating rounded deltas. Rapid Shift+wheel requests are coalesced. The analyser's custom painters read actual fonts, so point-size scaling replaces the previous font-delta-only path. Funding and banded-grid row heights now follow zoom; focused header editors no longer grow on repeated reset cycles.
- Analyser menus retain Zoom after rebuilding their other actions: in, out, reset and Set zoom presets. Reset also restores automatic analyser row heights after an accidental manual shrink. Description columns are capped at 40% of the viewport instead of retaining their largest previous best-fit width. Group rows and totals are more compact; native total-cell bounds are fitted inside the shorter rows to avoid clipping figures. Balance Sheet tree zoom/reset is covered too.
- The unwanted single-pixel header rule beside Stock's dummy spacer is removed without removing the spacer. Fixed-column/fixed-row boundaries use a distinct three-device-pixel SteelBlue divider, including the pinned Funding loan-name header and the fixed date/title pane. The latter overlays the existing header painter without replacing it; open date editors remain inside the separator. Category titles span scrolling records and keep their native rendering. These are control decorations, not workbook cell fills, and stay three pixels through zoom.
- Funding's missing **Filter to facility [name]** was caused by reading a sibling menu item's effective Visible state while its popup was closed. Both loan filters now receive their availability independently. Facility and funder identity mapping remains distinct; filter/rebuild uses physical source-record mapping.
- Undo and Redo now calculate restored source worksheets even when their old interface is no longer registered with the calculation engine, then run the established active-sheet calculation/refresh. The registered-sheet path already recalculated. No unconditional full-workbook calculation was added; the existing multi-active-object policy is retained. Rollback uses the same source-sheet handling.
- Check Sheet Summit destinations use their explicit worksheet target even when an adjacent Excel link caption is blank. One ordinary click opens a unique destination; ambiguous destinations still offer a choice. Native mouse hit testing avoids the GridView RowCellClick suppression seen when editors are enabled. Drag selection, modifiers, copy and the Yes/No editors retain their own behaviours. Untouched AGL already displayed all 27 destinations; this does not prove the client's exact disappearance had only this cause.

## Jon's functional checks - continue existing numbering

76. **Funding editor zoom:** try repeated 50%, 200% and Reset, plus Shift+wheel with the date editors visible. Text and row heights should stay aligned, refresh together and return to the original size. Check an unfinished header edit survives zoom and commits once normally.
77. **Analyser zoom and totals:** try SOCI, Cashflow and Balance Sheet; zoom both ways, use right-click Set zoom/Reset, resize the first column, and accidentally shrink a row then Reset. No expanding description-column ratchet, stuck tiny rows or clipped totals. Check financial figures remain unchanged.
78. **Fixed boundaries and Stock spacer:** Funding's pinned loan headings and fixed date/title region should be clearly separated from scrolling records. Fixed columns in other grids should have a three-pixel blue divider. Stock's intended blank spacer remains without the old thin header artifact.
79. **Facility filter:** right-click a loan and select Filter to facility [its actual name]. Only that facility's loan columns should remain. Repeat with a funder, scroll/zoom, then Remove filters or click the orange filter warning. Confirm edits and copied headings still refer to the correct physical loan.
80. **Undo calculation:** change a disposable input with a visible calculated result, Undo/Redo, and repeat after switching interface before opening history. Source-sheet results should follow the restored input without an extra manual calculation. Save a disposable copy and check it in Excel/VBA before financial acceptance.
81. **Check Sheet links:** a single click on a Summit destination opens it; multi-destination entries offer a choice. Check return navigation, multi-cell copy and Yes/No overrides. The separate public-heading warm-switch warning and reported Transactional DB remainder are still open, not declared repaired here.

## Validation

Both standard Debug and Release executables are built as test version 2.95. Logs: `obj/build295-debug.log` and `obj/build295-release.log`. All native workbook tests use disposable copies; no original is saved.

| Fixture | Result | Evidence under obj/ClientReportTests |
| --- | --- | --- |
| AnalyserZoom295Fixture | 43 assertions in each configuration, including native footer-cell bounds and repeated zoom/reset | Debug `252bb7d36b7c419aaf29a6ce55dac9ab`; Release `e4c67bfd30b64048b15ec7a18afc8459` |
| InplaceZoomFixture | 118 Release assertions; one layout for both 4 and 100 helper controls in horizontal and vertical grids; pending edit preserved | Dedicated executable `obj/ProductionZoomFixture/InplaceZoomFixture.exe` |
| Client294Fixture | 58 Release assertions: real selection, copy/paste, scroll, filters, focus, editors and shared zoom | `3e82fe77a19340188df0b449960f4810` |
| UndoWorksheetCalculationFixture | 17 in each configuration; registered/unregistered source sheets, Undo/Redo and calculation policy | Debug `e8b024f356624348afa08d8b35be225e`; Release `a9af2f98963a4f0db535c6ea625337fd` |
| PatternEditabilityFixture | 17 in each configuration; fill locks, editor/paste admission and relocking | Debug `0bb1ad667a79409c9e6359c92583d92c`; Release `083b5b6a5d774be9bd9db267aa9e0c8d` |
| CheckSheetLinksFixture | 36 in each configuration; actual mouse down/up, direct and ambiguous destinations, selection and return context | Debug `4cb07ff7afdd464e87fb682858c135f9`; Release `827f3d6524fb4f279b2a87373d8abef1` |
| FundingFacilityFilterFixture | 11 in each configuration; real DIT, nonadjacent physical loan columns, warning, filter replacement and removal | Debug `36a2226e472c4b92a21d1f57426a5623`; Release `49774e921c3a4f16be1fd2a6ac9fb870` |
| FundingFilterMenuFixture | 68 in each configuration; closed-popup availability, both three-pixel fixed boundaries through 50/100/200/reset and horizontal scroll, unchanged header-editor pixels and active-editor bounds | Debug `eb945c58217243099664e707c43897ad`; Release `c7b75de3cd7145f18cda1781d89ffa37` |

Generated Stock, Funding and analyser screenshots were inspected. Initial compact-total screenshots exposed clipping; the final tested row/cell geometry corrects it. Some successful WebView teardown logs contain the known Chrome_WidgetWin_0 unregister 1412 message; those fixture processes exited zero. The small local layout measurements are diagnostic evidence, not client-machine performance guarantees.

Original SHA-256 values remain unchanged:

- Blank: `E05274ACD3D4821AC38F013AC45A7B57CE335BD524938D9F21E1B640D0CC5E90`
- Demo: `1ED79726D8D1129C699242C9E8B3E988AC920EA7108D3C0B105BF9B0BC8A5E9C`
- AGL v26_0005: `30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47`

Physical client/mixed-DPI acceptance, full application reopen and Excel/VBA round-trip acceptance remain manual. No master, business formula, range structure or underlying fill-based protection rule is modified by this delivery. No new production benchmark tracing is enabled.

Native API references checked against installed DevExpress 25.2.4: [BeginUpdate](https://docs.devexpress.com/WindowsForms/DevExpress.XtraGrid.GridControl.BeginUpdate), [RowCellClick editor-mode behaviour](https://docs.devexpress.com/WindowsForms/DevExpress.XtraGrid.Views.Grid.GridView.RowCellClick?v=25.2), and [VGrid fixed-row dividers](https://docs.devexpress.com/WindowsForms/DevExpress.XtraVerticalGrid.BaseOptionsView.FixedLineWidth?v=25.2).
