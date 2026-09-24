# Summit 2.96 - readable grid zoom and restored-window sizing

Status: Ready to test. This follows the reported Funding date-editor clipping, analyser overpainting and Rent navigation crash in 2.95. No workbook, calculation engine, cell-fill editability rule or spreadsheet-engine integration is changed.

## Changes under test

- Repeating header editors take their effective font from the current grid, including window density and local zoom. Native text-area measurements include the editor border and dropdown button. Date rows and their header pane retain enough room for the text and fixed native indentation.
- Native DevExpress view-info has a separate PaintAppearance. The repeating-header painter uses an independent appearance snapshot: aliasing the linked editor's live appearance let native layout reset its scaled font to the default. Measurements use disposable repository clones; assigning and disposing a temporary measurement font on live view-info is also unsafe. The tests check actual text bounds, rendered glyph size and original-font validity, not just outer editor bounds. No whole-repository clone is added to the painting path.
- Maximise/restore re-applies absolute base layout and then restores the local zoom in one native update. It retains an unfinished header edit without posting it to the workbook. Minimum rendering dimensions do not inflate subsequent zoom/reset baselines.
- The agreed minimum zoom is 60%, with a separate 8pt readable text-size floor. Column widths and row heights respect that floor rather than continuing to shrink around unchanged text. Legacy lower preferences are clamped. Native chrome does not shrink below usable dimensions, so the most compact layout is deliberately not a uniform 60% photographic reduction.
- Shift+mouse-wheel zoom is anchored to the cell beneath the pointer, subject to native scrolling granularity and viewport edges. Funding uses native cumulative pixel offsets rather than estimating rows from the hovered row's height; this preserves the target across mixed category/date row heights. Normal wheel remains vertical and Ctrl+wheel horizontal. Context-menu zoom/reset remains available.
- Analyser data/group/total rows have approximately 10% more vertical space. Descriptions and totals are clipped to their own visible columns; horizontal scrolling cannot paint values beneath the fixed description column. A long description no longer suppresses a financial total. Values keep their original right-aligned column position, including partially visible edge columns.
- Initial analyser description fitting waits until the grid is visibly docked and has its real available width. Later resizing preserves manual column widths. Native indentation and group expanders are retained.
- Lazy banded tabs may construct repeating-column helpers before their GridControl is attached. Initial sizing now defers owner-dependent measurement until attachment; static scale/zoom subscriptions remain valid and are removed when the view is disposed. This addresses the reported Working-2.95 Rent navigation NullReferenceException.
- Repeated 200% Funding resize exposed a separate DevExpress 25.2 header-cache defect: when there are temporarily no visible rows, old header lines can retain disposed native brushes. Zoom/base-layout batches discard that obsolete drawing geometry before the font cache changes and again immediately before the final native record-header refresh. A temporarily empty viewport is restored using its saved row and native visibility API. This neither disposes vendor brushes nor rebuilds the datasource or commits an editor.

## Jon's functional checks

76. **Funding editor zoom and restore (retest):** try 60%, 100%, 200% and Reset; maximise and restore repeatedly with dates visible and an unfinished date edit open. Dates should fit, change size together and remain aligned with their rows. The unfinished edit must not be committed by resize/zoom alone.
77. **Analyser layout (retest):** on SOCI, Cashflow and Balance Sheet, zoom, scroll sideways, resize the description column and Reset. Check totals are readable, labels do not overpaint figures, and row heights are comfortable. Deliberately narrow columns may truncate text but must not change values.
82. **Pointer-centred zoom:** hover over a middle cell and use Shift+wheel in Funding, an ordinary grid and an analyser. The same cell should remain near the pointer, within native scrolling limits. Funding uses pixel offsets; ordinary table/tree layouts may still scroll by whole rows. Repeat near fixed panes and viewport edges; confirm the minimum is 60%, text remains readable, and ordinary/Ctrl scrolling is unchanged.
83. **Rent tab navigation crash:** open Rent assumptions, visit its tabs, navigate to another assumptions interface and return. Repeat with window resizing and the normal screen-sharing setup. Editors should appear without a constructor exception and still accept/cancel normal input.

## Validation

Both standard Debug and Release builds pass (`obj/build296-debug.log` and `obj/build296-release.log`), with test version 2.96. Final native evidence:

| Check | Result | Evidence under `obj/ClientReportTests/` |
| --- | --- | --- |
| Shared GridView/VGrid/Tree pointer zoom, readable minimum geometry and reset | 63 assertions in each build | Release `0a864a81b8cd48fd874efc913294dcd4`; Debug with application configuration `195d4f6b95bd4e808033e4554e260157` |
| Repeating-header batching, unattached views, linked-owner repositories, pending edits, inactive/active fonts and native text bounds | 352 assertions in each build | Release `39eaaacebdb54b57a01928c64673c02a`; Debug with application configuration `f6ca66a23d9849a29d67b5a596914e99` |
| Analyser initial fit, manual widths, numerical clipping and group/footer pointer anchor | 31 assertions in each build | Release `59180bd557ff481dba86d68e67ed01ab`; Debug `4c4dcba1e4fe49a1a5ad3dde7745a45e` |
| Analyser using the application configuration | 31 assertions | Final Release `25f1b224599240a682b74afdd7cf5f9d` |
| Actual Funding DIT at 60/100/200%, repeated 5000/1800-width and maximise/restore, pending date edit, native brush validity and pointer zoom | 223 assertions in both Release modes | Normal `4159ac9271724b9b8991809f614c0fb5`; application configuration `3ffb8d57c8d54831ba06974acbdaa831` |
| Rent lazy tabs and Rent/Stock navigation without workbook changes | 17 assertions in each build | Release `9e7a43a46e5f43298775bd99dbb56c1a`; final Debug with application configuration `fb6c568be22844a5a49d251cdbb39303` |
| Frozen Working-2.96 Rent navigation smoke test | 17 assertions with application configuration | `b5216e659f46457e94199e3492268922` |

The previously failing actual Funding pointer case is retained within one pixel in both Release modes. The native Funding screenshots were inspected after the complete resize sequence. The quiet-diagnostics gate passes in Debug and Release, and `git diff --check` passes. Normal diagnostic tracing stays disabled; the separately opted-in Check Sheet trial switch is unchanged. The final generic pointer fixture probes past a native row separator before asserting a data-cell anchor; post-zoom position tolerances are unchanged.

Native test DPI was 96. Presentation scales 1, 1.5 and 2.5, window-density changes and wide/restored layouts are exercised, but those are not proof of physical high-DPI or mixed-monitor behaviour. Jon's 5K maximise/restore check remains required. Source workbook hashes are checked around disposable-copy tests; no originals are saved. No client acceptance or financial/Excel-VBA round-trip acceptance is implied by these presentation tests. The separate Check Sheet public-heading refresh and reported Transactional DB remainder remain open.

A dependency-complete frozen copy of 2.95 was created at `bin/Working-2.95/Abovo-summit.exe`; its runtime test-version property and executable hash were verified at copy time. Subsequent user testing reported a Rent/lazy-tab constructor crash in that executable. It remains untouched for comparison but is no longer recommended as the working fallback. No blanket financial or UI certification is implied by its earlier targeted test passes.

Replacement frozen test copy: `bin/Working-2.96/Abovo-summit.exe`. All 299 copied application/dependency files matched Debug by SHA256 at copy time; browser profiles, old executable backups, publishing artifacts and logs were excluded. Its runtime test-version property is verified as 2.96 and its Rent navigation smoke test passes. Executable SHA256: `AD4800FD6F211B9FB635E1F1CE917FFB5EE201D63789F4C926ABED10D42B493C`. This is a test working copy, not a final client release or a guarantee against unrelated defects.

Native editor sizing reference: [DevExpress RepositoryItem.AutoHeight, version 25.2](https://docs.devexpress.com/WindowsForms/DevExpress.XtraEditors.Repository.RepositoryItem.AutoHeight?v=25.2). The fixed-height repeating editors use the actual native text rectangle rather than assuming the outer control bounds guarantee readable text.
