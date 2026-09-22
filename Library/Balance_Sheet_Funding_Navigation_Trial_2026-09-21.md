# Balance Sheet and Funding navigation trial — 2.51

21 September 2026. Test delivery, not financial sign-off. Source masters and supplied client workbooks were not saved or modified. No additional live/hidden worksheets, changes to Transactional_Records, synchroniser sizing rules, XML routes, or authoritative calculation formulas.

## Balance Sheet

The analyser's Statement of Financial Position now reads the calculated SOFP output in Transactional DB. The previous UseInBS-driven grid remains hidden for wrapper compatibility; it is not the displayed or exported Balance Sheet source.

- A strict v26 layout/formula adapter resolves the output by its unique title, validates relative labels, opening references and forty formula columns, and re-resolves after structural shifts. Unknown layouts show an explicit unavailable message; unknown line formulas preserve the headline and disable guessed detail.
- Figures show the opening balance and forty forecast years. Five section subtotals have their statement lines beneath them. Standalone financial totals remain independent workbook values, not additive transactions.
- Drill: statement section → line → opening / cumulative transaction contributions → workbook K:N hierarchy → description/record. Openings are not allocated to today's transactions. Signs and reserve criteria come from the four validated formula families. A transaction meeting both reserve criteria contributes twice. Unmirrored rows remain identifiable by their worksheet address.
- Grid and chart share one document. Changing the view or drilling does not calculate the workbook. The existing calculation/rebind/deferred structural lifecycle invalidates and rebuilds the document. Normal edits retain expansion, focus and widths. Disconnection clears stale values.
- The native TreeList supports cell multiselect, Copy and Copy with headings. Formatting comes from SOFP cells; explanatory cumulative detail inherits its statement line's formatting. Excel export includes all hierarchy nodes, forty-one numeric values, provenance and rule/diagnostic columns, with a warning not to sum parents and children together.

### Snapshot storage and compatibility

Only an explicit snapshot operation creates the additional regions. Existing dedicated snapshot worksheets are required; no load-time schema migration or new sheet is introduced.

| Region | Content |
| --- | --- |
| Existing transaction snapshot/comparison names | Unchanged transaction-table contract |
| Local `TDB_BS_Extra_Inputs` on both sheets | A:BV, immediately below Transactional_Records through the last SOFP output row; includes the reserve's extra source rows and opening/headline values |
| Local `TDB_BS_Bundle` on TDB Snapshot, BX1 downward | Version marker, SHA-256, chunk count and XML containing frozen hierarchy, values, styles, periods and mapping fingerprint; each XML cell ≤30,000 characters |
| Local `TDB_BS_Bundle` on TDB Comparison, BX1:BX3 | Version marker, mapping fingerprint and transaction geometry |

Comparison headline cells retain ordinary Excel formulas at the corresponding output coordinates. Grid/chart detail differences are the union of independently evaluated live and captured nodes—not today's classification applied to raw transaction differences.

Capture preflights client-owned content outside the recognised owned regions before clearing either sheet. A failure after mutation clears partial output; a preflight failure retains the previous contents. Protection/permissions and dirty state are preserved. XML is bounded, DTD/external resolution disabled, checksum and hierarchy checked. Excel's 15-significant-digit numeric persistence is tolerated only at serialization-rounding scale. Structural invalidation clears headers/markers cheaply and works without an analyser instance.

Old snapshots without a BS bundle remain usable for SOCI/Cashflow. BS explicitly requests a new snapshot; it never silently uses live openings. The supplied Stori v25 file has no dedicated snapshot sheets: its Live Balance Sheet works, but snapshot capture requires a compatible upgraded template.

## Funding / general DIT keyboard navigation

The reproduced upward jump was a delayed DevExpress page-focus scroll to the top of the entire tall VGrid, after the refresh itself had retained the cell. DIT pages now suppress this whole-grid focus scroll. Explicit navigation reveals the destination editor instead. Normal single controls retain their scrolling behaviour. The native scroll contract was checked against the installed 25.2 assemblies and [DevExpress XtraScrollableControl documentation](https://docs.devexpress.com/WindowsForms/DevExpress.XtraEditors.XtraScrollableControl._properties).

- Refresh preserves the focused cell/record, grid axes and outer page scroll, without taking focus from another window. Ordinary posted changes use the lightweight refresh path.
- Tab goes to the next editable cell across the row, then the next row/grid; Shift+Tab reverses. Right/Left stay within the current row. Up/Down seek an editable cell in the same column, then the nearest column at the bottom/top of the adjacent visible grid. At the outer boundary, stop rather than wrap to the top.
- Targets are re-evaluated after posting/calculation using current dataset protection/rules. Only visible/enabled grids and visible rows participate. Native popup-list navigation remains native while the popup is open; F4 / Alt+Down remain available. Modified shortcuts and selection keys keep their existing handling.
- Native editors and the painted column/first-column editors share the navigation route. Header edits commit on leave/navigation, not every value change. A failed validation stops navigation. Temporary VGrid header editors are parented outside the layout table and survive calculation/layout callbacks safely.
- MultiEditorRow/custom joint-venture paths retain native navigation. This is not a claim of complete coverage of every bespoke editor, nested detail view, hidden/collapsed category or client XML definition.

## Automated / native validation performed

All fixtures use private copies and verify unchanged source hashes. These targeted tools do not replace a general application integration suite.

- **Demo:** all 1,886 headline/opening values agree with Trad View to 0.001 workbook units; 20,213 populated branch-period checks reconcile. Adapter approximately 301–329 ms in these local runs; not a bitness benchmark.
- **Blank:** all 1,886 values agree; 15,662 branch-period checks reconcile. Zero-state evidence, not independent financial evidence.
- **Stori v25_0704:** 1,845 financial values agree. Its 41 Check values differ only because Transactional DB rounds and Trad View does not; that existing workbook behaviour is retained and explained. 17,835 branch-period checks reconcile.
- **Snapshot:** collision sentinel rejection before clear; exact entry protection permissions; zero initial differences; every frozen value survives native XLSB save/reopen; opening/classification changes stay frozen; missing/duplicate locator and unsupported-formula rejection; mapping changes rejected for old BS capture.
- **Structural:** real WorkbookManager five-row Cash Journals insert and delete, mirrored output relocation and recapture, snapshot invalidation without an analyser. Injected a qualifying reserve row outside Transactional_Records and under both criteria; confirmed double contribution, physical extra-input capture and frozen detail.
- **Native analyser:** all three statement charts, 830 group and 1,944 leaf series; forty-one BS points; sources/modes, right-click drill, wheel zoom, expansion/focus/width preservation, deferred stale-data clearing, 1100/1900/3000-width renders; BS grid render inspected; every exported BS node/value checked. Typed rent edit changes live BS; ChangeManager Undo/Redo restores the expected BS headlines.
- **Native Funding:** actual Funding Facilities/Loans at 450px page scroll, full refresh keeps the page/cell; native active-editor key events for Tab/Shift+Tab/arrows; real text commit and further editing; header buffering/commit; lower off-screen header reveal; rejected validation; all permitted Funding routes. CPI/RPI tests cross-grid boundaries; Rent/Voids tests existing banded-header combo and cell navigation.
- **Excel persistence:** native Excel read-only open → SaveCopyAs → Summit reload, macros/events/calculation disabled. Frozen values, hierarchy, sheet order and global names preserved. Excel rewrote the binary VBA container; all 337 module identities and source hashes remain identical across original Demo, Summit save and Excel save. No VBA source or password was extracted to the repository, and VBA was not executed.

Reproducible tools: `Tools/Test-BalanceSheet.ps1`, `Tools/Test-AnalyserChart.ps1`, `Tools/Test-EditorNavigation.ps1`, `Tools/Test-BalanceSheetExcelRoundtrip.ps1`, `Tools/Verify-VbaModuleHashes.py`. The latter requires disposable oletools outside the repository and emits hashes only. Ignored evidence is under `obj/BalanceSheetTests`, `obj/AnalyserChartTests`, `obj/EditorNavigationTests` and `obj/BalanceSheetExcelRoundtrip`.

## Required client acceptance / risks

1. **Funding first:** on a disposable populated file, scroll well down Facilities/Loans; edit text, numeric, date, combo and first-column editors; Tab/right/up/down and Shift+Tab. Confirm value, workbook address, undo entry, scroll and next-cell editability. Check F4/Alt+Down and popup arrows, rejected inputs, paste, mouse-click elsewhere and hiding/reopening the window. Test restored/maximised and the client's 5k/DPI configuration.
2. **Other DITs:** keyboard through normal and pivoted CPI/RPI, Rent/Voids, multiple separated grids, protected/rule-disabled cells, collapsed categories and reordered columns. Large client ranges may need further navigation performance tuning. Read-only regions must remain read-only. MultiEditorRow navigation is deliberately unchanged.
3. **Financial sign-off:** accountant compares opening plus every forecast year with Trad View and the Check Sheet. Examine asset/liability signs, cash, reserve overlap and cumulative versus annual meaning. Do not interpret workbook parity as proof that the underlying model is financially correct.
4. **Persistence:** create a fresh BS-inclusive snapshot, change an opening and a transaction heading, save/reopen Summit and verify Live/Snapshot/Differences. Run an approved Excel/VBA calculate/save on a copy and reopen Summit. Macros were deliberately disabled in the automated round-trip, so that path still requires acceptance.
5. **Structure:** further Development, Funding and component add/delete cases, with analyser absent/hidden/visible and snapshots present/absent; external Excel structural edits and reconciliation; save/reload. Cash Journals paths are covered, not this whole matrix. No new mirror sizing rules were added.
6. **Failure/version coverage:** old snapshots, workbooks lacking dedicated sheets, bespoke SOFP layouts, formula errors and conflicting snapshot contents must give useful diagnostics, not plausible stale figures. A valid known-profile document is required for BS drill. Unknown versions/bespoke formula families require reviewed mappings.
7. **Performance/architecture:** compare Release x86/x64 on representative extensively populated client files using the existing benchmark process. BS build/export and snapshot XML add memory/storage; current native fixture timings are not a full architecture comparison or proof against extreme-file memory pressure.

## Release gate

Debug/Release builds and final source/diff checks are recorded in the accompanying audit entry. Functional/client and accountant financial acceptance remain pending; no source workbook changes or automatic commit/push are part of this trial.
