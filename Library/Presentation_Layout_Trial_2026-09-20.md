# Presentation layout trial — 20 September 2026

Current test release: **2.44**. Structure revision: **1751**. The original trial below was delivered as 2.38.

## 2.44 — swap the two images

At the user's request, Compact now displays the outward-arrow image and Restore the original three-panel image. Only the image assignment changes; actions, tooltips and hidden captions are unchanged.

## 2.43 — distinct Restore icon

Compact retains the approved three-panel SVG. Restore uses a matching window outline with outward arrows, without visible text. A shared state updater switches the vector and tooltip together, including newly opened interfaces while panels are hidden. Two cached vectors avoid allocating graphics for each toggle. The fixture verifies distinct states, original-icon restoration and unchanged icon display size.

Validation: Debug and Release builds both passed and updated their normal bin directories after the user closed the debug session. The group presentation fixture passed both icon states, new-interface state inheritance, size preservation, panel persistence and the existing responsive-layout checks. Source workbook SHA-256 unchanged. The 2.42 Debug deployment lock is resolved in this delivery.

## 2.42 — icon-only Compact / Restore

The shared DIT/Analyser panel toggle keeps its existing SVG icon without a visible caption. Its tooltip and underlying caption alternate between Compact and Restore; panel visibility/width persistence and the icon are unchanged. The native group fixture checks both states and the state inherited by a newly opened interface.

Release build and the group presentation fixture passed, including icon-only state assertions and unchanged source workbook SHA-256. Debug compilation passed but deployment to bin/Debug was blocked by the user's running Visual Studio/Summit debug session; close that session and rebuild before testing Debug 2.42.

## 2.41 — restored group windows, sidebar readability and percentage refresh

The shared GIT/DIT/Analyser automatic density now follows the owning group window's logical client width, capped by the existing monitor scale. Explicit user font preferences remain intact; the child document width does not feed back into that density. Navigator and summary panels reserve the document area (25% and 22% upper width limits). Hide/restore retains exact manual widths in an unchanged window, but reapplies limits if the window changed size while hidden. The title bar height follows its font.

DIT's native WindowsUI toolbar is outside the designer-scaled absolute TablePanel row. The active section uses its tab viewport width rather than the entire group form width. Resize relayout does not invoke the legacy ResizeFonts editor-close path, recreate bindings, or calculate the workbook. Host resize notifications are coalesced.

Sidebar HTML text increases to 9.5pt (from 8/8.5), multiplied once by the explicit user preference. Summary tables use available width and omit entirely empty spacer rows, preserving labelled and zero-valued checks. File Details uses a two-column float layout supported by the embedded browser, replacing unsupported CSS grid. Browser zoom and CSS font scaling do not both own the preference.

The reported AbovoDESpinEdit.RefreshData NullReferenceException originated in the editable single-percentage creation path: metadata was assigned to Tag but TargetWorksheet/TargetCell were not initialised. Both are now supplied. Refresh ignores disposed/unbound editors; released DITs ignore late refresh and single-cell posting callbacks. Normal edit writes still go through ModelChangeManager. No workbook file, formulas or calculation policy are changed.

Validation includes Debug/Release builds, isolated native restored-window renders, the Repairs Stock Survey Allocation percentage editor, unbound/disposed refresh and source SHA-256 preservation. Manual acceptance: restore/maximise at 5k and on a smaller monitor; change font preference; hide/show each panel; edit/undo a Stock Survey Allocation percentage and close/reopen the interface/model. Confirm no editor is committed merely by resizing. ActiveX/MDI DrawToBitmap is not a trustworthy composite screenshot: inspect the separately rendered DIT and browser DOM metrics. Physical monitor/DPI acceptance is still required.

Final group fixture passed: 1280/1900/2800/5000/restored widths, actual native maximise/restore, panel hiding with intervening resize, the bound percentage editor, refresh after interface resource release, existing Funding button sizing and file-instance 100/150/200/100% checks. At 1900px the navigator/sidebar measured 315/340px, leaving a 1195px document and a 74px toolbar; at 1280px the toolbar intentionally wraps to 140px. DOM body font was 9.5pt. Debug and Release builds passed; Demo source hash unchanged. Final group renders: `obj/PresentationLayoutTests/0fbf79c3f2a54261b14f8ef26d28d28b`. The full earlier run also passed the unchanged Stress Test layout regression.

The preceding work was checkpointed as `d531635` before this tranche. This trial is presentation-only: it does not change workbook formulae, named ranges, calculation policy, model schema or edit/history routing. No source XLSB was saved during validation.

## Changes and boundaries

- Stress Test's live multivariable workspace uses the available page width instead of a 2,200-pixel cap. Input columns share that width; existing display/user scaling controls fonts, rows, headings and command controls. The unused covenant-detail column collapses while multivariable mode is off and returns when it is on.
- A shared native-control layout helper measures SimpleButton captions after their final font is assigned. DIT command hosts also account for the measured height. WindowsUI captions, glyphs and their circle backgrounds scale together, with space reserved for wrapped toolbars. Bitmap resizing always starts from the resource image, not a previously resized image; generated images are released with their owner.
- File-instance controls receive their initial scaling pass, not only a resize pass. Main-screen and file-details HTML use point-based fonts from the same existing display scale and explicitly opt out of the second browser zoom multiplier. Saved user font settings are not rewritten.
- Optional `NavigatorCaption` and `NavigatorGroupCaption` XML elements supply short display labels. All original `CSName`, `GroupName`, GSID/CSID values, worksheet references and range definitions remain intact. Older embedded XML without aliases still falls back to its existing names.
- Navigators use single-line captions and measure their contents. Width is capped at 45% of the host window to retain a usable document area; exceptionally long captions on narrow windows ellipsize rather than wrap. Combined's top-level black group styling is retained.
- DIT and Analyser toolbars gain **Hide panels / Restore panels**. State belongs to their containing GroupInterfaceTemplate, not the model globally. The navigator and summary panel retain their visibility and width; other windows are unaffected. Attaching another interface to an already-hidden host gives it the Restore caption. Hiding is not disposal or a data rebuild.

## Validation

- Debug and Release builds passed.
- XML parsed successfully. Removing the 67 new caption elements and restoring XMLVer produced the exact original XML document: no route/range/definition changes.
- `Tools/Test-PresentationLayout.ps1` opens the authoritative Demo master without Save/SaveAs and checks its SHA-256 afterwards. Native fixtures verified Funding's full command caption, no-wrap navigation, separate host state, exact visibility/width restoration, DIT and Analyser toggle registration, initial file-details formatting, and Stress Test widths of 1,570 / 2,770 / 4,970 pixels for the corresponding 1,600 / 2,800 / 5,000-pixel form sizes.
- Repeated user scale changes of 100 → 150 → 200 → 100% (not persisted) verified non-cumulative fonts and glyph sizes, and that each glyph fits its circle.
- `Tools/Test-FileInstanceLayout.ps1` passed nine width/font variants (750 / 1,100 / 1,600 pixels; 9 / 14 / 18 points), alignment, non-overlap, native button hit targets, order and DSA caption.
- `Tools/Test-StandaloneHistory.ps1` passed the existing Demo FFR and Stress Test history regression: edits, undo/redo, native FFR editor refresh, hidden/reopen state, deduplicated interface history and final model-close disposal. Source SHA-256 was unchanged.
- Generated native-control renderings were inspected for Stress Test, Funding and file-instance layout. Rendering exposed and corrected two additional issues: retained narrow Stress Test columns, and glyphs outgrowing WindowsUI's default fixed circle. The supported `ButtonBackgroundImages` collection now scales the circle too.

These are isolated runtime/layout fixtures, not a complete UI or Excel/VBA acceptance suite. Test renderings are disposable under `obj/PresentationLayoutTests` and `obj/FileInstanceLayoutTests`.

## Manual acceptance / risks

1. On the 5k monitor, inspect first launch, maximize, restore, side-by-side windows and moving between monitors. Repeat on the client's normal monitor. Physical Windows DPI transitions remain unverified by the size fixtures.
2. Check Stress Test's Stresses and Mitigations tabs, capture controls and both multivariable modes. At narrower widths, horizontal scrolling is preferable to unreadably compressed input cells. Verify a normal edit, undo/redo and retained form reopening.
3. Check Funding's Add Funding Columns and other DIT command rows at the user's preferred font size. Native auto-size can expose previously hidden layout crowding in uncommon multi-command sections.
4. In both separate and Combined windows, follow short navigator labels to the expected interface. Check nested black/blue group colours, long captions and preserved expansions. Confirm older embedded XML still navigates using its original labels.
5. Set left and right panels to different states, hide all, then restore. Repeat with one panel already hidden, both visible, an autohide panel, different widths and two model windows. Grid expansion/selection, input edits, history and snapshots must remain unchanged. Open another interface in the same hidden host and restore from its toolbar.
6. Inspect main-screen, file-details, button hover/pressed/disabled states, initial paint and font changes. Confirm the existing user preference persists after restart and no captions clip. Test narrow toolbars with the added panel toggle.
7. Regression smoke test: save a disposable model copy in Summit, reopen in Excel/VBA and return to Summit. No workbook serialization changes were introduced, but this remains part of release acceptance.

Native API reference: [WindowsUIButtonPanel.ButtonBackgroundImages](https://docs.devexpress.com/WindowsForms/DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel.ButtonBackgroundImages?v=25.2).

## Reminder

With the user's explicit approval, the existing monthly bitness-check heartbeat was temporarily replaced by a 20-minute inactivity check. The reminder has since fired and the original monthly benchmark prompt and schedule have been restored. This is app automation, not a Summit code change.

## 2.39 follow-up — Stress Test refresh typography and toolbar

The user's 5k screenshot exposed refresh paths that the original first-paint checks missed. `RefreshCovenantSummary` replaced the scaled labels with default-font labels; `ProcessBreachesGrid` called the shared formatter after sizing, resetting its fonts and disabling column auto-width; regenerated chart titles used a fixed 10-point font. These paths now reapply the same display font as the stress input grid after recreation/binding. Summary labels retain workbook colours, displayed values and bold treatment. Old summary blocks are disposed when replaced. The results grid fills its allotted pane and retains its movement-symbol font.

The WindowsUI toolbar is now left-aligned, retains Home-first/History-last visual order, has explicit captions, and uses a vector History glyph instead of the missing bitmap resource. The top TablePanel row auto-sizes to the measured control height; the content row takes the remaining space. An absolute row had requested 138 pixels but rendered the toolbar only 76 pixels high after the form's designer scaling. A minimum height is used without a maximum-width constraint, avoiding the native zero-width edge case discovered during font preference testing.

The final presentation-scale hook reapplies absolute fonts after the generic scaler, and summary fonts are normalised after ControlAdded notifications. This prevents the same user preference from being multiplied twice during refresh. No calculation, workbook, XML routing, editor or history behaviour is changed in this follow-up.

The fixture now tests summary recreation and results rebinding, not only resizing. It checks the toolbar's actual height, non-overlap, visual order and native hit targets; test renders live under `obj/PresentationLayoutTests`. Manual 5k/mixed-DPI acceptance remains necessary. In particular, inspect the summary/results after an edit, undo, covenant selection, chart regeneration and switching Stress Test pages; inspect all toolbar captions and clicks at normal and larger user font settings.

Follow-up validation: Debug and Release builds passed. The native Stress Test fixture passed at form widths 1,600 / 2,800 / 5,000 and after repeated 100 / 150 / 200 / 100% preference changes. Recreated summary and rebound results matched the input grid's font (19.56 / 29.33 / 39.11 / 19.56 points on this host); toolbar height followed 130 / 186 / 243 / 130 pixels, with all eight buttons reachable in the intended order and no content overlap. The file-instance layout fixture passed all nine existing width/font variants. The Demo standalone history fixture passed edits, undo/redo, retained hidden/reopen state and model-close disposal. Source workbook hashes were unchanged. Tests are isolated fixtures; the user's running application and physical monitor/DPI transitions were not driven.

Native layout reference: [DevExpress TablePanel auto-size rows](https://docs.devexpress.com/WindowsForms/DevExpress.Utils.Layout.TablePanelEntity.Style?v=25.2).

## 2.40 follow-up — restored Stress Test windows

The restored-window screenshot showed that the monitor-based density was still being applied to a much smaller client area. Stress Test now caps the automatic portion of its scale by the current logical client width (1920-pixel design reference, never below the normal baseline), while retaining the user's explicit font preference and the approved full-width scale. This is scoped to Stress Test; other interfaces retain their existing scale policy. The shared WindowsUI helper accepts an optional scale for this caller.

The toolbar is hosted above the content, outside the designer-scaled TablePanel. Content bounds use its final measured height, preventing a stale docking pass from leaving a gap or overlapping the page after a resize. Capture hosts reserve the two editor rows' actual height and use percentage-height inner rows. Narrow capture labels shorten to Name/Test without changing their workbook bindings. Summary blocks stack vertically when their measured labels/headings cannot fit side by side; they return to two columns on wider windows. Chart markers adapt to the available width and are reapplied after regeneration. Input columns retain readable widths and horizontal scrolling when they cannot all fit; wide grids still stretch. Results-header minimum widths are measured from their captions.

The fixture covers 1280x900, 1792x1200 (close to the user's screenshot), 2800x1500 and 5000x1800, repeated narrowing/widening, actual maximise/restore, summary recreation, results rebinding, capture-control containment, toolbar hit targets and explicit font preference changes. Generated renders are inspected, not just the outer control width. Mixed-monitor DPI changes, custom client font settings and manual editing while resizing remain acceptance checks. No workbook values, calculation policy, structure XML, edit routing or history semantics are changed.

Validation passed on 20 September: Debug and Release builds; the full presentation fixture (including actual maximise/restore and 100/150/200/100% preference changes); all nine file-instance layout variants; and the Demo standalone FFR/Stress Test history regression. Restored text was 10 points at 1280/1792 client widths and 14.58 at 2800, returning to the approved 19.56 points at 5000 on this host with its existing preference. The toolbar reduced from 130 to 74 pixels on restore. All three capture controls were wholly inside their ancestor client rectangles. Original workbook SHA-256 values remained unchanged. The 1280-wide render intentionally uses horizontal scrolling and vertically stacked summary blocks; no font is reduced below the existing normal baseline to fit more columns.
