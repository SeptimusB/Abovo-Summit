# Pattern-based current editability - 2.87

Status: **Ready to test** with Jon. Not committed or client accepted.

The user clarified that conditional formatting here means Summit assessing a cell's current editability from its XLSB fill pattern, not implementing a new Excel conditional-format renderer.

## Cause and narrow repair

The 2.84 readability change replaced the shared normal/banded-grid locked-cell lavender cue with the cell's background/font colours. This made currently unavailable inputs appear like ordinary blue/grey input cells. The actual `HasRules` / `Fill.PatternType <> Solid` admission checks and refreshed lock state remained effective in the reproduced Stock case.

The shared `GridView_CustomDrawCell` path again renders rule-locked cells lavender, with readable Abovo blue text matching the existing vertical-grid cue (not the old white-on-lavender text). It applies to all DIT normal/banded grids using that handler, not just Stock. Selected/hover treatment, workbook fill rules, edit/paste admission, calculations and source workbooks are unchanged. This is not a claim to have solved the separate outstanding full conditional-format-parity or missing-XML-rule cases.

## Validation

- Debug and Release builds pass; both normal executables updated to 2.87.
- `Tools/PatternEditabilityFixture.cs`: **17 assertions pass in each build**, on separate disposable Blank copies. Native editing and paste eligibility reject initially locked cells, allow them after the description edit, relock after Undo/clear and unlock after Redo. Edits use ChangeManager.
- The regression fails on the pre-repair 2.86 Release at the unavailable-input background assertion, after all 14 edit/lock-transition assertions pass. It passes after the repair.
- Native drawing-event checks verify lavender background and nonwhite text, and the workbook DarkGray fill remains unchanged. The Release captured Stock grid was visually inspected.
- Source workbook SHA-256 unchanged. Debug quiet-diagnostics regression passes. No production tracing enabled.
- Existing embedded-browser teardown logs error 1412 after completed assertions; this is not a new browser fix or acceptance claim.

Evidence: `obj/ClientReportTests/1c554a7f432b42b1852b165977a5e235` (pre-repair), `12e056b3dc884cb0a5288bfb683ffcd8` (Release), `9ee3fc06c7f042a0a81418d469f7200d` (Debug). Private test files/images/logs remain ignored.

## Stable functional checklist addition

**50. Currently unavailable inputs:** on Stock and other affected DITs, confirm the unavailable-cell cue is restored with readable text. Enter/clear the controlling description and use Undo/Redo: appearance and edit/paste availability should change together. Existing checklist numbers 1-49 are unchanged. This item is for Jon's functional test before client handoff.
