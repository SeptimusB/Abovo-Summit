# Funding schedule figures, colours and dialog scrolling — 2.91

Status: **Ready to test** with Jon. Debug and Release updated; no commit or push requested. Originals and masters unchanged.

## Changes

- Funding Schedule starts 130 logical pixels taller (1040 × 850 client area); the settings pane remains scrollable on smaller windows.
- DIT's application-wide wheel filter now checks message ownership, the enabled owner and the active form. Merely overlapping the DIT's screen rectangle is not sufficient. Schedule dialogs, Options and popup controls retain their own wheel processing.
- Editable scheduled amount cells use a slightly stronger but still pale Summit-only tint. A narrow colour marker also identifies scheduled targets while selected or unavailable; existing selection and unavailable-cell backgrounds remain authoritative. The actual workbook fill, protection and fill-pattern admission rules are unchanged.
- Optional **Populate target cells with a fixed figure**, off by default. Applies to the chosen facility only and every new date in each selected section. Existing/orphaned entries remain excluded by safe blank-row selection.
- Uses grid amount units. Percentage sections explicitly take percentage points: 5.125 becomes 0.05125 in Excel. Mixed amount/percentage selections are rejected when fixed figures are enabled; dates-only multi-section scheduling remains available.
- All dates are posted first inside the existing ChangeManager rollback boundary. For figures, a worksheet calculation updates date-dependent patterns, then each target must be empty, formula-free, unlocked and solid-filled. Existing numeric workbook validation is reused. No bypass of unavailable inputs is introduced.
- One Undo/Redo covers dates and figures together. As before, structural spare capacity is not undone. Fixed-figure choices/units are included in the schedule XML configuration summary; numeric values are ordinary worksheet inputs. Save/recovery use the established native persistence.
- `ProcessChanges` retains its original public signature and behaviour. It delegates to the new `ProcessValidatedChanges` entry point; the schedule supplies an admission callback that runs before each write and within rollback coverage.

## Verification

- Debug and Release builds pass; normal executable paths updated to 2.91. Quiet diagnostics regression passes; tracing is not re-enabled.
- Fixed-figure Release fixture: 17 assertions, including decimal precision, selected facility only, unchanged physical fills, Undo/Redo, XLSB save/reopen, unavailable-target rollback/metadata cleanup, mixed-unit rejection and percentage conversion. Evidence: `obj/ClientReportTests/bbe7bd4fdc804846bcd2915e988728af`.
- Native Debug schedule fixture: 78 assertions, including real dialog Apply with a figure, taller form, wheel ownership, opt-in default, multi-section dates, actual target-marker rendering and save/reopen. Evidence: `obj/ClientReportTests/4005648d018649e59d2a1bf792857193`.
- Final Release native fixture also verifies editable targets' actual rendered background against their saved group tint. Evidence: `obj/ClientReportTests/705ee1fcf7274ac19b1d86e9bebea050`.
- Existing blank-model fill-based editing regression: 17 assertions pass (typing, paste, defining-cell edits, Undo/Redo, relocking and readable unavailable cells). Evidence: `obj/ClientReportTests/de81fba57cdb4a5ba1a0fe0aff015372`.
- Macro-disabled Excel read-only open / separate SaveCopyAs / native reload preserves schedule XML, worksheet order, all 1,682 names and 6,231 Funding cell formulas/constants/dates/locks/fills in the disposable fixed-figure workbook. Evidence: `obj/FundingScheduleExcel/d6c17f6fec2a4b18ba940267207b2062`. This does not certify interactive VBA or financial results.
- Dialog and target-cell screenshots visually inspected. Test runners verify the source workbook hashes remain unchanged. The previously observed WebView class-unregistration warning 1412 can occur at runner teardown after successful assertions; no new schedule exception was observed.

## Stable functional tests — Jon first

57. **Dialog size and wheel:** open Funding Schedule and Options over a scrolled DIT; wheel over their controls and blank areas. Confirm the form behind does not move and the dialog's own scrolling/editors still work. Check restored windows and display scaling.
58. **Target colour:** add a coloured schedule; confirm date and selected-facility cells are identified, other loans are unchanged, unavailable inputs remain locked, and colour survives Save As/reopen.
59. **Fixed figure:** tick the new option; apply decimals to an eligible repayment facility. Check all new dates receive the figure, other loans/existing values stay unchanged, and one Undo/Redo covers both dates and figures. Try 5.125 in a percentage section (5.125%), an unavailable target (no schedule values retained), and mixed amount/percentage sections (clear rejection). Include a schedule requiring extra rows, then trusted Excel/VBA reopening.

Earlier numbers 1–56 remain unchanged. Jon functional checks, Alex's workstation/layout checks and trusted Excel/VBA/accountant acceptance remain outstanding.
