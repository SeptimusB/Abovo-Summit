# FFR native interfaces technical audit

Audited against `Library/TestFileMigrated.xlsb` on 19 August 2026. Excel automation was read-only with macros, events, prompts, and link updates disabled; the workbook was closed without saving.

## Interface contract

The FFR window presents these workbook sheets as native DevExpress grids without hosting or exposing a `SpreadsheetControl`:

1. `FFR Validation Summary` (`A1:J31`)
2. `Front Sheet` (`A1:J63`)
3. `FFR Inputs Adj Stmt` (`A1:AI74`)
4. `FFR Workings` (`A1:AK546`)
5. `Statements` (`A1:AG153`)
6. `Assumptions & tenure inputs` (`A1:AG206`)
7. `Compliance Questions` (`A1:J107`)
8. `FFR Key Defn` (`A1:N228`)

Each tab is created lazily. Seven worksheet-scale tabs use an in-memory snapshot of the live workbook, so a grid edit cannot write the XLSB before change management. Cell text, resolved fill and font colours, emphasis, alignment, row heights, column widths, and simple cell-note tooltips are re-read from the workbook. Formulas and locked cells are read-only.

`Front Sheet` is purpose-built rather than presented as a literal `A1:J63` grid. It displays the workbook title, RP number (`B5`), first forecast-year end (`B6`), registered-subsidiary confirmation (`B7`), non-registered-entity confirmation (`B36`), and the two 25-row workbook lists (`B10:C34` and `B39:C63`) as native controls. The entity lists are shown side by side so both blocks remain visible without magnifying Excel row/column geometry. Their values, colours, locks, list validation and `C9` guidance note remain workbook-owned.

## Edit contract

- Editability requires both `Protection.Locked = False` and `Fill.PatternType = Solid`, matching the established `DataInterfaceTemplate` and Stress Test rule signal.
- Workbook list validation creates a native `ComboBoxEdit`; date cells use a native date editor; other unlocked cells use typed text/numeric submission based on the workbook value and number format.
- An accepted edit creates one `DataChangeEvent` and calls `ModelChangeManager.ProcessChange`. The manager owns typed writing, change logging, dirty state, calculation, and rollback on failure.
- The active sheet snapshot reloads from the calculated workbook after every edit. Inactive tabs reload when selected.
- `Statements` and `Assumptions & tenure inputs` contain no workbook-unlocked cells in the audited model and are therefore calculated/read-only surfaces.

The development workbook is not protected at worksheet level for these sheets. That is not treated as an editability signal; cell protection and the workbook fill-pattern rule remain authoritative.

## Workbook findings

| Sheet | Formula cells | Constant cells | Workbook-unlocked cells | List-validation surface |
|---|---:|---:|---:|---|
| FFR Validation Summary | 15 | 22 | 0 | None |
| Front Sheet | 1 | 13 | 105 | `B7`, `B36`, registered and unregistered entity input blocks |
| FFR Inputs Adj Stmt | 476 | 88 | 155 | `C20` |
| FFR Workings | 12,983 | 1,097 | workbook-driven input blocks | Multiple annual/input selectors including rows 376-385 and named input rows |
| Statements | 4,246 | 374 | 0 | None |
| Assumptions & tenure inputs | 4,896 | 441 | 0 | None |
| Compliance Questions | 98 | 194 | 148 | Question and annual response cells including rows 21-73 |
| FFR Key Defn | 178 | 437 | 33 | None in the audited workbook |

The workbook contains 80 simple notes on `FFR Workings`, 27 on `Compliance Questions`, three on `FFR Key Defn`, and smaller note sets on the front/input sheets. The native grid exposes notes as cell tooltips; their text remains workbook-owned.

## FFR return generation

The native `Create FFR return` action replaces the workbook `FFR_New_Extraction` button entry point while retaining the workbook mapping contract:

- validate the selected provider template using `Cover Sheet!B4`;
- enumerate `FFRRangeNames`/`FFRListHeading` mappings from the active model;
- copy values into the mapped destination ranges of a separate macro-enabled workbook;
- save the result as XLSM without modifying the active XLSB.

The export workbook is an external target, not an interactive edit to the active model, so it is intentionally not submitted through `ModelChangeManager`.

## Required manual validation

- Open a populated client XLSB and compare all eight native tabs with Excel, including calculated values, conditional colours, long row labels, and note tooltips.
- Confirm locked/non-solid cells reject editing and that accepted list, text, date, percentage, and numeric edits update once, calculate, mark dirty, survive save/reopen, and appear correctly when the workbook is reopened in Excel/VBA.
- Confirm formula-driven changes propagate between FFR sheets when switching tabs.
- Generate a return using the current provider-specific template, reopen the XLSM in Excel, and verify mapped values, formulas, names, macros, and regulator-template behaviour.
- Repeat the export with cancellation and an invalid template to confirm no partial file or model change remains.

## Runtime validation update

On 19 August 2026 the Debug application opened `TestFileMigrated.xlsb` through the normal `FormMainScreen.OpenModelProceedureBP` path. The purpose-built `Front Sheet` rendered its return details and both 25-row entity lists. `FFR Inputs Adj Stmt` and `FFR Workings` were then selected in the same FFR window and both loaded without an exception or hidden error dialog. No workbook input was changed during this visual check.

Debug builds may auto-open `Z:\Sandbox\TestFileMigrated.xlsb` when that file exists. The hook is compile-time excluded from Release and still uses the normal model-open path.

### 2026-08-20 validation-summary refinement

Version 6.72 replaces the literal A1:J31 presentation of FFR Validation Summary with a compact, purpose-built read-only view. The visible workbook contract is rows 5-28: A5:A16 supplies the production instructions and warning text, A18:C23 supplies the hard-error heading, count caption, four validations and messages, and A25:C28 supplies the equivalent two-row soft-notification section.

The two summary sections use centred native DevExpress grids following the established Stress Test presentation policy: no filtering, sorting, grouping, column menus, row indicators or edit path. The workspace is capped at 1,240 pixels and centred when wider space is available. Empty spreadsheet rows and unused columns are not presented.

All captions, calculated counts, messages, resolved colours, font emphasis and alignment are re-read from the live workbook when the tab is opened or refreshed. Runtime verification through the normal Debug startup path displayed hard counts 61, 84, 0, and 27, soft counts 33 and 24, and the exact workbook messages. The verification workbook was closed without saving.

## Build validation

Version 6.72 passed Debug and Release MSBuild validation on 20 August 2026. The repository has no automated workbook/UI integration suite. The workbook-open, compact validation summary, and Tabs 2-4 checks above ran in Summit's normal application startup path; the remaining four tabs, interactive edit round trips and provider-template export checks still require representative client validation.
