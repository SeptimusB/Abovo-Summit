# Abovo Summit project scope audit

Audit date: 2026-08-18

This is the durable, compact record of a read-only inspection of `TestFileMigrated.xlsb`, its embedded VBA project, `Structure.xml`, and every compiled file in `Abovo Business Suite.vbproj`. The machine-readable companion is `Abovo_Summit_Project_Index.json`. The index contains hashes, topology, procedure names, and static metrics, but deliberately contains no workbook payload, VBA source, or passwords.

## Executive scope

Abovo Summit is a hybrid financial-modelling system with three mutually dependent layers:

1. `TestFileMigrated.xlsb` is the live business-plan model and remains the source of truth for assumptions, formulas, output layouts, named ranges, protection state, transactional data, and many user-visible workflow rules.
2. The embedded VBA project remains an executable compatibility and operational layer. It contains startup/close behavior, mode switches, imports, structural expansion, protection, stress testing, and transactional-database mirroring.
3. The VB.NET WinForms application loads the workbook into DevExpress Spreadsheet, deserializes `Structure.xml`, builds native interfaces over workbook ranges, manages edits/calculation/presentation, and progressively replaces or orchestrates legacy VBA workflows.

Consequently, workbook-first behavior is a project invariant. Interface controls should reflect workbook values, formats, locks, names, and calculations. User-originated cell edits should pass through `ModelChangeManager`; structural or batch workbook operations should use the established workbook/rule/migration services and preserve the workbook's observed protection and application state. An unprotected worksheet in this development workbook is not, by itself, a defect.

## Audit method and integrity

- Opened the XLSB through Microsoft Excel automation with macros and events disabled, alerts suppressed, calculation set to manual, and the workbook opened read-only. It was closed without saving.
- Counted formulas, constants, and stored Excel errors from bulk `Formula` and `Value2` arrays rather than relying on `SpecialCells`.
- Parsed the MS-OVBA streams directly with `oletools`. The VBE project password was not required because project-view protection does not encrypt the stored module streams.
- Parsed all 146 `<Compile>` entries in the current project, plus `Structure.xml`, and traced the startup, load, model-service, data/edit, calculation, interface, migration, structure, transactional-DB, import, and Stress Test paths.
- Reviewed commits `bdfe09a` and `b416334` and the current uncommitted `StressTest.vb` refinement.
- No VBA was executed. No workbook was saved, unlocked, or rewritten. The workbook hash remained the audit identity below.

## Workbook inventory

| Measure | Current result |
| --- | ---: |
| File size | 8,578,841 bytes |
| SHA-256 | `7039BD6E7BAE82C0269F1F9E56D4C28E17EED7599FFB263169798D214AABF044` |
| Excel file-format code | 50 (`.xlsb`) |
| Worksheets | 297 |
| Visibility | 279 visible, 2 hidden, 16 very hidden |
| Protected worksheets | 35 |
| Defined names | 1,772; none broken |
| Formula cells | 649,034 |
| Non-empty constant cells | 90,621 |
| Stored error cells | 467 |
| Shapes / charts | 1,148 / 36 |
| Tables / pivot tables | 0 / 0 |
| External workbook links | 0 |
| Data connections | 0 |

The former estimate of approximately 579,605 formulas is superseded by the 649,034-cell bulk Excel-COM count. Array/shared formula members are counted as formula cells because each occupied model cell participates in the formula surface.

The largest formula surfaces are `Transactional DB` (78,404), `Dvpt NonCash` (34,953), `Development Capital` (30,033), `FFR Workings` (12,983), `Hidden - Tenure Totals Start` (11,112), `Development Revenue` (10,677), and `Development Expenditure` (10,276). The 14 very-hidden `__Abovo_TransDB_DevID_A` through `_N` sheets are migration/materialisation mirrors rather than ordinary UI sheets. `__Abovo_Metadata` and `__Abovo_Templates` are also very hidden.

The 467 stored errors reconcile with the earlier audit and are all in four locations:

- `OW - Live Covenant Calculation`: 240
- `OW - Covenant Calculation`: 131
- `OW - Charts Source Data`: 95
- `Hidden - Tenure Totals Start`: 1

The covenant/chart errors are consistent with unavailable-series suppression. The single tenure-totals lookup error still warrants validation against a representative populated model before being declared expected.

Current workbook contract values remain Business Plan mode, model version 25.0708, start date 1 April 2025, and stress-test mode `N`. `Transactional DB` is the dominant data surface and is coupled to dedicated synchronisation and materialisation services in both VBA and VB.NET.

## Embedded VBA project

The workbook contains 334 VBA components and 18,226 extracted source lines:

| Component kind | Count |
| --- | ---: |
| Standard modules (`.bas`) | 47 |
| Class/document modules (`.cls`) | 281 |
| UserForms (`.frm`) | 6 |
| Identified procedures | 286 |

Most worksheet document modules contain only attributes or small event handlers. The principal operational modules are:

- `DSA_Import.bas` — file/folder imports, rejection reporting, new imported/total sheets, formula/list maintenance, templates, and date stamping.
- `Dvpt_Columns.bas`, `Funding_Columns.bas`, `Repairs_Columns.bas`, `CapGrant_Columns.bas`, `Component_Columns.bas`, `CapExpend_Columns.bas`, `Interco_Funding.bas`, `OFA_Module.bas`, `Joint_Venture.bas`, and related modules — physical model expansion/contraction and formula propagation.
- `Stress_Testing_Module.bas` — stress-mode switching, live and scenario capture, base/scenario imports, dashboard generation, formula modification, and sensitivity capture/clear/delete.
- `Updates_Module.bas` and `Protection_Module.bas` — workbook upgrade/support and sheet protection behavior.
- `Auto_Start_Close_Module.bas` — `Auto_Open`, `Auto_Close`, valuation/business-plan/stress mode checks, recalculation/timing, and version warnings.
- `Summit_Compatibility.bas` — full and domain-specific Transactional DB mirror sizing/synchronisation.
- `Calc_Event_State_Module.bas` — application calculation/event/screen state management.
- `Menu_Module.bas`, `Colouring_Module.bas`, `Validation_Module.bas`, `FFR_Module.bas`, and `Global_Functions.bas` — legacy UI, formatting, validation, extraction, shared logging/state utilities.
- `CodeManagement.bas` — developer-only module export/import/recode/copy utilities that can modify VBProjects and files.

The complete source-free module/procedure/hash inventory is in the JSON index. The largest modules are `DSA_Import` (1,104 lines), `Dvpt_Columns` (805), `Stress_Testing_Module` (800), `Updates_Module` (704), `Global_Functions` (526), `Funding_Columns` (456), `Columns_Fill_Module` (371), `Summit_Compatibility` (360), `Protection_Module` (349), and `Auto_Start_Close_Module` (339).

### VBA privilege and safety findings

The scanner found expected auto-entry points (`Auto_Open`, workbook events, worksheet events, button clicks, and hyperlink handlers). Generic suspicious keywords mostly arise from normal workbook automation, but these active behaviors deserve explicit boundaries:

- `CodeManagement.bas` exports, imports, deletes, and rewrites VBA components; saves/re-encodes workbooks; deletes temporary module files; and creates `WScript.Shell`/FileSystemObject objects. It requires Trust Center access to the VBA object model and should be treated as development tooling, not an end-user runtime path.
- `FFR_Module.bas` saves an extraction workbook to a user-selected path.
- `Global_Functions.SystemLog` appends to a log file.
- Import modules open user-selected workbooks and copy values, formats, formulas, sheets, and named-range data into the current model.
- No active shell-command execution was found despite the scanner's generic `Shell` keyword alert.

The only URL indicators are the regulator portal, an old spreadsheet-formatting reference, and Microsoft documentation. No external data connection exists in the audited workbook.

## VB.NET application project

`Abovo Business Suite.vbproj` is a .NET Framework 4.8 WinExe with `My.MyApplication` as startup object, `Option Explicit On`, `Option Strict Off`, and `Option Infer On`. The current project compiles 146 files:

| Source measure | Lines |
| --- | ---: |
| Non-designer source | 63,911 |
| Designer source | 10,743 |
| Total compiled VB | 74,654 |

The earlier 150-file/~62,600-line figures were approximate and are superseded by this project-file-driven inventory. DevExpress 25.2 supplies Spreadsheet, Grid, Charts, Printing, PDF, RichEdit, navigation, and UI services. WebView2, protobuf-net, and Microsoft configuration packages are also referenced.

Largest non-designer files are `DataInterfaceTemplate.vb` (11,613 lines), `WorkbookMigrationManager.vb` (5,999), `DataManager.vb` (4,514), `StressTest.vb` (4,272), the still-compiled `DataInterfaceTemplateOld.vb` (4,040), `TransactionalDBSynchroniser.vb` (2,168), `BPIncomeExpenditureAnalyser.vb` (2,128), and `WorkbookStructureRuleManager.vb` (1,804).

### Runtime path

1. `Application.vb` starts `FormMainScreen`.
2. The main form validates the selected model path and delegates to `FileManager.OpenModel`.
3. `FileManager.ExcelModel` owns the DevExpress `SpreadsheetControl`/workbook and constructs `StructureManager`, `DataManager`, the calculation engine, `InterfaceManager`, `PresentationManager`, event services, Transactional DB services, dependency management, migration management, and workbook-structure rule management.
4. The workbook loads first. Summit validates the required workbook contract, installs custom calculation/change services, loads metadata, deserializes `Structure.xml`, and then builds the user interface over workbook objects.
5. `DataManager` interprets XML data-source definitions and workbook names/ranges. `PresentationManager` builds presentation sections. `DataInterfaceTemplate`/related interfaces bind DevExpress controls to those sources.
6. `ModelChangeManager` applies typed values, logs change events, marks the model dirty, and requests calculation through the workbook calculation engine. It rolls back a failed write.
7. `WorkbookManager` and `WorkbookStructureRuleManager` own physical row/column/range structural operations and semantic rules. `WorkbookMigrationManager` owns schema upgrades and mirror/template metadata.
8. `TransactionalDBSynchroniser` and `TransactionDBMaterialiser` maintain the workbook's Transactional DB/mirror relationship.
9. Save uses the loaded document's `SaveDocument` path, preserving the workbook format. Close disposes interfaces, dependencies, spreadsheet controls, and model services.

`Structure.xml` currently hashes to `F90B7004FDAD627D175A4E280BDFA1DC10FA2DF5C2D9362F511B9FA5F4397645`. The static index sees 672 `Worksheet` elements, 1,182 data-field definitions, 747 cell-range data sources, and 467 distinct named-range/repository tokens. After trimming XML whitespace, all 184 distinct worksheet values found by the current parser exist in the audited workbook. This is a parser-normalisation refinement over the earlier 186-token count, not evidence of missing sheets.

### Static maintainability signals

The full compiled-file scan found 16 line-start `On Error Resume Next` occurrences, 119 conservatively detected empty `Catch` bodies, one `NotImplementedException`, and 334 direct cell/formula/clear/copy write tokens. These are triage signals, not defect counts: many direct writes are intentional migrations or structural transactions, and several catches are best-effort UI cleanup. The concentration is more useful than the total:

- `WorkbookMigrationManager` contains 134 direct-write tokens and 38 empty catches.
- `DataInterfaceTemplate` contains 25 empty catches and is the largest runtime/UI file.
- `TransactionalDBSynchroniser` contains 18 empty catches.
- `StressTest` contains 31 direct-write tokens, most in capture/generation operations rather than interactive single-cell editing.
- `DataInterfaceTemplateOld.vb` remains compiled, increasing ambiguity and maintenance surface.
- The sole `NotImplementedException` is in `ChangeLogManager`; no active caller was established in this pass.

There is no automated test project. Debug build success is therefore necessary but not sufficient; representative workbook integration tests remain important for structural, migration, calculation, protection, and persistence paths.

## VBA-to-VB.NET relationship

Exact procedure-name overlap is limited because the VB.NET layer often re-expresses VBA behavior as services and UI handlers. Confirmed direct-name bridges include stress-mode/capture routines, `FFR_New_Extraction`, DSA folder/update routines, `DoesNRExist`, full Transactional DB sync, and row insertion. Semantically, the overlap is much broader:

- VBA structural column modules correspond to `WorkbookManager`, `WorkbookStructureRuleManager`, migrations, and transactional materialisation.
- VBA `Calc_Event_State_Module` corresponds to calculation/event state orchestration in the model services.
- VBA DSA/import flows correspond to VB import interfaces/services and workbook structural operations.
- VBA Stress Testing corresponds to `Interface/StressTest/StressTest.vb`, but workbook named ranges and calculation formulas remain authoritative.

Removing VBA solely because a similarly named VB routine exists would be unsafe. A replacement decision must be based on user entry points, workbook buttons/events, named-range side effects, protection changes, and imported-file behavior.

## Stress Test-specific conclusions

Commits `bdfe09a` and `b416334` materially expanded the native Stress Test interface and then aligned native editors with workbook locks/reference rows. The current working-tree refinement routes interactive scenario names, import modes, include flags, planner values, comparison selectors/start year, live capture slot, and live scenario name through `ModelChangeManager`. Failed or locked edits refresh from the workbook rather than leaving the UI ahead of model state. It also preserves each batch-touched sheet's entry protection state and marks direct multi-cell Stress Test operations dirty so they participate in the save workflow. This deliberately leaves intentionally unprotected development sheets unprotected.

The remaining direct writes in Stress Test fall into two groups:

1. Workbook operations such as mode switching, sensitivity capture, scenario clearing/capture, dashboard generation, range snapshot/restore, and result import. These are multi-cell commands and should remain workbook-first, but they need transactional/protection/failure-state scrutiny rather than being mistaken for ordinary editor changes.
2. A small number of presentation/control selections that select workbook outputs and may merit change logging if they represent persisted user inputs.

Important follow-up checks are:

- Ensure dashboard generation restores live assumptions, mode names, stress flags, title/cursor, and protection after cancellation or any import/calculation failure.
- Treat clear/delete/capture commands as explicit model transactions: confirm ranges and lock/reference-row rules, mark dirty, calculate through the established engine, and provide an undo/audit story where appropriate.
- Continue to reject UI edits to locked workbook cells and refresh native controls from the workbook after a rejected change.
- Validate the interface against a representative populated workbook, particularly sensitivity row insertion, scenario imports, comparative selectors, and the 10-scenario dashboard loop.

### 2026-08-18 implementation follow-up

- Full-model workbook migrations now execute in both Debug and Release. Migration failures abort opening the model; migration changes remain in memory until the user saves. Import-source workbooks do not enter this path.
- First-tab stresses and mitigation grids now capture the workbook value before the bound data source commits, restore it, and submit the edit through `ModelChangeManager`. Mitigations is presented as the workbook's one continuous `A52:H76` table rather than three independently sized grids, eliminating the artificial gaps and preserving workbook row order.
- The covenant selector suppresses its legacy automatic direct write and submits `Live Multivariable Planner!AD3` once through `ModelChangeManager`.
- Stress Test version 6.66 removes the rejected live-worksheet Tab #4. The interface no longer reparents or presents a `SpreadsheetControl`; it reproduces the `Multivariable Dashboard` with native DevExpress controls while retaining the workbook as the calculation authority.
- The native dashboard contains a workbook-driven scenario selector, sort controls, ten-year five-metric banded matrix with custom target/breach indicators, the Debt Information outputs, selected covenant status chart, legend, explanatory note, and four result cards. Changes to the linked dashboard control cells travel through `ModelChangeManager`, calculate through the established service, and refresh from the workbook. The exact cell/range and VBA contract is retained in `Multivariable_Dashboard_Technical_Audit.md`.
- Stress Test version 6.67 refines those dashboard selectors into purpose-built DevExpress controls: an exact first-year `SpinEdit`, sort-metric and covenant-chart `RadioGroup` controls, and a sort-order `ToggleSwitch`. This deliberately reproduces workbook behaviour without reproducing Excel's form-control presentation.
- Stress Test version 6.68 rebuilds Tabs #5 and #6 as native counterparts of `Comparative` and `Comparative 2`. Four shared scenario selectors, six shared series-visibility controls, five independent date-window editors, five workbook-sourced charts, and the two extrema summaries reproduce the sheets' linked-cell/formula contract without displaying a spreadsheet. Exact ranges, colours, and remaining client-data validation are recorded in `Comparative_Interfaces_Technical_Audit.md`.
- Stress Test version 6.69 gives the shared series checkboxes sufficient caption width and removes the repeated chart legends. The scenario names in the first column of each extrema summary now form the legend using the workbook's exact six series colours.
- Stress Test version 6.70 presents the five comparative start-year offsets as integer 0-20 sliders. Slider values still update the same workbook-linked calculation cells through `ModelChangeManager`, with recalculation deferred until a mouse drag or keyboard adjustment finishes.
- Stress Test version 6.74 applies one selection/clipboard policy across every standard, custom, and banded grid. Editable and read-only grids support multi-cell selection and Ctrl+C without column headers; the Sensitivity grid retains multi-row selection so its existing selected-record deletion remains row-safe. Clipboard copy is presentation-only and does not write to the workbook.
- Stress Test version 6.78 presents `Live Multivariable Planner!B5` as `Multivariable Name:` inside the Tab #1 Capture scenario panel, alongside the Record as test selector; the separate Quick Capture panel now contains only its button. The editor loads from the authoritative workbook cell and commits through `ModelChangeManager` before capture; the selected Test continues to receive the name in its workbook row-8 cell (`I8`, `N8`, through `BB8`) as part of the established capture transaction.
- Every Stress Test grid now applies `ShowForFocusedRow` at the shared view-policy level, so repository editor buttons are not permanently visible.
- Stress Test batch calculation now uses the existing workbook calculation service. Scenario capture and ten-scenario generation temporarily use DevExpress recursive calculation, restore the previous engine in `Finally`, and calculate before copying `StressLiveInfo` results.
- Dependent stresses-grid cells whose controlling value is zero are now non-editable, matching their disabled visual treatment.
- Stress Test grid editability now uses the same workbook signal as `DataInterfaceTemplate`: the linked cell must be unlocked and its current fill pattern must be `Solid`. This rule is enforced for both first-tab input grids and the native Multivariable Planner when an editor opens and again before a grid edit enters `ModelChangeManager`, so recalculated conditional-format state remains authoritative.
- First-tab editor selection is derived from each linked workbook cell. The three workbook flags (`D35`, `D65`, and `D66`) use the XLSB's exact `Y`/blank list; percentage, whole-number, and decimal editors follow the cell number format; rows whose year validation starts at 2 use the 2-40 list, while the remaining active year inputs use 1-40. Locked or non-solid-pattern cells never acquire an active editor.
- First-tab Stress and Mitigation presentation is backed by an in-memory snapshot rather than an editable `RangeDataSource`. This prevents the grid from writing the XLSB before change management. Accepted edits now travel once through `ModelChangeManager`, calculate through the existing engine, and then reload the snapshot from the authoritative workbook; rejected edits simply reload workbook state. The former restore-and-replay event sequence was removed because it could revert a valid editor change.
- Planner presentation now mirrors DIT's two-stage rule instead of deriving a display colour from the pattern. `Fill.PatternType` determines editability (`Solid` is active; every other pattern is rule-locked). Active cells retain the workbook's `Fill.BackgroundColor`; rule-locked editable columns receive the standard Lavender/WhiteSmoke disabled appearance. Representative workbook verification produced `I10=Solid/editable`, `J10=DarkGray/rule-locked`, `I35=Solid/editable`, and `J35=None/rule-locked`.
- The native Planner renders Test/Scenario 1 through 10 simultaneously in a horizontally scrollable `BandedGridView`, with fixed Assumption/Short Name columns and one five-column band per workbook scenario (`Assumptions1`/`AssumptionsA1`, columns `I:M`, through `Assumptions10`). Each band header mirrors the workbook's Test number, scenario name, and import mode. The toolbar selector now chooses which scenario metadata or clear action is being edited; it no longer selects the only visible data block. The hidden `D:H` block and `S0Data` remain the Base Case dashboard/captured result, not an editable Planner scenario. Dashboard and comparative selectors may still expose Base Case.
- The native Stress Sensitivity List preserves the workbook's complete `A:BI` record rather than reducing it to summary and file columns. Its five ten-year metric groups are presented as bands; Target and Golden Rule are followed by an explicit visual spacer, then the live/Base row and captured rows. Cell text remains the workbook's `DisplayText`, while custom drawing uses the linked cell's resolved `FillColor` and font. Consequently, the XLSB's existing conditional-format rules remain the sole authority for red/amber/green displacement colours.
- Stress Sensitivity data cells remain workbook-coloured, but column headers use a presentation-only high-contrast light-grey/dark-blue bold palette so year identifiers remain readable at normal scaling. The Stress Test Windows UI navigation panel uses centred content alignment, keeping its buttons above the working area across window sizes.
- Stress Sensitivity column captions preserve workbook line breaks and wrap within a taller header row, allowing long summary and year-identification labels to display across multiple lines without widening their data columns.
- Stress Sensitivity header wrapping explicitly enables DevExpress text options and automatic column-header height. Stress navigation buttons and separators clear their designer-authored left-alignment flags before the panel applies centred content alignment.
- Stress Sensitivity summary columns use wider presentation widths than the repeated metric-year columns and all column captions align to the bottom of the wrapped header row. The centred navigation collection is reversed after clearing its left-docking flags, restoring Home as the leftmost action while retaining centred placement.
- The Stress Sensitivity Summary band scrolls with the remainder of the grid rather than remaining fixed. Each Tab #3 scenario band now embeds two header editors modelled on ColumnInplaceEditorHelper: the workbook list validation from Multivariable Planner row 7 and the free-text scenario name from row 8. Custom drawing and click handling share exact cached bounds, while committed metadata edits continue through ModelChangeManager and refresh back from the workbook.
- Tab #3 now leaves only its three workbook actions in the top toolbar; each Test band owns its row 7 import-mode and row 8 scenario-name editors, whose font and resolved workbook fill are drawn from their source cells. A further Copy From selector lists the other nine Tests and copies the source `AssumptionsN` and `AssumptionsAN` values into the destination in one workbook-native operation, restoring protection and performing one recursive calculation. Test names and import-mode metadata remain destination-specific. Stress Test grids suppress user filtering, sorting, grouping, and column menus, while the workbook-required Stresses/Mitigations grouping remains programmatic. The section captions now use the plural workbook concepts Stresses and Mitigations.
- The Tab #3 Test band header uses five logical DevExpress band rows, rather than relying on a single oversized row, so all three editor rows remain above the data-column captions. Custom-drawn workbook-linked editors explicitly transfer their repository appearance and repaint their content area after the skin painter, preserving the source workbook's resolved fill and font colours; standalone Copy From selectors use Abovo blue with white text, including their popup lists and active editors.
- Tab #3 is split into `Tests` and `Targets and Golden Rules` sub-tabs. `Tests` retains the ten-scenario planner. `Targets and Golden Rules` is a workbook-driven representation of `Multivariable Planner!BH6:BS48`: the two covenant groups, metric names, units, year rows, spacer, cell display formats, resolved fills, fonts, and protection state come from the workbook. Workbook row 8 is the first real grid row rather than a custom-painted column header, ensuring that the five Target dashboard labels use visible text editors and the five Golden Rule comparison directions use visible validation-list combo editors. Generated banded-grid columns are explicitly made visible after the bands are rebuilt; a full application runtime probe confirmed 40 rows, the row-8 editor values, years 1-39, and calculated covenant values render in the resulting grid. Target result rows remain read-only through their workbook locks, while workbook-unlocked solid-pattern Golden Rule values and all row-8 settings commit through `ModelChangeManager` and reload from the workbook.
- The remaining native Stress Test grids (Sensitivity, Dashboard summaries, and Comparative summaries/charts) are intentionally read-only. Their workbook-linked cell appearance remains presentation-only and does not create an edit path.
- Debug and Release builds both completed successfully after these changes. Interactive workbook integration testing remains required because the repository has no automated UI/workbook test project.

## FFR-specific conclusions

Version 6.71 replaces the disposable legacy `FFRForm` implementation with lazy native DevExpress presentations of the authoritative `FFR Validation Summary`, `Front Sheet`, `FFR Inputs Adj Stmt`, `FFR Workings`, `Statements`, `Assumptions & tenure inputs`, `Compliance Questions`, and `FFR Key Defn` worksheets. No `SpreadsheetControl` is hosted or exposed.

- Every tab uses an in-memory snapshot rather than an editable workbook `RangeDataSource`. The live workbook remains the source of displayed values, resolved colours, font emphasis, alignment, row heights, column widths, cell notes, protection, and validation lists.
- Interactive editing requires an unlocked workbook cell with a `Solid` fill pattern. Accepted text, date, list, percentage, and numeric edits travel once through `ModelChangeManager`; the calculated workbook is then reloaded into the active view. `Statements` and `Assumptions & tenure inputs` are read-only because the audited workbook contains no unlocked cells on those sheets.
- The native `Create FFR return` command retains the `FFR_New_Extraction` mapping contract through `FFRRangeNames` and `FFRListHeading`, copies values into a separate provider-specific macro-enabled template, and never modifies the active XLSB.
- Exact ranges, workbook counts, edit rules, extraction boundaries, and required client round-trip checks are recorded in `FFR_Interfaces_Technical_Audit.md`.
- `Front Sheet` is now a purpose-built native form rather than a literal worksheet grid: return details remain linked to `B5:B7`, the second confirmation remains linked to `B36`, and the registered/non-registered 25-row entity lists remain linked to `B10:C34` and `B39:C63`. The two lists are presented side by side while retaining workbook validation, colours, locks and change management.
- Debug startup auto-opens `Z:\Sandbox\TestFileMigrated.xlsb` when present by passing it to the existing `FormMainScreen.OpenModelProceedureBP` routine. The hook is compile-time excluded from Release.
- Version 6.71 passed Debug and Release builds on 19 August 2026 with zero warnings and zero errors. Interactive eight-tab and provider-template validation remains required in the normally started application.
- Version 6.72 replaces the worksheet-scale FFR Validation Summary surface with a centred compact view of workbook rows 5-28. Production instructions remain workbook-driven, while the hard-error (A18:C23) and soft-notification (A25:C28) blocks use read-only DevExpress grids with the established Stress Test filtering, sorting and grouping restrictions. Runtime verification reproduced all current workbook counts and messages; Debug and Release builds passed on 20 August 2026.
- Version 6.73 removes unnecessary vertical scrollbars from the two compact validation grids and standardises all FFR grids on cell multiselect with header-free clipboard copying. Read-only workbook cells remain selectable for copying without becoming editable. Debug and Release builds passed on 20 August 2026.
- Version 6.99 added the first native Workings prototype at `Workings > Stock > Target Rent Letting Rates`. The read-only LiveGrid resolves its width from the two fixed fields plus the live `StockType` named range, refreshes after calculation, rebuilds through worksheet/named-range dependency invalidation, preserves workbook cell display/formatting, and supports cell multiselect with header-free copying. Read-only DevExpress inspection confirmed `StockType` as `Stock Assumptions!E5:X5` (20 columns) and the source block as `Target Rent Letting Rates!A9:V48` (40 rows by 22 columns, with no hidden rows/columns). Although Debug and Release builds passed, the first interactive open exposed an `IndexOutOfRangeException` because common dataset post-processing assumed the direct LiveGrid adapter materialised `DataRows`.
- Version 7.0 trims the LiveGrid's chunk-grown column definitions before resolving its structural width and returns the completed direct-range dataset before materialised-row post-processing. Debug and Release builds passed on 21 August 2026; interactive open, side-by-side refresh, structural stock-type expansion and Excel/VBA round-trip validation remain required.
- Version 7.01 makes the Target Rent Letting Rates prototype visually follow its compact workbook block. Column captions now concatenate source rows 6 and 7 with multiline wrapping, body rows use compact fixed height/padding, source column widths are converted to interface pixels, and Excel General numeric alignment resolves to the right. Debug and Release builds passed on 21 August 2026; interactive visual comparison and the remaining side-by-side/structural/round-trip checks are still required.
- Version 7.02 refreshes LiveGrid column captions from their workbook formula cells after each calculation, so stock-type renames update in place while structural named-range changes continue to rebuild the interface. The Target Rent Letting Rates child structure now declares its underlying worksheet for the worksheet button. Debug and Release builds passed on 21 August 2026. Interactive rename, stock-type addition, worksheet-button and Excel/VBA round-trip validation remain required.
- Version 7.03 adds the second native Workings child, `Workings > Stock > Target Rent Letting Numbers`, using the workbook-defined `Workings_TargetRentLettingNumbers` block (`A9:X48`). Its read-only LiveGrid presents three fixed year/proportion columns, the `StockType`-driven unit columns, and the workbook-calculated trailing total; captions, values and formatting remain workbook-owned and the worksheet button maps to `Target Rent Letting Numbers`. Debug and Release builds passed on 21 August 2026; interactive visual, recalculation, structural stock-type and round-trip validation remain required.
- Version 7.04 converts the remaining Stock Workings children to read-only workbook-native LiveGrids: Demolition Numbers, RTB Numbers, Other Disposal Numbers, Stock Conversion Numbers, Development Stock, Development Stock Numbers, Service Charge Numbers and Leaseholder Units. Multi-area `Workings_*` names can now project selected workbook areas with shared leading year columns, preserving compact assumption-style tabs without copying data or exposing intervening helper columns. The delivery contains 22 tabs across those eight children, including previously omitted Other Disposal totals and Stock Conversion first-tranche sales; formula-driven captions, cell display/formatting, recalculation refresh, structural invalidation and worksheet-button mappings remain workbook-owned. Debug and Release builds passed on 21 August 2026; interactive tab-by-tab visual, recalculation, structural expansion and Excel/VBA round-trip validation remain required.
- Version 7.05 prevents interface layout/best-fit from passing a null worksheet key for deliberately synthetic legacy/unbound data cells. `GetDSData` now returns the cell's cached typed dataset value when no workbook source sheet/address exists; workbook-linked cells continue to read their authoritative live source. This fixes the `ArgumentNullException` exposed when opening the third Workings interface without creating a write path or masking workbook-linked failures.
- Version 7.06 replaces LiveGrid's worksheet `RangeDataSource` binding with a detached read-only `UnboundSource` projection. Multi-area Workings names can now reuse shared Year columns across adjacent tabs without DevExpress attempting to bind the same worksheet cells twice; every displayed value, caption and cell style is still resolved from its original live workbook cell, and calculation refresh continues to request current workbook values. Debug and Release builds passed on 21 August 2026; interactive Demolition tab switching, recalculation, structural expansion and Excel/VBA round-trip validation remain required.
- Version 7.07 restores the LiveGrid projection coordinates to the compiled `AbovoUnboundSourceTag` after a stale source buffer removed those members while leaving their consumers in place. Full Debug and Release Rebuild targets passed on 21 August 2026; Visual Studio design-time compilation and interactive Demolition tab switching remain required.
- Version 7.08 completes the native Workings prototype across the authoritative 138-child navigator list. The inaccurate duplicate RTB Valuation Factors child was removed; Variable Management Costs, Bond Premium Amortisation, all 12 Accounts children, and the other omitted entries were added; CSIDs are sequential 0-137; and all 320 sections are read-only LiveGrids with explicit worksheet-button mappings. Of those sections, 287 resolve workbook defined names, 31 use validated direct worksheet ranges for sheets without suitable names, and the original two Target Rent prototypes retain their established dynamic XML definitions. Multi-block names are lazy split into workbook-titled tabs, including 22 Development Capital and 50 Dvpt NonCash blocks. LiveGrid now supports XML direct ranges, chooses the last aligned row block from names that retain earlier areas, and discovers formula-driven headers from the nearest workbook heading cells. The read-only generator and full mapping/validation contract are recorded in `Workings_Interface_Technical_Audit.md`.
- Version 7.10 implements Outputs children 0-20 through Business Plan Dashboard. Stock and traditional statement views use detached read-only LiveGrids; cashflow and alternative account views use a new detached read-only LiveVGrid projection with workbook measures down the page and periods across it. Both paths refresh values and formula-driven headings after calculation, preserve workbook formatting, multiselect and clipboard copy, and avoid worksheet range binding. BP Dashboard retains the established native special interface. Exact child order, ranges, layout choices and manual checks are recorded in `Outputs_Interface_Technical_Audit.md`; `Tools/GenerateOutputsStructure.ps1` reproducibly generates only the Outputs XML group.
- Version 7.11 removes empty presentation artefacts from the detached worksheet projections: fully blank `DisplayText` records are excluded, fully blank and unheaded fields are hidden, and LiveVGrid record-header helpers are hidden after DevExpress attaches them so fallback rows such as `Col_1` cannot reappear. Meaningful unlabeled data is retained. LiveGrid and LiveVGrid values remain workbook `DisplayText`, while numeric, leading-minus and parenthesised negatives are rendered red. Debug and Release rebuilds passed on 22 August 2026 with zero warnings and zero errors; interactive recalculation and visual comparison remain required.
- Version 7.12 returns Outputs cashflow children 3-11 and Alternative Cashflow Statement to detached horizontal LiveGrid presentation, preserving the source worksheet orientation like the stock interfaces; only the three non-cashflow alternative account views remain vertical. Output LiveGrids may use the current monitor working height rather than the common 70% cap, allowing Summary Comp Inc - Trad View to show its compact report. BP Dashboard now participates in the established calculation refresh registry, rebuilding its workbook-backed charts and table after completed calculations while open and unregistering on disposal. Debug and Release rebuilds passed on 22 August 2026 with zero warnings and zero errors; interactive height, cashflow orientation, dashboard refresh and Excel/VBA round-trip validation remain required.
- Version 7.13 returns the remaining Alternative account views to detached horizontal LiveGrid presentation, so Outputs contains no vertical grids, and adds the native workbook-backed Funding Dashboard as Outputs child 21. Its four validation-driven selectors (`D6`, `I6`, `R6`, `V6`) commit through `ModelChangeManager`; its four funding charts, selected covenant table/chart and two fixed status charts rebuild from the calculated workbook after changes. Read-only Excel inspection left `TestFileMigrated.xlsb` unchanged at SHA-256 `7039BD6E7BAE82C0269F1F9E56D4C28E17EED7599FFB263169798D214AABF044`. Full standard Debug and Release rebuilds passed on 22 August 2026 with zero warnings and zero errors. The exact control, source-range and manual round-trip contract is recorded in `Funding_Dashboard_Technical_Audit.md`.
- Version 7.14 fixes the Funding Dashboard opening failure in the selected-covenant table. The custom DevExpress `GridView` is installed before binding its snapshot, then explicitly populates and safely resolves the `Year` and `Value` fields. This keeps workbook styling and clipboard behaviour while eliminating the null column lookup created when the data source had initially populated a temporary default view. Full standard Debug and Release rebuilds passed on 22 August 2026 with zero warnings and zero errors.
- Version 7.15 restores the Funding Dashboard Operating Margin and EBITDA MRI status markers by retaining their workbook colours and expanding the hidden status axis to include the source value `1.1`. The selected covenant chart now reproduces the workbook's calculated forecast, target and breach-overlay series from `OW - Charts Source Data!AI`, `AJ` and `AL` instead of charting only the adjacent display-table value. Full standard Debug and Release rebuilds passed on 22 August 2026 with zero warnings and zero errors.

- Version 7.16 corrects the Funding Dashboard chart-opening exception by using native DevExpress point series for target, covenant-met and covenant-breached markers. This preserves the marker-only workbook presentation without applying DashStyle.Empty to an ordinary series line, an unsupported DevExpress 25.2 combination. Full standard Debug and Release rebuilds passed on 22 August 2026 with zero warnings and zero errors.

- Version 7.17 replaces the Funding Dashboard's selected-covenant grid/chart split with one read-only DevExpress XtraGrid. Workbook year and percentage DisplayText remain visible, while a row-aligned custom visual column draws the calculated forecast bar, breach colour and target marker from OW - Charts Source Data AI, AL and AJ. The complete grid snapshot rebuilds through the existing dashboard calculation refresh. Full standard Debug and Release rebuilds passed on 22 August 2026 with zero warnings and zero errors.

- Version 7.18 removes the selected-covenant Target and Exceeded helper fields from the visible XtraGrid and its customization surface. Funding Dashboard chart cards now allocate separate header and content rows so legends remain unobscured, while the two lower status charts centre the workbook's 1.1 marker series within a fixed vertical range. Full standard Debug and Release rebuilds passed on 22 August 2026 with zero warnings and zero errors.

- Version 7.19 disables automatic column population in the selected-covenant XtraGrid and creates only its three presentation columns explicitly. Target and Exceeded remain bound-row values used by the custom forecast renderer, but they no longer exist as GridColumn objects and therefore cannot reappear in the interface. Full standard Debug and Release rebuilds passed on 22 August 2026 with zero warnings and zero errors.

## Current project risks and missing assets

1. Workbook/VBA behavior remains part of the product contract; a purely VB.NET review cannot establish compatibility.
2. Worksheet protection credentials stored in source/XML are operational controls, not security boundaries.
3. `Templates/DefaultBPTemplate.xlsb` is referenced but absent, leaving Create New incomplete.
4. Pending workbook migrations run during both Debug and Release full-model loads. Migration errors fail the open, and in-memory migration changes remain subject to the normal explicit-save workflow.
5. Developer VBA tooling can rewrite projects/files and should be isolated from normal runtime entry points.
6. The project has a large legacy/duplicate compiled surface and no automated integration tests.

## Durable-audit files

Library/Abovo_Summit_Architecture_Review.md is the durable review of the workbook, Structure.xml, DataManager, PresentationManager and DataInterfaceTemplate ownership boundaries. Use it with this audit and AGENTS.md before workbook-facing architectural changes.

- `TestFileMigrated.xlsb` — unchanged source workbook.
- `TestFileMigrated.analysis.md` — workbook compatibility summary.
- `Abovo_Summit_Project_Scope_Audit.md` — this architecture, VBA, workbook, and risk narrative.
- `Abovo_Summit_Project_Index.json` — source-free machine-readable worksheet, VBA-module/procedure/hash, compiled-file, and XML index.
- `Abovo_Summit_Codex_Context_Pack.zip` — compact context pack for future Codex tasks.
