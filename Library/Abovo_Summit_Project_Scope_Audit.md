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
- The native dashboard selector treats `Multivariable Dashboard!E6` as an authoritative workbook input, respects its lock state, and routes changed selections through `ModelChangeManager`.
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
- The remaining native Stress Test grids (Sensitivity, Dashboard summaries, and Comparative summaries/charts) are intentionally read-only. Their workbook-linked cell appearance remains presentation-only and does not create an edit path.
- Debug and Release builds both completed successfully after these changes. Interactive workbook integration testing remains required because the repository has no automated UI/workbook test project.

## Current project risks and missing assets

1. Workbook/VBA behavior remains part of the product contract; a purely VB.NET review cannot establish compatibility.
2. Worksheet protection credentials stored in source/XML are operational controls, not security boundaries.
3. `Templates/DefaultBPTemplate.xlsb` is referenced but absent, leaving Create New incomplete.
4. Pending workbook migrations run during both Debug and Release full-model loads. Migration errors fail the open, and in-memory migration changes remain subject to the normal explicit-save workflow.
5. Developer VBA tooling can rewrite projects/files and should be isolated from normal runtime entry points.
6. The project has a large legacy/duplicate compiled surface and no automated integration tests.

## Durable-audit files

- `TestFileMigrated.xlsb` — unchanged source workbook.
- `TestFileMigrated.analysis.md` — workbook compatibility summary.
- `Abovo_Summit_Project_Scope_Audit.md` — this architecture, VBA, workbook, and risk narrative.
- `Abovo_Summit_Project_Index.json` — source-free machine-readable worksheet, VBA-module/procedure/hash, compiled-file, and XML index.
- `Abovo_Summit_Codex_Context_Pack.zip` — compact context pack for future Codex tasks.
