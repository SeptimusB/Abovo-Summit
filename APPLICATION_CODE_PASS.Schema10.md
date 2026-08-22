# Historical schema-10 Abovo Summit application code pass

> Superseded on 22 August 2026 when `SampleTestFileXLSB.xlsb` became the authoritative master and the migration/materialisation implementation was removed.

Date: 18 August 2026

## Scope

- Reviewed the VB.NET 4.8 project definition and all 150 compiled files (112 non-designer source files, approximately 62,600 lines).
- Used `Z:\Sandbox\TestFileMigrated.xlsb` as the baseline workbook. Its repository copy at `Library\TestFileMigrated.xlsb` has the same recorded SHA-256 hash.
- Reconciled the workbook analysis with the current application source.
- Removed project-source `SystemLog` and `Debug.Print` logging and archived high-confidence repository debris.

## Baseline workbook contract

The baseline is a schema-10 Excel Binary Workbook with 297 worksheets, 1,772 defined names, approximately 579,605 formulas, and no external workbook links or data connections. It contains the worksheets, names, migration metadata, template cache, and Development Identified mirror sheets required by the current application.

The embedded VBA project was not exposed by the read-only workbook analysis and remains outside the audited scope.

## Runtime architecture

1. `Application.vb` creates `FormMainScreen`.
2. The main form selects or creates a workbook and delegates opening to `FileManager.OpenModel`.
3. `FileManager.ExcelModel` owns the DevExpress spreadsheet document and the per-model service graph.
4. `WorkbookContractValidator` verifies `Global Assumptions`, the `A8` marker, and `Transactional DB` before model services initialize.
5. Post-load initialization creates calculation, history, change, migration, transactional-database, structure, data, presentation, event, and interface services.
6. `StructureManager` deserializes `Structure.xml` (or supplied XML/path) into the UI/data contract.
7. `ProcessAsAbovoBP` reads model metadata and creates the file interface.
8. Calculation is switched to manual, chain-based, multithreaded operation and coordinated through `EngineManagement` and `CustomCalcEngine`.
9. `WorkbookStructureRuleManager`, `WorkbookManager`, `TransactionalDBSynchroniser`, `TransactionDBMaterialiser`, and `WorkbookMigrationManager` maintain named-range geometry and the Transactional DB compatibility/mirror contract.

## Cleanup performed

- Removed 820 textual `SystemLog`/`Debug.Print` occurrences from compiled project source, including commented calls and the obsolete logger implementation.
- Removed the now-empty migration status call/routine left after diagnostic output was deleted.
- Retained user-facing errors, `AbovoTransaction` failure propagation, and the in-memory application log (`AbovoAppCls.WriteLog`), because these are operational behaviour rather than debug-console output.
- Moved 62 excluded or superseded files (about 8.35 MB) to `Archive\Legacy Repository Files`, preserving their relative paths. These comprise backup folders, `.bak` resources, ZIP snapshots, duplicate excluded forms/classes, copied source variants, alternate structure XML files, and scratch notes.
- Kept the live project sources/resources, `Structure.xml`, the baseline XLSB, and its analysis in their original locations.

## Current risks and recommendations

1. **Create New workflow deferred:** `FormMainScreen.NewBP` is intentionally blank and marked for redesign. The former `Templates\DefaultBPTemplate.xlsb` path remains defined but is currently unused and the template is absent.
2. **Release migration policy:** `ApplyPendingMigrations` is compiled only under `#If DEBUG`. The schema-10 baseline needs no upgrade, but Release builds will not migrate older compatible workbooks.
3. **Error suppression:** compiled source still contains 26 active `On Error Resume Next` statements and 91 empty `Catch` blocks. Many surround optional workbook probes, but they materially reduce diagnosability and should be replaced incrementally with narrow exception handling and transaction failures.
4. **Worksheet protection:** the workbook protection value is application data, not a security boundary; it appears in source/structure configuration.
5. **No automated tests:** there is no test project. Add at least contract tests for opening/saving the baseline, structure deserialization, schema detection, named-range resize rules, and Transactional DB materialisation.
6. **Large mixed-responsibility classes:** `DataInterfaceTemplate`, `WorkbookMigrationManager`, and `DataManager` are particularly large. Further changes should extract focused services before adding more behaviour.
7. **Dormant placeholder:** `ChangeLogManager.DataChange` contains a `NotImplementedException`; no active call site was found, but the placeholder should be implemented or removed before that feature is wired in.

## Verification

- No `SystemLog(...)` or `Debug.Print` call remains outside the archive.
- Visual Studio MSBuild Debug build succeeds with 0 errors and 0 warnings.
- The project contains no remaining project-excluded `.vb` or `.resx` file outside the archive.
