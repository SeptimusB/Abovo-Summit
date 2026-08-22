# Historical TestFileMigrated.xlsb schema-10 analysis

> Superseded on 22 August 2026 when `SampleTestFileXLSB.xlsb` became the authoritative master. This analysis is retained only as evidence of the removed migration/materialisation design.

Analysis date: 2026-08-17; full workbook/VBA refresh: 2026-08-18

## Provenance

- Repository copy: `Library/TestFileMigrated.xlsb`
- Supplied source: `Z:\Sandbox\TestFileMigrated.xlsb`
- Size: 8,578,841 bytes
- SHA-256: `7039BD6E7BAE82C0269F1F9E56D4C28E17EED7599FFB263169798D214AABF044`
- The source and repository-copy hashes match.

## Workbook inventory

- Excel Binary Workbook (`.xlsb`, Excel file-format code 50)
- 297 worksheets: 279 visible and 18 hidden/very hidden
- 1,772 defined names; no defined name contains `#REF!`
- 649,034 formula cells and 90,621 non-empty constant cells (bulk Excel-COM count; supersedes the earlier formula estimate)
- No external workbook links
- No data connections
- Embedded `vbaProject.bin` present (1,335,296 bytes)
- Embedded VBA project audited without execution: 334 components, 18,226 extracted source lines, and 286 identified procedures

Key workbook values:

- Mode: Business Plan
- Model version: 25.0708
- Business-plan start date: 1 April 2025
- Stress-test mode: N
- Transactional records: `Transactional DB!A6:BV1599`

## Abovo Summit compatibility

The workbook satisfies the application's initial validity check:

- `Global Assumptions` exists
- `Global Assumptions!A8` is `Business Plan Start Date`
- Company name is available at `Global Assumptions!C6`
- Start date is available at `Global Assumptions!C8`
- `Transactional DB` exists

Comparison with `Structure.xml` found:

- All 184 normalised distinct worksheet values found by the current parser exist in the workbook (the earlier 186-token result used a different tokenisation pass)
- All genuine XML-referenced defined names exist (the XML token `CR` denotes direct cell-range addressing rather than a workbook name)
- No broken workbook defined names

The workbook migration marker is schema 10, equal to `WorkbookMigrationManager.CurrentSchemaVersion`. Metadata records migrations 001, 002, and 010 and includes the 14 Development Identified source/mirror signatures. The required very-hidden metadata, template, and A-N mirror sheets are present.

## Existing calculation errors

There are 467 stored error cells, all `#N/A`:

- `OW - Live Covenant Calculation`: 240
- `OW - Covenant Calculation`: 131
- `OW - Charts Source Data`: 95
- `Hidden - Tenure Totals Start`: 1

The first 466 are concentrated in covenant/chart-series ranges where `#N/A` is commonly used to suppress unavailable chart points. The remaining cell is an `INDEX/MATCH` against `ImportedSchemes` and should be reviewed with a representative populated model before being classified as expected.

## Application architecture relevant to this workbook

1. `Application.vb` starts `FormMainScreen`.
2. `FormMainScreen.OpenModelProceedureBP` selects an `.xlsb`, `.abp`, or `.adsa` file and calls `FileManager.OpenModel`.
3. `FileManager.ExcelModel` owns the DevExpress `SpreadsheetControl` document plus structure, data, calculation, presentation, event, transactional-DB, migration, and interface services.
4. After loading, Summit installs custom calculation services and applies/reconciles workbook migrations in both Debug and Release full-model loads.
5. The initial workbook check requires `Global Assumptions!A8`; the UI contract is then deserialized from the deployed `Structure.xml`.
6. The application reads company/start-date values, builds the model interface, switches calculation to manual/chain-based operation, and calculates through the custom engine.

## Follow-up risks

The 18 August 2026 application code pass confirmed that four earlier concerns are already resolved in the current source:

- `FileManager.SaveFile` uses `SpreadsheetControl.SaveDocument()`, preserving the loaded workbook format.
- `StructureManager.CreateStructureFromXML` accepts inline XML, an explicit file path, or the deployed `Structure.xml` default.
- `Abovo_Model_Def.DefID` and `FileID` map directly to their corresponding XML elements.
- `FileManager.OpenModel` validates the workbook contract immediately after loading and before post-load service initialization.

Remaining risks:

1. A worksheet-protection value is stored in source and `Structure.xml`; worksheet protection must not be treated as a security boundary.
2. `ApplicationConfiguration.DefaultTemplateFile` points to `Templates\DefaultBPTemplate.xlsb`, but that template is not present in the repository.
3. Pending workbook migrations now run in both Debug and Release full-model loads. The baseline is already schema 10; migration errors fail the open rather than leaving a partially reconciled production model.
4. The project has no automated test project, so workbook compatibility currently depends on manual/integration validation.
5. VBA remains an operational compatibility layer. Its developer-only `CodeManagement` routines can rewrite VBProjects/files and should remain outside ordinary end-user runtime paths.

The full workbook/VBA/VB.NET scope audit and source-free module/file index are in `Abovo_Summit_Project_Scope_Audit.md` and `Abovo_Summit_Project_Index.json`.
