# TestFileMigrated.xlsb analysis

Analysis date: 2026-08-17

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
- Approximately 579,605 formula cells
- No external workbook links
- No data connections
- Embedded `vbaProject.bin` present (1,335,296 bytes)
- Excel reports a VBA project, but its components were not exposed through read-only automation; macro code was not audited or executed

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

- All 186 distinct worksheets referenced by the XML exist in the workbook
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
4. After loading, Summit installs custom calculation services and, in Debug builds, applies/reconciles workbook migrations.
5. The initial workbook check requires `Global Assumptions!A8`; the UI contract is then deserialized from the deployed `Structure.xml`.
6. The application reads company/start-date values, builds the model interface, switches calculation to manual/chain-based operation, and calculates through the custom engine.

## Follow-up risks

1. `FileManager.SaveFile` explicitly saves with `DocumentFormat.Xlsm` even though the main model input is `.xlsb`. The normal `SpreadsheetControl.SaveDocument()` path should preserve the loaded format, but the explicit method is a format-integrity risk if invoked.
2. `StructureManager.CreateStructureFromXML` ignores its `StringXML` argument and always reads `Structure.xml` from the application directory.
3. `Abovo_Model_Def` maps XML `DefID` into `FileID` and XML `FileID` into `DefID`; the field names are reversed in code.
4. Workbook service initialization and Debug migration occur before `ValidateOpenFile`; malformed/non-Abovo workbooks can fail before the intended validation message.
5. A worksheet-protection value is hard-coded in source and `Structure.xml`. It must not be treated as a security boundary.
6. `ApplicationConfiguration.DefaultTemplateFile` points to `Templates\DefaultBPTemplate.xlsb`, but that template is not present in the repository.
7. The project has no automated test project, so workbook compatibility currently depends on manual/integration validation.
8. The embedded VBA project was not readable through automation and therefore remains outside this review.

