# Business Plan upgrade and migration handover

**Prepared:** 19 September 2026  
**Repository:** `C:\Repos\Abovo Summit`  
**Branch:** `main`  
**Implementation baseline:** `b8b231c Add business plan comparison and upgrade workflow`  
**Status:** Current two-file comparison and population work is implemented and build-validated. The broader lineage-aware, bespoke-client upgrade architecture described below is proposed work and has not yet been implemented.

## 1. Purpose

This document transfers the business context, verified implementation state, workbook constraints, architectural conclusions, risks and recommended next work for Summit's Business Plan comparison and upgrade capability.

The feature exists because Abovo supplies Excel Business Plan models to client organisations. The models assess current and future financial viability and support statutory reporting. Abovo releases new model versions continually in response to product changes, defects and regulatory changes, particularly changes to the government FFR template. Client organisations populate those models and may also commission bespoke alterations to a generic Abovo base model.

An upgrade therefore has to preserve three different things:

1. the latest generic Abovo logic and statutory requirements;
2. the client's populated assumptions and activity plans;
3. legitimate client-specific changes made to the earlier base model.

The result must remain usable in Summit, in standalone Microsoft Excel and by the embedded VBA project. The client reports that upgrades currently consume about 30% of its working year.

## 2. Core conclusion

This is product-version and client-variant migration for a regulated financial application. It is not adequately described as copying spreadsheet cells.

The target design is a lineage-aware, versioned, auditable migration system with a three-way merge for bespoke models:

```text
old generic template ---- generic release changes ----> new generic template
        |
        +---- client population and bespoke changes --> old client workbook

result = clone of new generic template
       + migrated client population
       + approved client-specific changes
       + explicit conflict resolutions
```

For a normal non-bespoke upgrade the old generic template may not be required during execution, but it is still valuable for establishing lineage and interpreting renamed or transformed fields.

## 3. Workbook authority and non-negotiable constraints

The XLSB is the authoritative financial model. Summit is a VB.NET/DevExpress interface and service layer over it.

The following constraints come from `AGENTS.md` and the repository audits:

- Preserve Excel and VBA round-tripping. A model must remain usable in Summit, Excel/VBA and Summit again without losing formulas, names, VBA behaviour or user edits.
- Do not modify source masters during inspection. Read them detached/read-only and close without saving.
- Do not introduce hidden metadata, template or materialisation sheets merely to support Summit.
- Preserve formula-backed Transactional DB state and workbook-defined names.
- Preserve worksheet order. The audited contract model contains 35,920 formulas using 3-D references, so changing sheet order can change results without an obvious formula error.
- Interactive writes must use `ModelChangeManager`/`ChangeManager.ProcessChange` or `ProcessChanges`.
- Structural commands must preserve workbook-first behaviour, protection, dirty state, recovery behaviour and the established structure services.
- Dependency-sensitive calculations may temporarily use recursive calculation but must restore the previous calculation engine.
- Workbook-facing changes require Debug and Release builds plus manual Summit/Excel/VBA round-trip validation.
- Do not store workbook passwords or extracted VBA source in the repository or migration manifest.

The current authoritative repository masters remain those identified by `AGENTS.md` and `Library/Abovo_Summit_Project_Scope_Audit.md`. Migration testing must use copies rather than mutate those masters.

## 4. Current implementation

### 4.1 Entry point and user interface

`Interface/User Interface/BusinessPlanComparisonForm.vb` implements one form with two tabs:

- **Compare** selects one base workbook and one or more comparison workbooks.
- **Upgrade** selects an older populated source and an open Summit target, assesses compatibility, migrates assumption inputs, calculates, compares outputs and saves an upgraded XLSB plus an XLSX report.

The form is opened from the file-instance `CompareBPs` action in `FormMainScreen.vb`. The previous placeholder message has been replaced by a persistent `BusinessPlanComparisonForm` instance.

Open Summit models are listed alongside detached external XLSB files. External workbooks are loaded into separate DevExpress `Workbook` objects with manual calculation and are disposed when the form closes. They are labelled read-only sources and are never saved by the comparison form.

Both result controls are read-only DevExpress `TreeList` controls with cell multiselect and clipboard copying. Reports can be exported to XLSX.

### 4.2 Comparison service

`Services/BusinessPlanComparisonService.vb` currently compares:

- assumption input values;
- aggregated detailed SOCI values derived from `Transactional_Records`;
- `Outputs_CheckSheet` statuses.

Comparison detail is capped at 25,000 items. Numeric equality uses a tolerance of `0.005`.

The SOCI aggregation expects the canonical Transactional DB fields including `UseInSOCI`, `OrderedSOCIGroup`, `OrderedSOCIHeading` and period columns. Rows with `UseInSOCI > 0` are aggregated by hierarchy and period.

The Check Sheet comparison expects `Outputs_CheckSheet`, currently `Check Sheet!A9:H63` in the contract workbook. A blank status or `OK` is treated as passing; other statuses are reported.

### 4.3 Migration range discovery

The implemented population path discovers migration candidates from the **target model's Assumptions group** in its parsed structure definition. It traverses interface sections and data sources and considers:

- `RowExpandByNR`;
- named-range source `NRName`, `DefinedBy` and `ExpandsBy`;
- cell-range source `DataRange`, row/column defining ranges, extension ranges and offsets;
- data-field `RepeatingNR` and `EditRepNRHere` references;
- XML data-format declarations;
- the data source's `StructureRuleID`.

Candidates are deduplicated by named-range key. Reference counts and data formats are merged. A resolved target range is retained only when it contains at least one unlocked, non-formula input cell.

This is already broader and safer than selecting every defined name beginning with `Rep`. The `Rep` prefix remains a useful discovery convention, but it is not sufficient authority because:

- some writable ranges may not use the prefix;
- the same range can be referenced many times in XML;
- some named ranges are calculated, lookup-only or read-only;
- structure rules, types and editability are expressed elsewhere in XML and workbook protection/formulas.

### 4.4 Structural preflight

The preflight resolves each target migration range against source and target workbooks. It records missing ranges and incompatible geometry. Where a structural rule is associated with a range, it compares source and target record counts and schedules `WorkbookStructureRuleManager.AddRecords` when the source contains more records.

The implementation does not automatically delete target records when the target is larger. It also treats non-empty incompatible target geometry cautiously.

After structural expansion, the service rediscovers the target ranges before copying values because the structural operation can change their addresses and geometry.

### 4.5 Population

For each resolved range the service:

1. copies only the common source/target geometry;
2. requires the target cell to be unlocked and non-formula;
3. skips source formula cells;
4. deduplicates target worksheet/address pairs;
5. skips unchanged values;
6. converts using the XML-derived data type where possible;
7. preserves and copies the source value even when type conversion reports a problem;
8. creates `DataChangeEvent` records and applies them through `ChangeManager.ProcessChanges`.

If a batch write fails, `ApplyChangesBestEffort` recursively divides the batch to isolate individual failing cells. Errors are retained in the migration report. If an unrecoverable exception occurs after structural mutation begins, the model is marked as requiring recovery/Save As.

### 4.6 Calculation, validation and saving

The detached source is calculated using the dependency-sensitive/recursive path. The open Summit target uses `CalculateDependencySensitiveFile(..., Force:=True)`. The service then compares the source and upgraded target assumptions, detailed SOCI and Check Sheet.

The UI requires the user to choose a new XLSB path. `ExcelModel.SaveFileAsTo` refuses the original path when `RequireDifferentPath` is true. The TreeList report is exported beside the result as `<result filename>_report.xlsx`.

Important limitation: the currently open target workbook is mutated in memory before Save As. The target file on disk is protected by the different-path check, but the strict future contract should clone the latest template first, open that clone as the editable result and never mutate the template instance at all.

### 4.7 Structural rules already available

`Services/WorkbookStructureRuleManager.vb` contains explicit rules for:

- `OFA_RECORDS`
- `CAPEX_RECORDS`
- `CAPGRANT_RECORDS`
- `REPAIRS_RECORDS`
- `HOUSING_COMPONENT_RECORDS`
- `FUNDING_RECORDS`
- `DEVELOPMENT_IDENTIFIED_RECORDS`
- `DEVELOPMENT_MULTIYEAR_RECORDS`
- `JOURNAL_RECORDS`
- `STOCK_CONVERSION_RECORDS`

There are also simple append rules/aliases. The manager owns the established linked-range shifts, formulas and Transactional DB post-actions. Migration code must call these rules rather than independently resize their ranges.

The implementation baseline added a workbook-aware `GetRecordCount` overload so a source workbook can be measured without making it the active Summit model.

## 5. Embedded XML and model identity

`WorkbookModelProfile.ResolveStructureSource` already calls `EmbeddedWorkbookStructureReader.TryRead`. The reader opens the XLSB package read-only, examines `customXml/*.xml`, selects an `Abovo_Model_Def` matching the model type and returns its XML. Summit falls back to the installed `Structure.xml` when no suitable embedded definition is present.

This is the correct foundation for future self-describing Summit-compatible workbooks. The current XML contains structural and interface metadata, including model type, definition/file identifiers, group structures, defining ranges, Stress Test and FFR definitions. It does not yet provide enough lineage and migration metadata for safe automated upgrades.

Recommended additions for future XML/schema versions are:

- stable `ModelFamilyId`;
- immutable `ReleaseId` and `ParentReleaseId`;
- `StructureSchemaVersion`;
- FFR/reporting schema version and effective date;
- optional `ClientVariantId`;
- compatible source releases;
- explicit migratable/read-only/calculated classifications;
- semantic field IDs that survive Excel named-range renames;
- structural rule IDs and expansion-master relationships;
- transformation/migration recipe IDs;
- capabilities and validation requirements.

The filename remains a useful human clue and fallback because it generally identifies the base model. It must not be authoritative: filenames can be edited, duplicated or inconsistently versioned.

## 6. Newly clarified master-range convention

The client frequently uses one **master named range** as the controlling structural range. Expanding that master with complete formula contiguity causes all related ranges to be populated or resized correctly. This convention grew organically and may not be consistently declared.

This is a major architectural clue. It means many apparent range-to-range geometry differences may be consequences of one structural action rather than separate migration operations.

It must nevertheless be discovered and encoded safely. Do not infer that every `Rep*` range is a master or expand every related range independently. That risks double insertion, misaligned formulas and Transactional DB corruption.

Recommended representation:

```text
Expansion group
  Stable group ID
  Master named range
  Record-count range or calculation
  Structural rule ID
  Dependent named ranges
  Formula-contiguity requirements
  Expected relative geometry
  Transactional DB post-actions
  Minimum/maximum records
  Validation rules
```

The discovery process should analyse older and current workbooks for:

- ranges that grow together;
- formula patterns immediately above, within and below the expansion area;
- identical row-count drivers;
- existing XML `RowExpandByNR`, `ExpandsBy`, `RepeatingNR` and rule metadata;
- named ranges moved by the same structural command;
- dependencies on sheet order, merged cells, tables, validation, objects and VBA;
- resulting Transactional DB and Check Sheet behaviour.

Once confirmed, the relationship belongs in embedded XML or a versioned migration recipe. It should not remain a heuristic used on every production upgrade.

## 7. Required future upgrade workflows

### 7.1 Standard upgrade

Inputs:

1. old populated client workbook, read-only;
2. latest generic template, read-only;
3. editable result cloned from the latest template.

Process:

1. identify family and releases;
2. construct and review a migration plan;
3. clone the latest template;
4. perform structural expansions through declared master rules;
5. migrate typed inputs;
6. apply explicit version transformations;
7. calculate and validate;
8. save result, human report and machine-readable manifest.

### 7.2 Bespoke three-way upgrade

Inputs:

1. old generic template, read-only;
2. old populated/bespoke client workbook, read-only;
3. latest generic template, read-only;
4. editable result cloned from the latest template.

Comparison rules:

- changed only in the client branch: candidate client customisation;
- changed only in the latest generic branch: retain latest generic behaviour;
- changed differently in both: explicit conflict;
- value difference in a declared writable input: migrate as client data rather than classifying it as a customisation;
- formula, name, structure, formatting, protection, validation, custom XML or VBA/custom-UI difference: inspect as a possible bespoke model change.

Approved bespoke changes should become named, versioned client patch manifests or code-backed migration recipes. Re-discovering the same bespoke changes by workbook diff every year will preserve much of the current upgrade cost.

## 8. Proposed staged engine

The future engine should separate assessment from mutation:

1. **Catalogue**: establish model family, release, parentage, variant and file hashes.
2. **Inventory**: build semantic maps of inputs, expansion groups, formulas, names, sheets, protection, validation, XML, FFR and output/check definitions.
3. **Diff**: compare old generic/client and old/new generic models.
4. **Classify**: distinguish client data, generic changes, bespoke changes, conflicts and unsupported differences.
5. **Plan**: create an immutable ordered plan of clone, structural, transform, write, calculate and validation operations.
6. **Review**: show the plan and conflicts before mutation.
7. **Clone**: create/open a new result from the latest template.
8. **Apply**: run idempotent version migrations, client patches and data population.
9. **Calculate**: run Summit calculation and required dependency-sensitive passes.
10. **Validate**: Check Sheet, detailed SOCI, FFR, named ranges, formulas, sheet-order manifest and structural invariants.
11. **Round-trip**: save, inspect in Excel/VBA on a disposable output and reopen in Summit.
12. **Report**: save result, human-readable report and machine-readable manifest.

Migration recipes should be sequential, for example release A to B and then B to C, rather than one large set of pair-specific rules for every possible source and target combination.

## 9. Reporting and severity model

The current implementation stores geometry notices, type conversion problems and hard failures in one `Errors` collection. The next design should separate:

- **Information**: copied values, unchanged ranges and summaries;
- **Notice**: benign geometry change, planned expansion, ordinary cell difference;
- **Warning**: ambiguous mapping, partial common-area copy or unresolved bespoke classification;
- **Error**: required range missing, invalid structural operation, failed write, failed calculation, non-OK Check Sheet or failed mandatory validation;
- **Conflict**: old client and new generic model changed the same semantic element differently.

The interface should combine a high-level tree with a read-only DevExpress SpreadsheetControl report. Suggested report sheets are Summary, Errors, Notices, Inputs, Structure, Customisations, Conflicts, SOCI, FFR, Checks and Provenance.

## 10. FFR treatment

FFR is a versioned regulatory subsystem, not merely another set of assumption cells. An upgrade can change:

- government-defined layout and labels;
- input and output definitions;
- formulas and validations;
- report generation/export behaviour;
- effective dates and compatibility requirements.

FFR migration needs its own schema version, migration recipes and acceptance tests. Where a statutory version is incompatible, the engine should report an explicit blocking error rather than copy by coincidentally matching addresses.

## 11. SharePoint

All historical and current client models are reportedly stored in SharePoint. No SharePoint connection has been configured or used in this work.

The least disruptive first step is a locally synchronised SharePoint library through OneDrive. Summit can then inventory ordinary paths and continue using detached read-only workbook loading.

A later production integration should use Microsoft Graph/SharePoint with least-privilege, read-only access initially. It should inventory metadata and version history before downloading selected files. Client solvency and activity-plan workbooks contain sensitive financial information, so access should be scoped by site/library/client and audited. Do not bulk-open all client workbooks by default.

## 12. Key risks

1. **Implicit structure**: a relationship may exist only in formula contiguity, VBA or historical practice.
2. **False bespoke classification**: without the old generic template, ordinary old generic logic can be mistaken for a client customisation.
3. **Address-based matching**: renamed/split/merged fields need semantic IDs and explicit transforms.
4. **Double expansion**: independently expanding a master and its dependent ranges can corrupt geometry.
5. **Formula/VBA incompatibility**: value migration can appear successful while Excel/VBA behaviour changes.
6. **Sheet-order damage**: 3-D formulas make worksheet order part of the model contract.
7. **Partial mutation**: current recovery marking does not make a series of structural operations globally atomic.
8. **Regulatory drift**: FFR changes may require replacement logic rather than copied data.
9. **Calculated-output equivalence**: some expected output differences arise from intentional new generic logic; source-versus-result equality is not always the correct acceptance test.
10. **Sensitive data**: SharePoint discovery and reports must not expose client workbooks beyond authorised scope.

## 13. Decisions that should precede major implementation

The next architecture review should answer:

1. What embedded identity can be added to all future models without harming Excel/VBA compatibility?
2. Can Abovo provide the exact old generic template for each populated/bespoke client workbook?
3. Which master expansion groups are already fully represented by structure rules, and which remain implicit?
4. Which types of bespoke change are supported automatically: values, formulas, formatting, validation, names, sheets, VBA/custom UI?
5. Which changes must always be implemented as reviewed client patches?
6. What constitutes an acceptable financial comparison when the new generic release intentionally changes calculations?
7. What FFR versions and upgrade paths must be supported?
8. Should migration create a fully detached result before assessment, or only immediately before execution?
9. What information may be retained in migration manifests and reports for client support?

## 14. Recommended first Astra work package

Do not begin by generalising the current copier. First perform a read-only architecture and evidence pass:

1. read this handover, the test matrix, `AGENTS.md`, the scope/index audits and contract audit/evidence;
2. inspect the implementation baseline `b8b231c` and current source files;
3. map all current structural rules to the XML ranges they control;
4. select representative old generic, old populated/bespoke and new generic workbooks supplied by the user;
5. inspect exact copies read-only with macros/events disabled;
6. identify master/dependent expansion groups and validate formula contiguity;
7. propose the embedded identity/schema extension and migration-plan object model;
8. propose a three-way diff classification and conflict policy;
9. identify the smallest safe implementation slice;
10. return a reviewable architecture proposal before changing production migration code.

## 15. Source map

Primary implementation:

- `Interface/User Interface/BusinessPlanComparisonForm.vb`
- `Services/BusinessPlanComparisonService.vb`
- `Services/WorkbookStructureRuleManager.vb`
- `Services/FileManager.vb`
- `Interface/User Interface/FormMainScreen.vb`
- `Services/WorkbookModelProfiles.vb`
- `Services/StructureCreation/StructureManager.vb`
- `Services/DataService/ChangeManagerV2.vb`
- `Services/Calculation Engine and Custom Functions/CalcEngine.vb`
- `Services/TransactionalDB/TransactionalDBSynchroniser.vb`

Required architectural/audit references:

- `AGENTS.md`
- `Library/Abovo_Summit_Project_Scope_Audit.md`
- `Library/Abovo_Summit_Project_Index.json`
- `Library/Contract_XLSB_Audit_2026-08-24.md`
- `Library/Contract_XLSB_Audit_Evidence_2026-08-24.json`
- `Structure.xml`

Related historical workbook examples mentioned by the user:

- `D:/Downloads/Stori BP v25_0704 2026-27 V4 update 100826 CLEAN pass 3.xlsb`
- `D:/Downloads/Group budget business plan v3.1 (2.7% Pay + Future projects) v25_0104.xlsb`
- `D:/Downloads/Aspire Business Plan 2025-26 vFinal v25_0702.xlsb`

These external paths are examples, not repository-controlled fixtures. Confirm that they still exist before relying on them.

## 16. Validation already completed for the implementation baseline

On 19 September 2026, immediately before commit `b8b231c`:

- Debug solution build: succeeded with zero warnings and zero errors.
- Release solution build: succeeded with zero warnings and zero errors.
- No automated workbook/UI integration suite exists.
- The existing runtime trial and generated report were user-accepted as a good initial implementation, but this is not a substitute for the matrix in `Upgrade_Migration_Test_Matrix_2026-09-19.md`.

