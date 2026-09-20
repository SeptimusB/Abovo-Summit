# Business Plan upgrade and migration: evidence review and proposed architecture

Date: 19 September 2026  
Reviewed checkout: `C:\Repos\Abovo Summit`, `main`, `68b96f4`  
Implementation baseline: `b8b231c`  
Status: proposal for review; no production code, XML or XLSB changes made in this review.

## 1. Recommendation

Keep the existing Compare/Upgrade interface and reuse Summit's workbook services. Do not expand the current range copier into an automatic bespoke-model merger yet. First introduce a single, immutable, reviewable migration plan which executes against a **new result cloned from a verified template**, never against an input workbook.

The key unit is a **logical structural family**, not a named range. One master operation can grow many input, formula and Transactional DB ranges. Named ranges locate fields within that family; they are not independent instructions to insert rows or columns.

For bespoke upgrades, first distinguish ordinary population/expansion from customisation. A client workbook with more records than its original template will naturally have different formulas, names and geometry. Those expected differences must be accounted for before identifying bespoke changes. Otherwise we risk treating normal fill-down formulas as custom code and carrying old generic calculations into the new release.

The first deliverable should be a safe, auditable two-file population pipeline, plus read-only three-way classification. Automatic application of bespoke formula, VBA or regulatory changes should follow only after a representative three-way pilot.

## 2. Evidence and limits

### Verified in this review

- Read the handover, prompt, test matrix and index; the project scope audit/index; the contract XLSB audit/evidence; and repository working rules.
- Inspected comparison/upgrade discovery, copying, type conversion, calculation, reporting and Save As; the form and main-screen integration; model-profile/embedded-XML selection; structure definitions and all rule registrations; batch change-manager rollback; relevant interactive expansion and Transactional DB paths.
- Working tree was clean at the start and again following the reported system crash. Only this review document is being added.
- The application has a Compare/Upgrade form with open-model and external-file selection, tree reports, population and XLSX report export. It does not yet implement a three-way bespoke merge.
- `Structure.xml` Assumptions contains 174 datasources, 602 `CellRangeDataSource` elements, no `NamedRangeSource` elements, nine explicit datasource `StructureRuleID` declarations, and 46 repeated-header expansion declarations representing 40 distinct name/method pairs. These are declaration counts, **not** a count of unsupported interfaces or proven migration coverage.
- `WorkbookStructureRuleManager` registers ten specialist families and 36 simple families.
- The three supplied historical client files, repository Blank/Demo masters, and the Sandbox New Blank all exist. This is filesystem verification, not a fresh inspection of their internal workbook contracts.

### Not validated in this review

No XLSB was opened, recalculated, changed or saved. No VBA, Excel round-trip, populated-model migration or GUI acceptance test was run. Earlier workbook findings in the audits remain historical evidence; they are not current runtime proof for the three supplied client files. There is no automated workbook/UI integration suite in the repository.

No build was needed for this documentation-only review. Prior build results remain recorded in the handover, not rerun here.

### Corrections/clarifications to durable context

- Actual debug auto-open is `C:\Sandbox\BP v26_0001 - New Blank.xlsb` in `FormMainScreen.vb:24`. The repository working rules still describe the Demo master as the debug default. Neither setting nor either master has been changed here.
- The current calculation entry point is in `Services/EngineManagement.vb:445`, not the `.../CalcEngine.vb` path listed in the handover source map.
- A forced open-model dependency calculation currently uses `ChainBased` plus `CalculateFullRebuild`, temporarily disables deferred TDB skipping, and restores previous settings. External comparison workbooks use `Recursive` plus `CalculateFull`. Earlier notes describing all dependency-sensitive calculation as Recursive are not the present implementation.
- Save As does protect against reusing the target's current path, but that does not make the target immutable during population. It is the open target object that is mutated.
- Existing structure operations already preserve important entry state and mark recovery on partial failure. The gap is migration orchestration continuing after those failures, not the absence of all lower-level safeguards.

## 3. Current implementation: strengths and material gaps

### Preserve these foundations

- Values pass through `ChangeManager.ProcessChanges`, with typed writes, before/after snapshots, grouped history, dirty state and rollback verification.
- Structural growth uses established semantic rules rather than copying old worksheets wholesale.
- Structure rules restore visibility/protection and calculation mode, process linked sheets in a deliberate order, then synchronise Transactional DB and invalidate dependent interfaces.
- The current population path avoids overwriting formula/locked target cells and deduplicates both range references and target cell addresses.
- Individual type problems and financial differences are exposed in a tree and saved report. This is a useful basis for a reviewer workflow.

### Findings requiring attention before general client upgrades

| Priority | Verified behaviour and evidence | Consequence / proposed correction |
|---|---|---|
| High | `CellValuesEqual` uses `NumericTolerance = 0.005` for all numeric inputs; population skips equal cells (`BusinessPlanComparisonService.vb:275,1138`). | A source rate of 1.40% and target rate of 1.00% differ by 0.004 and can be skipped. Compare input values exactly by declared type. Define tolerances separately for each financial metric. |
| High | Population writes into the selected open target; Save As follows execution (`BusinessPlanComparisonForm.vb:346,401,420`). | The template's in-memory state is not read-only, including on failure. Clone first into an isolated result session. |
| High | Destination checking excludes only the target path (`BusinessPlanComparisonForm.vb:369`; `FileManager.vb:567`). | A source/old-template/other participating file path is not excluded by this protection. Reject all input identities and paths, including aliases, before any save. |
| High | Assessment enumerates writable names on selected assumption worksheets; execution discovers target-XML ranges (`BusinessPlanComparisonService.vb:149,644,722,1109`). | The preview is not the plan executed. Both buttons must use the same frozen plan and coverage manifest. |
| High | Structural failure is logged and the next expansion/copy proceeds (`BusinessPlanComparisonService.vb:218`). Best-effort value batches are recursively retried without checking recovery state (`:345`). | Lower services can already mark `RecoveryRequired`; proceeding can operate on untrusted geometry or a failed rollback. Stop the affected run on structural/recovery failure. Retain an explicitly incomplete diagnostic result only. |
| High | Copying uses the common row/column rectangle and range-relative offsets (`BusinessPlanComparisonService.vb:256,268`). Counts for both files use target rule definitions (`:680`). | Same name does not prove same layout/meaning across releases. Resolve source and target through separate versioned contracts and semantic field mappings. |
| High | Unresolved target names are silently omitted (`BusinessPlanComparisonService.vb:805`). Scope/area identity is not part of range keys (`:824,1160`). | A missing required input can disappear from coverage, and local or multi-area names have no explicit migration semantics. Every declared input needs a discovered, excluded, unsupported or failed disposition. |
| High | SOCI ignores nonnumeric values, including error values; accepts no recognised periods; missing dictionary entries become zero (`BusinessPlanComparisonService.vb:519,602`). | Incomplete/failed calculation can look like agreement. Validate coverage and error cells before numerical comparison; missing is not zero. |
| High | A SOCI preparation problem is counted as a comparison structural issue, but migration carries only detailed numeric-difference and check counts (`BusinessPlanComparisonService.vb:316,507`). `Success=True` is set when execution reaches its end (`:331`). | A saved/completed run is not necessarily financially verified. Separate execution, coverage, validation and reviewer acceptance states. Propagate every validation failure. |
| Medium | `RO`/calculated fields suppress type metadata but there is no corresponding field-level cell eligibility mask during copying (`BusinessPlanComparisonService.vb:272,833,872`). | An unlocked cell inside a larger discovered range can still be copied despite a read-only field definition. Compile field-level eligibility and address maps explicitly. |
| Medium | `EditRepNRHere` is passed as a candidate named range (`BusinessPlanComparisonService.vb:792`), although XML uses `TRUE`; repeated-header expansion methods are not compiled into upgrade operations. | Runtime UI metadata and migration semantics are being conflated. Parse booleans as booleans, and map header masters to approved structural capabilities. |
| Medium | Type orientation is inferred by matching field count to range row/column counts, row count first (`BusinessPlanComparisonService.vb:855`). | Pivoted, square, repeated or multi-area ranges can be ambiguous. Use explicit field addresses/axes from the resolved contract, not dimensional coincidence. Conflicting types on overlaps must be reported. |
| Medium | Direct Compare reads current values without an explicit calculation step; upgrade calculates source and target through different adapters (`BusinessPlanComparisonForm.vb:296`; service `:1060`). | State freshness and calculation comparability are not established. Record calculation provenance and use tested per-release calculation profiles. |
| Medium | Check Sheet treats blank status cells as ignorable; comparison uses display labels/ordered headings and prints differences as `N0` (`BusinessPlanComparisonService.vb:549,567`). | Uncalculated checks or small real differences can be obscured. Require expected check coverage and show precision appropriate to units and tolerances. |
| Medium | Progress pumps `Application.DoEvents` and only the populate button is explicitly disabled; result save precedes report export (`BusinessPlanComparisonForm.vb:392,407,420`). | Reentrancy and partial publication need explicit handling. Use an operation guard and staged output publication with a durable run manifest. |

Additional limits: Compare considers base-side names, so target-only assumption additions need explicit coverage; the worksheet set is a union of open-model assumptions, not a separate source contract. Source formula-backed inputs are skipped during population and need a visible policy/disposition. Display-text comparison is insufficient for raw string/number/date distinctions. External source workbooks are loaded with `New Workbook`; no release-specific compatibility setup is visible on that path. Custom functions are also registered through `GlobalCustomFunctions` elsewhere, so this is **not proof that external workbooks lack UDFs**: the required function set, settings and outputs must be tested, not assumed missing or present.

## 4. Master/dependent structure: what is already known

### Specialist rule inventory

Source: `Services/WorkbookStructureRuleManager.vb:293-645,1020`.

| Family / rule | Count/master and boundary | Coordinated effect |
|---|---|---|
| Other Fixed Assets / `OFA_RECORDS` | `Rep_OFA_010`, count adjustment -1; insertion `LastOFACol` | Assumptions plus OFA Workings. Formula/format template and visible column-width donor differ. |
| Capital Expenditure / `CAPEX_RECORDS` | `Rep_CapExpend_010`, -1; `LastCapExpendCol` with offset -1 | Assumptions plus OFA Additions; insert inside the boundary before the hidden template. |
| Capital Grant / `CAPGRANT_RECORDS` | `CapGrantInclusion`, -1; `LastCapGrantCol` with offset -1 | Assumptions plus workings; dependency/name expansion must precede TDB growth. |
| Repairs / `REPAIRS_RECORDS` | `StockCondCats`; `LastStockCol` | Eight sheets: assumptions, condition inputs/results, rates, drivers, costs, depreciation and replacement costs. XML's explicit rule bridges its different `RowExpandByNR` (`RepIncStkCat`). |
| Housing components / `HOUSING_COMPONENT_RECORDS` | `DepnType`, -1; `LastCompCol` with offset -1 | Housing assumptions, stock numbers, Components 1-12 and totals; copy from shifted template column. |
| Funding facilities / `FUNDING_RECORDS` | `FacilityNames`; `LoanDescRev1` | 32 funding/interest/balance/fee sheets. This is different from expanding the separate Funders/Facility lookup rows. |
| Identified development / `DEVELOPMENT_IDENTIFIED_RECORDS` | `HouseTypeInID`, -1; `LastIDColNum` | Seven ordered sheets; `Dvpt Component Depn` before `Dvpt NonCash`. |
| Multi-year development / `DEVELOPMENT_MULTIYEAR_RECORDS` | `HouseTypeInMY`, -1; `LastMYColNum` | Same seven-sheet family, separate logical records/anchor. |
| Journal / `JOURNAL_RECORDS` | `Rep_Jour_01`, -1; insert at its end before sentinel | Whole Journal Assumptions rows, including hidden debit/credit formulas; TDB source `IR_Journals`. |
| Stock conversion / `STOCK_CONVERSION_RECORDS` | `IR_StockDispAss_01`; append after range | Whole assumption rows, retaining formulas to the right. |

The funding rule is present in code but is not one of the nine explicit datasource rule declarations counted above. Explicit declaration count alone therefore does not determine rule availability.

### The 36 simple rule registrations

These append after a named range, copy from the preceding row/column, then synchronise by source worksheet. Their existence is verified; their correctness for every historical release is not.

| Worksheet | Master names registered | Axis |
|---|---|---|
| Rent Assumptions | `Rep_Rent_02` | Rows |
| Service Charge Assumptions | `Rep_ServChg_01`, `Rep_ServChg_02` | Rows |
| Stock Disposal Assumptions | `IR_Stock_Disp_Desc` | Rows |
| Owner Occupier Assumptions | `IR_Own_Occ_Desc` | Rows |
| Specific Income Assumptions | `SummaryOtherIncCat`, `IR_Spec_Inc_Ass1` | Columns |
| Capital Grant Assumptions | `CapitalGrantCats` | Columns |
| Intercompany Income Assumptions | `IR_Intco_Inc_Ass1`, `IR_Intco_Inc_Ass2` | Columns |
| Management Costs Assumptions | `IR_Oneoff_Cost_Ass` | Columns |
| Repairs & Maint. Rephasing | `IR_Rep_Rephase` | Columns |
| Additional Pension Assumptions | `IR_AddPension_01`, `IR_AddPension_02`, `Rep_AddPension_050` | Rows |
| Development BP Assumptions | `IR_Dvpt_CapInt` | Rows |
| Economic Assumptions | `IR_RPI`, `IR_ServChg`, `IR_RTBReal`, `IR_OtherIncReal`, `IR_CapGrantReal`, `IR_LeaseIncReal`, `IR_StaffCostReal`, `IR_BPRepairReal`, `IR_OtherExpReal`, `IR_DemCostReal`, `IR_DevRentReal`, `IR_DevServiceChg`, `IR_DevMgtReal`, `IR_DevRepairsReal`, `IR_DevOtherReal`, `IR_FundFeesReal` | Rows |
| Capitalisation Assumptions | `IR_Mgt_Cost_Cap`, `IR_Repairs_Cap` | Rows |
| Cash Journals | `IR_Cash_Journals`, `IR_MvtWC` | Rows |

Exact IDs are defined at `WorkbookStructureRuleManager.vb:656-978`; use that registry rather than synthesising IDs from uppercased range names.

### XML and legacy paths to reconcile

- `StructureRuleID`, `StructureAddCommand`, `StructureDeleteCommand`, `RowExpandByNR`, `RowsExpandModel` and `SkipLastRecords` express datasource operations/boundaries.
- `NRDSName`, `DataRange`, row/column definition names, offsets, extra-record and mapped-right flags express addressing/shape.
- `RepeatingNR`, `RepeatsByNR`, `RepeatsByCR`, pre/post counts, `EditRepNRHereDataFormat` and `EditRepNRHereExpansionMethod` express embedded editable period/date/category headers.
- `Pivot`, merge direction and field `Index` describe presentation and must not be mistaken for physical workbook orientation.
- `DataFormat`, `RO`, calculated/dummy flags, bounds and repository references supply type/validation intent. Workbook formulas/protection remain additional authority; neither source can simply override the other silently.
- XML contains `<MasterNR>` for `Rep_Rent_02` and `Funders`, but current `ISEDatasource` does **not** deserialize `MasterNR`. `RowExpandByNR` supplies the working legacy path. Merely adding more `MasterNR` tags will not create a migration contract.
- `Funders`/`Facility` share a whole-row expansion in Funding Assumptions (`Structure.xml:11915`). `Funders` has no registration in the 46-rule registry. DIT falls back to `WorkbookManager.InsertRows` when there is no semantic rule (`DataInterfaceTemplate.vb:12708`). Migration currently does not provide that fallback.
- Repeated headers also use legacy `NRRI`, `NRCI` and `NRRIbyCOL` paths. The last infers physical axis from range dimensions (`DataInterfaceTemplate.vb:12335`). A reviewed migration capability must encode the axis, not repeat this inference across unknown releases.

### Evidence still required from exact workbooks

For each approved release/family, capture scope and ordered areas of all relevant names; logical record count versus physical capacity; sentinel/template exclusions; master insertion boundary; dependent-sheet order; formula templates and relative-reference patterns; hidden/zero-width template geometry; validation and protection; period/category mappings; and TDB source/mirror counts.

Inventory inspection can be read-only. **Proving** expansion contiguity requires a later authorised experiment on disposable copies: add one and five records, inspect every linked formula/name boundary, recalculate, then exercise Excel/VBA and reopen in Summit. Historical audit evidence identifies high-risk boundaries but does not replace these tests for a different release.

## 5. Proposed model identity and metadata

Keep the existing embedded `Abovo_Model_Def` as the UI definition. Add an explicitly versioned migration-contract section, with backward-compatible optional fields and validation. Do not introduce hidden sheets or silently write metadata into old files on open.

Identity fields should distinguish:

- `ModelFamilyId`: stable business-plan family, not the display name.
- `GenericReleaseId` and `ParentGenericReleaseId`: version lineage.
- `ContractSchemaVersion` and `ContractId`: migration contract identity; independent of UI CSIDs/GSIDs.
- `ClientVariantId` and ordered `ApprovedPatchIds`: optional bespoke lineage.
- `FfrTemplateId`, `FfrRelease`, reporting period and mapping-contract version.
- `CalculationProfileId`: engine policy, required UDFs, date system and supported external-link handling.

Keep exact file hashes, application/library versions, input paths, extraction timestamp and run IDs in an external provenance manifest. Avoid a self-referential whole-file hash embedded within the file it hashes. For legacy files, use reviewed sidecar contracts tied to known fingerprints; filenames are discovery hints only. Ambiguous identity blocks automatic execution but not read-only inventory.

Proposed declarative entities:

| Entity | Required meaning |
|---|---|
| `StructuralFamily` | Stable ID, master locator, logical record axis/count, sentinel rules, approved capability/version, ordered dependent families/sheets, preconditions and postconditions. |
| `Field` | Stable semantic ID, ownership (`ClientInput`, `GenericFormula`, `Derived`, `ClientCustomisation`), type, units, blank policy, validation/repository identity and physical locator. |
| `Locator` | Workbook/sheet-scoped name or explicit sheet range, ordered multi-area mapping, field and record axes, offsets and shape; resolved afresh after structural operations. |
| `RecordKey` | Stable supplied key where available; otherwise reviewed composite/positional mapping with provenance. Labels alone are not unique identifiers. |
| `PeriodAxis` | Actual period identity, ordinal plan year, calendar/fiscal basis and authorised mapping rule. |
| `Metric` / `Check` | Stable identifier, expected coverage, units, precision, tolerance, evaluation and applicable release/period. |
| `MigrationRecipe` | Source/target contract versions, aliases/transforms, supported capabilities, dependency order and test evidence. |

Example only, not an implemented schema:

```xml
<MigrationContract SchemaVersion="1" ContractId="abovo.bp.v26_0001">
  <Identity ModelFamilyId="abovo.bp" GenericReleaseId="v26_0001" />
  <StructuralFamily Id="capital-grants"
                    CapabilityId="CAPGRANT_RECORDS" CapabilityVersion="1">
    <Master Scope="Workbook" Name="CapGrantInclusion"
            RecordAxis="Columns" TrailingTemplateRecords="1" />
    <Field Id="capital-grant.description" Type="String" Ownership="ClientInput">
      <Locator Scope="Workbook" Name="Rep_CapGrant_010" />
    </Field>
  </StructuralFamily>
</MigrationContract>
```

A real field locator needs verified offsets/area mappings; the example deliberately does not invent them. Metadata can select allowlisted built-in capabilities, not arbitrary scripts, VBA or reflection targets. Conflicting, duplicate or unsupported embedded contracts must be reported, not resolved by whichever XML part happens to be found first. Current embedded lookup selects by root/model type (`WorkbookModelProfiles.vb:303`) and needs strengthening for this purpose.

## 6. Immutable plan and execution stages

### Object model

`MigrationPlan` contains: plan/schema version; operation ID; immutable input artifact identities/hashes; source/old-base/new-template contract versions; frozen options and approved decisions; ordered steps and dependencies; input coverage; expected postconditions; financial acceptance policy; and output destinations.

Each `PlanStep` contains: stable step ID; kind; semantic family/field IDs; source locator and expected content/type fingerprint; destination semantic locator; required predecessor steps; capability/version; expected before/after geometry; risk classification; validation rules; and permitted failure policy. Do not store live DevExpress `Cell`, `Range`, model-array slot or UI-control references in the plan.

Keep runtime facts in a separate `MigrationRun`: step attempts, outcomes, timings, resolved coordinates, actual counts, before/after evidence, recovery checkpoints and diagnostics. A reviewer change produces a new plan revision/hash, not a mutation of the approved plan.

### Ordered stages

1. **Acquire and freeze.** Select source, optional exact old generic base, and new generic template. Initially require saved inputs; later support an explicit immutable capture of unsaved open state. Hold original input sessions read-only to the operation. Verify identity, available resources and destination exclusions.
2. **Inventory.** Extract names/scopes/areas, input/formula masks, structural families, calculation dependencies/profile, validation, package-part and regulatory identities. Record unsupported objects as coverage, not silently ignored content.
3. **Classify and plan.** Compile version-to-version mappings; deduplicate families and physical writes; resolve record/period mappings; identify customisations after normalisation; produce conflicts and coverage. No result mutation yet.
4. **Review.** Show operations, omissions, errors/conflicts and expected financial changes. Approve a plan revision. Refuse structurally unsafe plans; permit clearly labelled drafts for mapped-but-invalid data when selected policy allows it.
5. **Clone.** Make an exact copy of the latest template at a new staging location and load it into an isolated result session with existing Summit services. Verify copy identity. Do not repurpose the user's open template model slot. Keep analysers/ordinary interface refreshes detached during the job.
6. **Prepare structure.** Execute each family once in dependency order through approved structure capabilities. Resolve master/dependent names after every growth. Verify geometry, formulas, template boundaries and protection before downstream copying. Keep current TDB post-actions initially; batching them across families is a later tested optimisation.
7. **Apply approved bespoke changes.** Apply only resolved, supported patches to the new generic result. Structural bespoke patches may need to precede family expansion; their ordering belongs in the plan graph. Revalidate mappings afterwards.
8. **Populate.** Apply exact typed input values through the change manager using the compiled eligibility/address map. Shared cells have one write owner. Preserve blank, zero and empty string distinctions according to field policy. Report source formulas in input positions explicitly; do not silently drop them or overwrite latest formulas.
9. **Reconcile and calculate.** Verify TDB mirror structure, invalidate inappropriate result snapshots, apply the supported dependency-sensitive calculation profile, and ensure checks/outputs are current. Calculate isolated source evidence if required, not the user's original workbook session.
10. **Validate.** Compare coverage, data, formulas/structure, Check Sheet, SOCI and relevant FFR mappings/results. Classify differences and obtain sign-off for expected changes.
11. **Publish.** Save staged result, reopen/verify persistence, generate reports and provenance, and publish a new output plus matching `_report.xlsx` and JSON manifest. Never announce accepted status merely because saving worked. A completion manifest marks a complete output set; two separate file writes are not a single atomic transaction.

Ordinary comparison can reuse inventory, normalisation and metric extraction without mutation. It must state whether values are saved caches or freshly calculated evidence.

## 7. Three-way rules for bespoke models

Let **B** be the exact old generic template, **C** the old populated/client-customised model, and **N** the new generic template. Result **R** starts from N. A read-only analysis can use normalised semantic views; a later disposable base expansion may provide stronger evidence of expected generated structure. Neither approach saves changes to B or C.

Separate user data from schema/logic before classifying deltas. Normalise translated formulas and generated ranges relative to records/fields, without erasing genuine constants, sign changes, external links or 3-D sheet-order changes.

For schema/logic elements:

| Comparison after normalisation | Classification / default |
|---|---|
| C equals B; N changed | Generic-only change: retain N. |
| N equals B; C changed | Client-only candidate: apply only if supported and approved. |
| C and N made the same semantic change | Already incorporated: keep N, do not duplicate. |
| C and N differ from B differently | Conflict: explicit resolution, no last-writer-wins. |
| Mapping/base unavailable or comparison incomplete | Unknown/unsupported: cannot establish a safe three-way merge. |

Client input data is migrated using field ownership and mapping, **not** simply the above formula merge test. A changed new-template default does not automatically displace an explicit client input. Conversely, a newly mandatory constraint cannot be waived because the source has a value. Report this as a validation/policy conflict.

| Element | Required treatment |
|---|---|
| Values | Exact typed input migration; distinguish blank/default/explicit value, dates, units and currencies. Validate without silently rounding or reinterpreting source data. |
| Formulas | Keep new generic formulas unless an approved bespoke patch says otherwise. Compare normalised expression and dependencies. Treat source formulas in input positions as explicit cases. No formula-text substitution across arbitrary addresses. |
| Defined names | Match semantic ID plus scope and ordered area definitions. Handle aliases, removed/renamed names, constants/formula names and reference targets; do not collapse local names. |
| Geometry | Model approved master expansion as an operation. Generated record growth is not bespoke. Added/deleted fields, changed record stride or unrelated insertions require mapping/review. No independent resizing of overlapping ranges. |
| Formatting | New template is the default. Transfer approved client-only presentation changes only where identity is stable; retain number-format and unit meaning. Apply no Summit-only recolouring to workbook-backed cells. |
| Validation | Preserve new-release validation unless explicitly resolved. Report changed lists/codes, dependencies and source values outside the new domain. |
| Protection | Preserve latest template rules and entry protection state. A now-locked/now-formula field is not an invitation to unprotect and overwrite it. |
| Sheet order / visibility | Preserve N by default. Any bespoke change affecting 3-D references, hidden templates or macro sheet addressing requires specialised validation. |
| Custom XML | Select approved versioned definitions by identity. Merge semantic contract changes deliberately; do not copy old UI/schema XML wholesale over N. Preserve unrelated package parts. |
| VBA and custom UI | Inventory/hashes and explicit compatibility review. No generic automatic VBA merge or old-project overwrite. Approved bespoke modules/menu changes require separately tested patches and Excel command tests. Do not retain extracted VBA/passwords in repository evidence. |
| FFR | Versioned reporting subsystem with its own template, period, field mapping and validation. Preserve current statutory-release logic; never paste an old FFR block blindly onto a new template. |
| External links / unsupported features | Inventory and obtain an explicit policy. Never refresh uncontrolled external sources during analysis or convert unresolved links to zeros. |

Store approved customisations as reusable, release-qualified patches with tests. A patch should declare source/target contract preconditions and its effect; applying it twice must be detected. Arbitrary VBA/FFR merges remain outside the first implementation slice.

## 8. Transactions, recovery and idempotency

- Inputs are immutable job artifacts. File path equality alone is insufficient; guard input aliases/identities and open-session mutation. Recheck hashes before execution/publication.
- Use an exclusive migration operation guard. Disable competing actions and do not rely on `DoEvents` or an individual disabled button to protect a job.
- Keep mutable DevExpress workbook work on its owning execution context. A progress/splash UI is separate from the calculation engine. Parallelise only isolated, supported work after memory and threading tests; four large simultaneous workbooks are a real memory risk, particularly in x86.
- Value batches continue using change-manager rollback, but migration must inspect transaction outcome, actual committed count and model integrity. Do not recursively retry following failed rollback or structural recovery status. Benign bad values can be copied raw into an explicitly incomplete/draft result and reported where mapping and target editability are certain.
- Structural rollback is not achieved by ordinary cell undo. On partial family failure, stop execution and discard/restart the disposable result from a known checkpoint. Preserve diagnostic evidence, not an apparently valid model.
- Checkpoints are full isolated result artifacts/manifests at deliberate stage boundaries. A checkpoint must itself pass save/reopen checks before it is resumable. Initial implementation should restart from the clean template rather than attempt a complex partial resume.
- Resolve destination cells only after structural dependencies finish. Assert expected source content and target preconditions before writing.
- Make execution logically idempotent: identical frozen inputs/recipe/decisions produce equivalent semantic results; a repeat run starts from a fresh clone. A resume verifies step postconditions instead of adding records again. Do not demand byte-identical XLSB files with volatile values/timestamps.
- For save/report failure, mark the publication incomplete and retain explicit recovery paths. Do not overwrite existing results without separate confirmation; never overwrite any input. Include output hash in the external manifest after save.

## 9. Reports and reviewer interface

Retain the two tabs. Add an optional old-generic-base selector for bespoke analysis. Use a plan tree grouped by model, domain, structural family, field and record, with a problems-only default and a switch to all notifications/coverage.

An issue should open aligned native SpreadsheetControl views of source, old base where applicable, latest template and result. Input panes are read-only, macros not executed, and navigation points to the actual recorded cell/name rather than a rendered imitation. Open these views lazily; do not permanently load four full GUI workbooks just to show one error. Result corrections use the established change manager and invalidate the relevant validation/sign-off.

Every report item should include: stable issue code; severity; stage/step; semantic field/family; artifact identity; sheet/name/scope/area/cell; typed before/source/result values; expected versus actual; reviewer explanation/action; resolution and validation state. Preserve precision and escape HTML/XML content.

| Severity | Meaning / example |
|---|---|
| Information | Routine operation or unchanged verified coverage. |
| Notice | Successful expected resizing, copied-cell counts, approved ordinary geometry changes. Not a client error. |
| Warning | Review advisable; a safe default used or a mapped field needs human confirmation. |
| Error | Missing required input, failed operation/calculation/check, uncovered/truncated validation, invalid mapped value. Can coexist with a saved diagnostic draft, never silently accepted. |
| Conflict | Competing semantic changes or unresolved business decision; explicit resolution required. |

Separate report severity from run status: `Planned`, `Running`, `DraftWithIssues`, `Failed`, `Validated`, `Accepted`, `PublicationIncomplete`. The precise names can change, but saved, validated and accepted must remain distinct.

Outputs: existing `_report.xlsx` with Summary, Problems, Coverage, Operations and Financial Reconciliation sheets; optional self-contained HTML problems report; and a structured JSON manifest/issue list for repeatability. An errors-only XLSX can be a filtered derivative, not the only audit trail. All reports use the same underlying result objects.

Coverage counts must include expected/discovered/mapped/copied/unchanged/excluded/unsupported/failed inputs and expected/evaluated/failed metrics. Pagination/detail limits affect display only, not validation or total counts. Repeated XML references must neither multiply writes nor hide conflicting definitions.

## 10. Financial acceptance

### Structural/data gate

All required fields have a disposition; no unnoticed truncation, misalignment, lost records or unresolved required mappings. Input values reconcile exactly after approved transforms. New-template formulas remain present except approved patches. Names, areas, sheet order, protection and formula continuity meet the release contract. TDB source and mirror geometry reconcile. Snapshot data is derived state, not migrated client input.

### Calculation gate

Record engine/library versions, UDF requirements, calculation settings/date system, calculation completion and external-link policy. Formula errors are errors, not missing numeric values to skip. Check expected period/check coverage explicitly. Where source and target use different engines/profiles, establish parity on the fixture set before relying on their deltas.

### Check Sheet gate

Compare check IDs and expected populations separately for source and result. Distinguish existing source failures, resolved failures, new failures and intentional release changes. All-empty/unavailable checks are unverified, not passed. A pre-existing source problem does not silently excuse the result; documented exceptions require reviewer approval.

### SOCI and supporting metrics

Match stable statement IDs and actual fiscal periods, not ordered display labels alone. Show unrounded values and explicitly declare units (for example pounds versus thousands), absolute tolerance and any justified relative tolerance. Missing rows/periods are coverage events, not implicit zeros.

Same-release population should preserve the same detailed SOCI within an agreed output-only numerical tolerance, and reconcile it to the authoritative detailed statement output. A TDB aggregation is a useful additional cross-check, not automatically proof of full statement equivalence. Also test financial position, cashflow/funding and covenant metrics selected for the pilot.

For new releases, equality is not always the correct target. Report a reconciliation bridge: source result; known generic-release effects; approved bespoke effects; mapping/period effects; and unexplained residual. Attribute effects only where a controlled comparison or approved release specification supports them. Nonlinear interactions may prevent clean additive attribution; leave an explicit reviewed interaction/residual rather than inventing a balancing explanation. The final acceptance gate is no unexplained material difference and all required checks resolved/approved, not merely a small difference count.

### FFR

Validate the designated FFR release, reporting period, mapping completeness and available template checks; reconcile financial fields to the accepted model. Old/new regulatory templates may not have one-to-one cells. Record discontinued/new fields and expected regulatory changes separately. Summit validation is not a substitute for any external submission/acceptance process.

## 11. Smallest safe implementation and pilot

### Slice A: safe plan-first population

1. Introduce immutable artifact/contract/plan/run DTOs and one discovery pipeline, used by Assess and Upgrade.
2. Fix exact input comparison, eligibility masks, unresolved-input coverage, scope/area identification and output-validation status propagation.
3. Clone the template into an isolated result; exclude all input save destinations; add operation/recovery guards and staged result/report publication.
4. Reuse existing approved structural capabilities for known contracts. Promote Funders and one representative repeated-header operation to explicit capabilities after disposable-copy evidence; do not invent a universal fallback.
5. Provide deterministic coverage and problems reports. Implement no automatic bespoke VBA/FFR/formula merge in this slice.

This is intentionally a supported subset with honest exclusions, not an assertion that every old workbook is safely migratable. Retain existing interactive calculation/expansion behaviour outside the migration session.

### Slice B: read-only three-way pilot

Inventory one matched old-base/client/new-template set. Normalise ordinary expansion, classify client-only/generic-only/conflicting changes, and review with Abovo. Add approved version maps and a small explicit patch set. Use findings to refine metadata before generalising.

### Slice C: reviewed bespoke application

Execute supported approved patches with regression fixtures, then populate/recalculate/reconcile. Add more release families and regulatory versions only with evidence. Later consider SharePoint cataloguing; locally synced, identified files are sufficient to start.

### Pilot fixtures

First, a same-release round trip using disposable copies of the repository Demo and Blank masters provides a control without release-change ambiguity. Next use Stori and the confirmed latest template, with unsupported coverage visible. The subsequently supplied `Master BP v25_0700 (1).xlsb` is now a strong provisional baseline for the read-only three-way Stori pilot: its Version History matches Stori's except for four documented additions. Validate the actual structural/logic deltas and ancestry assumptions before claiming a verified bespoke merge; see `Stori_Baseline_Version_History_Review_2026-09-19.md`.

Broaden to Group budget and Aspire to exercise other release generations and client variation; do not infer their generic ancestry from filenames alone.

Available files verified on disk:

- `C:\Repos\Abovo Summit\Library\Blank BP v26_0001.xlsb`
- `C:\Repos\Abovo Summit\Library\Demo BP v26_0001.xlsb`
- `C:\Sandbox\BP v26_0001 - New Blank.xlsb`
- `D:\Downloads\Stori BP v25_0704 2026-27 V4 update 100826 CLEAN pass 3.xlsb`
- `D:\Downloads\Group budget business plan v3.1 (2.7% Pay + Future projects) v25_0104.xlsb`
- `D:\Downloads\Aspire Business Plan 2025-26 vFinal v25_0702.xlsb`

These are input references, not a request to copy confidential client workbooks into Git.

## 12. Test plan additions and acceptance evidence

Use `Upgrade_Migration_Test_Matrix_2026-09-19.md` as the existing broader matrix. Add the following concrete risk-led tests:

| Test | Required evidence |
|---|---|
| Rate precision | 1.00% versus 1.40%, small percentages, small currencies and integer fields: inputs copied/compared exactly; output tolerance does not influence copying. |
| Immutable inputs | Hashes unchanged before/after success, cancel, structural fault, calculation error and save failure; open input values/dirty state unchanged. Destination equal to any input/alias rejected. |
| One plan | Assessment operations and execution step IDs identical for a frozen plan; changed input, XML or recipe invalidates approval. |
| Duplicate/overlapping definitions | Many XML references to one master cause one expansion; duplicate physical cells written once; conflicting types or sources are conflicts, not first-wins. |
| Scope and multi-area names | Global/local duplicates, two local scopes, discontiguous areas, renamed/removed names and formula names are correctly mapped or explicitly unsupported. |
| Field eligibility | XML read-only/calculated/dummy fields inside an otherwise writable range stay excluded; target locked/formula cells are never overwritten. |
| Source formula input | Source formula in a nominal input field is explicitly classified/reported; no silent omission or uncontrolled formula transfer. |
| Master propagation | Funders/Facility, one repeated-date header, Cash Journals, Capital Grant, Repairs and Development: one/five-record growth preserves all dependent names, formulas and template exclusions. |
| Cross-axis mapping | Physical rows presented as columns, mixed field/repeat dimensions, square ranges, multi-area records and blank template capacity cannot be mapped by a dimension guess. |
| Three-way normalisation | Ordinary record growth produces no false bespoke patch; intentional formula constant/operator edit is still detected. Test same change on both branches and true conflicts. |
| Recovery | Inject failures after each linked-sheet insert, TDB sync, batch write and rollback. Migration stops on recovery-required state, never retries on corrupt geometry, and originals remain untouched. |
| Coverage false pass | Missing SOCI period/header, formula error, no included rows, blank checks, unsupported output range and report display truncation cannot yield Validated. |
| Financial bridge | Same-release control reconciles; approved new-release change produces explained delta; unexplained material residual prevents acceptance. |
| Persistence | Save/reopen result and reports, verify counts, formulas, names, XML, package features and hashes; test disk/report failures and clear incomplete-publication status. |
| Excel/VBA round trip | Summit result -> Excel/VBA -> Summit; exercise custom add-in/menu commands, supported calculation, protection, structural add-lines and TDB reconciliation. Retain package-part hashes without extracted VBA. |
| FFR release | Correct template/period/mappings, explicit new/removed fields, template validation and reconciliation; old template data cannot overwrite latest regulatory logic. |
| Session/performance | Inputs already open/closed during selection; job prevents reentrancy; repeat run adds no duplicate records; benchmark x86/x64 load, extraction, structure, copy, calc, save and peak memory on populated fixtures. |

Build Debug and Release for every implementation slice affecting workbooks/structure/calculation. Keep unit-testable planning/type/identity/conflict logic separate from DevExpress mutation; introduce automated tests there before depending on a future UI harness. Report which fixture, engine, platform and checks actually ran. Build success alone is not workbook acceptance.

## 13. Decisions and missing evidence

Needed before a real bespoke merge:

1. Confirm the original generic lineage and intended variant for the first pilot. The subsequently supplied `D:\Downloads\Master BP v25_0700 (1).xlsb` has matching history through 25.0700 and the same worksheet sequence as Stori; Stori logs four further changes. This is sufficient for a provisional read-only comparison, but not proof that every difference is bespoke or that this exact file was the ancestor. See `Stori_Baseline_Version_History_Review_2026-09-19.md` for verified evidence and limits.
2. Confirm whether the latest pilot target is the Sandbox New Blank or the repository-authoritative Blank. They are distinct files; this review has not replaced either baseline.
3. A short example of a known intended bespoke change and an expected generic-release change, to establish a labelled three-way test case.
4. Agreed financial materiality/precision and which supporting metrics/FFR release must pass. Technical defaults can be proposed, but financial sign-off belongs to Abovo.

Recommended defaults for approval: no input overwrite; no automatic record deletion; exact input copying; latest generic formulas retained; unresolved mapping/structural conflicts block execution; safe mapped type-invalid values may appear in an explicitly labelled draft with errors; acceptance requires complete required coverage and resolved/approved financial differences.

An unconfirmed exact ancestor does **not** prevent implementing Slice A, conducting a same-release control, or analysing the supplied provisional base. It does prevent honestly claiming all of a client's bespoke changes have been identified and preserved automatically without further evidence and review.

## 14. Evidence pointers

Paths below are relative to `C:\Repos\Abovo Summit`; line numbers refer to reviewed commit `68b96f4`.

- `Services/BusinessPlanComparisonService.vb`: Compare 123; assessment 149; population 195; structural loop 218; value eligibility/copy 268; financial comparison 314; batch retry 345; SOCI/check extraction 497-642; preflight 644; XML discovery 722; field/type mapping 779-1059; calculation 1060; input equality/name resolution 1134-1168.
- `Interface/User Interface/BusinessPlanComparisonForm.vb`: participant collection 62; external load 258; assessment 318; execution/save/report 346-458.
- `Services/WorkbookStructureRuleManager.vb`: rule resolution/count 246; specialist families 293; simple families 656; dependency order 1020; insertion/recovery 1057; post-actions 1641; scoped snapshot/count handling 1705/1840.
- `Services/StructureCreation/StructureManager.vb`: datasource and field contracts 234-376.
- `Structure.xml`: repeated header example 443; Funders/Facility master example 11915.
- `Interface/User Interface/DataInterfaceTemplate.vb`: repeated-header axis selection 12269; semantic/legacy NRRI dispatch 12393; fallback row/column expansion 12708/12756.
- `Services/DataService/ChangeManagerV2.vb`: batch application/rollback 94.
- `Services/EngineManagement.vb`: dependency-sensitive calculation 445.
- `Services/TransactionalDB/TransactionalDBSynchroniser.vb`: same-worksheet rule selection 178; deferred reconnect/recovery 1084.
- `Services/WorkbookModelProfiles.vb`: profile/fallback 22; current BP marker validation 89; embedded-definition selection 303.
- `Services/FileManager.vb`: normal/custom Save As 454/567; shared function registration 292.
- `Interface/User Interface/FormMainScreen.vb`: debug target 24; compare form integration 774.

The required audit/context documents provide historical workbook evidence and constraints. This review deliberately distinguishes that evidence from current source inspection and from the runtime validation still to be performed.
