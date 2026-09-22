# Master VBA and DIT structural audit

Date: 21 September 2026. Checkout: `C:/Repos/Abovo Summit`, `main`, working tree including the 2.54 trial. This is an engineering audit, not an accountant's financial sign-off.

22 September follow-up: the user approved the Development/axis/copied-input/linked-range repairs. Implementation and current evidence are in `Structural_Repair_Trial_2026-09-22.md`. Statements below about these paths being unrepaired describe the original audit, not the subsequent trial. OFA/Repairs anchor differences and remaining deletion-boundary review are still outstanding.

## Conclusion

The structural contract is larger than a named range and an Add Lines macro. Worksheet events often coordinate assumption rows, several working-sheet columns, formula templates, sentinel names and Transactional DB mirrors. A generic insertion of the displayed range does not implement that contract.

The approved Funding insert/delete repair is implemented and tested separately. The wider review reproduced Development copy-order discrepancies, a wrong-axis Specific Income insertion, Cash Journal input duplication and a Service Charge input/workings synchronisation failure. These additional production paths have **not** been repaired by this audit.

Do not issue a blanket statement that every Add Lines action is now safe. In particular, avoid using the newly identified paths on a client's only copy pending the next repair and acceptance cycle.

## Scope, preservation and coverage

- Authoritative unpopulated master: `Library/Blank BP v26_0001.xlsb`, SHA-256 `0b4c06800fe998e8733a5d8eb9cdefa04f517750275bb0cfdd532928900f5b79`.
- Authoritative populated master: `Library/Demo BP v26_0001.xlsb`, SHA-256 `1ed79726d8d1129c699242c9e8b3e988ac920ea7108d3c0b105bf9b0bc8a5e9c`.
- Both have **337 VBA modules, 18,470 source lines and 287 procedures**. There are 89 modules containing procedures and 248 without procedures. All modules were inventoried; the full executable text of every procedure-bearing module was reviewed, including worksheet/workbook events, forms and administrative utilities. Comment/attribute-only text was inventoried, not represented as executable behaviour. Credential literals were redacted for review.
- All 287 procedure-source hashes match between Blank and Demo. Six form-module full-source hashes differ; no procedure-body differences were found. This does not assert that form resources or the workbooks themselves are identical.
- The complete metadata inventory is `Library/Master_VBA_Inventory_2026-09-21.json`: module/procedure hashes, source locations, literal range/sheet references and conservative call candidates. It contains no raw VBA or passwords. Lexical call candidates are not proof that a branch executes, particularly with dynamic `Application.Run`, worksheet events, or unqualified `Range`/`ActiveWorkbook` references.
- Reviewed native `WorkbookStructureRuleManager`, its 46 registered rules (10 specialised, 36 simple), the DIT dispatcher and legacy WorkbookManager paths, custom SpecificRowColumnEvents wrappers, Structure.xml expansion declarations and relevant Transactional DB synchronisation paths.
- The XML scan found 58 expansion declarations, including repeated uses of the same rule; 11 declarations resolve to legacy handling. Runtime-generated editor actions require separate tracing; an XML count is not a complete UI-button count.
- Masters and the supplied client original were not modified. Probes loaded disposable copies; diagnostic mutations were unsaved. Funding tests saved only private result/reference copies. Excel automation was a separate instance, with macros/events disabled, no external link update, source books opened read-only and closed without saving. No VBA procedure was executed. VBA extraction dependencies and source remain outside the repository; only metadata and audit conclusions are retained here.

## Findings requiring further repair

### 1. Development: inter-sheet copy order (high priority, reproduced)

The VBA groups the seven Development sheets before copying new columns. Native insertion currently inserts and copies one sheet at a time. The Depreciation and NonCash sheets refer to one another, so simply swapping their order cannot solve both directions.

A three-column differential probe using the same native target/anchor/template settings found **720 formula differences in Dvpt Component Depn for identified schemes and 720 for multi-year schemes**, comparing sequential copy with all-sheet insertion before copy. Newly copied depreciation formulas point into a different scheme's NonCash column in the sequential result.

This is a controlled copy-order defect probe, not yet an independent Excel/VBA-equivalence acceptance test. The eventual repair must also respect VBA's identified-scheme first-column/remaining-columns sequence, special-position insertion and protected leading columns, not just change a Boolean. A separate approval question was sent for both Development paths; they remain unchanged here.

### 2. Seven simple rules use the wrong physical axis (one runtime reproduction, all seven geometry checked)

| Rule / source range | Blank physical range | Native configured axis |
| --- | --- | --- |
| SIMPLE_SUMMARYOTHERINCCAT / SummaryOtherIncCat | Specific Income Assumptions!A7:A13 | Columns |
| SIMPLE_IR_SPEC_INC_ASS1 / IR_Spec_Inc_Ass1 | Specific Income Assumptions!A21:A25 | Columns |
| SIMPLE_CAPITALGRANTCATS / CapitalGrantCats | Capital Grant Assumptions!A6:A8 | Columns |
| SIMPLE_IR_INTCO_INC_ASS1 / IR_Intco_Inc_Ass1 | Intercompany Income Assumptions!A6:A10 | Columns |
| SIMPLE_IR_INTCO_INC_ASS2 / IR_Intco_Inc_Ass2 | Intercompany Income Assumptions!A15:A24 | Columns |
| SIMPLE_IR_ONEOFF_COST_ASS / IR_Oneoff_Cost_Ass | Management Costs Assumptions!A60:B71 | Columns |
| SIMPLE_IR_REP_REPHASE / IR_Rep_Rephase | Repairs & Maint. Rephasing!A8:A18 | Columns |

Actual `AddRecords(SIMPLE_IR_SPEC_INC_ASS1, 3)` returned success but changed A21:A25 into **A21:D25**: still five rows, four columns instead of one. Presentation transposition must not determine the physical worksheet axis. Correcting the axis alone is insufficient for Specific Income because its VBA also expands working sheets (finding 4).

### 3. Simple row inserts duplicate existing inputs (reproduced)

`InsertRowsForTarget` copies `PasteSpecial.All` from the preceding row and has no clearing policy for editable inputs. On an unsaved Blank copy, a sentinel text value placed in the unlocked description cell of the last Cash Journal row appeared in **all three newly added rows**.

The VBA `Row_Insertion.InsertRows` clears unlocked cells after copying; the native legacy `WorkbookManager.InsertRows` clears unlocked constants while retaining formulas. The simple structural-rule path does neither. A repair needs an explicit, workbook-compatible policy for new editable constants, defaults and formula-backed unlocked cells; do not indiscriminately clear every formula or extend Funding's template policy to unrelated families.

### 4. Coupled assumption/workings operations are missing (Service Charge reproduced; other paths statically traced)

| Family | Master VBA contract | Current native gap / evidence |
| --- | --- | --- |
| Service Charge | Sheet225 change of IR_ServChg_01 calls row insertion and Service_Charge_Columns over five working sheets | Actual SIMPLE_REP_SERVCHG_02 add-three enlarged B19:C24 to B19:C27, but IR_ServChg_01 remained A19:A23 and LastUnitSCColumn remained Unit Service Charges!J6. It returned success. Both named-input coverage and the workings boundary failed to grow. |
| Specific Income | Sheet275 grows IR_Spec_Inc_Ass1 rows and four Specific Income working sheets; separate assumptions-column action also exists | Registered rule has wrong axis and only one assumptions-sheet target; no equivalent coordinated workings action found in the native route. |
| Other Income | Sheet53 grows IR_Oth_Inc_Ass5 rows plus Other Income Workings columns; IR_Oth_Inc_Ass6 is a separate row-only case | XML resolves both to legacy row insertion. The required workings operation was not found in that native path. Separate assumptions-column VBA path must also be mapped. |
| Joint Venture | Sheet79 invokes Joint_Venture_Columns over seven sheets, then four TDB mirrors | IC_JointVenture_01's XML/runtime NRCI route is a generic single-sheet column insertion; no registered semantic JV structure rule. Existing TDB mirrors are not a substitute for resizing the seven working sheets. |
| Intercompany Funding | Sheet226 routes IC_IntercoFunding_01/02 to Interco_Funding_Columns across twelve sheets, then six mirrors; separate specialised row action exists | No corresponding semantic rule/coordinated native column routine found. Trace the exact exposed UI action before claiming every client can invoke the defective path; presence of TDB rules alone is not implementation of the VBA operation. |

The five Service Charge working sheets are Unit Service Charges, Service Charge Numbers, Service Charge Income, Service Charge Voids and Service Charge Bad Debts. The Rep range includes a different boundary from the IR range: extending only Rep is not sufficient. This resembles the already repaired Journals sentinel/name issue, but requires its own contract rather than a generic blanket resize.

### 5. OFA and Repairs: anchor/template differences (static discrepancy, not yet adjudicated)

- Master OFA insertion is one column before LastOFACol with the VBA template/copy behaviour across two sheets. Native uses the anchor itself, a preceding-column template, and formats/column-width only on its assumptions sheet versus VBA's full paste.
- Master Repairs insertion is one column before LastStockCol across eight sheets. Native uses the anchor itself and preceding-column template.
- A copy-order-only probe found no differences for these families. That **does not validate different insertion anchors, template contents, deletion boundaries or Excel round-tripping**. Do not alter these established paths without an independent Excel comparison and explicit repair scope.

### 6. Deletion bounds and legacy dispatch (static risks)

Simple registered rules generally retain a minimum of one record. VBA event handlers have family-specific limits and sentinels (often three, Service Charge row minimum four, category/column minimum two, and additional special leading-column limits). Some positive-only VBA paths expose no equivalent negative action. Do not infer a deletion contract from range width alone.

Funding's approved correction now counts only LoanDescsOrd, protects the first ten ordinary columns and excludes revolvers. Other families still require their own minimum, protected-template and allowed-position matrix.

DIT's unmatched NRRI fallback calls legacy `WorkbookManager.InsertRows`; the examined branch does not propagate its returned error before rebuilding, and its cursor handling lacks an equivalent finally block. NRRIbyCOL also uses a rows-versus-columns shape heuristic. These should become explicit structural dispatch decisions, with errors propagated and no interface success reported after a failed mutation. No production changes made here.

## Multi-sheet copy-order probe coverage

This probe isolates native operation ordering using three new columns; it is not an end-to-end Excel validation of every family.

| Native family | Formula cells compared | Sequential vs staged differences | Interpretation |
| --- | ---: | ---: | --- |
| OFA_RECORDS | 3,995 | 0 | No ordering mismatch in this sample; anchor/copy-policy questions remain |
| CAPEX_RECORDS | 1,615 | 0 | Ordering only checked |
| CAPGRANT_RECORDS | 1,529 | 0 | Ordering only checked; category row rule separately has wrong axis |
| REPAIRS_RECORDS | 21,266 | 0 | Ordering only checked; anchor question remains |
| HOUSING_COMPONENT_RECORDS | 156,219 | 0 | Ordering only checked |
| FUNDING_RECORDS | 86,964 | 2,436 | Motivated approved staged-insert repair; later Excel acceptance below supersedes this probe |
| DEVELOPMENT_IDENTIFIED_RECORDS | 124,449 | 720 | Further repair required |
| DEVELOPMENT_MULTIYEAR_RECORDS | 123,105 | 720 | Further repair required |

JOURNAL_RECORDS and STOCK_CONVERSION_RECORDS are row families, not part of that eight-family column-order comparison. Journals' separate IR/mirror repair and round-trip evidence are in `Presentation_Journals_Funding_Trial_2026-09-21.md`. Stock Conversion was traced through its rule/linked targets, but not independently Excel-roundtrip tested by this audit.

The remaining 36 simple rules were all inventoried. They cover Rent Weeks, two Service Charge ranges, Stock Disposal, Owner Occupier, the seven wrong-axis ranges above, three Additional Pension ranges, capitalised Development interest, sixteen Economic Assumption ranges, two capitalisation ranges and two Cash Journal ranges. This is metadata/route coverage, not 36 passing insertion/deletion tests.

## Full VBA project: other relevant observations

The review included workbook lifecycle, application/event/calculation state, error logging, column generation, all row/column families, DSA import/remove/template operations, Stress Test capture/import/sensitivity, FFR/WG exports, rent/management/repair imports, amendments history, menus, protection, user forms, colouring, code management, administration and custom functions.

Important distinctions for future work:

1. **Disabled formula generation is intentional in these masters.** Category_Changes, InitialColumnCheck, AllColumns, Launch_AllColumns and CreateFormGenMenu contain early exits. Do not reactivate them or treat missing execution as a Summit gap; generated formulas are retained in the authoritative files.
2. **ABVCalculate contains the known unconditional early exit.** Preserve Summit's intended recalculation points; do not reproduce this VBA accident. This affects how one interprets callers in FFR/Stress Test/menu code, not evidence that every native caller is incorrect.
3. **State-restoration hazards exist in legacy VBA.** Examples include early returns after InitiateCodeRun in DSA/administrative flows and SetTransDBMirrorRangeSize exiting after changing application flags. Broad On Error Resume Next is used in many modules. Summit must preserve entry state and propagate failures, rather than mechanically translate these behaviours.
4. **Two suspect function returns are visible statically:** DevelopmentTransDBNRs assigns DvptSheets rather than its own function result; intFDoesFileExist assigns DoesFileExist rather than its own result. Their actual reachable callers and impacts require targeted runtime investigation; no source-master edits made.
5. **Custom functions retain legacy constraints.** PMCost/RespCost use Integer parameters and suppressed errors, approximate-match lookups and an unused FinalYear argument; ThisDate/AllDates use sentinel dates. Native equivalence should follow the separately documented custom-function contract and dedicated test cases, not a casual rewrite during structural repair.
6. **Administrative macros are not ordinary insertion contracts.** MinimiseInputRowsAndColumns, RemoveSampleData, ReleaseFile, code export/replacement and bulk colouring can intentionally make broad/destructive changes. They were reviewed as part of the full project but not run, copied into Summit, or treated as authorization to modify masters.
7. **External import/export and scenario workflows remain separate acceptance areas.** Static review covers their calls and assumptions; it does not validate every imported third-party file, FFR template version, DSA layout, UDF result or Stress Test output. ActiveWorkbook/ActiveSheet and dynamic calls deserve special care in a multi-model native application.

## Approved Funding repair and verification boundary

See `Funding_Insert_and_DIT_Save_2026-09-21.md` for implementation and client-file findings. Funding now inserts all 32 linked sheets before copying the shifted VBA template, resizes names, synchronises eleven mirrors, and applies a narrowly scoped 3-D-reference correction using the DevExpress formula syntax tree. Its deletion path protects ordinary/revolver boundaries.

Both Blank and Demo passed actual native eight-column add/delete tests. Each expanded result matched **104,014 linked-sheet formula cells** against an independent macro-disabled Excel grouped-insertion reference and again after Excel save-copy/reopen. All 337 VBA module hashes survived each tested Excel round-trip. Excel returned no circular-reference address in that test; this is not a guarantee for the already damaged client file or a financial/VBA execution sign-off.

No automatic repair-on-load, new worksheet, hidden metadata or master schema migration was introduced. The already expanded client workbook remains unchanged and requires a separate agreed recovery step.

## Recommended next repair sequence

1. Development identified/multi-year operation order, template selection and protected boundaries, with independent Excel equivalents.
2. Replace the seven wrong-axis registrations; implement complete Specific Income and Service Charge contracts (not just visible range extension).
3. Explicit row-template input/default clearing policy, with non-zero populated sentinel tests and formula preservation.
4. Other Income, Joint Venture and Intercompany Funding coordinated operation rules; cover both row and column variants actually exposed by DIT.
5. Independent OFA/Repairs anchor/template comparisons; remaining family deletion minima and error propagation.

For each family: encode physical axis, affected sheet group, insertion anchor, template, new-input policy, retained sentinels, explicit named-range resize set, TDB mirrors, protected/minimum records and calculation/invalidation obligations. Put these reviewed contracts behind the Structure Manager; XML should select a validated semantic rule rather than ask the UI to guess it.

Acceptance per family must include blank and non-zero populated copies, add one/many, permitted delete and rejected protected delete, exact formula/name/validation/protection checks, TDB dimensions, Summit save -> Excel reopen/calculate -> save-copy -> Summit reopen, unchanged VBA hashes, Check Sheet comparison, detailed SOCI/Financial Position comparison and client-accountant review. Automated VBA execution is not part of this audit.

## Reproducibility and evidence

- `Tools/Audit-MasterVba.py`: in-memory complete VBA inventory and sanitized executable-text review. Requires oletools installed outside the repository.
- `Tools/Test-StructuralOrderAudit.ps1`: eight-family operation-order differential; `obj/structural-order-audit.log` and `obj/StructuralOrderAudit/a15b00351b4c4922989109e3dbc47777`.
- `Tools/Test-DitStructureAudit.ps1`: all rule geometries/XML routes and deliberate unsaved Service Charge, Specific Income and Cash Journal defect probes; `obj/dit-structure-final-audit.log`. A successful diagnostic run confirms the stated defects, not product correctness.
- `Tools/Test-FundingStructure.ps1`, `Tools/Test-FundingExcelRoundtrip.ps1`, `Tools/Test-Structural3D.ps1`: approved repair tests; detailed paths recorded in the Funding report.
- Source hashes were checked after test runs. Local `obj` files are disposable evidence, not distributed application dependencies; the conclusions and metadata inventory above are durable.

Production authority: Funding insertion/deletion and DIT Save were approved. Additional audit findings are reported, not silently repaired. Client DPI investigation and manual dialogs/multi-model UI tests remain outstanding. No commit or push was requested for this turn.
