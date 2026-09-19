# GPT-6 Astra handover prompt

Use GPT-6 Astra with high or xhigh reasoning. Paste the following as the first message after switching models.

---

You are continuing a long-running Abovo Summit development task in `C:\Repos\Abovo Summit`.

Your first objective is to take ownership of the Business Plan upgrade/migration architecture. Do not start by modifying production code. First complete a read-only evidence and architecture review, then return a concrete, reviewable proposal for the user.

Read these files completely before forming the proposal:

1. `AGENTS.md`
2. `Library/Upgrade_Migration_Handover_2026-09-19.md`
3. `Library/Upgrade_Migration_Test_Matrix_2026-09-19.md`
4. `Library/Upgrade_Migration_Handover_Index_2026-09-19.json`
5. `Library/Abovo_Summit_Project_Scope_Audit.md`
6. `Library/Abovo_Summit_Project_Index.json`
7. `Library/Contract_XLSB_Audit_2026-08-24.md`
8. `Library/Contract_XLSB_Audit_Evidence_2026-08-24.json`

Then inspect the current implementation and its integration, beginning with commit `b8b231c` and these files:

- `Interface/User Interface/BusinessPlanComparisonForm.vb`
- `Services/BusinessPlanComparisonService.vb`
- `Services/WorkbookStructureRuleManager.vb`
- `Services/FileManager.vb`
- `Services/WorkbookModelProfiles.vb`
- `Services/StructureCreation/StructureManager.vb`
- `Services/DataService/ChangeManagerV2.vb`
- `Services/TransactionalDB/TransactionalDBSynchroniser.vb`
- `Structure.xml`

Business context:

- Abovo supplies Excel XLSB Business Plan models used by client organisations to assess financial viability and meet statutory requirements.
- Models change continually, especially the government FFR template.
- Clients populate generic models and some commission bespoke changes to a base model.
- Abovo must carry client inputs and approved bespoke changes onto later generic releases. This consumes about 30% of its working year.
- Filenames generally hint at the base release, but future Summit-compatible workbooks will contain embedded XML describing their structure.
- Most upgradeable named ranges use a `Rep` prefix, but XML, cell formulas/protection and structural rules must remain authoritative.
- The client often has a master named range whose expansion, when formula contiguity is preserved, expands or populates all associated ranges. These master/dependent relationships are historically inconsistent and must be discovered and encoded rather than guessed in production.
- All historical models are held in SharePoint, but no direct SharePoint access is currently configured. A locally synced library is the initial low-risk access route.

Required product contract:

- The XLSB is authoritative.
- Source client files, old generic templates and new generic templates must remain read-only.
- Clone the new template first and mutate only the result.
- Preserve formulas, defined names, formatting, protection, worksheet order, VBA/custom UI and Excel/VBA round-tripping.
- Never store workbook/VBA passwords or extracted VBA source.
- Route interactive/value changes through the established change manager and structural changes through `WorkbookStructureRuleManager`.
- Do not introduce hidden metadata/template/materialisation sheets without explicit approval.
- Preserve Transactional DB reconciliation and Check Sheet validation.
- Treat FFR as a versioned regulatory subsystem, not merely a cell range.
- Preserve unrelated user changes in the working tree.

The target bespoke upgrade is a four-artifact, three-way merge:

1. old generic template, read-only;
2. old populated/bespoke client workbook, read-only;
3. latest generic template, read-only;
4. editable result cloned from the latest template.

Distinguish:

- client input data;
- changes made only in the old client branch;
- changes made only in the new generic release;
- genuine conflicts where both branches changed the same semantic element;
- intentional new-release output differences;
- unsupported or unsafe changes.

Your first deliverable should:

1. restate the verified current implementation and identify any discrepancy with the handover;
2. map the structural rules and XML fields that already express master/dependent range behaviour;
3. identify implicit structural conventions that still need workbook evidence;
4. propose an embedded model identity and migration metadata schema;
5. propose the immutable migration-plan object model and ordered execution stages;
6. define three-way classification and conflict rules for values, formulas, names, geometry, formatting, validation, protection, sheet order, custom XML, VBA/custom UI and FFR;
7. define transaction, rollback, recovery and idempotency behaviour;
8. define human and machine-readable reports with Information, Notice, Warning, Error and Conflict severities;
9. define financial acceptance rules for Check Sheet, detailed SOCI and FFR when generic calculations intentionally change;
10. recommend the smallest safe implementation slice and a representative workbook pilot;
11. list any material user decisions or missing files required before implementation.

Use direct evidence from the repository. Distinguish verified facts, reasonable inferences and proposals. Do not claim that workbook/VBA behaviour has been validated unless it was actually tested. Do not modify or save any XLSB during this first review. If workbook inspection is needed, work from disposable exact copies, use read-only Excel automation with macros/events disabled and close without saving.

Work autonomously through read-only inspection. Ask the user only when a missing choice, file or authority would materially change the design. Use concise prose, but make the architecture complete enough to implement without rediscovering the domain.

---

Official OpenAI guidance describes GPT-6 Astra as suited to long, multistep software-engineering work and recommends explicitly prompting for persistence, instruction priority and calibrated verification. The task prompt above incorporates those points while retaining the repository's stricter workbook-safety rules: <https://developers.openai.com/api/docs/guides/latest-model>.

