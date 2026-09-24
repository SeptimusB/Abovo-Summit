# Outstanding work after the Summit 2.66 checkpoint

This is the current work list, not a claim that every historical audit item is an active defect. Implemented repairs and their engineering checks are checkpointed; client, accountant and interactive Excel/VBA acceptance remain separate.

22 September client-review update: [checkpoint results and review gates](Client_Review_Checkpoints_2026-09-22.md) and [all 90 report issue paragraphs](Client_Report_2026-09-22.md) retain stakeholder colours separately from engineering evidence. The automated ten-Funding-column/recovery/Excel-copy test passes its stated gates, with the documented formula-text normalization caveat. 2.68 adds verified SHG/date-caption, post-edit width and Journal/Other Current Assets decimal-type repairs. This does not close the remaining report, performance or financial-acceptance work.

## Agreed requirements

- Minimum RAM: **32 GB**.
- CPU: **Core i5 acceptable; Core i7 or better preferred**. A minimum processor generation/model has not yet been specified.
- OS: **Windows 10 or later**. The exact supported Windows/Excel build matrix and release-support policy still need defining.
- **Performance takes priority over continuously checked accuracy, provided occasional integrity checks are retained.** Normal input and save workflows should not repeatedly pay for full scans/rebuilds. Define when checks run, keep pending/unverified results identifiable, and surface detected integrity failures. This does not authorise silent workbook damage, loss of formulas/VBA/user edits, or labelling unchecked output as verified.
- This checkpoint records the new policy; it does not silently introduce a further calculation/validation change. Existing value-save deferral, XML IsCalculated behaviour, result-reader gates and structural/failure safeguards remain as implemented in 2.66.
- Earlier proposed 8/16 GB VM profiles are superseded by the 32 GB minimum. A constrained VM is a repeatable comparison environment, not proof of real client performance.

## Priority work

**2.82 client-testing checkpoint:** diagnostic tracing/benchmarks are temporarily off; re-enable with the build property recorded in [the delivery note](Client_Test_Delivery_2.82_2026-09-23.md) before collecting new timings. [Client review v01](Client_Review_2.82_v01.docx) retains every original issue and separates orange believed fixes from open questions. The deferred recent-files HTML launcher remains to discuss at this release handoff, not to implement without agreement.

**Client input/presentation review (23 September):** [confirmed findings and clarification queue](Client_Input_Review_2026-09-23.md); approved checkpoint 1 is implemented in [2.81 typed-input/source-mapping repairs](Client_Input_Repairs_2026-09-23.md). Next are workbook conditional formatting/dependent editors, then agreed layouts. Engineering checks and client acceptance remain distinct; 2.80 structural repairs are retained for client testing.

1. **Reduce save and recovery interruptions.** Native serialization still dominates normal save, and recovery took approximately 13 seconds on the high-spec development machine. Measure blocked-UI time and investigate a safe way to reduce/defer interruptions without concurrent access to the live workbook. Recovery is scheduled, not nonblocking background serialization.

   23 September / 2.78: [timing, snooze and discarded-warning trial](Recovery_Timing_and_Discard_Warnings_2026-09-23.md) adds independent idle/maximum intervals, a pre-write snooze-until-one-minute-idle notice, exact-model recheck prompts for remembered failures, and guarded clearance when an entirely unsaved session is discarded. Maximum timing may interrupt only after safety gates; native serialization cannot be cancelled midway. Client acceptance remains open.

   23 September / 2.77: [recovery packaging optimisation](Recovery_Packaging_Optimisation_2026-09-23.md) removes unnecessary worksheet recompression while retaining every metadata check. Identical-export median metadata preparation falls from 2,420 to 879 ms. A separate full AGL recovery takes 14,627 ms, dominated by 13,573 ms of native snapshot export. Detailed phase traces and preservation/failure/Excel-copy regressions are in place. This is a modest stage-level gain, not a solved UI-blocking or ordinary-XLSB-save problem; client measurements remain required.

2. **Validate the performance-first integrity policy.** The 2.67 opt-in Integrity trial adds a configurable interval, two-minute session-idle requirement and staged inspection; running calculations are never interrupted. Native safety/state/history tests pass on Blank/Demo copies. Complete physical-input, multiple-model, client performance and Excel/VBA acceptance, and classify existing reported chart/menu errors. Decide publication gates, additional bespoke/3-D checks and explicit/batch check entry points separately. See `Idle_Integrity_Trial_2026-09-22.md`. Existing rollback, entry protection and save/close safeguards remain.

3. **Continue Funding/Development/Transactional DB optimisation.** Obtain fresh ten-record UI traces with the client's usual Analyser/Snapshot arrangement. Target measured linked-column and mirror-row shifts, especially the expensive warm analyser-plus-snapshot case. Multi-year Development performance still needs a controlled comparison. Prior formula/name and round-trip evidence remains a regression guard, not a requirement to repeat every full scan on every interactive step.

4. **Structural-audit acceptance and residual VBA review.** The 23 September / 2.80 [OFA/Repairs and deletion audit](Structural_Audit_Repairs_2026-09-23.md) repairs the confirmed insertion/template defects, row/column deletion boundaries, Economic service-charge mirror collision and Repairs Include-range drift. It records broad AGL add/delete tests, maximum/minimum boundaries, current-master regression and independent Excel comparisons. These are no longer wholly unimplemented tasks. Complete client UI and trusted Excel/VBA/financial acceptance, retain bespoke/older-model geometry review, and investigate the reachability/impact of the two suspect legacy VBA function-return assignments before proposing master changes. No source master or original AGL was modified.

5. **Run representative performance tests.** Establish a 32 GB, agreed Core-i5-class baseline plus at least one actual client machine. Use Release without a debugger; compare Summit x86/x64, first/repeated open/save, single edits, full calculation, Funding/Development insertion and recovery. Record CPU/OS/Excel versions, peak process memory, UI-blocked time and workbook identity. Test normal client storage separately from local SSD results. No VM is configured by this checkpoint.

## Acceptance of implemented work

6. **Client/accountant and Excel/VBA acceptance.** The user reported Balance Sheet functional checks passing; accountant review remains open. Check SOCI/Financial Position/Check Sheet and Summit -> trusted Excel/VBA -> Summit on copies, including the repaired Funding client file, additions/deletions, normal/recovered saves, FFR and Stress Test workflows. Recovery validation found Excel implicit-intersection changes in 1,140 Development Expenditure formulas; 44,906 native recalculated cells agree, but this is not an independent financial/VBA sign-off.

7. **Client UI and recovery acceptance.** Funding keyboard navigation and month-end editor; DIT Save/Save As and clean-state button; sidebar compact/restore strips; normal/5k/mixed-DPI icons, fonts and restored layouts. P016's confirmed Enter/Shift+Enter traversal is implemented in 2.69: Up/Down selects vertical, Left/Right/Tab selects horizontal; next-grid/standalone-input movement and reverse/wrapping/scrolling remain within the visible tab. Both native Debug/Release suites pass; physical-keyboard/client-workbook acceptance remains open. See `DIT_Enter_Navigation_Trial_2026-09-22.md`. Test recovery enable/disable, scheduling, notice/completion/failure, newer-copy prompt, read-only prior history, guided XLSB Save As, locked/network destinations and multiple models. Engineered/native passes cover their stated scope and do not imply client-environment acceptance.

## Next product increments / deferred decisions

8. **Structure Manager and bespoke upgrades.** Review the three-pane trial with Abovo. Extend it into a reviewed, versioned structure/upgrade contract: stable identities, master/related-range families, mixed month/year axes, bespoke definitions, conflict resolution and controlled XML publication/embedding. Add the controlled old-populated/original-template/new-template migration execution with source/template read-only, a separate result, report Save As and reconciliations. The existing generic population service and metadata editor are not that complete bespoke executor. Shared registry/SharePoint integration remains a separate decision.

9. **Optional Excel structural offload.** Investigation and XLSB/XLSM measurements are recorded, but there is no production handoff implementation. Decide whether an isolated prototype is worthwhile given checkpoint/open/reload overhead. It would need a controlled compatible VBA entry point, installed-Excel/trust checks, private result validation, history checkpoint and failure recovery. The additional Teams discussion treats a continuously parallel Excel shadow/comparison workbook as nice-to-have, not a current prerequisite; it does not authorise removing rollback or integrity safeguards. No macro-security bypass or Excel Interop package has been added. Keep XLSB as normal format; XLSM recovery is already implemented.

10. **Recent-files launcher (deferred).** At client-test release, revisit the requested persistent newest-first recent-file list, rendered as HTML launch links and refreshed on open. It has not been implemented.

11. **Client/live release preparation.** Complete the agreed acceptance matrix, confirm supported OS/Excel/architecture combinations and client installers, configure/test the existing Model/Structure Manager password gate for live use (trial remains deliberately off), then publish the agreed release. No new build/version is introduced merely by this documentation/commit checkpoint.

## Implemented checkpoint scope

The checkpoint includes the Balance Sheet analyser/snapshot repair, Funding editor navigation and month-end dates, FileInstance/sidebar/Check Sheet presentation, DIT save controls, Journals/Funding/approved linked structural repairs, diagnostics and measured optimisations, XLSB compatibility/fast-save state, opt-in XLSM recovery/history and shared progress notices, with their tests and evidence reports. The existing charts and Structure Manager trial were committed before this checkpoint.

Local `.codex/` configuration, ignored build/test outputs, private repaired/benchmark workbooks, extracted VBA and dependencies are not part of the commit. Original client/master workbooks are not changed. A commit records implementation; it does not close the acceptance and performance tasks above.
