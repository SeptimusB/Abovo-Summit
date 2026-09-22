# Outstanding work after the Summit 2.66 checkpoint

This is the current work list, not a claim that every historical audit item is an active defect. Implemented repairs and their engineering checks are checkpointed; client, accountant and interactive Excel/VBA acceptance remain separate.

## Agreed requirements

- Minimum RAM: **32 GB**.
- CPU: **Core i5 acceptable; Core i7 or better preferred**. A minimum processor generation/model has not yet been specified.
- OS: **Windows 10 or later**. The exact supported Windows/Excel build matrix and release-support policy still need defining.
- **Performance takes priority over continuously checked accuracy, provided occasional integrity checks are retained.** Normal input and save workflows should not repeatedly pay for full scans/rebuilds. Define when checks run, keep pending/unverified results identifiable, and surface detected integrity failures. This does not authorise silent workbook damage, loss of formulas/VBA/user edits, or labelling unchecked output as verified.
- This checkpoint records the new policy; it does not silently introduce a further calculation/validation change. Existing value-save deferral, XML IsCalculated behaviour, result-reader gates and structural/failure safeguards remain as implemented in 2.66.
- Earlier proposed 8/16 GB VM profiles are superseded by the 32 GB minimum. A constrained VM is a repeatable comparison environment, not proof of real client performance.

## Priority work

1. **Reduce save and recovery interruptions.** Native serialization still dominates normal save, and recovery took approximately 13 seconds on the high-spec development machine. Measure blocked-UI time and investigate a safe way to reduce/defer interruptions without concurrent access to the live workbook. The current recovery feature is idle-scheduled, not nonblocking background serialization.

2. **Define the performance-first integrity policy.** Agree and implement a measured occasional-check schedule: periodic/idle checks, appropriate completed-batch milestones and explicit user checks, rather than per-edit/per-save full verification. Decide which output/publication operations need a current check, how provisional results are indicated, and how a failed check affects the next action. Retain transactional rollback, entry protection and Excel/VBA round-trip safeguards. Do not simply remove all checking.

3. **Continue Funding/Development/Transactional DB optimisation.** Obtain fresh ten-record UI traces with the client's usual Analyser/Snapshot arrangement. Target measured linked-column and mirror-row shifts, especially the expensive warm analyser-plus-snapshot case. Multi-year Development performance still needs a controlled comparison. Prior formula/name and round-trip evidence remains a regression guard, not a requirement to repeat every full scan on every interactive step.

4. **Close the remaining structural-audit gaps.** Independently compare OFA and Repairs insertion anchors/templates with master VBA; check deletion minima/protected boundaries for the remaining families and broaden one/many/add/delete coverage. Development, wrong-axis registrations, copied-input clearing and linked-family repairs are already implemented; do not list them as still wholly unrepaired. Investigate the reachability/impact of the two suspect legacy VBA function-return assignments before proposing any master changes. See `Master_VBA_and_DIT_Structural_Audit_2026-09-21.md` and its subsequent trial reports.

5. **Run representative performance tests.** Establish a 32 GB, agreed Core-i5-class baseline plus at least one actual client machine. Use Release without a debugger; compare Summit x86/x64, first/repeated open/save, single edits, full calculation, Funding/Development insertion and recovery. Record CPU/OS/Excel versions, peak process memory, UI-blocked time and workbook identity. Test normal client storage separately from local SSD results. No VM is configured by this checkpoint.

## Acceptance of implemented work

6. **Client/accountant and Excel/VBA acceptance.** The user reported Balance Sheet functional checks passing; accountant review remains open. Check SOCI/Financial Position/Check Sheet and Summit -> trusted Excel/VBA -> Summit on copies, including the repaired Funding client file, additions/deletions, normal/recovered saves, FFR and Stress Test workflows. Recovery validation found Excel implicit-intersection changes in 1,140 Development Expenditure formulas; 44,906 native recalculated cells agree, but this is not an independent financial/VBA sign-off.

7. **Client UI and recovery acceptance.** Funding keyboard navigation and month-end editor; DIT Save/Save As and clean-state button; sidebar compact/restore strips; normal/5k/mixed-DPI icons, fonts and restored layouts. Test recovery enable/disable, scheduling, notice/completion/failure, newer-copy prompt, read-only prior history, guided XLSB Save As, locked/network destinations and multiple models. Engineered/native tests have passed; client-environment acceptance is not implied.

## Next product increments / deferred decisions

8. **Structure Manager and bespoke upgrades.** Review the three-pane trial with Abovo. Extend it into a reviewed, versioned structure/upgrade contract: stable identities, master/related-range families, mixed month/year axes, bespoke definitions, conflict resolution and controlled XML publication/embedding. Add the controlled old-populated/original-template/new-template migration execution with source/template read-only, a separate result, report Save As and reconciliations. The existing generic population service and metadata editor are not that complete bespoke executor. Shared registry/SharePoint integration remains a separate decision.

9. **Optional Excel structural offload.** Investigation and XLSB/XLSM measurements are recorded, but there is no production handoff implementation. Decide whether an isolated prototype is worthwhile given checkpoint/open/reload overhead. It would need a controlled compatible VBA entry point, installed-Excel/trust checks, private result validation, history checkpoint and failure recovery. No macro-security bypass or Excel Interop package has been added. Keep XLSB as normal format; XLSM recovery is already implemented.

10. **Recent-files launcher (deferred).** At client-test release, revisit the requested persistent newest-first recent-file list, rendered as HTML launch links and refreshed on open. It has not been implemented.

11. **Client/live release preparation.** Complete the agreed acceptance matrix, confirm supported OS/Excel/architecture combinations and client installers, configure/test the existing Model/Structure Manager password gate for live use (trial remains deliberately off), then publish the agreed release. No new build/version is introduced merely by this documentation/commit checkpoint.

## Implemented checkpoint scope

The checkpoint includes the Balance Sheet analyser/snapshot repair, Funding editor navigation and month-end dates, FileInstance/sidebar/Check Sheet presentation, DIT save controls, Journals/Funding/approved linked structural repairs, diagnostics and measured optimisations, XLSB compatibility/fast-save state, opt-in XLSM recovery/history and shared progress notices, with their tests and evidence reports. The existing charts and Structure Manager trial were committed before this checkpoint.

Local `.codex/` configuration, ignored build/test outputs, private repaired/benchmark workbooks, extracted VBA and dependencies are not part of the commit. Original client/master workbooks are not changed. A commit records implementation; it does not close the acceptance and performance tasks above.
