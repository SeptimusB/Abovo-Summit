# Excel / DevExpress checkpoint 2e: recovery inspection and verified reopen

Continuation: `Library/Excel_DevExpress_Engine_Stage2f_2026-09-25.md` adds explicit current-history candidate serialization and revision-bound history evidence, without connecting live dirty acknowledgement or ordinary Save.

## Scope

This continues the isolated, default-off publication trial. Normal Debug and Release **2.98 remain unchanged**. No live Save, startup scan, history/dirty acknowledgement, UI engine preference or workbook schema is switched on. This is not a client test release.

Two new boundaries are implemented:

- `WorkbookPublicationRecovery.InspectAsync(intentPath, expectedTarget)` examines an explicitly selected transaction without changing files or opening a spreadsheet engine.
- `WorkbookPublicationReceipt.ReopenAsync()` opens the exact successfully published bytes as a **new** session/baseline, using the original immutable engine and VBA-security options.

The existing `ModelChangeManagerV2` remains the only intended production edit/history integration point. Neither a recovery inspection nor a reopened revision-zero session marks a live model clean. An older saved revision must not acknowledge newer edits.

## Interrupted publication

New intent/completion records are version 2 and include a bounded digest of recognized provenance streams. Source URLs and raw stream contents are not copied into XML. Version 1 records from earlier trials require manual inspection; they are not silently upgraded.

The parser rejects DTDs, oversized XML (16 KiB), unknown/duplicate fields, invalid hashes/session/revision, mismatched targets/formats/modes and payload paths outside the exact generated transaction folder. The caller supplies the intended workbook path independently. The journal must be `intent.xml` in an immediate `~Summit-save-<guid>` child of that workbook's parent. Existing local-NTFS/reparse/hardlink/alternate-stream restrictions still apply. Inspection pins directories and guards file handles against replacement while reading.

| Observed state | Meaning and permitted action |
| --- | --- |
| Prepared | Staged candidate exists; original remains, or new-name Save As has not published. No automatic write. |
| OriginalRetained | Original is verified in `previous.<ext>`, the candidate is verified, and the public filename is vacant. Explicit restoration is available. |
| PublishedUnacknowledged | Candidate bytes occupy the target, staging name is absent, and the replacement backup is intact where required. Do not repeat the save merely because its completion record is absent. |
| PublishedAcknowledged | The published state also has a matching completion record. This is file evidence, not a live-model dirty acknowledgement. |
| Conflict / Incomplete | Changed, occupied, missing or inconsistent evidence. No automatic publication or restoration. |

`RestoreOriginalAsync(inspection)` reopens/rechecks the evidence, including the exact journal hash, then restores the **verified original handle only to a vacant target**. A late competing writer is never overwritten. The unsaved candidate remains available for review. Cancellation is honored before the rename; after it, the result reports restoration even if writing the advisory completion record fails. No recursive cleanup or backup deletion is performed.

The XML is unsigned diagnostic evidence, not authentication, a financial-integrity certificate or permission to enable macros. This API does not discover or act on arbitrary recovery folders automatically. Caller-selected recovery and an eventual client-facing startup workflow remain distinct responsibilities.

## Reopen/rebase

Only an in-memory successful publication receipt supplies verified reopen. Its exact file SHA256 and provenance digest are checked under guarded handles before creating any native backend, and checked again after opening. Changed bytes/markers reject this path; a disk journal cannot manufacture an equivalent receipt or reconstruct security settings.

The old session remains closed. The new session has a new ID, revision zero and the published file hash as its source baseline. Old results/candidates cannot be applied to it. It can accept another opted-in edit, create another verified candidate and publish that candidate. Initial automatic engine selection/fallback remains subject to the same original options and existing Excel security policy. There is no mid-session engine switch or Trust Center change.

If reopening fails or is cancelled, the successfully saved file is retained. It is not overwritten with an older backup. A new owner created before a failed post-open check is closed. Permanently hung native calls still require the separate process-supervision work.

## Validation

- Isolated main and harness Debug/Release builds pass with zero errors/warnings, using `bin/EngineStage2e-Debug` and `bin/EngineStage2e-Release`. A pilot harness rebuild encountered its own still-running Debug test executable; the final build was rerun after that test exited.
- Per configuration: **109 recovery/reopen assertions**, including three actual owned-process abrupt exits at Prepared, OriginalRetained and Published. These use a controlled backend and real private NTFS files; they do not terminate Excel or certify power-loss durability. Reopen tests also verify identical immutable options and cleanup after provenance changes during opening.
- Existing suites pass per configuration: safety 72, deadlines 6, grid 16, security 12, negative projection 10, edits 62, candidates 54 and terminal publication 76. Generic source-unchanged assertions also pass.
- Generated XLSX/XLSM/XLSB: both engines pass publication, receipt-based same-engine reopen, further edit/save, and other-engine full-calculation comparison. **20 assertions plus 12 output comparisons per configuration.**
- AGL XLSB Debug and XLSM Release: both directions pass same-engine reopen and continued edit/save, then other-engine comparison over **45,586 positions**. Eight assertions plus four full output comparisons pass per configuration.
- Original source, repository Blank/Demo and normal Debug/Release executable hashes are unchanged. Native tests check that pre-existing Excel processes remain and owned processes exit.

The private AGL workflow changes Funding G82 by +300, publishes to its new disposable source copy, reopens, adds a further +100, saves under another new private name and compares the established financial, Check Sheet and PMCost/RespCost rectangles. No original/customer workbook or master is saved. Preservation guidance is applied without redesigning formulas, visual formatting or fill-based editability.

Single-run reopen measurements: AGL XLSB Debug DevExpress **7,853 ms**, Excel **6,064 ms**; AGL XLSM Release DevExpress **15,796 ms**, Excel **15,387 ms**. These are diagnostics only and exclude candidate creation/publication, subsequent full calculation and UI rebinding. They demonstrate why closing/reopening on every ordinary Save would be inappropriate; they are not frequent-Save performance claims, controlled medians or structural benchmarks. Evidence stays in ignored `obj/EngineStage2e-*.log` and private GUID-named fixture folders.

## Remaining integration gates

1. Serialize **current** model/history/recovery XML; opening-baseline XML preservation is only valid for this value-only trial. Connect edits, undo/redo and exact-revision save acknowledgement through the existing change manager, not a second history stack.
2. Design ordinary frequent Save without unnecessarily closing/reopening the owner; integrate authoritative formatting/result refresh and the approved settings into the live UI.
3. Add caller-driven startup discovery/selection and plain client messages, qualified retention/cleanup, storage/attribute/owner/auditing policy, actual storage/power-loss failure tests and hung-native-process supervision.
4. Run the approved Stage 3 complete structural comparison: existing DevExpress .NET, VB.NET-driven Excel and verified master VBA. No new structural timings are claimed here.
5. Complete live UI, Excel/VBA event/menu round trips and Jon/Alex acceptance before production engine switching.

The no-overwrite restore uses the existing handle-based rename with replacement disabled, matching Microsoft's [FILE_RENAME_INFO contract](https://learn.microsoft.com/en-us/windows/win32/api/winbase/ns-winbase-file_rename_info). The two-rename publication remains recoverable, **not an uninterrupted atomic swap**.
