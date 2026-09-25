# Excel / DevExpress checkpoint 2d: terminal publication trial

## Scope

Continuation of checkpoint 2c, behind a separate **default-false** `EnablePublicationTrial` flag (requires candidate-save opt-in). Production Debug/Release **2.98 is unchanged**. This is not a client test release, a live Save replacement or a new financial/calculation policy.

`PublishAndCloseAsync(candidate, target, mode)` accepts an exact session/revision candidate. `CreateNew` writes a new same-format filename and refuses every existing destination. `ReplaceSource` can name only that session's source, and checks its opening SHA256 and provenance markers after the native engine has actually closed. Overwriting an unrelated Save As destination and format conversion are deliberately unavailable.

The operation consumes its isolated native session. It does **not** reopen/rebase it, serialize current production history XML, mark a live model clean, or acknowledge newer edits. Existing `ModelChangeManagerV2` remains the required production integration point. All value-only XML/VBA/array restrictions from checkpoint 2c still apply.

## Publication protocol

1. Reserve the session revision; reject overlapping edits, invalidation, exports and publication. Validate the candidate bytes and stage a flushed copy on the destination volume. Keep the verified candidate available independently.
2. Write a flushed `intent.xml` in an exact generated `~Summit-save-<guid>` directory, containing source/target/candidate/backup paths, hashes, session and revision. No financial data, VBA source or passwords are placed in the record.
3. Close the native owner on its STA and release the original lease only after disposal. Cleanup failure or timeout prevents publication; a late native return cannot resume abandoned publication.
4. Open the current source through a Windows handle which denies other data writes/deletes. Check source bytes and markers again. Keep parent directories pinned against renaming/reparse replacement. Preserve and verify the source discretionary ACL on the replacement.
5. For replacement, rename the **verified source handle** to `previous.<extension>` without overwriting. Rename the verified staged handle to the requested destination, also without overwriting. A competing file which appears in between is never replaced. On a pre-publication failure, restore the original name only if vacant; otherwise retain the original backup and report the conflict.
6. Return the receipt only after publication and its completion record. Cancellation is honored before the first rename, not halfway through acknowledgement. A written file followed by an acknowledgement failure is explicitly reported as `WorkbookPublicationException.Published=True`, with recovery paths; callers must not blindly retry.

**This is a recoverable two-rename protocol, not an uninterrupted atomic swap.** The public name is briefly absent between renames. An abrupt exit there leaves the original in the recorded backup location. Automatic startup inspection/recovery and a client-facing failure workflow are required before production use. Power-loss durability and arbitrary filesystem/network behavior have not been certified. The filesystem commit deliberately has no abandon-on-timeout wrapper that could falsely report an already-written file as cancelled.

## File/security scope

- This trial accepts local fixed NTFS paths only. It refuses UNC/network paths, reparse/offline/encrypted files or directories, hardlinked files, read-only replacements, unsupported alternate streams and long paths. Hydrated sync-folder detection is not implemented; cloud/sync folders are not qualified merely because they reside on NTFS.
- Ordinary `Zone.Identifier` is retained. The first real AGL publication was correctly refused because the input also carries `MBAM.Zone.Identifier`. Both recognized markers are now captured at session opening and preserved **byte-for-byte**, including embedded NULs, within a 64 KiB-per-marker limit. Secondary provenance does not grant macro permission, remove Windows marking or change Trust Center settings. Unknown streams still cause refusal.
- Source/marker changes during hand-off, changed staged bytes/markers and destination collisions are rejected. Only owned native Excel processes are closed.
- A custom-ACL regression found that passing an unmodified, freshly-read `FileSecurity` to `SetAccessControl` can be a no-op. The adapter now explicitly applies the access section and verifies it before touching the source name. Complete NTFS owner/SACL/attribute preservation is not claimed; that policy needs production qualification.
- Generated candidates, intent records and backups are retained in private test folders. There is no production cleanup/retention policy yet, and no recursive deletion is used.

## Validation

Both isolated Debug and Release builds pass with zero errors/warnings. Builds use `bin/EngineStage2d-Debug` and `bin/EngineStage2d-Release`; logs and disposable workbooks remain under ignored `obj/EngineStage2d-*.log` and GUID-named private directories.

- Per configuration: safety 72, deadlines 6, native result grid 16, security 12, negative projection 10, value edit 62 and candidate-save 54 assertions pass.
- Terminal-publication safety: **76 assertions** pass per configuration.
- Generated native fixtures: all **12 engine/format/mode combinations** pass per configuration (38 publication assertions plus 12 cross-engine output comparisons).
- AGL XLSB Debug and XLSM Release: both engine directions pass source-copy replacement, original retention and other-engine full-calculation parity over **45,586 positions** (8 publication assertions plus 2 output comparisons per configuration). The generic source-unchanged assertion also passes in every mode.
- Normal Debug/Release executables, repository Blank/Demo and original Sandbox AGL SHA256 hashes remain unchanged. Pre-existing Excel PIDs remain present; no owned trial Excel process remains.

Final single-run hand-off measurements, from starting publication through native close, staging/hash/security checks, guarded renames and completion record, excluding candidate creation/verification and subsequent independent reopen:

| Private format/build | DevExpress hand-off | Excel hand-off |
|---|---:|---:|
| AGL XLSB Debug | 152 ms | 1,825 ms |
| AGL XLSM Release | 154 ms | 1,896 ms |

These are diagnostic development-machine samples, **not total Save latency**, benchmark medians or Stage 3 structural timings. Frequent production Save must not require this terminal close/reopen path without further lifecycle design and measurement.

The safety matrix covers opt-in, revision currency, overlapping edits/publication, destination/format validation, partial staging failure, source changes/removal, locked source, late competing writers, rollback, cancellation, acknowledgement failure, native disposal failure/deadline, security streams, hardlinks, read-only files and custom ACLs.

Native tests publish generated XLSX/XLSM/XLSB with both engines through both modes, then reopen in the other engine and fully calculate the named-input formula and dynamic spill. The AGL test changes Funding Assumptions G82 by +300 on a **new private copy**, publishes over that copy only, and compares the established **45,586 output positions** after other-engine full calculation. The untouched supplied baseline and pre-existing Excel process IDs are checked.

## Remaining gates

1. Integrate current model/history XML, typed edit validation, undo/redo, dirty/recovery state and authoritative formatting with the existing change manager and UI. Opening-baseline XML preservation cannot replace current history serialization.
2. Add publication-journal inspection/recovery, reopen/rebase semantics and save acknowledgement tied to the exact live model revision. Terminal publication is not yet suitable for ordinary frequent Save.
3. Qualify storage locations, file attributes/owner/auditing, crashes/power loss, retention and permanently hung native calls. Do not silently choose a less guarded overwrite route on failure.
4. Execute the approved three-route complete structural comparison: DevExpress .NET, VB.NET-driven Excel, verified master VBA. This checkpoint contains no new insertion timings or structural parity claim.
5. Live UI and full Excel/VBA event/menu/form round trips remain required before enabling automatic Excel preference for clients.

## API references

- Microsoft [CreateFileW sharing rules](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilew).
- Microsoft [SetFileInformationByHandle](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-setfileinformationbyhandle) and [FILE_RENAME_INFO](https://learn.microsoft.com/en-us/windows/win32/api/winbase/ns-winbase-file_rename_info), used with replacement disabled.
- Microsoft [ReplaceFileW](https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-replacefilew) documents multi-file failure states. This trial does not treat a preflight hash plus a later path-based replace as a compare-and-swap primitive.
