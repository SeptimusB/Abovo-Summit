# Excel / DevExpress checkpoint 2f: current history in verified saves

## Scope and authority

This is an isolated continuation, not a client release or a production engine switch. Normal Debug/Release **2.98 remain unchanged**. Candidate saving and publication remain explicitly opted in. The original AGL file, repository Blank/Demo, local attachments and licences are not changed or published.

The existing `ModelChangeManagerV2` now captures immutable **display-history evidence** for a particular opening file hash/path, engine session/revision, manager identity, user revision and calculation revision. It uses the existing recovery-history serializer: newest 1,000 rows, bounded fields, no executable Undo commands. It is not a second edit/history stack. Restored history remains read-only evidence; this work does not add Undo across process restarts.

## Implemented boundaries

1. `ModelChangeManagerV2.CaptureSaveHistory(result)` runs on its owning thread, requires the same open model/workbook with no active grouped/bulk edit, and verifies that history and engine results share the same opening file bytes. The returned snapshot cannot change when later edits arrive.
2. `WorkbookCalculationSession.CreateSaveCandidateWithHistoryAsync(...)` explicitly requests history inclusion. It rejects mismatched source paths, sessions or engine revisions before exporting. The existing no-history API is unchanged.
3. Native export first passes the existing value-only custom-XML/VBA/array preservation gate. Only then does `WorkbookHistoryPackage` update the owned `urn:abovo:summit:recovery-history:1` payload. First creation includes standard custom-XML properties/relationships/content types; subsequent saves replace that same payload. Both `workbook.bin.rels` (XLSB) and `workbook.xml.rels` (XLSM/XLSX) are supported. Duplicate/unsupported history rejects the candidate.
4. Before/after decompressed-entry hashes enforce the exact permitted ZIP delta. All other parts must remain byte-identical during the history write. The two shared relationship/content-type documents must be unchanged except the exact required additions. Namespace prefix/serialization declarations are normalized for semantic comparison, not field values or meaningful whitespace. Native cached-value/definition read-back and critical-package verification still follow.
5. Candidate and publication receipts carry the captured history ID/hash. `IsSavedHistoryCurrent(snapshot, receipt)` also checks the live manager, source path, user/calculation revisions and current serialized history. It returns false for newer edits or history-only changes. It **does not clear dirty state**, even when true.

The history writer is shared with existing recovery saving, not copied into an engine-specific implementation. Its document serializer was extracted without changing row/field semantics. It now avoids adding XML indentation and refuses an unknown existing history version. No new worksheet, hidden template, VBA change or load-time migration is introduced.

## Important limits

- A matching history receipt is necessary evidence, **not proof that the engine owns every live edit, format, name or structure change**. Live change routing and dirty acknowledgement are still unconnected. Test fixtures use a small real change-manager workbook on a dedicated STA and explicitly send corresponding edits to the native engine; they do not exercise the live DIT dispatch path.
- This stage updates only recovery/history XML. It preserves other XML. Do not serialize an unsaved Structure Manager draft automatically. Funding schedule XML must remain coupled to its associated workbook names and structural transaction; adding that XML alone would be incorrect.
- Publication still closes the owner. Receipt-based reopen remains available for terminal hand-off testing, not a proposed implementation of every frequent Save. Avoiding repeated close/open, source-baseline renewal and authoritative live state ownership are the next implementation gates.
- The history delta audit is intentionally conservative: at most 20,000 ZIP entries / 2 GiB expanded content, with streamed hashing, bounded history XML and existing package limits. These whole-package validation passes are trial verification costs, not a fast-save performance claim. Any optimized validation policy must retain equivalent coverage.
- No automatic startup scan, storage-policy expansion, hanging-Excel termination, financial certification, client UI change or full Excel/VBA menu/event acceptance is implied.

## Validation

- Isolated main and harness Debug/Release builds pass with zero errors/warnings: `bin/EngineStage2f-Debug` and `bin/EngineStage2f-Release`. Normal executable hashes remain unchanged.
- **39 targeted history assertions per configuration:** real typed edits and existing Undo/dirty state, owner-thread/closing/grouped-edit rejection, wrong source/path/session/revision, cancellation, immutable snapshots while live history advances, history-only changes, receipts without history, newer edits after publication, bounded recovery import, first creation/replacement in three formats, unsupported versions/duplicates and safe candidate cleanup. Recovered rows have negative IDs and no executable Undo action.
- Existing regressions pass per configuration: safety 72, deadlines 6, native-grid 16, security 12, negative projection 10, value edits 62, candidate save 54, terminal publication 76 and recovery/reopen 109. Generic unchanged-source checks also pass.
- Generated XLSM/XLSB with custom XML, a named input/formula and dynamic spill: **29 assertions and eight independent full-output comparisons per configuration**, covering two successive saves through both engines and other-engine full-calculation reopen.
- AGL XLSB Debug and XLSM Release: **15 assertions and four independent 45,586-position comparisons per configuration** pass, including actual Funding G82 history and both current edits after the second save.
- Original AGL, repository Blank/Demo and normal Debug/Release hashes are unchanged. All native runs confirm that pre-existing Excel processes remain and their own processes exit. Final source-diff whitespace and project-index JSON checks pass.

Pilot corrections were confined to this checkpoint: XML auto-indentation initially caused an over-strict delta comparison; the writer now adds no formatting whitespace, preserving meaningful text. The test's range address (`A1:A1`) was corrected to the single-cell address expected by the real change manager. An earlier AGL pilot compared correct engine outputs but its small history fixture used `Data!A1`; the final tests explicitly use and verify Funding G82. These pilot runs are not counted as additional final assertions or performance evidence.

The native workflow performs two successive +100 edits, captures the real change-manager history, publishes under private new names, reopens with the same engine, and independently full-calculates with the other engine. The AGL input is Funding Assumptions G82. The existing five bounded financial/Check Sheet/custom-function rectangles contain 45,586 positions. History must name the actual edited cell, contain both edits after the second save, and remain in one owned XML part.

All generated workbooks and raw outputs are local ignored evidence. Only code, tests and technical notes are eligible for the code-only publication branch.

Observed two-save timings (seconds, individual-run ranges, not controlled medians):

| Private model | Engine | Native export | Candidate verification, including history/package audit and native cached read-back |
| --- | --- | ---: | ---: |
| XLSB, Debug | DevExpress | 23.985–24.113 | 13.800–13.825 |
| XLSB, Debug | Excel | 4.451–4.726 | 8.154–8.436 |
| XLSM, Release | DevExpress | 9.728–10.066 | 37.089–38.033 |
| XLSM, Release | Excel | 6.801–7.174 | 16.612–17.055 |

These exclude initial open/edit calculation, terminal publication, receipt reopen and independent other-engine full calculation. They do not measure the incremental cost of history alone or the future live Save policy. Full trial read-back is expensive; keeping an owner open and qualifying an appropriately bounded production verification policy remain performance work, not delivered speed improvements.

## Next gates

1. Connect authoritative input changes/Undo/Redo and current related model XML through the existing change manager, with exact saved-revision acknowledgement and an ordinary-save lifecycle that keeps a usable owner.
2. Implement live result/format refresh and configurable engine preference without changing fill-based editability or Excel security policy.
3. Complete recovery UX, storage/retention policy and native-process supervision, then the approved Stage 3 complete Funding/Development command comparison: DevExpress .NET, VB.NET-driven Excel, verified master VBA.
4. Run production UI and Excel/VBA round trips and Jon/Alex acceptance before switching client builds.
