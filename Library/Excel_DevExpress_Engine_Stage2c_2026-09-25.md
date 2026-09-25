# Excel / DevExpress checkpoint 2c: guarded save candidates

## Status and authority

Isolated engine continuation of checkpoint 2b. Production Debug/Release **2.98 remains unchanged**. This is not a client test release or a live engine switch. The user approved staged Excel-first implementation with DevExpress fallback, existing Excel security, and code/tests/technical-notes-only checkpoint publication.

`EnableCandidateSaveTrial` is a separate, default-false option. `CreateSaveCandidateAsync` writes only a unique `~Summit-candidate` file in a generated child of a caller-supplied private test directory. It cannot replace the source, choose an existing filename, mark the model clean, or acknowledge a production Save. The source is still leased read-only; releasing that lease for an original-file replacement remains future work. There is no security-changing relocation of the original before macro approval.

## Implemented boundary

- One STA owns edits, native export, package preservation and native candidate reopen. Edits, external revision changes and concurrent exports are rejected while an export is active.
- Export requires a current immutable calculation result plus 1–64 current input/definition checkpoints. The session rechecks the input definitions and bounded output values before export and afterwards. No extra full calculation is silently added to candidate export.
- Excel uses `SaveCopyAs`; DevExpress uses a stream export in the original XLSB/XLSM/XLSX format. The existing reviewed XLSB long-`CONCATENATE` workaround is reused, with its temporary formula changes restored to the in-memory owner. Native worksheet order, names, code names and protection are compared during export/reopen. Chart-sheet candidates remain outside this initial trial.
- Candidate reopening disables macros/events/link updates and does not request calculation. Persisted values must agree with the captured results, including typed blanks, booleans and spreadsheet errors. Numbers allow only a small serialization tolerance (absolute 1e-9 or relative 1e-12); definition checkpoints retain exact typed input comparison.
- A receipt binds session ID, revision, source SHA256 and candidate SHA256. `IsCurrentCandidate` checks session currency; it is **not** a disk-integrity check. `ValidateSaveCandidateAsync` separately rechecks bytes and the security marker. Neither is an atomic publication API: a future replacement must keep its own revision/source/file guards through publication.
- Original `Zone.Identifier` content is retained and checked. Existing Excel policy is not relaxed. XML parsing prohibits DTDs, and mismatched file formats, duplicate ZIP entries, ambiguous duplicate XML payloads, missing critical relationships, package signatures and signed VBA projects fail closed.
- Cancellation removes the exact owned candidate when the native operation returns. A deadline quarantines the owner and rejects late success; a permanently hung native call still needs process-level supervision. Native export/reopen/close errors quarantine because bounded cell checkpoints cannot prove full native-state restoration. A rejected package/value comparison can leave an otherwise verified unchanged owner usable.

## Preservation findings and resolution

Native saves are not byte-preserving package copies. In the disposable fixtures, DevExpress renumbered custom XML parts and replaced schema-reference lists. Excel also regenerated companion properties for a DevExpress-created XML part. AGL's SharePoint schema XML additionally lost element-only line indentation. These are now distinguished from changes to payload content.

For this **value-only trial**, the engine exposes no XML, VBA or structure-edit API. The candidate step first matches every linked custom XML payload, ignoring namespace-prefix/attribute ordering and insignificant document/element-only line indentation. Leaf text, mixed content, processing instructions and `xml:space` content are retained. Changed/missing/ambiguous payloads are rejected. Only after a match does it restore the complete original XML bundle, companion properties and content-type overrides, repointing native relationships without changing their IDs. This preserves original IDs and schema declarations rather than accepting an exporter's replacements.

The original opaque `xl/vbaProject.bin` is also retained for this value-only path. DevExpress reserialized it even without an exposed code edit. No VBA source or password is extracted, stored or published. Dynamic-array metadata is **not transplanted or synthesized**: the candidate's same-format metadata must match the source. Final package verification and native reopen follow preservation.

This policy must **not** be reused for structural changes or live history/schedule XML mutation. Production must serialize current model XML, not restore an opening baseline over newer history. Workbook-event, menu/form and signed-project round trips remain separate acceptance work. Matching the tested calculation outputs is not proof of every VBA path or every workbook object.

## Validation

Builds and all harness outputs are isolated under `bin/EngineStage2c-Debug` / `bin/EngineStage2c-Release` and ignored `obj/EngineStage2c-*.log`. No customer workbook, licence, attachment or diagnostic package is published.

The final suite covers:

- Existing safety 72, deadlines 6, native-grid 16, security 12, negative projection 10 and value-edit 62 assertions per configuration.
- Candidate safety 54 assertions: opt-in, input limits, revision binding, edited/tampered candidates, exact owned-file cleanup, cancellation, concurrent edits, native failures, timeout/late cleanup, XML payload/link/property changes, DTD rejection, preserved whitespace, metadata corruption, format/signature rejection and downloaded-file marker retention/tampering.
- Generated XLSX/XLSM/XLSB: 37 assertions per configuration. Both engines export, reopen in the other engine, fully calculate a named-range formula and dynamic spill, then invalidate the old candidate by restoring the original input. Source bytes and pre-existing Excel processes are unchanged.
- AGL XLSB Debug and converted XLSM Release: Funding Assumptions G82 is changed from 2,700 to 3,000 in memory. Candidate cached outputs match the current owner. Each candidate is then opened in the **other engine**, fully calculated and compared across the established **45,586 output positions**. Both directions pass; restoration makes the old candidate stale. Required Excel VBA functions run only under existing policy. Original files remain unchanged; only owned Excel sessions close.

The generated fixture uses a supported constant dynamic array. An initial `SEQUENCE(3)` fixture returned `#NAME?` in the installed DevExpress setup before saving; this is not claimed as a supported cross-engine function.

The candidate verifier is deliberately expensive during qualification. Final single-run measurements in seconds:

| Format/build | DevExpress export | DevExpress verification | Excel export | Excel verification |
|---|---:|---:|---:|---:|
| XLSB Debug | 17.986 | 10.942 | 4.643 | 7.922 |
| XLSM Release | 8.673 | 31.113 | 7.663 | 18.390 |

These portions exclude initial opening/calculation, the candidate's initial source-state/package capture, separate cross-engine full-calculation acceptance, UI/history and final publication. They are diagnostic runs on the development PC, not repeat-run medians, live Save latency or the Stage 3 insertion comparison. Earlier runs varied materially, particularly XML reopen times. Do not ship a whole-workbook reopen check on every ordinary Save without evaluating a safe revision-based/occasional qualification policy.

## Next gates

Checkpoint 2d subsequently adds separately gated **terminal** publication and new-name Save As. See `Excel_DevExpress_Engine_Stage2d_2026-09-25.md` for its two-rename/recovery limits; the live Save and history gates below remain open.

1. Integrate existing `ModelChangeManagerV2` validation/history/dirty/recovery ownership and current XML serialization, not a second edit/history stack.
2. Add guarded source publication/Save As, external-change detection, backup and storage-failure handling. A candidate receipt is not permission to replace the original or clear newer dirty edits.
3. Provide process-level recovery for native calls that never return, and define source-lease/private-working-copy lifecycle.
4. Execute the approved complete-command three-route structural comparison: DevExpress .NET, VB.NET driving Excel, and verified master VBA. Value-only package preservation cannot be assumed valid after inserts/deletes.
5. Perform live UI and manual Excel/VBA round trips, including events, forms/menu, recovery/history and production file locking, before enabling automatic Excel preference for clients.

## API references

- Microsoft [`Workbook.SaveCopyAs`](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas) documents copy export without changing the open workbook.
- DevExpress [`Workbook.SaveDocument`](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Workbook.SaveDocument(System.String-DevExpress.Spreadsheet.DocumentFormat)); the implementation was compiled/tested against installed 25.2.4, not assumed from the documentation site's default version.
