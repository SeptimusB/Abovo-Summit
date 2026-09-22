# Idle integrity trial - Summit 2.67

Status: Ready to test. Engineering checks below are not accountant, client-machine or interactive Excel/VBA acceptance.

## User contract

Options now has an **Integrity** tab: **Check integrity every [X] minutes if idle**. It is initially disabled, defaults to 30 minutes and accepts 1-120 minutes. Settings persist per Windows user using the same application-settings mechanism as recovery backup.

The due interval is a minimum interval, not a deadline. Before each work unit, the scheduler requires at least two minutes without keyboard/mouse input in the current Windows session. Activity in another application also postpones checks. The interval is reset after a completed/failed pass or a changed setting. Opening a model schedules its first check; closing removes its state.

**Never interrupt a running calculation, rebuild or refresh.** The unit finishes, restores engine/mode/deferred-sheet settings and refreshes visible result consumers before yielding. Input only prevents the next unit from starting. There is no cancellation inside workbook operations, `DoEvents`, worker-thread workbook access, automatic structural repair, file save, external-link refresh or VBA execution in this feature. It does not make the remaining long calculation nonblocking: returning during it can still mean a wait.

Input while a unit runs is observed on the next dispatch using Windows session input timestamps, not solely UI messages (which can be queued behind calculation). Failed input observation defers work. Tick wrap and out-of-order/future timestamps are treated conservatively. A subsequent unit requires two minutes idle again.

Pending native editors, modal dialogs, normal/recovery progress operations, saves, grouped edits or structural operations block work. Recovery-required/authoring-preview models are not checked. A live, non-minimised notice owner must exist. Models are visited round-robin, at most one unit per timer dispatch. Calculation/progress work stays on the owning UI thread.

## Stages and boundaries

1. Reuse the model's pending/full-rebuild flags. If results are pending, call the existing guarded full calculation (full dependency rebuild only when required), restore the prior state, then refresh the active UI consumers. An already-current revision does not pay for repeated full calculation. The existing value-save policy and XML `IsCalculated` route are unchanged.
2. Read the current Check Sheet through the same BP/DSA validation logic as close, without another calculation. A completed, same-revision Check Sheet failure remains known to the existing close guard.
3. Inspect global/local defined names for broken references and existing XLSB export compatibility hazards. Quoted string literals containing `#REF!` are not broken references. At most 100 names/25 ms target per batch.
4. Inspect supported Transactional DB mirror geometry using the synchroniser's existing compatibility rules, including matching legacy row mirrors such as Journals. It never invokes the synchroniser's resizing or formula-writing paths. This is not exhaustive validation of every bespoke mirror/formula or 3-D worksheet block.
5. Inspect cached errors and XLSB export compatibility across existing cells, at most 5,000 cells/25 ms target per batch. One native API call is indivisible, so 25 ms is a budget, not a hard latency guarantee.
6. Publish the completed revision, counts and sampled locations to System Messages, with a brief completion notice. Samples are bounded to ten per category/100 total; totals count every finding. No automatic file/report export is performed. Existing System Messages copy/export facilities remain available.

A partial pass holds bounded counters and iterators, not a workbook clone. If the model's calculation revision changes, the partial pass is discarded before accessing its iterators again. It cannot certify a mixture of revisions. An exception records **incomplete**, retains safety flags and retries at the next interval. Disabling checks discards a paused pass. A previous report is historical; its text explicitly says that later edits invalidate it.

Only literal `=NA()` sentinel errors are separated from findings. Other N/A values may be intentional chart gaps, but are retained for review rather than globally ignored. No scan result claims that all business logic, client bespokes, VBA or financial solvency have been validated. Current cached results are inspected; unchanged workbooks are not forcibly recalculated just to reevaluate time-sensitive volatile formulas each interval.

## Implementation

- `Services/IdleIntegrityManager.vb`: persisted options, session idle clock, safe scheduler, revision-local staged inspection and bounded reporting.
- `Services/FileManager.vb`: open/close tracking, shared calculation scope, factored cached Check Sheet validation and existing close-failure integration.
- `Services/TransactionalDB/TransactionalDBSynchroniser.vb`: read-only mirror sizing inspection reusing its compatibility rules.
- `Services/RecoveryBackupManager.vb`: shared editor/owner guards and explicit exclusion while an integrity unit is active. Existing recovery scheduling policy is retained.
- `Interface/User Interface/ApplicationOptionsForm.vb`: third native options tab, bounds, Apply/Cancel and clear blocking/repair limitations.
- `Tools/Test-IdleIntegrity.ps1` / `IdleIntegrityFixture.cs`: reproducible x86 native tests on synthetic and private master copies.

No workbook schema, structure XML, original master/client file, formula, name or VBA code is changed by this trial.

## Verified engineering evidence

Debug and Release builds passed. Synthetic native tests passed in both configurations: disabled/not-due/input gating, notification owner, save/modal/grouped/structural/unposted-editor/recovery guards, atomic calculation completion, no next stage after input, two-minute resume requirement, changed-revision restart, repeated-current calculation skip, disable/close cleanup and calculation-failure state restoration. The Integrity page was rendered and visually checked at ordinary test DPI; cancelling edited settings did not apply them.

Release Demo and Debug Blank were opened as private copies, given a typed/journalled stock edit, inspected, and checked for unchanged formulas, constants, names, worksheet order, protection, dirty state and native history. ChangeManager history survived, and Undo/Redo restored/reapplied the input. Engine, calculation mode and deferred-sheet settings were restored. Both master source hashes remain unchanged:

- Blank: `0B4C06800FE998E8733A5D8EB9CDEFA04F517750275BB0CFDD532928900F5B79`
- Demo: `1ED79726D8D1129C699242C9E8B3E988AC920EA7108D3C0B105BF9B0BC8A5E9C`

Final runs checked 1,210,966 existing cells and 1,761 names per model with no Check Sheet or supported mirror-geometry findings. Demo reported 467 N/A, 563 VALUE errors and one existing XLSB normalization candidate; Blank reported 475 N/A, 563 VALUE errors and the same normalization candidate. Each separately counted 98 literal NA() sentinels. Samples include `Hidden - BP Menu Sheet`, `OW - Charts Source Data` and the previously known `Multivariable Dashboard!B41` normalization. These are observed native-calculation findings, not 1,031/1,039 demonstrated workbook corruptions. Classification against Excel/VBA and chart intent is still outstanding.

Observed initial calculation/refresh units: Demo 7.112 seconds, Blank 6.369 seconds. Total active fixture time was about 10.6/9.9 seconds over 290 units. The deterministic fixture advances units immediately; production adds 150 ms timer spacing, notices, idle waits and any user activity. These are not minimum-spec or controlled bitness benchmarks.

The first preservation test found native history 1 -> 0. A probe established that this occurred during delayed **file-open UI dispatch before the first integrity unit**, not during checking. Final tests drained that dispatch and then made a real journalled edit: native history remained 1 -> 1 through the check. The existing recovery regression and Debug save/calculation-state regression also passed.

## Manual acceptance still required

- Enable at a short trial interval; leave Summit idle for two minutes and confirm start/completion and System Messages.
- Resume input during the calculation and between scan batches; the calculation must finish, while subsequent inspection waits for another two idle minutes. Check the actual Windows input path, not just injected test timing.
- Leave a native grid/date/combo editor unposted; test multiple open models, minimised windows and overlap with recovery/save/structural actions.
- Check populated client models with visible Analyser Live/Snapshot/Differences, normal grids and pending result-reader gates. Confirm displayed results, journal/Undo and post-check Save/Excel/VBA/Summit round-tripping on copies.
- Review reported chart N/A and native menu/formula errors with Abovo before defining any additional suppression/baseline rules.
- Measure UI-blocked calculation time on the agreed 32 GB/Core-i5-class baseline; check options wrapping at client DPI.

## Reference boundary

Windows idle observation follows [Microsoft GetLastInputInfo documentation](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getlastinputinfo), including its session scope and non-monotonic timestamp warning. DevExpress Support search surfaced its [non-UI-thread workbook discussion](https://supportcenter.devexpress.com/ticket/details/t326464/reading-data-from-control-in-a-non-ui-thread), but the opened page did not expose the full answer, so no threading capability was inferred from it. This implementation deliberately retains the locally verified 25.2.4 UI-thread calculation service rather than introducing background workbook access.
