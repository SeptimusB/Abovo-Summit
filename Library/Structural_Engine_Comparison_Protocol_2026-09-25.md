# Stage 3: native .NET versus VBA structural comparison

Approved by the user on 25 September 2026: compare **all three** routes below as part of the Excel-first implementation. This is the agreed experiment design, not measured performance or completed Stage 3 integration.

| Route | Workbook owner and command implementation | Purpose |
|---|---|---|
| DevExpress .NET | Existing Summit VB.NET structure services operating on a DevExpress workbook | Preserve the no-Excel route and measure the corrected production commands |
| VB.NET driving Excel | Summit's managed code calls the owned Excel object model directly, without invoking the insertion macro | Measure Excel's native range/formula fix-up with explicit managed orchestration |
| Excel/VBA | Owned Excel invokes the corresponding verified master VBA business action | Compare with the established Excel workflow and avoid per-cell COM round trips |

The second route is not a third calculation engine: Excel still owns and calculates the workbook. .NET drives its operations. Prefer bounded/batched range calls, not a cell-by-cell transcription of VBA. No third-party engine is introduced by this comparison.

## Common inputs and action mapping

- Begin with a fresh private copy of the reviewed original AGL XLSB for each route/run: SHA256 `30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47`. Verify it again before running. Never overwrite the customer original or a master.
- Inspect the current authoritative Blank/Demo action mapping and compare relevant VBA module identities/hashes with the client baseline. Do not assume a similarly named procedure or a model version proves equivalence. Record any bespoke behavior or intentional difference before comparing results. Raw VBA/passwords remain out of source control.
- Primary cases: **ten Funding columns** and **ten Development records**, matching the client's previous trials. Add single-record, end-of-range, add/delete reversal and protected-boundary cases. Use larger counts only as a clearly labelled scalability trial, not a substitute for the ten-record result.
- Compare complete business commands: all linked sheets, grouped inserts, templates, formula/name fix-up, Transactional DB mirrors, clearing/defaults, protection restoration and the intended calculation point. A raw `Rows.Insert` or `Columns.Insert` microbenchmark is not an equivalent command.
- Preserve the first ten ordinary loans and revolver boundary in Funding deletion tests. Reuse the repaired structure rules and remaining established template/deletion guards; do not relax them to win a benchmark.
- Match macro security to existing user policy. If the actual VBA action cannot run, report that route unavailable. Do not change Trust Center, remove downloaded-file markers or invoke unverified macros.

## Time accounting

Record cold opening separately. Run the routes sequentially, with the same workbook/calculation state, logging session ID, engine/build/bitness, input hash, action/count, route, result and stage durations.

Measure the total user-visible operation and these components where observable:

1. Dispatch/setup and transaction/checkpoint cost.
2. Native insertion/deletion and template copying.
3. Related formulas, defined names, linked sheets and Transactional DB synchronization.
4. Full calculation/rebuild at the same intended validation boundary.
5. Equivalent integrity/round-trip validation (reported separately, never silently removed from one route).
6. UI value/format read-back and refresh.
7. Save plus reopen/read-back, separately from an in-memory command.

VBA without an approved timing hook gives an aggregate macro time; label internal stages unavailable rather than guessing them or subtracting unrelated runs. If an instrumented wrapper/copy is needed, keep it disposable and demonstrate that it preserves the macro's behavior. Do not modify the master solely to obtain timings.

After a correctness pass, collect at least three uncontended timings per route and report median/range, with cold and warm runs distinct. Record current high-spec machine results as such; client minimum-spec/VM validation follows. Distinguish in-process manipulation time from the full user wait. A failed correctness run is not a recommended fast route.

## Acceptance checks

- Expected record counts, positions, inserted defaults and cleared input cells; adjacent populated records preserved.
- Correct formula references including 3-D/grouped-sheet formulas, names and all linked/mirror ranges. Compare formula structure independently of cached numbers.
- Legacy/dynamic array anchors and spill geometry, not just formula text or today's single-cell results.
- Original fill/pattern and lock behavior, data validation, number formats and worksheet protection restoration.
- Full-calculated financial outputs and Check Sheet deltas against the matching baseline/oracle. Existing unrelated failures and expected accounting imbalances stay separately identified.
- VBA and Summit custom XML/history retained appropriately. Compare payloads semantically where timestamps/history legitimately differ; do not dismiss unexplained differences.
- Save to a private candidate, normal Excel/VBA and DevExpress reopen, then repeat relevant checks. No silent repairs, overwritten customer files, hidden migrations or stripped metadata.
- Failed/cancelled/timeout action must not publish a success or adopt a half-updated workbook. Retain or restore the verified checkpoint using the established transaction path.

## Decision record

Select routes **per complete command**, based on correctness, total latency, security availability and maintainability. Excel may calculate while the implementation of a structural command is managed code or VBA, but there must still be one authoritative workbook session. Do not alternate engines on a dirty workbook without an explicit verified handover. DevExpress remains the configured/open-time fallback. Publish code, tests and technical summaries only; comparison workbooks and attachments remain local.
