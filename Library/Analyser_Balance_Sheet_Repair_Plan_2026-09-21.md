# Analyser Balance Sheet repair plan

21 September 2026. **Implemented as test-release 2.51 following approval of this plan.** Financial/client acceptance remains pending. See `Balance_Sheet_Funding_Navigation_Trial_2026-09-21.md` for implementation boundaries, actual validation and remaining tests. The detailed plan below records the design and acceptance gates, not a claim that every manual gate has passed.

## Decision

No extra XLSB worksheets are required. Read the already-calculated SOFP output through a separate, read-only Balance Sheet adapter. Keep SOCI/Cashflow on their existing transaction feed. Persistent Balance Sheet snapshots use additional explicitly defined regions on the existing `TDB Snapshot` and `TDB Comparison` worksheets, created only by the user's snapshot command, not by opening a file. Native XLSB/Excel persistence has been checked on private results; interactive VBA execution and financial acceptance remain required.

Do not extend `Transactional_Records` over the intervening SOCI, report or check blocks. Do not fill its unfinished BM/BN/UseInBS fields merely to make the current grid display something: transactions are annual movements, whereas the Balance Sheet is an opening balance plus accumulated signed movements.

## Source findings and evidence

Read-only native XLSB inspection and private, unsaved Summit copies of:

- `Library/Demo BP v26_0001.xlsb` (nonzero populated evidence).
- `C:/Sandbox/BP v26_0001 - New Blank.xlsb` (current Debug auto-open; zero-state evidence).

The originals were not saved; audit wrappers verified unchanged SHA-256 hashes. See `Analyser_Balance_Sheet_Diagnosis_2026-09-21.md` for the missing BS source metadata and the initial 1,886-value comparison per file, both before and after dependency rebuild.

In these specific files, the output title is at `Transactional DB!A1790`, the statement occupies rows 1791:1849, labels are in B/H, opening balances in **O**, and the forty years in Q:BD. It reconciles with `Financial Position - Trad View!C9:AQ67`. P is a spacer, not an opening balance. These addresses are audit evidence, **not production constants**: structural insertion moves the output block.

The follow-up reconstructs every one of the 37 primitive statement lines from its opening balance and the exact transaction criteria/signs below across forty years (1,480 annual values per file). Maximum absolute deviation was 6.984919309616089e-10 workbook units for Demo and zero for Blank. This establishes drilldown feasibility for these files; it is not validation of unknown client versions or bespoke layouts.

Local ignored evidence: `obj/BalanceSheetAudit/4375402f228b4a7791ecb4b43bb6f04f/audit.txt` (Demo) and `obj/BalanceSheetAudit/2f5a50437fd04fef86bf895722dbbc2f/audit.txt` (Blank). The diagnostic deliberately uses the inspected layout; production must validate and resolve a structure definition.

## What the transaction rows actually mean

The current formulas define these four families. Limits shown are those in the inspected files; derive them from a validated model definition/formula contract after insertions.

| Statement family | Annual movement to add to the previous closing balance | Workbook example |
| --- | --- | --- |
| Fixed/current assets and current liabilities, except Cash and Bank | For matching H (SOFP Heading): **minus Cash**, **plus Non Cash** in rows 164:1470; **minus all matching rows** in 1471:1599 | Q1793 housing properties |
| Creditors, provisions and reserves, except Income and Expenditure Reserve | For matching H: **plus Cash**, **minus Non Cash** in 164:1470; **plus all matching rows** in 1471:1599 | Q1827 long-term loans |
| Cash and Bank | **Plus every Cash row** (F = Cash) in 164:1599, regardless of H | Q1808 |
| Income and Expenditure Reserve | Plus rows 164:1615 with BF (SOCI ordinal) below BF13 and BQ (UseInSOCI) equal to 1, **plus** rows whose G matches G1784 (Movement to Pension Provisions) | Q1842 |

BF13 is currently 58, derived from `Hidden - SOCI Structure`. It must not become a literal 58 in production. Two matching reserve terms must retain the formula's two contributions, not silently deduplicate them. The reserve formulas extend sixteen rows beyond `Transactional_Records`; none of those extra rows qualified in these two files, but capture/validation must not assume that for another version.

The 37 openings link to Trad View column C, which in turn indexes `Financial Position - Alt View`; they are not contained in the transaction table. Examples: Trad C11 indexes Alt D10:D50 (housing), C45 indexes AI10:AI50 (loans), C60 indexes AW10:AW50 (reserve), and C26 indexes R10:R50 (cash). Never invent a split of these opening balances between today's transaction records.

Totals use the existing SOFP sum/reference formulas; the Check is a rounded difference. Do not derive category membership from captions alone, infer liability signs from financial terminology, or sum both a subtotal and its children.

## Proposed drilldown

```text
Statement section (e.g. Fixed Assets)
  Statement line (e.g. Housing Properties)
    Opening balance (separate, traceable to Trad/Alt View)
    Cumulative transaction contributions
      Workbook Level 1 -> Level 2 -> Level 3 -> Level 4
        Description / originating record
```

Skip empty or repeated hierarchy labels. Retain all defined statement headings, including zero-valued ones. Group the movements using K:N and O (with copy fields BT/BU checked against their source), using **only rows and signs actually selected by that statement formula**. `UseInBS > 0` is incomplete and must not determine this membership.

Examples verified in Demo:

- Housing has 85 matching transaction rows. `Component Non Cash Adjustments > Component Totals > Value / Cost on disposal - RTB/RTA > General Needs` comes from row 1053 (`TransCopy_HAComponents_H`). Row 817 also contributes land-bank capital expenditure through an INDEX into `Non Tenure Capital Expenditure`, with no TransCopy name covering that record.
- Long-term loans has 30 matching rows. `Cashflow detailed > Loan Repayments > [loan description]` uses `TransCopy_LoanDescsOrd_A`; drawdowns include `TransCopy_LoanDescsOrd_B` and direct Loan Drawdowns formula rows outside those named mirrors.
- Cash and Bank has 677 matching rows; its hierarchy legitimately includes receipts/payments across the model, not just rows labelled Cash and Bank.
- The reserve has 230 qualifying rows in Demo. The existing Level/Description metadata can explain them, but its separate SOCI/pension rules must remain visible in the trace.

For a closing-balance chart, each transaction's plotted contribution in year n is its signed **cumulative** amount through year n. The opening-balance series stays constant across the forecast. Their sum must equal the parent at every point. Label cumulative contributions clearly; optionally offer a separately labelled annual-movement inspection later. Do not switch chart semantics silently on drilldown.

Root headline values and visible numeric formatting remain workbook-owned. The adapter must share the same hierarchy/values between grid and chart. A statement total is a node, never an extra additive transaction. Keep a visible reconciliation diagnostic if detail does not explain a headline; do not manufacture a balancing record.

Stable identities need a statement-line definition plus source provenance. Prefer named mirror + record identity when available, but handle unmirrored formula rows too. Re-resolve after geometry changes; old absolute row numbers, nonunique descriptions or invented cross-version IDs are not safe keys. Unknown/ambiguous mappings should show the headline with drilldown unavailable and an explanatory message.

## Synchroniser impact

**Integration is needed, but not new mirror expansion rules for this repair.** The existing transaction rows already explain the balances.

`TransactionalDBSynchroniser.SynchroniseRules` already disconnects the analyser before shifts, checks snapshot validity, restores workbook mode/engine/history, and defers successful structural refresh. Reuse those hooks:

1. Disconnect/invalidate the new BS adapter as part of analyser disconnection, including its row-location and hierarchy caches.
2. On structural success, retain the current deferred-refresh behaviour so population is not slowed by an extra full calculation. On return/explicit refresh, resolve the moved output and transaction regions from the current workbook and rebuild once.
3. Extend snapshot validity/invalidation to the complete BS snapshot bundle as well as the original transaction geometry. This must work when no analyser exists, or it is retained but hidden.
4. On partial mutation/reconnect failure, never display stale BS values as current. Follow the existing error/recovery path and report missing/invalid mappings.
5. Cover structural entry points that do not reach this particular mirror batch, external Excel reconciliation, undo/redo and model reload with the same generation/invalidation contract. Do not rely exclusively on an open form receiving an event.

Keep the existing compatibility mappings, sizing/footer rules, formula propagation, structural-shift behaviour and Excel/VBA semantics unchanged. Much of the new work belongs in the **BS adapter and SnapshotManager**, not in synchroniser sizing code. A narrow synchroniser notification may be required if the existing disconnect/defer hooks cannot provide model-level invalidation.

## Snapshot and difference contract

- Capture an internally consistent bundle: current statement layout/identities, all opening balances, all 41 statement values, required transaction values and classification metadata, and the model-specific contribution rules/signs. Reuse the existing captured transaction table where it is sufficient; explicitly capture any additional contributors (including the reserve's wider source span).
- Store the added snapshot regions on the two existing dedicated worksheets without changing `Transactional_Records`. New names/layout are a proposed schema extension, not approved load-time migration. First check for conflicting client-owned content; no silent overwrites.
- Capture all regions in the same successful snapshot operation. Current creation clears both dedicated sheets and writes transaction/comparison data; extend that transaction coherently. A partial failure must invalidate all newly produced output, never publish a mixture of old and new.
- Snapshot must remain frozen after edits and save/reopen. Differences = independently evaluated Live contribution minus independently evaluated captured Snapshot contribution, including opening balances. Do **not** cumulatively sum raw differences using only today's classification: a record may have moved headings.
- Current geometry/header checks alone are insufficient for BS mapping validity. Check source addresses/dimensions, period headers, statement mapping/rule version and the required captured fields. Keep the established invalidation on genuine `Transactional_Records` structure change. Normal value edits do not invalidate a snapshot.
- Existing snapshots without BS data remain valid for SOCI/Cashflow if their existing checks pass. BS Snapshot/Differences should explicitly say "Balance Sheet was not included in this snapshot; create a new snapshot". Never substitute current openings or silently recapture.
- Preserve the existing low-cost invalidation approach (invalidate headers/markers rather than expensive cell-by-cell clears), extended to all bundle components. Capture clears/replaces the full bundle later. Save dirty state and entry protection state correctly.
- Round-trip test Summit save -> Excel/VBA open/calculate/save -> Summit reopen. Sheet order, VBA parts, original names, formulas and source master bytes must remain unchanged except explicitly approved snapshot additions in the test result.

## Implementation sequence / acceptance gates

1. **Define mapping and resolver**: explicit line IDs, section hierarchy, opening/period columns and four formula families. Prefer model structure definitions, with a strict validated adapter for existing v26 files; no generic formula interpreter or guessed caption matching. Add negative tests for moved, missing, duplicate, bespoke and unsupported sections. Unsupported files get a useful diagnostic.
2. **Live statement**: bind a separate read-only BS view to workbook SOFP outputs; preserve workbook formatting, full zero headings, clipboard/multiselect and UI state. Compare all 1,886 values to Trad on both inspected files. No extra whole-workbook calculation merely for chart/grid toggles.
3. **Drilldown**: build lazy cached contribution arrays after an established calculation, not during paint/cell edits. Reconcile every branch/year; test all four families, openings, reversed signs, duplicate captions, unmirrored rows, zero rows and reserve overlapping criteria. Test real edits, undo/redo and the user-selected populated client file.
4. **Persistence**: approve/test the additional regions on disposable files, implement atomic capture/invalidation and old-snapshot behaviour, then enable BS Snapshot/Differences. Test opening edits, classification changes with unchanged geometry, failed capture, persistent reopen and missing metadata.
5. **Structural and regression**: add/delete Cash Journals, Development, Funding and component records with analyser absent/hidden/visible, with/without snapshots. Verify shifted SOFP locator and no broken formulas/names. Compare Check Sheet and original SOCI/Cashflow; do not change authoritative workbook formulas to make tests pass.
6. **Performance and release**: measure initial adapter build, ordinary edit refresh, drill, source toggle, snapshot capture and structural add-lines separately, both x86/x64 on a populated file. Maintain the population-first/deferred-analysis policy. Debug/Release builds; Excel/VBA roundtrip; then client grid/chart, font/DPI and saved-state acceptance.

Scope guard: this plan does not author extra live worksheets, materialise replacement business calculations in hidden sheets, alter Excel/VBA formulas or widen transaction synchroniser rules. The read-only drill audit demonstrates a feasible explanation of the existing calculation, not permission to replace it.
