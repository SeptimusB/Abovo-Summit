# Analyser Balance Sheet diagnosis

Date: 21 September 2026. Diagnosis only; no product-code or source-workbook changes.

## Scope and verification

- Inspected `Library/Demo BP v26_0001.xlsb` and the current Debug auto-open file `C:/Sandbox/BP v26_0001 - New Blank.xlsb`.
- Source inspection used DevExpress's native XLSB reader in manual calculation mode. Recalculation checks used private, unsaved copies through Summit's normal model-open and `CalculateDependencySensitiveFile(..., True)` paths.
- Original SHA-256 hashes were unchanged. The running user Debug process was not touched.
- Compared Transactional DB rows 1791:1849 against Financial Position - Trad View rows 9:67. Row labels identify equivalent lines. Opening balances map O to C; forty forecast years map Q:BD to D:AQ.
- Each file has 46 numeric lines (including totals and Check), each with 41 values: 1,886 comparisons per file, both after open and after a full dependency rebuild. No differences above 0.001 workbook units and no nonnumeric mismatches were found. The populated Demo provides the nonzero evidence; the blank file provides zero-state coverage, not independent financial validation.

## Confirmed causes

1. All three analyser tabs share `Transactional_Records`, currently `Transactional DB!A6:BV1599`. See `BPIncomeExpenditureAnalyserV2.vb` lines 157-168 and 409-476. This is the transaction table, not the calculated Balance Sheet output.
2. The Balance Sheet filter is `[UseInBS] > 0`. Only worksheet row 98 and rows 133:163 qualify (32 rows). Their figures are zero, with no opening balances or populated BS grouping fields. Row 98 belongs to the cashflow placeholder section. None of the actual transaction rows qualify.
3. `OrderedBSGroup` (BM) and `OrderedBSHeading` (BN) have no populated data cells anywhere in the named range. The analyser explicitly groups by these fields (lines 339-345). Consequently its grid groups collapse into blank captions; the chart represents the same defect as `(Unclassified)`. These are missing source fields, not a datasource type-inference or refresh fault.
4. The worksheet contains a separate, working `SOFP Outputs - Transactional DB` section at row 1790, with the statement in rows 1791:1849. It lies outside `Transactional_Records` and is not bound to the analyser. It also does not have the analyser's BM/BN/UseInBS metadata.
5. The analyser guesses the opening-balance column as the column immediately before the first forecast period (`GetOpeningBalanceColumn`, line 3042). That is P in the transaction table. The calculated SOFP uses O for opening balances and P is a spacer. Simply expanding the named range would not repair this mismatch.
6. `TransactionalDBSnapshotManager` copies only `Transactional_Records` (lines 46-101), then creates differences for the selected period/opening columns. The calculated SOFP block and its opening balances are therefore not captured by the current snapshot. A Live-only binding change would leave Snapshot/Differences incorrect.

## Numerical cross-check

Demo first forecast year, values displayed in GBP thousands, agree in both sources:

| Item | Transactional DB | Trad View | Displayed value |
| --- | --- | --- | ---: |
| Housing properties | Q1793 | D11 | 819,753 |
| Total fixed assets | Q1802 | D20 | 862,168 |
| Total current assets | Q1811 | D29 | 27,878 |
| Total assets less current liabilities | Q1824 | D42 | 884,638 |
| Total financing and reserves | Q1847 | D65 | 884,638 |
| Check | Q1849 | D67 | 0 |

The underlying statement already rolls forward opening balances and signed cash/noncash movements. For example, Q1793 starts from O1793; R1793 starts from Q1793. Summing annual transaction amounts alone is not a closing Balance Sheet. Income and expenditure reserves have a separate SOCI rollforward formula and must not be treated as an ordinary SOFP-heading sum.

## Recommended repair, not yet implemented

- Give the Balance Sheet an explicit statement source/mapping, reading the workbook-calculated SOFP output and its existing section/line hierarchy for both figures and charts.
- Preserve meaningful drilldown by tracing a line's opening balance and signed cumulative transaction contributions. Reconcile each drill level to the workbook-owned headline. Do not present annual movements as closing balances.
- Extend snapshot/difference handling to include the Balance Sheet's captured figures and required opening/detail data. Do not reconstruct a historical snapshot using today's opening balances. Existing snapshots without those data need an explicit unavailable/recreate policy.
- Do not simply extend `Transactional_Records`: intervening report/check sections and mixed row semantics could cause double counting and disturb existing SOCI/CF/snapshot behaviour.
- Any named-range or snapshot-layout extension needs explicit approval and Excel/VBA round-trip validation. A UI-only statement adapter should not silently alter the XLSB schema.

## Follow-up validation

- Repeat the 1,886-cell reconciliation for nonzero populated files after edits and structural inserts, locating the SOFP block structurally rather than hardcoding row 1791.
- Verify zero lines remain visible, opening balances, negative liabilities, accumulated depreciation, reserves, totals and the Check line.
- Reconcile every drill level; preserve expansion/width state and chart/grid parity.
- Verify captured Snapshot remains frozen, Differences equals Live minus Snapshot, old snapshot handling, structural invalidation, Summit save/reopen, and Excel/VBA round trip.
- Previous chart trial tests proved chart-to-grid parity only. They did not establish Balance Sheet-to-workbook correctness; the earlier Unclassified note was a symptom of this incomplete feed.

Local diagnostic outputs (ignored obj directory):
- Demo: `obj/BalanceSheetAudit/42102eb5d73e4e4896ba12510374efe4/audit.txt`
- Current blank: `obj/BalanceSheetAudit/dcbf3c423c134d3780fe77a79380d71b/audit.txt`
