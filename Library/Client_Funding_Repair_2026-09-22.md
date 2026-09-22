# Client Funding workbook repair - 22 September 2026

Status: Ready to test. This is an explicitly authorised, separate-copy repair of one identified client workbook, not a general migration, financial certification or new application release. Application test version remains 2.61. No source master was modified.

## File identity and delivery

- Original: `D:/Downloads/Test BP v26_0001 - FormGenRemoved - PopInSummit - funding cols insert (1).xlsb`
- Original SHA-256: `CB3F42A2D8CD5BB073BF571AD4EA5E84495B8B698A6552A5C6714C1482E35AA8` (unchanged).
- Repaired copy: `D:/Downloads/Test BP v26_0001 - FormGenRemoved - PopInSummit - funding cols insert (1) - repaired 2026-09-22.xlsb`
- Repaired SHA-256: `3796F6B858DC5C45E5FC0278350650EA27C5FF6FD9D899D450095F0A18E63243`; size 7,747,298 bytes.
- Companion report: same repaired basename plus `_report.md`.

The client workbook is the inspected 26.0001 base with existing client changes. It is not interchangeable with the separate AGL 26.0005 performance-test workbook. Its 339 VBA modules were compared with its own original, not replaced by the 337-module repository master project.

## Repairs made

1. Repaired 15,124 existing Funding-related formulas, including 820 ordinary-loan Transactional DB mirror formulas. Each Funding patch was proved through DevExpress's native formula parser to alter only one of the known insertion-reference signatures; operators, literals, worksheet spans, rows and reference anchoring were retained. The eight already-inserted ordinary-loan columns O:V remain, and the four revolvers remain W:Z. No client input columns were discarded.
2. Extended six undersized names from ten to eighteen ordinary loans: `LoanDescsOrd`, `Rep_Fund_04`, `Rep_Fund_05`, `Rep_Fund_09`, `Rep_Fund_10`, `Rep_Fund_17`.
3. Corrected explicit Funding source ranges inside the ordinary-loan mirrors from E:N to E:V before invoking the existing Transactional DB synchroniser. `TransCopy_LoanDescsOrd_A` and `_B` now each contain eighteen records plus their footer (nineteen rows). The nine existing facility mirrors remain twenty-two records plus footer. Client-specific expansions elsewhere are retained; TDB formulas were matched by named block and relative record, never by assuming master absolute rows.
4. With separate user approval, restored the already-erased `Multivariable Dashboard!B41` description formula. The intact Blank, Demo and earlier baseline formulas agree. The existing application save guard nests the long CONCATENATE call without changing its order or separators, avoiding the previously reproduced XLSB export loss. B41 calculates without error in Summit and Excel. It currently displays blank because the existing scenario does not activate the formula's `MAX(...)>8` branch.

The total is 15,125 explicitly repaired existing formula cells including B41, plus formulas populated into the sixteen mirror rows by the established synchroniser. New sheets, hidden metadata, capacities and schema migrations were not introduced. No VBA was executed or rewritten. Other existing dashboard indicators, including a no-result `#N/A` tile, were not redefined.

## Validation completed

- Source and both authoritative master file hashes remain unchanged.
- All 714,535 original formula cells were checked against the precise approved repairs and expected TDB row relocation: zero unexplained differences.
- All 87,848 original constant/input cells are unchanged. Existing non-empty cells retain number formats and locked flags; worksheet order, visibility and protection states are retained. Global names are checked, including sensitive definitions without disclosing their contents; name count remains 1,761 across 285 worksheets.
- The 820 patched single-cell legacy array formulas are authored using ArrayFormulaInvariant, rather than flattened. Original multi-cell/dynamic arrays are rejected by the repair preflight. Non-TDB array classifications are checked during preservation validation; wholesale TDB array classification is not asserted because native import/export classification can differ.
- All 43 Check Sheet statuses are OK, with B63=0. The saved SOCI, financial position, detailed cashflow and Check Sheet contain no formula-error cells.
- A separate Excel instance opened both files read-only, with links, macros and events disabled. The repaired full dependency rebuild took 2.435 seconds in this run; no circular-reference address was returned. This is a diagnostic, not a benchmark guarantee.
- Summit's saved values and Excel's full-recalculated values agree on all four output sheets above, with absolute numeric tolerance 1e-7.
- Excel saved a separate verification copy; native reopening retained worksheet order, global-name count, and 12,483 selected formulas and values, including the restored dashboard formula.
- Unsaved first/last-added-loan probes independently exercise O and V: mirror A returns the injected 5 / 8 totals; mirror B returns the injected 123 / 456 drawdowns. Those probes never save the workbook.
- All 339 VBA module identities and source hashes are identical through original -> Summit repair -> Excel copy. Aggregate source hash: `c37706c7b1432e8661f4db7c2098c1e7bdd60c16338b68c3b77cf8b5da87414b`. The binary VBA container is rewritten by the serializers; byte-for-byte binary identity is not claimed.
- Original ZIP package parts are retained except `xl/calcChain.bin`, which Excel can rebuild. All ten media and nine custom-XML package entries remain. File size reduction reflects native reserialization; it is not evidence of removed worksheets or VBA. Funding and dashboard before/after native renders were reviewed.
- Both current Debug and Release builds passed. This task changed only offline repair/verification tools and evidence documentation, not production application behaviour or the test-release number.

## Safety and remaining acceptance

The repair is locked to the original SHA-256 and fails on unexpected reference signatures or changed input formulas. Bulk changes use the actual model safety and worksheet-protection services, then existing mirror synchronisation and model save services. Calculation mode, engine, history and entry worksheet protection are restored. This repair must not run automatically on model open.

Automated checks do not replace accountant review or interactive Excel/VBA acceptance. On a further disposable copy, the client should open the repaired file in current Summit, inspect existing Funding assumptions, add/delete a permitted loan, save, reopen in Excel, exercise the relevant VBA action, then reopen in Summit and recheck outputs. Exercise Stress Test scenario description generation, including more than eight descriptions. Financial interpretation and unsupported bespoke workflows have not been certified. Keep the original until client acceptance.

## Reproduction and evidence

One-off native tooling: `Tools/ClientFundingRepair.cs`, `Tools/ClientFundingRepairSheets.txt`, `Tools/Inspect-ClientFundingRepair.ps1`. The 32-sheet list is frozen for this known repair; it is not an inferred universal Funding contract. `inspect` compares against an independently Excel-expanded matching master; `prepare` proves its resulting manifest; `repair -RestoreDashboard` produces only a private copy. Do not use the post-save final manifest as a new Funding-only preflight manifest.

Validation tooling: `Tools/Validate-ClientFundingRepair.ps1`, `Tools/ValidateClientFundingRepair.cs`, `Tools/Check-ClientFundingExcel.ps1`, `Tools/Test-ClientFundingBoundaries.ps1`, and the existing `Test-SaveExcelRoundtrip.ps1`, `Inspect-SaveRoundtrip.ps1`, `Verify-VbaModuleHashes.py`.

Local disposable evidence (not required at application runtime and not committed as workbook payloads):

- Independently Excel-expanded master: `obj/FundingStructureTests/ba0f49a3374e42029718b64d6f716cb6/excel-grouped-reference.xlsb`.
- Funding-only preflight inventory: `obj/ClientFundingRepair/91cf2f26ae5c48bd99597c2d586e8896/inventory.json`.
- Final private candidate, manifest and render/check evidence: `obj/ClientFundingRepair/58bfaea459fa48d1bb5af7142371785c/`.
- Final preservation report: `obj/ClientFundingValidation/5b5174083ac241429a602713e9b9dcdf/preservation.json`.
- Excel rebuild verification copy: `obj/ClientFundingExcel/0abb4ae2e7f742e8aec13781110c0e80/excel-recalculated.xlsb`.
- Final logs: `obj/client-funding-dashboard-{final,preservation,excel,parity,native,boundaries,vba}.log`.

Early private trials correctly failed on undersized explicit TDB source references and legacy-array preflight; those candidates were superseded and never delivered. The complete Funding-only repair passed before the separately authorised dashboard restoration; the combined final copy was then validated again.
