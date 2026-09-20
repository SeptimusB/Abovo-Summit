# Stori migration: three-way evidence review

Date: 19 September 2026. Read-only investigation; no migrated result, recalculation, workbook save or production-code change.

## Executive conclusion

The four changes recorded in Stori's Version History are supported by the inspected workbook structure and formulas. The latest 26.0001 blanks do **not** contain the Stori-specific stock grouping, income-driven management costs or covenant definitions. The stock and management-cost capacity expansions must also be reproduced through the appropriate structural operations before population.

Version numbers are not reliable ancestry identifiers here. Stori and the later generic model use **25.0701 through 25.0704 for different changes**. The user confirmed that strict version management was not employed. A higher number is therefore not evidence that a bespoke change was incorporated.

The old base remains a strong provisional ancestor, not a proven byte-identical historical ancestor. Migration should use the three actual artifacts and a reviewed difference manifest. File names and version history help establish provenance; they cannot replace content comparison.

## Inputs and preservation

| Role | File | SHA-256 |
|---|---|---|
| Old generic base | `D:\Downloads\Master BP v25_0700 (1).xlsb` | `21E1E4B032821A9864046F1C1493819198417D2158D588019F626C9F55BA59B8` |
| Populated Stori | `D:\Downloads\Stori BP v25_0704 2026-27 V4 update 100826 CLEAN pass 3.xlsb` | `FD618BC0E3EF71C328C31AFFF90F10EEF1731415BD14A24484B111035C4693E3` |
| Sandbox latest candidate | `C:\Sandbox\BP v26_0001 - New Blank.xlsb` | `E05274ACD3D4821AC38F013AC45A7B57CE335BD524938D9F21E1B640D0CC5E90` |
| Repository contractual blank | `C:\Repos\Abovo Summit\Library\Blank BP v26_0001.xlsb` | `0B4C06800FE998E8733A5D8EB9CDEFA04F517750275BB0CFDD532928900F5B79` |

The subsequently supplied `Master BP v25_0700 (2).xlsb` was verified as the same file by SHA-256; it was not treated as another baseline.

Exact disposable copies were opened in a separate hidden Excel instance, read-only, with macros/events disabled, external-link updates disabled and calculation manual before opening. Copies were closed without saving. Originals and inspection copies were hash-checked after inspection. No worksheet was unprotected. VBA was inspected statically, without execution; only hashes, module names and comparison metadata were retained, never extracted VBA source or passwords. Disposable parser dependencies were kept outside the repository.

## Method and limits

- Inventory covers all 279 worksheets in the old base and Stori, and all 285 worksheets in each latest blank; all defined names were enumerated.
- Formula expressions and constants were compared by worksheet/cell, and formula-pattern differences were used to triage structural movement versus new logic.
- Excel returns null for some bulk formula requests involving merged cells. Those sheets were re-read recursively in smaller blocks, excluding duplicate merged-cell followers. An empty bulk response was not accepted as an empty sheet.
- Formula-pattern normalization removes address offsets for triage only. It is **not** proof that two formulas resolve to equivalent precedents. Targeted comparisons below supply concrete reference evidence.
- Dropdowns, input protection and covenant formats were inspected at identified sample cells. Conditional-format rules were examined in Live Multivariable Planner AC:AE. This is not an exhaustive visual, validation or object-model comparison.
- ZIP-part hashes and readable VBA module bodies were compared. Differences in chart caches, embedded objects, form binaries or compiled VBA are not certified as semantically harmless by a matching source-body hash.
- Cached values were not recalculated or financially certified. No SOCI acceptance comparison, Check Sheet acceptance run, Summit save, Excel macro execution or round-trip test was performed.

The cached Check Sheet results contain 43 formula status cells in column E, all `OK`, in each inspected file; no populated cached warning was found in the corresponding formula messages in column F. This is a stored baseline observation, **not a fresh calculation pass** and not migration acceptance.

## 1. Verify the four recorded changes

### 25.0701: six more stock columns and Transactional DB update

Evidence:

- `StockType`: old base `Stock Assumptions!E5:X5` (20 physical columns); Stori `E5:AD5` (26).
- Stori's first-period Transactional DB formulas referencing Existing Rental Income occupy 52 rows, versus 40 in the base/latest: two streams for each physical stock column.
- Example: base `Transactional DB!Q275` indexes `Existing Rental Income!E13:X52`; the corresponding Stori stream at `Q305` indexes `E13:AD52`.
- Related downstream reference movement is visible in Cashflow, FFR stock-related cells and covenant stock-number workings.

Status: **structurally corroborated**. Latest blanks retain 20 columns. Expansion is a capacity operation, not itself a reason to copy old calculation formulas over the new template.

### 25.0702: 140 more management-cost columns and Transactional DB update

Evidence:

- `Rep_MgtC_01`: old base `Management Costs Assumptions!D19:W19` (20); Stori `D19:FG19` (160). Related named input rows also extend through FG.
- Transactional DB first-period formulas referencing Existing Management Costs grow from 51 to 331 rows: an increase of 280, consistent with two streams for each of the 140 additional columns.
- Example: base `Transactional DB!Q414` indexes `Existing Management Costs!C11:V50`; Stori `Q556` indexes `C11:FF50`.
- Management Cost Factors references move from column 50 to 330 in `Hidden - CF Structure!A48`, `Cashflow detailed!AZ5`, `Cashflow det b4 interco!AW5` and `Development Cashflow Detailed!AV5`.

Status: **structurally corroborated**. Latest blanks retain 20 management-cost columns. The master expansion and dependent calculations/TDB streams must be handled together, not by independently resizing every overlapping `Rep_` range.

### 25.0703: stock grouping and new management-cost driver system

This is a linked feature, not merely more capacity:

1. Stori adds `StockGrouping = 'Global Assumptions'!A34:A48`: 15 physical slots, with 14 populated group labels in the inspected file.
2. A new Stock Grouping input row is inserted at `Stock Assumptions!6:6`. `E6` has list validation `=StockGrouping`; Owned/Managed moves to row 7. The latest blanks still have Owned/Managed at row 6 and have no StockGrouping name.
3. `CostDrivers` moves from `Management Cost Drivers!C6:Z6` to `Other Income Workings!E57:CN57`.
4. The new driver block in Other Income Workings has 45 income columns (E:AW), a separator (AX), aggregate stock (AY), 26 individual stock columns (AZ:BY), and 15 stock-group columns (BZ:CN).
5. Rows 59-60 identify first nonzero income and its base value; rows 62-101 supply the 40-year driver values. Stock-group values use the new grouping inputs and SUMIF aggregation.
6. Management Costs Assumptions row 53 still validates against `CostDrivers`, but row 56 now reads the base value from Other Income Workings. `D56` specifically switches its INDEX source to `E60:CN60` and MATCH lookup to `CostDrivers`.
7. Variable Management Costs uses the new driver list/workings for both management-cost streams.

Status: **formulas, names and sampled dropdown behavior corroborated**. The feature is absent from both latest blanks. Current `Structure.xml` uses `Rep_CostDrivers`, but contains no StockGrouping input definition. Recreating workbook logic alone will not expose the new grouping input correctly in Summit.

### 25.0704: client covenants, stress-test formats and chart correction

| Covenant position | Generic base and latest blank | Stori |
|---|---|---|
| 1 | Gearing | Borrowings to Net Worth (%) |
| 2 | Operating Margin (%) | Rental Income to Interest Payable (ratio) |
| 3 | EBITDA MRI (%) | Nationwide Interest Cover (ratio) |
| 4 | Debt / Unit (currency) | Deficit (%) |
| 5 | Debt | Debt |

Concrete changes:

- `OW - Covenant Workings!AZ8:BX47` adds 40 years of client covenant workings: 21 populated formula columns, 840 formulas, with separator columns between groups.
- `Covenants!C11:C50`, `G11:G50`, `K11:K50`, `O11:O50` point to BD, BH, BX and BL of those workings instead of the old definitions.
- `Covenants!P11:P50` and `Multivariable Planner!BL10:BL49` remove the old `/1000` scaling because the fourth covenant is now a percentage, not debt per unit. Their number formats change accordingly.
- Sample cells `Covenants!G11` and `K11` are decimal ratios in Stori, while the base/latest formats are percentages. `O11`, `P11` and `Multivariable Planner!BL10` are percentages in Stori.
- Live Multivariable Planner rules on `AC8:AE48` change numeric-format covenant IDs from 4/5 to 2/3/5, and percentage IDs from 1/2/3 to 1/4. The inspected breach-color rules retain their formulas.
- `OW - Charts Source Data!D67` negates the entire old sum of Cashflow detailed Q9, S9 and T9.
- The extracted cells in `OW - Covenant Calculation` and `OW - Live Covenant Calculation` are unchanged; their input definitions/workings and downstream presentation change instead.

Status: **the recorded changes are present**, including the otherwise easy-to-miss unit changes and conditional formatting. This verifies implementation differences, not the correctness of the covenant contract or lender requirements.

The latest blanks retain the generic definitions. `Structure.xml:13061` onward also hard-codes generic covenant captions and formats. Those definitions require a client-specific mapping if the bespoke covenant system is carried forward.

## 2. Check the latest template

### Version-history collision

Both branches share the base's history through row 693. Their next entries differ:

| Label | Stori history, August 2025 | Latest generic history, 2026 |
|---|---|---|
| 25.0701 | Stock capacity +6 | FFR 2026 update |
| 25.0702 | Management-cost capacity +140 | Housing asset/tax/OFA named-range updates |
| 25.0703 | New cost drivers and stock grouping | TDB Other Fixed Asset buffer correction |
| 25.0704 | Client covenants and chart sign | New WG FVA sheet |

The later generic history also records validation improvements, Interest Payable Breakdown, WG/FFR-related updates, Credit Rating, value-for-money changes, a TDB heading correction and snapshot/comparison sheets. Some history dates are not chronological. Neither filenames nor version numbers should be used as a monotonic migration ordering rule.

### Two different 26.0001 artifacts

The Sandbox blank and repository blank have matching extracted cell formula/constant contents and matching complete name inventories. Both lack the Stori features above. However, they are **not interchangeable binary artifacts**:

- File hashes differ; package differences include styles, chart/object parts and VBA.
- Static VBA inspection finds 339 modules in the Sandbox blank and 337 in the repository blank. `Sheet232.cls` and `Sheet233.cls` exist only in the Sandbox copy and contain no body text after Attribute lines are excluded.
- The shared `Menu_Module.bas` has a source-body difference in `BPMenu`. Six form-module differences are Attribute-only. Menu behavior was not executed or certified.

The repository blank remains the contractual master under AGENTS.md. This review does not replace it with the Sandbox file. A real migration run must freeze the exact chosen target hash and test its Excel menu behavior; matching worksheet contents are insufficient.

### Preserve newly generated formulas

Old models and the latest blanks have different formula-generation states. In inspected 40-year calculation blocks:

| Block | Old base | Stori | Latest blanks |
|---|---|---|---|
| Existing Rental Income, rows 13-52 | 1 of 20 columns has 40 formulas | 24 of 26 have 40 formulas | All 20 have 40 formulas |
| Existing Management Costs, rows 11-50 | 1 of 20 columns has 40 formulas | 97 of 160 have 40 formulas | All 20 have 40 formulas |

No partially filled formula column was found in those sampled blocks: each has either 40 or zero formula expressions. This is not an assertion that every dependency family has been checked for contiguity.

The old blank columns are not evidence that the new formulas should be deleted. Expand the new template through approved structural operations, preserving its fully generated formula pattern; then populate the inputs and apply only the reviewed bespoke formula changes.

## 3. Differences beyond the four history entries

### A. Formula-bearing user inputs: confirmed migration omission risk

There are additional formulas in editable assumption cells. Examples verified directly in protected worksheets:

- `Management Costs Assumptions!E25`: unlocked, containing a balancing formula over other management-cost inputs; the old/latest blank cell is empty and unlocked.
- `Stock Condition Inputs!H15`: unlocked arithmetic formula; old/latest cell is empty and unlocked.
- `Stock Condition Inputs!AB35`: unlocked arithmetic formula in Stori's expanded area; outside the original active input width in the old/latest templates.

Five further formula-bearing entries were found in `Accounts Assumptions!D10`, `D33`, `D50`, `D63` and `D64`, including arithmetic and unit conversion. Those cells' protection/eligibility was not individually checked in this pass, so they are coverage-review candidates rather than certified writable migration targets.

These are not automatically unlogged bespoke calculation-engine changes: many are the client's way of entering assumptions. But the current upgrader explicitly skips **all source cells with formulas**, at `Services/BusinessPlanComparisonService.vb:272`. Consequently, eligible input formulas can disappear even when their calculated values should have been migrated.

Recommended policy for review: every eligible source input receives a disposition. Preserve approved formula intent where it is meaningful and mappable; otherwise copy its freshly calculated value with an explicit formula-to-value notification. Never silently omit it, and never overwrite a target-owned calculation formula just because the source cell was editable. A value-only result preserves the starting assumption, not the source's future balancing behavior.

The user's existing values-copy requirement supplies the default: eligible formula-bearing inputs should contribute their calculated values, with a notification. Carrying an actual formula forward belongs to the separately reviewed bespoke-logic plan.

One stock-condition input uses a 1.23 multiplier where several nearby formulas use 1.023. This is a **review question, not a diagnosed error**: business intent has not been established. Do not automatically correct historical inputs.

### B. Additional capacity growth

| Named range | Old base physical size | Stori physical size |
|---|---|---|
| StockCondCats | 13 columns | 26 columns |
| HouseTypeInID | 12 columns | 24 columns |
| DepnType | 5 columns | 9 columns |
| Rep_OInc_05 / Rep_OInc_06 | 16 rows | 46 rows |
| Rep_OInc_07 | 11 rows | 19 rows |
| Transactional_Records | 1,594 x 74 | 2,393 x 74 |

These are physical extents, not counts of populated business records. They go beyond the two logged stock/management-cost expansions, but are consistent with normal population and structural growth. They require migration capacity planning; they are not client-facing errors merely because the target starts smaller.

### C. Name inventory changes are not lost business inputs

Old base: 1,756 names; Stori: 1,748. Stori adds StockGrouping and removes nine hidden function-compatibility placeholders whose RefersTo is `#NAME?`: eight `_xleta.*` placeholders and `_xlfn.ANCHORARRAY`. Thus the net reduction of eight names is **not** evidence that eight business ranges were deleted. There are also 565 changed same-name references, many resulting from expansion and shifted rows/columns.

Latest blanks: 1,761 names. New names include Rep_OFA_065, Rep_Tax_10 and WG-related names. Retain those new-release contracts rather than replacing the complete name collection with the old file's collection.

### D. Downstream formula movements examined

Targeted inspection explains several apparent unrecorded formula changes as structural consequences:

- `FFR Inputs Adj Stmt!C35:I36`: Stock Numbers references move by the six stock columns.
- `Investments!I5`, `O5`, `U5`: Funding Assumptions precedents shift by one row.
- `Development Cashflow Detailed!AQ9:AQ48`: Existing Management Costs precedents move by 280 columns, consistent with the two expanded management-cost streams.
- The old portion of OW - Covenant Workings has 40 changed stock-number references in AI8:AI47, also consistent with stock expansion.

These should not be presented as new bespoke features solely because their raw formula strings differ. This targeted explanation does not certify every moved reference throughout the workbook.

### E. VBA and other workbook objects

Stori and the old base each have 333 readable VBA modules. No modules are added or removed, and **all readable module bodies match after normalizing line endings and excluding Attribute declarations**. Six form modules have Attribute differences. This provides no evidence of Stori-specific VBA source logic in these files; it is not proof that form designer binaries, project metadata or compiled state are identical.

Package hashes also identify chart, drawing, comment and other object changes. Many may be cached values, save-time metadata or population effects. Their business meaning has not been exhaustively classified. They remain a separate object/visual review gate, not automatically approved patches or known client customizations.

## Recommended migration plan and acceptance gates

1. **Freeze artifact identities.** Use the source, provisional old base and chosen latest template hashes. Record the shared history prefix and the actual feature evidence; do not require fictitious strict historic versioning.
2. **Compile capacity once per structural family.** Stock, management costs, repairs, development types, depreciation components, other income and any other discovered input families need capability-specific expansion, dependency ordering and postconditions. Re-resolve all named ranges after expansion.
3. **Apply a reviewed Stori feature bundle.** Stock grouping, drivers, management-cost calculation wiring, client covenants, units/formatting and chart sign form explicit operations with before/after checks. Preserve unrelated generic 2026 release work.
4. **Update the client interface definition.** Include grouping inputs and workbook-correct covenant labels/formats. Do not use generic percent/amount metadata for Stori's changed ratios and percentages.
5. **Populate every eligible input with an explicit disposition.** Deduplicate overlapping ranges; distinguish typed values, formula-bearing inputs, target-owned calculations, missing mappings and excluded read-only ranges. Preserve invalid source content where permitted and report it, rather than silently coercing or omitting it.
6. **Validate an isolated result.** Recalculate source and result in controlled copies, then compare the detailed SOCI period by period, covenant values/units and check results. Expected generic statutory/FFR effects must be explained separately from migration errors. Require input-coverage and structural checks as well as headline financial agreement.
7. **Round-trip and preserve originals.** Save result and report under new names. Reopen in Microsoft Excel, test menu/VBA and representative structural operations, then reopen in Summit. Check formulas, names, dropdowns, protection, snapshot invalidation and client features. Source and templates remain immutable.

Recommended next implementation is the reviewable preflight/coverage report and explicit feature/capacity plan, before any automatic bespoke patch application. Financial acceptance and Abovo approval of the covenant/driver semantics remain necessary.

## Evidence and related documents

- `Stori_Baseline_Version_History_Review_2026-09-19.md`: earlier baseline/history evidence; the name-count question raised there is resolved above.
- `Upgrade_Migration_Architecture_Review_2026-09-19.md`: service review, preservation contract and proposed orchestration.
- Disposable detailed inspection evidence and scripts: `obj/MigrationReview/` (ignored; includes client cell data and must not be committed).
- Durable evidence retains findings and identities, not workbook payloads or extracted VBA source. There was no production-code change and therefore no application build/test-release increment in this investigation.
