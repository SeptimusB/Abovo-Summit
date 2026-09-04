# Abovo Summit contract XLSB health and performance audit

Audit date: 24 August 2026

## Standing review requirement

Review this document and `Contract_XLSB_Audit_Evidence_2026-08-24.json` before any workbook, calculation engine, Transactional DB, custom-function, structural range, or validation change.

## Authority and method

- Authoritative master: `Z:\Sandbox\TestFileClean.xlsb`
- Repository copy: `Library/TestFileClean.xlsb`
- Exact copy SHA-256: `B248B1C733E1E3293536FBE1DBC9576D56FD4D34BEB74A3898B7F3D7333BCFBE`
- The authoritative master was not opened or saved by the audit. Excel inspected a disposable exact local copy read-only, with macros/events disabled and links not updated.
- Thirty-five protected sheets were unlocked only in memory on the disposable copy; no password is stored in project evidence.

## Executive conclusion

The contract workbook is structurally healthy and natively fast on the audit machine. The master and library copy match exactly; saved Check Sheet results are all OK; no broken defined names, external links, data connections, circular reference, or full-column calculation formulas were found. Native Excel warm full calculation was about 0.42 seconds and a full dependency rebuild about 2.15 seconds.

The priority is correctness across Excel/VBA and Summit rather than wholesale formula optimisation. The most important confirmed issue is a Summit `PMCOST` parameter-metadata mismatch. The greatest latent workbook risk is the heavy dependence on worksheet order through 35,920 3-D formulas.

## Core metrics

- Worksheets: 281 (279 visible; 35 protected)
- Defined names: 1,758; 262 multi-area; no broken, external, or volatile names
- Formula cells: 638,360 across 5,625 R1C1 patterns
- Constants: 87,287
- Formulas at least 500 characters: 6,582; maximum 901 characters
- 3-D reference formulas: 35,920
- Volatile function occurrences: 1,313, all `CELL`
- Saved formula-error cells: 467; 466 are deliberate chart `#N/A` gaps and one is an unguarded lookup
- Excel read-only open: 2.37 s; warm full: 0.42 s; full rebuild: 2.15 s

## Findings

### F1 - High: Summit PMCOST metadata does not match the eight-position compatibility signature

Evidence: The workbook has 380 PMCOST calls and passes eight compatibility arguments, but FinalYear is deliberately unused, so seven inputs affect the result. clsPMCost.vb lines 20-27 advertises five value parameters plus two references (seven positions), while lines 85-92 accepts parameters 0 through 7 (eight positions).

Implication: DevExpress validation, parsing, or evaluation may diverge from Excel even when current sample values appear to calculate.

Recommendation: Preserve the eight-position compatibility signature, document that FinalYear is unused, advertise six required value positions plus two required references, and add parity tests before relying on Summit recalculation.

### F2 - High: Cross-engine UDF parity is not regression-tested

Evidence: Development Expenditure contains 380 PMCOST and 380 RESPCOST calls. Summit registers native replacements, but a macro-disabled Excel full rebuild makes seven Check Sheet rows #NAME? while the saved workbook checks are all OK.

Implication: Excel/VBA and Summit may produce different values or error behaviour for missing matches, invalid periods, or malformed ranges.

Recommendation: Create an automated parity harness comparing the 760 call results and key outputs after Excel/VBA and DevExpress recalculation. Make RESPCOST parameters required if they are unconditionally dereferenced and add explicit error guards.

### F3 - Medium: 35,920 formulas depend on worksheet ordering through 3-D references

Evidence: Examples include SUM('Component 1:Component 12'!...) and grouped Development/All Schemes blocks.

Implication: Moving or inserting a sheet inside or outside a boundary can silently change results without producing a broken reference.

Recommendation: Retain these efficient formulas for now, but add an expected sheet-order/block-membership manifest and validate it on open, save, and structural commands.

### F4 - Medium: Custom functions perform repeated formula-engine MATCH evaluations

Evidence: Both Summit UDF implementations build and evaluate a MATCH expression inside a loop for every call. With 760 workbook calls and up to 40 periods, this can create thousands of nested FormulaEngine evaluations.

Implication: This is the most credible Summit-specific calculation hotspot even though native Excel recalculation is already fast.

Recommendation: Replace repeated Engine.Evaluate calls with direct range reads and an indexed year-to-rate lookup, preserving Excel MATCH semantics and testing parity.

### F5 - Medium: Formula complexity is concentrated and highly repetitive

Evidence: There are 6,582 formulas of 500-901 characters, mainly on Development Capital and Development Expenditure. INDEX/MATCH and IF logic dominate the workbook.

Implication: Maintenance and structural-change risk are greater than the current timing impact.

Recommendation: Refactor only measured hotspots using helper rows/columns or cached match indices. Do not introduce LET/XLOOKUP until the minimum supported Excel version is agreed.

### F6 - Low-Medium: Volatile CELL calls are duplicated

Evidence: 1,313 CELL occurrences exist: 924 on Hidden - BP Menu Sheet, 388 on Hidden - Sheet Lists, and one on Stress Sensitivity List. The menu formulas call CELL four times per formula to derive a sheet name.

Implication: Every recalculation revisits workbook metadata unnecessarily; the current measured impact is modest.

Recommendation: Evaluate each CELL result once in a helper cell (or LET only if supported) and reuse it, reducing the menu occurrences by about 75%.

### F7 - Low-Medium: Several sheets have inflated used ranges

Evidence: 5 Yr Monthly Cashflow uses 337 rows although the last populated row is 102; 5 Yr Quarterly Cashflow uses 262 versus 27. Transactional DB extends to 78 columns while populated content ends at 74.

Implication: This can increase rendering, navigation, print, and file-maintenance overhead more than formula calculation time.

Recommendation: On a disposable copy, remove only demonstrably obsolete formatting/objects beyond the real model area, then run full Excel/VBA/Summit round-trip tests.

### F8 - Low: One unguarded #N/A remains outside intentional chart suppression

Evidence: Of 467 saved formula-error cells, 466 are deliberate NA() chart gaps. Hidden - Tenure Totals Start!B6 is an unguarded INDEX/MATCH #N/A.

Implication: The error is currently latent because the saved Check Sheet is OK, but it can leak into future imports or interfaces.

Recommendation: Confirm the expected blank-state behaviour and use IFNA/validation if absence from ImportedSchemes is normal.

## Recommended sequence

1. Correct and regression-test the Summit custom-function contracts before changing workbook formulas.
2. Add an automated Excel/VBA versus Summit parity pack covering the 760 UDF calls, Check Sheet, key outputs, Transactional DB and structural insertions.
3. Add a sheet-order/block-membership validator for all 3-D ranges.
4. Instrument Summit calculation, datasource refresh and UI refresh separately; native Excel time is already good.
5. Trial de-volatilising menu-sheet `CELL` formulas and used-range cleanup on disposable copies only.
6. Refactor long/repeated formulas only where profiling proves value and the supported Excel version permits the chosen functions.

## Evidence limits

- Embedded VBA was neither extracted nor executed. The saved Check Sheet is all OK; seven #NAME? rows after the audit rebuild are an expected limitation of macro-disabled Excel and are not treated as workbook defects.
- Excel timings exclude Summit's DevExpress calculation, UDF implementation, RangeDataSource work, grid refresh and UI rendering.
- Per-sheet COM Calculate timings were dominated by call overhead and are not used to rank performance.
- No workbook or application code was changed by this audit.

## Post-audit authority decision - 4 September 2026

This report remains the health and calculation baseline for the workbook audited on 24 August. A reserved-capacity prototype was trialled on 3 September but withdrawn following client review. The production contract now uses the validated pre-capacity workbook and the established structural-shift behavior shared by Summit and Excel/VBA.

The current master is 11,813,448 bytes with SHA-256 `D2CA9A14B432C6914612917129F20446DDD33A9BD7740AFAC704A5C357790C17`. It has 283 worksheets, 1,758 names and `Transactional_Records = Transactional DB!A6:BV1599` (1,594 by 74), with no `_Capacity` or `_Continuation` names. Targeted validation found no broken names and no non-OK saved Check Sheet results. See `TestFileClean.analysis.md` and `Abovo_Summit_Project_Scope_Audit.md` for the current authority record.
