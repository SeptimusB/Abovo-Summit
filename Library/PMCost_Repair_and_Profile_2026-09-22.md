# PMCost signature repair and maintenance-cost profiling — 22 September 2026

Status: Ready to test — Summit 2.71. Engineering checks are not client/accountant acceptance.

## 2.71 implementation follow-up

The user approved promotion of the native MATCH prototype. Both production
functions now invoke DevExpress MATCH directly with approximate mode 1. Each
invocation uses current range references; no cross-call caching is introduced.
Arithmetic, accumulation order and the unused FinalYear position are preserved.

Missing matches now return the native worksheet error (normally #N/A), rather
than indexing -1 and throwing. Invalid/missing reference parameters return
#VALUE!, an undersized rates range returns #REF!, and input/rate errors propagate.
These are intentionally visible errors, not invented zero results. The VBA's
broad error suppression on malformed inputs is not reproduced or claimed equal.

The profile fixture now freezes the old string-based evaluator in test code and
compares it with the actual production implementation, rather than comparing
two copies of the new evaluator.

Release paired profile: 37,688 assertions passed. All valid synthetic results are
exactly equal; old missing-lookup exceptions are checked against the new #N/A.
On the frozen AGL recovery, 475 actual calls to each function per rebuild:

| Path | Whole rebuild ms | PMCost self-time ms | RespCost self-time ms |
| --- | ---: | ---: | ---: |
| Old evaluator | 8,046 | 36.67 | 44.81 |
| Production direct MATCH | 7,562 | 26.99 | 27.41 |
| Production direct MATCH | 7,552 | 25.47 | 27.11 |
| Old evaluator | 7,549 | 36.87 | 44.90 |

Combined function self-time averages 81.625 -> 53.49 ms (about 34% less).
This is a localized saving, not a 34% whole-workbook improvement. All four
full-workbook formula/value digests equal the digest recorded below. The supplied
workbook was loaded read-only into memory and was not saved.

Log: obj/MaintenanceCostProfile/dd466b2054294fd8af490a315934c9e2/profile.log.
See Recovery_User_Changes_2026-09-22.md for the accompanying recovery changes.

The 2.71 native fixture passed 840 assertions with the read-only AGL evidence:
all 570 PMCost-containing cells numeric, zero NAME errors, existing VALUE/N/A
counts and deliberate Check Sheet warnings unchanged. Debug synthetic tests
passed 266 assertions, including both cost functions, missing/misaligned ranges,
error propagation, changed rate inputs with ordinary dirty calculation, and
synthetic XLSB/XLSM export/reload. Debug and Release builds passed.

The sections below retain the earlier 2.70 decision/evidence chronologically;
their "not applied" statements describe the prototype stage, not current code.
For current manual acceptance, restart into 2.71 and follow the same disposable
Summit -> Excel/VBA -> Summit checks below.

## 2.70 production change (historical)

Only clsPMCost.vb's function contract changed (plus the test-release version):
- Registered name: PMCost, matching the master VBA declaration.
- Eight required arguments: Year, AllUnits, UnitCost, FirstManage, LastManage, FinalYear, ApplRates, ApplYears.
- The missing sixth scalar metadata entry is now present; FinalYear stays unused in the cost formula.
- The two range parameters remain seventh/eighth.
- Existing evaluation arithmetic, error behavior and RespCost implementation are unchanged.
- Uppercase, lowercase and VBA-style formula spellings remain accepted.

Master declarations were inspected in memory without executing VBA or retaining extracted source. Metadata reconfirmed PMCost has eight parameters, RespCost has seven, and FinalYear occurs only in each VBA declaration. This matches the previously retained module hashes in Master_VBA_Inventory_2026-09-21.json.

No workbook formulas, names, VBA modules, structure XML, hidden sheets or user files were changed.

## Validation

- Debug and Release AnyCPU builds passed, producing bin/Debug/Abovo-summit.exe and bin/Release/Abovo-summit.exe.
- Tools/Test-PMCost.ps1 passed against both configurations: 233 synthetic assertions per configuration.
- Tests cover invariant/localized names, required argument types/count/order, PMCost/PMCOST/pmcost spelling, changed numeric FinalYear values without changed output, hand-calculated examples, sparse approximate lookup, nested INDEX reference arguments, malformed arity rejection, and synthetic XLSB/XLSM native save/reload/full-recalculation.
- Release additionally loaded the frozen AGL recovery evidence without saving and passed 807 total assertions. All 570 PMCost-containing cells returned numbers. Full recursive rebuild produced zero NAME errors, compared with 52,161 in the pre-repair native reload test.
- The three deliberate Check Sheet warnings E23/E33/E39 remain Check.
- Other existing errors remain: the repaired full rebuild observed 582 VALUE and 572 N/A cells (the latter includes the 98 literal markers). This repair does not suppress other errors or claim that all cached values remain identical after a correct full rebuild.
- Original recovery and frozen-evidence SHA-256 remain 405BC929A6A99A77B986EB05A0279E94196B99F025388D88B099A711F5D742B5.
- The untouched Blank master hash remains 0B4C06800FE998E8733A5D8EB9CDEFA04F517750275BB0CFDD532928900F5B79.
- git diff --check passed.

Reproduce:
`powershell -NoProfile -ExecutionPolicy Bypass -File Tools/Test-PMCost.ps1 -Configuration Debug`

`powershell -NoProfile -ExecutionPolicy Bypass -File Tools/Test-PMCost.ps1 -Configuration Release -Workbook "<disposable recovery evidence path>"`

Tests restore the previous calculation engine and never save the supplied client workbook. Their synthetic exports are confined to a fresh obj/PMCostTests directory.

## Speed investigation — prototype only

Both classes currently construct a MATCH formula string and call FormulaEngine.Evaluate on every year-loop iteration. A test-only alternative calls the existing DevExpress MATCH function's Evaluate method directly, using native parameters and explicit approximate-match mode 1. It retains the current arithmetic and accumulation order and does not replace MATCH with a hand-written approximation or introduce shared/stale caches.

**This optimization has not been applied to either production class.** The user requested investigation; the experiment is isolated in Tools/MaintenanceCostProfile.cs and Tools/Profile-MaintenanceCosts.ps1.

### Synthetic comparison

2,375 paired cases per function, four passes including warm-up:
- Three measured PMCost passes: baseline 184.05 / 179.85 / 207.99 ms; direct 157.11 / 154.51 / 170.96 ms.
- Three measured RespCost passes: baseline 181.41 / 181.05 / 205.82 ms; direct 155.44 / 154.24 / 171.12 ms.
- Roughly 14–18% lower function self-time in this mixed synthetic workload.
- Exact numeric/error-result agreement; approximate matching, duplicate keys, sparse bands, missing/unsorted keys and empty-loop cases included.
- Invalid-range exceptions are captured only in the synthetic test wrapper so type/location can be compared. No such suppression is enabled in the real-workbook profile or production.
- 38,332 total profile assertions passed, including the real-workbook digest checks below.

### Actual AGL recovery workbook

Four rebuilds in baseline → direct → direct → baseline order, same in-memory workbook, no saves:

| Path | Whole rebuild ms | PMCost self-time ms | RespCost self-time ms |
| --- | ---: | ---: | ---: |
| Baseline | 8,184 | 38.08 | 46.58 |
| Direct MATCH | 8,023 | 24.66 | 26.98 |
| Direct MATCH | 7,888 | 25.47 | 26.63 |
| Baseline | 8,178 | 41.55 | 46.15 |

Each rebuild actually executed 475 calls to each function. This is smaller than the 570 PMCost-containing formula cells because conditional formulas do not always take the custom-function branch.

Average combined self-time: 86.18 ms → 51.87 ms, about 40% faster **for these two functions**, saving approximately 34 ms per full rebuild. The functions therefore represent only about 1% of the measured baseline rebuild. Whole-rebuild variation is larger than the measured UDF saving; this small run must not be advertised as a proven 2–3% whole-model speedup, or as a solution to the multi-second insertion/save delays.

All four hashes of worksheet formula text and typed calculated values match:
`058674A51B8B404A68DC5A2CACAC3F5B0F8E4A3006FAF7E72CC869B4E9597B2B`

Local log:
`obj/MaintenanceCostProfile/fa303febb1f74cc4a78a3d2ae061a59a/profile.log`

Conclusion: a modest, localized optimization is feasible. Prioritize the larger workbook calculation/structural costs if the goal is a noticeable overall delay reduction. Before promoting the prototype, extend dirty-edit/dependency tests, locale/error cases and trusted Excel/VBA parity checks; full-rebuild checksum equality alone does not establish all interactive behavior.

## Separate existing findings

- A missing approximate MATCH can give index zero, then index -1 into ApplRates, throwing ArgumentOutOfRangeException in both current classes. Synthetic tests reproduced this. The master VBA has broad error suppression; resolving the intended edge behavior is a separate compatibility repair, not a speed-only refactor.
- RespCost advertises optional parameters which the implementation unconditionally reads. This investigation did not change that contract.
- Excel/VBA financial equivalence is still not certified: no macros were executed, and synthetic native round trips are not the same as opening/saving a populated model in Excel with its VBA enabled.

## Manual acceptance

Restart Summit to load 2.70; reopen a disposable recovery file and request a rebuild. Confirm the intentional balance warnings remain and no new NAME-error cascade appears. Then save to a new XLSB, open/recalculate/save through the trusted Excel/VBA workflow, and reopen in Summit. The client/accountant should check development maintenance costs and affected outputs. Never overwrite the sole client original for this test.

## References

- [DevExpress 25.2 custom-function contract](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Functions.ICustomFunction?v=25.2).
- [DevExpress built-in function access](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Functions.WorkbookFunctions.LookupAndReference).
- Library/Contract_XLSB_Audit_2026-08-24.md, findings F1/F2/F4.
- Library/AGL_Recovery_Error_Comparison_2026-09-22.md, pre-repair evidence.
