# Business Plan upgrade and migration validation matrix

**Prepared:** 19 September 2026  
**Related handover:** `Upgrade_Migration_Handover_2026-09-19.md`  
**Implementation baseline:** `b8b231c`

## 1. Purpose

This matrix defines the evidence required before an upgraded workbook can be treated as safe. A successful build or a report with no exceptions is not sufficient. The output must preserve the latest generic model, correctly migrate client data and approved customisations, calculate credibly and round-trip through Excel/VBA.

## 2. Test artifact roles

Each test file must have an explicit role and hash recorded before execution:

| Role | Writable? | Purpose |
| --- | --- | --- |
| Old generic template | No | Base for distinguishing client population/customisation |
| Old populated client workbook | No | Source client data and bespoke branch |
| New generic template | No | Authoritative destination product release |
| New result clone | Yes | Only workbook allowed to be mutated |
| Migration report | Yes | Human review evidence |
| Migration manifest | Yes | Machine-readable provenance and operations |

The original three input files must retain identical hashes after the test.

## 3. Representative pilot set

At minimum include:

1. a current blank generic workbook;
2. a moderately populated current workbook;
3. an older generic workbook with its corresponding populated client workbook;
4. a bespoke client workbook and its exact old generic base;
5. a source with more records than the new template for each important structural rule;
6. a source with fewer records than the target;
7. a workbook spanning an FFR schema change;
8. a workbook containing known type mismatches and missing/renamed ranges;
9. a heavily populated workbook representative of client scale.

## 4. Validation matrix

### A. Identity and provenance

| ID | Test | Expected evidence |
| --- | --- | --- |
| A01 | Identify model family/release from embedded XML | Stable IDs recorded; filename shown only as supporting evidence |
| A02 | Missing embedded identity | Explicit fallback classification and warning; no guessed silent match |
| A03 | Filename conflicts with embedded identity | Embedded identity wins; conflict reported |
| A04 | Client variant identification | Variant or unknown state recorded |
| A05 | Input hashes | All source/template hashes recorded and unchanged after migration |
| A06 | Result lineage | Manifest links result to all source/template hashes and release IDs |

### B. Read-only and clone guarantees

| ID | Test | Expected evidence |
| --- | --- | --- |
| B01 | Assess standard upgrade | No input workbook content or metadata changes |
| B02 | Assess bespoke upgrade | No input workbook content or metadata changes |
| B03 | Cancel before execution | No clone or partial output unless explicitly retained as a draft |
| B04 | Begin execution | Result is cloned before the first structural/value mutation |
| B05 | Save failure | Source/templates unchanged; result remains recoverable under a new path |

### C. Range discovery and classification

| ID | Test | Expected evidence |
| --- | --- | --- |
| C01 | XML-declared writable named range | Included once even when referenced by many interfaces |
| C02 | `Rep*` input omitted from XML | Reported as an XML coverage candidate; not silently migrated solely by prefix |
| C03 | `Rep*` calculated/read-only range | Excluded from input writes |
| C04 | Non-`Rep` XML-declared input | Included |
| C05 | Formula or locked target cell | Excluded from ordinary value migration |
| C06 | Renamed semantic field | Mapped by stable semantic ID or explicit recipe |
| C07 | Split/merged field | Explicit transform required and reported |
| C08 | Repeated named-range reference | One migration operation with reference provenance |

### D. Master/dependent structural ranges

| ID | Test | Expected evidence |
| --- | --- | --- |
| D01 | Known master expansion | One master rule runs; all declared dependents align |
| D02 | Master and dependent both appear in XML | No double expansion |
| D03 | Formula contiguity before expansion | Pattern recorded and accepted |
| D04 | Broken formula contiguity | Blocking error or explicit reviewed repair recipe |
| D05 | Dependent range fails to move with master | Structural validation fails and result is not accepted |
| D06 | Source contains more records | Required count added through `WorkbookStructureRuleManager` |
| D07 | Source contains fewer records | No automatic deletion without an approved rule |
| D08 | Multiple expansion groups | Deterministic dependency order and no overlapping mutation |
| D09 | Structural failure mid-plan | Recovery/rollback behaviour leaves a clear, recoverable result |
| D10 | Re-run same plan | Idempotent result; no duplicate added records |

### E. Values and types

| ID | Test | Expected evidence |
| --- | --- | --- |
| E01 | Text, Boolean, integer, decimal, percent, currency and date inputs | Correct typed values and display formats |
| E02 | Invalid source value for declared type | Raw value copied where policy permits; precise warning/error reported |
| E03 | Empty source into populated target | Policy applied explicitly and recorded |
| E04 | Formula source in input range | Formula is not copied as an ordinary input value |
| E05 | Duplicate target address from overlapping declarations | One deterministic write |
| E06 | Batch write failure | Failing cells isolated; successful cells and failure status accurately reported |
| E07 | Undo/history policy | Upgrade appears as coherent audited groups without corrupting history |

### F. Three-way bespoke merge

| ID | Test | Expected evidence |
| --- | --- | --- |
| F01 | Client-only input change | Migrated as client data |
| F02 | Client-only formula/name/format change | Classified as candidate bespoke change |
| F03 | New-generic-only change | Retained from new generic template |
| F04 | Both branches make identical change | Applied once without conflict |
| F05 | Both branches change same semantic element differently | Conflict requiring explicit resolution |
| F06 | Old generic logic differs from new generic | Not misclassified as a client customisation |
| F07 | Approved bespoke patch replay | Deterministic, versioned and idempotent |
| F08 | Unsupported bespoke VBA/custom UI change | Explicit blocking/manual-review result; no raw VBA source stored |

### G. Calculation and financial outputs

| ID | Test | Expected evidence |
| --- | --- | --- |
| G01 | Summit dependency-sensitive calculation | Completes and restores original engine/settings |
| G02 | Check Sheet | No non-OK result unless explicitly accepted and documented |
| G03 | Detailed SOCI | Differences classified as expected generic change, migrated-data issue or unexplained error |
| G04 | Transactional DB | Named-range geometry, formulas and reconciliations valid |
| G05 | Custom functions | Required calls calculate consistently with the supported baseline |
| G06 | Sheet-order manifest | All 3-D reference blocks retain expected membership/order |
| G07 | Blank and sparsely populated models | Required headings/structures remain available and calculate |
| G08 | Heavily populated model | Correct results within acceptable time/memory bounds |

### H. FFR

| ID | Test | Expected evidence |
| --- | --- | --- |
| H01 | Same FFR schema version | Declared fields migrate and validate |
| H02 | Supported FFR version upgrade | Version recipe transforms inputs and layout correctly |
| H03 | Unsupported FFR source version | Blocking compatibility error before mutation |
| H04 | Government template layout/formula change | New generic logic retained; old inputs semantically mapped |
| H05 | FFR export/generation | Output opens and matches the applicable regulatory template |

### I. Excel/VBA round-trip

| ID | Test | Expected evidence |
| --- | --- | --- |
| I01 | Save result in Summit | XLSB saves under new name without source mutation |
| I02 | Open result in Microsoft Excel | No repair prompt or damaged names/formulas |
| I03 | Use relevant VBA/custom menu actions | Expected commands still work |
| I04 | Recalculate in Excel/VBA | Check Sheet and selected outputs agree within defined tolerances |
| I05 | Save in Excel and reopen in Summit | Inputs, formulas, names, VBA and structure remain usable |
| I06 | Reconcile external Excel edit | Established conflict/reconciliation behaviour remains intact |

### J. Reporting and recovery

| ID | Test | Expected evidence |
| --- | --- | --- |
| J01 | Benign geometry expansion | Notice, not error |
| J02 | Ordinary value difference | Information/detail, not error |
| J03 | Ambiguous semantic mapping | Warning or conflict with both candidates shown |
| J04 | Failed structural/write/calculation/check operation | Error with operation, workbook, range/cell and recovery status |
| J05 | Human report | Summary and drill-down sheets are readable and copyable |
| J06 | Machine manifest | Contains identities, hashes, plan, operations, decisions, results and tool version |
| J07 | Sensitive content | No passwords or extracted VBA source |
| J08 | Interrupted operation | Result is either rolled back or clearly marked recovery-only |

### K. Performance and usability

| ID | Test | Expected evidence |
| --- | --- | --- |
| K01 | Assessment only | Does not perform unnecessary full calculations or bind live analysers |
| K02 | Structural expansion with analyser absent/hidden/visible | Timings recorded and UI remains responsive with progress messaging |
| K03 | Large comparison | Detail limit/truncation is explicit and summary counts remain correct |
| K04 | Report rendering | Large reports remain navigable without excessive best-fit/rebinding cost |
| K05 | Cancel/close | Detached source workbooks dispose cleanly; open Summit models remain valid |

## 5. Minimum acceptance gate

No migration should be described as successful unless:

1. all source/template hashes are unchanged;
2. the result was cloned from the intended new generic release;
3. all planned structural rules completed and their invariants passed;
4. all required inputs were copied or explicitly reported;
5. no unresolved blocking conflict remains;
6. Summit calculation completed;
7. the Check Sheet meets the agreed acceptance policy;
8. SOCI and FFR differences are classified rather than merely counted;
9. the XLSB opens and recalculates in supported Excel/VBA without repair;
10. the Excel-saved result reopens correctly in Summit;
11. the human report and machine manifest were saved;
12. the original workbooks are byte-for-byte unchanged.

## 6. Evidence record template

For every pilot retain:

```text
Test date/time:
Summit version/commit:
Machine and bitness:
Old generic path/hash/release:
Old client path/hash/release/variant:
New generic path/hash/release:
Result path/hash:
Migration plan ID/hash:
Structural operations:
Values copied/skipped/failed:
Notices/warnings/errors/conflicts:
Calculation timings:
Check Sheet result:
SOCI classification summary:
FFR version/result:
Excel/VBA round-trip result:
Summit reopen result:
Reviewer and decision:
```

