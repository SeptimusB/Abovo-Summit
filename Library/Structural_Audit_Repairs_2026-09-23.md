# Structural audit and repair trial — 23 September 2026

Test release: **2.80**. This is an engineering audit and private-copy validation, not accountant or interactive Excel/VBA acceptance.

## Scope and source identity

Requested: finish OFA/Repairs comparisons against master VBA, remaining deletion boundaries and add/delete tests using C:/Sandbox/Insert Comp/AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb.

The user approved repairing confirmed OFA/Repairs and Capital Expenditure defects and deletion safeguards, then separately approved the service-charge mirror-name collision and Repairs Include-range handling. Source files and masters are not edited.

SHA-256 identities, rechecked during the trial:

- AGL: 30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47.
- Library/Blank BP v26_0001.xlsb: E05274ACD3D4821AC38F013AC45A7B57CE335BD524938D9F21E1B640D0CC5E90.
- Library/Demo BP v26_0001.xlsb: 1ED79726D8D1129C699242C9E8B3E988AC920EA7108D3C0B105BF9B0BC8A5E9C.

Current Blank has 339 VBA modules / 287 procedures; AGL has 335 / 287. The inspected OFA, Repairs and shared Row_Insertion procedures match by procedure hash. That is not a claim of complete model/base parity.

## Confirmed defects and repairs

| Area | Reproduction on AGL / 2.79 | Repair |
| --- | --- | --- |
| OFA insertion | Inserts after, rather than before, the hidden template. One inserted record loses three default/inclusion formulas; three records lose nine. | Insert at LastOFACol minus one; copy full template contents, retain the shifted template, use a visible input column for new widths. |
| Repairs insertion | Three categories leave 17 dependent names undersized and 1,001 cell/formula differences from the independent Excel insertion, including totals omitting new categories. | Insert before the hidden template across all eight linked sheets so dependent ranges/totals extend. |
| OFA/Repairs deletion | Hidden template counted as a deletable record; minima also permit excessive deletion. | Exclude template/end columns and enforce the VBA boundaries below. |
| Capital Expenditure deletion | A new I5 marker survives Delete Last: the command removes the shifted hidden template instead. Blank inverse testing misses this. | Exclude both template and end column, retain the VBA E:F minimum; align DIT SkipLastRecords with genuine inputs. |
| Row minima | CPI/RPI and Cash Journals can be reduced to two rows, below shared VBA InsertRows' minimum of three. | Simple row rules, Journal and Stock Conversion retain three logical rows. More restrictive explicit family minima remain. |
| Economic service-charge rows | IR_ServChg is mistaken for the source of TransCopy_IR_ServChg_01, which belongs to the distinct Service Charge source IR_ServChg_01. A three-row inverse loses three TDB rows, changes 96 names and 45,888 affected formula/constant cells. | Prefer exact source names before interpreting numbered mirror aliases; do not overwrite explicit column-family rules. Integrity inspection and mutations share the resolver. |
| Repairs Include name | AGL's RepIncStkCat is D55:O55 (12 categories), but genuine categories extend to X (21). Even the master VBA insert does not expand this template-excluding range at its edge. | Validate companion geometry before mutation; align with all genuine StockCondCats columns after explicit Repairs add/delete, excluding the template. This is an approved correction beyond literal VBA parity. |

The Include repair changes a named range, not values/formulas. It is not a load-time migration or save repair. Old short names are aligned only during the structural command; unexpected worksheet/row/left-edge/overlong geometry is rejected. Original AGL remains unchanged.

### Verified column-deletion contract

Letters refer to unmodified AGL. Logical counts exclude hidden templates/end markers.

| Family | Last deletable column | Minimum retained | Maximum delete on AGL |
| --- | --- | --- | --- |
| OFA | K | D:F — 3 | 5 |
| Capital Expenditure | H | E:F — 2 | 2 |
| Capital Grant | G | D:F — 3 | 1 |
| Repairs | X | D:I — 6 | 15 |
| Housing Components | L | D:F — 3 | 6 |

Last deletable coordinates and maximum counts match actual VBA conditions, not just macro message text.
The retained leading columns are also protected from selected-record deletion, so selecting individual records cannot bypass the VBA's fixed left boundary.

## Validation method and reproducibility

Tools/StructuralBoundaryAudit.cs, run through Tools/Test-ColumnFamily.ps1, creates unique private copies under obj/ColumnFamilyTests. It records global/local names and affected-sheet plus TDB formulas/constants, checks add/delete, populated new-record deletion, associated mirror dimensions, forbidden/maximum deletions, protection/visibility and calculation-mode/engine restoration. No source workbook is saved.

Example (from repository root):

~~~powershell
Tools/Test-ColumnFamily.ps1 -Workbook 'C:/Sandbox/Insert Comp/AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb' -Rule aligned:REPAIRS_RECORDS -Count 3 -Fixture StructuralBoundaryAudit.cs
~~~

Other modes: marker:CAPEX_RECORDS; boundary:OFA_RECORDS -Count 2 (retained count); minimum:OFA_RECORDS (delete down to minimum and reject one more); inventory (all registered rejection checks and independent column boundaries).

Tools/Test-OfaRepairsExcelAudit.ps1 independently transcribes the inspected master insertion sequence into private Excel, with macros/events/links/calculation disabled. It compares formulas/constants/array flags, number formats/locked state, affected names and Excel SaveCopyAs/reopen results. The AllowVerifiedIncludeRepair switch separately asserts/reports the approved RepIncStkCat correction, without suppressing other differences. It does not execute workbook macros, TDB sync or financial calculation in Excel.

Baseline 2.79 matrix: all 36 simple-row +3/-3 cases completed; 35 passed and Economic service-charge failed as above. Capital Expenditure, Capital Grant, Housing Components, Stock Conversion and Journal each passed blank +1/-1 and +3/-3 inverse checks. This demonstrates why populated and independent-reference tests are necessary too.

An initial Repairs fixture expected a relative A1 name rather than an absolute reference. Its sole mismatch was dollar-sign spelling in the expectation. The fixture was corrected and rerun, not treated as a product pass without retesting.

Logs: obj/structural-audit279-* and obj/structural280*. Private outputs are not repository deliverables. Full payloads, passwords and raw VBA are excluded. VBA preservation compares module identities/source hashes, not binary zip equality.

## Delivery checks and remaining acceptance

**Ready to test — 2.80.** Debug and Release compile successfully. Final validation:

- AGL OFA and Repairs: one and three records added/deleted; formula/constant restoration across affected sheets and TDB; all 1,762 names restored, except the separately asserted approved Include correction. Repaired record counts, mirrors and entry worksheet/calculation state pass.
- Capital Expenditure populated-marker deletion: one and three records pass. The new input disappears, not the hidden template.
- Actual maximum deletions to the VBA minimum, followed by rejection of one more: OFA, Repairs, Capital Expenditure, Capital Grant, Housing Components, CPI/RPI, Cash Journals, Journal and Stock Conversion all pass.
- Explicit below-minimum attempts for OFA, Repairs, Capital Expenditure, CPI/RPI, Cash Journals, Journal and Stock Conversion are rejected. All 62 registered rules pass inventory rejection checks. The five reviewed column families also match independently derived VBA limits; protected-leading selected deletion is checked in the published builds.
- Economic service-charge +1/-1 and +3/-3 now restore all names and affected/TDB cells. Service Charge linked +3/-3 and Journal +3/-3 regressions pass, retaining exact-source and numbered-alias mirrors.
- Authoritative Blank copies: OFA +1/-1 and Repairs +1/-1 pass. Demo copies: Repairs +1/-1 and populated Capital Expenditure +1/-1 pass. No master changed.
- Published Debug malformed Repairs Include-row test: rejected before insertion; all names, 168,563 affected/TDB formula/constant cells, record count, worksheet state and calculation settings unchanged.

Strict independent Excel comparisons (case-sensitive formula/constant text, array flags, number formats and cell locking):

| AGL insert | Cells compared | Unexpected cell/format/name differences against Excel VBA sequence | Excel save/reopen names preserved |
| --- | ---: | --- | ---: |
| OFA +1 | 3,980 | 0 | 1,762 |
| OFA +3 | 4,426 | 0 | 1,762 |
| Repairs +1 | 22,266 | 0, plus one separately verified Include-name correction | 1,762 |
| Repairs +3 | 22,468 | 0, plus one separately verified Include-name correction | 1,762 |

All 335 AGL VBA module identities/source hashes remain identical through the four native expanded results and their Excel saved copies. Aggregate source-hash identity: 1539cc5c003414274742dda616da675ebf68b420e4143bd1ca5a7e17d68fa470. This proves source preservation, not VBA runtime execution.

These cases complement the earlier linked-family/Funding/Development trials in Structural_Repair_Trial_2026-09-22.md; they do not exhaust every arbitrary selection, bespoke schema or legacy VBA entry point. Earlier failing baseline runs and the superseded diagnostic address-format failure remain labelled in local logs, not counted as final passes.

Still required: client DIT selection/add/delete and new-record visibility; financial Check Sheet/SOCI/SOFP review; trusted Excel/VBA calculate/save followed by Summit reopen. Macro-disabled Excel operations do not close those acceptance gates. No release-wide performance claim is made from concurrently running diagnostics.
