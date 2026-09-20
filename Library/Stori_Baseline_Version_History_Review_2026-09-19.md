# Stori: candidate baseline and Version History evidence

Date: 19 September 2026. Scope: read-only workbook inspection; no migration, recalculation or application-code changes.

## Conclusion

`D:\Downloads\Master BP v25_0700 (1).xlsb` is a strong **provisional baseline** for investigating Stori. Its populated Version History cell values match Stori's exactly; Stori adds only the four entries at rows 694-697. Both workbooks have the same 279 worksheet names in the same order. Internal `ModelVersion` values agree with the releases in their filenames.

This supports beginning a read-only baseline/client comparison without waiting for a separate generic 25.0704 file. It does not prove exact file ancestry, that all changes were logged, or whether each recorded change was generic or commissioned for Stori. An exact byte-for-byte ancestor is not asserted. Production bespoke migration still requires verified mappings and approval of the changes to preserve.

## Inputs and preservation

| Role | Original file | Internal release |
|---|---|---|
| Candidate old base | `D:\Downloads\Master BP v25_0700 (1).xlsb` | 25.0700 |
| Populated client | `D:\Downloads\Stori BP v25_0704 2026-27 V4 update 100826 CLEAN pass 3.xlsb` | 25.0704 |

`ModelVersion` resolves to `Global Assumptions!C26` in both files. The master stores numeric 25.07, displayed as 25.0700; trailing zeros must not be treated as a different release identifier.

Original SHA-256 hashes, verified unchanged after inspection:

- Master: `21E1E4B032821A9864046F1C1493819198417D2158D588019F626C9F55BA59B8`
- Stori: `FD618BC0E3EF71C328C31AFFF90F10EEF1731415BD14A24484B111035C4693E3`

Inspection used exact disposable copies in an isolated hidden Excel instance. Macros and events were disabled, external-link updates disabled, and calculation set to manual before opening either copy read-only. Both copies were closed without saving; copy hashes also remained identical. No VBA source was extracted. Cached values were inspected, not financially validated.

## Version History findings

Inspected `Version History!A1:H714` in both files. Comparing populated cells, including row positions and values, finds exactly four additions in Stori and no other differences. Row 714 is a model-version-range marker, not a release entry.

The shared latest entry is row 693, release 25.0700, dated 29 May 2025 at 15:30. It records changes to the housing-revaluation SOCI heading on Hidden - SOCI Structure and corresponding Transactional DB headings.

| Stori row | Release / recorded date | Recorded work, summarised | Migration investigation |
|---|---|---|---|
| 694 | 25.0701 / 20 Aug 2025, 11:43 | Add six Stock Assumptions columns and update Trans DB. | Establish the stock master/dependent expansion and distinguish generated growth from bespoke logic. |
| 695 | 25.0702 / 20 Aug 2025, 15:59 | Add 140 Management Costs Assumptions columns and update Trans DB. | Establish the management-cost structural family and formula continuity; do not copy this geometry independently into each overlapping range. |
| 696 | 25.0703 / 20 Aug 2025, 16:14 | Add Management Cost Drivers and first-nonzero-income workings to Other Income Workings; add Existing Stock Grouping inputs to Global Assumptions and a Stock Grouping dropdown to Stock Assumptions; update the management-cost Drivers dropdown and Variable Management Costs references. | A linked inputs/validation/formula change, not just capacity. Check which equivalents exist in the latest template and map any genuinely missing client requirements. |
| 697 | 25.0704 / 21 Aug 2025, 10:58 | Update covenant definitions and their stress-test use; change Live Multivariable Planner conditional formatting in AC:AE; amend OW - Charts Source Data D67. | Review formulas, covenant semantics, stress-test consumers and formatting together. History does not give the old/new formulas or intended numerical effect. |

These are documented deltas, not automatically approved patches. In particular, copying old covenant or management-cost formulas over newer generic logic would be unsafe without checking the latest template.

## Supporting structural observations

The master has 1,756 defined names; Stori has 1,748. This review did not inventory the eight-name net difference or determine whether it represents additions, deletions, scope changes or repairs.

Selected named-range physical dimensions differ:

| Name | Master | Stori |
|---|---|---|
| Transactional_Records | 1,594 rows x 74 columns | 2,393 rows x 74 columns |
| StockCondCats | 1 x 13 | 1 x 26 |
| HouseTypeInID | 1 x 12 | 1 x 24 |
| DepnType | 1 x 5 | 1 x 9 |

These counts include physical template capacity; they are not asserted logical populated-record counts. Other inspected names, including Funders, Facility, FacilityNames, IR_Cash_Journals and Rep_Jour_01, have matching dimensions. Additional growth may be ordinary client population and is not proof of bespoke changes or an incomplete history.

## Next evidence steps

1. Read-only comparison of the four documented changes against actual names, formulas, validation and dependent ranges; normalise approved structural expansion before classifying differences.
2. Compare those requirements with the confirmed latest target template. Mark already-incorporated changes, mapped equivalents, missing features and genuine conflicts separately. Confirm whether the target is the Sandbox New Blank or repository Blank before a real migration.
3. Review the remaining unexplained differences with Abovo. Only then compile approved structural/input/patch operations for an isolated result copy, followed by calculation, Check Sheet, detailed SOCI reconciliation and Excel/VBA round-trip tests.

No original workbook, repository master or production code was modified. This note supplements `Upgrade_Migration_Architecture_Review_2026-09-19.md`; it does not establish financial acceptance or production migration safety.

## Subsequent three-way review

See [Stori three-way comparison](Stori_Three_Way_Comparison_2026-09-19.md) for verification of the four changes, latest-template inspection and differences outside the history entries. That review resolves the net-eight-name reduction as one business-name addition and nine hidden function-placeholder removals. It also establishes that the later generic branch reuses Stori's 25.0701-25.0704 labels for different work; version numbers must not be used as unique ancestry identifiers.
