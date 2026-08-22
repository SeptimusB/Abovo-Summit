# BP Income and Expenditure Analyser V2 Technical Audit

## Delivery

Version 7.22 preserves BPIncomeExpenditureAnalyser as Analysis V1 and adds an
independent BPIncomeExpenditureAnalyserV2 source, designer and resource set as
Analysis V2. Both are worksheet-free special Outputs children under the
Analysis navigation group. V1 was not modified.

V2 remains a read-only projection of the authoritative
Transactional DB!Transactional_Records named range. The Transactional DB
synchronizer disconnects and reconnects both open analyser versions around
structural range changes.

## V2 corrections

- Balance Sheet grouping now targets the Balance Sheet view and uses the
  workbook fields OrderedBSGroup and OrderedBSHeading.
- SOCI, Cashflow and Balance Sheet grouping, summaries and exports resolve
  required columns by field name rather than fixed helper-column ordinals.
- Period columns are discovered from workbook year headings; the Balance Sheet
  includes every forecast year and separately identifies its opening balance.
- The active-grid selector distinguishes all three tabs.
- Balance Sheet export is available, export field/count offsets are corrected,
  and export failures are reported rather than suppressed.
- Duplicate paint-handler registrations, per-paint Font/Pen leaks, null summary
  dereferences and the unmatched paint-event BeginUpdate/EndUpdate path are
  removed.
- Group traversal stops at the first invalid DevExpress group-row handle rather
  than assuming a maximum of 10,000 groups.
- Each grid explicitly enforces read-only behavior, cell multiselect and
  clipboard copy.
- Disposal disconnects the RangeDataSource, unregisters the calculation object
  and clears the model's V2 reference before child controls are disposed.
- Column type detection samples the live named range rather than assigning
  types from fixed column numbers.

## Automated validation

- Structure.xml contains 34 unique Outputs children with Analysis V1 at
  CSID 32 and Analysis V2 at CSID 33.
- Neither Analysis child defines an interface worksheet.
- The V2 source, designer and resource are compiled/embedded independently.
- Full standard Debug and Release rebuilds passed on 22 August 2026 with zero
  warnings and zero errors.

## Required manual validation

- Open Analysis V1 and Analysis V2 and confirm they can be viewed independently.
- In V2, switch among SOCI, Cashflow and Balance Sheet, then exercise Expand All
  and Collapse All on each tab.
- Confirm period headings, opening balance, totals, red negatives, cell
  multiselect and clipboard copy against the XLSB.
- Export each V2 tab and confirm the selected statement is the default export.
- With V2 open, perform a supported Transactional DB structural update and
  confirm the range reconnects without a binding exception.
- Close and reopen V2 and confirm calculations do not retain a disposed
  analyser instance.

## Version 7.23 exclusive range ownership

Transactional DB!Transactional_Records cannot back two independent live
DevExpress RangeDataSource instances. Each ExcelModel now owns a general
ModelResourceRegistry whose keyed entries have exactly one owner and an ordered
release callback.

Analysis V1 and V2 both claim the Transactional_Records RangeDataSource key.
Before either analyser is constructed, the existing owner is disconnected from
its grids, removed from the calculation engine by object identity, removed from
the document manager, disposed and cleared from the model. V2 also unregisters
itself during its normal Dispose path. Model shutdown releases all registered
resources before workbook controls and services are torn down.

Required switching test: open V1, then V2, then V1 again. At each transition,
confirm the previous document closes and the replacement opens without the
DevExpress already-associated range binding exception.

The registry behavior test passed exclusive registration, duplicate rejection,
ordered callback release and release-all. Full standard Debug and Release
rebuilds passed on 22 August 2026 with zero warnings and zero errors.

## Version 7.24 period discovery correction

The first V2 runtime test exposed a stripped regular-expression escape in
GetPeriodColumns: the expression tested for literal d characters and therefore
reported that Transactional_Records had no period columns. V2 now recognizes
four-digit/two-digit year headings without escape characters and checks the
generated field name, customization caption and visible caption. If DevExpress
normalizes all three, it falls back to the authoritative first row of the
Transactional_Records range and maps matching source positions to grid columns.
The built matcher behavior test recognized direct and multiline workbook years
and rejected a non-year caption. Full standard Debug and Release rebuilds
passed on 22 August 2026 with zero warnings and zero errors.

## Version 7.25 group-summary conversion correction

The next V2 runtime test exposed legacy group-summary values that DevExpress
materialised as blank or non-numeric strings. Every active expansion, styling
and custom-draw path now obtains summary integers through one guarded parser.
It accepts numeric objects, current/invariant-culture numeric strings and
parenthesised negatives, while rejecting blanks, text, non-finite values and
integer overflow without throwing. This removes the opening FormatException
without changing V1 or the Transactional_Records workbook range.
The focused parser test covered numeric text, blank text, non-numeric text,
numeric objects and parenthesised negatives. Full standard Debug and Release
rebuilds passed on 22 August 2026 with zero warnings and zero errors.
