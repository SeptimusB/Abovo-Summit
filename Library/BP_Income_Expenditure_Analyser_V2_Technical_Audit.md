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

## Test version 1.92: live datasource and description width

In the sparse BP v26_0001 model, navigation away from and back to Analysis V2
recreated its `RangeDataSource` and showed current data, while both
`GridControl.RefreshDataSource` and `GridView.RefreshData` left the existing
live binding stale after edits. Post-calculation refresh now disconnects and
recreates the live or comparison datasource without another workbook calculation,
then restores the existing grid state. Snapshot mode remains fixed. The normal
calculation service has already completed its ordinary and deferred Transactional
DB passes before this refresh callback.

The description column now measures the workbook-backed group captions as well
as ordinary data cells after binding, including custom-drawn totals. It does
not shrink an existing wider column. The workbook was not modified. Debug and
Release x64 builds passed; the live Summit edit, snapshot, selection-state and
multi-DPI checks still require manual validation.

## Test version 1.93: dependency-sensitive live refresh

The 1.92 manual test confirmed the description width correction but showed
that recreating the datasource alone did not update calculated values. That
observation rules out a binding-only fix: the normal chain/deferred worksheet
pass may still leave cross-sheet XLSB formula values stale. On live/comparison
refresh, V2 now performs the existing recursive dependency-sensitive workbook
calculation before disconnecting and rebinding its datasource. That recursive
pass temporarily permits the custom calculation service to calculate the
Transactional DB, comparison and check sheets, restoring both the service flag
and prior engine afterward. Snapshot mode remains unchanged. The workbook was
not modified. Debug and Release x64 builds passed. A populated-workbook edit
with Analysis V2 open must confirm updated values and measure the added time;
expanded groups and snapshot behavior also need a manual check.

## Test version 1.94: one-time calculation-chain rebuild

The live 1.93 test confirmed correctness but found a greater-than-five-second
pause on every edit. A read-only, unsaved DevExpress probe of the repository
Demo master reproduced the stale output: changing `TransRents` from 87.5 to
90 left `Transactional_Records` General Needs year one at 24,934 after an
ordinary chain calculation. After one `CalculateFullRebuild` in ChainBased
mode, the same edit changed that output to 25,646.4. With Summit's actual
deferred-sheet service registered, the rebuild took about 9 seconds and later
staged rent edits took about 0.6-0.8 seconds. A separate unsaved Blank-master
probe rebuilt its chain, then populated one dummy stock category only in
memory: staged calculation changed the output from zero to 25,025 in about
1.2 seconds; the following rent edit changed it to 25,740 in about 0.6 seconds.
No workbook file or persistent dummy row was changed.

`CalculateDependencySensitiveFile` now builds the ChainBased dependency tree
once before the initial analyser binding, temporarily including Transactional
DB and Check Sheet in that rebuild. Ordinary calculations leave this graph
prepared, use the established deferred-sheet pass, and rebind the live analyser
without another full calculation. Bulk structural mutation invalidates the
graph; the next analyser binding rebuilds it. The previous engine and custom
service setting are restored in `Finally`. Debug and Release x64 builds passed.
When an analyser is reopened after edits made with only one DIT worksheet active,
its existing graph is retained but an ordinary-plus-deferred calculation runs
before rebinding, so those edits are not hidden behind stale cached values.
Manual Summit testing must verify a populated rent edit with Analysis V2 open,
first stock population from a blank model, rebind/expansion retention, snapshot
mode, and an add-lines structural change followed by another edit. The one-time
rebuild cost at first analyser open is an expected trade-off to measure.

## Test version 1.95: DIT navigation registration and timings

Static review found that a revisited DataInterfaceTemplate registered a new
calculation-engine object without retaining its returned ID or re-registering
its worksheet. Deactivation then removed the original ID, leaving the revisited
object active. The engine also appended new object and worksheet slots rather
than reusing vacated slots. Reactivation now removes any stale registration by
object identity, retains the new ID, and registers its worksheet without an
extra calculation. Deactivation and resource release remove by object identity;
the engine tolerates objects without worksheets, avoids duplicate worksheet
registration, and reuses vacant slots.

The existing workbook calculation on navigation is deliberately unchanged.
Trace output prefixed `[Navigation Benchmark]` records ordinary and deferred
calculation, interface refresh, DIT section build/refresh, active registration
counts, allocated slots, and failures. Debug and Release builds passed. No
workbook was changed. Manual testing should repeat A -> B -> A navigation,
including with Analysis V2 open, and verify stable object/worksheet slot counts,
correct post-navigation edits and analyser values, and no loss of interface
state. The timings will establish whether navigation calculation can safely be
made conditional in a separate trial.

## Test version 1.96: per-model navigation calculation gate

The populated-model navigation measurements showed that a no-edit DIT return
still paid for ordinary and deferred workbook calculation, and paid again for
live analyser refresh when Analysis V2 was open. Navigation now skips that
workbook pass only when a per-model generation indicates that a successful
whole-workbook calculation already covers the latest known change. The DIT
still rebuilds/refreshes its visible section after reactivation, so this does
not replace interface rendering with a cached control image.

ChangeManager edits, undo and redo continue through CalculateWSs, which marks
the model's navigation calculation pending before any calculation. Direct
imports and model cell writes using SetDirtyFlag, bulk structural operations,
and Stress Test calculation/mode writes also invalidate the marker. A
successful CalcFile or dependency-sensitive whole-workbook pass certifies it
only when no intervening mutation occurred and no bulk mutation remains in
progress. A single-worksheet calculation never certifies the workbook; the
next DIT return performs the full staged pass. This gate does not alter edit
calculation, the prepared dependency graph, or analyser datasource semantics.
No XLSB was modified.

Debug and Release builds passed for Any CPU and x64. Manual testing is still
required: populated-model A -> B -> A navigation without edits, first with
Analysis V2 closed and then open, should log "CalcFile skipped" on the return
and preserve current figures; an ordinary DIT edit followed by navigation
must recalculate when needed; verify a blank-model first population, undo/redo,
multi-model independence, and structural add-lines followed by Analysis V2.
Population/edit timings should remain the main performance criterion; this
trial targets only unnecessary no-edit navigation work.

## Test version 1.97: avoid false change on interface registration

The client's 1.96 log confirmed the no-edit gate could reduce a DIT return to
7-17 ms, but also showed 767, 1,833 and 845 ms full passes on other returns
with advancing model generations. Static tracing identified one non-edit
source of those increments: constructing a DIT calls AddActiveWorksheet with
CalculateNow=True, which invokes CalculateWSs. The 1.96 marker treated every
such call as a workbook mutation even though registering a visible worksheet
does not write it. AddActiveWorksheet now retains its initial calculation but
passes InvalidateNavigation=False. ChangeManager, explicit CalculateWSs calls,
direct-write signals and structural invalidation still use the default True
path. The initial per-model unknown state still requires a full pass before
any no-edit navigation can be skipped.

The pasted run also includes a 4,672 ms CSID 43 return, with 3,820 ms in the
ordinary calculation phase. The log alone does not establish whether this
was a genuine edit or another invalidation. A further no-edit A -> B -> A
trial should establish whether the registration correction removes spurious
generations. Confirm calculated values after a real edit, undo/redo and a
structural change before treating the gate as validated.

## Test version 1.98: population timing capture

The client's paired solo-interface and Analysis V2 navigation traces showed
fast no-edit returns with the analyser open, but lacked calculation timings
for the actual edits and the five-line Cash Journals insertion. This test
version instruments those missing boundaries without changing workbook
calculation or structural behavior. ChangeManager logs each posted edit's
setup, typed write, calculation, journal/refresh and total time. The calculation
engine separately reports a single active worksheet calculation or a staged
whole-workbook pass, including ordinary, deferred Transactional DB and
interface refresh phases. The structure manager logs workbook insertion,
Transactional DB synchronisation and dependency invalidation. DIT logs the
structural action, section rebuild, fonts/rules and total add-lines time.
All new records use the `[Population Benchmark]` prefix and include model,
worksheet or rule rather than cell contents.

Debug and Release builds passed for both Any CPU and x64. The initial
standard Debug attempt could not copy its output while a prior Summit test
process held the executable open; it passed after that process exited.
The manual comparison should use the same unchanged workbook baseline twice:
Stock -> Rents -> Stock, edit, Rents, edit, Stock, Cash Journals, add five
lines; repeat with Analysis V2 already open. Check displayed analyser values
as well as timings, and keep the analyser's first-load cost separate.
