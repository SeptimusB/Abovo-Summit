# SpreadsheetGear: complete Funding insertion trial

24 September 2026. Isolated research checkpoint, not a Summit test release.

## Decision

**Do not integrate the raw SpreadsheetGear save/reload path into Summit.** It is
fast, but this trial fails two independent compatibility gates: Excel grouped-sheet
3-D reference semantics and saved-package preservation. AutoFill does not resolve
either problem. Production code, version, references, original AGL and repository
masters were not changed by this work. No commit or push was requested or made.

SpreadsheetGear remains a promising calculation engine. These results do not negate
the earlier calculation-only parity result, but they rule out calling this a
low-risk drop-in structural helper without additional compatibility work.

## Method and authority

The user approved the bounded full Funding experiment, isolated VBA execution and
temporary unprotection of test copies. Each native run starts from identical XLSM
bytes. All experimental saves are new files beneath the ignored private trial
directory; the originals are never opened for writing.

Common input: the earlier Excel-converted copy of
`C:/Sandbox/Insert Comp/AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb`.

- Original XLSB SHA256: `30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47`.
- Common XLSM SHA256: `62B44DE2764C4D1D1C9C7029C9297EEF4D461E300BAB4AD7DB092BBE1D23193C`.
- Both hashes, and the new input-copy hash, were unchanged after testing.
- x64 Release; SpreadsheetGear 9.3.85.102; Excel 16.0 build 20326; DevExpress 25.2.4.
- Same Core Ultra 9 285K machine as the preceding trial. Not minimum-spec results.

The native port reads the reviewed 32-sheet target list from Summit's current rule.
It inserts ten whole columns before the ordinary-loan template on all 32 sheets,
then copies the shifted template. It expands all nine facility and two ordinary-loan
Transactional DB mirrors, following the VBA's whole-row insertion and FillDefault
sequence. Source and mirror dimensions are asserted. SpreadsheetGear restores entry
worksheet protection settings and does not change sheet visibility.

Excel executes the **actual unmodified embedded** `Run_Insert_Funding_Columns` and
`FundingSynchroniseTransactionalDBSheet` procedures. Only the outer InputBox and
completion-message wrapper is omitted. Each Excel instance is checked against the
pre-existing process list, opens the private input read-only, disables events and
link updates, and uses ByUI macro policy. No Trust Center setting is changed. The
VBA source was inspected in memory; neither it nor credentials are persisted here.

An explicit full rebuild follows each operation: the VBA's historical ABVCalculate
early exit is not treated as a completed calculation. SaveCopyAs writes a separate
Excel result. The two pre-existing Excel processes were left alone; trial instances
closed normally.

## Timings: useful speed, invalid structural result

Seconds, rounded. SpreadsheetGear uses two fresh processes per variant in balanced
Copy / AutoFill / AutoFill / Copy order; the table uses the median of each pair.
Excel is one full-command reference pass, not a repeated median. No competing trial
inspection/build workload ran during the accepted timing passes. OS caches were not
flushed. Process creation/licence activation is outside the load timer.

| Stage | SpreadsheetGear Copy | SpreadsheetGear AutoFill | Excel embedded VBA |
| --- | ---: | ---: | ---: |
| Load common XLSM | 4.10 | 4.05 | 12.89 |
| Complete native Funding command, including mirrors/protection | 1.51 | 1.49 | 13.39 |
| Full rebuild | 0.77 | 0.78 | 4.32 |
| Save new XLSM | 1.41 | 1.41 | 6.22 |

SpreadsheetGear's copy/fill portion across all 32 sheets was about 31 ms for Copy
and 33 ms for AutoFill. **No worthwhile AutoFill improvement was demonstrated**
for the full-column operation. Its earlier small-range microbenchmark advantage
must not be extrapolated to this command. Roughly 0.9 seconds of this prototype's
command time is unprotection; the preservation guards must not simply be removed.

Independent full manifests for one Copy/AutoFill pair match every formula,
constant, name, array definition, worksheet state and audited populated-cell
appearance fingerprint; package feature counts also match. Both variants share
the same compatibility failures against the Excel result.

A separate native DevExpress reload of the inserted SpreadsheetGear XLSM took
**28.39 seconds**. Adding that observation to the Copy medians gives about
**36.2 seconds before Summit's initial staging save, real UI rebind and any added
compatibility work**. This is an illustrative sum of measured stages, not a timed
end-to-end Summit operation or an approved output. The prototype was deliberately
not connected to Summit's live model or UI after the compatibility gate failed.

No new full DevExpress or Aspose Funding command was measured. The older source-only
and historical x86 timings are not directly comparable with this x64 full command.

## Gate 1: grouped-sheet references fail in the serial port

The SpreadsheetGear result has **12,050 formula-text differences across nine sheets**
against the Excel/VBA result. The differences are unadjusted 3-D references, not
missing Transactional DB dimensions. For example:

```text
Loan Closing Balances!AC8
Excel: =SUM('Loan Opening Balances:Loan Repayments'!AC8)
Gear:  =SUM('Loan Opening Balances:Loan Repayments'!S8)
```

Differences by sheet: Hidden - Funds Available 250; Interest Rates 2,750;
Hidden - Loan Repayments 2,750; Hidden - Loan Interest Charge 2,750;
Hidden - Loan Commitment Fees 2,750; Loan Closing Balances 600; Cashflow 80;
Cashflow detailed 80; Total Cash Trial Balance 40.

Both full outputs contain 1,126,411 formulas, with matching sheet order, visibility
and protected/unprotected states. All scoped names match except the same three
Gear compatibility names listed below. The independent full-command appearance
comparison is recorded separately from the save-only control.

That comparison also finds 19,015 fill-pattern differences, including 1,968
Solid/non-Solid changes, and smaller font/colour differences across 1,232,586
populated cells. Sample patterns are red hatching in the financially incorrect
result and red Check headings. Because calculated states differ, these must not
be attributed solely to serialization or treated as an independent formatting
diagnosis. Importantly, **the passing save-only fill check does not clear the
full insertion's fill-based editability**.

The five calculated probes contain 998 numeric differences from Excel, plus one
blank/text-type and six text differences. These are not accountant-level rounding
differences. Check Sheet row 37 changes from `OK` to `Check`; the warning then flows
into worksheet headings. Copy and AutoFill repeat the same result.

Comparisons use absolute tolerance 0.000001 or relative tolerance 1e-10, whichever
is greater. The AGL insertion is not expected to be financially neutral: its AC
source/template column contains 13 non-zero numeric inputs, which the actual VBA
copies into the inserted columns. The Excel reference therefore differs from its
pre-insert baseline at 847 numeric probe positions while the Check Sheet remains
unchanged. This is an observed property of the supplied file and macro, not a new
instruction to clear those inputs. The comparison oracle is the same complete
Excel operation, not an assumption that inserting populated templates changes no
figures; financial approval remains with Abovo.

Excel independently opens and fully recalculates the saved Gear result and obtains
the **same wrong probe values** as Gear: all 38,109 populated probe positions match.
Thus a later Excel CalculateFullRebuild does not repair the structural damage.

A customer-data-free reproducer establishes the important distinction:

| Operation on First and Last sheets | Result formula | Calculated value |
| --- | --- | ---: |
| SpreadsheetGear serial insert B:C on each sheet | `=SUM(First:Last!B1)` | 0 |
| Excel serial insert B:C on each sheet | `=SUM(First:Last!B1)` | 0 |
| Excel grouped insert B:C on both sheets | `=SUM(First:Last!D1)` | 12 |

The starting B1 values are 5 and 7. This is therefore a mismatch between the port's
serial operation and the VBA's grouped semantics, **not evidence that SpreadsheetGear
alone is defective**. A supported grouping mechanism or an explicitly designed
3-D reference-preservation adapter would be required. Summit already has such a
guard for DevExpress; it cannot be assumed to apply to another engine unchanged.

## Gate 2: save-only package preservation fails

A separate control performs **no insertion**, only load, full rebuild and new XLSM save.
Independent DevExpress and ZIP/XML inspection found:

- All 1,077,821 cell formulas retained their text; all 283 worksheets retained order,
  visibility and protected/unprotected state.
- The VBA project binary is byte-identical.
- Selected formatting and locks match across **1,182,776 populated cells**, including
  zero changes to the Solid/non-Solid fill gate. This does not cover blank-cell styles.
- All three custom XML payloads, their properties and relationships are absent from
  the saved output: nine package parts before, zero afterwards.
- `xl/metadata.xml` and cell metadata links are absent. Independent DevExpress reads
  the original 64,820 dynamic arrays as legacy arrays after saving (98 legacy before,
  64,918 after). Formula text alone does not prove dynamic-array edit behaviour survives.
- Three compatibility names are added: `_xlfn.IFERROR`, `_xlfn.IFNA`, `_xlfn.SUMIFS`.
  One Version History constants fingerprint differs; it remains unclassified.
- Conditional-format rule/container counts and drawing/chart-related package parts
  change. These are review findings, not proof that every such difference is harmful.
  In particular a missing relationship part is not a count of missing charts.

Excel opens and fully recalculates this save-only control with all 38,108 populated
probe positions matching the earlier Excel baseline. Good present-day results do
not clear the missing structure XML or changed array semantics.

The vendor's [published limitations](https://www.spreadsheetgear.com/support/help/spreadsheetgear.net.9.0/SpreadsheetGear_2023_Limitations.html)
also identify lack of dynamic-array support. No client workbook was uploaded, no
vendor ticket sent, and no blind metadata/XML transplant was attempted.

## Validation boundaries and next decision

The trial harness builds Debug and Release for x86 and x64 with zero warnings/errors.
Each build passes 19 SpreadsheetGear and 13 licensed Aspose synthetic smoke checks.
Production Summit was neither rebuilt nor version-bumped for this research.

The native speed is promising, but **integration is held at the preservation gate**.
Live-model staging, transactional rollback, actual interface rebinding and the
full Blank/Demo/Stori matrix remain unimplemented/unmeasured. No output from this
trial is approved for client use. These explicit gaps are preferable to reporting
an invalid result as a successful end-to-end speedup.

Lowest-change direction: keep DevExpress as the current authority and investigate
a narrow Excel/VBA helper on compatible installations, with safe save/reload,
licensing/policy detection, rollback and fallback. Gear could still be trialled as
an in-memory calculation helper that does not rewrite authoritative packages, or
with a separately approved preservation layer. Neither is implemented here.

## Reproduction and evidence

Harness: `Tools/SpreadsheetEngineTrial/FundingTrial.cs`, `FundingThreeDProbe.cs`,
`Audit-FundingPackages.py`, existing native adapters and comparison tools.

Private evidence directory:
`obj/AsposeTrial/funding-d8bcacf09cdf4b308742ce2d4bf6645d`.

- Accepted timing passes: `gear-copy-3/4.json`, `gear-fill-3/4.json`, `excel-1.json`.
- Earlier passes 1/2 are pilot/repeat evidence, superseded for the timing table after
  eliminating repeated reading of the existing credential field in the harness.
- `synthetic-3d.json`, `formula-differences.json`, `final-probes.json`.
- `baseline-audit.json`, `style-differences.json`, `input-manifest.json`,
  `gear-baseline-manifest.json`, `dx-reload.json`.
- `gear-recalc-parity.json`, `baseline-excel-recalc-parity.json` and paired private
  Excel-reopened files. The explicit round-trip comparison flag permits their
  deliberately different input hashes; same-input checks remain the default.

Additional complete-output evidence: `full-command-audit.json`, `copy-fill-audit.json`,
the three full-command manifests, `full-command-style-differences.json` and
`excel-insert-versus-baseline.json`.

See the harness README for commands. Never distribute the experimental workbooks.
