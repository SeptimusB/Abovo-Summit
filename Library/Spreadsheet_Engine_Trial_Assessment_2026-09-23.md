# Spreadsheet engine speed investigation

23 September 2026. Aspose.Cells 26.9.0 is installed and licensed for an isolated trial. Initial real-model I/O and preservation testing is recorded below. It is **not integrated into Summit** and is not approved for client workbook output.

## Recommendation: minimise change

Keep DevExpress 25.2.4, the existing UI, workbook ownership, history and structural services for this client release. A wholesale engine substitution is not a DLL swap: production code binds DevExpress IWorkbook, Worksheet, CellRange, RangeDataSource and native controls directly.

Trial **commercial Aspose.Cells for .NET first**, in a separate console executable against disposable copies. Do not buy or alter Summit on the strength of vendor performance claims. A successful round-trip and a material end-to-end speed improvement must precede an integration proposal. The freely licensed Aspose.Cells FOSS product is a different product with unsuitable format/runtime constraints for this XLSB/.NET Framework application.

## Shortlist and evidence

| Library | Verified published capability | Decision for Summit |
| --- | --- | --- |
| Aspose.Cells for .NET | [XLSB and XLSM read/write](https://docs.aspose.com/cells/net/supported-file-formats/), [.NET Framework 4.8](https://docs.aspose.com/cells/net/system-requirements/), [VBA/macro preservation, formulas, names and structural operations](https://docs.aspose.com/cells/net/feature-overview/), [custom calculation engine callbacks](https://reference.aspose.com/cells/net/aspose.cells/abstractcalculationengine/) | Strongest first trial candidate; compatibility with the real Abovo models and speed remain unproved. |
| GemBox.Spreadsheet | [XLSB read/write](https://www.gemboxsoftware.com/spreadsheet/docs/supported-file-formats.html), [XLSB unsupported-feature preservation](https://www.gemboxsoftware.com/spreadsheet/docs/preservation.html), [.NET Framework 4.6.2+](https://www.gemboxsoftware.com/spreadsheet) | Reserve candidate. Its [VBA object-model example](https://www.gemboxsoftware.com/spreadsheet/examples/c-sharp-vb-net-excel-vba-macros/124) explicitly says XLSM only. That does not establish whether opaque XLSB VBA survives preservation. Resolve with vendor evidence and a private round-trip before a performance trial. |
| Syncfusion XlsIO | [Documentation lists XLSB as limited support](https://help.syncfusion.com/document-processing/excel/excel-library/net/overview) | Defer until the vendor confirms the precise XLSB read/write, formula and VBA preservation limitations. |
| SpreadsheetGear | [Vendor FAQ says XLSB reading/writing is not supported](https://www.spreadsheetgear.com/support/knowledge-base/faq) | Not the minimal-change choice: would introduce a format-conversion boundary. |

Aspose offers a [30-day temporary licence](https://docs.aspose.com/cells/net/evaluation-version-limitations/). Use that rather than its unlicensed mode, which inserts an evaluation worksheet. An added vendor sheet would invalidate a clean preservation comparison and could affect 3-D references. No claim is made here about relative pricing or the suitable redistribution licence.

## Trial boundaries

1. **I/O first:** identical frozen copies of current Blank/Demo, populated AGL, and older Stori. Measure load and save separately without calculation. Compare a no-op round-trip before considering edits. Never upload client workbooks to a vendor without approval.
2. **Preservation:** formulas and formula structures (including long CONCATENATE and dynamic-array metadata), scoped/multi-area names, 3-D spans/order, typed inputs, protection/validation, VBA module source identities, document properties and all existing custom XML. Open in Excel without repair, then reopen in Summit. Compare known relevant outputs and distinguish existing intentional errors from new ones.
3. **Calculation:** reproduce PMCost and RespCost through the candidate's custom-function API, then measure dirty-sheet, full and rebuild calculation. Preserving embedded VBA is not evidence that its UDFs will execute in another .NET engine. Financial parity is a separate gate.
4. **Structural work last:** port only the existing verified Funding +10 and Development +10 commands into the experimental harness. Generic InsertColumns does not replace the cross-sheet, mirror-range and template rules. Check add/delete inverses and output integrity outside the timed operation.
5. **End-to-end comparison:** at least three fresh-process runs per case, separate cold and warm observations, report median/range and peak memory. Compare like-for-like architecture and calculation settings. An optional 64-bit helper would need separate measurement, not be conflated with an engine improvement.

For a later file-based helper, the meaningful time is **Summit staging save + candidate load + operation + candidate save + Summit reopen/rebind**. A faster engine operation alone may still produce a slower user experience. Do not change the shipping implementation unless this total improves enough to justify the additional boundary and tests.

## Installation checkpoint - 23 September 2026

User authorised proceeding with installation. Added `Tools/SpreadsheetEngineTrial`, a standalone
.NET Framework 4.8 console project, with an exact Aspose.Cells 26.9.0 package reference and
NuGet lock file. Restore uses the ignored `packages/engine-trial` folder. Summit's project,
solution, dependency references, executable, settings and 2.92 test-release version are unchanged
by this trial. No workbook, licence, VBA source or password was added to the repository.

Debug and Release builds each pass at actual 32-bit and 64-bit process widths, with zero
compiler warnings/errors. Each run passes 13 installation assertions: process width;
cross-sheet decimal calculation; recalculation after an input change; and, for both XLSB
and XLSM, in-memory serialization, formula/input/date retention and reopened calculation.
Architecture-specific compiler intermediates prevent incremental builds from reusing a
different architecture's executable. These are synthetic in-memory tests, not workbook
artifacts or financial/Excel compatibility evidence. No source/client workbook was opened.

At the initial unlicensed checkpoint, all four runs reported `Licensed=False`. The evaluation sheet was observed after serialization;
repeated saves of the same synthetic workbook introduced further evaluation-sheet state.
Do not benchmark preservation on this basis. At that checkpoint the harness accepted no real workbook arguments.
It can load an explicitly supplied licence through `SUMMIT_ASPOSE_LICENSE_PATH`, without
copying, embedding or logging its contents. No licence registration or purchase was made.

### Licensed checkpoint

The user subsequently supplied a local Aspose.Cells for .NET licence. It was read directly
from its supplied location outside the repository, using the process-local environment
variable and the supported `License.SetLicense(Stream)` API. No licence contents were
displayed, copied into the project or persisted in application settings.

All four existing test executables (Debug/Release, actual x86/x64) report `Licensed=True`
and pass the same 13 assertions each. Both XLSB and XLSM reopened with exactly the two
synthetic worksheets; the additional evaluation sheets no longer appear. The original
environment variable was restored afterwards. No real workbook was opened, no performance
measurement was taken, and Summit itself was not rebuilt or changed.

The next checkpoint extends this harness for disposable-copy I/O/preservation testing. PMCost/RespCost
adaptation, real-model calculation, structural insertions and end-to-end timings remain pending.

## First real-model preservation checkpoint

Only the separate `Tools/SpreadsheetEngineTrial` project and this trial report were changed.
Summit's production project, references, executable, version and workbook originals were not
changed by this work. Generated copies/results are under ignored `obj/AsposeTrial` directories;
no licence contents or extracted VBA source are stored there or in tracked documentation.

The harness independently reads both input and output with DevExpress 25.2.4 in manual
calculation mode. It never calls calculation or executes VBA. After an Aspose no-op XLSB
load/save of a copy of the repository Demo master:

- All **285 sheets**, their order, visibility and worksheet-protection state match.
- All **1,761 scoped names** and their definitions/hidden flags match.
- All **678,983 formulas**, including the reader's array/dynamic-array flags, match.
- All **767,141 populated cells** retain their input types/values or cached formula results.
- The entire `vbaProject.bin` and `xl/metadata.bin` are byte-identical.
- Three existing custom XML payloads and their properties are byte-identical. Their three
  relationship parts differ only in XML whitespace; Id, Type and Target are unchanged.
- The original repository Demo hash remains
  `1ED79726D8D1129C699242C9E8B3E988AC920EA7108D3C0B105BF9B0BC8A5E9C`.

**Formatting is not a clean pass.** The combined formatting fingerprint differs on 119
sheets. The detailed populated-cell check reports 81 pattern differences, 77 background
colour differences, 55,330 pattern-colour differences, 12 bold changes and 373 font-colour
changes. No differences were found in the tested populated cells' Locked/Hidden flags,
number formats, font name/size/italic, alignment or wrapping. Empty-versus-black unused
pattern colours may be representational, but this has not been established exhaustively.

Five samples were checked independently in Microsoft Excel 16.0 build 20326. Their base
patterns/colours/fonts/locks and displayed pattern/conditional-rule counts agree between
original and Aspose output. For example, Excel displays the same pattern at Development
BP Assumptions Q66 and Q98, while DevExpress exposes DarkGray before and None after the
Aspose round-trip. This is **an unresolved cross-engine interpretation/preservation issue**,
not proof that Aspose erased Excel's formatting. It matters because Summit reads `Cell.Fill`
for its existing editability and presentation rules. Do not alter those rules to accommodate
the trial without discussing the evidence with Jon.

The follow-up fill-gate scan found **76 populated cells crossing Solid/non-Solid**:
72 Solid to None, four DarkGray to Solid, and a further five DarkGray to None (the latter
five do not cross the gate). This is a concrete compatibility risk for mapped cells using
Summit's fill-based rules, not proof that all 76 cells are mapped/editable in a particular
interface. No production editability rule was changed. Evidence: `style-gate-differences.json`
in the same preservation directory. Blank-cell gates remain unaudited.

Both private Demo files open successfully through Excel's normal read-only object-model
path, with the same sheet/name/VBA/custom-XML counts. Macros/events/link updates are disabled,
calculation is manual, and neither file is saved. Microsoft's documented default for
[Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open)
is normal load without object-model recovery. This does not prove VBA execution, complete
visual parity, a subsequent Excel save/reopen, or financial correctness.

Preservation evidence: `obj/AsposeTrial/b7a636c32f854e8db562f8b7f1370d4d`;
`demo-source-manifest.json`, `aspose-demo-manifest.json`, `style-differences.json`,
`excel-styles.json`. The earlier `excel-open.json` includes an unavailable/null RepairMode
probe; it is superseded by the documented normal-load test in `excel-styles.json`.

## Performance method and cautions

`Test-Io.ps1` runs three fresh Release/x64 processes per engine/output-format/sample,
sequentially and with alternating engine order. Inputs are identical frozen XLSB copies.
No OS cache flush is performed: these are **warm-cache** timings. XLSM rows measure saving
an XLSB input as XLSM, not loading an XLSM source. Native load/save times exclude licensing
startup, calculation, Summit safeguards, staging, reconciliation, UI reload and rebind.

The raw DevExpress comparator deliberately does not run Summit's CONCATENATE or recovery
metadata guards. Do not compare a known unguarded serialization defect with production as
if Summit lacked those fixes. None of the trial outputs is an approved user workbook.

The desktop was not isolated: a read-only process sample during AGL timing found substantial
After Effects CPU use (21.83 CPU-seconds across cores during a short sample containing a
one-second wait plus process-enumeration overhead, versus 0.64 for the trial process).
No unrelated application was stopped. Treat
AGL timing ranges as **provisional under contention**, not a controlled performance claim.
Repeat during an agreed quiet period before making a purchasing/integration decision.

### Recorded native I/O results

Each entry is **median seconds [minimum–maximum]** from three fresh processes. These are
observations from this desktop, not a controlled quiet-machine comparison; AGL was notably
affected by contention. Peak MiB is the median of each process's peak working set, not a
measurement of managed allocations alone.

| Input / output | Engine | Load seconds | Save seconds | Peak MiB |
| --- | --- | --- | --- | ---: |
| Demo / XLSB | Aspose 26.9 | 1.428 [1.406–1.439] | 2.481 [2.479–2.484] | 301.5 |
| Demo / XLSB | DevExpress 25.2.4 | 7.660 [7.573–7.842] | 10.304 [10.187–10.311] | 619.8 |
| Demo / XLSM | Aspose 26.9 | 1.440 [1.395–1.491] | 1.858 [1.836–1.888] | 314.2 |
| Demo / XLSM | DevExpress 25.2.4 | 7.643 [7.634–7.711] | 6.624 [6.559–6.672] | 604.6 |
| AGL / XLSB | Aspose 26.9 | 2.068 [1.940–6.629] | 4.255 [4.216–4.660] | 423.1 |
| AGL / XLSB | DevExpress 25.2.4 | 19.289 [13.116–37.121] | 35.648 [17.232–41.237] | 822.5 |
| AGL / XLSM | Aspose 26.9 | 1.933 [1.920–6.277] | 3.195 [3.163–8.299] | 432.3 |
| AGL / XLSM | DevExpress 25.2.4 | 29.578 [13.756–34.497] | 20.329 [11.566–24.727] | 803.8 |

All 24 runs completed, reported 64-bit execution, and retained their source SHA256. The
AGL original/private-input hash is
`30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47`.
Evidence: `obj/AsposeTrial/81d253a932204533a8fd31b00b2f4fbd/runs.json` and individual result
files. **No purchasing or shipping decision should use these raw medians as an expected
Summit user-facing speed-up.**

Excel also opened the private AGL original and Aspose XLSB/XLSM outputs normally, read-only,
with 283 worksheets, 1,762 names, VBA present and six COM-reported custom XML parts in all
three. Demo Aspose XLSM opened with its matching 285/1,761 counts. The Excel COM custom XML
count includes standard parts and is distinct from the three package custom XML payloads
inspected above. No VBA execution, calculation or saving was performed in these checks.

### Quiet-period repeat - 23 September 2026

After Jon confirmed the machine was quiet, repeated the same 24 tests (three fresh Release/x64
processes per engine, workbook and output format), with unchanged I/O settings and identical
Demo/AGL source hashes. This run took place at approximately 22:19-22:23 UTC. No source files,
production code, references, calculation policy or editability rules were changed.

A brief pre-run process sample showed approximately 4.0% aggregate readable-process CPU
across 24 logical processors. The post-run sample showed approximately 0.3% other-process
CPU and no AfterFX entry. These are spot checks, not continuous hardware monitoring; this
remains a warm-cache desktop test, not a controlled lab or minimum-spec client result.

Use these repeat measurements in preference to the earlier contention-affected table.
Values are **median seconds [minimum-maximum]**; peak MiB is the median process peak.

| Input / output | Engine | Load seconds | Save seconds | Peak MiB |
| --- | --- | --- | --- | ---: |
| Demo / XLSB | Aspose 26.9 | 1.436 [1.421-1.507] | 2.165 [2.145-2.181] | 302.0 |
| Demo / XLSB | DevExpress 25.2.4 | 6.166 [6.154-6.224] | 6.964 [6.833-6.998] | 620.6 |
| Demo / XLSM | Aspose 26.9 | 1.427 [1.414-1.432] | 1.708 [1.688-1.728] | 310.9 |
| Demo / XLSM | DevExpress 25.2.4 | 6.268 [6.086-6.277] | 4.812 [4.716-4.829] | 605.2 |
| AGL / XLSB | Aspose 26.9 | 1.777 [1.775-1.780] | 3.125 [3.106-3.127] | 422.6 |
| AGL / XLSB | DevExpress 25.2.4 | 9.637 [9.574-9.677] | 11.126 [11.112-11.132] | 824.3 |
| AGL / XLSM | Aspose 26.9 | 1.763 [1.751-1.769] | 2.431 [2.413-2.442] | 434.0 |
| AGL / XLSM | DevExpress 25.2.4 | 9.708 [9.675-9.779] | 7.735 [7.706-7.759] | 807.0 |

For native XLSB saving, DevExpress/Aspose median-time ratios are about **3.2x for Demo** and
**3.6x for AGL**, not the roughly 8.4x suggested by the earlier contended AGL medians. Both
engines benefited from the quiet period, particularly DevExpress. XLSM serialization was
also faster with Aspose, but all source loads in this test were still XLSB.

All 24 runs completed successfully at 64-bit and reported unchanged source hashes. The
runner also rechecked the originals after each case. Raw records and outputs are retained
in `obj/AsposeTrial/96ba7de65f524f198c72ed96ec3ecc15`, including `runs.json`.
This repeat measures speed only: the earlier preservation findings and unresolved gates
are **not cleared** by these timings, and these are not Summit end-to-end measurements.

## Installed Microsoft Excel remains a separate candidate

Jon reiterated that many clients have their own licensed desktop Excel installation.
Do not frame the choice as only DevExpress versus purchasing Aspose. Compare three routes:

1. Existing DevExpress workbook/UI path, with Summit's production safeguards.
2. An isolated Aspose helper, only after its preservation/interoperability gates pass.
3. An optional installed-Excel helper using the **verified, model-compatible VBA routines**
   for Funding/Development/other supported structural operations on disposable working copies.

The Excel path could reduce duplication of business-model VBA while retaining Summit's
interface, but it is not implemented or performance-validated by these I/O tests. Compare
the complete **Summit staging save -> Excel startup/open -> verified operation -> Excel save
-> Summit reload/rebind** time, correctness and failure recovery. Detect compatible desktop
Excel and respect macro/trust/blocked-file policy; do not silently weaken security or disturb
the user's existing Excel instances. Keep a supported fallback where Excel is absent or
unavailable. A licensed installation alone does not establish that the automation workflow
or a particular workbook's macros are compatible.

The Excel opens recorded elsewhere in this report had macros disabled and measured neither
VBA structural performance nor this full hand-off. No Excel performance result is implied.

## Populated AGL and XLSM preservation results

The independent AGL XLSB comparison passes the tested core fields: **283 sheets, 1,762
names, 1,077,821 formulas, 64,918 reader-reported array ranges and 1,182,776 populated cells**.
Formula strings, complete array definitions/ranges, constants, cached results, sheet order,
visibility/protection and name definitions match. VBA and binary metadata are byte-identical;
the three custom XML payloads/properties match. This first AGL pass deliberately omits style
comparisons; it does not override the Demo fill-gate finding.

**XLSM is not an exact DevExpress representation match.** Both Demo and AGL retain sheet/name
counts, inputs, cached results and VBA bytes, but formula strings differ on seven sheets.
AGL has 8,545 differing strings; the bounded examples show whitespace around operators or
after separators being removed, with quoted text retained. Do not count these examples as
proven arithmetic corruption, or claim every differing expression has been semantically
validated: a formula parser or calculation comparison is still needed.

DevExpress classifies AGL's 64,918 source array ranges as 98 legacy plus 64,820 dynamic in
the Aspose XLSM, and Demo's 50,298 as 98 legacy plus 50,200 dynamic. Binary metadata becomes
`xl/metadata.xml`, as expected for an XML format. Three corresponding AGL cells were read
in Excel on both sides: Management Costs Assumptions D61, Dvpt BP Rev and Exp Assumptions D51,
and Leaseholder Units C13. In all three, Formula/Formula2 agree and Excel reports HasArray=False
and HasSpill=False on both sides. One sampled range per affected sheet also retains the same
formula in the DevExpress array collections. Thus the evidence points to an import/metadata
interpretation issue requiring investigation, **not a demonstrated conversion of Excel's
legacy arrays into spills**. It is not a full-array or recalculation acceptance test.

Evidence in the benchmark output directory: `*-core-v2.json`,
`AGL-xlsm-formula-differences.json`, `excel-open.json`, `excel-array-samples.json`.
Schema 2 fingerprints each array range once. The initial AGL schema-1 scan was intentionally
stopped because per-cell array lookups were expensive; no source was altered and no completed
manifest from that interrupted scan is used as evidence.

## Harness validation and next gates

- Final isolated Debug/Release builds for x86 and x64: zero warnings/errors; all four
  licensed runs pass 13 synthetic assertions each.
- Original-path access, existing-output overwrite and unlicensed real-model save requests
  are rejected. The unlicensed test creates no workbook output.
- All timing sources and both original files retain their starting SHA256 hashes.
- No Aspose reference was added to Summit's production project or packages.config.
- No financial recalculation, UDF port, structural command, client-file repair, live UI
  trial, commit or push was performed in this checkpoint.

Next, investigate the Solid-fill and array-metadata interpretation differences using the
existing Summit reader and small synthetic fixtures. Do not weaken its locking rules or
patch user workbooks. Then finish complete styles/validation/conditional-format/drawing
preservation, test Blank/Stori, and measure genuine minimal-change staging/reload paths,
including the installed-Excel option. The quiet-period native I/O repeat above is complete.
Only after the relevant preservation gates should custom calculations and
structural operations move beyond experimental copies.

## Current outcome

Quiet-period timings confirm a substantial native I/O advantage for Aspose on these two
workbooks, but they do not select the production approach or rule out an Excel helper.
The first Demo formula/name/VBA checks are promising, but formatting interpretation needs
resolution and the broader preservation gates remain open. Keep DevExpress and the current
client release unchanged. Blank/Stori preservation, full blank-cell/conditional-format/
validation/drawing checks, financial callbacks, structural operations and a genuine Summit
end-to-end comparison remain separate work; raw serialization speed does not settle them.
