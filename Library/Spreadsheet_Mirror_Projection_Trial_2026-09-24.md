# Gear projection with DevExpress / Excel workers — second checkpoint

24 September 2026. **Isolated research only; not a Summit release or production-engine approval.** Production Debug/Release remain 2.98 and byte-identical. Original AGL, Blank and Demo XLSB files are unchanged. No client file was replaced; existing Excel processes were not used or closed.

## Outcome

The calculation-only Gear importer now works through the complete native-worker/reload lifecycle. The optional Excel-backed AGL Funding trial passes the bounded reference checks. The DevExpress-backed trial runs to completion but **fails compatibility checks**: its serialized Transactional DB loses 3,729 array declarations, and Gear then reports three TDB Check Sheet failures. A source-aware native row-copy attempt did not resolve this. Do not promote either route into production yet; the Excel path also still needs wider function, workload, VBA/UI and failure testing.

An explicit native-result fallback has been prototyped and tested with both workers. It returns expanded dynamic-spill values at an exact revision, rejects stale/cross-session/oversized results, and never writes result values over workbook formulas. This is not automatic dependency classification or a live Summit UI implementation.

## Changed research code

- `Tools/SpreadsheetEngineTrial/GearImportProbe.cs`: calculation projection selects `ReadObjects=false`, keeps `ReadVBA=true`, and records that it is non-authoritative.
- `FundingTrial.cs`: a reduced Gear projection is prohibited from saving **before any output file is created**. Historical full-Gear diagnostic commands remain available, not approved for customer serialization.
- `MirrorTrial.cs`: all initial/reopen/rebase Gear instances use that projection; no speculative Gear Funding insert runs in parallel. Only the selected native worker performs structure changes. Added exact-revision native range reads; failed native calculation invalidates subsequent reads/saves.
- `MirrorValues.cs`: session/base/revision/rectangle/shape validation; 100,000-cell transport bound; late-read/newer-edit, wrong-session and native-calculation-failure regressions.
- `MirrorArrayCopy.cs`: diagnostic native single-row array-preserving copy and synthetic save/reload checks; rejects intersecting multi-cell spill templates before copying. The AGL attempt reports zero repairs because the copied cells still report dynamic immediately after copy; it does **not** fix the eventual loss.
- `MirrorArrayDiagnosis.cs`: read-only package versus native-reload interpretation for the missing TDB arrays. `Program.cs` exposes these isolated commands.

Gear's documented object-import switch skips charts, pictures and drawing objects, but not chart sheets. It is appropriate only for the disposable calculation projection, never the complete saved model. [SpreadsheetGear ReadObjects documentation](https://spreadsheetgear.com/support/help/spreadsheetgear.net.9.0/SpreadsheetGear2023~SpreadsheetGear.IWorkbookSet~ReadObjects.html).

The array experiment uses DevExpress's native array APIs; it does not invent or transplant `xl/metadata.xml`. [DevExpress 25.2 DynamicArrayFormulas](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Worksheet.DynamicArrayFormulas?v=25.2), [array-formula operations](https://docs.devexpress.com/OfficeFileAPI/14942/spreadsheet-document-api/formulas/array-formulas).

## AGL measurements

All use the immutable common XLSM input SHA256 `62B44DE2764C4D1D1C9C7029C9297EEF4D461E300BAB4AD7DB092BBE1D23193C`, converted earlier from the approved AGL source. These are individual high-spec development-machine runs, **not client guarantees or live Summit UI timings**. Debug synthetic tests/audits ran during parts of the second DevExpress experiment; do not interpret its small timing differences statistically.

| Measurement | DevExpress worker, first pass | DevExpress, array-copy attempt | Excel worker |
|---|---:|---:|---:|
| Native load + rebuild | 36.54 s | 34.87 s | 20.09 s |
| Gear projection load + rebuild | 4.76 s | 5.67 s | 5.03 s |
| Three review edit/journal/Gear-calc observations | 0.50–0.77 s | 0.56–0.74 s | 0.49–0.76 s |
| Worker save + full calculation + private package check | 14.67 s | 13.76 s | 6.83 s |
| Funding structural command + native rebuild | 151.53 s | 157.09 s | 17.73 s |
| Structural checkpoint save | 15.79 s | 15.65 s | 6.74 s |
| Gear rebase + rebuild | 4.75 s | 4.69 s | 4.52 s |
| Full Funding/save/rebase barrier | **172.10 s** | **177.46 s** | **29.02 s** |

The native worker save runs separately from Gear; an additional Gear edit while it saves took approximately 0.48–0.60 s. Tests confirm Save N did not save or clear the later N+1 edit. None of these timings includes a production journal adapter, actual DIT rebind, transport eligibility scan or full financial acceptance.

## Independent correctness checks

Oracle: the earlier actual Excel/VBA ten-column insertion, `obj/AsposeTrial/funding-d8bcacf09cdf4b308742ce2d4bf6645d/excel-1.json` / `.xlsm`. The current Excel worker again uses the approved embedded Funding routines with ByUI macro policy, not force-enabling VBA.

**Excel worker:** all **38,108** populated probe values across five established sheets match within the audit tolerance (absolute 1e-6 / relative 1e-10); no name differences; **1,126,327** formula-cell locations match; array anchor/range/metadata geometry matches. All nine Custom XML parts and package drawing/chart/control families are retained. The 335-module source audit finds only the already-known `VB_Base` attributes in six forms changing during native Excel serialization; no other source changes. This does not certify compiled VBA, form designer storage or interactive UI execution. Full formula-text parity was established for the earlier native Excel mirror; the current audit establishes locations/geometry/probes, not a new full formula-text claim.

**DevExpress worker, both passes:**

- Same 1,126,327 formula-cell locations and retained XML/drawing/chart/control families. All 335 VBA module source bodies match the baseline exactly.
- TDB array anchors: **56,458 expected versus 52,729 saved**; **3,729 missing**, no extra anchors. Those cells retain `cm="1"` and formula text but lack `<f t="array" ref="...">`. A native DevExpress reload also recognizes **zero** of those 3,729 as dynamic arrays; this is not merely Gear failing to interpret a fully declared array.
- Native full-formula-text comparison of the first pass finds 570 differences on Development Expenditure; sampled expressions differ in `RespCost` versus `RESPCOST` casing. No other sheets differ in formula text. This is not a claim that all formula semantics or inputs have been audited.
- Five defined-name string differences reduce a one-cell rectangle to its equivalent one-cell reference (for example `$A$8:$A$8` → `$A$8`). Do not mistake this representation change for the array failure.
- Probe differences comprise **11 numeric + 11 text/blank differences**: three TDB Check Sheet failures, their messages/totals, and propagated validation headings. The reported financial amounts otherwise match within these particular probe ranges; do not generalize that to full-model financial correctness.
- The row-copy experiment validates copied array identity immediately and reports **zero arrays needing reinstatement**. The saved/reopened result still loses the same 3,729 declarations. The exact transition remains to be isolated between subsequent operations, calculation and serialization. Do not state that CopyFrom alone is the proven cause or that the attempted guard repaired it.

## Native fallback / failure tests

Passed in final x64 Debug and Release DevExpress synthetic runs and Debug/Release Excel synthetic runs (macros disabled):

1. Gear reduced-copy Save rejected with no file created.
2. Journal duplicates applied once; out-of-order command rejected.
3. Injected failed save leaves target unchanged; locked/stale destination rejected.
4. Later edit remains dirty after an earlier revision save.
5. DevExpress controlled worker termination → immutable-baseline restart/replay; Excel graceful restart/replay. Excel crash/hang cleanup remains unproven.
6. A value-only change expands the native spill from three to five, then six rows; Gear remains fixed-size, while explicit native snapshots return the full current values.
7. Worker rejects stale/oversized native reads; consumer rejects a valid-but-late result after a newer edit and rejects another session's result.
8. In the DevExpress synthetic fixture, expanding to seven rows hits its existing merged explanatory row and causes a native `KeyNotFoundException`. The exception is retained in the evidence. The worker now refuses subsequent reads and saves and leaves the prior target unchanged. This is containment, **not a fix for that DevExpress calculation edge case**.

The small array-copy fixture passes native reopen and Gear value checks and refuses a multi-cell source before writing. Its ordinary CopyFrom preserved all ten dynamic arrays; therefore it did not reproduce the AGL loss or exercise the reinstatement branch. An initial SEQUENCE-based fixture returned native `#NAME?`; the final guard test uses a supported array constant. Do not infer comprehensive SEQUENCE support from this experiment.

Excel automation remains a separate owned, serial process and does not change Trust Center settings. A production implementation needs bounded busy-call retry/cancellation and STA ownership; Office COM objects are not free-threaded. [Microsoft Office threading guidance](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=visualstudio).

## Evidence and reproducibility

Private root: `obj/AsposeTrial/mirror-20260924-v2/` (ignored; contains client data, not vendor-shareable).

- `dx-agl-1/report.json`, `oracle-audit.json`, `formula-comparison.json`: first complete DX path and failing compatibility gate.
- `dx-agl-2/report.json`, `oracle-audit.json`, `array-state.json`: source-aware copy attempt; failure remains.
- `excel-agl-1/report.json`, `oracle-audit.json`: fresh Excel path.
- `dx-synthetic-debug-final`, `dx-synthetic-release-final`, `excel-synthetic-1` (Debug), `excel-synthetic-release-final`: result freshness and safety tests.
- `array-copy-debug-3`, `array-copy-release-1`: final supported-array fixture. Earlier numbered failed experiments are retained, not overwritten.
- Both standalone Debug and Release builds pass; no production source/binary update or test-version bump for this research-only checkpoint.

`harnessCompleted=true` means the scripted scenario ran, **not** compatibility approval. Consult the independent audit; both AGL DevExpress passes fail it. The prototype's existing private package gate checks part presence/custom-XML identity, not every array declaration; it is insufficient as a production integrity gate.

Verified unchanged SHA256:

```text
Debug 2.98  921B4891473C5E7FB0E37FFE22933C340E1AF18673C830DF71AA6B07A1ECF8C3
Release2.98 DA7CFD8056CCC72C62D9FF833B387DC7520BAFD382502BD0829AD82083358C86
Blank       E05274ACD3D4821AC38F013AC45A7B57CE335BD524938D9F21E1B640D0CC5E90
Demo        1ED79726D8D1129C699242C9E8B3E988AC920EA7108D3C0B105BF9B0BC8A5E9C
AGL source  30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47
```

## Superseding third checkpoint

The array loss is now isolated to stale native dynamic-array registrations after row insertion, reproduced without AGL or VBA. A batch native workaround passes the AGL oracle in a direct run and the complete persistent worker: all 38,108 probe values and saved array geometry match. It is expensive: the DevExpress Funding/save/rebase barrier is now about 303 seconds, versus the previous Excel observation of 29 seconds. Original XLSB and converted XLSM load with different native array kinds; do not generalize this prototype defect to production without a separate test. See [DevExpress_Array_Insert_Diagnosis_2026-09-24.md](DevExpress_Array_Insert_Diagnosis_2026-09-24.md) for current source, evidence, export gate and remaining limits. Results above remain the historical second checkpoint.

## Next engineering gate at the second checkpoint (now investigated)

Keep native authority. Instrument the array manifest **after copy, after each later structural step, after full calculation and immediately before serialization**, then compare the written declarations; use that to produce a minimal DevExpress reproduction/repair. Re-run the independent AGL oracle after any correction. Do not restore old XML/array flags blindly or switch to scalar formulas to conceal the discrepancy.

After that, extend the native result fallback to conservatively classify consumers and test actual changed financial inputs/UDFs, Development add/delete, evolving Summit XML/history, failure/recovery and Excel/VBA reopen. No production routing, automatic engine switch or customer overwrite is authorized by this prototype checkpoint. The user's production 2.98 testing can continue independently.
