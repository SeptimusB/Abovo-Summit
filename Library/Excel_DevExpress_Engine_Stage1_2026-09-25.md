# Excel-first / DevExpress fallback: stage 1 checkpoint

## Authority and scope

The user approved staged implementation after reviewing SpreadsheetGear and Spread.NET. The finished product should prefer **compatible installed desktop Excel automatically**, keep an explicitly selectable DevExpress route, and retain the no-Excel product contract. Use required workbook VBA functions only when existing Excel policy permits. Never change Trust Center settings or force-enable macros. This supersedes the deferred 23 September proposal and the Gear-primary research direction.

**This checkpoint is a tested read-only calculation foundation, not a production engine switch.** The normal 2.98 Debug/Release executables, live edit/save/calculation paths, workbook-owned fills/locking, masters and customer originals are unchanged. There is no newly released executable or client acceptance implied by this checkpoint.

Publishing permission is code/tests/technical notes only. Local complete-history checkpoint: `f6b0ec8`. Published sanitized checkpoint: `d01d11597a0b1f8516609f22d9458588fbf6c099`, branch `codex/checkpoint-2.98-code-only`, parent `38616ccbf6f86e7b23eeabd4f010fc04fb880090`. Do not push local `main` as-is: two older unpushed ancestors contain Word attachments. Newer Word reviews and the synthetic support package also remain local. No licences, settings, build outputs, raw VBA, passwords or disposable customer workbooks are published.

## Implemented boundary

- `Services/WorkbookEngines/WorkbookEngineContracts.vb`: preference/options, detached typed cells/errors, immutable bounded rectangles and revision-stamped results. Settings default to Automatic, but are **not yet connected to Options or model opening**.
- `WorkbookEngineStaHost.vb`: one background STA owner and message loop; serialized queue prevents native message pumping from re-entering another workbook operation. Native objects never cross to the UI. In-flight cancellation discards the result after the native call completes; it does not abort COM.
- `WorkbookCalculationSession.vb`: readonly source lease and SHA256, one selected backend, revision/session/hash checks, bounded batch reads, exact requested-area verification, stale/late-result refusal, idempotent close. Excel may fall back only during initial opening after cleanup. Native calculation/read failure faults the session; it never silently switches engine mid-session.
- `ExcelCalculationBackend.vb`: owned separate Excel process, no PIA/NuGet runtime requirement, readonly original open, updates/events disabled, manual calculation, actual dynamic-array feature probe and known-answer workbook PMCost/RespCost probes in an unsaved scratch workbook. ByUI security is used only around the authorized open, then ForceDisable restored. Download-zone and external-link/connection/XLM checks run before opening. Existing user Excel processes are not attached to or closed.
- `DevExpressCalculationBackend.vb`: private native Workbook, manual/recursive calculation and workbook-local native custom functions. No mutation of the application's global function registry.

Excel values are read as a rectangle in one COM transfer, normalized into numbers/text/booleans/blanks/typed errors. Modern/unknown Excel error HRESULTs cannot become financial numbers. Results are **not** written into the UI workbook or over formulas. Formatting, permission refresh and displayed-value projection are the next stage.

## Reproducible validation

Harness: `Tools/WorkbookEngineFoundationTests` (.NET Framework 4.8, x86 by default). It references a separate candidate Summit build. Existing evaluation licences are not needed for Excel or the production DevExpress installation; no Gear/Aspose/Spread dependency is introduced.

Both candidate Debug and Release builds pass, to `bin/EngineStage1-Debug` and `bin/EngineStage1-Release`. Normal outputs are not overwritten. Logs below are ignored local evidence, not attachments for publication.

- Deterministic safety: **73 assertions** against each candidate build. Includes STA ownership, non-reentrant message pumping, cancellation before/during operations, bounded transfers, typed errors, stale and cross-session results, opening fallback, cleanup, wrong-area rejection and failed-session refusal.
- Security: **12 assertions**. Private disposable synthetic copies with Internet/Restricted NTFS zone markers are rejected, markers remain unchanged, Intranet marker can be read without rewriting it, and external-connection/XLM package entries are refused. A discovered .NET Framework alternate-stream path limitation was corrected with a readonly native handle. The fixture deletes only its own GUID-named files.
- Native synthetic, x86 Debug: **40,963 cells agree**, owned Excel exits, pre-existing Excel processes remain, VBA-disabled/required-function policy selects DevExpress, source bytes unchanged.
- Native AGL original-format XLSB, x86 Debug: **45,586 positions agree** across Detailed Comp Inc - Trad View, Financial Position - Trad View, Cashflow detailed, Check Sheet and Development Expenditure. This includes blanks; it is not whole-workbook certification. Required real VBA probes pass. Owned Excel exits; pre-existing Excel processes remain; required-VBA-disabled opening falls back; source bytes unchanged.
- Native converted AGL XLSM, x86 Release: **45,586 positions agree** over the same five rectangles. Required VBA probes, source-byte preservation, owned-process cleanup and VBA-disabled fallback also pass. This is a separate run, not an inference from XLSB.

Local evidence: `obj/EngineStage1-safety.log`, `obj/EngineStage1-safety-release.log`, `obj/EngineStage1-security.log`, `obj/EngineStage1-synthetic-x86.log`, `obj/EngineStage1-AGL-xlsb-x86.log`, `obj/EngineStage1-AGL-xlsm-x86.log`.

### Timings: single AGL XLSB run, not an end-user guarantee

| Operation | DevExpress | Excel |
|---|---:|---:|
| Readonly open and compatibility setup | 7,909 ms | 7,063 ms |
| First full rebuild | 6,603 ms | 4,127 ms |
| First transfer, 45,586 positions | 25 ms | 108 ms |
| Warm full calculation | 6,394 ms | 434 ms |
| Warm result transfer | 16 ms | 31 ms |

Excel shutdown took another 3.17 seconds after close; this is not included in calculation figures. An earlier synthetic cleanup assertion used only 500 ms and failed prematurely; the final test waits up to ten seconds and passes. Initial limited x64 XLSM probing was exploratory and is not a substitute for the final x86 coverage. Do not compare cold/warm runs or formats as though they were identical workloads. Client minimum-spec performance, bitness combinations, activation/security failures and live edit-to-display latency remain to measure.

The separate x86 Release XLSM run measured DevExpress/Excel open 12,979/15,591 ms, first rebuild 25,168/4,110 ms, warm full calculation 6,905/414 ms and warm transfer 16/26 ms. The converted dynamic-array representation has a materially different cold setup cost; these measurements do not imply a production XLSB regression. Final hashes confirm normal Debug/Release 2.98, Blank/Demo masters and original AGL unchanged.

## Remaining stages and release gates

1. **Safety and display projection:** add bounded operation supervision and explicit hung-process recovery; validate deployed Office security/activation and startup add-ins. Prove a read-only DevExpress display can consume Excel values without independent recalculation or overwriting formulas, including conditional fills and dynamic arrays. Keep one authoritative workbook owner.
2. **Interactive integration:** connect per-model engine preference/status, model opening, typed `ModelChangeManager` writes, expected-before values, undo/redo, revision/dirty state and rollback. Batch reads/formatting; preserve fill-based editability. Validate Funding date entry/unlocking and Check Sheet/header freshness. No enabled preference UI before the selected engine genuinely owns this path.
3. **Structural and persistence integration:** move actual Funding/Development/other structural commands, names, snapshots, custom XML/history, save/save-as and recovery behind the same owner. Preserve original protection, transactional restoration, exact save revision, XLSB/VBA metadata, verified private candidate and atomic replacement. Reuse proven structural rules; do not equate a row insert with a complete business command.
4. **Failure and client acceptance:** missing/incompatible/unlicensed Excel, policy-blocked VBA, prompts/hangs/crashes, multi-model cleanup, no-Excel route, discard/close/recovery, dirty N+1 during save N, x86/x64 and minimum-spec tests; full Summit → Excel/VBA → Summit round trips plus Jon/Alex/accountant acceptance.

Known stage-1 limitations: native calls currently have bounded busy retries but **no overall timeout or crash-isolated supervisor**; a blocked Excel security/password/add-in prompt can leave an operation waiting. The read-only baseline lease is deliberate for tests and must become a security-preserving private working-copy policy for the live session. No edit, structural or save API is exposed. No production integration until these gates pass; no claim that staging results alone proves Excel is safe for arbitrary client workbooks.

## References

- [Microsoft AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity): ByUI/ForceDisable policy and XLM caveat.
- [Microsoft Office threading support](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=visualstudio): STA ownership and busy calls.
- [Excel error codes](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcverror): preserve typed errors including spill errors.
- [DevExpress workbook-local custom functions](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Functions.WorkbookFunctions.CustomFunctions).
- Possible projection experiment, **not implemented**: [ICustomCalculationService](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraSpreadsheet.Services.ICustomCalculationService) and [CellCalculationArgs.Value](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Formulas.CellCalculationArgs.Value). The service requires the chain-based engine, so it is not a drop-in for Summit's recursive calculation paths.
