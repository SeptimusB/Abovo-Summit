# Excel offload and workbook-format investigation — 22 September 2026

## Decision

Keep XLSB as the normal business-plan format. Offer opt-in XLSM recovery copies separately. Desktop Excel is a feasible **optional future structural-operation provider**, but no production Excel offload, VBA task runner, macro-trust change, or Interop package installation is included in this delivery.

The user prioritises the opening experience despite saves being more frequent. Excel must be installed and compatible; macro trust, blocked downloads, Protected View and deployment policy remain separate gates. The user controls the VBA and can add a versioned automation entry point.

## Controlled format measurements

Source: a private frozen copy of `AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb`, SHA-256 `1E3C483E3D3B8B7D250C9B3B7DD742CFEB2DAE3AA5F9AD8B578D41C317BAF9AD`, 11,717,641 bytes. The supplied original was subsequently replaced/changed externally during this task; these measurements concern the frozen copy, not its later contents. Tools only read/copied the supplied original.

Three alternating runs per format, against equivalent Excel-resaved baselines. Desktop Excel: 64-bit 16.0.20326.20144; DevExpress: 25.2.4, x86 document API. Macro execution, events, link updates and calculation were disabled/manual. These are warm-machine I/O measurements, not cold-cache, full Summit UI, or financial/VBA benchmarks. Excel uses SaveCopyAs; DevExpress uses SaveDocument. Identical formula-compatibility preparation was applied before each DevExpress save and measured separately (~6 seconds).

| Median seconds | XLSB open | XLSM open | XLSB save | XLSM save |
| --- | ---: | ---: | ---: | ---: |
| Desktop Excel | 4.709 | 15.106 | 3.412 | 6.395 |
| DevExpress x86 | 10.754 | 29.813 | 13.285 | 7.940 |

DevExpress XLSM saves about 5.35 seconds faster here, but opening costs about 19.06 seconds more: roughly four saves balance the extra open time. This arithmetic does not outweigh the stated opening-experience priority. Excel itself is faster with XLSB for **both** operations on this file. Recovery is different: it is saved repeatedly and opened rarely.

Native output reopen probes: XLSB 10.659 seconds; XLSM 28.710 seconds. Excel baselines: 17,644,316 / 21,552,704 bytes; DevExpress outputs: 14,045,815 / 17,971,504 bytes respectively. Resaving changes package size; this is not evidence that original file contents were identical byte-for-byte.

Evidence: `obj/WorkbookFormatBenchmarks/a6663ba8cac54e77953ba20d91d64218/` (`excel.json`, `devexpress.json`, verification logs). Reproduce with `Benchmarks/Compare-WorkbookFormats.ps1` and `Tools/WorkbookFormatBenchmark.cs`.

### Preservation checks and limits

- Exact native digest equality after preparation/save/reopen, per format: formula text, constant value/type, workbook and worksheet-scoped names, worksheet order/names and protection.
- All 335 VBA module identities and source hashes identical across the frozen original, both Excel baselines, six Excel saves and six DevExpress saves. Aggregate `1539cc5c003414274742dda616da675ebf68b420e4143bd1ca5a7e17d68fa470`.
- Three custom XML payloads and custom document-property values/types match semantically across those 15 files. XML prefixes/part filenames and binary VBA container bytes can change without changing those payloads.
- These checks do **not** establish every style/validation/array-group detail, macro execution or financial equivalence after Excel calculation.
- The later recovery trial found a distinct **direct XLSB → XLSM metadata omission** not exposed by the Excel-resaved benchmark baselines. See the recovery trial document. Do not generalise the baseline benchmark into a guarantee that arbitrary direct conversions round-trip safely.

## Excel structural-core probe

On the frozen AGL copy, shifting/filling the 14 `TransCopy_DevptSingle_A:N` mirrors for ten rows took 7.307 seconds in desktop Excel: 2.422 seconds shifting, 0.548 seconds copying, with the rest in COM/name/setup work. This isolated probe deliberately did not expand source Development ranges and was **never saved**. It is not a complete valid insertion or an end-to-end offload benchmark.

Funding/source Development insertion probes stopped at worksheet protection. No passwords were extracted, no protection bypassed and no macros executed. The existing approved VBA protection logic belongs inside a future controlled entry point. Historical Summit and client-reported Excel timings concern different runs and must not be presented as a controlled speed-up factor.

Tool: `Tools/Measure-ExcelStructuralCore.ps1`; output log `obj/excel-structural-core-agl.log`. All probe changes were discarded. Existing user Excel instances were not used or closed.

## Proposed optional Excel provider

1. Preflight installed desktop Excel, a supported version, workbook contract, allowed structural operation, model identity, macro-trust state and VBA 32/64-bit compatibility. Do not bypass Trust Center, blocked files or Protected View.
2. Save a private current checkpoint; retain the live Summit model and original until a complete result is validated. A dedicated STA worker process owns its Excel instance and cleans up only that instance.
3. Prefer an explicitly invoked, fixed, versioned VBA entry point, such as `SummitAutomation_Run`, over an unqualified Auto_Open task flag. Supply a bounded operation code, count, contract version and operation ID. Never run a macro name read from arbitrary workbook content.
4. Suppress ordinary open-event work only through the controlled VBA contract/events policy; let the entry point perform all source/master expansions and linked TDB updates, followed by one intended calculation and integrity check.
5. Save the private result, close owned Excel, load/validate in Summit, then swap workbook/services/bindings as one accepted result. Preserve UI position where possible. Existing Undo commands cannot refer to the old workbook: explicitly establish a history checkpoint and retain audit text, not executable stale commands.
6. On failure/cancellation keep the original/live workbook. Native fallback must start from the untouched checkpoint, never a partially modified Excel result.

A rough lower bound for handoff I/O alone is ~34 seconds (native save + Excel startup/open/save + native reopen), **not a measured full workflow**. Macro calculation, validation, UI/service reconstruction and compatibility preparation add to it. Offload should therefore be limited to measured large operations; normal edits and small inserts should remain native.

No package is required for late-bound COM; a typed Interop reference is optional and does not install Excel. Out-of-process COM can cross x86/x64 boundaries; the VBA itself must still be compatible with installed Excel bitness.

## Primary references

- [DevExpress supported import/export formats](https://docs.devexpress.com/OfficeFileAPI/405569/spreadsheet-document-api/import-and-export).
- [DevExpress slow spreadsheet load support discussion](https://supportcenter.devexpress.com/ticket/details/t1305325/loading-spreadsheet-very-slow-compared-to-excel).
- [ExcelDataSource XLSB limitations](https://supportcenter.devexpress.com/ticket/details/t1191066/exceldatasource-xlsb-format-support): this lightweight data-source component is not the full workbook API.
- [Office primary interop assemblies](https://learn.microsoft.com/en-us/visualstudio/vsto/office-primary-interop-assemblies?view=visualstudio), [cross-bitness process interoperability](https://learn.microsoft.com/en-us/windows/win32/winprog64/process-interoperability), [VBA 32/64-bit compatibility](https://learn.microsoft.com/en-us/office/client-developer/shared/compatibility-between-the-32-bit-and-64-bit-versions-of-office).
- [AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity), [Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open), [Application.Run](https://learn.microsoft.com/en-us/office/vba/api/excel.application.run). ForceDisabled does not suppress Excel 4 macro prompts; the benchmark rejects packages with macro sheets. DisplayAlerts is not a macro-security bypass.

No workbook was uploaded to an external service.
