# XLSB formula export and saved-calculation trial — 22 September 2026

Test release: **2.58**. User authority: “Proceed. It might be concatenate”, following the AGL insert comparison. All supplied workbooks and repository masters remain unchanged; testing uses generated copies under `obj` or in-memory loads. No VBA macros were executed. No insertion algorithm or model schema was changed in this follow-up.

## Findings and repair

1. The installed DevExpress 25.2.4 XLSB exporter replaces `CONCATENATE` and `SUM` calls with more than 30 arguments with the **formula** `=#VALUE!`. The failure occurs without insertion, and is absent in the XLSX control export. Cached display values can initially disguise the lost formula. The dashboard's `Multivariable Dashboard!B41` has 37 CONCATENATE arguments.
2. Summit's saved Funding file contains formulas matching the Excel-produced file, but stale calculation caches. Recalculating the original baseline and both Funding files in memory, with Summit's actual custom functions and the Recursive engine, makes the two Funding files agree across Detailed SOCI, Financial Position, Cashflow and Check Sheet. This establishes consistency with the supplied Excel result, **not accountant approval of the model or proof that blank inserted columns should leave every financial outcome unchanged**.
3. `WorkbookXlsbFormulaCompatibility` preflights formula cells and global/local defined names before saving. It uses the public formula syntax tree, not textual substitutions. CONCATENATE calls with 31–60 arguments become shallow nested CONCATENATE calls, retaining argument order and separators. Existing 30-or-fewer calls remain unchanged. Legacy array ranges retain their geometry; dynamic-array rewrites are refused. Unsupported built-in calls above 30 arguments, and CONCATENATE above the tested 60-argument bound, stop saving with the cell/name location rather than guessing a transformation.
4. All four FileManager model save routes share `SavePreparedWorkbook`. After compatibility preparation, the workbook fully rebuilds with the Recursive engine and deferred-sheet suppression disabled; it is saved while that calculated state is current. The previous engine, calculation mode, history-enabled setting and deferred-sheet policy are restored. Preparation runs on save, **not on each edit or each Add Lines operation**. No edit-time performance improvement is claimed.
5. Failed/cancelled saves restore normalized formulas and retain the previous dirty state. A missing `DocumentSaved` event is not treated as success. Successful saves retain the compatible formula spelling in memory and in the saved file. Preparation/saving has activity feedback and `[XLSB Save Benchmark]` tracing; native Save As dialogs are not covered by the wait form.

### Important limits

- Formulas already replaced with `=#VALUE!` cannot be reconstructed from that expression. The supplied older Summit Funding and Development results have this existing dashboard damage. The new guard does **not** invent a replacement or silently copy a formula from another model. Repeat tests from the intact AGL original, or separately authorise a source-verified repair copy.
- Synthetic 255-argument expressions exposed further exporter limits even after nesting. Several candidate nesting shapes were rejected during development; they are not shipped as purported universal repairs. The shipped automatic rewrite is limited to the verified 31–60 range.
- The guard covers the model's FileManager Save/Save As routes, not a universal patch to DevExpress or every independent exporter in the repository.
- Source model calculation, Excel/VBA macro behaviour and accountant acceptance remain separate from serialization/cache checks. There is no claim that every formula family in every client bespoke model has been validated.

## Validation

- Both Debug and Release compile from the active `C:/Repos/Abovo Summit` checkout.
- Synthetic native XLSB tests cover **every count 29–60**, nested IF, commas in literal text, blank references, booleans, numbers, errors, protected sheets, multi-cell legacy arrays, global/local defined names, cancelled save, injected write failure, unsupported SUM and extreme CONCATENATE rejection. Independent macro-disabled Excel full recalculation agrees with all 36 synthetic formula results.
- Actual model Save As/private reopen passes for the AGL original, repository Blank and Demo. The AGL and Blank final checks compare value contents (including shared-string text) rather than workbook-internal CellValue identities. Early harness equality assertions were too strict for independently loaded text values; final semantic comparisons and independent Excel checks supersede those runs.
- **32-bit** Demo and populated AGL Funding runs pass Save As and output-cache/protection/engine checks. The Funding run also passes ordinary Save and injected write-failure restoration. A final 32-bit Debug Blank fixture additionally exercises cancellation without a DocumentSaved event.
- Independent Excel reads, with macros, events, links and automatic calculation disabled, show **zero differences** (numeric tolerance `1e-7`) between the prepared Funding save and the supplied Excel Funding file across the four selected output sheets. Baseline prepared save also agrees with the intact baseline's saved values. Do not count these duplicated statement views as independent financial proofs.
- The prepared baseline was opened read-only in Excel and copied with SaveCopyAs to another disposable XLSB. Native reopen preserves sheet order, defined-name count and **12,483 selected cells' formulas and values** across Dashboard and the four output sheets.
- All **335 VBA module identities and source hashes** are identical in the AGL original, prepared Summit save and Excel round-trip copy. No VBA source or password was written to the repository; this is preservation evidence, not macro-execution testing.

Representative AGL preparation in the diagnostic process: roughly 5.6 seconds preflight plus 7.4 seconds calculation, with total save around 28.5 seconds. These are diagnostic timings, not controlled x86/x64 benchmarks. Save is now deliberately more thorough. The earlier ~148-second Development insertion remains a separate performance task; no speed-up is claimed here.

## Reproduction and evidence

- `Tools/Test-FormulaExport.ps1` / `FormulaExportFixture.cs`: unguarded exporter reproduction, XLSB versus XLSX.
- `Tools/Test-AglCalculation.ps1` / `AglCalculationProbe.cs`: read-only, in-memory rebuild comparison with Summit custom functions.
- `Tools/Test-SavePreparation.ps1` / `SavePreparationFixture.cs`: synthetic tests, or `-Workbook <path>` actual model saves; `-Architecture x86` tests a real 32-bit child. Sources are hashed before/after.
- `Tools/Test-SaveExcelRoundtrip.ps1`: isolated, macro-disabled Excel saved-output comparison and private SaveCopyAs.
- `Tools/Inspect-SaveRoundtrip.ps1`: native reopen of that Excel copy.
- `Tools/Verify-VbaModuleHashes.py`: read-only VBA module hashing, with temporary external oletools runtime.

Evidence logs: `obj/save-synthetic-final.log`, `obj/agl-calculation-probe.log`, `obj/save-agl-base-final.log`, `obj/save-funding-x86-final.log`, `obj/save-demo-x86.log`, `obj/save-blank-final.log`, `obj/save-excel-base.log`, `obj/save-excel-funding.log`, `obj/save-native-excel-roundtrip.log`, `obj/save-vba-agl.log`, `obj/save-repair-build-debug.log`, `obj/save-repair-build-release.log`.

AGL calculation evidence: `obj/AglCalculation/9c056a8a977544cf99ac8cd278483dca/`. Baseline prepared save: `obj/SavePreparationTests/459af107ced7455ab1c7cbd6c0b6cbdd/`. Excel baseline round trip: `obj/SaveExcelRoundtrip/4ccb48aa8057448b8e3aa1bb9eac68fd/`. Funding independent Excel comparison: `obj/SaveExcelRoundtrip/7d94628606c5490aa9ed2bb64b8a6272/`.

## Next acceptance test

Start with the intact original AGL file, add the ten Funding columns, Save As a new XLSB, then reopen that copy in Excel and check the statements, Check Sheet and populated dashboard description. Repeat the Development action separately. Obtain accountant/Excel-VBA acceptance before treating this as a final release. Insertion-speed optimisation and a clean, unpaused Funding benchmark remain pending.
