# Funding and Development insertion tracing - test 2.56

Follow-up: test 2.57 publishes both Debug and Release normally, adds explicit unprotect/protect timings, and changes only the shared worksheet re-protection verifier policy. See `Library/Worksheet_Protection_Performance_2026-09-22.md`. The 2.56 publication note below is historical, not the current Debug delivery status.

## Scope

The Funding header-button route now uses the existing owner-bound progress notice, starting after the user accepts the record count. It remains open through the workbook edit and DIT section rebuild, then briefly displays Complete. Successful inserts also get a System Messages entry with the total elapsed seconds when the transaction has not already published success. Cancellation creates no notice; errors dismiss it before the error dialog. The temporary callback is restored in Finally so rebuilt controls cannot retain the DIT. The existing footer Add Lines route retains its own progress notice, without nesting another one.

Detailed passive timings are enabled for FUNDING_RECORDS, DEVELOPMENT_IDENTIFIED_RECORDS and DEVELOPMENT_MULTIYEAR_RECORDS. No structural sequence, formula, named-range rule, protection policy, calculation policy or XLSB schema was changed for this instrumentation.

## Capturing a run

Run the Debug executable under Visual Studio and copy its Output log from before the insert until the interface is usable again. Keep the complete trace, including:

- `[Structure Insert Benchmark]`: operation ID, model, rule, requested count, UTC start, version, process bitness, insertion column, engine and calculation mode. Stage-start records make a long-running stage identifiable before it returns. Each completed stage has milliseconds; the final record aggregates stages and reports total and outcome.
- Stages: 3-D reference capture; begin mutation; physical column shift per worksheet/batch; template copy per worksheet/batch; 3-D copy adjustment; 3-D reference apply; named-range resize; EndUpdate; calculation-engine/mode restoration; mutation-guard release; post-actions.
- Existing `[TDB Mirror Benchmark]`, `[TDB Sync Benchmark]` and `[Population Benchmark] Structure post-actions` split post-actions further.
- `[Population Benchmark] DIT grid-insert` separates the header button's workbook/event work from section rebuilding. Footer controls continue to emit `DIT add-lines`.

Stages are sequential, not nested in the aggregate. `other` includes preflight, protection/visibility changes, bookkeeping and trace-listener overhead. These are elapsed wall-clock measurements; no cell values or formula payloads are logged and no diagnostic workbook scans are added. An interrupted scope reports interrupted rather than success. An elapsed stage alone is not proof of successful completion: use the final outcome.

Timings are written to the existing .NET Trace output, not embedded in the workbook. A saved XLSB alone cannot reveal the elapsed time of its Excel macro.

## Stori comparison protocol

1. Preserve one unchanged starting Stori XLSB and create two identical working copies, labelled Summit and Excel. Supply the exact original filename/version; do not infer parity from the filename alone.
2. In the Summit copy, add the chosen number of Funding records, then Development records. Record the counts, order, Identified versus Multi-year choice, whether the Analyser was open, and whether these were the first operations after opening. Save As a distinct Summit result. Capture the full trace.
3. Start from the identical baseline Excel copy and perform the same VBA actions, counts and order. Record the approximate time from accepting each count until Excel is usable, and save a distinct Excel result. Preserve any Check Sheet/circular-reference messages. Use the same machine and do not benchmark both applications concurrently.
4. Provide the baseline, Summit result, Excel result and timings. Compare input preservation, names and dimensions, corresponding formulas/array types, Transactional DB mirrors, Check Sheet, SOCI and financial position. A VBA-aware financial check remains separate from macro-disabled structural comparison.
5. If also testing Excel edits on the Summit-produced file, make a third copy and label it Summit-then-Excel. That tests round-tripping; it is not directly comparable with the Summit-only result because it contains additional inserted records.

No automatic optimisation is made on the basis of unmeasured expectations. Use these traces to choose the next repair/performance target.

## Validation

Debug and Release builds pass. The isolated trace fixture checks success, failure, interruption, repeated disposal and diagnostic-listener failure isolation. Funding +8 passes through the actual event-service route: its progress callback occurs before mutation, all 32 physical shifts are traced, all eleven mirrors resize, protection/calculation settings are restored, and add/delete restores the original names/formulas. Identified Development +3 passes in Release with both batches and all seven sheets traced, fourteen mirrors checked, all 1,679 global names and 106,041 linked-sheet formulas restored after deletion. Multi-year Development +3 passes in Debug with its single batch, seven sheets, fourteen mirrors, protection/calculation restoration, deletion safeguards and the same name/formula restoration checks. All three fixtures verify the original master hash is unchanged.

These are correctness/instrumentation checks on disposable Blank-master copies, not Stori performance benchmarks. The Funding and Identified fixtures ran concurrently; their elapsed times must not be used as a comparative baseline. The populated Stori runs, visible progress acceptance and equivalent interactive Excel/VBA operations are for the user's next test. No source master or Stori file has been changed. No commit or push was requested.

Reproduction: `Tools/Test-StructuralInsertBenchmark.ps1`, `Tools/Test-FundingStructure.ps1`, and `Tools/Test-ColumnFamily.ps1` with `-Rule DEVELOPMENT_IDENTIFIED_RECORDS` or `DEVELOPMENT_MULTIYEAR_RECORDS` and `-Count 3`. Local logs: `obj/structural256-*.log`; detailed operation logs are `insert-trace.log` in each fixture's reported output directory.

Final delivery status: Release is published to `bin/Release/Abovo-summit.exe`. Final Debug compiles successfully to `obj/Structural256Debug/Abovo-summit.exe`, but publishing to `bin/Debug` is awaiting closure of the user-started Summit Debug session. The earlier 2.56 Debug build already contained the tracing/progress changes, but not the last System Messages success-entry addition. Do not describe the final Debug binary as refreshed until the normal Debug build succeeds after closure. Source Blank/Demo hashes match the 2.55 audit and the normal Git whitespace check passes.
