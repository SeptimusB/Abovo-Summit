# Client test delivery 2.82

23 September 2026. User requested a temporary removal of diagnostic tracing and benchmarks, a commit of the agreed changes, and a numbered orange-annotated client issue report. No push or installer publication was requested.

## Included checkpoints

- [2.80 structural audit repairs](Structural_Audit_Repairs_2026-09-23.md): OFA/Repairs insertion, deletion boundaries, Include range handling and service-charge mirror-name resolution. Client acceptance remains pending.
- [2.81 input repairs](Client_Input_Repairs_2026-09-23.md): monetary precision, numeric validation, source display formats, Covenant mapping and empty survey dates. The report records the narrower AGL coverage and older Other Fees mapping limitation.
- 2.82 quiet client build, with no change to workbook calculation policy, structure, transactions, protection or recovery safety checks in this quiet-build step.

## Temporary diagnostics policy

The project property `SummitDiagnosticsEnabled` defaults to `false` for every configuration and architecture. `SummitDiagnostics.WriteLine` is conditional on the corresponding compiler constant, so quiet builds omit diagnostic calls and argument evaluation. Application Debug/Trace/console diagnostic outputs use that gate. Passive benchmark clocks do not start or query a Stopwatch when disabled. Detailed structural insertion timing scopes are not created. The `--benchmark-bitness` entry point displays a disabled message and closes before loading or benchmarking a workbook.

Normal progress/splash messages, errors and warnings, system messages, change history, recovery notices and integrity reports remain. Real idle clocks, time-sliced integrity scan budgets, save scheduling and other functional timers remain unchanged. Visual Studio module/thread notifications and third-party native WebView2 messages are not Summit tracing and may still appear.

Re-enable temporarily by building with `/p:SummitDiagnosticsEnabled=true`; use a rebuild when changing this setting. No global DEBUG switch was changed, so Debug-only auto-open and developer controls retain their existing behaviour. Standalone benchmark/test scripts remain in the repository for future use.

## Validation

- Debug and Release AnyCPU builds pass at version 2.82.
- `Tools/Test-QuietDiagnostics.ps1` passes against both executable folders. Compiled-source inventory has no ungated application diagnostic outputs; helper compilation with the flag off omits even argument evaluation and active timing, while flag-on restores both. Both delivery executables contain `Enabled=False`. The three functional integrity Stopwatches remain present.
- Existing native client-report regression passes with diagnostics disabled: 48 assertions covering independent SHG fields, management-period caption, stable column widths, decimal Journal/Other Current Assets input, grouped history, protection and normal XLSB Save As/reopen. Evidence: `obj/ClientReportTests/44c52e7eab314b50bd39677748f70f7a/test.log`. Source hash unchanged. Existing third-party Chrome_WidgetWin_0 shutdown message remains; not claimed repaired.
- The final helper-only compatibility edit replaces null-conditional calls with equivalent explicit null guards to allow the .NET Framework CodeDom regression. Both configurations were rebuilt and both diagnostic gate tests rerun afterward.
- Demo, Blank and user AGL SHA-256 values match the baseline in the input review. No original workbook was saved. XML/JSON parsing and whitespace checks pass.

## Client document

Output: `D:/Downloads/Summit v2.51 comments - review v01.docx`. Identical committed copy: `Library/Client_Review_2.82_v01.docx`.

SHA-256: `48BAABBE8686628953DA5799CB41F8B52E752A498CAE66BAEDC933C6227C0EFF`.

The original `Summit v2.51 comments.docx` remains byte-for-byte unchanged, SHA-256 `955D44B6C714ED9BA6D5F9B4560C983764B38B83161A71B88C2BF400211BD9ED`. All original issue wording and order are retained. Five empty trailing paragraphs were removed only from the revision to avoid a blank final page.

25 reviewed items are marked orange as believed fixed; only the decimal/paste part of the mixed Survey Input issue is coloured. Original orange markings remain qualified where current client-display confirmation is needed. Remaining original green retains its agreed meaning, not a fix certification. 44 item-specific review notes/questions distinguish open, partly repaired and ambiguous work. The generator is `Tools/Prepare-ClientReviewDocument.py`; its paragraph mapping is tied to this exact source revision and it refuses to overwrite outputs.

Packaged LibreOffice rendering is unavailable on this Windows runtime. A private Word instance opened the revision read-only with macros disabled, exported a QA PDF and closed without saving. Bundled PDFium rendered all ten content pages; every page was visually inspected. After removing the inherited blank page, all ten content-page PNG hashes remained identical to the inspected versions. QA files live under ignored `obj/ClientReviewDocument/v01`, not in the delivered document.

## Remaining client work

Retest orange items on disposable copies, with physical keyboard/mouse and the client's restored/maximised window and display scaling. Excel/VBA recalculation and accountant acceptance remain separate. Conditional-format refresh, older AGL Other Fees mapping, dependent fee-description inputs and the document's unresolved layout/wording questions are not marked fixed.

At this client-testing checkpoint, remind the user of the deferred recent-files HTML launcher; it is not implemented in this delivery.
