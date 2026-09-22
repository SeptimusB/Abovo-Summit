# Client review and unattended validation checkpoints

## Authority and review rules

The client report is `D:/Downloads/Summit v2.51 comments.docx`, SHA-256 `955D44B6C714ED9BA6D5F9B4560C983764B38B83161A71B88C2BF400211BD9ED`. The complete document body was read: 132 paragraphs including blanks, no embedded images, comments, footnotes or tracked changes. Direct text colours include orange `E97132`, green `196B24` and inherited/default text. Colour meaning is awaiting the user's legend; do not infer acceptance from a colour. Paragraph references below are extraction locators, not printed page numbers. The original document is not edited.

The report concerns 2.51; the starting implementation is 2.67. Historical fixes are candidates for regression testing, not grounds to mark client acceptance complete. The bundled DOCX renderer is unavailable on this Windows runtime; review uses complete OOXML text/run-colour inspection, not claimed page-layout verification.

Each item records two independent states: client/user status (including exact source colour), and engineering evidence (unreproduced, reproduced, implemented, automated check passed, or manual acceptance required). A technical pass never implies client or accountant approval.

## Checkpoint gates

1. **CP0 baseline:** preserve and commit the existing 2.67 integrity trial after Debug/Release builds and isolated safety regressions. Exclude local `.codex/`, private workbooks and generated evidence. Record the starting source hashes and leave original workbooks untouched.
2. **CP1 intake:** retain every report item and source colour in a stable issue register. Cross-reference the outstanding work list and known repairs. Gather specific ambiguity questions without blocking independent work.
3. **CP2 recovery and Funding:** automate ten-column insertion, safe due-recovery scheduling, source preservation, recovery reopen/history, normal XLSB Save As and macro-disabled Excel round-trip on disposable copies. Capture timings and exceptions. Do not mutate the user's live process or settings.
4. **CP3 bounded repairs:** reproduce one coherent defect family before editing. Preserve workbook authority; distinguish display formatting from typed numeric storage, protection from missing editors, and user layout requests from financial semantics. Test through real shared services/editors, then build both configurations and checkpoint that batch separately.
5. **CP4 review handoff:** publish pass/fail/blocked results, exact evidence paths, remaining questions and client test steps. Keep performance observations separate from controlled benchmarks and financial acceptance. No push or release publication is implied.

## Stop and defer conditions

- Missing colour legend, unspecified workbook/field or ambiguous intended behaviour: preserve the report and ask; no speculative fix.
- Excel/VBA financial interpretation, bespoke template policy, minimum CPU/OS matrix, deployment/password policy and optional Excel offload: require the corresponding decisions; they cannot be settled by a UI fixture.
- Unexplained formula/name/value differences or source-hash change: stop that branch, keep failure evidence and last good checkpoint. Do not save the original or label a partial run passed.
- Existing client Excel sessions, open Summit files and persistent user settings are not test targets. Use privately owned processes and copies. Never terminate processes by application name.
- Running calculation/serialization is atomic. User input may defer a subsequent scheduled unit; tests must not interrupt a live workbook operation.

## Baseline evidence

22 September: Debug and Release builds pass. Release idle-integrity synthetic safety fixture passes (`obj/IdleIntegrityTests/117d603084554c40a7459081c46ac29e`). Debug recovery synthetic fixture passes (`obj/RecoveryTests/0ab6cd9fc40c47d1b6c8f7e5c3cb4668`), including dirty/Undo preservation, failure retention, unrelated-name collision, idle/save/grouped-edit guards, notification owner and settings cancellation. These are engineering checks, not client or financial acceptance.

Latest client Funding trace: test 2.67 x86 Debug, ten records, 126.484 seconds including 4.929 seconds interface rebuild. Source column shifts 60.772 seconds, TDB sync 36.859 seconds. No recovery event appears in that supplied log. A first-chance InvalidOperationException is not diagnosed without its message/stack; capture it in the private combined test.
