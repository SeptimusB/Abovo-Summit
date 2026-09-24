# Summit functional test ledger — 2.98 update

Updated 24 September 2026. Stable IDs 1–83 are retained; new items are 84–86. This is the current numbered testing supplement, not a replacement or rewrite of the illustrated Word review v05. Superseded instructions below take precedence over earlier test descriptions.

## How to report results

Reply, for example, **84 pass**, **85 fail — details**, or **68 passed; send to Alex**. Include the workbook and executable version for failures. Partial passes remain partial. A Jon pass moves the agreed scope to **newly ready for Alex**; it never means client agreement. Green/Agreed remains reserved for Alex's returned green Word text. Blue is Jon testing; orange heading/light-blue detail is Alex testing when carried into the next Word revision.

- **Newly ready for Alex:** Jon has passed **1, 2, 3 and 5** and requested **Send to Alex Test Response 3**. Builds/native tests alone are not acceptance; Alex's agreement is still pending.
- **Existing Test response 2 handoff:** 6–7 (save/break/file-warning behavior only), 14 (Economic link indentation only), 16 (read-only Management Costs readability only). Item 5 now moves to Test Response 3. The live Check Sheet heading and TDB remainder are explicitly excluded.
- **Latest feedback:** zoom/scroll and scheduler received positive feedback, but their amended requirements still need the targeted tests below. Alex's double-click failure is not reproduced; no speculative repair was made.
- **Calculation repair:** 10/60/80 are **Ready to test — Jon, 2.98**. The former 2.97 defect is retained in the result log. Fresh full Check Sheet reads, final-revision warning publication and mapped-grid refresh now share a guarded lifecycle. See [2.98 evidence](Check_Sheet_Coherence_2.98_2026-09-24.md). No new Alex handoff is inferred.
- **New work for Jon:** 84 contiguous schedules, 85 read-only Model Version; amended 68/82 wheel shortcuts. 86 is an Alex reproduction/retest, not a new fix.

## Jon's original checks, with current scope

| ID | Brief test action | Current stage / cross-reference |
|---:|---|---|
| 1 | Create Snapshot in an older file missing its sheets: Transactional DB → TDB Snapshot → TDB Comparison. | Send to Alex Test Response 3 |
| 2 | Compare Live/Snapshot/Differences in grids and charts; later edits must not alter the snapshot. | Send to Alex Test Response 3 |
| 3 | Save a separate copy and reopen; snapshot and differences survive. | Send to Alex Test Response 3; Excel round trip also 48 |
| 4 | Open/close interfaces and Options, change scale and restore it; no collection-modified crash. | Fail — reported mapped-table scope: zoom and horizontal scrolling needed. No renewed scale/collection crash reported; distinct from Rent crash 83 |
| 5 | Break, check, close without saving, reopen: discarded Check Sheet finding must not remain. | Send to Alex Test Response 3 |
| 6 | Save a break, reopen, correct and recheck: saved warning state follows that saved file. | Alex — Test response 2; warm refresh excluded |
| 7 | Check Sheet imbalance permits normal Save/Save As and default recovery continuation. | Alex — Test response 2 |
| 8 | Genuine formula/reference errors remain distinct from a soft balance warning. | Jon |
| 9 | Check Sheet overrides stay Yes/No after entry, clearing and Undo/Redo, never 1/0. | Jon |
| 10 | Check Sheet watch interval/idle settings detect changed files and update messages, header and visible grid. | Ready to test — Jon, 2.98; related 60/80 |
| 11 | Explicit trial enables Check Sheet/edit timings; disable after trial acceptance. | Jon; reminder remains applicable |
| 12 | Company-name input usable at normal/restored size. | Covered by amended 61 |
| 13 | Opening Stock description and blank spacer layout. | Covered by 62/74/78 |
| 14 | Service Charge/Specific Income links fit and align. | Indentation: Alex — Test response 2. Caption fitting: Jon |
| 15 | Fixed Summary Other Income Category offers no Add Lines. | Jon |
| 16 | Read-only Management Costs text readable; dependent category appearance refreshes. | Readability: Alex — Test response 2. Refresh/locking: Jon |
| 17 | Management Description/categories pinned; Year 1 heading locked, amounts editable. | Jon |

## Existing Alex test queue

These were already assigned to Alex; they are not newly passed in 2.97 and are not Agreed.

| ID | Brief test action | Current stage |
|---:|---|---|
| 18 | Main-screen Open/New/Compare/circular icons and Program Information fit on Alex's display. | Alex |
| 19 | File summary Start Date, no obsolete Edit link, no unwanted horizontal scrolling. | Alex |
| 20 | Funding summary YE Net Debt/Peak Debt and neighbouring figures align. | Alex |
| 21 | Company-name edit/Undo/Redo updates all headings, file tab and summary. | Alex |
| 22 | DIT Save/Save As; Save disabled when clean and enabled after edit. | Alex |
| 23 | Return is far-right left arrow with separator; History uses curved arrow. | Alex |
| 24 | Navigator labels fit; maximise/restore retains reopening strips and panel state. | Alex |
| 25 | Repeating selectors show Year 1/Year 2 consistently; stored values remain numeric. | Alex |
| 26 | Tab/arrows/Enter/Shift+Enter maintain expected focus and scroll, especially Funding. | Alex |
| 27 | Copy dropdown/single value/rectangle; reject invalid paste; headings and editor closure. | Alex; Funding selection also 73 |
| 28 | Repeated edits do not widen Rent/Voids, Other Income or Management grids. | Alex |
| 29 | Service Charge and Development weekly rent/service charge show two decimals. | Alex |
| 30 | Blank survey dates are blank; entry/clear/Undo/Redo do not show 30/12/1899. | Alex |
| 31 | Survey decimal inputs and rectangular Excel paste retain values. | Alex; layout not signed off |
| 32 | SHG Profiling/Basis independent; corrected completion-period caption. | Alex |
| 33 | Development Expenditure money retains decimals on paste/save/reopen; years integer. | Alex |
| 34 | Negative percentages allowed where the workbook permits, without bypassing rules. | Alex |
| 35 | First Interest Payment Month writes a real month-end date; clear/Undo/Redo. | Alex |
| 36 | Covenant BP Year/Year correct and read-only; permitted amounts editable with decimals. | Alex |
| 37 | Housing Assets, Other Current Assets and Journal amounts retain decimals on reopen. | Alex |

## Wider regression and client acceptance

| ID | Brief test action | Current stage |
|---:|---|---|
| 38 | Recovery idle/maximum intervals, visible notice, Snooze and Esc before writing. | Jon / client environment |
| 39 | Recovery only after user changes; enable/disable and multiple open files. | Jon / client environment |
| 40 | Newer recovery prompt, prior history and guided normal XLSB Save As. | Jon |
| 41 | Locked/unavailable destination reports failure and preserves last recovery. | Jon / client storage |
| 42 | Full Integrity targets selected file; Run Now, idle schedule and between-stage pause. | Jon |
| 43 | Funding add/delete, protected boundaries, save and Excel reopen without new circular references. | Alex structural acceptance; not a new fix |
| 44 | Development add/delete including multi-year, names and TDB alignment. | Jon / Alex structural acceptance |
| 45 | Journals add/delete updates IR_Journals and both TDB mirrors. | Alex |
| 46 | OFA/Repairs insert/delete, CapEx template boundary, Include and service-charge mirrors. | Alex — user said client will test |
| 47 | Stress Test and FFR editing/history/Undo/Redo, switching non-modal windows. | Jon / Alex |
| 48 | Disposable Summit → trusted Excel/VBA → Summit round trip retains functionality. | Jon / Alex; native tests do not replace this |
| 49 | Accountant compares Balance Sheet to Financial Position — Trad View. | Accountant pending; Jon's earlier functional pass retained |

## Subsequent tests and amended instructions

| ID | Brief test action | Current stage / current instruction |
|---:|---|---|
| 50 | Unavailable-cell cue and edit/paste locking follow source fill as description is entered/cleared/undone. | Jon reported improvement; remaining cross-interface acceptance not assumed |
| 51 | Standard/exception Rent Weeks accept positive decimals, paste/Undo; reject zero/negative; clear year relocks. | Jon |
| 52 | Basic Funding schedule open/cancel/apply and duplicate dates. | Retained ID; current rules in 55/66/71/84; old holiday/preview-button instructions withdrawn |
| 53 | Recovery prompt: open recovery/original/cancel/delete; declining Delete preserves copy; confirmed Delete targets only that copy. | Jon; disposable files |
| 54 | Options OK applies/persists; Cancel does not; separate Apply button absent. | Jon |
| 55 | Facility/section schedule chooses correct loan and multiple sections; preview, apply, focus, grouped Undo/Redo. | Jon; use 84 for expansion placement |
| 56 | Schedule tint XML persists through Save As, recovery and Excel/Summit reopening; Excel fills unchanged. | Jon |
| 57 | Scheduler/Options wheel does not scroll the form behind; dialog sizing at restored/5K scale. | Jon |
| 58 | Date and chosen-loan target tints visible; other loans and unavailable-cell cues unchanged. | Jon |
| 59 | Fixed decimals/percent figures, preservation and Undo/Redo. | Covered by 67/72; locked targets now soft-skip, not all-or-nothing |
| 60 | Full Check Sheet watch detects and clears a known break. | Earlier detection/clear passed Jon; header/grid coherence now Ready to test — Jon, 2.98 under 10/80 |
| 61 | Company field bounded in restored/full-screen windows; long name readable and propagates. | Jon |
| 62 | Opening Stock description, spacer and description-driven unlock/relock. | Jon; current width/spacer criteria 74/78 |
| 63 | Combo first click stays open; Shift+Down/Alt+Down/F4; read-only respected; normal keyboard navigation. | Jon |
| 64 | Right-click Add schedule targets the clicked loan. | Jon; old header-removal/old wheel instructions superseded by 68/69 |
| 65 | Facility Name list comes from Facility; Funder list comes from Funders. | Jon |
| 66 | Monthly/quarterly/semiannual/annual dates preserve initial day; inclusive end; short-month clamp without drift; one preview column. | Jon; no holiday lookup/first/last-day options |
| 67 | Enter 123456.789 and negative fixed figures; right loan/sections, focus, tints, duplicate dates, Undo/Redo, save/reopen. | Jon |
| 68 | All grids: ordinary wheel vertical, **Ctrl+wheel zoom**, **Shift+wheel horizontal**; menus agree and popups/modals do not leak scrolling. | Jon — amended in 2.97 |
| 69 | Sticky loan-name strip/tooltip identifies Loan (Funder/Facility); header/section scheduling uses correct context. | Jon |
| 70 | Filter to funder/facility, orange indicator, clear; persist settings; edits still hit correct physical loan. | Jon; also 79 |
| 71 | Alternate occurrence count and inclusive end date; right-click section default and multi-section selection. | Jon |
| 72 | Locked/occupied/formula fixed-amount targets warn softly and skip; other eligible figures/dates kept; no unlock/overwrite. | Jon |
| 73 | Funding multiselect; Copy with Header and Row titles; filtered/unfiltered, editor closure and clipboard values. | Jon |
| 74 | Stock Description approximately half the earlier widened minimum; user resizing retained. | Jon |
| 75 | Tree-style grids use the same zoom/pan/reset without changing data or selection meaning. | Jon; new modifiers 68 |
| 76 | Funding date editors fit at 60/100/200%; maximise/restore; pending edit retained without unintended posting. | Jon; 5K physical check remains |
| 77 | SOCI/Cashflow/Balance Sheet analyser text/totals do not overlap; suitable row heights, bounded first column, manual resize/reset. | Jon |
| 78 | Funding fixed elements have a clear ~3px boundary; Stock blank spacer has no stray bordering line. | Jon |
| 79 | Filter to facility appears with the right name; different funder/facility cases hide correct columns. | Jon |
| 80 | Undo recalculates current/restored source sheet and visible results; Funding G82 → Check Sheet → Undo/Redo → navigate. Header, colours and figures must agree; no transient TDB Cashflow remainder. | Ready to test — Jon, 2.98; ordinary hidden-Check-Sheet edits retain existing calculation path |
| 81 | Check Sheet Summit links single-click to correct destination; choices/Return/copy and Yes/No overrides still work. | Jon |
| 82 | **Ctrl+wheel** zoom follows mouse; minimum 60% remains readable; right-click zoom/set/reset works. | Positive general feedback; modifier change and full matrix still Jon |
| 83 | Rent interface/tab navigation with repeating editors, including rapid switches; no constructor crash. | Jon; Working-2.95 withdrawn, use 2.96+ |
| 84 | **New:** schedule fits earliest wholly empty contiguous block; otherwise extend tail for full schedule + five spare rows. Test scattered holes, another loan's existing amounts, multiple sections, focus and Undo/Redo. | Jon — 2.97. Added blank capacity remains on Undo/failure per existing policy |
| 85 | **New:** Global Assumptions Model Version cannot be typed/pasted over; text selectable; Company Name still editable. | Jon — 2.97 |
| 86 | **Reported, not reproduced:** Alex double-clicks a populated repeating header to copy prior period; compare committed year vs immediately pending edit. | Alex retest / call if still failing; no speculative fix |

## Open development questions — not presented as repaired tests

The Word v05 questions remain: Repairs' three-table refresh; Survey nested scrolling/wrapping; Development scheme-name/default and missing year editors; House Type/Include-dependent appearance; some requested tab moves/captions; Economic rule examples; Funding date hints/first-date locking and Other Fees older-file geometry; Housing Asset tab/component organization. These require a specific example or a scoped repair before entering the ready-to-test queue. Engine research, representative 32GB/client-machine benchmarks and bespoke Structure Manager upgrades remain separate workstreams.

The previously deferred recent-files HTML launcher is still a reminder for the next client-testing release, not implemented here.

## Evidence and reconciliation

- Original numbered 1–49 checklist: this task's response of 23 September 2026, 13:45; labels retained rather than mapping client paragraph IDs to guessed numbers.
- Current illustrated source: [Word review v05](Client_Review_2.93_v05.docx), read without alteration. It preserves only the original client's blank-button item as Agreed; no new green sign-off is inferred.
- Subsequent stable IDs: 2.87–2.96 release reports in Library; overlapping test instructions are explicitly cross-referenced above.
- [2.97 scoped changes and evidence](Production_Changes_2.97_2026-09-24.md).

### Result log

Append dated results against the same IDs; do not renumber or erase failures when a later repair arrives. Jon's pass is not Alex's green sign-off.

- 24 September 2026, **1, 2, 3, 5 — Send to Alex Test Response 3**: Jon's functional passes and explicit handoff recorded. This does not imply Alex's agreement or close the separate live Check Sheet coherence issue.
- 24 September 2026, **4 — fail, reported mapped-table scope**: mapped tables need zoom and horizontal scrolling. The report does not establish a recurrence of the scale/collection-modified crash; keep those scopes separate.

- 24 September 2026, **80 - investigation only; remains open**: five fresh private AGL model lifecycles compared identical observed stale states after Funding G82 edit and Undo. One/two deferred passes and incremental workbook calculation did not restore all observed balances; recursive full calculation did, but warning publication and visible-grid refresh remained separate steps. See [matrix evidence](Check_Sheet_Calc_Matrix_2026-09-24.md) and [proposed refresh design](Check_Sheet_Refresh_Design_2026-09-24.md). Debug 2.97 remains reserved for Jon; no production repair, new test release or Alex handoff is implied.

- 24 September 2026, **80 and related 10/60 — Ready to test, Jon, 2.98**: approved focused repair; both native candidate builds pass 124/124 assertions with the actual 5K maximised host and matching 32-bit address-space flags. Normal Debug and Release installed after Jon closed Debug; runtime versions verified. Retest break -> Check Sheet -> Undo/Redo -> away/back and visible watcher refresh. Check Sheet remains a soft warning; workbook fills/locks and original AGL are unchanged. First Funding rebind can advance revision and require another check; unchanged-revision navigation does not. This supersedes the earlier investigation-only stage, not its evidence or failure history.
