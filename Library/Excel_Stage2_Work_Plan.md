# Stage 2 completion gates

Request: continue Stage 2 until complete. Production 2.98 remains usable and unchanged while isolated builds are qualified. No new client release is implied by an intermediate checkpoint.

Stage 2 means interactive ownership, not more independent save demonstrations:

- [x] Atomic engine value batch API, prerequisite calculations and one compensated failure boundary; isolated native tests pass. Live paste/schedule callers are a separate remaining integration gate.
- [x] Existing ModelChangeManager internal engine binding owns typed edits, journal, dirty revision and Undo/Redo; no second history stack. Normal editor dispatch/opening remains unconnected.
- [ ] Detached, revision-bound values **and formatting/permissions** feed existing native interfaces without replacing formulas or relying on a second calculation engine.
- [ ] Model opening and per-model preference/status select compatible Excel automatically, with policy-respecting opening-only DevExpress fallback.
- [ ] Check Sheet/header and Funding date-dependent editability refresh at the accepted final revision.
- [ ] Close/discard and saves cannot lose engine-owned edits or publish stale state. Stage 3 structural operations stay disabled for this route until implemented, not silently sent to a different workbook.
- [ ] Both configurations and real native fixtures pass; then separately versioned Ready to test build and explicit Jon/Alex manual gates.

Existing checkpoints 2a–2f supply result binding, single-value transactions and guarded terminal saves/history. They do not complete the items above. Do not expose a preference which only changes a status label or bypasses an existing edit/calculation path.

Checkpoint 2g (`Excel_DevExpress_Engine_Stage2g_2026-09-25.md`) adds the real change-manager bridge, effective appearance, calculation-generation checks and a partial shared display-reader migration. Check Sheet/company-warning and Funding date/amount dependency tests pass in both engines on generated fixtures. The engine is still not selected by normal model opening. Continue the remaining consumer/editor/lifecycle gates; do not end at this checkpoint.

The current read surface contains many direct Cell.Value/DisplayText and RangeDataSource consumers. A scalar custom calculation callback has already failed dynamic-spill tests and must not be adopted as a universal cache injection mechanism. Native control adapters must consume detached results, or another supported, fully tested projection must be established.

Checkpoint 2h (`Excel_DevExpress_Engine_Stage2h_2026-09-25.md`) connects one-use native editor admission, explicit paste/cut/clear/copy commands, real standalone controls, dropdown/date reads and common calculation routing. Both native engines pass 133 bridge assertions; the ordinary production route passes 550 native UI regression checks. Next: remaining range/validation consumers and actual native-DIT acceptance, then ongoing save/lifecycle and real opening/options. This is still not an enabled client route or Stage 2 completion.
