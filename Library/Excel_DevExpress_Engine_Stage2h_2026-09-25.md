# Stage 2h: editor admission and common calculation routing

Stage 2 remains in progress. This is an internal checkpoint, not a client release. Normal Debug/Release 2.98 are unchanged. Continue the remaining gates in `Excel_Stage2_Work_Plan.md`; do not stop at this checkpoint.

## Changes

- `ModelEngineEditTicket` captures the native before-state when an editor opens. It is manager/cell/session/revision/calculation-generation bound, single-use and excluded from XML serialization. Ordinary `ProcessChange` and named-range posts require this ticket on an engine-bound model. A missing, mismatched, replayed or superseded editor cannot overwrite a native value.
- Existing grid, mapped Yes/No, header and standalone registration paths now carry tickets. Header helper capture occurs before accepting keyboard/mouse input. The original fill-based permission and existing XML read-only decisions are retained. Multi-editor rows use their actual focused cell index for native admission.
- `ProcessEngineCommand` captures all explicit paste/cut/clear/copy targets together and uses the previously qualified compensated native batch. It does not borrow an old text-editor ticket. DIT paste, cut/clear and copy-previous-column branches are connected to it. Arbitrary legacy mutation callbacks remain refused.
- Common `CalculateWSs`, `CalcFile`, dependency-sensitive reads, model integrity calculation and Check Sheet watch calculation select the bound owner. Native committed edits/Undo refresh the existing interfaces and Check Sheet/header, without inventing another history stack. A failed presentation subscriber after commit emits a warning, not a false failed/rolled-back edit result.
- Additional audited DIT/custom-control/dropdown cell reads use the current owner facade. Native date display uses DevExpress's supported `ValueObject.GetDateTimeValue` with the workbook's actual 1900/1904 system. This remains a partial read-consumer migration.
- Funding schedule application explicitly refuses the bound trial before any metadata/structural change. Its Stage 3 commands must not modify the stale display workbook. Normal DevExpress scheduling is unchanged.

## Validation

- Both isolated application configurations compile: `bin/EngineStage2h-Debug` and `bin/EngineStage2h-Release`.
- Native bridge: 133 assertions covering both engines, real `ModelPostingTextBox` Enter/post, typed commands, before-state/ticket rejection, date systems, grouped Undo/Redo, dynamic spill, current conditional fill, Check Sheet/header, ordinary calculation/rebuild routing and unchanged source cells. Final logs `obj/EngineStage2h-Debug-bridge-final.log` and `obj/EngineStage2h-Release-bridge-final.log` include a deliberately failing presentation subscriber after commit.
- Release suite also passes batch 48, presentation 33, edit, history 39, candidate save, terminal publication 76, recovery 109, safety 72, deadline, grid and security 12. Per-mode logs: `obj/EngineStage2h-Release-*.log`.
- Existing ordinary production DIT regression passes all **550 checks** on a disposable Demo copy, including save and independent reopen. Evidence: `obj/ClientReportTests/fd4c005cbab9488797d1755721db5212`. A WebView2 shutdown class-unregistration diagnostic remains; process exit is successful. This is normal-route regression coverage, not whole-model native-DIT acceptance.

## Still not enabled or certified

Normal model opening/options do not bind this route yet. Native whole-DIT/analyser acceptance, remaining range-backed readers and validation-expression consumers, continued Save/Save As/recovery/close ownership and exact dirty acknowledgement still need integration. Per-cell Excel appearance transfer remains an end-to-end performance qualification issue. Structural operations, hard-hang process supervision, security/prompt/storage cases and manual Excel/VBA round trips remain their stated gates. Do not infer production readiness from internal fixture counts.
