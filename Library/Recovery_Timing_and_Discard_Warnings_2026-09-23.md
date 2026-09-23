# Recovery timing and discarded Check Sheet warnings — tests 2.78–2.79

## Agreed behaviour

The user corrected the initial timing answer: **Always every X minutes may
interrupt active work**, after the current edit/operation finishes. When idle
is the alternative, non-interrupting start policy. The later request adds a
clickable **Snooze until idle for 1 minute** before a recovery write starts.
Neither policy makes native workbook serialization a background operation.

Options > Recovery backup now has independent When idle for and Always every
checkboxes and minute editors (1–120). Either or both may be selected. Existing
master enablement is retained; new settings default to idle enabled, two
minutes, and maximum interval disabled with a ten-minute value. The maximum
policy is not silently enabled for existing installations.

Idle measures Windows-session keyboard/mouse activity, not just Summit events.
The maximum interval can become due during activity, but never bypasses pending
editors, dialogs, normal saves, integrity work, structural/grouped edits, model
health, Check Sheet policy, recovery-mode or notification-owner guards. Only
new, committed, unsaved user revisions qualify. Successful or failed attempts
restart both timers; failures retain the previous completed recovery copy and
back off. Save As retargets and resets the schedule.

## Snooze opportunity

A native, non-modal notice precedes each eligible recovery write by at least
five seconds (the scheduler polls at five-second intervals). It does not take
focus when shown. It names the plan and offers Snooze until idle for 1 minute.
Closing the notice also snoozes. No workbook is exported while it is displayed.
Integrity work waits while the clickable notice is present.

2.79 adds **Press Esc to cancel** to the notice. Plain Esc in a Summit control
cancels the pending save and invokes the same one-minute-idle snooze, even when
focus remains in the working grid. A temporary UI-thread message filter is
installed only while the notice is visible and removed on close/disposal.
Other keys, system shortcuts and unrelated modal dialogs retain their normal
handling; the consumed Esc is not also delivered to the grid editor. This uses
the documented [WinForms message-filter lifecycle](https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.imessagefilter?view=netframework-4.8.1).
No global keyboard hook or in-flight write cancellation is introduced.

Debug/Release 2.79 builds pass with zero warnings/errors. Both native recovery
fixtures exercise the registered filter pipeline with editor-targeted messages,
verify that Esc preserves editor text and existing recovery bytes, and check
filter removal, modal/system-key exclusion and continued-activity snoozing.
Evidence: `obj/recovery-escape-debug.log`, `obj/recovery-escape-release.log` and
`obj/recovery-escape-{Debug,Release}-build.log`. The updated notice render was
visually checked. Physical-keyboard/client acceptance is still required.

Snooze cancels that plan's pending write and preserves the existing recovery.
It overrides an overdue Always deadline until a full minute has passed since
the click AND the Windows session has been idle for a full minute. More input
extends the wait. Then a new advance notice is shown, allowing another snooze.
A user revision changed during the notice restarts its grace period. Closing
the model, disabling recovery or entering an unsafe operation removes the
pending notice. A completed attempt returns to the configured timings.

Once the native workbook write starts, it cannot safely be interrupted; the
existing progress/completion notice replaces the advance notice. Normal user
Save/Save As commands are not snoozed or changed by this feature. Snooze is
session-local and per plan, not a persistent disable setting.

## Remembered Check Sheet failures

On reopening a held plan, the heading says **Check sheet: recheck required**.
A once-per-opening warning distinguishes a previous session's failure from a
newly verified failure and offers an integrity check. Yes targets that exact
model even if periodic checks/recovery are disabled; No leaves the warning and
hold intact. A fresh result replaces the remembered state. Existing override,
continue-on-error and optional Save-after-clearance policies remain.

If the saved file was acceptable when opened and no normal Save or Save As
has completed in this session, closing can explicitly discard the session's
warning together with its unsaved changes. Cancel does not clear anything.
Actual close clears only after the original's SHA-256 still matches the bytes
fingerprinted before and after loading. Changed, inaccessible, previously
failing/held, saved or unsafe/recovery-mode files are excluded.

The opening baseline reads the loaded Check Sheet; it is **not a new full
calculation or financial certification**. Existing remembered holds are not
cleared from cached results. A fresh successful check before any user changes
can establish an acceptable opening baseline. A permitted backup of known
errors receives its own path's persisted hold, so discarding the original
session does not incorrectly clear that recovery file's warning.

## Engineering evidence

- Debug and Release builds: `obj/recovery-lifecycle-Debug-build.log` and
  `obj/recovery-lifecycle-Release-build.log`.
- Recovery fixtures exercise idle/maximum boundaries, active-input maximum
  saves, revision deduplication, independent preference persistence, failure
  backoff, actual snooze-button/close actions, renewed grace, input extension,
  modal/edit/save/integrity guards, disable/close cleanup, and existing atomic
  recovery/history/metadata preservation. Timing samples/deadlines are injected
  deterministically; this is not a physical-keyboard idle-clock acceptance test.
  Logs: `obj/recovery-timing-debug.log`, `obj/recovery-timing-release.log`.
- Native synthetic Check Sheet tests cover cancelled close, discard, normal
  saves before/after failure, pre-existing failure/hold, fresh pre-edit pass,
  external changes, failed saves, known-bad recovery, persistence in a fresh
  process, Yes/No prompt targeting and no repeat prompt. Logs:
  `obj/recovery-lifecycle-integrity-debug.log` and Release counterpart.
- Private AGL full-load/UI test verifies opening fingerprint/baseline, both
  remembered/current captions, exact Check Sheet link target and immediate
  warning removal. Evidence: `obj/IntegrityCompatibility/919fe3c8e51a40b1b1ca94224505f170`.
  Source and private input bytes remained unchanged. A WebView shutdown-class
  diagnostic followed the passed assertions; the fixture exited successfully.
- Existing staged-integrity, fast-save state, native Save-button binding and
  integrity-target selection suites pass in `obj/recovery-regression-*.log`.
- Options renders cover 75/97/100/150/200 percent and an 18pt default-font case.
  New captions use native [DevExpress best-size measurement](https://docs.devexpress.com/WindowsForms/DevExpress.XtraEditors.BaseControl.CalcBestSize?v=25.2).
  The notice and options screenshots were inspected; physical mixed-DPI/client
  acceptance remains required.

AGL, repository Blank and repository Demo hashes match the prior checkpoint.
No client/master workbook, VBA or formula/schema was changed. Test settings
belong to isolated fixture executables, not the user's Summit configuration.

## Client test

1. Enable recovery and Always every 1 minute. Edit a disposable plan and keep
   working. On the advance notice, click Snooze; continue input beyond a minute
   and confirm there is no recovery write. Stop for a full minute and allow the
   renewed notice to proceed. Check the progress/completion and System Messages.
2. Try idle-only and both timings, close the notice, then disable recovery while
   a notice is visible. Test with an active editor, another model and a dialog.
3. From a healthy unsaved-on-disk baseline, create a deliberate Check Sheet
   failure and run Integrity. Close without saving, then reopen the unchanged
   original: the discarded-session warning should be gone. Cancel close and
   saved/pre-existing failures must retain their warnings instead.
4. Reopen a remembered failure, choose Yes, and confirm the exact model is
   checked. A successful check clears it and still offers the optional Save.

Financial acceptance and trusted Excel/VBA round-tripping remain separate
release gates. Native snapshot-export time remains the principal recovery
interruption; snooze controls when it begins, not how long it takes.
