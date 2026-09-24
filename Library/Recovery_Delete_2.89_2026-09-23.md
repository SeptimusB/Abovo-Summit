# Recovery discovery: delete option — 2.89

Status: Ready to test in Release. Debug also builds and passes in `bin/Debug-RecoveryTest`; installation into normal `bin/Debug` is pending closure of the user's running session. No existing client recovery copy was removed during development.

When a verified newer recovery copy is found, a native Summit-styled prompt offers **Open recovery**, **Open original**, **Delete recovery…**, and **Cancel**. Paths and saved times are displayed. Open original leaves the recovery untouched; Cancel opens neither. Delete asks a second, default-No confirmation. Successful deletion opens the original without modifying it or turning off future recovery saves.

Deletion is restricted to the exact displayed prefixed or supported legacy recovery filename and requires matching Summit recovery provenance, an available original, unchanged displayed size/timestamp, no file-level reparse point, and no Summit/exclusive-file lock. Changed, unrelated, malformed or open copies are refused. Failure or cancellation returns to the choices. Only the selected copy is removed when both naming generations exist.

Deletion requests the Windows Recycle Bin using [Microsoft's FileSystem.DeleteFile API](https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.fileio.filesystem.deletefile?view=netframework-4.8.1). The confirmation explicitly warns that deletion may be permanent on network drives or locations without a Recycle Bin. No permanent-delete retry is implemented. Original workbook data, XML and VBA are not edited.

## Verification

`Tools/RecoveryPromptFixture.cs` passes 31 assertions in Release and staged Debug: actual native choices; cancellation and declined deletion; original byte preservation; malformed/mismatched/missing-original/changed/open-file guards; and actual Windows deletion of synthetic prefixed and legacy copies, without deleting the other copy. The native prompt screenshot was inspected for readable paths, wrapping and button fit. Only generated disposable recovery files were recycled; these remain recoverable through the local Windows Recycle Bin.

- Release: `obj/ClientReportTests/5fa598b0653c4e3aaede622bd837d86b`.
- Staged Debug: `obj/ClientReportTests/efcd10bf095343b99c71a6c83fae944c`.
- Quiet-build diagnostics regression retained. Existing dirty-worktree changes are preserved; no commit/push performed.

## Stable functional test 53 — Jon

On a disposable plan with a newer recovery, confirm all four choices. Decline Delete and confirm the recovery remains. Confirm Delete and verify only that copy is removed and the original opens. Check that Open original retains the recovery and Cancel opens neither. Test your usual display scaling and storage location; network/permanent-delete behavior has not been exercised automatically. Existing functional items 1–52 keep their numbers.
