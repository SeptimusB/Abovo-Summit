# Recovery packaging optimisation — Summit 2.77

Ready to test. This is a measured reduction in recovery post-processing, not a solution to the remaining UI-blocking native export. Normal XLSB saving, calculation policy, integrity gates, recovery scheduling and the 2.76 optional Save offer after warning clearance are unchanged.

## Change and safety boundary

`RecoveryXlsmCompatibility.Prepare` previously opened every worksheet XML entry in `ZipArchiveMode.Update` solely to validate its cell-metadata indices. Opening those entries for update caused unnecessary decompression/recompression at disposal. It now validates all the same entries in Read mode, then opens Update mode only for the existing small metadata/relationship/content-type patch. One exclusive FileStream stays open throughout both phases, so the validated private file cannot be replaced or concurrently modified between them. No worksheet/formula/VBA payload is rewritten by this bridge. Existing metadata and no-metadata files return without changing even the ZIP bytes; unsupported profiles and malformed inputs remain rejected before update begins.

The exact verified binary-profile gate, all-cell cm/vm checks, XML limits/DTD prohibition, last-good-copy retention, dirty/user-revision eligibility, history, provenance and atomic publication remain. No checks were removed or deferred. New `[Recovery Packaging Benchmark]` entries separate snapshot export (including state restoration), disk flush, metadata, history, verification and replacement; failures name their stage. The existing total recovery benchmark is retained.

Microsoft describes the distinct archive-mode buffering/write behaviour in [ZipArchiveMode](https://learn.microsoft.com/en-us/dotnet/api/system.io.compression.ziparchivemode?view=netframework-4.8.1). The installed .NET Framework behaviour was measured directly; the change does not require a framework or DevExpress update.

## Measurements

Release, x86, local machine, no debugger. Source AGL 26.0005 SHA-256: `30172E674A491C8E69F8A137286EA246C99912111A474F4C1FBEF77F835E7B47`. Timed package comparisons use the identical private native export, with 64,820 cell metadata indices. Hash verification is outside the timed phases.

| Metadata preparation | Run 1 (cold) | Run 2 | Run 3 | Median |
| --- | ---: | ---: | ---: | ---: |
| Previous implementation | 2,551 ms | 2,420 ms | 2,344 ms | 2,420 ms |
| Read-only scan / small update | 1,803 ms | 879 ms | 875 ms | 879 ms |

About **1.54 seconds / 64% off the median metadata stage**, not 64% off the whole save. The first use also generates the small native XML metadata sample. Read-only worksheet scans alone were 669–773 ms, versus 2,225–2,687 ms in Update mode across these runs. No controlled peak-memory or client-hardware claim is made.

A separate full Summit AGL edit/recovery test measured **14,627 ms** inside the recovery writer: snapshot export 13,573 ms, flush 8 ms, metadata 937 ms, history 97 ms, verification 8 ms, replacement 1 ms. This is not a matched total-save before/after trial. A simpler raw Workbook export took 9,406 ms; it is a different setup and must not be substituted for the real-model export figure. Native export remains dominant and the UI remains blocked during it.

## Validation

- Debug and Release builds pass. One Release copy attempt was blocked by our still-running private fixture; it was rerun successfully after that fixture exited. No user process was stopped.
- Package fixture: original uncompressed payload hashes match for every worksheet, VBA and other part, except the two intentionally updated package indexes; exactly one metadata part is added. A second Prepare is an exact-byte no-op. Ten edge cases cover absent/existing metadata, unsupported cm/vm, unverified profile, dangling relationship/type, malformed/DTD worksheet XML and missing relationships. Rejected/no-op files retain exact bytes and release their exclusive handles.
- Debug and Release recovery synthetic suites pass scheduling, dirty/Undo/history preservation, failure restoration, ownership/collision/locking, notification and metadata-import cases. Existing integrity/optional-save and normal-save state suites pass in both configurations.
- The private populated AGL test passes committed edit -> recovery -> independent native import -> complete formula/constant/name/sheet-order/protection digest match -> full Summit recovered model -> read-only recovered history -> normal XLSB Save As. Original bytes and live dirty/Undo state are preserved.
- Separate Excel automation opens both outputs without repair and saves only new copies, with macros, events, link updates and calculation disabled. Existing custom XML/properties and the single recovery-history part survive all four outputs. All **335 VBA module identities/source hashes** agree with the original across native recovery, native normal XLSB and both Excel copies. VBA container bytes themselves differ after serialization, so container identity is not claimed. No raw VBA source or new extraction dependencies were written into the repository.
- AGL source and repository Blank/Demo hashes are unchanged. Client financial/VBA execution, shared/network storage and physical UI/input acceptance remain manual; this targeted change does not repeat the older full Excel formula-normalisation/value study.

Reproduction: `Tools/Profile-RecoveryPackaging.ps1 -Workbook <source.xlsb>`; `-RawExport` reuses an existing private XLSM export. The script checks its input hash. Normal safety tools: `Test-RecoveryBackup.ps1`, `Test-IntegrityCompatibility.ps1`, `Test-SavePreparation.ps1 -StateTracking -Architecture x86`, and `Test-RecoveryExcelRoundtrip.ps1` under `Tools/`.

Evidence: old package profile `obj/RecoveryPackaging/750b273bb47e43fda07f4dee04d1e3cc/profile.log`; new `obj/RecoveryPackaging/bc767f04ee4f4435893120a6eff0cb62/profile.log`; complete recovery `obj/recovery-packaging-agl.log`, outputs `obj/RecoveryTests/1142c1473ce44bb1a41fc2b30fd86643`; Excel `obj/recovery-packaging-excel.log`, outputs `obj/RecoveryExcelRoundtrip/25c49e6ccab2430595ac89eb5b1e86fe`; module hashes `obj/recovery-packaging-vba.log`. Build and synthetic/state regression logs use the `obj/recovery-packaging-` prefix. These ignored artifacts are not masters.

## Remaining work / next test

On a disposable populated plan, make a normal input edit and allow one scheduled recovery; send the new packaging and total benchmark lines. Confirm progress notification, retained Save/Undo, no repeated backup without a new user edit, and normal recovery reopening. First and subsequent recovery writes should both be measured.

Normal XLSB native export is unchanged. Do not move the live bound workbook into Task.Run or treat SaveDocumentAsync as permission for simultaneous editing: [DevExpress's asynchronous save contract](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Workbook.SaveDocumentAsync%28System.IO.Stream-DevExpress.Spreadsheet.DocumentFormat-System.IProgress-System.Int32-%29) explicitly excludes concurrent document access and permits events on another thread (the online page currently displays 26.1; production remains installed 25.2.4). A larger nonblocking design would require a separately owned, safely captured workbook and measured capture/memory/reopen costs, or a reviewed exclusive-operation UI lifecycle. No such architecture is introduced here. Keep save/recovery interruption reduction open, rather than marking it solved by this small gain.
