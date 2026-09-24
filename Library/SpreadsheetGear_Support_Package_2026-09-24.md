# SpreadsheetGear support handoff - v01

24 September 2026. User requested files to email to SpreadsheetGear support;
nothing has been emailed or uploaded. No production application changes or
test-release version increment are involved.

- ZIP: [SpreadsheetGear_Support_Package_v01_2026-09-24.zip](Support/SpreadsheetGear_Support_Package_v01_2026-09-24.zip)
- Draft: [SpreadsheetGear_Support_Email_v01_2026-09-24.txt](Support/SpreadsheetGear_Support_Email_v01_2026-09-24.txt)
- ZIP size: 51,055 bytes.
- ZIP SHA256: `A3F608FA9951D8CEE4B6F5FF2383CDB153C2FA04FD94982C71D21D505ECB70FA`.
- Source: `Tools/SpreadsheetGearSupportRepro`.
- Internal verification/packaging helper: `Tools/Build-SpreadsheetGearSupportPackage.ps1`.

## Vendor response received through the user - 24 September 2026

Tim Andersen, SpreadsheetGear Technical Support, confirms all three reproductions are current product limitations:

- There is no grouped-sheet insertion or alternative API that performs the corresponding 3-D formula-reference fix-ups. A feature request exists, but support gives no implementation roadmap or date.
- Open XML custom XML is not read or written. This is also a feature request without a committed roadmap.
- Dynamic arrays are converted to legacy CSE arrays, losing spill behaviour. Support says preservation of the associated XML cannot be treated as a trivial partial implementation. Dynamic-array support is planned for version 11 after version 10; this is direction, not a release-date commitment or a present capability.

The free-mode insertion restriction is separate from those limitations; a signed trial removes the free-mode range limits, not the unsupported features. Our licensed reproduction already distinguished those cases.

Source: the support email pasted by Jon into this task. No further email, upload, trial activation or vendor contact was made by Codex. This does not resolve the separate DevExpress-checkpoint object-import issue.

Production recommendation: retain DevExpress as Summit's authoritative editing/calculation/save engine. Gear is not a drop-in replacement and must not serialize authoritative customer workbooks. A calculation-only projection backed by a native authoritative worker remains research, with spill-sensitive value edits and correct structural read-back still requiring safeguards. Restoring custom XML or `xl/metadata.xml` alone cannot restore lost formula semantics. No engine migration or production code change is authorised or performed by this documentation update.

## Original reproduction scope

Entirely synthetic minimal reproduction; no customer workbook or proprietary
application/VBA source is required. NuGet SpreadsheetGear 9.3.85 / assembly
9.3.85.102; net48 x64; optional Excel 16.0 build 20326 controls.

1. Two serial column inserts leave a 3-D sum at B1 in both Gear and Excel; Excel's
   grouped-sheet insertion moves it to D1. This asks for a supported grouped
   equivalent, not a claim that Gear differs from Excel serial behavior.
2. An unedited XLSM load/save drops all three parts of a synthetic custom XML item.
3. The same save drops dynamic-array metadata. Excel changes a ROW/INDEX spill
   from three to five rows after an input edit in the original, but the Gear output
   remains a three-cell legacy array. This is a preservation/support inquiry in
   light of the vendor's documented dynamic-array limitation.

## Validation and distribution boundary

Clean ZIP-extracted Debug and Release builds succeed with zero warnings/errors.
Default no-licence runs pass in both configurations; they explicitly skip the
whole-column operation which exceeds free-use limits. The full licensed Release
run plus isolated Excel controls reproduces all three observations. The package
contains exact sample/output files, observations, verification and per-file hashes.
There are 347 recorded pre-packaging assertions and 18 final ZIP-entry hash checks.

Privacy checks inspect both source/evidence files and nested XLSM/XLSX XML. No
licence keys, customer information, workbook/VBA passwords, macros, binary workbook
payloads, external workbook relationships or local absolute paths are included.
Excel's incidental absPath property is removed only from newly generated synthetic
files; no worksheet/formula, custom XML or array metadata is patched. Input hashes
remain unchanged during testing. Existing Excel sessions are left running.

A separate read-only render confirms the fixture labels fit. Its non-Excel
renderer interprets the array anchor differently, so numerical verification relies
on the recorded native Excel controls, not that preview. No renderer export or
recalculation is written to any delivered file.

The original client/master workbooks, production Summit files and unrelated dirty
worktree edits remain untouched. This package does not establish a safe production
engine replacement or perform a new application benchmark. Vendor guidance is recorded above; the original reproduction and delivered ZIP remain unchanged.
