# Client input repairs — checkpoint 1

Test release: **2.81**, 23 September 2026. Scope: the user-approved first implementation batch from [the diagnostic review](Client_Input_Review_2026-09-23.md). The earlier 2.80 structural repairs are retained. No commit/push is implied by this checkpoint.

## Implemented scope

- Survey Input: the five Rep_SCond_01–05 cost sources now use monetary storage, not integer conversion. Year/count fields remain unchanged.
- Development Expenditure: Rep_DevRE_01a, 03a, 05a and 07a body values now retain decimals. The independently defined repeating year-header editors remain I; Planned Maintenance percentages remain P.
- Housing Asset: Rep_HAA_04, 05, 05a and 06 monetary inputs retain decimals, including their duplicate presentations. Property counts, amortisation/useful-life periods, calculated offsets and tab grouping are not indiscriminately changed.
- Both ordinary/banded grids and vertical grids display M/SM fields using the actual source cell's DisplayText. This preserves the workbook's currency, precision, sign, blank/zero treatment and scaling. Weekly rent and service charge show two decimals; Debt £Mn retains its three-decimal display. A workbook format that intentionally displays whole pounds still displays whole pounds, without rounding the stored input. Existing conditional-format appearance limitations remain a separate batch.
- Percentage spin editors no longer invent 0–100% limits when XML has no limits. Explicit XML limits remain effective; NOMIN/NOMAX mean no XML bound, not permission to ignore a workbook rule.
- A shared pre-write numeric gate validates finite values, integer types, XML limits and the target cell's Decimal/WholeNumber validation rules, including relative formula criteria. It is used by native grid validation and paste preflight. Mixed invalid pastes are rejected before any destination is written. No temporary worksheet write is used to validate a value, and no whole-workbook integrity scan/rebuild is added per input.
- Covenant table: editable source inputs D:H, with locked/read-only BP Year and Year sourced from A:B. Local offsets change from C-3/C-2 to C-7/C-6 because the existing MergeAcross contract measures from the end of Rep_Cov_01. Shared offset semantics are unchanged. Both Debt monetary fields now retain decimals.
- Funding Other Fees: C238:C240 uses integer **BP Year**, not a calendar date; Excel's current 1–40 whole-number rule is enforced from the workbook rather than hard-coded into the application. Missing fee-description editors are not silently added in this batch. Tests seed a description on their private copy through ChangeManager so the legitimate conditional gate is satisfied.
- Empty Stock Survey dates stay empty during initialisation, refresh and failed-edit restoration. Real dates, clearing, Undo and Redo retain their existing typed ChangeManager route.

## Validation boundaries

The shared gate covers numeric field types in ordinary DIT grid/vertical-grid entry and paste. It is not a new universal Excel validation engine: existing dropdown/month-end validation remains separate; Excel Custom/List/Date/Time rules and standalone/header editor behaviour are not newly rewritten. Blank clearing retains existing behaviour. Existing source-cell protection and conditional eligibility are not relaxed.

All actual workbook edits in the fixtures use production editor/ChangeManager or grouped paste paths on private copies. Original Demo, Blank and AGL files are not written. The small synthetic validation workbook tests relative criteria without modifying the target cells. Test output lives under ignored obj folders, not the source/master directories.

## Evidence

- Release Demo full native run: obj/ClientReportTests/b13a27b43d0847498c51550e5ac7370c, **540 assertions**. Actual editor precision, typed paste, one-step Undo/Redo, dirty/history, source format/protection, 20 Economic percentage views, Covenant source mapping and read-only years, BP-year rejection, date entry/clear, normal Summit save to a distinct XLSB and independent native reopen. This preceded the final paint-only monetary sign/blank and teardown cleanup; the later focused runs use the final build.
- Final-build Demo Covenant/numeric edge run: obj/ClientReportTests/a4946abb8b4e41a389a3b63ebce561ab, **80 assertions**. Includes negative and above-100% multi-cell paste, atomic rejection with no added history, grouped Undo, relative worksheet numeric limits, explicit XML limits, NOMIN/NOMAX, non-finite rejection and no trial-write validation, plus save/reopen.
- Independent read-only Excel inspection of Summit's saved Demo copy: obj/ClientInputExcel/151bf45feaf04f86a81ae3dafc30576e. Excel reads O169/O170 as 2345.67 with two-decimal currency display, Covenant F8 as 115%, H8 as 2345.678 with three decimals, and one-off fee C238 as year 6. Macros/events/link updates/calculation were disabled; closed without saving. This is stored-value/format compatibility, not financial calculation acceptance.
- Expanded master Excel validation/type evidence: obj/ClientInputExcel/957fcd75c175446a805353930dfc7623, 58 cells. Covenant monetary fields permit decimals, while Housing period fields remain explicitly out of scope despite their permissive Excel validation.
- VBA preservation: all **337 module identities/source hashes** match the original Demo in both saved Demo outputs (aggregate 4cc0fa759eb79f66299ff54255b8cdd81ab962a7ad9ab487d9893a8cdacef5ff). The VBA binary container is rewritten, so binary-payload inequality alone is not a source-change test. Existing Tools/Verify-VbaModuleHashes.py ran in memory with dependencies outside the repository; no raw VBA was written or macros executed.
- The initial wider AGL probe (obj/ClientReportTests/d3c5dd82bc1f48769fb8ce3d2746f5b4) intentionally stopped at its no-formula-overwrite test guard: Economic Assumptions!D207 contains **=D31** in this populated file. Read-only Excel confirms that formula in obj/ClientInputExcel/8ec4ff93786d48c69c88658c57fc4aec. It was not overwritten; the original file hash stayed unchanged. The same run also encountered Demo-specific empty/conditionally unavailable Development examples, so no all-AGL-fields pass is claimed. Applicable targeted AGL checks are recorded separately below.

**Additional AGL compatibility finding:** the current XML's static Funding Other Fees cells C238:C240/D238:D240 are not portable to AGL's shifted layout. Read-only Excel shows its B238 is the protected "Indexation to date" label and C238 is protected, not a one-off fee BP-year input. The attempted current-template fee case could not reach that editor (obj/ClientReportTests/b0b16d39a9fc428cbf13428abfdf4698); that run is not counted as a complete pass. Its test-only seed had altered a protected label on the disposable copy, which was discarded without saving. The fixture now guards that setup against protection and limits it to requested fee tests. The original AGL is unchanged. Named-range-relative Other Fees mapping must be addressed with the next dependent-editor batch; no forced unlock or load migration is introduced here. Do not treat the current-template fee-year repair as an AGL compatibility repair.

Final runs:

- AGL Release, applicable Survey/date/Covenant/Housing cases: obj/ClientReportTests/05cf78be74874fcd925f0fe2d6a74f7f, **203 assertions passed**, including saved/reopened values and atomic/relative-limit tests. The older-layout Funding fee case is excluded explicitly, not marked passed.
- Blank Debug, empty Survey date: obj/ClientReportTests/e43ff1c4d21744a09947c6232bf20434, **12 assertions passed**; genuine empty initial value, actual date entry, clearing, Undo/Redo and saved/reopened value.
- Earlier SHG/profile/basis, width, Journal/Other Current Assets decimal regressions: obj/ClientReportTests/10751cac64a44734a235e90ccf4d7e3f, **48 assertions passed**, including saved/reopened monetary values.
- VBA module hashes also match for all **335 AGL** and **339 Blank** modules in their final saved copies. No macro execution/financial recalculation acceptance is claimed.
- Debug and Release build successfully at the normal bin/Debug and bin/Release locations, version 2.81. XML/index JSON parse and git diff whitespace checks pass. Original Demo, Blank and AGL SHA-256 values remain exactly those recorded in the diagnostic review. There is no workbook-source modification, commit or push in this checkpoint.

The embedded browser logs its existing Chrome_WidgetWin_0 shutdown warning (1412) after clean fixture teardown; this is not classified as repaired here. Native test success is not a substitute for client keyboard/mouse/DPI acceptance or accountant sign-off.

## Reproduction

Use Windows PowerShell (.NET Framework), sequentially; native UI fixtures must not compete for focus.

~~~powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File Tools/Test-ClientReport.ps1 -Fixture ClientInputRepairFixture.cs
powershell.exe -NoProfile -ExecutionPolicy Bypass -File Tools/Test-ClientReport.ps1 -Fixture ClientInputRepairFixture.cs -ReviewCases '35'
powershell.exe -NoProfile -ExecutionPolicy Bypass -File Tools/Test-ClientReport.ps1 -Fixture ClientInputRepairFixture.cs -ReviewCases '15,16,35,39' -Workbook 'C:/Sandbox/Insert Comp/AGL - BP 2627 updated Mar26 v25 v2 v26_0005.xlsb'
powershell.exe -NoProfile -ExecutionPolicy Bypass -File Tools/Test-ClientReport.ps1 -Configuration Debug -Fixture ClientInputRepairFixture.cs -ReviewCases '15' -Workbook 'C:/Repos/Abovo Summit/Library/Blank BP v26_0001.xlsb'
powershell.exe -NoProfile -ExecutionPolicy Bypass -File Tools/Test-ClientReport.ps1
~~~

DevExpress's [SpinEdit bound documentation](https://docs.devexpress.com/WindowsForms/DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit.MinValue) explains the default unrestricted 0/0 pair; installed 25.2 assemblies and native tests verify this implementation. Native cell display and numeric validation criteria use the installed Spreadsheet API. No new package is required.

## Next checkpoint / client acceptance

Still open: workbook-based conditional-format refresh, Funding fee-description inputs/dependent eligibility **and named-range-relative Other Fees coordinates for older/populated layouts**, exact locked-cell reproductions, Housing tab grouping, date ghost/hints, first Funding date location, and the separately itemised layouts. Questions in the diagnostic review remain unanswered, not assumed fixed.

Client retest: decimals through keyboard/paste in Survey and Development/Housing costs; negative Economic percentages and Covenant values above 100%; proper Covenant years; weekly money displays; a blank Survey date; one-off fee BP year after a description exists. Save a **new test copy**, reopen in Excel, confirm expected inputs/formats, then reopen in Summit. Trusted Excel/VBA recalculation and financial acceptance remain required; no application repair here changes financial formulas or VBA.
