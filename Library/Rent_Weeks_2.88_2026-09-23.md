# Rent Weeks — test release 2.88

Status: Ready to test (Jon functional test; not Alex acceptance).

The standard Rent Weeks field (`RentWks`) and exception weeks (`Rep_Rent_03`) now accept positive decimal numbers. The former XML whole-number restriction and 50–54 bounds are removed. Zero, negatives, nonnumeric and nonfinite inputs are rejected before writing. Existing blank clearing remains available.

An opt-in `MinExclusive` XML flag provides a strict minimum without changing inclusive minimum rules elsewhere. It is used only by these two fields. Single-cell entry, native grid editors and clipboard validation share that rule. These editors display their decimal values; workbook number formats remain untouched, so Excel may display a stored fractional value rounded under its original format.

The existing exception-year prerequisite, workbook protection and fill-pattern editability rules and unavailable-cell colours are unchanged. Writes remain typed, calculated and journalled through ChangeManager.

## Validation

`Tools/RentWeeksFixture.cs` passes 58 native assertions for each of Release Demo, Debug Demo and Debug Blank. Cases include 0.125, 49.5, 52.142857, 54.5 and 1000.25; invalid entry; atomic mixed valid/zero paste rejection; undo/redo; exception-year clear/reselect; protection/format preservation; and separate XLSB save/reopen retaining both decimals. Both configurations build. Source hashes are unchanged.

Evidence under ignored `obj/ClientReportTests`: `d5538da389d74ddeb90acadbc8c6faae` (Release Demo), `fe86e77316f94557862a486a435f29e0` (Debug Demo), `560c6af9914e46eba3ec257d367f804a` (Debug Blank). The known embedded-browser teardown error 1412 occurs after successful assertions; this change does not claim to repair it.

## Stable functional test 51

**51. Rent Weeks positive decimals:** enter a decimal in standard Rent Weeks and in an exception with a year selected; verify the displayed/stored value, paste and undo/redo. Confirm zero/negative values are rejected, and clearing the exception year still makes its weeks unavailable. Save a disposable copy and reopen in Summit/Excel. Existing tests 1–50 retain their numbers. Physical keyboard/client and Excel/VBA acceptance remain manual.
