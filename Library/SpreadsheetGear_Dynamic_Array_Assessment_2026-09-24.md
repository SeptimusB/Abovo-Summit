# SpreadsheetGear dynamic-array conversion assessment

24 September 2026. Read-only assessment, not a production change or test release.

## Exact scope and method

Compared the existing common AGL XLSM input and the existing Excel/VBA ten-column
Funding output from the full Funding trial. No workbook was opened in Excel,
recalculated, changed or saved during this assessment. Production masters remain
unchanged. Counts are specific to this AGL fixture, not a newly audited master.

Evidence directory: `obj/AsposeTrial/funding-d8bcacf09cdf4b308742ce2d4bf6645d`.
Read-only analyzer: `obj/gear_array_readonly_audit.py`.
Input `AGL-input.xlsm` SHA256 from the earlier verified trial:
`62B44DE2764C4D1D1C9C7029C9297EEF4D461E300BAB4AD7DB092BBE1D23193C`.
Reference result: `excel-1.xlsm`.

The analyzer reads worksheet cell metadata associations and array definitions from
the ZIP/XML package. Original cells are matched to the Excel output after the
ten-column shift at AC on the reviewed 32-sheet Funding group and the eleven
whole-row Transactional DB mirror expansions. All eleven mirror start/end
boundaries are independently checked against the resulting defined names. Every
original dynamic-array anchor has a matching output anchor.

## Results

- Original dynamic-array formula anchors: **64,820**, all with a saved **1 x 1**
  array range. This describes current saved extent, not proof of permanently
  scalar behaviour for all possible inputs.
- Original anchors moved and/or changed in formula text or array address:
  **38,006**. This is not a count of financial-value changes or errors.
- Of those, **480** are on the grouped Funding sheets (160 Loan Opening Balances,
  320 Loan Closing Balances); **37,226** are in Transactional DB; **300** are
  elsewhere (80 OW - Charts Source Data, 20 Credit Rating, 200 Interest Payable
  Breakdown).
- Added dynamic-array anchors: **4,510**, all in Transactional DB. Total after
  the full operation: **69,330**.
- 64,473 original dynamic-array formulas contain INDEX. The scanned formulas
  use INDEX, MATCH, IF, ISNA, SUMPRODUCT, SUMIF, SUM and TREND, rather than modern
  spilling-function names. This is promising, but not a semantic conversion proof.
- `xl/metadata.xml` is **733 bytes**, **422 bytes compressed**, in the input.
  The nine custom-XML-related parts total **18,270 bytes** uncompressed. The
  important preservation issue is the worksheet links and formula semantics,
  not the size of this metadata part.

## Interpretation and next validation

Converting appropriate formulas to single-cell legacy CSE arrays is a credible
candidate. Plain scalar formulas should be used only after proving that implicit
intersection does not change their calculation. A currently one-cell array can
still expand for different inputs, so blanket removal of dynamic metadata is not
approved. Fixed-size legacy arrays also need insertion/deletion and Excel/VBA
testing, not just a matching initial balance.

Recommended next experiment: audit and classify the actual current Blank/Demo
masters, create a separate legacy candidate only with user approval, then compare
recalculation, input-boundary changes, Funding/Development add-delete and native
Excel/Gear/DevExpress save-reopen behaviour. Retain all current masters and files.
An approved master conversion would not automatically migrate existing clients.

This does **not** resolve the independent grouped-versus-serial 3-D reference
problem documented in `SpreadsheetGear_Funding_Full_Trial_2026-09-24.md`, nor the
custom XML loss. Gear's existing save-only control matching 38,108 populated
probes is encouraging but does not establish full-model equivalence.

References:
- Microsoft Formula versus Formula2:
  https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/range-formula-vs-formula2
- SpreadsheetGear 9 documentation, dynamic-array limitation:
  https://www.spreadsheetgear.com/support/help/spreadsheetgear.net.9.0/SpreadsheetGear_2023_Limitations.html
