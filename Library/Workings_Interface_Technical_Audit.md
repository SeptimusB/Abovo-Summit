# Workings interface technical audit

Audit date: 2026-08-21

## Delivered scope

The Workings navigator is generated from the ordered list supplied by the product owner and reconciled to `Library/TestFileClean.xlsb`.

- 12 navigator groups.
- 138 child structures with sequential CSIDs 0-137.
- 320 lazy-built, read-only LiveGrid sections.
- 287 sections backed by workbook defined names.
- 31 sections backed by validated direct worksheet ranges where no suitable defined name exists.
- The two original Target Rent prototypes retain their established dynamic XML definitions.
- Every child declares `DefaultWorksheet`, so the worksheet action opens the correct XLSB sheet.

The inaccurate duplicate `RTB Valuation Factors` child was removed. The identically labelled `Development Stock` entries are intentional: one belongs to Stock and one to Development, matching the supplied authoritative list.

Previously omitted children now include Variable Management Costs, Bond Premium Amortisation, all 12 Accounts workings, and the complete Output Workings list.

## Workbook-first projection

`Tools/GenerateWorkingsStructure.ps1` opens the authoritative XLSB read-only in an isolated Excel instance with macros and events disabled, and closes without saving. It does not alter workbook formulas, names, VBA, protection, formatting, or values.

For worksheets with `Workings_*` names, each aligned named block becomes a section. Multi-block children are split into lazy tabs and use nearby workbook text for descriptive labels. Representative splits include:

- Development Capital: 22 tabs.
- Dvpt NonCash: 50 tabs.
- All Schemes FFR: 14 tabs.
- Dvpt Imp FFR Info: 14 tabs.
- OW - Captured Data: Live Stress Information, Base Case, and Test 1-10.

Some legacy names retain an earlier block as well as the block identified by their suffix. The resolver selects the last aligned row block and its modal row extent, preventing unrelated vertical blocks from being combined into one grid.

Worksheets without a suitable data name use a validated direct range derived from their used data region. Formula-driven column captions are read from explicit XML header rows where supplied, otherwise from the two nearest non-empty workbook heading cells above each source column.

All LiveGrid values are supplied through a detached read-only projection. This permits tabs to share Year columns without DevExpress creating overlapping worksheet range bindings. Display text, fill, font, number presentation, alignment, column width, and recalculation refresh remain sourced from the live workbook cells.

## Validation

Automated validation performed against the read-only XLSB:

- Workings child count: 138.
- Workings section count: 320.
- Same-group duplicate child labels: 0.
- Non-LiveGrid sections: 0.
- Missing worksheets, named sources, or direct ranges: 0.
- CSID/index mismatches: 0.
- XML parse errors: 0.

Required interactive checks remain:

1. Open representative single- and multi-tab children in every navigator group.
2. Switch repeatedly between tabs sharing source columns.
3. Recalculate beside editable assumptions and confirm visible values/headings refresh.
4. Add and rename a structural item such as StockType and confirm dependent sections rebuild.
5. Use the worksheet action from representative children in every group.
6. Save, reopen in Microsoft Excel/VBA, then reopen in Summit to confirm round-trip integrity.
