# Abovo Summit architecture review

Reviewed 20 August 2026 against the development baseline, Library/TestFileMigrated.xlsb (read-only), Structure.xml, DataManager, PresentationManager and DataInterfaceTemplate.

## Purpose and authority

Summit is a native interface and change-management layer over an authoritative Excel XLSB/XLSM model. The workbook, including names, formulas, styles, protection, validation, embedded VBA and workbook-specific behaviour, remains the source of truth. Summit must not make a workbook unusable when reopened independently in Microsoft Excel/VBA.

AGENTS.md contains mandatory operating rules. The Project Scope Audit records delivery and validation history. This document records stable architectural conclusions supported by code/workbook evidence.

## Ownership map

| Layer | Owns | Must not own |
|---|---|---|
| Workbook contract | Calculation, formulas, styles, locks, validation, Excel/VBA behaviour | Disposable interface state |
| Structure metadata | Declared sections, elements, datasource shape and control intent | Replacement workbook rules/formulas |
| DataManager | DataCellRange bindings, typed values, calculated refresh, validation and rule snapshots | Direct unlogged interactive writes |
| PresentationManager | Presentation/section/element lifecycle and declared control mapping | Calculation semantics |
| Native interface | Layout, formatting, selection/copy, editors and UI events | A parallel model or inferred formats |
| ModelChangeManager | Typed write, audit, dirty state, calculation and rollback | Structural commands outside established services |

## Workbook and metadata flow

The FFR workbook contains compact and calculation-heavy sheets and uses formula cells, constants, locked/unlocked inputs, list validation, notes and resolved formatting as interface inputs. Read-only inspection must disable macros/events/prompts/link updates and close without saving.

PresentationManager finds a ChildStructure by group/child IDs, gets its default worksheet, enumerates InterfaceSections, then processes declared Grid, VGrid, LiveGrid, TextBox, Label, ComboBox and DateBox elements. Structure.xml therefore declares interface intent and binding structure; it never authorises bypassing workbook-first behaviour.

## DataManager responsibilities

DataManager materialises workbook-backed typed datasets and cell bindings, including source worksheet/address, calculated flags, validation lists and rule state. DataCellRange.UpdateCalcs re-reads calculated values from the live workbook. UpdateLocks derives a rule lock from the established solid-fill input signal; it does not replace worksheet/cell protection.

DataValidationsSet resolves list validation from direct ranges, formulas, worksheet/workbook names and literal lists. Native controls must obtain validation choices from that source instead of copying them into UI code.

AbovoRangeDataSource deliberately creates a read-only RangeDataSource with formulas not preserved. It is a display adapter, not an interactive write path.

For workbook Workings displays, LiveGrid resolves its current width from the declared data-field definitions and their repeating named ranges while retaining the configured range as the row/anchor contract. A LiveGrid may also project selected areas from a workbook-defined multi-area `Workings_*` name, optionally prepending the shared year columns; this allows one worksheet block to be split into compact tabs without copying data or exposing intervening helper columns. The resulting read-only RangeDataSource refreshes calculated values in place. Source-worksheet or named-range structural changes invalidate the dependent section, detach the old grid data source and rebuild the presentation atomically.

## Native interface policy

DataInterfaceTemplate owns dynamic DevExpress Grid, BandedGrid and VGrid lifecycle, binding, editors, events and disposal. Its VGrid policy expands the VGrid to visible-content height and lets the containing page own vertical scrolling; wide record sets retain horizontal scrolling.

LiveGrid workbook displays preserve source-cell display text, fill, font treatment and alignment, and use the same cell-multiselect/header-free clipboard-copy policy as other read-only grids.

Use DevExpress Ultimate Edition 25.2 native controls first. All grids are assumed to require cell multiselect and clipboard copy, including read-only grids. A VGrid may not replace a GridView/BandedGrid merely for compactness unless its selection/copy requirements are preserved.

## Mandatory rendering and editing contract

- Render workbook-backed values with their source cell display value, fill, font treatment, number display and alignment.
- Editable means the workbook contract permits it: at minimum an unlocked cell with the established solid-fill input signal. Formula/locked cells stay read-only.
- A single-cell interactive edit creates exactly one DataChangeEvent and passes through ModelChangeManager.
- Reload the relevant snapshot after a calculated change. Never bind an editable DevExpress range directly to the authoritative workbook.
- Retain cell notes as tooltips where the source surface uses them.
- Preserve standalone Excel/VBA round-tripping. Never impair an XLSB/XLSM through a Summit change.

## Change-review checklist

1. Excel/VBA round trip remains intact.
2. Values, formats, locks and validation come from the live workbook, not copied constants.
3. Interactive writes use ModelChangeManager; structural work uses established transaction/structure services.
4. Grids retain multiselect/copy and avoid internal vertical scrolling when the intended complete block is shown by an outer page.
5. Debug and Release builds pass, followed by proportionate manual workbook/Excel validation.

## Known validation gap

There is no automated workbook/UI integration suite. Manual validation remains required for Excel/VBA reopen, older-schema migration, Stress Test scenarios, Transactional DB reconciliation, FFR editing/clipboard behaviour and regulator-template export.
