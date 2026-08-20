# Abovo Summit working rules

## Product contract

- Summit is a VB.NET/DevExpress layer over an authoritative Excel XLSB model for Abovo.
- Most interface functionality is available through DevExpress Ultimate Edition 25.2; prefer its supported native controls and features before introducing custom or third-party implementations.
- Wherever practical, match the XLSB's visual appearance in the native interface unless the user specifically instructs otherwise.
- Increment the test-release version defined by AbovoAppCls/AbovoApp for each test delivery; reset the count only for the final user release.
- Preserve Excel round-tripping: a workbook must remain usable in Summit, then Microsoft Excel/VBA, then Summit again without losing formulas, names, VBA behavior, or user edits.
- Do not make any change to a utilised XLSB or XLSM that could impair its standalone Microsoft Excel or VBA functionality.
- Treat `Library/TestFileMigrated.xlsb`, its embedded VBA, workbook names/formulas/locks, and `Structure.xml` as the behavioral contract. Do not infer parity from similarly named VB methods alone.
- Read `Library/Abovo_Summit_Project_Scope_Audit.md` and `Library/Abovo_Summit_Project_Index.json` before broad workbook, migration, Transactional DB, or Stress Test changes.

## Workbook editing and calculation

- Route interactive single-cell workbook edits through `ModelChangeManager` so writes are typed, logged, marked dirty, calculated, and rolled back on failure.
- Multi-cell and structural commands may write workbook ranges directly only when they preserve workbook-first behavior, entry protection state, dirty state, failure restoration, and the established transaction/structure services.
- An unprotected worksheet in the development workbook is not itself a defect. Respect cell `Protection.Locked` for interface editability and restore each worksheet's entry protection state.
- Normal interactive calculation uses `WBCalcEngine.CalculateWSs`. Dependency-sensitive bulk operations may temporarily select DevExpress `CalculationEngineType.Recursive`, but must restore the previous engine in `Finally`.
- Do not reproduce the VBA `ABVCalculate` early-exit accident; preserve the intended recalculation point.
- Assume every DevExpress grid requires cell multiselect and clipboard copy, including grids that are read-only.
- Source every worksheet-backed grid cell's visible formatting from its underlying spreadsheet cell; do not replace workbook-owned colours, font treatment, number display, or alignment with inferred UI styling.

## Migrations and persistence

- Full-model workbook migrations run in both Debug and Release. Migration failures must fail the open rather than leave a partially reconciled model.
- Migration changes remain in memory until the normal user save. Imported source workbooks must not be treated as full-model migration targets.
- Preserve formula-backed Transactional DB state and the external Excel-edit reconciliation/conflict rules.

## Safety and durable context

- Never store workbook/VBA passwords, raw extracted VBA source, or disposable extraction dependencies in the repository or context pack.
- Do not modify `Library/TestFileMigrated.xlsb` during inspection. Use read-only Excel automation with macros/events disabled and close without saving.
- Preserve unrelated user changes in a dirty worktree. Do not reset or discard them.
- Keep durable architectural conclusions and completed validation results in the Library audit/context pack, but avoid duplicating large source or workbook payloads.

## Validation

- Build both Debug and Release after migration, calculation, or workbook-facing changes.
- The repository has no automated workbook/UI integration suite. Call out required manual round-trip tests, especially older-schema Release migration, Excel/VBA reopen, Stress Test scenario generation/import, sensitivity row expansion, and Transactional DB reconciliation.
