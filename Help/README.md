# Summit Help library

The Help button in each DataInterfaceTemplate opens `index.html` with the current GroupStructure ID, ChildStructure ID, active tab, and worksheet. The HTML application then displays the matching interface guidance.

## Editing guidance

- Edit `data/overrides.js` for durable client-authored summaries and additional HTML. These entries take precedence over generated wording.
- Edit `assets/summit-help.css` to change the appearance.
- `data/interfaces.js` is generated from `Structure.xml` and read-only metadata from the authoritative XLSB. Do not put manual changes there because regeneration replaces it.
- Run `Tools/Generate-SummitHelp.ps1` from the repository root after approved changes to `Structure.xml` or the authoritative workbook.

Example override:

```javascript
window.SummitHelpOverrides = {
  "0:0": {
    summary: "<p>Use Global Assumptions to set the organisation and plan start date.</p>",
    additionalHtml: "<div class=\"notice\"><strong>Important</strong><p>Review model dates after calculation.</p></div>"
  }
};
```

The generator opens the XLSB read-only with Excel macros and events disabled and closes it without saving. It exports only user-facing sheet comments and User Guide text; it does not export formulas, workbook values, or VBA.
