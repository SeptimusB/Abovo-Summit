# DevExpress to Gear import diagnosis - 24 September 2026

Status: isolated, read-only diagnostic; not a production fix or complete compatibility approval. Summit remains on its existing DevExpress-only implementation. Excel is not required by this diagnostic or proposed DevExpress/Gear path.

## Confirmed finding

Gear 9.3.85.102 rejects the DevExpress 25.2.4 AGL checkpoint with `IOException: Corrupt OpenXML document` when importing objects. The same bytes open when its documented `IWorkbookSet.ReadObjects` property is false. Keeping VBA import enabled works; disabling only VBA import does not help.

| ReadObjects | ReadVBA | Result |
|---|---|---|
| true | true | Open fails |
| true | false | Open fails |
| false | true | Opens, 283 sheets |
| false | false | Opens, 283 sheets |

This isolates the failure to a path enabled by object import, not the general worksheet loader. The exact offending XML element is not identified. It does not establish which vendor is at fault. SpreadsheetGear documents that ReadObjects skips charts, pictures and drawing objects, but does not skip chart sheets: [official API](https://spreadsheetgear.com/support/help/spreadsheetgear.net.9.0/SpreadsheetGear2023~SpreadsheetGear.IWorkbookSet~ReadObjects.html).

Read-only package inspection found no missing default Normal style, out-of-range style references, unresolved internal part relationships, duplicate parts or broken content types. Both packages contain 58 drawing XML and 44 chart XML parts. Original custom XML payloads are identical and cell-metadata counts match. Some chart number formats are normalized by DevExpress; this is a candidate for a smaller reproduction, not a proven cause. These checks are not full OpenXML schema validation or visual/VBA certification.

## Calculated checkpoint comparison

Using ReadObjects=false, ReadVBA=true and the established model UDF adapters, Gear opened and fully rebuilt both the immutable Excel-converted AGL baseline and its pre-Funding DevExpress checkpoint. All **38,108** cells across the five established probe ranges matched: no numeric, text or blank/type differences (absolute tolerance 1e-6, relative tolerance 1e-10). This comparison is Gear-before versus Gear-after, not an independent financial or Excel result certification.

One x64 Release pass: baseline load 3.766 s / rebuild 0.694 s; DevExpress checkpoint load 3.720 s / rebuild 0.686 s. No workbook was saved. Both inputs retained their SHA256 hashes. Original AGL, Blank and Demo also retain their pre-trial hashes. Debug and Release builds of the isolated harness passed with zero warnings/errors.

## Recommended boundary, not yet implemented

Subsequent vendor confirmation: [support response](SpreadsheetGear_Support_Package_2026-09-24.md) identifies grouped-sheet reference fix-up, custom XML and dynamic arrays as unsupported features, not configuration errors. The object-import workaround does not resolve them. Retaining DevExpress production authority remains the recommendation; a dual-engine calculation projection is research only.

- Keep the existing DevExpress-only application as the Excel-free baseline; optimise it independently where useful.
- In an optional dual-engine path, DevExpress owns the complete workbook and authoritative serialization; Gear imports a calculation-only copy without drawing objects.
- Never save that reduced Gear copy over the authoritative client file. Its omitted objects, custom XML and dynamic-array limitations remain material.
- Re-run the complete DevExpress-worker Funding trial with this import option, then Development and deletion cases. The current result covers the checkpoint before Funding, not structural compatibility.
- Preserve a native calculation/read-back path for dynamic arrays: previous synthetic tests proved a value-only edit can expand a spill that Gear does not expand. The object-import workaround does not fix that separate limitation.

No production option, native workbook, master, package reference or application version was changed. No vendor contact, commit or push was made.

## Reproduction and evidence

Harness: `Tools/SpreadsheetEngineTrial/GearImportProbe.cs`; commands `gear-import PRIVATE_INPUT NEW_REPORT` and `gear-projection PRIVATE_INPUT NEW_REPORT`. Both accept only private trial paths below ignored `obj/AsposeTrial/`; neither saves a workbook. The projection recalculates only its in-memory instance.

Private evidence under `obj/AsposeTrial/mirror-20260924-v1/dx-agl-2/`:

- `import-options.json`: import matrix and full exceptions.
- `projection-before.json`, `projection-after.json`: calculated probes, timings, source hashes.
- `projection-parity.json`: five-range comparison, zero differences.
- Baseline SHA256: `62B44DE2764C4D1D1C9C7029C9297EEF4D461E300BAB4AD7DB092BBE1D23193C`.
- Checkpoint SHA256: `E31F8354534AFAA71A62C4F7114FDD182687C3A10E64715C153CD92A8362618E`.

Do not distribute private client files or reports as vendor support material without approval. See `Spreadsheet_Two_Engine_Prototype_2026-09-24.md` for the preceding worker tests and remaining production gates.
