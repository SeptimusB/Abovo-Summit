# Model Manager definition trial — Summit 2.33

## Scope and access

Open **Compare Business Plans > Model Manager > Open Model Manager**.

This is the persistent definition/review layer of the proposed client model manager, not an automated bespoke migration engine. Existing Compare and Upgrade commands are unchanged. Manager rules do not feed the current population service yet. Embedding XML does not apply formulas, expand ranges, populate assumptions, calculate, or approve a migration.

The trial supports:

- Generic versus client-specific classification, separately from Template versus PopulatedModel role.
- Stable definition identity and immutable local revisions; latest revisions are shown by name in Saved definitions.
- Portable XML import/export, and reading/writing that definition as an XLSB custom XML part.
- Source, original-template and latest-template filenames and SHA-256 fingerprints. Paths are session-local, so the portable definition does not expose machine-specific paths. Reselect files after reopening to verify the evidence.
- Editable review items: description, worksheet, source/target named ranges, controlling master range, related ranges, data type, disposition and notes.
- Dispositions Unreviewed, PreserveCustomisation, AssumptionInput, CapacityOnly and DoNotCarryForward. All definitions remain Draft in this release. These dispositions record human intent, not executable approval.
- Replacing evidence with different file content requires confirmation and resets rules to Unreviewed.
- Native grids with cell multiselect and clipboard copy.

Local revisions are stored under `%LOCALAPPDATA%\Abovo\Summit\ModelDefinitions\<definition GUID>\revision-000001.xml`. This is a local library, not yet a shared SharePoint registry. Use XML export/import to transfer definitions. Existing revision/output filenames cannot be overwritten.

## Multiple XML parts in XLSB

Verified that the repository Blank and supplied Stori each already contain three custom XML parts: Office/SharePoint properties, content type schema and form-template metadata. Neither inspected source has embedded Summit Structure XML. The current Structure reader therefore still falls back to its packaged Structure.xml.

The new part has root `Abovo_Model_Manager`, namespace `urn:abovo:summit:model-manager:1`, schema version 1. It is independent of `Abovo_Model_Def` used for the UI structure. A synthetic workbook containing Structure XML, an unrelated XML part and the new manager definition was also tested, so coexistence with actual Structure XML is verified separately from the real-file tests.

DevExpress exposes `Workbook.CustomXmlParts`, including XLSB support: [official documentation](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Workbook.CustomXmlParts). Tests used the locally installed 25.2.4 assemblies, not only the online documentation.

The authoring service copies the saved XLSB, adds/updates only its owned part, registers package relationships/content types as required, then verifies every other original package part by SHA-256 before publishing the new filename. All existing worksheet, VBA and SharePoint parts must remain byte-identical. It refuses package digital signatures, duplicate manager definitions, unknown schemas and existing output paths. VBA project signatures, if present inside an untouched VBA part, are not modified by this package-only operation. Normal later workbook saves still follow the existing Summit/Excel serializers.

XML parsing prohibits DTDs/external resolution and limits document size. Unknown elements/attributes and missing identity fields are rejected. There are no script/formula execution hooks in this XML. Full-model open reads the optional metadata into `ExcelModel.ManagedDefinition`; malformed metadata reports a system warning without blocking normal model use. No hidden sheets or load-time workbook migrations are introduced.

## Trial files

Originals remain unchanged:

- Stori: `D:\Downloads\Stori BP v25_0704 2026-27 V4 update 100826 CLEAN pass 3.xlsb`
- Blank: `C:\Repos\Abovo Summit\Library\Blank BP v26_0001.xlsb`

Managed copies:

- `C:\Sandbox\Model Manager Trial\Stori - Model Manager Trial.xlsb`
- `C:\Sandbox\Model Manager Trial\Blank BP v26_0001 - Model Manager Trial.xlsb`

Stori contains four Unreviewed prompts covering stock capacity, management-cost capacity, grouping/cost drivers and bespoke covenants. They are review starting points only, not assertions that the provisional old baseline or executable migration is approved. Blank is marked Generic/Template. No assumptions, formulas or business results were changed in these copies.

## Password access

`ModelManager.RequirePassword=false` in App.config explicitly disables the gate for this trial, in both build configurations, as requested. Before live deployment, configure a salted verifier and set the flag to true. Missing/invalid flags are treated as requiring a password; missing/malformed verifiers deny access. The verifier uses PBKDF2-HMAC-SHA256 with 200,000 iterations, a random 16-byte salt and a 32-byte derived key. `ModelManagerAccess.CreateVerifier` produces a verifier; do not store plaintext passwords or credentials in source/workbooks.

This is an application entry-point gate, not workbook encryption or protection against somebody who can modify Summit's executable/configuration. Production deployment must protect the configuration and establish credential provisioning/reset policy before claiming an administrative security boundary. Passwords are deliberately not embedded in model XML.

## Validation performed

- Debug and Release AnyCPU builds.
- Native form construction and bitmap layout inspection (not a substitute for interactive multi-DPI testing).
- Separate XML definition serialization/import, immutable revision sequence and embedded readback.
- Missing manager part is a normal result; missing identity, duplicate definitions, unsupported schema, unknown elements, invalid client classification and attempted Approved status are rejected.
- Existing output and source overwrite attempts are rejected.
- Correct/incorrect password verification and missing-verifier rejection.
- Synthetic XLSB: Structure XML + unrelated part + manager part survive DevExpress save/reopen; the existing Structure reader still picks the correct root; updating manager XML does not add duplicates.
- Stori copy: 1,355 original parts unchanged (excluding the two relationship/content-type registration parts).
- Blank copy: 1,389 original parts unchanged (same exclusions).
- Both real-file copies: DevExpress -> XLSB -> Microsoft Excel -> XLSB -> DevExpress retained four custom XML parts and manager identity/classification/rule count. Originals and managed test inputs remained unchanged. Excel used a separate automation instance, macros/events disabled, manual calculation, read-only open and SaveCopyAs to disposable test files.

Reproducible scripts: `Tools/Test-ModelManager.ps1`, `Tools/Test-ModelManagerRoundtrip.ps1`. Disposable workbook/form-render artifacts are under ignored `obj/ModelManagerTests`. The latter needs installed desktop Excel. It does not execute macros or validate financial results.

## Manual test and risk plan

1. Launch `bin/Debug/Abovo-summit.exe` (2.33); reach the new manager tab, import each managed copy and check Generic/Client, role and Stori review items.
2. Edit a rule and notes, save two revisions, close/reopen, select by name, and import the earlier revision through Import XML. Confirm cancel/discard prompts protect unsaved edits.
3. Reselect each evidence file; check matching hashes. Select a changed copy deliberately and confirm the warning and Unreviewed reset. A matching hash establishes file identity, not financial correctness.
4. Save to a new XLSB filename, confirm the source/template stay unchanged, and reopen the result in Summit. Check `ManagedDefinition`; try corrupt metadata on a disposable copy and confirm a warning rather than failed model open.
5. Exercise multiselect, copy and enum editing on the rules and saved-definition grids, and layout at client DPI. No normal workbook cell formatting was changed.
6. Check normal Excel/VBA menus and workbook behaviour manually on copies; automation kept macros disabled. Financial output/Check Sheet parity is not claimed by these metadata-only tests.
7. Before live use, test password configuration, wrong/cancelled credentials, missing verifier, and configuration permissions. Trial gate is deliberately off.
8. Future execution work must consume reviewed, versioned rules; lock source/template inputs read-only; modify a fresh result; use established structural transactions; test master/related-range formula contiguity; report source formula inputs and unsupported special cases; calculate and compare detailed SOCI/Check Sheet. Never interpret Generic or a disposition alone as permission to execute an upgrade.

## Remaining stages

Shared model registry/SharePoint integration, reviewed operation handlers, preflight coverage and ambiguity resolution, a controlled three-file migration executor, result/report Save As, approval/version provenance and client acceptance tests remain separate implementation stages. The existing generic population service must not be represented as already satisfying those bespoke guarantees.
