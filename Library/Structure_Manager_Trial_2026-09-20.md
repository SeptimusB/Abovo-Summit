# Structure Manager authoring trial — 2.46

## Entry and scope

Open **Compare Business Plans → Structure Manager → Open Structure Manager**.
The separate Summit process accepts `--structure-manager`, checks the existing
Model Manager access gate, and opens only the authoring window. It does not open
the normal main screen or its Debug auto-open workbook. Password enforcement
remains explicitly disabled for the agreed trial; the live gate and verifier
settings are unchanged. The gate is UI access control, not workbook encryption.

Three resizable native panes provide:

1. A searchable structure tree, retaining group-scoped child IDs.
2. A read-only SpreadsheetControl and a property grid for the selected XML node.
3. A real DataInterfaceTemplate instance rendered by the existing Summit data
   and presentation services, with editing and action commands disabled.

Load the workbook first. Its normal profile chooses embedded structure where
available, otherwise packaged Structure.xml. **Load structure XML** can select
a different draft against the same workbook. Selecting a tree binding selects
its workbook range and the appropriate DIT section. Clicking ordinary Grid/VGrid,
live Grid or mapped-table preview cells locates their resolved workbook address;
the tree selects a matching range binding where available. This is range-level
linking, not a complete reverse map for every field/custom control. Multiple
bindings may legitimately refer to the same cell.

Editable properties in this trial are interface/section/datasource captions,
navigation aliases, default worksheet, range bindings, read-only flags, field
format, tooltip, units and minimum display width. Other properties, including
IDs, period expansion definitions and custom commands, are inspection-only.
The password-bearing legacy RejData element is not displayed in the property
grid; existing XML content is preserved, not copied into documentation.

**Use selected cells** requires confirmation and an existing contiguous binding
of the same dimensions. Named bindings require an exact existing name match;
ambiguous names must be entered explicitly in the properties. No name or range
is created/resized in the workbook. Direct property changes also reject geometry
changes when the previous binding resolves. A broken old binding can be repaired
to an existing valid range. Mixed month/year progression is never inferred from
column position.

Apply properties rebuilds the selected interface preview. Save draft as writes
a **new XML file only**; it refuses existing paths, including the loaded XML.
The original XML tree, comments, attributes and unmodelled elements are preserved
instead of reserializing the narrower runtime classes. XML parsing rejects DTDs
and external entities and limits document length. IDs and changed references are
validated; this is not yet a whole-workbook semantic validator or publication gate.

## Isolation and persistence

- A byte-for-byte temporary XLSB copy is loaded in the dedicated process. The
  source is opened for read access only while making that copy.
- Existing models in the ordinary Summit process cannot enter this registry.
  The authoring session refuses to start if its process already has open models.
- Spreadsheet UI Save, Save As, Open, New and file-drop operations are disabled
  using the native [DevExpress operation restrictions](https://docs.devexpress.com/WindowsForms/401195/controls-and-libraries/spreadsheet/operation-restrictions).
- The workbook and DIT previews are read-only. The private ChangeManager also
  rejects single and batch writes. DIT actions and mapped-table navigation are
  disabled in this mode; ordinary Summit instances keep existing behaviour.
- Closing/replacing the preview disposes the old DIT before its services. Closing
  the session releases the model and deletes only its exact temporary XLSB and
  empty private directory. A process/OS crash may leave a temporary copy; there is
  no broad cleanup operation against the user's temporary directory.
- Unsaved XML/property edits prompt before discard. No production XML publication,
  workbook save, load-time schema migration, new metadata sheet or calculation
  engine change is introduced.

The existing **Definitions / bespoke rules** editor remains available, including
generic/client classification, three-file evidence, immutable revisions and its
existing metadata-only XML embedding into newly named copies. Its metadata and
the structure draft are not yet an executable, unified upgrade contract.

## Deliberate first-trial limits / next increment

- Specialist class forms (Analyser, FFR, Stress Test etc.) are inspectable as XML
  but are not instantiated as DIT previews.
- No add/remove/reparent interface authoring yet; no automatic CSID renumbering.
- No workbook formula editing, structural expansion, master/dependent-range
  orchestration, arbitrary script execution or automatic bespoke migration.
- No one-click publication/embedding of the authored Structure XML. Exported
  drafts require review before installation; XMLVer is preserved, not presented
  as a newly released production definition.
- Binding validation does not prove business meaning or calculation equivalence.
  Future work must add stable semantic identity, explicit master-range families,
  mixed-period axes, reviewed bespoke rules, approved capabilities and a frozen
  migration plan, while preserving source/template read-only and result-copy rules.

## Verification performed

Debug and Release builds passed at 2.46. The native fixture
`Tools/Test-StructureAuthoring.ps1` passed with the authoritative Blank in Debug
and Demo in Release. It exercises:

- private temporary copy and read-only UI/central single-edit and paste guards;
- actual DIT creation, caption edit/rebuild and a valid binding rebuild;
- tree → worksheet and DIT cell → worksheet/range-tree selection;
- invalid binding/type rejection, atomic rollback preserving tree identities;
- scoped duplicate-ID rejection, prohibited DTDs, unknown XML/comment roundtrip;
- new-file-only save and refusal to overwrite the output;
- shared-model session refusal, native pane bounds at 1280/1900/3000 pixels;
- Stock, Check Sheet and Development Profiling tab/preview switches, with one
  active DIT calculation registration and read-only controls after each switch;
- cleanup, deletion of the owned temporary copy and reopen/close in one process;
- unchanged source workbook and production Structure.xml SHA-256.

The actual Release `--structure-manager` startup produced only the Structure
Manager form (checked by its process-owned window title), then closed cleanly.
The existing Model Manager tests passed (including password fail-closed behaviour,
metadata package preservation and native save/reopen), as did the populated Demo
FFR/Stress Test history regression (edits, undo/redo, retained state and cleanup).

Whole-form DrawToBitmap is not visual acceptance evidence for the spreadsheet:
its native rendering paints outside its child bounds in that capture path. Native
control bounds were checked, but the physical 5k/mixed-DPI appearance needs review.

## Required manual acceptance / risks

1. Keep a normal model open, launch Structure Manager from Compare, and open the
   same saved file. Confirm no normal window/model state changes. Unsaved normal
   session values are intentionally not included: this reads the saved XLSB.
2. Check the three panes maximised/restored/on the 5k screen and a client display.
   Drag both vertical dividers and the properties divider; check legibility,
   selection visibility, scrolling and long paths/captions.
3. Find Stock, a funding VGrid, Survey Input, Check Sheet and Development Unit
   Profiling. Select their bindings, click preview cells and compare addresses.
   Confirm monthly periods and later annual records retain their explicit mapping.
4. Edit a caption and tooltip, apply and inspect; change to a valid same-size
   binding. Try a missing name, wrong worksheet, incompatible shape and invalid
   type. Rejections must not change the draft. Shared ranges are not all globally
   retargeted: each referencing node remains a separate interface definition.
5. Try typing, paste/cut/delete, add-lines, mapped links, Ctrl+S/Save As and file
   drop in the previews. Copy and selection must work; workbook mutation/save and
   secondary navigation must not. Test valid/invalid range-selection confirmation.
6. Export a fresh XML filename, reopen it, confirm properties/comments survive;
   try an existing output filename and cancel pending-edit/close prompts.
7. Repeatedly switch interfaces and close/reopen workbooks, then close the manager.
   Confirm ordinary Summit remains usable and its source files unchanged.
8. Before any client publication, test the live password configuration in a
   disposable installation (missing/wrong verifier must fail closed). Review the
   exported XML, then separately run Summit save → Excel/VBA → Summit roundtrip,
   Check Sheet and SOCI comparisons on a disposable result. Those publication and
   financial-equivalence checks have not been established by this UI trial.

No source/master XLSB was modified.
