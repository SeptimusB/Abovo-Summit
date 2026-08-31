# Core BP Structure.xml audit

Audit date: 30 August 2026

Source: `C:\Repos\Abovo Summit\Structure.xml`

Authoritative master: `Z:\Sandbox\TestFileClean.xlsb`

Repository workbook: `Library\TestFileClean.xlsb`

## Scope and method

This report was regenerated from the current user-revised XML rather than carrying forward the previous conclusions.

- The XML was parsed with line information and checked for structural counts, expansion metadata, duplicate scalar siblings, navigation targets, identity anomalies and known serializer tag-casing defects.
- Both workbooks were opened in separate hidden, macro-disabled, read-only Excel sessions and closed without saving.
- Workbook worksheets, defined names, name casing and datasource named-range worksheet ownership were compared with the XML.
- The authoritative master and repository compatibility copy were hashed independently and their worksheet inventories compared.
- No workbook, application code or XML content was modified by the audit.

## Current structure

- Groups: 3
- Child structures: 217
- Interface sections: 490
- Datasources: 529
- Cell-range sources: 959
- Distinct declared worksheets: 210

## High-confidence findings

No high-confidence missing worksheet, missing named-range, serializer-casing, expansion-contract or broken-navigation defect remains in the current XML.

## Duplicate scalar elements

No exactly identical scalar sibling groups remain.
No conflicting repeated scalar sibling groups remain.

## Accepted design exceptions

- `Funding Assumptions V2` deliberately uses temporary CSID 138 while V1/V2 evaluation remains in progress. It is excluded from findings.
- The two Workings children named `Development Stock` represent separate workbook workings that both exist in the XLSB. Their shared display name is excluded from findings.
- Plaintext root `RejData` storage is an explicitly accepted client risk given the client's XML-editing capability. Its value remains deliberately omitted and it is excluded from findings.

## Remaining hygiene observation

- The XML contains trailing whitespace on 38 lines: 396, 408, 469, 5966, 6274, 6292, 6310, 6329, 6351, 6370, 6389, 6407, 6425, 6443, 6462, 6484, 6503, 6522, 6803, 7345, 7664, 7682, 7700, 7719, 7741, 7760, 7779, 7797, 7815, 7833, 7852, 7874, 7893, 7912, 15771, 16448, 16461 and 16462.

## Resolved by the latest XML edit

- The duplicate Stock New Lettings `TipText` has been removed.
- The conflicting Capital Grant `RepeatingHeaderText` values have been removed.
- The second Taxation section is now correctly named `Capital Allowances and Gains`; no duplicate section names remain within a child structure.
- Both broken GOTO targets now match their actual target sections.
- The unsupported conflicting `RowsExpandBy=CC` has been removed from the Housing Asset expansion datasource; its operative `RowsExpandModel=NRRI` and `RowExpandByNR=DepnType` contract remains.
- The duplicate Repairs preliminary-rate `FieldName`, both duplicate `PositID` pairs and the empty Funding Details `TipText` have been removed.
- The remaining duplicate Capital Expenditure `FieldName` has been removed; no conflicting scalar groups remain.
- Leaseholder spelling, Repairs naming/spacing, Repairs & Maint. Rates, Development Consol and Import Options, and Development Build On Cost Profiling Assumptions have been reconciled; no child `Name`/`CSName` mismatch remains.
- The two Other Disposal live-grid fields now use `Disposal Numbers`; the second also has integer `DataFormat=I` metadata.

Previously resolved findings remain resolved: the Impairment and non-tenure worksheet ownership corrections, Repairs expansion master, Other Fixed/Current Asset cleanup, Development summary `BandID` cleanup, `DataFormat=0` removal, serializer tag casing and identical scalar duplicates.

## Updated workbook baseline

- `Library\TestFileClean.xlsb` has been replaced with an exact copy of the revised authoritative `Z:\Sandbox\TestFileClean.xlsb`.
- Both files are 11,814,925 bytes with SHA-256 `E58DC1CF506422C598E68628E71A0FFCC701CB935B0E42C1C768566FBB72144C`.
- Both contain 283 worksheets and 1,758 defined names, including `TDB Snapshot` and `TSB Comparison`.

## Checks passed

- XML is well formed.
- Both workbooks contain every one of the 210 distinct worksheets declared by the XML.
- All genuine `NRDSName`, `RowExpandByNR`, `OffSetNR`, `RepeatingNR` and `LiveGridSourceName` references resolve and match workbook casing in both workbooks.
- Named-range-backed datasource ownership agrees with the workbook worksheet owning the name.
- The 13 `NRDSName=CR` entries are intentional explicit-cell-range sentinels, not missing workbook names.
- No known noncanonical serializer tag casing remains.
- Group IDs and child IDs are unique within their groups.
- There is no `MRCI` token. The supported `NRCI` expansion occurs in the dedicated Joint Venture interface.
- The authoritative master and repository compatibility workbook are byte-identical.

## Evidence boundary

The post-copy workbook inventory and name checks were read-only. Neither workbook was calculated or saved during validation. The additional `TDB Snapshot` and `TSB Comparison` sheets are recorded as part of the user-supplied revised master; their formulas and business purpose were not separately audited here.
