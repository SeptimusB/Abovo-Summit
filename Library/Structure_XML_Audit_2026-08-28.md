# Core BP Structure.xml audit

Audit date: 28 August 2026

Source: `C:\Repos\Abovo Summit\Structure.xml`

Contract workbook: `Library\TestFileMigrated.xlsb`

## Scope and method

This report was regenerated from the current user-revised XML rather than carried forward from the previous conclusions.

- The XML was parsed with line information and checked for structural counts, expansion metadata, duplicate scalar siblings, navigation targets, identity anomalies and known serializer tag-casing defects.
- The repository workbook was opened in a separate hidden, macro-disabled, read-only Excel instance. It was closed without saving.
- Workbook worksheets, defined names, name casing and named-range worksheet ownership were compared with the XML.
- The authoritative master and repository compatibility copy were hashed independently.
- No workbook, application code or XML content was modified by the audit.

## Current structure

- Groups: 3
- Child structures: 217
- Interface sections: 490
- Datasources: 529
- Cell-range sources: 959
- Declared worksheets: 210

## High-confidence inconsistency

### Housing Asset retains contradictory expansion metadata

`RowsExpandModel=NRRI` at line 15283 and `RowExpandByNR=DepnType` at line 15284 define the operative row-expansion contract. `RowsExpandBy=CC` remains at line 15288.

`RowsExpandBy` is not represented in the current `ISEDatasource` serializer, so the `CC` value is ignored and conflicts with the active metadata.

## Duplicate scalar elements

No exactly identical scalar sibling groups remain.

Seven parent nodes still contain repeated scalar properties with differing values or content:

| Element | Affected parents | Evidence |
| --- | ---: | --- |
| `FieldName` | 2 | Empty versus `Movement` at lines 3581/3596; empty versus `Year` at lines 4090/4104. |
| `PositID` | 2 | Values 7 and 5 at lines 5238/5240 and 6804/6806. |
| `TipText` | 2 | Semantically equivalent whitespace variants at lines 433/441; empty versus populated at lines 11878/11880. |
| `RepeatingHeaderText` | 1 | `Void Rate` versus `Capital Grant Income Year` at lines 2560/2573. |

These are scalar serializer properties, not collections. Each group therefore requires an explicit choice of intended value.

## Navigation and identity inconsistencies

- The Service Charge link targets the nonexistent section `Real Service Charge Increases / (Decreases)` at line 973. The actual section is `Real Service Charge` at line 10146.
- Taxation Assumptions contains two sections named `Corporation Tax Rates` at lines 13200 and 13400. The second has `ISName=Capital Allowances and Gains`.
- Workings contains two children named `Development Stock`, with CSIDs 6 and 58 at lines 16983 and 20323.
- `Funding Assumptions V2` uses CSID 138 while occupying zero-based child position 34. The resolver supports this, but it remains inconsistent with the remainder of the group.
- `Repairs rephasing ` contains trailing whitespace at line 3461.
- `Development Build & On Cost  Profiling` contains doubled whitespace at line 9351.
- `Leasholder Income Assumptions` and related fields misspell “Leaseholder”, beginning at line 1887.

## Other suspicious metadata

- The current XML revision introduces trailing whitespace on 34 lines, from line 396 through line 15774. This does not appear to change XML values, but it causes `git diff --check` to fail and makes future reviews noisier.
- `New Lettings Other Disposal Numbers` still uses the copied field name `DemolitionNumber` at line 16856.
- Material child `Name`/`CSName` differences remain in Repairs rephasing, Repairs & Maint Rates, Development Consol Options and Development Build/On Cost Profiling.
- Sensitive workbook credential material remains stored directly in the root XML at line 6. Its value is deliberately not reproduced.

## Resolved findings

- Both `ImpairmentDirect` sources now match `Development BP Assumptions`.
- Both annual and periodic non-tenure expenditure sources now match `Non Tenure Capital Assumptions`.
- Repairs now uses the locally represented `RepIncStkCat` expansion master at lines 3204 and 3216.
- Unsupported/misleading `RowsExpandBy` values were removed from Other Fixed Asset and Other Current Asset.
- The 28 conflicting summary `BandID` declarations were removed.
- The one-off `DataFormat=0` declaration was removed.
- Known case-sensitive serializer tags remain canonical.
- Exactly identical duplicate scalar siblings remain eliminated.

## Checks passed

- XML is well formed.
- All declared worksheets exist in the contract workbook.
- All genuine `NRDSName`, `RowExpandByNR`, `OffSetNR`, `RepeatingNR` and `LiveGridSourceName` references resolve and match workbook casing.
- No named-range worksheet-ownership mismatches remain.
- No known noncanonical serializer tags remain.
- Group IDs and child IDs are unique.
- There is no `MRCI` token. The supported `NRCI` expansion occurs in the dedicated Joint Venture interface.
- `Z:\Sandbox\TestFileClean.xlsb` and `Library\TestFileMigrated.xlsb` remain byte-identical: 11,841,039 bytes, SHA-256 `B248B1C733E1E3293536FBE1DBC9576D56FD4D34BEB74A3898B7F3D7333BCFBE`.
