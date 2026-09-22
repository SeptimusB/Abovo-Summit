# Client report issue register

Source: `D:\Downloads\Summit v2.51 comments.docx`

SHA-256: `955D44B6C714ED9BA6D5F9B4560C983764B38B83161A71B88C2BF400211BD9ED`

The original is unchanged. Locators are OOXML paragraph numbers, not page numbers. Run-level colours are retained in the companion JSON. Green means **agreed**, not error-free; orange means the user believes it is fixed; verification is still required. Report 2.51 predates the current implementation. Historical repairs are not client acceptance.

Review gates and subsequent evidence: [checkpoint record](Client_Review_Checkpoints_2026-09-22.md).

## General

### P004

Top icons not fully fitting (Open, New, Compare)

- Source colour: `E97132`.
- Stakeholder status: User believes this is fixed; verification still required.
- Engineering: Not yet independently verified

### P005

Icons a bit large for the circle button?

- Source colour: `E97132`.
- Stakeholder status: User believes this is fixed; verification still required.
- Engineering: Not yet independently verified

### P006

“Program Information” (on main screen) not fully showing.

- Source colour: `E97132`.
- Stakeholder status: User believes this is fixed; verification still required.
- Engineering: Not yet independently verified

### P007

Main screen – Start Date edit link doesn’t work.

- Source colour: `E97132`.
- Stakeholder status: User believes this is fixed; verification still required.
- Engineering: Start-date Edit link previously removed; client visual confirmation pending.

### P008

BP Summary – Funding Status – YE Net Debt to be aligned with YE Peak Debt, etc.

- Source colour: `E97132`.
- Stakeholder status: User believes this is fixed; verification still required.
- Engineering: Funding-status alignment previously changed; client visual confirmation pending.

### P009

BP Summary – has scroll left/right in each window – doesn’t need it.

- Source colour: `E97132`.
- Stakeholder status: User believes this is fixed; verification still required.
- Engineering: Not yet independently verified

### P010

Where does it get “Demonstration Housing” from (top left of Assumptions screen) – does this update when the Global Assumptions name changes?

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P011

Blank button on Assumptions window.

- Source colour: `196B24`.
- Stakeholder status: Agreed (user-confirmed meaning); errors may remain.
- Engineering: Not yet independently verified

### P012

Save button on Assumptions window.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Native Debug Save/Save As regression passed; client acceptance remains separate.

### P013

Assumptions Navigator a tad wider, so “Management Cost Assumptions &gt;” fits on 1 line.

- Source colour: `196B24`.
- Stakeholder status: Agreed (user-confirmed meaning); errors may remain.
- Engineering: Navigator captions/width previously changed; physical DPI acceptance pending.

### P014

Some “From Yr” is “Year 1” (Voids and Bad Debts tabs) other is “1” (Service and Support Charges tab). &lt;the repeating columinplaceeditors&gt;

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P015

“Return” button looks like undo, could this be a “left back arrow” as per internet browser.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P016

Can pressing “Enter” take to next cell?

- Source colour: `196B24`.
- Stakeholder status: Agreed (user-confirmed meaning); errors may remain.
- Engineering: Confirmed Enter/Shift+Enter contract implemented in 2.69: remembered horizontal/vertical axis, visible-tab wrapping, grids/header/standalone inputs and scrolling. Debug/Release native regressions passed, including validation recovery and Undo. Physical keyboard/client-workbook acceptance remains open; see trial and checkpoint records.

### P017

Paste tricky to use, e.g. want to copy one dropdown cell to a range of others, e.g. Management Cost Assumptions tab and Summary Cost Category column.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Single-value multi-cell paste implemented; exact client selection requires regression.

### P018

Validation Check sheet – where is this, as it doesn’t allow user to save with a check outstanding?

- Source colour: `196B24`.
- Stakeholder status: Agreed (user-confirmed meaning); errors may remain.
- Engineering: Check Sheet interface exists; save policy subsequently changed. Exact check needed.

## Global Assumptions

### P021

Hard to enter Company Name.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P022

Check Sheet tab – showing checks when Excel BP is clear?

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Opening Stock

### P025

Wider “Stock Description” column in blank BP.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P026

Gap between “Current Stock Numbers” and “Pre BP-Start Date New Build”.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Rent Assumptions

### P029

Column widths get wider per each input (e.g. Voids tab).

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Shared post-edit refitting defect repaired and tested on dropdown/numeric grids; confirm this Rent interface with the client.

## Service Charge Assumptions

### P032

Formatting on “Click to access Economic Assumptions” indented.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P033

Summary Categories tab - Button to link to inflationary increases not showing all text.

- Source colour: `196B24`.
- Stakeholder status: Agreed (user-confirmed meaning); errors may remain.
- Engineering: Not yet independently verified

### P034

Service and Support Charges tab - “Avg Service Charge per week” format to 2 d.p.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Specific Income Assumptions

### P037

Formatting on “Click to access Economic Assumptions” indented.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P038

Button to link to inflationary increases not showing all text.

- Source colour: `196B24`.
- Stakeholder status: Agreed (user-confirmed meaning); errors may remain.
- Engineering: Not yet independently verified

### P039

“Summary Other Income Category” doesn’t have “add lines” macro in BP.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Other Income Assumptions

### P042

Columns get too wide after typing in numbers.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Shared post-edit refitting defect repaired and tested on dropdown/numeric grids; confirm this Other Income interface with the client.

## Management Cost Assumptions

### P045

Summary Categories tab – formatting of default categories.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P046

Management Costs Assumptions tab – conditional formatting on Summary Cost Cat and SOCI Expense Desc columns.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P047

Management Costs Assumptions tab – year 1 should be locked in both Staff Costs and Other Costs tables.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P048

Column widths get wider per each input e.g. Other Costs table.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Shared post-edit refitting defect repaired and tested on dropdown/numeric grids; confirm this Management Costs interface with the client.

### P049

Need Description column locked, so can see when scroll to right.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Repairs & Maintenance

### P052

Repairs and Maintenance Assumptions tab – conditional formatting on “Summary Cost Category” column.  Table descriptions don’t fit.  Formatting on 3 tables not instant (had to switch away from section).

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P053

Stock Survey Allocation tab – date default as 30/12/1899, as can’t seem to change?

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P054

Survey Input – headers need wrapping and columns narrowing.  Scroll within a scroll.  Can’t seem to paste into input cells (not allowing decimal places?) – this is more than likely going to need to be able to paste in from an Excel table.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Development Assumptions

### P057

Import and Edit Schemes – need to scroll.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P058

BP Input Schemes – not letting me type a Scheme Name. Not picking up default “Unidentified”.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Dvpt Details:

### P060

Header says “Property” – can this be “House Type”. “Identified Scheme” be “Scheme Name” (not picking up.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P061

“Inflate Capital Costs” spelling.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P062

Two headings “Period Completed (to)”, 2nd one should be “Period Units into Mgmt (to)”.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Reproduced and repaired in XML: management-period caption now matches the master row. Editor/range unchanged.

### P063

Conditional Formatting required based on populating first description “House Type”.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P064

Conditional Formatting when “Include?” set to “No”.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P065

Capital Costs tab – space before “On Costs Calculation Method”.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P066

Grant/HFG tab – “Partial Sale of Property %” should be on Capital Sales tab.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P067

Grant/HFG tab – SHG Profiling and SHG Calculation basis seem to populate each other.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Reproduced and repaired: distinct master-verified SHG bindings. Actual dropdown edits and independent Undo passed.

### P068

Capital Sales tab – Sales Value per Unit formatting.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P069

Development Revenue tab – Weekly Rent and Service Charge / Property display as 2 d.p.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P070

Development Revenue tab – Void Rate from Year, Bad Debt Rate from Year.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P071

Development Revenue tab – can Staircasing Units table be a separate tab called “Staircasing Assumptions?

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P072

Can Capitalised Interest input table be a tab on Development Consol and Import Options”?

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P073

Valuations/Impairment tab – “From Year” col should be “% of OMV”.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P074

Rename “BP Inputs” tab “BP Dvpt Summary”

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Dvpt Expenditure Assumptions:

### P077

Not letting input/paste to 2 d.p. to Management Costs table.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P078

Planned Maintenance table not input cells for “From Yr”.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P079

All tables should be “From Year”.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Economic Assumptions

### P082

CPI and RPI tab – each table needs a label whether CPI or RPI.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P083

Real Rents tab – need conditional formatting. Scroll within scroll.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P084

Not allowing negative % inputs.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P085

Conditional formatting.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P086

Column widths of real % increases to be uniform.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P087

Real Dvpt R&amp;M Costs tab – inputs aren’t like BP – is this Real Dvpt Service Costs, and it is missing Dvpt R&amp;M?

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Funding Assumptions

### P090

Cell doesn’t highlight when selected, so not sure of typing in right cell.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Facilities & Loans tab:

### P092

Wider first column to see descriptions.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P093

Scroll left/right not visible with additional columns – have to go to bottom of ‘tab’.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P094

Loan description screen freeze, so when scrolling down, can see which loan entering against.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P095

Jumps back to top of sheet a lot when inputting.

- Source colour: `196B24`.
- Stakeholder status: Agreed (user-confirmed meaning); errors may remain.
- Engineering: Native Funding refresh, focus, scroll, keyboard and subsequent-edit regression passed; physical client acceptance remains open.

### P096

Add/Delete loan columns not visible/prominent.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P097

Facility Name dropdown not correct one.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P098

Loan Description needs to wrap text.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P099

Annuity section needs date formatting/ghost.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P100

Interest Rates “Fixed until and including” needs date formatting/ghost.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P101

First Interest Payment Month – causes issue as this is technically a date, but Summit doesn’t populate BP with a date.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Native month-end date/serial, validation, blank, Undo/Redo and grouped-invalid-date regression passed.

## Variable and Cash Rates:

### P103

Needs wider first column

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P104

First date cell to be locked.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P105

Other Fees – tables locked.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Covenant Assumptions

### P108

BP Year and Year columns missing.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P109

Input cells locked.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Taxation Assumptions

### P112

Capital Allowances tab – Description column not showing.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P113

“Capital Allowances %” heading spell check.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Capitalisation Assumptions

### P116

Management Costs Capitalisation tab – can this be transposed?  And will be “From Yr”.  Scroll within scroll.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P117

Repairs Maint Costs Capitalisation tab – transposed too?  “Depreciate Capitalised Management Costs?” heading should be “Depreciate Capitalised Repairs &amp; Mainte Costs?”.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Housing Asset Assumptions

### P120

Housing Asset Assumptions tab – not allowing decimal places in inputs.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P121

Need a way to enter other component descriptions, which would appear next to Land and Structure.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P122

Housing Asset Grant Assumptions tab – has depreciation tables on.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P123

Remaining Useful Life tab – has grant assumptions.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P124

Housing Asset Depreciation tab – looks ok, includes both depreciation and useful life inputs.  “Useful Life for Structure” heading spell check.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

## Other Current Asset Assumptions

### P127

Doesn’t allow decimal places in inputs.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Reproduced integer rounding in Rep_OCA_01; changed this monetary field only to decimal type M. Actual editor, Undo/Redo and XLSB save/reopen passed. Workbook display format retained.

## Journal Assumptions

### P130

Doesn’t allow decimal places.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Reproduced integer rounding in the Journal Amount field; changed the monetary field to type M, leaving Year integer. Actual editor, Undo/Redo and XLSB save/reopen passed. Workbook display format retained.

### P131

Scroll within scroll.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Not yet independently verified

### P132

“Add lines” row insert doesn’t update the “IR_Journals” range name, and causes a Check Sheet check.

- Source colour: `inherited`.
- Stakeholder status: Unclassified; no acceptance inferred.
- Engineering: Private five-row add, both formula mirrors, save/reopen, delete, noncontiguous delete and damaged-name rejection regression passed.
