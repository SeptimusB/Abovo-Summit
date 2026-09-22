param([Parameter(Mandatory=$true)][string]$TestOutput,[switch]$CompareOnly)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo 'bin/Release'
$inputFile=Join-Path $TestOutput 'excel-reference-input.xlsb'
$expanded=Join-Path $TestOutput 'funding-expanded.xlsb'
$expected=Join-Path $TestOutput 'excel-grouped-reference.xlsb'
$roundtrip=Join-Path $TestOutput 'excel-roundtrip.xlsb'
if(!$CompareOnly -and ((Test-Path $expected) -or (Test-Path $roundtrip))){throw 'Use a fresh fixture directory; existing evidence is not overwritten'}
$sheets=Get-Content (Join-Path $TestOutput 'funding-sheets.txt')
$excel=$null;$seed=$null;$book=$null
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}}
if(!$CompareOnly){try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false
    $seed=$excel.Workbooks.Add();$excel.Calculation=-4135
    $book=$excel.Workbooks.Open($inputFile,0,$true)
    $insert=$book.Names.Item('LoanDescRev1').RefersToRange.Column-1
    $first=$true
    foreach($name in $sheets){$ws=$book.Worksheets.Item($name);$ws.Select($first);$first=$false;Release-Com $ws}
    $ws=$book.Worksheets.Item('Funding Assumptions');$ws.Activate()
    # Reproduce the inspected Funding VBA group operation without executing VBA.
    $from=$ws.Columns.Item($insert);$to=$ws.Columns.Item($insert+7)
    $range=$ws.Range($from,$to);$range.Select();[void]$excel.Selection.Insert(-4161)
    Release-Com $range;Release-Com $from;Release-Com $to
    $range=$ws.Columns.Item($insert+8);$range.Select();$excel.Selection.Copy();Release-Com $range
    $from=$ws.Columns.Item($insert);$to=$ws.Columns.Item($insert+7)
    $range=$ws.Range($from,$to);$range.Select();$excel.ActiveSheet.Paste()
    $excel.CutCopyMode=[Enum]::ToObject([Microsoft.Office.Interop.Excel.XlCutCopyMode],0);$ws.Select()
    Release-Com $range;Release-Com $from;Release-Com $to;Release-Com $ws
    $book.SaveCopyAs($expected);$book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($expanded,0,$true)
    $excel.CalculateFullRebuild()
    $cycle=$excel.CircularReference
    if($null -ne $cycle){throw ('Excel reports circular reference at '+$cycle.Parent.Name+'!'+$cycle.Address())}
    Write-Output 'INFO: Macro-disabled Excel full rebuild returned no circular-reference address; not a financial/VBA sign-off.'
    $book.SaveCopyAs($roundtrip);$book.Close($false);Release-Com $book;$book=$null
} finally {
    if($book){$book.Close($false);Release-Com $book}
    if($seed){$seed.Close($false);Release-Com $seed}
    if($excel){$excel.Quit();Release-Com $excel}
}}
$refs=@('System.Core','System.Drawing','System.Data')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Docs.v25.2.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll')|ForEach-Object {Join-Path $bin $_}
foreach($dll in $refs | Where-Object {Test-Path $_}){[void][Reflection.Assembly]::LoadFrom($dll)}
Add-Type -ReferencedAssemblies $refs -TypeDefinition @'
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DevExpress.Spreadsheet;
public static class FundingComparison {
 public static void Compare(string actual,string expected,string[] sheets,bool allNames) {
  using(var a=new Workbook())using(var e=new Workbook()){
   a.Options.CalculationMode=e.Options.CalculationMode=WorkbookCalculationMode.Manual;
   a.LoadDocument(actual);e.LoadDocument(expected);
   int checkedCells=0,differences=0;
   foreach(string sheet in sheets){
    var ac=a.Worksheets[sheet];var ec=e.Worksheets[sheet];
    var addresses=new HashSet<string>(ac.GetUsedRange().ExistingCells.Where(c=>c.HasFormula).Select(c=>c.GetReferenceA1()));
    addresses.UnionWith(ec.GetUsedRange().ExistingCells.Where(c=>c.HasFormula).Select(c=>c.GetReferenceA1()));
    foreach(string address in addresses){
     string af=ac.Cells[address].FormulaInvariant,ef=ec.Cells[address].FormulaInvariant;checkedCells++;
     if(!String.Equals(af,ef,StringComparison.OrdinalIgnoreCase)){
      if(differences++<8)Console.WriteLine("DIFF "+sheet+"!"+address+" native="+af+" excel="+ef);
     }
    }
   }
   if(differences>0)throw new Exception(differences+" formula differences from Excel");
   if(allNames)foreach(var n in a.DefinedNames){var en=e.DefinedNames.GetDefinedName(n.Name);if(en==null||n.RefersTo!=en.RefersTo)throw new Exception("Round-trip name changed: "+n.Name);}
   Console.WriteLine("PASS: "+checkedCells+" linked-sheet formulas match "+Path.GetFileName(expected));
  }
 }
}
'@
[FundingComparison]::Compare($expanded,$expected,$sheets,$false)
[FundingComparison]::Compare($expanded,$roundtrip,$sheets,$true)
Write-Output ('PASS: Native Funding matches Excel grouped insertion and survives Excel round-trip. '+$TestOutput)
