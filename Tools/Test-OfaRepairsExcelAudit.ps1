param(
    [Parameter(Mandatory=$true)][string]$TestOutput,
    [Parameter(Mandatory=$true)][ValidateSet('OFA','Repairs')][string]$Family,
    [int]$Count=3,
    [switch]$CompareOnly,
    [switch]$AllowVerifiedIncludeRepair
)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$TestOutput=(Resolve-Path -LiteralPath $TestOutput).Path
if(!$TestOutput.StartsWith((Join-Path $repo 'obj')+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Only private obj test outputs are permitted'}
$inputFile=Join-Path $TestOutput 'excel-reference-input.xlsb'
$actual=Join-Path $TestOutput 'expanded.xlsb'
$expected=Join-Path $TestOutput 'excel-vba-reference.xlsb'
$roundtrip=Join-Path $TestOutput 'excel-roundtrip.xlsb'
# Independent contract: inspected current master OFA_Columns/Repairs_Columns.
# No application metadata is consulted to choose anchors or copy modes.
$anchor=if($Family -eq 'OFA'){'LastOFACol'}else{'LastStockCol'}
$sheets=if($Family -eq 'OFA'){@('Other Fixed Asset Assumptions','OFA Workings')}else{
 @('Repairs & Maint. Assumptions','Stock Condition Inputs','Repairs & Maint. Rates','Stock Condition Results','Repairs & Maint. Drivers','Repairs & Maintenance Costs','Repairs & Maintenance Depn','Cost & Depn on Replacement')
}
$hiddenTemplates=if($Family -eq 'OFA'){@('Other Fixed Asset Assumptions')}else{@('Repairs & Maint. Assumptions','Stock Condition Inputs','Repairs & Maint. Rates')}
function Release-Com($obj){if($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj)}}
if(!$CompareOnly){
 if((Test-Path -LiteralPath $expected) -or (Test-Path -LiteralPath $roundtrip)){throw 'Refusing to overwrite existing audit evidence'}
 $excel=$null;$seed=$null;$book=$null
 try{
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false
  $seed=$excel.Workbooks.Add();$excel.Calculation=-4135
  $book=$excel.Workbooks.Open($inputFile,0,$true)
  $at=$book.Names.Item($anchor).RefersToRange.Column-1
  foreach($name in $hiddenTemplates){$ws=$book.Worksheets.Item($name);$ws.Columns.Item($at).Hidden=$false;Release-Com $ws}
  $first=$true;foreach($name in $sheets){$ws=$book.Worksheets.Item($name);[void]$ws.Select($first);$first=$false;Release-Com $ws}
  $ws=$book.Worksheets.Item($sheets[0]);$ws.Activate()
  $range=$ws.Range($ws.Columns.Item($at),$ws.Columns.Item($at+$Count-1));$range.Select();[void]$excel.Selection.Insert(-4161);Release-Com $range
  $ws.Columns.Item($at+$Count).Select();$excel.Selection.Copy()
  $range=$ws.Range($ws.Columns.Item($at),$ws.Columns.Item($at+$Count-1));$range.Select();$excel.ActiveSheet.Paste();Release-Com $range
  $excel.CutCopyMode=[Enum]::ToObject([Microsoft.Office.Interop.Excel.XlCutCopyMode],0)
  $ws.Select();Release-Com $ws
  $template=$book.Names.Item($anchor).RefersToRange.Column-1
  foreach($name in $hiddenTemplates){$ws=$book.Worksheets.Item($name);$ws.Columns.Item($template).Hidden=$true;Release-Com $ws}
  $book.SaveCopyAs($expected);$book.Close($false);Release-Com $book;$book=$null
  $book=$excel.Workbooks.Open($actual,0,$true)
  $book.SaveCopyAs($roundtrip);$book.Close($false);Release-Com $book;$book=$null
  Write-Output 'Reference and roundtrip created with macros/events/calculation/link updates disabled. No financial calculation claim.'
 }finally{
  if($book){$book.Close($false);Release-Com $book};if($seed){$seed.Close($false);Release-Com $seed};if($excel){$excel.Quit();Release-Com $excel}
 }
}
$bin=Join-Path $repo 'bin/Release'
$refs=@('System.Core','System.Drawing','System.Data')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Docs.v25.2.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll')|ForEach-Object {Join-Path $bin $_}
foreach($dll in $refs | Where-Object {Test-Path $_}){[void][Reflection.Assembly]::LoadFrom($dll)}
Add-Type -ReferencedAssemblies $refs -TypeDefinition @'
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DevExpress.Spreadsheet;
public static class OfaRepairsComparison {
 static string Text(Cell c){return c.HasFormula?"F|"+c.HasArrayFormula+"|"+c.FormulaInvariant:"V|"+c.Value.Type+"|"+c.Value.ToString();}
 public static void Compare(string actual,string expected,string[] sheets,bool allNames,bool allowIncludeRepair){
  using(var a=new Workbook())using(var e=new Workbook()){
   a.Options.CalculationMode=e.Options.CalculationMode=WorkbookCalculationMode.Manual;a.LoadDocument(actual);e.LoadDocument(expected);
   int cells=0,diffs=0,names=0,nameDiffs=0,formats=0,formatDiffs=0,approvedRepairs=0;
   foreach(string s in sheets){
    var ac=a.Worksheets[s];var ec=e.Worksheets[s];
    var addresses=new HashSet<string>(ac.GetUsedRange().ExistingCells.Where(c=>c.HasFormula||!c.Value.IsEmpty).Select(c=>c.GetReferenceA1()));
    addresses.UnionWith(ec.GetUsedRange().ExistingCells.Where(c=>c.HasFormula||!c.Value.IsEmpty).Select(c=>c.GetReferenceA1()));
    foreach(string address in addresses){var av=ac.Cells[address];var ev=ec.Cells[address];cells++;
     if(!String.Equals(Text(av),Text(ev),StringComparison.Ordinal)){if(diffs++<15)Console.WriteLine("CELL_DIFF\t"+s+"!"+address+"\tnative="+Text(av)+"\texcel="+Text(ev));}
     formats++;if(av.NumberFormat!=ev.NumberFormat||av.Protection.Locked!=ev.Protection.Locked){if(formatDiffs++<8)Console.WriteLine("FORMAT_DIFF\t"+s+"!"+address+"\tnative="+av.NumberFormat+"/"+av.Protection.Locked+"\texcel="+ev.NumberFormat+"/"+ev.Protection.Locked);}
    }
   }
   foreach(var n in a.DefinedNames){
    var en=e.DefinedNames.GetDefinedName(n.Name);
    if(!allNames&&!sheets.Any(s=>n.RefersTo.Contains("'"+s+"'!")||n.RefersTo.Contains(s+"!")))continue;
    names++;if(en==null||n.RefersTo!=en.RefersTo){
     if(allowIncludeRepair&&!allNames&&n.Name=="RepIncStkCat"){
      var c=a.DefinedNames.GetDefinedName("StockCondCats").Range;var r=n.Range;
      if(r.Worksheet!=c.Worksheet||r.LeftColumnIndex!=c.LeftColumnIndex||r.RightColumnIndex!=c.RightColumnIndex-1||r.TopRowIndex!=c.TopRowIndex-1||r.RowCount!=1)throw new Exception("Invalid approved Include-range correction");
      approvedRepairs++;Console.WriteLine("APPROVED_NAME_REPAIR\t"+n.Name+"\tnative="+n.RefersTo+"\texcel="+(en==null?"missing":en.RefersTo));
     }else if(nameDiffs++<15)Console.WriteLine("NAME_DIFF\t"+n.Name+"\tnative="+n.RefersTo+"\texcel="+(en==null?"missing":en.RefersTo));
    }
   }
   if(allNames)foreach(var ws in a.Worksheets)foreach(var n in ws.DefinedNames){
    var en=e.Worksheets[ws.Name].DefinedNames.GetDefinedName(n.Name);names++;
    if(en==null||n.RefersTo!=en.RefersTo){if(nameDiffs++<15)Console.WriteLine("LOCAL_NAME_DIFF\t"+ws.Name+"!"+n.Name);}
   }
   Console.WriteLine("EXCEL_AUDIT\t"+Path.GetFileName(expected)+"\tcells="+cells+"\tcellDifferences="+diffs+"\tnames="+names+"\tnameDifferences="+nameDiffs+"\tformats="+formats+"\tformatDifferences="+formatDiffs+"\tapprovedIncludeRepairs="+approvedRepairs);
  }
 }
}
'@
[OfaRepairsComparison]::Compare($actual,$expected,$sheets,$false,$AllowVerifiedIncludeRepair)
[OfaRepairsComparison]::Compare($actual,$roundtrip,$sheets,$true,$false)
Write-Output 'Audit complete; EXCEL_AUDIT difference counts determine findings, not this process exit code.'
