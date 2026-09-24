param([Parameter(Mandatory=$true)][string]$SavedWorkbook,[string]$Configuration='Release')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$bin=Join-Path $repo ('bin/'+$Configuration)
$out=Join-Path $repo ('obj/FundingScheduleExcel/'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $out)
$roundtrip=Join-Path $out 'excel-roundtrip.xlsb'
$hash=(Get-FileHash -LiteralPath $SavedWorkbook).Hash
function Release-Com($item){if($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($item)}}
$excel=$null;$seed=$null;$book=$null
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false
 $seed=$excel.Workbooks.Add();$excel.Calculation=-4135;$excel.CalculateBeforeSave=$false
 $book=$excel.Workbooks.Open($SavedWorkbook,0,$true)
 $book.SaveCopyAs($roundtrip)
 $book.Close($false);Release-Com $book;$book=$null
} finally {
 if($book){$book.Close($false);Release-Com $book}
 if($seed){$seed.Close($false);Release-Com $seed}
 if($excel){$excel.Quit();Release-Com $excel}
 if((Get-FileHash -LiteralPath $SavedWorkbook).Hash -ne $hash){throw 'Input workbook changed'}
}
$refs=@('System.Core','System.Drawing','System.Xml','System.Xml.Linq')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Docs.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
foreach($ref in $refs|Where-Object {Test-Path -LiteralPath $_}){[void][Reflection.Assembly]::LoadFrom($ref)}
Add-Type -ReferencedAssemblies $refs -TypeDefinition @"
using System;
using System.Linq;
using System.Xml.Linq;
using DevExpress.Spreadsheet;
public static class FundingScheduleRoundtrip {
 public static void Check(string before,string after){using(var a=new Workbook())using(var b=new Workbook()){
  a.Options.CalculationMode=b.Options.CalculationMode=WorkbookCalculationMode.Manual;
  a.LoadDocument(before);b.LoadDocument(after);
  if(!a.Worksheets.Select(s=>s.Name).SequenceEqual(b.Worksheets.Select(s=>s.Name)))throw new Exception("Sheet order changed");
  if(a.DefinedNames.Count!=b.DefinedNames.Count)throw new Exception("Name count changed");
  foreach(var n in a.DefinedNames){var other=b.DefinedNames.GetDefinedName(n.Name);if(other==null||n.RefersTo!=other.RefersTo||n.Hidden!=other.Hidden)throw new Exception("Name changed: "+n.Name);}
  foreach(var part in a.CustomXmlParts){var root=part.CustomXmlPartDocument.DocumentElement;
   if(root.NamespaceURI!="urn:abovo:summit:funding-schedules:1"&&root.NamespaceURI!="urn:abovo:test:unrelated")continue;
   var other=b.CustomXmlParts.SingleOrDefault(p=>p.CustomXmlPartDocument.DocumentElement.NamespaceURI==root.NamespaceURI);
   if(other==null||!XNode.DeepEquals(XElement.Parse(root.OuterXml),XElement.Parse(other.CustomXmlPartDocument.DocumentElement.OuterXml)))throw new Exception("Custom XML changed: "+root.NamespaceURI);
   Console.WriteLine("PASS: Excel retains custom XML "+root.NamespaceURI);
  }
  var source=a.Worksheets["Funding Assumptions"];var target=b.Worksheets[source.Name];int checkedCells=0;
  if(source.IsProtected!=target.IsProtected)throw new Exception("Funding protection changed");
  foreach(var cell in source.GetUsedRange().ExistingCells){var other=target.Cells[cell.RowIndex,cell.ColumnIndex];
   if(!String.Equals(cell.FormulaInvariant,other.FormulaInvariant,StringComparison.OrdinalIgnoreCase))throw new Exception("Formula changed: "+cell.GetReferenceA1());
   if(!cell.HasFormula&&!cell.Value.Equals(other.Value))throw new Exception("Input/date changed: "+cell.GetReferenceA1());
   if(cell.Protection.Locked!=other.Protection.Locked||cell.Fill.PatternType!=other.Fill.PatternType)throw new Exception("Editability changed: "+cell.GetReferenceA1());
   checkedCells++;
  }
  Console.WriteLine("PASS: Excel read-only open / SaveCopyAs / native reload preserves sheet order, "+a.DefinedNames.Count+" names and "+checkedCells+" Funding cell formulas, constants, dates, locks and fill patterns.");
 }}
}
"@
[FundingScheduleRoundtrip]::Check($SavedWorkbook,$roundtrip)
Write-Output ('OUTPUT='+$out)
Write-Output 'Macros, events and calculation disabled; interactive VBA/financial acceptance remains manual. Input hash unchanged.'
