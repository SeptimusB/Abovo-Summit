param([Parameter(Mandatory=$true)][string]$SummitSave,[Parameter(Mandatory=$true)][string]$ExcelCopy)
$ErrorActionPreference='Stop'
$bin=Join-Path (Split-Path -Parent $PSScriptRoot) 'bin/Release'
$refs=@('System.Core','System.Drawing')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Docs.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
foreach($ref in $refs|Where-Object {Test-Path -LiteralPath $_}){[void][Reflection.Assembly]::LoadFrom($ref)}
Add-Type -ReferencedAssemblies $refs -TypeDefinition @'
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DevExpress.Spreadsheet;
public static class SaveRoundtripCheck {
 public static void Check(string a,string b){using(var left=new Workbook())using(var right=new Workbook()){
  left.Options.CalculationMode=right.Options.CalculationMode=WorkbookCalculationMode.Manual;
  using(var s=File.OpenRead(a))left.LoadDocument(s,DocumentFormat.Xlsb);
  using(var s=File.OpenRead(b))right.LoadDocument(s,DocumentFormat.Xlsb);
  if(!left.Worksheets.Select(s=>s.Name).SequenceEqual(right.Worksheets.Select(s=>s.Name)))throw new Exception("Worksheet order changed");
  if(left.DefinedNames.Count!=right.DefinedNames.Count)throw new Exception("Named-range count changed");
  const string pending="Abovo.Summit.ResultsPending";
  if(left.DocumentProperties.Custom.Names.Contains(pending)){
   if(!right.DocumentProperties.Custom.Names.Contains(pending)||left.DocumentProperties.Custom[pending].ToString(CultureInfo.InvariantCulture)!=right.DocumentProperties.Custom[pending].ToString(CultureInfo.InvariantCulture))throw new Exception("Pending-results marker changed");
   Console.WriteLine("PASS: Excel copy preserves pending-results custom property="+right.DocumentProperties.Custom[pending]);
  }
  int count=0;
  foreach(string name in new[]{"Multivariable Dashboard","Detailed Comp Inc - Trad View","Financial Position - Trad View","Cashflow detailed","Check Sheet"}){
   var target=right.Worksheets[name];
   foreach(var cell in left.Worksheets[name].GetUsedRange().ExistingCells){
    var other=target.Cells[cell.RowIndex,cell.ColumnIndex];
    if(cell.FormulaInvariant!=other.FormulaInvariant)throw new Exception("Formula changed: "+name+"!"+cell.GetReferenceA1());
    bool equal=cell.Value.IsNumeric&&other.Value.IsNumeric?Math.Abs(cell.Value.NumericValue-other.Value.NumericValue)<=1e-7:cell.Value.ToString(CultureInfo.InvariantCulture)==other.Value.ToString(CultureInfo.InvariantCulture);
    if(!equal)throw new Exception("Value changed: "+name+"!"+cell.GetReferenceA1());count++;
   }
  }
  Console.WriteLine("PASS: Excel copy reopened natively; worksheet order, name count, "+count+" selected cells/formulas/caches preserved; no explicit calculation.");
 }}
}
'@
[SaveRoundtripCheck]::Check($SummitSave,$ExcelCopy)
