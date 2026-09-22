param([Parameter(Mandatory=$true)][string]$Workbook)
$ErrorActionPreference='Stop'
$bin=Join-Path (Split-Path -Parent $PSScriptRoot) 'bin/Release'
$refs=@('System.Core','System.Drawing','System.Data')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll','DevExpress.Docs.v25.2.dll')|ForEach-Object {Join-Path $bin $_}
foreach($r in $refs|Where-Object {Test-Path -LiteralPath $_}){[void][Reflection.Assembly]::LoadFrom($r)}
Add-Type -ReferencedAssemblies $refs -TypeDefinition @'
using System;
using System.IO;
using System.Reflection;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Functions;
public static class FundingBoundaries {
 public static void Test(string path,string bin){using(var w=new Workbook()){
  var app=Assembly.LoadFrom(Path.Combine(bin,"Abovo-summit.exe"));
  foreach(string t in new[]{"PMCostFunction","ResponsiveCostFunction"}){var f=(ICustomFunction)Activator.CreateInstance(app.GetType("Abovo."+t));if(!w.Functions.GlobalCustomFunctions.Contains(f.Name))w.Functions.GlobalCustomFunctions.Add(f);}
  w.Options.CalculationMode=WorkbookCalculationMode.Manual;using(var s=File.OpenRead(path))w.LoadDocument(s,DocumentFormat.Xlsb);
  w.Options.CalculationEngineType=CalculationEngineType.Recursive;
  // Unsaved sensitivity probe: inject calculated source values, not client inputs.
  // Tests both the first and last newly added ordinary-loan columns.
  w.Worksheets["Loan Interest Paid"].Cells["O8"].Value=7;w.Worksheets["Loan Interest Paid"].Cells["V8"].Value=11;
  w.Worksheets["Loan Repayments"].Cells["O8"].Value=-2;w.Worksheets["Loan Repayments"].Cells["V8"].Value=-3;
  w.Worksheets["Loan Drawdowns"].Cells["O8"].Value=123;w.Worksheets["Loan Drawdowns"].Cells["V8"].Value=456;
  w.CalculateFullRebuild();
  foreach(string n in new[]{"TransCopy_LoanDescsOrd_A","TransCopy_LoanDescsOrd_B"}){
   var r=w.DefinedNames.GetDefinedName(n).Range;if(r.RowCount!=19)throw new Exception("Wrong mirror size");
   double x=r.Worksheet.Cells[r.TopRowIndex+10,16].Value.NumericValue,y=r.Worksheet.Cells[r.TopRowIndex+17,16].Value.NumericValue;
   double ex=n.EndsWith("_A")?5:123,ey=n.EndsWith("_A")?8:456;
   if(x!=ex||y!=ey)throw new Exception(n+" failed: "+x+", "+y);
   Console.WriteLine("PASS: "+n+" first/last added loan results="+x+", "+y);
  }
  Console.WriteLine("PASS: Unsaved boundary probes calculated; delivered workbook never modified.");
 }}
}
'@
$hash=(Get-FileHash -LiteralPath $Workbook).Hash
try{[FundingBoundaries]::Test($Workbook,$bin)}finally{if((Get-FileHash -LiteralPath $Workbook).Hash -ne $hash){throw 'Workbook changed'}}
