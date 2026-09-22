param([string]$Workbook=(Join-Path (Split-Path -Parent $PSScriptRoot) 'Library/Blank BP v26_0001.xlsb'))
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot;$bin=Join-Path $repo 'bin/Release'
$refs=@('System.Core','System.Drawing','System.Data')
$refs+=@('DevExpress.Spreadsheet.v25.2.Core.dll','DevExpress.Docs.v25.2.dll','DevExpress.Data.v25.2.dll','DevExpress.Office.v25.2.Core.dll','DevExpress.Drawing.v25.2.dll','DevExpress.Printing.v25.2.Core.dll')|ForEach-Object {Join-Path $bin $_}
foreach($dll in $refs | Where-Object {Test-Path $_}){[void][Reflection.Assembly]::LoadFrom($dll)}
Add-Type -ReferencedAssemblies $refs -TypeDefinition @'
using System;
using System.Linq;
using System.Text.RegularExpressions;
using DevExpress.Spreadsheet;
public static class StructuralRanges {
 public static void Inspect(string path){using(var w=new Workbook()){
  w.Options.CalculationMode=WorkbookCalculationMode.Manual;w.LoadDocument(path);
  foreach(var n in w.DefinedNames.Where(n=>Regex.IsMatch(n.Name,@"^(IR_ServChg_01|Rep_ServChg_02|LastUnitSCColumn|IR_Spec_Inc_Ass1|LastSIDriversCol|Rep_SInc_01|LastSIAssumpCol|IR_Oth_Inc_Ass[56]|Rep_OInc_04|LastOIWorkingsCol|LastOIAssumpCol|IC_JointVenture_01|LastJointVentureCol|IC_IntercoFunding_0[12]|LastInterco(Loan|Invest)AssCol|DateInterF0[123]|InterCoDatesFund)$",RegexOptions.IgnoreCase)))
   Console.WriteLine(n.Name+"\t"+n.RefersTo+"\t"+(n.Range==null?"":n.Range.RowCount+"x"+n.Range.ColumnCount));
  foreach(string sheet in new[]{"Unit Service Charges","Specific Income Drivers","Other Income Workings","Joint Venture Assumptions","Interco Funding Assumptions"}){
   var s=w.Worksheets[sheet];Console.WriteLine("SHEET "+sheet);
   foreach(var c in s.Range.FromLTRB(0,0,15,8).ExistingCells.Where(c=>c.HasFormula||!c.Value.IsEmpty))Console.WriteLine(c.GetReferenceA1()+" "+(c.HasFormula?c.FormulaInvariant:c.Value.ToString()));
  }
  foreach(var n in w.DefinedNames.Where(n=>n.Name.StartsWith("IC_ServChg_")||n.Name.StartsWith("IR_JointVenture_")))Console.WriteLine("EXTRA "+n.Name+" "+n.RefersTo);
  foreach(var s in w.Worksheets)foreach(var c in s.GetUsedRange().ExistingCells.Where(c=>c.HasArrayFormula&&c.FormulaInvariant.Contains(":")))
   if(s.Name.StartsWith("JV "))Console.WriteLine("ARRAY "+s.Name+"!"+c.GetReferenceA1()+" range="+c.GetArrayFormulaRange().GetReferenceA1()+" formula="+c.FormulaInvariant);
 }}
}
'@
$before=(Get-FileHash -LiteralPath $Workbook).Hash
[StructuralRanges]::Inspect($Workbook)
if((Get-FileHash -LiteralPath $Workbook).Hash -ne $before){throw 'Source changed'}
