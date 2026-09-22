using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Linq;
using DevExpress.Spreadsheet;
public static class Structural3DFixture {
 static Type repair;
 static object Capture(IWorkbook w,string[] sheets,int at,int amount){return repair.GetMethod("Capture").Invoke(null,new object[]{w,sheets,new[]{new KeyValuePair<int,int>(at,amount)}});}
 static void Apply(object saved){repair.GetMethod("Apply").Invoke(saved,new object[]{-1});}
 static void Equal(string actual,string expected){if(actual.Replace("'","")!=expected.Replace("'",""))throw new Exception("Expected "+expected+"; got "+actual);}
 static Workbook Split(bool reuse,int at,int count,int offset){
  var w=new Workbook();w.Options.CalculationMode=WorkbookCalculationMode.Manual;
  w.Worksheets[0].Name="Rates";w.Worksheets.Add("Fees");w.Worksheets.Add("Summary");
  foreach(var sheet in w.Worksheets)for(int c=0;c<12;c++){
   sheet.Cells[0,c].Value=c+1;
   sheet.Cells[2,c].FormulaInvariant="=SUM('Rates:Fees'!$B1:$C$2)+SUM('Rates:Fees'!1:2)+A1";
   sheet.Cells[3,c].FormulaInvariant="=SUM('Rates:Fees'!"+sheet.Cells[0,c].GetReferenceA1()+")";
   sheet.Cells[4,c].ArrayFormulaInvariant="=SUM('Rates:Fees'!$B$1)";
   sheet.Cells[5,c].FormulaInvariant="=SUM('Rates:Fees'!$A$1)+SUM(Rates!A1:B2)";
  }
  w.DefinedNames.Add("SpanTotal","=SUM('Rates:Fees'!$B$1:$L$2)");
  w.DefinedNames.Add("LeftTotal","=SUM('Rates:Fees'!$A$1)");
  string[] sheets={"Rates","Fees"};object saved=Capture(w,sheets,at,1);
  foreach(string s in sheets)w.Worksheets[s].Columns.Insert(at,1);Apply(saved);
  foreach(string s in sheets){var ws=w.Worksheets[s];int source=offset<0?at-1:at+1;ws.Range.FromLTRB(at,0,at,5).CopyFrom(ws.Range.FromLTRB(source,0,source,5),PasteSpecial.All);
   repair.GetMethod("CopyColumn").Invoke(null,new object[]{ws,source,at,offset<0?1:2,0,5});}
  saved=reuse?repair.GetMethod("CaptureFollowingInsertion").Invoke(saved,new object[]{at+1,count,sheets.Select(s=>w.Worksheets[s].Range.FromLTRB(at,0,at+1,5)).ToArray()}):Capture(w,sheets,at+1,count);
  foreach(string s in sheets)w.Worksheets[s].Columns.Insert(at+1,count);Apply(saved);
  foreach(string s in sheets){var ws=w.Worksheets[s];int source=offset<0?at:at+1+count;ws.Range.FromLTRB(at+1,0,at+count,5).CopyFrom(ws.Range.FromLTRB(source,0,source,5),PasteSpecial.All);
   repair.GetMethod("CopyColumn").Invoke(null,new object[]{ws,source,at+1,count+(offset<0?0:1),0,5});}
  repair.GetMethod("Verify").Invoke(saved,null);
  return w;
 }
 static void SplitTests(){
  for(int offset=-1;offset<=0;offset++)for(int at=1;at<=5;at++)for(int count=1;count<=3;count++)using(var full=Split(false,at,count,offset))using(var fast=Split(true,at,count,offset)){
   foreach(var ws in full.Worksheets)for(int row=0;row<=5;row++)for(int col=0;col<12+count+1;col++){
    var a=ws.Cells[row,col];var b=fast.Worksheets[ws.Name].Cells[row,col];
    if(a.FormulaInvariant!=b.FormulaInvariant||a.Value!=b.Value||a.HasArrayFormula!=b.HasArrayFormula)throw new Exception("Split inventory mismatch "+ws.Name+"!"+a.GetReferenceA1());
   }
   foreach(var n in full.DefinedNames)Equal(fast.DefinedNames.GetDefinedName(n.Name).RefersTo,n.RefersTo);
  }
  using(var w=new Workbook()){
   w.Worksheets[0].Name="Rates";w.Worksheets.Add("Fees");var s=w.Worksheets.Add("Summary");s.Cells[0,0].FormulaInvariant="=SUM('Rates:Fees'!B1)";
   var saved=Capture(w,new[]{"Rates","Fees"},1,1);bool rejected=false;
   try{repair.GetMethod("CaptureFollowingInsertion").Invoke(saved,new object[]{2,1,new CellRange[0]});}catch(TargetInvocationException e){rejected=e.InnerException is InvalidOperationException;}
   if(!rejected)throw new Exception("Unapplied capture reuse allowed");
   w.Worksheets[0].Columns.Insert(1,1);w.Worksheets[1].Columns.Insert(1,1);Apply(saved);
   rejected=false;try{repair.GetMethod("CaptureFollowingInsertion").Invoke(saved,new object[]{5,1,new CellRange[0]});}catch(TargetInvocationException e){rejected=e.InnerException is InvalidOperationException;}
   if(!rejected)throw new Exception("Non-adjacent capture reuse allowed");
   s.Cells[0,0].FormulaInvariant="=SUM('Rates:Fees'!Z1)";rejected=false;
   try{repair.GetMethod("Verify").Invoke(saved,null);}catch(TargetInvocationException e){rejected=e.InnerException is InvalidOperationException;}
   if(!rejected)throw new Exception("Incorrect reference passed final verification");
   w.Worksheets[0].Name="Renamed";rejected=false;
   try{repair.GetMethod("CaptureFollowingInsertion").Invoke(saved,new object[]{2,1,new CellRange[0]});}catch(TargetInvocationException e){rejected=e.InnerException is InvalidOperationException;}
   if(!rejected)throw new Exception("Changed sheet identity allowed");
  }
  Console.WriteLine("PASS: 30 full-scan versus inventory split-insertion cases; copies, arrays, names, unaffected refs, guards and final verification.");
 }
 [STAThread] public static int Main(string[] args){try{
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  repair=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe")).GetType("Abovo.WorkbookStructural3DReferences");
  var screen=repair.GetMethod("MayContainThreeDReference",BindingFlags.Static|BindingFlags.NonPublic);
  foreach(string formula in new[]{"=SUM(Rates:Fees!A1)","=SUM('First Sheet:Last Sheet'!A1)","=SUM('O''Brien:Last'!A1)","=SUM('First:O''Brien'!A1)","=SUM('First!Part:Last'!A1)","=SUM('[Book.xlsb]First:Last'!A1)","=SUM(前期:後期!A1)","=SUM(Fees:Rates!A1)"})
   if(!(bool)screen.Invoke(null,new object[]{formula}))throw new Exception("Missed sheet span: "+formula);
  foreach(string formula in new[]{"=SUM(A1:B2)","=SUM('One Sheet'!A1:B2)+Other!C3","=INDEX(One!A1:B3,1,1)","=A1",""})
   if((bool)screen.Invoke(null,new object[]{formula}))throw new Exception("Ordinary reference classified as a span: "+formula);
  using(var w=new Workbook()){
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;
   w.Worksheets[0].Name="Rates";w.Worksheets.Add("Fees");var summary=w.Worksheets.Add("Summary");
   var rates=w.Worksheets[0];
   string formula="=SUM('Rates:Fees'!$B1:C$2)+SUM('Rates:Fees'!1:2)+A1";
   summary.Cells["A5"].FormulaInvariant=formula;
   rates.Cells["B3"].FormulaInvariant="=SUM('Rates:Fees'!B1)+SUM('Rates:Fees'!$B$2)+SUM('Rates:Fees'!1:2)";
   rates.Cells["B4"].ArrayFormulaInvariant="=SUM('Rates:Fees'!B1)";
   w.DefinedNames.Add("SpanTotal","=SUM('Rates:Fees'!$B$1:$C$2)");
   var saved=Capture(w,new[]{"Rates","Fees"},1,2);
   rates.Columns.Insert(1,2);w.Worksheets[1].Columns.Insert(1,2);Apply(saved);
   Equal(summary.Cells["A5"].FormulaInvariant,"=SUM('Rates:Fees'!$D1:E$2)+SUM('Rates:Fees'!1:2)+A1");
   Equal(w.DefinedNames.GetDefinedName("SpanTotal").RefersTo,"=SUM('Rates:Fees'!$D$1:$E$2)");
   if(!rates.Cells["D4"].HasArrayFormula)throw new Exception("Array type lost during insert");
   Equal(rates.Cells["D4"].FormulaInvariant,"=SUM('Rates:Fees'!D1)");
   rates.Range["B4:C4"].CopyFrom(rates.Range["D4"],PasteSpecial.All);
   repair.GetMethod("CopyColumn").Invoke(null,new object[]{rates,3,1,2,3,3});
   if(!rates.Cells["B4"].HasArrayFormula||!rates.Cells["C4"].HasArrayFormula)throw new Exception("Copied array type lost");
   rates.Range["B3:C3"].CopyFrom(rates.Range["D3"],PasteSpecial.All);
   repair.GetMethod("CopyColumn").Invoke(null,new object[]{rates,3,1,2,2,2});
   Equal(rates.Cells["B3"].FormulaInvariant,"=SUM('Rates:Fees'!B1)+SUM('Rates:Fees'!$D$2)+SUM('Rates:Fees'!1:2)");
   Equal(rates.Cells["C3"].FormulaInvariant,"=SUM('Rates:Fees'!C1)+SUM('Rates:Fees'!$D$2)+SUM('Rates:Fees'!1:2)");
   saved=Capture(w,new[]{"Rates","Fees"},1,-2);rates.Columns.Remove(1,2);w.Worksheets[1].Columns.Remove(1,2);Apply(saved);
   Equal(summary.Cells["A5"].FormulaInvariant,formula);
   Equal(w.DefinedNames.GetDefinedName("SpanTotal").RefersTo,"=SUM('Rates:Fees'!$B$1:$C$2)");
   if(!rates.Cells["B4"].HasArrayFormula)throw new Exception("Array type lost during delete");
   Equal(rates.Cells["B4"].FormulaInvariant,"=SUM('Rates:Fees'!B1)");
   bool rejected=false;try{Capture(w,new[]{"Rates"},1,2);}catch(TargetInvocationException e){rejected=e.InnerException is InvalidOperationException;}
   if(!rejected)throw new Exception("Partial worksheet span was not rejected");
   Equal(summary.Cells["A5"].FormulaInvariant,formula);
   // A new capture must see changed worksheet names/order and formula edits.
   rates.Name="O'Brien";w.Worksheets[1].Name="Last!Sheet";
   summary.Cells["A6"].FormulaInvariant="=SUM('O''Brien:Last!Sheet'!B1)";
   saved=Capture(w,new[]{"O'Brien","Last!Sheet"},1,1);
   rates.Columns.Insert(1,1);w.Worksheets[1].Columns.Insert(1,1);Apply(saved);
   Equal(summary.Cells["A6"].FormulaInvariant,"=SUM('O''Brien:Last!Sheet'!C1)");
   Console.WriteLine("PASS: 3-D insert/delete, local refs, mixed absolute refs, range, whole rows, defined name, repeated copies, partial-span preflight.");
  }SplitTests();return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
