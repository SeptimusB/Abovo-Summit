using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Functions;

public static class PMCostFixture {
 static Assembly app;
 static int assertions;
 static void Check(bool ok,string message){if(!ok)throw new Exception(message);assertions++;}
 static ICustomFunction Register(IWorkbook workbook,string type="PMCostFunction"){
  var fn=(ICustomFunction)Activator.CreateInstance(app.GetType("Abovo."+type,true));
  if(!workbook.Functions.GlobalCustomFunctions.Contains(fn.Name))workbook.Functions.GlobalCustomFunctions.Add(fn);
  return fn;
 }
 [STAThread]public static int Main(string[] args){try{
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string path=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(path)?Assembly.LoadFrom(path):null;};
  Run(args);Console.WriteLine("PASS: "+assertions+" assertions");return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
 static void Run(string[] args){
  app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  Synthetic(args[1]);if(args.Length>2)Client(args[2]);
 }
 static void Values(IWorkbook workbook,List<double> expected){
  var sheet=workbook.Worksheets["Tests"];
  for(int i=0;i<expected.Count;i++){
   var cell=sheet.Cells[i,0];
   Check(cell.HasFormula,"Formula retained at "+cell.GetReferenceA1());
   Check(cell.Value.IsNumeric,"Numeric result at "+cell.GetReferenceA1()+": "+cell.Value);
   Check(Math.Abs(cell.Value.NumericValue-expected[i])<1e-8,"Result at "+cell.GetReferenceA1()+": expected "+expected[i]+", got "+cell.Value);
  }
 }
 static void Synthetic(string output){
  using(var w=new Workbook()){
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;
   var fn=Register(w);Register(w,"ResponsiveCostFunction");
   Check(fn.Name=="PMCost","Invariant name matches VBA PMCost");
   Check(fn.GetName(CultureInfo.GetCultureInfo("en-GB"))=="PMCost","Localized name matches VBA");
   Check(fn.GetName(CultureInfo.InvariantCulture)=="PMCost","Invariant display name matches VBA");
   Check(fn.Parameters.Length==8,"Eight compatibility positions");
   for(int i=0;i<8;i++){
    Check(fn.Parameters[i].Type==(i<6?ParameterType.Value:ParameterType.Reference),"Parameter type "+(i+1));
    Check(fn.Parameters[i].Attributes==DevExpress.Spreadsheet.Functions.ParameterAttributes.Required,"Required parameter "+(i+1));
   }
   var sheet=w.Worksheets[0];sheet.Name="Tests";
   var rates=w.Worksheets.Add("Rates");
   for(int i=0;i<6;i++){rates.Cells[i,0].Value=0.1*Math.Pow(2,i);rates.Cells[i,1].Value=i+1;}
   // Explicit expected results are independent of the implementation's MATCH loop.
   var expected=new List<double>();int row=0;
   Action<string,double> add=(formula,value)=>{sheet.Cells[row++,0].FormulaInvariant=formula;expected.Add(value);};
   foreach(string name in new[]{"PMCost","PMCOST","pmcost"}){
    foreach(int unused in new[]{-99,0,9999}){
     add("="+name+"(3,30,10,1,3,"+unused+",Rates!$A$1:$A$6,Rates!$B$1:$B$6)",70);
    }
   }
   add("=PMCost(0,30,10,1,3,1,Rates!$A$1:$A$6,Rates!$B$1:$B$6)",0);
   add("=PMCost(1,30,10,1,3,1,Rates!$A$1:$A$6,Rates!$B$1:$B$6)",10);
   add("=PMCost(2,30,10,1,3,1,Rates!$A$1:$A$6,Rates!$B$1:$B$6)",30);
   add("=PMCost(4,30,10,1,3,1,Rates!$A$1:$A$6,Rates!$B$1:$B$6)",140);
   add("=PMCost(6,30,10,1,3,1,Rates!$A$1:$A$6,Rates!$B$1:$B$6)",560);
   add("=PMCost(3,0,10,1,3,1,Rates!$A$1:$A$6,Rates!$B$1:$B$6)",0);
   add("=PMCost(3,30,0,1,3,1,Rates!$A$1:$A$6,Rates!$B$1:$B$6)",0);
   add("=PMCost(4,30,10,2,2,1,Rates!$A$1:$A$6,Rates!$B$1:$B$6)",120);
   add("=IF(TRUE,PMCost(3,30,10,1,3,123,INDEX(Rates!$A$1:$A$6,0,1),Rates!$B$1:$B$6),0)",70);
   rates.Cells["D1"].Value=0.1;rates.Cells["D2"].Value=0.5;rates.Cells["D3"].Value=1;
   rates.Cells["E1"].Value=1;rates.Cells["E2"].Value=3;rates.Cells["E3"].Value=5;
   add("=PMCost(4,20,10,1,2,0,Rates!$D$1:$D$3,Rates!$E$1:$E$3)",100);
   add("=RespCost(3,30,1,3,999,Rates!$A$1:$A$6,Rates!$B$1:$B$6)",7);
   add("=RespCost(4,20,1,2,0,Rates!$D$1:$D$3,Rates!$E$1:$E$3)",10);
   foreach(string formula in new[]{"=PMCost(3,30,10,1,3,Rates!A1:A6,Rates!B1:B6)","=PMCost(3,30,10,1,3,0,0,Rates!A1:A6,Rates!B1:B6)"}){
    bool rejected=false;try{sheet.Cells["C1"].FormulaInvariant=formula;}catch(ArgumentException){rejected=true;}
    Check(rejected,"Malformed argument count rejected");
   }
   var prior=w.Options.CalculationEngineType;
   try{w.Options.CalculationEngineType=CalculationEngineType.Recursive;w.CalculateFullRebuild();Values(w,expected);}
   finally{w.Options.CalculationEngineType=prior;}
   Console.WriteLine("PASS: casing, signature, unused FinalYear, nested INDEX and numeric examples");
   foreach(var format in new[]{DocumentFormat.Xlsb,DocumentFormat.Xlsm}){
    string path=Path.Combine(output,"pmcost-roundtrip."+(format==DocumentFormat.Xlsb?"xlsb":"xlsm"));
    w.SaveDocument(path,format);
    using(var reload=new Workbook()){
     reload.Options.CalculationMode=WorkbookCalculationMode.Manual;Register(reload);Register(reload,"ResponsiveCostFunction");
     Check(reload.LoadDocument(path,format),"Synthetic "+format+" opens");
     for(int i=0;i<row;i++)Check(String.Equals(sheet.Cells[i,0].FormulaInvariant,reload.Worksheets["Tests"].Cells[i,0].FormulaInvariant,StringComparison.OrdinalIgnoreCase),"Formula roundtrip "+format+" row "+i);
     prior=reload.Options.CalculationEngineType;
     try{reload.Options.CalculationEngineType=CalculationEngineType.Recursive;reload.CalculateFullRebuild();Values(reload,expected);}
     finally{reload.Options.CalculationEngineType=prior;}
    }
    Console.WriteLine("PASS: synthetic "+format+" save/reload/recalculation");
   }
   ErrorAndDependencyChecks(w);
  }
 }
 static void ErrorAndDependencyChecks(IWorkbook w){
  var s=w.Worksheets.Add("Guards");var rates=w.Worksheets["Rates"];
  rates.Cells["G1"].Value=2;rates.Cells["G2"].Value=4;
  foreach(bool pm in new[]{true,false}){
   string name=pm?"PMCost":"RespCost",prefix=pm?"1,30,10,1,3,0,":"1,30,1,3,0,";
   s.Cells["A1"].FormulaInvariant="="+name+"("+prefix+"Rates!A1:A2,Rates!G1:G2)";
   w.CalculateFullRebuild();Check(s.Cells["A1"].Value.ToString()=="#N/A",name+" missing lookup returns #N/A");
   s.Cells["A1"].FormulaInvariant="="+name+"("+(pm?"3,30,10,1,3,0,":"3,30,1,3,0,")+"Rates!A1:A1,Rates!B1:B6)";
   w.CalculateFullRebuild();Check(s.Cells["A1"].Value.ToString()=="#REF!",name+" short rates return #REF!");
   s.Cells["A1"].FormulaInvariant="="+name+"("+(pm?"NA(),30,10,1,3,0,":"NA(),30,1,3,0,")+"Rates!A1:A6,Rates!B1:B6)";
   w.CalculateFullRebuild();Check(s.Cells["A1"].Value.ToString()=="#N/A",name+" scalar error propagated");
   s.Cells["A1"].FormulaInvariant="="+name+"("+(pm?"3,30,10,1,3,0,":"3,30,1,3,0,")+"Rates!A1:A6,Rates!B1:B6)";
   w.CalculateFullRebuild();
   rates.Cells["A3"].Value=0.8;w.Calculate();
   Check(Math.Abs(s.Cells["A1"].Value.NumericValue-(pm?110:11))<1e-8,name+" rate edit observed without stale cache");
   rates.Cells["A3"].Value=0.4;w.Calculate();
   Check(Math.Abs(s.Cells["A1"].Value.NumericValue-(pm?70:7))<1e-8,name+" restored rate observed");
  }
  s.Cells["A1"].FormulaInvariant="=RESPCOST(1)";w.CalculateFullRebuild();
  Check(s.Cells["A1"].Value.IsError,"Missing optional RespCost references yield worksheet error");
 }
 static void Client(string path){
  using(var w=new Workbook()){
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;Register(w);Register(w,"ResponsiveCostFunction");
   var format=String.Equals(Path.GetExtension(path),".xlsm",StringComparison.OrdinalIgnoreCase)?DocumentFormat.Xlsm:DocumentFormat.Xlsb;
   using(var stream=File.OpenRead(path))Check(w.LoadDocument(stream,format),"Client evidence opens");
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;
   var calls=new List<Cell>();var formulas=new List<string>();
   foreach(var s in w.Worksheets)foreach(var c in s.GetUsedRange().ExistingCells)if(c.HasFormula&&Regex.IsMatch(c.FormulaInvariant,@"\bPMCost\s*\(",RegexOptions.IgnoreCase)){calls.Add(c);formulas.Add(c.FormulaInvariant);}
   Check(calls.Count>0,"Client evidence contains PMCost formulas");
   var engine=w.Options.CalculationEngineType;var clock=Stopwatch.StartNew();
   try{w.Options.CalculationEngineType=CalculationEngineType.Recursive;w.CalculateFullRebuild();}
   finally{w.Options.CalculationEngineType=engine;}
   int errors=0,numeric=0;for(int i=0;i<calls.Count;i++){
    Check(calls[i].FormulaInvariant==formulas[i],"Client PMCost formula retained");
    if(calls[i].Value.IsError){errors++;if(errors<6)Console.WriteLine("CLIENT PMCost ERROR: "+calls[i].Worksheet.Name+"!"+calls[i].GetReferenceA1()+"="+calls[i].Value);}
    if(calls[i].Value.IsNumeric)numeric++;
   }
   int nameErrors=0;var counts=new Dictionary<string,int>();
   foreach(var s in w.Worksheets)foreach(var c in s.GetUsedRange().ExistingCells)if(c.Value.IsError){string key=c.Value.ToString();if(!counts.ContainsKey(key))counts[key]=0;counts[key]++;if(key=="#NAME?")nameErrors++;}
   Console.WriteLine("CLIENT: PMCost calls="+calls.Count+", numeric="+numeric+", errors="+errors+", workbook NAME errors="+nameErrors+", rebuild ms="+clock.ElapsedMilliseconds);
   Console.WriteLine("CLIENT remaining errors: "+String.Join(", ",counts.Select(p=>p.Key+"="+p.Value)));
   Check(errors==0,"Client PMCost calls calculate without errors");
   Check(nameErrors==0,"No client NAME errors after rebuild");
   foreach(string a in new[]{"E23","E33","E39"})Console.WriteLine("CLIENT Check Sheet!"+a+"="+w.Worksheets["Check Sheet"].Cells[a].Value);
   Console.WriteLine("PASS: client file calculated in memory only; not saved");
  }
 }
}
