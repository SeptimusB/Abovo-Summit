using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Functions;

public static class AglCalculationProbe {
 [STAThread] public static int Main(string[] args) {try {
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  Run(args);return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
 static void Run(string[] args) {
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  using(var w=new Workbook()) {
   foreach(string type in new[]{"PMCostFunction","ResponsiveCostFunction"}) {var fn=(ICustomFunction)Activator.CreateInstance(app.GetType("Abovo."+type));if(!w.Functions.GlobalCustomFunctions.Contains(fn.Name))w.Functions.GlobalCustomFunctions.Add(fn);}
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;
   using(var stream=File.OpenRead(args[1]))w.LoadDocument(stream,DocumentFormat.Xlsb);
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;
   Console.WriteLine("FILE="+args[1]); Dump(w,"saved");
   w.Options.CalculationEngineType=CalculationEngineType.Recursive;
   var timer=Stopwatch.StartNew();Console.WriteLine("START full recursive rebuild");
   w.CalculateFullRebuild();Console.WriteLine("REBUILD MS="+timer.ElapsedMilliseconds);Dump(w,"recursive");
   using(var writer=new StreamWriter(args[2])) {
    foreach(string sheet in new[]{"Detailed Comp Inc - Trad View","Financial Position - Trad View","Cashflow detailed","Check Sheet"})
     foreach(var cell in w.Worksheets[sheet].GetUsedRange().ExistingCells)
      if(cell.HasFormula)writer.WriteLine(sheet+"\t"+cell.GetReferenceA1()+"\t"+cell.Value.ToString(CultureInfo.InvariantCulture));
   }
  }
 }
 static void Dump(IWorkbook w,string stage) {
  foreach(var item in new[]{new[]{"Cashflow detailed","BO9"},new[]{"Cashflow detailed","BR9"},new[]{"Detailed Comp Inc - Trad View","J73"},new[]{"Financial Position - Trad View","K26"}}) {
   var c=w.Worksheets[item[0]].Cells[item[1]];Console.WriteLine(stage+" "+item[0]+"!"+item[1]+" = "+c.Value+" formula="+c.FormulaInvariant);
  }
  foreach(string name in new[]{"Loan Interest Paid","Loan Drawdowns","Loan Repayments"}) {
   var ws=w.Worksheets[name];Console.WriteLine(stage+" "+name+" row 8:");
   for(int col=4;col<43;col++){var c=ws.Cells[7,col];if(!c.Value.IsEmpty)Console.WriteLine(c.GetReferenceA1()+"="+c.Value+" "+c.FormulaInvariant);}
  }
 }
}
