// Read-only diagnostic: run with Test-ClientReport.ps1 -Fixture WorkbookLoadProbe.cs.
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using DevExpress.Spreadsheet;

public static class WorkbookLoadProbe {
 [STAThread] public static int Main(string[] args) {
  try {
   AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
   using(var workbook=new Workbook()) {
    workbook.Options.CalculationMode=WorkbookCalculationMode.Manual;
    workbook.Options.Import.ThrowExceptionOnInvalidDocument=true;
    using(var input=File.OpenRead(args[2])) {
     if(!workbook.LoadDocument(input,DocumentFormat.Xlsb))throw new Exception("Read-only load returned false");
    }
    Console.WriteLine("PASS Read-only XLSB load; sheets="+workbook.Worksheets.Count+"; privateBytes="+Process.GetCurrentProcess().PrivateMemorySize64);
   }
   return 0;
  } catch(Exception e) {Console.Error.WriteLine(e);return 1;}
 }
}
