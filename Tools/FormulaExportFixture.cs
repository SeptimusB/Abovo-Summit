using System;
using System.IO;
using System.Linq;
using System.Reflection;
using DevExpress.Spreadsheet;

public static class FormulaExportFixture {
 [STAThread] public static int Main(string[] args) {
  try {
   AppDomain.CurrentDomain.AssemblyResolve += (s,e) => {
    string path=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");
    return File.Exists(path)?Assembly.LoadFrom(path):null;
   };
   Run(args[1]); return 0;
  } catch(Exception e) {Console.Error.WriteLine(e);return 1;}
 }
 static void Run(string output) {
  using(var w=new Workbook()) {
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;
   var sheet=w.Worksheets[0]; int row=0;
   foreach(int count in new[]{2,29,30,31,37,60,127,128,255}) {
    foreach(string fn in new[]{"CONCATENATE","SUM"}) {
     sheet.Cells[row,0].Value=fn+" "+count;
     sheet.Cells[row++,1].FormulaInvariant="="+fn+"("+String.Join(",",Enumerable.Repeat("1",count))+")";
    }
   }
   sheet.Cells[row,0].Value="nested 37";
   sheet.Cells[row++,1].FormulaInvariant="=CONCATENATE(CONCATENATE("+String.Join(",",Enumerable.Repeat("1",30))+"),"+String.Join(",",Enumerable.Repeat("1",7))+")";
   w.Calculate();
   foreach(DocumentFormat format in new[]{DocumentFormat.Xlsb,DocumentFormat.Xlsx}) {
    string path=Path.Combine(output,"formula-export."+(format==DocumentFormat.Xlsb?"xlsb":"xlsx"));
    w.SaveDocument(path,format);
    using(var read=new Workbook()) {
     read.Options.CalculationMode=WorkbookCalculationMode.Manual;
     read.LoadDocument(path);
     for(int r=0;r<row;r++) Console.WriteLine(format+" "+sheet.Cells[r,0].Value+" : "+(read.Worksheets[0].Cells[r,1].FormulaInvariant==sheet.Cells[r,1].FormulaInvariant?"SAME":"CHANGED "+read.Worksheets[0].Cells[r,1].FormulaInvariant)+" : "+read.Worksheets[0].Cells[r,1].Value);
    }
   }
  }
 }
}
