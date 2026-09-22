using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
public static class JournalFundingInspection {
    static void Dump(IWorkbook w) {
        foreach(var n in w.DefinedNames.Where(n=>n.Name.IndexOf("Jour",StringComparison.OrdinalIgnoreCase)>=0))
            Console.WriteLine("NAME "+n.Name+" = "+n.RefersTo);
        var check=w.Worksheets["Check Sheet"];
        foreach(var cell in check.GetUsedRange().ExistingCells.Where(c=>c.FormulaInvariant.IndexOf("Jour",StringComparison.OrdinalIgnoreCase)>=0 || c.DisplayText.IndexOf("Journal",StringComparison.OrdinalIgnoreCase)>=0))
            Console.WriteLine("CHECK "+cell.GetReferenceA1()+" = "+cell.FormulaInvariant+" => "+cell.DisplayText);
    }
    [STAThread] public static int Main(string[] args) {
        try {
            AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
            var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
            app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
            var manager=app.GetType("Abovo.FileManager");manager.GetMethod("Initialise").Invoke(null,new object[]{null});
            var copy=Path.Combine(args[2],"inspection-copy.xlsb");File.Copy(args[1],copy);
            var open=manager.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
            if(loaded.BError)throw new Exception(loaded.StringReturn);
            dynamic model=((Array)manager.GetField("ExcelModels").GetValue(null)).GetValue(0);
            IWorkbook w=(IWorkbook)model.WB;
            Dump(w);
            var fund=w.Worksheets["Funding Assumptions"];
            foreach(var c in fund.Range["E3:N5"].ExistingCells) Console.WriteLine("FUND "+c.GetReferenceA1()+" = "+c.DisplayText);
            foreach(var c in fund.Range["E137:F137"].ExistingCells) Console.WriteLine("DATE "+c.GetReferenceA1()+" value="+c.Value+" format="+c.NumberFormat+" locked="+c.Protection.Locked);
            foreach(var validation in fund.DataValidations.GetDataValidations(fund.Range["E137:R137"])) {
                Console.WriteLine("VALIDATION "+validation.GetType().FullName);
                var v=validation.Criteria;
                Console.WriteLine("CRITERIA text="+v.IsText+" range="+v.IsRange+" formula="+v.IsFormula+" formulaValue="+v.FormulaInvariant);
                if(v.IsRange)foreach(var c in v.RangeValue.ExistingCells)Console.WriteLine("CHOICE "+c.GetReferenceA1()+"="+c.Value+" format="+c.NumberFormat);
                if(v.IsFormula){var name=w.DefinedNames.GetDefinedName(v.FormulaInvariant.TrimStart('='));if(name!=null)foreach(var c in name.Range.ExistingCells)Console.WriteLine("CHOICE "+c.GetReferenceA1()+"="+c.Value+" format="+c.NumberFormat);}
                foreach(var p in validation.GetType().GetProperties().Where(p=>p.GetIndexParameters().Length==0))
                    try{Console.WriteLine("  "+p.Name+"="+p.GetValue(validation,null));}catch{}
            }
            if(args[3]=="True") {
                var names=new[]{"IR_Journals","Rep_Jour_01","TransCopy_IR_Journals_01","TransCopy_IR_Journals_02"};
                var original=names.ToDictionary(n=>n,n=>w.DefinedNames.GetDefinedName(n).Range.RowCount);
                var before=names.ToDictionary(n=>n,n=>w.DefinedNames.GetDefinedName(n).RefersTo);
                var protection=w.Worksheets.ToDictionary(s=>s.Name,s=>s.IsProtected);
                dynamic result=model.WorkbookStructureRules.AddRecords("JOURNAL_RECORDS",5);
                Console.WriteLine("INSERT error="+result.BError+" "+result.StringReturn);
                if(result.BError)throw new Exception(result.StringReturn);
                Dump(w);
                foreach(var n in names)if(w.DefinedNames.GetDefinedName(n).Range.RowCount!=original[n]+5)throw new Exception("Range not expanded: "+n);
                foreach(var n in names.Skip(2)) {
                    var range=w.DefinedNames.GetDefinedName(n).Range;
                    var cells=range.ExistingCells.Where(c=>c.RowIndex>=range.BottomRowIndex-5 && c.HasFormula).ToList();
                    if(cells.Count<200 || cells.Any(c=>c.FormulaInvariant.Contains("#REF!")))throw new Exception("Missing/broken mirror formulas: "+n);
                    Console.WriteLine("MIRROR "+n+" final six rows formula count="+cells.Count);
                }
                if(w.Worksheets.Any(s=>s.IsProtected!=protection[s.Name]) || !model.IsDirty)throw new Exception("Protection/dirty regression");
                var saved=Path.Combine(args[2],"journal-expanded.xlsb");w.SaveDocument(saved,DocumentFormat.Xlsb);
                using(var reopened=new Workbook()){
                    reopened.Options.CalculationMode=WorkbookCalculationMode.Manual;reopened.LoadDocument(saved,DocumentFormat.Xlsb);
                    foreach(var n in names)if(reopened.DefinedNames.GetDefinedName(n).Range.RowCount!=original[n]+5)throw new Exception("Saved range mismatch "+n);
                }
                result=model.WorkbookStructureRules.DeleteRecords("JOURNAL_RECORDS",Enumerable.Range(original["IR_Journals"],5));
                if(result.BError)throw new Exception(result.StringReturn);
                foreach(var n in names)if(w.DefinedNames.GetDefinedName(n).RefersTo!=before[n])throw new Exception("Delete failed to restore range "+n);
                Console.WriteLine("PASS: five-row add, both formula mirrors, dirty/protection, save/reopen and delete back to original geometry.");
                result=model.WorkbookStructureRules.AddRecords("JOURNAL_RECORDS",3);
                if(result.BError)throw new Exception(result.StringReturn);
                result=model.WorkbookStructureRules.DeleteRecords("JOURNAL_RECORDS",new[]{1,3});
                if(result.BError)throw new Exception(result.StringReturn);
                foreach(var n in names)if(w.DefinedNames.GetDefinedName(n).Range.RowCount!=original[n]+1)throw new Exception("Noncontiguous deletion range "+n);
                Console.WriteLine("PASS: noncontiguous deletion updates both source names and both mirrors.");
                var inputName=w.DefinedNames.GetDefinedName("IR_Journals");var inputRange=inputName.Range;
                inputName.Range=inputRange.Worksheet.Range.FromLTRB(inputRange.LeftColumnIndex,inputRange.TopRowIndex,inputRange.RightColumnIndex,inputRange.BottomRowIndex-1);
                var damaged=names.ToDictionary(n=>n,n=>w.DefinedNames.GetDefinedName(n).RefersTo);
                result=model.WorkbookStructureRules.AddRecords("JOURNAL_RECORDS",5);
                if(!result.BError || names.Any(n=>w.DefinedNames.GetDefinedName(n).RefersTo!=damaged[n]))throw new Exception("Pre-existing name mismatch was not rejected before mutation");
                Console.WriteLine("PASS: pre-existing inconsistent journal names fail before any structural mutation.");
            }
            Console.WriteLine("PASS: inspected private in-memory copy; source not saved.");return 0;
        }catch(Exception e){Console.Error.WriteLine(e);return 1;}
    }
}
