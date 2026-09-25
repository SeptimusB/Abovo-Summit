using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;

static class NativeSpillProbe
{
    static void Dump(Workbook book,string label){var sheet=book.Worksheets[0];var cell=sheet.Cells["D1"];Console.WriteLine("SPILL "+label+" engine="+book.Options.CalculationEngineType+" dynamic="+cell.HasDynamicArrayFormula+" range="+cell.GetDynamicArrayFormulaRange()?.GetReferenceA1()+" legacy="+cell.HasArrayFormula+" legacyRange="+cell.GetArrayFormulaRange()?.GetReferenceA1()+" values="+String.Join(",",new[]{"A1","D1","D2","D3"}.Select(a=>sheet.Cells[a].Value.ToString()))+" formula="+cell.FormulaInvariant);}
    internal static void Run(string original){
        string root=Path.Combine(Path.GetDirectoryName(original),"spill-probe-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);
        foreach(var extension in new[]{".xlsx",".xlsb"}){
            string path=Path.Combine(root,"probe"+extension);
            using(var source=new Workbook()){source.Worksheets[0].Cells["A1"].Value=10d;source.Worksheets[0].Cells["A1"].Protection.Locked=false;source.Worksheets[0].Cells["D1"].DynamicArrayFormulaInvariant="=A1+{0;1;2}";source.CalculateFullRebuild();Dump(source,"created"+extension);source.SaveDocument(path,extension==".xlsb"?DocumentFormat.Xlsb:DocumentFormat.Xlsx);}
            using(var book=new Workbook()){
                book.Options.CalculationMode=WorkbookCalculationMode.Manual;book.Options.CalculationEngineType=CalculationEngineType.Recursive;
                book.LoadDocument(path);book.Options.CalculationMode=WorkbookCalculationMode.Manual;Dump(book,"loaded"+extension);
                book.Worksheets[0].Cells["A1"].Value=11d;book.CalculateFullRebuild();Dump(book,"rebuilt"+extension);
                book.Options.CalculationEngineType=CalculationEngineType.ChainBased;book.CalculateFullRebuild();Dump(book,"chain"+extension);
                string converted=Path.Combine(root,"converted"+extension+".xlsm");book.SaveDocument(converted,DocumentFormat.Xlsm);
                var helper=typeof(Abovo.ModelChangeManagerV2).Assembly.GetType("Abovo.RecoveryXlsmCompatibility");var flags=BindingFlags.Static|BindingFlags.NonPublic;
                helper.GetMethod("Prepare",flags).Invoke(null,new object[]{converted,(bool)helper.GetMethod("HasVerifiedBinaryProfile",flags).Invoke(null,new object[]{path})});
                using(var reloaded=new Workbook()){reloaded.Options.CalculationMode=WorkbookCalculationMode.Manual;reloaded.LoadDocument(converted);Dump(reloaded,"converted"+extension);reloaded.CalculateFullRebuild();Dump(reloaded,"converted-rebuilt"+extension);}
            }
            var session=WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(WorkbookEnginePreference.ExcelRequired,false,false,120000,true)).GetAwaiter().GetResult();
            try{
                var areas=new[]{new WorkbookReadArea("Sheet1",0,0,3,4)};
                var initial=session.CalculateAndReadAsync(0,WorkbookCalculationKind.Rebuild,areas).GetAwaiter().GetResult();
                var input=session.CaptureCellAsync(0,new WorkbookReadArea("Sheet1",0,0,1,1)).GetAwaiter().GetResult();
                var after=session.ApplyValueAsync(input,11d,WorkbookValuePermission.UnlockedCell,areas).GetAwaiter().GetResult();
                Console.WriteLine("SPILL EXCEL "+extension+" before="+initial.Blocks[0].ValueAt(2,3)+" after="+after.Results.Blocks[0].ValueAt(2,3));
            }finally{session.CloseAsync().GetAwaiter().GetResult();}
            NativeProcessChecks.RequireOwnedProcesses(new[]{session}).GetAwaiter().GetResult();
        }
        string imported=Path.Combine(root,"excel-source.xlsb");
        var converter=NativeFixtureConversion.ExcelXlsb(Path.Combine(root,"probe.xlsx"),imported).GetAwaiter().GetResult();
        NativeProcessChecks.RequireOwnedIdentities(new[]{converter}).GetAwaiter().GetResult();
        foreach(string variant in new[]{"imported-direct","promoted-dynamic"}){
            string path=Path.Combine(root,variant+".xlsb");
            using(var book=new Workbook()){
                book.Options.CalculationMode=WorkbookCalculationMode.Manual;book.LoadDocument(imported);Dump(book,variant+"-initial");
                if(variant=="promoted-dynamic"){
                    var sheet=book.Worksheets[0];var range=sheet.Cells["D1"].GetArrayFormulaRange();string formula=sheet.Cells["D1"].FormulaInvariant;
                    range.ArrayFormulaInvariant=String.Empty;sheet.Cells["D1"].DynamicArrayFormulaInvariant=formula;
                    book.CalculateFullRebuild();Dump(book,variant+"-before-export");
                }
                book.SaveDocument(path,DocumentFormat.Xlsb);
            }
            using(var verify=new Workbook()){verify.Options.CalculationMode=WorkbookCalculationMode.Manual;verify.LoadDocument(path);Dump(verify,variant+"-reloaded");verify.Worksheets[0].Cells["A1"].Value=11d;verify.CalculateFullRebuild();Dump(verify,variant+"-recalculated");}
            var session=WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(WorkbookEnginePreference.ExcelRequired,false,false,120000,true)).GetAwaiter().GetResult();
            try{
                var areas=new[]{new WorkbookReadArea("Sheet1",0,0,3,4)};var initial=session.CalculateAndReadAsync(0,WorkbookCalculationKind.Rebuild,areas).GetAwaiter().GetResult();
                var input=session.CaptureCellAsync(0,new WorkbookReadArea("Sheet1",0,0,1,1)).GetAwaiter().GetResult();
                var after=session.ApplyValueAsync(input,11d,WorkbookValuePermission.UnlockedCell,areas).GetAwaiter().GetResult();
                Console.WriteLine("SPILL EXCEL "+variant+" before="+initial.Blocks[0].ValueAt(2,3)+" after="+after.Results.Blocks[0].ValueAt(2,3));
            }finally{session.CloseAsync().GetAwaiter().GetResult();}
            NativeProcessChecks.RequireOwnedProcesses(new[]{session}).GetAwaiter().GetResult();
        }
    }
}
