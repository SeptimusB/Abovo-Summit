using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Abovo;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;

static class ExpressionTests
{
    static int assertions;
    static void Check(bool ok,string message){if(!ok)throw new Exception(message);Console.WriteLine("EXPRESSION PASS "+(++assertions)+" "+message);}
    static async Task Reject(Func<Task> work,string label){try{await work();}catch(Exception e)when(e is ArgumentException||e is InvalidOperationException||e is NotSupportedException){Check(true,label);return;}throw new Exception("Expected rejection: "+label);}
    internal static async Task Run(string original)
    {
        var owners=new List<WorkbookCalculationSession>();
        var root=Path.Combine(Path.GetDirectoryName(original),"expression-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);
        var source=Path.Combine(root,"inputs.xlsx");
        using(var book=new Workbook())
        {
            var ws=book.Worksheets[0];ws.Name="Inputs";ws.Cells["A1"].Value=2;ws.Cells["A1"].Protection.Locked=false;
            ws.Cells["A2"].Value=7;ws.Cells["B1"].Formula="=A1*3";ws.Cells["D1"].Value=9;ws.Cells["D2"].Value=10;
            ws.DataValidations.Add(ws.Range["D1:D2"],DataValidationType.Decimal,DataValidationOperator.GreaterThan,"=A1");
            var other=book.Worksheets.Add("Other sheet");other.Cells["A1"].Value=17;
            book.DefinedNames.Add("NativeLimit","='Inputs'!$A$1");
            book.CalculateFull();book.SaveDocument(source,DocumentFormat.Xlsx);
        }
        foreach(var preference in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired})
        {
            var s=await WorkbookCalculationSession.OpenAsync(source,new WorkbookEngineOptions(preference,false,false,120000,true));
            owners.Add(s);
            using(var model=new EngineChangeManagerTests.Model(source))
            try
            {
                var area=new WorkbookReadArea("Inputs",0,0,1,1);
                var initial=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{area});
                Check(Equals(await s.EvaluateCurrentAsync(initial,area,"=SUM(A1,B1)"),8d),s.EngineName+" expression uses native precedent values");
                Check(Equals(await s.EvaluateCurrentAsync(initial,area,"=NativeLimit"),2d),"scalar named-range reference evaluated");
                Check(Equals(await s.EvaluateCurrentAsync(initial,area,"='Other sheet'!A1+1"),18d),"quoted cross-sheet reference evaluated");
                Check(Equals(await s.EvaluateCurrentAsync(initial,area,"=A1<B1"),true),"Boolean predicate remains Boolean");
                Check((await s.EvaluateCurrentAsync(initial,area,"=1/0")) is WorkbookCellError,"formula error is not silently zero");
                await Reject(()=>s.EvaluateCurrentAsync(initial,area,"=A1:A2"),"multi-cell result rejected without model writes");
                Check(s.IsCurrent(initial),"unsupported array leaves owner usable");
                await Reject(()=>s.EvaluateCurrentAsync(initial,area,"='[other.xlsx]Sheet1'!A1"),"external workbook reference refused");
                await Reject(()=>s.EvaluateCurrentAsync(initial,area,"="+new string('1',255)),"oversized expression refused before engine");
                await model.Do(()=>{
                    typeof(ModelChangeManagerV2).GetMethod("BindEngineEditingTrial",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(model.Manager,new object[]{s,initial});
                    var ws=model.Value.WB.Worksheets[0];
                    Check(ws.Cells["D2"].ModelEvaluate("=RC[-3]",ReferenceStyle.R1C1).NumericValue==7,"relative validation rebased at the actual target cell");
                    var numeric=typeof(ModelChangeManagerV2).Assembly.GetTypes().Single(t=>t.Name=="DataInterfaceTemplate").GetMethod("NumericInputError",BindingFlags.Static|BindingFlags.NonPublic);
                    var tag=new DataObject.DataColumnTag{DataType="N"};
                    Check(numeric.Invoke(null,new object[]{ws.Cells["D2"],tag,6d})!=null&&numeric.Invoke(null,new object[]{ws.Cells["D2"],tag,8d})==null,"real numeric input validator respects rebased native limit");
                    var response=model.Manager.ProcessEngineCommand(new[]{new DataChangeEvent{ModelID=0,WSName="Inputs",CellAddress="A1",ChangedValue=5d,DataFormat="N"}},new[]{WorkbookValuePermission.UnlockedCell},"Change limit");
                    Check(response.BSuccess,"precedent changed through existing manager");
                    Check(ws.Cells["D1"].ModelEvaluate("=SUM(A1,B1)").NumericValue==20,"validation expression uses changed engine values, not old presentation caches");
                    Check(numeric.Invoke(null,new object[]{ws.Cells["D1"],tag,4d})!=null&&numeric.Invoke(null,new object[]{ws.Cells["D1"],tag,6d})==null,"real input validator follows changed native precedent");
                    Check(ws.Cells["A1"].Value.NumericValue==2&&ws.Cells["B1"].HasFormula,"expression evaluation never rewrites presentation workbook");
                    return true;
                });
                await Reject(()=>s.EvaluateCurrentAsync(initial,area,"=A1"),"stale expression anchor refused");
            }
            finally{await s.CloseAsync();}
        }
        await NativeProcessChecks.RequireOwnedProcesses(owners);
        Console.WriteLine("EXPRESSION ASSERTIONS="+assertions);
    }
}
