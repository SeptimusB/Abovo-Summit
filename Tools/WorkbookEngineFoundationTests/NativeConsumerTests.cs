using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Abovo;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;

// Verify consumers against deliberately stale presentation caches, not just
// against another workbook which happens to contain the expected values.
static class NativeConsumerTests
{
    const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Static|BindingFlags.Instance;
    static int checks;
    static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("CONSUMER_NATIVE PASS "+(++checks)+" "+text);}
    static object Call(object owner,string name,params object[] args)=>owner.GetType().GetMethod(name,F).Invoke(owner,args);
    static string Create(string source,bool date1904)
    {
        string folder=Path.Combine(Path.GetDirectoryName(source),"native-consumers-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);string path=Path.Combine(folder,"private.xlsx");
        using(var wb=new Workbook()){
            wb.DocumentSettings.Calculation.Use1904DateSystem=date1904;
            var ws=wb.Worksheets[0];ws.Name="Data";ws.Cells["A1"].Value=10d;
            ws.Cells["B1"].Value=CellValue.FromDateTime(new DateTime(2026,9,30),date1904);ws.Cells["B1"].NumberFormat="dd-mmm-yyyy";
            ws.Range["A1:B1"].Protection.Locked=false;ws.Range["A1:B1"].Fill.PatternType=PatternType.Solid;
            ws.Cells["A2"].Formula="=1/A1";wb.DefinedNames.Add("IR_TestInputs","=Data!$A$1:$B$1");
            wb.CalculateFullRebuild();wb.SaveDocument(path,DocumentFormat.Xlsx);
        }
        return path;
    }
    static string Scan(EngineChangeManagerTests.Model model)
    {
        var type=typeof(ModelChangeManagerV2).Assembly.GetType("Abovo.IdleIntegrityManager+IntegrityPass");
        var pass=Activator.CreateInstance(type,F,null,new object[]{model.Value,true},null);
        try{
            var stage=type.GetField("Stage",F);stage.SetValue(pass,Enum.Parse(stage.FieldType,"Cells"));
            int steps=0;while(!(bool)type.GetProperty("Finished",F).GetValue(pass,null)){
                if(++steps>1000)throw new Exception("Integrity did not finish its bounded scan");Call(pass,"Advance",new object[]{null});
            }
            return (string)Call(pass,"Report");
        }finally{((IDisposable)pass).Dispose();}
    }
    internal static async Task Run(string source)
    {
        var owners=new List<WorkbookCalculationSession>();
        foreach(bool date1904 in new[]{false,true})foreach(var engine in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired}){
            string path=Create(source,date1904);byte[] original=File.ReadAllBytes(path);
            var session=await WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(engine,false,false,120000,true));owners.Add(session);
            using(var model=new EngineChangeManagerTests.Model(path))try{
                var first=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Rebuild,new[]{new WorkbookReadArea("Data",0,0,2,2)});
                await model.Do(()=>{
                    Call(model.Manager,"BindEngineEditingTrial",session,first);
                    var changed=model.Manager.ProcessEngineCommand(new[]{new DataChangeEvent{ModelID=0,WSName="Data",CellAddress="A1",DataFormat="N",ChangedValue=0d,Description="Consumer error test"}},new[]{WorkbookValuePermission.UnlockedCell},"Consumer error test");
                    Check(changed.BSuccess,session.EngineName+" native edit produces a current formula error");
                    Check(!model.Value.WB.Worksheets[0].Cells["A2"].Value.IsError,"presentation formula cache intentionally remains old");
                    Check(Scan(model).Contains("Cell error #DIV/0!: Data!A2"),"integrity scans native errors instead of stale caches");
                    using(var baseline=new Workbook()){
                        Check(baseline.LoadDocument(path),"comparison baseline opens");
                        var left=new BusinessPlanComparisonParticipant{Key="base",DisplayName="Base",FilePath=path,Workbook=baseline};
                        var right=new BusinessPlanComparisonParticipant{Key="current",DisplayName="Current",FilePath=path,Workbook=model.Value.WB,ModelID=0};
                        var comparison=BusinessPlanComparisonService.Compare(left,new[]{right},new HashSet<string>{"Data"});
                        Check(comparison.DifferenceCount==1&&comparison.Items.Any(i=>i.Address=="Data!A1"&&i.ComparedValue=="0"),"comparison reads unsaved native inputs");
                    }
                    object[] typed={model.Value.WB.Worksheets[0].Cells["B1"],"D",null,null};
                    var error=typeof(BusinessPlanComparisonService).GetMethod("GetTypedCellValue",F).Invoke(null,typed);
                    Check(error==null&&(DateTime)typed[2]==new DateTime(2026,9,30),"typed comparison respects date system "+(date1904?1904:1900));
                    Check(model.Manager.Undo().BSuccess&&!Scan(model).Contains("Cell error #DIV/0!"),"Undo clears current native error from integrity report");
                    return true;
                });
            }finally{await session.CloseAsync();}
            Check(File.ReadAllBytes(path).SequenceEqual(original),"original generated workbook unchanged");
        }
        await NativeProcessChecks.RequireOwnedProcesses(owners);Console.WriteLine("CONSUMER_NATIVE ASSERTIONS="+checks);
    }
}
