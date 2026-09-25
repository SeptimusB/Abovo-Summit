using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Abovo;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;

static class NativeStructuralGateTests
{
    static int checks;
    static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("STRUCTURE_GATE PASS "+(++checks)+" "+text);}
    static void Refused(AbovoAppCls.AbovoTransaction result,string text){Check(result.BError&&result.EventCancelled&&!result.BSuccess&&result.StrResponseMessage.Contains("calculation-engine mode"),text);}
    internal static async Task Run(string original)
    {
        string folder=Path.Combine(Path.GetDirectoryName(original),"native-gates-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
        string path=Path.Combine(folder,"private.xlsx");
        using(var book=new Workbook()){var sheet=book.Worksheets[0];sheet.Name="Data";sheet.Cells["A1"].Value=10d;sheet.Cells["A1"].Protection.Locked=false;book.SaveDocument(path,DocumentFormat.Xlsx);}
        byte[] baseline=File.ReadAllBytes(path);
        var owners=new System.Collections.Generic.List<WorkbookCalculationSession>();
        foreach(var engine in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired}){
            var session=await WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(engine,false,false,120000,true));owners.Add(session);
            using(var model=new EngineChangeManagerTests.Model(path))try{
                await model.Do(()=>{ModelSafetyManager.BeginBulkWorkbookMutation(0);Check(ModelSafetyManager.IsBulkWorkbookMutationInProgress(0),"ordinary model retains structural admission");ModelSafetyManager.EndBulkWorkbookMutation(0);return true;});
                var current=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{new WorkbookReadArea("Data",0,0,1,1)});
                await model.Do(()=>{
                    typeof(ModelChangeManagerV2).GetMethod("BindEngineEditingTrial",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(model.Manager,new object[]{session,current});
                    var rules=new WorkbookStructureRuleManager(0);
                    Refused(rules.AddRecords(WorkbookStructureRuleManager.RuleFundingRecords,10),"Funding insert refused before worksheet/name resolution");
                    Refused(rules.AddRecords(WorkbookStructureRuleManager.RuleDevelopmentIdentifiedRecords,10),"Development insert refused before worksheet/name resolution");
                    Refused(rules.DeleteRecords(WorkbookStructureRuleManager.RuleFundingRecords,new[]{11}),"selected deletion refused before boundary/mutation work");
                    Refused(rules.DeleteLastRecords(WorkbookStructureRuleManager.RuleFundingRecords,1),"delete-last refused before mutation work");
                    Refused(new TransactionalDBSynchroniser(0).FullTransactionalDBSync(),"full Transactional DB synchronisation cannot mutate display workbook");
                    try{ModelSafetyManager.BeginBulkWorkbookMutation(0);throw new Exception("Expected bulk refusal");}catch(InvalidOperationException){Check(!ModelSafetyManager.IsBulkWorkbookMutationInProgress(0),"generic bulk guard refuses before incrementing mutation state");}
                    try{TransactionalDBSnapshotManager.CreateSnapshotAndComparison(0);throw new Exception("Expected snapshot refusal");}catch(NotSupportedException){Check(model.Value.WB.Worksheets.Count==1,"snapshot refusal cannot add sheets or names");}
                    Check(session.IsCurrent(current)&&!model.Value.IsDirty&&!model.Manager.CanUndo&&model.Value.WB.Worksheets[0].Cells["A1"].Value.NumericValue==10d,"all refusals preserve current result, history, dirty state and display workbook");
                    var edit=model.Manager.ProcessEngineCommand(new[]{new DataChangeEvent{ModelID=0,WSName="Data",CellAddress="A1",ChangedValue=11d,DataFormat="N",Description="After structural refusal"}},new[]{WorkbookValuePermission.UnlockedCell},"After structural refusal");
                    Check(edit.BSuccess&&model.Value.WB.Worksheets[0].Cells["A1"].ModelValue().NumericValue==11d,"ordinary native editing remains available after refused structure actions");
                    return true;
                });
            }finally{await session.CloseAsync();}
        }
        Check(File.ReadAllBytes(path).SequenceEqual(baseline),"original generated file unchanged");
        await NativeProcessChecks.RequireOwnedProcesses(owners);
        Console.WriteLine("STRUCTURE_GATE ASSERTIONS="+checks);
    }
}
