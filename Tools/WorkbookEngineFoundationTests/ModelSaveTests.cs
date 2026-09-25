using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using Abovo;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;

static class ModelSaveTests
{
    const BindingFlags F=BindingFlags.Instance|BindingFlags.NonPublic;
    static int checks;
    static void Check(bool ok,string message){if(!ok)throw new Exception(message);Console.WriteLine("MODEL_SAVE PASS "+(++checks)+" "+message);}
    static object Call(object obj,string name,params object[] args){try{return obj.GetType().GetMethod(name,F).Invoke(obj,args);}catch(TargetInvocationException e){ExceptionDispatchInfo.Capture(e.InnerException).Throw();throw;}}
    static WorkbookCalculationSession Owner(ModelChangeManagerV2 manager)=>(WorkbookCalculationSession)manager.GetType().GetField("engineTrial",F).GetValue(manager);
    static void Edit(ModelChangeManagerV2 manager,double value){var response=manager.ProcessEngineCommand(new[]{new DataChangeEvent{ModelID=0,WSName="Data",CellAddress="A1",DataFormat="N",ChangedValue=value,Description="Native model save trial"}},new[]{WorkbookValuePermission.SolidFillRule},"Native model save trial");Check(response.BSuccess&&!response.BError,"model input committed through existing change manager");}
    static WorkbookPublicationReceipt Save(ModelChangeManagerV2 manager,string path,string folder,CancellationToken cancellation=default(CancellationToken))=>(WorkbookPublicationReceipt)Call(manager,"SaveEngineEditingModel",path,folder,cancellation);
    static async Task Refused(Func<Task> action,string message){try{await action();}catch(Exception e)when(e is IOException||e is InvalidOperationException||e is OperationCanceledException){Check(true,message);return;}throw new Exception("Expected rejection: "+message);}
    internal static async Task Run(string original)
    {
        var processes=NativeProcessChecks.ExcelIds();
        var owners=new List<WorkbookCalculationSession>();
        foreach(var extension in new[]{".xlsx",".xlsb"})foreach(var preference in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.Automatic})
        {
            string folder=Path.Combine(Path.GetDirectoryName(Path.GetFullPath(original)),"model-save-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
            string path=Path.Combine(folder,"private"+extension),other=Path.Combine(folder,"saved-as"+extension);
            using(var book=new Workbook()){
                var sheet=book.Worksheets[0];sheet.Name="Data";sheet.Cells["A1"].Value=10d;sheet.Cells["B1"].Value=20d;
                sheet.Range["A1:B1"].Protection.Locked=false;sheet.Range["A1:B1"].Fill.PatternType=PatternType.Solid;sheet.Cells["C1"].Formula="=A1+B1";
                book.CustomXmlParts.Add("<Model xmlns='urn:abovo:save-test'><Value>original</Value></Model>");
                book.CalculateFull();book.SaveDocument(path,extension==".xlsb"?DocumentFormat.Xlsb:DocumentFormat.Xlsx);
            }
            var session=await WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(preference,false,false,120000,true,true,true));
            owners.Add(session);
            Check(session.EngineName==(preference==WorkbookEnginePreference.Automatic?"Excel":"DevExpress"),"native save test uses intended engine "+preference+extension);
            using(var model=new EngineChangeManagerTests.Model(path))
            try
            {
                var initial=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{new WorkbookReadArea("Data",0,0,1,3)});
                await model.Do(()=>{Call(model.Manager,"BindEngineEditingTrial",session,initial);Edit(model.Manager,11d);return true;});
                EventHandler badView=(sender,args)=>{throw new InvalidOperationException("Synthetic dirty-state subscriber");};
                await model.Do(()=>{model.Value.DirtyStateChanged+=badView;return true;});
                var first=await model.Do(()=>Save(model.Manager,path,folder));
                owners.Add(await model.Do(()=>Owner(model.Manager)));
                await model.Do(()=>{model.Value.DirtyStateChanged-=badView;var current=Owner(model.Manager);
                    Check(!model.Value.IsDirty&&model.Manager.EngineEditingResult!=null&&current.SessionId!=session.SessionId,"successful verified save hands off owner and acknowledges exact dirty revision "+preference+extension);
                    Check(current.EngineName==session.EngineName,"Automatic save/reopen retains its selected engine");
                    Check(model.Manager.CanUndo&&model.Value.WB.Worksheets[0].Cells["A1"].ModelValue().NumericValue==11d,"save retains native input and existing Undo stack");
                    Check(model.Value.WB.Worksheets[0].Cells["A1"].Value.NumericValue==10d,"save does not overwrite presentation-map inputs");
                    var undo=model.Manager.Undo();Check(undo.BSuccess&&model.Value.IsDirty&&model.Value.WB.Worksheets[0].Cells["A1"].ModelValue().NumericValue==10d,"Undo after save changes new owner and restores dirty");return true;});
                var second=await model.Do(()=>Save(model.Manager,other,folder));
                owners.Add(await model.Do(()=>Owner(model.Manager)));
                await model.Do(()=>{Check(!model.Value.IsDirty&&model.Value.FileName==Path.GetFullPath(other)&&Owner(model.Manager).SourceHash==second.Hash,"Save As changes live identity only to its verified reopened file");
                    var redo=model.Manager.Redo();Check(redo.BSuccess&&model.Value.WB.Worksheets[0].Cells["A1"].ModelValue().NumericValue==11d,"Redo survives a second session handoff");return true;});
                var currentResult=await model.Do(()=>model.Manager.EngineEditingResult);
                await Refused(()=>model.Do(()=>Save(model.Manager,path,folder)),"Save As collision leaves both existing files unchanged");
                Check(await model.Do(()=>model.Value.IsDirty&&ReferenceEquals(model.Manager.EngineEditingResult,currentResult)),"refused destination retains current owner and dirty state");
                using(var cancelled=new CancellationTokenSource()){
                    cancelled.Cancel();await Refused(()=>model.Do(()=>Save(model.Manager,other,folder,cancelled.Token)),"cancelled save leaves live native model available");
                }
                await model.Do(()=>{
                    typeof(ModelChangeManagerV2).GetField("writingAndCalculating",F).SetValue(model.Manager,true);
                    try{bool refused=false;try{FileManager.CloseModel(0);}catch(InvalidOperationException){refused=true;}Check(refused&&ReferenceEquals(FileManager.ExcelModels[0],model.Value)&&!model.Value.IsClosing,"reentrant close cannot remove an admitted-operation model slot");}
                    finally{typeof(ModelChangeManagerV2).GetField("writingAndCalculating",F).SetValue(model.Manager,false);}return true;
                });
                await model.Do(()=>{model.Value.WB.CustomXmlParts.Add("<Unrouted xmlns='urn:abovo:save-test-extra'/>");return true;});
                await Refused(()=>model.Do(()=>Save(model.Manager,other,folder)),"unrouted custom XML changes block save before publication");
                await model.Do(()=>{model.Value.WB.CustomXmlParts.Remove(model.Value.WB.CustomXmlParts.Last());return true;});
                // A separate reader deliberately denies rename after native close.
                // Only disposable files are involved; retain the verified unsaved copy.
                var earlierCandidates=Directory.GetFiles(folder,"~Summit-candidate"+extension,SearchOption.AllDirectories);
                using(var denyRename=new FileStream(other,FileMode.Open,FileAccess.Read,FileShare.Read))
                    await Refused(()=>model.Do(()=>Save(model.Manager,other,folder)),"publication failure is reported with retained verified unsaved work");
                var retained=Directory.GetFiles(folder,"~Summit-candidate"+extension,SearchOption.AllDirectories).Except(earlierCandidates).Single();
                using(var verify=new Workbook()){verify.Options.CalculationMode=WorkbookCalculationMode.Manual;verify.LoadDocument(retained);
                    Check(verify.Worksheets[0].Cells["A1"].Value.NumericValue==11d,"failed publication retains the latest edited value in its verified candidate");}
                Check(await model.Do(()=>model.Value.IsDirty&&model.Manager.EngineEditingResult==null),"failed terminal save cannot clear dirty or expose old cached values");
                Check(File.Exists(first.BackupPath)&&File.Exists(other),"previous save and published workbook remain recoverable");
            }
            finally{await model.Do(()=>{Call(model.Manager,"CloseEngineEditingOwner");return true;});await session.CloseAsync();}
            using(var inspection=new Workbook()){
                inspection.Options.CalculationMode=WorkbookCalculationMode.Manual;inspection.LoadDocument(path);
                Check(inspection.Worksheets[0].Cells["A1"].Value.NumericValue==11d,"first saved source remains at its committed value");
                inspection.LoadDocument(other);
                Check(inspection.Worksheets[0].Cells["A1"].Value.NumericValue==10d,"Save As file was not overwritten by the failed later publication");
            }
        }
        await NativeProcessChecks.RequireOwnedProcesses(owners);
        Console.WriteLine("MODEL_SAVE ASSERTIONS="+checks);
    }
}
