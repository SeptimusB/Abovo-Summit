using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Threading.Tasks;
using Abovo;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;

static class NativeRecoveryTests
{
    const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Static|BindingFlags.Instance;
    static int checks;
    static void Check(bool ok,string message){if(!ok)throw new Exception(message);Console.WriteLine("BACKUP PASS "+(++checks)+" "+message);}
    static object Call(object instance,string name,params object[] args){try{return instance.GetType().GetMethod(name,F).Invoke(instance,args);}catch(TargetInvocationException e){ExceptionDispatchInfo.Capture(e.InnerException).Throw();throw;}}
    static object Static(string type,string name,params object[] args){try{return typeof(ModelChangeManagerV2).Assembly.GetType("Abovo."+type).GetMethod(name,F).Invoke(null,args);}catch(TargetInvocationException e){ExceptionDispatchInfo.Capture(e.InnerException).Throw();throw;}}
    static async Task Refused(Func<Task> action,string name){try{await action();}catch(Exception e)when(e is InvalidOperationException||e is IOException){Check(true,name);return;}throw new Exception("Expected refusal: "+name);}
    internal static async Task Run(string original)
    {
        var owners=new List<WorkbookCalculationSession>();
        foreach(var extension in new[]{".xlsx",".xlsm",".xlsb"})foreach(var preference in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired})
        {
            string folder=Path.Combine(Path.GetDirectoryName(original),"backup-native-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
            string path=Path.Combine(folder,"private"+extension);
            using(var book=new Workbook()){
                var ws=book.Worksheets[0];ws.Name="Data";ws.Cells["A1"].Value=10d;ws.Cells["A1"].Protection.Locked=false;
                ws.Cells["B1"].Formula="=A1*2";ws.Cells["D1"].DynamicArrayFormulaInvariant="=A1+{0;1;2}";
                book.CustomXmlParts.Add("<Model xmlns='urn:abovo:recovery-test'>Original</Model>");
                book.CalculateFullRebuild();book.SaveDocument(extension==".xlsb"?Path.Combine(folder,"generated.xlsx"):path,extension==".xlsm"?DocumentFormat.Xlsm:DocumentFormat.Xlsx);
            }
            if(extension==".xlsb"){
                var conversion=await NativeFixtureConversion.ExcelXlsb(Path.Combine(folder,"generated.xlsx"),path);
                await NativeProcessChecks.RequireOwnedIdentities(new[]{conversion});
            }
            byte[] baseline=File.ReadAllBytes(path);
            var s=await WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(preference,false,false,120000,true,true,true));owners.Add(s);
            using(var model=new EngineChangeManagerTests.Model(path))
            try
            {
                var initial=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{new WorkbookReadArea("Data",0,0,3,4)});
                Console.WriteLine("BACKUP INITIAL "+preference+extension+" spill="+initial.Blocks[0].ValueAt(2,3));
                await model.Do(()=>{Call(model.Manager,"BindEngineEditingTrial",s,initial);return true;});
                await Refused(()=>model.Do(()=>Static("RecoveryBackupStore","Write",model.Value)),"clean native model does not autosave "+preference+extension);
                for(int iteration=1;iteration<=2;iteration++)
                {
                    double value=10+iteration;
                    await model.Do(()=>{var edit=model.Manager.ProcessEngineCommand(new[]{new DataChangeEvent{ModelID=0,WSName="Data",CellAddress="A1",ChangedValue=value,DataFormat="N",Description="Recovery input "+iteration}},new[]{WorkbookValuePermission.UnlockedCell},"Recovery input");Check(edit.BSuccess,"typed edit admitted before recovery");return true;});
                    var accepted=await model.Do(()=>model.Manager.EngineEditingResult);
                    Check(Equals(accepted.Blocks[0].ValueAt(2,3),value+2),"accepted revision updates the native dynamic spill before recovery");
                    string recovery=await model.Do(()=>(string)Static("RecoveryBackupStore","Write",model.Value));
                    Check(Path.GetFileName(recovery)=="~private_recovery.xlsm","prefixed XLSM recovery name retained");
                    Check((string)Static("RecoveryBackupStore","ReadSource",recovery)==path,"source filename retained for recovery guidance");
                    using(var verify=new Workbook()){
                        verify.Options.CalculationMode=WorkbookCalculationMode.Manual;verify.LoadDocument(recovery);
                        var ws=verify.Worksheets["Data"];
                        Console.WriteLine("BACKUP VALUES "+preference+extension+" input="+ws.Cells["A1"].Value+" formula="+ws.Cells["B1"].Value+" spill="+ws.Cells["D3"].Value+" acceptedSpill="+accepted.Blocks[0].ValueAt(2,3)+" array="+ws.Cells["D1"].FormulaInvariant);
                        Check(ws.Cells["A1"].Value.NumericValue==value&&ws.Cells["B1"].Value.NumericValue==value*2&&ws.Cells["D3"].Value.NumericValue==value+2,"recovery contains current input, formula cache and dynamic spill");
                        Check(ws.Cells["B1"].HasFormula&&ws.Cells["D1"].GetDynamicArrayFormulaRange()!=null,"recovery retains formulas and dynamic-array behavior");
                        Check(verify.CustomXmlParts.Any(p=>p.CustomXmlPartDocument.OuterXml.Contains("urn:abovo:recovery-test"))&&verify.CustomXmlParts.Any(p=>p.CustomXmlPartDocument.OuterXml.Contains("Recovery input "+iteration)),"custom model XML and current history retained");
                    }
                    Check(await model.Do(()=>model.Value.IsDirty&&model.Manager.CanUndo&&s.IsCurrent(accepted)&&ReferenceEquals(model.Manager.EngineEditingResult,accepted)&&model.Value.FileName==path),"recovery never clears dirty, replaces owner, changes filename or loses Undo");
                    if(iteration==1){
                        byte[] good=File.ReadAllBytes(recovery);
                        await model.Do(()=>{model.Value.WB.CustomXmlParts.Add("<Unrouted xmlns='urn:abovo:recovery-reject'/>");return true;});
                        await Refused(()=>model.Do(()=>Static("RecoveryBackupStore","Write",model.Value)),"unrouted XML refuses backup without replacing previous recovery");
                        Check(File.ReadAllBytes(recovery).SequenceEqual(good),"previous recovery bytes retained on failed backup");
                        await model.Do(()=>{model.Value.WB.CustomXmlParts.Remove(model.Value.WB.CustomXmlParts.Last());return true;});
                    }
                }
                Check(await model.Do(()=>model.Manager.Undo().BSuccess&&model.Value.WB.Worksheets["Data"].Cells["A1"].ModelValue().NumericValue==11d),"native owner remains editable after two recovery exports");
                var current=await model.Do(()=>model.Manager.EngineEditingResult);var history=await model.Do(()=>model.Manager.CaptureSaveHistory(current));
                var probe=await s.CaptureCellsAsync(s.Revision,new[]{new WorkbookReadArea("Data",0,0,1,1)});
                var normal=await s.CreateSaveCandidateWithHistoryAsync(current,folder,probe,history);
                Check(Static("RecoveryBackupStore","ReadSource",normal.Path)==null,"normal save does not inherit temporary recovery identity");
                Console.WriteLine("NORMAL CANDIDATE "+normal.Path);
                var roundtrip=await WorkbookCalculationSession.OpenAsync(normal.Path,new WorkbookEngineOptions(WorkbookEnginePreference.ExcelRequired,false,false,120000,true));owners.Add(roundtrip);
                try{
                    var cell=await roundtrip.CaptureCellAsync(0,new WorkbookReadArea("Data",0,0,1,1));
                    var result=await roundtrip.ApplyValueAsync(cell,20d,WorkbookValuePermission.UnlockedCell,new[]{new WorkbookReadArea("Data",0,3,3,1)});
                    Check(Equals(result.Results.Blocks[0].ValueAt(2,0),22d),"normal saved file recalculates its dynamic spill in Excel");
                }finally{await roundtrip.CloseAsync();}
                using(var verify=new Workbook()){
                    verify.Options.CalculationMode=WorkbookCalculationMode.Manual;verify.LoadDocument(normal.Path);
                    var ws=verify.Worksheets["Data"];
                    // DX's XLSB importer represents Excel's cm-marked dynamic
                    // array as a legacy range. The Excel round trip above is
                    // the dynamic-behaviour assertion for this file format.
                    Check(extension==".xlsb"?ws.Cells["D1"].HasArrayFormula:ws.Cells["D1"].HasDynamicArrayFormula,"normal save retains the format's native array representation after recovery");
                    ws.Cells["A1"].Value=20d;verify.CalculateFullRebuild();
                    Check(ws.Cells["D3"].Value.NumericValue==22d,"normal saved file remains a working dynamic array, not only matching cached results");
                }
                var candidate=await s.CreateRecoveryCandidateAsync(current,folder,probe,history);
                await Refused(()=>s.PublishAndCloseAsync(candidate,path,WorkbookPublicationMode.ReplaceSource),"recovery candidate cannot replace original even when formats match");
                Check(s.IsCurrent(current),"refused recovery publication does not close live owner");
            }
            finally{await model.Do(()=>{Call(model.Manager,"CloseEngineEditingOwner");return true;});await s.CloseAsync();}
            Check(File.ReadAllBytes(path).SequenceEqual(baseline),"source workbook bytes unchanged");
        }
        await NativeProcessChecks.RequireOwnedProcesses(owners);
        Console.WriteLine("BACKUP ASSERTIONS="+checks);
    }
}
