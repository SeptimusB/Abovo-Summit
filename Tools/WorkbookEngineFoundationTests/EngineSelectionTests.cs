using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Abovo;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;

static class EngineSelectionTests
{
    const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
    static readonly Type Selection=typeof(ModelChangeManagerV2).Assembly.GetType("Abovo.WorkbookEngines.ModelEngineSelection");
    static int checks;
    static void Check(bool value,string text){if(!value)throw new Exception(text);Console.WriteLine("ENGINE_SELECTION PASS "+(++checks)+" "+text);Console.Out.Flush();}
    static void Activate(FileManager.ExcelModel model,WorkbookEngineOptions options,Func<WorkbookEnginePreference,IWorkbookCalculationBackend> factory=null){try{Selection.GetMethod("Activate",F).Invoke(null,new object[]{model,options,factory});}catch(TargetInvocationException e){System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(e.InnerException).Throw();throw;}}
    static string Status(FileManager.ExcelModel model)=>(string)Selection.GetMethod("Status",F).Invoke(null,new object[]{model});
    static IEnumerable<Control> Children(Control control){foreach(Control child in control.Controls){yield return child;foreach(var nested in Children(child))yield return nested;}}
    internal static async Task Run(string original)
    {
        string folder=Path.Combine(Path.GetDirectoryName(original),"engine-selection-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);string path=Path.Combine(folder,"private.xlsx");
        using(var book=new Workbook()){book.Worksheets[0].Name="Data";book.Worksheets[0].Cells["A1"].Value=10d;book.Worksheets[0].Cells["A1"].Protection.Locked=false;book.Worksheets[0].Cells["B1"].Formula="=A1*2";book.SaveDocument(path,DocumentFormat.Xlsx);}
        var bytes=File.ReadAllBytes(path);var owners=new List<WorkbookCalculationSession>();
        foreach(string scenario in new[]{"devexpress","unavailable","recovery","stress","excel"})using(var model=new EngineChangeManagerTests.Model(path)){
            WorkbookCalculationSession owner=null;
            try{await model.Do(()=>{
                int factories=0;
                Func<WorkbookEnginePreference,IWorkbookCalculationBackend> factory=kind=>{factories++;if(kind==WorkbookEnginePreference.ExcelRequired)throw new InvalidOperationException("Simulated unavailable Excel");return (IWorkbookCalculationBackend)Activator.CreateInstance(typeof(ModelChangeManagerV2).Assembly.GetType("Abovo.WorkbookEngines.DevExpressCalculationBackend"),true);};
                if(scenario=="recovery")typeof(FileManager.ExcelModel).GetField("RecoverySourcePath",F)?.SetValue(model.Value,"original.xlsb");
                if(scenario=="recovery"){
                    // The member is a property in current Summit builds.
                    typeof(FileManager.ExcelModel).GetProperty("RecoverySourcePath",F)?.SetValue(model.Value,"original.xlsb",null);
                }
                if(scenario=="stress"){model.Value.WB.Worksheets[0].Cells["C1"].Value="Y";model.Value.WB.DefinedNames.Add("StressTestMode","=Data!$C$1");}
                var preference=scenario=="devexpress"?WorkbookEnginePreference.DevExpressOnly:scenario=="excel"?WorkbookEnginePreference.ExcelRequired:WorkbookEnginePreference.Automatic;
                Activate(model.Value,new WorkbookEngineOptions(preference,false,false,120000,true,true,true),scenario=="excel"?null:factory);
                owner=(WorkbookCalculationSession)typeof(ModelChangeManagerV2).GetField("engineTrial",F).GetValue(model.Manager);if(owner!=null)owners.Add(owner);
                if(scenario=="excel"){
                    Check(model.Manager.HasEngineEditingTrial&&Status(model.Value).StartsWith("Excel "),"compatible Excel selected and identified");
                    Check(model.Value.WB.Worksheets[0].Cells["B1"].ModelValue().NumericValue==20d,"opening exposes calculated native values");
                    try{Activate(model.Value,new WorkbookEngineOptions(WorkbookEnginePreference.DevExpressOnly));throw new Exception("Expected live switch rejection");}catch(InvalidOperationException){Check(Status(model.Value).StartsWith("Excel "),"live engine switch rejected without replacing owner");}
                    var edit=model.Manager.ProcessEngineCommand(new[]{new DataChangeEvent{ModelID=0,WSName="Data",CellAddress="A1",ChangedValue=11d,DataFormat="N"}},new[]{WorkbookValuePermission.UnlockedCell},"Opening test");
                    Check(edit.BSuccess&&model.Value.WB.Worksheets[0].Cells["B1"].ModelValue().NumericValue==22d,"opened owner handles normal typed command");
                    using(var form=new ApplicationOptionsForm(0)){
                        Check(Children(form).Any(c=>c.Name=="CalculationEngineStatus"&&c.Text.Contains("Excel ")),"Options identifies current model engine");
                        var choice=Children(form).OfType<ComboBoxEdit>().Single(c=>c.Name=="CalculationEngineChoice");choice.SelectedIndex=1;
                        Check(Status(model.Value).StartsWith("Excel "),"changing Options selection does not migrate a live model");
                    }
                }else{
                    Check(!model.Manager.HasEngineEditingTrial&&Status(model.Value).StartsWith("DevExpress"),scenario+" retains ordinary DevExpress model");
                    Check(factories==(scenario=="unavailable"?2:0),scenario+" opens only the required native backends");
                    Check(!model.Value.IsDirty&&!model.Manager.CanUndo,scenario+" preserves dirty and history state");
                }
                return true;
            });}finally{if(owner!=null)await owner.CloseAsync();}
        }
        Check(File.ReadAllBytes(path).SequenceEqual(bytes),"generated opening source unchanged");await NativeProcessChecks.RequireOwnedProcesses(owners);Console.WriteLine("ENGINE_SELECTION ASSERTIONS="+checks);
    }
}
