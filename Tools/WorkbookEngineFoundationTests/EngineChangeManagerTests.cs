using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Threading;
using System.Threading.Tasks;
using Abovo;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;

static class EngineChangeManagerTests
{
    static int assertions;
    static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("BRIDGE PASS "+(++assertions)+" "+text);}
    static WorkbookReadArea Cell(int column)=>new WorkbookReadArea("Data",0,column,1,1);
    static object Invoke(object owner,string name,params object[] args)
    {try{return owner.GetType().GetMethod(name,BindingFlags.Instance|BindingFlags.NonPublic).Invoke(owner,args);}catch(TargetInvocationException e){System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(e.InnerException).Throw();throw;}}
    static void Succeeded(AbovoAppCls.AbovoTransaction result){if(!result.BSuccess||result.BError)throw new Exception(result.StrResponseMessage);}
    static async Task Reject(Func<Task> work,string label)
    {try{await work();}catch(Exception e)when(e is InvalidOperationException||e is ArgumentException){Check(true,label);return;}throw new Exception("Expected rejection: "+label);}
    internal sealed class Model:IDisposable
    {
        readonly BlockingCollection<Action> queue=new BlockingCollection<Action>();readonly Thread owner;
        Workbook book;FileManager.ExcelModel[] previous;
        internal FileManager.ExcelModel Value;internal ModelChangeManagerV2 Manager;
        internal Model(string source)
        {
            owner=new Thread(()=>{foreach(var work in queue.GetConsumingEnumerable())work();}){IsBackground=true};owner.SetApartmentState(ApartmentState.STA);owner.Start();
            Do(()=>{previous=FileManager.ExcelModels;book=new Workbook();book.Options.CalculationMode=WorkbookCalculationMode.Manual;book.LoadDocument(source);
                Value=(FileManager.ExcelModel)FormatterServices.GetUninitializedObject(typeof(FileManager.ExcelModel));Value.WB=book;Value.FileName=source;Value.WBCalcEngine=new CalcEngine(0);FileManager.ExcelModels=new[]{Value};int id=0;Manager=new ModelChangeManagerV2(ref id);Value.ChangeManager=Manager;return true;}).GetAwaiter().GetResult();
        }
        internal Task<T> Do<T>(Func<T> action){var t=new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);queue.Add(()=>{try{t.SetResult(action());}catch(Exception e){t.SetException(e);}});return t.Task;}
        public void Dispose(){Do(()=>{FileManager.ExcelModels=previous;book.Dispose();return true;}).GetAwaiter().GetResult();queue.CompleteAdding();owner.Join();queue.Dispose();}
    }
    static string Fixture(string source,bool date1904)
    {
        string folder=Path.Combine(Path.GetDirectoryName(source),"bridge-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);string path=Path.Combine(folder,"inputs.xlsx");
        using(var book=new Workbook())
        {
            book.DocumentSettings.Calculation.Use1904DateSystem=date1904;
            var ws=book.Worksheets[0];ws.Name="Data";
            ws.Cells["A1"].Value=10d;ws.Cells["B1"].Value=20d;ws.Cells["C1"].Value=30d;
            ws.Range["A1:C1"].Protection.Locked=false;ws.Range["A1:C1"].Fill.PatternType=PatternType.Solid;
            ws.Cells["A2"].Formula="=SUM(A1:C1)";ws.Cells["D1"].Formula="=A1*B1";
            ws.Cells["F1"].DynamicArrayFormulaInvariant="=A1+{0;1;2}";
            var check=book.Worksheets.Add("Check Sheet");check.Cells["A2"].Value="Synthetic model balance";
            check.Cells["B2"].Formula="=IF(Data!A2=60,0,1)";check.Cells["C2"].Value="No";check.Cells["C2"].Protection.Locked=false;
            check.DataValidations.Add(check.Range["C2"],DataValidationType.List,"Yes,No");
            check.Cells["D2"].Formula="=IF(C2=\"Yes\",0,B2)";check.Cells["E2"].Formula="=IF(B2=0,\"OK\",\"Check\")";
            check.Cells["F2"].Value="Fixture imbalance";check.Cells["H2"].Value="Data";
            book.DefinedNames.Add("Outputs_CheckSheet","='Check Sheet'!$A$2:$H$2");
            book.CalculateFullRebuild();book.SaveDocument(path,DocumentFormat.Xlsx);
        }
        return path;
    }
    static ModelEngineInput Input(WorkbookCellSnapshot cell,object value,string format="N")=>new ModelEngineInput(new DataChangeEvent{ModelID=0,WSName=cell.Area.Worksheet,CellAddress=cell.Area.Address.Split(':')[0],ChangedValue=value,DataFormat=format,Description="Engine input",TimeStamp=DateTime.UtcNow,UserName="Synthetic test"},cell,WorkbookValuePermission.UnlockedCell);
    static bool Warning(Model model)=>(bool)typeof(FileManager.ExcelModel).GetProperty("CheckSheetWarningActive",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(model.Value);
    static long UserRevision(Model model)=>(long)typeof(FileManager.ExcelModel).GetProperty("UserChangeRevision",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(model.Value);
    internal static async Task Run(string source)
    {
        int[] before=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).OrderBy(x=>x).ToArray();
        foreach(bool date1904 in new[]{false,true})
        foreach(var preference in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired})
        {
            var path=Fixture(source,date1904);var s=await WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(preference,false,false,120000,true,true));
            using(var model=new Model(path))
            try
            {
                var display=new[]{new WorkbookReadArea("Data",0,0,4,6)};
                var first=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Rebuild,display);
                Check(Equals(first.Blocks[0].ValueAt(2,5),12d),s.EngineName+" fixture genuinely spills before editing");
                await model.Do(()=>{Invoke(model.Manager,"BindEngineEditingTrial",s,first);return true;});
                Check(await model.Do(()=>model.Manager.HasEngineEditingTrial&&!model.Value.IsDirty&&model.Manager.EngineEditingResult==first),s.EngineName+" clean model binds real result owner, date1904="+date1904);
                await Reject(()=>Task.Run(()=>model.Manager.ProcessEngineChanges(new ModelEngineInput[0],"Wrong thread")),"change manager remains owner-thread confined");
                var cells=await s.CaptureCellsAsync(s.Revision,new[]{Cell(0),Cell(1),Cell(2)});
                var result=await model.Do(()=>model.Manager.ProcessEngineChanges(new[]{Input(cells[0],"12.5"),Input(cells[1],"25"),Input(cells[2],"3")},"Grouped input"));Succeeded(result);
                Check(result.IntegerReturn==3&&s.Revision==1,"three typed changes commit as one native revision");
                Check(await model.Do(()=>model.Value.IsDirty&&model.Manager.CanUndo&&!model.Manager.CanRedo&&model.Manager.GetHistoryTable().Rows.Count==3),"existing journal/dirty/Undo owns the group");
                Check(await model.Do(()=>model.Value.WB.Worksheets[0].Cells["A1"].Value.NumericValue==10&&model.Value.WB.Worksheets[0].Cells["A2"].HasFormula),"presentation workbook is not overwritten or independently edited");
                Check(await model.Do(()=>model.Value.WB.Worksheets[0].Cells["A1"].ModelValue().NumericValue==12.5&&model.Value.WB.Worksheets[0].Cells["A2"].ModelValue().NumericValue==40.5),"shared display reader obtains authoritative input and dependent value without modifying source cells");
                Check(await model.Do(()=>model.Value.WB.Worksheets[0].Cells["A2"].ModelDisplayText()=="40.5"),"uncached formatted read uses same calculation without recalculating");
                Check(await model.Do(()=>model.Manager.EngineEditingResult.Blocks[0].ValueAt(1,0).Equals(40.5d)&&model.Manager.EngineEditingResult.Blocks[0].ValueAt(2,5).Equals(14.5d)),"formula and dynamic spill display use native authoritative result");
                Check(await model.Do(()=>Warning(model)&&model.Value.WB.Worksheets["Check Sheet"].Cells["E2"].ModelDisplayText()=="Check"),"committed engine edit updates Check Sheet and company warning together");
                var history=await model.Do(()=>model.Manager.CaptureSaveHistory(model.Manager.EngineEditingResult));Check(history.EngineRevision==1&&history.UserRevision>0,"saved history capture binds current native and model revisions");
                Check(await model.Do(()=>!model.Manager.ProcessChange(new DataChangeEvent{WSName="Data",CellAddress="A1",ChangedValue=999d,DataFormat="N"}).BSuccess),"legacy edit cannot modify another workbook behind engine owner");
                await Reject(()=>model.Do(()=>{using(var output=new MemoryStream())Invoke(model.Value,"WriteRecoverySnapshot",output);return true;}),"legacy recovery cannot save stale presentation workbook");
                Succeeded(await model.Do(()=>model.Manager.Undo()));
                Check(await model.Do(()=>!model.Manager.CanUndo&&model.Manager.CanRedo&&model.Manager.EngineEditingResult.Blocks[0].ValueAt(1,0).Equals(60d)),"existing Undo atomically restores native group and outputs");
                Check(await model.Do(()=>!Warning(model)&&model.Value.WB.Worksheets["Check Sheet"].Cells["E2"].ModelDisplayText()=="OK"),"Undo clears Check Sheet and company warning at accepted revision");
                Succeeded(await model.Do(()=>model.Manager.Redo()));
                Check(await model.Do(()=>model.Manager.CanUndo&&!model.Manager.CanRedo&&model.Manager.EngineEditingResult.Blocks[0].ValueAt(1,0).Equals(40.5d)),"existing Redo restores native outputs");
                var stale=await model.Do(()=>model.Manager.ProcessEngineChanges(new[]{Input(cells[0],999d)},"stale"));
                Check(stale.BError&&await model.Do(()=>model.Manager.GetHistoryTable().Rows.Count)==3,"stale expected-before creates no history entry");
                Succeeded(await model.Do(()=>model.Manager.Undo()));
                cells=await s.CaptureCellsAsync(s.Revision,new[]{Cell(0),Cell(1),Cell(2)});
                var date=new DateTime(2026,9,30);var dateSerial=CellValue.FromDateTime(date,date1904).NumericValue;
                Succeeded(await model.Do(()=>model.Manager.ProcessEngineChanges(new[]{Input(cells[0],date,"D"),Input(cells[1],"=literal","S"),Input(cells[2],true,"BOOLEAN")},"Typed values")));
                var actual=await s.CaptureCellsAsync(s.Revision,new[]{Cell(0),Cell(1),Cell(2)});
                Check(Equals(actual[0].State.Value,dateSerial)&&Equals(actual[1].State.Value,"=literal")&&Equals(actual[2].State.Value,true),"date system, literal formula-looking text and Boolean preserved");
                Check(await model.Do(()=>model.Value.WB.Worksheets[0].Cells["A1"].ModelDateValue()==date),"native date display respects workbook date system");
                Check(await model.Do(()=>!model.Manager.CanRedo&&model.Manager.GetHistoryTable().Select("State='Superseded'").Length==3),"new edit supersedes previous undone group in same journal");
                Succeeded(await model.Do(()=>model.Manager.Undo()));
                Check((await s.CaptureCellAsync(s.Revision,Cell(0))).State.Value.Equals(10d),"typed group Undo restores input");
                var captured=await s.CaptureCellAsync(s.Revision,Cell(0));
                await s.ApplyValueAsync(captured,77d,WorkbookValuePermission.UnlockedCell);
                var rejected=await model.Do(()=>model.Manager.Redo());
                Check(rejected.BError&&await model.Do(()=>model.Manager.CanRedo),"history refuses externally changed native target without consuming Undo state");
            }
            finally{await s.CloseAsync();}
        }
        foreach(var preference in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired})
        {
            var path=Fixture(source,false);
            using(var book=new Workbook())
            {
                book.LoadDocument(path);var sheet=book.Worksheets[0];sheet.Cells["A1"].Value=0d;sheet.Cells["B1"].Value=CellValue.Empty;
                var cf=sheet.ConditionalFormattings.AddFormulaExpressionConditionalFormatting(sheet.Range["B1"],"=$A$1=0");cf.Formatting.Fill.PatternType=PatternType.Gray125;
                book.CalculateFull();book.SaveDocument(path,DocumentFormat.Xlsx);
            }
            var s=await WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(preference,false,false,120000,true));
            using(var model=new Model(path))
            try
            {
                var read=new[]{new WorkbookReadArea("Data",0,0,4,6)};var baseline=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,read,includePresentation:true);
                await model.Do(()=>{Invoke(model.Manager,"BindEngineEditingTrial",s,baseline);return true;});
                var cells=await s.CaptureCellsAsync(0,new[]{Cell(0),Cell(1)});
                Check(!cells[1].State.SolidFill&&cells[1].State.HasConditionalFormatting,s.EngineName+" effective pattern blocks amount before defining date");
                var blocked=new ModelEngineInput(Input(cells[1],100d).Change,cells[1],WorkbookValuePermission.SolidFillRule,false,true);
                var skipped=await model.Do(()=>model.Manager.ProcessEngineChanges(new[]{blocked},"Unavailable amount"));
                Check(skipped.BSuccess&&skipped.StrResponseMessage.Contains("1 unavailable")&&await model.Do(()=>!model.Value.IsDirty&&!model.Manager.CanUndo),"soft-skip is reported without journal or dirty change");
                cells=await s.CaptureCellsAsync(s.Revision,new[]{Cell(0),Cell(1)});
                var date=Input(cells[0],new DateTime(2026,9,30),"D");var amount=new ModelEngineInput(Input(cells[1],100d).Change,cells[1],WorkbookValuePermission.SolidFillRule,true);
                Succeeded(await model.Do(()=>model.Manager.ProcessEngineChanges(new[]{date,amount},"Schedule values")));
                var current=await s.CaptureCellAsync(s.Revision,Cell(1));
                Check(current.State.SolidFill&&Equals(current.State.Value,100d),s.EngineName+" recalculated defining date unlocks amount within one transaction");
                Check(await model.Do(()=>model.Manager.EngineEditingResult.Presentation[0].CellAt(0,1).Appearance.SolidFill),"grouped result publishes current effective fill with values");
                Succeeded(await model.Do(()=>model.Manager.Undo()));current=await s.CaptureCellAsync(s.Revision,Cell(1));
                Check(!current.State.SolidFill&&current.State.Value==null,"Undo restores both original amount and calculated lock pattern");
                Check(await model.Do(()=>model.Value.WB.Worksheets[0].Cells["B1"].ModelFill().PatternType==PatternType.Gray125),"shared display reader uses restored authoritative conditional pattern");
                Check(await model.Do(()=>!model.Manager.EngineEditingResult.Presentation[0].CellAt(0,1).Appearance.SolidFill),"Undo publishes restored fill at its result generation");
                Succeeded(await model.Do(()=>model.Manager.Redo()));current=await s.CaptureCellAsync(s.Revision,Cell(1));
                Check(current.State.SolidFill&&Equals(current.State.Value,100d),"Redo retains prerequisite calculation and unlocks amount again");
            }
            finally{await s.CloseAsync();}
        }
        var timer=Stopwatch.StartNew();int[] after;
        foreach(var preference in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired})
        {
            var path=Fixture(source,false);
            using(var book=new Workbook()){book.LoadDocument(path);book.DefinedNames.Add("EditorTargets","=Data!$A$1:$C$1");book.SaveDocument(path,DocumentFormat.Xlsx);}
            var s=await WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(preference,false,false,120000,true));
            using(var model=new Model(path))
            try
            {
                var read=new[]{new WorkbookReadArea("Data",0,0,4,6)};var baseline=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,read);
                await model.Do(()=>{Invoke(model.Manager,"BindEngineEditingTrial",s,baseline);return true;});
                var ticket=await model.Do(()=>model.Manager.CaptureEngineEditor("Data","$A$1",WorkbookValuePermission.UnlockedCell));
                var other=await model.Do(()=>model.Manager.CaptureEngineEditor("Data","B1",WorkbookValuePermission.UnlockedCell));
                var change=new DataChangeEvent{ModelID=0,WSName="Data",CellAddress="$A$1",ChangedValue=11d,DataFormat="N",EngineTicket=ticket,Description="Real editor",TimeStamp=DateTime.UtcNow};
                var wrong=change;wrong.CellAddress="B1";
                Check((await model.Do(()=>model.Manager.ProcessChange(wrong))).BError&&s.Revision==0,"editor token cannot target a different cell");
                Succeeded(await model.Do(()=>model.Manager.ProcessChange(change)));
                Check((await s.CaptureCellAsync(s.Revision,Cell(0))).State.Value.Equals(11d)&&await model.Do(()=>model.Manager.GetHistoryTable().Rows.Count)==1,s.EngineName+" ordinary ProcessChange commits only its admitted native cell");
                Check((await model.Do(()=>model.Manager.ProcessChange(change))).BError,"one-use ticket cannot be replayed");
                wrong.CellAddress="B1";wrong.EngineTicket=other;
                Check((await model.Do(()=>model.Manager.ProcessChange(wrong))).BError,"second editor becomes stale after a different edit");
                Succeeded(await model.Do(()=>model.Manager.Undo()));
                Check((await s.CaptureCellAsync(s.Revision,Cell(0))).State.Value.Equals(10d),"ordinary admitted edit uses existing Undo");
                ticket=await model.Do(()=>model.Manager.CaptureEngineEditor("Data","B1",WorkbookValuePermission.UnlockedCell));
                var heading=new DataChangeEvent{ModelID=0,TargetNR="EditorTargets",TargetNRIndex=1,NROrientation=System.Windows.Forms.Orientation.Horizontal,ChangedValue=22d,DataFormat="N",EngineTicket=ticket,Description="Heading editor"};
                Succeeded(await model.Do(()=>model.Manager.ProcessChangeByNRAddressing(heading)));
                Check((await s.CaptureCellAsync(s.Revision,Cell(1))).State.Value.Equals(22d),"named-range editor resolves and verifies its native target");
                await Reject(()=>model.Do(()=>model.Manager.CaptureEngineEditor("Data","D1",WorkbookValuePermission.UnlockedCell)),"formula cell cannot obtain an edit token");
                ticket=await model.Do(()=>model.Manager.CaptureEngineEditor("Data","A1",WorkbookValuePermission.UnlockedCell));
                change.EngineTicket=ticket;
                await s.CalculateAndReadAsync(s.Revision,WorkbookCalculationKind.Full,read);
                Check((await model.Do(()=>model.Manager.ProcessChange(change))).BError,"calculation-only refresh invalidates an already open editor");
                Check((await s.CaptureCellAsync(s.Revision,Cell(0))).State.Value.Equals(10d),"stale editor leaves native input untouched");
                var userRevision=await model.Do(()=>UserRevision(model));
                await model.Do(()=>{model.Value.WBCalcEngine.CalcFile(3);return true;});
                Check(await model.Do(()=>model.Manager.EngineEditingResult!=null&&!model.Value.ResultsPending&&UserRevision(model)==userRevision),"common rebuild restores current native display without inventing user edits");
                var revision=s.Revision;
                await model.Do(()=>{model.Value.WBCalcEngine.CalculateDependencySensitiveFile("Fixture");return true;});
                Check(s.Revision==revision&&await model.Do(()=>model.Value.WB.Worksheets[0].Cells["B1"].ModelValue().NumericValue==22d&&model.Value.WB.Worksheets[0].Cells["B1"].Value.NumericValue==20d),"dependency reader stays on selected owner and leaves template untouched");
                EventHandler brokenView=(sender,args)=>{throw new InvalidOperationException("Synthetic presentation subscriber");};
                await model.Do(()=>{model.Value.WBCalcEngine.CalculationCompleted+=brokenView;return true;});
                Succeeded(await model.Do(()=>model.Manager.ProcessEngineCommand(new[]{new DataChangeEvent{ModelID=0,WSName="Data",CellAddress="A1",ChangedValue=15d,DataFormat="N"},new DataChangeEvent{ModelID=0,WSName="Data",CellAddress="B1",ChangedValue=25d,DataFormat="N"}},new[]{WorkbookValuePermission.UnlockedCell,WorkbookValuePermission.UnlockedCell},"Paste command")));
                await model.Do(()=>{model.Value.WBCalcEngine.CalculationCompleted-=brokenView;return true;});
                Check(await model.Do(()=>model.Value.WB.Worksheets[0].Cells["A2"].ModelValue().NumericValue==70d),"explicit paste command captures targets and commits one native result");
                Succeeded(await model.Do(()=>model.Manager.Undo()));
                Check(await model.Do(()=>model.Value.WB.Worksheets[0].Cells["A1"].ModelValue().NumericValue==10d&&model.Value.WB.Worksheets[0].Cells["B1"].ModelValue().NumericValue==22d),"one Undo restores the admitted paste command");
                await model.Do(()=>{
                    using(var editor=new ModelPostingTextBox()){
                        int id=0;string sheet="Data",address="C1";
                        typeof(ModelPostingTextBox).GetMethod("Initialise",BindingFlags.Instance|BindingFlags.NonPublic|BindingFlags.Public).Invoke(editor,new object[]{id,sheet,address});
                        typeof(System.Windows.Forms.Control).GetMethod("OnEnter",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(editor,new object[]{EventArgs.Empty});
                        editor.EditValue="Native text";
                        Invoke(editor,"ProcessChange",editor,EventArgs.Empty);
                        Check(model.Value.WB.Worksheets[0].Cells["C1"].ModelValue().TextValue=="Native text","real standalone editor captures on Enter and posts through native manager");
                    }
                    return true;
                });
            }
            finally{await s.CloseAsync();}
        }
        timer.Restart();
        do{await Task.Delay(100);after=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).OrderBy(x=>x).ToArray();}while(timer.ElapsedMilliseconds<10000&&!before.SequenceEqual(after));
        Check(before.SequenceEqual(after),"owned Excel exits and existing user processes remain");
        Console.WriteLine("BRIDGE ASSERTIONS="+assertions);
    }
}
