using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using System.Security.Cryptography;
using System.Text;
using DevExpress.Spreadsheet;

public static class IdleIntegrityFixture {
    const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Static|BindingFlags.Instance;
    static Assembly app;
    static Type scheduler,files;
    static DateTime now;
    static string output;
    static Timer warningCloser;
    static int warningCount;
    static object Call(object target,string name,params object[] args) {
        return (target as Type??target.GetType()).GetMethods(F).Single(m=>m.Name==name&&m.IsStatic==(target is Type)&&m.GetParameters().Length==args.Length).Invoke(target is Type?null:target,args);
    }
    static object Field(object target,string name) {return target.GetType().GetField(name,F).GetValue(target);}
    static object Prop(object target,string name) {return target.GetType().GetProperty(name,F).GetValue(target);}
    static void Check(bool value,string message) {if(!value)throw new Exception(message);Console.WriteLine("PASS: "+message);}
    static object State(object model) {return ((IDictionary)scheduler.GetField("Plans",F).GetValue(null))[model];}
    static void Due(object model) {State(model).GetType().GetField("DueUtc",F).SetValue(State(model),now.AddMinutes(-1));}
    static bool Step(double idleSeconds) {return (bool)Call(scheduler,"ProcessIdle",now,TimeSpan.FromSeconds(idleSeconds));}
    static string Digest(IWorkbook w) {
        using(var hash=SHA256.Create()) {
            using(var stream=new CryptoStream(Stream.Null,hash,CryptoStreamMode.Write))using(var b=new BinaryWriter(stream,Encoding.UTF8)) {
                foreach(var s in w.Worksheets) {
                    b.Write(s.Name);b.Write(s.IsProtected);
                    foreach(var c in s.GetUsedRange().ExistingCells) {
                        if(!c.HasFormula&&c.Value.IsEmpty)continue;
                        b.Write(c.GetReferenceA1());b.Write(c.HasFormula?c.FormulaInvariant:c.Value.Type+":"+c.Value.ToString());
                    }
                    foreach(var n in s.DefinedNames){b.Write(n.Name);b.Write(n.RefersTo);}
                }
                foreach(var n in w.DefinedNames){b.Write(n.Name);b.Write(n.RefersTo);}
            }
            return BitConverter.ToString(hash.Hash);
        }
    }
    static string Finish(object model) {
        var clock=Stopwatch.StartNew();int units=0;
        do {
            // The real scheduler returns to the message loop after every unit.
            // Pump only in this fixture, never inside a production workbook operation.
            Application.DoEvents();
            if(!Step(121)){System.Threading.Thread.Sleep(10);if(clock.Elapsed.TotalMinutes>10)throw new Exception("Integrity gates never cleared");continue;}
            units++;if(units>30000)throw new Exception("Pass did not finish");
        }
        while(Field(State(model),"Work")!=null);
        string report=(string)Field(State(model),"LastReport");
        Check(report!=null,"Only a complete pass records a report");
        Console.WriteLine("UNITS="+units+" ACTIVE_SECONDS="+clock.Elapsed.TotalSeconds.ToString("F3"));
        Console.WriteLine(report);
        File.WriteAllText(Path.Combine(output,"report.txt"),report);
        return report;
    }
    static void CheckQuiet(bool v,string m) {if(!v)throw new Exception(m);}
    static void DrainNotices() {var timer=Stopwatch.StartNew();while(timer.ElapsedMilliseconds<1250){Application.DoEvents();System.Threading.Thread.Sleep(10);}}
    [STAThread] public static int Main(string[] args) {
        try {
            output=args[1];now=DateTime.UtcNow;
            AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var path=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(path)?Assembly.LoadFrom(path):null;};
            app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
            Trace.Listeners.Add(new ConsoleTraceListener());
            app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
            files=app.GetType("Abovo.FileManager");Call(files,"Initialise",new object[]{null});
            scheduler=app.GetType("Abovo.IdleIntegrityManager");Call(scheduler,"Configure",false,30,false);
            Call(app.GetType("Abovo.RecoveryBackupManager"),"Configure",false,10,false);
            // Prevent real timers from interfering with the deterministic injected clock.
            ((Timer)scheduler.GetField("Clock",F).GetValue(null)).Stop();
            warningCloser=new Timer{Interval=50};
            warningCloser.Tick+=(sender,e)=>{
                foreach(Form form in Application.OpenForms.Cast<Form>().ToArray())
                    if(form.Text=="Recovery autosaves paused"){warningCount++;form.DialogResult=DialogResult.No;form.Close();}
            };
            warningCloser.Start(); // Only acknowledge this fixture's own warning.
            if(args.Length>2)Real(args[2]);else Synthetic();
            return 0;
        } catch(Exception ex) {Console.Error.WriteLine(ex);return 1;}
        finally{if(warningCloser!=null)warningCloser.Dispose();}
    }
    static void Synthetic() {
        var type=app.GetType("Abovo.FileManager+ExcelModel");var ctor=type.GetConstructors().Single();
        dynamic model=ctor.Invoke(new object[]{0,Enum.ToObject(ctor.GetParameters()[1].ParameterType,0)});
        var models=Array.CreateInstance(type,1);models.SetValue(model,0);files.GetField("ExcelModels").SetValue(null,models);files.GetField("ExcelModelCount").SetValue(null,0);
        IWorkbook w=(IWorkbook)model.WB;w.Options.CalculationMode=WorkbookCalculationMode.Manual;
        var s=w.Worksheets[0];s.Name="Inputs";s.Cells["A1"].Value=10;s.Cells["B1"].FormulaInvariant="=A1*2";
        s.Cells["C1"].FormulaInvariant="=1/0";s.Cells["D1"].FormulaInvariant="=NA()";s.Cells["E1"].FormulaInvariant="=MATCH(42,A1:A1,0)";
        s.Cells["F1"].FormulaInvariant="=CONCATENATE("+String.Join(",",Enumerable.Repeat("A1",37))+")";
        for(int row=1;row<7000;row++)s.Cells[row,0].Value=row;
        var checks=w.Worksheets.Add("Check Sheet");checks.Cells["A1"].Value="Cash check";checks.Cells["E1"].Value="FAIL";checks.Cells["F1"].Value="Test issue";
        w.DefinedNames.Add("Outputs_CheckSheet","='Check Sheet'!$A$1:$H$1");
        w.DefinedNames.Add("BrokenTest","=#REF!");w.DefinedNames.Add("LiteralTest","=\"#REF!\"");
        w.Worksheets.Add("Transactional DB");w.DefinedNames.Add("IR_Journals","=Inputs!$A$1:$B$3");
        w.DefinedNames.Add("TransCopy_IR_Journals_01","='Transactional DB'!$A$1:$B$2");
        s.Protect("fixture-only",WorksheetProtectionPermissions.Default);
        string source=Path.Combine(output,"source.xlsm");model.ModelSpreadsheetControl.SaveDocument(source,DocumentFormat.Xlsm);model.FileName=source;model.FileInfo=new FileInfo(source);
        model.ChangeManager=(dynamic)Activator.CreateInstance(app.GetType("Abovo.ModelChangeManagerV2"),new object[]{0});
        model.TransDBSync=(dynamic)Activator.CreateInstance(app.GetType("Abovo.TransactionalDBSynchroniser"),new object[]{0});
        Application.DoEvents();model.IsDirty=false;
        string before=Digest(w);var mode=w.Options.CalculationMode;var engine=w.Options.CalculationEngineType;int history=w.History.Count;
        Call(scheduler,"Track",(object)model);Due((object)model);
        Check(!Step(121),"Disabled scheduler does nothing");
        Call(scheduler,"Configure",true,30,false);Check(!Step(121),"Schedule is not due before its interval");Due((object)model);
        Check(!Step(121),"No invisible calculation without a live notification window");
        using(var owner=new Form()) {
            owner.ShowInTaskbar=false;owner.Opacity=0;owner.Show();Application.DoEvents();
            Check(!Step(119.99),"At least two full idle minutes required");
            files.GetField("InternalBIsSaving",F).SetValue(null,true);
            try{Check(!Step(121),"Save blocks the next integrity task");}finally{files.GetField("InternalBIsSaving",F).SetValue(null,false);}
            using((IDisposable)model.ChangeManager.BeginChangeGroup("pending")){Check(!Step(121),"Pending grouped edit blocks integrity work");}
            var safety=app.GetType("Abovo.ModelSafetyManager");Call(safety,"BeginBulkWorkbookMutation",0);
            try{Check(!Step(121),"Structural mutation blocks integrity work");}finally{Call(safety,"EndBulkWorkbookMutation",0);}
            using(var modal=new Form()) {modal.Shown+=(o,e)=>{Check(!Step(121),"Modal dialog blocks integrity work");modal.Close();};modal.ShowDialog(owner);}
            using(var editor=(Control)Activator.CreateInstance(Type.GetType("DevExpress.XtraEditors.TextEdit, DevExpress.XtraEditors.v25.2",true))) {
                owner.Controls.Add(editor);editor.Focus();((dynamic)editor).IsModified=true;
                Check(!Step(121),"An unposted native editor blocks integrity work");
                ((dynamic)editor).IsModified=false;owner.Controls.Remove(editor);
            }
            model.RecoverySaveAsRequired=true;Check(!Step(121),"Recovery-required model is not automatically checked");model.RecoverySaveAsRequired=false;
            Check(Step(121),"Due idle calculation runs as one unit");
            Check(!(bool)model.ResultsPending&&w.Options.CalculationMode==mode&&w.Options.CalculationEngineType==engine,"Atomic calculation finishes and restores state");
            object pass=Field(State((object)model),"Work");Check(Field(pass,"Stage").ToString()=="Names"&&(bool)Prop(model,"RecoveryAutosaveSuspended"),"Calculation, refresh and Check Sheet failure recording finish together before yielding");
            Check(!Step(0)&&Object.ReferenceEquals(pass,Field(State((object)model),"Work")),"Input after calculation pauses next task and retains safe progress");
            Check(!Step(119),"Resume still requires another two idle minutes");
            DrainNotices();
            Check(Step(121)&&(bool)Prop(model,"RecoveryAutosaveSuspended"),"Remaining stages resume without losing the already recorded Check Sheet hold");
            Check(warningCount==0&&(bool)Prop(model,"CheckSheetWarningActive")&&Field(State((object)model),"LastReport")==null,"Non-modal warning state appears before final report; no popup");
            string report=Finish((object)model);
            Check(report.Contains("Check Sheet: row 1")&&report.Contains("BrokenTest")&&!report.Contains("Broken name: LiteralTest"),"Check Sheet and broken names reported without mistaking string literals");
            Check(report.Contains("Inputs!C1")&&report.Contains("Inputs!E1")&&!report.Contains("Inputs!D1"),"Only explicit NA() sentinel excluded; other N/A and division errors reported");
            Check(report.Contains("TransCopy_IR_Journals_01")&&report.Contains("Save compatibility notice: Inputs!F1")&&report.Contains("1 supported save-normalization notices"),"Mirror mismatch stays actionable; verified save normalization is a separate notice");
            Check(Digest(w)==before&&!(bool)model.IsDirty&&w.History.Count==history,"Formula/name/input/protection/dirty/history state unchanged by checks");
            Check((bool)model.CloseValidationRequired,"Known Check Sheet failure remains visible to existing close validation");
            Check((bool)Prop(model,"RecoveryAutosaveSuspended")&&warningCount==0,"Completing the report retains pause without a popup");
            DrainNotices();Due((object)model);Check(Step(121),"Unchanged workbook can be inspected at next interval");
            Check(!(bool)Field(Field(State((object)model),"Work"),"Calculated"),"Unchanged results do not trigger repeated full calculation");
            pass=Field(State((object)model),"Work");model.IsDirty=true;Check(Step(121),"Changed revision starts a fresh pass");
            Check(!Object.ReferenceEquals(pass,Field(State((object)model),"Work")),"No partial results mixed across revisions");
            Call(scheduler,"Configure",false,30,false);Check(Field(State((object)model),"Work")==null&&!Step(121),"Disabling discards a paused pass without further tasks");
            model.RequireFullRebuild();bool failed=false;
            try{Call((object)model,"CalculateForIdleIntegrity",new Action<string>(x=>{throw new InvalidOperationException("Injected calculation-stage failure");}));}catch(TargetInvocationException){failed=true;}
            Check(failed&&model.NeedsFullRebuild&&w.Options.CalculationMode==mode&&w.Options.CalculationEngineType==engine,"Failed calculation leaves results pending and restores calculation settings");
            checks.Cells["E1"].Value="OK";model.RequireFullRebuild();
            DrainNotices();Call(scheduler,"RequestNow",(object)model);now=DateTime.UtcNow;
            uint stamp=(uint)Call(app.GetType("Abovo.IntegrityInputClock"),"InputStamp");State((object)model).GetType().GetField("RequestedInputStamp",F).SetValue(State((object)model),(uint?)(stamp^1u));
            files.GetField("InternalBIsSaving",F).SetValue(null,true);try{Check(!Step(0)&&(bool)Field(State((object)model),"ManualStartPending"),"Explicit first-unit intent survives a save gate without bypassing it");}finally{files.GetField("InternalBIsSaving",F).SetValue(null,false);}
            Check(Step(0),"Explicit Run now starts despite input since the click, with schedules disabled");
            Check(!(bool)Prop(model,"CheckSheetWarningActive")&&!(bool)Prop(model,"RecoveryAutosaveSuspended")&&Field(Field(State((object)model),"Work"),"Stage").ToString()=="Names","Fresh passing Check Sheet clears remembered warning in the calculation unit, before any second scheduler tick");
            stamp=(uint)Call(app.GetType("Abovo.IntegrityInputClock"),"InputStamp");State((object)model).GetType().GetField("RequestedInputStamp",F).SetValue(State((object)model),(uint?)(stamp^1u));
            Check(!Step(0)&&!(bool)Prop(model,"CheckSheetWarningActive"),"Input after calculation still pauses the long scan but cannot strand the old heading warning");
            DrainNotices();string manualReport=Finish((object)model);Check(manualReport.Contains("Integrity check completed")&&!(bool)Field(State((object)model),"ManualRequested"),"Requested staged check completes and does not enable recurring checks");
            owner.Close();
        }
        using(var options=(Form)Activator.CreateInstance(app.GetType("Abovo.ApplicationOptionsForm"))) {
            options.Show();Application.DoEvents();dynamic tabs=options.Controls[0].Controls[0];tabs.SelectedTabPageIndex=2;Application.DoEvents();
            dynamic enabled=Field(options,"IntegrityEnabled"),interval=Field(options,"IntegrityMinutes");
            Check(tabs.SelectedTabPage.Text=="Integrity"&&!enabled.Checked&&Convert.ToInt32(interval.EditValue)==30,"Integrity tab displays opt-in 30 minute defaults");
            enabled.Checked=true;Check(interval.Enabled&&interval.Properties.MinValue==1&&interval.Properties.MaxValue==120,"Integrity checkbox and interval bounds work");
            using(var bitmap=new System.Drawing.Bitmap(options.Width,options.Height)){options.DrawToBitmap(bitmap,new System.Drawing.Rectangle(0,0,options.Width,options.Height));bitmap.Save(Path.Combine(output,"options.png"));}
            options.Close();Check(!(bool)scheduler.GetProperty("Enabled",F).GetValue(null),"Cancel does not persist edited settings");
        }
        model.CloseModel();Check(State((object)model)==null,"Close removes model and iterator state");
    }
    static void Real(string path) {
        string copy=Path.Combine(output,"input.xlsb");File.Copy(path,copy);
        var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
        Check(!result.BError,"Open private model copy");dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)result.IntegerReturn);
        IWorkbook w=(IWorkbook)model.WB;
        var engine=w.Options.CalculationEngineType;var mode=w.Options.CalculationMode;bool skip=model.WBCalculationService.DontCalcTDBS;
        using(var owner=new Form()) {
            owner.Opacity=0;owner.ShowInTaskbar=false;owner.Show();Application.DoEvents();
            var stock=w.DefinedNames.GetDefinedName("CurrStNo").Range[0,0];double original=stock.Value.NumericValue;
            dynamic change=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));change.ModelID=(int)model.ModelID;change.WSName=stock.Worksheet.Name;change.CellAddress=stock.GetReferenceA1();change.ChangedValue=original+1;change.DataFormat="I";change.Description="Idle integrity test input";
            Check(model.ChangeManager.ProcessChange(change).BSuccess,"Private workbook typed input is journalled before idle check");
            Application.DoEvents();string before=Digest(w);bool dirty=model.IsDirty;int history=w.History.Count;int journal=((System.Data.DataTable)model.ChangeManager.GetHistoryTable()).Rows.Count;
            Call(scheduler,"Configure",true,30,false);Due((object)model);string report=Finish((object)model);
            Check(!report.Contains("Mirror geometry"),"Real master mirror sizes match synchroniser rules");
            string after=Digest(w);Console.WriteLine("STATE digest="+(before==after)+" dirty="+dirty+"->"+model.IsDirty+" history="+history+"->"+w.History.Count);
            Check(before==after&&dirty==(bool)model.IsDirty&&history==w.History.Count,"All formulas/constants/names/order/protection and dirty/history preserved");
            Check(engine==w.Options.CalculationEngineType&&mode==w.Options.CalculationMode&&skip==(bool)model.WBCalculationService.DontCalcTDBS,"Real calculation scope restored");
            Check(!(bool)model.ResultsPending,"Real workbook results current after integrity pass");
            Check(journal==((System.Data.DataTable)model.ChangeManager.GetHistoryTable()).Rows.Count&&model.ChangeManager.Undo().BSuccess&&stock.Value.NumericValue==original,"Journal survives idle check and Undo restores the input");
            Check(model.ChangeManager.Redo().BSuccess&&stock.Value.NumericValue==original+1,"Redo still restores the user's edit");
            owner.Close();
        }
        Call(scheduler,"Configure",false,30,false);Call(files,"CloseModel",(int)model.ModelID);
    }
}
