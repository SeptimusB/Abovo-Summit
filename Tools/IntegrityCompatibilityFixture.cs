using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Functions;

public static class IntegrityCompatibilityFixture {
 const BindingFlags F=BindingFlags.Instance|BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic;
 static Assembly app;static Type recovery,store,modelType,files;static string bin,output;
 static object Call(object o,string name,params object[] a){return (o as Type??o.GetType()).GetMethods(F).Single(m=>m.Name==name&&m.IsStatic==(o is Type)&&m.GetParameters().Length==a.Length).Invoke(o is Type?null:o,a);}
 static object Get(object o,string name){return o.GetType().GetProperty(name,F).GetValue(o,null);}
 static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("PASS: "+text);}
 static string Hash(string path){using(var h=SHA256.Create())return BitConverter.ToString(h.ComputeHash(File.ReadAllBytes(path)));}
 static dynamic Model(int id){var c=modelType.GetConstructors().Single();return c.Invoke(new object[]{id,Enum.ToObject(c.GetParameters()[1].ParameterType,0)});}
 static void Calc(dynamic m){Call((object)m,"CalculateForIdleIntegrity",new object[]{null});}
 static bool Paused(dynamic m){return (bool)Get(m,"RecoveryAutosaveSuspended");}
 static object Validation(dynamic m){return Call((object)m,"ReadCheckSheetValidation");}
 static bool Record(dynamic m){object v=Validation(m);return (bool)Call((object)m,"RecordIdleCheckSheetResult",Get(m,"CalculationRevision"),Get(v,"HasFailures"));}
 static bool Gap(Cell c){return (bool)Call(Activator.CreateInstance(app.GetType("Abovo.ExpectedChartGapClassifier"),true),"IsExpected",c);}
 static void Change(dynamic m,double value){dynamic e=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));e.ModelID=0;e.WSName="Inputs";e.CellAddress="A1";e.ChangedValue=value;e.DataFormat="N";e.Description="Hold test";Check(m.ChangeManager.ProcessChange(e).BSuccess,"Journalled user edit");}
 [STAThread]public static int Main(string[] args){try{
  bin=args[0];output=args[1];AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(bin,new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  app=Assembly.LoadFrom(Path.Combine(bin,"Abovo-summit.exe"));recovery=app.GetType("Abovo.RecoveryBackupManager");store=app.GetType("Abovo.RecoveryBackupStore");modelType=app.GetType("Abovo.FileManager+ExcelModel");files=app.GetType("Abovo.FileManager");
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);Call(files,"Initialise",new object[]{null});Trace.Listeners.Add(new ConsoleTraceListener());
  if(args[2]=="--probe-pause"){dynamic m=Model(1);m.FileName=args[3];Call(recovery,"RestoreCheckSheetPause",(object)m,false);Check(Paused(m)==Boolean.Parse(args[4]),"Fresh process pause="+args[4]);m.CloseModel();return 0;}
  if(args[2]=="--synthetic"){Synthetic();DiscardAndPromptCases();}else WorkbookChecks(args[2],args[3]=="--ui");return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
 static void RestartProbe(string path,bool expected){
  var p=new ProcessStartInfo(Assembly.GetExecutingAssembly().Location,"\""+bin+"\" \""+output+"\" --probe-pause \""+path+"\" "+expected){UseShellExecute=false,CreateNoWindow=true,RedirectStandardOutput=true,RedirectStandardError=true};
  using(var child=Process.Start(p)){string stdout=child.StandardOutput.ReadToEnd(),stderr=child.StandardError.ReadToEnd();child.WaitForExit();Check(child.ExitCode==0,"Fresh-process persisted pause check: "+stdout+stderr);}
 }
 static void Synthetic(){dynamic m=Model(0);var models=Array.CreateInstance(modelType,1);models.SetValue(m,0);files.GetField("ExcelModels").SetValue(null,models);files.GetField("ExcelModelCount").SetValue(null,0);
  IWorkbook w=m.WB;w.Worksheets[0].Name="Inputs";w.Worksheets[0].Cells["A1"].Value=1;
  var c=w.Worksheets.Add("Check Sheet");c.Cells["A1"].Value="Fixture balance";c.Cells["B1"].FormulaInvariant="=IF(Inputs!A1>0,1,0)";c.Cells["C1"].Value="No";c.Cells["C1"].Protection.Locked=false;
  c.DataValidations.Add(c.Range["C1"],DataValidationType.List,"Yes,No");c.Cells["D1"].FormulaInvariant="=IF(C1=\"Yes\",0,B1)";c.Cells["E1"].FormulaInvariant="=IF(B1<>0,\"Check\",\"OK\")";c.Cells["F1"].Value="Fixture mismatch";w.DefinedNames.Add("Outputs_CheckSheet", "='Check Sheet'!$A$1:$H$1");
  string original=Path.Combine(output,"hold.xlsb");m.FileName=original;m.FileInfo=new FileInfo(original);m.ModelSpreadsheetControl.SaveDocument(original,DocumentFormat.Xlsb);m.ChangeManager=(dynamic)Activator.CreateInstance(app.GetType("Abovo.ModelChangeManagerV2"),new object[]{0});m.IsDirty=false;
  Change(m,2);string backup=(string)Call(store,"Write",(object)m),good=Hash(backup);Calc(m);Check(Record(m)&&Paused(m),"Fresh Check Sheet failure sets hold");
  Check(!Record(m),"Same failure does not repeat transition warning");RestartProbe(original.ToUpperInvariant(),true);
  Call(recovery,"ConfigureCheckSheetPolicy",true,false);Check(!Paused(m)&&(bool)Get(m,"CheckSheetWarningActive"),"Opt-in permits backup but retains failure warning");
  Check(File.Exists((string)Call(store,"Write",(object)m)),"Opt-in permits actual recovery write on failed Check Sheet");good=Hash(backup);
  Call(recovery,"ConfigureCheckSheetPolicy",false,false);Check(Paused(m),"Turning opt-in off immediately restores pause without another failure");
  bool rejected=false;try{Call(store,"Write",(object)m);}catch(TargetInvocationException){rejected=true;}Check(rejected&&Hash(backup)==good,"Direct recovery blocked; previous complete copy unchanged");
  Call(recovery,"Configure",true,1,false);Call(recovery,"Track",(object)m);
  var plans=(IDictionary)recovery.GetField("Plans",F).GetValue(null);object state=plans[(object)m];state.GetType().GetField("DueUtc",F).SetValue(state,DateTime.UtcNow.AddMinutes(-1));state.GetType().GetField("IdleDueUtc",F).SetValue(state,DateTime.UtcNow.AddMinutes(-1));
  using(var owner=new Form()){owner.ShowInTaskbar=false;owner.Opacity=0;owner.Show();Application.DoEvents();Call(recovery,"ProcessRecovery",DateTime.UtcNow,TimeSpan.FromMinutes(3));Check(Hash(backup)==good,"Due scheduler also respects Check Sheet hold");owner.Close();}
  Call(recovery,"Configure",false,10,false);Call(recovery,"Forget",(object)m);
  Check(m.SaveFile()&&Paused(m),"Normal Save is available without clearing hold");string copy=Path.Combine(output,"save-as.xlsb");Check(m.SaveFileAsTo(copy,true)&&Paused(m),"Save As carries hold to new name");RestartProbe(copy,true);RestartProbe(Path.Combine(output,"unrelated.xlsb"),false);
  long prior=(long)Get(m,"CalculationRevision");Change(m,-1);Check(Paused(m),"Correction alone does not clear hold");
  Check(!(bool)Call((object)m,"RecordIdleCheckSheetResult",prior,false)&&Paused(m),"Stale pass cannot clear hold");Check(!(bool)Call((object)m,"RecordIdleCheckSheetResult",Get(m,"CalculationRevision"),false)&&Paused(m),"Uncalculated pass cannot clear hold");
  Calc(m);Check(Record(m)&&!Paused(m),"Fresh current successful check clears hold");RestartProbe(copy,false);Check(File.Exists((string)Call(store,"Write",(object)m)),"New unsaved input may recover after fresh pass");
  ResolvedSave(m);
  Change(m,2);Calc(m);Record(m);c.Cells["C1"].Value="Yes";m.RequireFullRebuild();Calc(m);object result=Validation(m);
  Check(!(bool)Get(result,"HasFailures")&&((IList)result.GetType().GetField("OverriddenIssues").GetValue(result)).Count==1&&c.Cells["E1"].Value.TextValue=="Check","Workbook override accepted despite visible Check status");Check(Record(m)&&!Paused(m),"Fresh override-aware pass releases hold");
  c.Cells["D1"].FormulaInvariant="=0";m.RequireFullRebuild();Calc(m);Check((bool)Get(Validation(m),"HasFailures"),"Unrelated zero is not an override");c.Cells["D1"].FormulaInvariant="=IF(C1=\"Yes\",0,B1)";c.Cells["B1"].FormulaInvariant="=1/0";m.RequireFullRebuild();Calc(m);Check((bool)Get(Validation(m),"HasFailures"),"Override cannot hide a formula error");
  var chart=w.Worksheets.Add("OW - Live Covenant Calculation");chart.Cells["A1"].Value=0;chart.Cells["A2"].FormulaInvariant="=IF(A1=0,NA(),1)";chart.Cells["A3"].FormulaInvariant="=INDEX(A2:A2,1,1)";chart.Cells["A4"].FormulaInvariant="=IF(ISNA(INDEX(A2:A2,1,1)),NA(),1)";chart.Cells["A5"].FormulaInvariant="=IF(A6=0,NA(),1)";chart.Cells["A6"].FormulaInvariant="=MATCH(99,A1:A1,0)";chart.Cells["A7"].FormulaInvariant="=IF(ISNA(A6),NA(),1)";chart.Cells["A8"].FormulaInvariant="=IF(A6=0,NA(),1/0)";w.CalculateFullRebuild();
  foreach(string a in new[]{"A2","A3","A4"})Check(Gap(chart.Cells[a]),"Verified selected/inherited chart marker "+a);
  foreach(string a in new[]{"A5","A6","A7","A8"})Check(!Gap(chart.Cells[a]),"Unexpected predicate/lookup error remains visible "+a);
  c.Cells["B1"].Value=0;m.RequireFullRebuild();Calc(m);Record(m);m.FileName=original;Record(m);m.CloseModel();
 }
 delegate bool WindowVisitor(IntPtr hwnd,IntPtr data);
 [DllImport("kernel32.dll")]static extern uint GetCurrentThreadId();
 [DllImport("user32.dll")]static extern bool EnumThreadWindows(uint thread,WindowVisitor visitor,IntPtr data);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetWindowText(IntPtr hwnd,System.Text.StringBuilder text,int max);
 [DllImport("user32.dll")]static extern IntPtr SendMessage(IntPtr hwnd,uint msg,IntPtr wParam,IntPtr lParam);
 static object CloseAnswer(dynamic model,int answer){
  bool answered=false;uint thread=GetCurrentThreadId();
  using(var timer=new System.Windows.Forms.Timer()){timer.Interval=20;timer.Tick+=(s,e)=>EnumThreadWindows(thread,(hwnd,data)=>{
   var title=new System.Text.StringBuilder(200);GetWindowText(hwnd,title,200);if(title.ToString()=="Close business plan"){answered=true;SendMessage(hwnd,0x111,(IntPtr)answer,IntPtr.Zero);}return true;},IntPtr.Zero);
   timer.Start();object result=Call((object)model,"CommitToCloseModel");timer.Stop();Check(answered,"Owned close dialog answered "+answer);return result;
  }
 }
 static void DiscardAndPromptCases(){
  Call(recovery,"Configure",false,10,false);
  foreach(string test in new[]{"discard","saved-before-edit","saved-failure","failed-on-open","previous-hold","previous-pass","external-change","changed-during-open","failed-save","bad-recovery"}){
   dynamic m=Model(0);var models=Array.CreateInstance(modelType,1);models.SetValue(m,0);files.GetField("ExcelModels").SetValue(null,models);files.GetField("ExcelModelCount").SetValue(null,0);
   IWorkbook w=m.WB;w.Worksheets[0].Name="Inputs";w.Worksheets[0].Cells["A1"].Value=test=="failed-on-open"?1:-1;
   var sheet=w.Worksheets.Add("Check Sheet");sheet.Cells["A1"].Value="Fixture";sheet.Cells["E1"].FormulaInvariant="=IF(Inputs!A1>0,\"Check\",\"OK\")";w.DefinedNames.Add("Outputs_CheckSheet","='Check Sheet'!$A$1:$H$1");
   string path=Path.Combine(output,"discard-"+test+".xlsb");m.FileName=path;m.FileInfo=new FileInfo(path);Calc(m);m.ModelSpreadsheetControl.SaveDocument(path,DocumentFormat.Xlsb);m.IsDirty=false;
   m.ChangeManager=(dynamic)Activator.CreateInstance(app.GetType("Abovo.ModelChangeManagerV2"),new object[]{0});
   if(test.StartsWith("previous-")){Call(recovery,"PersistCheckSheetPause",(object)m,true);Call(recovery,"RestoreCheckSheetPause",(object)m,false);Check((bool)Get(m,"CheckSheetWarningNeedsRecheck"),"Restored warning is labelled previous/recheck");}
   Call((object)m,"CaptureCheckSheetDiskBeforeLoad");
   if(test=="changed-during-open")File.AppendAllText(path,"changed during load");
   Call((object)m,"CaptureCheckSheetOpenState");
   if(test=="previous-pass"){Calc(m);Record(m);Check(!Paused(m),"Fresh check clears previous hold before new unsaved session");}
   if(test=="saved-before-edit")m.ModelSpreadsheetControl.SaveDocument(path,DocumentFormat.Xlsb);
   string original=Hash(path);Change(m,2);Calc(m);Record(m);
   Check(Paused(m)&&!(bool)Get(m,"CheckSheetWarningNeedsRecheck"),"Current failed result replaces remembered state: "+test);
   string recoveryPath=null;
   if(test=="bad-recovery"){Call(recovery,"ConfigureCheckSheetPolicy",true,false);recoveryPath=(string)Call(store,"Write",(object)m);Call(recovery,"ConfigureCheckSheetPolicy",false,false);}
   if(test=="saved-failure")Check(m.SaveFile(),"Saving a failure keeps it outside the discard exception");
   if(test=="external-change")File.AppendAllText(path,"external change");
   if(test=="failed-save"){bool failed=false;try{Call((object)m,"SavePreparedWorkbook",new Action(()=>{throw new IOException("Injected failure before write");}),false);}catch(TargetInvocationException){failed=true;}Check(failed,"Failed save did not create a saved point");}
   bool clears=test=="discard"||test=="previous-pass"||test=="failed-save"||test=="bad-recovery";
   if(test=="discard"){
    dynamic cancelled=CloseAnswer(m,2);Check((string)cancelled.StringReturn=="Cancel"&&Paused(m),"Cancelled close retains warning");
    dynamic proceed=CloseAnswer(m,7);Check((string)proceed.StringReturn=="Proceed"&&Paused(m),"Discard decision defers clearing until actual close");
   }
   m.CloseModel();RestartProbe(path,!clears);
   if(clears)Check(Hash(path)==original,"Discard leaves saved original bytes intact: "+test);
   if(recoveryPath!=null){RestartProbe(recoveryPath,true);Call(recovery,"PersistCheckSheetPausePath",recoveryPath,false,-1);}
   Call(recovery,"PersistCheckSheetPausePath",path,false,-1);
  }
  // Exercise the real DevExpress Yes/No prompt in this fixture's UI thread.
  dynamic target=Model(7);target.FileName=Path.Combine(output,"prompt.xlsb");var all=Array.CreateInstance(modelType,8);all.SetValue(target,7);files.GetField("ExcelModels").SetValue(null,all);files.GetField("ExcelModelCount").SetValue(null,7);
  Call(recovery,"PersistCheckSheetPause",(object)target,true);var integrity=app.GetType("Abovo.IdleIntegrityManager");Call(integrity,"Configure",false,30,false);
  foreach(var answer in new[]{DialogResult.No,DialogResult.Yes}){
   Call(recovery,"Track",(object)target);int shown=0;
   using(var owner=new Form())using(var timer=new System.Windows.Forms.Timer()){
    owner.Opacity=0;owner.ShowInTaskbar=false;owner.Show();timer.Interval=20;timer.Tick+=(s,e)=>{foreach(Form form in Application.OpenForms.Cast<Form>().ToArray())if(form.Text=="Previous Check Sheet warning"){shown++;form.DialogResult=answer;}};timer.Start();
    Call(recovery,"ProcessRecovery",DateTime.UtcNow,TimeSpan.Zero);timer.Stop();Check(shown==1,"Remembered-warning prompt shown even with backups disabled: "+answer);
    Call(recovery,"ProcessRecovery",DateTime.UtcNow,TimeSpan.Zero);Check(shown==1,"Prompt is once per opening");owner.Close();
   }
   var plans=(IDictionary)integrity.GetField("Plans",F).GetValue(null);
   Check(answer==DialogResult.Yes?plans.Contains((object)target)&&(bool)plans[(object)target].GetType().GetField("ManualRequested",F).GetValue(plans[(object)target]):!plans.Contains((object)target),"Yes targets the exact model; No does not start a check");
   Call(integrity,"Forget",(object)target);
  }
  Call(recovery,"PersistCheckSheetPause",(object)target,false);target.CloseModel();
 }
 static void ResolvedSave(dynamic m){
  Check(m.SaveFile(),"Save previous input changes before clean warning-clear test");Calc(m);
  Check(!(bool)m.IsDirty&&!(bool)Get(m,"ManualSaveAvailable"),"Clean saved model has no optional save offer");
  Call((object)m,"RestoreRecoveryAutosaveHold",true);
  long revision=(long)Get(m,"CalculationRevision"),user=(long)Get(m,"UserChangeRevision");IWorkbook w=m.WB;int history=w.History.Count;
  Check(!(bool)Call((object)m,"RecordIdleCheckSheetResult",revision-1,false)&&!(bool)Get(m,"ManualSaveAvailable"),"Stale pass cannot offer Save");
  Check(Record(m)&&!Paused(m)&&(bool)Get(m,"ManualSaveAvailable"),"Clearing a remembered warning enables optional manual Save");
  Check(!(bool)m.IsDirty&&!(bool)Get(m,"HasUnsavedUserChanges")&&!(bool)m.CloseValidationRequired&&!(bool)m.ResultsPending&&!(bool)m.NeedsFullRebuild&&(long)Get(m,"CalculationRevision")==revision&&(long)Get(m,"UserChangeRevision")==user&&w.History.Count==history,"Save offer invents no edits, autosave work, close prompt, calculation or history");
  Check(!Record(m)&&(bool)Get(m,"ManualSaveAvailable"),"Repeated successful checks retain an unconsumed save offer");
  using(var stream=new MemoryStream())Call((object)m,"WriteRecoverySnapshot",stream);
  Check((bool)Get(m,"ManualSaveAvailable")&&!(bool)Get(m,"HasUnsavedUserChanges"),"Recovery serialization does not consume optional manual Save");
  Check(!(bool)Call((object)m,"SavePreparedWorkbook",new Action(()=>{}),false)&&(bool)Get(m,"ManualSaveAvailable")&&!(bool)m.IsDirty,"Cancelled save retains optional Save without inventing edits");
  bool failed=false;try{Call((object)m,"SavePreparedWorkbook",new Action(()=>{throw new IOException("Injected optional-save failure");}),false);}catch(TargetInvocationException){failed=true;}
  Check(failed&&(bool)Get(m,"ManualSaveAvailable")&&!(bool)m.IsDirty,"Failed save retains optional Save");
  int saves=0;EventHandler saved=(s,e)=>saves++;m.ModelSpreadsheetControl.DocumentSaved+=saved;
  try{
   Check(m.SaveFile()&&saves==1&&!(bool)Get(m,"ManualSaveAvailable")&&!(bool)m.IsDirty,"Optional Save really writes the workbook once, then becomes unavailable");
   Calc(m);Check(!Record(m)&&!(bool)Get(m,"ManualSaveAvailable"),"Unchanged healthy recheck does not re-enable Save");
   Check(m.SaveFile()&&saves==1,"Subsequent clean Save remains a no-op");
   Call((object)m,"RestoreRecoveryAutosaveHold",true);Check(Record(m),"Second genuine clearance offers saving again");
   Check(m.SaveFileAsTo(Path.Combine(output,"resolved-save-as.xlsb"),true)&&saves==2&&!(bool)Get(m,"ManualSaveAvailable"),"Successful Save As also consumes the save offer");
   Calc(m);Check((bool)Call((object)m,"RecordIdleCheckSheetResult",Get(m,"CalculationRevision"),true)&&!(bool)Get(m,"ManualSaveAvailable"),"Recording a failure alone does not enable optional Save");
   Check(Record(m),"Fresh pass clears the synthetic failure");
  }finally{m.ModelSpreadsheetControl.DocumentSaved-=saved;}
 }
 static void WorkbookChecks(string path,bool ui){
  using(var w=new Workbook()){
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;w.LoadDocument(path);
   int expected=0,remaining=0;object gaps=Activator.CreateInstance(app.GetType("Abovo.ExpectedChartGapClassifier"),true);
   foreach(string s in new[]{"OW - Charts Source Data","OW - Covenant Calculation","OW - Live Covenant Calculation"})foreach(var c in w.Worksheets[s].GetUsedRange().ExistingCells)if(c.Value.IsError){if((bool)Call(gaps,"IsExpected",c))expected++;else{remaining++;if(remaining<10)Console.WriteLine("REVIEW: "+s+"!"+c.GetReferenceA1()+" "+c.FormulaInvariant);}}
   Console.WriteLine("CHART_EXPECTED="+expected+" CHART_REVIEW="+remaining);Check(expected>0,"Native workbook intentional gaps recognised conservatively");
  }
  if(ui)Ui(path);
 }
 static object Field(object o,string name){return o.GetType().GetField(name,F).GetValue(o);}
 static void Drain(){var timer=Stopwatch.StartNew();while(timer.ElapsedMilliseconds<1400){Application.DoEvents();System.Threading.Thread.Sleep(10);}}
 static void Ui(string source){
  string copy=Path.Combine(output,"private-ui.xlsb");File.Copy(source,copy);var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});Check(!loaded.BError,"Private model opened for native UI tests");
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)loaded.IntegerReturn);
  Check(Field(model,"_openingCheckSheetHash")!=null,"Full model load captures a disk fingerprint");
  bool acceptable=!(bool)Get(Validation(model),"HasFailures")&&!(bool)Get(model,"CheckSheetWarningActive");
  Check((bool)Field(model,"_openingCheckSheetAcceptable")==acceptable,"Full model load records the actual opening Check Sheet state without a rebuild");
  string originalHash=Hash(copy);
  try{
   using(var owner=new Form())using(var card=(Control)Activator.CreateInstance(app.GetType("FileInstanceInterface"),new object[]{(int)model.ModelID}))using(var group=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})){
    owner.ShowInTaskbar=false;owner.Opacity=0;owner.Controls.Add(card);card.Dock=DockStyle.Fill;owner.ClientSize=new System.Drawing.Size(1100,650);owner.Show();Call(card,"PopulateFileInfo");
    group.ShowInTaskbar=false;group.Opacity=0;group.ClientSize=new System.Drawing.Size(1400,850);group.Show();Drain();
    dynamic button=Field(group,"CheckSheetButton");Check(button.Visibility.ToString()=="Never","Clean company header has no warning");
    Call((object)model,"SetCheckSheetWarning",true,false);Drain();
    Check(button.Visibility.ToString()=="Always"&&button.Caption=="(Check sheet)"&&button.ItemAppearance.Normal.ForeColor==System.Drawing.Color.Firebrick,"Warning is a red native button before company name");
    dynamic bar=Get(group,"BarTopBar");Check(Object.ReferenceEquals(bar.ItemLinks[0].Item,button)&&bar.ItemLinks.Count==2,"Warning precedes a single company header");
    var browser=(WebBrowser)Get(card,"WebBrowserBPInfo");Check(browser.DocumentText.Contains("summit-checksheet://open")&&browser.DocumentText.Contains("(Check sheet)"),"Main file summary has the same warning link");
    Call((object)model,"SetCheckSheetWarning",true,true);Drain();
    Check(button.Caption=="(Check sheet: recheck required)"&&browser.DocumentText.Contains("(Check sheet: recheck required)"),"Remembered warning is labelled recheck in both live interfaces");
    Call((object)model,"SetCheckSheetWarning",true,false);Drain();
    Check(button.Caption=="(Check sheet)"&&!browser.DocumentText.Contains("(Check sheet: recheck required)"),"Current failure replaces remembered-warning captions");
    button.PerformClick();Drain();
    var instances=(IEnumerable)model.WBInterface.GroupInterfaces;object destination=null;
    foreach(dynamic entry in instances){if((int)entry.GSID==0){destination=entry.RenderedForm;break;}}
    Check(destination!=null,"Warning click opens the matching plan's Assumptions interface");
    var active=(Control)Field(destination,"ActiveInterface");dynamic tabs=Get(active,"XtraTabControlNewGIT");Check(tabs.SelectedTabPage.Text.ToString().Trim()=="Check Sheet","Warning click selects Check Sheet tab, not default Global section");
    var destinationForm=(Form)destination;using(var bitmap=new System.Drawing.Bitmap(destinationForm.Width,destinationForm.Height)){destinationForm.DrawToBitmap(bitmap,new System.Drawing.Rectangle(0,0,destinationForm.Width,destinationForm.Height));bitmap.Save(Path.Combine(output,"warning-header.png"));}
    Call((object)model,"SetCheckSheetWarning",false,false);Drain();Check(button.Visibility.ToString()=="Never"&&!browser.DocumentText.Contains("summit-checksheet://open"),"Fresh success removes both warnings without reopening interfaces");
    group.Close();owner.Close();
   }
  }finally{Call((object)model,"SetCheckSheetWarning",false,false);model.CloseModel();Check(Hash(copy)==originalHash,"Full-load/UI warning tests leave the saved workbook unchanged");}
 }
}
