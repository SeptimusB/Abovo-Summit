using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using DevExpress.Spreadsheet;

class CheckSheetFullFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static int checks;
 static object Get(object o,string n){var f=o.GetType().GetField(n,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(n,F).GetValue(o,null);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static object Invoke(Type t,string n,params object[] a){return t.GetMethod(n,F).Invoke(null,a);}
 static void Check(bool ok,string text){if(!ok)throw new Exception(text);checks++;Console.WriteLine("PASS "+text);}
 [STAThread] static int Main(string[] args){try{
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string path=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(path)?Assembly.LoadFrom(path):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));Invoke(app.GetType("Abovo.AbovoAppCls"),"Initialise");
  var files=app.GetType("Abovo.FileManager");Invoke(files,"Initialise",new object[]{null});
  string copy=Path.Combine(args[1],"private-full-watch.xlsb");File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!loaded.BError,"Private workbook opens");dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)loaded.IntegerReturn);
  IWorkbook wb=model.WB;var watch=app.GetType("Abovo.CheckSheetWatch");Invoke(watch,"Configure",false,5,2,false,false);
  var check=wb.Worksheets["Check Sheet"];var source=wb.Worksheets["Stock Assumptions"];
  // Synthetic probes are confined to the disposable in-memory copy, never saved.
  var input=source.Cells["ZZ1"];var dependent=source.Cells["ZZ2"];var output=check.Cells["B22"];
  wb.Options.CalculationMode=WorkbookCalculationMode.Manual;
  input.Value=0;dependent.Formula="=ZZ1*10";output.Formula="='Stock Assumptions'!ZZ2";check.Cells["C22"].Value="No";
  Invoke(watch,"RunCheck",(object)model);Check(output.Value.IsNumeric&&output.Value.NumericValue==0,"Baseline dependent check calculates to zero");
  dynamic service=Get(model,"WBCalculationService");service.DontCalcTDBS=true;
  input.Value=1;check.Calculate();
  Console.WriteLine("OLD SHEET-ONLY CHECK="+output.Value+" skipped="+service.DontCalcTDBS);
  Check(output.Value.IsNumeric&&output.Value.NumericValue==0,"Old sheet-only path misses the changed upstream value while Check Sheet is deferred");
  var engine=wb.Options.CalculationEngineType;var mode=wb.Options.CalculationMode;
  bool dirty=(bool)model.IsDirty,pending=(bool)model.ResultsPending,rebuild=(bool)Get(model,"NeedsFullRebuild");
  long revision=(long)Get(model,"CalculationRevision"),user=(long)Get(model,"UserChangeRevision");int history=wb.History.Count;
  var settings=GetSettings(app);string savedState=(string)Get(settings,"CheckSheetSuspendedFiles");
  var timer=System.Diagnostics.Stopwatch.StartNew();Invoke(watch,"RunCheck",(object)model);
  Console.WriteLine("FULL WATCH MS="+timer.ElapsedMilliseconds);
  Check(output.Value.IsNumeric&&output.Value.NumericValue==10&&dependent.Value.NumericValue==10,"Full watch picks up upstream dependency and Check Sheet failure");
  Check((bool)Get(model,"CheckSheetWarningActive"),"Imbalance updates the soft Check Sheet state");
  Check(service.DontCalcTDBS&&wb.Options.CalculationMode==mode&&wb.Options.CalculationEngineType==engine,"Deferred-sheet switch and calculation engine/mode restored");
  Check((bool)model.IsDirty==dirty&&(bool)model.ResultsPending==pending&&(bool)Get(model,"NeedsFullRebuild")==rebuild,"Full trial does not alter dirty/pending/rebuild certification");
  Check((long)Get(model,"CalculationRevision")==revision&&(long)Get(model,"UserChangeRevision")==user&&wb.History.Count==history,"No invented edit revision or history");
  Check((string)Get(settings,"CheckSheetSuspendedFiles")==savedState,"Unsaved trial does not persist a warning");
  input.Value=0;Invoke(watch,"RunCheck",(object)model);Check(output.Value.NumericValue==0&&!(bool)Get(model,"CheckSheetWarningActive"),"Correcting upstream input clears the soft warning on the next full watch");
  output.Formula="=1/0";bool rejected=false;
  try{Invoke(watch,"RunCheck",(object)model);}catch(TargetInvocationException){rejected=true;}
  Check(rejected&&service.DontCalcTDBS&&wb.Options.CalculationMode==mode&&wb.Options.CalculationEngineType==engine,"Formula-error outcome preserves engine and deferral state");
  Check(!(bool)app.GetType("Abovo.FormSplashScreen").GetProperty("OperationInProgress",F).GetValue(null,null),"Progress/operation guard released on error");
  Console.WriteLine("PASS "+checks+" assertions");return 0;
 }catch(Exception ex){Console.WriteLine(ex);return 1;}}
 static object GetSettings(Assembly app){return app.GetType("Abovo.RecoveryBackupSettings").GetField("Instance",F).GetValue(null);}
}
