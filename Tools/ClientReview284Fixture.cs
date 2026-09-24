using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.Spreadsheet;

class ClientReview284Fixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static object Field(object o,string n){var f=o.GetType().GetField(n,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(n,F).GetValue(o,null);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static object Static(Type t,string n,params object[] a){return t.GetMethod(n,F).Invoke(null,a);}
 static int count;
 static void Check(bool v,string m){if(!v)throw new Exception(m);count++;Console.WriteLine("PASS: "+m);}
 static void Pump(){Application.DoEvents();}
 static System.Collections.Generic.IEnumerable<Control> Controls(Control c){foreach(Control child in c.Controls){yield return child;foreach(var d in Controls(child))yield return d;}}
 [STAThread] static int Main(string[] args){try {
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  string copy=Path.Combine(args[1],"private-284.xlsb");File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!result.BError,"Private workbook opens");dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)result.IntegerReturn);
  var backup=app.GetType("Abovo.RecoveryBackupManager");var settings=app.GetType("Abovo.RecoveryBackupSettings").GetField("Instance",F).GetValue(null);
  string originalSettings=(string)Field(settings,"CheckSheetSuspendedFiles");var clock=app.GetType("Abovo.CheckSheetWatch");
  try {
   Static(clock,"Configure",false,5,2,false,false);
   IWorkbook w=(IWorkbook)model.WB;var sheet=w.Worksheets["Check Sheet"];var cell=sheet.Cells["C22"];
   var numberFormat=cell.NumberFormat;dynamic edit=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));edit.ModelID=(int)model.ModelID;edit.WSName=sheet.Name;edit.CellAddress="C22";edit.DataFormat="B";edit.Description="Yes No regression";edit.ChangedValue="Yes";
   Check(model.ChangeManager.ProcessChange(edit).BSuccess&&cell.Value.IsText&&cell.Value.TextValue=="Yes","Legacy Boolean editor writes literal Yes, never 1");
   edit.ChangedValue=null;Check(model.ChangeManager.ProcessChange(edit).BSuccess&&cell.Value.IsEmpty,"Clearing Yes No keeps a blank, not a number");
   Check(model.ChangeManager.Undo().BSuccess&&cell.Value.IsText&&cell.Value.TextValue=="Yes","Undo restores literal Yes");
   Check(model.ChangeManager.Redo().BSuccess&&cell.Value.IsEmpty,"Redo restores blank");
   edit.ChangedValue="No";edit.DataFormat="S";Check(model.ChangeManager.ProcessChange(edit).BSuccess&&cell.Value.TextValue=="No"&&cell.NumberFormat==numberFormat,"No and workbook number format preserved");
   var rev=(long)Field(model,"CalculationRevision");Call(model,"RecordIdleCheckSheetResult",rev,true,true);
   Check((bool)Field(model,"CheckSheetWarningActive"),"Current imbalance is a soft visible state");
   Check((string)Field(settings,"CheckSheetSuspendedFiles")==originalSettings,"An unsaved failed check does not persist a pause");
   Static(backup,"ConfigureCheckSheetPolicy",true,false);Check(!(bool)Field(model,"RecoveryAutosaveSuspended"),"Allowed recovery is not blocked by a balance finding");
   Static(backup,"ConfigureCheckSheetPolicy",false,false);Check((bool)Field(model,"RecoveryAutosaveSuspended"),"Explicit pause option still works");
   Call(model,"RecordIdleCheckSheetResult",rev,false,true);Check(!(bool)Field(model,"CheckSheetWarningActive"),"Passing Check Sheet clears its indicator");
   Call(model,"RestoreRecoveryAutosaveHold",true);Check(!(bool)Field(model,"CheckSheetWarningActive")&&(bool)Field(model,"CheckSheetWarningNeedsRecheck"),"Historical state is hidden until fresh automatic check");
   Call(model,"RecordIdleCheckSheetResult",rev,false,true);
   var pending=(bool)model.ResultsPending;var dirty=(bool)model.IsDirty;var userRev=(long)Field(model,"UserChangeRevision");var engine=w.Options.CalculationEngineType;var mode=w.Options.CalculationMode;var history=w.History.Count;
   Static(clock,"RunCheck",(object)model);
   Check((bool)model.ResultsPending==pending&&(bool)model.IsDirty==dirty&&(long)Field(model,"UserChangeRevision")==userRev,"Worksheet watch does not certify full results or manufacture user edits");
   Check(w.Options.CalculationMode==mode&&w.Options.CalculationEngineType==engine&&w.History.Count==history,"Watch preserves calculation settings and workbook history");
   Check((string)Field(settings,"CheckSheetSuspendedFiles")==originalSettings,"Watch does not persist settings or write source XML");
   var checkCell=sheet.Cells["B22"];string checkFormula=checkCell.Formula;var checkValue=checkCell.Value;
   try {checkCell.Formula="=1/0";sheet.Calculate();var invalid=Call(model,"ReadCheckSheetValidation");Check(!string.IsNullOrEmpty((string)Field(invalid,"ValidationError")),"Excel formula errors are not treated as soft balance findings");}
   finally {if(string.IsNullOrEmpty(checkFormula))checkCell.Value=checkValue;else checkCell.Formula=checkFormula;sheet.Calculate();}
   using(var form=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})) {
    form.Opacity=0;form.ShowInTaskbar=false;form.ClientSize=new Size(1700,1000);form.Show();Pump();
    var integrity=app.GetType("Abovo.IdleIntegrityManager");
    Call(model,"RestoreRecoveryAutosaveHold",true);Static(integrity,"RequestOnOpen",(object)model);
    Check((bool)Static(integrity,"ProcessIdle",DateTime.UtcNow.AddSeconds(1),TimeSpan.Zero),"Remembered state schedules an automatic check without the idle wait");
    Check(!(bool)Field(model,"CheckSheetWarningNeedsRecheck")&&!(bool)Field(model,"CheckSheetWarningActive"),"Automatic open check clears a remembered warning on a balanced workbook");
    Static(integrity,"Forget",(object)model);Static(integrity,"Track",(object)model);
    Static(clock,"Configure",true,1,2,false,false);
    Check(!(bool)Static(clock,"ProcessIdle",DateTime.UtcNow.AddMinutes(5),TimeSpan.FromSeconds(30)),"Worksheet watch waits for the configured idle period");
    files.GetField("InternalBIsSaving",F).SetValue(null,true);
    try {Check(!(bool)Static(clock,"ProcessIdle",DateTime.UtcNow.AddMinutes(5),TimeSpan.FromMinutes(3)),"Worksheet watch cannot overlap a save");}finally{files.GetField("InternalBIsSaving",F).SetValue(null,false);}
    Check((bool)Static(clock,"ProcessIdle",DateTime.UtcNow.AddMinutes(5),TimeSpan.FromMinutes(3)),"Due worksheet watch runs at a safe idle boundary");
    Check(!(bool)Static(clock,"ProcessIdle",DateTime.UtcNow.AddMinutes(10),TimeSpan.FromMinutes(3)),"Unchanged workbook revision is not checked repeatedly");Static(clock,"Configure",false,5,2,false,false);
    Call(form,"ShowInterface",(int)model.ModelID,14,false,"None",null,-1);Pump();
    var dit=(Control)Field(form,"ActiveInterface");Call(dit,"BuildSection",1,false,false);Pump();
    var grids=Controls(dit).OfType<DevExpress.XtraGrid.GridControl>().ToList();
    var view=grids.Select(g=>g.MainView).OfType<DevExpress.XtraGrid.Views.Grid.GridView>().First(v=>v.Columns.Count>10);
    for(int i=0;i<3;i++){var c=view.Columns[i];var band=c as DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn;Check(c.Fixed==DevExpress.XtraGrid.Columns.FixedStyle.Left||(band!=null&&band.OwnerBand.Fixed==DevExpress.XtraGrid.Columns.FixedStyle.Left),"Management descriptor "+i+" remains visible during scrolling");}
    var data=Field(view.Tag,"DataSet");
    Call(data,"UpdateLocks");var rows=(Array)Field(data,"DataRows");var points=(Array)Field(rows.GetValue(0),"DataCells");
    for(int i=1;i<3;i++){var point=points.GetValue(i);var source=w.Worksheets[(string)Field(point,"SourceSheet")].Cells[(string)Field(point,"SourceAddress")];Check((Color)Field(point,"BGColor")==source.Fill.BackgroundColor&&(Color)Field(point,"FoColor")==source.Font.Color,"Refreshed descriptor colours come from source cell "+i);}
    Call(form,"ShowInterface",(int)model.ModelID,0,false,"None",null,-1);Pump();dit=(Control)Field(form,"ActiveInterface");
    var company=Controls(dit).First(c=>c.Name.StartsWith("TextBox_"));Check(company.Width>400,"Company editor uses more available width");
    using(var options=(Form)Activator.CreateInstance(app.GetType("Abovo.ApplicationOptionsForm"))){options.ShowInTaskbar=false;options.Opacity=0;options.Show();Pump();var tabs=Controls(options).OfType<DevExpress.XtraTab.XtraTabControl>().First();Check(tabs.TabPages.Cast<DevExpress.XtraTab.XtraTabPage>().Any(p=>p.Text=="Check Sheet trial"),"Separate configurable trial tab exists");options.Close();}
    form.Close();
   }
   Check((string)Field(settings,"CheckSheetSuspendedFiles")==originalSettings,"No check or interface render changes persisted validation settings");
   Call(model,"RecordIdleCheckSheetResult",(long)Field(model,"CalculationRevision"),true,true);
   Check(model.SaveFile(),"Explicit Save succeeds for a soft balance finding");
   Check(((string)Field(settings,"CheckSheetSuspendedFiles")).Contains(copy),"Successful Save persists only the saved file's balance state");
   Call(model,"RecordIdleCheckSheetResult",(long)Field(model,"CalculationRevision"),false,true);
   Check(((string)Field(settings,"CheckSheetSuspendedFiles")).Contains(copy),"Unsaved clearance also remains session-only");
   Check(model.SaveFile(),"User can save the cleared balance state");
   Check((string)Field(settings,"CheckSheetSuspendedFiles")==originalSettings,"Successful Save clears persisted state for this private file only");
  } finally {Static(backup,"PersistCheckSheetPausePath",copy,false,-1);files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});}
  Console.WriteLine("PASS: "+count+" checks");return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
}
