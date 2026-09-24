// Native, private-copy regression for item 80. Never saves the source or activates a window.
// Run with -UseApplicationConfig -MatchApplicationAddressSpace so the host matches Summit.
using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Reflection;
using System.Reflection.Emit;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraEditors;

class CheckSheetCoherence298 {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static Assembly app; static dynamic model; static IWorkbook book; static Form host; static Control dit;
 static GridControl grid; static GridView view; static int assertions,statusEvents,valueEvents;
 static readonly List<string> publications=new List<string>();
 static bool LargeAddressAware(string path) {using(var reader=new BinaryReader(File.OpenRead(path))){reader.BaseStream.Position=0x3c;int offset=reader.ReadInt32();reader.BaseStream.Position=offset+22;return (reader.ReadUInt16()&0x20)!=0;}}
 [DllImport("user32.dll")] static extern int GetGuiResources(IntPtr process,int flags);
 static void Resources(string stage) {
  using(var p=Process.GetCurrentProcess()){p.Refresh();Console.WriteLine("RESOURCES "+stage+" private="+p.PrivateMemorySize64+" virtual="+p.VirtualMemorySize64+" working="+p.WorkingSet64+" peakWorking="+p.PeakWorkingSet64+" managed="+GC.GetTotalMemory(false)+" gdi="+GetGuiResources(p.Handle,0)+" user="+GetGuiResources(p.Handle,1));}
  if(host!=null){var children=Children(host).ToArray();var largest=children.OrderByDescending(c=>(long)c.Width*c.Height).FirstOrDefault();Console.WriteLine("BOUNDS "+stage+" host="+host.ClientSize+" dit="+(dit==null?"none":dit.Size.ToString())+" maxChildWidth="+(children.Length==0?0:children.Max(c=>c.Width))+" maxChildHeight="+(children.Length==0?0:children.Max(c=>c.Height))+" largest="+(largest==null?"none":largest.GetType().Name+largest.Size));}
 }
 static object Get(object o,string name) { for(var t=o.GetType();t!=null;t=t.BaseType) { var f=t.GetField(name,F|BindingFlags.DeclaredOnly); if(f!=null)return f.GetValue(o);var p=t.GetProperty(name,F|BindingFlags.DeclaredOnly);if(p!=null)return p.GetValue(o,null); }throw new MissingMemberException(name); }
 static object Call(object o,string name,params object[] args) { return o.GetType().GetMethod(name,F).Invoke(o,args); }
 static object Static(string type,string name,params object[] args) { return app.GetType(type).GetMethod(name,F).Invoke(null,args); }
 static void Check(bool valid,string text) { if(!valid)throw new Exception(text);Console.WriteLine("PASS "+(++assertions)+" "+text); }
 static long Number(string name) { return Convert.ToInt64(Get((object)model,name)); }
 static long FullCount() { return Convert.ToInt64(Static("Abovo.CheckSheetWatch","GetFullCalculationCount",(object)model)); }
 static void Pump() { Application.DoEvents(); }
 static IEnumerable<Control> Children(Control c) { foreach(Control child in c.Controls) { yield return child;foreach(var nested in Children(child))yield return nested; } }
 static void Subscribe(string eventName,EventHandler handler,bool add) { var e=model.GetType().GetEvent(eventName,F);(add?e.GetAddMethod(true):e.GetRemoveMethod(true)).Invoke((object)model,new object[]{handler}); }
 static Form QuietHost(Type baseType) {
  var assembly=AppDomain.CurrentDomain.DefineDynamicAssembly(new AssemblyName("QuietCheckSheet"+Guid.NewGuid().ToString("N")),AssemblyBuilderAccess.Run);
  var type=assembly.DefineDynamicModule("Main").DefineType("QuietHost",TypeAttributes.Public,baseType);
  var signature=new[]{typeof(int),typeof(int),typeof(string)};
  var ctor=type.DefineConstructor(MethodAttributes.Public,CallingConventions.Standard,signature);var il=ctor.GetILGenerator();
  il.Emit(OpCodes.Ldarg_0);il.Emit(OpCodes.Ldarg_1);il.Emit(OpCodes.Ldarg_2);il.Emit(OpCodes.Ldarg_3);il.Emit(OpCodes.Call,baseType.GetConstructor(signature));il.Emit(OpCodes.Ret);
  var original=typeof(Form).GetProperty("ShowWithoutActivation",F).GetGetMethod(true);
  var getter=type.DefineMethod(original.Name,MethodAttributes.Family|MethodAttributes.Virtual|MethodAttributes.HideBySig|MethodAttributes.SpecialName,typeof(bool),Type.EmptyTypes);
  il=getter.GetILGenerator();il.Emit(OpCodes.Ldc_I4_1);il.Emit(OpCodes.Ret);type.DefineMethodOverride(getter,original);
  original=baseType.GetProperty("CreateParams",F).GetGetMethod(true);
  getter=type.DefineMethod(original.Name,MethodAttributes.Family|MethodAttributes.Virtual|MethodAttributes.HideBySig|MethodAttributes.SpecialName,typeof(CreateParams),Type.EmptyTypes);
  il=getter.GetILGenerator();var local=il.DeclareLocal(typeof(CreateParams));il.Emit(OpCodes.Ldarg_0);il.Emit(OpCodes.Call,original);il.Emit(OpCodes.Stloc,local);il.Emit(OpCodes.Ldloc,local);il.Emit(OpCodes.Ldloc,local);il.Emit(OpCodes.Callvirt,typeof(CreateParams).GetProperty("ExStyle").GetGetMethod());il.Emit(OpCodes.Ldc_I4,0x08000000);il.Emit(OpCodes.Or);il.Emit(OpCodes.Callvirt,typeof(CreateParams).GetProperty("ExStyle").GetSetMethod());il.Emit(OpCodes.Ldloc,local);il.Emit(OpCodes.Ret);type.DefineMethodOverride(getter,original);
  return (Form)Activator.CreateInstance(type.CreateType(),new object[]{(int)model.ModelID,-1,"Normal"});
 }
 static void Show(int csid) {
  var timer=Stopwatch.StartNew();long before=FullCount();
  Resources("before-navigation-"+csid);
  Call(host,"ShowInterface",(int)model.ModelID,csid,false,"None",null,csid==33?1:0);
  dit=(Control)Get(host,"ActiveInterface");int section=csid==33?1:2;
  // Use the native lazy-tab lifecycle, never force a destructive private rebuild.
  ((DevExpress.XtraTab.XtraTabControl)Get(dit,"XtraTabControlNewGIT")).SelectedTabPageIndex=section;Pump();
  if(csid==0){grid=Children(dit).OfType<GridControl>().First(g=>g.GetType().Name=="ReadOnlyMappedTableGrid");view=(GridView)grid.MainView;}
  Console.WriteLine("NAVIGATION csid="+csid+" ms="+timer.ElapsedMilliseconds+" additionalFull="+(FullCount()-before)+" revision="+Number("CalculationRevision"));
 }
 static void Change(string sheet,string address,object value,string format) {
  var timer=Stopwatch.StartNew();long before=FullCount();
  dynamic edit=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));edit.ModelID=(int)model.ModelID;edit.WSName=sheet;edit.CellAddress=address;edit.ChangedValue=value;edit.DataFormat=format;edit.Description="Private Check Sheet coherence regression";
  dynamic result=model.ChangeManager.ProcessChange(edit);Check((bool)result.BSuccess,"Journalled "+sheet+"!"+address+" edit accepted");Pump();
  Console.WriteLine("EDIT "+sheet+"!"+address+" ms="+timer.ElapsedMilliseconds+" additionalFull="+(FullCount()-before));
 }
 static int[] FailureRows() {
  var result=Call((object)model,"ReadCheckSheetValidation");var error=Convert.ToString(Get(result,"ValidationError"));if(!String.IsNullOrEmpty(error))throw new Exception("Validation read failed: "+error);
  return ((IList)Get(result,"Issues")).Cast<object>().Select(i=>Convert.ToInt32(Get(i,"CheckRow"))).OrderBy(i=>i).ToArray();
 }
 static int GridRow(int sheetRow) { var rows=(IList)Get(grid,"sourceRows");for(int i=0;i<rows.Count;i++)if((int)rows[i]==sheetRow-1)return i;throw new Exception("Unmapped check row "+sheetRow); }
 static int ProjectedForeground(string address) {
  Color colour;bool conditional=((Dictionary<string,Color>)Get(grid,"conditionalForeground")).TryGetValue(address,out colour);
  Color source=book.Worksheets["Check Sheet"].Cells[address].Font.Color;if(!conditional)colour=source;
  Console.WriteLine("FOREGROUND "+address+" conditional="+conditional+" projected="+colour.ToArgb()+" source="+source.ToArgb());return colour.ToArgb();
 }
 static void Coherent(string stage,params int[] expected) {
  Resources(stage);
  var actual=FailureRows();Console.WriteLine("STATE "+stage+" rows="+String.Join(",",actual)+" calcRevision="+Number("CalculationRevision")+" accepted="+Number("LastAcceptedCheckSheetRevision")+" full="+FullCount()+" values="+valueEvents+" status="+statusEvents);
  Check(actual.SequenceEqual(expected),stage+": exact fresh failures "+String.Join(",",expected));
  Check((bool)Get((object)model,"CheckSheetWarningActive")== (expected.Length>0),stage+": public warning agrees with workbook");
  Check((Get(Get(host,"CheckSheetButton"),"Visibility").ToString()=="Always")== (expected.Length>0),stage+": company-heading visibility agrees");
  Check(Number("LastAcceptedCheckSheetRevision")==Number("CalculationRevision"),stage+": final committed revision was accepted");
  var rows=(IList)Get(grid,"sourceRows");var sheet=book.Worksheets["Check Sheet"];int mismatches=0;
  for(int row=0;row<rows.Count;row++)for(int column=0;column<=5;column++)if(view.GetRowCellDisplayText(row,view.Columns["C"+column])!=sheet.Cells[(int)rows[row],column].DisplayText)mismatches++;
  Check(mismatches==0,stage+": all mapped labels, figures, override text and statuses match worksheet");
  Check(book.Worksheets["Check Sheet"].Cells["E37"].DisplayText=="OK",stage+": no transient TDB Cashflow failure remains");
 }
 static Dictionary<string,object> Flags() {
  return new[]{"IsDirty","UserChangeRevision","CalculationRevision","ResultsPending","NeedsFullRebuild","NeedsSaveRebuild","NeedsFormulaPreflight","DeferredSaveResultsPending"}.ToDictionary(name=>name,name=>Get((object)model,name));
 }
 static void SameFlags(Dictionary<string,object> before,string stage) { Check(before.All(pair=>Object.Equals(pair.Value,Get((object)model,pair.Key))),stage+": dirty, user/calculation revisions and pending/rebuild/preflight flags unchanged"); }
 static bool Publish(long revision,bool failure,out bool changed) { var arguments=new object[]{revision,failure,false,true};bool accepted=(bool)Call((object)model,"TryRecordCheckSheetResult",arguments);changed=(bool)arguments[2];return accepted; }
 static void Override(string value) { dynamic result=Call(grid,"PostChoice",book.Worksheets["Check Sheet"].Cells["C23"],value);Check((bool)result.BSuccess,"Native mapped override accepts literal "+value);Pump(); }
 static void Watch() {
  var flags=Flags();var engine=book.Options.CalculationEngineType;var mode=book.Options.CalculationMode;bool skip=model.WBCalculationService.DontCalcTDBS;int history=book.History.Count;
  Static("Abovo.CheckSheetWatch","RunCheck",(object)model);Pump();SameFlags(flags,"Watcher");
  Check(book.Options.CalculationEngineType==engine&&book.Options.CalculationMode==mode&&(bool)model.WBCalculationService.DontCalcTDBS==skip&&book.History.Count==history,"Watcher restores calculation engine/mode/deferral and leaves native history intact");
 }
 [STAThread] static int Main(string[] args) { try {
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string path=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(path)?Assembly.LoadFrom(path):null;};
  app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));Console.WriteLine("ASSEMBLY "+app.Location+" "+FileVersionInfo.GetVersionInfo(app.Location).FileVersion);
  Check(!Environment.Is64BitProcess,"Fixture runs32-bit, matching the reported Summit execution mode");
  bool appLaa=LargeAddressAware(app.Location),fixtureLaa=LargeAddressAware(Assembly.GetExecutingAssembly().Location);Console.WriteLine("ADDRESS_SPACE targetLAA="+appLaa+" fixtureLAA="+fixtureLaa);
  Check(appLaa==fixtureLaa,"Fixture matches the application large-address-aware PE flag");
  Check(Convert.ToDecimal(app.GetType("Abovo.AbovoAppCls").GetProperty("DecVersionNumber",F).GetValue(null,null))==2.98m,"Loaded Summit test release2.98");
  Static("Abovo.AbovoAppCls","Initialise");Static("Abovo.FileManager","Initialise",new object[]{null});
  string copy=Path.Combine(args[1],"private-checksheet-coherence.xlsb");File.Copy(args[2],copy);
  var files=app.GetType("Abovo.FileManager");var open=files.GetMethod("OpenModel");dynamic opened=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!(bool)opened.BError,"Private AGL workbook opens");model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)opened.IntegerReturn);book=model.WB;
  Static("Abovo.CheckSheetWatch","Configure",false,5,2,false,false);
  var settings=app.GetType("Abovo.RecoveryBackupSettings").GetField("Instance",F).GetValue(null);string persisted=(string)Get(settings,"CheckSheetSuspendedFiles");
  Subscribe("CheckSheetStatusChanged",(s,e)=>{statusEvents++;},true);
  Subscribe("CheckSheetValuesRefreshed",(s,e)=>{valueEvents++;var rows=String.Join(",",FailureRows());publications.Add(rows);Console.WriteLine("PUBLICATION "+valueEvents+" revision="+Number("LastAcceptedCheckSheetRevision")+" rows="+rows);},true);
  Check(book.Worksheets["Funding Assumptions"].Cells["G82"].Value.NumericValue==2700,"Actual AGL input starts at2700");
  var originalEngine=book.Options.CalculationEngineType;var originalMode=book.Options.CalculationMode;bool originalSkip=model.WBCalculationService.DontCalcTDBS;
  host=QuietHost(app.GetType("GroupInterfaceTemplate"));host.Opacity=0;host.ShowInTaskbar=false;host.ClientSize=new Size(1800,1100);host.Show();Pump();Show(33);
  long before=FullCount();Change("Funding Assumptions","G82",3000d,"N");Check(FullCount()==before,"Ordinary Funding edit does not trigger a new full Check Sheet calculation while hidden");
  Show(0);Coherent("warm-checksheet",23,33,39);Check(FullCount()==before+1,"First requested Check Sheet display uses one full calculation");
  var failedColours=(Dictionary<string,Color>)Get(grid,"conditionalForeground");
  var failedForeground=new[]{"E23","B23"}.ToDictionary(address=>address,address=>ProjectedForeground(address));
  Check(failedForeground.Keys.Any(address=>failedColours.ContainsKey(address)),"Failed row23 has workbook-rule conditional foreground in the mapped cache");
  before=FullCount();dynamic result=model.ChangeManager.Undo();Check((bool)result.BSuccess,"Undo while Check Sheet visible accepted");Pump();Coherent("undo-visible");Check(FullCount()==before+1,"Undo performs one final-revision full refresh, not provisional duplicate calculations");
  Check(failedForeground.Select(pair=>ProjectedForeground(pair.Key)!=pair.Value).Any(changedColour=>changedColour),"Undo refresh changes the failed row23 foreground back to its workbook-resolved appearance");
  Check(book.Worksheets["Funding Assumptions"].Cells["G82"].Value.NumericValue==2700,"Undo restores source input2700");
  before=FullCount();result=model.ChangeManager.Redo();Check((bool)result.BSuccess,"Redo while Check Sheet visible accepted");Pump();Coherent("redo-visible",23,33,39);Check(FullCount()==before+1,"Redo performs one final-revision full refresh");
  Check(failedForeground.Select(pair=>ProjectedForeground(pair.Key)==pair.Value).All(sameColour=>sameColour),"Redo refresh restores the same workbook-rule failed foreground");
  // A first Funding rebind after history currently invalidates calculation. Do not
  // mislabel that changed revision as unchanged or bypass its freshness check.
  long priorRevision=Number("CalculationRevision");before=FullCount();Show(33);Show(0);
  Console.WriteLine("POST_HISTORY_REBIND revision="+priorRevision+"->"+Number("CalculationRevision")+" additionalFull="+(FullCount()-before));
  Check(Number("LastAcceptedCheckSheetRevision")==Number("CalculationRevision"),"Any revision-changing Funding rebind is freshly checked before reuse");
  before=FullCount();priorRevision=Number("CalculationRevision");Show(33);Show(0);Coherent("repeat-navigation",23,33,39);
  Check(Number("CalculationRevision")==priorRevision&&FullCount()==before,"Repeated truly unchanged-revision navigation does not request another full calculation");
  var existingCheckSheet=dit;Show(33);before=FullCount();Change("Funding Assumptions","G82",3100d,"N");Check(FullCount()==before,"Second Funding edit while Check Sheet hidden remains on the ordinary edit path");
  // Return to the cached DIT, without BuildSection or assigning its selected tab.
  Call(host,"ShowInterface",(int)model.ModelID,0,false,"None",null,0);dit=(Control)Get(host,"ActiveInterface");Pump();
  grid=Children(dit).OfType<GridControl>().First(g=>g.GetType().Name=="ReadOnlyMappedTableGrid");view=(GridView)grid.MainView;
  Check(Object.ReferenceEquals(existingCheckSheet,dit)&&((DevExpress.XtraTab.XtraTabControl)Get(dit,"XtraTabControlNewGIT")).SelectedTabPageIndex==2,"Warm return retains the already-created Check Sheet and selected tab");
  Coherent("warm-reactivated",23,33,39);Check(FullCount()==before+1,"Warm cached-DIT reactivation checks its new revision exactly once");

  // Corrupt only the UI projection, never workbook data, to prove same-status watcher refresh.
  var values=(DataTable)Get(grid,"values");values.Rows[GridRow(23)]["C4"]="fixture stale";
  int priorValues=valueEvents,priorStatus=statusEvents;Watch();Check(valueEvents==priorValues+1&&statusEvents==priorStatus,"Unchanged failed watcher result emits values event without a spurious status transition");Coherent("watcher-visible",23,33,39);

  // Native pending Yes/No editor: no activation, mouse movement or keyboard injection.
  int overrideRow=GridRow(23);var overrideColumn=view.Columns["C2"];view.FocusedRowHandle=overrideRow;view.FocusedColumn=overrideColumn;view.MakeRowVisible(overrideRow);view.MakeColumnVisible(overrideColumn);view.ShowEditor();Pump();
  BaseEdit pending=view.ActiveEditor;Check(pending!=null,"Actual mapped Yes/No editor opens without activating the host");pending.EditValue="Yes";
  string sourceOverride=book.Worksheets["Check Sheet"].Cells["C23"].Value.TextValue;var pendingFlags=Flags();int top=view.TopRowIndex,left=view.LeftCoord;var widths=view.Columns.Cast<DevExpress.XtraGrid.Columns.GridColumn>().Select(c=>c.Width).ToArray();
  view.ClearSelection();view.SelectCell(overrideRow,overrideColumn);values.Rows[overrideRow]["C4"]="fixture pending";
  bool changed;before=FullCount();Check(Publish(Number("CalculationRevision"),true,out changed)&&!changed,"Current already-failed result is accepted separately from status transition");Pump();
  Check(Object.ReferenceEquals(pending,view.ActiveEditor)&&Convert.ToString(pending.EditValue)=="Yes"&&book.Worksheets["Check Sheet"].Cells["C23"].Value.TextValue==sourceOverride,"Values notification neither commits nor discards pending Yes input");
  Check((bool)Get(grid,"checkSheetRefreshPending")&&view.GetRowCellDisplayText(overrideRow,view.Columns["C4"])=="fixture pending","Notification queues a refresh while the override editor remains active");SameFlags(pendingFlags,"Pending publication");
  view.HideEditor();Pump();Call(grid,"TryRefreshCheckSheetValues");Pump();
  Check(view.ActiveEditor==null&&book.Worksheets["Check Sheet"].Cells["C23"].Value.TextValue==sourceOverride,"Cancelling the pending editor leaves workbook override unchanged");
  Check(view.TopRowIndex==top&&view.LeftCoord==left&&view.FocusedRowHandle==overrideRow&&view.FocusedColumn==overrideColumn&&view.GetSelectedCells().Any(c=>c.RowHandle==overrideRow&&c.Column==overrideColumn)&&view.Columns.Cast<DevExpress.XtraGrid.Columns.GridColumn>().Select(c=>c.Width).SequenceEqual(widths),"Deferred refresh preserves focus, viewport, selection and column widths");
  Check(FullCount()==before,"Publication/deferred repaint does not calculate");Coherent("pending-cancelled",23,33,39);

  long accepted=Number("LastAcceptedCheckSheetRevision");priorValues=valueEvents;priorStatus=statusEvents;var rejectionFlags=Flags();
  Check(!Publish(Number("CalculationRevision")-1,false,out changed)&&!changed,"Wrong-revision result is rejected, not reported as an unchanged acceptance");
  Check(valueEvents==priorValues&&statusEvents==priorStatus&&Number("LastAcceptedCheckSheetRevision")==accepted&&(bool)Get((object)model,"CheckSheetWarningActive"),"Rejected stale balanced result cannot clear warning or emit refresh");SameFlags(rejectionFlags,"Rejected publication");
  int sibling=0;EventHandler throws=(s,e)=>{throw new InvalidOperationException("Intentional private disposed-view simulation");};EventHandler survives=(s,e)=>{sibling++;};Subscribe("CheckSheetValuesRefreshed",throws,true);Subscribe("CheckSheetValuesRefreshed",survives,true);
  try { Check(Publish(Number("CalculationRevision"),true,out changed)&&sibling==1,"Failing presentation subscriber cannot escape or suppress sibling notification"); }finally {Subscribe("CheckSheetValuesRefreshed",throws,false);Subscribe("CheckSheetValuesRefreshed",survives,false);}Pump();
  bool reentrantAccepted=true;int reentrantCalls=0;EventHandler reentrant=(s,e)=>{bool innerChanged;reentrantCalls++;reentrantAccepted=Publish(Number("CalculationRevision"),false,out innerChanged);};Subscribe("CheckSheetValuesRefreshed",reentrant,true);
  try {Check(Publish(Number("CalculationRevision"),true,out changed)&&reentrantCalls==1&&!reentrantAccepted,"Reentrant publication is rejected without recursive events or false clearance");}finally {Subscribe("CheckSheetValuesRefreshed",reentrant,false);}Pump();
  priorValues=valueEvents;priorStatus=statusEvents;accepted=Number("LastAcceptedCheckSheetRevision");bool priorClosing=model.IsClosing;
  try {model.IsClosing=true;Check(!Publish(Number("CalculationRevision"),false,out changed)&&!changed,"Closing model rejects publication");}finally {model.IsClosing=priorClosing;}
  Check(valueEvents==priorValues&&statusEvents==priorStatus&&Number("LastAcceptedCheckSheetRevision")==accepted,"Rejected closing-model result leaves events and accepted revision unchanged");

  string numberFormat=book.Worksheets["Check Sheet"].Cells["C23"].NumberFormat;Override("Yes");Coherent("override-yes",33,39);
  Check(book.Worksheets["Check Sheet"].Cells["C23"].Value.IsText&&book.Worksheets["Check Sheet"].Cells["C23"].Value.TextValue=="Yes"&&view.GetRowCellDisplayText(GridRow(23),view.Columns["C2"])=="Yes","Override stays literal Yes, not1");
  Change("Check Sheet","C23",null,"S");Coherent("override-cleared",23,33,39);Check(book.Worksheets["Check Sheet"].Cells["C23"].Value.IsEmpty&&view.GetRowCellDisplayText(GridRow(23),view.Columns["C2"])=="","Cleared override remains blank, not0");
  result=model.ChangeManager.Undo();Check((bool)result.BSuccess,"Override-clear Undo accepted");Pump();Coherent("override-undo",33,39);Check(book.Worksheets["Check Sheet"].Cells["C23"].Value.TextValue=="Yes","Undo restores literal Yes");
  result=model.ChangeManager.Redo();Check((bool)result.BSuccess,"Override-clear Redo accepted");Pump();Coherent("override-redo",23,33,39);Override("No");Coherent("override-no",23,33,39);
  Check(book.Worksheets["Check Sheet"].Cells["C23"].Value.TextValue=="No"&&book.Worksheets["Check Sheet"].Cells["C23"].NumberFormat==numberFormat,"No text and original workbook number format retained");
  Check(!publications.Any(rows=>rows.Split(',').Contains("37")),"No provisional TDB Cashflow failure was published");
  Check(book.Options.CalculationEngineType==originalEngine&&book.Options.CalculationMode==originalMode&&(bool)model.WBCalculationService.DontCalcTDBS==originalSkip,"Navigation, history and overrides restore entry calculation settings");
  Check((string)Get(settings,"CheckSheetSuspendedFiles")==persisted,"Unsaved private checks do not persist a recovery warning");
  host.Dispose();model.ModelSpreadsheetControl.Dispose();Console.WriteLine("PASS "+assertions+" assertions; original never saved; no global input");return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;} }
}
