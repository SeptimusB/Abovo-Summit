// Diagnostic only: private AGL copy, native Summit edit/Undo/navigation, no save.
using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Reflection.Emit;
using System.Diagnostics;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;

class CheckSheetFundingDiagnostic297 {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static Assembly app; static dynamic model; static IWorkbook book; static Form host; static Control dit; static int events,checks;
 static object Get(object o,string name){for(var t=o.GetType();t!=null;t=t.BaseType){var f=t.GetField(name,F|BindingFlags.DeclaredOnly);if(f!=null)return f.GetValue(o);var p=t.GetProperty(name,F|BindingFlags.DeclaredOnly);if(p!=null)return p.GetValue(o,null);}throw new MissingMemberException(name);}
 static object Call(object o,string name,params object[] a){return o.GetType().GetMethod(name,F).Invoke(o,a);}
 static object Static(string type,string name,params object[] a){return app.GetType(type).GetMethod(name,F).Invoke(null,a);}
 static void Check(bool ok,string s){if(!ok)throw new Exception(s);Console.WriteLine("PASS "+(++checks)+" "+s);}
 static IEnumerable<Control> Controls(Control c){foreach(Control x in c.Controls){yield return x;foreach(var n in Controls(x))yield return n;}}
 static void Pump(){Application.DoEvents();}
 static void Timed(string label,Action action){var sw=Stopwatch.StartNew();action();Console.WriteLine("TIME "+label+" ms="+sw.ElapsedMilliseconds);}
 static Form QuietHost(Type baseType,object[] arguments){
  var assembly=AppDomain.CurrentDomain.DefineDynamicAssembly(new AssemblyName("CheckDiagnosticQuiet"+Guid.NewGuid().ToString("N")),AssemblyBuilderAccess.Run);
  var type=assembly.DefineDynamicModule("Main").DefineType("QuietCheckHost",TypeAttributes.Public,baseType);
  var signature=new[]{typeof(int),typeof(int),typeof(string)};var constructor=type.DefineConstructor(MethodAttributes.Public,CallingConventions.Standard,signature);var il=constructor.GetILGenerator();
  il.Emit(OpCodes.Ldarg_0);il.Emit(OpCodes.Ldarg_1);il.Emit(OpCodes.Ldarg_2);il.Emit(OpCodes.Ldarg_3);il.Emit(OpCodes.Call,baseType.GetConstructor(signature));il.Emit(OpCodes.Ret);
  var b=typeof(Form).GetProperty("ShowWithoutActivation",F).GetGetMethod(true);var m=type.DefineMethod(b.Name,MethodAttributes.Family|MethodAttributes.Virtual|MethodAttributes.HideBySig|MethodAttributes.SpecialName,typeof(bool),Type.EmptyTypes);il=m.GetILGenerator();il.Emit(OpCodes.Ldc_I4_1);il.Emit(OpCodes.Ret);type.DefineMethodOverride(m,b);
  var cb=baseType.GetProperty("CreateParams",F).GetGetMethod(true);var cm=type.DefineMethod(cb.Name,MethodAttributes.Family|MethodAttributes.Virtual|MethodAttributes.HideBySig|MethodAttributes.SpecialName,typeof(CreateParams),Type.EmptyTypes);il=cm.GetILGenerator();var v=il.DeclareLocal(typeof(CreateParams));
  il.Emit(OpCodes.Ldarg_0);il.Emit(OpCodes.Call,cb);il.Emit(OpCodes.Stloc,v);il.Emit(OpCodes.Ldloc,v);il.Emit(OpCodes.Ldloc,v);il.Emit(OpCodes.Callvirt,typeof(CreateParams).GetProperty("ExStyle").GetGetMethod());il.Emit(OpCodes.Ldc_I4,0x08000000);il.Emit(OpCodes.Or);il.Emit(OpCodes.Callvirt,typeof(CreateParams).GetProperty("ExStyle").GetSetMethod());il.Emit(OpCodes.Ldloc,v);il.Emit(OpCodes.Ret);type.DefineMethodOverride(cm,cb);
  return (Form)Activator.CreateInstance(type.CreateType(),arguments);
 }
 static void Show(int csid){Call(host,"ShowInterface",(int)model.ModelID,csid,false,"None",null,csid==33?1:0);dit=(Control)Get(host,"ActiveInterface");if(csid==0){Call(dit,"BuildSection",2,false,false);((DevExpress.XtraTab.XtraTabControl)Get(dit,"XtraTabControlNewGIT")).SelectedTabPageIndex=2;}else{Call(dit,"BuildSection",1,false,false);((DevExpress.XtraTab.XtraTabControl)Get(dit,"XtraTabControlNewGIT")).SelectedTabPageIndex=1;}Pump();}
 static void State(string stage){
  var result=Call((object)model,"ReadCheckSheetValidation");var issues=(IList)Get(result,"Issues");var overridden=(IList)Get(result,"OverriddenIssues");
  Console.WriteLine("STATE "+stage+" input="+book.Worksheets["Funding Assumptions"].Cells["G82"].Value+" engine="+book.Options.CalculationEngineType+" mode="+book.Options.CalculationMode+" skip="+model.WBCalculationService.DontCalcTDBS+" failures="+issues.Count+" overridden="+overridden.Count+" error="+Get(result,"ValidationError")+" warning="+Get((object)model,"CheckSheetWarningActive")+" heading="+(host==null?"none":Get(Get(host,"CheckSheetButton"),"Visibility"))+" events="+events+" pending="+model.ResultsPending+" revision="+Get((object)model,"CalculationRevision")+" nav="+model.WBCalcEngine.NavigationCalculationCurrent);
  var sheet=book.Worksheets["Check Sheet"];foreach(object issue in issues){int row=(int)Get(issue,"CheckRow");Console.WriteLine("ISSUE row="+row+" label="+Get(issue,"Label")+" status="+Get(issue,"Status")+" raw="+sheet.Cells["B"+row].Value+" net="+sheet.Cells["D"+row].Value+" formula="+sheet.Cells["B"+row].FormulaInvariant);}
  if(dit!=null)foreach(var grid in Controls(dit).OfType<GridControl>().Where(g=>g.GetType().Name=="ReadOnlyMappedTableGrid")){
   var view=(GridView)grid.MainView;int mismatch=0;for(int i=0;i<view.DataRowCount;i++){var rows=(IList)Get(grid,"sourceRows");int row=(int)rows[i];string ui=view.GetRowCellDisplayText(i,view.Columns["C4"]);string wb=sheet.Cells[row,4].DisplayText;if(ui!=wb){mismatch++;Console.WriteLine("UISTALE row="+(row+1)+" ui="+ui+" workbook="+wb);}}Console.WriteLine("UI rows="+view.DataRowCount+" mismatch="+mismatch);
  }
 }
 static void Change(double value){dynamic change=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));change.ModelID=(int)model.ModelID;change.WSName="Funding Assumptions";change.CellAddress="G82";change.ChangedValue=value;change.DataFormat="N";change.Description="Private diagnostic G82";dynamic result=model.ChangeManager.ProcessChange(change);Check((bool)result.BSuccess,"G82 edit accepted: "+value);}
 static void Undo(){dynamic result=model.ChangeManager.Undo();Check((bool)result.BSuccess,"Undo accepted");}
 static void Engine(string label,CalculationEngineType engine,bool skip,Action calculate){var old=book.Options.CalculationEngineType;bool oldSkip=model.WBCalculationService.DontCalcTDBS;try{book.Options.CalculationEngineType=engine;model.WBCalculationService.DontCalcTDBS=skip;Timed(label,calculate);State(label);}finally{book.Options.CalculationEngineType=old;model.WBCalculationService.DontCalcTDBS=oldSkip;}}
 [STAThread] static int Main(string[] args){try{
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));Static("Abovo.AbovoAppCls","Initialise");Static("Abovo.FileManager","Initialise",new object[]{null});
  string copy=Path.Combine(args[1],"private-checksheet-funding.xlsb");File.Copy(args[2],copy);var files=app.GetType("Abovo.FileManager");var open=files.GetMethod("OpenModel");dynamic opened=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});Check(!(bool)opened.BError,"Private workbook opened");model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)opened.IntegerReturn);book=model.WB;
  Static("Abovo.CheckSheetWatch","Configure",false,5,2,false,false);
  EventInfo warning=model.GetType().GetEvent("CheckSheetStatusChanged",F);warning.GetAddMethod(true).Invoke((object)model,new object[]{new EventHandler((s,e)=>{events++;Console.WriteLine("STATUS_EVENT count="+events+" warning="+Get((object)model,"CheckSheetWarningActive"));})});
  Check(book.Worksheets["Funding Assumptions"].Cells["G82"].Value.NumericValue==2700,"AGL G82 is the reported original2700");State("loaded");
  host=QuietHost(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,-1,"Normal"});host.Opacity=0;host.ShowInTaskbar=false;host.ClientSize=new Size(1800,1100);host.Show();Pump();
  Timed("open funding",()=>Show(33));State("funding baseline");Timed("edit3000",()=>Change(3000));State("edited funding");
  Timed("navigate Check Sheet",()=>Show(0));State("broken Check Sheet visible");Timed("history Undo while Check Sheet visible",Undo);State("undo without navigation");
  Timed("explicit grid RefreshData only",()=>Call(dit,"RefreshData",false));State("undo after display-only refresh");
  Timed("navigate away and back",()=>{Show(33);Show(0);});State("undo after reactivation");
  Engine("sheet calculate skipped",CalculationEngineType.ChainBased,true,()=>book.Worksheets["Check Sheet"].Calculate());
  Engine("deferred worksheets only",CalculationEngineType.ChainBased,true,()=>model.WBCalculationService.CalculateDeferredWorksheets(book));
  Engine("chain incremental all",CalculationEngineType.ChainBased,false,()=>book.Calculate());
  Engine("chain full all",CalculationEngineType.ChainBased,false,()=>book.CalculateFull());
  Engine("recursive full all",CalculationEngineType.Recursive,false,()=>book.CalculateFull());
  Timed("grid refresh after full",()=>Call(dit,"RefreshData",false));State("balanced after full display refresh");
  Timed("second edit3000",()=>Change(3000));Timed("watch break",()=>Static("Abovo.CheckSheetWatch","RunCheck",(object)model));State("watch publishes broken");
  Timed("second undo",Undo);State("second undo before watch");Timed("watch clear",()=>Static("Abovo.CheckSheetWatch","RunCheck",(object)model));State("watch publishes cleared");
  Check(book.Worksheets["Funding Assumptions"].Cells["G82"].Value.NumericValue==2700,"Undo restored source input");
  host.Dispose();model.ModelSpreadsheetControl.Dispose();Console.WriteLine("DIAGNOSTIC COMPLETE assertions="+checks+" no workbook saved; original untouched");return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
