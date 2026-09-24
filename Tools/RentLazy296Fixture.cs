// Native Rent lazy-tab regression. Opens only a private workbook copy; never saves.
using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraTab;

class RentLazy296Fixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static int checks;
 [DllImport("user32.dll")]static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]static extern bool SetForegroundWindow(IntPtr window);
 [DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
 static object Read(object o,string n){for(var t=o.GetType();t!=null;t=t.BaseType){var f=t.GetField(n,F|BindingFlags.DeclaredOnly);if(f!=null)return f.GetValue(o);var p=t.GetProperty(n,F|BindingFlags.DeclaredOnly);if(p!=null)return p.GetValue(o,null);}throw new MissingMemberException(o.GetType().FullName,n);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static IEnumerable<Control> Controls(Control c){foreach(Control child in c.Controls){yield return child;foreach(var inner in Controls(child))yield return inner;}}
 static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("PASS "+(++checks)+" "+text);}
 static void Pump(int ms=120){var until=DateTime.UtcNow.AddMilliseconds(ms);do{Application.DoEvents();Thread.Sleep(10);}while(DateTime.UtcNow<until);}
 static void HeaderEligibility(object dit,GridControl grid,BandedGridView view){
  var source=grid.DataSource;int ds=Convert.ToInt32(Read(Read(source,"UBSTag"),"DSIndex"));var data=((IList)Read(Read(dit,"DataPres"),"DataSets"))[ds];
  var rows=(IList)Read(data,"DataRows");int headers=0;
  foreach(BandedGridColumn column in view.Columns){
   if(column.Tag==null||!Convert.ToBoolean(Read(column.Tag,"HasIncolumnEditor")))continue;headers++;
   var header=Read(column.Tag,"InColumnEditorCombo")??Read(column.Tag,"InColumnEditorDate");var helper=Read(header,"InPlaceColumnHelper");var value=Read(helper,"EditValue");
   int index=Convert.ToInt32(Call(dit,"GetGridColumnIndex",column)),eligible=0,locked=0;
   for(int i=0;i<rows.Count;i++){if((bool)Call(dit,"CanPasteToDataPoint",data,i,index))eligible++;var point=((IList)Read(rows[i],"DataCells"))[index];if(point!=null&&Convert.ToBoolean(Read(point,"IsLocked")))locked++;}
   Console.WriteLine("ELIGIBILITY ds="+ds+" field="+column.FieldName+" absolute="+column.AbsoluteIndex+" header="+(value??"<null>")+" visible="+column.Visible+" eligible="+eligible+" locked="+locked+" rows="+rows.Count);
  }
  Check(headers>0,"Read-only eligibility inventory includes actual workbook defining headers");
 }
 static Form QuietHost(Type baseType,object[] args){
  var assembly=AppDomain.CurrentDomain.DefineDynamicAssembly(new AssemblyName("RentQuietHost"+Guid.NewGuid().ToString("N")),AssemblyBuilderAccess.Run);
  var type=assembly.DefineDynamicModule("Main").DefineType("RentQuietHost",TypeAttributes.Public,baseType);
  var signature=new[]{typeof(int),typeof(int),typeof(string)};var ctor=type.DefineConstructor(MethodAttributes.Public,CallingConventions.Standard,signature);var il=ctor.GetILGenerator();
  il.Emit(OpCodes.Ldarg_0);il.Emit(OpCodes.Ldarg_1);il.Emit(OpCodes.Ldarg_2);il.Emit(OpCodes.Ldarg_3);il.Emit(OpCodes.Call,baseType.GetConstructor(signature));il.Emit(OpCodes.Ret);
  var showBase=typeof(Form).GetProperty("ShowWithoutActivation",F).GetGetMethod(true);var show=type.DefineMethod(showBase.Name,MethodAttributes.Family|MethodAttributes.Virtual|MethodAttributes.HideBySig|MethodAttributes.SpecialName,typeof(bool),Type.EmptyTypes);il=show.GetILGenerator();il.Emit(OpCodes.Ldc_I4_1);il.Emit(OpCodes.Ret);type.DefineMethodOverride(show,showBase);
  var cpBase=baseType.GetProperty("CreateParams",F).GetGetMethod(true);var cp=type.DefineMethod(cpBase.Name,MethodAttributes.Family|MethodAttributes.Virtual|MethodAttributes.HideBySig|MethodAttributes.SpecialName,typeof(CreateParams),Type.EmptyTypes);il=cp.GetILGenerator();var value=il.DeclareLocal(typeof(CreateParams));
  il.Emit(OpCodes.Ldarg_0);il.Emit(OpCodes.Call,cpBase);il.Emit(OpCodes.Stloc,value);il.Emit(OpCodes.Ldloc,value);il.Emit(OpCodes.Ldloc,value);il.Emit(OpCodes.Callvirt,typeof(CreateParams).GetProperty("ExStyle").GetGetMethod());il.Emit(OpCodes.Ldc_I4,0x08000000);il.Emit(OpCodes.Or);il.Emit(OpCodes.Callvirt,typeof(CreateParams).GetProperty("ExStyle").GetSetMethod());il.Emit(OpCodes.Ldloc,value);il.Emit(OpCodes.Ret);type.DefineMethodOverride(cp,cpBase);
  return (Form)Activator.CreateInstance(type.CreateType(),args);
 }
 [STAThread]static int Main(string[] args){var previous=GetForegroundWindow();try{
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  var copy=Path.Combine(args[1],"private-rent-lazy296.xlsb");File.Copy(args[2],copy);var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!loaded.BError,"Private Rent model opens");dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)loaded.IntegerReturn);long revision=Convert.ToInt64(Read((object)model,"UserChangeRevision"));bool dirty=(bool)model.IsDirty;
  using(var host=QuietHost(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})){
   host.Opacity=0;host.ShowInTaskbar=false;host.StartPosition=FormStartPosition.Manual;host.Location=new Point(20,20);host.ClientSize=new Size(1800,1100);host.Show();Pump(300);
   using(var graphics=host.CreateGraphics())Console.WriteLine("DPI actual host="+host.DeviceDpi+" graphics="+graphics.DpiX+"/"+graphics.DpiY);
   Call(host,"ShowInterface",(int)model.ModelID,2,false,"None",null,0);Pump();var dit=(Control)Read(host,"ActiveInterface");Check(Convert.ToInt32(Read(dit,"CSID"))==2,"Rent Assumptions opens");var tabs=(XtraTabControl)Read(dit,"XtraTabControlNewGIT");Check(tabs.TabPages.Count>=4,"Rent has its expected lazy tabs");int banded=0;
   for(int i=0;i<tabs.TabPages.Count;i++){
    if(!tabs.TabPages[i].PageVisible||!tabs.TabPages[i].PageEnabled){Console.WriteLine("SKIP non-navigable Rent tab "+i+": "+tabs.TabPages[i].Text);continue;}
    tabs.SelectedTabPageIndex=i;Pump(200);Check(tabs.SelectedTabPageIndex==i&&tabs.SelectedTabPage.Controls.Count>0,"Rent lazy tab builds: "+tabs.SelectedTabPage.Text);
    foreach(var grid in Controls(tabs.SelectedTabPage).OfType<GridControl>()){
     var view=grid.MainView as BandedGridView;if(view==null)continue;banded++;Check(view.GridControl==grid&&view.Columns.Count>0,"Banded Rent header view attaches after construction: "+tabs.SelectedTabPage.Text);
     using(var graphics=grid.CreateGraphics())Console.WriteLine("DPI actual grid="+grid.DeviceDpi+" graphics="+graphics.DpiX+"/"+graphics.DpiY+" tab="+tabs.SelectedTabPage.Text);
     if(args.Length>4&&args[4]=="header-diagnostic")HeaderEligibility(dit,grid,view);
     app.GetType("Abovo.GridPresentation").GetMethod("SetZoom",F).Invoke(null,new object[]{grid,130,false});Pump();app.GetType("Abovo.GridPresentation").GetMethod("SetZoom",F).Invoke(null,new object[]{grid,100,false});Pump();
    }
   }
   Check(banded>0,"Rent exercise includes repeating banded header editors");
   Call(host,"ShowInterface",(int)model.ModelID,1,false,"None",null,0);Pump();Check(Convert.ToInt32(Read(Read(host,"ActiveInterface"),"CSID"))==1,"Can navigate from Rent to Stock Assumptions");
   Call(host,"ShowInterface",(int)model.ModelID,2,false,"None",null,2);Pump();dit=(Control)Read(host,"ActiveInterface");Check(Convert.ToInt32(Read(dit,"CSID"))==2,"Can return to Rent after switching interfaces");
   Check(GetForegroundWindow()!=host.Handle,"Private Rent host never takes foreground focus");Check(Convert.ToInt64(Read((object)model,"UserChangeRevision"))==revision&&(bool)model.IsDirty==dirty,"Lazy tabs and zoom do not create workbook edits or change dirty state");host.Close();
  }
  model.ModelSpreadsheetControl.Dispose();Console.WriteLine("PASS "+checks+" assertions; original workbook untouched; private copy not saved.");return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}finally{uint process;GetWindowThreadProcessId(GetForegroundWindow(),out process);if(previous!=IntPtr.Zero&&process==(uint)System.Diagnostics.Process.GetCurrentProcess().Id)SetForegroundWindow(previous);}}
}
