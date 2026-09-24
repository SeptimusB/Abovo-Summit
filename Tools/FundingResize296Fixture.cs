// Native actual-DIT resize regression. Opens only a private workbook copy; never saves.
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
using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraVerticalGrid;
using DevExpress.XtraVerticalGrid.Rows;

class FundingResize296Fixture {
 [DllImport("user32.dll")]static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]static extern bool SetForegroundWindow(IntPtr window);
 [DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
 [DllImport("user32.dll")]static extern uint GetGuiResources(IntPtr process,uint flags);
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static int checks,headerPaints,invalidHeaderBrushes;static string evidence;static Type presentation,application;static object helper;static Control dit;static VGridControl grid;static EditorRow dateRow;static BaseEdit active;
 static DateTime pending=new DateTime(2042,7,24);static Dictionary<string,double[]> sizes=new Dictionary<string,double[]>();
 static Form QuietHost(Type baseType,object[] arguments){
  var assembly=AppDomain.CurrentDomain.DefineDynamicAssembly(new AssemblyName("FundingQuietHost"+Guid.NewGuid().ToString("N")),AssemblyBuilderAccess.Run);
  var type=assembly.DefineDynamicModule("Main").DefineType("FundingQuietHost",TypeAttributes.Public,baseType);
  var signature=new[]{typeof(int),typeof(int),typeof(string)};var constructor=type.DefineConstructor(MethodAttributes.Public,CallingConventions.Standard,signature);var il=constructor.GetILGenerator();
  il.Emit(OpCodes.Ldarg_0);il.Emit(OpCodes.Ldarg_1);il.Emit(OpCodes.Ldarg_2);il.Emit(OpCodes.Ldarg_3);il.Emit(OpCodes.Call,baseType.GetConstructor(signature));il.Emit(OpCodes.Ret);
  var showBase=typeof(Form).GetProperty("ShowWithoutActivation",F).GetGetMethod(true);var show=type.DefineMethod(showBase.Name,MethodAttributes.Family|MethodAttributes.Virtual|MethodAttributes.HideBySig|MethodAttributes.SpecialName,typeof(bool),Type.EmptyTypes);il=show.GetILGenerator();il.Emit(OpCodes.Ldc_I4_1);il.Emit(OpCodes.Ret);type.DefineMethodOverride(show,showBase);
  var createBase=baseType.GetProperty("CreateParams",F).GetGetMethod(true);var create=type.DefineMethod(createBase.Name,MethodAttributes.Family|MethodAttributes.Virtual|MethodAttributes.HideBySig|MethodAttributes.SpecialName,typeof(CreateParams),Type.EmptyTypes);il=create.GetILGenerator();var value=il.DeclareLocal(typeof(CreateParams));
  il.Emit(OpCodes.Ldarg_0);il.Emit(OpCodes.Call,createBase);il.Emit(OpCodes.Stloc,value);il.Emit(OpCodes.Ldloc,value);il.Emit(OpCodes.Ldloc,value);il.Emit(OpCodes.Callvirt,typeof(CreateParams).GetProperty("ExStyle").GetGetMethod());il.Emit(OpCodes.Ldc_I4,0x08000000);il.Emit(OpCodes.Or);il.Emit(OpCodes.Callvirt,typeof(CreateParams).GetProperty("ExStyle").GetSetMethod());il.Emit(OpCodes.Ldloc,value);il.Emit(OpCodes.Ret);type.DefineMethodOverride(create,createBase);
  return (Form)Activator.CreateInstance(type.CreateType(),arguments);
 }
 static object Read(object o,string n){for(var type=o.GetType();type!=null;type=type.BaseType){var field=type.GetField(n,F|BindingFlags.DeclaredOnly);if(field!=null)return field.GetValue(o);var property=type.GetProperty(n,F|BindingFlags.DeclaredOnly);if(property!=null)return property.GetValue(o,null);}throw new MissingMemberException(o.GetType().FullName,n);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static IEnumerable<Control> Controls(Control c){foreach(Control child in c.Controls){yield return child;foreach(var inner in Controls(child))yield return inner;}}
 static IEnumerable<BaseRow> Rows(IEnumerable rows){foreach(BaseRow row in rows){yield return row;foreach(var child in Rows(row.ChildRows))yield return child;}}
 static void Check(bool value,string text){if(!value)throw new Exception(text);Console.WriteLine("PASS "+(++checks)+" "+text);}
 static void Resources(string stage){using(var process=System.Diagnostics.Process.GetCurrentProcess())Console.WriteLine("RESOURCE "+stage+" gdi="+GetGuiResources(process.Handle,0)+" user="+GetGuiResources(process.Handle,1)+" privateBytes="+process.PrivateMemorySize64+" recordHeader="+grid.RecordHeaderHeight+" recordWidth="+grid.RecordWidth+" leftRecord="+grid.LeftVisibleRecord+" topRow="+grid.TopVisibleRowIndex+" rowHeader="+grid.RowHeaderWidth+" client="+grid.ClientSize);}
 static void HeaderBrushes(object sender,DevExpress.XtraVerticalGrid.Events.CustomDrawRecordHeaderEventArgs e){
  if(e.Record!=0||(int)presentation.GetMethod("ZoomPercent",F).Invoke(null,new object[]{grid})!=200)return;
  var info=grid.ViewInfo;var panel=info.RecordHeaderPanelViewInfo;int invalid=0,stale=0,index=0;var bad=new List<string>();
  foreach(DevExpress.XtraVerticalGrid.ViewInfo.LineInfo line in panel.LinesInfo){
   bool matches=ReferenceEquals(line.Brush,info.RC.RowHorzLineBrush)||ReferenceEquals(line.Brush,info.RC.RowVertLineBrush);if(!matches)stale++;
   try{using(var clone=(Brush)line.Brush.Clone()){} }catch(Exception ex){invalid++;bad.Add(index+":"+line.Rect+":"+ex.GetType().Name+":matchesRC="+matches);}
   index++;
  }
  Console.WriteLine("HEADERBRUSH rows="+info.RowsViewInfo.Count+" editorRows="+info.RowsViewInfo.Cast<DevExpress.XtraVerticalGrid.ViewInfo.BaseRowViewInfo>().Count(row=>row.Row is EditorRow)+" lines="+panel.LinesInfo.Count+" invalid="+invalid+" stale="+stale+" panel="+panel.Bounds+" headers="+panel.HeadersBounds+" client="+grid.ClientSize+" top="+grid.TopVisibleRowIndex+" bad="+String.Join(";",bad.ToArray()));
  headerPaints++;invalidHeaderBrushes+=invalid+stale;
 }
 static void Pump(int ms=180){var until=DateTime.UtcNow.AddMilliseconds(ms);do{Application.DoEvents();Thread.Sleep(10);}while(DateTime.UtcNow<until);}
 static void Zoom(int percent){presentation.GetMethod("SetZoom",F).Invoke(null,new object[]{grid,percent,false});Pump();}
 static string Values(CellRange range){return String.Join("|",range.ExistingCells.Select(c=>c.GetReferenceA1()+":"+c.Formula+":"+c.Value.ToString()+":"+c.NumberFormat+":"+c.Fill.PatternType+":"+c.Fill.BackgroundColor+":"+c.Font.Name+":"+c.Font.Size+":"+c.Font.Bold+":"+c.Font.Italic+":"+c.Font.Color+":"+c.Protection.Locked).ToArray());}
 static void Verify(Form host,int percent,string phase,bool compareGeometry=true){
  Check(grid.ViewInfo.RowsViewInfo.Count>0,phase+": native data rows remain visible immediately after settled resize");
  grid.MakeRowVisible(dateRow);grid.Refresh();Pump();
  Check((int)presentation.GetMethod("ZoomPercent",F).Invoke(null,new object[]{grid})==percent,phase+": local zoom retained");
  float owner=grid.Appearance.RecordValue.GetFont().SizeInPoints;
  using(var font=(Font)application.GetMethod("GetDisplayFont").Invoke(null,new object[]{"Medium",dit,false,false,false})){
   Console.WriteLine("FONT "+phase+" owner="+owner+" baseline="+font.SizeInPoints+" expected="+Math.Max(8.0,font.SizeInPoints*percent/100.0));
   Check(Math.Abs(owner-Math.Max(8.0,font.SizeInPoints*percent/100.0))<0.08,phase+": body font equals density times zoom with eight-point readability floor");
  }
  Check(ReferenceEquals(active,Read(helper,"_ActiveEditor"))&&!active.IsDisposed&&Convert.ToDateTime(active.EditValue)==pending,phase+": same pending date editor/value retained");
  var item=(RepositoryItemDateEdit)Read(helper,"_Item");float header=item.Appearance.GetFont().SizeInPoints;
  Check(Math.Abs(header-owner)<0.06&&header>=7.99&&Math.Abs(active.Properties.Appearance.GetFont().SizeInPoints-header)<0.05,phase+": date repository and active editor match the readable owner font");
  Check(new[]{item.AppearanceFocused,item.AppearanceDisabled,item.AppearanceReadOnly}.All(a=>Math.Abs(a.GetFont().SizeInPoints-header)<0.05),phase+": all date appearance states agree");
  var headerBounds=(Rectangle)Read(helper,"_LastHeaderBounds");var editorBounds=(Rectangle)Call(helper,"GetEditorBounds",headerBounds);
  Check(!headerBounds.IsEmpty&&headerBounds.Contains(editorBounds)&&active.Size==editorBounds.Size,phase+": active date editor stays within the row-header geometry");
  int activeRight=grid.PointToClient(active.Parent.PointToScreen(new Point(active.Right,active.Top))).X;
  Check(activeRight<=headerBounds.Right-3,phase+": active date editor preserves the fixed-pane three-pixel boundary");
  using(var graphics=grid.CreateGraphics()){
   Check(item.Appearance.GetFont().GetHeight(graphics)>0,phase+": original repository font remains valid after native grid painting");
   Console.WriteLine("METRIC "+phase+" host="+host.ClientSize+" grid="+grid.Size+" dpi="+graphics.DpiX+"/"+graphics.DpiY+" hostDpi="+host.DeviceDpi+" owner="+owner+" header="+header+" record="+grid.RecordWidth+" rowheader="+grid.RowHeaderWidth+" row="+dateRow.Height+" editor="+editorBounds);
  }
  if(compareGeometry){
   string key=host.ClientSize.Width+"x"+host.ClientSize.Height+"/"+percent;
   var current=new double[]{owner,header,grid.RecordWidth,grid.RowHeaderWidth,dateRow.Height,grid.RecordHeaderHeight};double[] previous;
   if(sizes.TryGetValue(key,out previous))Check(current.Zip(previous,(a,b)=>Math.Abs(a-b)).Select((d,i)=>d<=(i<2?0.08:2)).All(x=>x),phase+": repeated window size has no cumulative font/width/height growth");
   else sizes.Add(key,current);
  }
 }
 static void WheelAroundPointer(){
  Zoom(100);grid.MakeRowVisible(dateRow);grid.LeftVisibleRecord=2;grid.Refresh();Pump();
  var info=grid.ViewInfo.GetRowValueInfo(dateRow,3,0);Check(info!=null&&!info.Bounds.IsEmpty,"Pointer regression has a visible Funding data cell");
  var before=info.Bounds;var point=new Point(before.Left+before.Width/2,before.Top+before.Height/2);Check(grid.ClientRectangle.Contains(point),"Pointer regression targets inside the real Funding viewport");
  var hit=grid.CalcHitInfo(point);Console.WriteLine("POINTERCAPTURE row="+dateRow.Properties.Caption+" hitrow="+(hit.Row==null?"null":hit.Row.Properties.Caption)+" record="+hit.RecordIndex+" type="+hit.HitInfoType+" before="+before+" top="+grid.TopVisibleRowIndex+" left="+grid.LeftVisibleRecord);
  {
   presentation.GetMethod("SetZoomAt",F).Invoke(null,new object[]{grid,110,(Point?)point,false});Pump();
   Check((int)presentation.GetMethod("ZoomPercent",F).Invoke(null,new object[]{grid})==110,"Pointer zoom applies a ten-percent step without moving the physical cursor");
   var after=grid.ViewInfo.GetRowValueInfo(dateRow,3,0);Console.WriteLine("POINTERAFTER rows="+grid.ViewInfo.RowsViewInfo.Count+" top="+grid.TopVisibleRowIndex+" left="+grid.LeftVisibleRecord+" cell="+(after==null?"null":after.Bounds.ToString()));Check(after!=null&&!after.Bounds.IsEmpty,"Hovered Funding source cell remains visible after zoom");
   var rectangle=after.Bounds;double dx=rectangle.Left+rectangle.Width/2.0-point.X,dy=rectangle.Top+rectangle.Height/2.0-point.Y;
   Console.WriteLine("POINTER before="+before+" after="+rectangle+" point="+point+" offset="+dx+","+dy);
   Check(Math.Abs(dx)<=grid.RecordWidth/2.0+2&&Math.Abs(dy)<=rectangle.Height+2,"Hovered Funding cell stays within native record/row scroll granularity of pointer");
   Check(ReferenceEquals(active,Read(helper,"_ActiveEditor"))&&!active.IsDisposed&&Convert.ToDateTime(active.EditValue)==pending,"Pointer zoom retains the same uncommitted date editor");
  }
  Zoom(100);
 }
 [STAThread]static int Main(string[] args){IntPtr previousWindow=GetForegroundWindow();try{
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);evidence=args[1];
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));application=app.GetType("Abovo.AbovoAppCls");presentation=app.GetType("Abovo.GridPresentation");application.GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  string copy=Path.Combine(args[1],"private-funding-resize296.xlsb");File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!loaded.BError,"Disposable model opens");dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)loaded.IntegerReturn);IWorkbook wb=model.WB;
  using(var host=QuietHost(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})){
   host.Opacity=0;host.ShowInTaskbar=false;host.StartPosition=FormStartPosition.Manual;host.Location=new Point(20,20);host.ClientSize=new Size(5000,1400);host.Show();Pump(500);
   Call(host,"ShowInterface",(int)model.ModelID,33,false,"None",null,1);Pump();dit=(Control)Read(host,"ActiveInterface");
   foreach(var tabs in Controls(dit).OfType<DevExpress.XtraTab.XtraTabControl>())tabs.SelectedTabPageIndex=1;
   Call(dit,"BuildSection",1,false,false);Pump(500);
   grid=Controls(dit).OfType<VGridControl>().First(g=>g.Visible&&Rows(g.Rows).Any(r=>r.Properties.Caption.Contains("Facility")&&r.Properties.Caption.Contains("Name")));
   grid.CustomDrawRecordHeader+=HeaderBrushes;
   presentation.GetMethod("Configure",F).Invoke(null,new object[]{grid,"Regression/FundingResize296/"+Guid.NewGuid()});Zoom(100);
   dateRow=Rows(grid.Rows).OfType<EditorRow>().First(r=>{var tag=Call(dit,"GetVGridColumnTag",r,0);return tag!=null&&Convert.ToString(Read(tag,"BandID"))=="Facility Decreases"&&Read(tag,"InColumnEditorDate")!=null;});
   var rowTag=Call(dit,"GetVGridColumnTag",dateRow,0);var headerDate=Read(rowTag,"InColumnEditorDate");var tagDate=Read(headerDate,"Tag");helper=Read(tagDate,"InPlaceVGridRowHelper");
   var range=wb.DefinedNames.GetDefinedName(Convert.ToString(Read(tagDate,"EditingNRName"))).Range;string before=Values(range);long revision=Convert.ToInt64(Read((object)model,"UserChangeRevision"));bool dirty=(bool)model.IsDirty;
   Call(helper,"ShowEditorFromKeyboard");Pump();active=(BaseEdit)Read(helper,"_ActiveEditor");Check(active!=null,"Actual Funding repeating-date editor opens");Check(GetForegroundWindow()!=host.Handle,"Private Funding host does not become the foreground window");active.EditValue=pending;
   Zoom(50);Check((int)presentation.GetMethod("ZoomPercent",F).Invoke(null,new object[]{grid})==60,"Legacy50% zoom requests clamp to the new60% minimum");
   foreach(int percent in new[]{60,100,200}){
    Zoom(percent);
    foreach(int width in new[]{5000,1800,5000,1800}){
     Resources("before width "+width+" zoom "+percent);
     host.WindowState=FormWindowState.Normal;host.ClientSize=new Size(width,1400);Pump(500);
     Check(Math.Abs(host.ClientSize.Width-width)<=2,"Actual host reaches requested width "+width);
     Verify(host,percent,"width "+width+" zoom "+percent);
     Check(Values(range)==before&&Convert.ToInt64(Read((object)model,"UserChangeRevision"))==revision&&(bool)model.IsDirty==dirty,"Resize/zoom does not commit the pending date or dirty the workbook");
    }
    var restored=host.ClientSize;host.WindowState=FormWindowState.Maximized;Pump(500);Check(host.WindowState==FormWindowState.Maximized,"Actual maximised state reached at "+percent+"%");Verify(host,percent,"maximised zoom "+percent,false);
    host.WindowState=FormWindowState.Normal;Pump(500);Check(host.WindowState==FormWindowState.Normal&&host.ClientSize==restored,"Actual restore returns prior client size at "+percent+"%");Verify(host,percent,"restored zoom "+percent);
   }
   Zoom(100);Verify(host,100,"final reset");
   WheelAroundPointer();
   Call(helper,"Editor_KeyDown",active,new KeyEventArgs(Keys.Escape));Pump();
   Check(Values(range)==before&&Convert.ToInt64(Read((object)model,"UserChangeRevision"))==revision&&(bool)model.IsDirty==dirty,"Cancelling fixture edit leaves all source dates/history/dirty state unchanged");
   Check(headerPaints>0&&invalidHeaderBrushes==0,"Native record-header line brushes stay valid and current throughout repeated resize/zoom");
   // Capture only after all resize/zoom checks, avoiding extra paint or layout during the regression.
   using(var picture=new Bitmap(grid.Width,grid.Height)){
    grid.DrawToBitmap(picture,grid.ClientRectangle);
    picture.Save(Path.Combine(evidence,"funding-final-reset.png"),System.Drawing.Imaging.ImageFormat.Png);
   }
   host.Close();
  }
  model.ModelSpreadsheetControl.Dispose();Console.WriteLine("PASS "+checks+" assertions; original workbook untouched; private copy was not saved.");return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}finally{
  IntPtr current=GetForegroundWindow();uint process;GetWindowThreadProcessId(current,out process);
  if(previousWindow!=IntPtr.Zero&&(current==IntPtr.Zero||process==(uint)System.Diagnostics.Process.GetCurrentProcess().Id))SetForegroundWindow(previousWindow);
 }}
}
