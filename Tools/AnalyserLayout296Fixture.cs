using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraTab;
using DevExpress.Utils.Drawing;

class AnalyserLayout296Fixture {
 sealed class QuietHost:Form {
  protected override bool ShowWithoutActivation {get{return true;}}
  protected override CreateParams CreateParams {get{var p=base.CreateParams;p.ExStyle|=0x08000000;return p;}}
 }
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static int checks;
 [DllImport("user32.dll")]static extern bool SetProcessDpiAwarenessContext(IntPtr context);
 static object Get(object o,string n){var f=o.GetType().GetField(n,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(n,F).GetValue(o,null);}
 static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("PASS "+(++checks)+" "+text);}
 static void Pump(){Application.DoEvents();}
 static void Picture(Control c,string path){using(var b=new Bitmap(c.Width,c.Height)){c.DrawToBitmap(b,new Rectangle(Point.Empty,b.Size));b.Save(path);}}
 static void Zoom(Type p,Control c,int n){p.GetMethod("SetZoom",F).Invoke(null,new object[]{c,n,false});Pump();c.Refresh();Pump();}
 static void CheckPointer(Type presentation,GridView view,int kind){
  Zoom(presentation,view.GridControl,100);view.TopRowIndex=0;view.LeftCoord=0;view.LayoutChanged();Pump();
  var info=(GridViewInfo)view.GetViewInfo();Rectangle rowBounds=Rectangle.Empty;
  foreach(GridRowInfo row in info.RowsInfo){
   if(kind==1&&row.IsGroupRow&&row.Bounds.Top>120&&row.Bounds.Bottom<view.GridControl.Height-100){rowBounds=row.Bounds;break;}
   if(kind==2)foreach(GridRowFooterInfo footer in row.RowFooters)if(footer.Bounds.Top>100&&footer.Bounds.Bottom<view.GridControl.Height-100){rowBounds=footer.Bounds;break;}
   if(kind==1&&!rowBounds.IsEmpty)break;
  }
  Check(!rowBounds.IsEmpty,"Visible native analyser row available for pointer kind "+kind);
  var column=view.VisibleColumns.Cast<DevExpress.XtraGrid.Columns.GridColumn>().Where(c=>c.FieldName!="ItemDesc").Skip(3).First();
  var header=info.ColumnsInfo[column].Bounds;var point=new Point(header.Left+header.Width/2,rowBounds.Top+rowBounds.Height/2);
  var capture=presentation.GetMethod("CaptureAnchor",F);var anchor=capture.Invoke(null,new object[]{view.GridControl,point});
  Check(anchor!=null&&(int)Get(anchor,"GridRowKind")==kind,"Zoom captures analyser "+(kind==1?"group heading":"group total"));
  presentation.GetMethod("SetZoomAt",F).Invoke(null,new object[]{view.GridControl,120,point,false});Pump();
  var after=(Rectangle)presentation.GetMethod("AnchorBounds",F).Invoke(null,new object[]{view.GridControl,anchor});
  double fx=(double)Get(anchor,"FractionX"),fy=(double)Get(anchor,"FractionY");
  double dx=Math.Abs(after.Left+after.Width*fx-point.X),dy=Math.Abs(after.Top+after.Height*fy-point.Y);
  Console.WriteLine("POINTER kind="+kind+" before="+rowBounds+" after="+after+" deltaX="+dx+" deltaY="+dy);
  Check(!after.IsEmpty&&dx<=3&&dy<=Math.Max(24,after.Height*2),"Analyser virtual row stays near the pointer after zoom");
 }
 [STAThread] static int Main(string[] args){try{
  bool highDpi=args.Length>4&&args[4]=="hidpi";
  if(highDpi)Console.WriteLine("PERMONITORV2 "+SetProcessDpiAwarenessContext(new IntPtr(-4)));
  Application.EnableVisualStyles();
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  if(args.Length>4&&args[4]=="api"){
  foreach(var m in typeof(GraphicsCache).GetMethods().Where(m=>m.Name.Contains("Clip")))Console.WriteLine("CLIPAPI "+m+" params="+string.Join(",",m.GetParameters().Select(p=>p.Name)));
  foreach(var p in typeof(DevExpress.XtraGrid.Drawing.GridFooterCellInfoArgs).GetProperties(F))Console.WriteLine("FOOTERAPI "+p);
  foreach(var f in typeof(DevExpress.XtraGrid.Drawing.GridFooterCellInfoArgs).GetFields(F))Console.WriteLine("FOOTERFIELD "+f);
  foreach(var p in typeof(GridGroupRowInfo).GetProperties(F).Where(p=>p.Name.Contains("Text")||p.Name.Contains("Caption")))Console.WriteLine("GROUPAPI "+p+" write="+p.CanWrite);
  foreach(var f in typeof(GridGroupRowInfo).GetFields(F).Where(f=>f.Name.Contains("Text")||f.Name.Contains("Caption")))Console.WriteLine("GROUPFIELD "+f);
  return 0;}
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  var path=Path.Combine(args[1],"analyser-layout296.xlsb");File.Copy(args[2],path);
  var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{path,new FileInfo(path),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!result.BError,"Private model opened");int model=(int)result.IntegerReturn;
  bool probe=args.Length>4&&args[4]=="probe";
  var presentation=app.GetType("Abovo.GridPresentation");
  using(var host=new QuietHost{Opacity=0,ShowInTaskbar=false,ClientSize=new Size(1600,950)})
  using(var analyser=(Control)Activator.CreateInstance(app.GetType("BPIncomeExpenditureAnalyserV2"),new object[]{model,null})){
   host.Controls.Add(analyser);analyser.Dock=DockStyle.Fill;host.Show();Pump();
   var tabs=(XtraTabControl)Get(analyser,"XtraTabControlAnalyser");
   for(int tab=0;tab<2;tab++){
    tabs.SelectedTabPageIndex=tab;Pump();var wrap=Get(analyser,tab==0?"WrapCG_SOCI":"WrapCG_CF");var view=(GridView)Get(wrap,"WrappedGridView");var grid=view.GridControl;
    Console.WriteLine("INITIALDESCRIPTION tab="+tab+" width="+view.Columns["ItemDesc"].Width+" max="+view.Columns["ItemDesc"].MaxWidth+" viewport="+grid.ClientSize.Width);
    Picture(grid,Path.Combine(args[1],"analyser-"+tab+"-initial.png"));
    Check(view.Columns["ItemDesc"].Width>=250&&view.Columns["ItemDesc"].Width<=grid.ClientSize.Width*.41,"Initial docked layout fits group captions without a fixture refit");
    presentation.GetMethod("Configure",F).Invoke(null,new object[]{grid,"Regression/layout296/"+Guid.NewGuid()});
    int footerCells=0,badFooterCells=0,groupFigures=0,badGroupFigures=0,clippedFigures=0;string sample=null,groupSample=null,figureSample=null;
    view.CustomDrawRowFooterCell+=(s,e)=>{
     if(e.Column==null||e.Column.FieldName=="ItemDesc")return;
     var info=(GridViewInfo)view.GetViewInfo();var header=info.ColumnsInfo[e.Column];if(header==null)return;
     footerCells++;
     var fixedBounds=info.ColumnsInfo[view.Columns["ItemDesc"]].Bounds;
     int left=Math.Max(header.Bounds.Left,fixedBounds.Right),right=Math.Min(header.Bounds.Right,grid.ClientSize.Width);
     bool bad=right>left&&(e.Bounds.Left<left-1||e.Bounds.Right>right+1);
     if(bad){badFooterCells++;if(sample==null)sample=e.Column.FieldName+" footer="+e.Bounds+" header="+header.Bounds+" fixed="+fixedBounds;}
     if(header.Bounds.Left>=info.ViewRects.FixedLeft.Right&&header.Bounds.Right<=grid.ClientSize.Width){
      string value=e.Info.DisplayText;double number;
      if(double.TryParse(value,NumberStyles.Any,CultureInfo.CurrentCulture,out number)){
       if(number<0&&value.StartsWith("-",StringComparison.Ordinal))value="("+value.Substring(1)+")";
       float needed=e.Cache.CalcTextSize(value,e.Appearance.GetFont()).Width;
       if(needed>header.Bounds.Width-2){clippedFigures++;if(figureSample==null)figureSample=value+" needs="+needed+" column="+header.Bounds.Width+" font="+e.Appearance.GetFont();}
      }
     }
    };
    view.CustomDrawGroupRow+=(s,e)=>{
     var info=(GridViewInfo)view.GetViewInfo();var desc=info.ColumnsInfo[view.Columns["ItemDesc"]].Bounds;
     var label=(Rectangle)analyser.GetType().GetMethod("StatementDescriptionBounds",F).Invoke(null,new object[]{view,e.Bounds});
     if(label.Right>desc.Right||label.Left<desc.Left)throw new Exception("Group caption escaped description region");
     foreach(var column in view.VisibleColumns.Cast<DevExpress.XtraGrid.Columns.GridColumn>().Where(c=>c.FieldName!="ItemDesc")){
      var header=info.ColumnsInfo[column];if(header==null)continue;
      int left=Math.Max(header.Bounds.Left+1,info.ViewRects.FixedLeft.Right),right=Math.Min(header.Bounds.Right-2,grid.ClientSize.Width);
      if(right<=left)continue;
      var shortRect=(Rectangle)analyser.GetType().GetMethod("CalcRowSummaryRect",F).Invoke(analyser,new object[]{"1",e,column});
      var longRect=(Rectangle)analyser.GetType().GetMethod("CalcRowSummaryRect",F).Invoke(analyser,new object[]{new string('W',200),e,column});
      groupFigures++;if(shortRect.IsEmpty||shortRect!=longRect||shortRect.Left<desc.Right){badGroupFigures++;if(groupSample==null)groupSample=column.FieldName+" short="+shortRect+" long="+longRect+" header="+header.Bounds+" fixed="+info.ViewRects.FixedLeft;}
     }
    };
    foreach(int zoom in new[]{60,100,150,200}){
     Zoom(presentation,grid,zoom);view.TopRowIndex=0;view.LeftCoord=0;Pump();
     int baseFooter=Math.Max(12,(int)Math.Ceiling(view.Appearance.GroupFooter.GetFont().GetHeight(grid.DeviceDpi))+Math.Max(3,(int)Math.Round(4*grid.DeviceDpi/96.0*zoom/100.0)));
     var info=(GridViewInfo)view.GetViewInfo();Console.WriteLine("LAYOUT tab="+tab+" dpi="+grid.DeviceDpi+" zoom="+zoom+" footer="+info.GroupFooterHeight+" old="+baseFooter);
     if(!probe)Check(info.GroupFooterHeight>=Math.Ceiling(baseFooter*1.1),"Total row has requested extra spacing at "+zoom+"%");
     var desc=info.ColumnsInfo[view.Columns["ItemDesc"]].Bounds;
     Console.WriteLine("DESCRIPTION "+desc+" fixedRect="+info.ViewRects.FixedLeft);
     foreach(var c in view.VisibleColumns.Cast<DevExpress.XtraGrid.Columns.GridColumn>().Take(3))Console.WriteLine("HEADER "+c.FieldName+" "+info.ColumnsInfo[c].Bounds);
     Picture(grid,Path.Combine(args[1],"analyser-"+tab+"-"+zoom+".png"));
     view.LeftCoord=100;Pump();Picture(grid,Path.Combine(args[1],"analyser-"+tab+"-"+zoom+"-scroll.png"));
    }
    Zoom(presentation,grid,100);host.ClientSize=new Size(1120,850);view.Columns["ItemDesc"].Width=130;view.LeftCoord=120;Pump();grid.Refresh();Pump();Picture(grid,Path.Combine(args[1],"analyser-"+tab+"-narrow.png"));
    host.Width+=40;Pump();Check(view.Columns["ItemDesc"].Width==130,"Manual description width survives later resize");host.Width-=40;Pump();
    Console.WriteLine("FOOTERCELLS total="+footerCells+" outsideHeader="+badFooterCells+" first="+sample);
    if(!probe)Check(footerCells>0&&badFooterCells==0,"Footer numbers stay in their visible header column after zoom and horizontal scroll");
    Console.WriteLine("CLIPPEDFIGURES total="+clippedFigures+" first="+figureSample);
    if(!probe)Check(clippedFigures==0,"Complete financial totals fit fully visible period columns at every tested zoom");
    Console.WriteLine("GROUPFIGURES total="+groupFigures+" omittedOrOverlapping="+badGroupFigures+" first="+groupSample);
    if(!probe)Check(groupFigures>0&&badGroupFigures==0,"Group totals remain visible regardless of caption length and stay outside the fixed description");
    CheckPointer(presentation,view,1);CheckPointer(presentation,view,2);
    if(highDpi)Console.WriteLine("DPI-COVERAGE actual="+grid.DeviceDpi+"; config=PerMonitorV2; above96="+(grid.DeviceDpi>96));
    host.ClientSize=new Size(1600,950);
    Zoom(presentation,grid,100);
   }
   host.Close();
  }
  files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{model});Console.WriteLine("PASS "+checks+" layout checks");return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
