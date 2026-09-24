using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraVerticalGrid;
using DevExpress.XtraVerticalGrid.Rows;
using DevExpress.XtraTreeList;

// Presentation-only native regression. No workbook is opened or changed.
class GridPointerZoom296Fixture {
 class QuietForm:Form { protected override bool ShowWithoutActivation { get { return true; } } }
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Static|BindingFlags.Instance;
 static Type presentation; static int checks;
 static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("PASS "+(++checks)+" "+text);}
 static void Pump(Control c){Application.DoEvents();c.Refresh();Application.DoEvents();}
 static void Zoom(Control c,int percent){presentation.GetMethod("SetZoom",F).Invoke(null,new object[]{c,percent,false});Pump(c);}
 static void At(Control c,int percent,Point point){presentation.GetMethod("SetZoomAt",F).Invoke(null,new object[]{c,percent,point,false});Pump(c);}
 static int Percent(Control c){return (int)presentation.GetMethod("ZoomPercent",F).Invoke(null,new object[]{c});}
 static void Wheel(Control c,Func<int> horizontal,string name){
  Zoom(c,100);int before=horizontal();
  Check(!(bool)presentation.GetMethod("ModifiedWheel",F).Invoke(null,new object[]{c,-120,0}),name+" ordinary wheel remains native vertical");
  Check((bool)presentation.GetMethod("ModifiedWheel",F).Invoke(null,new object[]{c,-120,4}),name+" Shift wheel handled");Pump(c);
  Check(horizontal()>before&&Percent(c)==100,name+" Shift wheel scrolls horizontally without zooming");
  Check((bool)presentation.GetMethod("ModifiedWheel",F).Invoke(null,new object[]{c,120,8}),name+" Ctrl wheel handled");
  var clock=System.Diagnostics.Stopwatch.StartNew();while(clock.ElapsedMilliseconds<350){Application.DoEvents();System.Threading.Thread.Sleep(10);}Pump(c);
  Check(Percent(c)==110,name+" Ctrl wheel zooms");
  var menu=c.ContextMenuStrip.Items.Find("SummitGridZoom",false).FirstOrDefault();Check(menu!=null&&menu.Text.Contains("Ctrl + wheel"),name+" context menu advertises new modifier");Zoom(c,100);
 }
 static DataTable Data(){var t=new DataTable();for(int i=0;i<32;i++)t.Columns.Add("C"+i);for(int r=0;r<100;r++)t.Rows.Add(Enumerable.Range(0,32).Select(c=>(object)("R"+r+"C"+c)).ToArray());return t;}
 static void Readable(Control c,Font font,int height,string name){using(var g=c.CreateGraphics()){Check(font.SizeInPoints>=7.99f,name+" font has readable floor");Check(height>=Math.Ceiling(font.GetHeight(g)),name+" row fits the font at device DPI "+c.DeviceDpi);}}
 static void Near(Rectangle before,Rectangle after,Point point,bool discreteX,string name){
  Check(!after.IsEmpty,name+" anchor cell remains visible");
  double fx=(double)(point.X-before.Left)/before.Width,fy=(double)(point.Y-before.Top)/before.Height;
  double dx=Math.Abs(after.Left+after.Width*fx-point.X),dy=Math.Abs(after.Top+after.Height*fy-point.Y);
  Console.WriteLine("ANCHOR "+name+" before="+before+" after="+after+" point="+point+" delta="+dx+","+dy);
  Check(dx<=(discreteX?after.Width+3:3),name+" horizontal anchor stays within native scroll precision");
  Check(dy<=after.Height+3,name+" vertical anchor stays within one native row");
 }
 static void Standard(){using(var host=new QuietForm{Opacity=0,ShowInTaskbar=false,ClientSize=new Size(1000,650)})using(var grid=new GridControl()){
  var view=new GridView(grid);grid.MainView=view;grid.DataSource=Data();grid.Dock=DockStyle.Fill;host.Controls.Add(grid);view.OptionsView.ColumnAutoWidth=false;view.OptionsView.ShowGroupPanel=false;view.RowHeight=26;view.Appearance.Row.Font=new Font("Segoe UI",12);view.Appearance.Row.Options.UseFont=true;
  host.Show();Pump(grid);foreach(DevExpress.XtraGrid.Columns.GridColumn c in view.Columns){c.Width=110;c.MinWidth=20;}view.Columns[0].Fixed=DevExpress.XtraGrid.Columns.FixedStyle.Left;
  presentation.GetMethod("Configure",F).Invoke(null,new object[]{grid,"Regression/Pointer/"+Guid.NewGuid()});Pump(grid);Zoom(grid,50);Check(Percent(grid)==60,"Grid legacy50% request clamps to60%");Readable(grid,view.Appearance.Row.GetFont(),view.RowHeight,"Grid");
  Check(view.Columns[1].Width>=74,"Grid columns respect the eight-point text floor");
  Zoom(grid,100);view.TopRowIndex=12;view.LeftCoord=350;Pump(grid);
  foreach(int target in new[]{110,90,150,100}){
   var point=new Point(820,350);var hit=view.CalcHitInfo(point);
   // App DPI configuration can place this nominal point on a row separator.
   // Test an actual data-cell anchor; retain all post-zoom position assertions.
   for(int probe=0;probe<10&&!hit.InRowCell;probe++){point.Y++;hit=view.CalcHitInfo(point);}
   Check(hit.InRowCell,"Grid pointer starts over data");var before=((GridViewInfo)view.GetViewInfo()).GetGridCellInfo(hit.RowHandle,hit.Column).Bounds;
   At(grid,target,point);var info=((GridViewInfo)view.GetViewInfo()).GetGridCellInfo(hit.RowHandle,hit.Column);Near(before,info==null?Rectangle.Empty:info.Bounds,point,false,"Grid "+target);
  }
  Zoom(grid,60);int h=view.RowHeight;Zoom(grid,100);Zoom(grid,60);Check(view.RowHeight==h,"Grid minimum row height does not grow after reset");
  Wheel(grid,()=>view.LeftCoord,"Grid");using(var b=new Bitmap(grid.Width,grid.Height)){grid.DrawToBitmap(b,new Rectangle(Point.Empty,b.Size));}
 }}
 static void Vertical(){using(var host=new QuietForm{Opacity=0,ShowInTaskbar=false,ClientSize=new Size(1000,650)})using(var grid=new VGridControl()){
  grid.DataSource=Data();grid.Dock=DockStyle.Fill;grid.RowHeaderWidth=160;grid.RecordWidth=110;grid.RecordHeaderHeight=28;grid.OptionsView.AutoScaleBands=false;grid.Appearance.RecordValue.Font=new Font("Segoe UI",12);grid.Appearance.RecordValue.Options.UseFont=true;
  for(int i=0;i<32;i++){var row=new EditorRow();row.Properties.FieldName="C"+i;row.Properties.Caption="Row "+i;row.Height=26;grid.Rows.Add(row);}host.Controls.Add(grid);host.Show();Pump(grid);
  presentation.GetMethod("Configure",F).Invoke(null,new object[]{grid,"Regression/Pointer/"+Guid.NewGuid()});Pump(grid);Zoom(grid,50);Check(Percent(grid)==60,"VGrid legacy50% request clamps to60%");Readable(grid,grid.Appearance.RecordValue.GetFont(),grid.Rows[0].Height,"VGrid");
  Check(grid.RecordWidth>=74,"VGrid records respect the eight-point text floor");
  Zoom(grid,100);grid.TopVisibleRowIndex=4;grid.LeftVisibleRecord=5;Pump(grid);
  foreach(int target in new[]{110,90,150,100}){
   var point=new Point(820,350);var hit=grid.CalcHitInfo(point);Check(hit.Row!=null&&hit.RecordIndex>=0,"VGrid pointer starts over data");var before=grid.ViewInfo.GetRowValueInfo(hit.Row,hit.RecordIndex,Math.Max(0,hit.CellIndex)).Bounds;
   At(grid,target,point);var info=grid.ViewInfo.GetRowValueInfo(hit.Row,hit.RecordIndex,Math.Max(0,hit.CellIndex));Near(before,info==null?Rectangle.Empty:info.Bounds,point,true,"VGrid "+target);
  }
  Zoom(grid,60);int h=grid.Rows[0].Height;Zoom(grid,100);Zoom(grid,60);Check(grid.Rows[0].Height==h,"VGrid minimum row height does not grow after reset");Wheel(grid,()=>grid.LeftVisibleRecord,"VGrid");
 }}
 static void Tree(){using(var host=new QuietForm{Opacity=0,ShowInTaskbar=false,ClientSize=new Size(1000,650)})using(var tree=new TreeList()){
  tree.Dock=DockStyle.Fill;tree.OptionsView.AutoWidth=false;tree.RowHeight=26;tree.Appearance.Row.Font=new Font("Segoe UI",12);tree.Appearance.Row.Options.UseFont=true;
  for(int c=0;c<32;c++){var col=tree.Columns.Add();col.Caption="C"+c;col.FieldName="C"+c;col.Visible=true;col.Width=110;col.MinWidth=20;}
  for(int r=0;r<100;r++)tree.AppendNode(Enumerable.Range(0,32).Select(c=>(object)("R"+r+"C"+c)).ToArray(),null);tree.Columns[0].Fixed=DevExpress.XtraTreeList.Columns.FixedStyle.Left;
  host.Controls.Add(tree);host.Show();Pump(tree);presentation.GetMethod("Configure",F).Invoke(null,new object[]{tree,"Regression/Pointer/"+Guid.NewGuid()});Pump(tree);Zoom(tree,50);Check(Percent(tree)==60,"Tree legacy50% request clamps to60%");Readable(tree,tree.Appearance.Row.GetFont(),tree.RowHeight,"Tree");
  Check(tree.Columns[1].Width>=74,"Tree columns respect the eight-point text floor");
  Zoom(tree,100);tree.TopVisibleNodeIndex=12;tree.LeftCoord=350;Pump(tree);
  foreach(int target in new[]{110,90,150,100}){
   var point=new Point(820,350);var hit=tree.CalcHitInfo(point);Check(hit.Node!=null&&hit.Column!=null,"Tree pointer starts over data");var before=tree.ViewInfo.RowsInfo[hit.Node].Cells.First(c=>c.Column==hit.Column).Bounds;
   At(tree,target,point);var row=tree.ViewInfo.RowsInfo[hit.Node];var cell=row==null?null:row.Cells.FirstOrDefault(c=>c.Column==hit.Column);Near(before,cell==null?Rectangle.Empty:cell.Bounds,point,false,"Tree "+target);
  }
  Zoom(tree,60);int h=tree.RowHeight;Zoom(tree,100);Zoom(tree,60);Check(tree.RowHeight==h,"Tree minimum row height does not grow after reset");Wheel(tree,()=>tree.LeftCoord,"Tree");
 }}
 [STAThread]static int Main(string[] args){try{
  Application.EnableVisualStyles();Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));presentation=app.GetType("Abovo.GridPresentation");Standard();Vertical();Tree();Console.WriteLine("PASS TOTAL="+checks);return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
}
