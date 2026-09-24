using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraTab;
using DevExpress.XtraTreeList;

// Native, hidden UI against an unsaved private model copy; no originals are saved.
class AnalyserZoom295Fixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static int assertions;
 static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("PASS "+(++assertions)+" "+text);}
 static object Field(object o,string name){var f=o.GetType().GetField(name,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(name,F).GetValue(o,null);}
 static void Pump(){Application.DoEvents();}
 static void Picture(Control c,string path){using(var b=new Bitmap(c.Width,c.Height)){c.DrawToBitmap(b,new Rectangle(Point.Empty,b.Size));b.Save(path);}}
 static void Zoom(Type p,Control c,int n){p.GetMethod("SetZoom",F).Invoke(null,new object[]{c,n,false});Pump();c.Refresh();Pump();}
 [STAThread] static int Main(string[] args){try{
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var dllPath=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(dllPath)?Assembly.LoadFrom(dllPath):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  var path=Path.Combine(args[1],"analyser295.xlsb");File.Copy(args[2],path);
  var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{path,new FileInfo(path),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!result.BError,"Private model opened");int model=(int)result.IntegerReturn;
  var presentation=app.GetType("Abovo.GridPresentation");
  using(var host=new Form{Opacity=0,ShowInTaskbar=false,ClientSize=new Size(1600,950)})
  using(var analyser=(Control)Activator.CreateInstance(app.GetType("BPIncomeExpenditureAnalyserV2"),new object[]{model,null})){
   host.Controls.Add(analyser);analyser.Dock=DockStyle.Fill;host.Show();Pump();
   var tabs=(XtraTabControl)Field(analyser,"XtraTabControlAnalyser");
   for(int tab=0;tab<2;tab++){
    tabs.SelectedTabPageIndex=tab;Pump();
    var wrap=Field(analyser,tab==0?"WrapCG_SOCI":"WrapCG_CF");var view=(GridView)Field(wrap,"WrappedGridView");var grid=view.GridControl;
    presentation.GetMethod("Configure",F).Invoke(null,new object[]{grid,"Regression/analyser295/"+Guid.NewGuid()});
    Zoom(presentation,grid,100);float size=view.Appearance.GroupRow.Font.SizeInPoints;
    var info=(GridViewInfo)view.GetViewInfo();int baseHeight=info.GroupFooterHeight;
    Console.WriteLine("FOOTER font="+view.Appearance.GroupFooter.GetFont()+" group="+view.Appearance.GroupRow.GetFont()+" row="+baseHeight+" cell="+info.GroupFooterCellHeight);
    var footerRows=new System.Collections.Generic.Dictionary<int,Rectangle>();
    view.CustomDrawRowFooter+=(s,e)=>{footerRows[e.RowHandle]=e.Bounds;};
    view.CustomDrawRowFooterCell+=(s,e)=>{Rectangle row;if(footerRows.TryGetValue(e.RowHandle,out row)&&e.Bounds.Bottom>row.Bottom)throw new Exception("Footer figure extends below its row: "+e.Bounds+" vs "+row);};
    Check(info.GroupFooterCellHeight<=baseHeight,"Total cell height fits its compact row");
    Check(view.Columns["ItemDesc"].Width<=grid.ClientSize.Width*.41,"Description leaves most viewport for figures");
    for(int repeat=0;repeat<3;repeat++){
     Zoom(presentation,grid,50);Check(view.Appearance.GroupRow.Font.SizeInPoints<size,"Custom-drawn group text shrinks");int small=((GridViewInfo)view.GetViewInfo()).GroupFooterHeight;
     Zoom(presentation,grid,200);Check(((GridViewInfo)view.GetViewInfo()).GroupFooterHeight>small,"Total height follows zoom");
     Zoom(presentation,grid,100);Check(Math.Abs(view.Appearance.GroupRow.Font.SizeInPoints-size)<.05,"Repeated zoom restores exact group font");
     Check(Math.Abs(((GridViewInfo)view.GetViewInfo()).GroupFooterHeight-baseHeight)<=1,"Repeated zoom restores total height");
    }
    for(int repeat=0;repeat<5;repeat++)analyser.GetType().GetMethod("ApplyDescriptionColumnBestFit",F).Invoke(null,new object[]{view});
    Check(view.Columns["ItemDesc"].Width<=grid.ClientSize.Width*.41,"Repeated refresh does not grow description column");
    var menu=grid.ContextMenuStrip;var menuOpening=typeof(ContextMenuStrip).GetMethod("OnOpening",F);
    for(int repeat=0;repeat<3;repeat++){var e=new CancelEventArgs();menuOpening.Invoke(menu,new object[]{e});Check(!e.Cancel&&menu.Items.Find("SummitGridZoom",false).Length==1,"Rebuilt analyser menu retains one Zoom action");}
    var zoomItem=(ToolStripMenuItem)menu.Items.Find("SummitGridZoom",false)[0];
    Check(zoomItem.DropDownItems.OfType<ToolStripMenuItem>().Any(x=>x.Text=="Set zoom"),"Explicit zoom percentages available");
    view.RowHeight=1;Zoom(presentation,grid,50);
    zoomItem.DropDownItems.OfType<ToolStripMenuItem>().First(x=>x.Text.StartsWith("Reset zoom")).PerformClick();Pump();
    Check(view.RowHeight==-1&&(int)presentation.GetMethod("ZoomPercent",F).Invoke(null,new object[]{grid})==100,"Reset restores automatic row height after accidental shrink");
    Picture(grid,Path.Combine(args[1],"analyser-"+tab+"-100.png"));
   }
   tabs.SelectedTabPageIndex=2;Pump();var statement=Field(analyser,"BalanceSheetView");var tree=(TreeList)Field(statement,"Tree");
   Zoom(presentation,tree,100);float treeSize=tree.Appearance.Row.Font.SizeInPoints;
   for(int i=0;i<3;i++){Zoom(presentation,tree,50);Zoom(presentation,tree,200);Zoom(presentation,tree,100);}
   Check(Math.Abs(tree.Appearance.Row.Font.SizeInPoints-treeSize)<.05,"Balance Sheet tree font resets repeatedly");
   Check(tree.ContextMenuStrip.Items.Find("SummitGridZoom",false).Length==1,"Balance Sheet retains zoom menu");
   Picture(tree,Path.Combine(args[1],"balance-100.png"));host.Close();
  }
  files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{model});
  Console.WriteLine("PASS "+assertions+" assertions");return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
