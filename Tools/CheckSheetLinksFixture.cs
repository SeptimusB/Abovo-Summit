using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using DevExpress.Spreadsheet;
using DevExpress.Utils;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;

class CheckSheetLinksFixture {
 const BindingFlags F = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static;
 [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int m, IntPtr w, IntPtr l);
 static int count;
 static object Get(object o,string n) { var f=o.GetType().GetField(n,F); return f!=null?f.GetValue(o):o.GetType().GetProperty(n,F).GetValue(o,null); }
 static object Call(object o,string n,params object[] a) { return o.GetType().GetMethod(n,F).Invoke(o,a); }
 static void Check(bool value,string text) { if(!value)throw new Exception(text); Console.WriteLine("PASS "+(++count)+" "+text); }
 static IEnumerable<Control> Controls(Control c) { foreach(Control child in c.Controls) { yield return child; foreach(var d in Controls(child)) yield return d; } }
 static void Pump() { Application.DoEvents(); }
 static void Click(GridControl grid,GridView view,int row,DevExpress.XtraGrid.Columns.GridColumn column) {
  view.MakeRowVisible(row); view.MakeColumnVisible(column); Pump();
  grid.FindForm().Activate(); grid.Focus(); Pump();
  for(int y=5;y<grid.ClientSize.Height;y+=5) for(int x=5;x<grid.ClientSize.Width;x+=5) {
   var hit=view.CalcHitInfo(new Point(x,y)); if(!hit.InRowCell||hit.RowHandle!=row||hit.Column!=column)continue;
   Console.WriteLine("CLICK row="+row+" field="+column.FieldName+" x="+x+" y="+y+" selectedBefore="+view.GetSelectedCells().Length);
   var savedCursor=Cursor.Position; var p=(IntPtr)((y<<16)|x);
   try { Cursor.Position=grid.PointToScreen(new Point(x,y)); SendMessage(grid.Handle,0x200,IntPtr.Zero,p); SendMessage(grid.Handle,0x201,(IntPtr)1,p); SendMessage(grid.Handle,0x202,IntPtr.Zero,p); Pump(); }
   finally { Cursor.Position=savedCursor; }
   Console.WriteLine("CLICK focused="+view.FocusedRowHandle+" field="+view.FocusedColumn.FieldName+" selectedAfter="+view.GetSelectedCells().Length); return;
  }
  throw new Exception("Link cell is not visible for native mouse click");
 }
 [STAThread] static int Main(string[] args) { try {
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=> { var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll"); return File.Exists(p)?Assembly.LoadFrom(p):null; };
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager"); files.GetMethod("Initialise").Invoke(null,new object[]{null});
  string copy=Path.Combine(args[1],"private-checksheet-links.xlsb"); File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel"); dynamic opened=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!opened.BError,"Private workbook opens");
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)opened.IntegerReturn);
  IWorkbook book=model.WB; var sheet=book.Worksheets["Check Sheet"];
  using(var host=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,-1,"Normal"})) {
   host.Opacity=0; host.ShowInTaskbar=false; host.ClientSize=new Size(1700,1050); host.Show(); Pump();
   Call(host,"ShowInterface",(int)model.ModelID,0,false,"None",null,0); Pump();
   var dit=(Control)Get(host,"ActiveInterface"); Call(dit,"BuildSection",2,false,false);
   var tabs=(DevExpress.XtraTab.XtraTabControl)Get(dit,"XtraTabControlNewGIT"); tabs.SelectedTabPageIndex=2; Pump();
   var grid=Controls(dit).OfType<GridControl>().First(g=>g.GetType().Name=="ReadOnlyMappedTableGrid");
   var view=(GridView)grid.MainView; var summit=view.Columns["SummitInterface"];
   Console.WriteLine("EDITOR mode="+view.OptionsBehavior.EditorShowMode+" editable="+view.OptionsBehavior.Editable+" multi="+view.OptionsSelection.MultiSelectMode);
   Check(summit!=null&&summit.Visible&&summit.Caption=="Summit interface","Explicit Summit destination column remains visible");
   dynamic mapping=model.WBStructure.GroupStructures[0].ChildStructures[0].InterfaceSections[2].IElements[0].MappedTable;
   int unique=-1, ambiguous=-1, linked=0; object uniqueRoute=null; int ambiguousCount=0;
   var resolver=grid.GetType().GetMethod("FindInterfaces");
   for(int row=0;row<57;row++) {
    string target=sheet.Cells[row+6,7].DisplayText.Trim(); if(target.Length==0)continue;
    var routes=(IList)resolver.Invoke(null,new object[]{model.WBStructure,target,mapping.InterfaceLinks});
    if(routes.Count==0)continue; linked++;
    Check(view.GetRowCellDisplayText(row,summit).Length>0,"Destination is shown for "+target);
    if(routes.Count==1&&unique<0) { unique=row; uniqueRoute=routes[0]; }
    if(routes.Count>1&&ambiguous<0) { ambiguous=row; ambiguousCount=routes.Count; }
   }
   Check(linked>=20&&unique>=0&&ambiguous>=0,"Representative unique and multiple Summit destinations are available");
   var caption=sheet.Cells[unique+6,6]; string oldFormula=caption.FormulaInvariant; var oldValue=caption.Value;
   try {
    caption.Value=CellValue.Empty; Call(grid,"RefreshData");
    Check(view.GetRowCellDisplayText(unique,summit).Contains((string)Get(uniqueRoute,"LinkTip")),"Blank Excel hyperlink caption does not erase a valid Summit destination");
    Check(view.GetRowCellDisplayText(unique,view.Columns["C6"])==sheet.Cells[unique+6,7].DisplayText.Trim(),"Explicit worksheet link is also retained");
    view.ClearSelection(); view.SelectCell(unique,view.Columns[0]); view.SelectCell(unique,summit);
    var argsClick=new RowCellClickEventArgs(DXMouseEventArgs.GetMouseArgs(new MouseEventArgs(MouseButtons.Left,1,0,0,0)),unique,summit);
    Call(grid,"ClickCell",view,argsClick);
    Check(Object.ReferenceEquals(dit,Get(host,"ActiveInterface")),"Finishing multi-cell selection does not navigate");
    view.ClearSelection(); Click(grid,view,ambiguous,summit);
    var menu=(ContextMenuStrip)Get(grid,"navigationMenu");
    Console.WriteLine("MENU visible="+menu.Visible+" items="+menu.Items.Count+" expected="+ambiguousCount);
    Check(menu.Visible&&menu.Items.Count==ambiguousCount,"Ambiguous destinations still offer an explicit choice after one click"); menu.Close();
    view.ClearSelection(); Click(grid,view,unique,summit);
    var destination=(Control)Get(host,"ActiveInterface");
    Check(!Object.ReferenceEquals(dit,destination)&&!menu.Visible,"One native mouse click opens the unique Summit destination directly");
    var link=Get(destination,"ActiveLinkElement");
    Check((string)Get(link,"LinkData")== (string)Get(uniqueRoute,"LinkData")&&(int)Get(link,"LinkReturnID")==0&&(int)Get(link,"LinkReturnGroup")==0,"Navigation retains correct destination and return-to-Check-Sheet context");
   } finally { if(string.IsNullOrEmpty(oldFormula))caption.Value=oldValue; else caption.FormulaInvariant=oldFormula; }
   host.Close();
  }
  // Only the private in-memory fixture was changed. No Save, recovery or workbook close-save.
  Console.WriteLine("PASS: "+count+" Check Sheet link checks"); return 0;
 } catch(Exception e) { Console.Error.WriteLine(e); return 1; } }
}
