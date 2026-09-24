using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;

class PatternEditabilityFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static object Field(object o,string n){var f=o.GetType().GetField(n,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(n,F).GetValue(o,null);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static int count;
 static void Check(bool value,string message){if(!value)throw new Exception(message);count++;Console.WriteLine("PASS "+message);}
 static IEnumerable<Control> Controls(Control c){foreach(Control child in c.Controls){yield return child;foreach(var d in Controls(child))yield return d;}}
 [STAThread] static int Main(string[] args){try {
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  string copy=Path.Combine(args[1],"private-pattern.xlsb");File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  if(result.BError)throw new Exception("Open failed");
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)result.IntegerReturn);
  try {
   IWorkbook w=(IWorkbook)model.WB;
   foreach(string n in new[]{"StockType","Rep_Sto_2","Rep_Sto_3","CurrStNo"}){
    var r=w.DefinedNames.GetDefinedName(n).Range;
    Console.WriteLine(n+" "+r.GetReferenceA1());
    for(int i=0;i<Math.Min(3,r.ColumnCount);i++) {var c=r[0,i];Console.WriteLine(c.GetReferenceA1()+" value="+c.DisplayText+" pattern="+c.Fill.PatternType+" locked="+c.Protection.Locked+" bg="+c.Fill.BackgroundColor+" font="+c.Font.Color);}
   }
   using(var form=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})){
    form.Opacity=0;form.ShowInTaskbar=false;form.ClientSize=new Size(1700,1000);form.Show();Application.DoEvents();
    Call(form,"ShowInterface",(int)model.ModelID,1,false,"None",null,-1);Application.DoEvents();
    var dit=(Control)Field(form,"ActiveInterface");
    foreach(var v in Controls(dit).OfType<GridControl>().Select(g=>g.MainView).OfType<GridView>()){
     var data=Field(v.Tag,"DataSet");Console.WriteLine("GRID "+Field(data,"Name"));
     var rows=(Array)Field(data,"DataRows");if(rows.Length==0)continue;var cells=(Array)Field(rows.GetValue(0),"DataCells");
     for(int i=0;i<v.Columns.Count;i++){var col=v.Columns[i];var dp=cells.GetValue(col.AbsoluteIndex);Console.WriteLine(i+" "+col.Caption+" index="+col.AbsoluteIndex+" tagRules="+Field(col.Tag,"HasRules")+" tagRO="+Field(col.Tag,"IsReadOnly")+" dp="+(dp==null?"null":Field(dp,"SourceAddress")+" lock="+Field(dp,"IsLocked")));}
     Check((bool)Field(cells.GetValue(1),"IsLocked"),"Empty Stock description locks dependent category");
     Check(!(bool)Call(dit,"CanPasteToDataPoint",data,0,1),"Paste refuses locked category");
     v.FocusedRowHandle=0;v.FocusedColumn=v.Columns[1];v.ShowEditor();Check(v.ActiveEditor==null,"Native editor refuses locked category");
     dynamic change=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));change.ModelID=(int)model.ModelID;change.WSName="Stock Assumptions";change.CellAddress="E5";change.DataFormat="S";change.Description="Stock pattern regression";change.ChangedValue="Pattern test";
     Check(model.ChangeManager.ProcessChange(change).BSuccess,"Description edit through ChangeManager succeeds");Application.DoEvents();
     Check(!(bool)Field(cells.GetValue(1),"IsLocked"),"Description edit unlocks dependent category without reopening");
     Check((bool)Call(dit,"CanPasteToDataPoint",data,0,1),"Paste allows newly unlocked category");
     v.FocusedRowHandle=0;v.FocusedColumn=v.Columns[1];v.ShowEditor();Check(v.ActiveEditor!=null,"Native editor opens newly unlocked category");v.CloseEditor();
     Check(model.ChangeManager.Undo().BSuccess,"Undo description succeeds");Application.DoEvents();
     Check((bool)Field(cells.GetValue(1),"IsLocked"),"Undo relocks dependent category");
     Check(!(bool)Call(dit,"CanPasteToDataPoint",data,0,1),"Paste refuses relocked category");
     Check(model.ChangeManager.Redo().BSuccess,"Redo description succeeds");Application.DoEvents();
     Check(!(bool)Field(cells.GetValue(1),"IsLocked"),"Redo unlocks dependent category");
     change.ChangedValue=null;Check(model.ChangeManager.ProcessChange(change).BSuccess,"Clearing description succeeds");Application.DoEvents();
     Check((bool)Field(cells.GetValue(1),"IsLocked"),"Clearing description relocks category");
     Color paintedBack=Color.Empty,paintedFore=Color.Empty;
     v.CustomDrawCell+=(s,e)=>{if(e.RowHandle==0&&e.Column==v.Columns[1]){paintedBack=e.Appearance.BackColor;paintedFore=e.Appearance.ForeColor;}};
     v.ClearSelection();v.FocusedRowHandle=1;v.FocusedColumn=v.Columns[0];v.ClearSelection();v.GridControl.Refresh();Application.DoEvents();
     using(var bitmap=new Bitmap(v.GridControl.Width,v.GridControl.Height)){v.GridControl.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(args[1],"stock-pattern.png"));}
     Check(paintedBack.ToArgb()==Color.Lavender.ToArgb(),"Pattern-locked cell retains shared unavailable-input background");
     Check(paintedFore.ToArgb()!=Color.White.ToArgb()&&!paintedFore.IsEmpty,"Unavailable-input text is readable, not white");
     Check(w.Worksheets["Stock Assumptions"].Cells["E6"].Fill.PatternType==PatternType.DarkGray,"Rendering does not alter workbook fill pattern");
    }
    form.Close();
   }
  }finally{files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});}
  Console.WriteLine("PASS "+count+" assertions");return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
}
