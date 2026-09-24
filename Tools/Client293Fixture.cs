using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraVerticalGrid;
using DevExpress.XtraVerticalGrid.Rows;
using DevExpress.Spreadsheet;

class Client293Fixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static int count;
 [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h,int m,IntPtr w,IntPtr l);
 static object Field(object o,string n){var f=o.GetType().GetField(n,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(n,F).GetValue(o,null);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("PASS "+(++count)+" "+text);}
 static IEnumerable<Control> Controls(Control c){foreach(Control d in c.Controls){yield return d;foreach(var n in Controls(d))yield return n;}}
 static IEnumerable<BaseRow> Rows(IEnumerable list){foreach(BaseRow r in list){yield return r;foreach(var n in Rows(r.ChildRows))yield return n;}}
 static void Pump(){Application.DoEvents();}
 static void Picture(Control c,string path){using(var b=new Bitmap(c.Width,c.Height)){c.DrawToBitmap(b,new Rectangle(Point.Empty,b.Size));b.Save(path);}}
 static void Click(Control c,int x,int y){var p=(IntPtr)((y<<16)|x);SendMessage(c.Handle,0x201,(IntPtr)1,p);SendMessage(c.Handle,0x202,IntPtr.Zero,p);Pump();}
 static void Type(Control c,string s){var target=Controls(c).OfType<TextBoxBase>().FirstOrDefault() as Control ?? c;Console.WriteLine("Typing into "+target.GetType().FullName);foreach(char ch in s)SendMessage(target.Handle,0x102,(IntPtr)ch,IntPtr.Zero);Pump();}
 static List<DateTime> Dates(Assembly a,DateTime start,int months,DateTime end){dynamic r=Activator.CreateInstance(a.GetType("Abovo.FundingScheduleRequest"));r.StartDate=start;r.IntervalMonths=months;r.EndDate=end;return ((IEnumerable)a.GetType("Abovo.FundingScheduleGenerator").GetMethod("Generate").Invoke(null,new object[]{r,null})).Cast<object>().Select(x=>(DateTime)Field(x,"DateToAdd")).ToList();}
 [STAThread] static int Main(string[] args){try {
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  Check(Dates(app,new DateTime(2026,9,24),12,new DateTime(2049,9,24)).SequenceEqual(Enumerable.Range(2026,24).Select(y=>new DateTime(y,9,24))),"Alex annual September 24 example, inclusive through 2049");
  Check(Dates(app,new DateTime(2026,8,31),6,new DateTime(2027,8,31)).SequenceEqual(new[]{new DateTime(2026,8,31),new DateTime(2027,2,28),new DateTime(2027,8,31)}),"Semi-annual short-month adjustment does not drift");
  Check(Dates(app,new DateTime(2028,1,31),1,new DateTime(2028,3,31))[1]==new DateTime(2028,2,29),"Monthly leap year adjustment");
  Check(Dates(app,new DateTime(2026,1,31),3,new DateTime(2026,7,30)).Count==2,"Quarterly excludes dates after end date");
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  var copy=Path.Combine(args[1],"private-client293.xlsb");File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  if(result.BError)throw new Exception(result.StringReturn);
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)result.IntegerReturn);IWorkbook wb=model.WB;
  try {
   var targets=(Array)app.GetType("Abovo.FundingScheduleWriter").GetField("Targets").GetValue(null);
   using(var dialog=(Form)Activator.CreateInstance(app.GetType("FundingScheduleForm"),new object[]{(int)model.ModelID,targets,null,-1})) {
    dialog.Opacity=0;dialog.ShowInTaskbar=false;dialog.Show();Pump();
    var start=(DateEdit)Field(dialog,"StartChoice");var end=(DateEdit)Field(dialog,"EndDateChoice");var interval=(ComboBoxEdit)Field(dialog,"IntervalChoice");
    start.DateTime=new DateTime(2026,9,24);end.DateTime=new DateTime(2028,9,24);interval.SelectedIndex=2;
    Call(dialog,"RefreshPreviewAsync");Pump();
    Check(interval.Properties.Items.Count==4&&interval.Properties.Items[2].ToString()=="Semi-annually","Four requested intervals only");
    Check(!Controls(dialog).OfType<Control>().Any(c=>c.Text.Contains("Holidays")||c.Text=="Date rule"||c.Text=="Occurrences"),"No holiday, date-rule or occurrence controls");
    var view=(GridView)Field(dialog,"View");Check(view.VisibleColumns.Count==1&&view.VisibleColumns[0].FieldName=="DateToAdd","Single-column date preview");
    Check(((List<DateTime>)Field(dialog,"DatesToApply")).Count==5,"Automatic preview uses inclusive end date");
    ((CheckEdit)Field(dialog,"PopulateChoice")).Checked=true;Call(dialog,"RefreshPreviewAsync");
    var figure=(SpinEdit)Field(dialog,"FigureChoice");figure.Focus();figure.SelectAll();Type(figure,"123456.789");
    Console.WriteLine("FIGURE text="+figure.Text+" value="+figure.Value);Check(figure.Value==123456.789m,"Real character input accepts multi-digit decimal figure");
    figure.SelectAll();Type(figure,"-98765.4321");Check(figure.Value==-98765.4321m,"Negative multi-digit decimal entry is retained");figure.SelectAll();Type(figure,"123456.789");
    var button=(SimpleButton)Field(dialog,"ApplyButton");Check(button.Enabled,"Typing a figure does not disable the valid date preview");
    Picture(dialog,Path.Combine(args[1],"scheduler.png"));button.PerformClick();Check(dialog.DialogResult==DialogResult.OK&&(decimal?)Field(dialog,"FixedFigure")==123456.789m,"One OK accepts the typed figure and dates");
   }
   using(var host=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})) {
    host.Opacity=0;host.ShowInTaskbar=false;host.ClientSize=new Size(1700,1000);host.Show();Pump();
    Call(host,"ShowInterface",(int)model.ModelID,0,false,"None",null,-1);Pump();var dit=(Control)Field(host,"ActiveInterface");
    var company=Controls(dit).OfType<TextEdit>().First(e=>e.Text==wb.DefinedNames.GetDefinedName("SelectTrust").Range[0,0].DisplayText);
    Check(company.Width<=company.MaximumSize.Width&&company.MaximumSize.Width>400&&company.Width<dit.Width,"Company editor is usefully wide but bounded");Picture(dit,Path.Combine(args[1],"company.png"));
    Call(host,"ShowInterface",(int)model.ModelID,1,false,"None",null,-1);Pump();dit=(Control)Field(host,"ActiveInterface");
    var stock=Controls(dit).OfType<GridControl>().Select(g=>g.MainView).OfType<GridView>().First(v=>v.Columns.Count>4);
    var description=stock.Columns[0];Check(Convert.ToInt32(Field(description.Tag,"MinimumWidthChars"))==42&&description.Width>=description.MinWidth&&description.MinWidth>200,"Stock minimum width reaches pivoted data and native column");
    var dummy=stock.Columns.Cast<DevExpress.XtraGrid.Columns.GridColumn>().FirstOrDefault(c=>(bool)Field(c.Tag,"IsDummyColumn"));
    Check(dummy!=null&&dummy.Visible&&dummy.VisibleIndex==5&&dummy.Width>=40,"Stock separator is visible immediately after Current Stock Numbers");Picture(dit,Path.Combine(args[1],"stock.png"));
    Call(host,"ShowInterface",(int)model.ModelID,33,false,"None",null,1);Pump();dit=(Control)Field(host,"ActiveInterface");
    foreach(var tabs in Controls(dit).OfType<DevExpress.XtraTab.XtraTabControl>())tabs.SelectedTabPageIndex=1;
    Call(dit,"BuildSection",1,false,false);Pump();
    Console.WriteLine("FUNDING CSID="+Field(dit,"CSID")+" grids="+Controls(dit).OfType<VGridControl>().Count());
    var grid=Controls(dit).OfType<VGridControl>().First(g=>g.Visible&&Rows(g.Rows).Any(r=>r.Properties.Caption.Replace("\n"," ").Contains("Facility")));
    Check(!grid.OptionsView.ShowRecordHeaders,"Extra Funding record-header strip removed");Check(!grid.OptionsBehavior.RecordsMouseWheel,"Funding ordinary wheel configured for vertical rows");
    var facilityRow=Rows(grid.Rows).First(r=>r.Properties.Caption.Contains("Facility")&&r.Properties.Caption.Contains("Name"));
    var repo=(DevExpress.XtraEditors.Repository.RepositoryItemComboBox)facilityRow.Properties.RowEdit;
    var expected=wb.DefinedNames.GetDefinedName("Facility").Range.ExistingCells.Where(c=>!c.Value.IsEmpty).Select(c=>c.DisplayText).ToList();
    Check(expected.All(s=>repo.Items.Cast<object>().Any(x=>x.ToString()==s)),"Facility Name choices are sourced from Facility, not Funders");
    grid.FocusedRow=facilityRow;grid.FocusedRecord=1;Call(dit,"SetClipboardTarget",grid);Call(dit,"UpdateFundingScheduleMenu");
    var menu=(ToolStripMenuItem)Field(dit,"FundingScheduleMenuItem");Check(menu.Available&&menu.Text=="Add schedule…","Funding context action is available for a loan cell");
    object[] context={false};var selected=dit.GetType().GetMethod("ContextFundingFacility",F).Invoke(dit,context);
    var cell=(Cell)Call(dit,"InputSourceCell",grid.DataSource,grid.FocusedRecord,(int)Call(dit,"GetVGridColumnIndex",facilityRow,0));
    Check((int)Field(selected,"ColumnIndex")==cell.ColumnIndex,"Context action targets clicked loan's actual worksheet column");
    var tab=(DevExpress.XtraTab.XtraTabControl)Field(dit,"XtraTabControlNewGIT");
    var filter=Activator.CreateInstance(app.GetType("DataInterfaceTemplate+ScrollRediverter"),new object[]{dit,tab});
    var oldCursor=Cursor.Position;
    try {
     host.Activate();grid.Focus();grid.RecordWidth=Math.Max(200,grid.Width/3);grid.LeftVisibleRecord=0;Pump();
     Cursor.Position=grid.PointToScreen(new Point(Math.Min(grid.Width-20,400),30));
     var wParam=(IntPtr)unchecked((int)((uint)(ushort)(short)-120<<16)|8);
     object[] wheel={Message.Create(grid.Handle,0x20A,wParam,IntPtr.Zero)};
     bool wheelHandled=(bool)Call(filter,"PreFilterMessage",wheel);Console.WriteLine("WHEEL ctrl handled="+wheelHandled+" left="+grid.LeftVisibleRecord+" records="+grid.RecordCount+" width="+grid.Width+" recwidth="+grid.RecordWidth+" point="+Cursor.Position+" rect="+tab.SelectedTabPage.RectangleToScreen(tab.SelectedTabPage.ClientRectangle)+" enabled="+host.Enabled+" active="+(Form.ActiveForm==host));
     Check(wheelHandled&&grid.LeftVisibleRecord>0,"Ctrl+wheel message scrolls loan records horizontally");
     int left=grid.LeftVisibleRecord;int top=grid.TopVisibleRowIndex;var page=tab.SelectedTabPage;int scroll=page.VerticalScroll.Value;
     wheel[0]=Message.Create(grid.Handle,0x20A,(IntPtr)unchecked((int)((uint)(ushort)(short)-120<<16)),IntPtr.Zero);
     Check((bool)Call(filter,"PreFilterMessage",wheel),"Ordinary wheel is routed to the current Funding view");Pump();
     Console.WriteLine("WHEEL top="+top+"->"+grid.TopVisibleRowIndex+" page="+scroll+"->"+page.VerticalScroll.Value+" left="+left+"->"+grid.LeftVisibleRecord);
     Check(grid.LeftVisibleRecord==left&&(grid.TopVisibleRowIndex>top||page.VerticalScroll.Value>scroll),"Ordinary wheel scrolls vertically without moving loan columns");
    }finally{Cursor.Position=oldCursor;}
    // Exercise the exact repository installed in the Funding grid, on a native edit control.
    using(var popupHost=new Form())using(var combo=new ComboBoxEdit()){
     popupHost.Opacity=0;popupHost.ShowInTaskbar=false;combo.Properties.Assign(repo);combo.Width=250;popupHost.Controls.Add(combo);popupHost.Show();Pump();combo.Focus();Pump();
     Check(!combo.IsPopupOpen,"Focusing the combo no longer opens its popup before the click");
     Click(combo,combo.Width-8,combo.Height/2);Check(combo.IsPopupOpen,"First dropdown-button click leaves the popup open");combo.ClosePopup();
     var helper=app.GetType("Abovo.InplaceEditorFormatting");var key=new KeyEventArgs(Keys.Shift|Keys.Down);helper.GetMethod("OpenPopupForShortcut").Invoke(null,new object[]{combo,key});
     Check(combo.IsPopupOpen&&key.Handled,"Shift+Down opens the dropdown without navigating");combo.ClosePopup();combo.Properties.ReadOnly=true;
     key=new KeyEventArgs(Keys.Shift|Keys.Down);Check(!(bool)helper.GetMethod("OpenPopupForShortcut").Invoke(null,new object[]{combo,key})&&!combo.IsPopupOpen,"Popup shortcut respects read-only inputs");
    }
    host.Close();
   }
  } finally {files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});}
  Console.WriteLine("PASS "+count+" assertions");return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
