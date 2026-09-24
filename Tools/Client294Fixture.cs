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

class Client294Fixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static int count;
 [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h,int m,IntPtr w,IntPtr l);
 [DllImport("user32.dll")] static extern bool SetKeyboardState(byte[] state);
 [DllImport("user32.dll")] static extern bool GetKeyboardState(byte[] state);
 static object Field(object o,string n){var f=o.GetType().GetField(n,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(n,F).GetValue(o,null);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("PASS "+(++count)+" "+text);}
 static IEnumerable<Control> Controls(Control c){foreach(Control d in c.Controls){yield return d;foreach(var n in Controls(d))yield return n;}}
 static IEnumerable<BaseRow> Rows(IEnumerable list){foreach(BaseRow r in list){yield return r;foreach(var n in Rows(r.ChildRows))yield return n;}}
 static void Pump(){Application.DoEvents();}
 static void Picture(Control c,string path){using(var b=new Bitmap(c.Width,c.Height)){c.DrawToBitmap(b,new Rectangle(Point.Empty,b.Size));b.Save(path);}}
 static void Click(Control c,int x,int y){var p=(IntPtr)((y<<16)|x);SendMessage(c.Handle,0x201,(IntPtr)1,p);SendMessage(c.Handle,0x202,IntPtr.Zero,p);Pump();}
 static Point CellPoint(VGridControl g,BaseRow row,int record){for(int y=5;y<g.Height-20;y+=5)for(int x=g.RowHeaderWidth+10;x<Math.Min(g.Width,g.RowHeaderWidth+g.RecordWidth*3);x+=20){var h=g.CalcHitInfo(new Point(x,y));if(h.Row==row&&h.RecordIndex==record)return new Point(x+5,y+5);}throw new Exception("Cell hit not found");}
 static void Drag(VGridControl g,Point from,Point to){var cursor=Cursor.Position;var keys=new byte[256];GetKeyboardState(keys);try{g.Focus();Cursor.Position=g.PointToScreen(from);var pressed=(byte[])keys.Clone();pressed[1]=128;SetKeyboardState(pressed);SendMessage(g.Handle,0x201,(IntPtr)1,(IntPtr)((from.Y<<16)|from.X));Console.WriteLine("DRAG down mode="+g.OptionsBehavior.EditorShowMode+" active="+g.ActiveEditor);for(int i=1;i<=12;i++){var p=new Point(from.X+(to.X-from.X)*i/12,from.Y+(to.Y-from.Y)*i/12);Cursor.Position=g.PointToScreen(p);SetKeyboardState(pressed);SendMessage(g.Handle,0x200,(IntPtr)1,(IntPtr)((p.Y<<16)|p.X));Pump();}Console.WriteLine("DRAG move="+g.GetSelectedCells().Count());SetKeyboardState(keys);SendMessage(g.Handle,0x202,IntPtr.Zero,(IntPtr)((to.Y<<16)|to.X));Pump();}finally{SetKeyboardState(keys);Cursor.Position=cursor;}}
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
  var copy=Path.Combine(args[1],"private-client294.xlsb");File.Copy(args[2],copy);
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
    Check(!Controls(dialog).OfType<Control>().Any(c=>c.Text.Contains("Holidays")||c.Text=="Date rule"),"No holiday or date-rule controls");
    var view=(GridView)Field(dialog,"View");Check(view.VisibleColumns.Count==1&&view.VisibleColumns[0].FieldName=="DateToAdd","Single-column date preview");
    Check(((List<DateTime>)Field(dialog,"DatesToApply")).Count==5,"Automatic preview uses inclusive end date");
    var endMode=(ComboBoxEdit)Field(dialog,"EndChoice");var occurrences=(SpinEdit)Field(dialog,"CountChoice");endMode.SelectedIndex=1;occurrences.Value=3;Call(dialog,"RefreshPreviewAsync");
    Check(((List<DateTime>)Field(dialog,"DatesToApply")).Count==3&&!end.Enabled&&occurrences.Enabled,"Occurrence alternative includes exactly three dates and disables end date");
    endMode.SelectedIndex=0;Call(dialog,"RefreshPreviewAsync");
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
    var description=stock.Columns[0];Check(Convert.ToInt32(Field(description.Tag,"MinimumWidthChars"))==21&&description.Width>=description.MinWidth&&description.MinWidth>100,"Stock minimum width reaches pivoted data and native column");
    var dummy=stock.Columns.Cast<DevExpress.XtraGrid.Columns.GridColumn>().FirstOrDefault(c=>(bool)Field(c.Tag,"IsDummyColumn"));
    Check(dummy!=null&&dummy.Visible&&dummy.VisibleIndex==5&&dummy.Width>=40,"Stock separator is visible immediately after Current Stock Numbers");Picture(dit,Path.Combine(args[1],"stock.png"));
    Call(host,"ShowInterface",(int)model.ModelID,33,false,"None",null,1);Pump();dit=(Control)Field(host,"ActiveInterface");
    foreach(var tabs in Controls(dit).OfType<DevExpress.XtraTab.XtraTabControl>())tabs.SelectedTabPageIndex=1;
    Call(dit,"BuildSection",1,false,false);Pump();
    Console.WriteLine("FUNDING CSID="+Field(dit,"CSID")+" grids="+Controls(dit).OfType<VGridControl>().Count());
    var grid=Controls(dit).OfType<VGridControl>().First(g=>g.Visible&&Rows(g.Rows).Any(r=>r.Properties.Caption.Replace("\n"," ").Contains("Facility")));
    var presentation=app.GetType("Abovo.GridPresentation");presentation.GetMethod("Configure",F).Invoke(null,new object[]{grid,"Regression/"+Guid.NewGuid()});
    Check(grid.OptionsView.ShowRecordHeaders,"Loan record header strip restored");Check(!grid.OptionsBehavior.RecordsMouseWheel,"Funding ordinary wheel configured for vertical rows");
    var facilityRow=Rows(grid.Rows).First(r=>r.Properties.Caption.Contains("Facility")&&r.Properties.Caption.Contains("Name"));
    var repo=(DevExpress.XtraEditors.Repository.RepositoryItemComboBox)facilityRow.Properties.RowEdit;
    var expected=wb.DefinedNames.GetDefinedName("Facility").Range.ExistingCells.Where(c=>!c.Value.IsEmpty).Select(c=>c.DisplayText).ToList();
    Check(expected.All(s=>repo.Items.Cast<object>().Any(x=>x.ToString()==s)),"Facility Name choices are sourced from Facility, not Funders");
    grid.FocusedRow=facilityRow;grid.FocusedRecord=1;Call(dit,"SetClipboardTarget",grid);Call(dit,"UpdateFundingScheduleMenu");
    var menu=(ToolStripMenuItem)Field(dit,"FundingScheduleMenuItem");Check(menu.Available&&menu.Text=="Add schedule…","Funding context action is available for a loan cell");
    object[] context={false};var selected=dit.GetType().GetMethod("ContextFundingFacility",F).Invoke(dit,context);
    var cell=(Cell)Call(dit,"InputSourceCell",grid.DataSource,grid.FocusedRecord,(int)Call(dit,"GetVGridColumnIndex",facilityRow,0));
    Check((int)Field(selected,"ColumnIndex")==cell.ColumnIndex,"Context action targets clicked loan's actual worksheet column");
    var decrease=Rows(grid.Rows).OfType<EditorRow>().First(r=>{var t=Call(dit,"GetVGridColumnTag",r,0);return t!=null&&Convert.ToString(Field(t,"BandID"))=="Facility Decreases";});
    grid.FocusedRow=decrease;Check((string)Call(dit,"ContextFundingDateRange")=="IR_Fund_Fac_Dec","Right-click Facility Decreases selects that schedule target");
    var funders=wb.DefinedNames.GetDefinedName("Rep_Fund_03").Range;var facilities=wb.DefinedNames.GetDefinedName("FacilityNames").Range;var loans=wb.DefinedNames.GetDefinedName("LoanDescs").Range;
    // Disposable fixture identities, with deliberately non-adjacent matches.
    funders[0,1].SetValue("Filter fixture A");funders[0,3].SetValue("Filter fixture A");facilities[0,1].SetValue("Filter facility A");facilities[0,3].SetValue("Filter facility B");loans[0,1].SetValue("Loan second");loans[0,3].SetValue("Loan fourth");
    Call(dit,"RefreshData",false);grid.FocusedRow=facilityRow;grid.FocusedRecord=1;Call(dit,"SetContextFundingFilter",true);Pump();
    Console.WriteLine("FILTER count="+grid.RecordCount+" map="+String.Join(",",Enumerable.Range(0,grid.RecordCount).Select(i=>grid.GetDataSourceRecordIndex(i))));
    foreach(DictionaryEntry pair in (IDictionary)Field(dit,"FundingViews")){Console.WriteLine("STATE "+Field(pair.Value,"Funder")+" / "+Field(pair.Value,"Facility"));foreach(DictionaryEntry id in (IDictionary)Field(pair.Value,"Identities"))Console.WriteLine("ID "+id.Key+" col="+Field(id.Value,"ColumnIndex")+" "+Field(id.Value,"Hint"));}
    Check(grid.RecordCount==2&&grid.GetDataSourceRecordIndex(0)==1&&grid.GetDataSourceRecordIndex(1)==3,"Native filter hides other loans while retaining source indexes 1 and 3");
    Check(grid.GetRecordHeaderText(0)=="Loan second"&&grid.GetRecordHeaderText(1)=="Loan fourth","Filtered sticky headers show correct loan names");
    Check((bool)Field(Field(dit,"FundingFilterButton"),"Visible"),"Orange filter indicator is visible");
    grid.FocusedRow=facilityRow;grid.FocusedRecord=1;context[0]=false;selected=dit.GetType().GetMethod("ContextFundingFacility",F).Invoke(dit,context);
    Check((int)Field(selected,"ColumnIndex")==loans.LeftColumnIndex+3,"Filtered context action resolves the fourth physical loan");
    grid.ClearSelection();grid.SelectCell(0,facilityRow,0);grid.SelectCell(1,facilityRow,0);grid.SelectCell(0,decrease,0);grid.SelectCell(1,decrease,0);
    Check(grid.GetSelectedCells().Count()==4,"Funding accepts rectangular multi-cell selection");Call(dit,"CopyFundingWithTitles",grid);var clip=Clipboard.GetText();Console.WriteLine("COPY "+clip.Replace("\r\n"," | "));
    Check(clip.Contains("Loan second")&&clip.Contains("Loan fourth")&&clip.Contains("Facility Decreases")&&clip.Contains("Filter facility B"),"Copy includes selected loan headers, section/date row titles and mapped values");
    var targetsSelected=((IEnumerable)Call(dit,"GetSelectedClipboardDataCells",grid)).Cast<object>().Select(t=>(int)Field(t,"DataRowIndex")).Distinct().OrderBy(i=>i).ToArray();
    Check(targetsSelected.SequenceEqual(new[]{1,3}),"Filtered clipboard edit/clear targets use physical source indexes");
    grid.HideEditor();grid.ClearSelection();grid.TopVisibleRowIndex=0;var funderRow=Rows(grid.Rows).OfType<EditorRow>().First(r=>r.Properties.Caption.Trim()=="Funder");
    var p1=CellPoint(grid,facilityRow,0);var p2=CellPoint(grid,funderRow,1);Drag(grid,p1,p2);
    Console.WriteLine("DRAG selected="+grid.GetSelectedCells().Count());Check(grid.GetSelectedCells().Count()>=4,"Real mouse drag selects a Funding cell rectangle");
    var selectedCount=grid.GetSelectedCells().Count();
    Call(dit,"ClipboardTarget_MouseDown",grid,new MouseEventArgs(MouseButtons.Right,1,p2.X,p2.Y,0));
    context[0]=false;selected=dit.GetType().GetMethod("ContextFundingFacility",F).Invoke(dit,context);
    Check(grid.GetSelectedCells().Count()==selectedCount&&(int)Field(selected,"ColumnIndex")==loans.LeftColumnIndex+3,"Right click retains selection and targets the clicked filtered loan");
    grid.ClearSelection();grid.FocusedRow=facilityRow;grid.FocusedRecord=0;
    var hidden0=facilities[0,0].Value;var hidden2=facilities[0,2].Value;
    var firstFacility=expected[0];var secondFacility=expected.Count>1?expected[1]:expected[0];
    Clipboard.SetText(firstFacility+"\t"+secondFacility);Call(dit,"CustomPasteIntoVGrid",grid);Pump();
    Check(facilities[0,1].DisplayText==firstFacility&&facilities[0,3].DisplayText==secondFacility&&facilities[0,0].Value.Equals(hidden0)&&facilities[0,2].Value.Equals(hidden2),"Filtered two-column paste writes only physical loan columns 2 and 4");
    // Rebuild the grid via the product's refresh action, which reloads preferences.
    Call(dit,"RebuildAllSections");Pump();
    grid=Controls(dit).OfType<VGridControl>().First(g=>g.Visible&&Rows(g.Rows).Any(r=>r.Properties.Caption.Contains("Facility")));
    Console.WriteLine("RESTORED count="+grid.RecordCount+" name="+grid.Name+" indicator="+Field(Field(dit,"FundingFilterButton"),"Visible"));foreach(DictionaryEntry state in (IDictionary)Field(dit,"FundingViews"))Console.WriteLine("RESTORED key="+Field(state.Value,"Key")+" funder="+Field(state.Value,"Funder"));
    Check(grid.RecordCount==2&&grid.GetDataSourceRecordIndex(0)==1&&grid.GetDataSourceRecordIndex(1)==3&&(bool)Field(Field(dit,"FundingFilterButton"),"Visible"),"Rebuilt Funding grid restores saved filter and visible warning");
    presentation.GetMethod("Configure",F).Invoke(null,new object[]{grid,"Regression/"+Guid.NewGuid()});
    facilityRow=Rows(grid.Rows).First(r=>r.Properties.Caption.Contains("Facility")&&r.Properties.Caption.Contains("Name"));repo=(DevExpress.XtraEditors.Repository.RepositoryItemComboBox)facilityRow.Properties.RowEdit;
    decrease=Rows(grid.Rows).OfType<EditorRow>().First(r=>{var t=Call(dit,"GetVGridColumnTag",r,0);return t!=null&&Convert.ToString(Field(t,"BandID"))=="Facility Decreases";});Call(dit,"SetClipboardTarget",grid);
    host.Activate();grid.Focus();
    float baseFont=grid.Appearance.RecordValue.GetFont().SizeInPoints;int baseWidth=grid.RecordWidth;
    presentation.GetMethod("SetZoom",F).Invoke(null,new object[]{grid,130,false});Pump();Console.WriteLine("ZOOM font="+baseFont+"->"+grid.Appearance.RecordValue.GetFont().SizeInPoints+" width="+baseWidth+"->"+grid.RecordWidth);
    Check(grid.Appearance.RecordValue.GetFont().SizeInPoints>baseFont&&grid.RecordWidth>baseWidth,"Grid zoom enlarges body font and loan column geometry");
    Check((int)presentation.GetMethod("ZoomPercent",F).Invoke(null,new object[]{grid})==130,"Grid zoom records its independent scale");
    grid.FocusedRow=facilityRow;grid.FocusedRecord=0;grid.ShowEditor();Pump();Check(grid.ActiveEditor!=null&&grid.ActiveEditor.Properties.Appearance.GetFont().SizeInPoints>baseFont,"Zoom also scales the active Funding editor");
    Check(!Object.ReferenceEquals(grid.ActiveEditor.Properties,repo),"Zoom does not mutate the shared repository editor");grid.HideEditor();
    var htag=Call(dit,"GetVGridColumnTag",decrease,0);var headerDate=Field(htag,"InColumnEditorDate");var editorTag=Field(headerDate,"Tag");var headerHelper=Field(editorTag,"InPlaceVGridRowHelper");Check(headerHelper!=null,"Funding date header helper available for zoom regression");
    Call(headerHelper,"ShowEditorFromKeyboard");Pump();var activeHeader=(BaseEdit)Field(headerHelper,"_ActiveEditor");Check(activeHeader!=null&&activeHeader.Properties.Appearance.GetFont().SizeInPoints>baseFont,"Zoom scales the repeating date editor too");Call(headerHelper,"CommitForSave");
    Picture(dit,Path.Combine(args[1],"funding-filtered-zoom.png"));presentation.GetMethod("SetZoom",F).Invoke(null,new object[]{grid,100,false});
    Check(Math.Abs(grid.Appearance.RecordValue.GetFont().SizeInPoints-baseFont)<1.1,"Reset restores original body font");
    grid.FocusedRow=facilityRow;grid.FocusedRecord=0;grid.ShowEditor();Pump();Check(grid.ActiveEditor!=null&&grid.ActiveEditor.Properties.Appearance.GetFont().SizeInPoints<baseFont+1.1,"Reset restores the active editor font too");grid.HideEditor();
    Console.WriteLine("CLEAR commit="+Call(dit,"CommitEditorsForSave")+" editor="+grid.ActiveEditor);Call(dit,"ClearFundingFilters");Pump();Console.WriteLine("CLEAR records="+grid.RecordCount+" indicator="+Field(Field(dit,"FundingFilterButton"),"Visible"));Check(grid.RecordCount>2&&!(bool)Field(Field(dit,"FundingFilterButton"),"Visible"),"Remove filters restores all loans and clears the orange indicator");
    var tab=(DevExpress.XtraTab.XtraTabControl)Field(dit,"XtraTabControlNewGIT");
    var filter=Activator.CreateInstance(app.GetType("DataInterfaceTemplate+ScrollRediverter"),new object[]{dit,tab});
    var oldCursor=Cursor.Position;
    try {
     host.Activate();grid.Focus();grid.RecordWidth=Math.Max(200,grid.Width/3);grid.LeftVisibleRecord=0;Pump();
     Cursor.Position=grid.PointToScreen(new Point(Math.Min(grid.Width-20,400),Math.Min(grid.Height-20,200)));
     var wheelPosition=(IntPtr)((Cursor.Position.Y<<16)|(Cursor.Position.X&0xffff));
     var routed=(Control)filter.GetType().GetMethod("FindGridScrollOwnerAtScreenPoint",F).Invoke(null,new object[]{tab.SelectedTabPage,Cursor.Position});Console.WriteLine("ROUTED="+routed.Name+" target="+grid.Name+" rows="+Rows(grid.Rows).Count()+" height="+grid.Height+" focus="+grid.ContainsFocus);
     grid.TopVisibleRowIndex=1;Console.WriteLine("TOP SET="+grid.TopVisibleRowIndex);grid.TopVisibleRowIndex=0;
     foreach(var bar in Controls(grid).OfType<DevExpress.XtraEditors.VScrollBar>())Console.WriteLine("SCROLLBAR visible="+bar.Visible+" enabled="+bar.Enabled+" maximum="+bar.Maximum+" large="+bar.LargeChange);
     var wParam=(IntPtr)unchecked((int)((uint)(ushort)(short)-120<<16)|8);
     object[] wheel={Message.Create(grid.Handle,0x20A,wParam,wheelPosition)};
     bool wheelHandled=(bool)Call(filter,"PreFilterMessage",wheel);Console.WriteLine("WHEEL ctrl handled="+wheelHandled+" left="+grid.LeftVisibleRecord+" records="+grid.RecordCount+" width="+grid.Width+" recwidth="+grid.RecordWidth+" point="+Cursor.Position+" rect="+tab.SelectedTabPage.RectangleToScreen(tab.SelectedTabPage.ClientRectangle)+" enabled="+host.Enabled+" active="+(Form.ActiveForm==host));
     Check(wheelHandled&&grid.LeftVisibleRecord>0,"Ctrl+wheel message scrolls loan records horizontally");
     int left=grid.LeftVisibleRecord;int top=grid.TopVisibleRowIndex;var page=tab.SelectedTabPage;int scroll=page.VerticalScroll.Value;
     wheel[0]=Message.Create(grid.Handle,0x20A,(IntPtr)unchecked((int)((uint)(ushort)(short)-120<<16)),wheelPosition);
     Check((bool)Call(filter,"PreFilterMessage",wheel),"Ordinary wheel is routed to the current Funding view");Pump();
     // Native controls may consume the first notch while finalising a restored
     // viewport. Exercise the following notches too, as a physical wheel does.
     for(int notch=0;notch<3&&grid.TopVisibleRowIndex==top&&page.VerticalScroll.Value==scroll;notch++){System.Threading.Thread.Sleep(20);Call(filter,"PreFilterMessage",wheel);Pump();}
     Console.WriteLine("WHEEL top="+top+"->"+grid.TopVisibleRowIndex+" page="+scroll+"->"+page.VerticalScroll.Value+" left="+left+"->"+grid.LeftVisibleRecord);
     Check(grid.LeftVisibleRecord==left&&(grid.TopVisibleRowIndex>top||page.VerticalScroll.Value>scroll),"Ordinary wheel scrolls vertically without moving loan columns");
     wheel[0]=Message.Create(grid.Handle,0x20A,(IntPtr)((120<<16)|4),wheelPosition);
     bool shiftHandled=(bool)Call(filter,"PreFilterMessage",wheel);Pump();
     Check(shiftHandled&&(int)presentation.GetMethod("ZoomPercent",F).Invoke(null,new object[]{grid})==110,"Shift+wheel message zooms the Funding grid");
     presentation.GetMethod("SetZoom",F).Invoke(null,new object[]{grid,100,false});
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
    var zoomKey="Regression/persistent-zoom/"+Guid.NewGuid();
    for(int pass=0;pass<2;pass++)using(var gridHost=new Form())using(var ordinaryGrid=new GridControl())using(var ordinaryView=new GridView(ordinaryGrid)){
     gridHost.Opacity=0;gridHost.ShowInTaskbar=false;gridHost.ClientSize=new Size(600,350);ordinaryGrid.Dock=DockStyle.Fill;gridHost.Controls.Add(ordinaryGrid);ordinaryGrid.MainView=ordinaryView;
     var data=new System.Data.DataTable();data.Columns.Add("First");data.Columns.Add("Second");data.Rows.Add("A","B");ordinaryGrid.DataSource=data;
     ordinaryView.Appearance.Row.Font=new Font("Segoe UI",12);ordinaryView.Appearance.Row.Options.UseFont=true;
     presentation.GetMethod("Configure",F).Invoke(null,new object[]{ordinaryGrid,zoomKey});gridHost.Show();Pump();
     if(pass==0)presentation.GetMethod("SetZoom",F).Invoke(null,new object[]{ordinaryGrid,140,true});
     Check((int)presentation.GetMethod("ZoomPercent",F).Invoke(null,new object[]{ordinaryGrid})==140&&ordinaryView.Appearance.Row.GetFont().SizeInPoints>12,pass==0?"Ordinary GridView zoom scales text":"Recreated GridView restores saved zoom");
     Check(ordinaryGrid.ContextMenuStrip.Items.Find("SummitGridZoom",false).Length==1,"Ordinary grid has native Zoom in/out/reset context actions");
     ordinaryView.OptionsView.ColumnAutoWidth=false;foreach(DevExpress.XtraGrid.Columns.GridColumn col in ordinaryView.Columns)col.Width=500;
     presentation.GetMethod("ModifiedWheel",F).Invoke(null,new object[]{ordinaryGrid,-120,8});Check(ordinaryView.LeftCoord>0,"Ctrl+wheel pans an ordinary GridView horizontally");
     if(pass==1)presentation.GetMethod("SetZoom",F).Invoke(null,new object[]{ordinaryGrid,100,true});gridHost.Close();
    }
    using(var treeHost=new Form())using(var tree=new DevExpress.XtraTreeList.TreeList()){
     treeHost.Opacity=0;treeHost.ShowInTaskbar=false;treeHost.ClientSize=new Size(600,350);tree.Dock=DockStyle.Fill;treeHost.Controls.Add(tree);
     tree.OptionsView.AutoWidth=false;tree.Appearance.Row.Font=new Font("Segoe UI",12);tree.Appearance.Row.Options.UseFont=true;
     tree.Columns.AddVisible("First").Width=500;tree.Columns.AddVisible("Second").Width=500;tree.AppendNode(new object[]{"A","B"},null);
     presentation.GetMethod("Configure",F).Invoke(null,new object[]{tree,"Regression/tree/"+Guid.NewGuid()});treeHost.Show();Pump();
     presentation.GetMethod("SetZoom",F).Invoke(null,new object[]{tree,140,false});
     Check(tree.Appearance.Row.GetFont().SizeInPoints>12&&tree.ContextMenuStrip.Items.Find("SummitGridZoom",false).Length==1,"Tree-style grids share zoom and context menu actions");
     presentation.GetMethod("ModifiedWheel",F).Invoke(null,new object[]{tree,-120,8});Check(tree.LeftCoord>0,"Ctrl+wheel pans a tree-style grid horizontally");treeHost.Close();
    }
    host.Close();
   }
  } finally {files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});}
  Console.WriteLine("PASS "+count+" assertions");return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
