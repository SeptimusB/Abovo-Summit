using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;
using System.Xml.Serialization;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraVerticalGrid;
using DevExpress.XtraVerticalGrid.Rows;
using DevExpress.Spreadsheet;
class FundingScheduleFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static int count;
 static void Check(bool ok,string message){if(!ok)throw new Exception(message);count++;Console.WriteLine("PASS "+message);}
 static object Field(object o,string n){return o.GetType().GetField(n,F).GetValue(o);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static List<DateTime> Generate(Assembly app,DateTime start,int months,int rule,int number,DateTime? end=null,object calendar=null){
  dynamic r=Activator.CreateInstance(app.GetType("Abovo.FundingScheduleRequest"));r.StartDate=start;r.IntervalMonths=months;((object)r).GetType().GetProperty("Rule").SetValue(r,Enum.ToObject(app.GetType("Abovo.FundingDateRule"),rule),null);r.Occurrences=number;r.EndDate=end;r.DelayBankHolidays=calendar!=null;
  return ((IEnumerable)app.GetType("Abovo.FundingScheduleGenerator").GetMethod("Generate").Invoke(null,new object[]{r,calendar})).Cast<object>().Select(x=>(DateTime)x.GetType().GetProperty("DateToAdd").GetValue(x,null)).ToList();
 }
 static IEnumerable<Control> Controls(Control c){foreach(Control d in c.Controls){yield return d;foreach(var n in Controls(d))yield return n;}}
 static void Pump(){Application.DoEvents();}
 [STAThread] static int Main(string[] args) { try {
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  Application.ThreadException+=(s,e)=>{Console.WriteLine("UI EXCEPTION "+e.Exception);Environment.Exit(1);};
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  string mode=args.Length>4?args[4]:"";
  if(mode=="calendar"){
   foreach(var region in new[]{"england-and-wales","scotland","northern-ireland"}){
    dynamic task=app.GetType("Abovo.FundingHolidayCalendar").GetMethod("LoadAsync").Invoke(null,new object[]{region});dynamic calendar=task.GetAwaiter().GetResult();
    Check(calendar.Years.Contains(2026)&&calendar.Dates.Contains(new DateTime(2026,12,25)),"Official calendar loads: "+region);
    Console.WriteLine("CALENDAR "+region+" cached="+calendar.FromCache+" years="+String.Join(", ",((IEnumerable)calendar.Years).Cast<int>().OrderBy(y=>y)));
   }return 0;
  }
  Check(Generate(app,new DateTime(2024,1,31),1,0,3).SequenceEqual(new[]{new DateTime(2024,1,31),new DateTime(2024,2,29),new DateTime(2024,3,31)}),"Month-end clamping does not drift");
  Check(Generate(app,new DateTime(2026,1,31),3,0,3).SequenceEqual(new[]{new DateTime(2026,1,31),new DateTime(2026,4,30),new DateTime(2026,7,31)}),"Quarterly original-day anchor");
  var annual=Generate(app,new DateTime(2024,2,29),12,0,5);Check(annual[1]==new DateTime(2025,2,28)&&annual[4]==new DateTime(2028,2,29),"Annual leap-day anchor");
  Check(Generate(app,new DateTime(2026,1,15),1,1,1)[0]==new DateTime(2026,2,1),"First-day rule respects start boundary");
  Check(Generate(app,new DateTime(2026,8,1),1,3,1)[0]==new DateTime(2026,8,3),"First working day excludes weekends");
  Check(Generate(app,new DateTime(2026,5,1),1,4,1)[0]==new DateTime(2026,5,29),"Last working day stays within month");
  Check(Generate(app,new DateTime(2026,5,1),1,4,12,new DateTime(2026,5,29)).Count==1,"End boundary uses adjusted working date");
  Check(Generate(app,new DateTime(2026,1,1),1,0,12,new DateTime(2026,3,1)).Count==3,"End date inclusive");
  dynamic holidays=Activator.CreateInstance(app.GetType("Abovo.FundingHolidayCalendar"));holidays.Region="England and Wales";holidays.Years.Add(2026);holidays.Dates.Add(new DateTime(2026,8,31));
  Check(Generate(app,new DateTime(2026,8,31),1,0,1,null,holidays)[0]==new DateTime(2026,9,1),"Holiday delayed to next working day");
  Check(Generate(app,new DateTime(2026,8,1),1,4,1,null,holidays)[0]==new DateTime(2026,8,28),"Last working day includes enabled bank holidays");
  bool failed=false;try{Generate(app,new DateTime(2040,1,1),1,0,2,null,holidays);}catch(TargetInvocationException){failed=true;}Check(failed,"Unpublished future holiday coverage rejected");
  failed=false;try{Generate(app,new DateTime(2026,1,1),1,0,1001);}catch(TargetInvocationException){failed=true;}Check(failed,"Oversize schedule rejected");
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  var copy=Path.Combine(args[1],"private-funding.xlsb");File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  if(result.BError)throw new Exception(result.StringReturn);
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)result.IntegerReturn);
  IWorkbook wb=model.WB;
  using(var options=(Form)Activator.CreateInstance(app.GetType("Abovo.ApplicationOptionsForm"))){
   Check(!Controls(options).OfType<SimpleButton>().Any(b=>b.Text=="Apply"),"Settings has no redundant Apply button");
   Check(options.AcceptButton!=null&&((SimpleButton)options.AcceptButton).Text=="OK","Settings OK remains the default confirmation");
  }
  if(args[3]=="True")foreach(var n in wb.DefinedNames.Where(n=>n.Name.StartsWith("Rep_Fund",StringComparison.OrdinalIgnoreCase)))Console.WriteLine(n.Name+" "+n.Range.GetReferenceA1());
  if(args[3]=="True")foreach(var name in wb.DefinedNames.Where(n=>n.Name.StartsWith("IR_Fund_",StringComparison.OrdinalIgnoreCase))) {
   var r=name.Range;Console.WriteLine(name.Name+" "+r.GetReferenceA1()+" "+r.RowCount+"x"+r.ColumnCount);
   for(int i=0;i<Math.Min(2,r.RowCount);i++){var c=r[i,0];Console.WriteLine("  "+c.GetReferenceA1()+" value="+c.Value+" formula="+c.HasFormula+" locked="+c.Protection.Locked+" pattern="+c.Fill.PatternType);}
  }
  var writer=app.GetType("Abovo.FundingScheduleWriter");var targets=((IEnumerable)writer.GetField("Targets").GetValue(null)).Cast<object>().ToArray();
  if(args[3]=="True"){foreach(dynamic t in targets){var r=wb.DefinedNames.GetDefinedName((string)t.ValueRange).Range;Console.WriteLine(t.Caption+" values="+r.GetReferenceA1());}files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});return 0;}
  var blank=writer.GetMethod("BlankRows");var apply=writer.GetMethod("Apply");
  var xml=new XmlDocument();xml.Load(Path.Combine(args[1],"Structure.xml"));
  var mappings=new List<object>();var mapper=app.GetType("Abovo.DataManager").GetMethod("ResolveAnchoredInputRange");
  foreach(XmlNode n in xml.SelectNodes("//ChildStructure[CSID='33']//CellRangeDataSource[DataRangeAnchorNR]")){
   var serializer=new XmlSerializer(app.GetType("Abovo.CellRangeDataSource"));var definition=serializer.Deserialize(new StringReader(n.OuterXml));mappings.Add(definition);
   var resolved=(CellRange)mapper.Invoke(null,new object[]{wb.Worksheets["Funding Assumptions"],definition});
   Check(resolved.GetReferenceA1()==wb.Worksheets["Funding Assumptions"].Range[n["DataRange"].InnerText].GetReferenceA1(),"Named anchor matches original field: "+n["DataRange"].InnerText);
  }
  Func<object,List<int>> blanks=t=>(List<int>)blank.Invoke(null,new object[]{wb,t});
  var sample=new DateTime(2028,4,28);
  if(mode!="ui")foreach(dynamic t in targets){
   List<int> positions=blanks((object)t);Check(positions.Count>=2,"Safe date capacity: "+t.Caption);
   CellRange range=wb.DefinedNames.GetDefinedName((string)t.DateRange).Range;
   var before=range.ExistingCells.ToDictionary(c=>c.GetReferenceA1(),c=>c.Value);
   bool protection=range.Worksheet.IsProtected;var firstFormula=range[0,0].FormulaInvariant;
   int added=(int)apply.Invoke(null,new object[]{(int)model.ModelID,t,new List<DateTime>{sample,sample}});
   Check(added==0&&range[positions[0],0].Value.DateTimeValue==sample&&range[positions[1],0].Value.DateTimeValue==sample,"Duplicate dates are separate rows: "+t.Caption);
   Check(range.Worksheet.IsProtected==protection&&range[0,0].FormulaInvariant==firstFormula,"Protection/formula-owned first date preserved: "+t.Caption);
   Check(model.ChangeManager.Undo().BSuccess&&range.ExistingCells.All(c=>before.ContainsKey(c.GetReferenceA1())?before[c.GetReferenceA1()].Equals(c.Value):c.Value.IsEmpty),"Grouped date undo: "+t.Caption);
   Check(model.ChangeManager.Redo().BSuccess&&range[positions[1],0].Value.DateTimeValue==sample,"Grouped date redo: "+t.Caption);
   Check(model.ChangeManager.Undo().BSuccess,"Reset private schedule: "+t.Caption);
  }
  if(mode!="ui")foreach(string dateName in new[]{"IR_Fund_Repay","IR_Fund_Int_Rec","IR_Fund_Inv_Inc"}){
   dynamic t=targets.First(x=>(string)x.GetType().GetProperty("DateRange").GetValue(x,null)==dateName);
   var range=wb.DefinedNames.GetDefinedName(dateName).Range;int top=range.TopRowIndex,bottom=range.BottomRowIndex,oldRows=range.RowCount;
   var ws=range.Worksheet;bool protectedBefore=ws.IsProtected;
   var constants=ws.GetUsedRange().ExistingCells.Where(c=>!c.HasFormula&&!c.Value.IsEmpty).Select(c=>Tuple.Create(c.RowIndex,c.ColumnIndex,c.Value)).ToArray();
   var mapPositions=mappings.Select(x=>((CellRange)mapper.Invoke(null,new object[]{ws,x})).TopRowIndex).ToArray();
   int requested=blanks((object)t).Count+3;var dates=Enumerable.Range(0,requested).Select(i=>sample.AddMonths(i)).ToList();
   Check((int)apply.Invoke(null,new object[]{(int)model.ModelID,t,dates})==8,"Expansion adds shortage plus five: "+t.Caption);
   range=wb.DefinedNames.GetDefinedName(dateName).Range;
   Check(range.RowCount==oldRows+8&&blanks(t).Count==5,"Five spare safe date rows remain: "+t.Caption);
   Check(constants.All(c=>ws.Cells[c.Item1>bottom?c.Item1+8:c.Item1,c.Item2].Value.Equals(c.Item3)),"Existing Funding constants preserved: "+t.Caption);
   Check(ws.IsProtected==protectedBefore,"Expansion restores protection: "+t.Caption);
   Check(mappings.Select((x,i)=>((CellRange)mapper.Invoke(null,new object[]{ws,x})).TopRowIndex==(mapPositions[i]>bottom?mapPositions[i]+8:mapPositions[i])).All(x=>x),"All literal-input anchors track inserted rows: "+t.Caption);
   Check(model.ChangeManager.Undo().BSuccess&&range.RowCount==oldRows+8,"Undo dates retains added capacity: "+t.Caption);
   Check(model.ChangeManager.Redo().BSuccess,"Expansion date redo: "+t.Caption);
  }
  using(var host=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})){
   host.Opacity=0;host.ShowInTaskbar=false;host.ClientSize=new Size(1700,1050);host.Show();Pump();
   Call(host,"ShowInterface",(int)model.ModelID,33,false,"None",null,1);Pump();var dit=(Control)Field(host,"ActiveInterface");
   Console.WriteLine("DIT GSID="+Field(dit,"GSID")+" CSID="+Field(dit,"CSID"));
   foreach(var tab in Controls(dit).OfType<DevExpress.XtraTab.XtraTabControl>()) {Console.WriteLine("TABS "+String.Join(", ",tab.TabPages.Cast<DevExpress.XtraTab.XtraTabPage>().Select(p=>p.Text))+" selected="+tab.SelectedTabPageIndex);tab.SelectedTabPageIndex=1;}
   Pump();
   foreach(var e in ((IEnumerable)Field(dit,"VGridCategoryExtenders")).Cast<object>()){var c=(CategoryRow)Field(e,"ActiveCategoryRow");Console.WriteLine("CATEGORY "+c.Properties.Caption+" range="+(c.Tag==null?"null":Field(c.Tag,"ActionNR")));}
   var extenders=((IEnumerable)Field(dit,"VGridCategoryExtenders")).Cast<object>().Where(e=>Field(e,"ScheduleAction")!=null).ToArray();
   Console.WriteLine("HEADER ACTIONS "+extenders.Length+": "+String.Join(", ",extenders.Select(e=>((CategoryRow)Field(e,"ActiveCategoryRow")).Properties.Caption)));
   Check(extenders.Length==7,"All seven Facilities and Loans date categories have header actions");
   var extender=extenders.First(e=>((CategoryRow)Field(e,"ActiveCategoryRow")).Properties.Caption.Trim()=="Repayments");
   var grid=(VGridControl)Field(extender,"view");var row=(CategoryRow)Field(extender,"ActiveCategoryRow");
   foreach(var category in grid.Rows.OfType<CategoryRow>())category.Expanded=false;
   grid.MakeRowVisible(row);grid.FocusedRow=row;grid.Refresh();Pump();
   using(var bitmap=new Bitmap(grid.Width,Math.Min(950,grid.Height))){grid.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(args[1],"funding-header.png"));}
   var rect=(Rectangle)Field(extender,"ScheduleButtonRect");Console.WriteLine("HEADER RECT "+rect);Check(rect.Width>=12&&rect.Height>=12,"Header plus has usable bounds");
   var beforeCancel=wb.Worksheets["Funding Assumptions"].GetUsedRange().ExistingCells.ToDictionary(c=>c.GetReferenceA1(),c=>c.Value);
   bool shown=false,previewed=false;
   using(var timer=new Timer()){timer.Interval=150;timer.Tick+=(s,e)=>{foreach(Form f in Application.OpenForms.Cast<Form>().ToArray())if(f.GetType().Name=="FundingScheduleForm"){
     shown=true;var target=f.GetType().GetProperty("SelectedTarget").GetValue(f,null);Check((string)target.GetType().GetProperty("DateRange").GetValue(target,null)=="IR_Fund_Repay","Header click selects associated section");
     previewed=((SimpleButton)Field(f,"ApplyButton")).Enabled;
     using(var bitmap=new Bitmap(f.Width,f.Height)){f.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(args[1],"funding-preview.png"));}
     timer.Stop();f.DialogResult=DialogResult.Cancel;f.Close();break;
   }};timer.Start();
   var point=new Point(rect.Left+rect.Width/2,rect.Top+rect.Height/2);var mouse=new MouseEventArgs(MouseButtons.Left,1,point.X,point.Y,0);
   Call(extender,"OnMouseDown",grid,mouse);Call(extender,"OnMouseUp",grid,mouse);Pump();timer.Stop();}
   Check(shown&&previewed,"Native header plus opens a populated preview; cancel is read-only");
   Check(wb.Worksheets["Funding Assumptions"].GetUsedRange().ExistingCells.All(c=>beforeCancel.ContainsKey(c.GetReferenceA1())?beforeCancel[c.GetReferenceA1()].Equals(c.Value):c.Value.IsEmpty),"Preview and Cancel leave Funding cell values unchanged");
   dynamic repayment=targets.First(x=>(string)x.GetType().GetProperty("DateRange").GetValue(x,null)=="IR_Fund_Repay");
   var repaymentRange=wb.DefinedNames.GetDefinedName("IR_Fund_Repay").Range;
   List<int> freeRows=blanks((object)repayment);int firstScheduleRow=repaymentRange.TopRowIndex+freeRows[0];
   bool applied=false;
   using(var timer=new Timer()){timer.Interval=150;timer.Tick+=(s,e)=>{foreach(Form f in Application.OpenForms.Cast<Form>().ToArray())if(f.GetType().Name=="FundingScheduleForm"){
     timer.Stop();((DateEdit)Field(f,"EndDateChoice")).DateTime=((DateEdit)Field(f,"StartChoice")).DateTime.AddMonths(1);
     ((CheckEdit)Field(f,"PopulateChoice")).Checked=true;((SpinEdit)Field(f,"FigureChoice")).Value=123.456m;
     Check(!((SimpleButton)Field(f,"ApplyButton")).Enabled,"Changing options invalidates the previous preview");
     ((System.Threading.Tasks.Task)Call(f,"RefreshPreviewAsync")).GetAwaiter().GetResult();
     Check(((SimpleButton)Field(f,"ApplyButton")).Enabled,"Changed preview enables Apply");
     Check((decimal)f.GetType().GetProperty("FixedFigure").GetValue(f,null)==123.456m,"Fixed-figure editor is included in the current preview");
     applied=true;((SimpleButton)Field(f,"ApplyButton")).PerformClick();break;
   }};timer.Start();((Action)Field(extender,"ScheduleAction"))();Pump();Pump();timer.Stop();}
   Check(applied&&!wb.Worksheets["Funding Assumptions"].Cells[firstScheduleRow,repaymentRange.LeftColumnIndex].Value.IsEmpty,"Native Apply posts the preview dates");
   dynamic defaultFacility=((IEnumerable)app.GetType("Abovo.FundingScheduleGroups").GetMethod("Facilities").Invoke(null,new object[]{wb,false})).Cast<object>().First();
   Check(wb.Worksheets["Funding Assumptions"].Cells[firstScheduleRow,(int)defaultFacility.ColumnIndex].Value.NumericValue==123.456,"Native dialog Apply posts the optional fixed figure");
   var focused=((IEnumerable)Field(dit,"VertGridControls")).Cast<VGridControl>().FirstOrDefault(g=>g.ContainsFocus);
   var destinations=((IEnumerable)Call(dit,"NavigationPositions")).Cast<object>().Where(p=>Field(p,"VRow")!=null).Where(p=>{
    dynamic tag=Call(dit,"GetVGridColumnTag",Field(p,"VRow"),0);if(tag==null||tag.BandID!="Repayments")return false;
    var hostGrid=(VGridControl)Field(p,"Host");int col=(int)Call(dit,"GetVGridColumnIndex",Field(p,"VRow"),0);
    var c=(Cell)Call(dit,"InputSourceCell",hostGrid.DataSource,Math.Max(0,(int)Field(p,"Record")),col);return c!=null&&c.RowIndex==firstScheduleRow;
   }).ToList();
   if(destinations.Any(p=>(int)Field(p,"Record")>=0)){
    Check(focused!=null&&focused.FocusedRow!=null&&focused.FocusedRecord>=0,"Apply restores focus to a Funding amount editor");
    int column=(int)Call(dit,"GetVGridColumnIndex",focused.FocusedRow,0);
    var focusedCell=(Cell)Call(dit,"InputSourceCell",focused.DataSource,focused.FocusedRecord,column);
    Check(focusedCell!=null&&focusedCell.RowIndex==firstScheduleRow&&focusedCell.ColumnIndex>repaymentRange.LeftColumnIndex&&focused.FocusedRecord==destinations.Where(p=>(int)Field(p,"Record")>=0).Min(p=>(int)Field(p,"Record")),"Focused amount is first editable cell beside the first new date");
   }else{
    Check(destinations.Any(p=>Field(p,"VHeader")!=null),"Blank template retains a navigable first date without unlocking loans");
    Check(destinations.Where(p=>Field(p,"VHeader")!=null).Any(p=>{var editor=Field(Field(p,"VHeader"),"_ActiveEditor") as Control;return editor!=null&&editor.ContainsFocus;}),"With no editable loan, focus returns to the first new date header");
   }
   // Exercise the facility-targeted context menu, automatic preview and grouped application.
   var currentGrid=Controls(dit).OfType<VGridControl>().First(g=>!g.IsDisposed&&g.Visible);
   var details=currentGrid.Rows.OfType<CategoryRow>().First(r=>r.Properties.Caption=="Funding Details");details.Expanded=true;
   currentGrid.FocusedRow=details.ChildRows[0];currentGrid.FocusedRecord=1;
   Call(dit,"SetClipboardTarget",currentGrid);Call(dit,"UpdateFundingScheduleMenu");
   var scheduleMenu=(ToolStripMenuItem)Field(dit,"FundingScheduleMenuItem");
   Check(scheduleMenu.Available&&currentGrid.OptionsView.ShowRecordHeaders,"Facility context action and sticky loan header strip are available");
   bool multiApplied=false;int phase=0;DateTime previewDue=DateTime.MinValue;
   using(var timer=new Timer()){timer.Interval=100;timer.Tick+=(s,e)=>{foreach(Form f in Application.OpenForms.Cast<Form>().ToArray())if(f.GetType().Name=="FundingScheduleForm"){
     if(phase==0){
      Check(((ComboBoxEdit)Field(f,"FacilityChoice")).SelectedIndex==1,"Facility context menu preselects the clicked loan");
      var sections=(CheckedListBoxControl)Field(f,"SectionChoice");Check(sections.Items.Count==11,"Facility action offers all eleven compatible sections across tabs");
      sections.SetItemChecked(1,true);((DateEdit)Field(f,"EndDateChoice")).DateTime=((DateEdit)Field(f,"StartChoice")).DateTime.AddMonths(1);
      phase=1;previewDue=DateTime.UtcNow.AddSeconds(3);return;
     }
     if(((SimpleButton)Field(f,"ApplyButton")).Enabled){
      Check(((IEnumerable)f.GetType().GetProperty("SelectedTargets").GetValue(f,null)).Cast<object>().Count()==2,"Automatic preview includes both selected sections");
     Check(((ICollection)f.GetType().GetProperty("DatesToApply").GetValue(f,null)).Count==2,"Automatic preview uses the latest inclusive end date");
      Check(f.ClientSize.Height>=800,"Schedule dialog is taller");
      var filter=(IMessageFilter)Field(dit,"InterfaceScrollRediverter");
      var wheel=Message.Create(f.Handle,0x20A,new IntPtr(120<<16),IntPtr.Zero);
      Check(!filter.PreFilterMessage(ref wheel),"DIT does not capture the schedule dialog's wheel input");
      using(var options=(Form)Activator.CreateInstance(app.GetType("Abovo.ApplicationOptionsForm"))){var optionsWheel=Message.Create(options.Handle,0x20A,new IntPtr(120<<16),IntPtr.Zero);Check(!filter.PreFilterMessage(ref optionsWheel),"DIT does not capture Options wheel input");}
      Check(f.GetType().GetProperty("FixedFigure").GetValue(f,null)==null&&!((SpinEdit)Field(f,"FigureChoice")).Enabled,"Fixed figures are opt-in");
      Check(!Controls(f).OfType<SimpleButton>().Any(b=>b.Text.Contains("preview")),"No mandatory Preview button");
      using(var bitmap=new Bitmap(f.Width,f.Height)){f.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(args[1],"multi-preview.png"));}
      timer.Stop();multiApplied=true;((SimpleButton)Field(f,"ApplyButton")).PerformClick();return;
     }
     if(DateTime.UtcNow>previewDue)throw new Exception("Automatic preview did not become valid: "+((LabelControl)Field(f,"Status")).Text);
   }};timer.Start();scheduleMenu.PerformClick();Pump();Pump();timer.Stop();}
   Check(multiApplied,"Facility context action applies one schedule to multiple sections");
   var selectedGrid=Controls(dit).OfType<VGridControl>().FirstOrDefault(g=>!g.IsDisposed&&g.ContainsFocus);
   var activeHeaders=((IDictionary)Field(dit,"VGridInplaceEditorHelpersBySection")).Values.Cast<IEnumerable>().SelectMany(x=>x.Cast<object>()).Where(h=>{var edit=Field(h,"_ActiveEditor") as Control;return edit!=null&&edit.ContainsFocus;}).ToList();
   Console.WriteLine("FOCUS record="+(selectedGrid==null?-99:selectedGrid.FocusedRecord)+" activeDates="+activeHeaders.Count);
   Check((selectedGrid!=null&&selectedGrid.FocusedRecord==1)||activeHeaders.Count>0,"Focus returns to selected facility or defining date when that loan amount is locked");
   var tintMap=(IDictionary)Field(dit,"FundingTints");Check(tintMap.Count>=12,"Date and amount tints are populated after native Apply");
   int renderedTargets=0,editableTints=0;
   foreach(var vg in Controls(dit).OfType<VGridControl>().Where(g=>g.Visible&&!g.IsDisposed)){
    vg.CustomDrawRowValueCell+=(s,e)=>{
     var c=(Cell)Call(dit,"InputSourceCell",vg.DataSource,vg.GetDataSourceRecordIndex(e.RecordIndex),(int)Call(dit,"GetVGridColumnIndex",e.Row,e.CellIndex));
     if(c==null)return;string key=c.RowIndex+":"+c.ColumnIndex;if(!tintMap.Contains(key))return;
     if((bool)tintMap[key].GetType().GetProperty("Active").GetValue(tintMap[key],null)){Check(e.Handled,"Scheduled amount receives the actual native colour-marker painter");renderedTargets++;}
     if(!c.Protection.Locked&&c.Fill.PatternType==PatternType.Solid){var expected=(Color)tintMap[key].GetType().GetProperty("Colour").GetValue(tintMap[key],null);Check(e.Appearance.BackColor==expected||e.Appearance.BackColor==SystemColors.Highlight,"Editable scheduled amount has its tint or native selection highlight");editableTints++;}
    };
    // Funding now has an internal viewport and sticky loan header. Bring an
    // eligible tinted cell into that viewport before checking its rendering.
    bool foundTint=false;
    foreach(BaseRow category in vg.Rows){foreach(BaseRow tintRow in category.ChildRows){if(!(tintRow is EditorRow))continue;for(int record=0;record<vg.RecordCount;record++){
     var c=(Cell)Call(dit,"InputSourceCell",vg.DataSource,vg.GetDataSourceRecordIndex(record),(int)Call(dit,"GetVGridColumnIndex",tintRow,0));
     if(c!=null&&tintMap.Contains(c.RowIndex+":"+c.ColumnIndex)&&!c.Protection.Locked&&c.Fill.PatternType==PatternType.Solid){vg.MakeRowVisible(tintRow);vg.LeftVisibleRecord=record;foundTint=true;break;}
    }if(foundTint)break;}if(foundTint)break;}
    Pump();
    using(var bitmap=new Bitmap(vg.Width,Math.Min(1500,vg.Height))){vg.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(args[1],"schedule-targets.png"));}
   }
   Check(renderedTargets>0,"Scheduled target cells were rendered, not just stored in metadata");
   Check(editableTints>0,"Editable scheduled target background verified");
   foreach(var list in ((IDictionary)Field(dit,"VGridInplaceEditorHelpersBySection")).Values.Cast<IEnumerable>())foreach(var helper in list.Cast<object>()){
    var tint=helper.GetType().GetProperty("PresentationTint").GetValue(helper,null) as Delegate;if(tint!=null&&tint.DynamicInvoke()!=null){Check(true,"Defining date header receives its group tint");goto DateTintChecked;}
   }throw new Exception("No defining date tint found");DateTintChecked:;
   foreach(int tabIndex in new[]{2,4}){
    foreach(var tab in Controls(dit).OfType<DevExpress.XtraTab.XtraTabControl>())tab.SelectedTabPageIndex=tabIndex;
    Pump();
    var visibleActions=((IEnumerable)Field(dit,"VGridCategoryExtenders")).Cast<object>().Where(e=>Field(e,"ScheduleAction")!=null&&((VGridControl)Field(e,"view")).Visible).ToArray();
    Check(visibleActions.Length==(tabIndex==2?4:3),"All date sections offer schedules on tab "+tabIndex);
   }
   host.Close();
  }
  string saved=Path.Combine(args[1],"funding-schedule-saved.xlsb");Check((bool)model.SaveFileAsTo(saved,true),"Save scheduled workbook copy");
  using(var reopened=new Workbook()){reopened.Options.CalculationMode=WorkbookCalculationMode.Manual;Check(reopened.LoadDocument(saved),"Scheduled XLSB reopens");foreach(dynamic t in targets){var r=wb.DefinedNames.GetDefinedName((string)t.DateRange).Range;var n=reopened.DefinedNames.GetDefinedName((string)t.DateRange).Range;Check(r.GetReferenceA1()==n.GetReferenceA1()&&r.ExistingCells.All(c=>c.Value.Equals(n.Worksheet.Cells[c.GetReferenceA1()].Value)),"Saved date range preserved: "+t.Caption);}}
  files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});Console.WriteLine("PASS "+count+" assertions");return 0;
 }catch(Exception e){Console.WriteLine(e);return 1;}}
}
