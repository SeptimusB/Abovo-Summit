// Diagnostic only: never saves a source workbook. UI edits use the production
// editor/ChangeManager path on a private copy and are undone after each probe.
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraVerticalGrid;
using DevExpress.XtraVerticalGrid.Rows;

public static class ClientInputReviewFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static object Field(object o,string n){if(o==null)return null;var f=o.GetType().GetField(n,F);if(f!=null)return f.GetValue(o);var p=o.GetType().GetProperty(n,F);return p==null?null:p.GetValue(o,null);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static IEnumerable<object> Items(object o){return o==null?Enumerable.Empty<object>():((IEnumerable)o).Cast<object>();}
 static string S(object o){return Convert.ToString(o,CultureInfo.InvariantCulture);}
 static string Flat(object o){return System.Text.RegularExpressions.Regex.Replace(S(o),@"\s+"," ").Trim();}
 static string Q(object o){return "\""+S(o).Replace("\"","\"\"")+"\"";}
 static void Row(StreamWriter w,params object[] values){w.WriteLine(String.Join(",",values.Select(Q)));}
 static void Pump(){Application.DoEvents();Application.RaiseIdle(EventArgs.Empty);Application.DoEvents();}
 static IWorkbook workbook;
 static HashSet<object> dumped=new HashSet<object>();
 static HashSet<string> tested=new HashSet<string>();
 [STAThread] public static int Main(string[] args){try {
  System.Threading.Thread.CurrentThread.CurrentCulture=CultureInfo.GetCultureInfo("en-GB");
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  var copy=Path.Combine(args[1],"private-input-review.xlsb");File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  if(result.BError)throw new Exception("Private model did not open");
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)result.IntegerReturn);workbook=(IWorkbook)model.WB;
  try {
   using(var output=new StreamWriter(Path.Combine(args[1],"input-inventory.csv")))
   using(var tests=new StreamWriter(Path.Combine(args[1],"editor-tests.csv")))
   using(var form=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})){
    output.AutoFlush=true;tests.AutoFlush=true;
    Row(output,"CSID","LoadedWhileTab","DataSet","Heading","Type","RO","Calculated","Rules","Min","Max","Repository","Sheet","Cell","Value","Display","Format","Formula","WorkbookLocked","DataLocked","PasteAllowed","DecimalParsed","ParsedValue","FillPattern","CFCount");
    Row(tests,"CSID","Tab","Heading","Sheet","Cell","Type","Editor","Min","Max","Requested","Stored","Precise","Undo","PasteParsed","PasteValue","GridDisplay");
    form.Opacity=0;form.ShowInTaskbar=false;form.ClientSize=new Size(1750,1000);form.Show();Pump();
    var cases=args.Length>4&&!String.IsNullOrEmpty(args[4])?args[4].Split(',').Select(Int32.Parse).ToArray():new[]{3,14,15,16,26,30,32,33,35,39,41,42};
    foreach(int csid in cases){
     Console.WriteLine("REVIEW CSID="+csid);Console.Out.Flush();
     Call(form,"ShowInterface",0,csid,false,"None",null,-1);Pump();
     var dit=(Control)Field(form,"ActiveInterface");dynamic tabs=Field(dit,"XtraTabControlNewGIT");
     for(int ti=0;ti<tabs.TabPages.Count;ti++){
      string tab=Flat(tabs.TabPages[ti].Text);
      if(tab.StartsWith("XtraTabPage",StringComparison.Ordinal))continue;
      // All the reported tables, avoiding unrelated import operations.
      Call(dit,"BuildSection",ti,false,false);tabs.SelectedTabPageIndex=ti;Pump();
      Dump(dit,csid,tab,output);
      if(args[3]!="True"){ProbeEditors(dit,model,csid,tab,tests);
       if(csid==14&&tab=="Management Costs Assumptions")ProbePaste(dit,model);
       if(csid==15&&tab=="Stock Survey Allocation")foreach(var control in Descendants(dit).Where(c=>c.GetType().Name=="AbovoDEDateEdit"))Console.WriteLine("DATE_EDITOR target="+Field(control,"TargetCell")+" value="+Field(control,"EditValue")+" text="+control.Text+" enabled="+control.Enabled);
      }
     }
    }
    form.Close();
   }
  } finally {files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});}
  Console.WriteLine("REVIEW FINISHED; original not saved; no app source changes");return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
 static void Dump(object dit,int csid,string tab,StreamWriter output){
  int dsIndex=-1;
  foreach(var ds in Items(Field(Field(dit,"DataPres"),"DataSets"))){
   dsIndex++;if(ds==null||!dumped.Add(ds))continue;
   var columns=Items(Field(ds,"DataColumns")).ToArray();var rows=Items(Field(ds,"DataRows")).ToArray();
   for(int c=0;c<columns.Length;c++){
    var tag=Field(columns[c],"ColumnTag");if(tag==null)continue;
    for(int r=0;r<Math.Min(3,rows.Length);r++){
     var cells=Items(Field(rows[r],"DataCells")).ToArray();if(c>=cells.Length)continue;var dp=cells[c];if(dp==null)continue;
     string sn=S(Field(dp,"SourceSheet")),addr=S(Field(dp,"SourceAddress"));if(sn==""||addr=="")continue;
     var cell=workbook.Worksheets[sn].Cells[addr];object[] parse={S(Field(tag,"DataType"))=="P"?"-2.5%":"1234.56",tag,null};
     bool parsed=(bool)Call(dit,"ConvertPastedTextValue",parse);
     Row(output,csid,tab,dsIndex,Flat(Field(tag,"ColumnHeading")),Field(tag,"DataType"),Field(tag,"IsReadOnly"),Field(tag,"IsCalculated"),Field(tag,"HasRules"),Field(tag,"MinVal"),Field(tag,"MaxVal"),Field(tag,"RepositaryID"),sn,addr,cell.Value,cell.DisplayText,cell.NumberFormat,cell.HasFormula,cell.Protection.Locked,Field(dp,"IsLocked"),Call(dit,"CanPasteToDataPoint",ds,r,c),parsed,parse[2],cell.Fill.PatternType,cell.Worksheet.ConditionalFormattings.Count);
    }
   }
  }
 }
 static string Caption(object p){var vr=(EditorRow)Field(p,"VRow");return Flat(vr!=null?vr.Properties.Caption:Field(Field(p,"Column"),"Caption"));}
 static void ProbeEditors(object dit,dynamic model,int csid,string tab,StreamWriter tests){
  var positions=Items(Call(dit,"NavigationPositions")).Where(p=>Field(p,"Header")==null&&Field(p,"VHeader")==null).ToList();
  foreach(var point in positions){
   string caption=Caption(point);var view=Field(point,"View") as GridView;var vertical=Field(point,"Host") as VGridControl;
   var column=(GridColumn)Field(point,"Column");var vr=(EditorRow)Field(point,"VRow");
   var tag=view!=null?Field(column,"Tag"):Call(dit,"GetVGridColumnTag",vr,0);string type=S(Field(tag,"DataType"));
   bool select= csid==3&&caption.Contains("Avg") || csid==16&&type=="I" || csid==26&&tab=="Development Revenue"&&(caption.Contains("Rent")||caption.Contains("Charge")) ||
     csid==30&&type=="I" || csid==32&&type=="P" || csid==39&&(type=="I"||type=="M"||type=="R"||type=="SM") || (csid==41||csid==42)&&caption.Contains("Amount");
   string key=csid+"|"+tab+"|"+caption;if(!select||!tested.Add(key))continue;
   // One representative per CPI/real-rent table, not every forecast column.
   if(csid==32&&tested.Any(k=>k.StartsWith(csid+"|"+tab+"|")&&k!=key))continue;
   if(csid==16&&tested.Any(k=>k.StartsWith("16|")&&k!=key))continue;
   dynamic source=view!=null?view.GridControl.DataSource:vertical.DataSource;
   int index=(int)source.UBSTag.DSIndex;
   int ci=(int)(view!=null?Call(dit,"GetGridColumnIndex",column):Call(dit,"GetVGridColumnIndex",vr,0));
   int ri=view!=null?view.GetDataSourceRowIndex((int)Field(point,"RowHandle")):(int)Field(point,"Record");
   dynamic pres=Field(dit,"DataPres");dynamic dp=pres.DataSets[index].DataRows[ri].DataCells[ci];
   var sheet=workbook.Worksheets[(string)dp.SourceSheet];var cell=sheet.Cells[(string)dp.SourceAddress];var old=cell.Value;
   bool originalProtection=sheet.IsProtected;string format=cell.NumberFormat;double requested=type=="P"?-0.025:1234.56;
   Call(dit,"ActivateEditorPosition",point);Pump();var editor=view!=null?view.ActiveEditor:vertical.ActiveEditor;
   if(editor==null){Row(tests,csid,tab,caption,sheet.Name,cell.GetReferenceA1(),type,"NOT OPENED");continue;}
   string min=S(Field(editor.Properties,"MinValue")),max=S(Field(editor.Properties,"MaxValue"));
   editor.EditValue=requested;
   bool posted=view!=null?view.PostEditor()&&view.UpdateCurrentRow():vertical.PostEditor();
   if(view!=null)view.CloseEditor();else vertical.CloseEditor();Pump();
   var stored=cell.Value;string display=view!=null?view.GetRowCellDisplayText((int)Field(point,"RowHandle"),column):vertical.GetCellDisplayText(vr,(int)Field(point,"Record"));bool changed=!stored.Equals(old),undoOK=!changed;
   if(changed){dynamic undo=model.ChangeManager.Undo();undoOK=!undo.BError&&cell.Value.Equals(old);Pump();}
   object[] parse={type=="P"?"-2.5%":"1234.56",tag,null};bool parsed=(bool)Call(dit,"ConvertPastedTextValue",parse);
   bool precise=posted&&stored.IsNumeric&&Math.Abs(stored.NumericValue-requested)<1e-9;
   Row(tests,csid,tab,caption,sheet.Name,cell.GetReferenceA1(),type,editor.GetType().Name,min,max,requested,stored,precise,undoOK,parsed,parse[2],display);
   Console.WriteLine("INPUT CS="+csid+" "+caption+" "+sheet.Name+"!"+cell.GetReferenceA1()+" type="+type+" requested="+requested+" stored="+stored+" precise="+precise+" undo="+undoOK);
   if(!undoOK||originalProtection!=sheet.IsProtected||format!=cell.NumberFormat)throw new Exception("Diagnostic edit did not restore original state");
  }
 }
 static IEnumerable<Control> Descendants(Control c){foreach(Control child in c.Controls){yield return child;foreach(var nested in Descendants(child))yield return nested;}}
 static void ProbePaste(object dit,dynamic model){
  var points=Items(Call(dit,"NavigationPositions")).Where(p=>Field(p,"View")!=null&&Field(p,"Header")==null&&Caption(p)=="Summary Cost Category").Take(2).ToList();
  if(points.Count<2)throw new Exception("Expected two Management Cost dropdown destinations");
  var view=(GridView)Field(points[0],"View");var column=(GridColumn)Field(points[0],"Column");dynamic source=view.GridControl.DataSource;
  int index=(int)source.UBSTag.DSIndex;dynamic pres=Field(dit,"DataPres");object ds=pres.DataSets[index];
  int ci=(int)Call(dit,"GetGridColumnIndex",column);var rows=points.Select(p=>view.GetDataSourceRowIndex((int)Field(p,"RowHandle"))).ToArray();
  var cells=rows.Select(r=>{dynamic dp=((dynamic)ds).DataRows[r].DataCells[ci];return workbook.Worksheets[(string)dp.SourceSheet].Cells[(string)dp.SourceAddress];}).ToArray();
  var old=cells.Select(c=>c.Value).ToArray();
  Call(dit,"ActivateEditorPosition",points[0]);Pump();var combo=view.ActiveEditor as ComboBoxEdit;
  if(combo==null)throw new Exception("Expected native Management category dropdown");
  string value=combo.Properties.Items.Cast<object>().Select(S).First(v=>!String.IsNullOrWhiteSpace(v)&&v!="<Blank>"&&cells.All(c=>c.Value.TextValue.Trim()!=v.Trim()));
  view.CloseEditor();
  var targetType=dit.GetType().GetNestedType("ClipboardDataCellTarget",F);var list=(IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(targetType));
  foreach(int row in rows)list.Add(Activator.CreateInstance(targetType,new object[]{index,row,ci}));
  Call(dit,"ApplyPasteMatrix",index,ds,new List<string[]>{new[]{value}},rows[0],ci,false,list);Pump();
  bool broadcast=cells.All(c=>c.Value.TextValue.Trim()==value.Trim());
  dynamic undo=model.ChangeManager.Undo();bool restored=!undo.BError&&cells.Select((c,i)=>c.Value.Equals(old[i])).All(b=>b);
  Console.WriteLine("PASTE single-dropdown-to-two="+broadcast+" grouped-undo="+restored);
  if(!broadcast||!restored)throw new Exception("Dropdown broadcast/undo failed");
  // Dismiss only this runner's expected rejection dialog, not an external app.
  int rejected=0;using(var timer=new System.Windows.Forms.Timer()){timer.Interval=100;timer.Tick+=(s,e)=>{
   foreach(Form f in Application.OpenForms.Cast<Form>().ToArray())if(f.Text=="Paste rejected"){rejected++;f.DialogResult=DialogResult.OK;f.Close();}
  };timer.Start();
   Call(dit,"ApplyPasteMatrix",index,ds,new List<string[]>{new[]{value},new[]{"__INVALID_REVIEW_CATEGORY__"}},rows[0],ci,false,null);
   timer.Stop();
  }
  bool unchanged=cells.Select((c,i)=>c.Value.Equals(old[i])).All(b=>b);
  Console.WriteLine("PASTE mixed-valid-invalid rejected="+rejected+" all-destinations-unchanged="+unchanged);
  if(rejected!=1||!unchanged)throw new Exception("Invalid dropdown paste was not atomic");
  // Change an existing dependency through the same guarded paste service.
  // Observe cached colour versus workbook colour, then undo; do not paint it.
  var descriptionPoint=Items(Call(dit,"NavigationPositions")).First(p=>Field(p,"View")==view&&Field(p,"Header")==null&&Caption(p)=="Description");
  int descriptionColumn=(int)Call(dit,"GetGridColumnIndex",Field(descriptionPoint,"Column"));
  int descriptionRow=view.GetDataSourceRowIndex((int)Field(descriptionPoint,"RowHandle"));
  dynamic descriptionData=((dynamic)ds).DataRows[descriptionRow].DataCells[descriptionColumn];
  var descriptionCell=workbook.Worksheets[(string)descriptionData.SourceSheet].Cells[(string)descriptionData.SourceAddress];var oldDescription=descriptionCell.Value;
  dynamic categoryData=((dynamic)ds).DataRows[rows[0]].DataCells[ci];
  string before=S(categoryData.BGColor);
  Call(dit,"ApplyPasteMatrix",index,ds,new List<string[]>{new[]{""}},descriptionRow,descriptionColumn,false,null);Pump();
  string cached=S(categoryData.BGColor),actual=S(cells[0].Fill.BackgroundColor);
  Console.WriteLine("CF description-cleared="+descriptionCell.Value.IsEmpty+" category="+cells[0].GetReferenceA1()+" before="+before+" cached-after="+cached+" workbook-after="+actual+" stale="+(cached!=actual));
  var affected=cells[0];var engine=workbook.FormulaEngine;
  foreach(var rule in affected.Worksheet.ConditionalFormattings.OfType<FormulaExpressionConditionalFormatting>().Where(r=>affected.RowIndex>=r.Range.TopRowIndex&&affected.RowIndex<=r.Range.BottomRowIndex&&affected.ColumnIndex>=r.Range.LeftColumnIndex&&affected.ColumnIndex<=r.Range.RightColumnIndex)){
   var origin=new ExpressionContext(rule.Range.LeftColumnIndex,rule.Range.TopRowIndex,affected.Worksheet);
   var expression=engine.Parse(rule.Expression,origin);origin.ReferenceStyle=ReferenceStyle.R1C1;
   var at=new ExpressionContext(affected.ColumnIndex,affected.RowIndex,affected.Worksheet){ReferenceStyle=ReferenceStyle.R1C1};
   Console.WriteLine("CF_RULE expression="+rule.Expression+" applies="+engine.Evaluate(expression.ToString(origin),at)+" rule-fill="+rule.Formatting.Fill.BackgroundColor+" rule-pattern="+rule.Formatting.Fill.PatternType+" cached="+cached);
  }
  undo=model.ChangeManager.Undo();if(undo.BError||!descriptionCell.Value.Equals(oldDescription))throw new Exception("Conditional-format probe undo failed");Pump();
 }
}
