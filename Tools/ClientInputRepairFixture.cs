// Native DIT regression. All writes/saves are confined to the runner's private copy.
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
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraVerticalGrid;
using DevExpress.XtraVerticalGrid.Rows;

public static class ClientInputRepairFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static Assembly app; static IWorkbook book; static int checks;
 static Dictionary<string,CellValue> saved=new Dictionary<string,CellValue>();
 static Dictionary<string,string> formats=new Dictionary<string,string>();
 static object Field(object o,string n){if(o==null)return null;var f=o.GetType().GetField(n,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(n,F).GetValue(o,null);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static IEnumerable<object> Items(object o){return ((IEnumerable)o).Cast<object>();}
 static void Check(bool ok,string message){if(!ok)throw new Exception("FAIL "+message);checks++;Console.WriteLine("PASS "+message);Console.Out.Flush();}
 static void Pump(){Application.DoEvents();Application.RaiseIdle(EventArgs.Empty);Application.DoEvents();}
 static string Key(Cell c){return c.Worksheet.Name+"!"+c.GetReferenceA1();}
 sealed class Target {
  public object Dit,Point,Tag,Data;public Cell Cell;public int DS,Row,Column;public GridView View;public VGridControl Grid;public GridColumn GridColumn;public EditorRow VRow;
  public BaseEdit Editor{get{return View!=null?View.ActiveEditor:Grid.ActiveEditor;}}
  public string Display{get{return View!=null?View.GetRowCellDisplayText((int)Field(Point,"RowHandle"),GridColumn):Grid.GetCellDisplayText(VRow,(int)Field(Point,"Record"));}}
  public void Close(){if(View!=null)View.CloseEditor();else Grid.CloseEditor();}
 }
 static Target Resolve(object dit,object p){
  var t=new Target{Dit=dit,Point=p,View=Field(p,"View") as GridView,Grid=Field(p,"Host") as VGridControl,GridColumn=Field(p,"Column") as GridColumn,VRow=Field(p,"VRow") as EditorRow};
  if(t.View==null&&t.VRow==null)return null;
  t.Tag=t.View!=null?Field(t.GridColumn,"Tag"):Call(dit,"GetVGridColumnTag",t.VRow,0);
  dynamic source=t.View!=null?t.View.GridControl.DataSource:t.Grid.DataSource;t.DS=(int)source.UBSTag.DSIndex;
  t.Column=(int)(t.View!=null?Call(dit,"GetGridColumnIndex",t.GridColumn):Call(dit,"GetVGridColumnIndex",t.VRow,0));
  t.Row=t.View!=null?t.View.GetDataSourceRowIndex((int)Field(p,"RowHandle")):(int)Field(p,"Record");
  dynamic pres=Field(dit,"DataPres");t.Data=pres.DataSets[t.DS];dynamic dp=((dynamic)t.Data).DataRows[t.Row].DataCells[t.Column];
  t.Cell=book.Worksheets[(string)dp.SourceSheet].Cells[(string)dp.SourceAddress];return t;
 }
 static List<Target> Targets(object dit){return Items(Call(dit,"NavigationPositions")).Where(p=>Field(p,"Header")==null&&Field(p,"VHeader")==null&&Field(p,"Standalone")==null).Select(p=>Resolve(dit,p)).Where(t=>t!=null).ToList();}
 static bool Number(Cell c,double v){return c.Value.IsNumeric&&Math.Abs(c.Value.NumericValue-v)<1e-9;}
 static void Undo(dynamic model,Cell c,CellValue original){dynamic r=model.ChangeManager.Undo();Pump();Check(!r.BError&&c.Value.Equals(original),"Undo "+Key(c));}
 static void Redo(dynamic model,Cell c,CellValue expected){dynamic r=model.ChangeManager.Redo();Pump();Check(!r.BError&&c.Value.Equals(expected),"Redo "+Key(c));}
 static void Remember(Cell c){saved[Key(c)]=c.Value;formats[Key(c)]=c.NumberFormat;}
 static void Edit(Target t,dynamic model,double value,bool money){
  var old=t.Cell.Value;string fmt=t.Cell.NumberFormat;bool protection=t.Cell.Worksheet.IsProtected;
  Check(!t.Cell.Protection.Locked&&!t.Cell.HasFormula,"Eligible workbook input "+Key(t.Cell));
  Call(t.Dit,"ActivateEditorPosition",t.Point);Pump();Check(t.Editor!=null,"Native editor opens "+Key(t.Cell));
  t.Editor.EditValue=value;Check(t.View!=null?t.View.PostEditor()&&t.View.UpdateCurrentRow():t.Grid.PostEditor(),"Editor posts "+Key(t.Cell));t.Close();Pump();
  Check(Number(t.Cell,value),"Precise native value "+Key(t.Cell)+"="+value);
  if(money)Check(t.Display==t.Cell.DisplayText,"Workbook display "+Key(t.Cell)+"="+t.Display);
  var after=t.Cell.Value;Undo(model,t.Cell,old);Redo(model,t.Cell,after);
  Check(t.Cell.NumberFormat==fmt&&t.Cell.Worksheet.IsProtected==protection,"Format/protection unchanged "+Key(t.Cell));
  Check((bool)model.IsDirty&&(bool)model.ChangeManager.CanUndo,"Input dirty/history retained");Remember(t.Cell);
 }
 static void Paste(Target t,string text){
  var type=t.Dit.GetType().GetNestedType("ClipboardDataCellTarget",F);var list=(IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(type));
  list.Add(Activator.CreateInstance(type,new object[]{t.DS,t.Row,t.Column}));
  Call(t.Dit,"ApplyPasteMatrix",t.DS,t.Data,new List<string[]>{new[]{text}},t.Row,t.Column,false,list,null);Pump();
 }
 static void PasteRoundTrip(Target t,dynamic model,double value,string text){var old=t.Cell.Value;Paste(t,text);Check(Number(t.Cell,value),"Typed paste "+Key(t.Cell)+"="+text);var after=t.Cell.Value;Undo(model,t.Cell,old);Redo(model,t.Cell,after);Remember(t.Cell);}
 static void RejectPaste(Target t,string text){
  var old=t.Cell.Value;int rejected=0;
  using(var timer=new System.Windows.Forms.Timer()){timer.Interval=80;timer.Tick+=(s,e)=>{foreach(Form f in Application.OpenForms.Cast<Form>().ToArray())if(f.Text=="Paste rejected"){rejected++;f.DialogResult=DialogResult.OK;f.Close();}};timer.Start();Paste(t,text);timer.Stop();}
  Check(rejected==1&&t.Cell.Value.Equals(old),"Invalid numeric paste rejected unchanged: "+text);
 }
 static void BatchAndLimits(Target first,Target second,dynamic model){
  var old1=first.Cell.Value;var old2=second.Cell.Value;int rejected=0;int history=(int)Field(Field((object)model.ChangeManager,"UndoStack"),"Count");
  using(var timer=new System.Windows.Forms.Timer()){timer.Interval=80;timer.Tick+=(s,e)=>{foreach(Form f in Application.OpenForms.Cast<Form>().ToArray())if(f.Text=="Paste rejected"){rejected++;f.DialogResult=DialogResult.OK;f.Close();}};timer.Start();
   Call(first.Dit,"ApplyPasteMatrix",first.DS,first.Data,new List<string[]>{new[]{"-2.5%"},new[]{"1e25"}},first.Row,first.Column,false,null,null);timer.Stop();
  }
  Check(rejected==1&&first.Cell.Value.Equals(old1)&&second.Cell.Value.Equals(old2),"Mixed valid/invalid percentage paste is atomic");
  Check((int)Field(Field((object)model.ChangeManager,"UndoStack"),"Count")==history,"Rejected batch adds no history entry");
  Call(first.Dit,"ApplyPasteMatrix",first.DS,first.Data,new List<string[]>{new[]{"-2.5%"},new[]{"115%"}},first.Row,first.Column,false,null,null);Pump();
  Check(Number(first.Cell,-0.025)&&Number(second.Cell,1.15),"Mixed negative/above-100 percentage batch");
  dynamic undo=model.ChangeManager.Undo();Pump();Check(!undo.BError&&first.Cell.Value.Equals(old1)&&second.Cell.Value.Equals(old2),"One Undo restores numeric paste batch");
  var validate=first.Dit.GetType().GetMethod("NumericInputError",F);dynamic tag=Activator.CreateInstance(first.Tag.GetType());tag.DataType="P";
  Func<Cell,double,string> error=(c,v)=>(string)validate.Invoke(null,new object[]{c,tag,v});
  using(var sample=new Workbook()){
   var ws=sample.Worksheets[0];ws.Cells["A1"].Value=-0.1;ws.Cells["A2"].Value=0.5;
   ws.DataValidations.Add(ws.Range["B1:B2"],DataValidationType.Decimal,DataValidationOperator.Between,"=$A1",2);
   Check(error(ws.Cells["B1"],0)==null&&error(ws.Cells["B2"],0)!=null,"Relative workbook numeric bounds resolved at each target");
   tag.MinVal="-0.05";tag.MaxVal="1.5";
   Check(error(ws.Cells["B1"],-0.06)!=null&&error(ws.Cells["B1"],1.6)!=null&&error(ws.Cells["B1"],1.1)==null,"Explicit XML bounds retained");
   tag.MinVal="NOMIN";tag.MaxVal="NOMAX";
   Check(error(ws.Cells["B1"],-0.075)==null&&error(ws.Cells["B1"],3)!=null,"NOMIN/NOMAX never remove workbook bounds");
   Check(error(ws.Cells["B1"],Double.NaN)!=null&&error(ws.Cells["B1"],Double.PositiveInfinity)!=null,"Non-finite input rejected");
   Check(ws.Cells["B1"].Value.IsEmpty&&ws.Cells["B2"].Value.IsEmpty,"Validation does not trial-write worksheet cells");
  }
 }
 static IEnumerable<Control> Children(Control c){foreach(Control child in c.Controls){yield return child;foreach(var nested in Children(child))yield return nested;}}
 [STAThread] public static int Main(string[] args){try{
  System.Threading.Thread.CurrentThread.CurrentCulture=CultureInfo.GetCultureInfo("en-GB");
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  string copy=Path.Combine(args[1],"private-input-repair.xlsb");File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});Check(!result.BError,"Open private workbook");
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)result.IntegerReturn);book=(IWorkbook)model.WB;
  var cases=args.Length>4&&!String.IsNullOrEmpty(args[4])?args[4].Split(',').Select(Int32.Parse).ToArray():new[]{3,15,16,26,30,32,33,35,39};
  try{
   // The fee editor deliberately requires a description. Seed it through the
   // production typed/history service, never by unlocking or direct assignment.
   if(cases.Contains(33)){
    Check(!book.Worksheets["Funding Assumptions"].Cells["B238"].Protection.Locked,"Current-template fee-description coordinate is an input; older layouts need a separate mapping test");
    dynamic setup=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));setup.ModelID=(int)model.ModelID;setup.WSName="Funding Assumptions";setup.CellAddress="B238";setup.ChangedValue="Private regression fee";setup.DataFormat="S";setup.Description="Private fee input setup";
    Check(model.ChangeManager.ProcessChange(setup).BSuccess,"Seed private fee description via ChangeManager");
   }
   var wanted=new HashSet<string>{"Service Charge Assumptions!E19","Dvpt BP Rev and Exp Assumptions!D8","Dvpt BP Rev and Exp Assumptions!D16","Dvpt BP Rev and Exp Assumptions!D25","Dvpt BP Rev and Exp Assumptions!D34","Housing Asset Assumptions!D14","Housing Asset Assumptions!D59","Housing Asset Assumptions!D61","Housing Asset Assumptions!D69","Development BP Assumptions!O169","Development BP Assumptions!O170"};
   for(int i=1;i<=5;i++){var r=book.DefinedNames.GetDefinedName("Rep_SCond_0"+i).Range;wanted.Add(Key(r.Worksheet.Cells[r.TopRowIndex,r.LeftColumnIndex]));}
   var owners=new Dictionary<string,int>{{"Service Charge Assumptions",3},{"Stock Condition Inputs",16},{"Development BP Assumptions",26},{"Dvpt BP Rev and Exp Assumptions",30},{"Housing Asset Assumptions",39}};
   wanted.RemoveWhere(k=>!cases.Contains(owners[k.Substring(0,k.LastIndexOf('!'))]));
   var tested=new HashSet<string>();bool dateTested=false,covenantTested=false,yearTested=false;int economic=0;
   using(var form=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})){
    form.Opacity=0;form.ShowInTaskbar=false;form.ClientSize=new Size(1750,1000);form.Show();Pump();
    foreach(int csid in cases){
     Console.WriteLine("CASE="+csid);Call(form,"ShowInterface",0,csid,false,"None",null,-1);Pump();var dit=(Control)Field(form,"ActiveInterface");dynamic tabs=Field(dit,"XtraTabControlNewGIT");
     for(int ti=0;ti<tabs.TabPages.Count;ti++){
      string caption=Convert.ToString(tabs.TabPages[ti].Text).Trim();if(caption.StartsWith("XtraTabPage"))continue;
      Call(dit,"BuildSection",ti,false,false);tabs.SelectedTabPageIndex=ti;Pump();var targets=Targets(dit);
      foreach(var t in targets.Where(t=>wanted.Contains(Key(t.Cell))))if(tested.Add(Key(t.Cell))){Edit(t,model,1234.56,true);PasteRoundTrip(t,model,2345.67,"2345.67");}
      if(csid==32){var t=targets.FirstOrDefault(candidate=>Convert.ToString(Field(candidate.Tag,"DataType"))=="P");if(t!=null){Edit(t,model,-0.025,false);PasteRoundTrip(t,model,-0.03125,"-3.125%");economic++;}}
      if(csid==35&&!covenantTested){
       dynamic data=Field(dit,"DataPres");dynamic ds=data.DataSets[0];dynamic a=ds.DataRows[0].DataCells[0];dynamic b=ds.DataRows[0].DataCells[1];
       Check((string)a.SourceAddress=="A8"&&(string)b.SourceAddress=="B8","Covenant years map to A8/B8");
       Check(!(bool)Call(dit,"CanPasteToDataPoint",ds,0,0)&&!(bool)Call(dit,"CanPasteToDataPoint",ds,0,1),"Covenant year columns stay read-only");
       foreach(string address in new[]{"D8","E8","F8","G8","H8"}){var t=targets.First(candidate=>candidate.Cell.GetReferenceA1()==address);Edit(t,model,address=="G8"||address=="H8"?1234.567:1.125,address=="G8"||address=="H8");PasteRoundTrip(t,model,address=="G8"||address=="H8"?2345.678:1.15,address=="G8"||address=="H8"?"2345.678":"115%");}
       RejectPaste(targets.First(t=>t.Cell.GetReferenceA1()=="F8"),"1e25");
       BatchAndLimits(targets.First(t=>t.Cell.GetReferenceA1()=="F8"),targets.First(t=>t.Cell.GetReferenceA1()=="F9"),model);covenantTested=true;
      }
      if(csid==33&&!yearTested){var t=targets.FirstOrDefault(candidate=>candidate.Cell.GetReferenceA1()=="C238");if(t!=null){Check(Convert.ToString(Field(t.Tag,"DataType"))=="I","One-off fee is a BP year");Edit(t,model,5,false);PasteRoundTrip(t,model,6,"6");RejectPaste(t,"41");RejectPaste(t,"6.5");yearTested=true;}}
      if(csid==15&&!dateTested){var editor=Children(dit).OfType<DateEdit>().FirstOrDefault(c=>c.GetType().Name=="AbovoDEDateEdit");if(editor!=null){
       var cell=book.Worksheets["Repairs & Maint. Assumptions"].Cells[Convert.ToString(Field(editor,"TargetCell"))];var old=cell.Value;
       Check(!cell.Value.IsEmpty||editor.EditValue==null,"Empty survey date displays empty");
       editor.EditValue=new DateTime(2026,10,31);Pump();Check(cell.Value.DateTimeValue==new DateTime(2026,10,31),"Survey date posts actual date");var after=cell.Value;Undo(model,cell,old);Redo(model,cell,after);
       editor.EditValue=null;Pump();Check(cell.Value.IsEmpty&&editor.EditValue==null,"Clearing survey date remains blank");Undo(model,cell,after);Remember(cell);dateTested=true;
      }}
     }
    }
    Check(wanted.All(tested.Contains),"All requested monetary targets tested; missing="+String.Join(",",wanted.Except(tested)));
    Check((!cases.Contains(15)||dateTested)&&(!cases.Contains(35)||covenantTested)&&(!cases.Contains(33)||yearTested)&&(!cases.Contains(32)||economic>=10),"Requested date, covenant, fee-year and economic sections covered");
    string path=Path.Combine(args[1],"input-repairs-saved.xlsb");Check((bool)model.SaveFileAsTo(path,true),"Normal Summit save to separate XLSB");
    using(var reopened=new Workbook()){reopened.Options.CalculationMode=WorkbookCalculationMode.Manual;Check(reopened.LoadDocument(path),"Saved workbook reopens independently");foreach(var pair in saved){int split=pair.Key.LastIndexOf('!');var cell=reopened.Worksheets[pair.Key.Substring(0,split)].Cells[pair.Key.Substring(split+1)];Check(cell.Value.Equals(pair.Value)&&cell.NumberFormat==formats[pair.Key],"Reopen value/format "+pair.Key);}}
    form.Close();
   }
  }finally{files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});}
  Console.WriteLine("PASS CHECKS="+checks);return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
