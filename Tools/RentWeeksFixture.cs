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
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;

class RentWeeksFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static object Field(object o,string n){var f=o.GetType().GetField(n,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(n,F).GetValue(o,null);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static int count;
 static void Check(bool value,string message){if(!value)throw new Exception(message);count++;Console.WriteLine("PASS "+message);}
 static IEnumerable<Control> Controls(Control c){foreach(Control child in c.Controls){yield return child;foreach(var d in Controls(child))yield return d;}}
 static void Pump(){Application.DoEvents();}
 static bool Number(Cell c,double n){return c.Value.IsNumeric&&Math.Abs(c.Value.NumericValue-n)<1e-10;}
 static void Paste(object dit,int index,object data,int column,string[] values){Call(dit,"ApplyPasteMatrix",index,data,values.Select(v=>new[]{v}).ToList(),0,column,false,null);Pump();}
 [STAThread] static int Main(string[] args){try {
  System.Threading.Thread.CurrentThread.CurrentCulture=CultureInfo.GetCultureInfo("en-GB");
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  string copy=Path.Combine(args[1],"private-rent.xlsb");File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!result.BError,"Private workbook opens");
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)result.IntegerReturn);
  try {
   IWorkbook w=(IWorkbook)model.WB;
   foreach(string n in new[]{"RentWks","Rep_Rent_02","Rep_Rent_03"}){
    var r=w.DefinedNames.GetDefinedName(n).Range;var c=r[0,0];
    Console.WriteLine(n+" "+r.GetReferenceA1()+" value="+c.Value+" pattern="+c.Fill.PatternType+" locked="+c.Protection.Locked+" format="+c.NumberFormat);
    foreach(var rule in c.Worksheet.DataValidations.GetDataValidations(c))Console.WriteLine("VALIDATION "+rule.ValidationType+" "+rule.Operator+" "+rule.Criteria+" "+rule.Criteria2);
   }
   if(args[3]=="True")return 0;
   var standard=w.DefinedNames.GetDefinedName("RentWks").Range[0,0];
   var exception=w.DefinedNames.GetDefinedName("Rep_Rent_03").Range[0,0];
   string fmt1=standard.NumberFormat,fmt2=exception.NumberFormat;
   bool sheetProtection=standard.Worksheet.IsProtected;
   using(var form=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})){
    form.Opacity=0;form.ShowInTaskbar=false;form.ClientSize=new Size(1750,1000);form.Show();Pump();
    Call(form,"ShowInterface",(int)model.ModelID,2,false,"None",null,-1);Pump();var dit=(Control)Field(form,"ActiveInterface");
    var editor=Controls(dit).OfType<BaseEdit>().First(c=>c.GetType().Name=="AbovoDETextEdit"&&(string)Field(c,"TargetCell")==standard.GetReferenceA1());
    var tag=Field(editor,"Tag");Check((string)Field(tag,"DataType")=="R"&&(bool)Field(tag,"MinExclusive"),"Standard is a positive real-number input");
    foreach(double value in new[]{0.125,49.5,52.142857,54.5,1000.25}){
     var old=standard.Value;editor.EditValue=value;Call(editor,"MarkDirty");Call(dit,"SingleCell_Value_Push",editor,EventArgs.Empty);Pump();
     Check(Number(standard,value),"Standard native input retains "+value);
     Check(model.ChangeManager.Undo().BSuccess&&standard.Value.Equals(old),"Standard Undo");Pump();
     Check(model.ChangeManager.Redo().BSuccess&&Number(standard,value),"Standard Redo");Pump();
    }
    foreach(object value in new object[]{0d,-0.5d,Double.NaN,Double.PositiveInfinity,"not a number"}){
     var old=standard.Value;editor.EditValue=value;Call(editor,"MarkDirty");Call(dit,"SingleCell_Value_Push",editor,EventArgs.Empty);
     Check(standard.Value.Equals(old)&&!string.IsNullOrEmpty(editor.ErrorText),"Standard invalid input rejected without writing: "+value);
     Call(editor,"RefreshData");editor.ErrorText="";
    }
    var view=Controls(dit).OfType<GridControl>().Select(g=>g.MainView).OfType<GridView>().First();
    dynamic source=view.GridControl.DataSource;int index=(int)source.UBSTag.DSIndex;var data=Field(view.Tag,"DataSet");
    var columns=(Array)Field(data,"DataColumns");int column=1;var weekTag=Field(columns.GetValue(column),"ColumnTag");
    Check((string)Field(weekTag,"DataType")=="R"&&(bool)Field(weekTag,"MinExclusive"),"Exception is a positive real-number input");
    // Choose an actual offered year, leaving the original workbook untouched.
    view.FocusedRowHandle=0;view.FocusedColumn=view.Columns[0];view.ShowEditor();
    var yearEditor=view.ActiveEditor as ComboBoxEdit;Check(yearEditor!=null,"Exception year uses existing dropdown");
    string year=yearEditor.Properties.Items.Cast<object>().Select(Convert.ToString).First(s=>!string.IsNullOrWhiteSpace(s)&&s!="<Blank>");
    yearEditor.EditValue=year;Check(view.PostEditor()&&view.UpdateCurrentRow(),"Exception year selected");view.CloseEditor();Pump();
    Check((bool)Call(dit,"CanPasteToDataPoint",data,0,column),"Selected year enables exception weeks under existing fill rule");
    foreach(double value in new[]{0.125,49.5,52.142857,54.5,1000.25}){
     view.FocusedRowHandle=0;view.FocusedColumn=view.Columns[column];view.ShowEditor();Check(view.ActiveEditor!=null,"Exception editor opens");
     view.ActiveEditor.EditValue=value;Check(view.PostEditor()&&view.UpdateCurrentRow(),"Exception editor posts");view.CloseEditor();Pump();
     Check(Number(exception,value),"Exception native input retains "+value);
    }
    Check(view.GetRowCellDisplayText(0,view.Columns[column]).Contains("1000.25"),"Exception displays the decimal value");
    var validate=dit.GetType().GetMethod("NumericInputError",F);
    foreach(double value in new[]{0d,-0.5d,Double.NaN,Double.PositiveInfinity})Check(!string.IsNullOrEmpty((string)validate.Invoke(null,new object[]{exception,weekTag,value})),"Grid commit/paste gate rejects "+value);
    var beforePaste=exception.Value;Paste(dit,index,data,column,new[]{"52.142857"});Check(Number(exception,52.142857),"Decimal clipboard input preserved");
    Check(model.ChangeManager.Undo().BSuccess&&exception.Value.Equals(beforePaste),"Paste Undo");Pump();
    Check(model.ChangeManager.Redo().BSuccess&&Number(exception,52.142857),"Paste Redo");Pump();
    var second=w.DefinedNames.GetDefinedName("Rep_Rent_03").Range[1,0];var firstValue=exception.Value;var secondValue=second.Value;
    // Enable a second row using the same valid year, then reject a mixed paste atomically.
    dynamic yearChange=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));yearChange.ModelID=(int)model.ModelID;yearChange.WSName=exception.Worksheet.Name;yearChange.CellAddress=w.DefinedNames.GetDefinedName("Rep_Rent_02").Range[1,0].GetReferenceA1();yearChange.ChangedValue=year;yearChange.DataFormat="S";yearChange.Description="Private rent test year";Check(model.ChangeManager.ProcessChange(yearChange).BSuccess,"Second exception year enabled");Pump();
    int rejected=0;using(var timer=new System.Windows.Forms.Timer()){timer.Interval=100;timer.Tick+=(s,e)=>{foreach(Form f in Application.OpenForms.Cast<Form>().ToArray())if(f.Text=="Paste rejected"){rejected++;f.DialogResult=DialogResult.OK;f.Close();}};timer.Start();Paste(dit,index,data,column,new[]{"53.25","0"});timer.Stop();}
    Check(rejected==1&&exception.Value.Equals(firstValue)&&second.Value.Equals(secondValue),"Mixed valid/zero paste rejects all writes");
    // Clear the prerequisite through history and verify fill-based editability is unchanged.
    yearChange.CellAddress=w.DefinedNames.GetDefinedName("Rep_Rent_02").Range[0,0].GetReferenceA1();yearChange.ChangedValue=null;
    Check(model.ChangeManager.ProcessChange(yearChange).BSuccess,"Year clearing succeeds");Pump();
    Check(!(bool)Call(dit,"CanPasteToDataPoint",data,0,column),"Cleared year still locks exception weeks");
    Check(model.ChangeManager.Undo().BSuccess,"Restore exception year with Undo");Pump();
    Check(standard.NumberFormat==fmt1&&exception.NumberFormat==fmt2&&standard.Worksheet.IsProtected==sheetProtection,"Workbook formatting and protection unchanged");
    string saved=Path.Combine(args[1],"rent-weeks-saved.xlsb");Check((bool)model.SaveFileAsTo(saved,true),"Save separate XLSB");
    using(var reopened=new Workbook()){reopened.Options.CalculationMode=WorkbookCalculationMode.Manual;Check(reopened.LoadDocument(saved),"Saved XLSB reopens");Check(Number(reopened.Worksheets[standard.Worksheet.Name].Cells[standard.GetReferenceA1()],1000.25)&&Number(reopened.Worksheets[exception.Worksheet.Name].Cells[exception.GetReferenceA1()],52.142857),"Both decimal inputs survive save/reopen");}
    form.Close();
   }
  }finally{files.GetMethod("CloseModel",new[]{typeof(int)}).Invoke(null,new object[]{(int)model.ModelID});}
  Console.WriteLine("PASS "+count+" assertions");return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
}
