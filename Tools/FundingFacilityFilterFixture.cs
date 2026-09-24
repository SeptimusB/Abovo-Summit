using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.XtraVerticalGrid;
using DevExpress.XtraVerticalGrid.Rows;

// Real DIT with a disposable private workbook copy. No Save/SaveAs.
class FundingFacilityFilterFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static int checks;
 static object Field(object o,string n){var f=o.GetType().GetField(n,F);return f!=null?f.GetValue(o):o.GetType().GetProperty(n,F).GetValue(o,null);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static IEnumerable<Control> Controls(Control c){foreach(Control child in c.Controls){yield return child;foreach(var inner in Controls(child))yield return inner;}}
 static IEnumerable<BaseRow> Rows(IEnumerable rows){foreach(BaseRow row in rows){yield return row;foreach(var child in Rows(row.ChildRows))yield return child;}}
 static void Check(bool value,string text){if(!value)throw new Exception(text);Console.WriteLine("PASS "+(++checks)+" "+text);}
 [STAThread]static int Main(string[] args){try{
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  string copy=Path.Combine(args[1],"private-facility-filter.xlsb");File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!loaded.BError,"Disposable model opens");dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)loaded.IntegerReturn);
  IWorkbook wb=model.WB;var funders=wb.DefinedNames.GetDefinedName("Rep_Fund_03").Range;var facilities=wb.DefinedNames.GetDefinedName("FacilityNames").Range;var loans=wb.DefinedNames.GetDefinedName("LoanDescs").Range;
  facilities[0,1].Value="Shared fixture facility";facilities[0,3].Value="Shared fixture facility";funders[0,1].Value="First fixture funder";funders[0,3].Value="Other fixture funder";loans[0,1].Value="Loan second";loans[0,3].Value="Loan fourth";
  var sourceBefore=Enumerable.Range(0,facilities.ColumnCount).Select(i=>facilities[0,i].Value.ToString()+"|"+funders[0,i].Value.ToString()+"|"+loans[0,i].Value.ToString()).ToArray();
  using(var host=(Form)Activator.CreateInstance(app.GetType("GroupInterfaceTemplate"),new object[]{(int)model.ModelID,0,"Normal"})){
   host.Opacity=0;host.ShowInTaskbar=false;host.ClientSize=new Size(1700,1000);host.Show();Application.DoEvents();
   Call(host,"ShowInterface",(int)model.ModelID,33,false,"None",null,1);Application.DoEvents();var dit=(Control)Field(host,"ActiveInterface");
   foreach(var tabs in Controls(dit).OfType<DevExpress.XtraTab.XtraTabControl>())tabs.SelectedTabPageIndex=1;
   Call(dit,"BuildSection",1,false,false);Application.DoEvents();
   var grid=Controls(dit).OfType<VGridControl>().First(g=>g.Visible&&Rows(g.Rows).Any(r=>r.Properties.Caption.Contains("Facility")&&r.Properties.Caption.Contains("Name")));
   var facilityRow=Rows(grid.Rows).First(r=>r.Properties.Caption.Contains("Facility")&&r.Properties.Caption.Contains("Name"));int originalCount=grid.RecordCount;
   Check(originalCount>4,"Funding initially displays all loan columns");
   grid.FocusedRow=facilityRow;grid.FocusedRecord=1;Call(dit,"SetClipboardTarget",grid);Call(dit,"UpdateFundingScheduleMenu");
   var facilityMenu=(ToolStripMenuItem)Field(dit,"FundingFacilityMenu");
   Check(facilityMenu.Available&&facilityMenu.Enabled&&facilityMenu.Text=="Filter to facility Shared fixture facility","Actual DIT menu offers the clicked loan's facility filter");
   facilityMenu.PerformClick();Application.DoEvents();
   Check(grid.RecordCount==2&&grid.GetDataSourceRecordIndex(0)==1&&grid.GetDataSourceRecordIndex(1)==3,"Facility menu hides other columns and retains nonadjacent physical loans 2 and 4");
   Check(grid.GetRecordHeaderText(0)=="Loan second"&&grid.GetRecordHeaderText(1)=="Loan fourth","Filtered fixed headers retain correct loan names");
   Check((bool)Field(Field(dit,"FundingFilterButton"),"Visible"),"Applied filter warning is shown");
   grid.FocusedRecord=1;grid.FocusedRow=facilityRow;object[] context={false};var identity=dit.GetType().GetMethod("ContextFundingFacility",F).Invoke(dit,context);
   Check((int)Field(identity,"ColumnIndex")==loans.LeftColumnIndex+3&&(string)Field(identity,"FunderName")=="Other fixture funder","Filtered context still maps to fourth physical loan, not second source record");
   Call(dit,"UpdateFundingScheduleMenu");((ToolStripMenuItem)Field(dit,"FundingFunderMenu")).PerformClick();Application.DoEvents();
   Check(grid.RecordCount==1&&grid.GetDataSourceRecordIndex(0)==3,"Switching to funder filter uses the clicked filtered loan identity");
   Call(dit,"UpdateFundingScheduleMenu");var clear=(ToolStripMenuItem)Field(dit,"FundingClearMenu");Check(clear.Available,"Remove filters is available");clear.PerformClick();Application.DoEvents();
   Check(grid.RecordCount==originalCount&&!(bool)Field(Field(dit,"FundingFilterButton"),"Visible"),"Remove filters restores all physical loans and clears warning");
   Check(sourceBefore.SequenceEqual(Enumerable.Range(0,facilities.ColumnCount).Select(i=>facilities[0,i].Value.ToString()+"|"+funders[0,i].Value.ToString()+"|"+loans[0,i].Value.ToString())),"Filtering did not alter workbook identities or source order");
   host.Close();
  }
  model.ModelSpreadsheetControl.Dispose();Console.WriteLine("PASS "+checks+" assertions; original workbook untouched.");return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
}
