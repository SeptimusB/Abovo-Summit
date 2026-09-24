using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using DevExpress.Spreadsheet;

// Native, in-memory workbook only. No source workbook is opened or saved.
public static class UndoWorksheetCalculationFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static Assembly app; static int count;
 static void Check(bool ok,string message) { if(!ok) throw new Exception(message); Console.WriteLine("PASS "+(++count)+" "+message); }
 static double Number(Worksheet sheet,string cell) { return sheet.Cells[cell].Value.NumericValue; }
 static long UserRevision(object model) { return (long)model.GetType().GetProperty("UserChangeRevision",F).GetValue(model,null); }
 static dynamic Change(dynamic model,Worksheet sheet,string cell,double value) {
  dynamic change=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));
  change.ModelID=(int)model.ModelID; change.WSName=sheet.Name; change.CellAddress=cell;
  change.ChangedValue=value; change.DataFormat="N"; change.Description="Native undo calculation regression";
  return model.ChangeManager.ProcessChange(change);
 }
 [STAThread] public static int Main(string[] args) { try {
  Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string path=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(path)?Assembly.LoadFrom(path):null;};
  app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  Type files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  Type type=app.GetType("Abovo.FileManager+ExcelModel");var constructor=type.GetConstructors().Single();
  dynamic model=constructor.Invoke(new object[]{0,Enum.ToObject(constructor.GetParameters()[1].ParameterType,0)});
  Array models=Array.CreateInstance(type,1);models.SetValue(model,0);files.GetField("ExcelModels").SetValue(null,models);files.GetField("ExcelModelCount").SetValue(null,0);
  IWorkbook book=model.WB;book.Options.CalculationMode=WorkbookCalculationMode.Manual;
  Worksheet source=book.Worksheets[0];source.Name="Inputs";source.Cells["A1"].Value=10;source.Cells["B1"].FormulaInvariant="=A1*2";source.Cells["C1"].FormulaInvariant="=IF(A1=10,\"Locked\",\"Open\")";
  Worksheet second=book.Worksheets.Add("Other inputs");second.Cells["A1"].Value=2;second.Cells["B1"].FormulaInvariant="=A1*3";
  Worksheet report=book.Worksheets.Add("Current report");report.Cells["A1"].FormulaInvariant="=Inputs!B1+'Other inputs'!B1";
  book.CalculateFull();
  dynamic manager=Activator.CreateInstance(app.GetType("Abovo.ModelChangeManagerV2"),new object[]{0});model.ChangeManager=manager;
  dynamic engine=model.WBCalcEngine;int active=engine.AddActiveObject(new object());engine.AddActiveWorksheet(active,source,false);
  var mode=book.Options.CalculationMode;var engineType=book.Options.CalculationEngineType;object custom=model.WBCalculationService;bool deferred=custom!=null&&((dynamic)custom).DontCalcTDBS;
  Check(Change(model,source,"A1",15).BSuccess&&Number(source,"B1")==30,"Registered current worksheet edit calculates");
  Check(manager.Undo().BSuccess&&Number(source,"A1")==10&&Number(source,"B1")==20&&source.Cells["C1"].Value.TextValue=="Locked","Registered current sheet Undo refreshes formula and editability-driving result");
  Check(manager.Redo().BSuccess&&Number(source,"B1")==30&&source.Cells["C1"].Value.TextValue=="Open","Registered current sheet Redo refreshes formulas");
  engine.RemoveActiveObject(active);
  Check(engine.ActiveWorksheetRegistrationCount==0,"No interface worksheet remains registered");
  dynamic result=manager.Undo();
  if(args.Length>4&&args[4]=="reproduce") {
   Check(result.BSuccess&&Number(source,"A1")==10&&Number(source,"B1")==30,"Original bug reproduced: Undo restores input but leaves unregistered sheet formula stale");return 0;
  }
  Check(result.BSuccess&&Number(source,"A1")==10&&Number(source,"B1")==20,"Unregistered source worksheet Undo calculates restored values");
  Check(manager.Redo().BSuccess&&Number(source,"B1")==30,"Unregistered source worksheet Redo calculates restored values");
  active=engine.AddActiveObject(new object());engine.AddActiveWorksheet(active,report,false);
  Check(manager.Undo().BSuccess&&Number(source,"B1")==20&&Number(report,"A1")==26,"Undo calculates restored source before registered current report refresh");
  Check(manager.Redo().BSuccess&&Number(source,"B1")==30&&Number(report,"A1")==36,"Redo calculates source before registered current report refresh");
  using((IDisposable)manager.BeginChangeGroup("Two worksheets")) {
   Check(Change(model,source,"A1",20).BSuccess,"Grouped first source edit");
   Check(Change(model,second,"A1",4).BSuccess,"Grouped second source edit");
  }
  source.Calculate();second.Calculate();report.Calculate();
  Check(manager.Undo().BSuccess&&Number(source,"B1")==30&&Number(second,"B1")==6&&Number(report,"A1")==36,"Grouped multi-sheet Undo completes before report refresh");
  Check(manager.Redo().BSuccess&&Number(source,"B1")==40&&Number(second,"B1")==12&&Number(report,"A1")==52,"Grouped multi-sheet Redo completes before report refresh");
  long revision=UserRevision(model);bool disturbed=false;
  EventHandler interfere=(s,e)=>{if(!disturbed){disturbed=true;source.Cells["A1"].Value=999;}};
  EventInfo completed=((object)engine).GetType().GetEvent("CalculationCompleted");completed.AddEventHandler(engine,interfere);
  result=manager.Undo();completed.RemoveEventHandler(engine,interfere);
  Check(result.BError&&!result.BSuccess&&disturbed,"Refresh-time mutation rejects the undo transaction");
  Check(Number(source,"A1")==20&&Number(source,"B1")==40&&Number(second,"A1")==4&&Number(second,"B1")==12&&Number(report,"A1")==52,"Rollback recalculates restored source sheets and current report");
  Check((bool)manager.CanUndo&&!(bool)manager.CanRedo&&UserRevision(model)==revision,"Failed undo retains history stacks and committed-user revision");
  Check(book.Options.CalculationMode==mode&&book.Options.CalculationEngineType==engineType&&(custom==null||((dynamic)custom).DontCalcTDBS==deferred),"Sheet calculation preserves engine, mode and deferred-sheet policy");
  Check(manager.Undo().BSuccess&&Number(report,"A1")==36,"Undo remains usable after a restored rollback");
  Console.WriteLine("PASS "+count+" assertions; no workbook opened or saved.");
  model.ModelSpreadsheetControl.Dispose();return 0;
 } catch(Exception ex) { Console.Error.WriteLine(ex);return 1; } }
}
