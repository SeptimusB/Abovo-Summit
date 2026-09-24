// Investigation only. Each invocation starts an identical fresh private AGL model.
using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Reflection.Emit;
using System.Diagnostics;
using System.Windows.Forms;
using System.Text;
using System.Security.Cryptography;
using System.Globalization;
using DevExpress.Spreadsheet;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;

class CheckSheetCalculationMatrix297 {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static Assembly app;static dynamic model;static IWorkbook book;static Form host;static Control dit;static int events,checks;static string mode;static int traceColumn=-1;
 static object Get(object o,string n){for(var t=o.GetType();t!=null;t=t.BaseType){var f=t.GetField(n,F|BindingFlags.DeclaredOnly);if(f!=null)return f.GetValue(o);var p=t.GetProperty(n,F|BindingFlags.DeclaredOnly);if(p!=null)return p.GetValue(o,null);}throw new MissingMemberException(n);}
 static object Call(object o,string n,params object[] a){return o.GetType().GetMethod(n,F).Invoke(o,a);}
 static object Static(string t,string n,params object[] a){return app.GetType(t).GetMethod(n,F).Invoke(null,a);}
 static void Check(bool ok,string text){if(!ok)throw new Exception(text);Console.WriteLine("PASS "+(++checks)+" "+text);}
 static IEnumerable<Control> Children(Control c){foreach(Control x in c.Controls){yield return x;foreach(var n in Children(x))yield return n;}}
 static string Value(Cell c){return c.Value.IsNumeric?c.Value.NumericValue.ToString("R",CultureInfo.InvariantCulture):c.Value.ToString();}
 static string Hash(string text){using(var h=SHA256.Create())return BitConverter.ToString(h.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-","");}
 static Form QuietHost(Type baseType){
  var ab=AppDomain.CurrentDomain.DefineDynamicAssembly(new AssemblyName("QuietMatrix"+Guid.NewGuid().ToString("N")),AssemblyBuilderAccess.Run);var tb=ab.DefineDynamicModule("Main").DefineType("QuietHost",TypeAttributes.Public,baseType);
  var sig=new[]{typeof(int),typeof(int),typeof(string)};var ctor=tb.DefineConstructor(MethodAttributes.Public,CallingConventions.Standard,sig);var il=ctor.GetILGenerator();il.Emit(OpCodes.Ldarg_0);il.Emit(OpCodes.Ldarg_1);il.Emit(OpCodes.Ldarg_2);il.Emit(OpCodes.Ldarg_3);il.Emit(OpCodes.Call,baseType.GetConstructor(sig));il.Emit(OpCodes.Ret);
  var b=typeof(Form).GetProperty("ShowWithoutActivation",F).GetGetMethod(true);var m=tb.DefineMethod(b.Name,MethodAttributes.Family|MethodAttributes.Virtual|MethodAttributes.HideBySig|MethodAttributes.SpecialName,typeof(bool),Type.EmptyTypes);il=m.GetILGenerator();il.Emit(OpCodes.Ldc_I4_1);il.Emit(OpCodes.Ret);tb.DefineMethodOverride(m,b);
  var cb=baseType.GetProperty("CreateParams",F).GetGetMethod(true);var cm=tb.DefineMethod(cb.Name,MethodAttributes.Family|MethodAttributes.Virtual|MethodAttributes.HideBySig|MethodAttributes.SpecialName,typeof(CreateParams),Type.EmptyTypes);il=cm.GetILGenerator();var v=il.DeclareLocal(typeof(CreateParams));il.Emit(OpCodes.Ldarg_0);il.Emit(OpCodes.Call,cb);il.Emit(OpCodes.Stloc,v);il.Emit(OpCodes.Ldloc,v);il.Emit(OpCodes.Ldloc,v);il.Emit(OpCodes.Callvirt,typeof(CreateParams).GetProperty("ExStyle").GetGetMethod());il.Emit(OpCodes.Ldc_I4,0x08000000);il.Emit(OpCodes.Or);il.Emit(OpCodes.Callvirt,typeof(CreateParams).GetProperty("ExStyle").GetSetMethod());il.Emit(OpCodes.Ldloc,v);il.Emit(OpCodes.Ret);tb.DefineMethodOverride(cm,cb);
  return (Form)Activator.CreateInstance(tb.CreateType(),new object[]{(int)model.ModelID,-1,"Normal"});
 }
 static void Show(int csid){Call(host,"ShowInterface",(int)model.ModelID,csid,false,"None",null,csid==33?1:0);dit=(Control)Get(host,"ActiveInterface");int section=csid==33?1:2;Call(dit,"BuildSection",section,false,false);((DevExpress.XtraTab.XtraTabControl)Get(dit,"XtraTabControlNewGIT")).SelectedTabPageIndex=section;Application.DoEvents();}
 static void Change(double number){dynamic e=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));e.ModelID=(int)model.ModelID;e.WSName="Funding Assumptions";e.CellAddress="G82";e.ChangedValue=number;e.DataFormat="N";e.Description="Private calculation matrix";dynamic r=model.ChangeManager.ProcessChange(e);Check((bool)r.BSuccess,"G82 change accepted");}
 static int Snapshot(string stage){var result=Call((object)model,"ReadCheckSheetValidation");var issues=(IList)Get(result,"Issues");Check(String.IsNullOrEmpty(Convert.ToString(Get(result,"ValidationError"))),"No validation read error at "+stage);
  var data=new StringBuilder();var range=book.DefinedNames.GetDefinedName("Outputs_CheckSheet").Range;var ws=range.Worksheet;
  for(int row=range.TopRowIndex;row<=range.BottomRowIndex;row++){string line="CHECKROW row="+(row+1)+" label="+ws.Cells[row,0].Value+" B="+Value(ws.Cells[row,1])+" D="+Value(ws.Cells[row,3])+" E="+Value(ws.Cells[row,4]);Console.WriteLine(line);data.AppendLine(line);}
  var tdb=book.Worksheets["Transactional DB"];foreach(int r in new[]{2677,2745})for(int c=16;c<=55;c++){var cell=tdb.Cells[r,c];string line="TDB row="+(r+1)+" column="+(c+1)+" value="+Value(cell);Console.WriteLine(line);data.AppendLine(line);if(r==2677&&traceColumn<0&&cell.Value.IsNumeric&&Math.Abs(cell.Value.NumericValue)>0.000001)traceColumn=c;}
  int mismatch=0;if(dit!=null)foreach(var grid in Children(dit).OfType<GridControl>().Where(g=>g.GetType().Name=="ReadOnlyMappedTableGrid")){var view=(GridView)grid.MainView;var rows=(IList)Get(grid,"sourceRows");for(int r=0;r<view.DataRowCount;r++)if(view.GetRowCellDisplayText(r,view.Columns["C4"])!=ws.Cells[(int)rows[r],4].DisplayText)mismatch++;}
  Console.WriteLine("STATE mode="+mode+" stage="+stage+" digest="+Hash(data.ToString())+" input="+Value(book.Worksheets["Funding Assumptions"].Cells["G82"])+" failures="+issues.Count+" rows="+String.Join(",",issues.Cast<object>().Select(i=>Get(i,"CheckRow")))+" warning="+Get((object)model,"CheckSheetWarningActive")+" heading="+(host==null?"none":Get(Get(host,"CheckSheetButton"),"Visibility"))+" events="+events+" mismatch="+mismatch+" engine="+book.Options.CalculationEngineType+" skip="+model.WBCalculationService.DontCalcTDBS);
  if(mode=="deferred1"&&traceColumn>=0){var cell=tdb.Cells[2677,traceColumn];Console.WriteLine("TRACE "+stage+" target="+cell.GetReferenceA1()+" value="+Value(cell)+" formula="+cell.FormulaInvariant);int count=0;foreach(CellRange precedent in cell.DirectPrecedents){Console.WriteLine("PRECEDENTRANGE "+precedent.Worksheet.Name+"!"+precedent.GetReferenceA1());for(int r=precedent.TopRowIndex;r<=precedent.BottomRowIndex&&count<24;r++)for(int c=precedent.LeftColumnIndex;c<=precedent.RightColumnIndex&&count<24;c++){var p=precedent.Worksheet.Cells[r,c];Console.WriteLine("PRECEDENT "+p.Worksheet.Name+"!"+p.GetReferenceA1()+" value="+Value(p)+" formula="+p.FormulaInvariant);count++;}}}
  return issues.Count;
 }
 static void Publish(){var result=Call((object)model,"ReadCheckSheetValidation");Call((object)model,"RecordIdleCheckSheetResult",Get((object)model,"CalculationRevision"),Get(result,"HasFailures"),true);}
 [STAThread] static int Main(string[] args){try{
  mode=args[4];Check(new[]{"deferred1","deferred2","chain","recursive","full"}.Contains(mode),"Known candidate "+mode);Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));Console.WriteLine("VERSION "+FileVersionInfo.GetVersionInfo(app.Location).FileVersion);Static("Abovo.AbovoAppCls","Initialise");Static("Abovo.FileManager","Initialise",new object[]{null});
  string copy=Path.Combine(args[1],"private-calc-matrix.xlsb");File.Copy(args[2],copy);var files=app.GetType("Abovo.FileManager");var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});Check(!(bool)loaded.BError,"Private workbook opened");model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)loaded.IntegerReturn);book=model.WB;Static("Abovo.CheckSheetWatch","Configure",false,5,2,false,false);
  model.GetType().GetEvent("CheckSheetStatusChanged",F).GetAddMethod(true).Invoke((object)model,new object[]{new EventHandler((s,e)=>{events++;Console.WriteLine("STATUS_EVENT "+events+" warning="+Get((object)model,"CheckSheetWarningActive"));})});Check(book.Worksheets["Funding Assumptions"].Cells["G82"].Value.NumericValue==2700,"Initial G82 equals2700");Snapshot("load");
  host=QuietHost(app.GetType("GroupInterfaceTemplate"));host.Opacity=0;host.ShowInTaskbar=false;host.ClientSize=new Size(1800,1100);host.Show();Application.DoEvents();Show(33);Change(3000);Show(0);Check(Snapshot("broken")>0,"Broken state reproduced");
  // Isolate notification plumbing without changing calculation. Normal navigation did not publish.
  Publish();Check((bool)Get((object)model,"CheckSheetWarningActive"),"Diagnostic publication shows red warning");dynamic undone=model.ChangeManager.Undo();Check((bool)undone.BSuccess,"Undo accepted");Check(Snapshot("identical-start")>0,"Stale state after Undo reproduced");
  var engine=book.Options.CalculationEngineType;var calcMode=book.Options.CalculationMode;bool skip=model.WBCalculationService.DontCalcTDBS;var timer=Stopwatch.StartNew();
  try{if(mode.StartsWith("deferred")){model.WBCalculationService.CalculateDeferredWorksheets(book);if(mode=="deferred2")model.WBCalculationService.CalculateDeferredWorksheets(book);}else{book.Options.CalculationEngineType=mode=="chain"?CalculationEngineType.ChainBased:CalculationEngineType.Recursive;model.WBCalculationService.DontCalcTDBS=false;if(mode=="full")book.CalculateFull();else book.Calculate();}}
  finally{book.Options.CalculationEngineType=engine;book.Options.CalculationMode=calcMode;model.WBCalculationService.DontCalcTDBS=skip;}
  Console.WriteLine("TIME candidate="+mode+" ms="+timer.ElapsedMilliseconds);Check(book.Options.CalculationEngineType==engine&&book.Options.CalculationMode==calcMode&&(bool)model.WBCalculationService.DontCalcTDBS==skip,"Engine mode and deferral restored");int remaining=Snapshot("after-candidate");Publish();Call(dit,"RefreshData",false);Snapshot("after-publish-refresh");Check((bool)Get((object)model,"CheckSheetWarningActive")== (remaining>0),"Published warning matches calculated balance");Check(book.Worksheets["Funding Assumptions"].Cells["G82"].Value.NumericValue==2700,"Source input restored");host.Dispose();model.ModelSpreadsheetControl.Dispose();Console.WriteLine("COMPLETE mode="+mode+" remaining="+remaining+" assertions="+checks+" original never saved");return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
