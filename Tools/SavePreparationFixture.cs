using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using System.Windows.Forms;
using DevExpress.Spreadsheet;

public static class SavePreparationFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Static;
 static Type guard;
 static void Check(bool value,string message){if(!value)throw new Exception(message);Console.WriteLine("PASS: "+message);}
 static string Formula(int n){return "=CONCATENATE("+String.Join(",",Enumerable.Repeat("A1",n))+")";}
 static void Save(IWorkbook w,Func<bool> action){guard.GetMethod("Save",F).Invoke(null,new object[]{w,action,null,true});}
 [STAThread] public static int Main(string[] args){try {
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));guard=app.GetType("Abovo.WorkbookXlsbFormulaCompatibility");
  Trace.Listeners.Add(new ConsoleTraceListener());
  Console.WriteLine("PROCESS BITNESS="+(IntPtr.Size*8));
  if(args.Length>2&&args[2]=="--state"){StateLifecycle(app,args[1]);return 0;}
  if(args.Length>4&&args[3]=="--save-stage"){SaveStageProbe(app,args);return 0;}
  if(args.Length>3&&args[3]=="--incremental"){IncrementalProbe(app,args);return 0;}
  if(args.Length>2){ModelSave(app,args);return 0;}
  Synthetic(args[1]);return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
 static void Synthetic(string output){
  using(var w=new Workbook()){
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;w.History.IsEnabled=true;var s=w.Worksheets[0];s.Cells["A1"].Value="text";
   int row=0;foreach(int count in Enumerable.Range(29,32))s.Cells[row++,1].FormulaInvariant=Formula(count);
   s.Cells[row++,1].FormulaInvariant="=IF(TRUE,"+Formula(37).Substring(1)+",\"literal,commas,(,)\")";
   s.Cells[row++,1].FormulaInvariant="=CONCATENATE(\""+new string(',',40)+"\",A1)";
   s.Cells["A3"].Value=true;s.Cells["A4"].Value=-1.25;
   s.Cells[row++,1].FormulaInvariant="=CONCATENATE(\"quoted,comma\",A2,A3,A4,"+String.Join(",",Enumerable.Repeat("A1",33))+")";
   s.Cells[row++,1].FormulaInvariant="=CONCATENATE(#N/A,"+String.Join(",",Enumerable.Repeat("A1",36))+")";
   s.Range["C2:C3"].ArrayFormulaInvariant=Formula(37);
   w.DefinedNames.Add("GlobalText",Formula(37));s.DefinedNames.Add("LocalText",Formula(37));
   w.Calculate();var expected=Enumerable.Range(0,row).Select(r=>s.Cells[r,1].Value).ToArray();
   s.Protect("test-fixture-only",WorksheetProtectionPermissions.Default);
   string path=Path.Combine(output,"guarded.xlsb");
   Save(w,()=>{w.Calculate();w.SaveDocument(path,DocumentFormat.Xlsb);return true;});
   Check(s.IsProtected&&w.Options.CalculationMode==WorkbookCalculationMode.Manual&&w.History.IsEnabled,"Protection, calculation mode and history preserved");
   using(var read=new Workbook()){
    read.Options.CalculationMode=WorkbookCalculationMode.Manual;read.LoadDocument(path);read.CalculateFullRebuild();
    for(int r=0;r<row;r++){var actual=read.Worksheets[0].Cells[r,1];if(actual.Value!=expected[r])Console.WriteLine("MISMATCH row="+r+" expected="+expected[r]+" actual="+actual.Value+" formula="+actual.FormulaInvariant);Check(actual.Value==expected[r]&&actual.FormulaInvariant!="=#VALUE!","Formula/value round trip row "+r);}
    Check(read.Worksheets[0].Cells["C2"].HasArrayFormula&&read.Worksheets[0].Cells["C2"].GetArrayFormulaRange().RowCount==2,"Legacy array geometry preserved");
    Check(read.DefinedNames.GetDefinedName("GlobalText").RefersTo!="=#VALUE!"&&read.Worksheets[0].DefinedNames.GetDefinedName("LocalText").RefersTo!="=#VALUE!","Global and local names preserved");
   }
  }
  using(var w=new Workbook()){
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;w.History.IsEnabled=true;var c=w.Worksheets[0].Cells["A1"];string original=Formula(37);c.FormulaInvariant=original;
   Save(w,()=>false);Check(c.FormulaInvariant==original,"Cancelled save restores original formula");
   bool failed=false;try{Save(w,()=>{throw new IOException("injected save failure");});}catch(TargetInvocationException){failed=true;}
   Check(failed&&c.FormulaInvariant==original&&w.History.IsEnabled,"Failed save restores formula and history");
   w.Worksheets[0].Cells["B1"].FormulaInvariant="=SUM("+String.Join(",",Enumerable.Repeat("1",31))+")";
   bool called=false;failed=false;try{Save(w,()=>{called=true;return true;});}catch(TargetInvocationException){failed=true;}
   Check(failed&&!called&&c.FormulaInvariant==original,"Unsupported long-argument call blocks save before any mutation");
   w.Worksheets[0].Cells["B1"].FormulaInvariant=Formula(255);called=false;failed=false;
   try{Save(w,()=>{called=true;return true;});}catch(TargetInvocationException){failed=true;}
   Check(failed&&!called&&c.FormulaInvariant==original,"Extreme CONCATENATE is rejected instead of silently damaged");
  }
 }
 static void ModelSave(Assembly app,string[] args){
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  string copy=Path.Combine(args[1],"input.xlsb");File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!loaded.BError,"Open private model copy");dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);IWorkbook w=(IWorkbook)model.WB;
  var trace=new SaveTrace();Trace.Listeners.Add(trace);
  string finalSavedPath=null;Dictionary<string,Dictionary<string,SavedCell>> finalSavedValues=null;
  using(var host=new Form())using(var ui=(Control)Activator.CreateInstance(app.GetType("FileInstanceInterface"),new object[]{0})){
  host.Opacity=0;host.ShowInTaskbar=false;host.Controls.Add(ui);ui.Dock=DockStyle.Fill;host.Show();
  var browser=(WebBrowser)ui.Controls.Find("WebBrowserBPInfo",true).Single();
  Call(ui,"PopulateFileInfo");AssertSummary(browser,copy);
  var previousAccess=(DateTime)model.PreviousFileAccessTime;
  Check(!(bool)model.IsDirty,"Opening the summary does not dirty the model");
  var clean=Call(model,"CalculateAndValidateForClose");Check(!(bool)CallProperty(clean,"HasFailures")&&trace.Modes.Count==0,"Unchanged loaded model closes without calculation or validation");
  var mode=w.Options.CalculationMode;var engine=w.Options.CalculationEngineType;var protection=w.Worksheets.ToDictionary(s=>s.Name,s=>s.IsProtected);bool skipDeferred=model.WBCalculationService.DontCalcTDBS;
  var originalDashboard=w.Worksheets["Multivariable Dashboard"].Cells["B41"].FormulaInvariant;
  Check((bool)model.NeedsFullRebuild&&!(bool)model.NeedsSaveRebuild,"Fresh model needs output verification, not a structural save rebuild");
  var firstStock=w.DefinedNames.GetDefinedName("CurrStNo").Range[0,0];
  int firstActive=mDummy(model,firstStock.Worksheet);
  dynamic firstChange=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));firstChange.ModelID=0;firstChange.WSName=firstStock.Worksheet.Name;firstChange.CellAddress=firstStock.GetReferenceA1();firstChange.ChangedValue=firstStock.Value.NumericValue+1;firstChange.DataFormat="I";firstChange.Description="First value-only save";
  dynamic firstResult=model.ChangeManager.ProcessChange(firstChange);model.WBCalcEngine.RemoveActiveObject(firstActive);
  Check(firstResult.BSuccess&&!(bool)model.NeedsSaveRebuild,"First stock edit does not request a structural save rebuild");
  string path=Path.Combine(args[1],"model-save.xlsb");
  Check((bool)model.SaveFileAsTo(path,true),"Actual model Save As succeeds");
  bool normalized=!trace.Preflights.Last().Contains("rewritten=0,");
  Check(trace.Modes.Last().Contains(normalized?"mode=rebuild,":"mode=deferred,"),"First value save defers initial rebuild unless formula compatibility repair requires it");AssertSummary(browser,path);
  Check(!(bool)model.NeedsFormulaPreflight&&trace.Preflights.Last().Contains("skipped=False"),"First save verifies the initially unknown formula structure");
  Check(((FileInfo)model.FileInfo).FullName==path&&((FileInfo)model.FileInfo).Length==new FileInfo(path).Length&&(DateTime)model.PreviousFileAccessTime==previousAccess,"Save As refreshes file metadata without changing previous-access history");
  string second=Path.Combine(args[1],"model-save-second.xlsb");Check((bool)model.SaveFileAsTo(second,true),"Second model Save As succeeds");
  Check(trace.Modes.Last().Contains(normalized?"mode=current,":"mode=deferred,"),"Unchanged Save As does not calculate pending results");AssertSummary(browser,second);
  Check(trace.Preflights.Last().Contains("skipped=True"),"Unchanged Save As skips the full formula scan");
  Check(w.Options.CalculationMode==mode&&w.Options.CalculationEngineType==engine&&w.Worksheets.All(s=>s.IsProtected==protection[s.Name])&&(bool)model.WBCalculationService.DontCalcTDBS==skipDeferred,"Model save preserves engine, mode, deferred-sheet policy and protection");
  using(var read=new Workbook()){
   read.Options.CalculationMode=WorkbookCalculationMode.Manual;Check(read.LoadDocument(path),"First saved workbook reloads for cache comparison");
   foreach(var key in new[]{new[]{"Multivariable Dashboard","B41"},new[]{"Cashflow detailed","BO9"},new[]{"Cashflow detailed","BR9"},new[]{"Detailed Comp Inc - Trad View","J73"},new[]{"Financial Position - Trad View","K26"}}){var c=read.Worksheets[key[0]].Cells[key[1]];Console.WriteLine("SAVED "+key[0]+"!"+key[1]+" = "+c.Value+" formula="+c.FormulaInvariant);}
   if(originalDashboard=="=#VALUE!")Console.WriteLine("LIMIT: Input dashboard formula was already erased; no source formula is invented.");
   else Check(read.Worksheets["Multivariable Dashboard"].Cells["B41"].FormulaInvariant!="=#VALUE!","Dashboard formula survives model save");
   foreach(string sheet in new[]{"Detailed Comp Inc - Trad View","Financial Position - Trad View","Cashflow detailed","Check Sheet"}){
    foreach(var c in w.Worksheets[sheet].GetUsedRange().ExistingCells.Where(c=>c.HasFormula)){var actual=read.Worksheets[sheet].Cells[c.RowIndex,c.ColumnIndex];CheckQuiet(Equivalent(c.Value,actual.Value),"Saved cache differs: "+sheet+"!"+c.GetReferenceA1()+" expected="+c.Value+" actual="+actual.Value+" formula="+c.FormulaInvariant);}
   }
   Check(true,"Output caches preserved through save/reopen");
  }
  model.EnsureDeferredSaveResultsCurrent("Testing first-save output gate");
  Check(!(bool)model.NeedsFullRebuild&&!(bool)model.NeedsSaveRebuild,"First-save pending output gate rebuilds before consuming results");
  var rent=w.DefinedNames.GetDefinedName("TransRents").Range[0,0];
  int active=mDummy(model,rent.Worksheet);
  Check(!rent.HasFormula&&!rent.Protection.Locked,"Selected rent input is an editable value cell");
  double beforeRent=rent.Value.IsNumeric?rent.Value.NumericValue:0;
  dynamic change=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));change.ModelID=0;change.WSName=rent.Worksheet.Name;change.CellAddress=rent.GetReferenceA1();change.ChangedValue=beforeRent+7.25;change.DataFormat="SM";change.Description="Save lifecycle regression";
  dynamic result=model.ChangeManager.ProcessChange(change);Check(result.BSuccess&&(bool)model.IsDirty&&!(bool)model.NeedsFullRebuild,"Typed rent edit marks pending results without rebuilding");
  Check(!(bool)model.NeedsFormulaPreflight,"Typed value edit preserves formula-check state");
  Check((bool)model.SaveFile()&&!(bool)model.IsDirty&&trace.Modes.Last().Contains("mode=deferred,"),"Ordinary edited Save defers full results without rebuilding");
  Check((bool)model.ResultsPending&&(bool)model.DeferredSaveResultsPending&&w.DocumentProperties.Custom["Abovo.Summit.ResultsPending"].BooleanValue,"Fast save persists a pending-results marker without marking inputs unsaved");
  AssertSummary(browser,"full results pending");
  Check((bool)model.EnsureDeferredSaveResultsCurrent("Testing output-reader gate"),"Output reader calculates deferred results");
  Check(!(bool)model.DeferredSaveResultsPending&&!(bool)model.ResultsPending,"Successful full calculation clears only the in-memory pending state");
  Check((bool)model.WBCalculationService.DontCalcTDBS==skipDeferred,"Ordinary save restores deferred-sheet policy");
  Check(trace.Preflights.Last().Contains("skipped=True"),"Ordinary edited save skips formula preflight");
  {
   var savedValues=CaptureOutputValues(w);
   try{w.Options.CalculationEngineType=CalculationEngineType.Recursive;model.WBCalculationService.DontCalcTDBS=false;w.CalculateFullRebuild();}
   finally{w.Options.CalculationEngineType=engine;model.WBCalculationService.DontCalcTDBS=skipDeferred;}
   int compared=VerifyOutputValues(w,savedValues,"Ordinary save differs from independent full rebuild");
   w.Options.CalculationEngineType=engine;Check(true,"Ordinary-save caches agree with independent full rebuild across "+compared+" output cells");
  }
  result=model.ChangeManager.Undo();Check(result.BSuccess&&rent.Value.NumericValue==beforeRent&&(bool)model.IsDirty&&!(bool)model.NeedsFullRebuild,"Undo value edit retains incremental save policy");
  result=model.ChangeManager.Redo();Check(result.BSuccess&&rent.Value.NumericValue==beforeRent+7.25&&(bool)model.IsDirty&&!(bool)model.NeedsFullRebuild,"Redo value edit retains incremental save policy");
  Check((bool)model.SaveFile()&&trace.Modes.Last().Contains("mode=deferred,"),"Save after undo/redo defers full results without rebuilding");
  Check(trace.Preflights.Last().Contains("skipped=True"),"Value Undo/Redo does not recheck formulas");
  model.WBCalcEngine.RemoveActiveObject(active);
  var stock=w.DefinedNames.GetDefinedName("CurrStNo").Range[0,0];active=mDummy(model,stock.Worksheet);
  double beforeStock=stock.Value.IsNumeric?stock.Value.NumericValue:0;
  dynamic stockData=model.WBData.GetISEDataStructure(0,0,1,0,0);
  dynamic calculatedColumn=null;dynamic calculatedPoint=null;
  foreach(dynamic column in stockData.DataColumns){if(column.ColumnTag.IsCalculated&&column.ColumnTag.DataType=="I"){calculatedColumn=column;calculatedPoint=stockData.DataRows[0].DataCells[column.Index];break;}}
  Check(calculatedPoint!=null,"Stock XML exposes an IsCalculated integer field");
  var calculatedCell=w.Worksheets[(string)calculatedPoint.SourceSheet].Cells[(string)calculatedPoint.SourceAddress];double beforeCalculated=calculatedCell.Value.NumericValue;
  change.WSName=stock.Worksheet.Name;change.CellAddress=stock.GetReferenceA1();change.ChangedValue=beforeStock+3;change.DataFormat="I";change.Description="Stock-value save regression";
  result=model.ChangeManager.ProcessChange(change);
  stockData.UpdateCalcs();
  Check(calculatedCell.Value.NumericValue==beforeCalculated+3&&stockData.DataRows[0].DataCells[calculatedColumn.Index].IntValue==(int)(beforeCalculated+3),"IsCalculated stock figure is calculated and returned to the grid data immediately, before Save");
  Check(result.BSuccess&&!(bool)model.NeedsFormulaPreflight&&!(bool)model.NeedsFullRebuild,"Active Stock Assumptions value edit leaves formula/structure flags clear");
  Check((bool)model.SaveFile()&&trace.Preflights.Last().Contains("skipped=True"),"Stock-figure save performs no formula scan");
  finalSavedPath=second;
  Check(trace.Modes.Last().Contains("mode=deferred,")&&(bool)model.DeferredSaveResultsPending,"Stock save uses fast path even with earlier pending results");
  int beforeClose=trace.Modes.Count;var fastClose=Call(model,"CalculateAndValidateForClose");
  Check(!(bool)CallProperty(fastClose,"HasFailures")&&trace.Modes.Count==beforeClose&&(bool)model.DeferredSaveResultsPending,"Saved healthy pending model closes without silently claiming calculated results");
  model.WBCalcEngine.CalculateDependencySensitiveFile("Fast-save analyser gate",false);
  Check(!(bool)model.ResultsPending,"Analyser dependency gate resolves saved pending results");
  finalSavedValues=CaptureOutputValues(w);
  try{w.Options.CalculationEngineType=CalculationEngineType.Recursive;model.WBCalculationService.DontCalcTDBS=false;w.CalculateFullRebuild();}
  finally{w.Options.CalculationEngineType=engine;model.WBCalculationService.DontCalcTDBS=skipDeferred;}
  Check(VerifyOutputValues(w,finalSavedValues,"Stock-save cache differs from independent full rebuild")>0,"Stock-save output caches agree with an independent full rebuild");
  model.WBCalcEngine.RemoveActiveObject(active);
  int calculations=trace.Modes.Count;clean=Call(model,"CalculateAndValidateForClose");Check(!(bool)CallProperty(clean,"HasFailures")&&trace.Modes.Count==calculations,"Saved healthy model closes without duplicate work");
  model.IsDirty=true;bool failed=false;
  try{model.GetType().GetMethod("SavePreparedWorkbook",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(model,new object[]{new Action(()=>{throw new IOException("injected model save failure");}),false});}catch(TargetInvocationException){failed=true;}
  Check(failed&&(bool)model.IsDirty&&w.Options.CalculationMode==mode&&w.Options.CalculationEngineType==engine&&w.Worksheets.All(s=>s.IsProtected==protection[s.Name]),"Model save failure retains dirty state and restores calculation/protection state");
  bool cancelled=(bool)model.GetType().GetMethod("SavePreparedWorkbook",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(model,new object[]{new Action(()=>{}),false});
  Check(!cancelled&&(bool)model.IsDirty,"No DocumentSaved event means cancelled save, never successful or clean");
  }
  model.CloseModel();Trace.Listeners.Remove(trace);model=null;w=null;
  // Large x86 models can exhaust native importer headroom if a second model is
  // loaded alongside the still-live, edited model. Verify persisted caches after
  // closing it instead; do not silently treat a false LoadDocument as success.
  GC.Collect();GC.WaitForPendingFinalizers();GC.Collect();
  files.GetMethod("Initialise").Invoke(null,new object[]{null});
  loaded=open.Invoke(null,new object[]{finalSavedPath,new FileInfo(finalSavedPath),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!loaded.BError,"Fast-saved workbook reopens through real Summit load path");
  dynamic reopened=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue((int)loaded.IntegerReturn);
  Check(!(bool)reopened.IsDirty&&!(bool)reopened.ResultsPending&&!(bool)reopened.DeferredSaveResultsPending,"Reopen refreshes pending results without inventing unsaved input changes");
  int reopenedCompared=VerifyOutputValues((IWorkbook)reopened.WB,finalSavedValues,"Reopened result differs from independent full calculation");
  Check(reopenedCompared>0,"Fast save/Summit reopen agrees across "+reopenedCompared+" output cells");
  reopened.CloseModel();
 }
 // CellValue text can depend on its workbook's shared-string storage. Keep
 // detached primitives so closing the source cannot invalidate the oracle.
 sealed class SavedCell {
  readonly bool numeric;readonly double number;readonly string text;
  public SavedCell(CellValue value){numeric=value.IsNumeric;if(numeric)number=value.NumericValue;else text=value.ToString(System.Globalization.CultureInfo.InvariantCulture);}
  public bool Matches(CellValue value){return numeric?value.IsNumeric&&Math.Abs(number-value.NumericValue)<=1e-7:!value.IsNumeric&&text==value.ToString(System.Globalization.CultureInfo.InvariantCulture);}
 }
 static Dictionary<string,Dictionary<string,SavedCell>> CaptureOutputValues(IWorkbook w){return new[]{"Detailed Comp Inc - Trad View","Financial Position - Trad View","Cashflow detailed","Check Sheet"}.ToDictionary(n=>n,n=>w.Worksheets[n].GetUsedRange().ExistingCells.ToDictionary(c=>c.GetReferenceA1(),c=>new SavedCell(c.Value)));}
 static int VerifyOutputValues(IWorkbook w,Dictionary<string,Dictionary<string,SavedCell>> expected,string context){int count=0;foreach(var sheet in expected)foreach(var cell in sheet.Value){CheckQuiet(cell.Value.Matches(w.Worksheets[sheet.Key].Cells[cell.Key].Value),context+": "+sheet.Key+"!"+cell.Key);count++;}return count;}
 static int mDummy(dynamic model,Worksheet sheet){int id=model.WBCalcEngine.AddActiveObject(new object());model.WBCalcEngine.AddActiveWorksheet(id,sheet,false);return id;}
 // Diagnostic alternatives only. No production calculation/save policy is changed.
 static void SaveStageProbe(Assembly app,string[] args){
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  string copy=Path.Combine(args[1],"probe-input.xlsb");File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!loaded.BError,"Save-stage probe opens a private copy");dynamic m=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);IWorkbook w=m.WB;
  // Normalize through the real guard and establish the same warm state for all candidates.
  guard.GetMethod("Save",F).Invoke(null,new object[]{w,new Func<bool>(()=>{Call(m,"EnsureSaveCalculationCurrent",new object[]{null});return true;}),new Action(()=>m.RequireFullRebuild()),true});
  var stock=w.DefinedNames.GetDefinedName("CurrStNo").Range[0,0];int active=mDummy(m,stock.Worksheet);
  dynamic change=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));change.ModelID=0;change.WSName=stock.Worksheet.Name;change.CellAddress=stock.GetReferenceA1();change.ChangedValue=stock.Value.NumericValue+3;change.DataFormat="I";change.Description="Private save-stage probe";
  dynamic result=m.ChangeManager.ProcessChange(change);Check(result.BSuccess,"Save-stage probe applies typed stock edit");
  string candidate=args[4],path=Path.Combine(args[1],candidate+".xlsb");var engine=w.Options.CalculationEngineType;var mode=w.Options.CalculationMode;bool skip=m.WBCalculationService.DontCalcTDBS;
  var customType=typeof(DevExpress.XtraSpreadsheet.Services.ICustomCalculationService);object service=w.GetService(customType);bool removed=false;
  Dictionary<string,Dictionary<string,SavedCell>> prepared=null;long prepareMs=0,exportMs=0,diskMs=0,restoreMs=0;
  try{
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;m.WBCalculationService.DontCalcTDBS=false;
   var timer=Stopwatch.StartNew();
   if(candidate.StartsWith("recursive"))w.Options.CalculationEngineType=CalculationEngineType.Recursive;
   if(candidate=="chain-no-service"){w.RemoveService(customType);removed=true;}
   w.CalculateFull();prepareMs=timer.ElapsedMilliseconds;
   prepared=CaptureProbeValues(w);
   timer.Restart();
   if(candidate=="recursive-buffer"){
    using(var buffer=new MemoryStream()){w.SaveDocument(buffer,DocumentFormat.Xlsb);exportMs=timer.ElapsedMilliseconds;timer.Restart();using(var output=File.Create(path))buffer.WriteTo(output);diskMs=timer.ElapsedMilliseconds;}
   }else{m.ModelSpreadsheetControl.SaveDocument(path,DocumentFormat.Xlsb);exportMs=timer.ElapsedMilliseconds;}
   timer.Restart();if(removed){w.AddService(customType,service);removed=false;}w.Options.CalculationEngineType=engine;restoreMs=timer.ElapsedMilliseconds;
  }finally{if(removed)w.AddService(customType,service);w.Options.CalculationEngineType=engine;w.Options.CalculationMode=mode;m.WBCalculationService.DontCalcTDBS=skip;}
  Console.WriteLine("SAVE_STAGE candidate="+candidate+", prepareMs="+prepareMs+", exportMs="+exportMs+", diskMs="+diskMs+", restoreMs="+restoreMs+", measuredTotalMs="+(prepareMs+exportMs+diskMs+restoreMs));
  w.Options.CalculationEngineType=CalculationEngineType.Recursive;m.WBCalculationService.DontCalcTDBS=false;w.CalculateFullRebuild();
  int oracleDiffs=CountProbeDifferences(w,prepared,"candidate versus rebuild");
  w.Options.CalculationEngineType=engine;m.WBCalculationService.DontCalcTDBS=skip;m.WBCalcEngine.RemoveActiveObject(active);m.CloseModel();m=null;w=null;
  GC.Collect();GC.WaitForPendingFinalizers();GC.Collect();
  using(var saved=new Workbook()){saved.Options.CalculationMode=WorkbookCalculationMode.Manual;Check(saved.LoadDocument(path),"Save-stage output reopens");int fileDiffs=CountProbeDifferences(saved,prepared,"file versus candidate");Console.WriteLine("SAVE_STAGE_RESULT candidate="+candidate+", oracleDifferences="+oracleDiffs+", persistenceDifferences="+fileDiffs);Check(oracleDiffs==0&&fileDiffs==0,"Candidate agrees with full rebuild and persisted caches");}
 }
 static Dictionary<string,Dictionary<string,SavedCell>> CaptureProbeValues(IWorkbook w){var values=CaptureOutputValues(w);values.Add("Transactional DB",w.Worksheets["Transactional DB"].GetUsedRange().ExistingCells.ToDictionary(c=>c.GetReferenceA1(),c=>new SavedCell(c.Value)));return values;}
 static int CountProbeDifferences(IWorkbook w,Dictionary<string,Dictionary<string,SavedCell>> expected,string stage){int differences=0;foreach(var sheet in expected)foreach(var cell in sheet.Value)if(!cell.Value.Matches(w.Worksheets[sheet.Key].Cells[cell.Key].Value)){differences++;if(differences<=6)Console.WriteLine("SAVE_STAGE_DIFFERENCE "+stage+": "+sheet.Key+"!"+cell.Key);}return differences;}
 static void AssertSummary(WebBrowser browser,string path){
  var wait=Stopwatch.StartNew();while(wait.ElapsedMilliseconds<5000){Application.DoEvents();if(browser.ReadyState==WebBrowserReadyState.Complete&&browser.Document!=null&&browser.Document.Body!=null&&browser.Document.Body.InnerText.Contains(path)){Check(true,"Native summary HTML shows "+Path.GetFileName(path));return;}System.Threading.Thread.Sleep(10);}
  throw new Exception("Summary HTML did not update to "+path);
 }
 static void CheckQuiet(bool value,string message){if(!value)throw new Exception(message);}
 static void IncrementalProbe(Assembly app,string[] args){
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  string copy=Path.Combine(args[1],"probe-input.xlsb");File.Copy(args[2],copy);
  var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!loaded.BError,"Probe opens a private model copy");dynamic m=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);IWorkbook w=m.WB;
  var engine=w.Options.CalculationEngineType;bool skip=m.WBCalculationService.DontCalcTDBS;int totalDiffs=0;
  Call(m,"EnsureSaveCalculationCurrent",new object[]{null});
  string[] sheets={"Detailed Comp Inc - Trad View","Financial Position - Trad View","Cashflow detailed","Check Sheet","Transactional DB"};
  for(int trial=0;trial<3;trial++){
   var cell=w.DefinedNames.GetDefinedName(trial==1?"TransRents":"CurrStNo").Range[0,0];
   Check(!cell.HasFormula&&!cell.Protection.Locked,"Probe input is an editable value: "+cell.Worksheet.Name+"!"+cell.GetReferenceA1());
   int id=m.WBCalcEngine.AddActiveObject(new object());m.WBCalcEngine.AddActiveWorksheet(id,cell.Worksheet,false);
   int second=-1;if(trial==2){second=m.WBCalcEngine.AddActiveObject(new object());m.WBCalcEngine.AddActiveWorksheet(second,w.Worksheets["Transactional DB"],false);}
   dynamic change=Activator.CreateInstance(app.GetType("Abovo.DataChangeEvent"));change.ModelID=0;change.WSName=cell.Worksheet.Name;change.CellAddress=cell.GetReferenceA1();change.ChangedValue=(cell.Value.IsNumeric?cell.Value.NumericValue:0)+(trial==1?7.25:3);change.DataFormat=trial==1?"SM":"I";change.Description="Incremental save calculation probe";
   dynamic result=m.ChangeManager.ProcessChange(change);Check(result.BSuccess,"Typed probe change applied with active worksheet calculation");
   w.Options.CalculationEngineType=CalculationEngineType.Recursive;m.WBCalculationService.DontCalcTDBS=false;
   try{
    var clock=Stopwatch.StartNew();w.Calculate();long incremental=clock.ElapsedMilliseconds;
    var cached=sheets.ToDictionary(n=>n,n=>w.Worksheets[n].GetUsedRange().ExistingCells.ToDictionary(c=>c.GetReferenceA1(),c=>c.Value));
    clock.Restart();w.CalculateFull();long fullMs=clock.ElapsedMilliseconds;
    var full=sheets.ToDictionary(n=>n,n=>w.Worksheets[n].GetUsedRange().ExistingCells.ToDictionary(c=>c.GetReferenceA1(),c=>c.Value));
    clock.Restart();w.CalculateFullRebuild();long rebuild=clock.ElapsedMilliseconds;int differences=0,fullDifferences=0;
    foreach(var sheet in cached)foreach(var c in sheet.Value){var rebuilt=w.Worksheets[sheet.Key].Cells[c.Key].Value;if(Equivalent(c.Value,rebuilt))continue;differences++;if(differences<=8)Console.WriteLine("PROBE DIFFERENCE "+sheet.Key+"!"+c.Key+" incremental="+c.Value+" rebuild="+rebuilt);}
    foreach(var sheet in full)foreach(var c in sheet.Value)if(!Equivalent(c.Value,w.Worksheets[sheet.Key].Cells[c.Key].Value))fullDifferences++;
    Console.WriteLine("INCREMENTAL_PROBE trial="+trial+", input="+cell.Worksheet.Name+", activeObjects="+(trial==2?2:1)+", incrementalMs="+incremental+", fullMs="+fullMs+", fullRebuildMs="+rebuild+", differences="+differences+", fullDifferences="+fullDifferences);totalDiffs+=differences;
    Check(fullDifferences==0,"Retained full-result calculation agrees with the full rebuild");
   }finally{w.Options.CalculationEngineType=engine;m.WBCalculationService.DontCalcTDBS=skip;m.WBCalcEngine.RemoveActiveObject(id);if(second>=0)m.WBCalcEngine.RemoveActiveObject(second);}
  }
  m.CloseModel();Check(totalDiffs==0,"Incremental candidate agrees with full rebuild in every probe");
 }
 class SaveTrace:TraceListener {
  public readonly System.Collections.Generic.List<string> Modes=new System.Collections.Generic.List<string>();
  public readonly System.Collections.Generic.List<string> Preflights=new System.Collections.Generic.List<string>();
  public override void Write(string text){}
  public override void WriteLine(string text){if(text==null)return;if(text.StartsWith("[Save Calculation Benchmark]"))Modes.Add(text);if(text.StartsWith("[XLSB Save Benchmark] formulaPreflight="))Preflights.Add(text);}
 }
 static object Call(object target,string name,params object[] args){return target.GetType().GetMethod(name,BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance).Invoke(target,args);}
 static void StateLifecycle(Assembly app,string output){
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  var mt=app.GetType("Abovo.FileManager+ExcelModel");var openMode=app.GetType("Abovo.FileManager+WorkbookOpenMode");
  var models=Array.CreateInstance(mt,2);files.GetField("ExcelModels").SetValue(null,models);
  dynamic m=Activator.CreateInstance(mt,new object[]{0,Enum.ToObject(openMode,0)});models.SetValue(m,0);
  dynamic other=Activator.CreateInstance(mt,new object[]{1,Enum.ToObject(openMode,0)});models.SetValue(other,1);
  IWorkbook w=(IWorkbook)m.WB;var s=w.Worksheets[0];s.Name="Inputs";s.Cells["A1"].Value=2;s.Cells["B1"].FormulaInvariant="=A1*2";
  var checks=w.Worksheets.Add("Check Sheet");checks.Cells["E1"].Value="OK";w.DefinedNames.Add("Outputs_CheckSheet","='Check Sheet'!$A$1:$H$1");
  m.IsDirty=false;Check((bool)m.NeedsFullRebuild&&!(bool)m.CloseValidationRequired,"Loaded clean model may close without calculating");
  var trace=new SaveTrace();Trace.Listeners.Add(trace);int serial=0;
  Action<string> save=expected=>{
   bool checkRequired=m.NeedsFormulaPreflight;
   int before=trace.Modes.Count;string path=Path.Combine(output,"state-"+(serial++)+".xlsb");
   bool saved=(bool)Call(m,"SavePreparedWorkbook",new Action(()=>m.ModelSpreadsheetControl.SaveDocument(path,DocumentFormat.Xlsb)),false);
   Check(saved&&trace.Modes.Count==before+1&&trace.Modes.Last().Contains("mode="+expected+","),"Save mode "+expected+"; actual="+trace.Modes.Last());
   Check(!(bool)m.NeedsSaveRebuild&&!(bool)m.IsDirty,"Successful save leaves inputs saved and no structural rebuild pending");
   Check(!(bool)m.NeedsFormulaPreflight&&trace.Preflights.Last().Contains("skipped="+(!checkRequired).ToString()),"Formula preflight runs only when required, then remembers verified state");
  };
  var cleanResult=Call(m,"CalculateAndValidateForClose");Check(!(bool)CallProperty(cleanResult,"HasFailures")&&trace.Modes.Count==0,"Clean close skips calculation and Check Sheet validation");
  Check((bool)m.NeedsFullRebuild&&!(bool)m.NeedsSaveRebuild,"Initial unverified graph is not a structural mutation");
  save("deferred");save("deferred");
  Check((bool)m.NeedsFullRebuild&&(bool)m.DeferredSaveResultsPending,"Repeated first saves preserve initial pending rebuild without executing it");
  Check((bool)m.EnsureDeferredSaveResultsCurrent("Initial output read")&&s.Cells["B1"].Value.NumericValue==4,"Initial output reader performs deferred full rebuild");
  save("current");
  Check((bool)other.NeedsFullRebuild,"One workbook cannot clear another workbook's rebuild flag");
  Check((bool)other.NeedsFormulaPreflight,"One workbook cannot certify another workbook's formula structure");
  s.Cells["A1"].Value=3;m.IsDirty=true;Check((bool)m.IsDirty&&!(bool)m.NeedsFullRebuild,"Ordinary service value edit marks pending calculation without requesting rebuild");
  save("deferred");Check((bool)m.ResultsPending&&(bool)m.DeferredSaveResultsPending,"Fast save separates saved inputs and pending results");
  Check(s.Cells["B1"].Value.NumericValue==4,"Save does not calculate a deferred result");
  Check((bool)m.EnsureDeferredSaveResultsCurrent("Fixture output")&&s.Cells["B1"].Value.NumericValue==6,"Output gate calculates deferred dependent result");
  Check(!(bool)m.EnsureDeferredSaveResultsCurrent("Fixture repeated output"),"Repeated output gate avoids duplicate calculation");
  Check(!w.DocumentProperties.Custom["Abovo.Summit.ResultsPending"].IsEmpty&&w.DocumentProperties.Custom["Abovo.Summit.ResultsPending"].BooleanValue,"In-memory calculation does not silently change the saved-file marker");
  save("current");Check(!w.DocumentProperties.Custom["Abovo.Summit.ResultsPending"].BooleanValue,"Next current save clears the persisted marker");
  using(var owner=new Form {Opacity=0,ShowInTaskbar=false,Text="Fixture results"}) {
   mt.GetField("ChangeManager").SetValue(m,Activator.CreateInstance(app.GetType("Abovo.ModelChangeManagerV2"),new object[]{0}));
   int refreshed=0;
   var binding=Activator.CreateInstance(app.GetType("Abovo.ModelFormHistoryBinding"),BindingFlags.Instance|BindingFlags.NonPublic,null,new object[]{owner,0,new Action(()=>refreshed++)},null);
   owner.Show();Application.DoEvents();int beforeRefresh=refreshed;
   s.Cells["A1"].Value=3.25;m.IsDirty=true;save("deferred");
   Call(binding,"RefreshWhenShown",owner,EventArgs.Empty);
   Check(!(bool)m.DeferredSaveResultsPending&&s.Cells["B1"].Value.NumericValue==6.5&&refreshed==beforeRefresh+1,"Retained standalone activation calculates pending results even without a history notification");
   Call(binding,"RefreshWhenShown",owner,EventArgs.Empty);
   Check(refreshed==beforeRefresh+1,"Unchanged standalone activation does not calculate or redraw again");
   ((IDisposable)binding).Dispose();
  }
  s.Cells["A1"].Value=3.5;m.IsDirty=true;save("deferred");
  var pendingSafety=app.GetType("Abovo.ModelSafetyManager");pendingSafety.GetMethod("BeginBulkWorkbookMutation").Invoke(null,new object[]{0});
  bool readerBlocked=false;try{m.EnsureDeferredSaveResultsCurrent("Fixture blocked reader");}catch(InvalidOperationException){readerBlocked=true;}
  Check(readerBlocked&&(bool)m.DeferredSaveResultsPending,"Blocked result calculation retains pending state");
  m.WBCalcEngine.CalcFile((byte)1,null);
  Check((bool)m.DeferredSaveResultsPending,"Internal structural calculation does not certify pending results mid-mutation");
  pendingSafety.GetMethod("EndBulkWorkbookMutation").Invoke(null,new object[]{0});save("rebuild");
  // Formula-capable native UI events are also raised by this opt-in API mode.
  m.ModelSpreadsheetControl.Options.Events.RaiseOnModificationsViaAPI=true;
  s.Cells["B1"].FormulaInvariant="=A1*3";
  m.ModelSpreadsheetControl.Options.Events.RaiseOnModificationsViaAPI=false;
  Check((bool)m.NeedsFullRebuild,"Native spreadsheet formula edit requests rebuild");save("rebuild");
  m.ModelSpreadsheetControl.Options.Events.RaiseOnModificationsViaAPI=true;
  var added=w.DefinedNames.Add("UnsafeFormula","=SUM("+String.Join(",",Enumerable.Repeat("1",31))+")");
  m.ModelSpreadsheetControl.Options.Events.RaiseOnModificationsViaAPI=false;
  // DefinedNameAdded is a UI notification, not raised by this API operation.
  // Exercise its retained callback explicitly; application API writers use
  // RequireFullRebuild/the structural service boundary before changing names.
  var nameHandler=(Delegate)m.GetType().GetField("NativeNameAddedHandler",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(m);
  nameHandler.DynamicInvoke((object)m.ModelSpreadsheetControl,null);
  Check((bool)m.NeedsFormulaPreflight,"Native name-added callback invalidates cached preflight");
  bool nameWritten=false,nameRejected=false;try{Call(m,"SavePreparedWorkbook",new Action(()=>nameWritten=true),false);}catch(TargetInvocationException){nameRejected=true;}
  Check(nameRejected&&!nameWritten&&(bool)m.NeedsFormulaPreflight,"Unsafe formula added after verification is blocked before saving");
  w.DefinedNames.Remove(added);m.RequireFullRebuild();save("rebuild");
  var safety=app.GetType("Abovo.ModelSafetyManager");safety.GetMethod("BeginBulkWorkbookMutation").Invoke(null,new object[]{0});
  Check((bool)m.NeedsFullRebuild,"Bulk structure command requests rebuild before mutation");
  Check((bool)m.CloseValidationRequired&&(bool)CallProperty(Call(m,"CalculateAndValidateForClose"),"HasFailures"),"An in-progress operation cannot take the clean-close shortcut");
  bool called=false,failed=false;try{Call(m,"SavePreparedWorkbook",new Action(()=>called=true),false);}catch(TargetInvocationException){failed=true;}
  Check(failed&&!called&&(bool)m.NeedsFullRebuild,"Save refuses an in-progress structural operation");
  safety.GetMethod("EndBulkWorkbookMutation").Invoke(null,new object[]{0});save("rebuild");
  // A trusted explicit full rebuild can satisfy the next save.
  m.RequireFullRebuild();var engine=w.Options.CalculationEngineType;w.Options.CalculationEngineType=CalculationEngineType.Recursive;
  m.WBCalcEngine.CalcFile((byte)3,null);w.Options.CalculationEngineType=engine;
  Check(!(bool)m.NeedsFullRebuild&&(bool)m.NeedsFormulaPreflight,"Rebuild completion does not clear the separate formula-preflight flag");save("current");
  s.Cells["A1"].Value=4;m.IsDirty=true;var closeResult=Call(m,"CalculateAndValidateForClose");
  Check(!(bool)CallProperty(closeResult,"HasFailures"),"Dirty close still validates workbook checks");save("current");
  checks.Cells["E1"].Value="Check";m.IsDirty=true;closeResult=Call(m,"CalculateAndValidateForClose");
  Check((bool)CallProperty(closeResult,"HasFailures"),"Failing dirty workbook is not accepted");save("current");
  Check((bool)m.CloseValidationRequired,"Known failing validation is not forgotten merely because a copy was saved");
  checks.Cells["E1"].Value="OK";m.IsDirty=true;closeResult=Call(m,"CalculateAndValidateForClose");Check(!(bool)CallProperty(closeResult,"HasFailures"),"Fixed checks can validate again");save("current");
  m.RecoverySaveAsRequired=true;Check((bool)m.CloseValidationRequired,"Recovery safeguard survives clean state");m.RecoverySaveAsRequired=false;
  m.IsDirty=true;bool cancelled=(bool)Call(m,"SavePreparedWorkbook",new Action(()=>{}),false);
  Check(!cancelled&&(bool)m.IsDirty&&(bool)m.NeedsFullRebuild,"Cancelled save stays dirty and conservatively invalidates preparation");save("rebuild");
  m.IsDirty=true;failed=false;try{Call(m,"SavePreparedWorkbook",new Action(()=>{throw new IOException("injected write failure");}),false);}catch(TargetInvocationException){failed=true;}
  Check(failed&&(bool)m.IsDirty&&(bool)m.NeedsFullRebuild,"Failed save retains dirty and rebuild flags");save("rebuild");
  // A later change must not be certified by an earlier completion token.
  long revision=(long)CallProperty(m,"CalculationRevision");m.RequireFullRebuild();Call(m,"MarkFullCalculationCurrent",revision,true);
  Check((bool)m.NeedsFullRebuild,"Stale completion cannot clear a newer mutation");
  m.CloseModel();other.CloseModel();Trace.Listeners.Remove(trace);
 }
 static object CallProperty(object target,string name){return target.GetType().GetProperty(name,BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance).GetValue(target,null);}
 static bool Equivalent(CellValue a,CellValue b){if(a.IsNumeric&&b.IsNumeric)return Math.Abs(a.NumericValue-b.NumericValue)<=1e-7;return a.ToString(System.Globalization.CultureInfo.InvariantCulture)==b.ToString(System.Globalization.CultureInfo.InvariantCulture);}
}
