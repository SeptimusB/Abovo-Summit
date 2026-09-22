using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;

// Runs against a private input copy, in a fresh process. Never executes VBA.
public static class InsertionOptimisationFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static object Field(object o,string n){return o.GetType().GetField(n,F).GetValue(o);}
 static void Check(bool value,string message){if(!value)throw new Exception(message);Console.WriteLine("PASS: "+message);}
 static string Fingerprint(object captured,IWorkbook w){
  using(var stream=new MemoryStream())using(var writer=new BinaryWriter(stream,Encoding.UTF8,true)){
   int count=0;
   foreach(object item in (IEnumerable)Field(captured,"Items")){
    var sheet=(Worksheet)Field(item,"Sheet");var name=(DefinedName)Field(item,"Name");
    writer.Write(sheet==null?"":sheet.Name);writer.Write((int)Field(item,"Row"));writer.Write((int)Field(item,"Column"));
    writer.Write(name==null?"":name.Name);
    foreach(CellReferenceExpression r in (IEnumerable)Field(item,"References")){
     var text=new StringBuilder();r.BuildExpressionString(text,w);writer.Write(text.ToString());
    }
    count++;
   }
   writer.Flush();using(var sha=SHA256.Create())return count+":"+BitConverter.ToString(sha.ComputeHash(stream.ToArray())).Replace("-","");
  }
 }
 [STAThread] public static int Main(string[] args){try{
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(args.Length>6?args[6]:Path.Combine(args[0],"Abovo-summit.exe"));
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  var copy=Path.Combine(args[2],"input-copy.xlsb");File.Copy(args[1],copy);
  var open=files.GetMethod("OpenModel");var timer=Stopwatch.StartNew();
  dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!loaded.BError,"Opened private workbook in "+timer.ElapsedMilliseconds+" ms; bitness="+(IntPtr.Size*8));
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);IWorkbook w=(IWorkbook)model.WB;
  object manager=model.WorkbookStructureRules;string ruleID=args[3];int count=int.Parse(args[4]);string mode=args[5];
  var rule=manager.GetType().GetMethod("GetRule",F).Invoke(manager,new object[]{ruleID});
  if(mode=="recursive")rule.GetType().GetField("UseRecursiveEngineForMutation",F).SetValue(rule,true);
  var targets=((IEnumerable)Field(rule,"Targets")).Cast<object>().Select(t=>(string)Field(t,"WorksheetName")).ToArray();
  if(mode=="no-history")w.History.IsEnabled=false;
  if(mode=="automatic")w.Options.CalculationMode=WorkbookCalculationMode.Automatic;
  var protection=w.Worksheets.ToDictionary(s=>s.Name,s=>s.IsProtected);var visibility=w.Worksheets.ToDictionary(s=>s.Name,s=>s.Visible);
  var calc=w.Options.CalculationMode;var engine=w.Options.CalculationEngineType;bool history=w.History.IsEnabled;
  string range=(string)Field(rule,"RecordCountNamedRange");int before=w.DefinedNames.GetDefinedName(range).Range.ColumnCount;
  using(var listener=new TextWriterTraceListener(Path.Combine(args[2],"operation-trace.log"))){
   Trace.Listeners.Add(listener);
   try{
    if(mode=="tdb-bounded" || mode=="tdb-rows" || mode=="tdb-union"){
     var sync=(object)model.TransDBSync;
     var resize=sync.GetType().GetMethod("ResizeMirrorRangeInBatch",F);
     string prefix=ruleID=="FUNDING_RECORDS"?"TransCopy_FacilityNames_":ruleID=="DEVELOPMENT_IDENTIFIED_RECORDS"?"TransCopy_DevptSingle_":"TransCopy_DevptMulti_";
     var names=w.DefinedNames.Where(n=>n.Name.StartsWith(prefix) || (ruleID=="FUNDING_RECORDS" && n.Name.StartsWith("TransCopy_LoanDescsOrd_"))).OrderByDescending(n=>n.Range.TopRowIndex).Select(n=>n.Name).ToArray();
     var used=w.Worksheets["Transactional DB"].GetUsedRange();int left=mode=="tdb-rows"?-1:used.LeftColumnIndex,right=used.RightColumnIndex;
     w.Options.CalculationEngineType=CalculationEngineType.Recursive;w.Options.CalculationMode=WorkbookCalculationMode.Manual;w.History.IsEnabled=false;w.BeginUpdate();timer.Restart();
     try{
      if(mode=="tdb-union"){
       var ws=w.Worksheets["Transactional DB"];
       var blocks=names.Select(n=>w.DefinedNames.GetDefinedName(n).Range).Select(r=>ws.Range.FromLTRB(left,r.BottomRowIndex,right,r.BottomRowIndex+count-1)).ToArray();
       ws.InsertCells(ws.Range.Union(blocks),InsertCellsMode.ShiftCellsDown);
       Console.WriteLine("UNION shiftMs="+timer.ElapsedMilliseconds);
       foreach(string name in names){var r=w.DefinedNames.GetDefinedName(name).Range;int at=r.BottomRowIndex-count;ws.Range.FromLTRB(r.LeftColumnIndex,at,r.RightColumnIndex,r.BottomRowIndex-1).CopyFrom(ws.Range.FromLTRB(r.LeftColumnIndex,at-1,r.RightColumnIndex,at-1),PasteSpecial.All);}
      }else foreach(string name in names){int required=w.DefinedNames.GetDefinedName(name).Range.RowCount+count;dynamic result=resize.Invoke(sync,new object[]{w,name,required,left,right});Check(!result.BError,"Mirror "+name+": "+result.StringReturn);}
     }
     finally{w.EndUpdate();}
     Console.WriteLine("TDBPROBE mode="+mode+", mirrors="+names.Length+", ms="+timer.ElapsedMilliseconds);
     // Diagnostic only: deliberately unsynchronised source ranges. Never save.
     w.Options.CalculationEngineType=engine;w.Options.CalculationMode=calc;w.History.IsEnabled=history;
    }else if(mode=="scan"){
     var candidate=new Regex("(?:'(?:[^']|'')*:(?:[^']|'')*'|[A-Za-z_][\\w.]*:[A-Za-z_][\\w.]*)!",RegexOptions.Compiled);
     var screen=app.GetType("Abovo.WorkbookStructural3DReferences").GetMethod("MayContainThreeDReference",F);
     var fastScreen=screen==null?null:(Func<string,bool>)Delegate.CreateDelegate(typeof(Func<string,bool>),screen);
     long readTicks=0,matchTicks=0;int formulas=0,candidates=0,newCandidates=0;timer.Restart();
     foreach(var ws in w.Worksheets){
      long start=timer.ElapsedMilliseconds;int sheetCandidates=0;
      foreach(var cell in ws.GetUsedRange().ExistingCells){
       long tick=Stopwatch.GetTimestamp();bool formula=cell.HasFormula;string text=formula?cell.FormulaInvariant:null;readTicks+=Stopwatch.GetTimestamp()-tick;
       if(!formula)continue;formulas++;tick=Stopwatch.GetTimestamp();bool match=candidate.IsMatch(text);matchTicks+=Stopwatch.GetTimestamp()-tick;
       if(match){candidates++;sheetCandidates++;}
       if(fastScreen!=null){bool fast=fastScreen(text);if(fast)newCandidates++;CheckScreen(!match||fast,ws.Name,cell.GetReferenceA1());}
      }
      if(timer.ElapsedMilliseconds-start>500)Console.WriteLine("SCAN sheet="+ws.Name+", ms="+(timer.ElapsedMilliseconds-start)+", candidates="+sheetCandidates);
     }
     Console.WriteLine("SCAN total="+timer.ElapsedMilliseconds+", read="+(1000*readTicks/Stopwatch.Frequency)+", regex="+(1000*matchTicks/Stopwatch.Frequency)+", formulas="+formulas+", candidates="+candidates+", newCandidates="+newCandidates);
    }else if(mode=="capture"){
     var anchor=w.DefinedNames.GetDefinedName((string)Field(rule,"InsertAnchorNamedRange")).Range;
     int at=anchor.LeftColumnIndex+(int)Field(rule,"InsertIndexOffset");
     var repair=app.GetType("Abovo.WorkbookStructural3DReferences");
     for(int iteration=0;iteration<2;iteration++){
      timer.Restart();var captured=repair.GetMethod("Capture").Invoke(null,new object[]{w,targets,new[]{new KeyValuePair<int,int>(at,count)}});long elapsed=timer.ElapsedMilliseconds;
      Console.WriteLine("CAPTURE iteration="+iteration+", ms="+elapsed+", fingerprint="+Fingerprint(captured,w));
     }
    }else{
     Console.WriteLine("INSERT mode="+mode+", history="+history+", engine="+engine);
     // Exercise failure after source mutation, before TDB. This disposable
     // model must be recovery-only, never saved or silently retried.
     if(mode=="sync-failure")model.TransDBSync=null;
     timer.Restart();dynamic result=model.WorkbookStructureRules.AddRecords(ruleID,count);long elapsed=timer.ElapsedMilliseconds;
     if(mode=="sync-failure"){
      Check(result.BError&&((string)result.StringReturn).Contains("service is unavailable"),"Synchronisation failure reported");
      Check((bool)model.RecoverySaveAsRequired&&model.IntegrityState.ToString()=="RecoveryRequired","Incomplete operation marked recovery-required");
     }else Check(!result.BError,"Inserted "+count+" "+ruleID+" in "+elapsed+" ms: "+result.StringReturn);
     Check(w.DefinedNames.GetDefinedName(range).Range.ColumnCount==before+count,"Record range grew correctly");
     Check((bool)model.IsDirty,"Model marked dirty");
     Check(w.Options.CalculationMode==calc&&w.Options.CalculationEngineType==engine&&w.History.IsEnabled==history,"Insertion restored calculation/history state before save");
     Check(!((bool)app.GetType("Abovo.ModelSafetyManager").GetMethod("IsBulkWorkbookMutationInProgress").Invoke(null,new object[]{0})),"Bulk mutation guard released");
     Check(!(bool)Field(manager,"IsExecuting"),"Structural execution guard released");
     listener.Flush();
     if(mode!="sync-failure" && Convert.ToDecimal(app.GetType("Abovo.AbovoAppCls").GetProperty("DecVersionNumber").GetValue(null,null))>=2.60m){
      string trace;using(var read=new StreamReader(new FileStream(Path.Combine(args[2],"operation-trace.log"),FileMode.Open,FileAccess.Read,FileShare.ReadWrite)))trace=read.ReadToEnd();
      Check(trace.Split(new[]{"[TDB Sync Benchmark]"},StringSplitOptions.None).Length==2,"TDB synchronised exactly once");
      int sync=trace.IndexOf("stage=synchroniseTDB, engineShared=True, elapsed=");
      int restored=trace.IndexOf("stage=restoreEngine, elapsed=");
      int notify=trace.IndexOf("stage=postActions, state=started");
      if(ruleID=="FUNDING_RECORDS")Check(sync<0&&restored>=0&&notify>restored,"Funding retains its original engine/post-action sequence");
      else Check(sync>=0&&restored>sync&&notify>restored,"TDB then engine restoration then interface notifications");
      if(ruleID=="DEVELOPMENT_IDENTIFIED_RECORDS"&&count>1&&Convert.ToDecimal(app.GetType("Abovo.AbovoAppCls").GetProperty("DecVersionNumber").GetValue(null,null))>=2.61m){
       Check(Regex.Matches(trace,@"stage=capture3D, batch=\d+, elapsed=").Count==1&&Regex.Matches(trace,@"stage=advance3D, batch=2, elapsed=").Count==1,"One full capture then one incremental capture");
       Check(Regex.Matches(trace,@"stage=verify3D, elapsed=").Count==1,"One final tracked-reference verification");
      }
     }
     if(mode!="sync-failure")Check((bool)model.SaveFileAsTo(Path.Combine(args[2],"expanded.xlsb"),true),"Saved through production preparation/calculation/export path");
    }
   }finally{listener.Flush();Trace.Listeners.Remove(listener);}
  }
  Check(w.Worksheets.All(s=>s.IsProtected==protection[s.Name]&&s.Visible==visibility[s.Name]),"Protection and visibility restored");
  Check(w.Options.CalculationMode==calc&&w.Options.CalculationEngineType==engine&&w.History.IsEnabled==history,"Calculation engine, mode and history state restored");
  Console.WriteLine("OUTPUT="+args[2]);return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
 static void CheckScreen(bool valid,string sheet,string address){if(!valid)throw new Exception("Candidate omitted by new screen: "+sheet+"!"+address);}
}
