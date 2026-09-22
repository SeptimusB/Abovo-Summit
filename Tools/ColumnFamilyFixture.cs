using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using DevExpress.Spreadsheet;

public static class ColumnFamilyFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static object Field(object o,string n){return o.GetType().GetField(n,F).GetValue(o);}
 static void Check(bool v,string m){if(!v)throw new Exception(m);Console.WriteLine("PASS: "+m);}
 static Dictionary<string,string> Formulas(IWorkbook w,IEnumerable<string> sheets){
  var result=new Dictionary<string,string>();
  foreach(string s in sheets)foreach(var c in w.Worksheets[s].GetUsedRange().ExistingCells.Where(c=>c.HasFormula))result.Add(s+"!"+c.GetReferenceA1(),c.FormulaInvariant);
  return result;
 }
 [STAThread] public static int Main(string[] args){try{
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  var copy=Path.Combine(args[2],"input-copy.xlsb");File.Copy(args[1],copy);
  var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!loaded.BError,"Opened private workbook");
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);IWorkbook w=(IWorkbook)model.WB;
  object manager=model.WorkbookStructureRules;string ruleID=args[3];int count=int.Parse(args[4]);
  var rule=manager.GetType().GetMethod("GetRule",F).Invoke(manager,new object[]{ruleID});
  var targets=((IEnumerable)Field(rule,"Targets")).Cast<object>().Select(t=>(string)Field(t,"WorksheetName")).ToArray();
  var protection=w.Worksheets.ToDictionary(s=>s.Name,s=>s.IsProtected);var visibility=w.Worksheets.ToDictionary(s=>s.Name,s=>s.Visible);
  var names=w.DefinedNames.ToDictionary(n=>n.Name,n=>n.RefersTo);var formulas=Formulas(w,targets);
  var calc=w.Options.CalculationMode;var engine=w.Options.CalculationEngineType;
  string range=(string)Field(rule,"RecordCountNamedRange");int before=w.DefinedNames.GetDefinedName(range).Range.ColumnCount;
  // Reference input contains no credential. Unprotect only in private memory.
  var security=app.GetType("Abovo.WSSecurity");
  foreach(string s in targets){if(w.Worksheets[s].IsProtected)security.GetMethod("UNProtectWS").Invoke(null,new object[]{0,s});w.Worksheets[s].Visible=true;}
  w.SaveDocument(Path.Combine(args[2],"excel-reference-input.xlsb"),DocumentFormat.Xlsb);
  foreach(string s in targets){if(protection[s])security.GetMethod("ProtectWS").Invoke(null,new object[]{0,s,Type.Missing});w.Worksheets[s].Visible=visibility[s];}
  dynamic result;
  using(var trace=new StringWriter())using(var listener=new TextWriterTraceListener(trace)){
   Trace.Listeners.Add(listener);
   try{
    result=model.WorkbookStructureRules.AddRecords(ruleID,count);listener.Flush();
    string log=trace.ToString();File.WriteAllText(Path.Combine(args[2],"insert-trace.log"),log);
    foreach(string stage in new[]{"capture3D","unprotect","protect","shiftColumns","copyTemplate","copy3D","apply3D","endUpdate","restoreCalculationMode","postActions"})
     Check(log.Contains("stage="+stage),"Development trace stage: "+stage);
    Check(log.Contains("state=finished")&&log.Contains("outcome=ok"),"Development trace completes successfully");
    int batches=ruleID=="DEVELOPMENT_IDENTIFIED_RECORDS"&&count>1?2:1;
    Check(log.Split(new[]{Environment.NewLine},StringSplitOptions.None).Count(s=>s.Contains("stage=shiftColumns")&&s.Contains("elapsed="))==targets.Length*batches,"Development trace covers every sheet and batch");
   }finally{Trace.Listeners.Remove(listener);}
  }
  Check(!result.BError,"Added "+count+" "+ruleID+": "+result.StringReturn);
  Check(w.DefinedNames.GetDefinedName(range).Range.ColumnCount==before+count,"Record range grew correctly");
  string mirrorPrefix=ruleID=="DEVELOPMENT_IDENTIFIED_RECORDS"?"TransCopy_DevptSingle_":"TransCopy_DevptMulti_";
  var mirrors=w.DefinedNames.Where(n=>n.Name.StartsWith(mirrorPrefix)).ToArray();
  Check(mirrors.Length==14&&mirrors.All(n=>n.Range.RowCount==before+count),"All fourteen Development mirror dimensions");
  Check(w.Worksheets.All(s=>s.IsProtected==protection[s.Name]&&s.Visible==visibility[s.Name]),"Protection and visibility restored");
  Check(w.Options.CalculationMode==calc&&w.Options.CalculationEngineType==engine&&model.IsDirty,"Calculation state restored and model dirty");
  Check((bool)model.SaveFileAsTo(Path.Combine(args[2],"expanded.xlsb"),true),"Saved expanded result through model");
  int minimum=(int)Field(rule,"ProtectedLeadingRecordCount");
  foreach(int index in new[]{0,minimum-1,before+count-1}){
   result=model.WorkbookStructureRules.DeleteRecords(ruleID,new[]{index});
   Check(result.BError&&!model.IsDirty,"Protected leading/sentinel delete rejected: "+index);
  }
  result=model.WorkbookStructureRules.DeleteLastRecords(ruleID,count);Check(!result.BError,"Delete added records: "+result.StringReturn);
  var nameDiff=names.Where(n=>w.DefinedNames.GetDefinedName(n.Key)==null||w.DefinedNames.GetDefinedName(n.Key).RefersTo!=n.Value).ToArray();
  foreach(var n in nameDiff.Take(6))Console.WriteLine("NAME_DIFF "+n.Key+" before="+n.Value+" after="+w.DefinedNames.GetDefinedName(n.Key).RefersTo);
  Check(nameDiff.Length==0,"All "+names.Count+" names restored");
  var after=Formulas(w,targets);var diffs=formulas.Where(p=>!after.ContainsKey(p.Key)||after[p.Key]!=p.Value).ToArray();
  foreach(var p in diffs.Take(6))Console.WriteLine("FORMULA_DIFF "+p.Key+" before="+p.Value+" after="+(after.ContainsKey(p.Key)?after[p.Key]:"missing"));
  Check(diffs.Length==0&&after.Count==formulas.Count,"All "+formulas.Count+" linked-sheet formulas restored");
  Check(w.Worksheets.All(s=>s.IsProtected==protection[s.Name]&&s.Visible==visibility[s.Name]),"Delete restores worksheet state");
  Console.WriteLine("PASS: "+ruleID+" count="+count+". Original untouched.");return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
