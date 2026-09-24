using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DevExpress.Spreadsheet;

// Fault injections and saves are confined to disposable copies made by this fixture.
public static class SnapshotCreationFixture {
 static Assembly app; static int assertions;
 const string S="TDB Snapshot", C="TDB Comparison", T="Transactional DB";
 static object Call(string type,string method,params object[] args){return app.GetType("Abovo."+type).GetMethod(method).Invoke(null,args);}
 static void Check(bool value,string text){if(!value)throw new Exception(text);assertions++;Console.WriteLine("PASS: "+text);}
 static void Capture(int id){Call("TransactionalDBSnapshotManager","CreateSnapshotAndComparison",id);}
 static string[] BusinessOrder(IWorkbook w){return w.Worksheets.Cast<Worksheet>().Where(s=>s.Name!=S&&s.Name!=C).Select(s=>s.Name).ToArray();}
 static void Order(IWorkbook w){Check(w.Worksheets[S].Index==w.Worksheets[T].Index+1&&w.Worksheets[C].Index==w.Worksheets[T].Index+2,"Snapshot and Comparison immediately follow Transactional DB");}
 static void Remove(IWorkbook w,string name){if(w.Worksheets.Contains(name))w.Worksheets.Remove(w.Worksheets[name]);}
 static void Reject(Action a,string what){try{a();}catch(TargetInvocationException e){if(e.InnerException is InvalidOperationException){Check(true,what);return;}throw;}throw new Exception("Expected rejection: "+what);}
 static string Content(IWorkbook w){
  using(var hash=SHA256.Create())using(var output=new CryptoStream(Stream.Null,hash,CryptoStreamMode.Write))using(var writer=new StreamWriter(output,Encoding.UTF8)){
   foreach(var sheet in w.Worksheets.Cast<Worksheet>().Where(s=>s.Name!=S&&s.Name!=C)){
    writer.WriteLine(sheet.Name);writer.WriteLine(sheet.IsProtected);
    foreach(var name in sheet.DefinedNames.Cast<DefinedName>().OrderBy(n=>n.Name)){writer.WriteLine(name.Name);writer.WriteLine(name.RefersTo);}
    foreach(var cell in sheet.GetExistingCells()){
     if(cell.Value.IsEmpty&&!cell.HasFormula)continue;
     writer.WriteLine(cell.GetReferenceA1());writer.WriteLine(cell.HasFormula?cell.FormulaInvariant:cell.Value.Type.ToString()+":"+cell.Value.ToString());
    }
   }
   foreach(var name in w.DefinedNames.Cast<DefinedName>().OrderBy(n=>n.Name)){writer.WriteLine(name.Name);writer.WriteLine(name.RefersTo);}
   writer.Flush();output.FlushFinalBlock();return BitConverter.ToString(hash.Hash);
  }
 }
 static double[][] Values(object doc){return ((IEnumerable)doc.GetType().GetProperty("Nodes").GetValue(doc,null)).Cast<object>().Select(n=>(double[])n.GetType().GetProperty("Values").GetValue(n,null)).ToArray();}
 [STAThread] public static int Main(string[] args){try{
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  Application.EnableVisualStyles();app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));Call("AbovoAppCls","Initialise");Call("FileManager","Initialise",new object[]{null});
  string copy=Path.Combine(args[1],"private-snapshot.xlsb");File.Copy(args[2],copy);
  int originalCount;using(var initial=new Workbook()){initial.Options.CalculationMode=WorkbookCalculationMode.Manual;initial.LoadDocument(copy);originalCount=initial.Worksheets.Count;}
  var open=app.GetType("Abovo.FileManager").GetMethod("OpenModel");dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  Check(!result.BError,"Private workbook opens");int id=result.IntegerReturn;
  dynamic model=((Array)app.GetType("Abovo.FileManager").GetField("ExcelModels").GetValue(null)).GetValue(id);IWorkbook w=model.WB;
  Check(w.Worksheets.Count==originalCount,"Ordinary open adds no worksheets");
  model.WBCalcEngine.CalculateDependencySensitiveFile("Snapshot fixture",true);
  Remove(w,S);Remove(w,C);int baseline=w.Worksheets.Count;string[] originalOrder=BusinessOrder(w);string business=Content(w);
  Check(!(bool)Call("TransactionalDBSnapshotManager","HasValidSnapshot",id)&&!(bool)Call("TransactionalDBSnapshotManager","HasPersistedSnapshot",id),"Missing-sheet probes report no snapshot");
  Call("TransactionalDBSnapshotManager","InvalidateSnapshot",id);Check(w.Worksheets.Count==baseline,"Probes and invalidation do not add sheets");
  string protection=Guid.NewGuid().ToString("N");w.Protect(protection,true,false);
  try{Reject(()=>Capture(id),"Protected workbook structure fails before adding sheets");Check(w.Worksheets.Count==baseline,"Protected structure retains original sheet count");}finally{w.Unprotect(protection);}
  var sourceName=w.DefinedNames.GetDefinedName("Transactional_Records")??w.Worksheets[T].DefinedNames.GetDefinedName("Transactional_Records");
  var sourceRange=sourceName.Range;sourceName.Range=w.Worksheets[T].Range["A1"];
  try{Reject(()=>Capture(id),"Invalid source headers reject before adding sheets");Check(w.Worksheets.Count==baseline,"Invalid source retains original sheet count");}finally{sourceName.Range=sourceRange;}
  // An existing unrelated destination must not cause its missing partner to be created.
  var other=w.Worksheets.Add(C);other.Cells["A1"].Value="Client-owned content";
  Reject(()=>Capture(id),"Collision rejects before provisioning or clearing");
  Check(!w.Worksheets.Contains(S)&&other.Cells["A1"].Value.TextValue=="Client-owned content","Collision retains existing content and absent partner");other.Cells["A1"].ClearContents();
  Capture(id);Order(w);Check(Object.ReferenceEquals(other,w.Worksheets[C])&&w.Worksheets.Count==baseline+2,"Missing Snapshot created; existing Comparison reused");
  var snapshot=w.Worksheets[S];Remove(w,C);Capture(id);Order(w);Check(Object.ReferenceEquals(snapshot,w.Worksheets[S])&&w.Worksheets.Count==baseline+2,"Missing Comparison created; existing Snapshot reused");
  Remove(w,S);Remove(w,C);Capture(id);Order(w);Check(w.Worksheets.Count==baseline+2,"Both missing destinations created by explicit capture");
  Check((bool)model.IsDirty&&(bool)Call("TransactionalDBSnapshotManager","HasValidSnapshot",id),"Successful explicit capture is valid and dirty for normal Save");
  snapshot=w.Worksheets[S];other=w.Worksheets[C];snapshot.Move(0);other.Move(w.Worksheets.Count-1);Capture(id);Order(w);
  Check(Object.ReferenceEquals(snapshot,w.Worksheets[S])&&Object.ReferenceEquals(other,w.Worksheets[C]),"Misplaced dedicated sheets are reused, not duplicated");
  Capture(id);Order(w);Check(w.Worksheets.Count==baseline+2,"Repeated capture remains idempotent");
  Check(originalOrder.SequenceEqual(BusinessOrder(w)),"All business worksheets retain relative order");
  Check(business==Content(w),"Business formulas, constants, names and entry protection unchanged");
  var spans=new System.Collections.Generic.HashSet<string>();
  foreach(var sheet in w.Worksheets.Cast<Worksheet>().Where(s=>s.Name!=S&&s.Name!=C))foreach(var cell in sheet.GetExistingCells()){
   if(!cell.HasFormula||!cell.FormulaInvariant.Contains(":"))continue;
   foreach(Match m in Regex.Matches(cell.FormulaInvariant,@"'((?:[^']|'')+):((?:[^']|'')+)'!|\b([A-Za-z_][A-Za-z_0-9]*):([A-Za-z_][A-Za-z_0-9]*)!")){
    string first=(m.Groups[1].Success?m.Groups[1].Value:m.Groups[3].Value).Replace("''","'");string last=(m.Groups[2].Success?m.Groups[2].Value:m.Groups[4].Value).Replace("''","'");
    if(!w.Worksheets.Contains(first)||!w.Worksheets.Contains(last))throw new Exception("Unresolved 3D boundary");
    int lo=Math.Min(w.Worksheets[first].Index,w.Worksheets[last].Index),hi=Math.Max(w.Worksheets[first].Index,w.Worksheets[last].Index);
    if((w.Worksheets[S].Index>=lo&&w.Worksheets[S].Index<=hi)||(w.Worksheets[C].Index>=lo&&w.Worksheets[C].Index<=hi))throw new Exception("Snapshot inside business 3D span "+first+":"+last);
    spans.Add(first+":"+last);
   }
  }
  Check(spans.Count>0,"Snapshot pair lies outside every recognised business 3D span ("+spans.Count+")");
  object live=Call("BalanceSheetStatement","Read",w), frozen=Call("BalanceSheetSnapshot","Read",w,live);var values=Values(frozen);
  object difference=Call("BalanceSheetStatement","Difference",live,frozen);
  Check(Values(difference).SelectMany(v=>v).All(v=>Double.IsNaN(v)||Math.Abs(v)<0.001),"New Balance Sheet snapshot equals live at every captured level");
  string saved=Path.Combine(args[1],"snapshot-roundtrip.xlsb");
  var save=model.GetType().GetMethod("SavePreparedWorkbook",BindingFlags.NonPublic|BindingFlags.Instance);
  Check((bool)save.Invoke(model,new object[]{new Action(()=>model.ModelSpreadsheetControl.SaveDocument(saved,DocumentFormat.Xlsb)),false}),"Snapshot saves through normal XLSB compatibility path");
  using(var reopened=new Workbook()){
   reopened.Options.CalculationMode=WorkbookCalculationMode.Manual;reopened.LoadDocument(saved);Order(reopened);
   Check(originalOrder.SequenceEqual(BusinessOrder(reopened)),"Saved/reopened business worksheet order preserved");
   object restored=Call("BalanceSheetSnapshot","Read",reopened,Call("BalanceSheetStatement","Read",reopened));
   Check(Values(restored).Zip(values,(a,b)=>a.SequenceEqual(b)).All(v=>v)&&Values(restored).Length==values.Length,"Every frozen Balance Sheet value survives XLSB save/reopen");
  }
  Console.WriteLine("SNAPSHOT="+saved);Console.WriteLine("PASS assertions="+assertions);return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}finally{Application.Exit();}}
}
