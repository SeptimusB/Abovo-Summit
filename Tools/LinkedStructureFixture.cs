using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DevExpress.Spreadsheet;

public static class LinkedStructureFixture {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static object Field(object o,string n){return o.GetType().GetField(n,F).GetValue(o);}
 static void Check(bool v,string m){if(!v)throw new Exception(m);Console.WriteLine("PASS: "+m);}
 static Dictionary<string,string> Formulas(IWorkbook w,IEnumerable<string> sheets){
  var result=new Dictionary<string,string>();
  foreach(string s in sheets)foreach(var c in w.Worksheets[s].GetUsedRange().ExistingCells.Where(c=>c.HasFormula))result.Add(s+"!"+c.GetReferenceA1(),c.FormulaInvariant);
  return result;
 }
 static void Mirrors(IWorkbook w,object sync,string sourceSheet,string rowRange,bool rows){
  int checkedCount=0;
  foreach(object rule in (IEnumerable)Field(sync,"SyncRules")){
   var source=w.DefinedNames.GetDefinedName((string)Field(rule,"SourceNamedRange"));
   if(source==null||source.Range.Worksheet.Name!=sourceSheet)continue;
   int expected=source.Range.ColumnCount+(int)Field(rule,"SourceColumnAdjustment")+1;
   foreach(string name in (string[])Field(rule,"TargetNamedRanges")){
    var target=w.DefinedNames.GetDefinedName(name);
    if(target==null||target.Range.RowCount!=expected)throw new Exception("Mirror size mismatch: "+name);checkedCount++;
   }
  }
  if(rows){int expected=w.DefinedNames.GetDefinedName(rowRange).Range.RowCount+1;
   foreach(var n in w.DefinedNames.Where(n=>n.Name=="TransCopy_"+rowRange||n.Name.StartsWith("TransCopy_"+rowRange+"_"))){
    if(n.Range.RowCount!=expected)throw new Exception("Row mirror size mismatch: "+n.Name);checkedCount++;
   }
  }
  Console.WriteLine("PASS: "+checkedCount+" associated TDB mirror dimensions");
 }
 [STAThread] public static int Main(string[] args){try{
  System.Diagnostics.Trace.Listeners.Add(new System.Diagnostics.ConsoleTraceListener());
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
  var targets=((IEnumerable)Field(rule,"Targets")).Cast<object>().Select(t=>(string)Field(t,"WorksheetName")).ToList();
  var linked=Field(rule,"LinkedColumns");string linkedAnchor=linked==null?null:(string)Field(linked,"EndAnchorNamedRange");
  if(linked!=null)targets.AddRange(((IEnumerable)Field(linked,"Targets")).Cast<object>().Select(t=>(string)Field(t,"WorksheetName")));
  bool rows=Field(rule,"Axis").ToString()=="Rows";
  string range=(string)Field(rule,"RecordCountNamedRange");var nr=w.DefinedNames.GetDefinedName(range).Range;
  int before=rows?nr.RowCount:nr.ColumnCount,other=rows?nr.ColumnCount:nr.RowCount;
  int endBefore=linked==null?0:w.DefinedNames.GetDefinedName(linkedAnchor).Range.LeftColumnIndex;
  var protection=w.Worksheets.ToDictionary(s=>s.Name,s=>s.IsProtected);var visibility=w.Worksheets.ToDictionary(s=>s.Name,s=>s.Visible);
  var names=w.DefinedNames.ToDictionary(n=>n.Name,n=>n.RefersTo);var formulas=Formulas(w,targets);
  var localNames=w.Worksheets.SelectMany(s=>s.DefinedNames.Select(n=>new {Sheet=s.Name,Name=n.Name,Formula=n.RefersTo})).ToList();
  var calc=w.Options.CalculationMode;var engine=w.Options.CalculationEngineType;
  var security=app.GetType("Abovo.WSSecurity");
  // Deliberately populate the copied row/column with a distinctive existing input.
  Cell sentinel=null;
  if(rows){int row=nr.BottomRowIndex-((bool)Field(rule,"InsertAtAnchorEnd")?1:0);sentinel=nr.Worksheet.Range.FromLTRB(nr.LeftColumnIndex,row,nr.Worksheet.GetUsedRange().RightColumnIndex,row).ExistingCells.FirstOrDefault(c=>!c.Protection.Locked&&!c.HasFormula);}
  else {var a=w.DefinedNames.GetDefinedName((string)Field(rule,"InsertAnchorNamedRange")).Range;int col=a.LeftColumnIndex+(int)Field(rule,"InsertIndexOffset");sentinel=a.Worksheet.Range.FromLTRB(col,nr.TopRowIndex,col,a.Worksheet.GetUsedRange().BottomRowIndex).ExistingCells.FirstOrDefault(c=>!c.Protection.Locked&&!c.HasFormula);}
  int sentinelRow=-1,sentinelCol=-1;
  if(sentinel!=null){sentinelRow=sentinel.RowIndex;sentinelCol=sentinel.ColumnIndex;if(sentinel.Worksheet.IsProtected)security.GetMethod("UNProtectWS").Invoke(null,new object[]{0,sentinel.Worksheet.Name});sentinel.Value="EXISTING_INPUT_MUST_NOT_DUPLICATE";}
  foreach(string s in targets){if(w.Worksheets[s].IsProtected)security.GetMethod("UNProtectWS").Invoke(null,new object[]{0,s});w.Worksheets[s].Visible=true;}
  w.SaveDocument(Path.Combine(args[2],"excel-reference-input.xlsb"),DocumentFormat.Xlsb);
  foreach(string s in targets){if(protection[s])security.GetMethod("ProtectWS").Invoke(null,new object[]{0,s,Type.Missing});w.Worksheets[s].Visible=visibility[s];}
  dynamic result=model.WorkbookStructureRules.AddRecords(ruleID,count);
  Check(!result.BError,"Added "+count+" "+ruleID+": "+result.StringReturn);
  nr=w.DefinedNames.GetDefinedName(range).Range;
  Check((rows?nr.RowCount:nr.ColumnCount)==before+count&&(rows?nr.ColumnCount:nr.RowCount)==other,"Only intended axis grew");
  if(linked!=null)Check(w.DefinedNames.GetDefinedName(linkedAnchor).Range.LeftColumnIndex==endBefore+count,"Linked workings grew with input records");
  Mirrors(w,(object)model.TransDBSync,nr.Worksheet.Name,range,rows);
  if(sentinelRow>=0){
   var ws=nr.Worksheet;int matches=ws.GetUsedRange().ExistingCells.Count(c=>!c.HasFormula&&c.Value.TextValue=="EXISTING_INPUT_MUST_NOT_DUPLICATE");
   Check(matches==1,"Existing input retained exactly once, not copied into new records");
  }
  Check(w.Worksheets.All(s=>s.IsProtected==protection[s.Name]&&s.Visible==visibility[s.Name]),"Protection and visibility restored");
  Check(w.Options.CalculationMode==calc&&w.Options.CalculationEngineType==engine&&model.IsDirty,"Calculation state restored and model dirty");
  Check((bool)model.SaveFileAsTo(Path.Combine(args[2],"expanded.xlsb"),true),"Saved expanded result through model");
  result=model.WorkbookStructureRules.DeleteRecords(ruleID,new[]{before+count});
  Check(result.BError&&!model.IsDirty,"Out-of-range delete rejected without dirtying");
  int minimum=(int)Field(rule,"MinimumRecordCount");
  result=model.WorkbookStructureRules.ValidateDeleteLastRecords(ruleID,before+count-minimum+1);
  Check(result.BError&&!model.IsDirty,"Below-minimum delete rejected without mutation");
  result=model.WorkbookStructureRules.DeleteLastRecords(ruleID,count);Check(!result.BError,"Delete added records: "+result.StringReturn);
  Mirrors(w,(object)model.TransDBSync,nr.Worksheet.Name,range,rows);
  var nameDiff=names.Where(n=>w.DefinedNames.GetDefinedName(n.Key)==null||w.DefinedNames.GetDefinedName(n.Key).RefersTo!=n.Value).ToArray();
  foreach(var n in nameDiff.Take(8))Console.WriteLine("NAME_DIFF "+n.Key+" before="+n.Value+" after="+w.DefinedNames.GetDefinedName(n.Key).RefersTo);
  Check(nameDiff.Length==0,"All "+names.Count+" names restored");
  Check(localNames.All(n=>w.Worksheets[n.Sheet].DefinedNames.GetDefinedName(n.Name)!=null&&w.Worksheets[n.Sheet].DefinedNames.GetDefinedName(n.Name).RefersTo==n.Formula),"All "+localNames.Count+" worksheet-local names restored");
  var after=Formulas(w,targets);var diffs=formulas.Where(p=>!after.ContainsKey(p.Key)||after[p.Key]!=p.Value).ToArray();
  foreach(var p in diffs.Take(8))Console.WriteLine("FORMULA_DIFF "+p.Key+" before="+p.Value+" after="+(after.ContainsKey(p.Key)?after[p.Key]:"missing"));
  Check(diffs.Length==0&&after.Count==formulas.Count,"All "+formulas.Count+" linked-sheet formulas restored");
  Check(w.Worksheets.All(s=>s.IsProtected==protection[s.Name]&&s.Visible==visibility[s.Name]),"Delete restores worksheet state");
  Console.WriteLine("PASS: "+ruleID+" count="+count+". Original untouched.");return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
