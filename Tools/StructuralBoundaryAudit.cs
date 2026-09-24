using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using DevExpress.Spreadsheet;

// Diagnostic only. The runner copies its source and never saves the original.
// A completed audit may contain findings: this is not a blanket acceptance test.
public static class StructuralBoundaryAudit {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static object Field(object o,string n){return o.GetType().GetField(n,F).GetValue(o);}
 static readonly List<string> Findings=new List<string>();
 static void Observe(bool pass,string message){Console.WriteLine((pass?"PASS: ":"FINDING: ")+message);if(!pass)Findings.Add(message);}
 static Dictionary<string,string> Names(IWorkbook w){
  var d=w.DefinedNames.ToDictionary(n=>"global|"+n.Name,n=>n.RefersTo);
  foreach(var s in w.Worksheets)foreach(var n in s.DefinedNames)d.Add(s.Name+"|"+n.Name,n.RefersTo);return d;
 }
 static string CellText(Cell c){return c.HasFormula?"F|"+c.FormulaInvariant:"V|"+c.Value.Type+"|"+c.Value.ToString();}
 static Dictionary<string,string> Cells(IWorkbook w,IEnumerable<string> sheets){
  var d=new Dictionary<string,string>();foreach(string name in sheets.Distinct()){
   Console.WriteLine("CAPTURE_BEGIN\t"+name);int visited=0;
   foreach(var c in w.Worksheets[name].GetUsedRange().ExistingCells){
    if(c.HasFormula||!c.Value.IsEmpty)d[name+"!"+c.GetReferenceA1()]=CellText(c);
    if(++visited%50000==0)Console.WriteLine("CAPTURE_PROGRESS\t"+name+"\t"+visited);
   }Console.WriteLine("CAPTURE_END\t"+name+"\t"+visited);
  }return d;
 }
 static void Compare(Dictionary<string,string> before,Dictionary<string,string> after,string what){
  var keys=new HashSet<string>(before.Keys);keys.UnionWith(after.Keys);
  var diffs=keys.Where(k=>!before.ContainsKey(k)||!after.ContainsKey(k)||before[k]!=after[k]).ToArray();
  Observe(diffs.Length==0,what+": compared="+keys.Count+", differences="+diffs.Length);
  foreach(string k in diffs.Take(12))Console.WriteLine("DIFF\t"+k+"\tbefore="+(before.ContainsKey(k)?before[k]:"<empty>")+"\tafter="+(after.ContainsKey(k)?after[k]:"<empty>"));
 }
 static int Count(IWorkbook w,object r){var nr=w.DefinedNames.GetDefinedName((string)Field(r,"RecordCountNamedRange")).Range;return (Field(r,"Axis").ToString()=="Rows"?nr.RowCount:nr.ColumnCount)+(int)Field(r,"RecordCountAdjustment");}
 static string Geometry(IWorkbook w,object r){string range=(string)Field(r,"RecordCountNamedRange");var nr=w.DefinedNames.GetDefinedName(range);return nr==null?"MISSING":nr.RefersTo;}
 static void Mirrors(IWorkbook w,object sync,object structuralRule){
  var source=w.DefinedNames.GetDefinedName((string)Field(structuralRule,"TransactionDBSyncNamedRange"));
  if(source==null)return;int checkedCount=0;
  var covered=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
  foreach(object r in (IEnumerable)Field(sync,"SyncRules")){
   foreach(string name in (string[])Field(r,"TargetNamedRanges"))covered.Add(name);
   var s=w.DefinedNames.GetDefinedName((string)Field(r,"SourceNamedRange"));
   if(s==null||s.Range.Worksheet!=source.Range.Worksheet)continue;
   foreach(string name in (string[])Field(r,"TargetNamedRanges")){
    var target=w.DefinedNames.GetDefinedName(name);int expected=s.Range.ColumnCount+(int)Field(r,"SourceColumnAdjustment")+1;
    Observe(target!=null&&target.Range.RowCount==expected,"Mirror "+name+" rows="+(target==null?-1:target.Range.RowCount)+" expected="+expected);checkedCount++;
   }
  }
  foreach(var target in w.DefinedNames.Where(n=>n.Name.StartsWith("TransCopy_",StringComparison.OrdinalIgnoreCase)&&!covered.Contains(n.Name))){
   string name=target.Name.Substring(10);var s=w.DefinedNames.GetDefinedName(name);
   if(s==null&&System.Text.RegularExpressions.Regex.IsMatch(name,"_0[1-6]$"))s=w.DefinedNames.GetDefinedName(name.Substring(0,name.Length-3));
   if(s==null||!String.Equals(s.Name,source.Name,StringComparison.OrdinalIgnoreCase))continue;
   Observe(target.Range.RowCount==s.Range.RowCount+1,"Row mirror "+target.Name+" matches exact source "+s.Name);checkedCount++;
  }
  Console.WriteLine("MIRRORS_CHECKED\t"+checkedCount);
 }
 static void Inventory(IWorkbook w,object manager){
  foreach(DictionaryEntry e in (IDictionary)Field(manager,"Rules")){
   object r=e.Value;int n=Count(w,r);dynamic m=manager;
   Console.WriteLine("RULE\t"+e.Key+"\t"+Field(r,"Axis")+"\t"+Field(r,"RecordCountNamedRange")+"\t"+Geometry(w,r)+"\tcount="+n+"\tminimum="+Field(r,"MinimumRecordCount")+"\tprotected="+Field(r,"ProtectedLeadingRecordCount")+"\tadjustment="+Field(r,"RecordCountAdjustment")+"\tanchor="+Field(r,"InsertAnchorNamedRange")+"\toffset="+Field(r,"InsertIndexOffset"));
   int min=(int)Field(r,"MinimumRecordCount");
   dynamic rejected=m.ValidateDeleteLastRecords((string)e.Key,n-min+1);
   Console.WriteLine("BOUNDARY\t"+e.Key+"\tdeleteToOneBelowNativeMinimumRejected="+rejected.BError);
   Observe(rejected.BError,"Reject below minimum: "+e.Key);
   int leading=(int)Field(r,"ProtectedLeadingRecordCount");
   if(leading>0){
    var before=Names(w);dynamic selected=m.DeleteRecords((string)e.Key,new[]{0,leading-1});
    Observe(selected.BError,"Reject selected protected leading records: "+e.Key);
    Compare(before,Names(w),"Rejected leading-record deletion leaves names unchanged");
   }
   if(new[]{"OFA_RECORDS","REPAIRS_RECORDS","CAPEX_RECORDS","CAPGRANT_RECORDS","HOUSING_COMPONENT_RECORDS"}.Contains((string)e.Key)){
    var anchor=w.DefinedNames.GetDefinedName((string)Field(r,"InsertAnchorNamedRange")).Range;
    var start=w.DefinedNames.GetDefinedName((string)Field(r,"DeleteAnchorNamedRange")).Range;
    int vbaLast=anchor.LeftColumnIndex-1,limit=(string)e.Key=="REPAIRS_RECORDS"?9:6;
    Console.WriteLine("VBA_COLUMN_BOUNDARY\t"+e.Key+"\tvbaLastColumn1="+vbaLast+"\tnativeLastColumn1="+(start.LeftColumnIndex+n)+"\tvbaMaximumDelete="+(vbaLast-limit)+"\tnativeMaximumDelete="+(n-min)+"\tanchor="+anchor.GetReferenceA1());
    Observe(vbaLast==start.LeftColumnIndex+n&&vbaLast-limit==n-min,"VBA deletion contract: "+e.Key);
   }
  }
  foreach(string name in new[]{"LastOFACol","Rep_OFA_010","Rep_OFA_030","LastStockCol","StockCondCats","RepIncStkCat","Rep_CapExpend_010"}){
   var nr=w.DefinedNames.GetDefinedName(name).Range;Console.WriteLine("ANCHOR\t"+name+"\t"+nr.GetReferenceA1());
   for(int col=Math.Max(0,nr.RightColumnIndex-2);col<=nr.RightColumnIndex+1;col++){
    var ws=nr.Worksheet;Console.WriteLine("COLUMN\t"+ws.Name+"\tindex0="+col+"\twidth="+ws.Columns[col].Width+"\tvisible="+ws.Columns[col].Visible);
   }
  }
 }
 [STAThread] public static int Main(string[] args){try{
  CultureInfo.DefaultThreadCurrentCulture=CultureInfo.GetCultureInfo("en-GB");
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
  var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
  string copy=Path.Combine(args[2],"input-copy.xlsb");File.Copy(args[1],copy);
  var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
  if(loaded.BError)throw new Exception(loaded.StringReturn);
  dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);IWorkbook w=(IWorkbook)model.WB;object manager=model.WorkbookStructureRules;
  string id=args[3];int count=int.Parse(args[4]);bool boundary=id.StartsWith("boundary:");if(boundary)id=id.Substring(9);
  bool marker=id.StartsWith("marker:");if(marker)id=id.Substring(7);
  bool aligned=id.StartsWith("aligned:");if(aligned)id=id.Substring(8);
  bool minimumTrial=id.StartsWith("minimum:");if(minimumTrial)id=id.Substring(8);
  bool invalid=id.StartsWith("invalid:");if(invalid)id=id.Substring(8);
  if(id=="inventory"){Inventory(w,manager);Console.WriteLine("AUDIT_RESULT\tinventory\tfindings="+Findings.Count);return Findings.Count==0?0:2;}
  object rule=manager.GetType().GetMethod("GetRule",F).Invoke(manager,new object[]{id});if(rule==null)throw new Exception("Unknown rule "+id);
  var sheets=((IEnumerable)Field(rule,"Targets")).Cast<object>().Select(t=>(string)Field(t,"WorksheetName")).ToList();
  object linked=Field(rule,"LinkedColumns");if(linked!=null)sheets.AddRange(((IEnumerable)Field(linked,"Targets")).Cast<object>().Select(t=>(string)Field(t,"WorksheetName")));
  sheets.Add("Transactional DB");sheets=sheets.Distinct().ToList();
  var protection=w.Worksheets.ToDictionary(s=>s.Name,s=>s.IsProtected);var visible=w.Worksheets.ToDictionary(s=>s.Name,s=>s.Visible);
  var calc=w.Options.CalculationMode;var engine=w.Options.CalculationEngineType;
  Console.WriteLine("CAPTURE_NAMES_BEGIN");var names=Names(w);Console.WriteLine("CAPTURE_NAMES_END\t"+names.Count);
  var cells=Cells(w,sheets);int before=Count(w,rule);string geometry=Geometry(w,rule);
  if(aligned&&id=="REPAIRS_RECORDS"){
   var c=w.DefinedNames.GetDefinedName("StockCondCats").Range;
   string address=c.Worksheet.Range.FromLTRB(c.LeftColumnIndex,c.TopRowIndex-1,c.RightColumnIndex-1,c.TopRowIndex-1).GetReferenceA1();
   string repaired="='"+c.Worksheet.Name+"'!"+System.Text.RegularExpressions.Regex.Replace(address,@"([A-Z]+)([0-9]+)",m=>"$"+m.Groups[1].Value+"$"+m.Groups[2].Value);
   Console.WriteLine("APPROVED_NAME_REPAIR\tRepIncStkCat\told="+names["global|RepIncStkCat"]+"\texpectedAfterInverse="+repaired);
   names["global|RepIncStkCat"]=repaired;
  }
  var insertionAnchor=w.DefinedNames.GetDefinedName((string)Field(rule,"InsertAnchorNamedRange")).Range;
  int insertionColumn=insertionAnchor.LeftColumnIndex+(int)Field(rule,"InsertIndexOffset");
  Console.WriteLine("CASE\t"+id+"\tcount="+count+"\tbefore="+before+"\tgeometry="+geometry);
  if(invalid){
   if(id!="REPAIRS_RECORDS")throw new Exception("Unsupported invalid-geometry trial");
   var include=w.DefinedNames.GetDefinedName("RepIncStkCat");var range=include.Range;
   include.Range=range.Worksheet.Range.FromLTRB(range.LeftColumnIndex,range.TopRowIndex+1,range.RightColumnIndex,range.BottomRowIndex+1);
   var invalidNames=Names(w);dynamic rejected=model.WorkbookStructureRules.AddRecords(id,1);
   Observe(rejected.BError,"Reject unexpected Include-row geometry before insertion");
   Compare(invalidNames,Names(w),"Rejected malformed-range add leaves names unchanged");
   Compare(cells,Cells(w,sheets),"Rejected malformed-range add leaves formulas/constants unchanged");
   Observe(Count(w,rule)==before&&w.Options.CalculationMode==calc&&w.Options.CalculationEngineType==engine,"Rejected malformed-range add leaves count/calculation state unchanged");
   Observe(w.Worksheets.All(s=>s.IsProtected==protection[s.Name]&&s.Visible==visible[s.Name]),"Rejected malformed-range add preserves protection/visibility");
   Console.WriteLine("AUDIT_RESULT\tinvalid:"+id+"\tfindings="+Findings.Count);return Findings.Count==0?0:2;
  }
  if(minimumTrial){
   int min=(int)Field(rule,"MinimumRecordCount"),remove=before-min;
   dynamic v=model.WorkbookStructureRules.ValidateDeleteLastRecords(id,remove);
   Observe(!v.BError,"Allow deletion down to VBA minimum "+min);
   dynamic trial=model.WorkbookStructureRules.DeleteLastRecords(id,remove);
   Observe(!trial.BError&&Count(w,rule)==min,"Actual deletion retains minimum "+min+": "+trial.StringReturn);
   trial=model.WorkbookStructureRules.DeleteLastRecords(id,1);
   Observe(trial.BError&&Count(w,rule)==min,"Next deletion rejected without reducing minimum");
   Mirrors(w,(object)model.TransDBSync,rule);
   Observe(w.Options.CalculationMode==calc&&w.Options.CalculationEngineType==engine,"Minimum trial restores calculation state");
   Observe(w.Worksheets.All(s=>s.IsProtected==protection[s.Name]&&s.Visible==visible[s.Name]),"Minimum trial restores protection/visibility");
   Console.WriteLine("AUDIT_RESULT\tminimum:"+id+"\tfindings="+Findings.Count);return Findings.Count==0?0:2;
  }
  if(boundary){
   // Count specifies retained records. These cases deliberately try a forbidden
   // boundary on a private copy; never save/reuse the mutated result.
   int delete=before-count;dynamic validation=model.WorkbookStructureRules.ValidateDeleteLastRecords(id,delete);
   Console.WriteLine("VBA_BOUNDARY\t"+id+"\tretain="+count+"\tdelete="+delete+"\tnativeRejected="+validation.BError);
   dynamic trial=model.WorkbookStructureRules.DeleteLastRecords(id,delete);
   Observe(trial.BError,"VBA-forbidden deletion rejected: "+trial.StringReturn);
   Console.WriteLine("BOUNDARY_AFTER\t"+id+"\tcount="+Count(w,rule)+"\tgeometry="+Geometry(w,rule));
   Observe(w.Options.CalculationMode==calc&&w.Options.CalculationEngineType==engine,"Boundary attempt restores calculation state");
   Observe(w.Worksheets.All(s=>s.IsProtected==protection[s.Name]&&s.Visible==visible[s.Name]),"Boundary attempt restores protection/visibility");
   Console.WriteLine("AUDIT_RESULT\tboundary:"+id+"\tretain="+count+"\tfindings="+Findings.Count);return Findings.Count==0?0:2;
  }
  if(id=="OFA_RECORDS"||id=="REPAIRS_RECORDS"){
   // Only private comparison input is unprotected. Never obtain/persist passwords.
   var security=app.GetType("Abovo.WSSecurity");
   foreach(string s in sheets){if(w.Worksheets[s].IsProtected)security.GetMethod("UNProtectWS").Invoke(null,new object[]{0,s});w.Worksheets[s].Visible=true;}
   w.SaveDocument(Path.Combine(args[2],"excel-reference-input.xlsb"),DocumentFormat.Xlsb);
   foreach(string s in sheets){if(protection[s])security.GetMethod("ProtectWS").Invoke(null,new object[]{0,s,Type.Missing});w.Worksheets[s].Visible=visible[s];}
  }
  dynamic result=model.WorkbookStructureRules.AddRecords(id,count);Observe(!result.BError,"Add "+count+": "+result.StringReturn);
  if(result.BError)return 2;
  Observe(Count(w,rule)==before+count,"Record count grew: "+before+" -> "+Count(w,rule)+"; "+Geometry(w,rule));
  Mirrors(w,(object)model.TransDBSync,rule);
  if(aligned&&id=="REPAIRS_RECORDS"){
   var cats=w.DefinedNames.GetDefinedName("StockCondCats").Range;var inc=w.DefinedNames.GetDefinedName("RepIncStkCat").Range;
   Observe(inc.ColumnCount==cats.ColumnCount-1&&inc.RightColumnIndex==cats.RightColumnIndex-1,"Include covers all genuine Repairs categories and excludes template");
  }
  Observe(w.Options.CalculationMode==calc&&w.Options.CalculationEngineType==engine,"Add restores calculation state");
  Observe(w.Worksheets.All(s=>s.IsProtected==protection[s.Name]&&s.Visible==visible[s.Name]),"Add restores protection/visibility");
  if(marker){
   var ws=w.Worksheets[sheets[0]];
   var input=ws.Range.FromLTRB(insertionColumn,0,insertionColumn,ws.GetUsedRange().BottomRowIndex).ExistingCells.FirstOrDefault(c=>!c.Protection.Locked&&!c.HasFormula);
   if(input==null)throw new Exception("No editable constant in the first newly inserted column");
   input.Value="AUDIT_NEW_RECORD_MUST_BE_DELETED";
   Console.WriteLine("NEW_RECORD_MARKER\t"+ws.Name+"!"+input.GetReferenceA1());
  }
  if(id=="OFA_RECORDS"||id=="REPAIRS_RECORDS")Observe((bool)model.SaveFileAsTo(Path.Combine(args[2],"expanded.xlsb"),true),"Saved private expanded comparison through model");
  var afterAdd=Names(w);int n=Count(w,rule),minimum=(int)Field(rule,"MinimumRecordCount");
  result=model.WorkbookStructureRules.ValidateDeleteLastRecords(id,n-minimum+1);Observe(result.BError,"Reject below native minimum="+minimum);
  result=model.WorkbookStructureRules.DeleteRecords(id,new[]{n});Observe(result.BError,"Reject out-of-range selected index="+n);
  int leading=(int)Field(rule,"ProtectedLeadingRecordCount");
  if(leading>0){result=model.WorkbookStructureRules.DeleteRecords(id,new[]{0,leading-1});Observe(result.BError,"Reject selected protected leading records");}
  Compare(afterAdd,Names(w),"Rejected commands leave names unchanged");
  result=model.WorkbookStructureRules.DeleteLastRecords(id,count);Observe(!result.BError,"Delete-last "+count+": "+result.StringReturn);
  Observe(Count(w,rule)==before,"Logical record count restored");
  Mirrors(w,(object)model.TransDBSync,rule);
  Compare(names,Names(w),"All global/local names restored");Compare(cells,Cells(w,sheets),"Affected worksheet and TDB formulas/constants restored");
  Observe(w.Options.CalculationMode==calc&&w.Options.CalculationEngineType==engine,"Delete restores calculation state");
  Observe(w.Worksheets.All(s=>s.IsProtected==protection[s.Name]&&s.Visible==visible[s.Name]),"Delete restores protection/visibility");
  Console.WriteLine("AUDIT_RESULT\t"+id+"\tcount="+count+"\tfindings="+Findings.Count);return Findings.Count==0?0:2;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
