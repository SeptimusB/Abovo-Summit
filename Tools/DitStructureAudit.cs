using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using DevExpress.Spreadsheet;

// Diagnostic only. Deliberately reproduces suspected defects in an unsaved
// private copy. It neither fixes nor saves any source workbook.
public static class DitStructureAudit {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static object Field(object o,string n){return o.GetType().GetField(n,F).GetValue(o);}
 [STAThread] public static int Main(string[] args){
  try {
   AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
   var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
   app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
   var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
   string copy=Path.Combine(args[2],"input-copy.xlsb");File.Copy(args[1],copy);
   var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
   if(loaded.BError)throw new Exception(loaded.StringReturn);
   dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);
   IWorkbook w=(IWorkbook)model.WB;object manager=model.WorkbookStructureRules;
   foreach(DictionaryEntry entry in (IDictionary)Field(manager,"Rules")) {
    object rule=entry.Value;string name=(string)Field(rule,"RecordCountNamedRange");var dn=w.DefinedNames.GetDefinedName(name);
    Console.WriteLine("GEOMETRY\t"+entry.Key+"\t"+Field(rule,"Axis")+"\t"+name+"\t"+(dn==null?"MISSING":dn.RefersTo)+"\t"+(dn==null?"":dn.Range.RowCount+"x"+dn.Range.ColumnCount));
   }
   var xml=new XmlDocument();xml.Load(Path.Combine(args[2],"Structure.xml"));
   foreach(XmlNode ds in xml.SelectNodes("//*[RowsExpandModel or StructureRuleID or RowExpandByNR]")) {
    Func<string,string> val=tag=>ds.SelectSingleNode(tag)==null?"":ds.SelectSingleNode(tag).InnerText;
    string token=val("RowExpandByNR"), declared=val("StructureRuleID");
    string resolved=(string)manager.GetType().GetMethod("ResolveRuleID").Invoke(manager,new object[]{declared.Length>0?declared:token});
    Console.WriteLine("XML_ROUTE\t"+val("ISDName")+"\t"+val("RowsExpandModel")+"\t"+token+"\t"+(resolved??"LEGACY"));
   }
   var sc=w.DefinedNames.GetDefinedName("Rep_ServChg_02");
   var scInput=w.DefinedNames.GetDefinedName("IR_ServChg_01");
   var scEnd=w.DefinedNames.GetDefinedName("LastUnitSCColumn");
   string scBefore=sc.RefersTo,scInputBefore=scInput.RefersTo,scEndBefore=scEnd.RefersTo;
   dynamic scResult=model.WorkbookStructureRules.AddRecords("SIMPLE_REP_SERVCHG_02",3);
   Console.WriteLine("SERVICE_CHARGE_INSERT\tbefore="+scBefore+"\tafter="+sc.RefersTo+"\terror="+scResult.BError);
   Console.WriteLine("SERVICE_CHARGE_LINKED_RANGE\tinputBefore="+scInputBefore+"\tinputAfter="+scInput.RefersTo+"\tworkingsEndBefore="+scEndBefore+"\tworkingsEndAfter="+scEnd.RefersTo);
   var specific=w.DefinedNames.GetDefinedName("IR_Spec_Inc_Ass1");
   string before=specific.RefersTo;int rows=specific.Range.RowCount,cols=specific.Range.ColumnCount;
   dynamic result=model.WorkbookStructureRules.AddRecords("SIMPLE_IR_SPEC_INC_ASS1",3);
   Console.WriteLine("SPECIFIC_INSERT\tbefore="+before+"\tafter="+specific.RefersTo+"\terror="+result.BError);
   Console.WriteLine("SPECIFIC_AXIS_DEFECT_REPRODUCED\t"+(!result.BError && specific.Range.RowCount==rows && specific.Range.ColumnCount==cols+3));
   var cash=w.DefinedNames.GetDefinedName("IR_Cash_Journals").Range;var ws=cash.Worksheet;
   int last=cash.BottomRowIndex,col=cash.LeftColumnIndex;bool wasProtected=ws.IsProtected;
   var security=app.GetType("Abovo.WSSecurity");if(wasProtected)security.GetMethod("UNProtectWS").Invoke(null,new object[]{0,ws.Name});
   if(ws.Cells[last,col].Protection.Locked)throw new Exception("Cash journal description test cell is locked; fixture must not alter it.");
   ws.Cells[last,col].Value="AUDIT_ONLY_NOT_A_NEW_TRANSACTION";
   if(wasProtected)security.GetMethod("ProtectWS").Invoke(null,new object[]{0,ws.Name,Type.Missing});
   result=model.WorkbookStructureRules.AddRecords("SIMPLE_IR_CASH_JOURNALS",3);
   int duplicates=Enumerable.Range(last+1,3).Count(r=>ws.Cells[r,col].Value.TextValue=="AUDIT_ONLY_NOT_A_NEW_TRANSACTION");
   Console.WriteLine("CASH_INPUT_DUPLICATION\tnewRowsWithCopiedInput="+duplicates+"\terror="+result.BError);
   Console.WriteLine("Audit finished on an unsaved private workbook. These are defect probes, not passing integration acceptance tests.");
   return 0;
  } catch(Exception e){Console.Error.WriteLine(e);return 1;}
 }
}
