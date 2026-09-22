using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DevExpress.Spreadsheet;
public static class StructuralOrderAudit {
 const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
 static object Field(object o,string n){return o.GetType().GetField(n,F).GetValue(o);}
 [STAThread] public static int Main(string[] args){
  try{
   AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
   var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
   app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
   var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
   var privateCopy=Path.Combine(args[2],"audit-input.xlsb");File.Copy(args[1],privateCopy);
   var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{privateCopy,new FileInfo(privateCopy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
   if(loaded.BError)throw new Exception(loaded.StringReturn);
   dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);
   // Existing workbook credential stays in memory only. Never emit it.
   string credential=(string)model.WBStructure.RejData;
   object manager=model.WorkbookStructureRules;
   var rules=(IDictionary)Field(manager,"Rules");
   var insert=manager.GetType().GetMethod("InsertColumnsForTarget",F);
   foreach(DictionaryEntry entry in rules){
    object rule=entry.Value;var targets=((IEnumerable)Field(rule,"Targets")).Cast<object>().ToArray();
    Console.WriteLine("RULE\t"+entry.Key+"\t"+Field(rule,"Axis")+"\t"+Field(rule,"InsertAnchorNamedRange")+"\t"+Field(rule,"InsertIndexOffset")+"\t"+Field(rule,"InsertAllColumnsBeforeCopy")+"\t"+String.Join(";",targets.Select(t=>Field(t,"WorksheetName")+":"+Field(t,"TemplateOffset")+":"+Field(t,"CopyMode"))));
    if(Convert.ToString(Field(rule,"Axis"))!="Columns"||targets.Length<2)continue;
    using(var sequential=new Workbook())using(var grouped=new Workbook()){
     sequential.Options.CalculationMode=grouped.Options.CalculationMode=WorkbookCalculationMode.Manual;
     sequential.LoadDocument(privateCopy);grouped.LoadDocument(privateCopy);
     int at=sequential.DefinedNames.GetDefinedName((string)Field(rule,"InsertAnchorNamedRange")).Range.LeftColumnIndex+(int)Field(rule,"InsertIndexOffset");
     var actions=new List<Action>();
     foreach(var target in targets){
      string name=(string)Field(target,"WorksheetName");
      foreach(var w in new IWorkbook[]{sequential,grouped}){var ws=w.Worksheets[name];if(ws.IsProtected)ws.Unprotect(credential);ws.Visible=true;}
      insert.Invoke(manager,new object[]{sequential.Worksheets[name],at,3,target,null});
      insert.Invoke(manager,new object[]{grouped.Worksheets[name],at,3,target,actions});
     }
     foreach(var action in actions)action();
     int differences=0,checkedCells=0;var affected=new Dictionary<string,int>();
     foreach(var target in targets){
      string name=(string)Field(target,"WorksheetName");var a=sequential.Worksheets[name];var b=grouped.Worksheets[name];
      var addresses=new HashSet<string>(a.GetUsedRange().ExistingCells.Where(c=>c.HasFormula).Select(c=>c.GetReferenceA1()));
      addresses.UnionWith(b.GetUsedRange().ExistingCells.Where(c=>c.HasFormula).Select(c=>c.GetReferenceA1()));
      foreach(string address in addresses){checkedCells++;if(a.Cells[address].FormulaInvariant==b.Cells[address].FormulaInvariant)continue;
       if(!affected.ContainsKey(name))affected[name]=0;affected[name]++;
       if(differences++<5)Console.WriteLine("ORDER_DIFF\t"+entry.Key+"\t"+name+"!"+address+"\tsequential="+a.Cells[address].FormulaInvariant+"\tgrouped="+b.Cells[address].FormulaInvariant);
      }
     }
     Console.WriteLine("ORDER_RESULT\t"+entry.Key+"\tchecked="+checkedCells+"\tdifferences="+differences+"\t"+String.Join(";",affected.Select(p=>p.Key+":"+p.Value)));
    }
    GC.Collect();GC.WaitForPendingFinalizers();
   }
   Console.WriteLine("PASS: All rule metadata inventoried; all multi-sheet column copy orders compared on unsaved private models. Not a complete VBA/financial validation.");return 0;
  }catch(Exception e){Console.Error.WriteLine(e);return 1;}
 }
}
