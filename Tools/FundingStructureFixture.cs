using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using DevExpress.Spreadsheet;

public static class FundingStructureFixture {
    const BindingFlags F=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
    static object Field(object o,string n){return o.GetType().GetField(n,F).GetValue(o);}
    static void Check(bool v,string m){if(!v)throw new Exception(m);Console.WriteLine("PASS: "+m);}
    static Dictionary<string,string> Formulas(IWorkbook w,IEnumerable<string> sheets){
        var result=new Dictionary<string,string>();
        foreach(string s in sheets)foreach(var c in w.Worksheets[s].GetUsedRange().ExistingCells.Where(c=>c.HasFormula))
            result.Add(s+"!"+c.GetReferenceA1(),c.FormulaInvariant);
        return result;
    }
    [STAThread] public static int Main(string[] args){
        try{
            AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
            var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
            app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
            var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
            string copy=Path.Combine(args[2],"input-copy.xlsb");File.Copy(args[1],copy);
            var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
            Check(!loaded.BError,"Opened private workbook");
            dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);
            IWorkbook w=(IWorkbook)model.WB;
            object rules=model.WorkbookStructureRules;
            var rule=rules.GetType().GetMethod("GetRule",F).Invoke(rules,new object[]{"FUNDING_RECORDS"});
            var targets=((IEnumerable)Field(rule,"Targets")).Cast<object>().Select(t=>Convert.ToString(Field(t,"WorksheetName"))).ToArray();
            Check(targets.Length==32,"All 32 VBA Funding sheets targeted");
            File.WriteAllLines(Path.Combine(args[2],"funding-sheets.txt"),targets);
            var protection=w.Worksheets.ToDictionary(s=>s.Name,s=>s.IsProtected);
            var visibility=w.Worksheets.ToDictionary(s=>s.Name,s=>s.Visible);
            var mode=w.Options.CalculationMode;var engine=w.Options.CalculationEngineType;
            var names=w.DefinedNames.ToDictionary(n=>n.Name,n=>n.RefersTo);
            var formulas=Formulas(w,targets);
            int count=8,facilities=w.DefinedNames.GetDefinedName("FacilityNames").Range.ColumnCount;
            int ordinary=w.DefinedNames.GetDefinedName("LoanDescsOrd").Range.ColumnCount;
            var mirrors=w.DefinedNames.Where(n=>n.Name.StartsWith("TransCopy_FacilityNames_")||n.Name.StartsWith("TransCopy_LoanDescsOrd_")).ToDictionary(n=>n.Name,n=>n.Range.RowCount);
            Check(mirrors.Count==11,"Nine facility and two ordinary-loan mirrors present");
            // Unprotected/visible disposable reference input for Excel's grouped
            // operation. No password or raw VBA is exported or executed.
            var security=app.GetType("Abovo.WSSecurity");
            foreach(string s in targets){security.GetMethod("UNProtectWS").Invoke(null,new object[]{0,s});w.Worksheets[s].Visible=true;}
            w.SaveDocument(Path.Combine(args[2],"excel-reference-input.xlsb"),DocumentFormat.Xlsb);
            foreach(string s in targets){if(protection[s])security.GetMethod("ProtectWS").Invoke(null,new object[]{0,s,Type.Missing});w.Worksheets[s].Visible=visibility[s];}
            dynamic result;
            using(var trace=new StringWriter())
            using(var listener=new TextWriterTraceListener(trace))
            using(var owner=new System.Windows.Forms.Form()){
                Trace.Listeners.Add(listener);
                try{
                    var tagType=app.GetType("Abovo.PresentationManager+AttachedGridCommandButton");
                    var tag=Activator.CreateInstance(tagType);
                    tagType.GetField("CommandData").SetValue(tag,"ProcessAddFundingRecords");
                    tagType.GetField("RequestedRecordCount").SetValue(tag,count);
                    bool progressStarted=false;
                    tagType.GetField("StructuralProgress").SetValue(tag,new Action<string>(text=>{
                        progressStarted=text!=null;
                        Check(text!=null&&text.Contains(count.ToString()),"Funding progress receives accepted count");
                        Check(w.DefinedNames.GetDefinedName("FacilityNames").Range.ColumnCount==facilities,"Progress begins before workbook mutation");
                    }));
                    result=model.EventCoordinator.TriggerEvent("GridButton",tag,owner);
                    listener.Flush();
                    var log=trace.ToString();File.WriteAllText(Path.Combine(args[2],"insert-trace.log"),log);
                    Check(progressStarted,"Funding button service invokes progress callback");
                    foreach(string stage in new[]{"capture3D","unprotect","protect","shiftColumns","copyTemplate","copy3D","apply3D","endUpdate","restoreCalculationMode","postActions"})
                        Check(log.Contains("stage="+stage),"Funding trace stage: "+stage);
                    Check(log.Contains("state=finished")&&log.Contains("outcome=ok"),"Funding trace completes successfully");
                    Check(log.Split(new[]{Environment.NewLine},StringSplitOptions.None).Count(s=>s.Contains("stage=shiftColumns")&&s.Contains("elapsed="))==32,"Funding trace covers all 32 shifts");
                }finally{Trace.Listeners.Remove(listener);}
            }
            Check(!result.BError,"Eight Funding columns inserted: "+result.StringReturn);
            Check(w.DefinedNames.GetDefinedName("FacilityNames").Range.ColumnCount==facilities+count,"FacilityNames expanded");
            Check(w.DefinedNames.GetDefinedName("LoanDescsOrd").Range.ColumnCount==ordinary+count,"Ordinary loan names expanded");
            foreach(var m in mirrors)Check(w.DefinedNames.GetDefinedName(m.Key).Range.RowCount==m.Value+count,"Mirror resized: "+m.Key);
            Check(w.Worksheets.All(s=>s.IsProtected==protection[s.Name] && s.Visible==visibility[s.Name]),"Protection and visibility preserved");
            Check(w.Options.CalculationMode==mode && w.Options.CalculationEngineType==engine && model.IsDirty,"Calculation settings restored and workbook dirty");
            Check((bool)model.SaveFileAsTo(Path.Combine(args[2],"funding-expanded.xlsb"),true),"Expanded XLSB saved through model service");
            foreach(int index in new[]{0,9,ordinary+count}){
                var before=w.DefinedNames.GetDefinedName("FacilityNames").RefersTo;
                result=model.WorkbookStructureRules.DeleteRecords("FUNDING_RECORDS",new[]{index});
                Check(result.BError && before==w.DefinedNames.GetDefinedName("FacilityNames").RefersTo && !model.IsDirty,"Protected ordinary/revolver deletion rejected without mutation: "+index);
            }
            result=model.WorkbookStructureRules.DeleteLastRecords("FUNDING_RECORDS",count);
            Check(!result.BError,"Delete last removes ordinary columns, not revolvers");
            foreach(var n in names)Check(w.DefinedNames.GetDefinedName(n.Key).RefersTo==n.Value,"Restored name: "+n.Key);
            var restored=Formulas(w,targets);
            var differences=formulas.Where(p=>!restored.ContainsKey(p.Key)||restored[p.Key]!=p.Value).Take(8).ToArray();
            foreach(var p in differences)Console.WriteLine("DIFFERENCE "+p.Key+" before="+p.Value+" after="+(restored.ContainsKey(p.Key)?restored[p.Key]:"missing"));
            Check(differences.Length==0 && formulas.Count==restored.Count,"All linked-sheet formulas restored after add/delete");
            result=model.WorkbookStructureRules.ValidateDeleteLastRecords("FUNDING_RECORDS",1);
            Check(result.BError,"Cannot delete below ten ordinary-loan columns");
            Check(w.Worksheets.All(s=>s.IsProtected==protection[s.Name] && s.Visible==visibility[s.Name]),"Delete preserves worksheet state");
            Console.WriteLine("PASS: Funding add/delete safeguard fixture. Original workbook not saved.");return 0;
        }catch(Exception e){Console.Error.WriteLine(e);return 1;}
    }
}
