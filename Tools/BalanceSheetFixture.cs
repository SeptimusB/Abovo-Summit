using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using System.Diagnostics;
using DevExpress.Spreadsheet;

// All destructive/negative-case mutations are fault injections on a disposable
// copy, not interactive application edits and never on a supplied source file.
public static class BalanceSheetFixture {
    static Assembly app;
    static object Static(string type,string method,params object[] args) { return app.GetType("Abovo."+type).GetMethod(method).Invoke(null,args); }
    static dynamic Read(IWorkbook w) { return Static("BalanceSheetStatement","Read",w); }
    static List<dynamic> Nodes(dynamic d) { return ((IEnumerable)d.Nodes).Cast<object>().Select(x=>(dynamic)x).ToList(); }
    static void Check(bool good,string what) { if(!good)throw new Exception(what);Console.WriteLine("PASS: "+what); }
    static void Reconcile(dynamic d) {
        var nodes=Nodes((object)d);
        Check(String.IsNullOrEmpty((string)d.Diagnostic),"No unsupported/drill reconciliation diagnostics");
        int points=0;
        foreach(var parent in nodes) {
            string id=parent.Id;
            var children=nodes.Where(n=>(string)n.ParentId==id).ToList();
            if(children.Count==0)continue;
            for(int p=0;p<41;p++) {
                double sum=children.Sum(n=>(double)n.Values[p]);
                if(Math.Abs(sum-(double)parent.Values[p])>0.001)throw new Exception("Branch mismatch: "+id+" period="+p);
                points++;
            }
        }
        Check(points>1500,"Every populated hierarchy branch reconciles: "+points+" values");
    }
    static void Trad(IWorkbook w,dynamic doc) {
        int count=0, textChecks=0, roundedChecks=0;
        foreach(var n in Nodes((object)doc).Where(n=>(bool)n.IsHeadline)) {
            int offset=Int32.Parse(((string)n.Id).Substring(3));
            for(int p=0;p<41;p++) {
                var v=w.Worksheets["Financial Position - Trad View"].Cells[8+offset,2+p].Value;
                if(offset==58 && !v.IsNumeric) {
                    Console.WriteLine("NOTE: Trad Check is nonnumeric in period "+p+": '"+v.ToString()+"'; TDB Check="+n.Values[p]);
                    textChecks++;continue;
                }
                if(offset==58 && v.IsNumeric && Math.Abs(v.NumericValue-(double)n.Values[p])>0.001 &&
                   Math.Abs(v.NumericValue-(double)n.Values[p])<=0.500001 &&
                   w.Worksheets["Transactional DB"].Cells[(int)doc.OutputTop+offset,p==0?14:15+p].FormulaInvariant.ToUpperInvariant().Contains("ROUND(")) {
                    roundedChecks++;continue;
                }
                if(!v.IsNumeric || Math.Abs(v.NumericValue-(double)n.Values[p])>0.001)throw new Exception("Trad mismatch "+n.Id+" "+p+" Trad="+v.ToString()+" TDB="+n.Values[p]);
                count++;
            }
        }
        Check(count+textChecks+roundedChecks==1886,"Numeric headline/opening values match Trad View: "+count+"; workbook-rounded Check cells: "+roundedChecks+"; nonnumeric Trad Check cells: "+textChecks);
    }
    static void Reject(Action action,string message) {
        try { action(); }catch(TargetInvocationException e) {if(e.InnerException is InvalidOperationException){Console.WriteLine("PASS: "+message+": "+e.InnerException.Message);return;}throw;}
        throw new Exception("Expected rejection: "+message);
    }
    static void Calc(IWorkbook w) {
        var previous=w.Options.CalculationEngineType;
        try { w.Options.CalculationEngineType=CalculationEngineType.Recursive;w.CalculateFull(); }
        finally { w.Options.CalculationEngineType=previous; }
    }
    [STAThread] public static int Main(string[] args) {
        try {
            AppDomain.CurrentDomain.AssemblyResolve+=(sender,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
            Application.EnableVisualStyles();
            app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
            Static("AbovoAppCls","Initialise");Static("FileManager","Initialise",new object[]{null});
            string copy=Path.Combine(args[2],"private-source.xlsb");File.Copy(args[1],copy);
            var open=app.GetType("Abovo.FileManager").GetMethod("OpenModel");
            dynamic result=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
            Check(!result.BError,"Opened private workbook");
            dynamic model=((Array)app.GetType("Abovo.FileManager").GetField("ExcelModels").GetValue(null)).GetValue(0);
            IWorkbook w=model.WB;
            model.WBCalcEngine.CalculateDependencySensitiveFile("Balance Sheet fixture",true);
            var watch=Stopwatch.StartNew();dynamic live=Read(w);
            Console.WriteLine("METRIC adapter "+watch.ElapsedMilliseconds+" ms; nodes="+live.Nodes.Count);
            Trad(w,live);Reconcile(live);
            if(!w.Worksheets.Contains("TDB Snapshot") || !w.Worksheets.Contains("TDB Comparison")) {
                Console.WriteLine("PASS: Live legacy workbook reconciles. Snapshot tests not applicable: dedicated sheets absent; no schema added.");
                return 0;
            }
            Reject(()=>Static("BalanceSheetSnapshot","Read",w,(object)live),"Legacy snapshot has no BS capture");
            var s=w.Worksheets["TDB Snapshot"];var c=w.Worksheets["TDB Comparison"];
            // Explicit collision must fail before either output is cleared.
            s.Cells[0,100].Value="Client-owned sentinel";
            Reject(()=>Static("TransactionalDBSnapshotManager","CreateSnapshotAndComparison",0),"Client-owned content protected");
            Check(s.Cells[0,100].Value.ToString()=="Client-owned sentinel","Collision sentinel retained");s.Cells[0,100].ClearContents();
            var permissions=WorksheetProtectionPermissions.SelectLockedCells|WorksheetProtectionPermissions.SelectUnlockedCells;
            Static("WSSecurity","UNProtectWS",0,s.Name);
            Static("WSSecurity","UNProtectWS",0,c.Name);
            Static("WSSecurity","ProtectWS",0,s.Name,permissions);
            Static("WSSecurity","ProtectWS",0,c.Name,permissions);
            permissions=s.GetProtectionPermissions();
            watch.Restart();Static("TransactionalDBSnapshotManager","CreateSnapshotAndComparison",0);
            Console.WriteLine("METRIC snapshot "+watch.ElapsedMilliseconds+" ms");
            Check(s.IsProtected && c.IsProtected && s.GetProtectionPermissions()==permissions,"Capture restores protection and exact permissions");
            dynamic frozen=Static("BalanceSheetSnapshot","Read",w,(object)live);
            Reconcile(frozen);Trad(w,frozen);
            dynamic zero=Static("BalanceSheetStatement","Difference",(object)live,(object)frozen);
            Check(Nodes((object)zero).All(n=>((double[])n.Values).All(v=>Math.Abs(v)<0.001)),"New capture differences zero at every hierarchy level");
            string saved=Path.Combine(args[2],"snapshot-roundtrip.xlsb");w.SaveDocument(saved,DocumentFormat.Xlsb);
            using(var reopened=new Workbook()) {
                reopened.Options.CalculationMode=WorkbookCalculationMode.Manual;
                reopened.LoadDocument(saved,DocumentFormat.Xlsb);
                dynamic loaded=Read(reopened);
                dynamic restored=Static("BalanceSheetSnapshot","Read",reopened,(object)loaded);
                Check(Nodes((object)restored).Count==Nodes((object)frozen).Count,"Frozen detail survives XLSB save/reopen");
                Check(Nodes((object)restored).Zip(Nodes((object)frozen),(a,b)=>((double[])a.Values).SequenceEqual((double[])b.Values)).All(x=>x),"Every frozen value survives save/reopen");
            }
            // Perturb an opening precedent and a classification without changing
            // transaction geometry. This simulates external Excel edits.
            var tdb=w.Worksheets["Transactional DB"];
            var opening=w.Worksheets["Financial Position - Trad View"].Cells[10,2];
            var openingFormula=opening.FormulaInvariant;var openingValue=opening.Value;
            opening.Value=openingValue.NumericValue+123.5;
            int input=163;
            while(input<1470 && (!tdb.Cells[input,7].Value.IsText || tdb.Cells[input,7].Value.ToString()=="" || !tdb.Cells[input,16].Value.IsNumeric || Math.Abs(tdb.Cells[input,16].Value.NumericValue)<0.001))input++;
            if(input==1470)input=163;
            var headingCell=tdb.Cells[input,7];string headingFormula=headingCell.FormulaInvariant;var headingValue=headingCell.Value;
            var movementCell=tdb.Cells[input,16];string movementFormula=movementCell.FormulaInvariant;var movementValue=movementCell.Value;
            var cashCell=tdb.Cells[input,5];string cashFormula=cashCell.FormulaInvariant;var cashValue=cashCell.Value;
            headingCell.Value=tdb.Cells[(int)live.OutputTop+2,7].Value;cashCell.Value="Cash";movementCell.Value=777.25;
            Calc(w);dynamic changed=Read(w);Reconcile(changed);
            dynamic unchanged=Static("BalanceSheetSnapshot","Read",w,(object)changed);
            Check(Nodes((object)unchanged).Zip(Nodes((object)frozen),(a,b)=>((double[])a.Values).SequenceEqual((double[])b.Values)).All(x=>x),"Openings/classifications remain frozen after live edits");
            dynamic delta=Static("BalanceSheetStatement","Difference",(object)changed,(object)frozen);Reconcile(delta);
            Check(Nodes((object)delta).Any(n=>((double[])n.Values).Any(v=>Math.Abs(v)>0.001)),"Differences reflect opening and transaction edits");
            Restore(opening,openingFormula,openingValue);Restore(headingCell,headingFormula,headingValue);Restore(movementCell,movementFormula,movementValue);Restore(cashCell,cashFormula,cashValue);Calc(w);
            // Formula and locator negative tests must never invent a mapping.
            var rule=tdb.Cells[(int)live.OutputTop+2,16];string formula=rule.FormulaInvariant;rule.FormulaInvariant=formula+"+1";rule.Calculate();
            dynamic unsupported=Read(w);
            Check(Nodes((object)unsupported).Any(n=>(string)n.Id=="bs/02" && ((string)n.Diagnostic).Length>0),"Bespoke formula keeps headline but disables guessed drill");
            Reject(()=>Static("BalanceSheetSnapshot","Read",w,(object)unsupported),"Rule changes invalidate frozen BS mapping");rule.FormulaInvariant=formula;Calc(w);
            var marker=tdb.Cells[(int)live.OutputTop-1,0];var markerValue=marker.Value;marker.Value="Missing SOFP";
            Reject(()=>Read(w),"Missing locator rejected");marker.Value=markerValue;
            tdb.Cells[(int)live.OutputBottom+5,0].Value=markerValue;
            Reject(()=>Read(w),"Ambiguous locator rejected");tdb.Cells[(int)live.OutputBottom+5,0].ClearContents();
            // Real workbook structural service and transactional mirror lifecycle.
            int before=(int)live.OutputTop;
            dynamic inserted=Static("WorkbookManager","InsertRows",0,"IR_Cash_Journals",5,false,false);
            Check(!inserted.BError,"Real five-row Cash Journals insertion");
            Check(!(bool)Static("TransactionalDBSnapshotManager","HasValidSnapshot",0),"Structural change invalidates complete snapshot without an analyser");
            model.WBCalcEngine.CalculateDependencySensitiveFile("Balance Sheet shifted fixture",true);
            dynamic shifted=Read(w);Check((int)shifted.OutputTop==before+5,"SOFP locator re-resolves after actual mirror expansion");Reconcile(shifted);Trad(w,shifted);
            Reject(()=>Static("BalanceSheetSnapshot","Read",w,(object)shifted),"Shifted BS cannot reuse old capture");
            Check(s.IsProtected && c.IsProtected && s.GetProtectionPermissions()==permissions,"Invalidation preserves protection");
            Static("TransactionalDBSnapshotManager","CreateSnapshotAndComparison",0);
            dynamic recaptured=Static("BalanceSheetSnapshot","Read",w,(object)shifted);Reconcile(recaptured);
            // Reserve rules deliberately have a wider range than Transactional_Records.
            // One qualifying row may legitimately contribute twice, not be deduplicated.
            var records=w.DefinedNames.GetDefinedName("Transactional_Records").Range;
            int extraRow=records.BottomRowIndex+8;
            var extraCells=new[]{tdb.Cells[extraRow,6],tdb.Cells[extraRow,57],tdb.Cells[extraRow,68],tdb.Cells[extraRow,16]};
            var extraFormulas=extraCells.Select(cell=>cell.FormulaInvariant).ToArray();
            var extraValues=extraCells.Select(cell=>cell.Value).ToArray();
            double reserveBefore=(double)Nodes((object)shifted).First(n=>(string)n.Id=="bs/51").Values[1];
            try {
                var ruleText=tdb.Cells[(int)shifted.OutputTop+51,16].FormulaInvariant;
                var pension=System.Text.RegularExpressions.Regex.Match(ruleText.Replace("$",""),",G([0-9]+),Q");
                Check(pension.Success,"Reserve criterion resolves from actual formula");
                extraCells[0].Value=tdb.Cells[Int32.Parse(pension.Groups[1].Value)-1,6].Value;
                extraCells[1].Value=0;extraCells[2].Value=1;extraCells[3].Value=123.25;
                Calc(w);dynamic overlap=Read(w);Reconcile(overlap);
                var reserve=Nodes((object)overlap).First(n=>(string)n.Id=="bs/51");
                Check(Math.Abs((double)reserve.Values[1]-reserveBefore-246.5)<0.001,"Extra-span reserve row contributes twice under overlapping criteria");
                Static("TransactionalDBSnapshotManager","CreateSnapshotAndComparison",0);
                dynamic overlapCapture=Static("BalanceSheetSnapshot","Read",w,(object)overlap);Reconcile(overlapCapture);
                Check(s.Cells[extraRow,16].Value.NumericValue==123.25,"Snapshot physically includes extra reserve input");
                Check(Nodes((object)overlapCapture).Any(n=>((string)n.Source).Contains("Q"+(extraRow+1)+":") && ((string)n.Rule).StartsWith("Coefficient 2")),"Captured detail retains both reserve contributions");
            } finally {
                for(int i=0;i<extraCells.Length;i++)Restore(extraCells[i],extraFormulas[i],extraValues[i]);Calc(w);
            }
            dynamic deleted=Static("WorkbookManager","DeleteRows",0,"IR_Cash_Journals",5,false,false,3);
            Check(!deleted.BError,"Real five-row Cash Journals deletion");
            Check(!(bool)Static("TransactionalDBSnapshotManager","HasValidSnapshot",0),"Deletion invalidates snapshot without analyser");
            model.WBCalcEngine.CalculateDependencySensitiveFile("Balance Sheet contracted fixture",true);
            dynamic contracted=Read(w);Check((int)contracted.OutputTop==before,"SOFP locator re-resolves after mirror contraction");Reconcile(contracted);Trad(w,contracted);
            Console.WriteLine("PASS: Balance Sheet live, frozen, differences, persistence, structural and negative cases. Source never saved.");
            return 0;
        }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}
    }
    static void Restore(Cell cell,string formula,CellValue value){if(String.IsNullOrEmpty(formula))cell.Value=value;else cell.FormulaInvariant=formula;}
}
