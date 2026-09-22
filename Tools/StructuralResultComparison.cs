using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using DevExpress.Spreadsheet;

// Read-only engineering comparison. No application model, macros, calculation or save.
public static class StructuralResultComparison {
    static readonly string[] RecordNames={"FacilityNames","LoanDescsOrd","HouseTypeInID","HouseTypeInMY","LastIDColNum","LastMYColNum","Transactional_Records"};
    static readonly JavaScriptSerializer Json=new JavaScriptSerializer {MaxJsonLength=int.MaxValue,RecursionLimit=100};
    static bool Sensitive(string name){return Regex.IsMatch(name,"password|passwd|credential|rejdata|secret|token",RegexOptions.IgnoreCase);}
    static void Save(string path,object value){File.WriteAllText(path,Json.Serialize(value));}
    static string Value(CellValue v){
        if(v.IsEmpty)return "empty:";
        if(v.IsNumeric)return "number:"+v.NumericValue.ToString("R",CultureInfo.InvariantCulture);
        if(v.IsBoolean)return "boolean:"+v.BooleanValue;
        if(v.IsError)return "error:"+v.ToString();
        return "text:"+v.ToString();
    }
    static bool SameValue(CellValue a,CellValue b){
        if(a.IsNumeric && b.IsNumeric)return Math.Abs(a.NumericValue-b.NumericValue)<=0.0000001;
        return Value(a)==Value(b);
    }
    static string Geometry(DefinedName n){
        try {var r=n.Range;return r==null?"non-range":r.Worksheet.Name+"!"+r.GetReferenceA1()+" ["+r.RowCount+"x"+r.ColumnCount+"]";}
        catch{return "unresolved";}
    }
    static Dictionary<string,DefinedName> Names(IWorkbook w){
        var result=new Dictionary<string,DefinedName>(StringComparer.OrdinalIgnoreCase);
        foreach(var n in w.DefinedNames)if(!Sensitive(n.Name))result["global::"+n.Name]=n;
        foreach(var s in w.Worksheets)foreach(var n in s.DefinedNames)if(!Sensitive(n.Name))result[s.Name+"::"+n.Name]=n;
        return result;
    }
    static Workbook Load(string path){
        var w=new Workbook();w.Options.CalculationMode=WorkbookCalculationMode.Manual;
        using(var stream=File.OpenRead(path)){w.LoadDocument(stream,DocumentFormat.Xlsb);}
        w.Options.CalculationMode=WorkbookCalculationMode.Manual;
        return w;
    }
    static Dictionary<string,object> Inventory(string path,string output){
        Console.WriteLine("Inspecting "+Path.GetFileName(path));
        using(var w=Load(path)){
            var ranges=new Dictionary<string,object>();
            foreach(var n in w.DefinedNames.Where(n=>RecordNames.Contains(n.Name)||n.Name.StartsWith("TransCopy_FacilityNames_")||n.Name.StartsWith("TransCopy_LoanDescsOrd_")||n.Name.StartsWith("TransCopy_DevptSingle_")||n.Name.StartsWith("TransCopy_DevptMulti_"))) {
                try {var r=n.Range;ranges[n.Name]=new {reference=n.RefersTo,rows=r.RowCount,columns=r.ColumnCount};}
                catch {ranges[n.Name]="unresolved";}
            }
            var version=w.DefinedNames.GetDefinedName("ModelVersion");
            var checks=new List<object>();var sheet=w.Worksheets.FirstOrDefault(s=>s.Name=="Check Sheet");
            if(sheet!=null)foreach(var c in sheet.GetUsedRange().ExistingCells){
                var v=c.Value;string text=v.ToString();
                if(v.IsError||Regex.IsMatch(text,"^(OK|ERROR|CHECK|FAIL|WARNING|NOT OK)$",RegexOptions.IgnoreCase))
                    checks.Add(new {cell=c.GetReferenceA1(),value=Value(v)});
            }
            var inv=new Dictionary<string,object>{{"path",path},{"modelVersion",version==null?"missing":version.Range[0,0].Value.ToString()},
                {"sheetNames",w.Worksheets.Select(s=>s.Name).ToArray()},{"globalNames",w.DefinedNames.Count},{"allNames",Names(w).Count},
                {"protectedSheets",w.Worksheets.Count(s=>s.IsProtected)},{"ranges",ranges},{"savedChecks",checks},
                {"outputSheets",w.Worksheets.Where(s=>Regex.IsMatch(s.Name,"SOCI|Comp Income|Comprehensive|Financial Position|Detailed Cashflow",RegexOptions.IgnoreCase)).Select(s=>s.Name).ToArray()}};
            Save(output,inv);
            foreach(string name in RecordNames)if(ranges.ContainsKey(name))Console.WriteLine(name+" "+Json.Serialize(ranges[name]));
            return inv;
        }
    }
    static Dictionary<long,Cell> Cells(Worksheet s){
        return s.GetUsedRange().ExistingCells.Where(c=>c.HasFormula||!c.Value.IsEmpty).ToDictionary(c=>(long)c.RowIndex*16384+c.ColumnIndex,c=>c);
    }
    static string Clip(string s){return s==null?null:(s.Length<=1000?s:s.Substring(0,1000)+" [truncated]");}
    static string FormulaKey(string s){
        if(s==null)return null;
        return String.Concat(Regex.Split(s,"(\"(?:[^\"]|\"\")*\")").Select((part,i)=>i%2==0?part.ToUpperInvariant():part));
    }
    static void Compare(string left,string right,string output,bool includeTdb=false){
        Console.WriteLine("Comparing "+Path.GetFileName(left)+" AGAINST "+Path.GetFileName(right));
        using(var a=Load(left))using(var b=Load(right)){
            var namesA=Names(a);var namesB=Names(b);var nameDiffs=new List<object>();
            foreach(string key in namesA.Keys.Union(namesB.Keys,StringComparer.OrdinalIgnoreCase)){
                DefinedName x=null,y=null;namesA.TryGetValue(key,out x);namesB.TryGetValue(key,out y);
                if(x==null||y==null||x.RefersTo!=y.RefersTo)
                    nameDiffs.Add(new {name=key,left=x==null?null:Geometry(x),right=y==null?null:Geometry(y),sameGeometry=x!=null&&y!=null&&Geometry(x)==Geometry(y)});
            }
            var sheets=new List<object>();var samples=new List<object>();
            int formulaTotal=0,formulaDiffTotal=0,inputDiffTotal=0,valueDiffTotal=0,arrayDiffTotal=0,caseOnlyTotal=0;
            foreach(var sa in a.Worksheets){
                var sb=b.Worksheets.FirstOrDefault(s=>s.Name==sa.Name);if(sb==null)continue;
                // This sheet's per-cell native introspection is prohibitively slow.
                // Compare its formulas/constants separately through bulk read-only Excel.
                if(sa.Name=="Transactional DB"&&!includeTdb){Console.WriteLine("Deferred Transactional DB to independent Excel bulk comparison");continue;}
                Console.WriteLine("Checking "+sa.Name);
                var ca=Cells(sa);var cb=Cells(sb);int nf=0,fd=0,id=0,vd=0,ad=0,cd=0,sampleCount=0;
                bool outputs=Regex.IsMatch(sa.Name,"SOCI|Comp Inc|Comprehensive|Financial Position|Cashflow detailed|^Check Sheet$",RegexOptions.IgnoreCase);
                var sampleKinds=new Dictionary<string,int>();
                foreach(long key in ca.Keys.Union(cb.Keys)){
                    Cell x=null,y=null;ca.TryGetValue(key,out x);cb.TryGetValue(key,out y);
                    bool xf=x!=null&&x.HasFormula,yf=y!=null&&y.HasFormula;
                    string kind=null,l=null,r=null;
                    if(xf||yf){
                        nf++;string fx=xf?x.FormulaInvariant:null,fy=yf?y.FormulaInvariant:null;
                        if(fx!=fy){if(FormulaKey(fx)==FormulaKey(fy)){cd++;}else{fd++;kind="formula";l=fx;r=fy;}}
                        if(sa.Name!="Transactional DB"&&xf&&yf && x.HasArrayFormula!=y.HasArrayFormula){ad++;if(kind==null){kind="array-type";l=x.HasArrayFormula.ToString();r=y.HasArrayFormula.ToString();}}
                    } else if(x==null||y==null||!SameValue(x.Value,y.Value)){
                        id++;kind="constant";l=x==null?null:Value(x.Value);r=y==null?null:Value(y.Value);
                    }
                    if(outputs && (x==null||y==null||!SameValue(x.Value,y.Value))){
                        vd++;if(kind==null){kind="saved-output";l=x==null?null:Value(x.Value);r=y==null?null:Value(y.Value);}
                    }
                    if(kind!=null && sampleCount<30 && samples.Count<1500){
                        int n=sampleKinds.ContainsKey(kind)?sampleKinds[kind]:0;
                        if(n<10){samples.Add(new {sheet=sa.Name,cell=(x??y).GetReferenceA1(),kind,left=Clip(l),right=Clip(r)});sampleKinds[kind]=n+1;sampleCount++;}
                    }
                }
                formulaTotal+=nf;formulaDiffTotal+=fd;inputDiffTotal+=id;valueDiffTotal+=vd;arrayDiffTotal+=ad;caseOnlyTotal+=cd;
                sheets.Add(new {sheet=sa.Name,formulaCells=nf,formulaDifferences=fd,formulaCaseOnlyDifferences=cd,constantDifferences=id,savedOutputDifferences=vd,arrayTypeDifferences=ad,leftProtected=sa.IsProtected,rightProtected=sb.IsProtected});
                Save(output+".partial",new {complete=false,nameDifferences=nameDiffs,formulaTotal,formulaDiffTotal,caseOnlyTotal,inputDiffTotal,valueDiffTotal,arrayDiffTotal,sheets,samples});
                if(fd+id+vd+ad>0)Console.WriteLine(sa.Name+": formulas="+fd+", constants="+id+", savedOutputs="+vd+", arrays="+ad);
            }
            Save(output,new {complete=true,excludedSheets=includeTdb?new string[0]:new[]{"Transactional DB"},arrayTypeExcludedSheets=new[]{"Transactional DB"},sameSheetOrder=a.Worksheets.Select(s=>s.Name).SequenceEqual(b.Worksheets.Select(s=>s.Name)),left,right,absoluteNumericTolerance=0.0000001,leftOnlySheets=a.Worksheets.Select(s=>s.Name).Except(b.Worksheets.Select(s=>s.Name)).ToArray(),rightOnlySheets=b.Worksheets.Select(s=>s.Name).Except(a.Worksheets.Select(s=>s.Name)).ToArray(),nameDifferences=nameDiffs,formulaCells=formulaTotal,formulaDifferences=formulaDiffTotal,formulaCaseOnlyDifferences=caseOnlyTotal,constantDifferences=inputDiffTotal,savedOutputDifferences=valueDiffTotal,arrayTypeDifferences=arrayDiffTotal,sheets,samples});
            Console.WriteLine("TOTAL: formulaCells="+formulaTotal+", formulaDifferences="+formulaDiffTotal+", constantDifferences="+inputDiffTotal+", outputDifferences="+valueDiffTotal+", nameDifferences="+nameDiffs.Count);
        }
    }
    [STAThread] public static int Main(string[] args){try{
        AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
        if(args.Length==4){Compare(args[2],args[3],Path.Combine(args[1],"comparison.json"),true);return 0;}
        string[] labels={"baseline","funding-summit","funding-excel","development-summit","development-excel"};
        for(int i=0;i<labels.Length;i++){Inventory(args[i+2],Path.Combine(args[1],labels[i]+".json"));GC.Collect();GC.WaitForPendingFinalizers();}
        Compare(args[3],args[4],Path.Combine(args[1],"funding-comparison.json"));GC.Collect();GC.WaitForPendingFinalizers();
        Compare(args[5],args[6],Path.Combine(args[1],"development-comparison.json"));
        return 0;
    }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
