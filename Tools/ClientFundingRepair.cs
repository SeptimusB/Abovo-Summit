using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;

// Explicit offline investigation/repair. Never a load-time migration.
public static class ClientFundingRepair {
    public class Patch {public string sheet;public string cell;public int row;public int column;public string before;public string expected;public bool array;public bool linked;public bool locked;}
    public class NamePatch {public string name;public string before;public string expected;}
    public class Manifest {public Patch[] differences;public NamePatch[] nameDifferences;}
    static JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength=int.MaxValue };
    static Workbook Load(string path) {
        var w=new Workbook();w.Options.CalculationMode=WorkbookCalculationMode.Manual;
        using(var stream=File.OpenRead(path)) w.LoadDocument(stream,DocumentFormat.Xlsb);
        w.Options.CalculationMode=WorkbookCalculationMode.Manual;return w;
    }
    class Refs:ExpressionVisitor {
        public readonly List<CellReferenceExpression> Items=new List<CellReferenceExpression>();
        public override void Visit(CellReferenceExpression e){Items.Add(e);base.Visit(e);}
    }
    static List<CellReferenceExpression> References(IWorkbook w,string formula) {
        var v=new Refs();w.FormulaEngine.Parse(formula).Expression.Visit(v);return v.Items;
    }
    static bool ThreeDFunding(IWorkbook w,string formula,HashSet<string> sheets){
        if(!sheets.Any(s=>formula.IndexOf(s,StringComparison.OrdinalIgnoreCase)>=0))return false;
        return References(w,formula).Any(r=>r.SheetReference!=null&&sheets.Contains(r.SheetReference.StartSheetName));
    }
    static string Text(IWorkbook w,IExpression e){var b=new StringBuilder();e.BuildExpressionString(b,w);return "="+b;}
    static void Require(bool value,string reason){if(!value)throw new InvalidOperationException(reason);}
    static void Prove(IWorkbook w,string before,string expected,string sheet,int column,HashSet<string> sheets){
        var pa=w.FormulaEngine.Parse(before);var pe=w.FormulaEngine.Parse(expected);var va=new Refs();var ve=new Refs();
        pa.Expression.Visit(va);pe.Expression.Visit(ve);Require(va.Items.Count==ve.Items.Count,"Reference count differs");int changed=0;
        for(int i=0;i<va.Items.Count;i++){
            var a=va.Items[i];var b=ve.Items[i];var x=a.CellArea;var y=b.CellArea;
            if(Text(w,a)==Text(w,b))continue;
            Require(a.SheetReference!=null&&b.SheetReference!=null&&a.SheetReference.Type==b.SheetReference.Type&&a.SheetReference.StartSheetName==b.SheetReference.StartSheetName&&a.SheetReference.EndSheetName==b.SheetReference.EndSheetName&&sheets.Contains(a.SheetReference.StartSheetName),"Worksheet reference changed");
            Require(x.TopRowIndex==y.TopRowIndex&&x.BottomRowIndex==y.BottomRowIndex&&x.TopLeft.RowType==y.TopLeft.RowType&&x.BottomRight.RowType==y.BottomRight.RowType&&x.TopLeft.ColumnType==y.TopLeft.ColumnType&&x.BottomRight.ColumnType==y.BottomRight.ColumnType,"Row/anchoring differs");
            bool three=!String.IsNullOrEmpty(a.SheetReference.EndSheetName)&&a.SheetReference.StartSheetName!=a.SheetReference.EndSheetName;
            bool ordinary=x.LeftColumnIndex==4&&x.RightColumnIndex==13&&y.LeftColumnIndex==4&&y.RightColumnIndex==21;
            bool copied=sheets.Contains(sheet)&&column>=14&&column<=21&&x.LeftColumnIndex==column+8&&x.RightColumnIndex==column+8&&y.LeftColumnIndex==column&&y.RightColumnIndex==column;
            bool copy3D=three&&sheets.Contains(sheet)&&column>=14&&column<=21&&x.LeftColumnIndex==13&&x.RightColumnIndex==13&&y.LeftColumnIndex==column&&y.RightColumnIndex==column;
            bool moved3D=three&&x.LeftColumnIndex>=14&&x.RightColumnIndex<=17&&y.LeftColumnIndex==x.LeftColumnIndex+8&&y.RightColumnIndex==x.RightColumnIndex+8;
            Require(ordinary||copied||copy3D||moved3D,"Difference is not a proved Funding insertion signature");
            a.CellArea=b.CellArea.Clone();changed++;
        }
        Require(changed>0&&Text(w,pa.Expression)==Text(w,pe.Expression),"Non-reference formula content differs");
    }
    static void Preview(string source,string output,string sheetName="Funding Assumptions",int row=40,int column=11){
        using(var control=new DevExpress.XtraSpreadsheet.SpreadsheetControl())using(var form=new Form()){
            form.ClientSize=new Size(1600,800);control.Dock=DockStyle.Fill;form.Controls.Add(control);
            control.Document.Options.CalculationMode=WorkbookCalculationMode.Manual;control.LoadDocument(source);
            var ws=control.Document.Worksheets[sheetName];control.Document.Worksheets.ActiveWorksheet=ws;ws.ScrollTo(row,column);
            form.CreateControl();var handle=form.Handle;control.CreateControl();form.PerformLayout();control.Refresh();
            using(var bitmap=new Bitmap(control.Width,control.Height)){control.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(output);}
        }
    }
    static string TdbExpected(IWorkbook a,IWorkbook b,Cell cell,HashSet<string> sheets){
        var parsed=a.FormulaEngine.Parse(cell.FormulaInvariant);var refs=new Refs();parsed.Expression.Visit(refs);int changed=0;
        foreach(var r in refs.Items){var x=r.CellArea;
            if(r.SheetReference==null||!sheets.Contains(r.SheetReference.StartSheetName)||x.LeftColumnIndex!=4||x.RightColumnIndex!=13)continue;
            r.CellArea=new CellArea(x.TopLeft,new CellReferencePosition(21,x.BottomRight.Row,x.BottomRight.ColumnType,x.BottomRight.RowType));changed++;
        }
        if(changed==0)return cell.FormulaInvariant;
        var name=new[]{"TransCopy_LoanDescsOrd_A","TransCopy_LoanDescsOrd_B"}.Single(n=>{var r=a.DefinedNames.GetDefinedName(n).Range;return cell.RowIndex>=r.TopRowIndex&&cell.RowIndex<r.BottomRowIndex;});
        var source=a.DefinedNames.GetDefinedName(name).Range;var donor=b.DefinedNames.GetDefinedName(name).Range;
        var donorCell=donor.Worksheet.Cells[donor.TopRowIndex+cell.RowIndex-source.TopRowIndex,cell.ColumnIndex];
        var dp=b.FormulaEngine.Parse(donorCell.FormulaInvariant);var dv=new Refs();dp.Expression.Visit(dv);int delta=cell.RowIndex-donorCell.RowIndex;
        foreach(var r in dv.Items){var s=r.SheetReference;if(s!=null&&!String.IsNullOrEmpty(s.StartSheetName))continue;var x=r.CellArea;
            r.CellArea=new CellArea(new CellReferencePosition(x.LeftColumnIndex,x.TopRowIndex+(x.TopLeft.RowType==PositionType.Relative?delta:0),x.TopLeft.ColumnType,x.TopLeft.RowType),new CellReferencePosition(x.RightColumnIndex,x.BottomRowIndex+(x.BottomRight.RowType==PositionType.Relative?delta:0),x.BottomRight.ColumnType,x.BottomRight.RowType));
        }
        string expected=Text(a,parsed.Expression);Require(expected==Text(b,dp.Expression),"TDB repair differs from Excel-inserted master mirror: "+cell.GetReferenceA1());return expected;
    }
    [STAThread] public static int Main(string[] args) {try{
        AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
        var sheets=new HashSet<string>(File.ReadAllLines(args[4]),StringComparer.OrdinalIgnoreCase);
        if(args.Length>5){
            var manifest=Json.Deserialize<Manifest>(File.ReadAllText(args[2]));
            using(var original=Load(args[1])){
                foreach(var p in manifest.differences){
                    Require(original.Worksheets[p.sheet].Cells[p.cell].FormulaInvariant==p.before,"Source formula changed: "+p.sheet+"!"+p.cell);
                    var cell=original.Worksheets[p.sheet].Cells[p.cell];
                    Require(!cell.HasDynamicArrayFormula,"Do not rewrite a dynamic array");
                    if(cell.HasArrayFormula){var area=cell.GetArrayFormulaRange();Require(area.RowCount==1&&area.ColumnCount==1,"Do not split a client multi-cell array");}
                    Prove(original,p.before,p.expected,p.sheet,p.column,sheets);
                }
                foreach(var p in manifest.nameDifferences){Require(original.DefinedNames.GetDefinedName(p.name).RefersTo==p.before,"Source name changed");Prove(original,p.before,p.expected,"",-1,sheets);}
                Console.WriteLine("PASS: "+manifest.differences.Length+" formula and "+manifest.nameDifferences.Length+" name repairs prove known column-reference signatures only.");
            }
            Preview(args[1],Path.Combine(args[3],"funding-before.png"));
            if(args[5]=="prepare")return 0;
            Apply(args,manifest,sheets);return 0;
        }
        using(var a=Load(args[1]))using(var b=Load(args[2])){
            var differences=new List<object>();var counts=new List<object>();
            foreach(var sa in a.Worksheets){
                var sb=b.Worksheets[sa.Name];int checkedCount=0,diffs=0,missing=0;
                bool linked=sheets.Contains(sa.Name);
                foreach(var c in sa.GetUsedRange().ExistingCells){
                    if(!c.HasFormula)continue;var f=c.FormulaInvariant;
                    if(!linked&&!ThreeDFunding(a,f,sheets))continue;
                    checkedCount++;var e=sb.Cells[c.RowIndex,c.ColumnIndex];var ef=sa.Name=="Transactional DB"?TdbExpected(a,b,c,sheets):e.FormulaInvariant;
                    if(String.Equals(f,ef,StringComparison.OrdinalIgnoreCase))continue;
                    diffs++;
                    differences.Add(new{sheet=sa.Name,cell=c.GetReferenceA1(),row=c.RowIndex,column=c.ColumnIndex,before=f,expected=ef,array=c.HasArrayFormula,locked=c.Protection.Locked,linked=linked});
                }
                if(linked)foreach(var c in sb.GetUsedRange().ExistingCells.Where(c=>c.HasFormula))if(!sa.Cells[c.RowIndex,c.ColumnIndex].HasFormula)missing++;
                if(checkedCount>0){Console.WriteLine(sa.Name+": checked="+checkedCount+", differences="+diffs+", missing="+missing);counts.Add(new{sheet=sa.Name,checkedCount,diffs,missing});}
            }
            var names=a.DefinedNames.Where(n=>n.Name.StartsWith("TransCopy_FacilityNames_")||n.Name.StartsWith("TransCopy_LoanDescsOrd_")||n.Name=="FacilityNames"||n.Name=="LoanDescsOrd"||n.Name=="LoanDescRev1").Select(n=>new{name=n.Name,reference=n.RefersTo,rows=n.Range.RowCount,columns=n.Range.ColumnCount}).ToArray();
            var nameDifferences=a.DefinedNames.Where(n=>ThreeDFunding(a,n.RefersTo,sheets)).Where(n=>b.DefinedNames.GetDefinedName(n.Name)!=null&&n.RefersTo!=b.DefinedNames.GetDefinedName(n.Name).RefersTo).Select(n=>new{name=n.Name,before=n.RefersTo,expected=b.DefinedNames.GetDefinedName(n.Name).RefersTo}).ToArray();
            File.WriteAllText(Path.Combine(args[3],"inventory.json"),Json.Serialize(new{counts,differences,names,nameDifferences}));
            Console.WriteLine("OUTPUT="+args[3]);
        }return 0;
    }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
    static void Apply(string[] args,Manifest manifest,HashSet<string> sheets){
        var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
        var files=app.GetType("Abovo.FileManager");files.GetMethod("Initialise").Invoke(null,new object[]{null});
        string copy=Path.Combine(args[3],"input-copy.xlsb");File.Copy(args[1],copy);
        var open=files.GetMethod("OpenModel");dynamic loaded=open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});Require(!loaded.BError,"Open failed");
        dynamic model=((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);IWorkbook w=(IWorkbook)model.WB;
        var safety=app.GetType("Abovo.ModelSafetyManager");var security=app.GetType("Abovo.WSSecurity");
        var mode=w.Options.CalculationMode;var engine=w.Options.CalculationEngineType;bool history=w.History.IsEnabled;
        var protection=w.Worksheets.ToDictionary(s=>s.Name,s=>s.IsProtected);var visibility=w.Worksheets.ToDictionary(s=>s.Name,s=>s.Visible);
        string dashboardFormula=null;
        if(args.Length>6&&args[6]=="restore-dashboard"){
            string library=Path.GetFullPath(Path.Combine(args[0],"../../Library"));
            using(var blank=Load(Path.Combine(library,"Blank BP v26_0001.xlsb")))using(var demo=Load(Path.Combine(library,"Demo BP v26_0001.xlsb"))){
                dashboardFormula=blank.Worksheets["Multivariable Dashboard"].Cells["B41"].FormulaInvariant;
                Require(dashboardFormula==demo.Worksheets["Multivariable Dashboard"].Cells["B41"].FormulaInvariant,"Master dashboard formulas differ");
                Require(dashboardFormula.IndexOf("CONCATENATE(",StringComparison.OrdinalIgnoreCase)>=0,"Expected dashboard description expression is missing");
                Require(w.Worksheets["Multivariable Dashboard"].Cells["B41"].FormulaInvariant=="=#VALUE!","Do not overwrite a client dashboard formula");
            }
            Preview(args[1],Path.Combine(args[3],"dashboard-before.png"),"Multivariable Dashboard",32,0);
        }
        safety.GetMethod("BeginBulkWorkbookMutation").Invoke(null,new object[]{0});
        try{
            w.Options.CalculationMode=WorkbookCalculationMode.Manual;w.Options.CalculationEngineType=CalculationEngineType.Recursive;w.History.IsEnabled=false;w.BeginUpdate();
            try{
                foreach(var group in manifest.differences.GroupBy(p=>p.sheet)){
                    var ws=w.Worksheets[group.Key];var permissions=ws.GetProtectionPermissions();bool protect=ws.IsProtected;
                    if(protect)security.GetMethod("UNProtectWS").Invoke(null,new object[]{0,ws.Name});
                    try{foreach(var p in group){var cell=ws.Cells[p.cell];Require(cell.FormulaInvariant==p.before,"Formula changed during load: "+p.sheet+"!"+p.cell);if(cell.HasArrayFormula)cell.GetArrayFormulaRange().ArrayFormulaInvariant=p.expected;else cell.FormulaInvariant=p.expected;}}
                    finally{if(protect)security.GetMethod("ProtectWS").Invoke(null,new object[]{0,ws.Name,permissions});}
                }
                foreach(var p in manifest.nameDifferences){Require(w.DefinedNames.GetDefinedName(p.name).RefersTo==p.before,"Name changed during load");w.DefinedNames.GetDefinedName(p.name).RefersTo=p.expected;}
                if(dashboardFormula!=null){
                    var ws=w.Worksheets["Multivariable Dashboard"];bool protect=ws.IsProtected;var permissions=ws.GetProtectionPermissions();
                    if(protect)security.GetMethod("UNProtectWS").Invoke(null,new object[]{0,ws.Name});
                    try{ws.Cells["B41"].FormulaInvariant=dashboardFormula;}finally{if(protect)security.GetMethod("ProtectWS").Invoke(null,new object[]{0,ws.Name,permissions});}
                }
            }finally{w.EndUpdate();}
            dynamic result=model.TransDBSync.SynchroniseForNamedRange("LoanDescsOrd");Require(!result.BError,"TDB sync failed: "+result.StringReturn);
            Require(w.DefinedNames.GetDefinedName("LoanDescsOrd").Range.ColumnCount==18,"Ordinary count");
            foreach(string name in new[]{"TransCopy_LoanDescsOrd_A","TransCopy_LoanDescsOrd_B"})Require(w.DefinedNames.GetDefinedName(name).Range.RowCount==19,"Mirror count "+name);
            foreach(var p in manifest.differences.Where(p=>p.sheet!="Transactional DB"))Require(w.Worksheets[p.sheet].Cells[p.cell].FormulaInvariant==p.expected,"Repair did not persist");
            model.IsDirty=true;
        }finally{
            try{w.Options.CalculationEngineType=engine;w.Options.CalculationMode=mode;w.History.IsEnabled=history;}
            finally{safety.GetMethod("EndBulkWorkbookMutation").Invoke(null,new object[]{0});}
        }
        Require(w.Worksheets.All(s=>s.IsProtected==protection[s.Name]&&s.Visible==visibility[s.Name]),"Sheet state changed");
        Require((bool)model.SaveFileAsTo(Path.Combine(args[3],"repaired.xlsb"),true),"Prepared save failed");
        if(dashboardFormula!=null){
            var c=w.Worksheets["Multivariable Dashboard"].Cells["B41"];Require(!c.Value.IsError&&c.FormulaInvariant!="=#VALUE!","Restored dashboard did not calculate");
            manifest.differences=manifest.differences.Concat(new[]{new Patch{sheet="Multivariable Dashboard",cell="B41",row=40,column=1,before="=#VALUE!",expected=c.FormulaInvariant}}).ToArray();
            Console.WriteLine("PASS: Dashboard B41 restored from identical Blank/Demo source formulas; compatibility-normalised save calculates successfully.");
        }
        File.WriteAllText(Path.Combine(args[3],"final-manifest.json"),Json.Serialize(manifest));
        Console.WriteLine("PASS: Native Funding repair and two-mirror resize saved privately.");
        var checks=w.Worksheets["Check Sheet"].GetUsedRange().ExistingCells.Where(c=>!c.Value.IsEmpty).Select(c=>new{cell=c.GetReferenceA1(),value=c.Value.ToString(),error=c.Value.IsError}).ToArray();
        File.WriteAllText(Path.Combine(args[3],"calculated-checks.json"),Json.Serialize(checks));
        Preview(Path.Combine(args[3],"repaired.xlsb"),Path.Combine(args[3],"funding-after.png"));
        if(dashboardFormula!=null)Preview(Path.Combine(args[3],"repaired.xlsb"),Path.Combine(args[3],"dashboard-after.png"),"Multivariable Dashboard",32,0);
    }
}
