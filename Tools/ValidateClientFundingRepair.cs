using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;
public static class ValidateClientFundingRepair {
 public class Patch {public string sheet;public string cell;public string expected;}
 public class NP {public string name;public string expected;}
 public class Manifest {public Patch[] differences;public NP[] nameDifferences;}
 class Refs:ExpressionVisitor {public List<CellReferenceExpression> Items=new List<CellReferenceExpression>();public override void Visit(CellReferenceExpression e){Items.Add(e);base.Visit(e);}}
 static JavaScriptSerializer Json=new JavaScriptSerializer {MaxJsonLength=int.MaxValue};
 static Workbook Load(string path){var w=new Workbook();w.Options.CalculationMode=WorkbookCalculationMode.Manual;using(var s=File.OpenRead(path))w.LoadDocument(s,DocumentFormat.Xlsb);w.Options.CalculationMode=WorkbookCalculationMode.Manual;return w;}
 static string Value(CellValue v){return v.Type+":"+v.ToString(CultureInfo.InvariantCulture);}
 static string Canonical(IWorkbook w,string f){if(String.IsNullOrEmpty(f))return f;var p=w.FormulaEngine.Parse(f);var b=new StringBuilder();p.Expression.BuildExpressionString(b,w);return b.ToString();}
 static int Map(int r){return r+(r>=1047?8:0)+(r>=1064?8:0);}
 static string Translate(IWorkbook w,string formula,string owner){
  if(String.IsNullOrEmpty(formula))return formula;
  var parsed=w.FormulaEngine.Parse(formula);var visitor=new Refs();parsed.Expression.Visit(visitor);bool changed=false;
  foreach(var r in visitor.Items){var s=r.SheetReference;
   if(s!=null&&s.StartSheetName=="Transactional DB"||(owner=="Transactional DB"&&(s==null||String.IsNullOrEmpty(s.StartSheetName)))){
    var a=r.CellArea;if(a.TopRowIndex==0&&a.BottomRowIndex==1048575)continue;
    int top=Map(a.TopRowIndex),bottom=Map(a.BottomRowIndex);if(top==a.TopRowIndex&&bottom==a.BottomRowIndex)continue;
    r.CellArea=new CellArea(new CellReferencePosition(a.LeftColumnIndex,top,a.TopLeft.ColumnType,a.TopLeft.RowType),new CellReferencePosition(a.RightColumnIndex,bottom,a.BottomRight.ColumnType,a.BottomRight.RowType));changed=true;
   }
  }
  if(!changed)return formula;var text=new StringBuilder();parsed.Expression.BuildExpressionString(text,w);return "="+text;
 }
 [STAThread] public static int Main(string[] args){try{
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  var m=Json.Deserialize<Manifest>(File.ReadAllText(args[3]));var patches=m.differences.ToDictionary(p=>p.sheet+"!"+p.cell,p=>p.expected);var names=m.nameDifferences.ToDictionary(p=>p.name,p=>p.expected);
  using(var a=Load(args[1]))using(var b=Load(args[2])){
   foreach(var book in new[]{a,b}){
    Console.WriteLine("STATUS "+(book==a?"original":"repaired"));
    foreach(string address in new[]{"A1","A37","B37","E37","A39","B39","E39","B63"}){var c=book.Worksheets["Check Sheet"].Cells[address];Console.WriteLine(address+"="+Value(c.Value)+"; formula="+c.FormulaInvariant);}
    foreach(string address in new[]{"C6","C10"}){var c=book.Worksheets["Global Assumptions"].Cells[address];Console.WriteLine("Global "+address+"="+Value(c.Value)+"; formula="+c.FormulaInvariant);}
    foreach(string address in book==a?new[]{"Q1923","Q1991"}:new[]{"Q1939","Q2007"}){var c=book.Worksheets["Transactional DB"].Cells[address];Console.WriteLine("TDB "+address+"="+Value(c.Value)+"; formula="+c.FormulaInvariant);}
    foreach(var c in book.Worksheets["Transactional DB"].GetUsedRange().ExistingCells.Where(c=>c.Value.IsError).Take(12))Console.WriteLine("TDB ERROR "+c.GetReferenceA1()+" "+Value(c.Value)+" "+c.FormulaInvariant);
   }
   int formulas=0,constants=0;var failures=new List<object>();var formulaSamples=new List<object>();var counts=new List<object>();
   Action<string,string,string,string,string> fail=(sheet,cell,kind,x,y)=>{if(failures.Count<1000)failures.Add(new{sheet,cell,kind,before=x,after=y});};
   if(!a.Worksheets.Select(s=>s.Name).SequenceEqual(b.Worksheets.Select(s=>s.Name)))throw new Exception("Sheet order changed");
   foreach(var sa in a.Worksheets){var sb=b.Worksheets[sa.Name];int fd=0,id=0,sd=0,ad=0;bool tdb=sa.Name=="Transactional DB";
    if(sa.IsProtected!=sb.IsProtected||sa.Visible!=sb.Visible)fail(sa.Name,"","sheet-state",sa.IsProtected+"/"+sa.Visible,sb.IsProtected+"/"+sb.Visible);
    foreach(var c in sa.GetUsedRange().ExistingCells){
     if(!c.HasFormula&&c.Value.IsEmpty)continue;var d=sb.Cells[tdb?Map(c.RowIndex):c.RowIndex,c.ColumnIndex];
     if(c.HasFormula){formulas++;string expected;
      if(!patches.TryGetValue(sa.Name+"!"+c.GetReferenceA1(),out expected))expected=c.FormulaInvariant;
      expected=Translate(a,expected,sa.Name);
      if(!String.Equals(expected,d.FormulaInvariant,StringComparison.OrdinalIgnoreCase)&&!String.Equals(Canonical(a,expected),Canonical(b,d.FormulaInvariant),StringComparison.OrdinalIgnoreCase)){fd++;fail(sa.Name,c.GetReferenceA1(),"formula",expected,d.FormulaInvariant);}
      if(!tdb&&c.HasArrayFormula!=d.HasArrayFormula){ad++;fail(sa.Name,c.GetReferenceA1(),"array-type",c.HasArrayFormula.ToString(),d.HasArrayFormula.ToString());}
     }else {constants++;if(d.HasFormula||Value(c.Value)!=Value(d.Value)){id++;fail(sa.Name,c.GetReferenceA1(),"constant",Value(c.Value),Value(d.Value));}}
     if(c.NumberFormat!=d.NumberFormat||c.Protection.Locked!=d.Protection.Locked){sd++;fail(sa.Name,c.GetReferenceA1(),"format/lock",c.NumberFormat+"/"+c.Protection.Locked,d.NumberFormat+"/"+d.Protection.Locked);}
    }
    if(fd+id+sd+ad>0)Console.WriteLine(sa.Name+": formulas="+fd+", constants="+id+", formatLocks="+sd+", arrays="+ad);
    counts.Add(new{sheet=sa.Name,formulas=fd,constants=id,formatLocks=sd,arrays=ad});
   }
   if(a.DefinedNames.Count!=b.DefinedNames.Count)fail("names","count","name-count",a.DefinedNames.Count.ToString(),b.DefinedNames.Count.ToString());
   foreach(var n in a.DefinedNames){
    if(n.Name.IndexOf("rejdata",StringComparison.OrdinalIgnoreCase)>=0){
     var sensitive=b.DefinedNames.GetDefinedName(n.Name);
     if(sensitive==null||n.RefersTo!=sensitive.RefersTo)fail("names","[redacted]","sensitive-name","unchanged","changed");
     continue;
    }
    string expected;
    if(!names.TryGetValue(n.Name,out expected))expected=Translate(a,n.RefersTo,"");var bn=b.DefinedNames.GetDefinedName(n.Name);
    if(bn==null||!String.Equals(expected,bn.RefersTo,StringComparison.OrdinalIgnoreCase))fail("names",n.Name,"name",expected,bn==null?"missing":bn.RefersTo);
   }
   var totals=new{formulas,constants,counts,failures};File.WriteAllText(Path.Combine(args[4],"preservation.json"),Json.Serialize(totals));
   Console.WriteLine("TOTAL originalFormulas="+formulas+", originalConstants="+constants+", failureSamples="+failures.Count);
   int outputErrors=0;
   foreach(string sn in new[]{"Detailed Comp Inc - Trad View","Financial Position - Trad View","Cashflow detailed","Check Sheet"}){
    var errors=b.Worksheets[sn].GetUsedRange().ExistingCells.Where(c=>c.Value.IsError).Select(c=>new{cell=c.GetReferenceA1(),formula=c.FormulaInvariant,value=Value(c.Value)}).ToArray();Console.WriteLine(sn+" saved errors="+errors.Length);
    File.WriteAllText(Path.Combine(args[4],sn.Replace(" ","-")+"-errors.json"),Json.Serialize(errors));
    outputErrors+=errors.Length;
   }
   if(failures.Count>0||outputErrors>0)throw new Exception("Repair validation failed; see the preservation and output-error reports.");
  }return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}
