using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Functions;

// Paired regression/profile: frozen pre-2.71 string evaluator vs production.
// No client workbook saves.
public static class MaintenanceCostProfile {
 sealed class Probe : ICustomFunction {
  readonly ICustomFunction original;readonly string name;readonly bool planned;
  public bool Direct;public long Calls,Ticks;public bool CaptureInvalidRange;
  public readonly Dictionary<int,string> Exceptions=new Dictionary<int,string>();
  public Probe(ICustomFunction fn,string alias,bool pm){original=fn;name=alias;planned=pm;}
  public string Name{get{return name;}}public string GetName(CultureInfo c){return name;}
  public DevExpress.Spreadsheet.Functions.ParameterInfo[] Parameters{get{return original.Parameters;}}
  public ParameterType ReturnType{get{return original.ReturnType;}}public bool Volatile{get{return original.Volatile;}}
  public void Reset(){Calls=Ticks=0;Exceptions.Clear();}
  public ParameterValue Evaluate(IList<ParameterValue> p,EvaluationContext c){
   long start=Stopwatch.GetTimestamp();Calls++;
   try{return Direct?original.Evaluate(p,c):EvaluateLegacy(p,c);}
   catch(ArgumentOutOfRangeException e){if(!CaptureInvalidRange)throw;Exceptions[c.Row]=e.GetType().FullName;return ParameterValue.ErrorInvalidValueInFunction;}
   finally{Ticks+=Stopwatch.GetTimestamp()-start;}
  }
  ParameterValue EvaluateLegacy(IList<ParameterValue> p,EvaluationContext c){
   int year=Convert.ToInt32(p[0].NumericValue),units=Convert.ToInt32(p[1].NumericValue),offset=planned?1:0;
   double unitCost=planned?p[2].NumericValue:1;
   int first=Convert.ToInt32(p[2+offset].NumericValue),last=Convert.ToInt32(p[3+offset].NumericValue);
   Convert.ToInt32(p[4+offset].NumericValue); // Preserve original unused-argument conversion.
   CellRange rates=p[5+offset].RangeValue,years=p[6+offset].RangeValue;
   double annual=(double)units/(last-first+1),total=0;
   var engine=c.Sheet.Workbook.FormulaEngine;
   var context=new DevExpress.Spreadsheet.Formulas.ExpressionContext(c.Column,c.Row,c.Sheet,c.Culture,ReferenceStyle.R1C1,DevExpress.Spreadsheet.Formulas.ExpressionStyle.Normal);
   for(int i=year,j=0;i>=first&&j<(last-first+1);i--,j++){
    var found=engine.Evaluate("=MATCH("+(i-first+1)+", "+years.GetReferenceR1C1(ReferenceElement.IncludeSheetName|ReferenceElement.RowAbsolute|ReferenceElement.ColumnAbsolute,null)+")",context);
    double rate=rates[(int)found.NumericValue-1].Value.NumericValue;
    // Preserve the original multiplication order for each function.
    total+=planned?unitCost*annual*rate:annual*rate;
   }
   return total;
  }
  public double Ms{get{return 1000.0*Ticks/Stopwatch.Frequency;}}
 }
 static Assembly app;static int assertions;
 static void Check(bool condition,string message){if(!condition)throw new Exception(message);assertions++;}
 static Probe Create(IWorkbook w,bool planned,string alias){
  var inner=(ICustomFunction)Activator.CreateInstance(app.GetType(planned?"Abovo.PMCostFunction":"Abovo.ResponsiveCostFunction",true));
  var p=new Probe(inner,alias??inner.Name,planned);
  if(w.Functions.GlobalCustomFunctions.Contains(p.Name))w.Functions.GlobalCustomFunctions.Remove(p.Name);
  w.Functions.GlobalCustomFunctions.Add(p);return p;
 }
 [STAThread]public static int Main(string[] args){try{
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string f=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");return File.Exists(f)?Assembly.LoadFrom(f):null;};
  Run(args);Console.WriteLine("PASS: "+assertions+" legacy/production profile checks");return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
 static void Run(string[] args){
  app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
  Synthetic();if(args.Length>1)Client(args[1]);
 }
 static void Synthetic(){
  using(var w=new Workbook()){
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;
   var bp=Create(w,true,"BASEPM");var fp=Create(w,true,"FASTPM");fp.Direct=true;
   var br=Create(w,false,"BASERESP");var fr=Create(w,false,"FASTRESP");fr.Direct=true;
   // The old implementation threw for missing MATCH. Production must return
   // #N/A instead; valid inputs must retain exactly the same numeric results.
   foreach(var p in new[]{bp,fp,br,fr})p.CaptureInvalidRange=true;
   var sheet=w.Worksheets[0];sheet.Name="Results";var rates=w.Worksheets.Add("Rates");
   for(int i=0;i<40;i++){rates.Cells[i,0].Value=(i+1)*0.02;rates.Cells[i,1].Value=i+1;}
   int row=0;
   // Exact/approximate lookups, duplicate keys, missing and unsorted keys,
   // sparse year bands and empty loops.
   string[] yearVectors={"1,2,3,4,5,6","1,3,5,7,9,11","1,1,2,2,4,4","3,5,7,9,11,13","3,1,4,2,6,5"};
   foreach(string vector in yearVectors){
    string[] v=vector.Split(',');int col=3+Array.IndexOf(yearVectors,vector)*3;
    for(int i=0;i<6;i++){rates.Cells[i,col].Value=(i+1)*0.1;rates.Cells[i,col+1].Value=Int32.Parse(v[i]);}
    for(int y=0;y<15;y++)for(int duration=1;duration<=5;duration++){
     string r=rates.Range.FromLTRB(col,0,col,5).GetReferenceA1(ReferenceElement.IncludeSheetName|ReferenceElement.RowAbsolute|ReferenceElement.ColumnAbsolute);
     string yr=rates.Range.FromLTRB(col+1,0,col+1,5).GetReferenceA1(ReferenceElement.IncludeSheetName|ReferenceElement.RowAbsolute|ReferenceElement.ColumnAbsolute);
     Pair(sheet,row++,y,30,2.5,1,duration,r,yr);
    }
   }
   // Bulk workload uses ordinary rate tables; errors above are deliberately compared.
   for(int i=0;i<2000;i++)Pair(sheet,row++,1+i%40,50+i%100,7.25,1,5+i%20,"Rates!$A$1:$A$40","Rates!$B$1:$B$40");
   var prior=w.Options.CalculationEngineType;
   try{
    w.Options.CalculationEngineType=CalculationEngineType.Recursive;
    for(int run=0;run<4;run++){
     foreach(var p in new[]{bp,fp,br,fr})p.Reset();
     w.CalculateFullRebuild();
     for(int r=0;r<row;r++){
      if(bp.Exceptions.ContainsKey(r))Check(sheet.Cells[r,1].Value.ToString()=="#N/A","Missing PM lookup returns #N/A, not an exception");else Compare(sheet.Cells[r,0],sheet.Cells[r,1]);
      if(br.Exceptions.ContainsKey(r))Check(sheet.Cells[r,3].Value.ToString()=="#N/A","Missing Resp lookup returns #N/A, not an exception");else Compare(sheet.Cells[r,2],sheet.Cells[r,3]);
     }
     Check(fp.Exceptions.Count==0&&fr.Exceptions.Count==0,"Production lookup evaluations never throw range exceptions");
     Console.WriteLine("SYNTHETIC "+(run==0?"warmup":"run "+run)+": pairs/function="+row+", PM old="+bp.Ms.ToString("F2",CultureInfo.InvariantCulture)+" ms direct="+fp.Ms.ToString("F2",CultureInfo.InvariantCulture)+" ms; Resp old="+br.Ms.ToString("F2",CultureInfo.InvariantCulture)+" ms direct="+fr.Ms.ToString("F2",CultureInfo.InvariantCulture)+" ms");
    }
   }finally{w.Options.CalculationEngineType=prior;}
  }
 }
 static void Pair(Worksheet sheet,int row,int year,int units,double cost,int first,int last,string rates,string years){
  string prefix=year+","+units+",",tail=first+","+last+",999,"+rates+","+years+")";
  sheet.Cells[row,0].FormulaInvariant="=BASEPM("+prefix+cost.ToString(CultureInfo.InvariantCulture)+","+tail;
  sheet.Cells[row,1].FormulaInvariant="=FASTPM("+prefix+cost.ToString(CultureInfo.InvariantCulture)+","+tail;
  sheet.Cells[row,2].FormulaInvariant="=BASERESP("+prefix+tail;
  sheet.Cells[row,3].FormulaInvariant="=FASTRESP("+prefix+tail;
 }
 static void Compare(Cell a,Cell b){
  Check(a.Value.Type==b.Value.Type,"Type parity "+a.GetReferenceA1());
  if(a.Value.IsNumeric)Check(a.Value.NumericValue==b.Value.NumericValue,"Exact numeric parity "+a.GetReferenceA1()+": "+a.Value+" vs "+b.Value);
  else Check(a.Value.ToString(CultureInfo.InvariantCulture)==b.Value.ToString(CultureInfo.InvariantCulture),"Error/text parity "+a.GetReferenceA1());
 }
 static string Digest(IWorkbook w){
  using(var hash=SHA256.Create())using(var sink=new CryptoStream(Stream.Null,hash,CryptoStreamMode.Write))using(var writer=new BinaryWriter(sink,Encoding.UTF8)){
   foreach(var sheet in w.Worksheets){writer.Write(sheet.Name);foreach(var cell in sheet.GetUsedRange().ExistingCells){
    if(!cell.HasFormula&&cell.Value.IsEmpty)continue;
    writer.Write(cell.RowIndex);writer.Write(cell.ColumnIndex);writer.Write(cell.HasFormula?cell.FormulaInvariant:"");
    writer.Write((int)cell.Value.Type);if(cell.Value.IsNumeric)writer.Write(cell.Value.NumericValue);else writer.Write(cell.Value.ToString(CultureInfo.InvariantCulture));
   }}
   writer.Flush();sink.FlushFinalBlock();return BitConverter.ToString(hash.Hash).Replace("-","");
  }
 }
 static void Client(string path){
  using(var w=new Workbook()){
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;var pm=Create(w,true,null);var resp=Create(w,false,null);
   Console.WriteLine("CLIENT loading read-only evidence into memory");var timer=Stopwatch.StartNew();
   using(var stream=File.OpenRead(path))Check(w.LoadDocument(stream,Path.GetExtension(path).Equals(".xlsm",StringComparison.OrdinalIgnoreCase)?DocumentFormat.Xlsm:DocumentFormat.Xlsb),"Client load");
   w.Options.CalculationMode=WorkbookCalculationMode.Manual;Console.WriteLine("CLIENT loaded ms="+timer.ElapsedMilliseconds);
   var engine=w.Options.CalculationEngineType;string baseline=null;
   try{w.Options.CalculationEngineType=CalculationEngineType.Recursive;
    foreach(bool direct in new[]{false,true,true,false}){
     pm.Direct=resp.Direct=direct;pm.Reset();resp.Reset();timer.Restart();w.CalculateFullRebuild();long elapsed=timer.ElapsedMilliseconds;
     string digest=Digest(w);if(baseline==null)baseline=digest;else Check(baseline==digest,"All worksheet formulas and calculated values identical across baseline/direct rebuilds");
     Console.WriteLine("CLIENT "+(direct?"direct":"baseline")+": rebuild="+elapsed+" ms; PM calls="+pm.Calls+" time="+pm.Ms.ToString("F2",CultureInfo.InvariantCulture)+" ms; Resp calls="+resp.Calls+" time="+resp.Ms.ToString("F2",CultureInfo.InvariantCulture)+" ms; digest="+digest);
    }
   }finally{w.Options.CalculationEngineType=engine;}
   Console.WriteLine("PASS: full-workbook result digests match; client file not saved");
  }
 }
}
