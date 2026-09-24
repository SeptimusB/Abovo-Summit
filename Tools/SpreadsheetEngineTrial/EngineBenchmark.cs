using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using AC = Aspose.Cells;
using DX = DevExpress.Spreadsheet;
using Gear = SpreadsheetGear;

// Research harness only: private paths, no original saves, no Summit UI or security changes.
internal static partial class EngineBenchmark
{
    const int Rows = 4096;
    internal static string Repo { get { return Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"../../../../../..")); } }
    internal static string Address(int row, int col, int height = 1, int width = 1)
    {
        return Column(col) + (row + 1) + ":" + Column(col + width - 1) + (row + height);
    }
    static string Column(int c) { string s=""; for(c++;c>0;c=(c-1)/26) s=(char)('A'+(c-1)%26)+s; return s; }
    static double Time(Action action) { var t=Stopwatch.StartNew(); action(); return t.Elapsed.TotalMilliseconds; }
    static void Check(bool yes, string text) { if(!yes) throw new InvalidOperationException(text); }
    static bool Equal(double a,double b) { return Math.Abs(a-b) <= Math.Max(0.000001, Math.Abs(b)*1e-10); }
    static object[][] Jagged(object[,] a) { return Enumerable.Range(0,a.GetLength(0)).Select(r=>Enumerable.Range(0,a.GetLength(1)).Select(c=>a[r,c]).ToArray()).ToArray(); }
    internal static object Error(string error)
    {
        switch(error) { case "Div0": return "#DIV/0!"; case "Value": return "#VALUE!"; case "Ref": return "#REF!"; case "Name":return "#NAME?"; case "Num":return "#NUM!"; case "NA":return "#N/A"; case "Null":return "#NULL!"; default:return error; }
    }
    public static int Run(string[] args)
    {
        CultureInfo.CurrentCulture=CultureInfo.InvariantCulture;
        if(args.Length==4 && args[1]=="udf-excel")
        {
            string udfInput=IoTrial.InTrial(args[2]),report=IoTrial.InTrial(args[3]);
            Check(!File.Exists(report),"New report required.");string hash=IoTrial.FileHash(udfInput);
            object result;
            using(var engine=new ExcelEngine(true)){engine.Open(udfInput);result=engine.VerifyWorkbookFunctions();}
            Check(hash==IoTrial.FileHash(udfInput),"Private source changed.");
            IoTrial.WriteJson(report,new {success=true,sourceUnchanged=true,inputHash=hash,functions=result});
            Console.WriteLine("PASS Excel VBA known-answer probes.");return 0;
        }
        if(args.Length==3 && args[1]=="fixture") { Fixture(IoTrial.InTrial(args[2])); return 0; }
        if(args.Length==5 && args[1]=="convert")
        {
            string source=IoTrial.InTrial(args[2]), target=IoTrial.InTrial(args[3]), report=IoTrial.InTrial(args[4]);
            Check(!File.Exists(target)&&!File.Exists(report),"New output paths required.");
            string hash=IoTrial.FileHash(source);
            using(var engine=new ExcelEngine(false))
            {
                double open=Time(()=>engine.Open(source));
                double save=Time(()=>engine.SaveCopyAsXlsm(target));
                Check(hash==IoTrial.FileHash(source),"Private source changed.");
                IoTrial.WriteJson(report,new { source,target,sourceHash=hash,sourceUnchanged=true,openMs=open,saveMs=save,engine.Version,macrosEnabled=false,eventsEnabled=false,calculationRequested=false });
            }
            return 0;
        }
        if(args.Length!=5) throw new ArgumentException("bench fixture PATH; bench convert XLSB XLSM REPORT; bench ENGINE INPUT REPORT synthetic|agl");
        string input=IoTrial.InTrial(args[2]), resultPath=IoTrial.InTrial(args[3]);
        bool synthetic=args[4]=="synthetic";
        Check(synthetic || args[4]=="agl","Choose synthetic or agl.");
        Check(!File.Exists(resultPath),"Report must be new.");
        string before=IoTrial.FileHash(input);
        var reportData=new Dictionary<string,object> { {"engine",args[1]}, {"input",input}, {"inputHash",before}, {"kind",args[4]}, {"utc",DateTime.UtcNow.ToString("o")}, {"bits",IntPtr.Size*8}, {"scope","Native engine only; no Summit UI, save preflight, transaction or rebind overhead"} };
        var times=new Dictionary<string,object>(); reportData["timingsMs"]=times;
        try
        {
            using(var engine=Create(args[1],!synthetic))
            {
                Console.WriteLine("START "+args[1]+" "+args[4]+" load");
                times["load"]=Time(()=>engine.Open(input)); reportData["version"]=engine.Version;
                try
                {
                    if(synthetic) Synthetic(engine,times,reportData);
                    else Model(engine,times,reportData);
                }
                finally { reportData["engineDiagnostics"]=engine.Diagnostics; }
            }
            reportData["success"]=true;
        }
        catch(Exception error) { reportData["success"]=false; reportData["error"]=error.GetType().Name+": "+error.Message; reportData["errorStack"]=error.StackTrace; Console.WriteLine("FAILED "+reportData["error"]); }
        finally
        {
            reportData["sourceUnchanged"]=before==IoTrial.FileHash(input);
            reportData["peakWorkingSetBytes"]=Process.GetCurrentProcess().PeakWorkingSet64;
            IoTrial.WriteJson(resultPath,reportData);
            Console.WriteLine("RESULT "+resultPath);
            Check((bool)reportData["sourceUnchanged"],"Private input changed.");
        }
        return (bool)reportData["success"] ? 0 : 1;
    }
    static Engine Create(string kind,bool model)
    {
#if SPREAD_TRIAL
        if(kind=="spreadnet")return new SpreadEngine(model);
#endif
        switch(kind) { case "devexpress":return new DxEngine(model);case "aspose":return new AsposeEngine(model);case "spreadsheetgear":return new GearEngine(model);case "excel":return new ExcelEngine(model);default:throw new ArgumentException("Unknown engine."); }
    }
    static void Fixture(string path)
    {
        Check(!File.Exists(path),"Fixture must be new.");
        using(var book=new DX.Workbook())
        {
            book.Options.CalculationMode=DX.WorkbookCalculationMode.Manual;
            var data=book.Worksheets[0];data.Name="Data";
            var rates=book.Worksheets.Add("Rates");var summary=book.Worksheets.Add("Summary");book.Worksheets.Add("Bulk");
            for(int r=0;r<40;r++) { rates.Cells[r,0].Value=r+1; rates.Cells[r,1].Value=(r+1)*2; }
            rates.Cells[0,3].Value=1.25;
            for(int r=0;r<Rows;r++)
            {
                int n=r+1;
                data.Cells[r,0].Value=n;data.Cells[r,1].Value=100+r%97;data.Cells[r,2].Value=0.125*(r%6+1);data.Cells[r,3].Value=1+r%40;
                string[] f={"=B"+n+"*C"+n,"=ROUND(E"+n+"*Rates!$D$1,2)","=IF(D"+n+">20,F"+n+",-F"+n+")","=INDEX(Rates!$B$1:$B$40,MATCH(D"+n+",Rates!$A$1:$A$40,0))","=G"+n+"+H"+n,"=SUM(E"+n+":I"+n+")"};
                for(int c=0;c<f.Length;c++)data.Cells[r,c+4].FormulaInvariant=f[c];
            }
            book.DefinedNames.Add("BenchInputs","Data!$A$1:$D$"+Rows);
            summary.Cells[0,1].FormulaInvariant="=SUM(Data!F1:F"+Rows+")";
            summary.Cells[1,1].FormulaInvariant="=SUM(Data!G1:G"+Rows+")";
            summary.Cells[2,1].FormulaInvariant="=SUM(Data!J1:J"+Rows+")";
            book.Options.CalculationEngineType=DX.CalculationEngineType.Recursive;book.CalculateFullRebuild();
            using(var stream=new FileStream(path,FileMode.CreateNew,FileAccess.Write))book.SaveDocument(stream,DX.DocumentFormat.Xlsm);
        }
        Console.WriteLine("Created macro-free fixture: "+Rows+" rows; 24,579 formulas.");
    }
    static void Synthetic(Engine engine,Dictionary<string,object> times,Dictionary<string,object> result)
    {
        times["firstFullRebuild"]=Time(()=>engine.Calculate(2));
        var full=new List<double>();for(int i=0;i<3;i++)full.Add(Time(()=>engine.Calculate(1))); times["warmFull"]=full;
        ValidateSynthetic(engine,1.25);
        engine.Calculate(0); // Establish a dirty-dependency chain before measuring edits.
        var incremental=new List<double>();
        foreach(double multiplier in new[]{1.5,1.75,1.25})
        {
            engine.Set("Rates",0,3,new object[,]{{multiplier}});
            incremental.Add(Time(()=>engine.Calculate(0)));ValidateSynthetic(engine,multiplier);
        }
        times["incrementalCalculation"]=incremental;
        var bulk=new object[10000,20];for(int r=0;r<10000;r++)for(int c=0;c<20;c++)bulk[r,c]=(double)(r*20+c+1);
        times["bulkWrite200000"]=Time(()=>engine.Set("Bulk",0,0,bulk));object[,] read=null;
        times["bulkRead200000"]=Time(()=>read=engine.Get("Bulk",0,0,10000,20));
        Check(Equal(read.Cast<object>().Sum(Convert.ToDouble),20000100000d),"Bulk values checksum differs.");
        string name=engine.NameReference("BenchInputs");
        times["insert10Columns"]=Time(()=>engine.Columns("Data",3,10,true));
        Check(engine.NameReference("BenchInputs")!=name,"Column insertion failed to adjust the name.");
        times["copyColumnInto10"]=Time(()=>engine.Copy("Data",0,2,Rows,1,0,3,Rows,10));
        var copiedColumns=engine.Get("Data",0,3,Rows,10);
        for(int r=0;r<Rows;r++)for(int c=0;c<10;c++)
            Check(Equal(Convert.ToDouble(copiedColumns[r,c]),0.125*(r%6+1)),"Inserted column copy differs.");
        times["delete10Columns"]=Time(()=>engine.Columns("Data",3,10,false));
        Check(engine.NameReference("BenchInputs")==name,"Column deletion failed to restore the name.");
        times["insert10Rows"]=Time(()=>engine.Rows("Data",99,10,true));
        times["copyRowInto10"]=Time(()=>engine.Copy("Data",98,0,1,10,99,0,10,10));
        var copiedRows=engine.Get("Data",99,0,10,4);var copiedInputs=new double[]{99,101,0.375,19};
        for(int r=0;r<10;r++)for(int c=0;c<4;c++)
            Check(Equal(Convert.ToDouble(copiedRows[r,c]),copiedInputs[c]),"Inserted row copy differs.");
        times["delete10Rows"]=Time(()=>engine.Rows("Data",99,10,false));
        Check(engine.NameReference("BenchInputs")==name,"Row deletion failed to restore the name.");
        engine.Calculate(2);ValidateSynthetic(engine,1.25);
        result["knownAnswersPassed"]=true;result["formulaCellsCheckedPerState"]=Rows*6;
        result["statesChecked"]=5; result["structuralChecksPassed"]=true;
        result["copiedPayloadChecksPassed"]=true;
    }
    static void ValidateSynthetic(Engine engine,double multiplier)
    {
        var values=engine.Get("Data",0,4,Rows,6);var sums=new double[3];
        for(int r=0;r<Rows;r++)
        {
            double e=(100+r%97)*0.125*(r%6+1),f=Math.Round(e*multiplier,2,MidpointRounding.AwayFromZero),g=r%40+1>20?f:-f,h=(r%40+1)*2,j=e+f+g+h+g+h;
            var expected=new[]{e,f,g,h,g+h,j};
            for(int c=0;c<6;c++)Check(Equal(Convert.ToDouble(values[r,c]),expected[c]),"Known answer differs at Data!"+Address(r,c+4));
            sums[0]+=f;sums[1]+=g;sums[2]+=j;
        }
        var actual=engine.Get("Summary",0,1,3,1);for(int i=0;i<3;i++)Check(Equal(Convert.ToDouble(actual[i,0]),sums[i]),"Summary mismatch.");
    }
    static void Model(Engine engine,Dictionary<string,object> times,Dictionary<string,object> result)
    {
        Console.WriteLine("START full model calculation");
        times["firstFullRebuild"]=Time(()=>engine.Calculate(2));
        var full=new List<double>();for(int i=0;i<3;i++) { Console.WriteLine("START warm full "+(i+1));full.Add(Time(()=>engine.Calculate(1))); }times["warmFull"]=full;
        var probes=new Dictionary<string,object>();
        var originalValues=new Dictionary<string,object[,]>();
        foreach(var sheet in new[]{"Detailed Comp Inc - Trad View","Financial Position - Trad View","Cashflow detailed","Check Sheet","Development Expenditure"})
        {
            var size=engine.Size(sheet);
            Check(size.Item1<=5000 && size.Item2<=300,"Unexpected model probe extent.");
            var values=engine.Get(sheet,0,0,size.Item1,size.Item2);originalValues[sheet]=values;
            probes[sheet]=new { rows=size.Item1,columns=size.Item2,values=Jagged(values) };
        }
        result["calculatedProbes"]=probes;
        // Unsaved source-sheet primitives, not full Summit/VBA multi-sheet structural commands.
        var mutations=new List<object>();
        foreach(var item in new[]{new[]{"Funding Assumptions","LoanDescRev1"},new[]{"Development BP Assumptions","LastIDColNum"}})
        {
            string sheet=item[0];var anchor=engine.NameLocation(item[1]);int at=anchor.Item3+(item[1]=="LoanDescRev1"?-1:0);
            Check(anchor.Item1==sheet&&at>0,"Unexpected insertion anchor.");
            bool protectedOnEntry=engine.Protected(sheet);
            string reference=engine.NameReference(item[1]);var size=engine.Size(sheet);
            // Common extents verified in the identical Excel-converted AGL fixture.
            // Aspose's MaxDataRow omits trailing formatted rows; copy those too.
            int copyRows=sheet=="Funding Assumptions"?580:688;
            Check(size.Item1<=copyRows,"AGL geometry differs from this benchmark fixture.");
            double insert,copy,delete;object fill=null;
            Console.WriteLine("START disposable source-sheet insert/copy/delete: "+sheet);
            using(protectedOnEntry?engine.TemporarilyUnprotect(sheet,TrialProtection.ReadExistingCredential()):null)
            {
                insert=Time(()=>engine.Columns(sheet,at,10,true));
                int sourceCol=item[1]=="LoanDescRev1"?at+10:at-1;
                copy=Time(()=>engine.Copy(sheet,0,sourceCol,copyRows,1,0,at,copyRows,10));
                fill=engine.AutoFillProbe(sheet,sourceCol,at,copyRows,10);
                delete=Time(()=>engine.Columns(sheet,at,10,false));
            }
            Check(engine.NameReference(item[1])==reference,"Insertion/deletion did not restore anchor.");
            Check(engine.Protected(sheet)==protectedOnEntry,"Entry protection not restored.");
            mutations.Add(new {sheet,insertMs=insert,copyMs=copy,autoFill=fill,deleteMs=delete,rows=copyRows,columns=10,anchorRestored=true,entryProtectionRestored=true});
        }
        result["sourceSheetMicrobenchmarks"]=mutations;
        Console.WriteLine("START post-delete calculation and output checks");
        times["postDeleteValidationCalculation"]=Time(()=>engine.Calculate(2));
        foreach(var pair in originalValues)
        {
            var expected=pair.Value;var actual=engine.Get(pair.Key,0,0,expected.GetLength(0),expected.GetLength(1));
            for(int r=0;r<expected.GetLength(0);r++)for(int c=0;c<expected.GetLength(1);c++)
            {
                object a=expected[r,c],b=actual[r,c];
                if(Object.Equals(a,""))a=null;if(Object.Equals(b,""))b=null;
                bool same=Object.Equals(a,b);
                if(a is IConvertible&&b is IConvertible&&!(a is string)&&!(b is string)&&!(a is bool)&&!(b is bool))
                    same=Equal(Convert.ToDouble(a),Convert.ToDouble(b));
                Check(same,"Post-delete output differs: "+pair.Key+"!"+Address(r,c));
            }
        }
        result["postDeleteOutputProbesMatch"]=true;
        result["fullStructuralCommandValidated"]=false; result["saved"]=false;
    }

    internal abstract class Engine:IDisposable
    {
        internal abstract string Version {get;}
        internal virtual object Diagnostics {get{return new { };}}
        internal abstract void Open(string path);
        internal abstract void Calculate(int mode);
        internal abstract object[,] Get(string sheet,int row,int col,int rows,int cols);
        internal abstract void Set(string sheet,int row,int col,object[,] values);
        internal abstract void Columns(string sheet,int at,int count,bool insert);
        internal abstract void Rows(string sheet,int at,int count,bool insert);
        internal abstract void Copy(string sheet,int row,int col,int rows,int cols,int destRow,int destCol,int destRows,int destCols);
        internal abstract string NameReference(string name);
        internal abstract Tuple<string,int,int> NameLocation(string name);
        internal abstract Tuple<int,int> Size(string sheet);
        internal abstract bool Protected(string sheet);
        internal abstract IDisposable TemporarilyUnprotect(string sheet,string password);
        internal virtual object AutoFillProbe(string sheet,int sourceCol,int destCol,int rows,int cols){return null;}
        public abstract void Dispose();
    }
    sealed partial class DxEngine:Engine
    {
        DX.Workbook book=new DX.Workbook(); readonly bool model;
        internal DxEngine(bool model) { this.model=model;book.Options.CalculationMode=DX.WorkbookCalculationMode.Manual; }
        internal override string Version {get{return typeof(DX.Workbook).Assembly.GetName().Version.ToString();}}
        internal override void Open(string path)
        {
            if(model)
            {
                string bin=Path.Combine(Repo,"bin/Release");
                AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string f=Path.Combine(bin,new AssemblyName(e.Name).Name+".dll");return File.Exists(f)?Assembly.LoadFrom(f):null;};
                var app=Assembly.LoadFrom(Path.Combine(bin,"Abovo-summit.exe"));
                foreach(string name in new[]{"Abovo.PMCostFunction","Abovo.ResponsiveCostFunction"})
                {
                    var f=(DX.Functions.ICustomFunction)Activator.CreateInstance(app.GetType(name,true));
                    if(book.Functions.GlobalCustomFunctions.Contains(f.Name))book.Functions.GlobalCustomFunctions.Remove(f.Name);
                    book.Functions.GlobalCustomFunctions.Add(f);
                }
            }
            using(var stream=File.OpenRead(path))Check(book.LoadDocument(stream,Path.GetExtension(path)==".xlsb"?DX.DocumentFormat.Xlsb:DX.DocumentFormat.Xlsm),"DX load failed.");
            book.Options.CalculationMode=DX.WorkbookCalculationMode.Manual;book.Options.CalculationEngineType=DX.CalculationEngineType.Recursive;
        }
        internal override void Calculate(int mode) {if(mode==2)book.CalculateFullRebuild();else if(mode==1)book.CalculateFull();else book.Calculate();}
        internal override object[,] Get(string sheet,int row,int col,int rows,int cols)
        {var a=new object[rows,cols];var ws=book.Worksheets[sheet];for(int r=0;r<rows;r++)for(int c=0;c<cols;c++){var v=ws.Cells[row+r,col+c].Value;a[r,c]=v.IsEmpty?null:v.IsNumeric?(object)v.NumericValue:v.IsBoolean?(object)v.BooleanValue:v.ToString(CultureInfo.InvariantCulture);}return a;}
        internal override void Set(string sheet,int row,int col,object[,] values) {DX.WorksheetExtensions.Import(book.Worksheets[sheet],values,row,col);}
        internal override void Columns(string sheet,int at,int count,bool insert) {var c=book.Worksheets[sheet].Columns;if(insert)c.Insert(at,count);else c.Remove(at,count);}
        internal override void Rows(string sheet,int at,int count,bool insert) {var r=book.Worksheets[sheet].Rows;if(insert)r.Insert(at,count);else r.Remove(at,count);}
        internal override void Copy(string sheet,int row,int col,int rows,int cols,int dr,int dc,int dh,int dw) {var ws=book.Worksheets[sheet];ws.Range[Address(dr,dc,dh,dw)].CopyFrom(ws.Range[Address(row,col,rows,cols)],DX.PasteSpecial.All);}
        internal override string NameReference(string name) {return book.DefinedNames.GetDefinedName(name).RefersTo;}
        internal override Tuple<string,int,int> NameLocation(string name) {var r=book.DefinedNames.GetDefinedName(name).Range;return Tuple.Create(r.Worksheet.Name,r.TopRowIndex,r.LeftColumnIndex);}
        internal override Tuple<int,int> Size(string sheet) {var r=book.Worksheets[sheet].GetUsedRange();return Tuple.Create(r.BottomRowIndex+1,r.RightColumnIndex+1);}
        internal override bool Protected(string sheet) {return book.Worksheets[sheet].IsProtected;}
        internal override IDisposable TemporarilyUnprotect(string sheet,string password)
        {
            var ws=book.Worksheets[sheet];var permissions=ws.GetProtectionPermissions();
            Check(ws.Unprotect(password)&&!ws.IsProtected,"Temporary unprotect failed.");
            return new TrialProtection(()=>{ws.Protect(password,permissions);Check(ws.IsProtected&&ws.GetProtectionPermissions()==permissions,"Protection restore failed.");});
        }
        public override void Dispose(){book.Dispose();}
    }
    sealed class AsposeEngine:Engine
    {
        AC.Workbook book;readonly bool model; readonly TrialFunctions.AsposeFunctions udf=new TrialFunctions.AsposeFunctions();
        internal AsposeEngine(bool model){this.model=model;using(var test=new AC.Workbook())Check(test.IsLicensed,"Aspose licence required.");}
        internal override string Version {get{return AC.CellsHelper.GetVersion();}}
        internal override object Diagnostics {get{return new {udf.Calls,udf.InvalidCalls,incrementalCalculation="Aspose CalculateFormula; no claim of equivalent dirty-chain execution"};}}
        internal override void Open(string path){using(var stream=File.OpenRead(path))book=new AC.Workbook(stream);Check(!book.Settings.Date1904,"Trial date normalization expects the 1900 date system.");book.Settings.FormulaSettings.CalculationMode=AC.CalcModeType.Manual;book.Settings.FormulaSettings.EnableCalculationChain=true;}
        internal override void Calculate(int mode){book.Settings.FormulaSettings.EnableCalculationChain=mode==0;book.CalculateFormula(new AC.CalculationOptions{CustomEngine=model?udf:null,IgnoreError=false,Recursive=true});}
        internal override object[,] Get(string sheet,int row,int col,int rows,int cols)
        {
            var values=book.Worksheets[sheet].Cells.ExportArray(row,col,rows,cols);
            // ExportArray returns formatted dates as DateTime; compare Excel's numeric values,
            // not JSON's UTC conversion of a local midnight. Benchmark inputs use the 1900 system.
            for(int r=0;r<rows;r++)for(int c=0;c<cols;c++)
                if(values[r,c] is DateTime)values[r,c]=((DateTime)values[r,c]).ToOADate();
            return values;
        }
        internal override void Set(string sheet,int row,int col,object[,] a){book.Worksheets[sheet].Cells.ImportTwoDimensionArray(a,row,col);}
        internal override void Columns(string sheet,int at,int count,bool insert){var c=book.Worksheets[sheet].Cells;if(insert)c.InsertColumns(at,count,true);else c.DeleteColumns(at,count,true);}
        internal override void Rows(string sheet,int at,int count,bool insert){var c=book.Worksheets[sheet].Cells;if(insert)c.InsertRows(at,count,true);else c.DeleteRows(at,count,true);}
        internal override void Copy(string sheet,int row,int col,int rows,int cols,int dr,int dc,int dh,int dw)
        {var cells=book.Worksheets[sheet].Cells;for(int r=0;r<dh;r+=rows)for(int c=0;c<dw;c+=cols)cells.CreateRange(dr+r,dc+c,rows,cols).Copy(cells.CreateRange(row,col,rows,cols));}
        internal override string NameReference(string name){return book.Worksheets.Names[name].RefersTo;}
        internal override Tuple<string,int,int> NameLocation(string name){var r=book.Worksheets.Names[name].GetRange();return Tuple.Create(r.Worksheet.Name,r.FirstRow,r.FirstColumn);}
        internal override Tuple<int,int> Size(string sheet){var c=book.Worksheets[sheet].Cells;return Tuple.Create(c.MaxDataRow+1,c.MaxDataColumn+1);}
        internal override bool Protected(string sheet){return book.Worksheets[sheet].IsProtected;}
        internal override IDisposable TemporarilyUnprotect(string sheet,string password)
        {
            var ws=book.Worksheets[sheet];var snapshot=new AC.Workbook();
            try {snapshot.Worksheets[0].Protection.Copy(ws.Protection);ws.Unprotect(password);Check(!ws.IsProtected,"Temporary unprotect failed.");}
            catch {snapshot.Dispose();throw;}
            return new TrialProtection(()=>{try{ws.Protection.Copy(snapshot.Worksheets[0].Protection);Check(ws.IsProtected,"Protection restore failed.");}finally{snapshot.Dispose();}});
        }
        public override void Dispose(){if(book!=null)book.Dispose();}
    }
    sealed partial class GearEngine:Engine
    {
        Gear.IWorkbookSet set;Gear.IWorkbook book;readonly TrialFunctions.GearFunction pm=new TrialFunctions.GearFunction(true),resp=new TrialFunctions.GearFunction(false);
        internal GearEngine(bool model){SpreadsheetGearTrial.Activate();set=Gear.Factory.GetWorkbookSet(CultureInfo.InvariantCulture);set.Calculation=Gear.Calculation.Manual;set.BackgroundCalculation=false;set.CalculationOnDemand=false;set.EnableWebService=false;if(model){set.Add(pm);set.Add(resp);}}
        internal override string Version {get{return typeof(Gear.Factory).Assembly.GetName().Version.ToString();}}
        internal override object Diagnostics {get{return new {pmCalls=pm.Calls,respCalls=resp.Calls,invalidCalls=pm.InvalidCalls+resp.InvalidCalls};}}
        internal override void Open(string path){using(var stream=File.OpenRead(path))book=set.Workbooks.OpenFromStream(stream);set.Calculation=Gear.Calculation.Manual;}
        internal override void Calculate(int mode){if(mode==2)set.CalculateFullRebuild();else if(mode==1)set.CalculateFull();else set.Calculate();}
        internal override object[,] Get(string sheet,int row,int col,int rows,int cols)
        {object raw=book.Worksheets[sheet].Cells[Address(row,col,rows,cols)].Value;var a=raw as object[,]??new object[,]{{raw}};for(int r=0;r<rows;r++)for(int c=0;c<cols;c++)if(a[r,c] is Gear.ValueError)a[r,c]=Error(a[r,c].ToString());return a;}
        internal override void Set(string sheet,int row,int col,object[,] a){book.Worksheets[sheet].Cells[Address(row,col,a.GetLength(0),a.GetLength(1))].Value=a;}
        internal override void Columns(string sheet,int at,int count,bool insert){var r=book.Worksheets[sheet].Cells[Column(at)+":"+Column(at+count-1)];if(insert)r.Insert();else r.Delete();}
        internal override void Rows(string sheet,int at,int count,bool insert){var r=book.Worksheets[sheet].Cells[(at+1)+":"+(at+count)];if(insert)r.Insert();else r.Delete();}
        internal override void Copy(string sheet,int row,int col,int rows,int cols,int dr,int dc,int dh,int dw){var c=book.Worksheets[sheet].Cells;c[Address(row,col,rows,cols)].Copy(c[Address(dr,dc,dh,dw)]);}
        internal override string NameReference(string name){return book.Names[name].RefersTo;}
        internal override Tuple<string,int,int> NameLocation(string name){var r=book.Names[name].RefersToRange;return Tuple.Create(r.Worksheet.Name,r.Row,r.Column);}
        internal override Tuple<int,int> Size(string sheet){var r=book.Worksheets[sheet].UsedRange;return Tuple.Create(r.Row+r.RowCount,r.Column+r.ColumnCount);}
        internal override bool Protected(string sheet){return book.Worksheets[sheet].ProtectContents;}
        internal override IDisposable TemporarilyUnprotect(string sheet,string password)
        {
            var ws=book.Worksheets[sheet];var p=ws.Protection;var selection=ws.EnableSelection;
            var a=new[]{ws.ProtectDrawingObjects,ws.ProtectContents,ws.ProtectScenarios,ws.ProtectionMode,p.AllowFormattingCells,p.AllowFormattingColumns,p.AllowFormattingRows,p.AllowInsertingColumns,p.AllowInsertingRows,p.AllowInsertingHyperlinks,p.AllowDeletingColumns,p.AllowDeletingRows,p.AllowSorting,p.AllowFiltering,p.AllowUsingPivotTables};
            ws.Unprotect(password);Check(!ws.ProtectContents,"Temporary unprotect failed.");
            return new TrialProtection(()=>{ws.Protect(password,a[0],a[1],a[2],a[3],a[4],a[5],a[6],a[7],a[8],a[9],a[10],a[11],a[12],a[13],a[14]);ws.EnableSelection=selection;Check(ws.ProtectContents,"Protection restore failed.");});
        }
        internal override object AutoFillProbe(string sheet,int sourceCol,int destCol,int rows,int cols)
        {
            var cells=book.Worksheets[sheet].Cells;var formulas=new string[rows,cols];
            var values=Get(sheet,0,destCol,rows,cols);
            for(int r=0;r<rows;r++)for(int c=0;c<cols;c++)formulas[r,c]=cells[r,destCol+c].Formula;
            // Give AutoFill freshly inserted columns, not cells already populated by Copy.
            // Setup is excluded from its timing; these extra operations are trial-only.
            Columns(sheet,destCol,cols,false);Columns(sheet,destCol,cols,true);
            double time=Time(()=>cells[Address(0,sourceCol,rows,1)].AutoFill(cells[Address(0,Math.Min(sourceCol,destCol),rows,cols+1)],Gear.AutoFillType.FillCopy));
            for(int r=0;r<rows;r++)for(int c=0;c<cols;c++)
            {
                Check(cells[r,destCol+c].Formula==formulas[r,c],"AutoFill formula differs from Copy.");
                if(String.IsNullOrEmpty(formulas[r,c]))Check(Object.Equals(cells[r,destCol+c].Value,values[r,c]),"AutoFill constant differs from Copy.");
            }
            return new {milliseconds=time,mode="AutoFill FillCopy",freshlyInsertedTarget=true,formulasAndConstantsMatchCopy=true};
        }
        public override void Dispose(){while(set.Workbooks.Count>0)set.Workbooks[0].Close();}
    }
    sealed partial class ExcelEngine:Engine
    {
        dynamic app,books,seed,book;bool owned;readonly bool vba;
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
        internal ExcelEngine(bool vba)
        {
            this.vba=vba;var existing=Process.GetProcessesByName("EXCEL").Select(p=>p.Id).ToArray();
            try
            {
                app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));uint pid;GetWindowThreadProcessId(new IntPtr((int)app.Hwnd),out pid);
                Check(!existing.Contains((int)pid),"Refusing pre-existing Excel process.");owned=true;
                Console.WriteLine("OWNED_EXCEL_PID="+pid);
                app.Visible=false;app.EnableEvents=false;app.DisplayAlerts=false;app.AskToUpdateLinks=false;app.ScreenUpdating=false;
                app.AutomationSecurity=vba?2:3; // ByUI respects policy; never force-enable macros or change Trust Center.
                books=app.Workbooks;Check((int)books.Count==0,"Unexpected startup workbook.");seed=books.Add();app.Calculation=-4135;app.CalculateBeforeSave=false;
            }
            catch{Dispose();throw;}
        }
        internal override string Version {get{return "Excel "+app.Version+" build "+app.Build;}}
        internal override object Diagnostics {get{return new {vbaPermitted=vba,automationSecurity=vba?"ByUI":"ForceDisable",eventsEnabled=false,excelPeakWorkingSetBytes=ExcelMemory()};}}
        long ExcelMemory(){uint pid;GetWindowThreadProcessId(new IntPtr((int)app.Hwnd),out pid);using(var p=Process.GetProcessById((int)pid))return p.PeakWorkingSet64;}
        internal override void Open(string path)
        {
            using(var zip=ZipFile.OpenRead(path))Check(!zip.Entries.Any(e=>e.FullName.StartsWith("xl/macrosheets/",StringComparison.OrdinalIgnoreCase)||e.FullName.StartsWith("xl/intlmacrosheets/",StringComparison.OrdinalIgnoreCase)),"XLM macros outside trial scope.");
            book=books.Open(path,0,true);app.Calculation=-4135;
        }
        internal void SaveCopyAsXlsm(string path){book.SaveAs(path,52);}
        internal object VerifyWorkbookFunctions()
        {
            // Known non-zero answers prove VBA execution, independent of saved model caches.
            // Only a new, unsaved scratch workbook receives probe values.
            dynamic probe=null,sheets=null,sheet=null,rates=null,years=null;
            try
            {
                probe=books.Add();sheets=probe.Worksheets;sheet=sheets.Item[1];
                rates=sheet.Range["A1:A40"];years=sheet.Range["B1:B40"];
                var rv=new object[40,1];var yv=new object[40,1];
                for(int i=0;i<40;i++){rv[i,0]=(i+1)*0.02;yv[i,0]=(double)i+1;}
                rates.Value2=rv;years.Value2=yv;
                string prefix="'"+((string)book.Name).Replace("'","''")+"'!";
                double planned=Convert.ToDouble(app.Run(prefix+"PMCost",10,100,5.0,1,5,40,rates,years));
                double responsive=Convert.ToDouble(app.Run(prefix+"RespCost",10,100,1,5,40,rates,years));
                Check(Equal(planned,80)&&Equal(responsive,16),"Embedded VBA known answer differs.");
                TrialFunctions.VerifyKnownAnswers();
                return new {PMCost=planned,RespCost=responsive,expectedPMCost=80,expectedRespCost=16,trialPortKnownAnswersPassed=true};
            }
            finally
            {
                Release((object)years);Release((object)rates);Release((object)sheet);Release((object)sheets);
                if(probe!=null){try{probe.Close(false);}finally{Release((object)probe);}}
            }
        }
        internal override void Calculate(int mode){if(mode==2)app.CalculateFullRebuild();else if(mode==1)app.CalculateFull();else app.Calculate();}
        static void Release(object o){if(o!=null&&Marshal.IsComObject(o))Marshal.ReleaseComObject(o);}
        dynamic Sheet(string name){dynamic sheets=book.Worksheets;try{return sheets.Item[name];}finally{Release((object)sheets);}}
        internal override object[,] Get(string sheet,int row,int col,int rows,int cols)
        {
            dynamic s=Sheet(sheet),r=null;try{r=s.Range[Address(row,col,rows,cols)];object raw=r.Value2;var a=raw as object[,];var output=new object[rows,cols];for(int i=0;i<rows;i++)for(int j=0;j<cols;j++){object v=a==null?raw:a[i+a.GetLowerBound(0),j+a.GetLowerBound(1)];if(v is ErrorWrapper)v=((ErrorWrapper)v).ErrorCode;if(v is int && (int)v<=-2146826245 && (int)v>=-2146826288){int n=(int)v;v=n==-2146826281?"#DIV/0!":n==-2146826273?"#VALUE!":n==-2146826265?"#REF!":n==-2146826259?"#NAME?":n==-2146826252?"#NUM!":n==-2146826246?"#N/A":n==-2146826288?"#NULL!":"#ERROR:"+n;}output[i,j]=v;}return output;}finally{Release((object)r);Release((object)s);}
        }
        internal override void Set(string sheet,int row,int col,object[,] a){dynamic s=Sheet(sheet),r=null;try{r=s.Range[Address(row,col,a.GetLength(0),a.GetLength(1))];r.Value2=a;}finally{Release((object)r);Release((object)s);}}
        internal override void Columns(string sheet,int at,int count,bool insert){Mutate(sheet,Column(at)+":"+Column(at+count-1),insert,-4161);}
        internal override void Rows(string sheet,int at,int count,bool insert){Mutate(sheet,(at+1)+":"+(at+count),insert,-4121);}
        void Mutate(string sheet,string address,bool insert,int direction){dynamic s=Sheet(sheet),r=null;try{r=s.Range[address];if(insert)r.Insert(direction);else r.Delete(direction==-4161?-4159:-4162);}finally{Release((object)r);Release((object)s);}}
        internal override void Copy(string sheet,int row,int col,int rows,int cols,int dr,int dc,int dh,int dw){dynamic s=Sheet(sheet),from=null,to=null;try{from=s.Range[Address(row,col,rows,cols)];to=s.Range[Address(dr,dc,dh,dw)];from.Copy(to);}finally{Release((object)to);Release((object)from);Release((object)s);}}
        internal override string NameReference(string name){dynamic names=book.Names,n=null;try{n=names.Item(name);return (string)n.RefersTo;}finally{Release((object)n);Release((object)names);}}
        internal override Tuple<string,int,int> NameLocation(string name){dynamic names=book.Names,n=null,r=null,s=null;try{n=names.Item(name);r=n.RefersToRange;s=r.Worksheet;return Tuple.Create((string)s.Name,(int)r.Row-1,(int)r.Column-1);}finally{Release((object)s);Release((object)r);Release((object)n);Release((object)names);}}
        internal override Tuple<int,int> Size(string sheet){dynamic s=Sheet(sheet),r=null,rows=null,cols=null;try{r=s.UsedRange;rows=r.Rows;cols=r.Columns;return Tuple.Create((int)r.Row+(int)rows.Count-1,(int)r.Column+(int)cols.Count-1);}finally{Release((object)cols);Release((object)rows);Release((object)r);Release((object)s);}}
        internal override bool Protected(string sheet){dynamic s=Sheet(sheet);try{return (bool)s.ProtectContents;}finally{Release((object)s);}}
        internal override IDisposable TemporarilyUnprotect(string sheet,string password)
        {
            dynamic s=Sheet(sheet),p=null;
            try
            {
                p=s.Protection;int selection=(int)s.EnableSelection;
                var a=new bool[]{s.ProtectDrawingObjects,s.ProtectContents,s.ProtectScenarios,s.ProtectionMode,p.AllowFormattingCells,p.AllowFormattingColumns,p.AllowFormattingRows,p.AllowInsertingColumns,p.AllowInsertingRows,p.AllowInsertingHyperlinks,p.AllowDeletingColumns,p.AllowDeletingRows,p.AllowSorting,p.AllowFiltering,p.AllowUsingPivotTables};
                s.Unprotect(password);Check(!(bool)s.ProtectContents,"Temporary unprotect failed.");
                return new TrialProtection(()=>{try{s.Protect(password,a[0],a[1],a[2],a[3],a[4],a[5],a[6],a[7],a[8],a[9],a[10],a[11],a[12],a[13],a[14]);s.EnableSelection=selection;Check((bool)s.ProtectContents,"Protection restore failed.");}finally{Release((object)s);}});
            }
            catch{Release((object)s);throw;}
            finally{Release((object)p);}
        }
        public override void Dispose()
        {
            try{if(book!=null)book.Close(false);}finally{Release((object)book);book=null;try{if(seed!=null)try{seed.Close(false);}catch(COMException e){if(e.ErrorCode!=unchecked((int)0x80010114))throw;}}finally{Release((object)seed);seed=null;Release((object)books);books=null;if(app!=null){try{if(owned)app.Quit();}finally{Release((object)app);app=null;}}}}
        }
    }
}
