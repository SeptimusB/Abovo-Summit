using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using S = GrapeCity.Spreadsheet;
using C = GrapeCity.CalcEngine;
using X = GrapeCity.Spreadsheet.IO.OpenXml;

internal static partial class EngineBenchmark
{
    internal sealed class SpreadEngine : Engine
    {
        internal S.IWorkbookSet SetOfBooks;
        internal S.IWorkbook Book;
        readonly FarPoint.Win.Spread.FpSpread host;
        readonly bool model;
        readonly SpreadFunctions functions = new SpreadFunctions();
        bool registered;
        internal readonly List<object> IoWarnings = new List<object>();
        void CaptureWarnings(FarPoint.Excel.ExcelWarningList warnings)
        {
            for(int i=0;i<warnings.Count;i++) { var w=warnings[i]; IoWarnings.Add(new {code=w.Code.ToString(),w.Message,w.Sheet,w.Row,w.Column}); }
        }
        internal SpreadEngine(bool model)
        {
            this.model = model;
            if(Environment.GetEnvironmentVariable("SPREAD_TRIAL_HEADLESS")!="1")
            {
                host = new FarPoint.Win.Spread.FpSpread(FarPoint.Win.Spread.LegacyBehaviors.None);
                SetOfBooks = host.AsWorkbook().WorkbookSet;
            }
            else SetOfBooks = S.Factory.CreateWorkbookSet(CultureInfo.InvariantCulture);
            Configure();
            if(model&&host!=null) { host.AddCustomFunction(functions.Pm);host.AddCustomFunction(functions.Resp);registered=true; }
        }
        void Configure()
        {
            var c = SetOfBooks.CalculationEngine;
            c.Calculation = S.Calculation.Manual;
            c.BackgroundCalculation = false;
            c.CalculationOnDemandMode = S.CalculationOnDemandMode.Off;
            c.CalculateBeforeSave = false;
            var features=(c.CalcFeatures & ~S.CalcFeatures.Optimal) | S.CalcFeatures.DynamicArray;
            if(Environment.GetEnvironmentVariable("SPREAD_TRIAL_OPTIMAL")=="1")features|=S.CalcFeatures.Optimal;
            if(features!=c.CalcFeatures)c.CalcFeatures=features;
        }
        internal override string Version { get { return typeof(S.IWorkbook).Assembly.GetName().Version.ToString(); } }
        internal override object Diagnostics { get { return new { functions.Pm.Calls, pmInvalid = functions.Pm.InvalidCalls,
            respCalls = functions.Resp.Calls, respInvalid = functions.Resp.InvalidCalls,
            ioWarnings=IoWarnings, calcFeatures=SetOfBooks.CalculationEngine.CalcFeatures.ToString(),host=host==null?"Core factory":"FpSpread (not displayed)",
            mode = "Standalone primary engine, no native worker or repair of exported package" }; } }
        internal override void Open(string path)
        {
            if(host==null)
            {
                var context=new X.ImportContext(new X.UnitConverter(96),null){Options=X.ImportOptions.Default|X.ImportOptions.StopRecalculate};
                if(model)context.FunctionResolver=functions;
                using(var input=File.OpenRead(path))Book=SetOfBooks.Workbooks.Open(input,context,false);
            }
            else
            {
                using (var input = File.OpenRead(path)) Check(host.OpenExcel(input,FarPoint.Excel.ExcelOpenFlags.DocumentCaching | FarPoint.Excel.ExcelOpenFlags.DoNotRecalculateAfterLoad),"Native XLSM open rejected.");
                Book = host.AsWorkbook(); SetOfBooks = Book.WorkbookSet;
            }
            Check(Book!=null,"Spread.NET rejected workbook open.");
            Configure();
            if(model&&!registered)RegisterFunctions();
            Check(!Book.Date1904, "Fixture requires Excel's 1900 date system.");
        }
        internal override void Calculate(int mode) { SetOfBooks.CalculationEngine.Calculate(Book, mode != 0, mode == 2); }
        internal void RegisterFunctions()
        {
            if(host!=null)
            {
                if(!Object.ReferenceEquals(host.AsWorkbook(),Book))Check(host.Attach(Book),"UDF host attach rejected.");
                host.AddCustomFunction(functions.Pm);host.AddCustomFunction(functions.Resp);
            }
            else ((S.WorkbookBase)Book).FormulaEngine.AddResolver(functions);
        }
        internal void AttachCurrentBook() { if(host!=null&&!Object.ReferenceEquals(host.AsWorkbook(),Book))Check(host.Attach(Book),"Host attach failed."); }
        internal void InsertViaSheetView(string sheet,int index,int count,bool rows)
        {
            AttachCurrentBook();var view=host.Sheets.Cast<FarPoint.Win.Spread.SheetView>().Single(s=>s.SheetName==sheet);
            if(rows)view.AddRows(index,count);else view.AddColumns(index,count);
        }
        internal override object[,] Get(string sheet, int row, int col, int rows, int cols)
        {
            var ws=Book.Worksheets[sheet];
            var values=new object[rows,cols];
            for (int r = 0; r < rows; r++) for (int c = 0; c < cols; c++)
            {
                values[r,c]=ws.GetValue(row+r,col+c);
                if (values[r,c] is DateTime) values[r,c] = ((DateTime)values[r,c]).ToOADate();
                else if (values[r,c] is C.CalcError) values[r,c] = Error(values[r,c].ToString());
            }
            return values;
        }
        internal override void Set(string sheet, int row, int col, object[,] values)
        {
            Check(Book.Worksheets[sheet].SetValue(row,col,values),"Native block write rejected.");
        }
        internal override void Columns(string sheet, int at, int count, bool insert)
        {
            var ws = Book.Worksheets[sheet];
            var result=insert?ws.InsertColumns(at,count):ws.RemoveColumns(at,count);
            Check(result.Success,"Native column operation rejected: "+sheet);
        }
        internal override void Rows(string sheet, int at, int count, bool insert)
        {
            var ws = Book.Worksheets[sheet];
            var result=insert?ws.InsertRows(at,count):ws.RemoveRows(at,count);
            Check(result.Success,"Native row operation rejected: "+sheet);
        }
        internal override void Copy(string sheet, int row, int col, int rows, int cols, int dr, int dc, int dh, int dw)
        {
            var cells = Book.Worksheets[sheet].Cells;
            Check(dh%rows==0 && dw%cols==0,"Copy tiling requires complete tiles.");
            for(int r=0;r<dh;r+=rows)for(int c=0;c<dw;c+=cols)
                Check(cells[Address(row,col,rows,cols)].Copy(cells[Address(dr+r,dc+c,rows,cols)]), "Native range copy rejected.");
        }
        internal override string NameReference(string name) { return Book.Names.Get(name).GetRefersTo(0,0); }
        internal override Tuple<string,int,int> NameLocation(string name)
        {
            var r = Book.Names.Get(name).RefersToRange(0,0);
            return Tuple.Create(r.Worksheet.Name, r.Row, r.Column);
        }
        internal override Tuple<int,int> Size(string sheet)
        {
            var r = Book.Worksheets[sheet].UsedRange;
            return Tuple.Create(r.Row2+1,r.Column2+1);
        }
        internal override bool Protected(string sheet) { return Book.Worksheets[sheet].ProtectionMode != S.ProtectionMode.None; }
        internal override IDisposable TemporarilyUnprotect(string sheet, string password)
        {
            var ws = (S.Worksheet)Book.Worksheets[sheet]; var locks = ws.Locks;
            Check(ws.Unprotect(password) && !Protected(sheet), "Temporary unprotect failed.");
            return new TrialProtection(() => { Check(ws.Protect(locks,password) && ws.Locks == locks, "Protection restore failed."); });
        }
        internal void Save(string path, bool lossless)
        {
            path = IoTrial.InTrial(path);
            if(host!=null&&!Object.ReferenceEquals(host.AsWorkbook(),Book)) Check(host.Attach(Book),"Native host attach rejected.");
            var flags=FarPoint.Excel.ExcelSaveFlags.UseOOXMLFormat | FarPoint.Excel.ExcelSaveFlags.MacroEnabledWorkbook;
            if(lossless)flags |= FarPoint.Excel.ExcelSaveFlags.DocumentCaching;
            using (var output = new FileStream(path,FileMode.CreateNew,FileAccess.ReadWrite))
            {
                var warnings=new FarPoint.Excel.ExcelWarningList();
                var context=new X.ExportContext(new X.UnitConverter(96)){Options=lossless?X.ExportOptions.Default|X.ExportOptions.Lossless:X.ExportOptions.Default};
                bool saved=host==null?Book.SaveAs(output,S.IO.FileFormat.OpenXMLWorkbookMacroEnabled,null,context):host.SaveExcel(output,flags,warnings);
                CaptureWarnings(warnings);
                if(!saved)Console.WriteLine("SAVE WARNINGS "+IoTrial.Json.Serialize(IoWarnings));
                Check(saved,"Native XLSM save rejected.");
                Check(output.Length>0,"Native XLSM save produced empty output.");
            }
        }
        public override void Dispose() { if(host!=null)host.Dispose(); if (Book != null) { Book.Close(false); Book=null; } }
    }

    internal static int SpreadSmoke()
    {
        using (var engine = new SpreadEngine(false))
        {
            engine.Book = engine.SetOfBooks.Workbooks.Add("Smoke");
            var ws=engine.Book.Worksheets.Add(); ws.Name="Data";
            ws.Cells["A1"].Value=3d; ws.Cells["B1"].Formula2="=SEQUENCE(A1)";
            engine.Calculate(2);
            Console.WriteLine("SMOKE formula="+ws.Cells["B1"].Formula2+" values="+ws.Cells["B1"].Value+","+ws.Cells["B2"].Value+","+ws.Cells["B3"].Value+" features="+engine.SetOfBooks.CalculationEngine.CalcFeatures);
            Check(Convert.ToDouble(ws.Cells["B3"].Value)==3,"Dynamic spill failed.");
            ws.Cells["A1"].Value=5d; engine.Calculate(0);
            Check(Convert.ToDouble(ws.Cells["B5"].Value)==5,"Dynamic spill growth failed.");
            engine.RegisterFunctions();
            for(int i=0;i<40;i++){ws.Cells[i,3].Value=(i+1)*.02;ws.Cells[i,4].Value=i+1;}
            ws.Cells["G1"].Formula2="=PMCost(10,100,5,1,5,40,D1:D40,E1:E40)";
            ws.Cells["G2"].Formula2="=RespCost(10,100,1,5,40,D1:D40,E1:E40)";
            ws.Cells["J10"].Formula2="=A1*2";engine.Calculate(2);
            Console.WriteLine("UDF values="+ws.Cells["G1"].Value+","+ws.Cells["G2"].Value+"; scalar="+ws.Cells["J10"].Value+"; diagnostics="+IoTrial.Json.Serialize(engine.Diagnostics));
            Check(Convert.ToDouble(ws.Cells["G1"].Value)==80&&Convert.ToDouble(ws.Cells["G2"].Value)==16,"UDF known answers failed.");
            Check(Convert.ToDouble(ws.Cells["J10"].Value)==10,"Distant scalar reference failed.");
            Console.WriteLine("PASS native dynamic spill and incremental growth; "+engine.Version);
        }
        return 0;
    }

    internal static int SpreadIo(string[] args)
    {
        Check(args.Length == 5, "io INPUT NEW_OUTPUT NEW_REPORT default|lossless");
        string input=IoTrial.InTrial(args[1]),output=IoTrial.InTrial(args[2]),report=IoTrial.InTrial(args[3]);
        Check(!File.Exists(output)&&!File.Exists(report),"New outputs required.");
        string hash=IoTrial.FileHash(input);
        var result=new Dictionary<string,object>{{"input",input},{"output",output},{"inputHash",hash},{"mode",args[4]},{"bits",IntPtr.Size*8}};
        try
        {
            using(var engine=new SpreadEngine(false))
            {
                result["loadMs"]=Time(()=>engine.Open(input));
                result["version"]=engine.Version;result["sheetCount"]=engine.Book.Worksheets.Count;
                result["saveMs"]=Time(()=>engine.Save(output,args[4]=="lossless"));
            }
            using(var engine=new SpreadEngine(false))result["reopenMs"]=Time(()=>engine.Open(output));
            result["success"]=true;
        }
        catch(Exception e){result["success"]=false;result["error"]=e.ToString();Console.WriteLine(e);}
        result["sourceUnchanged"]=hash==IoTrial.FileHash(input);
        result["peakWorkingSetBytes"]=Process.GetCurrentProcess().PeakWorkingSet64;
        IoTrial.WriteJson(report,result);
        Console.WriteLine("RESULT "+report);
        return Object.Equals(result["success"],true)?0:1;
    }

    sealed class SpreadFunctions : C.IFunctionResolver
    {
        internal readonly SpreadFunction Pm = new SpreadFunction(true), Resp = new SpreadFunction(false);
        public C.Function Resolve(string name)
        {
            if(String.Equals(name,"PMCost",StringComparison.OrdinalIgnoreCase))return Pm;
            if(String.Equals(name,"RespCost",StringComparison.OrdinalIgnoreCase))return Resp;
            return null;
        }
        public C.Function[] Find(string prefix) { return new C.Function[]{Pm,Resp}.Where(f=>f.Name.StartsWith(prefix??"",StringComparison.OrdinalIgnoreCase)).ToArray(); }
    }
    sealed class SpreadFunction : C.Function
    {
        readonly bool pm;
        internal long Calls, InvalidCalls;
        internal SpreadFunction(bool pm) : base(pm?"PMCost":"RespCost",pm?8:7,pm?8:7,C.FunctionAttributes.SingleCell | C.FunctionAttributes.Number) { this.pm=pm; }
        protected override void Evaluate(C.IArguments args,C.IValue result)
        {
            Calls++;
            try
            {
                var ctx=args.EvaluationContext;int n=pm?6:5;
                var scalars=new double[n];for(int i=0;i<n;i++)scalars[i]=args[i].GetNumber(ctx);
                result.SetValue(ctx,TrialFunctions.Cost(pm,scalars,Vector(args[n],ctx,false),Vector(args[n+1],ctx,true)));
            }
            catch { InvalidCalls++; args.EvaluationContext.Error=C.CalcError.Value; }
        }
        static double[] Vector(C.IValue value,C.IEvaluationContext ctx,bool years)
        {
            var source=value.GetReferenceSource(ctx);var range=value.GetAccessibleReference(ctx,0);
            var values=new List<double>();
            for(int row=range.Row;row<=range.Row2;row++)for(int col=range.Column;col<=range.Column2;col++)
            {
                C.CellValue cell=new C.CellValue();source.GetValue(ctx,row,col,ref cell);
                string type=cell.Type2.ToString();
                if(type=="Empty"||type=="Blank")values.Add(years?Double.NaN:0);
                else if(type=="Number")values.Add(cell.Number);
                else throw new ArgumentException("Non-numeric lookup.");
            }
            return values.ToArray();
        }
    }
}
