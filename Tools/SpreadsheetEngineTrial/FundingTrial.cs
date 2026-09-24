using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Gear = SpreadsheetGear;

// Isolated complete Funding trial. Never called by Summit. Outputs are disposable.
internal static partial class EngineBenchmark
{
    static readonly string[] ProbeSheets = {"Detailed Comp Inc - Trad View","Financial Position - Trad View","Cashflow detailed","Check Sheet","Development Expenditure"};
    static string[] FundingSheets()
    {
        string source=File.ReadAllText(Path.Combine(Repo,"Services/WorkbookStructureRuleManager.vb"));
        var block=Regex.Match(source,@"(?s)AddColumnTargets\(FundingRule,(.*?)\)\s*'Funding_Columns");
        var sheets=Regex.Matches(block.Groups[1].Value,"\"([^\"]+)\"").Cast<Match>().Select(m=>m.Groups[1].Value).ToArray();
        Check(sheets.Length==32,"Reviewed Funding group changed; review before testing.");return sheets;
    }
    static string[] Mirrors(string source)
    {return Enumerable.Range(0,source=="FacilityNames"?9:2).Select(i=>"TransCopy_"+source+"_"+(char)('A'+i)).ToArray();}
    static object Probes(Engine engine)
    {
        var data=new Dictionary<string,object>();foreach(var s in ProbeSheets){var size=engine.Size(s);Check(size.Item1<5000&&size.Item2<300,"Unexpected extent.");data[s]=new {rows=size.Item1,columns=size.Item2,values=Jagged(engine.Get(s,0,0,size.Item1,size.Item2))};}return data;
    }
    internal static int RunFunding(string[] args)
    {
        if(args.Length==4&&args[1]=="dx-reload")
        {
            string file=IoTrial.InTrial(args[2]),path=IoTrial.InTrial(args[3]);Check(!File.Exists(path),"New report required.");string beforeHash=IoTrial.FileHash(file);
            using(var reader=new DxEngine(true))
            {
                double load=Time(()=>reader.Open(file));
                IoTrial.WriteJson(path,new {success=true,input=file,inputHash=beforeHash,sourceUnchanged=beforeHash==IoTrial.FileHash(file),loadMs=load,reader.Version,bits=IntPtr.Size*8,scope="DevExpress native reload only; NO Summit UI rebind or application service initialization"});
            }
            return 0;
        }
        Check(args.Length==5,"funding gear-copy|gear-fill|gear-baseline|excel|excel-baseline INPUT NEW_OUTPUT NEW_REPORT");
        CultureInfo.CurrentCulture=CultureInfo.InvariantCulture;
        string mode=args[1],input=IoTrial.InTrial(args[2]),output=IoTrial.InTrial(args[3]),report=IoTrial.InTrial(args[4]);
        Check(new[]{"gear-copy","gear-fill","gear-baseline","excel","excel-baseline"}.Contains(mode),"Unknown mode.");
        Check(Path.GetExtension(input)==".xlsm"&&Path.GetExtension(output)==".xlsm","Use the common converted XLSM fixture.");
        Check(!File.Exists(output)&&!File.Exists(report),"Refusing existing output.");
        string hash=IoTrial.FileHash(input);var times=new Dictionary<string,object>();
        var result=new Dictionary<string,object>{{"engine",mode},{"mode",mode},{"input",input},{"output",output},{"inputHash",hash},{"bits",IntPtr.Size*8},{"utc",DateTime.UtcNow.ToString("o")},{"timingsMs",times},{"scope","Whole native Funding operation and serialization; excludes Summit staging, transaction, UI rebind and validation"}};
        try
        {
            bool excel=mode.StartsWith("excel");bool mutate=!mode.EndsWith("baseline");
            using(Engine engine=excel?(Engine)new ExcelEngine(true):new GearEngine(true))
            {
                Console.WriteLine("START "+mode+" load");times["load"]=Time(()=>engine.Open(input));result["version"]=engine.Version;
                if(mutate){Console.WriteLine("START full Funding insertion");times["fundingCommand"]=Time(()=>{if(excel)((ExcelEngine)engine).FundingVba(times,result);else ((GearEngine)engine).FundingNative(mode=="gear-fill",times,result);});}
                Console.WriteLine("START full rebuild");times["fullRebuild"]=Time(()=>engine.Calculate(2));result["calculatedProbes"]=Probes(engine);
                Console.WriteLine("START save disposable output");times["save"]=Time(()=>{if(excel)((ExcelEngine)engine).SaveFunding(output);else ((GearEngine)engine).SaveFunding(output);});
                result["engineDiagnostics"]=engine.Diagnostics;
            }
            result["success"]=true;
        }
        catch(Exception e){result["success"]=false;result["error"]=e.GetType().Name+": "+e.Message;result["stack"]=e.StackTrace;Console.WriteLine("FAILED "+result["error"]);}
        finally{result["sourceUnchanged"]=hash==IoTrial.FileHash(input);IoTrial.WriteJson(report,result);Console.WriteLine("RESULT "+report);Check((bool)result["sourceUnchanged"],"Input changed.");}
        return (bool)result["success"]?0:1;
    }
    sealed partial class GearEngine
    {
        internal void SaveFunding(string path){RequireFullWorkbookForSave();using(var stream=new FileStream(path,FileMode.CreateNew,FileAccess.Write))book.SaveToStream(stream,Gear.FileFormat.OpenXMLWorkbookMacroEnabled);}
        internal void FundingNative(bool fill,Dictionary<string,object> times,Dictionary<string,object> result)
        {
            var sheets=FundingSheets();var facility=book.Names["FacilityNames"].RefersToRange;var ordinary=book.Names["LoanDescsOrd"].RefersToRange;
            int before=facility.ColumnCount,ordinaryBefore=ordinary.ColumnCount,at=book.Names["LoanDescRev1"].RefersToRange.Column-1;const int count=10;
            Check(at>0&&ordinary.Column+ordinary.ColumnCount-1==at,"Funding boundary differs.");
            foreach(string source in new[]{"FacilityNames","LoanDescsOrd"})foreach(string n in Mirrors(source))Check(book.Names[n].RefersToRange.RowCount==book.Names[source].RefersToRange.ColumnCount+1,"Existing mirror misaligned: "+n);
            var restores=new List<IDisposable>();
            try
            {
                times["unprotect"]=Time(()=>{string credential=TrialProtection.ReadExistingCredential();foreach(string s in sheets.Concat(new[]{"Transactional DB"}))if(Protected(s))restores.Add(TemporarilyUnprotect(s,credential));});
                times["insert32Sheets"]=Time(()=>{foreach(string s in sheets)Columns(s,at,count,true);});
                times["copy32Sheets"]=Time(()=>{foreach(string s in sheets){var cells=book.Worksheets[s].Cells;var template=cells[Column(at+count)+":"+Column(at+count)];if(fill)template.AutoFill(cells[Column(at)+":"+Column(at+count)],Gear.AutoFillType.FillCopy);else template.Copy(cells[Column(at)+":"+Column(at+count-1)]);}});
                times["mirrors11"]=Time(()=>{
                    // VBA order and whole-row geometry. Default AutoFill, as in master VBA;
                    // Copy-v-Fill experiment concerns only the 32 funding column templates.
                    foreach(string source in new[]{"FacilityNames","LoanDescsOrd"})foreach(string n in Mirrors(source))
                    {
                        var nr=book.Names[n];var range=nr.RefersToRange;int top=range.Row,left=range.Column,width=range.ColumnCount,oldRows=range.RowCount,bottom=top+oldRows-1;
                        int add=book.Names[source].RefersToRange.ColumnCount+1-range.RowCount;Check(add==count,"Unexpected mirror resize.");
                        var ws=range.Worksheet;Rows(ws.Name,bottom,add,true);
                        ws.Cells[Address(bottom-1,left,1,width)].AutoFill(ws.Cells[Address(bottom-1,left,add+1,width)],Gear.AutoFillType.FillDefault);
                        Check(nr.RefersToRange.RowCount==oldRows+add,"Mirror name failed to extend: "+n);
                    }
                });
                Check(book.Names["FacilityNames"].RefersToRange.ColumnCount==before+count&&book.Names["LoanDescsOrd"].RefersToRange.ColumnCount==ordinaryBefore+count,"Source names failed to extend.");
                foreach(string source in new[]{"FacilityNames","LoanDescsOrd"})foreach(string n in Mirrors(source))Check(book.Names[n].RefersToRange.RowCount==book.Names[source].RefersToRange.ColumnCount+1,"Mirror mismatch: "+n);
                result["geometryVerified"]=true;
            }
            finally{times["restoreProtection"]=Time(()=>{for(int i=restores.Count-1;i>=0;i--)restores[i].Dispose();});}
        }
    }
    sealed partial class ExcelEngine
    {
        internal void SaveFunding(string path){Check(!File.Exists(path),"Output exists.");book.SaveCopyAs(path);}
        int NameWidth(string name){dynamic names=book.Names,n=null,r=null,cols=null;try{n=names.Item(name);r=n.RefersToRange;cols=r.Columns;return (int)cols.Count;}finally{Release((object)cols);Release((object)r);Release((object)n);Release((object)names);}}
        int NameHeight(string name){dynamic names=book.Names,n=null,r=null,rows=null;try{n=names.Item(name);r=n.RefersToRange;rows=r.Rows;return (int)rows.Count;}finally{Release((object)rows);Release((object)r);Release((object)n);Release((object)names);}}
        internal void FundingVba(Dictionary<string,object> times,Dictionary<string,object> result)
        {
            int before=NameWidth("FacilityNames"),ordinary=NameWidth("LoanDescsOrd");
            string prefix="'"+((string)book.Name).Replace("'","''")+"'!";
            book.Activate();
            // Avoid only the outer InputBox/MsgBox wrapper. Execute unmodified embedded code.
            times["vbaInsert32Sheets"]=Time(()=>app.Run(prefix+"Run_Insert_Funding_Columns",10,"Null","BP"));
            Check(NameWidth("FacilityNames")==before+10&&NameWidth("LoanDescsOrd")==ordinary+10,"VBA insertion failed or aborted.");
            using(Protected("Transactional DB")?TemporarilyUnprotect("Transactional DB",TrialProtection.ReadExistingCredential()):null)
                times["vbaMirrors11"]=Time(()=>app.Run(prefix+"FundingSynchroniseTransactionalDBSheet"));
            foreach(string source in new[]{"FacilityNames","LoanDescsOrd"})foreach(string n in Mirrors(source))Check(NameHeight(n)==NameWidth(source)+1,"VBA mirror mismatch: "+n);
            Check(!(bool)app.EnableEvents,"Macro unexpectedly enabled events.");app.Calculation=-4135;app.ScreenUpdating=false;
            result["geometryVerified"]=true;result["actualEmbeddedVbaExecuted"]=true;
        }
    }
}
