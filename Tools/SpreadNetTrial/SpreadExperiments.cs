using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using S=GrapeCity.Spreadsheet;

internal static partial class EngineBenchmark
{
    internal static int SpreadBatch(string inputDirectory,string outputDirectory)
    {
        inputDirectory=Path.GetDirectoryName(IoTrial.InTrial(Path.Combine(inputDirectory,"guard.txt")));
        outputDirectory=Path.GetDirectoryName(IoTrial.InTrial(Path.Combine(outputDirectory,"guard.txt")));
        int failed=0;
        for(int i=1;i<=3;i++)failed+=Run(new[]{"bench","spreadnet",Path.Combine(inputDirectory,"synthetic.xlsm"),Path.Combine(outputDirectory,"synthetic-"+i+".json"),"synthetic"});
        failed+=Run(new[]{"bench","spreadnet",Path.Combine(inputDirectory,"AGL-excel-converted.xlsm"),Path.Combine(outputDirectory,"agl.json"),"agl"});
        return failed==0?0:1;
    }
    internal static int SpreadVerifyExcel(string[] args)
    {
        string directory=Path.GetDirectoryName(IoTrial.InTrial(Path.Combine(args[1],"guard.txt"))),report=IoTrial.InTrial(args[2]);
        Check(!File.Exists(report),"New report required.");var results=new List<object>();
        foreach(string source in Directory.GetFiles(directory,"spill-insert-*.xlsm"))
        {
            string input=IoTrial.InTrial(source),hash=IoTrial.FileHash(input);
            using(var e=new ExcelEngine(false))
            {
                e.Open(input);e.Calculate(2);object before=e.Get("Data",11,5,1,1)[0,0];
                e.Set("Data",0,0,new object[,]{{5d}});e.Calculate(2);
                object scalar=e.Get("Data",11,5,1,1)[0,0],spill=e.Get("Data",4,1,1,1)[0,0];
                results.Add(new{input,before,scalar,spill,passed=Equal(Convert.ToDouble(before),6)&&Equal(Convert.ToDouble(scalar),10)&&Equal(Convert.ToDouble(spill),5),sourceUnchanged=hash==IoTrial.FileHash(input)});
            }
        }
        IoTrial.WriteJson(report,new{results});Console.WriteLine(IoTrial.Json.Serialize(results));return 0;
    }
    internal static int SpreadExcelOpen(string[] args)
    {
        string input=IoTrial.InTrial(args[1]),report=IoTrial.InTrial(args[2]),hash=IoTrial.FileHash(input);
        Check(!File.Exists(report),"New report required.");var result=new Dictionary<string,object>{{"input",input},{"inputHash",hash},{"macrosEnabled",false},{"calculationRequested",false},{"saveRequested",false}};
        try
        {
            using(var e=new ExcelEngine(false)){result["loadMs"]=Time(()=>e.Open(input));result["fundingSize"]=e.Size("Funding Assumptions");result["loanName"]=e.NameReference("LoanDescRev1");result["success"]=true;}
        }
        catch(Exception ex){result["success"]=false;result["error"]=ex.ToString();}
        result["sourceUnchanged"]=hash==IoTrial.FileHash(input);IoTrial.WriteJson(report,result);Console.WriteLine(IoTrial.Json.Serialize(result));return Object.Equals(result["success"],true)?0:1;
    }
    internal static int SpreadDiagnose(string[] args)
    {
        string input=IoTrial.InTrial(args[1]),report=IoTrial.InTrial(args[2]),hash=IoTrial.FileHash(input);
        Check(!File.Exists(report),"New report required.");
        var result=new Dictionary<string,object>{{"inputHash",hash}};
        using(var e=new SpreadEngine(true))
        {
            result["loadMs"]=Time(()=>e.Open(input));
            result["fullMs"]=Time(()=>e.Calculate(2));
            result["diagnostics"]=e.Diagnostics;
            var cells=new List<object>();
            foreach(var location in new[]{new[]{"Global Assumptions","C6"},new[]{"Global Assumptions","C10"},new[]{"Development Expenditure","A1"},new[]{"Development Expenditure","AV417"}})
            { var c=e.Book.Worksheets[location[0]].Cells[location[1]];cells.Add(new{sheet=location[0],cell=location[1],c.Formula,c.Formula2,c.Value,c.HasArray}); }
            result["cells"]=cells;result["calculatedProbes"]=SpreadProbes(e);
        }
        result["sourceUnchanged"]=IoTrial.FileHash(input)==hash;IoTrial.WriteJson(report,result);
        Console.WriteLine("RESULT "+report);return 0;
    }
    internal static int SpreadCases(string directory)
    {
        directory=IoTrial.InTrial(Path.Combine(directory,"guard.txt"));
        directory=Path.GetDirectoryName(directory);
        Directory.CreateDirectory(directory);
        string report=Path.Combine(directory,"cases.json");
        Check(!File.Exists(report),"New report required.");
        var results=new List<object>();bool casesPassed=true;
        foreach(string mode in new[]{"serial","selected-range","selected-api","selected-command"})
        {
            using(var e=new SpreadEngine(false))
            {
                e.Book=e.SetOfBooks.Workbooks.Add("Group");
                var first=e.Book.Worksheets.Add("First");var last=e.Book.Worksheets.Add("Last");var result=e.Book.Worksheets.Add("Result");
                e.AttachCurrentBook();
                first.Cells["B1"].Value=5d;last.Cells["B1"].Value=7d;result.Cells["A1"].Formula2="=SUM(First:Last!B1)";
                e.Calculate(2);Check(Convert.ToDouble(result.Cells["A1"].Value)==12d,"3D baseline failed.");
                string before=result.Cells["A1"].Formula2;
                if(mode=="serial"){e.Columns("First",1,2,true);e.Columns("Last",1,2,true);}
                else
                {
                    e.Book.Worksheets[new[]{"First","Last"}].Select(true);
                    if(mode=="selected-range")first.Cells["B:C"].Insert(S.InsertShiftDirection.EntireColumn);
                    else if(mode=="selected-api")Check(first.InsertColumns(1,2).Success,"Grouped API rejected.");
                    else
                    {
                        first.Cells["B:C"].Select();
                        var command=S.Commands.ClipboardInsertCommand.InsertColumns((S.Worksheet)first);
                        Check(command!=null&&command.Execute((S.Workbook)e.Book).Success,"Grouped native command rejected.");
                    }
                    first.Select(true);
                }
                e.Calculate(2);
                bool groupPassed=mode=="serial"?Convert.ToDouble(result.Cells["A1"].Value)==0:
                    result.Cells["A1"].Formula2.Contains("D1")&&Convert.ToDouble(first.Cells["D1"].Value)==5&&Convert.ToDouble(last.Cells["D1"].Value)==7;
                casesPassed &= groupPassed;
                results.Add(new {test="3D "+mode,before,after=result.Cells["A1"].Formula2,value=result.Cells["A1"].Value,
                    first=first.Cells["D1"].Value,last=last.Cells["D1"].Value,selectedSheets=e.Book.SelectedSheets.Count,passed=groupPassed});
                Console.WriteLine(IoTrial.Json.Serialize(results.Last()));
            }
        }
        foreach(string insertMode in Environment.GetEnvironmentVariable("SPREAD_TRIAL_HEADLESS")=="1"?new[]{"worksheet","range"}:new[]{"worksheet","range","sheetview"})using(var e=new SpreadEngine(false))
        {
            e.Book=e.SetOfBooks.Workbooks.Add("Spill");
            var ws=e.Book.Worksheets.Add("Data");
            e.AttachCurrentBook();
            ws.Cells["A1"].Value=3d;ws.Cells["B1"].Formula2="=SEQUENCE(A1)";
            ws.Cells["D10"].Formula2="=A1*2";e.Calculate(2);
            Check(Convert.ToDouble(ws.Cells["D10"].Value)==6,"Pre-insert scalar failed.");
            string file=Path.Combine(directory,"spill-insert-"+insertMode+".xlsm");
            if(insertMode=="worksheet") {e.Rows("Data",4,2,true);e.Columns("Data",2,2,true);}
            else if(insertMode=="range") {ws.Cells["5:6"].Insert(S.InsertShiftDirection.EntireRow);ws.Cells["C:D"].Insert(S.InsertShiftDirection.EntireColumn);}
            else {e.InsertViaSheetView("Data",4,2,true);e.InsertViaSheetView("Data",2,2,false);}
            e.Calculate(2);
            casesPassed &= Convert.ToDouble(ws.Cells["F12"].Value)==6;
            results.Add(new{test="Scalar reference after "+insertMode+" insert",value=ws.Cells["F12"].Value,formula=ws.Cells["F12"].Formula2,passed=Convert.ToDouble(ws.Cells["F12"].Value)==6});
            Console.WriteLine("ARRAY after insert="+ws.Cells["F12"].Formula2+" value="+ws.Cells["F12"].Value+" old="+ws.Cells["D10"].Value+" A1="+ws.Cells["A1"].Value);
            e.Save(file,true);
            using(var reopen=new SpreadEngine(false))
            {
                reopen.Open(file);reopen.Set("Data",0,0,new object[,]{{5d}});reopen.Calculate(2);
                var sheet=reopen.Book.Worksheets["Data"];
                bool pass=Convert.ToDouble(sheet.Cells["B5"].Value)==5d&&Convert.ToDouble(sheet.Cells["F12"].Value)==10d;
                casesPassed &= pass;
                results.Add(new{test="spill growth and "+insertMode+" insert after save/reopen",passed=pass,
                    anchor=sheet.Cells["B1"].Formula2,afterInsert=sheet.Cells["F12"].Formula2,spillValue=sheet.Cells["B5"].Value});
            }
        }
        IoTrial.WriteJson(report,new{results,passed=casesPassed});
        Console.WriteLine(IoTrial.Json.Serialize(results));
        return casesPassed?0:1;
    }

    static object SpreadProbes(Engine e)
    {
        var probes=new Dictionary<string,object>();
        foreach(string sheet in new[]{"Detailed Comp Inc - Trad View","Financial Position - Trad View","Cashflow detailed","Check Sheet","Development Expenditure"})
        {
            var size=e.Size(sheet);Check(size.Item1<5000&&size.Item2<300,"Unexpected probe geometry.");
            probes[sheet]=new{rows=size.Item1,columns=size.Item2,values=Jagged(e.Get(sheet,0,0,size.Item1,size.Item2))};
        }
        return probes;
    }
    internal static int SpreadEdit(string[] args)
    {
        Check(args.Length==4,"edit ENGINE INPUT NEW_REPORT");
        string input=IoTrial.InTrial(args[2]),report=IoTrial.InTrial(args[3]),hash=IoTrial.FileHash(input);
        Check(!File.Exists(report),"New report required.");
        var result=new Dictionary<string,object>{{"engine",args[1]},{"input",input},{"inputHash",hash},{"bits",IntPtr.Size*8}};
        var times=new Dictionary<string,object>();result["timingsMs"]=times;
        try
        {
            using(var e=Create(args[1],true))
            {
                times["load"]=Time(()=>e.Open(input));times["baselineFull"]=Time(()=>e.Calculate(2));
                result["baseline"]=SpreadProbes(e);
                object old=e.Get("Funding Assumptions",81,6,1,1)[0,0];
                Check(old is IConvertible&&!(old is string),"Funding opening balance must be numeric.");
                double next=Convert.ToDouble(old)+300;
                using(e.Protected("Funding Assumptions")?e.TemporarilyUnprotect("Funding Assumptions",TrialProtection.ReadExistingCredential()):null)
                {
                    times["editWrite"]=Time(()=>e.Set("Funding Assumptions",81,6,new object[,]{{next}}));
                    times["editedIncremental"]=Time(()=>e.Calculate(0));result["editedIncremental"]=SpreadProbes(e);
                    times["editedFull"]=Time(()=>e.Calculate(1));result["editedFull"]=SpreadProbes(e);
                    times["restoreWrite"]=Time(()=>e.Set("Funding Assumptions",81,6,new object[,]{{old}}));
                    times["restoredIncremental"]=Time(()=>e.Calculate(0));result["restoredIncremental"]=SpreadProbes(e);
                }
                result["diagnostics"]=e.Diagnostics;
            }
            result["success"]=true;
        }
        catch(Exception ex){result["success"]=false;result["error"]=ex.ToString();Console.WriteLine(ex);}
        result["sourceUnchanged"]=hash==IoTrial.FileHash(input);
        IoTrial.WriteJson(report,result);Console.WriteLine("RESULT "+report);return Object.Equals(result["success"],true)?0:1;
    }
}
