using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Gear = SpreadsheetGear;

// Read-side isolation only. Never calculate, save, rewrite, or remove workbook content.
internal static class GearImportProbe
{
    internal static int Run(string input,string report)
    {
        input=IoTrial.InTrial(input);report=IoTrial.InTrial(report);
        if(File.Exists(report))throw new InvalidOperationException("New report required.");
        string hash=IoTrial.FileHash(input);SpreadsheetGearTrial.Activate();
        var tests=new List<object>();
        foreach(bool objects in new[]{true,false})foreach(bool vba in new[]{true,false})
        {
            var set=Gear.Factory.GetWorkbookSet(CultureInfo.InvariantCulture);
            set.Calculation=Gear.Calculation.Manual;set.BackgroundCalculation=false;
            set.CalculationOnDemand=false;set.EnableWebService=false;
            set.ReadObjects=objects;set.ReadVBA=vba;var time=Stopwatch.StartNew();
            try
            {
                using(var stream=File.OpenRead(input))
                {
                    var book=set.Workbooks.OpenFromStream(stream);
                    tests.Add(new {readObjects=objects,readVba=vba,opened=true,milliseconds=time.Elapsed.TotalMilliseconds,sheets=book.Worksheets.Count,names=book.Names.Count});
                    Console.WriteLine("OPEN objects="+objects+" VBA="+vba+" "+time.ElapsedMilliseconds+" ms");
                }
            }
            catch(Exception e)
            {
                tests.Add(new {readObjects=objects,readVba=vba,opened=false,milliseconds=time.Elapsed.TotalMilliseconds,error=e.ToString()});
                Console.WriteLine("FAIL objects="+objects+" VBA="+vba+" "+e.GetType().Name+": "+e.Message);
            }
            finally{while(set.Workbooks.Count>0)set.Workbooks[0].Close();}
        }
        bool unchanged=hash==IoTrial.FileHash(input);
        IoTrial.WriteJson(report,new {input,inputHash=hash,unchanged,tests,scope="Read-only import flags; no calculation, VBA execution, workbook save or production changes"});
        if(!unchanged)throw new InvalidOperationException("Input changed.");return 0;
    }
}

internal static partial class EngineBenchmark
{
    internal static int RunGearProjection(string input,string report)
    {
        input=IoTrial.InTrial(input);report=IoTrial.InTrial(report);Check(!File.Exists(report),"New report required.");
        string hash=IoTrial.FileHash(input);var timings=new Dictionary<string,object>();
        using(var engine=new GearEngine(true))
        {
            engine.ReadCalculationProjection();
            timings["loadWithoutObjects"]=Time(()=>engine.Open(input));
            timings["fullRebuild"]=Time(()=>engine.Calculate(2));
            IoTrial.WriteJson(report,new {success=true,engine="spreadsheetgear-no-objects",inputHash=hash,input,
                sourceUnchanged=hash==IoTrial.FileHash(input),timingsMs=timings,calculatedProbes=Probes(engine),
                scope="Read-only calculation projection diagnostic. Objects skipped in memory; VBA retained; no workbook saved or production setting changed."});
        }
        Console.WriteLine("PROJECTION "+report);return 0;
    }
    sealed partial class GearEngine
    {
        bool calculationProjection;
        internal void ReadCalculationProjection()
        {
            Check(book==null,"Select the calculation-only import policy before opening a workbook.");
            set.ReadObjects=false;set.ReadVBA=true;calculationProjection=true;
        }
        internal void RequireFullWorkbookForSave()
        {
            Check(!calculationProjection,"A calculation-only Gear projection cannot be saved; use the authoritative worker.");
        }
    }
}
