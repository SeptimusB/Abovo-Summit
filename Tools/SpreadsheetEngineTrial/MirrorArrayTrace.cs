using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DX = DevExpress.Spreadsheet;

internal static partial class EngineBenchmark
{
    // Native-only diagnostic: never publish a checkpoint to the real model.
    // Addresses can move during inserts; compare counts at those boundaries,
    // not unadjusted before/after address sets.
    internal static int RunMirrorArrayTrace(string input,string directory,bool preserveArrays=false)
    {
        input=IoTrial.InTrial(input);directory=IoTrial.InTrial(directory);
        Check(!Directory.Exists(directory),"New trace directory required.");
        Directory.CreateDirectory(directory);
        string hash=IoTrial.FileHash(input),output=Path.Combine(directory,"funding-checkpoint.xlsm");
        var times=new Dictionary<string,object>();var results=new Dictionary<string,object>();
        var stages=new List<object>();var watch=Stopwatch.StartNew();
        var report=new Dictionary<string,object>{{"input",input},{"inputHash",hash},{"harnessCompleted",false},
            {"scope","Isolated array-stage diagnosis; no production or original file writes."},{"preserveArrays",preserveArrays},{"timings",times},{"results",results},{"stages",stages}};
        try
        {
            using(var engine=new DxEngine(true))
            {
                Action<string> trace=label=>
                {
                    var state=engine.MirrorArrayManifest();
                    var record=new {stage=label,elapsedMs=watch.ElapsedMilliseconds,dynamicArrays=state.Length};
                    stages.Add(record);
                    IoTrial.WriteJson(Path.Combine(directory,"stage-"+stages.Count.ToString("D2")+".json"),new {record,arrays=state});
                    Console.WriteLine("ARRAY TRACE "+label+" count="+state.Length+" elapsedMs="+watch.ElapsedMilliseconds);
                };
                times["load"]=Time(()=>engine.Open(input));trace("loaded");
                times["initialRebuild"]=Time(()=>engine.Calculate(2));trace("initial-rebuild");
                engine.MirrorFunding(times,results,trace,preserveArrays);
                times["postInsertRebuild"]=Time(()=>engine.Calculate(2));trace("post-insert-rebuild");
                times["saveFullCalculate"]=Time(()=>engine.Calculate(1));trace("save-full-calculate");
                times["serialize"]=Time(()=>engine.MirrorSave(output,preserveArrays));trace("serialized");
                PreserveTrialCustomXml(input,output);
            }
            RunMirrorArrayState(output,Path.Combine(directory,"array-state.json"));
            using(var projection=NewMirrorProjection(true))
            {
                times["gearRebase"]=Time(()=>{projection.Open(output);projection.Calculate(2);});
                report["tests"]=new {correctedGearProbes=Probes(projection)};
            }
            report["structuralReadbackFile"]=output;report["harnessCompleted"]=true;
        }
        catch(Exception ex){report["error"]=ex.ToString();throw;}
        finally
        {
            report["sourceUnchanged"]=hash==IoTrial.FileHash(input);
            IoTrial.WriteJson(Path.Combine(directory,"report.json"),report);
        }
        Console.WriteLine("ARRAY TRACE COMPLETE "+directory);return 0;
    }
    sealed partial class DxEngine
    {
        internal object[] MirrorArrayManifest()
        {
            return book.Worksheets["Transactional DB"].DynamicArrayFormulas.Cast<DX.DynamicArrayFormula>()
                .Select(a=>(object)new {row=a.Range.TopRowIndex,column=a.Range.LeftColumnIndex,rows=a.Range.RowCount,columns=a.Range.ColumnCount}).ToArray();
        }
    }
}
