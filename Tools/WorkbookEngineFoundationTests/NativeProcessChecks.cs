using System;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

static class NativeProcessChecks
{
    internal static int[] ExcelIds(){return Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).OrderBy(x=>x).ToArray();}
    internal static async Task RequireOriginalProcesses(int[] before)
    {
        var timer=Stopwatch.StartNew();var after=ExcelIds();
        while(!before.SequenceEqual(after)&&timer.ElapsedMilliseconds<10000){await Task.Delay(100).ConfigureAwait(false);after=ExcelIds();}
        Console.WriteLine("EXCEL_EXIT before="+String.Join(",",before)+" after="+String.Join(",",after)+" waitMs="+timer.ElapsedMilliseconds);
        if(!before.SequenceEqual(after))throw new Exception("Owned Excel cleanup or pre-existing process inventory changed.");
    }
}
