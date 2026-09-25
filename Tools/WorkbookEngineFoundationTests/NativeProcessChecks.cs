using System;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using Abovo.WorkbookEngines;

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
    internal static async Task RequireOwnedProcesses(IEnumerable<WorkbookCalculationSession> sessions)
    {
        var owners=sessions.Where(s=>s!=null&&s.NativeProcess!=null).Select(s=>s.NativeProcess).ToArray();
        await RequireOwnedIdentities(owners).ConfigureAwait(false);
    }
    internal static async Task RequireOwnedIdentities(IEnumerable<WorkbookNativeProcessIdentity> identities)
    {
        var owners=identities.Where(p=>p!=null).ToArray();
        // Quit acknowledges shutdown before Windows necessarily removes the
        // process. Retain a bounded grace period and identify only our owners;
        // never pass cleanup because a foreign process appeared/disappeared.
        Console.WriteLine("OWNED_EXCEL_WAIT "+String.Join(";",owners.Select(p=>p.ProcessId+"@"+p.StartedUtc.ToString("O"))));
        Func<WorkbookNativeProcessIdentity,bool> alive=identity=>{
            try{using(var process=Process.GetProcessById(identity.ProcessId))return !process.HasExited&&process.StartTime.ToUniversalTime()==identity.StartedUtc;}
            catch(ArgumentException){return false;}
        };
        var timer=Stopwatch.StartNew();
        while(owners.Any(alive)&&timer.ElapsedMilliseconds<30000)await Task.Delay(100).ConfigureAwait(false);
        if(owners.Any(alive))throw new Exception("A tracked native Excel owner did not exit within 30 seconds: "+String.Join(";",owners.Where(alive).Select(p=>p.ProcessId+"@"+p.StartedUtc.ToString("O"))));
        Console.WriteLine("OWNED_EXCEL_EXIT tracked="+owners.Length+" waitMs="+timer.ElapsedMilliseconds+"; unrelated user processes ignored");
    }
}
