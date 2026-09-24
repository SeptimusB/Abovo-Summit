using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

internal static partial class EngineBenchmark
{
    public sealed class MirrorValues
    {
        public string Session,BaseHash,Sheet;
        public int Row,Column,Rows,Columns;
        public object[][] Cells;
    }
    static void ValidateMirrorRead(MirrorRequest request)
    {
        Check(!String.IsNullOrWhiteSpace(request.Sheet)&&request.Row>=0&&request.Column>=0&&request.Rows>0&&request.Columns>0,
            "Native result rectangle is invalid.");
        Check((long)request.Rows*request.Columns<=100000&&
            (long)request.Row+request.Rows<=1048576&&(long)request.Column+request.Columns<=16384,
            "Native result rectangle exceeds the bounded transport limit.");
    }
    static bool AcceptMirrorValues(MirrorReply reply,MirrorRequest request,MirrorConfig config,MirrorJournal journal)
    {
        var values=reply.Values;
        return reply.Ok&&reply.Revision==request.Revision&&reply.Revision==journal.Revision&&values!=null&&
            values.Session==config.Session&&values.BaseHash==config.InputHash&&values.Sheet==request.Sheet&&
            values.Row==request.Row&&values.Column==request.Column&&values.Rows==request.Rows&&values.Columns==request.Columns&&
            values.Cells!=null&&values.Cells.Length==request.Rows&&values.Cells.All(row=>row!=null&&row.Length==request.Columns);
    }
    static MirrorRequest SpillRead(int revision,int rows)
    {return new MirrorRequest {Action="read-values",Revision=revision,Sheet="ArrayProbe",Row=1,Column=1,Rows=rows,Columns=1};}

    static void TestNativeMirrorReads(MirrorConfig config,MirrorJournal journal,MirrorClient worker,GearEngine gear,
        Dictionary<string,object> results,Dictionary<string,object> times)
    {
        // The consumer displays a native result snapshot, never writes native
        // result values over formulas in either workbook. Unknown dependencies
        // must use this path; these tests do not certify a Gear eligibility list.
        var request=SpillRead(journal.Revision,5);var reply=worker.Call(request);
        Check(AcceptMirrorValues(reply,request,config,journal),"Fresh native result rejected.");
        for(int i=0;i<5;i++)Check(SameValue(reply.Values.Cells[i][0],(double)i+1),"Native spill fallback differs.");
        times["nativeFallbackFullCalculateAndFiveValues"]=reply.Milliseconds;
        results["nativeFallbackDisplaysExpandedSpillWithoutWorkbookRewrite"]=true;
        Check(!worker.Call(SpillRead(journal.Revision-1,5)).Ok,"Worker accepted stale read revision.");
        var invalid=SpillRead(journal.Revision,100001);
        Check(!worker.Call(invalid).Ok,"Unbounded native read accepted.");
        results["staleAndOversizedNativeReadRequestsRejected"]=true;

        // Simulate a UI edit arriving while a native read is in flight. Only
        // the newer revision may be displayed even when the old result is valid.
        worker.Send(request);
        // Six rows fit before the fixture's merged explanatory row at row 8.
        var edit=new MirrorEdit {Kind="set",Sheet="ArrayProbe",Row=0,Column=3,ValueKind="number",Before=5.0,Value=6.0};
        string path=journal.Append(edit);ApplyMirrorEdit(gear,edit);gear.Calculate(2);
        worker.Send(new MirrorRequest {Action="apply",Journal=path});
        var stale=worker.Read();Check(stale.Ok&&!AcceptMirrorValues(stale,request,config,journal),"Late native result overwrote a newer edit.");
        Check(worker.Read().Ok,"Concurrent native fallback edit failed.");
        request=SpillRead(journal.Revision,6);reply=worker.Call(request);
        Check(AcceptMirrorValues(reply,request,config,journal),"Newer native result rejected: "+reply.Error+"; worker revision="+reply.Revision+"; expected="+journal.Revision);
        for(int i=0;i<6;i++)Check(SameValue(reply.Values.Cells[i][0],(double)i+1),"Newer native spill result differs.");
        string session=reply.Values.Session;reply.Values.Session="another-session";
        Check(!AcceptMirrorValues(reply,request,config,journal),"Cross-session native result accepted.");reply.Values.Session=session;
        results["lateNativeResultsRejectedAfterNewerEdit"]=true;
        results["crossSessionNativeResultRejected"]=true;
        results["nativeResultFallbackScope"]="Explicit requested range, matching session/base/revision, native full calculation; no production routing or automatic dependency classification.";
        if(config.Backend=="devexpress")
        {
            // Keep the discovered blocked-spill edge case as a safety regression.
            // The fixture's row 8 is merged. A future engine may return #SPILL!
            // instead of throwing; either must not manufacture a successful value.
            var blocked=new MirrorEdit {Kind="set",Sheet="ArrayProbe",Row=0,Column=3,ValueKind="number",Before=6.0,Value=7.0};
            string blockedPath=journal.Append(blocked);
            Check(worker.Call(new MirrorRequest {Action="apply",Journal=blockedPath}).Ok,"Blocked-spill setup failed.");
            var failedRead=worker.Call(SpillRead(journal.Revision,7));
            results["blockedSpillNativeRead"]=new {failedRead.Ok,failedRead.Error};
            if(!failedRead.Ok)
            {
                string target=Path.Combine(config.Directory,"model.xlsm"),hash=IoTrial.FileHash(target);
                Check(failedRead.Values==null&&!worker.Call(SpillRead(journal.Revision,1)).Ok,"Faulted native worker published later values.");
                var save=worker.Call(new MirrorRequest {Action="save",Revision=journal.Revision,Target=target,ExpectedHash=hash});
                Check(!save.Ok&&IoTrial.FileHash(target)==hash,"Faulted native calculation replaced the saved file.");
                results["nativeCalculationFailureBlocksReadsAndSave"]=true;
            }
            else Check(failedRead.Values!=null&&Convert.ToString(failedRead.Values.Cells[0][0]).StartsWith("#"),"Blocked spill returned an apparently valid scalar.");
        }
    }
}
