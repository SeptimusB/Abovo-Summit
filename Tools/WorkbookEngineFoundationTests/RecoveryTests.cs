using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Abovo.WorkbookEngines;

static class RecoveryTests
{
    static int assertions;
    static void Check(bool ok,string label){if(!ok)throw new Exception(label);Console.WriteLine("RECOVERY PASS "+(++assertions)+" "+label);}
    static string Hash(string path){using(var f=File.OpenRead(path))using(var h=System.Security.Cryptography.SHA256.Create())return BitConverter.ToString(h.ComputeHash(f)).Replace("-","");}
    static async Task Reject(Func<Task> action,string label)
    {try{await action();}catch(Exception e)when(e is IOException||e is InvalidDataException||e is InvalidOperationException||e is ArgumentException||e is OperationCanceledException||e is XmlException||e is NotSupportedException){Check(true,label+" ["+e.GetType().Name+"]");return;}throw new Exception("Expected rejection: "+label);}
    static Task<T> Invoke<T>(Type type,object instance,string method,params object[] arguments)
    {try{return (Task<T>)type.GetMethod(method,BindingFlags.NonPublic|BindingFlags.Static|BindingFlags.Instance).Invoke(instance,arguments);}catch(TargetInvocationException e){System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(e.InnerException).Throw();throw;}}
    static string Journal(PublicationTests.Trial t)=>Directory.GetFiles(t.Root,"intent.xml",SearchOption.AllDirectories).Single();
    static Task<WorkbookRecoveryInspection> Inspect(PublicationTests.Trial t)=>WorkbookPublicationRecovery.InspectAsync(Journal(t),t.Source);
    static Task<WorkbookOriginalRestoration> Restore(WorkbookRecoveryInspection i,Action<string> hook,CancellationToken cancel=default(CancellationToken))=>Invoke<WorkbookOriginalRestoration>(typeof(WorkbookPublicationRecovery),null,"RestoreCoreAsync",i,cancel,hook);
    static Task<WorkbookCalculationSession> Reopen(WorkbookPublicationReceipt r,Func<WorkbookEnginePreference,IWorkbookCalculationBackend> factory,CancellationToken cancel=default(CancellationToken))=>Invoke<WorkbookCalculationSession>(typeof(WorkbookEngineOptions).Assembly.GetType("Abovo.WorkbookEngines.WorkbookPublicationReopen"),null,"OpenAsync",r,cancel,factory);
    static void Mark(string path,string text)
    {typeof(PublicationTests).GetMethod("Stream",BindingFlags.NonPublic|BindingFlags.Static).Invoke(null,new object[]{path,"Zone.Identifier",text});}
    static async Task<PublicationTests.Trial> Prepared(string fixture,bool gap=true)
    {
        var t=await PublicationTests.Trial.Create(fixture);
        try
        {
            await Reject(()=>Invoke<WorkbookPublicationReceipt>(typeof(WorkbookCalculationSession),t.Session,"PublishAndCloseCoreAsync",t.Candidate,t.Source,WorkbookPublicationMode.ReplaceSource,CancellationToken.None,new Action<string>(p=>{if(p=="Prepared")throw new IOException("stop before close");})),"prepared interruption retained");
        }finally{await t.Session.CloseAsync();}
        if(gap)File.Move(t.Source,t.Backup);
        return t;
    }
    internal static async Task Run(string fixture)
    {
        var t=await Prepared(fixture,false);string stage=t.Stage;
        var inspected=await Inspect(t);Check(inspected.State==WorkbookRecoveryState.Prepared,"original intact is not a completed save");
        await Reject(()=>WorkbookPublicationRecovery.RestoreOriginalAsync(inspected),"no restoration over existing original");
        Check(Hash(t.Source)==t.Original&&File.Exists(stage),"inspection leaves original and candidate unchanged");
        File.Move(t.Source,t.Backup);inspected=await Inspect(t);
        Check(inspected.State==WorkbookRecoveryState.OriginalRetained,"missing-name interruption recognized");
        var restored=await WorkbookPublicationRecovery.RestoreOriginalAsync(inspected);
        Check(restored.CompletionRecorded&&Hash(t.Source)==t.Original&&!File.Exists(t.Backup),"verified original restored to vacant name");
        Check(File.Exists(stage)&&Hash(stage)==t.Candidate.Hash,"unsaved candidate retained after rollback");
        Check((await Inspect(t)).State==WorkbookRecoveryState.Prepared,"post-restore inspection recognizes original, not saved revision");
        await Reject(()=>WorkbookPublicationRecovery.RestoreOriginalAsync(inspected),"stale restore cannot run twice");

        t=await Prepared(fixture);inspected=await Inspect(t);
        using(var cancel=new CancellationTokenSource())
        {cancel.Cancel();await Reject(()=>WorkbookPublicationRecovery.RestoreOriginalAsync(inspected,cancel.Token),"cancel before recovery leaves gap untouched");}
        Check(!File.Exists(t.Source)&&Hash(t.Backup)==t.Original,"cancellation preserves retained original");
        await Reject(()=>Restore(inspected,p=>{if(p=="BeforeRestore")File.WriteAllText(t.Source,"competing writer");}),"writer arriving after inspection is never overwritten");
        Check(File.ReadAllText(t.Source)=="competing writer"&&Hash(t.Backup)==t.Original,"late competing file and original both preserved");
        Check((await Inspect(t)).State==WorkbookRecoveryState.Conflict,"occupied target classified as conflict");

        t=await Prepared(fixture);inspected=await Inspect(t);
        using(var cancel=new CancellationTokenSource())
        {restored=await Restore(inspected,p=>{if(p=="Restored")cancel.Cancel();},cancel.Token);Check(restored.CompletionRecorded&&Hash(t.Source)==t.Original,"late cancellation does not discard successful restoration");}
        t=await Prepared(fixture);inspected=await Inspect(t);
        restored=await Restore(inspected,p=>{if(p=="Restored")throw new IOException("injected advisory record failure");});
        Check(!restored.CompletionRecorded&&Hash(t.Source)==t.Original,"post-rename record failure reports restored, not retry");

        foreach(string change in new[]{"journal","candidate","backup","marker","missing","directory"})
        {
            t=await Prepared(fixture);inspected=await Inspect(t);stage=t.Stage;
            if(change=="journal"){var x=XDocument.Load(Journal(t));x.Root.Element("Revision").Value="999";x.Save(Journal(t));}
            if(change=="candidate")File.AppendAllText(stage,"changed");
            if(change=="backup")File.AppendAllText(t.Backup,"changed");
            if(change=="marker")Mark(t.Backup,"[ZoneTransfer]\r\nZoneId=3\r\n");
            if(change=="missing")File.Delete(stage);
            if(change=="directory")Directory.CreateDirectory(t.Source);
            await Reject(()=>WorkbookPublicationRecovery.RestoreOriginalAsync(inspected),"changed "+change+" blocks stale recovery");
            Check(!File.Exists(t.Source)&&File.Exists(t.Backup),"changed evidence never publishes or loses original: "+change);
        }
        foreach(string malformed in new[]{"legacy","duplicate","traversal","target","unknown","hash","session","revision","oversized","dtd"})
        {
            t=await Prepared(fixture);string retained=t.Backup;var x=XDocument.Load(Journal(t));
            switch(malformed)
            {
                case "legacy":x.Root.SetAttributeValue("version",1);break;
                case "duplicate":x.Root.Add(new XElement("Target",t.Source));break;
                case "traversal":x.Root.Element("Backup").Value=Path.Combine(t.Root,"outside.xlsm");break;
                case "target":x.Root.Element("Target").Value=t.Target;break;
                case "unknown":x.Root.Add(new XElement("Execute","something"));break;
                case "hash":x.Root.Element("SourceHash").Value="invalid";break;
                case "session":x.Root.Element("Session").Value=Guid.Empty.ToString();break;
                case "revision":x.Root.Element("Revision").Value="-1";break;
                case "oversized":x.Root.Add(new XElement("Padding",new string('x',20000)));break;
            }
            if(malformed=="dtd")File.WriteAllText(Journal(t),"<!DOCTYPE SummitPublication [<!ENTITY external SYSTEM 'file:///C:/Windows/win.ini'>]><SummitPublication>&external;</SummitPublication>");else x.Save(Journal(t));
            await Reject(()=>Inspect(t),"malformed "+malformed+" journal rejected");
            Check(!File.Exists(t.Source)&&Hash(retained)==t.Original,"invalid journal leaves original untouched: "+malformed);
        }
        t=await Prepared(fixture);
        await Reject(()=>WorkbookPublicationRecovery.InspectAsync(Journal(t),t.Target),"journal cannot choose a different caller target");
        string outside=Path.Combine(t.Root,"outside.xml");File.Copy(Journal(t),outside);
        await Reject(()=>WorkbookPublicationRecovery.InspectAsync(outside,t.Source),"journal outside generated sibling folder rejected");
        using(var locked=new FileStream(t.Backup,FileMode.Open,FileAccess.ReadWrite,FileShare.None))await Reject(()=>Inspect(t),"locked original is not interpreted as absent");
        Check(Hash(t.Backup)==t.Original,"locked inspection preserves bytes");

        foreach(var mode in new[]{WorkbookPublicationMode.CreateNew,WorkbookPublicationMode.ReplaceSource})
        {
            t=await PublicationTests.Trial.Create(fixture);string target=mode==WorkbookPublicationMode.CreateNew?t.Target:t.Source;
            var receipt=await t.Session.PublishAndCloseAsync(t.Candidate,target,mode);
            inspected=await WorkbookPublicationRecovery.InspectAsync(receipt.JournalPath,target);
            Check(inspected.State==WorkbookRecoveryState.PublishedAcknowledged,"published completion recognized: "+mode);
            string complete=Path.Combine(Path.GetDirectoryName(receipt.JournalPath),"published.xml");File.Delete(complete);
            Check((await WorkbookPublicationRecovery.InspectAsync(receipt.JournalPath,target)).State==WorkbookRecoveryState.PublishedUnacknowledged,"lost acknowledgement does not repeat save: "+mode);
            await Reject(()=>WorkbookPublicationRecovery.RestoreOriginalAsync(inspected),"published workbook is never automatically rolled back: "+mode);
            int factories=0;var newBackend=new PublicationTests.Backend();var reopened=await Reopen(receipt,kind=>{factories++;Check(kind==WorkbookEnginePreference.DevExpressOnly,"reopen retains original engine preference");return newBackend;});
            try
            {
                Check(factories==1&&reopened.SourceHash==receipt.Hash&&reopened.SessionId!=receipt.SessionId&&reopened.Revision==0,"reopen rebases to exact saved bytes and a new session");
                Check(ReferenceEquals(newBackend.SeenOptions,t.Backend.SeenOptions),"reopen passes identical immutable security and trial options");
                Check(!reopened.IsCurrentCandidate(t.Candidate),"old candidate cannot acknowledge new session edits");
                reopened.InvalidateResults();Check(reopened.Revision==1&&receipt.Revision==0,"new revisions do not change saved receipt");
            }finally{await reopened.CloseAsync();}
            File.AppendAllText(target,"external edit");factories=0;
            await Reject(()=>Reopen(receipt,_=>{factories++;return new PublicationTests.Backend();}),"changed published file rejected before native open");
            Check(factories==0,"changed file cannot run VBA through stale receipt");
        }
        t=await PublicationTests.Trial.Create(fixture);
        var saved=await t.Session.PublishAndCloseAsync(t.Candidate,t.Target,WorkbookPublicationMode.CreateNew);
        Mark(t.Target,"[ZoneTransfer]\r\nZoneId=3\r\n");
        await Reject(()=>Reopen(saved,_=>throw new Exception("must not reach backend")),"changed security marker blocks verified reopen");
        Check(Hash(t.Target)==saved.Hash,"refused reopen does not undo successful save");
        t=await PublicationTests.Trial.Create(fixture);
        saved=await t.Session.PublishAndCloseAsync(t.Candidate,t.Target,WorkbookPublicationMode.CreateNew);
        using(var canceled=new CancellationTokenSource())
        {canceled.Cancel();await Reject(()=>Reopen(saved,_=>throw new Exception("must not create native owner"),canceled.Token),"cancelled reopen creates no owner");}
        await Reject(()=>Reopen(saved,_=>throw new InvalidOperationException("injected native opening failure")),"failed reopen leaves saved file intact");
        Check(Hash(t.Target)==saved.Hash,"failed reopen never restores old source over successful save");
        var retry=await Reopen(saved,_=>new PublicationTests.Backend());await retry.CloseAsync();
        Check(Hash(t.Target)==saved.Hash,"failed reopen can retry under unchanged policy");
        bool disposed=false;var lateMarker=new PublicationTests.Backend{Closing=()=>disposed=true};
        await Reject(()=>Reopen(saved,_=>{Mark(t.Target,"[ZoneTransfer]\r\nZoneId=1\r\n");return lateMarker;}),"marker change during native open rejects new owner");
        Check(disposed&&Hash(t.Target)==saved.Hash,"failed post-open verification closes owner without changing saved bytes");
        foreach(string boundary in new[]{"Prepared","OriginalRetained","Published"})await AbruptExit(fixture,boundary);
        Console.WriteLine("RECOVERY ASSERTIONS="+assertions);
    }

    internal static async Task CrashChild(string fixture,string boundary)
    {
        // A separate owned test process, fake native backend, private file only.
        // Environment.Exit skips finally/disposal, reproducing lost process state.
        var t=await PublicationTests.Trial.Create(fixture);
        Console.WriteLine("CRASH_ROOT="+t.Root);Console.Out.Flush();
        await Invoke<WorkbookPublicationReceipt>(typeof(WorkbookCalculationSession),t.Session,"PublishAndCloseCoreAsync",t.Candidate,t.Source,WorkbookPublicationMode.ReplaceSource,CancellationToken.None,new Action<string>(p=>{if(p==boundary)Environment.Exit(73);}));
    }
    static async Task AbruptExit(string fixture,string boundary)
    {
        var start=new ProcessStartInfo(Assembly.GetExecutingAssembly().Location,"\""+Path.GetDirectoryName(typeof(WorkbookEngineOptions).Assembly.Location)+"\" \""+fixture+"\" crash-"+boundary)
        {UseShellExecute=false,CreateNoWindow=true,WindowStyle=ProcessWindowStyle.Hidden,RedirectStandardOutput=true,RedirectStandardError=true};
        using(var child=Process.Start(start))
        {
            var output=child.StandardOutput.ReadToEndAsync();var errors=child.StandardError.ReadToEndAsync();
            if(!await Task.Run(()=>child.WaitForExit(30000))){child.Kill();throw new TimeoutException("Owned crash fixture did not exit.");}
            Check(child.ExitCode==73,"owned process exited abruptly at "+boundary+" "+await errors);
            string root=(await output).Split(new[]{'\r','\n'},StringSplitOptions.RemoveEmptyEntries).Single(line=>line.StartsWith("CRASH_ROOT=",StringComparison.Ordinal)).Substring(11);
            string target=Path.Combine(root,"input"+Path.GetExtension(fixture));
            string journal=Directory.GetFiles(root,"intent.xml",SearchOption.AllDirectories).Single();
            var inspected=await WorkbookPublicationRecovery.InspectAsync(journal,target);
            var expected=boundary=="Prepared"?WorkbookRecoveryState.Prepared:boundary=="OriginalRetained"?WorkbookRecoveryState.OriginalRetained:WorkbookRecoveryState.PublishedUnacknowledged;
            Check(inspected.State==expected,"on-disk state recognized after actual process exit: "+boundary);
            if(boundary=="OriginalRetained")
            {var result=await WorkbookPublicationRecovery.RestoreOriginalAsync(inspected);Check(Hash(result.Path)==Hash(fixture)&&File.Exists(result.CandidatePath),"abrupt-exit original restored and unsaved candidate retained");}
            else Check(File.Exists(target),"abrupt-exit target retained: "+boundary);
        }
    }

    internal static async Task Native(string original,bool agl)
    {
        var before=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).ToArray();string originalHash=Hash(original);
        string root=Path.Combine(Path.GetDirectoryName(original),"reopen-native-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);
        var fixtures=agl?new[]{original}:new[]{".xlsx",".xlsm",".xlsb"}.Select(ext=>SaveCandidateTests.Fixture(root,ext)).ToArray();
        foreach(string fixture in fixtures)foreach(var engine in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired})
        {
            string path=Path.Combine(root,Guid.NewGuid().ToString("N")+Path.GetExtension(fixture));File.Copy(fixture,path);
            var options=new WorkbookEngineOptions(engine,agl,agl,120000,true,true,true);
            var session=await WorkbookCalculationSession.OpenAsync(path,options);
            var cell=agl?new WorkbookReadArea("Funding Assumptions",81,6,1,1):new WorkbookReadArea("Data",0,0,1,1);
            var reads=agl?new[]{new WorkbookReadArea("Detailed Comp Inc - Trad View",0,0,98,42),new WorkbookReadArea("Financial Position - Trad View",0,0,67,43),new WorkbookReadArea("Cashflow detailed",0,0,51,94),new WorkbookReadArea("Check Sheet",0,0,63,11),new WorkbookReadArea("Development Expenditure",0,0,613,54)}:new[]{new WorkbookReadArea("Data",0,0,3,4)};
            WorkbookValueEditReceipt edit;
            WorkbookPublicationReceipt receipt;
            try
            {
                var input=await session.CaptureCellAsync(0,cell);
                edit=await session.ApplyValueAsync(input,Convert.ToDouble(input.State.Value)+300d,WorkbookValuePermission.UnlockedCell,reads);
                var candidate=await session.CreateSaveCandidateAsync(edit.Results,root,new[]{edit.After});
                receipt=await session.PublishAndCloseAsync(candidate,path,WorkbookPublicationMode.ReplaceSource);
            }finally{await session.CloseAsync();}
            Check((await WorkbookPublicationRecovery.InspectAsync(receipt.JournalPath,path)).State==WorkbookRecoveryState.PublishedAcknowledged,"native publication journal verified: "+engine);
            var timer=Stopwatch.StartNew();session=await receipt.ReopenAsync();
            Console.WriteLine("REOPEN_NATIVE engine="+engine+" format="+Path.GetExtension(path)+" reopenMs="+timer.ElapsedMilliseconds);
            string second=Path.Combine(root,"continued-"+Path.GetFileName(path));
            try
            {
                Check(session.SessionId!=receipt.SessionId&&session.SourceHash==receipt.Hash&&session.Revision==0,"native new baseline identity");
                SaveCandidateTests.Compare(edit.Results,await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,reads),"verified reopen "+engine);
                var input=await session.CaptureCellAsync(0,cell);
                edit=await session.ApplyValueAsync(input,Convert.ToDouble(input.State.Value)+100d,WorkbookValuePermission.UnlockedCell,reads);
                var candidate=await session.CreateSaveCandidateAsync(edit.Results,root,new[]{edit.After});
                receipt=await session.PublishAndCloseAsync(candidate,second,WorkbookPublicationMode.CreateNew);
                Check(Hash(second)==receipt.Hash,"second edit can publish from rebased owner");
            }finally{await session.CloseAsync();}
            var other=engine==WorkbookEnginePreference.ExcelRequired?WorkbookEnginePreference.DevExpressOnly:WorkbookEnginePreference.ExcelRequired;
            session=await WorkbookCalculationSession.OpenAsync(second,new WorkbookEngineOptions(other,agl,agl));
            try{SaveCandidateTests.Compare(edit.Results,await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,reads),"continued edit other engine "+other);}finally{await session.CloseAsync();}
        }
        Check(Hash(original)==originalHash,"native reopen tests preserve supplied baseline");
        var clock=Stopwatch.StartNew();int[] after;
        do{after=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).ToArray();if(!after.Except(before).Any())break;await Task.Delay(100);}while(clock.ElapsedMilliseconds<10000);
        Check(before.OrderBy(n=>n).SequenceEqual(after.OrderBy(n=>n)),"reopen closes only owned Excel processes");
        Console.WriteLine("REOPEN_NATIVE ASSERTIONS="+assertions);
    }
}
