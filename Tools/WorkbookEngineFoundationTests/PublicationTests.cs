using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Abovo.WorkbookEngines;

static class PublicationTests
{
    static int assertions;
    static WorkbookReadArea Cell=>new WorkbookReadArea("Data",0,0,1,1);
    static WorkbookEngineOptions Options(bool publish=true,int timeout=120000,WorkbookEnginePreference engine=WorkbookEnginePreference.DevExpressOnly,bool udf=false)=>new WorkbookEngineOptions(engine,udf,udf,timeout,true,true,publish);
    static void Check(bool yes,string text){if(!yes)throw new Exception(text);Console.WriteLine("PUBLISH PASS "+(++assertions)+" "+text);}
    static string Hash(string path){using(var f=File.OpenRead(path))using(var h=System.Security.Cryptography.SHA256.Create())return BitConverter.ToString(h.ComputeHash(f)).Replace("-","");}
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true,EntryPoint="CreateFileW")]
    static extern SafeFileHandle OpenNative(string p,uint access,uint share,IntPtr security,uint creation,uint flags,IntPtr template);
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)]
    static extern bool CreateHardLinkW(string name,string existing,IntPtr security);
    static void Stream(string file,string name,string text)
    {using(var h=OpenNative(file+":"+name,0x40000000,7,IntPtr.Zero,2,0,IntPtr.Zero)){if(h.IsInvalid)throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());using(var s=new FileStream(h,FileAccess.Write))using(var w=new StreamWriter(s))w.Write(text);}}
    static string Zone(string path)=>(string)typeof(WorkbookEngineOptions).Assembly.GetType("Abovo.WorkbookEngines.ExcelCalculationBackend").GetMethod("ReadInternetZone",BindingFlags.NonPublic|BindingFlags.Static).Invoke(null,new object[]{path});
    static async Task<Exception> Reject(Func<Task> action,string text)
    {try{await action();}catch(Exception e)when(e is IOException||e is ArgumentException||e is InvalidOperationException||e is OperationCanceledException||e is TimeoutException||e is NotSupportedException){Check(true,text+" ["+e.GetType().Name+"]");return e;}throw new Exception("Expected rejection: "+text);}
    static Task<WorkbookPublicationReceipt> Publish(WorkbookCalculationSession s,WorkbookSaveCandidate c,string target,WorkbookPublicationMode mode,Action<string> hook=null,CancellationToken cancel=default(CancellationToken))
    {
        try{return (Task<WorkbookPublicationReceipt>)typeof(WorkbookCalculationSession).GetMethod("PublishAndCloseCoreAsync",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(s,new object[]{c,target,mode,cancel,hook});}
        catch(TargetInvocationException ex){System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw();throw;}
    }
    internal sealed class Backend:IWorkbookCandidateBackend
    {
        public string Name=>"Publication test";public string Version=>"1";public string Source;public int ThreadId;public Action Closing;public WorkbookEngineOptions SeenOptions;
        void Own(){if(Thread.CurrentThread.ManagedThreadId!=ThreadId)throw new Exception("Wrong owner thread");}
        public void OpenReadOnly(string p,WorkbookEngineOptions o){ThreadId=Thread.CurrentThread.ManagedThreadId;Source=p;SeenOptions=o;}
        public void Calculate(WorkbookCalculationKind k){Own();}
        public WorkbookValueBlock Read(WorkbookReadArea a){Own();return new WorkbookValueBlock(a,new object[,]{{10d}});}
        public WorkbookCellState ReadCell(WorkbookReadArea a){Own();return new WorkbookCellState(10d,"","0",false,true,false,false,false);}
        public void WriteValue(WorkbookReadArea a,object v){Own();throw new Exception("Should reject before native write");}
        public void ExportCopy(string p){Own();File.Copy(Source,p);}
        public WorkbookCandidateReadback ReadCopy(string p,IList<WorkbookReadArea> a,IList<WorkbookReadArea> c){Own();return new WorkbookCandidateReadback(a.Select(Read),c.Select(ReadCell));}
        public void Dispose(){Own();Closing?.Invoke();}
    }
    internal sealed class Trial
    {
        public string Root,Source,Original;public Backend Backend;public WorkbookCalculationSession Session;public WorkbookSaveCandidate Candidate;public WorkbookCellSnapshot CellState;
        public static async Task<Trial> Create(string fixture,bool enabled=true,int timeout=120000,Action<string> seed=null)
        {
            var t=new Trial();t.Root=Path.Combine(Path.GetDirectoryName(fixture),"pub-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(t.Root);
            t.Source=Path.Combine(t.Root,"input"+Path.GetExtension(fixture));File.Copy(fixture,t.Source);t.Original=Hash(t.Source);t.Backend=new Backend();
            seed?.Invoke(t.Source);
            t.Session=await WorkbookCalculationSession.OpenAsync(t.Source,Options(enabled,timeout),default(CancellationToken),_=>t.Backend);
            t.CellState=await t.Session.CaptureCellAsync(0,Cell);var result=await t.Session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{Cell});
            t.Candidate=await t.Session.CreateSaveCandidateAsync(result,t.Root,new[]{t.CellState});return t;
        }
        public string Target=>Path.Combine(Root,"saved"+Path.GetExtension(Source));
        public string Stage=>XDocument.Load(Directory.GetFiles(Root,"intent.xml",SearchOption.AllDirectories).Single()).Root.Element("Candidate").Value;
        public string Backup=>XDocument.Load(Directory.GetFiles(Root,"intent.xml",SearchOption.AllDirectories).Single()).Root.Element("Backup").Value;
    }
    internal static async Task Run(string fixture)
    {
        var t=await Trial.Create(fixture,false);
        try{await Reject(()=>t.Session.PublishAndCloseAsync(t.Candidate,t.Target,WorkbookPublicationMode.CreateNew),"publication requires separate opt-in");Check(t.Session.IsCurrentCandidate(t.Candidate)&&!File.Exists(t.Target),"disabled publication leaves owner and disk intact");}finally{await t.Session.CloseAsync();}
        t=await Trial.Create(fixture);
        try
        {
            await Reject(()=>t.Session.PublishAndCloseAsync(t.Candidate,"relative.xlsm",WorkbookPublicationMode.CreateNew),"relative destination rejected");
            await Reject(()=>t.Session.PublishAndCloseAsync(t.Candidate,t.Target+".xlsx",WorkbookPublicationMode.CreateNew),"format conversion rejected");
            await Reject(()=>t.Session.PublishAndCloseAsync(t.Candidate,t.Source,WorkbookPublicationMode.CreateNew),"Save As cannot overwrite source");
            await Reject(()=>t.Session.PublishAndCloseAsync(t.Candidate,t.Target,WorkbookPublicationMode.ReplaceSource),"replacement cannot name unrelated file");
            await Reject(()=>t.Session.PublishAndCloseAsync(t.Candidate,t.Target,(WorkbookPublicationMode)99),"invalid mode rejected");
            File.WriteAllText(t.Target,"other writer");string external=Hash(t.Target);
            await Reject(()=>t.Session.PublishAndCloseAsync(t.Candidate,t.Target,WorkbookPublicationMode.CreateNew),"existing Save As target rejected");
            Check(Hash(t.Target)==external&&t.Session.IsCurrentCandidate(t.Candidate),"existing target untouched; session reusable");
            using(var cancel=new CancellationTokenSource()){cancel.Cancel();await Reject(()=>Publish(t.Session,t.Candidate,t.Source,WorkbookPublicationMode.ReplaceSource,cancel:cancel.Token),"pre-cancelled publication rejected");}
            t.Session.InvalidateResults();await Reject(()=>t.Session.PublishAndCloseAsync(t.Candidate,t.Source,WorkbookPublicationMode.ReplaceSource),"newer revision cannot acknowledge old save");
        }finally{await t.Session.CloseAsync();}
        foreach(var mode in new[]{WorkbookPublicationMode.CreateNew,WorkbookPublicationMode.ReplaceSource})
        {
            t=await Trial.Create(fixture);
            try
            {
                string target=mode==WorkbookPublicationMode.CreateNew?t.Target:t.Source;
                var receipt=await Publish(t.Session,t.Candidate,target,mode,p=>
                {
                    if(p=="Prepared")
                    {
                        Reject(()=>t.Session.ApplyValueAsync(t.CellState,20d,WorkbookValuePermission.UnlockedCell),"edit cannot overlap publication").GetAwaiter().GetResult();
                        Reject(()=>Task.Run(()=>t.Session.InvalidateResults()),"invalidation cannot overlap publication").GetAwaiter().GetResult();
                        Reject(()=>t.Session.PublishAndCloseAsync(t.Candidate,t.Target,WorkbookPublicationMode.CreateNew),"second publication cannot overlap").GetAwaiter().GetResult();
                    }
                    if(p=="BeforeCommit")
                    {
                        Reject(()=>Task.Run(()=>File.AppendAllText(t.Source,"change")),"source lease guards write during commit").GetAwaiter().GetResult();
                        Reject(()=>Task.Run(()=>Directory.Move(t.Root,t.Root+"-moved")),"parent directory stays pinned during commit").GetAwaiter().GetResult();
                    }
                });
                Check(receipt.SessionId==t.Session.SessionId&&receipt.Revision==0&&receipt.Hash==Hash(target),"published receipt names exact bytes and revision: "+mode);
                Check(!t.Session.IsCurrentCandidate(t.Candidate),"publication consumes old native session");
                Check(File.Exists(receipt.JournalPath)&&File.Exists(Path.Combine(Path.GetDirectoryName(receipt.JournalPath),"published.xml")),"intent and successful acknowledgement recorded");
                Check(mode==WorkbookPublicationMode.CreateNew?Hash(t.Source)==t.Original:Hash(receipt.BackupPath)==t.Original,"original bytes retained: "+mode);
            }finally{await t.Session.CloseAsync();}
        }
        foreach(string fault in new[]{"cancel-prepared","staging-write","staged-bytes","source-changed","source-missing","target-appeared","rename-fails","competing-writer","late-cancel","ack-fails","source-locked","source-zone","staged-zone","source-stream","stage-stream","boundary-zone","secondary-marker"})
        {
            t=await Trial.Create(fixture);FileStream external=null;
            using(var cancel=new CancellationTokenSource())
            try
            {
                bool create=fault=="target-appeared";string target=create?t.Target:t.Source;
                Action<string> hook=p=>
                {
                    if(p=="Staging"&&fault=="staging-write")throw new IOException("injected full/unwritable storage");
                    if(p=="Prepared"&&fault=="cancel-prepared")cancel.Cancel();
                    if(p=="Prepared"&&fault=="staged-bytes")File.AppendAllText(t.Stage,"tamper");
                    if(p=="Closed"&&fault=="source-changed")File.AppendAllText(t.Source,"external");
                    if(p=="Closed"&&fault=="source-missing")File.Move(t.Source,t.Source+".external");
                    if(p=="Closed"&&fault=="source-locked")external=new FileStream(t.Source,FileMode.Open,FileAccess.Read,FileShare.Read);
                    if(p=="Closed"&&fault=="source-zone")Stream(t.Source,"Zone.Identifier","[ZoneTransfer]\r\nZoneId=3\r\n");
                    if(p=="Prepared"&&fault=="staged-zone")Stream(t.Stage,"Zone.Identifier","[ZoneTransfer]\r\nZoneId=3\r\n");
                    if(p=="Closed"&&fault=="source-stream")Stream(t.Source,"Unreviewed","preserve this");
                    if(p=="Prepared"&&fault=="stage-stream")Stream(t.Stage,"Unreviewed","preserve this");
                    if(p=="BeforeCommit"&&fault=="boundary-zone")Stream(t.Stage,"Zone.Identifier","[ZoneTransfer]\r\nZoneId=3\r\n");
                    if(p=="Closed"&&fault=="secondary-marker")Stream(t.Source,"MBAM.Zone.Identifier","unexpected");
                    if(p=="BeforeCommit"&&fault=="target-appeared")File.WriteAllText(t.Target,"competitor");
                    if(p=="OriginalRetained"&&fault=="rename-fails")throw new IOException("injected rename failure");
                    if(p=="OriginalRetained"&&fault=="competing-writer")File.WriteAllText(t.Source,"competitor");
                    if(p=="OriginalRetained"&&fault=="late-cancel")cancel.Cancel();
                    if(p=="Published"&&fault=="ack-fails")throw new IOException("injected acknowledgement failure");
                };
                if(fault=="late-cancel")
                {var receipt=await Publish(t.Session,t.Candidate,target,WorkbookPublicationMode.ReplaceSource,hook,cancel.Token);Check(Hash(target)==receipt.Hash&&Hash(receipt.BackupPath)==t.Original,"late cancellation cannot erase successful publication acknowledgement");}
                else
                {
                    var e=await Reject(()=>Publish(t.Session,t.Candidate,target,create?WorkbookPublicationMode.CreateNew:WorkbookPublicationMode.ReplaceSource,hook,cancel.Token),"injected "+fault);
                    if(fault=="competing-writer")Check(File.ReadAllText(t.Source)=="competitor"&&Hash(t.Backup)==t.Original,"competing file never overwritten; verified original retained for recovery");
                    else if(fault=="target-appeared")Check(File.ReadAllText(t.Target)=="competitor"&&Hash(t.Source)==t.Original,"late Save As collision preserves both files");
                    else if(fault=="source-changed")Check(Hash(t.Source)!=t.Original&&!File.Exists(t.Backup),"external source change is not replaced");
                    else if(fault=="source-missing")Check(!File.Exists(t.Source)&&Hash(t.Source+".external")==t.Original,"externally renamed source is not recreated");
                    else if(fault=="ack-fails")Check(((WorkbookPublicationException)e).Published&&Hash(t.Source)==t.Candidate.Hash&&Hash(t.Backup)==t.Original,"written-but-unacknowledged is explicit, with previous bytes retained");
                    else Check(Hash(t.Source)==t.Original,"failure leaves or restores original: "+fault);
                    if(fault=="cancel-prepared")Check(t.Session.IsCurrentCandidate(t.Candidate),"cancellation before close keeps native owner reusable");
                }
            }
            finally{external?.Dispose();await t.Session.CloseAsync();}
        }
        t=await Trial.Create(fixture,seed:p=>{Stream(p,"Zone.Identifier","[ZoneTransfer]\r\nZoneId=3\r\n");Stream(p,"MBAM.Zone.Identifier","[ZoneTransfer]\r\nZoneId=3\r\n\0");});
        try
        {
            var receipt=await t.Session.PublishAndCloseAsync(t.Candidate,t.Source,WorkbookPublicationMode.ReplaceSource);
            Check(Zone(t.Source)=="[ZoneTransfer]\r\nZoneId=3\r\n"&&Zone(receipt.BackupPath)==Zone(t.Source),"downloaded-file marker preserved on publication and original backup");
            foreach(string path in new[]{t.Source,receipt.BackupPath})using(var h=OpenNative(path+":MBAM.Zone.Identifier",0x80000000,7,IntPtr.Zero,3,0,IntPtr.Zero))using(var file=new FileStream(h,FileAccess.Read))using(var reader=new StreamReader(file))Check(reader.ReadToEnd()=="[ZoneTransfer]\r\nZoneId=3\r\n\0","opaque secondary provenance marker including NUL is preserved");
        }finally{await t.Session.CloseAsync();}
        t=await Trial.Create(fixture,seed:p=>
        {
            var access=new System.Security.AccessControl.FileSecurity();access.SetAccessRuleProtection(true,false);
            access.AddAccessRule(new System.Security.AccessControl.FileSystemAccessRule(System.Security.Principal.WindowsIdentity.GetCurrent().User,System.Security.AccessControl.FileSystemRights.FullControl,System.Security.AccessControl.AccessControlType.Allow));
            File.SetAccessControl(p,access);
        });
        try
        {
            string acl=File.GetAccessControl(t.Source).GetSecurityDescriptorSddlForm(System.Security.AccessControl.AccessControlSections.Access);
            var receipt=await t.Session.PublishAndCloseAsync(t.Candidate,t.Source,WorkbookPublicationMode.ReplaceSource);
            Check(File.GetAccessControl(t.Source).GetSecurityDescriptorSddlForm(System.Security.AccessControl.AccessControlSections.Access)==acl,"source discretionary access policy preserved on replacement");
            Check(Hash(receipt.BackupPath)==t.Original,"custom-permission original retained");
        }finally{await t.Session.CloseAsync();}
        foreach(bool linked in new[]{true,false})
        {
            t=await Trial.Create(fixture);
            try
            {
                var e=await Reject(()=>Publish(t.Session,t.Candidate,t.Source,WorkbookPublicationMode.ReplaceSource,p=>
                {if(p=="Closed"){if(linked){if(!CreateHardLinkW(t.Source+".linked",t.Source,IntPtr.Zero))throw new Exception("Hardlink fixture failed");}else File.SetAttributes(t.Source,FileAttributes.ReadOnly);}}),linked?"hardlinked source rejected":"read-only source rejected");
                Check(Hash(t.Source)==t.Original,"unsupported source not changed");
            }finally{File.SetAttributes(t.Source,FileAttributes.Normal);await t.Session.CloseAsync();}
        }
        t=await Trial.Create(fixture);
        t.Backend.Closing=()=>{throw new IOException("injected native close failure");};
        await Reject(()=>t.Session.PublishAndCloseAsync(t.Candidate,t.Source,WorkbookPublicationMode.ReplaceSource),"native cleanup failure forbids publication");
        Check(Hash(t.Source)==t.Original,"failed cleanup cannot replace source");
        t=await Trial.Create(fixture,timeout:250);
        using(var release=new ManualResetEventSlim())
        {
            t.Backend.Closing=()=>release.Wait();
            await Reject(()=>t.Session.PublishAndCloseAsync(t.Candidate,t.Source,WorkbookPublicationMode.ReplaceSource),"hung native cleanup wait is bounded");
            release.Set();await t.Session.NativeCleanupCompletion;Check(Hash(t.Source)==t.Original,"late cleanup cannot trigger abandoned publication");
        }
        Console.WriteLine("PUBLICATION ASSERTIONS="+assertions);
    }
    internal static async Task Native(string original,bool agl)
    {
        var before=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).ToArray();string originalHash=Hash(original);
        string root=Path.Combine(Path.GetDirectoryName(original),"pub-native-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);
        var fixtures=agl?new[]{original}:new[]{".xlsx",".xlsm",".xlsb"}.Select(ext=>SaveCandidateTests.Fixture(root,ext)).ToArray();
        foreach(string fixture in fixtures)foreach(var engine in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired})
        foreach(var mode in agl?new[]{WorkbookPublicationMode.ReplaceSource}:new[]{WorkbookPublicationMode.CreateNew,WorkbookPublicationMode.ReplaceSource})
        {
            string path=Path.Combine(root,Guid.NewGuid().ToString("N")+Path.GetExtension(fixture));File.Copy(fixture,path);string baseline=Hash(path);
            var s=await WorkbookCalculationSession.OpenAsync(path,Options(engine:engine,udf:agl));
            try
            {
                var cell=agl?new WorkbookReadArea("Funding Assumptions",81,6,1,1):Cell;
                var reads=agl?new[]{new WorkbookReadArea("Detailed Comp Inc - Trad View",0,0,98,42),new WorkbookReadArea("Financial Position - Trad View",0,0,67,43),new WorkbookReadArea("Cashflow detailed",0,0,51,94),new WorkbookReadArea("Check Sheet",0,0,63,11),new WorkbookReadArea("Development Expenditure",0,0,613,54)}:new[]{new WorkbookReadArea("Data",0,0,3,4)};
                var input=await s.CaptureCellAsync(0,cell);var edit=await s.ApplyValueAsync(input,Convert.ToDouble(input.State.Value)+300d,WorkbookValuePermission.UnlockedCell,reads);
                var candidate=await s.CreateSaveCandidateAsync(edit.Results,root,new[]{edit.After});string target=mode==WorkbookPublicationMode.ReplaceSource?path:Path.Combine(root,"new-"+Path.GetFileName(path));
                var timer=Stopwatch.StartNew();var receipt=await s.PublishAndCloseAsync(candidate,target,mode);
                Console.WriteLine("PUBLICATION_NATIVE engine="+s.EngineName+" mode="+mode+" format="+Path.GetExtension(path)+" handoffMs="+timer.ElapsedMilliseconds);
                Check(Hash(target)==candidate.Hash&&receipt.Revision==1,"native publication writes candidate byte-for-byte");
                Check(mode==WorkbookPublicationMode.ReplaceSource?Hash(receipt.BackupPath)==baseline:Hash(path)==baseline,"native original remains intact or recoverable");
                var other=engine==WorkbookEnginePreference.ExcelRequired?WorkbookEnginePreference.DevExpressOnly:WorkbookEnginePreference.ExcelRequired;
                var check=await WorkbookCalculationSession.OpenAsync(target,Options(engine:other,udf:agl));
                try{var snap=await check.CaptureCellAsync(0,cell);Check(Equals(snap.State.Value,edit.After.State.Value),"published edited input reopens in other engine");SaveCandidateTests.Compare(edit.Results,await check.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,reads),"published "+s.EngineName+" -> "+check.EngineName);}finally{await check.CloseAsync();}
            }finally{await s.CloseAsync();}
        }
        Check(Hash(original)==originalHash,"native publication tests never modify supplied baseline");
        var clock=Stopwatch.StartNew();int[] after;
        do{after=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).ToArray();if(!after.Except(before).Any())break;await Task.Delay(100);}while(clock.ElapsedMilliseconds<10000);
        Check(before.OrderBy(n=>n).SequenceEqual(after.OrderBy(n=>n)),"publication closes only owned Excel processes");
        Console.WriteLine("PUBLICATION_NATIVE ASSERTIONS="+assertions);
    }
}
