using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Abovo.WorkbookEngines;
using System.IO.Compression;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

static class Program
{
    static int assertions;
    static WorkbookReadArea One => new WorkbookReadArea("Data",0,0,1,1);
    static void Check(bool yes,string message)
    {
        if(!yes)throw new Exception(message);
        Console.WriteLine("PASS "+(++assertions)+" "+message);
    }
    static string Hash(string path)
    {using(var input=File.OpenRead(path))using(var hash=SHA256.Create())return BitConverter.ToString(hash.ComputeHash(input));}
    static async Task Fails(Func<Task> action,string message)
    {
        try{await action();}catch(Exception e){if(e is InvalidOperationException||e is ArgumentException||e is ObjectDisposedException||e is OperationCanceledException||e is InvalidDataException){Check(true,message);return;}throw;}
        throw new Exception("Expected rejection: "+message);
    }
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true,EntryPoint="CreateFileW")]
    static extern SafeFileHandle CreateNativeFile(string path,uint access,uint share,IntPtr security,uint creation,uint flags,IntPtr template);
    static void Security(string path)
    {
        var type=typeof(WorkbookEngineOptions).Assembly.GetType("Abovo.WorkbookEngines.ExcelCalculationBackend");
        var inspect=type.GetMethod("InspectPackage",BindingFlags.NonPublic|BindingFlags.Static);
        var read=type.GetMethod("ReadInternetZone",BindingFlags.NonPublic|BindingFlags.Static);
        foreach(var scenario in new[]{"internet","restricted","intranet","connections","xlm"})
        {
            var target=Path.Combine(Path.GetDirectoryName(path),"engine-security-"+Guid.NewGuid().ToString("N")+".xlsm");
            File.Copy(path,target);
            try
            {
                string zone=null;
                if(scenario=="internet"||scenario=="restricted"||scenario=="intranet")
                {
                    zone="[ZoneTransfer]\r\nZoneId="+(scenario=="internet"?"3":scenario=="restricted"?"4":"1")+"\r\n";
                    using(var handle=CreateNativeFile(target+":Zone.Identifier",0x40000000,7,IntPtr.Zero,2,0,IntPtr.Zero))
                    {
                        if(handle.IsInvalid)throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                        using(var stream=new FileStream(handle,FileAccess.Write))using(var writer=new StreamWriter(stream))writer.Write(zone);
                    }
                    Check((string)read.Invoke(null,new object[]{target})==zone,"native security marker read: "+scenario);
                }
                else
                {
                    using(var zip=ZipFile.Open(target,ZipArchiveMode.Update))
                    using(var writer=new StreamWriter(zip.CreateEntry(scenario=="connections"?"xl/connections.xml":"xl/macrosheets/sheet1.xml").Open()))writer.Write("<test/>");
                }
                bool rejected=false;
                try{inspect.Invoke(null,new object[]{target});}catch(TargetInvocationException e){if(!(e.InnerException is InvalidOperationException))throw;rejected=true;}
                Check(rejected==(scenario!="intranet"),"security preflight decision: "+scenario);
                if(zone!=null)Check((string)read.Invoke(null,new object[]{target})==zone,"security marker preserved: "+scenario);
            }
            finally{File.Delete(target);}
        }
    }
    sealed class FakeBackend:IWorkbookCalculationBackend
    {
        public string Name{get;set;}="Fake";
        public string Version=>"1";
        public bool FailOpen,FailCalc,Disposed,Pump,Release,WrongArea;
        public int Opens,Calculates,OwnerThread,Depth,MaxDepth;
        public readonly TaskCompletionSource<bool> Started=new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        public void OpenReadOnly(string path,WorkbookEngineOptions options){OwnerThread=Thread.CurrentThread.ManagedThreadId;Opens++;Check(Thread.CurrentThread.GetApartmentState()==ApartmentState.STA,"backend opens on STA");if(FailOpen)throw new InvalidOperationException("Simulated Excel unavailable");}
        void Owner(){Check(Thread.CurrentThread.ManagedThreadId==OwnerThread,"native object stays on owning thread");}
        public void Calculate(WorkbookCalculationKind kind)
        {
            Owner();Depth++;MaxDepth=Math.Max(MaxDepth,Depth);Calculates++;Started.TrySetResult(true);
            try{if(FailCalc)throw new InvalidOperationException("Simulated native failure");while(Pump&&!Volatile.Read(ref Release)){Application.DoEvents();Thread.Sleep(1);}}
            finally{Depth--;}
        }
        public WorkbookValueBlock Read(WorkbookReadArea area){Owner();var data=new object[area.Rows,area.Columns];data[0,0]=42d;return new WorkbookValueBlock(WrongArea?new WorkbookReadArea("Wrong",area.Row,area.Column,area.Rows,area.Columns):area,data);}
        public void Dispose(){Owner();Disposed=true;}
    }
    static WorkbookEngineOptions Options(WorkbookEnginePreference preference=WorkbookEnginePreference.Automatic,bool udf=false,bool vba=false)
    {return new WorkbookEngineOptions(preference,vba,udf);}
    static async Task Safety(string path)
    {
        Check(new WorkbookEngineOptions().Preference==WorkbookEnginePreference.Automatic,"default preference is compatible Excel first");
        var conversion=typeof(WorkbookEngineOptions).Assembly.GetType("Abovo.WorkbookEngines.ExcelCalculationBackend").GetMethod("ConvertValue",BindingFlags.NonPublic|BindingFlags.Static);
        Check(((WorkbookCellError)conversion.Invoke(null,new object[]{unchecked((int)0x800A07FD)})).Text=="#SPILL!","modern spill HRESULT remains an error");
        Check(conversion.Invoke(null,new object[]{unchecked((int)0x800A0802)}) is WorkbookCellError,"unknown modern error cannot become a financial number");
        Check(conversion.Invoke(null,new object[]{new System.Runtime.InteropServices.ErrorWrapper(2007)}) is WorkbookCellError,"wrapped CVErr remains typed error");
        Check(conversion.Invoke(null,new object[]{-2146826243d}) is double,"ordinary numeric value is not mistaken for an error");
        Check(new WorkbookReadArea("Data",1048575,16383,1,1).Address=="XFD1048576:XFD1048576","last Excel cell address");
        await Fails(()=>Task.Run(()=>new WorkbookReadArea("Data",0,0,100001,1)),"oversized read rejected");
        await Fails(()=>Task.Run(()=>new WorkbookReadArea("Data",int.MaxValue,0,1,1)),"overflow row rejected");
        var raw=new object[,]{{1d}};var block=new WorkbookValueBlock(One,raw);raw[0,0]=2d;
        Check((double)block.ValueAt(0,0)==1d,"value blocks detached from mutable backend buffers");
        await Fails(()=>Task.Run(()=>new WorkbookValueBlock(One,new object[,]{{new object()}})),"native/non-value objects rejected at transport boundary");

        var excel=new FakeBackend{FailOpen=true,Name="Excel"};var dx=new FakeBackend{Name="DevExpress"};
        var session=await WorkbookCalculationSession.OpenAsync(path,Options(),default(CancellationToken),kind=>kind==WorkbookEnginePreference.ExcelRequired?excel:dx);
        Check(session.EngineName=="DevExpress"&&session.FallbackReason.Contains("Simulated")&&excel.Disposed,"opening failure selects fallback only after Excel cleanup");
        var result=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{One});
        Check(session.IsCurrent(result)&&result.Blocks.Count==1&&(double)result.Blocks[0].ValueAt(0,0)==42d,"detached result has exact session/hash/revision");
        session.InvalidateResults();Check(!session.IsCurrent(result),"new revision invalidates old result");
        await Fails(()=>session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{One}),"stale request rejected before calculation");
        Check(dx.Calculates==1,"stale request did not touch native engine");
        using(var canceled=new CancellationTokenSource()){canceled.Cancel();await Fails(()=>session.CalculateAndReadAsync(1,WorkbookCalculationKind.Full,new[]{One},canceled.Token),"pre-cancelled calculation rejected");}
        Check(dx.Calculates==1,"cancelled queued request did not calculate");
        dx.FailCalc=true;await Fails(()=>session.CalculateAndReadAsync(1,WorkbookCalculationKind.Full,new[]{One}),"native calculation failure surfaced");
        dx.FailCalc=false;await Fails(()=>session.CalculateAndReadAsync(1,WorkbookCalculationKind.Full,new[]{One}),"failed session cannot publish subsequent results");
        Check(excel.Opens==1&&dx.Opens==1,"failure never switched engines mid-session");
        await session.CloseAsync();await session.CloseAsync();Check(dx.Disposed&&!session.IsCurrent(result),"close is idempotent and invalidates all results");

        var only=new FakeBackend();int factories=0;
        session=await WorkbookCalculationSession.OpenAsync(path,Options(WorkbookEnginePreference.DevExpressOnly),default(CancellationToken),kind=>{factories++;Check(kind==WorkbookEnginePreference.DevExpressOnly,"DevExpress-only never probes Excel");return only;});
        await session.CloseAsync();Check(factories==1,"one authoritative backend selected");
        excel=new FakeBackend{FailOpen=true};factories=0;
        await Fails(()=>WorkbookCalculationSession.OpenAsync(path,Options(WorkbookEnginePreference.ExcelRequired),default(CancellationToken),kind=>{factories++;return excel;}),"Excel-required fails instead of masking capability failure");
        Check(factories==1&&excel.Disposed,"failed required Excel cleaned up");

        var pump=new FakeBackend{Pump=true};
        session=await WorkbookCalculationSession.OpenAsync(path,Options(),default(CancellationToken),kind=>pump);
        var first=session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{One});await pump.Started.Task;
        var second=session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{One});
        await Task.Delay(100);Check(pump.Calculates==1&&pump.MaxDepth==1,"COM/message pumping cannot re-enter a second workbook operation");
        session.InvalidateResults();Volatile.Write(ref pump.Release,true);
        await Fails(()=>first,"late reply rejected after newer revision");
        await Fails(()=>second,"queued old revision rejected after newer revision");
        Check(pump.Calculates==1,"stale queued request never calculated");
        result=await session.CalculateAndReadAsync(1,WorkbookCalculationKind.Full,new[]{One});Check(session.IsCurrent(result),"new revision can calculate after stale-result rejection");
        await session.CloseAsync();

        var mismatch=new FakeBackend{WrongArea=true};
        session=await WorkbookCalculationSession.OpenAsync(path,Options(),default(CancellationToken),kind=>mismatch);
        await Fails(()=>session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{One}),"wrong worksheet block rejected before publication");
        mismatch.WrongArea=false;
        await Fails(()=>session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{One}),"bad transport response faults the session");
        await session.CloseAsync();

        using(var canceled=new CancellationTokenSource())
        {
            canceled.Cancel();factories=0;
            await Fails(()=>WorkbookCalculationSession.OpenAsync(path,Options(),canceled.Token,kind=>{factories++;return new FakeBackend();}),"pre-cancelled opening stops cleanly");
            Check(factories==0,"pre-cancelled opening never creates native objects");
        }
        await Fails(()=>WorkbookCalculationSession.OpenAsync(Path.ChangeExtension(path,".txt"),Options()),"unsupported file type rejected before native opening");

        var firstBackend=new FakeBackend();var otherBackend=new FakeBackend();
        session=await WorkbookCalculationSession.OpenAsync(path,Options(),default(CancellationToken),kind=>firstBackend);
        var other=await WorkbookCalculationSession.OpenAsync(path,Options(),default(CancellationToken),kind=>otherBackend);
        try
        {
            result=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{One});
            Check(!other.IsCurrent(result),"same path and revision cannot confuse different sessions");
            await Fails(()=>session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,Enumerable.Repeat(One,129)),"too many transport blocks rejected");
        }
        finally{await session.CloseAsync();await other.CloseAsync();}

        pump=new FakeBackend{Pump=true};session=await WorkbookCalculationSession.OpenAsync(path,Options(),default(CancellationToken),kind=>pump);
        using(var canceled=new CancellationTokenSource())
        {
            first=session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{One},canceled.Token);await pump.Started.Task;
            canceled.Cancel();Check(!first.IsCompleted,"in-flight native calculation is not unsafely aborted");
            Volatile.Write(ref pump.Release,true);await Fails(()=>first,"cancellation discards completed native result");
        }
        result=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{One});Check(session.IsCurrent(result),"safe cancellation retains a reusable session");
        await session.CloseAsync();
    }
    static async Task Native(string path,bool model)
    {
        // Rectangles match the prior independently verified AGL engine probes.
        var requests=model?new[]{new WorkbookReadArea("Detailed Comp Inc - Trad View",0,0,98,42),
            new WorkbookReadArea("Financial Position - Trad View",0,0,67,43),
            new WorkbookReadArea("Cashflow detailed",0,0,51,94),new WorkbookReadArea("Check Sheet",0,0,63,11),
            new WorkbookReadArea("Development Expenditure",0,0,613,54)}:
            new[]{new WorkbookReadArea("Data",0,0,4096,10),new WorkbookReadArea("Summary",0,1,3,1)};
        var before=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).ToArray();
        WorkbookCalculationResult reference=null;
        foreach(var preference in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.Automatic})
        {
            WorkbookCalculationSession session=null;
            try
            {
                var timer=Stopwatch.StartNew();session=await WorkbookCalculationSession.OpenAsync(path,Options(preference,model,model));
                Console.WriteLine("OPEN engine="+session.EngineName+" version="+session.EngineVersion+" ms="+timer.ElapsedMilliseconds+" fallback="+session.FallbackReason);
                Check(session.EngineName==(preference==WorkbookEnginePreference.Automatic?"Excel":"DevExpress"),"real requested engine selected");
                var result=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Rebuild,requests);
                Console.WriteLine("CALC engine="+session.EngineName+" ms="+result.CalculationMilliseconds+" transferMs="+result.TransferMilliseconds);
                Check(session.IsCurrent(result),"real calculated result accepted");
                await ResultGridTests.NativeResult(session,result);
                if(reference==null)reference=result;
                else
                {
                    int cells=0,differences=0;
                    for(int b=0;b<result.Blocks.Count;b++)for(int r=0;r<result.Blocks[b].Area.Rows;r++)for(int c=0;c<result.Blocks[b].Area.Columns;c++)
                    {
                        object a=reference.Blocks[b].ValueAt(r,c),v=result.Blocks[b].ValueAt(r,c);cells++;
                        bool equal=a is double&&v is double?Math.Abs((double)a-(double)v)<=Math.Max(0.000001,Math.Abs((double)a)*1e-10):
                            (a is WorkbookCellError&&v is WorkbookCellError?((WorkbookCellError)a).Text==((WorkbookCellError)v).Text:Equals(a,v));
                        if(!equal){differences++;if(differences<=8)Console.WriteLine("DIFFERENCE block="+b+" row="+r+" col="+c+" dx="+a+" excel="+v);}
                    }
                    Check(differences==0,"native parity across "+cells+" cells (not full-model certification)");
                }
                var warm=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,requests);
                Console.WriteLine("WARM engine="+session.EngineName+" ms="+warm.CalculationMilliseconds+" transferMs="+warm.TransferMilliseconds);
            }
            finally{if(session!=null)await session.CloseAsync();}
        }
        var exitTimer=Stopwatch.StartNew();int[] after;
        do
        {
            await Task.Delay(100);
            after=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).ToArray();
        }while(exitTimer.ElapsedMilliseconds<10000&&!before.OrderBy(x=>x).SequenceEqual(after.OrderBy(x=>x)));
        Console.WriteLine("EXCEL_EXIT ms="+exitTimer.ElapsedMilliseconds+" before="+String.Join(",",before)+" after="+String.Join(",",after));
        Check(before.OrderBy(x=>x).SequenceEqual(after.OrderBy(x=>x)),"owned Excel exited; pre-existing Excel processes unchanged");
        var fallback=await WorkbookCalculationSession.OpenAsync(path,Options(WorkbookEnginePreference.Automatic,true,false));
        try{Check(fallback.EngineName=="DevExpress"&&fallback.FallbackReason.Contains("disabled"),"real VBA-disabled policy falls back before use");}
        finally{await fallback.CloseAsync();}
    }
    static async Task Run(string[] args)
    {
        var before=Hash(args[1]);
        if(args[2].StartsWith("crash-",StringComparison.Ordinal)){await RecoveryTests.CrashChild(args[1],args[2].Substring(6));throw new Exception("Crash fixture returned unexpectedly");}
        if(args[2]=="recovery"||args[2]=="reopen-native"||args[2]=="reopen-agl")
        {
            if(args[2]=="recovery")await RecoveryTests.Run(args[1]);else await RecoveryTests.Native(args[1],args[2]=="reopen-agl");
            Check(Hash(args[1])==before,"source bytes unchanged");return;
        }
        if(args[2]=="publish"||args[2]=="publish-native"||args[2]=="publish-agl")
        {
            if(args[2]=="publish")await PublicationTests.Run(args[1]);else await PublicationTests.Native(args[1],args[2]=="publish-agl");
            Check(Hash(args[1])==before,"source bytes unchanged");return;
        }
        if(args[2]=="safety")await Safety(args[1]);else if(args[2]=="security")Security(args[1]);else if(args[2]=="projection")ProjectionProbe.Run();else if(args[2]=="grid")await ResultGridTests.Run(args[1]);else if(args[2]=="deadline")await DeadlineTests.Run(args[1]);else if(args[2]=="edit")await ValueEditTests.Run(args[1]);else if(args[2]=="edit-native")await ValueEditTests.Native(args[1], false);else if(args[2]=="edit-agl")await ValueEditTests.Native(args[1], true);else if(args[2]=="save")await SaveCandidateTests.Run(args[1]);else if(args[2]=="save-native")await SaveCandidateTests.Native(args[1],false);else if(args[2]=="save-agl")await SaveCandidateTests.Native(args[1],true);else await Native(args[1],args[2]=="agl");
        Check(Hash(args[1])==before,"source bytes unchanged");Console.WriteLine("ASSERTIONS="+assertions);
    }
    [STAThread]
    static int Main(string[] args)
    {
        if(args.Length!=3||!new[]{"safety","security","synthetic","agl","projection","grid","deadline","edit","edit-native","edit-agl","save","save-native","save-agl","publish","publish-native","publish-agl","recovery","reopen-native","reopen-agl","crash-Prepared","crash-OriginalRetained","crash-Published"}.Contains(args[2])){Console.Error.WriteLine("APP_BIN PRIVATE_WORKBOOK safety|security|synthetic|agl|projection|grid|deadline|edit|edit-native|edit-agl|save|save-native|save-agl|publish|publish-native|publish-agl|recovery|reopen-native|reopen-agl");return 2;}
        AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string file=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");if(!File.Exists(file))file=Path.Combine(args[0],new AssemblyName(e.Name).Name+".exe");return File.Exists(file)?Assembly.LoadFrom(file):null;};
        try{Run(args).GetAwaiter().GetResult();return 0;}catch(Exception e){Console.Error.WriteLine(e);return 1;}
    }
}
