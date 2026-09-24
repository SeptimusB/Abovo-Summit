using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using DX = DevExpress.Spreadsheet;
using Gear = SpreadsheetGear;

// Research only. No Summit UI integration, original-file writes or Trust Center changes.
// The two engines never share a workbook object or a thread. All worker calls are serial.
internal static partial class EngineBenchmark
{
    public sealed class MirrorConfig
    {
        public string Backend, Input, InputHash, Directory, Session;
        public bool Model;
    }
    public sealed class MirrorEdit
    {
        public int Schema = 1, Revision, Row, Column;
        public string Session, BaseHash, PreviousHash, Kind, Sheet, ValueKind;
        public object Before, Value;
    }
    public sealed class MirrorRequest
    {
        public string Action, Journal, Target, ExpectedHash;
        public int Revision;
        public bool FailBeforePublish;
        public string Sheet;
        public int Row,Column,Rows,Columns;
    }
    public sealed class MirrorReply
    {
        public bool Ok, Duplicate;
        public int Revision;
        public string Error, Path, Hash;
        public double Milliseconds;
        public object Data;
        public MirrorValues Values;
    }
    static string DigestText(string value)
    { using(var sha=SHA256.Create())return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value))).Replace("-",""); }
    static string SessionFile(MirrorConfig c,string path)
    {
        path=IoTrial.InTrial(path);
        Check(path.StartsWith(c.Directory+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase),"Path outside this trial session.");
        Check(!String.Equals(path,c.Input,StringComparison.OrdinalIgnoreCase),"Immutable baseline cannot be an output.");return path;
    }
    static void DurableJson(string path,object data)
    {
        path=IoTrial.InTrial(path);var bytes=Encoding.UTF8.GetBytes(IoTrial.Json.Serialize(data));
        string pending=path+".pending";
        using(var f=new FileStream(pending,FileMode.CreateNew,FileAccess.Write,FileShare.None)) { f.Write(bytes,0,bytes.Length);f.Flush(true); }
        File.Move(pending,path); // Incomplete writes are never visible as committed journal records.
    }
    static bool SameValue(object a,object b)
    {
        if(a is string&&((string)a).Length==0)a=null;if(b is string&&((string)b).Length==0)b=null;
        if(a==null||b==null)return a==b;
        if(!(a is string)&&!(a is bool)&&!(b is string)&&!(b is bool))return Equal(Convert.ToDouble(a),Convert.ToDouble(b));
        return Object.Equals(a,b);
    }
    static object TypedValue(MirrorEdit e)
    {
        switch(e.ValueKind) {
            case "number":return Convert.ToDouble(e.Value,CultureInfo.InvariantCulture);
            case "date":return DateTime.ParseExact((string)e.Value,"yyyy-MM-dd",CultureInfo.InvariantCulture).ToOADate();
            case "text":return (string)e.Value;
            case "boolean":return Convert.ToBoolean(e.Value,CultureInfo.InvariantCulture);
            case "blank":return null;
            default:throw new InvalidOperationException("Unknown journal value type.");
        }
    }
    static void ApplyMirrorEdit(Engine engine,MirrorEdit edit)
    {
        Check(edit.Kind=="set","Unsupported edit.");
        Check(SameValue(engine.Get(edit.Sheet,edit.Row,edit.Column,1,1)[0,0],edit.Before),"Edit precondition differs; replay stopped.");
        bool editable=engine is DxEngine ? ((DxEngine)engine).MirrorEditable(edit) : engine is GearEngine ? ((GearEngine)engine).MirrorEditable(edit) : ((ExcelEngine)engine).MirrorEditable(edit);
        Check(editable,"Trial refuses a formula or locked cell.");
        using(engine.Protected(edit.Sheet)?engine.TemporarilyUnprotect(edit.Sheet,TrialProtection.ReadExistingCredential()):null)
            engine.Set(edit.Sheet,edit.Row,edit.Column,new object[,]{{TypedValue(edit)}});
    }
    static Dictionary<string,string> ImportantParts(string file)
    {
        var parts=new Dictionary<string,string>();using(var zip=ZipFile.OpenRead(file))
        {
            foreach(var e in zip.Entries.Where(e=>e.FullName.EndsWith("vbaProject.bin",StringComparison.OrdinalIgnoreCase)))parts[e.FullName]=EntryHash(e);
            foreach(var props in zip.Entries.Where(e=>e.FullName.StartsWith("customXml/itemProps",StringComparison.Ordinal)&&e.FullName.EndsWith(".xml")))
            {
                var document=ReadMirrorXml(props);string id=document.Root.Attributes().Single(a=>a.Name.LocalName=="itemID").Value;
                var matching=zip.Entries.Where(e=>e.FullName.StartsWith("customXml/_rels/",StringComparison.Ordinal)).Where(e=>ReadMirrorXml(e).Root.Elements().Any(r=>(string)r.Attribute("Target")==Path.GetFileName(props.FullName))).ToArray();
                Check(matching.Length==1,"Ambiguous custom XML property relationship.");
                string payload="customXml/"+Path.GetFileName(matching[0].FullName).Replace(".rels","");var data=zip.GetEntry(payload);Check(data!=null,"Custom XML relationship has no payload.");
                parts.Add("customXml:"+id+":payload",EntryHash(data));parts.Add("customXml:"+id+":properties",EntryHash(props));
            }
        }
        return parts;
    }
    static XDocument ReadMirrorXml(ZipArchiveEntry entry)
    {using(var stream=entry.Open())using(var reader=XmlReader.Create(stream,new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null}))return XDocument.Load(reader);}
    static string EntryHash(ZipArchiveEntry entry)
    {using(var stream=entry.Open())using(var hash=SHA256.Create())return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-","");}
    static object ValidateMirrorPackage(string baseline,string candidate)
    {
        var before=ImportantParts(baseline);var after=ImportantParts(candidate);
        // Both native serializers can regenerate compiled VBA containers. Retain
        // presence here and require the separate module-source/form audit, rather
        // than falsely treating binary inequality as an executable source change.
        Check(before.Count==after.Count&&before.All(p=>after.ContainsKey(p.Key)&&(after[p.Key]==p.Value||p.Key.EndsWith("vbaProject.bin"))),"VBA/custom XML package identity or payload changed; candidate not published.");
        bool vbaBinaryUnchanged=before.Where(p=>p.Key.EndsWith("vbaProject.bin")).All(p=>after[p.Key]==p.Value);
        using(var source=ZipFile.OpenRead(baseline))using(var target=ZipFile.OpenRead(candidate))
        {
            bool hadMetadata=source.GetEntry("xl/metadata.xml")!=null;
            Check(!hadMetadata||target.GetEntry("xl/metadata.xml")!=null,"Dynamic array metadata part missing; candidate not published.");
            Check(target.GetEntry("xl/workbook.xml")!=null,"Candidate is not an XLSM package.");
            return new {preservedCustomXmlPayloadsAndProperties=before.Count(p=>p.Key.StartsWith("customXml:")),vbaBinaryUnchanged,requiresIndependentVbaAndFormValidation=!vbaBinaryUnchanged,metadataPartRetained=hadMetadata,scope="Private trial package gate, not financial or Excel round-trip certification. Excel may regenerate VBA form attributes; audit separately."};
        }
    }
    static void PreserveTrialCustomXml(string baseline,string candidate)
    {
        // The prototype's journal is external and has no XML/history mutation command.
        // ONLY in that deliberately limited case are these original payloads still current.
        // Production must serialize the latest Summit model XML, never blindly restore it.
        using(var source=ZipFile.OpenRead(baseline))using(var target=ZipFile.Open(candidate,ZipArchiveMode.Update))
        {
            var originals=source.Entries.Where(e=>e.FullName.StartsWith("customXml/",StringComparison.Ordinal)).ToArray();
            Check(target.Entries.Count(e=>e.FullName.StartsWith("customXml/",StringComparison.Ordinal))==originals.Length,"Custom XML package geometry changed unexpectedly.");
            foreach(var entry in originals)
            {
                var existing=target.GetEntry(entry.FullName);Check(existing!=null,"Custom XML part was removed.");existing.Delete();
                using(var from=entry.Open())using(var to=target.CreateEntry(entry.FullName).Open())from.CopyTo(to);
            }
        }
    }
    internal static int MirrorWorker(string configPath)
    {
        Console.OutputEncoding=new UTF8Encoding(false);
        var requests=new StreamReader(Console.OpenStandardInput(),Encoding.UTF8,true);
        var config=IoTrial.Json.Deserialize<MirrorConfig>(File.ReadAllText(IoTrial.InTrial(configPath)));
        Check(new[]{"excel","devexpress"}.Contains(config.Backend),"Unknown worker backend.");
        Check(IoTrial.FileHash(IoTrial.InTrial(config.Input))==config.InputHash,"Baseline hash differs.");
        int revision=0;string chain=config.InputHash;var seen=new Dictionary<int,string>();bool faulted=false;
        var load=Stopwatch.StartNew();
        using(var engine=Create(config.Backend,config.Model))
        {
            engine.Open(config.Input);engine.Calculate(2);
            Reply(new MirrorReply {Ok=true,Milliseconds=load.Elapsed.TotalMilliseconds,Data=new {engine.Version,backend=config.Backend}});
            string line;while((line=requests.ReadLine())!=null)
            {
                var timer=Stopwatch.StartNew();var reply=new MirrorReply {Revision=revision};
                try
                {
                    var request=IoTrial.Json.Deserialize<MirrorRequest>(line);
                    if(request.Action=="exit") {reply.Ok=true;Reply(reply);break;}
                    Check(!faulted,"Worker mutation failed; discard worker and replay immutable baseline.");
                    if(request.Action=="apply")
                    {
                        string file=SessionFile(config,request.Journal);string raw=File.ReadAllText(file);string hash=DigestText(raw);
                        var edit=IoTrial.Json.Deserialize<MirrorEdit>(raw);
                        Check(edit.Schema==1&&edit.Session==config.Session&&edit.BaseHash==config.InputHash,"Journal belongs to a different model/session/schema.");
                        if(edit.Revision<=revision) {Check(seen.ContainsKey(edit.Revision)&&seen[edit.Revision]==hash,"Conflicting duplicate command.");reply.Duplicate=true;}
                        else
                        {
                            Check(edit.Revision==revision+1&&edit.PreviousHash==chain,"Journal gap or hash-chain mismatch.");
                            try
                            {
                                if(edit.Kind=="set")ApplyMirrorEdit(engine,edit);
                                else if(edit.Kind=="funding10")
                                {
                                    Check(config.Model,"Funding requires the reviewed AGL model fixture.");
                                    var timings=new Dictionary<string,object>();var results=new Dictionary<string,object>();
                                    if(engine is ExcelEngine)((ExcelEngine)engine).FundingVba(timings,results);
                                    else ((DxEngine)engine).MirrorFunding(timings,results);
                                    engine.Calculate(2);reply.Data=new {timings,results};
                                }
                                else throw new InvalidOperationException("Unsupported journal command.");
                            }
                            catch {faulted=true;throw;}
                            seen.Add(edit.Revision,hash);chain=hash;revision=edit.Revision;
                        }
                    }
                    else if(request.Action=="save")
                    {
                        Check(request.Revision==revision,"Save revision has not been fully applied.");
                        string target=SessionFile(config,request.Target);
                        Check(IoTrial.FileHash(target)==request.ExpectedHash,"Destination changed externally; no overwrite.");
                        string candidate=SessionFile(config,Path.Combine(config.Directory,"candidate-"+Guid.NewGuid().ToString("N")+".xlsm"));
                        try
                        {
                            engine.Calculate(1);
                            if(engine is ExcelEngine)((ExcelEngine)engine).SaveFunding(candidate);
                            else {((DxEngine)engine).MirrorSave(candidate);PreserveTrialCustomXml(config.Input,candidate);}
                            reply.Data=ValidateMirrorPackage(config.Input,candidate);
                        }
                        catch {faulted=true;throw;} // Native calculation/export validation failure requires restart/replay.
                        Check(!request.FailBeforePublish,"Injected failure before publication.");
                        // This is a private trial target, never the user's original. Hold a read
                        // lease excluding writes during hash check + replace. Production also
                        // needs application ownership and recovery around external renames.
                        using(var lease=new FileStream(target,FileMode.Open,FileAccess.Read,FileShare.Read|FileShare.Delete))
                        {
                            using(var sha=SHA256.Create())Check(BitConverter.ToString(sha.ComputeHash(lease)).Replace("-","")==request.ExpectedHash,"Destination changed during save.");
                            File.Replace(candidate,target,target+".before-"+Guid.NewGuid().ToString("N"));
                        }
                        reply.Path=target;reply.Hash=IoTrial.FileHash(target);
                    }
                    else if(request.Action=="read-values")
                    {
                        Check(request.Revision==revision,"Native result request is stale or not yet applied.");
                        ValidateMirrorRead(request);
                        try
                        {
                            engine.Calculate(1);
                            reply.Values=new MirrorValues {Session=config.Session,BaseHash=config.InputHash,Sheet=request.Sheet,
                                Row=request.Row,Column=request.Column,Rows=request.Rows,Columns=request.Columns,
                                Cells=Jagged(engine.Get(request.Sheet,request.Row,request.Column,request.Rows,request.Columns))};
                        }
                        catch {faulted=true;throw;} // Failed native calculation is not a publishable state.
                    }
                    else if(request.Action=="verify")
                    {
                        engine.Calculate(2);reply.Data=config.Model?Probes(engine):new {values=Jagged(engine.Get("ArrayProbe",1,1,5,1))};
                    }
                    else if(request.Action=="diagnostics")reply.Data=new {engine.Diagnostics,workerPeakWorkingSetBytes=Process.GetCurrentProcess().PeakWorkingSet64};
                    else throw new InvalidOperationException("Unknown worker action.");
                    reply.Ok=true;
                }
                catch(Exception e){reply.Error=e.GetBaseException().GetType().Name+": "+e.GetBaseException().Message;}
                reply.Revision=revision;reply.Milliseconds=timer.Elapsed.TotalMilliseconds;Reply(reply);
            }
        }
        return 0;
    }
    static void Reply(MirrorReply reply){Console.WriteLine("MIRROR "+IoTrial.Json.Serialize(reply));Console.Out.Flush();}
    sealed class MirrorClient:IDisposable
    {
        internal readonly Process Process;
        internal MirrorReply Ready;
        internal MirrorClient(string config)
        {
            Process=new Process {StartInfo=new ProcessStartInfo(Assembly.GetExecutingAssembly().Location,"mirror-worker \""+config+"\""){
                UseShellExecute=false,CreateNoWindow=true,WindowStyle=ProcessWindowStyle.Hidden,RedirectStandardInput=true,RedirectStandardOutput=true,RedirectStandardError=true,StandardOutputEncoding=Encoding.UTF8}};
            Process.ErrorDataReceived+=(s,e)=>{if(e.Data!=null)Console.WriteLine("Worker stderr: "+e.Data);};
            Process.Start();Process.BeginErrorReadLine();Ready=Read();Check(Ready.Ok,"Worker failed to initialize.");
        }
        internal void Send(MirrorRequest request){Process.StandardInput.WriteLine(IoTrial.Json.Serialize(request));Process.StandardInput.Flush();}
        internal MirrorReply Read()
        {
            while(true)
            {
                var lineTask=Process.StandardOutput.ReadLineAsync();
                if(!lineTask.Wait(TimeSpan.FromMinutes(5)))throw new TimeoutException("Worker has not replied; do not publish or mark saved.");
                string line=lineTask.Result;if(line==null)throw new EndOfStreamException("Worker ended before acknowledgement.");
                if(line.StartsWith("MIRROR ",StringComparison.Ordinal))return IoTrial.Json.Deserialize<MirrorReply>(line.Substring(7));
                Console.WriteLine("Worker: "+line);
            }
        }
        internal MirrorReply Call(MirrorRequest request){Send(request);return Read();}
        public void Dispose()
        {
            if(!Process.HasExited)
            {
                try {Call(new MirrorRequest {Action="exit"});}catch(IOException) { }
                if(!Process.WaitForExit(30000))throw new TimeoutException("Owned worker did not close; check before another run.");
            }
            Process.Dispose();
        }
    }
    sealed class MirrorJournal
    {
        readonly MirrorConfig config;string chain;
        internal int Revision;
        internal List<string> Files=new List<string>();
        internal MirrorJournal(MirrorConfig c){config=c;chain=c.InputHash;}
        internal string Append(MirrorEdit e)
        {
            e.Revision=Revision+1;e.Session=config.Session;e.BaseHash=config.InputHash;e.PreviousHash=chain;
            string path=Path.Combine(config.Directory,"command-"+e.Revision.ToString("D6")+".json");
            DurableJson(path,e);chain=IoTrial.FileHash(path);Revision++;Files.Add(path);return path;
        }
    }
    static GearEngine NewMirrorProjection(bool model)
    {
        var gear=new GearEngine(model);
        gear.ReadCalculationProjection();
        return gear;
    }
    internal static int RunMirror(string[] args)
    {
        if(args.Length==4&&args[1]=="fixture")
        {
            using(var excel=new ExcelEngine(false)) {excel.Open(IoTrial.InTrial(args[2]));excel.PrepareMirrorSynthetic(IoTrial.InTrial(args[3]));}
            return 0;
        }
        Check(args.Length==5,"mirror excel|devexpress model|synthetic PRIVATE_XLSM NEW_DIRECTORY");
        string backend=args[1],input=IoTrial.InTrial(args[3]);bool model=args[2]=="model";
        Check(new[]{"excel","devexpress"}.Contains(backend)&&new[]{"model","synthetic"}.Contains(args[2]),"Unsupported mirror mode.");
        string dir=Path.GetFullPath(args[4]);Check(!Directory.Exists(dir),"New trial directory required.");
        // Check the new directory's existing parents before creation, then its children.
        IoTrial.InTrial(dir);Directory.CreateDirectory(dir);IoTrial.InTrial(Path.Combine(dir,"config.json"));
        string baseFile=Path.Combine(dir,"baseline.xlsm");File.Copy(input,baseFile);
        var config=new MirrorConfig {Backend=backend,Model=model,Input=baseFile,InputHash=IoTrial.FileHash(baseFile),Directory=dir,Session=Guid.NewGuid().ToString("N")};
        string configFile=Path.Combine(dir,"config.json");DurableJson(configFile,config);
        string target=Path.Combine(dir,"model.xlsm");File.Copy(baseFile,target);
        var journal=new MirrorJournal(config);var times=new Dictionary<string,object>();var results=new Dictionary<string,object>();
        var report=new Dictionary<string,object>{{"backend",backend},{"model",model},{"inputHash",config.InputHash},{"timingsMs",times},{"tests",results},{"scope","Isolated persistent-worker research; no production UI, original-file save, user validation or financial certification"}};
        GearEngine gear=null;MirrorClient worker=null;
        try
        {
            Console.WriteLine("START persistent "+backend+" worker");worker=new MirrorClient(configFile);times["workerColdLoadAndRebuild"]=worker.Ready.Milliseconds;
            gear=NewMirrorProjection(model);times["gearLoadAndRebuild"]=Time(()=>{gear.Open(baseFile);gear.Calculate(2);});
            results["gearImportPolicy"]="Calculation only: ReadObjects=false, ReadVBA=true; authoritative serialization prohibited.";
            string forbiddenProjection=Path.Combine(dir,"forbidden-projection-save.xlsm");
            bool projectionSaveRejected=false;
            try {gear.SaveFunding(forbiddenProjection);}
            catch(InvalidOperationException e){projectionSaveRejected=e.Message.Contains("calculation-only Gear projection");}
            Check(projectionSaveRejected&&!File.Exists(forbiddenProjection),"Reduced Gear copy was not prevented from saving before file creation.");
            results["projectionSaveRejectedWithoutFileCreation"]=true;
            var location=model?gear.MirrorInput():Tuple.Create("ArrayProbe",0,3);
            string sheet=location.Item1;int row=location.Item2,col=location.Item3;object initial=gear.Get(sheet,row,col,1,1)[0,0];
            results["editCell"]=sheet+"!"+Address(row,col);
            var editTimes=new List<double>();var applyTimes=new List<double>();
            // No worker data read-back for ordinary edits. Final assertions below are
            // deliberately separate verification traffic, not part of the proposed UI path.
            foreach(double value in new[]{Convert.ToDouble(initial)+1,Convert.ToDouble(initial)+2,Convert.ToDouble(initial)})
            {
                var edit=new MirrorEdit {Kind="set",Sheet=sheet,Row=row,Column=col,ValueKind="number",Before=gear.Get(sheet,row,col,1,1)[0,0],Value=value};
                editTimes.Add(Time(()=>{string path=journal.Append(edit);worker.Send(new MirrorRequest {Action="apply",Journal=path});ApplyMirrorEdit(gear,edit);gear.Calculate(0);gear.Get(sheet,row,col,1,1);}));
                var ack=worker.Read();Check(ack.Ok,ack.Error);applyTimes.Add(ack.Milliseconds);
            }
            times["durableEditSendGearCalculateAndRead"]=editTimes;times["workerApplyOnly"]=applyTimes;
            Check(worker.Call(new MirrorRequest {Action="apply",Journal=journal.Files.Last()}).Duplicate,"Duplicate replay was not recognized.");results["duplicateCommandAppliedOnce"]=true;
            var gap=new MirrorEdit {Kind="set",Revision=journal.Revision+2,Session=config.Session,BaseHash=config.InputHash,PreviousHash=IoTrial.FileHash(journal.Files.Last())};
            string gapFile=Path.Combine(dir,"gap-probe.json");DurableJson(gapFile,gap);
            Check(!worker.Call(new MirrorRequest {Action="apply",Journal=gapFile}).Ok,"Out-of-order journal accepted.");results["outOfOrderRejected"]=true;
            string prior=IoTrial.FileHash(target);
            var failed=worker.Call(new MirrorRequest {Action="save",Target=target,ExpectedHash=prior,Revision=journal.Revision,FailBeforePublish=true});
            results["injectedFailure"]=failed.Error;Check(!failed.Ok&&failed.Error.Contains("Injected failure before publication")&&IoTrial.FileHash(target)==prior,"Failure injection was not reached or changed target.");results["failedSaveLeavesTargetUnchanged"]=true;
            var saveRequest=new MirrorRequest {Action="save",Target=target,ExpectedHash=prior,Revision=journal.Revision};
            worker.Send(saveRequest);
            var concurrentEdit=new MirrorEdit {Kind="set",Sheet=sheet,Row=row,Column=col,ValueKind="number",Before=initial,Value=Convert.ToDouble(initial)+3};
            times["gearEditWhileWorkerSaving"]=Time(()=>{var path=journal.Append(concurrentEdit);ApplyMirrorEdit(gear,concurrentEdit);gear.Calculate(0);worker.Send(new MirrorRequest {Action="apply",Journal=path});});
            var saved=worker.Read();Check(saved.Ok,saved.Error);times["workerSaveWithCalculationAndPackageCheck"]=saved.Milliseconds;
            Check(saved.Revision==journal.Revision-1,"Save barrier included a later command.");results["laterEditRemainsDirtyAfterEarlierSave"]=true;
            var later=worker.Read();Check(later.Ok,later.Error);
            using(var reopened=NewMirrorProjection(model)) {reopened.Open(target);Check(SameValue(reopened.Get(sheet,row,col,1,1)[0,0],initial),"Saved revision contains the later edit.");}
            // Restore input before reference comparison; this is itself journalled.
            var reset=new MirrorEdit {Kind="set",Sheet=sheet,Row=row,Column=col,ValueKind="number",Before=concurrentEdit.Value,Value=initial};
            var resetFile=journal.Append(reset);ApplyMirrorEdit(gear,reset);gear.Calculate(0);Check(worker.Call(new MirrorRequest {Action="apply",Journal=resetFile}).Ok,"Reset failed.");
            Check(!worker.Call(new MirrorRequest {Action="save",Target=target,ExpectedHash="stale-writer",Revision=journal.Revision}).Ok,"Stale destination token accepted.");results["staleDestinationRejected"]=true;
            using(var locked=new FileStream(target,FileMode.Open,FileAccess.Read,FileShare.None))
                Check(!worker.Call(new MirrorRequest {Action="save",Target=target,ExpectedHash=saved.Hash,Revision=journal.Revision}).Ok,"Locked target overwritten.");
            results["lockedTargetRejected"]=true;
            results["initialWorkerDiagnostics"]=worker.Call(new MirrorRequest {Action="diagnostics"}).Data;
            if(backend=="devexpress")
            {
                worker.Process.Kill();Check(worker.Process.WaitForExit(10000),"Trial worker did not terminate.");
                results["controlledWorkerCrashInjected"]=true;
            }
            worker.Dispose();worker=null;
            Console.WriteLine("START immutable-baseline restart and journal replay");
            times["restartAndReplay"]=Time(()=>{worker=new MirrorClient(configFile);foreach(var path in Directory.GetFiles(dir,"command-*.json").OrderBy(p=>p,StringComparer.Ordinal)){var reply=worker.Call(new MirrorRequest {Action="apply",Journal=path});Check(reply.Ok,reply.Error);}});
            results[backend=="devexpress"?"crashRestartReplayedAllAcknowledgedEdits":"gracefulRestartReplayedAllAcknowledgedEdits"]=true;
            if(model)
            {
                Console.WriteLine("START complete Funding structural barrier");
                string command=journal.Append(new MirrorEdit {Kind="funding10"});var barrier=Stopwatch.StartNew();worker.Send(new MirrorRequest {Action="apply",Journal=command});
                // Gear cannot perform the master's grouped-sheet 3-D fix-up. Do not
                // mutate or publish a speculative projection during this barrier.
                results["gearStructuralMutationSkipped"]=true;
                var structure=worker.Read();results["workerFunding"]=structure;Check(structure.Ok,structure.Error);
                var snapshot=worker.Call(new MirrorRequest {Action="save",Target=target,ExpectedHash=saved.Hash,Revision=journal.Revision});Check(snapshot.Ok,snapshot.Error);
                times["structuralWorkerSave"]=snapshot.Milliseconds;
                // Safe coarse first prototype: read back only at structural boundaries.
                // Rebase from authoritative snapshot instead of trying to patch a subset
                // of formulas/names while missing dependent ranges or new array anchors.
                gear.Dispose();gear=NewMirrorProjection(true);
                times["structuralGearRebaseAndRebuild"]=Time(()=>{gear.Open(target);gear.Calculate(2);});
                times["fullStructuralBarrierIncludingSaveAndRebase"]=barrier.Elapsed.TotalMilliseconds;
                results["correctedGearProbes"]=Probes(gear);results["structuralReadbackFile"]=target;
            }
            else
            {
                // A value-only edit can expand a dynamic array: structural-only
                // reconciliation must not silently assume those displays remain valid.
                var spill=new MirrorEdit {Kind="set",Sheet="ArrayProbe",Row=0,Column=3,ValueKind="number",Before=initial,Value=5.0};
                var path=journal.Append(spill);ApplyMirrorEdit(gear,spill);gear.Calculate(2);Check(worker.Call(new MirrorRequest {Action="apply",Journal=path}).Ok,"Spill edit failed.");
                results["gearSpillValues"]=Jagged(gear.Get("ArrayProbe",1,1,5,1));
                results["workerSpillVerification"]=worker.Call(new MirrorRequest {Action="verify"});
                results["structuralOnlyReadbackNotApprovedForDynamicSpills"]=true;
                TestNativeMirrorReads(config,journal,worker,gear,results,times);
            }
            results["finalWorkerDiagnostics"]=worker.Call(new MirrorRequest {Action="diagnostics"}).Data;
            report["harnessCompleted"]=true;
        }
        catch(Exception e){report["harnessCompleted"]=false;report["error"]=e.GetBaseException().GetType().Name+": "+e.GetBaseException().Message;Console.WriteLine("TRIAL STOPPED: "+report["error"]);}
        finally
        {
            if(gear!=null)gear.Dispose();if(worker!=null)worker.Dispose();
            report["immutableInputUnchanged"]=IoTrial.FileHash(input)==config.InputHash&&IoTrial.FileHash(baseFile)==config.InputHash;
            report["gearProcessPeakWorkingSetBytes"]=Process.GetCurrentProcess().PeakWorkingSet64;
            IoTrial.WriteJson(Path.Combine(dir,"report.json"),report);Console.WriteLine("RESULT "+Path.Combine(dir,"report.json"));
        }
        return (bool)report["harnessCompleted"]&&(bool)report["immutableInputUnchanged"]?0:1;
    }
    sealed partial class GearEngine
    {
        internal bool MirrorEditable(MirrorEdit e){var cell=book.Worksheets[e.Sheet].Cells[e.Row,e.Column];return !cell.HasFormula&&!(bool)cell.Locked;}
        internal Tuple<string,int,int> MirrorInput()
        {
            var cells=book.Worksheets["Stock Assumptions"].Cells;
            for(int r=8;r<200;r++)for(int c=5;c<30;c++){var cell=cells[r,c];if(!cell.HasFormula&&!(bool)cell.Locked&&cell.Value is double)return Tuple.Create("Stock Assumptions",r,c);}
            throw new InvalidOperationException("No reviewed numeric input candidate.");
        }
    }
    sealed partial class ExcelEngine
    {
        internal void PrepareMirrorSynthetic(string output)
        {
            dynamic sheet=Sheet("ArrayProbe"),cell=null;
            try{cell=sheet.Range["D1"];cell.Locked=false;SaveFunding(output);}
            finally{Release((object)cell);Release((object)sheet);}
        }
        internal bool MirrorEditable(MirrorEdit e){dynamic s=Sheet(e.Sheet),r=null;try{r=s.Range[Address(e.Row,e.Column)];return !(bool)r.HasFormula&&!(bool)r.Locked;}finally{Release((object)r);Release((object)s);}}
    }
    sealed partial class DxEngine
    {
        internal bool MirrorEditable(MirrorEdit e){var cell=book.Worksheets[e.Sheet].Cells[e.Row,e.Column];return !cell.HasFormula&&!cell.Protection.Locked;}
        internal void MirrorSave(string path,bool validateArrays=true)
        {
            var sheet=book.Worksheets.FirstOrDefault(s=>s.Name=="Transactional DB");
            var expected=sheet==null?null:MirrorArrayDeclarations(sheet);
            using(var f=new FileStream(path,FileMode.CreateNew,FileAccess.Write))book.SaveDocument(f,DX.DocumentFormat.Xlsm);
            if(validateArrays&&expected!=null)AssertTrialArrayDeclarations(path,sheet.Name,expected);
        }
        internal void MirrorFunding(Dictionary<string,object> times,Dictionary<string,object> results,Action<string> trace=null,bool preserveArrays=true)
        {
            var sheets=FundingSheets();int at=book.DefinedNames.GetDefinedName("LoanDescRev1").Range.LeftColumnIndex-1;const int count=10;
            var app=Assembly.LoadFrom(Path.Combine(Repo,"bin/Release/Abovo-summit.exe"));var type=app.GetType("Abovo.WorkbookStructural3DReferences",true);
            object guard=null;
            times["captureProduction3DGuard"]=Time(()=>guard=type.GetMethod("Capture").Invoke(null,new object[]{book,sheets,new[]{new KeyValuePair<int,int>(at,count)}}));
            var restores=new List<IDisposable>();
            try
            {
                // Isolated native port. Unprotect all protected sheets temporarily so the
                // production AST helper never accesses the live application ModelID table.
                string password=TrialProtection.ReadExistingCredential();foreach(DX.Worksheet ws in book.Worksheets)if(ws.IsProtected)restores.Add(TemporarilyUnprotect(ws.Name,password));
                times["insert32"]=Time(()=>{foreach(var sheet in sheets)Columns(sheet,at,count,true);});
                trace?.Invoke("after-insert32");
                times["applyProduction3DGuard"]=Time(()=>type.GetMethod("Apply").Invoke(guard,new object[]{-1}));
                trace?.Invoke("after-3d-guard");
                times["copy32"]=Time(()=>{foreach(var name in sheets){var ws=book.Worksheets[name];int bottom=ws.GetUsedRange().BottomRowIndex;ws.Range[Column(at)+":"+Column(at+count-1)].CopyFrom(ws.Range[Column(at+count)+":"+Column(at+count)],DX.PasteSpecial.All);type.GetMethod("CopyColumn").Invoke(null,new object[]{ws,at+count,at,count,0,bottom});}});
                trace?.Invoke("after-copy32");
                MirrorDynamicArrayBatch arrayBatch=null;
                if(preserveArrays)times["suspendTdbArrays"]=Time(()=>arrayBatch=new MirrorDynamicArrayBatch(book.Worksheets["Transactional DB"]));
                times["mirrors11"]=Time(()=>{foreach(string source in new[]{"FacilityNames","LoanDescsOrd"})foreach(string name in Mirrors(source))
                {
                    var range=book.DefinedNames.GetDefinedName(name).Range;var ws=range.Worksheet;int bottom=range.BottomRowIndex,left=range.LeftColumnIndex,width=range.ColumnCount;
                    int add=book.DefinedNames.GetDefinedName(source).Range.ColumnCount+1-range.RowCount;Check(add==10,"Mirror precondition differs: "+name);
                    if(arrayBatch!=null){Check(ws.Name=="Transactional DB","Unexpected mirror worksheet.");arrayBatch.InsertRows(bottom,add);}else Rows(ws.Name,bottom,add,true);
                    trace?.Invoke("after-insert-"+name);
                    int restored=0;
                    if(arrayBatch!=null)arrayBatch.CopyRow(ws.Range[Address(bottom-1,left,1,width)],ws.Range[Address(bottom,left,add,width)]);
                    else restored=CopyMirrorRowPreservingDynamicArrays(ws,ws.Range[Address(bottom-1,left,1,width)],ws.Range[Address(bottom,left,add,width)]);
                    trace?.Invoke("after-copy-"+name);
                    results["restoredDynamicArrays:"+name]=restored;
                    Check(book.DefinedNames.GetDefinedName(name).Range.RowCount==book.DefinedNames.GetDefinedName(source).Range.ColumnCount+1,"Mirror name not extended.");
                }});
                if(arrayBatch!=null){times["restoreTdbArrays"]=Time(()=>results["restoredTdbArrays"]=arrayBatch.Complete());trace?.Invoke("after-batch-array-restore");}
                type.GetMethod("Verify").Invoke(guard,null);results["geometryVerified"]=true;results["productionAstGuardReused"]=true;
                trace?.Invoke("after-verify");
            }
            finally{for(int i=restores.Count-1;i>=0;i--)restores[i].Dispose();}
            trace?.Invoke("after-restore-protection");
        }
    }
}
