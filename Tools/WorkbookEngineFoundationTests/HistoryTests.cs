using System;
using System.Collections.Concurrent;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Abovo;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;

static class HistoryTests
{
    static int assertions;
    static readonly XNamespace Ns="urn:abovo:summit:recovery-history:1";
    static readonly Type Store=typeof(ModelHistorySnapshot).Assembly.GetType("Abovo.RecoveryHistoryStore");
    static readonly WorkbookReadArea Cell=new WorkbookReadArea("Data",0,0,1,1);
    static void Check(bool yes,string label){if(!yes)throw new Exception(label);Console.WriteLine("HISTORY PASS "+(++assertions)+" "+label);}
    static string Hash(string path){using(var s=File.OpenRead(path))using(var h=System.Security.Cryptography.SHA256.Create())return BitConverter.ToString(h.ComputeHash(s)).Replace("-","");}
    static object Call(Type type,object instance,string method,params object[] args)
    {try{return type.GetMethod(method,BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static).Invoke(instance,args);}catch(TargetInvocationException e){System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(e.InnerException).Throw();throw;}}
    static async Task Reject(Func<Task> action,string label)
    {try{await action();}catch(Exception e)when(e is IOException||e is InvalidDataException||e is InvalidOperationException||e is ArgumentException||e is OperationCanceledException||e is System.Xml.XmlException){Check(true,label);return;}throw new Exception("Expected rejection: "+label);}
    static string Folder(string file){var p=Path.Combine(Path.GetDirectoryName(file),"history-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(p);return p;}
    static WorkbookEngineOptions Options(WorkbookEnginePreference engine=WorkbookEnginePreference.DevExpressOnly,bool udf=false)=>new WorkbookEngineOptions(engine,udf,udf,120000,true,true,true);
    static string HistoryPart(string path)
    {using(var z=ZipFile.OpenRead(path))return z.Entries.Where(e=>e.FullName.StartsWith("customXml/")&&e.FullName.EndsWith(".xml")).Single(e=>{using(var s=e.Open())return XDocument.Load(s).Root.Name==Ns+"RecoveryHistory";}).FullName;}
    static XDocument HistoryXml(string path){using(var z=ZipFile.OpenRead(path))using(var s=z.GetEntry(HistoryPart(path)).Open())return XDocument.Load(s);}
    static void MutateHistory(string path,Action<XDocument> change)
    {string part=HistoryPart(path);var doc=HistoryXml(path);change(doc);using(var z=ZipFile.Open(path,ZipArchiveMode.Update)){z.GetEntry(part).Delete();using(var s=z.CreateEntry(part).Open())doc.Save(s);}}

    // Real production change manager, owned by one STA. Only its absent UI host
    // is stubbed; typed edits, dirty revisions and history serialization are real.
    sealed class Model:IDisposable
    {
        readonly BlockingCollection<Action> queue=new BlockingCollection<Action>();readonly Thread owner;
        Workbook wb;FileManager.ExcelModel[] previous;readonly WorkbookReadArea area;
        internal FileManager.ExcelModel Value;internal ModelChangeManagerV2 Manager;
        public Model(string source,WorkbookReadArea location=null,double initial=10d)
        {
            area=location??Cell;
            owner=new Thread(()=>{foreach(var a in queue.GetConsumingEnumerable())a();}){IsBackground=true};owner.SetApartmentState(ApartmentState.STA);owner.Start();
            Do(()=>{previous=FileManager.ExcelModels;wb=new Workbook();wb.Worksheets[0].Name=area.Worksheet;wb.Worksheets[0].Cells[area.Row,area.Column].Value=initial;
                Value=(FileManager.ExcelModel)FormatterServices.GetUninitializedObject(typeof(FileManager.ExcelModel));Value.WB=wb;Value.FileName=source;Value.WBCalcEngine=new CalcEngine(0);
                FileManager.ExcelModels=new[]{Value};int id=0;Manager=new ModelChangeManagerV2(ref id);Value.ChangeManager=Manager;return true;}).GetAwaiter().GetResult();
        }
        public Task<T> Do<T>(Func<T> action)
        {var t=new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);queue.Add(()=>{try{t.SetResult(action());}catch(Exception e){t.SetException(e);}});return t.Task;}
        public Task<ModelHistorySnapshot> Capture(WorkbookCalculationResult result)=>Do(()=>Manager.CaptureSaveHistory(result));
        public Task<bool> Current(ModelHistorySnapshot s,WorkbookPublicationReceipt r)=>Do(()=>Manager.IsSavedHistoryCurrent(s,r));
        public Task<bool> Edit(double number)=>Do(()=>{var result=Manager.ProcessChange(new DataChangeEvent{ModelID=0,WSName=area.Worksheet,CellAddress=area.Address.Split(':')[0],Description="History trial edit",DataFormat="N",ChangedValue=number,TimeStamp=DateTime.UtcNow,UserName="Synthetic test"});if(!result.BSuccess)throw new Exception(result.StrResponseMessage);return true;});
        public void Dispose(){Do(()=>{FileManager.ExcelModels=previous;wb.Dispose();return true;}).GetAwaiter().GetResult();queue.CompleteAdding();owner.Join();queue.Dispose();}
    }

    internal static async Task Run(string input)
    {
        string root=Folder(input),source=Path.Combine(root,"source"+Path.GetExtension(input));File.Copy(input,source);string original=Hash(source);
        var backend=new PublicationTests.Backend();var session=await WorkbookCalculationSession.OpenAsync(source,Options(),default(CancellationToken),_=>backend);
        using(var model=new Model(source))
        try
        {
            var result=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{Cell});var cell=await session.CaptureCellAsync(0,Cell);
            await model.Edit(20d);var snapshot=await model.Capture(result);
            Check(snapshot.SessionId==session.SessionId&&snapshot.EngineRevision==0&&snapshot.SourceHash==original&&snapshot.UserRevision>0,"history binds real user revision and engine opening baseline");
            await Reject(()=>Task.Run(()=>model.Manager.CaptureSaveHistory(result)),"cross-thread history capture rejected");
            await Reject(()=>model.Do(()=>{using(model.Manager.BeginChangeGroup("pending"))return model.Manager.CaptureSaveHistory(result);}),"in-progress grouped edit cannot be snapshotted");
            await model.Do(()=>{model.Value.IsClosing=true;return true;});await Reject(()=>model.Capture(result),"closing model rejected");await model.Do(()=>{model.Value.IsClosing=false;return true;});
            string other=Path.Combine(root,"same-bytes"+Path.GetExtension(source));File.Copy(source,other);
            await model.Do(()=>{model.Value.FileName=other;return true;});var differentPath=await model.Capture(result);
            await Reject(()=>session.CreateSaveCandidateWithHistoryAsync(result,root,new[]{cell},differentPath),"identical bytes at another opening path cannot substitute history");
            await model.Do(()=>{model.Value.FileName=source;return true;});
            File.AppendAllText(other,"different");await model.Do(()=>{model.Value.FileName=other;return true;});await Reject(()=>model.Capture(result),"different source bytes rejected before capture");await model.Do(()=>{model.Value.FileName=source;return true;});
            using(var cancel=new CancellationTokenSource()){cancel.Cancel();await Reject(()=>session.CreateSaveCandidateWithHistoryAsync(result,root,new[]{cell},snapshot,cancel.Token),"cancelled history candidate not exported");}
            var second=await WorkbookCalculationSession.OpenAsync(source,Options(),default(CancellationToken),_=>new PublicationTests.Backend());
            try{var r=await second.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{Cell});var c=await second.CaptureCellAsync(0,Cell);await Reject(()=>second.CreateSaveCandidateWithHistoryAsync(r,root,new[]{c},snapshot),"another engine session cannot consume snapshot");}finally{await second.CloseAsync();}
            await model.Edit(25d); // A later live edit must not mutate captured XML.
            var candidate=await session.CreateSaveCandidateWithHistoryAsync(result,root,new[]{cell},snapshot);
            Check(candidate.HistorySnapshotId==snapshot.Id&&candidate.HistoryHash==snapshot.Hash,"candidate receipt binds exact history identity and hash");
            var xml=HistoryXml(candidate.Path);Check(xml.Root.Elements(Ns+"Edit").Count()==1&&!xml.Descendants(Ns+"Action").Any(),"actual typed edit persisted without executable Undo command");
            Check(await model.Do(()=>model.Manager.GetHistoryTable().Rows.Count)==2&&xml.Root.Elements().Single().Element(Ns+"NewValue").Value=="20","snapshot stays immutable when live history advances before export");
            Check(Hash(source)==original&&await model.Do(()=>model.Value.IsDirty&&model.Manager.CanUndo),"candidate leaves source, live dirty state and existing Undo intact");
            var noHistory=await session.CreateSaveCandidateAsync(result,root,new[]{cell});
            var wrong=await session.PublishAndCloseAsync(noHistory,Path.Combine(root,"without-history"+Path.GetExtension(source)),WorkbookPublicationMode.CreateNew);
            Check(!await model.Current(snapshot,wrong),"receipt without captured history cannot acknowledge it");
        }
        finally{await session.CloseAsync();}

        // Publication and same-manager evidence, without claiming model clean.
        session=await WorkbookCalculationSession.OpenAsync(source,Options(),default(CancellationToken),_=>new PublicationTests.Backend());
        using(var model=new Model(source))
        try
        {
            var result=await session.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{Cell});var cell=await session.CaptureCellAsync(0,Cell);await model.Edit(30d);var snapshot=await model.Capture(result);
            session.InvalidateResults();var newer=await session.CalculateAndReadAsync(1,WorkbookCalculationKind.Full,new[]{Cell});var newCell=await session.CaptureCellAsync(1,Cell);
            await Reject(()=>session.CreateSaveCandidateWithHistoryAsync(newer,root,new[]{newCell},snapshot),"older engine revision cannot attach history to a newer save");
            snapshot=await model.Capture(newer);var candidate=await session.CreateSaveCandidateWithHistoryAsync(newer,root,new[]{newCell},snapshot);
            string target=Path.Combine(root,"with-history"+Path.GetExtension(source));var receipt=await session.PublishAndCloseAsync(candidate,target,WorkbookPublicationMode.CreateNew);
            Check(await model.Current(snapshot,receipt)&&await model.Do(()=>model.Value.IsDirty),"matching history receipt is evidence only; dirty remains set");
            var history=(DataTable)await model.Do(()=>Call(Store,null,"Read",target,model.Manager.GetHistoryTable()));
            Check(history.Rows.Count==1&&history.Rows[0]["Action"].Equals("")&&(int)history.Rows[0]["GroupID"]<0,"recovered rows remain display-only with negative IDs");
            await model.Do(()=>{Call(typeof(ModelChangeManagerV2),model.Manager,"RestoreRecoveryHistory",history);return true;});
            Check(!await model.Current(snapshot,receipt),"history-only changes detected even without a newer user/calculation revision");
            await model.Do(()=>{Call(typeof(ModelChangeManagerV2),model.Manager,"RestoreRecoveryHistory",history.Clone());return true;});
            Check(await model.Current(snapshot,receipt),"identical history evidence becomes current again without changing dirty state");
            await model.Edit(40d);Check(!await model.Current(snapshot,receipt),"later typed input cannot be cleared by old successful save");
            await model.Do(()=>{int id=0;model.Value.ChangeManager=new ModelChangeManagerV2(ref id);return true;});
            await Reject(()=>model.Current(snapshot,receipt),"replaced live change manager cannot accept old receipt");
        }
        finally{await session.CloseAsync();}

        using(var model=new Model(source))
        {
            var schema=await model.Do(()=>model.Manager.GetHistoryTable());
            for(int i=0;i<1005;i++)schema.Rows.Add(i,new DateTime(2026,1,1).AddSeconds(i),i==1004?new string('x',3000):"Edit "+i,"Data","A1","10","20","Synthetic","Applied","N",1,"Undo");
            var doc=(XDocument)Call(Store,null,"ToDocument",schema);
            Check(doc.Root.Elements().Count()==1000&&doc.Root.Elements().First().Element(Ns+"GroupID").Value=="1004","history bounded to newest 1000 entries");
            Check(doc.Root.Elements().First().Element(Ns+"Description").Value.Length<2100&&!doc.Descendants(Ns+"Action").Any(),"large fields truncated and executable actions omitted");
            foreach(string ext in new[]{".xlsx",".xlsm",".xlsb"})
            {
                string dir=Path.Combine(root,ext.Substring(1));Directory.CreateDirectory(dir);string path=SaveCandidateTests.Fixture(dir,ext);
                Call(Store,null,"WriteDocument",path,doc);string part=HistoryPart(path);var read=(DataTable)Call(Store,null,"Read",path,schema);
                Check(read.Rows.Count==1000&&read.Rows.Cast<DataRow>().All(r=>(int)r["GroupID"]<0&&r["Action"].Equals("")),"read-only bounded import "+ext);
                var small=(XDocument)Call(Store,null,"ToDocument",schema.Clone());Call(Store,null,"WriteDocument",path,small);
                Check(HistoryPart(path)==part&&!HistoryXml(path).Root.Elements().Any(),"replace owned history without duplication "+ext);
                MutateHistory(path,d=>d.Root.SetAttributeValue("version","99"));string hash=Hash(path);
                await Reject(()=>Task.Run(()=>Call(Store,null,"WriteDocument",path,small)),"unknown history version not overwritten "+ext);Check(Hash(path)==hash,"rejected version leaves workbook bytes intact "+ext);
            }
            // A future/ambiguous existing history part rejects the whole new
            // candidate, not just its history, without damaging the live owner.
            foreach(string fault in new[]{"future-version","duplicate"})
            {
                string dir=Path.Combine(root,fault);Directory.CreateDirectory(dir);string path=SaveCandidateTests.Fixture(dir,".xlsm");
                Call(Store,null,"WriteDocument",path,doc);
                if(fault=="future-version")MutateHistory(path,d=>d.Root.SetAttributeValue("version","99"));
                else{var duplicate=HistoryXml(path);using(var zip=ZipFile.Open(path,ZipArchiveMode.Update))using(var stream=zip.CreateEntry("customXml/duplicate-history.xml").Open())duplicate.Save(stream);}
                string hash=Hash(path);var owner=await WorkbookCalculationSession.OpenAsync(path,Options(),default(CancellationToken),_=>new PublicationTests.Backend());
                await model.Do(()=>{model.Value.FileName=path;return true;});
                try
                {
                    var calc=await owner.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{Cell});var probe=await owner.CaptureCellAsync(0,Cell);var snap=await model.Capture(calc);
                    await Reject(()=>owner.CreateSaveCandidateWithHistoryAsync(calc,dir,new[]{probe},snap),"malformed history rejects candidate: "+fault);
                    Check(owner.IsCurrent(calc)&&Hash(path)==hash&&!Directory.GetDirectories(dir,"candidate-*").Any(),"failed history candidate removed; source and owner intact: "+fault);
                }finally{await owner.CloseAsync();}
            }
        }
        Check(Hash(source)==original,"all history tests leave opening file unchanged");Console.WriteLine("HISTORY ASSERTIONS="+assertions);
    }

    internal static async Task Native(string input,bool agl)
    {
        int[] processes=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).ToArray();string root=Folder(input);
        string[] paths=agl?new[]{input}:new[]{".xlsm",".xlsb"}.Select(ext=>SaveCandidateTests.Fixture(root,ext)).ToArray();
        foreach(string path in paths)
        foreach(var engine in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired})
        {
            string hash=Hash(path);var area=agl?new WorkbookReadArea("Funding Assumptions",81,6,1,1):Cell;
            var reads=agl?new[]{new WorkbookReadArea("Detailed Comp Inc - Trad View",0,0,98,42),new WorkbookReadArea("Financial Position - Trad View",0,0,67,43),new WorkbookReadArea("Cashflow detailed",0,0,51,94),new WorkbookReadArea("Check Sheet",0,0,63,11),new WorkbookReadArea("Development Expenditure",0,0,613,54)}:new[]{new WorkbookReadArea("Data",0,0,3,4)};
            var session=await WorkbookCalculationSession.OpenAsync(path,Options(engine,agl));
            var initial=await session.CaptureCellAsync(0,area);
            using(var model=new Model(path,area,Convert.ToDouble(initial.State.Value)))
            try
            {
                for(int iteration=0;iteration<2;iteration++)
                {
                    var old=await session.CaptureCellAsync(session.Revision,area);double value=Convert.ToDouble(old.State.Value)+100d;
                    var edit=await session.ApplyValueAsync(old,value,WorkbookValuePermission.UnlockedCell,reads);await model.Edit(value);var history=await model.Capture(edit.Results);
                    var candidate=await session.CreateSaveCandidateWithHistoryAsync(edit.Results,root,new[]{edit.After},history);
                    string target=Path.Combine(root,engine+"-"+Guid.NewGuid().ToString("N")+Path.GetExtension(path));var receipt=await session.PublishAndCloseAsync(candidate,target,WorkbookPublicationMode.CreateNew);
                    Check(await model.Current(history,receipt)&&await model.Do(()=>model.Value.IsDirty),engine+" published current history without clearing live dirty "+Path.GetExtension(path));
                    Check(HistoryXml(target).Root.Elements().Count()==iteration+1,engine+" latest full history, one part, cycle "+iteration);
                    var latest=HistoryXml(target).Root.Elements().First();Check(latest.Element(Ns+"Worksheet").Value==area.Worksheet&&latest.Element(Ns+"Cell").Value==area.Address.Split(':')[0],engine+" history points at the actual edited cell");
                    Console.WriteLine("HISTORY_NATIVE engine="+engine+" format="+Path.GetExtension(path)+" cycle="+iteration+" exportMs="+candidate.ExportMilliseconds+" verifyMs="+candidate.VerificationMilliseconds);
                    var other=engine==WorkbookEnginePreference.ExcelRequired?WorkbookEnginePreference.DevExpressOnly:WorkbookEnginePreference.ExcelRequired;
                    var reopened=await WorkbookCalculationSession.OpenAsync(target,Options(other,agl));
                    try{var check=await reopened.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,reads);SaveCandidateTests.Compare(edit.Results,check,"history "+engine+" -> "+other+" cycle="+iteration);}finally{await reopened.CloseAsync();}
                    if(iteration==0){session=await receipt.ReopenAsync();await model.Do(()=>{model.Value.FileName=target;return true;});}
                }
            }
            finally{await session.CloseAsync();}
            Check(Hash(path)==hash,"native history trial never overwrote original "+engine);
        }
        var clock=Stopwatch.StartNew();int[] after;
        do{after=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).ToArray();if(processes.OrderBy(x=>x).SequenceEqual(after.OrderBy(x=>x)))break;await Task.Delay(100);}while(clock.ElapsedMilliseconds<10000);
        Check(processes.OrderBy(x=>x).SequenceEqual(after.OrderBy(x=>x)),"native history tests closed only owned Excel processes");Console.WriteLine("HISTORY_NATIVE ASSERTIONS="+assertions);
    }
}
