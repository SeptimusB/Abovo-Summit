using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;

static class SaveCandidateTests
{
    static int assertions;
    static WorkbookReadArea Cell(string sheet="Data",int row=0,int col=0)=>new WorkbookReadArea(sheet,row,col,1,1);
    static WorkbookEngineOptions Options(WorkbookEnginePreference engine=WorkbookEnginePreference.DevExpressOnly,bool udf=false,int timeout=120000)=>new WorkbookEngineOptions(engine,udf,udf,timeout,true,true);
    static void Check(bool value,string text){if(!value)throw new Exception(text);Console.WriteLine("SAVE PASS "+(++assertions)+" "+text);}
    static async Task Reject(Func<Task> action,string text)
    {try{await action();}catch(Exception e)when(e is InvalidOperationException||e is ArgumentException||e is IOException||e is InvalidDataException||e is System.Xml.XmlException||e is OperationCanceledException||e is TimeoutException){Check(true,text+" ["+e.GetType().Name+"]");return;}throw new Exception("Expected rejection: "+text);}
    static string Hash(string path){using(var s=File.OpenRead(path))using(var h=System.Security.Cryptography.SHA256.Create())return BitConverter.ToString(h.ComputeHash(s)).Replace("-","");}
    internal static void Compare(WorkbookCalculationResult expected,WorkbookCalculationResult actual,string label)
    {
        int count=0;
        for(int n=0;n<expected.Blocks.Count;n++)for(int r=0;r<expected.Blocks[n].Area.Rows;r++)for(int c=0;c<expected.Blocks[n].Area.Columns;c++)
        {
            object a=expected.Blocks[n].ValueAt(r,c),b=actual.Blocks[n].ValueAt(r,c);bool same=Equals(a,b);
            if(a is WorkbookCellError&&b is WorkbookCellError)same=((WorkbookCellError)a).Text==((WorkbookCellError)b).Text;
            if(a is double&&b is double)same=Math.Abs((double)a-(double)b)<=Math.Max(1e-9,Math.Abs((double)a)*1e-12);
            if(!same)throw new Exception(label+" mismatch "+expected.Blocks[n].Area.Worksheet+" r="+r+" c="+c+" expected="+a+" actual="+b);
            count++;
        }
        Check(true,label+" across "+count+" output positions");
    }
    sealed class Backend:IWorkbookCandidateBackend
    {
        public string Name=>"Controlled save owner";public string Version=>"1";
        public string Source,LastPath;public int Owner,Exports;public double Value=10d;
        public Action ExportAction;public bool BadRead,FailRead,WrongArea;
        public void OpenReadOnly(string path,WorkbookEngineOptions o){Owner=Thread.CurrentThread.ManagedThreadId;Source=path;}
        void Own(){if(Owner!=Thread.CurrentThread.ManagedThreadId)throw new Exception("Wrong native owner");}
        public void Calculate(WorkbookCalculationKind k){Own();}
        public WorkbookValueBlock Read(WorkbookReadArea a){Own();var data=new object[a.Rows,a.Columns];data[0,0]=Value;return new WorkbookValueBlock(a,data);}
        public WorkbookCellState ReadCell(WorkbookReadArea a){Own();return new WorkbookCellState(Value,"","0.00",false,true,false,false,false);}
        public void WriteValue(WorkbookReadArea a,object v){Own();Value=(double)v;}
        public void ExportCopy(string path){Own();Exports++;LastPath=path;File.Copy(Source,path);ExportAction?.Invoke();}
        public WorkbookCandidateReadback ReadCopy(string path,IList<WorkbookReadArea> areas,IList<WorkbookReadArea> cells)
        {Own();if(FailRead)throw new InvalidDataException("Controlled candidate open failure");var blocks=areas.Select(a=>Read(WrongArea?Cell("wrong"):a)).ToList();if(BadRead)blocks[0]=new WorkbookValueBlock(areas[0],new object[,]{{99d}});return new WorkbookCandidateReadback(blocks,cells.Select(ReadCell));}
        public void Dispose(){Own();}
    }
    static string Folder(string path){var root=Path.Combine(Path.GetDirectoryName(path),"save-trial-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);return root;}
    internal static async Task Run(string path)
    {
        string root=Folder(path),original=Hash(path);var b=new Backend();
        var s=await WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(),default(CancellationToken),_=>b);
        try{var result=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{Cell()});var cell=await s.CaptureCellAsync(0,Cell());await Reject(()=>s.CreateSaveCandidateAsync(result,root,new[]{cell}),"default session cannot export");Check(b.Exports==0,"default export opt-out never touches disk");}finally{await s.CloseAsync();}
        b=new Backend();s=await WorkbookCalculationSession.OpenAsync(path,Options(),default(CancellationToken),_=>b);
        try
        {
            var result=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{Cell()});var cell=await s.CaptureCellAsync(0,Cell());
            await Reject(()=>s.CreateSaveCandidateAsync(result,"relative",new[]{cell}),"relative candidate directory rejected");
            await Reject(()=>s.CreateSaveCandidateAsync(result,root,new WorkbookCellSnapshot[0]),"missing definition checkpoints rejected");
            await Reject(()=>s.CreateSaveCandidateAsync(result,root,Enumerable.Repeat(cell,65)),"unbounded definition checkpoints rejected");
            using(var cancel=new CancellationTokenSource()){cancel.Cancel();await Reject(()=>s.CreateSaveCandidateAsync(result,root,new[]{cell},cancel.Token),"pre-cancelled candidate does not export");}
            var candidate=await s.CreateSaveCandidateAsync(result,root,new[]{cell});
            Check(candidate.SessionId==s.SessionId&&candidate.Revision==s.Revision&&candidate.SourceHash==original&&candidate.Hash==Hash(candidate.Path),"candidate receipt binds exact bytes, source and revision");
            await s.ValidateSaveCandidateAsync(candidate);Check(s.IsCurrentCandidate(candidate)&&s.IsCurrent(result),"candidate verification leaves current results and revision unchanged");
            using(var stream=new FileStream(candidate.Path,FileMode.Append,FileAccess.Write))stream.WriteByte(0);
            await Reject(()=>s.ValidateSaveCandidateAsync(candidate),"changed candidate bytes cannot pass publication preflight");
            var changed=await s.ApplyValueAsync(cell,20d,WorkbookValuePermission.UnlockedCell,new[]{Cell()});
            Check(!s.IsCurrentCandidate(candidate),"newer edit invalidates old candidate");
            await Reject(()=>s.ValidateSaveCandidateAsync(candidate),"old candidate cannot acknowledge newer edit");
            await Reject(()=>s.CreateSaveCandidateAsync(result,root,new[]{cell}),"old calculated result cannot export");
            result=changed.Results;cell=changed.After;
            foreach(var failure in new[]{"read-value","read-area","cancel","metadata"})
            {
                using(var cancel=new CancellationTokenSource())
                {
                    b.FailRead=failure=="read-open";b.BadRead=failure=="read-value";b.WrongArea=failure=="read-area";
                    b.ExportAction=()=>{if(failure=="partial-export")throw new IOException("Controlled interrupted export");if(failure=="cancel")cancel.Cancel();if(failure=="metadata"){using(var zip=ZipFile.Open(b.LastPath,ZipArchiveMode.Update))using(var writer=new StreamWriter(zip.CreateEntry("customXml/unexpected.xml").Open()))writer.Write("<unexpected/>");}};
                    await Reject(()=>s.CreateSaveCandidateAsync(result,root,new[]{cell},cancel.Token),"reject "+failure);
                    Check(!File.Exists(b.LastPath)&&s.IsCurrent(result)&&s.Revision==1,"remove only owned failed candidate; native state reusable: "+failure);
                }
            }
            b.ExportAction=null;b.FailRead=b.BadRead=b.WrongArea=false;
            candidate=await s.CreateSaveCandidateAsync(result,root,new[]{cell});await s.ValidateSaveCandidateAsync(candidate);
            Check(s.IsCurrentCandidate(candidate),"fresh candidate succeeds after recoverable export failures");
        }finally{await s.CloseAsync();}
        foreach(var failure in new[]{"partial-export","read-open","changed-owner"})
        {
            b=new Backend();s=await WorkbookCalculationSession.OpenAsync(path,Options(),default(CancellationToken),_=>b);
            try
            {
                var result=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{Cell()});var cell=await s.CaptureCellAsync(0,Cell());
                b.FailRead=failure=="read-open";
                b.ExportAction=()=>{if(failure=="partial-export")throw new IOException("Controlled interrupted native export");if(failure=="changed-owner")b.Value=99d;};
                await Reject(()=>s.CreateSaveCandidateAsync(result,root,new[]{cell}),"quarantine uncertain native state: "+failure);
                Check(!s.IsCurrent(result)&&!File.Exists(b.LastPath),"faulted export cannot keep serving old results or leave an accepted candidate");
                await Reject(()=>s.CaptureCellAsync(s.Revision,Cell()),"faulted native export owner cannot be reused");
            }finally{await s.CloseAsync();}
        }
        foreach(bool timeout in new[]{false,true})
        {
            b=new Backend();s=await WorkbookCalculationSession.OpenAsync(path,Options(timeout:timeout?250:5000),default(CancellationToken),_=>b);
            using(var release=new ManualResetEventSlim())
            {
                var entered=new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                b.ExportAction=()=>{entered.TrySetResult(true);release.Wait();};
                try
                {
                    var result=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{Cell()});var cell=await s.CaptureCellAsync(0,Cell());
                    var task=s.CreateSaveCandidateAsync(result,root,new[]{cell});await entered.Task;
                    await Reject(()=>s.ApplyValueAsync(cell,22d,WorkbookValuePermission.UnlockedCell),"edit cannot overlap candidate export");
                    await Reject(()=>Task.Run(()=>s.InvalidateResults()),"external revision cannot overlap candidate export");
                    await Reject(()=>s.CreateSaveCandidateAsync(result,root,new[]{cell}),"second export cannot overlap first");
                    if(timeout){await Reject(()=>task,"hung export wait is bounded");Check(!s.IsCurrent(result),"timed-out owner is quarantined");release.Set();await s.NativeCleanupCompletion;Check(!File.Exists(b.LastPath),"late export is removed, never accepted");}
                    else{release.Set();var candidate=await task;await s.ValidateSaveCandidateAsync(candidate);Check(s.IsCurrentCandidate(candidate),"serialized export accepted after native return");}
                }finally{release.Set();await s.CloseAsync();}
            }
        }
        Check(Hash(path)==original,"all controlled save trials leave original bytes unchanged");
        await PackageSafety(path);
        Console.WriteLine("SAVE ASSERTIONS="+assertions);
    }
    static object Invoke(Type t,object target,string method,params object[] args)
    {
        try{return t.GetMethod(method,BindingFlags.NonPublic|BindingFlags.Static|BindingFlags.Instance).Invoke(target,args);}
        catch(TargetInvocationException e){System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(e.InnerException).Throw();throw;}
    }
    static void ChangeXml(string file,string name,Action<XDocument> change)
    {
        using(var zip=ZipFile.Open(file,ZipArchiveMode.Update))
        {var entry=zip.GetEntry(name);XDocument xml;using(var s=entry.Open())xml=XDocument.Load(s);change(xml);entry.Delete();using(var s=zip.CreateEntry(name).Open())xml.Save(s);}
    }
    static async Task PackageSafety(string original)
    {
        string root=Folder(original),source=Fixture(root,".xlsx"),sourceHash=Hash(source);
        var type=typeof(WorkbookEngineOptions).Assembly.GetType("Abovo.WorkbookEngines.WorkbookCandidatePackage");
        object baseline=Invoke(type,null,"Capture",source);
        foreach(string kind in new[]{"payload","missing-link","extra-part","metadata","schema-properties","preserved-space","dtd","format","signature"})
        {
            string file=Path.Combine(root,kind+".xlsx");File.Copy(source,file);
            if(kind=="payload")ChangeXml(file,"customXml/item1.xml",d=>d.Root.Descendants().First(e=>e.Name.LocalName=="Edit").SetAttributeValue("value","after"));
            if(kind=="preserved-space")ChangeXml(file,"customXml/item1.xml",d=>d.Root.Descendants().First(e=>e.Name.LocalName=="Space").Value="a");
            if(kind=="schema-properties")ChangeXml(file,"customXml/itemProps1.xml",d=>{d.Root.SetAttributeValue(d.Root.Name.Namespace+"itemID",Guid.NewGuid().ToString());d.Root.Elements().Remove();});
            if(kind=="metadata")ChangeXml(file,"xl/metadata.xml",d=>d.Descendants().First(e=>e.Name.LocalName=="dynamicArrayProperties").SetAttributeValue("fDynamic","0"));
            if(kind=="missing-link")ChangeXml(file,"xl/_rels/workbook.xml.rels",d=>d.Root.Elements().Where(e=>((string)e.Attribute("Type")).EndsWith("/customXml")).Remove());
            if(kind=="format")ChangeXml(file,"[Content_Types].xml",d=>d.Root.Elements().First(e=>(string)e.Attribute("PartName")=="/xl/workbook.xml").SetAttributeValue("ContentType","application/vnd.ms-excel.sheet.macroEnabled.main+xml"));
            if(kind=="signature")using(var zip=ZipFile.Open(file,ZipArchiveMode.Update))using(var w=new StreamWriter(zip.CreateEntry("xl/vbaProjectSignature.bin").Open()))w.Write("synthetic signature marker");
            if(kind=="extra-part"||kind=="dtd")using(var zip=ZipFile.Open(file,ZipArchiveMode.Update))
            {string name=kind=="extra-part"?"customXml/unexpected.xml":"customXml/item1.xml";zip.GetEntry(name)?.Delete();using(var w=new StreamWriter(zip.CreateEntry(name).Open()))w.Write(kind=="dtd"?"<!DOCTYPE t [<!ENTITY x 'external'>]><t>&x;</t>":"<unexpected/>");}
            Func<Task> check=()=>Task.Run(()=>{Invoke(type,baseline,"PrepareValueOnlyCandidate",file);Invoke(type,baseline,"Verify",file);});
            if(kind=="schema-properties"){await check();Check(true,"restore original XML companion properties only after matching unchanged payload");}
            else await Reject(check,"reject changed/missing custom XML or array semantics: "+kind);
        }
        Check(Hash(source)==sourceHash,"package verification and preservation never change source bytes");
        var excel=typeof(WorkbookEngineOptions).Assembly.GetType("Abovo.WorkbookEngines.ExcelCalculationBackend");
        string zone="[ZoneTransfer]\r\nZoneId=3\r\n";
        Invoke(excel,null,"CopyInternetZone",source,zone);
        var b=new Backend();var s=await WorkbookCalculationSession.OpenAsync(source,Options(),default(CancellationToken),_=>b);
        try
        {
            var result=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{Cell()});var cell=await s.CaptureCellAsync(0,Cell());
            var candidate=await s.CreateSaveCandidateAsync(result,root,new[]{cell});
            Check((string)Invoke(excel,null,"ReadInternetZone",candidate.Path)==zone,"candidate retains downloaded-file security marker");
            Invoke(excel,null,"CopyInternetZone",candidate.Path,"[ZoneTransfer]\r\nZoneId=1\r\n");
            await Reject(()=>s.ValidateSaveCandidateAsync(candidate),"security-marker change invalidates candidate even when ZIP hash is unchanged");
            Check((string)Invoke(excel,null,"ReadInternetZone",source)==zone,"source marker is never removed or downgraded");
        }finally{await s.CloseAsync();}
    }
    internal static string Fixture(string root,string extension)
    {
        string path=Path.Combine(root,"input"+extension);
        using(var wb=new Workbook())
        {
            var ws=wb.Worksheets[0];ws.Name="Data";ws.Cells["A1"].Value=10d;ws.Cells["A1"].NumberFormat="0.00";ws.Cells["A1"].Protection.Locked=false;ws.Cells["A1"].FillColor=System.Drawing.Color.LightBlue;
            ws.Cells["B1"].Formula="=SaveTrialInput*2";ws.Cells["C1"].DynamicArrayFormula="={1;2;3}";
            ws.Cells["D1"].Formula="=CONCATENATE("+string.Join(",",Enumerable.Repeat("\"x\"",37))+")";
            wb.DefinedNames.Add("SaveTrialInput","Data!$A$1");
            var conditional=ws.ConditionalFormattings.AddFormulaExpressionConditionalFormatting(ws.Range["A1"],"=$A$1>=100");
            conditional.Formatting.Fill.PatternType=PatternType.Gray125;
            conditional.Formatting.Font.Color=System.Drawing.Color.DarkRed;
            wb.CustomXmlParts.Add("<Trial xmlns='urn:abovo:save-trial'><History><Edit value='before'/></History><Space xml:space='preserve'> a </Space></Trial>");
            wb.CalculateFullRebuild();
            if(extension==".xlsb")
            {
                // Native XLSB fixture already uses the production compatibility
                // service; long formula behavior is also exercised on AGL.
                ws.Cells["D1"].Formula="=\""+new string('x',37)+"\"";
                wb.CalculateFullRebuild();
            }
            wb.SaveDocument(path,extension==".xlsb"?DocumentFormat.Xlsb:extension==".xlsm"?DocumentFormat.Xlsm:DocumentFormat.Xlsx);
        }
        return path;
    }
    internal static async Task Native(string original,bool agl)
    {
        var before=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).ToArray();
        string root=Folder(original);var paths=agl?new[]{original}:new[]{".xlsx",".xlsm",".xlsb"}.Select(ext=>Fixture(root,ext)).ToArray();
        foreach(string path in paths)
        foreach(var engine in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired})
        {
            string baseline=Hash(path);var s=await WorkbookCalculationSession.OpenAsync(path,Options(engine,agl));
            try
            {
                var area=agl?Cell("Funding Assumptions",81,6):Cell();
                var reads=agl?new[]{new WorkbookReadArea("Detailed Comp Inc - Trad View",0,0,98,42),new WorkbookReadArea("Financial Position - Trad View",0,0,67,43),new WorkbookReadArea("Cashflow detailed",0,0,51,94),new WorkbookReadArea("Check Sheet",0,0,63,11),new WorkbookReadArea("Development Expenditure",0,0,613,54)}:new[]{new WorkbookReadArea("Data",0,0,3,4)};
                var input=await s.CaptureCellAsync(0,area);
                var edit=await s.ApplyValueAsync(input,Convert.ToDouble(input.State.Value)+300d,WorkbookValuePermission.UnlockedCell,reads);
                if(!agl)Console.WriteLine("SAVE_FIXTURE before export B1="+edit.Results.Blocks[0].ValueAt(0,1)+" C1="+edit.Results.Blocks[0].ValueAt(0,2)+" C3="+edit.Results.Blocks[0].ValueAt(2,2));
                var probes=new List<WorkbookCellSnapshot>{edit.After};
                if(!agl){probes.Add(await s.CaptureCellAsync(s.Revision,Cell("Data",0,1)));probes.Add(await s.CaptureCellAsync(s.Revision,Cell("Data",0,2)));}
                var candidate=await s.CreateSaveCandidateAsync(edit.Results,root,probes);
                await s.ValidateSaveCandidateAsync(candidate);
                Check(s.IsCurrentCandidate(candidate)&&s.IsCurrent(edit.Results),s.EngineName+" same-format candidate accepted "+Path.GetExtension(path));
                Console.WriteLine("SAVE_NATIVE engine="+s.EngineName+" format="+Path.GetExtension(path)+" exportMs="+candidate.ExportMilliseconds+" verificationMs="+candidate.VerificationMilliseconds+" path="+candidate.Path);
                // Independent native reopen and full calculation. The candidate
                // now has the original security marker; VBA obeys existing policy.
                var other=engine==WorkbookEnginePreference.ExcelRequired?WorkbookEnginePreference.DevExpressOnly:WorkbookEnginePreference.ExcelRequired;
                var check=await WorkbookCalculationSession.OpenAsync(candidate.Path,Options(other,agl));
                try{var snap=await check.CaptureCellAsync(0,area);Check(Equals(snap.State.Value,edit.After.State.Value),s.EngineName+" saved input reopens in "+check.EngineName);var calc=await check.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,reads);Compare(edit.Results,calc,s.EngineName+" save -> "+check.EngineName+" full calculation");if(!agl)Check(Equals(calc.Blocks[0].ValueAt(2,2),3d)&&Equals(calc.Blocks[0].ValueAt(0,1),620d),s.EngineName+" formula and dynamic spill survive cross-engine save/reopen");}finally{await check.CloseAsync();}
                await s.ApplyValueAsync(edit.After,input.State.Value,WorkbookValuePermission.UnlockedCell,reads);
                Check(!s.IsCurrentCandidate(candidate),s.EngineName+" undo makes saved candidate stale");
            }finally{await s.CloseAsync();}
            Check(Hash(path)==baseline,"native candidate never overwrote source "+engine+" "+Path.GetExtension(path));
        }
        var clock=Stopwatch.StartNew();int[] after;
        do{after=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).ToArray();if(!after.Except(before).Any())break;await Task.Delay(100);}while(clock.ElapsedMilliseconds<10000);
        Check(!after.Except(before).Any()&&!before.Except(after).Any(),"candidate tests close only owned Excel processes");
        Console.WriteLine("SAVE_NATIVE ASSERTIONS="+assertions);
    }
}
