using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Abovo.WorkbookEngines;
using DevExpress.Spreadsheet;

static class ValueEditTests
{
    static int assertions;
    static WorkbookReadArea Cell(string sheet="Data", int row=0, int column=0) => new WorkbookReadArea(sheet,row,column,1,1);
    static WorkbookEngineOptions Options(WorkbookEnginePreference preference=WorkbookEnginePreference.Automatic, bool functions=false, int timeout=120000) =>
        new WorkbookEngineOptions(preference,functions,functions,timeout,true);
    static void Check(bool ok,string text)
    { if(!ok)throw new Exception(text); Console.WriteLine("EDIT PASS "+(++assertions)+" "+text); }
    static async Task Reject(Func<Task> action,string text)
    { try{await action();}catch(Exception e)when(e is InvalidOperationException||e is ArgumentException||e is AggregateException||e is OperationCanceledException||e is TimeoutException||e is InvalidDataException){Check(true,text);return;}throw new Exception("Expected rejection: "+text); }
    sealed class Backend:IWorkbookValueEditBackend
    {
        public string Name=>"Controlled edit owner";public string Version=>"1";
        public object Value=10d;
        public bool Locked,Solid=true,Array,Merged,Protected,FailWrite,FailRestore,WrongArea,FailReadCell;
        public int Calculates,FailCalculates,Writes,Owner;
        public string Formula="",Format="0.00";
        public Action OnCalculate,OnRead;
        public readonly TaskCompletionSource<bool> Disposed=new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        public void OpenReadOnly(string path,WorkbookEngineOptions options){Owner=Thread.CurrentThread.ManagedThreadId;}
        void Own(){if(Owner!=Thread.CurrentThread.ManagedThreadId)throw new Exception("Wrong native thread");}
        public void Calculate(WorkbookCalculationKind kind){Own();Calculates++;OnCalculate?.Invoke();if(FailCalculates-->0)throw new InvalidOperationException("Controlled calculation failure");}
        public WorkbookValueBlock Read(WorkbookReadArea area){Own();var values=new object[area.Rows,area.Columns];values[0,0]=Value;OnRead?.Invoke();return new WorkbookValueBlock(WrongArea?Cell("Wrong"):area,values);}
        public WorkbookCellState ReadCell(WorkbookReadArea area){Own();if(FailReadCell)throw new InvalidOperationException("Controlled input-state read failure");return new WorkbookCellState(Value,Formula,Format,Locked,Solid,Array,Merged,Protected);}
        public void WriteValue(WorkbookReadArea area,object value){Own();Writes++;if(FailRestore && Equals(value,10d))throw new InvalidOperationException("Controlled restore failure");Value=value;if(FailWrite){FailWrite=false;throw new InvalidOperationException("Partial native write");}}
        public void Dispose(){Own();Disposed.TrySetResult(true);}
    }
    internal static async Task Run(string path)
    {
        WorkbookCellSnapshot foreign=null;
        var b=new Backend();var s=await WorkbookCalculationSession.OpenAsync(path,new WorkbookEngineOptions(),default(CancellationToken),_=>b);
        try{var cell=await s.CaptureCellAsync(0,Cell());foreign=cell;await Reject(()=>s.ApplyValueAsync(cell,20d,WorkbookValuePermission.UnlockedCell),"default session cannot edit");Check(b.Writes==0,"default opt-out did not touch native data");}finally{await s.CloseAsync();}
        b=new Backend();s=await WorkbookCalculationSession.OpenAsync(path,Options(),default(CancellationToken),_=>b);
        try
        {
            var first=await s.CaptureCellAsync(0,Cell());
            await Reject(()=>s.ApplyValueAsync(foreign,20d,WorkbookValuePermission.UnlockedCell),"snapshot from another session cannot edit this workbook");
            await Reject(()=>s.ApplyValueAsync(first,20d,WorkbookValuePermission.UnlockedCell,Enumerable.Repeat(Cell(),129)),"oversized edit read-back rejected before mutation");
            foreach(var invalid in new object[]{double.NaN,double.PositiveInfinity,1,DateTime.Today,new WorkbookCellError("#N/A"),new string('x',32768)})
                await Reject(()=>s.ApplyValueAsync(first,invalid,WorkbookValuePermission.UnlockedCell),"reject untyped or unsupported input");
            Check(b.Writes==0&&s.Revision==0,"invalid inputs do not write or advance revision");
            var no=await s.ApplyValueAsync(first,10d,WorkbookValuePermission.UnlockedCell);
            Check(!no.Changed&&b.Calculates==0&&s.Revision==0,"unchanged input has no calculation or new revision");
            var old=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{Cell()});
            var edit=await s.ApplyValueAsync(first,12.5d,WorkbookValuePermission.UnlockedCell,new[]{Cell()});
            Check(edit.Changed&&edit.Before==first&&Equals(edit.After.State.Value,12.5d)&&s.Revision==1,"typed edit returns immutable before/after receipt");
            Check(!s.IsCurrent(old)&&s.IsCurrent(edit.Results)&&Equals(edit.Results.Blocks[0].ValueAt(0,0),12.5d),"read-back and input share the same committed revision");
            await Reject(()=>s.ApplyValueAsync(first,99d,WorkbookValuePermission.UnlockedCell),"old edit snapshot rejected");
            var undo=await s.ApplyValueAsync(edit.After,edit.Before.State.Value,WorkbookValuePermission.UnlockedCell);
            Check(Equals(undo.After.State.Value,10d)&&s.Revision==2,"receipt supports expected-before restoration");
            var snap=await s.CaptureCellAsync(s.Revision,Cell());
            var cached=await s.CalculateAndReadAsync(s.Revision,WorkbookCalculationKind.Full,new[]{Cell()});b.Value=11d;
            await Reject(()=>s.ApplyValueAsync(snap,20d,WorkbookValuePermission.UnlockedCell),"native content changed without a revision is detected");
            Check(!s.IsCurrent(cached)&&s.Revision==snap.Revision+1,"untracked native change invalidates earlier cached outputs");b.Value=10d;
            b.Locked=true;snap=await s.CaptureCellAsync(s.Revision,Cell());
            await Reject(()=>s.ApplyValueAsync(snap,20d,WorkbookValuePermission.UnlockedCell),"lock-bit rule rejects locked input");
            edit=await s.ApplyValueAsync(snap,20d,WorkbookValuePermission.SolidFillRule);
            Check(Equals(edit.After.State.Value,20d),"existing fill-rule path remains distinct from lock-bit rule on an unprotected sheet");b.Value=10d;
            b.Solid=false;snap=await s.CaptureCellAsync(s.Revision,Cell());
            await Reject(()=>s.ApplyValueAsync(snap,20d,WorkbookValuePermission.SolidFillRule),"non-solid pattern rejects fill-rule input");
            b.Solid=true;b.Protected=true;snap=await s.CaptureCellAsync(s.Revision,Cell());
            await Reject(()=>s.ApplyValueAsync(snap,20d,WorkbookValuePermission.SolidFillRule),"never bypass actual protected locked cell");b.Protected=false;b.Locked=false;
            foreach(var kind in new[]{"formula","array","merge"})
            {b.Formula=kind=="formula"?"=1":"";b.Array=kind=="array";b.Merged=kind=="merge";snap=await s.CaptureCellAsync(s.Revision,Cell());await Reject(()=>s.ApplyValueAsync(snap,20d,WorkbookValuePermission.UnlockedCell),"trial refuses "+kind);}
            b.Formula="";b.Array=false;b.Merged=false;
            snap=await s.CaptureCellAsync(s.Revision,Cell());b.Format="0.0";
            await Reject(()=>s.ApplyValueAsync(snap,20d,WorkbookValuePermission.UnlockedCell),"changed format snapshot is rejected");b.Format="0.00";
            foreach(var failure in new[]{"partial-write","calculation","read-area","cancel","cancel-read"})
            {
                snap=await s.CaptureCellAsync(s.Revision,Cell());var rev=s.Revision;
                using(var cancel=new CancellationTokenSource())
                {
                    b.FailWrite=failure=="partial-write";b.FailCalculates=failure=="calculation"?1:0;b.WrongArea=failure=="read-area";
                    b.OnCalculate=failure=="cancel"?(Action)(()=>cancel.Cancel()):null;
                    b.OnRead=failure=="cancel-read"?(Action)(()=>cancel.Cancel()):null;
                    await Reject(()=>s.ApplyValueAsync(snap,20d,WorkbookValuePermission.UnlockedCell,new[]{Cell()},cancel.Token),"compensate "+failure);
                    b.OnCalculate=null;b.OnRead=null;b.WrongArea=false;
                    Check(Equals(b.Value,10d)&&s.Revision==rev+1,"restoration succeeds; invalidated revision retained for "+failure);
                    Check((await s.CaptureCellAsync(s.Revision,Cell())).State.Value.Equals(10d),"session reusable after verified "+failure+" restoration");
                }
            }
            using(var cancel=new CancellationTokenSource())
            {cancel.Cancel();snap=await s.CaptureCellAsync(s.Revision,Cell());int writes=b.Writes;await Reject(()=>s.ApplyValueAsync(snap,20d,WorkbookValuePermission.UnlockedCell,null,cancel.Token),"pre-cancelled input rejected");Check(writes==b.Writes,"pre-cancel does not write");}
            b.FailCalculates=1;b.FailRestore=true;snap=await s.CaptureCellAsync(s.Revision,Cell());
            await Reject(()=>s.ApplyValueAsync(snap,20d,WorkbookValuePermission.UnlockedCell),"restoration failure quarantines the session");
            await Reject(()=>s.CaptureCellAsync(s.Revision,Cell()),"quarantined session refuses further reads");
        }finally{await s.CloseAsync();}
        foreach(var timeout in new[]{false,true})
        {
            b=new Backend();s=await WorkbookCalculationSession.OpenAsync(path,Options(timeout:timeout?250:5000),default(CancellationToken),_=>b);
            using(var release=new ManualResetEventSlim())
            using(var cancel=new CancellationTokenSource())
            {
                var entered=new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);int held=0;
                b.OnCalculate=()=>{if(Interlocked.Increment(ref held)==1){entered.TrySetResult(true);release.Wait();}};
                try
                {
                    var snap=await s.CaptureCellAsync(0,Cell());
                    var operation=s.ApplyValueAsync(snap,20d,WorkbookValuePermission.UnlockedCell,new[]{Cell()},cancel.Token);
                    if(await Task.WhenAny(entered.Task,Task.Delay(5000))!=entered.Task)throw new Exception("Held edit never entered calculation");
                    Check(Equals(b.Value,20d)&&s.Revision==1,"held edit invalidates old results before calculation");
                    await Reject(()=>s.ApplyValueAsync(snap,30d,WorkbookValuePermission.UnlockedCell),"concurrent input cannot overwrite pending edit");
                    bool invalidationBlocked=false;try{s.InvalidateResults();}catch(InvalidOperationException){invalidationBlocked=true;}
                    Check(invalidationBlocked,"external revision changes cannot interleave with a pending edit");
                    if(timeout)
                    {
                        if(await Task.WhenAny(operation,Task.Delay(5000))!=operation)throw new Exception("Edit timeout did not return");
                        await Reject(()=>operation,"timed-out edit produces no successful receipt");
                        Check(!b.Disposed.Task.IsCompleted&&s.NativeCleanupCompletion!=null,"timed-out native edit retains owner until it returns");
                        release.Set();await s.NativeCleanupCompletion;
                        Check(Equals(b.Value,10d)&&b.Disposed.Task.IsCompleted,"late timed-out edit restores then closes on owner thread");
                    }
                    else
                    {
                        cancel.Cancel();release.Set();await Reject(()=>operation,"in-flight cancellation restores before returning failure");
                        Check(Equals((await s.CaptureCellAsync(s.Revision,Cell())).State.Value,10d),"cancelled session remains usable only after verified restoration");
                    }
                }
                finally{release.Set();await s.CloseAsync();}
            }
        }
        foreach(var preflight in new[]{false,true})
        {
            b=new Backend();s=await WorkbookCalculationSession.OpenAsync(path,Options(),default(CancellationToken),_=>b);
            try
            {
                var cached=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Full,new[]{Cell()});
                var snap=await s.CaptureCellAsync(0,Cell());b.FailReadCell=true;
                if(preflight)await Reject(()=>s.ApplyValueAsync(snap,20d,WorkbookValuePermission.UnlockedCell),"failed native edit preflight quarantines without writing");
                else await Reject(()=>s.CaptureCellAsync(0,Cell()),"failed native cell capture quarantines");
                Check(!s.IsCurrent(cached)&&b.Writes==0,"native read failure invalidates previous outputs without mutation");
                await Reject(()=>s.CalculateAndReadAsync(s.Revision,WorkbookCalculationKind.Full,new[]{Cell()}),"failed input reader cannot continue serving results");
            }finally{await s.CloseAsync();}
        }
        Console.WriteLine("EDIT ASSERTIONS="+assertions);
    }

    static string CreateNativeFixture(string original)
    {
        var folder=Path.Combine(Path.GetDirectoryName(original),"value-edit-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
        var path=Path.Combine(folder,"input.xlsx");
        using(var wb=new Workbook())
        {
            var ws=wb.Worksheets[0];ws.Name="Data";ws.Cells["A1"].Value=10d;ws.Cells["A1"].NumberFormat="0.00";
            ws.Cells["A1"].Protection.Locked=false;ws.Cells["A1"].FillColor=System.Drawing.Color.LightBlue;
            ws.Cells["A2"].Formula="=A1*2";ws.Cells["A2"].Protection.Locked=false;
            ws.Cells["B1"].Value=5d;ws.Cells["B1"].FillColor=System.Drawing.Color.LightBlue;
            ws.Cells["C1"].Protection.Locked=false;
            ws.MergeCells(ws.Range["E1:F1"]);ws.Cells["E1"].Protection.Locked=false;
            ws.Cells["G1"].DynamicArrayFormula="={1;2}";ws.Range["G1:G2"].Protection.Locked=false;
            var guard=wb.Worksheets.Add("Protected");guard.Cells["A1"].Protection.Locked=false;guard.Protect("",WorksheetProtectionPermissions.Default);
            wb.CalculateFullRebuild();wb.SaveDocument(path,DocumentFormat.Xlsx);
        }
        return path;
    }
    static WorkbookReadArea[] Outputs()=>new[]{new WorkbookReadArea("Detailed Comp Inc - Trad View",0,0,98,42),new WorkbookReadArea("Financial Position - Trad View",0,0,67,43),new WorkbookReadArea("Cashflow detailed",0,0,51,94),new WorkbookReadArea("Check Sheet",0,0,63,11),new WorkbookReadArea("Development Expenditure",0,0,613,54)};
    static void Compare(WorkbookCalculationResult a,WorkbookCalculationResult b,string label)
    {
        int cells=0;
        for(int n=0;n<a.Blocks.Count;n++)for(int r=0;r<a.Blocks[n].Area.Rows;r++)for(int c=0;c<a.Blocks[n].Area.Columns;c++)
        {object x=a.Blocks[n].ValueAt(r,c),y=b.Blocks[n].ValueAt(r,c);bool same=x is double&&y is double?Math.Abs((double)x-(double)y)<=Math.Max(1e-6,Math.Abs((double)x)*1e-10):x is WorkbookCellError&&y is WorkbookCellError?x.ToString()==y.ToString():Equals(x,y);if(!same)throw new Exception(label+" differs at "+a.Blocks[n].Area.Worksheet+" "+r+","+c+": "+x+" != "+y);cells++;}
        Check(true,label+" across "+cells+" positions");
    }
    internal static async Task Native(string source,bool agl)
    {
        string path=agl?source:CreateNativeFixture(source);
        var inputHash=Hash(path);
        var before=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).ToArray();
        WorkbookCalculationResult reference=null;
        foreach(var preference in new[]{WorkbookEnginePreference.DevExpressOnly,WorkbookEnginePreference.ExcelRequired})
        {
            var s=await WorkbookCalculationSession.OpenAsync(path,Options(preference,agl));
            try
            {
                var area=agl?Cell("Funding Assumptions",81,6):Cell();
                var outputs=agl?Outputs():new[]{new WorkbookReadArea("Data",0,0,2,1)};
                var baseline=await s.CalculateAndReadAsync(0,WorkbookCalculationKind.Rebuild,outputs);
                var input=await s.CaptureCellAsync(0,area);Console.WriteLine("EDIT_NATIVE engine="+s.EngineName+" input="+input.State.Value+" locked="+input.State.Locked+" solid="+input.State.SolidFill+" protected="+input.State.WorksheetProtected);
                var editTimer=Stopwatch.StartNew();
                var changed=await s.ApplyValueAsync(input,Convert.ToDouble(input.State.Value)+300d,WorkbookValuePermission.UnlockedCell,outputs);
                Console.WriteLine("EDIT_NATIVE_TIME engine="+s.EngineName+" totalMs="+editTimer.ElapsedMilliseconds+" calculationMs="+changed.Results.CalculationMilliseconds+" readBackMs="+changed.Results.TransferMilliseconds+"; single integration trial, not comparative benchmark");
                Check(changed.Changed&&s.IsCurrent(changed.Results),s.EngineName+" typed input and outputs publish together");
                if(reference==null)reference=changed.Results;else Compare(reference,changed.Results,"native changed-value parity");
                var undo=await s.ApplyValueAsync(changed.After,input.State.Value,WorkbookValuePermission.UnlockedCell,outputs);
                Compare(baseline,undo.Results,s.EngineName+" restored outputs");
                if(!agl)
                {
                    foreach(var value in new object[]{12.5d,-0.0525d,45000d,true,false,"=1+1","+123","0012","01/02/2026","'literal","a\nb",null,""})
                    {
                        input=await s.CaptureCellAsync(s.Revision,area);
                        var receipt=await s.ApplyValueAsync(input,value,WorkbookValuePermission.UnlockedCell,outputs);
                        var expectedValue=value is string && (string)value==""?null:value;
                        Check(Equals(receipt.After.State.Value,expectedValue)&&receipt.After.State.Formula==""&&receipt.After.State.NumberFormat=="0.00",s.EngineName+" literal/typed value preserved: "+(value?.ToString()??"blank"));
                    }
                    foreach(var blocked in new[]{Cell("Data",1,0),Cell("Data",0,1),Cell("Data",1,6),Cell("Data",0,4),Cell("Protected",0,1)})
                    {input=await s.CaptureCellAsync(s.Revision,blocked);await Reject(()=>s.ApplyValueAsync(input,42d,WorkbookValuePermission.UnlockedCell),s.EngineName+" rejects guarded "+blocked.Worksheet+"!"+blocked.Address);}
                    input=await s.CaptureCellAsync(s.Revision,Cell("Data",0,1));
                    var fillEdit=await s.ApplyValueAsync(input,42d,WorkbookValuePermission.SolidFillRule);
                    Check(Equals(fillEdit.After.State.Value,42d)&&fillEdit.After.State.Locked&&fillEdit.After.State.SolidFill,s.EngineName+" raw solid-fill rule permits input without changing lock or fill");
                    input=await s.CaptureCellAsync(s.Revision,Cell("Data",0,2));await Reject(()=>s.ApplyValueAsync(input,42d,WorkbookValuePermission.SolidFillRule),s.EngineName+" preserves fill-rule rejection");
                    input=await s.CaptureCellAsync(s.Revision,Cell("Protected"));var ok=await s.ApplyValueAsync(input,42d,WorkbookValuePermission.UnlockedCell);
                    Check(ok.After.State.WorksheetProtected&&Equals(ok.After.State.Value,42d),s.EngineName+" keeps worksheet protection for unlocked input");
                }
                Console.WriteLine("EDIT_NATIVE_DONE engine="+s.EngineName+" revision="+s.Revision);
            }finally{await s.CloseAsync();}
        }
        var timer=Stopwatch.StartNew();int[] after;
        do{after=Process.GetProcessesByName("EXCEL").Select(p=>{using(p)return p.Id;}).ToArray();if(!after.Except(before).Any())break;await Task.Delay(100);}while(timer.ElapsedMilliseconds<10000);
        Check(!after.Except(before).Any()&&!before.Except(after).Any(),"native edit sessions close only owned Excel");
        Check(Hash(path)==inputHash,"native edited source bytes unchanged");
        Console.WriteLine("EDIT_NATIVE ASSERTIONS="+assertions+" source="+path+"; no native save or source overwrite");
    }
    static string Hash(string path)
    {using(var stream=File.OpenRead(path))using(var hash=System.Security.Cryptography.SHA256.Create())return BitConverter.ToString(hash.ComputeHash(stream));}
}
