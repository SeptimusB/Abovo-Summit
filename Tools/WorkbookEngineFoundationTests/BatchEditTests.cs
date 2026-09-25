using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Abovo.WorkbookEngines;

static class BatchEditTests
{
    static int assertions;
    static WorkbookReadArea Cell(int column) => new WorkbookReadArea("Data",0,column,1,1);
    static void Check(bool ok,string text) { if(!ok)throw new Exception(text);Console.WriteLine("BATCH PASS "+(++assertions)+" "+text); }
    static async Task Reject(Func<Task> action,string text)
    { try{await action();}catch(Exception e)when(e is InvalidOperationException||e is ArgumentException||e is AggregateException||e is OperationCanceledException||e is InvalidDataException){Check(true,text);return;}throw new Exception("Expected rejection: "+text); }
    sealed class Backend:IWorkbookValueEditBackend
    {
        public string Name=>"Batch test"; public string Version=>"1";
        public readonly object[] Values={10d,20d,30d};
        public readonly bool[] Locked=new bool[3];
        public int Writes,Calculates,FailWrite=-1,FailCalc,Owner;
        public bool FailRestore,WrongRead;
        public Action OnWrite,OnCalc;
        void Own(){if(Thread.CurrentThread.ManagedThreadId!=Owner)throw new Exception("Wrong native thread");}
        public void OpenReadOnly(string path,WorkbookEngineOptions options){Owner=Thread.CurrentThread.ManagedThreadId;}
        public WorkbookCellState ReadCell(WorkbookReadArea a){Own();return new WorkbookCellState(Values[a.Column],"","0.00",Locked[a.Column],true,false,false,false);}
        public void WriteValue(WorkbookReadArea a,object v){Own();Writes++;if(FailRestore&&Equals(v,10d))throw new InvalidOperationException("Restore failed");Values[a.Column]=v;OnWrite?.Invoke();if(a.Column==FailWrite){FailWrite=-1;throw new InvalidOperationException("Partial write");}}
        public void Calculate(WorkbookCalculationKind kind){Own();Calculates++;OnCalc?.Invoke();if(FailCalc-->0)throw new InvalidOperationException("Calculation failure");}
        public WorkbookValueBlock Read(WorkbookReadArea a){Own();var values=new object[a.Rows,a.Columns];for(int i=0;i<a.Columns;i++)values[0,i]=Values[a.Column+i];return new WorkbookValueBlock(WrongRead?new WorkbookReadArea("Wrong",a.Row,a.Column,a.Rows,a.Columns):a,values);}
        public void Dispose(){Own();}
    }
    static WorkbookValueChange Change(WorkbookCellSnapshot s,object value,bool calc=false,bool skip=false) => new WorkbookValueChange(s,value,WorkbookValuePermission.UnlockedCell,calc,skip);
    internal static async Task Run(string path)
    {
        using(var book=new DevExpress.Spreadsheet.Workbook())
        {
            var sheet=book.Worksheets[0];sheet.Cells["A1"].Value=0d;
            sheet.Cells["B1"].Fill.PatternType=DevExpress.Spreadsheet.PatternType.Solid;
            var rule=sheet.ConditionalFormattings.AddFormulaExpressionConditionalFormatting(sheet.Range["B1"],"=$A$1=0");
            rule.Formatting.Fill.PatternType=DevExpress.Spreadsheet.PatternType.Gray125;
            book.CalculateFull();Check(sheet.Cells["B1"].Fill.PatternType==DevExpress.Spreadsheet.PatternType.Gray125,"DevExpress Cell.Fill reflects active conditional pattern");
            sheet.Cells["A1"].Value=1d;book.CalculateFull();Check(sheet.Cells["B1"].Fill.PatternType==DevExpress.Spreadsheet.PatternType.Solid,"DevExpress Cell.Fill returns base pattern when condition clears");
        }
        var b=new Backend();var options=new WorkbookEngineOptions(enableValueEditTrial:true);
        var s=await WorkbookCalculationSession.OpenAsync(path,options,default(CancellationToken),_=>b);
        try
        {
            var areas=new[]{Cell(0),Cell(1),Cell(2)};var read=new WorkbookReadArea("Data",0,0,1,3);
            await Reject(()=>s.CaptureCellsAsync(0,new WorkbookReadArea[0]),"empty capture rejected");
            await Reject(()=>s.CaptureCellsAsync(0,new[]{Cell(0),Cell(0)}),"duplicate capture rejected");
            await Reject(()=>s.CaptureCellsAsync(0,new[]{read}),"non-cell capture rejected");
            await Reject(()=>s.CaptureCellsAsync(0,Enumerable.Range(0,10001).Select(n=>new WorkbookReadArea("Data",n,0,1,1))),"capture bounded");
            var cells=await s.CaptureCellsAsync(0,areas);
            Check(cells.Count==3&&cells.All(x=>x.SessionId==s.SessionId&&x.Revision==0),"batch capture at one owner/revision");
            await Reject(()=>s.ApplyValuesAsync(new WorkbookValueChange[0]),"empty changes rejected");
            await Reject(()=>s.ApplyValuesAsync(new[]{Change(cells[0],11d),Change(cells[0],12d)}),"duplicate target rejected before write");
            var no=await s.ApplyValuesAsync(cells.Select((x,i)=>Change(x,b.Values[i])));
            Check(!no.Changed&&b.Writes==0&&b.Calculates==0&&s.Revision==0,"unchanged batch preserves revision without calculation");
            b.Values[2]=31d;
            await Reject(()=>s.ApplyValuesAsync(cells.Select(x=>Change(x,99d))),"changed final target prevents entire batch");
            Check(b.Writes==0&&s.Revision==1,"all originals checked before first write and stale outputs invalidated");b.Values[2]=30d;
            cells=await s.CaptureCellsAsync(s.Revision,areas);
            var done=await s.ApplyValuesAsync(cells.Select((x,i)=>Change(x,100d+i)),new[]{read});
            Check(done.Changed&&done.Changes.Count==3&&s.Revision==2&&b.Calculates==1,"three writes use one calculation and revision");
            Check(s.IsCurrent(done.Results)&&Equals(done.Results.Blocks[0].ValueAt(0,2),102d),"outputs share final command revision");
            var undo=await s.ApplyValuesAsync(done.Changes.Select(x=>Change(x.After,x.Before.State.Value)),new[]{read});
            Check(b.Values.SequenceEqual(new object[]{10d,20d,30d})&&undo.Changes.Count==3,"whole receipt restores original values");
            await Reject(()=>s.ApplyValuesAsync(cells.Select(x=>Change(x,99d))),"obsolete whole command rejected");
            foreach(var fault in new[]{"write","calc","read","cancel-write","cancel-calc","late-lock"})
            {
                cells=await s.CaptureCellsAsync(s.Revision,areas);var revision=s.Revision;
                using(var cancel=new CancellationTokenSource())
                {
                    b.FailWrite=fault=="write"?1:-1;b.FailCalc=fault=="calc"?1:0;b.WrongRead=fault=="read";
                    b.OnWrite=fault=="cancel-write"?(Action)(()=>cancel.Cancel()):null;b.OnCalc=fault=="cancel-calc"?(Action)(()=>cancel.Cancel()):null;
                    if(fault=="late-lock")b.Locked[2]=true;
                    if(fault=="late-lock")cells=await s.CaptureCellsAsync(s.Revision,areas);
                    await Reject(()=>s.ApplyValuesAsync(cells.Select(x=>Change(x,99d)),new[]{read},cancel.Token),"compensate entire batch: "+fault);
                    Check(b.Values.SequenceEqual(new object[]{10d,20d,30d})&&s.Revision==revision+1,"all values restored; revision invalidated: "+fault);
                    b.WrongRead=false;b.OnWrite=null;b.OnCalc=null;b.Locked[2]=false;
                    Check((await s.CaptureCellsAsync(s.Revision,areas)).Count==3,"restored session usable: "+fault);
                }
            }
            b.Locked[1]=true;cells=await s.CaptureCellsAsync(s.Revision,areas);
            done=await s.ApplyValuesAsync(cells.Select((x,i)=>Change(x,200d+i,i==2,true)),new[]{read});
            Check(done.Changes.Count==2&&done.Skipped.Count==1&&Equals(b.Values[1],20d),"explicit soft skip leaves unavailable input untouched");
            Check(done.Skipped[0].Column==1,"skipped receipt identifies exact input");b.Locked[1]=false;
            await s.ApplyValuesAsync(done.Changes.Select(x=>Change(x.After,x.Before.State.Value)));
            cells=await s.CaptureCellsAsync(s.Revision,areas);int calculations=b.Calculates;
            done=await s.ApplyValuesAsync(new[]{Change(cells[0],11d),Change(cells[1],21d,true),Change(cells[2],31d)},new[]{read});
            Check(b.Calculates==calculations+2,"prerequisite calculation then final calculation");
            await s.ApplyValuesAsync(done.Changes.Select(x=>Change(x.After,x.Before.State.Value)));
            cells=await s.CaptureCellsAsync(s.Revision,areas);
            using(var cancel=new CancellationTokenSource()){cancel.Cancel();int writes=b.Writes;await Reject(()=>s.ApplyValuesAsync(cells.Select(x=>Change(x,1d)),null,cancel.Token),"pre-cancelled command rejected");Check(b.Writes==writes,"pre-cancellation performs no writes");}
            b.FailWrite=1;b.FailRestore=true;
            await Reject(()=>s.ApplyValuesAsync(cells.Select(x=>Change(x,99d))),"failed compensation quarantines session");
            Check(Equals(b.Values[1],20d),"compensation continues past another restoration failure");
            await Reject(()=>s.CaptureCellsAsync(s.Revision,areas),"quarantined owner cannot continue");
        }
        finally{await s.CloseAsync();}
        foreach(string failure in new[]{"calculate","read"})
        {
            b=new Backend();s=await WorkbookCalculationSession.OpenAsync(path,options,default(CancellationToken),_=>b);
            try
            {
                var cell=await s.CaptureCellAsync(0,Cell(0));b.FailCalc=failure=="calculate"?1:0;b.WrongRead=failure=="read";
                await Reject(()=>s.ApplyValuesAsync(new[]{Change(cell,10d)},new[]{Cell(0)}),"unchanged batch native "+failure+" failure surfaced");
                Check(b.Writes==0,"unchanged failed native read performs no write");
                await Reject(()=>s.CaptureCellAsync(s.Revision,Cell(0)),"unverified native failure without compensation quarantines session");
            }
            finally{await s.CloseAsync();}
        }
        Console.WriteLine("BATCH ASSERTIONS="+assertions);
    }
}
