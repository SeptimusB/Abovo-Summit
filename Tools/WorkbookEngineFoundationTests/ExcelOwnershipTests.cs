using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Abovo.WorkbookEngines;

// Controlled process-wide boundary tests. No user Excel, windows or files.
static class ExcelOwnershipTests
{
    static int checks;
    static void Check(bool ok,string message){if(!ok)throw new Exception(message);Console.WriteLine("OWNERSHIP PASS "+(++checks)+" "+message);}
    public sealed class Books
    {
        public readonly List<Book> Items=new List<Book>();
        public bool FailCount;
        public int Count{get{if(FailCount)throw new IOException("Synthetic inaccessible collection");return Items.Count;}}
        public Book this[int index]=>Items[index-1];
    }
    public sealed class Book
    {
        readonly Books owner;
        public bool Closed,Saved;
        public Book(Books owner){this.owner=owner;owner.Items.Add(this);}
        public void Close(bool save){Closed=true;Saved=save;owner.Items.Remove(this);}
    }
    public sealed class Application
    {
        public Books Workbooks{get;}=new Books();
        public bool QuitCalled{get;private set;}
        public bool UserControl{get;set;}
        public bool ScreenUpdating{get;set;}
        public bool DisplayAlerts{get;set;}
        public bool Visible{get;set;}
        public int Calculations;
        public int CalculationState=>0;
        public void CalculateFull(){Calculations++;}
        public void Quit(){QuitCalled=true;}
    }
    internal static Task Run()
    {
        var done=new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread=new Thread(()=>{try{RunOnOwner();done.SetResult(true);}catch(Exception e){done.SetException(e);}}){IsBackground=true};
        thread.SetApartmentState(ApartmentState.STA);thread.Start();return done.Task;
    }
    static void RunOnOwner()
    {
        var type=typeof(WorkbookCalculationSession).Assembly.GetType("Abovo.WorkbookEngines.ExcelCalculationBackend");
        foreach(string scenario in new[]{"exclusive","foreign","uninspectable"}){
            var backend=(IWorkbookCalculationBackend)Activator.CreateInstance(type,true);
            var app=new Application();var book=new Book(app.Workbooks);var seed=new Book(app.Workbooks);
            foreach(var pair in new Dictionary<string,object>{{"app",app},{"books",app.Workbooks},{"book",book},{"seed",seed},{"owned",true}})
                type.GetField(pair.Key,BindingFlags.Instance|BindingFlags.NonPublic).SetValue(backend,pair.Value);
            backend.Calculate(WorkbookCalculationKind.Full);Check(app.Calculations==1,"exclusive owner can calculate");
            Book foreign=null;
            if(scenario=="foreign"){
                foreign=new Book(app.Workbooks);
                try{backend.Calculate(WorkbookCalculationKind.Full);throw new Exception("Expected calculation refusal");}catch(InvalidOperationException){Check(app.Calculations==1,"foreign workbook prevents process-wide calculation");}
                try{((IWorkbookValueEditBackend)backend).WriteValue(new WorkbookReadArea("Data",0,0,1,1),5d);throw new Exception("Expected edit refusal");}catch(InvalidOperationException){Check(true,"foreign owner refuses further writes");}
                try{((IWorkbookCandidateBackend)backend).ExportCopy("unused.xlsm");throw new Exception("Expected save refusal");}catch(InvalidOperationException){Check(!File.Exists("unused.xlsm"),"foreign owner refuses export before creating a file");}
            }
            if(scenario=="uninspectable")app.Workbooks.FailCount=true;
            bool failed=false;try{backend.Dispose();}catch(AggregateException){failed=true;}
            Check(book.Closed&&seed.Closed&&!book.Saved&&!seed.Saved,"only tracked workbooks receive discard-close");
            if(scenario=="exclusive")Check(app.QuitCalled&&!failed,"empty owned instance exits cleanly");
            else{
                Check(failed&&!app.QuitCalled&&app.UserControl&&app.Visible&&app.DisplayAlerts&&app.ScreenUpdating,"foreign or unknown collection is handed back, never quit");
                if(foreign!=null)Check(!foreign.Closed&&!foreign.Saved&&app.Workbooks.Items.Contains(foreign),"unrelated workbook remains untouched");
            }
            backend.Dispose();Check(true,"cleanup remains idempotent after hand-back");
        }
        Console.WriteLine("OWNERSHIP ASSERTIONS="+checks);
    }
}
