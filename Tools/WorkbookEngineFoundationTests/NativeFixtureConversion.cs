using System;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Abovo.WorkbookEngines;

static class NativeFixtureConversion
{
    // Fixture creation only. The source is a generated private XLSX, not a
    // customer workbook. Use the production isolation/security setup; this
    // reflected native SaveAs is deliberately not a product conversion API.
    internal static Task<WorkbookNativeProcessIdentity> ExcelXlsb(string source,string target)
    {
        var result=new TaskCompletionSource<WorkbookNativeProcessIdentity>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread=new Thread(()=>{
            IWorkbookCalculationBackend backend=null;
            try{
                var type=typeof(WorkbookCalculationSession).Assembly.GetType("Abovo.WorkbookEngines.ExcelCalculationBackend");
                backend=(IWorkbookCalculationBackend)Activator.CreateInstance(type,true);
                backend.OpenReadOnly(source,new WorkbookEngineOptions(WorkbookEnginePreference.ExcelRequired,false,false));
                object book=type.GetField("book",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(backend);
                book.GetType().InvokeMember("SaveAs",BindingFlags.InvokeMethod,null,book,new object[]{target,50});
                var identity=(WorkbookNativeProcessIdentity)type.GetProperty("NativeProcess",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(backend,null);
                backend.Dispose();backend=null;result.SetResult(identity);
            }catch(Exception e){try{backend?.Dispose();}catch(Exception cleanup){e=new AggregateException(e,cleanup);}result.SetException(e);}
        }){IsBackground=true};
        thread.SetApartmentState(ApartmentState.STA);thread.Start();return result.Task;
    }
}
