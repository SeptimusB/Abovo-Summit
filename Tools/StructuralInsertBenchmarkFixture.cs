using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

public static class StructuralInsertBenchmarkFixture {
    static Type benchmark;
    static IDisposable Start() {
        return (IDisposable)Activator.CreateInstance(benchmark, new object[]{0,"TEST_INSERT",3,"test=true"});
    }
    static IDisposable Measure(object owner) {
        return (IDisposable)benchmark.GetMethod("Measure").Invoke(owner,new object[]{"testStage","sheet=Test"});
    }
    static void Check(bool ok,string message) {
        if(!ok)throw new Exception(message);
        Console.WriteLine("PASS: "+message);
    }
    sealed class FailingListener : TraceListener {
        public override void Write(string value){throw new IOException("Synthetic diagnostic listener failure");}
        public override void WriteLine(string value){throw new IOException("Synthetic diagnostic listener failure");}
    }
    public static int Main(string[] args) {
        try {
            AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{
                var path=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");
                return File.Exists(path)?Assembly.LoadFrom(path):null;
            };
            benchmark=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe")).GetType("Abovo.StructuralInsertBenchmark");
            using(var output=new StringWriter())using(var listener=new TextWriterTraceListener(output)) {
                Trace.Listeners.Add(listener);
                try {
                    var scope=Start();var stage=Measure(scope);stage.Dispose();stage.Dispose();
                    benchmark.GetMethod("Complete").Invoke(scope,new object[]{false});scope.Dispose();scope.Dispose();
                    using(var failed=Start())benchmark.GetMethod("Complete").Invoke(failed,new object[]{true});
                    try {using(var interrupted=Start())using(Measure(interrupted))throw new InvalidOperationException("Synthetic operation interruption");}
                    catch(InvalidOperationException){}
                    listener.Flush();string log=output.ToString();
                    Check(log.Contains("outcome=ok")&&log.Contains("outcome=failed")&&log.Contains("outcome=interrupted"),"Success, failure and interruption distinguished");
                    var lines=log.Split(new[]{Environment.NewLine},StringSplitOptions.RemoveEmptyEntries);
                    Check(lines.Count(s=>s.Contains("state=finished"))==3,"Repeated Dispose does not duplicate operation totals");
                    Check(lines.Count(s=>s.Contains("stage=testStage")&&s.Contains("elapsed="))==2,"Stage measured once, including exception unwinding");
                    Check(lines.Any(s=>s.Contains(", version=")&&s.Contains("bitness=")),"Release/version and process architecture recorded");
                } finally {Trace.Listeners.Remove(listener);}
            }
            using(var listener=new FailingListener()) {
                Trace.Listeners.Add(listener);
                try {using(var scope=Start()){using(Measure(scope)){}benchmark.GetMethod("Complete").Invoke(scope,new object[]{false});}}
                finally {Trace.Listeners.Remove(listener);}
            }
            Check(true,"Diagnostic listener failure does not escape into operation");return 0;
        } catch(Exception ex){Console.Error.WriteLine(ex);return 1;}
    }
}
