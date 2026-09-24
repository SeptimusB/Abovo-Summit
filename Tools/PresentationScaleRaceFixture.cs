using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

class PresentationScaleRaceFixture {
    const BindingFlags Flags = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
    static Type manager;
    static int assertions;
    static object Call(string name, params object[] args) { return manager.GetMethod(name, Flags).Invoke(null, args); }
    static void Check(bool value, string name) { if (!value) throw new Exception(name); assertions++; Console.WriteLine("PASS: " + name); }
    static void Idle() { Call("ApplyScaleToNewForms", null, EventArgs.Empty); }

    public class ProbeForm : Form {
        public int ScaleCalls;
        public Action DuringScale;
        protected override bool ShowWithoutActivation { get { return true; } }
        public ProbeForm() { ShowInTaskbar=false; Opacity=0; StartPosition=FormStartPosition.Manual; Location=new Point(-20000,-20000); }
        public void ApplyPresentationScale() { ScaleCalls++; if (DuringScale != null) DuringScale(); }
    }

    class ChurningForms : IEnumerable {
        readonly Form first, second;
        readonly int failures;
        public int Attempts;
        public ChurningForms(Form a, Form b, int failures) { first=a; second=b; this.failures=failures; }
        public IEnumerator GetEnumerator() {
            Attempts++;
            yield return first;
            if (Attempts <= failures) throw new InvalidOperationException("Collection was modified; enumeration operation may not execute.");
            yield return second;
        }
    }

    [STAThread] static int Main(string[] args) { try {
        AppDomain.CurrentDomain.AssemblyResolve += (s,e) => { string p=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll"); return File.Exists(p)?Assembly.LoadFrom(p):null; };
        var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
        manager=app.GetType("Abovo.PresentationScaleManager");
        app.GetType("Abovo.FontManager").GetMethod("InitialiseApplication").Invoke(null,null);
        int oldPercent=(int)manager.GetProperty("InterfaceScalePercent").GetValue(null,null);
        try { Call("SetInterfaceScale",100,false); Exercise(); }
        finally { Call("SetInterfaceScale",oldPercent,false); }
        Console.WriteLine("PASS: "+assertions+" scale-race checks; no workbook opened or settings saved");
        return 0;
    } catch(Exception ex) { Console.Error.WriteLine(ex); return 1; } }

    static void Exercise() {
        using(var first=new ProbeForm()) using(var second=new ProbeForm()) {
            bool reproduced=false;
            try { foreach(Form f in new ChurningForms(first,second,1)) { } }
            catch(InvalidOperationException) { reproduced=true; }
            Check(reproduced,"Original enumeration pattern reproduces the reported exception");
            var churn=new ChurningForms(first,second,1);
            var snapshot=(List<Form>)Call("SnapshotOpenForms",churn);
            Check(churn.Attempts==2 && snapshot.Count==2 && snapshot[0]==first && snapshot[1]==second,"A changing snapshot retries without retaining a partial copy");
            churn=new ChurningForms(first,second,Int32.MaxValue);
            snapshot=(List<Form>)Call("SnapshotOpenForms",churn);
            Check(churn.Attempts==3 && snapshot.Count==0,"Sustained changes defer after three bounded attempts without throwing");

            first.Show(); second.Show(); Idle();
            Check(first.ScaleCalls==1 && second.ScaleCalls==1,"Existing forms initialise exactly once");
            float original=first.Font.SizeInPoints;
            for(int i=0;i<30;i++) Idle();
            Check(first.ScaleCalls==1 && first.Font.SizeInPoints==original,"Repeated idle passes do not double-scale forms");
            first.DuringScale=Idle;
            Call("SetInterfaceScale",125,false);
            Check(first.ScaleCalls==2 && Math.Abs(first.Font.SizeInPoints-original*1.25f)<0.01f,"Explicit scale change survives re-entrant idle and applies once");
            first.DuringScale=null;
            for(int i=0;i<30;i++) Idle();
            Check(first.ScaleCalls==2,"Queued or repeated work does not replay a scale ratio");
            // Simulate the desired scale changing during a pass that had to defer.
            manager.GetField("CurrentPercent",Flags).SetValue(null,150);
            Idle();
            Check(first.ScaleCalls==3 && Math.Abs(first.Font.SizeInPoints-original*1.5f)<0.01f,"Next idle pass catches up an outstanding scale change");
            using(var late=new ProbeForm()) {
                late.Show(); Idle();
                Check(late.ScaleCalls==1,"A later form receives one initial scale hook at the current scale");
                late.Close(); Call("ApplyCurrentScaleToForm",late);
                Check(late.ScaleCalls==1,"Closed forms are ignored");
            }
            using(var closesItself=new ProbeForm()) {
                closesItself.DuringScale=closesItself.Close;
                closesItself.Show(); Idle();
                Check(closesItself.IsDisposed,"A form may close in its scale hook without a later layout call");
            }
            int before=first.ScaleCalls;
            Exception backgroundError=null;
            int iterations=0;
            using(var started=new ManualResetEvent(false)) {
                var worker=new Thread(() => { try {
                    started.Set();
                    for(int i=0;i<200;i++) {
                        using(var splashLike=new ProbeForm()) { splashLike.Show(); Application.DoEvents(); splashLike.Close(); }
                        Interlocked.Increment(ref iterations);
                    }
                } catch(Exception ex) { backgroundError=ex; } });
                worker.IsBackground=true; worker.SetApartmentState(ApartmentState.STA); worker.Start();
                Check(started.WaitOne(5000),"Secondary UI thread starts");
                DateTime deadline=DateTime.UtcNow.AddSeconds(30);
                while(worker.IsAlive && DateTime.UtcNow<deadline) { Idle(); Application.DoEvents(); Thread.Yield(); }
                Check(worker.Join(1000),"Secondary UI churn completes within its timeout");
            }
            Check(backgroundError==null && iterations==200,"Two UI threads tolerate 200 actual form open/close cycles");
            Idle(); Check(first.ScaleCalls==before,"Background window churn does not rescale stable forms");
            Call("SetInterfaceScale",100,false);
            Check(Math.Abs(first.Font.SizeInPoints-original)<0.01f,"Returning to 100 percent restores the original font size");
        }
    }
}
