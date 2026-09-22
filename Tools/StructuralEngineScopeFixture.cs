using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using DevExpress.Spreadsheet;

// Private workbook and native controls. Never saves the source or executes VBA.
// Run through Test-ColumnFamily.ps1 -Fixture StructuralEngineScopeFixture.cs.
public static class StructuralEngineScopeFixture {
    const BindingFlags F = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static;
    static object Field(object o, string name) { return o.GetType().GetField(name, F).GetValue(o); }
    static void Check(bool value, string message) { if (!value) throw new Exception(message); Console.WriteLine("PASS: " + message); }
    [STAThread] public static int Main(string[] args) {
        try {
            AppDomain.CurrentDomain.AssemblyResolve += (s,e) => {
                var path = Path.Combine(args[0],new AssemblyName(e.Name).Name + ".dll");
                return File.Exists(path) ? Assembly.LoadFrom(path) : null;
            };
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
            var app = Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
            app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null,null);
            var files = app.GetType("Abovo.FileManager");
            files.GetMethod("Initialise").Invoke(null,new object[]{null});
            string copy = Path.Combine(args[2],"scope-copy.xlsb"); File.Copy(args[1],copy);
            var open = files.GetMethod("OpenModel");
            dynamic loaded = open.Invoke(null,new object[]{copy,new FileInfo(copy),Enum.ToObject(open.GetParameters()[2].ParameterType,0)});
            Check(!loaded.BError,"Opened private workbook");
            dynamic model = ((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);
            IWorkbook w = model.WB;
            var protection = w.Worksheets.ToDictionary(s=>s.Name,s=>s.IsProtected);
            var visibility = w.Worksheets.ToDictionary(s=>s.Name,s=>s.Visible);
            var mode = w.Options.CalculationMode; var engine = w.Options.CalculationEngineType;
            bool history = w.History.IsEnabled;
            model.WBCalcEngine.CalculateDependencySensitiveFile("Engine scope fixture",true);
            var snapshots = app.GetType("Abovo.TransactionalDBSnapshotManager");
            snapshots.GetMethod("CreateSnapshotAndComparison").Invoke(null,new object[]{0});
            Check((bool)snapshots.GetMethod("HasValidSnapshot").Invoke(null,new object[]{0}),"Snapshot valid before insertion");
            using (var host = new Form{Opacity=0,ShowInTaskbar=false,ClientSize=new Size(1500,900)})
            using (var analyser = (Control)Activator.CreateInstance(app.GetType("BPIncomeExpenditureAnalyserV2"),new object[]{0,null})) {
                model.ExpendAnalyserV2 = (dynamic)analyser;
                analyser.Dock=DockStyle.Fill;host.Controls.Add(analyser);host.Show();Application.DoEvents();
                Check(Field(analyser,"DSAnalDataRange")!=null,"Analyser has a live binding before insertion");
                using (var listener = new TextWriterTraceListener(Path.Combine(args[2],"scope-trace.log"))) {
                    Trace.Listeners.Add(listener);
                    try {
                        dynamic result = model.WorkbookStructureRules.AddRecords(args[3],int.Parse(args[4]));
                        Check(!result.BError,"Insertion with visible analyser succeeds: " + result.StringReturn);
                    } finally { listener.Flush();Trace.Listeners.Remove(listener); }
                }
                Check(w.Options.CalculationMode==mode && w.Options.CalculationEngineType==engine && w.History.IsEnabled==history,"Entry calculation/history restored before UI refresh");
                Check(!(bool)app.GetType("Abovo.ModelSafetyManager").GetMethod("IsBulkWorkbookMutationInProgress").Invoke(null,new object[]{0}),"Nested bulk guards released");
                Check(!(bool)Field((object)model.WorkbookStructureRules,"IsExecuting"),"Execution guard released");
                Check((bool)Field(analyser,"StructuralRefreshDeferred"),"Analyser explicitly stale after insertion");
                Check(Field(analyser,"DSAnalDataRange")==null,"Stale live range binding disconnected");
                Check(!(bool)snapshots.GetMethod("HasValidSnapshot").Invoke(null,new object[]{0}),"Structural change invalidates snapshot");
                analyser.GetType().GetMethod("RefreshDeferredIfNeeded").Invoke(analyser,null);
                Check(!(bool)Field(analyser,"StructuralRefreshDeferred") && Field(analyser,"DSAnalDataRange")!=null,"Explicit refresh restores current live analyser");
                Check(w.Options.CalculationMode==mode && w.Options.CalculationEngineType==engine && w.History.IsEnabled==history,"Refresh retains entry calculation/history settings");
                model.ExpendAnalyserV2=null;
            }
            Check(w.Worksheets.All(s=>s.IsProtected==protection[s.Name] && s.Visible==visibility[s.Name]),"Worksheet protection/visibility unchanged");
            Console.WriteLine("PASS: Engine-sharing analyser/snapshot lifecycle. Original untouched.");return 0;
        } catch (Exception e) { Console.Error.WriteLine(e);return 1; }
    }
}
