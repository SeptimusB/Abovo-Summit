// Disposable-file diagnostics only. No production Save policy or source workbook writes.
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Diagnostics;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Web.Script.Serialization;
using DevExpress.Spreadsheet;

public static class WorkbookFormatBenchmark {
    public class Sample {
        public string Engine="DevExpress", Format, Input, Output, CalculationEngine;
        public int Run, ProcessBits;
        public long OpenMs, SaveMs, PreparationMs, DisposeMs, Bytes, SavedFileOpenMs;
        public bool? StructureAndFormulaHashMatch;
        public string ExpectedHash, ReopenedHash;
    }
    static string Digest(IWorkbook w) {
        using(var hash=SHA256.Create()) {
            using(var crypto=new CryptoStream(Stream.Null,hash,CryptoStreamMode.Write))
            using(var writer=new BinaryWriter(crypto,Encoding.UTF8)) {
                foreach(var sheet in w.Worksheets) {
                    writer.Write(sheet.Name);writer.Write(sheet.IsProtected);
                    foreach(var cell in sheet.GetUsedRange().ExistingCells) {
                        if(!cell.HasFormula && cell.Value.IsEmpty)continue;
                        writer.Write(cell.RowIndex);writer.Write(cell.ColumnIndex);
                        writer.Write(cell.HasFormula);
                        writer.Write(cell.HasFormula?cell.FormulaInvariant:cell.Value.Type+":"+cell.Value.ToString());
                    }
                    foreach(var name in sheet.DefinedNames.OrderBy(n=>n.Name,StringComparer.Ordinal)) {
                        writer.Write(name.Name);writer.Write(name.RefersTo);
                    }
                }
                foreach(var name in w.DefinedNames.OrderBy(n=>n.Name,StringComparer.Ordinal)) {
                    writer.Write(name.Name);writer.Write(name.RefersTo);
                }
            }
            return BitConverter.ToString(hash.Hash).Replace("-","");
        }
    }
    [STAThread] public static int Main(string[] args) {
        try {
            AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{
                var path=Path.Combine(args[0],new AssemblyName(e.Name).Name+".dll");
                return File.Exists(path)?Assembly.LoadFrom(path):null;
            };
            var app=Assembly.LoadFrom(Path.Combine(args[0],"Abovo-summit.exe"));
            var guard=app.GetType("Abovo.WorkbookXlsbFormulaCompatibility").GetMethod("Save",BindingFlags.Static|BindingFlags.NonPublic);
            var rows=new List<Sample>();int runs=Int32.Parse(args[2]);
            if(runs==0)rows=new JavaScriptSerializer().Deserialize<List<Sample>>(File.ReadAllText(Path.Combine(args[1],"devexpress.json")));
            for(int run=1;run<=runs;run++)foreach(var ext in (run%2==1?new[]{"xlsb","xlsm"}:new[]{"xlsm","xlsb"})) {
                var row=new Sample {Format=ext,Run=run,ProcessBits=IntPtr.Size*8,
                    Input=Path.Combine(args[1],"baseline."+ext),Output=Path.Combine(args[1],"dx-"+run+"."+ext)};
                var watch=new Stopwatch();var w=new Workbook();
                try {
                    w.Options.CalculationMode=WorkbookCalculationMode.Manual;
                    watch.Start();w.LoadDocument(row.Input);watch.Stop();row.OpenMs=watch.ElapsedMilliseconds;
                    w.Options.CalculationMode=WorkbookCalculationMode.Manual;
                    row.CalculationEngine=w.Options.CalculationEngineType.ToString();
                    // Apply the same compatibility normalization to BOTH formats;
                    // time it separately, not as native serialization speed.
                    var preparation=Stopwatch.StartNew();
                    guard.Invoke(null,new object[]{w,new Func<bool>(()=>{
                        watch.Restart();w.SaveDocument(row.Output,ext=="xlsb"?DocumentFormat.Xlsb:DocumentFormat.Xlsm);
                        watch.Stop();row.SaveMs=watch.ElapsedMilliseconds;return true;
                    }),null,true});
                    row.PreparationMs=preparation.ElapsedMilliseconds-row.SaveMs;
                } finally {watch.Restart();w.Dispose();row.DisposeMs=watch.ElapsedMilliseconds;}
                GC.Collect();GC.WaitForPendingFinalizers();
                row.Bytes=new FileInfo(row.Output).Length;rows.Add(row);
                File.WriteAllText(Path.Combine(args[1],"devexpress.json"),new JavaScriptSerializer().Serialize(rows));
                Console.WriteLine("DX "+ext+" run="+run+" open="+row.OpenMs+" save="+row.SaveMs+" prepare="+row.PreparationMs);
                GC.Collect();GC.WaitForPendingFinalizers();
            }
            // Validate only after all timings, so enumeration does not warm the
            // exporter or add per-cell diagnostic work to any timing sample.
            foreach(var ext in new[]{"xlsb","xlsm"}) {
                var row=rows.First(r=>r.Format==ext);
                using(var expected=new Workbook()) {
                    expected.Options.CalculationMode=WorkbookCalculationMode.Manual;
                    expected.LoadDocument(row.Input);expected.Options.CalculationMode=WorkbookCalculationMode.Manual;
                    guard.Invoke(null,new object[]{expected,new Func<bool>(()=>{
                        Console.WriteLine("VERIFY "+ext+" baseline");
                        row.ExpectedHash=Digest(expected);return true;
                    }),null,true});
                }
                using(var actual=new Workbook()) {
                    actual.Options.CalculationMode=WorkbookCalculationMode.Manual;
                    var reopening=Stopwatch.StartNew();actual.LoadDocument(row.Output);row.SavedFileOpenMs=reopening.ElapsedMilliseconds;
                    actual.Options.CalculationMode=WorkbookCalculationMode.Manual;
                    Console.WriteLine("VERIFY "+ext+" saved");row.ReopenedHash=Digest(actual);
                }
                row.StructureAndFormulaHashMatch=row.ExpectedHash==row.ReopenedHash;
                File.WriteAllText(Path.Combine(args[1],"devexpress.json"),new JavaScriptSerializer().Serialize(rows));
                Console.WriteLine("VERIFY "+ext+" formulasInputsNamesOrderProtection="+row.StructureAndFormulaHashMatch);
                GC.Collect();GC.WaitForPendingFinalizers();
            }
            if(rows.Any(r=>r.StructureAndFormulaHashMatch==false))throw new InvalidDataException("Formula/input/name preservation check failed; inspect the result before recommending either format.");
            return 0;
        } catch(Exception e) {Console.Error.WriteLine(e);return 1;}
    }
}
