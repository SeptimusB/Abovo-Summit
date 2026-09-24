using System;
using System.Linq;
using System.Reflection;

// Isolated public-API discovery; never called by Summit.
internal static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        try
        {
            if (args.Length > 0 && args[0] == "bench") return EngineBenchmark.Run(args);
            if (args.Length > 0 && args[0] == "smoke") return EngineBenchmark.SpreadSmoke();
            if (args.Length > 0 && args[0] == "io") return EngineBenchmark.SpreadIo(args);
            if (args.Length == 2 && args[0] == "cases") return EngineBenchmark.SpreadCases(args[1]);
            if (args.Length > 0 && args[0] == "edit") return EngineBenchmark.SpreadEdit(args);
            if (args.Length > 0 && args[0] == "diagnose") return EngineBenchmark.SpreadDiagnose(args);
            if (args.Length > 0 && args[0] == "verify-excel") return EngineBenchmark.SpreadVerifyExcel(args);
            if (args.Length > 0 && args[0] == "excel-open") return EngineBenchmark.SpreadExcelOpen(args);
            if (args.Length > 0 && new[]{"inspect","inspect-core","compare-styles","compare-formulas"}.Contains(args[0]))return IoTrial.Run(args);
            if (args.Length == 3 && args[0] == "batch")return EngineBenchmark.SpreadBatch(args[1],args[2]);
            var assembly = typeof(GrapeCity.Spreadsheet.IWorkbook).Assembly;
            Console.WriteLine(assembly.FullName);
            foreach (var type in assembly.GetExportedTypes().Concat(typeof(GrapeCity.CalcEngine.Function).Assembly.GetExportedTypes()).Concat(typeof(FarPoint.Win.Spread.FpSpread).Assembly.GetExportedTypes()).Where(t => args.Length == 0
                ? t.Name.IndexOf("Factory", StringComparison.OrdinalIgnoreCase) >= 0
                : args.Contains(t.FullName)))
            {
                Console.WriteLine("TYPE " + type.FullName + " base="+type.BaseType+" abstract=" + type.IsAbstract);
                foreach (var member in type.GetMembers(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly))
                    Console.WriteLine(member);
            }
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
