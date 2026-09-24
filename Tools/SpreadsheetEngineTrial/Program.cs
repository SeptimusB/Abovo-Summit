using System;
using System.IO;
using Aspose.Cells;

// Standalone experiment. No production integration or original-model writes.
internal static class Program
{
    private static int assertions;

    private static void Require(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
        assertions++;
        Console.WriteLine("PASS " + message);
    }

    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length == 3 && args[0] == "mirror-array-trace") return EngineBenchmark.RunMirrorArrayTrace(args[1],args[2]);
            if (args.Length == 3 && args[0] == "mirror-array-trace-fixed") return EngineBenchmark.RunMirrorArrayTrace(args[1],args[2],true);
            if (args.Length == 2 && args[0] == "mirror-array-insert") return EngineBenchmark.RunMirrorArrayInsertProbe(args[1]);
            if (args.Length == 3 && args[0] == "mirror-excel-array-open") return EngineBenchmark.RunMirrorExcelArrayOpen(args[1],args[2]);
            if (args.Length == 2 && args[0] == "mirror-array-copy") return EngineBenchmark.RunMirrorArrayCopyProbe(args[1]);
            if (args.Length == 3 && args[0] == "mirror-array-state") return EngineBenchmark.RunMirrorArrayState(args[1],args[2]);
            if (args.Length == 3 && args[0] == "mirror-array-native") return EngineBenchmark.RunMirrorNativeArrayState(args[1],args[2]);
            if (args.Length == 3 && args[0] == "gear-import") return GearImportProbe.Run(args[1], args[2]);
            if (args.Length == 3 && args[0] == "gear-projection") return EngineBenchmark.RunGearProjection(args[1], args[2]);
            if (args.Length != 0 && args[0] == "mirror") return EngineBenchmark.RunMirror(args);
            if (args.Length == 2 && args[0] == "mirror-worker") return EngineBenchmark.MirrorWorker(args[1]);
            if (args.Length == 1 && args[0].StartsWith("sg-", StringComparison.Ordinal))
                return SpreadsheetGearTrial.Run(args[0]);

            // Licence contents are never logged, embedded, copied or discovered by scanning.
            var licencePath = Environment.GetEnvironmentVariable("SUMMIT_ASPOSE_LICENSE_PATH");
            var licenceProvided = !String.IsNullOrWhiteSpace(licencePath);
            if (licenceProvided)
            {
                using (var licenceFile = File.OpenRead(licencePath))
                    new License().SetLicense(licenceFile);
            }

            if (args.Length == 2 && args[0] == "funding-3d") return EngineBenchmark.FundingThreeDProbe(args[1]);
            if (args.Length != 0 && args[0] == "funding") return EngineBenchmark.RunFunding(args);
            if (args.Length != 0 && args[0] == "bench") return EngineBenchmark.Run(args);
            if (args.Length != 0) return IoTrial.Run(args);
            Console.WriteLine("Aspose.Cells " + CellsHelper.GetVersion() + "; process=" + (IntPtr.Size * 8) + "-bit; runtime=" + Environment.Version);
#if TRIAL_X64
            Require(IntPtr.Size == 8, "Actual process is 64-bit");
#else
            Require(IntPtr.Size == 4, "Actual process is 32-bit");
#endif
            Console.WriteLine("SYNTHETIC IN-MEMORY TEST ONLY. No client workbook is opened or saved.");
            using (var workbook = new Workbook())
            {
                Console.WriteLine("Licensed=" + workbook.IsLicensed);
                if (licenceProvided && !workbook.IsLicensed)
                    throw new InvalidOperationException("The supplied licence did not activate Aspose.Cells.");
                var inputs = workbook.Worksheets[0];
                inputs.Name = "Inputs";
                inputs.Cells["A1"].PutValue(12.5);
                inputs.Cells["A2"].PutValue(3.0);
                inputs.Cells["A3"].PutValue(new DateTime(2026, 9, 23));
                var dateStyle = inputs.Cells["A3"].GetStyle();
                dateStyle.Custom = "dd-mmm-yyyy";
                inputs.Cells["A3"].SetStyle(dateStyle);
                var output = workbook.Worksheets.Add("Output");
                output.Cells["B1"].Formula = "='Inputs'!A1*'Inputs'!A2";
                workbook.CalculateFormula();
                Require(output.Cells["B1"].DoubleValue == 37.5, "Cross-sheet decimal calculation");
                inputs.Cells["A1"].PutValue(10.0);
                workbook.CalculateFormula();
                Require(output.Cells["B1"].DoubleValue == 30.0, "Recalculation after input change");

                foreach (var format in new[] { SaveFormat.Xlsb, SaveFormat.Xlsm })
                {
                    using (var bytes = new MemoryStream())
                    {
                        workbook.Save(bytes, format);
                        Require(bytes.Length > 0, format + " in-memory serialization");
                        bytes.Position = 0;
                        using (var reopened = new Workbook(bytes))
                        {
                            Require(reopened.Worksheets["Output"].Cells["B1"].IsFormula, format + " formula retained");
                            Require(reopened.Worksheets["Inputs"].Cells["A1"].DoubleValue == 10.0, format + " input retained");
                            Require(reopened.Worksheets["Inputs"].Cells["A3"].DateTimeValue == new DateTime(2026, 9, 23), format + " typed date retained");
                            reopened.CalculateFormula();
                            Require(reopened.Worksheets["Output"].Cells["B1"].DoubleValue == 30.0, format + " reopened calculation");
                            Console.WriteLine(format + " reopened worksheet count=" + reopened.Worksheets.Count + " (evaluation can add its own sheet).");
                        }
                    }
                }
            }
            Console.WriteLine("PASS " + assertions + " installation assertions. NOT an Abovo compatibility or speed result.");
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(exception.GetType().Name + ": " + exception.Message);
            return 1;
        }
    }
}
