using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Gear = SpreadsheetGear;
using DX = DevExpress.Spreadsheet;

// Isolated evaluation only. Does not reference Summit or open a client workbook.
internal static class SpreadsheetGearTrial
{
    private const string LicenceVariable = "SUMMIT_SPREADSHEETGEAR_SIGNED_LICENSE";
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("Abovo.Summit.IsolatedSpreadsheetGearTrial.v1");
    private static int assertions;

    internal static string LicencePath
    {
        get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Abovo", "SummitEngineTrials", "SpreadsheetGear.lic.dpapi"); }
    }

    internal static int Run(string command)
    {
        if (command != "sg-smoke" && command != "sg-license-install")
            throw new ArgumentException("Use sg-smoke or sg-license-install.");
        Activate(command == "sg-license-install");
        Console.WriteLine("SpreadsheetGear " + typeof(Gear.Factory).Assembly.GetName().Version +
            "; process=" + (IntPtr.Size * 8) + "-bit; signed licence accepted.");
        if (command == "sg-license-install")
        {
            Console.WriteLine("Licence stored encrypted for the current Windows user, outside the repository.");
            return 0;
        }
        return Smoke();
    }

    internal static void Activate(bool install = false)
    {
        byte[] plain = null;
        try
        {
            string signed = Environment.GetEnvironmentVariable(LicenceVariable);
            if (String.IsNullOrWhiteSpace(signed))
            {
                if (install) throw new InvalidOperationException("Installation requires the process-local licence variable.");
                ValidatePath(LicencePath);
                if (!File.Exists(LicencePath)) throw new InvalidOperationException("No isolated SpreadsheetGear licence is installed.");
                plain = ProtectedData.Unprotect(File.ReadAllBytes(LicencePath), Entropy, DataProtectionScope.CurrentUser);
                signed = Encoding.UTF8.GetString(plain);
            }
            try
            {
                // Must precede every other Factory call. Never log vendor licence errors or the key.
                Gear.Factory.SetSignedLicense(signed.Trim());
            }
            catch
            {
                throw new InvalidOperationException("SpreadsheetGear rejected the signed licence; no unlicensed fallback was used.");
            }
            if (install)
            {
                ValidatePath(LicencePath);
                if (File.Exists(LicencePath)) throw new InvalidOperationException("An encrypted trial licence already exists; it was not overwritten.");
                Directory.CreateDirectory(Path.GetDirectoryName(LicencePath));
                ValidatePath(LicencePath);
                plain = Encoding.UTF8.GetBytes(signed.Trim());
                var encrypted = ProtectedData.Protect(plain, Entropy, DataProtectionScope.CurrentUser);
                using (var stream = new FileStream(LicencePath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    stream.Write(encrypted, 0, encrypted.Length);
            }
        }
        finally
        {
            if (plain != null) Array.Clear(plain, 0, plain.Length);
        }
    }

    private static void ValidatePath(string file)
    {
        if (File.Exists(file) && (File.GetAttributes(file) & FileAttributes.ReparsePoint) != 0)
            throw new InvalidOperationException("A reparse-point licence file is not permitted.");
        for (var dir = new DirectoryInfo(Path.GetDirectoryName(file)); dir != null; dir = dir.Parent)
            if (dir.Exists && (dir.Attributes & FileAttributes.ReparsePoint) != 0)
                throw new InvalidOperationException("A reparse-point licence directory is not permitted.");
    }

    private static void Require(bool condition, string description)
    {
        if (!condition) throw new InvalidOperationException("Smoke test failed: " + description);
        assertions++;
        Console.WriteLine("PASS " + description);
    }

    private static int Smoke()
    {
#if TRIAL_X64
        Require(IntPtr.Size == 8, "Actual process is 64-bit");
#else
        Require(IntPtr.Size == 4, "Actual process is 32-bit");
#endif
        Console.WriteLine("Synthetic in-memory tests only. No client file, Excel instance or production setting is touched.");
        var set = Gear.Factory.GetWorkbookSet(CultureInfo.InvariantCulture);
        set.Calculation = Gear.Calculation.Manual;
        set.CalculationOnDemand = false;
        set.BackgroundCalculation = false;
        set.EnableWebService = false;
        var book = set.Workbooks.Add();
        try
        {
            var input = book.Worksheets[0]; input.Name = "Inputs";
            var output = book.Worksheets.Add(); output.Name = "Output";
            input.Cells["A1"].Value = 12.5;
            input.Cells["B1"].Value = 3.0;
            input.Cells["C1"].Value = new DateTime(2026, 9, 24).ToOADate();
            input.Cells["C1"].NumberFormat = "dd-mmm-yyyy";
            book.Names.Add("Rate", "=Inputs!$B$1");
            output.Cells["A1"].Formula = "=Inputs!A1*Rate";
            set.CalculateFullRebuild();
            Require(Convert.ToDouble(output.Cells["A1"].Value) == 37.5, "Cross-sheet and named-range calculation");
            input.Cells["A1"].Value = 10.0;
            set.Calculate();
            Require(Convert.ToDouble(output.Cells["A1"].Value) == 30.0, "Incremental recalculation after input edit");
            input.Cells["1:2"].Insert();
            set.Calculate();
            Require(output.Cells["A1"].Formula.Contains("A3"), "Row insertion fixes cross-sheet formula references");
            Require(book.Names["Rate"].RefersTo.Contains("$B$3"), "Row insertion fixes defined names");
            Require(Convert.ToDouble(output.Cells["A1"].Value) == 30.0, "Calculation remains valid after row insertion");
            input.Cells["1:2"].Delete();
            input.Cells["A:B"].Insert();
            set.Calculate();
            Require(output.Cells["A1"].Formula.Contains("C1"), "Column insertion fixes cross-sheet references");
            Require(book.Names["Rate"].RefersTo.Contains("$D$1"), "Column insertion fixes defined names");
            input.Cells["A:B"].Delete();
            set.CalculateFull();
            Require(output.Cells["A1"].Formula == "=Inputs!A1*Rate" &&
                book.Names["Rate"].RefersTo == "=Inputs!$B$1", "Add/delete restores original formula and defined name");

            var limits = book.Worksheets.Add(); limits.Name = "TrialLimits";
            limits.Cells["A1001"].Value = 7.0;
            limits.Cells["CW1"].Value = 9.0;
            Require(Convert.ToDouble(limits.Cells["A1001"].Value) == 7.0, "Trial exceeds free 1000-row limit");
            Require(Convert.ToDouble(limits.Cells["CW1"].Value) == 9.0, "Trial exceeds free 100-column limit");
            while (book.Worksheets.Count < 11) book.Worksheets.Add();
            Require(book.Worksheets.Count == 11, "Trial exceeds free 10-worksheet limit");
            for (int i = 0; i < 3; i++) set.Workbooks.Add();
            Require(set.Workbooks.Count == 4, "Trial exceeds free 3-workbook limit");

            byte[] bytes = book.SaveToMemory(Gear.FileFormat.OpenXMLWorkbookMacroEnabled);
            Require(bytes.Length > 0, "XLSM serializes in memory");
            book.Close();
            book = set.Workbooks.OpenFromMemory(bytes);
            set.CalculateFullRebuild();
            Require(Convert.ToDouble(book.Worksheets["Output"].Cells["A1"].Value) == 30.0, "XLSM reopens and recalculates");
            Require(book.Worksheets["Inputs"].Cells["C1"].NumberFormat == "dd-mmm-yyyy", "XLSM preserves date number format");
            using (var other = new DX.Workbook())
            using (var stream = new MemoryStream(bytes, false))
            {
                other.Options.CalculationMode = DX.WorkbookCalculationMode.Manual;
                Require(other.LoadDocument(stream, DX.DocumentFormat.Xlsm), "DevExpress independently opens synthetic XLSM");
                Require(other.Worksheets["Output"].Cells["A1"].HasFormula, "Independent reader sees the saved formula");
                Require(other.Worksheets["Inputs"].Cells["C1"].Value.NumericValue == new DateTime(2026, 9, 24).ToOADate(),
                    "Independent reader sees the numeric date");
            }
            Console.WriteLine("PASS " + assertions + " installation assertions. Not a client-workbook compatibility or performance result.");
            return 0;
        }
        finally
        {
            while (set.Workbooks.Count > 0) set.Workbooks[0].Close();
        }
    }
}
