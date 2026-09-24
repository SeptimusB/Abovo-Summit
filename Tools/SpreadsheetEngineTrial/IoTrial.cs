using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using DX = DevExpress.Spreadsheet;
using AC = Aspose.Cells;

// Private-file diagnostics only. Never recalculates, executes VBA, repairs or edits a source.
internal static class IoTrial
{
    internal static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = Int32.MaxValue };
    private static string TrialRoot { get { return Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../../../..", "obj/AsposeTrial")); } }
    internal static string InTrial(string path)
    {
        var full = Path.GetFullPath(path);
        if (!full.StartsWith(TrialRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("File operations are restricted to obj/AsposeTrial disposable copies.");
        if (File.Exists(full) && (File.GetAttributes(full) & FileAttributes.ReparsePoint) != 0)
            throw new InvalidOperationException("Reparse-point files are not allowed in the trial.");
        for (var parent = new DirectoryInfo(Path.GetDirectoryName(full)); parent != null; parent = parent.Parent)
            if ((parent.Attributes & FileAttributes.ReparsePoint) != 0)
                throw new InvalidOperationException("Reparse-point paths are not allowed in the trial.");
        return full;
    }
    internal static void WriteJson(string path, object value)
    {
        path = InTrial(path);
        using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
        using (var text = new StreamWriter(stream, new UTF8Encoding(false))) text.Write(Json.Serialize(value));
    }
    private static string Hash(Stream stream)
    {
        using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "");
    }
    internal static string FileHash(string file) { using (var stream = File.OpenRead(file)) return Hash(stream); }
    private static string Digest(Action<BinaryWriter> action)
    {
        using (var hash = SHA256.Create())
        {
            using (var crypto = new CryptoStream(Stream.Null, hash, CryptoStreamMode.Write))
            using (var writer = new BinaryWriter(crypto, Encoding.UTF8)) action(writer);
            return BitConverter.ToString(hash.Hash).Replace("-", "");
        }
    }
    internal static int Run(string[] args)
    {
        CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
        if (args.Length == 4 && args[0] == "compare-formulas")
        {
            WriteJson(args[3], CompareFormulas(InTrial(args[1]), InTrial(args[2])));
            return 0;
        }
        if (args.Length == 4 && args[0] == "compare-styles")
        {
            WriteJson(args[3], CompareStyles(InTrial(args[1]), InTrial(args[2])));
            return 0;
        }
        if (args.Length == 3 && (args[0] == "inspect" || args[0] == "inspect-core"))
        {
            WriteJson(args[2], Inspect(InTrial(args[1]), args[0] == "inspect"));
            return 0;
        }
        if (args.Length != 5 || args[0] != "io" || (args[1] != "aspose" && args[1] != "devexpress"))
            throw new ArgumentException("Use: io aspose|devexpress private-input private-output new-result.json; or inspect private-input new-manifest.json");
        var source = InTrial(args[2]); var output = InTrial(args[3]); var resultPath = InTrial(args[4]);
        if (File.Exists(output) || File.Exists(resultPath) || source == output)
            throw new InvalidOperationException("Outputs must be new files; overwrite is prohibited.");
        var ext = Path.GetExtension(output).ToLowerInvariant();
        if (ext != ".xlsb" && ext != ".xlsm") throw new ArgumentException("Only XLSB/XLSM are tested.");
        var result = new Dictionary<string, object> {
            {"engine", args[1]}, {"bits", IntPtr.Size * 8}, {"input", source}, {"output", output},
            {"calculationRequested", false}, {"utc", DateTime.UtcNow.ToString("o")},
            {"scope", "Native load/serialization only; not Summit end-to-end save or financial parity"}
        };
        var before = FileHash(source);
        try
        {
            if (args[1] == "aspose")
            {
                using (var licenceProbe = new AC.Workbook())
                    if (!licenceProbe.IsLicensed) throw new InvalidOperationException("Real-model trials require an activated licence.");
                result["licensed"] = true;
            }
            using (var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var destination = new FileStream(output, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                if (args[1] == "aspose")
                {
                    var timer = Stopwatch.StartNew();
                    using (var book = new AC.Workbook(input))
                    {
                        result["loadMs"] = timer.ElapsedMilliseconds;
                        result["version"] = AC.CellsHelper.GetVersion();
                        timer.Restart(); book.Save(destination, ext == ".xlsb" ? AC.SaveFormat.Xlsb : AC.SaveFormat.Xlsm);
                        result["saveMs"] = timer.ElapsedMilliseconds;
                    }
                }
                else
                {
                    using (var book = new DX.Workbook())
                    {
                        book.Options.CalculationMode = DX.WorkbookCalculationMode.Manual;
                        var timer = Stopwatch.StartNew();
                        if (!book.LoadDocument(input, Path.GetExtension(source).ToLowerInvariant() == ".xlsb" ? DX.DocumentFormat.Xlsb : DX.DocumentFormat.Xlsm))
                            throw new InvalidDataException("DevExpress import failed.");
                        result["loadMs"] = timer.ElapsedMilliseconds;
                        book.Options.CalculationMode = DX.WorkbookCalculationMode.Manual;
                        result["version"] = typeof(DX.Workbook).Assembly.GetName().Version.ToString();
                        result["compatibilityPreflightIncluded"] = false;
                        // Raw-engine comparator; production CONCATENATE/metadata guards are NOT run.
                        timer.Restart(); book.SaveDocument(destination, ext == ".xlsb" ? DX.DocumentFormat.Xlsb : DX.DocumentFormat.Xlsm);
                        result["saveMs"] = timer.ElapsedMilliseconds;
                    }
                }
            }
            result["outputBytes"] = new FileInfo(output).Length;
            result["success"] = true;
        }
        catch (Exception error)
        {
            result["success"] = false; result["error"] = error.GetType().Name + ": " + error.Message;
            throw;
        }
        finally
        {
            result["sourceHash"] = before; result["sourceUnchanged"] = before == FileHash(source);
            result["peakWorkingSetBytes"] = Process.GetCurrentProcess().PeakWorkingSet64;
            WriteJson(resultPath, result);
            Console.WriteLine(Json.Serialize(result));
            if (!(bool)result["sourceUnchanged"]) throw new InvalidOperationException("Source hash changed.");
        }
        return 0;
    }
    private static string Value(DX.CellValue value)
    {
        if (value.IsNumeric) return "N:" + value.NumericValue.ToString("R", CultureInfo.InvariantCulture);
        if (value.IsText) return "T:" + value.TextValue;
        if (value.IsBoolean) return "B:" + value.BooleanValue;
        if (value.IsError) return "E:" + value.ErrorValue;
        return value.Type + ":" + value.ToString();
    }
    private static void Position(BinaryWriter w, DX.Cell c) { w.Write(c.RowIndex); w.Write(c.ColumnIndex); }
    private static object CompareFormulas(string before, string after)
    {
        var differences = new List<object>();
        var arraySamples = new List<object>();
        using (var a = new DX.Workbook())
        using (var b = new DX.Workbook())
        {
            a.Options.CalculationMode = b.Options.CalculationMode = DX.WorkbookCalculationMode.Manual;
            if (!a.LoadDocument(before) || !b.LoadDocument(after)) throw new InvalidDataException("Import failed.");
            a.Options.CalculationMode = b.Options.CalculationMode = DX.WorkbookCalculationMode.Manual;
            foreach (DX.Worksheet sheet in a.Worksheets)
            {
                var other = b.Worksheets[sheet.Name]; int count = 0;
                var samples = new List<object>();
                foreach (var c in sheet.GetUsedRange().ExistingCells.Where(c=>c.HasFormula))
                {
                    var x = c.FormulaInvariant; var y = other.Cells[c.RowIndex,c.ColumnIndex].FormulaInvariant;
                    if (x == y) continue;
                    count++;
                    if (samples.Count < 3) samples.Add(new { cell = c.GetReferenceA1(), before = x, after = y });
                }
                if (count > 0) differences.Add(new { sheet = sheet.Name, count, samples });
                var oldArrays = sheet.ArrayFormulas.Cast<DX.ArrayFormula>().ToDictionary(x=>x.Range.GetReferenceA1(), x=>x.Formula);
                foreach(var d in other.DynamicArrayFormulas.Cast<DX.DynamicArrayFormula>())
                {
                    string oldFormula;
                    if (oldArrays.TryGetValue(d.Range.GetReferenceA1(), out oldFormula))
                    {
                        arraySamples.Add(new { sheet = sheet.Name, range = d.Range.GetReferenceA1(), sameFormula = oldFormula == d.Formula });
                        break;
                    }
                }
            }
        }
        return new { before, after, differences, arraySamples, limits = "Diagnostic formula text, not calculated financial equivalence. Array samples are one range per affected sheet." };
    }
    private static object CompareStyles(string before, string after)
    {
        var counts = new SortedDictionary<string, int>();
        var samples = new SortedDictionary<string, List<object>>();
        var patternPairs = new SortedDictionary<string, int>();
        int solidGateChanges = 0;
        int checkedCells = 0;
        using (var a = new DX.Workbook())
        using (var b = new DX.Workbook())
        {
            a.Options.CalculationMode = b.Options.CalculationMode = DX.WorkbookCalculationMode.Manual;
            if (!a.LoadDocument(before) || !b.LoadDocument(after)) throw new InvalidDataException("Independent import failed.");
            a.Options.CalculationMode = b.Options.CalculationMode = DX.WorkbookCalculationMode.Manual;
            foreach (DX.Worksheet sheet in a.Worksheets)
            {
                var target = b.Worksheets[sheet.Name];
                if (target == null) throw new InvalidDataException("Missing sheet.");
                foreach (var c in sheet.GetUsedRange().ExistingCells.Where(c => c.HasFormula || !c.Value.IsEmpty))
                {
                    var d = target.Cells[c.RowIndex, c.ColumnIndex];
                    Action<string, object, object> compare = (field, x, y) => {
                        if (Object.Equals(x,y)) return;
                        if (!counts.ContainsKey(field)) { counts[field] = 0; samples[field] = new List<object>(); }
                        counts[field]++;
                        if (field == "fillPattern")
                        {
                            var pair = x + " -> " + y;
                            patternPairs[pair] = patternPairs.ContainsKey(pair) ? patternPairs[pair] + 1 : 1;
                            if (Object.Equals(x, DX.PatternType.Solid) != Object.Equals(y, DX.PatternType.Solid)) solidGateChanges++;
                        }
                        if (samples[field].Count < 5) samples[field].Add(new { sheet = sheet.Name, cell = c.GetReferenceA1(), before = x.ToString(), after = y.ToString() });
                    };
                    compare("locked", c.Protection.Locked, d.Protection.Locked);
                    compare("hiddenFormula", c.Protection.Hidden, d.Protection.Hidden);
                    compare("numberFormat", c.NumberFormat, d.NumberFormat);
                    compare("fillPattern", c.Fill.PatternType, d.Fill.PatternType);
                    compare("fillBackground", c.Fill.BackgroundColor.ToArgb(), d.Fill.BackgroundColor.ToArgb());
                    compare("fillPatternColor", c.Fill.PatternColor.ToArgb(), d.Fill.PatternColor.ToArgb());
                    compare("fontName", c.Font.Name, d.Font.Name);
                    compare("fontSize", c.Font.Size, d.Font.Size);
                    compare("fontBold", c.Font.Bold, d.Font.Bold);
                    compare("fontItalic", c.Font.Italic, d.Font.Italic);
                    compare("fontColor", c.Font.Color.ToArgb(), d.Font.Color.ToArgb());
                    compare("horizontalAlignment", c.Alignment.Horizontal, d.Alignment.Horizontal);
                    compare("verticalAlignment", c.Alignment.Vertical, d.Alignment.Vertical);
                    compare("wrap", c.Alignment.WrapText, d.Alignment.WrapText);
                    checkedCells++;
                }
            }
        }
        return new { before, after, checkedCells, counts, samples, patternPairs, solidGateChanges, limits = "Populated cells only; does not audit blank-cell styles, validation, conditional-format rules or drawings." };
    }
    private static object Inspect(string file, bool includeStyles)
    {
        var sheets = new List<object>(); var names = new SortedDictionary<string, string>(StringComparer.Ordinal);
        using (var book = new DX.Workbook())
        {
            book.Options.CalculationMode = DX.WorkbookCalculationMode.Manual;
            if (!book.LoadDocument(file)) throw new InvalidDataException("Independent reader failed.");
            book.Options.CalculationMode = DX.WorkbookCalculationMode.Manual;
            foreach (DX.DefinedName n in book.DefinedNames) names["workbook:" + n.Name] = n.RefersTo + "|hidden=" + n.Hidden;
            foreach (DX.Worksheet sheet in book.Worksheets)
            {
                foreach (DX.DefinedName n in sheet.DefinedNames) names[sheet.Name + ":" + n.Name] = n.RefersTo + "|hidden=" + n.Hidden;
                var cells = sheet.GetUsedRange().ExistingCells.Where(c => c.HasFormula || !c.Value.IsEmpty).OrderBy(c => c.RowIndex).ThenBy(c => c.ColumnIndex).ToArray();
                var row = new Dictionary<string, object> {
                    {"name", sheet.Name}, {"visibility", sheet.VisibilityType.ToString()}, {"protected", sheet.IsProtected},
                    {"populatedCells", cells.Length}, {"formulaCells", cells.Count(c=>c.HasFormula)},
                    {"formulas", Digest(w => { foreach(var c in cells.Where(c=>c.HasFormula)) { Position(w,c); w.Write(c.FormulaInvariant); } })},
                    // AGL has many array ranges: inspect each range once, not for every cell.
                    {"arrayFormulaRanges", sheet.ArrayFormulas.Cast<DX.ArrayFormula>().Count()},
                    {"arrayFormulas", Digest(w => { foreach(var a in sheet.ArrayFormulas.Cast<DX.ArrayFormula>().OrderBy(a=>a.Range.GetReferenceA1(), StringComparer.Ordinal)) { w.Write(a.Range.GetReferenceA1()); w.Write(a.Formula); } })},
                    {"dynamicArrayFormulaRanges", sheet.DynamicArrayFormulas.Cast<DX.DynamicArrayFormula>().Count()},
                    {"dynamicArrayFormulas", Digest(w => { foreach(var a in sheet.DynamicArrayFormulas.Cast<DX.DynamicArrayFormula>().OrderBy(a=>a.Range.GetReferenceA1(), StringComparer.Ordinal)) { w.Write(a.Range.GetReferenceA1()); w.Write(a.Formula); } })},
                    {"constants", Digest(w => { foreach(var c in cells.Where(c=>!c.HasFormula)) { Position(w,c); w.Write(Value(c.Value)); } })},
                    {"cachedFormulaValues", Digest(w => { foreach(var c in cells.Where(c=>c.HasFormula)) { Position(w,c); w.Write(Value(c.Value)); } })},
                    {"populatedCellAppearanceAndLocks", includeStyles ? Digest(w => { foreach(var c in cells) {
                        Position(w,c); w.Write(c.Protection.Locked); w.Write(c.Protection.Hidden); w.Write(c.NumberFormat);
                        w.Write(c.Fill.PatternType.ToString()); w.Write(c.Fill.BackgroundColor.ToArgb()); w.Write(c.Fill.PatternColor.ToArgb());
                        w.Write(c.Font.Name); w.Write(c.Font.Size); w.Write(c.Font.Bold); w.Write(c.Font.Italic); w.Write(c.Font.Color.ToArgb());
                        w.Write(c.Alignment.Horizontal.ToString()); w.Write(c.Alignment.Vertical.ToString()); w.Write(c.Alignment.WrapText);
                    } }) : null}
                };
                sheets.Add(row);
                if (sheets.Count % 25 == 0) Console.WriteLine("INSPECT progress " + sheets.Count + "/" + book.Worksheets.Count);
            }
        }
        var package = new SortedDictionary<string, string>(StringComparer.Ordinal);
        using (var zip = ZipFile.OpenRead(file))
            foreach (var entry in zip.Entries.Where(e => e.FullName.StartsWith("customXml/", StringComparison.Ordinal) || e.FullName.EndsWith("vbaProject.bin", StringComparison.Ordinal) || e.FullName.Contains("metadata")))
                using (var stream = entry.Open()) package[entry.FullName] = Hash(stream);
        Console.WriteLine("INSPECT " + Path.GetFileName(file) + " sheets=" + sheets.Count + " names=" + names.Count);
        return new { schema = 2, file, sha256 = FileHash(file), reader = "DevExpress 25.2.4; manual; no calculation", sheets, names, package,
            stylesIncluded = includeStyles,
            limits = "Formatting, when included, covers populated cells only. Not full validation/conditional-format/chart/blank-style/VBA execution or financial equivalence proof. Package differences require investigation." };
    }
}
