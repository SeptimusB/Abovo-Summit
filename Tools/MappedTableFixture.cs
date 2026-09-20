#define TRACE
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.Spreadsheet;

public static class MappedTableFixture
{
    [STAThread]
    public static int Main(string[] args)
    {
        try { Run(args); return 0; }
        catch (Exception e) { Console.Error.WriteLine(e); return 1; }
    }
    private static void Run(string[] args)
    {
        string bin = args[0];
        AppDomain.CurrentDomain.AssemblyResolve += (sender, eventArgs) => {
            string path = Path.Combine(bin, new AssemblyName(eventArgs.Name).Name + ".dll");
            return File.Exists(path) ? Assembly.LoadFrom(path) : null;
        };
        Assembly app = Assembly.LoadFrom(Path.Combine(bin, "Abovo-summit.exe"));
        Trace.Listeners.Add(new ConsoleTraceListener());
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
        Application.EnableVisualStyles();
        app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null, null);
        Type manager = app.GetType("Abovo.FileManager");
        manager.GetMethod("Initialise").Invoke(null, new object[] { null });
        MethodInfo open = manager.GetMethod("OpenModel");
        object mode = Enum.ToObject(open.GetParameters()[2].ParameterType, 0);
        dynamic result = open.Invoke(null, new object[] { args[1], new FileInfo(args[1]), mode });
        if (result.BError) throw new InvalidOperationException(result.StringReturn);
        Type ditType = app.GetType("DataInterfaceTemplate");
        var watch = Stopwatch.StartNew();
        using (var accounts = (Control)Activator.CreateInstance(ditType,
            new object[] { 0, 0, 37, null, "Normal", null }))
        {
            Console.WriteLine("Accounts constructor: " + watch.ElapsedMilliseconds + " ms");
            if (args.Length > 2 && args[2] == "CheckSheet")
            {
                using (var global = (Control)Activator.CreateInstance(ditType,
                    new object[] { 0, 0, 0, null, "Normal", null }))
                {
                    ditType.GetMethod("BuildSection", BindingFlags.Instance | BindingFlags.NonPublic)
                        .Invoke(global, new object[] { 2, false, false });
                    Console.WriteLine("PASS: Check Sheet DIT section built.");
                    Control grid = FindGrid(global);
                    if (grid == null) throw new Exception("No mapped grid found.");
                    dynamic nativeGrid = grid;
                    dynamic view = nativeGrid.MainView;
                    if (view.RowCount != 57 || view.Columns.Count != 8 ||
                        !view.OptionsBehavior.Editable || !view.OptionsSelection.MultiSelect)
                        throw new Exception("Mapped geometry/selection contract failed.");
                    dynamic model = ((Array)manager.GetField("ExcelModels").GetValue(null)).GetValue(0);
                    IWorkbook workbook = (IWorkbook)model.WB;
                    Worksheet sheet = workbook.Worksheets["Check Sheet"];
                    for (int row = 0; row < 57; row++)
                        for (int column = 0; column < 6; column++)
                            if (view.GetRowCellDisplayText(row, view.Columns[column]) !=
                                sheet.Cells[row + 6, column].DisplayText)
                                throw new Exception("Mapped cell mismatch.");
                    Console.WriteLine("PASS: all 342 A:F values match their worksheet cells; multiselect enabled.");
                    dynamic mapping = model.WBStructure.GroupStructures[0].ChildStructures[0]
                        .InterfaceSections[2].IElements[0].MappedTable;
                    MethodInfo resolver = grid.GetType().GetMethod("FindInterfaces");
                    for (int row = 0; row < 57; row++) {
                        string target = string.IsNullOrWhiteSpace(sheet.Cells[row + 6, 6].DisplayText) ? "" :
                            sheet.Cells[row + 6, 7].DisplayText.Trim();
                        if (view.GetRowCellDisplayText(row, view.Columns[6]) != target)
                            throw new Exception("Explicit worksheet destination mismatch.");
                        string routeText = view.GetRowCellDisplayText(row, view.Columns[7]);
                        if (routeText.Contains(" V2")) throw new Exception("Legacy V2 label visible.");
                        if (target == "Transactional DB" && !routeText.Contains("Outputs / Analysis"))
                            throw new Exception("Primary analyser route missing.");
                        if (target == "Funding Assumptions" && !routeText.Contains("Funding Assumptions"))
                            throw new Exception("Primary funding route missing.");
                    }
                    if (view.Columns[6].Caption != "Worksheet" || view.Columns[7].Caption != "Summit interface")
                        throw new Exception("Explicit destination captions missing.");
                    int linked = 0;
                    for (int row = 8; row < 39; row++) {
                        string target = sheet.Cells[row, 7].DisplayText;
                        if (string.IsNullOrWhiteSpace(target)) continue;
                        var routes = (System.Collections.IList)resolver.Invoke(null,
                            new object[] { model.WBStructure, target, mapping.InterfaceLinks });
                        if (routes.Count == 0) throw new Exception("No Summit link for " + target);
                        if (workbook.Worksheets[target] == null) throw new Exception("Invalid worksheet link.");
                        linked++;
                    }
                    Console.WriteLine("PASS: all " + linked + " column G destinations resolve to Summit and worksheet.");
                    var editors = (System.Collections.IDictionary)grid.GetType()
                        .GetField("choiceEditors", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(grid);
                    if (editors.Count != 5) throw new Exception("Expected exactly five Yes/No inputs.");
                    foreach (string address in new[] { "C22", "C23", "C29", "C33", "C39" })
                        if (!editors.Contains(address)) throw new Exception("Missing Yes/No editor: " + address);
                    MethodInfo postChoice = grid.GetType().GetMethod("PostChoice", BindingFlags.NonPublic | BindingFlags.Instance);
                    Cell input = sheet.Cells["C22"];
                    string original = input.Value.TextValue;
                    string alternate = original == "Yes" ? "No" : "Yes";
                    bool protectedBefore = sheet.IsProtected;
                    // Exercise the actual GridView change event, not a direct workbook write.
                    view.SetRowCellValue(15, view.Columns[2], alternate);
                    if (input.Value.TextValue != alternate || !model.IsDirty || !model.ChangeManager.CanUndo)
                        throw new Exception("Valid combo change did not reach workbook/change history.");
                    dynamic invalid = postChoice.Invoke(grid, new object[] { input, "Maybe" });
                    if (!invalid.BError || input.Value.TextValue != alternate)
                        throw new Exception("Invalid combo choice was accepted.");
                    invalid = postChoice.Invoke(grid, new object[] { sheet.Cells["B22"], "Yes" });
                    if (!invalid.BError) throw new Exception("Formula cell was editable.");
                    dynamic undo = model.ChangeManager.Undo();
                    if (undo.BError || input.Value.TextValue != original ||
                        view.GetRowCellDisplayText(15, view.Columns[2]) != original || !model.ChangeManager.CanRedo)
                        throw new Exception("Undo did not refresh the mapped value.");
                    dynamic redo = model.ChangeManager.Redo();
                    if (redo.BError || input.Value.TextValue != alternate ||
                        view.GetRowCellDisplayText(15, view.Columns[2]) != alternate)
                        throw new Exception("Redo did not refresh the mapped value.");
                    model.ChangeManager.Undo();
                    if (input.Value.TextValue != original || sheet.IsProtected != protectedBefore)
                        throw new Exception("Input/protection restoration failed.");
                    Console.WriteLine("PASS: five workbook-defined combos; edit, invalid/formula rejection, dirty marking, undo/redo and protection.");
                    view.SelectCell(2, view.Columns[0]);
                    nativeGrid.RefreshData();
                    if (!view.IsCellSelected(2, view.Columns[0])) throw new Exception("Selection lost on refresh.");
                    Console.WriteLine("PASS: selection preserved during refresh.");
                    Cell probe = sheet.Cells["B9"];
                    string formula = probe.FormulaInvariant;
                    CellValue value = probe.Value;
                    try {
                        probe.Value = 1;
                        nativeGrid.RefreshData();
                        var colours = (System.Collections.IDictionary)grid.GetType()
                            .GetField("conditionalForeground", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(grid);
                        if (!colours.Contains(sheet.Cells["E9"].GetReferenceA1()) ||
                            colours.Contains(sheet.Cells["E10"].GetReferenceA1())) {
                            Console.WriteLine("Probe B9=" + probe.DisplayText);
                            foreach (object key in colours.Keys) Console.WriteLine("Conditional address=" + key);
                            throw new Exception("Conditional-format relative references failed.");
                        }
                    } finally {
                        probe.Value = value;
                        if (!string.IsNullOrEmpty(formula)) probe.FormulaInvariant = formula;
                        nativeGrid.RefreshData();
                    }
                    Console.WriteLine("PASS: workbook error colour follows the correct row after refresh.");
                    using (var host = new Form()) {
                        grid.Parent.Controls.Remove(grid);
                        host.ClientSize = new Size(1500, 960);
                        host.Controls.Add(grid);
                        grid.Dock = DockStyle.Fill;
                        host.CreateControl();
                        grid.CreateControl();
                        nativeGrid.ForceInitialize();
                        host.PerformLayout();
                        if (view.Columns[0].Caption != sheet.Cells["A6"].DisplayText ||
                            view.Columns[0].Width < 200)
                            throw new Exception("Captions or best-fit widths lost on first display.");
                        using (var bitmap = new Bitmap(1500, 960)) {
                            grid.DrawToBitmap(bitmap, new Rectangle(0, 0, 1500, 960));
                            string render = Path.Combine(Application.StartupPath, "check-sheet.png");
                            bitmap.Save(render);
                            Console.WriteLine("RENDER: " + render);
                        }
                        view.FocusedRowHandle = 15;
                        view.FocusedColumn = view.Columns[2];
                        view.ShowEditor();
                        if (view.ActiveEditor == null || view.ActiveEditor.GetType().Name != "ComboBoxEdit")
                            throw new Exception("The actual override cell did not open a native dropdown.");
                        view.ActiveEditor.EditValue = alternate;
                        // Programmatic EditValue assignment does not mark a native
                        // editor dirty; emulate the flag set by a user's selection.
                        view.ActiveEditor.IsModified = true;
                        view.PostEditor();
                        view.CloseEditor();
                        if (input.Value.TextValue != alternate)
                            throw new Exception("In-place editor did not commit to the workbook.");
                        view.ShowEditor();
                        model.ChangeManager.Undo();
                        if (input.Value.TextValue != original || view.ActiveEditor != null ||
                            view.GetRowCellDisplayText(15, view.Columns[2]) != original)
                            throw new Exception("Undo left a stale in-place editor.");
                        view.FocusedRowHandle = 2;
                        view.FocusedColumn = view.Columns[2];
                        view.ShowEditor();
                        if (view.ActiveEditor != null) throw new Exception("Ordinary Override row accepted editing.");
                        Console.WriteLine("PASS: native in-place dropdown commits; undo closes stale editor; other rows reject editing.");
                    }
                }
            }
        }
        // No Save/SaveAs or interactive CloseModel. Process exit releases the
        // in-memory model; the input file is never written.
    }
    private static Control FindGrid(Control parent) {
        if (parent.GetType().Name == "ReadOnlyMappedTableGrid") return parent;
        foreach (Control child in parent.Controls) {
            Control result = FindGrid(child);
            if (result != null) return result;
        }
        return null;
    }
}
