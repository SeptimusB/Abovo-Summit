using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using DevExpress.Spreadsheet;

// Real services and native controls, with transparent/off-taskbar windows.
// All edits are in memory. No Save/SaveAs and no source-package changes.
public static class StandaloneHistoryFixture
{
    [STAThread]
    public static int Main(string[] args) {
        try { Run(args); return 0; }
        catch (Exception e) { Console.Error.WriteLine(e); return 1; }
    }
    static void Run(string[] args) {
        AppDomain.CurrentDomain.AssemblyResolve += (sender, e) => {
            string path = Path.Combine(args[0], new AssemblyName(e.Name).Name + ".dll");
            return File.Exists(path) ? Assembly.LoadFrom(path) : null;
        };
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
        Application.EnableVisualStyles();
        Assembly app = Assembly.LoadFrom(Path.Combine(args[0], "Abovo-summit.exe"));
        app.GetType("Abovo.AbovoAppCls").GetMethod("Initialise").Invoke(null, null);
        Type files = app.GetType("Abovo.FileManager");
        files.GetMethod("Initialise").Invoke(null, new object[] { null });
        MethodInfo open = files.GetMethod("OpenModel");
        dynamic opened = open.Invoke(null, new object[] { args[1], new FileInfo(args[1]),
            Enum.ToObject(open.GetParameters()[2].ParameterType, 0) });
        Check(!opened.BError, "Model open");
        dynamic model = ((Array)files.GetField("ExcelModels").GetValue(null)).GetValue(0);
        IWorkbook book = (IWorkbook)model.WB;
        dynamic changes = model.ChangeManager;
        using (Control instance = (Control)Activator.CreateInstance(app.GetType("FileInstanceInterface"), new object[] { 0 }))
        using (Form ffr = (Form)Activator.CreateInstance(app.GetType("FFRForm"), new object[] { 0 }))
        using (Form stress = (Form)Activator.CreateInstance(app.GetType("StressTest"), new object[] { 0 })) {
            ffr.Opacity = 0; ffr.ShowInTaskbar = false;
            stress.Opacity = 0; stress.ShowInTaskbar = false;
            instance.GetType().GetField("FFRer").SetValue(instance, ffr);
            instance.GetType().GetField("StressTester").SetValue(instance, stress);
            Set(instance, "FFRInit", true); Set(instance, "STInit", true);
            Invoke(instance, "ShowFFRInterface");
            Invoke(instance, "ShowStressTestInterface");
            Application.DoEvents();
            Check(ffr.Visible && stress.Visible && !ffr.Modal && !stress.Modal && ffr.Enabled, "Both windows modeless");
            IList visits = (IList)model.InterfaceHistory.SnapshotItems();
            Check(visits.Count == 2, "Both standalone interfaces recorded once");
            Check(ffr.Controls.Find("HistoryButton", true).Length == 1, "FFR History button");
            dynamic navigator = Field(stress, "WindowsUIButtonPanelStressNavigator");
            bool historyButton = false;
            foreach (dynamic button in navigator.Buttons)
                if (Convert.ToString(button.Tag) == "History") historyButton = true;
            Check(historyButton, "Stress History button");
            Console.WriteLine("PASS: FFR and Stress Test open modeless; both have history access and unique Interface History entries.");

            dynamic tabs = Field(ffr, "FFRTabs");
            tabs.SelectedTabPageIndex = 1; // Front Sheet
            Application.DoEvents();
            Control front = ffr.Controls.Find("FFR_Front_Sheet", true)[0].Controls[0];
            Cell ffrCell = book.Worksheets["Front Sheet"].Cells["B5"];
            string originalFFR = ffrCell.DisplayText;
            Invoke(front, "CommitWorkbookCell", ffrCell, "History trial", "FFR history fixture");
            Check(ffrCell.DisplayText == "History trial", "FFR edit journalled");
            dynamic undo = changes.Undo();
            Console.WriteLine("FFR undo: error=" + undo.BError + ", message=" + undo.StrResponseMessage +
                ", original='" + originalFFR + "', workbook='" + ffrCell.DisplayText +
                "', editor='" + EditText(front, "RPNumberEdit") + "'");
            Check(!undo.BError && ffrCell.DisplayText == originalFFR &&
                EditText(front, "RPNumberEdit") == originalFFR, "FFR visible undo refresh");
            dynamic redo = changes.Redo();
            Check(!redo.BError && EditText(front, "RPNumberEdit") == "History trial", "FFR redo refresh");
            changes.Undo();
            Console.WriteLine("PASS: FFR cell edit, undo/redo and visible editor values.");

            tabs.SelectedTabPageIndex = 7; // Key Definitions, editable native VGrid
            Application.DoEvents();
            Control keyView = ffr.Controls.Find("FFR_FFR_Key_Defn", true)[0].Controls[0];
            dynamic keyGrid = Field(keyView, "Grid");
            var rows = (System.Collections.Generic.Dictionary<string, int>)Field(keyView, "SourceRows");
            var target = rows.First(pair => !book.Worksheets["FFR Key Defn"].Cells[pair.Value, 2].Protection.Locked &&
                string.IsNullOrEmpty(book.Worksheets["FFR Key Defn"].Cells[pair.Value, 2].FormulaInvariant));
            Cell keyCell = book.Worksheets["FFR Key Defn"].Cells[target.Value, 2];
            string originalKey = keyCell.DisplayText;
            dynamic editorRow = FindVGridRow(keyGrid.Rows, target.Key);
            Check(editorRow != null, "FFR key row exists");
            object originalGridValue = ((DataTable)keyGrid.DataSource).Rows[0][target.Key];
            keyGrid.SetCellValue(editorRow, 0, "History key trial");
            Check(keyCell.DisplayText == "History key trial", "FFR native VGrid event journalled");
            keyGrid.FocusedRow = FindVGridRow(keyGrid.Rows, target.Key);
            keyGrid.FocusedRecord = 0;
            keyGrid.ShowEditor();
            Check(keyGrid.ActiveEditor != null, "FFR native VGrid editor opened");
            undo = changes.Undo();
            Check(!undo.BError && keyCell.DisplayText == originalKey && keyGrid.ActiveEditor == null, "FFR VGrid undo closes stale editor");
            Check(object.Equals(((DataTable)keyGrid.DataSource).Rows[0][target.Key], originalGridValue), "FFR VGrid snapshot refreshed");
            Console.WriteLine("PASS: FFR Key Definitions VGrid edit/undo and stale editor closure.");
            tabs.SelectedTabPageIndex = 1;

            Cell name = book.Worksheets["Live Multivariable Planner"].Cells["B5"];
            string originalName = name.DisplayText;
            int beforeCount = ((DataTable)changes.GetHistoryTable()).Rows.Count;
            Check((bool)Invoke(stress, "CommitLiveMultivariableName", "History trial"), "Stress edit accepted");
            Check(name.DisplayText == "History trial", "Stress edit workbook value");
            undo = changes.Undo();
            Check(!undo.BError && name.DisplayText == originalName &&
                EditText(stress, "TextEditMultivariableName") == originalName, "Stress visible undo refresh");
            redo = changes.Redo();
            Check(!redo.BError && name.DisplayText == "History trial" &&
                EditText(stress, "TextEditMultivariableName") == "History trial", "Stress redo refresh");
            Check(((DataTable)changes.GetHistoryTable()).Rows.Count == beforeCount + 1, "Refresh did not post phantom history");
            stress.Close();
            Check(!stress.IsDisposed && !stress.Visible, "Stress close retains state");
            changes.Undo();
            Invoke(instance, "ShowStressTestInterface");
            Check(EditText(stress, "TextEditMultivariableName") == originalName, "Hidden Stress refresh on reopen");
            ffr.Close();
            Check(!ffr.IsDisposed && !ffr.Visible, "FFR close retains state");
            Invoke(instance, "ShowFFRInterface");
            Check((int)tabs.SelectedTabPageIndex == 1, "FFR selected tab preserved");
            Check(((IList)model.InterfaceHistory.SnapshotItems()).Count == 2, "History deduplicates reopening");
            Console.WriteLine("PASS: Stress undo/redo, no phantom posts, hidden/reopen refresh and retained tabs; history deduplicated.");
            object stressBinding = Field(stress, "HistoryBinding");
            object ffrBinding = Field(ffr, "HistoryBinding");
            Invoke(instance, "CloseStandaloneInterfaces");
            Check(ffr.IsDisposed && stress.IsDisposed, "Model close disposes standalone windows");
            Check((bool)Field(stressBinding, "DisposedBinding") && (bool)Field(ffrBinding, "DisposedBinding"), "Keyboard/history subscriptions detached");
            Console.WriteLine("PASS: model-close cleanup detaches both history bindings and disposes both forms.");
        }
    }
    static string EditText(object owner, string name) { return Convert.ToString((object)((dynamic)Field(owner, name)).EditValue); }
    static object FindVGridRow(IEnumerable rows, string field) {
        foreach (dynamic row in rows) {
            if ((string)row.Properties.FieldName == field) return row;
            object match = FindVGridRow(row.ChildRows, field);
            if (match != null) return match;
        }
        return null;
    }
    static object Field(object target, string name) {
        for (Type type = target.GetType(); type != null; type = type.BaseType) {
            var field = type.GetField(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            if (field != null) return field.GetValue(target);
            var property = type.GetProperty(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            if (property != null) return property.GetValue(target, null);
        }
        throw new Exception("Missing field " + name);
    }
    static void Set(object target, string name, object value) {
        target.GetType().GetField(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).SetValue(target, value);
    }
    static object Invoke(object target, string name, params object[] args) {
        return target.GetType().GetMethod(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).Invoke(target, args);
    }
    static void Check(bool value, string description) { if (!value) throw new Exception(description); }
}
