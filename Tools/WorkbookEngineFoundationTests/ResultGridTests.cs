using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Abovo.WorkbookEngines;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;

static class ResultGridTests
{
    sealed class Backend : IWorkbookCalculationBackend
    {
        public string Name => "Controlled authority";
        public string Version => "1";
        public double Number = 42;
        public void OpenReadOnly(string path, WorkbookEngineOptions options) { }
        public void Calculate(WorkbookCalculationKind kind) { }
        public WorkbookValueBlock Read(WorkbookReadArea area) => new WorkbookValueBlock(area,
            new object[,] { { Number, "Loan", true }, { -123.456d, null, new WorkbookCellError("#DIV/0!") } });
        public void Dispose() { }
    }
    static int assertions;
    static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
        Console.WriteLine("GRID PASS " + (++assertions) + " " + message);
    }
    internal static Task Run(string path)
    {
        var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() => {
            try { RunOnOwner(path); completion.SetResult(true); }
            catch (Exception error) { completion.SetException(error); }
        });
        thread.IsBackground = true;
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return completion.Task;
    }
    internal static Task NativeResult(WorkbookCalculationSession session, WorkbookCalculationResult result)
    {
        var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() => {
            try
            {
                var timer = Stopwatch.StartNew(); int cells = 0;
                foreach (var block in result.Blocks)
                {
                    var source = new WorkbookResultGridSource(session, block.Area);
                    if (!source.TryPublish(result)) throw new Exception("Native result became stale before display");
                    using (var form = new Form())
                    using (var grid = new GridControl())
                    using (var view = new GridView(grid))
                    {
                        grid.MainView = view; form.Controls.Add(grid); grid.DataSource = source;
                        view.OptionsBehavior.Editable = false;
                        view.OptionsSelection.MultiSelect = true;
                        view.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect;
                        view.OptionsClipboard.AllowCopy = DevExpress.Utils.DefaultBoolean.True;
                        form.CreateControl(); grid.CreateControl(); grid.ForceInitialize(); view.PopulateColumns();
                        if (view.DataRowCount != block.Area.Rows || view.Columns.Count != block.Area.Columns) throw new Exception("Native result grid geometry differs");
                        foreach (int row in new[] { 0, block.Area.Rows / 2, block.Area.Rows - 1 })
                        foreach (int column in new[] { 0, block.Area.Columns / 2, block.Area.Columns - 1 })
                        {
                            if (!Equals(view.GetRowCellValue(row, "Column" + column), block.ValueAt(row, column)))
                                throw new Exception("Native grid differs from its engine's accepted values");
                            cells++;
                        }
                    }
                }
                Console.WriteLine("GRID_NATIVE engine=" + result.Engine + " rectangles=" + result.Blocks.Count + " sampledCells=" + cells + " bindMs=" + timer.ElapsedMilliseconds);
                completion.SetResult(true);
            }
            catch (Exception error) { completion.SetException(error); }
        });
        thread.IsBackground = true; thread.SetApartmentState(ApartmentState.STA); thread.Start();
        return completion.Task;
    }
    static void RunOnOwner(string path)
    {
        var backend = new Backend();
        var session = WorkbookCalculationSession.OpenAsync(path, new WorkbookEngineOptions(), default(CancellationToken), _ => backend).GetAwaiter().GetResult();
        var area = new WorkbookReadArea("Funding Assumptions", 80, 6, 2, 3);
        try
        {
            var source = new WorkbookResultGridSource(session, area);
            var properties = ((ITypedList)source).GetItemProperties(null);
            Check(source.Count == 2 && properties.Count == 3 && !source.AllowEdit && !source.AllowNew && !source.AllowRemove,
                "fixed geometry and read-only binding contract");
            Check(source[0].WorksheetRow == 80 && source[0].GetValue(0) is WorkbookCellError && !source.HasCurrentValues,
                "unpublished values unavailable with exact source-row mapping");
            var result = session.CalculateAndReadAsync(0, WorkbookCalculationKind.Full, new[] { area }).GetAwaiter().GetResult();
            int resets = 0;
            source.ListChanged += (s, e) => { if (e.ListChangedType == ListChangedType.Reset) resets++; };
            Check(source.TryPublish(result) && resets == 1, "one batched notification publishes a complete result");
            Check((double)properties[0].GetValue(source[0]) == 42 && (string)properties[1].GetValue(source[0]) == "Loan" &&
                (bool)properties[2].GetValue(source[0]) && source[1].GetValue(1) == null && source[1].GetValue(2) is WorkbookCellError,
                "numbers, text, booleans, blanks and errors retain distinct types");
            bool refused = false;
            try { properties[0].SetValue(source[0], 8d); } catch (NotSupportedException) { refused = true; }
            Check(refused && (double)source[0].GetValue(0) == 42, "property setter cannot alter an engine result");
            refused = false;
            try { ((IList)source).RemoveAt(0); } catch (NotSupportedException) { refused = true; }
            Check(refused && source.Count == 2, "grid row deletion cannot alter the workbook");

            using (var form = new Form())
            using (var grid = new GridControl())
            using (var view = new GridView(grid))
            {
                grid.MainView = view; grid.Dock = DockStyle.Fill; form.Controls.Add(grid);
                view.OptionsBehavior.Editable = false;
                view.OptionsSelection.MultiSelect = true;
                view.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect;
                view.OptionsClipboard.AllowCopy = DevExpress.Utils.DefaultBoolean.True;
                grid.DataSource = source;
                form.CreateControl(); grid.CreateControl(); grid.ForceInitialize(); view.PopulateColumns();
                Check(view.DataRowCount == 2 && view.Columns.Count == 3 && Convert.ToDouble(view.GetRowCellValue(0, "Column0")) == 42,
                    "native DevExpress GridView binds the detached engine result");
                Check(view.GetRowCellValue(1, "Column2") is WorkbookCellError, "native grid retains typed spreadsheet error");
                session.InvalidateResults();
                Check(!source.HasCurrentValues && source[0].GetValue(0) is WorkbookCellError && !source.TryPublish(result),
                    "revision change makes stale data unavailable even before the UI refresh runs");
                source.InvalidateDisplay();
                grid.RefreshDataSource();
                Check(view.GetRowCellValue(0, "Column0") is WorkbookCellError,
                    "controller invalidation removes the previous native-grid value");
                backend.Number = 84;
                result = session.CalculateAndReadAsync(1, WorkbookCalculationKind.Full, new[] { area }).GetAwaiter().GetResult();
                Check(source.TryPublish(result), "next current revision accepted");
                grid.RefreshDataSource();
                Check(Convert.ToDouble(view.GetRowCellValue(0, "Column0")) == 84 && source[0].WorksheetRow == 80,
                    "existing native grid refreshes without replacing its source or row mapping");
                view.SelectCells(0, view.Columns[0], 1, view.Columns[1]);
                Check(view.GetSelectedCells().Length == 4, "native cell multiselect remains available");
            }
            var access = Task.Run(() => { try { source[0].GetValue(0); return false; } catch (InvalidOperationException) { return true; } }).GetAwaiter().GetResult();
            Check(access, "cross-thread grid access is rejected");
            var other = new WorkbookResultGridSource(session, new WorkbookReadArea("Wrong", 80, 6, 2, 3));
            refused = false;
            try { other.TryPublish(result); } catch (ArgumentException) { refused = true; }
            Check(refused && !other.HasCurrentValues, "wrong worksheet cannot receive a result");
            session.CloseAsync().GetAwaiter().GetResult();
            Check(!source.HasCurrentValues && source[0].GetValue(0) is WorkbookCellError, "closed session cannot expose cached values as current");
        }
        finally { session.CloseAsync().GetAwaiter().GetResult(); }
        Console.WriteLine("GRID ASSERTIONS=" + assertions);
    }
}
