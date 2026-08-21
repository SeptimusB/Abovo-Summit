Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Globalization
Imports System.Windows.Forms
Imports Abovo
Imports Abovo.FileManager
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.XtraVerticalGrid.Events
Imports DevExpress.XtraVerticalGrid.Rows

''' <summary>
''' Funding-Assumptions-style presentation of FFR Inputs Adj Stmt. This local
''' MergeDownAndPivot projection makes worksheet columns VGrid records and
''' worksheet rows fields; the workbook remains the source of truth.
''' </summary>
Public Class FFRInputsAdjStmtVGridView
    Inherits XtraUserControl

    Private Const SheetName As String = "FFR Inputs Adj Stmt"
    Private ReadOnly ModelID As Integer
    Private ReadOnly Workbook As IWorkbook
    Private ReadOnly ChangeManager As ModelChangeManager
    Private ReadOnly Host As New XtraScrollableControl()
    Private ReadOnly Workspace As New Panel()
    Private ReadOnly ActualGrid As New VGridControl()
    Private ReadOnly LoansGrid As New VGridControl()
    Private ReadOnly ActualSource As New PivotSource()
    Private ReadOnly LoansSource As New PivotSource()
    Private Loading As Boolean

    Public Event WorkbookCellChanged As EventHandler

    Public Sub New(ByVal setModelID As Integer)
        ModelID = setModelID
        Workbook = FileManager.GetWorkBook(ModelID)
        ChangeManager = ExcelModels(ModelID).ChangeManager
        Dock = DockStyle.Fill
        BuildSurface()
        RefreshFromWorkbook()
    End Sub

    Public ReadOnly Property WorksheetName As String
        Get
            Return SheetName
        End Get
    End Property

    Private Sub BuildSurface()
        Host.Dock = DockStyle.Fill
        Host.AutoScroll = True
        AddHandler Host.Resize, AddressOf HostResize
        Controls.Add(Host)
        Workspace.Size = New Size(1080, 1040)
        Host.Controls.Add(Workspace)
        ActualGrid.Name = "FFRActualStockGrid"
        LoansGrid.Name = "FFRLoansGrid"
        AddSection("FFR Inputs Adjustment Statement", "FFR Actual Stock Inputs", ActualGrid, 20, 18, 1080, 620)
        AddSection("", "Statement of Cash Flow - Movements in Loans", LoansGrid, 20, 662, 1080, 360)
        ConfigureGrid(ActualGrid)
        ConfigureGrid(LoansGrid)
    End Sub

    Private Sub AddSection(ByVal documentTitle As String, ByVal caption As String, ByVal grid As VGridControl, ByVal left As Integer, ByVal top As Integer, ByVal width As Integer, ByVal height As Integer)
        If documentTitle.Length > 0 Then
            Dim title As New LabelControl With {.Text = documentTitle, .Location = New Point(left, top), .AutoSizeMode = LabelAutoSizeMode.None, .Size = New Size(width, 26)}
            title.Appearance.Font = New Font(Font, FontStyle.Bold)
            title.Appearance.ForeColor = Color.FromArgb(32, 58, 89)
            Workspace.Controls.Add(title)
            top += 28
        End If
        Dim label As New LabelControl With {.Name = grid.Name & "_Caption", .Text = caption, .Location = New Point(left, top), .AutoSizeMode = LabelAutoSizeMode.None, .Size = New Size(width, 22)}
        label.Appearance.Font = New Font(Font, FontStyle.Bold)
        label.Appearance.ForeColor = Color.FromArgb(0, 90, 180)
        Workspace.Controls.Add(label)
        grid.Location = New Point(left, top + 24)
        grid.Size = New Size(width, height - 24)
        Workspace.Controls.Add(grid)
    End Sub

    Private Sub ConfigureGrid(ByVal grid As VGridControl)
        grid.OptionsBehavior.Editable = True
        grid.OptionsBehavior.ResizeRowHeaders = False
        grid.OptionsView.ShowButtons = False
        grid.ScrollVisibility = ScrollVisibility.Horizontal
        grid.RowHeaderWidth = 390
        grid.RecordWidth = 78
        AddHandler grid.CustomDrawRowValueCell, AddressOf GridCustomDrawRowValueCell
        AddHandler grid.ShowingEditor, AddressOf GridShowingEditor
        AddHandler grid.CellValueChanged, AddressOf GridCellValueChanged
    End Sub

    Public Sub RefreshFromWorkbook()
        If Workbook Is Nothing OrElse Not Workbook.Worksheets.Contains(SheetName) Then Return
        Loading = True
        Try
            BuildPivot(ActualSource, ActualGrid, 4, 49, 2, 9, "Actual stock")
            BuildPivot(LoansSource, LoansGrid, 52, 69, 2, 31, "Loan movements")
            ResizeFirstGridToContents()
        Finally
            Loading = False
        End Try
    End Sub

    Private Sub ResizeFirstGridToContents()
        'A VGrid exposes worksheet rows as fields.  Keep its entire first
        'block visible and let the page own vertical scrolling, just as the
        'Funding Assumptions presenter does for tiled datasets.
        Dim categoryCount As Integer = ActualSource.Bands.Values.Distinct().Count()
        Dim desiredHeight As Integer = 88 + (ActualSource.Rows.Count * 21) + (categoryCount * 24)
        ActualGrid.Height = Math.Max(300, desiredHeight)

        Dim loansCaption As Control = Workspace.Controls(LoansGrid.Name & "_Caption")
        If loansCaption IsNot Nothing Then
            loansCaption.Top = ActualGrid.Bottom + 22
            LoansGrid.Top = loansCaption.Bottom + 2
        End If
        Workspace.Height = LoansGrid.Bottom + 24
        Host.AutoScrollMinSize = New Size(Workspace.Width + 40, Workspace.Height + 20)
        CentreWorkspace()
    End Sub

    Private Sub HostResize(ByVal sender As Object, ByVal e As EventArgs)
        CentreWorkspace()
    End Sub

    Private Sub CentreWorkspace()
        If Host Is Nothing OrElse Workspace Is Nothing Then Return
        Workspace.Left = Math.Max(20, (Host.ClientSize.Width - Workspace.Width) \ 2)
        Workspace.Top = 0
    End Sub

    Private Sub BuildPivot(ByVal source As PivotSource, ByVal grid As VGridControl, ByVal firstRow As Integer, ByVal lastRow As Integer, ByVal firstColumn As Integer, ByVal lastColumn As Integer, ByVal defaultBand As String)
        Dim ws As Worksheet = Workbook.Worksheets(SheetName)
        source.Reset()
        Dim table As New DataTable()
        table.Columns.Add("__SourceColumn", GetType(Integer))
        For col As Integer = firstColumn To lastColumn
            Dim record As DataRow = table.NewRow()
            record("__SourceColumn") = col
            table.Rows.Add(record)
        Next
        For row As Integer = firstRow To lastRow
            Dim heading As String = ws.Cells(row, 1).DisplayText.Trim()
            If heading.Length = 0 Then heading = ws.Cells(row, 0).DisplayText.Trim()
            Dim hasContent As Boolean = heading.Length > 0
            For col As Integer = firstColumn To lastColumn
                If ws.Cells(row, col).DisplayText.Length > 0 Then hasContent = True
            Next
            If Not hasContent Then Continue For
            Dim field As String = "Row_" & row.ToString(CultureInfo.InvariantCulture)
            table.Columns.Add(field, GetType(Object))
            source.Rows.Add(field, row)
            source.Headings.Add(field, If(heading.Length = 0, " ", heading))
            source.Bands.Add(field, BandForRow(row, defaultBand))
            For recordIndex As Integer = 0 To table.Rows.Count - 1
                Dim col As Integer = CInt(table.Rows(recordIndex)("__SourceColumn"))
                table.Rows(recordIndex)(field) = CellValue(ws.Cells(row, col))
            Next
        Next
        grid.BeginUpdate()
        Try
            grid.DataSource = table
            grid.Rows.Clear()
            Dim categories As New Dictionary(Of String, CategoryRow)(StringComparer.Ordinal)
            For Each item In source.Rows
                Dim field As String = item.Key
                Dim band As String = source.Bands(field)
                Dim category As CategoryRow = Nothing
                If Not categories.TryGetValue(band, category) Then
                    category = New CategoryRow("Category_" & categories.Count.ToString())
                    category.Properties.Caption = band
                    category.Height = 24
                    categories.Add(band, category)
                    grid.Rows.Add(category)
                End If
                Dim editorRow As New EditorRow(field)
                editorRow.Properties.FieldName = field
                editorRow.Properties.Caption = source.Headings(field)
                editorRow.Properties.ReadOnly = False
                editorRow.Height = 21
                category.ChildRows.Add(editorRow)
            Next
            grid.ForceInitialize()
        Finally
            grid.EndUpdate()
        End Try
    End Sub

    Private Function BandForRow(ByVal row As Integer, ByVal fallback As String) As String
        If row <= 18 Then Return "FFR Actual Stock Inputs"
        If row <= 22 Then Return "FFR Actual Stock Inputs"
        If row <= 32 Then Return "Housing Units Owned & Managed"
        If row <= 41 Then Return "BP opening units"
        If row <= 50 Then Return "Difference to closing actual units"
        If row <= 61 Then Return "Forecast movements"
        If row <= 66 Then Return "Loan adjustments"
        Return fallback
    End Function

    Private Sub GridShowingEditor(ByVal sender As Object, ByVal e As CancelEventArgs)
        Dim grid As VGridControl = TryCast(sender, VGridControl)
        Dim source As PivotSource = SourceFor(grid)
        If grid Is Nothing OrElse source Is Nothing OrElse grid.FocusedRow Is Nothing Then e.Cancel = True : Return
        Dim row As Integer
        If Not source.Rows.TryGetValue(grid.FocusedRow.Properties.FieldName, row) Then e.Cancel = True : Return
        e.Cancel = Not IsEditable(Workbook.Worksheets(SheetName).Cells(row, SourceColumn(grid)))
    End Sub

    Private Sub GridCellValueChanged(ByVal sender As Object, ByVal e As CellValueChangedEventArgs)
        If Loading Then Return
        Dim grid As VGridControl = TryCast(sender, VGridControl)
        Dim source As PivotSource = SourceFor(grid)
        If grid Is Nothing OrElse source Is Nothing OrElse e.Row Is Nothing Then Return
        Dim row As Integer
        If Not source.Rows.TryGetValue(e.Row.Properties.FieldName, row) Then Return
        Dim cell As Cell = Workbook.Worksheets(SheetName).Cells(row, SourceColumn(grid))
        If Not IsEditable(cell) Then RefreshFromWorkbook() : Return
        Dim change As New DataChangeEvent With {.ModelID = ModelID, .Description = "FFR input updated", .WSName = SheetName, .CellAddress = cell.GetReferenceA1(), .OriginalValue = CellValue(cell), .ChangedValue = e.Value, .DataFormat = "S", .TimeStamp = Now(), .UserName = Environment.UserName}
        'ProcessChange performs the authoritative workbook calculation.  Like
        'DataInterfaceTemplate.UpdateCalcs/RefreshData, rebuild the complete
        'pivot afterwards so calculated rows (including Closing Actual Units)
        'are read back from the calculated workbook rather than retaining a
        'stale VGrid value.
        If ChangeManager.ProcessChange(change).BError Then
            RefreshFromWorkbook()
            Return
        End If

        'CellValueChanged is raised while the VGrid is still committing its
        'bound DataTable value.  Defer the DIT-equivalent RefreshData pass to
        'the message queue so the calculated workbook results cannot be
        'overwritten by that final local commit.
        BeginInvoke(New MethodInvoker(AddressOf RefreshFromWorkbook))
        RaiseEvent WorkbookCellChanged(Me, EventArgs.Empty)
    End Sub

    Private Sub GridCustomDrawRowValueCell(ByVal sender As Object, ByVal e As CustomDrawRowValueCellEventArgs)
        Dim grid As VGridControl = TryCast(sender, VGridControl)
        Dim source As PivotSource = SourceFor(grid)
        If grid Is Nothing OrElse source Is Nothing OrElse e.Row Is Nothing Then Return
        Dim row As Integer
        If Not source.Rows.TryGetValue(e.Row.Properties.FieldName, row) Then Return
        Dim cell As Cell = Workbook.Worksheets(SheetName).Cells(row, SourceColumn(grid, e.RecordIndex))
        e.CellText = cell.DisplayText
        e.Appearance.BackColor = If(cell.FillColor.IsEmpty, Color.White, cell.FillColor)
        e.Appearance.ForeColor = DisplayForeground(cell)
        e.Appearance.Font = New Font(Font, If(cell.Font.Bold, FontStyle.Bold, FontStyle.Regular))
    End Sub

    Private Function SourceFor(ByVal grid As VGridControl) As PivotSource
        If Object.ReferenceEquals(grid, ActualGrid) Then Return ActualSource
        If Object.ReferenceEquals(grid, LoansGrid) Then Return LoansSource
        Return Nothing
    End Function

    Private Function SourceColumn(ByVal grid As VGridControl, Optional ByVal recordIndex As Integer = -1) As Integer
        If recordIndex < 0 Then recordIndex = grid.FocusedRecord
        Dim table As DataTable = TryCast(grid.DataSource, DataTable)
        If table Is Nothing OrElse recordIndex < 0 OrElse recordIndex >= table.Rows.Count Then Return 0
        Return CInt(table.Rows(recordIndex)("__SourceColumn"))
    End Function

    Private Shared Function CellValue(ByVal cell As Cell) As Object
        If cell Is Nothing OrElse cell.Value.IsEmpty Then Return String.Empty
        If cell.Value.IsNumeric Then Return cell.Value.NumericValue
        Return cell.DisplayText
    End Function

    Private Shared Function DisplayForeground(ByVal cell As Cell) As Color
        If cell IsNot Nothing AndAlso cell.DisplayText.Trim().StartsWith("(", StringComparison.Ordinal) Then Return Color.Red
        Return If(cell Is Nothing OrElse cell.Font.Color.IsEmpty, Color.FromArgb(32, 58, 89), cell.Font.Color)
    End Function

    Private Shared Function IsEditable(ByVal cell As Cell) As Boolean
        Return cell IsNot Nothing AndAlso Not cell.Protection.Locked AndAlso cell.Fill.PatternType = PatternType.Solid
    End Function

    Private Class PivotSource
        Public ReadOnly Rows As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        Public ReadOnly Headings As New Dictionary(Of String, String)(StringComparer.Ordinal)
        Public ReadOnly Bands As New Dictionary(Of String, String)(StringComparer.Ordinal)
        Public Sub Reset()
            Rows.Clear() : Headings.Clear() : Bands.Clear()
        End Sub
    End Class
End Class
