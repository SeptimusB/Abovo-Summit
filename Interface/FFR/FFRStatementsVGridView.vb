Imports System.Data
Imports System.Drawing
Imports System.Globalization
Imports System.Windows.Forms
Imports Abovo
Imports Abovo.FileManager
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraEditors
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.XtraVerticalGrid.Events
Imports DevExpress.XtraVerticalGrid.Rows

'''''' <summary>
'''''' Read-only vertical presentation of the authoritative Statements sheet.
'''''' Column A workbook section headings form the VGrid categories; each
'''''' workbook period is a record and every displayed value/style is read from
'''''' its underlying worksheet cell.
'''''' </summary>
Public Class FFRStatementsVGridView
    Inherits XtraUserControl

    Private ReadOnly SheetName As String
    Private Const FirstPresentationRow As Integer = 7
    Private ReadOnly LastPresentationRow As Integer
    Private ReadOnly DisplayTitle As String
    Private ReadOnly EditableColumnC As Boolean
    Private Const MinimumWorkspaceWidth As Integer = 1080

    Private ReadOnly ModelID As Integer
    Private ReadOnly Workbook As IWorkbook
    Private ReadOnly Host As New Panel()
    Private ReadOnly Workspace As New Panel()
    Private ReadOnly Grid As New VGridControl()
    Private ReadOnly SourceRows As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
    Private Loading As Boolean

    Public Sub New(ByVal setModelID As Integer,
                   Optional ByVal worksheetNameOverride As String = "Statements",
                   Optional ByVal displayTitleOverride As String = "FFR Statements",
                   Optional ByVal lastPresentationRowOverride As Integer = 151,
                   Optional ByVal editableColumnCOverride As Boolean = False)
        ModelID = setModelID
        SheetName = worksheetNameOverride
        DisplayTitle = displayTitleOverride
        LastPresentationRow = lastPresentationRowOverride
        EditableColumnC = editableColumnCOverride
        Workbook = GetWorkBook(ModelID)
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
        Host.AutoScroll = False
        AddHandler Host.Resize, AddressOf HostResize
        Controls.Add(Host)

        Workspace.Width = PreferredWorkspaceWidth
        Host.Controls.Add(Workspace)

        Dim title As New LabelControl With {
            .Text = DisplayTitle,
            .Location = New Point(0, 16),
            .AutoSizeMode = LabelAutoSizeMode.None,
            .Size = New Size(PreferredWorkspaceWidth, 26)}
        title.Appearance.Font = New Font(Font, FontStyle.Bold)
        title.Appearance.ForeColor = Color.FromArgb(32, 58, 89)
        Workspace.Controls.Add(title)

        Grid.Location = New Point(0, 44)
        Grid.Width = PreferredWorkspaceWidth
        Grid.OptionsBehavior.Editable = EditableColumnC
        Grid.OptionsBehavior.ResizeRowHeaders = False
        Grid.OptionsView.ShowButtons = False
        Grid.OptionsSelectionAndFocus.MultiSelect = True
        Grid.OptionsSelectionAndFocus.MultiSelectMode = MultiSelectMode.CellSelect
        Grid.ScrollVisibility = ScrollVisibility.Vertical
        Grid.RowHeaderWidth = If(EditableColumnC, 720, 410)
        Grid.RecordWidth = If(EditableColumnC, 360, 78)
        AddHandler Grid.CustomDrawRowValueCell, AddressOf GridCustomDrawRowValueCell
        AddHandler Grid.CustomDrawRowHeaderCell, AddressOf GridCustomDrawRowHeaderCell
        AddHandler Grid.KeyDown, AddressOf GridKeyDown
        AddHandler Grid.ShowingEditor, AddressOf GridShowingEditor
        AddHandler Grid.CellValueChanged, AddressOf GridCellValueChanged
        Workspace.Controls.Add(Grid)
        ResizeSurface()
    End Sub

    Public Sub RefreshFromWorkbook()
        If Workbook Is Nothing OrElse Not Workbook.Worksheets.Contains(SheetName) Then Return
        Loading = True
        Grid.BeginUpdate()
        Try
            BuildPivot()
        Finally
            Grid.EndUpdate()
            Loading = False
        End Try
    End Sub

    Private Sub BuildPivot()
        Dim ws As Worksheet = Workbook.Worksheets(SheetName)
        Dim used As CellRange = ws.GetUsedRange()
        Dim table As New DataTable()
        table.Columns.Add("__SourceColumn", GetType(Integer))
        Dim lastColumn As Integer = If(EditableColumnC, 2, used.RightColumnIndex)
        For col As Integer = 2 To lastColumn
            Dim record As DataRow = table.NewRow()
            record("__SourceColumn") = col
            table.Rows.Add(record)
        Next

        SourceRows.Clear()
        Grid.DataSource = Nothing
        Grid.Rows.Clear()
        For row As Integer = FirstPresentationRow To LastPresentationRow
            If IsCategoryRow(ws, row) Then Continue For
            If ws.Cells(row, 1).DisplayText.Trim().Length = 0 Then Continue For

            Dim field As String = "Row_" & row.ToString(CultureInfo.InvariantCulture)
            table.Columns.Add(field, GetType(Object))
            SourceRows.Add(field, row)
            For recordIndex As Integer = 0 To table.Rows.Count - 1
                Dim col As Integer = CInt(table.Rows(recordIndex)("__SourceColumn"))
                table.Rows(recordIndex)(field) = CellValue(ws.Cells(row, col))
            Next
        Next

        Grid.DataSource = table
        Grid.ForceInitialize()
        Grid.Rows.Clear()
        Dim categories As New Dictionary(Of Integer, CategoryRow)()
        For Each item In SourceRows
            Dim sourceRow As Integer = item.Value
            Dim category As CategoryRow = EnsureCategory(ws, CategoryForRow(ws, sourceRow), categories)
            Dim editorRow As New EditorRow(item.Key)
            editorRow.Properties.FieldName = item.Key
            editorRow.Properties.Caption = EditorCaption(ws, sourceRow)
            editorRow.Properties.ToolTip = GetCellNote(ws.Cells(sourceRow, 1), ws.Cells(sourceRow, 2), ws.Cells(sourceRow, 0))
            editorRow.Properties.ReadOnly = True
            editorRow.Height = 18
            category.ChildRows.Add(editorRow)
        Next
        Grid.ForceInitialize()
        ResizeSurface()
    End Sub

    Private Function CategoryForRow(ByVal ws As Worksheet, ByVal row As Integer) As Integer
        For probe As Integer = row To FirstPresentationRow Step -1
            If IsCategoryRow(ws, probe) Then Return probe
        Next
        Return -1
    End Function

    Private Function EnsureCategory(ByVal ws As Worksheet,
                                    ByVal sourceRow As Integer,
                                    ByVal categories As Dictionary(Of Integer, CategoryRow)) As CategoryRow
        If sourceRow < 0 Then sourceRow = -1
        Dim category As CategoryRow = Nothing
        If categories.TryGetValue(sourceRow, category) Then Return category

        category = New CategoryRow("Category_" & sourceRow.ToString(CultureInfo.InvariantCulture))
        category.Properties.Caption = If(sourceRow < 0, DisplayTitle, ws.Cells(sourceRow, 0).DisplayText.Trim())
        category.Height = 22
        category.Tag = sourceRow
        Grid.Rows.Add(category)
        categories.Add(sourceRow, category)
        Return category
    End Function

    Private Shared Function IsCategoryRow(ByVal ws As Worksheet, ByVal row As Integer) As Boolean
        Dim caption As String = ws.Cells(row, 0).DisplayText.Trim()
        Return caption.Length > 0 AndAlso ws.Cells(row, 1).DisplayText.Trim().Length = 0 AndAlso Not IsNumeric(caption)
    End Function

    Private Shared Function EditorCaption(ByVal ws As Worksheet, ByVal row As Integer) As String
        Dim lineNumber As String = ws.Cells(row, 0).DisplayText.Trim()
        Dim description As String = ws.Cells(row, 1).DisplayText.Trim()
        Return If(lineNumber.Length = 0, description, lineNumber & " - " & description)
    End Function

    Private Sub HostResize(ByVal sender As Object, ByVal e As EventArgs)
        ResizeSurface()
    End Sub

    Private Sub ResizeSurface()
        Dim width As Integer = If(EditableColumnC, Math.Max(620, Host.ClientSize.Width), Math.Max(PreferredWorkspaceWidth, Host.ClientSize.Width))
        Workspace.Width = width
        Workspace.Left = Math.Max(0, (Host.ClientSize.Width - width) \ 2)
        Grid.Width = width
        If EditableColumnC Then
            Grid.RowHeaderWidth = Math.Max(360, CInt(width * 0.6R))
            Grid.RecordWidth = Math.Max(200, width - Grid.RowHeaderWidth - 24)
        End If
        Grid.Height = Math.Max(400, Host.ClientSize.Height - Grid.Top)
        Workspace.Height = Grid.Bottom
    End Sub

    Private ReadOnly Property PreferredWorkspaceWidth As Integer
        Get
            Return If(EditableColumnC, 1280, MinimumWorkspaceWidth)
        End Get
    End Property

    Private Sub GridKeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
        If Not e.Control OrElse e.Alt OrElse e.KeyCode <> Keys.C Then Return
        Grid.CopyToClipboard()
        e.Handled = True
        e.SuppressKeyPress = True
    End Sub

    Private Sub GridShowingEditor(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        If Not EditableColumnC OrElse Grid.FocusedRow Is Nothing Then e.Cancel = True : Return
        Dim sourceRow As Integer
        e.Cancel = Not SourceRows.TryGetValue(Grid.FocusedRow.Properties.FieldName, sourceRow) OrElse
                   Workbook.Worksheets(SheetName).Cells(sourceRow, 2).Protection.Locked
    End Sub

    Private Sub GridCellValueChanged(ByVal sender As Object, ByVal e As CellValueChangedEventArgs)
        If Loading OrElse Not EditableColumnC OrElse e.Row Is Nothing Then Return
        Dim sourceRow As Integer
        If Not SourceRows.TryGetValue(e.Row.Properties.FieldName, sourceRow) Then Return
        Dim cell As Cell = Workbook.Worksheets(SheetName).Cells(sourceRow, 2)
        If cell.Protection.Locked Then RefreshFromWorkbook() : Return
        Dim change As New DataChangeEvent With {.ModelID = ModelID, .Description = "FFR key definition updated", .WSName = SheetName, .CellAddress = cell.GetReferenceA1(), .OriginalValue = CellValue(cell), .ChangedValue = e.Value, .DataFormat = "S", .TimeStamp = Now(), .UserName = Environment.UserName}
        ExcelModels(ModelID).ChangeManager.ProcessChange(change)
        RefreshFromWorkbook()
    End Sub

    Private Sub GridCustomDrawRowValueCell(ByVal sender As Object, ByVal e As CustomDrawRowValueCellEventArgs)
        If e.Row Is Nothing Then Return
        Dim sourceRow As Integer
        If Not SourceRows.TryGetValue(e.Row.Properties.FieldName, sourceRow) Then Return
        Dim cell As Cell = Workbook.Worksheets(SheetName).Cells(sourceRow, SourceColumn(e.RecordIndex))
        e.CellText = cell.DisplayText
        e.Appearance.BackColor = If(cell.FillColor.IsEmpty, Color.White, cell.FillColor)
        e.Appearance.ForeColor = DisplayForeground(cell)
        e.Appearance.Font = New Font(Font, If(cell.Font.Bold, FontStyle.Bold, FontStyle.Regular))
    End Sub

    Private Sub GridCustomDrawRowHeaderCell(ByVal sender As Object, ByVal e As CustomDrawRowHeaderCellEventArgs)
        Dim category As CategoryRow = TryCast(e.Row, CategoryRow)
        If category IsNot Nothing Then
            Dim categoryRow As Integer = CInt(category.Tag)
            If categoryRow < 0 Then Return
            Dim cell As Cell = Workbook.Worksheets(SheetName).Cells(categoryRow, 0)
            e.Appearance.ForeColor = If(cell.Font.Color.IsEmpty, Color.FromArgb(0, 85, 170), cell.Font.Color)
            e.Appearance.Font = New Font(Font, If(cell.Font.Bold, FontStyle.Bold, FontStyle.Regular))
            e.DefaultDraw()
            e.Handled = True
            Return
        End If

        Dim sourceRow As Integer
        If Not SourceRows.TryGetValue(e.Row.Properties.FieldName, sourceRow) Then Return
        Dim sourceCell As Cell = Workbook.Worksheets(SheetName).Cells(sourceRow, 1)
        e.Appearance.ForeColor = If(sourceCell.Font.Color.IsEmpty, Color.FromArgb(32, 58, 89), sourceCell.Font.Color)
        e.Appearance.Font = New Font(Font, If(sourceCell.Font.Bold, FontStyle.Bold, FontStyle.Regular))
        e.DefaultDraw()
        e.Handled = True
    End Sub

    Private Function SourceColumn(ByVal recordIndex As Integer) As Integer
        Dim table As DataTable = TryCast(Grid.DataSource, DataTable)
        If table Is Nothing OrElse recordIndex < 0 OrElse recordIndex >= table.Rows.Count Then Return 0
        Return CInt(table.Rows(recordIndex)("__SourceColumn"))
    End Function

    Private Shared Function CellValue(ByVal cell As Cell) As Object
        If cell Is Nothing OrElse cell.Value.IsEmpty Then Return String.Empty
        If cell.Value.IsNumeric Then Return cell.Value.NumericValue
        If cell.Value.IsBoolean Then Return cell.Value.BooleanValue
        If cell.Value.IsDateTime Then Return cell.Value.DateTimeValue
        Return cell.DisplayText
    End Function

    Private Shared Function GetCellNote(ParamArray cells() As Cell) As String
        For Each sourceCell As Cell In cells
            Dim comments = sourceCell.Worksheet.Comments.GetComments(sourceCell)
            If comments Is Nothing OrElse comments.Count = 0 Then Continue For
            Dim text As String = String.Join(Environment.NewLine, comments.Select(Function(comment) comment.Text.Trim()).Where(Function(value) value.Length > 0))
            If text.Length > 0 Then Return text
        Next
        Return String.Empty
    End Function

    Private Shared Function DisplayForeground(ByVal cell As Cell) As Color
        If cell IsNot Nothing AndAlso cell.DisplayText.Trim().StartsWith("(", StringComparison.Ordinal) Then Return Color.Red
        Return If(cell Is Nothing OrElse cell.Font.Color.IsEmpty, Color.FromArgb(32, 58, 89), cell.Font.Color)
    End Function
End Class

Public Class FFRAssumptionsTenureVGridView
    Inherits FFRStatementsVGridView

    Public Sub New(ByVal setModelID As Integer)
        MyBase.New(setModelID,
                   "Assumptions & tenure inputs",
                   "FFR Assumptions & Tenure Inputs",
                   205)
    End Sub
End Class

Public Class FFRComplianceQuestionsVGridView
    Inherits FFRStatementsVGridView

    Public Sub New(ByVal setModelID As Integer)
        MyBase.New(setModelID,
                   "Compliance Questions",
                   "FFR Compliance Questions",
                   106)
    End Sub
End Class

Public Class FFRKeyDefinitionsVGridView
    Inherits FFRStatementsVGridView
    Public Sub New(ByVal setModelID As Integer)
        MyBase.New(setModelID, "FFR Key Defn", "FFR Key Definitions", 227, True)
    End Sub
End Class
