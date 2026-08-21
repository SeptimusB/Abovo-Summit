Imports System.ComponentModel
Imports System.Data
Imports System.Diagnostics
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

'''''' <summary>
'''''' MergeDownAndPivot-style native presentation of FFR Workings.  Workbook
'''''' columns are tiled as VGrid records; worksheet section headings become
'''''' category rows and all values/styles continue to come from the workbook.
'''''' </summary>
Public Class FFRWorkingsVGridView
    Inherits XtraUserControl

    Private Const SheetName As String = "FFR Workings"
    Private Const MinimumWorkspaceWidth As Integer = 1080
    'Includes the FFR Workings validation panels.  Their Column A headings
    '(Data validations - Hard Errors / Common Soft errors) are rendered as
    'ordinary VGrid categories alongside the workings sections.
    Private Const LastPresentationRow As Integer = 545
    Private ReadOnly ModelID As Integer
    Private ReadOnly Workbook As IWorkbook
    Private ReadOnly ChangeManager As ModelChangeManager
    Private ReadOnly Host As New Panel()
    Private ReadOnly Workspace As New Panel()
    Private ReadOnly Grid As New VGridControl()
    Private ReadOnly SourceRows As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
    Private Loading As Boolean
    Private LastChangedAddress As String = String.Empty

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
        Host.AutoScroll = False
        AddHandler Host.Resize, AddressOf HostResize
        Controls.Add(Host)
        Workspace.Width = MinimumWorkspaceWidth
        Host.Controls.Add(Workspace)

        Dim title As New LabelControl With {.Text = "FFR Workings", .Location = New Point(0, 16), .AutoSizeMode = LabelAutoSizeMode.None, .Size = New Size(MinimumWorkspaceWidth, 26)}
        title.Appearance.Font = New Font(Font, FontStyle.Bold)
        title.Appearance.ForeColor = Color.FromArgb(32, 58, 89)
        Workspace.Controls.Add(title)

        Grid.Location = New Point(0, 44)
        Grid.Width = MinimumWorkspaceWidth
        Grid.OptionsBehavior.Editable = True
        Grid.OptionsBehavior.ResizeRowHeaders = False
        Grid.OptionsView.ShowButtons = False
        'All FFR grids support selecting a range of displayed cells and copying
        'that range.  This is presentation-only: workbook edits still flow only
        'through GridCellValueChanged and ModelChangeManager.
        Grid.OptionsSelectionAndFocus.MultiSelect = True
        Grid.OptionsSelectionAndFocus.MultiSelectMode = MultiSelectMode.CellSelect
        Grid.ScrollVisibility = ScrollVisibility.Vertical
        Grid.RowHeaderWidth = 410
        Grid.RecordWidth = 78
        AddHandler Grid.CustomDrawRowValueCell, AddressOf GridCustomDrawRowValueCell
        AddHandler Grid.CustomDrawRowHeaderCell, AddressOf GridCustomDrawRowHeaderCell
        AddHandler Grid.ShowingEditor, AddressOf GridShowingEditor
        AddHandler Grid.CellValueChanged, AddressOf GridCellValueChanged
        AddHandler Grid.MouseWheel, AddressOf GridMouseWheel
        AddHandler Grid.KeyDown, AddressOf GridKeyDown
        Workspace.Controls.Add(Grid)
        ResizeWorkspaceToHost()
    End Sub

    Public Sub RefreshFromWorkbook()
        If Workbook Is Nothing OrElse Not Workbook.Worksheets.Contains(SheetName) Then Return
        Dim scrollX As Integer = -Host.AutoScrollPosition.X
        Dim scrollY As Integer = -Host.AutoScrollPosition.Y
        Loading = True
        Grid.BeginUpdate()
        Try
            BuildPivot()
        Finally
            Grid.EndUpdate()
            Loading = False
        End Try
        Host.AutoScrollPosition = New Point(scrollX, scrollY)
    End Sub

    Private Sub BuildPivot()
        Dim ws As Worksheet = Workbook.Worksheets(SheetName)
        Dim used As CellRange = ws.GetUsedRange()
        Dim table As New DataTable()
        table.Columns.Add("__SourceColumn", GetType(Integer))
        For col As Integer = 4 To used.RightColumnIndex
            Dim record As DataRow = table.NewRow()
            record("__SourceColumn") = col
            table.Rows.Add(record)
        Next

        SourceRows.Clear()
        Grid.DataSource = Nothing
        Grid.Rows.Clear()
        Dim categories As New Dictionary(Of Integer, CategoryRow)()
        'GetUsedRange can end at the last populated formula block in a loaded
        'model even when the defined FFR layout continues below it.  The FFR
        'presentation boundary is deliberate and must not be truncated by that
        'heuristic.
        For row As Integer = 4 To LastPresentationRow
            If IsSectionHeading(ws, row) Then
                Continue For
            End If

            Dim label As String = ws.Cells(row, 1).DisplayText.Trim()
            If label.Length = 0 Then Continue For
            Dim field As String = "Row_" & row.ToString(CultureInfo.InvariantCulture)
            table.Columns.Add(field, GetType(Object))
            SourceRows.Add(field, row)
            For recordIndex As Integer = 0 To table.Rows.Count - 1
                Dim col As Integer = CInt(table.Rows(recordIndex)("__SourceColumn"))
                table.Rows(recordIndex)(field) = CellValue(ws.Cells(row, col))
            Next
        Next

        Dim finalMappedCaption As String = String.Empty
        If SourceRows.Count > 0 Then
            Dim finalMappedRow As Integer = SourceRows.Last().Value
            finalMappedCaption = EditorCaption(ws, finalMappedRow)
        End If

        'DevExpress creates an EditorRow for each field when the data source is
        'initialised.  Remove those root rows and create the explicit workbook
        'hierarchy below; otherwise the automatic root rows remain visible and
        'mask the Column A category hierarchy.
        Grid.DataSource = table
        Grid.ForceInitialize()
        Grid.Rows.Clear()
        For Each item In SourceRows
            Dim row As Integer = item.Value
            Dim headingRow As Integer = HeadingRowForRow(ws, row)
            Dim category As CategoryRow = EnsureCategory(ws, headingRow, categories)
            Dim editorRow As New EditorRow(item.Key)
            editorRow.Properties.FieldName = item.Key
            editorRow.Properties.Caption = EditorCaption(ws, row)
            editorRow.Properties.ToolTip = GetCellNote(ws.Cells(row, 1), ws.Cells(row, 4), ws.Cells(row, 0))
            editorRow.Properties.ReadOnly = False
            editorRow.Height = 18
            category.ChildRows.Add(editorRow)
        Next
        Grid.ForceInitialize()
        ResizeGridViewport()
        ResizeWorkspaceToHost()
#If DEBUG Then
        Debug.WriteLine("FFR Workings VGrid: " & SourceRows.Count.ToString(CultureInfo.InvariantCulture) & " mapped rows; " & categories.Count.ToString(CultureInfo.InvariantCulture) & " categories; final mapped=[" & finalMappedCaption & "]; grid height=" & Grid.Height.ToString(CultureInfo.InvariantCulture) & ".")
#End If
    End Sub

    Private Function HeadingRowForRow(ByVal ws As Worksheet, ByVal row As Integer) As Integer
        For probe As Integer = row To 4 Step -1
            If IsSectionHeading(ws, probe) Then Return probe
        Next
        Return -1
    End Function

    Private Function EnsureCategory(ByVal ws As Worksheet,
                                    ByVal headingRow As Integer,
                                    ByVal categories As Dictionary(Of Integer, CategoryRow)) As CategoryRow
        If headingRow < 0 Then headingRow = -1

        Dim category As CategoryRow = Nothing
        If categories.TryGetValue(headingRow, category) Then Return category

        Dim caption As String = If(headingRow < 0, "FFR Workings", ws.Cells(headingRow, 0).DisplayText.Trim())
        category = New CategoryRow("Category_" & headingRow.ToString(CultureInfo.InvariantCulture))
        category.Properties.Caption = caption
        category.Height = If(IsValidationTopHeading(caption), 24, 21)

        Dim parentValidationRow As Integer = ValidationTopHeadingRow(ws, headingRow)
        If IsValidationTopHeading(caption) Then
            category.Tag = "ValidationTop"
            Grid.Rows.Add(category)
        ElseIf parentValidationRow >= 0 Then
            Dim parent As CategoryRow = EnsureCategory(ws, parentValidationRow, categories)
            category.Tag = "ValidationSub"
            parent.ChildRows.Add(category)
        Else
            category.Tag = "Section"
            Grid.Rows.Add(category)
        End If

        categories.Add(headingRow, category)
        Return category
    End Function

    Private Shared Function ValidationTopHeadingRow(ByVal ws As Worksheet, ByVal row As Integer) As Integer
        If row < 0 Then Return -1
        For probe As Integer = row To 4 Step -1
            If IsSectionHeading(ws, probe) AndAlso IsValidationTopHeading(ws.Cells(probe, 0).DisplayText.Trim()) Then Return probe
        Next
        Return -1
    End Function

    Private Shared Function IsValidationTopHeading(ByVal caption As String) As Boolean
        Return caption.StartsWith("Data validations -", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function IsSectionHeading(ByVal ws As Worksheet, ByVal row As Integer) As Boolean
        Dim first As String = ws.Cells(row, 0).DisplayText.Trim()
        If first.Length = 0 OrElse ws.Cells(row, 1).DisplayText.Trim().Length <> 0 Then Return False
        Return Not IsNumeric(first)
    End Function

    Private Shared Function EditorCaption(ByVal ws As Worksheet, ByVal row As Integer) As String
        Dim lineNumber As String = ws.Cells(row, 0).DisplayText.Trim()
        Dim description As String = ws.Cells(row, 1).DisplayText.Trim()
        If lineNumber.Length = 0 Then Return description
        Return lineNumber & " - " & description
    End Function

    Private Sub HostResize(ByVal sender As Object, ByVal e As EventArgs)
        ResizeWorkspaceToHost()
        ResizeGridViewport()
    End Sub

    Private Sub GridMouseWheel(ByVal sender As Object, ByVal e As MouseEventArgs)
        'The native VGrid navigator owns vertical movement for this long
        'workings surface.  Do not redirect the wheel to the host panel.
    End Sub

    Private Sub GridKeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
        If Not e.Control OrElse e.Alt OrElse e.KeyCode <> Keys.C Then Return

        'VGridControl has native tabular clipboard output for its selected
        'cell range.  Explicitly invoke it so Ctrl+C works consistently for
        'read-only and editable workbook cells alike.
        Grid.CopyToClipboard()
        e.Handled = True
        e.SuppressKeyPress = True
    End Sub

    Private Sub ResizeWorkspaceToHost()
        Dim availableWidth As Integer = Math.Max(0, Host.ClientSize.Width)
        Dim requiredWidth As Integer = Math.Max(MinimumWorkspaceWidth, availableWidth)
        Workspace.Width = requiredWidth
        Grid.Width = requiredWidth
        Workspace.Left = 0
        Workspace.Top = 0
    End Sub

    Private Sub ResizeGridViewport()
        Grid.Height = Math.Max(400, Host.ClientSize.Height - Grid.Top)
        Workspace.Height = Grid.Bottom
    End Sub

    Private Sub GridShowingEditor(ByVal sender As Object, ByVal e As CancelEventArgs)
        If Grid.FocusedRow Is Nothing Then e.Cancel = True : Return
        Dim row As Integer
        If Not SourceRows.TryGetValue(Grid.FocusedRow.Properties.FieldName, row) Then e.Cancel = True : Return
        e.Cancel = Not IsEditable(Workbook.Worksheets(SheetName).Cells(row, SourceColumn(Grid.FocusedRecord)))
    End Sub

    Private Sub GridCellValueChanged(ByVal sender As Object, ByVal e As CellValueChangedEventArgs)
        If Loading OrElse e.Row Is Nothing Then Return
        Dim row As Integer
        If Not SourceRows.TryGetValue(e.Row.Properties.FieldName, row) Then Return
        If e.RecordIndex < 0 Then Return
        Dim cell As Cell = Workbook.Worksheets(SheetName).Cells(row, SourceColumn(e.RecordIndex))
        If Not IsEditable(cell) Then RefreshFromWorkbook() : Return
        Dim changedValue As Object = NormalizeEditValue(e.Value)
        LastChangedAddress = cell.GetReferenceA1()
#If DEBUG Then
        Debug.WriteLine("FFR Workings VGrid edit: " & SheetName & "!" & LastChangedAddress & " old=[" & cell.DisplayText & "] new=[" & Convert.ToString(changedValue, CultureInfo.CurrentCulture) & "]")
#End If
        Dim change As New DataChangeEvent With {.ModelID = ModelID, .Description = "FFR workings input updated", .WSName = SheetName, .CellAddress = cell.GetReferenceA1(), .OriginalValue = CellValue(cell), .ChangedValue = changedValue, .DataFormat = DataFormatForCell(cell, changedValue), .TimeStamp = Now(), .UserName = Environment.UserName}
        If ChangeManager.ProcessChange(change).BError Then RefreshFromWorkbook() : Return
        BeginInvoke(New MethodInvoker(AddressOf RefreshAfterWorkbookCalculation))
        RaiseEvent WorkbookCellChanged(Me, EventArgs.Empty)
    End Sub

    Private Sub GridCustomDrawRowValueCell(ByVal sender As Object, ByVal e As CustomDrawRowValueCellEventArgs)
        If e.Row Is Nothing Then Return
        Dim row As Integer
        If Not SourceRows.TryGetValue(e.Row.Properties.FieldName, row) Then Return
        Dim cell As Cell = Workbook.Worksheets(SheetName).Cells(row, SourceColumn(e.RecordIndex))
        e.CellText = cell.DisplayText
        e.Appearance.BackColor = If(cell.FillColor.IsEmpty, Color.White, cell.FillColor)
        e.Appearance.ForeColor = DisplayForeground(cell)
        e.Appearance.Font = New Font(Font, If(cell.Font.Bold, FontStyle.Bold, FontStyle.Regular))
    End Sub

    Private Sub GridCustomDrawRowHeaderCell(ByVal sender As Object, ByVal e As CustomDrawRowHeaderCellEventArgs)
        Dim category As CategoryRow = TryCast(e.Row, CategoryRow)
        If category IsNot Nothing Then
            Dim style As String = TryCast(category.Tag, String)
            If style = "ValidationTop" Then
                e.Appearance.ForeColor = Color.FromArgb(0, 85, 170)
                e.Appearance.Font = New Font(Font, FontStyle.Bold)
            ElseIf style = "ValidationSub" Then
                e.Appearance.ForeColor = Color.FromArgb(32, 58, 89)
                e.Appearance.Font = New Font(Font, FontStyle.Bold)
            Else
                Return
            End If

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

    Private Sub RefreshAfterWorkbookCalculation()
        If IsDisposed OrElse Not IsHandleCreated Then Return

        'ProcessChange performs the normal workbook calculation.  Recalculate and
        'then rebuild on the UI queue so formula-backed cells (including cells
        'dependent on the changed FFR input) are read after that calculation.
        ExcelModels(ModelID).WBCalcEngine.CalculateWSs()
        'FFR Workings is a standalone workbook-backed VGrid rather than a
        'registered DataInterfaceTemplate worksheet.  Calculate it explicitly
        'so same-sheet dependent formulas (for example E18 after E10 changes)
        'are current before the presentation reads them back.
        Workbook.Worksheets(SheetName).Calculate()
        RefreshFromWorkbook()
        Grid.RefreshDataSource()
        Grid.Refresh()
#If DEBUG Then
        If LastChangedAddress.Length > 0 Then
            Debug.WriteLine("FFR Workings VGrid refreshed: " & SheetName & "!" & LastChangedAddress & " now=[" & Workbook.Worksheets(SheetName).Cells(LastChangedAddress).DisplayText & "]")
        End If
#End If
    End Sub

    Private Shared Function NormalizeEditValue(ByVal value As Object) As Object
        If value Is Nothing OrElse value Is DBNull.Value OrElse
           String.IsNullOrWhiteSpace(Convert.ToString(value, CultureInfo.CurrentCulture)) Then Return Nothing
        Return value
    End Function

    Private Shared Function DataFormatForCell(ByVal cell As Cell, ByVal value As Object) As String
        If value Is Nothing Then Return "S"
        If cell.Value.IsDateTime OrElse TypeOf value Is DateTime Then Return "D"
        If cell.Value.IsBoolean OrElse TypeOf value Is Boolean Then Return "B"
        If If(cell.NumberFormat, String.Empty).Contains("%") Then Return "P"
        If cell.Value.IsNumeric OrElse IsNumeric(value) Then Return "N"
        Return "S"
    End Function

    Private Shared Function IsEditable(ByVal cell As Cell) As Boolean
        Return cell IsNot Nothing AndAlso Not cell.Protection.Locked AndAlso cell.Fill.PatternType = PatternType.Solid
    End Function
End Class
