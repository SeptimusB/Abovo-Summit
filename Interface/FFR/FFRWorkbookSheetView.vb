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
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid

''' <summary>
''' Native, workbook-backed presentation for one FFR worksheet.  The grid is an
''' in-memory snapshot: it never binds an editable DevExpress RangeDataSource to
''' the authoritative workbook.  Accepted edits are submitted once through the
''' model change manager and the snapshot is then rebuilt from the calculated
''' workbook.
''' </summary>
Public Class FFRWorkbookSheetView
    Inherits XtraUserControl

    Private Const SourceRowField As String = "__SourceRow"
    Private Const DefaultRowHeight As Integer = 24

    Private ReadOnly ModelID As Integer
    Private ReadOnly SheetName As String
    Private ReadOnly Workbook As IWorkbook
    Private ReadOnly ChangeManager As ModelChangeManager
    Private ReadOnly SheetGrid As New GridControl()
    Private ReadOnly SheetView As New GridView()
    Private ReadOnly CellToolTips As New ToolTipController()
    Private ReadOnly SourceColumns As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
    Private ReadOnly ValidationEditors As New Dictionary(Of String, RepositoryItemComboBox)(StringComparer.Ordinal)
    Private ReadOnly TextEditor As New RepositoryItemTextEdit()
    Private ReadOnly DateEditor As New RepositoryItemDateEdit()
    Private ReadOnly StatusLabel As New LabelControl()

    Private LoadingSnapshot As Boolean
    Private DisposedView As Boolean
    Private Snapshot As DataTable
    Private UsedTopRow As Integer
    Private UsedLeftColumn As Integer

    Public Event WorkbookCellChanged As EventHandler

    Public Sub New(SetModelID As Integer, SetSheetName As String)
        ModelID = SetModelID
        SheetName = SetSheetName
        Workbook = FileManager.GetWorkBook(ModelID)
        ChangeManager = ExcelModels(ModelID).ChangeManager

        Dock = DockStyle.Fill
        BuildNativeSurface()
        RefreshFromWorkbook()
    End Sub

    Public ReadOnly Property WorksheetName As String
        Get
            Return SheetName
        End Get
    End Property

    Private Sub BuildNativeSurface()
        Dim Header As New PanelControl With {
            .Dock = DockStyle.Top,
            .Height = 34,
            .BorderStyle = BorderStyles.NoBorder,
            .Padding = New Padding(10, 6, 10, 4)
        }
        StatusLabel.Dock = DockStyle.Fill
        StatusLabel.Appearance.ForeColor = Color.FromArgb(32, 58, 89)
        StatusLabel.Appearance.Options.UseForeColor = True
        Header.Controls.Add(StatusLabel)

        SheetGrid.Dock = DockStyle.Fill
        SheetGrid.MainView = SheetView
        SheetGrid.ViewCollection.Add(SheetView)
        SheetGrid.ToolTipController = CellToolTips
        SheetGrid.RepositoryItems.AddRange(New RepositoryItem() {TextEditor, DateEditor})

        With SheetView
            .OptionsBehavior.AllowAddRows = DefaultBoolean.False
            .OptionsBehavior.AllowDeleteRows = DefaultBoolean.False
            .OptionsBehavior.Editable = True
            .OptionsBehavior.EditorShowMode = EditorShowMode.MouseDownFocused
            .OptionsCustomization.AllowColumnMoving = False
            .OptionsCustomization.AllowFilter = False
            .OptionsCustomization.AllowGroup = False
            .OptionsCustomization.AllowSort = False
            .OptionsCustomization.AllowQuickHideColumns = False
            .OptionsMenu.EnableColumnMenu = False
            .OptionsMenu.EnableFooterMenu = False
            .OptionsSelection.EnableAppearanceFocusedRow = False
            .OptionsSelection.MultiSelect = True
            .OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect
            .OptionsClipboard.CopyColumnHeaders = DefaultBoolean.False
            .OptionsView.ColumnAutoWidth = False
            .OptionsView.ShowAutoFilterRow = False
            .OptionsView.ShowFilterPanelMode = ShowFilterPanelMode.Never
            .OptionsView.ShowGroupPanel = False
            .OptionsView.ShowIndicator = True
            .OptionsView.ShowHorizontalLines = DefaultBoolean.True
            .OptionsView.ShowVerticalLines = DefaultBoolean.True
            .OptionsView.RowAutoHeight = False
            .RowHeight = DefaultRowHeight
            .IndicatorWidth = 54
        End With

        TextEditor.NullText = String.Empty
        DateEditor.NullText = String.Empty
        DateEditor.Buttons.Clear()
        DateEditor.Buttons.Add(New EditorButton(ButtonPredefines.Combo))

        AddHandler SheetView.CustomColumnDisplayText, AddressOf SheetViewCustomColumnDisplayText
        AddHandler SheetView.RowCellStyle, AddressOf SheetViewRowCellStyle
        AddHandler SheetView.CustomRowCellEdit, AddressOf SheetViewCustomRowCellEdit
        AddHandler SheetView.ShowingEditor, AddressOf SheetViewShowingEditor
        AddHandler SheetView.CellValueChanged, AddressOf SheetViewCellValueChanged
        AddHandler SheetView.CustomDrawRowIndicator, AddressOf SheetViewCustomDrawRowIndicator
        AddHandler SheetView.ShownEditor, AddressOf SheetViewShownEditor
        AddHandler SheetView.CalcRowHeight, AddressOf SheetViewCalcRowHeight
        AddHandler CellToolTips.GetActiveObjectInfo, AddressOf CellToolTipsGetActiveObjectInfo

        Controls.Add(SheetGrid)
        Controls.Add(Header)
    End Sub

    Public Sub RefreshFromWorkbook()
        If DisposedView OrElse Workbook Is Nothing OrElse
           Not Workbook.Worksheets.Contains(SheetName) Then Return

        Dim PreviousTopRow As Integer = SheetView.TopRowIndex
        Dim PreviousSourceRow As Integer = -1
        Dim PreviousSourceColumn As Integer = -1
        Dim FocusedCell As Cell = Nothing
        If TryGetSourceCell(SheetView.FocusedRowHandle, SheetView.FocusedColumn, FocusedCell) Then
            PreviousSourceRow = FocusedCell.RowIndex
            PreviousSourceColumn = FocusedCell.ColumnIndex
        End If

        LoadingSnapshot = True
        SheetGrid.BeginUpdate()
        Try
            ClearValidationEditors()
            SourceColumns.Clear()

            Dim Worksheet As Worksheet = Workbook.Worksheets(SheetName)
            Dim UsedRange As CellRange = Worksheet.GetUsedRange()
            UsedTopRow = UsedRange.TopRowIndex
            UsedLeftColumn = UsedRange.LeftColumnIndex

            Snapshot = BuildSnapshot(Worksheet, UsedRange)
            SheetGrid.DataSource = Snapshot
            ConfigureColumns(Worksheet, UsedRange)
            StatusLabel.Text = BuildStatusText(Worksheet, UsedRange)

            RestoreFocus(PreviousSourceRow, PreviousSourceColumn)
            SheetView.TopRowIndex = Math.Max(0, Math.Min(PreviousTopRow, Math.Max(0, Snapshot.Rows.Count - 1)))
        Finally
            SheetGrid.EndUpdate()
            LoadingSnapshot = False
        End Try
    End Sub

    Private Function BuildSnapshot(Worksheet As Worksheet, UsedRange As CellRange) As DataTable
        Dim Result As New DataTable(SheetName)
        Result.Locale = CultureInfo.CurrentCulture
        Result.Columns.Add(SourceRowField, GetType(Integer))

        For ColumnIndex As Integer = UsedRange.LeftColumnIndex To UsedRange.RightColumnIndex
            Dim FieldName As String = SourceColumnField(ColumnIndex)
            Result.Columns.Add(FieldName, GetType(Object))
            SourceColumns(FieldName) = ColumnIndex
        Next

        For RowIndex As Integer = UsedRange.TopRowIndex To UsedRange.BottomRowIndex
            Dim Row As DataRow = Result.NewRow()
            Row(SourceRowField) = RowIndex
            For ColumnIndex As Integer = UsedRange.LeftColumnIndex To UsedRange.RightColumnIndex
                Row(SourceColumnField(ColumnIndex)) = CellToSnapshotValue(Worksheet.Cells(RowIndex, ColumnIndex))
            Next
            Result.Rows.Add(Row)
        Next

        Return Result
    End Function

    Private Sub ConfigureColumns(Worksheet As Worksheet, UsedRange As CellRange)
        SheetView.PopulateColumns()
        Dim SourceRowColumn As GridColumn = SheetView.Columns(SourceRowField)
        If SourceRowColumn IsNot Nothing Then SourceRowColumn.Visible = False

        Dim VisibleIndex As Integer = 0
        For ColumnIndex As Integer = UsedRange.LeftColumnIndex To UsedRange.RightColumnIndex
            Dim Column As GridColumn = SheetView.Columns(SourceColumnField(ColumnIndex))
            If Column Is Nothing Then Continue For

            Column.Caption = ColumnName(ColumnIndex)
            Column.Tag = ColumnIndex
            Column.VisibleIndex = VisibleIndex
            Column.OptionsColumn.AllowEdit = True
            Column.OptionsColumn.AllowFocus = True
            Column.OptionsColumn.AllowMerge = DefaultBoolean.False
            Column.MinWidth = 36
            Column.Width = Math.Max(44, Math.Min(420, CInt(Math.Round(Worksheet.Columns(ColumnIndex).Width * 1.15F))))
            VisibleIndex += 1
        Next
    End Sub

    Private Function BuildStatusText(Worksheet As Worksheet, UsedRange As CellRange) As String
        Dim EditableCount As Integer
        For RowIndex As Integer = UsedRange.TopRowIndex To UsedRange.BottomRowIndex
            For ColumnIndex As Integer = UsedRange.LeftColumnIndex To UsedRange.RightColumnIndex
                If IsWorkbookCellEditable(Worksheet.Cells(RowIndex, ColumnIndex)) Then EditableCount += 1
            Next
        Next

        If EditableCount = 0 Then
            Return SheetName & "  •  calculated workbook output (read-only)"
        End If
        Return SheetName & "  •  " & EditableCount.ToString("N0", CultureInfo.CurrentCulture) &
               " workbook-controlled input cells"
    End Function

    Private Sub SheetViewCustomColumnDisplayText(sender As Object, e As CustomColumnDisplayTextEventArgs)
        If e.ListSourceRowIndex < 0 OrElse e.Column Is Nothing Then Return
        Dim RowHandle As Integer = SheetView.GetRowHandle(e.ListSourceRowIndex)
        Dim SourceCell As Cell = Nothing
        If TryGetSourceCell(RowHandle, e.Column, SourceCell) Then e.DisplayText = SourceCell.DisplayText
    End Sub

    Private Sub SheetViewRowCellStyle(sender As Object, e As RowCellStyleEventArgs)
        Dim SourceCell As Cell = Nothing
        If Not TryGetSourceCell(e.RowHandle, e.Column, SourceCell) Then Return

        Dim Background As Color = SourceCell.FillColor
        If Background.IsEmpty OrElse Background.A = 0 Then Background = Color.White
        Dim Foreground As Color = SourceCell.Font.Color
        If Foreground.IsEmpty OrElse Foreground.A = 0 Then Foreground = Color.FromArgb(32, 58, 89)

        e.Appearance.BackColor = Background
        e.Appearance.ForeColor = Foreground
        e.Appearance.Options.UseBackColor = True
        e.Appearance.Options.UseForeColor = True
        e.Appearance.Options.UseTextOptions = True
        e.Appearance.TextOptions.WordWrap = WordWrap.Wrap

        Select Case SourceCell.Alignment.Horizontal
            Case SpreadsheetHorizontalAlignment.Center
                e.Appearance.TextOptions.HAlignment = HorzAlignment.Center
            Case SpreadsheetHorizontalAlignment.Right
                e.Appearance.TextOptions.HAlignment = HorzAlignment.Far
            Case Else
                e.Appearance.TextOptions.HAlignment = HorzAlignment.Near
        End Select

        Dim FontStyle As FontStyle = FontStyle.Regular
        If SourceCell.Font.Bold Then FontStyle = FontStyle Or FontStyle.Bold
        If SourceCell.Font.Italic Then FontStyle = FontStyle Or FontStyle.Italic
        If SourceCell.Font.UnderlineType <> UnderlineType.None Then FontStyle = FontStyle Or FontStyle.Underline
        e.Appearance.Font = New Font(Font.FontFamily, Font.Size, FontStyle)
        e.Appearance.Options.UseFont = True

        If SheetView.IsCellSelected(e.RowHandle, e.Column) Then
            e.Appearance.BackColor = Color.Beige
            e.Appearance.ForeColor = Color.Black
        End If
    End Sub

    Private Sub SheetViewCustomRowCellEdit(sender As Object, e As CustomRowCellEditEventArgs)
        Dim SourceCell As Cell = Nothing
        If Not TryGetSourceCell(e.RowHandle, e.Column, SourceCell) OrElse
           Not IsWorkbookCellEditable(SourceCell) Then Return

        Dim ValidationItems As List(Of String) = WorkbookValidationItems(SourceCell)
        If ValidationItems.Count > 0 Then
            e.RepositoryItem = GetValidationEditor(ValidationItems)
        ElseIf IsDateCell(SourceCell) Then
            e.RepositoryItem = DateEditor
        Else
            e.RepositoryItem = TextEditor
        End If
    End Sub

    Private Sub SheetViewShowingEditor(sender As Object, e As CancelEventArgs)
        Dim SourceCell As Cell = Nothing
        If Not TryGetSourceCell(SheetView.FocusedRowHandle, SheetView.FocusedColumn, SourceCell) OrElse
           Not IsWorkbookCellEditable(SourceCell) Then e.Cancel = True
    End Sub

    Private Sub SheetViewShownEditor(sender As Object, e As EventArgs)
        Dim Combo As ComboBoxEdit = TryCast(SheetView.ActiveEditor, ComboBoxEdit)
        If Combo IsNot Nothing Then Combo.ShowPopup()
    End Sub

    Private Sub SheetViewCellValueChanged(sender As Object, e As CellValueChangedEventArgs)
        If LoadingSnapshot Then Return

        Dim SourceCell As Cell = Nothing
        If Not TryGetSourceCell(e.RowHandle, e.Column, SourceCell) OrElse
           Not IsWorkbookCellEditable(SourceCell) Then
            RefreshFromWorkbook()
            Return
        End If

        Dim ChangedValue As Object = NormalizeEditValue(e.Value)
        Dim ChangeEvent As New DataChangeEvent With {
            .ModelID = ModelID,
            .Description = "FFR input updated",
            .WSName = SheetName,
            .CellAddress = SourceCell.GetReferenceA1(),
            .OriginalValue = CellToObject(SourceCell),
            .ChangedValue = ChangedValue,
            .DataFormat = DataFormatForCell(SourceCell, ChangedValue),
            .TimeStamp = Now(),
            .UserName = Environment.UserName
        }

        ChangeManager.ProcessChange(ChangeEvent)
        RefreshFromWorkbook()
        RaiseEvent WorkbookCellChanged(Me, EventArgs.Empty)
    End Sub

    Private Sub SheetViewCustomDrawRowIndicator(sender As Object, e As RowIndicatorCustomDrawEventArgs)
        If e.Info.IsRowIndicator AndAlso e.RowHandle >= 0 Then
            Dim Value As Object = SheetView.GetRowCellValue(e.RowHandle, SourceRowField)
            If Value IsNot Nothing AndAlso Value IsNot DBNull.Value Then
                e.Info.DisplayText = (Convert.ToInt32(Value) + 1).ToString(CultureInfo.InvariantCulture)
            End If
        End If
    End Sub

    Private Sub SheetViewCalcRowHeight(sender As Object, e As RowHeightEventArgs)
        If e.RowHandle < 0 Then Return
        Dim Value As Object = SheetView.GetRowCellValue(e.RowHandle, SourceRowField)
        If Value Is Nothing OrElse Value Is DBNull.Value Then Return

        Dim SourceRow As Integer = Convert.ToInt32(Value, CultureInfo.InvariantCulture)
        Dim WorkbookHeight As Single = Workbook.Worksheets(SheetName).Rows(SourceRow).Height
        e.RowHeight = Math.Max(DefaultRowHeight, Math.Min(180, CInt(Math.Ceiling(WorkbookHeight * 1.35F))))
    End Sub

    Private Sub CellToolTipsGetActiveObjectInfo(
        sender As Object,
        e As ToolTipControllerGetActiveObjectInfoEventArgs)

        Dim Hit = SheetView.CalcHitInfo(e.ControlMousePosition)
        If Not Hit.InRowCell OrElse Hit.Column Is Nothing Then Return

        Dim SourceCell As Cell = Nothing
        If Not TryGetSourceCell(Hit.RowHandle, Hit.Column, SourceCell) Then Return
        Dim Comments = SourceCell.Worksheet.Comments.GetComments(SourceCell)
        If Comments Is Nothing OrElse Comments.Count = 0 Then Return

        Dim Text As String = String.Empty
        For Each Comment As DevExpress.Spreadsheet.Comment In Comments
            If Text.Length > 0 Then Text &= Environment.NewLine & Environment.NewLine
            Text &= Comment.Text
        Next
        If Text.Length > 0 Then
            e.Info = New ToolTipControlInfo(
                SheetName & "!" & SourceCell.GetReferenceA1(), Text)
        End If
    End Sub

    Private Function TryGetSourceCell(RowHandle As Integer, Column As GridColumn, ByRef SourceCell As Cell) As Boolean
        SourceCell = Nothing
        If RowHandle < 0 OrElse Column Is Nothing OrElse Not SourceColumns.ContainsKey(Column.FieldName) Then Return False

        Dim SourceRowValue As Object = SheetView.GetRowCellValue(RowHandle, SourceRowField)
        If SourceRowValue Is Nothing OrElse SourceRowValue Is DBNull.Value Then Return False

        SourceCell = Workbook.Worksheets(SheetName).Cells(
            Convert.ToInt32(SourceRowValue, CultureInfo.InvariantCulture),
            SourceColumns(Column.FieldName))
        Return SourceCell IsNot Nothing
    End Function

    Private Function IsWorkbookCellEditable(SourceCell As Cell) As Boolean
        If SourceCell Is Nothing OrElse SourceCell.Protection.Locked Then Return False
        Return SourceCell.Fill.PatternType = PatternType.Solid
    End Function

    Private Function WorkbookValidationItems(SourceCell As Cell) As List(Of String)
        Dim Result As New List(Of String)()
        Dim Validation As DataValidation = SourceCell.Worksheet.DataValidations.GetDataValidation(SourceCell)
        If Validation Is Nothing OrElse Validation.ValidationType <> DataValidationType.List Then Return Result

        Dim Criteria As ValueObject = Validation.Criteria
        If Criteria Is Nothing OrElse Criteria.IsEmpty Then Return Result

        If Criteria.IsText Then
            For Each Item As String In Criteria.TextValue.Split(New Char() {","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                AddValidationItem(Result, Item)
            Next
        ElseIf Criteria.IsRange Then
            AddValidationRangeItems(Result, Criteria.RangeValue)
        ElseIf Criteria.IsFormula Then
            Dim Formula As String = Criteria.FormulaInvariant
            If Not String.IsNullOrWhiteSpace(Formula) Then
                Dim Reference As String = Formula.Trim().TrimStart("="c)
                Try
                    If Workbook.DefinedNames.Contains(Reference) Then
                        AddValidationRangeItems(Result, Workbook.DefinedNames.GetDefinedName(Reference).Range)
                    Else
                        AddValidationRangeItems(Result, Workbook.Range(Reference))
                    End If
                Catch
                    'The workbook remains the validation authority.  If a
                    'formula cannot be resolved here, editing still travels
                    'through the normal typed change path.
                End Try
            End If
        End If

        Return Result
    End Function

    Private Shared Sub AddValidationRangeItems(Result As List(Of String), Source As CellRange)
        If Source Is Nothing Then Return
        For RowIndex As Integer = 0 To Source.RowCount - 1
            For ColumnIndex As Integer = 0 To Source.ColumnCount - 1
                AddValidationItem(Result, Source(RowIndex, ColumnIndex).DisplayText)
            Next
        Next
    End Sub

    Private Shared Sub AddValidationItem(Result As List(Of String), Value As String)
        Dim Item As String = If(Value, String.Empty).Trim()
        If Item.Length > 0 AndAlso Not Result.Contains(Item) Then Result.Add(Item)
    End Sub

    Private Function GetValidationEditor(Items As List(Of String)) As RepositoryItemComboBox
        Dim Key As String = String.Join(ChrW(30), Items)
        Dim Editor As RepositoryItemComboBox = Nothing
        If ValidationEditors.TryGetValue(Key, Editor) Then Return Editor

        Editor = New RepositoryItemComboBox()
        Editor.Items.AddRange(Items.Cast(Of Object).ToArray())
        Editor.TextEditStyle = TextEditStyles.DisableTextEditor
        Editor.NullText = String.Empty
        SheetGrid.RepositoryItems.Add(Editor)
        ValidationEditors.Add(Key, Editor)
        Return Editor
    End Function

    Private Sub ClearValidationEditors()
        For Each Editor As RepositoryItemComboBox In ValidationEditors.Values
            SheetGrid.RepositoryItems.Remove(Editor)
            Editor.Dispose()
        Next
        ValidationEditors.Clear()
    End Sub

    Private Sub RestoreFocus(SourceRow As Integer, SourceColumn As Integer)
        If SourceRow < 0 OrElse SourceColumn < 0 OrElse Snapshot Is Nothing Then Return
        Dim Column As GridColumn = SheetView.Columns(SourceColumnField(SourceColumn))
        If Column Is Nothing Then Return

        For ListIndex As Integer = 0 To Snapshot.Rows.Count - 1
            If Convert.ToInt32(Snapshot.Rows(ListIndex)(SourceRowField), CultureInfo.InvariantCulture) = SourceRow Then
                SheetView.FocusedRowHandle = SheetView.GetRowHandle(ListIndex)
                SheetView.FocusedColumn = Column
                Exit For
            End If
        Next
    End Sub

    Private Shared Function CellToSnapshotValue(SourceCell As Cell) As Object
        If SourceCell Is Nothing OrElse SourceCell.Value.IsEmpty Then Return DBNull.Value
        If SourceCell.Value.IsNumeric Then Return SourceCell.Value.NumericValue
        If SourceCell.Value.IsBoolean Then Return SourceCell.Value.BooleanValue
        If SourceCell.Value.IsDateTime Then Return SourceCell.Value.DateTimeValue
        Return SourceCell.Value.TextValue
    End Function

    Private Shared Function CellToObject(SourceCell As Cell) As Object
        Return CellToSnapshotValue(SourceCell)
    End Function

    Private Shared Function NormalizeEditValue(Value As Object) As Object
        If Value Is Nothing OrElse Value Is DBNull.Value OrElse
           String.IsNullOrWhiteSpace(Convert.ToString(Value, CultureInfo.CurrentCulture)) Then Return Nothing
        Return Value
    End Function

    Private Shared Function DataFormatForCell(SourceCell As Cell, Value As Object) As String
        If Value Is Nothing Then Return "S"
        If IsDateCell(SourceCell) OrElse TypeOf Value Is DateTime Then Return "D"
        If SourceCell.Value.IsBoolean OrElse TypeOf Value Is Boolean Then Return "B"

        Dim NumberFormat As String = If(SourceCell.NumberFormat, String.Empty)
        If NumberFormat.Contains("%") Then Return "P"
        If SourceCell.Value.IsNumeric OrElse IsNumeric(Value) Then Return "N"
        Return "S"
    End Function

    Private Shared Function IsDateCell(SourceCell As Cell) As Boolean
        If SourceCell Is Nothing Then Return False
        If SourceCell.Value.IsDateTime Then Return True
        Dim Format As String = If(SourceCell.NumberFormat, String.Empty).ToLowerInvariant()
        Return Format.Contains("dd") OrElse Format.Contains("yy")
    End Function

    Private Shared Function SourceColumnField(ColumnIndex As Integer) As String
        Return "C" & ColumnIndex.ToString(CultureInfo.InvariantCulture)
    End Function

    Private Shared Function ColumnName(ColumnIndex As Integer) As String
        Dim Value As Integer = ColumnIndex + 1
        Dim Result As String = String.Empty
        While Value > 0
            Value -= 1
            Result = ChrW(AscW("A"c) + (Value Mod 26)) & Result
            Value \= 26
        End While
        Return Result
    End Function

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso Not DisposedView Then
            DisposedView = True
            ClearValidationEditors()
            Snapshot?.Dispose()
            TextEditor.Dispose()
            DateEditor.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class
