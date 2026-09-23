Imports System.Data
Imports System.ComponentModel
Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms
Imports Abovo
Imports Abovo.FileManager
Imports Abovo.PresentationManager
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas
Imports DevExpress.Utils
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Grid

''' <summary>
''' A read-only-by-default MappedTable projection. Optional Yes/No inputs use
''' workbook validation and ChangeManager; navigation columns never edit the sheet.
''' </summary>
Public NotInheritable Class ReadOnlyMappedTableGrid
    Inherits GridControl

    Friend Property IsAuthoringPreview As Boolean
    Friend Function AuthoringSelectedCell() As Cell
        Return SourceCell(view.FocusedRowHandle, view.FocusedColumn)
    End Function

    Private ReadOnly sourceSheet As Worksheet
    Private ReadOnly sourceRange As CellRange
    Private ReadOnly view As GridView
    Private ReadOnly values As New DataTable()
    Private ReadOnly workbookFonts As New Dictionary(Of String, Font)()
    Private ReadOnly conditionalForeground As New Dictionary(Of String, Color)()
    Private ReadOnly ownerModelID As Integer
    Private ReadOnly host As GroupInterfaceTemplate
    Private ReadOnly mapping As MappedTable
    Private ReadOnly returnGroup As Integer
    Private ReadOnly returnInterface As Integer
    Private ReadOnly returnName As String
    Private ReadOnly navigationMenu As New ContextMenuStrip()
    Private ReadOnly sourceRows As New List(Of Integer)()
    Private ReadOnly linkColumnIndex As Integer = -1
    Private ReadOnly targetColumnIndex As Integer = -1
    Private Const InterfaceField As String = "SummitInterface"
    Private ReadOnly choiceEditors As New Dictionary(Of String, RepositoryItemComboBox)()
    Private ReadOnly changeManager As ModelChangeManagerV2
    Private refreshing As Boolean
    Private posting As Boolean

    Private ReadOnly Property ShowDestinations As Boolean
        Get
            Return mapping.ShowLinkDestinations AndAlso linkColumnIndex >= sourceRange.LeftColumnIndex AndAlso
                linkColumnIndex <= sourceRange.RightColumnIndex AndAlso targetColumnIndex >= 0
        End Get
    End Property

    Public ReadOnly Property WorksheetName As String
        Get
            Return sourceSheet.Name
        End Get
    End Property

    Public Sub New(modelID As Integer, sheet As Worksheet, definition As MappedTable,
                   parentGroup As GroupInterfaceTemplate, groupID As Integer,
                   interfaceID As Integer, interfaceName As String)
        sourceSheet = sheet
        mapping = definition
        ownerModelID = modelID
        host = parentGroup
        returnGroup = groupID
        returnInterface = interfaceID
        returnName = interfaceName
        sourceRange = sheet.Range(mapping.ReadOnlyRange)
        changeManager = ExcelModels(modelID).ChangeManager
        If sourceRange.RowCount > 10000 OrElse sourceRange.ColumnCount > 100 Then
            Throw New InvalidOperationException("The read-only mapped table range is too large.")
        End If
        If Not String.IsNullOrWhiteSpace(mapping.LinkColumn) Then
            linkColumnIndex = sheet.Cells(mapping.LinkColumn.Trim() & "1").ColumnIndex
            targetColumnIndex = sheet.Cells(mapping.LinkTargetColumn.Trim() & "1").ColumnIndex
        End If

        view = New GridView(Me)
        MainView = view
        ViewCollection.Add(view)
        view.OptionsBehavior.Editable = mapping.AllowYesNoEdits
        view.OptionsBehavior.ReadOnly = Not mapping.AllowYesNoEdits
        'Do not regenerate columns when the lazy tab first creates its handle:
        'that discards workbook captions and widths applied before display.
        view.OptionsBehavior.AutoPopulateColumns = False
        view.OptionsSelection.MultiSelect = True
        view.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect
        view.OptionsSelection.EnableAppearanceFocusedCell = False
        view.OptionsSelection.EnableAppearanceFocusedRow = False
        view.OptionsView.ShowGroupPanel = False
        view.OptionsView.ShowIndicator = False
        If String.Equals(sheet.Name, "Check Sheet", StringComparison.OrdinalIgnoreCase) Then
            view.OptionsView.ShowHorizontalLines = DefaultBoolean.False
            view.OptionsView.ShowVerticalLines = DefaultBoolean.False
        End If
        view.OptionsView.ColumnAutoWidth = False
        view.OptionsView.RowAutoHeight = True
        view.OptionsCustomization.AllowFilter = False
        view.OptionsCustomization.AllowSort = False
        view.OptionsCustomization.AllowGroup = False
        view.OptionsCustomization.AllowColumnMoving = False
        view.OptionsMenu.EnableColumnMenu = False
        view.OptionsMenu.EnableGroupPanelMenu = False
        view.OptionsMenu.EnableFooterMenu = False
        view.OptionsClipboard.AllowCopy = DefaultBoolean.True
        view.OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never

        For columnIndex As Integer = sourceRange.LeftColumnIndex To sourceRange.RightColumnIndex
            values.Columns.Add("C" & columnIndex.ToString(), GetType(String))
        Next
        If ShowDestinations Then values.Columns.Add(InterfaceField, GetType(String))
        For rowIndex As Integer = sourceRange.TopRowIndex To sourceRange.BottomRowIndex
            'Keep workbook section headings, total rows and intentional blank rows.
            sourceRows.Add(rowIndex)
            values.Rows.Add(values.NewRow())
        Next
        DataSource = values
        For Each dataColumn As DataColumn In values.Columns
            view.Columns.AddVisible(dataColumn.ColumnName)
        Next
        Dim textDisplay As New DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit With {
            .ReadOnly = True}
        RepositoryItems.Add(textDisplay)
        ConfigureChoiceEditors()
        For Each column As DevExpress.XtraGrid.Columns.GridColumn In view.Columns
            Dim sheetColumn As Integer = StyleColumnIndex(column)
            column.Caption = If(mapping.HeaderRow > 0,
                sheet.Cells(mapping.HeaderRow - 1, sheetColumn).DisplayText,
                sheet.Columns(sheetColumn).Heading)
            If ShowDestinations AndAlso sheetColumn = linkColumnIndex Then
                column.Caption = If(column.FieldName = InterfaceField, "Summit interface", "Worksheet")
            End If
            Dim canEdit As Boolean = column.FieldName <> InterfaceField AndAlso
                sourceRows.Any(Function(row) choiceEditors.ContainsKey(sheet.Cells(row, sheetColumn).GetReferenceA1()))
            column.OptionsColumn.AllowEdit = canEdit
            column.OptionsColumn.ReadOnly = Not canEdit
            column.OptionsFilter.AllowFilter = False
            column.ColumnEdit = textDisplay
            column.ToolTip = If(sheetColumn = linkColumnIndex,
                If(ShowDestinations, If(column.FieldName = InterfaceField,
                    "Click to open the named Summit interface. Multiple destinations offer a choice.",
                    "Click to open this worksheet in the model's spreadsheet window."),
                    "Click a link for Summit interfaces or the workbook sheet. Enter also opens the choices."),
                column.Caption)
        Next
        AddHandler view.RowCellStyle, AddressOf StyleCell
        AddHandler view.CustomDrawColumnHeader,
            Sub(sender, e)
                If e.Column Is Nothing OrElse mapping.HeaderRow < 1 Then Return
                Dim cell As Cell = sourceSheet.Cells(mapping.HeaderRow - 1,
                    StyleColumnIndex(e.Column))
                e.Appearance.Font = GetWorkbookFont(cell)
                e.Appearance.ForeColor = cell.Font.Color
                e.Appearance.BackColor = If(cell.Fill.BackgroundColor.IsEmpty, Color.White, cell.Fill.BackgroundColor)
                e.Appearance.Options.UseFont = True
                e.Appearance.Options.UseForeColor = True
                e.Appearance.Options.UseBackColor = True
            End Sub
        AddHandler view.RowCellClick, AddressOf ClickCell
        AddHandler view.KeyDown, AddressOf GridKeyDown
        AddHandler view.CustomRowCellEdit, AddressOf ChooseCellEditor
        AddHandler view.ShowingEditor, AddressOf ShowingCellEditor
        AddHandler view.CellValueChanged, AddressOf CellValueChanged
        If changeManager IsNot Nothing Then AddHandler changeManager.HistoryChanged, AddressOf HistoryChanged
        AddHandler view.PopupMenuShowing, Sub(sender, e) e.Allow = False
        WorkbookGridClipboardSupport.AddCopyWithHeadersMenu(Me)
        RefreshData()
    End Sub

    Public Sub RefreshData()
        If IsDisposed OrElse posting OrElse refreshing Then Return
        refreshing = True
        view.BeginUpdate()
        Try
            'An undo/refresh must not leave the old in-place value covering the workbook.
            view.HideEditor()
            RefreshConditionalColours()
            Dim routes As New Dictionary(Of String, List(Of ElementInterfaceLinkTag))(StringComparer.OrdinalIgnoreCase)
            'Update the same rows in place: keep selection, scroll and column widths.
            For row As Integer = 0 To sourceRows.Count - 1
                For column As Integer = 0 To sourceRange.ColumnCount - 1
                    Dim text As String = sourceSheet.Cells(sourceRows(row),
                        sourceRange.LeftColumnIndex + column).DisplayText
                    If ShowDestinations AndAlso sourceRange.LeftColumnIndex + column = linkColumnIndex Then
                        text = LinkTarget(sourceRows(row))
                    End If
                    If Not Object.Equals(values.Rows(row)(column), text) Then
                        values.Rows(row)(column) = text
                    End If
                Next
                If ShowDestinations Then
                    Dim target As String = LinkTarget(sourceRows(row))
                    Dim label As String = String.Empty
                    If target.Length > 0 Then
                        If Not routes.ContainsKey(target) Then
                            routes.Add(target, FindInterfaces(ExcelModels(ownerModelID).WBStructure, target, mapping.InterfaceLinks))
                        End If
                        label = If(routes(target).Count = 0, "No Summit interface",
                            String.Join(Environment.NewLine, routes(target).Select(Function(route) route.LinkTip)))
                    End If
                    If Not Object.Equals(values.Rows(row)(InterfaceField), label) Then values.Rows(row)(InterfaceField) = label
                End If
            Next
        Finally
            view.EndUpdate()
            refreshing = False
        End Try
        Invalidate()
    End Sub

    Private Sub RefreshConditionalColours()
        conditionalForeground.Clear()
        'Check Sheet's error colour is a workbook formula rule (=B9<>0),
        'not a hard-coded interpretation of the word displayed in Status.
        'Parse at the rule origin and rebase through the supported formula API.
        'Only evaluate during refresh, never while painting.
        Dim engine = ExcelModels(ownerModelID).WB.FormulaEngine
        For Each rule In sourceSheet.ConditionalFormattings.
                OfType(Of FormulaExpressionConditionalFormatting)().OrderBy(Function(item) item.Priority)
            If rule.Formatting.Font.Color.IsEmpty Then Continue For
            Dim top As Integer = Math.Max(rule.Range.TopRowIndex, sourceRange.TopRowIndex)
            Dim bottom As Integer = Math.Min(rule.Range.BottomRowIndex, sourceRange.BottomRowIndex)
            Dim left As Integer = Math.Max(rule.Range.LeftColumnIndex, sourceRange.LeftColumnIndex)
            Dim right As Integer = Math.Min(rule.Range.RightColumnIndex, sourceRange.RightColumnIndex)
            If top > bottom OrElse left > right Then Continue For
            Dim expression = engine.Parse(rule.Expression, New ExpressionContext(
                rule.Range.LeftColumnIndex, rule.Range.TopRowIndex, sourceSheet))
            Dim origin As New ExpressionContext(rule.Range.LeftColumnIndex,
                rule.Range.TopRowIndex, sourceSheet) With {.ReferenceStyle = ReferenceStyle.R1C1}
            Dim relativeFormula As String = expression.ToString(origin)
            For row As Integer = top To bottom
                For column As Integer = left To right
                    Dim context As New ExpressionContext(column, row, sourceSheet) With {
                        .ReferenceStyle = ReferenceStyle.R1C1}
                    Dim result = engine.Evaluate(relativeFormula, context)
                    Dim address As String = sourceSheet.Cells(row, column).GetReferenceA1()
                    If result.IsBoolean AndAlso result.BooleanValue AndAlso
                        Not conditionalForeground.ContainsKey(address) Then
                        conditionalForeground.Add(address, rule.Formatting.Font.Color)
                    End If
                Next
            Next
        Next
    End Sub

    Public Sub FitWorkbookColumns()
        If IsDisposed Then Return
        ForceInitialize()
        If String.Equals(sourceSheet.Name, "Check Sheet", StringComparison.OrdinalIgnoreCase) AndAlso ShowDestinations Then
            view.BeginUpdate()
            Try
                FitCheckSheetColumns()
            Finally
                view.EndUpdate()
            End Try
            Return
        End If
        view.BestFitColumns()
        'Measure with the actual workbook font (BestFit may otherwise use the
        'application font). Retain room for a full error message, with scrolling.
        For Each column As DevExpress.XtraGrid.Columns.GridColumn In view.Columns
            Dim sheetColumn As Integer = StyleColumnIndex(column)
            Dim width As Integer = Math.Max(column.Width,
                CInt(sourceSheet.Columns(sheetColumn).WidthInPixels * DeviceDpi / 96.0))
            If mapping.HeaderRow > 0 Then
                Dim heading As Cell = sourceSheet.Cells(mapping.HeaderRow - 1, sheetColumn)
                width = Math.Max(width, TextRenderer.MeasureText(column.Caption,
                    GetWorkbookFont(heading)).Width + 24)
            End If
            For row As Integer = 0 To sourceRows.Count - 1
                Dim cell As Cell = sourceSheet.Cells(sourceRows(row), sheetColumn)
                Dim font As Font = GetWorkbookFont(cell)
                width = Math.Max(width, TextRenderer.MeasureText(
                    Convert.ToString(values.Rows(row)(column.FieldName)), font).Width + 24)
            Next
            column.Width = Math.Max(55, Math.Min(600, width))
        Next
        Height = Math.Min(950, Math.Max(360, view.RowCount * Math.Max(22, view.RowHeight) + 50))
    End Sub

    Private Sub FitCheckSheetColumns()
        'Check results are short; the workbook's print widths and longest error
        'message must not push both navigation destinations off the screen.
        Dim dpi As Single = DeviceDpi / 96.0F
        view.OptionsView.ColumnAutoWidth = False
        view.OptionsView.ColumnHeaderAutoHeight = DefaultBoolean.True
        view.Appearance.HeaderPanel.TextOptions.WordWrap = WordWrap.Wrap
        For Each column As DevExpress.XtraGrid.Columns.GridColumn In view.Columns
            Dim logicalWidth As Integer
            If column.FieldName = InterfaceField Then
                logicalWidth = 250
            ElseIf StyleColumnIndex(column) = linkColumnIndex Then
                logicalWidth = 215
            Else
                Select Case StyleColumnIndex(column) - sourceRange.LeftColumnIndex
                    Case 0 : logicalWidth = 230
                    Case 1, 3 : logicalWidth = 55
                    Case 2 : logicalWidth = 75
                    Case 4 : logicalWidth = 65
                    Case Else : logicalWidth = 185
                End Select
            End If
            column.MinWidth = CInt(Math.Min(logicalWidth, If(logicalWidth >= 185, 120, 45)) * dpi)
            column.Width = CInt(logicalWidth * dpi)
        Next
        view.OptionsView.ColumnAutoWidth = True
        'Keep multiline messages/destination names intact. Only the UI column
        'widths change; worksheet geometry, fonts and Yes/No editors are untouched.
        Height = Math.Min(950, Math.Max(360, view.RowCount * Math.Max(22, view.RowHeight) + 50))
    End Sub

    Private Function SourceCell(rowHandle As Integer,
                                column As DevExpress.XtraGrid.Columns.GridColumn) As Cell
        If rowHandle < 0 OrElse column Is Nothing Then Return Nothing
        Dim row As Integer = view.GetDataSourceRowIndex(rowHandle)
        If row < 0 OrElse row >= sourceRows.Count Then Return Nothing
        Return sourceSheet.Cells(sourceRows(row), StyleColumnIndex(column))
    End Function

    Private Function StyleColumnIndex(column As DevExpress.XtraGrid.Columns.GridColumn) As Integer
        Return If(column.FieldName = InterfaceField, linkColumnIndex,
                  sourceRange.LeftColumnIndex + column.AbsoluteIndex)
    End Function

    Private Function LinkTarget(row As Integer) As String
        If linkColumnIndex < 0 OrElse targetColumnIndex < 0 OrElse
            String.IsNullOrWhiteSpace(sourceSheet.Cells(row, linkColumnIndex).DisplayText) Then Return String.Empty
        Return sourceSheet.Cells(row, targetColumnIndex).DisplayText.Trim()
    End Function

    Private Function YesNoChoices(cell As Cell) As List(Of String)
        Dim result As New List(Of String)()
        If Not mapping.AllowYesNoEdits OrElse cell Is Nothing OrElse cell.Protection.Locked OrElse
            Not String.IsNullOrEmpty(cell.FormulaInvariant) Then Return result
        Dim validation As DataValidation = sourceSheet.DataValidations.GetDataValidation(cell)
        If validation Is Nothing OrElse validation.ValidationType <> DataValidationType.List OrElse
            validation.Criteria Is Nothing OrElse Not validation.Criteria.IsText Then Return result
        result = validation.Criteria.TextValue.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries).
            Select(Function(item) item.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        If result.Count <> 2 OrElse Not result.Contains("Yes", StringComparer.OrdinalIgnoreCase) OrElse
            Not result.Contains("No", StringComparer.OrdinalIgnoreCase) Then result.Clear()
        Return result
    End Function

    Private Sub ConfigureChoiceEditors()
        If Not mapping.AllowYesNoEdits Then Return
        For Each row As Integer In sourceRows
            For column As Integer = sourceRange.LeftColumnIndex To sourceRange.RightColumnIndex
                Dim cell As Cell = sourceSheet.Cells(row, column)
                Dim choices = YesNoChoices(cell)
                If choices.Count = 0 Then Continue For
                Dim editor As New RepositoryItemComboBox With {
                    .TextEditStyle = TextEditStyles.DisableTextEditor,
                    .NullText = String.Empty}
                editor.Items.AddRange(choices.Cast(Of Object)().ToArray())
                editor.Appearance.Font = GetWorkbookFont(cell)
                editor.AppearanceDropDown.Font = GetWorkbookFont(cell)
                AddHandler editor.Closed,
                    Sub(sender, e)
                        If Not posting AndAlso Not refreshing Then
                            view.PostEditor()
                            view.CloseEditor()
                        End If
                    End Sub
                RepositoryItems.Add(editor)
                choiceEditors.Add(cell.GetReferenceA1(), editor)
            Next
        Next
    End Sub

    Private Sub ChooseCellEditor(sender As Object, e As CustomRowCellEditEventArgs)
        If e.Column Is Nothing OrElse e.Column.FieldName = InterfaceField Then Return
        Dim cell As Cell = SourceCell(e.RowHandle, e.Column)
        Dim editor As RepositoryItemComboBox = Nothing
        If cell IsNot Nothing AndAlso choiceEditors.TryGetValue(cell.GetReferenceA1(), editor) Then
            e.RepositoryItem = editor
        End If
    End Sub

    Private Sub ShowingCellEditor(sender As Object, e As CancelEventArgs)
        e.Cancel = posting OrElse refreshing OrElse view.FocusedColumn Is Nothing OrElse
            view.FocusedColumn.FieldName = InterfaceField OrElse
            YesNoChoices(SourceCell(view.FocusedRowHandle, view.FocusedColumn)).Count = 0
    End Sub

    Private Sub CellValueChanged(sender As Object, e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs)
        If refreshing OrElse posting Then Return
        Dim result As AbovoAppCls.AbovoTransaction = PostChoice(SourceCell(e.RowHandle, e.Column), e.Value)
        If result.BError Then
            XtraMessageBox.Show(Me, result.StrResponseMessage, "Check Sheet", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Function PostChoice(cell As Cell, value As Object) As AbovoAppCls.AbovoTransaction
        Dim result As New AbovoAppCls.AbovoTransaction("Mapped table Yes/No input")
        Dim choice As String = YesNoChoices(cell).FirstOrDefault(
            Function(item) String.Equals(item, Convert.ToString(value).Trim(), StringComparison.OrdinalIgnoreCase))
        If choice Is Nothing OrElse changeManager Is Nothing Then
            result.BError = True
            result.StrResponseMessage = "Only an unlocked workbook Yes/No input can be changed. Select Yes or No."
            RefreshData()
            Return result
        End If
        If String.Equals(cell.Value.TextValue, choice, StringComparison.Ordinal) Then
            result.BSuccess = True
            Return result
        End If
        posting = True
        Try
            Dim description As String = sourceSheet.Cells(cell.RowIndex, sourceRange.LeftColumnIndex).DisplayText
            Return changeManager.ProcessChange(New DataChangeEvent With {
                .ModelID = ownerModelID,
                .Description = sourceSheet.Name & " - " & description & " override updated",
                .WSName = sourceSheet.Name, .CellAddress = cell.GetReferenceA1(),
                .OriginalValue = cell.Value.TextValue, .ChangedValue = choice, .DataFormat = "S",
                .TimeStamp = Now(), .UserName = Environment.UserName})
        Finally
            posting = False
            RefreshData()
        End Try
    End Function

    Private Sub HistoryChanged(sender As Object, e As ChangeHistoryChangedEventArgsV2)
        If e.WorksheetNames.Contains(sourceSheet.Name, StringComparer.OrdinalIgnoreCase) Then RefreshData()
    End Sub

    Private Function GetWorkbookFont(cell As Cell) As Font
        Dim style As FontStyle = FontStyle.Regular
        If cell.Font.Bold Then style = style Or FontStyle.Bold
        If cell.Font.Italic Then style = style Or FontStyle.Italic
        If cell.Font.UnderlineType <> UnderlineType.None Then style = style Or FontStyle.Underline
        If cell.Font.Strikethrough Then style = style Or FontStyle.Strikeout
        Dim key As String = cell.Font.Name & "|" & cell.Font.Size.ToString(
            Globalization.CultureInfo.InvariantCulture) & "|" & CInt(style).ToString()
        Dim result As Font = Nothing
        If Not workbookFonts.TryGetValue(key, result) Then
            result = New Font(cell.Font.Name, CSng(cell.Font.Size), style)
            workbookFonts.Add(key, result)
        End If
        Return result
    End Function

    Private Sub StyleCell(sender As Object, e As RowCellStyleEventArgs)
        Dim cell As Cell = SourceCell(e.RowHandle, e.Column)
        If cell Is Nothing Then Return
        e.Appearance.BackColor = If(cell.Fill.BackgroundColor.IsEmpty, Color.White, cell.Fill.BackgroundColor)
        e.Appearance.ForeColor = cell.Font.Color
        Dim conditionalColour As Color
        If conditionalForeground.TryGetValue(cell.GetReferenceA1(), conditionalColour) Then
            e.Appearance.ForeColor = conditionalColour
        End If
        e.Appearance.Font = GetWorkbookFont(cell)
        e.Appearance.Options.UseBackColor = True
        e.Appearance.Options.UseForeColor = True
        e.Appearance.Options.UseFont = True
        e.Appearance.Options.UseTextOptions = True
        e.Appearance.TextOptions.WordWrap = WordWrap.Wrap
        Select Case cell.Alignment.Horizontal
            Case SpreadsheetHorizontalAlignment.Center
                e.Appearance.TextOptions.HAlignment = HorzAlignment.Center
            Case SpreadsheetHorizontalAlignment.Right
                e.Appearance.TextOptions.HAlignment = HorzAlignment.Far
            Case Else
                e.Appearance.TextOptions.HAlignment = If(cell.Value.IsNumeric, HorzAlignment.Far, HorzAlignment.Near)
        End Select
        If view.IsCellSelected(e.RowHandle, e.Column) Then
            e.Appearance.BackColor = Color.Wheat
            e.Appearance.ForeColor = Color.Black
        End If
    End Sub

    Private Sub ClickCell(sender As Object, e As RowCellClickEventArgs)
        If IsAuthoringPreview Then Return
        'Finishing a copy rectangle on G is selection, not navigation.
        If e.Button = MouseButtons.Left AndAlso e.Clicks = 1 AndAlso
            ModifierKeys = Keys.None AndAlso view.GetSelectedCells().Length <= 1 Then
            ShowLinkChoices(SourceCell(e.RowHandle, e.Column), e.Column.FieldName = InterfaceField)
        End If
    End Sub

    Private Sub GridKeyDown(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso e.KeyCode = Keys.C Then
            WorkbookGridClipboardSupport.CopyGridSelection(Me)
            e.Handled = True
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Keys.Enter AndAlso Not IsAuthoringPreview Then
            Dim cell As Cell = SourceCell(view.FocusedRowHandle, view.FocusedColumn)
            If cell IsNot Nothing AndAlso cell.ColumnIndex = linkColumnIndex Then
                ShowLinkChoices(cell, view.FocusedColumn.FieldName = InterfaceField)
                e.Handled = True
                e.SuppressKeyPress = True
            End If
        End If
    End Sub

    Private Sub ShowLinkChoices(cell As Cell, Optional summitOnly As Boolean = False)
        If cell Is Nothing OrElse cell.ColumnIndex <> linkColumnIndex OrElse
            String.IsNullOrWhiteSpace(cell.DisplayText) OrElse targetColumnIndex < 0 Then Return
        Dim targetName As String = LinkTarget(cell.RowIndex)
        If targetName.Length = 0 Then Return
        Dim model As ExcelModel = ExcelModels(ownerModelID)
        Dim targetSheet As Worksheet = model.WB.Worksheets.FirstOrDefault(
            Function(sheet) String.Equals(sheet.Name, targetName, StringComparison.OrdinalIgnoreCase))
        If ShowDestinations AndAlso Not summitOnly AndAlso targetSheet IsNot Nothing Then
            model.ShowSpreadsheet(targetSheet)
            Return
        End If
        For Each item As ToolStripItem In navigationMenu.Items.Cast(Of ToolStripItem)().ToArray()
            item.Dispose()
        Next
        navigationMenu.Items.Clear()
        If targetSheet Is Nothing Then
            navigationMenu.Items.Add("Worksheet not found: " & targetName).Enabled = False
        Else
            Dim matches = FindInterfaces(model.WBStructure, targetSheet.Name, mapping.InterfaceLinks)
            For Each match As ElementInterfaceLinkTag In matches
                Dim destination As ElementInterfaceLinkTag = match
                destination.LinkReturnGroup = returnGroup
                destination.LinkReturnID = returnInterface
                destination.LinkReturnName = returnName
                navigationMenu.Items.Add("Open in Summit: " & match.LinkTip, Nothing,
                    Sub() model.EventCoordinator.TriggerEvent("Link", destination, host))
            Next
            If matches.Count = 0 Then
                navigationMenu.Items.Add("No Summit interface for this worksheet").Enabled = False
            End If
            If Not summitOnly Then
                navigationMenu.Items.Add(New ToolStripSeparator())
                navigationMenu.Items.Add("Open worksheet: " & targetSheet.Name, Nothing,
                    Sub() model.ShowSpreadsheet(targetSheet))
            End If
        End If
        navigationMenu.Show(Me, PointToClient(Cursor.Position))
    End Sub

    Public Shared Function FindInterfaces(modelStructure As Abovo_Model_Def,
                                          worksheetName As String,
                                          explicitLinks As List(Of MappedWorksheetInterfaceLink)) As List(Of ElementInterfaceLinkTag)
        Dim result As New List(Of ElementInterfaceLinkTag)()
        If modelStructure Is Nothing OrElse modelStructure.GroupStructures Is Nothing Then Return result
        For Each group In modelStructure.GroupStructures
            For Each child In group.ChildStructures
                Dim sectionIndex As Integer = -1
                Dim matches As Boolean = String.Equals(child.DefaultWorksheet, worksheetName,
                                                        StringComparison.OrdinalIgnoreCase)
                If Not matches AndAlso explicitLinks IsNot Nothing Then
                    matches = explicitLinks.Any(Function(link)
                        Return String.Equals(link.Worksheet, worksheetName, StringComparison.OrdinalIgnoreCase) AndAlso
                            String.Equals(link.Group, group.GSName, StringComparison.OrdinalIgnoreCase) AndAlso
                            String.Equals(link.InterfaceName, child.CSName, StringComparison.OrdinalIgnoreCase)
                    End Function)
                End If
                If Not matches Then
                    sectionIndex = child.InterfaceSections.FindIndex(Function(section)
                        If section.IElements.Any(Function(element) element.MappedTable IsNot Nothing AndAlso
                            String.Equals(element.MappedTable.Worksheet, worksheetName, StringComparison.OrdinalIgnoreCase)) Then Return True
                        Return section.ISDatasources.Any(Function(data)
                            Return data.CellRangeSources.Any(Function(source)
                                Return String.Equals(source.WSName, worksheetName, StringComparison.OrdinalIgnoreCase)
                            End Function)
                        End Function)
                    End Function)
                    matches = sectionIndex >= 0
                End If
                Dim groupID, childID As Integer
                If matches AndAlso Integer.TryParse(group.GSID, groupID) AndAlso
                    Integer.TryParse(child.CSID, childID) AndAlso childID >= 0 Then
                    result.Add(New ElementInterfaceLinkTag With {
                        .LinkGroupID = groupID, .LinkData = child.CSName,
                        .LinkTargetSection = If(sectionIndex >= 0, child.InterfaceSections(sectionIndex).ISName, Nothing),
                        .LinkTip = group.GSName & " / " & child.CSName &
                            If(sectionIndex >= 0, " / " & child.InterfaceSections(sectionIndex).ISName, String.Empty)})
                End If
            Next
        Next
        Return result
    End Function

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            If changeManager IsNot Nothing Then RemoveHandler changeManager.HistoryChanged, AddressOf HistoryChanged
            navigationMenu.Dispose()
            If ContextMenuStrip IsNot Nothing Then ContextMenuStrip.Dispose()
        End If
        MyBase.Dispose(disposing)
        If disposing Then
            values.Dispose()
            For Each font As Font In workbookFonts.Values
                font.Dispose()
            Next
            workbookFonts.Clear()
        End If
    End Sub
End Class
