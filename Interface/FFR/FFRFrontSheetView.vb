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
''' Purpose-built native presentation of the workbook's FFR Front Sheet.
''' Values, validation, protection and colours remain workbook-owned; accepted
''' edits are submitted once through ModelChangeManager and then reloaded.
''' </summary>
Public Class FFRFrontSheetView
    Inherits XtraUserControl

    Private Const SheetName As String = "Front Sheet"
    Private Const SourceRowField As String = "__SourceRow"
    Private Const EntryField As String = "Entry"
    Private Const EntityField As String = "Entity"

    Private ReadOnly ModelID As Integer
    Private ReadOnly Workbook As IWorkbook
    Private ReadOnly ChangeManager As ModelChangeManager

    Private ReadOnly WorkbookTitle As New LabelControl()
    Private ReadOnly SheetTitle As New LabelControl()
    Private ReadOnly RPNumberEdit As New TextEdit()
    Private ReadOnly FirstForecastYearEdit As New DateEdit()
    Private ReadOnly RegisteredConfirmation As New ComboBoxEdit()
    Private ReadOnly OtherConfirmation As New ComboBoxEdit()
    Private ReadOnly RegisteredQuestion As New LabelControl()
    Private ReadOnly OtherQuestion As New LabelControl()
    Private ReadOnly RegisteredNote As New LabelControl()

    Private ReadOnly RegisteredData As DataTable = CreateEntityTable()
    Private ReadOnly OtherData As DataTable = CreateEntityTable()
    Private ReadOnly RegisteredGrid As New GridControl()
    Private ReadOnly OtherGrid As New GridControl()
    Private ReadOnly RegisteredView As New GridView()
    Private ReadOnly OtherView As New GridView()
    Private ReadOnly RegisteredEntryEditor As New RepositoryItemComboBox()
    Private ReadOnly RegisteredEntityEditor As New RepositoryItemTextEdit()
    Private ReadOnly OtherEntryEditor As New RepositoryItemComboBox()
    Private ReadOnly OtherEntityEditor As New RepositoryItemTextEdit()

    Private LoadingView As Boolean
    Private DisposedView As Boolean

    Public Event WorkbookCellChanged As EventHandler

    Public Sub New(SetModelID As Integer)
        ModelID = SetModelID
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
        BackColor = Color.White

        Dim MainLayout As New TableLayoutPanel With {
            .BackColor = Color.White,
            .ColumnCount = 1,
            .Dock = DockStyle.Fill,
            .Padding = New Padding(22, 14, 22, 16),
            .RowCount = 4
        }
        MainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        MainLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 70.0F))
        MainLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 112.0F))
        MainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        MainLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0F))

        Dim TitlePanel As New PanelControl With {
            .Dock = DockStyle.Fill,
            .BorderStyle = BorderStyles.NoBorder,
            .BackColor = Color.White
        }
        WorkbookTitle.Dock = DockStyle.Top
        WorkbookTitle.Height = 28
        WorkbookTitle.Appearance.Font = New Font("Segoe UI", 9.5F, FontStyle.Regular)
        WorkbookTitle.Appearance.ForeColor = Color.FromArgb(0, 91, 170)
        WorkbookTitle.Appearance.Options.UseFont = True
        WorkbookTitle.Appearance.Options.UseForeColor = True
        SheetTitle.Dock = DockStyle.Fill
        SheetTitle.Appearance.Font = New Font("Segoe UI", 17.0F, FontStyle.Bold)
        SheetTitle.Appearance.ForeColor = Color.FromArgb(32, 58, 89)
        SheetTitle.Appearance.Options.UseFont = True
        SheetTitle.Appearance.Options.UseForeColor = True
        TitlePanel.Controls.Add(SheetTitle)
        TitlePanel.Controls.Add(WorkbookTitle)

        Dim DetailsGroup As New GroupControl With {
            .Dock = DockStyle.Fill,
            .Text = "Return details"
        }
        DetailsGroup.AppearanceCaption.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        DetailsGroup.AppearanceCaption.Options.UseFont = True
        Dim DetailsLayout As New TableLayoutPanel With {
            .ColumnCount = 4,
            .Dock = DockStyle.Fill,
            .Padding = New Padding(12, 8, 12, 8),
            .RowCount = 1
        }
        DetailsLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 150.0F))
        DetailsLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 38.0F))
        DetailsLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 310.0F))
        DetailsLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 62.0F))
        DetailsLayout.Controls.Add(CreateFieldLabel("RP Number"), 0, 0)
        ConfigureTopEditor(RPNumberEdit)
        DetailsLayout.Controls.Add(RPNumberEdit, 1, 0)
        DetailsLayout.Controls.Add(CreateFieldLabel("Year end of first forecast year"), 2, 0)
        ConfigureTopEditor(FirstForecastYearEdit)
        FirstForecastYearEdit.Properties.Buttons.Clear()
        FirstForecastYearEdit.Properties.Buttons.Add(New EditorButton(ButtonPredefines.Combo))
        FirstForecastYearEdit.Properties.DisplayFormat.FormatType = FormatType.DateTime
        FirstForecastYearEdit.Properties.DisplayFormat.FormatString = "dd/MM/yyyy"
        FirstForecastYearEdit.Properties.EditFormat.FormatType = FormatType.DateTime
        FirstForecastYearEdit.Properties.EditFormat.FormatString = "dd/MM/yyyy"
        DetailsLayout.Controls.Add(FirstForecastYearEdit, 3, 0)
        DetailsGroup.Controls.Add(DetailsLayout)

        Dim EntityLayout As New TableLayoutPanel With {
            .BackColor = Color.White,
            .ColumnCount = 2,
            .Dock = DockStyle.Fill,
            .Padding = New Padding(0, 12, 0, 0),
            .RowCount = 1
        }
        EntityLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        EntityLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        EntityLayout.Controls.Add(
            CreateEntityGroup(
                "Registered subsidiaries",
                RegisteredQuestion,
                RegisteredConfirmation,
                RegisteredNote,
                RegisteredGrid,
                RegisteredView,
                RegisteredData,
                "RP Code and Name"), 0, 0)
        EntityLayout.Controls.Add(
            CreateEntityGroup(
                "Non-registered entities and joint ventures",
                OtherQuestion,
                OtherConfirmation,
                Nothing,
                OtherGrid,
                OtherView,
                OtherData,
                "Name of unregistered entity or joint venture"), 1, 0)
        EntityLayout.SetCellPosition(EntityLayout.GetControlFromPosition(0, 0), New TableLayoutPanelCellPosition(0, 0))
        EntityLayout.GetControlFromPosition(0, 0).Margin = New Padding(0, 0, 7, 0)
        EntityLayout.GetControlFromPosition(1, 0).Margin = New Padding(7, 0, 0, 0)

        Dim Footer As New LabelControl With {
            .Dock = DockStyle.Fill,
            .Text = "Blue workbook fields are editable. All values and validation remain controlled by the open model.",
            .AutoSizeMode = LabelAutoSizeMode.None
        }
        Footer.Appearance.ForeColor = Color.DimGray
        Footer.Appearance.Options.UseForeColor = True
        Footer.Appearance.TextOptions.VAlignment = VertAlignment.Center

        MainLayout.Controls.Add(TitlePanel, 0, 0)
        MainLayout.Controls.Add(DetailsGroup, 0, 1)
        MainLayout.Controls.Add(EntityLayout, 0, 2)
        MainLayout.Controls.Add(Footer, 0, 3)
        Controls.Add(MainLayout)

        AddHandler RPNumberEdit.Validated, AddressOf RPNumberValidated
        AddHandler FirstForecastYearEdit.Validated, AddressOf FirstForecastYearValidated
        AddHandler RegisteredConfirmation.Validated, AddressOf RegisteredConfirmationValidated
        AddHandler OtherConfirmation.Validated, AddressOf OtherConfirmationValidated
    End Sub

    Private Shared Function CreateFieldLabel(Text As String) As LabelControl
        Dim Label As New LabelControl With {
            .AutoSizeMode = LabelAutoSizeMode.None,
            .Dock = DockStyle.Fill,
            .Text = Text
        }
        Label.Appearance.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
        Label.Appearance.ForeColor = Color.FromArgb(32, 58, 89)
        Label.Appearance.Options.UseFont = True
        Label.Appearance.Options.UseForeColor = True
        Label.Appearance.TextOptions.VAlignment = VertAlignment.Center
        Return Label
    End Function

    Private Shared Sub ConfigureTopEditor(Editor As BaseEdit)
        Editor.Dock = DockStyle.Fill
        Editor.Margin = New Padding(4, 12, 26, 12)
    End Sub

    Private Function CreateEntityGroup(
        Caption As String,
        Question As LabelControl,
        Confirmation As ComboBoxEdit,
        Note As LabelControl,
        Grid As GridControl,
        View As GridView,
        Data As DataTable,
        EntityCaption As String) As GroupControl

        Dim Group As New GroupControl With {
            .Dock = DockStyle.Fill,
            .Text = Caption
        }
        Group.AppearanceCaption.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        Group.AppearanceCaption.Options.UseFont = True
        Dim QuestionPanel As New PanelControl With {
            .BorderStyle = BorderStyles.NoBorder,
            .Dock = DockStyle.Top,
            .Height = If(Note Is Nothing, 86, 126),
            .Padding = New Padding(12, 8, 12, 8)
        }
        Question.AutoSizeMode = LabelAutoSizeMode.None
        Question.Dock = DockStyle.Top
        Question.Height = 38
        Question.Appearance.ForeColor = Color.FromArgb(32, 58, 89)
        Question.Appearance.Options.UseForeColor = True
        Question.Appearance.TextOptions.WordWrap = WordWrap.Wrap
        Confirmation.Dock = DockStyle.Top
        Confirmation.Height = 28
        Confirmation.Properties.TextEditStyle = TextEditStyles.DisableTextEditor
        Confirmation.Margin = New Padding(0, 4, 0, 0)
        QuestionPanel.Controls.Add(Confirmation)
        QuestionPanel.Controls.Add(Question)
        If Note IsNot Nothing Then
            Note.AutoSizeMode = LabelAutoSizeMode.None
            Note.Dock = DockStyle.Bottom
            Note.Height = 34
            Note.Appearance.ForeColor = Color.DimGray
            Note.Appearance.Options.UseForeColor = True
            Note.Appearance.TextOptions.WordWrap = WordWrap.Wrap
            QuestionPanel.Controls.Add(Note)
        End If

        ConfigureEntityGrid(Grid, View, Data, EntityCaption)
        Group.Controls.Add(Grid)
        Group.Controls.Add(QuestionPanel)
        Return Group
    End Function

    Private Sub ConfigureEntityGrid(
        Grid As GridControl,
        View As GridView,
        Data As DataTable,
        EntityCaption As String)

        Grid.Dock = DockStyle.Fill
        Grid.MainView = View
        Grid.ViewCollection.Add(View)
        Dim GridEntryEditor As RepositoryItemComboBox =
            If(View Is RegisteredView, RegisteredEntryEditor, OtherEntryEditor)
        Dim GridEntityEditor As RepositoryItemTextEdit =
            If(View Is RegisteredView, RegisteredEntityEditor, OtherEntityEditor)
        Grid.RepositoryItems.AddRange(New RepositoryItem() {GridEntryEditor, GridEntityEditor})
        Grid.DataSource = Data

        GridEntryEditor.TextEditStyle = TextEditStyles.DisableTextEditor
        GridEntityEditor.NullText = String.Empty

        With View
            .OptionsBehavior.AllowAddRows = DefaultBoolean.False
            .OptionsBehavior.AllowDeleteRows = DefaultBoolean.False
            .OptionsBehavior.Editable = True
            .OptionsCustomization.AllowColumnMoving = False
            .OptionsCustomization.AllowFilter = False
            .OptionsCustomization.AllowGroup = False
            .OptionsCustomization.AllowSort = False
            .OptionsMenu.EnableColumnMenu = False
            .OptionsSelection.EnableAppearanceFocusedRow = False
            .OptionsSelection.MultiSelect = True
            .OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect
            .OptionsClipboard.CopyColumnHeaders = DefaultBoolean.False
            .OptionsView.ColumnAutoWidth = True
            .OptionsView.ShowGroupPanel = False
            .OptionsView.ShowIndicator = False
            .OptionsView.ShowHorizontalLines = DefaultBoolean.True
            .OptionsView.ShowVerticalLines = DefaultBoolean.True
            .RowHeight = 25
            .ColumnPanelRowHeight = 32
        End With

        View.Columns.Clear()
        Dim SourceRowColumn As New GridColumn With {
            .FieldName = SourceRowField,
            .Visible = False
        }
        Dim EntryColumn As New GridColumn With {
            .Caption = "Entry",
            .FieldName = EntryField,
            .Visible = True,
            .VisibleIndex = 0,
            .Width = 72,
            .MaxWidth = 90
        }
        Dim EntityColumn As New GridColumn With {
            .Caption = EntityCaption,
            .FieldName = EntityField,
            .Visible = True,
            .VisibleIndex = 1,
            .Width = 420
        }
        View.Columns.AddRange(New GridColumn() {SourceRowColumn, EntryColumn, EntityColumn})

        AddHandler View.CustomColumnDisplayText, AddressOf EntityViewCustomColumnDisplayText
        AddHandler View.RowCellStyle, AddressOf EntityViewRowCellStyle
        AddHandler View.CustomRowCellEdit, AddressOf EntityViewCustomRowCellEdit
        AddHandler View.ShowingEditor, AddressOf EntityViewShowingEditor
        AddHandler View.ShownEditor, AddressOf EntityViewShownEditor
        AddHandler View.CellValueChanged, AddressOf EntityViewCellValueChanged
    End Sub

    Public Sub RefreshFromWorkbook()
        If DisposedView OrElse Workbook Is Nothing OrElse Not Workbook.Worksheets.Contains(SheetName) Then Return

        LoadingView = True
        RegisteredGrid.BeginUpdate()
        OtherGrid.BeginUpdate()
        Try
            Dim Worksheet As Worksheet = Workbook.Worksheets(SheetName)
            WorkbookTitle.Text = Worksheet.Cells("A1").DisplayText
            SheetTitle.Text = Worksheet.Cells("A2").DisplayText
            RegisteredQuestion.Text = Worksheet.Cells("A7").DisplayText
            OtherQuestion.Text = Worksheet.Cells("A36").DisplayText
            RegisteredNote.Text = GetCellNote(Worksheet.Cells("C9"))

            LoadEditor(RPNumberEdit, Worksheet.Cells("B5"))
            LoadEditor(FirstForecastYearEdit, Worksheet.Cells("B6"))
            LoadConfirmation(RegisteredConfirmation, Worksheet.Cells("B7"))
            LoadConfirmation(OtherConfirmation, Worksheet.Cells("B36"))
            LoadEntryEditor(RegisteredEntryEditor, Worksheet.Cells("B10"))
            LoadEntryEditor(OtherEntryEditor, Worksheet.Cells("B39"))
            LoadEntityData(RegisteredData, Worksheet, 9, 33)
            LoadEntityData(OtherData, Worksheet, 38, 62)
        Finally
            OtherGrid.EndUpdate()
            RegisteredGrid.EndUpdate()
            LoadingView = False
            RegisteredView.RefreshData()
            OtherView.RefreshData()
        End Try
    End Sub

    Private Sub LoadEditor(Editor As BaseEdit, SourceCell As Cell)
        If SourceCell.Value.IsDateTime Then
            Editor.EditValue = SourceCell.Value.DateTimeValue
        ElseIf SourceCell.Value.IsEmpty Then
            Editor.EditValue = Nothing
        Else
            Editor.EditValue = CellToObject(SourceCell)
        End If
        ApplyWorkbookAppearance(Editor, SourceCell)
        Editor.Properties.ReadOnly = Not IsWorkbookCellEditable(SourceCell)
    End Sub

    Private Sub LoadConfirmation(Editor As ComboBoxEdit, SourceCell As Cell)
        Editor.Properties.Items.Clear()
        Editor.Properties.Items.AddRange(WorkbookValidationItems(SourceCell).Cast(Of Object).ToArray())
        Editor.EditValue = If(SourceCell.Value.IsEmpty, Nothing, CellToObject(SourceCell))
        ApplyWorkbookAppearance(Editor, SourceCell)
        Editor.Properties.ReadOnly = Not IsWorkbookCellEditable(SourceCell)
    End Sub

    Private Sub LoadEntryEditor(Editor As RepositoryItemComboBox, SourceCell As Cell)
        Editor.Items.Clear()
        Editor.Items.AddRange(WorkbookValidationItems(SourceCell).Cast(Of Object).ToArray())
    End Sub

    Private Shared Sub ApplyWorkbookAppearance(Editor As BaseEdit, SourceCell As Cell)
        Dim Background As Color = SourceCell.FillColor
        If Background.IsEmpty OrElse Background.A = 0 Then Background = Color.White
        Dim Foreground As Color = SourceCell.Font.Color
        If Foreground.IsEmpty OrElse Foreground.A = 0 Then Foreground = Color.FromArgb(32, 58, 89)
        Editor.Properties.Appearance.BackColor = Background
        Editor.Properties.Appearance.ForeColor = Foreground
        Editor.Properties.Appearance.Options.UseBackColor = True
        Editor.Properties.Appearance.Options.UseForeColor = True
    End Sub

    Private Shared Sub LoadEntityData(Data As DataTable, Worksheet As Worksheet, FirstRow As Integer, LastRow As Integer)
        Data.BeginLoadData()
        Try
            Data.Clear()
            For SourceRow As Integer = FirstRow To LastRow
                Dim Row As DataRow = Data.NewRow()
                Row(SourceRowField) = SourceRow
                Row(EntryField) = CellToSnapshotValue(Worksheet.Cells(SourceRow, 1))
                Row(EntityField) = CellToSnapshotValue(Worksheet.Cells(SourceRow, 2))
                Data.Rows.Add(Row)
            Next
        Finally
            Data.EndLoadData()
        End Try
    End Sub

    Private Sub RPNumberValidated(sender As Object, e As EventArgs)
        CommitEditor("B5", RPNumberEdit.EditValue, "FFR RP number updated")
    End Sub

    Private Sub FirstForecastYearValidated(sender As Object, e As EventArgs)
        CommitEditor("B6", FirstForecastYearEdit.EditValue, "FFR first forecast year updated")
    End Sub

    Private Sub RegisteredConfirmationValidated(sender As Object, e As EventArgs)
        CommitEditor("B7", RegisteredConfirmation.EditValue, "FFR registered-subsidiary confirmation updated")
    End Sub

    Private Sub OtherConfirmationValidated(sender As Object, e As EventArgs)
        CommitEditor("B36", OtherConfirmation.EditValue, "FFR non-registered-entity confirmation updated")
    End Sub

    Private Sub CommitEditor(Address As String, Value As Object, Description As String)
        If LoadingView Then Return
        CommitWorkbookCell(Workbook.Worksheets(SheetName).Cells(Address), Value, Description)
    End Sub

    Private Sub EntityViewCustomColumnDisplayText(sender As Object, e As CustomColumnDisplayTextEventArgs)
        If LoadingView Then Return
        If e.ListSourceRowIndex < 0 OrElse e.Column Is Nothing Then Return
        Dim View As GridView = DirectCast(sender, GridView)
        Dim SourceCell As Cell = Nothing
        If TryGetEntitySourceCell(View, View.GetRowHandle(e.ListSourceRowIndex), e.Column, SourceCell) Then
            e.DisplayText = SourceCell.DisplayText
        End If
    End Sub

    Private Sub EntityViewRowCellStyle(sender As Object, e As RowCellStyleEventArgs)
        If LoadingView Then Return
        Dim View As GridView = DirectCast(sender, GridView)
        Dim SourceCell As Cell = Nothing
        If Not TryGetEntitySourceCell(View, e.RowHandle, e.Column, SourceCell) Then Return

        Dim Background As Color = SourceCell.FillColor
        If Background.IsEmpty OrElse Background.A = 0 Then Background = Color.White
        Dim Foreground As Color = SourceCell.Font.Color
        If Foreground.IsEmpty OrElse Foreground.A = 0 Then Foreground = Color.FromArgb(32, 58, 89)
        e.Appearance.BackColor = Background
        e.Appearance.ForeColor = Foreground
        e.Appearance.Options.UseBackColor = True
        e.Appearance.Options.UseForeColor = True
        If e.Column.FieldName = EntryField Then e.Appearance.TextOptions.HAlignment = HorzAlignment.Center
    End Sub

    Private Sub EntityViewCustomRowCellEdit(sender As Object, e As CustomRowCellEditEventArgs)
        If LoadingView Then Return
        Dim View As GridView = DirectCast(sender, GridView)
        Dim SourceCell As Cell = Nothing
        If Not TryGetEntitySourceCell(View, e.RowHandle, e.Column, SourceCell) OrElse
           Not IsWorkbookCellEditable(SourceCell) Then Return
        If e.Column.FieldName = EntryField Then
            e.RepositoryItem = If(View Is RegisteredView, RegisteredEntryEditor, OtherEntryEditor)
        Else
            e.RepositoryItem = If(View Is RegisteredView, RegisteredEntityEditor, OtherEntityEditor)
        End If
    End Sub

    Private Sub EntityViewShowingEditor(sender As Object, e As CancelEventArgs)
        Dim View As GridView = DirectCast(sender, GridView)
        Dim SourceCell As Cell = Nothing
        If Not TryGetEntitySourceCell(View, View.FocusedRowHandle, View.FocusedColumn, SourceCell) OrElse
           Not IsWorkbookCellEditable(SourceCell) Then e.Cancel = True
    End Sub

    Private Sub EntityViewShownEditor(sender As Object, e As EventArgs)
        Dim View As GridView = DirectCast(sender, GridView)
        Dim Combo As ComboBoxEdit = TryCast(View.ActiveEditor, ComboBoxEdit)
        If Combo IsNot Nothing Then Combo.ShowPopup()
    End Sub

    Private Sub EntityViewCellValueChanged(sender As Object, e As CellValueChangedEventArgs)
        If LoadingView Then Return
        Dim View As GridView = DirectCast(sender, GridView)
        Dim SourceCell As Cell = Nothing
        If Not TryGetEntitySourceCell(View, e.RowHandle, e.Column, SourceCell) Then
            RefreshFromWorkbook()
            Return
        End If
        CommitWorkbookCell(SourceCell, e.Value, "FFR entity list updated")
    End Sub

    Private Function TryGetEntitySourceCell(
        View As GridView,
        RowHandle As Integer,
        Column As GridColumn,
        ByRef SourceCell As Cell) As Boolean

        SourceCell = Nothing
        If RowHandle < 0 OrElse Column Is Nothing OrElse
           (Column.FieldName <> EntryField AndAlso Column.FieldName <> EntityField) Then Return False
        Dim SourceRowValue As Object = View.GetRowCellValue(RowHandle, SourceRowField)
        If SourceRowValue Is Nothing OrElse SourceRowValue Is DBNull.Value Then Return False
        Dim SourceColumn As Integer = If(Column.FieldName = EntryField, 1, 2)
        SourceCell = Workbook.Worksheets(SheetName).Cells(
            Convert.ToInt32(SourceRowValue, CultureInfo.InvariantCulture), SourceColumn)
        Return True
    End Function

    Private Sub CommitWorkbookCell(SourceCell As Cell, Value As Object, Description As String)
        If LoadingView OrElse SourceCell Is Nothing OrElse Not IsWorkbookCellEditable(SourceCell) Then
            RefreshFromWorkbook()
            Return
        End If

        Dim ChangedValue As Object = NormalizeEditValue(Value)
        If ValuesEqual(CellToObject(SourceCell), ChangedValue) Then Return

        Dim ChangeEvent As New DataChangeEvent With {
            .ModelID = ModelID,
            .Description = Description,
            .WSName = SheetName,
            .CellAddress = SourceCell.GetReferenceA1(),
            .OriginalValue = CellToObject(SourceCell),
            .ChangedValue = ChangedValue,
            .DataFormat = DataFormatForCell(SourceCell, ChangedValue),
            .TimeStamp = Now(),
            .UserName = Environment.UserName
        }

        Dim Result = ChangeManager.ProcessChange(ChangeEvent)
        RefreshFromWorkbook()
        If Result IsNot Nothing AndAlso Not Result.BError Then RaiseEvent WorkbookCellChanged(Me, EventArgs.Empty)
    End Sub

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
            Dim Reference As String = Criteria.FormulaInvariant.Trim().TrimStart("="c)
            Try
                If Workbook.DefinedNames.Contains(Reference) Then
                    AddValidationRangeItems(Result, Workbook.DefinedNames.GetDefinedName(Reference).Range)
                Else
                    AddValidationRangeItems(Result, Workbook.Range(Reference))
                End If
            Catch
                'The workbook remains authoritative if a validation formula is not resolvable here.
            End Try
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

    Private Shared Function GetCellNote(SourceCell As Cell) As String
        Dim Comments = SourceCell.Worksheet.Comments.GetComments(SourceCell)
        If Comments Is Nothing OrElse Comments.Count = 0 Then Return String.Empty
        Dim Result As String = String.Empty
        For Each Comment As DevExpress.Spreadsheet.Comment In Comments
            If Result.Length > 0 Then Result &= Environment.NewLine
            Result &= Comment.Text.Trim()
        Next
        Return Result
    End Function

    Private Shared Function IsWorkbookCellEditable(SourceCell As Cell) As Boolean
        Return SourceCell IsNot Nothing AndAlso
               Not SourceCell.Protection.Locked AndAlso
               SourceCell.Fill.PatternType = PatternType.Solid
    End Function

    Private Shared Function CreateEntityTable() As DataTable
        Dim Result As New DataTable()
        Result.Locale = CultureInfo.CurrentCulture
        Result.Columns.Add(SourceRowField, GetType(Integer))
        Result.Columns.Add(EntryField, GetType(Object))
        Result.Columns.Add(EntityField, GetType(Object))
        Return Result
    End Function

    Private Shared Function CellToSnapshotValue(SourceCell As Cell) As Object
        Dim Value As Object = CellToObject(SourceCell)
        Return If(Value Is Nothing, DBNull.Value, Value)
    End Function

    Private Shared Function CellToObject(SourceCell As Cell) As Object
        If SourceCell Is Nothing OrElse SourceCell.Value.IsEmpty Then Return Nothing
        If SourceCell.Value.IsNumeric Then Return SourceCell.Value.NumericValue
        If SourceCell.Value.IsBoolean Then Return SourceCell.Value.BooleanValue
        If SourceCell.Value.IsDateTime Then Return SourceCell.Value.DateTimeValue
        Return SourceCell.Value.TextValue
    End Function

    Private Shared Function NormalizeEditValue(Value As Object) As Object
        If Value Is Nothing OrElse Value Is DBNull.Value OrElse
           String.IsNullOrWhiteSpace(Convert.ToString(Value, CultureInfo.CurrentCulture)) Then Return Nothing
        Return Value
    End Function

    Private Shared Function ValuesEqual(Original As Object, Changed As Object) As Boolean
        If Original Is Nothing AndAlso Changed Is Nothing Then Return True
        If Original Is Nothing OrElse Changed Is Nothing Then Return False
        If TypeOf Original Is DateTime OrElse TypeOf Changed Is DateTime Then
            Dim OriginalDate As DateTime
            Dim ChangedDate As DateTime
            Return DateTime.TryParse(Convert.ToString(Original, CultureInfo.CurrentCulture), OriginalDate) AndAlso
                   DateTime.TryParse(Convert.ToString(Changed, CultureInfo.CurrentCulture), ChangedDate) AndAlso
                   OriginalDate.Date = ChangedDate.Date
        End If
        Return String.Equals(
            Convert.ToString(Original, CultureInfo.CurrentCulture),
            Convert.ToString(Changed, CultureInfo.CurrentCulture),
            StringComparison.CurrentCulture)
    End Function

    Private Shared Function DataFormatForCell(SourceCell As Cell, Value As Object) As String
        If Value Is Nothing Then Return "S"
        If SourceCell.Value.IsDateTime OrElse TypeOf Value Is DateTime Then Return "D"
        If SourceCell.Value.IsBoolean OrElse TypeOf Value Is Boolean Then Return "B"
        If If(SourceCell.NumberFormat, String.Empty).Contains("%") Then Return "P"
        If SourceCell.Value.IsNumeric OrElse IsNumeric(Value) Then Return "N"
        Return "S"
    End Function

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso Not DisposedView Then
            DisposedView = True
            RegisteredData.Dispose()
            OtherData.Dispose()
            RegisteredEntryEditor.Dispose()
            RegisteredEntityEditor.Dispose()
            OtherEntryEditor.Dispose()
            OtherEntityEditor.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class
