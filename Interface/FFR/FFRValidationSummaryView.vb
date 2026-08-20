Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports Abovo
Imports Abovo.FileManager
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.Grid

''' <summary>
''' Compact native presentation of the workbook's FFR Validation Summary.
''' The workbook remains the sole source of headings, instructions, calculated
''' counts, messages and resolved cell appearance.
''' </summary>
Public Class FFRValidationSummaryView
    Inherits XtraUserControl

    Private Const SheetName As String = "FFR Validation Summary"
    Private Const SourceRowField As String = "SourceRow"
    Private Const ValidationField As String = "Validation"
    Private Const CountField As String = "Count"
    Private Const MessageField As String = "Message"
    Private Const WorkspaceMaximumWidth As Integer = 1240
    Private Const WorkspaceMinimumWidth As Integer = 760
    Private Const WorkspaceHeight As Integer = 592

    Private ReadOnly Workbook As IWorkbook
    Private ReadOnly StatusLabel As New LabelControl()
    Private ReadOnly ScrollHost As New XtraScrollableControl()
    Private ReadOnly Workspace As New TableLayoutPanel()
    Private ReadOnly ProductionLabel As New LabelControl()
    Private ReadOnly InstructionLabels As New List(Of LabelControl)()
    Private ReadOnly HardGroup As New GroupControl()
    Private ReadOnly SoftGroup As New GroupControl()
    Private ReadOnly HardGrid As GridControl
    Private ReadOnly SoftGrid As GridControl
    Private DisposedView As Boolean

    Public Sub New(SetModelID As Integer)
        Workbook = FileManager.GetWorkBook(SetModelID)
        Dock = DockStyle.Fill

        BuildNativeSurface()
        HardGrid = CreateSummaryGrid()
        SoftGrid = CreateSummaryGrid()
        HardGroup.Controls.Add(HardGrid)
        SoftGroup.Controls.Add(SoftGrid)
        RefreshFromWorkbook()
    End Sub

    Public ReadOnly Property WorksheetName As String
        Get
            Return SheetName
        End Get
    End Property

    Private Sub BuildNativeSurface()
        BackColor = Color.White

        Dim Header As New PanelControl With {
            .Dock = DockStyle.Top,
            .Height = 34,
            .BorderStyle = BorderStyles.NoBorder,
            .Padding = New Padding(10, 6, 10, 4)
        }
        StatusLabel.Dock = DockStyle.Fill
        StatusLabel.Text = SheetName & "  •  calculated workbook output (read-only)"
        StatusLabel.Appearance.ForeColor = Color.FromArgb(32, 58, 89)
        StatusLabel.Appearance.Options.UseForeColor = True
        Header.Controls.Add(StatusLabel)

        ScrollHost.Dock = DockStyle.Fill
        ScrollHost.BackColor = Color.White
        ScrollHost.AutoScroll = True

        Workspace.BackColor = Color.White
        Workspace.ColumnCount = 1
        Workspace.RowCount = 4
        Workspace.Padding = New Padding(12, 6, 12, 8)
        Workspace.Margin = New Padding(0)
        Workspace.Size = New Size(WorkspaceMaximumWidth, WorkspaceHeight)
        Workspace.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        Workspace.RowStyles.Add(New RowStyle(SizeType.Absolute, 28.0F))
        Workspace.RowStyles.Add(New RowStyle(SizeType.Absolute, 250.0F))
        Workspace.RowStyles.Add(New RowStyle(SizeType.Absolute, 166.0F))
        Workspace.RowStyles.Add(New RowStyle(SizeType.Absolute, 132.0F))

        ProductionLabel.Dock = DockStyle.Fill
        ProductionLabel.AutoSizeMode = LabelAutoSizeMode.None
        ProductionLabel.Appearance.TextOptions.VAlignment = VertAlignment.Center
        Workspace.Controls.Add(ProductionLabel, 0, 0)

        Dim Instructions As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .ColumnCount = 1,
            .RowCount = 11,
            .Margin = New Padding(0)
        }
        Instructions.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        For ItemIndex As Integer = 0 To 10
            Dim Height As Single = If(ItemIndex = 5 OrElse ItemIndex = 6, 25.0F, 22.0F)
            Instructions.RowStyles.Add(New RowStyle(SizeType.Absolute, Height))
            Dim Instruction As New LabelControl With {
                .Dock = DockStyle.Fill,
                .AutoSizeMode = LabelAutoSizeMode.None,
                .Margin = New Padding(0)
            }
            Instruction.Appearance.TextOptions.VAlignment = VertAlignment.Center
            Instruction.Appearance.TextOptions.WordWrap = WordWrap.Wrap
            InstructionLabels.Add(Instruction)
            Instructions.Controls.Add(Instruction, 0, ItemIndex)
        Next
        Workspace.Controls.Add(Instructions, 0, 1)

        ConfigureSummaryGroup(HardGroup)
        ConfigureSummaryGroup(SoftGroup)
        Workspace.Controls.Add(HardGroup, 0, 2)
        Workspace.Controls.Add(SoftGroup, 0, 3)

        ScrollHost.Controls.Add(Workspace)
        AddHandler ScrollHost.Resize, AddressOf PositionWorkspace

        Controls.Add(ScrollHost)
        Controls.Add(Header)
        PositionWorkspace(Nothing, EventArgs.Empty)
    End Sub

    Private Shared Sub ConfigureSummaryGroup(Group As GroupControl)
        Group.Dock = DockStyle.Fill
        Group.Margin = New Padding(0, 4, 0, 4)
        Group.Padding = New Padding(0)
        Group.AppearanceCaption.Font = New Font(Group.AppearanceCaption.Font, FontStyle.Bold)
        Group.AppearanceCaption.Options.UseFont = True
    End Sub

    Private Function CreateSummaryGrid() As GridControl
        Dim Grid As New GridControl With {.Dock = DockStyle.Fill}
        Dim View As New GridView(Grid)
        Grid.MainView = View
        Grid.ViewCollection.Add(View)

        With View
            .OptionsBehavior.Editable = False
            .OptionsCustomization.AllowColumnMoving = False
            .OptionsCustomization.AllowFilter = False
            .OptionsCustomization.AllowGroup = False
            .OptionsCustomization.AllowSort = False
            .OptionsCustomization.AllowQuickHideColumns = False
            .OptionsMenu.EnableColumnMenu = False
            .OptionsMenu.EnableFooterMenu = False
            .OptionsSelection.EnableAppearanceFocusedCell = False
            .OptionsSelection.EnableAppearanceFocusedRow = False
            .OptionsSelection.MultiSelect = True
            .OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect
            .OptionsClipboard.CopyColumnHeaders = DefaultBoolean.False
            .OptionsView.ColumnAutoWidth = True
            .OptionsView.ShowAutoFilterRow = False
            .OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never
            .OptionsView.ShowGroupPanel = False
            .OptionsView.ShowIndicator = False
            .OptionsView.ShowHorizontalLines = DefaultBoolean.True
            .OptionsView.ShowVerticalLines = DefaultBoolean.False
            .RowHeight = 25
            .ColumnPanelRowHeight = 25
            .VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Never
            .Appearance.HeaderPanel.Font = New Font(.Appearance.HeaderPanel.Font, FontStyle.Bold)
            .Appearance.HeaderPanel.ForeColor = Color.FromArgb(0, 75, 160)
            .Appearance.HeaderPanel.Options.UseFont = True
            .Appearance.HeaderPanel.Options.UseForeColor = True
        End With

        AddHandler View.RowCellStyle, AddressOf SummaryRowCellStyle
        Return Grid
    End Function

    Public Sub RefreshFromWorkbook()
        If DisposedView OrElse Workbook Is Nothing OrElse
           Not Workbook.Worksheets.Contains(SheetName) Then Return

        Dim Sheet As Worksheet = Workbook.Worksheets(SheetName)
        RefreshInstructionSurface(Sheet)
        RefreshSummaryGroup(HardGroup, HardGrid, Sheet, 17, 18, 19, 4)
        RefreshSummaryGroup(SoftGroup, SoftGrid, Sheet, 24, 25, 26, 2)
    End Sub

    Private Sub RefreshInstructionSurface(Sheet As Worksheet)
        ProductionLabel.Text = Sheet.Cells(4, 0).DisplayText
        ApplyWorkbookCellAppearance(ProductionLabel.Appearance, Sheet.Cells(4, 0))

        For ItemIndex As Integer = 0 To InstructionLabels.Count - 1
            Dim SourceCell As Cell = Sheet.Cells(5 + ItemIndex, 0)
            InstructionLabels(ItemIndex).Text = SourceCell.DisplayText
            ApplyWorkbookCellAppearance(InstructionLabels(ItemIndex).Appearance, SourceCell)
        Next
    End Sub

    Private Sub RefreshSummaryGroup(
        Group As GroupControl,
        Grid As GridControl,
        Sheet As Worksheet,
        GroupCaptionRow As Integer,
        ColumnCaptionRow As Integer,
        FirstDataRow As Integer,
        RowCount As Integer)

        Dim GroupCaptionCell As Cell = Sheet.Cells(GroupCaptionRow, 0)
        Group.Text = GroupCaptionCell.DisplayText
        ApplyWorkbookCellAppearance(Group.AppearanceCaption, GroupCaptionCell)

        Dim PreviousTable As DataTable = TryCast(Grid.DataSource, DataTable)
        Dim Table As New DataTable(Group.Text)
        Table.Columns.Add(SourceRowField, GetType(Integer))
        Table.Columns.Add(ValidationField, GetType(String))
        Table.Columns.Add(CountField, GetType(String))
        Table.Columns.Add(MessageField, GetType(String))

        For SourceRow As Integer = FirstDataRow To FirstDataRow + RowCount - 1
            Table.Rows.Add(
                SourceRow,
                Sheet.Cells(SourceRow, 0).DisplayText,
                Sheet.Cells(SourceRow, 1).DisplayText,
                Sheet.Cells(SourceRow, 2).DisplayText)
        Next

        Grid.DataSource = Table
        Grid.Tag = SheetName
        ConfigureSummaryColumns(DirectCast(Grid.MainView, GridView), Sheet.Cells(ColumnCaptionRow, 1).DisplayText)
        If PreviousTable IsNot Nothing Then PreviousTable.Dispose()
    End Sub

    Private Shared Sub ConfigureSummaryColumns(View As GridView, CountCaption As String)
        View.PopulateColumns()
        View.Columns(SourceRowField).Visible = False

        Dim ValidationColumn As GridColumn = View.Columns(ValidationField)
        ValidationColumn.Caption = " "
        ValidationColumn.VisibleIndex = 0
        ValidationColumn.Width = 390
        ValidationColumn.MinWidth = 280

        Dim CountColumn As GridColumn = View.Columns(CountField)
        CountColumn.Caption = CountCaption
        CountColumn.VisibleIndex = 1
        CountColumn.Width = 90
        CountColumn.MinWidth = 72
        CountColumn.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Center
        CountColumn.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center

        Dim MessageColumn As GridColumn = View.Columns(MessageField)
        MessageColumn.Caption = " "
        MessageColumn.VisibleIndex = 2
        MessageColumn.Width = 650
        MessageColumn.MinWidth = 360

        For Each Column As GridColumn In View.Columns
            Column.OptionsColumn.AllowEdit = False
            Column.OptionsColumn.AllowFocus = True
            Column.OptionsColumn.AllowSort = DefaultBoolean.False
            Column.OptionsFilter.AllowFilter = False
        Next
    End Sub

    Private Sub SummaryRowCellStyle(sender As Object, e As RowCellStyleEventArgs)
        If e.RowHandle < 0 OrElse e.Column Is Nothing Then Return

        Dim View As GridView = TryCast(sender, GridView)
        If View Is Nothing Then Return
        Dim SourceRowValue As Object = View.GetRowCellValue(e.RowHandle, SourceRowField)
        If SourceRowValue Is Nothing OrElse SourceRowValue Is DBNull.Value Then Return

        Dim SourceColumn As Integer
        Select Case e.Column.FieldName
            Case ValidationField
                SourceColumn = 0
            Case CountField
                SourceColumn = 1
            Case MessageField
                SourceColumn = 2
            Case Else
                Return
        End Select

        Dim SourceCell As Cell = Workbook.Worksheets(SheetName).Cells(
            Convert.ToInt32(SourceRowValue), SourceColumn)
        ApplyWorkbookCellAppearance(e.Appearance, SourceCell)
        e.Appearance.TextOptions.VAlignment = VertAlignment.Center
        If SourceColumn = 1 Then e.Appearance.TextOptions.HAlignment = HorzAlignment.Center
    End Sub

    Private Sub PositionWorkspace(sender As Object, e As EventArgs)
        Dim AvailableWidth As Integer = Math.Max(0, ScrollHost.ClientSize.Width - 32)
        Dim DesiredWidth As Integer = Math.Min(WorkspaceMaximumWidth, Math.Max(WorkspaceMinimumWidth, AvailableWidth))
        Workspace.Size = New Size(DesiredWidth, WorkspaceHeight)
        Workspace.Location = New Point(Math.Max(16, (ScrollHost.ClientSize.Width - DesiredWidth) \ 2), 6)
    End Sub

    Private Sub ApplyWorkbookCellAppearance(Appearance As AppearanceObject, SourceCell As Cell)
        If Appearance Is Nothing OrElse SourceCell Is Nothing Then Return

        Dim Background As Color = SourceCell.FillColor
        If Background.IsEmpty OrElse Background.A = 0 Then Background = Color.White
        Dim Foreground As Color = SourceCell.Font.Color
        If Foreground.IsEmpty OrElse Foreground.A = 0 Then Foreground = Color.FromArgb(32, 58, 89)

        Appearance.BackColor = Background
        Appearance.ForeColor = Foreground
        Appearance.Options.UseBackColor = True
        Appearance.Options.UseForeColor = True
        Appearance.Options.UseTextOptions = True

        Select Case SourceCell.Alignment.Horizontal
            Case SpreadsheetHorizontalAlignment.Center
                Appearance.TextOptions.HAlignment = HorzAlignment.Center
            Case SpreadsheetHorizontalAlignment.Right
                Appearance.TextOptions.HAlignment = HorzAlignment.Far
            Case Else
                Appearance.TextOptions.HAlignment = HorzAlignment.Near
        End Select

        Dim Style As FontStyle = FontStyle.Regular
        If SourceCell.Font.Bold Then Style = Style Or FontStyle.Bold
        If SourceCell.Font.Italic Then Style = Style Or FontStyle.Italic
        If SourceCell.Font.UnderlineType <> UnderlineType.None Then Style = Style Or FontStyle.Underline
        Appearance.Font = New Font(Font.FontFamily, Font.Size, Style)
        Appearance.Options.UseFont = True
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso Not DisposedView Then
            DisposedView = True
            RemoveHandler ScrollHost.Resize, AddressOf PositionWorkspace
            TryCast(HardGrid.DataSource, DataTable)?.Dispose()
            TryCast(SoftGrid.DataSource, DataTable)?.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class
