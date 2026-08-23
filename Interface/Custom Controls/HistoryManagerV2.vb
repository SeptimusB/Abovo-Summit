Imports System.ComponentModel
Imports Abovo
Imports DevExpress.Images
Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.Grid

''' <summary>
''' Model-scoped, read-only change journal with linear undo/redo controls.
''' The form never edits workbook cells directly; every command delegates to
''' ModelChangeManagerV2.
''' </summary>
Public NotInheritable Class HistoryManagerV2
    Inherits XtraForm

    Private ReadOnly ModelID As Integer
    Private ReadOnly Manager As ModelChangeManagerV2
    Private ReadOnly Root As New TableLayoutPanel()
    Private ReadOnly Toolbar As New FlowLayoutPanel()
    Private ReadOnly HistoryGrid As New GridControl()
    Private ReadOnly HistoryView As New GridView()
    Private ReadOnly StatusLabel As New LabelControl()
    Private ReadOnly UndoButton As New SimpleButton()
    Private ReadOnly RedoButton As New SimpleButton()
    Private ReadOnly UndoSelectedButton As New SimpleButton()
    Private ReadOnly RedoSelectedButton As New SimpleButton()
    Private ReadOnly RefreshButton As New SimpleButton()
    Private ReadOnly CloseButton As New SimpleButton()
    Private ReadOnly UndoRowEditor As New RepositoryItemButtonEdit()
    Private ReadOnly RedoRowEditor As New RepositoryItemButtonEdit()
    Private ClosingForModel As Boolean

    Public Sub New(ByVal setModelID As Integer)
        ModelID = setModelID
        Manager = FileManager.ExcelModels(ModelID).ChangeManager
        Text = "Change History"
        StartPosition = FormStartPosition.CenterParent
        MinimumSize = New Size(900, 520)
        Size = New Size(1280, 760)
        BuildSurface()
        AddHandler Manager.HistoryChanged, AddressOf Manager_HistoryChanged
        RefreshHistory()
    End Sub

    Private Sub BuildSurface()
        Root.Dock = DockStyle.Fill
        Root.ColumnCount = 1
        Root.RowCount = 3
        Root.RowStyles.Add(New RowStyle(SizeType.Absolute, 52.0F))
        Root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        Root.RowStyles.Add(New RowStyle(SizeType.Absolute, 32.0F))
        Controls.Add(Root)

        Toolbar.Dock = DockStyle.Fill
        Toolbar.FlowDirection = FlowDirection.LeftToRight
        Toolbar.WrapContents = False
        Toolbar.Padding = New Padding(8, 8, 8, 4)
        ConfigureButton(UndoButton, "Undo", AddressOf UndoButton_Click)
        ConfigureButton(RedoButton, "Redo", AddressOf RedoButton_Click)
        ConfigureButton(UndoSelectedButton, "Undo to selected", AddressOf UndoSelectedButton_Click)
        ConfigureButton(RedoSelectedButton, "Redo to selected", AddressOf RedoSelectedButton_Click)
        ConfigureButton(RefreshButton, "Refresh", AddressOf RefreshButton_Click)
        ConfigureButton(CloseButton, "Close", AddressOf CloseButton_Click)
        Toolbar.Controls.AddRange(New Control() {UndoButton, RedoButton, UndoSelectedButton, RedoSelectedButton, RefreshButton, CloseButton})
        Root.Controls.Add(Toolbar, 0, 0)

        HistoryGrid.Dock = DockStyle.Fill
        HistoryGrid.MainView = HistoryView
        HistoryGrid.ViewCollection.Add(HistoryView)
        ConfigureRowActionEditor(UndoRowEditor, "Undo through this action", "images/actions/undo_16x16.png", AddressOf UndoRowButton_Click)
        ConfigureRowActionEditor(RedoRowEditor, "Redo through this action", "images/actions/redo_16x16.png", AddressOf RedoRowButton_Click)
        HistoryGrid.RepositoryItems.AddRange(New RepositoryItem() {UndoRowEditor, RedoRowEditor})
        HistoryView.OptionsBehavior.Editable = True
        HistoryView.OptionsSelection.MultiSelect = True
        HistoryView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect
        HistoryView.OptionsClipboard.CopyColumnHeaders = DefaultBoolean.False
        HistoryView.OptionsView.ShowGroupPanel = False
        HistoryView.OptionsView.ShowAutoFilterRow = True
        HistoryView.OptionsView.ColumnAutoWidth = False
        HistoryView.OptionsMenu.EnableColumnMenu = True
        AddHandler HistoryView.FocusedRowChanged, AddressOf HistorySelectionChanged
        AddHandler HistoryView.CustomRowCellEdit, AddressOf HistoryCustomRowCellEdit
        AddHandler HistoryView.ShowingEditor, AddressOf HistoryShowingEditor
        Root.Controls.Add(HistoryGrid, 0, 1)

        StatusLabel.Dock = DockStyle.Fill
        StatusLabel.Padding = New Padding(10, 5, 5, 2)
        StatusLabel.Appearance.ForeColor = Color.FromArgb(32, 58, 89)
        Root.Controls.Add(StatusLabel, 0, 2)
    End Sub

    Private Shared Sub ConfigureRowActionEditor(ByVal editor As RepositoryItemButtonEdit,
                                                ByVal toolTip As String,
                                                ByVal imageResourceName As String,
                                                ByVal clickHandler As ButtonPressedEventHandler)
        editor.TextEditStyle = TextEditStyles.HideTextEditor
        editor.Buttons.Clear()
        Dim actionButton As New EditorButton(ButtonPredefines.Glyph) With {
            .ToolTip = toolTip, .Caption = If(toolTip.StartsWith("Undo"), "↶", "↷")}
        actionButton.ImageOptions.Image = ImageResourceCache.Default.GetImage(imageResourceName)
        editor.Buttons.Add(actionButton)
        AddHandler editor.ButtonClick, clickHandler
    End Sub

    Private Shared Sub ConfigureButton(ByVal button As SimpleButton,
                                       ByVal caption As String,
                                       ByVal clickHandler As EventHandler)
        button.Text = caption
        button.AutoSize = True
        button.Margin = New Padding(3)
        AddHandler button.Click, clickHandler
    End Sub

    Public Sub ShowForUser(Optional ByVal owner As IWin32Window = Nothing)
        RefreshHistory()
        If Visible Then
            WindowState = FormWindowState.Normal
            BringToFront()
            Activate()
        ElseIf owner IsNot Nothing Then
            Show(owner)
        Else
            Show()
        End If
    End Sub

    Public Sub CloseForModel()
        If IsDisposed Then Return
        ClosingForModel = True
        RemoveHandler Manager.HistoryChanged, AddressOf Manager_HistoryChanged
        Close()
        Dispose()
    End Sub

    Private Sub Manager_HistoryChanged(ByVal sender As Object,
                                       ByVal e As ChangeHistoryChangedEventArgsV2)
        If IsDisposed Then Return
        If InvokeRequired Then
            BeginInvoke(New MethodInvoker(AddressOf RefreshHistory))
        Else
            RefreshHistory()
        End If
    End Sub

    Private Sub RefreshHistory()
        If IsDisposed Then Return
        Dim selectedGroup As Integer = FocusedGroupID()
        HistoryGrid.DataSource = Manager.GetHistoryTable()
        HistoryView.PopulateColumns()
        ConfigureColumns()
        HistoryView.ClearSorting()
        If HistoryView.Columns("TimeStamp") IsNot Nothing Then
            HistoryView.Columns("TimeStamp").SortOrder = DevExpress.Data.ColumnSortOrder.Descending
        End If
        If selectedGroup >= 0 Then FocusGroup(selectedGroup)
        UpdateCommandState()
    End Sub

    Private Sub ConfigureColumns()
        If HistoryView.Columns.Count = 0 Then Return
        For Each column As GridColumn In HistoryView.Columns
            column.OptionsColumn.ReadOnly = True
            column.OptionsColumn.AllowEdit = False
        Next
        HistoryView.Columns("GroupID").Caption = "Action #"
        HistoryView.Columns("GroupID").Width = 70
        HistoryView.Columns("TimeStamp").Caption = "Time"
        HistoryView.Columns("TimeStamp").DisplayFormat.FormatType = FormatType.DateTime
        HistoryView.Columns("TimeStamp").DisplayFormat.FormatString = "dd/MM/yyyy HH:mm:ss"
        HistoryView.Columns("TimeStamp").Width = 135
        HistoryView.Columns("Description").Width = 260
        HistoryView.Columns("Worksheet").Width = 180
        HistoryView.Columns("Cell").Width = 75
        HistoryView.Columns("OriginalValue").Caption = "Original value"
        HistoryView.Columns("OriginalValue").Width = 135
        HistoryView.Columns("NewValue").Caption = "New value"
        HistoryView.Columns("NewValue").Width = 135
        HistoryView.Columns("User").Width = 110
        HistoryView.Columns("State").Width = 85
        HistoryView.Columns("DataType").Visible = False
        HistoryView.Columns("GroupSize").Caption = "Cells"
        HistoryView.Columns("GroupSize").Width = 55
        HistoryView.Columns("Action").Caption = String.Empty
        HistoryView.Columns("Action").Width = 42
        HistoryView.Columns("Action").Fixed = FixedStyle.Right
        HistoryView.Columns("Action").OptionsColumn.ReadOnly = False
        HistoryView.Columns("Action").OptionsColumn.AllowEdit = True
        HistoryView.Columns("Action").OptionsColumn.AllowFocus = True
    End Sub

    Private Sub HistoryCustomRowCellEdit(ByVal sender As Object, ByVal e As CustomRowCellEditEventArgs)
        If e.Column Is Nothing OrElse e.Column.FieldName <> "Action" Then Return
        Select Case Convert.ToString(HistoryView.GetRowCellValue(e.RowHandle, "Action"))
            Case "Undo"
                e.RepositoryItem = UndoRowEditor
            Case "Redo"
                e.RepositoryItem = RedoRowEditor
        End Select
    End Sub

    Private Sub HistoryShowingEditor(ByVal sender As Object, ByVal e As CancelEventArgs)
        If HistoryView.FocusedColumn Is Nothing OrElse
           HistoryView.FocusedColumn.FieldName <> "Action" Then Return
        e.Cancel = String.IsNullOrWhiteSpace(
            Convert.ToString(HistoryView.GetFocusedRowCellValue("Action")))
    End Sub

    Private Sub UndoRowButton_Click(ByVal sender As Object, ByVal e As ButtonPressedEventArgs)
        Dim groupID As Integer = FocusedGroupID()
        If groupID >= 0 Then ExecuteHistoryCommand(Function() Manager.UndoTo(groupID))
    End Sub

    Private Sub RedoRowButton_Click(ByVal sender As Object, ByVal e As ButtonPressedEventArgs)
        Dim groupID As Integer = FocusedGroupID()
        If groupID >= 0 Then ExecuteHistoryCommand(Function() Manager.RedoTo(groupID))
    End Sub

    Private Sub UpdateCommandState()
        UndoButton.Enabled = Manager.CanUndo
        RedoButton.Enabled = Manager.CanRedo
        Dim state As String = Convert.ToString(HistoryView.GetFocusedRowCellValue("State"))
        UndoSelectedButton.Enabled = state = ChangeHistoryStateV2.Applied.ToString()
        RedoSelectedButton.Enabled = state = ChangeHistoryStateV2.Undone.ToString()
        StatusLabel.Text = "Undo " & If(Manager.CanUndo, "available", "unavailable") & "   |   Redo " & If(Manager.CanRedo, "available", "unavailable") & "   |   Structural workbook changes are recorded separately and are not automatically reversible."
    End Sub

    Private Function FocusedGroupID() As Integer
        Dim value As Object = HistoryView.GetFocusedRowCellValue("GroupID")
        If value Is Nothing OrElse value Is DBNull.Value Then Return -1
        Dim result As Integer
        Return If(Integer.TryParse(value.ToString(), result), result, -1)
    End Function

    Private Sub FocusGroup(ByVal groupID As Integer)
        For rowHandle As Integer = 0 To HistoryView.DataRowCount - 1
            If Convert.ToInt32(HistoryView.GetRowCellValue(rowHandle, "GroupID")) = groupID Then
                HistoryView.FocusedRowHandle = rowHandle
                Exit For
            End If
        Next
    End Sub

    Private Sub ExecuteHistoryCommand(ByVal command As Func(Of AbovoAppCls.AbovoTransaction))
        Dim result As AbovoAppCls.AbovoTransaction = command()
        If result Is Nothing Then Return
        If result.BError Then
            XtraMessageBox.Show(Me, result.StrResponseMessage, "Change History", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        RefreshHistory()
    End Sub

    Private Sub UndoButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        ExecuteHistoryCommand(Function() Manager.Undo())
    End Sub

    Private Sub RedoButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        ExecuteHistoryCommand(Function() Manager.Redo())
    End Sub

    Private Sub UndoSelectedButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim groupID As Integer = FocusedGroupID()
        If groupID < 0 Then Return
        ExecuteHistoryCommand(Function() Manager.UndoTo(groupID))
    End Sub

    Private Sub RedoSelectedButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim groupID As Integer = FocusedGroupID()
        If groupID < 0 Then Return
        ExecuteHistoryCommand(Function() Manager.RedoTo(groupID))
    End Sub

    Private Sub RefreshButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        RefreshHistory()
    End Sub

    Private Sub CloseButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Hide()
    End Sub

    Private Sub HistorySelectionChanged(ByVal sender As Object,
                                        ByVal e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs)
        UpdateCommandState()
    End Sub

    Protected Overrides Sub OnFormClosing(ByVal e As FormClosingEventArgs)
        If Not ClosingForModel AndAlso e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            Hide()
            Return
        End If
        MyBase.OnFormClosing(e)
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        If keyData = (Keys.Control Or Keys.Z) Then
            ExecuteHistoryCommand(Function() Manager.Undo())
            Return True
        End If
        If keyData = (Keys.Control Or Keys.Y) OrElse keyData = (Keys.Control Or Keys.Shift Or Keys.Z) Then
            ExecuteHistoryCommand(Function() Manager.Redo())
            Return True
        End If
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function
End Class
