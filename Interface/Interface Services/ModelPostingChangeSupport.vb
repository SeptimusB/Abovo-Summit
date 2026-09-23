Imports Abovo.FileManager
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraEditors
Imports System.Windows.Forms

Namespace Abovo

    Friend Module ModelPostingChangeSupport

        'Call only inside the view's posting-suppressed workbook refresh scope.
        Friend Sub HideGridEditors(root As Control)
            Dim grid = TryCast(root, DevExpress.XtraGrid.GridControl)
            If grid IsNot Nothing Then
                For Each view In grid.ViewCollection.OfType(Of DevExpress.XtraGrid.Views.Base.ColumnView)()
                    view.HideEditor()
                Next
            End If
            Dim vertical = TryCast(root, DevExpress.XtraVerticalGrid.VGridControl)
            If vertical IsNot Nothing Then vertical.HideEditor()
            For Each child As Control In root.Controls.Cast(Of Control)().ToArray()
                HideGridEditors(child)
            Next
        End Sub

        Friend Function PostModelCellValue(ByVal modelID As Integer,
                                           ByVal worksheetName As String,
                                           ByVal cellAddress As String,
                                           ByVal changedValue As Object,
                                           ByVal dataFormat As String,
                                           ByVal description As String) As AbovoAppCls.AbovoTransaction
            If ExcelModels Is Nothing OrElse modelID < 0 OrElse modelID >= ExcelModels.Length OrElse
               ExcelModels(modelID) Is Nothing OrElse ExcelModels(modelID).ChangeManager Is Nothing Then
                Return New AbovoAppCls.AbovoTransaction With {
                    .BError = True,
                    .StrResponseMessage = "The model change manager is not available."}
            End If

            Dim workbook = ExcelModels(modelID).WB
            If workbook Is Nothing OrElse String.IsNullOrWhiteSpace(worksheetName) OrElse
               Not workbook.Worksheets.Contains(worksheetName) OrElse String.IsNullOrWhiteSpace(cellAddress) Then
                Return New AbovoAppCls.AbovoTransaction With {
                    .BError = True,
                    .StrResponseMessage = "The workbook target for this editor is not available."}
            End If

            Dim targetCell As Cell = workbook.Worksheets(worksheetName).Cells(cellAddress)
            Dim change As New DataChangeEvent With {
                .ModelID = modelID,
                .Description = description,
                .WSName = worksheetName,
                .CellAddress = cellAddress,
                .OriginalValue = targetCell.Value,
                .ChangedValue = changedValue,
                .DataFormat = dataFormat,
                .TimeStamp = Now(),
                .UserName = Environment.UserName}
            Return ExcelModels(modelID).ChangeManager.ProcessChange(change)
        End Function

        Friend Function EditorValueFromCell(ByVal cell As Cell,
                                            ByVal dataFormat As String) As Object
            If cell Is Nothing OrElse cell.Value.IsEmpty Then Return Nothing
            Select Case If(dataFormat, String.Empty).Trim().ToUpperInvariant()
                Case "D"
                    If cell.Value.IsDateTime Then Return cell.Value.DateTimeValue
                    If cell.Value.IsNumeric Then Return DateTime.FromOADate(cell.Value.NumericValue)
                    Return cell.DisplayText
                Case "B"
                    If cell.Value.IsBoolean Then Return cell.Value.BooleanValue
                    If cell.Value.IsNumeric Then Return cell.Value.NumericValue <> 0
                    Return cell.DisplayText
                Case "N", "C", "M", "SM", "R", "P", "I", "Y"
                    If cell.Value.IsNumeric Then Return cell.Value.NumericValue
                    Return Nothing
                Case Else
                    Return cell.DisplayText
            End Select
        End Function

        Friend Function TryProcessModelHistoryShortcut(ByVal owner As IWin32Window,
                                                       ByVal modelID As Integer,
                                                       ByVal keyData As Keys) As Boolean
            Dim redo As Boolean
            If keyData = (Keys.Control Or Keys.Z) Then
                redo = False
            ElseIf keyData = (Keys.Control Or Keys.Y) OrElse
                   keyData = (Keys.Control Or Keys.Shift Or Keys.Z) Then
                redo = True
            Else
                Return False
            End If

            If ExcelModels Is Nothing OrElse modelID < 0 OrElse modelID >= ExcelModels.Length OrElse
               ExcelModels(modelID) Is Nothing OrElse ExcelModels(modelID).ChangeManager Is Nothing Then
                Return False
            End If

            Dim result As AbovoAppCls.AbovoTransaction =
                If(redo, ExcelModels(modelID).ChangeManager.Redo(), ExcelModels(modelID).ChangeManager.Undo())
            If result IsNot Nothing AndAlso result.BError Then
                XtraMessageBox.Show(owner, result.StrResponseMessage, "Change History",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            Return True
        End Function

    End Module

    'A pending editor is not yet a workbook edit. Keep Save clickable for it so
    'the normal commit/validation path still works on the first click. This
    'binding observes state only: no posting, calculation or dirty writes.
    Friend NotInheritable Class ModelSaveButtonBinding
        Implements IDisposable

        Private ReadOnly Owner As Control
        Private ReadOnly Model As ExcelModel
        Private ReadOnly SaveButton As DevExpress.XtraBars.Docking2010.WindowsUIButton
        Private ReadOnly ObserveEditor As Boolean
        Private DisposedBinding As Boolean

        Friend Sub New(control As Control, modelInstance As ExcelModel,
                       panel As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel,
                       Optional includePendingEditor As Boolean = False)
            Owner = control
            Model = modelInstance
            ObserveEditor = includePendingEditor
            SaveButton = panel.Buttons.OfType(Of DevExpress.XtraBars.Docking2010.WindowsUIButton)().
                FirstOrDefault(Function(button) Convert.ToString(button.Tag) = "SaveBP")
            AddHandler Model.DirtyStateChanged, AddressOf StateChanged
            AddHandler Model.ManualSaveAvailabilityChanged, AddressOf StateChanged
            AddHandler Owner.Disposed, AddressOf OwnerDisposed
            If ObserveEditor Then AddHandler Application.Idle, AddressOf CheckPendingEditor
            RefreshState()
        End Sub

        Private Sub StateChanged(sender As Object, e As EventArgs)
            If DisposedBinding OrElse Owner.IsDisposed OrElse Owner.Disposing Then Return
            If Owner.InvokeRequired Then
                If Not Owner.IsHandleCreated Then Return
                Try
                    Owner.BeginInvoke(New MethodInvoker(AddressOf RefreshState))
                Catch ex As InvalidOperationException
                    'The UI handle may have gone away during model shutdown.
                End Try
            Else
                RefreshState()
            End If
        End Sub

        Private Sub CheckPendingEditor(sender As Object, e As EventArgs)
            If DisposedBinding OrElse Not Owner.Visible Then Return
            RefreshState()
        End Sub

        Private Sub RefreshState()
            If DisposedBinding OrElse Owner.IsDisposed OrElse Owner.Disposing OrElse SaveButton Is Nothing Then Return
            Dim canSave = Not Model.IsClosing AndAlso
                (Model.ManualSaveAvailable OrElse (ObserveEditor AndAlso HasPendingEditor()))
            If SaveButton.Enabled <> canSave Then SaveButton.Enabled = canSave
        End Sub

        Private Function HasPendingEditor() As Boolean
            'Walk only the focus branch, not every grid/cell in a large DIT.
            Dim focused As Control = Owner
            While focused IsNot Nothing AndAlso focused.ContainsFocus
                Dim editor = TryCast(focused, BaseEdit)
                If editor IsNot Nothing Then Return editor.IsModified AndAlso Not editor.Properties.ReadOnly
                focused = focused.Controls.Cast(Of Control)().FirstOrDefault(Function(child) child.ContainsFocus)
            End While
            Return False
        End Function

        Private Sub OwnerDisposed(sender As Object, e As EventArgs)
            Dispose()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If DisposedBinding Then Return
            DisposedBinding = True
            RemoveHandler Model.DirtyStateChanged, AddressOf StateChanged
            RemoveHandler Model.ManualSaveAvailabilityChanged, AddressOf StateChanged
            RemoveHandler Owner.Disposed, AddressOf OwnerDisposed
            If ObserveEditor Then RemoveHandler Application.Idle, AddressOf CheckPendingEditor
        End Sub
    End Class

    'Standalone model windows need the same workbook history shortcuts and
    'post-undo refresh as DIT, including when a native editor owns the key message.
    Friend NotInheritable Class ModelFormHistoryBinding
        Implements IMessageFilter, IDisposable

        Private ReadOnly Owner As Form
        Private ReadOnly ModelID As Integer
        Private ReadOnly Manager As ModelChangeManagerV2
        Private ReadOnly RefreshAction As Action
        Private NeedsRefresh As Boolean = True
        Private Refreshing As Boolean
        Private DisposedBinding As Boolean

        Friend Sub New(form As Form, id As Integer, refreshFromWorkbook As Action)
            Owner = form
            ModelID = id
            Manager = ExcelModels(id).ChangeManager
            RefreshAction = refreshFromWorkbook
            AddHandler Manager.HistoryChanged, AddressOf HistoryChanged
            AddHandler Owner.VisibleChanged, AddressOf RefreshWhenShown
            AddHandler Owner.Activated, AddressOf RefreshWhenShown
            AddHandler Owner.Disposed, AddressOf OwnerDisposed
            Application.AddMessageFilter(Me)
        End Sub

        Private Sub HistoryChanged(sender As Object, e As ChangeHistoryChangedEventArgsV2)
            NeedsRefresh = True
            'Ordinary posts already refresh their own editor. Do not rebuild a
            'grid inside CellValueChanged; refresh other windows on activation.
            If e.IsUndoRedo Then RefreshWhenShown(Me, EventArgs.Empty)
        End Sub

        Private Sub RefreshWhenShown(sender As Object, e As EventArgs)
            If DisposedBinding OrElse Refreshing OrElse
                Owner.IsDisposed OrElse Not Owner.Visible Then Return
            If BIsSaving OrElse ModelSafetyManager.IsBulkWorkbookMutationInProgress(ModelID) Then Return
            Dim model = ExcelModels(ModelID)
            If model Is Nothing OrElse model.IsClosing Then Return
            If Not NeedsRefresh AndAlso Not model.DeferredSaveResultsPending Then Return
            Refreshing = True
            Try
                'An already-open output form can be activated without passing
                'through FileInstance's Show command after a fast input save.
                model.EnsureDeferredSaveResultsCurrent("Refreshing " & Owner.Text & "...")
                RefreshAction()
                NeedsRefresh = False
            Finally
                Refreshing = False
            End Try
        End Sub

        Public Function PreFilterMessage(ByRef message As Message) As Boolean Implements IMessageFilter.PreFilterMessage
            If DisposedBinding OrElse Not Owner.Visible OrElse
                (message.Msg <> &H100 AndAlso message.Msg <> &H104) Then Return False
            Dim target As Control = Control.FromChildHandle(message.HWnd)
            If target Is Nothing OrElse (target IsNot Owner AndAlso Not Owner.Contains(target)) Then Return False
            Dim keys As Keys = CType(message.WParam.ToInt32(), Keys) Or Control.ModifierKeys
            Return ModelPostingChangeSupport.TryProcessModelHistoryShortcut(Owner, ModelID, keys)
        End Function

        Private Sub OwnerDisposed(sender As Object, e As EventArgs)
            Dispose()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If DisposedBinding Then Return
            DisposedBinding = True
            Application.RemoveMessageFilter(Me)
            RemoveHandler Manager.HistoryChanged, AddressOf HistoryChanged
            RemoveHandler Owner.VisibleChanged, AddressOf RefreshWhenShown
            RemoveHandler Owner.Activated, AddressOf RefreshWhenShown
            RemoveHandler Owner.Disposed, AddressOf OwnerDisposed
        End Sub
    End Class

    Friend NotInheritable Class ModelPostingHistoryBinding
        Private ReadOnly Owner As Control
        Private ReadOnly Manager As ModelChangeManagerV2
        Private ReadOnly WorksheetName As String
        Private ReadOnly RefreshAction As Action

        Friend Sub New(ByVal ownerControl As Control,
                       ByVal modelID As Integer,
                       ByVal worksheetName As String,
                       ByVal refreshFromWorkbook As Action)
            Owner = ownerControl
            WorksheetName = worksheetName
            RefreshAction = refreshFromWorkbook
            If ExcelModels Is Nothing OrElse modelID < 0 OrElse modelID >= ExcelModels.Length OrElse
               ExcelModels(modelID) Is Nothing Then Return
            Manager = ExcelModels(modelID).ChangeManager
            If Manager Is Nothing Then Return
            AddHandler Manager.HistoryChanged, AddressOf HistoryChanged
            AddHandler Owner.Disposed, AddressOf OwnerDisposed
        End Sub

        Private Sub HistoryChanged(ByVal sender As Object,
                                   ByVal e As ChangeHistoryChangedEventArgsV2)
            If Not e.IsUndoRedo OrElse Owner Is Nothing OrElse Owner.IsDisposed Then Return
            If e.WorksheetNames IsNot Nothing AndAlso
               Not e.WorksheetNames.Contains(WorksheetName, StringComparer.OrdinalIgnoreCase) Then Return
            If Owner.IsHandleCreated Then
                Owner.BeginInvoke(New MethodInvoker(Sub()
                    If Not Owner.IsDisposed Then RefreshAction()
                End Sub))
            End If
        End Sub

        Private Sub OwnerDisposed(ByVal sender As Object, ByVal e As EventArgs)
            If Manager IsNot Nothing Then RemoveHandler Manager.HistoryChanged, AddressOf HistoryChanged
            If Owner IsNot Nothing Then RemoveHandler Owner.Disposed, AddressOf OwnerDisposed
        End Sub
    End Class

End Namespace
