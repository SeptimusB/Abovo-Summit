Imports Abovo.FileManager
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraEditors
Imports System.Windows.Forms

Namespace Abovo

    Friend Module ModelPostingChangeSupport

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
