Imports Abovo.FileManager
Imports DevExpress.XtraEditors
Namespace Abovo
    Public Class ModelPostingDateBox

        Inherits DateEdit

        Private ModelID As Integer
        Private TargetWorksheet As String
        Private TargetCell As String
        Private SuppressPosting As Boolean
        Private HistoryBinding As ModelPostingHistoryBinding

        Sub New()



        End Sub
        Sub Initialise(ByRef SetModelID As Integer, ByRef SetTargetWorksheet As String, ByRef SetTargetCell As String)

            ModelID = SetModelID
            TargetWorksheet = SetTargetWorksheet
            TargetCell = SetTargetCell

            RefreshFromWorkbook()
            RemoveHandler MyBase.Validated, AddressOf ProcessChange
            AddHandler MyBase.Validated, AddressOf ProcessChange
            HistoryBinding = New ModelPostingHistoryBinding(Me, ModelID, TargetWorksheet, AddressOf RefreshFromWorkbook)

        End Sub
        Protected Sub ProcessChange(ByVal sender As Object, ByVal e As System.EventArgs)

            If SuppressPosting Then Return
            Dim result = PostModelCellValue(ModelID, TargetWorksheet, TargetCell,
                                            If(EditValue Is Nothing OrElse Convert.IsDBNull(EditValue), Nothing, EditValue),
                                            "D", "Date value updated")
            If result.BError Then RefreshFromWorkbook()

        End Sub

        Private Sub RefreshFromWorkbook()
            SuppressPosting = True
            Try
                Dim cell = ExcelModels(ModelID).WB.Worksheets(TargetWorksheet).Cells(TargetCell)
                EditValue = EditorValueFromCell(cell, "D")
            Finally
                SuppressPosting = False
            End Try
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message,
                                                   ByVal keyData As Keys) As Boolean
            If TryProcessModelHistoryShortcut(Me, ModelID, keyData) Then Return True
            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

    End Class

End Namespace
