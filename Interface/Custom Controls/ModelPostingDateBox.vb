Imports Abovo.FileManager
Imports DevExpress.XtraEditors
Namespace Abovo
    Public Class ModelPostingDateBox

        Inherits DateEdit

        Private ModelID As Integer
        Private TargetWorksheet As String
        Private TargetCell As String

        Sub New()



        End Sub
        Sub Initialise(ByRef SetModelID As Integer, ByRef SetTargetWorksheet As String, ByRef SetTargetCell As String)

            ModelID = SetModelID
            TargetWorksheet = SetTargetWorksheet
            TargetCell = SetTargetCell

            Try

                EditValue = DateTime.FromOADate(ExcelModels(ModelID).WB.Worksheets(TargetWorksheet).Cells(TargetCell).Value.NumericValue)

            Catch ex As Exception

                EditValue = ""

            End Try

            AddHandler MyBase.EditValueChanged, AddressOf ProcessChange

        End Sub
        Protected Sub ProcessChange(ByVal sender As Object, ByVal e As System.EventArgs)

            Dim NewVal As String = EditValue

            If EditValue IsNot Nothing Then

                ExcelModels(ModelID).WB.Worksheets(TargetWorksheet).Cells(TargetCell).Value = NewVal

            End If

        End Sub

    End Class

End Namespace
