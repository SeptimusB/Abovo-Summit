
Imports Abovo.FileManager
Imports DevExpress.CodeParser



Namespace Abovo
    Public Class ModelPostingTextBox

        Inherits DevExpress.XtraEditors.TextEdit

        Private ModelID As Integer
        Private TargetWorksheet As String
        Private TargetCell As String
        Private PriorVal As Object

        Sub New()



        End Sub
        Sub Initialise(ByRef SetModelID As Integer, ByRef SetTargetWorksheet As String, ByRef SetTargetCell As String)

            ModelID = SetModelID
            TargetWorksheet = SetTargetWorksheet
            TargetCell = SetTargetCell

            Try
                EditValue = ExcelModels(ModelID).WB.Worksheets(TargetWorksheet).Cells(TargetCell).DisplayText

            Catch ex As Exception

                EditValue = ""

            End Try

            PriorVal = EditValue

            AddHandler MyBase.EditValueChanged, AddressOf ProcessChange

        End Sub
        Protected Sub ProcessChange(ByVal sender As Object, ByVal e As System.EventArgs)

            Dim NewVal As String = EditValue.ToString

            If EditValue IsNot Nothing Then

                ExcelModels(ModelID).WB.Worksheets(TargetWorksheet).Cells(TargetCell).Value = NewVal
                PriorVal = EditValue

            End If

        End Sub


    End Class
End Namespace
