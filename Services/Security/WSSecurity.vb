Imports Abovo
Imports Abovo.FileManager
Namespace Abovo


    Public Class WSSecurity

        Public Shared Sub UNProtectWS(ModelID As Integer, WSName As String)

            If Not ExcelModels(ModelID).WB.Worksheets(WSName).IsProtected Then Exit Sub

            Dim pwd As String = ExcelModels(ModelID).WBStructure.RejData
            ExcelModels(ModelID).WB.Worksheets(WSName).Unprotect(pwd)
            If ExcelModels(ModelID).WB.Worksheets(WSName).IsProtected Then
                Throw New InvalidOperationException(
                    "Worksheet '" & WSName & "' could not be unprotected.")
            End If


        End Sub

        Public Shared Sub ProtectWS(ModelID As Integer, WSName As String)

            If ExcelModels(ModelID).WB.Worksheets(WSName).IsProtected Then Exit Sub

            Dim pwd As String = ExcelModels(ModelID).WBStructure.RejData

            ExcelModels(ModelID).WB.Worksheets(WSName).Protect(
                pwd,
                DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default)
            If Not ExcelModels(ModelID).WB.Worksheets(WSName).IsProtected Then
                Throw New InvalidOperationException(
                    "Worksheet '" & WSName & "' could not be protected.")
            End If


        End Sub

    End Class
End Namespace
