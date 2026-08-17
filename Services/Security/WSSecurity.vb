Imports Abovo
Imports Abovo.FileManager
Namespace Abovo


    Public Class WSSecurity

        Public Shared Sub UNProtectWS(ModelID As Integer, WSName As String)

            If Not ExcelModels(ModelID).WB.Worksheets(WSName).IsProtected Then Exit Sub

            Dim pwd As String = ExcelModels(ModelID).WBStructure.RejData
            Try
                ExcelModels(ModelID).WB.Worksheets(WSName).Unprotect(pwd)
            Catch ex As Exception

            End Try


        End Sub

        Public Shared Sub ProtectWS(ModelID As Integer, WSName As String)

            If ExcelModels(ModelID).WB.Worksheets(WSName).IsProtected Then Exit Sub

            Dim pwd As String = ExcelModels(ModelID).WBStructure.RejData

            Try
                ExcelModels(ModelID).WB.Worksheets(WSName).Protect(pwd, DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default)
            Catch ex As Exception

            End Try


        End Sub

    End Class
End Namespace
