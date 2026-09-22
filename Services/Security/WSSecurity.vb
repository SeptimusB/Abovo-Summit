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

        Public Shared Sub ProtectWS(ModelID As Integer, WSName As String,
                                    Optional permissions As DevExpress.Spreadsheet.WorksheetProtectionPermissions = DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default)

            If ExcelModels(ModelID).WB.Worksheets(WSName).IsProtected Then Exit Sub

            Dim pwd As String = ExcelModels(ModelID).WBStructure.RejData

            ProtectWorksheetForEditing(ExcelModels(ModelID).WB,
                                       ExcelModels(ModelID).WB.Worksheets(WSName), pwd, permissions)
            If Not ExcelModels(ModelID).WB.Worksheets(WSName).IsProtected Then
                Throw New InvalidOperationException(
                    "Worksheet '" & WSName & "' could not be protected.")
            End If


        End Sub

        Friend Shared Sub ProtectWorksheetForEditing(WB As DevExpress.Spreadsheet.IWorkbook,
                                                      WS As DevExpress.Spreadsheet.Worksheet,
                                                      Password As String,
                                                      Permissions As DevExpress.Spreadsheet.WorksheetProtectionPermissions)
            'Worksheet protection prevents accidental edits; it is not file encryption.
            'Use Excel-compatible legacy verification for sheets Summit re-protects.
            'Repeated strong password hashing adds avoidable structural-command latency.
            'Keep the password, cell locks and caller's permissions unchanged, and do
            'not change the workbook-wide policy for any subsequent protection calls.
            Dim PreviousStrongVerifier = WB.Options.Protection.UseStrongPasswordVerifier
            Try
                WB.Options.Protection.UseStrongPasswordVerifier = False
                WS.Protect(Password, Permissions)
            Finally
                WB.Options.Protection.UseStrongPasswordVerifier = PreviousStrongVerifier
            End Try
        End Sub

    End Class
End Namespace
