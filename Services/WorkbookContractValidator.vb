Imports Abovo.AbovoAppCls
Imports DevExpress.Spreadsheet

Namespace Abovo

    Public NotInheritable Class WorkbookContractValidator

        Private Sub New()
        End Sub

        Public Shared Function Validate(ByVal Workbook As IWorkbook) As AbovoTransaction

            Dim Profile As WorkbookModelProfile = Nothing
            Return Validate(Workbook, Profile)

        End Function

        Public Shared Function Validate(
            ByVal Workbook As IWorkbook,
            ByRef Profile As WorkbookModelProfile) As AbovoTransaction

            If Workbook Is Nothing Then
                Return Failure("The workbook could not be loaded.")
            End If

            Dim FailureMessage As String = String.Empty
            Profile = WorkbookModelProfileRegistry.Resolve(Workbook, FailureMessage)

            If Profile Is Nothing Then Return Failure(FailureMessage)

            Return New AbovoTransaction With {
                .BError = False,
                .BSuccess = True,
                .StringReturn = Profile.ModelType,
                .StrResponseMessage = "Workbook contract validated."
            }

        End Function

        Private Shared Function Failure(ByVal Message As String) As AbovoTransaction

            Return New AbovoTransaction With {
                .BError = True,
                .BSuccess = False,
                .IntReturnCode = -1,
                .StringReturn = Message,
                .StrResponseMessage = Message
            }

        End Function

    End Class

End Namespace
