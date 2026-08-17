Imports Abovo.AbovoAppCls
Imports DevExpress.Spreadsheet

Namespace Abovo

    Public NotInheritable Class WorkbookContractValidator

        Private Const GlobalAssumptionsSheetName As String = "Global Assumptions"
        Private Const TransactionalDatabaseSheetName As String = "Transactional DB"
        Private Const ValidationMarkerAddress As String = "A8"
        Private Const ValidationMarkerText As String = "Business Plan Start Date"

        Private Sub New()
        End Sub

        Public Shared Function Validate(ByVal Workbook As IWorkbook) As AbovoTransaction

            If Workbook Is Nothing Then
                Return Failure("The workbook could not be loaded.")
            End If

            If Not Workbook.Worksheets.Contains(GlobalAssumptionsSheetName) Then
                Return Failure("The workbook is missing the 'Global Assumptions' worksheet.")
            End If

            Dim GlobalAssumptions As Worksheet =
                Workbook.Worksheets(GlobalAssumptionsSheetName)

            If Not String.Equals(GlobalAssumptions.Cells(ValidationMarkerAddress).DisplayText,
                                 ValidationMarkerText,
                                 StringComparison.Ordinal) Then

                Return Failure(
                    "The workbook does not contain the expected Abovo validation marker at " &
                    GlobalAssumptionsSheetName & "!" & ValidationMarkerAddress & ".")

            End If

            If Not Workbook.Worksheets.Contains(TransactionalDatabaseSheetName) Then
                Return Failure("The workbook is missing the 'Transactional DB' worksheet.")
            End If

            Return New AbovoTransaction With {
                .BError = False,
                .BSuccess = True,
                .StringReturn = "AbovoBP",
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
