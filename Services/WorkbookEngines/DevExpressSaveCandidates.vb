Option Strict On

Imports System.IO
Imports DevExpress.Spreadsheet

Namespace Abovo.WorkbookEngines
    Partial Friend NotInheritable Class DevExpressCalculationBackend
        Implements IWorkbookCandidateBackend

        Public Sub ExportCopy(path As String) Implements IWorkbookCandidateBackend.ExportCopy
            RequireOwner()
            Dim format = NativeFormat(path)
            Dim originalPath = book.Path, modified = book.DocumentProperties.Modified, author = book.DocumentProperties.LastModifiedBy
            Dim identity = SheetIdentity()
            Try
                Using output As New FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None)
                    If format = DocumentFormat.Xlsb Then
                        ' The established wrapper restores its temporary formula
                        ' rewrites when False is returned. A candidate must not
                        ' silently commit those rewrites to the live session.
                        WorkbookXlsbFormulaCompatibility.Save(book, Function()
                            book.SaveDocument(output, format)
                            Return False
                        End Function)
                    Else
                        book.SaveDocument(output, format)
                    End If
                    output.Flush(True)
                End Using
                If book.Path <> originalPath Then Throw New InvalidOperationException("Candidate export changed the open workbook path.")
                If SheetIdentity() <> identity Then Throw New InvalidOperationException("Candidate export changed the open worksheet identities.")
            Finally
                book.DocumentProperties.Modified = modified
                book.DocumentProperties.LastModifiedBy = author
            End Try
        End Sub

        Private Function SheetIdentity() As String
            If book.ChartSheets.Count <> 0 Then Throw New NotSupportedException("Chart-sheet VBA identity preservation is outside the initial candidate trial.")
            Return String.Join(vbLf, book.Worksheets.Select(Function(s) s.Name & ":" & s.CodeName & ":" & s.IsProtected.ToString()))
        End Function

        Private Shared Function NativeFormat(path As String) As DocumentFormat
            Select Case IO.Path.GetExtension(path).ToLowerInvariant()
                Case ".xlsb" : Return DocumentFormat.Xlsb
                Case ".xlsm" : Return DocumentFormat.Xlsm
                Case ".xlsx" : Return DocumentFormat.Xlsx
                Case Else : Throw New ArgumentException("Unsupported candidate format.")
            End Select
        End Function

        Public Function ReadCopy(path As String, areas As IList(Of WorkbookReadArea), cells As IList(Of WorkbookReadArea)) As WorkbookCandidateReadback Implements IWorkbookCandidateBackend.ReadCopy
            RequireOwner()
            Dim original = book
            Dim identity = SheetIdentity()
            Using copy As New Workbook()
                copy.Options.CalculationMode = WorkbookCalculationMode.Manual
                Using input = File.OpenRead(path)
                    If Not copy.LoadDocument(input, NativeFormat(path)) Then Throw New InvalidDataException("Candidate could not be reopened.")
                End Using
                copy.Options.CalculationMode = WorkbookCalculationMode.Manual
                Try
                    book = copy
                    If SheetIdentity() <> identity Then Throw New InvalidDataException("Candidate worksheet order, name, code name or protection changed.")
                    Return New WorkbookCandidateReadback(areas.Select(Function(a) Read(a)).ToList(), cells.Select(Function(a) ReadCell(a)).ToList())
                Finally
                    book = original
                End Try
            End Using
        End Function
    End Class
End Namespace
