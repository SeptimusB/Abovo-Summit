Option Strict Off

Imports System.IO

Namespace Abovo.WorkbookEngines
    Partial Friend NotInheritable Class ExcelCalculationBackend
        Implements IWorkbookCandidateBackend

        Friend Shared Sub CopyInternetZone(path As String, zone As String)
            If zone Is Nothing Then
                If ReadInternetZone(path) IsNot Nothing Then Throw New InvalidDataException("Unexpected security marking on candidate.")
                Return
            End If
            Using handle = OpenNativeFile(path & ":Zone.Identifier", &H40000000UI, 0UI, IntPtr.Zero, 2UI, 0UI, IntPtr.Zero)
                If handle.IsInvalid Then Throw New ComponentModel.Win32Exception(Runtime.InteropServices.Marshal.GetLastWin32Error(), "Candidate security marking could not be preserved.")
                Using stream As New FileStream(handle, FileAccess.Write), writer As New StreamWriter(stream)
                    writer.Write(zone)
                    writer.Flush()
                    stream.Flush(True)
                End Using
            End Using
        End Sub

        Public Sub ExportCopy(path As String) Implements IWorkbookCandidateBackend.ExportCopy
            RequireOwner()
            RequireExclusiveWorkbooks()
            If File.Exists(path) Then Throw New IOException("Candidate path already exists.")
            Dim originalPath = CStr(book.FullName), saved = CBool(book.Saved)
            Dim identity = SheetIdentity()
            Dim calculation = CInt(app.Calculation), beforeSave = CBool(app.CalculateBeforeSave)
            Try
                app.EnableEvents = False : app.DisplayAlerts = False
                app.Calculation = -4135 : app.CalculateBeforeSave = False
                book.SaveCopyAs(path)
                If CStr(book.FullName) <> originalPath OrElse CBool(book.Saved) <> saved Then Throw New InvalidOperationException("Candidate export changed the open workbook identity or saved state.")
                If SheetIdentity() <> identity Then Throw New InvalidOperationException("Candidate export changed the open worksheet identities.")
            Finally
                app.CalculateBeforeSave = beforeSave : app.Calculation = calculation
            End Try
        End Sub

        Private Function SheetIdentity() As String
            Dim sheets As Object = Nothing, allSheets As Object = Nothing, sheet As Object = Nothing
            Try
                sheets = book.Worksheets : allSheets = book.Sheets
                If CInt(sheets.Count) <> CInt(allSheets.Count) Then Throw New NotSupportedException("Non-worksheet VBA identity preservation is outside the initial candidate trial.")
                Dim identities As New List(Of String)()
                For i = 1 To CInt(sheets.Count)
                    Try
                        sheet = sheets.Item(i)
                        identities.Add(CStr(sheet.Name) & ":" & CStr(sheet.CodeName) & ":" & CBool(sheet.ProtectContents).ToString())
                    Finally
                        Release(sheet)
                    End Try
                Next
                Return String.Join(vbLf, identities)
            Finally
                Release(sheets) : Release(allSheets)
            End Try
        End Function

        Public Function ReadCopy(path As String, areas As IList(Of WorkbookReadArea), cells As IList(Of WorkbookReadArea)) As WorkbookCandidateReadback Implements IWorkbookCandidateBackend.ReadCopy
            RequireOwner()
            RequireExclusiveWorkbooks()
            Dim original = book, copy As Object = Nothing
            Dim identity = SheetIdentity()
            Try
                ' Already approved source; candidate is never a new macro trust
                ' decision. No VBA, events, links or recalculation on verification.
                app.AutomationSecurity = 3 : app.EnableEvents = False : app.DisplayAlerts = False
                app.Calculation = -4135 : app.CalculateBeforeSave = False
                copy = books.Open(path, 0, True)
                If Not CBool(copy.ReadOnly) Then Throw New InvalidOperationException("Candidate was not reopened read-only.")
                book = copy
                If SheetIdentity() <> identity Then Throw New InvalidDataException("Candidate worksheet order, name, code name or protection changed.")
                Return New WorkbookCandidateReadback(areas.Select(Function(a) Read(a)).ToList(), cells.Select(Function(a) ReadCell(a)).ToList())
            Finally
                book = original
                Try
                    If copy IsNot Nothing Then copy.Close(False)
                Finally
                    Release(copy)
                End Try
            End Try
        End Function
    End Class
End Namespace
