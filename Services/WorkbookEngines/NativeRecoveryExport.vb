Option Strict Off

Imports System.IO
Imports DevExpress.Spreadsheet

Namespace Abovo.WorkbookEngines
    Friend Interface IWorkbookRecoveryExportBackend
        Sub ExportRecoveryCopy(path As String, source As String)
    End Interface

    Partial Friend NotInheritable Class DevExpressCalculationBackend
        Implements IWorkbookRecoveryExportBackend

        Public Sub ExportRecoveryCopy(path As String, source As String) Implements IWorkbookRecoveryExportBackend.ExportRecoveryCopy
            RequireOwner()
            If Not String.Equals(IO.Path.GetExtension(path), ".xlsm", StringComparison.OrdinalIgnoreCase) Then Throw New ArgumentException("Recovery output must be XLSM.")
            Dim properties = book.DocumentProperties.Custom
            Dim names = {RecoveryBackupStore.SourceProperty, RecoveryBackupStore.VersionProperty,
                         RecoveryBackupStore.DateProperty, RecoveryBackupStore.PendingProperty}
            Dim original = names.ToDictionary(Function(name) name, Function(name) properties(name))
            Try
                properties(names(0)) = IO.Path.GetFullPath(source)
                properties(names(1)) = "1"
                properties(names(2)) = DateTime.UtcNow.ToString("O")
                properties(names(3)) = False 'Only a current calculated revision is exported.
                ExportCopy(path)
            Finally
                For Each pair In original
                    properties(pair.Key) = pair.Value
                Next
            End Try
        End Sub
    End Class

    Partial Friend NotInheritable Class ExcelCalculationBackend
        Implements IWorkbookRecoveryExportBackend

        Public Sub ExportRecoveryCopy(path As String, source As String) Implements IWorkbookRecoveryExportBackend.ExportRecoveryCopy
            RequireOwner()
            If Not String.Equals(IO.Path.GetExtension(path), ".xlsm", StringComparison.OrdinalIgnoreCase) Then Throw New ArgumentException("Recovery output must be XLSM.")
            If File.Exists(path) Then Throw New IOException("Recovery candidate already exists.")
            'SaveCopyAs cannot convert formats. Convert a private same-format
            'copy with macros/events/calculation disabled, never the live owner.
            Dim temporary = IO.Path.Combine(IO.Path.GetDirectoryName(path), "~conversion-" & Guid.NewGuid().ToString("N") & IO.Path.GetExtension(source))
            Dim originalPath = CStr(book.FullName), originalSaved = CBool(book.Saved), identity = SheetIdentity()
            Dim copy As Object = Nothing, properties As Object = Nothing
            Try
                ExportCopy(temporary)
                CopyInternetZone(temporary, ReadInternetZone(source))
                app.AutomationSecurity = 3 : app.EnableEvents = False : app.DisplayAlerts = False
                app.Calculation = -4135 : app.CalculateBeforeSave = False
                copy = books.Open(temporary, 0, True)
                properties = copy.CustomDocumentProperties
                PutRecoveryProperty(properties, RecoveryBackupStore.SourceProperty, IO.Path.GetFullPath(source), 4)
                PutRecoveryProperty(properties, RecoveryBackupStore.VersionProperty, "1", 4)
                PutRecoveryProperty(properties, RecoveryBackupStore.DateProperty, DateTime.UtcNow.ToString("O"), 4)
                PutRecoveryProperty(properties, RecoveryBackupStore.PendingProperty, False, 2)
                copy.SaveAs(path, 52) 'xlOpenXMLWorkbookMacroEnabled
                If CInt(copy.FileFormat) <> 52 OrElse Not String.Equals(CStr(copy.FullName), path, StringComparison.OrdinalIgnoreCase) Then
                    Throw New InvalidDataException("Excel did not write the requested recovery format/path.")
                End If
                If CStr(book.FullName) <> originalPath OrElse CBool(book.Saved) <> originalSaved OrElse SheetIdentity() <> identity Then
                    Throw New InvalidOperationException("Recovery export changed the live workbook identity or saved state.")
                End If
            Finally
                Release(properties)
                Try
                    If copy IsNot Nothing Then copy.Close(False)
                Finally
                    Release(copy)
                    If File.Exists(temporary) Then File.Delete(temporary) 'Exact disposable conversion copy only.
                End Try
            End Try
        End Sub

        Private Shared Sub PutRecoveryProperty(properties As Object, name As String, value As Object, kind As Integer)
            Dim item As Object = Nothing
            Try
                'Do not catch arbitrary COM errors as if the property was absent.
                For index As Integer = 1 To CInt(properties.Count)
                    item = properties.Item(index)
                    If String.Equals(CStr(item.Name), name, StringComparison.Ordinal) Then
                        item.Value = value
                        Return
                    End If
                    Release(item)
                Next
                item = properties.Add(name, False, kind, value)
            Finally
                Release(item)
            End Try
        End Sub
    End Class
End Namespace
