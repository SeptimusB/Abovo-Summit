Namespace Abovo
    Public Class ExportServices
        Public Shared Exporter As ExportForm
        Public Shared AmReady As Boolean
        Public Shared Sub SetExportMode(Mode As String)
            If Exporter Is Nothing Then Exporter = New ExportForm
            Exporter.SetMode(Mode)
        End Sub
        Public Shared Sub OpenFileInExcel(path As String)

            Dim startexternal As New Process()

            startexternal.StartInfo.FileName = path

            startexternal.StartInfo.UseShellExecute = True

            startexternal.Start()


            'Dim excelProcess As Process = New Process()

            '' Specify the path to Excel executable

            'excelProcess.StartInfo.FileName = "excel.exe " & path

            '' Start Excel

            'excelProcess.Start()
            'Try

            '    If Not String.IsNullOrEmpty(path) Then
            '        Dim excelApp As Object = CreateObject("Excel.Application")
            '        excelApp.Visible = True
            '        excelApp.Workbooks.Open(path)
            '    Else
            '        Throw New ArgumentException("File path cannot be empty.")
            '    End If

            'Catch ex As Exception

            '    MsgBox("Sorry, there was an error opening the Excel file: " & ex.Message, MsgBoxStyle.Critical, "Error")
            'End Try
        End Sub

    End Class

End Namespace

