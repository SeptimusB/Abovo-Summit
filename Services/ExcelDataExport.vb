Imports Abovo
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraGrid.Views.Grid
Imports Abovo.LogDebugDev

Namespace Abovo
    Public Class ExcelDataExport

        Public CreatedWorkbook As IWorkbook
        Public CreatedWorksheets(-1) As Worksheet
        Public WorksheetCount As Integer = 0

        Public Sub New()

            Dim SSViewer As New SpreadsheetViewer

            CreatedWorkbook = SSViewer.SpreadsheetControlViewer.Document
            CreatedWorksheets(0) = CreatedWorkbook.Worksheets(0)
            WorksheetCount += 1

        End Sub
        Public Function CreateGridWorksheet(ByVal GV As GridView, ByVal WsOptions As ExcelExportAdditions) As Boolean

            CreatedWorkbook.Worksheets.Insert(0, WsOptions.SheetName)

            Dim ws As Worksheet = CreatedWorkbook.Worksheets(0)


            Dim RowCount As Integer = GV.RowCount
            Dim DataRowCount As Integer = GV.DataRowCount
            Dim GroupRowCOunt As Integer = RowCount - DataRowCount
            Dim GColumnCount As Integer = GV.Columns.Count
            Dim ActiveRow As Integer = 0
            Dim ActiveColumn As Integer = 0
            Dim NumberFormatString As String = "#,###;[red](#,###);0"
            SystemLog("Row Count: " & RowCount)
            SystemLog("Data Row Count: " & DataRowCount)
            SystemLog("Group Row Count: " & GroupRowCOunt)

            Dim WSCell As Cell = ws.Cells(ActiveRow, ActiveColumn)
            WSCell.Value = "Exported Data " & Now().ToString

            WSCell.Font.Bold = True
            WSCell.Font.Size = 14
            WSCell.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left

            WSCell.NumberFormat = NumberFormatString

            ActiveRow += 1
            Dim I As Integer

            For I = 0 To DataRowCount - 1

                'Dim CellValue As Object = GV.GetRowCellValue(I, Col)
                'Dim NewValue As Double = Convert.ToDouble(CellValue) * 0.9

            Next

            Return True

        End Function

        Structure ExcelExportAdditions

            Public SheetName As String
            Public Title As String
            Public DoNotes As Boolean
            Public DoTotals As Boolean
            Public DoSubTotals As Boolean
            Public DoGroupTotals As Boolean
            Public DoGroupSubTotals As Boolean
            Public NoteText As String
            Public ColumnFormats() As ColumnFormat


        End Structure

        Structure ColumnFormat
            Public Caption As String
            Public ColumnIndex As Integer
            Public ColumnFormat As String
        End Structure

    End Class
    Public Class GridExportPackage

        Public GridView As GridView
        Public Description As String
        Public IsDefault As Boolean
        Public IDStart As Integer
        Public IDCount As Integer
        Public GroupA As Integer
        Public GroupB As Integer
        Public GroupC As Integer

        Public GroupD As Integer


    End Class
End Namespace