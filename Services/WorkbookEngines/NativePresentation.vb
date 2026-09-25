Option Strict Off

Imports System.Drawing
Imports DevExpress.Spreadsheet

Namespace Abovo.WorkbookEngines
    Partial Friend NotInheritable Class DevExpressCalculationBackend
        Implements IWorkbookPresentationBackend

        Public Function ReadPresentation(values As WorkbookValueBlock) As WorkbookPresentationBlock Implements IWorkbookPresentationBackend.ReadPresentation
            RequireOwner()
            Dim area = values.Area
            Dim sheet = book.Worksheets(area.Worksheet)
            Dim cells(area.Rows - 1, area.Columns - 1) As WorkbookPresentationCell
            For r As Integer = 0 To area.Rows - 1
                For c As Integer = 0 To area.Columns - 1
                    Dim cell = sheet.Cells(area.Row + r, area.Column + c)
                    Dim font = cell.Font
                    Dim style = System.Drawing.FontStyle.Regular
                    If font.Bold Then style = style Or System.Drawing.FontStyle.Bold
                    If font.Italic Then style = style Or System.Drawing.FontStyle.Italic
                    If font.Strikethrough Then style = style Or System.Drawing.FontStyle.Strikeout
                    If font.UnderlineType <> UnderlineType.None Then style = style Or System.Drawing.FontStyle.Underline
                    Dim appearance As New WorkbookCellAppearance(cell.NumberFormat, cell.Fill.BackgroundColor, font.Color,
                        font.Name, font.Size, style, cell.Alignment.Horizontal, cell.Alignment.Vertical,
                        cell.Alignment.WrapText, cell.Alignment.Indent, cell.Alignment.RotationAngle,
                        cell.Protection.Locked, cell.Fill.PatternType = PatternType.Solid)
                    cells(r, c) = New WorkbookPresentationCell(appearance, cell.DisplayText)
                Next
            Next
            Return New WorkbookPresentationBlock(area, cells)
        End Function
    End Class

    Partial Friend NotInheritable Class ExcelCalculationBackend
        Implements IWorkbookPresentationBackend

        Public Function ReadPresentation(values As WorkbookValueBlock) As WorkbookPresentationBlock Implements IWorkbookPresentationBackend.ReadPresentation
            RequireOwner()
            Dim area = values.Area
            Dim cells(area.Rows - 1, area.Columns - 1) As WorkbookPresentationCell
            Dim sheets As Object = Nothing, sheet As Object = Nothing, range As Object = Nothing, rangeCells As Object = Nothing
            Dim cell As Object = Nothing, display As Object = Nothing, font As Object = Nothing, fill As Object = Nothing
            Try
                sheets = book.Worksheets : sheet = sheets.Item(area.Worksheet)
                range = sheet.Range(area.Address) : rangeCells = range.Cells
                Using formatter As New WorkbookDisplayFormatter(CBool(book.Date1904))
                    For r As Integer = 0 To area.Rows - 1
                        For c As Integer = 0 To area.Columns - 1
                            Try
                                cell = rangeCells.Item(r + 1, c + 1)
                                display = cell.DisplayFormat : font = display.Font : fill = display.Interior
                                Dim style = System.Drawing.FontStyle.Regular
                                If CBool(font.Bold) Then style = style Or System.Drawing.FontStyle.Bold
                                If CBool(font.Italic) Then style = style Or System.Drawing.FontStyle.Italic
                                If CBool(font.Strikethrough) Then style = style Or System.Drawing.FontStyle.Strikeout
                                If CInt(font.Underline) <> -4142 Then style = style Or System.Drawing.FontStyle.Underline
                                Dim background = If(CInt(fill.ColorIndex) = -4142, Color.Empty, ColorTranslator.FromOle(CInt(fill.Color)))
                                Dim numberFormat = CStr(display.NumberFormat)
                                Dim appearance As New WorkbookCellAppearance(numberFormat, background, ColorTranslator.FromOle(CInt(font.Color)),
                                    CStr(font.Name), CDbl(font.Size), style, HorizontalAlignment(CInt(display.HorizontalAlignment)),
                                    VerticalAlignment(CInt(display.VerticalAlignment)), CBool(display.WrapText), CInt(display.IndentLevel),
                                    CInt(display.Orientation), CBool(cell.Locked), CInt(fill.Pattern) = 1)
                                cells(r, c) = New WorkbookPresentationCell(appearance, formatter.Format(values.ValueAt(r, c), numberFormat))
                            Finally
                                Release(fill) : Release(font) : Release(display) : Release(cell)
                            End Try
                        Next
                    Next
                End Using
                Return New WorkbookPresentationBlock(area, cells)
            Finally
                Release(rangeCells) : Release(range) : Release(sheet) : Release(sheets)
            End Try
        End Function

        Private Shared Function HorizontalAlignment(value As Integer) As SpreadsheetHorizontalAlignment
            Select Case value
                Case 1 : Return SpreadsheetHorizontalAlignment.General
                Case -4131 : Return SpreadsheetHorizontalAlignment.Left
                Case -4152 : Return SpreadsheetHorizontalAlignment.Right
                Case -4108 : Return SpreadsheetHorizontalAlignment.Center
                Case 5 : Return SpreadsheetHorizontalAlignment.Fill
                Case -4130 : Return SpreadsheetHorizontalAlignment.Justify
                Case -4117 : Return SpreadsheetHorizontalAlignment.Distributed
                Case 7 : Return SpreadsheetHorizontalAlignment.CenterContinuous
                Case Else : Throw New NotSupportedException("Unsupported workbook horizontal alignment: " & value.ToString())
            End Select
        End Function
        Private Shared Function VerticalAlignment(value As Integer) As SpreadsheetVerticalAlignment
            Select Case value
                Case -4160 : Return SpreadsheetVerticalAlignment.Top
                Case -4107 : Return SpreadsheetVerticalAlignment.Bottom
                Case -4108 : Return SpreadsheetVerticalAlignment.Center
                Case -4130 : Return SpreadsheetVerticalAlignment.Justify
                Case -4117 : Return SpreadsheetVerticalAlignment.Distributed
                Case Else : Throw New NotSupportedException("Unsupported workbook vertical alignment: " & value.ToString())
            End Select
        End Function
    End Class
End Namespace
