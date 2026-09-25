Option Strict Off

Namespace Abovo.WorkbookEngines
    Partial Friend NotInheritable Class ExcelCalculationBackend
        Implements IWorkbookValueEditBackend

        Public Function ReadCell(area As WorkbookReadArea) As WorkbookCellState Implements IWorkbookValueEditBackend.ReadCell
            RequireOwner()
            RequireSingleCell(area)
            Dim sheets As Object = Nothing, sheet As Object = Nothing, cell As Object = Nothing, fill As Object = Nothing
            Dim display As Object = Nothing, conditions As Object = Nothing
            Try
                sheets = book.Worksheets : sheet = sheets.Item(area.Worksheet) : cell = sheet.Range(area.Address)
                ' Match DevExpress Cell.Fill: it includes active conditional
                ' formatting, unlike Excel Range.Interior (the base format).
                display = cell.DisplayFormat : fill = display.Interior : conditions = cell.FormatConditions
                Return New WorkbookCellState(ConvertValue(cell.Value2), If(CBool(cell.HasFormula), CStr(cell.Formula2), ""),
                    CStr(display.NumberFormat), CBool(cell.Locked), CInt(fill.Pattern) = 1,
                    CBool(cell.HasArray) OrElse CBool(cell.HasSpill), CBool(cell.MergeCells), CBool(sheet.ProtectContents), CInt(conditions.Count) > 0)
            Finally
                Release(conditions) : Release(fill) : Release(display) : Release(cell) : Release(sheet) : Release(sheets)
            End Try
        End Function

        Private Shared Sub RequireSingleCell(area As WorkbookReadArea)
            If area Is Nothing OrElse area.Rows <> 1 OrElse area.Columns <> 1 Then Throw New ArgumentException("A single cell is required.")
        End Sub

        Public Sub WriteValue(area As WorkbookReadArea, value As Object) Implements IWorkbookValueEditBackend.WriteValue
            RequireOwner()
            RequireExclusiveWorkbooks()
            RequireSingleCell(area)
            WorkbookCellState.ValidateValue(value)
            Dim sheets As Object = Nothing, sheet As Object = Nothing, cell As Object = Nothing
            Try
                sheets = book.Worksheets : sheet = sheets.Item(area.Worksheet) : cell = sheet.Range(area.Address)
                If CBool(sheet.ProtectContents) AndAlso CBool(cell.Locked) Then Throw New InvalidOperationException("The worksheet protects this cell.")
                If value Is Nothing Then
                    cell.ClearContents()
                ElseIf TypeOf value Is String Then
                    ' Literal entry: Excel must not reinterpret =, +, numeric or
                    ' date-looking text as a formula/number or alter its format.
                    cell.Value2 = "'" & CStr(value)
                Else
                    cell.Value2 = value
                End If
            Finally
                Release(cell) : Release(sheet) : Release(sheets)
            End Try
        End Sub
    End Class
End Namespace
