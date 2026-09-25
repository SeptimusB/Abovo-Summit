Option Strict On

Imports DevExpress.Spreadsheet

Namespace Abovo.WorkbookEngines
    Partial Friend NotInheritable Class DevExpressCalculationBackend
        Implements IWorkbookValueEditBackend

        Public Function ReadCell(area As WorkbookReadArea) As WorkbookCellState Implements IWorkbookValueEditBackend.ReadCell
            RequireOwner()
            Dim cell = SingleCell(area)
            Return New WorkbookCellState(Read(area).ValueAt(0, 0), If(cell.HasFormula, cell.FormulaInvariant, ""),
                cell.NumberFormat, cell.Protection.Locked, cell.Fill.PatternType = PatternType.Solid,
                cell.GetArrayFormulaRange() IsNot Nothing OrElse cell.GetDynamicArrayFormulaRange() IsNot Nothing,
                cell.GetMergedRanges().Count > 0, cell.Worksheet.IsProtected,
                cell.Worksheet.ConditionalFormattings.GetConditionalFormattings(cell).Any())
        End Function

        Private Function SingleCell(area As WorkbookReadArea) As Cell
            If area Is Nothing OrElse area.Rows <> 1 OrElse area.Columns <> 1 Then Throw New ArgumentException("A single cell is required.")
            If Not book.Worksheets.Contains(area.Worksheet) Then Throw New IO.InvalidDataException("Worksheet not found: " & area.Worksheet)
            Return book.Worksheets(area.Worksheet).Cells(area.Row, area.Column)
        End Function

        Public Sub WriteValue(area As WorkbookReadArea, value As Object) Implements IWorkbookValueEditBackend.WriteValue
            RequireOwner()
            WorkbookCellState.ValidateValue(value)
            Dim cell = SingleCell(area)
            If cell.Worksheet.IsProtected AndAlso cell.Protection.Locked Then Throw New InvalidOperationException("The worksheet protects this cell.")
            If value Is Nothing Then
                cell.Value = CellValue.Empty
            ElseIf TypeOf value Is Double Then
                cell.Value = CDbl(value)
            ElseIf TypeOf value Is Boolean Then
                cell.Value = CBool(value)
            Else
                Dim text = CStr(value)
                ' DevExpress consumes a leading quote as an entry prefix even
                ' through Cell.Value. Escape it so the literal text is retained.
                cell.Value = If(text.StartsWith("'", StringComparison.Ordinal), "'" & text, text)
            End If
        End Sub
    End Class
End Namespace
