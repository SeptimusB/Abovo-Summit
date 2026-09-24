Option Infer On
Imports System.Globalization
Imports Abovo
Imports Abovo.DataObject
Imports Abovo.FileManager
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.XtraVerticalGrid.Rows

Partial Public Class DataInterfaceTemplate
    Private Shared Function ExclusiveMinimumInputError(tag As SingleCellDataTag, value As Object) As String
        Dim column As New DataColumnTag With {
            .DataType = tag.DataType,
            .MinVal = If(tag.MinValSet, tag.MinVal.ToString(CultureInfo.InvariantCulture), Nothing),
            .MaxVal = If(tag.MaxValSet, tag.MaxVal.ToString(CultureInfo.InvariantCulture), Nothing),
            .MinExclusive = tag.MinExclusive}
        Return NumericInputError(tag.TargetWorksheet.Cells(tag.TargetCell), column, value)
    End Function

    Private Shared Function NumericLimit(text As String, sentinel As String) As Decimal?
        If String.IsNullOrWhiteSpace(text) OrElse text.Trim().Equals(sentinel, StringComparison.OrdinalIgnoreCase) Then Return Nothing
        Return Decimal.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture)
    End Function

    Private Shared Sub ConfigurePercentageEditor(editor As RepositoryItemSpinEdit, tag As DataColumnTag)
        Dim minimum = NumericLimit(tag.MinVal, "NOMIN")
        Dim maximum = NumericLimit(tag.MaxVal, "NOMAX")
        'DevExpress's default 0/0 pair means unrestricted, not a 0% minimum.
        'Workbook rules are checked at commit/paste, against the actual cell.
        If minimum.HasValue OrElse maximum.HasValue Then
            editor.MinValue = If(minimum, Decimal.MinValue)
            editor.MaxValue = If(maximum, Decimal.MaxValue)
        End If
        editor.Increment = If(String.IsNullOrWhiteSpace(tag.DefIncrement), 0.0025D, Decimal.Parse(tag.DefIncrement, CultureInfo.InvariantCulture))
    End Sub

    Private Function InputSourceCell(source As AbovoUnboundSource, row As Integer, column As Integer) As Cell
        If source Is Nothing OrElse source.UBSTag Is Nothing OrElse DataPres Is Nothing Then Return Nothing
        Dim index = source.UBSTag.DSIndex
        If index < 0 OrElse index >= DataPres.DataSets.Count Then Return Nothing
        Dim data = DataPres.DataSets(index)
        If data Is Nothing OrElse row < 0 OrElse row >= data.DataRows.Count OrElse column < 0 OrElse column >= data.DataColumns.Count Then Return Nothing
        Dim point = data.DataRows(row).DataCells(column)
        If point Is Nothing OrElse String.IsNullOrWhiteSpace(point.SourceSheet) OrElse String.IsNullOrWhiteSpace(point.SourceAddress) Then Return Nothing
        Return ExcelModels(ModelID).WB.Worksheets(point.SourceSheet).Cells(point.SourceAddress)
    End Function

    Private Sub InputGridDisplayText(sender As Object, e As CustomColumnDisplayTextEventArgs)
        Dim view = TryCast(sender, GridView)
        Dim tag = TryCast(e.Column?.Tag, DataColumnTag)
        If view Is Nothing OrElse Not IsMoneyField(tag) OrElse e.ListSourceRowIndex < 0 Then Return
        Dim cell = InputSourceCell(TryCast(view.GridControl.DataSource, AbovoUnboundSource), e.ListSourceRowIndex, GetGridColumnIndex(e.Column))
        If cell IsNot Nothing Then e.DisplayText = cell.DisplayText
    End Sub

    Private Sub InputVGridDisplayText(sender As Object, e As DevExpress.XtraVerticalGrid.Events.CustomRecordDisplayTextEventArgs)
        Dim grid = TryCast(sender, VGridControl)
        If grid Is Nothing OrElse e.Properties Is Nothing Then Return
        Dim row = e.Properties.Row
        Dim multi = TryCast(row, MultiEditorRow)
        Dim item As Integer = If(multi Is Nothing, 0, multi.PropertiesCollection.IndexOf(e.Properties))
        Dim tag = GetVGridColumnTag(row, item)
        If Not IsMoneyField(tag) Then Return
        Dim cell = InputSourceCell(TryCast(grid.DataSource, AbovoUnboundSource), grid.GetDataSourceRecordIndex(e.Record), GetVGridColumnIndex(row, item))
        If cell IsNot Nothing Then e.DisplayText = cell.DisplayText
    End Sub

    Private Shared Function IsMoneyField(tag As DataColumnTag) As Boolean
        Return tag IsNot Nothing AndAlso (tag.DataType = "M" OrElse tag.DataType = "SM")
    End Function

    'The same gate is used before native editor commits and before ANY paste
    'writes. It never temporarily writes to the worksheet to test validation.
    Friend Shared Function NumericInputError(cell As Cell, tag As DataColumnTag, value As Object) As String
        If tag Is Nothing OrElse value Is Nothing OrElse Convert.IsDBNull(value) Then Return Nothing
        Select Case tag.DataType
            Case "I", "Y", "N", "P", "C", "M", "SM", "R"
            Case Else
                Return Nothing
        End Select
        Try
            Dim number = Convert.ToDouble(value, CultureInfo.CurrentCulture)
            If Double.IsNaN(number) OrElse Double.IsInfinity(number) Then Return "Enter a finite number."
            If (tag.DataType = "I" OrElse tag.DataType = "Y") AndAlso number <> Math.Truncate(number) Then Return "Enter a whole number."
            Dim minimum = NumericLimit(tag.MinVal, "NOMIN")
            Dim maximum = NumericLimit(tag.MaxVal, "NOMAX")
            If minimum.HasValue AndAlso tag.MinExclusive AndAlso number <= CDbl(minimum.Value) Then Return "Enter a number greater than " & minimum.Value.ToString(CultureInfo.CurrentCulture) & "."
            If minimum.HasValue AndAlso number < CDbl(minimum.Value) Then Return "The minimum is " & minimum.Value.ToString(CultureInfo.CurrentCulture) & "."
            If maximum.HasValue AndAlso number > CDbl(maximum.Value) Then Return "The maximum is " & maximum.Value.ToString(CultureInfo.CurrentCulture) & "."
            If cell Is Nothing Then Return "The source cell could not be resolved."
            For Each rule In cell.Worksheet.DataValidations.GetDataValidations(cell)
                If rule.ValidationType <> DataValidationType.Decimal AndAlso rule.ValidationType <> DataValidationType.WholeNumber Then Continue For
                If rule.ValidationType = DataValidationType.WholeNumber AndAlso number <> Math.Truncate(number) Then Return "The workbook requires a whole number."
                Dim first = NumericCriterion(rule.Criteria, rule, cell)
                Dim valid As Boolean
                Select Case rule.Operator
                    Case DataValidationOperator.Between
                        valid = number >= first AndAlso number <= NumericCriterion(rule.Criteria2, rule, cell)
                    Case DataValidationOperator.NotBetween
                        valid = number < first OrElse number > NumericCriterion(rule.Criteria2, rule, cell)
                    Case DataValidationOperator.Equal
                        valid = number = first
                    Case DataValidationOperator.NotEqual
                        valid = number <> first
                    Case DataValidationOperator.GreaterThan
                        valid = number > first
                    Case DataValidationOperator.GreaterThanOrEqual
                        valid = number >= first
                    Case DataValidationOperator.LessThan
                        valid = number < first
                    Case DataValidationOperator.LessThanOrEqual
                        valid = number <= first
                    Case Else
                        Return "Unsupported workbook numeric validation."
                End Select
                If Not valid Then Return "This value is outside the workbook's permitted numeric range."
            Next
            Return Nothing
        Catch ex As Exception
            Return "Cannot validate this number: " & ex.Message
        End Try
    End Function

    Private Shared Function NumericCriterion(value As ValueObject, rule As DataValidation, cell As Cell) As Double
        If value Is Nothing Then Throw New FormatException("Missing workbook validation limit.")
        If Not value.IsFormula Then Return value.NumericValue
        Dim engine = cell.Worksheet.Workbook.FormulaEngine
        Dim origin As New ExpressionContext(rule.Range.LeftColumnIndex, rule.Range.TopRowIndex, cell.Worksheet)
        Dim expression = engine.Parse(value.FormulaInvariant, origin)
        origin.ReferenceStyle = ReferenceStyle.R1C1
        Dim target As New ExpressionContext(cell.ColumnIndex, cell.RowIndex, cell.Worksheet) With {.ReferenceStyle = ReferenceStyle.R1C1}
        Dim result = engine.Evaluate(expression.ToString(origin), target)
        If Not result.IsNumeric Then Throw New FormatException("The workbook validation limit is not numeric.")
        Return result.NumericValue
    End Function
End Class
