Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas

Namespace Abovo
    Friend NotInheritable Class WorkbookIntegritySupport
        Friend Shared Function IsYesNoInput(cell As Cell) As Boolean
            If cell Is Nothing OrElse cell.HasFormula OrElse cell.Protection.Locked Then Return False
            Dim validation = cell.Worksheet.DataValidations.GetDataValidation(cell)
            If validation Is Nothing OrElse validation.ValidationType <> DataValidationType.List OrElse
                validation.Criteria Is Nothing OrElse Not validation.Criteria.IsText Then Return False
            Dim choices = validation.Criteria.TextValue.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries).
                Select(Function(v) v.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToArray()
            Return choices.Length = 2 AndAlso choices.Contains("Yes", StringComparer.OrdinalIgnoreCase) AndAlso choices.Contains("No", StringComparer.OrdinalIgnoreCase)
        End Function
        Friend Shared Function HasAcceptedOverride(range As CellRange, row As Integer, status As String) As Boolean
            'The visible E status deliberately remains Check. D is the workbook's
            'effective count after its unlocked Yes/No override in C is applied.
            If Not String.Equals(status, "Check", StringComparison.OrdinalIgnoreCase) Then Return False
            Dim raw = range(row, 1), choice = range(row, 2), effective = range(row, 3)
            If Not raw.ModelValue().IsNumeric OrElse raw.ModelValue().NumericValue = 0 OrElse
               choice.Protection.Locked OrElse choice.HasFormula OrElse Not choice.ModelValue().IsText OrElse
               Not String.Equals(choice.ModelValue().TextValue, "Yes", StringComparison.OrdinalIgnoreCase) OrElse
               Not effective.ModelValue().IsNumeric OrElse effective.ModelValue().NumericValue <> 0 Then Return False
            Dim validation = choice.Worksheet.DataValidations.GetDataValidation(choice)
            If validation Is Nothing OrElse validation.ValidationType <> DataValidationType.List OrElse
               validation.Criteria Is Nothing OrElse Not validation.Criteria.IsText Then Return False
            Dim choices = validation.Criteria.TextValue.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries).
                Select(Function(v) v.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToArray()
            If choices.Length <> 2 OrElse Not choices.Contains("Yes", StringComparer.OrdinalIgnoreCase) OrElse
               Not choices.Contains("No", StringComparer.OrdinalIgnoreCase) Then Return False
            Dim formula = Regex.Replace(effective.FormulaInvariant, "[\s$]", "")
            Return String.Equals(formula, "=IF(" & choice.GetReferenceA1() & "=""Yes"",0," & raw.GetReferenceA1() & ")", StringComparison.OrdinalIgnoreCase)
        End Function
    End Class

    'Only recognise a currently selected NA() branch, or an INDEX/direct-cell
    'reference traced to one. Errors in predicates/indices remain findings.
    Friend NotInheritable Class ExpectedChartGapClassifier
        Private ReadOnly Cache As New Dictionary(Of Cell, Boolean)()
        Private ReadOnly Visiting As New HashSet(Of Cell)()
        Private Shared Function ChartSheet(name As String) As Boolean
            Return name = "OW - Charts Source Data" OrElse name = "OW - Covenant Calculation" OrElse name = "OW - Live Covenant Calculation"
        End Function
        Friend Function IsExpected(cell As Cell) As Boolean
            Return IsExpected(cell, 0)
        End Function
        Private Function IsExpected(cell As Cell, depth As Integer) As Boolean
            If cell Is Nothing OrElse depth > 16 OrElse Not cell.ModelValue().IsError OrElse cell.ModelValue().ToString() <> "#N/A" OrElse Not cell.HasFormula Then Return False
            If String.Equals(cell.FormulaInvariant.Replace(" ", ""), "=NA()", StringComparison.OrdinalIgnoreCase) Then Return True
            If Not ChartSheet(cell.Worksheet.Name) Then Return False
            Dim known As Boolean
            If Cache.TryGetValue(cell, known) Then Return known
            If Not Visiting.Add(cell) Then Return False
            Try
                known = ProvesGap(cell.Worksheet.Workbook.FormulaEngine.Parse(cell.FormulaInvariant, Context(cell)).Expression, cell, depth + 1)
                Cache(cell) = known
                Return known
            Catch
                Return False 'Unsupported or ambiguous expressions are still reported.
            Finally
                Visiting.Remove(cell)
            End Try
        End Function
        Private Shared Function Text(expression As IExpression, cell As Cell) As String
            Dim builder As New StringBuilder()
            expression.BuildExpressionString(builder, cell.Worksheet.Workbook, Context(cell))
            Return "=" & builder.ToString()
        End Function
        Private Shared Function Context(cell As Cell) As ExpressionContext
            Return New ExpressionContext(cell.ColumnIndex, cell.RowIndex, cell.Worksheet, Globalization.CultureInfo.InvariantCulture, ReferenceStyle.A1, ExpressionStyle.Normal)
        End Function
        Private Function ProvesGap(expression As IExpression, cell As Cell, depth As Integer) As Boolean
            If depth > 16 Then Return False
            Dim fn = TryCast(expression, FunctionExpression)
            If fn IsNot Nothing Then
                Dim args = fn.InnerExpressions
                Select Case fn.Function.Name.ToUpperInvariant()
                    Case "NA"
                        Return args.Count = 0
                    Case "IF"
                        If args.Count <> 3 Then Return False
                        Dim predicate = TryCast(args(0), FunctionExpression), chooseFirst As Boolean
                        If predicate IsNot Nothing AndAlso predicate.Function.Name.Equals("ISNA", StringComparison.OrdinalIgnoreCase) Then
                            If predicate.InnerExpressions.Count <> 1 OrElse Not ProvesGap(predicate.InnerExpressions(0), cell, depth + 1) Then Return False
                            chooseFirst = True
                        Else
                            'Do not certify an error swallowed inside a predicate.
                            Dim check = Text(args(0), cell)
                            If Regex.IsMatch(check, "\b(?:ISNA|ISERR|ISERROR|IFERROR|IFNA)\s*\(", RegexOptions.IgnoreCase) Then Return False
                            Dim value = cell.ModelEvaluate(check)
                            If value.IsBoolean Then
                                chooseFirst = value.BooleanValue
                            ElseIf value.IsNumeric Then
                                chooseFirst = value.NumericValue <> 0
                            Else
                                Return False
                            End If
                        End If
                        Return ProvesGap(args(If(chooseFirst, 1, 2)), cell, depth + 1)
                    Case "INDEX"
                        If args.Count < 2 OrElse args.Count > 3 Then Return False
                        Dim target = ReferenceRange(args(0), cell)
                        If target Is Nothing Then Return False
                        Dim row = PositiveIndex(args(1), cell), column = 1
                        If args.Count = 3 Then
                            column = PositiveIndex(args(2), cell)
                        ElseIf target.ColumnCount <> 1 Then
                            Return False
                        End If
                        If row < 1 OrElse row > target.RowCount OrElse column < 1 OrElse column > target.ColumnCount Then Return False
                        Return IsExpected(target(row - 1, column - 1), depth + 1)
                End Select
            ElseIf TypeOf expression Is ReferenceExpression Then
                Dim target = ReferenceRange(expression, cell)
                If target IsNot Nothing AndAlso target.RowCount = 1 AndAlso target.ColumnCount = 1 Then
                    Return IsExpected(target(0, 0), depth + 1)
                End If
            End If
            Return False
        End Function
        Private Shared Function ReferenceRange(expression As IExpression, cell As Cell) As CellRange
            'Resolve native references, not the evaluated error value. Do not
            'mistake an arbitrary function's precedent list for its result range.
            If Not TypeOf expression Is ReferenceExpression AndAlso Not TypeOf expression Is RangeExpression Then Return Nothing
            Dim ranges = cell.Worksheet.Workbook.FormulaEngine.Parse(Text(expression, cell), Context(cell)).GetRanges(Context(cell))
            If ranges.Count <> 1 Then Return Nothing
            Return ranges(0)
        End Function
        Private Shared Function PositiveIndex(expression As IExpression, cell As Cell) As Integer
            Dim value = cell.ModelEvaluate(Text(expression, cell))
            If Not value.IsNumeric OrElse value.NumericValue < 1 OrElse value.NumericValue > Integer.MaxValue OrElse value.NumericValue <> Math.Truncate(value.NumericValue) Then Return 0
            Return CInt(value.NumericValue)
        End Function
    End Class
End Namespace
