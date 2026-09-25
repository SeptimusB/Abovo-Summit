Option Strict Off

Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas

Namespace Abovo.WorkbookEngines
    Friend Interface IWorkbookScalarExpressionBackend
        Function EvaluateScalar(context As WorkbookReadArea, formula As String) As Object
    End Interface

    Partial Public NotInheritable Class WorkbookCalculationSession
        ' Validation/diagnostic expressions only: no cell writes, names, active
        ' selection changes or temporary financial formulas in another engine.
        Public Async Function EvaluateCurrentAsync(anchor As WorkbookCalculationResult, context As WorkbookReadArea,
                                                   formula As String, Optional cancellation As CancellationToken = Nothing) As Task(Of Object)
            RequireSingleCell(context)
            If String.IsNullOrWhiteSpace(formula) OrElse Not formula.StartsWith("=", StringComparison.Ordinal) OrElse formula.Length > 255 Then
                Throw New ArgumentException("A scalar workbook expression must start with '=' and contain at most 255 characters.", NameOf(formula))
            End If
            ' External workbooks/connections are not part of the opening policy.
            ' Structured table references need separately qualified support too.
            If formula.Contains("[") OrElse formula.Contains("]") Then Throw New NotSupportedException("External or structured references are not supported by this scalar expression route.")
            If Not IsCurrent(anchor) Then Throw New InvalidOperationException("A current calculation is required before evaluating validation.")
            Dim operation = owner.InvokeAsync(Function()
                If Not IsCurrent(anchor) Then Throw New StaleResultException()
                cancellation.ThrowIfCancellationRequested()
                Dim evaluator = TryCast(backend, IWorkbookScalarExpressionBackend)
                If evaluator Is Nothing Then Throw New NotSupportedException("The selected engine cannot evaluate this validation expression.")
                Dim value As Object
                Try
                    value = evaluator.EvaluateScalar(context, formula)
                    ' Enforce the same detached scalar/error transport boundary.
                    Dim validated As New WorkbookValueBlock(context, New Object(,) {{value}})
                    value = validated.ValueAt(0, 0)
                Catch ex As NotSupportedException
                    Throw
                Catch
                    SyncLock gate
                        failed = True
                    End SyncLock
                    Throw
                End Try
                cancellation.ThrowIfCancellationRequested()
                If Not IsCurrent(anchor) Then Throw New StaleResultException()
                Return value
            End Function, cancellation)
            Return Await AwaitOperation(operation, "validation expression").ConfigureAwait(False)
        End Function
    End Class

    Partial Friend NotInheritable Class DevExpressCalculationBackend
        Implements IWorkbookScalarExpressionBackend
        Public Function EvaluateScalar(context As WorkbookReadArea, formula As String) As Object Implements IWorkbookScalarExpressionBackend.EvaluateScalar
            RequireOwner()
            Dim expressionContext As New ExpressionContext(context.Column, context.Row, book.Worksheets(context.Worksheet),
                CultureInfo.InvariantCulture, ReferenceStyle.A1, ExpressionStyle.Array)
            Dim value = book.FormulaEngine.Evaluate(formula, expressionContext)
            If value.IsRange Then
                If value.RangeAreas.Count <> 1 OrElse value.RangeValue.RowCount <> 1 OrElse value.RangeValue.ColumnCount <> 1 Then
                    Throw New NotSupportedException("The validation expression did not return a scalar value.")
                End If
                value = value.RangeValue(0, 0).Value
            End If
            If value.IsArray Then
                Dim array = value.ArrayValue
                If array.Length <> 1 Then Throw New NotSupportedException("The validation expression did not return a scalar value.")
                value = array(0, 0)
            End If
            If value.IsNumeric Then Return value.NumericValue
            If value.IsBoolean Then Return value.BooleanValue
            If value.IsText Then Return value.TextValue
            If value.IsError Then Return New WorkbookCellError(value.ToString())
            If value.IsEmpty Then Return Nothing
            Throw New NotSupportedException("The validation expression did not return a scalar value.")
        End Function
    End Class

    Partial Friend NotInheritable Class ExcelCalculationBackend
        Implements IWorkbookScalarExpressionBackend
        Public Function EvaluateScalar(context As WorkbookReadArea, formula As String) As Object Implements IWorkbookScalarExpressionBackend.EvaluateScalar
            RequireOwner()
            Dim sheets As Object = Nothing, sheet As Object = Nothing, nativeResult As Object = Nothing
            Try
                sheets = book.Worksheets : sheet = sheets.Item(context.Worksheet)
                nativeResult = sheet.Evaluate(formula)
                Dim raw As Object = nativeResult
                If raw IsNot Nothing AndAlso Marshal.IsComObject(raw) Then raw = nativeResult.Value2
                Dim array = TryCast(raw, Array)
                If array IsNot Nothing Then
                    If array.Rank <> 2 OrElse array.Length <> 1 Then Throw New NotSupportedException("The validation expression did not return a scalar value.")
                    raw = array.GetValue(array.GetLowerBound(0), array.GetLowerBound(1))
                End If
                Return ConvertValue(raw)
            Finally
                If nativeResult IsNot Nothing AndAlso Marshal.IsComObject(nativeResult) Then Release(nativeResult)
                Release(sheet) : Release(sheets)
            End Try
        End Function
    End Class
End Namespace
