Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas

Namespace Abovo
    'DX 25.2.4 XLSB export replaces built-in calls with >30 arguments with
    '=#VALUE!, although XLSX export and Excel support the original formula.
    'Only CONCATENATE has a reviewed rewrite. Do not guess at other functions.
    Friend NotInheritable Class WorkbookXlsbFormulaCompatibility
        Private Class Normalizer
            Inherits ExpressionVisitor
            Public Changed As Boolean
            Public Overrides Sub Visit(expression As FunctionExpression)
                MyBase.Visit(expression)
                If expression.InnerExpressions.Count <= 30 Then Return
                If Not String.Equals(expression.Function.Name, "CONCATENATE", StringComparison.OrdinalIgnoreCase) Then
                    Throw New InvalidOperationException("XLSB export cannot safely preserve " & expression.Function.Name & " with more than 30 arguments. No file has been saved.")
                End If
                Dim arguments = expression.InnerExpressions.ToList()
                'Larger synthetic expressions hit additional XLSB exporter
                'limits even when nested. Fail closed outside the verified
                '31..60 argument repair (the AGL dashboard uses 37).
                If expression.InnerExpressions.Count > 60 Then Throw New InvalidOperationException("CONCATENATE exceeds the verified XLSB repair limit of 60 arguments. No file has been saved.")
                'Keep groups shallow and modest in size.
                Do While arguments.Count > 16
                    Dim groups As New List(Of IExpression)
                    For index As Integer = 0 To arguments.Count - 1 Step 16
                        groups.Add(New FunctionExpression(expression.Function, arguments.Skip(index).Take(16).ToList()))
                    Next
                    arguments = groups
                Loop
                expression.InnerExpressions = arguments
                Changed = True
            End Sub
        End Class

        Friend Shared Function Normalize(wb As IWorkbook, formula As String) As String
            'Cheap superset filter; the parser handles commas in strings,
            'arrays, quoted sheet names and nested calls correctly.
            If String.IsNullOrEmpty(formula) OrElse formula.Count(Function(c) c = ","c) < 30 Then Return formula
            Dim parsed = wb.FormulaEngine.Parse(formula)
            If parsed Is Nothing OrElse parsed.Expression Is Nothing Then Throw New InvalidOperationException("A formula could not be checked for safe XLSB export.")
            Dim visitor As New Normalizer()
            parsed.Expression.Visit(visitor)
            If Not visitor.Changed Then Return formula
            Dim text As New StringBuilder()
            parsed.Expression.BuildExpressionString(text, wb)
            Return "=" & text.ToString()
        End Function

        Private Class Edit
            Public Original As String
            Public Replacement As String
            Public Write As Action(Of String)
        End Class

        Private Shared Function CollectEdits(wb As IWorkbook) As List(Of Edit)
            Dim edits As New List(Of Edit)
            'Finish preflight before changing anything or touching the file.
            For Each ws As Worksheet In wb.Worksheets
                For Each cell As Cell In ws.GetUsedRange().ExistingCells
                    If Not cell.HasFormula Then Continue For
                    Dim original = cell.FormulaInvariant
                    Dim replacement As String
                    Try
                        replacement = Normalize(wb, original)
                    Catch ex As Exception
                        Throw New InvalidOperationException(ws.Name & "!" & cell.GetReferenceA1() & ": " & ex.Message, ex)
                    End Try
                    If replacement = original Then Continue For
                    If cell.HasDynamicArrayFormula Then Throw New InvalidOperationException("Cannot safely rewrite dynamic array at " & ws.Name & "!" & cell.GetReferenceA1())
                    If cell.HasArrayFormula Then
                        Dim area = cell.GetArrayFormulaRange()
                        If cell.RowIndex <> area.TopRowIndex OrElse cell.ColumnIndex <> area.LeftColumnIndex Then Continue For
                        edits.Add(New Edit With {.Original = original, .Replacement = replacement, .Write = Sub(value) area.ArrayFormulaInvariant = value})
                    Else
                        Dim target = cell
                        edits.Add(New Edit With {.Original = original, .Replacement = replacement, .Write = Sub(value) target.FormulaInvariant = value})
                    End If
                Next
            Next
            Dim names = wb.DefinedNames.Concat(wb.Worksheets.SelectMany(Function(ws) ws.DefinedNames)).Distinct()
            For Each name As DefinedName In names
                Dim target = name, original = name.RefersTo
                Dim replacement As String
                Try
                    replacement = Normalize(wb, original)
                Catch ex As Exception
                    Throw New InvalidOperationException("Defined name '" & name.Name & "': " & ex.Message, ex)
                End Try
                If replacement <> original Then edits.Add(New Edit With {.Original = original, .Replacement = replacement, .Write = Sub(value) target.RefersTo = value})
            Next
            Return edits
        End Function

        Friend Shared Sub Save(wb As IWorkbook, saveAction As Func(Of Boolean), Optional formulasChanging As Action = Nothing,
                               Optional checkFormulas As Boolean = True)
            Dim clock = Abovo.SummitDiagnostics.DiagnosticTimer.StartNew()
            'Only the owning model can reuse a previously verified formula state.
            'Other callers retain the conservative full preflight by default.
            Dim edits = If(checkFormulas, CollectEdits(wb), New List(Of Edit)())
            Abovo.SummitDiagnostics.WriteLine("[XLSB Save Benchmark] formulaPreflight=" & clock.ElapsedMilliseconds.ToString() &
                            " ms, rewritten=" & edits.Count.ToString() & ", skipped=" & (Not checkFormulas).ToString())
            Dim historyEnabled = wb.History.IsEnabled
            Dim previousMode = wb.Options.CalculationMode
            Dim applied As New List(Of Edit)
            Dim succeeded As Boolean = False
            Try
                wb.Options.CalculationMode = WorkbookCalculationMode.Manual
                If edits.Count > 0 Then wb.History.IsEnabled = False
                If edits.Count > 0 AndAlso formulasChanging IsNot Nothing Then formulasChanging()
                For Each item In edits
                    applied.Add(item)
                    item.Write(item.Replacement)
                Next
                succeeded = saveAction()
            Finally
                Try
                    If Not succeeded Then
                        For Each item In applied.AsEnumerable().Reverse()
                            item.Write(item.Original)
                        Next
                    End If
                Finally
                    Try
                        wb.History.IsEnabled = historyEnabled
                    Finally
                        wb.Options.CalculationMode = previousMode
                    End Try
                End Try
            End Try
        End Sub
    End Class
End Namespace
