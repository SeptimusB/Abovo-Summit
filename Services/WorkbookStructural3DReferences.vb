Imports System.Globalization
Imports System.Linq
Imports System.Text
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas

Namespace Abovo
    'DX 25.2 leaves 3-D references unchanged on column insert/delete/copy.
    'Use its public formula syntax tree, never textual formula substitutions.
    'Opted into by reviewed linked-column families. No schema changes.
    Friend Class WorkbookStructural3DReferences
        Private ReadOnly Workbook As IWorkbook
        Private ReadOnly Items As New List(Of SavedFormula)
        Private ChangedSheets As HashSet(Of String)
        Private CapturedEdits As KeyValuePair(Of Integer, Integer)()
        Private SheetOrder As String()
        Private Applied As Boolean

        Private Class ReferenceCollector
            Inherits ExpressionVisitor
            Public ReadOnly References As New List(Of CellReferenceExpression)
            Public Overrides Sub Visit(expression As CellReferenceExpression)
                If expression.SheetReference IsNot Nothing AndAlso
                   Not String.IsNullOrEmpty(expression.SheetReference.EndSheetName) AndAlso
                   expression.SheetReference.StartSheetName <> expression.SheetReference.EndSheetName Then
                    References.Add(expression)
                End If
                MyBase.Visit(expression)
            End Sub
        End Class

        Private Class SavedFormula
            Public Sheet As Worksheet
            Public Row As Integer
            Public Column As Integer
            Public Name As DefinedName
            Public References As List(Of CellReferenceExpression)
        End Class

        Private Sub New(wb As IWorkbook)
            Workbook = wb
        End Sub

        Public Shared Function Capture(wb As IWorkbook, sheets As IEnumerable(Of String),
                                       edits As IEnumerable(Of KeyValuePair(Of Integer, Integer))) As WorkbookStructural3DReferences
            'Positive count inserts; negative count deletes at zero-based start.
            Dim result As New WorkbookStructural3DReferences(wb)
            Dim changed As New HashSet(Of String)(sheets, StringComparer.OrdinalIgnoreCase)
            Dim changes = edits.ToArray()
            result.ChangedSheets = changed
            result.CapturedEdits = changes
            result.SheetOrder = wb.Worksheets.Select(Function(ws) ws.Name).ToArray()
            'Operation-local only: worksheet names/order may change between commands.
            'Build the span lookup once, not for every parsed formula.
            Dim order = wb.Worksheets.Select(Function(ws, i) New With {.Name = ws.Name, .Index = i}).ToDictionary(Function(x) x.Name, Function(x) x.Index, StringComparer.OrdinalIgnoreCase)
            Dim selectedPrefix(wb.Worksheets.Count) As Integer
            For index As Integer = 0 To wb.Worksheets.Count - 1
                selectedPrefix(index + 1) = selectedPrefix(index) + If(changed.Contains(wb.Worksheets(index).Name), 1, 0)
            Next
            For Each ws As Worksheet In wb.Worksheets
                For Each cell As Cell In ws.GetUsedRange().ExistingCells
                    If Not cell.HasFormula Then Continue For
                    Dim formula = cell.FormulaInvariant
                    If Not MayContainThreeDReference(formula) Then Continue For
                    Dim destination = If(changed.Contains(ws.Name), MapColumn(cell.ColumnIndex, changes), cell.ColumnIndex)
                    If destination < 0 Then Continue For 'Formula is itself being deleted.
                    Dim refs = ReadReferences(wb, formula)
                    If Not Transform(refs, order, selectedPrefix, changes) Then Continue For
                    ValidateArrayFormula(cell)
                    result.Items.Add(New SavedFormula With {.Sheet = ws, .Row = cell.RowIndex, .Column = destination, .References = refs})
                Next
            Next
            For Each dn As DefinedName In wb.DefinedNames
                If Not MayContainThreeDReference(dn.RefersTo) Then Continue For
                Dim refs = ReadReferences(wb, dn.RefersTo)
                If Transform(refs, order, selectedPrefix, changes) Then result.Items.Add(New SavedFormula With {.Name = dn, .References = refs})
            Next
            Return result
        End Function

        'Only for adjacent positive batches of the reviewed Development command,
        'while its bulk-mutation guard excludes unrelated edits. An old reference
        'unaffected by insertion at X cannot become affected at X + count.
        'Newly copied/overwritten columns are the exception: inspect those afresh.
        Public Function CaptureFollowingInsertion(start As Integer, count As Integer,
                                                   copiedRanges As IEnumerable(Of CellRange)) As WorkbookStructural3DReferences
            If Not Applied OrElse CapturedEdits.Length <> 1 OrElse
               CapturedEdits(0).Value <= 0 OrElse count <= 0 OrElse
               start <> CapturedEdits(0).Key + CapturedEdits(0).Value OrElse
               Not SheetOrder.SequenceEqual(Workbook.Worksheets.Select(Function(ws) ws.Name)) Then
                Throw New InvalidOperationException("Reference capture reuse requires an applied, adjacent insertion with unchanged worksheet order.")
            End If
            Dim ranges = copiedRanges.ToArray()
            If ranges.Any(Function(r) Not ChangedSheets.Contains(r.Worksheet.Name) OrElse
                                      Not Object.ReferenceEquals(r.Worksheet, Workbook.Worksheets(r.Worksheet.Name))) Then
                Throw New InvalidOperationException("Copied reference ranges must belong to the insertion's worksheets.")
            End If
            Dim result As New WorkbookStructural3DReferences(Workbook) With {
                .ChangedSheets = New HashSet(Of String)(ChangedSheets, StringComparer.OrdinalIgnoreCase),
                .CapturedEdits = {New KeyValuePair(Of Integer, Integer)(start, count)},
                .SheetOrder = SheetOrder
            }
            Dim order = SheetOrder.Select(Function(name, index) New With {.Name = name, .Index = index}).
                ToDictionary(Function(x) x.Name, Function(x) x.Index, StringComparer.OrdinalIgnoreCase)
            Dim selectedPrefix(SheetOrder.Length) As Integer
            For index As Integer = 0 To SheetOrder.Length - 1
                selectedPrefix(index + 1) = selectedPrefix(index) + If(ChangedSheets.Contains(SheetOrder(index)), 1, 0)
            Next
            Dim captureCell As Action(Of Cell) =
                Sub(cell)
                    If Not cell.HasFormula Then Return
                    Dim formula = cell.FormulaInvariant
                    If Not MayContainThreeDReference(formula) Then Return
                    Dim refs = ReadReferences(Workbook, formula)
                    If Not Transform(refs, order, selectedPrefix, result.CapturedEdits) Then Return
                    ValidateArrayFormula(cell)
                    result.Items.Add(New SavedFormula With {.Sheet = cell.Worksheet, .Row = cell.RowIndex,
                        .Column = If(ChangedSheets.Contains(cell.Worksheet.Name), MapColumn(cell.ColumnIndex, result.CapturedEdits), cell.ColumnIndex),
                        .References = refs})
                End Sub
            For Each item In Items
                If item.Name IsNot Nothing Then
                    Dim refs = ReadReferences(Workbook, item.Name.RefersTo)
                    If Transform(refs, order, selectedPrefix, result.CapturedEdits) Then
                        result.Items.Add(New SavedFormula With {.Name = item.Name, .References = refs})
                    End If
                ElseIf Not ranges.Any(Function(r) Object.ReferenceEquals(r.Worksheet, item.Sheet) AndAlso
                    item.Column >= r.LeftColumnIndex AndAlso item.Column <= r.RightColumnIndex AndAlso
                    item.Row >= r.TopRowIndex AndAlso item.Row <= r.BottomRowIndex) Then
                    captureCell(item.Sheet.Cells(item.Row, item.Column))
                End If
            Next
            For Each range In ranges
                For Each cell As Cell In range.ExistingCells
                    captureCell(cell)
                Next
            Next
            Return result
        End Function

        Public Sub Verify()
            If Not Applied Then Throw New InvalidOperationException("Structural references have not been applied.")
            For Each item In Items
                Dim formula = If(item.Name IsNot Nothing, item.Name.RefersTo, item.Sheet.Cells(item.Row, item.Column).FormulaInvariant)
                Dim actual = ReadReferences(Workbook, formula)
                If actual.Count <> item.References.Count Then Throw New InvalidOperationException("Structural reference verification failed: reference count changed.")
                For index As Integer = 0 To actual.Count - 1
                    Dim expectedText As New StringBuilder(), actualText As New StringBuilder()
                    item.References(index).BuildExpressionString(expectedText, Workbook)
                    actual(index).BuildExpressionString(actualText, Workbook)
                    If expectedText.ToString() <> actualText.ToString() Then Throw New InvalidOperationException("Structural reference verification failed: reference changed.")
                Next
            Next
        End Sub

        Friend Shared Function MayContainThreeDReference(formula As String) As Boolean
            'A conservative lexical screen, NOT a formula parser or rewrite. Check
            'the sheet qualifier immediately before ! for a span colon. This avoids
            'regex backtracking over long INDEX/MATCH formulas without skipping
            'quoted/escaped, Unicode or external sheet names. False positives are
            'harmless: ReadReferences still uses the native public syntax tree.
            Dim colon = formula.IndexOf(":"c)
            If colon < 0 Then Return False
            Dim bang = formula.IndexOf("!"c, colon + 1)
            While bang >= 0
                Dim index = bang - 1
                If formula(index) = "'"c Then
                    index -= 1
                    While index >= 0
                        Dim value = formula(index)
                        If value = ":"c Then Return True
                        If value = "'"c Then
                            If index > 0 AndAlso formula(index - 1) = "'"c Then
                                index -= 2 'Escaped apostrophe inside the sheet name.
                                Continue While
                            End If
                            Exit While
                        End If
                        index -= 1
                    End While
                Else
                    While index >= 0
                        Dim value = formula(index)
                        If value = ":"c Then Return True
                        If Char.IsWhiteSpace(value) OrElse "!+-*/^&=<>(),;{}[]'""".IndexOf(value) >= 0 Then Exit While
                        index -= 1
                    End While
                End If
                bang = formula.IndexOf("!"c, bang + 1)
            End While
            Return False
        End Function

        Private Shared Function ReadReferences(wb As IWorkbook, formula As String) As List(Of CellReferenceExpression)
            Dim parsed = wb.FormulaEngine.Parse(formula)
            If parsed Is Nothing OrElse parsed.Expression Is Nothing Then Throw New InvalidOperationException("A 3-D formula could not be parsed before structural resize.")
            Dim visitor As New ReferenceCollector()
            parsed.Expression.Visit(visitor)
            Return visitor.References.Select(Function(r) r.Clone()).ToList()
        End Function

        Private Shared Function Transform(refs As List(Of CellReferenceExpression), order As Dictionary(Of String, Integer),
                                          selectedPrefix As Integer(), edits As KeyValuePair(Of Integer, Integer)()) As Boolean
            Dim affected As Boolean = False
            For Each reference In refs
                Dim sheetRange = reference.SheetReference
                If sheetRange.Type = SheetReferenceType.External Then Continue For
                Dim first, last As Integer
                If Not order.TryGetValue(sheetRange.StartSheetName, first) OrElse Not order.TryGetValue(sheetRange.EndSheetName, last) Then Throw New InvalidOperationException("A 3-D worksheet reference cannot be resolved before structural resize.")
                Dim low = Math.Min(first, last), high = Math.Max(first, last)
                Dim selected = selectedPrefix(high + 1) - selectedPrefix(low)
                If selected = 0 Then Continue For
                If selected <> high - low + 1 Then Throw New InvalidOperationException("structural resize crosses a partially selected 3-D worksheet span: " & sheetRange.StartSheetName & ":" & sheetRange.EndSheetName & ". No columns have been changed.")
                Dim area = reference.CellArea
                Dim left = area.LeftColumnIndex, right = area.RightColumnIndex
                If left = 0 AndAlso right = 16383 Then Continue For 'Whole rows retain their full width.
                For Each edit In edits
                    Dim start = edit.Key, amount = edit.Value
                    If amount > 0 Then
                        If left >= start Then left += amount
                        If right >= start Then right += amount
                    Else
                        Dim finish = start - amount
                        If left >= finish Then
                            left += amount
                        ElseIf left >= start Then
                            left = start
                        End If
                        If right >= finish Then
                            right += amount
                        ElseIf right >= start Then
                            right = start - 1
                        End If
                    End If
                Next
                If left < 0 OrElse right > 16383 OrElse left > right Then Throw New InvalidOperationException("structural resize would invalidate an existing 3-D reference. No columns have been changed.")
                If left = area.LeftColumnIndex AndAlso right = area.RightColumnIndex Then Continue For
                reference.CellArea = New CellArea(
                    New CellReferencePosition(left, area.TopLeft.Row, area.TopLeft.ColumnType, area.TopLeft.RowType),
                    New CellReferencePosition(right, area.BottomRight.Row, area.BottomRight.ColumnType, area.BottomRight.RowType))
                affected = True
            Next
            Return affected
        End Function

        Private Shared Function MapColumn(column As Integer, edits As KeyValuePair(Of Integer, Integer)()) As Integer
            For Each edit In edits
                If edit.Value > 0 Then
                    If column >= edit.Key Then column += edit.Value
                ElseIf column >= edit.Key - edit.Value Then
                    column += edit.Value
                ElseIf column >= edit.Key Then
                    Return -1
                End If
            Next
            Return column
        End Function

        Public Sub Apply(modelID As Integer)
            For Each item In Items.Where(Function(i) i.Name IsNot Nothing)
                item.Name.RefersTo = ReplaceReferences(Workbook, item.Name.RefersTo, item.References)
            Next
            For Each group In Items.Where(Function(i) i.Sheet IsNot Nothing).GroupBy(Function(i) i.Sheet)
                Dim ws = group.Key, wasProtected = ws.IsProtected
                Try
                    If wasProtected Then WSSecurity.UNProtectWS(modelID, ws.Name)
                    For Each item In group
                        Dim cell = ws.Cells(item.Row, item.Column)
                        SetFormulaPreservingArray(cell, ReplaceReferences(Workbook, cell.FormulaInvariant, item.References))
                    Next
                Finally
                    If wasProtected Then WSSecurity.ProtectWS(modelID, ws.Name)
                End Try
            Next
            Applied = True
        End Sub

        Private Shared Function ReplaceReferences(wb As IWorkbook, formula As String,
                                                   expected As List(Of CellReferenceExpression)) As String
            Dim parsed = wb.FormulaEngine.Parse(formula)
            If parsed Is Nothing OrElse parsed.Expression Is Nothing Then Throw New InvalidOperationException("A 3-D formula could not be parsed during structural resize.")
            Dim visitor As New ReferenceCollector()
            parsed.Expression.Visit(visitor)
            If visitor.References.Count <> expected.Count Then Throw New InvalidOperationException("3-D formula shape changed unexpectedly during structural resize.")
            For index As Integer = 0 To expected.Count - 1
                Dim target = visitor.References(index), source = expected(index)
                If target.SheetReference.StartSheetName <> source.SheetReference.StartSheetName OrElse
                   target.SheetReference.EndSheetName <> source.SheetReference.EndSheetName Then Throw New InvalidOperationException("3-D worksheet span changed unexpectedly during structural resize.")
                target.CellArea = source.CellArea.Clone()
            Next
            Dim text As New StringBuilder()
            parsed.Expression.BuildExpressionString(text, wb)
            Return "=" & text.ToString()
        End Function

        Public Shared Sub CopyColumn(ws As Worksheet, template As Integer, first As Integer,
                                     count As Integer, top As Integer, bottom As Integer)
            For Each source As Cell In ws.Range.FromLTRB(template, top, template, bottom).ExistingCells
                If Not source.HasFormula OrElse Not MayContainThreeDReference(source.FormulaInvariant) Then Continue For
                ValidateArrayFormula(source)
                Dim refs = ReadReferences(ws.Workbook, source.FormulaInvariant)
                If refs.Count = 0 Then Continue For
                For column As Integer = first To first + count - 1
                    Dim translated = refs.Select(Function(r) r.Clone()).ToList()
                    For Each reference In translated
                        Dim area = reference.CellArea
                        If area.LeftColumnIndex = 0 AndAlso area.RightColumnIndex = 16383 Then Continue For
                        Dim tl = area.TopLeft, br = area.BottomRight
                        Dim left = tl.Column + If(tl.ColumnType = PositionType.Relative, column - template, 0)
                        Dim right = br.Column + If(br.ColumnType = PositionType.Relative, column - template, 0)
                        If left < 0 OrElse right > 16383 Then Throw New InvalidOperationException("A copied 3-D reference is outside the worksheet.")
                        reference.CellArea = New CellArea(New CellReferencePosition(left, tl.Row, tl.ColumnType, tl.RowType), New CellReferencePosition(right, br.Row, br.ColumnType, br.RowType))
                    Next
                    Dim target = ws.Cells(source.RowIndex, column)
                    SetFormulaPreservingArray(target, ReplaceReferences(ws.Workbook, target.FormulaInvariant, translated))
                Next
            Next
        End Sub
        Private Shared Sub ValidateArrayFormula(cell As Cell)
            If cell.HasDynamicArrayFormula Then Throw New InvalidOperationException("Structural resize cannot safely update a dynamic 3-D array formula at " & cell.Worksheet.Name & "!" & cell.GetReferenceA1())
            If cell.HasArrayFormula Then
                Dim area = cell.GetArrayFormulaRange()
                If area.RowCount <> 1 OrElse area.ColumnCount <> 1 Then Throw New InvalidOperationException("Structural resize cannot safely split a multi-cell 3-D array at " & cell.Worksheet.Name & "!" & area.GetReferenceA1())
            End If
        End Sub

        Private Shared Sub SetFormulaPreservingArray(cell As Cell, formula As String)
            ValidateArrayFormula(cell)
            If cell.HasArrayFormula Then
                'Single-cell legacy arrays are used by JV Interest Received.
                'Retain array semantics; ordinary FormulaInvariant would flatten it.
                cell.GetArrayFormulaRange().ArrayFormulaInvariant = formula
            Else
                cell.FormulaInvariant = formula
            End If
        End Sub
    End Class
End Namespace
