Option Infer On
Imports System.Globalization
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Security.Cryptography
Imports DevExpress.Spreadsheet

Namespace Abovo
    'Read-only v26 adapter. The workbook remains the owner of headline calculations.
    Public Class BalanceSheetCellStyle
        Public Property Format As String
        Public Property Foreground As Integer
        Public Property Background As Integer
        Public Property FontName As String
        Public Property Bold As Boolean
        Public Property Italic As Boolean
        Public Property Underline As Boolean
        Public Property Alignment As Integer
    End Class

    Public Class BalanceSheetNode
        Public Property Id As String
        Public Property ParentId As String = ""
        Public Property Caption As String
        Public Property Source As String
        Public Property Rule As String
        Public Property Diagnostic As String = ""
        Public Property IsTotal As Boolean
        Public Property IsHeadline As Boolean
        Public Property Values As Double()
        Public Property Styles As Integer()
    End Class

    Public Class BalanceSheetDocument
        Public Property Version As Integer = 1
        Public Property Geometry As String
        Public Property Fingerprint As String
        Public Property OutputTop As Integer
        Public Property OutputBottom As Integer
        Public Property ExtraInputBottom As Integer
        Public Property Periods As String()
        Public Property Nodes As New List(Of BalanceSheetNode)
        Public Property Styles As New List(Of BalanceSheetCellStyle)
        Public Property Diagnostic As String = ""
        Public Function Children(id As String) As List(Of BalanceSheetNode)
            Return Nodes.Where(Function(n) n.ParentId = id).ToList()
        End Function
    End Class

    Public NotInheritable Class BalanceSheetStatement
        Private Const Title As String = "SOFP Outputs - Transactional DB"
        'Relative rows are a versioned schema, not absolute worksheet positions.
        Private Shared ReadOnly Labels As String() = {
            "Fixed Assets", "Other Fixed Assets - Intangible", "Housing Properties Either At Cost Or Valuation",
            "Work In Progress Fixed Assets", "Depreciation", "Investment Properties", "Investment in joint ventures and associates",
            "Other Fixed Assets - Investments", "Homebuy - Loan", "Other Fixed Assets - Tangible", "Less Other Fixed Assets - Tangible Depn",
            "Total Fixed Assets", "", "Current Assets", "Work In Progress - Properties For Sale", "Tenant Debtors", "Other Current Assets",
            "Cash And Bank", "Short Term Investments", "Intercompany Investments", "Total Current Assets", "", "Current Liabilities",
            "Short Term Loans", "Intercompany Loans", "Bank Overdrafts", "Creditors: Amounts Falling Due In One Year", "Creditors: RTB Receipts Due To Council",
            "Deferred Capital Grant: Due Within One Year", "Total Current Liabilities", "", "Net Current Assets", "", "Total Assets Less Current Liabilities",
            "", "Creditors", "Long Term Loans: Amounts Falling Due After More Than One Year", "Capitalised Loan Fees", "Finance Lease Obligations",
            "Fair value derivative financial instruments", "Long Term Creditors: Amounts Falling Due After More Than One Year", "Recycled Grant",
            "Other Capital Grants", "Deferred Capital Grant: Due After More Than One Year", "Amortised Capital Grant", "Homebuy - Grant", "Total Creditors",
            "", "Provisions And Reserves", "Pension Provisions", "Other Provisions", "Income And Expenditure Reserve", "Revaluation reserve",
            "Other Reserves", "Total Provisions And Reservies", "", "Total Financing And Reserves", "", "Check"}
        Private Shared ReadOnly Sections As Integer() = {0, 13, 22, 35, 48}
        Private Shared ReadOnly SectionEnds As Integer() = {11, 20, 29, 46, 54}

        Public Shared Function Read(workbook As IWorkbook) As BalanceSheetDocument
            Dim timer = Abovo.SummitDiagnostics.DiagnosticTimer.StartNew()
            Dim result As New BalanceSheetDocument
            Dim range = workbook.DefinedNames.GetDefinedName(TransactionalDBSnapshotManager.SourceRangeName)?.Range
            If range Is Nothing OrElse range.Worksheet.Name <> "Transactional DB" Then Throw New InvalidOperationException("Transactional_Records is unavailable.")
            Dim sheet = range.Worksheet
            Dim used = sheet.GetUsedRange()
            Dim markers As New List(Of Integer)
            For r = range.BottomRowIndex + 1 To used.BottomRowIndex
                If sheet.Cells(r, 0).Value.ToString().Trim() = Title Then markers.Add(r)
            Next
            If markers.Count <> 1 Then Throw New InvalidOperationException("Balance Sheet requires one recognised SOFP output block; found " & markers.Count.ToString() & ".")
            result.OutputTop = markers(0) + 1
            result.OutputBottom = result.OutputTop + Labels.Length - 1
            result.ExtraInputBottom = range.BottomRowIndex
            result.Geometry = range.GetReferenceA1()
            Dim periods As New List(Of String) From {"Opening balance"}
            For c = 16 To 55
                Dim heading = sheet.Cells(range.TopRowIndex, c).DisplayText
                If Not Regex.IsMatch(heading, "[0-9]{4}/[0-9]{2}") Then Throw New InvalidOperationException("Unsupported Balance Sheet period layout at " & sheet.Cells(range.TopRowIndex, c).GetReferenceA1() & ".")
                periods.Add(heading.Replace(vbCr, "").Replace(vbLf, " "))
            Next
            result.Periods = periods.ToArray()
            Dim contract As New StringBuilder("SOFP-v26-1|" & result.Geometry & "|" & result.OutputTop.ToString() & "|" & String.Join("|", periods))
            Dim styleKeys As New Dictionary(Of String, Integer)
            Dim lineNodes As New Dictionary(Of Integer, BalanceSheetNode)
            For i = 0 To Labels.Length - 1
                Dim row = result.OutputTop + i
                Dim caption = sheet.Cells(row, 1).Value.ToString().Trim()
                If Normal(caption) <> Normal(Labels(i)) Then Throw New InvalidOperationException("Unsupported Balance Sheet layout at " & sheet.Cells(row, 1).GetReferenceA1() & ": expected '" & Labels(i) & "', found '" & caption & "'.")
                contract.Append("|" & caption)
                For p = 0 To 40
                    contract.Append("|" & NormalFormula(sheet.Cells(row, PeriodColumn(p)).FormulaInvariant))
                Next
                If caption.Length = 0 OrElse Sections.Contains(i) Then Continue For
                Dim node As New BalanceSheetNode With {.Id = "bs/" & i.ToString("00"), .Caption = caption,
                    .Source = "'Transactional DB'!" & sheet.Cells(row, 14).GetReferenceA1() & ":" & sheet.Cells(row, 55).GetReferenceA1(),
                    .Rule = sheet.Cells(row, 16).FormulaInvariant, .IsHeadline = True,
                    .IsTotal = SectionEnds.Contains(i) OrElse {31, 33, 56, 58}.Contains(i),
                    .Values = New Double(40) {}, .Styles = New Integer(41) {}}
                node.Styles(0) = CaptureStyle(sheet.Cells(row, 1), result, styleKeys)
                For p = 0 To 40
                    node.Values(p) = Number(sheet.Cells(row, PeriodColumn(p)).Value, False)
                    node.Styles(p + 1) = CaptureStyle(sheet.Cells(row, PeriodColumn(p)), result, styleKeys)
                Next
                lineNodes.Add(i, node)
                If i = 58 Then node.Rule &= " | The workbook Check rounds its result; Trad View may show an unrounded check."
            Next
            result.Fingerprint = Digest(contract.ToString())
            'Each section uses its workbook subtotal as its parent, never a second
            'additive transaction. Standalone totals retain their own workbook values.
            For i = 0 To Labels.Length - 1
                Dim node As BalanceSheetNode = Nothing
                If Not lineNodes.TryGetValue(i, node) Then Continue For
                If SectionEnds.Contains(i) Then Continue For
                Dim lineIndex = i
                Dim sectionIndex = Array.FindIndex(Sections, Function(start) lineIndex > start AndAlso lineIndex < SectionEnds(Array.IndexOf(Sections, start)))
                If sectionIndex >= 0 Then
                    Dim parent = lineNodes(SectionEnds(sectionIndex))
                    If Not result.Nodes.Contains(parent) Then
                        parent.Caption = Labels(Sections(sectionIndex))
                        parent.IsTotal = False
                        result.Nodes.Add(parent)
                    End If
                    node.ParentId = parent.Id
                End If
                result.Nodes.Add(node)
                If Not node.IsTotal Then
                    Dim before = result.Nodes.Count
                    Try
                        AddDetail(workbook, range, row:=result.OutputTop + i, node:=node, document:=result, contract:=contract)
                    Catch ex As InvalidOperationException
                        If result.Nodes.Count > before Then result.Nodes.RemoveRange(before, result.Nodes.Count - before)
                        node.Diagnostic = "Drill-down unavailable: " & ex.Message
                    End Try
                End If
            Next
            For Each sectionEnd In SectionEnds
                Dim parent = lineNodes(sectionEnd)
                Dim children = result.Children(parent.Id)
                CheckReconciliation(parent, children)
            Next
            result.Diagnostic = String.Join(Environment.NewLine, result.Nodes.Where(Function(n) n.Diagnostic.Length > 0).Select(Function(n) n.Caption & ": " & n.Diagnostic))
            result.Fingerprint = Digest(contract.ToString())
            Abovo.SummitDiagnostics.WriteLine("[Balance Sheet] nodes=" & result.Nodes.Count.ToString() & ", total=" & timer.ElapsedMilliseconds.ToString() & " ms")
            Return result
        End Function

        Private Shared Sub AddDetail(workbook As IWorkbook, source As CellRange, row As Integer, node As BalanceSheetNode, document As BalanceSheetDocument, contract As StringBuilder)
            Dim sheet = source.Worksheet
            Dim formula = NormalFormula(sheet.Cells(row, 16).FormulaInvariant)
            Dim spans = Regex.Matches(formula, "Q(?<start>[0-9]+):Q(?<end>[0-9]+)")
            If spans.Count = 0 Then Throw New InvalidOperationException("Unrecognised contribution formula at Q" & (row + 1).ToString() & ".")
            Dim first = Integer.Parse(spans(0).Groups("start").Value) - 1
            Dim last = Integer.Parse(spans(0).Groups("end").Value) - 1
            Dim tailFirst = first, tailLast = last
            Dim family As String
            Dim sign As Integer = 1
            Dim thresholdRow As Integer = -1, pensionRow As Integer = -1
            If formula.Contains(",BF") Then
                family = "reserve"
                Dim threshold = Regex.Match(formula, """<""&BF(?<row>[0-9]+)")
                Dim pension = Regex.Match(formula, ",G(?<row>[0-9]+),Q")
                If Not threshold.Success OrElse Not pension.Success Then Throw New InvalidOperationException("Unsupported reserve criteria.")
                thresholdRow = Integer.Parse(threshold.Groups("row").Value) - 1
                pensionRow = Integer.Parse(pension.Groups("row").Value) - 1
            ElseIf formula.StartsWith("=O" & (row + 1).ToString() & "+SUMIF(F", StringComparison.Ordinal) Then
                family = "cash"
            Else
                family = "heading"
                If spans.Count <> 3 Then Throw New InvalidOperationException("Unsupported heading contribution ranges.")
                tailFirst = Integer.Parse(spans(2).Groups("start").Value) - 1
                tailLast = Integer.Parse(spans(2).Groups("end").Value) - 1
                sign = If(formula.StartsWith("=O" & (row + 1).ToString() & "-", StringComparison.Ordinal), -1, 1)
                If tailFirst <> last + 1 Then Throw New InvalidOperationException("Contribution ranges are not contiguous.")
            End If
            If first <= source.TopRowIndex OrElse tailLast >= document.OutputTop OrElse last >= document.OutputTop Then Throw New InvalidOperationException("Contribution bounds are outside the supported input region.")
            document.ExtraInputBottom = Math.Max(document.ExtraInputBottom, Math.Max(last, tailLast))
            For p = 1 To 40
                Dim col = ColumnName(PeriodColumn(p)), previous = ColumnName(PeriodColumn(p - 1))
                Dim rr = (row + 1).ToString(), a = (first + 1).ToString(), b = (last + 1).ToString()
                Dim expected = "=" & previous & rr
                If family = "cash" Then
                    expected &= "+SUMIF(F" & a & ":F" & b & ",""Cash""," & col & a & ":" & col & b & ")"
                ElseIf family = "reserve" Then
                    expected &= "+SUMIFS(" & col & a & ":" & col & b & ",BF" & a & ":BF" & b & ",""<""&BF" & (thresholdRow + 1).ToString() & ",BQ" & a & ":BQ" & b & ",1)+SUMIF(G" & a & ":G" & b & ",G" & (pensionRow + 1).ToString() & "," & col & a & ":" & col & b & ")"
                Else
                    Dim plus = If(sign = 1, "+", "-"), minus = If(sign = 1, "-", "+")
                    Dim t = (tailFirst + 1).ToString(), z = (tailLast + 1).ToString()
                    expected &= plus & "SUMIFS(" & col & a & ":" & col & b & ",H" & a & ":H" & b & ",H" & rr & ",F" & a & ":F" & b & ",""Cash"")" & minus & "SUMIFS(" & col & a & ":" & col & b & ",H" & a & ":H" & b & ",H" & rr & ",F" & a & ":F" & b & ",""Non Cash"")" & plus & "SUMIF(H" & t & ":H" & z & ",H" & rr & "," & col & t & ":" & col & z & ")"
                End If
                If NormalFormula(expected) <> NormalFormula(sheet.Cells(row, PeriodColumn(p)).FormulaInvariant) Then Throw New InvalidOperationException("Unsupported formula at " & sheet.Cells(row, PeriodColumn(p)).GetReferenceA1() & ".")
            Next
            If Not Regex.IsMatch(sheet.Cells(row, 14).FormulaInvariant, "^='Financial Position - Trad View'!\$?C\$?[0-9]+$", RegexOptions.IgnoreCase) Then Throw New InvalidOperationException("Opening balance is not linked to Trad View.")
            Dim opening = MakeChild(node, "opening", "Opening balance", Enumerable.Repeat(node.Values(0), 41).ToArray())
            opening.Source = sheet.Cells(row, 14).FormulaInvariant
            opening.Rule = "Opening balance carried across all forecast years."
            document.Nodes.Add(opening)
            Dim movements = MakeChild(node, "movements", "Cumulative transaction contributions", New Double(40) {})
            movements.Rule = node.Rule
            document.Nodes.Add(movements)
            Dim groups As New Dictionary(Of String, BalanceSheetNode)(StringComparer.Ordinal)
            Dim mirrors = workbook.DefinedNames.Where(Function(n) n.Name.StartsWith("TransCopy_", StringComparison.OrdinalIgnoreCase) AndAlso n.Range IsNot Nothing AndAlso n.Range.Worksheet Is sheet).OrderBy(Function(n) n.Name, StringComparer.Ordinal).ToList()
            Dim heading = sheet.Cells(row, 7).Value.ToString()
            Dim thresholdValue = If(family = "reserve", Number(sheet.Cells(Math.Max(0, thresholdRow), 57).Value, False), 0)
            If family = "reserve" AndAlso Double.IsNaN(thresholdValue) Then Throw New InvalidOperationException("The SOCI reserve threshold is not numeric.")
            Dim pensionHeading = If(family = "reserve", sheet.Cells(Math.Max(0, pensionRow), 6).Value.ToString(), "")
            contract.Append("|criteria:" & node.Id & ":" & family & ":" & heading & ":" & pensionHeading & ":" & thresholdValue.ToString("R", CultureInfo.InvariantCulture))
            Dim criterion = If(family = "reserve", pensionHeading, heading)
            If family <> "cash" AndAlso criterion.IndexOfAny(New Char() {"*"c, "?"c, "~"c}) >= 0 Then Throw New InvalidOperationException("Wildcard contribution criteria require an explicit model mapping.")
            For r = first To Math.Max(last, tailLast)
                Dim coefficient As Integer = 0
                Dim cash = sheet.Cells(r, 5).Value.ToString()
                If family = "cash" Then
                    If Same(cash, "Cash") Then coefficient = 1
                ElseIf family = "reserve" Then
                    Dim ordinal = Number(sheet.Cells(r, 57).Value, True)
                    If Number(sheet.Cells(r, 68).Value, True) = 1 AndAlso ordinal < thresholdValue Then coefficient += 1
                    If Same(sheet.Cells(r, 6).Value.ToString(), pensionHeading) Then coefficient += 1
                ElseIf Same(sheet.Cells(r, 7).Value.ToString(), heading) Then
                    If r >= tailFirst Then
                        coefficient = sign
                    ElseIf Same(cash, "Cash") Then
                        coefficient = sign
                    ElseIf Same(cash, "Non Cash") Then
                        coefficient = -sign
                    End If
                End If
                If coefficient = 0 Then Continue For
                Dim values(40) As Double
                For p = 1 To 40
                    Dim raw = Number(sheet.Cells(r, PeriodColumn(p)).Value, True)
                    If Double.IsNaN(raw) Then Throw New InvalidOperationException("Nonnumeric contributing value at " & sheet.Cells(r, PeriodColumn(p)).GetReferenceA1() & ".")
                    values(p) = values(p - 1) + coefficient * raw
                    movements.Values(p) += coefficient * raw
                Next
                Dim parent = movements
                Dim lastLabel As String = ""
                For c = 10 To 13
                    Dim label = sheet.Cells(r, c).Value.ToString().Trim()
                    If label.Length = 0 OrElse Same(label, lastLabel) Then Continue For
                    lastLabel = label
                    Dim key = parent.Id & "/level" & c.ToString() & "/" & Digest(label)
                    Dim group As BalanceSheetNode = Nothing
                    If Not groups.TryGetValue(key, group) Then
                        group = MakeChild(parent, "level" & c.ToString() & "/" & Digest(label), label, New Double(40) {})
                        groups.Add(key, group)
                        document.Nodes.Add(group)
                    End If
                    AddValues(group.Values, values)
                    parent = group
                Next
                Dim sourceRow = r
                Dim owner = mirrors.FirstOrDefault(Function(n) sourceRow >= n.Range.TopRowIndex AndAlso sourceRow <= n.Range.BottomRowIndex)
                Dim identity = If(owner Is Nothing, "row/" & r.ToString(), owner.Name & "/" & (r - owner.Range.TopRowIndex).ToString())
                Dim description = sheet.Cells(r, 14).Value.ToString().Trim()
                If description.Length = 0 Then description = "Transaction " & (r + 1).ToString()
                Dim leaf = MakeChild(parent, "record/" & identity, description, values)
                leaf.Source = "'Transactional DB'!" & sheet.Cells(r, 16).GetReferenceA1() & ":" & sheet.Cells(r, 55).GetReferenceA1() & If(owner Is Nothing, " (unmirrored)", " (" & owner.Name & ")")
                leaf.Rule = "Coefficient " & coefficient.ToString() & "; " & family & "; " & sheet.Cells(r, 16).FormulaInvariant
                document.Nodes.Add(leaf)
            Next
            'Movement values above are annual additions; turn them into the same
            'cumulative semantics used by each leaf, group and closing-balance chart.
            For p = 1 To 40
                movements.Values(p) += movements.Values(p - 1)
            Next
            CheckReconciliation(node, New List(Of BalanceSheetNode) From {opening, movements})
            If node.Diagnostic.Length > 0 Then Throw New InvalidOperationException(node.Diagnostic)
        End Sub

        Private Shared Function MakeChild(parent As BalanceSheetNode, suffix As String, caption As String, values As Double()) As BalanceSheetNode
            Return New BalanceSheetNode With {.Id = parent.Id & "/" & suffix, .ParentId = parent.Id, .Caption = caption, .Values = values,
                .Styles = CType(parent.Styles.Clone(), Integer()), .Source = parent.Source, .Rule = parent.Rule}
        End Function

        Private Shared Sub CheckReconciliation(parent As BalanceSheetNode, children As List(Of BalanceSheetNode))
            For p = 0 To 40
                Dim period = p
                Dim sum = children.Sum(Function(n) n.Values(period))
                If Double.IsNaN(sum) OrElse Double.IsNaN(parent.Values(p)) OrElse Math.Abs(sum - parent.Values(p)) > 0.001 Then
                    parent.Diagnostic = "Detail does not reconcile with the workbook headline in period " & p.ToString() & "."
                    Return
                End If
            Next
        End Sub

        Public Shared Function Difference(live As BalanceSheetDocument, snapshot As BalanceSheetDocument) As BalanceSheetDocument
            If live.Fingerprint <> snapshot.Fingerprint Then Throw New InvalidOperationException("Balance Sheet structure has changed since the snapshot. Create a new snapshot.")
            Dim result As New BalanceSheetDocument With {.Geometry = live.Geometry, .Fingerprint = live.Fingerprint,
                .Periods = live.Periods, .OutputTop = live.OutputTop, .OutputBottom = live.OutputBottom, .Styles = live.Styles.ToList()}
            Dim previous = snapshot.Nodes.ToDictionary(Function(n) n.Id, StringComparer.Ordinal)
            Dim current = live.Nodes.ToDictionary(Function(n) n.Id, StringComparer.Ordinal)
            For Each key In live.Nodes.Select(Function(n) n.Id).Union(snapshot.Nodes.Select(Function(n) n.Id), StringComparer.Ordinal)
                Dim a As BalanceSheetNode = Nothing, b As BalanceSheetNode = Nothing
                current.TryGetValue(key, a) : previous.TryGetValue(key, b)
                Dim source = If(a, b)
                Dim values(40) As Double
                For p = 0 To 40
                    values(p) = If(a Is Nothing, 0, a.Values(p)) - If(b Is Nothing, 0, b.Values(p))
                Next
                Dim styles = CType(source.Styles.Clone(), Integer())
                If a Is Nothing Then
                    For p = 0 To styles.Length - 1
                        result.Styles.Add(snapshot.Styles(styles(p)))
                        styles(p) = result.Styles.Count - 1
                    Next
                End If
                result.Nodes.Add(New BalanceSheetNode With {.Id = key, .ParentId = source.ParentId, .Caption = source.Caption,
                    .Source = source.Source, .Rule = source.Rule, .Diagnostic = String.Join(" ", {If(a Is Nothing, "", a.Diagnostic), If(b Is Nothing, "", b.Diagnostic)}).Trim(),
                    .Values = values, .Styles = styles, .IsTotal = source.IsTotal, .IsHeadline = source.IsHeadline})
            Next
            For Each node In result.Nodes.ToArray()
                Dim children = result.Children(node.Id)
                If children.Count > 0 Then CheckReconciliation(node, children)
            Next
            result.Diagnostic = String.Join(Environment.NewLine, result.Nodes.Where(Function(n) n.Diagnostic.Length > 0).Select(Function(n) n.Caption & ": " & n.Diagnostic))
            Return result
        End Function

        Private Shared Sub AddValues(target As Double(), values As Double())
            For p = 0 To target.Length - 1
                target(p) += values(p)
            Next
        End Sub
        Public Shared Function Digest(text As String) As String
            Using hash = SHA256.Create()
                Return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "")
            End Using
        End Function
        Private Shared Function Normal(value As String) As String
            Return Regex.Replace(value.Trim(), "\s+", " ").ToUpperInvariant()
        End Function
        Private Shared Function Same(left As String, right As String) As Boolean
            Return String.Equals(left, right, StringComparison.OrdinalIgnoreCase)
        End Function
        Private Shared Function NormalFormula(value As String) As String
            Return value.Replace("$", "").Replace(" ", "").ToUpperInvariant()
        End Function
        Private Shared Function Number(value As CellValue, blankIsZero As Boolean) As Double
            If value.IsNumeric Then Return value.NumericValue
            If value.IsEmpty AndAlso blankIsZero Then Return 0
            Return Double.NaN
        End Function
        Public Shared Function PeriodColumn(period As Integer) As Integer
            Return If(period = 0, 14, 15 + period)
        End Function
        Private Shared Function ColumnName(index As Integer) As String
            Dim name = ""
            Do
                name = ChrW(65 + index Mod 26) & name
                index = index \ 26 - 1
            Loop While index >= 0
            Return name
        End Function
        Private Shared Function CaptureStyle(cell As Cell, document As BalanceSheetDocument, keys As Dictionary(Of String, Integer)) As Integer
            Dim style As New BalanceSheetCellStyle With {.Format = cell.NumberFormat, .Foreground = cell.Font.Color.ToArgb(),
                .Background = If(cell.Fill.BackgroundColor.IsEmpty, Drawing.Color.White, cell.Fill.BackgroundColor).ToArgb(),
                .FontName = cell.Font.Name, .Bold = cell.Font.Bold, .Italic = cell.Font.Italic,
                .Underline = cell.Font.UnderlineType <> UnderlineType.None, .Alignment = CInt(cell.Alignment.Horizontal)}
            Dim key = String.Join("|", style.Format, style.Foreground, style.Background, style.FontName, style.Bold, style.Italic, style.Underline, style.Alignment)
            Dim index As Integer
            If keys.TryGetValue(key, index) Then Return index
            index = document.Styles.Count
            keys.Add(key, index) : document.Styles.Add(style)
            Return index
        End Function
    End Class
End Namespace
