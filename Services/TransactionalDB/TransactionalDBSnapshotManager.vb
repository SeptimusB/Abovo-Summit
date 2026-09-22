Imports System.Text.RegularExpressions
Imports DevExpress.Spreadsheet

Namespace Abovo

    Public NotInheritable Class TransactionalDBSnapshotManager

        Public Const SourceWorksheetName As String = "Transactional DB"
        Public Const SourceRangeName As String = "Transactional_Records"
        Public Const SnapshotWorksheetName As String = "TDB Snapshot"
        Public Const ComparisonWorksheetName As String = "TDB Comparison"
        Public Const SnapshotRangeName As String = "TDB_Snapshot_Records"
        Public Const ComparisonRangeName As String = "TDB_Comparison_Records"

        Private Sub New()
        End Sub

        Public Shared Sub CreateSnapshotAndComparison(ByVal modelID As Integer)
            Dim snapshotSheet As Worksheet = Nothing
            Dim comparisonSheet As Worksheet = Nothing
            Dim benchmark As System.Diagnostics.Stopwatch =
                System.Diagnostics.Stopwatch.StartNew()
            Dim rangeRows As Integer = 0
            Dim rangeColumns As Integer = 0
            Dim comparisonFormulaCount As Integer = 0
            Dim setupMs As Long = 0
            Dim clearMs As Long = 0
            Dim nameMs As Long = 0
            Dim copyMs As Long = 0
            Dim formulaMs As Long = 0
            Dim endUpdateMs As Long = 0
            Dim calculateMs As Long = 0
            Dim verifyMs As Long = 0
            Dim outcome As String = "failed"
            Dim mutationStarted As Boolean = False
            Dim protectedSheets As New Dictionary(Of Worksheet, WorksheetProtectionPermissions)

            Try
                If FileManager.ExcelModels Is Nothing OrElse
                   modelID < 0 OrElse modelID >= FileManager.ExcelModels.Length OrElse
                   FileManager.ExcelModels(modelID) Is Nothing OrElse
                   FileManager.ExcelModels(modelID).WB Is Nothing Then

                    Throw New InvalidOperationException("The active business-plan workbook is not available.")
                End If

                Dim workbook As IWorkbook = FileManager.ExcelModels(modelID).WB
                FileManager.ExcelModels(modelID).EnsureDeferredSaveResultsCurrent("Preparing snapshot results...")
                Dim sourceSheet As Worksheet = RequireWorksheet(workbook, SourceWorksheetName)
                snapshotSheet = RequireWorksheet(workbook, SnapshotWorksheetName)
                comparisonSheet = RequireWorksheet(workbook, ComparisonWorksheetName)

                Dim sourceName As DefinedName = sourceSheet.DefinedNames.GetDefinedName(SourceRangeName)
                If sourceName Is Nothing Then sourceName = workbook.DefinedNames.GetDefinedName(SourceRangeName)
                If sourceName Is Nothing OrElse sourceName.Range Is Nothing Then
                    Throw New InvalidOperationException(
                        "The named range ''" & SourceRangeName & "'' was not found.")
                End If

                Dim sourceRange As CellRange = sourceName.Range
                rangeRows = sourceRange.RowCount
                rangeColumns = sourceRange.ColumnCount
                If Not String.Equals(
                    sourceRange.Worksheet.Name,
                    SourceWorksheetName,
                    StringComparison.OrdinalIgnoreCase) Then

                    Throw New InvalidOperationException(
                        "The named range ''" & SourceRangeName & "'' does not refer to the ''" &
                        SourceWorksheetName & "'' worksheet.")
                End If

                Dim snapshotRange As CellRange =
                    CreateMatchingLocalRange(snapshotSheet, sourceRange)
                Dim comparisonRange As CellRange =
                    CreateMatchingLocalRange(comparisonSheet, sourceRange)

                Dim comparisonColumns As Boolean() =
                    GetComparisonValueColumns(sourceRange)
                BalanceSheetSnapshot.ValidateDestination(snapshotSheet, SnapshotRangeName)
                BalanceSheetSnapshot.ValidateDestination(comparisonSheet, ComparisonRangeName)
                Dim balanceSheet As BalanceSheetDocument = Nothing
                Try
                    balanceSheet = BalanceSheetStatement.Read(workbook)
                Catch ex As InvalidOperationException
                    'Unsupported/bespoke layouts must not disable existing SOCI/CF
                    'snapshots. Their BS page reports the unsupported mapping.
                    Diagnostics.Trace.WriteLine("[Balance Sheet snapshot] Not captured: " & ex.Message)
                End Try
                For Each sheet In {snapshotSheet, comparisonSheet}
                    If sheet.IsProtected Then
                        protectedSheets.Add(sheet, sheet.GetProtectionPermissions())
                        WSSecurity.UNProtectWS(modelID, sheet.Name)
                        If sheet.IsProtected Then Throw New InvalidOperationException("Cannot write the protected snapshot worksheet.")
                    End If
                Next
                setupMs = benchmark.ElapsedMilliseconds

                workbook.BeginUpdate()
                Try
                    Dim phaseStartMs As Long = benchmark.ElapsedMilliseconds
                    mutationStarted = True
                    FileManager.ExcelModels(modelID).RequireFullRebuild()
                    snapshotSheet.GetUsedRange().ClearContents()
                    comparisonSheet.GetUsedRange().ClearContents()
                    clearMs = benchmark.ElapsedMilliseconds - phaseStartMs

                    phaseStartMs = benchmark.ElapsedMilliseconds
                    CreateOrResizeLocalNamedRange(
                        snapshotSheet,
                        SnapshotRangeName,
                        snapshotRange)
                    CreateOrResizeLocalNamedRange(
                        comparisonSheet,
                        ComparisonRangeName,
                        comparisonRange)
                    nameMs = benchmark.ElapsedMilliseconds - phaseStartMs

                    ''The snapshot is deliberately values-only. The comparison begins
                    ''as the same values so blanks and text remain literal values.
                    phaseStartMs = benchmark.ElapsedMilliseconds
                    snapshotRange.CopyFrom(sourceRange, PasteSpecial.Values)
                    comparisonRange.CopyFrom(snapshotRange, PasteSpecial.Values)
                    copyMs = benchmark.ElapsedMilliseconds - phaseStartMs

                    ''Row zero is the RangeDataSource header. Identity, grouping,
                    ''ordering and UseIn... columns must remain unchanged so the
                    ''comparison can drive the same analyser grids and filters.
                    phaseStartMs = benchmark.ElapsedMilliseconds
                    Dim absoluteReferenceElements As ReferenceElement =
                        ReferenceElement.ColumnAbsolute Or ReferenceElement.RowAbsolute
                    For rowIndex As Integer = 1 To sourceRange.RowCount - 1
                        For columnIndex As Integer = 0 To sourceRange.ColumnCount - 1
                            If Not comparisonColumns(columnIndex) Then Continue For

                            Dim value As CellValue =
                                sourceRange(rowIndex, columnIndex).Value
                            If Not IsComparisonValue(value) Then Continue For

                            Dim sourceCell As Cell =
                                sourceRange(rowIndex, columnIndex)
                            Dim address As String =
                                sourceCell.GetReferenceA1(absoluteReferenceElements)

                            comparisonRange(rowIndex, columnIndex).FormulaInvariant =
                                "=" & QualifiedCellReference(SourceWorksheetName, address) &
                                "-" & QualifiedCellReference(SnapshotWorksheetName, address)
                            comparisonFormulaCount += 1
                        Next
                    Next
                    formulaMs = benchmark.ElapsedMilliseconds - phaseStartMs
                    If balanceSheet IsNot Nothing Then BalanceSheetSnapshot.Capture(workbook, balanceSheet)
                Finally
                    Dim phaseStartMs As Long = benchmark.ElapsedMilliseconds
                    workbook.EndUpdate()
                    endUpdateMs = benchmark.ElapsedMilliseconds - phaseStartMs
                End Try

                Dim calculateStartMs As Long = benchmark.ElapsedMilliseconds
                comparisonRange.Calculate()
                If balanceSheet IsNot Nothing Then comparisonSheet.DefinedNames.GetDefinedName(BalanceSheetSnapshot.ExtraRangeName).Range.Calculate()
                calculateMs = benchmark.ElapsedMilliseconds - calculateStartMs
                Dim verifyStartMs As Long = benchmark.ElapsedMilliseconds
                VerifyComparisonValues(
                    sourceRange,
                    snapshotRange,
                    comparisonRange,
                    comparisonColumns)
                If balanceSheet IsNot Nothing Then
                    BalanceSheetSnapshot.Read(workbook, balanceSheet)
                    For Each node In balanceSheet.Nodes.Where(Function(n) n.IsHeadline)
                        Dim row = balanceSheet.OutputTop + Integer.Parse(node.Id.Substring(3))
                        For p = 0 To 40
                            Dim actual = comparisonSheet.Cells(row, BalanceSheetStatement.PeriodColumn(p)).Value
                            If Not Double.IsNaN(node.Values(p)) AndAlso (Not actual.IsNumeric OrElse Math.Abs(actual.NumericValue) > 0.001) Then Throw New InvalidOperationException("Balance Sheet comparison did not initialise to zero.")
                        Next
                    Next
                End If
                verifyMs = benchmark.ElapsedMilliseconds - verifyStartMs
                FileManager.ExcelModels(modelID).IsDirty = True
                outcome = "ok"
            Catch
                ''A failed run must not leave a partially-populated comparison which
                ''could be mistaken for a valid snapshot. These sheets are dedicated
                ''scratch outputs, so a clean blank state is the safe recovery state.
                If mutationStarted Then
                    ClearPartialOutput(snapshotSheet, comparisonSheet)
                    FileManager.ExcelModels(modelID).IsDirty = True
                End If
                Throw
            Finally
                For Each entry In protectedSheets
                    WSSecurity.ProtectWS(modelID, entry.Key.Name, entry.Value)
                Next
                Dim measuredMs As Long =
                    setupMs + clearMs + nameMs + copyMs + formulaMs +
                    endUpdateMs + calculateMs + verifyMs
                System.Diagnostics.Trace.WriteLine(
                    "[Snapshot Benchmark] model=" & modelID.ToString() &
                    ", rows=" & rangeRows.ToString() &
                    ", columns=" & rangeColumns.ToString() &
                    ", comparisonFormulas=" & comparisonFormulaCount.ToString() &
                    ", setup=" & setupMs.ToString() & " ms" &
                    ", clear=" & clearMs.ToString() & " ms" &
                    ", names=" & nameMs.ToString() & " ms" &
                    ", copy=" & copyMs.ToString() & " ms" &
                    ", formulas=" & formulaMs.ToString() & " ms" &
                    ", endUpdate=" & endUpdateMs.ToString() & " ms" &
                    ", calculate=" & calculateMs.ToString() & " ms" &
                    ", verify=" & verifyMs.ToString() & " ms" &
                    ", other=" &
                    Math.Max(0, benchmark.ElapsedMilliseconds - measuredMs).ToString() & " ms" &
                    ", total=" & benchmark.ElapsedMilliseconds.ToString() & " ms" &
                    ", outcome=" & outcome)
            End Try
        End Sub

        Private Shared Function IsComparisonValue(ByVal value As CellValue) As Boolean
            Return value.IsNumeric OrElse value.IsDateTime
        End Function

        Private Shared Sub VerifyComparisonValues(
            ByVal sourceRange As CellRange,
            ByVal snapshotRange As CellRange,
            ByVal comparisonRange As CellRange,
            ByVal comparisonColumns As Boolean())

            Const tolerance As Double = 0.0000001R
            For columnIndex As Integer = 0 To sourceRange.ColumnCount - 1
                If Not comparisonColumns(columnIndex) Then Continue For

                For rowIndex As Integer = 1 To sourceRange.RowCount - 1
                    Dim sourceValue As CellValue =
                        sourceRange(rowIndex, columnIndex).Value
                    If Not IsComparisonValue(sourceValue) Then Continue For

                    Dim comparisonCell As Cell =
                        comparisonRange(rowIndex, columnIndex)
                    If Not comparisonCell.HasFormula Then
                        Throw New InvalidOperationException(
                            "The comparison formula was not created at " &
                            comparisonCell.GetReferenceA1() & ".")
                    End If

                    If Not sourceValue.IsNumeric Then Continue For
                    Dim snapshotValue As CellValue =
                        snapshotRange(rowIndex, columnIndex).Value
                    Dim comparisonValue As CellValue = comparisonCell.Value
                    If Not snapshotValue.IsNumeric OrElse
                       Not comparisonValue.IsNumeric OrElse
                       Math.Abs(
                           comparisonValue.NumericValue -
                           (sourceValue.NumericValue - snapshotValue.NumericValue)) >
                       tolerance Then
                        Throw New InvalidOperationException(
                            "The comparison value is invalid at " &
                            comparisonCell.GetReferenceA1() & ".")
                    End If
                Next
            Next
        End Sub

        Public Shared Function HasValidSnapshot(ByVal modelID As Integer) As Boolean
            Try
                If FileManager.ExcelModels Is Nothing OrElse
                   modelID < 0 OrElse modelID >= FileManager.ExcelModels.Length OrElse
                   FileManager.ExcelModels(modelID) Is Nothing OrElse
                   FileManager.ExcelModels(modelID).WB Is Nothing Then Return False

                Dim workbook As IWorkbook = FileManager.ExcelModels(modelID).WB
                Dim sourceSheet As Worksheet = RequireWorksheet(workbook, SourceWorksheetName)
                Dim snapshotSheet As Worksheet = RequireWorksheet(workbook, SnapshotWorksheetName)
                Dim comparisonSheet As Worksheet = RequireWorksheet(workbook, ComparisonWorksheetName)
                Dim sourceRange As CellRange = ResolveNamedRange(workbook, sourceSheet, SourceRangeName)
                Dim snapshotRange As CellRange = ResolveNamedRange(workbook, snapshotSheet, SnapshotRangeName)
                Dim comparisonRange As CellRange = ResolveNamedRange(workbook, comparisonSheet, ComparisonRangeName)

                If sourceRange Is Nothing OrElse
                   snapshotRange Is Nothing OrElse
                   comparisonRange Is Nothing OrElse
                   sourceRange.Worksheet Is Nothing OrElse
                   snapshotRange.Worksheet Is Nothing OrElse
                   comparisonRange.Worksheet Is Nothing OrElse
                   Not String.Equals(sourceRange.Worksheet.Name, SourceWorksheetName, StringComparison.OrdinalIgnoreCase) OrElse
                   Not String.Equals(snapshotRange.Worksheet.Name, SnapshotWorksheetName, StringComparison.OrdinalIgnoreCase) OrElse
                   Not String.Equals(comparisonRange.Worksheet.Name, ComparisonWorksheetName, StringComparison.OrdinalIgnoreCase) OrElse
                   Not HasMatchingGeometry(sourceRange, snapshotRange) OrElse
                   Not HasMatchingGeometry(sourceRange, comparisonRange) Then Return False

                For columnIndex As Integer = 0 To sourceRange.ColumnCount - 1
                    If Not String.Equals(
                        sourceRange(0, columnIndex).Value.ToString(),
                        snapshotRange(0, columnIndex).Value.ToString(),
                        StringComparison.Ordinal) Then Return False

                    If Not String.Equals(
                        sourceRange(0, columnIndex).Value.ToString(),
                        comparisonRange(0, columnIndex).Value.ToString(),
                        StringComparison.Ordinal) Then Return False
                Next

                Return True
            Catch
                Return False
            End Try
        End Function

        Public Shared Function HasPersistedSnapshot(ByVal modelID As Integer) As Boolean
            Try
                If FileManager.ExcelModels Is Nothing OrElse
                   modelID < 0 OrElse modelID >= FileManager.ExcelModels.Length OrElse
                   FileManager.ExcelModels(modelID) Is Nothing OrElse
                   FileManager.ExcelModels(modelID).WB Is Nothing Then Return False

                Dim workbook As IWorkbook = FileManager.ExcelModels(modelID).WB
                Dim snapshotSheet As Worksheet = RequireWorksheet(workbook, SnapshotWorksheetName)
                Dim comparisonSheet As Worksheet = RequireWorksheet(workbook, ComparisonWorksheetName)
                Dim snapshotRange As CellRange = ResolveNamedRange(workbook, snapshotSheet, SnapshotRangeName)
                Dim comparisonRange As CellRange = ResolveNamedRange(workbook, comparisonSheet, ComparisonRangeName)

                If snapshotRange Is Nothing OrElse
                   comparisonRange Is Nothing OrElse
                   snapshotRange.Worksheet Is Nothing OrElse
                   comparisonRange.Worksheet Is Nothing OrElse
                   Not String.Equals(snapshotRange.Worksheet.Name, SnapshotWorksheetName, StringComparison.OrdinalIgnoreCase) OrElse
                   Not String.Equals(comparisonRange.Worksheet.Name, ComparisonWorksheetName, StringComparison.OrdinalIgnoreCase) OrElse
                   Not HasMatchingGeometry(snapshotRange, comparisonRange) Then Return False

                Dim HasHeader As Boolean = False

                For columnIndex As Integer = 0 To snapshotRange.ColumnCount - 1
                    Dim snapshotHeader As String = snapshotRange(0, columnIndex).Value.ToString()
                    Dim comparisonHeader As String = comparisonRange(0, columnIndex).Value.ToString()

                    If Not String.Equals(snapshotHeader, comparisonHeader, StringComparison.Ordinal) Then Return False
                    If Not String.IsNullOrWhiteSpace(snapshotHeader) Then HasHeader = True
                Next

                Return HasHeader
            Catch
                Return False
            End Try
        End Function

        Public Shared Sub InvalidateSnapshot(ByVal modelID As Integer)
            If FileManager.ExcelModels Is Nothing OrElse
               modelID < 0 OrElse modelID >= FileManager.ExcelModels.Length OrElse
               FileManager.ExcelModels(modelID) Is Nothing OrElse
               FileManager.ExcelModels(modelID).WB Is Nothing Then Return

            Dim workbook As IWorkbook = FileManager.ExcelModels(modelID).WB
            Dim snapshotSheet As Worksheet = Nothing
            Dim comparisonSheet As Worksheet = Nothing

            Try
                snapshotSheet = RequireWorksheet(workbook, SnapshotWorksheetName)
                comparisonSheet = RequireWorksheet(workbook, ComparisonWorksheetName)
            Catch
                Return
            End Try

            Dim protectedSheets As New Dictionary(Of Worksheet, WorksheetProtectionPermissions)
            For Each sheet In {snapshotSheet, comparisonSheet}
                If sheet.IsProtected Then protectedSheets.Add(sheet, sheet.GetProtectionPermissions())
            Next
            workbook.BeginUpdate()
            Try
                For Each sheet In protectedSheets.Keys
                    WSSecurity.UNProtectWS(modelID, sheet.Name)
                Next
                'Snapshot validity is persisted by the two local named-range
                'headers. Clearing only those headers invalidates both analyser
                'datasources immediately and after save/reopen without paying to
                'clear hundreds of thousands of obsolete body cells. The next
                'snapshot creation clears and replaces both dedicated sheets.
                ClearSnapshotHeader(workbook, snapshotSheet, SnapshotRangeName)
                ClearSnapshotHeader(workbook, comparisonSheet, ComparisonRangeName)
                BalanceSheetSnapshot.Invalidate(workbook)
            Finally
                Try
                    workbook.EndUpdate()
                Finally
                    For Each entry In protectedSheets
                        WSSecurity.ProtectWS(modelID, entry.Key.Name, entry.Value)
                    Next
                End Try
            End Try

            FileManager.ExcelModels(modelID).IsDirty = True
        End Sub

        Private Shared Sub ClearSnapshotHeader(ByVal workbook As IWorkbook,
                                               ByVal worksheet As Worksheet,
                                               ByVal rangeName As String)
            Dim snapshotRange As CellRange =
                ResolveNamedRange(workbook, worksheet, rangeName)

            If snapshotRange Is Nothing OrElse snapshotRange.RowCount = 0 Then Return

            worksheet.Range.FromLTRB(
                snapshotRange.LeftColumnIndex,
                snapshotRange.TopRowIndex,
                snapshotRange.RightColumnIndex,
                snapshotRange.TopRowIndex).ClearContents()
        End Sub

        Private Shared Function RequireWorksheet(ByVal workbook As IWorkbook,
                                                 ByVal worksheetName As String) As Worksheet
            For Each worksheet As Worksheet In workbook.Worksheets
                If String.Equals(
                    worksheet.Name,
                    worksheetName,
                    StringComparison.OrdinalIgnoreCase) Then Return worksheet
            Next

            Throw New InvalidOperationException(
                "The worksheet ''" & worksheetName & "'' was not found.")
        End Function

        Private Shared Function CreateMatchingLocalRange(ByVal targetSheet As Worksheet,
                                                         ByVal sourceRange As CellRange) As CellRange
            Return targetSheet.Range.FromLTRB(
                sourceRange.LeftColumnIndex,
                sourceRange.TopRowIndex,
                sourceRange.RightColumnIndex,
                sourceRange.BottomRowIndex)
        End Function

        Private Shared Function ResolveNamedRange(ByVal workbook As IWorkbook,
                                                  ByVal worksheet As Worksheet,
                                                  ByVal rangeName As String) As CellRange
            Dim definedName As DefinedName = worksheet.DefinedNames.GetDefinedName(rangeName)
            If definedName Is Nothing Then definedName = workbook.DefinedNames.GetDefinedName(rangeName)
            If definedName Is Nothing Then Return Nothing
            Return definedName.Range
        End Function

        Private Shared Function HasMatchingGeometry(ByVal sourceRange As CellRange,
                                                    ByVal targetRange As CellRange) As Boolean
            Return sourceRange.LeftColumnIndex = targetRange.LeftColumnIndex AndAlso
                   sourceRange.TopRowIndex = targetRange.TopRowIndex AndAlso
                   sourceRange.RightColumnIndex = targetRange.RightColumnIndex AndAlso
                   sourceRange.BottomRowIndex = targetRange.BottomRowIndex
        End Function

        Private Shared Function GetComparisonValueColumns(ByVal sourceRange As CellRange) As Boolean()
            Dim result(sourceRange.ColumnCount - 1) As Boolean
            Dim firstPeriodColumn As Integer = -1

            For columnIndex As Integer = 0 To sourceRange.ColumnCount - 1
                If Not IsPeriodHeading(sourceRange(0, columnIndex).DisplayText) Then Continue For
                result(columnIndex) = True
                If firstPeriodColumn < 0 Then firstPeriodColumn = columnIndex
            Next

            If firstPeriodColumn < 0 Then
                Throw New InvalidOperationException(
                    "The Transactional_Records range does not contain any recognised period columns.")
            End If

            ''The analyser defines the column immediately before the first period
            ''as Opening Balance and displays it only on the balance-sheet view.
            If firstPeriodColumn > 0 Then result(firstPeriodColumn - 1) = True
            Return result
        End Function

        Private Shared Function IsPeriodHeading(ByVal candidate As String) As Boolean
            If String.IsNullOrWhiteSpace(candidate) Then Return False

            Dim lines As String() =
                candidate.Replace(vbCrLf, vbLf).
                          Replace(vbCr, vbLf).
                          Split(New String() {vbLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each line As String In lines
                If Regex.IsMatch(
                    line.Trim(),
                    "^[0-9]{4}/[0-9]{2}$",
                    RegexOptions.CultureInvariant) Then Return True
            Next

            Return False
        End Function

        Private Shared Sub CreateOrResizeLocalNamedRange(ByVal worksheet As Worksheet,
                                                        ByVal rangeName As String,
                                                        ByVal targetRange As CellRange)
            Dim definedName As DefinedName = worksheet.DefinedNames.GetDefinedName(rangeName)

            If definedName Is Nothing Then
                Dim elements As ReferenceElement =
                    ReferenceElement.IncludeSheetName Or
                    ReferenceElement.ColumnAbsolute Or
                    ReferenceElement.RowAbsolute

                worksheet.DefinedNames.Add(
                    rangeName,
                    targetRange.GetReferenceA1(elements))
            Else
                definedName.Range = targetRange
            End If
        End Sub

        Private Shared Function QualifiedCellReference(ByVal worksheetName As String,
                                                       ByVal address As String) As String
            Return "'" & worksheetName.Replace("'", "''") & "'!" & address
        End Function

        Private Shared Sub ClearPartialOutput(ByVal snapshotSheet As Worksheet,
                                             ByVal comparisonSheet As Worksheet)
            Try
                If snapshotSheet IsNot Nothing Then snapshotSheet.GetUsedRange().ClearContents()
                If comparisonSheet IsNot Nothing Then comparisonSheet.GetUsedRange().ClearContents()
            Catch
                ''Preserve the original exception; the next run starts by clearing
                ''both sheets again before publishing a new snapshot.
            End Try
        End Sub

    End Class

End Namespace
