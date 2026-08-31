Imports System.Diagnostics
Imports System.Globalization
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

        Public Shared Function CreateSnapshotAndComparison(ByVal modelID As Integer) As TimeSpan
            Dim timer As Stopwatch = Stopwatch.StartNew()
            Dim completed As Boolean = False
            Dim processedCellCount As Integer = 0
            Dim snapshotSheet As Worksheet = Nothing
            Dim comparisonSheet As Worksheet = Nothing

            Try
                If FileManager.ExcelModels Is Nothing OrElse
                   modelID < 0 OrElse modelID >= FileManager.ExcelModels.Length OrElse
                   FileManager.ExcelModels(modelID) Is Nothing OrElse
                   FileManager.ExcelModels(modelID).WB Is Nothing Then

                    Throw New InvalidOperationException("The active business-plan workbook is not available.")
                End If

                Dim workbook As IWorkbook = FileManager.ExcelModels(modelID).WB
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

                processedCellCount = sourceRange.RowCount * sourceRange.ColumnCount
                Dim comparisonColumns As Boolean() =
                    GetComparisonValueColumns(sourceRange)

                workbook.BeginUpdate()
                Try
                    snapshotSheet.GetUsedRange().ClearContents()
                    comparisonSheet.GetUsedRange().ClearContents()

                    CreateOrResizeLocalNamedRange(
                        snapshotSheet,
                        SnapshotRangeName,
                        snapshotRange)
                    CreateOrResizeLocalNamedRange(
                        comparisonSheet,
                        ComparisonRangeName,
                        comparisonRange)

                    ''The snapshot is deliberately values-only. The comparison begins
                    ''as the same values so blanks and text remain literal values.
                    snapshotRange.CopyFrom(sourceRange, PasteSpecial.Values)
                    comparisonRange.CopyFrom(snapshotRange, PasteSpecial.Values)

                    Dim absoluteReferenceElements As ReferenceElement =
                        ReferenceElement.ColumnAbsolute Or ReferenceElement.RowAbsolute

                    ''Row zero is the RangeDataSource header. Identity, grouping,
                    ''ordering and UseIn... columns must remain unchanged so the
                    ''comparison can drive the same analyser grids and filters.
                    For rowIndex As Integer = 1 To sourceRange.RowCount - 1
                        For columnIndex As Integer = 0 To sourceRange.ColumnCount - 1
                            If Not comparisonColumns(columnIndex) Then Continue For

                            Dim value As CellValue = sourceRange(rowIndex, columnIndex).Value
                            If Not value.IsNumeric AndAlso Not value.IsDateTime Then Continue For

                            Dim sourceCell As Cell = sourceRange(rowIndex, columnIndex)
                            Dim address As String =
                                sourceCell.GetReferenceA1(absoluteReferenceElements)

                            comparisonRange(rowIndex, columnIndex).FormulaInvariant =
                                "=" & QualifiedCellReference(SourceWorksheetName, address) &
                                "-" & QualifiedCellReference(SnapshotWorksheetName, address)
                        Next
                    Next
                Finally
                    workbook.EndUpdate()
                End Try

                comparisonRange.Calculate()
                FileManager.ExcelModels(modelID).IsDirty = True
                completed = True
                Return timer.Elapsed
            Catch
                ''A failed run must not leave a partially-populated comparison which
                ''could be mistaken for a valid snapshot. These sheets are dedicated
                ''scratch outputs, so a clean blank state is the safe recovery state.
                ClearPartialOutput(snapshotSheet, comparisonSheet)
                Throw
            Finally
                timer.Stop()
                Debug.Print(
                    "Transactional DB snapshot " & If(completed, "completed", "failed") &
                    " in " & timer.Elapsed.TotalSeconds.ToString("0.000", CultureInfo.InvariantCulture) &
                    " seconds for " & processedCellCount.ToString("N0", CultureInfo.InvariantCulture) &
                    " cells.")
            End Try
        End Function

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

            workbook.BeginUpdate()
            Try
                snapshotSheet.GetUsedRange().ClearContents()
                comparisonSheet.GetUsedRange().ClearContents()
            Finally
                workbook.EndUpdate()
            End Try

            FileManager.ExcelModels(modelID).IsDirty = True
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
