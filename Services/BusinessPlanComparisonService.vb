Imports System.Globalization
Imports System.Text.RegularExpressions
Imports DevExpress.Spreadsheet

Namespace Abovo

    Public NotInheritable Class BusinessPlanComparisonParticipant
        Public Property Key As String
        Public Property DisplayName As String
        Public Property FilePath As String
        Public Property Workbook As IWorkbook
        Public Property ModelID As Integer = -1

        Public ReadOnly Property IsOpenSummitModel As Boolean
            Get
                Return ModelID >= 0
            End Get
        End Property
    End Class

    Public NotInheritable Class BusinessPlanComparisonItem
        Public Property ID As Integer
        Public Property ParentID As Integer
        Public Property TargetModel As String
        Public Property Area As String
        Public Property Section As String
        Public Property Item As String
        Public Property Address As String
        Public Property BaseValue As String
        Public Property ComparedValue As String
        Public Property Difference As String
        Public Property Status As String
    End Class

    Public NotInheritable Class BusinessPlanComparisonResult
        Public ReadOnly Property Items As New List(Of BusinessPlanComparisonItem)()
        Public Property DifferenceCount As Integer
        Public Property StructuralIssueCount As Integer
        Public Property CheckIssueCount As Integer
        Public Property Truncated As Boolean
    End Class

    Public NotInheritable Class BusinessPlanMigrationAssessment
        Public ReadOnly Property Lines As New List(Of String)()
        Public Property CanPopulate As Boolean
        Public Property CommonInputCellCount As Integer
        Public Property StructuralMismatchCount As Integer
        Public Property MissingRangeCount As Integer
    End Class

    Public NotInheritable Class BusinessPlanMigrationResult
        Public ReadOnly Property Lines As New List(Of String)()
        Public ReadOnly Property Errors As New List(Of String)()
        Public ReadOnly Property ReportItems As New List(Of BusinessPlanComparisonItem)()
        Public ReadOnly Property ErrorCount As Integer
            Get
                Return Errors.Count
            End Get
        End Property
        Public Property ChangedCellCount As Integer
        Public Property AddedRecordCount As Integer
        Public Property SkippedRangeCount As Integer
        Public Property DetailedSOCIDifferenceCount As Integer
        Public Property CheckIssueCount As Integer
        Public Property Success As Boolean
    End Class

    Friend NotInheritable Class BusinessPlanMigrationRange
        Public Property Key As String
        Public Property RangeName As String
        Public Property WorksheetName As String
        Public Property RuleID As String
        Public Property ReferenceCount As Integer = 1
        Public Property DefaultDataFormat As String
        Public ReadOnly Property CellDataFormats As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Public Function GetDataFormat(ByVal rowIndex As Integer,
                                      ByVal columnIndex As Integer) As String
            Dim result As String = Nothing
            If CellDataFormats.TryGetValue(rowIndex.ToString(CultureInfo.InvariantCulture) & "," &
                                           columnIndex.ToString(CultureInfo.InvariantCulture), result) Then
                Return result
            End If
            Return DefaultDataFormat
        End Function
    End Class

    Friend NotInheritable Class BusinessPlanExpansionRequirement
        Public Property RuleID As String
        Public Property SourceCount As Integer
        Public Property TargetCount As Integer
        Public ReadOnly Property AddCount As Integer
            Get
                Return Math.Max(0, SourceCount - TargetCount)
            End Get
        End Property
    End Class

    Friend NotInheritable Class BusinessPlanMigrationPreflight
        Public ReadOnly Property Ranges As New List(Of BusinessPlanMigrationRange)()
        Public ReadOnly Property Expansions As New List(Of BusinessPlanExpansionRequirement)()
        Public ReadOnly Property Problems As New List(Of String)()
        Public ReadOnly Property Notes As New List(Of String)()
    End Class

    Friend NotInheritable Class BusinessPlanSOCIValue
        Public Property GroupName As String
        Public Property HeadingName As String
        Public Property Level1Name As String
        Public Property Level2Name As String
        Public Property PeriodName As String
        Public Property Value As Double
    End Class

    Public NotInheritable Class BusinessPlanComparisonService

        Private Const MaximumDetailItems As Integer = 25000
        Private Const NumericTolerance As Double = 0.005R

        Private Sub New()
        End Sub

        Public Shared Function Compare(ByVal basePlan As BusinessPlanComparisonParticipant,
                                       ByVal comparisonPlans As IEnumerable(Of BusinessPlanComparisonParticipant),
                                       ByVal assumptionWorksheetNames As HashSet(Of String)) As BusinessPlanComparisonResult
            RequireParticipant(basePlan)
            Dim result As New BusinessPlanComparisonResult()
            Dim nextID As Integer = 1

            For Each comparisonPlan As BusinessPlanComparisonParticipant In comparisonPlans.GroupBy(Function(item) item.Key).Select(Function(group) group.First())
                RequireParticipant(comparisonPlan)
                If String.Equals(comparisonPlan.Key, basePlan.Key, StringComparison.OrdinalIgnoreCase) Then Continue For
                Dim targetRoot As Integer = AddItem(
                    result, nextID, 0, comparisonPlan.FilePath,
                    "Model", comparisonPlan.DisplayName,
                    IO.Path.GetFileName(comparisonPlan.FilePath), String.Empty,
                    basePlan.DisplayName,
                    comparisonPlan.DisplayName,
                    String.Empty, "Compared")

                CompareAssumptions(basePlan, comparisonPlan, assumptionWorksheetNames, targetRoot, result, nextID)
                CompareSOCI(basePlan, comparisonPlan, targetRoot, result, nextID)
                CompareCheckSheet(basePlan, comparisonPlan, targetRoot, result, nextID)
            Next

            Return result
        End Function

        Public Shared Function AssessMigration(ByVal sourcePlan As BusinessPlanComparisonParticipant,
                                               ByVal targetPlan As BusinessPlanComparisonParticipant,
                                               ByVal worksheetNames As HashSet(Of String)) As BusinessPlanMigrationAssessment
            RequireParticipant(sourcePlan)
            RequireParticipant(targetPlan)
            Dim result As New BusinessPlanMigrationAssessment()
            Dim targetRanges As Dictionary(Of String, CellRange) = GetInputNamedRanges(targetPlan.Workbook, worksheetNames)

            For Each pair As KeyValuePair(Of String, CellRange) In targetRanges.OrderBy(Function(entry) entry.Key)
                Dim sourceRange As CellRange = ResolveNamedRange(sourcePlan.Workbook, pair.Key)
                If sourceRange Is Nothing Then
                    result.MissingRangeCount += 1
                    result.Lines.Add("Missing in source: " & pair.Key)
                    Continue For
                End If

                Dim targetRange As CellRange = pair.Value
                If sourceRange.RowCount <> targetRange.RowCount OrElse
                   sourceRange.ColumnCount <> targetRange.ColumnCount Then
                    result.StructuralMismatchCount += 1
                    result.Lines.Add(
                        pair.Key & ": source " & sourceRange.RowCount & " x " & sourceRange.ColumnCount &
                        ", target " & targetRange.RowCount & " x " & targetRange.ColumnCount)
                    Continue For
                End If

                For rowIndex As Integer = 0 To targetRange.RowCount - 1
                    For columnIndex As Integer = 0 To targetRange.ColumnCount - 1
                        Dim targetCell As Cell = targetRange(rowIndex, columnIndex)
                        Dim sourceCell As Cell = sourceRange(rowIndex, columnIndex)
                        If IsInputCell(targetCell) AndAlso IsInputCell(sourceCell) Then
                            result.CommonInputCellCount += 1
                        End If
                    Next
                Next
            Next

            result.CanPopulate = result.StructuralMismatchCount = 0 AndAlso result.MissingRangeCount = 0
            If result.CanPopulate Then
                result.Lines.Insert(0, "All current assumption range geometries match. A write profile can now be validated against this file pair.")
            Else
                result.Lines.Insert(0, "Population is not yet safe: structural mappings or special-case handlers are required.")
            End If
            Return result
        End Function

        Public Shared Function PopulateAssumptions(ByVal sourcePlan As BusinessPlanComparisonParticipant,
                                                   ByVal targetPlan As BusinessPlanComparisonParticipant,
                                                   ByVal worksheetNames As HashSet(Of String),
                                                   Optional ByVal progress As Action(Of String) = Nothing) As BusinessPlanMigrationResult
            RequireParticipant(sourcePlan)
            RequireParticipant(targetPlan)
            If Not targetPlan.IsOpenSummitModel Then
                Throw New InvalidOperationException("The upgrade target must be an open Summit model.")
            End If

            Dim targetModel As FileManager.ExcelModel = FileManager.ExcelModels(targetPlan.ModelID)
            Dim result As New BusinessPlanMigrationResult()
            ReportProgress(progress, "Checking assumption structures...")
            Dim preflight As BusinessPlanMigrationPreflight = BuildMigrationPreflight(sourcePlan, targetPlan, targetModel)
            result.Lines.Add("Source: " & sourcePlan.FilePath)
            result.Lines.Add("Blank target: " & targetPlan.FilePath)
            result.Lines.AddRange(preflight.Notes)
            For Each problem As String In preflight.Problems
                AddMigrationError(result, problem)
            Next

            Dim structuralMutationStarted As Boolean = False
            Try
                For Each expansion As BusinessPlanExpansionRequirement In preflight.Expansions
                    ReportProgress(progress, "Adding " & expansion.AddCount.ToString() & " record(s) for " & expansion.RuleID & "...")
                    Try
                        structuralMutationStarted = True
                        Dim structuralResult As AbovoAppCls.AbovoTransaction =
                            targetModel.WorkbookStructureRules.AddRecords(expansion.RuleID, expansion.AddCount)
                        If structuralResult Is Nothing OrElse structuralResult.BError Then
                            Dim message As String = If(structuralResult Is Nothing,
                                                       "No result was returned.",
                                                       If(String.IsNullOrWhiteSpace(structuralResult.StringReturn),
                                                          structuralResult.StrResponseMessage,
                                                          structuralResult.StringReturn))
                            AddMigrationError(result, "Could not expand '" & expansion.RuleID & "': " & message)
                        Else
                            result.AddedRecordCount += expansion.AddCount
                            result.Lines.Add("Expanded " & expansion.RuleID & " by " & expansion.AddCount.ToString() & " record(s).")
                        End If
                    Catch ex As Exception
                        AddMigrationError(result, "Could not expand '" & expansion.RuleID & "': " & ex.Message)
                    End Try
                Next

                ReportProgress(progress, "Copying writable assumption values...")
                Dim changes As New List(Of DataChangeEvent)()
                Dim changedTargets As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim currentRanges As Dictionary(Of String, BusinessPlanMigrationRange) =
                    GetMigrationRanges(targetModel).ToDictionary(Function(item) item.Key, StringComparer.OrdinalIgnoreCase)
                For Each originalDefinition As BusinessPlanMigrationRange In preflight.Ranges
                    Dim definition As BusinessPlanMigrationRange = Nothing
                    If Not currentRanges.TryGetValue(originalDefinition.Key, definition) Then definition = originalDefinition
                    Dim sourceRange As CellRange = ResolveMigrationRange(sourcePlan.Workbook, definition)
                    Dim targetRange As CellRange = ResolveMigrationRange(targetPlan.Workbook, definition)
                    If sourceRange Is Nothing OrElse targetRange Is Nothing Then
                        result.SkippedRangeCount += 1
                        AddMigrationError(result, "Skipped unavailable range: " & definition.Key)
                        Continue For
                    End If

                    Dim copyRows As Integer = Math.Min(sourceRange.RowCount, targetRange.RowCount)
                    Dim copyColumns As Integer = Math.Min(sourceRange.ColumnCount, targetRange.ColumnCount)
                    If sourceRange.RowCount <> targetRange.RowCount OrElse
                       sourceRange.ColumnCount <> targetRange.ColumnCount Then
                        AddMigrationError(
                            result,
                            "Range '" & definition.Key & "' has different geometry after expansion; copied the common " &
                            copyRows.ToString() & " x " & copyColumns.ToString() & " area (source " &
                            sourceRange.RowCount.ToString() & " x " & sourceRange.ColumnCount.ToString() &
                            ", target " & targetRange.RowCount.ToString() & " x " & targetRange.ColumnCount.ToString() & ").")
                    End If

                    For rowIndex As Integer = 0 To copyRows - 1
                        For columnIndex As Integer = 0 To copyColumns - 1
                            Dim sourceCell As Cell = sourceRange(rowIndex, columnIndex)
                            Dim targetCell As Cell = targetRange(rowIndex, columnIndex)
                            If Not IsInputCell(targetCell) OrElse sourceCell.HasFormula Then Continue For
                            Dim cellKey As String = targetCell.Worksheet.Name & "!" & targetCell.GetReferenceA1()
                            If Not changedTargets.Add(cellKey) Then Continue For
                            If CellValuesEqual(sourceCell, targetCell) Then Continue For
                            Dim typedValue As Object = Nothing
                            Dim dataFormat As String = "S"
                            Dim conversionIssue As String = GetTypedCellValue(
                                sourceCell, definition.GetDataFormat(rowIndex, columnIndex), typedValue, dataFormat)
                            If Not String.IsNullOrWhiteSpace(conversionIssue) Then
                                AddMigrationError(result, cellKey & ": " & conversionIssue & " The source value was copied without type coercion.")
                            End If
                            changes.Add(New DataChangeEvent With {
                                .ModelID = targetPlan.ModelID,
                                .Description = "Upgrade " & targetCell.Worksheet.Name & "!" & targetCell.GetReferenceA1() &
                                               " from " & IO.Path.GetFileName(sourcePlan.FilePath),
                                .WSName = targetCell.Worksheet.Name,
                                .CellAddress = targetCell.GetReferenceA1(),
                                .ChangedValue = typedValue,
                                .DataFormat = dataFormat,
                                .TimeStamp = Now(),
                                .UserName = Environment.UserName})
                        Next
                    Next
                Next

                result.ChangedCellCount = ApplyChangesBestEffort(
                    targetModel, changes, result,
                    "Upgrade assumptions from " & IO.Path.GetFileName(sourcePlan.FilePath))

                ReportProgress(progress, "Calculating source and upgraded target...")
                Try
                    CalculateParticipant(sourcePlan, "Business-plan upgrade source")
                Catch ex As Exception
                    AddMigrationError(result, "Source calculation failed: " & ex.Message)
                End Try
                Try
                    targetModel.WBCalcEngine.CalculateDependencySensitiveFile(
                        "Business-plan upgraded target", Force:=True)
                Catch ex As Exception
                    AddMigrationError(result, "Target calculation failed: " & ex.Message)
                End Try

                ReportProgress(progress, "Comparing the detailed statement and checks...")
                Try
                    Dim comparison As BusinessPlanComparisonResult = Compare(sourcePlan, {targetPlan}, worksheetNames)
                    result.ReportItems.AddRange(comparison.Items)
                    result.DetailedSOCIDifferenceCount = comparison.Items.Where(
                        Function(item) String.Equals(item.Area, "Outputs", StringComparison.OrdinalIgnoreCase) AndAlso
                                       String.Equals(item.Status, "Different", StringComparison.OrdinalIgnoreCase) AndAlso
                                       Not String.IsNullOrWhiteSpace(item.Address)).Count()
                    result.CheckIssueCount = comparison.CheckIssueCount
                Catch ex As Exception
                    AddMigrationError(result, "Post-upgrade comparison failed: " & ex.Message)
                End Try

                result.Lines.Add("Copied " & result.ChangedCellCount.ToString("N0") & " writable assumption cell(s).")
                result.Lines.Add("Detailed SOCI differences: " & result.DetailedSOCIDifferenceCount.ToString("N0") & ".")
                result.Lines.Add("Check Sheet issues: " & result.CheckIssueCount.ToString("N0") & ".")
                AppendMigrationReport(result, sourcePlan, targetPlan)
                result.Success = True
                Return result
            Catch
                If structuralMutationStarted Then
                    ModelSafetyManager.MarkRecoveryRequired(
                        targetPlan.ModelID,
                        "Upgrade assumptions from " & IO.Path.GetFileName(sourcePlan.FilePath),
                        "An unrecoverable migration failure occurred after workbook expansion began.",
                        "Business Plan Upgrade")
                End If
                Throw
            End Try
        End Function

        Private Shared Function ApplyChangesBestEffort(ByVal targetModel As FileManager.ExcelModel,
                                                       ByVal changes As List(Of DataChangeEvent),
                                                       ByVal result As BusinessPlanMigrationResult,
                                                       ByVal description As String) As Integer
            If changes Is Nothing OrElse changes.Count = 0 Then Return 0
            Dim transaction As AbovoAppCls.AbovoTransaction =
                targetModel.ChangeManager.ProcessChanges(changes, description)
            If transaction IsNot Nothing AndAlso Not transaction.BError Then Return changes.Count

            If changes.Count = 1 Then
                Dim change As DataChangeEvent = changes(0)
                Dim message As String = If(transaction Is Nothing,
                                           "No result was returned.",
                                           If(String.IsNullOrWhiteSpace(transaction.StringReturn),
                                              transaction.StrResponseMessage,
                                              transaction.StringReturn))
                AddMigrationError(result,
                                  change.WSName & "!" & change.CellAddress &
                                  " could not be copied: " & message)
                Return 0
            End If

            Dim midpoint As Integer = changes.Count \ 2
            Dim firstHalf As List(Of DataChangeEvent) = changes.Take(midpoint).ToList()
            Dim secondHalf As List(Of DataChangeEvent) = changes.Skip(midpoint).ToList()
            Return ApplyChangesBestEffort(targetModel, firstHalf, result, description) +
                   ApplyChangesBestEffort(targetModel, secondHalf, result, description)
        End Function

        Private Shared Sub AddMigrationError(ByVal result As BusinessPlanMigrationResult,
                                             ByVal message As String)
            If result Is Nothing OrElse String.IsNullOrWhiteSpace(message) Then Return
            If Not result.Errors.Contains(message, StringComparer.OrdinalIgnoreCase) Then
                result.Errors.Add(message)
            End If
        End Sub

        Private Shared Sub AppendMigrationReport(ByVal result As BusinessPlanMigrationResult,
                                                 ByVal sourcePlan As BusinessPlanComparisonParticipant,
                                                 ByVal targetPlan As BusinessPlanComparisonParticipant)
            Dim nextID As Integer = If(result.ReportItems.Count = 0,
                                       1,
                                       result.ReportItems.Max(Function(item) item.ID) + 1)
            Dim rootID As Integer = nextID
            nextID += 1
            result.ReportItems.Add(New BusinessPlanComparisonItem With {
                .ID = rootID,
                .ParentID = 0,
                .TargetModel = targetPlan.FilePath,
                .Area = "Upgrade",
                .Section = "Upgrade report",
                .Item = IO.Path.GetFileName(sourcePlan.FilePath) & " -> " & IO.Path.GetFileName(targetPlan.FilePath),
                .Status = If(result.ErrorCount = 0, "Completed", "Completed with errors")})

            For Each line As String In result.Lines
                result.ReportItems.Add(New BusinessPlanComparisonItem With {
                    .ID = nextID,
                    .ParentID = rootID,
                    .TargetModel = targetPlan.FilePath,
                    .Area = "Upgrade",
                    .Section = "Summary",
                    .Item = line,
                    .Status = "Information"})
                nextID += 1
            Next
            For Each migrationError As String In result.Errors
                result.ReportItems.Add(New BusinessPlanComparisonItem With {
                    .ID = nextID,
                    .ParentID = rootID,
                    .TargetModel = targetPlan.FilePath,
                    .Area = "Upgrade",
                    .Section = "Errors",
                    .Item = migrationError,
                    .Status = "Error"})
                nextID += 1
            Next
        End Sub
        Private Shared Sub CompareAssumptions(ByVal basePlan As BusinessPlanComparisonParticipant,
                                              ByVal comparisonPlan As BusinessPlanComparisonParticipant,
                                              ByVal worksheetNames As HashSet(Of String),
                                              ByVal targetRoot As Integer,
                                              ByVal result As BusinessPlanComparisonResult,
                                              ByRef nextID As Integer)
            Dim areaRoot As Integer = AddItem(result, nextID, targetRoot, comparisonPlan.FilePath,
                                              "Assumptions", "Assumption inputs", String.Empty,
                                              String.Empty, String.Empty, String.Empty, String.Empty, "Compared")
            Dim baseRanges As Dictionary(Of String, CellRange) = GetInputNamedRanges(basePlan.Workbook, worksheetNames)
            Dim worksheetRoots As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Dim comparedCells As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each pair As KeyValuePair(Of String, CellRange) In baseRanges.OrderBy(Function(entry) entry.Key)
                If result.Items.Count >= MaximumDetailItems Then result.Truncated = True : Exit For
                Dim baseRange As CellRange = pair.Value
                Dim comparedRange As CellRange = ResolveNamedRange(comparisonPlan.Workbook, pair.Key)
                Dim worksheetName As String = baseRange.Worksheet.Name
                Dim worksheetRoot As Integer = 0
                If Not worksheetRoots.TryGetValue(worksheetName, worksheetRoot) Then
                    worksheetRoot = AddItem(result, nextID, areaRoot, comparisonPlan.FilePath,
                                            "Assumptions", worksheetName, String.Empty, String.Empty,
                                            String.Empty, String.Empty, String.Empty, "Compared")
                    worksheetRoots.Add(worksheetName, worksheetRoot)
                End If

                If comparedRange Is Nothing Then
                    AddItem(result, nextID, worksheetRoot, comparisonPlan.FilePath,
                            "Assumptions", pair.Key, "Named range missing", pair.Key,
                            baseRange.GetReferenceA1(), "Missing", String.Empty, "Structure")
                    result.StructuralIssueCount += 1
                    Continue For
                End If

                Dim rangeRoot As Integer = 0
                If baseRange.RowCount <> comparedRange.RowCount OrElse
                   baseRange.ColumnCount <> comparedRange.ColumnCount Then
                    rangeRoot = AddItem(result, nextID, worksheetRoot, comparisonPlan.FilePath,
                                        "Assumptions", pair.Key, "Range geometry", pair.Key,
                                        baseRange.RowCount & " x " & baseRange.ColumnCount,
                                        comparedRange.RowCount & " x " & comparedRange.ColumnCount,
                                        String.Empty, "Structure")
                    result.StructuralIssueCount += 1
                End If

                Dim rows As Integer = Math.Min(baseRange.RowCount, comparedRange.RowCount)
                Dim columns As Integer = Math.Min(baseRange.ColumnCount, comparedRange.ColumnCount)
                For rowIndex As Integer = 0 To rows - 1
                    For columnIndex As Integer = 0 To columns - 1
                        Dim baseCell As Cell = baseRange(rowIndex, columnIndex)
                        Dim comparedCell As Cell = comparedRange(rowIndex, columnIndex)
                        Dim cellKey As String = baseCell.Worksheet.Name & "!" & baseCell.GetReferenceA1()
                        If Not comparedCells.Add(cellKey) Then Continue For
                        If Not IsInputCell(baseCell) AndAlso Not IsInputCell(comparedCell) Then Continue For
                        If CellValuesEqual(baseCell, comparedCell) Then Continue For
                        If rangeRoot = 0 Then
                            rangeRoot = AddItem(result, nextID, worksheetRoot, comparisonPlan.FilePath,
                                                "Assumptions", pair.Key, String.Empty, pair.Key,
                                                String.Empty, String.Empty, String.Empty, "Different")
                        End If
                        AddItem(result, nextID, rangeRoot, comparisonPlan.FilePath,
                                "Assumptions", pair.Key,
                                comparedCell.Worksheet.Name & "!" & comparedCell.GetReferenceA1(),
                                comparedCell.Worksheet.Name & "!" & comparedCell.GetReferenceA1(),
                                CellDisplay(baseCell), CellDisplay(comparedCell),
                                NumericDifference(baseCell, comparedCell), "Different")
                        result.DifferenceCount += 1
                        If result.Items.Count >= MaximumDetailItems Then result.Truncated = True : Exit For
                    Next
                    If result.Truncated Then Exit For
                Next
                If result.Truncated Then Exit For
            Next
        End Sub

        Private Shared Sub CompareSOCI(ByVal basePlan As BusinessPlanComparisonParticipant,
                                       ByVal comparisonPlan As BusinessPlanComparisonParticipant,
                                       ByVal targetRoot As Integer,
                                       ByVal result As BusinessPlanComparisonResult,
                                       ByRef nextID As Integer)
            Dim areaRoot As Integer = AddItem(result, nextID, targetRoot, comparisonPlan.FilePath,
                                              "Outputs", "SOCI", "Statement of Comprehensive Income",
                                              String.Empty, String.Empty, String.Empty, String.Empty, "Compared")
            Dim baseValues As Dictionary(Of String, Double)
            Dim comparedValues As Dictionary(Of String, Double)
            Try
                baseValues = BuildSOCIAggregates(basePlan.Workbook)
                comparedValues = BuildSOCIAggregates(comparisonPlan.Workbook)
            Catch ex As Exception
                AddItem(result, nextID, areaRoot, comparisonPlan.FilePath, "Outputs", "SOCI",
                        ex.Message, String.Empty, String.Empty, String.Empty, String.Empty, "Problem")
                result.StructuralIssueCount += 1
                Return
            End Try
            Dim groupRoots As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
            Dim headingRoots As New Dictionary(Of String, Integer)(StringComparer.Ordinal)

            For Each key As String In baseValues.Keys.Union(comparedValues.Keys).OrderBy(Function(entry) entry)
                Dim baseValue As Double = If(baseValues.ContainsKey(key), baseValues(key), 0R)
                Dim comparedValue As Double = If(comparedValues.ContainsKey(key), comparedValues(key), 0R)
                If Math.Abs(baseValue - comparedValue) < NumericTolerance Then Continue For
                Dim parts As String() = key.Split(ChrW(30))
                Dim groupName As String = parts(0)
                Dim headingName As String = parts(1)
                Dim periodName As String = parts(2)
                Dim level1Name As String = If(parts.Length > 3, parts(2), String.Empty)
                Dim level2Name As String = If(parts.Length > 4, parts(3), String.Empty)
                periodName = If(parts.Length > 4, parts(4), parts(2))
                Dim groupRoot As Integer = 0
                If Not groupRoots.TryGetValue(groupName, groupRoot) Then
                    groupRoot = AddItem(result, nextID, areaRoot, comparisonPlan.FilePath,
                                        "Outputs", groupName, String.Empty, String.Empty,
                                        String.Empty, String.Empty, String.Empty, "Different")
                    groupRoots.Add(groupName, groupRoot)
                End If
                Dim headingKey As String = groupName & ChrW(30) & headingName
                Dim headingRoot As Integer = 0
                If Not headingRoots.TryGetValue(headingKey, headingRoot) Then
                    headingRoot = AddItem(result, nextID, groupRoot, comparisonPlan.FilePath,
                                          "Outputs", headingName, String.Empty, String.Empty,
                                          String.Empty, String.Empty, String.Empty, "Different")
                    headingRoots.Add(headingKey, headingRoot)
                End If
                AddItem(result, nextID, headingRoot, comparisonPlan.FilePath,
                        "Outputs", headingName,
                        String.Join(" / ", {level1Name, level2Name}.Where(Function(value) Not String.IsNullOrWhiteSpace(value))),
                        periodName,
                        baseValue.ToString("N0"), comparedValue.ToString("N0"),
                        (comparedValue - baseValue).ToString("N0"), "Different")
                result.DifferenceCount += 1
            Next
        End Sub

        Private Shared Sub CompareCheckSheet(ByVal basePlan As BusinessPlanComparisonParticipant,
                                             ByVal comparisonPlan As BusinessPlanComparisonParticipant,
                                             ByVal targetRoot As Integer,
                                             ByVal result As BusinessPlanComparisonResult,
                                             ByRef nextID As Integer)
            Dim areaRoot As Integer = AddItem(result, nextID, targetRoot, comparisonPlan.FilePath,
                                              "Validation", "Check Sheet", String.Empty, String.Empty,
                                              String.Empty, String.Empty, String.Empty, "Checked")
            AddCheckSheetIssues(basePlan.Workbook, "Base", basePlan.FilePath, areaRoot, result, nextID)
            AddCheckSheetIssues(comparisonPlan.Workbook, "Compared", comparisonPlan.FilePath, areaRoot, result, nextID)
        End Sub

        Private Shared Sub AddCheckSheetIssues(ByVal workbook As IWorkbook,
                                               ByVal label As String,
                                               ByVal fileName As String,
                                               ByVal parentID As Integer,
                                               ByVal result As BusinessPlanComparisonResult,
                                               ByRef nextID As Integer)
            Dim validationRange As CellRange = ResolveNamedRange(workbook, "Outputs_CheckSheet")
            If validationRange Is Nothing OrElse validationRange.ColumnCount < 8 Then
                AddItem(result, nextID, parentID, fileName, "Validation", label,
                        "Outputs_CheckSheet is missing or malformed", String.Empty,
                        String.Empty, String.Empty, String.Empty, "Problem")
                result.CheckIssueCount += 1
                Return
            End If

            Dim issueCount As Integer = 0
            For rowIndex As Integer = 0 To validationRange.RowCount - 1
                Dim statusText As String = validationRange(rowIndex, 4).DisplayText.Trim()
                If String.IsNullOrWhiteSpace(statusText) OrElse
                   String.Equals(statusText, "OK", StringComparison.OrdinalIgnoreCase) Then Continue For
                AddItem(result, nextID, parentID, fileName, "Validation", label,
                        validationRange(rowIndex, 0).DisplayText.Trim(),
                        validationRange(rowIndex, 0).GetReferenceA1(),
                        String.Empty, statusText,
                        validationRange(rowIndex, 5).DisplayText.Trim(), "Problem")
                issueCount += 1
                result.CheckIssueCount += 1
            Next
            If issueCount = 0 Then
                AddItem(result, nextID, parentID, fileName, "Validation", label,
                        "All populated checks are OK", validationRange.GetReferenceA1(),
                        String.Empty, String.Empty, String.Empty, "OK")
            End If
        End Sub

        Private Shared Function BuildSOCIAggregates(ByVal workbook As IWorkbook) As Dictionary(Of String, Double)
            Dim source As CellRange = ResolveNamedRange(workbook, "Transactional_Records")
            If source Is Nothing OrElse source.RowCount < 2 Then
                Throw New InvalidOperationException("Transactional_Records is missing or empty.")
            End If
            Dim headers As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Dim periods As New List(Of Integer)()
            For columnIndex As Integer = 0 To source.ColumnCount - 1
                Dim heading As String = source(0, columnIndex).DisplayText.Trim()
                If Not headers.ContainsKey(heading) Then headers.Add(heading, columnIndex)
                If Regex.IsMatch(heading, "^[0-9]{4}/[0-9]{2}$", RegexOptions.CultureInvariant) Then periods.Add(columnIndex)
            Next
            For Each required As String In {"UseInSOCI", "OrderedSOCIGroup", "OrderedSOCIHeading"}
                If Not headers.ContainsKey(required) Then
                    Throw New InvalidOperationException("Transactional_Records is missing '" & required & "'.")
                End If
            Next
            Dim values As New Dictionary(Of String, Double)(StringComparer.Ordinal)
            For rowIndex As Integer = 1 To source.RowCount - 1
                Dim useCell As Cell = source(rowIndex, headers("UseInSOCI"))
                If Not useCell.Value.IsNumeric OrElse useCell.Value.NumericValue <= 0R Then Continue For
                Dim groupName As String = source(rowIndex, headers("OrderedSOCIGroup")).DisplayText.Trim()
                Dim headingName As String = source(rowIndex, headers("OrderedSOCIHeading")).DisplayText.Trim()
                Dim level1Name As String = If(headers.ContainsKey("Level 1 Copy"), source(rowIndex, headers("Level 1 Copy")).DisplayText.Trim(), String.Empty)
                Dim level2Name As String = If(headers.ContainsKey("Level 2 Copy"), source(rowIndex, headers("Level 2 Copy")).DisplayText.Trim(), String.Empty)
                If String.IsNullOrWhiteSpace(groupName) Then groupName = "Uncategorised"
                If String.IsNullOrWhiteSpace(headingName) Then headingName = "Uncategorised"
                For Each periodIndex As Integer In periods
                    Dim valueCell As Cell = source(rowIndex, periodIndex)
                    If Not valueCell.Value.IsNumeric Then Continue For
                    Dim key As String = String.Join(ChrW(30), {groupName, headingName, level1Name, level2Name,
                                                              source(0, periodIndex).DisplayText.Trim()})
                    If values.ContainsKey(key) Then
                        values(key) += valueCell.Value.NumericValue
                    Else
                        values.Add(key, valueCell.Value.NumericValue)
                    End If
                Next
            Next
            Return values
        End Function

        Private Shared Function BuildMigrationPreflight(ByVal sourcePlan As BusinessPlanComparisonParticipant,
                                                        ByVal targetPlan As BusinessPlanComparisonParticipant,
                                                        ByVal targetModel As FileManager.ExcelModel) As BusinessPlanMigrationPreflight
            Dim result As New BusinessPlanMigrationPreflight()
            result.Ranges.AddRange(GetMigrationRanges(targetModel))
            Dim expansionIndex As New Dictionary(Of String, BusinessPlanExpansionRequirement)(StringComparer.OrdinalIgnoreCase)

            For Each definition As BusinessPlanMigrationRange In result.Ranges
                Dim sourceRange As CellRange = ResolveMigrationRange(sourcePlan.Workbook, definition)
                Dim targetRange As CellRange = ResolveMigrationRange(targetPlan.Workbook, definition)
                If sourceRange Is Nothing Then
                    result.Problems.Add("Not present in source; skipped: " & definition.Key)
                    Continue For
                End If
                If targetRange Is Nothing Then
                    result.Problems.Add("Target range is unavailable: " & definition.Key)
                    Continue For
                End If
                If sourceRange.RowCount = targetRange.RowCount AndAlso
                   sourceRange.ColumnCount = targetRange.ColumnCount Then Continue For

                If TargetRangeHasContents(targetRange) Then
                    result.Problems.Add(
                        definition.Key & " has populated target cells but incompatible geometry (source " &
                        sourceRange.RowCount.ToString() & " x " & sourceRange.ColumnCount.ToString() &
                        ", target " & targetRange.RowCount.ToString() & " x " & targetRange.ColumnCount.ToString() & ").")
                    Continue For
                End If
                If String.IsNullOrWhiteSpace(definition.RuleID) Then
                    result.Problems.Add(
                        definition.Key & " needs resizing but has no declared structural rule (source " &
                        sourceRange.RowCount.ToString() & " x " & sourceRange.ColumnCount.ToString() &
                        ", target " & targetRange.RowCount.ToString() & " x " & targetRange.ColumnCount.ToString() & ").")
                    Continue For
                End If

                Dim sourceCount As Integer = targetModel.WorkbookStructureRules.GetRecordCount(definition.RuleID, sourcePlan.Workbook)
                Dim targetCount As Integer = targetModel.WorkbookStructureRules.GetRecordCount(definition.RuleID, targetPlan.Workbook)
                If sourceCount < 0 OrElse targetCount < 0 Then
                    result.Problems.Add("Could not determine the record count for " & definition.RuleID & " (" & definition.Key & ").")
                    Continue For
                End If
                If sourceCount < targetCount Then
                    result.Problems.Add(
                        definition.Key & " is smaller in the source; automatic population never deletes target structure (source " &
                        sourceCount.ToString() & ", target " & targetCount.ToString() & ").")
                    Continue For
                End If
                If sourceCount = targetCount Then
                    result.Problems.Add(
                        definition.Key & " has different geometry although " & definition.RuleID &
                        " has the same record count in both files.")
                    Continue For
                End If

                Dim existing As BusinessPlanExpansionRequirement = Nothing
                If expansionIndex.TryGetValue(definition.RuleID, existing) Then
                    If existing.SourceCount <> sourceCount OrElse existing.TargetCount <> targetCount Then
                        result.Problems.Add("Conflicting record counts were found for " & definition.RuleID & ".")
                    End If
                Else
                    existing = New BusinessPlanExpansionRequirement With {
                        .RuleID = definition.RuleID,
                        .SourceCount = sourceCount,
                        .TargetCount = targetCount}
                    expansionIndex.Add(definition.RuleID, existing)
                    result.Expansions.Add(existing)
                End If
            Next

            Dim referenceCount As Integer = result.Ranges.Sum(Function(item) item.ReferenceCount)
            result.Notes.Insert(0,
                                "Unique writable Structure.xml assumption ranges considered: " &
                                result.Ranges.Count.ToString("N0") &
                                " (" & referenceCount.ToString("N0") & " XML reference(s)).")
            Return result
        End Function

        Private Shared Function GetMigrationRanges(ByVal targetModel As FileManager.ExcelModel) As List(Of BusinessPlanMigrationRange)
            Dim ranges As New Dictionary(Of String, BusinessPlanMigrationRange)(StringComparer.OrdinalIgnoreCase)
            If targetModel Is Nothing OrElse targetModel.WBStructure Is Nothing OrElse
               targetModel.WBStructure.GroupStructures Is Nothing Then Return ranges.Values.ToList()

            For Each group As GroupStructure In targetModel.WBStructure.GroupStructures
                If group Is Nothing OrElse Not String.Equals(group.GSName, "Assumptions", StringComparison.OrdinalIgnoreCase) OrElse
                   group.ChildStructures Is Nothing Then Continue For
                For Each child As ChildStructure In group.ChildStructures
                    If child Is Nothing OrElse child.InterfaceSections Is Nothing Then Continue For
                    For Each section As CSInterfaceSection In child.InterfaceSections
                        If section Is Nothing OrElse section.ISDatasources Is Nothing Then Continue For
                        For Each dataSource As ISEDatasource In section.ISDatasources
                            If dataSource Is Nothing OrElse IsConfiguredTrue(dataSource.RO) Then Continue For
                            Dim ruleID As String = dataSource.StructureRuleID
                            If String.IsNullOrWhiteSpace(ruleID) Then
                                ruleID = targetModel.WorkbookStructureRules.ResolveRuleID(dataSource.RowExpandByNR)
                            End If
                            AddMigrationNamedRange(ranges, targetModel.WB, dataSource.RowExpandByNR, Nothing, ruleID)

                            If dataSource.NamedRangeSources IsNot Nothing Then
                                For Each namedSource As NamedRangeDataSource In dataSource.NamedRangeSources
                                    If namedSource Is Nothing Then Continue For
                                    AddMigrationNamedRange(ranges, targetModel.WB, namedSource.NRName, Nothing, ruleID,
                                                           namedSource.DataFieldDefinitions)
                                    AddMigrationNamedRange(ranges, targetModel.WB, namedSource.DefinedBy, Nothing, ruleID)
                                    AddMigrationNamedRange(ranges, targetModel.WB, namedSource.ExpandsBy, Nothing, ruleID)
                                    If namedSource.DataFieldDefinitions IsNot Nothing Then
                                        For Each field As DataFieldDefinition In namedSource.DataFieldDefinitions
                                            AddMigrationFieldRanges(ranges, targetModel.WB, field, Nothing, ruleID)
                                        Next
                                    End If
                                Next
                            End If

                            If dataSource.CellRangeSources Is Nothing Then Continue For
                            For Each cellSource As CellRangeDataSource In dataSource.CellRangeSources
                                If cellSource Is Nothing OrElse IsConfiguredTrue(cellSource.RO) OrElse
                                   IsConfiguredTrue(cellSource.IsCalculated) Then Continue For
                                AddMigrationNamedRange(ranges, targetModel.WB, cellSource.DataRange, cellSource.WSName, ruleID,
                                                       cellSource.DataFieldDefinitions)
                                AddMigrationNamedRange(ranges, targetModel.WB, cellSource.RowsDefinedByData, cellSource.WSName, ruleID)
                                AddMigrationNamedRange(ranges, targetModel.WB, cellSource.ColsDefinedByData, cellSource.WSName, ruleID)
                                AddMigrationNamedRange(ranges, targetModel.WB, cellSource.DataRangeExtensionData, cellSource.WSName, ruleID)
                                AddMigrationNamedRange(ranges, targetModel.WB, cellSource.OffSetNR, cellSource.WSName, ruleID)
                                If cellSource.DataFieldDefinitions Is Nothing Then Continue For
                                For Each field As DataFieldDefinition In cellSource.DataFieldDefinitions
                                    AddMigrationFieldRanges(ranges, targetModel.WB, field, cellSource.WSName, ruleID)
                                Next
                            Next
                        Next
                    Next
                Next
            Next
            Return ranges.Values.OrderBy(Function(item) item.Key).ToList()
        End Function

        Private Shared Sub AddMigrationFieldRanges(ByVal ranges As Dictionary(Of String, BusinessPlanMigrationRange),
                                                   ByVal workbook As IWorkbook,
                                                   ByVal field As DataFieldDefinition,
                                                   ByVal worksheetName As String,
                                                   ByVal ruleID As String,
                                                  Optional ByVal fieldDefinitions As IEnumerable(Of DataFieldDefinition) = Nothing,
                                                  Optional ByVal defaultDataFormat As String = Nothing)
            If field Is Nothing OrElse IsConfiguredTrue(field.RO) OrElse IsConfiguredTrue(field.IsColCalculated) Then Return
            Dim editDataFormat As String = If(String.IsNullOrWhiteSpace(field.EditRepNRHereDataFormat),
                                               field.DataFormat,
                                               field.EditRepNRHereDataFormat)
            AddMigrationNamedRange(ranges, workbook, field.RepeatingNR, worksheetName, ruleID,
                                   Nothing, editDataFormat)
            AddMigrationNamedRange(ranges, workbook, field.EditRepNRHere, worksheetName, ruleID,
                                   Nothing, editDataFormat)
        End Sub

        Private Shared Sub AddMigrationNamedRange(ByVal ranges As Dictionary(Of String, BusinessPlanMigrationRange),
                                                  ByVal workbook As IWorkbook,
                                                  ByVal candidate As String,
                                                  ByVal worksheetName As String,
                                                  ByVal ruleID As String,
                                                  Optional ByVal fieldDefinitions As IEnumerable(Of DataFieldDefinition) = Nothing,
                                                  Optional ByVal defaultDataFormat As String = Nothing)
            If String.IsNullOrWhiteSpace(candidate) Then Return
            Dim rangeName As String = candidate.Trim()
            Dim resolvedRange As CellRange = ResolveNamedRange(workbook, rangeName)
            If resolvedRange Is Nothing Then Return
            Dim hasWritableInput As Boolean = False
            For rowIndex As Integer = 0 To resolvedRange.RowCount - 1
                For columnIndex As Integer = 0 To resolvedRange.ColumnCount - 1
                    If IsInputCell(resolvedRange(rowIndex, columnIndex)) Then hasWritableInput = True : Exit For
                Next
                If hasWritableInput Then Exit For
            Next
            If Not hasWritableInput Then Return
            Dim existing As BusinessPlanMigrationRange = Nothing
            If ranges.TryGetValue(rangeName, existing) Then
                existing.ReferenceCount += 1
                If String.IsNullOrWhiteSpace(existing.RuleID) AndAlso Not String.IsNullOrWhiteSpace(ruleID) Then existing.RuleID = ruleID
                MergeMigrationDataFormats(existing, resolvedRange, workbook, worksheetName,
                                          fieldDefinitions, defaultDataFormat)
                Return
            End If
            Dim definition As New BusinessPlanMigrationRange With {
                .Key = rangeName,
                .RangeName = rangeName,
                .WorksheetName = worksheetName,
                .RuleID = ruleID}
            MergeMigrationDataFormats(definition, resolvedRange, workbook, worksheetName,
                                      fieldDefinitions, defaultDataFormat)
            ranges.Add(rangeName, definition)
        End Sub

        Private Shared Sub MergeMigrationDataFormats(ByVal definition As BusinessPlanMigrationRange,
                                                     ByVal resolvedRange As CellRange,
                                                     ByVal workbook As IWorkbook,
                                                     ByVal worksheetName As String,
                                                     ByVal fieldDefinitions As IEnumerable(Of DataFieldDefinition),
                                                     ByVal defaultDataFormat As String)
            If definition Is Nothing OrElse resolvedRange Is Nothing Then Return
            If Not String.IsNullOrWhiteSpace(defaultDataFormat) AndAlso
               String.IsNullOrWhiteSpace(definition.DefaultDataFormat) Then
                definition.DefaultDataFormat = defaultDataFormat.Trim()
            End If
            If fieldDefinitions Is Nothing Then Return

            Dim formats As List(Of String) = BuildFieldDataFormats(fieldDefinitions, workbook, worksheetName)
            If formats.Count = 0 Then Return
            If formats.Count = 1 Then
                If String.IsNullOrWhiteSpace(definition.DefaultDataFormat) Then
                    definition.DefaultDataFormat = formats(0)
                End If
                Return
            End If

            If formats.Count = resolvedRange.RowCount Then
                For rowIndex As Integer = 0 To formats.Count - 1
                    If String.IsNullOrWhiteSpace(formats(rowIndex)) Then Continue For
                    For columnIndex As Integer = 0 To resolvedRange.ColumnCount - 1
                        MergeCellDataFormat(definition, rowIndex, columnIndex, formats(rowIndex))
                    Next
                Next
            ElseIf formats.Count = resolvedRange.ColumnCount Then
                For columnIndex As Integer = 0 To formats.Count - 1
                    If String.IsNullOrWhiteSpace(formats(columnIndex)) Then Continue For
                    For rowIndex As Integer = 0 To resolvedRange.RowCount - 1
                        MergeCellDataFormat(definition, rowIndex, columnIndex, formats(columnIndex))
                    Next
                Next
            End If
        End Sub

        Private Shared Function BuildFieldDataFormats(ByVal fieldDefinitions As IEnumerable(Of DataFieldDefinition),
                                                      ByVal workbook As IWorkbook,
                                                      ByVal worksheetName As String) As List(Of String)
            Dim formats As New List(Of String)()
            For Each field As DataFieldDefinition In fieldDefinitions
                If field Is Nothing Then Continue For
                Dim format As String = If(IsConfiguredTrue(field.RO) OrElse
                                           IsConfiguredTrue(field.IsColCalculated) OrElse
                                           IsConfiguredTrue(field.IsDummy),
                                           Nothing,
                                           field.DataFormat)
                Dim repeatCount As Integer = 1
                If IsConfiguredTrue(field.RepeatsByNR) Then
                    Dim repeatRange As CellRange = ResolveNamedRange(workbook, field.RepeatingNR)
                    If repeatRange IsNot Nothing Then repeatCount = Math.Max(repeatRange.RowCount, repeatRange.ColumnCount)
                ElseIf IsConfiguredTrue(field.RepeatsByCR) AndAlso
                       Not String.IsNullOrWhiteSpace(worksheetName) AndAlso
                       Not String.IsNullOrWhiteSpace(field.RepeatsByCRData) Then
                    Try
                        Dim repeatRange As CellRange = workbook.Worksheets(worksheetName.Trim()).Range(field.RepeatsByCRData.Trim())
                        repeatCount = Math.Max(repeatRange.RowCount, repeatRange.ColumnCount)
                    Catch
                        repeatCount = 1
                    End Try
                End If
                repeatCount += ParseOptionalInteger(field.RepeatingPreRows) + ParseOptionalInteger(field.RepeatingPostRows)
                For index As Integer = 1 To Math.Max(1, repeatCount)
                    formats.Add(format)
                Next
            Next
            Return formats
        End Function

        Private Shared Sub MergeCellDataFormat(ByVal definition As BusinessPlanMigrationRange,
                                               ByVal rowIndex As Integer,
                                               ByVal columnIndex As Integer,
                                               ByVal dataFormat As String)
            If String.IsNullOrWhiteSpace(dataFormat) Then Return
            Dim key As String = rowIndex.ToString(CultureInfo.InvariantCulture) & "," &
                                columnIndex.ToString(CultureInfo.InvariantCulture)
            If Not definition.CellDataFormats.ContainsKey(key) Then
                definition.CellDataFormats.Add(key, dataFormat.Trim())
            End If
        End Sub

        Private Shared Function ParseOptionalInteger(ByVal value As String) As Integer
            Dim parsed As Integer
            If Integer.TryParse(If(value, String.Empty).Trim(), NumberStyles.Integer,
                                CultureInfo.InvariantCulture, parsed) Then Return parsed
            Return 0
        End Function
        Private Shared Function ResolveMigrationRange(ByVal workbook As IWorkbook,
                                                      ByVal definition As BusinessPlanMigrationRange) As CellRange
            If definition Is Nothing Then Return Nothing
            Return ResolveNamedRange(workbook, definition.RangeName)
        End Function

        Private Shared Function TargetRangeHasContents(ByVal targetRange As CellRange) As Boolean
            If targetRange Is Nothing Then Return False
            For rowIndex As Integer = 0 To targetRange.RowCount - 1
                For columnIndex As Integer = 0 To targetRange.ColumnCount - 1
                    Dim cell As Cell = targetRange(rowIndex, columnIndex)
                    If IsInputCell(cell) AndAlso Not cell.Value.IsEmpty Then Return True
                Next
            Next
            Return False
        End Function

        Private Shared Function GetTypedCellValue(ByVal sourceCell As Cell,
                                                  ByVal declaredDataFormat As String,
                                                  ByRef value As Object,
                                                  ByRef dataFormat As String) As String
            dataFormat = If(declaredDataFormat, String.Empty).Trim().ToUpperInvariant()
            If sourceCell Is Nothing OrElse sourceCell.Value.IsEmpty Then
                value = Nothing
                If String.IsNullOrWhiteSpace(dataFormat) Then dataFormat = "S"
                Return Nothing
            End If

            If String.IsNullOrWhiteSpace(dataFormat) Then
                If sourceCell.Value.IsDateTime Then
                    value = sourceCell.Value.DateTimeValue
                    dataFormat = "D"
                ElseIf sourceCell.Value.IsBoolean Then
                    value = sourceCell.Value.BooleanValue
                    dataFormat = "BOOL"
                ElseIf sourceCell.Value.IsNumeric Then
                    value = sourceCell.Value.NumericValue
                    dataFormat = "N"
                ElseIf sourceCell.Value.IsText Then
                    value = sourceCell.Value.TextValue
                    dataFormat = "S"
                Else
                    value = sourceCell.DisplayText
                    dataFormat = "S"
                End If
                Return Nothing
            End If

            Select Case dataFormat
                Case "S", "FL", "DUMMY"
                    value = If(sourceCell.Value.IsText, sourceCell.Value.TextValue, sourceCell.DisplayText)
                Case "BOOL", "BOOLEAN", "B"
                    If sourceCell.Value.IsBoolean Then
                        value = sourceCell.Value.BooleanValue
                    ElseIf sourceCell.Value.IsNumeric Then
                        value = sourceCell.Value.NumericValue
                    Else
                        Dim textValue As String = sourceCell.DisplayText.Trim()
                        Dim parsedBoolean As Boolean
                        Dim parsedNumber As Double
                        If Boolean.TryParse(textValue, parsedBoolean) OrElse
                           textValue.Equals("YES", StringComparison.OrdinalIgnoreCase) OrElse
                           textValue.Equals("NO", StringComparison.OrdinalIgnoreCase) OrElse
                           textValue.Equals("Y", StringComparison.OrdinalIgnoreCase) OrElse
                           textValue.Equals("N", StringComparison.OrdinalIgnoreCase) OrElse
                           Double.TryParse(textValue, NumberStyles.Any, CultureInfo.CurrentCulture, parsedNumber) OrElse
                           Double.TryParse(textValue, NumberStyles.Any, CultureInfo.InvariantCulture, parsedNumber) Then
                            value = textValue
                        Else
                            value = sourceCell.DisplayText
                            dataFormat = "S"
                            Return "The value does not match declared type '" & declaredDataFormat & "'."
                        End If
                    End If
                Case "I", "Y"
                    If sourceCell.Value.IsNumeric Then
                        value = sourceCell.Value.NumericValue
                        If Math.Abs(sourceCell.Value.NumericValue - Math.Truncate(sourceCell.Value.NumericValue)) > NumericTolerance Then
                            dataFormat = "N"
                            Return "The value is not a whole number required by declared type '" & declaredDataFormat & "'."
                        End If
                    Else
                        Dim parsedInteger As Integer
                        If Integer.TryParse(sourceCell.DisplayText, NumberStyles.Integer Or NumberStyles.AllowThousands,
                                            CultureInfo.CurrentCulture, parsedInteger) OrElse
                           Integer.TryParse(sourceCell.DisplayText, NumberStyles.Integer Or NumberStyles.AllowThousands,
                                            CultureInfo.InvariantCulture, parsedInteger) Then
                            value = parsedInteger
                        Else
                            value = sourceCell.DisplayText
                            dataFormat = "S"
                            Return "The value does not match declared type '" & declaredDataFormat & "'."
                        End If
                    End If
                Case "N", "P", "C", "M", "SM", "R"
                    If sourceCell.Value.IsNumeric Then
                        value = sourceCell.Value.NumericValue
                    Else
                        Dim parsedNumber As Double
                        If Double.TryParse(sourceCell.DisplayText, NumberStyles.Any, CultureInfo.CurrentCulture, parsedNumber) OrElse
                           Double.TryParse(sourceCell.DisplayText, NumberStyles.Any, CultureInfo.InvariantCulture, parsedNumber) Then
                            value = parsedNumber
                        Else
                            value = sourceCell.DisplayText
                            dataFormat = "S"
                            Return "The value does not match declared type '" & declaredDataFormat & "'."
                        End If
                    End If
                Case "D", "DM"
                    If sourceCell.Value.IsDateTime Then
                        value = sourceCell.Value.DateTimeValue
                    ElseIf sourceCell.Value.IsNumeric Then
                        value = sourceCell.Value.NumericValue
                    Else
                        Dim parsedDate As DateTime
                        Dim parsedSerial As Double
                        If DateTime.TryParse(sourceCell.DisplayText, CultureInfo.CurrentCulture,
                                             DateTimeStyles.AllowWhiteSpaces, parsedDate) OrElse
                           DateTime.TryParse(sourceCell.DisplayText, CultureInfo.InvariantCulture,
                                             DateTimeStyles.AllowWhiteSpaces, parsedDate) Then
                            value = parsedDate
                        ElseIf Double.TryParse(sourceCell.DisplayText, NumberStyles.Any,
                                               CultureInfo.CurrentCulture, parsedSerial) OrElse
                               Double.TryParse(sourceCell.DisplayText, NumberStyles.Any,
                                               CultureInfo.InvariantCulture, parsedSerial) Then
                            value = parsedSerial
                        Else
                            value = sourceCell.DisplayText
                            dataFormat = "S"
                            Return "The value does not match declared type '" & declaredDataFormat & "'."
                        End If
                    End If
                Case Else
                    value = If(sourceCell.Value.IsText, sourceCell.Value.TextValue, sourceCell.DisplayText)
            End Select
            Return Nothing
        End Function
        Private Shared Sub CalculateParticipant(ByVal participant As BusinessPlanComparisonParticipant,
                                                ByVal reason As String)
            If participant.IsOpenSummitModel Then
                FileManager.ExcelModels(participant.ModelID).WBCalcEngine.CalculateDependencySensitiveFile(reason, Force:=True)
                Return
            End If
            Dim previousEngine As CalculationEngineType = participant.Workbook.Options.CalculationEngineType
            Try
                participant.Workbook.Options.CalculationEngineType = CalculationEngineType.Recursive
                participant.Workbook.CalculateFull()
            Finally
                participant.Workbook.Options.CalculationEngineType = previousEngine
            End Try
        End Sub

        Private Shared Sub ReportProgress(ByVal progress As Action(Of String), ByVal message As String)
            If progress IsNot Nothing Then progress(message)
        End Sub

        Private Shared Function IsConfiguredTrue(ByVal value As String) As Boolean
            If String.IsNullOrWhiteSpace(value) Then Return False
            Return String.Equals(value.Trim(), "TRUE", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(value.Trim(), "YES", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(value.Trim(), "1", StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Function GetAssumptionWorksheetNames(ByVal model As FileManager.ExcelModel) As HashSet(Of String)
            Dim names As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If model Is Nothing OrElse model.WBStructure Is Nothing OrElse model.WBStructure.GroupStructures Is Nothing Then Return names
            For Each group As GroupStructure In model.WBStructure.GroupStructures
                If group Is Nothing OrElse Not String.Equals(group.GSName, "Assumptions", StringComparison.OrdinalIgnoreCase) Then Continue For
                For Each child As ChildStructure In group.ChildStructures
                    If child Is Nothing Then Continue For
                    If Not String.IsNullOrWhiteSpace(child.DefaultWorksheet) Then names.Add(child.DefaultWorksheet.Trim())
                    If child.InterfaceSections Is Nothing Then Continue For
                    For Each section As CSInterfaceSection In child.InterfaceSections
                        If section Is Nothing OrElse section.ISDatasources Is Nothing Then Continue For
                        For Each dataSource As ISEDatasource In section.ISDatasources
                            If dataSource Is Nothing OrElse dataSource.CellRangeSources Is Nothing Then Continue For
                            For Each cellSource As CellRangeDataSource In dataSource.CellRangeSources
                                If cellSource IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(cellSource.WSName) Then names.Add(cellSource.WSName.Trim())
                            Next
                        Next
                    Next
                Next
            Next
            Return names
        End Function

        Private Shared Function GetInputNamedRanges(ByVal workbook As IWorkbook,
                                                    ByVal worksheetNames As HashSet(Of String)) As Dictionary(Of String, CellRange)
            Dim ranges As New Dictionary(Of String, CellRange)(StringComparer.OrdinalIgnoreCase)
            For Each definedName As DefinedName In workbook.DefinedNames
                If definedName Is Nothing Then Continue For
                Dim sourceRange As CellRange = Nothing
                Try
                    sourceRange = definedName.Range
                Catch
                    Continue For
                End Try
                If sourceRange Is Nothing Then Continue For
                If Not worksheetNames.Contains(sourceRange.Worksheet.Name) Then Continue For
                Dim hasInput As Boolean = False
                For rowIndex As Integer = 0 To sourceRange.RowCount - 1
                    For columnIndex As Integer = 0 To sourceRange.ColumnCount - 1
                        If IsInputCell(sourceRange(rowIndex, columnIndex)) Then hasInput = True : Exit For
                    Next
                    If hasInput Then Exit For
                Next
                If hasInput AndAlso Not ranges.ContainsKey(definedName.Name) Then ranges.Add(definedName.Name, sourceRange)
            Next
            Return ranges
        End Function

        Private Shared Function IsInputCell(ByVal cell As Cell) As Boolean
            Return cell IsNot Nothing AndAlso Not cell.Protection.Locked AndAlso Not cell.HasFormula
        End Function

        Private Shared Function CellValuesEqual(ByVal left As Cell, ByVal right As Cell) As Boolean
            If left.Value.IsEmpty AndAlso right.Value.IsEmpty Then Return True
            If left.Value.IsNumeric AndAlso right.Value.IsNumeric Then
                Return Math.Abs(left.Value.NumericValue - right.Value.NumericValue) < NumericTolerance
            End If
            If left.Value.IsDateTime AndAlso right.Value.IsDateTime Then Return left.Value.DateTimeValue = right.Value.DateTimeValue
            If left.Value.IsBoolean AndAlso right.Value.IsBoolean Then Return left.Value.BooleanValue = right.Value.BooleanValue
            Return String.Equals(CellDisplay(left), CellDisplay(right), StringComparison.Ordinal)
        End Function

        Private Shared Function CellDisplay(ByVal cell As Cell) As String
            If cell Is Nothing OrElse cell.Value.IsEmpty Then Return String.Empty
            Return cell.DisplayText
        End Function

        Private Shared Function NumericDifference(ByVal baseCell As Cell, ByVal comparedCell As Cell) As String
            If baseCell.Value.IsNumeric AndAlso comparedCell.Value.IsNumeric Then
                Return (comparedCell.Value.NumericValue - baseCell.Value.NumericValue).ToString("N2")
            End If
            Return String.Empty
        End Function

        Private Shared Function ResolveNamedRange(ByVal workbook As IWorkbook, ByVal rangeName As String) As CellRange
            If workbook Is Nothing OrElse String.IsNullOrWhiteSpace(rangeName) Then Return Nothing
            Try
                Dim definedName As DefinedName = workbook.DefinedNames.GetDefinedName(rangeName)
                Return If(definedName Is Nothing, Nothing, definedName.Range)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub RequireParticipant(ByVal participant As BusinessPlanComparisonParticipant)
            If participant Is Nothing OrElse participant.Workbook Is Nothing Then
                Throw New ArgumentException("The selected business plan is no longer available.", NameOf(participant))
            End If
        End Sub

        Private Shared Function AddItem(ByVal result As BusinessPlanComparisonResult,
                                        ByRef nextID As Integer,
                                        ByVal parentID As Integer,
                                        ByVal targetModel As String,
                                        ByVal area As String,
                                        ByVal section As String,
                                        ByVal item As String,
                                        ByVal address As String,
                                        ByVal baseValue As String,
                                        ByVal comparedValue As String,
                                        ByVal difference As String,
                                        ByVal status As String) As Integer
            Dim currentID As Integer = nextID
            nextID += 1
            result.Items.Add(New BusinessPlanComparisonItem With {
                .ID = currentID, .ParentID = parentID, .TargetModel = targetModel,
                .Area = area, .Section = section, .Item = item, .Address = address,
                .BaseValue = baseValue, .ComparedValue = comparedValue,
                .Difference = difference, .Status = status})
            Return currentID
        End Function

    End Class

End Namespace
