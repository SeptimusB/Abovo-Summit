Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports Abovo.FileManager
Imports Abovo.AbovoAppCls
Imports DevExpress.Spreadsheet

Namespace Abovo

    Public Class TransactionalDBSynchroniser

        Private ReadOnly ModelID As Integer
        Private ReadOnly SyncRules As List(Of TransactionalDBSyncRule)
        Private IsSynchronising As Boolean = False

        Private Class TransactionalDBSyncRule

            Public SourceNamedRange As String
            Public SourceColumnAdjustment As Integer
            Public TargetNamedRanges As String()

            Public Sub New(ByVal SetSourceNamedRange As String,
                           ByVal SetSourceColumnAdjustment As Integer,
                           ParamArray SetTargetNamedRanges() As String)

                SourceNamedRange = SetSourceNamedRange
                SourceColumnAdjustment = SetSourceColumnAdjustment
                TargetNamedRanges = SetTargetNamedRanges

            End Sub

        End Class

        Private Class MirrorResizeRequest

            Public TargetNamedRange As String
            Public RequiredRows As Integer

        End Class

        Private Class MirrorResizeWorkItem

            Public Request As MirrorResizeRequest
            Public WorksheetName As String
            Public TopRowIndex As Integer
            Public BottomRowIndex As Integer
            Public LeftColumnIndex As Integer
            Public RightColumnIndex As Integer
            Public CurrentRows As Integer

        End Class

        Public Sub New(ByVal SetModelID As Integer)

            ModelID = SetModelID
            SyncRules = CreateCompatibilityRules()

        End Sub

        Private Function GetWorkbook() As IWorkbook

            If ExcelModels Is Nothing Then Return Nothing
            If ModelID < 0 OrElse ModelID >= ExcelModels.Length Then Return Nothing
            If ExcelModels(ModelID) Is Nothing Then Return Nothing

            Return ExcelModels(ModelID).WB

        End Function

        Private Function CreateCompatibilityRules() As List(Of TransactionalDBSyncRule)

            Dim Rules As New List(Of TransactionalDBSyncRule)

            '=========================================================================
            ' These rules mirror the Summit_Compatibility VBA module in the XLSB.
            '
            ' SourceColumnAdjustment corresponds to VBA expressions such as:
            '   Range("HouseTypeInID").Columns.Count - 1
            '
            ' Every TransCopy range contains one additional final/footer row, so the
            ' required mirror row count is SourceColumnCount + Adjustment + 1.
            '=========================================================================

            Rules.Add(New TransactionalDBSyncRule(
                      "HouseTypeInID", -1,
                      "TransCopy_DevptSingle_A", "TransCopy_DevptSingle_B", "TransCopy_DevptSingle_C",
                      "TransCopy_DevptSingle_D", "TransCopy_DevptSingle_E", "TransCopy_DevptSingle_F",
                      "TransCopy_DevptSingle_G", "TransCopy_DevptSingle_H", "TransCopy_DevptSingle_I",
                      "TransCopy_DevptSingle_J", "TransCopy_DevptSingle_K", "TransCopy_DevptSingle_L",
                      "TransCopy_DevptSingle_M", "TransCopy_DevptSingle_N"))

            Rules.Add(New TransactionalDBSyncRule(
                      "HouseTypeInMY", -1,
                      "TransCopy_DevptMulti_A", "TransCopy_DevptMulti_B", "TransCopy_DevptMulti_C",
                      "TransCopy_DevptMulti_D", "TransCopy_DevptMulti_E", "TransCopy_DevptMulti_F",
                      "TransCopy_DevptMulti_G", "TransCopy_DevptMulti_H", "TransCopy_DevptMulti_I",
                      "TransCopy_DevptMulti_J", "TransCopy_DevptMulti_K", "TransCopy_DevptMulti_L",
                      "TransCopy_DevptMulti_M", "TransCopy_DevptMulti_N"))

            Rules.Add(New TransactionalDBSyncRule(
                      "IC_JointVenture_01", 0,
                      "TransCopy_JV_01", "TransCopy_JV_02", "TransCopy_JV_03", "TransCopy_JV_04"))

            Rules.Add(New TransactionalDBSyncRule(
                      "CapGrantInclusion", 0,
                      "TransCopy_CapGrantAssumpt", "TransCopy_CapGrant_01"))

            Rules.Add(New TransactionalDBSyncRule(
                      "StockCondCats", 0,
                      "TransCopy_RepMainAssumpt_A", "TransCopy_RepMainAssumpt_B",
                      "TransCopy_RepMainAssumpt_C", "TransCopy_RepMainAssumpt_D",
                      "TransCopy_RepMainAssumpt_E", "TransCopy_RepMainAssumpt_F",
                      "TransCopy_RepMainAssumpt_G", "TransCopy_RepMainAssumpt_H"))

            Rules.Add(New TransactionalDBSyncRule(
                      "Rep_CapExpend_010", -1,
                      "TransCopy_CapEx1"))

            Rules.Add(New TransactionalDBSyncRule(
                      "FacilityNames", 0,
                      "TransCopy_FacilityNames_A", "TransCopy_FacilityNames_B", "TransCopy_FacilityNames_C",
                      "TransCopy_FacilityNames_D", "TransCopy_FacilityNames_E", "TransCopy_FacilityNames_F",
                      "TransCopy_FacilityNames_G", "TransCopy_FacilityNames_H", "TransCopy_FacilityNames_I"))

            Rules.Add(New TransactionalDBSyncRule(
                      "LoanDescsOrd", 0,
                      "TransCopy_LoanDescsOrd_A", "TransCopy_LoanDescsOrd_B"))

            Rules.Add(New TransactionalDBSyncRule(
                      "IC_IntercoFunding_01", 0,
                      "TransCopy_IC_IntercoFunding_01_A", "TransCopy_IC_IntercoFunding_01_B",
                      "TransCopy_IC_IntercoFunding_01_C"))

            Rules.Add(New TransactionalDBSyncRule(
                      "IC_IntercoFunding_02", 0,
                      "TransCopy_IC_IntercoFunding_02_A", "TransCopy_IC_IntercoFunding_02_B",
                      "TransCopy_IC_IntercoFunding_02_C"))

            Rules.Add(New TransactionalDBSyncRule(
                      "DepnType", 0,
                      "TransCopy_HAComponents_A", "TransCopy_HAComponents_B", "TransCopy_HAComponents_C",
                      "TransCopy_HAComponents_D", "TransCopy_HAComponents_E", "TransCopy_HAComponents_F",
                      "TransCopy_HAComponents_G", "TransCopy_HAComponents_H", "TransCopy_HAComponents_I",
                      "TransCopy_HAComponents_J", "TransCopy_HAComponents_K", "TransCopy_HAComponents_L",
                      "TransCopy_HAComponents_M", "TransCopy_HAComponents_N", "TransCopy_HAComponents_O",
                      "TransCopy_HAComponents_P", "TransCopy_HAComponents_Q", "TransCopy_HAComponents_R",
                      "TransCopy_HAComponents_S", "TransCopy_HAComponents_T", "TransCopy_HAComponents_U",
                      "TransCopy_HAComponents_V", "TransCopy_HAComponents_W", "TransCopy_HAComponents_X",
                      "TransCopy_HAComponents_Y"))

            Rules.Add(New TransactionalDBSyncRule(
                      "Rep_OFA_010", -1,
                      "TransCopy_OFA_A", "TransCopy_OFA_B", "TransCopy_OFA_C", "TransCopy_OFA_D"))

            Return Rules

        End Function

        '=========================================================================
        ' PUBLIC ENTRY POINTS
        '=========================================================================

        Public Function FullTransactionalDBSync() As AbovoTransaction

            Return SynchroniseRules(SyncRules, Nothing)

        End Function

        Public Function SynchroniseAll() As AbovoTransaction

            Return FullTransactionalDBSync()

        End Function

        Public Function SynchroniseForNamedRange(ByVal ChangedNamedRange As String) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}

            If IsSynchronising Then Return Result
            If String.IsNullOrWhiteSpace(ChangedNamedRange) Then Return Result

            Dim WB As IWorkbook = GetWorkbook()
            If WB Is Nothing Then Return Result

            Dim MatchingRules As New List(Of TransactionalDBSyncRule)
            Dim ChangedWorksheet As String = GetNamedRangeWorksheetName(ChangedNamedRange)

            'If the changed range can be resolved, synchronise all compatibility
            'rules whose source named range lives on the same worksheet.  This is
            'intentional: structural edits frequently target a companion/repeating
            'range rather than the exact source range used to size TransactionDB.
            If Not String.IsNullOrWhiteSpace(ChangedWorksheet) Then

                For Each Rule As TransactionalDBSyncRule In SyncRules

                    Dim SourceWorksheet As String = GetNamedRangeWorksheetName(Rule.SourceNamedRange)

                    If String.Equals(SourceWorksheet, ChangedWorksheet, StringComparison.OrdinalIgnoreCase) Then
                        MatchingRules.Add(Rule)
                    End If

                Next

            Else

                'If the changed range cannot currently be resolved, still honour an
                'exact source-range name match.  This keeps the synchroniser useful
                'during transitional structural operations.
                For Each Rule As TransactionalDBSyncRule In SyncRules
                    If String.Equals(Rule.SourceNamedRange, ChangedNamedRange, StringComparison.OrdinalIgnoreCase) Then
                        MatchingRules.Add(Rule)
                    End If
                Next

            End If

            Result = SynchroniseRules(MatchingRules, ChangedNamedRange)

            Return Result

        End Function

        Public Function SynchroniseForWorksheet(ByVal WorksheetName As String) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}

            If IsSynchronising Then Return Result
            If String.IsNullOrWhiteSpace(WorksheetName) Then Return Result

            Dim MatchingRules As New List(Of TransactionalDBSyncRule)

            For Each Rule As TransactionalDBSyncRule In SyncRules

                Dim SourceWorksheet As String = GetNamedRangeWorksheetName(Rule.SourceNamedRange)

                If String.Equals(SourceWorksheet, WorksheetName, StringComparison.OrdinalIgnoreCase) Then
                    MatchingRules.Add(Rule)
                End If

            Next

            Return SynchroniseRules(MatchingRules, Nothing)

        End Function

        'Existing project entry point used by SpecificRowColumnEvents.
        Public Sub SynchroniseOFA()

            SynchroniseSourceNamedRange("Rep_OFA_010")

        End Sub

        Public Sub SynchroniseDevelopment()
            SynchroniseSourceNamedRange("HouseTypeInID")
            SynchroniseSourceNamedRange("HouseTypeInMY")
        End Sub

        Public Sub SynchroniseJV()
            SynchroniseSourceNamedRange("IC_JointVenture_01")
        End Sub

        Public Sub SynchroniseCapGrant()
            SynchroniseSourceNamedRange("CapGrantInclusion")
        End Sub

        Public Sub SynchroniseRepairs()
            SynchroniseSourceNamedRange("StockCondCats")
        End Sub

        Public Sub SynchroniseCapEx()
            SynchroniseSourceNamedRange("Rep_CapExpend_010")
        End Sub

        Public Sub SynchroniseFunding()
            SynchroniseSourceNamedRange("FacilityNames")
            SynchroniseSourceNamedRange("LoanDescsOrd")
        End Sub

        Public Sub SynchroniseInterCompanyFunding()
            SynchroniseSourceNamedRange("IC_IntercoFunding_01")
            SynchroniseSourceNamedRange("IC_IntercoFunding_02")
        End Sub

        Public Sub SynchroniseHousingComponents()
            SynchroniseSourceNamedRange("DepnType")
        End Sub

        '=========================================================================
        ' TRANSACTION DB DEPENDENCY DIAGNOSTIC
        '
        'Run manually from the Visual Studio Immediate window when required:
        '
        ' FileManager.ExcelModels(ApplicationConfiguration.ActiveModelID).
        '     TransDBSync.DiagnoseTransactionDBDependencies()
        '
        'This routine is deliberately NOT called during normal synchronisation.
        '=========================================================================

        '=====================================================================
        ' TRANSACTION DB INTERNAL STRUCTURE / FORMULA DIAGNOSTIC
        '
        'Read-only profile intended to determine whether Transactional DB can
        'be simplified into a periodically refreshed cache/data sheet rather
        'than remain part of the workbook's live dependency structure.
        '=====================================================================

        Public Sub DiagnoseTransactionDBInternalStructure()

            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Return
            End If

            Const SheetName As String = "Transactional DB"
            Const MaximumExamples As Integer = 100

            Dim WS As Worksheet = Nothing

            Try
                WS = WB.Worksheets(SheetName)
            Catch
                WS = Nothing
            End Try

            Dim DiagnosticTimer As Stopwatch = Stopwatch.StartNew()


            If WS Is Nothing Then
                Return
            End If

            Dim UsedRange As CellRange = WS.GetUsedRange()


            Dim DevelopmentTop As Integer = Integer.MaxValue
            Dim DevelopmentBottom As Integer = -1
            Dim DevelopmentNameCount As Integer = 0

            For Each DN As DefinedName In WB.DefinedNames

                Try

                    If DN IsNot Nothing AndAlso
                       DN.Range IsNot Nothing AndAlso
                       DN.Range.Worksheet IsNot Nothing AndAlso
                       String.Equals(DN.Range.Worksheet.Name,
                                     SheetName,
                                     StringComparison.OrdinalIgnoreCase) AndAlso
                       (DN.Name.StartsWith("TransCopy_DevptSingle_",
                                           StringComparison.OrdinalIgnoreCase) OrElse
                        DN.Name.StartsWith("TransCopy_DevptMulti_",
                                           StringComparison.OrdinalIgnoreCase)) Then

                        DevelopmentNameCount += 1
                        DevelopmentTop = Math.Min(DevelopmentTop,
                                                  DN.Range.TopRowIndex)
                        DevelopmentBottom = Math.Max(DevelopmentBottom,
                                                     DN.Range.BottomRowIndex)

                    End If

                Catch
                End Try

            Next

            If DevelopmentTop = Integer.MaxValue Then

                DevelopmentTop = -1
                DevelopmentBottom = -1

            Else


            End If

            Dim FormulaCount As Integer = 0
            Dim ValueOnlyExistingCells As Integer = 0
            Dim SelfReferenceFormulaCount As Integer = 0
            Dim TransCopyTokenFormulaCount As Integer = 0
            Dim DevelopmentRegionFormulaCount As Integer = 0
            Dim DevelopmentSheetReferenceCount As Integer = 0
            Dim SelfExamples As Integer = 0
            Dim DevelopmentExamples As Integer = 0

            Dim FormulaCountByRow As New Dictionary(Of Integer, Integer)

            Dim DevelopmentTokens() As String = {
                "'Development BP Assumptions'!",
                "'Development Stock'!",
                "'Development Capital'!",
                "'Development Revenue'!",
                "'Development Expenditure'!",
                "'Dvpt NonCash'!",
                "'Dvpt Component Depn'!"
            }

            For Each Cell As Cell In WS.GetExistingCells()

                Dim Formula As String = Cell.Formula

                If String.IsNullOrWhiteSpace(Formula) Then

                    ValueOnlyExistingCells += 1
                    Continue For

                End If

                FormulaCount += 1

                If FormulaCountByRow.ContainsKey(Cell.RowIndex) Then
                    FormulaCountByRow(Cell.RowIndex) += 1
                Else
                    FormulaCountByRow.Add(Cell.RowIndex, 1)
                End If

                If DevelopmentTop >= 0 AndAlso
                   Cell.RowIndex >= DevelopmentTop AndAlso
                   Cell.RowIndex <= DevelopmentBottom Then

                    DevelopmentRegionFormulaCount += 1

                End If

                If Formula.IndexOf("'Transactional DB'!",
                                   StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   Formula.IndexOf("Transactional DB!",
                                   StringComparison.OrdinalIgnoreCase) >= 0 Then

                    SelfReferenceFormulaCount += 1

                    If SelfExamples < MaximumExamples Then


                        SelfExamples += 1

                    End If

                End If

                If Formula.IndexOf("TransCopy_",
                                   StringComparison.OrdinalIgnoreCase) >= 0 Then

                    TransCopyTokenFormulaCount += 1

                End If

                Dim ReferencesDevelopment As Boolean = False

                For Each Token As String In DevelopmentTokens

                    If Formula.IndexOf(Token,
                                       StringComparison.OrdinalIgnoreCase) >= 0 Then

                        ReferencesDevelopment = True
                        Exit For

                    End If

                Next

                If ReferencesDevelopment Then

                    DevelopmentSheetReferenceCount += 1

                    If DevelopmentExamples < MaximumExamples Then


                        DevelopmentExamples += 1

                    End If

                End If

            Next



            Dim HotRows As New List(Of KeyValuePair(Of Integer, Integer))(FormulaCountByRow)

            HotRows.Sort(
                Function(A As KeyValuePair(Of Integer, Integer),
                         B As KeyValuePair(Of Integer, Integer)) As Integer

                    Dim CountCompare As Integer = B.Value.CompareTo(A.Value)

                    If CountCompare <> 0 Then Return CountCompare

                    Return A.Key.CompareTo(B.Key)

                End Function)

            Dim HotRowLimit As Integer = Math.Min(25, HotRows.Count)

            For Index As Integer = 0 To HotRowLimit - 1


            Next


            Dim NamesOnTransDB As Integer = 0
            Dim NamesTouchingDevelopmentEnvelope As Integer = 0
            Dim NamesSpanningDevelopmentEnvelope As Integer = 0

            For Each DN As DefinedName In WB.DefinedNames

                Try

                    If DN Is Nothing OrElse
                       DN.Range Is Nothing OrElse
                       DN.Range.Worksheet Is Nothing OrElse
                       Not String.Equals(DN.Range.Worksheet.Name,
                                         SheetName,
                                         StringComparison.OrdinalIgnoreCase) Then Continue For

                    NamesOnTransDB += 1

                    If DevelopmentTop >= 0 AndAlso
                       DN.Range.BottomRowIndex >= DevelopmentTop AndAlso
                       DN.Range.TopRowIndex <= DevelopmentBottom Then

                        NamesTouchingDevelopmentEnvelope += 1

                        Dim SpansEnvelope As Boolean =
                            DN.Range.TopRowIndex < DevelopmentTop AndAlso
                            DN.Range.BottomRowIndex > DevelopmentBottom

                        If SpansEnvelope Then
                            NamesSpanningDevelopmentEnvelope += 1
                        End If


                    End If

                Catch ex As Exception


                End Try

            Next




            If FormulaCount = 0 Then


            ElseIf SelfReferenceFormulaCount = 0 AndAlso
                   DevelopmentSheetReferenceCount = 0 Then


            Else


            End If



        End Sub

        Public Sub DiagnoseTransactionDBDependencies()

            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Return
            End If

            Const TransactionDBSheetName As String = "Transactional DB"
            Const MaximumExamplesPerType As Integer = 250

            Dim DiagnosticTimer As Stopwatch = Stopwatch.StartNew()


            Dim TransactionDBSheet As Worksheet = Nothing

            Try
                TransactionDBSheet = WB.Worksheets(TransactionDBSheetName)
            Catch
                TransactionDBSheet = Nothing
            End Try

            If TransactionDBSheet Is Nothing Then
                Return
            End If

            Dim TransUsedRange As CellRange = TransactionDBSheet.GetUsedRange()



            Dim GlobalNamesOnTransDB As Integer = 0

            For Each DN As DefinedName In WB.DefinedNames

                Try

                    If DN IsNot Nothing AndAlso
                       DN.Range IsNot Nothing AndAlso
                       DN.Range.Worksheet IsNot Nothing AndAlso
                       String.Equals(DN.Range.Worksheet.Name,
                                     TransactionDBSheetName,
                                     StringComparison.OrdinalIgnoreCase) Then

                        GlobalNamesOnTransDB += 1


                    End If

                Catch ex As Exception


                End Try

            Next

            Dim LocalNamesOnTransDB As Integer = 0

            For Each ScopeWS As Worksheet In WB.Worksheets

                For Each DN As DefinedName In ScopeWS.DefinedNames

                    Try

                        If DN IsNot Nothing AndAlso
                           DN.Range IsNot Nothing AndAlso
                           DN.Range.Worksheet IsNot Nothing AndAlso
                           String.Equals(DN.Range.Worksheet.Name,
                                         TransactionDBSheetName,
                                         StringComparison.OrdinalIgnoreCase) Then

                            LocalNamesOnTransDB += 1


                        End If

                    Catch ex As Exception


                    End Try

                Next

            Next



            Dim FormulaCellsScanned As Long = 0
            Dim DirectTransactionDBReferences As Long = 0
            Dim TransCopyNameReferences As Long = 0
            Dim FormulasContainingBoth As Long = 0

            Dim DirectExamplesPrinted As Integer = 0
            Dim NamedExamplesPrinted As Integer = 0

            Dim QuotedSheetToken As String = "'" & TransactionDBSheetName & "'!"
            Dim UnquotedSheetToken As String = TransactionDBSheetName & "!"

            For Each WS As Worksheet In WB.Worksheets

                'References inside Transactional DB itself are not external consumers
                'of the sheet and are therefore excluded from the direct-reference
                'count.  They are still represented by the named-range inventory above.
                If String.Equals(WS.Name,
                                 TransactionDBSheetName,
                                 StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                Dim SheetTimer As Stopwatch = Stopwatch.StartNew()
                Dim UsedRange As CellRange = WS.GetUsedRange()

                Dim SheetFormulaCount As Long = 0
                Dim SheetDirectCount As Long = 0
                Dim SheetNamedCount As Long = 0

                For RowIndex As Integer = UsedRange.TopRowIndex To UsedRange.BottomRowIndex

                    For ColumnIndex As Integer = UsedRange.LeftColumnIndex To UsedRange.RightColumnIndex

                        Dim Cell As Cell = WS.Cells(RowIndex, ColumnIndex)
                        Dim FormulaText As String = Convert.ToString(Cell.Formula)

                        If String.IsNullOrWhiteSpace(FormulaText) Then Continue For

                        FormulaCellsScanned += 1
                        SheetFormulaCount += 1

                        Dim HasDirectReference As Boolean =
                            FormulaText.IndexOf(QuotedSheetToken,
                                                StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                            FormulaText.IndexOf(UnquotedSheetToken,
                                                StringComparison.OrdinalIgnoreCase) >= 0

                        Dim HasTransCopyName As Boolean =
                            FormulaText.IndexOf("TransCopy_",
                                                StringComparison.OrdinalIgnoreCase) >= 0

                        If HasDirectReference Then

                            DirectTransactionDBReferences += 1
                            SheetDirectCount += 1

                            If DirectExamplesPrinted < MaximumExamplesPerType Then

                                DirectExamplesPrinted += 1


                            End If

                        End If

                        If HasTransCopyName Then

                            TransCopyNameReferences += 1
                            SheetNamedCount += 1

                            If NamedExamplesPrinted < MaximumExamplesPerType Then

                                NamedExamplesPrinted += 1


                            End If

                        End If

                        If HasDirectReference AndAlso HasTransCopyName Then
                            FormulasContainingBoth += 1
                        End If

                    Next

                Next

                If SheetFormulaCount > 0 OrElse
                   SheetDirectCount > 0 OrElse
                   SheetNamedCount > 0 Then


                End If

            Next


            If DirectTransactionDBReferences = 0 Then


            Else


            End If


        End Sub

        '=========================================================================
        ' CORE SYNCHRONISATION
        '=========================================================================

        Private Function SynchroniseSourceNamedRange(ByVal SourceNamedRange As String) As AbovoTransaction

            Dim Rules As New List(Of TransactionalDBSyncRule)

            For Each Rule As TransactionalDBSyncRule In SyncRules
                If String.Equals(Rule.SourceNamedRange, SourceNamedRange, StringComparison.OrdinalIgnoreCase) Then
                    Rules.Add(Rule)
                End If
            Next

            Return SynchroniseRules(Rules, Nothing)

        End Function

        Private Function SynchroniseRules(ByVal Rules As IEnumerable(Of TransactionalDBSyncRule),
                                          ByVal LegacyChangedNamedRange As String) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}

            If IsSynchronising Then Return Result

            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            Dim Requests As New Dictionary(Of String, MirrorResizeRequest)(StringComparer.OrdinalIgnoreCase)

            If Rules IsNot Nothing Then

                For Each Rule As TransactionalDBSyncRule In Rules
                    AddRuleResizeRequests(Rule, Requests)

                Next

            End If

            If Not String.IsNullOrWhiteSpace(LegacyChangedNamedRange) Then
                AddLegacyDirectMirrorRequests(LegacyChangedNamedRange, Requests)
            End If

            If Requests.Count = 0 Then
                Result.StringReturn = "No TransactionDB mirror changes required."
                Return Result

            End If

            'Resolve all ranges before changing the worksheet, then process from
            'bottom to top.  This minimises the amount of lower worksheet content
            'that each full-row insert/delete has to shift.
            Dim WorkItems As New List(Of MirrorResizeWorkItem)

            For Each Request As MirrorResizeRequest In Requests.Values

                Dim DN As DefinedName = Nothing

                Try
                    DN = WB.DefinedNames.GetDefinedName(Request.TargetNamedRange)
                Catch
                    DN = Nothing
                End Try

                If DN Is Nothing OrElse DN.Range Is Nothing OrElse DN.Range.Worksheet Is Nothing Then
                    Continue For
                End If

                Dim Rng As CellRange = DN.Range

                If Rng.RowCount = Request.RequiredRows Then Continue For

                WorkItems.Add(New MirrorResizeWorkItem With {
                    .Request = Request,
                    .WorksheetName = Rng.Worksheet.Name,
                    .TopRowIndex = Rng.TopRowIndex,
                    .BottomRowIndex = Rng.BottomRowIndex,
                    .LeftColumnIndex = Rng.LeftColumnIndex,
                    .RightColumnIndex = Rng.RightColumnIndex,
                    .CurrentRows = Rng.RowCount
                })

            Next

            WorkItems.Sort(
                Function(A As MirrorResizeWorkItem, B As MirrorResizeWorkItem) As Integer

                    Dim SheetCompare As Integer =
                        String.Compare(A.WorksheetName,
                                       B.WorksheetName,
                                       StringComparison.OrdinalIgnoreCase)

                    If SheetCompare <> 0 Then Return SheetCompare

                    'Descending row order within each worksheet.
                    Return B.TopRowIndex.CompareTo(A.TopRowIndex)

                End Function)


            For Each Item As MirrorResizeWorkItem In WorkItems


            Next

            If WorkItems.Count = 0 Then
                Result.StringReturn = "TransactionDB mirrors already aligned."
                Return Result
            End If

            Dim RDSDisconnected As Boolean = False
            Dim RDSV2Disconnected As Boolean = False
            Dim PreviousCalculationMode As WorkbookCalculationMode = WB.Options.CalculationMode
            Dim PreviousCalculationEngineType As CalculationEngineType = WB.Options.CalculationEngineType
            Dim CalculationEngineChanged As Boolean = False
            Dim PreviousHistoryEnabled As Boolean = WB.History.IsEnabled
            Dim HistoryChanged As Boolean = False
            Dim UpdateStarted As Boolean = False
            Dim StructuralChangeAttempted As Boolean = False
            Dim SyncTimer As Stopwatch = Stopwatch.StartNew()

            Try

                IsSynchronising = True

                Dim HistoryTimer As Stopwatch = Stopwatch.StartNew()

                If WB.History.IsEnabled Then
                    WB.History.IsEnabled = False
                    HistoryChanged = True
                End If


                Dim EngineTimer As Stopwatch = Stopwatch.StartNew()

                If WB.Options.CalculationEngineType <> CalculationEngineType.Recursive Then
                    WB.Options.CalculationEngineType = CalculationEngineType.Recursive
                    CalculationEngineChanged = True
                End If


                If WB.Options.CalculationMode <> WorkbookCalculationMode.Manual Then
                    WB.Options.CalculationMode = WorkbookCalculationMode.Manual
                End If

                WB.BeginUpdate()
                UpdateStarted = True

                If ExcelModels(ModelID).ExpendAnalyser IsNot Nothing Then

                    Dim DisconnectTimer As Stopwatch = Stopwatch.StartNew()

                    ExcelModels(ModelID).ExpendAnalyser.DisconectRDS()
                    RDSDisconnected = True


                End If

                If ExcelModels(ModelID).ExpendAnalyserV2 IsNot Nothing Then

                    ExcelModels(ModelID).ExpendAnalyserV2.DisconnectRDS()
                    RDSV2Disconnected = True

                End If

                For Each Item As MirrorResizeWorkItem In WorkItems

                    Dim MirrorTimer As Stopwatch = Stopwatch.StartNew()
                    StructuralChangeAttempted = True

                    Dim ResizeResult As AbovoTransaction =
                        ResizeMirrorRangeInBatch(WB,
                                                 Item.Request.TargetNamedRange,
                                                 Item.Request.RequiredRows)


                    If ResizeResult.BError Then
                        Result.BError = True
                        Result.StringReturn &= Item.Request.TargetNamedRange &
                                               ": " & ResizeResult.StringReturn &
                                               Environment.NewLine
                    End If

                Next

            Catch ex As Exception

                Result.BError = True
                Result.StringReturn &= ex.Message

            Finally

                If UpdateStarted Then

                    Dim EndUpdateTimer As Stopwatch = Stopwatch.StartNew()
                    WB.EndUpdate()


                End If

                If StructuralChangeAttempted Then
                    Try
                        TransactionalDBSnapshotManager.InvalidateSnapshot(ModelID)

                        If ExcelModels(ModelID).ExpendAnalyserV2 IsNot Nothing Then
                            ExcelModels(ModelID).ExpendAnalyserV2.NotifySnapshotInvalidated()
                        End If
                    Catch ex As Exception
                        Result.BError = True
                        Result.StringReturn &=
                            "Transactional DB snapshot invalidation: " & ex.Message &
                            Environment.NewLine
                    End Try
                End If

                Dim RestoreCalcTimer As Stopwatch = Stopwatch.StartNew()
                WB.Options.CalculationMode = PreviousCalculationMode


                If CalculationEngineChanged Then

                    Dim RestoreEngineTimer As Stopwatch = Stopwatch.StartNew()
                    WB.Options.CalculationEngineType = PreviousCalculationEngineType


                End If

                If HistoryChanged Then

                    Dim RestoreHistoryTimer As Stopwatch = Stopwatch.StartNew()
                    WB.History.IsEnabled = PreviousHistoryEnabled


                End If

                If RDSDisconnected AndAlso ExcelModels(ModelID).ExpendAnalyser IsNot Nothing Then

                    Dim ReconnectTimer As Stopwatch = Stopwatch.StartNew()
                    ExcelModels(ModelID).ExpendAnalyser.ReconnectRDS()


                End If

                If RDSV2Disconnected AndAlso ExcelModels(ModelID).ExpendAnalyserV2 IsNot Nothing Then

                    ExcelModels(ModelID).ExpendAnalyserV2.ReconnectRDS()

                End If


                IsSynchronising = False

            End Try

            Return Result

        End Function

        Private Function GetTransactionDBShiftBounds(ByVal WS As Worksheet,
                                                     ByVal RangeLeft As Integer,
                                                     ByVal RangeRight As Integer,
                                                     ByRef ShiftLeft As Integer,
                                                     ByRef ShiftRight As Integer) As Boolean

            If WS Is Nothing Then Return False

            Dim UsedRangeTimer As Stopwatch = Stopwatch.StartNew()
            Dim UsedRange As CellRange = WS.GetUsedRange()

            'Shift every USED column on the TransactionDB sheet, not merely the
            'mirror's own columns.  This preserves any row-aligned helper data to
            'the right of the TransCopy ranges while avoiding an EntireRow insert
            'across all 16,384 worksheet columns.
            ShiftLeft = Math.Min(UsedRange.LeftColumnIndex, RangeLeft)
            ShiftRight = Math.Max(UsedRange.RightColumnIndex, RangeRight)


            Dim CanUseFastPath As Boolean =
                ShiftLeft >= 0 AndAlso
                ShiftRight >= ShiftLeft AndAlso
                ShiftRight <= 16383


            Return CanUseFastPath

        End Function

        Private Function ResizeMirrorRangeInBatch(ByVal WB As IWorkbook,
                                                  ByVal TargetNamedRange As String,
                                                  ByVal RequiredRows As Integer) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}

            Try

                Dim DN As DefinedName = WB.DefinedNames.GetDefinedName(TargetNamedRange)

                If DN Is Nothing OrElse DN.Range Is Nothing OrElse DN.Range.Worksheet Is Nothing Then
                    Result.BError = True
                    Result.StringReturn = "Named range was not found."
                    Return Result
                End If

                Dim OriginalRange As CellRange = DN.Range
                Dim WS As Worksheet = OriginalRange.Worksheet

                Dim RangeLeft As Integer = OriginalRange.LeftColumnIndex
                Dim RangeRight As Integer = OriginalRange.RightColumnIndex
                Dim RangeTop As Integer = OriginalRange.TopRowIndex
                Dim RangeBottom As Integer = OriginalRange.BottomRowIndex
                Dim CurrentRows As Integer = OriginalRange.RowCount

                If CurrentRows = RequiredRows Then Return Result

                'The final row is the mirror footer.  Formula/data rows are added
                'or removed immediately before that footer, exactly as in the
                'Summit_Compatibility SetTransDBMirrorRangeSize VBA routine.
                If RequiredRows > CurrentRows Then

                    Dim RowsToAdd As Integer = RequiredRows - CurrentRows
                    Dim LastFormulaRow As Integer = RangeBottom - 1

                    If LastFormulaRow < RangeTop Then
                        Result.BError = True
                        Result.StringReturn = "Mirror range has no template formula row."
                        Return Result
                    End If

                    Dim TemplateHeight As Single = WS.Rows(LastFormulaRow).Height

                    Dim ShiftLeft As Integer = RangeLeft
                    Dim ShiftRight As Integer = RangeRight

                    Dim UseFastPath As Boolean =
                        GetTransactionDBShiftBounds(WS,
                                                    RangeLeft,
                                                    RangeRight,
                                                    ShiftLeft,
                                                    ShiftRight)

                    Dim InsertTimer As Stopwatch = Stopwatch.StartNew()

                    If UseFastPath Then

                        Dim InsertRange As CellRange =
                            WS.Range.FromLTRB(ShiftLeft,
                                              RangeBottom,
                                              ShiftRight,
                                              RangeBottom + RowsToAdd - 1)

                        WS.InsertCells(InsertRange, InsertCellsMode.ShiftCellsDown)

                    Else

                        WS.Rows.Insert(RangeBottom, RowsToAdd)

                    End If

                    Dim InsertElapsed As Long = InsertTimer.ElapsedMilliseconds

                    Dim SourceTemplate As CellRange =
                        WS.Range.FromLTRB(RangeLeft,
                                          LastFormulaRow,
                                          RangeRight,
                                          LastFormulaRow)

                    Dim NewRows As CellRange =
                        WS.Range.FromLTRB(RangeLeft,
                                          RangeBottom,
                                          RangeRight,
                                          RangeBottom + RowsToAdd - 1)

                    Dim FillTimer As Stopwatch = Stopwatch.StartNew()

                    NewRows.CopyFrom(SourceTemplate, PasteSpecial.All)

                    If Not UseFastPath Then
                        For RowIndex As Integer = RangeBottom To RangeBottom + RowsToAdd - 1
                            WS.Rows(RowIndex).Height = TemplateHeight
                        Next
                    End If

                    Dim FillElapsed As Long = FillTimer.ElapsedMilliseconds

                    Dim ExpandedRange As CellRange =
                        WS.Range.FromLTRB(RangeLeft,
                                          RangeTop,
                                          RangeRight,
                                          RangeBottom + RowsToAdd)

                    If DN.IsGlobal Then
                        WB.DefinedNames.GetDefinedName(TargetNamedRange).Range = ExpandedRange
                    Else
                        WS.DefinedNames.GetDefinedName(TargetNamedRange).Range = ExpandedRange
                    End If


                Else

                    Dim RowsToDelete As Integer = CurrentRows - RequiredRows

                    'Never delete the footer.  Remove formula/data rows immediately
                    'above it as the compatibility VBA does.
                    If RequiredRows < 1 OrElse RowsToDelete >= CurrentRows Then
                        Result.BError = True
                        Result.StringReturn = "Requested mirror size is invalid."
                        Return Result
                    End If

                    Dim FirstDeleteRow As Integer = RangeBottom - RowsToDelete

                    Dim ShiftLeft As Integer = RangeLeft
                    Dim ShiftRight As Integer = RangeRight

                    Dim UseFastPath As Boolean =
                        GetTransactionDBShiftBounds(WS,
                                                    RangeLeft,
                                                    RangeRight,
                                                    ShiftLeft,
                                                    ShiftRight)

                    Dim DeleteTimer As Stopwatch = Stopwatch.StartNew()

                    If UseFastPath Then

                        Dim DeleteRange As CellRange =
                            WS.Range.FromLTRB(ShiftLeft,
                                              FirstDeleteRow,
                                              ShiftRight,
                                              FirstDeleteRow + RowsToDelete - 1)

                        WS.DeleteCells(DeleteRange, DeleteMode.ShiftCellsUp)

                    Else

                        WS.Rows.Remove(FirstDeleteRow, RowsToDelete)

                    End If

                    Dim DeleteElapsed As Long = DeleteTimer.ElapsedMilliseconds

                    Dim ContractedRange As CellRange =
                        WS.Range.FromLTRB(RangeLeft,
                                          RangeTop,
                                          RangeRight,
                                          RangeBottom - RowsToDelete)

                    If DN.IsGlobal Then
                        WB.DefinedNames.GetDefinedName(TargetNamedRange).Range = ContractedRange
                    Else
                        WS.DefinedNames.GetDefinedName(TargetNamedRange).Range = ContractedRange
                    End If


                End If

            Catch ex As Exception

                Result.BError = True
                Result.StringReturn = ex.Message

            End Try

            Return Result

        End Function

        Private Sub AddRuleResizeRequests(ByVal Rule As TransactionalDBSyncRule,
                                          ByVal Requests As Dictionary(Of String, MirrorResizeRequest))

            If Rule Is Nothing Then Return
            If Not WorkbookManager.DoesNRExist(ModelID, Rule.SourceNamedRange) Then Return

            Dim SourceColumns As Integer = WorkbookManager.GetRangeColumns(ModelID, Rule.SourceNamedRange)
            Dim DataRowsRequired As Integer = SourceColumns + Rule.SourceColumnAdjustment

            If DataRowsRequired < 0 Then DataRowsRequired = 0

            Dim TotalMirrorRowsRequired As Integer = DataRowsRequired + 1

            For Each TargetNamedRange As String In Rule.TargetNamedRanges

                If String.IsNullOrWhiteSpace(TargetNamedRange) Then Continue For
                If Not WorkbookManager.DoesNRExist(ModelID, TargetNamedRange) Then Continue For

                Requests(TargetNamedRange) = New MirrorResizeRequest With {
                    .TargetNamedRange = TargetNamedRange,
                    .RequiredRows = TotalMirrorRowsRequired
                }

            Next

        End Sub

        Private Sub AddLegacyDirectMirrorRequests(ByVal ChangedNamedRange As String,
                                                  ByVal Requests As Dictionary(Of String, MirrorResizeRequest))

            If Not WorkbookManager.DoesNRExist(ModelID, ChangedNamedRange) Then Return

            Dim RequiredRows As Integer = WorkbookManager.GetRangeRows(ModelID, ChangedNamedRange) + 1
            Dim BaseMirrorName As String = "TransCopy_" & ChangedNamedRange

            AddLegacyMirrorRequest(BaseMirrorName, RequiredRows, Requests)

            For Suffix As Integer = 1 To 6
                AddLegacyMirrorRequest(BaseMirrorName & "_" & Suffix.ToString("00"), RequiredRows, Requests)
            Next

        End Sub

        Private Sub AddLegacyMirrorRequest(ByVal TargetNamedRange As String,
                                           ByVal RequiredRows As Integer,
                                           ByVal Requests As Dictionary(Of String, MirrorResizeRequest))

            If Not WorkbookManager.DoesNRExist(ModelID, TargetNamedRange) Then Return

            Requests(TargetNamedRange) = New MirrorResizeRequest With {
                .TargetNamedRange = TargetNamedRange,
                .RequiredRows = RequiredRows
            }

        End Sub

        Private Function GetNamedRangeWorksheetName(ByVal NamedRange As String) As String

            If String.IsNullOrWhiteSpace(NamedRange) Then Return Nothing

            Dim WB As IWorkbook = GetWorkbook()
            If WB Is Nothing Then Return Nothing

            Try

                Dim DN As DefinedName = WB.DefinedNames.GetDefinedName(NamedRange)

                If DN Is Nothing OrElse DN.Range Is Nothing OrElse DN.Range.Worksheet Is Nothing Then Return Nothing

                Return DN.Range.Worksheet.Name

            Catch

                Return Nothing

            End Try

        End Function

    End Class

End Namespace
