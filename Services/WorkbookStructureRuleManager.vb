Imports System
Imports System.Collections.Generic
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.WSSecurity
Imports DevExpress.Spreadsheet

Namespace Abovo

    '=========================================================================
    ' WorkbookStructureRuleManager
    '
    ' Executes semantic structural operations ("add records", "delete records")
    ' across all workbook sheets that form one logical model structure.
    '
    ' The UI does not need to know whether a record is physically represented by
    ' a row or a column.  That belongs to the structural rule.  This is especially
    ' important for MergeDownAndPivot interfaces, where an apparent interface row
    ' is commonly an underlying workbook column.
    '=========================================================================
    Public Class WorkbookStructureRuleManager

        Public Const RuleOFARecords As String = "OFA_RECORDS"
        Public Const RuleCapExRecords As String = "CAPEX_RECORDS"
        Public Const RuleCapGrantRecords As String = "CAPGRANT_RECORDS"
        Public Const RuleRepairsRecords As String = "REPAIRS_RECORDS"
        Public Const RuleHousingComponentRecords As String = "HOUSING_COMPONENT_RECORDS"
        Public Const RuleFundingRecords As String = "FUNDING_RECORDS"
        Public Const RuleDevelopmentIdentifiedRecords As String = "DEVELOPMENT_IDENTIFIED_RECORDS"
        Public Const RuleDevelopmentMultiYearRecords As String = "DEVELOPMENT_MULTIYEAR_RECORDS"
        Public Const RuleJournalRecords As String = "JOURNAL_RECORDS"
        Public Const RuleStockConversionRecords As String = "STOCK_CONVERSION_RECORDS"

        Private ReadOnly ModelID As Integer
        Private ReadOnly Rules As Dictionary(Of String, WorkbookStructureRule)
        Private IsExecuting As Boolean = False

        Public Sub New(ByVal SetModelID As Integer)

            ModelID = SetModelID
            Rules = CreateRules()

        End Sub

        Public Function AddRecords(ByVal RuleID As String,
                                   ByVal RecordCount As Integer) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}

            If RecordCount <= 0 Then
                Result.StringReturn = "No records requested."
                Return Result
            End If

            Dim Rule As WorkbookStructureRule = GetRule(RuleID)

            If Rule Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook structure rule '" & RuleID & "' was not found."
                Return Result
            End If

            Return ExecuteInsert(Rule, RecordCount)

        End Function

        Public Function DeleteRecords(ByVal RuleID As String,
                                      ByVal RecordIndexes As IEnumerable(Of Integer)) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim Rule As WorkbookStructureRule = GetRule(RuleID)

            If Rule Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook structure rule '" & RuleID & "' was not found."
                Return Result
            End If

            Dim Indexes As New List(Of Integer)

            If RecordIndexes IsNot Nothing Then
                For Each Index As Integer In RecordIndexes
                    If Index >= 0 AndAlso Not Indexes.Contains(Index) Then Indexes.Add(Index)
                Next
            End If

            If Indexes.Count = 0 Then
                Result.BError = True
                Result.EventCancelled = True
                Result.StringReturn = "No records were selected."
                Return Result
            End If

            Dim CurrentRecordCount As Integer = GetDeletionRecordCount(Rule)

            If Rule.RuleID = RuleFundingRecords Then
                If CurrentRecordCount < 0 Then
                    Result.BError = True
                    Result.EventCancelled = True
                    Result.StringReturn = "Funding ordinary-loan boundaries are inconsistent. No columns have been deleted."
                    Return Result
                End If
            End If

            If Indexes.Exists(Function(Index) Index < Rule.ProtectedLeadingRecordCount) Then
                Result.BError = True
                Result.EventCancelled = True
                Result.StringReturn = "The first " & Rule.ProtectedLeadingRecordCount.ToString() &
                    " records in '" & Rule.Description & "' are protected and cannot be deleted."
                Return Result
            End If

            If CurrentRecordCount >= 0 Then

                For Each Index As Integer In Indexes
                    If Index >= CurrentRecordCount Then
                        Result.BError = True
                        Result.EventCancelled = True
                        Result.StringReturn = "The selected record index is outside the current '" &
                                              Rule.Description & "' range."
                        Return Result
                    End If
                Next

                If Rule.MinimumRecordCount > 0 AndAlso
                   CurrentRecordCount - Indexes.Count < Rule.MinimumRecordCount Then

                    Result.BError = True
                    Result.EventCancelled = True
                    Result.StringReturn = "At least " & Rule.MinimumRecordCount.ToString &
                                          " " & Rule.Description & " must be retained."
                    Return Result

                End If

            End If

            Indexes.Sort()
            Indexes.Reverse()

            Return ExecuteDelete(Rule, Indexes)

        End Function

        Public Function ValidateDeleteLastRecords(ByVal RuleID As String,
                                                   ByVal RecordCount As Integer) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}

            If RecordCount <= 0 Then
                Result.BError = True
                Result.EventCancelled = True
                Result.StringReturn = "Please enter at least 1 record to delete."
                Return Result
            End If

            Dim Rule As WorkbookStructureRule = GetRule(RuleID)

            If Rule Is Nothing Then
                Result.BError = True
                Result.StringReturn =
                    "Workbook structure rule '" & RuleID & "' was not found."
                Return Result
            End If

            Dim CurrentRecordCount As Integer = GetDeletionRecordCount(Rule)

            If CurrentRecordCount < 0 Then
                Result.BError = True
                Result.StringReturn =
                    "The current record count for '" &
                    Rule.Description &
                    "' could not be determined."
                Return Result
            End If

            If RecordCount > CurrentRecordCount Then
                Result.BError = True
                Result.EventCancelled = True
                Result.StringReturn =
                    "Cannot delete " &
                    RecordCount.ToString &
                    " records because only " &
                    CurrentRecordCount.ToString &
                    " currently exist."
                Return Result
            End If

            If Rule.MinimumRecordCount > 0 AndAlso
               CurrentRecordCount - RecordCount < Rule.MinimumRecordCount Then

                Dim MaximumDeleteCount As Integer =
                    Math.Max(0, CurrentRecordCount - Rule.MinimumRecordCount)

                Result.BError = True
                Result.EventCancelled = True

                If MaximumDeleteCount = 0 Then

                    Result.StringReturn =
                        "No further records can be deleted." &
                        Environment.NewLine &
                        "At least " &
                        Rule.MinimumRecordCount.ToString &
                        " " &
                        Rule.Description &
                        " must be retained."

                Else

                    Result.StringReturn =
                        "You can delete a maximum of " &
                        MaximumDeleteCount.ToString &
                        " record(s)." &
                        Environment.NewLine &
                        "At least " &
                        Rule.MinimumRecordCount.ToString &
                        " " &
                        Rule.Description &
                        " must be retained."

                End If

                Return Result
            End If

            Return Result

        End Function

        Public Function DeleteLastRecords(ByVal RuleID As String,
                                          ByVal RecordCount As Integer) As AbovoTransaction

            Dim ValidationResult As AbovoTransaction =
                ValidateDeleteLastRecords(
                    RuleID,
                    RecordCount)

            If ValidationResult.BError Then Return ValidationResult

            Dim Rule As WorkbookStructureRule = GetRule(RuleID)
            Dim CurrentRecordCount As Integer = GetDeletionRecordCount(Rule)

            Dim DeleteIndexes As New List(Of Integer)

            For Index As Integer = CurrentRecordCount - 1 To CurrentRecordCount - RecordCount Step -1

                DeleteIndexes.Add(Index)

            Next

            Return DeleteRecords(RuleID, DeleteIndexes)

        End Function

        Public Function HasRule(ByVal RuleID As String) As Boolean

            If String.IsNullOrWhiteSpace(RuleID) Then Return False
            Return Rules.ContainsKey(RuleID)

        End Function

        Public Function ResolveRuleID(ByVal ExpansionToken As String) As String

            If String.IsNullOrWhiteSpace(ExpansionToken) Then Return Nothing

            If Rules.ContainsKey(ExpansionToken) Then Return ExpansionToken

            For Each Pair As KeyValuePair(Of String, WorkbookStructureRule) In Rules

                Dim Rule As WorkbookStructureRule = Pair.Value

                If String.Equals(ExpansionToken, Rule.RecordCountNamedRange, StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(ExpansionToken, Rule.DeleteAnchorNamedRange, StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(ExpansionToken, Rule.TransactionDBSyncNamedRange, StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(ExpansionToken, Rule.InsertAnchorNamedRange, StringComparison.OrdinalIgnoreCase) Then

                    Return Rule.RuleID

                End If

                If Rule.AdditionalRecordRanges.Contains(ExpansionToken, StringComparer.OrdinalIgnoreCase) Then Return Rule.RuleID

            Next

            Return Nothing

        End Function
        Public Function GetRecordCount(ByVal RuleID As String) As Integer

            Return GetRecordCount(RuleID, GetWorkbook())

        End Function

        Public Function GetRecordCount(ByVal RuleID As String,
                                       ByVal Workbook As IWorkbook) As Integer

            Return GetCurrentRecordCount(GetRule(RuleID), Workbook)

        End Function
        Private Function GetRule(ByVal RuleID As String) As WorkbookStructureRule

            If String.IsNullOrWhiteSpace(RuleID) Then Return Nothing

            Dim Rule As WorkbookStructureRule = Nothing
            If Rules.TryGetValue(RuleID, Rule) Then Return Rule

            Return Nothing

        End Function

        Private Function CreateRules() As Dictionary(Of String, WorkbookStructureRule)

            Dim RetRules As New Dictionary(Of String, WorkbookStructureRule)(StringComparer.OrdinalIgnoreCase)

            'The rules below mirror the linked-sheet column families in the XLSB VBA.
            'The UI deals in logical records.  All of these families are physically
            'stored as workbook columns, including MergeDownAndPivot interfaces.

            '-----------------------------------------------------------------
            ' Other Fixed Assets
            '-----------------------------------------------------------------
            'Rep_OFA_010 includes one trailing record which the interface
            'already excludes with <SkipLastRecords>1</SkipLastRecords>.
            'Keep structural record counting aligned with the interface.
            Dim OFARule As New WorkbookStructureRule With {
                .RuleID = RuleOFARecords,
                .Description = "Other Fixed Asset records",
                .Axis = WorkbookStructureAxis.Columns,
                .InsertAnchorNamedRange = "LastOFACol",
                .DeleteAnchorNamedRange = "Rep_OFA_030",
                .RecordCountNamedRange = "Rep_OFA_010",
                .RecordCountAdjustment = -1,
                .MinimumRecordCount = 3,
                .TransactionDBSyncNamedRange = "Rep_OFA_010"
            }

            'The immediate structural/template column is zero-width in the
            'current OFA workbook. Keep it as the format template, but obtain
            'the physical width from the preceding visible logical record.
            OFARule.Targets.Add(New WorkbookStructureTarget With {
                .WorksheetName = "Other Fixed Asset Assumptions",
                .CopyMode = WorkbookStructureCopyMode.FormatsAndColumnWidth,
                .TemplateOffset = -1,
                .ColumnWidthTemplateOffset = -2
            })
            OFARule.Targets.Add(New WorkbookStructureTarget With {
                .WorksheetName = "OFA Workings",
                .CopyMode = WorkbookStructureCopyMode.AllAndColumnWidth,
                .TemplateOffset = -1
            })
            RetRules.Add(OFARule.RuleID, OFARule)

            '-----------------------------------------------------------------
            ' Capital Expenditure
            ' VBA: Capital Expenditure Assumptions + OFA Additions
            '-----------------------------------------------------------------
            'Rep_CapExpend_010 includes one trailing blank/template column.
            'This mirrors <SkipLastRecords>1</SkipLastRecords> in the Capital
            'Expenditure interface definition. Structural operations therefore
            'count/select only genuine logical records.
            Dim CapExRule As New WorkbookStructureRule With {
                .RuleID = RuleCapExRecords,
                .Description = "Capital Expenditure records",
                .Axis = WorkbookStructureAxis.Columns,
                .InsertAnchorNamedRange = "LastCapExpendCol",
                .InsertIndexOffset = -1,
                .DeleteAnchorNamedRange = "Rep_CapExpend_010",
                .RecordCountNamedRange = "Rep_CapExpend_010",
                .RecordCountAdjustment = -1,
                .MinimumRecordCount = 3,
                .TransactionDBSyncNamedRange = "Rep_CapExpend_010"
            }
            'Insert before the immediate zero-width template column. This keeps
            'the insertion inside both the assumptions ranges and the linked
            'OFA Additions E:I formula ranges, allowing workbook dependencies to
            'expand before Transactional DB mirror rows are added.
            '
            'Use that template for formulas/formats, but the preceding visible
            'logical record for the assumptions column width.
            CapExRule.Targets.Add(New WorkbookStructureTarget With {
                .WorksheetName = "Capital Expenditure Assumptions",
                .CopyMode = WorkbookStructureCopyMode.AllAndColumnWidth,
                .TemplateOffset = 0,
                .ColumnWidthTemplateOffset = -1
            })

            CapExRule.Targets.Add(New WorkbookStructureTarget With {
                .WorksheetName = "OFA Additions",
                .CopyMode = WorkbookStructureCopyMode.AllAndColumnWidth,
                .TemplateOffset = 0
            })

            RetRules.Add(CapExRule.RuleID, CapExRule)

            '-----------------------------------------------------------------
            ' Capital Grant
            ' VBA: Capital Grant Assumptions + Capital Grant Workings
            '-----------------------------------------------------------------
            Dim CapGrantRule As New WorkbookStructureRule With {
                .RuleID = RuleCapGrantRecords,
                .Description = "Capital Grant records",
                .Axis = WorkbookStructureAxis.Columns,
                .InsertAnchorNamedRange = "LastCapGrantCol",
                .InsertIndexOffset = -1,
                .DeleteAnchorNamedRange = "Rep_CapGrant_010",
                .RecordCountNamedRange = "CapGrantInclusion",
                .RecordCountAdjustment = -1,
                .MinimumRecordCount = 3,
                .TransactionDBSyncNamedRange = "CapGrantInclusion"
            }

            'Insert immediately before the final hidden/template grant column.
            'That places the insertion inside the workbook-owned D:H grant ranges,
            'so Excel/DevExpress expands dependent workings names and Transactional
            'DB formula references naturally.  Inserting after the template leaves
            'those references at D:H and new mirror rows then return #REF!.
            '
            'Use the hidden template for formulas/formats and the preceding visible
            'record for the assumptions column width.
            CapGrantRule.Targets.Add(New WorkbookStructureTarget With {
                .WorksheetName = "Capital Grant Assumptions",
                .CopyMode = WorkbookStructureCopyMode.AllAndColumnWidth,
                .TemplateOffset = 0,
                .ColumnWidthTemplateOffset = -1
            })

            CapGrantRule.Targets.Add(New WorkbookStructureTarget With {
                .WorksheetName = "Capital Grant Workings",
                .CopyMode = WorkbookStructureCopyMode.AllAndColumnWidth,
                .TemplateOffset = 0
            })

            RetRules.Add(CapGrantRule.RuleID, CapGrantRule)

            '-----------------------------------------------------------------
            ' Repairs / stock-condition categories
            ' VBA RepairsSheets():
            '-----------------------------------------------------------------
            Dim RepairsRule As New WorkbookStructureRule With {
                .RuleID = RuleRepairsRecords,
                .Description = "Repairs and Maintenance categories",
                .Axis = WorkbookStructureAxis.Columns,
                .InsertAnchorNamedRange = "LastStockCol",
                .DeleteAnchorNamedRange = "StockCondCats",
                .RecordCountNamedRange = "StockCondCats",
                .MinimumRecordCount = 5,
                .TransactionDBSyncNamedRange = "StockCondCats"
            }
            'StockCondCats ends with a zero-width template column immediately
            'before the Q calculation-total column. Keep that column as the
            'formula/format template, but use the preceding visible category for
            'the width of newly-added assumption columns.
            RepairsRule.Targets.Add(New WorkbookStructureTarget With {
                .WorksheetName = "Repairs & Maint. Assumptions",
                .CopyMode = WorkbookStructureCopyMode.AllAndColumnWidth,
                .TemplateOffset = -1,
                .ColumnWidthTemplateOffset = -2
            })
            AddColumnTargets(RepairsRule,
                             "Stock Condition Inputs",
                             "Repairs & Maint. Rates",
                             "Stock Condition Results",
                             "Repairs & Maint. Drivers",
                             "Repairs & Maintenance Costs",
                             "Repairs & Maintenance Depn",
                             "Cost & Depn on Replacement")
            RetRules.Add(RepairsRule.RuleID, RepairsRule)

            '-----------------------------------------------------------------
            ' Housing components
            ' VBA ComponentSheets():
            '-----------------------------------------------------------------
            'DepnType includes the final structural/template record which the
            'Housing Asset interface excludes with SkipLastRecords=1.
            Dim ComponentsRule As New WorkbookStructureRule With {
                .RuleID = RuleHousingComponentRecords,
                .Description = "Housing Asset component records",
                .Axis = WorkbookStructureAxis.Columns,
                .InsertAnchorNamedRange = "LastCompCol",
                .InsertIndexOffset = -1,
                .DeleteAnchorNamedRange = "DepnType",
                .RecordCountNamedRange = "DepnType",
                .RecordCountAdjustment = -1,
                .MinimumRecordCount = 3,
                .TransactionDBSyncNamedRange = "DepnType"
            }
            AddColumnTargets(ComponentsRule,
                             "Housing Asset Assumptions",
                             "Depn Type Stock Numbers",
                             "Component 1",
                             "Component 2",
                             "Component 3",
                             "Component 4",
                             "Component 5",
                             "Component 6",
                             "Component 7",
                             "Component 8",
                             "Component 9",
                             "Component 10",
                             "Component 11",
                             "Component 12",
                             "Component Totals")

            'Every component sheet uses the final DepnType column as its copy
            'template. Because insertion now occurs before that column, resolve
            'the template at the insertion point (it shifts right during insert).
            For Each Target As WorkbookStructureTarget In ComponentsRule.Targets
                Target.TemplateOffset = 0
            Next
            RetRules.Add(ComponentsRule.RuleID, ComponentsRule)

            '-----------------------------------------------------------------
            ' Funding facilities
            ' VBA FundingSheets().  This is intentionally a single rule because
            ' the macro inserts/deletes the same physical facility column through
            ' all of the linked funding calculation sheets.
            'The existing bulk-mutation guard selects Recursive only while the
            'linked structure is inconsistent, then restores the entry engine.
            'Do not maintain the chain repeatedly between these physical shifts.
            '-----------------------------------------------------------------
            Dim FundingRule As New WorkbookStructureRule With {
                .RuleID = RuleFundingRecords,
                .Description = "Funding facility records",
                .UseRecursiveEngineForMutation = True,
                .Axis = WorkbookStructureAxis.Columns,
                .InsertAnchorNamedRange = "LoanDescRev1",
                .InsertIndexOffset = -1,
                .InsertAllColumnsBeforeCopy = True,
                .PreserveThreeDReferences = True,
                .ProtectedLeadingRecordCount = 10,
                .DeleteAnchorNamedRange = "FacilityNames",
                .RecordCountNamedRange = "FacilityNames",
                .MinimumRecordCount = 10,
                .TransactionDBSyncNamedRange = "FacilityNames"
            }
            AddColumnTargets(FundingRule,
                             "Funding Assumptions",
                             "Hidden- Intermed Interest Rates",
                             "Interest Rates",
                             "Hidden - Revolver Op Bal",
                             "Hidden - Revolver Adj Op Bal",
                             "Hidden - Revolver Bal Pre Int",
                             "Hidden - Revolver Close Bal",
                             "Hidden - Revolver Draw Repay",
                             "Hidden - Open Loan Facilities",
                             "Hidden - Close Loan Facilities",
                             "Hidden - Loan Opening Balances",
                             "Hidden - Loan Drawdowns",
                             "Hidden - Loan Interest",
                             "Hidden - Loan Repayments",
                             "Hidden - Loan Interest Charge",
                             "Hidden - Interest Pay Dates",
                             "Hidden - Interest Paid Factors",
                             "Hidden - Interest Accrual",
                             "Hidden - Loan Fees Amortisation",
                             "Hidden - Premium Amortisation",
                             "Hidden - Loan Commitment Fees",
                             "Loan Interest Payable",
                             "Loan Closing Facilities",
                             "Loan Opening Balances",
                             "Loan Drawdowns",
                             "Loan Interest Paid",
                             "Loan Repayments",
                             "Loan Closing Balances",
                             "Loan Fixed Variable",
                             "Loan Commitment Fees",
                             "Loan Fees Amortisation",
                             "Bond Premium Amortisation")
            'Funding_Columns inserts BEFORE the ordinary-loan template and copies
            'that original column only after EVERY linked sheet has been shifted.
            For Each Target As WorkbookStructureTarget In FundingRule.Targets
                Target.TemplateOffset = 0
            Next
            RetRules.Add(FundingRule.RuleID, FundingRule)

            '-----------------------------------------------------------------
            ' Development - Identified schemes
            ' VBA DvptSheets().  HouseTypeInID contains the final/template column,
            ' hence the -1 count adjustment used by Summit_Compatibility.
            '-----------------------------------------------------------------
            Dim DvptIDRule As New WorkbookStructureRule With {
                .RuleID = RuleDevelopmentIdentifiedRecords,
                .Description = "Identified Development records",
                .UseRecursiveEngineForMutation = True,
                .Axis = WorkbookStructureAxis.Columns,
                .InsertAnchorNamedRange = "LastIDColNum",
                .InsertAllColumnsBeforeCopy = True,
                .InsertFirstColumnSeparately = True,
                .PreserveThreeDReferences = True,
                .ProtectedLeadingRecordCount = 10,
                .DeleteAnchorNamedRange = "HouseTypeInID",
                .RecordCountNamedRange = "HouseTypeInID",
                .RecordCountAdjustment = -1,
                .MinimumRecordCount = 10,
                .TransactionDBSyncNamedRange = "HouseTypeInID"
            }
            AddDevelopmentTargets(DvptIDRule)
            RetRules.Add(DvptIDRule.RuleID, DvptIDRule)

            '-----------------------------------------------------------------
            ' Development - Multi-year schemes
            '-----------------------------------------------------------------
            Dim DvptMYRule As New WorkbookStructureRule With {
                .RuleID = RuleDevelopmentMultiYearRecords,
                .Description = "Multi-year Development records",
                .UseRecursiveEngineForMutation = True,
                .Axis = WorkbookStructureAxis.Columns,
                .InsertAnchorNamedRange = "LastMYColNum",
                .InsertAllColumnsBeforeCopy = True,
                .PreserveThreeDReferences = True,
                .ProtectedLeadingRecordCount = 3,
                .DeleteAnchorNamedRange = "HouseTypeInMY",
                .RecordCountNamedRange = "HouseTypeInMY",
                .RecordCountAdjustment = -1,
                .MinimumRecordCount = 3,
                .TransactionDBSyncNamedRange = "HouseTypeInMY"
            }
            AddDevelopmentTargets(DvptMYRule)
            RetRules.Add(DvptMYRule.RuleID, DvptMYRule)

            '-----------------------------------------------------------------
            ' Journal assumptions
            '
            'The workbook keeps one final blank/sentinel row in Rep_Jour_01.
            'The interface already excludes it with SkipLastRecords=1.
            '
            'Unlike the column-based structures above there is no separate
            '"Last..." named range.  The insertion point is therefore the END
            'of Rep_Jour_01, i.e. immediately before the final sentinel row.
            '
            'The row immediately above that insertion point is a fully formed
            'journal row containing the hidden debit/credit formulas, making it
            'the correct copy template.
            '-----------------------------------------------------------------
            Dim JournalRule As New WorkbookStructureRule With {
                .RuleID = RuleJournalRecords,
                .Description = "Journal records",
                .Axis = WorkbookStructureAxis.Rows,
                .InsertAnchorNamedRange = "Rep_Jour_01",
                .InsertAtAnchorEnd = True,
                .DeleteAnchorNamedRange = "Rep_Jour_01",
                .RecordCountNamedRange = "Rep_Jour_01",
                .RecordCountAdjustment = -1,
                .MinimumRecordCount = 1,
                .TransactionDBSyncNamedRange = "IR_Journals"
            }

            JournalRule.Targets.Add(New WorkbookStructureTarget With {
                .WorksheetName = "Journal Assumptions",
                .CopyMode = WorkbookStructureCopyMode.All,
                .TemplateOffset = -1,
                .ClearUnlockedConstants = True
            })

            RetRules.Add(JournalRule.RuleID, JournalRule)

            '-----------------------------------------------------------------
            ' Stock Conversion assumptions
            '
            'The interface already identifies IR_StockDispAss_01 as the row
            'expansion family. Unlike Journal there is no final sentinel row:
            'new records are appended AFTER the current named range.
            '
            'The previous final row is a complete conversion record and is used
            'as the copy template so formulas/formatting to the right of the
            'visible interface fields remain Excel-native.
            '-----------------------------------------------------------------
            Dim StockConversionRule As New WorkbookStructureRule With {
                .RuleID = RuleStockConversionRecords,
                .Description = "Stock Conversion records",
                .Axis = WorkbookStructureAxis.Rows,
                .InsertAnchorNamedRange = "IR_StockDispAss_01",
                .InsertAfterAnchorEnd = True,
                .DeleteAnchorNamedRange = "IR_StockDispAss_01",
                .RecordCountNamedRange = "IR_StockDispAss_01",
                .MinimumRecordCount = 1,
                .TransactionDBSyncNamedRange = "IR_StockDispAss_01"
            }

            StockConversionRule.Targets.Add(New WorkbookStructureTarget With {
                .WorksheetName = "Stock Conversion Assumptions",
                .CopyMode = WorkbookStructureCopyMode.All,
                .TemplateOffset = -1,
                .ClearUnlockedConstants = True
            })

            RetRules.Add(StockConversionRule.RuleID, StockConversionRule)

            '-----------------------------------------------------------------
            ' Remaining simple NRRI range families
            '
            'Default append rules for NRRI ranges. Reviewed coupled families are
            'configured below with their input boundaries and linked workings.
            '-----------------------------------------------------------------

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_REP_RENT_02",
                "Rent Weeks",
                "Rep_Rent_02",
                "Rent Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_REP_SERVCHG_01",
                "Summary Categories",
                "Rep_ServChg_01",
                "Service Charge Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_REP_SERVCHG_02",
                "Service and Support Charges",
                "Rep_ServChg_02",
                "Service Charge Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_STOCK_DISP_DESC",
                "Stock Disposal Assumptions",
                "IR_Stock_Disp_Desc",
                "Stock Disposal Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_OWN_OCC_DESC",
                "Owner Occupier Assumptions",
                "IR_Own_Occ_Desc",
                "Owner Occupier Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_SUMMARYOTHERINCCAT",
                "Specific Income Assumptions Categories",
                "SummaryOtherIncCat",
                "Specific Income Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_SPEC_INC_ASS1",
                "Specific Other Income",
                "IR_Spec_Inc_Ass1",
                "Specific Income Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_CAPITALGRANTCATS",
                "Summary Categories",
                "CapitalGrantCats",
                "Capital Grant Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_INTCO_INC_ASS1",
                "Annual Intercompany Income",
                "IR_Intco_Inc_Ass1",
                "Intercompany Income Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_INTCO_INC_ASS2",
                "Periodic Intercompany Income",
                "IR_Intco_Inc_Ass2",
                "Intercompany Income Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_ONEOFF_COST_ASS",
                "One Off Other Spend",
                "IR_Oneoff_Cost_Ass",
                "Management Costs Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_REP_REPHASE",
                "Rent Weeks",
                "IR_Rep_Rephase",
                "Repairs & Maint. Rephasing",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_ADDPENSION_01",
                "Additional Pension Assumptions",
                "IR_AddPension_01",
                "Additional Pension Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_ADDPENSION_02",
                "Other Financial Costs",
                "IR_AddPension_02",
                "Additional Pension Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_REP_ADDPENSION_050",
                "Actuarial Movements",
                "Rep_AddPension_050",
                "Additional Pension Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_DVPT_CAPINT",
                "Capitalised Interest",
                "IR_Dvpt_CapInt",
                "Development BP Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_RPI",
                "CPI and RPI",
                "IR_RPI",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_SERVCHG",
                "Real Service Charge",
                "IR_ServChg",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_RTBREAL",
                "Real RTB Valuation",
                "IR_RTBReal",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_OTHERINCREAL",
                "Real Other Income",
                "IR_OtherIncReal",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_CAPGRANTREAL",
                "Real Capital Grant",
                "IR_CapGrantReal",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_LEASEINCREAL",
                "Real Leaseholder Income",
                "IR_LeaseIncReal",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_STAFFCOSTREAL",
                "Real Staff Costs",
                "IR_StaffCostReal",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_BPREPAIRREAL",
                "Real R & M Costs",
                "IR_BPRepairReal",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_OTHEREXPREAL",
                "Real Other Expenditure",
                "IR_OtherExpReal",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_DEMCOSTREAL",
                "Real Demolition",
                "IR_DemCostReal",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_DEVRENTREAL",
                "Real Dvpt Rent",
                "IR_DevRentReal",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_DEVSERVICECHG",
                "Real Dvpt Service Charge",
                "IR_DevServiceChg",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_DEVMGTREAL",
                "Real Dvpt Mgmt Costs",
                "IR_DevMgtReal",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_DEVREPAIRSREAL",
                "Real Dvpt R & M Costs",
                "IR_DevRepairsReal",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_DEVOTHERREAL",
                "Real Dvpt Capital",
                "IR_DevOtherReal",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_FUNDFEESREAL",
                "Real Funding Fees",
                "IR_FundFeesReal",
                "Economic Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_MGT_COST_CAP",
                "Management Costs Capitalisation",
                "IR_Mgt_Cost_Cap",
                "Capitalisation Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_REPAIRS_CAP",
                "Repairs Maint. Costs Capitalisation",
                "IR_Repairs_Cap",
                "Capitalisation Assumptions",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_CASH_JOURNALS",
                "Cash Journals",
                "IR_Cash_Journals",
                "Cash Journals",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            AddSimpleAppendRangeRule(
                RetRules,
                "SIMPLE_IR_MVTWC",
                "Movements in Balance Sheet accounts - External",
                "IR_MvtWC",
                "Cash Journals",
                WorkbookStructureAxis.Rows,
                WorkbookStructureCopyMode.All)

            'Worksheet events coordinate these row/column families. The displayed
            'Rep range is not the physical input boundary (notably Service Charge).
            Dim ServiceRule = RetRules("SIMPLE_REP_SERVCHG_02")
            ServiceRule.InsertAnchorNamedRange = "IR_ServChg_01"
            ServiceRule.DeleteAnchorNamedRange = "IR_ServChg_01"
            ServiceRule.RecordCountNamedRange = "IR_ServChg_01"
            ServiceRule.TransactionDBSyncNamedRange = "IR_ServChg_01"
            ServiceRule.MinimumRecordCount = 4
            ServiceRule.AdditionalRecordRanges.Add("Rep_ServChg_02")
            ServiceRule.LinkedColumns = CreateLinkedColumns("LastUnitSCColumn",
                "Unit Service Charges", "Service Charge Numbers", "Service Charge Income",
                "Service Charge Voids", "Service Charge Bad Debts")

            RetRules("SIMPLE_IR_SPEC_INC_ASS1").MinimumRecordCount = 3
            RetRules("SIMPLE_IR_SPEC_INC_ASS1").LinkedColumns = CreateLinkedColumns("LastSIDriversCol",
                "Specific Income Drivers", "Specific Income Workings", "Specific Income Voids", "Specific Income Bad Debts")
            AddSimpleAppendRangeRule(RetRules, "OTHER_INCOME_RECORDS", "Other Income records",
                "IR_Oth_Inc_Ass5", "Other Income Assumptions", WorkbookStructureAxis.Rows, WorkbookStructureCopyMode.All)
            RetRules("OTHER_INCOME_RECORDS").MinimumRecordCount = 3
            RetRules("OTHER_INCOME_RECORDS").LinkedColumns = CreateLinkedColumns("LastOIWorkingsCol", "Other Income Workings")
            AddSimpleAppendRangeRule(RetRules, "OTHER_INCOME_ADJUSTMENTS", "Other Income adjustments",
                "IR_Oth_Inc_Ass6", "Other Income Assumptions", WorkbookStructureAxis.Rows, WorkbookStructureCopyMode.All)
            RetRules("OTHER_INCOME_ADJUSTMENTS").MinimumRecordCount = 3

            AddGroupedTemplateRule(RetRules, "JOINT_VENTURE_RECORDS", "Joint Venture records",
                "IC_JointVenture_01", "LastJointVentureCol", 3,
                "Joint Venture Assumptions", "JV Opening balances", "JV Investments", "JV Share of Profits",
                "JV Payments Repayments", "JV Closing Balances", "JV Interest Received")
            For Each Suffix In {"01", "02", "03", "04"}
                Dim InputRange = "IR_JointVenture_" & Suffix
                Dim ID = "JOINT_VENTURE_ROWS_" & Suffix
                AddSimpleAppendRangeRule(RetRules, ID, "Joint Venture input rows", InputRange,
                    "Joint Venture Assumptions", WorkbookStructureAxis.Rows, WorkbookStructureCopyMode.All)
                RetRules(ID).MinimumRecordCount = 3
                RetRules(ID).AdditionalRecordRanges.Add(InputRange & "a")
            Next
            For Each Suffix In {"01", "02"}
                Dim InputRange = "IC_ServChg_" & Suffix
                Dim ID = "SERVICE_CHARGE_COLUMNS_" & Suffix
                AddSimpleAppendRangeRule(RetRules, ID, "Service Charge input columns", InputRange,
                    "Service Charge Assumptions", WorkbookStructureAxis.Columns, WorkbookStructureCopyMode.FormatsAndColumnWidth)
                RetRules(ID).MinimumRecordCount = 2
                RetRules(ID).AdditionalRecordRanges.Add(InputRange & "a")
                RetRules(ID).FillRelativeHeaderRowOffset = 2
            Next
            Dim IntercoSheets = {"Interco Funding Assumptions", "Hidden - InterCo Int Rates",
                "Hidden - InterCo Opening Bal", "Hidden - InterCo Increases", "Hidden - InterCo Decreases",
                "Hidden - InterCo Closing Bal", "Hidden - InterCo Interest", "InterCo Opening Balances",
                "InterCo Increases", "InterCo Decreases", "InterCo Closing Balances", "InterCo Interest"}
            AddGroupedTemplateRule(RetRules, "INTERCO_LOAN_RECORDS", "Intercompany loans",
                "IC_IntercoFunding_01", "LastIntercoLoanAssCol", 3, IntercoSheets)
            AddGroupedTemplateRule(RetRules, "INTERCO_INVESTMENT_RECORDS", "Intercompany investments",
                "IC_IntercoFunding_02", "LastIntercoInvestAssCol", 3, IntercoSheets)
            'The actual exposed dates are on Interco Funding Assumptions. Do not
            'substitute the older, different Funding Assumptions year-insertion macro.
            For Each DateRange In {"DateInterF01", "DateInterF02", "DateInterF03"}
                Dim ID = "INTERCO_DATES_" & DateRange.ToUpperInvariant()
                AddSimpleAppendRangeRule(RetRules, ID, "Intercompany dates", DateRange,
                    "Interco Funding Assumptions", WorkbookStructureAxis.Rows, WorkbookStructureCopyMode.All)
                RetRules(ID).MinimumRecordCount = 5
                'DX 25.2's chain engine can retain invalid shared-formula
                'precedents after these rows move. Use its supported recursive
                'engine only for this bulk command and restore entry state.
                RetRules(ID).UseRecursiveEngineForMutation = True
            Next

            AddGroupedTemplateRule(RetRules, "SPECIFIC_INCOME_CATEGORIES", "Specific Income categories",
                "Rep_SInc_01", "LastSIAssumpCol", 3, "Specific Income Assumptions")
            AddGroupedTemplateRule(RetRules, "OTHER_INCOME_CATEGORIES", "Other Income categories",
                "Rep_OInc_04", "LastOIAssumpCol", 2, "Other Income Assumptions")

            Return RetRules

        End Function

        Private Sub AddSimpleAppendRangeRule(
            ByVal RetRules As Dictionary(Of String, WorkbookStructureRule),
            ByVal RuleID As String,
            ByVal Description As String,
            ByVal NamedRange As String,
            ByVal WorksheetName As String,
            ByVal Axis As WorkbookStructureAxis,
            ByVal CopyMode As WorkbookStructureCopyMode)

            If RetRules Is Nothing Then Return
            If String.IsNullOrWhiteSpace(RuleID) Then Return
            If String.IsNullOrWhiteSpace(NamedRange) Then Return
            If String.IsNullOrWhiteSpace(WorksheetName) Then Return

            Dim Rule As New WorkbookStructureRule With {
                .RuleID = RuleID,
                .Description = Description,
                .Axis = Axis,
                .InsertAnchorNamedRange = NamedRange,
                .InsertAfterAnchorEnd = True,
                .DeleteAnchorNamedRange = NamedRange,
                .RecordCountNamedRange = NamedRange,
                .MinimumRecordCount = 1,
                .TransactionDBSyncNamedRange = NamedRange
            }

            Rule.Targets.Add(New WorkbookStructureTarget With {
                .WorksheetName = WorksheetName,
                .CopyMode = CopyMode,
                .TemplateOffset = -1,
                .ClearUnlockedConstants = (Axis = WorkbookStructureAxis.Rows)
            })

            RetRules.Add(Rule.RuleID, Rule)

        End Sub

        Private Shared Function CreateLinkedColumns(Anchor As String, ParamArray Sheets() As String) As WorkbookLinkedColumns
            Dim Linked As New WorkbookLinkedColumns With {.EndAnchorNamedRange = Anchor}
            For Each Sheet In Sheets
                Linked.Targets.Add(New WorkbookStructureTarget With {
                    .WorksheetName = Sheet, .CopyMode = WorkbookStructureCopyMode.AllAndColumnWidth,
                    .TemplateOffset = -1, .AdditionalCopiedColumns = 1})
            Next
            Return Linked
        End Function

        Private Sub AddGroupedTemplateRule(Rules As Dictionary(Of String, WorkbookStructureRule),
            ID As String, Description As String, Records As String, Anchor As String, Minimum As Integer,
            ParamArray Sheets() As String)
            Dim Rule As New WorkbookStructureRule With {
                .RuleID = ID, .Description = Description, .Axis = WorkbookStructureAxis.Columns,
                .RecordCountNamedRange = Records, .DeleteAnchorNamedRange = Records,
                .InsertAnchorNamedRange = Anchor, .InsertIndexOffset = -1,
                .MinimumRecordCount = Minimum, .TransactionDBSyncNamedRange = Records,
                .InsertAllColumnsBeforeCopy = True, .PreserveThreeDReferences = True}
            For Index As Integer = 0 To Sheets.Length - 1
                Rule.Targets.Add(New WorkbookStructureTarget With {
                    .WorksheetName = Sheets(Index), .TemplateOffset = 0,
                    .CopyMode = WorkbookStructureCopyMode.AllAndColumnWidth,
                    .AdditionalCopiedColumns = 1, .ClearUnlockedConstants = (Index = 0),
                    .InputClearColumnOffset = 1})
            Next
            Rules.Add(ID, Rule)
        End Sub

        Private Sub AddDevelopmentTargets(ByVal Rule As WorkbookStructureRule)

            'Match DvptSheets() and insert EVERY target before copying. NonCash
            'and Component Depn reference each other: reordering alone cannot
            'preserve both directions. Identified schemes use VBA's first-one,
            'then-remainder batches, under one transaction and one mirror sync.
            AddColumnTargets(Rule,
                             "Development BP Assumptions",
                             "Development Stock",
                             "Development Capital",
                             "Development Revenue",
                             "Development Expenditure",
                             "Dvpt NonCash",
                             "Dvpt Component Depn")

        End Sub

        Private Sub AddColumnTargets(ByVal Rule As WorkbookStructureRule,
                                     ParamArray WorksheetNames() As String)

            If Rule Is Nothing OrElse WorksheetNames Is Nothing Then Return

            For Each WorksheetName As String In WorksheetNames

                Rule.Targets.Add(New WorkbookStructureTarget With {
                    .WorksheetName = WorksheetName,
                    .CopyMode = WorkbookStructureCopyMode.AllAndColumnWidth,
                    .TemplateOffset = -1
                })

            Next

        End Sub

        Private Function ExecuteInsert(ByVal Rule As WorkbookStructureRule,
                                       ByVal RecordCount As Integer) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}

            If IsExecuting Then
                Result.BError = True
                Result.StringReturn = "A workbook structural operation is already in progress."
                Return Result
            End If

            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            Dim Anchor As DefinedName = WB.DefinedNames.GetDefinedName(Rule.InsertAnchorNamedRange)

            If Anchor Is Nothing OrElse Anchor.Range Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Insert anchor named range '" & Rule.InsertAnchorNamedRange & "' was not found."
                Return Result
            End If

            Dim InsertIndex As Integer

            Select Case Rule.Axis

                Case WorkbookStructureAxis.Columns

                    If Rule.InsertAfterAnchorEnd Then
                        InsertIndex = Anchor.Range.RightColumnIndex + 1
                    ElseIf Rule.InsertAtAnchorEnd Then
                        InsertIndex = Anchor.Range.RightColumnIndex
                    Else
                        InsertIndex = Anchor.Range.LeftColumnIndex
                    End If

                Case WorkbookStructureAxis.Rows

                    If Rule.InsertAfterAnchorEnd Then
                        InsertIndex = Anchor.Range.BottomRowIndex + 1
                    ElseIf Rule.InsertAtAnchorEnd Then
                        InsertIndex = Anchor.Range.BottomRowIndex
                    Else
                        InsertIndex = Anchor.Range.TopRowIndex
                    End If

                Case Else
                    Result.BError = True
                    Result.StringReturn = "Unsupported structural axis."
                    Return Result

            End Select

            InsertIndex += Rule.InsertIndexOffset

            If InsertIndex < 0 Then
                Result.BError = True
                Result.StringReturn = "The insertion position for '" & Rule.Description & "' is invalid."
                Return Result
            End If

            'Snapshot the logical record-count range before the structural edit.
            '
            'DevExpress does not always expand a named range when columns/rows are
            'inserted immediately adjacent to its current edge.  The interface
            'presentation is rebuilt from these named ranges, so a stale range can
            'leave the workbook structurally correct while the interface still shows
            'the old record count.
            Dim RecordRangeSnapshot As StructuralNamedRangeSnapshot =
                SnapshotNamedRange(WB, Rule.RecordCountNamedRange)

            Dim JournalInputSnapshot As StructuralNamedRangeSnapshot = Nothing

            For Each Target As WorkbookStructureTarget In Rule.Targets
                If GetWorksheet(WB, Target.WorksheetName) Is Nothing Then
                    Result.BError = True
                    Result.StringReturn = "Worksheet '" & Target.WorksheetName & "' was not found."
                    Result.StrResponseMessage = Result.StringReturn
                    Return Result
                End If
            Next

            Dim ChangedWorksheets As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim UpdateStarted As Boolean = False
            Dim MutationStarted As Boolean = False
            Dim BulkMutationGuardStarted As Boolean = False
            Dim PreviousCalculationMode As WorkbookCalculationMode = WB.Options.CalculationMode
            Dim PreviousCalculationEngine As CalculationEngineType = WB.Options.CalculationEngineType
            Dim benchmark As System.Diagnostics.Stopwatch =
                System.Diagnostics.Stopwatch.StartNew()
            Dim mutationMs As Long = 0
            Dim DeferredColumnCopies As List(Of Action) =
                If(Rule.InsertAllColumnsBeforeCopy, New List(Of Action)(), Nothing)
            Dim ThreeDReferences As WorkbookStructural3DReferences = Nothing
            Dim PreSyncedTransactionDBResult As AbovoTransaction = Nothing
            Dim preSyncMs As Long = 0
            Dim ReuseDevelopmentCapture As Boolean =
                Rule.RuleID = RuleDevelopmentIdentifiedRecords AndAlso RecordCount > 1 AndAlso
                Rule.InsertFirstColumnSeparately AndAlso Rule.InsertAllColumnsBeforeCopy AndAlso
                Rule.PreserveThreeDReferences AndAlso Rule.Axis = WorkbookStructureAxis.Columns AndAlso
                Rule.Targets.All(Function(t) t.TemplateOffset = -1 AndAlso t.AdditionalCopiedColumns = 0)

            'Detailed tracing is limited to the three insertion families under review.
            'These scopes only measure existing work; they do not read formula values.
            Using Timing As StructuralInsertBenchmark =
                If(Rule.RuleID = RuleFundingRecords OrElse
                   Rule.RuleID = RuleDevelopmentIdentifiedRecords OrElse
                   Rule.RuleID = RuleDevelopmentMultiYearRecords,
                   New StructuralInsertBenchmark(ModelID, Rule.RuleID, RecordCount,
                       "sheets=" & Rule.Targets.Count.ToString() &
                       ", insertColumn=" & (InsertIndex + 1).ToString() &
                       ", calculationMode=" & PreviousCalculationMode.ToString() &
                       ", engine=" & PreviousCalculationEngine.ToString()), Nothing)
                Try

                    JournalInputSnapshot = SnapshotJournalInputRange(WB, Rule, RecordRangeSnapshot)

                    Dim AdditionalSnapshots = Rule.AdditionalRecordRanges.Select(Function(n) SnapshotNamedRange(WB, n)).ToList()
                    If RecordRangeSnapshot Is Nothing OrElse AdditionalSnapshots.Any(Function(s) s Is Nothing) Then Throw New InvalidOperationException("A required input/display range is missing. No workbook ranges were changed.")
                    Dim LinkedEnd = ValidateLinkedColumns(WB, Rule)
                    If Rule.LinkedColumns IsNot Nothing Then
                        'Preflight unsupported 3-D spans before the primary rows change.
                        WorkbookStructural3DReferences.Capture(WB, Rule.LinkedColumns.Targets.Select(Function(t) t.WorksheetName),
                            {New KeyValuePair(Of Integer, Integer)(LinkedEnd - 1, RecordCount)})
                    End If

                    If Rule.PreserveThreeDReferences Then
                        Using Stage = Timing?.Measure("capture3D", "batch=1")
                            ThreeDReferences = WorkbookStructural3DReferences.Capture(WB,
                                Rule.Targets.Select(Function(t) t.WorksheetName),
                                {New KeyValuePair(Of Integer, Integer)(InsertIndex, If(Rule.InsertFirstColumnSeparately, 1, RecordCount))})
                        End Using
                    End If

                    Using Stage = Timing?.Measure("beginMutation")
                        IsExecuting = True
                        ModelSafetyManager.BeginBulkWorkbookMutation(ModelID)
                        BulkMutationGuardStarted = True

                        'Match deletion and the original VBA structural routines: do not
                        'recalculate the workbook between linked-sheet insertions.
                        WB.Options.CalculationMode = WorkbookCalculationMode.Manual
                        If Rule.UseRecursiveEngineForMutation Then WB.Options.CalculationEngineType = CalculationEngineType.Recursive

                        WB.BeginUpdate()
                        UpdateStarted = True
                    End Using

                    Dim Batches As New List(Of KeyValuePair(Of Integer, Integer))
                    If Rule.InsertFirstColumnSeparately AndAlso RecordCount > 1 Then
                        Batches.Add(New KeyValuePair(Of Integer, Integer)(InsertIndex, 1))
                        Batches.Add(New KeyValuePair(Of Integer, Integer)(InsertIndex + 1, RecordCount - 1))
                    Else
                        Batches.Add(New KeyValuePair(Of Integer, Integer)(InsertIndex, RecordCount))
                    End If

                    Dim BatchNumber As Integer = 0
                    For Each Batch In Batches
                        BatchNumber += 1
                        If DeferredColumnCopies IsNot Nothing Then DeferredColumnCopies.Clear()
                        'Carry forward the first capture for the reviewed adjacent
                        'Development batch. Include the copied template column as
                        'well as the new column; all other rules retain full capture.
                        If Rule.PreserveThreeDReferences AndAlso Batch.Key <> InsertIndex Then
                            If ReuseDevelopmentCapture Then
                                Using Stage = Timing?.Measure("advance3D", "batch=" & BatchNumber.ToString())
                                    Dim CopiedRanges As New List(Of CellRange)
                                    For Each Target In Rule.Targets
                                        Dim WS = GetWorksheet(WB, Target.WorksheetName), Used = WS.GetUsedRange()
                                        CopiedRanges.Add(WS.Range.FromLTRB(InsertIndex, Used.TopRowIndex,
                                            InsertIndex + Target.AdditionalCopiedColumns, Used.BottomRowIndex))
                                    Next
                                    ThreeDReferences = ThreeDReferences.CaptureFollowingInsertion(Batch.Key, Batch.Value, CopiedRanges)
                                End Using
                            Else
                                Using Stage = Timing?.Measure("capture3D", "batch=" & BatchNumber.ToString())
                                    ThreeDReferences = WorkbookStructural3DReferences.Capture(WB,
                                        Rule.Targets.Select(Function(t) t.WorksheetName), {Batch})
                                End Using
                            End If
                        End If
                        For Each Target As WorkbookStructureTarget In Rule.Targets

                            Dim WS As Worksheet = GetWorksheet(WB, Target.WorksheetName)

                            If WS Is Nothing Then
                                Throw New InvalidOperationException("Worksheet '" & Target.WorksheetName & "' was not found.")
                            End If

                            Dim WasProtected As Boolean = WS.IsProtected
                            Dim WasVisible As Boolean = WS.Visible

                            Try

                                WS.Visible = True
                                If WasProtected Then
                                    Using Stage = Timing?.Measure("unprotect", "phase=shift, batch=" & BatchNumber.ToString() & ", sheet=" & WS.Name)
                                        UNProtectWS(ModelID, WS.Name)
                                    End Using
                                End If

                                Select Case Rule.Axis

                                    Case WorkbookStructureAxis.Columns
                                        MutationStarted = True
                                        InsertColumnsForTarget(WS, Batch.Key, Batch.Value, Target, DeferredColumnCopies, Timing, BatchNumber)

                                    Case WorkbookStructureAxis.Rows
                                        MutationStarted = True
                                        InsertRowsForTarget(WS, Batch.Key, Batch.Value, Target)

                                End Select

                                ChangedWorksheets.Add(WS.Name)

                            Finally

                                If WasProtected Then
                                    Using Stage = Timing?.Measure("protect", "phase=shift, batch=" & BatchNumber.ToString() & ", sheet=" & WS.Name)
                                        ProtectWS(ModelID, WS.Name)
                                    End Using
                                End If
                                WS.Visible = WasVisible

                            End Try

                        Next

                        'A formula copied before another linked sheet is shifted can be
                        'rewritten to the wrong record. Match Excel's grouped
                        'insert: finish all physical insertions, THEN copy each template.
                        If ThreeDReferences IsNot Nothing Then
                            Using Stage = Timing?.Measure("apply3D", "batch=" & BatchNumber.ToString())
                                ThreeDReferences.Apply(ModelID)
                            End Using
                        End If
                        If DeferredColumnCopies IsNot Nothing Then
                            For Each CopyColumns As Action In DeferredColumnCopies
                                CopyColumns()
                            Next
                        End If
                    Next

                    If ReuseDevelopmentCapture Then
                        Using Stage = Timing?.Measure("verify3D")
                            ThreeDReferences.Verify()
                        End Using
                    End If

                    'Force the interface-driving named range to the intended new size.
                    'This is deliberately done before TransactionDB synchronisation and
                    'interface dependency invalidation.
                    Using Stage = Timing?.Measure("resizeNamesAndLinkedRanges")
                        ResizeNamedRangeFromSnapshot(WB,
                                                     RecordRangeSnapshot,
                                                     Rule.Axis,
                                                     RecordCount)
                        ResizeNamedRangeFromSnapshot(WB, JournalInputSnapshot, Rule.Axis, RecordCount)
                        For Each Snapshot In AdditionalSnapshots
                            ResizeNamedRangeFromSnapshot(WB, Snapshot, Rule.Axis, RecordCount)
                        Next
                        If Rule.FillRelativeHeaderRowOffset >= 0 Then
                            Dim CurrentRange = WB.DefinedNames.GetDefinedName(Rule.RecordCountNamedRange).Range
                            Dim WS = CurrentRange.Worksheet, WasProtected = WS.IsProtected
                            Try
                                If WasProtected Then UNProtectWS(ModelID, WS.Name)
                                Dim Row = CurrentRange.TopRowIndex + Rule.FillRelativeHeaderRowOffset
                                WS.Range.FromLTRB(CurrentRange.LeftColumnIndex + 1, Row, CurrentRange.RightColumnIndex, Row).
                                    CopyFrom(WS.Cells(Row, CurrentRange.LeftColumnIndex), PasteSpecial.All)
                            Finally
                                If WasProtected Then ProtectWS(ModelID, WS.Name)
                            End Try
                        End If
                        If Rule.LinkedColumns IsNot Nothing Then
                            InsertLinkedColumns(WB, Rule.LinkedColumns, LinkedEnd - 1, RecordCount, ChangedWorksheets)
                        End If
                    End Using
                    'Development already uses Recursive for source shifts.
                    'Keep that engine through TDB instead of rebuilding ChainBased,
                    'immediately switching back, and rebuilding it a second time.
                    'Finish the source update first; the synchroniser retains its own
                    'update/history/recovery guards. The existing Finally restores
                    'our entry engine/mode before dependency/UI notifications run.
                    'Funding retains its original sequencing: paired trials did
                    'not demonstrate a repeatable gain from sharing this scope.
                    If Rule.UseRecursiveEngineForMutation AndAlso
                       (Rule.RuleID = RuleDevelopmentIdentifiedRecords OrElse
                        Rule.RuleID = RuleDevelopmentMultiYearRecords) Then
                        Using Stage = Timing?.Measure("endSourceUpdate")
                            UpdateStarted = False
                            WB.EndUpdate()
                        End Using
                        Dim syncStartMs As Long = benchmark.ElapsedMilliseconds
                        Try
                            Using Stage = Timing?.Measure("synchroniseTDB", "engineShared=True")
                                PreSyncedTransactionDBResult = SynchroniseTransactionDB(Rule)
                            End Using
                        Finally
                            preSyncMs = benchmark.ElapsedMilliseconds - syncStartMs
                        End Try
                    End If
                    Result.BError = False
                    Result.EventCancelled = False
                    Result.StringReturn = RecordCount.ToString & " " & Rule.Description & " added."

                Catch ex As Exception

                    Result.BError = True
                    Result.StringReturn = ex.Message
                    Result.StrResponseMessage = ex.Message
                    System.Diagnostics.Trace.WriteLine("[Structure failure] " & Rule.RuleID & ": " & ex.ToString())
                    If MutationStarted Then
                        ModelSafetyManager.MarkRecoveryRequired(
                            ModelID, "Insert " & Rule.Description, ex.Message,
                            "Workbook Structure", Rule.RuleID)
                    End If

                Finally

                    Try
                        Using Stage = Timing?.Measure("endUpdate")
                            If UpdateStarted Then WB.EndUpdate()
                        End Using
                    Catch ex As Exception
                        Result.BError = True
                        Result.StringReturn &= Environment.NewLine & "End workbook update: " & ex.Message
                        If MutationStarted Then
                            ModelSafetyManager.MarkRecoveryRequired(
                                ModelID, "Insert " & Rule.Description, Result.StringReturn,
                                "Workbook Structure", Rule.RuleID)
                        End If
                    End Try
                    Try
                        Using Stage = Timing?.Measure("restoreEngine")
                            If Rule.UseRecursiveEngineForMutation Then WB.Options.CalculationEngineType = PreviousCalculationEngine
                        End Using
                    Catch ex As Exception
                        Result.BError = True
                        Result.StringReturn &= Environment.NewLine & "Restore calculation engine: " & ex.Message
                        If MutationStarted Then ModelSafetyManager.MarkRecoveryRequired(ModelID, "Restore calculation engine", Result.StringReturn, "Workbook Structure", Rule.RuleID)
                    End Try
                    Try
                        Using Stage = Timing?.Measure("restoreCalculationMode")
                            WB.Options.CalculationMode = PreviousCalculationMode
                        End Using
                    Catch ex As Exception
                        Result.BError = True
                        Result.StringReturn &= Environment.NewLine & "Restore calculation mode: " & ex.Message
                    End Try
                    IsExecuting = False
                    If BulkMutationGuardStarted Then
                        Using Stage = Timing?.Measure("endMutationGuard")
                            ModelSafetyManager.EndBulkWorkbookMutation(ModelID)
                            BulkMutationGuardStarted = False
                        End Using
                    End If

                End Try

                mutationMs = benchmark.ElapsedMilliseconds - preSyncMs
                If Not Result.BError Then

                    Dim PostActionResult As AbovoTransaction
                    Using Stage = Timing?.Measure("postActions")
                        PostActionResult = RunPostActions(Rule, ChangedWorksheets,
                                                         PreSyncedTransactionDBResult, preSyncMs)
                    End Using

                    If PostActionResult.BError Then
                        Result.BError = True
                        Result.BSuccess = False
                        Result.StringReturn =
                            "Workbook structure changed, but Transactional DB synchronisation failed: " &
                            PostActionResult.StringReturn
                        Result.StrResponseMessage = Result.StringReturn
                        ModelSafetyManager.MarkRecoveryRequired(
                            ModelID, "Insert " & Rule.Description, Result.StringReturn,
                            "Workbook Structure", Rule.RuleID)
                    End If

                End If

                If MutationStarted Then ExcelModels(ModelID).IsDirty = True
                System.Diagnostics.Trace.WriteLine(
                    "[Population Benchmark] Structure insert: model=" & ModelID.ToString() &
                    ", rule=" & Rule.RuleID &
                    ", records=" & RecordCount.ToString() &
                    ", workbookMutation=" & mutationMs.ToString() & " ms" &
                    ", postActions=" &
                    (benchmark.ElapsedMilliseconds - mutationMs).ToString() & " ms" &
                    ", total=" & benchmark.ElapsedMilliseconds.ToString() & " ms" &
                    ", outcome=" & If(Result.BError, "failed", "ok"))
                Timing?.Complete(Result.BError)
                Return Result
            End Using

        End Function

        Private Function BuildDeleteBlocks(ByVal RecordIndexes As IEnumerable(Of Integer)) As List(Of StructuralDeleteBlock)

            Dim SortedIndexes As New List(Of Integer)

            If RecordIndexes IsNot Nothing Then
                For Each RecordIndex As Integer In RecordIndexes
                    If Not SortedIndexes.Contains(RecordIndex) Then SortedIndexes.Add(RecordIndex)
                Next
            End If

            SortedIndexes.Sort()

            Dim Blocks As New List(Of StructuralDeleteBlock)

            If SortedIndexes.Count = 0 Then Return Blocks

            Dim BlockStart As Integer = SortedIndexes(0)
            Dim PreviousIndex As Integer = SortedIndexes(0)
            Dim BlockCount As Integer = 1

            For Index As Integer = 1 To SortedIndexes.Count - 1

                Dim CurrentIndex As Integer = SortedIndexes(Index)

                If CurrentIndex = PreviousIndex + 1 Then
                    BlockCount += 1
                Else
                    Blocks.Add(New StructuralDeleteBlock With {
                        .StartRecordIndex = BlockStart,
                        .RecordCount = BlockCount
                    })
                    BlockStart = CurrentIndex
                    BlockCount = 1
                End If

                PreviousIndex = CurrentIndex

            Next

            Blocks.Add(New StructuralDeleteBlock With {
                .StartRecordIndex = BlockStart,
                .RecordCount = BlockCount
            })

            Blocks.Sort(Function(A As StructuralDeleteBlock, B As StructuralDeleteBlock)
                            Return B.StartRecordIndex.CompareTo(A.StartRecordIndex)
                        End Function)

            Return Blocks

        End Function

        Private Function ExecuteDelete(ByVal Rule As WorkbookStructureRule,
                                       ByVal RecordIndexes As List(Of Integer)) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}

            If IsExecuting Then
                Result.BError = True
                Result.StringReturn = "A workbook structural operation is already in progress."
                Return Result
            End If

            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            Dim Anchor As DefinedName = WB.DefinedNames.GetDefinedName(Rule.DeleteAnchorNamedRange)

            If Anchor Is Nothing OrElse Anchor.Range Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Delete anchor named range '" & Rule.DeleteAnchorNamedRange & "' was not found."
                Return Result
            End If

            Dim FirstRecordIndex As Integer

            Select Case Rule.Axis
                Case WorkbookStructureAxis.Columns
                    FirstRecordIndex = Anchor.Range.LeftColumnIndex
                Case WorkbookStructureAxis.Rows
                    FirstRecordIndex = Anchor.Range.TopRowIndex
                Case Else
                    Result.BError = True
                    Result.StringReturn = "Unsupported structural axis."
                    Return Result
            End Select

            'As with insertion, retain the pre-delete logical range dimensions.
            Dim RecordRangeSnapshot As StructuralNamedRangeSnapshot =
                SnapshotNamedRange(WB, Rule.RecordCountNamedRange)
            Dim JournalInputSnapshot As StructuralNamedRangeSnapshot = Nothing

            Dim DeleteBlocks As List(Of StructuralDeleteBlock) =
                BuildDeleteBlocks(RecordIndexes)
            Dim ThreeDReferences As WorkbookStructural3DReferences = Nothing

            For Each Target As WorkbookStructureTarget In Rule.Targets
                If GetWorksheet(WB, Target.WorksheetName) Is Nothing Then
                    Result.BError = True
                    Result.StringReturn = "Worksheet '" & Target.WorksheetName & "' was not found."
                    Result.StrResponseMessage = Result.StringReturn
                    Return Result
                End If
            Next

            Dim ChangedWorksheets As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim UpdateStarted As Boolean = False
            Dim MutationStarted As Boolean = False
            Dim BulkMutationGuardStarted As Boolean = False
            Dim PreviousCalculationMode As WorkbookCalculationMode = WB.Options.CalculationMode
            Dim PreviousCalculationEngine As CalculationEngineType = WB.Options.CalculationEngineType

            Try

                JournalInputSnapshot = SnapshotJournalInputRange(WB, Rule, RecordRangeSnapshot)

                Dim AdditionalSnapshots = Rule.AdditionalRecordRanges.Select(Function(n) SnapshotNamedRange(WB, n)).ToList()
                If RecordRangeSnapshot Is Nothing OrElse AdditionalSnapshots.Any(Function(s) s Is Nothing) Then Throw New InvalidOperationException("A required input/display range is missing. No workbook ranges were changed.")
                Dim LinkedEnd = ValidateLinkedColumns(WB, Rule)
                Dim LinkedFirst = LinkedEnd - GetCurrentRecordCount(Rule, WB)
                If Rule.LinkedColumns IsNot Nothing Then
                    WorkbookStructural3DReferences.Capture(WB, Rule.LinkedColumns.Targets.Select(Function(t) t.WorksheetName),
                        DeleteBlocks.Select(Function(b) New KeyValuePair(Of Integer, Integer)(LinkedFirst + b.StartRecordIndex, -b.RecordCount)))
                End If

                If Rule.PreserveThreeDReferences Then
                    ThreeDReferences = WorkbookStructural3DReferences.Capture(WB,
                        Rule.Targets.Select(Function(t) t.WorksheetName),
                        DeleteBlocks.Select(Function(b) New KeyValuePair(Of Integer, Integer)(FirstRecordIndex + b.StartRecordIndex, -b.RecordCount)))
                End If

                IsExecuting = True
                ModelSafetyManager.BeginBulkWorkbookMutation(ModelID)
                BulkMutationGuardStarted = True

                WB.Options.CalculationMode = WorkbookCalculationMode.Manual
                If Rule.UseRecursiveEngineForMutation Then WB.Options.CalculationEngineType = CalculationEngineType.Recursive

                WB.BeginUpdate()
                UpdateStarted = True

                For Each Target As WorkbookStructureTarget In Rule.Targets

                    Dim WS As Worksheet = GetWorksheet(WB, Target.WorksheetName)

                    If WS Is Nothing Then
                        Throw New InvalidOperationException("Worksheet '" & Target.WorksheetName & "' was not found.")
                    End If

                    Dim WasProtected As Boolean = WS.IsProtected
                    Dim WasVisible As Boolean = WS.Visible

                    Try

                        WS.Visible = True
                        If WasProtected Then UNProtectWS(ModelID, WS.Name)

                        For Each DeleteBlock As StructuralDeleteBlock In DeleteBlocks

                            Select Case Rule.Axis
                                Case WorkbookStructureAxis.Columns
                                    MutationStarted = True
                                    WS.Columns.Remove(FirstRecordIndex + DeleteBlock.StartRecordIndex,
                                                      DeleteBlock.RecordCount)
                                Case WorkbookStructureAxis.Rows
                                    MutationStarted = True
                                    WS.Rows.Remove(FirstRecordIndex + DeleteBlock.StartRecordIndex,
                                                   DeleteBlock.RecordCount)
                            End Select

                        Next

                        ChangedWorksheets.Add(WS.Name)

                    Finally

                        If WasProtected Then ProtectWS(ModelID, WS.Name)
                        WS.Visible = WasVisible

                    End Try

                Next

                If ThreeDReferences IsNot Nothing Then ThreeDReferences.Apply(ModelID)
                ResizeNamedRangeFromSnapshot(WB,
                                             RecordRangeSnapshot,
                                             Rule.Axis,
                                             -RecordIndexes.Count)
                ResizeNamedRangeFromSnapshot(WB, JournalInputSnapshot, Rule.Axis, -RecordIndexes.Count)
                For Each Snapshot In AdditionalSnapshots
                    ResizeNamedRangeFromSnapshot(WB, Snapshot, Rule.Axis, -RecordIndexes.Count)
                Next
                If Rule.LinkedColumns IsNot Nothing Then
                    DeleteLinkedColumns(WB, Rule.LinkedColumns, LinkedFirst, DeleteBlocks, ChangedWorksheets)
                End If
                Result.BError = False
                Result.EventCancelled = False
                Result.StringReturn = RecordIndexes.Count.ToString & " " & Rule.Description & " deleted."

            Catch ex As Exception

                Result.BError = True
                Result.StringReturn = ex.Message
                Result.StrResponseMessage = ex.Message
                System.Diagnostics.Trace.WriteLine("[Structure failure] " & Rule.RuleID & ": " & ex.ToString())
                If MutationStarted Then
                    ModelSafetyManager.MarkRecoveryRequired(
                        ModelID, "Delete " & Rule.Description, ex.Message,
                        "Workbook Structure", Rule.RuleID)
                End If

            Finally

                Try
                    If UpdateStarted Then WB.EndUpdate()
                Catch ex As Exception
                    Result.BError = True
                    Result.StringReturn &= Environment.NewLine & "End workbook update: " & ex.Message
                    If MutationStarted Then
                        ModelSafetyManager.MarkRecoveryRequired(
                            ModelID, "Delete " & Rule.Description, Result.StringReturn,
                            "Workbook Structure", Rule.RuleID)
                    End If
                End Try
                Try
                    If Rule.UseRecursiveEngineForMutation Then WB.Options.CalculationEngineType = PreviousCalculationEngine
                Catch ex As Exception
                    Result.BError = True
                    Result.StringReturn &= Environment.NewLine & "Restore calculation engine: " & ex.Message
                    If MutationStarted Then ModelSafetyManager.MarkRecoveryRequired(ModelID, "Restore calculation engine", Result.StringReturn, "Workbook Structure", Rule.RuleID)
                End Try
                Try
                    WB.Options.CalculationMode = PreviousCalculationMode
                Catch ex As Exception
                    Result.BError = True
                    Result.StringReturn &= Environment.NewLine & "Restore calculation mode: " & ex.Message
                End Try
                IsExecuting = False
                If BulkMutationGuardStarted Then
                    ModelSafetyManager.EndBulkWorkbookMutation(ModelID)
                    BulkMutationGuardStarted = False
                End If

            End Try

            If Not Result.BError Then

                Dim PostActionResult As AbovoTransaction =
                    RunPostActions(Rule, ChangedWorksheets)

                If PostActionResult.BError Then
                    Result.BError = True
                    Result.BSuccess = False
                    Result.StringReturn =
                        "Workbook structure changed, but Transactional DB synchronisation failed: " &
                        PostActionResult.StringReturn
                    Result.StrResponseMessage = Result.StringReturn
                    ModelSafetyManager.MarkRecoveryRequired(
                        ModelID, "Delete " & Rule.Description, Result.StringReturn,
                        "Workbook Structure", Rule.RuleID)
                End If

            End If

            If MutationStarted Then ExcelModels(ModelID).IsDirty = True
            Return Result

        End Function

        Private Function ValidateLinkedColumns(WB As IWorkbook, Rule As WorkbookStructureRule) As Integer
            If Rule.LinkedColumns Is Nothing Then Return 0
            If Rule.Axis <> WorkbookStructureAxis.Rows Then Throw New InvalidOperationException("Linked workings require an input-row rule.")
            Dim EndName = WB.DefinedNames.GetDefinedName(Rule.LinkedColumns.EndAnchorNamedRange)
            If EndName Is Nothing OrElse EndName.Range Is Nothing Then Throw New InvalidOperationException("Linked workings boundary is missing: " & Rule.LinkedColumns.EndAnchorNamedRange)
            Dim EndIndex = EndName.Range.LeftColumnIndex
            If EndIndex - GetCurrentRecordCount(Rule, WB) < 0 Then Throw New InvalidOperationException("Linked workings boundaries are inconsistent.")
            For Each Target In Rule.LinkedColumns.Targets
                If GetWorksheet(WB, Target.WorksheetName) Is Nothing Then Throw New InvalidOperationException("Linked worksheet is missing: " & Target.WorksheetName)
                If Rule.Targets.Any(Function(t) String.Equals(t.WorksheetName, Target.WorksheetName, StringComparison.OrdinalIgnoreCase)) Then Throw New InvalidOperationException("Input and linked workings cannot share a worksheet.")
            Next
            If Rule.RuleID = "SIMPLE_REP_SERVCHG_02" Then
                Dim Input = WB.DefinedNames.GetDefinedName("IR_ServChg_01").Range
                Dim Rep = WB.DefinedNames.GetDefinedName("Rep_ServChg_02").Range
                If Input.TopRowIndex <> Rep.TopRowIndex OrElse Rep.RowCount <> Input.RowCount + 1 Then Throw New InvalidOperationException("Service Charge input/display boundaries are inconsistent; reopen an intact workbook before inserting.")
            End If
            Return EndIndex
        End Function

        Private Sub InsertLinkedColumns(WB As IWorkbook, Linked As WorkbookLinkedColumns,
            InsertIndex As Integer, Count As Integer, Changed As HashSet(Of String))
            Dim Copies As New List(Of Action)
            Dim References = WorkbookStructural3DReferences.Capture(WB, Linked.Targets.Select(Function(t) t.WorksheetName),
                {New KeyValuePair(Of Integer, Integer)(InsertIndex, Count)})
            For Each Target In Linked.Targets
                Dim WS = GetWorksheet(WB, Target.WorksheetName)
                Dim WasProtected = WS.IsProtected, WasVisible = WS.Visible
                Try
                    WS.Visible = True
                    If WasProtected Then UNProtectWS(ModelID, WS.Name)
                    InsertColumnsForTarget(WS, InsertIndex, Count, Target, Copies)
                    Changed.Add(WS.Name)
                Finally
                    If WasProtected Then ProtectWS(ModelID, WS.Name)
                    WS.Visible = WasVisible
                End Try
            Next
            References.Apply(ModelID)
            For Each Copy In Copies
                Copy()
            Next
        End Sub

        Private Sub DeleteLinkedColumns(WB As IWorkbook, Linked As WorkbookLinkedColumns,
            FirstIndex As Integer, Blocks As List(Of StructuralDeleteBlock), Changed As HashSet(Of String))
            Dim References = WorkbookStructural3DReferences.Capture(WB, Linked.Targets.Select(Function(t) t.WorksheetName),
                Blocks.Select(Function(b) New KeyValuePair(Of Integer, Integer)(FirstIndex + b.StartRecordIndex, -b.RecordCount)))
            For Each Target In Linked.Targets
                Dim WS = GetWorksheet(WB, Target.WorksheetName)
                Dim WasProtected = WS.IsProtected, WasVisible = WS.Visible
                Try
                    WS.Visible = True
                    If WasProtected Then UNProtectWS(ModelID, WS.Name)
                    For Each Block In Blocks
                        WS.Columns.Remove(FirstIndex + Block.StartRecordIndex, Block.RecordCount)
                    Next
                    Changed.Add(WS.Name)
                Finally
                    If WasProtected Then ProtectWS(ModelID, WS.Name)
                    WS.Visible = WasVisible
                End Try
            Next
            References.Apply(ModelID)
        End Sub

        Private Sub InsertColumnsForTarget(ByVal WS As Worksheet,
                                           ByVal InsertIndex As Integer,
                                           ByVal RecordCount As Integer,
                                           ByVal Target As WorkbookStructureTarget,
                                           Optional DeferredCopies As List(Of Action) = Nothing,
                                           Optional Timing As StructuralInsertBenchmark = Nothing,
                                           Optional BatchNumber As Integer = 1)

            Dim TimingContext As String = "batch=" & BatchNumber.ToString() & ", sheet=" & WS.Name &
                ", column=" & (InsertIndex + 1).ToString() & ", count=" & RecordCount.ToString()

            Dim TemplateColumnIndexBeforeInsert As Integer =
                InsertIndex + Target.TemplateOffset

            If TemplateColumnIndexBeforeInsert < 0 Then
                Throw New InvalidOperationException("The template column for worksheet '" & WS.Name & "' is invalid.")
            End If

            Dim UsedRange As CellRange = WS.GetUsedRange()
            Dim UsedTop As Integer = UsedRange.TopRowIndex
            Dim UsedBottom As Integer = UsedRange.BottomRowIndex

            Dim WidthTemplateOffset As Integer = Target.TemplateOffset

            If Target.ColumnWidthTemplateOffset <> Integer.MinValue Then
                WidthTemplateOffset = Target.ColumnWidthTemplateOffset
            End If

            Dim WidthTemplateColumnIndex As Integer =
                InsertIndex + WidthTemplateOffset

            If WidthTemplateColumnIndex < 0 Then
                Throw New InvalidOperationException(
                    "The column-width template for worksheet '" &
                    WS.Name & "' is invalid.")
            End If

            Dim TemplateWidth As Single =
                WS.Columns(WidthTemplateColumnIndex).Width

            Using Stage = Timing?.Measure("shiftColumns", TimingContext)
                WS.Columns.Insert(InsertIndex, RecordCount)
            End Using

            'A template at or to the right of the insertion point moves with the
            'worksheet insertion. Resolve its post-insert coordinate before copying.
            Dim TemplateColumnIndex As Integer = TemplateColumnIndexBeforeInsert
            If TemplateColumnIndex >= InsertIndex Then
                TemplateColumnIndex += RecordCount
            End If

            If DeferredCopies IsNot Nothing Then
                DeferredCopies.Add(
                    Sub()
                        Dim WasProtected As Boolean = WS.IsProtected
                        Dim WasVisible As Boolean = WS.Visible
                        Try
                            WS.Visible = True
                            If WasProtected Then
                                Using Stage = Timing?.Measure("unprotect", "phase=copy, " & TimingContext)
                                    UNProtectWS(ModelID, WS.Name)
                                End Using
                            End If
                            Using Stage = Timing?.Measure("copyTemplate", TimingContext)
                                CopyInsertedColumns(WS, InsertIndex, RecordCount, Target,
                                                    TemplateColumnIndex, UsedTop, UsedBottom, TemplateWidth)
                            End Using
                            Using Stage = Timing?.Measure("copy3D", TimingContext)
                                WorkbookStructural3DReferences.CopyColumn(WS, TemplateColumnIndex,
                                                                          InsertIndex, RecordCount + Target.AdditionalCopiedColumns, UsedTop, UsedBottom)
                            End Using
                        Finally
                            If WasProtected Then
                                Using Stage = Timing?.Measure("protect", "phase=copy, " & TimingContext)
                                    ProtectWS(ModelID, WS.Name)
                                End Using
                            End If
                            WS.Visible = WasVisible
                        End Try
                    End Sub)
            Else
                Using Stage = Timing?.Measure("copyTemplate", TimingContext)
                    CopyInsertedColumns(WS, InsertIndex, RecordCount, Target,
                                        TemplateColumnIndex, UsedTop, UsedBottom, TemplateWidth)
                End Using
            End If

        End Sub

        Private Sub CopyInsertedColumns(WS As Worksheet, InsertIndex As Integer,
                                        RecordCount As Integer, Target As WorkbookStructureTarget,
                                        TemplateColumnIndex As Integer, UsedTop As Integer,
                                        UsedBottom As Integer, TemplateWidth As Single)

            Dim SourceTemplate As CellRange =
                WS.Range.FromLTRB(TemplateColumnIndex, UsedTop, TemplateColumnIndex, UsedBottom)

            Dim TargetRange As CellRange =
                WS.Range.FromLTRB(InsertIndex, UsedTop,
                    InsertIndex + RecordCount + If(TemplateColumnIndex = InsertIndex + RecordCount, 0, Target.AdditionalCopiedColumns) - 1, UsedBottom)

            Select Case Target.CopyMode

                Case WorkbookStructureCopyMode.FormatsAndColumnWidth
                    TargetRange.CopyFrom(SourceTemplate, PasteSpecial.Formats)

                    For ColIndex As Integer = InsertIndex To InsertIndex + RecordCount - 1
                        WS.Columns(ColIndex).Width = TemplateWidth
                    Next

                Case WorkbookStructureCopyMode.All
                    TargetRange.CopyFrom(SourceTemplate, PasteSpecial.All)

                Case WorkbookStructureCopyMode.AllAndColumnWidth
                    TargetRange.CopyFrom(SourceTemplate, PasteSpecial.All)

                    For ColIndex As Integer = InsertIndex To InsertIndex + RecordCount - 1
                        WS.Columns(ColIndex).Width = TemplateWidth
                    Next

            End Select

            If Target.ClearUnlockedConstants Then
                ClearInsertedInputs(WS.Range.FromLTRB(InsertIndex + Target.InputClearColumnOffset, UsedTop,
                    InsertIndex + Target.InputClearColumnOffset + RecordCount - 1, UsedBottom), Target)
            End If

        End Sub

        Private Sub InsertRowsForTarget(ByVal WS As Worksheet,
                                        ByVal InsertIndex As Integer,
                                        ByVal RecordCount As Integer,
                                        ByVal Target As WorkbookStructureTarget)

            Dim TemplateRowIndexBeforeInsert As Integer =
                InsertIndex + Target.TemplateOffset

            If TemplateRowIndexBeforeInsert < 0 Then
                Throw New InvalidOperationException("The template row for worksheet '" & WS.Name & "' is invalid.")
            End If

            Dim UsedRange As CellRange = WS.GetUsedRange()
            Dim UsedLeft As Integer = UsedRange.LeftColumnIndex
            Dim UsedRight As Integer = UsedRange.RightColumnIndex
            Dim TemplateHeight As Single = WS.Rows(TemplateRowIndexBeforeInsert).Height

            WS.Rows.Insert(InsertIndex, RecordCount)

            Dim TemplateRowIndex As Integer = TemplateRowIndexBeforeInsert
            If TemplateRowIndex >= InsertIndex Then
                TemplateRowIndex += RecordCount
            End If

            Dim SourceTemplate As CellRange =
                WS.Range.FromLTRB(UsedLeft, TemplateRowIndex, UsedRight, TemplateRowIndex)

            Dim TargetRange As CellRange =
                WS.Range.FromLTRB(UsedLeft, InsertIndex, UsedRight, InsertIndex + RecordCount - 1)

            Select Case Target.CopyMode

                Case WorkbookStructureCopyMode.FormatsAndColumnWidth
                    TargetRange.CopyFrom(SourceTemplate, PasteSpecial.Formats)

                Case Else
                    TargetRange.CopyFrom(SourceTemplate, PasteSpecial.All)

            End Select

            For RowIndex As Integer = InsertIndex To InsertIndex + RecordCount - 1
                WS.Rows(RowIndex).Height = TemplateHeight
            Next

            ClearInsertedInputs(TargetRange, Target)

        End Sub

        Private Shared Sub ClearInsertedInputs(Range As CellRange, Target As WorkbookStructureTarget)
            If Not Target.ClearUnlockedConstants Then Return
            'Match the existing native row-insertion contract: a new record must
            'not duplicate a populated input. Formula-backed unlocked cells remain
            'formulas; removing those would break Excel/Summit round-tripping.
            For Each Cell As Cell In Range.ExistingCells.ToList()
                If Not Cell.Protection.Locked AndAlso Not Cell.HasFormula Then Cell.Value = CellValue.Empty
            Next
        End Sub

        Private Function RunPostActions(ByVal Rule As WorkbookStructureRule,
                                        ByVal ChangedWorksheets As IEnumerable(Of String),
                                        Optional ByVal PreSyncedResult As AbovoTransaction = Nothing,
                                        Optional ByVal PreSyncMs As Long = 0) As AbovoTransaction

            Dim benchmark As System.Diagnostics.Stopwatch =
                System.Diagnostics.Stopwatch.StartNew()
            Dim Result As AbovoTransaction =
                If(PreSyncedResult, SynchroniseTransactionDB(Rule))
            Dim syncMs As Long = PreSyncMs + benchmark.ElapsedMilliseconds
            Dim dependencyStartMs As Long = benchmark.ElapsedMilliseconds
            'Invalidate every interface section dependent on any linked worksheet.
            'The entry calculation engine/mode have been restored before this point.
            If ExcelModels IsNot Nothing AndAlso
               ModelID >= 0 AndAlso ModelID < ExcelModels.Length AndAlso
               ExcelModels(ModelID) IsNot Nothing AndAlso
               ExcelModels(ModelID).InterfaceDependencies IsNot Nothing Then

                For Each WorksheetName As String In ChangedWorksheets
                    ExcelModels(ModelID).InterfaceDependencies.WorksheetStructureChanged(WorksheetName)
                Next

            End If

            System.Diagnostics.Trace.WriteLine(
                "[Population Benchmark] Structure post-actions: model=" &
                ModelID.ToString() &
                ", rule=" & Rule.RuleID &
                ", tdbSync=" & syncMs.ToString() & " ms" &
                ", engineShared=" & (PreSyncedResult IsNot Nothing).ToString() &
                ", dependencies=" &
                (benchmark.ElapsedMilliseconds - dependencyStartMs).ToString() & " ms" &
                ", total=" & (PreSyncMs + benchmark.ElapsedMilliseconds).ToString() & " ms" &
                ", outcome=" & If(Result.BError, "failed", "ok"))
            Return Result

        End Function

        Private Function SynchroniseTransactionDB(ByVal Rule As WorkbookStructureRule) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            'Do TransactionDB once after all linked workbook sheets are structurally
            'consistent.  This avoids synchronising an intermediate half-updated state.
            If Not String.IsNullOrWhiteSpace(Rule.TransactionDBSyncNamedRange) Then

                If ExcelModels IsNot Nothing AndAlso
                   ModelID >= 0 AndAlso ModelID < ExcelModels.Length AndAlso
                   ExcelModels(ModelID) IsNot Nothing AndAlso
                   ExcelModels(ModelID).TransDBSync IsNot Nothing Then

                    Dim SyncResult As AbovoTransaction =
                        ExcelModels(ModelID).TransDBSync.SynchroniseForNamedRange(Rule.TransactionDBSyncNamedRange)

                    If SyncResult.BError Then
                        Result.BError = True
                        Result.BSuccess = False
                        Result.StringReturn = SyncResult.StringReturn
                        Result.StrResponseMessage = SyncResult.StrResponseMessage
                    End If

                Else

                    Result.BError = True
                    Result.BSuccess = False
                    Result.StringReturn = "Transactional DB synchronisation service is unavailable."
                    Result.StrResponseMessage = Result.StringReturn

                End If

            End If

            Return Result

        End Function

        Private Function SnapshotJournalInputRange(WB As IWorkbook, rule As WorkbookStructureRule,
                                                  records As StructuralNamedRangeSnapshot) As StructuralNamedRangeSnapshot
            If rule.RuleID <> RuleJournalRecords Then Return Nothing
            Dim input = SnapshotNamedRange(WB, "IR_Journals")
            'IR excludes the sentinel. Insertion just below its last row does not
            'auto-expand the name. Resize from the pre-edit snapshot exactly once,
            'before synchronising BOTH debit and credit Transactional DB mirrors.
            If input Is Nothing OrElse records Is Nothing OrElse
               input.RangeWorksheetName <> records.RangeWorksheetName OrElse
               input.TopRowIndex <> records.TopRowIndex OrElse
               input.BottomRowIndex <> records.BottomRowIndex - 1 Then
                Throw New InvalidOperationException("Journal range geometry is inconsistent: IR_Journals must contain all Rep_Jour_01 records except its final sentinel row. No rows have been changed; review this workbook's journal names before continuing.")
            End If
            Return input
        End Function

        Private Function SnapshotNamedRange(ByVal WB As IWorkbook,
                                                   ByVal NamedRange As String) As StructuralNamedRangeSnapshot

            If WB Is Nothing OrElse String.IsNullOrWhiteSpace(NamedRange) Then Return Nothing

            Dim DN As DefinedName = Nothing
            Dim ScopeWorksheet As Worksheet = Nothing

            'Workbook/global scope first.
            Try
                DN = WB.DefinedNames.GetDefinedName(NamedRange)
            Catch
                DN = Nothing
            End Try

            'Then allow for a worksheet-local name.
            If DN Is Nothing Then

                For Each WS As Worksheet In WB.Worksheets

                    Try

                        DN = WS.DefinedNames.GetDefinedName(NamedRange)

                        If DN IsNot Nothing Then
                            ScopeWorksheet = WS
                            Exit For
                        End If

                    Catch
                        DN = Nothing
                    End Try

                Next

            End If

            If DN Is Nothing OrElse DN.Range Is Nothing OrElse DN.Range.Worksheet Is Nothing Then Return Nothing

            Dim Rng As CellRange = DN.Range

            Return New StructuralNamedRangeSnapshot With {
                .Name = NamedRange,
                .IsGlobal = DN.IsGlobal,
                .ScopeWorksheetName = If(ScopeWorksheet Is Nothing, Nothing, ScopeWorksheet.Name),
                .RangeWorksheetName = Rng.Worksheet.Name,
                .LeftColumnIndex = Rng.LeftColumnIndex,
                .TopRowIndex = Rng.TopRowIndex,
                .RightColumnIndex = Rng.RightColumnIndex,
                .BottomRowIndex = Rng.BottomRowIndex
            }

        End Function

        Private Sub ResizeNamedRangeFromSnapshot(ByVal WB As IWorkbook,
                                                 ByVal Snapshot As StructuralNamedRangeSnapshot,
                                                 ByVal Axis As WorkbookStructureAxis,
                                                 ByVal SizeDelta As Integer)

            If WB Is Nothing OrElse Snapshot Is Nothing OrElse SizeDelta = 0 Then Return

            Dim WS As Worksheet = GetWorksheet(WB, Snapshot.RangeWorksheetName)

            If WS Is Nothing Then
                Throw New InvalidOperationException(
                    "Worksheet '" & Snapshot.RangeWorksheetName &
                    "' for named range '" & Snapshot.Name & "' was not found.")
            End If

            Dim NewRight As Integer = Snapshot.RightColumnIndex
            Dim NewBottom As Integer = Snapshot.BottomRowIndex

            Select Case Axis

                Case WorkbookStructureAxis.Columns
                    NewRight += SizeDelta

                    If NewRight < Snapshot.LeftColumnIndex Then
                        Throw New InvalidOperationException(
                            "The resized named range '" & Snapshot.Name & "' would have no columns.")
                    End If

                Case WorkbookStructureAxis.Rows
                    NewBottom += SizeDelta

                    If NewBottom < Snapshot.TopRowIndex Then
                        Throw New InvalidOperationException(
                            "The resized named range '" & Snapshot.Name & "' would have no rows.")
                    End If

                Case Else
                    Throw New InvalidOperationException("Unsupported structural axis.")

            End Select

            Dim NewRange As CellRange =
                WS.Range.FromLTRB(Snapshot.LeftColumnIndex,
                                  Snapshot.TopRowIndex,
                                  NewRight,
                                  NewBottom)

            Dim DN As DefinedName = Nothing

            If Snapshot.IsGlobal Then

                DN = WB.DefinedNames.GetDefinedName(Snapshot.Name)

            Else

                Dim ScopeWS As Worksheet = GetWorksheet(WB, Snapshot.ScopeWorksheetName)

                If ScopeWS IsNot Nothing Then
                    DN = ScopeWS.DefinedNames.GetDefinedName(Snapshot.Name)
                End If

            End If

            If DN Is Nothing Then
                Throw New InvalidOperationException(
                    "Named range '" & Snapshot.Name & "' could not be restored after the structural edit.")
            End If

            DN.Range = NewRange

#If DEBUG Then
#End If

        End Sub

        Private Function GetDeletionRecordCount(Rule As WorkbookStructureRule) As Integer
            If Rule.RuleID <> RuleFundingRecords Then Return GetCurrentRecordCount(Rule)
            'FacilityNames includes the protected revolvers. VBA deletes only
            'ordinary loans, never the first ten or any revolver/template column.
            Try
                Dim WB = GetWorkbook()
                Dim Ordinary = WB.DefinedNames.GetDefinedName("LoanDescsOrd").Range
                Dim Facilities = WB.DefinedNames.GetDefinedName("FacilityNames").Range
                Dim Revolver = WB.DefinedNames.GetDefinedName("LoanDescRev1").Range
                If Ordinary.Worksheet IsNot Facilities.Worksheet OrElse
                   Ordinary.Worksheet IsNot Revolver.Worksheet OrElse
                   Ordinary.LeftColumnIndex <> Facilities.LeftColumnIndex OrElse
                   Ordinary.RightColumnIndex <> Revolver.LeftColumnIndex - 1 Then Return -1
                Return Ordinary.ColumnCount
            Catch
                Return -1
            End Try
        End Function

        Private Function GetCurrentRecordCount(ByVal Rule As WorkbookStructureRule) As Integer

            Return GetCurrentRecordCount(Rule, GetWorkbook())

        End Function

        Private Function GetCurrentRecordCount(ByVal Rule As WorkbookStructureRule,
                                               ByVal WB As IWorkbook) As Integer

            If Rule Is Nothing Then Return -1
            If String.IsNullOrWhiteSpace(Rule.RecordCountNamedRange) Then Return -1
            If WB Is Nothing Then Return -1

            Try

                Dim DN As DefinedName = WB.DefinedNames.GetDefinedName(Rule.RecordCountNamedRange)

                If DN Is Nothing OrElse DN.Range Is Nothing Then Return -1

                Dim Count As Integer

                Select Case Rule.Axis
                    Case WorkbookStructureAxis.Columns
                        Count = DN.Range.ColumnCount
                    Case WorkbookStructureAxis.Rows
                        Count = DN.Range.RowCount
                    Case Else
                        Return -1
                End Select

                Dim PhysicalCount As Integer = Count

                Count += Rule.RecordCountAdjustment
                Count = Math.Max(0, Count)

#If DEBUG Then
#End If

                Return Count

            Catch
                Return -1
            End Try

        End Function

        Private Function GetWorkbook() As IWorkbook

            If ExcelModels Is Nothing Then Return Nothing
            If ModelID < 0 OrElse ModelID >= ExcelModels.Length Then Return Nothing
            If ExcelModels(ModelID) Is Nothing Then Return Nothing

            Return ExcelModels(ModelID).WB

        End Function

        Private Function GetWorksheet(ByVal WB As IWorkbook,
                                      ByVal WorksheetName As String) As Worksheet

            If WB Is Nothing OrElse String.IsNullOrWhiteSpace(WorksheetName) Then Return Nothing

            Try
                Return WB.Worksheets(WorksheetName)
            Catch
                Return Nothing
            End Try

        End Function

    End Class

    Friend Class StructuralDeleteBlock

        Public StartRecordIndex As Integer
        Public RecordCount As Integer

    End Class

    Friend Class StructuralNamedRangeSnapshot

        Public Name As String
        Public IsGlobal As Boolean
        Public ScopeWorksheetName As String
        Public RangeWorksheetName As String
        Public LeftColumnIndex As Integer
        Public TopRowIndex As Integer
        Public RightColumnIndex As Integer
        Public BottomRowIndex As Integer

    End Class

    Public Class WorkbookStructureRule

        Public RuleID As String
        Public Description As String
        Public Axis As WorkbookStructureAxis
        Public InsertAnchorNamedRange As String

        'Optional offset from the named insertion anchor. A value of -1 is used
        'when the anchor identifies a sentinel immediately after the real template.
        Public InsertIndexOffset As Integer = 0

        'Opt-in for reviewed grouped worksheet insertion contracts.
        Public InsertAllColumnsBeforeCopy As Boolean = False
        Public InsertFirstColumnSeparately As Boolean = False
        Public PreserveThreeDReferences As Boolean = False
        Public ProtectedLeadingRecordCount As Integer = 0

        'Normally InsertAnchorNamedRange identifies the exact insertion column/row
        'and its left/top edge is used. Some workbook structures instead expose
        'the whole logical range as the only stable anchor. In that case insert
        'at the range's right/bottom edge.
        Public InsertAtAnchorEnd As Boolean = False

        'Some row/column families contain only genuine records and therefore
        'append immediately AFTER the current anchor range rather than inserting
        'before a trailing template/sentinel record.
        Public InsertAfterAnchorEnd As Boolean = False

        Public DeleteAnchorNamedRange As String
        Public RecordCountNamedRange As String
        Public RecordCountAdjustment As Integer = 0
        Public MinimumRecordCount As Integer = 0
        Public TransactionDBSyncNamedRange As String
        Public Targets As New List(Of WorkbookStructureTarget)
        Public AdditionalRecordRanges As New List(Of String)
        Public LinkedColumns As WorkbookLinkedColumns
        Public FillRelativeHeaderRowOffset As Integer = -1
        Public UseRecursiveEngineForMutation As Boolean = False

    End Class

    Public Class WorkbookStructureTarget

        Public WorksheetName As String
        Public CopyMode As WorkbookStructureCopyMode = WorkbookStructureCopyMode.All

        'Source column/row used for formulas, values and formats.
        Public TemplateOffset As Integer = -1

        'Optional separate source for the physical column width.
        'Integer.MinValue means "use TemplateOffset", preserving existing rules.
        Public ColumnWidthTemplateOffset As Integer = Integer.MinValue

        Public ClearUnlockedConstants As Boolean = False
        Public InputClearColumnOffset As Integer = 0
        Public AdditionalCopiedColumns As Integer = 0

    End Class

    Public Class WorkbookLinkedColumns
        Public EndAnchorNamedRange As String
        Public Targets As New List(Of WorkbookStructureTarget)
    End Class

    Public Enum WorkbookStructureAxis
        Rows = 0
        Columns = 1
    End Enum

    Public Enum WorkbookStructureCopyMode
        All = 0
        FormatsAndColumnWidth = 1
        AllAndColumnWidth = 2
    End Enum

End Namespace
