Imports System.Data.SqlClient
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.GeneralFunctions
Imports Abovo.PresentationManager
Imports Abovo.WSSecurity
Imports DevExpress.CodeParser
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports DevExpress.XtraLayout.Customization
Imports DevExpress.XtraRichEdit.Model
Imports Microsoft.VisualBasic.Devices

Namespace Abovo
    Public Class SpecificRowColumnEvents

        Private Shared BusPlanFile As DevExpress.Spreadsheet.IWorkbook

        Public Shared Sub InsertOFAColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            InsertRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleOFARecords,
                                "Other Fixed Asset",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub DeleteOFAColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            DeleteRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleOFARecords,
                                "Other Fixed Asset",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub
        Public Shared Sub InsertHAColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            InsertRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleHousingComponentRecords,
                                "Housing Asset component",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub DeleteHAColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            DeleteRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleHousingComponentRecords,
                                "Housing Asset component",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub
        Public Shared Sub InsertCapExColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            InsertRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleCapExRecords,
                                "Capital Expenditure",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub DeleteCapExColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            DeleteRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleCapExRecords,
                                "Capital Expenditure",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub InsertCapGrantColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            InsertRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleCapGrantRecords,
                                "Capital Grant",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub DeleteCapGrantColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            DeleteRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleCapGrantRecords,
                                "Capital Grant",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub InsertRepairsColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            InsertRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleRepairsRecords,
                                "Repairs and Maintenance",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub DeleteRepairsColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            DeleteRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleRepairsRecords,
                                "Repairs and Maintenance",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub InsertFundingColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            InsertRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleFundingRecords,
                                "Funding facility",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub DeleteFundingColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            DeleteRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleFundingRecords,
                                "Funding facility",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub InsertDevelopmentIdentifiedColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            InsertRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleDevelopmentIdentifiedRecords,
                                "Identified Development",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub DeleteDevelopmentIdentifiedColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            DeleteRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleDevelopmentIdentifiedRecords,
                                "Identified Development",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub InsertDevelopmentMultiYearColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            InsertRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleDevelopmentMultiYearRecords,
                                "Multi-year Development",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub DeleteDevelopmentMultiYearColumns(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            DeleteRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleDevelopmentMultiYearRecords,
                                "Multi-year Development",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub InsertJournalRows(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            InsertRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleJournalRecords,
                                "Journal",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub DeleteJournalRows(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            DeleteRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleJournalRecords,
                                "Journal",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub InsertStockConversionRows(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            InsertRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleStockConversionRecords,
                                "Stock Conversion",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Public Shared Sub DeleteStockConversionRows(ModelID As Integer, ByRef GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

            DeleteRecordsByRule(ModelID,
                                WorkbookStructureRuleManager.RuleStockConversionRecords,
                                "Stock Conversion",
                                GridCommandTag,
                                SetTransaction,
                                ActioningForm)

        End Sub

        Private Shared Sub InsertRecordsByRule(ByVal ModelID As Integer,
                                               ByVal RuleID As String,
                                               ByVal RecordDescription As String,
                                               ByRef GridCommandTag As AttachedGridCommandButton,
                                               ByVal SetTransaction As AbovoTransaction,
                                               ByVal ActioningForm As Form)

            Dim InsertRecordCount As Integer = 0

            If GridCommandTag IsNot Nothing Then
                InsertRecordCount = GridCommandTag.RequestedRecordCount
            End If

            If InsertRecordCount <= 0 Then
                InsertRecordCount =
                    GetNumericalIntegerInput("How many records do you wish to add?",
                                             "Insert " & RecordDescription & " Records",
                                             3, 0, 100)
            End If

            If InsertRecordCount <= 0 Then
                SetTransaction.BError = True
                SetTransaction.EventCancelled = True
                Return
            End If

            Try

                ActioningForm.Cursor = Cursors.WaitCursor

                Dim Result As AbovoTransaction =
                    ExcelModels(ModelID).WorkbookStructureRules.AddRecords(RuleID, InsertRecordCount)

                CopyTransactionResult(Result, SetTransaction)

                If Result.BError Then
                    XtraMessageBox.Show("An error occurred while trying to add the " &
                                        RecordDescription & " records." &
                                        Environment.NewLine & Result.StringReturn)
                End If

            Finally

                ActioningForm.Cursor = Cursors.Default

            End Try

        End Sub

        Private Shared Sub DeleteRecordsByRule(ByVal ModelID As Integer,
                                               ByVal RuleID As String,
                                               ByVal RecordDescription As String,
                                               ByRef GridCommandTag As AttachedGridCommandButton,
                                               ByVal SetTransaction As AbovoTransaction,
                                               ByVal ActioningForm As Form)

            If GridCommandTag IsNot Nothing AndAlso
               GridCommandTag.DeleteLastRecords AndAlso
               GridCommandTag.RequestedRecordCount > 0 Then

                Try

                    ActioningForm.Cursor = Cursors.WaitCursor

                    Dim Result As AbovoTransaction =
                        ExcelModels(ModelID).WorkbookStructureRules.DeleteLastRecords(
                            RuleID,
                            GridCommandTag.RequestedRecordCount)

                    CopyTransactionResult(Result, SetTransaction)

                    If Result.BError AndAlso
                       Not String.IsNullOrWhiteSpace(Result.StringReturn) Then

                        XtraMessageBox.Show(Result.StringReturn)
                    End If

                Finally

                    ActioningForm.Cursor = Cursors.Default

                End Try

                Return
            End If

            If GridCommandTag Is Nothing OrElse GridCommandTag.AttachedGrid Is Nothing Then
                SetTransaction.BError = True
                SetTransaction.EventCancelled = True
                SetTransaction.StringReturn = "The structural action is not attached to a grid."
                Return
            End If

            Dim GC As GridControl = GridCommandTag.AttachedGrid
            Dim GV As DevExpress.XtraGrid.Views.Grid.GridView =
                TryCast(GC.FocusedView, DevExpress.XtraGrid.Views.Grid.GridView)

            If GV Is Nothing Then
                SetTransaction.BError = True
                SetTransaction.EventCancelled = True
                SetTransaction.StringReturn = "The structural action requires a GridView."
                Return
            End If

            Dim SelectedRecordIndexes As New List(Of Integer)
            Dim SelectedRowsString As String = ""

            For Each RowHandle As Integer In GV.GetSelectedRows()

                If RowHandle >= 0 Then

                    Dim DataSourceIndex As Integer = GV.GetDataSourceRowIndex(RowHandle)

                    If DataSourceIndex >= 0 AndAlso Not SelectedRecordIndexes.Contains(DataSourceIndex) Then
                        SelectedRecordIndexes.Add(DataSourceIndex)
                        SelectedRowsString &= (DataSourceIndex + 1).ToString & ", "
                    End If

                End If

            Next

            If SelectedRecordIndexes.Count = 0 Then
                XtraMessageBox.Show("Please select at least one record to delete.")
                SetTransaction.BError = True
                SetTransaction.EventCancelled = True
                Return
            End If

            SelectedRowsString = SelectedRowsString.TrimEnd(", ".ToCharArray())

            Dim MsgText As String =
                "You have selected " & SelectedRecordIndexes.Count.ToString &
                " records for deletion. Record numbers: " & SelectedRowsString &
                ". Do you want to delete these records?"

            Dim Args As New XtraMessageBoxArgs With {
                .Caption = "Confirm deletion",
                .Text = MsgText,
                .Buttons = New DialogResult() {DialogResult.Yes, DialogResult.No}
            }

            If XtraMessageBox.Show(Args) = DialogResult.No Then
                SetTransaction.BError = True
                SetTransaction.EventCancelled = True
                Return
            End If

            Try

                ActioningForm.Cursor = Cursors.WaitCursor

                Dim Result As AbovoTransaction =
                    ExcelModels(ModelID).WorkbookStructureRules.DeleteRecords(RuleID, SelectedRecordIndexes)

                CopyTransactionResult(Result, SetTransaction)

                If Result.BError AndAlso Not Result.EventCancelled Then
                    XtraMessageBox.Show("An error occurred while trying to delete the selected " &
                                        RecordDescription & " records." &
                                        Environment.NewLine & Result.StringReturn)
                ElseIf Result.BError AndAlso Result.EventCancelled AndAlso
                       Not String.IsNullOrWhiteSpace(Result.StringReturn) Then
                    XtraMessageBox.Show(Result.StringReturn)
                End If

            Finally

                ActioningForm.Cursor = Cursors.Default

            End Try

        End Sub

        Private Shared Sub CopyTransactionResult(ByVal Source As AbovoTransaction,
                                                 ByVal Target As AbovoTransaction)

            If Source Is Nothing OrElse Target Is Nothing Then Return

            Target.BError = Source.BError
            Target.EventCancelled = Source.EventCancelled
            Target.StringReturn = Source.StringReturn

        End Sub

        'OFA workbook manipulation is now handled by WorkbookStructureRuleManager.

        Class HAColumnsInsertion

            Public Sub New(ModelID As Integer, GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)


                Dim InsertRecordCount As Integer
                Dim Title As String

                Title = "Insert Other Fixed Asset Columns"

                '   Assign number of columns to be inserted
                InsertRecordCount = GetNumericalIntegerInput("How many records do you wish to add?", "Insert Housing Asset Records", 3, 0, 20)

                If InsertRecordCount > 0 Then

                    ActioningForm.Cursor = Cursors.WaitCursor
                    Insert_HAA_Columns(ModelID, InsertRecordCount, SetTransaction, ActioningForm)
                    'ExcelModels(ModelID).TransDBSync.SynchroniseHAA()

                Else

                    ActioningForm.Cursor = Cursors.Default
                    SetTransaction.BError = True
                    SetTransaction.EventCancelled = True

                End If

            End Sub

            Private Sub Insert_HAA_Columns(ModelID As Integer, ByVal SetInsertRecordCount As Integer, SetTransaction As AbovoTransaction, ActioningForm As Form)

                On Error GoTo ErrHandler

                Dim WB As DevExpress.Spreadsheet.IWorkbook = FileManager.GetWorkBook(ModelID)
                Dim ActiveSheets As New List(Of String) From {"Housing Asset Assumptions", "HAA Workings"}
                Dim ActiveWS As DevExpress.Spreadsheet.Worksheet
                Dim LastHAACol As DevExpress.Spreadsheet.Cell = WB.Range("LastHAACol")(0, 0)
                Dim RCOlIndex As Integer = LastHAACol.ColumnIndex

                WB.BeginUpdate()

                ActiveWS = WB.Worksheets("Other Fixed Asset Assumptions")

                ActiveWS.Visible = True

                UNProtectWS(ModelID, "Other Fixed Asset Assumptions")

                ActiveWS.Columns.Insert(RCOlIndex, SetInsertRecordCount)

                For i = 0 To SetInsertRecordCount - 1

                    ActiveWS.Columns(RCOlIndex + i).CopyFrom(ActiveWS.Columns(RCOlIndex - 1), PasteSpecial.Formats)
                    ActiveWS.Columns(RCOlIndex + i).CopyFrom(ActiveWS.Columns(RCOlIndex - 1), PasteSpecial.ColumnWidths)

                Next i

                ActiveWS = WB.Worksheets("OFA Workings")

                ActiveWS.Visible = True

                UNProtectWS(ModelID, "OFA Workings")

                ActiveWS.Columns.Insert(RCOlIndex, SetInsertRecordCount)

                For i = 0 To SetInsertRecordCount - 1

                    ActiveWS.Columns(RCOlIndex + i).CopyFrom(ActiveWS.Columns(RCOlIndex - 1), PasteSpecial.All)

                Next i

                For Each WSName In ActiveSheets

                    If Left(WSName, 6) = "Hidden" Then WB.Worksheets(WSName).Visible = False
                    ProtectWS(ModelID, WSName)

                Next WSName

Exiter:
                On Error Resume Next

                WB.EndUpdate()

                ActioningForm.Cursor = Cursors.Default
                Exit Sub

ErrHandler:

                ActioningForm.Cursor = Cursors.Default
                XtraMessageBox.Show("An error occurred while trying to delete the selected records. Please ensure you have selected valid records and try again.")
                SetTransaction.BError = True
                SetTransaction.EventCancelled = True

            End Sub

        End Class

        Class HAColumnsDeletion

            Public Sub New(ModelID As Integer, GridCommandTag As AttachedGridCommandButton, SetTransaction As AbovoTransaction, ActioningForm As Form)

                On Error GoTo ErrHandler

                Dim GC As GridControl = GridCommandTag.AttachedGrid
                Dim GV As DevExpress.XtraGrid.Views.Grid.GridView = GC.FocusedView
                Dim SelectRows As Integer() = GV.GetSelectedRows()
                Dim selectedDataRowsCount As Integer = 0

                Dim SelectedRowsString As String = ""

                For Each rowHandle As Integer In GV.GetSelectedRows()

                    If rowHandle >= 0 Then

                        SelectedRowsString += (rowHandle + 1).ToString() & ", "
                        selectedDataRowsCount += 1

                    End If

                Next rowHandle

                If selectedDataRowsCount = 0 Then

                    XtraMessageBox.Show("Please select at least one record to delete.")
                    SetTransaction.BError = True
                    SetTransaction.EventCancelled = True
                    Return

                End If

                SelectedRowsString = If(SelectedRowsString = "", "None", SelectedRowsString.TrimEnd(", ".ToCharArray()))
                Dim MsgText As String = $"You have selected {selectedDataRowsCount} rows for deletion. Row numbers: {SelectedRowsString}. Do you want to delete these records?"
                Dim args As XtraMessageBoxArgs = New XtraMessageBoxArgs() With {
                                   .Caption = "Confirm deletion",
                                   .Text = MsgText,
                                   .Buttons = New DialogResult() {DialogResult.Yes, DialogResult.No}}

                If XtraMessageBox.Show(args) = DialogResult.No Then

                    SetTransaction.BError = True
                    SetTransaction.EventCancelled = True
                    Return

                End If

                ActioningForm.Cursor = Cursors.WaitCursor

                Dim WB As DevExpress.Spreadsheet.IWorkbook = FileManager.GetWorkBook(ModelID)
                Dim ActiveSheets As New List(Of String) From {"Other Fixed Asset Assumptions", "OFA Workings"}

                Dim ActiveWS As DevExpress.Spreadsheet.Worksheet

                Dim FirstOFACol As DevExpress.Spreadsheet.Cell = WB.Range("Rep_OFA_030")(0, 0)
                Dim RCOlIndex As Integer = FirstOFACol.ColumnIndex

                WB.BeginUpdate()

                For Each WSName In ActiveSheets

                    ActiveWS = WB.Worksheets(WSName)

                    ActiveWS.Visible = True

                    UNProtectWS(ModelID, WSName)

                    For i = SelectRows.Count - 1 To 0 Step -1

                        ActiveWS.Columns(RCOlIndex + SelectRows(i)).Delete()

                    Next i

                Next WSName

                For Each WSName In ActiveSheets

                    If Left(WSName, 6) = "Hidden" Then WB.Worksheets(WSName).Visible = False
                    ProtectWS(ModelID, WSName)

                Next WSName

                SetTransaction.BError = False
                SetTransaction.EventCancelled = False

Exiter:

                On Error Resume Next

                WB.EndUpdate()

                ActioningForm.Cursor = Cursors.Default


                Exit Sub

ErrHandler:

                ActioningForm.Cursor = Cursors.Default
                XtraMessageBox.Show("An error occurred while trying to delete the selected records. Please ensure you have selected valid records and try again.")
                SetTransaction.BError = True
                SetTransaction.EventCancelled = True

            End Sub


        End Class
        Public Shared Function ProcessSpecificRowColumnEvents(ModelID As Integer, RowColEventTag As RowColEventTag) As Boolean

            Return True

        End Function




        Public Shared Sub ProcessDvtColumns(ModelID As Integer, TargetRow As Integer, TargetCol As Integer, NewValue As Object, WSName As String)

            Dim wb As DevExpress.Spreadsheet.IWorkbook = FileManager.GetWorkBook(ModelID)
            Dim ws As DevExpress.Spreadsheet.Worksheet = wb.Worksheets(WSName)

            'Example: If a value is entered in column 2 (B), update column 3 (C) to be double that value

            If TargetCol = 2 Then ' Column B
                Dim newVal As Double
                If Double.TryParse(NewValue.ToString(), newVal) Then
                    ws.Cells(TargetRow, 3).Value = newVal * 2 ' Update Column C
                End If
            End If

        End Sub

        Public Class DvptColumns
            '            Const MinColsID = 10  ' Minimum dvpt categories for "Identified"
            '            Const MinColsMY = 3 ' Minimum dvpt categories for "Multi-year"
            '            Option Explicit On

            '            Function DvptSheets()

            '                DvptSheets = Array("Development BP Assumptions", "Development Stock",
            '            "Development Capital", "Development Revenue", "Development Expenditure",
            '            "Dvpt NonCash", "Dvpt Component Depn")

            '            End Function
            '            Function DevelopmentTransDBNRs()

            '                DvptSheets = Array("Development BP Assumptions", "Development Stock",
            '            "Development Capital", "Development Revenue", "Development Expenditure",
            '            "Dvpt NonCash", "Dvpt Component Depn")

            '            End Function
            '            Sub Insert_Dvpt_Columns()

            '                'CodeSafe JW 26/4/22
            '                ' Calling procedure for InsertDvptColumns
            '                Dim y As Integer
            '                Dim ResponseMessage, Title As String

            '                InitiateCodeRun "Insert_Dvpt_Columns", True

            '    Title = "Insert Identified Development Columns"

            '                '   Assign number of columns to be inserted
            '                y = Application.InputBox(Prompt:="How many columns do you wish to insert?", Title:=Title, Default:=1, Type:=1)                                     '   Integer type

            '                Call InsertDvptColumns(y, "Null", "BP")
            '                ResumeCodeRun "Insert_Dvpt_Columns"

            '    EndCodeRun True

            'End Sub
            '            Sub InsertDvptColumns(y As Integer, TargRng As String, CalledFrom As String)

            '                'IN TRACE 4/5/22 JW Note 84176
            '                ''FLUXHIGH 4/5/22 JW Note 84172/84176/84177
            '                'CodeSafe JW 26/4/22 'QUESTION 4/5/22

            '                ' Procedure to insert at the right any number of "Identified" scheme columns
            '                ' on the "Development BP Assumptions" sheet and its dependent sheets.
            '                ' Called by Insert_Dvpt_Columns

            '                Dim LastColNum As Integer, x As Integer
            '                Dim i As Integer, j As Integer, LastRow As Integer
            '                Dim CalcMethod As Integer
            '                Dim s

            '                InitiateCodeRun "InsertDvptColumns"
            '    '   Assign column number of starting point for insertion
            '                x = Range("LastIDColNum").Column

            '                '   Assign number of columns to be inserted
            '                If CalledFrom = "Replicator" Then y = y - Range(TargRng).Columns.Count

            '                IncActiveLine

            '                If y > 0 Then

            '                    'Range("MacStartTime") = Time   ' timer if needed

            '                    For Each s In DvptSheets()

            '                        With Sheets(s)

            '                            .Visible = True
            '                            .Unprotect Password:=PW

            '            End With

            '                    Next s

            '                    IncActiveLine

            '                    Sheets(DvptSheets).Select
            '                    Sheets("Development BP Assumptions").Activate


            '                    '   Set and select range of columns to be inserted
            '                    Range(Columns(x), Columns(x + 1 - 1)).Select
            '                    Selection.Insert Shift:=xlToRight
            '        Columns(x - 1).Select

            '                    'JW 5/5/22 Note 84176
            '                    DoEvents

            '        Selection.Copy

            '                    IncActiveLine

            '                    Range(Columns(x), Columns(x + 1 - 1)).Select
            '                    Cells(1, x + 1 - 1).Activate




            '        ActiveSheet.Paste
            '                    Range("B1").Select



            '        Sheets("Development BP Assumptions").Select
            '                    Range("LastIDColNum").Offset(0, 1).Select

            '                    ' Clear contents of unprotected cells in the new columns
            '                    LastRow = ActiveCell.SpecialCells(xlLastCell).Row

            '                    'Deprecated on instruction AY JW Note 86992



            '                    '        For i = 1 To LastRow
            '                    '
            '                    '            For j = 1 To 1
            '                    '
            '                    '                Range("A1").Cells(i, x).Select
            '                    '                If Not ActiveCell.Locked Then ActiveCell.ClearContents
            '                    '
            '                    '                'MsgBox "Stop 3" 'Note 84176
            '                    '
            '                    '            Next j
            '                    '
            '                    '        Next i


            '                    If y > 1 Then


            '            y = y - 1
            '                        x = x + 1

            '                        ' INSERT OTHER COLUMNS
            '                        Sheets(DvptSheets).Select
            '                        Sheets("Development BP Assumptions").Activate

            '                        IncActiveLine

            '                        '   Set and select range of columns to be inserted
            '                        Range(Columns(x), Columns(x + y - 1)).Select


            '            Selection.Insert Shift:=xlToRight

            '            IncActiveLine

            '                        Columns(x - 1).Select


            '            Selection.Copy
            '                        Range(Columns(x), Columns(x + y - 1)).Select
            '                        Cells(1, x + y - 1).Activate

            '                        IncActiveLine


            '            ActiveSheet.Paste
            '                        Range("B1").Select

            '                        Sheets("Development BP Assumptions").Select
            '                        Range("LastIDColNum").Offset(0, 1).Select

            '                    End If

            '                    ABVCalculate



            '                    Range("A1").Cells(1, x).Activate

            '                    If CalledFrom = "BP" Then

            '                        For Each s In DvptSheets()

            '                            Call fncPrtSht(Sheets(s).Name)
            '                            If Left(s, 6) = "Hidden" Then Sheets(s).Visible = False

            '                        Next s
            '                        ResumeCodeRun "InsertDvptColumns"

            '        End If

            '                    Call DevptSynchroniseTransactionalDBSheet

            '                    If CalledFrom = "BP" Then

            '                        Application.ScreenUpdating = True
            '                        MsgBox "Insertion of columns complete", vbOKOnly, "Insert columns"

            '        End If

            '                End If

            '                EndCodeRun


            '            End Sub


            '            Sub Delete_Dvpt_Columns()

            '                'CodeSafe JW 26/4/22

            '                ' Calling procedure for DeleteDvptColumns

            '                InitiateCodeRun "Delete_Dvpt_Columns", True

            '    Dim x As Integer, y As Integer
            '                Dim Title As String
            '                Title = "Delete Identified Development Columns"

            '                '   Assign column number of starting point for deletion
            '                x = Range("LastIDColNum").Column - 1

            '                '   Assign number of columns to be deleted
            '                y = Application.InputBox(Prompt:="How many columns do you wish to delete?", Title:=Title, Default:=1, Type:=1)     '   Integer type

            '                Call DeleteDvptColumns(x, y, "BP")

            '                ResumeCodeRun "Delete_Dvpt_Columns"

            '    Call DevptSynchroniseTransactionalDBSheet

            '                EndCodeRun True

            'End Sub
            '            Sub DeleteDvptColumns(x As Integer, y As Integer, CalledFrom As String)

            '                'CodeSafe JW 26/4/22

            '                ' Procedure to delete from the right any number of "Identified" scheme columns
            '                ' on the "Development BP Assumptions" sheet and its dependent sheets.
            '                ' Called by Delete_Dvpt_Columns
            '                Dim LastColNum As Integer
            '                Dim FirstColNum As Integer
            '                Dim CalcMethod As Integer
            '                Dim ResponseMessage
            '                Dim s

            '                InitiateCodeRun "DeleteDvptColumns"

            '    If y > 0 Then

            '                    FirstColNum = Range("HouseTypeIn").Column

            '                    If x - y < FirstColNum + MinColsID - 1 And CalledFrom = "BP" Then  ' check minimum dvpt categories retained

            '                        MsgBox "No more than " & x - FirstColNum - MinColsID + 1 & " columns may be deleted at this time"

            '        Else

            '                        If x - y < FirstColNum + MinColsID - 1 Then

            '                            y = x - FirstColNum - MinColsID + 1

            '                        End If

            '                        If y > 0 Then

            '                            For Each s In DvptSheets()

            '                                With Sheets(s)

            '                                    .Visible = True
            '                                    .Unprotect Password:=PW

            '                    End With

            '                            Next s

            '                            Sheets(DvptSheets).Select
            '                            Sheets("Development BP Assumptions").Activate

            '                            '   Set and select range of columns to be deleted
            '                            Range(Columns(x), Columns(x - y + 1)).Select
            '                            Selection.Delete Shift:=xlToLeft
            '                Range("B1").Select

            '                            Sheets("Development BP Assumptions").Select
            '                            Range("A1").Cells(1, x - y).Activate


            '                ABVCalculate

            '                            For Each s In DvptSheets()

            '                                Call fncPrtSht(Sheets(s).Name)

            '                                If Left(s, 6) = "Hidden" Then Sheets(s).Visible = False

            '                            Next s

            '                            Call DevptSynchroniseTransactionalDBSheet

            '                            ResumeCodeRun "DeleteDvptColumns"

            '                If CalledFrom = "BP" Then MsgBox "Deletion of columns complete"

            '            End If

            '                    End If

            '                End If

            '                EndCodeRun

            '            End Sub
            '            Sub Insert_Multi_Dvpt_Columns()

            '                'CodeSafe JW 26/4/22

            '                ' Procedure to insert at the right any number of "Multi-year" scheme columns
            '                ' on the "Development BP Assumptions" sheet and its dependent sheets.
            '                Dim y As Integer, LastColNum As Integer
            '                Dim Title As String

            '                InitiateCodeRun "Insert_Multi_Dvpt_Columns", True

            '    Title = "Insert Multi-year Development Columns"

            '                '   Assign number of columns to be inserted
            '                y = Application.InputBox(Prompt:="How many columns do you wish to insert?", Title:=Title, Default:=1, Type:=1)                                     '   Integer type

            '                If y > 0 Then

            '                    Call Run_Insert_Multi_Dvpt_Columns(y, "Null", "BP")

            '                    ResumeCodeRun "Insert_Multi_Dvpt_Columns"

            '    End If

            '                Call DevptSynchroniseTransactionalDBSheet

            '                EndCodeRun True

            'End Sub
            '            Sub Run_Insert_Multi_Dvpt_Columns(y As Integer, TargRng As String, CalledFrom As String)

            '                'CodeSafe JW 26/4/22

            '                Dim LastColNum As Integer, x As Integer
            '                Dim i As Integer, j As Integer, LastRow As Integer
            '                Dim CalcMethod As Integer
            '                Dim s

            '                InitiateCodeRun "Run_Insert_Multi_Dvpt_Columns"
            '    '   Assign column number of starting point for insertion
            '                x = Range("LastMYColNum").Column

            '                '   Assign number of columns to be inserted
            '                If CalledFrom <> "BP" Then y = y - Range(TargRng).Columns.Count

            '                If y > 0 Then

            '                    For Each s In DvptSheets()

            '                        With Sheets(s)

            '                            .Visible = True
            '                            .Unprotect Password:=PW

            '            End With

            '                    Next s

            '                    Sheets(DvptSheets).Select
            '                    Sheets("Development BP Assumptions").Activate


            '                    '   Set and select range of columns to be inserted


            '        Range(Columns(x), Columns(x + y - 1)).Select
            '                    Selection.Insert Shift:=xlToRight
            '        Columns(x - 1).Select


            '        Selection.Copy

            '                    Range(Columns(x), Columns(x + y - 1)).Select
            '                    Cells(1, x + y - 1).Activate


            '        ActiveSheet.Paste

            '                    Range("B1").Select

            '                    Sheets("Development BP Assumptions").Select

            '                    Range("LastMYColNum").Offset(0, 1).Select

            '                    ' Clear contents of unprotected cells in the new columns

            '                    LastRow = ActiveCell.SpecialCells(xlLastCell).Row


            ''        For i = 1 To LastRow
            '                    '
            '                    '            For j = 1 To y
            '                    '
            '                    '                Range("A1").Cells(i, x + j - 1).Select
            '                    '                If Not ActiveCell.Locked Then ActiveCell.ClearContents
            '                    '
            '                    '            Next j
            '                    '
            '                    '        Next i


            '        ABVCalculate

            '                    Range("A1").Cells(1, x).Activate

            '                    If CalledFrom = "BP" Then

            '                        For Each s In DvptSheets()

            '                            Call fncPrtSht(Sheets(s).Name)
            '                            If Left(s, 6) = "Hidden" Then Sheets(s).Visible = False

            '                        Next s

            '                        ResumeCodeRun "Run_Insert_Multi_Dvpt_Columns"

            '        End If

            '                    Call DevptSynchroniseTransactionalDBSheet

            '                    If CalledFrom = "BP" Then

            '                        Application.ScreenUpdating = True
            '                        MsgBox "Insertion of columns complete", vbOKOnly, "Insert columns"

            '        End If

            '                End If

            '                EndCodeRun

            '            End Sub
            '            Sub Delete_Multi_Dvpt_Columns()

            '                'CodeSafe JW 26/4/22

            '                ' Procedure to delete from the right any number of "Multi-year" scheme columns
            '                ' on the "Development BP Assumptions" sheet and its dependent sheets.
            '                Dim x As Integer, y As Integer, LastColNum As Integer
            '                Dim FirstColNum As Integer
            '                Dim CalcMethod As Integer
            '                Dim ResponseMessage, Title As String
            '                Dim s

            '                InitiateCodeRun "Delete_Multi_Dvpt_Columns", True

            '    Title = "Delete Multi-year Development Columns"

            '                '   Assign column number of starting point for deletion
            '                x = Range("LastMYColNum").Column - 1

            '                '   Assign number of columns to be deleted
            '                y = Application.InputBox(Prompt:="How many columns do you wish to delete?", Title:=Title, Default:=1, Type:=1)           '   Integer type

            '                If y > 0 Then

            '                    FirstColNum = Range("LastIDColNum").Column + 1

            '                    If x - y >= FirstColNum + MinColsMY - 1 Then  ' check minimum dvpt categories retained

            '                        Application.ScreenUpdating = False
            '                        Application.DisplayAlerts = False

            '                        CalcMethod = Application.Calculation
            '                        Application.Calculation = xlManual

            '                        For Each s In DvptSheets()

            '                            With Sheets(s)

            '                                .Visible = True
            '                                .Unprotect Password:=PW

            '                End With
            '                        Next s
            '                        Sheets(DvptSheets).Select
            '                        Sheets("Development BP Assumptions").Activate

            '                        '   Set and select range of columns to be deleted
            '                        Range(Columns(x), Columns(x - y + 1)).Select
            '                        Selection.Delete Shift:=xlToLeft
            '            Range("B1").Select

            '                        Sheets("Development BP Assumptions").Select
            '                        Range("A1").Cells(1, x - y).Activate

            '            ABVCalculate

            '                        For Each s In DvptSheets()

            '                            Call fncPrtSht(Sheets(s).Name)
            '                            If Left(s, 6) = "Hidden" Then Sheets(s).Visible = False

            '                        Next s

            '                        ResumeCodeRun "Delete_Multi_Dvpt_Columns"

            '            Call DevptSynchroniseTransactionalDBSheet

            '                        Application.ScreenUpdating = True
            '                        ResponseMessage = MsgBox(Prompt:="Deletion of columns complete", Title:=Title)

            '                    Else

            '                        ResponseMessage = MsgBox(Prompt:="No more than " & x - FirstColNum - MinColsMY + 1 & " columns may be deleted at this time", Buttons:=vbExclamation, Title:=Title)

            '                    End If

            '                End If

            '                EndCodeRun True

            'End Sub
            '            Sub Delete_Spec_Dvpt_Column()


            '                ' Procedure to delete a single selected column from
            '                ' the "Development BP Assumptions" sheet and its dependent sheets.
            '                Dim y As Integer, LastColNum As Integer
            '                Dim FirstColNum As Integer
            '                Dim LastIDColNum As Integer, FirstMYColNum As Integer
            '                Dim ChosenCell As Range, ChosenColumn As String, DefaultAddr As String
            '                Dim CalcMethod As Integer
            '                Dim ResponseMessage, Answer, Title As String
            '                Dim s

            '                InitiateCodeRun "Delete_Spec_Dvpt_Column", True

            '    Title = "Delete Selected Development Column"

            '                '   Assign column to be deleted
            '                DefaultAddr = ActiveCell.Address

            '                On Error GoTo Cancelled

            '    Set ChosenCell = Application.InputBox(Prompt:="Please select one cell in the column you wish to delete", Title:=Title, Default:=DefaultAddr, Type:=8)                                    '   Address type

            '    y = ChosenCell.Column

            '                If y <> 0 Then

            '                    FirstColNum = Range("HouseTypeIn").Column + MinColsID
            '                    LastIDColNum = Range("LastIDColNum").Column - 1
            '                    FirstMYColNum = Range("LastIDColNum").Column + MinColsMY + 1
            '                    LastColNum = Range("LastMYColNum").Column - 1

            '                    Columns(y).Select

            '                    If y <= 26 Then

            '                        ChosenColumn = Mid(ActiveCell.Address, 2, 1)

            '                    Else

            '                        ChosenColumn = Mid(ActiveCell.Address, 2, 2)

            '                    End If

            '                    If (y >= FirstColNum And y <= LastIDColNum) Or (y >= FirstMYColNum And y <= LastColNum) Then

            '                        Answer = MsgBox(Prompt:="Are you sure that you want" & Chr(13) & "to delete Column " & ChosenColumn & "?", Buttons:=vbYesNo + vbQuestion)
            '                        If Answer = vbYes Then


            '                            For Each s In DvptSheets()

            '                                With Sheets(s)

            '                                    .Visible = True
            '                                    .Unprotect Password:=PW

            '                    End With

            '                            Next s
            '                            Sheets(DvptSheets).Select

            '                            '   Set and select range of the column to be deleted
            '                            Columns(y).Select
            '                            Selection.Delete Shift:=xlToLeft
            '                Range("B1").Select

            '                            Sheets("Development BP Assumptions").Select
            '                            Range("A1").Cells(1, y).Activate


            '                ABVCalculate

            '                            For Each s In DvptSheets()

            '                                Call fncPrtSht(Sheets(s).Name)

            '                                If Left(s, 6) = "Hidden" Then Sheets(s).Visible = False

            '                            Next s

            '                            ResumeCodeRun "Delete_Spec_Dvpt_Column"

            '                Call DevptSynchroniseTransactionalDBSheet

            '                            Application.ScreenUpdating = True
            '                            ResponseMessage = MsgBox(Prompt:="Deletion of column " & ChosenColumn & " is complete", Title:=Title)

            '                        Else

            '                            Answer = MsgBox(Prompt:="Column " & ChosenColumn & " has NOT been deleted", Buttons:=vbExclamation, Title:=Title)

            '                        End If

            '                    Else

            '                        ResponseMessage = MsgBox(Prompt:="Columns may not be deleted from the first " & MinColsID & " columns" & Chr(13) & "or after the last Identified Scheme column.", Buttons:=vbExclamation, Title:=Title)

            '                    End If

            '                End If

            'Cancelled:

            '                EndCodeRun True

            'End Sub
            '            Sub Insert_Spec_Dvpt_Column()

            '                ' Procedure to delete a single selected column from
            '                ' the "Development BP Assumptions" sheet and its dependent sheets.
            '                Dim x As Integer, y As Integer, LastColNum As Integer, LastRow As Integer
            '                Dim i As Integer, j As Integer
            '                Dim FirstColNum As Integer
            '                Dim LastIDColNum As Integer, FirstMYColNum As Integer
            '                Dim ChosenCell As Range, ChosenColumn As String, DefaultAddr As String
            '                Dim CalcMethod As Integer
            '                Dim ResponseMessage, Answer, Title As String
            '                Dim s

            '                InitiateCodeRun "Insert_Spec_Dvpt_Column", True

            '    Title = "Insert Selected Development Columns"

            '                '   Assign column to be deleted
            '                DefaultAddr = ActiveCell.Address

            '                On Error GoTo Cancelled

            '    Set ChosenCell = Application.InputBox(Prompt:="Please select one cell in the column where" & Chr(13) & "where you wish to insert new columns", Title:=Title, Default:=DefaultAddr, Type:=8)                                                     '   Address type

            '    x = ChosenCell.Column

            '                If x <> 0 Then

            '                    FirstColNum = Range("HouseTypeIn").Column + MinColsID
            '                    LastIDColNum = Range("LastIDColNum").Column - 1
            '                    FirstMYColNum = Range("LastIDColNum").Column + MinColsMY + 1
            '                    LastColNum = Range("LastMYColNum").Column - 1

            '                    Columns(x).Select

            '                    If x <= 26 Then

            '                        ChosenColumn = Mid(ActiveCell.Address, 2, 1)

            '                    Else

            '                        ChosenColumn = Mid(ActiveCell.Address, 2, 2)

            '                    End If

            '                    If (x >= FirstColNum And x <= LastIDColNum) Or (x >= FirstMYColNum And x <= LastColNum) Then

            '                        '   Assign number of columns to be inserted
            '                        y = Application.InputBox(Prompt:="How many columns do you wish to insert?", Title:=Title, Default:=1, Type:=1)                  '   Integer type

            '                        If y > 0 Then

            '                            For Each s In DvptSheets()

            '                                With Sheets(s)

            '                                    .Visible = True
            '                                    .Unprotect Password:=PW

            '                    End With

            '                            Next s

            '                            Sheets(DvptSheets).Select
            '                            Sheets("Development BP Assumptions").Activate

            '                            '   Set and select range of columns to be inserted
            '                            Range(Columns(x), Columns(x + 1 - 1)).Select
            '                            Selection.Insert Shift:=xlToRight
            '                Columns(x - 1).Select
            '                            Selection.Copy
            '                            Range(Columns(x), Columns(x + 1 - 1)).Select
            '                            Cells(1, x + 1 - 1).Activate
            '                            ActiveSheet.Paste
            '                            Range("B1").Select

            '                            Sheets("Development BP Assumptions").Select
            '                            Range("LastIDColNum").Offset(0, 1).Select

            '                            ' Clear contents of unprotected cells in the new columns
            '                            LastRow = ActiveCell.SpecialCells(xlLastCell).Row

            '                            '                For i = 1 To LastRow
            '                            '
            '                            '                    For j = 1 To 1
            '                            '
            '                            '                        Range("A1").Cells(i, x).Select
            '                            '                        If Not ActiveCell.Locked Then ActiveCell.ClearContents
            '                            '
            '                            '                    Next j
            '                            '
            '                            '                Next i

            '                            If y > 1 Then
            '                                y = y - 1
            '                                x = x + 1

            '                                ' INSERT OTHER COLUMNS
            '                                Sheets(DvptSheets).Select
            '                                Sheets("Development BP Assumptions").Activate

            '                                '   Set and select range of columns to be inserted
            '                                Range(Columns(x), Columns(x + y - 1)).Select
            '                                Selection.Insert Shift:=xlToRight
            '                    Columns(x - 1).Select
            '                                Selection.Copy
            '                                Range(Columns(x), Columns(x + y - 1)).Select
            '                                Cells(1, x + y - 1).Activate
            '                                ActiveSheet.Paste
            '                                Range("B1").Select

            '                                Sheets("Development BP Assumptions").Select
            '                                Range("LastIDColNum").Offset(0, 1).Select

            '                            End If

            '                ABVCalculate

            '                            For Each s In DvptSheets()

            '                                Call fncPrtSht(Sheets(s).Name)

            '                                If Left(s, 6) = "Hidden" Then Sheets(s).Visible = False

            '                            Next s

            '                            ResumeCodeRun "Insert_Spec_Dvpt_Column"

            '                Call DevptSynchroniseTransactionalDBSheet

            '                            ResponseMessage = MsgBox(Prompt:="Insertion at column " & ChosenColumn & " is complete", Title:=Title)

            '                        Else

            '                            Answer = MsgBox(Prompt:="Columns have NOT been inserted", Buttons:=vbExclamation, Title:=Title)

            '                        End If

            '                    Else

            '                        ResponseMessage = MsgBox(Prompt:="Columns may not be inserted among the first " & MinColsID & " columns" & Chr(13) & "or after the last Identified Scheme column.", Buttons:=vbExclamation, Title:=Title)

            '                    End If

            '                End If

            'Cancelled:

            '                EndCodeRun True

            'End Sub




        End Class

    End Class

End Namespace
