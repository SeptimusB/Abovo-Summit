Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.WorkbookManager
Imports Abovo.WSSecurity
Imports DevExpress.CodeParser
Imports DevExpress.Pdf.Native.BouncyCastle.Asn1.X509
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraRichEdit.Layout

Namespace Abovo
    Public Class WorkbookManager

#Region "Worksheet manipulation"

        Private Shared Sub NotifyTransactionalDBStructuralChange(ByVal ModelID As Integer, ByVal TargetNamedRange As String)

            Try

                If ExcelModels Is Nothing Then Return
                If ModelID < 0 OrElse ModelID >= ExcelModels.Length Then Return
                If ExcelModels(ModelID) Is Nothing Then Return
                If ExcelModels(ModelID).TransDBSync Is Nothing Then Return

                ExcelModels(ModelID).TransDBSync.SynchroniseForNamedRange(TargetNamedRange)

            Catch ex As Exception

#If DEBUG Then
#End If

            End Try

        End Sub

        Public Shared Function DoesNRExist(ModelID As Integer, RangeName As String) As Boolean

            Dim WB As DevExpress.Spreadsheet.IWorkbook = ExcelModels(ModelID).WB
            Dim TargetdefinedRange As DevExpress.Spreadsheet.DefinedName = WB.DefinedNames.GetDefinedName(RangeName)
            If TargetdefinedRange Is Nothing Then
                Return False
            Else
                Return True
            End If

        End Function
        Public Shared Function GetRangeRows(ModelID As Integer, RangeName As String) As Integer

            Dim WB As DevExpress.Spreadsheet.IWorkbook = ExcelModels(ModelID).WB

            Dim TargetdefinedRange As DevExpress.Spreadsheet.DefinedName = WB.DefinedNames.GetDefinedName(RangeName)
            Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = TargetdefinedRange.Range

            Return CRTargetRange.RowCount

        End Function

        Public Shared Function GetRangeColumns(ModelID As Integer, RangeName As String) As Integer

            Dim WB As DevExpress.Spreadsheet.IWorkbook = ExcelModels(ModelID).WB

            Dim TargetdefinedRange As DevExpress.Spreadsheet.DefinedName = WB.DefinedNames.GetDefinedName(RangeName)
            Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = TargetdefinedRange.Range

            Return CRTargetRange.ColumnCount

        End Function

        Public Shared Function SetRangeRowsToSize(ModelID As Integer,
                                                     RangeName As String,
                                                     SetSize As Integer,
                                                     Optional ByVal IgnoreFinalRow As Boolean = False,
                                                     Optional ByVal MinimumRowsToRetain As Integer = 3) As AbovoTransaction

            Dim ThisTrans As New AbovoTransaction
            Dim CurrentSize As Integer = GetRangeRows(ModelID, RangeName)
            If CurrentSize = SetSize Then
                ThisTrans.BError = False
                ThisTrans.StringReturn = "No change needed"
                Return ThisTrans
            End If
            If CurrentSize < SetSize Then
                ThisTrans = InsertRows(ModelID, RangeName, SetSize - CurrentSize, , IgnoreFinalRow)
            Else
                ThisTrans = DeleteRows(ModelID, RangeName, CurrentSize - SetSize, , IgnoreFinalRow, MinimumRowsToRetain)
            End If
            Return ThisTrans

        End Function

        Public Shared Function SetRangeColsToSize(ModelID As Integer, RangeName As String, SetSize As Integer) As AbovoTransaction

            Dim ThisTrans As New AbovoTransaction
            Dim CurrentSize As Integer = GetRangeColumns(ModelID, RangeName)
            If CurrentSize = SetSize Then
                ThisTrans.BError = False
                ThisTrans.StringReturn = "No change needed"
                Return ThisTrans
            End If
            If CurrentSize < SetSize Then
                ThisTrans = InsertColumns(ModelID, RangeName, SetSize - CurrentSize)
            Else
                ThisTrans = DeleteColumns(ModelID, RangeName, CurrentSize - SetSize)
            End If
            Return ThisTrans

        End Function

        Public Shared Function InsertRows(ModelID As Integer, TargetNamedRange As String, Optional ByVal RowsToAdd As Integer = 1, Optional ByVal JustFormats As Boolean = False, Optional ByVal IgnoreFinalRow As Boolean = False) As AbovoTransaction

            Dim ThisTrans As New AbovoTransaction

            If RowsToAdd <= 0 Then

                ThisTrans.BError = False
                ThisTrans.StringReturn = "No rows requested."
                Return ThisTrans

            End If

            Dim WB As IWorkbook = ExcelModels(ModelID).WB

            If WB Is Nothing Then

                ThisTrans.BError = True
                ThisTrans.StringReturn = "Workbook is not available."
                Return ThisTrans

            End If

            Dim TargetDefinedName As DefinedName = WB.DefinedNames.GetDefinedName(TargetNamedRange)

            If TargetDefinedName Is Nothing OrElse TargetDefinedName.Range Is Nothing Then

                ThisTrans.BError = True
                ThisTrans.StringReturn = "Named range '" & TargetNamedRange & "' was not found."
                Return ThisTrans

            End If

            Dim WSTarget As Worksheet = TargetDefinedName.Range.Worksheet
            Dim bIsGlobal As Boolean = TargetDefinedName.IsGlobal

            'Snapshot the named-range coordinates before inserting anything.
            'Do not keep using TargetDefinedName.Range after the structural edit:
            'DevExpress may already have adjusted the live defined-name reference.
            Dim RangeLeft As Integer = TargetDefinedName.Range.LeftColumnIndex
            Dim RangeTop As Integer = TargetDefinedName.Range.TopRowIndex
            Dim RangeRight As Integer = TargetDefinedName.Range.RightColumnIndex
            Dim RangeBottom As Integer = TargetDefinedName.Range.BottomRowIndex

            Dim PriorBottomRow As Integer = RangeBottom

            If IgnoreFinalRow Then PriorBottomRow -= 1

            If PriorBottomRow < 0 Then

                ThisTrans.BError = True
                ThisTrans.StringReturn = "The insertion row for named range '" & TargetNamedRange & "' is invalid."
                Return ThisTrans

            End If

            'GetUsedRange includes cells that contain values, formulas OR formatting.
            'The old routine copied a complete Excel row (16,384 cells).  Restricting
            'the template copy to this span preserves all actual worksheet content
            'while avoiding work on thousands of unused cells.
            Dim UsedRange As CellRange = WSTarget.GetUsedRange()

            Dim UsedLeft As Integer = UsedRange.LeftColumnIndex
            Dim UsedRight As Integer = UsedRange.RightColumnIndex

            'The named range itself must always be included even on a sparse sheet.
            If RangeLeft < UsedLeft Then UsedLeft = RangeLeft
            If RangeRight > UsedRight Then UsedRight = RangeRight

            Dim TemplateRowHeight As Single = WSTarget.Rows(PriorBottomRow).Height
            Dim WasProtected As Boolean = WSTarget.IsProtected
            Dim UpdateStarted As Boolean = False

#If DEBUG Then
            Dim PerfTimer As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Dim InsertElapsed As Long = 0
            Dim TemplateElapsed As Long = 0
#End If

            Try

                WB.BeginUpdate()
                UpdateStarted = True

                If WasProtected Then UNProtectWS(ModelID, WSTarget.Name)

                'Insert all rows in one structural operation.  DevExpress updates
                'workbook references as part of the row insertion.
                WSTarget.Rows.Insert(PriorBottomRow + 1, RowsToAdd)

#If DEBUG Then
                InsertElapsed = PerfTimer.ElapsedMilliseconds
#End If

                Dim SourceTemplate As CellRange =
                    WSTarget.Range.FromLTRB(UsedLeft,
                                            PriorBottomRow,
                                            UsedRight,
                                            PriorBottomRow)

                Dim FirstNewRow As CellRange =
                    WSTarget.Range.FromLTRB(UsedLeft,
                                            PriorBottomRow + 1,
                                            UsedRight,
                                            PriorBottomRow + 1)

                'Maintain the legacy behaviour for the first inserted row:
                'copy the source row, then remove unlocked constants while retaining
                'formulas, locked values and formatting.
                If JustFormats Then

                    FirstNewRow.CopyFrom(SourceTemplate, PasteSpecial.Formats)

                Else

                    FirstNewRow.CopyFrom(SourceTemplate, PasteSpecial.All)

                    For Col As Integer = UsedLeft To UsedRight

                        Dim CellExamine As Cell = WSTarget.Cells(PriorBottomRow + 1, Col)

                        If Not CellExamine.Protection.Locked AndAlso Not CellExamine.HasFormula Then
                            CellExamine.ClearContents()
                        End If

                    Next

                End If

                WSTarget.Rows(PriorBottomRow + 1).Height = TemplateRowHeight

                'The old implementation performed two full-row CopyFrom calls for
                'every additional row.  CopyFrom can repeat a smaller source range
                'through a larger destination range, and adjusts relative references.
                'This reduces N copies to one contiguous operation.
                If RowsToAdd > 1 Then

                    Dim RemainingRows As CellRange =
                        WSTarget.Range.FromLTRB(UsedLeft,
                                                PriorBottomRow + 2,
                                                UsedRight,
                                                PriorBottomRow + RowsToAdd)

                    If JustFormats Then
                        RemainingRows.CopyFrom(FirstNewRow, PasteSpecial.Formats)
                    Else
                        RemainingRows.CopyFrom(FirstNewRow, PasteSpecial.Formulas Or PasteSpecial.Formats)
                    End If

                    For RowIndex As Integer = PriorBottomRow + 2 To PriorBottomRow + RowsToAdd
                        WSTarget.Rows(RowIndex).Height = TemplateRowHeight
                    Next

                End If

#If DEBUG Then
                TemplateElapsed = PerfTimer.ElapsedMilliseconds - InsertElapsed
#End If

                'Explicitly set the intended named-range extent.  We use the saved
                'pre-insert coordinates so the result does not depend on whether the
                'live DefinedName was automatically adjusted during Rows.Insert.
                Dim ExpandedRange As CellRange =
                    WSTarget.Range.FromLTRB(RangeLeft,
                                            RangeTop,
                                            RangeRight,
                                            RangeBottom + RowsToAdd)

                If bIsGlobal Then
                    WB.DefinedNames.GetDefinedName(TargetNamedRange).Range = ExpandedRange
                Else
                    WSTarget.DefinedNames.GetDefinedName(TargetNamedRange).Range = ExpandedRange
                End If

                ThisTrans.BError = False

#If DEBUG Then
#End If

            Catch ex As Exception

                ThisTrans.BError = True
                ThisTrans.StringReturn = ex.Message

            Finally

                If WasProtected AndAlso WSTarget IsNot Nothing Then
                    ProtectWS(ModelID, WSTarget.Name)
                End If

                If UpdateStarted Then WB.EndUpdate()

            End Try

            If Not ThisTrans.BError Then
                NotifyTransactionalDBStructuralChange(ModelID, TargetNamedRange)
            End If

            Return ThisTrans

        End Function
        Public Shared Function InsertColumns(ModelID As Integer, TargetNamedRange As String, Optional ByVal ColsToAdd As Integer = 1, Optional ByVal JustFormats As Boolean = False, Optional ByVal IgnoreFinalCol As Boolean = False) As AbovoTransaction

            Dim ThisTrans As New AbovoTransaction

            If ColsToAdd <= 0 Then

                ThisTrans.BError = False
                ThisTrans.StringReturn = "No columns requested."
                Return ThisTrans

            End If

            Dim WB As IWorkbook = ExcelModels(ModelID).WB

            If WB Is Nothing Then

                ThisTrans.BError = True
                ThisTrans.StringReturn = "Workbook is not available."
                Return ThisTrans

            End If

            Dim TargetDefinedName As DefinedName = WB.DefinedNames.GetDefinedName(TargetNamedRange)

            If TargetDefinedName Is Nothing OrElse TargetDefinedName.Range Is Nothing Then

                ThisTrans.BError = True
                ThisTrans.StringReturn = "Named range '" & TargetNamedRange & "' was not found."
                Return ThisTrans

            End If

            Dim WSTarget As Worksheet = TargetDefinedName.Range.Worksheet
            Dim bIsGlobal As Boolean = TargetDefinedName.IsGlobal

            Dim RangeLeft As Integer = TargetDefinedName.Range.LeftColumnIndex
            Dim RangeTop As Integer = TargetDefinedName.Range.TopRowIndex
            Dim RangeRight As Integer = TargetDefinedName.Range.RightColumnIndex
            Dim RangeBottom As Integer = TargetDefinedName.Range.BottomRowIndex

            Dim PriorRightCol As Integer = RangeRight

            If IgnoreFinalCol Then PriorRightCol -= 1

            If PriorRightCol <= 0 Then

                ThisTrans.BError = True
                ThisTrans.StringReturn = "The insertion column for named range '" & TargetNamedRange & "' is invalid."
                Return ThisTrans

            End If

            Dim UsedRange As CellRange = WSTarget.GetUsedRange()

            Dim UsedTop As Integer = UsedRange.TopRowIndex
            Dim UsedBottom As Integer = UsedRange.BottomRowIndex

            If RangeTop < UsedTop Then UsedTop = RangeTop
            If RangeBottom > UsedBottom Then UsedBottom = RangeBottom

            Dim TemplateColumnIndex As Integer = PriorRightCol - 1
            Dim TemplateColumnWidth As Single = WSTarget.Columns(TemplateColumnIndex).Width

            Dim WasProtected As Boolean = WSTarget.IsProtected
            Dim UpdateStarted As Boolean = False

#If DEBUG Then
            Dim PerfTimer As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Dim InsertElapsed As Long = 0
            Dim TemplateElapsed As Long = 0
#End If

            Try

                WB.BeginUpdate()
                UpdateStarted = True

                If WasProtected Then UNProtectWS(ModelID, WSTarget.Name)

                'Preserve the legacy insertion point: new columns are inserted to
                'the left of the named range's prior right-hand column.
                WSTarget.Columns.Insert(PriorRightCol, ColsToAdd)

#If DEBUG Then
                InsertElapsed = PerfTimer.ElapsedMilliseconds
#End If

                Dim SourceTemplate As CellRange =
                    WSTarget.Range.FromLTRB(TemplateColumnIndex,
                                            UsedTop,
                                            TemplateColumnIndex,
                                            UsedBottom)

                Dim FirstNewColumn As CellRange =
                    WSTarget.Range.FromLTRB(PriorRightCol,
                                            UsedTop,
                                            PriorRightCol,
                                            UsedBottom)

                If JustFormats Then

                    FirstNewColumn.CopyFrom(SourceTemplate, PasteSpecial.Formats)

                Else

                    FirstNewColumn.CopyFrom(SourceTemplate, PasteSpecial.All)

                    'The previous routine did this one cell at a time after copying
                    'an entire 1,048,576-cell worksheet column.  Restrict the scan to
                    'the used rows only.
                    For RowIndex As Integer = UsedTop To UsedBottom

                        Dim CellExamine As Cell = WSTarget.Cells(RowIndex, PriorRightCol)

                        If Not CellExamine.Protection.Locked AndAlso Not CellExamine.HasFormula Then
                            CellExamine.ClearContents()
                        End If

                    Next

                End If

                WSTarget.Columns(PriorRightCol).Width = TemplateColumnWidth

                'Preserve the current routine's exact expansion semantics: after the
                'first copied column it copies the template across ColsToAdd further
                'columns.  This includes the shifted former final column.  Although
                'that behaviour deserves separate review, it is intentionally NOT
                'changed in this performance-only pass.
                Dim LastTemplateColumn As Integer = PriorRightCol + ColsToAdd

                If LastTemplateColumn > PriorRightCol Then

                    Dim RemainingColumns As CellRange =
                        WSTarget.Range.FromLTRB(PriorRightCol + 1,
                                                UsedTop,
                                                LastTemplateColumn,
                                                UsedBottom)

                    If JustFormats Then
                        RemainingColumns.CopyFrom(FirstNewColumn, PasteSpecial.Formats)
                    Else
                        RemainingColumns.CopyFrom(FirstNewColumn, PasteSpecial.All)
                    End If

                    For ColIndex As Integer = PriorRightCol + 1 To LastTemplateColumn
                        WSTarget.Columns(ColIndex).Width = TemplateColumnWidth
                    Next

                End If

#If DEBUG Then
                TemplateElapsed = PerfTimer.ElapsedMilliseconds - InsertElapsed
#End If

                'InsertColumns historically relied on DevExpress to alter the named
                'range implicitly.  Make the result explicit so later enhancements
                'have one predictable contract.
                Dim ExpandedRange As CellRange =
                    WSTarget.Range.FromLTRB(RangeLeft,
                                            RangeTop,
                                            RangeRight + ColsToAdd,
                                            RangeBottom)

                If bIsGlobal Then
                    WB.DefinedNames.GetDefinedName(TargetNamedRange).Range = ExpandedRange
                Else
                    WSTarget.DefinedNames.GetDefinedName(TargetNamedRange).Range = ExpandedRange
                End If

                ThisTrans.BError = False

#If DEBUG Then
#End If

            Catch ex As Exception

                ThisTrans.BError = True
                ThisTrans.StringReturn = ex.Message

            Finally

                If WasProtected AndAlso WSTarget IsNot Nothing Then
                    ProtectWS(ModelID, WSTarget.Name)
                End If

                If UpdateStarted Then WB.EndUpdate()

            End Try

            If Not ThisTrans.BError Then
                NotifyTransactionalDBStructuralChange(ModelID, TargetNamedRange)
            End If

            Return ThisTrans

        End Function
        Public Shared Function DeleteRows(ModelID As Integer,
                                              TargetNamedRange As String,
                                              Optional ByVal RowsToDelete As Integer = 1,
                                              Optional ByVal JustFormats As Boolean = False,
                                              Optional ByVal IgnoreFinalRow As Boolean = False,
                                              Optional ByVal MinimumRowsToRetain As Integer = 3) As AbovoTransaction

            'On Error GoTo Err_Handler_A

            Dim WB As DevExpress.Spreadsheet.IWorkbook = ExcelModels(ModelID).WB
            Dim ThisTrans As New AbovoTransaction
            Dim TargetDefinedName As DevExpress.Spreadsheet.DefinedName = WB.DefinedNames.GetDefinedName(TargetNamedRange)
            Dim WSTarget As DevExpress.Spreadsheet.Worksheet = TargetDefinedName.Range.Worksheet
            Dim OutMessage As String = ""
            Dim CurrentRangeRows As Integer = TargetDefinedName.Range.RowCount
            Dim PriorBottomRow As Integer = TargetDefinedName.Range.BottomRowIndex

            If IgnoreFinalRow Then

                PriorBottomRow -= 1
                CurrentRangeRows -= 1

            End If

            Dim ActualRowsToDelete As Integer = RowsToDelete

            If CurrentRangeRows - RowsToDelete < MinimumRowsToRetain Then
                ActualRowsToDelete = CurrentRangeRows - MinimumRowsToRetain
                OutMessage = "Maximum rows deletable is " & ActualRowsToDelete & ". "
            End If

            If ActualRowsToDelete <= 0 Then
                OutMessage += " No rows can be deleted."
                ThisTrans.BError = True
                ThisTrans.StringReturn = OutMessage
                Return ThisTrans
            End If

            UNProtectWS(ModelID, WSTarget.Name)

            WSTarget.Rows.Remove(PriorBottomRow - ActualRowsToDelete + 1, ActualRowsToDelete)

            ProtectWS(ModelID, WSTarget.Name)

            TargetDefinedName = Nothing
            WSTarget = Nothing
            WB = Nothing

            ThisTrans.BError = False
            ThisTrans.StringReturn = OutMessage

            ExcelModels(ModelID).ModelSpreadsheetControl.Options.Clipboard.AllowFormulasInBiff8 = False

            NotifyTransactionalDBStructuralChange(ModelID, TargetNamedRange)

            Return ThisTrans

            Exit Function

Err_Handler_A:

            ThisTrans.BError = True
            ThisTrans.StringReturn = Err.Description

            Return ThisTrans

        End Function

        Public Shared Function DeleteColumns(ModelID As Integer, TargetNamedRange As String, Optional ByVal ColsToDelete As Integer = 1, Optional ByVal JustFormats As Boolean = False) As AbovoTransaction

            'On Error GoTo Err_Handler_A

            Dim WB As DevExpress.Spreadsheet.IWorkbook = ExcelModels(ModelID).WB
            Dim ThisTrans As New AbovoTransaction
            Dim TargetDefinedName As DevExpress.Spreadsheet.DefinedName = WB.DefinedNames.GetDefinedName(TargetNamedRange)
            Dim WSTarget As DevExpress.Spreadsheet.Worksheet = TargetDefinedName.Range.Worksheet
            Dim OutMessage As String = ""
            Dim CurrentRangeCols As Integer = TargetDefinedName.Range.ColumnCount
            Dim PriorRightCol As Integer = TargetDefinedName.Range.RightColumnIndex

            Dim ActualColsToDelete As Integer = ColsToDelete

            If CurrentRangeCols - ColsToDelete <= 3 Then
                ActualColsToDelete = CurrentRangeCols - 3
                OutMessage = "Maximum columns deletable is " & ActualColsToDelete & ". "
            End If

            If ActualColsToDelete < 0 Then
                OutMessage += " No columns can be deleted."
                ThisTrans.BError = True
                ThisTrans.StringReturn = OutMessage
                Return ThisTrans
            End If

            UNProtectWS(ModelID, WSTarget.Name)

            WSTarget.Columns.Remove(PriorRightCol - ActualColsToDelete + 1, ActualColsToDelete)

            ProtectWS(ModelID, WSTarget.Name)

            TargetDefinedName = Nothing
            WSTarget = Nothing
            WB = Nothing

            ThisTrans.BError = False
            ThisTrans.StringReturn = OutMessage

            ExcelModels(ModelID).ModelSpreadsheetControl.Options.Clipboard.AllowFormulasInBiff8 = False

            NotifyTransactionalDBStructuralChange(ModelID, TargetNamedRange)

            Return ThisTrans

            Exit Function

Err_Handler_A:

            ThisTrans.BError = True
            ThisTrans.StringReturn = Err.Description

            Return ThisTrans

        End Function


        Public Shared Function DevExpressInsertRows(ModelID As Integer, TargetNamedRange As String, Optional ByVal RowsToAdd As Integer = 1, Optional ByVal JustFormats As Boolean = False, Optional ByVal IgnoreFinalRow As Boolean = False) As AbovoTransaction

            'On Error GoTo Err_Handler_A

            Dim WB As DevExpress.Spreadsheet.IWorkbook = ExcelModels(ModelID).WB
            Dim ThisTrans As New AbovoTransaction
            Dim TargetDefinedName As DevExpress.Spreadsheet.DefinedName = WB.DefinedNames.GetDefinedName(TargetNamedRange)
            Dim bIsGlobal As Boolean = TargetDefinedName.IsGlobal
            Dim CRTarget As DevExpress.Spreadsheet.CellRange

            Dim WSTarget As DevExpress.Spreadsheet.Worksheet = TargetDefinedName.Range.Worksheet

            Dim PriorBottomRow As Integer = TargetDefinedName.Range.BottomRowIndex

            If IgnoreFinalRow Then PriorBottomRow -= 1

            Dim Lref As Integer = TargetDefinedName.Range.LeftColumnIndex
            Dim Rref As Integer = TargetDefinedName.Range.RightColumnIndex
            Dim Tref As Integer = TargetDefinedName.Range.TopRowIndex
            Dim Bref As Integer = TargetDefinedName.Range.BottomRowIndex

            TargetDefinedName = Nothing
            'TargetDefinedName.RefersTo = Nothing


            WSTarget.Rows.Insert(PriorBottomRow + 1, RowsToAdd)

            CRTarget = WSTarget.Range.FromLTRB(Lref, Tref, Rref, Bref + RowsToAdd)

            If bIsGlobal Then

                WB.DefinedNames.GetDefinedName(TargetNamedRange).Range = CRTarget

            Else

                WSTarget.DefinedNames.GetDefinedName(TargetNamedRange).Range = CRTarget

            End If

            Dim CRCopySource As CellRange = WSTarget.Range.FromLTRB(Lref, PriorBottomRow, Rref, PriorBottomRow)

            Dim DestRow As CellRange

            For i = 1 To RowsToAdd

                DestRow = WSTarget.Range.FromLTRB(CRTarget.LeftColumnIndex, PriorBottomRow + i, CRTarget.RightColumnIndex, PriorBottomRow + i)

                If JustFormats Then
                    DestRow.CopyFrom(CRCopySource, PasteSpecial.Formats, True)
                Else
                    DestRow.CopyFrom(CRCopySource, PasteSpecial.Values Or PasteSpecial.Formats, True)
                End If

            Next

            DestRow = Nothing
            CRCopySource = Nothing
            CRTarget = Nothing
            WSTarget = Nothing
            WB = Nothing

            ThisTrans.BError = False

            Return ThisTrans

            Exit Function

Err_Handler_A:

            ThisTrans.BError = True
            ThisTrans.StringReturn = Err.Description

            Return ThisTrans

        End Function
        Public Shared Function InsertColumn(WB As DevExpress.Spreadsheet.Workbook, TargetNamedRange As String, Optional ByVal ColssToAdd As Integer = 1) As AbovoTransaction

            On Error GoTo Err_Handler_A

            Dim ThisTrans As New AbovoTransaction

            Dim TargetdefinedRange As DevExpress.Spreadsheet.DefinedName = WB.DefinedNames.GetDefinedName(TargetNamedRange)

            Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = TargetdefinedRange.Range
            Dim CRTargetWorksheet As DevExpress.Spreadsheet.Worksheet = CRTargetRange.Worksheet
            Dim IntRangeRows As Integer = CRTargetRange.RowCount
            Dim IntRangeCols As Integer = CRTargetRange.ColumnCount

            On Error Resume Next

            CRTargetWorksheet.Unprotect(UnlockPassword)

            On Error GoTo Err_Handler_A

            WB.BeginUpdate()

            Dim IntRightCol As Integer = CRTargetRange.RightColumnIndex

            CRTargetWorksheet.Columns.Insert(IntRightCol)
            CRTargetWorksheet.Columns(IntRightCol).CopyFrom(CRTargetWorksheet.Columns(IntRightCol + 1), PasteSpecial.Formulas)
            CRTargetWorksheet.Columns(IntRightCol).CopyFrom(CRTargetWorksheet.Columns(IntRightCol + 1), PasteSpecial.Formats)
            CRTargetWorksheet.Columns(IntRightCol + 1).ClearContents
            CRTargetRange.Resize(IntRangeRows, IntRangeCols + 1)

            On Error Resume Next

            CRTargetWorksheet.Protect(UnlockPassword, WorksheetProtectionPermissions.Default)

            On Error GoTo Err_Handler_A

            If ColssToAdd > 1 Then

            End If

            WB.EndUpdate()

            ThisTrans.BError = False

            Return ThisTrans

            Exit Function

Err_Handler_A:

            ThisTrans.BError = True
            ThisTrans.StringReturn = Err.Description

            Return ThisTrans

        End Function
#End Region
        Public Shared Sub CopyRowToRangeBottom(ModelID As Integer, RangeName As String)

            Dim WB As IWorkbook = ExcelModels(ModelID).WB

            Dim TargetDefinedRange As DevExpress.Spreadsheet.DefinedName = WB.DefinedNames.GetDefinedName(RangeName)
            Dim CRTargetRange As CellRange
            Dim CRTargetWorksheet As DevExpress.Spreadsheet.Worksheet
            Dim topRow As CellRange

RestartFrom:

            WB = ExcelModels(ModelID).WB
            TargetDefinedRange = WB.DefinedNames.GetDefinedName(RangeName)
            CRTargetRange = TargetDefinedRange.Range
            CRTargetWorksheet = CRTargetRange.Worksheet

            topRow = CRTargetWorksheet.Range.FromLTRB(CRTargetRange.LeftColumnIndex, CRTargetRange.TopRowIndex, CRTargetRange.RightColumnIndex, CRTargetRange.TopRowIndex)

            ' Find Empty Row
            Dim bottomRow As CellRange = Nothing

            For i As Integer = CRTargetRange.TopRowIndex To CRTargetRange.TopRowIndex + CRTargetRange.RowCount

                Dim rowCell As CellRange = CRTargetWorksheet.Range.FromLTRB(CRTargetRange.LeftColumnIndex, i, CRTargetRange.RightColumnIndex, i)

                If rowCell.ToArray().All(Function(cell) cell.Value.IsEmpty) Then

                    bottomRow = rowCell
                    Exit For

                End If

            Next

            If (bottomRow Is Nothing) Then

                ' Extend
                topRow = Nothing
                WB = Nothing
                TargetDefinedRange = Nothing
                CRTargetRange = Nothing
                CRTargetWorksheet = Nothing

                InsertRows(ModelID, RangeName, 1)

                GoTo RestartFrom

                'bottomRow = CRTargetWorksheet.Range.FromLTRB(CRTargetRange.LeftColumnIndex, CRTargetRange.BottomRowIndex + 1, CRTargetRange.RightColumnIndex, CRTargetRange.BottomRowIndex + 1)

                'CRTargetWorksheet.DefinedNames.GetDefinedName(RangeName).Range = CRTargetRange.Resize(CRTargetRange.RowCount + 1, CRTargetRange.ColumnCount)

            End If

            ' Copy
            bottomRow.CopyFrom(topRow, PasteSpecial.Values, True)

        End Sub

    End Class

End Namespace