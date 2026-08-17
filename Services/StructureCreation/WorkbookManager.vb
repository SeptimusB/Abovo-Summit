Imports Abovo.AbovoAppCls
Imports DevExpress.Spreadsheet
Imports Abovo.FileManager
Namespace Abovo
    Public Class WorkbookManager

#Region "Worksheet manipulation"
        Public Shared Function GetRangeRows(WB As DevExpress.Spreadsheet.Workbook, RangeName As String) As Integer

            Dim TargetdefinedRange As DevExpress.Spreadsheet.DefinedName = WB.DefinedNames.GetDefinedName(RangeName)
            Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = TargetdefinedRange.Range

            Return CRTargetRange.RowCount

        End Function
        Public Shared Function InsertRows(WB As DevExpress.Spreadsheet.Workbook, TargetNamedRange As String, Optional ByVal RowsToAdd As Integer = 1) As AbovoTransaction

            On Error GoTo Err_Handler_A

            Dim ThisTrans As New AbovoTransaction

            Dim TargetdefinedRange As DevExpress.Spreadsheet.DefinedName = WB.DefinedNames.GetDefinedName(TargetNamedRange)

            Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = TargetdefinedRange.Range
            Dim CRTargetWorksheet As DevExpress.Spreadsheet.Worksheet = CRTargetRange.Worksheet
            Dim IntRangeRows As Integer = CRTargetRange.RowCount
            Dim IntRangeCols As Integer = CRTargetRange.ColumnCount

            CRTargetWorksheet.Unprotect(UnlockPassword)

            On Error Resume Next

            WB.BeginUpdate()

            On Error GoTo Err_Handler_A

            Dim IntBottomRow As Integer = CRTargetRange.BottomRowIndex

            CRTargetWorksheet.Rows.Insert(IntBottomRow)
            CRTargetWorksheet.Rows(IntBottomRow).CopyFrom(CRTargetWorksheet.Rows(IntBottomRow + 1), PasteSpecial.Formulas)
            CRTargetWorksheet.Rows(IntBottomRow).CopyFrom(CRTargetWorksheet.Rows(IntBottomRow + 1), PasteSpecial.Formats)
            CRTargetWorksheet.Rows(IntBottomRow + 1).ClearContents
            CRTargetRange.Resize(IntRangeRows + 1, IntRangeCols)

            On Error Resume Next
            CRTargetWorksheet.Protect(UnlockPassword, WorksheetProtectionPermissions.Default)

            If RowsToAdd > 1 Then

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

    End Class

End Namespace