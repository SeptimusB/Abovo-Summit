Imports DevExpress.Spreadsheet
Imports Abovo.FileManager

Namespace Abovo
    Public Class ExcelFunctions

        'Public Shared Function VBCountA(ByVal ModelID As Integer, ByVal SearchRange As String) As Integer

        '    Dim WB As IWorkbook = ExcelModels(ModelID).WB

        '    Dim TargetDefinedRange As DevExpress.Spreadsheet.DefinedName = WB.DefinedNames.GetDefinedName(SearchRange)
        '    Dim CRTargetRange As CellRange = TargetDefinedRange.Range
        '    Dim CRTargetWorksheet As DevExpress.Spreadsheet.Worksheet = CRTargetRange.Worksheet



        '    Dim topRow As CellRange = CRTargetWorksheet.Range.FromLTRB(CRTargetRange.LeftColumnIndex, CRTargetRange.TopRowIndex, CRTargetRange.RightColumnIndex, CRTargetRange.TopRowIndex)

        '    ' Find Empty Row
        '    Dim bottomRow As CellRange = Nothing
        '    Dim i As Integer

        '    For i As Integer = CRTargetRange.TopRowIndex To CRTargetRange.TopRowIndex + CRTargetRange.RowCount

        '        Dim rowCell As CellRange = CRTargetWorksheet.Range.FromLTRB(CRTargetRange.LeftColumnIndex, i, CRTargetRange.RightColumnIndex, i)

        '        If rowCell.ToArray().All(Function(cell) cell.Value.IsEmpty) Then Exit For

        '    Next

        '    Return i

        'End Function

    End Class

End Namespace
