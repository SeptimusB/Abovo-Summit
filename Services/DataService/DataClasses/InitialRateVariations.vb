Imports System.ComponentModel
Imports System.Text
Imports System
Imports Microsoft.Office.Interop
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraSpreadsheet.Model
Imports System.Collections.ObjectModel
Imports DevExpress.CodeParser
Imports DevExpress.Utils.Extensions
Imports Abovo.AbovoAppCls

Namespace Abovo
    Public Class InitialRateVariations

        Public IsInitialised As Boolean
        Public BActive As Boolean
        Public IsDirty As Boolean
        Public IsError As Boolean
        Public ErrorText As String
        Public NumLines As Short
        Public VariationLines() As InitialRateVariationLine
        Public ActiveRange As String = "IR_NewLetRate"

        'Private ReadOnly ActiveRangeSize As Integer

        Sub New()

            IsInitialised = False
            BActive = False
            IsDirty = False
            NumLines = 0
            IsError = False
            ErrorText = ""
        End Sub

        Public Sub Initialise()

            '            WriteLog("Initilising IRVs")

            '            'On Error GoTo Err_handler

            '            'Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = AbovoBP.WBCoreBP.DefinedNames.GetDefinedName(ActiveRange).Range
            '            Dim CRTargetWorksheet As DevExpress.Spreadsheet.Worksheet = CRTargetRange.Worksheet
            '            Dim clCell As DevExpress.Spreadsheet.Cell
            '            Dim RowRef As Short
            '            Dim i As Short

            '            For i = 0 To CRTargetRange.RowCount - 1

            '                clCell = CRTargetRange(i)

            '                If Len(clCell.DisplayText) > 0 Then

            '                    RowRef = clCell.RowIndex
            '                    BActive = True

            '                    Dim x As Short = AddLine(ConvertToNumZeroed(clCell.DisplayText))

            '                    Dim j As Short

            '                    For j = 1 To AbovoBP.StockSize

            '                        clCell = CRTargetWorksheet(RowRef, 3 + j)

            '                        If Len(clCell.DisplayText) > 0 Then

            '                            VariationLines(x).LineItems(j) = ConvertToNumZeroed(clCell.DisplayText)

            '                        End If

            '                    Next j

            '                End If

            '            Next i
            '            WriteLog("IRVs initialised")

            'Exiter:

            '            IsInitialised = True
            '            Exit Sub

            'Err_handler:

            '            ErrorText = Err.Description
            '            WriteLog("Error: " & ErrorText)
            '            ErrorText = ""
            '            IsError = True

        End Sub

        Public Function AddLine(SetYear As Short) As Short

            'If AbovoBP.GetRangeRows(ActiveRange) = NumLines Then AbovoBP.InsertRows(ActiveRange)
            ReDim Preserve VariationLines(NumLines + 1)
            VariationLines(NumLines + 1) = New InitialRateVariationLine(NumLines + 1) With {.intYear = SetYear}
            BActive = True
            NumLines += 1

            Return NumLines

        End Function
        Public Function CopyLineToNew(LineNumToCopy As Short, SetYear As Short) As Short

            'If AbovoBP.GetRangeRows(ActiveRange) = NumLines Then AbovoBP.InsertRows(ActiveRange)

            ReDim Preserve VariationLines(NumLines + 1)
            VariationLines(NumLines) = New InitialRateVariationLine(NumLines) With {.intYear = SetYear}

            Dim i As Short

            For i = 0 To AbovoBP.StockSize - 1

                VariationLines(NumLines).LineItems(i) = VariationLines(LineNumToCopy).LineItems(i)

            Next i

            NumLines += 1

            Return NumLines

        End Function
        Sub SetValue(VLine As Short, StockItem As Short, value As Single)

            IsDirty = True
            VariationLines(VLine).LineItems(StockItem) = value

        End Sub
        Function GetValue(VLine As Short, StockItem As Short) As Single

            Return VariationLines(VLine).LineItems(StockItem)

        End Function

        Structure InitialRateVariationLine

            Public intYear As Short
            Public intIdentifier As Short
            Public SetCell As String
            Public NumLines As Short
            Public LineItems() As Single
            Sub New(intSetIdentifier As Short)

                Dim i As Short

                intYear = 0
                intIdentifier = intSetIdentifier
                SetCell = ""

                ReDim LineItems(AbovoBP.StockSize)

                For i = 0 To AbovoBP.StockSize - 1

                    LineItems(i) = 0

                Next i

            End Sub

        End Structure

    End Class


End Namespace


