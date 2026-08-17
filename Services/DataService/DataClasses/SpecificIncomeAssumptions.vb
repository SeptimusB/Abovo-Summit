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
    Public Class SpecificIncomeAssumptions

        Public BActive As Boolean
        Public NumLines As Integer
        Public VariationLines() As InitialRateVariationLine

        Sub New()

            BActive = False
            NumLines = 0

        End Sub
        Function AddLine(SetYear As Integer) As Integer

            VariationLines(NumLines + 1) = New InitialRateVariationLine(NumLines + 1) With {.intYear = SetYear}
            VariationLines(NumLines + 1).intYear = SetYear
            NumLines += 1
            Return NumLines

        End Function
        Sub CopyLineToNew(LineNumToCopy As Integer, SetYear As Integer)
            VariationLines(NumLines + 1) = New InitialRateVariationLine(NumLines + 1) With {.intYear = SetYear}
            Dim i As Integer

            For i = 1 To AbovoBP.StockSize

                VariationLines(NumLines + 1).LineItems(i) = VariationLines(LineNumToCopy).LineItems(i)

            Next i

            NumLines += 1

        End Sub
        Sub SetValue(VLine As Integer, StockItem As Integer, value As Decimal)

            VariationLines(VLine).LineItems(StockItem) = value

        End Sub
        Function GetValue(VLine As Integer, StockItem As Integer) As Decimal

            Return VariationLines(VLine).LineItems(StockItem)

        End Function



        Structure SummaryOIC

            Public EntryName() As String



        End Structure
        Structure InitialRateVariationLine

            Public intYear As Integer
            Public intIdentifier As Integer
            Public SetCell As String
            Public NumLines As Integer
            Public LineItems() As Decimal
            Sub New(intSetIdentifier As Integer)

                Dim i As Integer

                intYear = 0
                intIdentifier = intSetIdentifier
                SetCell = ""

                For i = 1 To AbovoBP.StockSize

                    Lineitems(i) = New Decimal
                    Lineitems(i) = 0

                Next i

            End Sub

        End Structure


    End Class


End Namespace
