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
Imports Abovo.AbovoBP
Imports Abovo.DataObject
Imports DevExpress.XtraRichEdit.Model
Imports DevExpress.XtraLayout.Customization
Imports DevExpress.UIAutomation


Namespace Abovo
    Public Class ManagementCostAssumptions

        Public BActive As Boolean
        Public NumLines As Integer
        Public VariationLines() As InitialRateVariationLine
        Public ObjData As DataObject

        Sub New()

            BActive = False
            NumLines = 0

        End Sub
        Sub Initialise()

            Dim Trans As New AbovoTransaction

            Dim MergeList As New List(Of Integer)

            ObjData = New DataObject("ManagementCosts")
            ' SystemLog("Created object - index is " & ObjData.Presentation.DSIndex)

            '''''''''''''''''''''''''''

            Dim Titles() As String = {"Summary Cost Category"}
            Dim TitleList As New List(Of String)(Titles)

            Trans = ObjData.AddStandardNamedRange("CostCats", "S", TitleList)

            Dim Titles2() As String = {"Description"}
            TitleList = New List(Of String)(Titles2)

            Dim Trans2 As AbovoTransaction = ObjData.AddStandardNamedRange("IR_Oneoff_Cost_Ass", "SD", TitleList, 1)
            MergeList.Add(Trans2.IntegerReturn)

            Dim Titles3() As String = {"Year", "Cost Category", "Staff/Other", "Amount"}
            TitleList = New List(Of String)(Titles3)

            Dim Trans3 As AbovoTransaction = ObjData.AddStandardNamedRange("Rep_MgtC_09", "YDSDSDN", TitleList, 2)

            MergeList.Add(Trans3.IntegerReturn)

            'ManagementCostData.GetStatus()

            Dim Trans4 As AbovoTransaction = ObjData.UnionAcross(MergeList, "One Off Other Spend")

            ''''''''''''''''''''''''''''''''''''''''

            Dim MergeList2 As New List(Of Integer)
            Dim Titles4() As String = {"Category Description"}
            TitleList = New List(Of String)(Titles4)

            Dim Trans5 As AbovoTransaction = ObjData.AddHorizontalHeaderDataRange("Services", "S", TitleList)

            MergeList2.Add(Trans5.IntegerReturn)

            Dim Titles5() As String = {"Summary Category"}
            TitleList = New List(Of String)(Titles5)

            Dim Trans6 As AbovoTransaction = ObjData.AddHorizontalHeaderDataRange("Rep_MgtC_01", "S", TitleList)
            MergeList2.Add(Trans6.IntegerReturn)

            Dim Titles6() As String = {"SOCI Expense"}
            TitleList = New List(Of String)(Titles6)

            Dim Trans7 As AbovoTransaction = ObjData.AddHorizontalHeaderDataRange("Rep_MgtC_01a", "S", TitleList)
            MergeList2.Add(Trans7.IntegerReturn)

            Dim Trans8 As AbovoTransaction = ObjData.UnionAcross(MergeList2, "Management Cost Categories")

            ''''''''''''''''''''''''''''''




            Dim Titles7() As String = {"Year"}
            TitleList = New List(Of String)(Titles7)

            Dim Trans9 As AbovoTransaction = ObjData.AddStandardNamedRange("IR_Staff_Cost_Ass", "Y", TitleList)

            Dim Trans10 As AbovoTransaction = ObjData.AddPivottedDataRangeByDefinedColums("Rep_MgtC_06", Trans9.IntegerReturn, "N", 1)

            Dim MergeList3 As New List(Of Integer)
            MergeList3.Add(Trans8.IntegerReturn)
            MergeList3.Add(Trans10.IntegerReturn)

            Dim Trans11 As AbovoTransaction = ObjData.UnionAcross(MergeList3, "StaffCostsByYear")

            ''''''''''''''

            Dim Titles8() As String = {"Year"}
            TitleList = New List(Of String)(Titles8)

            Dim Trans12 As AbovoTransaction = ObjData.AddStandardNamedRange("IR_Other_Cost_Ass", "Y", TitleList)

            Dim Trans13 As AbovoTransaction = ObjData.AddPivottedDataRangeByDefinedColums("Rep_MgtC_07", Trans12.IntegerReturn, "N", 1)

            Dim MergeList4 As New List(Of Integer)
            MergeList4.Add(Trans8.IntegerReturn)
            MergeList4.Add(Trans13.IntegerReturn)

            Dim Trans14 As AbovoTransaction = ObjData.UnionAcross(MergeList4, "OtherCostsByYear")

            ''''''''''''''''''''''''''''''

            Dim Titles9() As String = {"Staff Cost %", "Other Cost %"}
            TitleList = New List(Of String)(Titles9)
            Dim Trans15 As AbovoTransaction = ObjData.AddPivottedDataRangeByTitleList("Rep_MgtC_04", "PP", TitleList)

            Dim Titles10() As String = {"Variable Cost driver"}
            TitleList = New List(Of String)(Titles10)
            Dim Trans16 As AbovoTransaction = ObjData.AddPivottedDataRangeByTitleList("Rep_MgtC_05", "S", TitleList)

            'DataObject.Presentation.AddPresentation(Trans8.IntegerReturn, "Grid", "Management Cost Categories")
            'ObjData.Presentation.AddPresentation(Trans11.IntegerReturn, "Grid", "Staff Costs")
            'ObjData.Presentation.AddPresentation(Trans14.IntegerReturn, "Grid", "Other Costs")
            'ObjData.Presentation.AddPresentation(Trans16.IntegerReturn, "Grid", "Variable Cost driver")
            'ObjData.GetStatus()

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

                    LineItems(i) = New Decimal
                    LineItems(i) = 0

                Next i

            End Sub

        End Structure


    End Class


End Namespace

