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
    Public Class BPStockCollection

        Public StockItems(AbovoBP.StockSize) As StockItem
        Public InitialRateVariations As InitialRateVariations
        Public IsDirty As Boolean
        Public Sub New()

            IsDirty = False

            Dim i As Short

            For i = 1 To AbovoBP.StockSize

                StockItems(i) = New StockItem(i)

            Next i

        End Sub
        Public Sub SetIRV()

            InitialRateVariations = New InitialRateVariations()

        End Sub
        Structure StockItem

            Public intIdentifier As Byte

            Public StockDescription As String
            Public OwnedManaged As String
            Public SOCIStockType As String
            Public SOCIRentType As String
            Public CurrentStockNumbers As Int32
            Public PreBPlanStartDateNewBuild As Int32
            Public PreBPlanStartDateDemolitions As Int32
            Public PreBPlanStartDateRTBs As Int32
            Public PreBPlanStartDateOtherDisposals As Int32
            Public NewLettings As Int32
            Public NewLetInitialRate As Double
            Public ExistingStocksCalc As Int32
            Public TotalOpeningStockCalc As Int32

            Sub New(intSetIdentifier As Integer)

                intIdentifier = intSetIdentifier

                StockDescription = ""
                OwnedManaged = ""
                SOCIStockType = ""
                SOCIRentType = ""
                CurrentStockNumbers = 0
                PreBPlanStartDateNewBuild = 0
                PreBPlanStartDateDemolitions = 0
                PreBPlanStartDateRTBs = 0
                PreBPlanStartDateOtherDisposals = 0
                NewLettings = 0
                NewLetInitialRate = 0
                ExistingStocksCalc = 0
                TotalOpeningStockCalc = 0

            End Sub
            Public Sub FUpdateStockTotals()

                ExistingStocksCalc = CurrentStockNumbers + PreBPlanStartDateNewBuild - PreBPlanStartDateDemolitions - PreBPlanStartDateRTBs - PreBPlanStartDateOtherDisposals
                TotalOpeningStockCalc = ExistingStocksCalc + NewLettings

            End Sub

        End Structure



    End Class

End Namespace