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
Imports DevExpress.XtraTreeList.Data
Imports System.Threading
Imports DevExpress.XtraSpreadsheet.Model.History
Imports DevExpress.Xpo


Namespace Abovo

    Public Class AbovoBP ' Class to hold  Abovo Business Plan
#Region "Global definitiions and properties"





        Private Shared InternalUnlockPassword As String = "23_t4qhe"
        Public Property UnlockPassword As String

            Get
                Return InternalUnlockPassword
                Exit Property
            End Get

            Set(ByVal NewUnlockPassword As String)
                InternalUnlockPassword = NewUnlockPassword
            End Set

        End Property
        Private Shared internalFixedStockSize As Integer = 20

        Private InternalBIsSaving As Boolean
        Public ReadOnly Property BIsSaving As Boolean

            Get
                Return InternalBIsSaving
                Exit Property
            End Get

        End Property

        Private InternalCompanyName As String = ""
        Public Property CompanyName As String
            Get
                Return InternalCompanyName
                Exit Property
            End Get
            Set(ByVal NewCompanyName As String)
                InternalCompanyName = NewCompanyName
            End Set
        End Property


        Public Shared Property StockSize As Short
            Get
                Return internalFixedStockSize
                Exit Property
            End Get
            Set(ByVal NewStockSize As Short)
                internalFixedStockSize = NewStockSize
            End Set
        End Property

        Private internalBPState As Integer

        '0 - initialising
        '1 - ready, empty
        '2 - ready, loaded clean
        '3 - ready, dirty
        '4 - saved undirty
        Public Property BPState As Byte
            Get
                Return internalBPState
                Exit Property
            End Get
            Set(ByVal NewBPState As Byte)
                internalBPState = NewBPState
            End Set
        End Property
        Private BUnappliedAssumpInternal As Boolean
        Public Property BUnappliedAssump As Boolean
            Get
                Return BUnappliedAssumpInternal
                Exit Property
            End Get
            Set(ByVal NewBPState As Boolean)
                BUnappliedAssumpInternal = NewBPState
            End Set
        End Property


        Public Shared BPDetails As BPStructure

        Structure BPStructure

            Public intIdentifier As Byte

            Public CompanyName As String
            Public StartDate As Date
            Sub New(intSetIdentifier As Integer)

                intIdentifier = intSetIdentifier

            End Sub

        End Structure


        'Core Properties
        Public Property StrFileName As String
        Public Property FileNameFormat As DocumentFormat

        Public WBCoreBP As DevExpress.Spreadsheet.Workbook

        'Assumptions Classes
        Public Shared Stock As BPStockCollection
        Public Shared SIAs As SpecificIncomeAssumptions
        Public Shared IRVs As InitialRateVariations

        'Output Objects
        Public Shared DSExistingStocksRange As RangeDataSource
        Public Shared DSAnalDataRange As RangeDataSource

#End Region
#Region "Open close save validate transactions"
        Public Sub InitialiseBP()

            internalBPState = 0
            Stock = New BPStockCollection()
            InitiateWorkbook()
            internalBPState = 1
            InternalBIsSaving = False

        End Sub
        Private Sub InitiateWorkbook()

            WBCoreBP = New DevExpress.Spreadsheet.Workbook

            Dim customFunction As New Abovo.PMCostFunction()

            If Not WBCoreBP.Functions.GlobalCustomFunctions.Contains(customFunction.Name) Then

                WBCoreBP.Functions.GlobalCustomFunctions.Add(customFunction)

            End If

            Dim customFunction2 As New Abovo.ResponsiveCostFunction()

            If Not WBCoreBP.Functions.GlobalCustomFunctions.Contains(customFunction2.Name) Then

                WBCoreBP.Functions.GlobalCustomFunctions.Add(customFunction2)

            End If

            WBCoreBP.Options.CalculationEngineType = CalculationEngineType.ChainBased
            WBCoreBP.DocumentSettings.Calculation.EnableMultiThreading = True

        End Sub
        Public Async Function SaveBPBackground() As Task

            Dim cancellationSource As New CancellationTokenSource(TimeSpan.FromSeconds(60))
            Dim cancellationToken As CancellationToken = cancellationSource.Token

            Dim SaveTrans As New AbovoTransaction

            WBCoreBP.Calculate()

            InternalBIsSaving = True

            Try
                Using workbook1 As New Workbook()
                    Using workbook2 As New Workbook()
                        Await WBCoreBP.SaveDocumentAsync(StrFileName, cancellationToken, New Progress(Of Integer)(Sub(progress) Console.WriteLine($"{progress}%")))
                    End Using
                End Using
            Catch e1 As OperationCanceledException
                Console.WriteLine("Cancelled by timeout.")
                Console.ReadLine()
            Finally
                cancellationSource.Dispose()
            End Try




            'Await WBCoreBP.SaveDocumentAsync(StrFileName, New Progress(Of Integer)(Sub(progress) Console.WriteLine($"{progress}%" & " " & Now())))
            InternalBIsSaving = False
            internalBPState = 2



        End Function

        Public Sub SaveBP()

            Dim TimeStart As Date = Now()
            WriteLog("Starting corethread save of " & StrFileName)

            InternalBIsSaving = True
            FileManager.ActiveWB.SaveDocument("SavedAs" & StrFileName, DocumentFormat.Xlsm)
            internalBPState = 2

            InternalBIsSaving = False
            WriteLog("Complete. Time taken: " & (Now() - TimeStart).ToString)
            If IsDev Then MsgBox("Saved. Time taken: " & (Now() - TimeStart).ToString)
        End Sub
        Public Async Sub SaveBPAsync()

            Dim TimeStart As Date = Now()
            WriteLog("Starting async save of " & StrFileName)

            InternalBIsSaving = True
            Await WBCoreBP.SaveDocumentAsync(StrFileName, DocumentFormat.Xlsm)
            internalBPState = 2

            InternalBIsSaving = False
            WriteLog("Complete. Time taken: " & (Now() - TimeStart).ToString)
            If IsDev Then MsgBox("Saved. Time taken: " & (Now() - TimeStart).ToString)
        End Sub
        Public Sub SaveBPAs(StrNewFileName As String, DFFFormat As DocumentFormat)

            WBCoreBP.SaveDocumentAsync(StrNewFileName, DFFFormat)
            internalBPState = 2
            StrFileName = StrNewFileName

        End Sub
        Public Sub CloseBP(CallingForm As Form)

            If internalBPState = 3 Then
                Dim c As New AbovoMessageBox("Current file not saved", MsgBoxStyle.YesNoCancel, CallingForm, "BP Not Saved")
                If c.GetResponse = DialogResult.Cancel Then Exit Sub
            End If
            WBCoreBP.Dispose()
            WBCoreBP = Nothing
            internalBPState = 1

        End Sub
        Public Shared Function ValidateOpenFile() As Boolean

            Return True

        End Function
        Private Async Sub AsycLoadFileHandler()
            Dim task As Task(Of Boolean) = LoadBPAsync(StrFileName)
            Dim result As Boolean = Await task
        End Sub
        Private Async Function LoadBPAsync(FileName As String) As Task(Of Boolean)
            Dim TimeStart As Date = Now()
            WriteLog("Starting async load of " & FileName)
            Await WBCoreBP.LoadDocumentAsync(FileName)
            WriteLog("Complete. Time taken: " & (Now() - TimeStart).ToString)
            If IsDev Then MsgBox("Complete. Time taken: " & (Now() - TimeStart).ToString)
            Return True
        End Function
        Private Sub LoadBPCentral(FileName As String)
            Dim TimeStart As Date = Now()
            WriteLog("Starting corethread load of " & FileName)
            WBCoreBP.LoadDocument(FileName)
            WriteLog("Complete. Time taken: " & (Now() - TimeStart).ToString)
            'If IsDev Then MsgBox("Complete. Time taken: " & (Now() - TimeStart).ToString)
            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        Public Function LoadBP(strPath As String) As AbovoTransaction

            'On Error GoTo Err_Handler_1



            Dim ObjResponse As New AbovoTransaction

            If internalBPState > 1 Then

                ObjResponse.BError = True

            End If

            If IsNothing(WBCoreBP) Then InitiateWorkbook()

            StrFileName = strPath
            LoadBPCentral(StrFileName)

            If ValidateOpenFile() Then

                ObjResponse.BError = False
                internalBPState = 2

            Else

                Beep()
                ObjResponse.BError = True
                ObjResponse.StrResponseMessage = "Sorry, this is not a valid Abovo BP file"
                GoTo Exiter

            End If

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()

            BPDetails = New BPStructure

            PopulateStock()
            CreateDataRanges()


Exiter:

            Return ObjResponse
            Exit Function

Err_Handler_1:

Err_Clean:

            ObjResponse.BError = True
            ObjResponse.IntReturnCode = -1

        End Function

#End Region

#Region "Worksheet manipulation"
        Public Function GetRangeRows(RangeName As String) As Integer

            Dim TargetdefinedRange As DevExpress.Spreadsheet.DefinedName = WBCoreBP.DefinedNames.GetDefinedName(RangeName)
            Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = TargetdefinedRange.Range

            Return CRTargetRange.RowCount

        End Function
        Public Function InsertRows(TargetNamedRange As String, Optional ByVal RowsToAdd As Integer = 1) As AbovoTransaction

            On Error GoTo Err_Handler_A

            Dim ThisTrans As New AbovoTransaction

            Dim TargetdefinedRange As DevExpress.Spreadsheet.DefinedName = WBCoreBP.DefinedNames.GetDefinedName(TargetNamedRange)

            Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = TargetdefinedRange.Range
            Dim CRTargetWorksheet As DevExpress.Spreadsheet.Worksheet = CRTargetRange.Worksheet
            Dim IntRangeRows As Integer = CRTargetRange.RowCount
            Dim IntRangeCols As Integer = CRTargetRange.ColumnCount

            CRTargetWorksheet.Unprotect(InternalUnlockPassword)

            On Error Resume Next



            On Error GoTo Err_Handler_A

            Dim IntBottomRow As Integer = CRTargetRange.BottomRowIndex

            CRTargetWorksheet.Rows.Insert(IntBottomRow)
            CRTargetWorksheet.Rows(IntBottomRow).CopyFrom(CRTargetWorksheet.Rows(IntBottomRow + 1), PasteSpecial.Formulas)
            CRTargetWorksheet.Rows(IntBottomRow).CopyFrom(CRTargetWorksheet.Rows(IntBottomRow + 1), PasteSpecial.Formats)
            CRTargetWorksheet.Rows(IntBottomRow + 1).ClearContents
            CRTargetRange.Resize(IntRangeRows + 1, IntRangeCols)

            On Error Resume Next
            CRTargetWorksheet.Protect(InternalUnlockPassword, WorksheetProtectionPermissions.Default)

            If RowsToAdd > 1 Then

            End If



            ThisTrans.BError = False

            Return ThisTrans

            Exit Function

Err_Handler_A:

        End Function
        Public Function InsertColumn(TargetNamedRange As String, Optional ByVal ColssToAdd As Integer = 1) As AbovoTransaction

            On Error GoTo Err_Handler_A

            Dim ThisTrans As New AbovoTransaction

            Dim TargetdefinedRange As DevExpress.Spreadsheet.DefinedName = WBCoreBP.DefinedNames.GetDefinedName(TargetNamedRange)

            Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = TargetdefinedRange.Range
            Dim CRTargetWorksheet As DevExpress.Spreadsheet.Worksheet = CRTargetRange.Worksheet
            Dim IntRangeRows As Integer = CRTargetRange.RowCount
            Dim IntRangeCols As Integer = CRTargetRange.ColumnCount

            On Error Resume Next

            CRTargetWorksheet.Unprotect(InternalUnlockPassword)

            On Error GoTo Err_Handler_A



            Dim IntRightCol As Integer = CRTargetRange.RightColumnIndex

            CRTargetWorksheet.Columns.Insert(IntRightCol)
            CRTargetWorksheet.Columns(IntRightCol).CopyFrom(CRTargetWorksheet.Columns(IntRightCol + 1), PasteSpecial.Formulas)
            CRTargetWorksheet.Columns(IntRightCol).CopyFrom(CRTargetWorksheet.Columns(IntRightCol + 1), PasteSpecial.Formats)
            CRTargetWorksheet.Columns(IntRightCol + 1).ClearContents
            CRTargetRange.Resize(IntRangeRows, IntRangeCols + 1)

            On Error Resume Next

            CRTargetWorksheet.Protect(InternalUnlockPassword, WorksheetProtectionPermissions.Default)

            On Error GoTo Err_Handler_A

            If ColssToAdd > 1 Then

            End If



            ThisTrans.BError = False

            Return ThisTrans

            Exit Function

Err_Handler_A:

        End Function
#End Region

        'Sub WriteCell(StrSheet As String, StrRef As String, varValue As VariantType)

        'End Sub
        Public Sub WriteStock()

            WBCoreBP.DocumentSettings.R1C1ReferenceStyle = True

            Dim wsStock As DevExpress.Spreadsheet.Worksheet

            Dim clCell As DevExpress.Spreadsheet.Cell


            Dim i, sRef As Integer


            'wsStock.Unprotect("")


            wsStock = WBCoreBP.Worksheets("Stock Assumptions")

            'Try

            For i = 0 To internalFixedStockSize - 1

                sRef = 4 + i

                clCell = wsStock(4, sRef)
                If clCell.Value <> Stock.StockItems(i).StockDescription Then clCell.Value = Stock.StockItems(i).StockDescription

                clCell = wsStock(5, sRef)
                If clCell.Value <> Stock.StockItems(i).OwnedManaged Then clCell.Value = Stock.StockItems(i).OwnedManaged

                clCell = wsStock(6, sRef)
                If clCell.Value <> Stock.StockItems(i).SOCIStockType Then clCell.Value = Stock.StockItems(i).SOCIStockType

                clCell = wsStock(7, sRef)
                If clCell.Value <> Stock.StockItems(i).SOCIRentType Then clCell.Value = Stock.StockItems(i).SOCIRentType

                clCell = wsStock(16, sRef)
                If clCell.Value <> Stock.StockItems(i).CurrentStockNumbers Then clCell.Value = Stock.StockItems(i).CurrentStockNumbers

                clCell = wsStock(19, sRef)
                If clCell.Value <> Stock.StockItems(i).PreBPlanStartDateNewBuild Then clCell.Value = Stock.StockItems(i).PreBPlanStartDateNewBuild

                clCell = wsStock(21, sRef)
                If clCell.Value <> Stock.StockItems(i).PreBPlanStartDateDemolitions Then clCell.Value = Stock.StockItems(i).PreBPlanStartDateDemolitions

                clCell = wsStock(22, sRef)
                If clCell.Value <> Stock.StockItems(i).PreBPlanStartDateRTBs Then clCell.Value = Stock.StockItems(i).PreBPlanStartDateRTBs

                clCell = wsStock(27, sRef)
                If clCell.Value <> Stock.StockItems(i).NewLettings Then clCell.Value = Stock.StockItems(i).NewLettings

                clCell = wsStock(34, sRef)
                If clCell.Value <> Stock.StockItems(i).NewLetInitialRate Then clCell.Value = Stock.StockItems(i).NewLetInitialRate

            Next i

            'Catch ex As Exception



            'Finally

            'CalcBP(2)
            'End Try


        End Sub
        Public Sub ListStock()

            For i = 1 To StockSize

                MsgBox(AbovoBP.Stock.StockItems(i).StockDescription)

            Next i

        End Sub


    End Class


End Namespace

