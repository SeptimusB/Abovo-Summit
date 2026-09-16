Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas
Namespace Abovo
    Public Class CustomCalcEngine
        Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService
        Private _DontCalcTDBS As Boolean = False
        Private _TransDBSheetID As Integer = -1
        Private _CheckSheetID As Integer = -1
        Private _ComparisonSheetID As Integer = -1
        Public ModelID As Integer = -1
        Public Property DontCalcTDBS As Boolean
            Get
                Return _DontCalcTDBS
            End Get
            Set(value As Boolean)
                _DontCalcTDBS = value
            End Set
        End Property
        Public Property TransDBSheetID As Integer
            Get
                Return _TransDBSheetID
            End Get
            Set(value As Integer)
                _TransDBSheetID = value
            End Set
        End Property
        Public Property CheckSheetID As Integer
            Get
                Return _CheckSheetID
            End Get
            Set(value As Integer)
                _CheckSheetID = value
            End Set
        End Property
        Public Property ComparisonSheetID As Integer
            Get
                Return _ComparisonSheetID
            End Get
            Set(value As Integer)
                _ComparisonSheetID = value
            End Set
        End Property
        Public Sub New(SetModelID As Integer)
            ModelID = SetModelID
        End Sub

        Public Function OnBeginCalculation() As Boolean Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.OnBeginCalculation
            Return True
        End Function
        Public Sub OnBeginCellCalculation(ByVal args As CellCalculationArgs) Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.OnBeginCellCalculation
            If _DontCalcTDBS AndAlso IsDeferredSheet(args.SheetId) Then
                args.Handled = True
            End If
        End Sub
        Public Function OnBeginCircularReferencesCalculation() As Boolean Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.OnBeginCircularReferencesCalculation
            Return False
        End Function
        Public Sub OnEndCalculation() Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.OnEndCalculation
        End Sub
        Public Sub OnEndCellCalculation(ByVal cellKey As CellKey, ByVal startValue As CellValue, ByVal endValue As CellValue) Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.OnEndCellCalculation
        End Sub
        Public Sub OnEndCircularReferencesCalculation(ByVal cellKeys As IList(Of CellKey)) Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.OnEndCircularReferencesCalculation
        End Sub
        Public Function ShouldMarkupCalculateAlwaysCells() As Boolean Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.ShouldMarkupCalculateAlwaysCells
            Return False
        End Function

        Private Function IsDeferredSheet(ByVal SheetID As Integer) As Boolean
            Return SheetID >= 0 AndAlso
                   (SheetID = _TransDBSheetID OrElse
                    SheetID = _ComparisonSheetID OrElse
                    SheetID = _CheckSheetID)
        End Function

        Public Sub CalculateDeferredWorksheets(ByVal Workbook As IWorkbook)
            If Workbook Is Nothing Then Throw New ArgumentNullException(NameOf(Workbook))

            Dim PreviousDontCalcTDBS As Boolean = _DontCalcTDBS
            _DontCalcTDBS = False
            Try
                CalculateDeferredWorksheet(Workbook, "Transactional DB")
                If TransactionalDBSnapshotManager.HasValidSnapshot(ModelID) Then
                    CalculateDeferredWorksheet(Workbook, "TDB Comparison")
                End If
                CalculateDeferredWorksheet(Workbook, "Check Sheet")
            Finally
                _DontCalcTDBS = PreviousDontCalcTDBS
            End Try
        End Sub

        Private Sub CalculateDeferredWorksheet(ByVal Workbook As IWorkbook,
                                               ByVal WorksheetName As String)
            If Not Workbook.Worksheets.Contains(WorksheetName) Then Return
            Workbook.Worksheets(WorksheetName).Calculate()
        End Sub
    End Class
End Namespace
