Option Strict On

Imports System.IO
Imports System.Threading
Imports DevExpress.Spreadsheet

Namespace Abovo.WorkbookEngines
    Friend NotInheritable Class DevExpressCalculationBackend
        Implements IWorkbookCalculationBackend

        Private ReadOnly threadId As Integer = Thread.CurrentThread.ManagedThreadId
        Private book As Workbook
        Public ReadOnly Property Name As String Implements IWorkbookCalculationBackend.Name
            Get
                Return "DevExpress"
            End Get
        End Property
        Public ReadOnly Property Version As String Implements IWorkbookCalculationBackend.Version
            Get
                Return GetType(Workbook).Assembly.GetName().Version.ToString()
            End Get
        End Property
        Private Sub RequireOwner()
            If Thread.CurrentThread.ManagedThreadId <> threadId Then Throw New InvalidOperationException("Workbook accessed outside its owner thread.")
        End Sub
        Public Sub OpenReadOnly(path As String, options As WorkbookEngineOptions) Implements IWorkbookCalculationBackend.OpenReadOnly
            RequireOwner()
            book = New Workbook()
            book.Options.CalculationMode = WorkbookCalculationMode.Manual
            book.Options.CalculationEngineType = CalculationEngineType.Recursive
            ' Workbook-local functions do not mutate the application's global
            ' function registry while another workbook may be calculating.
            If options.RequireSummitFunctions Then
                book.Functions.CustomFunctions.Add(New PMCostFunction())
                book.Functions.CustomFunctions.Add(New ResponsiveCostFunction())
            End If
            Dim format = If(String.Equals(IO.Path.GetExtension(path), ".xlsb", StringComparison.OrdinalIgnoreCase), DocumentFormat.Xlsb,
                            If(String.Equals(IO.Path.GetExtension(path), ".xlsm", StringComparison.OrdinalIgnoreCase), DocumentFormat.Xlsm, DocumentFormat.Xlsx))
            Using stream = File.OpenRead(path)
                If Not book.LoadDocument(stream, format) Then Throw New InvalidDataException("DevExpress could not open the workbook.")
            End Using
            book.DocumentSettings.Calculation.Mode = CalculationMode.Manual
        End Sub
        Public Sub Calculate(kind As WorkbookCalculationKind) Implements IWorkbookCalculationBackend.Calculate
            RequireOwner()
            Select Case kind
                Case WorkbookCalculationKind.Incremental : book.Calculate()
                Case WorkbookCalculationKind.Full : book.CalculateFull()
                Case WorkbookCalculationKind.Rebuild : book.CalculateFullRebuild()
                Case Else : Throw New ArgumentOutOfRangeException(NameOf(kind))
            End Select
        End Sub
        Public Function Read(area As WorkbookReadArea) As WorkbookValueBlock Implements IWorkbookCalculationBackend.Read
            RequireOwner()
            If Not book.Worksheets.Contains(area.Worksheet) Then Throw New InvalidDataException("Worksheet not found: " & area.Worksheet)
            Dim sheet = book.Worksheets(area.Worksheet)
            Dim data(area.Rows - 1, area.Columns - 1) As Object
            For r As Integer = 0 To area.Rows - 1
                For c As Integer = 0 To area.Columns - 1
                    Dim value = sheet.Cells(area.Row + r, area.Column + c).Value
                    If value.IsError Then
                        data(r, c) = New WorkbookCellError(value.ToString())
                    ElseIf value.IsNumeric Then
                        data(r, c) = value.NumericValue
                    ElseIf value.IsBoolean Then
                        data(r, c) = value.BooleanValue
                    ElseIf value.IsText Then
                        data(r, c) = value.TextValue
                    End If
                Next
            Next
            Return New WorkbookValueBlock(area, data)
        End Function
        Public Sub Dispose() Implements IDisposable.Dispose
            RequireOwner()
            If book IsNot Nothing Then book.Dispose()
            book = Nothing
        End Sub
    End Class
End Namespace
