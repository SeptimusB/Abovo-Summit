Option Strict On

Imports System.Collections.ObjectModel
Imports System.Drawing
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Export

Namespace Abovo.WorkbookEngines
    ' Detached workbook-owned appearance. No native cells, COM objects, fonts
    ' or mutable buffers cross the calculation owner's boundary.
    Public NotInheritable Class WorkbookCellAppearance
        Public ReadOnly Property NumberFormat As String
        Public ReadOnly Property Background As Color
        Public ReadOnly Property Foreground As Color
        Public ReadOnly Property FontName As String
        Public ReadOnly Property FontSize As Double
        Public ReadOnly Property FontStyle As FontStyle
        Public ReadOnly Property Horizontal As SpreadsheetHorizontalAlignment
        Public ReadOnly Property Vertical As SpreadsheetVerticalAlignment
        Public ReadOnly Property WrapText As Boolean
        Public ReadOnly Property Indent As Integer
        Public ReadOnly Property Rotation As Integer
        Public ReadOnly Property Locked As Boolean
        Public ReadOnly Property SolidFill As Boolean
        Public Sub New(numberFormat As String, background As Color, foreground As Color, fontName As String,
                       fontSize As Double, fontStyle As FontStyle, horizontal As SpreadsheetHorizontalAlignment,
                       vertical As SpreadsheetVerticalAlignment, wrapText As Boolean, indent As Integer,
                       rotation As Integer, locked As Boolean, solidFill As Boolean)
            If String.IsNullOrWhiteSpace(fontName) OrElse fontSize <= 0 OrElse Double.IsNaN(fontSize) OrElse Double.IsInfinity(fontSize) Then Throw New ArgumentException("Invalid workbook font.")
            Me.NumberFormat = If(numberFormat, "General") : Me.Background = background : Me.Foreground = foreground
            Me.FontName = fontName : Me.FontSize = fontSize : Me.FontStyle = fontStyle
            Me.Horizontal = horizontal : Me.Vertical = vertical : Me.WrapText = wrapText
            Me.Indent = indent : Me.Rotation = rotation : Me.Locked = locked : Me.SolidFill = solidFill
        End Sub
    End Class

    Public NotInheritable Class WorkbookPresentationCell
        Public ReadOnly Property Appearance As WorkbookCellAppearance
        Public ReadOnly Property Text As String
        Public Sub New(appearance As WorkbookCellAppearance, text As String)
            If appearance Is Nothing Then Throw New ArgumentNullException(NameOf(appearance))
            Me.Appearance = appearance : Me.Text = If(text, "")
        End Sub
    End Class

    Public NotInheritable Class WorkbookPresentationBlock
        Private ReadOnly cells As WorkbookPresentationCell(,)
        Public ReadOnly Property Area As WorkbookReadArea
        Public Sub New(area As WorkbookReadArea, cells As WorkbookPresentationCell(,))
            If area Is Nothing OrElse cells Is Nothing OrElse cells.GetLowerBound(0) <> 0 OrElse cells.GetLowerBound(1) <> 0 OrElse
                cells.GetLength(0) <> area.Rows OrElse cells.GetLength(1) <> area.Columns Then Throw New ArgumentException("Presentation rectangle differs from request.")
            For Each cell In cells
                If cell Is Nothing Then Throw New ArgumentException("A presentation cell is missing.")
            Next
            Me.Area = area : Me.cells = DirectCast(cells.Clone(), WorkbookPresentationCell(,))
        End Sub
        Public Function CellAt(row As Integer, column As Integer) As WorkbookPresentationCell
            Return cells(row, column)
        End Function
    End Class

    Public Interface IWorkbookPresentationBackend
        Inherits IWorkbookCalculationBackend
        Function ReadPresentation(values As WorkbookValueBlock) As WorkbookPresentationBlock
    End Interface

    ' Supported DevExpress number formatting only; no business-plan formulas
    ' are loaded, evaluated or written in this tiny formatting workbook.
    Friend NotInheritable Class WorkbookDisplayFormatter
        Implements IDisposable
        Private ReadOnly formats As New Dictionary(Of String, CellValueToStringConverter)(StringComparer.Ordinal)
        Private ReadOnly book As New Workbook()
        Friend Sub New(date1904 As Boolean)
            book.Options.CalculationMode = WorkbookCalculationMode.Manual
            book.Options.Culture = Globalization.CultureInfo.CurrentCulture
            book.DocumentSettings.Calculation.Use1904DateSystem = date1904
        End Sub
        Friend Function Format(value As Object, numberFormat As String) As String
            If value Is Nothing Then Return ""
            If TypeOf value Is WorkbookCellError Then Return DirectCast(value, WorkbookCellError).Text
            If TypeOf value Is String Then Return CStr(value)
            If TypeOf value Is Boolean Then Return If(CBool(value), "TRUE", "FALSE")
            Dim converter As CellValueToStringConverter = Nothing
            Dim key = If(numberFormat, "General")
            If Not formats.TryGetValue(key, converter) Then
                converter = New CellValueToStringConverter() With {.SkipErrorValues = False, .PreferredCulture = book.Options.Culture}
                converter.SetPreferredNumberFormat(book, key)
                If formats.Count >= 2048 Then formats.Clear()
                formats.Add(key, converter)
            End If
            Dim text As Object = Nothing
            converter.Convert(book.Worksheets(0).Cells(0, 0), CellValue.FromObject(value), GetType(String), text)
            If text Is Nothing Then Throw New InvalidOperationException("The workbook number format could not be rendered.")
            Return CStr(text)
        End Function
        Public Sub Dispose() Implements IDisposable.Dispose
            formats.Clear() : book.Dispose()
        End Sub
    End Class

    Partial Public NotInheritable Class WorkbookCalculationSession
        Private calculationGeneration As Long
        ' Invalidate an earlier render even when inputs did not change (for
        ' example NOW(), volatile VBA or two queued calculation requests).
        Private Sub CalculateNative(kind As WorkbookCalculationKind)
            SyncLock gate
                RequireAvailable()
                calculationGeneration += 1
            End SyncLock
            backend.Calculate(kind)
        End Sub
        Private Function MakeResult(revision As Long, calculationMs As Long, transferMs As Long,
                                    values As IList(Of WorkbookValueBlock),
                                    Optional presentation As IList(Of WorkbookPresentationBlock) = Nothing) As WorkbookCalculationResult
            SyncLock gate
                RequireRevision(revision)
                Return New WorkbookCalculationResult(SessionId, revision, SourceHash, EngineName, calculationMs, transferMs,
                                                     values, calculationGeneration, presentation)
            End SyncLock
        End Function
        Private Function ReadPresentationOnOwner(values As IList(Of WorkbookValueBlock)) As List(Of WorkbookPresentationBlock)
            Dim reader = TryCast(backend, IWorkbookPresentationBackend)
            If reader Is Nothing Then Throw New NotSupportedException("The selected engine cannot supply authoritative appearance.")
            Dim result As New List(Of WorkbookPresentationBlock)()
            For Each block In values
                Dim display = reader.ReadPresentation(block)
                If display Is Nothing OrElse Not SameArea(display.Area, block.Area) Then Throw New IO.InvalidDataException("Engine appearance differs from the requested rectangle.")
                result.Add(display)
            Next
            Return result
        End Function
    End Class
End Namespace
