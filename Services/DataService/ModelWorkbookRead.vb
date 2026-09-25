Option Strict On

Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports Abovo.WorkbookEngines
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas
Imports DevExpress.Spreadsheet.Functions

Namespace Abovo
    ' Read facade for existing worksheet-backed controls. Without a selected
    ' engine owner, values are the original native DevExpress reads. With one,
    ' cache misses are read from that same calculated owner, never from WB's
    ' stale formula cache. Original cells remain the structural/address map.
    Friend NotInheritable Class ModelEngineView
        Private Shared ReadOnly views As New ConditionalWeakTable(Of IWorkbook, ModelEngineView)()
        Private ReadOnly ownerThread As Integer = Threading.Thread.CurrentThread.ManagedThreadId
        Private session As WorkbookCalculationSession
        Private anchor As WorkbookCalculationResult
        Private ReadOnly blocks As New List(Of WorkbookCalculationResult)()
        Private Sub New(session As WorkbookCalculationSession, initial As WorkbookCalculationResult)
            Me.session = session
            Publish(initial)
        End Sub
        Friend Shared Sub Bind(book As IWorkbook, session As WorkbookCalculationSession, initial As WorkbookCalculationResult)
            Dim view As New ModelEngineView(session, initial)
            views.Add(book, view)
            ' Workbook is a public wrapper; Worksheet.Workbook can expose its
            ' distinct document implementation. Bind that exact native identity
            ' too, rather than assuming reference equality across the API.
            For Each sheet In book.Worksheets
                Dim existing As ModelEngineView = Nothing
                If Not views.TryGetValue(sheet.Workbook, existing) Then
                    views.Add(sheet.Workbook, view)
                ElseIf existing IsNot view Then
                    Throw New InvalidOperationException("Workbook already has another result owner.")
                End If
            Next
        End Sub
        Friend Shared Function Find(cell As Cell) As ModelEngineView
            Dim view As ModelEngineView = Nothing
            views.TryGetValue(cell.Worksheet.Workbook, view)
            Return view
        End Function
        Friend Shared Sub Publish(book As IWorkbook, result As WorkbookCalculationResult)
            Dim view As ModelEngineView = Nothing
            If views.TryGetValue(book, view) Then view.Publish(result)
        End Sub
        Friend Shared Sub ReplaceOwner(book As IWorkbook, previous As WorkbookCalculationSession,
                                       replacement As WorkbookCalculationSession, result As WorkbookCalculationResult)
            Dim view As ModelEngineView = Nothing
            If Not views.TryGetValue(book, view) Then Throw New InvalidOperationException("The model has no bound result owner.")
            view.RequireOwner()
            If view.session IsNot previous OrElse replacement Is Nothing OrElse Not replacement.IsCurrent(result) Then
                Throw New InvalidOperationException("The saved model cannot adopt this calculation owner.")
            End If
            If previous.NativeCleanupCompletion Is Nothing OrElse previous.NativeCleanupCompletion.Status <> Threading.Tasks.TaskStatus.RanToCompletion Then
                Throw New InvalidOperationException("The previous calculation owner has not closed successfully.")
            End If
            view.session = replacement
            view.Publish(result)
        End Sub
        Private Sub RequireOwner()
            If Threading.Thread.CurrentThread.ManagedThreadId <> ownerThread Then Throw New InvalidOperationException("Model display accessed outside its owner thread.")
        End Sub
        Private Sub Publish(result As WorkbookCalculationResult)
            RequireOwner()
            blocks.Clear() : anchor = Nothing
            If result Is Nothing Then Return
            If Not session.IsCurrent(result) Then Throw New InvalidOperationException("Cannot display stale engine results.")
            anchor = result : blocks.Add(result)
        End Sub
        Private Shared Function Contains(area As WorkbookReadArea, cell As Cell) As Boolean
            Return area.Worksheet.Equals(cell.Worksheet.Name, StringComparison.OrdinalIgnoreCase) AndAlso
                cell.RowIndex >= area.Row AndAlso cell.RowIndex < area.Row + area.Rows AndAlso
                cell.ColumnIndex >= area.Column AndAlso cell.ColumnIndex < area.Column + area.Columns
        End Function
        Private Function ResultFor(cell As Cell, presentation As Boolean) As WorkbookCalculationResult
            RequireOwner()
            If Not session.IsCurrent(anchor) Then Throw New InvalidOperationException("Model results are pending. Refresh before displaying or editing cells.")
            For Each result In blocks
                If result.Blocks.Any(Function(block) Contains(block.Area, cell)) AndAlso
                    (Not presentation OrElse result.Presentation.Any(Function(block) Contains(block.Area, cell))) Then Return result
            Next
            ' Small spatial batches amortise native value/appearance calls while
            ' bounding memory. Headers and sparse inputs do not read whole sheets.
            Dim row = (cell.RowIndex \ 8) * 8, column = (cell.ColumnIndex \ 4) * 4
            Dim area As New WorkbookReadArea(cell.Worksheet.Name, row, column, 8, 4)
            Dim fetched = session.ReadCurrentAsync(anchor, {area}, includePresentation:=presentation).GetAwaiter().GetResult()
            If Not session.IsCurrent(fetched) Then Throw New InvalidOperationException("Model display was superseded during refresh.")
            If blocks.Count >= 2048 Then blocks.RemoveAt(0)
            blocks.Add(fetched)
            Return fetched
        End Function
        Friend Function Value(cell As Cell) As CellValue
            Dim result = ResultFor(cell, False)
            Dim block = result.Blocks.First(Function(item) Contains(item.Area, cell))
            Dim raw = block.ValueAt(cell.RowIndex - block.Area.Row, cell.ColumnIndex - block.Area.Column)
            Return ToCellValue(raw)
        End Function
        Friend Shared Function ToCellValue(raw As Object) As CellValue
            If raw Is Nothing Then Return CellValue.Empty
            If TypeOf raw Is WorkbookCellError Then
                Select Case DirectCast(raw, WorkbookCellError).Text
                    Case "#DIV/0!" : Return CellValue.ErrorDivisionByZero
                    Case "#N/A" : Return CellValue.ErrorValueNotAvailable
                    Case "#NAME?" : Return CellValue.ErrorName
                    Case "#NULL!" : Return CellValue.ErrorNullIntersection
                    Case "#NUM!" : Return CellValue.ErrorNumber
                    Case "#REF!" : Return CellValue.ErrorReference
                    Case Else : Return CellValue.ErrorInvalidValueInFunction ' DisplayText retains the exact modern/unknown error.
                End Select
            End If
            Return CellValue.FromObject(raw)
        End Function
        Friend Function Presentation(cell As Cell) As WorkbookPresentationCell
            Dim result = ResultFor(cell, True)
            Dim block = result.Presentation.First(Function(item) Contains(item.Area, cell))
            Return block.CellAt(cell.RowIndex - block.Area.Row, cell.ColumnIndex - block.Area.Column)
        End Function
        Friend Function CaptureRange(range As CellRange) As WorkbookCalculationResult
            RequireOwner()
            If Not session.IsCurrent(anchor) Then Throw New InvalidOperationException("Model results are pending. Refresh before binding a range.")
            If CLng(range.RowCount) * range.ColumnCount > 500000 Then Throw New ArgumentException("The display range exceeds the bounded transfer limit.", NameOf(range))
            Dim areas As New List(Of WorkbookReadArea)()
            Dim stepRows = Math.Max(1, 100000 \ range.ColumnCount)
            For offset = 0 To range.RowCount - 1 Step stepRows
                areas.Add(New WorkbookReadArea(range.Worksheet.Name, range.TopRowIndex + offset,
                    range.LeftColumnIndex, Math.Min(stepRows, range.RowCount - offset), range.ColumnCount))
            Next
            Dim result = session.ReadCurrentAsync(anchor, areas).GetAwaiter().GetResult()
            If Not session.IsCurrent(result) Then Throw New InvalidOperationException("The display range was superseded during transfer.")
            If blocks.Count >= 2048 Then blocks.RemoveAt(0)
            blocks.Add(result)
            Return result
        End Function
        Friend Function IsCurrent(result As WorkbookCalculationResult) As Boolean
            RequireOwner()
            Return HasCurrentValues AndAlso session.IsCurrent(result)
        End Function
        Friend ReadOnly Property HasCurrentValues As Boolean
            Get
                RequireOwner()
                Return session.IsCurrent(anchor)
            End Get
        End Property
        Friend Function Evaluate(cell As Cell, formula As String) As CellValue
            RequireOwner()
            Dim raw = session.EvaluateCurrentAsync(anchor, New WorkbookReadArea(cell.Worksheet.Name,
                cell.RowIndex, cell.ColumnIndex, 1, 1), formula).GetAwaiter().GetResult()
            Return ToCellValue(raw)
        End Function
    End Class

    Public Module ModelWorkbookRead
        <Extension()> Public Function ModelResultsAvailable(cell As Cell) As Boolean
            Dim view = ModelEngineView.Find(cell)
            Return view Is Nothing OrElse view.HasCurrentValues
        End Function
        ' Painting may re-enter while an STA wait pumps WM_PAINT. Strict model
        ' reads still fail closed; display callbacks alone show a pending marker.
        <Extension()> Public Function ModelPaintText(cell As Cell) As String
            If Not cell.ModelResultsAvailable() Then Return "…"
            Return cell.ModelDisplayText()
        End Function
        <Extension()> Public Function ModelEvaluate(cell As Cell, formula As String,
                                                   Optional style As ReferenceStyle = ReferenceStyle.A1) As ParameterValue
            Dim context As New ExpressionContext(cell.ColumnIndex, cell.RowIndex, cell.Worksheet,
                Globalization.CultureInfo.InvariantCulture, style, ExpressionStyle.Normal)
            Dim engine = cell.Worksheet.Workbook.FormulaEngine
            Dim view = ModelEngineView.Find(cell)
            If view Is Nothing Then Return engine.Evaluate(formula, context)
            ' Parsing/rebasing is symbolic only; evaluate against the selected
            ' owner, never the presentation document's old cached precedents.
            Dim parsed = engine.Parse(formula, context)
            context.ReferenceStyle = ReferenceStyle.A1
            Dim absoluteFormula = parsed.ToString(context)
            If Not absoluteFormula.StartsWith("=", StringComparison.Ordinal) Then absoluteFormula = "=" & absoluteFormula
            Return view.Evaluate(cell, absoluteFormula)
        End Function
        <Extension()> Public Function HasModelEngineView(cell As Cell) As Boolean
            Return ModelEngineView.Find(cell) IsNot Nothing
        End Function
        <Extension()> Public Function ModelValue(cell As Cell) As CellValue
            Dim view = ModelEngineView.Find(cell)
            Return If(view Is Nothing, cell.Value, view.Value(cell))
        End Function
        <Extension()> Public Function ModelDisplayText(cell As Cell) As String
            Dim view = ModelEngineView.Find(cell)
            Return If(view Is Nothing, cell.DisplayText, view.Presentation(cell).Text)
        End Function
        <Extension()> Public Function ModelDateValue(cell As Cell) As DateTime
            Dim value = cell.ModelValue()
            If Not cell.HasModelEngineView() Then Return value.DateTimeValue
            Dim serial As ValueObject = value
            Return serial.GetDateTimeValue(cell.Worksheet.Workbook.DocumentSettings.Calculation.Use1904DateSystem)
        End Function
        <Extension()> Public Function ModelFont(cell As Cell) As WorkbookReadFont
            Dim view = ModelEngineView.Find(cell)
            If view IsNot Nothing Then Return New WorkbookReadFont(view.Presentation(cell).Appearance)
            Return New WorkbookReadFont(cell.Font)
        End Function
        <Extension()> Public Function ModelFill(cell As Cell) As WorkbookReadFill
            Dim view = ModelEngineView.Find(cell)
            If view IsNot Nothing Then
                Dim a = view.Presentation(cell).Appearance
                Return New WorkbookReadFill(a.Background, If(a.SolidFill, PatternType.Solid, PatternType.Gray125))
            End If
            Return New WorkbookReadFill(cell.Fill.BackgroundColor, cell.Fill.PatternType)
        End Function
        <Extension()> Public Function ModelHorizontalAlignment(cell As Cell) As SpreadsheetHorizontalAlignment
            Dim view = ModelEngineView.Find(cell)
            Return If(view Is Nothing, cell.Alignment.Horizontal, view.Presentation(cell).Appearance.Horizontal)
        End Function
        <Extension()> Public Function ModelVerticalAlignment(cell As Cell) As SpreadsheetVerticalAlignment
            Dim view = ModelEngineView.Find(cell)
            Return If(view Is Nothing, cell.Alignment.Vertical, view.Presentation(cell).Appearance.Vertical)
        End Function
    End Module

    Public Structure WorkbookReadFont
        Public ReadOnly Name As String
        Public ReadOnly Size As Double
        Public ReadOnly Color As Color
        Public ReadOnly Bold As Boolean
        Public ReadOnly Italic As Boolean
        Public ReadOnly Strikethrough As Boolean
        Public ReadOnly UnderlineType As UnderlineType
        Friend Sub New(font As SpreadsheetFont)
            Name = font.Name : Size = font.Size : Color = font.Color : Bold = font.Bold
            Italic = font.Italic : Strikethrough = font.Strikethrough : UnderlineType = font.UnderlineType
        End Sub
        Friend Sub New(font As WorkbookCellAppearance)
            Name = font.FontName : Size = font.FontSize : Color = font.Foreground
            Bold = font.FontStyle.HasFlag(FontStyle.Bold) : Italic = font.FontStyle.HasFlag(FontStyle.Italic)
            Strikethrough = font.FontStyle.HasFlag(FontStyle.Strikeout)
            UnderlineType = If(font.FontStyle.HasFlag(FontStyle.Underline), DevExpress.Spreadsheet.UnderlineType.Single, DevExpress.Spreadsheet.UnderlineType.None)
        End Sub
    End Structure
    Public Structure WorkbookReadFill
        Public ReadOnly BackgroundColor As Color
        Public ReadOnly PatternType As PatternType
        Friend Sub New(background As Color, pattern As PatternType)
            BackgroundColor = background : PatternType = pattern
        End Sub
    End Structure
End Namespace
