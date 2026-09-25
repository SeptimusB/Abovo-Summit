Imports System.Globalization
Imports Abovo.ModelWorkbookRead
Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils.Drawing
Imports DevExpress.Utils
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.XtraVerticalGrid.Events

''' <summary>
''' Shared clipboard parsing and workbook-cell conversion for workbook-backed
''' grids.  This module never writes cells: callers must route accepted values
''' through ModelChangeManager.
''' </summary>
Friend Module WorkbookGridClipboardSupport

    Private Const HeaderClipboardFormat As String = "AbovoSummit.ClipboardHeaders.v1"

    Friend Sub SetClipboardTextWithHeaders(ByVal value As String,
                                           ByVal headingRows As Integer,
                                           ByVal headingColumns As Integer)
        Dim data As New DataObject()
        data.SetText(If(value, String.Empty), TextDataFormat.UnicodeText)
        data.SetData(HeaderClipboardFormat,
                     headingRows.ToString(Globalization.CultureInfo.InvariantCulture) & "," &
                     headingColumns.ToString(Globalization.CultureInfo.InvariantCulture))
        Clipboard.SetDataObject(data, True)
    End Sub

    Friend Sub CopyGridWithHeaders(ByVal grid As GridControl)
        CopyGridSelection(grid, True)
    End Sub

    Friend Sub CopyGridSelection(ByVal grid As GridControl,
                                 Optional ByVal includeHeadings As Boolean = False)
        If grid Is Nothing Then Return
        Dim view As GridView = TryCast(grid.FocusedView, GridView)
        If view Is Nothing Then Return
        Dim selectedRows As New List(Of Integer)()
        Dim selectedColumns As New List(Of Integer)()
        For Each cell In view.GetSelectedCells()
            If cell.RowHandle < 0 OrElse cell.Column Is Nothing OrElse
               cell.Column.VisibleIndex < 0 Then Continue For
            Dim visibleRow As Integer = view.GetVisibleIndex(cell.RowHandle)
            If visibleRow < 0 Then Continue For
            selectedRows.Add(visibleRow)
            selectedColumns.Add(cell.Column.VisibleIndex)
        Next
        For Each rowHandle As Integer In view.GetSelectedRows()
            If rowHandle < 0 Then Continue For
            Dim visibleRow As Integer = view.GetVisibleIndex(rowHandle)
            If visibleRow >= 0 Then selectedRows.Add(visibleRow)
        Next
        If selectedColumns.Count = 0 AndAlso selectedRows.Count > 0 Then
            selectedColumns.AddRange(view.VisibleColumns.Cast(Of GridColumn)().
                                     Select(Function(column) column.VisibleIndex))
        End If

        'Only use the focused cell when there is no explicit selection. A nearby
        'focused editor may be outside the selected rectangle.
        If (selectedRows.Count = 0 OrElse selectedColumns.Count = 0) AndAlso
           view.FocusedRowHandle >= 0 AndAlso view.FocusedColumn IsNot Nothing AndAlso
           view.FocusedColumn.VisibleIndex >= 0 Then
            Dim focusedVisibleRow As Integer = view.GetVisibleIndex(view.FocusedRowHandle)
            If focusedVisibleRow >= 0 Then
                selectedRows.Add(focusedVisibleRow)
                selectedColumns.Add(view.FocusedColumn.VisibleIndex)
            End If
        End If
        If selectedRows.Count = 0 OrElse selectedColumns.Count = 0 Then Return

        Dim firstColumn As Integer = selectedColumns.Min()
        Dim lastColumn As Integer = selectedColumns.Max()
        Dim columns As List(Of GridColumn) = view.VisibleColumns.Cast(Of GridColumn)().
            Where(Function(column) column.VisibleIndex >= firstColumn AndAlso
                                  column.VisibleIndex <= lastColumn).
            OrderBy(Function(column) column.VisibleIndex).ToList()
        If columns.Count = 0 Then Return

        Dim clipboardRows As New List(Of String)()
        Dim headingRowCount As Integer = 0
        If includeHeadings Then
            Dim topHeadings As New List(Of String)()
            Dim lowerHeadings As New List(Of String)()
            For Each column As GridColumn In columns
                Dim parts As String() = If(column.Caption, String.Empty).
                    Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).
                    Split(New Char() {ControlChars.Lf}, StringSplitOptions.None)
                Dim firstLine As String = parts(0).Trim()
                Dim remainingLines As String =
                    If(parts.Length > 1, String.Join(" ", parts.Skip(1)).Trim(),
                       String.Empty)
                If remainingLines.Length > 0 AndAlso
                   (firstLine.StartsWith("Year ", StringComparison.OrdinalIgnoreCase) OrElse
                    firstLine.StartsWith("Period ", StringComparison.OrdinalIgnoreCase)) Then
                    topHeadings.Add(firstLine)
                    lowerHeadings.Add(remainingLines)
                Else
                    topHeadings.Add(String.Join(" ", parts.Select(Function(part) part.Trim())).
                                    Trim())
                    lowerHeadings.Add(String.Empty)
                End If
            Next
            clipboardRows.Add(String.Join(ControlChars.Tab, topHeadings))
            If lowerHeadings.Any(Function(value) value.Length > 0) Then
                clipboardRows.Add(String.Join(ControlChars.Tab, lowerHeadings))
            End If
            headingRowCount = clipboardRows.Count
        End If

        For visibleRow As Integer = selectedRows.Min() To selectedRows.Max()
            Dim rowHandle As Integer = view.GetVisibleRowHandle(visibleRow)
            If rowHandle < 0 Then Continue For
            Dim values As New List(Of String)()
            For Each column As GridColumn In columns
                values.Add(If(view.GetRowCellDisplayText(rowHandle, column),
                              String.Empty).
                           Replace(ControlChars.Tab, " ").
                           Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " "))
            Next
            clipboardRows.Add(String.Join(ControlChars.Tab, values))
        Next
        If clipboardRows.Count = headingRowCount Then Return

        Dim copiedText As String = String.Join(vbCrLf, clipboardRows)
        If includeHeadings Then
            SetClipboardTextWithHeaders(copiedText, headingRowCount, 0)
        Else
            Clipboard.SetText(copiedText, TextDataFormat.UnicodeText)
        End If
    End Sub

    Friend Sub CopyVGridWithHeaders(ByVal grid As VGridControl)
        If grid Is Nothing Then Return
        Dim previous As Boolean = grid.OptionsBehavior.CopyToClipboardWithRowHeaders
        Try
            grid.OptionsBehavior.CopyToClipboardWithRowHeaders = True
            grid.CopyToClipboard()
        Finally
            grid.OptionsBehavior.CopyToClipboardWithRowHeaders = previous
        End Try
        If Clipboard.ContainsText() Then
            SetClipboardTextWithHeaders(Clipboard.GetText(TextDataFormat.UnicodeText), 0, 1)
        End If
    End Sub

    Friend Sub AddCopyWithHeadersMenu(ByVal grid As GridControl)
        If grid Is Nothing Then Return
        Dim menu As ContextMenuStrip = If(grid.ContextMenuStrip, New ContextMenuStrip())
        If menu.Items.Find("SummitCopyWithHeaders", False).Length > 0 Then Return
        menu.Items.Add(New ToolStripMenuItem("Copy", Nothing,
            Sub() CopyGridSelection(grid)))
        Dim headerItem As New ToolStripMenuItem("Copy with headings", Nothing,
            Sub() CopyGridWithHeaders(grid)) With {.Name = "SummitCopyWithHeaders"}
        menu.Items.Add(headerItem)
        grid.ContextMenuStrip = menu
    End Sub

    Friend Sub AddCopyWithHeadersMenu(ByVal grid As VGridControl)
        If grid Is Nothing Then Return
        Dim menu As ContextMenuStrip = If(grid.ContextMenuStrip, New ContextMenuStrip())
        If menu.Items.Find("SummitCopyWithHeaders", False).Length > 0 Then Return
        menu.Items.Add(New ToolStripMenuItem("Copy", Nothing,
            Sub() grid.CopyToClipboard()))
        Dim headerItem As New ToolStripMenuItem("Copy with headings", Nothing,
            Sub() CopyVGridWithHeaders(grid)) With {.Name = "SummitCopyWithHeaders"}
        menu.Items.Add(headerItem)
        grid.ContextMenuStrip = menu
    End Sub

    Friend Sub ConfigureVGridCellMultiSelect(ByVal grid As VGridControl)
        If grid Is Nothing Then Return
        grid.OptionsSelectionAndFocus.MultiSelect = True
        grid.OptionsSelectionAndFocus.MultiSelectMode = MultiSelectMode.CellSelect
        grid.OptionsBehavior.EditorShowMode = EditorShowMode.MouseUp
        grid.OptionsBehavior.RecordsMouseWheel = False
        grid.Appearance.SelectedCell.BackColor = SystemColors.Highlight
        grid.Appearance.SelectedCell.ForeColor = SystemColors.HighlightText
        grid.Appearance.SelectedCell.Options.UseBackColor = True
        grid.Appearance.SelectedCell.Options.UseForeColor = True
    End Sub

    Friend Sub ApplyVGridSelectedCellAppearance(ByVal e As CustomDrawRowValueCellEventArgs)
        If e Is Nothing OrElse e.RowValueInfo Is Nothing Then Return
        If (e.RowValueInfo.State And ObjectState.Selected) <> ObjectState.Selected Then Return
        e.Appearance.BackColor = SystemColors.Highlight
        e.Appearance.ForeColor = SystemColors.HighlightText
        e.Appearance.Options.UseBackColor = True
        e.Appearance.Options.UseForeColor = True
    End Sub

    Friend Function ReadClipboardMatrix() As List(Of String())
        Dim result As New List(Of String())()
        If Not Clipboard.ContainsText() Then Return result

        Dim clipboardText As String = Clipboard.GetText(TextDataFormat.UnicodeText)
        If String.IsNullOrEmpty(clipboardText) Then Return result

        Dim normalised As String = clipboardText.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        While normalised.EndsWith(vbLf, StringComparison.Ordinal)
            normalised = normalised.Substring(0, normalised.Length - 1)
        End While
        If normalised.Length = 0 Then Return result

        For Each clipboardRow As String In normalised.Split(New String() {vbLf}, StringSplitOptions.None)
            result.Add(clipboardRow.Split(New Char() {ControlChars.Tab}, StringSplitOptions.None))
        Next
        If Clipboard.ContainsData(HeaderClipboardFormat) Then
            Dim metadata As String = TryCast(Clipboard.GetData(HeaderClipboardFormat), String)
            Dim parts() As String = If(metadata, String.Empty).Split(","c)
            Dim headingRows As Integer
            Dim headingColumns As Integer
            If parts.Length = 2 AndAlso Integer.TryParse(parts(0), headingRows) AndAlso
               Integer.TryParse(parts(1), headingColumns) AndAlso
               headingRows >= 0 AndAlso headingColumns >= 0 AndAlso
               headingRows <= result.Count Then
                result.RemoveRange(0, headingRows)
                If result.Count = 0 Then Return result
                If headingColumns > 0 Then
                    For index As Integer = 0 To result.Count - 1
                        If result(index).Length < headingColumns Then Return New List(Of String())()
                        result(index) = result(index).Skip(headingColumns).ToArray()
                    Next
                End If
            End If
        End If
        Return result
    End Function

    Friend Function InferDataFormat(ByVal cell As Cell) As String
        If cell Is Nothing Then Return "S"
        If cell.ModelValue().IsDateTime Then Return "D"
        If cell.ModelValue().IsBoolean Then Return "B"

        Dim numberFormat As String = If(cell.ModelNumberFormat(), String.Empty)
        If cell.HasModelEngineView() AndAlso cell.ModelValue().IsNumeric AndAlso
            (numberFormat.IndexOf("yy", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
             numberFormat.IndexOf("dd", StringComparison.OrdinalIgnoreCase) >= 0) Then Return "D"
        If numberFormat.Contains("%") Then Return "P"
        If cell.ModelValue().IsNumeric Then Return "N"
        Return "S"
    End Function

    Friend Function TryConvertClipboardValue(ByVal rawValue As String,
                                             ByVal dataFormat As String,
                                             ByRef convertedValue As Object) As Boolean
        If String.IsNullOrEmpty(rawValue) Then
            convertedValue = Nothing
            Return True
        End If

        Select Case If(dataFormat, String.Empty).ToUpperInvariant()
            Case "S"
                convertedValue = rawValue
                Return True
            Case "B"
                Dim booleanValue As Boolean
                If Boolean.TryParse(rawValue, booleanValue) Then
                    convertedValue = booleanValue
                    Return True
                End If
                Dim booleanNumber As Double
                If Double.TryParse(rawValue, NumberStyles.Any, CultureInfo.CurrentCulture, booleanNumber) Then
                    convertedValue = booleanNumber <> 0
                    Return True
                End If
                Return False
            Case "D"
                Dim dateValue As DateTime
                If DateTime.TryParse(rawValue, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, dateValue) Then
                    convertedValue = dateValue
                    Return True
                End If
                Return False
            Case "I", "Y"
                Dim integerValue As Integer
                If Integer.TryParse(rawValue, NumberStyles.Any, CultureInfo.CurrentCulture, integerValue) Then
                    convertedValue = integerValue
                    Return True
                End If
                Return False
            Case "N", "C", "M", "SM", "R", "P"
                Dim numericText As String = rawValue.Trim()
                Dim clipboardPercent As Boolean = numericText.EndsWith("%", StringComparison.Ordinal)
                If clipboardPercent Then numericText = numericText.Substring(0, numericText.Length - 1).Trim()
                Dim numericValue As Double
                If Not Double.TryParse(numericText, NumberStyles.Any, CultureInfo.CurrentCulture, numericValue) Then Return False
                If clipboardPercent Then numericValue /= 100.0
                convertedValue = numericValue
                Return True
            Case Else
                convertedValue = rawValue
                Return True
        End Select
    End Function

End Module
