Imports System.Globalization
Imports System.Drawing
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils.Drawing
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.XtraVerticalGrid.Events

''' <summary>
''' Shared clipboard parsing and workbook-cell conversion for workbook-backed
''' grids.  This module never writes cells: callers must route accepted values
''' through ModelChangeManager.
''' </summary>
Friend Module WorkbookGridClipboardSupport

    Friend Sub ConfigureVGridCellMultiSelect(ByVal grid As VGridControl)
        If grid Is Nothing Then Return
        grid.OptionsSelectionAndFocus.MultiSelect = True
        grid.OptionsSelectionAndFocus.MultiSelectMode = MultiSelectMode.CellSelect
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
        Return result
    End Function

    Friend Function InferDataFormat(ByVal cell As Cell) As String
        If cell Is Nothing Then Return "S"
        If cell.Value.IsDateTime Then Return "D"
        If cell.Value.IsBoolean Then Return "B"

        Dim numberFormat As String = If(cell.NumberFormat, String.Empty)
        If numberFormat.Contains("%") Then Return "P"
        If cell.Value.IsNumeric Then Return "N"
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
