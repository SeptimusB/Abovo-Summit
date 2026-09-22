Imports System.Globalization
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraEditors.Repository

Namespace Abovo
    'Excel owns the allowed dates. The old string repository lost their year and
    'serial value; use a date editor and validate all ChangeManager writes/pastes.
    Public NotInheritable Class FundingPaymentDateSupport
        Private Const ChoiceName As String = "FundMonthDesc"

        Public Shared Function AllowedDates(workbook As IWorkbook) As HashSet(Of DateTime)
            Dim range = workbook.DefinedNames.GetDefinedName(ChoiceName)?.Range
            If range Is Nothing Then Throw New InvalidOperationException("The workbook's funding month date list is missing.")
            Dim dates As New HashSet(Of DateTime)
            For Each cell As Cell In range.ExistingCells
                If cell.Value.IsEmpty Then Continue For
                If Not cell.Value.IsNumeric AndAlso Not cell.Value.IsDateTime Then Throw New InvalidOperationException("The funding month list contains a non-date value.")
                Dim dateValue = cell.Value.DateTimeValue.Date
                If dateValue.Day <> DateTime.DaysInMonth(dateValue.Year, dateValue.Month) Then
                    Throw New InvalidOperationException("The funding month list contains a date which is not a month end.")
                End If
                dates.Add(dateValue)
            Next
            If dates.Count = 0 Then Throw New InvalidOperationException("The funding month date list is empty.")
            Return dates
        End Function

        Public Shared Sub Configure(editor As RepositoryItemDateEdit, workbook As IWorkbook)
            editor.Mask.UseMaskAsDisplayFormat = False
            editor.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
            editor.DisplayFormat.FormatString = "mmm"
            editor.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime
            editor.EditFormat.FormatString = "dd-MMM-yyyy"
            AddHandler editor.DisableCalendarDate,
                Sub(s, e)
                    If e.View = DevExpress.XtraEditors.Controls.DateEditCalendarViewType.MonthInfo Then
                        e.IsDisabled = Not AllowedDates(workbook).Contains(e.DateTime.Date)
                    End If
                End Sub
        End Sub

        Public Shared Function ValidateChange(cell As Cell, ByRef value As Object) As DateTime?
            If Not String.Equals(cell.Worksheet.Name, "Funding Assumptions", StringComparison.OrdinalIgnoreCase) Then Return Nothing
            For Each validation In cell.Worksheet.DataValidations.GetDataValidations(cell)
                If validation.ValidationType <> DataValidationType.List OrElse validation.Criteria Is Nothing OrElse
                   Not validation.Criteria.IsFormula OrElse
                   Not String.Equals(validation.Criteria.FormulaInvariant.TrimStart("="c), ChoiceName, StringComparison.OrdinalIgnoreCase) Then Continue For
                If value Is Nothing OrElse Convert.IsDBNull(value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(value)) Then
                    If validation.AllowBlank Then
                        value = Nothing
                        Return Nothing
                    End If
                    Throw New FormatException("Choose a funding payment month.")
                End If
                Dim dateValue As DateTime
                If TypeOf value Is DateTime Then
                    dateValue = DirectCast(value, DateTime)
                Else
                    Dim serial As Double
                    If Double.TryParse(Convert.ToString(value, CultureInfo.CurrentCulture), NumberStyles.Any, CultureInfo.CurrentCulture, serial) Then
                        dateValue = DateTime.FromOADate(serial)
                    ElseIf Not DateTime.TryParse(Convert.ToString(value), CultureInfo.CurrentCulture, DateTimeStyles.None, dateValue) Then
                        Throw New FormatException("Choose a full month-end date from the funding month list, not a month name.")
                    End If
                End If
                If dateValue.TimeOfDay <> TimeSpan.Zero OrElse Not AllowedDates(cell.Worksheet.Workbook).Contains(dateValue) Then
                    Throw New FormatException("First Interest Payment Month must be one of the month-end dates in the workbook's funding month list.")
                End If
                Return dateValue
            Next
            Return Nothing
        End Function
    End Class
End Namespace
