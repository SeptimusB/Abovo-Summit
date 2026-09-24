Imports System.Globalization
Imports System.IO
Imports System.Net.Http
Imports System.Text.Json
Imports System.Threading.Tasks
Imports DevExpress.Spreadsheet

Namespace Abovo
    Public Enum FundingDateRule
        SameDay
        FirstDay
        LastDay
        FirstWorkingDay
        LastWorkingDay
    End Enum

    Public Class FundingScheduleRequest
        Public Property StartDate As DateTime = Date.Today
        Public Property IntervalMonths As Integer = 1
        Public Property Rule As FundingDateRule
        Public Property Occurrences As Integer = 12
        Public Property EndDate As DateTime?
        Public Property DelayBankHolidays As Boolean
    End Class

    Public Class FundingScheduleDate
        Public Property Number As Integer
        Public Property ScheduledDate As DateTime
        Public Property DateToAdd As DateTime
        Public Property Note As String
    End Class

    'Only official published years are treated as known. No projected holidays,
    'new package, VBA dependency or workbook metadata is required.
    Public Class FundingHolidayCalendar
        Public Property Region As String
        Public Property FetchedUtc As DateTime
        Public Property FromCache As Boolean
        Public ReadOnly Dates As New HashSet(Of DateTime)
        Public ReadOnly Years As New HashSet(Of Integer)

        Public Sub RequireCoverage(value As DateTime)
            If Not Years.Contains(value.Year) Then
                Throw New InvalidOperationException("GOV.UK has not supplied a complete " & Region &
                    " calendar for " & value.Year.ToString() & ". Shorten the schedule or turn off bank-holiday adjustment; future holidays will not be guessed.")
            End If
        End Sub

        Public Shared Function Parse(json As String, regionKey As String, fetched As DateTime) As FundingHolidayCalendar
            Dim result As New FundingHolidayCalendar With {.Region = regionKey, .FetchedUtc = fetched}
            Dim minimum = If(regionKey = "england-and-wales", 8, If(regionKey = "scotland", 9, If(regionKey = "northern-ireland", 10, 0)))
            If minimum = 0 Then Throw New ArgumentException("Select a UK bank-holiday region.")
            Using document = JsonDocument.Parse(json)
                For Each entry In document.RootElement.GetProperty(regionKey).GetProperty("events").EnumerateArray()
                    result.Dates.Add(DateTime.ParseExact(entry.GetProperty("date").GetString(), "yyyy-MM-dd", CultureInfo.InvariantCulture))
                Next
            End Using
            For Each holidayYear In result.Dates.GroupBy(Function(d) d.Year)
                If holidayYear.Count() >= minimum Then result.Years.Add(holidayYear.Key)
            Next
            If result.Years.Count = 0 Then Throw New InvalidDataException("No complete published holiday years were received.")
            Return result
        End Function

        Public Shared Async Function LoadAsync(regionKey As String) As Task(Of FundingHolidayCalendar)
            Dim folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Abovo", "Summit", "Calendars")
            Dim cache = Path.Combine(folder, "uk-bank-holidays.json")
            Try
                'Legacy hosts may default to TLS 1.0. Scope TLS 1.2 to this
                'download, retaining the normal platform certificate checks.
                Using handler As New HttpClientHandler With {.SslProtocols = System.Security.Authentication.SslProtocols.Tls12},
                      client As New HttpClient(handler) With {.Timeout = TimeSpan.FromSeconds(8)}
                    Dim json = Await client.GetStringAsync("https://www.gov.uk/bank-holidays.json")
                    Dim result = Parse(json, regionKey, DateTime.UtcNow)
                    Try
                        Directory.CreateDirectory(folder)
                        Dim temporary = cache & ".tmp"
                        File.WriteAllText(temporary, json)
                        If File.Exists(cache) Then
                            File.Replace(temporary, cache, Nothing)
                        Else
                            File.Move(temporary, cache)
                        End If
                    Catch cacheError As IOException
                        'A usable download does not depend on permission to cache it.
                    Catch cacheError As UnauthorizedAccessException
                    End Try
                    Return result
                End Using
            Catch downloadError As Exception
                If File.Exists(cache) AndAlso DateTime.UtcNow - File.GetLastWriteTimeUtc(cache) <= TimeSpan.FromDays(7) Then
                    Dim result = Parse(File.ReadAllText(cache), regionKey, File.GetLastWriteTimeUtc(cache))
                    result.FromCache = True
                    Return result
                End If
                Throw New InvalidOperationException("Cannot refresh the official UK holiday calendar. Check the connection and retry, or turn off bank-holiday adjustment. No dates have been applied.", downloadError)
            End Try
        End Function
    End Class

    Public NotInheritable Class FundingScheduleGenerator
        Public Shared Function Generate(request As FundingScheduleRequest, Optional holidays As FundingHolidayCalendar = Nothing) As List(Of FundingScheduleDate)
            If request Is Nothing Then Throw New ArgumentNullException(NameOf(request))
            If Not {1, 3, 6, 12}.Contains(request.IntervalMonths) Then Throw New ArgumentException("Choose monthly, quarterly, semi-annually or annually.")
            If Not [Enum].IsDefined(GetType(FundingDateRule), request.Rule) Then Throw New ArgumentException("Choose a date rule.")
            If request.StartDate.Year < 1900 Then Throw New ArgumentException("Choose a start date from 1900 onwards.")
            If request.EndDate.HasValue Then
                If request.EndDate.Value.Date < request.StartDate.Date Then Throw New ArgumentException("End date must not be before start date.")
            ElseIf request.Occurrences < 1 OrElse request.Occurrences > 1000 Then
                Throw New ArgumentException("Choose between 1 and 1,000 occurrences.")
            End If
            If request.DelayBankHolidays AndAlso holidays Is Nothing Then Throw New InvalidOperationException("Refresh the UK bank-holiday calendar first.")
            Dim result As New List(Of FundingScheduleDate)
            Dim anchor As New DateTime(request.StartDate.Year, request.StartDate.Month, 1)
            Dim monthIndex As Integer = 0
            Do
                Dim month = anchor.AddMonths(monthIndex * request.IntervalMonths)
                If request.EndDate.HasValue AndAlso month > request.EndDate.Value.Date Then Exit Do
                Dim lastDay = DateTime.DaysInMonth(month.Year, month.Month)
                Dim nominal = month.AddDays(Math.Min(request.StartDate.Day, lastDay) - 1)
                If request.Rule = FundingDateRule.FirstDay OrElse request.Rule = FundingDateRule.FirstWorkingDay Then nominal = month
                If request.Rule = FundingDateRule.LastDay OrElse request.Rule = FundingDateRule.LastWorkingDay Then nominal = month.AddDays(lastDay - 1)
                Dim actual = nominal
                If request.Rule = FundingDateRule.LastWorkingDay Then
                    While IsNonWorking(actual, request, holidays)
                        actual = actual.AddDays(-1)
                    End While
                ElseIf request.Rule = FundingDateRule.FirstWorkingDay OrElse request.DelayBankHolidays Then
                    While IsNonWorking(actual, request, holidays)
                        actual = actual.AddDays(1)
                    End While
                End If
                If actual >= request.StartDate.Date AndAlso (Not request.EndDate.HasValue OrElse actual <= request.EndDate.Value.Date) Then
                    result.Add(New FundingScheduleDate With {.Number = result.Count + 1, .ScheduledDate = nominal, .DateToAdd = actual,
                        .Note = If(actual = nominal, "", If(request.Rule = FundingDateRule.LastWorkingDay, "Previous working day", "Next working day"))})
                    If Not request.EndDate.HasValue AndAlso result.Count = request.Occurrences Then Exit Do
                    If result.Count > 1000 Then Throw New ArgumentException("This schedule exceeds 1,000 entries. Use a shorter end date.")
                End If
                monthIndex += 1
                If anchor.Year + (monthIndex * request.IntervalMonths \ 12) > 9998 Then Throw New ArgumentException("The schedule extends beyond the supported date range.")
            Loop
            If result.Count = 0 Then Throw New ArgumentException("No dates fall within the chosen period.")
            Return result
        End Function

        Private Shared Function IsNonWorking(value As DateTime, request As FundingScheduleRequest, holidays As FundingHolidayCalendar) As Boolean
            If request.DelayBankHolidays Then holidays.RequireCoverage(value)
            Return value.DayOfWeek = DayOfWeek.Saturday OrElse value.DayOfWeek = DayOfWeek.Sunday OrElse
                (request.DelayBankHolidays AndAlso holidays.Dates.Contains(value))
        End Function
    End Class

    Public Class FundingScheduleTarget
        Public Property Caption As String
        Public Property DateRange As String
        Public Property ValueRange As String
        Public Overrides Function ToString() As String
            Return Caption
        End Function
    End Class

    Public Class FundingSchedulePlacement
        Public Property StartOffset As Integer
        Public Property RowsToAdd As Integer
    End Class

    Public Class FundingScheduleApplication
        Public Property AddedRows As Integer
        Public ReadOnly FirstDateRows As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
    End Class

    Public NotInheritable Class FundingScheduleWriter
        'A narrow allow-list: never infer a structural mutation from a caption.
        Public Shared ReadOnly Targets As FundingScheduleTarget() = {
            Target("Facility Increases", "Fac_Inc", "06"), Target("Facility Decreases", "Fac_Dec", "07"),
            Target("Drawdowns", "Draw", "09"), Target("Repayments", "Repay", "10"),
            Target("Margins", "Marg", "19"), Target("Adjustment Spread", "MLA", "20"), Target("Commitment Fees", "Fees", "22"),
            Target("Variable Rate", "Var_Int", "16"), Target("Minimum Balance", "Min_Bal", "24"),
            Target("Interest Receivable", "Int_Rec", "25"), Target("Interest Payable", "Int_Pay", "26"),
            Target("Investment Increases", "Inv_Inc", "31"), Target("Investment Decreases", "Inv_Dec", "32"), Target("Investment Interest Rate", "Inv_Rate", "30")}

        Private Shared Function Target(caption As String, dateSuffix As String, valueSuffix As String) As FundingScheduleTarget
            Return New FundingScheduleTarget With {.Caption = caption, .DateRange = "IR_Fund_" & dateSuffix, .ValueRange = "Rep_Fund_" & valueSuffix}
        End Function

        Public Shared Function Ranges(workbook As IWorkbook, target As FundingScheduleTarget) As Tuple(Of CellRange, CellRange)
            If target Is Nothing OrElse Not Targets.Any(Function(t) t.DateRange = target.DateRange AndAlso t.ValueRange = target.ValueRange) Then Throw New ArgumentException("Unsupported Funding schedule section.")
            Dim dates = workbook.DefinedNames.GetDefinedName(target.DateRange)?.Range
            Dim values = workbook.DefinedNames.GetDefinedName(target.ValueRange)?.Range
            If dates Is Nothing OrElse values Is Nothing OrElse dates.ColumnCount <> 1 OrElse dates.Worksheet.Name <> "Funding Assumptions" OrElse
               dates.Worksheet IsNot values.Worksheet OrElse dates.TopRowIndex <> values.TopRowIndex OrElse
               (values.BottomRowIndex < dates.BottomRowIndex AndAlso Not (target.DateRange = "IR_Fund_Inv_Rate" AndAlso values.RowCount = 1)) Then
                Throw New InvalidOperationException("This workbook's Funding date/input ranges do not match the supported layout. No dates have been written.")
            End If
            'Rep_Fund_30 is deliberately only an anchor row. Its XML repeats
            'down by IR_Fund_Inv_Rate, using the same physical amount column.
            If target.DateRange = "IR_Fund_Inv_Rate" Then values = values.Worksheet.Range.FromLTRB(values.LeftColumnIndex, dates.TopRowIndex, values.RightColumnIndex, dates.BottomRowIndex)
            Return Tuple.Create(dates, values)
        End Function

        Public Shared Function BlankRows(workbook As IWorkbook, target As FundingScheduleTarget) As List(Of Integer)
            Dim pair = Ranges(workbook, target)
            Dim result As New List(Of Integer)
            For offset = 0 To pair.Item1.RowCount - 1
                Dim cell = pair.Item1(offset, 0)
                If cell.HasFormula OrElse Not cell.Value.IsEmpty OrElse cell.Protection.Locked OrElse cell.Fill.PatternType <> PatternType.Solid Then Continue For
                'Do not attach a new date to orphaned amounts or formula-owned inputs.
                Dim occupied = False
                For col = 0 To pair.Item2.ColumnCount - 1
                    Dim amount = pair.Item2(offset, col)
                    If amount.HasFormula OrElse Not amount.Value.IsEmpty Then occupied = True : Exit For
                Next
                If Not occupied Then result.Add(offset)
            Next
            Return result
        End Function

        Public Shared Function Placement(workbook As IWorkbook, target As FundingScheduleTarget, count As Integer) As FundingSchedulePlacement
            If count < 1 OrElse count > 1000 Then Throw New ArgumentException("Choose between 1 and 1,000 schedule dates.")
            Dim range = Ranges(workbook, target).Item1
            Dim blanks = BlankRows(workbook, target)
            Dim length = 0, previous = -2
            For Each offset In blanks
                length = If(offset = previous + 1, length + 1, 1)
                If length = count Then Return New FundingSchedulePlacement With {.StartOffset = offset - count + 1}
                previous = offset
            Next
            'Earlier isolated holes are not combined into one schedule. Extend
            'only the final safe blank run, keeping every new schedule together.
            Dim tail = 0
            For i = blanks.Count - 1 To 0 Step -1
                If blanks(i) <> range.RowCount - tail - 1 Then Exit For
                tail += 1
            Next
            Dim template = range(range.RowCount - 1, 0)
            If template.Protection.Locked OrElse template.HasFormula OrElse template.Fill.PatternType <> PatternType.Solid Then
                Throw New InvalidOperationException(target.Caption & ": the final date row is not an editable expansion template.")
            End If
            Return New FundingSchedulePlacement With {.StartOffset = range.RowCount - tail, .RowsToAdd = count - tail + 5}
        End Function

        Public Shared Function Apply(modelID As Integer, target As FundingScheduleTarget, dates As IList(Of DateTime)) As Integer
            Return ApplyBatch(modelID, {target}, dates, Nothing, Drawing.Color.Empty, "")
        End Function

        Public Shared Function ApplyBatch(modelID As Integer, sections As IEnumerable(Of FundingScheduleTarget), dates As IList(Of DateTime),
                                          facility As FundingScheduleFacility, colour As Drawing.Color, configuration As String) As Integer
            Return ApplyBatchWithAmount(modelID, sections, dates, facility, colour, configuration, Nothing)
        End Function

        Public Shared Function IsPercentage(workbook As IWorkbook, target As FundingScheduleTarget) As Boolean
            Return Ranges(workbook, target).Item2(0, 0).NumberFormat.Contains("%")
        End Function

        Public Shared Function ApplyBatchWithAmount(modelID As Integer, sections As IEnumerable(Of FundingScheduleTarget), dates As IList(Of DateTime),
                                                    facility As FundingScheduleFacility, colour As Drawing.Color, configuration As String,
                                                    fixedFigure As Decimal?) As Integer
            Return ApplyBatchCore(modelID, sections, dates, facility, colour, configuration, fixedFigure, Nothing)
        End Function

        Public Shared Function ApplyAvailableAmounts(modelID As Integer, sections As IEnumerable(Of FundingScheduleTarget), dates As IList(Of DateTime),
                                                     facility As FundingScheduleFacility, colour As Drawing.Color, configuration As String,
                                                     fixedFigure As Decimal?, skipped As IList(Of String)) As Integer
            If skipped Is Nothing Then Throw New ArgumentNullException(NameOf(skipped))
            Return ApplyBatchCore(modelID, sections, dates, facility, colour, configuration, fixedFigure, skipped)
        End Function

        Public Shared Function ApplyAvailableAmountsWithLocations(modelID As Integer, sections As IEnumerable(Of FundingScheduleTarget), dates As IList(Of DateTime),
                                                                 facility As FundingScheduleFacility, colour As Drawing.Color, configuration As String,
                                                                 fixedFigure As Decimal?, skipped As IList(Of String)) As FundingScheduleApplication
            If skipped Is Nothing Then Throw New ArgumentNullException(NameOf(skipped))
            Dim result As New FundingScheduleApplication
            result.AddedRows = ApplyBatchCore(modelID, sections, dates, facility, colour, configuration, fixedFigure, skipped, result.FirstDateRows)
            Return result
        End Function

        Private Shared Function ApplyBatchCore(modelID As Integer, sections As IEnumerable(Of FundingScheduleTarget), dates As IList(Of DateTime),
                                               facility As FundingScheduleFacility, colour As Drawing.Color, configuration As String,
                                               fixedFigure As Decimal?, skipped As IList(Of String), Optional firstDateRows As IDictionary(Of String, Integer) = Nothing) As Integer
            Dim targetsToWrite = sections.GroupBy(Function(t) t.DateRange).Select(Function(g) g.First()).ToList()
            If targetsToWrite.Count = 0 Then Throw New ArgumentException("Select at least one section.")
            If dates Is Nothing OrElse dates.Count = 0 OrElse dates.Count > 1000 Then Throw New ArgumentException("Preview between 1 and 1,000 dates first.")
            If dates.Any(Function(d) d.Year < 1900 OrElse d.TimeOfDay <> TimeSpan.Zero) Then Throw New ArgumentException("The preview contains an invalid date.")
            Dim model = FileManager.ExcelModels(modelID)
            If model Is Nothing OrElse model.IsClosing Then Throw New InvalidOperationException("The workbook is closing.")
            If model.ChangeManager.IsReadOnlyPreview OrElse model.IntegrityState = ModelIntegrityState.RecoveryRequired Then Throw New InvalidOperationException("This workbook is read-only or requires recovery; a schedule cannot be applied.")
            If fixedFigure.HasValue Then
                If facility Is Nothing Then Throw New ArgumentException("Choose a facility for the fixed figure.")
                If targetsToWrite.Select(Function(t) IsPercentage(model.WB, t)).Distinct().Count() > 1 Then Throw New ArgumentException("Apply amount and percentage schedules separately when using a fixed figure.")
            End If
            If facility IsNot Nothing Then
                FundingScheduleGroups.Read(model.WB) 'Reject unsupported metadata before structural changes.
                If Not FundingScheduleGroups.Facilities(model.WB, facility.IsInvestment).Any(Function(f) f.ColumnIndex = facility.ColumnIndex) Then Throw New ArgumentException("Select a valid facility.")
            End If
            'Preflight every section before adding any capacity.
            For Each target As FundingScheduleTarget In targetsToWrite
                Dim pair = Ranges(model.WB, target)
                If facility IsNot Nothing AndAlso (FundingScheduleGroups.IsInvestment(target) <> facility.IsInvestment OrElse facility.ColumnIndex < pair.Item2.LeftColumnIndex OrElse facility.ColumnIndex > pair.Item2.RightColumnIndex) Then Throw New ArgumentException("The selected facility is not part of " & target.Caption & ".")
                Placement(model.WB, target, dates.Count)
            Next
            Dim extra As Integer = 0, revertMetadata As Action = Nothing
            ModelSafetyManager.BeginBulkWorkbookMutation(modelID)
            Try
                For Each target As FundingScheduleTarget In targetsToWrite
                    Dim plan = Placement(model.WB, target, dates.Count)
                    If plan.RowsToAdd = 0 Then Continue For
                    Dim expanded = WorkbookManager.InsertRows(modelID, target.DateRange, plan.RowsToAdd)
                    If expanded Is Nothing OrElse expanded.BError Then Throw New InvalidOperationException(If(expanded Is Nothing, "No expansion result was returned.", expanded.StringReturn))
                    extra += plan.RowsToAdd
                Next
                Dim changes As New List(Of DataChangeEvent)
                Dim amounts As New List(Of DataChangeEvent)
                Dim entries As New List(Of Tuple(Of FundingScheduleTarget, Cell, DateTime))
                'All addresses are resolved after all expansions (later sections move).
                For Each target As FundingScheduleTarget In targetsToWrite
                    Dim plan = Placement(model.WB, target, dates.Count), range = Ranges(model.WB, target).Item1
                    If plan.RowsToAdd <> 0 Then Throw New InvalidOperationException("The expanded range did not provide enough consecutive safe blank date rows.")
                    For i = 0 To dates.Count - 1
                        Dim cell = range(plan.StartOffset + i, 0)
                        entries.Add(Tuple.Create(target, cell, dates(i)))
                        changes.Add(New DataChangeEvent With {.ModelID = modelID, .WSName = range.Worksheet.Name,
                            .CellAddress = cell.GetReferenceA1(), .ChangedValue = dates(i), .DataFormat = "D",
                            .Description = target.Caption & " schedule date " & (i + 1).ToString()})
                        If fixedFigure.HasValue Then
                            Dim amount = range.Worksheet.Cells(cell.RowIndex, facility.ColumnIndex)
                            Dim percent = IsPercentage(model.WB, target)
                            amounts.Add(New DataChangeEvent With {.ModelID = modelID, .WSName = range.Worksheet.Name,
                                .CellAddress = amount.GetReferenceA1(), .ChangedValue = CDbl(fixedFigure.Value / If(percent, 100D, 1D)),
                                .DataFormat = If(percent, "P", "N"), .Description = target.Caption & " schedule figure " & (i + 1).ToString()})
                        End If
                    Next
                Next
                If facility IsNot Nothing Then revertMetadata = FundingScheduleGroups.Add(model.WB, facility, colour, entries, configuration)
                changes.AddRange(amounts)
                Dim datesCalculated As Boolean = False
                Dim admit As Func(Of Cell, DataChangeEvent, Boolean) =
                    Function(cell, change)
                        If change.DataFormat = "D" Then Return True
                        'Dates control the conditional fill patterns. Calculate
                        'once inside the rollback boundary, then use the same
                        'fill/protection/formula rules as ordinary input.
                        If Not datesCalculated Then
                            cell.Worksheet.Calculate()
                            datesCalculated = True
                        End If
                        If cell.HasFormula OrElse Not cell.Value.IsEmpty OrElse cell.Protection.Locked OrElse cell.Fill.PatternType <> PatternType.Solid Then
                            If skipped IsNot Nothing Then
                                skipped.Add(cell.GetReferenceA1() & " (" & change.Description & ")")
                                Return False
                            End If
                            Throw New InvalidOperationException("The fixed figure cannot be entered in " & cell.GetReferenceA1() & ". The cell is unavailable or already contains data. No schedule values were kept; check the facility settings or add dates only.")
                        End If
                        Dim problem = DataInterfaceTemplate.NumericInputError(cell, New DataObject.DataColumnTag With {.DataType = change.DataFormat}, change.ChangedValue)
                        If problem IsNot Nothing Then Throw New InvalidOperationException(cell.GetReferenceA1() & ": " & problem)
                        Return True
                    End Function
                Dim posted = model.ChangeManager.ProcessAdmittedChanges(changes, "Funding schedule (" & dates.Count.ToString() & " dates × " & targetsToWrite.Count.ToString() & " sections)", Nothing, admit)
                If posted Is Nothing OrElse posted.BError OrElse Not posted.BSuccess Then Throw New InvalidOperationException(If(posted Is Nothing, "No date posting result was returned.", posted.StrResponseMessage))
                If firstDateRows IsNot Nothing Then
                    For Each section In entries.GroupBy(Function(entry) entry.Item1.DateRange)
                        firstDateRows(section.Key) = section.First().Item2.RowIndex
                    Next
                End If
                Return extra
            Catch ex As Exception
                If revertMetadata IsNot Nothing Then revertMetadata()
                If extra > 0 Then Throw New InvalidOperationException(ex.Message & " Extra blank rows were added and remain in the workbook; the date changes were not committed. Review the section before retrying.", ex)
                Throw
            Finally
                Try
                    If extra > 0 Then model.InterfaceDependencies.WorksheetStructureChanged("Funding Assumptions")
                Finally
                    ModelSafetyManager.EndBulkWorkbookMutation(modelID)
                End Try
            End Try
        End Function
    End Class
End Namespace
