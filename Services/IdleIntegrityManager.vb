Imports System.Configuration
Imports System.Diagnostics
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet

Namespace Abovo
    Friend NotInheritable Class IdleIntegritySettings
        Inherits ApplicationSettingsBase
        Friend Shared ReadOnly Instance As New IdleIntegritySettings()
        <UserScopedSetting(), DefaultSettingValue("False")>
        Public Property Enabled As Boolean
            Get
                Return CBool(Me(NameOf(Enabled)))
            End Get
            Set(value As Boolean)
                Me(NameOf(Enabled)) = value
            End Set
        End Property
        <UserScopedSetting(), DefaultSettingValue("30")>
        Public Property Minutes As Integer
            Get
                Return CInt(Me(NameOf(Minutes)))
            End Get
            Set(value As Integer)
                Me(NameOf(Minutes)) = value
            End Set
        End Property
    End Class

    'Observe Windows-session input even when an atomic UI-thread calculation is
    'running. No input is consumed, no message pumping and no worker touches WB.
    Friend NotInheritable Class IntegrityInputClock
        <StructLayout(LayoutKind.Sequential)>
        Private Structure LastInput
            Public Size As UInteger
            Public Tick As UInteger
        End Structure
        <DllImport("user32.dll")>
        Private Shared Function GetLastInputInfo(ByRef info As LastInput) As Boolean
        End Function
        Private Shared Seen As Boolean
        Private Shared Previous As UInteger
        Private Shared ReadOnly Stable As Stopwatch = Stopwatch.StartNew()

        Friend Shared Function InputStamp() As UInteger?
            Dim info As New LastInput With {.Size = CUInt(Marshal.SizeOf(GetType(LastInput)))}
            If Not GetLastInputInfo(info) Then Return Nothing
            Return info.Tick
        End Function

        Friend Shared Function IdleDuration() As TimeSpan
            Dim info As New LastInput With {.Size = CUInt(Marshal.SizeOf(GetType(LastInput)))}
            If Not GetLastInputInfo(info) Then
                Seen = False
                Stable.Restart()
                Return TimeSpan.Zero 'Fail closed if input state is unavailable.
            End If
            If Not Seen OrElse info.Tick <> Previous Then
                Seen = True
                Previous = info.Tick
                Stable.Restart()
            End If
            'Both counters are unsigned 32-bit ticks; account for wrap. A future
            'or out-of-order timestamp is conservatively treated as recent input.
            Dim current = CLng(Environment.TickCount) And &HFFFFFFFFL
            Dim elapsed = (current - CLng(info.Tick) + &H100000000L) Mod &H100000000L
            If elapsed > Integer.MaxValue Then Return TimeSpan.Zero
            Return TimeSpan.FromMilliseconds(Math.Min(elapsed, Stable.ElapsedMilliseconds))
        End Function
    End Class

    Friend NotInheritable Class IdleIntegrityManager
        Private NotInheritable Class Schedule
            Friend DueUtc As DateTime
            Friend Work As IntegrityPass
            Friend LastReport As String
            Friend LastRevision As Long = -1
            Friend LastCompletedUtc As DateTime
            Friend ManualRequested As Boolean
            Friend ManualStartPending As Boolean
            Friend RequestedInputStamp As UInteger?
            Friend Quiet As Boolean
        End Class
        Private Shared ReadOnly Plans As New Dictionary(Of FileManager.ExcelModel, Schedule)()
        Private Shared ReadOnly Clock As New Windows.Forms.Timer With {.Interval = 150}
        Private Shared Initialised As Boolean
        Private Shared Busy As Boolean
        Private Shared NextModel As Integer
        Friend Shared Property Enabled As Boolean
        Friend Shared Property Minutes As Integer = 30
        Friend Shared ReadOnly Property OperationInProgress As Boolean
            Get
                Return Busy
            End Get
        End Property

        Friend Shared Sub Initialise()
            If Initialised Then Return
            Initialised = True
            Try
                Enabled = IdleIntegritySettings.Instance.Enabled
                Minutes = Math.Max(1, Math.Min(120, IdleIntegritySettings.Instance.Minutes))
            Catch ex As ConfigurationErrorsException
                Abovo.SummitDiagnostics.WriteLine("[Idle Integrity] Settings unavailable: " & ex.Message)
            End Try
            AddHandler Clock.Tick, AddressOf Tick
            AddHandler Application.ApplicationExit, Sub()
                                                        Clock.Stop()
                                                        For Each model In Plans.Keys.ToArray()
                                                            Forget(model)
                                                        Next
                                                    End Sub
            IntegrityInputClock.IdleDuration()
            Clock.Start()
        End Sub

        Friend Shared Sub Configure(enableChecks As Boolean, interval As Integer, Optional persist As Boolean = True)
            If interval < 1 OrElse interval > 120 Then Throw New ArgumentOutOfRangeException(NameOf(interval))
            Initialise()
            If persist Then
                IdleIntegritySettings.Instance.Enabled = enableChecks
                IdleIntegritySettings.Instance.Minutes = interval
                IdleIntegritySettings.Instance.Save()
            End If
            If Enabled = enableChecks AndAlso Minutes = interval Then Return
            Enabled = enableChecks
            Minutes = interval
            For Each state In Plans.Values
                DiscardPass(state)
                state.DueUtc = DateTime.UtcNow.AddMinutes(Minutes)
            Next
        End Sub

        Friend Shared Sub Track(model As FileManager.ExcelModel)
            Initialise()
            Forget(model)
            Plans(model) = New Schedule With {.DueUtc = DateTime.UtcNow.AddMinutes(Minutes)}
            CheckSheetWatch.Track(model)
        End Sub

        Friend Shared Sub Forget(model As FileManager.ExcelModel)
            Dim state As Schedule = Nothing
            If Plans.TryGetValue(model, state) Then DiscardPass(state)
            Plans.Remove(model)
            CheckSheetWatch.Forget(model)
        End Sub

        Friend Shared Sub RequestNow(model As FileManager.ExcelModel)
            Initialise()
            If model Is Nothing OrElse model.IsClosing OrElse model.WB Is Nothing Then Throw New InvalidOperationException("Select an open business plan first.")
            If Not Plans.ContainsKey(model) Then Track(model)
            Dim state = Plans(model)
            DiscardPass(state)
            state.ManualRequested = True
            state.Quiet = False
            state.ManualStartPending = True
            state.RequestedInputStamp = IntegrityInputClock.InputStamp()
            state.DueUtc = DateTime.UtcNow
            'The command can be launched while another plan's messages are on
            'screen. Record the selected target in every plan's command log.
            SystemMessageManager.Publish(-1, "Integrity check requested for " & IO.Path.GetFileName(model.FileName) &
                ". The check will start when the workbook is available.",
                SystemMessageSeverity.Information, "Integrity", model.FileName)
        End Sub

        Friend Shared Sub RequestOnOpen(model As FileManager.ExcelModel)
            Initialise()
            If Not Plans.ContainsKey(model) Then Track(model)
            Dim state = Plans(model)
            DiscardPass(state)
            state.Quiet = True
            state.ManualRequested = True
            state.ManualStartPending = True
            state.DueUtc = DateTime.UtcNow
        End Sub

        Private Shared Sub DiscardPass(state As Schedule)
            If state.Work IsNot Nothing Then state.Work.Dispose()
            state.Work = Nothing
        End Sub

        Private Shared Sub Tick(sender As Object, e As EventArgs)
            'An input clock/API failure must never permit speculative idle work.
            Try
                If CheckSheetWatch.ProcessIdle(DateTime.UtcNow, IntegrityInputClock.IdleDuration()) Then Return
                If Plans.Count = 0 OrElse (Not Enabled AndAlso Not Plans.Values.Any(Function(s) s.ManualRequested)) Then Return
                ProcessIdle(DateTime.UtcNow, IntegrityInputClock.IdleDuration())
            Catch ex As Exception
                Abovo.SummitDiagnostics.WriteLine("[Idle Integrity] Scheduler deferred: " & ex.Message)
            End Try
        End Sub

        Friend Shared Function ProcessIdle(now As DateTime, idle As TimeSpan) As Boolean
            If Busy OrElse
               FileManager.BIsSaving OrElse FormSplashScreen.OperationInProgress OrElse
               RecoveryBackupManager.AdvanceNoticeVisible OrElse RecoveryBackupManager.PendingEditor() Then Return False
            If FileManager.ExcelModels IsNot Nothing AndAlso FileManager.ExcelModels.Any(
                Function(m) m IsNot Nothing AndAlso (ModelSafetyManager.IsBulkWorkbookMutationInProgress(m.ModelID) OrElse
                    (m.ChangeManager IsNot Nothing AndAlso m.ChangeManager.ChangeInProgress))) Then Return False
            Dim entries = Plans.ToArray()
            For offset = 0 To entries.Length - 1
                Dim index = (NextModel + offset) Mod entries.Length
                Dim model = entries(index).Key, state = entries(index).Value
                If Not Enabled AndAlso Not state.ManualRequested Then Continue For
                Dim inputStamp = IntegrityInputClock.InputStamp()
                'The explicit command authorises its first safe unit without an idle
                'wait. Mouse movement while Options closes must not turn Run now into
                'an unannounced two-minute delay. All editor/dialog/mutation gates above
                'and below still apply; later units continue to yield to user input.
                Dim manualReady = state.ManualRequested AndAlso (state.ManualStartPending OrElse
                    (state.RequestedInputStamp.HasValue AndAlso inputStamp.HasValue AndAlso
                     state.RequestedInputStamp.Value = inputStamp.Value))
                If Not manualReady AndAlso idle < TimeSpan.FromMinutes(2) Then Continue For
                If model.IsClosing OrElse model.WB Is Nothing OrElse model.ModelSpreadsheetControl Is Nothing OrElse model.ModelSpreadsheetControl.IsDisposed Then
                    Forget(model)
                    Continue For
                End If
                If now < state.DueUtc OrElse model.ModelSpreadsheetControl.InvokeRequired OrElse
                   model.IntegrityState <> ModelIntegrityState.Healthy OrElse model.RecoverySaveAsRequired OrElse
                   ModelSafetyManager.IsBulkWorkbookMutationInProgress(model.ModelID) OrElse
                   (model.ChangeManager IsNot Nothing AndAlso (model.ChangeManager.ChangeInProgress OrElse model.ChangeManager.IsReadOnlyPreview)) Then Continue For
                Dim owner = RecoveryBackupManager.NoticeOwner(model)
                If owner Is Nothing Then Continue For
                If state.Work IsNot Nothing AndAlso state.Work.Revision <> model.CalculationRevision Then
                    DiscardPass(state) 'Never certify a mixture of revisions or retain stale cell iterators.
                    Publish(model, "Workbook changed: the incomplete integrity pass will restart.")
                End If
                Busy = True
                Dim timer = Abovo.SummitDiagnostics.DiagnosticTimer.StartNew()
                Dim stage As String = "Start"
                Try
                    state.ManualStartPending = False
                    If state.Work Is Nothing Then
                        state.Work = New IntegrityPass(model, state.Quiet)
                        If Not state.Quiet Then Publish(model, "Checking workbook formulas and references.")
                    End If
                    stage = state.Work.Stage.ToString()
                    state.Work.Advance(owner)
                    If state.Work.Revision <> model.CalculationRevision Then
                        DiscardPass(state)
                    ElseIf state.Work.Finished Then
                        state.LastReport = state.Work.Report()
                        state.LastRevision = state.Work.Revision
                        state.LastCompletedUtc = now
                        Publish(model, state.LastReport, If(state.Work.IssueCount = 0, SystemMessageSeverity.Success, SystemMessageSeverity.Warning))
                        If Not state.Quiet Then
                        Using notice As New FormSplashScreen(owner, "Integrity check complete", IO.Path.GetFileName(model.FileName))
                            notice.Complete(If(state.Work.IssueCount = 0, "No issues found by the implemented checks.", state.Work.IssueCount.ToString() & " issue(s) found. See System Messages for details."))
                        End Using
                        End If
                        DiscardPass(state)
                        state.ManualRequested = False
                        state.DueUtc = now.AddMinutes(Minutes)
                    End If
                Catch ex As Exception
                    DiscardPass(state)
                    state.ManualRequested = False
                    state.DueUtc = now.AddMinutes(Minutes)
                    Publish(model, "Integrity check incomplete at " & stage & ": " & ex.Message & ". No pass recorded; will retry at the next interval.", SystemMessageSeverity.Warning)
                Finally
                    Busy = False
                    NextModel = (index + 1) Mod entries.Length
                    Abovo.SummitDiagnostics.WriteLine("[Idle Integrity Benchmark] model=" & model.ModelID.ToString() & ", stage=" & stage & ", total=" & timer.ElapsedMilliseconds.ToString() & " ms")
                End Try
                Return True 'One safe unit only; return to the UI message loop.
            Next
            Return False
        End Function

        Private Shared Sub Publish(model As FileManager.ExcelModel, text As String, Optional severity As SystemMessageSeverity = SystemMessageSeverity.Information)
            SystemMessageManager.Publish(model.ModelID, text, severity, "Integrity", model.FileName)
        End Sub

        Private Enum CheckStage
            Calculation
            CheckSheet
            Names
            Mirrors
            Cells
            Complete
        End Enum

        Private NotInheritable Class IntegrityPass
            Implements IDisposable
            Private ReadOnly Model As FileManager.ExcelModel
            Friend ReadOnly Revision As Long
            Friend Stage As CheckStage
            Friend IssueCount As Integer
            Friend CheckSheetFailed As Boolean
            Private ReadOnly Samples As New List(Of String)()
            Private ReadOnly SampleCounts As New Dictionary(Of String, Integer)()
            Private ReadOnly Counts As New SortedDictionary(Of String, Integer)()
            Private Names As IEnumerator(Of DefinedName)
            Private Cells As IEnumerator(Of Cell)
            Private SheetIndex As Integer
            Private CellsChecked As Long
            Private NamesChecked As Integer
            Private ExplicitNaCells As Integer
            Private ExpectedChartGaps As Integer
            Private OverridesAccepted As Integer
            Private CompatibilityNotices As Integer
            Private ReadOnly ChartGaps As New ExpectedChartGapClassifier()
            Private ReadOnly Notices As New List(Of String)()
            Private Calculated As Boolean
            Private ReadOnly Quiet As Boolean
            Private ReadOnly Started As DateTime = DateTime.UtcNow
            Friend ReadOnly Property Finished As Boolean
                Get
                    Return Stage = CheckStage.Complete
                End Get
            End Property

            Friend Sub New(model As FileManager.ExcelModel, Optional quiet As Boolean = False)
                Me.Model = model
                Me.Quiet = quiet
                Revision = model.CalculationRevision
            End Sub

            Private Sub AddIssue(category As String, detail As String)
                IssueCount += 1
                If Not Counts.ContainsKey(category) Then Counts(category) = 0
                Counts(category) += 1
                If Not SampleCounts.ContainsKey(category) Then SampleCounts(category) = 0
                If Samples.Count < 100 AndAlso SampleCounts(category) < 10 Then
                    Samples.Add(category & ": " & detail)
                    SampleCounts(category) += 1
                End If
            End Sub

            Friend Sub Advance(owner As Form)
                Select Case Stage
                    Case CheckStage.Calculation
                        Calculated = Model.ResultsPending
                        If Calculated AndAlso Quiet Then
                            Model.CalculateForIdleIntegrity(Nothing)
                            If Model.WBCalcEngine IsNot Nothing Then Model.WBCalcEngine.RefreshAfterDeferredCalculation()
                            Stage = CheckStage.CheckSheet
                            ValidateCheckSheet(owner)
                        ElseIf Calculated Then
                            Using notice As New FormSplashScreen(owner, "Integrity: updating calculations", IO.Path.GetFileName(Model.FileName) & Environment.NewLine & "Please wait. This calculation cannot be interrupted safely.")
                                If Not notice.IsShowing Then Throw New InvalidOperationException("Progress notice unavailable; calculation deferred")
                                Model.CalculateForIdleIntegrity(AddressOf notice.Update)
                                'Required completion of the calculation unit: refresh results
                                'before accepting input, never leave the visible grid half-updated.
                                If Model.WBCalcEngine IsNot Nothing Then Model.WBCalcEngine.RefreshAfterDeferredCalculation()
                                If Model.InstanceInterface IsNot Nothing AndAlso Not Model.InstanceInterface.IsDisposed Then Model.InstanceInterface.PopulateFileInfo()
                                'Publish the brief, revision-checked Check Sheet result before
                                'yielding. Input during calculation must not leave a previous
                                'recovery hold waiting for the next idle window.
                                Stage = CheckStage.CheckSheet
                                ValidateCheckSheet(owner)
                                notice.Complete("Calculations and Check Sheet checked. Remaining integrity checks continue in idle stages.")
                            End Using
                        Else
                            Stage = CheckStage.CheckSheet
                            ValidateCheckSheet(owner)
                        End If
                    Case CheckStage.CheckSheet
                        ValidateCheckSheet(owner)
                    Case CheckStage.Names
                        Dim timer = Stopwatch.StartNew(), count = 0
                        While count < 100 AndAlso timer.ElapsedMilliseconds < 25
                            If Not Names.MoveNext() Then
                                Names.Dispose()
                                Names = Nothing
                                Stage = CheckStage.Mirrors
                                Return
                            End If
                            Dim name = Names.Current
                            'A deleted schedule anchor is inactive optional UI metadata,
                            'not a broken financial formula. Require explicit XML ownership.
                            If name.RefersTo.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso FundingScheduleGroups.IsPresentationAnchor(Model.WB, name) Then
                                count += 1
                                NamesChecked += 1
                                Continue While
                            End If
                            Dim expression = Regex.Replace(name.RefersTo, """(?:""""|[^""])*""", "")
                            If expression.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0 Then AddIssue("Broken name", name.Name & " refers to " & name.RefersTo)
                            InspectExportFormula(name.RefersTo, "Name " & name.Name)
                            NamesChecked += 1
                            count += 1
                        End While
                    Case CheckStage.Mirrors
                        If Model.TransDBSync IsNot Nothing Then
                            For Each issue In Model.TransDBSync.InspectMirrorGeometry(Model.WB)
                                AddIssue("Mirror geometry", issue)
                            Next
                        End If
                        Stage = CheckStage.Cells
                    Case CheckStage.Cells
                        ScanCells()
                End Select
            End Sub

            Private Sub ValidateCheckSheet(owner As Form)
                Dim timer = Abovo.SummitDiagnostics.DiagnosticTimer.StartNew()
                Try
                    'Never certify caches from an interrupted or superseded calculation.
                    If Model.ResultsPending OrElse Revision <> Model.CalculationRevision Then
                        Throw New InvalidOperationException("The workbook changed before Check Sheet validation. A fresh check is required.")
                    End If
                    Dim result = Model.ReadCheckSheetValidation()
                    CheckSheetFailed = result.HasFailures
                    If String.IsNullOrWhiteSpace(result.ValidationError) Then
                        If Model.RecordIdleCheckSheetResult(Revision, result.HasFailures) Then
                            RecoveryBackupManager.NotifyCheckSheetState(Model, result, owner)
                        End If
                        If Not Quiet Then Publish(Model, "Check Sheet: " & If(result.HasFailures, "some figures do not balance yet.", "balanced."), SystemMessageSeverity.Information)
                    Else
                        Publish(Model, result.ValidationError, SystemMessageSeverity.Warning)
                    End If
                    If Not String.IsNullOrWhiteSpace(result.ValidationError) Then AddIssue("Check Sheet", result.ValidationError)
                    For Each issue In result.Issues
                        Notices.Add("Check Sheet (balance only): row " & issue.CheckRow.ToString() & ", " & issue.Label & ": " & issue.Status)
                    Next
                    OverridesAccepted = result.OverriddenIssues.Count
                    For Each issue In result.OverriddenIssues.Take(10)
                        Notices.Add("Accepted Check Sheet override: row " & issue.CheckRow.ToString() & ", " & issue.Label & ". Visible status remains " & issue.Status & "; effective check count is zero.")
                    Next
                    Names = Model.WB.DefinedNames.Concat(Model.WB.Worksheets.SelectMany(Function(s) s.DefinedNames)).Distinct().GetEnumerator()
                    Stage = CheckStage.Names
                Finally
                    Abovo.SummitDiagnostics.WriteLine("[Idle Integrity Benchmark] model=" & Model.ModelID.ToString() & ", stage=CheckSheet, total=" & timer.ElapsedMilliseconds.ToString() & " ms")
                End Try
            End Sub

            Private Sub InspectExportFormula(formula As String, location As String)
                Try
                    If WorkbookXlsbFormulaCompatibility.Normalize(Model.WB, formula) <> formula Then
                        CompatibilityNotices += 1
                        If CompatibilityNotices <= 10 Then Notices.Add("Save compatibility notice: " & location & " will receive the verified formula normalization on XLSB save; not a calculation failure.")
                    End If
                Catch ex As Exception
                    AddIssue("XLSB compatibility", location & ": " & ex.Message)
                End Try
            End Sub

            Private Sub ScanCells()
                Dim timer = Stopwatch.StartNew(), count = 0
                While count < 5000 AndAlso timer.ElapsedMilliseconds < 25
                    If Cells Is Nothing Then
                        If SheetIndex >= Model.WB.Worksheets.Count Then
                            Stage = CheckStage.Complete
                            Return
                        End If
                        Cells = Model.WB.Worksheets(SheetIndex).GetUsedRange().ExistingCells.GetEnumerator()
                        SheetIndex += 1
                    End If
                    If Not Cells.MoveNext() Then
                        Cells.Dispose()
                        Cells = Nothing
                        Continue While
                    End If
                    Dim cell = Cells.Current
                    CellsChecked += 1
                    count += 1
                    Dim formula = If(cell.HasFormula, cell.FormulaInvariant, "")
                    Dim location = cell.Worksheet.Name & "!" & cell.GetReferenceA1()
                    Dim currentValue = cell.ModelValue()
                    If currentValue.IsError Then
                        If String.Equals(formula.Replace(" ", ""), "=NA()", StringComparison.OrdinalIgnoreCase) Then
                            ExplicitNaCells += 1 'Explicit NA() sentinels are chart gaps, not broken calculations.
                        ElseIf ChartGaps.IsExpected(cell) Then
                            ExpectedChartGaps += 1
                        Else
                            Dim errorText = If(cell.HasModelEngineView(), cell.ModelDisplayText(), currentValue.ToString())
                            AddIssue("Cell error " & errorText, location & " = " & errorText)
                        End If
                    End If
                    If cell.HasFormula Then InspectExportFormula(formula, location)
                End While
            End Sub

            Friend Function Report() As String
                Dim text = "Integrity check completed for revision " & Revision.ToString() & " (" & Started.ToLocalTime().ToString("HH:mm:ss") & "). " &
                    IssueCount.ToString() & " issue(s); " & CellsChecked.ToString("N0") & " existing cells and " & NamesChecked.ToString("N0") & " names inspected. " &
                    ExplicitNaCells.ToString() & " explicit NA() sentinel(s) counted separately."
                text &= Environment.NewLine & "Separate non-blocking findings: " & ExpectedChartGaps.ToString() & " verified chart gaps; " & OverridesAccepted.ToString() & " accepted Check Sheet overrides; " & CompatibilityNotices.ToString() & " supported save-normalization notices."
                text &= Environment.NewLine & RecoveryBackupManager.CheckSheetPolicyDescription(Model) & " Backup settings and unsaved-input rules still apply."
                If Counts.Count > 0 Then text &= Environment.NewLine & String.Join("; ", Counts.Select(Function(p) p.Key & "=" & p.Value.ToString()))
                If Samples.Count > 0 Then text &= Environment.NewLine & String.Join(Environment.NewLine, Samples)
                If Notices.Count > 0 Then text &= Environment.NewLine & String.Join(Environment.NewLine, Notices)
                If IssueCount > Samples.Count Then text &= Environment.NewLine & "Sampled locations: up to 10 per category, 100 total; totals above include all findings."
                Return text & Environment.NewLine & "Inspection only: no automatic repair or save. Supported checks do not constitute financial or Excel/VBA certification. Later edits invalidate this report."
            End Function

            Public Sub Dispose() Implements IDisposable.Dispose
                If Names IsNot Nothing Then Names.Dispose()
                If Cells IsNot Nothing Then Cells.Dispose()
                Names = Nothing
                Cells = Nothing
            End Sub
        End Class
    End Class
End Namespace
