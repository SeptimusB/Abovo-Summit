Imports System.Configuration
Imports System.Linq
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet

Namespace Abovo
    Friend NotInheritable Class CheckSheetWatchSettings
        Inherits ApplicationSettingsBase
        Friend Shared ReadOnly Instance As New CheckSheetWatchSettings()
        <UserScopedSetting(), DefaultSettingValue("False")>
        Public Property Enabled As Boolean
            Get
                Return CBool(Me(NameOf(Enabled)))
            End Get
            Set(value As Boolean)
                Me(NameOf(Enabled)) = value
            End Set
        End Property
        <UserScopedSetting(), DefaultSettingValue("5")>
        Public Property Minutes As Integer
            Get
                Return CInt(Me(NameOf(Minutes)))
            End Get
            Set(value As Integer)
                Me(NameOf(Minutes)) = value
            End Set
        End Property
        <UserScopedSetting(), DefaultSettingValue("2")>
        Public Property IdleMinutes As Integer
            Get
                Return CInt(Me(NameOf(IdleMinutes)))
            End Get
            Set(value As Integer)
                Me(NameOf(IdleMinutes)) = value
            End Set
        End Property
        <UserScopedSetting(), DefaultSettingValue("False")>
        Public Property Benchmark As Boolean
            Get
                Return CBool(Me(NameOf(Benchmark)))
            End Get
            Set(value As Boolean)
                Me(NameOf(Benchmark)) = value
            End Set
        End Property
    End Class

    'Full-calculation / soft-balance trial. Never certifies the whole model or changes its dirty,
    'pending-results, user-revision or rebuild flags. Runs on the owning UI thread.
    Friend NotInheritable Class CheckSheetWatch
        Private NotInheritable Class Schedule
            Friend Due As DateTime
            Friend Revision As Long = -1
            Friend LastError As String
            Friend FullCalculationCount As Integer
        End Class
        Private Shared ReadOnly Plans As New Dictionary(Of FileManager.ExcelModel, Schedule)
        Private Shared ReadOnly Running As New HashSet(Of FileManager.ExcelModel)
        Private Shared Loaded As Boolean
        Friend Shared Property Enabled As Boolean
        Friend Shared Property Minutes As Integer = 5
        Friend Shared Property IdleMinutes As Integer = 2
        Friend Shared Property Benchmark As Boolean
        Friend Shared Sub Initialise()
            If Loaded Then Return
            Loaded = True
            Try
                Dim s = CheckSheetWatchSettings.Instance
                Enabled = s.Enabled : Minutes = Math.Max(1, Math.Min(120, s.Minutes))
                IdleMinutes = Math.Max(1, Math.Min(120, s.IdleMinutes)) : Benchmark = s.Benchmark
            Catch ex As ConfigurationErrorsException
                Enabled = False : Benchmark = False
            End Try
        End Sub
        Friend Shared Sub Configure(enabledValue As Boolean, minutesValue As Integer, idleValue As Integer, benchmarkValue As Boolean, Optional persist As Boolean = True)
            If minutesValue < 1 OrElse minutesValue > 120 OrElse idleValue < 1 OrElse idleValue > 120 Then Throw New ArgumentOutOfRangeException("minutes")
            Initialise()
            If persist Then
                Dim s = CheckSheetWatchSettings.Instance
                s.Enabled = enabledValue : s.Minutes = minutesValue : s.IdleMinutes = idleValue : s.Benchmark = benchmarkValue
                s.Save()
            End If
            Enabled = enabledValue : Minutes = minutesValue : IdleMinutes = idleValue : Benchmark = benchmarkValue
            For Each state In Plans.Values
                state.Due = DateTime.UtcNow.AddMinutes(Minutes)
            Next
        End Sub
        Friend Shared Sub Track(model As FileManager.ExcelModel)
            Initialise()
            Plans(model) = New Schedule With {.Due = DateTime.UtcNow.AddMinutes(Minutes)}
        End Sub
        Friend Shared Sub Forget(model As FileManager.ExcelModel)
            Plans.Remove(model)
        End Sub
        Friend Shared Function GetFullCalculationCount(model As FileManager.ExcelModel) As Integer
            Dim state As Schedule = Nothing
            Return If(Plans.TryGetValue(model, state), state.FullCalculationCount, 0)
        End Function

        'Explicit Check Sheet reads and successful history/edit completion share
        'the watcher calculation, but never turn an ordinary hidden-sheet edit
        'into a full calculation. Accepted revisions avoid repeated navigation work.
        Friend Shared Sub EnsureCurrent(model As FileManager.ExcelModel)
            If model Is Nothing OrElse model.IsClosing OrElse model.WB Is Nothing OrElse
               model.ModelSpreadsheetControl Is Nothing OrElse model.ModelSpreadsheetControl.IsDisposed OrElse
               model.ModelSpreadsheetControl.InvokeRequired OrElse Running.Contains(model) OrElse
               FileManager.BIsSaving OrElse ModelSafetyManager.IsBulkWorkbookMutationInProgress(model.ModelID) OrElse
               (model.ChangeManager IsNot Nothing AndAlso (model.ChangeManager.ChangeInProgress OrElse model.ChangeManager.IsReadOnlyPreview)) Then Return
            If model.LastAcceptedCheckSheetRevision = model.CalculationRevision Then Return
            Try
                RunCheck(model)
            Catch ex As Exception
                'A display/check failure after a committed edit must never roll
                'back that edit while leaving its history group committed.
                SystemMessageManager.Publish(model.ModelID,
                    "Check Sheet could not be refreshed. Please run Integrity to check the latest figures. " & ex.Message,
                    SystemMessageSeverity.Warning, "Check Sheet", model.FileName)
            End Try
        End Sub

        Friend Shared Sub RefreshVisibleAfterCommittedChange(model As FileManager.ExcelModel)
            If model Is Nothing OrElse model.WBCalcEngine Is Nothing OrElse
               Not model.WBCalcEngine.HasVisibleCheckSheet Then Return
            model.WBCalcEngine.QueueVisibleCheckSheetRead()
        End Sub
        Friend Shared Function ProcessIdle(now As DateTime, idle As TimeSpan) As Boolean
            Initialise()
            If Not Enabled OrElse idle.TotalMinutes < IdleMinutes OrElse FileManager.BIsSaving OrElse
                IdleIntegrityManager.OperationInProgress OrElse FormSplashScreen.OperationInProgress OrElse
                RecoveryBackupManager.AdvanceNoticeVisible OrElse RecoveryBackupManager.PendingEditor() Then Return False
            If FileManager.ExcelModels IsNot Nothing AndAlso FileManager.ExcelModels.Any(Function(m) m IsNot Nothing AndAlso
                (ModelSafetyManager.IsBulkWorkbookMutationInProgress(m.ModelID) OrElse
                 (m.ChangeManager IsNot Nothing AndAlso m.ChangeManager.ChangeInProgress))) Then Return False
            For Each pair In Plans.ToArray()
                Dim model = pair.Key, state = pair.Value
                If model.IsClosing OrElse model.WB Is Nothing Then
                    Forget(model) : Continue For
                End If
                If now < state.Due OrElse state.Revision = model.CalculationRevision OrElse
                    model.LastAcceptedCheckSheetRevision = model.CalculationRevision OrElse model.RecoverySaveAsRequired OrElse
                    model.IntegrityState <> ModelIntegrityState.Healthy OrElse model.ModelSpreadsheetControl Is Nothing OrElse
                    model.ModelSpreadsheetControl.IsDisposed OrElse model.ModelSpreadsheetControl.InvokeRequired OrElse
                    (model.ChangeManager IsNot Nothing AndAlso model.ChangeManager.IsReadOnlyPreview) OrElse
                    RecoveryBackupManager.NoticeOwner(model) Is Nothing Then Continue For
                state.Due = now.AddMinutes(Minutes)
                Try
                    RunCheck(model)
                    If model.LastAcceptedCheckSheetRevision = model.CalculationRevision Then
                        state.Revision = model.LastAcceptedCheckSheetRevision
                    End If
                    state.LastError = Nothing
                Catch ex As Exception
                    If state.LastError <> ex.Message Then SystemMessageManager.Publish(model.ModelID,
                        "Check Sheet could not be checked. " & ex.Message, SystemMessageSeverity.Warning, "Check Sheet", model.FileName)
                    state.LastError = ex.Message
                End Try
                Return True
            Next
            Return False
        End Function
        Friend Shared Sub RunCheck(model As FileManager.ExcelModel)
            If model Is Nothing OrElse model.IsClosing OrElse model.WB Is Nothing OrElse Running.Contains(model) OrElse
               model.ModelSpreadsheetControl Is Nothing OrElse model.ModelSpreadsheetControl.IsDisposed OrElse
               model.ModelSpreadsheetControl.InvokeRequired OrElse FileManager.BIsSaving OrElse
               ModelSafetyManager.IsBulkWorkbookMutationInProgress(model.ModelID) OrElse
               (model.ChangeManager IsNot Nothing AndAlso (model.ChangeManager.ChangeInProgress OrElse model.ChangeManager.IsReadOnlyPreview)) Then Return
            Dim sheet = model.WB.Worksheets.FirstOrDefault(Function(s) s.Name.Equals("Check Sheet", StringComparison.OrdinalIgnoreCase))
            If sheet Is Nothing Then Throw New InvalidOperationException("This model has no Check Sheet worksheet.")
            Dim revision = model.CalculationRevision
            Dim timer = SummitDiagnostics.DiagnosticTimer.StartNew(Benchmark)
            Dim outcome = "failed"
            Dim previousEngine = model.WB.Options.CalculationEngineType
            Dim previousMode = model.WB.Options.CalculationMode
            Dim service = model.WBCalculationService
            Dim previousSkip = If(service Is Nothing, False, service.DontCalcTDBS)
            Running.Add(model)
            Try
                Using notice As New FormSplashScreen(RecoveryBackupManager.NoticeOwner(model), "Checking plan balances", "Calculating the workbook before reading Check Sheet…")
                    'Check Sheet can depend on cached results elsewhere. The
                    'custom workbook-chain deferral also skips Check Sheet cells;
                    'a standalone sheet pass can instead consume external caches.
                    'Force the full workbook, including deferred sheets, without
                    'performing a dependency rebuild or certifying integrity.
                    Try
                        model.WB.Options.CalculationMode = WorkbookCalculationMode.Manual
                        If service IsNot Nothing Then service.DontCalcTDBS = False
                        model.WB.Options.CalculationEngineType = CalculationEngineType.Recursive
                        Dim state As Schedule = Nothing
                        If Not Plans.TryGetValue(model, state) Then
                            state = New Schedule With {.Due = DateTime.UtcNow.AddMinutes(Minutes)}
                            Plans.Add(model, state)
                        End If
                        state.FullCalculationCount += 1
                        model.WB.CalculateFull()
                    Finally
                        Try
                            model.WB.Options.CalculationEngineType = previousEngine
                        Finally
                            Try
                                model.WB.Options.CalculationMode = previousMode
                            Finally
                                If service IsNot Nothing Then service.DontCalcTDBS = previousSkip
                            End Try
                        End Try
                    End Try
                    Dim result = model.ReadCheckSheetValidation()
                    If Not String.IsNullOrWhiteSpace(result.ValidationError) Then Throw New InvalidOperationException(result.ValidationError)
                    'The worksheetOnly flag limits certification to soft Check Sheet
                    'state; pending rebuild/formula-integrity flags are not cleared.
                    Dim statusChanged As Boolean
                    If Not model.TryRecordCheckSheetResult(revision, result.HasFailures, statusChanged, worksheetOnly:=True) Then
                        outcome = "superseded"
                        Return
                    End If
                    If statusChanged Then
                        RecoveryBackupManager.NotifyCheckSheetState(model, result, Nothing)
                    End If
                    outcome = If(result.HasFailures, "unbalanced", "balanced")
                End Using
            Finally
                Running.Remove(model)
                SummitDiagnostics.WriteTrialLine("[Check Sheet Trial Benchmark] model=" & model.ModelID.ToString() &
                    ", mode=full, calculationAndRead=" & timer.ElapsedMilliseconds.ToString() & " ms, outcome=" & outcome)
            End Try
        End Sub
    End Class
End Namespace
