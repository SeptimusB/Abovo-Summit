Imports System.Configuration
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
Imports System.Windows.Forms
Imports System.Xml
Imports System.Xml.Linq
Imports DevExpress.XtraEditors

Namespace Abovo
    Friend NotInheritable Class RecoveryBackupSettings
        Inherits ApplicationSettingsBase
        Friend Shared ReadOnly Instance As New RecoveryBackupSettings()
        <UserScopedSetting(), DefaultSettingValue("False")>
        Public Property ContinueOnCheckSheetError As Boolean
            Get
                Return CBool(Me(NameOf(ContinueOnCheckSheetError)))
            End Get
            Set(value As Boolean)
                Me(NameOf(ContinueOnCheckSheetError)) = value
            End Set
        End Property
        <UserScopedSetting(), DefaultSettingValue("False")>
        Public Property Enabled As Boolean
            Get
                Return CBool(Me(NameOf(Enabled)))
            End Get
            Set(value As Boolean)
                Me(NameOf(Enabled)) = value
            End Set
        End Property
        <UserScopedSetting(), DefaultSettingValue("10")>
        Public Property Minutes As Integer
            Get
                Return CInt(Me(NameOf(Minutes)))
            End Get
            Set(value As Integer)
                Me(NameOf(Minutes)) = value
            End Set
        End Property
        <UserScopedSetting(), DefaultSettingValue("True")>
        Public Property WhenIdle As Boolean
            Get
                Return CBool(Me(NameOf(WhenIdle)))
            End Get
            Set(value As Boolean)
                Me(NameOf(WhenIdle)) = value
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
        Public Property AlwaysEvery As Boolean
            Get
                Return CBool(Me(NameOf(AlwaysEvery)))
            End Get
            Set(value As Boolean)
                Me(NameOf(AlwaysEvery)) = value
            End Set
        End Property
        <UserScopedSetting(), DefaultSettingValue("")>
        Public Property CheckSheetSuspendedFiles As String
            Get
                Return CStr(Me(NameOf(CheckSheetSuspendedFiles)))
            End Get
            Set(value As String)
                Me(NameOf(CheckSheetSuspendedFiles)) = value
            End Set
        End Property
    End Class

    'One completed recovery per saved model, in its existing folder. No Excel
    'installation, macro execution, original-file writes or metadata sheets.
    Friend NotInheritable Class RecoveryBackupStore
        Friend Const SourceProperty As String = "Abovo.Summit.RecoverySource"
        Friend Const VersionProperty As String = "Abovo.Summit.RecoveryVersion"
        Friend Const DateProperty As String = "Abovo.Summit.RecoveryCreatedUtc"
        Friend Const PendingProperty As String = "Abovo.Summit.ResultsPending"

        Friend Shared Function RecoveryPath(source As String) As String
            Return Path.Combine(Path.GetDirectoryName(Path.GetFullPath(source)),
                                "~" & Path.GetFileNameWithoutExtension(source) & "_recovery.xlsm")
        End Function

        Private Shared Function LegacyRecoveryPath(source As String) As String
            Return Path.Combine(Path.GetDirectoryName(Path.GetFullPath(source)),
                                Path.GetFileNameWithoutExtension(source) & "_recovery.xlsm")
        End Function

        Friend Shared Function ReadSource(filePath As String) As String
            If Not File.Exists(filePath) Then Return Nothing
            Using package = ZipFile.OpenRead(filePath)
                Dim part = package.GetEntry("docProps/custom.xml")
                If part Is Nothing Then Return Nothing
                Using content = part.Open(), reader = XmlReader.Create(content, New XmlReaderSettings With {
                    .DtdProcessing = DtdProcessing.Prohibit, .XmlResolver = Nothing, .MaxCharactersInDocument = 1048576})
                    Dim document = XDocument.Load(reader)
                    If document.Root Is Nothing Then Return Nothing
                    Dim properties = document.Root.Elements().
                        Where(Function(p) p.Attribute("name") IsNot Nothing).
                        ToDictionary(Function(p) CStr(p.Attribute("name")), Function(p) p.Value)
                    Dim origin As String = Nothing, version As String = Nothing
                    If Not properties.TryGetValue(VersionProperty, version) OrElse version <> "1" OrElse
                       Not properties.TryGetValue(SourceProperty, origin) OrElse String.IsNullOrWhiteSpace(origin) OrElse Not Path.IsPathRooted(origin) Then Return Nothing
                    Return Path.GetFullPath(origin)
                End Using
            End Using
        End Function

        Friend Shared Function RecoveryOrigin(file As String) As String
            Try
                Dim source = ReadSource(file)
                If source IsNot Nothing AndAlso
                   (String.Equals(RecoveryPath(source), Path.GetFullPath(file), StringComparison.OrdinalIgnoreCase) OrElse
                    String.Equals(LegacyRecoveryPath(source), Path.GetFullPath(file), StringComparison.OrdinalIgnoreCase)) Then Return source
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException OrElse TypeOf ex Is System.Xml.XmlException OrElse TypeOf ex Is ArgumentException
                Trace.WriteLine("[Recovery] Cannot read recovery metadata: " & ex.Message)
            End Try
            Return Nothing
        End Function

        Friend Shared Function NewerRecovery(source As String) As String
            If Not File.Exists(source) Then Return Nothing
            Dim newest As String = Nothing
            Dim newestTime = File.GetLastWriteTimeUtc(source)
            'Read earlier releases' filenames too. Never rename/delete an existing
            'recovery as part of discovery; new writes use only the prefixed path.
            For Each candidate In {RecoveryPath(source), LegacyRecoveryPath(source)}
                If Not File.Exists(candidate) Then Continue For
                Dim candidateTime = File.GetLastWriteTimeUtc(candidate)
                If candidateTime > newestTime AndAlso
                   String.Equals(RecoveryOrigin(candidate), Path.GetFullPath(source), StringComparison.OrdinalIgnoreCase) Then
                    newest = candidate
                    newestTime = candidateTime
                End If
            Next
            Return newest
        End Function

        Friend Shared Function SelectOpenPath(owner As IWin32Window, source As String) As String
            Try
                Dim candidate = NewerRecovery(source)
                If candidate Is Nothing Then Return source
                Dim answer = XtraMessageBox.Show(owner,
                    "A newer recovery copy is available." & Environment.NewLine & Environment.NewLine &
                    "Original: " & source & Environment.NewLine & "Saved: " & File.GetLastWriteTime(source).ToString("dd/MM/yyyy HH:mm:ss") &
                    Environment.NewLine & Environment.NewLine & "Recovery: " & candidate & Environment.NewLine &
                    "Saved: " & File.GetLastWriteTime(candidate).ToString("dd/MM/yyyy HH:mm:ss") & Environment.NewLine & Environment.NewLine &
                    "Open the recovery copy?" & Environment.NewLine &
                    "Yes: recover committed inputs, then use Save As to save an XLSB business plan. The original folder and filename will be suggested; replacing it requires confirmation." & Environment.NewLine &
                    "No: open the original. Cancel: do not open either file.",
                    "Newer recovery copy", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                If answer = DialogResult.Cancel Then Return Nothing
                Return If(answer = DialogResult.Yes, candidate, source)
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
                XtraMessageBox.Show(owner, "The recovery copy could not be checked. The original will be opened." & Environment.NewLine & ex.Message,
                                    "Recovery copy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return source
            End Try
        End Function

        Friend Shared Function Write(model As FileManager.ExcelModel) As String
            If model IsNot Nothing AndAlso model.RecoveryAutosaveSuspended Then
                Throw New InvalidOperationException("Recovery autosaves are paused because Check Sheet validation failed. Correct the problems and complete a fresh successful check first.")
            End If
            If model Is Nothing OrElse model.WB Is Nothing OrElse model.IsClosing OrElse Not model.HasUnsavedUserChanges OrElse
               model.IntegrityState <> ModelIntegrityState.Healthy OrElse model.RecoverySaveAsRequired OrElse
               (model.ChangeManager IsNot Nothing AndAlso (model.ChangeManager.ChangeInProgress OrElse model.ChangeManager.IsReadOnlyPreview)) OrElse
               ModelSafetyManager.IsBulkWorkbookMutationInProgress(model.ModelID) Then
                Throw New InvalidOperationException("The model is not ready for a recovery backup.")
            End If
            Dim source = Path.GetFullPath(model.FileName)
            If Not File.Exists(source) Then Throw New IOException("Save the business plan to a file before enabling recovery for it.")
            Dim destination = RecoveryPath(source)
            If FileManager.IsFileOpen(destination) Then Throw New IOException("The recovery copy is currently open in Summit; save that recovered plan as XLSB first.")
            If File.Exists(destination) AndAlso Not String.Equals(ReadSource(destination), source, StringComparison.OrdinalIgnoreCase) Then
                Throw New IOException("A file already uses the recovery filename but is not this plan's Summit recovery copy. It has not been overwritten.")
            End If
            Dim temporary = Path.Combine(Path.GetDirectoryName(destination), ".summit-recovery-" & Guid.NewGuid().ToString("N") & ".xlsm")
            Dim timer = Stopwatch.StartNew(), phaseTimer = Stopwatch.StartNew()
            Dim snapshotMs As Long, flushMs As Long, metadataMs As Long, historyMs As Long, verifyMs As Long, replaceMs As Long
            Dim phase As String = "snapshotExport", saved As Boolean = False
            Try
                Using output As New FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None)
                    model.WriteRecoverySnapshot(output)
                    snapshotMs = phaseTimer.ElapsedMilliseconds
                    phase = "flush" : phaseTimer.Restart()
                    output.Flush(True)
                End Using
                flushMs = phaseTimer.ElapsedMilliseconds
                phase = "metadata" : phaseTimer.Restart()
                RecoveryXlsmCompatibility.Prepare(temporary, model.RecoveryHasVerifiedBinaryMetadata)
                metadataMs = phaseTimer.ElapsedMilliseconds
                phase = "history" : phaseTimer.Restart()
                If model.ChangeManager IsNot Nothing Then RecoveryHistoryStore.Write(temporary, model.ChangeManager)
                historyMs = phaseTimer.ElapsedMilliseconds
                phase = "verify" : phaseTimer.Restart()
                If Not String.Equals(ReadSource(temporary), source, StringComparison.OrdinalIgnoreCase) Then Throw New InvalidDataException("Recovery metadata verification failed.")
                verifyMs = phaseTimer.ElapsedMilliseconds
                phase = "replace" : phaseTimer.Restart()
                'Never truncate the last good copy. A failed write/replace keeps it.
                If File.Exists(destination) Then
                    If Not String.Equals(ReadSource(destination), source, StringComparison.OrdinalIgnoreCase) Then Throw New IOException("The recovery destination changed during saving.")
                    File.Replace(temporary, destination, Nothing)
                Else
                    File.Move(temporary, destination)
                End If
                replaceMs = phaseTimer.ElapsedMilliseconds
                saved = True
                'An explicitly permitted backup of known errors must retain its
                'own hold even if the unsaved original session is later discarded.
                If model.CheckSheetWarningActive Then RecoveryBackupManager.PersistCheckSheetPausePath(destination, True)
                Return destination
            Finally
                Trace.WriteLine("[Recovery Packaging Benchmark] model=" & model.ModelID.ToString() &
                    ", snapshotExport=" & snapshotMs.ToString() & " ms, flush=" & flushMs.ToString() &
                    " ms, metadata=" & metadataMs.ToString() & " ms, history=" & historyMs.ToString() &
                    " ms, verify=" & verifyMs.ToString() & " ms, replace=" & replaceMs.ToString() &
                    " ms, total=" & timer.ElapsedMilliseconds.ToString() & " ms, outcome=" & If(saved, "ok", "failed") &
                    If(saved, String.Empty, ", failedStage=" & phase & ", stageElapsed=" & phaseTimer.ElapsedMilliseconds.ToString() & " ms"))
                If File.Exists(temporary) Then File.Delete(temporary) 'Only this invocation's private, incomplete output.
            End Try
        End Function
    End Class

    'This is deliberately separate from the threaded, non-interactive wait form.
    'No workbook writing occurs while this native UI-thread notice is visible.
    Friend NotInheritable Class RecoverySaveNotice
        Inherits XtraForm
        Implements IMessageFilter
        Friend Event SnoozeRequested As EventHandler
        Private dismissing As Boolean
        Private escapeFilterInstalled As Boolean
        Friend Property ClosingNow As Boolean

        Friend Sub New(fileName As String)
            Text = "Recovery save"
            ShowInTaskbar = False
            FormBorderStyle = FormBorderStyle.FixedToolWindow
            MaximizeBox = False : MinimizeBox = False
            StartPosition = FormStartPosition.Manual
            AutoScaleMode = AutoScaleMode.Dpi
            ClientSize = New System.Drawing.Size(480, 170)
            Padding = New Padding(14)
            Dim layout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2}
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Dim description As New LabelControl With {.Dock = DockStyle.Top, .AutoSizeMode = LabelAutoSizeMode.Vertical,
                .Text = "Recovery save in about 5 seconds" & Environment.NewLine & fileName & Environment.NewLine & Environment.NewLine &
                        "The original is unchanged. Press Esc to cancel this save and snooze until idle for 1 minute."}
            Dim snooze As New SimpleButton With {.Name = "SnoozeRecovery", .Text = "Snooze until idle for 1 minute", .AutoSize = True,
                .Anchor = AnchorStyles.Left, .Margin = New Padding(3, 12, 3, 3)}
            AddHandler snooze.Click, Sub() RaiseEvent SnoozeRequested(Me, EventArgs.Empty)
            layout.Controls.Add(description, 0, 0)
            layout.Controls.Add(snooze, 0, 1)
            Controls.Add(layout)
            AddHandler Shown, Sub()
                                  ClientSize = New System.Drawing.Size(Math.Max(ClientSize.Width, snooze.Right + Padding.Right + 20),
                                      layout.PreferredSize.Height + Padding.Vertical)
                                  If Owner IsNot Nothing Then Location = New System.Drawing.Point(
                                      Math.Max(Screen.FromControl(Owner).WorkingArea.Left, Owner.Right - Width - 18),
                                      Math.Max(Screen.FromControl(Owner).WorkingArea.Top, Owner.Top + 45))
                              End Sub
        End Sub

        Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
            Get
                Return True
            End Get
        End Property

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            'The notice deliberately does not take focus. Scope the shortcut to
            'this UI thread while it is visible, so Esc also works in the grid.
            If Not IsDisposed AndAlso Visible AndAlso Not escapeFilterInstalled Then
                Application.AddMessageFilter(Me)
                escapeFilterInstalled = True
            End If
        End Sub

        Public Function PreFilterMessage(ByRef message As Message) As Boolean Implements IMessageFilter.PreFilterMessage
            If Not escapeFilterInstalled OrElse IsDisposed OrElse Not Visible OrElse ClosingNow OrElse dismissing OrElse
               message.Msg <> &H100 OrElse message.WParam.ToInt64() <> CInt(Keys.Escape) OrElse ModifierKeys <> Keys.None Then Return False
            'Do not intercept other applications or Esc in an unrelated modal
            'dialog. Consume it here so it does not also undo the grid editor.
            If Control.FromChildHandle(message.HWnd) Is Nothing OrElse Application.OpenForms.Cast(Of Form)().Any(Function(f) f.Modal) Then Return False
            RaiseEvent SnoozeRequested(Me, EventArgs.Empty)
            Return True
        End Function

        Private Sub RemoveEscapeFilter()
            If Not escapeFilterInstalled Then Return
            Application.RemoveMessageFilter(Me)
            escapeFilterInstalled = False
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            RemoveEscapeFilter()
            ClosingNow = True
            If Not dismissing Then RaiseEvent SnoozeRequested(Me, EventArgs.Empty)
            MyBase.OnFormClosing(e)
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then RemoveEscapeFilter()
            MyBase.Dispose(disposing)
        End Sub

        Friend Sub Dismiss()
            dismissing = True
            Close()
            Dispose()
        End Sub
    End Class

    Friend NotInheritable Class RecoveryBackupManager
        Private Class Schedule
            Friend DueUtc As DateTime
            Friend IdleDueUtc As DateTime
            Friend PromptPreviousCheck As Boolean
            Friend Snoozed As Boolean
            Friend SnoozeNotBeforeUtc As DateTime
            Friend NoticeReadyUtc As DateTime
            Friend NoticeRevision As Long
            Friend Revision As Long = -1
            Friend Path As String
        End Class
        Private Shared ReadOnly Plans As New Dictionary(Of FileManager.ExcelModel, Schedule)()
        Private Shared ReadOnly Clock As New Windows.Forms.Timer With {.Interval = 5000}
        Private Shared Initialised As Boolean
        Private Shared Busy As Boolean
        Private Shared PendingModel As FileManager.ExcelModel
        Private Shared PendingNotice As RecoverySaveNotice
        Private Shared PausePersistencePending As Boolean
        Friend Shared Property Enabled As Boolean
        Friend Shared Property Minutes As Integer = 10
        Friend Shared Property WhenIdle As Boolean = True
        Friend Shared Property IdleMinutes As Integer = 2
        Friend Shared Property AlwaysEvery As Boolean
        Friend Shared Property ContinueOnCheckSheetError As Boolean
        Friend Shared ReadOnly Property AdvanceNoticeVisible As Boolean
            Get
                Return PendingNotice IsNot Nothing AndAlso Not PendingNotice.IsDisposed AndAlso PendingNotice.Visible
            End Get
        End Property

        Friend Shared Sub Initialise()
            If Initialised Then Return
            Initialised = True
            Try
                Enabled = RecoveryBackupSettings.Instance.Enabled
                ContinueOnCheckSheetError = RecoveryBackupSettings.Instance.ContinueOnCheckSheetError
                Minutes = Math.Max(1, Math.Min(120, RecoveryBackupSettings.Instance.Minutes))
                WhenIdle = RecoveryBackupSettings.Instance.WhenIdle
                IdleMinutes = Math.Max(1, Math.Min(120, RecoveryBackupSettings.Instance.IdleMinutes))
                AlwaysEvery = RecoveryBackupSettings.Instance.AlwaysEvery
            Catch ex As ConfigurationErrorsException
                Trace.WriteLine("[Recovery] Settings unavailable: " & ex.Message)
            End Try
            AddHandler Clock.Tick, AddressOf Tick
            AddHandler Application.ApplicationExit, Sub()
                                                        Clock.Stop()
                                                        CancelPendingNotice()
                                                    End Sub
            Clock.Start()
        End Sub

        Friend Shared Sub Configure(enableBackup As Boolean, interval As Integer, Optional persist As Boolean = True)
            If interval < 1 OrElse interval > 120 Then Throw New ArgumentOutOfRangeException(NameOf(interval))
            Initialise()
            If persist Then
                RecoveryBackupSettings.Instance.Enabled = enableBackup
                RecoveryBackupSettings.Instance.Minutes = interval
                RecoveryBackupSettings.Instance.Save()
            End If
            Enabled = enableBackup
            Minutes = interval
            CancelPendingNotice()
            For Each entry In Plans.Values
                entry.DueUtc = DateTime.UtcNow.AddMinutes(Minutes)
                entry.IdleDueUtc = DateTime.UtcNow.AddMinutes(IdleMinutes)
            Next
        End Sub

        Friend Shared Sub ConfigureTiming(whenIdleEnabled As Boolean, idleInterval As Integer,
                                          alwaysEnabled As Boolean, interval As Integer,
                                          Optional persist As Boolean = True)
            If idleInterval < 1 OrElse idleInterval > 120 Then Throw New ArgumentOutOfRangeException(NameOf(idleInterval))
            If interval < 1 OrElse interval > 120 Then Throw New ArgumentOutOfRangeException(NameOf(interval))
            If Not whenIdleEnabled AndAlso Not alwaysEnabled Then Throw New ArgumentException("Choose when idle, always every, or both recovery timings.")
            Initialise()
            If persist Then
                RecoveryBackupSettings.Instance.WhenIdle = whenIdleEnabled
                RecoveryBackupSettings.Instance.IdleMinutes = idleInterval
                RecoveryBackupSettings.Instance.AlwaysEvery = alwaysEnabled
                RecoveryBackupSettings.Instance.Minutes = interval
                RecoveryBackupSettings.Instance.Save()
            End If
            WhenIdle = whenIdleEnabled : IdleMinutes = idleInterval
            AlwaysEvery = alwaysEnabled : Minutes = interval
            CancelPendingNotice()
            For Each state In Plans.Values
                state.DueUtc = DateTime.UtcNow.AddMinutes(Minutes)
                state.IdleDueUtc = DateTime.UtcNow.AddMinutes(IdleMinutes)
            Next
        End Sub

        Friend Shared Sub Track(model As FileManager.ExcelModel)
            Initialise()
            Forget(model)
            RestoreCheckSheetPause(model)
            Plans(model) = New Schedule With {.DueUtc = DateTime.UtcNow.AddMinutes(Minutes),
                .IdleDueUtc = DateTime.UtcNow.AddMinutes(IdleMinutes), .Path = model.FileName,
                .PromptPreviousCheck = model.CheckSheetWarningNeedsRecheck}
        End Sub

        Friend Shared Sub ConfigureCheckSheetPolicy(continueAutosave As Boolean, Optional persist As Boolean = True)
            Initialise()
            If persist Then
                RecoveryBackupSettings.Instance.ContinueOnCheckSheetError = continueAutosave
                RecoveryBackupSettings.Instance.Save()
            End If
            ContinueOnCheckSheetError = continueAutosave
            'This is permission to back up known errors, not a successful check.
            'Retain both the per-file failure record and the visible warning.
        End Sub

        Friend Shared Sub Forget(model As FileManager.ExcelModel)
            If PendingModel Is model Then CancelPendingNotice()
            Plans.Remove(model)
        End Sub

        Private Shared Sub CancelPendingNotice()
            Dim notice = PendingNotice
            PendingNotice = Nothing : PendingModel = Nothing
            If notice IsNot Nothing AndAlso Not notice.IsDisposed AndAlso Not notice.ClosingNow Then notice.Dismiss()
        End Sub

        Private Shared Sub Snooze(model As FileManager.ExcelModel)
            Dim state As Schedule = Nothing
            If Not Plans.TryGetValue(model, state) Then Return
            state.Snoozed = True
            state.SnoozeNotBeforeUtc = DateTime.UtcNow.AddMinutes(1)
            CancelPendingNotice()
            SystemMessageManager.Publish(model.ModelID,
                "Recovery save snoozed until one full minute without keyboard or mouse input. The pending save was cancelled; the previous recovery copy is unchanged.",
                SystemMessageSeverity.Information, "Recovery backup", model.FileName)
        End Sub

        Private Shared Function ReadCheckSheetPauses() As HashSet(Of String)
            Dim paths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each entry In If(RecoveryBackupSettings.Instance.CheckSheetSuspendedFiles, "").
                Split({ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
                If Not IO.Path.IsPathRooted(entry) Then Throw New ConfigurationErrorsException("A recovery pause has an invalid file path.")
                paths.Add(IO.Path.GetFullPath(entry))
            Next
            Return paths
        End Function

        Friend Shared Sub RestoreCheckSheetPause(model As FileManager.ExcelModel,
                                                Optional reportHold As Boolean = True)
            If String.IsNullOrWhiteSpace(model.FileName) Then Return
            Try
                Dim paths = ReadCheckSheetPauses()
                Dim originalHeld = Not String.IsNullOrWhiteSpace(model.RecoverySourcePath) AndAlso paths.Contains(IO.Path.GetFullPath(model.RecoverySourcePath))
                model.RestoreRecoveryAutosaveHold(PausePersistencePending OrElse originalHeld OrElse paths.Contains(IO.Path.GetFullPath(model.FileName)))
                If reportHold AndAlso model.CheckSheetWarningActive Then
                    SystemMessageManager.Publish(model.ModelID,
                        "This file has an unresolved previous Check Sheet validation. " & CheckSheetPolicyDescription(model) & " Click the red (Check sheet) indicator beside the company name to review it. A fresh successful check clears the warning.",
                        SystemMessageSeverity.Warning, "Recovery backup", model.FileName)
                End If
            Catch ex As Exception
                model.RestoreRecoveryAutosaveHold(True)
                SystemMessageManager.Publish(model.ModelID,
                    "Recovery autosaves are paused because saved Check Sheet pause settings could not be read: " & ex.Message,
                    SystemMessageSeverity.Warning, "Recovery backup", model.FileName)
            End Try
        End Sub

        Friend Shared Function PersistCheckSheetPause(model As FileManager.ExcelModel, paused As Boolean) As Boolean
            If String.IsNullOrWhiteSpace(model.FileName) Then Return True 'Unsaved plans cannot autosave yet.
            Return PersistCheckSheetPausePath(model.FileName, paused, model.ModelID)
        End Function

        Friend Shared Function PersistCheckSheetPausePath(filePath As String, paused As Boolean,
                                                          Optional modelID As Integer = -1) As Boolean
            Try
                Dim paths = ReadCheckSheetPauses()
                Dim path = IO.Path.GetFullPath(filePath)
                Dim changed = If(paused, paths.Add(path), paths.Remove(path))
                If Not changed AndAlso Not PausePersistencePending Then Return True
                RecoveryBackupSettings.Instance.CheckSheetSuspendedFiles = String.Join(Environment.NewLine, paths.OrderBy(Function(p) p, StringComparer.OrdinalIgnoreCase))
                PausePersistencePending = True
                RecoveryBackupSettings.Instance.Save()
                PausePersistencePending = False
                Return True
            Catch ex As Exception
                PausePersistencePending = True
                SystemMessageManager.Publish(modelID,
                    "The recovery autosave pause could not be saved to application settings. Autosaves remain paused in this session; do not rely on the pause surviving a restart until settings can be saved. " & ex.Message,
                    SystemMessageSeverity.Error, "Recovery backup", filePath)
                Return False
            End Try
        End Function

        Friend Shared Sub NotifyCheckSheetState(model As FileManager.ExcelModel,
                                                result As FileManager.ExcelModel.CloseModelValidationResult,
                                                owner As Form)
            If model.CheckSheetWarningActive Then
                Dim details = result.ValidationError
                If result.Issues.Count > 0 Then
                    details = String.Join(Environment.NewLine, result.Issues.Take(5).
                        Select(Function(issue) "Row " & issue.CheckRow.ToString() & ": " & issue.Label & " - " & issue.Status & " " & issue.Message))
                    If result.Issues.Count > 5 Then details &= Environment.NewLine & "More findings are listed in the integrity report."
                End If
                Dim message = "Check Sheet needs attention for " & IO.Path.GetFileName(model.FileName) & ". " & CheckSheetPolicyDescription(model) &
                    Environment.NewLine & Environment.NewLine & details & Environment.NewLine & Environment.NewLine &
                    "Click the red (Check sheet) indicator beside the company name to review the Check Sheet. After correction, run a fresh integrity check to clear the warning." &
                    Environment.NewLine & "Normal Save and Save As remain available; they do not clear the warning."
                SystemMessageManager.Publish(model.ModelID, message, SystemMessageSeverity.Warning, "Recovery backup", model.FileName)
            Else
                SystemMessageManager.Publish(model.ModelID,
                    "Check Sheet validation now passes. Recovery autosaves are permitted again when enabled and new unsaved user changes are due.",
                    SystemMessageSeverity.Success, "Recovery backup", model.FileName)
            End If
        End Sub

        Friend Shared Function CheckSheetPolicyDescription(model As FileManager.ExcelModel) As String
            If model.RecoveryAutosaveSuspended Then Return "Recovery autosaves are PAUSED; the last completed recovery copy is retained."
            If model.CheckSheetWarningActive Then Return "Recovery autosaves may continue with these errors because 'Continue autosave if Check Sheet error?' is enabled."
            Return "Recovery autosaves are not paused by Check Sheet validation."
        End Function

        Friend Shared Sub OpenCheckSheet(model As FileManager.ExcelModel)
            If model Is Nothing OrElse model.IsClosing Then Return
            Try
                Dim target = model.WB.Worksheets.FirstOrDefault(Function(s) s.Name.Equals("Check Sheet", StringComparison.OrdinalIgnoreCase))
                If target Is Nothing Then Throw New InvalidOperationException("The workbook has no Check Sheet worksheet.")
                Dim links = ReadOnlyMappedTableGrid.FindInterfaces(model.WBStructure, target.Name, Nothing)
                If links.Count > 0 AndAlso model.WBInterface IsNot Nothing Then
                    Dim link = links(0)
                    Dim group = model.WBStructure.GroupStructures.First(Function(g) g.GSID = link.LinkGroupID.ToString())
                    model.WBInterface.ShowGroupInterface(model.ModelID, link.LinkGroupID, "Normal", group.GSName, model.InstanceInterface, link)
                Else
                    model.ShowSpreadsheet(target)
                End If
            Catch ex As Exception
                SystemMessageManager.Publish(model.ModelID, "Could not open Check Sheet: " & ex.Message, SystemMessageSeverity.Warning, "Recovery backup", model.FileName)
            End Try
        End Sub

        Friend Shared Function PendingEditor() As Boolean
            For Each form As Form In Application.OpenForms
                If form.Modal Then Return True
                Dim focus As Control = form
                While focus IsNot Nothing AndAlso focus.ContainsFocus
                    Dim editor = TryCast(focus, BaseEdit)
                    If editor IsNot Nothing AndAlso editor.IsModified AndAlso Not editor.Properties.ReadOnly Then Return True
                    Dim spreadsheet = TryCast(focus, DevExpress.XtraSpreadsheet.SpreadsheetControl)
                    If spreadsheet IsNot Nothing AndAlso spreadsheet.IsCellEditorActive Then Return True
                    focus = focus.Controls.Cast(Of Control)().FirstOrDefault(Function(c) c.ContainsFocus)
                End While
            Next
            Return False
        End Function

        Friend Shared Function NoticeOwner(model As FileManager.ExcelModel) As Form
            'ActiveForm is Nothing while another application has focus. Still use
            'a live Summit window; do not activate it or restore a minimised one.
            Dim eligible As Func(Of Form, Boolean) =
                Function(f) f IsNot Nothing AndAlso Not f.IsDisposed AndAlso f.IsHandleCreated AndAlso
                            Not f.InvokeRequired AndAlso f.Visible AndAlso f.WindowState <> FormWindowState.Minimized AndAlso
                            Not TypeOf f Is DevExpress.XtraWaitForm.WaitForm AndAlso Not TypeOf f Is RecoverySaveNotice
            If eligible(Form.ActiveForm) Then Return Form.ActiveForm
            If PendingNotice IsNot Nothing AndAlso Form.ActiveForm Is PendingNotice AndAlso eligible(PendingNotice.Owner) Then Return PendingNotice.Owner
            Dim modelForm = If(model.ModelSpreadsheetControl Is Nothing, Nothing, model.ModelSpreadsheetControl.FindForm())
            If eligible(modelForm) Then Return modelForm
            Return Application.OpenForms.Cast(Of Form)().FirstOrDefault(eligible)
        End Function

        Private Shared Sub Tick(sender As Object, e As EventArgs)
            'Session input includes mouse movement and activity in other apps.
            'If the input clock is unavailable, only the explicit maximum-interval
            'policy may run; never infer that the user is idle.
            Dim idle = TimeSpan.Zero
            Try
                idle = IntegrityInputClock.IdleDuration()
            Catch ex As Exception
                Trace.WriteLine("[Recovery] Idle clock unavailable: " & ex.Message)
            End Try
            ProcessRecovery(DateTime.UtcNow, idle)
        End Sub

        Friend Shared Sub ProcessRecovery(now As DateTime, idle As TimeSpan)
            If Busy Then Return
            If IdleIntegrityManager.OperationInProgress OrElse FormSplashScreen.OperationInProgress OrElse FileManager.BIsSaving OrElse PendingEditor() Then
                CancelPendingNotice()
                Return
            End If
            If FileManager.ExcelModels IsNot Nothing AndAlso FileManager.ExcelModels.Any(
                Function(m) m IsNot Nothing AndAlso (ModelSafetyManager.IsBulkWorkbookMutationInProgress(m.ModelID) OrElse
                    (m.ChangeManager IsNot Nothing AndAlso m.ChangeManager.ChangeInProgress))) Then
                CancelPendingNotice()
                Return
            End If
            For Each pair In Plans.ToArray()
                Dim model = pair.Key, state = pair.Value
                If model.IsClosing OrElse model.ModelSpreadsheetControl Is Nothing OrElse model.ModelSpreadsheetControl.IsDisposed Then
                    Forget(model)
                    Continue For
                End If
                If state.Path <> model.FileName Then
                    If PendingModel Is model Then CancelPendingNotice()
                    state.Path = model.FileName
                    state.Revision = -1
                    state.DueUtc = now.AddMinutes(Minutes)
                    state.IdleDueUtc = now.AddMinutes(IdleMinutes)
                End If
                If state.PromptPreviousCheck Then
                    If Not model.CheckSheetWarningNeedsRecheck Then
                        state.PromptPreviousCheck = False
                    Else
                        Dim promptOwner = NoticeOwner(model)
                        If promptOwner Is Nothing Then Continue For
                        CancelPendingNotice()
                        state.PromptPreviousCheck = False 'Once per opening, before modal message pumping.
                        Busy = True
                        Try
                            Dim answer = XtraMessageBox.Show(promptOwner,
                                "A previous session's Check Sheet validation failed for:" & Environment.NewLine & model.FileName &
                                Environment.NewLine & Environment.NewLine & "This is a remembered warning, not a new failure found in the reopened file." &
                                Environment.NewLine & CheckSheetPolicyDescription(model) & Environment.NewLine & Environment.NewLine &
                                "Would you like to run an integrity check now?" & Environment.NewLine &
                                "No leaves the warning in place; you can run the check later from Options.",
                                "Previous Check Sheet warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
                            If answer = DialogResult.Yes Then IdleIntegrityManager.RequestNow(model)
                        Finally
                            Busy = False
                        End Try
                        Return
                    End If
                End If
                If Not Enabled Then
                    If PendingModel Is model Then CancelPendingNotice()
                    Continue For
                End If
                Dim maximumDue = AlwaysEvery AndAlso now >= state.DueUtc
                Dim idleDue = WhenIdle AndAlso now >= state.IdleDueUtc AndAlso idle.TotalMinutes >= IdleMinutes
                Dim timingDue = If(state.Snoozed, now >= state.SnoozeNotBeforeUtc AndAlso idle.TotalMinutes >= 1, maximumDue OrElse idleDue)
                If Not timingDue OrElse model.RecoveryAutosaveSuspended OrElse Not model.HasUnsavedUserChanges OrElse model.UserChangeRevision = state.Revision OrElse
                   model.RecoverySaveAsRequired OrElse model.IntegrityState <> ModelIntegrityState.Healthy OrElse
                   (model.ChangeManager IsNot Nothing AndAlso (model.ChangeManager.ChangeInProgress OrElse model.ChangeManager.IsReadOnlyPreview)) OrElse
                   ModelSafetyManager.IsBulkWorkbookMutationInProgress(model.ModelID) Then
                    If PendingModel Is model Then CancelPendingNotice()
                    Continue For
                End If
                If PendingModel IsNot Nothing AndAlso PendingModel IsNot model Then Continue For
                Dim owner = NoticeOwner(model)
                If owner Is Nothing Then
                    If PendingModel Is model Then CancelPendingNotice()
                    Continue For 'Retry when a window is restored; never save silently.
                End If
                If PendingNotice IsNot Nothing AndAlso
                   (PendingNotice.IsDisposed OrElse Not PendingNotice.Visible OrElse PendingNotice.Owner IsNot owner OrElse state.NoticeRevision <> model.UserChangeRevision) Then CancelPendingNotice()
                If PendingNotice Is Nothing Then
                    Try
                        PendingModel = model
                        state.NoticeReadyUtc = now.AddSeconds(5)
                        state.NoticeRevision = model.UserChangeRevision
                        PendingNotice = New RecoverySaveNotice(IO.Path.GetFileName(model.FileName))
                        AddHandler PendingNotice.SnoozeRequested, Sub() Snooze(model)
                        PendingNotice.Show(owner)
                    Catch ex As Exception
                        CancelPendingNotice()
                        SystemMessageManager.Publish(model.ModelID, "Recovery notice could not be shown; no backup was written: " & ex.Message,
                            SystemMessageSeverity.Warning, "Recovery backup", model.FileName)
                        state.DueUtc = now.AddMinutes(Minutes) : state.IdleDueUtc = now.AddMinutes(IdleMinutes)
                        state.SnoozeNotBeforeUtc = now.AddMinutes(1)
                    End Try
                    Return 'Keep pumping normally; never sleep or export during the notice.
                End If
                If now < state.NoticeReadyUtc Then Return
                Dim afterSnooze = state.Snoozed
                CancelPendingNotice()
                Busy = True
                Dim timer = Stopwatch.StartNew()
                Try
                    Using notice As New FormSplashScreen(owner, "Saving recovery copy",
                                                        IO.Path.GetFileName(model.FileName) & Environment.NewLine &
                                                        "Saving committed inputs to XLSM. The original business plan is unchanged; please wait...")
                        If Not notice.IsShowing Then Throw New InvalidOperationException("The recovery progress notice could not be displayed; saving has been deferred.")
                        SystemMessageManager.Publish(model.ModelID, "Saving recovery copy: " & IO.Path.GetFileName(model.FileName) &
                            If(afterSnooze, " (one minute idle after snooze).", If(maximumDue, " (maximum interval reached).", " (idle interval reached).")) & " The original plan will not be changed.", SystemMessageSeverity.Information, "Recovery backup", model.FileName)
                        Try
                            Dim saved = RecoveryBackupStore.Write(model)
                            state.Revision = model.UserChangeRevision
                            notice.Complete("Recovery copy saved. Normal Save still writes your XLSB plan.")
                            SystemMessageManager.Publish(model.ModelID, "Recovery copy saved: " & saved & ". The original plan and unsaved changes are unchanged.", SystemMessageSeverity.Success, "Recovery backup", saved)
                        Catch
                            notice.Fail("Recovery copy could not be saved. The original and previous recovery copy are unchanged. See System Messages for details.")
                            Throw
                        End Try
                    End Using
                Catch ex As Exception
                    SystemMessageManager.Publish(model.ModelID, "Recovery backup failed; the original and previous completed recovery were retained. " & ex.Message,
                                                 SystemMessageSeverity.Warning, "Recovery backup", model.FileName)
                Finally
                    state.DueUtc = DateTime.UtcNow.AddMinutes(Minutes)
                    state.IdleDueUtc = DateTime.UtcNow.AddMinutes(IdleMinutes)
                    state.Snoozed = False
                    Busy = False
                    Trace.WriteLine("[Recovery Save Benchmark] model=" & model.ModelID.ToString() & ", total=" & timer.ElapsedMilliseconds.ToString() & " ms")
                End Try
                Exit For 'Never serialise several large plans in one timer tick.
            Next
        End Sub
    End Class
End Namespace
