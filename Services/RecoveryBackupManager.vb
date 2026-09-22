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
                If source IsNot Nothing AndAlso String.Equals(RecoveryPath(source), Path.GetFullPath(file), StringComparison.OrdinalIgnoreCase) Then Return source
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException OrElse TypeOf ex Is System.Xml.XmlException OrElse TypeOf ex Is ArgumentException
                Trace.WriteLine("[Recovery] Cannot read recovery metadata: " & ex.Message)
            End Try
            Return Nothing
        End Function

        Friend Shared Function NewerRecovery(source As String) As String
            Dim candidate = RecoveryPath(source)
            If Not File.Exists(source) OrElse Not File.Exists(candidate) Then Return Nothing
            If File.GetLastWriteTimeUtc(candidate) <= File.GetLastWriteTimeUtc(source) Then Return Nothing
            If String.Equals(RecoveryOrigin(candidate), Path.GetFullPath(source), StringComparison.OrdinalIgnoreCase) Then Return candidate
            Return Nothing
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
            If model Is Nothing OrElse model.WB Is Nothing OrElse model.IsClosing OrElse Not model.IsDirty OrElse
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
            Try
                Using output As New FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None)
                    model.WriteRecoverySnapshot(output)
                    output.Flush(True)
                End Using
                RecoveryXlsmCompatibility.Prepare(temporary, model.RecoveryHasVerifiedBinaryMetadata)
                If model.ChangeManager IsNot Nothing Then RecoveryHistoryStore.Write(temporary, model.ChangeManager)
                If Not String.Equals(ReadSource(temporary), source, StringComparison.OrdinalIgnoreCase) Then Throw New InvalidDataException("Recovery metadata verification failed.")
                'Never truncate the last good copy. A failed write/replace keeps it.
                If File.Exists(destination) Then
                    If Not String.Equals(ReadSource(destination), source, StringComparison.OrdinalIgnoreCase) Then Throw New IOException("The recovery destination changed during saving.")
                    File.Replace(temporary, destination, Nothing)
                Else
                    File.Move(temporary, destination)
                End If
                Return destination
            Finally
                If File.Exists(temporary) Then File.Delete(temporary) 'Only this invocation's private, incomplete output.
            End Try
        End Function
    End Class

    Friend NotInheritable Class RecoveryBackupManager
        Implements IMessageFilter
        Private Class Schedule
            Friend DueUtc As DateTime
            Friend Revision As Long = -1
            Friend Path As String
        End Class
        Private Shared ReadOnly Plans As New Dictionary(Of FileManager.ExcelModel, Schedule)()
        Private Shared ReadOnly Clock As New Windows.Forms.Timer With {.Interval = 5000}
        Private Shared ReadOnly InputObserver As New RecoveryBackupManager()
        Private Shared Initialised As Boolean
        Private Shared Busy As Boolean
        Private Shared LastInputUtc As DateTime = DateTime.UtcNow
        Friend Shared Property Enabled As Boolean
        Friend Shared Property Minutes As Integer = 10

        Friend Shared Sub Initialise()
            If Initialised Then Return
            Initialised = True
            Try
                Enabled = RecoveryBackupSettings.Instance.Enabled
                Minutes = Math.Max(1, Math.Min(120, RecoveryBackupSettings.Instance.Minutes))
            Catch ex As ConfigurationErrorsException
                Trace.WriteLine("[Recovery] Settings unavailable: " & ex.Message)
            End Try
            Application.AddMessageFilter(InputObserver)
            AddHandler Clock.Tick, AddressOf Tick
            AddHandler Application.ApplicationExit, Sub() Clock.Stop()
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
            For Each entry In Plans.Values
                entry.DueUtc = DateTime.UtcNow.AddMinutes(Minutes)
            Next
        End Sub

        Friend Shared Sub Track(model As FileManager.ExcelModel)
            Initialise()
            Plans(model) = New Schedule With {.DueUtc = DateTime.UtcNow.AddMinutes(Minutes), .Path = model.FileName}
        End Sub

        Friend Shared Sub Forget(model As FileManager.ExcelModel)
            Plans.Remove(model)
        End Sub

        Public Function PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
            If m.Msg = &H100 OrElse m.Msg = &H104 OrElse m.Msg = &H102 OrElse m.Msg = &H201 OrElse m.Msg = &H204 OrElse m.Msg = &H20A Then LastInputUtc = DateTime.UtcNow
            Return False
        End Function

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
                            Not TypeOf f Is DevExpress.XtraWaitForm.WaitForm
            If eligible(Form.ActiveForm) Then Return Form.ActiveForm
            Dim modelForm = If(model.ModelSpreadsheetControl Is Nothing, Nothing, model.ModelSpreadsheetControl.FindForm())
            If eligible(modelForm) Then Return modelForm
            Return Application.OpenForms.Cast(Of Form)().FirstOrDefault(eligible)
        End Function

        Private Shared Sub Tick(sender As Object, e As EventArgs)
            If Not Enabled OrElse Busy OrElse IdleIntegrityManager.OperationInProgress OrElse FormSplashScreen.OperationInProgress OrElse FileManager.BIsSaving OrElse
               (DateTime.UtcNow - LastInputUtc).TotalSeconds < 15 OrElse PendingEditor() Then Return
            For Each pair In Plans.ToArray()
                Dim model = pair.Key, state = pair.Value
                If model.IsClosing OrElse model.ModelSpreadsheetControl Is Nothing OrElse model.ModelSpreadsheetControl.IsDisposed Then
                    Forget(model)
                    Continue For
                End If
                If state.Path <> model.FileName Then
                    state.Path = model.FileName
                    state.Revision = -1
                    state.DueUtc = DateTime.UtcNow.AddMinutes(Minutes)
                End If
                If DateTime.UtcNow < state.DueUtc OrElse Not model.IsDirty OrElse model.CalculationRevision = state.Revision OrElse
                   model.RecoverySaveAsRequired OrElse model.IntegrityState <> ModelIntegrityState.Healthy OrElse
                   (model.ChangeManager IsNot Nothing AndAlso (model.ChangeManager.ChangeInProgress OrElse model.ChangeManager.IsReadOnlyPreview)) OrElse
                   ModelSafetyManager.IsBulkWorkbookMutationInProgress(model.ModelID) Then Continue For
                Dim owner = NoticeOwner(model)
                If owner Is Nothing Then Continue For 'Retry when a window is restored; never save silently.
                Busy = True
                Dim timer = Stopwatch.StartNew()
                Try
                    Using notice As New FormSplashScreen(owner, "Saving recovery copy",
                                                        IO.Path.GetFileName(model.FileName) & Environment.NewLine &
                                                        "Saving committed inputs to XLSM. The original business plan is unchanged; please wait...")
                        If Not notice.IsShowing Then Throw New InvalidOperationException("The recovery progress notice could not be displayed; saving has been deferred.")
                        SystemMessageManager.Publish(model.ModelID, "Saving recovery copy: " & IO.Path.GetFileName(model.FileName) & ". The original plan will not be changed.", SystemMessageSeverity.Information, "Recovery backup", model.FileName)
                        Try
                            Dim saved = RecoveryBackupStore.Write(model)
                            state.Revision = model.CalculationRevision
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
                    Busy = False
                    Trace.WriteLine("[Recovery Save Benchmark] model=" & model.ModelID.ToString() & ", total=" & timer.ElapsedMilliseconds.ToString() & " ms")
                End Try
                Exit For 'Never serialise several large plans in one timer tick.
            Next
        End Sub
    End Class
End Namespace
