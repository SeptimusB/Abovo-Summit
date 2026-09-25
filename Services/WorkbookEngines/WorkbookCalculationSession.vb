Option Strict On

Imports System.Diagnostics
Imports System.IO
Imports System.Security.Cryptography
Imports System.Threading
Imports System.Threading.Tasks

Namespace Abovo.WorkbookEngines
    Partial Public NotInheritable Class WorkbookCalculationSession
        Private ReadOnly owner As New WorkbookEngineStaHost()
        Private ReadOnly gate As New Object()
        Private backend As IWorkbookCalculationBackend
        Private sourceLease As FileStream
        Private closing As Boolean
        Private failed As Boolean
        Private currentRevision As Long
        Private ReadOnly timeoutMilliseconds As Integer
        Private cleanupTask As Task
        Private ReadOnly valueEditTrial As Boolean
        Private candidateSaveTrial As Boolean
        Private sourcePath As String
        Private sourceZone As String
        Private savePending As Boolean
        Private editPending As Boolean
        Public ReadOnly Property SessionId As Guid = Guid.NewGuid()
        Public ReadOnly Property SourceHash As String
        Public ReadOnly Property EngineName As String
        Public ReadOnly Property EngineVersion As String
        Public ReadOnly Property FallbackReason As String

        Private Sub New(timeoutMilliseconds As Integer, valueEditTrial As Boolean)
            Me.timeoutMilliseconds = timeoutMilliseconds
            Me.valueEditTrial = valueEditTrial
        End Sub

        Public ReadOnly Property Revision As Long
            Get
                SyncLock gate
                    Return currentRevision
                End SyncLock
            End Get
        End Property

        Public Shared Async Function OpenAsync(path As String, options As WorkbookEngineOptions,
                                               Optional cancellation As CancellationToken = Nothing,
                                               Optional factory As Func(Of WorkbookEnginePreference, IWorkbookCalculationBackend) = Nothing) As Task(Of WorkbookCalculationSession)
            If options Is Nothing Then Throw New ArgumentNullException(NameOf(options))
            If String.IsNullOrWhiteSpace(path) Then Throw New ArgumentException("A workbook path is required.", NameOf(path))
            Dim fullPath = IO.Path.GetFullPath(path)
            Dim extension = IO.Path.GetExtension(fullPath)
            If Not {".xlsb", ".xlsm", ".xlsx"}.Contains(extension, StringComparer.OrdinalIgnoreCase) Then
                Throw New ArgumentException("The calculation adapter supports XLSB, XLSM and XLSX workbooks only.", NameOf(path))
            End If
            Dim session As New WorkbookCalculationSession(options.OperationTimeoutMilliseconds, options.EnableValueEditTrial)
            session.candidateSaveTrial = options.EnableCandidateSaveTrial
            session.sourcePath = fullPath
            Dim openError As Exception = Nothing
            Try
                Dim opening = session.owner.InvokeAsync(Of Boolean)(Function()
                    cancellation.ThrowIfCancellationRequested()
                    ' Retain a read-only lease: the baseline cannot change under
                    ' an open native model. Stage one never saves this document.
                    session.sourceLease = New FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read)
                    If session.candidateSaveTrial Then session.sourceZone = ExcelCalculationBackend.ReadInternetZone(fullPath)
                    Using hash = SHA256.Create()
                        session._SourceHash = BitConverter.ToString(hash.ComputeHash(session.sourceLease)).Replace("-", "")
                        session.sourceLease.Position = 0
                    End Using
                    Dim createBackend As Func(Of WorkbookEnginePreference, IWorkbookCalculationBackend) = If(factory, AddressOf CreateNativeBackend)
                    If options.Preference <> WorkbookEnginePreference.DevExpressOnly Then
                        Try
                            session.backend = createBackend(WorkbookEnginePreference.ExcelRequired)
                            session.backend.OpenReadOnly(fullPath, options)
                        Catch ex As Exception When Not TypeOf ex Is OperationCanceledException AndAlso
                                                       options.Preference = WorkbookEnginePreference.Automatic
                            ' Fallback is allowed only before handing ownership
                            ' to the caller, never after edits or published reads.
                            SyncLock session.gate
                                session.RequireAvailable()
                            End SyncLock
                            If session.backend IsNot Nothing Then session.backend.Dispose()
                            session.backend = Nothing
                            session._FallbackReason = ex.GetType().Name & ": " & ex.Message
                        End Try
                    End If
                    cancellation.ThrowIfCancellationRequested()
                    SyncLock session.gate
                        session.RequireAvailable()
                    End SyncLock
                    If session.backend Is Nothing Then
                        session.backend = createBackend(WorkbookEnginePreference.DevExpressOnly)
                        session.backend.OpenReadOnly(fullPath, options)
                    End If
                    session._EngineName = session.backend.Name
                    session._EngineVersion = session.backend.Version
                    Return True
                End Function, cancellation)
                Await session.AwaitOperation(opening, "opening").ConfigureAwait(False)
                Return session
            Catch ex As Exception
                openError = ex
            End Try
            If TypeOf openError Is TimeoutException Then
                ' AwaitOperation has quarantined the owner and queued cleanup.
                ' Do not wait through a second timeout before informing the caller.
                Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(openError).Throw()
            End If
            ' VB does not permit Await in Catch/Finally.
            Try
                Await session.CloseAsync().ConfigureAwait(False)
            Catch cleanupError As Exception
                Throw New AggregateException("Workbook opening and cleanup failed.", openError, cleanupError)
            End Try
            Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(openError).Throw()
            Return Nothing
        End Function

        Private Shared Function CreateNativeBackend(preference As WorkbookEnginePreference) As IWorkbookCalculationBackend
            If preference = WorkbookEnginePreference.ExcelRequired Then Return New ExcelCalculationBackend()
            Return New DevExpressCalculationBackend()
        End Function

        Public Function InvalidateResults() As Long
            SyncLock gate
                RequireAvailable()
                If editPending OrElse savePending Then Throw New InvalidOperationException("A workbook edit or candidate export is still in progress.")
                currentRevision += 1
                Return currentRevision
            End SyncLock
        End Function

        Public Function IsCurrent(result As WorkbookCalculationResult) As Boolean
            SyncLock gate
                Return Not closing AndAlso Not failed AndAlso result IsNot Nothing AndAlso
                       result.SessionId = SessionId AndAlso result.Revision = currentRevision AndAlso
                       result.SourceHash = SourceHash
            End SyncLock
        End Function

        Public Async Function CalculateAndReadAsync(expectedRevision As Long, kind As WorkbookCalculationKind,
                                                    areas As IEnumerable(Of WorkbookReadArea),
                                                    Optional cancellation As CancellationToken = Nothing) As Task(Of WorkbookCalculationResult)
            If Not [Enum].IsDefined(GetType(WorkbookCalculationKind), kind) Then Throw New ArgumentOutOfRangeException(NameOf(kind))
            If areas Is Nothing Then Throw New ArgumentNullException(NameOf(areas))
            Dim requests As New List(Of WorkbookReadArea)()
            Dim cells As Long
            For Each area In areas
                If area Is Nothing Then Throw New ArgumentException("Null read area.", NameOf(areas))
                cells += CLng(area.Rows) * area.Columns
                If requests.Count = 128 OrElse cells > 500000 Then Throw New ArgumentException("Read batch exceeds transport limits.", NameOf(areas))
                requests.Add(area)
            Next
            If requests.Count = 0 Then Throw New ArgumentException("At least one read area is required.", NameOf(areas))
            SyncLock gate
                RequireRevision(expectedRevision)
            End SyncLock
            Dim calculation = owner.InvokeAsync(Function()
                SyncLock gate
                    RequireRevision(expectedRevision)
                End SyncLock
                Dim blocks As New List(Of WorkbookValueBlock)()
                Dim timer = Stopwatch.StartNew()
                Dim calculationMs As Long
                Try
                    backend.Calculate(kind)
                    SyncLock gate
                        RequireRevision(expectedRevision)
                    End SyncLock
                    calculationMs = timer.ElapsedMilliseconds
                    For Each area In requests
                        cancellation.ThrowIfCancellationRequested()
                        Dim block = backend.Read(area)
                        If block Is Nothing OrElse Not SameArea(block.Area, area) Then Throw New InvalidDataException("The calculation engine returned a different worksheet area from the requested area.")
                        blocks.Add(block)
                    Next
                Catch ex As OperationCanceledException
                    Throw
                Catch ex As StaleResultException
                    ' A newer edit is not a native engine failure.
                    Throw
                Catch
                    SyncLock gate
                        failed = True
                    End SyncLock
                    Throw
                End Try
                Dim result As New WorkbookCalculationResult(SessionId, expectedRevision, SourceHash, EngineName,
                                                            calculationMs, timer.ElapsedMilliseconds - calculationMs, blocks)
                SyncLock gate
                    RequireRevision(expectedRevision)
                End SyncLock
                Return result
            End Function, cancellation)
            Return Await AwaitOperation(calculation, "calculation").ConfigureAwait(False)
        End Function

        Private Shared Function SameArea(actual As WorkbookReadArea, requested As WorkbookReadArea) As Boolean
            Return String.Equals(actual.Worksheet, requested.Worksheet, StringComparison.OrdinalIgnoreCase) AndAlso
                   actual.Row = requested.Row AndAlso actual.Column = requested.Column AndAlso
                   actual.Rows = requested.Rows AndAlso actual.Columns = requested.Columns
        End Function

        Private Sub RequireAvailable()
            If closing Then Throw New ObjectDisposedException(NameOf(WorkbookCalculationSession))
            If failed Then Throw New InvalidOperationException("The calculation session failed. Reopen from a verified checkpoint; automatic mid-session fallback is not allowed.")
        End Sub

        Private Sub RequireRevision(expected As Long)
            RequireAvailable()
            If expected <> currentRevision Then Throw New StaleResultException()
        End Sub

        Private NotInheritable Class StaleResultException
            Inherits InvalidOperationException
            Friend Sub New()
                MyBase.New("The workbook changed while these results were being prepared. Request current results again.")
            End Sub
        End Class

        Private Async Function AwaitOperation(Of T)(operation As Task(Of T), operationName As String) As Task(Of T)
            Using deadline As New CancellationTokenSource()
                Dim elapsed = Task.Delay(timeoutMilliseconds, deadline.Token)
                If Await Task.WhenAny(operation, elapsed).ConfigureAwait(False) IsNot operation Then
                    SyncLock gate
                        failed = True
                    End SyncLock
                    ObserveFault(operation)
                    Dim pendingCleanup = StartCleanup()
                    Throw New TimeoutException("Workbook " & operationName & " exceeded its time limit. No result was accepted. The session is closed to further use; native cleanup is queued, without interrupting or switching engines.")
                End If
                deadline.Cancel()
                Return Await operation.ConfigureAwait(False)
            End Using
        End Function

        Private Shared Sub ObserveFault(task As Task)
            task.ContinueWith(Sub(completed)
                                  Dim observed = completed.Exception
                              End Sub, CancellationToken.None, TaskContinuationOptions.OnlyOnFaulted, TaskScheduler.Default)
        End Sub

        Public ReadOnly Property NativeCleanupCompletion As Task
            Get
                SyncLock gate
                    Return cleanupTask
                End SyncLock
            End Get
        End Property

        Private Function StartCleanup() As Task
            SyncLock gate
                closing = True
                If cleanupTask Is Nothing Then
                    cleanupTask = owner.StopAsync(Sub()
                        Try
                            If backend IsNot Nothing Then backend.Dispose()
                        Finally
                            backend = Nothing
                            If sourceLease IsNot Nothing Then sourceLease.Dispose()
                            sourceLease = Nothing
                        End Try
                    End Sub)
                    ObserveFault(cleanupTask)
                End If
                Return cleanupTask
            End SyncLock
        End Function

        Public Async Function CloseAsync() As Task
            Dim cleanup = StartCleanup()
            Using deadline As New CancellationTokenSource()
                If Await Task.WhenAny(cleanup, Task.Delay(timeoutMilliseconds, deadline.Token)).ConfigureAwait(False) IsNot cleanup Then
                    Throw New TimeoutException("Native workbook cleanup is still waiting. No further operations or engine fallback are permitted for this session.")
                End If
                deadline.Cancel()
                Await cleanup.ConfigureAwait(False)
            End Using
        End Function
    End Class
End Namespace
