Option Strict On

Imports System.Diagnostics
Imports System.IO
Imports System.Security.Cryptography
Imports System.Threading
Imports System.Threading.Tasks

Namespace Abovo.WorkbookEngines
    Public NotInheritable Class WorkbookCalculationSession
        Private ReadOnly owner As New WorkbookEngineStaHost()
        Private ReadOnly gate As New Object()
        Private backend As IWorkbookCalculationBackend
        Private sourceLease As FileStream
        Private closing As Boolean
        Private failed As Boolean
        Private currentRevision As Long
        Public ReadOnly Property SessionId As Guid = Guid.NewGuid()
        Public ReadOnly Property SourceHash As String
        Public ReadOnly Property EngineName As String
        Public ReadOnly Property EngineVersion As String
        Public ReadOnly Property FallbackReason As String

        Private Sub New()
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
            Dim session As New WorkbookCalculationSession()
            Dim openError As Exception = Nothing
            Try
                Await session.owner.InvokeAsync(Of Boolean)(Function()
                    cancellation.ThrowIfCancellationRequested()
                    ' Retain a read-only lease: the baseline cannot change under
                    ' an open native model. Stage one never saves this document.
                    session.sourceLease = New FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read)
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
                            If session.backend IsNot Nothing Then session.backend.Dispose()
                            session.backend = Nothing
                            session._FallbackReason = ex.GetType().Name & ": " & ex.Message
                        End Try
                    End If
                    cancellation.ThrowIfCancellationRequested()
                    If session.backend Is Nothing Then
                        session.backend = createBackend(WorkbookEnginePreference.DevExpressOnly)
                        session.backend.OpenReadOnly(fullPath, options)
                    End If
                    session._EngineName = session.backend.Name
                    session._EngineVersion = session.backend.Version
                    Return True
                End Function, cancellation).ConfigureAwait(False)
                Return session
            Catch ex As Exception
                openError = ex
            End Try
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
            Return Await owner.InvokeAsync(Function()
                SyncLock gate
                    RequireRevision(expectedRevision)
                End SyncLock
                Dim blocks As New List(Of WorkbookValueBlock)()
                Dim timer = Stopwatch.StartNew()
                Dim calculationMs As Long
                Try
                    backend.Calculate(kind)
                    calculationMs = timer.ElapsedMilliseconds
                    For Each area In requests
                        cancellation.ThrowIfCancellationRequested()
                        Dim block = backend.Read(area)
                        If block Is Nothing OrElse Not SameArea(block.Area, area) Then Throw New InvalidDataException("The calculation engine returned a different worksheet area from the requested area.")
                        blocks.Add(block)
                    Next
                Catch ex As OperationCanceledException
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
            End Function, cancellation).ConfigureAwait(False)
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
            If expected <> currentRevision Then Throw New InvalidOperationException("The workbook changed while these results were being prepared. Request current results again.")
        End Sub

        Public Function CloseAsync() As Task
            SyncLock gate
                closing = True
            End SyncLock
            Return owner.StopAsync(Sub()
                Try
                    If backend IsNot Nothing Then backend.Dispose()
                Finally
                    backend = Nothing
                    If sourceLease IsNot Nothing Then sourceLease.Dispose()
                    sourceLease = Nothing
                End Try
            End Sub)
        End Function
    End Class
End Namespace
