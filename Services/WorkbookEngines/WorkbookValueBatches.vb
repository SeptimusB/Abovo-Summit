Option Strict On

Imports System.Collections.ObjectModel
Imports System.Threading
Imports System.Threading.Tasks

Namespace Abovo.WorkbookEngines
    Public NotInheritable Class WorkbookValueChange
        Public ReadOnly Property Expected As WorkbookCellSnapshot
        Public ReadOnly Property Value As Object
        Public ReadOnly Property Permission As WorkbookValuePermission
        Public ReadOnly Property CalculateBefore As Boolean
        Public ReadOnly Property SkipUnavailable As Boolean
        Public Sub New(expected As WorkbookCellSnapshot, value As Object, permission As WorkbookValuePermission,
                       Optional calculateBefore As Boolean = False, Optional skipUnavailable As Boolean = False)
            If expected Is Nothing Then Throw New ArgumentNullException(NameOf(expected))
            WorkbookCellState.ValidateValue(value)
            If Not [Enum].IsDefined(GetType(WorkbookValuePermission), permission) Then Throw New ArgumentOutOfRangeException(NameOf(permission))
            Me.Expected = expected
            Me.Value = If(TypeOf value Is String AndAlso CStr(value).Length = 0, Nothing, value)
            Me.Permission = permission : Me.CalculateBefore = calculateBefore : Me.SkipUnavailable = skipUnavailable
        End Sub
    End Class

    Public NotInheritable Class WorkbookValueBatchReceipt
        Public ReadOnly Property SessionId As Guid
        Public ReadOnly Property Revision As Long
        Public ReadOnly Property Changes As ReadOnlyCollection(Of WorkbookValueEditReceipt)
        Public ReadOnly Property Skipped As ReadOnlyCollection(Of WorkbookReadArea)
        Public ReadOnly Property Results As WorkbookCalculationResult
        Public ReadOnly Property Changed As Boolean
            Get
                Return Changes.Any(Function(item) item.Changed)
            End Get
        End Property
        Friend Sub New(session As Guid, revision As Long, changes As IList(Of WorkbookValueEditReceipt),
                       skipped As IList(Of WorkbookReadArea), results As WorkbookCalculationResult)
            SessionId = session : Me.Revision = revision
            Me.Changes = New List(Of WorkbookValueEditReceipt)(changes).AsReadOnly()
            Me.Skipped = New List(Of WorkbookReadArea)(skipped).AsReadOnly() : Me.Results = results
        End Sub
    End Class

    Partial Public NotInheritable Class WorkbookCalculationSession
        Friend Sub RequireEditingSource(path As String)
            SyncLock gate
                RequireAvailable()
                If Not valueEditTrial OrElse Not String.Equals(IO.Path.GetFullPath(path), sourcePath, StringComparison.OrdinalIgnoreCase) Then Throw New InvalidOperationException("The engine edit owner belongs to a different source or is read-only.")
            End SyncLock
        End Sub
        ' One owner invocation and revision for the complete value command.
        ' Formulas, arrays, structure and protection need separately qualified commands.
        Public Async Function CaptureCellsAsync(expectedRevision As Long, areas As IEnumerable(Of WorkbookReadArea),
                                               Optional cancellation As CancellationToken = Nothing) As Task(Of ReadOnlyCollection(Of WorkbookCellSnapshot))
            Dim requests = ValueBatchAreas(areas)
            SyncLock gate
                RequireRevision(expectedRevision)
            End SyncLock
            Dim reading = owner.InvokeAsync(Function()
                SyncLock gate
                    RequireRevision(expectedRevision)
                End SyncLock
                Dim snapshots As New List(Of WorkbookCellSnapshot)()
                Try
                    Dim editor = ValueEditor()
                    For Each area In requests
                        cancellation.ThrowIfCancellationRequested()
                        snapshots.Add(New WorkbookCellSnapshot(SessionId, expectedRevision, area, editor.ReadCell(area)))
                    Next
                Catch ex As OperationCanceledException
                    Throw
                Catch
                    SyncLock gate
                        failed = True
                    End SyncLock
                    Throw
                End Try
                SyncLock gate
                    RequireRevision(expectedRevision)
                End SyncLock
                Return snapshots.AsReadOnly()
            End Function, cancellation)
            Return Await AwaitOperation(reading, "input batch capture").ConfigureAwait(False)
        End Function

        Private Shared Function ValueBatchAreas(areas As IEnumerable(Of WorkbookReadArea)) As List(Of WorkbookReadArea)
            If areas Is Nothing Then Throw New ArgumentNullException(NameOf(areas))
            Dim requests As New List(Of WorkbookReadArea)()
            Dim keys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each area In areas
                RequireSingleCell(area)
                If requests.Count = 10000 Then Throw New ArgumentException("Value batches are limited to 10,000 distinct cells.")
                If Not keys.Add(area.Worksheet & "!" & area.Address) Then Throw New ArgumentException("Duplicate input cell in batch.")
                requests.Add(area)
            Next
            If requests.Count = 0 Then Throw New ArgumentException("At least one input is required.")
            Return requests
        End Function

        Public Async Function ApplyValuesAsync(changes As IEnumerable(Of WorkbookValueChange),
                                               Optional readAreas As IEnumerable(Of WorkbookReadArea) = Nothing,
                                               Optional cancellation As CancellationToken = Nothing,
                                               Optional includePresentation As Boolean = False) As Task(Of WorkbookValueBatchReceipt)
            If Not valueEditTrial Then Throw New InvalidOperationException("Value-edit trials were not enabled for this session.")
            If changes Is Nothing Then Throw New ArgumentNullException(NameOf(changes))
            Dim inputs As New List(Of WorkbookValueChange)()
            For Each item In changes
                If item Is Nothing OrElse inputs.Count = 10000 Then Throw New ArgumentException("Invalid or oversized value batch.")
                If item.Expected.SessionId <> SessionId Then Throw New InvalidOperationException("The input belongs to another session.")
                inputs.Add(item)
            Next
            ValueBatchAreas(inputs.Select(Function(item) item.Expected.Area))
            Dim revision = inputs(0).Expected.Revision
            If inputs.Any(Function(item) item.Expected.Revision <> revision) Then Throw New InvalidOperationException("All input snapshots must share one revision.")
            Dim reads As New List(Of WorkbookReadArea)()
            Dim cells As Long
            If readAreas IsNot Nothing Then
                For Each area In readAreas
                    If area Is Nothing Then Throw New ArgumentException("Null read area.")
                    cells += CLng(area.Rows) * area.Columns
                    If reads.Count = 128 OrElse cells > 500000 Then Throw New ArgumentException("Read batch exceeds transport limits.")
                    reads.Add(area)
                Next
            End If
            If includePresentation AndAlso cells > 10000 Then Throw New ArgumentException("A styled read is limited to 10,000 visible cells.")
            cancellation.ThrowIfCancellationRequested()
            SyncLock gate
                RequireRevision(revision)
                If editPending OrElse savePending Then Throw New InvalidOperationException("Another edit or candidate export is still in progress.")
                editPending = True
            End SyncLock
            Try
                Dim editing = owner.InvokeAsync(Function() ApplyValuesOnOwner(inputs, reads, cancellation, includePresentation), CancellationToken.None)
                Return Await AwaitOperation(editing, "value batch").ConfigureAwait(False)
            Finally
                SyncLock gate
                    editPending = False
                End SyncLock
            End Try
        End Function

        Private Function ApplyValuesOnOwner(inputs As List(Of WorkbookValueChange), reads As List(Of WorkbookReadArea),
                                            cancellation As CancellationToken, includePresentation As Boolean) As WorkbookValueBatchReceipt
            cancellation.ThrowIfCancellationRequested()
            Dim revision = inputs(0).Expected.Revision
            SyncLock gate
                RequireRevision(revision)
            End SyncLock
            Dim editor = ValueEditor()
            ' Check all original states before the first mutation.
            For Each item In inputs
                cancellation.ThrowIfCancellationRequested()
                Dim actual As WorkbookCellState
                Try
                    actual = editor.ReadCell(item.Expected.Area)
                Catch
                    SyncLock gate
                        failed = True
                    End SyncLock
                    Throw
                End Try
                If Not item.Expected.State.Matches(actual) Then
                    SyncLock gate
                        RequireRevision(revision)
                        currentRevision += 1
                    End SyncLock
                    Throw New InvalidOperationException("An input or its definition changed; refresh before applying this command.")
                End If
            Next
            Dim touched As New List(Of WorkbookValueChange)()
            Dim skipped As New List(Of WorkbookReadArea)()
            Dim changed As Boolean = False
            Dim timer = Diagnostics.Stopwatch.StartNew()
            Dim calculationMs As Long
            Dim nativeOperationActive As Boolean
            Try
                For Each item In inputs
                    cancellation.ThrowIfCancellationRequested()
                    If item.CalculateBefore AndAlso changed Then
                        Dim started = timer.ElapsedMilliseconds
                        nativeOperationActive = True
                        CalculateNative(WorkbookCalculationKind.Full)
                        nativeOperationActive = False
                        calculationMs += timer.ElapsedMilliseconds - started
                        cancellation.ThrowIfCancellationRequested()
                    End If
                    nativeOperationActive = True
                    Dim current = editor.ReadCell(item.Expected.Area)
                    nativeOperationActive = False
                    If Not WorkbookCellState.SameValue(item.Expected.State.Value, current.Value) OrElse Not item.Expected.State.SameInputDefinition(current) Then Throw New InvalidOperationException("An earlier input unexpectedly changed another target's value or definition.")
                    If Not current.AllowsValueEdit(item.Permission) Then
                        If Not item.SkipUnavailable Then Throw New InvalidOperationException("The target is not editable under the selected workbook rule.")
                        skipped.Add(item.Expected.Area)
                        Continue For
                    End If
                    If WorkbookCellState.SameValue(current.Value, item.Value) Then Continue For
                    If Not changed Then
                        SyncLock gate
                            RequireRevision(revision)
                            currentRevision += 1 : revision = currentRevision
                        End SyncLock
                        changed = True
                    End If
                    ' A throwing native write can still have changed its target.
                    touched.Add(item)
                    editor.WriteValue(item.Expected.Area, item.Value)
                    Dim written = editor.ReadCell(item.Expected.Area)
                    If Not WorkbookCellState.SameValue(written.Value, item.Value) OrElse Not current.SameInputDefinition(written) Then Throw New InvalidOperationException("The engine did not retain the exact input and definition.")
                Next
                If changed OrElse reads.Count > 0 Then
                    Dim started = timer.ElapsedMilliseconds
                    nativeOperationActive = True
                    CalculateNative(WorkbookCalculationKind.Full)
                    nativeOperationActive = False
                    calculationMs += timer.ElapsedMilliseconds - started
                End If
                cancellation.ThrowIfCancellationRequested()
                Dim committed As New List(Of WorkbookValueEditReceipt)()
                For Each item In touched
                    Dim after = editor.ReadCell(item.Expected.Area)
                    If Not WorkbookCellState.SameValue(after.Value, item.Value) OrElse Not item.Expected.State.SameInputDefinition(after) Then Throw New InvalidOperationException("Calculation changed an input or its definition.")
                    committed.Add(New WorkbookValueEditReceipt(item.Expected, New WorkbookCellSnapshot(SessionId, revision, item.Expected.Area, after), True, Nothing))
                Next
                Dim blocks As New List(Of WorkbookValueBlock)()
                For Each area In reads
                    cancellation.ThrowIfCancellationRequested()
                    nativeOperationActive = True
                    Dim block = backend.Read(area)
                    If block Is Nothing OrElse Not SameArea(block.Area, area) Then Throw New IO.InvalidDataException("Batch read-back returned a different rectangle.")
                    nativeOperationActive = False
                    blocks.Add(block)
                Next
                cancellation.ThrowIfCancellationRequested()
                SyncLock gate
                    RequireRevision(revision)
                End SyncLock
                nativeOperationActive = True
                Dim display = If(includePresentation, ReadPresentationOnOwner(blocks), Nothing)
                nativeOperationActive = False
                cancellation.ThrowIfCancellationRequested()
                Dim result = If(blocks.Count = 0, Nothing, MakeResult(revision, calculationMs, timer.ElapsedMilliseconds - calculationMs, blocks, display))
                Return New WorkbookValueBatchReceipt(SessionId, revision, committed, skipped, result)
            Catch editError As Exception
                Dim failures As New List(Of Exception)()
                ' A read/calculation failure with no writes still invalidates
                ' previously displayed values. Do not advertise cached success.
                If Not changed AndAlso Not TypeOf editError Is OperationCanceledException Then
                    SyncLock gate
                        currentRevision += 1
                        If nativeOperationActive Then failed = True
                    End SyncLock
                End If
                For Each item In touched.AsEnumerable().Reverse()
                    Try
                        editor.WriteValue(item.Expected.Area, item.Expected.State.Value)
                    Catch ex As Exception
                        failures.Add(ex)
                    End Try
                Next
                If changed Then
                    Try
                        CalculateNative(WorkbookCalculationKind.Full)
                        For Each item In inputs
                            If Not item.Expected.State.Matches(editor.ReadCell(item.Expected.Area)) Then Throw New InvalidOperationException("Could not verify the complete original input state.")
                        Next
                    Catch ex As Exception
                        failures.Add(ex)
                    End Try
                End If
                If failures.Count > 0 Then
                    SyncLock gate
                        failed = True
                    End SyncLock
                    failures.Insert(0, editError)
                    Throw New AggregateException("The batch and its restoration failed. Reopen from a verified checkpoint.", failures)
                End If
                Throw
            End Try
        End Function
    End Class
End Namespace
