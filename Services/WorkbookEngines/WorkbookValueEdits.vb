Option Strict On

Imports System.Threading
Imports System.Threading.Tasks

Namespace Abovo.WorkbookEngines
    ' Matches the existing two DataManager permission paths. XML/row/editor
    ' restrictions still belong to ModelChangeManager and its UI caller.
    Public Enum WorkbookValuePermission
        UnlockedCell = 0
        SolidFillRule = 1
    End Enum

    Public NotInheritable Class WorkbookCellState
        Public ReadOnly Property Value As Object
        Public ReadOnly Property Formula As String
        Public ReadOnly Property NumberFormat As String
        Public ReadOnly Property Locked As Boolean
        Public ReadOnly Property SolidFill As Boolean
        Public ReadOnly Property ArrayMember As Boolean
        Public ReadOnly Property Merged As Boolean
        Public ReadOnly Property WorksheetProtected As Boolean

        Public Sub New(value As Object, formula As String, numberFormat As String, locked As Boolean,
                       solidFill As Boolean, arrayMember As Boolean, merged As Boolean, worksheetProtected As Boolean)
            If TypeOf value Is WorkbookCellError Then
                ' Errors are readable but not a supported input/rollback type.
            Else
                ValidateValue(value)
            End If
            Me.Value = value : Me.Formula = If(formula, "") : Me.NumberFormat = If(numberFormat, "")
            Me.Locked = locked : Me.SolidFill = solidFill : Me.ArrayMember = arrayMember
            Me.Merged = merged : Me.WorksheetProtected = worksheetProtected
        End Sub

        Public Function AllowsValueEdit(permission As WorkbookValuePermission) As Boolean
            If Not [Enum].IsDefined(GetType(WorkbookValuePermission), permission) Then Throw New ArgumentOutOfRangeException(NameOf(permission))
            ' Conservative trial scope: never split arrays/merges, overwrite a
            ' formula, or bypass actual worksheet protection. No unprotect call.
            If Formula.Length <> 0 OrElse ArrayMember OrElse Merged OrElse TypeOf Value Is WorkbookCellError OrElse
               (WorksheetProtected AndAlso Locked) Then Return False
            Return If(permission = WorkbookValuePermission.SolidFillRule, SolidFill, Not Locked)
        End Function

        Friend Shared Sub ValidateValue(value As Object)
            If value Is Nothing OrElse TypeOf value Is Boolean Then Return
            If TypeOf value Is String AndAlso DirectCast(value, String).Length <= 32767 Then Return
            If TypeOf value Is Double AndAlso Not Double.IsNaN(CDbl(value)) AndAlso Not Double.IsInfinity(CDbl(value)) Then Return
            Throw New ArgumentException("Input must be a finite Double, Boolean, literal String (up to 32767 characters), or Nothing. Dates must already be converted by the model's date policy.")
        End Sub

        Friend Shared Function SameValue(left As Object, right As Object) As Boolean
            If TypeOf left Is WorkbookCellError AndAlso TypeOf right Is WorkbookCellError Then
                Return DirectCast(left, WorkbookCellError).Text = DirectCast(right, WorkbookCellError).Text
            End If
            Return Object.Equals(left, right)
        End Function

        Friend Function Matches(other As WorkbookCellState) As Boolean
            Return other IsNot Nothing AndAlso SameValue(Value, other.Value) AndAlso SameDefinition(other)
        End Function

        Friend Function SameDefinition(other As WorkbookCellState) As Boolean
            Return other IsNot Nothing AndAlso Formula = other.Formula AndAlso NumberFormat = other.NumberFormat AndAlso
                   Locked = other.Locked AndAlso SolidFill = other.SolidFill AndAlso ArrayMember = other.ArrayMember AndAlso
                   Merged = other.Merged AndAlso WorksheetProtected = other.WorksheetProtected
        End Function
    End Class

    Public NotInheritable Class WorkbookCellSnapshot
        Public ReadOnly Property SessionId As Guid
        Public ReadOnly Property Revision As Long
        Public ReadOnly Property Area As WorkbookReadArea
        Public ReadOnly Property State As WorkbookCellState
        Friend Sub New(sessionId As Guid, revision As Long, area As WorkbookReadArea, state As WorkbookCellState)
            If state Is Nothing Then Throw New ArgumentNullException(NameOf(state))
            Me.SessionId = sessionId : Me.Revision = revision : Me.Area = area : Me.State = state
        End Sub
    End Class

    Public NotInheritable Class WorkbookValueEditReceipt
        Public ReadOnly Property Before As WorkbookCellSnapshot
        Public ReadOnly Property After As WorkbookCellSnapshot
        Public ReadOnly Property Changed As Boolean
        Public ReadOnly Property Results As WorkbookCalculationResult
        Friend Sub New(before As WorkbookCellSnapshot, after As WorkbookCellSnapshot, changed As Boolean, results As WorkbookCalculationResult)
            Me.Before = before : Me.After = after : Me.Changed = changed : Me.Results = results
        End Sub
    End Class

    ' Optional extension: all calls remain on the session's STA. The production
    ' change manager must own journal/dirty state when this trial is integrated.
    Public Interface IWorkbookValueEditBackend
        Inherits IWorkbookCalculationBackend
        Function ReadCell(area As WorkbookReadArea) As WorkbookCellState
        Sub WriteValue(area As WorkbookReadArea, value As Object)
    End Interface

    Partial Public NotInheritable Class WorkbookCalculationSession
        Public Async Function CaptureCellAsync(expectedRevision As Long, area As WorkbookReadArea,
                                                Optional cancellation As CancellationToken = Nothing) As Task(Of WorkbookCellSnapshot)
            RequireSingleCell(area)
            SyncLock gate
                RequireRevision(expectedRevision)
            End SyncLock
            Dim reading = owner.InvokeAsync(Function()
                SyncLock gate
                    RequireRevision(expectedRevision)
                End SyncLock
                Dim editor = ValueEditor()
                Dim state As WorkbookCellState
                Try
                    state = editor.ReadCell(area)
                Catch
                    SyncLock gate
                        failed = True
                    End SyncLock
                    Throw
                End Try
                SyncLock gate
                    RequireRevision(expectedRevision)
                End SyncLock
                Return New WorkbookCellSnapshot(SessionId, expectedRevision, area, state)
            End Function, cancellation)
            Return Await AwaitOperation(reading, "cell capture").ConfigureAwait(False)
        End Function

        Private Shared Sub RequireSingleCell(area As WorkbookReadArea)
            If area Is Nothing OrElse area.Rows <> 1 OrElse area.Columns <> 1 Then Throw New ArgumentException("A single worksheet cell is required.", NameOf(area))
        End Sub

        Private Function ValueEditor() As IWorkbookValueEditBackend
            Dim editor = TryCast(backend, IWorkbookValueEditBackend)
            If editor Is Nothing Then Throw New NotSupportedException("This engine does not implement the isolated value-edit boundary.")
            Return editor
        End Function

        Public Async Function ApplyValueAsync(expected As WorkbookCellSnapshot, value As Object,
                                              permission As WorkbookValuePermission,
                                              Optional readAreas As IEnumerable(Of WorkbookReadArea) = Nothing,
                                              Optional cancellation As CancellationToken = Nothing) As Task(Of WorkbookValueEditReceipt)
            If Not valueEditTrial Then Throw New InvalidOperationException("Value-edit trials were not enabled for this session.")
            If expected Is Nothing Then Throw New ArgumentNullException(NameOf(expected))
            If expected.SessionId <> SessionId Then Throw New InvalidOperationException("The input snapshot belongs to another workbook session.")
            WorkbookCellState.ValidateValue(value)
            ' Clearing a text editor means an empty input, not a formula that
            ' returns an empty string. Both native engines store this as blank.
            If TypeOf value Is String AndAlso CStr(value).Length = 0 Then value = Nothing
            If Not [Enum].IsDefined(GetType(WorkbookValuePermission), permission) Then Throw New ArgumentOutOfRangeException(NameOf(permission))
            Dim areas As New List(Of WorkbookReadArea)()
            Dim count As Long
            If readAreas IsNot Nothing Then
                For Each area In readAreas
                    If area Is Nothing Then Throw New ArgumentException("Null read area.", NameOf(readAreas))
                    count += CLng(area.Rows) * area.Columns
                    If areas.Count = 128 OrElse count > 500000 Then Throw New ArgumentException("Read batch exceeds transport limits.", NameOf(readAreas))
                    areas.Add(area)
                Next
            End If
            cancellation.ThrowIfCancellationRequested()
            SyncLock gate
                RequireRevision(expected.Revision)
                If editPending OrElse savePending Then Throw New InvalidOperationException("A workbook edit or candidate export is still in progress.")
                editPending = True
            End SyncLock
            Try
                ' The action owns cancellation and compensation. The generic
                ' read-only dispatcher must not discard an already-committed
                ' edit receipt if cancellation arrives just after commit.
                Dim editing = owner.InvokeAsync(Function() ApplyValueOnOwner(expected, value, permission, areas, cancellation), CancellationToken.None)
                Return Await AwaitOperation(editing, "value edit").ConfigureAwait(False)
            Finally
                SyncLock gate
                    editPending = False
                End SyncLock
            End Try
        End Function

        Private Function ApplyValueOnOwner(expected As WorkbookCellSnapshot, value As Object, permission As WorkbookValuePermission,
                                           areas As List(Of WorkbookReadArea), cancellation As CancellationToken) As WorkbookValueEditReceipt
            cancellation.ThrowIfCancellationRequested()
            SyncLock gate
                RequireRevision(expected.Revision)
            End SyncLock
            Dim editor = ValueEditor()
            Dim before As WorkbookCellState
            Try
                before = editor.ReadCell(expected.Area)
            Catch
                SyncLock gate
                    failed = True
                End SyncLock
                Throw
            End Try
            If Not expected.State.Matches(before) Then
                SyncLock gate
                    RequireRevision(expected.Revision)
                    currentRevision += 1 ' A detected untracked change invalidates cached outputs too.
                End SyncLock
                Throw New InvalidOperationException("The cell or its permission/format state changed; refresh it before editing.")
            End If
            If Not before.AllowsValueEdit(permission) Then Throw New InvalidOperationException("The target is not editable under the selected workbook rule.")
            If WorkbookCellState.SameValue(before.Value, value) Then Return New WorkbookValueEditReceipt(expected, expected, False, Nothing)
            Dim revision As Long
            SyncLock gate
                RequireRevision(expected.Revision)
                currentRevision += 1 ' Invalidate old results before the first write.
                revision = currentRevision
            End SyncLock
            Try
                cancellation.ThrowIfCancellationRequested()
                editor.WriteValue(expected.Area, value)
                Dim written = editor.ReadCell(expected.Area)
                If Not WorkbookCellState.SameValue(written.Value, value) OrElse Not before.SameDefinition(written) Then
                    Throw New InvalidOperationException("The engine did not retain the exact typed input and cell definition.")
                End If
                cancellation.ThrowIfCancellationRequested()
                Dim timer = Diagnostics.Stopwatch.StartNew()
                backend.Calculate(WorkbookCalculationKind.Full)
                Dim calculationMs = timer.ElapsedMilliseconds
                cancellation.ThrowIfCancellationRequested()
                SyncLock gate
                    RequireRevision(revision)
                End SyncLock
                Dim after = editor.ReadCell(expected.Area)
                If Not WorkbookCellState.SameValue(after.Value, value) OrElse Not before.SameDefinition(after) Then
                    Throw New InvalidOperationException("Calculation changed the input or its cell definition unexpectedly.")
                End If
                Dim blocks As New List(Of WorkbookValueBlock)()
                For Each area In areas
                    cancellation.ThrowIfCancellationRequested()
                    Dim block = backend.Read(area)
                    If block Is Nothing OrElse Not SameArea(block.Area, area) Then Throw New IO.InvalidDataException("Edit read-back returned a different worksheet rectangle.")
                    blocks.Add(block)
                Next
                cancellation.ThrowIfCancellationRequested()
                SyncLock gate
                    RequireRevision(revision)
                End SyncLock
                Dim results = If(blocks.Count = 0, Nothing, New WorkbookCalculationResult(SessionId, revision, SourceHash, EngineName,
                    calculationMs, timer.ElapsedMilliseconds - calculationMs, blocks))
                Return New WorkbookValueEditReceipt(expected, New WorkbookCellSnapshot(SessionId, revision, expected.Area, after), True, results)
            Catch editError As Exception
                Try
                    If Not before.Matches(editor.ReadCell(expected.Area)) Then editor.WriteValue(expected.Area, before.Value)
                    backend.Calculate(WorkbookCalculationKind.Full)
                    If Not before.Matches(editor.ReadCell(expected.Area)) Then Throw New InvalidOperationException("The original cell state was not restored.")
                Catch rollbackError As Exception
                    SyncLock gate
                        failed = True
                    End SyncLock
                    Throw New AggregateException("The value edit and its restoration failed. The session cannot be reused or saved.", editError, rollbackError)
                End Try
                ' Failed/cancelled edits produce no receipt/history. The bumped
                ' revision remains so callers must refresh even after restoration.
                Throw
            End Try
        End Function
    End Class
End Namespace
