Option Strict On

Imports System.Collections.ObjectModel
Imports System.Diagnostics
Imports System.IO
Imports System.Security.Cryptography
Imports System.Threading
Imports System.Threading.Tasks

Namespace Abovo.WorkbookEngines
    ' A receipt for verified bytes, NOT a successful user Save. Publication,
    ' source conflict handling, history XML and dirty acknowledgement come later.
    Public NotInheritable Class WorkbookSaveCandidate
        Public ReadOnly Property SessionId As Guid
        Public ReadOnly Property Revision As Long
        Public ReadOnly Property SourceHash As String
        Public ReadOnly Property Path As String
        Public ReadOnly Property Hash As String
        Public ReadOnly Property ExportMilliseconds As Long
        Public ReadOnly Property VerificationMilliseconds As Long
        Friend Sub New(session As WorkbookCalculationSession, path As String, hash As String, exportMs As Long, verificationMs As Long)
            SessionId = session.SessionId : Revision = session.Revision : SourceHash = session.SourceHash
            Me.Path = path : Me.Hash = hash : ExportMilliseconds = exportMs : VerificationMilliseconds = verificationMs
        End Sub
    End Class

    Public NotInheritable Class WorkbookCandidateReadback
        Public ReadOnly Property Blocks As ReadOnlyCollection(Of WorkbookValueBlock)
        Public ReadOnly Property Cells As ReadOnlyCollection(Of WorkbookCellState)
        Public Sub New(blocks As IEnumerable(Of WorkbookValueBlock), cells As IEnumerable(Of WorkbookCellState))
            Me.Blocks = New List(Of WorkbookValueBlock)(blocks).AsReadOnly()
            Me.Cells = New List(Of WorkbookCellState)(cells).AsReadOnly()
        End Sub
    End Class

    Public Interface IWorkbookCandidateBackend
        Inherits IWorkbookValueEditBackend
        Sub ExportCopy(path As String)
        ' Reopen with macros/events/links/calculation disabled. Never repair a
        ' candidate on open: its persisted values, not a fresh calculation, count.
        Function ReadCopy(path As String, areas As IList(Of WorkbookReadArea), cells As IList(Of WorkbookReadArea)) As WorkbookCandidateReadback
    End Interface

    Partial Public NotInheritable Class WorkbookCalculationSession
        Public Function IsCurrentCandidate(candidate As WorkbookSaveCandidate) As Boolean
            SyncLock gate
                Return Not closing AndAlso Not failed AndAlso Not editPending AndAlso Not savePending AndAlso candidate IsNot Nothing AndAlso
                    candidate.SessionId = SessionId AndAlso candidate.Revision = currentRevision AndAlso candidate.SourceHash = SourceHash
            End SyncLock
        End Function

        Public Async Function CreateSaveCandidateAsync(expected As WorkbookCalculationResult, directory As String,
                                                       checkpoints As IEnumerable(Of WorkbookCellSnapshot),
                                                       Optional cancellation As CancellationToken = Nothing) As Task(Of WorkbookSaveCandidate)
            If Not candidateSaveTrial Then Throw New InvalidOperationException("Candidate save trials were not enabled for this session.")
            If expected Is Nothing OrElse Not IsCurrent(expected) Then Throw New InvalidOperationException("Current calculated results are required for a candidate export.")
            If String.IsNullOrWhiteSpace(directory) OrElse Not IO.Path.IsPathRooted(directory) OrElse Not IO.Directory.Exists(directory) Then
                Throw New ArgumentException("An existing absolute private trial directory is required.", NameOf(directory))
            End If
            Dim probes As New List(Of WorkbookCellSnapshot)()
            If checkpoints Is Nothing Then Throw New ArgumentNullException(NameOf(checkpoints))
            For Each cell In checkpoints
                If cell Is Nothing OrElse cell.SessionId <> SessionId OrElse cell.Revision <> expected.Revision Then Throw New ArgumentException("Checkpoint belongs to another session or revision.")
                If probes.Count = 64 Then Throw New ArgumentException("At most 64 cell-definition checkpoints are supported.")
                probes.Add(cell)
            Next
            If probes.Count = 0 Then Throw New ArgumentException("At least one cell-definition checkpoint is required.")
            cancellation.ThrowIfCancellationRequested()
            SyncLock gate
                RequireRevision(expected.Revision)
                If editPending OrElse savePending Then Throw New InvalidOperationException("A workbook edit or candidate export is still in progress.")
                savePending = True
            End SyncLock
            Try
                ' Owner handles cancellation so a completed receipt cannot be
                ' discarded by the read-only dispatcher's post-action check.
                Dim saving = owner.InvokeAsync(Function() CreateCandidateOnOwner(expected, directory, probes, cancellation), CancellationToken.None)
                Return Await AwaitOperation(saving, "candidate export").ConfigureAwait(False)
            Finally
                SyncLock gate
                    savePending = False
                End SyncLock
            End Try
        End Function

        Private Function CreateCandidateOnOwner(expected As WorkbookCalculationResult, directory As String,
                                                probes As List(Of WorkbookCellSnapshot), cancellation As CancellationToken) As WorkbookSaveCandidate
            cancellation.ThrowIfCancellationRequested()
            SyncLock gate
                RequireRevision(expected.Revision)
            End SyncLock
            Dim saver = TryCast(backend, IWorkbookCandidateBackend)
            If saver Is Nothing Then Throw New NotSupportedException("This engine has no candidate export adapter.")
            Dim areas = expected.Blocks.Select(Function(b) b.Area).ToList()
            Dim cellAreas = probes.Select(Function(p) p.Area).ToList()
            ' Validate before any file creation, including exact native state.
            VerifySessionState(expected, probes, saver)
            Dim package = WorkbookCandidatePackage.Capture(sourcePath)
            Dim folder = IO.Path.Combine(IO.Path.GetFullPath(directory), "candidate-" & Guid.NewGuid().ToString("N"))
            IO.Directory.CreateDirectory(folder)
            Dim path = IO.Path.Combine(folder, "~Summit-candidate" & IO.Path.GetExtension(sourcePath))
            Dim accepted As Boolean = False
            Try
                Dim clock = Stopwatch.StartNew()
                Try
                    saver.ExportCopy(path)
                Catch
                    ' A native save can fail after touching more workbook state
                    ' than our bounded checkpoints capture. Do not assume that
                    ' a readable input proves the whole owner remains usable.
                    SyncLock gate
                        failed = True
                    End SyncLock
                    Throw
                End Try
                Dim exportMs = clock.ElapsedMilliseconds
                cancellation.ThrowIfCancellationRequested()
                SyncLock gate
                    RequireRevision(expected.Revision)
                End SyncLock
                package.PrepareValueOnlyCandidate(path)
                ExcelCalculationBackend.CopyInternetZone(path, sourceZone)
                Using lease As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)
                    package.Verify(path)
                    Dim readback As WorkbookCandidateReadback
                    Try
                        readback = saver.ReadCopy(path, areas, cellAreas)
                    Catch
                        ' Includes candidate-close/COM cleanup failures. A spare
                        ' native workbook must not remain part of later calcs.
                        SyncLock gate
                            failed = True
                        End SyncLock
                        Throw
                    End Try
                    If readback Is Nothing OrElse readback.Blocks.Count <> expected.Blocks.Count OrElse readback.Cells.Count <> probes.Count Then Throw New InvalidDataException("Candidate read-back geometry differs from the request.")
                    For i = 0 To expected.Blocks.Count - 1
                        VerifyBlock(expected.Blocks(i), readback.Blocks(i))
                    Next
                    For i = 0 To probes.Count - 1
                        If Not probes(i).State.Matches(readback.Cells(i)) Then Throw New InvalidDataException("Candidate cell definition/value differs at " & probes(i).Area.Worksheet & "!" & probes(i).Area.Address)
                    Next
                    VerifySessionState(expected, probes, saver)
                    If ExcelCalculationBackend.ReadInternetZone(path) <> sourceZone OrElse ExcelCalculationBackend.ReadInternetZone(sourcePath) <> sourceZone Then Throw New InvalidDataException("Workbook security marking changed during candidate export.")
                    Dim hash = HashStream(lease)
                    cancellation.ThrowIfCancellationRequested()
                    SyncLock gate
                        RequireRevision(expected.Revision)
                        accepted = True
                        Return New WorkbookSaveCandidate(Me, path, hash, exportMs, clock.ElapsedMilliseconds - exportMs)
                    End SyncLock
                End Using
            Catch
                ' A failed export is not an edit. Retain a usable engine only
                ' if its input/definition checkpoints and outputs still agree.
                Try
                    VerifySessionState(expected, probes, saver)
                Catch
                    SyncLock gate
                        failed = True
                    End SyncLock
                End Try
                Throw
            Finally
                If Not accepted Then
                    ' Exact generated file only; never delete caller directories
                    ' recursively or touch the source workbook.
                    If File.Exists(path) Then File.Delete(path)
                    If IO.Directory.Exists(folder) AndAlso Not IO.Directory.EnumerateFileSystemEntries(folder).Any() Then IO.Directory.Delete(folder)
                End If
            End Try
        End Function

        Private Sub VerifySessionState(expected As WorkbookCalculationResult, probes As List(Of WorkbookCellSnapshot), saver As IWorkbookCandidateBackend)
            Try
                For Each probe In probes
                    If Not probe.State.Matches(saver.ReadCell(probe.Area)) Then Throw New InvalidDataException("The session changed outside its revisioned edit boundary.")
                Next
                For Each block In expected.Blocks
                    VerifyBlock(block, saver.Read(block.Area))
                Next
            Catch
                SyncLock gate
                    failed = True
                End SyncLock
                Throw
            End Try
        End Sub

        Private Shared Sub VerifyBlock(expected As WorkbookValueBlock, actual As WorkbookValueBlock)
            If actual Is Nothing OrElse Not SameArea(expected.Area, actual.Area) Then Throw New InvalidDataException("Candidate worksheet rectangle differs from the request.")
            For r = 0 To expected.Area.Rows - 1
                For c = 0 To expected.Area.Columns - 1
                    Dim left = expected.ValueAt(r, c), right = actual.ValueAt(r, c)
                    Dim same = WorkbookCellState.SameValue(left, right)
                    ' Native save precision can round the final binary ULP.
                    If TypeOf left Is Double AndAlso TypeOf right Is Double Then same = Math.Abs(CDbl(left) - CDbl(right)) <= Math.Max(0.000000001, Math.Abs(CDbl(left)) * 0.000000000001)
                    If Not same Then Throw New InvalidDataException("Candidate values differ at " & expected.Area.Worksheet & " row " & (expected.Area.Row + r + 1).ToString() & ", column " & (expected.Area.Column + c + 1).ToString())
                Next
            Next
        End Sub

        Public Async Function ValidateSaveCandidateAsync(candidate As WorkbookSaveCandidate, Optional cancellation As CancellationToken = Nothing) As Task
            If candidate Is Nothing OrElse Not IsCurrentCandidate(candidate) Then Throw New InvalidOperationException("Candidate belongs to an old or unavailable session revision.")
            Dim checking = owner.InvokeAsync(Function()
                cancellation.ThrowIfCancellationRequested()
                SyncLock gate
                    RequireRevision(candidate.Revision)
                End SyncLock
                Using stream As New FileStream(candidate.Path, FileMode.Open, FileAccess.Read, FileShare.Read)
                    If HashStream(stream) <> candidate.Hash OrElse ExcelCalculationBackend.ReadInternetZone(candidate.Path) <> sourceZone Then Throw New InvalidDataException("Candidate bytes or security marking changed after verification.")
                End Using
                SyncLock gate
                    RequireRevision(candidate.Revision)
                End SyncLock
                Return True
            End Function, cancellation)
            Await AwaitOperation(checking, "candidate validation").ConfigureAwait(False)
        End Function

        Private Shared Function HashStream(stream As Stream) As String
            stream.Position = 0
            Using hash = SHA256.Create()
                Return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "")
            End Using
        End Function
    End Class
End Namespace
