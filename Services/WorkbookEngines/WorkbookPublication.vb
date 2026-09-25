Option Strict On

Imports System.IO
Imports System.Security.AccessControl
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Xml.Linq
Imports Microsoft.Win32.SafeHandles

Namespace Abovo.WorkbookEngines
    Public Enum WorkbookPublicationMode
        CreateNew = 0
        ReplaceSource = 1
    End Enum

    ' Terminal hand-off only. This is not a production dirty/history acknowledgement.
    Public NotInheritable Class WorkbookPublicationReceipt
        Public ReadOnly Property SessionId As Guid
        Public ReadOnly Property Revision As Long
        Public ReadOnly Property Path As String
        Public ReadOnly Property Hash As String
        Public ReadOnly Property BackupPath As String
        Public ReadOnly Property JournalPath As String
        Public ReadOnly Property MarkerHash As String
        Friend ReadOnly Property OpeningOptions As WorkbookEngineOptions
        Friend Sub New(candidate As WorkbookSaveCandidate, path As String, backup As String, journal As String,
                       markerHash As String, options As WorkbookEngineOptions)
            SessionId = candidate.SessionId : Revision = candidate.Revision : Hash = candidate.Hash
            Me.Path = path : BackupPath = backup : JournalPath = journal
            Me.MarkerHash = markerHash : OpeningOptions = options
        End Sub

        ' A new owner/baseline, not continuation of the old revision counter and
        ' not permission to mark a live model or its newer history clean.
        Public Function ReopenAsync(Optional cancellation As CancellationToken = Nothing) As Task(Of WorkbookCalculationSession)
            Return WorkbookPublicationReopen.OpenAsync(Me, cancellation, Nothing)
        End Function
    End Class

    Public NotInheritable Class WorkbookPublicationException
        Inherits IOException
        Public ReadOnly Property CandidatePath As String
        Public ReadOnly Property TargetPath As String
        Public ReadOnly Property BackupPath As String
        Public ReadOnly Property JournalPath As String
        Public ReadOnly Property Published As Boolean
        Friend Sub New(cause As Exception, candidate As String, target As String, backup As String, journal As String, published As Boolean)
            MyBase.New(If(published, "The workbook was written, but save acknowledgement failed. Do not retry blindly.",
                          "The workbook save was not acknowledged. Retained candidate and recovery paths must be checked before retrying."), cause)
            CandidatePath = candidate : TargetPath = target : BackupPath = backup : JournalPath = journal : Me.Published = published
        End Sub
    End Class

    Partial Public NotInheritable Class WorkbookCalculationSession
        Public Function PublishAndCloseAsync(candidate As WorkbookSaveCandidate, target As String, mode As WorkbookPublicationMode,
                                              Optional cancellation As CancellationToken = Nothing) As Task(Of WorkbookPublicationReceipt)
            Return PublishAndCloseCoreAsync(candidate, target, mode, cancellation, Nothing)
        End Function

        ' Per-call fault seam for the isolated harness, never a global live hook.
        Friend Async Function PublishAndCloseCoreAsync(candidate As WorkbookSaveCandidate, target As String, mode As WorkbookPublicationMode,
                                                       cancellation As CancellationToken, checkpoint As Action(Of String)) As Task(Of WorkbookPublicationReceipt)
            If Not publicationTrial Then Throw New InvalidOperationException("Publication trials were not enabled for this session.")
            If Not [Enum].IsDefined(GetType(WorkbookPublicationMode), mode) Then Throw New ArgumentOutOfRangeException(NameOf(mode))
            If Not IsCurrentCandidate(candidate) Then Throw New InvalidOperationException("A current save candidate is required.")
            target = WorkbookPublicationFile.Normalize(target)
            If Not String.Equals(IO.Path.GetExtension(target), IO.Path.GetExtension(sourcePath), StringComparison.OrdinalIgnoreCase) Then Throw New ArgumentException("Publication cannot convert workbook formats.")
            Dim sameSource = String.Equals(target, sourcePath, StringComparison.OrdinalIgnoreCase)
            If sameSource <> (mode = WorkbookPublicationMode.ReplaceSource) Then Throw New ArgumentException("ReplaceSource must name the current source; CreateNew must name a different file.")
            cancellation.ThrowIfCancellationRequested()
            SyncLock gate
                RequireRevision(candidate.Revision)
                If editPending OrElse savePending Then Throw New InvalidOperationException("A workbook edit or save is already in progress.")
                savePending = True
            End SyncLock
            Dim transaction As WorkbookPublicationTransaction = Nothing
            Try
                transaction = Await Task.Run(Function() WorkbookPublicationTransaction.Prepare(sourcePath, sourceZone, publicationStreams, candidate, target, mode, openingOptions, cancellation, checkpoint)).ConfigureAwait(False)
                checkpoint?.Invoke("Prepared")
                cancellation.ThrowIfCancellationRequested()
                SyncLock gate
                    RequireRevision(candidate.Revision)
                    ' Close the owner before releasing its source lease. No new
                    ' work may be accepted once this terminal boundary is taken.
                    Dim pendingCleanup = StartCleanup()
                End SyncLock
                Await CloseAsync().ConfigureAwait(False)
                checkpoint?.Invoke("Closed")
                ' Never put publication behind AwaitOperation's abandon-on-timeout
                ' wrapper: a filesystem rename may already have committed. Native
                ' close above is bounded and MUST actually succeed before this runs.
                Return Await Task.Run(Function() transaction.Commit(cancellation, checkpoint)).ConfigureAwait(False)
            Finally
                If transaction IsNot Nothing Then transaction.Dispose()
                SyncLock gate
                    savePending = False
                End SyncLock
            End Try
        End Function
    End Class

    Friend NotInheritable Class WorkbookPublicationTransaction
        Implements IDisposable
        Private pins As List(Of SafeFileHandle)
        Private source As String
        Private zone As String
        Private markers As Dictionary(Of String, Byte())
        Private candidate As WorkbookSaveCandidate
        Private target As String
        Private stage As String
        Private backup As String
        Private journal As String
        Private mode As WorkbookPublicationMode
        Private options As WorkbookEngineOptions

        Friend Shared Function Prepare(source As String, zone As String, markers As Dictionary(Of String, Byte()), candidate As WorkbookSaveCandidate, target As String,
                                       mode As WorkbookPublicationMode, options As WorkbookEngineOptions, cancellation As CancellationToken, checkpoint As Action(Of String)) As WorkbookPublicationTransaction
            Dim transaction As New WorkbookPublicationTransaction With {.source = source, .zone = zone, .markers = markers, .candidate = candidate, .target = target, .mode = mode, .options = options}
            Try
                transaction.pins = WorkbookPublicationFile.PinDirectories({source, candidate.Path, target})
                If Not WorkbookPublicationFile.MarkersMatch(source, markers) Then Throw New IOException("Source provenance marking changed since opening.")
                If mode = WorkbookPublicationMode.CreateNew AndAlso (File.Exists(target) OrElse Directory.Exists(target)) Then Throw New IOException("Save As destination already exists; choose a new filename.")
                Using input As New FileStream(candidate.Path, FileMode.Open, FileAccess.Read, FileShare.Read)
                    WorkbookPublicationFile.RequireSupportedStreams(candidate.Path)
                    Dim candidateMarkers = WorkbookPublicationFile.CaptureMarkers(candidate.Path)
                    For Each pair In candidateMarkers
                        If String.Equals(pair.Key, ":Zone.Identifier:$DATA", StringComparison.OrdinalIgnoreCase) Then Continue For ' Existing candidate API uses text-normalized Windows marker.
                        If Not markers.ContainsKey(pair.Key) OrElse Not markers(pair.Key).SequenceEqual(pair.Value) Then Throw New IOException("Candidate provenance marking differs from the source.")
                    Next
                    If WorkbookCalculationSessionHash(input) <> candidate.Hash OrElse ExcelCalculationBackend.ReadInternetZone(candidate.Path) <> zone Then Throw New InvalidDataException("Candidate changed after verification.")
                    cancellation.ThrowIfCancellationRequested()
                    Dim folder = IO.Path.Combine(IO.Path.GetDirectoryName(target), "~Summit-save-" & Guid.NewGuid().ToString("N"))
                    transaction.stage = IO.Path.Combine(folder, "candidate" & IO.Path.GetExtension(source))
                    transaction.backup = If(mode = WorkbookPublicationMode.ReplaceSource, IO.Path.Combine(folder, "previous" & IO.Path.GetExtension(source)), Nothing)
                    transaction.journal = IO.Path.Combine(folder, "intent.xml")
                    If transaction.stage.Length >= 248 Then Throw New ArgumentException("Publication recovery path is too long.")
                    Directory.CreateDirectory(folder)
                    transaction.pins.AddRange(WorkbookPublicationFile.PinDirectories({transaction.stage}))
                    checkpoint?.Invoke("Staging")
                    Using output As New FileStream(transaction.stage, FileMode.CreateNew, FileAccess.Write, FileShare.None)
                        input.Position = 0 : input.CopyTo(output) : output.Flush(True)
                    End Using
                    WorkbookPublicationFile.WriteMarkers(transaction.stage, markers)
                End Using
                transaction.WriteRecord(transaction.journal, "Prepared")
                Return transaction
            Catch
                transaction.Dispose()
                ' Retain partial staging only; never touch the target on prepare.
                Throw
            End Try
        End Function

        Private Shared Function WorkbookCalculationSessionHash(stream As Stream) As String
            stream.Position = 0
            Using hash = Security.Cryptography.SHA256.Create()
                Return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "")
            End Using
        End Function

        Friend Function Commit(cancellation As CancellationToken, checkpoint As Action(Of String)) As WorkbookPublicationReceipt
            Dim retained As Boolean = False, published As Boolean = False
            Try
                Using original As New WorkbookPublicationFile(source, mode = WorkbookPublicationMode.ReplaceSource)
                    If original.Hash() <> candidate.SourceHash OrElse Not WorkbookPublicationFile.MarkersMatch(source, markers) Then Throw New IOException("The source workbook changed during hand-off. It has not been replaced.")
                    If mode = WorkbookPublicationMode.ReplaceSource Then
                        ' Preserve the original discretionary access policy. No
                        ' permission elevation or ignore-ACL-error fallback.
                        Dim originalAccess = File.GetAccessControl(source, AccessControlSections.Access).GetSecurityDescriptorSddlForm(AccessControlSections.Access)
                        Dim stageAccess As New FileSecurity()
                        ' A freshly read FileSecurity has no modified sections;
                        ' passing it straight to SetAccessControl can be a no-op.
                        stageAccess.SetSecurityDescriptorSddlForm(originalAccess, AccessControlSections.Access)
                        File.SetAccessControl(stage, stageAccess)
                        If File.GetAccessControl(stage, AccessControlSections.Access).GetSecurityDescriptorSddlForm(AccessControlSections.Access) <> originalAccess Then Throw New IOException("Replacement permissions could not be preserved.")
                    End If
                    Using replacement As New WorkbookPublicationFile(stage, True)
                        If replacement.Hash() <> candidate.Hash OrElse Not WorkbookPublicationFile.MarkersMatch(stage, markers) Then Throw New InvalidDataException("Staged workbook changed before publication.")
                        checkpoint?.Invoke("BeforeCommit")
                        cancellation.ThrowIfCancellationRequested()
                        ' Security streams can have separate sharing semantics.
                        ' Recheck them at the boundary, not just at preparation.
                        WorkbookPublicationFile.RequireSupportedStreams(source)
                        WorkbookPublicationFile.RequireSupportedStreams(stage)
                        If Not WorkbookPublicationFile.MarkersMatch(source, markers) OrElse Not WorkbookPublicationFile.MarkersMatch(stage, markers) Then Throw New IOException("Workbook security marking changed before publication.")
                        Try
                            If mode = WorkbookPublicationMode.ReplaceSource Then
                                original.RenameNew(backup) : retained = True
                                checkpoint?.Invoke("OriginalRetained")
                            End If
                            ' No cancellation or timeout may discard a successful
                            ' rename. This is the irreversible acknowledgement unit.
                            replacement.RenameNew(target) : published = True
                            checkpoint?.Invoke("Published")
                            If Not WorkbookPublicationFile.MarkersMatch(target, markers) Then Throw New IOException("Published security marking changed; acknowledgement refused.")
                            WriteRecord(IO.Path.Combine(IO.Path.GetDirectoryName(journal), "published.xml"), "Published")
                            Return New WorkbookPublicationReceipt(candidate, target, backup, journal, WorkbookPublicationFile.MarkerHash(markers), options)
                        Catch
                            If retained AndAlso Not published Then
                                Try
                                    ' Restore only if the original name is still
                                    ' vacant. Never clobber another writer's file.
                                    original.RenameNew(source)
                                Catch
                                    ' The verified original remains at backup.
                                End Try
                            End If
                            Throw
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                Throw New WorkbookPublicationException(ex, stage, target, backup, journal, published)
            End Try
        End Function

        Private Sub WriteRecord(path As String, state As String)
            Dim xml As New XDocument(New XElement("SummitPublication", New XAttribute("version", 2),
                New XElement("State", state), New XElement("Session", candidate.SessionId), New XElement("Revision", candidate.Revision),
                New XElement("Mode", mode.ToString()), New XElement("Source", source), New XElement("SourceHash", candidate.SourceHash),
                New XElement("Target", target), New XElement("Candidate", stage), New XElement("CandidateHash", candidate.Hash), New XElement("Backup", backup),
                New XElement("MarkerHash", WorkbookPublicationFile.MarkerHash(markers))))
            Using file As New FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read)
                xml.Save(file) : file.Flush(True)
            End Using
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If pins IsNot Nothing Then
                For Each pin In pins
                    pin.Dispose()
                Next
                pins = Nothing
            End If
        End Sub
    End Class
End Namespace
