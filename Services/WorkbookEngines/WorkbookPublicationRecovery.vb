Option Strict On

Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Xml
Imports System.Xml.Linq
Imports Microsoft.Win32.SafeHandles

Namespace Abovo.WorkbookEngines
    Public Enum WorkbookRecoveryState
        Prepared = 0
        OriginalRetained = 1
        PublishedUnacknowledged = 2
        PublishedAcknowledged = 3
        Conflict = 4
        Incomplete = 5
    End Enum

    ' A diagnostic snapshot, never an authenticated save receipt. The caller must
    ' independently select the intended target; a journal cannot choose it.
    Public NotInheritable Class WorkbookRecoveryInspection
        Public ReadOnly Property State As WorkbookRecoveryState
        Public ReadOnly Property JournalPath As String
        Public ReadOnly Property TargetPath As String
        Public ReadOnly Property CandidatePath As String
        Public ReadOnly Property BackupPath As String
        Public ReadOnly Property SessionId As Guid
        Public ReadOnly Property Revision As Long
        Friend ReadOnly Property JournalHash As String
        Friend Sub New(record As WorkbookRecoveryRecord, state As WorkbookRecoveryState)
            Me.State = state : JournalPath = record.Journal : JournalHash = record.Hash
            TargetPath = record.Target : CandidatePath = record.Candidate : BackupPath = record.Backup
            SessionId = record.SessionId : Revision = record.Revision
        End Sub
    End Class

    Public NotInheritable Class WorkbookOriginalRestoration
        Public ReadOnly Property Path As String
        Public ReadOnly Property Hash As String
        Public ReadOnly Property CandidatePath As String
        Public ReadOnly Property CompletionRecorded As Boolean
        Friend Sub New(record As WorkbookRecoveryRecord, recorded As Boolean)
            Path = record.Target : Hash = record.SourceHash : CandidatePath = record.Candidate : CompletionRecorded = recorded
        End Sub
    End Class

    Public NotInheritable Class WorkbookPublicationRecovery
        Private Sub New()
        End Sub

        ' Explicit path inspection only: no automatic disk scan or automatic
        ' write from unsigned XML. Malformed/legacy records fail closed.
        Public Shared Function InspectAsync(journalPath As String, expectedTarget As String,
                                            Optional cancellation As CancellationToken = Nothing) As Task(Of WorkbookRecoveryInspection)
            Return Task.Run(Function()
                                Using guarded = WorkbookRecoveryGuard.Open(journalPath, expectedTarget, False, cancellation)
                                    Return New WorkbookRecoveryInspection(guarded.Record, guarded.State)
                                End Using
                            End Function, cancellation)
        End Function

        ' Roll back only the vacant-name interruption. Never finish a save or
        ' replace a competing file based on a disk journal. Preserve candidate.
        Public Shared Function RestoreOriginalAsync(inspection As WorkbookRecoveryInspection,
                                                    Optional cancellation As CancellationToken = Nothing) As Task(Of WorkbookOriginalRestoration)
            Return RestoreCoreAsync(inspection, cancellation, Nothing)
        End Function

        Friend Shared Function RestoreCoreAsync(inspection As WorkbookRecoveryInspection, cancellation As CancellationToken,
                                                checkpoint As Action(Of String)) As Task(Of WorkbookOriginalRestoration)
            If inspection Is Nothing Then Throw New ArgumentNullException(NameOf(inspection))
            If inspection.State <> WorkbookRecoveryState.OriginalRetained Then Throw New InvalidOperationException("Only a verified original with a vacant destination can be restored.")
            Return Task.Run(Function()
                                Using guarded = WorkbookRecoveryGuard.Open(inspection.JournalPath, inspection.TargetPath, True, cancellation)
                                    Dim record = guarded.Record
                                    If record.Hash <> inspection.JournalHash OrElse guarded.State <> WorkbookRecoveryState.OriginalRetained Then Throw New IOException("Recovery evidence changed; inspect it again before taking action.")
                                    checkpoint?.Invoke("BeforeRestore")
                                    cancellation.ThrowIfCancellationRequested()
                                    guarded.RequireMarkers()
                                    guarded.Backup.RenameNew(record.Target)
                                    ' No cancellation after the rename. If writing
                                    ' the advisory receipt fails, still report the
                                    ' restoration rather than inviting blind retry.
                                    Dim recorded As Boolean = False
                                    Try
                                        checkpoint?.Invoke("Restored")
                                        Dim completion As New XDocument(New XElement("SummitOriginalRestored",
                                            New XAttribute("version", 1), New XElement("JournalHash", record.Hash),
                                            New XElement("Target", record.Target), New XElement("SourceHash", record.SourceHash)))
                                        Using output As New FileStream(IO.Path.Combine(IO.Path.GetDirectoryName(record.Journal), "restored.xml"), FileMode.CreateNew, FileAccess.Write, FileShare.Read)
                                            completion.Save(output) : output.Flush(True)
                                        End Using
                                        recorded = True
                                    Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
                                        ' File bytes are restored; advisory record
                                        ' failure must not turn this into a retry.
                                    End Try
                                    Return New WorkbookOriginalRestoration(record, recorded)
                                End Using
                            End Function, cancellation)
        End Function
    End Class

    Friend NotInheritable Class WorkbookRecoveryRecord
        Friend Journal As String, Hash As String, Target As String, Candidate As String, Backup As String
        Friend SourceHash As String, CandidateHash As String, MarkerHash As String
        Friend SessionId As Guid, Revision As Long, Mode As WorkbookPublicationMode
        Friend Xml As XDocument

        Friend Shared Function Read(journal As String, target As String) As WorkbookRecoveryRecord
            Dim xml = ReadXml(journal)
            Dim root = xml.Root
            Dim names = {"State", "Session", "Revision", "Mode", "Source", "SourceHash", "Target", "Candidate", "CandidateHash", "Backup", "MarkerHash"}
            If root Is Nothing OrElse root.Name <> XName.Get("SummitPublication") OrElse root.Attributes().Count() <> 1 OrElse
                CStr(root.Attribute("version")) <> "2" OrElse root.Elements().Count() <> names.Length OrElse
                names.Any(Function(name) root.Elements(name).Count() <> 1) OrElse
                root.Elements().Any(Function(element) element.HasElements OrElse element.HasAttributes) Then
                Throw New InvalidDataException("Unrecognized recovery record. Older records require manual inspection.")
            End If
            If root.Element("State").Value <> "Prepared" OrElse Not SamePath(root.Element("Target").Value, target) Then Throw New InvalidDataException("The recovery record does not describe the selected target.")
            Dim result As New WorkbookRecoveryRecord With {.Journal = journal, .Target = target, .Xml = xml,
                .SourceHash = root.Element("SourceHash").Value, .CandidateHash = root.Element("CandidateHash").Value,
                .MarkerHash = root.Element("MarkerHash").Value}
            If Not Guid.TryParseExact(root.Element("Session").Value, "D", result.SessionId) OrElse result.SessionId = Guid.Empty OrElse
                Not Long.TryParse(root.Element("Revision").Value, Globalization.NumberStyles.None, Globalization.CultureInfo.InvariantCulture, result.Revision) OrElse
                result.Revision < 0 Then Throw New InvalidDataException("Invalid recovery session or revision.")
            For Each digest In {result.SourceHash, result.CandidateHash, result.MarkerHash}
                If Not Text.RegularExpressions.Regex.IsMatch(digest, "\A[0-9A-F]{64}\z") Then Throw New InvalidDataException("Invalid recovery digest.")
            Next
            Select Case root.Element("Mode").Value
                Case "ReplaceSource" : result.Mode = WorkbookPublicationMode.ReplaceSource
                Case "CreateNew" : result.Mode = WorkbookPublicationMode.CreateNew
                Case Else : Throw New InvalidDataException("Unknown recovery mode.")
            End Select
            Dim source = WorkbookPublicationFile.Normalize(root.Element("Source").Value)
            Dim extension = IO.Path.GetExtension(target)
            If Not {".xlsx", ".xlsm", ".xlsb"}.Contains(extension, StringComparer.OrdinalIgnoreCase) OrElse
                Not String.Equals(extension, IO.Path.GetExtension(source), StringComparison.OrdinalIgnoreCase) OrElse
                SamePath(source, target) <> (result.Mode = WorkbookPublicationMode.ReplaceSource) Then Throw New InvalidDataException("Recovery mode or format does not match the target.")
            result.Candidate = IO.Path.Combine(IO.Path.GetDirectoryName(journal), "candidate" & extension)
            result.Backup = If(result.Mode = WorkbookPublicationMode.ReplaceSource, IO.Path.Combine(IO.Path.GetDirectoryName(journal), "previous" & extension), Nothing)
            If Not SamePath(root.Element("Candidate").Value, result.Candidate) OrElse
                Not String.Equals(root.Element("Backup").Value, If(result.Backup, ""), StringComparison.OrdinalIgnoreCase) Then Throw New InvalidDataException("Recovery payload paths must stay in their own transaction folder.")
            Return result
        End Function

        Friend Shared Function SamePath(left As String, right As String) As Boolean
            ' Require canonical paths, not '..', ADS or aliases in XML.
            Return String.Equals(left, right, StringComparison.OrdinalIgnoreCase)
        End Function

        Friend Shared Function ReadXml(path As String) As XDocument
            Dim settings As New XmlReaderSettings With {.DtdProcessing = DtdProcessing.Prohibit, .XmlResolver = Nothing, .MaxCharactersInDocument = 16384}
            Using input As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)
                If input.Length > 16384 Then Throw New InvalidDataException("Recovery record exceeds the size limit.")
                Using reader = XmlReader.Create(input, settings)
                    Return XDocument.Load(reader)
                End Using
            End Using
        End Function
    End Class

    Friend NotInheritable Class WorkbookRecoveryGuard
        Implements IDisposable
        Private pins As List(Of SafeFileHandle)
        Private ReadOnly files As New Dictionary(Of String, WorkbookPublicationFile)(StringComparer.OrdinalIgnoreCase)
        Friend Record As WorkbookRecoveryRecord
        Friend State As WorkbookRecoveryState
        Friend ReadOnly Property Backup As WorkbookPublicationFile
            Get
                Return files(Record.Backup)
            End Get
        End Property

        Friend Shared Function Open(journalPath As String, expectedTarget As String, restore As Boolean, cancellation As CancellationToken) As WorkbookRecoveryGuard
            cancellation.ThrowIfCancellationRequested()
            Dim journal = WorkbookPublicationFile.Normalize(journalPath)
            Dim target = WorkbookPublicationFile.Normalize(expectedTarget)
            Dim folder = IO.Path.GetDirectoryName(journal)
            If IO.Path.GetFileName(journal) <> "intent.xml" OrElse
                Not String.Equals(IO.Path.GetDirectoryName(folder), IO.Path.GetDirectoryName(target), StringComparison.OrdinalIgnoreCase) OrElse
                Not Text.RegularExpressions.Regex.IsMatch(IO.Path.GetFileName(folder), "\A~Summit-save-[0-9a-f]{32}\z") Then Throw New ArgumentException("Select an intent record directly beside the intended workbook, in its generated recovery folder.")
            Dim guarded As New WorkbookRecoveryGuard()
            Try
                guarded.pins = WorkbookPublicationFile.PinDirectories({journal, target})
                Dim intent = guarded.AddFile(journal, False)
                If intent Is Nothing Then Throw New FileNotFoundException("Recovery intent record is missing.", journal)
                guarded.Record = WorkbookRecoveryRecord.Read(journal, target)
                guarded.Record.Hash = intent.Hash()
                Dim record = guarded.Record
                Dim candidate = guarded.AddFile(record.Candidate, False)
                Dim backup = guarded.AddFile(record.Backup, restore)
                Dim destination = guarded.AddFile(record.Target, False)
                Dim completedPath = IO.Path.Combine(folder, "published.xml")
                Dim completed = guarded.AddFile(completedPath, False)
                If completed IsNot Nothing Then
                    Dim expected = New XDocument(record.Xml)
                    expected.Root.Element("State").Value = "Published"
                    If Not XNode.DeepEquals(expected, WorkbookRecoveryRecord.ReadXml(completedPath)) Then Throw New InvalidDataException("Publication completion record disagrees with intent.")
                End If
                cancellation.ThrowIfCancellationRequested()
                Dim candidateOK = guarded.Matches(candidate, record.Candidate, record.CandidateHash)
                Dim backupOK = guarded.Matches(backup, record.Backup, record.SourceHash)
                Dim originalOK = guarded.Matches(destination, record.Target, record.SourceHash)
                Dim publishedOK = guarded.Matches(destination, record.Target, record.CandidateHash)
                guarded.State = WorkbookRecoveryState.Incomplete
                If (candidate IsNot Nothing AndAlso Not candidateOK) OrElse (backup IsNot Nothing AndAlso Not backupOK) OrElse
                    Directory.Exists(record.Target) OrElse Directory.Exists(record.Candidate) OrElse (record.Backup IsNot Nothing AndAlso Directory.Exists(record.Backup)) Then
                    guarded.State = WorkbookRecoveryState.Conflict
                ElseIf publishedOK AndAlso candidate Is Nothing AndAlso (record.Mode = WorkbookPublicationMode.CreateNew OrElse backupOK) Then
                    guarded.State = If(completed Is Nothing, WorkbookRecoveryState.PublishedUnacknowledged, WorkbookRecoveryState.PublishedAcknowledged)
                ElseIf completed IsNot Nothing Then
                    guarded.State = WorkbookRecoveryState.Conflict
                ElseIf record.Mode = WorkbookPublicationMode.ReplaceSource AndAlso destination Is Nothing AndAlso backupOK AndAlso candidateOK Then
                    guarded.State = WorkbookRecoveryState.OriginalRetained
                ElseIf candidateOK AndAlso backup Is Nothing AndAlso
                    ((record.Mode = WorkbookPublicationMode.CreateNew AndAlso destination Is Nothing) OrElse (record.Mode = WorkbookPublicationMode.ReplaceSource AndAlso originalOK)) Then
                    guarded.State = WorkbookRecoveryState.Prepared
                ElseIf destination IsNot Nothing Then
                    guarded.State = WorkbookRecoveryState.Conflict
                End If
                Return guarded
            Catch
                guarded.Dispose() : Throw
            End Try
        End Function

        Private Function AddFile(path As String, rename As Boolean) As WorkbookPublicationFile
            If path Is Nothing OrElse Not IO.File.Exists(path) Then Return Nothing
            Dim file As New WorkbookPublicationFile(path, rename)
            files.Add(path, file)
            Return file
        End Function

        Private Function Matches(file As WorkbookPublicationFile, path As String, hash As String) As Boolean
            Return file IsNot Nothing AndAlso file.Hash() = hash AndAlso
                WorkbookPublicationFile.MarkerHash(WorkbookPublicationFile.CaptureMarkers(path)) = Record.MarkerHash
        End Function

        Friend Sub RequireMarkers()
            If Not Matches(Backup, Record.Backup, Record.SourceHash) OrElse
                Not Matches(files(Record.Candidate), Record.Candidate, Record.CandidateHash) Then Throw New IOException("Recovery payload changed before restoration.")
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            For Each file In files.Values
                file.Dispose()
            Next
            files.Clear()
            If pins IsNot Nothing Then
                For Each pin In pins
                    pin.Dispose()
                Next
                pins = Nothing
            End If
        End Sub
    End Class
End Namespace
