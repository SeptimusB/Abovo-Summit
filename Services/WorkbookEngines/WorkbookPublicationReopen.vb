Option Strict On

Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks

Namespace Abovo.WorkbookEngines
    Friend NotInheritable Class WorkbookPublicationReopen
        ' Only an in-memory publication receipt enters this path. Unsigned disk
        ' journals cannot grant a trusted reopen or reconstruct engine/security
        ' options. A recovered original goes through ordinary opening policy.
        Friend Shared Async Function OpenAsync(receipt As WorkbookPublicationReceipt, cancellation As CancellationToken,
                                               factory As Func(Of WorkbookEnginePreference, IWorkbookCalculationBackend)) As Task(Of WorkbookCalculationSession)
            If receipt Is Nothing Then Throw New ArgumentNullException(NameOf(receipt))
            cancellation.ThrowIfCancellationRequested()
            Dim pins = WorkbookPublicationFile.PinDirectories({receipt.Path})
            Dim session As WorkbookCalculationSession = Nothing
            Dim failure As Exception = Nothing
            Try
                Using file As New WorkbookPublicationFile(receipt.Path, False)
                    If file.Hash() <> receipt.Hash OrElse
                        WorkbookPublicationFile.MarkerHash(WorkbookPublicationFile.CaptureMarkers(receipt.Path)) <> receipt.MarkerHash Then
                        Throw New IOException("The saved workbook changed before reopening. Open it as a separate file; do not acknowledge this saved revision.")
                    End If
                    Try
                        session = Await WorkbookCalculationSession.OpenAsync(receipt.Path, receipt.OpeningOptions, cancellation, factory).ConfigureAwait(False)
                        cancellation.ThrowIfCancellationRequested()
                        If session.SourceHash <> receipt.Hash OrElse
                            WorkbookPublicationFile.MarkerHash(WorkbookPublicationFile.CaptureMarkers(receipt.Path)) <> receipt.MarkerHash Then
                            Throw New IOException("Saved workbook provenance changed while reopening.")
                        End If
                    Catch ex As Exception
                        failure = ex
                    End Try
                End Using
            Finally
                For Each pin In pins
                    pin.Dispose()
                Next
            End Try
            If failure IsNot Nothing Then
                If session IsNot Nothing Then
                    Try
                        Await session.CloseAsync().ConfigureAwait(False)
                    Catch cleanupError As Exception
                        Throw New AggregateException("The file is saved, but reopening and cleanup failed.", failure, cleanupError)
                    End Try
                End If
                Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw()
            End If
            Return session
        End Function
    End Class
End Namespace
