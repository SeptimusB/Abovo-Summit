Option Strict On

Imports System.IO
Imports System.Threading
Imports Abovo.WorkbookEngines

Namespace Abovo
    Partial Public NotInheritable Class ModelChangeManagerV2
        Friend Sub WriteEngineRecoverySnapshot(output As Stream)
            Dim model = RequireSaveHistoryOwner()
            If engineTrial Is Nothing OrElse IsReadOnlyPreview OrElse Not model.HasUnsavedUserChanges Then Throw New InvalidOperationException("No edited native model is available for recovery.")
            If ComputeEngineDefinitionFingerprint() <> engineDefinitionFingerprint Then Throw New InvalidOperationException("Unrouted model XML cannot be included in a native recovery snapshot.")
            If EngineEditingResult Is Nothing OrElse model.ResultsPending Then
                RecalculateEngine(If(model.NeedsFullRebuild, WorkbookCalculationKind.Rebuild, WorkbookCalculationKind.Full))
            End If
            Dim expected = EngineEditingResult, history = CaptureSaveHistory(expected)
            Dim candidate As WorkbookSaveCandidate = Nothing
            writingAndCalculating = True
            Try
                Dim probes = engineTrial.CaptureCellsAsync(expected.Revision, EngineSaveAreas(expected)).GetAwaiter().GetResult()
                candidate = engineTrial.CreateRecoveryCandidateAsync(expected, IO.Path.GetDirectoryName(model.FileName), probes, history).GetAwaiter().GetResult()
                If model.UserChangeRevision <> history.UserRevision OrElse model.CalculationRevision <> history.CalculationRevision OrElse
                    ComputeEngineDefinitionFingerprint() <> engineDefinitionFingerprint Then Throw New InvalidOperationException("The model changed during recovery export.")
                Using input As New FileStream(candidate.Path, FileMode.Open, FileAccess.Read, FileShare.Read)
                    engineTrial.ValidateSaveCandidateAsync(candidate).GetAwaiter().GetResult()
                    input.CopyTo(output)
                End Using
            Finally
                writingAndCalculating = False
                'The existing recovery store owns final replacement. Only this
                'private verified staging file is removed; source/last recovery
                'and the model's saved revision are never acknowledged here.
                If candidate IsNot Nothing AndAlso File.Exists(candidate.Path) Then
                    Try
                        File.Delete(candidate.Path)
                        Dim folder = IO.Path.GetDirectoryName(candidate.Path)
                        If Not Directory.EnumerateFileSystemEntries(folder).Any() Then Directory.Delete(folder)
                    Catch cleanup As IOException
                        SummitDiagnostics.WriteLine("[Recovery] Native staging cleanup: " & cleanup.Message)
                    End Try
                End If
            End Try
        End Sub
    End Class
End Namespace
