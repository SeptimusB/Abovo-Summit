Option Strict On

Imports System.IO
Imports System.Threading
Imports Abovo.WorkbookEngines

Namespace Abovo
    Partial Public NotInheritable Class ModelChangeManagerV2
        Private Function ComputeEngineDefinitionFingerprint() As String
            Dim parts = WB.CustomXmlParts.Select(Function(part) part.CustomXmlPartDocument.OuterXml).
                OrderBy(Function(xml) xml, StringComparer.Ordinal)
            Return ModelHistorySnapshot.Digest(String.Join(vbLf, parts))
        End Function

        'Internal until the ordinary save dialogs/recovery and remaining consumers
        'are qualified. Publication is terminal for the old native session; the
        'same engine/security policy is reopened before live dirty is cleared.
        Friend Function SaveEngineEditingModel(target As String, candidateDirectory As String,
                                               Optional cancellation As CancellationToken = Nothing) As WorkbookPublicationReceipt
            Dim model = RequireSaveHistoryOwner()
            If engineTrial Is Nothing Then Throw New InvalidOperationException("No native calculation owner is bound.")
            If IsReadOnlyPreview Then Throw New InvalidOperationException("A preview cannot save a model.")
            If ComputeEngineDefinitionFingerprint() <> engineDefinitionFingerprint Then
                Throw New InvalidOperationException("Model XML changed outside the native engine. Save is blocked until that change is reconciled.")
            End If
            target = IO.Path.GetFullPath(target)
            Dim sameSource = String.Equals(target, IO.Path.GetFullPath(model.FileName), StringComparison.OrdinalIgnoreCase)
            If Not sameSource AndAlso (File.Exists(target) OrElse Directory.Exists(target)) Then Throw New IOException("Choose a new filename; this destination already exists.")
            If Not String.Equals(IO.Path.GetExtension(target), IO.Path.GetExtension(model.FileName), StringComparison.OrdinalIgnoreCase) Then
                Throw New NotSupportedException("This native save path does not yet convert workbook formats.")
            End If
            If EngineEditingResult Is Nothing OrElse model.ResultsPending Then
                RecalculateEngine(If(model.NeedsFullRebuild, WorkbookCalculationKind.Rebuild, WorkbookCalculationKind.Full))
            End If
            Dim expected = EngineEditingResult
            Dim history = CaptureSaveHistory(expected)
            Dim previous = engineTrial
            Dim candidate As WorkbookSaveCandidate = Nothing
            Dim published As WorkbookPublicationReceipt = Nothing
            Dim replacement As WorkbookCalculationSession = Nothing
            writingAndCalculating = True
            Try
                Dim areas = Journal.AsEnumerable().Reverse().SelectMany(Function(group) group.Entries.AsEnumerable().Reverse()).
                    Where(Function(entry) entry.EngineBefore IsNot Nothing).
                    Select(Function(entry) New WorkbookReadArea(entry.WorksheetName,
                        WB.Worksheets(entry.WorksheetName).Cells(entry.CellAddress).RowIndex,
                        WB.Worksheets(entry.WorksheetName).Cells(entry.CellAddress).ColumnIndex, 1, 1)).
                    GroupBy(Function(area) EngineCellKey(area), StringComparer.OrdinalIgnoreCase).
                    Select(Function(group) group.First()).Take(64).ToList()
                If areas.Count = 0 Then
                    Dim first = expected.Blocks(0).Area
                    areas.Add(New WorkbookReadArea(first.Worksheet, first.Row, first.Column, 1, 1))
                End If
                Dim probes = previous.CaptureCellsAsync(expected.Revision, areas, cancellation).GetAwaiter().GetResult()
                candidate = previous.CreateSaveCandidateWithHistoryAsync(expected, candidateDirectory, probes, history, cancellation).GetAwaiter().GetResult()
                If model.UserChangeRevision <> history.UserRevision OrElse model.CalculationRevision <> history.CalculationRevision OrElse
                    ComputeEngineDefinitionFingerprint() <> engineDefinitionFingerprint Then Throw New InvalidOperationException("The model changed while preparing its save.")
                engineTrialResult = Nothing
                ModelEngineView.Publish(WB, Nothing)
                published = previous.PublishAndCloseAsync(candidate, target,
                    If(sameSource, WorkbookPublicationMode.ReplaceSource, WorkbookPublicationMode.CreateNew), cancellation).GetAwaiter().GetResult()
                'Once publication succeeded, cancellation cannot pretend nothing
                'was saved. Reopen/acknowledgement must complete or report the file.
                replacement = published.ReopenAsync().GetAwaiter().GetResult()
                Dim current = replacement.CalculateAndReadAsync(replacement.Revision, WorkbookCalculationKind.Rebuild,
                    engineTrialAreas, includePresentation:=engineTrialPresentation).GetAwaiter().GetResult()
                If Not IsSavedHistoryCurrentCore(model, history, published) OrElse ComputeEngineDefinitionFingerprint() <> engineDefinitionFingerprint Then
                    Throw New InvalidOperationException("The saved history no longer matches the live model. Dirty state has been retained.")
                End If
                ModelEngineView.ReplaceOwner(WB, previous, replacement, current)
                engineTrial = replacement : engineTrialResult = current
                replacement = Nothing 'The model now owns this session.
                model.AcknowledgeEngineSave(published.Path, history.UserRevision, history.CalculationRevision)
            Catch ex As Exception
                If replacement IsNot Nothing Then
                    Try
                        replacement.CloseAsync().GetAwaiter().GetResult()
                    Catch cleanup As Exception
                        ex = New AggregateException(ex, cleanup)
                    End Try
                End If
                If previous.IsCurrent(expected) Then
                    engineTrialResult = expected
                    ModelEngineView.Publish(WB, expected)
                End If
                Dim retained = If(published IsNot Nothing, published.Path, If(candidate IsNot Nothing, candidate.Path, Nothing))
                If retained IsNot Nothing Then
                    Throw New IOException(If(published IsNot Nothing,
                        "The workbook was saved but could not be resumed. Its saved file is retained at: ",
                        "Save did not complete. A verified copy of the unsaved work is retained at: ") & retained & Environment.NewLine & ex.Message, ex)
                End If
                Throw
            Finally
                writingAndCalculating = False
            End Try
            RefreshEngineConsumers()
            Return published
        End Function
    End Class
End Namespace
