Imports System
Imports System.Collections.Generic
Imports DevExpress.Spreadsheet

Namespace Abovo

    Public Enum ModelIntegrityState
        Healthy = 0
        RecoveryRequired = 1
    End Enum

    Public NotInheritable Class ModelSafetyManager

        Private Shared ReadOnly BulkMutationCounts As New Dictionary(Of Integer, Integer)

        Private Sub New()
        End Sub

        Public Shared Sub BeginBulkWorkbookMutation(ByVal modelID As Integer)
            SyncLock BulkMutationCounts
                Dim count As Integer = 0
                BulkMutationCounts.TryGetValue(modelID, count)
                BulkMutationCounts(modelID) = count + 1
            End SyncLock
        End Sub

        Public Shared Sub EndBulkWorkbookMutation(ByVal modelID As Integer)
            SyncLock BulkMutationCounts
                Dim count As Integer = 0
                If Not BulkMutationCounts.TryGetValue(modelID, count) Then Return

                If count <= 1 Then
                    BulkMutationCounts.Remove(modelID)
                Else
                    BulkMutationCounts(modelID) = count - 1
                End If
            End SyncLock
        End Sub

        Public Shared Function IsBulkWorkbookMutationInProgress(ByVal modelID As Integer) As Boolean
            SyncLock BulkMutationCounts
                Dim count As Integer = 0
                Return BulkMutationCounts.TryGetValue(modelID, count) AndAlso count > 0
            End SyncLock
        End Function

        Public Shared Sub MarkRecoveryRequired(ByVal modelID As Integer,
                                               ByVal operation As String,
                                               ByVal reason As String,
                                               Optional ByVal source As String = "Workbook operation",
                                               Optional ByVal location As String = "")
            Dim model As FileManager.ExcelModel = GetModel(modelID)
            If model IsNot Nothing Then
                model.IntegrityState = ModelIntegrityState.RecoveryRequired
                model.RecoverySaveAsRequired = True
                model.IntegrityReason = Normalise(reason)
                model.IntegrityOperation = Normalise(operation)
                model.IsDirty = True
            End If
            SystemMessageManager.Publish(modelID, BuildRecoveryMessage(operation, reason),
                SystemMessageSeverity.Error, source, location)
        End Sub

        Public Shared Sub PublishResult(ByVal modelID As Integer,
                                        ByVal operation As String,
                                        ByVal result As AbovoAppCls.AbovoTransaction,
                                        Optional ByVal source As String = "Workbook operation",
                                        Optional ByVal location As String = "")
            If result Is Nothing Then
                SystemMessageManager.Publish(modelID, Normalise(operation) & " returned no transaction result.",
                    SystemMessageSeverity.Error, source, location)
                Return
            End If
            If result.EventCancelled Then
                SystemMessageManager.Publish(modelID, FirstMessage(result, Normalise(operation) & " was cancelled."),
                    SystemMessageSeverity.Information, source, location)
            ElseIf result.BError Then
                SystemMessageManager.Publish(modelID, FirstMessage(result, Normalise(operation) & " failed."),
                    SystemMessageSeverity.Error, source, location)
            ElseIf result.BSuccess Then
                SystemMessageManager.Publish(modelID, FirstMessage(result, Normalise(operation) & " completed."),
                    SystemMessageSeverity.Success, source, location)
            End If
        End Sub

        Public Shared Function RequiresRecoverySaveAs(ByVal modelID As Integer) As Boolean
            Dim model As FileManager.ExcelModel = GetModel(modelID)
            Return model IsNot Nothing AndAlso model.RecoverySaveAsRequired
        End Function

        Private Shared Function GetModel(ByVal modelID As Integer) As FileManager.ExcelModel
            If FileManager.ExcelModels Is Nothing OrElse modelID < 0 OrElse
               modelID >= FileManager.ExcelModels.Length Then Return Nothing
            Return FileManager.ExcelModels(modelID)
        End Function

        Private Shared Function BuildRecoveryMessage(ByVal operation As String,
                                                     ByVal reason As String) As String
            Return Normalise(operation) & " failed after workbook changes may have begun. " &
                   "Rollback could not be verified. Do not overwrite the original workbook; " &
                   "use Save As to create a recovery copy." &
                   If(String.IsNullOrWhiteSpace(reason), String.Empty, " Details: " & Normalise(reason))
        End Function

        Private Shared Function FirstMessage(ByVal result As AbovoAppCls.AbovoTransaction,
                                             ByVal fallback As String) As String
            If Not String.IsNullOrWhiteSpace(result.StrResponseMessage) AndAlso
               Not String.Equals(result.StrResponseMessage, "Initiated", StringComparison.OrdinalIgnoreCase) Then
                Return result.StrResponseMessage.Trim()
            End If
            If Not String.IsNullOrWhiteSpace(result.StringReturn) Then Return result.StringReturn.Trim()
            Return fallback
        End Function

        Private Shared Function Normalise(ByVal value As String) As String
            Return If(value, String.Empty).Trim()
        End Function

    End Class

    Public NotInheritable Class WorkbookCellJournal
        Implements IDisposable

        Private ReadOnly Entries As New List(Of WorkbookRangeJournalEntry)
        Private ReadOnly BackupWorkbook As New Workbook()
        Private IsDisposed As Boolean

        Public Sub New()
            BackupWorkbook.Options.CalculationMode = WorkbookCalculationMode.Manual
            BackupWorkbook.History.IsEnabled = False
        End Sub

        Public Sub Capture(ByVal range As CellRange)
            If range Is Nothing OrElse range.Worksheet Is Nothing Then
                Throw New ArgumentNullException(NameOf(range))
            End If

            If IsDisposed Then Throw New ObjectDisposedException(NameOf(WorkbookCellJournal))

            'Keep the backup at identical worksheet coordinates. This lets
            'CopyFrom preserve relative formulas without translating them to A1,
            'while performing one native range copy instead of thousands of
            'DisplayText/value/formula calls through the managed API.
            Dim backupSheet As Worksheet = BackupWorkbook.Worksheets.Add()
            Dim backupRange As CellRange = backupSheet.Range.FromLTRB(
                range.LeftColumnIndex,
                range.TopRowIndex,
                range.RightColumnIndex,
                range.BottomRowIndex)
            backupRange.CopyFrom(range, PasteSpecial.Formulas Or PasteSpecial.Values)

            Entries.Add(New WorkbookRangeJournalEntry With {
                .Worksheet = range.Worksheet,
                .LeftColumnIndex = range.LeftColumnIndex,
                .TopRowIndex = range.TopRowIndex,
                .RightColumnIndex = range.RightColumnIndex,
                .BottomRowIndex = range.BottomRowIndex,
                .Snapshot = backupRange})
        End Sub

        Public Function Restore(ByRef failureDetails As String) As Boolean
            Dim failures As New List(Of String)

            For index As Integer = Entries.Count - 1 To 0 Step -1
                Dim entry As WorkbookRangeJournalEntry = Entries(index)
                Try
                    Dim target As CellRange = entry.TargetRange()
                    target.ClearContents()
                    target.CopyFrom(entry.Snapshot, PasteSpecial.Formulas Or PasteSpecial.Values)
                Catch ex As Exception
                    failures.Add(entry.Description() & ": " & ex.Message)
                End Try
            Next

            'Verification is deliberately deferred to the failure path. Successful
            'imports pay only for native bulk range copies.
            For Each entry As WorkbookRangeJournalEntry In Entries
                Try
                    Dim target As CellRange = entry.TargetRange()
                    For rowOffset As Integer = 0 To entry.Snapshot.RowCount - 1
                        For columnOffset As Integer = 0 To entry.Snapshot.ColumnCount - 1
                            Dim expected As CellSnapshotV2 =
                                CellSnapshotV2.Capture(entry.Snapshot(rowOffset, columnOffset))
                            If Not expected.Matches(target(rowOffset, columnOffset)) Then
                                failures.Add(entry.Description() & ": restored value did not verify.")
                                Exit For
                            End If
                        Next
                        If failures.Count > 0 Then Exit For
                    Next
                Catch ex As Exception
                    failures.Add("Rollback verification for " & entry.Description() & ": " & ex.Message)
                End Try
            Next

            failureDetails = String.Join(Environment.NewLine, failures)
            Return failures.Count = 0
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If IsDisposed Then Return
            IsDisposed = True
            BackupWorkbook.Dispose()
        End Sub

        Private NotInheritable Class WorkbookRangeJournalEntry
            Public Worksheet As Worksheet
            Public LeftColumnIndex As Integer
            Public TopRowIndex As Integer
            Public RightColumnIndex As Integer
            Public BottomRowIndex As Integer
            Public Snapshot As CellRange

            Public Function TargetRange() As CellRange
                Return Worksheet.Range.FromLTRB(
                    LeftColumnIndex,
                    TopRowIndex,
                    RightColumnIndex,
                    BottomRowIndex)
            End Function

            Public Function Description() As String
                Return Worksheet.Name & "!" & TargetRange().GetReferenceA1()
            End Function
        End Class

    End Class

End Namespace
