Imports System.Globalization
Imports DevExpress.Spreadsheet

Namespace Abovo

    Public Enum ChangeHistoryStateV2
        Applied
        Undone
        Superseded
    End Enum

    Public NotInheritable Class ChangeHistoryChangedEventArgsV2
        Inherits EventArgs
        Public ReadOnly Property IsUndoRedo As Boolean
        Public ReadOnly Property WorksheetNames As IReadOnlyCollection(Of String)

        Public Sub New(ByVal undoRedo As Boolean, ByVal sheets As IEnumerable(Of String))
            IsUndoRedo = undoRedo
            WorksheetNames = New List(Of String)(If(sheets, Enumerable.Empty(Of String)))
        End Sub
    End Class

    Partial Public NotInheritable Class ModelChangeManagerV2
        Private ReadOnly WB As IWorkbook
        Private ReadOnly UndoStack As New List(Of ChangeHistoryGroupV2)()
        Private ReadOnly RedoStack As New List(Of ChangeHistoryGroupV2)()
        Private ReadOnly Journal As New List(Of ChangeHistoryGroupV2)()
        Private NextGroupID As Integer
        Private ActiveGroup As ChangeHistoryGroupV2
        Private ActiveGroupDepth As Integer
        Private IsApplyingHistory As Boolean
        Private writingAndCalculating As Boolean
        Private RecoveryHistory As DataTable
        Friend Sub RestoreRecoveryHistory(table As DataTable)
            RecoveryHistory = table.Copy()
            RaiseHistoryChanged(False, Enumerable.Empty(Of String)())
        End Sub
        Friend ReadOnly Property RecoveryHistoryCount As Integer
            Get
                Return If(RecoveryHistory Is Nothing, 0, RecoveryHistory.Rows.Count)
            End Get
        End Property
        Friend ReadOnly Property ChangeInProgress As Boolean
            Get
                Return IsApplyingHistory OrElse writingAndCalculating OrElse ActiveGroup IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property ModelID As Integer
        'Only the isolated Structure Manager sets this; normal edit paths are unchanged.
        Friend Property IsReadOnlyPreview As Boolean
        Public Event HistoryChanged As EventHandler(Of ChangeHistoryChangedEventArgsV2)

        Public Sub New(ByRef setModelID As Integer)
            ModelID = setModelID
            WB = FileManager.ExcelModels(ModelID).WB
            MasterChangeLog.AddChangeLogEvent(New ChangeLogEvent With {
                .ModelID = ModelID, .Description = "File " & WB.Path & " opened",
                .WSName = "System Message", .TimeStamp = Now(),
                .UserName = Environment.UserName, .Status = 6,
                .Operation = "Notice"})
        End Sub

        Public ReadOnly Property CanUndo As Boolean
            Get
                Return UndoStack.Count > 0
            End Get
        End Property

        Public ReadOnly Property CanRedo As Boolean
            Get
                Return RedoStack.Count > 0
            End Get
        End Property

        Public Function BeginChangeGroup(ByVal description As String) As IDisposable
            If HasEngineEditingTrial Then Throw New InvalidOperationException("Use one engine batch for a grouped change.")
            If IsApplyingHistory Then Return New EmptyScopeV2()
            If ActiveGroup Is Nothing Then
                ActiveGroup = CreateGroup(description)
                ActiveGroupDepth = 0
            End If
            ActiveGroupDepth += 1
            Return New ChangeGroupScopeV2(Me)
        End Function

        Private Sub EndChangeGroup()
            If ActiveGroup Is Nothing Then Return
            ActiveGroupDepth -= 1
            If ActiveGroupDepth > 0 Then Return
            Dim completed As ChangeHistoryGroupV2 = ActiveGroup
            ActiveGroup = Nothing
            ActiveGroupDepth = 0
            If completed.Entries.Count > 0 Then CommitNewGroup(completed)
            If completed.Entries.Count > 0 Then
                RaiseHistoryChanged(False, completed.Entries.Select(Function(item) item.WorksheetName))
            End If
        End Sub

        Public Function ProcessChange(ByVal sentEvent As DataChangeEvent) As AbovoAppCls.AbovoTransaction
            If HasEngineEditingTrial Then Return ProcessEngineEditorChange(sentEvent)
            If IsApplyingHistory Then Return SuccessfulNoAction("A refresh-time post was ignored while history was being applied.")
            Dim worksheetName As String = NormalizeIdentifier(sentEvent.WSName)
            Dim address As String = NormalizeIdentifier(sentEvent.CellAddress)
            Try
                Return ProcessResolvedChange(WB.Worksheets(worksheetName).Cells(address), sentEvent)
            Catch ex As Exception
                Return FailedChange(sentEvent, worksheetName, address, ex)
            End Try
        End Function

        'A paste retains a snapshot per distinct target, writes typed values,
        'calculates once, and commits one undoable change-history group.
        Public Function ProcessChanges(ByVal changes As IEnumerable(Of DataChangeEvent),
                                       ByVal description As String) As AbovoAppCls.AbovoTransaction
            Return ProcessValidatedChanges(changes, description, Nothing)
        End Function

        'Optional admission check runs inside the existing rollback boundary,
        'after preceding defining-cell writes and before this cell is touched.
        Public Function ProcessValidatedChanges(ByVal changes As IEnumerable(Of DataChangeEvent),
                                                ByVal description As String,
                                                ByVal validate As Action(Of Cell, DataChangeEvent)) As AbovoAppCls.AbovoTransaction
            Return ProcessAdmittedChanges(changes, description, validate, Nothing)
        End Function

        'An explicit bulk command may omit unavailable targets. Omitted cells
        'are never written, snapshotted or recorded as edits. Other exceptions
        'still roll back the whole admitted batch, including its defining dates.
        Public Function ProcessAdmittedChanges(ByVal changes As IEnumerable(Of DataChangeEvent),
                                               ByVal description As String,
                                               ByVal validate As Action(Of Cell, DataChangeEvent),
                                               ByVal admit As Func(Of Cell, DataChangeEvent, Boolean)) As AbovoAppCls.AbovoTransaction
            If HasEngineEditingTrial Then Return NoAction("This model requires the engine batch edit path.")
            If IsReadOnlyPreview Then Return NoAction("Structure Manager previews are read-only.")
            If ChangeInProgress Then
                Return NoAction("Paste is unavailable during another change.")
            End If

            Dim ordered As List(Of DataChangeEvent) =
                If(changes, Enumerable.Empty(Of DataChangeEvent)()).ToList()
            If ordered.Count = 0 Then Return NoAction("There are no editable cells to paste.")

            Dim targets As New List(Of BatchChangeTargetV2)()
            Dim targetIndex As New Dictionary(Of String, BatchChangeTargetV2)(
                StringComparer.OrdinalIgnoreCase)
            Dim entries As New List(Of ChangeHistoryEntryV2)()
            Dim timer As Abovo.SummitDiagnostics.DiagnosticTimer =
                Abovo.SummitDiagnostics.DiagnosticTimer.StartNew()
            Dim writeMs As Long = 0
            Dim calculationMs As Long = 0
            Dim outcome As String = "failed"
            Dim activeChange As DataChangeEvent = ordered(0)

            writingAndCalculating = True
            Try
                For Each change As DataChangeEvent In ordered
                    activeChange = change
                    Dim worksheetName As String = NormalizeIdentifier(change.WSName)
                    Dim address As String = NormalizeIdentifier(change.CellAddress)
                    Dim cell As Cell = WB.Worksheets(worksheetName).Cells(address)
                    If admit IsNot Nothing AndAlso Not admit(cell, change) Then Continue For
                    If validate IsNot Nothing Then validate(cell, change)
                    Dim key As String =
                        cell.Worksheet.Name & "!" & cell.GetReferenceA1()
                    Dim target As BatchChangeTargetV2 = Nothing
                    If Not targetIndex.TryGetValue(key, target) Then
                        target = New BatchChangeTargetV2 With {
                            .Cell = cell,
                            .BeforeSnapshot = CellSnapshotV2.Capture(cell)}
                        targetIndex.Add(key, target)
                        targets.Add(target)
                    End If
                    target.LastChange = change
                    If cell.HasFormula Then FileManager.ExcelModels(ModelID).RequireFullRebuild()
                    WriteTypedValue(cell, change.ChangedValue, change.DataFormat)
                Next
                writeMs = timer.ElapsedMilliseconds

                If targets.Any(Function(target) Not target.BeforeSnapshot.Matches(target.Cell)) Then
                    FileManager.ExcelModels(ModelID).WBCalcEngine.CalculateWSs(
                        True, "Paste model=" & ModelID.ToString() &
                        ", cells=" & targets.Count.ToString())
                    calculationMs = timer.ElapsedMilliseconds - writeMs
                End If

                For Each target As BatchChangeTargetV2 In targets
                    If target.BeforeSnapshot.Matches(target.Cell) Then Continue For
                    Dim after As CellSnapshotV2 = CellSnapshotV2.Capture(target.Cell)
                    Dim change As DataChangeEvent = target.LastChange
                    entries.Add(New ChangeHistoryEntryV2 With {
                        .TimeStamp = If(change.TimeStamp = DateTime.MinValue,
                                        Now(), change.TimeStamp),
                        .Description = If(String.IsNullOrWhiteSpace(change.Description),
                                          description, change.Description),
                        .WorksheetName = target.Cell.Worksheet.Name,
                        .CellAddress = target.Cell.GetReferenceA1(),
                        .BeforeSnapshot = target.BeforeSnapshot,
                        .AfterSnapshot = after,
                        .OriginalDisplay = target.BeforeSnapshot.DisplayText,
                        .ChangedDisplay = after.DisplayText,
                        .UserName = change.UserName,
                        .DataFormat = change.DataFormat})
                Next
                outcome = "ok"
            Catch ex As Exception
                Dim rollbackFailures As New List(Of String)()
                For Each target As BatchChangeTargetV2 In targets.AsEnumerable().Reverse()
                    Try
                        target.BeforeSnapshot.Apply(target.Cell)
                    Catch rollbackError As Exception
                        rollbackFailures.Add(
                            target.Cell.Worksheet.Name & "!" &
                            target.Cell.GetReferenceA1() & ": " &
                            rollbackError.Message)
                    End Try
                Next
                If targets.Count > 0 Then
                    Try
                        FileManager.ExcelModels(ModelID).WBCalcEngine.CalculateWSs(
                            True, "Paste rollback model=" & ModelID.ToString())
                    Catch rollbackError As Exception
                        rollbackFailures.Add(
                            "Recalculate restored workbook: " &
                            rollbackError.Message)
                    End Try
                    For Each target As BatchChangeTargetV2 In targets
                        If Not target.BeforeSnapshot.Matches(target.Cell) Then
                            rollbackFailures.Add(
                                "Could not verify " & target.Cell.Worksheet.Name &
                                "!" & target.Cell.GetReferenceA1())
                        End If
                    Next
                End If
                If rollbackFailures.Count > 0 Then
                    ModelSafetyManager.MarkRecoveryRequired(
                        ModelID, description,
                        String.Join(Environment.NewLine, rollbackFailures),
                        "Change Manager")
                End If
                outcome = If(rollbackFailures.Count = 0,
                             "rolled back", "recovery required")
                Return FailedChange(
                    activeChange, activeChange.WSName,
                    activeChange.CellAddress, ex)
            Finally
                writingAndCalculating = False
                Abovo.SummitDiagnostics.WriteLine(
                    "[Paste Benchmark] model=" & ModelID.ToString() &
                    ", requested=" & ordered.Count.ToString() &
                    ", targets=" & targets.Count.ToString() &
                    ", write=" & writeMs.ToString() & " ms" &
                    ", calculation=" & calculationMs.ToString() & " ms" &
                    ", total=" & timer.ElapsedMilliseconds.ToString() & " ms" &
                    ", outcome=" & outcome)
            End Try

            Dim result As New AbovoAppCls.AbovoTransaction(
                "ModelChangeManagerV2.ProcessChanges") With {
                    .BSuccess = True,
                    .IntegerReturn = entries.Count,
                    .StrResponseMessage = entries.Count.ToString() & " cell(s) pasted."}
            If entries.Count = 0 Then Return result

            Dim group As ChangeHistoryGroupV2 = CreateGroup(description)
            For Each entry As ChangeHistoryEntryV2 In entries
                entry.GroupID = group.GroupID
                group.Entries.Add(entry)
            Next
            FileManager.ExcelModels(ModelID).MarkUserChange()
            CommitNewGroup(group)
            For Each entry As ChangeHistoryEntryV2 In group.Entries
                Try
                    MasterChangeLog.AddChangeLogEvent(
                        ToLogEvent(entry, 1, "Apply"))
                Catch logError As Exception
                    Abovo.SummitDiagnostics.WriteLine(
                        "[Paste] Change-log entry failed: " &
                        logError.Message)
                End Try
            Next
            RaiseHistoryChanged(
                False, group.Entries.Select(Function(entry) entry.WorksheetName))
            Return result
        End Function

        Public Function ProcessChangeByNRAddressing(ByVal sentEvent As DataChangeEvent) As AbovoAppCls.AbovoTransaction
            If IsApplyingHistory Then Return SuccessfulNoAction("A refresh-time post was ignored while history was being applied.")
            Try
                Dim targetRange As CellRange = WB.Range(NormalizeIdentifier(sentEvent.TargetNR))
                Dim targetCell As Cell = If(sentEvent.NROrientation = Orientation.Horizontal,
                                            targetRange(0, sentEvent.TargetNRIndex),
                                            targetRange(sentEvent.TargetNRIndex, 0))
                sentEvent.WSName = targetRange.Worksheet.Name
                sentEvent.CellAddress = targetCell.GetReferenceA1()
                If HasEngineEditingTrial Then Return ProcessEngineEditorChange(sentEvent)
                Return ProcessResolvedChange(targetCell, sentEvent)
            Catch ex As Exception
                Return FailedChange(sentEvent, NormalizeIdentifier(sentEvent.TargetNR), String.Empty, ex)
            End Try
        End Function

        Private Function ProcessResolvedChange(ByVal targetCell As Cell,
                                               ByVal sentEvent As DataChangeEvent) As AbovoAppCls.AbovoTransaction
            If IsReadOnlyPreview Then Return NoAction("Structure Manager previews are read-only.")
            If writingAndCalculating Then Return SuccessfulNoAction("A refresh-time post was ignored while an edit was being applied.")
            Dim result As New AbovoAppCls.AbovoTransaction("ModelChangeManagerV2.ProcessChange")
            If targetCell Is Nothing Then Return FailedChange(sentEvent, sentEvent.WSName, sentEvent.CellAddress, New InvalidOperationException("The target cell was not found."))

            Dim benchmark As Abovo.SummitDiagnostics.DiagnosticTimer =
                Abovo.SummitDiagnostics.DiagnosticTimer.StartNew(CheckSheetWatch.Benchmark)
            Dim before As CellSnapshotV2 = CellSnapshotV2.Capture(targetCell)
            Dim automaticGroup As Boolean = ActiveGroup Is Nothing
            Dim group As ChangeHistoryGroupV2 = If(ActiveGroup, CreateGroup(sentEvent.Description))
            Dim setupMs As Long = benchmark.ElapsedMilliseconds
            Dim writeMs As Long = 0
            Dim calculationMs As Long = 0
            Dim outcome As String = "failed"
            writingAndCalculating = True
            Try
                If targetCell.HasFormula Then FileManager.ExcelModels(ModelID).RequireFullRebuild()
                WriteTypedValue(targetCell, sentEvent.ChangedValue, sentEvent.DataFormat)
                writeMs = benchmark.ElapsedMilliseconds - setupMs
                FileManager.ExcelModels(ModelID).WBCalcEngine.CalculateWSs(
                    True,
                    "Edit model=" & ModelID.ToString() &
                    ", worksheet=" & targetCell.Worksheet.Name)
                calculationMs = benchmark.ElapsedMilliseconds - setupMs - writeMs

                'Calculation is allowed to normalise a posted value.  History must
                'therefore describe the authoritative post-calculation cell, not
                'the transient value written immediately before calculation.
                Dim after As CellSnapshotV2 = CellSnapshotV2.Capture(targetCell)
                If before.Matches(targetCell) Then
                    result.BSuccess = True
                    result.StrResponseMessage = "The workbook value is unchanged."
                    outcome = "unchanged"
                    Return result
                End If
                Dim entry As New ChangeHistoryEntryV2 With {
                    .GroupID = group.GroupID, .TimeStamp = If(sentEvent.TimeStamp = DateTime.MinValue, Now(), sentEvent.TimeStamp),
                    .Description = sentEvent.Description, .WorksheetName = targetCell.Worksheet.Name,
                    .CellAddress = targetCell.GetReferenceA1(), .BeforeSnapshot = before,
                    .AfterSnapshot = after, .OriginalDisplay = before.DisplayText,
                    .ChangedDisplay = after.DisplayText, .UserName = sentEvent.UserName,
                    .DataFormat = sentEvent.DataFormat}
                group.Entries.Add(entry)
                MasterChangeLog.AddChangeLogEvent(ToLogEvent(entry, 1, "Apply"))
                If automaticGroup Then CommitNewGroup(group)
                FileManager.ExcelModels(ModelID).MarkUserChange()
                result.BSuccess = True
                result.StrResponseMessage = "Change applied."
                writingAndCalculating = False
                If automaticGroup Then RaiseHistoryChanged(False, {targetCell.Worksheet.Name})
                outcome = "ok"
                Return result
            Catch ex As Exception
                Dim rollbackFailures As New List(Of String)
                Try
                    before.Apply(targetCell)
                Catch rollbackError As Exception
                    rollbackFailures.Add("Restore cell: " & rollbackError.Message)
                End Try
                Try
                    FileManager.ExcelModels(ModelID).WBCalcEngine.CalculateWSs(
                        True,
                        "Edit rollback model=" & ModelID.ToString() &
                        ", worksheet=" & targetCell.Worksheet.Name)
                Catch rollbackError As Exception
                    rollbackFailures.Add("Recalculate restored workbook: " & rollbackError.Message)
                End Try
                If rollbackFailures.Count > 0 Then
                    ModelSafetyManager.MarkRecoveryRequired(
                        ModelID,
                        sentEvent.Description,
                        String.Join(Environment.NewLine, rollbackFailures),
                        "Change Manager",
                        targetCell.Worksheet.Name & "!" & targetCell.GetReferenceA1())
                End If
                outcome = If(rollbackFailures.Count = 0, "rolled back", "recovery required")
                Return FailedChange(sentEvent, targetCell.Worksheet.Name, targetCell.GetReferenceA1(), ex)
            Finally
                writingAndCalculating = False
                Abovo.SummitDiagnostics.WriteTrialLine("[Edit Trial Benchmark] model=" & ModelID.ToString() &
                    ", worksheet=" & targetCell.Worksheet.Name & ", cell=" & targetCell.GetReferenceA1() &
                    ", total=" & benchmark.ElapsedMilliseconds.ToString() & " ms, outcome=" & outcome)
                Abovo.SummitDiagnostics.WriteLine(
                    "[Population Benchmark] Edit: model=" & ModelID.ToString() &
                    ", worksheet=" & targetCell.Worksheet.Name &
                    ", setup=" & setupMs.ToString() & " ms" &
                    ", write=" & writeMs.ToString() & " ms" &
                    ", calculation=" & calculationMs.ToString() & " ms" &
                    ", journalAndRefresh=" &
                    (benchmark.ElapsedMilliseconds - setupMs - writeMs - calculationMs).ToString() & " ms" &
                    ", total=" & benchmark.ElapsedMilliseconds.ToString() & " ms" &
                    ", outcome=" & outcome)
            End Try
        End Function

        Private Sub CommitNewGroup(ByVal group As ChangeHistoryGroupV2)
            For Each discarded As ChangeHistoryGroupV2 In RedoStack
                discarded.State = ChangeHistoryStateV2.Superseded
            Next
            RedoStack.Clear()
            group.State = ChangeHistoryStateV2.Applied
            Journal.Add(group)
            UndoStack.Add(group)
        End Sub

        Public Function Undo() As AbovoAppCls.AbovoTransaction
            If Not CanUndo Then Return NoAction("There is no change to undo.")
            Return ApplyHistoryGroup(UndoStack(UndoStack.Count - 1), False)
        End Function

        Public Function Redo() As AbovoAppCls.AbovoTransaction
            If Not CanRedo Then Return NoAction("There is no change to redo.")
            Return ApplyHistoryGroup(RedoStack(RedoStack.Count - 1), True)
        End Function

        Public Function UndoTo(ByVal groupID As Integer) As AbovoAppCls.AbovoTransaction
            Dim index As Integer = UndoStack.FindLastIndex(Function(item) item.GroupID = groupID)
            If index < 0 Then Return NoAction("The selected history item is not currently undoable.")
            Dim result As AbovoAppCls.AbovoTransaction = Nothing
            While UndoStack.Count > index
                result = Undo()
                If result.BError Then Return result
            End While
            Return result
        End Function

        Public Function RedoTo(ByVal groupID As Integer) As AbovoAppCls.AbovoTransaction
            Dim index As Integer = RedoStack.FindLastIndex(Function(item) item.GroupID = groupID)
            If index < 0 Then Return NoAction("The selected history item is not currently redoable.")
            Dim result As AbovoAppCls.AbovoTransaction = Nothing
            While RedoStack.Count > index
                result = Redo()
                If result.BError Then Return result
            End While
            Return result
        End Function

        Private Function ApplyHistoryGroup(ByVal group As ChangeHistoryGroupV2,
                                           ByVal redo As Boolean) As AbovoAppCls.AbovoTransaction
            If HasEngineEditingTrial Then Return ApplyEngineHistory(group, redo)
            Dim result As New AbovoAppCls.AbovoTransaction(If(redo, "Redo", "Undo"))
            Dim ordered As List(Of ChangeHistoryEntryV2) = If(redo, group.Entries.ToList(), group.Entries.AsEnumerable().Reverse().ToList())

            For Each entry As ChangeHistoryEntryV2 In ordered
                Dim cell As Cell = WB.Worksheets(entry.WorksheetName).Cells(entry.CellAddress)
                Dim expected As CellSnapshotV2 = If(redo, entry.BeforeSnapshot, entry.AfterSnapshot)
                If Not expected.Matches(cell) Then
                    result.BError = True
                    result.StrResponseMessage = "Cannot " & If(redo, "redo", "undo") & " because " & entry.WorksheetName & "!" & entry.CellAddress & " has changed since this action."
                    Return result
                End If
            Next

            IsApplyingHistory = True
            Dim applied As New List(Of ChangeHistoryEntryV2)()
            Try
                For Each entry As ChangeHistoryEntryV2 In ordered
                    Dim cell As Cell = WB.Worksheets(entry.WorksheetName).Cells(entry.CellAddress)
                    Dim targetSnapshot As CellSnapshotV2 = If(redo, entry.AfterSnapshot, entry.BeforeSnapshot)
                    If cell.HasFormula OrElse targetSnapshot.HasFormula Then FileManager.ExcelModels(ModelID).RequireFullRebuild()
                    targetSnapshot.Apply(cell)
                    applied.Add(entry)
                Next
                CalculateHistoryWorksheets(ordered)

                'Calculation and control refresh are capable of raising editor
                'events. Do not report success unless the workbook still contains
                'every snapshot that this undo/redo intended to apply.
                For Each entry As ChangeHistoryEntryV2 In ordered
                    Dim cell As Cell = WB.Worksheets(entry.WorksheetName).Cells(entry.CellAddress)
                    Dim targetSnapshot As CellSnapshotV2 = If(redo, entry.AfterSnapshot, entry.BeforeSnapshot)
                    If Not targetSnapshot.Matches(cell) Then
                        Throw New InvalidOperationException(
                            entry.WorksheetName & "!" & entry.CellAddress &
                            " was changed again while history was being applied.")
                    End If
                Next
            Catch ex As Exception
                Dim rollbackFailures As New List(Of String)
                For Each entry As ChangeHistoryEntryV2 In applied.AsEnumerable().Reverse()
                    Try
                        Dim cell As Cell = WB.Worksheets(entry.WorksheetName).Cells(entry.CellAddress)
                        Dim rollbackSnapshot As CellSnapshotV2 = If(redo, entry.BeforeSnapshot, entry.AfterSnapshot)
                        rollbackSnapshot.Apply(cell)
                    Catch rollbackError As Exception
                        rollbackFailures.Add(
                            entry.WorksheetName & "!" & entry.CellAddress & ": " &
                            rollbackError.Message)
                    End Try
                Next
                Try
                    CalculateHistoryWorksheets(applied)
                Catch rollbackError As Exception
                    rollbackFailures.Add("Recalculate restored workbook: " & rollbackError.Message)
                End Try
                result.BError = True
                If rollbackFailures.Count = 0 Then
                    result.StrResponseMessage =
                        "The " & If(redo, "redo", "undo") &
                        " failed; the previous workbook values were restored: " & ex.Message
                Else
                    result.StrResponseMessage =
                        "The " & If(redo, "redo", "undo") &
                        " failed and rollback could not be verified: " & ex.Message &
                        Environment.NewLine & String.Join(Environment.NewLine, rollbackFailures)
                    ModelSafetyManager.MarkRecoveryRequired(
                        ModelID,
                        If(redo, "Redo", "Undo"),
                        result.StrResponseMessage,
                        "Change Manager")
                End If
                Return result
            Finally
                IsApplyingHistory = False
            End Try

            If redo Then
                RedoStack.Remove(group)
                UndoStack.Add(group)
                group.State = ChangeHistoryStateV2.Applied
            Else
                UndoStack.Remove(group)
                RedoStack.Add(group)
                group.State = ChangeHistoryStateV2.Undone
            End If
            For Each entry As ChangeHistoryEntryV2 In ordered
                MasterChangeLog.AddChangeLogEvent(ToLogEvent(entry, If(redo, 5, 4), If(redo, "Redo", "Undo")))
            Next
            FileManager.ExcelModels(ModelID).MarkUserChange()
            result.BSuccess = True
            result.StrResponseMessage = If(redo, "Change redone.", "Change undone.")
            RaiseHistoryChanged(True, group.Entries.Select(Function(item) item.WorksheetName))
            Return result
        End Function

        Private Sub CalculateHistoryWorksheets(ByVal entries As IEnumerable(Of ChangeHistoryEntryV2))
            Dim engine As CalcEngine = FileManager.ExcelModels(ModelID).WBCalcEngine

            'History can be used after its DIT was deactivated, or from a different
            'interface. CalculateWSs only knows currently registered worksheets and
            'can otherwise do no calculation at all. Bring restored source sheets
            'current before the normal active-sheet calculation and UI refresh.
            'More than one active object already takes the workbook calculation
            'path; registered sheets are likewise calculated by CalculateWSs.
            If engine.ActiveObjectCount <= 1 Then
                Dim registered As New HashSet(Of String)(
                    engine.ActiveWSs.Where(Function(sheet) sheet IsNot Nothing).
                        Select(Function(sheet) sheet.Name),
                    StringComparer.OrdinalIgnoreCase)
                For Each worksheetName As String In entries.
                    Select(Function(entry) entry.WorksheetName).
                    Distinct(StringComparer.OrdinalIgnoreCase)
                    If Not registered.Contains(worksheetName) Then
                        WB.Worksheets(worksheetName).Calculate()
                    End If
                Next
            End If

            engine.CalculateWSs()
        End Sub

        Public Function GetHistoryTable() As DataTable
            Dim table As New DataTable("ModelHistoryV2")
            table.Columns.Add("GroupID", GetType(Integer))
            table.Columns.Add("TimeStamp", GetType(DateTime))
            table.Columns.Add("Description", GetType(String))
            table.Columns.Add("Worksheet", GetType(String))
            table.Columns.Add("Cell", GetType(String))
            table.Columns.Add("OriginalValue", GetType(String))
            table.Columns.Add("NewValue", GetType(String))
            table.Columns.Add("User", GetType(String))
            table.Columns.Add("State", GetType(String))
            table.Columns.Add("DataType", GetType(String))
            table.Columns.Add("GroupSize", GetType(Integer))
            table.Columns.Add("Action", GetType(String))
            If RecoveryHistory IsNot Nothing Then
                For Each row As DataRow In RecoveryHistory.Rows
                    table.ImportRow(row)
                Next
            End If
            For Each group As ChangeHistoryGroupV2 In Journal
                For Each entry As ChangeHistoryEntryV2 In group.Entries
                    table.Rows.Add(group.GroupID, entry.TimeStamp, entry.Description,
                                   entry.WorksheetName, entry.CellAddress,
                                   entry.OriginalDisplay, entry.ChangedDisplay,
                                   entry.UserName, group.State.ToString(),
                                   entry.DataFormat, group.Entries.Count,
                                   If(group.State = ChangeHistoryStateV2.Applied, "Undo",
                                      If(group.State = ChangeHistoryStateV2.Undone, "Redo", String.Empty)))
                Next
            Next
            Return table
        End Function

        Private Function CreateGroup(ByVal description As String) As ChangeHistoryGroupV2
            NextGroupID += 1
            Return New ChangeHistoryGroupV2 With {.GroupID = NextGroupID, .Description = description, .TimeStamp = Now()}
        End Function

        Private Function FailedChange(ByVal sentEvent As DataChangeEvent,
                                      ByVal worksheetName As String,
                                      ByVal address As String,
                                      ByVal ex As Exception) As AbovoAppCls.AbovoTransaction
            MasterChangeLog.AddChangeLogEvent(New ChangeLogEvent With {
                .ModelID = ModelID,
                .Description = sentEvent.Description & " failed: " & ex.Message,
                .WSName = worksheetName, .CellAddress = address,
                .OriginalValue = Convert.ToString(sentEvent.OriginalValue, CultureInfo.CurrentCulture),
                .ChangedValue = Convert.ToString(sentEvent.ChangedValue, CultureInfo.CurrentCulture),
                .TimeStamp = Now(), .UserName = sentEvent.UserName, .Status = 3,
                .DataType = sentEvent.DataFormat, .Operation = "Error"})
            Return New AbovoAppCls.AbovoTransaction With {
                .BError = True, .BSuccess = False,
                .StrResponseMessage = "Error processing change for " & worksheetName & "!" & address & ": " & ex.Message}
        End Function

        Private Shared Function NoAction(ByVal message As String) As AbovoAppCls.AbovoTransaction
            Return New AbovoAppCls.AbovoTransaction With {.BSuccess = False, .BError = False, .StrResponseMessage = message}
        End Function

        Private Shared Function SuccessfulNoAction(ByVal message As String) As AbovoAppCls.AbovoTransaction
            Return New AbovoAppCls.AbovoTransaction With {.BSuccess = True, .BError = False, .StrResponseMessage = message}
        End Function

        Private Function ToLogEvent(ByVal entry As ChangeHistoryEntryV2,
                                    ByVal status As Integer,
                                    ByVal operation As String) As ChangeLogEvent
            Return New ChangeLogEvent With {
                .ModelID = ModelID, .Description = entry.Description,
                .WSName = entry.WorksheetName, .CellAddress = entry.CellAddress,
                .OriginalValue = entry.OriginalDisplay, .ChangedValue = entry.ChangedDisplay,
                .TimeStamp = Now(), .UserName = entry.UserName, .Status = status,
                .DataType = entry.DataFormat, .GroupID = entry.GroupID,
                .Operation = operation}
        End Function

        Private Sub RaiseHistoryChanged(ByVal undoRedo As Boolean,
                                        ByVal worksheets As IEnumerable(Of String))
            Dim args As New ChangeHistoryChangedEventArgsV2(
                undoRedo,
                worksheets.Where(Function(name) Not String.IsNullOrWhiteSpace(name)).Distinct(StringComparer.OrdinalIgnoreCase))

            'History notification is presentation work, not part of the workbook
            'transaction. A stale/disposed UI subscriber must not escape into
            'ProcessResolvedChange after its group has already been committed and
            'cause the workbook cell to be rolled back while history stays Applied.
            Dim subscribers As EventHandler(Of ChangeHistoryChangedEventArgsV2) = HistoryChangedEvent
            If subscribers IsNot Nothing Then
                For Each subscriber As [Delegate] In subscribers.GetInvocationList()
                    Try
                        DirectCast(subscriber, EventHandler(Of ChangeHistoryChangedEventArgsV2)).Invoke(Me, args)
                    Catch ex As Exception
                        Abovo.SummitDiagnostics.WriteLine(
                            "HistoryChanged subscriber failed: " & ex.ToString())
                    End Try
                Next
            End If

            'Read only after successful history completion AND its presentation
            'callbacks. Those callbacks can populate bound editors and invalidate
            'calculation again. Coalescing onto the next UI message captures the
            'final revision without certifying provisional or rolled-back state.
            If args.WorksheetNames.Count > 0 AndAlso Not HasEngineEditingTrial Then
                Try
                    CheckSheetWatch.RefreshVisibleAfterCommittedChange(FileManager.ExcelModels(ModelID))
                Catch ex As Exception
                    Abovo.SummitDiagnostics.WriteLine("Check Sheet completion refresh failed: " & ex.ToString())
                End Try
            End If
        End Sub

        Private Shared Function NormalizeIdentifier(ByVal value As String) As String
            Return If(value Is Nothing, Nothing, value.Trim())
        End Function

        Private Shared Sub WriteTypedValue(ByVal targetCell As Cell,
                                           ByVal changedValue As Object,
                                           ByVal dataFormat As String)
            Dim value = ResolveTypedValue(targetCell, changedValue, dataFormat)
            If value Is Nothing Then
                targetCell.ClearContents()
            Else
                targetCell.Value = CellValue.FromObject(value)
            End If
        End Sub

        Private Shared Function ResolveTypedValue(ByVal targetCell As Cell,
                                                  ByVal changedValue As Object,
                                                  ByVal dataFormat As String) As Object
            Dim fundingDate = FundingPaymentDateSupport.ValidateChange(targetCell, changedValue)
            If fundingDate.HasValue Then
                'Also protect workbooks using an older embedded XML text definition.
                Return fundingDate.Value
            End If
            If changedValue Is Nothing OrElse Convert.IsDBNull(changedValue) Then
                Return Nothing
            End If
            'A workbook Yes/No validation is text even when an older interface
            'declares the editor Boolean. Never convert its Yes to the number 1.
            If WorkbookIntegritySupport.IsYesNoInput(targetCell) Then
                Dim choice = Convert.ToString(changedValue, CultureInfo.CurrentCulture).Trim()
                If TypeOf changedValue Is Boolean Then choice = If(CBool(changedValue), "Yes", "No")
                If choice.Length = 0 Then
                    Return Nothing
                ElseIf choice.Equals("Yes", StringComparison.OrdinalIgnoreCase) OrElse choice.Equals("No", StringComparison.OrdinalIgnoreCase) Then
                    Return If(choice.Equals("Yes", StringComparison.OrdinalIgnoreCase), "Yes", "No")
                Else
                    Throw New ArgumentException("Select Yes or No for this workbook input.")
                End If
            End If
            Select Case If(dataFormat, String.Empty).Trim().ToUpperInvariant()
                Case "S", "FL", "DUMMY", String.Empty
                    Return Convert.ToString(changedValue, CultureInfo.CurrentCulture)
                Case "B"
                    Return If(ConvertToBoolean(changedValue), 1, 0)
                Case "BOOL", "BOOLEAN"
                    'Excel form-control linked cells are genuine Boolean values.
                    'Keep the legacy B representation numeric for existing DIT data,
                    'but preserve Boolean semantics where workbook formulas test the
                    'linked cell directly.
                    Return ConvertToBoolean(changedValue)
                Case "I", "Y"
                    Return ConvertToInteger(changedValue)
                Case "D", "DM"
                    Return ConvertToDateOrSerial(changedValue)
                Case "N", "P", "C", "M", "SM", "R"
                    Return ConvertToDouble(changedValue)
                Case Else
                    Return changedValue
            End Select
        End Function

        Private Shared Function ConvertToBoolean(ByVal value As Object) As Boolean
            If TypeOf value Is Boolean Then Return DirectCast(value, Boolean)
            Dim text As String = Convert.ToString(value, CultureInfo.CurrentCulture).Trim()
            Dim parsed As Boolean
            If Boolean.TryParse(text, parsed) Then Return parsed
            If text.Equals("YES", StringComparison.OrdinalIgnoreCase) OrElse text.Equals("Y", StringComparison.OrdinalIgnoreCase) Then Return True
            If text.Equals("NO", StringComparison.OrdinalIgnoreCase) OrElse text.Equals("N", StringComparison.OrdinalIgnoreCase) Then Return False
            Return ConvertToDouble(value) <> 0
        End Function

        Private Shared Function ConvertToInteger(ByVal value As Object) As Integer
            If Not TypeOf value Is String AndAlso TypeOf value Is IConvertible Then
                Return Convert.ToInt32(value, CultureInfo.InvariantCulture)
            End If
            Dim parsed As Integer
            Dim text As String = Convert.ToString(value, CultureInfo.CurrentCulture).Trim()
            If Integer.TryParse(text, NumberStyles.Integer Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, parsed) Then Return parsed
            If Integer.TryParse(text, NumberStyles.Integer Or NumberStyles.AllowThousands, CultureInfo.InvariantCulture, parsed) Then Return parsed
            Throw New FormatException("'" & text & "' is not a valid whole number.")
        End Function

        Private Shared Function ConvertToDouble(ByVal value As Object) As Double
            If Not TypeOf value Is String AndAlso TypeOf value Is IConvertible Then
                Return Convert.ToDouble(value, CultureInfo.InvariantCulture)
            End If
            Dim parsed As Double
            Dim text As String = Convert.ToString(value, CultureInfo.CurrentCulture).Trim()
            If Double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, parsed) Then Return parsed
            If Double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, parsed) Then Return parsed
            Throw New FormatException("'" & text & "' is not a valid number.")
        End Function

        Private Shared Function ConvertToDateOrSerial(ByVal value As Object) As Object
            If TypeOf value Is DateTime Then Return DirectCast(value, DateTime)
            If Not TypeOf value Is String AndAlso TypeOf value Is IConvertible Then
                Return Convert.ToDouble(value, CultureInfo.InvariantCulture)
            End If
            Dim text As String = Convert.ToString(value, CultureInfo.CurrentCulture).Trim()
            Dim serial As Double
            If Double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, serial) Then Return serial
            If Double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, serial) Then Return serial
            Dim parsedDate As DateTime
            If DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, parsedDate) Then Return parsedDate
            If DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, parsedDate) Then Return parsedDate
            Throw New FormatException("'" & text & "' is not a valid date.")
        End Function

        Private NotInheritable Class ChangeGroupScopeV2
            Implements IDisposable
            Private Owner As ModelChangeManagerV2
            Public Sub New(ByVal manager As ModelChangeManagerV2)
                Owner = manager
            End Sub
            Public Sub Dispose() Implements IDisposable.Dispose
                If Owner Is Nothing Then Return
                Owner.EndChangeGroup()
                Owner = Nothing
            End Sub
        End Class

        Private NotInheritable Class EmptyScopeV2
            Implements IDisposable
            Public Sub Dispose() Implements IDisposable.Dispose
            End Sub
        End Class
    End Class

    Friend NotInheritable Class ChangeHistoryGroupV2
        Public GroupID As Integer
        Public Description As String
        Public TimeStamp As DateTime
        Public State As ChangeHistoryStateV2
        Public ReadOnly Entries As New List(Of ChangeHistoryEntryV2)()
    End Class

    Friend NotInheritable Class BatchChangeTargetV2
        Public Cell As Cell
        Public BeforeSnapshot As CellSnapshotV2
        Public LastChange As DataChangeEvent
    End Class

    Friend NotInheritable Class ChangeHistoryEntryV2
        Public EngineBefore As WorkbookEngines.WorkbookCellState
        Public EngineAfter As WorkbookEngines.WorkbookCellState
        Public EnginePermission As WorkbookEngines.WorkbookValuePermission
        Public EngineCalculateBefore As Boolean
        Public GroupID As Integer
        Public TimeStamp As DateTime
        Public Description As String
        Public WorksheetName As String
        Public CellAddress As String
        Public BeforeSnapshot As CellSnapshotV2
        Public AfterSnapshot As CellSnapshotV2
        Public OriginalDisplay As String
        Public ChangedDisplay As String
        Public UserName As String
        Public DataFormat As String
    End Class

    Friend NotInheritable Class CellSnapshotV2
        Public Value As CellValue
        Public HasFormula As Boolean
        Public FormulaInvariant As String
        Public DisplayText As String

        Public Shared Function Capture(ByVal cell As Cell) As CellSnapshotV2
            Return New CellSnapshotV2 With {
                .Value = cell.Value, .HasFormula = cell.HasFormula,
                .FormulaInvariant = If(cell.HasFormula, cell.FormulaInvariant, String.Empty),
                .DisplayText = cell.DisplayText}
        End Function

        Public Sub Apply(ByVal cell As Cell)
            If HasFormula Then
                cell.FormulaInvariant = FormulaInvariant
            ElseIf Value Is Nothing OrElse Value.IsEmpty Then
                cell.ClearContents()
            Else
                cell.Value = Value
            End If
        End Sub

        Public Function Matches(ByVal cell As Cell) As Boolean
            If cell Is Nothing OrElse cell.HasFormula <> HasFormula Then Return False
            If HasFormula Then Return String.Equals(cell.FormulaInvariant, FormulaInvariant, StringComparison.Ordinal)
            Return SnapshotKey(cell.Value) = SnapshotKey(Value)
        End Function

        Private Shared Function SnapshotKey(ByVal value As CellValue) As String
            If value Is Nothing OrElse value.IsEmpty Then Return "E:"
            If value.IsBoolean Then Return "B:" & value.BooleanValue.ToString()
            If value.IsNumeric Then Return "N:" & value.NumericValue.ToString("R", CultureInfo.InvariantCulture)
            If value.IsText Then Return "S:" & value.TextValue
            If value.IsError Then Return "X:" & value.ErrorValue.ToString()
            Return value.Type.ToString() & ":" & value.ToString()
        End Function
    End Class

End Namespace
