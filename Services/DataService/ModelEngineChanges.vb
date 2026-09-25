Option Strict On

Imports System.Globalization
Imports System.Threading
Imports Abovo.WorkbookEngines
Imports DevExpress.Spreadsheet

Namespace Abovo
    ' The UI must supply an accepted, revision-bound input snapshot and its
    ' existing XML/fill permission rule. No permission is inferred from colour.
    Public NotInheritable Class ModelEngineInput
        Public ReadOnly Property Change As DataChangeEvent
        Public ReadOnly Property Expected As WorkbookCellSnapshot
        Public ReadOnly Property Permission As WorkbookValuePermission
        Public ReadOnly Property CalculateBefore As Boolean
        Public ReadOnly Property SkipUnavailable As Boolean
        Public Sub New(change As DataChangeEvent, expected As WorkbookCellSnapshot, permission As WorkbookValuePermission,
                       Optional calculateBefore As Boolean = False, Optional skipUnavailable As Boolean = False)
            If expected Is Nothing Then Throw New ArgumentNullException(NameOf(expected))
            If Not [Enum].IsDefined(GetType(WorkbookValuePermission), permission) Then Throw New ArgumentOutOfRangeException(NameOf(permission))
            Me.Change = change : Me.Expected = expected : Me.Permission = permission
            Me.CalculateBefore = calculateBefore : Me.SkipUnavailable = skipUnavailable
        End Sub
    End Class

    Partial Public NotInheritable Class ModelChangeManagerV2
        Private engineTrial As WorkbookCalculationSession
        Private engineTrialAreas As List(Of WorkbookReadArea)
        Private engineTrialResult As WorkbookCalculationResult
        Private engineTrialPresentation As Boolean
        Public ReadOnly Property HasEngineEditingTrial As Boolean
            Get
                Return engineTrial IsNot Nothing
            End Get
        End Property
        Public ReadOnly Property EngineEditingResult As WorkbookCalculationResult
            Get
                Return If(engineTrial IsNot Nothing AndAlso engineTrial.IsCurrent(engineTrialResult), engineTrialResult, Nothing)
            End Get
        End Property

        ' Not wired into model opening until every display, save and structural
        ' path can honour this owner. The normal UI workbook stays unmodified;
        ' it is not a second authoritative copy of engine-owned edits.
        Friend Sub BindEngineEditingTrial(session As WorkbookCalculationSession, initial As WorkbookCalculationResult)
            Dim model = RequireSaveHistoryOwner()
            If session Is Nothing OrElse Not session.IsCurrent(initial) Then Throw New InvalidOperationException("Current calculated engine results are required.")
            If engineTrial IsNot Nothing OrElse Journal.Count <> 0 OrElse model.IsDirty Then Throw New InvalidOperationException("Bind only a clean, newly opened model.")
            Dim snapshot = CaptureSaveHistory(initial)
            session.RequireEditingSource(snapshot.SourcePath)
            If initial.Blocks.Count = 0 Then Throw New InvalidOperationException("A result display is required.")
            engineTrialAreas = initial.Blocks.Select(Function(block) block.Area).ToList()
            engineTrialPresentation = initial.Presentation.Count > 0
            engineTrialResult = initial : engineTrial = session
            ModelEngineView.Bind(WB, session, initial)
        End Sub

        Public Function ProcessEngineChanges(inputs As IEnumerable(Of ModelEngineInput), description As String,
                                             Optional cancellation As CancellationToken = Nothing) As AbovoAppCls.AbovoTransaction
            RequireSaveHistoryOwner()
            If engineTrial Is Nothing Then Throw New InvalidOperationException("No engine editing owner is bound.")
            If IsReadOnlyPreview Then Return NoAction("Structure Manager previews are read-only.")
            Dim source = If(inputs, Enumerable.Empty(Of ModelEngineInput)()).Take(10001).ToList()
            If source.Count = 0 OrElse source.Count > 10000 Then Throw New ArgumentException("An edit requires 1 to 10,000 inputs.")
            Dim edits As New List(Of WorkbookValueChange)()
            For Each item In source
                If item Is Nothing Then Throw New ArgumentException("Null input.")
                Dim area = item.Expected.Area
                Dim change = item.Change
                If change.ModelID <> ModelID OrElse Not String.Equals(change.WSName, area.Worksheet, StringComparison.OrdinalIgnoreCase) OrElse
                    Not String.Equals(change.CellAddress, area.Address.Split(":"c)(0), StringComparison.OrdinalIgnoreCase) Then Throw New ArgumentException("The input does not match its model and captured cell.")
                Dim cell = WB.Worksheets(area.Worksheet).Cells(area.Row, area.Column)
                Dim value = ResolveTypedValue(cell, change.ChangedValue, change.DataFormat)
                If TypeOf value Is DateTime Then value = CellValue.FromDateTime(CDate(value), WB.DocumentSettings.Calculation.Use1904DateSystem).NumericValue
                If value IsNot Nothing AndAlso Not TypeOf value Is String AndAlso Not TypeOf value Is Boolean Then value = Convert.ToDouble(value, CultureInfo.InvariantCulture)
                edits.Add(New WorkbookValueChange(item.Expected, value, item.Permission, item.CalculateBefore, item.SkipUnavailable))
            Next
            writingAndCalculating = True
            Dim receipt As WorkbookValueBatchReceipt
            Try
                ' All native operations run on their own non-reentrant STA. This
                ' synchronous boundary matches current ChangeManager callers;
                ' it never pumps UI messages or moves native objects to the UI.
                receipt = engineTrial.ApplyValuesAsync(edits, engineTrialAreas, cancellation, engineTrialPresentation).GetAwaiter().GetResult()
            Catch ex As Exception
                engineTrialResult = Nothing
                ModelEngineView.Publish(WB, Nothing)
                Return FailedChange(source(0).Change, source(0).Change.WSName, source(0).Change.CellAddress, ex)
            Finally
                writingAndCalculating = False
            End Try
            engineTrialResult = receipt.Results
            ModelEngineView.Publish(WB, receipt.Results)
            Dim skippedMessage = If(receipt.Skipped.Count = 0, "", " " & receipt.Skipped.Count.ToString() & " unavailable input(s) were left unchanged.")
            If Not receipt.Changed Then Return SuccessfulNoAction("The workbook values are unchanged." & skippedMessage)
            Dim group = CreateGroup(description)
            Dim byCell = source.ToDictionary(Function(item) EngineCellKey(item.Expected.Area), StringComparer.OrdinalIgnoreCase)
            For Each edit In receipt.Changes
                Dim item = byCell(EngineCellKey(edit.Before.Area))
                Dim entry As New ChangeHistoryEntryV2 With {
                    .GroupID = group.GroupID, .TimeStamp = If(item.Change.TimeStamp = DateTime.MinValue, Now(), item.Change.TimeStamp),
                    .Description = If(String.IsNullOrWhiteSpace(item.Change.Description), description, item.Change.Description),
                    .WorksheetName = edit.Before.Area.Worksheet, .CellAddress = edit.Before.Area.Address.Split(":"c)(0),
                    .OriginalDisplay = EngineDisplay(edit.Before.State.Value), .ChangedDisplay = EngineDisplay(edit.After.State.Value),
                    .UserName = item.Change.UserName, .DataFormat = item.Change.DataFormat,
                    .EngineBefore = edit.Before.State, .EngineAfter = edit.After.State, .EnginePermission = item.Permission,
                    .EngineCalculateBefore = item.CalculateBefore}
                group.Entries.Add(entry)
            Next
            CommitNewGroup(group)
            FileManager.ExcelModels(ModelID).MarkUserChange()
            For Each entry In group.Entries
                Try
                    MasterChangeLog.AddChangeLogEvent(ToLogEvent(entry, 1, "Apply"))
                Catch ex As Exception
                    Abovo.SummitDiagnostics.WriteLine("Engine history log: " & ex.Message)
                End Try
            Next
            RaiseHistoryChanged(False, group.Entries.Select(Function(entry) entry.WorksheetName))
            PublishEngineCheckSheet()
            Return New AbovoAppCls.AbovoTransaction With {.BSuccess = True, .IntegerReturn = group.Entries.Count, .StrResponseMessage = "Change applied." & skippedMessage}
        End Function

        Private Function ApplyEngineHistory(group As ChangeHistoryGroupV2, redo As Boolean) As AbovoAppCls.AbovoTransaction
            RequireSaveHistoryOwner()
            Dim entries = If(redo, group.Entries.ToList(), group.Entries.AsEnumerable().Reverse().ToList())
            Dim areas = entries.Select(Function(entry) New WorkbookReadArea(entry.WorksheetName,
                WB.Worksheets(entry.WorksheetName).Cells(entry.CellAddress).TopRowIndex,
                WB.Worksheets(entry.WorksheetName).Cells(entry.CellAddress).LeftColumnIndex, 1, 1)).ToList()
            IsApplyingHistory = True
            Try
                Dim snapshots = engineTrial.CaptureCellsAsync(engineTrial.Revision, areas).GetAwaiter().GetResult()
                Dim edits As New List(Of WorkbookValueChange)()
                For i As Integer = 0 To entries.Count - 1
                    Dim expected = If(redo, entries(i).EngineBefore, entries(i).EngineAfter)
                    Dim target = If(redo, entries(i).EngineAfter, entries(i).EngineBefore)
                    If expected Is Nothing OrElse target Is Nothing OrElse Not expected.Matches(snapshots(i).State) Then Throw New InvalidOperationException("Cannot apply history: an input has changed since this action.")
                    edits.Add(New WorkbookValueChange(snapshots(i), target.Value, entries(i).EnginePermission, redo AndAlso entries(i).EngineCalculateBefore))
                Next
                Dim receipt = engineTrial.ApplyValuesAsync(edits, engineTrialAreas, includePresentation:=engineTrialPresentation).GetAwaiter().GetResult()
                engineTrialResult = receipt.Results
                ModelEngineView.Publish(WB, receipt.Results)
            Catch ex As Exception
                engineTrialResult = Nothing
                ModelEngineView.Publish(WB, Nothing)
                Return New AbovoAppCls.AbovoTransaction With {.BError = True, .StrResponseMessage = "History could not be applied: " & ex.Message}
            Finally
                IsApplyingHistory = False
            End Try
            If redo Then
                RedoStack.Remove(group) : UndoStack.Add(group) : group.State = ChangeHistoryStateV2.Applied
            Else
                UndoStack.Remove(group) : RedoStack.Add(group) : group.State = ChangeHistoryStateV2.Undone
            End If
            FileManager.ExcelModels(ModelID).MarkUserChange()
            For Each entry In entries
                Try
                    MasterChangeLog.AddChangeLogEvent(ToLogEvent(entry, If(redo, 5, 4), If(redo, "Redo", "Undo")))
                Catch ex As Exception
                    Abovo.SummitDiagnostics.WriteLine("Engine history log: " & ex.Message)
                End Try
            Next
            RaiseHistoryChanged(True, entries.Select(Function(entry) entry.WorksheetName))
            PublishEngineCheckSheet()
            Return New AbovoAppCls.AbovoTransaction With {.BSuccess = True, .StrResponseMessage = If(redo, "Change redone.", "Change undone.")}
        End Function

        Private Shared Function EngineCellKey(area As WorkbookReadArea) As String
            Return area.Worksheet & "!" & area.Address
        End Function

        Private Sub PublishEngineCheckSheet()
            Dim model = FileManager.ExcelModels(ModelID)
            If EngineEditingResult Is Nothing OrElse WB.DefinedNames.GetDefinedName("Outputs_CheckSheet") Is Nothing Then Return
            Try
                Dim revision = model.CalculationRevision
                Dim result = model.ReadCheckSheetValidation()
                If Not String.IsNullOrWhiteSpace(result.ValidationError) Then Throw New InvalidOperationException(result.ValidationError)
                If EngineEditingResult Is Nothing Then Return
                Dim changed As Boolean
                If model.TryRecordCheckSheetResult(revision, result.HasFailures, changed, worksheetOnly:=True) AndAlso changed Then
                    RecoveryBackupManager.NotifyCheckSheetState(model, result, Nothing)
                End If
            Catch ex As Exception
                'This is post-commit presentation. Never undo a committed input
                'because a view/subscriber could not accept its latest result.
                SystemMessageManager.Publish(ModelID, "Check Sheet could not be refreshed. " & ex.Message,
                    SystemMessageSeverity.Warning, "Check Sheet", model.FileName)
            End Try
        End Sub
        Private Shared Function EngineDisplay(value As Object) As String
            Return Convert.ToString(value, CultureInfo.CurrentCulture)
        End Function
    End Class
End Namespace
