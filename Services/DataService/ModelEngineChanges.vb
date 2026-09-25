Option Strict On

Imports System.Globalization
Imports System.Threading
Imports Abovo.WorkbookEngines
Imports DevExpress.Spreadsheet

Namespace Abovo
    ' A single editor's captured-before state. This is transient, never saved
    ' in model XML or history, and cannot be reused after a post or refresh.
    Public NotInheritable Class ModelEngineEditTicket
        Friend ReadOnly Owner As ModelChangeManagerV2
        Friend ReadOnly Anchor As WorkbookCalculationResult
        Friend ReadOnly Snapshot As WorkbookCellSnapshot
        Friend ReadOnly Permission As WorkbookValuePermission
        Friend Used As Boolean
        Friend Sub New(owner As ModelChangeManagerV2, anchor As WorkbookCalculationResult,
                       snapshot As WorkbookCellSnapshot, permission As WorkbookValuePermission)
            Me.Owner = owner : Me.Anchor = anchor : Me.Snapshot = snapshot : Me.Permission = permission
        End Sub
    End Class

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
        Private engineDefinitionFingerprint As String
        Public ReadOnly Property HasEngineEditingTrial As Boolean
            Get
                Return engineTrial IsNot Nothing
            End Get
        End Property
        Friend ReadOnly Property CalculationOwnerDescription As String
            Get
                If engineTrial Is Nothing Then Return "DevExpress"
                Return engineTrial.EngineName & " " & engineTrial.EngineVersion & If(EngineEditingResult Is Nothing, " — results unavailable", "")
            End Get
        End Property

        Friend Sub AcceptInitialEngineCalculation()
            Dim model = RequireSaveHistoryOwner()
            If EngineEditingResult Is Nothing Then Throw New InvalidOperationException("Current opening results are required.")
            model.MarkFullCalculationCurrent(model.CalculationRevision, True)
            PublishEngineCheckSheet()
        End Sub
        Public ReadOnly Property EngineEditingResult As WorkbookCalculationResult
            Get
                Return If(engineTrial IsNot Nothing AndAlso engineTrial.IsCurrent(engineTrialResult), engineTrialResult, Nothing)
            End Get
        End Property

        'Called only after model interfaces have detached. Keep the closed view
        'bound: a late reader must not fall through to stale presentation values.
        Friend Sub CloseEngineEditingOwner()
            If engineTrial Is Nothing Then Return
            If Thread.CurrentThread.ManagedThreadId <> saveHistoryOwnerThread Then Throw New InvalidOperationException("Close the model on its owning thread.")
            If ChangeInProgress Then Throw New InvalidOperationException("Wait for the current model operation before closing.")
            engineTrialResult = Nothing
            ModelEngineView.Publish(WB, Nothing)
            engineTrial.CloseAsync().GetAwaiter().GetResult()
        End Sub

        Public Function CaptureEngineEditor(worksheetName As String, address As String,
                                             permission As WorkbookValuePermission) As ModelEngineEditTicket
            If engineTrial Is Nothing Then Return Nothing
            RequireSaveHistoryOwner()
            If IsReadOnlyPreview Then Throw New InvalidOperationException("This preview is read-only.")
            Dim anchor = EngineEditingResult
            If anchor Is Nothing Then Throw New InvalidOperationException("Refresh the model before editing.")
            Dim cell = WB.Worksheets(NormalizeIdentifier(worksheetName)).Cells(NormalizeIdentifier(address))
            Dim area As New WorkbookReadArea(cell.Worksheet.Name, cell.RowIndex, cell.ColumnIndex, 1, 1)
            Dim captured = engineTrial.CaptureCellAsync(anchor.Revision, area).GetAwaiter().GetResult()
            If Not engineTrial.IsCurrent(anchor) Then Throw New InvalidOperationException("The model changed while the editor was opening. Please reopen the editor.")
            If Not captured.State.AllowsValueEdit(permission) Then Throw New InvalidOperationException("This workbook cell is not currently editable.")
            Return New ModelEngineEditTicket(Me, anchor, captured, permission)
        End Function

        Friend Sub RecalculateEngine(kind As WorkbookCalculationKind, Optional publishCheckSheet As Boolean = True)
            Dim model = RequireSaveHistoryOwner()
            If engineTrial Is Nothing Then Throw New InvalidOperationException("No engine editing owner is bound.")
            Dim revision = model.CalculationRevision
            engineTrialResult = Nothing
            ModelEngineView.Publish(WB, Nothing)
            writingAndCalculating = True
            Try
                Dim result = engineTrial.CalculateAndReadAsync(engineTrial.Revision, kind, engineTrialAreas,
                    includePresentation:=engineTrialPresentation).GetAwaiter().GetResult()
                If revision <> model.CalculationRevision Then Throw New InvalidOperationException("The model changed while calculating.")
                engineTrialResult = result
                ModelEngineView.Publish(WB, result)
                model.MarkFullCalculationCurrent(revision, kind = WorkbookCalculationKind.Rebuild)
            Catch
                model.RequireFullRebuild()
                Throw
            Finally
                writingAndCalculating = False
            End Try
            If publishCheckSheet Then PublishEngineCheckSheet()
        End Sub

        Private Function ProcessEngineEditorChange(change As DataChangeEvent) As AbovoAppCls.AbovoTransaction
            If IsApplyingHistory OrElse writingAndCalculating Then Return SuccessfulNoAction("A refresh-time post was ignored while a change was being applied.")
            Try
                RequireSaveHistoryOwner()
                Dim ticket = change.EngineTicket
                If ticket Is Nothing OrElse Not Object.ReferenceEquals(ticket.Owner, Me) OrElse ticket.Used Then
                    Throw New InvalidOperationException("Reopen the editor before changing this value.")
                End If
                If Not engineTrial.IsCurrent(ticket.Anchor) Then Throw New InvalidOperationException("The model has been refreshed since this editor opened. Reopen the editor and check the value before trying again.")
                Dim area = ticket.Snapshot.Area
                Dim cell = WB.Worksheets(NormalizeIdentifier(change.WSName)).Cells(NormalizeIdentifier(change.CellAddress))
                If change.ModelID <> ModelID OrElse Not String.Equals(area.Worksheet, cell.Worksheet.Name, StringComparison.OrdinalIgnoreCase) OrElse
                    area.Row <> cell.RowIndex OrElse area.Column <> cell.ColumnIndex Then Throw New InvalidOperationException("The editor's captured cell does not match this change.")
                ticket.Used = True
                change.WSName = cell.Worksheet.Name : change.CellAddress = cell.GetReferenceA1()
                Return ProcessEngineChanges({New ModelEngineInput(change, ticket.Snapshot, ticket.Permission)}, change.Description)
            Catch ex As Exception
                Return FailedChange(change, change.WSName, change.CellAddress, ex)
            End Try
        End Function

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
            engineDefinitionFingerprint = ComputeEngineDefinitionFingerprint()
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
                ' synchronous boundary matches current ChangeManager callers.
                ' An STA wait can admit paint callbacks; pending display guards
                ' prevent stale reads. Native objects never move to the UI.
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
            FileManager.ExcelModels(ModelID).MarkFullCalculationCurrent(FileManager.ExcelModels(ModelID).CalculationRevision, False)
            For Each entry In group.Entries
                Try
                    MasterChangeLog.AddChangeLogEvent(ToLogEvent(entry, 1, "Apply"))
                Catch ex As Exception
                    Abovo.SummitDiagnostics.WriteLine("Engine history log: " & ex.Message)
                End Try
            Next
            RaiseHistoryChanged(False, group.Entries.Select(Function(entry) entry.WorksheetName))
            RefreshEngineConsumers()
            Return New AbovoAppCls.AbovoTransaction With {.BSuccess = True, .IntegerReturn = group.Entries.Count, .StrResponseMessage = "Change applied." & skippedMessage}
        End Function

        ' Explicit commands (paste/clear/copy) capture all targets together at
        ' command admission. They do not borrow a text editor's older ticket.
        Public Function ProcessEngineCommand(changes As IEnumerable(Of DataChangeEvent),
                                             permissions As IEnumerable(Of WorkbookValuePermission),
                                             description As String) As AbovoAppCls.AbovoTransaction
            RequireSaveHistoryOwner()
            If engineTrial Is Nothing Then Throw New InvalidOperationException("No engine editing owner is bound.")
            Dim anchor = EngineEditingResult
            If anchor Is Nothing Then Throw New InvalidOperationException("Refresh the model before applying this command.")
            Dim inputs = If(changes, Enumerable.Empty(Of DataChangeEvent)()).Take(10001).ToList()
            Dim rules = If(permissions, Enumerable.Empty(Of WorkbookValuePermission)()).Take(10001).ToList()
            If inputs.Count = 0 OrElse inputs.Count > 10000 OrElse inputs.Count <> rules.Count Then Throw New ArgumentException("Supply a permission for each of 1 to 10,000 target cells.")
            Dim areas As New List(Of WorkbookReadArea)
            For Each change In inputs
                If change.ModelID <> ModelID Then Throw New ArgumentException("The command belongs to another model.")
                Dim cell = WB.Worksheets(NormalizeIdentifier(change.WSName)).Cells(NormalizeIdentifier(change.CellAddress))
                areas.Add(New WorkbookReadArea(cell.Worksheet.Name, cell.RowIndex, cell.ColumnIndex, 1, 1))
            Next
            Dim captured = engineTrial.CaptureCellsAsync(anchor.Revision, areas).GetAwaiter().GetResult()
            If Not engineTrial.IsCurrent(anchor) Then Throw New InvalidOperationException("The command was superseded while its inputs were being captured.")
            Dim batch As New List(Of ModelEngineInput)
            For index = 0 To inputs.Count - 1
                Dim change = inputs(index)
                change.WSName = areas(index).Worksheet : change.CellAddress = areas(index).Address.Split(":"c)(0)
                batch.Add(New ModelEngineInput(change, captured(index), rules(index)))
            Next
            Return ProcessEngineChanges(batch, description)
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
            FileManager.ExcelModels(ModelID).MarkFullCalculationCurrent(FileManager.ExcelModels(ModelID).CalculationRevision, False)
            For Each entry In entries
                Try
                    MasterChangeLog.AddChangeLogEvent(ToLogEvent(entry, If(redo, 5, 4), If(redo, "Redo", "Undo")))
                Catch ex As Exception
                    Abovo.SummitDiagnostics.WriteLine("Engine history log: " & ex.Message)
                End Try
            Next
            RaiseHistoryChanged(True, entries.Select(Function(entry) entry.WorksheetName))
            RefreshEngineConsumers()
            Return New AbovoAppCls.AbovoTransaction With {.BSuccess = True, .StrResponseMessage = If(redo, "Change redone.", "Change undone.")}
        End Function

        Private Shared Function EngineCellKey(area As WorkbookReadArea) As String
            Return area.Worksheet & "!" & area.Address
        End Function

        Private Sub RefreshEngineConsumers()
            Try
                PublishEngineCheckSheet()
                FileManager.ExcelModels(ModelID).WBCalcEngine?.RefreshAfterDeferredCalculation()
            Catch ex As Exception
                'The edit and journal have already committed. A presentation
                'failure must not be reported as a failed or rolled-back input.
                SystemMessageManager.Publish(ModelID, "The change was applied, but an interface could not refresh. " & ex.Message,
                    SystemMessageSeverity.Warning, "Calculation refresh")
            End Try
        End Sub

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
