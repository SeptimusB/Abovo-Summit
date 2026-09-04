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

    Public NotInheritable Class ModelChangeManagerV2
        Private ReadOnly WB As IWorkbook
        Private ReadOnly UndoStack As New List(Of ChangeHistoryGroupV2)()
        Private ReadOnly RedoStack As New List(Of ChangeHistoryGroupV2)()
        Private ReadOnly Journal As New List(Of ChangeHistoryGroupV2)()
        Private NextGroupID As Integer
        Private ActiveGroup As ChangeHistoryGroupV2
        Private ActiveGroupDepth As Integer
        Private IsApplyingHistory As Boolean

        Public ReadOnly Property ModelID As Integer
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
            If IsApplyingHistory Then Return SuccessfulNoAction("A refresh-time post was ignored while history was being applied.")
            Dim worksheetName As String = NormalizeIdentifier(sentEvent.WSName)
            Dim address As String = NormalizeIdentifier(sentEvent.CellAddress)
            Try
                Return ProcessResolvedChange(WB.Worksheets(worksheetName).Cells(address), sentEvent)
            Catch ex As Exception
                Return FailedChange(sentEvent, worksheetName, address, ex)
            End Try
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
                Return ProcessResolvedChange(targetCell, sentEvent)
            Catch ex As Exception
                Return FailedChange(sentEvent, NormalizeIdentifier(sentEvent.TargetNR), String.Empty, ex)
            End Try
        End Function

        Private Function ProcessResolvedChange(ByVal targetCell As Cell,
                                               ByVal sentEvent As DataChangeEvent) As AbovoAppCls.AbovoTransaction
            Dim result As New AbovoAppCls.AbovoTransaction("ModelChangeManagerV2.ProcessChange")
            If targetCell Is Nothing Then Return FailedChange(sentEvent, sentEvent.WSName, sentEvent.CellAddress, New InvalidOperationException("The target cell was not found."))

            Dim before As CellSnapshotV2 = CellSnapshotV2.Capture(targetCell)
            Dim automaticGroup As Boolean = ActiveGroup Is Nothing
            Dim group As ChangeHistoryGroupV2 = If(ActiveGroup, CreateGroup(sentEvent.Description))
            Try
                WriteTypedValue(targetCell, sentEvent.ChangedValue, sentEvent.DataFormat)
                FileManager.ExcelModels(ModelID).WBCalcEngine.CalculateWSs()

                'Calculation is allowed to normalise a posted value.  History must
                'therefore describe the authoritative post-calculation cell, not
                'the transient value written immediately before calculation.
                Dim after As CellSnapshotV2 = CellSnapshotV2.Capture(targetCell)
                If before.Matches(targetCell) Then
                    result.BSuccess = True
                    result.StrResponseMessage = "The workbook value is unchanged."
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
                FileManager.ExcelModels(ModelID).IsDirty = True
                result.BSuccess = True
                result.StrResponseMessage = "Change applied."
                If automaticGroup Then RaiseHistoryChanged(False, {targetCell.Worksheet.Name})
                Return result
            Catch ex As Exception
                Dim rollbackFailures As New List(Of String)
                Try
                    before.Apply(targetCell)
                Catch rollbackError As Exception
                    rollbackFailures.Add("Restore cell: " & rollbackError.Message)
                End Try
                Try
                    FileManager.ExcelModels(ModelID).WBCalcEngine.CalculateWSs()
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
                Return FailedChange(sentEvent, targetCell.Worksheet.Name, targetCell.GetReferenceA1(), ex)
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
                    targetSnapshot.Apply(cell)
                    applied.Add(entry)
                Next
                FileManager.ExcelModels(ModelID).WBCalcEngine.CalculateWSs()

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
                    FileManager.ExcelModels(ModelID).WBCalcEngine.CalculateWSs()
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
            FileManager.ExcelModels(ModelID).IsDirty = True
            result.BSuccess = True
            result.StrResponseMessage = If(redo, "Change redone.", "Change undone.")
            RaiseHistoryChanged(True, group.Entries.Select(Function(item) item.WorksheetName))
            Return result
        End Function

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
            If subscribers Is Nothing Then Return

            For Each subscriber As [Delegate] In subscribers.GetInvocationList()
                Try
                    DirectCast(subscriber, EventHandler(Of ChangeHistoryChangedEventArgsV2)).Invoke(Me, args)
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine(
                        "HistoryChanged subscriber failed: " & ex.ToString())
                End Try
            Next
        End Sub

        Private Shared Function NormalizeIdentifier(ByVal value As String) As String
            Return If(value Is Nothing, Nothing, value.Trim())
        End Function

        Private Shared Sub WriteTypedValue(ByVal targetCell As Cell,
                                           ByVal changedValue As Object,
                                           ByVal dataFormat As String)
            If changedValue Is Nothing OrElse Convert.IsDBNull(changedValue) Then
                targetCell.ClearContents()
                Return
            End If
            Select Case If(dataFormat, String.Empty).Trim().ToUpperInvariant()
                Case "S", "FL", "DUMMY", String.Empty
                    targetCell.Value = Convert.ToString(changedValue, CultureInfo.CurrentCulture)
                Case "B"
                    targetCell.Value = If(ConvertToBoolean(changedValue), 1, 0)
                Case "I", "Y"
                    targetCell.Value = ConvertToInteger(changedValue)
                Case "D", "DM"
                    targetCell.Value = CellValue.FromObject(ConvertToDateOrSerial(changedValue))
                Case "N", "P", "C", "M", "SM", "R"
                    targetCell.Value = ConvertToDouble(changedValue)
                Case Else
                    targetCell.Value = CellValue.FromObject(changedValue)
            End Select
        End Sub

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

    Friend NotInheritable Class ChangeHistoryEntryV2
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
