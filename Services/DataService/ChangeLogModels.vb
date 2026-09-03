Imports System.Data

Namespace Abovo

    Public NotInheritable Class MasterChangeLog
        Private Shared ReadOnly SyncRoot As New Object()
        Public Shared ChangeLogEventID As Integer = -1
        Public Shared ChangeLog As DataTable
        Public Shared Event EntryAdded As EventHandler(Of ChangeLogEntryAddedEventArgs)

        Public Shared Function AddChangeLogEvent(ByVal sentEvent As ChangeLogEvent) As Integer
            Dim addedEvent As ChangeLogEvent
            SyncLock SyncRoot
                If ChangeLog Is Nothing Then Initialise()
                ChangeLogEventID += 1
                sentEvent.EventID = ChangeLogEventID
                ChangeLog.Rows.Add(New Object() {
                    ChangeLogEventID, sentEvent.TimeStamp, sentEvent.ModelID,
                    sentEvent.Description, sentEvent.WSName, sentEvent.CellAddress,
                    sentEvent.OriginalValue, sentEvent.ChangedValue, sentEvent.UserName,
                    sentEvent.Status, sentEvent.DataType, sentEvent.GroupID,
                    sentEvent.Operation})
                addedEvent = sentEvent
            End SyncLock

            Try
                RaiseEvent EntryAdded(Nothing, New ChangeLogEntryAddedEventArgs(addedEvent))
            Catch
                'A presentation subscriber must never make workbook logging fail.
            End Try
            Return addedEvent.EventID
        End Function

        Public Shared Function Snapshot() As DataTable
            SyncLock SyncRoot
                If ChangeLog Is Nothing Then Initialise()
                Return ChangeLog.Copy()
            End SyncLock
        End Function

        Public Shared Sub Initialise()
            SyncLock SyncRoot
                ChangeLogEventID = -1
                ChangeLog = New DataTable("MasterChangeLog")
                ChangeLog.Columns.AddRange(New DataColumn() {
                    New DataColumn("EventID", GetType(Integer)),
                    New DataColumn("TimeStamp", GetType(DateTime)),
                    New DataColumn("ModelID", GetType(Integer)),
                    New DataColumn("Description", GetType(String)),
                    New DataColumn("WSName", GetType(String)),
                    New DataColumn("CellAddress", GetType(String)),
                    New DataColumn("OriginalValue", GetType(String)),
                    New DataColumn("ChangedValue", GetType(String)),
                    New DataColumn("UserName", GetType(String)),
                    New DataColumn("Status", GetType(Integer)),
                    New DataColumn("DataType", GetType(String)),
                    New DataColumn("GroupID", GetType(Integer)),
                    New DataColumn("Operation", GetType(String))})
            End SyncLock
        End Sub
    End Class

    Public NotInheritable Class ChangeLogEntryAddedEventArgs
        Inherits EventArgs

        Public ReadOnly Property Entry As ChangeLogEvent

        Public Sub New(ByVal setEntry As ChangeLogEvent)
            Entry = setEntry
        End Sub
    End Class

    Public Structure ChangeLogEvent
        Public EventID As Integer
        Public ModelID As Integer
        Public Description As String
        Public WSName As String
        Public CellAddress As String
        Public OriginalValue As String
        Public ChangedValue As String
        Public TimeStamp As DateTime
        Public UserName As String
        '0 unprocessed, 1 applied, 2 rejected, 3 error, 4 undone,
        '5 redone, 6 notice only.
        Public Status As Integer
        Public DataType As String
        Public GroupID As Integer
        Public Operation As String
    End Structure

End Namespace
