Imports System.Data
Imports System.Globalization
Imports System.Net
Imports System.Text

Namespace Abovo

    Public Enum SystemMessageSeverity
        Information
        Success
        Warning
        [Error]
    End Enum

    Public NotInheritable Class SystemMessageRecord
        Public Property EventID As Integer
        Public Property TimeStamp As DateTime
        Public Property Severity As SystemMessageSeverity
        Public Property Message As String
        Public Property Source As String
        Public Property Location As String
        Public Property UserName As String

        Public ReadOnly Property SeverityText As String
            Get
                Return Severity.ToString()
            End Get
        End Property
    End Class

    Public NotInheritable Class SystemMessageManager
        Implements IDisposable

        Private Shared ReadOnly RegistryLock As New Object()
        Private Shared ReadOnly Registry As New Dictionary(Of Integer, ManagerRegistration)()

        Private ReadOnly ItemsLock As New Object()
        Private ReadOnly Items As New List(Of SystemMessageRecord)()
        Private ReadOnly ModelID As Integer
        Private IsDisposed As Boolean

        Public Event MessagesChanged As EventHandler

        Private Sub New(ByVal setModelID As Integer)
            ModelID = setModelID
            LoadExistingMessages()
            AddHandler MasterChangeLog.EntryAdded, AddressOf MasterChangeLog_EntryAdded
        End Sub

        Public Shared Function Acquire(ByVal modelID As Integer) As SystemMessageManager
            SyncLock RegistryLock
                Dim registration As ManagerRegistration = Nothing
                If Not Registry.TryGetValue(modelID, registration) Then
                    registration = New ManagerRegistration With {
                        .Manager = New SystemMessageManager(modelID),
                        .ReferenceCount = 0}
                    Registry.Add(modelID, registration)
                End If
                registration.ReferenceCount += 1
                Return registration.Manager
            End SyncLock
        End Function

        Public Shared Sub Publish(ByVal modelID As Integer,
                                  ByVal message As String,
                                  Optional ByVal severity As SystemMessageSeverity = SystemMessageSeverity.Information,
                                  Optional ByVal source As String = "Summit",
                                  Optional ByVal location As String = "")
            If String.IsNullOrWhiteSpace(message) Then Return
            MasterChangeLog.AddChangeLogEvent(New ChangeLogEvent With {
                .ModelID = modelID,
                .Description = message.Trim(),
                .WSName = source,
                .CellAddress = location,
                .TimeStamp = Now(),
                .UserName = Environment.UserName,
                .Status = SeverityToStatus(severity),
                .Operation = "SystemMessage"})
        End Sub

        Public Function SnapshotItems() As List(Of SystemMessageRecord)
            SyncLock ItemsLock
                Return New List(Of SystemMessageRecord)(Items)
            End SyncLock
        End Function

        Public Function SnapshotItemsAfter(ByVal eventID As Integer) As List(Of SystemMessageRecord)
            Dim result As New List(Of SystemMessageRecord)()
            SyncLock ItemsLock
                For Each item As SystemMessageRecord In Items
                    If item.EventID > eventID Then result.Add(item)
                Next
            End SyncLock
            Return result
        End Function

        Public Function CreateTextExport() As String
            Dim output As New StringBuilder()
            output.AppendLine("Abovo Summit system messages")
            output.AppendLine("Created " & Now().ToString("g", CultureInfo.CurrentCulture))
            output.AppendLine(New String("-"c, 80))
            For Each item As SystemMessageRecord In SnapshotItems()
                output.Append(item.TimeStamp.ToString("g", CultureInfo.CurrentCulture)).Append("  ")
                output.Append(item.SeverityText.ToUpperInvariant()).Append("  ")
                output.AppendLine(item.Message)
                If Not String.IsNullOrWhiteSpace(item.Source) OrElse
                   Not String.IsNullOrWhiteSpace(item.Location) Then
                    output.Append("    ").Append(item.Source)
                    If Not String.IsNullOrWhiteSpace(item.Location) Then
                        output.Append(" - ").Append(item.Location)
                    End If
                    output.AppendLine()
                End If
            Next
            Return output.ToString()
        End Function

        Public Function CreateHtmlExport() As String
            Dim output As New StringBuilder()
            output.Append("<!doctype html><html><head><meta charset='utf-8'>")
            output.Append("<style>body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#333}")
            output.Append("table{border-collapse:collapse;width:100%}th,td{padding:7px;border-bottom:1px solid #ddd;text-align:left;vertical-align:top}")
            output.Append("th{background:#075da8;color:white}.Error{color:#c62828}.Warning{color:#9a6700}.Success{color:#147a3d}</style></head><body>")
            output.Append("<h1>Abovo Summit system messages</h1><p>Created ")
            output.Append(WebUtility.HtmlEncode(Now().ToString("g", CultureInfo.CurrentCulture)))
            output.Append("</p><table><thead><tr><th>Time</th><th>Type</th><th>Message</th><th>Source</th><th>Location</th><th>User</th></tr></thead><tbody>")
            For Each item As SystemMessageRecord In SnapshotItems()
                output.Append("<tr class='").Append(item.SeverityText).Append("'><td>")
                output.Append(WebUtility.HtmlEncode(item.TimeStamp.ToString("g", CultureInfo.CurrentCulture))).Append("</td><td>")
                output.Append(WebUtility.HtmlEncode(item.SeverityText)).Append("</td><td>")
                output.Append(WebUtility.HtmlEncode(item.Message)).Append("</td><td>")
                output.Append(WebUtility.HtmlEncode(item.Source)).Append("</td><td>")
                output.Append(WebUtility.HtmlEncode(item.Location)).Append("</td><td>")
                output.Append(WebUtility.HtmlEncode(item.UserName)).Append("</td></tr>")
            Next
            output.Append("</tbody></table></body></html>")
            Return output.ToString()
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If IsDisposed Then Return
            SyncLock RegistryLock
                Dim registration As ManagerRegistration = Nothing
                If Registry.TryGetValue(ModelID, registration) AndAlso
                   Object.ReferenceEquals(registration.Manager, Me) Then
                    registration.ReferenceCount -= 1
                    If registration.ReferenceCount <= 0 Then
                        Registry.Remove(ModelID)
                        RemoveHandler MasterChangeLog.EntryAdded, AddressOf MasterChangeLog_EntryAdded
                        IsDisposed = True
                    End If
                End If
            End SyncLock
        End Sub

        Private Sub LoadExistingMessages()
            Dim snapshot As DataTable = MasterChangeLog.Snapshot()
            For Each row As DataRow In snapshot.Rows
                Dim entry As ChangeLogEvent = RowToEntry(row)
                If AppliesToModel(entry) Then Items.Add(ToMessage(entry))
            Next
        End Sub

        Private Sub MasterChangeLog_EntryAdded(ByVal sender As Object,
                                               ByVal e As ChangeLogEntryAddedEventArgs)
            If IsDisposed OrElse Not AppliesToModel(e.Entry) Then Return
            SyncLock ItemsLock
                Items.Add(ToMessage(e.Entry))
            End SyncLock
            RaiseEvent MessagesChanged(Me, EventArgs.Empty)
        End Sub

        Private Function AppliesToModel(ByVal entry As ChangeLogEvent) As Boolean
            Return entry.ModelID = ModelID OrElse entry.ModelID = -1
        End Function

        Private Shared Function ToMessage(ByVal entry As ChangeLogEvent) As SystemMessageRecord
            Dim source As String = If(entry.WSName, String.Empty)
            If source.Equals("System Message", StringComparison.OrdinalIgnoreCase) Then source = "Summit"
            Return New SystemMessageRecord With {
                .EventID = entry.EventID,
                .TimeStamp = If(entry.TimeStamp = DateTime.MinValue, Now(), entry.TimeStamp),
                .Severity = StatusToSeverity(entry.Status),
                .Message = If(entry.Description, String.Empty),
                .Source = source,
                .Location = If(entry.CellAddress, String.Empty),
                .UserName = If(entry.UserName, String.Empty)}
        End Function

        Private Shared Function RowToEntry(ByVal row As DataRow) As ChangeLogEvent
            Return New ChangeLogEvent With {
                .EventID = CInt(row("EventID")),
                .TimeStamp = CDate(row("TimeStamp")),
                .ModelID = CInt(row("ModelID")),
                .Description = Convert.ToString(row("Description")),
                .WSName = Convert.ToString(row("WSName")),
                .CellAddress = Convert.ToString(row("CellAddress")),
                .UserName = Convert.ToString(row("UserName")),
                .Status = CInt(row("Status")),
                .Operation = Convert.ToString(row("Operation"))}
        End Function

        Private Shared Function StatusToSeverity(ByVal status As Integer) As SystemMessageSeverity
            Select Case status
                Case 1, 5
                    Return SystemMessageSeverity.Success
                Case 2
                    Return SystemMessageSeverity.Warning
                Case 3
                    Return SystemMessageSeverity.Error
                Case Else
                    Return SystemMessageSeverity.Information
            End Select
        End Function

        Private Shared Function SeverityToStatus(ByVal severity As SystemMessageSeverity) As Integer
            Select Case severity
                Case SystemMessageSeverity.Success
                    Return 1
                Case SystemMessageSeverity.Warning
                    Return 2
                Case SystemMessageSeverity.Error
                    Return 3
                Case Else
                    Return 6
            End Select
        End Function

        Private NotInheritable Class ManagerRegistration
            Public Property Manager As SystemMessageManager
            Public Property ReferenceCount As Integer
        End Class
    End Class
End Namespace
