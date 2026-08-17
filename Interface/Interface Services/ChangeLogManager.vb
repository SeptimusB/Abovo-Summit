Imports DevExpress.XtraRichEdit.Model

Namespace Abovo
    Public Class ChangeLogManager

        Public Shared DoChangeTrack As Boolean = True
        Public Shared AppDataEvents As List(Of ChangeLogEvent)
        Public Shared EventCounter As Integer = -1

        Public Sub ChangeLog(SetSource As String, SetRealDesc As String, SetVal As String, SetSheet As String, SetCellAddress As String)

            EventCounter += 1

            Dim newevent As New ChangeLogEvent With {
                .EventID = EventCounter,
                .Source = SetSource,
                .RealDesc = SetRealDesc,
                .IsCancelled = False,
                .EventTime = Now(),
                .EventSheet = SetSheet,
                .EventAddress = SetCellAddress
            }

        End Sub
        Class ChangeLogEvent

            Implements IEnumerable

            Public Source As String
            Public EventID As Integer
            Public RealDesc As String
            Public EventTime As DateTime
            Public OrigVal As String
            Public ChangVal As String
            Public EventSheet As String
            Public EventAddress As String
            Public IsCancelled As Boolean = False

            Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
                Throw New NotImplementedException()
            End Function
        End Class

    End Class

End Namespace