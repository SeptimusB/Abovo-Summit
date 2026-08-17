Imports Abovo.GeneralFunctions
Imports System.Globalization
Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Namespace Abovo

    Public Class LogDebugDev

        Private Shared LastTime = Now()
        Private Shared ThisTime
        Public Shared StartTime = Now()
        Public Shared DoLog As Boolean = True
        Public Shared DoDataLog As Boolean = True
        Public Shared Counter As Integer = 0
        Public Shared bIsDevelopment As Boolean = True
        Public Shared Sub SystemLog(ItemToLog As String, Optional ByVal sender As Object = Nothing, Optional ByVal StartStop As String = "No")

            If StartStop = "Start" Then ResetTimer()

            If ItemToLog = "CounterReset" Then Counter = 0


            Dim ci As CultureInfo = CultureInfo.InvariantCulture

            If Not IsDebugRun Then Return

            ThisTime = Now()

            Dim Instduration As TimeSpan = ThisTime - LastTime
            Dim OAduration As TimeSpan = ThisTime - StartTime

            If ItemToLog = "Counter" Then

                Debug.Print("+" & Instduration.Seconds & "." & Format(Instduration.Milliseconds, "000") & "s (OA: " & OAduration.Minutes & ":" & Format(OAduration.Seconds, "00") & "." & Format(OAduration.Milliseconds, "000") & "). " & Counter.ToString)
                Counter += 1

            Else
                Debug.Print("+" & Instduration.Seconds & "." & Format(Instduration.Milliseconds, "000") & "s (OA: " & OAduration.Minutes & ":" & Format(OAduration.Seconds, "00") & "." & Format(OAduration.Milliseconds, "000") & "). " & ItemToLog)


            End If

            LastTime = Now()

            If StartStop = "End" Then EndTimer()

        End Sub
        Public Shared Sub ResetTimer(Optional ByVal LogEntry As String = "")

            Debug.Print("''''''''''''''''''''''")
            StartTime = Now()
            LastTime = Now()
            Debug.Print("Timer reset at " & ThisTime.ToString & ". " & LogEntry)
            Debug.Print("''''''''''''''''''''''")

        End Sub

        Public Shared Sub EndTimer(Optional ByVal LogEntry As String = "")

            Dim OAduration As TimeSpan = ThisTime - StartTime

            Debug.Print("''''''''''''''''''''''")
            Debug.Print("Timer end at " & ThisTime.ToString & ". OA: " & OAduration.Minutes & ":" & Format(OAduration.Seconds, "00") & "." & Format(OAduration.Milliseconds, "000") & ". " & LogEntry)
            Debug.Print("''''''''''''''''''''''")

            StartTime = Now()

        End Sub

    End Class

End Namespace
