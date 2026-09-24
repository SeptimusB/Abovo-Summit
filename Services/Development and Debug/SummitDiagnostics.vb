Imports System.Diagnostics

Namespace Abovo
    'Client builds default to quiet. Re-enable with /p:SummitDiagnosticsEnabled=true.
    'Do not use this switch for user messages, audit history or safety scheduling.
    Friend NotInheritable Class SummitDiagnostics
#If SUMMIT_DIAGNOSTICS Then
        Public Const Enabled As Boolean = True
#Else
        Public Const Enabled As Boolean = False
#End If

        <Conditional("SUMMIT_DIAGNOSTICS")>
        Public Shared Sub WriteLine(message As String)
            Trace.WriteLine(message)
        End Sub

        'Explicit runtime trial only; ordinary client traces remain compiled out.
        Friend Shared Sub WriteTrialLine(message As String)
            If CheckSheetWatch.Benchmark Then Trace.WriteLine(message)
        End Sub

        'Only passive timing uses this class. Functional timeouts retain Stopwatch.
        Friend NotInheritable Class DiagnosticTimer
            Private ReadOnly Clock As Stopwatch

            Private Sub New(trial As Boolean)
                If Enabled OrElse trial Then Clock = Stopwatch.StartNew()
            End Sub

            Public Shared Function StartNew(Optional trial As Boolean = False) As DiagnosticTimer
                Return New DiagnosticTimer(trial)
            End Function

            Public ReadOnly Property ElapsedMilliseconds As Long
                Get
                    Return If(Clock Is Nothing, 0L, Clock.ElapsedMilliseconds)
                End Get
            End Property

            Public Sub [Stop]()
                If Clock IsNot Nothing Then Clock.Stop()
            End Sub

            Public Sub Restart()
                If Clock IsNot Nothing Then Clock.Restart()
            End Sub
        End Class
    End Class
End Namespace
