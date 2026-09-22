Imports System.Diagnostics
Imports System.Globalization

Namespace Abovo
    'Passive diagnostics only: no workbook reads/writes, recalculation or UI pumping.
    Friend NotInheritable Class StructuralInsertBenchmark
        Implements IDisposable

        Private ReadOnly Clock As Stopwatch = Stopwatch.StartNew()
        Private ReadOnly Prefix As String
        Private ReadOnly Totals As New SortedDictionary(Of String, Long)(StringComparer.Ordinal)
        Private Outcome As String = "interrupted"
        Private Disposed As Boolean

        Public Sub New(ModelID As Integer, RuleID As String, Count As Integer, Context As String)
            Prefix = "[Structure Insert Benchmark] operation=" & Guid.NewGuid().ToString("N").Substring(0, 8) &
                ", model=" & ModelID.ToString() & ", rule=" & RuleID & ", records=" & Count.ToString()
            Write("state=started, utc=" & DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture) &
                  ", version=" & AbovoAppCls.DecVersionNumber.ToString("0.00", CultureInfo.InvariantCulture) &
                  ", bitness=" & (IntPtr.Size * 8).ToString() & ", " & Context)
        End Sub

        Public Function Measure(Stage As String, Optional Context As String = "") As IDisposable
            Return New TimedStage(Me, Stage, Context)
        End Function

        Public Sub Complete(Failed As Boolean)
            Outcome = If(Failed, "failed", "ok")
        End Sub

        Private Sub Write(Message As String)
            'A diagnostic listener must never turn a successful workbook edit into a failure.
            Try
                Trace.WriteLine(Prefix & ", " & Message)
            Catch
            End Try
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If Disposed Then Return
            Disposed = True
            Clock.Stop()
            Dim Measured As Long = 0
            Dim Parts As New List(Of String)
            For Each Item In Totals
                Measured += Item.Value
                Parts.Add(Item.Key & "=" & Item.Value.ToString() & " ms")
            Next
            Write("state=finished, " & String.Join(", ", Parts) &
                  ", other=" & Math.Max(0L, Clock.ElapsedMilliseconds - Measured).ToString() & " ms" &
                  ", total=" & Clock.ElapsedMilliseconds.ToString() & " ms, outcome=" & Outcome)
        End Sub

        Private NotInheritable Class TimedStage
            Implements IDisposable
            Private ReadOnly Owner As StructuralInsertBenchmark
            Private ReadOnly Name As String
            Private ReadOnly Context As String
            Private ReadOnly Clock As Stopwatch
            Private Disposed As Boolean

            Public Sub New(Owner As StructuralInsertBenchmark, Name As String, Context As String)
                Me.Owner = Owner
                Me.Name = Name
                Me.Context = If(Context = "", "", ", " & Context)
                Owner.Write("stage=" & Name & Me.Context & ", state=started")
                Clock = Stopwatch.StartNew()
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                If Disposed Then Return
                Disposed = True
                Clock.Stop()
                If Not Owner.Totals.ContainsKey(Name) Then Owner.Totals.Add(Name, 0)
                Owner.Totals(Name) += Clock.ElapsedMilliseconds
                Owner.Write("stage=" & Name & Context & ", elapsed=" & Clock.ElapsedMilliseconds.ToString() & " ms")
            End Sub
        End Class
    End Class
End Namespace
