Option Strict On

Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace Abovo.WorkbookEngines
    Friend NotInheritable Class WorkbookEngineStaHost
        Private ReadOnly ready As New TaskCompletionSource(Of Control)(TaskCreationOptions.RunContinuationsAsynchronously)
        Private ReadOnly stopped As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
        Private ReadOnly gate As New Object()
        Private ReadOnly pending As New Queue(Of Action)()
        Private closing As Boolean
        Private executing As Boolean
        Private scheduled As Boolean
        Private context As ApplicationContext

        Friend Sub New()
            Dim worker As New Thread(AddressOf RunLoop) With {.IsBackground = True, .Name = "Summit workbook calculation owner"}
            worker.SetApartmentState(ApartmentState.STA)
            worker.Start()
        End Sub

        Private Sub RunLoop()
            Try
                Using dispatcher As New Control()
                    Dim handle = dispatcher.Handle
                    context = New ApplicationContext()
                    ready.TrySetResult(dispatcher)
                    Application.Run(context)
                    context.Dispose()
                End Using
                stopped.TrySetResult(True)
            Catch ex As Exception
                ready.TrySetException(ex)
                stopped.TrySetException(ex)
            End Try
        End Sub

        Friend Async Function InvokeAsync(Of T)(action As Func(Of T), cancellation As CancellationToken) As Task(Of T)
            Dim dispatcher = Await ready.Task.ConfigureAwait(False)
            Dim completion As New TaskCompletionSource(Of T)(TaskCreationOptions.RunContinuationsAsynchronously)
            SyncLock gate
                If closing Then Throw New ObjectDisposedException(NameOf(WorkbookEngineStaHost))
                Enqueue(dispatcher, Sub()
                    If cancellation.IsCancellationRequested Then
                        completion.TrySetCanceled()
                        Return
                    End If
                    Try
                        Dim value = action()
                        ' Native calculation is an atomic operation. Cancellation
                        ' discards its result; it does not interrupt COM halfway.
                        If cancellation.IsCancellationRequested Then
                            completion.TrySetCanceled()
                        Else
                            completion.TrySetResult(value)
                        End If
                    Catch ex As Exception
                        completion.TrySetException(ex)
                    End Try
                End Sub)
            End SyncLock
            Return Await completion.Task.ConfigureAwait(False)
        End Function

        Friend Async Function StopAsync(cleanup As Action) As Task
            Dim dispatcher = Await ready.Task.ConfigureAwait(False)
            SyncLock gate
                If Not closing Then
                    closing = True
                    Enqueue(dispatcher, Sub()
                        Try
                            cleanup()
                        Catch ex As Exception
                            stopped.TrySetException(ex)
                        Finally
                            context.ExitThread()
                        End Try
                    End Sub)
                End If
            End SyncLock
            Await stopped.Task.ConfigureAwait(False)
        End Function

        Private Sub Enqueue(dispatcher As Control, action As Action)
            pending.Enqueue(action)
            Schedule(dispatcher)
        End Sub

        Private Sub Schedule(dispatcher As Control)
            If executing OrElse scheduled OrElse pending.Count = 0 Then Return
            scheduled = True
            dispatcher.BeginInvoke(New Action(Sub() DrainOne(dispatcher)))
        End Sub

        Private Sub DrainOne(dispatcher As Control)
            Dim action As Action
            SyncLock gate
                scheduled = False
                If executing OrElse pending.Count = 0 Then Return
                executing = True
                action = pending.Dequeue()
            End SyncLock
            Try
                action()
            Finally
                SyncLock gate
                    executing = False
                    Schedule(dispatcher)
                End SyncLock
            End Try
        End Sub
    End Class
End Namespace
