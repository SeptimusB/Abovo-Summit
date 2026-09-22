Namespace Abovo
    'An owner-bound wait form for deliberately long UI actions.
    'Workbook and control operations stay on their existing UI thread.
    Public NotInheritable Class FormSplashScreen
        Implements System.IDisposable

        Private ReadOnly Manager As DevExpress.XtraSplashScreen.SplashScreenManager
        Private IsOpen As Boolean
        Private Completed As Boolean
        Private Disposed As Boolean
        Private Shared ActiveOperations As Integer
        Friend ReadOnly Property IsShowing As Boolean
            Get
                Return IsOpen
            End Get
        End Property
        Friend Shared ReadOnly Property OperationInProgress As Boolean
            Get
                Return ActiveOperations > 0
            End Get
        End Property

        Public Sub New(Owner As System.Windows.Forms.Form,
                       Caption As String,
                       Description As String)
            System.Threading.Interlocked.Increment(ActiveOperations)
            If Owner Is Nothing OrElse Owner.IsDisposed OrElse
               Not Owner.IsHandleCreated OrElse Owner.InvokeRequired Then
                System.Diagnostics.Trace.WriteLine(
                    "[Progress Notice] Skipped: owner is not a live UI-thread form.")
                Return
            End If

            Try
                Manager = New DevExpress.XtraSplashScreen.SplashScreenManager(
                    Owner, GetType(Global.WaitFormA), False, False)
                Manager.ClosingDelay = 900
                Manager.ShowWaitForm()
                IsOpen = True
                Manager.SetWaitFormCaption(Caption)
                Manager.SetWaitFormDescription(Description)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(
                    "[Progress Notice] Unable to show: " & ex.Message)
                Close()
                If Manager IsNot Nothing Then
                    Try
                        Manager.Dispose()
                    Catch cleanupError As System.Exception
                        System.Diagnostics.Trace.WriteLine(
                            "[Progress Notice] Unable to dispose: " &
                            cleanupError.Message)
                    End Try
                End If
            End Try
        End Sub

        Public Sub Update(Description As String)
            If Not IsOpen Then Return
            Try
                Manager.SetWaitFormDescription(Description)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(
                    "[Progress Notice] Unable to update: " & ex.Message)
            End Try
        End Sub

        Public Sub Complete(Optional Description As String = "Complete")
            Finish("Complete", Description)
        End Sub

        Public Sub Fail(Description As String)
            Finish("Not saved", Description)
        End Sub

        Private Sub Finish(Caption As String, Description As String)
            If Not IsOpen Then Return
            Try
                Manager.SetWaitFormCaption(Caption)
                Manager.SetWaitFormDescription(Description)
                Completed = True
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(
                    "[Progress Notice] Unable to complete: " & ex.Message)
            End Try
            Close()
        End Sub

        Private Sub Close()
            If Not IsOpen Then Return
            IsOpen = False
            Try
                Manager.CloseWaitForm()
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(
                    "[Progress Notice] Unable to close: " & ex.Message)
            End Try
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            If Disposed Then Return
            Disposed = True
            System.Threading.Interlocked.Decrement(ActiveOperations)
            Close()
            If Manager Is Nothing Then Return
            If Not Completed Then
                Try
                    Manager.Dispose()
                Catch ex As System.Exception
                    System.Diagnostics.Trace.WriteLine(
                        "[Progress Notice] Unable to dispose: " & ex.Message)
                End Try
                Return
            End If

            'Keep the manager alive until its brief closing delay has elapsed.
            Dim CleanupTimer As New System.Windows.Forms.Timer With {.Interval = 1100}
            AddHandler CleanupTimer.Tick,
                Sub()
                    CleanupTimer.Stop()
                    CleanupTimer.Dispose()
                    Try
                        Manager.Dispose()
                    Catch ex As System.Exception
                        System.Diagnostics.Trace.WriteLine(
                            "[Progress Notice] Unable to dispose: " & ex.Message)
                    End Try
                End Sub
            CleanupTimer.Start()
        End Sub




    End Class

End Namespace

