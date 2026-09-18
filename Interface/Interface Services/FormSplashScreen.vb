Namespace Abovo
    'An owner-bound wait form for deliberately long UI actions.
    'Workbook and control operations stay on their existing UI thread.
    Public NotInheritable Class FormSplashScreen
        Implements System.IDisposable

        Private ReadOnly Manager As DevExpress.XtraSplashScreen.SplashScreenManager
        Private IsOpen As Boolean
        Private Completed As Boolean
        Private Disposed As Boolean

        Public Sub New(Owner As System.Windows.Forms.Form,
                       Caption As String,
                       Description As String)
            If Owner Is Nothing OrElse Owner.IsDisposed OrElse
               Not Owner.IsHandleCreated OrElse Owner.InvokeRequired Then
                System.Diagnostics.Trace.WriteLine(
                    "[Progress Notice] Skipped: owner is not a live UI-thread form.")
                Return
            End If

            Try
                Manager = New DevExpress.XtraSplashScreen.SplashScreenManager(
                    Owner, GetType(Global.WaitFormA), False, False)
                Manager.ClosingDelay = 400
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
            If Not IsOpen Then Return
            Try
                Manager.SetWaitFormCaption("Complete")
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
            Dim CleanupTimer As New System.Windows.Forms.Timer With {.Interval = 500}
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

