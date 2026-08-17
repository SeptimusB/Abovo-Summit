Public Class SplashScreenStart
    Sub New
        InitializeComponent()
        Me.labelCopyright.Text = "Copyright © Abovo Business Services Limited " & DateTime.Now.Year.ToString()
    End Sub

    Public Overrides Sub ProcessCommand(ByVal cmd As System.Enum, ByVal arg As Object)
        MyBase.ProcessCommand(cmd, arg)
    End Sub

    Public Enum SplashScreenCommand
        SomeCommandId
    End Enum

    Private Sub labelCopyright_Click(sender As Object, e As EventArgs) Handles labelCopyright.Click

    End Sub
End Class
