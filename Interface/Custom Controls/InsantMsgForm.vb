
Class InsantMsgForm
    Public Sub New(StrMsg As String)

        InitializeComponent()
        Me.Width = Screen.PrimaryScreen.Bounds.Width * 0.25
        Me.Width = Screen.PrimaryScreen.Bounds.Width * 0.125
        Me.MemoEditMsg.Width = Me.Width
        Me.MemoEditMsg.Height = Me.Height
        MemoEditMsg.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        MemoEditMsg.Properties.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Me.MemoEditMsg.Text = StrMsg
        WindowsUIButtonPanel1.Focus()
    End Sub

    Private WithEvents CloseTimer1 As New Timer With {.Interval = 2200, .Enabled = True}

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles CloseTimer1.Tick
        Me.Dispose()
        Me.Close()
    End Sub

End Class

