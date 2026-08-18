Public Class MessageForm
    Private Sub SimpleButtonCopy_Click(sender As Object, e As EventArgs) Handles SimpleButtonCopy.Click
        On Error Resume Next
        My.Computer.Clipboard.SetText(WebBrowserMessage.DocumentText)
    End Sub
    Public Sub New(StrMessage As String, MessType As MsgBoxStyle, Optional ByVal StrTitle As String = "None", Optional ByVal OverRideButt1 As String = "", Optional ByVal OverRideButt2 As String = "", Optional ByVal OverRideButt3 As String = "")

        ' This call is required by the designer.
        InitializeComponent()
        If StrTitle <> "None" Then Me.Text = StrTitle
        Dim StrDocumentText As String = "<html><body>" & StrMessage & "</body></html>"

        WebBrowserMessage.DocumentText = StrMessage
        Select Case MessType
            Case MsgBoxStyle.YesNoCancel
                SimpleButtonChoiceNo.Visible = True
                SimpleButtonYes.Visible = True
            Case MsgBoxStyle.YesNo

        End Select
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub SimpleButtonChoiceNo_Click(sender As Object, e As EventArgs) Handles SimpleButtonChoiceNo.Click
        DialogResult = Windows.Forms.DialogResult.No
    End Sub

    Private Sub SimpleButtonChoiceYes_Click(sender As Object, e As EventArgs) Handles SimpleButtonYes.Click
        DialogResult = Windows.Forms.DialogResult.Yes
    End Sub


End Class