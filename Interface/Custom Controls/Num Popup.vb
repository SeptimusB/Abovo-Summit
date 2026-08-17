Imports DevExpress.XtraEditors
Public Class Num_Popup

    Public MyValue As Integer
    Public MyAction As DialogResult
    'Public MySection As Integer
    Sub New(Title As String, Prompt As String, DefaultValue As Integer)

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        Me.Text = Title
        MyValue = 0
        TBValue.Text = "0"
        'MySection = SetSectionID

    End Sub
    Private Sub SBCancel_Click(sender As Object, e As EventArgs) Handles SBCancel.Click

        MyAction = DialogResult.Cancel

        Me.Close()

    End Sub

    Sub ProcessInput(sender As DevExpress.XtraEditors.SimpleButton, e As EventArgs) Handles SB0.Click, SB1.Click, SB2.Click, SB3.Click, SB4.Click, SB5.Click, SB6.Click, SB7.Click, SB8.Click, SB9.Click

        Dim Btn As SimpleButton = CType(sender, SimpleButton)

        If TBValue.Text = "0" Then

            If Btn.Text <> "0" Then TBValue.Text = Btn.Text

        Else

            TBValue.Text &= Btn.Text

        End If


    End Sub



    Private Sub SBExecute_Click(sender As Object, e As EventArgs) Handles SBExecute.Click

        MyValue = CInt(TBValue.Text)
        MyAction = DialogResult.OK
        Me.Close()

    End Sub

    Private Sub SBClear_Click(sender As Object, e As EventArgs) Handles SBClear.Click

        MyValue = 0
        TBValue.Text = "0"

    End Sub

    Private Sub SBBck_Click(sender As Object, e As EventArgs) Handles SBBck.Click
        If TBValue.Text = "0" Then Exit Sub
        If TBValue.Text.Length = 1 Then
            TBValue.Text = "0"
        Else
            TBValue.Text = TBValue.Text.Substring(0, TBValue.Text.Length - 1)
        End If
    End Sub

    Private Sub SimpleButtonMinus_Click(sender As Object, e As EventArgs) Handles SimpleButtonMinus.Click
        If TBValue.Text = "0" Then
            TBValue.Text = "-"
        ElseIf TBValue.Text = "-" Then
            TBValue.Text = "0"
        Else
            Exit Sub
        End If
    End Sub
End Class
