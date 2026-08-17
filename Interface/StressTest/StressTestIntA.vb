Imports Abovo
Public Class StressTestIntA

    Inherits System.Windows.Forms.UserControl

    Private ParST As StressTest
    Public Property ParentST As StressTest
        Get
            Return ParST
        End Get
        Set(value As StressTest)
            ParST = value
        End Set
    End Property

    Sub New()

        InitializeComponent()

    End Sub

    Private Sub SimpleButtonStressModeSwitch_Click(sender As Object, e As EventArgs)

        ParST.StressTestModeSwitch(True)

    End Sub
End Class
