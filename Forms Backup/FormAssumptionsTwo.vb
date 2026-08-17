Public Class FormAssumptionsTwo
    Sub New()

        InitializeComponent()

        AddHandler Me.TabbedViewDefault.QueryControl, AddressOf TabbedView2_QueryControl
    End Sub

    Sub TabbedView1_QueryControl(sender As Object, e As DevExpress.XtraBars.Docking2010.Views.QueryControlEventArgs)
        If e.Document Is assumptionsNavigatorDocument Then
            e.Control = New AssumptionsNavigatorInterface()
        End If
        If e.Document Is stockAssumptionsDocument Then
            e.Control = New StockAssumptionsInterface()
        End If
        If e.Document Is stockItemDetailDocument Then
            e.Control = New StockItemDetailInterface()
        End If
        If e.Control Is Nothing Then
            e.Control = New System.Windows.Forms.Control()
        End If
    End Sub



    Sub TabbedView2_QueryControl(sender As Object, e As DevExpress.XtraBars.Docking2010.Views.QueryControlEventArgs)


        If e.Document Is stockAssumptionsInterfaceDocument Then
            e.Control = New StockAssumptionsInterface()
        End If
        If e.Control Is Nothing Then
            e.Control = New System.Windows.Forms.Control()
        End If
    End Sub

    Private Sub AccordionControlElement9_Click(sender As Object, e As EventArgs) Handles AccordionControlElement9.Click

    End Sub

    Private Sub TileBar1_Click(sender As Object, e As EventArgs) 

    End Sub
End Class