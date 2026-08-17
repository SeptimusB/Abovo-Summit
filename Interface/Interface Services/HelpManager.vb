Imports DevExpress.Utils

Namespace Abovo
    Public Class HelpManager

        Public Shared Function CreateSuperTooltip(Title As String, TipText As String, HelpIP As String) As SuperToolTip

            Dim superTip As New SuperToolTip()
            Dim item As New ToolTipItem()

            item.Text = TipText
            Dim footer As New ToolTipItem()
            superTip.AllowHtmlText = DefaultBoolean.True

            footer.Text = "<a href=\""https://www.abovo-consult.co.uk/help/" & HelpIP & ".html\>Learn more</a>Learn more at www.abovo-consult.co.uk/help"
            superTip.Items.AddTitle(Title)
            superTip.Items.Add(item)
            superTip.Items.AddSeparator()
            superTip.Items.Add(footer)
            Return superTip

        End Function


    End Class

End Namespace
