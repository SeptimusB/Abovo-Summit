Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports System.Diagnostics
Imports System.IO

Namespace Abovo
    Public Class HelpManager

        Private Shared ActiveViewer As SummitHelpViewer

        Public Shared Function CreateSuperTooltip(Title As String, TipText As String, HelpIP As String) As SuperToolTip
            Dim superTip As New SuperToolTip()
            Dim item As New ToolTipItem With {.Text = TipText}
            Dim footer As New ToolTipItem With {
                .Text = "<a href=""https://www.abovo-consult.co.uk/help/" & HelpIP & ".html"">Learn more</a>"
            }
            superTip.AllowHtmlText = DefaultBoolean.True
            superTip.Items.AddTitle(Title)
            superTip.Items.Add(item)
            superTip.Items.AddSeparator()
            superTip.Items.Add(footer)
            Return superTip
        End Function

        Public Shared Sub ShowHelpHome(Owner As Form)
            ShowDITHelp(Owner, -1, -1, String.Empty, String.Empty, String.Empty)
        End Sub

        Public Shared Sub ShowDITHelp(Owner As Form,
                                      GSID As Integer,
                                      CSID As Integer,
                                      InterfaceName As String,
                                      SectionName As String,
                                      WorksheetName As String)
            Dim helpRoot As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Help")
            Dim helpFile As String = Path.Combine(helpRoot, "index.html")

            If Not File.Exists(helpFile) Then
                XtraMessageBox.Show(Owner,
                                    "The Summit help library could not be found at:" & Environment.NewLine & helpFile,
                                    "Summit Help",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information)
                Return
            End If

            Dim helpUri As String = New Uri(helpFile).AbsoluteUri &
                "?gsid=" & Uri.EscapeDataString(GSID.ToString()) &
                "&csid=" & Uri.EscapeDataString(CSID.ToString()) &
                "&name=" & Uri.EscapeDataString(If(InterfaceName, String.Empty)) &
                "&section=" & Uri.EscapeDataString(If(SectionName, String.Empty)) &
                "&worksheet=" & Uri.EscapeDataString(If(WorksheetName, String.Empty))

            If ActiveViewer Is Nothing OrElse ActiveViewer.IsDisposed Then ActiveViewer = New SummitHelpViewer(helpRoot)
            ActiveViewer.NavigateTo(helpUri)
            If Not ActiveViewer.Visible Then
                If Owner IsNot Nothing AndAlso Not Owner.IsDisposed Then ActiveViewer.Show(Owner) Else ActiveViewer.Show()
            End If
            If ActiveViewer.WindowState = FormWindowState.Minimized Then ActiveViewer.WindowState = FormWindowState.Normal
            ActiveViewer.BringToFront()
            ActiveViewer.Activate()
        End Sub

        Private NotInheritable Class SummitHelpViewer
            Inherits XtraForm

            Private ReadOnly HelpRoot As String
            Private ReadOnly Browser As WebBrowser
            Private ReadOnly BackButton As SimpleButton
            Private ReadOnly ForwardButton As SimpleButton

            Public Sub New(SetHelpRoot As String)
                HelpRoot = SetHelpRoot
                Text = "Summit Help"
                StartPosition = FormStartPosition.CenterParent
                Size = New Size(1100, 760)
                MinimumSize = New Size(720, 480)
                Icon = FormMainScreen.Icon

                Dim toolbar As New FlowLayoutPanel With {
                    .Dock = DockStyle.Top,
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    .FlowDirection = FlowDirection.LeftToRight,
                    .Padding = New Padding(6),
                    .WrapContents = False
                }
                BackButton = CreateButton("Back", AddressOf GoBack)
                ForwardButton = CreateButton("Forward", AddressOf GoForward)
                toolbar.Controls.Add(BackButton)
                toolbar.Controls.Add(ForwardButton)
                toolbar.Controls.Add(CreateButton("Home", AddressOf GoHome))
                toolbar.Controls.Add(CreateButton("Open help folder", AddressOf OpenHelpFolder))

                Browser = New WebBrowser With {
                    .Dock = DockStyle.Fill,
                    .AllowWebBrowserDrop = False,
                    .IsWebBrowserContextMenuEnabled = True,
                    .ScriptErrorsSuppressed = True,
                    .WebBrowserShortcutsEnabled = True
                }
                AddHandler Browser.Navigated, AddressOf BrowserNavigated
                Controls.Add(Browser)
                Controls.Add(toolbar)
                UpdateNavigationButtons()
            End Sub

            Public Sub NavigateTo(Url As String)
                Browser.Navigate(Url)
            End Sub

            Private Function CreateButton(Caption As String, Handler As EventHandler) As SimpleButton
                Dim button As New SimpleButton With {.Text = Caption, .AutoSize = True}
                AddHandler button.Click, Handler
                Return button
            End Function

            Private Sub GoBack(Sender As Object, E As EventArgs)
                If Browser.CanGoBack Then Browser.GoBack()
            End Sub

            Private Sub GoForward(Sender As Object, E As EventArgs)
                If Browser.CanGoForward Then Browser.GoForward()
            End Sub

            Private Sub GoHome(Sender As Object, E As EventArgs)
                Browser.Navigate(New Uri(Path.Combine(HelpRoot, "index.html")))
            End Sub

            Private Sub OpenHelpFolder(Sender As Object, E As EventArgs)
                Process.Start(New ProcessStartInfo(HelpRoot) With {.UseShellExecute = True})
            End Sub

            Private Sub BrowserNavigated(Sender As Object, E As WebBrowserNavigatedEventArgs)
                UpdateNavigationButtons()
            End Sub

            Private Sub UpdateNavigationButtons()
                BackButton.Enabled = Browser IsNot Nothing AndAlso Browser.CanGoBack
                ForwardButton.Enabled = Browser IsNot Nothing AndAlso Browser.CanGoForward
            End Sub
        End Class
    End Class
End Namespace
