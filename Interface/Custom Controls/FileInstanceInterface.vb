Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.PresentationManager
Imports DevExpress
Imports DevExpress.Utils
Imports DevExpress.XtraBars.Docking2010
Imports System.Net
Imports System.Text
Public Class FileInstanceInterface

    Inherits System.Windows.Forms.UserControl

    Public BPModelID As Integer

    Private ScaleUnits As Single
    Private MyFileInfos As System.IO.FileInfo
    Public MyFilePath As String
    Public MyCompanyName As String

    Public StressTester As StressTest
    Public FFRer As FFRForm
    Private STInit As Boolean
    Private FFRInit As Boolean
    Private MyChildInterfaces() As GroupInterfaceTemplate
    Public Property BPModelInstance As Integer

        Get

            Return BPModelID

        End Get

        Set(value As Integer)

            BPModelID = value

        End Set

    End Property

    Public ReadOnly Property ModelID As Integer
        Get
            Return BPModelID
        End Get
    End Property
    Public Sub New(ModelID As Integer)

        ' This call is required by the designer.
        InitializeComponent()
        BPModelID = ModelID
        ScaleUnits = 5
        'PopulateFileInfo()

        ' Add any initialization after the InitializeComponent() call.
        STInit = False
        FFRInit = False
        FileManager.RegisterModelInterface(BPModelID, Me)
        ConfigureModelActions()
        LayoutFileActionControls()

    End Sub

    Private Sub FileInstanceInterface_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        LayoutFileActionControls()
    End Sub

    Private Sub LayoutFileActionControls()
        If GroupBoxFileActions Is Nothing OrElse WebBrowserBPInfo Is Nothing OrElse
           WindowsUIButtonPanelBPActions Is Nothing OrElse
           WindowsUIButtonPanelSaveClose Is Nothing OrElse
           WindowsUIButtonPanelBPBadge Is Nothing Then Return

        Dim inset As Integer = Math.Max(8, PresentationScaleManager.Scale(12))
        Dim gap As Integer = Math.Max(6, PresentationScaleManager.Scale(8))
        Dim contentTop As Integer = Math.Max(inset * 2, GroupBoxFileActions.Font.Height + inset)
        Dim leftPanelWidth As Integer = Math.Max(PresentationScaleManager.Scale(100),
            TextRenderer.MeasureText("Save As", WindowsUIButtonPanelSaveClose.AppearanceButton.Normal.Font).Width + inset * 2)
        Dim contentWidth As Integer =
            Math.Max(20, GroupBoxFileActions.ClientSize.Width - leftPanelWidth - (inset * 2) - gap)
        Dim contentBottom As Integer = GroupBoxFileActions.ClientSize.Height - inset
        Dim contentLeft As Integer = inset + leftPanelWidth + gap
        Dim actionPanelHeight As Integer = GetNavigationPanelHeight(contentWidth)
        Dim detailsTop As Integer = contentTop + actionPanelHeight + gap

        'Keep the designer's top-and-left layout on initial load, resize and font changes.
        'The badge uses a native WindowsUI button, so its circle matches the other buttons.
        WindowsUIButtonPanelBPBadge.SetBounds(inset, contentTop, leftPanelWidth, actionPanelHeight)
        WindowsUIButtonPanelSaveClose.SetBounds(
            inset,
            detailsTop,
            leftPanelWidth,
            Math.Max(20, contentBottom - detailsTop))
        WindowsUIButtonPanelBPActions.SetBounds(
            contentLeft,
            contentTop,
            contentWidth,
            actionPanelHeight)
        WebBrowserBPInfo.SetBounds(
            contentLeft,
            detailsTop,
            contentWidth,
            Math.Max(20, contentBottom - detailsTop))
    End Sub

    Private Function GetNavigationPanelHeight(availableWidth As Integer) As Integer
        'Reserve enough rows for native button wrapping on narrower windows or larger fonts.
        'Measure the captions instead of shrinking the user's font or hiding trailing actions.
        Dim dpiScale As Single = CSng(DeviceDpi) / 96.0F
        Dim iconWidth As Integer = CInt(Math.Ceiling(42 * dpiScale))
        Dim rows As Integer = 1
        Dim rowWidth As Integer = 0
        Dim captionFont As Font = WindowsUIButtonPanelBPActions.AppearanceButton.Normal.Font
        For Each item As DevExpress.XtraEditors.ButtonPanel.IBaseButton In WindowsUIButtonPanelBPActions.Buttons
            Dim button As WindowsUIButton = TryCast(item, WindowsUIButton)
            If button Is Nothing OrElse Not button.Visible Then Continue For
            Dim captionWidth As Integer = TextRenderer.MeasureText(
                button.Caption, captionFont, Size.Empty, TextFormatFlags.NoPadding).Width
            Dim slotWidth As Integer = Math.Max(iconWidth, captionWidth) +
                CInt(Math.Ceiling(8 * dpiScale)) + 2 * WindowsUIButtonPanelBPActions.ButtonInterval
            If rowWidth > 0 AndAlso rowWidth + slotWidth > availableWidth Then
                rows += 1
                rowWidth = 0
            End If
            rowWidth += slotWidth
        Next
        Dim rowHeight As Integer = CInt(Math.Ceiling(50 * dpiScale)) + captionFont.Height
        Return rows * rowHeight + CInt(Math.Ceiling(8 * dpiScale))
    End Function
    Private Sub ConfigureModelActions()

        If ExcelModels Is Nothing OrElse
           BPModelID < 0 OrElse
           BPModelID >= ExcelModels.Length OrElse
           ExcelModels(BPModelID) Is Nothing Then Return

        Dim IsBusinessPlan As Boolean =
            ExcelModels(BPModelID).Profile IsNot Nothing AndAlso
            String.Equals(
                ExcelModels(BPModelID).Profile.ModelType,
                "AbovoBP",
                StringComparison.OrdinalIgnoreCase)

        Dim profile As WorkbookModelProfile = ExcelModels(BPModelID).Profile
        Dim badge As WindowsUIButton = TryCast(WindowsUIButtonPanelBPBadge.Buttons(0), WindowsUIButton)
        If badge IsNot Nothing AndAlso profile IsNot Nothing Then
            badge.Caption = profile.ButtonCaption
            badge.ToolTip = profile.DisplayName
            WindowsUIButtonPanelBPBadge.AccessibleName = profile.DisplayName
            If Not IsBusinessPlan Then
                'Other models use the neutral workbook glyph rather than the BP initials.
                For Each item As DevExpress.XtraEditors.ButtonPanel.IBaseButton In WindowsUIButtonPanelBPActions.Buttons
                    Dim action As WindowsUIButton = TryCast(item, WindowsUIButton)
                    If action IsNot Nothing AndAlso Convert.ToString(action.Tag) = "Spreadsheet" Then
                        badge.ImageOptions.Assign(action.ImageOptions)
                        Exit For
                    End If
                Next
            End If
        End If

        For Each Item As DevExpress.XtraEditors.ButtonPanel.IBaseButton In
            WindowsUIButtonPanelBPActions.Buttons

            Dim ActionButton As WindowsUIButton = TryCast(Item, WindowsUIButton)
            If ActionButton Is Nothing OrElse ActionButton.Tag Is Nothing Then Continue For

            Dim ActionTag As String = ActionButton.Tag.ToString()

            Select Case ActionTag
                Case "GoAssumpt"
                    ActionButton.Visible =
                        GetGroupID(BPModelID, "Assumptions") >= 0
                Case "GoWorkings"
                    ActionButton.Visible =
                        GetGroupID(BPModelID, "Workings") >= 0
                Case "GoOutputs"
                    ActionButton.Visible =
                        GetGroupID(BPModelID, "Outputs") >= 0
                Case "GoFFR", "GoData", "StressTest"
                    ActionButton.Visible = IsBusinessPlan
            End Select
        Next

    End Sub
    Public Sub ProcessBPInstance(ByVal BPModelID As Integer)

        Me.BPModelID = BPModelID

        ' Set the title of the interface


    End Sub
    Public Sub ShowInterface(InterfaceName As String, Optional ByVal LinkTag As ElementInterfaceLinkTag = Nothing)

        Dim GroupId As Integer = GetGroupID(BPModelID, InterfaceName)

        If GroupId < 0 Then
            MessageBox.Show(
                Me,
                "This model does not define an '" & InterfaceName & "' interface.",
                "Interface unavailable",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
            Return
        End If

        ExcelModels(BPModelID).WBInterface.ShowGroupInterface(BPModelID, GroupId, "Maximised", InterfaceName, Me)

    End Sub


    Private Sub WindowsUIButtonPanelBPActions_ButtonClick(sender As Object, e As DevExpress.XtraBars.Docking2010.ButtonEventArgs) Handles WindowsUIButtonPanelBPActions.ButtonClick

        Dim ButSender As WindowsUIButton = TryCast(e.Button, DevExpress.XtraBars.Docking2010.WindowsUIButton)

        If ButSender Is Nothing Then

            Return

        End If

        Dim tag As String = ButSender.Tag.ToString()

        Select Case tag

            Case "GoAssumpt"
                ' OpenAssumptionsInterface

                ShowInterface("Assumptions")

            Case "GoWorkings"

                ShowInterface("Workings")

            Case "GoOutputs"

                ShowInterface("Outputs")

            Case "GoData"
                ExcelModels(BPModelID).WBInterface.ShowGroupInterface(
                    BPModelID, -1, "Maximised", "Combined", Me)


            Case "GoFFR"
                ShowFFRInterface()

            Case "Spreadsheet"
                ExcelModels(BPModelID).ShowSpreadsheet()
            Case "StressTest"
                ShowStressTestInterface()

        End Select

    End Sub

    Public Sub ShowFFRInterface()
        If Not FFRInit OrElse FFRer Is Nothing OrElse FFRer.IsDisposed Then
            Me.Cursor = Cursors.WaitCursor
            Try
                FFRer = New FFRForm(BPModelID)
                FFRInit = True
            Finally
                Me.Cursor = Cursors.Default
            End Try
        End If

        FFRer.Show()
        FFRer.Activate()
        FFRer.BringToFront()

        If ExcelModels(BPModelID).InterfaceHistory IsNot Nothing Then
            ExcelModels(BPModelID).InterfaceHistory.RecordStandalone(
                InterfaceHistoryDestinationKind.FinancialForecastReturn,
                "Financial Forecast Return",
                "Model")
        End If
    End Sub

    Public Sub ShowStressTestInterface()
        If Not STInit OrElse StressTester Is Nothing OrElse StressTester.IsDisposed Then
            StressTester = New StressTest(BPModelID)
            STInit = True
        End If

        If ExcelModels(BPModelID).InterfaceHistory IsNot Nothing Then
            ExcelModels(BPModelID).InterfaceHistory.RecordStandalone(
                InterfaceHistoryDestinationKind.StressTest,
                "Stress Test",
                "Model")
        End If
        StressTester.Show()
        If StressTester.WindowState = FormWindowState.Minimized Then StressTester.WindowState = FormWindowState.Normal
        StressTester.Activate()
        StressTester.BringToFront()
    End Sub

    Public Sub CloseStandaloneInterfaces()
        'User close hides these windows; model close must dispose them while
        'their workbook is still available, including history subscriptions.
        Try
            If FFRer IsNot Nothing AndAlso Not FFRer.IsDisposed Then FFRer.ManualDispose()
        Finally
            If StressTester IsNot Nothing AndAlso Not StressTester.IsDisposed Then
                StressTester.Clearup()
                StressTester.Dispose()
            End If
        End Try
        FFRer = Nothing
        StressTester = Nothing
        FFRInit = False
        STInit = False
    End Sub

    Public Sub PopulateFileInfo()

        ScaleUnits = GetDisplayScale(Me)

        Dim prominentFontPixels As Integer = Math.Max(16, CInt(Math.Round(17.0F * ScaleUnits)))
        Dim detailFontPixels As Integer = Math.Max(14, CInt(Math.Round(15.0F * ScaleUnits)))
        Dim StrFileDescription As New StringBuilder()

        MyFilePath = ExcelModels(BPModelID).FileName
        MyCompanyName = ExcelModels(BPModelID).WBStructure.CompanyName

        Dim ModelDescription As String =
            If(ExcelModels(BPModelID).Profile Is Nothing,
               "Abovo model",
               ExcelModels(BPModelID).Profile.DisplayName)

        StrFileDescription.Append("<html><head><style>")
        StrFileDescription.Append("body{font-family:Verdana,sans-serif;font-size:")
        StrFileDescription.Append(detailFontPixels)
        StrFileDescription.Append("px;line-height:1.35;margin:8px;color:#202020;overflow-wrap:anywhere;}")
        StrFileDescription.Append(".primary{font-size:")
        StrFileDescription.Append(prominentFontPixels)
        StrFileDescription.Append("px;margin:0 0 5px 0;}.detail{margin:0 0 3px 0;}")
        StrFileDescription.Append("</style></head><body>")
        StrFileDescription.Append("<div class='primary'>Model type: ")
        StrFileDescription.Append(WebUtility.HtmlEncode(ModelDescription))
        StrFileDescription.Append("</div><div class='primary'>Model name: ")
        StrFileDescription.Append(WebUtility.HtmlEncode(ExcelModels(BPModelID).WBStructure.CompanyName))
        StrFileDescription.Append("</div><div class='primary'>Start Date: ")
        StrFileDescription.Append(WebUtility.HtmlEncode(Convert.ToString(ExcelModels(BPModelID).WBStructure.StartDate)))
        StrFileDescription.Append(" (<a href='editbpdate'>edit</a>)</div><div class='detail'>File Name: ")
        StrFileDescription.Append(WebUtility.HtmlEncode(ExcelModels(BPModelID).FileName))
        StrFileDescription.Append("</div><div class='detail'>Opened: ")
        StrFileDescription.Append(WebUtility.HtmlEncode(Now().ToString()))
        StrFileDescription.Append("</div><div class='detail'>Created: ")
        StrFileDescription.Append(WebUtility.HtmlEncode(ExcelModels(BPModelID).FileInfo.CreationTime.ToString()))
        StrFileDescription.Append("</div><div class='detail'>Previous File Access: ")
        StrFileDescription.Append(WebUtility.HtmlEncode(
            If(ExcelModels(BPModelID).PreviousFileAccessTime = DateTime.MinValue,
               "Not recorded",
               ExcelModels(BPModelID).PreviousFileAccessTime.ToString())))
        StrFileDescription.Append("</div><div class='detail'>Size: ")
        StrFileDescription.Append(WebUtility.HtmlEncode(
            Format((ExcelModels(BPModelID).FileInfo.Length / 1000000), "###.##") & "Mb"))
        StrFileDescription.Append("</div></body></html>")

        WebBrowserBPInfo.DocumentText = StrFileDescription.ToString()
        LayoutFileActionControls()

    End Sub
    Private Sub WindowsUIButtonPanelSaveClose_ButtonClick(sender As Object, e As ButtonEventArgs) Handles WindowsUIButtonPanelSaveClose.ButtonClick

        Dim ButSender As WindowsUIButton = TryCast(e.Button, DevExpress.XtraBars.Docking2010.WindowsUIButton)
        If ButSender Is Nothing Then
            Return
        End If
        Dim tag As String = ButSender.Tag.ToString()

        Select Case tag

            Case "SaveBP"

                Me.Cursor = Cursors.WaitCursor
                Try
                    ExcelModels(BPModelID).SaveFile()
                Finally
                    Me.Cursor = Cursors.Default
                End Try

            Case "SaveBPAs"

                'Me.Cursor = Cursors.WaitCursor
                ExcelModels(BPModelID).SaveFileAs()
                'Me.Cursor = Cursors.Default

            Case "CloseBP"

                ' Close the model and dispose of the interface
                If ExcelModels(BPModelID).CommitToCloseModel.StringReturn = "Proceed" Then

                    FileManager.CloseModel(BPModelID)
                    FormMainScreen.RemoveModel(BPModelID)
                    Me.Dispose()

                End If

        End Select

    End Sub
    Public Function SaveFileAs() As Boolean
        Me.Cursor = Cursors.WaitCursor
        Try
            Return ExcelModels(BPModelID).SaveFileAs()
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Function
    Sub SetScale()
        Me.WindowsUIButtonPanelBPActions.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelBPActions.AppearanceButton.Normal.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelBPActions.AppearanceButton.Hovered.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelBPActions.AppearanceButton.Pressed.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelSaveClose.AppearanceButton.Normal.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelSaveClose.AppearanceButton.Hovered.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelSaveClose.AppearanceButton.Pressed.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelBPBadge.AppearanceButton.Normal.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelBPBadge.AppearanceButton.Hovered.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelBPBadge.AppearanceButton.Pressed.Font = GetDisplayFont("Small", Me)
        Me.GroupBoxFileActions.Font = GetDisplayFont("Small", Me)
        LayoutFileActionControls()
    End Sub

    Private Sub ApplyPresentationScale()
        SetScale()
        If ExcelModels IsNot Nothing AndAlso BPModelID >= 0 AndAlso
           BPModelID < ExcelModels.Length AndAlso ExcelModels(BPModelID) IsNot Nothing Then
            PopulateFileInfo()
        End If
    End Sub
End Class
