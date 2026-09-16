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
           WindowsUIButtonPanelSaveClose Is Nothing Then Return

        Dim inset As Integer = Math.Max(8, PresentationScaleManager.Scale(12))
        Dim gap As Integer = Math.Max(6, PresentationScaleManager.Scale(8))
        Dim contentTop As Integer = Math.Max(inset * 2, GroupBoxFileActions.Font.Height + inset)
        Dim rightPanelWidth As Integer =
            Math.Max(PresentationScaleManager.Scale(135),
                     Math.Min(PresentationScaleManager.Scale(190),
                              CInt(GroupBoxFileActions.ClientSize.Width * 0.16R)))
        Dim actionPanelHeight As Integer =
            Math.Max(PresentationScaleManager.Scale(100),
                     Math.Min(PresentationScaleManager.Scale(140),
                              CInt(GroupBoxFileActions.ClientSize.Height * 0.24R)))
        Dim contentWidth As Integer =
            Math.Max(20, GroupBoxFileActions.ClientSize.Width - rightPanelWidth - (inset * 2) - gap)
        Dim contentBottom As Integer = GroupBoxFileActions.ClientSize.Height - inset
        Dim actionTop As Integer = Math.Max(contentTop + 20, contentBottom - actionPanelHeight)

        WindowsUIButtonPanelSaveClose.SetBounds(
            inset + contentWidth + gap,
            contentTop,
            rightPanelWidth,
            Math.Max(20, contentBottom - contentTop))
        WindowsUIButtonPanelBPActions.SetBounds(
            inset,
            actionTop,
            contentWidth,
            Math.Max(20, contentBottom - actionTop))
        WebBrowserBPInfo.SetBounds(
            inset,
            contentTop,
            contentWidth,
            Math.Max(20, actionTop - contentTop - gap))
    End Sub
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

            Case "GoOther"



            Case "GoFFR"
                ShowFFRInterface()




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

        StressTester.SetActive()
        If ExcelModels(BPModelID).InterfaceHistory IsNot Nothing Then
            ExcelModels(BPModelID).InterfaceHistory.RecordStandalone(
                InterfaceHistoryDestinationKind.StressTest,
                "Stress Test",
                "Model")
        End If
        StressTester.ShowDialog()
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
        StrFileDescription.Append("</div><div class='detail'>Last Previous Access: ")
        StrFileDescription.Append(WebUtility.HtmlEncode(ExcelModels(BPModelID).FileInfo.LastAccessTime.ToString()))
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

            Case "Spreadsheet"

                ' Close the model and dispose of the interface
                ExcelModels(BPModelID).ShowSpreadsheet()


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
