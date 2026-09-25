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
    Private SaveButtonBinding As ModelSaveButtonBinding
    Private CheckSheetModel As FileManager.ExcelModel
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
        SaveButtonBinding = New ModelSaveButtonBinding(Me, ExcelModels(BPModelID), WindowsUIButtonPanelSaveClose)
        WebBrowserBPInfo.Tag = PresentationLayout.BrowserOwnsScale
        CheckSheetModel = ExcelModels(BPModelID)
        AddHandler CheckSheetModel.CheckSheetStatusChanged, AddressOf CheckSheetStatusChanged
        AddHandler CheckSheetModel.MetadataChanged, AddressOf CheckSheetStatusChanged
        SetScale()
        LayoutFileActionControls()

    End Sub

    Private Sub CheckSheetStatusChanged(sender As Object, e As EventArgs)
        If Not IsDisposed AndAlso Not Disposing Then PopulateFileInfo()
    End Sub

    Private Sub ReleaseCheckSheetBinding(sender As Object, e As EventArgs) Handles Me.Disposed
        If CheckSheetModel IsNot Nothing Then RemoveHandler CheckSheetModel.CheckSheetStatusChanged, AddressOf CheckSheetStatusChanged
        If CheckSheetModel IsNot Nothing Then RemoveHandler CheckSheetModel.MetadataChanged, AddressOf CheckSheetStatusChanged
        CheckSheetModel = Nothing
    End Sub

    Private Sub FileInfoNavigating(sender As Object, e As WebBrowserNavigatingEventArgs) Handles WebBrowserBPInfo.Navigating
        If e.Url IsNot Nothing AndAlso e.Url.Scheme.Equals("summit-checksheet", StringComparison.OrdinalIgnoreCase) Then
            e.Cancel = True
            RecoveryBackupManager.OpenCheckSheet(CheckSheetModel)
        End If
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
        Dim buttonBackgrounds = TryCast(WindowsUIButtonPanelSaveClose.ButtonBackgroundImages, ImageCollection)
        If buttonBackgrounds IsNot Nothing Then
            leftPanelWidth = Math.Max(leftPanelWidth, CInt(buttonBackgrounds.ImageSize.Width * DeviceDpi / 96.0F) + inset * 2)
        End If
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
        Return PresentationLayout.ButtonPanelHeight(WindowsUIButtonPanelBPActions, availableWidth)
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
        ExcelModels(BPModelID).EnsureDeferredSaveResultsCurrent("Opening Financial Forecast Return...")
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
        If Not ModelSafetyManager.CanOpenDirectWorkbookSurface(BPModelID, "Stress Test scenario operations") Then Return
        ExcelModels(BPModelID).EnsureDeferredSaveResultsCurrent("Opening Stress Test...")
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

    Private ReadOnly InfoOpenedAt As DateTime = DateTime.Now

    Public Sub PopulateFileInfo()

        ScaleUnits = GetDisplayScale(Me)

        'HTML points share the native controls' scale; do not also apply browser zoom.
        Dim detailFontPoints As Single = GetDisplayFont("Small", Me).SizeInPoints
        Dim prominentFontPoints As Single = GetDisplayFont("Medium", Me).SizeInPoints

        MyFilePath = ExcelModels(BPModelID).FileName
        MyCompanyName = ExcelModels(BPModelID).WBStructure.CompanyName
        Dim modelPage = TryCast(Parent, DevExpress.XtraTab.XtraTabPage)
        If modelPage IsNot Nothing Then
            Dim modelPrefix = If(ExcelModels(BPModelID).Profile IsNot Nothing AndAlso
                ExcelModels(BPModelID).Profile.ModelType = "AbovoDSA", "DSA", "HABP")
            modelPage.Text = "(" & (BPModelID + 1).ToString() & ") " & modelPrefix & " " & MyCompanyName
            modelPage.Tooltip = MyFilePath
        End If

        Dim ModelDescription As String =
            If(ExcelModels(BPModelID).Profile Is Nothing,
               "Abovo model",
               ExcelModels(BPModelID).Profile.DisplayName)

        Dim model = ExcelModels(BPModelID)
        WebBrowserBPInfo.DocumentText = BuildFileSummaryHtml(ModelDescription, MyCompanyName,
            Convert.ToString(model.WBStructure.StartDate), model.FileName,
            InfoOpenedAt.ToString("dd/MM/yyyy HH:mm:ss"), model.FileInfo.CreationTime.ToString("dd/MM/yyyy HH:mm:ss"),
            If(model.PreviousFileAccessTime = DateTime.MinValue, "Not recorded", model.PreviousFileAccessTime.ToString("dd/MM/yyyy HH:mm:ss")),
            (model.FileInfo.Length / 1000000.0R).ToString("0.##") & " MB", detailFontPoints, prominentFontPoints,
            model.DeferredSaveResultsPending, model.RecoverySourcePath, model.CheckSheetWarningActive, model.CheckSheetWarningNeedsRecheck)
        LayoutFileActionControls()

    End Sub
    Friend Shared Function BuildFileSummaryHtml(modelType As String, company As String, startDate As String,
                                               path As String, opened As String, created As String,
                                               previous As String, size As String, bodyPoints As Single,
                                               titlePoints As Single, Optional resultsPending As Boolean = False,
                                               Optional recoverySource As String = Nothing,
                                               Optional checkSheetWarning As Boolean = False,
                                               Optional checkSheetRecheck As Boolean = False) As String
        'IE-compatible layout: no dependency on WebView2, flex/grid or script.
        Dim html As New StringBuilder("<!DOCTYPE html><html><head><meta http-equiv='X-UA-Compatible' content='IE=edge'><style>")
        html.Append("body{margin:0;padding:12px;font-family:'Segoe UI',Arial,sans-serif;color:#243746;background:white;font-size:")
        html.Append(bodyPoints.ToString("0.##", Globalization.CultureInfo.InvariantCulture))
        html.Append("pt;line-height:1.45}.card{border:1px solid #dce5ec;border-top:4px solid #005baa;padding:18px 22px;max-width:1050px}")
        html.Append(".type{color:#005baa;font-weight:600;margin-bottom:4px}h1{font-weight:600;margin:0 0 16px;font-size:")
        html.Append(titlePoints.ToString("0.##", Globalization.CultureInfo.InvariantCulture))
        html.Append("pt}.key{display:inline-block;vertical-align:top;margin:0 32px 16px 0}.label{color:#617280;font-size:90%;font-weight:400}")
        html.Append(".value{margin-top:3px}table{width:100%;border-collapse:collapse;table-layout:fixed}td{padding:9px 0;border-top:1px solid #edf1f4;vertical-align:top;word-wrap:break-word}td.label{width:32%;padding-right:12px}")
        html.Append("</style></head><body><div class='card'><div class='type'>").Append(WebUtility.HtmlEncode(modelType))
        html.Append("</div><h1>")
        If checkSheetWarning Then html.Append("<a href='summit-checksheet://open' style='color:#b22222' title='Review the Check Sheet; run integrity check in Options'>").
            Append(If(checkSheetRecheck, "(Check sheet: recheck required)", "(Check sheet)")).Append("</a> ")
        html.Append(WebUtility.HtmlEncode(company)).Append("</h1>")
        For Each pair In {New String() {"Plan start", startDate}, New String() {"File size", size}}
            html.Append("<div class='key'><div class='label'>").Append(pair(0)).Append("</div><div class='value'>")
            html.Append(WebUtility.HtmlEncode(pair(1))).Append("</div></div>")
        Next
        html.Append("<table>")
        If Not String.IsNullOrWhiteSpace(recoverySource) Then
            html.Append("<tr><td class='label'>Recovery copy</td><td><strong>This is not your normal business-plan file.</strong><br>Use Save As and choose Excel Binary Workbook (.xlsb). The original folder and filename below will be suggested. Confirm replacement, or choose a different name.<br>")
            html.Append(WebUtility.HtmlEncode(recoverySource)).Append("</td></tr>")
        End If
        If resultsPending Then
            html.Append("<tr><td class='label'>Calculation</td><td>Inputs saved; full results pending. Results update when required or when reopened in Summit.</td></tr>")
        End If
        For Each pair In {New String() {"File", path}, New String() {"Opened", opened},
                          New String() {"Created", created}, New String() {"Previous file access", previous}}
            html.Append("<tr><td class='label'>").Append(pair(0)).Append("</td><td>")
            html.Append(WebUtility.HtmlEncode(pair(1))).Append("</td></tr>")
        Next
        Return html.Append("</table></div></body></html>").ToString()
    End Function
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
        PresentationLayout.ApplyButtonPanel(WindowsUIButtonPanelBPActions, Me)
        PresentationLayout.ApplyButtonPanel(WindowsUIButtonPanelSaveClose, Me)
        PresentationLayout.ApplyButtonPanel(WindowsUIButtonPanelBPBadge, Me)
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
