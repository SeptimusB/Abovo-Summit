Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.GeneralFunctions
Imports Abovo.LogDebugDev
Imports DevExpress.CodeParser
Imports DevExpress.DataAccess.Native.Web
Imports DevExpress.Office.History
Imports DevExpress.Skins
Imports DevExpress.Spreadsheet
Imports DevExpress.UserSkins
Imports DevExpress.Utils.CommonDialogs
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.XtraCharts.Native
Imports DevExpress.XtraEditors
Imports DevExpress.XtraPrinting.Native
Imports DevExpress.XtraSplashScreen
Imports DevExpress.XtraSpreadsheet.Forms


Public Class FormMainScreen

#If DEBUG Then
    Private Const DebugAutoOpenModelPath As String = "Z:\Sandbox\TestFileMigrated.xlsb"
#End If

    Dim rs As New Resizer

    Private LrgFontSize As Integer
    Private MediumFontSize As Integer
    Private SmallFontSize As Integer
    Private ScaleFactor As Single
    Private ScaleUnits As Single
    Private LastHTMLFontSize As Integer
    Private AmMaximised As Boolean = False
    Dim SplashScreenManagerStartup As DevExpress.XtraSplashScreen.SplashScreenManager = New DevExpress.XtraSplashScreen.SplashScreenManager(Me, GetType(Global.SplashScreenStart), True, True)
    Private HTMLFontSize As Integer
    Private MyFileInfos As System.IO.FileInfo
    Public AssumptionForm As FormAssumptionsTwo
    Public ActiveModel As Integer
    Private FileInstances() As FileInstanceInterface
    Private FileInstanceIndex As Integer = -1
    Private ModelsClosedForShutdown As Boolean

    Public Sub New()





        'SplashScreenManagerStartup.ShowWaitForm()
        DevExpress.XtraEditors.WindowsFormsSettings.SmartMouseWheelProcessing = False
        WindowsFormsSettings.DefaultFont = New System.Drawing.Font("Segoe UI", 10)

        InitializeComponent()

        If String.IsNullOrWhiteSpace(ApplicationConfiguration.CurrentApplicationPath) Then
            ApplicationConfiguration.Initialize()
        End If
        If String.IsNullOrWhiteSpace(ApplicationConfiguration.BaseApplicationTitle) Then
            ApplicationConfiguration.BaseApplicationTitle = "abovo summit"
        End If

        ModelCollection.Initialise()
        FileManager.Initialise(Me)
        Abovo.AbovoAppCls.Initialise()

        Me.Text = ApplicationConfiguration.BaseApplicationTitle


        SetInitialSizes()

        rs.FindAllControls(Me)


        LrgFontSize = DefaultLrgFontSize
        MediumFontSize = DefaultMediumFontSize
        SmallFontSize = DefaultSmallFontSize


        XtraTabControlMainNavigator.LookAndFeel.UseDefaultLookAndFeel = False
        XtraTabControlMainNavigator.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        XtraTabControlMainNavigator.Appearance.BackColor = Color.White

        SetBrowserText()

        SplashScreenManagerStartup.CloseWaitForm()
        'AbovoBP.InitialiseBP()
        'OpenModelProceedure("Z:\Sandbox\TestFileXLSB.xlsb")

    End Sub
    Sub ResizeFonts()

        ScaleFactor = Me.Width / 2100

        Me.GroupBoxProgramDetails.Font = GetFont("Small", Me.ScaleFactor)

        Me.hideContainerRight.Font = GetFont("Small", Me.ScaleFactor)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.ActiveTab.Font = GetFont("Medium", Me.ScaleFactor)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.HidePanelButton.Font = GetFont("Medium", Me.ScaleFactor)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.HidePanelButtonActive.Font = GetFont("Medium", Me.ScaleFactor)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaption.Font = GetFont("Medium", Me.ScaleFactor)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaptionActive.Font = GetFont("Medium", Me.ScaleFactor)

        Me.BarAndDockingControllerMainScreen.AppearancesBar.Dock.Font = GetFont("Medium", Me.ScaleFactor)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.HidePanelButton.Font = GetFont("Medium", Me.ScaleFactor)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaption.Font = GetFont("Medium", Me.ScaleFactor)
        Me.XtraTabControlMainNavigator.AppearancePage.HeaderActive.Font = GetFont("Medium", Me.ScaleFactor, False, True)
        Me.XtraTabControlMainNavigator.AppearancePage.Header.Font = GetFont("Medium", Me.ScaleFactor)
        Me.XtraTabControlMainNavigator.Appearance.Font = GetFont("Medium", Me.ScaleFactor)
        Me.XtraTabControlMainNavigator.AppearancePage.HeaderHotTracked.Font = GetFont("Medium", Me.ScaleFactor)

        WindowsUIButtonPanelExitHelp.Font = GetFont("Small", Me.ScaleFactor)
        Me.WindowsUIButtonPanelBPActions.Font = GetFont("Small", Me.ScaleFactor)

        WindowsUIButtonPanelOpenCompare.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        WindowsUIButtonPanelOpenCompare.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        WindowsUIButtonPanelOpenCompare.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)

        WindowsUIButtonPanelBPActions.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        WindowsUIButtonPanelBPActions.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        WindowsUIButtonPanelBPActions.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)

        WindowsUIButtonPanelExitHelp.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        WindowsUIButtonPanelExitHelp.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        WindowsUIButtonPanelExitHelp.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)

        WindowsUIButtonPanelSaveClose.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        WindowsUIButtonPanelSaveClose.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        WindowsUIButtonPanelSaveClose.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)
        'Me.GroupBoxFileActions.Font = GetFont("Small", Me.ScaleFactor)

    End Sub
    Sub SetInitialSizes()

        If Screen.PrimaryScreen.Bounds.Width < 900 Then
            Me.Width = Screen.PrimaryScreen.Bounds.Width * 0.9
            Me.Height = Screen.PrimaryScreen.Bounds.Height * 0.9
        ElseIf Screen.PrimaryScreen.Bounds.Width < 1190 Then
            Me.Width = Screen.PrimaryScreen.Bounds.Width * 0.8
            Me.Height = Screen.PrimaryScreen.Bounds.Height * 0.8
        Else
            Me.Width = Screen.PrimaryScreen.Bounds.Width * 0.7
            Me.Height = Screen.PrimaryScreen.Bounds.Height * 0.7
        End If

        ResizeFonts()
        ResizeControls()

    End Sub
    Sub ResizeControls()

        Dim SetWidth As Integer = Me.Width * 0.17
        ScaleUnits = Me.Width * 0.007

        PictureBoxAbovoLogo.Top = ScaleUnits
        PictureBoxAbovoLogo.Left = ScaleUnits
        PictureBoxAbovoLogo.Width = SetWidth
        PictureBoxAbovoLogo.Height = CInt(PictureBoxAbovoLogo.Width * 0.483)

        DockPanelSettings.Width = SetWidth
        WindowsUIButtonPanelExitHelp.Left = ScaleUnits
        GroupBoxProgramDetails.Width = SetWidth
        XtraTabControlMainNavigator.Top = ScaleUnits
        XtraTabControlMainNavigator.Left = PictureBoxAbovoLogo.Right + ScaleUnits
        XtraTabControlMainNavigator.Width = Me.Width - SetWidth - (5 * ScaleUnits) - hideContainerRight.Width
        XtraTabControlMainNavigator.Height = Me.Height - (6 * ScaleUnits)
        XtraTabPageMainHABP.Height = XtraTabControlMainNavigator.PageClientBounds.Height
        WindowsUIButtonPanelOpenCompare.Left = ScaleUnits
        WindowsUIButtonPanelOpenCompare.Top = 3 * ScaleUnits
        XtraTabControlModels.Left = ScaleUnits
        XtraTabControlModels.Top = WindowsUIButtonPanelOpenCompare.Bottom + ScaleUnits
        XtraTabControlModels.Width = XtraTabControlMainNavigator.Width - (2 * ScaleUnits)
        XtraTabControlModels.Height = XtraTabPageMainHABP.Height - WindowsUIButtonPanelExitHelp.Height - (4 * ScaleUnits)
        WindowsUIButtonPanelExitHelp.Top = XtraTabControlMainNavigator.Bottom - WindowsUIButtonPanelExitHelp.Height
        WindowsUIButtonPanelExitHelp.Width = SetWidth
        WindowsUIButtonPanelExitHelp.Left = ScaleUnits
        'WebBrowserBPInfo.Top = (2 * ScaleUnits)
        'WebBrowserBPInfo.Width = GroupBoxFileActions.Width - WindowsUIButtonPanelSaveClose.Width - (3 * ScaleUnits)
        'WebBrowserBPInfo.Height = GroupBoxFileActions.Height - WindowsUIButtonPanelBPActions.Height - (4 * ScaleUnits)
        'WindowsUIButtonPanelSaveClose.Left = WebBrowserBPInfo.Right + ScaleUnits
        'WindowsUIButtonPanelSaveClose.Top = WebBrowserBPInfo.Top
        'WindowsUIButtonPanelSaveClose.Height = GroupBoxFileActions.Height
        WindowsUIButtonPanelOpenCompare.Width = XtraTabControlMainNavigator.PageClientBounds.Width - (2 * ScaleUnits)
        'WindowsUIButtonPanelBPActions.Top = WebBrowserBPInfo.Bottom + ScaleUnits
        'WindowsUIButtonPanelBPActions.Width = WebBrowserBPInfo.Width
        GroupBoxProgramDetails.Top = PictureBoxAbovoLogo.Bottom + ScaleUnits
        GroupBoxProgramDetails.Left = ScaleUnits
        GroupBoxProgramDetails.Height = WindowsUIButtonPanelExitHelp.Top - PictureBoxAbovoLogo.Bottom - (2 * ScaleUnits)

    End Sub
    Private Sub FormMainScreen_Load(sender As Object, e As EventArgs) Handles MyBase.Load
#If DEBUG Then
        If Not DesignMode AndAlso IO.File.Exists(DebugAutoOpenModelPath) Then
            BeginInvoke(New MethodInvoker(
                Sub() OpenModelProceedureBP(DebugAutoOpenModelPath)))
        End If
#End If
    End Sub
    Private Sub SetBrowserText()

        WebBrowserProgramDetails.DocumentText = "<html><body><B>" +
                                                "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits * 1.2) & "px'>abovo-summit version " & DecVersionNumber & " <br/></b>" +
                                                "</p>" +
                                                "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits) & "px'>© 2015-" +
                                                Year(Now()).ToString +
                                                " Abovo Business Services Limited.</p>" +
                                                "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits) & "px'><Support <a href='https://www.abovo-consult.co.uk'>www.abovo-consult.co.uk</a>" +
                                                "</p><p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits) & "px'><a href='mailto:support@abovo-consult.co.uk'>support@abovo-consult.co.uk</a><br>" +
                                                "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits) & "px'>Built using Microsoft&reg; Excel&reg; © " +
                                                "<a href='https://www.microsoft.com'>Microsoft</a> Inc. </p>" +
                                                "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits) & "px'>Portions © Developer Express Inc." +
                                                "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits) & "px'>Visit the <a href='goforum'>Abovo Forum</a>." +
                                                "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits) & "px'>View <a href='goforum'>System Log</a>." +
                                                "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits) & "px'>Portions © Developer Express Inc." +
                                                "</body></html>"

    End Sub

    Private Sub ButtonExit_Click(sender As Object, e As EventArgs)

        CloseApplication()

    End Sub

#Region "ApplicationFormEvents"
    Sub CloseApplication()

        Me.Close()

    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)

        If Not ModelsClosedForShutdown Then
            ModelsClosedForShutdown = CloseAllModelsFromFMS(Me)
            If Not ModelsClosedForShutdown Then
                e.Cancel = True
                Return
            End If
        End If

        MyBase.OnFormClosing(e)

    End Sub
    Sub AddEvHandler()

    End Sub
    Private Sub DivertURl(sender As Object, e As WebBrowserNavigatingEventArgs)

        Process.Start(e.Url.ToString, "default")
        e.Cancel = True

    End Sub

    Private Sub ButtonCompareBPs_Click(sender As Object, e As EventArgs)

        AddEvHandler()

    End Sub

    Private Sub WebBrowserProgramDetails_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs)

        If e.Url.ToString = "about:blank" Then Exit Sub
        Process.Start(e.Url.ToString)
        e.Cancel = True

    End Sub

#End Region
#Region "ApplicationProcessEvents"
    Sub OpenModelProceedureBP(Optional ByVal AutoFileToOpen As String = Nothing)

        Dim FileToOpen As String = AutoFileToOpen
        Dim OpenedModelID As Integer = -1

        XtraOpenFileDialogMainScreen.Filter = "Abovo Models|*.xlsb;*.abp;*.adsa"

        If String.IsNullOrWhiteSpace(FileToOpen) OrElse
           String.Equals(FileToOpen, "None", StringComparison.OrdinalIgnoreCase) Then

            If XtraOpenFileDialogMainScreen.ShowDialog() <> DialogResult.OK Then Return
            FileToOpen = XtraOpenFileDialogMainScreen.FileName

        End If

        FileToOpen = IO.Path.GetFullPath(FileToOpen)

        If Not IO.File.Exists(FileToOpen) Then
            MessageBox.Show(Me,
                            "The selected model does not exist:" & Environment.NewLine & FileToOpen,
                            "Open Abovo Model",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
            Return
        End If

        If FileManager.IsFileOpen(FileToOpen) Then
            MessageBox.Show(Me,
                            "This model is already open:" & Environment.NewLine & FileToOpen,
                            "Open Abovo Model",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Return
        End If

        Me.Cursor = Cursors.WaitCursor

        Try
            MyFileInfos = New IO.FileInfo(FileToOpen)
            ProgressPanel("Loading " & FileToOpen & "...", "Abovo BP", 0)

            Dim FileOpenResult As AbovoTransaction =
                FileManager.OpenModel(FileToOpen, MyFileInfos)

            If FileOpenResult.BError Then
                Dim ErrorMessage As String = FileOpenResult.StringReturn
                If String.IsNullOrWhiteSpace(ErrorMessage) Then
                    ErrorMessage = FileOpenResult.StrResponseMessage
                End If
                If String.IsNullOrWhiteSpace(ErrorMessage) Then
                    ErrorMessage = "The model could not be opened."
                End If

                ProgressPanel("Error opening file: " & ErrorMessage, "Abovo BP", 2)
                MessageBox.Show(Me,
                                ErrorMessage,
                                "Error Opening File",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)
                Return
            End If

            OpenedModelID = FileOpenResult.IntegerReturn
            ActiveModel = OpenedModelID

            Select Case FileOpenResult.StringReturn
                Case "AbovoBP"
                    PopulateControlsFileBP(ActiveModel)
                    PostLoadActionsBP(ActiveModel)
                Case "AbovoDSA"
                    'Reserved for the DSA-specific interface.
                Case Else
                    Throw New InvalidOperationException("The loaded model type is not recognised.")
            End Select

            ProgressPanel("Finished.", "Abovo BP", 2)

        Catch ex As Exception
            If OpenedModelID >= 0 Then
                Try
                    If ExcelModels IsNot Nothing AndAlso
                       OpenedModelID < ExcelModels.Length AndAlso
                       ExcelModels(OpenedModelID) IsNot Nothing Then
                        FileManager.CloseModel(OpenedModelID)
                    End If
                    RemoveModel(OpenedModelID)
                Catch
                    'Keep the original open failure as the user-facing error.
                End Try
            End If

            MessageBox.Show(Me,
                            "The model could not be opened." & Environment.NewLine & ex.Message,
                            "Error Opening File",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

#End Region
#Region "InterfaceEvents"
    Function PopulateControlsFileBP(ModelID As Integer) As Integer

        FileInstanceIndex += 1
        ReDim Preserve FileInstances(FileInstanceIndex)
        FileInstances(FileInstanceIndex) = New FileInstanceInterface(ModelID)


        Dim XtraTabPageMainHABP As New DevExpress.XtraTab.XtraTabPage
        FileInstances(FileInstanceIndex).Dock = DockStyle.Fill
        XtraTabPageMainHABP.Controls.Add(FileInstances(FileInstanceIndex))
        XtraTabPageMainHABP.Tag = ModelID
        XtraTabControlModels.TabPages.Add(XtraTabPageMainHABP)
        XtraTabControlModels.SelectedTabPage = XtraTabPageMainHABP
        XtraTabPageMainHABP.Name = "TabPageModelNo_" & ModelID
        FileInstances(FileInstanceIndex).PopulateFileInfo()
        XtraTabPageMainHABP.Appearance.Header.ForeColor = ExcelModels(ModelID).ColourSwatch
        XtraTabPageMainHABP.Text = "(" & ModelID + 1 & ") " & "HABP " & FileInstances(FileInstanceIndex).MyCompanyName
        XtraTabPageMainHABP.Tooltip = FileInstances(FileInstanceIndex).MyFilePath
        HideFirstTab()
        Return FileInstanceIndex

    End Function
    Public Sub PostLoadActionsBP(ModelID As Integer)

        If ExcelModels Is Nothing OrElse
           ModelID < 0 OrElse
           ModelID >= ExcelModels.Length OrElse
           ExcelModels(ModelID) Is Nothing Then Return

        Dim ThisBP As IWorkbook = ExcelModels(ModelID).WB
        Dim StressModeName As DefinedName =
            ThisBP.DefinedNames.GetDefinedName("StressTestMode")

        If StressModeName Is Nothing OrElse StressModeName.Range Is Nothing Then Return

        If String.Equals(StressModeName.Range(0, 0).Value.TextValue,
                         "Y",
                         StringComparison.OrdinalIgnoreCase) Then

            If MsgBox("WARNING: ths model is in Stress Test mode." & Chr(13) & "Do you wish to switch to Business Plan mode?", Buttons:=vbYesNo + vbQuestion) = vbYes Then

                StressModeName.Range(0, 0).SetValueFromText("N")

                Dim ModeName As DefinedName =
                    ThisBP.DefinedNames.GetDefinedName("Mode")

                If ModeName Is Nothing Then
                    ThisBP.DefinedNames.Add("Mode", """Business Plan""")
                Else
                    ModeName.RefersTo = """Business Plan"""
                End If

            End If

        End If

    End Sub

    Public Sub RemoveModel(ModelID As Integer)

        Dim XtraTabPageToRemove As DevExpress.XtraTab.XtraTabPage = Nothing

        For Each XTP As DevExpress.XtraTab.XtraTabPage In XtraTabControlModels.TabPages

            If XTP.Tag = ModelID Then
                XtraTabPageToRemove = XTP
                Exit For
            End If

        Next

        If Not IsNothing(XtraTabPageToRemove) Then
            XtraTabControlModels.TabPages.Remove(XtraTabPageToRemove)
            HideFirstTab()
        End If

    End Sub
    Sub HideFirstTab()

        If XtraTabControlModels.TabPages.Count > 1 Then

            XtraTabPageBlank.PageVisible = False

        Else

            XtraTabPageBlank.PageVisible = True

        End If

    End Sub
#End Region
#Region "General Services"
    Private Sub ProgressPanel(strDisplayText As String, strCaption As String, Optional ByVal intStage As Integer = 1)

        Select Case intStage

            Case 0

                SplashScreenManagerMainForm.ShowWaitForm()
                SplashScreenManagerMainForm.SetWaitFormCaption(strCaption)
                SplashScreenManagerMainForm.SetWaitFormDescription(strDisplayText)

            Case 1

                SplashScreenManagerMainForm.SetWaitFormDescription(strDisplayText)

            Case Else

                SplashScreenManagerMainForm.SetWaitFormCaption("Finished")
                SplashScreenManagerMainForm.SetWaitFormDescription(strDisplayText)
                SplashScreenManagerMainForm.CloseWaitForm()

        End Select

    End Sub


    Private Sub WindowsUIButtonPanelOpenCompare_ButtonClick(sender As Object, e As ButtonEventArgs) Handles WindowsUIButtonPanelOpenCompare.ButtonClick
        Dim ButSender As WindowsUIButton = TryCast(e.Button, DevExpress.XtraBars.Docking2010.WindowsUIButton)
        If ButSender Is Nothing Then
            Return
        End If
        Dim tag As String = ButSender.Tag.ToString()
        Select Case tag
            Case "OpenBP"
                ' OpenBusinessPlan

                OpenModelProceedureBP()

            Case "CreateNewBP"
                NewBP()

            Case "CompareBPs"

                MsgBox("Awaiting DevExpress Fix")
                Return

        End Select
    End Sub

    Private Sub NewBP()
        ' TODO: Revisit and redesign the Create New Business Plan workflow.
    End Sub


    Private Sub WindowsUIButtonPanelExitHelp_ButtonClick(sender As Object, e As ButtonEventArgs) Handles WindowsUIButtonPanelExitHelp.ButtonClick
        Dim ButSender As WindowsUIButton = TryCast(e.Button, DevExpress.XtraBars.Docking2010.WindowsUIButton)
        If ButSender Is Nothing Then
            Return
        End If
        Dim tag As String = CType(e.Button, WindowsUIButton).Tag.ToString()

        Select Case tag

            Case "CloseApp"

                ' Exit
                If CloseAllModelsFromFMS(Me) Then

                    ModelsClosedForShutdown = True
                    CloseApplication()

                End If


            Case "GetHelp"

                ' Navigate to page B 
                Process.Start("https://www.abovo-consult.co.uk/summit/help")

        End Select
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs)

        AbovoBP.IRVs.Initialise()

    End Sub

    Private Sub FormMainScreen_ResizeEnd(sender As Object, e As EventArgs) Handles MyBase.ResizeEnd

        ResizeFonts()
        ResizeControls()
        PictureBoxAbovoLogo.Height = CInt(PictureBoxAbovoLogo.Width * 0.483)
        GroupBoxProgramDetails.Top = PictureBoxAbovoLogo.Bottom + ScaleUnits


    End Sub
    Private Sub FormMainScreen_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize

        If Me.WindowState = FormWindowState.Maximized Then

            AmMaximised = True
            ResizeFonts()
            ResizeControls()

        End If

        If Not Me.WindowState = FormWindowState.Maximized Then

            If AmMaximised Then

                AmMaximised = False
                ' ResizeFonts()
                ResizeControls()

            End If

        End If

    End Sub

    Private Sub hideContainerRight_Click(sender As Object, e As EventArgs) Handles hideContainerRight.Click

    End Sub

    Private Sub AccordionControlOptHist_Click(sender As Object, e As EventArgs)

    End Sub

#End Region

End Class
