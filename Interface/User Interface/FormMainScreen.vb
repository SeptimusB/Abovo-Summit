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
    Private Const DebugAutoOpenModelPath As String = "C:\Sandbox\Deprecated\Test BP v26_0001 - FormGenRemoved - PopInSummit - MenuFixed.xlsb"
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
    Private DsaModelsTabControl As DevExpress.XtraTab.XtraTabControl
    Private DsaBlankPage As DevExpress.XtraTab.XtraTabPage
    Private ModelsClosedForShutdown As Boolean

    Public Sub New()





        'SplashScreenManagerStartup.ShowWaitForm()
        DevExpress.XtraEditors.WindowsFormsSettings.SmartMouseWheelProcessing = False

        InitializeComponent()
        InitialiseDSAWorkspace()

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

        ScaleFactor = GetDisplayScale(Me)

        Me.GroupBoxProgramDetails.Font = GetDisplayFont("Small", Me)

        Me.hideContainerRight.Font = GetDisplayFont("Small", Me)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.ActiveTab.Font = GetDisplayFont("Medium", Me)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.HidePanelButton.Font = GetDisplayFont("Medium", Me)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.HidePanelButtonActive.Font = GetDisplayFont("Medium", Me)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaption.Font = GetDisplayFont("Medium", Me)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaptionActive.Font = GetDisplayFont("Medium", Me)

        Me.BarAndDockingControllerMainScreen.AppearancesBar.Dock.Font = GetDisplayFont("Medium", Me)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.HidePanelButton.Font = GetDisplayFont("Medium", Me)
        Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaption.Font = GetDisplayFont("Medium", Me)
        Me.XtraTabControlMainNavigator.AppearancePage.HeaderActive.Font = GetDisplayFont("Medium", Me, False, True)
        Me.XtraTabControlMainNavigator.AppearancePage.Header.Font = GetDisplayFont("Medium", Me)
        Me.XtraTabControlMainNavigator.Appearance.Font = GetDisplayFont("Medium", Me)
        Me.XtraTabControlMainNavigator.AppearancePage.HeaderHotTracked.Font = GetDisplayFont("Medium", Me)

        WindowsUIButtonPanelExitHelp.Font = GetDisplayFont("Small", Me)
        Me.WindowsUIButtonPanelBPActions.Font = GetDisplayFont("Small", Me)

        WindowsUIButtonPanelOpenCompare.AppearanceButton.Normal.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelOpenCompare.AppearanceButton.Hovered.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelOpenCompare.AppearanceButton.Pressed.Font = GetDisplayFont("Small", Me)

        WindowsUIButtonPanelBPActions.AppearanceButton.Normal.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelBPActions.AppearanceButton.Hovered.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelBPActions.AppearanceButton.Pressed.Font = GetDisplayFont("Small", Me)

        WindowsUIButtonPanelExitHelp.AppearanceButton.Normal.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelExitHelp.AppearanceButton.Hovered.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelExitHelp.AppearanceButton.Pressed.Font = GetDisplayFont("Small", Me)

        WindowsUIButtonPanelSaveClose.AppearanceButton.Normal.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelSaveClose.AppearanceButton.Hovered.Font = GetDisplayFont("Small", Me)
        WindowsUIButtonPanelSaveClose.AppearanceButton.Pressed.Font = GetDisplayFont("Small", Me)
        'Me.GroupBoxFileActions.Font = GetFont("Small", Me.ScaleFactor)

    End Sub
    Sub SetInitialSizes()

        Dim AvailableArea As Rectangle = Screen.FromPoint(Cursor.Position).WorkingArea
        If AvailableArea.Width < 900 Then
            Me.Width = CInt(AvailableArea.Width * 0.9)
            Me.Height = CInt(AvailableArea.Height * 0.9)
        ElseIf AvailableArea.Width < 1190 Then
            Me.Width = CInt(AvailableArea.Width * 0.8)
            Me.Height = CInt(AvailableArea.Height * 0.8)
        Else
            Me.Width = CInt(AvailableArea.Width * 0.7)
            Me.Height = CInt(AvailableArea.Height * 0.7)
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
        Dim benchmarkArgs As String() = Environment.GetCommandLineArgs()
        If benchmarkArgs.Length >= 4 AndAlso benchmarkArgs.Length <= 5 AndAlso
           String.Equals(benchmarkArgs(1), "--benchmark-bitness", StringComparison.OrdinalIgnoreCase) Then
            BeginInvoke(New MethodInvoker(
                Sub() RunBitnessBenchmark(benchmarkArgs(2), benchmarkArgs(3),
                    If(benchmarkArgs.Length = 5, benchmarkArgs(4), "TransRents"))))
            Return
        End If
#If DEBUG Then
        Dim DebugAutoOpenModelPath As String = ResolveDebugAutoOpenModelPath()
        If Not DesignMode AndAlso Not String.IsNullOrWhiteSpace(DebugAutoOpenModelPath) Then
            BeginInvoke(New MethodInvoker(
                Sub() OpenModelProceedureBP(DebugAutoOpenModelPath)))
        End If
#End If
    End Sub

    Private Sub RunBitnessBenchmark(workbookPath As String, resultPath As String,
                                    editRangeName As String)
        Dim modelID As Integer = -1
        Try
            Using output As New System.IO.StreamWriter(resultPath, False, System.Text.Encoding.UTF8)
                output.WriteLine("bitness" & vbTab & "phase" & vbTab & "elapsed_ms" & vbTab &
                                 "private_bytes" & vbTab & "peak_working_set_bytes" & vbTab & "outcome")
                Dim timer As New System.Diagnostics.Stopwatch()
                timer.Start()
                Dim result As AbovoTransaction =
                    FileManager.OpenModel(workbookPath, New System.IO.FileInfo(workbookPath))
                timer.Stop()
                If result.BError Then Throw New InvalidOperationException(result.StringReturn)
                modelID = result.IntegerReturn
                WriteBitnessBenchmarkLine(output, "open", timer.ElapsedMilliseconds, "ok")

                WriteBitnessBenchmarkPhase(output, "staged_first",
                    Sub() ExcelModels(modelID).WBCalcEngine.CalcFile(1))
                WriteBitnessBenchmarkPhase(output, "staged_repeat",
                    Sub() ExcelModels(modelID).WBCalcEngine.CalcFile(1))
                WriteBitnessBenchmarkPhase(output, "recursive_dependencies",
                    Sub() ExcelModels(modelID).WBCalcEngine.CalculateDependencySensitiveFile(
                        "Bitness benchmark", True))
                WriteBitnessBenchmarkPhase(output, "full_rebuild",
                    Sub() ExcelModels(modelID).WBCalcEngine.CalcFile(3))
                RunBitnessEditBenchmark(output, modelID, editRangeName)
            End Using
        Catch ex As Exception
            Try
                System.IO.File.AppendAllText(resultPath,
                    IntPtr.Size * 8 & vbTab & "error" & vbTab & "0" & vbTab & "0" & vbTab &
                    "0" & vbTab & ex.ToString().Replace(vbCr, " ").Replace(vbLf, " ") &
                    Environment.NewLine)
            Catch
            End Try
        Finally
            If modelID >= 0 Then FileManager.CloseModel(modelID)
            ModelsClosedForShutdown = True
            Close()
        End Try
    End Sub

    Private Shared Sub RunBitnessEditBenchmark(output As System.IO.TextWriter,
                                                modelID As Integer, rangeName As String)
        Dim model = ExcelModels(modelID)
        Dim name As DefinedName = model.WB.DefinedNames.GetDefinedName(rangeName)
        If name Is Nothing Then
            name = model.WB.Worksheets("Rent Assumptions").DefinedNames.GetDefinedName(rangeName)
        End If
        If name Is Nothing OrElse name.Range Is Nothing Then
            WriteBitnessBenchmarkLine(output, "edit", 0, "skipped: name not found")
            Return
        End If

        Dim input As Cell = Nothing
        Dim cells As CellRange = name.Range
        For row As Integer = 0 To cells.RowCount - 1
            For col As Integer = 0 To cells.ColumnCount - 1
                Dim candidate As Cell = cells(row, col)
                If Not candidate.Protection.Locked AndAlso
                   Not candidate.HasFormula AndAlso candidate.Value.IsNumeric Then
                    input = candidate
                    Exit For
                End If
            Next
            If input IsNot Nothing Then Exit For
        Next
        If input Is Nothing Then
            WriteBitnessBenchmarkLine(output, "edit", 0, "skipped: no numeric unlocked input")
            Return
        End If

        Dim original As Double = input.Value.NumericValue
        Dim change As New DataChangeEvent With {
            .ModelID = modelID,
            .Description = "Architecture benchmark edit",
            .WSName = input.Worksheet.Name,
            .CellAddress = input.GetReferenceA1(),
            .OriginalValue = original,
            .ChangedValue = original + 1.0R,
            .DataFormat = "SM",
            .TimeStamp = DateTime.Now,
            .UserName = Environment.UserName
        }
        Dim timer As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
        Dim result As AbovoTransaction = model.ChangeManager.ProcessChange(change)
        timer.Stop()
        If result.BError Then Throw New InvalidOperationException(result.StringReturn)
        If Math.Abs(input.Value.NumericValue - original) < 0.0000001R Then
            Throw New InvalidOperationException("The benchmark edit did not change the selected cell.")
        End If
        WriteBitnessBenchmarkLine(output, "edit", timer.ElapsedMilliseconds,
                                  "ok:" & input.Worksheet.Name & "!" & input.GetReferenceA1())

        timer.Restart()
        result = model.ChangeManager.Undo()
        timer.Stop()
        If result.BError Then Throw New InvalidOperationException(result.StringReturn)
        If Math.Abs(input.Value.NumericValue - original) > 0.0000001R Then
            Throw New InvalidOperationException("Undo did not restore the benchmark input.")
        End If
        WriteBitnessBenchmarkLine(output, "undo", timer.ElapsedMilliseconds, "ok")
    End Sub

    Private Shared Sub WriteBitnessBenchmarkPhase(
        output As System.IO.TextWriter, phase As String, operation As Action)
        Dim timer As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
        operation()
        timer.Stop()
        WriteBitnessBenchmarkLine(output, phase, timer.ElapsedMilliseconds, "ok")
    End Sub

    Private Shared Sub WriteBitnessBenchmarkLine(
        output As System.IO.TextWriter, phase As String, elapsedMs As Long, outcome As String)
        Using currentProcess As System.Diagnostics.Process =
            System.Diagnostics.Process.GetCurrentProcess()
            currentProcess.Refresh()
            output.WriteLine(String.Join(vbTab, {
                (IntPtr.Size * 8).ToString(),
                phase,
                elapsedMs.ToString(),
                currentProcess.PrivateMemorySize64.ToString(),
                currentProcess.PeakWorkingSet64.ToString(),
                outcome
            }))
            output.Flush()
        End Using
    End Sub
#If DEBUG Then
    Private Shared Function ResolveDebugAutoOpenModelPath() As String
        Return If(IO.File.Exists(DebugAutoOpenModelPath),
                  DebugAutoOpenModelPath,
                  String.Empty)

    End Function
#End If
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
    Sub OpenModelProceedureBP(
        Optional ByVal AutoFileToOpen As String = Nothing,
        Optional ByVal ExpectedModelType As String = Nothing)

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

                If SplashScreenManagerMainForm.IsSplashFormVisible Then
                    SplashScreenManagerMainForm.CloseWaitForm()
                End If
                MessageBox.Show(Me,
                                ErrorMessage,
                                "Error Opening File",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)
                Return
            End If

            OpenedModelID = FileOpenResult.IntegerReturn
            ActiveModel = OpenedModelID

            If Not String.IsNullOrWhiteSpace(ExpectedModelType) AndAlso
               Not String.Equals(
                FileOpenResult.StringReturn,
                ExpectedModelType,
                StringComparison.OrdinalIgnoreCase) Then

                Throw New InvalidOperationException(
                    "The selected file is a " & FileOpenResult.StringReturn &
                    " model, not the requested " & ExpectedModelType & " model.")
            End If

            Select Case FileOpenResult.StringReturn
                Case "AbovoBP"
                    PopulateControlsFileBP(ActiveModel)
                    PostLoadActionsBP(ActiveModel)
                Case "AbovoDSA"
                    PopulateControlsFileDSA(ActiveModel)
                Case Else
                    Throw New InvalidOperationException("The loaded model type is not recognised.")
            End Select

            ProgressPanel("Model ready.", "Abovo BP", 2)

        Catch ex As Exception
            If SplashScreenManagerMainForm.IsSplashFormVisible Then
                SplashScreenManagerMainForm.CloseWaitForm()
            End If
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

    Private Sub InitialiseDSAWorkspace()

        XtraTabPageEvolveDSA.Appearance.PageClient.BackColor = Color.White
        XtraTabPageEvolveDSA.Controls.Clear()

        SimpleButtonTest.Text = "Open DSA"
        SimpleButtonTest.Location = New Point(16, 12)
        SimpleButtonTest.Size = New Size(180, 48)
        SimpleButtonTest.Anchor = AnchorStyles.Top Or AnchorStyles.Left

        DsaModelsTabControl = New DevExpress.XtraTab.XtraTabControl With {
            .Name = "XtraTabControlDSAModels",
            .Location = New Point(12, 72),
            .Size = New Size(
                Math.Max(200, XtraTabPageEvolveDSA.ClientSize.Width - 24),
                Math.Max(200, XtraTabPageEvolveDSA.ClientSize.Height - 84)),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or
                      AnchorStyles.Left Or AnchorStyles.Right
        }

        DsaBlankPage = New DevExpress.XtraTab.XtraTabPage With {
            .Name = "XtraTabPageDSABlank",
            .Text = "No DSA models open"
        }

        DsaModelsTabControl.TabPages.Add(DsaBlankPage)
        XtraTabPageEvolveDSA.Controls.Add(DsaModelsTabControl)
        XtraTabPageEvolveDSA.Controls.Add(SimpleButtonTest)
        SimpleButtonTest.BringToFront()

    End Sub

    Private Sub OpenDSAProcedure()

        XtraOpenFileDialogMainScreen.Filter =
            "Development Scheme Appraisals|*.xlsb;*.adsa"

        If XtraOpenFileDialogMainScreen.ShowDialog() <> DialogResult.OK Then Return

        OpenModelProceedureBP(
            XtraOpenFileDialogMainScreen.FileName,
            "AbovoDSA")

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
        XtraTabControlMainNavigator.SelectedTabPage = Me.XtraTabPageMainHABP
        XtraTabPageMainHABP.Name = "TabPageModelNo_" & ModelID
        FileInstances(FileInstanceIndex).PopulateFileInfo()
        XtraTabPageMainHABP.Appearance.Header.ForeColor = ExcelModels(ModelID).ColourSwatch
        XtraTabPageMainHABP.Text = "(" & ModelID + 1 & ") " & "HABP " & FileInstances(FileInstanceIndex).MyCompanyName
        XtraTabPageMainHABP.Tooltip = FileInstances(FileInstanceIndex).MyFilePath
        HideFirstTab()
        Return FileInstanceIndex

    End Function

    Function PopulateControlsFileDSA(ModelID As Integer) As Integer

        FileInstanceIndex += 1
        ReDim Preserve FileInstances(FileInstanceIndex)
        FileInstances(FileInstanceIndex) = New FileInstanceInterface(ModelID)

        Dim DsaPage As New DevExpress.XtraTab.XtraTabPage
        FileInstances(FileInstanceIndex).Dock = DockStyle.Fill
        DsaPage.Controls.Add(FileInstances(FileInstanceIndex))
        DsaPage.Tag = ModelID
        DsaModelsTabControl.TabPages.Add(DsaPage)
        DsaModelsTabControl.SelectedTabPage = DsaPage
        XtraTabControlMainNavigator.SelectedTabPage = XtraTabPageEvolveDSA
        DsaPage.Name = "TabPageModelNo_" & ModelID
        FileInstances(FileInstanceIndex).PopulateFileInfo()
        DsaPage.Appearance.Header.ForeColor =
            ExcelModels(ModelID).ColourSwatch
        DsaPage.Text =
            "(" & ModelID + 1 & ") DSA " &
            FileInstances(FileInstanceIndex).MyCompanyName
        DsaPage.Tooltip = FileInstances(FileInstanceIndex).MyFilePath
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

        If RemoveModelPage(XtraTabControlModels, ModelID) OrElse
           RemoveModelPage(DsaModelsTabControl, ModelID) Then
            HideFirstTab()
        End If

    End Sub

    Private Function RemoveModelPage(
        ByVal ModelTabs As DevExpress.XtraTab.XtraTabControl,
        ByVal ModelID As Integer) As Boolean

        If ModelTabs Is Nothing Then Return False

        Dim XtraTabPageToRemove As DevExpress.XtraTab.XtraTabPage = Nothing

        For Each XTP As DevExpress.XtraTab.XtraTabPage In ModelTabs.TabPages

            If XTP.Tag = ModelID Then
                XtraTabPageToRemove = XTP
                Exit For
            End If

        Next

        If Not IsNothing(XtraTabPageToRemove) Then
            ModelTabs.TabPages.Remove(XtraTabPageToRemove)
            Return True
        End If

        Return False

    End Function

    Sub HideFirstTab()

        If XtraTabControlModels.TabPages.Count > 1 Then

            XtraTabPageBlank.PageVisible = False

        Else

            XtraTabPageBlank.PageVisible = True

        End If

        If DsaModelsTabControl IsNot Nothing AndAlso
           DsaBlankPage IsNot Nothing Then

            DsaBlankPage.PageVisible =
                DsaModelsTabControl.TabPages.Count <= 1
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

                SplashScreenManagerMainForm.SetWaitFormCaption("Complete")
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

                Abovo.HelpManager.ShowHelpHome(Me)

            Case "Options"

                Abovo.PresentationScaleManager.ShowOptions(Me)

        End Select
    End Sub

    Private Sub SimpleButtonTest_Click(
        sender As Object,
        e As EventArgs) Handles SimpleButtonTest.Click

        OpenDSAProcedure()

    End Sub

    Private Sub FormMainScreen_ResizeEnd(sender As Object, e As EventArgs) Handles MyBase.ResizeEnd

        ResizeControls()
        PictureBoxAbovoLogo.Height = CInt(PictureBoxAbovoLogo.Width * 0.483)
        GroupBoxProgramDetails.Top = PictureBoxAbovoLogo.Bottom + ScaleUnits


    End Sub
    Private Sub FormMainScreen_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize

        If Me.WindowState = FormWindowState.Maximized Then

            AmMaximised = True
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
