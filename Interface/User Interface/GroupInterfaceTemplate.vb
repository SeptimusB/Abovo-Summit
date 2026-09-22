Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.LogDebugDev
Imports Abovo.PresentationManager

Imports DevExpress.CodeParser
Imports DevExpress.Data.Async.Helpers
Imports DevExpress.DataAccess.DataFederation
Imports DevExpress.DataAccess.Wizard.Model
Imports DevExpress.Skins
Imports DevExpress.Skins.XtraForm
Imports DevExpress.Utils
Imports DevExpress.Utils.Drawing
Imports DevExpress.XtraBars.Docking.Helpers
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.XtraBars.Docking2010.Views
Imports DevExpress.XtraBars.Docking2010.Views.Tabbed
Imports DevExpress.XtraBars.Navigation
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Filtering
Imports DevExpress.XtraLayout
Imports DevExpress.XtraPrinting
Imports DevExpress.XtraRichEdit.Import.Html
Imports Microsoft.VisualBasic.Devices
Imports System.Diagnostics
Imports System.Net

Public Class GroupInterfaceTemplate

    Inherits DevExpress.XtraEditors.XtraForm

    Public GSID As Integer
    Private Const CombinedGroupID As Integer = -1

    Public ReadOnly Property IsCombined As Boolean
        Get
            Return GSID = CombinedGroupID
        End Get
    End Property

    Private LrgFontSize As Integer
    Private MediumFontSize As Integer
    Private SmallFontSize As Integer
    Private ScaleFactor As Single
    Private ScaleUnits As Single
    Private LastHTMLFontSize As Integer
    Private AmMaximised As Boolean = False
    Private HTMLFontSize As Integer
    Private AccGroupCount As Integer
    Private AccGroups() As AccordionControlElement
    Private LastGroup As String
    Private OpenGroup As Boolean
    Private GroupElements() As AccordionControlElement
    Private GroupElementsCount As Integer
    Private DataInterfaces() As DataInterfaceTemplate
    Private ActiveInterface As Object
    Private DataInterfaceCount As Integer
    Private MyName As String
    Private MyModelID As Integer
    Private GITWindowState As FormWindowState
    Public ParentModelSSViewer As MainModelViewer
    Public ActiveWorksheet As String
    Private MyColourSwatch As Color
    Private OpenGroupID As Integer
    Private InterfaceMode As String
    Private SidebarMessageView As SystemMessageView
    Private SidebarHistoryView As InterfaceHistoryView
    Private SidebarHistoryContainer As AccordionContentContainer
    Private SidebarHistoryElement As AccordionControlElement
    Private SidebarRefreshTimer As Timer
    Private SidebarEventsAttached As Boolean
    Private PanelsHidden As Boolean 'Compacted contents; the edge strips remain available.
    Private ReadOnly SavedPanelStates As New Dictionary(Of DevExpress.XtraBars.Docking.DockPanel, Tuple(Of DevExpress.XtraBars.Docking.DockVisibility, Size, Size))
    Private ReadOnly PanelButtons As New List(Of WindowsUIButton)
    Private PresentationReady As Boolean
    Private PresentationResizeQueued As Boolean
    Private LastPresentationScale As Single = -1
    Private LastPanelLayoutWidth As Integer = -1
    Private LastPanelLayoutScale As Single = -1
    Private LastSidebarUserScale As Single = -1
    Private ReadOnly SidebarHtml As New Dictionary(Of WebBrowser, String)
    Private ReadOnly SidebarRenderedHtml As New Dictionary(Of WebBrowser, String)
    Private ReadOnly SidebarDocumentLoading As New HashSet(Of WebBrowser)
    Private ReadOnly SidebarDocumentUpdating As New HashSet(Of WebBrowser)

    Public Sub AttachPanelsButton(panel As WindowsUIButtonPanel)
        If panel.Buttons.OfType(Of WindowsUIButton)().Any(Function(b) Object.Equals(b.Tag, "TogglePanels")) Then Return
        Dim button As WindowsUIButton = PresentationLayout.CreatePanelsButton()
        PanelButtons.Add(button)
        panel.Buttons.Add(button)
        AddHandler panel.ButtonClick,
            Sub(sender, e)
                If Object.ReferenceEquals(e.Button, button) Then ToggleInterfacePanels()
            End Sub
        AddHandler panel.Disposed, Sub(sender, e) PanelButtons.Remove(button)
        UpdatePanelsButtons()
    End Sub

    Public Sub ToggleInterfacePanels()
        DockManagerAssumptions.BeginUpdate()
        Try
            If Not PanelsHidden Then
                SavedPanelStates.Clear()
                For Each panel In {DockPanelNewNavigator, DockPanelDetail}
                    SavedPanelStates.Add(panel, Tuple.Create(panel.Visibility, panel.Size, panel.OriginalSize))
                Next
                'Compact the contents, not the edge strips. Hiding the legacy
                'navigator panel destroys its auto-hide container and Click handler.
                PanelsHidden = True
                DockPanelNewNavigator.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden
                DockPanelDetail.Visibility = DevExpress.XtraBars.Docking.DockVisibility.AutoHide
            Else
                For Each entry In SavedPanelStates
                    entry.Key.OriginalSize = entry.Value.Item3
                    entry.Key.Visibility = entry.Value.Item1
                    entry.Key.Size = entry.Value.Item2
                Next
                PanelsHidden = False
            End If
            DockPanelNavigator.Visibility = DevExpress.XtraBars.Docking.DockVisibility.AutoHide
            'A previously hidden summary must also retain a way back in.
            If DockPanelDetail.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden Then
                DockPanelDetail.Visibility = DevExpress.XtraBars.Docking.DockVisibility.AutoHide
            End If
        Finally
            DockManagerAssumptions.EndUpdate()
        End Try
        ReconnectPanelStrips()
        DockPanelNavigator.HideImmediately()
        If DockPanelDetail.Visibility = DevExpress.XtraBars.Docking.DockVisibility.AutoHide Then
            DockPanelDetail.HideImmediately()
        End If
        UpdatePanelsButtons()
        If Not PanelsHidden Then ResizeControls()
        ResizeGIT()
    End Sub

    Private Sub ReconnectPanelStrips()
        'WithEvents follows any replacement container created by native docking.
        'The left strip normally survives; the right is recreated when unpinned.
        Dim navigatorStrip = DockPanelNavigator.ParentAutoHideContainer
        If navigatorStrip IsNot Nothing AndAlso Not Object.ReferenceEquals(hideContainerLeft, navigatorStrip) Then
            hideContainerLeft = navigatorStrip
        End If
        Dim summaryStrip = DockPanelDetail.ParentAutoHideContainer
        If summaryStrip IsNot Nothing AndAlso Not Object.ReferenceEquals(hideContainerRightDetail, summaryStrip) Then
            hideContainerRightDetail = summaryStrip
        End If
    End Sub

    Private Sub UpdatePanelsButtons()
        For Each button In PanelButtons
            PresentationLayout.UpdatePanelsButton(button, PanelsHidden)
        Next
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(SetModelID As Integer, SetGSID As Integer, SetInterfaceMode As String)

        MyColourSwatch = ExcelModels(SetModelID).ColourSwatch
        GSID = SetGSID
        AccGroupCount = -1
        InterfaceMode = SetInterfaceMode
        InitializeComponent()

        LastGroup = ""
        OpenGroup = False
        GroupElementsCount = -1
        ReDim DataInterfaces(-1)
        DataInterfaceCount = -1
        MyModelID = SetModelID
        'Me.LookAndFeel.UseDefaultLookAndFeel = False
        Me.BarManagerAssumptions.TransparentEditorsMode = True
        Me.BarAndDockingControllerAssumptions.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        MyColourSwatch = ExcelModels(SetModelID).ColourSwatch
        BarTopBar.Appearance.BackColor = Color.White
        BarHeaderItemDetail.Appearance.BackColor = Color.White

        Me.Text = ExcelModels(SetModelID).WBStructure.CompanyName
        AccordionControlNavigator.LookAndFeel.UseDefaultLookAndFeel = False

        MyName = If(IsCombined, "Combined",
                    ExcelModels(SetModelID).WBStructure.GroupStructures(GSID).GSName)
        Me.Text = ExcelModels(SetModelID).WBStructure.CompanyName & " / " & MyName & " Interface"
        DockPanelNewNavigator.Text = " " & MyName & " Navigator"
        DockPanelNavigator.Text = " " & MyName & " Navigator"
        Dim myTag As String = MyName & " Interface"

        SetInitialSizes()

        If IsCombined Then
            ApplyCombinedStructure(SetModelID)
        Else
            ApplyStructure(SetModelID, GSID)
        End If

        LoadDefaultInterface()
        InitialiseRightSidebar()
        RefreshSummaryData()
        PresentationReady = True
        LastPanelLayoutWidth = -1
        ApplyPresentationScale()

    End Sub
    Sub HideNewNavContainer(ByVal sender As Object, ByVal e As DevExpress.XtraBars.Docking2010.ButtonEventArgs) Handles DockPanelNewNavigator.CustomButtonClick

        If (e.Button Is DockPanelNewNavigator.CustomHeaderButtons(0)) Then DockPanelNewNavigator.Hide()

    End Sub
    Public Sub ShowModelSSViewer()

        If ActiveWorksheet Is Nothing Then Exit Sub

        If ParentModelSSViewer IsNot Nothing Then

            ExcelModels(MyModelID).WB.Worksheets.ActiveWorksheet = ExcelModels(MyModelID).WB.Worksheets.ActiveWorksheet(ActiveWorksheet)
            ParentModelSSViewer.ShowDialog()
            ParentModelSSViewer.BringToFront()

        End If

    End Sub
    Public Sub ShowHistoryViewer()

        If ActiveWorksheet Is Nothing Then Exit Sub

        If ParentModelSSViewer IsNot Nothing Then

            ExcelModels(MyModelID).WB.Worksheets.ActiveWorksheet = ExcelModels(MyModelID).WB.Worksheets.ActiveWorksheet(ActiveWorksheet)
            ParentModelSSViewer.ShowDialog()
            ParentModelSSViewer.BringToFront()

        End If

    End Sub
    Protected Overrides Function CreateFormBorderPainter() As DevExpress.Skins.XtraForm.FormPainter

        Return New CustomFormPainterGIT(Me, LookAndFeel)

    End Function
    Sub ApplyStructure(ModelID As Integer, SetGSID As Integer)

        'DefineMenus
        Dim CS As ChildStructure


        For Each CS In ExcelModels(ModelID).WBStructure.GroupStructures(SetGSID).ChildStructures

            Dim SetTag As New AbovoInterfaceTag

            If CS.SpecialElement Is Nothing Then

                SetTag.TargetID = CS.CSID
                SetTag.SpecialItem = False

            Else

                SetTag.TargetID = CS.CSID
                SetTag.SpecialItem = True
                SetTag.SpecialItemData = CS.SpecialElement

            End If

            If CS.IsMaster = "True" Then

                AddNavigatorItem(CS.NavigationText, SetTag, CS.IsMaster)

            Else

                AddNavigatorItem(CS.NavigationText, SetTag, CS.IsMaster, CS.NavigationGroupText)

            End If

        Next

        FinaliseNavigator()

    End Sub

    Private Sub ApplyCombinedStructure(ByVal modelID As Integer)
        'Keep the original three GroupStructure definitions and their local CSIDs.
        'Only the navigator is combined; every target retains its (GSID, CSID).
        Dim root As New AccordionControlElement With {
            .Name = "CombinedRoot", .Text = "Root", .HeaderVisible = False,
            .Expanded = True
        }

        For groupIndex As Integer = 0 To ExcelModels(modelID).WBStructure.GroupStructures.Count - 1
            Dim group As GroupStructure = ExcelModels(modelID).WBStructure.GroupStructures(groupIndex)
            If group Is Nothing OrElse group.ChildStructures Is Nothing Then Continue For

            Dim groupNode As New AccordionControlElement With {
                .Name = "CombinedGroup" & groupIndex.ToString(),
                .Text = group.GSName,
                .Style = DevExpress.XtraBars.Navigation.ElementStyle.Group
            }
            'Apply before the node is attached so the first paint is already black.
            SetCombinedGroupAppearance(groupNode.Appearance.Default)
            SetCombinedGroupAppearance(groupNode.Appearance.Normal)
            SetCombinedGroupAppearance(groupNode.Appearance.Hovered)
            SetCombinedGroupAppearance(groupNode.Appearance.Pressed)
            Dim currentSubgroup As AccordionControlElement = Nothing
            Dim currentSubgroupName As String = Nothing

            For Each child As ChildStructure In group.ChildStructures
                If child Is Nothing Then Continue For
                Dim childID As Integer
                If Not Integer.TryParse(child.CSID, childID) Then Continue For

                Dim target As New AbovoInterfaceTag With {
                    .TargetGroupID = groupIndex,
                    .TargetID = childID,
                    .SpecialItem = Not String.IsNullOrWhiteSpace(child.SpecialElement),
                    .SpecialItemData = child.SpecialElement
                }

                If String.Equals(child.IsMaster, "True", StringComparison.OrdinalIgnoreCase) Then
                    groupNode.Elements.Add(New AccordionControlElement With {
                        .Name = "CombinedItem" & groupIndex.ToString() & "_" & childID.ToString(),
                        .Text = child.NavigationText & " >", .Tag = target,
                        .Style = DevExpress.XtraBars.Navigation.ElementStyle.Group
                    })
                    currentSubgroup = Nothing
                    currentSubgroupName = Nothing
                    Continue For
                End If

                Dim subgroupName As String = If(child.GroupName, String.Empty).Trim()
                If subgroupName.Length = 0 OrElse subgroupName = "None" Then
                    currentSubgroup = Nothing
                    currentSubgroupName = Nothing
                ElseIf Not String.Equals(currentSubgroupName, subgroupName, StringComparison.Ordinal) Then
                    currentSubgroup = New AccordionControlElement With {
                        .Name = "CombinedSection" & groupIndex.ToString() & "_" & childID.ToString(),
                        .Text = child.NavigationGroupText,
                        .Style = DevExpress.XtraBars.Navigation.ElementStyle.Group
                    }
                    groupNode.Elements.Add(currentSubgroup)
                    currentSubgroupName = subgroupName
                End If

                Dim childNode As New AccordionControlElement With {
                    .Name = "CombinedItem" & groupIndex.ToString() & "_" & childID.ToString(),
                    .Text = child.NavigationText, .Tag = target,
                    .Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
                }
                If currentSubgroup Is Nothing Then
                    groupNode.Elements.Add(childNode)
                Else
                    currentSubgroup.Elements.Add(childNode)
                End If
            Next

            root.Elements.Add(groupNode)
        Next

        AccordionControlNavigator.BeginUpdate()
        Try
            AccordionControlNavigator.Elements.Clear()
            AccordionControlNavigator.Elements.Add(root)
        Finally
            AccordionControlNavigator.EndUpdate()
        End Try
        AddHandler AccordionControlNavigator.ElementClick, AddressOf AccordionControlNavigator_ElementClick
        ResizeFonts()
    End Sub

    Private Shared Sub SetCombinedGroupAppearance(
        ByVal appearance As DevExpress.Utils.AppearanceObject)
        appearance.BackColor = Color.Black
        appearance.BackColor2 = Color.Black
        appearance.ForeColor = Color.White
        appearance.Options.UseBackColor = True
        appearance.Options.UseForeColor = True
    End Sub

    Public Sub RefreshSummaryData(Optional ByVal refreshReason As String = "Direct")

        If IsDisposed OrElse Disposing OrElse
           ExcelModels Is Nothing OrElse
           MyModelID < 0 OrElse MyModelID >= ExcelModels.Length OrElse
           ExcelModels(MyModelID) Is Nothing OrElse
           ExcelModels(MyModelID).WB Is Nothing Then Return

        Dim Workbook As DevExpress.Spreadsheet.IWorkbook =
            ExcelModels(MyModelID).WB
        Dim ModelProfile As WorkbookModelProfile =
            ExcelModels(MyModelID).Profile
        Dim ModelDescription As String =
            If(ModelProfile Is Nothing,
               "Abovo model",
               ModelProfile.DisplayName)

        Dim HasBPSummary As Boolean = Workbook.Worksheets.Contains("BP Dashboard")
        Dim HasFundingSummary As Boolean = Workbook.Worksheets.Contains("Funding Assumptions")
        Dim HasDevelopmentSummary As Boolean = Workbook.Worksheets.Contains("Development Dashboard")

        AccordionControlElementBPStat.Visible = HasBPSummary
        AccordionControlElementFund.Visible = HasFundingSummary
        AccordionControlElementDev.Visible = HasDevelopmentSummary

        Debug.WriteLine("GroupInterfaceTemplate sidebar refresh started. ModelID=" &
                        MyModelID.ToString() & ", GSID=" & GSID.ToString() &
                        ", instance=" & System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(Me).ToString() &
                        ", reason=" & refreshReason)

        'Navigation, interface construction, calculation-completed notifications and
        'history changes refresh the displayed snapshot only.  A full workbook
        'calculation is deliberately reserved for the user's Refresh button.
        If String.Equals(refreshReason, "Manual", StringComparison.OrdinalIgnoreCase) AndAlso
           (HasBPSummary OrElse HasFundingSummary OrElse HasDevelopmentSummary) Then
            CalculateSidebarWorkbook(Workbook, refreshReason)
        End If

        If HasBPSummary Then
            Dim DataRange As DevExpress.Spreadsheet.CellRange =
                Workbook.Worksheets("BP Dashboard").Range("H6:L23")

            Dim DLList As New List(Of DevExpress.Spreadsheet.CellRange)
            DLList.Add(DataRange)

            SetSidebarDocument(WebBrowserBPSum, CompactSummaryHtml(
                ExcelModels(MyModelID).WBData.RenderIEHTMLCourceFromDR(DLList)))
        Else
            SetSidebarDocument(WebBrowserBPSum,
                "<html><body><p>" & ModelDescription &
                " does not define a Business Plan dashboard summary.</p></body></html>")
        End If

        If HasFundingSummary Then
            Dim FDSList As New List(Of DevExpress.Spreadsheet.CellRange)
            Dim DataRange As DevExpress.Spreadsheet.CellRange =
                Workbook.Worksheets("Funding Assumptions").Range("E3:I5")
            FDSList.Add(DataRange)

            DataRange =
                Workbook.Worksheets("Funding Assumptions").Range("J3:N5")
            FDSList.Add(DataRange)

            SetSidebarDocument(WebBrowserFundSum, CompactSummaryHtml(
                ExcelModels(MyModelID).WBData.RenderIEHTMLCourceFromDR(FDSList, New Integer() {0, 1})))
        Else
            SetSidebarDocument(WebBrowserFundSum,
                "<html><body><p>No funding summary is defined for this " &
                ModelDescription & ".</p></body></html>")
        End If

        SetSidebarDocument(WebBrowserAboutHelp, CreateSidebarHtml(
            "<h2>abovo-summit version " & WebUtility.HtmlEncode(DecVersionNumber.ToString()) & "</h2>" &
            "<p>© 2015-" & Year(Now()).ToString() & " Abovo Business Services Limited.</p>" &
            "<p><a href='summit-help'>Open Summit Help</a></p>" &
            "<p><a href='https://www.abovo-consult.co.uk'>www.abovo-consult.co.uk</a><br>" &
            "<a href='mailto:support@abovo-consult.co.uk'>support@abovo-consult.co.uk</a></p>" &
            "<p>Built using Microsoft&reg; Excel&reg; and DevExpress.</p>"))


        Dim StrFileDescription As String

        Dim FileInformation As System.IO.FileInfo = ExcelModels(MyModelID).FileInfo
        StrFileDescription = "<h2>Model details</h2>" &
            "<dl><dt>Model type</dt><dd>" & Html(ModelDescription) & "</dd>" &
            "<dt>Model name</dt><dd>" & Html(ExcelModels(MyModelID).WBStructure.CompanyName) & "</dd>" &
            "<dt>Start date</dt><dd>" & Html(ExcelModels(MyModelID).WBStructure.StartDate) & "</dd>" &
            "<dt>File</dt><dd>" & Html(ExcelModels(MyModelID).FileName) & "</dd>"
        If FileInformation IsNot Nothing Then
            StrFileDescription &= "<dt>Created</dt><dd>" & Html(FileInformation.CreationTime.ToString("g")) & "</dd>" &
                "<dt>Previous file access</dt><dd>" &
                Html(If(ExcelModels(MyModelID).PreviousFileAccessTime = DateTime.MinValue,
                        "Not recorded",
                        ExcelModels(MyModelID).PreviousFileAccessTime.ToString("G"))) & "</dd>" &
                "<dt>Size</dt><dd>" & Format((FileInformation.Length / 1000000), "###.##") & " MB</dd>"
        End If
        StrFileDescription &= "</dl>"

        SetSidebarDocument(WebBrowserFile, CreateSidebarHtml(StrFileDescription))
        If SidebarMessageView IsNot Nothing Then SidebarMessageView.RefreshMessages()
        DockPanelDetail.Text = If(ExcelModels(MyModelID).DeferredSaveResultsPending,
            "Summary — full results pending", "Summary — updated " & Now().ToString("HH:mm:ss"))
        Debug.WriteLine("GroupInterfaceTemplate sidebar refresh completed. ModelID=" &
                        MyModelID.ToString() & ", GSID=" & GSID.ToString() &
                        ", instance=" & System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(Me).ToString() &
                        ", reason=" & refreshReason)

    End Sub

    Private Sub CalculateSidebarWorkbook(ByVal workbook As DevExpress.Spreadsheet.IWorkbook,
                                         ByVal refreshReason As String)
        If workbook Is Nothing Then Return
        Dim Activity As FormSplashScreen = Nothing
        Dim previousEngine As DevExpress.Spreadsheet.CalculationEngineType =
            workbook.Options.CalculationEngineType
        Dim previousCursor As Cursor = Me.Cursor
        Dim previousUseWaitCursor As Boolean = Me.UseWaitCursor
        Try
            Activity = New FormSplashScreen(
                Me, "Calculating model", "Refreshing all workbook results...")
            Me.UseWaitCursor = True
            Me.Cursor = Cursors.WaitCursor
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Not ExcelModels(MyModelID).EnsureDeferredSaveResultsCurrent("Refreshing summary figures...") Then
                workbook.Options.CalculationEngineType =
                    DevExpress.Spreadsheet.CalculationEngineType.Recursive
                workbook.CalculateFull()
            End If
            Activity.Complete("Calculation complete.")
            Debug.WriteLine("GroupInterfaceTemplate sidebar workbook calculated. ModelID=" &
                            MyModelID.ToString() & ", reason=" & refreshReason)
        Catch ex As Exception
            Debug.WriteLine("GroupInterfaceTemplate sidebar workbook calculation failed. ModelID=" &
                            MyModelID.ToString() & ", error=" & ex.ToString())
            SystemMessageManager.Publish(MyModelID,
                "The sidebar summary could not be recalculated: " & ex.Message,
                SystemMessageSeverity.Warning,
                "Summary refresh")
        Finally
            If Activity IsNot Nothing Then Activity.Dispose()
            Try
                workbook.Options.CalculationEngineType = previousEngine
            Finally
                Me.UseWaitCursor = previousUseWaitCursor
                Me.Cursor = previousCursor
                System.Windows.Forms.Cursor.Current = previousCursor
            End Try
        End Try
    End Sub

    Private Shared Function CompactSummaryHtml(ByVal sourceHtml As String) As String
        If String.IsNullOrWhiteSpace(sourceHtml) Then Return sourceHtml
        Dim UserScale As Single = Abovo.PresentationScaleManager.UserScale
        Dim FontSize As String = (9.5F * UserScale).ToString("0.##", Globalization.CultureInfo.InvariantCulture)
        Dim RowHeight As String = Math.Max(14, CInt(Math.Round(18 * UserScale))).ToString()
        Dim HorizontalPadding As String = Math.Max(2, CInt(Math.Round(3 * UserScale))).ToString()
        Dim compactStyle As String =
            "<style type='text/css'>" &
            "html,body{margin:0!important;padding:2px!important;font-size:" & FontSize & "pt!important;}" &
            "table{width:100%!important;margin:0!important;border-collapse:collapse!important;}col{width:auto!important;}" &
            "tr{height:" & RowHeight & "px!important;min-height:" & RowHeight & "px!important;}" &
            "td,th{height:" & RowHeight & "px!important;min-height:0!important;padding:1px " &
            HorizontalPadding & "px!important;font-size:" & FontSize &
            "pt!important;line-height:1.15!important;white-space:normal!important;}" &
            "</style>"
        'Excel exports spacer rows as cells containing only non-breaking spaces.
        'Remove only those empty rows, never a zero or a labelled check.
        sourceHtml = System.Text.RegularExpressions.Regex.Replace(sourceHtml, "<tr\b[^>]*>.*?</tr>",
            Function(m)
                Dim contents = System.Text.RegularExpressions.Regex.Replace(m.Value, "<[^>]+>", "")
                Return If(String.IsNullOrWhiteSpace(WebUtility.HtmlDecode(contents)), "", m.Value)
            End Function, System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Singleline)
        Dim headEnd As Integer = sourceHtml.IndexOf("</head>", StringComparison.OrdinalIgnoreCase)
        If headEnd >= 0 Then Return sourceHtml.Insert(headEnd, compactStyle)
        Return compactStyle & sourceHtml
    End Function

    Private Sub InitialiseRightSidebar()
        'Keep the auto-hide tab available, but do not make the user wait for
        'the dock panel to slide closed or open.
        DockManagerAssumptions.AutoHideSpeed = 10000
        AccordionControlSum.AnimationType = DevExpress.XtraBars.Navigation.AnimationType.None
        AccordionControlSum.AllowSmoothScrolling = False
        SidebarMessageView = New SystemMessageView(MyModelID) With {.Dock = DockStyle.Fill}
        AccordionContentContainerSystemMessages.Controls.Add(SidebarMessageView)

        SidebarHistoryContainer = New AccordionContentContainer With {
            .Name = "AccordionContentContainerInterfaceHistory",
            .Size = New Size(540, 230)}
        SidebarHistoryView = New InterfaceHistoryView(
            MyModelID,
            ExcelModels(MyModelID).InterfaceHistory) With {.Dock = DockStyle.Fill}
        SidebarHistoryContainer.Controls.Add(SidebarHistoryView)
        AccordionControlSum.Controls.Add(SidebarHistoryContainer)

        SidebarHistoryElement = New AccordionControlElement With {
            .ContentContainer = SidebarHistoryContainer,
            .Name = "AccordionControlElementInterfaceHistory",
            .Style = ElementStyle.Item,
            .Text = "Interface History",
            .Expanded = False}
        AccordionControlSum.Elements.Add(SidebarHistoryElement)
        'Initial state only. Refreshes and navigation must retain the user's choices.
        For Each element As AccordionControlElement In AccordionControlSum.Elements
            element.Expanded = element Is AccordionControlElementBPStat OrElse
                               element Is AccordionControlElementFund
        Next
        ApplySidebarAccordionAppearance()

        For Each browser As WebBrowser In {WebBrowserBPSum, WebBrowserDevSum,
                                           WebBrowserFundSum, WebBrowserAboutHelp,
                                           WebBrowserFile}
            browser.Tag = PresentationLayout.BrowserOwnsScale
            browser.AllowWebBrowserDrop = False
            browser.IsWebBrowserContextMenuEnabled = False
            browser.ScriptErrorsSuppressed = True
            browser.WebBrowserShortcutsEnabled = True
            AddHandler browser.Navigating, AddressOf SidebarBrowser_Navigating
            AddHandler browser.DocumentCompleted, Sub(s, e) ApplySidebarDocument(DirectCast(s, WebBrowser))
            AddHandler browser.HandleCreated, Sub(s, e)
                                                  Dim readyBrowser = DirectCast(s, WebBrowser)
                                                  readyBrowser.BeginInvoke(New MethodInvoker(Sub() ApplySidebarDocument(readyBrowser)))
                                              End Sub
            AddHandler browser.HandleDestroyed, Sub(s, e)
                                                    Dim oldBrowser = DirectCast(s, WebBrowser)
                                                    SidebarRenderedHtml.Remove(oldBrowser)
                                                    SidebarDocumentLoading.Remove(oldBrowser)
                                                End Sub
        Next

        SidebarRefreshTimer = New Timer With {.Interval = 400}
        AddHandler SidebarRefreshTimer.Tick, AddressOf SidebarRefreshTimer_Tick
        AddHandler DockPanelDetail.CustomButtonClick, AddressOf DockPanelDetail_CustomButtonClick
        AddHandler DockPanelDetail.Expanding, AddressOf DockPanelDetail_Expanding
        ReconnectPanelStrips()

        If ExcelModels(MyModelID).WBCalcEngine IsNot Nothing Then
            AddHandler ExcelModels(MyModelID).WBCalcEngine.CalculationCompleted,
                AddressOf WorkbookCalculationCompleted
        End If
        If ExcelModels(MyModelID).ChangeManager IsNot Nothing Then
            AddHandler ExcelModels(MyModelID).ChangeManager.HistoryChanged,
                AddressOf WorkbookHistoryChanged
        End If
        SidebarEventsAttached = True

        SystemMessageManager.Publish(MyModelID,
            MyName & " interface opened.",
            SystemMessageSeverity.Information,
            "Interface")
    End Sub

    Friend Sub RequestSidebarRefresh()
        If IsDisposed OrElse Disposing OrElse SidebarRefreshTimer Is Nothing Then Return
        If InvokeRequired Then
            BeginInvoke(New MethodInvoker(AddressOf RequestSidebarRefresh))
            Return
        End If
        SidebarRefreshTimer.Stop()
        SidebarRefreshTimer.Start()
    End Sub

    Private Sub WorkbookCalculationCompleted(ByVal sender As Object, ByVal e As EventArgs)
        RequestSidebarRefresh()
    End Sub

    Private Sub WorkbookHistoryChanged(ByVal sender As Object,
                                       ByVal e As ChangeHistoryChangedEventArgsV2)
        RequestSidebarRefresh()
    End Sub

    Private Sub SidebarRefreshTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
        SidebarRefreshTimer.Stop()
        RefreshSummaryData("Automatic")
    End Sub

    Private Sub DockPanelDetail_Expanding(ByVal sender As Object,
                                         ByVal e As DevExpress.XtraBars.Docking.DockPanelCancelEventArgs)
        'Like the navigator, reopen explicitly from the strip Click below.
        'A delayed native hover expansion must not undo a newer compact/restore.
        e.Cancel = True
    End Sub

    Private Sub hideContainerRightDetail_Click(sender As Object, e As EventArgs) Handles hideContainerRightDetail.Click
        If IsDisposed OrElse Disposing Then Return
        If DockPanelDetail.Visibility = DevExpress.XtraBars.Docking.DockVisibility.AutoHide Then
            DockPanelDetail.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Visible
        End If
    End Sub

    Private Sub DockPanelDetail_CustomButtonClick(
        ByVal sender As Object,
        ByVal e As DevExpress.XtraBars.Docking2010.ButtonEventArgs)

        If DockPanelDetail.CustomHeaderButtons.Count > 1 AndAlso
           Object.ReferenceEquals(e.Button, DockPanelDetail.CustomHeaderButtons(1)) Then
            DockPanelDetail.Visibility = DevExpress.XtraBars.Docking.DockVisibility.AutoHide
            DockPanelDetail.HideImmediately()
            ReconnectPanelStrips()
            Return
        End If

        RefreshSummaryData("Manual")
        SystemMessageManager.Publish(MyModelID,
            "The sidebar summary was refreshed.",
            SystemMessageSeverity.Success,
            "Summary")
    End Sub

    Private Sub SidebarBrowser_Navigating(ByVal sender As Object,
                                          ByVal e As WebBrowserNavigatingEventArgs)
        If e.Url Is Nothing OrElse e.Url.ToString().Equals("about:blank", StringComparison.OrdinalIgnoreCase) Then Return
        e.Cancel = True
        If e.Url.ToString().IndexOf("summit-help", StringComparison.OrdinalIgnoreCase) >= 0 Then
            HelpManager.ShowHelpHome(Me)
            Return
        End If
        If e.Url.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) OrElse
           e.Url.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) OrElse
           e.Url.Scheme.Equals(Uri.UriSchemeMailto, StringComparison.OrdinalIgnoreCase) Then
            Process.Start(New ProcessStartInfo(e.Url.ToString()) With {.UseShellExecute = True})
        End If
    End Sub

    Private Shared Function Html(ByVal value As Object) As String
        Return WebUtility.HtmlEncode(If(value, String.Empty).ToString())
    End Function

    Private Sub SetSidebarDocument(browser As WebBrowser, htmlText As String)
        'DocumentText navigates asynchronously. Repeated pre-show assignments can
        'lose the pending stream and leave about:blank. Keep the latest content
        'until the ActiveX document is ready, then update it without navigating.
        SidebarHtml(browser) = htmlText
        ApplySidebarDocument(browser)
    End Sub

    Private Sub ApplySidebarDocument(browser As WebBrowser)
        If IsDisposed OrElse Disposing OrElse browser.IsDisposed OrElse
           Not browser.IsHandleCreated OrElse SidebarDocumentUpdating.Contains(browser) Then Return
        Dim htmlText As String = Nothing
        If Not SidebarHtml.TryGetValue(browser, htmlText) Then Return
        If browser.Document Is Nothing OrElse browser.ReadyState <> WebBrowserReadyState.Complete Then
            If SidebarDocumentLoading.Add(browser) Then browser.Navigate("about:blank")
            Return
        End If
        SidebarDocumentLoading.Remove(browser)
        Dim previous As String = Nothing
        If SidebarRenderedHtml.TryGetValue(browser, previous) AndAlso previous = htmlText AndAlso
           browser.Document.Body IsNot Nothing AndAlso Not String.IsNullOrEmpty(browser.Document.Body.InnerText) Then Return
        SidebarDocumentUpdating.Add(browser)
        Try
            Dim scroll As Point = If(browser.Document.Body Is Nothing, Point.Empty,
                                      New Point(browser.Document.Body.ScrollLeft, browser.Document.Body.ScrollTop))
            Dim document As HtmlDocument = browser.Document.OpenNew(True)
            document.Write(htmlText)
            SidebarRenderedHtml(browser) = htmlText
            'Complete the synchronous MSHTML stream, otherwise ReadyState stays
            'Interactive and a later refresh would restart about:blank.
            CallByName(document.DomDocument, "close", CallType.Method)
            If document.Window IsNot Nothing Then document.Window.ScrollTo(scroll)
        Finally
            SidebarDocumentUpdating.Remove(browser)
        End Try
    End Sub

    Private Function CreateSidebarHtml(ByVal body As String) As String
        Dim points As String = (9.5F * PresentationScaleManager.UserScale).ToString("0.##", Globalization.CultureInfo.InvariantCulture)
        Return "<!doctype html><html><head><meta charset='utf-8'><meta http-equiv='X-UA-Compatible' content='IE=edge'><style>" &
            "body{font-family:Segoe UI,Arial,sans-serif;font-size:" & points & "pt;color:#333;margin:6px;background:#fff;line-height:1.2}" &
            "h2{color:#075da8;font-size:1.15em;margin:0 0 6px}" &
            "a{color:#075da8}p{margin:3px 0}dl{margin:4px 0}dt{font-weight:600;float:left;clear:left;width:32%;padding:2px 0}" &
            "dd{margin:0 0 0 35%;padding:2px 0;word-wrap:break-word}</style></head><body>" & body & "</body></html>"
    End Function

    Private Sub GroupInterfaceTemplate_Disposed(ByVal sender As Object,
                                                ByVal e As EventArgs) Handles Me.Disposed
        If SidebarEventsAttached AndAlso
           ExcelModels IsNot Nothing AndAlso
           MyModelID >= 0 AndAlso MyModelID < ExcelModels.Length AndAlso
           ExcelModels(MyModelID) IsNot Nothing Then
            If ExcelModels(MyModelID).WBCalcEngine IsNot Nothing Then
                RemoveHandler ExcelModels(MyModelID).WBCalcEngine.CalculationCompleted,
                    AddressOf WorkbookCalculationCompleted
            End If
            If ExcelModels(MyModelID).ChangeManager IsNot Nothing Then
                RemoveHandler ExcelModels(MyModelID).ChangeManager.HistoryChanged,
                    AddressOf WorkbookHistoryChanged
            End If
        End If
        RemoveHandler DockPanelDetail.Expanding, AddressOf DockPanelDetail_Expanding
        SidebarEventsAttached = False
        If SidebarRefreshTimer IsNot Nothing Then
            SidebarRefreshTimer.Stop()
            SidebarRefreshTimer.Dispose()
            SidebarRefreshTimer = Nothing
        End If
        If SidebarMessageView IsNot Nothing Then
            SidebarMessageView.Dispose()
            SidebarMessageView = Nothing
        End If
        If SidebarHistoryView IsNot Nothing Then
            SidebarHistoryView.Dispose()
            SidebarHistoryView = Nothing
        End If
    End Sub

    Sub AddNavigatorItem(ItemName As String, IntTag As AbovoInterfaceTag, IsMaster As String, Optional ByVal GroupName As String = "None")



        If IsMaster = "True" Then 'Single Master element
            'GroupName <> LastGroup Then

            If OpenGroup Then ' close last group
                AccGroups(AccGroupCount).Elements.AddRange(GroupElements)
                GroupElementsCount = -1
                ReDim GroupElements(-1)
                OpenGroup = False
            End If

            AccGroupCount += 1

            ReDim Preserve AccGroups(AccGroupCount)

            AccGroups(AccGroupCount) = New AccordionControlElement With {.Tag = IntTag, .Text = ItemName, .Name = "Acg" & AccGroupCount.ToString}
            AccGroups(AccGroupCount).Text = AccGroups(AccGroupCount).Text & " >"
            AccGroups(AccGroupCount).Style = DevExpress.XtraBars.Navigation.ElementStyle.Group

            Exit Sub

        End If

        If GroupName = LastGroup Then 'Another item for group

            GroupElementsCount += 1
            ReDim Preserve GroupElements(GroupElementsCount)
            GroupElements(GroupElementsCount) = New AccordionControlElement With {.Tag = IntTag, .Text = ItemName, .Name = "Acge" & GroupElementsCount.ToString}
            GroupElements(GroupElementsCount).Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
            Exit Sub

        End If

        If OpenGroup Then ' close last group

            AccGroups(AccGroupCount).Elements.AddRange(GroupElements)
            GroupElementsCount = -1
            ReDim GroupElements(-1)
            OpenGroup = False

        End If

        'New non-master group and item
        AccGroupCount += 1
        ReDim Preserve AccGroups(AccGroupCount)
        AccGroups(AccGroupCount) = New AccordionControlElement With {.Tag = IntTag, .Text = GroupName, .Name = "Acg" & AccGroupCount.ToString}
        AccGroups(AccGroupCount).Style = DevExpress.XtraBars.Navigation.ElementStyle.Group
        LastGroup = GroupName
        GroupElementsCount += 1
        ReDim Preserve GroupElements(GroupElementsCount)
        GroupElements(GroupElementsCount) = New AccordionControlElement With {.Tag = IntTag, .Text = ItemName, .Name = "Acge" & GroupName & GroupElementsCount.ToString}
        GroupElements(GroupElementsCount).Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
        OpenGroup = True

    End Sub
    Sub FinaliseNavigator()

        If OpenGroup Then ' close last group

            AccGroups(AccGroupCount).Elements.AddRange(GroupElements)

        End If

        AccordionControlNavigator.Elements.Clear()

        Dim AccRoot(0) As AccordionControlElement
        AccRoot(0) = New AccordionControlElement With {.Tag = "-1", .Text = "Root", .Name = "RootItem", .HeaderVisible = False, .Expanded = True}

        AccRoot(0).Elements.AddRange(AccGroups)
        AccordionControlNavigator.Elements.AddRange(AccRoot)

        AddHandler AccordionControlNavigator.ElementClick, AddressOf AccordionControlNavigator_ElementClick
        ResizeFonts()

    End Sub
    Sub LoadDefaultInterface()

        'Dim NewSAI As New VideoPlayer("NavGuide")
        'NewSAI.Tag = -1
        'DocumentManagerAssumptions.View.AddDocument(NewSAI)
        'DocumentManagerAssumptions.View.ActivateDocument(NewSAI)

    End Sub

    Public Sub LoadDocument(TargetID As Integer, Optional ByVal ShowSpecial As Boolean = False, Optional ByVal SpecialData As String = "None", Optional ByVal TargetGSID As Integer = -1)
        If ShowSpecial Then
            ShowInterface(MyModelID, TargetID, True, SpecialData, Nothing, TargetGSID)
        Else
            ShowInterface(MyModelID, TargetID, False, "None", Nothing, TargetGSID)
        End If
    End Sub

    Private Sub AccordionControlNavigator_ElementClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.Navigation.ElementClickEventArgs)

        If e.Element.Tag Is Nothing Then

            Return

        End If

        Dim ItemTag As AbovoInterfaceTag = DirectCast(e.Element.Tag, AbovoInterfaceTag)

        If IsCombined Then
            If ItemTag.TargetID < 0 OrElse ItemTag.TargetGroupID < 0 Then Return
            ShowInterface(MyModelID, ItemTag.TargetID, ItemTag.SpecialItem,
                          ItemTag.SpecialItemData, Nothing, ItemTag.TargetGroupID)
            Return
        End If

        If ItemTag.SpecialItem = False Then

            If ItemTag.TargetID = -1 Then Return
            ShowInterface(MyModelID, ItemTag.TargetID)

        End If

        ShowInterface(MyModelID, ItemTag.TargetID, True, ItemTag.SpecialItemData)

    End Sub


    Sub SetInitialSizes()

        Dim AvailableArea As Rectangle = Screen.FromPoint(Cursor.Position).WorkingArea
        If AvailableArea.Width < 900 Then

            Me.Width = CInt(AvailableArea.Width * 0.85)
            Me.Height = CInt(AvailableArea.Height * 0.85)

        ElseIf AvailableArea.Width < 1190 Then

            Me.Width = CInt(AvailableArea.Width * 0.75)
            Me.Height = CInt(AvailableArea.Height * 0.75)

        Else

            Me.Width = CInt(AvailableArea.Width * 0.65)
            Me.Height = CInt(AvailableArea.Height * 0.65)

        End If

        ResizeFonts()
        ResizeControls()

    End Sub
    Sub ResizeControls()

        ScaleFactor = Me.Width / 2400

        Dim SetWidth As Integer = Me.Width * 0.22
        ScaleUnits = Me.Width * 0.007

        DockPanelNavigator.Width = SetWidth
        FitNavigatorWidth()
        'DockManagerAssumptions.
    End Sub
    Sub ResizeFonts()

        ScaleFactor = GetDisplayScale(Me)
        LastPresentationScale = ScaleFactor

        If Not hideContainerRightDetail.IsDisposed Then
            hideContainerRightDetail.Font = GetDisplayFont("Small", Me)
        End If
        Me.BarAndDockingControllerAssumptions.AppearancesDocking.ActiveTab.Font = GetDisplayFont("Medium", Me)
        Me.BarAndDockingControllerAssumptions.AppearancesDocking.HidePanelButton.Font = GetDisplayFont("Medium", Me)
        Me.BarAndDockingControllerAssumptions.AppearancesDocking.HidePanelButtonActive.Font = GetDisplayFont("Medium", Me)
        Me.BarAndDockingControllerAssumptions.AppearancesDocking.PanelCaption.Font = GetDisplayFont("Medium", Me)
        Me.BarAndDockingControllerAssumptions.AppearancesDocking.PanelCaptionActive.Font = GetDisplayFont("Medium", Me)
        Me.BarTopBar.BarAppearance.Normal.Font = GetDisplayFont("Medium", Me)
        Me.BarStaticItemDescription.ItemAppearance.Normal.Font = GetDisplayFont("Medium", Me)
        BarTopBar.OptionsBar.MinHeight = CInt(36 * ScaleFactor)
        Me.AccordionControlNavigator.Appearance.Group.Hovered.Font = GetDisplayFont("Medium", Me)
        Me.AccordionControlNavigator.Appearance.Group.Default.Font = GetDisplayFont("Medium", Me)
        Me.AccordionControlNavigator.Appearance.Group.Normal.Font = GetDisplayFont("Medium", Me)
        Me.AccordionControlNavigator.Appearance.Item.Normal.Font = GetDisplayFont("Small", Me)
        Me.AccordionControlNavigator.Appearance.Item.Default.Font = GetDisplayFont("Small", Me)
        Me.AccordionControlNavigator.Appearance.Item.Hovered.Font = GetDisplayFont("Small", Me)
        ApplySidebarAccordionAppearance()

        Me.AccordionControlNavigator.BeginUpdate()
        Try
            For Each Element As DevExpress.XtraBars.Navigation.AccordionControlElement In
                Me.AccordionControlNavigator.Elements
                ApplyNavigatorElementFont(Element)
            Next
        Finally
            Me.AccordionControlNavigator.EndUpdate()
        End Try

        FitNavigatorWidth()

    End Sub

    Private Sub ApplySidebarAccordionAppearance()
        'These are content-hosting Item elements, not navigational Groups. Keep
        'their style/containers and expansion state; change header appearance only.
        'Match headers only; do not change the renderer of hosted content.
        AccordionControlSum.BeginUpdate()
        Try
            Dim sidebarFont As New Font("Segoe UI", 9.5F * PresentationScaleManager.UserScale, FontStyle.Regular)
            For Each element As AccordionControlElement In AccordionControlSum.Elements
                For Each appearance As AppearanceObject In {element.Appearance.Default, element.Appearance.Normal,
                                        element.Appearance.Hovered, element.Appearance.Pressed, element.Appearance.Disabled}
                    appearance.Assign(AccordionControlNavigator.Appearance.Group.Default)
                    appearance.Font = sidebarFont
                    appearance.BackColor = AccordionControlNavigator.Appearance.Group.Default.BackColor
                    appearance.ForeColor = Color.White
                    appearance.Options.UseFont = True
                    appearance.Options.UseBackColor = True
                    appearance.Options.UseForeColor = True
                    appearance.TextOptions.WordWrap = WordWrap.NoWrap
                    appearance.TextOptions.Trimming = Trimming.EllipsisCharacter
                    appearance.Options.UseTextOptions = True
                Next
            Next
        Finally
            AccordionControlSum.EndUpdate()
        End Try
    End Sub

    Private Sub FitNavigatorWidth()
        If PanelsHidden OrElse AccordionControlNavigator Is Nothing OrElse ClientSize.Width < 1 Then Return
        Dim scale As Single = GetDisplayScale(Me)
        If LastPanelLayoutWidth = ClientSize.Width AndAlso Math.Abs(LastPanelLayoutScale - scale) < 0.001F Then Return
        LastPanelLayoutWidth = ClientSize.Width
        LastPanelLayoutScale = scale
        Dim wanted As Integer = CInt(200 * CSng(DeviceDpi) / 96.0F)
        For Each element As AccordionControlElement In AccordionControlNavigator.Elements
            wanted = Math.Max(wanted, MeasureNavigatorElement(element, 0))
        Next
        'Keep a useful document area on small/side-by-side windows. Exceptionally
        'long captions ellipsize instead of wrapping or shrinking the user's font.
        DockPanelNewNavigator.Width = Math.Min(wanted, Math.Max(180, CInt(ClientSize.Width * 0.25)))
        DockPanelDetail.Width = Math.Min(CInt(340 * scale * DeviceDpi / 96.0F), Math.Max(220, CInt(ClientSize.Width * 0.22)))
        Dim compactScale As Single = PresentationScaleManager.UserScale * DeviceDpi / 96.0F
        AccordionContentContainer1.Height = CInt(240 * compactScale)
        AccordionContentContainer5.Height = CInt(230 * compactScale)
    End Sub

    Private Function MeasureNavigatorElement(element As AccordionControlElement, depth As Integer) As Integer
        Dim dpi As Single = CSng(DeviceDpi) / 96.0F
        Dim width As Integer = 0
        If element.HeaderVisible Then
            width = TextRenderer.MeasureText(element.Text, element.Appearance.Normal.Font,
                Size.Empty, TextFormatFlags.SingleLine Or TextFormatFlags.NoPadding).Width + CInt((64 + depth * 26) * dpi)
        End If
        For Each child As AccordionControlElement In element.Elements
            width = Math.Max(width, MeasureNavigatorElement(child, depth + If(element.HeaderVisible, 1, 0)))
        Next
        Return width
    End Function

    Private Sub ApplyNavigatorElementFont(
        ByVal Element As DevExpress.XtraBars.Navigation.AccordionControlElement,
        Optional ByVal depth As Integer = 0)

        Dim FontClass As String =
            If(Element.Style = DevExpress.XtraBars.Navigation.ElementStyle.Group, "Medium", "Small")
        Dim ElementFont As Font = GetDisplayFont(FontClass, Me)

        Element.Appearance.Default.Font = ElementFont
        Element.Appearance.Normal.Font = ElementFont
        Element.Appearance.Hovered.Font = ElementFont
        Element.Appearance.Pressed.Font = ElementFont
        Element.Appearance.Disabled.Font = ElementFont

        Element.Appearance.Default.Options.UseFont = True
        Element.Appearance.Normal.Options.UseFont = True
        Element.Appearance.Hovered.Options.UseFont = True
        Element.Appearance.Pressed.Options.UseFont = True
        Element.Appearance.Disabled.Options.UseFont = True
        For Each appearance As AppearanceObject In {Element.Appearance.Default, Element.Appearance.Normal,
                                Element.Appearance.Hovered, Element.Appearance.Pressed, Element.Appearance.Disabled}
            appearance.TextOptions.WordWrap = WordWrap.NoWrap
            appearance.TextOptions.Trimming = Trimming.EllipsisCharacter
            appearance.Options.UseTextOptions = True
        Next

        If IsCombined AndAlso depth = 1 AndAlso
           Element.Style = DevExpress.XtraBars.Navigation.ElementStyle.Group Then
            SetCombinedGroupAppearance(Element.Appearance.Default)
            SetCombinedGroupAppearance(Element.Appearance.Normal)
            SetCombinedGroupAppearance(Element.Appearance.Hovered)
            SetCombinedGroupAppearance(Element.Appearance.Pressed)
        End If

        For Each Child As DevExpress.XtraBars.Navigation.AccordionControlElement In Element.Elements
            ApplyNavigatorElementFont(Child, depth + 1)
        Next
    End Sub
    Private Sub TabbedViewDefault_QueryControl(sender As Object, e As DevExpress.XtraBars.Docking2010.Views.QueryControlEventArgs) Handles TabbedViewDefault.QueryControl


    End Sub


    Private Sub FormMainScreen_ResizeEnd(sender As Object, e As EventArgs) Handles MyBase.ResizeEnd

        ResizeGIT()

    End Sub

    Public Sub ResizeGIT()

        If IsNothing(ActiveInterface) Then Exit Sub

        Try
            ActiveInterface.ResizeControlsCommand()
        Catch ex As Exception
        End Try

    End Sub
    Private Sub ApplyPresentationScale()
        If IsDisposed OrElse Disposing Then Return
        ResizeFonts()
        ResizeControls()
        ResizeGIT()
        If PresentationReady AndAlso Math.Abs(LastSidebarUserScale - PresentationScaleManager.UserScale) > 0.001F Then
            LastSidebarUserScale = PresentationScaleManager.UserScale
            RefreshSummaryData("Presentation")
        End If
    End Sub
    Private Sub GIT_ResizeEnd(sender As Object, e As EventArgs) Handles MyBase.ResizeEnd

        ResizeControls()


    End Sub
    Private Sub GIT_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If PresentationReady AndAlso IsHandleCreated AndAlso Not PresentationResizeQueued AndAlso Not Disposing AndAlso Not IsDisposed Then
            PresentationResizeQueued = True
            BeginInvoke(New MethodInvoker(
                Sub()
                    PresentationResizeQueued = False
                    If IsDisposed OrElse Disposing Then Return
                    If Math.Abs(LastPresentationScale - GetDisplayScale(Me)) > 0.001F Then ResizeFonts()
                    ResizeControls()
                    ResizeGIT()
                End Sub))
        End If
        If Me.WindowState = FormWindowState.Maximized Then

            AmMaximised = True

        Else

            AmMaximised = False

        End If

        If Me.WindowState <> GITWindowState Then

            If DockManagerAssumptions.ActivePanel IsNot Nothing Then DockManagerAssumptions.ActivePanel.Update()
            ResizeGIT()

        End If

        GITWindowState = Me.WindowState

    End Sub
    'Private Sub Button1_Click(sender As Object, e As EventArgs)

    '    MsgBox(AccordionControlNavigator.Width.ToString)
    '    MsgBox(TablePanelNavigator.Width.ToString)
    '    MsgBox(DockPanelNavigator.Width.ToString)
    '    MsgBox(DockPanelNavigator_Container.Width.ToString)

    'End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        e.Cancel = True
        Me.Hide()
        FormMainScreen.BringToFront()

    End Sub
    Public Sub ShowSpreadsheet(ActiveSheet As String)

    End Sub

    Private Sub RegisterAnalysisV1(ByVal analyser As BPIncomeExpenditureAnalyser)

        Dim registry As ModelResourceRegistry =
            ExcelModels(MyModelID).ResourceRegistry

        registry.RegisterExclusive(
            ModelResourceKeys.TransactionalRecordsRangeDataSource,
            analyser,
            Sub()
                Try
                    ReleaseAnalysisV1(analyser)
                Finally
                    DisposeAnalyserDocument(analyser)
                End Try
            End Sub)

        AddHandler analyser.Disposed,
            Sub()
                ReleaseAnalysisV1(analyser)
            End Sub

    End Sub

    Private Sub RegisterAnalysisV2(ByVal analyser As BPIncomeExpenditureAnalyserV2)

        ExcelModels(MyModelID).ResourceRegistry.RegisterExclusive(
            ModelResourceKeys.TransactionalRecordsRangeDataSource,
            analyser,
            Sub()
                Try
                    analyser.ReleaseAnalyserResources()
                Finally
                    DisposeAnalyserDocument(analyser)
                End Try
            End Sub)

    End Sub

    Private Sub ReleaseAnalysisV1(ByVal analyser As BPIncomeExpenditureAnalyser)

        If analyser Is Nothing Then Return

        Try
            analyser.DisconectRDS()
        Finally
            If ExcelModels IsNot Nothing AndAlso
               MyModelID >= 0 AndAlso MyModelID < ExcelModels.Length AndAlso
               ExcelModels(MyModelID) IsNot Nothing Then

                Dim model As ExcelModel = ExcelModels(MyModelID)

                model.ResourceRegistry.Release(
                    ModelResourceKeys.TransactionalRecordsRangeDataSource,
                    analyser)

                If model.WBCalcEngine IsNot Nothing Then
                    model.WBCalcEngine.RemoveActiveObject(analyser)
                End If

                If Object.ReferenceEquals(model.ExpendAnalyser, analyser) Then
                    model.ExpendAnalyser = Nothing
                End If
            End If
        End Try

    End Sub

    Private Sub DisposeAnalyserDocument(ByVal analyser As System.Windows.Forms.Control)

        If analyser Is Nothing Then Return

        Dim document As BaseDocument =
            DocumentManagerAssumptions.View.Documents.FirstOrDefault(
                Function(candidate) Object.ReferenceEquals(candidate.Control, analyser))

        If document IsNot Nothing Then document.Dispose()
        If Not analyser.IsDisposed Then analyser.Dispose()

    End Sub

    Private Function InterfaceDocumentKey(ByVal targetGroupID As Integer,
                                          ByVal targetChildID As Integer) As String
        If IsCombined Then Return targetGroupID.ToString() & ":" & targetChildID.ToString()
        Return targetChildID.ToString()
    End Function

    Public Sub ShowInterface(SetModelID As Integer, SetCSID As Integer, Optional ByVal ShowSpecial As Boolean = False, Optional ByVal SpecialData As String = "None", Optional ByVal Interfacelink As ElementInterfaceLinkTag = Nothing, Optional ByVal TargetGSID As Integer = -1)

        Dim resolvedGSID As Integer = If(IsCombined, TargetGSID, GSID)
        If resolvedGSID < 0 OrElse
           resolvedGSID >= ExcelModels(SetModelID).WBStructure.GroupStructures.Count Then
            Throw New ArgumentOutOfRangeException(NameOf(TargetGSID),
                "The selected interface group is not available.")
        End If
        Dim documentKey As String = InterfaceDocumentKey(resolvedGSID, SetCSID)
        If resolvedGSID = 2 Then ExcelModels(SetModelID).EnsureDeferredSaveResultsCurrent("Opening model outputs...")
        Dim documentTag As Object = If(IsCombined, CObj(documentKey), CObj(SetCSID))
        Dim groupName As String = ExcelModels(SetModelID).WBStructure.GroupStructures(resolvedGSID).GSName
        Dim doc As BaseDocument = DocumentManagerAssumptions.View.Documents.FirstOrDefault(
            Function(x) x.Control IsNot Nothing AndAlso
                        String.Equals(Convert.ToString(x.Control.Tag), documentKey,
                                      StringComparison.Ordinal))

        Dim DocCount As Integer = DocumentManagerAssumptions.View.Documents.Count

        Dim DataITemp As DataInterfaceTemplate = Nothing
        Dim BIA As BPIncomeExpenditureAnalyser = Nothing

        If doc IsNot Nothing Then

            DataITemp = TryCast(doc.Control, DataInterfaceTemplate)

            If DataITemp IsNot Nothing AndAlso Interfacelink Is Nothing Then

                DataITemp.ClearLinks()

            Else
                If Interfacelink IsNot Nothing AndAlso DataITemp IsNot Nothing Then
                    DataITemp.AddLink(Interfacelink)
                End If

            End If

            If DataITemp IsNot Nothing Then 'It is a DITemplate, so reactivate and deactivate others

                If DataITemp.AmActivated = False Then

                    DataITemp.Reactivate()

                    ActiveInterface = DataITemp

                    If DocCount > 1 Then

                        For Each D As BaseDocument In DocumentManagerAssumptions.View.Documents

                            If Not String.Equals(Convert.ToString(D.Control.Tag), documentKey, StringComparison.Ordinal) Then

                                Dim OtherDITemp As DataInterfaceTemplate = TryCast(D.Control, DataInterfaceTemplate)

                                If OtherDITemp IsNot Nothing Then

                                    OtherDITemp.Deactivate()

                                End If

                            End If

                        Next

                    End If


                End If
            End If


            DocumentManagerAssumptions.View.ActivateDocument(doc.Control)

            Dim resumedAnalyser As BPIncomeExpenditureAnalyserV2 =
                TryCast(doc.Control, BPIncomeExpenditureAnalyserV2)
            If resumedAnalyser IsNot Nothing Then
                resumedAnalyser.RefreshDeferredIfNeeded()
            End If

            Me.BarStaticItemDescription.Caption = " " & ExcelModels(SetModelID).WBStructure.CompanyName & " • " & If(IsCombined, MyName & " • " & groupName, MyName) & " • " & ExcelModels(SetModelID).WBStructure.GroupStructures(resolvedGSID).ResolveChildStructure(SetCSID).CSName
            Me.Text = Me.BarStaticItemDescription.Caption

        Else

            If DocCount > 0 Then

                For Each D As BaseDocument In DocumentManagerAssumptions.View.Documents

                    If Not String.Equals(Convert.ToString(D.Control.Tag), documentKey, StringComparison.Ordinal) Then

                        Dim OtherDITemp As DataInterfaceTemplate = TryCast(D.Control, DataInterfaceTemplate)

                        If OtherDITemp IsNot Nothing Then
                            OtherDITemp.Deactivate()
                        End If

                    End If

                Next

            End If

            Me.Cursor = Cursors.WaitCursor
            If ShowSpecial Then

                Select Case SpecialData

                    Case "WrkSheet"

                        'ExcelModels(SetModelID).SSInterface.Tag = -1
                        'DocumentManagerAssumptions.View.AddDocument(ExcelModels(SetModelID).SSInterface)
                        'DocumentManagerAssumptions.View.ActivateDocument(ExcelModels(SetModelID).SSInterface)


                    Case "StockAssumptionsInterface"
                        Dim NewSAI As New StockAssumptionsInterface(SetModelID, resolvedGSID, SetCSID)
                        NewSAI.Tag = documentTag
                        DocumentManagerAssumptions.View.AddDocument(NewSAI)
                        DocumentManagerAssumptions.View.ActivateDocument(NewSAI)

                    Case "BP_Dashboard"
                        Dim NewSAI As New BPDashboard(MyModelID)
                        NewSAI.Tag = documentTag
                        DocumentManagerAssumptions.View.AddDocument(NewSAI)
                        DocumentManagerAssumptions.View.ActivateDocument(NewSAI)

                    Case "FundingDashboard"
                        Dim NewSAI As New FundingDashboard(MyModelID)
                        NewSAI.Tag = documentTag
                        DocumentManagerAssumptions.View.AddDocument(NewSAI)
                        DocumentManagerAssumptions.View.ActivateDocument(NewSAI)

                    Case "BPIncomeExpenditureAnalyser"
                        ExcelModels(MyModelID).ResourceRegistry.ReleaseCurrent(
                            ModelResourceKeys.TransactionalRecordsRangeDataSource)

                        Dim NewSAI As New BPIncomeExpenditureAnalyser(MyModelID, Me)
                        NewSAI.Tag = documentTag
                        ExcelModels(MyModelID).ExpendAnalyser = NewSAI
                        RegisterAnalysisV1(NewSAI)

                        Try
                            DocumentManagerAssumptions.View.AddDocument(NewSAI)
                            DocumentManagerAssumptions.View.ActivateDocument(NewSAI)
                        Catch
                            ExcelModels(MyModelID).ResourceRegistry.ReleaseCurrent(
                                ModelResourceKeys.TransactionalRecordsRangeDataSource)
                            Throw
                        End Try

                    Case "BPIncomeExpenditureAnalyserV2"
                        ExcelModels(MyModelID).ResourceRegistry.ReleaseCurrent(
                            ModelResourceKeys.TransactionalRecordsRangeDataSource)

                        Using Activity As New FormSplashScreen(
                            Me, "Opening analysis", "Calculating the analyser datasource...")
                            Dim NewSAI As New BPIncomeExpenditureAnalyserV2(MyModelID, Me)
                            NewSAI.Tag = documentTag
                            ExcelModels(MyModelID).ExpendAnalyserV2 = NewSAI
                            RegisterAnalysisV2(NewSAI)

                            Try
                                Activity.Update("Binding the analyser grids...")
                                DocumentManagerAssumptions.View.AddDocument(NewSAI)
                                DocumentManagerAssumptions.View.ActivateDocument(NewSAI)
                                Activity.Complete("Analysis ready.")
                            Catch
                                ExcelModels(MyModelID).ResourceRegistry.ReleaseCurrent(
                                    ModelResourceKeys.TransactionalRecordsRangeDataSource)
                                Throw
                            End Try
                        End Using

                    Case "WebInterface"
                        Dim NewSAI As New WebInterfaceTemplate(MyModelID, resolvedGSID, SetCSID)
                        NewSAI.Tag = documentTag
                        DocumentManagerAssumptions.View.AddDocument(NewSAI)
                        DocumentManagerAssumptions.View.ActivateDocument(NewSAI)

                End Select

            Else

                'Open standard interface
                DataInterfaceCount += 1
                ReDim Preserve DataInterfaces(DataInterfaceCount)

                DataInterfaces(DataInterfaceCount) = New DataInterfaceTemplate(SetModelID, resolvedGSID, SetCSID, Me, InterfaceMode, Interfacelink)

                DataInterfaces(DataInterfaceCount).Tag = documentTag
                DocumentManagerAssumptions.View.AddDocument(DataInterfaces(DataInterfaceCount))
                DocumentManagerAssumptions.View.ActivateDocument(DataInterfaces(DataInterfaceCount))
                ActiveInterface = DataInterfaces(DataInterfaceCount)

            End If

            Me.BarStaticItemDescription.Caption = " " & ExcelModels(SetModelID).WBStructure.CompanyName & " • " & If(IsCombined, MyName & " • " & groupName, MyName) & " • " & ExcelModels(SetModelID).WBStructure.GroupStructures(resolvedGSID).ResolveChildStructure(SetCSID).CSName
            Me.Text = Me.BarStaticItemDescription.Caption



            Me.Cursor = Cursors.Default

        End If

        RecordInterfaceVisit(SetModelID, resolvedGSID, SetCSID, ShowSpecial, SpecialData)
    End Sub

    Private Sub RecordInterfaceVisit(ByVal setModelID As Integer,
                                     ByVal targetGSID As Integer,
                                     ByVal setCSID As Integer,
                                     ByVal showSpecial As Boolean,
                                     ByVal specialData As String)
        Try
            If ExcelModels(setModelID).InterfaceHistory Is Nothing Then Return

            Dim child As ChildStructure =
                ExcelModels(setModelID).WBStructure.GroupStructures(targetGSID).
                    ResolveChildStructure(setCSID)
            If child Is Nothing Then Return

            ExcelModels(setModelID).InterfaceHistory.RecordGroupInterface(
                Me,
                targetGSID,
                setCSID,
                child.CSName,
                If(IsCombined,
                   MyName & " / " & ExcelModels(setModelID).WBStructure.GroupStructures(targetGSID).GSName,
                   MyName),
                showSpecial,
                specialData)
        Catch ex As Exception
            Debug.WriteLine("Interface history record failed: " & ex.ToString())
        End Try
    End Sub
    Private Sub RightPanelButtonClick(sender As Object, e As ContextItemClickEventArgs) Handles AccordionControlSum.ContextButtonClick

        'Dim ButSender As WindowsUIButton = TryCast(e.Item, DevExpress.XtraBars.Docking2010.WindowsUIButton)

        'If ButSender Is Nothing Then
        '    Return
        'End If

        'Dim tag As String = ButSender.Tag.ToString()
        'Select Case tag
        '    Case "Refresh"
        '        ' OpenBusinessPlan

        '        RefreshSummaryData()

        '        'Case "CreateNewBP"
        '        '    MsgBox("No approved template found, awaiting code signing")

        '        'Case "CompareBPs"

        '        '    MsgBox("Awaiting DevExpress Fix")
        '        '    Return

        'End Select
    End Sub
    Public Property FormBorderColor() As Color
        Get
            Return MyColourSwatch
        End Get
        Set(ByVal value As Color)
            MyColourSwatch = value
        End Set
    End Property

    Private Sub KeydownListener(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Select Case e.KeyCode

            '  Ctrl + Shift + D
            Case Keys.P And (e.Control And Not e.Alt)
                MsgBox("Print")

                '  Ctrl + Z
            Case Keys.Z And (e.Control And Not e.Shift And Not e.Alt)
                If FileManager.ExcelModels(MyModelID) IsNot Nothing AndAlso
                   FileManager.ExcelModels(MyModelID).ChangeManager IsNot Nothing Then
                    Dim result = FileManager.ExcelModels(MyModelID).ChangeManager.Undo()
                    If result.BError Then XtraMessageBox.Show(Me, result.StrResponseMessage, "Undo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    e.Handled = True
                    e.SuppressKeyPress = True
                End If

            Case Keys.Y And (e.Control And Not e.Shift And Not e.Alt),
                 Keys.Z And (e.Control And e.Shift And Not e.Alt)
                If FileManager.ExcelModels(MyModelID) IsNot Nothing AndAlso
                   FileManager.ExcelModels(MyModelID).ChangeManager IsNot Nothing Then
                    Dim result = FileManager.ExcelModels(MyModelID).ChangeManager.Redo()
                    If result.BError Then XtraMessageBox.Show(Me, result.StrResponseMessage, "Redo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    e.Handled = True
                    e.SuppressKeyPress = True
                End If

                '  Shift and F1
            Case Keys.F1 And (e.Shift And Not e.Control And Not e.Alt)
                MsgBox("Help called")

        End Select

    End Sub

    Public Sub ReloadInterface(SetCSID As Integer, Optional ByVal ShowSpecial As Boolean = False, Optional ByVal SpecialData As String = "None", Optional ByVal TargetGSID As Integer = -1)

        Dim resolvedGSID As Integer = If(IsCombined, TargetGSID, GSID)
        Dim documentKey As String = InterfaceDocumentKey(resolvedGSID, SetCSID)
        Dim DocToReload As BaseDocument = DocumentManagerAssumptions.View.Documents.FirstOrDefault(
            Function(x) x.Control IsNot Nothing AndAlso
                        String.Equals(Convert.ToString(x.Control.Tag), documentKey,
                                      StringComparison.Ordinal))

        If DocToReload IsNot Nothing Then DocToReload.Dispose()

        If ShowSpecial Then

            ShowInterface(MyModelID, SetCSID, True, SpecialData, Nothing, TargetGSID)

        Else

            ShowInterface(MyModelID, SetCSID, False, "None", Nothing, TargetGSID)

        End If

    End Sub

    Private Sub hideContainerLeft_Click(sender As Object, e As EventArgs) Handles hideContainerLeft.Click

        'hideContainerLeft.Hide()

        If DockPanelNewNavigator.Visible Then
            DockPanelNewNavigator.Hide()
        Else
            DockPanelNewNavigator.Show()
        End If

    End Sub

    Private Sub barDockControlLeft_Click(sender As Object, e As EventArgs) Handles barDockControlLeft.Click

        'hideContainerLeft.Hide()

        'If DockPanelNewNavigator.Visible Then
        '    DockPanelNewNavigator.Hide()
        'Else
        '    DockPanelNewNavigator.Show()
        'End If

    End Sub

    Private Sub DockPanelNewNavigator_ClosingPanel(sender As Object, e As DevExpress.XtraBars.Docking.DockPanelCancelEventArgs) Handles DockPanelNewNavigator.ClosingPanel
        e.Cancel = True
        DockPanelNewNavigator.Hide()
    End Sub

    Private Sub DockPanelNavigator_Expanding(sender As Object, e As DevExpress.XtraBars.Docking.DockPanelCancelEventArgs) Handles DockPanelNavigator.Expanding
        e.Cancel = True
    End Sub

    Private Sub AccordionControlElementDev_Click(sender As Object, e As EventArgs) Handles AccordionControlElementDev.Click

    End Sub

End Class
Public Class CustomFormPainterGIT
    Inherits FormPainter
    Public Sub New(ByVal owner As System.Windows.Forms.Control, ByVal provider As DevExpress.Skins.ISkinProvider)
        MyBase.New(owner, provider)
    End Sub
    Private Function GetFormBorderColor() As Color
        Dim formBorderColor = (TryCast(Owner, GroupInterfaceTemplate)).FormBorderColor
        Return formBorderColor
    End Function
    Protected Overrides Sub DrawBackground(ByVal cache As GraphicsCache)
        Dim info = GetCaptionInfo()
        Dim ee = TryCast(info, ObjectInfoArgs)
        Dim formBorderColor = GetFormBorderColor()
        cache.FillRectangle(New SolidBrush(formBorderColor), ee.Bounds)
    End Sub
    Protected Overrides Sub DrawFrameCore(ByVal cache As GraphicsCache, ByVal info As SkinElementInfo, ByVal kind As FrameKind)
        Dim formBorderColor = GetFormBorderColor()
        cache.FillRectangle(formBorderColor, info.Bounds)
    End Sub
End Class
