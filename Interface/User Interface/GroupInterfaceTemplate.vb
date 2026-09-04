Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.LogDebugDev
Imports Abovo.PresentationManager

Imports DevExpress.CodeParser
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
    Private SidebarRefreshTimer As Timer
    Private SidebarEventsAttached As Boolean
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

        MyName = ExcelModels(SetModelID).WBStructure.GroupStructures(GSID).GSName
        Me.Text = ExcelModels(SetModelID).WBStructure.CompanyName & " / " & MyName & " Interface"
        DockPanelNewNavigator.Text = " " & MyName & " Navigator"
        DockPanelNavigator.Text = " " & MyName & " Navigator"
        Dim myTag As String = MyName & " Interface"

        SetInitialSizes()

        ApplyStructure(SetModelID, GSID)

        LoadDefaultInterface()
        InitialiseRightSidebar()
        RefreshSummaryData()

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

                AddNavigatorItem(CS.CSName, SetTag, CS.IsMaster)

            Else

                AddNavigatorItem(CS.CSName, SetTag, CS.IsMaster, CS.GroupName)

            End If

        Next

        FinaliseNavigator()

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
                        MyModelID.ToString() & ", reason=" & refreshReason)

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

            WebBrowserBPSum.DocumentText = CompactSummaryHtml(
                ExcelModels(MyModelID).WBData.RenderIEHTMLCourceFromDR(DLList))
        Else
            WebBrowserBPSum.DocumentText =
                "<html><body><p>" & ModelDescription &
                " does not define a Business Plan dashboard summary.</p></body></html>"
        End If

        If HasFundingSummary Then
            Dim FDSList As New List(Of DevExpress.Spreadsheet.CellRange)
            Dim DataRange As DevExpress.Spreadsheet.CellRange =
                Workbook.Worksheets("Funding Assumptions").Range("E3:I5")
            FDSList.Add(DataRange)

            DataRange =
                Workbook.Worksheets("Funding Assumptions").Range("J3:N5")
            FDSList.Add(DataRange)

            WebBrowserFundSum.DocumentText = CompactSummaryHtml(
                ExcelModels(MyModelID).WBData.RenderIEHTMLCourceFromDR(FDSList))
        Else
            WebBrowserFundSum.DocumentText =
                "<html><body><p>No funding summary is defined for this " &
                ModelDescription & ".</p></body></html>"
        End If

        WebBrowserAboutHelp.DocumentText = CreateSidebarHtml(
            "<h2>abovo-summit version " & WebUtility.HtmlEncode(DecVersionNumber.ToString()) & "</h2>" &
            "<p>© 2015-" & Year(Now()).ToString() & " Abovo Business Services Limited.</p>" &
            "<p><a href='summit-help'>Open Summit Help</a></p>" &
            "<p><a href='https://www.abovo-consult.co.uk'>www.abovo-consult.co.uk</a><br>" &
            "<a href='mailto:support@abovo-consult.co.uk'>support@abovo-consult.co.uk</a></p>" &
            "<p>Built using Microsoft&reg; Excel&reg; and DevExpress.</p>")


        Dim StrFileDescription As String

        Dim FileInformation As System.IO.FileInfo = ExcelModels(MyModelID).FileInfo
        StrFileDescription = "<h2>Model details</h2>" &
            "<dl><dt>Model type</dt><dd>" & Html(ModelDescription) & "</dd>" &
            "<dt>Model name</dt><dd>" & Html(ExcelModels(MyModelID).WBStructure.CompanyName) & "</dd>" &
            "<dt>Start date</dt><dd>" & Html(ExcelModels(MyModelID).WBStructure.StartDate) & "</dd>" &
            "<dt>File</dt><dd>" & Html(ExcelModels(MyModelID).FileName) & "</dd>"
        If FileInformation IsNot Nothing Then
            StrFileDescription &= "<dt>Created</dt><dd>" & Html(FileInformation.CreationTime.ToString("g")) & "</dd>" &
                "<dt>Last accessed</dt><dd>" & Html(FileInformation.LastAccessTime.ToString("g")) & "</dd>" &
                "<dt>Size</dt><dd>" & Format((FileInformation.Length / 1000000), "###.##") & " MB</dd>"
        End If
        StrFileDescription &= "</dl>"

        WebBrowserFile.DocumentText = CreateSidebarHtml(StrFileDescription)
        If SidebarMessageView IsNot Nothing Then SidebarMessageView.RefreshMessages()
        DockPanelDetail.Text = "Summary — updated " & Now().ToString("HH:mm:ss")
        Debug.WriteLine("GroupInterfaceTemplate sidebar refresh completed. ModelID=" &
                        MyModelID.ToString() & ", reason=" & refreshReason)

    End Sub

    Private Sub CalculateSidebarWorkbook(ByVal workbook As DevExpress.Spreadsheet.IWorkbook,
                                         ByVal refreshReason As String)
        If workbook Is Nothing Then Return
        Dim previousEngine As DevExpress.Spreadsheet.CalculationEngineType =
            workbook.Options.CalculationEngineType
        Dim previousCursor As Cursor = Me.Cursor
        Dim previousUseWaitCursor As Boolean = Me.UseWaitCursor
        Try
            Me.UseWaitCursor = True
            Me.Cursor = Cursors.WaitCursor
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            workbook.Options.CalculationEngineType =
                DevExpress.Spreadsheet.CalculationEngineType.Recursive
            workbook.CalculateFull()
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
        Const compactStyle As String =
            "<style type='text/css'>" &
            "html,body{margin:0!important;padding:2px!important;font-size:8pt!important;}" &
            "table{width:auto!important;margin:0!important;border-collapse:collapse!important;}" &
            "tr{height:18px!important;min-height:18px!important;}" &
            "td,th{height:18px!important;min-height:0!important;padding:1px 3px!important;" &
            "font-size:8pt!important;line-height:1.05!important;white-space:nowrap!important;}" &
            "</style>"
        Dim headEnd As Integer = sourceHtml.IndexOf("</head>", StringComparison.OrdinalIgnoreCase)
        If headEnd >= 0 Then Return sourceHtml.Insert(headEnd, compactStyle)
        Return compactStyle & sourceHtml
    End Function

    Private Sub InitialiseRightSidebar()
        SidebarMessageView = New SystemMessageView(MyModelID) With {.Dock = DockStyle.Fill}
        AccordionContentContainerSystemMessages.Controls.Add(SidebarMessageView)

        For Each browser As WebBrowser In {WebBrowserBPSum, WebBrowserDevSum,
                                           WebBrowserFundSum, WebBrowserAboutHelp,
                                           WebBrowserFile}
            browser.AllowWebBrowserDrop = False
            browser.IsWebBrowserContextMenuEnabled = False
            browser.ScriptErrorsSuppressed = True
            browser.WebBrowserShortcutsEnabled = True
            AddHandler browser.Navigating, AddressOf SidebarBrowser_Navigating
        Next

        SidebarRefreshTimer = New Timer With {.Interval = 400}
        AddHandler SidebarRefreshTimer.Tick, AddressOf SidebarRefreshTimer_Tick
        AddHandler DockPanelDetail.CustomButtonClick, AddressOf DockPanelDetail_CustomButtonClick

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

    Private Sub RequestSidebarRefresh()
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

    Private Sub DockPanelDetail_CustomButtonClick(
        ByVal sender As Object,
        ByVal e As DevExpress.XtraBars.Docking2010.ButtonEventArgs)

        If DockPanelDetail.CustomHeaderButtons.Count > 1 AndAlso
           Object.ReferenceEquals(e.Button, DockPanelDetail.CustomHeaderButtons(1)) Then
            DockPanelDetail.Visibility = DevExpress.XtraBars.Docking.DockVisibility.AutoHide
            BeginInvoke(New MethodInvoker(
                Sub()
                    If IsDisposed OrElse Disposing Then Return
                    If DockPanelDetail.Visibility = DevExpress.XtraBars.Docking.DockVisibility.AutoHide Then
                        DockPanelDetail.HideSliding()
                    End If
                End Sub))
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

    Private Function CreateSidebarHtml(ByVal body As String) As String
        Dim fontSize As Integer = Math.Max(8, CInt(ScaleUnits * 0.85F))
        Return "<!doctype html><html><head><meta charset='utf-8'><style>" &
            "body{font-family:Segoe UI,Arial,sans-serif;font-size:" & fontSize.ToString() &
            "px;color:#333;margin:6px;background:#fff;line-height:1.15}h2{color:#075da8;font-size:1.15em;margin:0 0 6px}" &
            "a{color:#075da8}p{margin:3px 0}dl{display:grid;grid-template-columns:minmax(85px,35%) 1fr;gap:3px 6px;margin:4px 0}" &
            "dt{font-weight:600}dd{margin:0;overflow-wrap:anywhere}</style></head><body>" & body & "</body></html>"
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

    End Sub
    Sub LoadDefaultInterface()

        'Dim NewSAI As New VideoPlayer("NavGuide")
        'NewSAI.Tag = -1
        'DocumentManagerAssumptions.View.AddDocument(NewSAI)
        'DocumentManagerAssumptions.View.ActivateDocument(NewSAI)

    End Sub

    Public Sub LoadDocument(TargetID As Integer, Optional ByVal ShowSpecial As Boolean = False, Optional ByVal SpecialData As String = "None")
        If ShowSpecial Then
            ShowInterface(MyModelID, TargetID, True, SpecialData)
        Else
            ShowInterface(MyModelID, TargetID)
        End If
    End Sub

    Private Sub AccordionControlNavigator_ElementClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.Navigation.ElementClickEventArgs)

        If e.Element.Tag Is Nothing Then

            Return

        End If

        Dim ItemTag As AbovoInterfaceTag = DirectCast(e.Element.Tag, AbovoInterfaceTag)

        If ItemTag.SpecialItem = False Then

            If ItemTag.TargetID = -1 Then Return
            ShowInterface(MyModelID, ItemTag.TargetID)

        End If

        ShowInterface(MyModelID, ItemTag.TargetID, True, ItemTag.SpecialItemData)

    End Sub


    Sub SetInitialSizes()

        If Screen.PrimaryScreen.Bounds.Width < 900 Then

            Me.Width = Screen.PrimaryScreen.Bounds.Width * 0.85
            Me.Height = Screen.PrimaryScreen.Bounds.Height * 0.85

        ElseIf Screen.PrimaryScreen.Bounds.Width < 1190 Then

            Me.Width = Screen.PrimaryScreen.Bounds.Width * 0.75
            Me.Height = Screen.PrimaryScreen.Bounds.Height * 0.75

        Else

            Me.Width = Screen.PrimaryScreen.Bounds.Width * 0.65
            Me.Height = Screen.PrimaryScreen.Bounds.Height * 0.65

        End If

        ResizeFonts()
        ResizeControls()

    End Sub
    Sub ResizeControls()

        ScaleFactor = Me.Width / 2400

        Dim SetWidth As Integer = Me.Width * 0.22
        ScaleUnits = Me.Width * 0.007

        DockPanelNavigator.Width = SetWidth
        'DockManagerAssumptions.
        Me.hideContainerRightDetail.Font = GetFont("Small", Me.ScaleFactor)
        Me.BarAndDockingControllerAssumptions.AppearancesDocking.ActiveTab.Font = GetFont("Medium", Me.ScaleFactor)
        Me.BarAndDockingControllerAssumptions.AppearancesDocking.HidePanelButton.Font = GetFont("Medium", Me.ScaleFactor)
        Me.BarAndDockingControllerAssumptions.AppearancesDocking.HidePanelButtonActive.Font = GetFont("Medium", Me.ScaleFactor)
        Me.BarAndDockingControllerAssumptions.AppearancesDocking.PanelCaption.Font = GetFont("Medium", Me.ScaleFactor)
        Me.BarAndDockingControllerAssumptions.AppearancesDocking.PanelCaptionActive.Font = GetFont("Medium", Me.ScaleFactor)
        Me.BarTopBar.BarAppearance.Normal.Font = GetFont("Medium", Me.ScaleFactor)
        Me.BarStaticItemDescription.ItemAppearance.Normal.Font = GetFont("Medium", Me.ScaleFactor)
        Me.AccordionControlNavigator.Appearance.Group.Hovered.Font = GetFont("Medium", Me.ScaleFactor * 0.9)
        Me.AccordionControlNavigator.Appearance.Group.Default.Font = GetFont("Medium", Me.ScaleFactor * 0.9)
        Me.AccordionControlNavigator.Appearance.Group.Normal.Font = GetFont("Medium", Me.ScaleFactor * 0.9)
        Me.AccordionControlNavigator.Appearance.Item.Normal.Font = GetFont("Small", Me.ScaleFactor)
        Me.AccordionControlNavigator.Appearance.Item.Default.Font = GetFont("Small", Me.ScaleFactor)
        Me.AccordionControlNavigator.Appearance.Item.Hovered.Font = GetFont("Small", Me.ScaleFactor)

    End Sub
    Sub ResizeFonts()

        ScaleFactor = Me.Width / 2100

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
    Private Sub GIT_ResizeEnd(sender As Object, e As EventArgs) Handles MyBase.ResizeEnd

        ResizeFonts()
        ResizeControls()


    End Sub
    Private Sub GIT_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize



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

    Public Sub ShowInterface(SetModelID As Integer, SetCSID As Integer, Optional ByVal ShowSpecial As Boolean = False, Optional ByVal SpecialData As String = "None", Optional ByVal Interfacelink As ElementInterfaceLinkTag = Nothing)


        Dim doc As BaseDocument = DocumentManagerAssumptions.View.Documents.FirstOrDefault(Function(x) x.Control.Tag.ToString() = SetCSID)

        Dim DocCount As Integer = DocumentManagerAssumptions.View.Documents.Count

        Dim DataITemp As DataInterfaceTemplate = Nothing
        Dim BIA As BPIncomeExpenditureAnalyser = Nothing

        If doc IsNot Nothing Then

            DataITemp = TryCast(doc.Control, DataInterfaceTemplate)

            If DataITemp IsNot Nothing AndAlso Interfacelink Is Nothing Then

                DataITemp.ClearLinks()

            Else

                DataITemp.AddLink(Interfacelink)

            End If

            If DataITemp IsNot Nothing Then 'It is a DITemplate, so reactivate and deactivate others

                If DataITemp.AmActivated = False Then

                    DataITemp.Reactivate()

                    ActiveInterface = DataITemp

                    If DocCount > 1 Then

                        For Each D As BaseDocument In DocumentManagerAssumptions.View.Documents

                            If Not D.Control.Tag.ToString = SetCSID Then

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

            Me.BarStaticItemDescription.Caption = " " & ExcelModels(SetModelID).WBStructure.CompanyName & " • " & MyName & " • " & ExcelModels(SetModelID).WBStructure.GroupStructures(GSID).ResolveChildStructure(SetCSID).CSName

        Else

            If DocCount > 0 Then

                For Each D As BaseDocument In DocumentManagerAssumptions.View.Documents

                    If Not D.Control.Tag.ToString = SetCSID Then

                        Dim OtherDITemp As DataInterfaceTemplate = TryCast(D.Control, DataInterfaceTemplate)

                        If OtherDITemp IsNot Nothing Then
                            OtherDITemp.Deactivate()
                        End If

                    End If

                Next

            End If

            Me.Cursor = Cursors.WaitCursor
            DataInterfaceCount += 1
            ReDim Preserve DataInterfaces(DataInterfaceCount)

            If ShowSpecial Then

                Select Case SpecialData

                    Case "WrkSheet"

                        'ExcelModels(SetModelID).SSInterface.Tag = -1
                        'DocumentManagerAssumptions.View.AddDocument(ExcelModels(SetModelID).SSInterface)
                        'DocumentManagerAssumptions.View.ActivateDocument(ExcelModels(SetModelID).SSInterface)


                    Case "StockAssumptionsInterface"
                        Dim NewSAI As New StockAssumptionsInterface(SetModelID, GSID, SetCSID)
                        NewSAI.Tag = SetCSID
                        DocumentManagerAssumptions.View.AddDocument(NewSAI)
                        DocumentManagerAssumptions.View.ActivateDocument(NewSAI)

                    Case "BP_Dashboard"
                        Dim NewSAI As New BPDashboard(MyModelID)
                        NewSAI.Tag = SetCSID
                        DocumentManagerAssumptions.View.AddDocument(NewSAI)
                        DocumentManagerAssumptions.View.ActivateDocument(NewSAI)

                    Case "FundingDashboard"
                        Dim NewSAI As New FundingDashboard(MyModelID)
                        NewSAI.Tag = SetCSID
                        DocumentManagerAssumptions.View.AddDocument(NewSAI)
                        DocumentManagerAssumptions.View.ActivateDocument(NewSAI)

                    Case "BPIncomeExpenditureAnalyser"
                        ExcelModels(MyModelID).ResourceRegistry.ReleaseCurrent(
                            ModelResourceKeys.TransactionalRecordsRangeDataSource)

                        Dim NewSAI As New BPIncomeExpenditureAnalyser(MyModelID, Me)
                        NewSAI.Tag = SetCSID
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

                        Dim NewSAI As New BPIncomeExpenditureAnalyserV2(MyModelID, Me)
                        NewSAI.Tag = SetCSID
                        ExcelModels(MyModelID).ExpendAnalyserV2 = NewSAI
                        RegisterAnalysisV2(NewSAI)

                        Try
                            DocumentManagerAssumptions.View.AddDocument(NewSAI)
                            DocumentManagerAssumptions.View.ActivateDocument(NewSAI)
                        Catch
                            ExcelModels(MyModelID).ResourceRegistry.ReleaseCurrent(
                                ModelResourceKeys.TransactionalRecordsRangeDataSource)
                            Throw
                        End Try

                    Case "WebInterface"
                        Dim NewSAI As New WebInterfaceTemplate(MyModelID, GSID, SetCSID)
                        NewSAI.Tag = SetCSID
                        DocumentManagerAssumptions.View.AddDocument(NewSAI)
                        DocumentManagerAssumptions.View.ActivateDocument(NewSAI)

                End Select

            Else

                'Open standard interface

                DataInterfaces(DataInterfaceCount) = New DataInterfaceTemplate(SetModelID, GSID, SetCSID, Me, InterfaceMode, Interfacelink)

                DataInterfaces(DataInterfaceCount).Tag = SetCSID

            End If

            DocumentManagerAssumptions.View.AddDocument(DataInterfaces(DataInterfaceCount))
            DocumentManagerAssumptions.View.ActivateDocument(DataInterfaces(DataInterfaceCount))
            ActiveInterface = DataInterfaces(DataInterfaceCount)
            Me.BarStaticItemDescription.Caption = " " & ExcelModels(SetModelID).WBStructure.CompanyName & " • " & MyName & " • " & ExcelModels(SetModelID).WBStructure.GroupStructures(GSID).ResolveChildStructure(SetCSID).CSName
            Me.Text = " " & ExcelModels(SetModelID).WBStructure.CompanyName & " • " & MyName & " • " & ExcelModels(SetModelID).WBStructure.GroupStructures(GSID).ResolveChildStructure(SetCSID).CSName



            Me.Cursor = Cursors.Default

        End If

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

    Public Sub ReloadInterface(SetCSID As Integer, Optional ByVal ShowSpecial As Boolean = False, Optional ByVal SpecialData As String = "None")

        Dim DocToReload As BaseDocument = DocumentManagerAssumptions.View.Documents.FirstOrDefault(Function(x) x.Control.Tag.ToString() = SetCSID)

        DocToReload.Dispose()

        If ShowSpecial Then

            ShowInterface(MyModelID, SetCSID, True, SpecialData)

        Else

            ShowInterface(MyModelID, SetCSID)

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
