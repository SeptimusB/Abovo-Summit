Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.LogDebugDev
Imports Abovo.PresentationManager
Imports DevExpress.CodeParser
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

        Me.Text = AbovoBP.BPDetails.CompanyName
        AccordionControlNavigator.LookAndFeel.UseDefaultLookAndFeel = False

        MyName = ExcelModels(SetModelID).WBStructure.GroupStructures(GSID).GSName
        Me.Text = ExcelModels(SetModelID).WBStructure.CompanyName & " / " & MyName & " Interface"
        DockPanelNavigator.Text = " " & MyName & " Navigator"
        Dim myTag As String = MyName & " Interface"

        SetInitialSizes()

        ApplyStructure(SetModelID, GSID)

        LoadDefaultInterface()

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

        SystemLog("LoadingStructure")

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
    Sub AddNavigatorItem(ItemName As String, IntTag As AbovoInterfaceTag, IsMaster As String, Optional ByVal GroupName As String = "None")
        SystemLog("Adding - " & ItemName & " " & IsMaster & " G:" & GroupName)



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

        Dim NewSAI As New VideoPlayer("NavGuide")
        NewSAI.Tag = -1
        DocumentManagerAssumptions.View.AddDocument(NewSAI)
        DocumentManagerAssumptions.View.ActivateDocument(NewSAI)

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
            Debug.Print("GIT is this wide: " & Me.Width)
            'Debug.Print("DMA Active is this wide: " & DockManagerAssumptions.ActivePanel.Width)
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
    Private Sub Button1_Click(sender As Object, e As EventArgs)

        MsgBox(AccordionControlNavigator.Width.ToString)
        MsgBox(TablePanelNavigator.Width.ToString)
        MsgBox(DockPanelNavigator.Width.ToString)
        MsgBox(DockPanelNavigator_Container.Width.ToString)

    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        e.Cancel = True
        Me.Hide()
        FormMainScreen.BringToFront()

    End Sub
    Public Sub ShowSpreadsheet(ActiveSheet As String)

    End Sub
    Public Sub ShowInterface(SetModelID As Integer, SetCSID As Integer, Optional ByVal ShowSpecial As Boolean = False, Optional ByVal SpecialData As String = "None", Optional ByVal Interfacelink As ElementInterfaceLinkTag = Nothing)

        SystemLog("Interface called" & SetCSID)

        Dim doc As BaseDocument = DocumentManagerAssumptions.View.Documents.FirstOrDefault(Function(x) x.Control.Tag.ToString() = SetCSID)

        Dim DocCount As Integer = DocumentManagerAssumptions.View.Documents.Count

        Dim DataITemp As DataInterfaceTemplate = Nothing
        Dim BIA As BPIncomeExpenditureAnalyser = Nothing

        If doc IsNot Nothing Then

            DataITemp = TryCast(doc.Control, DataInterfaceTemplate)

            If Interfacelink Is Nothing Then

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

            Me.BarStaticItemDescription.Caption = " " & AbovoBP.BPDetails.CompanyName & " • " & MyName & " • " & ExcelModels(SetModelID).WBStructure.GroupStructures(GSID).ChildStructures(SetCSID).CSName

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

                    Case "BPIncomeExpenditureAnalyser"
                        Dim NewSAI As New BPIncomeExpenditureAnalyser(MyModelID, Me)
                        NewSAI.Tag = SetCSID
                        DocumentManagerAssumptions.View.AddDocument(NewSAI)
                        DocumentManagerAssumptions.View.ActivateDocument(NewSAI)
                        ExcelModels(MyModelID).ExpendAnalyser = NewSAI

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
            Me.BarStaticItemDescription.Caption = " " & AbovoBP.BPDetails.CompanyName & " • " & MyName & " • " & ExcelModels(SetModelID).WBStructure.GroupStructures(GSID).ChildStructures(SetCSID).CSName
            Me.Text = " " & AbovoBP.BPDetails.CompanyName & " • " & MyName & " • " & ExcelModels(SetModelID).WBStructure.GroupStructures(GSID).ChildStructures(SetCSID).CSName



            Me.Cursor = Cursors.Default

        End If

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

                MsgBox("Undo GIT")

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