Imports Abovo
Imports DevExpress.Utils
Imports Abovo.AbovoAppCls
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.XtraEditors
Imports DevExpress.XtraBars.Docking2010.Views
Imports DevExpress.XtraBars.Docking2010.Views.Tabbed
Imports DevExpress.XtraBars.Docking.Helpers
Imports DevExpress.XtraLayout

Public Class FormAssumptionsTwo

    Inherits DevExpress.XtraEditors.XtraForm
    Private LrgFontSize As Integer
    Private MediumFontSize As Integer
    Private SmallFontSize As Integer
    Private ScaleFactor As Single
    Private ScaleUnits As Single
    Private LastHTMLFontSize As Integer
    Private AmMaximised As Boolean = False
    Private HTMLFontSize As Integer
    Public Sub New()
        ' This call is required by the designer.

        InitializeComponent()

        Me.LookAndFeel.UseDefaultLookAndFeel = False ' = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        Me.BarManagerAssumptions.TransparentEditorsMode = True   'LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        Me.BarAndDockingControllerAssumptions.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        BarHeaderItemDetail.Appearance.BackColor = Color.White
        BarTopBar.Appearance.BackColor = Color.White
        BarHeaderItemDetail.Appearance.BackColor = Color.White

        Me.Text = AbovoBP.BPDetails.CompanyName
        AccordionControlNavigator.LookAndFeel.UseDefaultLookAndFeel = False

        Dim myTag As String = "StockAssumptionsInterface"
        'Dim frm As New StockAssumptionsInterface()
        'frm.Tag = myTag
        'DocumentManagerAssumptions.View.AddDocument(frm)

        Me.BarStaticItemDescription.Caption = " " & AbovoBP.BPDetails.CompanyName & " • Assumptions • Stock Assumptions • Stock Numbers" '& DocumentManagerAssumptions.
        SetInitialSizes()

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

        'PictureBoxAbovoLogo.Top = ScaleUnits
        'PictureBoxAbovoLogo.Left = ScaleUnits
        DockPanelNavigator.Width = SetWidth

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




        'PictureBoxAbovoLogo.Height = CInt(PictureBoxAbovoLogo.Width * 0.483)

        'WindowsUIButtonPanelExitHelp.Left = ScaleUnits
        'GroupBoxProgramDetails.Width = SetWidth
        'XtraTabControlMainNavigator.Top = ScaleUnits
        'XtraTabControlMainNavigator.Left = PictureBoxAbovoLogo.Right + ScaleUnits
        'XtraTabControlMainNavigator.Width = Me.Width - SetWidth - (7 * ScaleUnits)
        'XtraTabControlMainNavigator.Height = Me.Height - (6 * ScaleUnits)
        'XtraTabPageMainHABP.Height = XtraTabControlMainNavigator.PageClientBounds.Height
        'WindowsUIButtonPanelOpenCompare.Left = ScaleUnits
        'WindowsUIButtonPanelOpenCompare.Top = 3 * ScaleUnits
        'GroupBoxFileActions.Left = ScaleUnits
        'GroupBoxFileActions.Top = WindowsUIButtonPanelOpenCompare.Bottom + ScaleUnits
        'GroupBoxFileActions.Width = XtraTabControlMainNavigator.Width - (2 * ScaleUnits)
        'GroupBoxFileActions.Height = XtraTabPageMainHABP.Height - WindowsUIButtonPanelExitHelp.Height - (5 * ScaleUnits)
        'WindowsUIButtonPanelExitHelp.Top = XtraTabControlMainNavigator.Bottom - WindowsUIButtonPanelExitHelp.Height
        'WindowsUIButtonPanelExitHelp.Width = SetWidth
        'WindowsUIButtonPanelExitHelp.Left = ScaleUnits
        'WebBrowserBPInfo.Top = (2 * ScaleUnits)
        'WebBrowserBPInfo.Width = GroupBoxFileActions.Width - WindowsUIButtonPanelSaveClose.Width - (2 * ScaleUnits)
        'WebBrowserBPInfo.Height = GroupBoxFileActions.Height - WindowsUIButtonPanelBPActions.Height - (4 * ScaleUnits)
        'WindowsUIButtonPanelSaveClose.Left = WebBrowserBPInfo.Right + ScaleUnits
        'WindowsUIButtonPanelSaveClose.Top = WebBrowserBPInfo.Top
        'WindowsUIButtonPanelSaveClose.Height = GroupBoxFileActions.Height
        'WindowsUIButtonPanelOpenCompare.Width = XtraTabControlMainNavigator.PageClientBounds.Width - (2 * ScaleUnits)
        'WindowsUIButtonPanelBPActions.Top = WebBrowserBPInfo.Bottom + ScaleUnits
        'WindowsUIButtonPanelBPActions.Width = WebBrowserBPInfo.Width
        'GroupBoxProgramDetails.Top = PictureBoxAbovoLogo.Bottom + ScaleUnits
        'GroupBoxProgramDetails.Left = ScaleUnits
        'GroupBoxProgramDetails.Height = WindowsUIButtonPanelExitHelp.Top - PictureBoxAbovoLogo.Bottom - (2 * ScaleUnits)
        'SetBrowserText()

    End Sub
    Sub ResizeFonts()

        ScaleFactor = Me.Width / 2100

    End Sub
    Private Sub TabbedViewDefault_QueryControl(sender As Object, e As DevExpress.XtraBars.Docking2010.Views.QueryControlEventArgs) Handles TabbedViewDefault.QueryControl

    End Sub


    Private Sub FormMainScreen_ResizeEnd(sender As Object, e As EventArgs) Handles MyBase.ResizeEnd

        ResizeFonts()
        ResizeControls()

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
                ResizeFonts()
                ResizeControls()

            End If

        End If

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        MsgBox(AccordionControlNavigator.Width.ToString)
        MsgBox(TablePanelNavigator.Width.ToString)
        MsgBox(DockPanelNavigator.Width.ToString)
        MsgBox(DockPanelNavigator_Container.Width.ToString)

    End Sub

    Private Sub FormAssumptionsTwo_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        e.Cancel = True
        Me.Hide()

    End Sub
    Private Sub AccordionControlElementStockNumbers_Click(sender As Object, e As EventArgs) Handles AccordionControlElementStockNumbers.Click

        Dim myTag As String = "StockAssumptionsInterface"
        Dim doc As BaseDocument = DocumentManagerAssumptions.View.Documents.FirstOrDefault(Function(x) x.Control.Tag.ToString() = myTag)

        If doc IsNot Nothing Then

            DocumentManagerAssumptions.View.ActivateDocument(doc.Control)

        Else

            'Dim frm As New StockAssumptionsInterface()
            'frm.Tag = myTag
            'frm.Name = "Stock Numbers"
            'DocumentManagerAssumptions.View.AddDocument(frm)
            'DocumentManagerAssumptions.View.ActivateDocument(frm)

        End If

        Me.BarStaticItemDescription.Caption = " " & AbovoBP.BPDetails.CompanyName & " • Assumptions • Stock Assumptions • Stock Numbers"
        DockPanelNavigator.HideSliding()

    End Sub
    Private Sub AccordionControlElementSpecificIncomeAssumptions_Click(sender As Object, e As EventArgs) Handles AccordionControlElementSpecificIncomeAssumptions.Click

        Dim myTag As String = "SpecificIncomeAssumptionsInterface"
        Dim doc As BaseDocument = DocumentManagerAssumptions.View.Documents.FirstOrDefault(Function(x) x.Control.Tag.ToString() = myTag)

        If doc IsNot Nothing Then

            DocumentManagerAssumptions.View.ActivateDocument(doc.Control)

        Else

            Dim frm As New SpecificIncomeAssumptionsInterface()
            frm.Tag = myTag
            frm.Name = "Specific Income"
            DocumentManagerAssumptions.View.AddDocument(frm)
            DocumentManagerAssumptions.View.ActivateDocument(frm)

        End If

        Me.BarStaticItemDescription.Caption = " " & AbovoBP.BPDetails.CompanyName & " • Assumptions • Disposal and Other Income • Specific Income"
        DockPanelNavigator.HideSliding()

    End Sub

    Private Sub AccordionControlElementLettingRateVariations_Click(sender As Object, e As EventArgs) Handles AccordionControlElementLettingRateVariations.Click

        Dim myTag As String = "InitialRateVaritationsInterface"
        Dim doc As BaseDocument = DocumentManagerAssumptions.View.Documents.FirstOrDefault(Function(x) x.Control.Tag.ToString() = myTag)

        If doc IsNot Nothing Then

            DocumentManagerAssumptions.View.ActivateDocument(doc.Control)

        Else

            Dim frm As New InitialRateVaritationsInterface()
            frm.Tag = myTag
            frm.Name = "Initial Rate Variations"
            DocumentManagerAssumptions.View.AddDocument(frm)
            DocumentManagerAssumptions.View.ActivateDocument(frm)

        End If

        Me.BarStaticItemDescription.Caption = " " & AbovoBP.BPDetails.CompanyName & " • Assumptions • Stock Assumptions • Initial Rate Variations"
        DockPanelNavigator.HideSliding()



    End Sub


End Class