Partial Public Class InterfaceAssumptionsEdit
    Inherits DevExpress.XtraEditors.XtraForm

    ''' <summary>
    ''' Required designer variable.
    ''' </summary>
    Private components As System.ComponentModel.IContainer = Nothing

    ''' <summary>
    ''' Clean up any resources being used.
    ''' </summary>
    ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso (components IsNot Nothing) Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

#Region "Windows Form Designer generated code"

    ''' <summary>
    ''' Required method for Designer support - do not modify
    ''' the contents of this method with the code editor.
    ''' </summary>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim AccordionContextButton1 As DevExpress.XtraBars.Navigation.AccordionContextButton = New DevExpress.XtraBars.Navigation.AccordionContextButton()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(InterfaceAssumptionsEdit))
        Dim AccordionContextButton2 As DevExpress.XtraBars.Navigation.AccordionContextButton = New DevExpress.XtraBars.Navigation.AccordionContextButton()
        Me.DockManagerAssumptions = New DevExpress.XtraBars.Docking.DockManager(Me.components)
        Me.hideContainerLeft = New DevExpress.XtraBars.Docking.AutoHideContainer()
        Me.DockPanelNavigator = New DevExpress.XtraBars.Docking.DockPanel()
        Me.DockPanelContainerNavigator = New DevExpress.XtraBars.Docking.ControlContainer()
        Me.AccordionControlNavigator = New DevExpress.XtraBars.Navigation.AccordionControl()
        Me.AccordionControlElement1 = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.AccordionControlElement2 = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.AccordionControlElementGroupRepMain = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.AccordionControlElement3 = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.DockPanelStockAssumptions = New DevExpress.XtraBars.Docking.DockPanel()
        Me.DockPanelContainerMain = New DevExpress.XtraBars.Docking.ControlContainer()
        Me.AccordionControlStockAssumptions = New DevExpress.XtraBars.Navigation.AccordionControl()
        Me.AccordionContentContainerInitialRateVariations = New DevExpress.XtraBars.Navigation.AccordionContentContainer()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.AccordionContentContainerStockNumbers = New DevExpress.XtraBars.Navigation.AccordionContentContainer()
        Me.GridControlStockNumbers = New DevExpress.XtraGrid.GridControl()
        Me.UnboundSourceStocks = New DevExpress.Data.UnboundSource(Me.components)
        Me.GridViewStockNumbers = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.AccordionControlElementStockNumbers = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.AccordionControlElementInitialRateVariations = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.DefaultBarAndDockingControllerAssuptions = New DevExpress.XtraBars.DefaultBarAndDockingController(Me.components)
        CType(Me.DockManagerAssumptions, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.hideContainerLeft.SuspendLayout()
        Me.DockPanelNavigator.SuspendLayout()
        Me.DockPanelContainerNavigator.SuspendLayout()
        CType(Me.AccordionControlNavigator, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DockPanelStockAssumptions.SuspendLayout()
        Me.DockPanelContainerMain.SuspendLayout()
        CType(Me.AccordionControlStockAssumptions, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.AccordionControlStockAssumptions.SuspendLayout()
        Me.AccordionContentContainerInitialRateVariations.SuspendLayout()
        Me.AccordionContentContainerStockNumbers.SuspendLayout()
        CType(Me.GridControlStockNumbers, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UnboundSourceStocks, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridViewStockNumbers, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DefaultBarAndDockingControllerAssuptions.Controller, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DockManagerAssumptions
        '
        Me.DockManagerAssumptions.AutoHideContainers.AddRange(New DevExpress.XtraBars.Docking.AutoHideContainer() {Me.hideContainerLeft})
        Me.DockManagerAssumptions.Form = Me
        Me.DockManagerAssumptions.RootPanels.AddRange(New DevExpress.XtraBars.Docking.DockPanel() {Me.DockPanelStockAssumptions})
        Me.DockManagerAssumptions.TopZIndexControls.AddRange(New String() {"DevExpress.XtraBars.BarDockControl", "DevExpress.XtraBars.StandaloneBarDockControl", "System.Windows.Forms.MenuStrip", "System.Windows.Forms.StatusStrip", "System.Windows.Forms.StatusBar", "DevExpress.XtraBars.Ribbon.RibbonStatusBar", "DevExpress.XtraBars.Ribbon.RibbonControl", "DevExpress.XtraBars.Navigation.OfficeNavigationBar", "DevExpress.XtraBars.Navigation.TileNavPane", "DevExpress.XtraBars.TabFormControl", "DevExpress.XtraBars.FluentDesignSystem.FluentDesignFormControl", "DevExpress.XtraBars.ToolbarForm.ToolbarFormControl"})
        '
        'hideContainerLeft
        '
        Me.hideContainerLeft.BackColor = System.Drawing.Color.White
        Me.hideContainerLeft.Controls.Add(Me.DockPanelNavigator)
        Me.hideContainerLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.hideContainerLeft.Location = New System.Drawing.Point(0, 0)
        Me.hideContainerLeft.Name = "hideContainerLeft"
        Me.hideContainerLeft.Size = New System.Drawing.Size(34, 971)
        '
        'DockPanelNavigator
        '
        Me.DockPanelNavigator.Appearance.BackColor = System.Drawing.Color.PowderBlue
        Me.DockPanelNavigator.Appearance.Options.UseBackColor = True
        Me.DockPanelNavigator.Controls.Add(Me.DockPanelContainerNavigator)
        Me.DockPanelNavigator.Dock = DevExpress.XtraBars.Docking.DockingStyle.Left
        Me.DockPanelNavigator.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DockPanelNavigator.ID = New System.Guid("4815b33f-7dce-4fc2-b22c-df3904a5eb9e")
        Me.DockPanelNavigator.Location = New System.Drawing.Point(0, 0)
        Me.DockPanelNavigator.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DockPanelNavigator.Name = "DockPanelNavigator"
        Me.DockPanelNavigator.Options.AllowDockAsTabbedDocument = False
        Me.DockPanelNavigator.Options.AllowDockBottom = False
        Me.DockPanelNavigator.Options.AllowDockFill = False
        Me.DockPanelNavigator.Options.AllowDockRight = False
        Me.DockPanelNavigator.Options.AllowDockTop = False
        Me.DockPanelNavigator.Options.AllowFloating = False
        Me.DockPanelNavigator.Options.FloatOnDblClick = False
        Me.DockPanelNavigator.Options.ShowCloseButton = False
        Me.DockPanelNavigator.Options.ShowMaximizeButton = False
        Me.DockPanelNavigator.Options.ShowMinimizeButton = False
        Me.DockPanelNavigator.OriginalSize = New System.Drawing.Size(445, 200)
        Me.DockPanelNavigator.Padding = New System.Windows.Forms.Padding(14, 16, 14, 16)
        Me.DockPanelNavigator.SavedDock = DevExpress.XtraBars.Docking.DockingStyle.Left
        Me.DockPanelNavigator.SavedIndex = 1
        Me.DockPanelNavigator.Size = New System.Drawing.Size(445, 971)
        Me.DockPanelNavigator.Text = "Navigation"
        Me.DockPanelNavigator.Visibility = DevExpress.XtraBars.Docking.DockVisibility.AutoHide
        '
        'DockPanelContainerNavigator
        '
        Me.DockPanelContainerNavigator.Controls.Add(Me.AccordionControlNavigator)
        Me.DockPanelContainerNavigator.Location = New System.Drawing.Point(3, 34)
        Me.DockPanelContainerNavigator.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DockPanelContainerNavigator.Name = "DockPanelContainerNavigator"
        Me.DockPanelContainerNavigator.Size = New System.Drawing.Size(434, 934)
        Me.DockPanelContainerNavigator.TabIndex = 0
        '
        'AccordionControlNavigator
        '
        Me.AccordionControlNavigator.Appearance.AccordionControl.BackColor = System.Drawing.Color.Linen
        Me.AccordionControlNavigator.Appearance.AccordionControl.Options.UseBackColor = True
        Me.AccordionControlNavigator.Appearance.Group.Default.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AccordionControlNavigator.Appearance.Group.Default.Options.UseFont = True
        Me.AccordionControlNavigator.Appearance.Group.Hovered.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Bold)
        Me.AccordionControlNavigator.Appearance.Group.Hovered.FontStyleDelta = System.Drawing.FontStyle.Bold
        Me.AccordionControlNavigator.Appearance.Group.Hovered.Options.UseFont = True
        Me.AccordionControlNavigator.Appearance.Item.Default.Font = New System.Drawing.Font("Segoe UI", 9.857143!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AccordionControlNavigator.Appearance.Item.Default.Options.UseFont = True
        Me.AccordionControlNavigator.Appearance.Item.Hovered.Font = New System.Drawing.Font("Segoe UI", 9.857143!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AccordionControlNavigator.Appearance.Item.Hovered.Options.UseFont = True
        Me.AccordionControlNavigator.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AccordionControlNavigator.Elements.AddRange(New DevExpress.XtraBars.Navigation.AccordionControlElement() {Me.AccordionControlElement1, Me.AccordionControlElementGroupRepMain})
        Me.AccordionControlNavigator.Location = New System.Drawing.Point(0, 0)
        Me.AccordionControlNavigator.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.AccordionControlNavigator.Name = "AccordionControlNavigator"
        Me.AccordionControlNavigator.Size = New System.Drawing.Size(434, 934)
        Me.AccordionControlNavigator.TabIndex = 0
        '
        'AccordionControlElement1
        '
        Me.AccordionControlElement1.Appearance.Default.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AccordionControlElement1.Appearance.Default.FontStyleDelta = System.Drawing.FontStyle.Bold
        Me.AccordionControlElement1.Appearance.Default.Options.UseFont = True
        Me.AccordionControlElement1.Elements.AddRange(New DevExpress.XtraBars.Navigation.AccordionControlElement() {Me.AccordionControlElement2})
        Me.AccordionControlElement1.Expanded = True
        Me.AccordionControlElement1.Name = "AccordionControlElement1"
        Me.AccordionControlElement1.Text = "Stock"
        '
        'AccordionControlElement2
        '
        Me.AccordionControlElement2.Appearance.Default.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AccordionControlElement2.Appearance.Default.Options.UseFont = True
        Me.AccordionControlElement2.Name = "AccordionControlElement2"
        Me.AccordionControlElement2.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
        Me.AccordionControlElement2.Text = "Stock Numbers"
        '
        'AccordionControlElementGroupRepMain
        '
        Me.AccordionControlElementGroupRepMain.Elements.AddRange(New DevExpress.XtraBars.Navigation.AccordionControlElement() {Me.AccordionControlElement3})
        Me.AccordionControlElementGroupRepMain.Expanded = True
        Me.AccordionControlElementGroupRepMain.Name = "AccordionControlElementGroupRepMain"
        Me.AccordionControlElementGroupRepMain.Text = "Element3"
        '
        'AccordionControlElement3
        '
        Me.AccordionControlElement3.Name = "AccordionControlElement3"
        Me.AccordionControlElement3.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
        Me.AccordionControlElement3.Text = "Element3"
        '
        'DockPanelStockAssumptions
        '
        Me.DockPanelStockAssumptions.Appearance.BackColor = System.Drawing.Color.LightSteelBlue
        Me.DockPanelStockAssumptions.Appearance.BorderColor = System.Drawing.Color.LightSteelBlue
        Me.DockPanelStockAssumptions.Appearance.Options.UseBackColor = True
        Me.DockPanelStockAssumptions.Appearance.Options.UseBorderColor = True
        Me.DockPanelStockAssumptions.Appearance.Options.UseFont = True
        Me.DockPanelStockAssumptions.Controls.Add(Me.DockPanelContainerMain)
        Me.DockPanelStockAssumptions.Dock = DevExpress.XtraBars.Docking.DockingStyle.Fill
        Me.DockPanelStockAssumptions.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DockPanelStockAssumptions.ID = New System.Guid("a4437271-bba7-4294-b505-4e499ca7c872")
        Me.DockPanelStockAssumptions.Location = New System.Drawing.Point(34, 0)
        Me.DockPanelStockAssumptions.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DockPanelStockAssumptions.Name = "DockPanelStockAssumptions"
        Me.DockPanelStockAssumptions.Options.AllowDockAsTabbedDocument = False
        Me.DockPanelStockAssumptions.Options.ShowAutoHideButton = False
        Me.DockPanelStockAssumptions.Options.ShowCloseButton = False
        Me.DockPanelStockAssumptions.OriginalSize = New System.Drawing.Size(2336, 200)
        Me.DockPanelStockAssumptions.Padding = New System.Windows.Forms.Padding(14, 16, 14, 16)
        Me.DockPanelStockAssumptions.Size = New System.Drawing.Size(1662, 971)
        Me.DockPanelStockAssumptions.Text = "Stock assumptions"
        '
        'DockPanelContainerMain
        '
        Me.DockPanelContainerMain.BackColor = System.Drawing.Color.Linen
        Me.DockPanelContainerMain.Controls.Add(Me.AccordionControlStockAssumptions)
        Me.DockPanelContainerMain.Location = New System.Drawing.Point(3, 34)
        Me.DockPanelContainerMain.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DockPanelContainerMain.Name = "DockPanelContainerMain"
        Me.DockPanelContainerMain.Padding = New System.Windows.Forms.Padding(10)
        Me.DockPanelContainerMain.Size = New System.Drawing.Size(1656, 934)
        Me.DockPanelContainerMain.TabIndex = 0
        '
        'AccordionControlStockAssumptions
        '
        Me.AccordionControlStockAssumptions.Controls.Add(Me.AccordionContentContainerInitialRateVariations)
        Me.AccordionControlStockAssumptions.Controls.Add(Me.AccordionContentContainerStockNumbers)
        Me.AccordionControlStockAssumptions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AccordionControlStockAssumptions.Elements.AddRange(New DevExpress.XtraBars.Navigation.AccordionControlElement() {Me.AccordionControlElementStockNumbers, Me.AccordionControlElementInitialRateVariations})
        Me.AccordionControlStockAssumptions.ExpandElementMode = DevExpress.XtraBars.Navigation.ExpandElementMode.Multiple
        Me.AccordionControlStockAssumptions.Location = New System.Drawing.Point(10, 10)
        Me.AccordionControlStockAssumptions.LookAndFeel.UseDefaultLookAndFeel = False
        Me.AccordionControlStockAssumptions.Margin = New System.Windows.Forms.Padding(20)
        Me.AccordionControlStockAssumptions.Name = "AccordionControlStockAssumptions"
        Me.AccordionControlStockAssumptions.Padding = New System.Windows.Forms.Padding(50)
        Me.AccordionControlStockAssumptions.ScrollBarMode = DevExpress.XtraBars.Navigation.ScrollBarMode.[Auto]
        Me.AccordionControlStockAssumptions.Size = New System.Drawing.Size(1636, 914)
        Me.AccordionControlStockAssumptions.TabIndex = 0
        '
        'AccordionContentContainerInitialRateVariations
        '
        Me.AccordionContentContainerInitialRateVariations.Controls.Add(Me.Label1)
        Me.AccordionContentContainerInitialRateVariations.Name = "AccordionContentContainerInitialRateVariations"
        Me.AccordionContentContainerInitialRateVariations.Size = New System.Drawing.Size(1617, 423)
        Me.AccordionContentContainerInitialRateVariations.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(792, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Label1"
        '
        'AccordionContentContainerStockNumbers
        '
        Me.AccordionContentContainerStockNumbers.Appearance.BackColor = System.Drawing.Color.White
        Me.AccordionContentContainerStockNumbers.Appearance.Options.UseBackColor = True
        Me.AccordionContentContainerStockNumbers.Controls.Add(Me.GridControlStockNumbers)
        Me.AccordionContentContainerStockNumbers.Name = "AccordionContentContainerStockNumbers"
        Me.AccordionContentContainerStockNumbers.Padding = New System.Windows.Forms.Padding(15)
        Me.AccordionContentContainerStockNumbers.Size = New System.Drawing.Size(1617, 590)
        Me.AccordionContentContainerStockNumbers.TabIndex = 2
        '
        'GridControlStockNumbers
        '
        Me.GridControlStockNumbers.DataSource = Me.UnboundSourceStocks
        Me.GridControlStockNumbers.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GridControlStockNumbers.Location = New System.Drawing.Point(15, 15)
        Me.GridControlStockNumbers.LookAndFeel.UseDefaultLookAndFeel = False
        Me.GridControlStockNumbers.MainView = Me.GridViewStockNumbers
        Me.GridControlStockNumbers.Margin = New System.Windows.Forms.Padding(50)
        Me.GridControlStockNumbers.Name = "GridControlStockNumbers"
        Me.GridControlStockNumbers.Padding = New System.Windows.Forms.Padding(50)
        Me.GridControlStockNumbers.Size = New System.Drawing.Size(1587, 560)
        Me.GridControlStockNumbers.TabIndex = 0
        Me.GridControlStockNumbers.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridViewStockNumbers})
        '
        'GridViewStockNumbers
        '
        Me.GridViewStockNumbers.Appearance.HeaderPanel.BackColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.HeaderPanel.Options.UseBackColor = True
        Me.GridViewStockNumbers.GridControl = Me.GridControlStockNumbers
        Me.GridViewStockNumbers.Name = "GridViewStockNumbers"
        Me.GridViewStockNumbers.OptionsView.ShowGroupPanel = False
        '
        'AccordionControlElementStockNumbers
        '
        Me.AccordionControlElementStockNumbers.Appearance.Default.BackColor = System.Drawing.Color.White
        Me.AccordionControlElementStockNumbers.Appearance.Default.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.AccordionControlElementStockNumbers.Appearance.Default.Options.UseBackColor = True
        Me.AccordionControlElementStockNumbers.Appearance.Default.Options.UseFont = True
        Me.AccordionControlElementStockNumbers.Appearance.Disabled.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.AccordionControlElementStockNumbers.Appearance.Disabled.Options.UseFont = True
        Me.AccordionControlElementStockNumbers.Appearance.Hovered.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Bold)
        Me.AccordionControlElementStockNumbers.Appearance.Hovered.FontStyleDelta = System.Drawing.FontStyle.Bold
        Me.AccordionControlElementStockNumbers.Appearance.Hovered.Options.UseFont = True
        Me.AccordionControlElementStockNumbers.Appearance.Normal.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.AccordionControlElementStockNumbers.Appearance.Normal.Options.UseFont = True
        Me.AccordionControlElementStockNumbers.Appearance.Pressed.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Bold)
        Me.AccordionControlElementStockNumbers.Appearance.Pressed.FontStyleDelta = System.Drawing.FontStyle.Bold
        Me.AccordionControlElementStockNumbers.Appearance.Pressed.Options.UseFont = True
        Me.AccordionControlElementStockNumbers.ContentContainer = Me.AccordionContentContainerStockNumbers
        AccordionContextButton1.AlignmentOptions.Panel = DevExpress.Utils.ContextItemPanel.Center
        AccordionContextButton1.AlignmentOptions.Position = DevExpress.Utils.ContextItemPosition.Far
        AccordionContextButton1.Caption = "Grid"
        AccordionContextButton1.Id = New System.Guid("4ecd2561-95ef-4864-8ec9-b8b6b7d8c162")
        AccordionContextButton1.ImageOptionsCollection.ItemNormal.SvgImage = CType(resources.GetObject("resource.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        AccordionContextButton1.Name = "accordionContextButtonStockNumbersGrid"
        AccordionContextButton2.AlignmentOptions.Panel = DevExpress.Utils.ContextItemPanel.Center
        AccordionContextButton2.AlignmentOptions.Position = DevExpress.Utils.ContextItemPosition.Far
        AccordionContextButton2.Caption = "Tile"
        AccordionContextButton2.Id = New System.Guid("1e6a00df-a673-41da-a266-85c845a004b8")
        AccordionContextButton2.ImageOptionsCollection.ItemNormal.SvgImage = CType(resources.GetObject("resource.SvgImage1"), DevExpress.Utils.Svg.SvgImage)
        AccordionContextButton2.Name = "accordionContextButton1"
        Me.AccordionControlElementStockNumbers.ContextButtons.Add(AccordionContextButton1)
        Me.AccordionControlElementStockNumbers.ContextButtons.Add(AccordionContextButton2)
        Me.AccordionControlElementStockNumbers.Expanded = True
        Me.AccordionControlElementStockNumbers.Name = "AccordionControlElementStockNumbers"
        Me.AccordionControlElementStockNumbers.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
        Me.AccordionControlElementStockNumbers.Text = "Stock Numbers"
        '
        'AccordionControlElementInitialRateVariations
        '
        Me.AccordionControlElementInitialRateVariations.Appearance.Default.BackColor = System.Drawing.Color.White
        Me.AccordionControlElementInitialRateVariations.Appearance.Default.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AccordionControlElementInitialRateVariations.Appearance.Default.Options.UseBackColor = True
        Me.AccordionControlElementInitialRateVariations.Appearance.Default.Options.UseFont = True
        Me.AccordionControlElementInitialRateVariations.Appearance.Disabled.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.AccordionControlElementInitialRateVariations.Appearance.Disabled.Options.UseFont = True
        Me.AccordionControlElementInitialRateVariations.Appearance.Hovered.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Bold)
        Me.AccordionControlElementInitialRateVariations.Appearance.Hovered.FontStyleDelta = System.Drawing.FontStyle.Bold
        Me.AccordionControlElementInitialRateVariations.Appearance.Hovered.Options.UseFont = True
        Me.AccordionControlElementInitialRateVariations.Appearance.Normal.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.AccordionControlElementInitialRateVariations.Appearance.Normal.Options.UseFont = True
        Me.AccordionControlElementInitialRateVariations.Appearance.Pressed.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Bold)
        Me.AccordionControlElementInitialRateVariations.Appearance.Pressed.FontStyleDelta = System.Drawing.FontStyle.Bold
        Me.AccordionControlElementInitialRateVariations.Appearance.Pressed.Options.UseFont = True
        Me.AccordionControlElementInitialRateVariations.ContentContainer = Me.AccordionContentContainerInitialRateVariations
        Me.AccordionControlElementInitialRateVariations.Expanded = True
        Me.AccordionControlElementInitialRateVariations.Name = "AccordionControlElementInitialRateVariations"
        Me.AccordionControlElementInitialRateVariations.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
        Me.AccordionControlElementInitialRateVariations.Text = "Initial Rate Variations"
        '
        'DefaultBarAndDockingControllerAssuptions
        '
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.BackstageView.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.BackstageView.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.Button.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.Button.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.ButtonDisabled.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.ButtonDisabled.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.ButtonHover.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.ButtonHover.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.ButtonPressed.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.ButtonPressed.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.Separator.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.Separator.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.Tab.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.Tab.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.TabDisabled.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.TabDisabled.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.TabHover.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.TabHover.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.TabSelected.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBackstageView.TabSelected.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.BarAppearance.Disabled.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.BarAppearance.Disabled.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.BarAppearance.Hovered.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.BarAppearance.Hovered.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.BarAppearance.Normal.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.BarAppearance.Normal.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.BarAppearance.Pressed.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.BarAppearance.Pressed.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.Dock.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.Dock.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.ItemsFont = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.MainMenuAppearance.Disabled.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.MainMenuAppearance.Disabled.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.MainMenuAppearance.Hovered.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.MainMenuAppearance.Hovered.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.MainMenuAppearance.Normal.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.MainMenuAppearance.Normal.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.MainMenuAppearance.Pressed.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.MainMenuAppearance.Pressed.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.StatusBarAppearance.Disabled.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.StatusBarAppearance.Disabled.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.StatusBarAppearance.Hovered.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.StatusBarAppearance.Hovered.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.StatusBarAppearance.Normal.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.StatusBarAppearance.Normal.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.StatusBarAppearance.Pressed.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.StatusBarAppearance.Pressed.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.AppearanceMenu.Normal.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.AppearanceMenu.Normal.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.AppearanceMenu.Pressed.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.AppearanceMenu.Pressed.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.HeaderItemAppearance.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.HeaderItemAppearance.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.MenuBar.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.MenuBar.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.MenuCaption.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.MenuCaption.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.SideStrip.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.SideStrip.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.SideStripNonRecent.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesBar.SubMenu.SideStripNonRecent.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.ActiveTab.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.ActiveTab.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.FloatFormCaption.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.FloatFormCaption.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.FloatFormCaptionActive.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.FloatFormCaptionActive.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.HideContainer.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.HidePanelButton.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.HidePanelButton.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.HidePanelButtonActive.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.HidePanelButtonActive.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.Panel.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.PanelCaption.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.PanelCaption.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.PanelCaptionActive.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Bold)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.PanelCaptionActive.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.Tabs.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocking.Tabs.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocumentManager.View.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesDocumentManager.View.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ApplicationButton.Disabled.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ApplicationButton.Disabled.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ApplicationButton.Hovered.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ApplicationButton.Hovered.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ApplicationButton.Normal.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ApplicationButton.Normal.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ApplicationButton.Pressed.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ApplicationButton.Pressed.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.Editor.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.Editor.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.FormCaption.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.FormCaption.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.Gallery.GroupCaption.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.Gallery.GroupCaption.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.Item.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.Item.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemDescription.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemDescription.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemDescriptionDisabled.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemDescriptionDisabled.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemDescriptionHovered.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemDescriptionHovered.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemDescriptionPressed.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemDescriptionPressed.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemDisabled.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemDisabled.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemHovered.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemHovered.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemPressed.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.ItemPressed.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.PageCategory.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.PageCategory.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.PageGroupCaption.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.PageGroupCaption.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.PageHeader.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.PageHeader.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.PageHeaderHovered.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.PageHeaderHovered.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.PageHeaderSelected.Font = New System.Drawing.Font("Segoe UI", 11.14286!)
        Me.DefaultBarAndDockingControllerAssuptions.Controller.AppearancesRibbon.PageHeaderSelected.Options.UseFont = True
        Me.DefaultBarAndDockingControllerAssuptions.Controller.LookAndFeel.FormTouchUIMode = DevExpress.Utils.DefaultBoolean.[False]
        Me.DefaultBarAndDockingControllerAssuptions.Controller.LookAndFeel.UseDefaultLookAndFeel = False
        '
        'FormAssumptions
        '
        Me.Appearance.BackColor = System.Drawing.Color.White
        Me.Appearance.Options.UseBackColor = True
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1696, 971)
        Me.Controls.Add(Me.DockPanelStockAssumptions)
        Me.Controls.Add(Me.hideContainerLeft)
        Me.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LookAndFeel.TouchUIMode = DevExpress.Utils.DefaultBoolean.[True]
        Me.LookAndFeel.UseDefaultLookAndFeel = False
        Me.Margin = New System.Windows.Forms.Padding(7, 8, 7, 8)
        Me.Name = "FormAssumptions"
        Me.Text = "Assumptions"
        CType(Me.DockManagerAssumptions, System.ComponentModel.ISupportInitialize).EndInit()
        Me.hideContainerLeft.ResumeLayout(False)
        Me.DockPanelNavigator.ResumeLayout(False)
        Me.DockPanelContainerNavigator.ResumeLayout(False)
        CType(Me.AccordionControlNavigator, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DockPanelStockAssumptions.ResumeLayout(False)
        Me.DockPanelContainerMain.ResumeLayout(False)
        CType(Me.AccordionControlStockAssumptions, System.ComponentModel.ISupportInitialize).EndInit()
        Me.AccordionControlStockAssumptions.ResumeLayout(False)
        Me.AccordionContentContainerInitialRateVariations.ResumeLayout(False)
        Me.AccordionContentContainerInitialRateVariations.PerformLayout()
        Me.AccordionContentContainerStockNumbers.ResumeLayout(False)
        CType(Me.GridControlStockNumbers, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UnboundSourceStocks, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridViewStockNumbers, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DefaultBarAndDockingControllerAssuptions.Controller, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DockManagerAssumptions As DevExpress.XtraBars.Docking.DockManager
    Friend WithEvents DockPanelStockAssumptions As DevExpress.XtraBars.Docking.DockPanel
    Friend WithEvents DockPanelContainerMain As DevExpress.XtraBars.Docking.ControlContainer
    Friend WithEvents DockPanelContainerNavigator As DevExpress.XtraBars.Docking.ControlContainer
    Friend WithEvents AccordionControlNavigator As DevExpress.XtraBars.Navigation.AccordionControl
    Friend WithEvents AccordionControlElement1 As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents AccordionControlElement2 As DevExpress.XtraBars.Navigation.AccordionControlElement
    Private WithEvents DockPanelNavigator As DevExpress.XtraBars.Docking.DockPanel
    Friend WithEvents DefaultBarAndDockingControllerAssuptions As DevExpress.XtraBars.DefaultBarAndDockingController
    Friend WithEvents AccordionControlStockAssumptions As DevExpress.XtraBars.Navigation.AccordionControl
    Friend WithEvents AccordionControlElementStockNumbers As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents AccordionControlElementInitialRateVariations As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents AccordionContentContainerInitialRateVariations As DevExpress.XtraBars.Navigation.AccordionContentContainer
    Friend WithEvents AccordionContentContainerStockNumbers As DevExpress.XtraBars.Navigation.AccordionContentContainer
    Friend WithEvents GridControlStockNumbers As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridViewStockNumbers As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents AccordionControlElementGroupRepMain As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents AccordionControlElement3 As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents UnboundSourceStocks As DevExpress.Data.UnboundSource
    Friend WithEvents hideContainerLeft As DevExpress.XtraBars.Docking.AutoHideContainer
    Friend WithEvents Label1 As Label

#End Region

End Class
