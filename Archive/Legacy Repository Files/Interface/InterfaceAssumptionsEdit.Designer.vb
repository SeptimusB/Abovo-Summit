<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class InterfaceAssumptionsEdit
    Inherits DevExpress.XtraEditors.XtraForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DockingContainer1 As DevExpress.XtraBars.Docking2010.Views.Tabbed.DockingContainer = New DevExpress.XtraBars.Docking2010.Views.Tabbed.DockingContainer()
        Dim WindowsUIButtonImageOptions1 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(InterfaceAssumptionsEdit))
        Dim WindowsUIButtonImageOptions2 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions3 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions4 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Me.DocumentGroup1 = New DevExpress.XtraBars.Docking2010.Views.Tabbed.DocumentGroup(Me.components)
        Me.userControl1Document = New DevExpress.XtraBars.Docking2010.Views.Tabbed.Document(Me.components)
        Me.stockAssumptionsInterfaceDocument = New DevExpress.XtraBars.Docking2010.Views.Tabbed.Document(Me.components)
        Me.DocumentManagerAssumptionsForm = New DevExpress.XtraBars.Docking2010.DocumentManager(Me.components)
        Me.XtraUserControlDocuments = New DevExpress.XtraEditors.XtraUserControl()
        Me.TabbedViewAssumptions = New DevExpress.XtraBars.Docking2010.Views.Tabbed.TabbedView(Me.components)
        Me.TablePanelAssumptions = New DevExpress.Utils.Layout.TablePanel()
        Me.WindowsUIButtonPanelAssumptionsHeader = New DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel()
        Me.LabelControlAssumptionDescription = New DevExpress.XtraEditors.LabelControl()
        Me.DockManagerAssumptions = New DevExpress.XtraBars.Docking.DockManager(Me.components)
        Me.hideContainerRight = New DevExpress.XtraBars.Docking.AutoHideContainer()
        Me.DockPanelSummary = New DevExpress.XtraBars.Docking.DockPanel()
        Me.DockPanel2_Container = New DevExpress.XtraBars.Docking.ControlContainer()
        Me.hideContainerLeft = New DevExpress.XtraBars.Docking.AutoHideContainer()
        Me.DockPanelNavigator = New DevExpress.XtraBars.Docking.DockPanel()
        Me.DockPanel1_Container = New DevExpress.XtraBars.Docking.ControlContainer()
        CType(Me.DocumentGroup1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.userControl1Document, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.stockAssumptionsInterfaceDocument, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DocumentManagerAssumptionsForm, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TabbedViewAssumptions, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TablePanelAssumptions, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TablePanelAssumptions.SuspendLayout()
        CType(Me.DockManagerAssumptions, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.hideContainerRight.SuspendLayout()
        Me.DockPanelSummary.SuspendLayout()
        Me.hideContainerLeft.SuspendLayout()
        Me.DockPanelNavigator.SuspendLayout()
        Me.SuspendLayout()
        '
        'DocumentGroup1
        '
        Me.DocumentGroup1.Items.AddRange(New DevExpress.XtraBars.Docking2010.Views.Tabbed.Document() {Me.userControl1Document, Me.stockAssumptionsInterfaceDocument})
        '
        'userControl1Document
        '
        Me.userControl1Document.Caption = "UserControl1"
        Me.userControl1Document.ControlName = "UserControl1"
        Me.userControl1Document.ControlTypeName = "UserControl1"
        '
        'stockAssumptionsInterfaceDocument
        '
        Me.stockAssumptionsInterfaceDocument.Caption = "StockAssumptionsInterface"
        Me.stockAssumptionsInterfaceDocument.ControlName = "StockAssumptionsInterface"
        Me.stockAssumptionsInterfaceDocument.ControlTypeName = "StockAssumptionsInterface"
        '
        'DocumentManagerAssumptionsForm
        '
        Me.DocumentManagerAssumptionsForm.ContainerControl = Me.XtraUserControlDocuments
        Me.DocumentManagerAssumptionsForm.View = Me.TabbedViewAssumptions
        Me.DocumentManagerAssumptionsForm.ViewCollection.AddRange(New DevExpress.XtraBars.Docking2010.Views.BaseView() {Me.TabbedViewAssumptions})
        '
        'XtraUserControlDocuments
        '
        Me.XtraUserControlDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.XtraUserControlDocuments.Location = New System.Drawing.Point(22, 100)
        Me.XtraUserControlDocuments.Name = "XtraUserControlDocuments"
        Me.XtraUserControlDocuments.Size = New System.Drawing.Size(1860, 1309)
        Me.XtraUserControlDocuments.TabIndex = 0
        '
        'TabbedViewAssumptions
        '
        Me.TabbedViewAssumptions.DocumentGroups.AddRange(New DevExpress.XtraBars.Docking2010.Views.Tabbed.DocumentGroup() {Me.DocumentGroup1})
        Me.TabbedViewAssumptions.Documents.AddRange(New DevExpress.XtraBars.Docking2010.Views.BaseDocument() {Me.userControl1Document, Me.stockAssumptionsInterfaceDocument})
        DockingContainer1.Element = Me.DocumentGroup1
        Me.TabbedViewAssumptions.RootContainer.Nodes.AddRange(New DevExpress.XtraBars.Docking2010.Views.Tabbed.DockingContainer() {DockingContainer1})
        '
        'TablePanelAssumptions
        '
        Me.TablePanelAssumptions.Columns.AddRange(New DevExpress.Utils.Layout.TablePanelColumn() {New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 55.0!), New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 50.0!)})
        Me.TablePanelAssumptions.Controls.Add(Me.WindowsUIButtonPanelAssumptionsHeader)
        Me.TablePanelAssumptions.Controls.Add(Me.LabelControlAssumptionDescription)
        Me.TablePanelAssumptions.Controls.Add(Me.XtraUserControlDocuments)
        Me.TablePanelAssumptions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TablePanelAssumptions.Location = New System.Drawing.Point(37, 0)
        Me.TablePanelAssumptions.Name = "TablePanelAssumptions"
        Me.TablePanelAssumptions.Rows.AddRange(New DevExpress.Utils.Layout.TablePanelRow() {New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 80.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 26.0!)})
        Me.TablePanelAssumptions.Size = New System.Drawing.Size(1904, 1431)
        Me.TablePanelAssumptions.TabIndex = 1
        Me.TablePanelAssumptions.UseSkinIndents = True
        '
        'WindowsUIButtonPanelAssumptionsHeader
        '
        Me.WindowsUIButtonPanelAssumptionsHeader.AllowDrop = True
        Me.WindowsUIButtonPanelAssumptionsHeader.ButtonInterval = 25
        WindowsUIButtonImageOptions1.Image = CType(resources.GetObject("WindowsUIButtonImageOptions1.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions2.Image = CType(resources.GetObject("WindowsUIButtonImageOptions2.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions3.Image = CType(resources.GetObject("WindowsUIButtonImageOptions3.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions4.Image = CType(resources.GetObject("WindowsUIButtonImageOptions4.Image"), System.Drawing.Image)
        Me.WindowsUIButtonPanelAssumptionsHeader.Buttons.AddRange(New DevExpress.XtraEditors.ButtonPanel.IBaseButton() {New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions1, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Open Workings for thie Business Plan", -1, True, Nothing, True, False, True, Nothing, -1, False), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions2, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Open Outputs", -1, True, Nothing, True, False, True, Nothing, -1, False), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions3, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Save all assumptions", -1, True, Nothing, True, False, True, Nothing, -1, False), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions4, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Close", -1, True, Nothing, True, False, True, Nothing, -1, False)})
        Me.TablePanelAssumptions.SetColumn(Me.WindowsUIButtonPanelAssumptionsHeader, 1)
        Me.WindowsUIButtonPanelAssumptionsHeader.ContentAlignment = System.Drawing.ContentAlignment.MiddleRight
        Me.WindowsUIButtonPanelAssumptionsHeader.ForeColor = System.Drawing.Color.SteelBlue
        Me.WindowsUIButtonPanelAssumptionsHeader.Location = New System.Drawing.Point(548, 20)
        Me.WindowsUIButtonPanelAssumptionsHeader.Name = "WindowsUIButtonPanelAssumptionsHeader"
        Me.TablePanelAssumptions.SetRow(Me.WindowsUIButtonPanelAssumptionsHeader, 0)
        Me.WindowsUIButtonPanelAssumptionsHeader.Size = New System.Drawing.Size(1334, 74)
        Me.WindowsUIButtonPanelAssumptionsHeader.TabIndex = 6
        Me.WindowsUIButtonPanelAssumptionsHeader.Text = "WindowsUIButtonPanelAssumptionsHeader"
        '
        'LabelControlAssumptionDescription
        '
        Me.LabelControlAssumptionDescription.Appearance.Font = New System.Drawing.Font("Segoe UI", 14.0!)
        Me.LabelControlAssumptionDescription.Appearance.Options.UseFont = True
        Me.TablePanelAssumptions.SetColumn(Me.LabelControlAssumptionDescription, 0)
        Me.LabelControlAssumptionDescription.Location = New System.Drawing.Point(22, 34)
        Me.LabelControlAssumptionDescription.Name = "LabelControlAssumptionDescription"
        Me.TablePanelAssumptions.SetRow(Me.LabelControlAssumptionDescription, 0)
        Me.LabelControlAssumptionDescription.Size = New System.Drawing.Size(520, 45)
        Me.LabelControlAssumptionDescription.TabIndex = 4
        Me.LabelControlAssumptionDescription.Text = "LabelControlAssumptionDescription"
        '
        'DockManagerAssumptions
        '
        Me.DockManagerAssumptions.AutoHideContainers.AddRange(New DevExpress.XtraBars.Docking.AutoHideContainer() {Me.hideContainerRight, Me.hideContainerLeft})
        Me.DockManagerAssumptions.Form = Me
        Me.DockManagerAssumptions.TopZIndexControls.AddRange(New String() {"DevExpress.XtraBars.BarDockControl", "DevExpress.XtraBars.StandaloneBarDockControl", "System.Windows.Forms.MenuStrip", "System.Windows.Forms.StatusStrip", "System.Windows.Forms.StatusBar", "DevExpress.XtraBars.Ribbon.RibbonStatusBar", "DevExpress.XtraBars.Ribbon.RibbonControl", "DevExpress.XtraBars.Navigation.OfficeNavigationBar", "DevExpress.XtraBars.Navigation.TileNavPane", "DevExpress.XtraBars.TabFormControl", "DevExpress.XtraBars.FluentDesignSystem.FluentDesignFormControl", "DevExpress.XtraBars.ToolbarForm.ToolbarFormControl"})
        '
        'hideContainerRight
        '
        Me.hideContainerRight.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.hideContainerRight.Controls.Add(Me.DockPanelSummary)
        Me.hideContainerRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.hideContainerRight.Location = New System.Drawing.Point(1941, 0)
        Me.hideContainerRight.Name = "hideContainerRight"
        Me.hideContainerRight.Size = New System.Drawing.Size(37, 1431)
        '
        'DockPanelSummary
        '
        Me.DockPanelSummary.Controls.Add(Me.DockPanel2_Container)
        Me.DockPanelSummary.Dock = DevExpress.XtraBars.Docking.DockingStyle.Right
        Me.DockPanelSummary.ID = New System.Guid("7c6d66d4-018d-4e1f-a621-4155d9020720")
        Me.DockPanelSummary.Location = New System.Drawing.Point(0, 0)
        Me.DockPanelSummary.Name = "DockPanelSummary"
        Me.DockPanelSummary.OriginalSize = New System.Drawing.Size(200, 200)
        Me.DockPanelSummary.SavedDock = DevExpress.XtraBars.Docking.DockingStyle.Right
        Me.DockPanelSummary.SavedIndex = 1
        Me.DockPanelSummary.Size = New System.Drawing.Size(200, 1431)
        Me.DockPanelSummary.Text = "Summary"
        Me.DockPanelSummary.Visibility = DevExpress.XtraBars.Docking.DockVisibility.AutoHide
        '
        'DockPanel2_Container
        '
        Me.DockPanel2_Container.Location = New System.Drawing.Point(9, 45)
        Me.DockPanel2_Container.Name = "DockPanel2_Container"
        Me.DockPanel2_Container.Size = New System.Drawing.Size(186, 1381)
        Me.DockPanel2_Container.TabIndex = 0
        '
        'hideContainerLeft
        '
        Me.hideContainerLeft.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.hideContainerLeft.Controls.Add(Me.DockPanelNavigator)
        Me.hideContainerLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.hideContainerLeft.Location = New System.Drawing.Point(0, 0)
        Me.hideContainerLeft.Name = "hideContainerLeft"
        Me.hideContainerLeft.Size = New System.Drawing.Size(37, 1431)
        '
        'DockPanelNavigator
        '
        Me.DockPanelNavigator.Controls.Add(Me.DockPanel1_Container)
        Me.DockPanelNavigator.Dock = DevExpress.XtraBars.Docking.DockingStyle.Left
        Me.DockPanelNavigator.ID = New System.Guid("e0152093-0007-456b-b9e4-856eae811f5f")
        Me.DockPanelNavigator.Location = New System.Drawing.Point(0, 0)
        Me.DockPanelNavigator.Name = "DockPanelNavigator"
        Me.DockPanelNavigator.OriginalSize = New System.Drawing.Size(200, 200)
        Me.DockPanelNavigator.SavedDock = DevExpress.XtraBars.Docking.DockingStyle.Left
        Me.DockPanelNavigator.SavedIndex = 0
        Me.DockPanelNavigator.Size = New System.Drawing.Size(200, 1431)
        Me.DockPanelNavigator.Text = "Navigator"
        Me.DockPanelNavigator.Visibility = DevExpress.XtraBars.Docking.DockVisibility.AutoHide
        '
        'DockPanel1_Container
        '
        Me.DockPanel1_Container.Location = New System.Drawing.Point(5, 45)
        Me.DockPanel1_Container.Name = "DockPanel1_Container"
        Me.DockPanel1_Container.Size = New System.Drawing.Size(186, 1381)
        Me.DockPanel1_Container.TabIndex = 0
        '
        'InterfaceAssumptionsEdit
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(14.0!, 36.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1978, 1431)
        Me.Controls.Add(Me.TablePanelAssumptions)
        Me.Controls.Add(Me.hideContainerLeft)
        Me.Controls.Add(Me.hideContainerRight)
        Me.Font = New System.Drawing.Font("Segoe UI", 11.14286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "InterfaceAssumptionsEdit"
        Me.Text = "InterfaceAssumptionsEdit"
        CType(Me.DocumentGroup1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.userControl1Document, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.stockAssumptionsInterfaceDocument, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DocumentManagerAssumptionsForm, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TabbedViewAssumptions, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TablePanelAssumptions, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TablePanelAssumptions.ResumeLayout(False)
        Me.TablePanelAssumptions.PerformLayout()
        CType(Me.DockManagerAssumptions, System.ComponentModel.ISupportInitialize).EndInit()
        Me.hideContainerRight.ResumeLayout(False)
        Me.DockPanelSummary.ResumeLayout(False)
        Me.hideContainerLeft.ResumeLayout(False)
        Me.DockPanelNavigator.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DocumentManagerAssumptionsForm As DevExpress.XtraBars.Docking2010.DocumentManager
    Friend WithEvents TabbedViewAssumptions As DevExpress.XtraBars.Docking2010.Views.Tabbed.TabbedView
    Friend WithEvents TablePanelAssumptions As DevExpress.Utils.Layout.TablePanel
    Friend WithEvents XtraUserControlDocuments As DevExpress.XtraEditors.XtraUserControl
    Friend WithEvents DockManagerAssumptions As DevExpress.XtraBars.Docking.DockManager
    Friend WithEvents hideContainerRight As DevExpress.XtraBars.Docking.AutoHideContainer
    Friend WithEvents DockPanelSummary As DevExpress.XtraBars.Docking.DockPanel
    Friend WithEvents DockPanel2_Container As DevExpress.XtraBars.Docking.ControlContainer
    Friend WithEvents DockPanelNavigator As DevExpress.XtraBars.Docking.DockPanel
    Friend WithEvents DockPanel1_Container As DevExpress.XtraBars.Docking.ControlContainer
    Friend WithEvents hideContainerLeft As DevExpress.XtraBars.Docking.AutoHideContainer
    Friend WithEvents DocumentGroup1 As DevExpress.XtraBars.Docking2010.Views.Tabbed.DocumentGroup
    Friend WithEvents userControl1Document As DevExpress.XtraBars.Docking2010.Views.Tabbed.Document
    Friend WithEvents stockAssumptionsInterfaceDocument As DevExpress.XtraBars.Docking2010.Views.Tabbed.Document
    Friend WithEvents LabelControlAssumptionDescription As DevExpress.XtraEditors.LabelControl
    Friend WithEvents WindowsUIButtonPanelAssumptionsHeader As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel
End Class
