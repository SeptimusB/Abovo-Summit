Imports DevExpress.Data
Imports Abovo.AbovoAppCls
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DataInterfaceTemplateOld
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim WindowsUIButtonImageOptions1 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DataInterfaceTemplate))
        Dim WindowsUIButtonImageOptions2 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions3 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions4 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions5 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Me.TablePanelDIT = New DevExpress.Utils.Layout.TablePanel()
        Me.WindowsUIButtonPanelActions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel()
        Me.AccordionControlM = New DevExpress.XtraBars.Navigation.AccordionControl()
        Me.AccordionControlElement1 = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        CType(Me.TablePanelDIT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TablePanelDIT.SuspendLayout()
        CType(Me.AccordionControlM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TablePanelDIT
        '
        Me.TablePanelDIT.AutoScroll = True
        Me.TablePanelDIT.AutoSize = True
        Me.TablePanelDIT.Columns.AddRange(New DevExpress.Utils.Layout.TablePanelColumn() {New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 500.0!)})
        Me.TablePanelDIT.Controls.Add(Me.WindowsUIButtonPanelActions)
        Me.TablePanelDIT.Controls.Add(Me.AccordionControlM)
        Me.TablePanelDIT.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TablePanelDIT.Location = New System.Drawing.Point(0, 0)
        Me.TablePanelDIT.Name = "TablePanelDIT"
        Me.TablePanelDIT.Rows.AddRange(New DevExpress.Utils.Layout.TablePanelRow() {New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 115.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 400.0!)})
        Me.TablePanelDIT.Size = New System.Drawing.Size(1300, 1882)
        Me.TablePanelDIT.TabIndex = 0
        Me.TablePanelDIT.UseSkinIndents = True
        '
        'WindowsUIButtonPanelActions
        '
        WindowsUIButtonImageOptions1.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions1.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        WindowsUIButtonImageOptions2.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions2.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        WindowsUIButtonImageOptions3.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions3.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        WindowsUIButtonImageOptions4.SvgImage = Global.My.Resources.Resources.charttype_spline
        WindowsUIButtonImageOptions5.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions5.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.WindowsUIButtonPanelActions.Buttons.AddRange(New DevExpress.XtraEditors.ButtonPanel.IBaseButton() {New DevExpress.XtraBars.Docking2010.WindowsUIButton("", True, WindowsUIButtonImageOptions1, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "View data as spreadsheet", -1, True, Nothing, True, False, True, "Spreadsheet", -1, True), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(Nothing, True, -1, True), New DevExpress.XtraBars.Docking2010.WindowsUIButton("", True, WindowsUIButtonImageOptions2, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "", -1, True, Nothing, True, False, True, "Refresh", -1, True), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(Nothing, True, -1, True), New DevExpress.XtraBars.Docking2010.WindowsUIButton("", True, WindowsUIButtonImageOptions3, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Show Main Menu", -1, True, Nothing, True, False, True, "MainMenu", -1, True), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(), New DevExpress.XtraBars.Docking2010.WindowsUIButton("", False, WindowsUIButtonImageOptions4, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Show History", -1, True, Nothing, True, False, True, "History", -1, False), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(), New DevExpress.XtraBars.Docking2010.WindowsUIButton("", True, WindowsUIButtonImageOptions5, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Return", -1, True, Nothing, True, False, False, "Return", -1, False)})
        Me.TablePanelDIT.SetColumn(Me.WindowsUIButtonPanelActions, 0)
        Me.WindowsUIButtonPanelActions.ContentAlignment = System.Drawing.ContentAlignment.MiddleLeft
        Me.WindowsUIButtonPanelActions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WindowsUIButtonPanelActions.Location = New System.Drawing.Point(19, 18)
        Me.WindowsUIButtonPanelActions.Name = "WindowsUIButtonPanelActions"
        Me.TablePanelDIT.SetRow(Me.WindowsUIButtonPanelActions, 0)
        Me.WindowsUIButtonPanelActions.Size = New System.Drawing.Size(1262, 109)
        Me.WindowsUIButtonPanelActions.TabIndex = 3
        Me.WindowsUIButtonPanelActions.Text = "WindowsUIButtonPanel1"
        '
        'AccordionControlM
        '
        Me.TablePanelDIT.SetColumn(Me.AccordionControlM, 0)
        Me.AccordionControlM.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AccordionControlM.Elements.AddRange(New DevExpress.XtraBars.Navigation.AccordionControlElement() {Me.AccordionControlElement1})
        Me.AccordionControlM.Location = New System.Drawing.Point(19, 133)
        Me.AccordionControlM.Name = "AccordionControlM"
        Me.AccordionControlM.Padding = New System.Windows.Forms.Padding(25, 0, 0, 0)
        Me.TablePanelDIT.SetRow(Me.AccordionControlM, 1)
        Me.AccordionControlM.Size = New System.Drawing.Size(1262, 1730)
        Me.AccordionControlM.TabIndex = 2
        '
        'AccordionControlElement1
        '
        Me.AccordionControlElement1.HeaderIndent = 75
        Me.AccordionControlElement1.Height = 70
        Me.AccordionControlElement1.Name = "AccordionControlElement1"
        Me.AccordionControlElement1.Text = "Element1"
        '
        'DataInterfaceTemplate
        '
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.TablePanelDIT)
        Me.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.Margin = New System.Windows.Forms.Padding(8, 11, 8, 11)
        Me.Name = "DataInterfaceTemplate"
        Me.Size = New System.Drawing.Size(1300, 1882)
        CType(Me.TablePanelDIT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TablePanelDIT.ResumeLayout(False)
        CType(Me.AccordionControlM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TablePanelDIT As DevExpress.Utils.Layout.TablePanel
    Friend WithEvents AccordionControlM As DevExpress.XtraBars.Navigation.AccordionControl
    Friend WithEvents AccordionControlElement1 As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents WindowsUIButtonPanelActions As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel

End Class
