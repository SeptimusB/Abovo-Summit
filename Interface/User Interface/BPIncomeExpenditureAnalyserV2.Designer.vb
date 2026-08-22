Imports DevExpress.Data
Imports Abovo.AbovoAppCls
Imports Abovo.CustomGrid
Imports Abovo.GeneralFunctions
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BPIncomeExpenditureAnalyserV2
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing Then
                ReleaseAnalyserResources()
                If components IsNot Nothing Then components.Dispose()
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(BPIncomeExpenditureAnalyserV2))
        Dim WindowsUIButtonImageOptions2 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions3 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions4 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions5 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Me.XtraTabControlAnalyser = New DevExpress.XtraTab.XtraTabControl()
        Me.XtraTabPageSOCIWrapped = New DevExpress.XtraTab.XtraTabPage()
        Me.XtraTabPageCFWrapped = New DevExpress.XtraTab.XtraTabPage()
        Me.XtraTabPageBSWrapped = New DevExpress.XtraTab.XtraTabPage()
        Me.TimeSpanChartRangeControlClient1 = New DevExpress.XtraEditors.TimeSpanChartRangeControlClient()
        Me.TablePanelAnalyser = New DevExpress.Utils.Layout.TablePanel()
        Me.WindowsUIButtonPanelAnalyser = New DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel()
        CType(Me.XtraTabControlAnalyser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtraTabControlAnalyser.SuspendLayout()
        CType(Me.TimeSpanChartRangeControlClient1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TablePanelAnalyser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TablePanelAnalyser.SuspendLayout()
        Me.SuspendLayout()
        '
        'XtraTabControlAnalyser
        '
        Me.XtraTabControlAnalyser.AppearancePage.Header.BackColor = System.Drawing.Color.White
        Me.XtraTabControlAnalyser.AppearancePage.Header.BorderColor = System.Drawing.Color.White
        Me.XtraTabControlAnalyser.AppearancePage.Header.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.XtraTabControlAnalyser.AppearancePage.Header.Options.UseBackColor = True
        Me.XtraTabControlAnalyser.AppearancePage.Header.Options.UseBorderColor = True
        Me.XtraTabControlAnalyser.AppearancePage.Header.Options.UseForeColor = True
        Me.XtraTabControlAnalyser.AppearancePage.HeaderActive.BackColor = System.Drawing.Color.White
        Me.XtraTabControlAnalyser.AppearancePage.HeaderActive.BorderColor = System.Drawing.Color.White
        Me.XtraTabControlAnalyser.AppearancePage.HeaderActive.ForeColor = System.Drawing.Color.Maroon
        Me.XtraTabControlAnalyser.AppearancePage.HeaderActive.Options.UseBackColor = True
        Me.XtraTabControlAnalyser.AppearancePage.HeaderActive.Options.UseBorderColor = True
        Me.XtraTabControlAnalyser.AppearancePage.HeaderActive.Options.UseForeColor = True
        Me.TablePanelAnalyser.SetColumn(Me.XtraTabControlAnalyser, 0)
        Me.XtraTabControlAnalyser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.XtraTabControlAnalyser.Location = New System.Drawing.Point(9, 65)
        Me.XtraTabControlAnalyser.Name = "XtraTabControlAnalyser"
        Me.TablePanelAnalyser.SetRow(Me.XtraTabControlAnalyser, 1)
        Me.XtraTabControlAnalyser.SelectedTabPage = Me.XtraTabPageSOCIWrapped
        Me.XtraTabControlAnalyser.Size = New System.Drawing.Size(794, 469)
        Me.XtraTabControlAnalyser.TabIndex = 0
        Me.XtraTabControlAnalyser.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtraTabPageSOCIWrapped, Me.XtraTabPageCFWrapped, Me.XtraTabPageBSWrapped})
        '
        'XtraTabPageSOCIWrapped
        '
        Me.XtraTabPageSOCIWrapped.Name = "XtraTabPageSOCIWrapped"
        Me.XtraTabPageSOCIWrapped.Size = New System.Drawing.Size(792, 434)
        Me.XtraTabPageSOCIWrapped.Text = "SOCI"
        '
        'XtraTabPageCFWrapped
        '
        Me.XtraTabPageCFWrapped.Name = "XtraTabPageCFWrapped"
        Me.XtraTabPageCFWrapped.Size = New System.Drawing.Size(792, 458)
        Me.XtraTabPageCFWrapped.Text = "Detailed Cashflow"
        '
        'XtraTabPageBSWrapped
        '
        Me.XtraTabPageBSWrapped.Name = "XtraTabPageBSWrapped"
        Me.XtraTabPageBSWrapped.Size = New System.Drawing.Size(792, 458)
        Me.XtraTabPageBSWrapped.Text = "Statement of Financial Position"
        '
        'TablePanelAnalyser
        '
        Me.TablePanelAnalyser.Columns.AddRange(New DevExpress.Utils.Layout.TablePanelColumn() {New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 200.0!)})
        Me.TablePanelAnalyser.Controls.Add(Me.WindowsUIButtonPanelAnalyser)
        Me.TablePanelAnalyser.Controls.Add(Me.XtraTabControlAnalyser)
        Me.TablePanelAnalyser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TablePanelAnalyser.Location = New System.Drawing.Point(0, 0)
        Me.TablePanelAnalyser.Name = "TablePanelAnalyser"
        Me.TablePanelAnalyser.Rows.AddRange(New DevExpress.Utils.Layout.TablePanelRow() {New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 90.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 300.0!)})
        Me.TablePanelAnalyser.Size = New System.Drawing.Size(1301, 870)
        Me.TablePanelAnalyser.TabIndex = 1
        Me.TablePanelAnalyser.UseSkinIndents = True
        '
        'WindowsUIButtonPanelAnalyser
        '
        Me.WindowsUIButtonPanelAnalyser.AppearanceButton.Normal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.WindowsUIButtonPanelAnalyser.AppearanceButton.Normal.Options.UseForeColor = True
        WindowsUIButtonImageOptions1.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions1.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        WindowsUIButtonImageOptions1.SvgImageSize = New System.Drawing.Size(16, 16)
        WindowsUIButtonImageOptions2.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions2.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        WindowsUIButtonImageOptions2.SvgImageSize = New System.Drawing.Size(16, 16)
        WindowsUIButtonImageOptions3.ImageUri.Uri = "SendXLS"
        WindowsUIButtonImageOptions3.SvgImage = Global.My.Resources.Resources.exporttoxls
        WindowsUIButtonImageOptions3.SvgImageSize = New System.Drawing.Size(16, 16)
        WindowsUIButtonImageOptions4.Image = CType(resources.GetObject("WindowsUIButtonImageOptions4.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions5.ImageUri.Uri = "SendPDF"
        WindowsUIButtonImageOptions5.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions5.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        WindowsUIButtonImageOptions5.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.WindowsUIButtonPanelAnalyser.Buttons.AddRange(New DevExpress.XtraEditors.ButtonPanel.IBaseButton() {New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions1, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Expand all rows", -1, True, Nothing, True, False, True, "ExpandAll", -1, True), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions2, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Collapse all rows", -1, True, Nothing, True, False, True, "CollapseAll", -1, True), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(Nothing, True, -1, True), New DevExpress.XtraBars.Docking2010.WindowsUIButton("", True, WindowsUIButtonImageOptions3, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Export the data to XLS document", -1, True, Nothing, True, False, True, "ExportXL", -1, True), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(Nothing, True, -1, True), New DevExpress.XtraBars.Docking2010.WindowsUIButton("", False, WindowsUIButtonImageOptions4, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "View side by side with assumptions", -1, True, Nothing, True, False, True, "SideBySide", -1, True), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(Nothing, True, -1, True), New DevExpress.XtraBars.Docking2010.WindowsUIButton("", True, WindowsUIButtonImageOptions5, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Go to main menu", -1, True, Nothing, True, False, True, "OpenHome", -1, True)})
        Me.TablePanelAnalyser.SetColumn(Me.WindowsUIButtonPanelAnalyser, 0)
        Me.WindowsUIButtonPanelAnalyser.ContentAlignment = System.Drawing.ContentAlignment.MiddleLeft
        Me.WindowsUIButtonPanelAnalyser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WindowsUIButtonPanelAnalyser.Location = New System.Drawing.Point(9, 9)
        Me.WindowsUIButtonPanelAnalyser.Name = "WindowsUIButtonPanelAnalyser"
        Me.TablePanelAnalyser.SetRow(Me.WindowsUIButtonPanelAnalyser, 0)
        Me.WindowsUIButtonPanelAnalyser.Size = New System.Drawing.Size(794, 54)
        Me.WindowsUIButtonPanelAnalyser.TabIndex = 1
        Me.WindowsUIButtonPanelAnalyser.Text = "WindowsUIButtonPanelAnalyser"
        '
        'BPIncomeExpenditureAnalyserV2
        '
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.TablePanelAnalyser)
        Me.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.Margin = New System.Windows.Forms.Padding(8, 11, 8, 11)
        Me.Name = "BPIncomeExpenditureAnalyserV2"
        Me.Size = New System.Drawing.Size(1301, 870)
        CType(Me.XtraTabControlAnalyser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtraTabControlAnalyser.ResumeLayout(False)
        CType(Me.TimeSpanChartRangeControlClient1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TablePanelAnalyser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TablePanelAnalyser.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents XtraTabControlAnalyser As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents TimeSpanChartRangeControlClient1 As DevExpress.XtraEditors.TimeSpanChartRangeControlClient
    Friend WithEvents TablePanelAnalyser As DevExpress.Utils.Layout.TablePanel
    Friend WithEvents WindowsUIButtonPanelAnalyser As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel
    Friend WithEvents XtraTabPageSOCIWrapped As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabPageCFWrapped As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabPageBSWrapped As DevExpress.XtraTab.XtraTabPage
End Class
