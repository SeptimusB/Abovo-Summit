<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FFRForm
    Inherits DevExpress.XtraEditors.XtraForm

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then components.Dispose()
        MyBase.Dispose(disposing)
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.TopPanel = New DevExpress.XtraEditors.PanelControl()
        Me.SheetCaption = New DevExpress.XtraEditors.LabelControl()
        Me.CreateReturnButton = New DevExpress.XtraEditors.SimpleButton()
        Me.RefreshButton = New DevExpress.XtraEditors.SimpleButton()
        Me.CloseButton = New DevExpress.XtraEditors.SimpleButton()
        Me.FFRTabs = New DevExpress.XtraTab.XtraTabControl()
        CType(Me.TopPanel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TopPanel.SuspendLayout()
        CType(Me.FFRTabs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        Me.TopPanel.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.TopPanel.Controls.Add(Me.SheetCaption)
        Me.TopPanel.Controls.Add(Me.CreateReturnButton)
        Me.TopPanel.Controls.Add(Me.RefreshButton)
        Me.TopPanel.Controls.Add(Me.CloseButton)
        Me.TopPanel.Dock = System.Windows.Forms.DockStyle.Top
        Me.TopPanel.Name = "TopPanel"
        Me.TopPanel.Padding = New System.Windows.Forms.Padding(12, 8, 12, 8)
        Me.TopPanel.Size = New System.Drawing.Size(1500, 52)
        Me.SheetCaption.Appearance.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Bold)
        Me.SheetCaption.Appearance.ForeColor = System.Drawing.Color.FromArgb(32, 58, 89)
        Me.SheetCaption.Appearance.Options.UseFont = True
        Me.SheetCaption.Appearance.Options.UseForeColor = True
        Me.SheetCaption.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.SheetCaption.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SheetCaption.Name = "SheetCaption"
        Me.SheetCaption.Text = "Financial Forecast Return"
        Me.CreateReturnButton.Dock = System.Windows.Forms.DockStyle.Right
        Me.CreateReturnButton.Name = "CreateReturnButton"
        Me.CreateReturnButton.Size = New System.Drawing.Size(160, 36)
        Me.CreateReturnButton.Text = "Create FFR return"
        Me.RefreshButton.Dock = System.Windows.Forms.DockStyle.Right
        Me.RefreshButton.Name = "RefreshButton"
        Me.RefreshButton.Size = New System.Drawing.Size(100, 36)
        Me.RefreshButton.Text = "Refresh"
        Me.CloseButton.Dock = System.Windows.Forms.DockStyle.Right
        Me.CloseButton.Name = "CloseButton"
        Me.CloseButton.Size = New System.Drawing.Size(112, 36)
        Me.CloseButton.Text = "Close"
        Me.FFRTabs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FFRTabs.Name = "FFRTabs"
        Me.ClientSize = New System.Drawing.Size(1500, 900)
        Me.Controls.Add(Me.FFRTabs)
        Me.Controls.Add(Me.TopPanel)
        Me.MinimumSize = New System.Drawing.Size(1000, 650)
        Me.Name = "FFRForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Financial Forecast Return"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.TopPanel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TopPanel.ResumeLayout(False)
        CType(Me.FFRTabs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents TopPanel As DevExpress.XtraEditors.PanelControl
    Friend WithEvents SheetCaption As DevExpress.XtraEditors.LabelControl
    Friend WithEvents CreateReturnButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents RefreshButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents CloseButton As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents FFRTabs As DevExpress.XtraTab.XtraTabControl
End Class
