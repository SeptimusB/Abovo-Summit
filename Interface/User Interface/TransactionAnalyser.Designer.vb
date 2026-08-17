Imports DevExpress.Data
Imports Abovo.AbovoAppCls
Imports Abovo.CustomGrid
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class TransactionAnalyser
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
        Me.XtraTabControl1 = New DevExpress.XtraTab.XtraTabControl()
        Me.XtraTabPage1 = New DevExpress.XtraTab.XtraTabPage()
        Me.SimpleButton1 = New DevExpress.XtraEditors.SimpleButton()
        Me.GridControlSOCIData = New Abovo.CustomGrid.CustomGridControl()
        Me.GridViewAnalysis = New Abovo.CustomGrid.CustomGridView()
        Me.XtraTabPage2 = New DevExpress.XtraTab.XtraTabPage()
        Me.TimeSpanChartRangeControlClient1 = New DevExpress.XtraEditors.TimeSpanChartRangeControlClient()
        CType(Me.XtraTabControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtraTabControl1.SuspendLayout()
        Me.XtraTabPage1.SuspendLayout()
        CType(Me.GridControlSOCIData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridViewAnalysis, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TimeSpanChartRangeControlClient1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'XtraTabControl1
        '
        Me.XtraTabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.XtraTabControl1.Location = New System.Drawing.Point(0, 0)
        Me.XtraTabControl1.Name = "XtraTabControl1"
        Me.XtraTabControl1.SelectedTabPage = Me.XtraTabPage1
        Me.XtraTabControl1.Size = New System.Drawing.Size(1569, 882)
        Me.XtraTabControl1.TabIndex = 0
        Me.XtraTabControl1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtraTabPage1, Me.XtraTabPage2})
        '
        'XtraTabPage1
        '
        Me.XtraTabPage1.Controls.Add(Me.SimpleButton1)
        Me.XtraTabPage1.Controls.Add(Me.GridControlSOCIData)
        Me.XtraTabPage1.Name = "XtraTabPage1"
        Me.XtraTabPage1.Size = New System.Drawing.Size(1565, 822)
        Me.XtraTabPage1.Text = "SOCI"
        '
        'SimpleButton1
        '
        Me.SimpleButton1.Location = New System.Drawing.Point(1377, 12)
        Me.SimpleButton1.Name = "SimpleButton1"
        Me.SimpleButton1.Size = New System.Drawing.Size(131, 40)
        Me.SimpleButton1.TabIndex = 1
        Me.SimpleButton1.Text = "Refresh"
        '
        'GridControlSOCIData
        '
        Me.GridControlSOCIData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GridControlSOCIData.Location = New System.Drawing.Point(0, 0)
        Me.GridControlSOCIData.MainView = Me.GridViewAnalysis
        Me.GridControlSOCIData.Name = "GridControlSOCIData"
        Me.GridControlSOCIData.Size = New System.Drawing.Size(1565, 822)
        Me.GridControlSOCIData.TabIndex = 0
        Me.GridControlSOCIData.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridViewAnalysis})
        '
        'GridViewAnalysis
        '
        Me.GridViewAnalysis.Appearance.EvenRow.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GridViewAnalysis.Appearance.EvenRow.Options.UseBackColor = True
        Me.GridViewAnalysis.Appearance.GroupFooter.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.GridViewAnalysis.Appearance.GroupFooter.Options.UseFont = True
        Me.GridViewAnalysis.Appearance.GroupRow.BackColor = System.Drawing.Color.Lavender
        Me.GridViewAnalysis.Appearance.GroupRow.FontStyleDelta = System.Drawing.FontStyle.Bold
        Me.GridViewAnalysis.Appearance.GroupRow.Options.UseBackColor = True
        Me.GridViewAnalysis.Appearance.GroupRow.Options.UseFont = True
        Me.GridViewAnalysis.Appearance.OddRow.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.GridViewAnalysis.Appearance.OddRow.Options.UseBackColor = True
        Me.GridViewAnalysis.DetailHeight = 380
        Me.GridViewAnalysis.GridControl = Me.GridControlSOCIData
        Me.GridViewAnalysis.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always
        Me.GridViewAnalysis.Name = "GridViewAnalysis"
        '
        'XtraTabPage2
        '
        Me.XtraTabPage2.Name = "XtraTabPage2"
        Me.XtraTabPage2.Size = New System.Drawing.Size(1565, 822)
        Me.XtraTabPage2.Text = "Cashflow"
        '
        'BPIncomeExpenditureAnalyser
        '
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.XtraTabControl1)
        Me.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.Margin = New System.Windows.Forms.Padding(8, 11, 8, 11)
        Me.Name = "BPIncomeExpenditureAnalyser"
        Me.Size = New System.Drawing.Size(1569, 882)
        CType(Me.XtraTabControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtraTabControl1.ResumeLayout(False)
        Me.XtraTabPage1.ResumeLayout(False)
        CType(Me.GridControlSOCIData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridViewAnalysis, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TimeSpanChartRangeControlClient1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents XtraTabControl1 As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents XtraTabPage1 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents GridControlSOCIData As CustomGridControl
    Friend WithEvents GridViewAnalysis As CustomGridView
    Friend WithEvents XtraTabPage2 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents SimpleButton1 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents TimeSpanChartRangeControlClient1 As DevExpress.XtraEditors.TimeSpanChartRangeControlClient
End Class
