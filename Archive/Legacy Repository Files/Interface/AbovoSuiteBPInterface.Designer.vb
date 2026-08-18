

Partial Class AbovoSuiteBPInterface

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
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.XtraTabPageOther = New DevExpress.XtraTab.XtraTabPage()
        Me.XtraTabControlOther = New DevExpress.XtraTab.XtraTabControl()
        Me.XtraTabPageSOCIData = New DevExpress.XtraTab.XtraTabPage()
        Me.GridControlSOCIData = New DevExpress.XtraGrid.GridControl()
        Me.GridViewSOCIData = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.XtraTabPage13 = New DevExpress.XtraTab.XtraTabPage()
        Me.XtraTabPageFFR = New DevExpress.XtraTab.XtraTabPage()
        Me.XtraTabControl1 = New DevExpress.XtraTab.XtraTabControl()
        Me.XtraTabPageOther.SuspendLayout()
        CType(Me.XtraTabControlOther, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtraTabControlOther.SuspendLayout()
        Me.XtraTabPageSOCIData.SuspendLayout()
        CType(Me.GridControlSOCIData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridViewSOCIData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.XtraTabControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtraTabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'XtraTabPageOther
        '
        Me.XtraTabPageOther.AccessibleRole = System.Windows.Forms.AccessibleRole.Cursor
        Me.XtraTabPageOther.Appearance.Header.BackColor = System.Drawing.Color.Green
        Me.XtraTabPageOther.Appearance.Header.Font = New System.Drawing.Font("Tahoma", 9.857143!)
        Me.XtraTabPageOther.Appearance.Header.ForeColor = System.Drawing.Color.White
        Me.XtraTabPageOther.Appearance.Header.Options.UseBackColor = True
        Me.XtraTabPageOther.Appearance.Header.Options.UseFont = True
        Me.XtraTabPageOther.Appearance.Header.Options.UseForeColor = True
        Me.XtraTabPageOther.Appearance.HeaderActive.ForeColor = System.Drawing.Color.Red
        Me.XtraTabPageOther.Appearance.HeaderActive.Options.UseForeColor = True
        Me.XtraTabPageOther.Controls.Add(Me.XtraTabControlOther)
        Me.XtraTabPageOther.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.XtraTabPageOther.Name = "XtraTabPageOther"
        Me.XtraTabPageOther.Size = New System.Drawing.Size(2036, 1217)
        Me.XtraTabPageOther.Text = "Other"
        '
        'XtraTabControlOther
        '
        Me.XtraTabControlOther.Dock = System.Windows.Forms.DockStyle.Fill
        Me.XtraTabControlOther.Location = New System.Drawing.Point(0, 0)
        Me.XtraTabControlOther.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.XtraTabControlOther.Name = "XtraTabControlOther"
        Me.XtraTabControlOther.SelectedTabPage = Me.XtraTabPageSOCIData
        Me.XtraTabControlOther.Size = New System.Drawing.Size(2036, 1217)
        Me.XtraTabControlOther.TabIndex = 0
        Me.XtraTabControlOther.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtraTabPageSOCIData, Me.XtraTabPage13})
        '
        'XtraTabPageSOCIData
        '
        Me.XtraTabPageSOCIData.Controls.Add(Me.GridControlSOCIData)
        Me.XtraTabPageSOCIData.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.XtraTabPageSOCIData.Name = "XtraTabPageSOCIData"
        Me.XtraTabPageSOCIData.Size = New System.Drawing.Size(2032, 1170)
        Me.XtraTabPageSOCIData.Text = "SOCI"
        '
        'GridControlSOCIData
        '
        Me.GridControlSOCIData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GridControlSOCIData.EmbeddedNavigator.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.GridControlSOCIData.Location = New System.Drawing.Point(0, 0)
        Me.GridControlSOCIData.MainView = Me.GridViewSOCIData
        Me.GridControlSOCIData.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.GridControlSOCIData.Name = "GridControlSOCIData"
        Me.GridControlSOCIData.Size = New System.Drawing.Size(2032, 1170)
        Me.GridControlSOCIData.TabIndex = 0
        Me.GridControlSOCIData.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridViewSOCIData})
        '
        'GridViewSOCIData
        '
        Me.GridViewSOCIData.Appearance.EvenRow.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GridViewSOCIData.Appearance.EvenRow.Options.UseBackColor = True
        Me.GridViewSOCIData.Appearance.GroupRow.BackColor = System.Drawing.Color.Lavender
        Me.GridViewSOCIData.Appearance.GroupRow.FontStyleDelta = System.Drawing.FontStyle.Bold
        Me.GridViewSOCIData.Appearance.GroupRow.Options.UseBackColor = True
        Me.GridViewSOCIData.Appearance.GroupRow.Options.UseFont = True
        Me.GridViewSOCIData.Appearance.OddRow.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.GridViewSOCIData.Appearance.OddRow.Options.UseBackColor = True
        Me.GridViewSOCIData.DetailHeight = 380
        Me.GridViewSOCIData.GridControl = Me.GridControlSOCIData
        Me.GridViewSOCIData.GroupFormat = "{1} {2}"
        Me.GridViewSOCIData.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always
        Me.GridViewSOCIData.Name = "GridViewSOCIData"
        Me.GridViewSOCIData.OptionsEditForm.PopupEditFormWidth = 880
        Me.GridViewSOCIData.OptionsView.ColumnAutoWidth = False
        '
        'XtraTabPage13
        '
        Me.XtraTabPage13.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.XtraTabPage13.Name = "XtraTabPage13"
        Me.XtraTabPage13.Size = New System.Drawing.Size(2032, 1170)
        Me.XtraTabPage13.Text = "Reports"
        '
        'XtraTabPageFFR
        '
        Me.XtraTabPageFFR.Appearance.Header.BackColor = System.Drawing.Color.Blue
        Me.XtraTabPageFFR.Appearance.Header.Font = New System.Drawing.Font("Tahoma", 9.857143!)
        Me.XtraTabPageFFR.Appearance.Header.ForeColor = System.Drawing.Color.White
        Me.XtraTabPageFFR.Appearance.Header.Options.UseBackColor = True
        Me.XtraTabPageFFR.Appearance.Header.Options.UseFont = True
        Me.XtraTabPageFFR.Appearance.Header.Options.UseForeColor = True
        Me.XtraTabPageFFR.Appearance.HeaderActive.ForeColor = System.Drawing.Color.Red
        Me.XtraTabPageFFR.Appearance.HeaderActive.Options.UseForeColor = True
        Me.XtraTabPageFFR.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.XtraTabPageFFR.Name = "XtraTabPageFFR"
        Me.XtraTabPageFFR.Size = New System.Drawing.Size(2036, 1217)
        Me.XtraTabPageFFR.Text = "FFR"
        '
        'XtraTabControl1
        '
        Me.XtraTabControl1.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.XtraTabControl1.Appearance.Font = New System.Drawing.Font("Tahoma", 9.857143!)
        Me.XtraTabControl1.Appearance.ForeColor = System.Drawing.Color.White
        Me.XtraTabControl1.Appearance.Options.UseBackColor = True
        Me.XtraTabControl1.Appearance.Options.UseFont = True
        Me.XtraTabControl1.Appearance.Options.UseForeColor = True
        Me.XtraTabControl1.AppearancePage.Header.Font = New System.Drawing.Font("Tahoma", 9.857143!)
        Me.XtraTabControl1.AppearancePage.Header.ForeColor = System.Drawing.Color.Red
        Me.XtraTabControl1.AppearancePage.Header.Options.UseFont = True
        Me.XtraTabControl1.AppearancePage.Header.Options.UseForeColor = True
        Me.XtraTabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.XtraTabControl1.Location = New System.Drawing.Point(0, 0)
        Me.XtraTabControl1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.XtraTabControl1.Name = "XtraTabControl1"
        Me.XtraTabControl1.SelectedTabPage = Me.XtraTabPageFFR
        Me.XtraTabControl1.Size = New System.Drawing.Size(2040, 1267)
        Me.XtraTabControl1.TabIndex = 3
        Me.XtraTabControl1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtraTabPageFFR, Me.XtraTabPageOther})
        '
        'AbovoSuiteBPInterface
        '
        Me.Appearance.BackColor = System.Drawing.Color.White
        Me.Appearance.Options.UseBackColor = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2040, 1267)
        Me.Controls.Add(Me.XtraTabControl1)
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Name = "AbovoSuiteBPInterface"
        Me.Text = "Abovo Suite a2.54 - No file"
        Me.XtraTabPageOther.ResumeLayout(False)
        CType(Me.XtraTabControlOther, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtraTabControlOther.ResumeLayout(False)
        Me.XtraTabPageSOCIData.ResumeLayout(False)
        CType(Me.GridControlSOCIData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridViewSOCIData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.XtraTabControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtraTabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplashScreenManager1 As DevExpress.XtraSplashScreen.SplashScreenManager
    Friend WithEvents OpenFileDialog1 As OpenFileDialog

    Friend WithEvents XtraTabPageOther As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabControlOther As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents XtraTabPageSOCIData As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabPage13 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabPageFFR As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabControl1 As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents GridControlSOCIData As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridViewSOCIData As DevExpress.XtraGrid.Views.Grid.GridView
    '
    'AbovoMCMInterface
    '


End Class
