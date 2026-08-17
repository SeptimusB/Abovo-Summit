<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ExportForm
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
        Dim WindowsUIButtonImageOptions1 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ExportForm))
        Dim WindowsUIButtonImageOptions2 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions3 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions4 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions5 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Me.TablePanelWholeExportForm = New DevExpress.Utils.Layout.TablePanel()
        Me.WindowsUIButtonPanelSave = New DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel()
        Me.TablePanelLeftSidePanel = New DevExpress.Utils.Layout.TablePanel()
        Me.WebBrowserMessage = New System.Windows.Forms.WebBrowser()
        Me.WindowsUIButtonPanelPreview = New DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel()
        Me.MemoEditNotes = New DevExpress.XtraEditors.MemoEdit()
        Me.WindowsUIButtonPanelClose = New DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel()
        Me.CheckedListBoxElementsToExport = New DevExpress.XtraEditors.CheckedListBoxControl()
        Me.CheckedListBoxThingsToExport = New DevExpress.XtraEditors.CheckedListBoxControl()
        Me.XtraTabControlExport = New DevExpress.XtraTab.XtraTabControl()
        Me.XtraTabPageExportXLS = New DevExpress.XtraTab.XtraTabPage()
        Me.SpreadsheetControlExport = New DevExpress.XtraSpreadsheet.SpreadsheetControl()
        Me.XtraTabPageExportPDF = New DevExpress.XtraTab.XtraTabPage()
        CType(Me.TablePanelWholeExportForm, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TablePanelWholeExportForm.SuspendLayout()
        CType(Me.TablePanelLeftSidePanel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TablePanelLeftSidePanel.SuspendLayout()
        CType(Me.MemoEditNotes.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CheckedListBoxElementsToExport, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CheckedListBoxThingsToExport, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.XtraTabControlExport, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtraTabControlExport.SuspendLayout()
        Me.XtraTabPageExportXLS.SuspendLayout()
        Me.SuspendLayout()
        '
        'TablePanelWholeExportForm
        '
        Me.TablePanelWholeExportForm.AutoSize = True
        Me.TablePanelWholeExportForm.Columns.AddRange(New DevExpress.Utils.Layout.TablePanelColumn() {New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 20.6!), New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 39.4!)})
        Me.TablePanelWholeExportForm.Controls.Add(Me.WindowsUIButtonPanelSave)
        Me.TablePanelWholeExportForm.Controls.Add(Me.TablePanelLeftSidePanel)
        Me.TablePanelWholeExportForm.Controls.Add(Me.XtraTabControlExport)
        Me.TablePanelWholeExportForm.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TablePanelWholeExportForm.Location = New System.Drawing.Point(0, 0)
        Me.TablePanelWholeExportForm.Name = "TablePanelWholeExportForm"
        Me.TablePanelWholeExportForm.Rows.AddRange(New DevExpress.Utils.Layout.TablePanelRow() {New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 125.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 1200.0!)})
        Me.TablePanelWholeExportForm.Size = New System.Drawing.Size(2494, 1522)
        Me.TablePanelWholeExportForm.TabIndex = 0
        Me.TablePanelWholeExportForm.UseSkinIndents = True
        '
        'WindowsUIButtonPanelSave
        '
        Me.WindowsUIButtonPanelSave.AppearanceButton.Normal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.WindowsUIButtonPanelSave.AppearanceButton.Normal.Options.UseForeColor = True
        WindowsUIButtonImageOptions1.Image = CType(resources.GetObject("WindowsUIButtonImageOptions1.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions1.SvgImageSize = New System.Drawing.Size(16, 16)
        WindowsUIButtonImageOptions2.Image = CType(resources.GetObject("WindowsUIButtonImageOptions2.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions2.SvgImageSize = New System.Drawing.Size(16, 16)
        WindowsUIButtonImageOptions3.Image = CType(resources.GetObject("WindowsUIButtonImageOptions3.Image"), System.Drawing.Image)
        Me.WindowsUIButtonPanelSave.Buttons.AddRange(New DevExpress.XtraEditors.ButtonPanel.IBaseButton() {New DevExpress.XtraBars.Docking2010.WindowsUIButton("Save in Existing File", True, WindowsUIButtonImageOptions1, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Save in Existing File", -1, True, Nothing, True, False, True, "SaveExisting", -1, False), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Save as New Document", True, WindowsUIButtonImageOptions2, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Save as New Document", -1, True, Nothing, True, False, True, "SaveNew", -1, False), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Clear Spreadsheet", True, WindowsUIButtonImageOptions3, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "", -1, True, Nothing, True, False, False, "Clear", -1, True)})
        Me.TablePanelWholeExportForm.SetColumn(Me.WindowsUIButtonPanelSave, 1)
        Me.WindowsUIButtonPanelSave.ContentAlignment = System.Drawing.ContentAlignment.TopRight
        Me.WindowsUIButtonPanelSave.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WindowsUIButtonPanelSave.Location = New System.Drawing.Point(865, 20)
        Me.WindowsUIButtonPanelSave.Name = "WindowsUIButtonPanelSave"
        Me.TablePanelWholeExportForm.SetRow(Me.WindowsUIButtonPanelSave, 0)
        Me.WindowsUIButtonPanelSave.Size = New System.Drawing.Size(1607, 119)
        Me.WindowsUIButtonPanelSave.TabIndex = 3
        Me.WindowsUIButtonPanelSave.Text = "WindowsUIButtonPanelSave"
        '
        'TablePanelLeftSidePanel
        '
        Me.TablePanelWholeExportForm.SetColumn(Me.TablePanelLeftSidePanel, 0)
        Me.TablePanelLeftSidePanel.Columns.AddRange(New DevExpress.Utils.Layout.TablePanelColumn() {New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 55.0!)})
        Me.TablePanelLeftSidePanel.Controls.Add(Me.WebBrowserMessage)
        Me.TablePanelLeftSidePanel.Controls.Add(Me.WindowsUIButtonPanelPreview)
        Me.TablePanelLeftSidePanel.Controls.Add(Me.MemoEditNotes)
        Me.TablePanelLeftSidePanel.Controls.Add(Me.WindowsUIButtonPanelClose)
        Me.TablePanelLeftSidePanel.Controls.Add(Me.CheckedListBoxElementsToExport)
        Me.TablePanelLeftSidePanel.Controls.Add(Me.CheckedListBoxThingsToExport)
        Me.TablePanelLeftSidePanel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TablePanelLeftSidePanel.Location = New System.Drawing.Point(22, 20)
        Me.TablePanelLeftSidePanel.Name = "TablePanelLeftSidePanel"
        Me.TablePanelWholeExportForm.SetRow(Me.TablePanelLeftSidePanel, 0)
        Me.TablePanelLeftSidePanel.Rows.AddRange(New DevExpress.Utils.Layout.TablePanelRow() {New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 26.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 26.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 26.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 26.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 26.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 156.857!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 26.0!)})
        Me.TablePanelWholeExportForm.SetRowSpan(Me.TablePanelLeftSidePanel, 2)
        Me.TablePanelLeftSidePanel.Size = New System.Drawing.Size(837, 1480)
        Me.TablePanelLeftSidePanel.TabIndex = 1
        Me.TablePanelLeftSidePanel.UseSkinIndents = True
        '
        'WebBrowserMessage
        '
        Me.TablePanelLeftSidePanel.SetColumn(Me.WebBrowserMessage, 0)
        Me.WebBrowserMessage.Location = New System.Drawing.Point(22, 20)
        Me.WebBrowserMessage.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrowserMessage.Name = "WebBrowserMessage"
        Me.TablePanelLeftSidePanel.SetRow(Me.WebBrowserMessage, 0)
        Me.WebBrowserMessage.Size = New System.Drawing.Size(793, 309)
        Me.WebBrowserMessage.TabIndex = 6
        '
        'WindowsUIButtonPanelPreview
        '
        Me.WindowsUIButtonPanelPreview.AppearanceButton.Normal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.WindowsUIButtonPanelPreview.AppearanceButton.Normal.Options.UseForeColor = True
        WindowsUIButtonImageOptions4.Image = CType(resources.GetObject("WindowsUIButtonImageOptions4.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions4.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.WindowsUIButtonPanelPreview.Buttons.AddRange(New DevExpress.XtraEditors.ButtonPanel.IBaseButton() {New DevExpress.XtraBars.Docking2010.WindowsUIButton("Process and Preview the Export", True, WindowsUIButtonImageOptions4, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Process and Preview the Export", -1, True, Nothing, True, False, True, "Preview", -1, False)})
        Me.TablePanelLeftSidePanel.SetColumn(Me.WindowsUIButtonPanelPreview, 0)
        Me.WindowsUIButtonPanelPreview.Location = New System.Drawing.Point(22, 1015)
        Me.WindowsUIButtonPanelPreview.Name = "WindowsUIButtonPanelPreview"
        Me.TablePanelLeftSidePanel.SetRow(Me.WindowsUIButtonPanelPreview, 5)
        Me.WindowsUIButtonPanelPreview.Size = New System.Drawing.Size(793, 151)
        Me.WindowsUIButtonPanelPreview.TabIndex = 5
        Me.WindowsUIButtonPanelPreview.Text = "WindowsUIButtonPanelPreview"
        '
        'MemoEditNotes
        '
        Me.TablePanelLeftSidePanel.SetColumn(Me.MemoEditNotes, 0)
        Me.MemoEditNotes.Location = New System.Drawing.Point(22, 866)
        Me.MemoEditNotes.Name = "MemoEditNotes"
        Me.TablePanelLeftSidePanel.SetRow(Me.MemoEditNotes, 4)
        Me.MemoEditNotes.Size = New System.Drawing.Size(793, 143)
        Me.MemoEditNotes.TabIndex = 4
        '
        'WindowsUIButtonPanelClose
        '
        Me.WindowsUIButtonPanelClose.AppearanceButton.Normal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.WindowsUIButtonPanelClose.AppearanceButton.Normal.Options.UseForeColor = True
        WindowsUIButtonImageOptions5.Image = CType(resources.GetObject("WindowsUIButtonImageOptions5.Image"), System.Drawing.Image)
        WindowsUIButtonImageOptions5.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.WindowsUIButtonPanelClose.Buttons.AddRange(New DevExpress.XtraEditors.ButtonPanel.IBaseButton() {New DevExpress.XtraBars.Docking2010.WindowsUIButton("Close", True, WindowsUIButtonImageOptions5, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Close the Export Form", -1, True, Nothing, True, False, True, "Close", -1, False)})
        Me.TablePanelLeftSidePanel.SetColumn(Me.WindowsUIButtonPanelClose, 0)
        Me.WindowsUIButtonPanelClose.ContentAlignment = System.Drawing.ContentAlignment.BottomCenter
        Me.WindowsUIButtonPanelClose.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WindowsUIButtonPanelClose.Location = New System.Drawing.Point(22, 1172)
        Me.WindowsUIButtonPanelClose.Name = "WindowsUIButtonPanelClose"
        Me.TablePanelLeftSidePanel.SetRow(Me.WindowsUIButtonPanelClose, 6)
        Me.WindowsUIButtonPanelClose.Size = New System.Drawing.Size(793, 286)
        Me.WindowsUIButtonPanelClose.TabIndex = 2
        Me.WindowsUIButtonPanelClose.Text = "WindowsUIButtonPanelAnalyser"
        '
        'CheckedListBoxElementsToExport
        '
        Me.TablePanelLeftSidePanel.SetColumn(Me.CheckedListBoxElementsToExport, 0)
        Me.CheckedListBoxElementsToExport.Items.AddRange(New DevExpress.XtraEditors.Controls.CheckedListBoxItem() {New DevExpress.XtraEditors.Controls.CheckedListBoxItem(Nothing, "Include Subtotals", System.Windows.Forms.CheckState.Checked), New DevExpress.XtraEditors.Controls.CheckedListBoxItem(Nothing, "Don't format"), New DevExpress.XtraEditors.Controls.CheckedListBoxItem(Nothing, "Open in Microsoft Excel after export")})
        Me.CheckedListBoxElementsToExport.Location = New System.Drawing.Point(22, 514)
        Me.CheckedListBoxElementsToExport.Name = "CheckedListBoxElementsToExport"
        Me.TablePanelLeftSidePanel.SetRow(Me.CheckedListBoxElementsToExport, 2)
        Me.CheckedListBoxElementsToExport.Size = New System.Drawing.Size(793, 326)
        Me.CheckedListBoxElementsToExport.TabIndex = 1
        '
        'CheckedListBoxThingsToExport
        '
        Me.TablePanelLeftSidePanel.SetColumn(Me.CheckedListBoxThingsToExport, 0)
        Me.CheckedListBoxThingsToExport.Location = New System.Drawing.Point(22, 335)
        Me.CheckedListBoxThingsToExport.Name = "CheckedListBoxThingsToExport"
        Me.TablePanelLeftSidePanel.SetRow(Me.CheckedListBoxThingsToExport, 1)
        Me.CheckedListBoxThingsToExport.Size = New System.Drawing.Size(793, 173)
        Me.CheckedListBoxThingsToExport.TabIndex = 0
        '
        'XtraTabControlExport
        '
        Me.TablePanelWholeExportForm.SetColumn(Me.XtraTabControlExport, 1)
        Me.XtraTabControlExport.Dock = System.Windows.Forms.DockStyle.Fill
        Me.XtraTabControlExport.Location = New System.Drawing.Point(865, 145)
        Me.XtraTabControlExport.Name = "XtraTabControlExport"
        Me.TablePanelWholeExportForm.SetRow(Me.XtraTabControlExport, 1)
        Me.XtraTabControlExport.SelectedTabPage = Me.XtraTabPageExportXLS
        Me.XtraTabControlExport.Size = New System.Drawing.Size(1607, 1355)
        Me.XtraTabControlExport.TabIndex = 0
        Me.XtraTabControlExport.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtraTabPageExportXLS, Me.XtraTabPageExportPDF})
        '
        'XtraTabPageExportXLS
        '
        Me.XtraTabPageExportXLS.Controls.Add(Me.SpreadsheetControlExport)
        Me.XtraTabPageExportXLS.Name = "XtraTabPageExportXLS"
        Me.XtraTabPageExportXLS.Size = New System.Drawing.Size(1603, 1302)
        Me.XtraTabPageExportXLS.Text = "Export to XLS"
        '
        'SpreadsheetControlExport
        '
        Me.SpreadsheetControlExport.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SpreadsheetControlExport.Enabled = False
        Me.SpreadsheetControlExport.Location = New System.Drawing.Point(0, 0)
        Me.SpreadsheetControlExport.Name = "SpreadsheetControlExport"
        Me.SpreadsheetControlExport.Size = New System.Drawing.Size(1603, 1302)
        Me.SpreadsheetControlExport.TabIndex = 0
        Me.SpreadsheetControlExport.Text = "SpreadsheetControlExport"
        '
        'XtraTabPageExportPDF
        '
        Me.XtraTabPageExportPDF.Appearance.PageClient.BackColor = System.Drawing.Color.White
        Me.XtraTabPageExportPDF.Appearance.PageClient.Options.UseBackColor = True
        Me.XtraTabPageExportPDF.Name = "XtraTabPageExportPDF"
        Me.XtraTabPageExportPDF.Size = New System.Drawing.Size(1603, 1302)
        Me.XtraTabPageExportPDF.Text = "Export to PDF"
        '
        'ExportForm
        '
        Me.Appearance.BackColor = System.Drawing.Color.White
        Me.Appearance.Options.UseBackColor = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(13.0!, 31.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2494, 1522)
        Me.Controls.Add(Me.TablePanelWholeExportForm)
        Me.IconOptions.Icon = CType(resources.GetObject("ExportForm.IconOptions.Icon"), System.Drawing.Icon)
        Me.Name = "ExportForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Export Data"
        CType(Me.TablePanelWholeExportForm, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TablePanelWholeExportForm.ResumeLayout(False)
        CType(Me.TablePanelLeftSidePanel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TablePanelLeftSidePanel.ResumeLayout(False)
        CType(Me.MemoEditNotes.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CheckedListBoxElementsToExport, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CheckedListBoxThingsToExport, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.XtraTabControlExport, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtraTabControlExport.ResumeLayout(False)
        Me.XtraTabPageExportXLS.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TablePanelWholeExportForm As DevExpress.Utils.Layout.TablePanel
    Friend WithEvents XtraTabPageExportXLS As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabPageExportPDF As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabControlExport As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents TablePanelLeftSidePanel As DevExpress.Utils.Layout.TablePanel
    Friend WithEvents CheckedListBoxElementsToExport As DevExpress.XtraEditors.CheckedListBoxControl
    Friend WithEvents CheckedListBoxThingsToExport As DevExpress.XtraEditors.CheckedListBoxControl
    Friend WithEvents WindowsUIButtonPanelClose As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel
    Friend WithEvents WindowsUIButtonPanelSave As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel
    Friend WithEvents WindowsUIButtonPanelPreview As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel
    Friend WithEvents MemoEditNotes As DevExpress.XtraEditors.MemoEdit
    Friend WithEvents SpreadsheetControlExport As DevExpress.XtraSpreadsheet.SpreadsheetControl
    Friend WithEvents WebBrowserMessage As WebBrowser
End Class
