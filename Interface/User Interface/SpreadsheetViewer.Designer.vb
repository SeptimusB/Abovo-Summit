<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SpreadsheetViewer
    Inherits DevExpress.XtraEditors.XtraForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SpreadsheetViewer))
        Dim WindowsUIButtonImageOptions2 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions3 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Me.TablePanelSpreadsheetView = New DevExpress.Utils.Layout.TablePanel()
        Me.SpreadsheetControlViewer = New DevExpress.XtraSpreadsheet.SpreadsheetControl()
        Me.SpreadsheetFormulaBar1 = New DevExpress.XtraSpreadsheet.SpreadsheetFormulaBar()
        Me.WindowsUIButtonPanelSpreadsheet = New DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel()
        CType(Me.TablePanelSpreadsheetView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TablePanelSpreadsheetView.SuspendLayout()
        Me.SuspendLayout()
        '
        'TablePanelSpreadsheetView
        '
        Me.TablePanelSpreadsheetView.Columns.AddRange(New DevExpress.Utils.Layout.TablePanelColumn() {New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 55.0!)})
        Me.TablePanelSpreadsheetView.Controls.Add(Me.SpreadsheetControlViewer)
        Me.TablePanelSpreadsheetView.Controls.Add(Me.SpreadsheetFormulaBar1)
        Me.TablePanelSpreadsheetView.Controls.Add(Me.WindowsUIButtonPanelSpreadsheet)
        Me.TablePanelSpreadsheetView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TablePanelSpreadsheetView.Location = New System.Drawing.Point(0, 0)
        Me.TablePanelSpreadsheetView.Name = "TablePanelSpreadsheetView"
        Me.TablePanelSpreadsheetView.Rows.AddRange(New DevExpress.Utils.Layout.TablePanelRow() {New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 400.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 995.0007!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 75.0!)})
        Me.TablePanelSpreadsheetView.Size = New System.Drawing.Size(1288, 1071)
        Me.TablePanelSpreadsheetView.TabIndex = 0
        Me.TablePanelSpreadsheetView.UseSkinIndents = True
        '
        'SpreadsheetControlViewer
        '
        Me.TablePanelSpreadsheetView.SetColumn(Me.SpreadsheetControlViewer, 0)
        Me.SpreadsheetControlViewer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SpreadsheetControlViewer.Location = New System.Drawing.Point(19, 61)
        Me.SpreadsheetControlViewer.Name = "SpreadsheetControlViewer"
        Me.TablePanelSpreadsheetView.SetRow(Me.SpreadsheetControlViewer, 1)
        Me.SpreadsheetControlViewer.Size = New System.Drawing.Size(1250, 893)
        Me.SpreadsheetControlViewer.TabIndex = 5
        Me.SpreadsheetControlViewer.Text = "SpreadsheetControl"
        '
        'SpreadsheetFormulaBar1
        '
        Me.SpreadsheetFormulaBar1.AccessibleName = "Formula bar"
        Me.TablePanelSpreadsheetView.SetColumn(Me.SpreadsheetFormulaBar1, 0)
        Me.SpreadsheetFormulaBar1.Location = New System.Drawing.Point(21, 21)
        Me.SpreadsheetFormulaBar1.Margin = New System.Windows.Forms.Padding(5, 6, 5, 6)
        Me.SpreadsheetFormulaBar1.MinimumSize = New System.Drawing.Size(0, 31)
        Me.SpreadsheetFormulaBar1.Name = "SpreadsheetFormulaBar1"
        Me.TablePanelSpreadsheetView.SetRow(Me.SpreadsheetFormulaBar1, 0)
        Me.SpreadsheetFormulaBar1.Size = New System.Drawing.Size(1246, 31)
        Me.SpreadsheetFormulaBar1.TabIndex = 4
        '
        'WindowsUIButtonPanelSpreadsheet
        '
        Me.WindowsUIButtonPanelSpreadsheet.AppearanceButton.Normal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.WindowsUIButtonPanelSpreadsheet.AppearanceButton.Normal.Options.UseForeColor = True
        WindowsUIButtonImageOptions1.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions1.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        WindowsUIButtonImageOptions1.SvgImageSize = New System.Drawing.Size(16, 16)
        WindowsUIButtonImageOptions2.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions2.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        WindowsUIButtonImageOptions2.SvgImageSize = New System.Drawing.Size(16, 16)
        WindowsUIButtonImageOptions3.ImageUri.Uri = "SendPDF"
        WindowsUIButtonImageOptions3.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions3.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        WindowsUIButtonImageOptions3.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.WindowsUIButtonPanelSpreadsheet.Buttons.AddRange(New DevExpress.XtraEditors.ButtonPanel.IBaseButton() {New DevExpress.XtraBars.Docking2010.WindowsUISeparator(Nothing, True, -1, True), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions1, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Collapse all rows", -1, True, Nothing, True, False, True, "CollapseAll", -1, False), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions2, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Expand all rows", -1, True, Nothing, True, False, True, "ExpandAll", -1, False), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(), New DevExpress.XtraBars.Docking2010.WindowsUIButton("", True, WindowsUIButtonImageOptions3, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Export the data to PDF", -1, True, Nothing, True, False, True, "Save", -1, True)})
        Me.TablePanelSpreadsheetView.SetColumn(Me.WindowsUIButtonPanelSpreadsheet, 0)
        Me.WindowsUIButtonPanelSpreadsheet.ContentAlignment = System.Drawing.ContentAlignment.MiddleLeft
        Me.WindowsUIButtonPanelSpreadsheet.Location = New System.Drawing.Point(19, 972)
        Me.WindowsUIButtonPanelSpreadsheet.Name = "WindowsUIButtonPanelSpreadsheet"
        Me.TablePanelSpreadsheetView.SetRow(Me.WindowsUIButtonPanelSpreadsheet, 2)
        Me.WindowsUIButtonPanelSpreadsheet.Size = New System.Drawing.Size(1250, 68)
        Me.WindowsUIButtonPanelSpreadsheet.TabIndex = 2
        Me.WindowsUIButtonPanelSpreadsheet.Text = "WindowsUIButtonPanelAnalyser"
        '
        'SpreadsheetViewer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 28.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1288, 1071)
        Me.Controls.Add(Me.TablePanelSpreadsheetView)
        Me.Name = "SpreadsheetViewer"
        Me.Text = "SpreadsheetViewer"
        CType(Me.TablePanelSpreadsheetView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TablePanelSpreadsheetView.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TablePanelSpreadsheetView As DevExpress.Utils.Layout.TablePanel
    Friend WithEvents WindowsUIButtonPanelSpreadsheet As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel
    Friend WithEvents SpreadsheetFormulaBar1 As DevExpress.XtraSpreadsheet.SpreadsheetFormulaBar

End Class
