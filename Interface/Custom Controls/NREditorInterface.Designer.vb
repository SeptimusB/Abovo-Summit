<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NREditorInterface
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim WindowsUIButtonImageOptions5 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(NREditorInterface))
        Dim WindowsUIButtonImageOptions6 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions7 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions8 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Me.TablePanel1 = New DevExpress.Utils.Layout.TablePanel()
        Me.LabelControlMsg = New DevExpress.XtraEditors.LabelControl()
        Me.WindowsUIButtonPanelActions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel()

        Me.GridView1 = New DevExpress.XtraGrid.Views.Grid.GridView()
        CType(Me.TablePanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TablePanel1.SuspendLayout()
        CType(Me.GridControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TablePanel1
        '
        Me.TablePanel1.Columns.AddRange(New DevExpress.Utils.Layout.TablePanelColumn() {New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 55.0!)})
        Me.TablePanel1.Controls.Add(Me.LabelControlMsg)
        Me.TablePanel1.Controls.Add(Me.WindowsUIButtonPanelActions)

        Me.TablePanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TablePanel1.Location = New System.Drawing.Point(0, 0)
        Me.TablePanel1.Margin = New System.Windows.Forms.Padding(2)
        Me.TablePanel1.Name = "TablePanel1"
        Me.TablePanel1.Rows.AddRange(New DevExpress.Utils.Layout.TablePanelRow() {New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 54.00003!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 115.3334!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 26.0!)})
        Me.TablePanel1.Size = New System.Drawing.Size(1022, 753)
        Me.TablePanel1.TabIndex = 0
        Me.TablePanel1.UseSkinIndents = True
        '
        'LabelControlMsg
        '
        Me.LabelControlMsg.Appearance.Font = New System.Drawing.Font("Segoe UI", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControlMsg.Appearance.Options.UseFont = True
        Me.TablePanel1.SetColumn(Me.LabelControlMsg, 0)
        Me.LabelControlMsg.Location = New System.Drawing.Point(19, 19)
        Me.LabelControlMsg.Name = "LabelControlMsg"
        Me.TablePanel1.SetRow(Me.LabelControlMsg, 0)
        Me.LabelControlMsg.Size = New System.Drawing.Size(198, 45)
        Me.LabelControlMsg.TabIndex = 2
        Me.LabelControlMsg.Text = "LabelControl1"
        '
        'WindowsUIButtonPanelActions
        '
        WindowsUIButtonImageOptions5.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions5.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        WindowsUIButtonImageOptions6.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions6.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        WindowsUIButtonImageOptions7.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions7.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        WindowsUIButtonImageOptions8.SvgImage = CType(resources.GetObject("WindowsUIButtonImageOptions8.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.WindowsUIButtonPanelActions.Buttons.AddRange(New DevExpress.XtraEditors.ButtonPanel.IBaseButton() {New DevExpress.XtraBars.Docking2010.WindowsUIButton("Cancel", True, WindowsUIButtonImageOptions5, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Cancel and Return", -1, True, Nothing, True, False, True, "Cancel", -1, False), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Add Records", True, WindowsUIButtonImageOptions6, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Add Rows to this definition", -1, True, Nothing, True, False, True, "AddRows", -1, False), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Remove Rows", True, WindowsUIButtonImageOptions7, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Delete Rows from this definition", -1, True, Nothing, True, False, True, "DelRows", -1, False), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Apply", True, WindowsUIButtonImageOptions8, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Apply Changes", -1, True, Nothing, True, False, True, "Apply", -1, False)})
        Me.TablePanel1.SetColumn(Me.WindowsUIButtonPanelActions, 0)
        Me.WindowsUIButtonPanelActions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WindowsUIButtonPanelActions.Location = New System.Drawing.Point(18, 71)
        Me.WindowsUIButtonPanelActions.Margin = New System.Windows.Forms.Padding(2)
        Me.WindowsUIButtonPanelActions.Name = "WindowsUIButtonPanelActions"
        Me.TablePanel1.SetRow(Me.WindowsUIButtonPanelActions, 1)
        Me.WindowsUIButtonPanelActions.Size = New System.Drawing.Size(986, 111)
        Me.WindowsUIButtonPanelActions.TabIndex = 1
        Me.WindowsUIButtonPanelActions.Text = "WindowsUIButtonPanel1"
        '
        'GridControl1
        '

        '
        'GridView1
        '
        Me.GridView1.DetailHeight = 258
        Me.GridView1.GridControl = Me.GridControl1
        Me.GridView1.Name = "GridView1"
        Me.GridView1.OptionsEditForm.PopupEditFormWidth = 587
        '
        'NREditorInterface
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 28.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1022, 753)
        Me.Controls.Add(Me.TablePanel1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "NREditorInterface"
        CType(Me.TablePanel1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TablePanel1.ResumeLayout(False)
        Me.TablePanel1.PerformLayout()
        CType(Me.GridControl1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TablePanel1 As DevExpress.Utils.Layout.TablePanel
    Friend WithEvents GridView1 As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents WindowsUIButtonPanelActions As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel
    Friend WithEvents LabelControlMsg As DevExpress.XtraEditors.LabelControl
End Class
