<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class InsantMsgForm
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
        Me.TablePanel1 = New DevExpress.Utils.Layout.TablePanel()
        Me.MemoEditMsg = New DevExpress.XtraEditors.MemoEdit()
        Me.WindowsUIButtonPanel1 = New DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel()
        CType(Me.TablePanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TablePanel1.SuspendLayout()
        CType(Me.MemoEditMsg.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TablePanel1
        '
        Me.TablePanel1.Columns.AddRange(New DevExpress.Utils.Layout.TablePanelColumn() {New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 55.0!)})
        Me.TablePanel1.Controls.Add(Me.WindowsUIButtonPanel1)
        Me.TablePanel1.Controls.Add(Me.MemoEditMsg)
        Me.TablePanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TablePanel1.Location = New System.Drawing.Point(0, 0)
        Me.TablePanel1.Name = "TablePanel1"
        Me.TablePanel1.Rows.AddRange(New DevExpress.Utils.Layout.TablePanelRow() {New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 26.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 71.0!)})
        Me.TablePanel1.Size = New System.Drawing.Size(1113, 641)
        Me.TablePanel1.TabIndex = 0
        Me.TablePanel1.UseSkinIndents = True
        '
        'MemoEditMsg
        '
        Me.TablePanel1.SetColumn(Me.MemoEditMsg, 0)
        Me.MemoEditMsg.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MemoEditMsg.Location = New System.Drawing.Point(22, 20)
        Me.MemoEditMsg.Name = "MemoEditMsg"
        Me.TablePanel1.SetRow(Me.MemoEditMsg, 0)
        Me.MemoEditMsg.Size = New System.Drawing.Size(1069, 528)
        Me.MemoEditMsg.TabIndex = 0
        '
        'WindowsUIButtonPanel1
        '
        Me.WindowsUIButtonPanel1.Buttons.AddRange(New DevExpress.XtraEditors.ButtonPanel.IBaseButton() {New DevExpress.XtraBars.Docking2010.WindowsUIButton(), New DevExpress.XtraBars.Docking2010.WindowsUIButton()})
        Me.TablePanel1.SetColumn(Me.WindowsUIButtonPanel1, 0)
        Me.WindowsUIButtonPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WindowsUIButtonPanel1.Location = New System.Drawing.Point(22, 554)
        Me.WindowsUIButtonPanel1.Name = "WindowsUIButtonPanel1"
        Me.TablePanel1.SetRow(Me.WindowsUIButtonPanel1, 1)
        Me.WindowsUIButtonPanel1.Size = New System.Drawing.Size(1069, 65)
        Me.WindowsUIButtonPanel1.TabIndex = 4
        Me.WindowsUIButtonPanel1.Text = "WindowsUIButtonPanel1"
        '
        'InsantMsgForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(15.0!, 38.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1113, 641)
        Me.Controls.Add(Me.TablePanel1)
        Me.Name = "InsantMsgForm"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Validation Error"
        CType(Me.TablePanel1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TablePanel1.ResumeLayout(False)
        CType(Me.MemoEditMsg.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TablePanel1 As DevExpress.Utils.Layout.TablePanel
    Friend WithEvents MemoEditMsg As DevExpress.XtraEditors.MemoEdit
    Friend WithEvents WindowsUIButtonPanel1 As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel
End Class
