Imports DevExpress.Data
Imports Abovo.AbovoAppCls
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BPDashboard
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
        Me.TablePanelDashboard = New DevExpress.Utils.Layout.TablePanel()
        CType(Me.TablePanelDashboard, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TablePanelDashboard
        '
        Me.TablePanelDashboard.Columns.AddRange(New DevExpress.Utils.Layout.TablePanelColumn() {New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 1.0!), New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 1.0!), New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 1.0!)})
        Me.TablePanelDashboard.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TablePanelDashboard.Location = New System.Drawing.Point(0, 0)
        Me.TablePanelDashboard.Name = "TablePanelDashboard"
        Me.TablePanelDashboard.Rows.AddRange(New DevExpress.Utils.Layout.TablePanelRow() {New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 1.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 1.0!)})
        Me.TablePanelDashboard.Size = New System.Drawing.Size(1569, 882)
        Me.TablePanelDashboard.TabIndex = 0
        Me.TablePanelDashboard.UseSkinIndents = True
        '
        'BPDashboard
        '
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.TablePanelDashboard)
        Me.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.Margin = New System.Windows.Forms.Padding(8, 11, 8, 11)
        Me.Name = "BPDashboard"
        Me.Size = New System.Drawing.Size(1569, 882)
        CType(Me.TablePanelDashboard, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TablePanelDashboard As DevExpress.Utils.Layout.TablePanel
End Class
