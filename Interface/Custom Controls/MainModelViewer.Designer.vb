<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainModelViewer
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
        Debug.Print("1")
        Me.TablePanelSSV = New DevExpress.Utils.Layout.TablePanel()
        Debug.Print("2")
        Me.SpreadsheetFormulaBarMMV = New DevExpress.XtraSpreadsheet.SpreadsheetFormulaBar()
        Debug.Print("3")
        CType(Me.TablePanelSSV, System.ComponentModel.ISupportInitialize).BeginInit()
        Debug.Print("4")
        Me.TablePanelSSV.SuspendLayout()
        Debug.Print("5")
        Me.SuspendLayout()

        '
        'TablePanelSSV
        '
        Me.TablePanelSSV.Columns.AddRange(New DevExpress.Utils.Layout.TablePanelColumn() {New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 5.0!)})
        Me.TablePanelSSV.Controls.Add(Me.SpreadsheetFormulaBarMMV)
        Me.TablePanelSSV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TablePanelSSV.Location = New System.Drawing.Point(0, 0)
        Me.TablePanelSSV.Name = "TablePanelSSV"
        Me.TablePanelSSV.Rows.AddRange(New DevExpress.Utils.Layout.TablePanelRow() {New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 74.66669!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 26.0!)})
        Me.TablePanelSSV.Size = New System.Drawing.Size(1521, 1053)
        Me.TablePanelSSV.TabIndex = 0
        Me.TablePanelSSV.UseSkinIndents = True

        '
        'SpreadsheetFormulaBarMMV
        '
        Me.SpreadsheetFormulaBarMMV.AccessibleName = "Formula bar"
        Me.TablePanelSSV.SetColumn(Me.SpreadsheetFormulaBarMMV, 0)
        Me.SpreadsheetFormulaBarMMV.Expanded = True
        Me.SpreadsheetFormulaBarMMV.Location = New System.Drawing.Point(22, 26)
        Me.SpreadsheetFormulaBarMMV.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.SpreadsheetFormulaBarMMV.MinimumSize = New System.Drawing.Size(0, 34)
        Me.SpreadsheetFormulaBarMMV.Name = "SpreadsheetFormulaBarMMV"
        Me.TablePanelSSV.SetRow(Me.SpreadsheetFormulaBarMMV, 0)
        Me.SpreadsheetFormulaBarMMV.Size = New System.Drawing.Size(1199, 52)
        Me.SpreadsheetFormulaBarMMV.TabIndex = 0

        '
        'MainModelViewer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 28.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1521, 1053)
        Me.Controls.Add(Me.TablePanelSSV)
        Me.Name = "MainModelViewer"
        Me.Text = "MainModelViewer"
        CType(Me.TablePanelSSV, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TablePanelSSV.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TablePanelSSV As DevExpress.Utils.Layout.TablePanel
    Friend WithEvents SpreadsheetFormulaBarMMV As DevExpress.XtraSpreadsheet.SpreadsheetFormulaBar
End Class
