<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TempModelViewer
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.SpreadsheetControl1 = New DevExpress.XtraSpreadsheet.SpreadsheetControl()
        Me.SuspendLayout()
        '
        'SpreadsheetControl1
        '
        Me.SpreadsheetControl1.Location = New System.Drawing.Point(-33, -67)
        Me.SpreadsheetControl1.Name = "SpreadsheetControl1"
        Me.SpreadsheetControl1.Size = New System.Drawing.Size(400, 200)
        Me.SpreadsheetControl1.TabIndex = 0
        Me.SpreadsheetControl1.Text = "SpreadsheetControl1"
        '
        'TempModelViewer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1273, 951)
        Me.Controls.Add(Me.SpreadsheetControl1)
        Me.Name = "TempModelViewer"
        Me.Text = "TempModelViewer"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SpreadsheetControl1 As DevExpress.XtraSpreadsheet.SpreadsheetControl
End Class
