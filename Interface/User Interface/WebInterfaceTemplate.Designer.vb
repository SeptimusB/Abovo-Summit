Imports DevExpress.Data
Imports Abovo.AbovoAppCls
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class WebInterfaceTemplate
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
        Me.WebView2Main = New Microsoft.Web.WebView2.WinForms.WebView2()
        CType(Me.WebView2Main, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'WebView2Main
        '
        Me.WebView2Main.AllowExternalDrop = True
        Me.WebView2Main.CreationProperties = Nothing
        Me.WebView2Main.DefaultBackgroundColor = System.Drawing.Color.White
        Me.WebView2Main.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WebView2Main.Location = New System.Drawing.Point(0, 0)
        Me.WebView2Main.Name = "WebView2Main"
        Me.WebView2Main.Size = New System.Drawing.Size(1569, 882)
        Me.WebView2Main.TabIndex = 0
        Me.WebView2Main.ZoomFactor = 1.0R
        '
        'WebInterfaceTemplate
        '
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.WebView2Main)
        Me.Font = New System.Drawing.Font("Segoe UI Variable Display", 12.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Me.Margin = New System.Windows.Forms.Padding(8, 11, 8, 11)
        Me.Name = "WebInterfaceTemplate"
        Me.Size = New System.Drawing.Size(1569, 882)
        CType(Me.WebView2Main, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents WebView2Main As Microsoft.Web.WebView2.WinForms.WebView2
End Class
