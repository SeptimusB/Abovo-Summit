<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class VideoPlayer
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(VideoPlayer))
        Me.wmp = New AxWMPLib.AxWindowsMediaPlayer()
        CType(Me.wmp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'wmp
        '
        Me.wmp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wmp.Enabled = True
        Me.wmp.Location = New System.Drawing.Point(0, 0)
        Me.wmp.Name = "wmp"
        Me.wmp.OcxState = CType(resources.GetObject("wmp.OcxState"), System.Windows.Forms.AxHost.State)
        Me.wmp.Size = New System.Drawing.Size(1204, 875)
        Me.wmp.TabIndex = 0
        '
        'OpenNav
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.wmp)
        Me.Name = "OpenNav"
        Me.Size = New System.Drawing.Size(1204, 875)
        CType(Me.wmp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents wmp As AxWMPLib.AxWindowsMediaPlayer
End Class
