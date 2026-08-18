<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MessageForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MessageForm))
        Me.SimpleButtonYes = New DevExpress.XtraEditors.SimpleButton()
        Me.SimpleButtonChoiceNo = New DevExpress.XtraEditors.SimpleButton()
        Me.WebBrowserMessage = New System.Windows.Forms.WebBrowser()
        Me.SimpleButtonCopy = New DevExpress.XtraEditors.SimpleButton()
        Me.SimpleButtonOK = New DevExpress.XtraEditors.SimpleButton()
        Me.SuspendLayout()
        '
        'SimpleButtonYes
        '
        Me.SimpleButtonYes.Location = New System.Drawing.Point(92, 306)
        Me.SimpleButtonYes.Margin = New System.Windows.Forms.Padding(4)
        Me.SimpleButtonYes.Name = "SimpleButtonYes"
        Me.SimpleButtonYes.Size = New System.Drawing.Size(130, 40)
        Me.SimpleButtonYes.TabIndex = 1
        Me.SimpleButtonYes.Text = "Yes"
        '
        'SimpleButtonChoiceNo
        '
        Me.SimpleButtonChoiceNo.Location = New System.Drawing.Point(388, 306)
        Me.SimpleButtonChoiceNo.Margin = New System.Windows.Forms.Padding(4)
        Me.SimpleButtonChoiceNo.Name = "SimpleButtonChoiceNo"
        Me.SimpleButtonChoiceNo.Size = New System.Drawing.Size(130, 40)
        Me.SimpleButtonChoiceNo.TabIndex = 2
        Me.SimpleButtonChoiceNo.Text = "No"
        Me.SimpleButtonChoiceNo.ToolTip = "Select No"
        '
        'WebBrowserMessage
        '
        Me.WebBrowserMessage.Location = New System.Drawing.Point(28, 13)
        Me.WebBrowserMessage.Margin = New System.Windows.Forms.Padding(4)
        Me.WebBrowserMessage.MinimumSize = New System.Drawing.Size(20, 19)
        Me.WebBrowserMessage.Name = "WebBrowserMessage"
        Me.WebBrowserMessage.ScriptErrorsSuppressed = True
        Me.WebBrowserMessage.Size = New System.Drawing.Size(594, 239)
        Me.WebBrowserMessage.TabIndex = 3
        '
        'SimpleButtonCopy
        '
        Me.SimpleButtonCopy.ImageOptions.Image = CType(resources.GetObject("SimpleButtonCopy.ImageOptions.Image"), System.Drawing.Image)
        Me.SimpleButtonCopy.ImageOptions.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.TopCenter
        Me.SimpleButtonCopy.Location = New System.Drawing.Point(601, 259)
        Me.SimpleButtonCopy.Name = "SimpleButtonCopy"
        Me.SimpleButtonCopy.Size = New System.Drawing.Size(40, 40)
        Me.SimpleButtonCopy.TabIndex = 4
        Me.SimpleButtonCopy.ToolTip = "Copy message text"
        '
        'SimpleButtonOK
        '
        Me.SimpleButtonOK.Location = New System.Drawing.Point(240, 306)
        Me.SimpleButtonOK.Margin = New System.Windows.Forms.Padding(4)
        Me.SimpleButtonOK.Name = "SimpleButtonOK"
        Me.SimpleButtonOK.Size = New System.Drawing.Size(130, 40)
        Me.SimpleButtonOK.TabIndex = 5
        Me.SimpleButtonOK.Text = "Okay"
        Me.SimpleButtonOK.ToolTip = "Select Okay"
        '
        'MessageForm
        '
        Me.Appearance.BackColor = System.Drawing.Color.White
        Me.Appearance.BorderColor = System.Drawing.Color.White
        Me.Appearance.Options.UseBackColor = True
        Me.Appearance.Options.UseBorderColor = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(643, 378)
        Me.ControlBox = False
        Me.Controls.Add(Me.SimpleButtonOK)
        Me.Controls.Add(Me.SimpleButtonCopy)
        Me.Controls.Add(Me.WebBrowserMessage)
        Me.Controls.Add(Me.SimpleButtonChoiceNo)
        Me.Controls.Add(Me.SimpleButtonYes)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.IconOptions.Icon = CType(resources.GetObject("MessageForm.IconOptions.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "MessageForm"
        Me.ShowInTaskbar = False
        Me.Text = "abovo-summit"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SimpleButtonYes As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents SimpleButtonChoiceNo As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents WebBrowserMessage As WebBrowser
    Friend WithEvents SimpleButtonCopy As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents SimpleButtonOK As DevExpress.XtraEditors.SimpleButton
End Class
