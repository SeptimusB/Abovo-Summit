Imports Abovo.CustomGrid
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class CustomGridWrapper
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()

        Me.WrappedCGC = New CustomGridControl()
        Me.WrappedGridView = New CustomGridView()
        CType(Me.WrappedCGC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WrappedGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'WrappedCGC
        '
        Me.WrappedCGC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WrappedCGC.Location = New System.Drawing.Point(0, 0)
        Me.WrappedCGC.MainView = Me.WrappedGridView
        Me.WrappedCGC.Name = "WrappedCGC"
        Me.WrappedCGC.Size = New System.Drawing.Size(1502, 931)
        Me.WrappedCGC.TabIndex = 0
        Me.WrappedCGC.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.WrappedGridView})
        '
        'GridView1
        '
        Me.WrappedGridView.GridControl = Me.WrappedCGC
        Me.WrappedGridView.Name = "WrappedGridView"
        '
        'CustomGridWrapper
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.WrappedCGC)
        Me.Name = "CustomGridWrapper"
        Me.Size = New System.Drawing.Size(1502, 931)
        CType(Me.WrappedCGC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WrappedGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub


End Class
