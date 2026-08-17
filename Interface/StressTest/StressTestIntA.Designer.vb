<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class StressTestIntA
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
        Me.ComboBoxBreachMode = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.SimpleButtonModeSwitch = New DevExpress.XtraEditors.SimpleButton()
        Me.LayoutControl1 = New DevExpress.XtraLayout.LayoutControl()
        Me.Root = New DevExpress.XtraLayout.LayoutControlGroup()
        Me.LayoutControlItem1 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.EmptySpaceItem1 = New DevExpress.XtraLayout.EmptySpaceItem()
        Me.LayoutControlItem2 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.SimpleLabelItem1 = New DevExpress.XtraLayout.SimpleLabelItem()
        CType(Me.ComboBoxBreachMode.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LayoutControl1.SuspendLayout()
        CType(Me.Root, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SimpleLabelItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ComboBoxBreachMode
        '
        Me.ComboBoxBreachMode.Location = New System.Drawing.Point(1235, 57)
        Me.ComboBoxBreachMode.Name = "ComboBoxBreachMode"
        Me.ComboBoxBreachMode.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.ComboBoxBreachMode.Properties.DropDownRows = 10
        Me.ComboBoxBreachMode.Properties.Items.AddRange(New Object() {"Test 1", "Test 2", "Test 3", "Test 4", "Test 5", "Test 6", "Test 7", "Test 8", "Test 9", "Test 10"})
        Me.ComboBoxBreachMode.Size = New System.Drawing.Size(766, 46)
        Me.ComboBoxBreachMode.StyleController = Me.LayoutControl1
        Me.ComboBoxBreachMode.TabIndex = 6
        '
        'SimpleButtonModeSwitch
        '
        Me.SimpleButtonModeSwitch.Location = New System.Drawing.Point(516, 20)
        Me.SimpleButtonModeSwitch.Name = "SimpleButtonModeSwitch"
        Me.SimpleButtonModeSwitch.Size = New System.Drawing.Size(491, 40)
        Me.SimpleButtonModeSwitch.StyleController = Me.LayoutControl1
        Me.SimpleButtonModeSwitch.TabIndex = 7
        Me.SimpleButtonModeSwitch.Text = "Change to Multivariable mode"
        '
        'LayoutControl1
        '
        Me.LayoutControl1.Controls.Add(Me.ComboBoxBreachMode)
        Me.LayoutControl1.Controls.Add(Me.SimpleButtonModeSwitch)
        Me.LayoutControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutControl1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControl1.Name = "LayoutControl1"
        Me.LayoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = New System.Drawing.Rectangle(2244, 592, 1137, 700)
        Me.LayoutControl1.Root = Me.Root
        Me.LayoutControl1.Size = New System.Drawing.Size(2021, 248)
        Me.LayoutControl1.TabIndex = 8
        Me.LayoutControl1.Text = "LayoutControl1"
        '
        'Root
        '
        Me.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.[True]
        Me.Root.GroupBordersVisible = False
        Me.Root.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.EmptySpaceItem1, Me.LayoutControlItem2, Me.SimpleLabelItem1, Me.LayoutControlItem1})
        Me.Root.Name = "Root"
        Me.Root.Size = New System.Drawing.Size(2021, 248)
        Me.Root.TextVisible = False
        '
        'LayoutControlItem1
        '
        Me.LayoutControlItem1.Control = Me.SimpleButtonModeSwitch
        Me.LayoutControlItem1.Location = New System.Drawing.Point(496, 0)
        Me.LayoutControlItem1.Name = "LayoutControlItem1"
        Me.LayoutControlItem1.Size = New System.Drawing.Size(497, 214)
        Me.LayoutControlItem1.TextVisible = False
        '
        'EmptySpaceItem1
        '
        Me.EmptySpaceItem1.Location = New System.Drawing.Point(0, 0)
        Me.EmptySpaceItem1.Name = "EmptySpaceItem1"
        Me.EmptySpaceItem1.Size = New System.Drawing.Size(496, 214)
        '
        'LayoutControlItem2
        '
        Me.LayoutControlItem2.Control = Me.ComboBoxBreachMode
        Me.LayoutControlItem2.CustomizationFormText = "LayoutControlItemCMBBoxGroup"
        Me.LayoutControlItem2.Location = New System.Drawing.Point(993, 37)
        Me.LayoutControlItem2.Name = "LayoutControlItem2"
        Me.LayoutControlItem2.Size = New System.Drawing.Size(994, 177)
        Me.LayoutControlItem2.TextSize = New System.Drawing.Size(201, 31)
        '
        'SimpleLabelItem1
        '
        Me.SimpleLabelItem1.Location = New System.Drawing.Point(993, 0)
        Me.SimpleLabelItem1.Name = "SimpleLabelItem1"
        Me.SimpleLabelItem1.Size = New System.Drawing.Size(994, 37)
        Me.SimpleLabelItem1.Text = "Record as test:"
        Me.SimpleLabelItem1.TextSize = New System.Drawing.Size(201, 31)
        '
        'StressTestIntA
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.LayoutControl1)
        Me.Name = "StressTestIntA"
        Me.Size = New System.Drawing.Size(2021, 248)
        CType(Me.ComboBoxBreachMode.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LayoutControl1.ResumeLayout(False)
        CType(Me.Root, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SimpleLabelItem1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ComboBoxBreachMode As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents SimpleButtonModeSwitch As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LayoutControl1 As DevExpress.XtraLayout.LayoutControl
    Friend WithEvents Root As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents LayoutControlItem1 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents EmptySpaceItem1 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents LayoutControlItem2 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents SimpleLabelItem1 As DevExpress.XtraLayout.SimpleLabelItem
End Class
