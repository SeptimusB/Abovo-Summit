<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class HistoryManager
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
        Dim WindowsUIButtonImageOptions1 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Dim WindowsUIButtonImageOptions2 As DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonImageOptions()
        Me.TablePanelHistory = New DevExpress.Utils.Layout.TablePanel()
        Me.WindowsUIButtonPanelHistoryActions = New DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel()
        Me.GridControlHistory = New DevExpress.XtraGrid.GridControl()
        Me.GridViewHistory = New DevExpress.XtraGrid.Views.Grid.GridView()
        CType(Me.TablePanelHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TablePanelHistory.SuspendLayout()
        CType(Me.GridControlHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridViewHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TablePanelHistory
        '
        Me.TablePanelHistory.Appearance.BackColor = System.Drawing.Color.White
        Me.TablePanelHistory.Appearance.Options.UseBackColor = True
        Me.TablePanelHistory.Columns.AddRange(New DevExpress.Utils.Layout.TablePanelColumn() {New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 55.0!), New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 50.0!)})
        Me.TablePanelHistory.Controls.Add(Me.WindowsUIButtonPanelHistoryActions)
        Me.TablePanelHistory.Controls.Add(Me.GridControlHistory)
        Me.TablePanelHistory.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TablePanelHistory.Location = New System.Drawing.Point(0, 0)
        Me.TablePanelHistory.Margin = New System.Windows.Forms.Padding(2)
        Me.TablePanelHistory.Name = "TablePanelHistory"
        Me.TablePanelHistory.Rows.AddRange(New DevExpress.Utils.Layout.TablePanelRow() {New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 85.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 26.0!), New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 26.0!)})
        Me.TablePanelHistory.Size = New System.Drawing.Size(1366, 954)
        Me.TablePanelHistory.TabIndex = 0
        Me.TablePanelHistory.UseSkinIndents = True
        '
        'WindowsUIButtonPanelHistoryActions
        '
        WindowsUIButtonImageOptions1.SvgImage = Global.My.Resources.Resources.actions_refresh
        WindowsUIButtonImageOptions2.SvgImage = Global.My.Resources.Resources.close
        Me.WindowsUIButtonPanelHistoryActions.Buttons.AddRange(New DevExpress.XtraEditors.ButtonPanel.IBaseButton() {New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions1, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Undo all visible actions", -1, True, Nothing, True, False, True, "UndoAll", -1, False), New DevExpress.XtraBars.Docking2010.WindowsUISeparator(), New DevExpress.XtraBars.Docking2010.WindowsUIButton("Button", False, WindowsUIButtonImageOptions2, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "Close", -1, True, Nothing, True, False, True, "Close", -1, False)})
        Me.TablePanelHistory.SetColumn(Me.WindowsUIButtonPanelHistoryActions, 0)
        Me.WindowsUIButtonPanelHistoryActions.ContentAlignment = System.Drawing.ContentAlignment.MiddleLeft
        Me.WindowsUIButtonPanelHistoryActions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WindowsUIButtonPanelHistoryActions.Location = New System.Drawing.Point(18, 17)
        Me.WindowsUIButtonPanelHistoryActions.Margin = New System.Windows.Forms.Padding(2)
        Me.WindowsUIButtonPanelHistoryActions.Name = "WindowsUIButtonPanelHistoryActions"
        Me.WindowsUIButtonPanelHistoryActions.Padding = New System.Windows.Forms.Padding(5)
        Me.TablePanelHistory.SetRow(Me.WindowsUIButtonPanelHistoryActions, 0)
        Me.WindowsUIButtonPanelHistoryActions.Size = New System.Drawing.Size(695, 81)
        Me.WindowsUIButtonPanelHistoryActions.TabIndex = 1
        '
        'GridControlHistory
        '
        Me.TablePanelHistory.SetColumn(Me.GridControlHistory, 0)
        Me.TablePanelHistory.SetColumnSpan(Me.GridControlHistory, 2)
        Me.GridControlHistory.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GridControlHistory.EmbeddedNavigator.Margin = New System.Windows.Forms.Padding(2)
        Me.GridControlHistory.Location = New System.Drawing.Point(18, 102)
        Me.GridControlHistory.MainView = Me.GridViewHistory
        Me.GridControlHistory.Margin = New System.Windows.Forms.Padding(2)
        Me.GridControlHistory.Name = "GridControlHistory"
        Me.TablePanelHistory.SetRow(Me.GridControlHistory, 1)
        Me.GridControlHistory.Size = New System.Drawing.Size(1330, 853)
        Me.GridControlHistory.TabIndex = 0
        Me.GridControlHistory.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridViewHistory})
        '
        'GridViewHistory
        '
        Me.GridViewHistory.DetailHeight = 218
        Me.GridViewHistory.GridControl = Me.GridControlHistory
        Me.GridViewHistory.Name = "GridViewHistory"
        Me.GridViewHistory.OptionsEditForm.PopupEditFormWidth = 489
        '
        'HistoryManager
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 28.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1366, 954)
        Me.Controls.Add(Me.TablePanelHistory)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "HistoryManager"
        Me.Text = "HistoryManager"
        CType(Me.TablePanelHistory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TablePanelHistory.ResumeLayout(False)
        CType(Me.GridControlHistory, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridViewHistory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TablePanelHistory As DevExpress.Utils.Layout.TablePanel
    Friend WithEvents WindowsUIButtonPanelHistoryActions As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel
    Friend WithEvents GridControlHistory As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridViewHistory As DevExpress.XtraGrid.Views.Grid.GridView
End Class
