Imports System.ComponentModel
Imports Abovo
Imports Abovo.ObjectFormatter
Imports DevExpress.Data
Imports DevExpress.XtraExport.Helpers
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraSpreadsheet


Public Class HistoryManager
    Private Formatter As ObjectFormatter
    Private ModelID As Integer

    Public Sub New()

        InitializeComponent()

        Formatter = New ObjectFormatter



        Dim GV As GridView = GridControlHistory.MainView

        Formatter.FormatGridControl(GridControlHistory)
        Formatter.FormatGridView(GV, GridControlHistory)

        GridControlHistory.DataSource = MasterChangeLog.ChangeLog

        GV.BeginSort()
        GV.SortInfo.ClearAndAddRange({
              New GridColumnSortInfo(GV.Columns(0), ColumnSortOrder.Descending)
            })
        GV.EndSort()

        GV.OptionsBehavior.Editable = False

        GV.Columns(0).Visible = False

        GV.Columns(1).ColumnEdit = Nothing
        GV.Columns(1).OptionsColumn.ReadOnly = True
        GV.Columns(1).Caption = "Time"
        GV.Columns(1).DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        GV.Columns(1).DisplayFormat.FormatString = "H:mm:ss"
        GV.Columns(1).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far

        GV.Columns(2).Caption = "File #"
        GV.Columns(2).OptionsColumn.ReadOnly = True
        GV.Columns(2).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far

        GV.Columns(3).Caption = "Desc."
        GV.Columns(3).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        GV.Columns(3).OptionsColumn.ReadOnly = True

        GV.Columns(4).Caption = "Area/<br>worksheet"
        GV.Columns(4).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        GV.Columns(4).OptionsColumn.ReadOnly = True

        GV.Columns(5).Caption = "Equiv.<br>cell"
        GV.Columns(5).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        GV.Columns(5).OptionsColumn.ReadOnly = True

        GV.Columns(6).Caption = "Original<br>Value"
        GV.Columns(6).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        GV.Columns(6).OptionsColumn.ReadOnly = True

        GV.Columns(7).Caption = "New<br>Value"
        GV.Columns(7).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        GV.Columns(7).OptionsColumn.ReadOnly = True

        GV.Columns(8).Visible = False
        GV.Columns(10).Visible = False

        GV.BestFitColumns()

        Me.WindowsUIButtonPanelHistoryActions.ForeColor = GeneralFunctions.AbovoBlue

        AddHandler GV.ShowingEditor, AddressOf GV_ShowningEditor


    End Sub


    Sub GV_ShowningEditor(sender As Object, ByVal e As CancelEventArgs)

        Dim view As GridView = DirectCast(sender, GridView)

        If view.FocusedColumn.FieldName = "TimeStamp" Then
            e.Cancel = True
        End If

        Dim EventID As Integer = view.GetFocusedRowCellValue("EventID")

        Dim ChangeLogEntry As DataRow = MasterChangeLog.ChangeLog.Rows(EventID)

        'If Not IsNothing(ChangeLogEntry) Then
        '    MemoEditChangeDetails.Text = ChangeLogEntry.ChangeDetails
        'End If

    End Sub

    Sub ManageClose(sender As Object, ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing

        e.Cancel = True
        Me.Hide()

    End Sub

End Class