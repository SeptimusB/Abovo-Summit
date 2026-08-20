Imports System.Drawing
Imports System.Windows.Forms
Imports DevExpress.XtraEditors

''' <summary>
''' Compact, workbook-backed presentation of the FFR Inputs Adjustment
''' Statement.  It keeps the input sections together and avoids rendering the
''' unused spreadsheet geometry between the stock and loan-adjustment blocks.
''' </summary>
Public Class FFRInputsAdjStmtView
    Inherits XtraUserControl

    Private Const SheetName As String = "FFR Inputs Adj Stmt"
    Private Const WorkspaceMaximumWidth As Integer = 1640
    Private Const WorkspaceMinimumWidth As Integer = 1040

    Private ReadOnly ScrollHost As New XtraScrollableControl()
    Private ReadOnly Workspace As New TableLayoutPanel()
    Private ReadOnly ActualStockView As FFRWorkbookSheetView
    Private ReadOnly LoanAdjustmentsView As FFRWorkbookSheetView
    Private DisposedView As Boolean

    Public Event WorkbookCellChanged As EventHandler

    Public Sub New(SetModelID As Integer)
        Dock = DockStyle.Fill

        ActualStockView = New FFRWorkbookSheetView(
            SetModelID,
            SheetName,
            "A5:J50",
            "FFR Actual Stock Inputs",
            True,
            True,
            True)
        LoanAdjustmentsView = New FFRWorkbookSheetView(
            SetModelID,
            SheetName,
            "A53:AF70",
            "Statement of Cash Flow – Movements in Loans",
            True,
            True,
            True)

        BuildNativeSurface()
        AddHandler ActualStockView.WorkbookCellChanged, AddressOf ChildWorkbookCellChanged
        AddHandler LoanAdjustmentsView.WorkbookCellChanged, AddressOf ChildWorkbookCellChanged
    End Sub

    Public ReadOnly Property WorksheetName As String
        Get
            Return SheetName
        End Get
    End Property

    Private Sub BuildNativeSurface()
        BackColor = Color.White

        ScrollHost.Dock = DockStyle.Fill
        ScrollHost.BackColor = Color.White
        ScrollHost.AutoScroll = True

        Workspace.BackColor = Color.White
        Workspace.ColumnCount = 1
        Workspace.RowCount = 3
        Workspace.Padding = New Padding(12, 10, 12, 16)
        Workspace.Margin = New Padding(0)
        Workspace.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        Workspace.RowStyles.Add(New RowStyle(SizeType.Absolute, 34.0F))
        Workspace.RowStyles.Add(New RowStyle(SizeType.Absolute, 760.0F))
        Workspace.RowStyles.Add(New RowStyle(SizeType.Absolute, 340.0F))

        Dim PageTitle As New LabelControl With {
            .Dock = DockStyle.Fill,
            .Text = "FFR Inputs Adjustment Statement",
            .AutoSizeMode = LabelAutoSizeMode.None
        }
        PageTitle.Appearance.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
        PageTitle.Appearance.ForeColor = Color.FromArgb(32, 58, 89)
        PageTitle.Appearance.Options.UseFont = True
        PageTitle.Appearance.Options.UseForeColor = True

        ActualStockView.Dock = DockStyle.Fill
        ActualStockView.Margin = New Padding(0, 0, 0, 8)
        LoanAdjustmentsView.Dock = DockStyle.Fill
        LoanAdjustmentsView.Margin = New Padding(0, 8, 0, 0)

        Workspace.Controls.Add(PageTitle, 0, 0)
        Workspace.Controls.Add(ActualStockView, 0, 1)
        Workspace.Controls.Add(LoanAdjustmentsView, 0, 2)
        ScrollHost.Controls.Add(Workspace)
        AddHandler ScrollHost.Resize, AddressOf PositionWorkspace
        Controls.Add(ScrollHost)
        PositionWorkspace(Nothing, EventArgs.Empty)
    End Sub

    Public Sub RefreshFromWorkbook()
        If DisposedView Then Return
        ActualStockView.RefreshFromWorkbook()
        LoanAdjustmentsView.RefreshFromWorkbook()
    End Sub

    Private Sub ChildWorkbookCellChanged(sender As Object, e As EventArgs)
        RefreshFromWorkbook()
        RaiseEvent WorkbookCellChanged(Me, EventArgs.Empty)
    End Sub

    Private Sub PositionWorkspace(sender As Object, e As EventArgs)
        Dim AvailableWidth As Integer = Math.Max(0, ScrollHost.ClientSize.Width - 32)
        Dim DesiredWidth As Integer = Math.Min(WorkspaceMaximumWidth, Math.Max(WorkspaceMinimumWidth, AvailableWidth))
        Workspace.Size = New Size(DesiredWidth, 1160)
        Workspace.Location = New Point(Math.Max(16, (ScrollHost.ClientSize.Width - DesiredWidth) \ 2), 6)
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso Not DisposedView Then
            DisposedView = True
            RemoveHandler ScrollHost.Resize, AddressOf PositionWorkspace
            RemoveHandler ActualStockView.WorkbookCellChanged, AddressOf ChildWorkbookCellChanged
            RemoveHandler LoanAdjustmentsView.WorkbookCellChanged, AddressOf ChildWorkbookCellChanged
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class
