Imports System.Runtime.Serialization
Imports Abovo
Imports Abovo.WorkbookManager
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.XtraSpreadsheet.Model
Imports DevExpress.XtraSpreadsheet.PrintLayoutEngine
Public Class NREditorInterface

    Private NRds As RangeDataSource
    Private Formatter As ObjectFormatter
    Private ModelID As Integer
    Private DoActionTok As ActionToken
    Private ActionRange As DevExpress.Spreadsheet.CellRange
    Private GridControl1 As DevExpress.XtraGrid.GridControl

    Private NROrient As String = "Rows"
    Public ReturnActionTok As ActionToken

    Sub New(Parent As Object, SetModelID As Integer, RunAction As ActionToken, SetTitle As String)

        DoActionTok = RunAction
        Me.Text = "Edit Rows"

        ModelID = SetModelID
        Formatter = New ObjectFormatter
        Me.GridControl1 = New DevExpress.XtraGrid.GridControl()

        ' This call is required by the designer.
        InitializeComponent()
        SetDataRange()
        If NROrient = "Cols" Then
            Me.Text = "Edit Columns"
            SwapOrientation()
        End If
        Me.LabelControlMsg.Text = SetTitle
        ProcessGrid()

    End Sub
    Private Sub ProcessGrid()

        Me.TablePanel1.Controls.Add(Me.GridControl1)
        Me.TablePanel1.SetColumn(Me.GridControl1, 0)
        Me.GridControl1.EmbeddedNavigator.Margin = New System.Windows.Forms.Padding(2)
        Me.GridControl1.Location = New System.Drawing.Point(18, 186)
        'Me.GridControl1.MainView = Me.GridView1
        Me.GridControl1.Margin = New System.Windows.Forms.Padding(2)
        Me.GridControl1.Name = "GridControl1"
        Me.TablePanel1.SetRow(Me.GridControl1, 2)
        Me.GridControl1.Size = New System.Drawing.Size(986, 549)
        Me.GridControl1.TabIndex = 0
        'Me.GridControl1.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridView1})

        GridControl1.Font = Formatter.DefaultLargeFont
        GridControl1.Dock = DockStyle.None
        GridControl1.Width = Me.Width / 2
        GridControl1.Left = (Me.Width - GridControl1.Width) / 2
        GridControl1.Top = 0
        Formatter.FormatGridView(GridControl1.Views(0), GridControl1, "Massive")
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(GridControl1.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
        view.BestFitColumns()

        Formatter.FormatGridControl(GridControl1)
        GridControl1.BackColor = Color.LightGray
    End Sub
    Private Sub SwapOrientation()

        Dim PanelButton As WindowsUIButton

        For Each ExButton In WindowsUIButtonPanelActions.Buttons

            If TryCast(ExButton, WindowsUIButton) IsNot Nothing Then
                PanelButton = CType(ExButton, WindowsUIButton)
            Else
                Continue For
            End If

            If PanelButton.Caption = "Add Rows" Then

                PanelButton.Caption = "Add Columns"
                PanelButton.Tag = "AddCols"

            End If

            If PanelButton.Caption = "Remove Rows" Then

                PanelButton.Caption = "Remove Columns"
                PanelButton.Tag = "DelCols"

            End If

        Next

    End Sub
    Sub HandleMenuComands(sender As Object, e As ButtonEventArgs) Handles WindowsUIButtonPanelActions.ButtonClick

        Dim ButSender As WindowsUIButton = TryCast(e.Button, DevExpress.XtraBars.Docking2010.WindowsUIButton)
        If ButSender Is Nothing Then
            Return
        End If
        Dim tag As String = ButSender.Tag.ToString()

        Select Case tag

            Case "Cancel"

                Me.DialogResult = DialogResult.Cancel
                Me.Close()

            Case "AddRows"

                Dim editor As New Num_Popup("Add Rows", ModelID, 3)

                editor.ShowDialog(Me)
                If editor.MyAction = DialogResult.OK Then
                    Me.Cursor = Cursors.WaitCursor
                    Dim NewRows As Integer = editor.MyValue
                    If NewRows > 0 Then
                        DisconnectRDS()
                        WorkbookManager.InsertRows(ModelID, DoActionTok.ActionStrData1, NewRows)
                        TransDBManager.CheckTransDBActions(ModelID, DoActionTok.ActionStrData1)
                        REconnect()
                    End If
                    Me.Cursor = Cursors.Default
                End If
                editor.Dispose()
                editor = Nothing

            Case "AddCols"

                Dim editor As New Num_Popup("Add Columns", ModelID, 3)

                editor.ShowDialog(Me)
                If editor.MyAction = DialogResult.OK Then
                    Me.Cursor = Cursors.WaitCursor
                    Dim NewCols As Integer = editor.MyValue
                    If NewCols > 0 Then
                        DisconnectRDS()
                        WorkbookManager.InsertColumns(ModelID, DoActionTok.ActionStrData1, NewCols)
                        TransDBManager.CheckTransDBActions(ModelID, DoActionTok.ActionStrData1)
                        REconnect()
                    End If
                    Me.Cursor = Cursors.Default
                End If
                editor.Dispose()
                editor = Nothing

            Case "DelRows"

                Dim editor As New Num_Popup("Delete Rows", ModelID, 3)

                editor.ShowDialog(Me)
                If editor.MyAction = DialogResult.OK Then
                    Me.Cursor = Cursors.WaitCursor
                    Dim NewRows As Integer = editor.MyValue
                    If NewRows > 0 Then
                        DisconnectRDS()
                        WorkbookManager.DeleteRows(ModelID, DoActionTok.ActionStrData1, NewRows)
                        TransDBManager.CheckTransDBActions(ModelID, DoActionTok.ActionStrData1)
                        REconnect()
                    End If
                    Me.Cursor = Cursors.Default
                End If
                editor.Dispose()
                editor = Nothing

            Case "DelCols"

                Dim editor As New Num_Popup("Delete Columns", ModelID, 3)

                editor.ShowDialog(Me)
                If editor.MyAction = DialogResult.OK Then
                    Me.Cursor = Cursors.WaitCursor
                    Dim NewRows As Integer = editor.MyValue
                    If NewRows > 0 Then
                        DisconnectRDS()
                        WorkbookManager.DeleteColumns(ModelID, DoActionTok.ActionStrData1, NewRows)
                        TransDBManager.CheckTransDBActions(ModelID, DoActionTok.ActionStrData1)
                        REconnect()
                    End If
                    Me.Cursor = Cursors.Default
                End If
                editor.Dispose()

            Case "Apply"

                Me.DialogResult = DialogResult.OK
                Me.Close()

        End Select

    End Sub
    Sub AddRows(NumRows As Integer)
        If Not NRds Is Nothing Then
            InsertRows(ModelID, DoActionTok.ActionStrData1, NumRows)
        End If
    End Sub
    Sub AddCols(NumCols As Integer)
        If Not NRds Is Nothing Then
            InsertColumns(ModelID, DoActionTok.ActionStrData1, NumCols)
        End If
    End Sub
    Sub DisconnectRDS()

        If Not NRds Is Nothing Then

            GridControl1.DataSource = Nothing
            GridControl1.DataBindings.Clear()
            GridControl1.Dispose()
            GridControl1 = Nothing
            NRds.Dispose()
            NRds = Nothing
            ActionRange = Nothing

        End If

    End Sub

    Sub REconnect()

        GridControl1 = New DevExpress.XtraGrid.GridControl
        GridControl1.Dock = DockStyle.Fill
        Me.Controls.Add(GridControl1)
        SetDataRange()
        ProcessGrid()

    End Sub

    Sub SetDataRange()

        ActionRange = FileManager.ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(DoActionTok.ActionStrData1).Range
        If ActionRange.RowCount < ActionRange.ColumnCount Then NROrient = "Cols"
        Dim RDSOptions As New RangeDataSourceOptions With {
            .UseFirstRowAsHeader = False,
            .PreserveFormulas = False,
            .SkipHiddenRows = True,
            .SkipHiddenColumns = True,
            .EditingOptions = DataSourceEditingOptions.AllowEdit
        }

        NRds = ActionRange.GetDataSource(RDSOptions)

        GridControl1.DataSource = NRds
        'GridControl1.RefreshDataSource()

    End Sub

    Public Sub ClearDss()

        GridControl1.DataSource = Nothing
        GridControl1.DataBindings.Clear()
        GridControl1.Dispose()
        GridControl1 = Nothing

        NRds.Dispose()
        NRds = Nothing
        ActionRange = Nothing
        Formatter = Nothing


    End Sub

    Protected Overrides Sub Finalize()

        GridControl1.DataSource = Nothing
        NRds = Nothing
        ActionRange = Nothing
        MyBase.Finalize()
        Me.Dispose()

    End Sub

    Private Sub TablePanel1_Paint(sender As Object, e As PaintEventArgs) Handles TablePanel1.Paint

    End Sub
End Class