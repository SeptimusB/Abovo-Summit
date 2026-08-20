Imports System.IO
Imports System.Windows.Forms
Imports Abovo
Imports Abovo.ExportServices
Imports Abovo.FileManager
Imports Abovo.GeneralFunctions
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraEditors
Imports DevExpress.XtraTab

Public Class FFRForm
    Private Shared ReadOnly FFRSheetNames As String() = {
        "FFR Validation Summary",
        "Front Sheet",
        "FFR Inputs Adj Stmt",
        "FFR Workings",
        "Statements",
        "Assumptions & tenure inputs",
        "Compliance Questions",
        "FFR Key Defn"
    }

    Private ReadOnly ModelID As Integer
    Private ReadOnly SheetViews As New Dictionary(Of XtraTabPage, Control)()
    Private ClosingForDisposal As Boolean

    Public Sub New(SetModelID As Integer)
        InitializeComponent()
        ModelID = SetModelID
        Text = "Financial Forecast Return for " & ExcelModels(ModelID).WBStructure.CompanyName
        SheetCaption.Text = Text
        BuildSheetTabs()
        EnsureSelectedSheetBuilt()
    End Sub

    'Retained for callers compiled against the former form API.
    Public Sub Initialise()
        EnsureSelectedSheetBuilt()
    End Sub

    Public Sub SetMode(ExMode As String)
        'The native FFR surface is always workbook-driven; there is no separate
        'preview/export mode to maintain outside the model.
    End Sub

    Private Sub BuildSheetTabs()
        FFRTabs.BeginUpdate()
        Try
            FFRTabs.TabPages.Clear()
            For Each SheetName As String In FFRSheetNames
                Dim Page As New XtraTabPage With {
                    .Name = "FFR_" & SheetName.Replace(" ", "_").Replace("&", "And"),
                    .Text = SheetName,
                    .Tag = SheetName
                }
                FFRTabs.TabPages.Add(Page)
            Next
            If FFRTabs.TabPages.Count > 0 Then FFRTabs.SelectedTabPageIndex = 0
        Finally
            FFRTabs.EndUpdate()
        End Try
    End Sub

    Private Sub EnsureSelectedSheetBuilt()
        Dim Page As XtraTabPage = FFRTabs.SelectedTabPage
        If Page Is Nothing Then Return

        Dim SheetName As String = Convert.ToString(Page.Tag)
        If String.IsNullOrWhiteSpace(SheetName) Then Return
        SheetCaption.Text = "Financial Forecast Return  •  " & SheetName

        Dim ExistingView As Control = Nothing
        If SheetViews.TryGetValue(Page, ExistingView) Then
            If TypeOf ExistingView Is FFRValidationSummaryView Then
                DirectCast(ExistingView, FFRValidationSummaryView).RefreshFromWorkbook()
            ElseIf TypeOf ExistingView Is FFRFrontSheetView Then
                DirectCast(ExistingView, FFRFrontSheetView).RefreshFromWorkbook()
            ElseIf TypeOf ExistingView Is FFRInputsAdjStmtView Then
                DirectCast(ExistingView, FFRInputsAdjStmtView).RefreshFromWorkbook()
            Else
                DirectCast(ExistingView, FFRWorkbookSheetView).RefreshFromWorkbook()
            End If
            Return
        End If

        If Not FileManager.GetWorkBook(ModelID).Worksheets.Contains(SheetName) Then
            Dim Missing As New LabelControl With {
                .Dock = DockStyle.Fill,
                .AutoSizeMode = LabelAutoSizeMode.None,
                .Text = "The workbook does not contain the required sheet '" & SheetName & "'."
            }
            Missing.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            Missing.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
            Page.Controls.Add(Missing)
            Return
        End If

        Cursor = Cursors.WaitCursor
        Page.SuspendLayout()
        Try
            Dim View As Control
            If String.Equals(SheetName, "FFR Validation Summary", StringComparison.Ordinal) Then
                View = New FFRValidationSummaryView(ModelID) With {.Dock = DockStyle.Fill}
            ElseIf String.Equals(SheetName, "Front Sheet", StringComparison.Ordinal) Then
                Dim FrontSheetView As New FFRFrontSheetView(ModelID) With {.Dock = DockStyle.Fill}
                AddHandler FrontSheetView.WorkbookCellChanged, AddressOf WorkbookCellChanged
                View = FrontSheetView
            ElseIf String.Equals(SheetName, "FFR Inputs Adj Stmt", StringComparison.Ordinal) Then
                Dim InputsView As New FFRInputsAdjStmtView(ModelID) With {.Dock = DockStyle.Fill}
                AddHandler InputsView.WorkbookCellChanged, AddressOf WorkbookCellChanged
                View = InputsView
            Else
                Dim WorkbookView As New FFRWorkbookSheetView(ModelID, SheetName) With {.Dock = DockStyle.Fill}
                AddHandler WorkbookView.WorkbookCellChanged, AddressOf WorkbookCellChanged
                View = WorkbookView
            End If
            SheetViews.Add(Page, View)
            Page.Controls.Add(View)
        Catch ex As Exception
            Page.Controls.Clear()
            Dim Failure As New MemoEdit With {
                .Dock = DockStyle.Fill,
                .Text = "The FFR sheet could not be loaded." & Environment.NewLine & Environment.NewLine & ex.Message
            }
            Failure.Properties.ReadOnly = True
            Page.Controls.Add(Failure)
        Finally
            Page.ResumeLayout()
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub WorkbookCellChanged(sender As Object, e As EventArgs)
        'Inactive views deliberately remain snapshots.  They reload from the
        'calculated workbook when selected, avoiding eight full redraws per edit.
        Dim ChangedSheetName As String
        If TypeOf sender Is FFRFrontSheetView Then
            ChangedSheetName = DirectCast(sender, FFRFrontSheetView).WorksheetName
        ElseIf TypeOf sender Is FFRInputsAdjStmtView Then
            ChangedSheetName = DirectCast(sender, FFRInputsAdjStmtView).WorksheetName
        Else
            ChangedSheetName = DirectCast(sender, FFRWorkbookSheetView).WorksheetName
        End If
        SheetCaption.Text = "Financial Forecast Return  •  " & ChangedSheetName
    End Sub

    Private Sub FFRTabsSelectedPageChanged(sender As Object, e As TabPageChangedEventArgs) Handles FFRTabs.SelectedPageChanged
        EnsureSelectedSheetBuilt()
    End Sub

    Private Sub RefreshButtonClick(sender As Object, e As EventArgs) Handles RefreshButton.Click
        EnsureSelectedSheetBuilt()
    End Sub

    Private Sub CloseButtonClick(sender As Object, e As EventArgs) Handles CloseButton.Click
        Hide()
    End Sub

    Private Sub CreateReturnButtonClick(sender As Object, e As EventArgs) Handles CreateReturnButton.Click
        CreateFFRReturn()
    End Sub

    Private Sub CreateFFRReturn()
        Dim SourceWorkbook As IWorkbook = FileManager.GetWorkBook(ModelID)
        If SourceWorkbook Is Nothing Then Return

        Using OpenDialog As New OpenFileDialog With {
            .Filter = "FFR macro-enabled workbooks|*.xlsm;*.xlsb|Excel workbooks|*.xlsx;*.xls",
            .Title = "Select the provider-specific FFR template"
        }
            If OpenDialog.ShowDialog(Me) <> DialogResult.OK Then Return

            Dim ReturnWorkbook As New Workbook()
            Try
                Cursor = Cursors.WaitCursor
                ReturnWorkbook.Options.CalculationMode = WorkbookCalculationMode.Manual
                ReturnWorkbook.Options.CalculationEngineType = CalculationEngineType.ChainBased
                ReturnWorkbook.DocumentSettings.Calculation.EnableMultiThreading = False
                ReturnWorkbook.LoadDocument(OpenDialog.FileName)

                If Not ReturnWorkbook.Worksheets.Contains("Cover Sheet") OrElse
                   Not String.Equals(
                       ReturnWorkbook.Worksheets("Cover Sheet").Cells("B4").DisplayText,
                       "Spreadsheet Import Template - Financial Forecast Return (FFR)",
                       StringComparison.Ordinal) Then
                    XtraMessageBox.Show(
                        Me,
                        "The selected file is not a provider-specific Financial Forecast Return template.",
                        "Create FFR return",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
                    Return
                End If

                Dim MappingCount As Integer = SourceWorkbook.Range("FFRRangeNames").RowCount
                ReturnWorkbook.BeginUpdate()
                Try
                    For MappingIndex As Integer = 1 To MappingCount - 1
                        Dim SourceName As String =
                            SourceWorkbook.Range("FFRListHeading")(MappingIndex, 8).DisplayText.Trim()
                        Dim DestinationName As String =
                            SourceWorkbook.Range("FFRListHeading")(MappingIndex, 3).DisplayText.Trim()
                        If SourceName.Length = 0 OrElse DestinationName.Length = 0 Then Continue For

                        ReturnWorkbook.Range(DestinationName).CopyFrom(
                            SourceWorkbook.Range(SourceName), PasteSpecial.Values)
                    Next
                Finally
                    ReturnWorkbook.EndUpdate()
                End Try

                Using SaveDialog As New SaveFileDialog With {
                    .Filter = "Excel macro-enabled workbook|*.xlsm",
                    .DefaultExt = "xlsm",
                    .AddExtension = True,
                    .OverwritePrompt = True,
                    .Title = "Save the completed FFR return"
                }
                    If SaveDialog.ShowDialog(Me) <> DialogResult.OK Then Return
                    ReturnWorkbook.SaveDocument(SaveDialog.FileName, DocumentFormat.Xlsm)

                    If XtraMessageBox.Show(
                            Me,
                            "The FFR return was saved successfully. Open it in Microsoft Excel?",
                            "Create FFR return",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information) = DialogResult.Yes Then
                        OpenFileInExcel(SaveDialog.FileName)
                    End If
                End Using
            Catch ex As Exception
                XtraMessageBox.Show(
                    Me,
                    "The FFR return could not be created: " & ex.Message,
                    "Create FFR return",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            Finally
                Cursor = Cursors.Default
                ReturnWorkbook.Dispose()
            End Try
        End Using
    End Sub

    Public Sub ManualDispose()
        ClosingForDisposal = True
        For Each View As Control In SheetViews.Values
            If TypeOf View Is FFRValidationSummaryView Then
                'Read-only workbook summary has no edit event to detach.
            ElseIf TypeOf View Is FFRFrontSheetView Then
                RemoveHandler DirectCast(View, FFRFrontSheetView).WorkbookCellChanged, AddressOf WorkbookCellChanged
            ElseIf TypeOf View Is FFRInputsAdjStmtView Then
                RemoveHandler DirectCast(View, FFRInputsAdjStmtView).WorkbookCellChanged, AddressOf WorkbookCellChanged
            Else
                RemoveHandler DirectCast(View, FFRWorkbookSheetView).WorkbookCellChanged, AddressOf WorkbookCellChanged
            End If
            View.Dispose()
        Next
        SheetViews.Clear()
        Close()
        Dispose()
    End Sub

    Private Sub FFRFormFormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If ClosingForDisposal Then Return
        e.Cancel = True
        Hide()
    End Sub
End Class
