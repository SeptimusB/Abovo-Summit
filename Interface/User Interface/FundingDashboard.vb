Imports System.Data
Imports System.Drawing
Imports System.Globalization
Imports System.Linq
Imports System.Windows.Forms
Imports Abovo
Imports Abovo.DataObject
Imports Abovo.FileManager
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils
Imports DevExpress.XtraCharts
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Grid

Public Class FundingDashboard
    Inherits XtraUserControl

    Private Const DashboardSheetName As String = "Funding Dashboard"
    Private Const ChartSourceSheetName As String = "OW - Charts Source Data"
    Private Const CovenantSourceSheetName As String = "OW - Live Covenant Calculation"
    Private Shared ReadOnly DashboardBlue As Color = Color.FromArgb(0, 91, 170)
    Private Shared ReadOnly DashboardHeader As Color = Color.FromArgb(89, 89, 89)
    Private Shared ReadOnly SeriesPalette As Color() = {
        Color.FromArgb(68, 114, 196),
        Color.FromArgb(237, 125, 49),
        Color.FromArgb(165, 165, 165),
        Color.FromArgb(255, 192, 0),
        Color.FromArgb(91, 155, 213),
        Color.FromArgb(112, 173, 71),
        Color.FromArgb(38, 68, 120),
        Color.FromArgb(244, 177, 131)
    }

    Private ReadOnly ModelID As Integer
    Private CalcEngID As Integer = -1
    Private IsRefreshingDashboard As Boolean
    Private IsUpdatingSelectors As Boolean
    Private RefreshPending As Boolean
    Private RootPanel As TableLayoutPanel

    Private NotInheritable Class CovenantVisualScale
        Public Property Maximum As Double
    End Class

    Public Sub New(SetModelID As Integer)
        ModelID = SetModelID
        Dock = DockStyle.Fill
        BackColor = Color.White
        BuildDashboard()

        CalcEngID = ExcelModels(ModelID).WBCalcEngine.AddActiveObject(Me)
        ExcelModels(ModelID).WBCalcEngine.AddActiveWorksheet(
            CalcEngID, Workbook.Worksheets(ChartSourceSheetName), False)
        ExcelModels(ModelID).WBCalcEngine.AddActiveWorksheet(
            CalcEngID, Workbook.Worksheets(CovenantSourceSheetName), False)
        ExcelModels(ModelID).WBCalcEngine.AddActiveWorksheet(
            CalcEngID, Workbook.Worksheets(DashboardSheetName), False)
        AddHandler Disposed, AddressOf FundingDashboard_Disposed
    End Sub

    Private ReadOnly Property Workbook As DevExpress.Spreadsheet.IWorkbook
        Get
            Return ExcelModels(ModelID).WB
        End Get
    End Property

    Private Sub BuildDashboard()
        SuspendLayout()
        Try
            DisposeDashboardControls()

            RootPanel = New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .BackColor = Color.White,
                .Padding = New Padding(12),
                .ColumnCount = 1,
                .RowCount = 3
            }
            RootPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            RootPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 44.0F))
            RootPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 58.0F))
            RootPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            RootPanel.Controls.Add(CreateHeading(), 0, 0)
            RootPanel.Controls.Add(CreateSelectorPanel(), 0, 1)
            RootPanel.Controls.Add(CreateDashboardBody(), 0, 2)
            Controls.Add(RootPanel)
        Finally
            ResumeLayout(True)
        End Try
    End Sub

    Private Function CreateHeading() As Control
        Dim DashboardSheet As Worksheet = Workbook.Worksheets(DashboardSheetName)
        Dim HeadingPanel As New PanelControl With {
            .Dock = DockStyle.Fill,
            .BorderStyle = BorderStyles.NoBorder,
            .BackColor = Color.White
        }
        Dim Heading As New LabelControl With {
            .Dock = DockStyle.Fill,
            .AutoSizeMode = LabelAutoSizeMode.None,
            .Text = DashboardSheet.Cells("A2").DisplayText
        }
        Heading.Appearance.Font = New Font("Segoe UI", 14.0F, FontStyle.Bold)
        Heading.Appearance.ForeColor = DashboardBlue
        Heading.Appearance.Options.UseFont = True
        Heading.Appearance.Options.UseForeColor = True
        Heading.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center
        Heading.ToolTip = DashboardSheet.Cells("A1").DisplayText
        HeadingPanel.Controls.Add(Heading)
        Return HeadingPanel
    End Function

    Private Function CreateSelectorPanel() As Control
        Dim Panel As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = False,
            .AutoScroll = True,
            .Padding = New Padding(4, 8, 4, 4)
        }
        AddSelector(Panel, "Display by Funder", "D6", "S", 175)
        AddSelector(Panel, "Display Facility", "I6", "S", 155)
        AddSelector(Panel, "Covenant Selector:", "R6", "S", 175)
        AddSelector(Panel, "From Year", "V6", "I", 95)
        Return Panel
    End Function

    Private Sub AddSelector(
        Parent As FlowLayoutPanel,
        Caption As String,
        CellAddress As String,
        DataFormat As String,
        EditorWidth As Integer)

        Dim SourceCell As Cell = Workbook.Worksheets(DashboardSheetName).Cells(CellAddress)
        Dim CaptionLabel As New LabelControl With {
            .Text = Caption,
            .AutoSizeMode = LabelAutoSizeMode.None,
            .Width = If(Caption = "Covenant Selector:", 135, 120),
            .Height = 28,
            .Margin = New Padding(10, 4, 4, 0)
        }
        CaptionLabel.Appearance.Font = New Font("Segoe UI", 10.0F)
        CaptionLabel.Appearance.ForeColor = Color.FromArgb(70, 70, 70)
        CaptionLabel.Appearance.Options.UseFont = True
        CaptionLabel.Appearance.Options.UseForeColor = True
        CaptionLabel.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center

        Dim Editor As New ComboBoxEdit With {
            .Width = EditorWidth,
            .Height = 28,
            .Margin = New Padding(0, 2, 18, 0),
            .Tag = New SelectorBinding(CellAddress, DataFormat)
        }
        Editor.Properties.TextEditStyle = TextEditStyles.DisableTextEditor
        Editor.Properties.Items.AddRange(WorkbookValidationItems(SourceCell))
        Editor.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Editor.Properties.Appearance.BackColor = WorkbookBackground(SourceCell)
        Editor.Properties.Appearance.ForeColor = WorkbookForeground(SourceCell)
        Editor.Properties.Appearance.Options.UseBackColor = True
        Editor.Properties.Appearance.Options.UseForeColor = True
        Editor.EditValue = CellToObject(SourceCell)
        AddHandler Editor.EditValueChanged, AddressOf Selector_EditValueChanged

        Parent.Controls.Add(CaptionLabel)
        Parent.Controls.Add(Editor)
    End Sub

    Private Function CreateDashboardBody() As Control
        Dim Body As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .ColumnCount = 2,
            .RowCount = 1,
            .Padding = New Padding(0, 4, 0, 0)
        }
        Body.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 64.0F))
        Body.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 36.0F))
        Body.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        Dim FundingCharts As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .ColumnCount = 2,
            .RowCount = 2,
            .Padding = New Padding(0, 0, 6, 0)
        }
        FundingCharts.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        FundingCharts.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        FundingCharts.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))
        FundingCharts.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))

        Dim DashboardSheet As Worksheet = Workbook.Worksheets(DashboardSheetName)
        Dim SelectedFunder As String = DashboardSheet.Cells("D6").DisplayText
        FundingCharts.Controls.Add(
            CreateChartCard("Funder: " & SelectedFunder & " - Drawn vs Available", CreateDrawnAvailableChart()), 0, 0)
        FundingCharts.Controls.Add(CreateChartCard("Loans by Funder", CreateLoansByFunderChart()), 1, 0)
        FundingCharts.Controls.Add(
            CreateChartCard("Funder: " & SelectedFunder & " - Loan Mix", CreateLoanMixChart()), 0, 1)
        FundingCharts.Controls.Add(
            CreateChartCard("Funder: " & SelectedFunder & " - Rates", CreateRatesChart()), 1, 1)

        Dim CovenantCharts As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .ColumnCount = 1,
            .RowCount = 3,
            .Padding = New Padding(6, 0, 0, 0)
        }
        CovenantCharts.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        CovenantCharts.RowStyles.Add(New RowStyle(SizeType.Percent, 56.0F))
        CovenantCharts.RowStyles.Add(New RowStyle(SizeType.Percent, 22.0F))
        CovenantCharts.RowStyles.Add(New RowStyle(SizeType.Percent, 22.0F))
        CovenantCharts.Controls.Add(
            CreateChartCard(DashboardSheet.Cells("R6").DisplayText, CreateSelectedCovenantView()), 0, 0)
        CovenantCharts.Controls.Add(
            CreateChartCard("Operating Margin", CreateStatusChart(40, 41)), 0, 1)
        CovenantCharts.Controls.Add(
            CreateChartCard("EBITDA MRI", CreateStatusChart(43, 44)), 0, 2)

        Body.Controls.Add(FundingCharts, 0, 0)
        Body.Controls.Add(CovenantCharts, 1, 0)
        Return Body
    End Function

    Private Function CreateChartCard(Title As String, Content As Control) As Control
        Dim Card As New PanelControl With {
            .Dock = DockStyle.Fill,
            .BorderStyle = BorderStyles.Simple,
            .BackColor = Color.White,
            .Margin = New Padding(4)
        }
        Dim Header As New LabelControl With {
            .Dock = DockStyle.Fill,
            .AutoSizeMode = LabelAutoSizeMode.None,
            .Text = Title
        }
        Header.Appearance.BackColor = DashboardHeader
        Header.Appearance.ForeColor = Color.White
        Header.Appearance.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        Header.Appearance.Options.UseBackColor = True
        Header.Appearance.Options.UseForeColor = True
        Header.Appearance.Options.UseFont = True
        Header.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Header.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center

        Dim CardLayout As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .ColumnCount = 1,
            .RowCount = 2,
            .Padding = New Padding(0),
            .Margin = New Padding(0)
        }
        CardLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        CardLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0F))
        CardLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        Content.Dock = DockStyle.Fill
        CardLayout.Controls.Add(Header, 0, 0)
        CardLayout.Controls.Add(Content, 0, 1)
        Card.Controls.Add(CardLayout)
        Return Card
    End Function

    Private Function CreateDrawnAvailableChart() As ChartControl
        Dim Chart As ChartControl = CreateBaseChart()
        AddColumnSeries(Chart, "Loans Drawn", 30, SeriesPalette(0), False)
        AddColumnSeries(Chart, "Unidentified Loans", 31, SeriesPalette(7), False)
        AddLineSeries(Chart, "Facilities Available", 29, DashboardBlue, False)
        ConfigureFundingChart(Chart, False)
        Return Chart
    End Function

    Private Function CreateLoansByFunderChart() As ChartControl
        Dim Chart As ChartControl = CreateBaseChart()
        Dim Source As Worksheet = Workbook.Worksheets(ChartSourceSheetName)
        Dim PaletteIndex As Integer
        For ColumnIndex As Integer = 8 To 15
            Dim SeriesName As String = Source.Cells(6, ColumnIndex).DisplayText.Trim()
            If SeriesName.Length = 0 OrElse SeriesName = "0" Then Continue For
            AddColumnSeries(Chart, SeriesName, ColumnIndex, SeriesPalette(PaletteIndex Mod SeriesPalette.Length), False)
            PaletteIndex += 1
        Next
        AddLineSeries(Chart, "Net Closing Debt", 16, DashboardBlue, False)
        ConfigureFundingChart(Chart, False)
        Return Chart
    End Function

    Private Function CreateLoanMixChart() As ChartControl
        Dim Chart As ChartControl = CreateBaseChart()
        AddColumnSeries(Chart, SourceHeader(18), 18, SeriesPalette(0), True)
        AddColumnSeries(Chart, SourceHeader(19), 19, SeriesPalette(7), True)
        ConfigureFundingChart(Chart, True)
        Return Chart
    End Function

    Private Function CreateRatesChart() As ChartControl
        Dim Chart As ChartControl = CreateBaseChart()
        AddLineSeries(Chart, SourceHeader(25), 25, SeriesPalette(0), True)
        AddLineSeries(Chart, SourceHeader(26), 26, SeriesPalette(7), True)
        AddLineSeries(Chart, SourceHeader(27), 27, SeriesPalette(1), True)
        ConfigureFundingChart(Chart, True)
        Return Chart
    End Function

    Private Function CreateBaseChart() As ChartControl
        Dim Chart As New ChartControl With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White
        }
        Chart.BorderOptions.Visibility = DefaultBoolean.False
        Return Chart
    End Function

    Private Sub AddColumnSeries(
        Chart As ChartControl,
        SeriesName As String,
        SourceColumn As Integer,
        SeriesColor As Color,
        FullStacked As Boolean)

        Dim View As ViewType = If(FullStacked, ViewType.FullStackedBar, ViewType.StackedBar)
        Dim Series As New DevExpress.XtraCharts.Series(SeriesName, View)
        AddSourcePoints(Series, SourceColumn)
        Series.LabelsVisibility = DefaultBoolean.False
        Series.View.Color = SeriesColor
        Chart.Series.Add(Series)
    End Sub

    Private Sub AddLineSeries(
        Chart As ChartControl,
        SeriesName As String,
        SourceColumn As Integer,
        SeriesColor As Color,
        PercentageValues As Boolean)

        Dim Series As New DevExpress.XtraCharts.Series(SeriesName, ViewType.Line)
        AddSourcePoints(Series, SourceColumn)
        Series.LabelsVisibility = DefaultBoolean.False
        Dim View As LineSeriesView = CType(Series.View, LineSeriesView)
        View.Color = SeriesColor
        View.MarkerVisibility = DefaultBoolean.False
        View.LineStyle.Thickness = 2
        Series.CrosshairLabelPattern = If(PercentageValues, "{S}: {V:p2}", "{S}: {V:n0}")
        Chart.Series.Add(Series)
    End Sub

    Private Sub AddSourcePoints(Series As DevExpress.XtraCharts.Series, SourceColumn As Integer)
        Dim Source As Worksheet = Workbook.Worksheets(ChartSourceSheetName)
        For RowIndex As Integer = 9 To 48
            Dim NumericValue As Double
            If TryGetNumericValue(Source.Cells(RowIndex, SourceColumn), NumericValue) Then
                Series.Points.Add(New SeriesPoint((RowIndex - 8).ToString(CultureInfo.InvariantCulture), NumericValue))
            End If
        Next
    End Sub

    Private Sub ConfigureFundingChart(Chart As ChartControl, PercentageAxis As Boolean)
        Chart.Legend.Visibility = DefaultBoolean.True
        Chart.Legend.AlignmentHorizontal = LegendAlignmentHorizontal.Center
        Chart.Legend.AlignmentVertical = LegendAlignmentVertical.TopOutside
        Chart.Legend.Direction = LegendDirection.LeftToRight
        Dim Diagram As XYDiagram = TryCast(Chart.Diagram, XYDiagram)
        If Diagram Is Nothing Then Return
        Diagram.AxisX.GridLines.Visible = False
        Diagram.AxisY.Interlaced = False
        Diagram.AxisY.GridLines.Visible = True
        Diagram.AxisY.Label.TextPattern = If(PercentageAxis, "{V:p0}", "{V:n0}")
    End Sub

    Private Function CreateSelectedCovenantView() As Control
        Dim Table As New DataTable()
        Table.Locale = CultureInfo.CurrentCulture
        Table.Columns.Add("Year", GetType(String))
        Table.Columns.Add("Value", GetType(String))
        Table.Columns.Add("Forecast", GetType(Double))
        Table.Columns.Add("Target", GetType(Double))
        Table.Columns.Add("Exceeded", GetType(Boolean))
        Dim DashboardSheet As Worksheet = Workbook.Worksheets(DashboardSheetName)
        Dim ChartSource As Worksheet = Workbook.Worksheets(ChartSourceSheetName)
        Dim MaximumValue As Double
        For RowIndex As Integer = 9 To 23
            Dim ForecastValue As Double
            Dim TargetValue As Double
            Dim ExceededValue As Double
            Dim HasForecast As Boolean =
                TryGetNumericValue(ChartSource.Cells(RowIndex, 34), ForecastValue)
            Dim HasTarget As Boolean =
                TryGetNumericValue(ChartSource.Cells(RowIndex, 35), TargetValue)
            Dim HasExceeded As Boolean =
                TryGetNumericValue(ChartSource.Cells(RowIndex, 37), ExceededValue)
            If HasForecast Then MaximumValue = Math.Max(MaximumValue, ForecastValue)
            If HasTarget Then MaximumValue = Math.Max(MaximumValue, TargetValue)
            Table.Rows.Add(
                DashboardSheet.Cells(RowIndex, 15).DisplayText,
                DashboardSheet.Cells(RowIndex, 16).DisplayText,
                If(HasForecast, CType(ForecastValue, Object), DBNull.Value),
                If(HasTarget, CType(TargetValue, Object), DBNull.Value),
                HasExceeded)
        Next

        Dim Grid As New GridControl With {.Dock = DockStyle.Fill}
        Grid.Tag = New CovenantVisualScale With {
            .Maximum = Math.Max(MaximumValue * 1.04R, 0.01R)
        }
        Dim View As New GridView(Grid)
        Grid.MainView = View
        Grid.ViewCollection.Add(View)
        View.OptionsBehavior.AutoPopulateColumns = False
        Dim YearColumn As New DevExpress.XtraGrid.Columns.GridColumn With {
            .FieldName = "Year",
            .Name = "FundingCovenantYearColumn",
            .VisibleIndex = 0
        }
        Dim ValueColumn As New DevExpress.XtraGrid.Columns.GridColumn With {
            .FieldName = "Value",
            .Name = "FundingCovenantValueColumn",
            .VisibleIndex = 1
        }
        Dim ForecastColumn As New DevExpress.XtraGrid.Columns.GridColumn With {
            .FieldName = "Forecast",
            .Name = "FundingCovenantVisualColumn",
            .VisibleIndex = 2
        }
        View.Columns.Add(YearColumn)
        View.Columns.Add(ValueColumn)
        View.Columns.Add(ForecastColumn)
        Grid.DataSource = Table
        View.OptionsBehavior.Editable = False
        View.OptionsBehavior.ReadOnly = True
        View.OptionsSelection.MultiSelect = True
        View.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect
        View.OptionsClipboard.CopyColumnHeaders = DefaultBoolean.False
        View.OptionsView.ShowColumnHeaders = False
        View.OptionsView.ShowGroupPanel = False
        View.OptionsView.ShowIndicator = False
        View.OptionsView.ShowHorizontalLines = DefaultBoolean.False
        View.OptionsView.ShowVerticalLines = DefaultBoolean.False
        View.OptionsView.RowAutoHeight = False
        View.RowHeight = 28
        View.Appearance.Row.Font = New Font("Segoe UI", 9.0F)
        View.Appearance.Row.ForeColor = Color.FromArgb(70, 70, 70)
        View.Appearance.Row.Options.UseFont = True
        View.Appearance.Row.Options.UseForeColor = True
        AddHandler View.RowCellStyle, AddressOf CovenantGridRowCellStyle
        AddHandler View.CustomDrawCell, AddressOf CovenantGridCustomDrawCell
        If YearColumn IsNot Nothing Then
            YearColumn.VisibleIndex = 0
            YearColumn.Width = 90
            YearColumn.OptionsColumn.FixedWidth = True
        End If
        If ValueColumn IsNot Nothing Then
            ValueColumn.VisibleIndex = 1
            ValueColumn.Width = 76
            ValueColumn.OptionsColumn.FixedWidth = True
            ValueColumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        End If
        If ForecastColumn IsNot Nothing Then
            ForecastColumn.VisibleIndex = 2
            ForecastColumn.MinWidth = 220
            ForecastColumn.OptionsColumn.AllowFocus = True
        End If
        For Each Column As DevExpress.XtraGrid.Columns.GridColumn In View.Columns
            Column.OptionsColumn.AllowEdit = False
            Column.OptionsColumn.ReadOnly = True
            Column.OptionsFilter.AllowFilter = False
            Column.OptionsColumn.AllowSort = DefaultBoolean.False
        Next

        Return Grid
    End Function

    Private Sub CovenantGridRowCellStyle(sender As Object, e As RowCellStyleEventArgs)
        If e.RowHandle < 0 OrElse e.Column Is Nothing Then Return
        If e.Column.FieldName <> "Year" AndAlso e.Column.FieldName <> "Value" Then Return
        Dim SourceColumn As Integer = If(e.Column.FieldName = "Year", 15, 16)
        Dim SourceCell As Cell = Workbook.Worksheets(DashboardSheetName).Cells(9 + e.RowHandle, SourceColumn)
        ApplyWorkbookCellAppearance(e.Appearance, SourceCell)
    End Sub

    Private Sub CovenantGridCustomDrawCell(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs)

        If e.RowHandle < 0 OrElse e.Column Is Nothing OrElse
           e.Column.FieldName <> "Forecast" Then Return
        Dim View As GridView = TryCast(sender, GridView)
        If View Is Nothing Then Return
        Dim Grid As GridControl = TryCast(View.GridControl, GridControl)
        Dim Scale As CovenantVisualScale =
            If(Grid Is Nothing, Nothing, TryCast(Grid.Tag, CovenantVisualScale))
        If Scale Is Nothing OrElse Scale.Maximum <= 0 Then Return
        Dim SourceRow As DataRowView = TryCast(View.GetRow(e.RowHandle), DataRowView)
        If SourceRow Is Nothing Then Return

        Dim DrawingGraphics As Graphics = e.Cache.Graphics
        Using BackgroundBrush As New SolidBrush(e.Appearance.BackColor)
            DrawingGraphics.FillRectangle(BackgroundBrush, e.Bounds)
        End Using
        Dim PlotBounds As Rectangle = Rectangle.Inflate(e.Bounds, -7, -5)
        If PlotBounds.Width <= 0 OrElse PlotBounds.Height <= 0 Then
            e.Handled = True
            Return
        End If

        Dim ForecastObject As Object = SourceRow("Forecast")
        If ForecastObject IsNot Nothing AndAlso ForecastObject IsNot DBNull.Value Then
            Dim ForecastValue As Double = Convert.ToDouble(
                ForecastObject, CultureInfo.InvariantCulture)
            Dim BarWidth As Integer = CInt(Math.Round(
                Math.Max(0, Math.Min(1, ForecastValue / Scale.Maximum)) *
                PlotBounds.Width))
            If BarWidth > 0 Then
                Dim ExceededObject As Object = SourceRow("Exceeded")
                Dim IsExceeded As Boolean =
                    ExceededObject IsNot Nothing AndAlso
                    ExceededObject IsNot DBNull.Value AndAlso
                    Convert.ToBoolean(ExceededObject, CultureInfo.InvariantCulture)
                Using BarBrush As New SolidBrush(
                    If(IsExceeded, Color.Red, SeriesPalette(0)))
                    DrawingGraphics.FillRectangle(
                        BarBrush,
                        New Rectangle(
                            PlotBounds.Left,
                            PlotBounds.Top + 3,
                            BarWidth,
                            Math.Max(1, PlotBounds.Height - 6)))
                End Using
            End If
        End If

        Dim TargetObject As Object = SourceRow("Target")
        If TargetObject IsNot Nothing AndAlso TargetObject IsNot DBNull.Value Then
            Dim TargetValue As Double = Convert.ToDouble(
                TargetObject, CultureInfo.InvariantCulture)
            Dim TargetX As Integer = PlotBounds.Left + CInt(Math.Round(
                Math.Max(0, Math.Min(1, TargetValue / Scale.Maximum)) *
                PlotBounds.Width))
            TargetX = Math.Max(PlotBounds.Left + 5, Math.Min(PlotBounds.Right - 5, TargetX))
            Dim MarkerTop As Integer = PlotBounds.Top
            Dim MarkerPoints As Point() = {
                New Point(TargetX, MarkerTop),
                New Point(TargetX - 5, MarkerTop + 9),
                New Point(TargetX + 5, MarkerTop + 9)
            }
            Using MarkerBrush As New SolidBrush(Color.FromArgb(0, 176, 80))
                DrawingGraphics.FillPolygon(MarkerBrush, MarkerPoints)
            End Using
        End If
        e.Handled = True
    End Sub

    Private Function CreateStatusChart(MetColumn As Integer, BreachedColumn As Integer) As ChartControl
        Dim Chart As ChartControl = CreateBaseChart()
        Dim Source As Worksheet = Workbook.Worksheets(ChartSourceSheetName)
        Dim Met As New DevExpress.XtraCharts.Series("Covenant Met", ViewType.Point)
        Dim Breached As New DevExpress.XtraCharts.Series("Covenant Breached", ViewType.Point)

        For RowIndex As Integer = 9 To 48
            Dim Value As Double
            If TryGetNumericValue(Source.Cells(RowIndex, MetColumn), Value) Then
                Met.Points.Add(New SeriesPoint(RowIndex - 8, Value))
            End If
            If TryGetNumericValue(Source.Cells(RowIndex, BreachedColumn), Value) Then
                Breached.Points.Add(New SeriesPoint(RowIndex - 8, Value))
            End If
        Next

        ConfigureStatusSeries(Met, Color.FromArgb(0, 176, 80), MarkerKind.Triangle)
        ConfigureStatusSeries(Breached, Color.Red, MarkerKind.Cross)
        Chart.Series.Add(Met)
        Chart.Series.Add(Breached)
        Chart.Legend.Visibility = DefaultBoolean.False

        Dim Diagram As XYDiagram = TryCast(Chart.Diagram, XYDiagram)
        If Diagram IsNot Nothing Then
            Diagram.AxisY.WholeRange.SetMinMaxValues(0, 2.2)
            Diagram.AxisY.VisualRange.SetMinMaxValues(0, 2.2)
            Diagram.AxisY.Visibility = DefaultBoolean.False
            Diagram.AxisY.GridLines.Visible = False
            Diagram.AxisX.GridLines.Visible = False
            Diagram.AxisX.Tickmarks.MinorVisible = False
            Diagram.AxisX.NumericScaleOptions.AutoGrid = False
            Diagram.AxisX.NumericScaleOptions.GridSpacing = 10
        End If
        Return Chart
    End Function

    Private Shared Sub ConfigureStatusSeries(
        Series As DevExpress.XtraCharts.Series,
        MarkerColor As Color,
        Marker As MarkerKind)

        Dim View As PointSeriesView = CType(Series.View, PointSeriesView)
        View.PointMarkerOptions.Size = 10
        View.PointMarkerOptions.Kind = Marker
        View.Color = MarkerColor
        Series.LabelsVisibility = DefaultBoolean.False
    End Sub

    Private Function SourceHeader(ColumnIndex As Integer) As String
        Dim Header As String = Workbook.Worksheets(ChartSourceSheetName).Cells(6, ColumnIndex).DisplayText.Trim()
        Return If(Header.Length = 0 OrElse Header = "0", "Series " & (ColumnIndex + 1).ToString(), Header)
    End Function

    Private Sub Selector_EditValueChanged(sender As Object, e As EventArgs)
        If IsRefreshingDashboard OrElse IsUpdatingSelectors Then Return
        Dim Editor As ComboBoxEdit = TryCast(sender, ComboBoxEdit)
        Dim Binding As SelectorBinding = If(Editor Is Nothing, Nothing, TryCast(Editor.Tag, SelectorBinding))
        If Binding Is Nothing Then Return

        Dim Target As Cell = Workbook.Worksheets(DashboardSheetName).Cells(Binding.CellAddress)
        If Target.Protection.Locked Then
            RefreshData()
            Return
        End If
        Dim ChangedValue As Object = Editor.EditValue
        If Binding.DataFormat = "I" Then
            Dim ParsedValue As Integer
            If Not Integer.TryParse(Convert.ToString(ChangedValue, CultureInfo.CurrentCulture), ParsedValue) Then
                RefreshData()
                Return
            End If
            ChangedValue = ParsedValue
        End If
        If ValuesEqual(CellToObject(Target), ChangedValue) Then Return

        IsUpdatingSelectors = True
        RefreshPending = False
        Try
            Dim Change As New DataChangeEvent With {
                .ModelID = ModelID,
                .Description = "Funding Dashboard selector updated",
                .WSName = DashboardSheetName,
                .CellAddress = Target.GetReferenceA1(),
                .OriginalValue = CellToObject(Target),
                .ChangedValue = ChangedValue,
                .DataFormat = Binding.DataFormat,
                .TimeStamp = Now(),
                .UserName = Environment.UserName
            }
            Dim Result = ExcelModels(ModelID).ChangeManager.ProcessChange(Change)
            If Result Is Nothing OrElse Result.BError Then RefreshPending = True
        Finally
            IsUpdatingSelectors = False
        End Try
        If RefreshPending Then RefreshPending = False
        RefreshData()
    End Sub

    Public Sub RefreshData()
        If IsDisposed OrElse Disposing Then Return
        If InvokeRequired Then
            BeginInvoke(New Action(AddressOf RefreshData))
            Return
        End If
        If IsUpdatingSelectors Then
            RefreshPending = True
            Return
        End If
        If IsRefreshingDashboard Then Return

        IsRefreshingDashboard = True
        Try
            BuildDashboard()
        Finally
            IsRefreshingDashboard = False
        End Try
    End Sub

    Private Sub DisposeDashboardControls()
        For Each Existing As Control In Controls.Cast(Of Control)().ToArray()
            Controls.Remove(Existing)
            Existing.Dispose()
        Next
    End Sub

    Private Sub FundingDashboard_Disposed(sender As Object, e As EventArgs)
        If CalcEngID < 0 Then Return
        If ExcelModels IsNot Nothing AndAlso
           ModelID >= 0 AndAlso ModelID < ExcelModels.Length AndAlso
           ExcelModels(ModelID) IsNot Nothing AndAlso
           ExcelModels(ModelID).WBCalcEngine IsNot Nothing Then
            ExcelModels(ModelID).WBCalcEngine.RemoveActiveObject(CalcEngID)
        End If
        CalcEngID = -1
    End Sub

    Private Function WorkbookValidationItems(SourceCell As Cell) As List(Of String)
        Dim Result As New List(Of String)()
        Dim Validation As DataValidation = SourceCell.Worksheet.DataValidations.GetDataValidation(SourceCell)
        If Validation Is Nothing OrElse Validation.ValidationType <> DataValidationType.List Then Return Result
        Dim Criteria As ValueObject = Validation.Criteria
        If Criteria Is Nothing OrElse Criteria.IsEmpty Then Return Result

        If Criteria.IsText Then
            For Each Item As String In Criteria.TextValue.Split(New Char() {","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                AddValidationItem(Result, Item)
            Next
        ElseIf Criteria.IsRange Then
            AddValidationRangeItems(Result, Criteria.RangeValue)
        ElseIf Criteria.IsFormula Then
            Dim Reference As String = Criteria.FormulaInvariant.Trim().TrimStart("="c)
            Try
                If Workbook.DefinedNames.Contains(Reference) Then
                    AddValidationRangeItems(Result, Workbook.DefinedNames.GetDefinedName(Reference).Range)
                Else
                    AddValidationRangeItems(Result, Workbook.Range(Reference))
                End If
            Catch
                'The workbook remains authoritative when an optional validation source cannot be resolved.
            End Try
        End If
        Return Result
    End Function

    Private Shared Sub AddValidationRangeItems(Result As List(Of String), Source As CellRange)
        If Source Is Nothing Then Return
        For RowIndex As Integer = 0 To Source.RowCount - 1
            For ColumnIndex As Integer = 0 To Source.ColumnCount - 1
                AddValidationItem(Result, Source(RowIndex, ColumnIndex).DisplayText)
            Next
        Next
    End Sub

    Private Shared Sub AddValidationItem(Result As List(Of String), Value As String)
        Dim Item As String = If(Value, String.Empty).Trim()
        If Item.Length > 0 AndAlso Not Result.Contains(Item) Then Result.Add(Item)
    End Sub

    Private Shared Function TryGetNumericValue(SourceCell As Cell, ByRef Value As Double) As Boolean
        Value = 0
        If SourceCell Is Nothing OrElse Not SourceCell.Value.IsNumeric Then Return False
        Value = SourceCell.Value.NumericValue
        Return Not Double.IsNaN(Value) AndAlso Not Double.IsInfinity(Value)
    End Function

    Private Shared Function CellToObject(SourceCell As Cell) As Object
        If SourceCell Is Nothing OrElse SourceCell.Value.IsEmpty Then Return Nothing
        If SourceCell.Value.IsNumeric Then Return SourceCell.Value.NumericValue
        If SourceCell.Value.IsBoolean Then Return SourceCell.Value.BooleanValue
        If SourceCell.Value.IsDateTime Then Return SourceCell.Value.DateTimeValue
        Return SourceCell.Value.TextValue
    End Function

    Private Shared Function ValuesEqual(Original As Object, Changed As Object) As Boolean
        If Original Is Nothing AndAlso Changed Is Nothing Then Return True
        If Original Is Nothing OrElse Changed Is Nothing Then Return False
        Dim OriginalNumber As Double
        Dim ChangedNumber As Double
        If Double.TryParse(Convert.ToString(Original, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, OriginalNumber) AndAlso
           Double.TryParse(Convert.ToString(Changed, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, ChangedNumber) Then
            Return Math.Abs(OriginalNumber - ChangedNumber) < 0.0000001R
        End If
        Return String.Equals(Convert.ToString(Original), Convert.ToString(Changed), StringComparison.Ordinal)
    End Function

    Private Shared Function WorkbookBackground(SourceCell As Cell) As Color
        Dim Result As Color = SourceCell.FillColor
        If Result.IsEmpty OrElse Result.A = 0 Then Result = Color.White
        Return Result
    End Function

    Private Shared Sub ApplyWorkbookCellAppearance(
        Appearance As DevExpress.Utils.AppearanceObject,
        SourceCell As Cell)

        If Appearance Is Nothing OrElse SourceCell Is Nothing Then Return
        Appearance.BackColor = WorkbookBackground(SourceCell)
        Appearance.ForeColor = WorkbookForeground(SourceCell)
        Appearance.Options.UseBackColor = True
        Appearance.Options.UseForeColor = True
        Appearance.Options.UseTextOptions = True
        Select Case SourceCell.Alignment.Horizontal
            Case SpreadsheetHorizontalAlignment.Center
                Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            Case SpreadsheetHorizontalAlignment.Right
                Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            Case Else
                Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
        End Select
        Dim Style As FontStyle = FontStyle.Regular
        If SourceCell.Font.Bold Then Style = Style Or FontStyle.Bold
        If SourceCell.Font.Italic Then Style = Style Or FontStyle.Italic
        If SourceCell.Font.UnderlineType <> UnderlineType.None Then Style = Style Or FontStyle.Underline
        Appearance.Font = New Font(Appearance.Font.FontFamily, Appearance.Font.Size, Style)
        Appearance.Options.UseFont = True
    End Sub

    Private Shared Function WorkbookForeground(SourceCell As Cell) As Color
        Dim Result As Color = SourceCell.Font.Color
        If Result.IsEmpty OrElse Result.A = 0 Then Result = Color.Black
        Return Result
    End Function

    Private NotInheritable Class SelectorBinding
        Public ReadOnly CellAddress As String
        Public ReadOnly DataFormat As String

        Public Sub New(SetCellAddress As String, SetDataFormat As String)
            CellAddress = SetCellAddress
            DataFormat = SetDataFormat
        End Sub
    End Class
End Class
