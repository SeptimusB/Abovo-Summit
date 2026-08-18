

Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports DevExpress.XtraCharts
Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.LogDebugDev
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.Utils.Drawing
Imports DevExpress.Data
Imports Abovo.DataObject
Imports DevExpress.XtraBars.Navigation
Imports DevExpress.XtraPrinting
Imports Abovo.PresentationManager
Imports Abovo.AbovoUnboundSource
Imports DevExpress.XtraLayout.Customization
Imports DevExpress.DataAccess.DataFederation
Imports DevExpress.UnitConversion
Imports DevExpress.Charts.Model
Imports DevExpress.Drawing
Imports DevExpress.Spreadsheet.Charts
Imports DevExpress.Snap.Core.Native



Public Class BPDashboard

    Public DataPM As PresentationManager
    Public DataPres As PresentationManager.DataPresentation
    Public ModelID As Integer
    Public PresID As Integer
    Public GSID As Integer
    Public CSID As Integer
    Public rs As New Resizer
    Private BIsDirty As Boolean
    Private Grid1ExpandedView As Boolean
    Public ScaleUnits As Single
    Public Scalefactor As Single
    Public DataSources() As AbovoUnboundSource
    Public DataSourceCount As Integer
    Public PresentedDS As Abovo.DataObject.DataCellRange
    Public PresentedColumn As Abovo.DataObject.SheetDataColumn
    Public PropertyArray() As UnboundSourceProperty
    Public PropertiesCount As Integer
    Public PropertyList As IEnumerable(Of UnboundSourceProperty)
    Public PropType As System.Type
    Public ColList As List(Of String)
    Public ColCount As Integer = 0
    Public ColName As String
    Public GridControls() As GridControl
    Public GridCount As Integer = -1

    Public AcControl As AccordionControl
    Public AcElements() As AccordionControlElement
    Public AcElementlist As List(Of AccordionControlElement)
    Public AcControlCount As Integer = -1
    Public AcElementCount As Integer = -1
    Public AcContainers() As AccordionContentContainer
    Public hyperlinkLabelControls() As HyperlinkLabelControl
    Public AcContainersCount As Integer = -1
    Public UsedGridViews(-1) As GridView
    Public GridViewCount As Integer = -1
    Public Formatter As ObjectFormatter
    Public CurrChartWS As DevExpress.Spreadsheet.Worksheet
    Public StockChart As ChartControl
    Public IncomePieChart As ChartControl
    Public ExpensesPieChart As ChartControl
    Public DebtLineChart As ChartControl
    Public CvntChart As ChartControl
    Public CvntChart2 As ChartControl
    Public CvntChart3 As ChartControl
    Public CvntChart4 As ChartControl
    Public CvntChart5 As ChartControl
    Public GearingBrowserMsg As WebBrowser


    Public Sub New(SetModelID As Integer)

        DataSourceCount = -1
        ' This call is required by the designer.
        InitializeComponent()
        Formatter = New ObjectFormatter

        ModelID = SetModelID

        'DataPM = ExcelModels(SetModelID).WBDataPres
        'Dim CheckTrans As AbovoTransaction
        'CheckTrans = DataPM.ValidatePresentation(GSID, CSID)

        'If CheckTrans.BError = False Then

        '    PresID = CheckTrans.IntegerReturn
        '    DataPres = DataPM.DataPresentations(PresID)

        'End If

        CurrChartWS = ExcelModels(ModelID).WB.Worksheets("OW - Charts Source Data")
        'ManagementCosts.ManagementCostData.GetStatus()
        AddStocksChart()
        AddBrowserGearingTable()
        AddDebtChart()

        AddOperatingIncomeChart()
        AddExpensesPieChart()
        AddCvntChart()
        Add2ndCovChart()
        Add3rdCovChart()
        Add4thCovChart()
        Add5thCovChart()
        Exit Sub


    End Sub
    Sub AddBrowserGearingTable()

        GearingBrowserMsg = New WebBrowser
        Dim OutputMsg As String = ""
        Dim DBWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets("BP Dashboard")
        Dim BPDBWS As DevExpress.Spreadsheet.CellRange = DBWS.Range("H9:L23")
        Dim CellExamine As DevExpress.Spreadsheet.Cell

        OutputMsg = "<html><body><table style=width:100% style=border:1px><B><tr>" +
        "<th bgcolor='white'></th>" +
        "<th bgcolor='grey'><center><p style='font-family:verdana;color:white'>Target</th>" +
        "<th bgcolor='grey'><center><p style='font-family:verdana;color:white'>Base</th>" +
        "<th bgcolor='grey'><center><p style='font-family:verdana;color:white'>Breaches</th>" +
        "</tr><tr>"

        CellExamine = BPDBWS(0, 0)
        OutputMsg += "<td><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(0, 2)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</p></td>"
        CellExamine = BPDBWS(0, 3)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(0, 4)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td></tr><tr>"

        CellExamine = BPDBWS(1, 0)
        OutputMsg += "<td><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(1, 2)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(1, 3)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(1, 4)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td></tr><tr></tr><td> </td><tr>"

        CellExamine = BPDBWS(3, 0)
        OutputMsg += "<td><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(3, 2)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(3, 3)
        OutputMsg += "<td><center><p style='font-family:verdana'>       " & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(3, 4)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td></tr><tr>"

        CellExamine = BPDBWS(4, 0)
        OutputMsg += "<td><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(4, 2)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(4, 3)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(4, 4)
        OutputMsg += "<td><center><p style='font-family:verdana'>       " & CellExamine.DisplayText & "</td></tr><tr></tr><td> </td><tr>"

        CellExamine = BPDBWS(6, 0)
        OutputMsg += "<td><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(6, 2)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(6, 3)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(6, 4)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td></tr><tr>"
        CellExamine = BPDBWS(7, 0)
        OutputMsg += "<td><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(7, 2)
        OutputMsg += "<td><center><p style='font-family:verdana'>   " & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(7, 3)
        OutputMsg += "<td><center><p style='font-family:verdana'>   " & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(7, 4)
        OutputMsg += "<td><center> <p style='font-family:verdana'> " & CellExamine.DisplayText & "</td></tr><tr></tr><td> </td><tr>"
        CellExamine = BPDBWS(9, 0)
        OutputMsg += "<td><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(9, 2)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(9, 3)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(9, 4)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td></tr><tr>"
        CellExamine = BPDBWS(10, 0)
        OutputMsg += "<td><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(10, 2)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(10, 3)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(10, 4)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td></tr><tr></tr><td> </td><tr>"
        CellExamine = BPDBWS(12, 0)
        OutputMsg += "<td><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(12, 2)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(12, 3)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(12, 4)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td></tr><tr>"
        CellExamine = BPDBWS(13, 0)
        OutputMsg += "<td><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(13, 2)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(13, 3)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(13, 4)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td></tr><tr>"
        CellExamine = BPDBWS(14, 0)
        OutputMsg += "<td><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(14, 2)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(14, 3)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"
        CellExamine = BPDBWS(14, 4)
        OutputMsg += "<td><center><p style='font-family:verdana'>" & CellExamine.DisplayText & "</td>"

        OutputMsg += "</tr></table></body></html>"

        CellExamine = BPDBWS(0, 0)
        GearingBrowserMsg.DocumentText = OutputMsg
        GearingBrowserMsg.Dock = DockStyle.Fill

        TablePanelDashboard.Controls.Add(GearingBrowserMsg)
        TablePanelDashboard.SetCell(GearingBrowserMsg, 0, 1)


    End Sub
    Sub AddStocksChart()

        StockChart = New ChartControl()
        StockChart.Dock = DockStyle.None

        ' Bind the chart to a data source:
        Dim CellExamine As DevExpress.Spreadsheet.Cell
        Dim CellExamineRight As DevExpress.Spreadsheet.Cell
        Dim Datapoints As New List(Of StockDataPoint)
        Dim StockNumWs As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets("FFRW - Stock Numbers")
        Dim DataRange As DevExpress.Spreadsheet.CellRange = StockNumWs.Range("J11:J41,$S$11:$S$41")
        Dim NeDp As StockDataPoint
        For x = 0 To 29

            CellExamine = DataRange(x, 0)

            CellExamineRight = DataRange(x, 0)
            NeDp = New StockDataPoint("Year " & (x + 1).ToString, "Existing", CellExamine.Value.NumericValue)

            Datapoints.Add(NeDp)
        Next
        DataRange = StockNumWs.Range("S11:S41")
        For x = 0 To 29

            CellExamine = DataRange(x, 0)

            CellExamineRight = DataRange(x, 0)
            NeDp = New StockDataPoint("Year " & (x + 1).ToString, "Development", CellExamine.Value.NumericValue)

            Datapoints.Add(NeDp)
        Next
        StockChart.DataSource = Datapoints
        StockChart.SeriesTemplate.ChangeView(ViewType.StackedBar)
        StockChart.SeriesTemplate.SeriesDataMember = "Stocktype"
        StockChart.SeriesTemplate.SetDataMembers("Year", "Numbers")
        ' Enable series point labels, specify their text pattern and position:
        StockChart.SeriesTemplate.LabelsVisibility = DevExpress.Utils.DefaultBoolean.False
        StockChart.SeriesTemplate.Label.TextPattern = "{V}"
        CType(StockChart.SeriesTemplate.Label, BarSeriesLabel).Position = BarSeriesLabelPosition.Center
        ' Customize series view settings (for example, bar width):
        Dim view As StackedBarSeriesView = CType(StockChart.SeriesTemplate.View, StackedBarSeriesView)
        view.BarWidth = 0.8
        ' Disable minor tickmarks on the x-axis:
        Dim diagram As XYDiagram = CType(StockChart.Diagram, XYDiagram)
        diagram.AxisX.Tickmarks.MinorVisible = False
        ' Add a chart title:DevExpress.XtraCharts.
        StockChart.Titles.Add(New DevExpress.XtraCharts.ChartTitle With {.Text = "Opening Stock"})
        ' Specify legend settings:
        StockChart.Legend.MarkerMode = LegendMarkerMode.CheckBoxAndMarker
        StockChart.Legend.AlignmentHorizontal = LegendAlignmentHorizontal.Center
        StockChart.Legend.AlignmentVertical = LegendAlignmentVertical.TopOutside
        StockChart.Dock = DockStyle.Fill
        TablePanelDashboard.Controls.Add(StockChart)
        TablePanelDashboard.SetCell(StockChart, 0, 0)

    End Sub
    Class StockDataPoint
        Public Property Year As String

        Public Property StockType As String

        Public Property Numbers As Integer

        Public Sub New(ByVal SetYear As String, ByVal SetStockType As String, ByVal SetNumbers As Integer)
            Me.Year = SetYear
            Me.StockType = SetStockType
            Me.Numbers = SetNumbers
        End Sub

    End Class
    Sub AddOperatingIncomeChart()

        IncomePieChart = New ChartControl()

        IncomePieChart.Width = Me.Width / 3
        IncomePieChart.Height = Me.Height / 2

        IncomePieChart.Titles.Add(New DevExpress.XtraCharts.ChartTitle() With {.Text = "Yr1 Operating Cash Income"})

        ' Create a pie series.
        Dim series1 As New DevExpress.XtraCharts.Series("Income by category", ViewType.Pie)

        ' Bind the series to data.
        Dim Datapoints As New List(Of DataPoint)
        Dim CellExamine As DevExpress.Spreadsheet.Cell
        Dim CellExamineRight As DevExpress.Spreadsheet.Cell
        Dim DataRange As DevExpress.Spreadsheet.CellRange = CurrChartWS.Range("A53:D59")
        Dim NeDp As DataPoint
        For x = 0 To 6
            NeDp = New DataPoint
            CellExamine = DataRange(x, 0)
            NeDp.Argument = CellExamine.Value.TextValue
            CellExamineRight = DataRange(x, 3)
            If CInt(CellExamineRight.Value.NumericValue) < 0 Then
                NeDp.Argument = NeDp.Argument & " (Loss of £" & CInt(CellExamineRight.Value.NumericValue).ToString & "k)"
                NeDp.Value = 0
            Else
                NeDp.Value = CInt(CellExamineRight.Value.NumericValue)
            End If

            Datapoints.Add(NeDp)
        Next


        series1.DataSource = Datapoints
        series1.ArgumentDataMember = "Argument"
        series1.ValueDataMembers.AddRange(New String() {"Value"})

        ' Add the series to the chart.
        IncomePieChart.Series.Add(series1)

        ' Format the the series labels.
        series1.Label.TextPattern = "{VP:p0} (£{V:n0}k)"

        ' Format the series legend items.
        series1.LegendTextPattern = "{A}"

        ' Adjust the position of series labels. 
        CType(series1.Label, PieSeriesLabel).Position = PieSeriesLabelPosition.TwoColumns

        ' Detect overlapping of series labels.
        CType(series1.Label, PieSeriesLabel).ResolveOverlappingMode = ResolveOverlappingMode.Default

        ' Access the view-type-specific options of the series.
        Dim myView As PieSeriesView = CType(series1.View, PieSeriesView)
        myView.ExplodeMode = DevExpress.XtraCharts.PieExplodeMode.None
        ' Specify a data filter to explode points.
        'myView.ExplodedPointsFilters.Add(New SeriesPointFilter(SeriesPointKey.Value_1, DataFilterCondition.GreaterThanOrEqual, 9))
        'myView.ExplodedPointsFilters.Add(New SeriesPointFilter(SeriesPointKey.Argument, DataFilterCondition.NotEqual, "Others"))

        'myView.ExplodedDistancePercentage = 15
        myView.RuntimeExploding = True

        ' Customize the legend.
        IncomePieChart.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True

        ' Add the chart to the form.
        IncomePieChart.Dock = DockStyle.Fill
        TablePanelDashboard.Controls.Add(IncomePieChart)
        TablePanelDashboard.SetCell(IncomePieChart, 1, 0)
    End Sub
    Sub AddExpensesPieChart()
        ExpensesPieChart = New ChartControl()

        ExpensesPieChart.Width = Me.Width / 3
        ExpensesPieChart.Height = Me.Height / 2

        ExpensesPieChart.Titles.Add(New DevExpress.XtraCharts.ChartTitle() With {.Text = "Yr1 Operating Cash Expenditure"})

        ' Create a pie series.
        Dim series1 As New DevExpress.XtraCharts.Series("Expenses by category", ViewType.Pie)

        ' Bind the series to data.
        Dim Datapoints As New List(Of DataPoint)
        Dim CellExamine As DevExpress.Spreadsheet.Cell
        Dim CellExamineRight As DevExpress.Spreadsheet.Cell
        Dim DataRange As DevExpress.Spreadsheet.CellRange = CurrChartWS.Range("A61:D72")
        Dim NeDp As DataPoint

        For x = 0 To 6
            NeDp = New DataPoint
            CellExamine = DataRange(x, 0)
            NeDp.Argument = CellExamine.Value.TextValue
            CellExamineRight = DataRange(x, 3)
            If CInt(CellExamineRight.Value.NumericValue) < 0 Then
                NeDp.Argument = NeDp.Argument & " (Loss of £" & CInt(CellExamineRight.Value.NumericValue).ToString & "k)"
                NeDp.Value = 0
            Else
                NeDp.Value = CInt(CellExamineRight.Value.NumericValue)
            End If

            Datapoints.Add(NeDp)
        Next


        series1.DataSource = Datapoints
        series1.ArgumentDataMember = "Argument"
        series1.ValueDataMembers.AddRange(New String() {"Value"})

        ' Add the series to the chart.
        ExpensesPieChart.Series.Add(series1)

        ' Format the the series labels.
        series1.Label.TextPattern = "{VP:p0} (£{V:n0}k)"

        ' Format the series legend items.
        series1.LegendTextPattern = "{A}"

        ' Adjust the position of series labels. 
        CType(series1.Label, PieSeriesLabel).Position = PieSeriesLabelPosition.TwoColumns

        ' Detect overlapping of series labels.
        CType(series1.Label, PieSeriesLabel).ResolveOverlappingMode = ResolveOverlappingMode.Default

        ' Access the view-type-specific options of the series.
        Dim myView As PieSeriesView = CType(series1.View, PieSeriesView)

        ' Specify a data filter to explode points.
        'myView.ExplodedPointsFilters.Add(New SeriesPointFilter(SeriesPointKey.Value_1, DataFilterCondition.GreaterThanOrEqual, 9))
        myView.ExplodedPointsFilters.Add(New SeriesPointFilter(SeriesPointKey.Argument, DataFilterCondition.NotEqual, "Others"))
        myView.ExplodeMode = DevExpress.XtraCharts.PieExplodeMode.UseFilters
        myView.ExplodedDistancePercentage = 15
        myView.RuntimeExploding = False
        myView.ExplodeMode = DevExpress.XtraCharts.PieExplodeMode.None
        ' Customize the legend.
        ExpensesPieChart.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True

        ' Add the chart to the form.
        ExpensesPieChart.Dock = DockStyle.Fill
        'ExpensesPieChart.Width = 1200
        'ExpensesPieChart.Height = 1000
        ExpensesPieChart.Top = StockChart.Bottom
        ExpensesPieChart.Left = IncomePieChart.Right
        ExpensesPieChart.Dock = DockStyle.Fill
        TablePanelDashboard.Controls.Add(ExpensesPieChart)
        TablePanelDashboard.SetCell(ExpensesPieChart, 1, 1)

    End Sub
    Public Class DataPoint
        Public Property Argument() As String
        Public Property Value() As Double

    End Class

    Sub AddDebtChart()
        Dim CoveWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets("Covenants")
        DebtLineChart = New ChartControl()
        ' Create a line series.
        Dim series1 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Forecast", ViewType.Line)
        Dim series2 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Target", ViewType.Line)
        Dim Datapoints As New List(Of DataPoint)
        Dim CellExamine As DevExpress.Spreadsheet.Cell

        Dim DataRange As DevExpress.Spreadsheet.CellRange = CoveWS.Range("S11:T50")
        Dim NeDp As DataPoint
        For x = 0 To 39

            NeDp = New DataPoint
            CellExamine = DataRange(x, 0)
            series1.Points.Add(New SeriesPoint("Year " & (x + 1).ToString, CellExamine.Value.NumericValue))
            CellExamine = DataRange(x, 1)
            series2.Points.Add(New SeriesPoint("Year " & (x + 1).ToString, CellExamine.Value.NumericValue))

        Next

        ' Add the series to the chart.
        DebtLineChart.Series.Add(series1)
        DebtLineChart.Series.Add(series2)

        ' Set the numerical argument scale types for the series,
        ' as it is qualitative, by default.
        series1.ArgumentScaleType = ScaleType.Auto
        ' Access the view-type-specific options of the series.

        Dim View1 As LineSeriesView = CType(series1.View, LineSeriesView)
        View1.MarkerVisibility = DevExpress.Utils.DefaultBoolean.False
        View1.LineStyle.Thickness = 4

        Dim View2 As LineSeriesView = CType(series2.View, LineSeriesView)
        View2.MarkerVisibility = DevExpress.Utils.DefaultBoolean.False
        View2.LineStyle.Thickness = 4

        ' Access the view-type-specific options of the series.
        CType(DebtLineChart.Diagram, XYDiagram).AxisY.Interlaced = True
        CType(DebtLineChart.Diagram, XYDiagram).AxisY.InterlacedColor = Color.FromArgb(20, 60, 60, 60)
        CType(DebtLineChart.Diagram, XYDiagram).AxisX.NumericScaleOptions.AutoGrid = False
        CType(DebtLineChart.Diagram, XYDiagram).AxisX.NumericScaleOptions.GridSpacing = 1
        ' Hide the legend (if necessary).
        DebtLineChart.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True
        ' Add a title to the chart (if necessary).
        DebtLineChart.Titles.Add(New DevExpress.XtraCharts.ChartTitle())
        DebtLineChart.Titles(0).Text = "Debt"
        ' Add the chart to the form.


        DebtLineChart.Dock = DockStyle.Fill
        TablePanelDashboard.Controls.Add(DebtLineChart)
        TablePanelDashboard.SetCell(DebtLineChart, 0, 2)

    End Sub
    Sub AddCvntChart()

        Dim CoveWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets("OW - Live Covenant Calculation")
        CvntChart = New ChartControl()
        ' Create a line series.
        Dim series1 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Okay", ViewType.ScatterLine)
        Dim series2 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Breach", ViewType.ScatterLine)
        Dim Datapoints As New List(Of DataPoint)
        Dim CellExamine As DevExpress.Spreadsheet.Cell

        Dim DataRange As DevExpress.Spreadsheet.CellRange = CoveWS.Range("I7:I46")
        Dim NeDp As DataPoint

        For x = 0 To 39
            NeDp = New DataPoint
            CellExamine = DataRange(x, 0)
            If CellExamine.DisplayText = "#N/A" Then
                series1.Points.Add(New SeriesPoint(x + 1, 0.5))
            Else
                series2.Points.Add(New SeriesPoint(x + 1, 0.5))
            End If

        Next

        ' Add the series to the chart.
        CvntChart.Series.Add(series1)
        CvntChart.Series.Add(series2)

        ' Set the numerical argument scale types for the series,
        ' as it is qualitative, by default.
        series1.ArgumentScaleType = ScaleType.Auto
        ' Access the view-type-specific options of the series.
        CType(series1.View, LineSeriesView).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True
        CType(series2.View, LineSeriesView).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True
        CType(series1.View, LineSeriesView).LineMarkerOptions.Size = 20
        CType(series1.View, LineSeriesView).LineMarkerOptions.Color = Color.Green
        CType(series1.View, LineSeriesView).LineMarkerOptions.Kind = MarkerKind.Circle
        CType(series2.View, LineSeriesView).LineMarkerOptions.Size = 20
        CType(series2.View, LineSeriesView).LineMarkerOptions.Color = Color.Red
        CType(series2.View, LineSeriesView).LineMarkerOptions.Kind = MarkerKind.Cross
        'CType(series1.View, LineSeriesView).LineStyle.DashStyle = DevExpress.XtraCharts.DashStyle.Empty
        CType(series1.View, LineSeriesView).LineStyle.Thickness = 1
        ' Access the view-type-specific options of the series.
        CType(CvntChart.Diagram, XYDiagram).AxisY.WholeRange.SetMinMaxValues(0, 1)
        CType(CvntChart.Diagram, XYDiagram).AxisY.Visibility = DevExpress.Utils.DefaultBoolean.False
        CType(CvntChart.Diagram, XYDiagram).AxisX.Visibility = DevExpress.Utils.DefaultBoolean.False
        CType(CvntChart.Diagram, XYDiagram).AxisX.Tickmarks.Visible = False
        CType(CvntChart.Diagram, XYDiagram).AxisY.Tickmarks.Visible = False
        CType(CvntChart.Diagram, XYDiagram).AxisX.GridLines.Visible = False
        CType(CvntChart.Diagram, XYDiagram).AxisY.GridLines.Visible = False
        CType(CvntChart.Diagram, XYDiagram).AxisY.InterlacedColor = Color.FromArgb(20, 60, 60, 60)
        CType(CvntChart.Diagram, XYDiagram).AxisX.NumericScaleOptions.AutoGrid = False
        CType(CvntChart.Diagram, XYDiagram).AxisX.Interlaced = False
        CType(CvntChart.Diagram, XYDiagram).AxisY.Interlaced = False
        CType(CvntChart.Diagram, XYDiagram).AxisX.NumericScaleOptions.GridSpacing = 1

        ' Hide the legend (if necessary).
        CvntChart.Legend.Visibility = DevExpress.Utils.DefaultBoolean.False
        ' Add a title to the chart (if necessary).
        CvntChart.Titles.Add(New DevExpress.XtraCharts.ChartTitle())
        CvntChart.Titles(0).Text = "Gearing Covenant Compliance"
        CvntChart.Titles(0).EnableAntialiasing = DevExpress.Utils.DefaultBoolean.True
        CvntChart.Titles(0).DXFont = New DXFont("Tahoma", 10, DXFontStyle.Bold)

        ' Add the chart to the form.

        CvntChart.Dock = DockStyle.Fill
        TablePanelDashboard.Controls.Add(CvntChart)
        TablePanelDashboard.SetCell(CvntChart, 1, 2)
    End Sub
    Sub Add2ndCovChart()

        Dim CoveWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets("OW - Live Covenant Calculation")
        CvntChart2 = New ChartControl()
        ' Create a line series.
        Dim series1 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Okay", ViewType.ScatterLine)
        Dim series2 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Breach", ViewType.ScatterLine)
        Dim Datapoints As New List(Of DataPoint)
        Dim CellExamine As DevExpress.Spreadsheet.Cell

        Dim DataRange As DevExpress.Spreadsheet.CellRange = CoveWS.Range("M7:M46")
        Dim NeDp As DataPoint

        For x = 0 To 39
            NeDp = New DataPoint
            CellExamine = DataRange(x, 0)
            If CellExamine.DisplayText = "#N/A" Then
                series1.Points.Add(New SeriesPoint(x + 1, 0.5))
            Else
                series2.Points.Add(New SeriesPoint(x + 1, 0.5))
            End If

        Next

        ' Add the series to the chart.
        CvntChart2.Series.Add(series1)
        CvntChart2.Series.Add(series2)
        CType(CvntChart2.Diagram, XYDiagram).AxisY.WholeRange.SetMinMaxValues(0, 1)
        ' Set the numerical argument scale types for the series,
        ' as it is qualitative, by default.
        series1.ArgumentScaleType = ScaleType.Auto
        ' Access the view-type-specific options of the series.
        CType(series1.View, LineSeriesView).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True
        CType(series2.View, LineSeriesView).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True
        CType(series1.View, LineSeriesView).LineMarkerOptions.Size = 20
        CType(series1.View, LineSeriesView).LineMarkerOptions.Color = Color.Green
        CType(series1.View, LineSeriesView).LineMarkerOptions.Kind = MarkerKind.Circle
        CType(series2.View, LineSeriesView).LineMarkerOptions.Size = 20
        CType(series2.View, LineSeriesView).LineMarkerOptions.Color = Color.Red
        CType(series2.View, LineSeriesView).LineMarkerOptions.Kind = MarkerKind.Cross
        'CType(series1.View, LineSeriesView).LineStyle.DashStyle = DevExpress.XtraCharts.DashStyle.Empty
        CType(series1.View, LineSeriesView).LineStyle.Thickness = 1
        ' Access the view-type-specific options of the series.
        CType(CvntChart2.Diagram, XYDiagram).AxisX.Tickmarks.Visible = False
        CType(CvntChart2.Diagram, XYDiagram).AxisY.Tickmarks.Visible = False
        CType(CvntChart2.Diagram, XYDiagram).AxisX.GridLines.Visible = False
        CType(CvntChart2.Diagram, XYDiagram).AxisY.GridLines.Visible = False
        CType(CvntChart2.Diagram, XYDiagram).AxisY.Visibility = DevExpress.Utils.DefaultBoolean.False
        CType(CvntChart2.Diagram, XYDiagram).AxisX.Visibility = DevExpress.Utils.DefaultBoolean.False
        CType(CvntChart2.Diagram, XYDiagram).AxisY.InterlacedColor = Color.FromArgb(20, 60, 60, 60)
        CType(CvntChart2.Diagram, XYDiagram).AxisX.NumericScaleOptions.AutoGrid = False
        CType(CvntChart2.Diagram, XYDiagram).AxisX.NumericScaleOptions.GridSpacing = 1

        ' Add a title to the chart (if necessary).
        CvntChart2.Titles.Add(New DevExpress.XtraCharts.ChartTitle())
        CvntChart2.Titles(0).Text = "Op Margin Covenant Compliance"
        CvntChart2.Titles(0).EnableAntialiasing = DevExpress.Utils.DefaultBoolean.True
        CvntChart2.Titles(0).DXFont = New DXFont("Tahoma", 10, DXFontStyle.Bold)
        CvntChart2.Legend.Visibility = DevExpress.Utils.DefaultBoolean.False
        ' Add the chart to the form.
        'CvntChart2.Top = CvntChart.Bottom
        'CvntChart2.Dock = DockStyle.None
        'CvntChart2.Width = 1200
        'CvntChart2.Height = 187
        'CvntChart2.Left = GearingBrowserMsg.Right
        'Me.Controls.Add(CvntChart2)

    End Sub
    Sub Add3rdCovChart()
        Dim CoveWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets("OW - Live Covenant Calculation")
        CvntChart3 = New ChartControl()
        ' Create a line series.
        Dim series1 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Okay", ViewType.ScatterLine)
        Dim series2 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Breach", ViewType.ScatterLine)
        Dim Datapoints As New List(Of DataPoint)
        Dim CellExamine As DevExpress.Spreadsheet.Cell

        Dim DataRange As DevExpress.Spreadsheet.CellRange = CoveWS.Range("Q7:Q46")
        Dim NeDp As DataPoint

        For x = 0 To 39
            NeDp = New DataPoint
            CellExamine = DataRange(x, 0)
            If CellExamine.DisplayText = "#N/A" Then
                series1.Points.Add(New SeriesPoint(x + 1, 0.5))
            Else
                series2.Points.Add(New SeriesPoint(x + 1, 0.5))
            End If

        Next

        ' Add the series to the chart.
        CvntChart3.Series.Add(series1)
        CvntChart3.Series.Add(series2)
        CType(CvntChart3.Diagram, XYDiagram).AxisY.WholeRange.SetMinMaxValues(0, 1)
        ' Set the numerical argument scale types for the series,
        ' as it is qualitative, by default.
        series1.ArgumentScaleType = ScaleType.Auto
        ' Access the view-type-specific options of the series.
        CType(series1.View, LineSeriesView).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True
        CType(series2.View, LineSeriesView).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True
        CType(series1.View, LineSeriesView).LineMarkerOptions.Size = 20
        CType(series1.View, LineSeriesView).LineMarkerOptions.Color = Color.Green
        CType(series1.View, LineSeriesView).LineMarkerOptions.Kind = MarkerKind.Circle
        CType(series2.View, LineSeriesView).LineMarkerOptions.Size = 20
        CType(series2.View, LineSeriesView).LineMarkerOptions.Color = Color.Red
        CType(series2.View, LineSeriesView).LineMarkerOptions.Kind = MarkerKind.Cross
        'CType(series1.View, LineSeriesView).LineStyle.DashStyle = DevExpress.XtraCharts.DashStyle.Empty
        CType(series1.View, LineSeriesView).LineStyle.Thickness = 1
        ' Access the view-type-specific options of the series.
        CType(CvntChart3.Diagram, XYDiagram).AxisX.Tickmarks.Visible = False
        CType(CvntChart3.Diagram, XYDiagram).AxisY.Tickmarks.Visible = False
        CType(CvntChart3.Diagram, XYDiagram).AxisX.GridLines.Visible = False
        CType(CvntChart3.Diagram, XYDiagram).AxisY.GridLines.Visible = False
        CType(CvntChart3.Diagram, XYDiagram).AxisY.Visibility = DevExpress.Utils.DefaultBoolean.False
        CType(CvntChart3.Diagram, XYDiagram).AxisX.Visibility = DevExpress.Utils.DefaultBoolean.False
        CType(CvntChart3.Diagram, XYDiagram).AxisY.InterlacedColor = Color.FromArgb(20, 60, 60, 60)
        CType(CvntChart3.Diagram, XYDiagram).AxisX.NumericScaleOptions.AutoGrid = False
        CType(CvntChart3.Diagram, XYDiagram).AxisX.NumericScaleOptions.GridSpacing = 1

        ' Add a title to the chart (if necessary).
        CvntChart3.Titles.Add(New DevExpress.XtraCharts.ChartTitle())
        CvntChart3.Titles(0).Text = "EBITDA MRI Covenant Compliance"
        CvntChart3.Titles(0).EnableAntialiasing = DevExpress.Utils.DefaultBoolean.True
        CvntChart3.Titles(0).DXFont = New DXFont("Tahoma", 10, DXFontStyle.Bold)
        CvntChart3.Legend.Visibility = DevExpress.Utils.DefaultBoolean.False
        ' Add the chart to the form.
        'CvntChart3.Top = CvntChart2.Bottom
        'CvntChart3.Dock = DockStyle.None
        'CvntChart3.Width = 1200
        'CvntChart3.Height = 187
        'CvntChart3.Left = GearingBrowserMsg.Right
        'Me.Controls.Add(CvntChart3)
    End Sub
    Sub Add4thCovChart()
        Dim CoveWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets("OW - Live Covenant Calculation")
        CvntChart4 = New ChartControl()
        ' Create a line series.
        Dim series1 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Okay", ViewType.ScatterLine)
        Dim series2 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Breach", ViewType.ScatterLine)
        Dim Datapoints As New List(Of DataPoint)
        Dim CellExamine As DevExpress.Spreadsheet.Cell

        Dim DataRange As DevExpress.Spreadsheet.CellRange = CoveWS.Range("U7:U46")
        Dim NeDp As DataPoint

        For x = 0 To 39
            NeDp = New DataPoint
            CellExamine = DataRange(x, 0)
            If CellExamine.DisplayText = "#N/A" Then
                series1.Points.Add(New SeriesPoint(x + 1, 0.5))
            Else
                series2.Points.Add(New SeriesPoint(x + 1, 0.5))
            End If

        Next

        ' Add the series to the chart.
        CvntChart4.Series.Add(series1)
        CvntChart4.Series.Add(series2)
        CType(CvntChart4.Diagram, XYDiagram).AxisY.WholeRange.SetMinMaxValues(0, 1)
        ' Set the numerical argument scale types for the series,
        ' as it is qualitative, by default.
        series1.ArgumentScaleType = ScaleType.Auto
        ' Access the view-type-specific options of the series.
        CType(series1.View, LineSeriesView).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True
        CType(series2.View, LineSeriesView).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True
        CType(series1.View, LineSeriesView).LineMarkerOptions.Size = 20
        CType(series1.View, LineSeriesView).LineMarkerOptions.Color = Color.Green
        CType(series1.View, LineSeriesView).LineMarkerOptions.Kind = MarkerKind.Circle
        CType(series2.View, LineSeriesView).LineMarkerOptions.Size = 20
        CType(series2.View, LineSeriesView).LineMarkerOptions.Color = Color.Red
        CType(series2.View, LineSeriesView).LineMarkerOptions.Kind = MarkerKind.Cross
        'CType(series1.View, LineSeriesView).LineStyle.DashStyle = DevExpress.XtraCharts.DashStyle.Empty
        CType(series1.View, LineSeriesView).LineStyle.Thickness = 1
        ' Access the view-type-specific options of the series.
        CType(CvntChart4.Diagram, XYDiagram).AxisX.Tickmarks.Visible = False
        CType(CvntChart4.Diagram, XYDiagram).AxisY.Tickmarks.Visible = False
        CType(CvntChart4.Diagram, XYDiagram).AxisX.GridLines.Visible = False
        CType(CvntChart4.Diagram, XYDiagram).AxisY.GridLines.Visible = False
        CType(CvntChart4.Diagram, XYDiagram).AxisY.Visibility = DevExpress.Utils.DefaultBoolean.False
        CType(CvntChart4.Diagram, XYDiagram).AxisX.Visibility = DevExpress.Utils.DefaultBoolean.False
        CType(CvntChart4.Diagram, XYDiagram).AxisY.InterlacedColor = Color.FromArgb(20, 60, 60, 60)
        CType(CvntChart4.Diagram, XYDiagram).AxisX.NumericScaleOptions.AutoGrid = False
        CType(CvntChart4.Diagram, XYDiagram).AxisX.NumericScaleOptions.GridSpacing = 1

        ' Add a title to the chart (if necessary).
        CvntChart4.Titles.Add(New DevExpress.XtraCharts.ChartTitle())
        CvntChart4.Titles(0).Text = "Debt / Unit Covenant Compliance"
        CvntChart4.Titles(0).EnableAntialiasing = DevExpress.Utils.DefaultBoolean.True
        CvntChart4.Titles(0).DXFont = New DXFont("Tahoma", 10, DXFontStyle.Bold)
        CvntChart4.Legend.Visibility = DevExpress.Utils.DefaultBoolean.False
        ' Add the chart to the form.
        'CvntChart4.Top = CvntChart3.Bottom
        'CvntChart4.Dock = DockStyle.None
        'CvntChart4.Width = 1200
        'CvntChart4.Height = 187
        'CvntChart4.Left = GearingBrowserMsg.Right
        'Me.Controls.Add(CvntChart4)
    End Sub
    Sub Add5thCovChart()
        Dim CoveWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets("OW - Live Covenant Calculation")
        CvntChart5 = New ChartControl()
        ' Create a line series.
        Dim series1 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Okay", ViewType.ScatterLine)
        Dim series2 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Breach", ViewType.ScatterLine)
        Dim Datapoints As New List(Of DataPoint)
        Dim CellExamine As DevExpress.Spreadsheet.Cell

        Dim DataRange As DevExpress.Spreadsheet.CellRange = CoveWS.Range("Y7:Y46")
        Dim NeDp As DataPoint

        For x = 0 To 39
            NeDp = New DataPoint
            CellExamine = DataRange(x, 0)
            If CellExamine.DisplayText = "#N/A" Then
                series1.Points.Add(New SeriesPoint(x + 1, 0.5))
            Else
                series2.Points.Add(New SeriesPoint(x + 1, 0.5))
            End If

        Next

        ' Add the series to the chart.
        CvntChart5.Series.Add(series1)
        CvntChart5.Series.Add(series2)
        CType(CvntChart5.Diagram, XYDiagram).AxisY.WholeRange.SetMinMaxValues(0, 1)
        ' Set the numerical argument scale types for the series,
        ' as it is qualitative, by default.
        series1.ArgumentScaleType = ScaleType.Auto
        ' Access the view-type-specific options of the series.
        CType(series1.View, LineSeriesView).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True
        CType(series2.View, LineSeriesView).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True
        CType(series1.View, LineSeriesView).LineMarkerOptions.Size = 20
        CType(series1.View, LineSeriesView).LineMarkerOptions.Color = Color.Green
        CType(series1.View, LineSeriesView).LineMarkerOptions.Kind = MarkerKind.Circle
        CType(series2.View, LineSeriesView).LineMarkerOptions.Size = 20
        CType(series2.View, LineSeriesView).LineMarkerOptions.Color = Color.Red
        CType(series2.View, LineSeriesView).LineMarkerOptions.Kind = MarkerKind.Cross
        'CType(series1.View, LineSeriesView).LineStyle.DashStyle = DevExpress.XtraCharts.DashStyle.Empty
        CType(series1.View, LineSeriesView).LineStyle.Thickness = 1
        ' Access the view-type-specific options of the series.
        CType(CvntChart5.Diagram, XYDiagram).AxisX.Tickmarks.Visible = False
        CType(CvntChart5.Diagram, XYDiagram).AxisY.Tickmarks.Visible = False
        CType(CvntChart5.Diagram, XYDiagram).AxisX.GridLines.Visible = False
        CType(CvntChart5.Diagram, XYDiagram).AxisY.GridLines.Visible = False
        CType(CvntChart5.Diagram, XYDiagram).AxisY.Visibility = DevExpress.Utils.DefaultBoolean.False
        CType(CvntChart5.Diagram, XYDiagram).AxisX.Visibility = DevExpress.Utils.DefaultBoolean.False
        CType(CvntChart5.Diagram, XYDiagram).AxisY.InterlacedColor = Color.FromArgb(20, 60, 60, 60)
        CType(CvntChart5.Diagram, XYDiagram).AxisX.NumericScaleOptions.AutoGrid = False
        CType(CvntChart5.Diagram, XYDiagram).AxisX.NumericScaleOptions.GridSpacing = 1

        ' Add a title to the chart (if necessary).
        CvntChart5.Titles.Add(New DevExpress.XtraCharts.ChartTitle())
        CvntChart5.Titles(0).Text = "Debt Covenant Compliance"
        CvntChart5.Titles(0).EnableAntialiasing = DevExpress.Utils.DefaultBoolean.True
        CvntChart5.Titles(0).DXFont = New DXFont("Tahoma", 10, DXFontStyle.Bold)
        CvntChart5.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True
        CvntChart5.Legend.AlignmentHorizontal = LegendAlignmentHorizontal.Right
        CvntChart5.Legend.AlignmentVertical = LegendAlignmentVertical.BottomOutside
        CvntChart5.Legend.Direction = LegendDirection.LeftToRight
        ' Add the chart to the form.
        'CvntChart5.Top = CvntChart4.Bottom
        'CvntChart5.Dock = DockStyle.None
        'CvntChart5.Width = 1200
        'CvntChart5.Height = 223
        'CvntChart5.Left = GearingBrowserMsg.Right
        'Me.Controls.Add(CvntChart5)
    End Sub
    Private Sub PopulateDashboard()

        If DataPres.Sections.Length < 0 Then Exit Sub

        AcControl = New AccordionControl() With {
            .Dock = DockStyle.Fill,
            .Parent = Me,
            .Width = Me.Width
            }

        Formatter.FormatAccordianControl(AcControl)

        AcControl.BeginUpdate()

        Dim Section As PresentationSection
        Dim ActiveDataSet As DataCellRange
        Dim hyperlinkLabelControl1 As New HyperlinkLabelControl()
        AcElementlist = New List(Of AccordionControlElement)


        For Each Section In DataPres.Sections


            DataSourceCount += 1
            AcElementCount += 1

            ReDim Preserve AcElements(AcElementCount)

            AcElements(AcElementCount) = AcControl.AddItem

            With AcElements(AcElementCount)

                .Text = Section.Name
                '.Style = ElementStyle.Group
                .Name = "Element" & AcElementCount.ToString
                .Expanded = True

            End With




            AcContainersCount += 1

            ReDim Preserve AcContainers(AcContainersCount)
            AcContainers(AcContainersCount) = New AccordionContentContainer()

            AcControl.Controls.Add(AcContainers(AcContainersCount))
            AcElements(AcElementCount).ContentContainer = AcContainers(AcContainersCount)

            Formatter.FormatAccordianControlContainer(AcContainers(AcContainersCount))

            For Each SectionElement In Section.SectionElements


                If SectionElement.Type = "Grid" Then

                    ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)


                    ColList = New List(Of String)

                    PropertiesCount = -1 'reset

                    ReDim Preserve DataSources(DataSourceCount)
                    Dim SetTag As New AbovoUnboundSourceTag With {.GSID = GSID, .CSID = CSID, .DSIndex = SectionElement.ControlSourceIndex}


                    DataSources(DataSourceCount) = New AbovoUnboundSource(DataSourceCount, SetTag)

                    For Each PresentedColumn In ActiveDataSet.DataColumns

                        ColCount += 1
                        ColName = "Col_" & ColCount.ToString

                        PropertiesCount += 1
                        ReDim Preserve PropertyArray(PropertiesCount)

                        Select Case PresentedColumn.ColumnTag.DataType
                            Case "S"
                                PropType = GetType(String)
                            Case "I", "Y"
                                PropType = GetType(Integer)
                            Case "N", "P"
                                PropType = GetType(Double)
                            Case "B"
                                PropType = GetType(Integer)
                            Case Else
                                PropType = GetType(String)
                        End Select
                        PropertyArray(PropertiesCount) = New UnboundSourceProperty With {
                            .UserTag = ColCount,
                            .Name = ColName,
                            .PropertyType = PropType,
                            .DisplayName = PresentedColumn.ColumnTag.ColumnHeading
                        }

                        ColList.Add(ColName)

                    Next

                    PropertyList = PropertyArray

                    DataSources(DataSourceCount).Properties.AddRange(PropertyList)

                    AddHandler DataSources(DataSourceCount).ValueNeeded, AddressOf UnboundDS_ValueNeeded
                    AddHandler DataSources(DataSourceCount).ValuePushed, AddressOf UnboundDS_ValuePushed

                    GridCount += 1
                    ReDim Preserve GridControls(GridCount)

                    GridControls(GridCount) = New GridControl() With {
                        .Name = "GridControl_" & GridCount.ToString,
                        .Parent = Me,
                        .Dock = DockStyle.Fill,
                        .DataSource = DataSources(DataSourceCount)
                    }

                    GridViewCount += 1
                    ReDim Preserve UsedGridViews(GridViewCount)

                    UsedGridViews(GridViewCount) = New DevExpress.XtraGrid.Views.Grid.GridView
                    GridControls(GridCount).ViewCollection.Add(UsedGridViews(GridViewCount))
                    GridControls(GridCount).MainView = UsedGridViews(GridViewCount)
                    UsedGridViews(GridViewCount).PopulateColumns()


                    Dim testUBS As AbovoUnboundSource = TryCast(GridControls(GridCount).DataSource, AbovoUnboundSource)

                    DataSources(DataSourceCount).SetRowCount(ActiveDataSet.RowCount)

                    Formatter.FormatGridControl(GridControls(GridCount))

                    AcContainers(AcContainersCount).Controls.Add(GridControls(GridCount))

                    AcContainers(AcContainersCount).Height = 900
                    AcContainers(AcContainersCount).Width = Me.Width
                    AcContainers(AcContainersCount).Appearance.BackColor = Color.Aquamarine
                    GridControls(GridCount).Dock = DockStyle.Fill
                    GridControls(GridCount).ForceInitialize()


                    Formatter.FormatGridView(UsedGridViews(GridViewCount), GridControls(GridCount))
                    UsedGridViews(GridViewCount).BestFitColumns()

                    UsedGridViews(GridViewCount).OptionsView.BestFitMaxRowCount = ActiveDataSet.RowCount

                Else



                End If

            Next







        Next

        AcControl.ExpandElementMode = ExpandElementMode.Multiple
        AcControl.ExpandAll()
        AcControl.EndUpdate()



    End Sub
    Private Sub InitAccordionControl()
        AcControl.BeginUpdate()
        Dim acRootGroupHome As New AccordionControlElement()
        Dim acItemActivity As New AccordionControlElement()
        Dim acItemNews As New AccordionControlElement()
        Dim acRootItemSettings As New AccordionControlElement()

        AddHandler AcControl.ElementClick, AddressOf AccordionControl1_ElementClick

        ' 
        ' Root Group 'Home'
        ' 
        acRootGroupHome.Elements.AddRange(New AccordionControlElement() {acItemActivity, acItemNews})
        acRootGroupHome.Expanded = True
        acRootGroupHome.ImageOptions.ImageUri.Uri = "Home;Office2013"
        acRootGroupHome.Name = "acRootGroupHome"
        acRootGroupHome.Text = "Home"
        ' 
        ' Child Item 'Activity'
        ' 
        acItemActivity.Name = "acItemActivity"
        acItemActivity.Style = ElementStyle.Item
        acItemActivity.Tag = "idActivity"
        acItemActivity.Text = "Activity"
        ' 
        ' Child Item 'News'
        ' 
        acItemNews.Name = "acItemNews"
        acItemNews.Style = ElementStyle.Item
        acItemNews.Tag = "idNews"
        acItemNews.Text = "News"
        ' 
        ' Root Item 'Settings' with ContentContainer
        ' 
        acRootItemSettings.ImageOptions.ImageUri.Uri = "Customization;Office2013"
        acRootItemSettings.Name = "acRootItemSettings"
        acRootItemSettings.Style = ElementStyle.Item
        acRootItemSettings.Text = "Settings"
        ' 
        ' itemSettingsControlContainer
        ' 
        Dim itemSettingsControlContainer As New AccordionContentContainer()
        Dim hyperlinkLabelControl1 As New HyperlinkLabelControl()
        Dim toggleSwitch1 As New ToggleSwitch()
        AcControl.Controls.Add(itemSettingsControlContainer)
        acRootItemSettings.ContentContainer = itemSettingsControlContainer
        itemSettingsControlContainer.Controls.Add(hyperlinkLabelControl1)
        itemSettingsControlContainer.Controls.Add(toggleSwitch1)
        itemSettingsControlContainer.Appearance.BackColor = System.Drawing.SystemColors.Control
        itemSettingsControlContainer.Appearance.Options.UseBackColor = True
        itemSettingsControlContainer.Height = 60
        ' 
        ' hyperlinkLabelControl1
        ' 
        hyperlinkLabelControl1.Location = New System.Drawing.Point(26, 33)
        hyperlinkLabelControl1.Size = New System.Drawing.Size(107, 13)
        hyperlinkLabelControl1.Text = "www.devexpress.com"
        AddHandler hyperlinkLabelControl1.HyperlinkClick, AddressOf HyperlinkLabelControl1_HyperlinkClick
        ' 
        ' toggleSwitch1
        ' 
        toggleSwitch1.EditValue = True
        toggleSwitch1.Location = New System.Drawing.Point(24, 3)
        toggleSwitch1.Properties.AllowFocused = False
        toggleSwitch1.Properties.AutoWidth = True
        toggleSwitch1.Properties.OffText = "Offline Mode"
        toggleSwitch1.Properties.OnText = "Onlne Mode"
        toggleSwitch1.Size = New System.Drawing.Size(134, 24)
        AddHandler toggleSwitch1.Toggled, AddressOf toggleSwitch1_Toggled

        AcControl.Elements.AddRange(New DevExpress.XtraBars.Navigation.AccordionControlElement() {acRootGroupHome, acRootItemSettings})

        acRootItemSettings.Expanded = True

        AcControl.EndUpdate()
    End Sub

    Private Sub AccordionControl1_ElementClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.Navigation.ElementClickEventArgs)
        If e.Element.Style = DevExpress.XtraBars.Navigation.ElementStyle.Group Then
            Return
        End If
        If e.Element.Tag Is Nothing Then
            Return
        End If
        Dim itemID As String = e.Element.Tag.ToString()
        If itemID = "idNews" Then
            '...
        End If
        'listBoxControl1.Items.Add(itemID & " clicked")
    End Sub

    Private Sub toggleSwitch1_Toggled(ByVal sender As Object, ByVal e As EventArgs)
        '...
    End Sub

    Private Sub HyperlinkLabelControl1_HyperlinkClick(ByVal sender As Object, ByVal e As DevExpress.Utils.HyperlinkClickEventArgs)
        Process.Start(e.Text)
    End Sub
    Private Sub UnboundDS_ValueNeeded(ByVal sender As Object, ByVal e As DevExpress.Data.UnboundSourceValueNeededEventArgs)

        Dim UDSSender As AbovoUnboundSource = sender
        e.Value = GetDSData(UDSSender.UBSTag.DSIndex, e.RowIndex, e.PropertyIndex)
    End Sub
    Private Sub UnboundDS_ValuePushed(ByVal sender As Object, ByVal e As DevExpress.Data.UnboundSourceValuePushedEventArgs)
        'something = e.Value ' Propagate the value into the storage.
    End Sub
    Private Function GetDSData(ByVal SetDSIndex As Integer, ByVal rowIndex As Integer, ByVal PropertyIndex As Integer) As Object


        Dim DP As CellDataPoint = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(PropertyIndex)


        Select Case DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag.DataType

            Case "S"

                Return DP.StringValue

            Case "B"

                Return DP.BoolValue

            Case "N", "P"

                Return DP.RealValue

            Case "I", "Y"

                Return DP.IntValue

            Case Else
                Return Nothing

        End Select

    End Function

    'Public Shared Sub CustomDrawColumnHeader(ByVal gridControl As GridControl, ByVal gridView As GridView)
    '    ' Handle this event to paint columns headers manually
    '    AddHandler gridView.CustomDrawColumnHeader, Sub(s, e)
    '                                                    If e.Column Is Nothing OrElse e.Column.FieldName <> "Category_Name" Then
    '                                                        Return
    '                                                    End If
    '                                                    ' Fill column headers with the specified colors.
    '                                                    e.Cache.FillRectangle(Color.Coral, e.Bounds)
    '                                                    e.Appearance.DrawString(e.Cache, e.Info.Caption, e.Info.CaptionRect)
    '                                                    ' Draw the filter and sort buttons.
    '                                                    For Each info As DrawElementInfo In e.Info.InnerElements
    '                                                        If Not info.Visible Then
    '                                                            Continue For
    '                                                        End If
    '                                                        ObjectPainter.DrawObject(e.Cache, info.ElementPainter, info.ElementInfo)
    '                                                    Next info
    '                                                    e.Handled = True
    '                                                End Sub
    'End Sub
    'Private Sub gridView1_CustomDrawColumnHeader(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.ColumnHeaderCustomDrawEventArgs)
    '    e.Graphics.FillRectangle(Brushes.Green, e.Bounds)
    '    Using pen As New Pen(Color.Black, 3)
    '        e.Graphics.DrawRectangle(pen, e.Bounds)
    '    End Using
    '    e.Info.InnerElements.DrawObjects(e.Info, e.Cache, Point.Empty)
    '    e.Handled = True
    'End Sub
#Region "Interface events"

    Sub ResizeMe()
        Dim ScaleFactor As Single = Me.Width / 1920

    End Sub
    Private Sub WindowsUIButtonPanelItemEdit_ButtonChecked(sender As Object, e As ButtonEventArgs)
        Dim tag As String = CType(e.Button, WindowsUIButton).Tag.ToString()
        Select Case tag
            Case "AddDispo"


                'GridViewStockNumbers.Columns(11).Visible = False
                'GridViewStockNumbers.Columns(9).Visible = False
                'GridViewStockNumbers.Columns(8).Visible = False
                'GridViewStockNumbers.Columns(7).Visible = False
                'GridViewStockNumbers.Columns(6).Visible = False
                'GridViewStockNumbers.Columns(5).Visible = False
                'Grid1ExpandedView = False
                'SortGridColums()
                'CustomDrawCell(GridControlStockGrid, GridViewStockNumbers)
                'GridViewStockNumbers.LeftCoord = 0
        End Select
    End Sub
    Private Sub WindowsUIButtonPanelItemEdit_ButtonCheck(sender As Object, e As DevExpress.XtraBars.Docking2010.ButtonEventArgs)
        Dim tag As String = CType(e.Button, WindowsUIButton).Tag.ToString()
        Select Case tag
            Case "AddDispo"

                'GridViewStockNumbers.Columns(10).Visible = False

                'GridViewStockNumbers.Columns(5).Visible = True

                'GridViewStockNumbers.Columns(6).Visible = True

                'GridViewStockNumbers.Columns(7).Visible = True

                'GridViewStockNumbers.Columns(8).Visible = True

                'GridViewStockNumbers.Columns(9).Visible = True

                'GridViewStockNumbers.Columns(10).Visible = True

                'GridViewStockNumbers.Columns(11).Visible = True

                'Grid1ExpandedView = True
                'SortGridColums()
                'CustomDrawCell(GridControlStockGrid, GridViewStockNumbers)
                'Dim Column As DevExpress.XtraGrid.Columns.GridColumn = GridViewStockNumbers.Columns(5)
                'Dim info As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridViewInfo = GridViewStockNumbers.GetViewInfo()
                'GridViewStockNumbers.LeftCoord = info.GetColumnLeftCoord(Column) + GridViewStockNumbers.Columns(0).Width
        End Select
    End Sub

    Private Sub WindowsUIButtonPanelBPActions_ButtonClick(sender As Object, e As DevExpress.XtraBars.Docking2010.ButtonEventArgs)
        Dim tag As String = CType(e.Button, WindowsUIButton).Tag.ToString()
        Select Case tag
            Case "ApplyAndSave"
                ' OpenAssumptionsInterface
                WriteStockToBPAndSave()
            Case "ApplyToFile"
                ' Navigate to page B 
                WriteStockToBP()
            Case "Ad3"
                    ' Navigate to page C
            Case "Ad4"
                    ' Navigate to page D 
            Case "Ad5"
                ' Navigate to page E 
        End Select
    End Sub
    Private Sub RepositoryItemComboBoxSOCIRent_QueryPopUp(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim lookUpEdit As LookUpEdit = TryCast(sender, LookUpEdit)
        lookUpEdit.Properties.PopulateColumns()
        lookUpEdit.Properties.Columns(0).Visible = False

    End Sub
    Private Sub RepositoryItemComboBoxSOCIStocktype_QueryPopUp(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim lookUpEdit As LookUpEdit = TryCast(sender, LookUpEdit)
        lookUpEdit.Properties.PopulateColumns()
        lookUpEdit.Properties.Columns(0).Visible = False
        lookUpEdit.Properties.Columns(1).Visible = False
    End Sub
#End Region
#Region "Data events"

    Private Sub WriteStockToBP()



    End Sub
    Private Sub WriteStockToBPAndSave()



    End Sub




    Private Sub SetArrayData(ByVal rowIndex As Integer, ByVal propertyName As String, ByVal value As Object)

        Select Case propertyName

            Case "PropertyStockDescription"
                AbovoBP.Stock.StockItems(rowIndex).StockDescription = value
            Case "PropertyOwnedManaged"
                AbovoBP.Stock.StockItems(rowIndex).OwnedManaged = value
            Case "PropertySOCIStockType"
                AbovoBP.Stock.StockItems(rowIndex).SOCIStockType = value
            Case "PropertySOCIRentType"
                AbovoBP.Stock.StockItems(rowIndex).SOCIRentType = value
            Case "PropertyCurrentStockNumbers"
                AbovoBP.Stock.StockItems(rowIndex).CurrentStockNumbers = value
            Case "PropertyInitialRateNewLettings"
                AbovoBP.Stock.StockItems(rowIndex).NewLetInitialRate = value
            Case "PropertyPreBPlanStartDateNewBuild"
                AbovoBP.Stock.StockItems(rowIndex).PreBPlanStartDateNewBuild = value
            Case "PropertyPreBPlanStartDateDemolitions"
                AbovoBP.Stock.StockItems(rowIndex).PreBPlanStartDateDemolitions = value
            Case "PropertyPreBPlanStartDateRTBs"
                AbovoBP.Stock.StockItems(rowIndex).PreBPlanStartDateRTBs = value
            Case "PropertyPreBPlanStartDateOtherDisposals"
                AbovoBP.Stock.StockItems(rowIndex).PreBPlanStartDateOtherDisposals = value
            Case "PropertyNewLettings"
                AbovoBP.Stock.StockItems(rowIndex).NewLettings = value
            Case Else

        End Select

        AbovoBP.Stock.StockItems(rowIndex).FUpdateStockTotals()

    End Sub
#End Region



    Private Sub GridViewStockNumbers_ValidatingEditor(sender As Object, e As DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventArgs)

        Dim view As ColumnView = sender
        Dim column As GridColumn = If(TryCast(e, EditFormValidateEditorEventArgs)?.Column, view.FocusedColumn)



        If column.Name = "colPropertyInitialRateNewLettings1" Then

            If (Convert.ToDecimal(e.Value) < 0) Or (Convert.ToDecimal(e.Value) > 1) Then
                MsgBox("Sorry, The value of initial New Lettings Rate must be more than 0 and less than 100", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If

            Exit Sub

        ElseIf column.Name = "colPropertyCurrentStockNumbers1" Then

            If Convert.ToInt32(e.Value) < 0 Then
                MsgBox("Sorry, the value of Current Stock Numbers must be more than 0", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If
            Exit Sub
        ElseIf column.Name = "colPropertyPreBPlanStartDateNewBuild1" Then

            If Convert.ToInt32(e.Value) < 0 Then
                MsgBox("Sorry, the value of New Build Numbers must be more than 0", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If
            Exit Sub
        ElseIf column.Name = "colPropertyPreBPlanStartDateDemolitions1" Then

            If Convert.ToInt32(e.Value) < 0 Then
                MsgBox("Sorry, the value of Demolitions must be more than 0", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If
            Exit Sub
        ElseIf column.Name = "colPropertyPreBPlanStartDateRTBs1" Then

            If Convert.ToInt32(e.Value) < 0 Then
                MsgBox("Sorry, the value of Right To Buys must be more than 0", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If
            Exit Sub
        ElseIf column.Name = "colPropertyPreBPlanStartDateOtherDisposals1" Then

            If Convert.ToInt32(e.Value) < 0 Then
                MsgBox("Sorry, the value of Other Disposals must be more than 0", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If
            Exit Sub
        ElseIf column.Name = "colPropertyNewLettings1" Then

            If Convert.ToInt32(e.Value) < 0 Then
                MsgBox("Sorry, the value of New Lettings must be more than 0", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If
            Exit Sub
        End If

    End Sub

    Private Sub GridViewStockNumbers_InvalidValueException(sender As Object, e As InvalidValueExceptionEventArgs)
        Dim view As ColumnView = sender

        view.HideEditor()
        Exit Sub
        'Dim view As ColumnView9 = sender
        'If view Is Nothing Then
        '    Return
        'End If
        'Dim column As GridColumn = If(TryCast(e, InvalidValueExceptionEventArgs)?.Column, view.FocusedColumn)
        'e.ExceptionMode = ExceptionMode.DisplayError
        'e.WindowCaption = "Input Error"
        'e.ErrorText = "The value should be greater than 0 and less than 100"
        '' Destroy the editor and discard the changes made within the edited cell
        'view.HideEditor()
    End Sub
    Sub SortGridColums()

        Dim np As Integer = 0
        'For i = 0 To GridViewStockNumbers.Columns.Count - 1
        '    If GridViewStockNumbers.Columns(i).Visible Then
        '        GridViewStockNumbers.Columns(i).VisibleIndex = np
        '        np += 1
        '    End If


        'Next i
    End Sub


    Private Sub RepositoryItemLookUpEditSOCIStockType_EditValueChanged(sender As Object, e As EventArgs)
        'Dim Editor As LookUpEdit = CType(sender, LookUpEdit)

        'Dim StrChosenStock As String = Convert.ToString(Editor.EditValue)
        'Dim IntFoundCatID As Integer = SOCIStock.GetSOCICategoryByName(StrChosenStock)
        'If IntFoundCatID = 0 Then GridViewStockNumbers.SetFocusedRowCellValue("PropertySOCIRentType", Convert.ToString("N/A"))
    End Sub

    Private Sub GridViewStockNumbers_CustomDrawColumnHeader(sender As Object, e As ColumnHeaderCustomDrawEventArgs)
        e.Cache.FillRectangle(Color.White, e.Bounds)
        e.Appearance.DrawString(e.Cache, e.Info.Caption, e.Info.CaptionRect)
        Using pen As New Pen(Color.Silver, 4)
            e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom, e.Bounds.Right + 4, e.Bounds.Bottom)
            'e.Graphics.DrawRectangle(pen, e.Bounds)
        End Using
        ' Draw the filter and sort buttons.
        For Each info As DrawElementInfo In e.Info.InnerElements
            If Not info.Visible Then
                Continue For
            End If
            ObjectPainter.DrawObject(e.Cache, info.ElementPainter, info.ElementInfo)
        Next info
        e.Handled = True
    End Sub



    Sub CustomDrawCell(ByVal gridControl As GridControl, ByVal gridView As GridView)
        ' Handle this event to paint cells manually
        Dim BDo As Boolean = False
        AddHandler gridView.CustomDrawCell,
            Sub(s, e)
                If Grid1ExpandedView Then
                    If e.Column.VisibleIndex = 0 Then
                        Using pen As New Pen(Color.Silver, 4)
                            e.Graphics.DrawLine(pen, e.Bounds.Right, e.Bounds.Top - 4, e.Bounds.Right, e.Bounds.Bottom + 15)
                            'e.Graphics.DrawRectangle(pen, e.Bounds)
                        End Using

                    End If
                    If e.Column.VisibleIndex = 12 Then
                        Using pen As New Pen(Color.Silver, 4)
                            e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Top - 4, e.Bounds.Left, e.Bounds.Bottom + 15)
                            'e.Graphics.DrawRectangle(pen, e.Bounds)
                        End Using

                    End If
                Else
                    If e.Column.VisibleIndex = 0 Then
                        Using pen As New Pen(Color.White, 4)
                            e.Graphics.DrawLine(pen, e.Bounds.Right, e.Bounds.Top, e.Bounds.Right, e.Bounds.Bottom)
                            'e.Graphics.DrawRectangle(pen, e.Bounds)
                        End Using

                    End If
                    If e.Column.VisibleIndex = 12 Then
                        Using pen As New Pen(Color.White, 4)
                            e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Top - 4, e.Bounds.Left, e.Bounds.Bottom + 15)
                            'e.Graphics.DrawRectangle(pen, e.Bounds)
                        End Using

                    End If
                End If
                If BDo Then
                    For Each info As DrawElementInfo In e.Cell.InnerElements
                        If Not info.Visible Then
                            Continue For
                        End If
                        ObjectPainter.DrawObject(e.Cache, info.ElementPainter, info.ElementInfo)
                    Next info
                    e.Handled = True
                End If

            End Sub
    End Sub



    Sub ResizeFonts()

        Scalefactor = Me.Width / 1700



        'Me.GridControlStockGrid.Font = GetFont("Small", Me.Scalefactor)

        'Me.GridViewStockNumbers.Appearance.OddRow.Font = GetFont("Small", Me.Scalefactor)
        'Me.GridViewStockNumbers.Appearance.ViewCaption.Font = GetFont("Small", Me.Scalefactor)

        'Dim x As Integer

        'For x = 0 To GridViewStockNumbers.Columns.Count - 1

        '    GridViewStockNumbers.Columns(x).AppearanceCell.Font = GetFont("Small", Me.Scalefactor)
        '    GridViewStockNumbers.Columns(x).AppearanceHeader.Font = GetFont("Small", Me.Scalefactor, True)

        'Next

        'RepositoryItemComboBoxOwnedManaged.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        'RepositoryItemLookUpEditSOCIStockType.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        'RepositoryItemLookUpEditSOCIRentType.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        'RepositoryItemIntegerEdit.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        'RepositoryItemComboBoxSOCIStockType.Appearance.Font = GetFont("Small", Me.Scalefactor, True)


        'colPropertyStockDescription1.Width = GridControlStockGrid.Width * 0.25
        'colPropertyOwnedManaged1.Width = GridControlStockGrid.Width * 0.125
        'colPropertySOCIStockType1.Width = GridControlStockGrid.Width * 0.125
        'colPropertySOCIRentType1.Width = GridControlStockGrid.Width * 0.125
        'colPropertyCurrentStockNumbers1.Width = GridControlStockGrid.Width * 0.12
        'colPropertyNewLettings1.Width = GridControlStockGrid.Width * 0.12
        'colPropertyTotalOpeningStockCalc1.Width = GridControlStockGrid.Width * 0.12


        'colPropertyPreBPlanStartDateNewBuild1.Width = GridControlStockGrid.Width * 0.1
        'colPropertyPreBPlanStartDateDemolitions1.Width = GridControlStockGrid.Width * 0.1
        'colPropertyPreBPlanStartDateRTBs1.Width = GridControlStockGrid.Width * 0.1
        'colPropertyPreBPlanStartDateOtherDisposals1.Width = GridControlStockGrid.Width * 0.1
        'colPropertyExistingStocksCalc1.Width = GridControlStockGrid.Width * 0.1

        'colPropertyNewLettings1.Width = GridControlStockGrid.Width * 0.1


        'Me.hideContainerRight.Font = GetFont("Small", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.ActiveTab.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDo                 cking.HidePanelButton.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.HidePanelButtonActive.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaption.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaptionActive.Font = GetFont("Medium", Me.ScaleFactor)


        ''XtraTabControlMainNavigator.Appearance.FontSizeDelta = MediumFontSize - XtraTabControlMainNavigator.Appearance.Font.Size
        'Me.BarAndDockingControllerMainScreen.AppearancesBar.Dock.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.HidePanelButton.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaption.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.XtraTabControlMainNavigator.AppearancePage.HeaderActive.Font = GetFont("Medium", Me.ScaleFactor, False, True)
        'Me.XtraTabControlMainNavigator.AppearancePage.Header.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.XtraTabControlMainNavigator.Appearance.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.XtraTabControlMainNavigator.AppearancePage.HeaderHotTracked.Font = GetFont("Medium", Me.ScaleFactor)

        'WindowsUIButtonPanelExitHelp.Font = GetFont("Small", Me.ScaleFactor)
        'Me.WindowsUIButtonPanelBPActions.Font = GetFont("Small", Me.ScaleFactor)

        'WindowsUIButtonPanelOpenCompare.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelOpenCompare.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelOpenCompare.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)

        'WindowsUIButtonPanelBPActions.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelBPActions.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelBPActions.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)

        'WindowsUIButtonPanelExitHelp.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelExitHelp.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelExitHelp.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)

        'WindowsUIButtonPanelSaveClose.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelSaveClose.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelSaveClose.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)
        'Me.GroupBoxFileActions.Font = GetFont("Small", Me.ScaleFactor)
        'Me.WindowsUIButtonPanelBPActions.ButtonBackgroundImages

    End Sub
    Sub ResizeControls()

        Dim SetWidth As Integer = Me.Width * 0.17
        ScaleUnits = Me.Width * 0.007

        'PictureBoxAbovoLogo.Top = ScaleUnits
        'PictureBoxAbovoLogo.Left = ScaleUnits
        'PictureBoxAbovoLogo.Width = SetWidth
        'PictureBoxAbovoLogo.Height = CInt(PictureBoxAbovoLogo.Width * 0.483)

        'DockPanelSettings.Width = SetWidth
        'WindowsUIButtonPanelExitHelp.Left = ScaleUnits
        'GroupBoxProgramDetails.Width = SetWidth
        'XtraTabControlMainNavigator.Top = ScaleUnits
        'XtraTabControlMainNavigator.Left = PictureBoxAbovoLogo.Right + ScaleUnits
        'XtraTabControlMainNavigator.Width = Me.Width - SetWidth - (5 * ScaleUnits) - hideContainerRight.Width
        'XtraTabControlMainNavigator.Height = Me.Height - (6 * ScaleUnits)
        'XtraTabPageMainHABP.Height = XtraTabControlMainNavigator.PageClientBounds.Height
        'WindowsUIButtonPanelOpenCompare.Left = ScaleUnits
        'WindowsUIButtonPanelOpenCompare.Top = 3 * ScaleUnits
        'GroupBoxFileActions.Left = ScaleUnits
        'GroupBoxFileActions.Top = WindowsUIButtonPanelOpenCompare.Bottom + ScaleUnits
        'GroupBoxFileActions.Width = XtraTabControlMainNavigator.Width - (2 * ScaleUnits)
        'GroupBoxFileActions.Height = XtraTabPageMainHABP.Height - WindowsUIButtonPanelExitHelp.Height - (4 * ScaleUnits)
        'WindowsUIButtonPanelExitHelp.Top = XtraTabControlMainNavigator.Bottom - WindowsUIButtonPanelExitHelp.Height
        'WindowsUIButtonPanelExitHelp.Width = SetWidth
        'WindowsUIButtonPanelExitHelp.Left = ScaleUnits
        'WebBrowserBPInfo.Top = (2 * ScaleUnits)
        'WebBrowserBPInfo.Width = GroupBoxFileActions.Width - WindowsUIButtonPanelSaveClose.Width - (3 * ScaleUnits)
        'WebBrowserBPInfo.Height = GroupBoxFileActions.Height - WindowsUIButtonPanelBPActions.Height - (4 * ScaleUnits)
        'WindowsUIButtonPanelSaveClose.Left = WebBrowserBPInfo.Right + ScaleUnits
        'WindowsUIButtonPanelSaveClose.Top = WebBrowserBPInfo.Top
        'WindowsUIButtonPanelSaveClose.Height = GroupBoxFileActions.Height
        'WindowsUIButtonPanelOpenCompare.Width = XtraTabControlMainNavigator.PageClientBounds.Width - (2 * ScaleUnits)
        'WindowsUIButtonPanelBPActions.Top = WebBrowserBPInfo.Bottom + ScaleUnits
        'WindowsUIButtonPanelBPActions.Width = WebBrowserBPInfo.Width
        'GroupBoxProgramDetails.Top = PictureBoxAbovoLogo.Bottom + ScaleUnits
        'GroupBoxProgramDetails.Left = ScaleUnits
        'GroupBoxProgramDetails.Height = WindowsUIButtonPanelExitHelp.Top - PictureBoxAbovoLogo.Bottom - (2 * ScaleUnits)
        'SetBrowserText()

    End Sub
    Private Sub StockAssumptionsInterface_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        ResizeFonts()
        ResizeControls()
    End Sub

End Class





