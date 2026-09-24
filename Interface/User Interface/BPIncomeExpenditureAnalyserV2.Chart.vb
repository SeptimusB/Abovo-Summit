Option Infer On

Imports System.Globalization
Imports System.Drawing
Imports System.Windows.Forms
Imports Abovo.CustomGrid
Imports DevExpress.Utils
Imports DevExpress.Utils.Menu
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.XtraCharts
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Columns

Partial Public Class BPIncomeExpenditureAnalyserV2
    Private ReadOnly ChartViews As New List(Of StatementChartView)
    Private ChartViewButton As WindowsUIButton
    Private ReadOnly ChartGridIcon As DevExpress.Utils.Svg.SvgImage =
        My.Resources.bo_appointment

    Private Sub InitialiseChartViews()
        ChartViews.Add(New StatementChartView(Me, WrapCG_SOCI, XtraTabPageSOCIWrapped))
        ChartViews.Add(New StatementChartView(Me, WrapCG_CF, XtraTabPageCFWrapped))
        ChartViews.Add(New StatementChartView(Me, WrapCG_BS, XtraTabPageBSWrapped))
        ChartViewButton = New WindowsUIButton("", False,
            New WindowsUIButtonImageOptions With {.SvgImage = My.Resources.charttype_spline},
            ButtonStyle.PushButton, "Show line chart", -1, True, Nothing, True, False, True,
            "ToggleChart", -1, True)
        WindowsUIButtonPanelAnalyser.Buttons.Add(New WindowsUISeparator())
        WindowsUIButtonPanelAnalyser.Buttons.Add(ChartViewButton)
        AddHandler XtraTabControlAnalyser.SelectedPageChanged, Sub(s, e) UpdateChartViewButton()
        UpdateChartViewButton()
        Abovo.PresentationLayout.ApplyButtonPanel(WindowsUIButtonPanelAnalyser, Me)
    End Sub

    Private Sub ToggleCurrentStatementChart()
        Dim current = ChartViews.FirstOrDefault(Function(v) v.IsForPage(XtraTabControlAnalyser.SelectedTabPage))
        If current IsNot Nothing Then current.SetShowingChart(Not current.ShowingChart)
    End Sub

    Private Sub UpdateChartViewButton()
        If ChartViewButton Is Nothing Then Return
        Dim current = ChartViews.FirstOrDefault(Function(v) v.IsForPage(XtraTabControlAnalyser.SelectedTabPage))
        ChartViewButton.Enabled = current IsNot Nothing
        Dim showing As Boolean = current IsNot Nothing AndAlso current.ShowingChart
        ChartViewButton.ImageOptions.SvgImage = If(showing, ChartGridIcon, My.Resources.charttype_spline)
        ChartViewButton.ToolTip = If(showing, "Show figures grid", "Show line chart")
    End Sub

    Private Sub InvalidateChartViews(disconnected As Boolean)
        For Each chartView In ChartViews
            chartView.SourceChanged(disconnected)
        Next
    End Sub

    'A read-only projection of the existing view, not a second workbook binding.
    'Group handles are transient: retain value paths across datasource replacement.
    Private NotInheritable Class StatementChartView
        Private ReadOnly Owner As BPIncomeExpenditureAnalyserV2
        Private ReadOnly Wrapper As CustomGridWrapper
        Private ReadOnly View As CustomGridView
        Private ReadOnly Page As DevExpress.XtraTab.XtraTabPage
        Private ReadOnly Chart As New ChartControl With {.Dock = DockStyle.Fill, .RuntimeHitTesting = True}
        Private ReadOnly ChartHost As New Panel With {.Dock = DockStyle.Fill, .Visible = False}
        Private ReadOnly Toolbar As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill, .AutoSize = True, .WrapContents = True,
            .MinimumSize = New Size(0, 32), .Visible = False,
            .Padding = New Padding(3), .BackColor = Color.White}
        Private ReadOnly NavigationMenu As New DXPopupMenu
        Public Property ShowingChart As Boolean
            Get
                Return IsChartShown
            End Get
            Private Set(value As Boolean)
                IsChartShown = value
            End Set
        End Property
        Private IsChartShown As Boolean
        Private ReadOnly UpButton As New SimpleButton With {.Text = "Up", .AutoSize = True}
        Private ReadOnly OverviewButton As New SimpleButton With {.Text = "Overview", .AutoSize = True}
        Private ReadOnly ResetZoomButton As New SimpleButton With {.Text = "Reset zoom", .AutoSize = True}
        Private ReadOnly Totals As New CheckEdit With {.Text = "Include statement totals", .AutoSizeInLayoutControl = True}
        Private ReadOnly Choices As New ComboBoxEdit With {.Width = 270}
        Private ReadOnly DrillButton As New SimpleButton With {.Text = "Drill down", .AutoSize = True}
        Private ReadOnly Trail As New LabelControl With {.AutoSizeMode = LabelAutoSizeMode.None, .Dock = DockStyle.Top}
        Private ReadOnly Status As New LabelControl With {.AutoSizeMode = LabelAutoSizeMode.None, .Dock = DockStyle.Bottom}
        Private ReadOnly Navigation As New List(Of String)
        Private ReadOnly Nodes As New List(Of ChartNode)
        Private Dirty As Boolean = True
        Private Updating As Boolean
        Private Connected As Boolean = True
        Private ReadOnly Property IsBalanceSheet As Boolean
            Get
                Return Object.ReferenceEquals(Wrapper, Owner.WrapCG_BS)
            End Get
        End Property
        Private BalanceDocument As Abovo.BalanceSheetDocument

        Private NotInheritable Class ChartNode
            Public RowHandle As Integer
            Public Caption As String
            Public Path As String
            Public CanDrill As Boolean
            Public BalanceNode As Abovo.BalanceSheetNode
            Public Overrides Function ToString() As String
                Return Caption
            End Function
        End Class

        Public Sub New(analyser As BPIncomeExpenditureAnalyserV2, grid As CustomGridWrapper,
                       tab As DevExpress.XtraTab.XtraTabPage)
            Owner = analyser
            Wrapper = grid
            View = grid.WrappedGridView
            Page = tab
            toolbar.Controls.Add(UpButton)
            toolbar.Controls.Add(OverviewButton)
            toolbar.Controls.Add(ResetZoomButton)
            toolbar.Controls.Add(Totals)
            toolbar.Controls.Add(Choices)
            toolbar.Controls.Add(DrillButton)
            Choices.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
            Totals.Width = 180
            Trail.Height = Math.Max(30, Page.Font.Height * 2 + 6)
            Trail.Appearance.TextOptions.WordWrap = WordWrap.Wrap
            Trail.Padding = New Padding(6)
            Status.Height = Math.Max(30, Page.Font.Height * 2 + 6)
            Status.Appearance.TextOptions.WordWrap = WordWrap.Wrap
            Status.Padding = New Padding(6)
            Chart.BorderOptions.Visibility = DefaultBoolean.False
            Chart.Font = Page.Font
            Chart.Legend.Font = Page.Font
            Chart.Legend.AlignmentHorizontal = LegendAlignmentHorizontal.RightOutside
            Chart.Legend.AlignmentVertical = LegendAlignmentVertical.Top
            Chart.Legend.Direction = LegendDirection.TopToBottom
            Chart.Legend.MaxHorizontalPercentage = 35
            Chart.Legend.MaxVerticalPercentage = 90
            Chart.Legend.Visibility = DefaultBoolean.True
            ChartHost.Controls.Add(Chart)
            ChartHost.Controls.Add(Trail)
            ChartHost.Controls.Add(Status)
            Dim layout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2, .Margin = New Padding(0)}
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            Dim body As New Panel With {.Dock = DockStyle.Fill, .Margin = New Padding(0)}
            body.Controls.Add(Wrapper)
            body.Controls.Add(ChartHost)
            layout.Controls.Add(toolbar, 0, 0)
            layout.Controls.Add(body, 0, 1)
            Page.Controls.Add(layout)
            AddHandler UpButton.Click, Sub(s, e) GoUp()
            AddHandler ResetZoomButton.Click, Sub(s, e)
                                                 Dim diagram = TryCast(Chart.Diagram, XYDiagram)
                                                 If diagram IsNot Nothing Then diagram.ResetZoom()
                                             End Sub
            AddHandler OverviewButton.Click, Sub(s, e)
                                                      Navigation.Clear()
                                                      Dirty = True
                                                      Render()
                                                  End Sub
            AddHandler Totals.CheckedChanged, Sub(s, e)
                                                  Dirty = True
                                                  Render()
                                              End Sub
            AddHandler DrillButton.Click, Sub(s, e) DrillInto(TryCast(Choices.SelectedItem, ChartNode))
            AddHandler Chart.MouseClick, AddressOf ChartClicked
            AddHandler Chart.KeyDown, Sub(s, e)
                                          If e.KeyCode = Keys.Back Then
                                              GoUp()
                                              e.Handled = True
                                          End If
                                      End Sub
            AddHandler Page.VisibleChanged, Sub(s, e)
                                               If Not Page.Visible Then NavigationMenu.HidePopup()
                                               Render()
                                           End Sub
            AddHandler Chart.Disposed, Sub(s, e) NavigationMenu.HidePopup()
            AddHandler View.ColumnFilterChanged, Sub(s, e)
                                                    If View.GridControl.DataSource IsNot Nothing Then SourceChanged(False)
                                                End Sub
            UpdateCommands()
        End Sub

        Public Sub SourceChanged(disconnected As Boolean)
            NavigationMenu.HidePopup()
            Connected = Not disconnected
            Dirty = True
            If disconnected Then
                'Never leave a plausible-looking stale chart during structural deferral.
                Chart.Series.Clear()
                Nodes.Clear()
                Choices.Properties.Items.Clear()
                Status.Text = "Analysis is disconnected. Refresh analysis before using these figures."
                UpdateCommands()
            Else
                Render()
            End If
        End Sub

        Public Function IsForPage(tab As DevExpress.XtraTab.XtraTabPage) As Boolean
            Return Object.ReferenceEquals(Page, tab)
        End Function

        Public Sub SetShowingChart(show As Boolean)
            ShowingChart = show
            NavigationMenu.HidePopup()
            Wrapper.Visible = Not show
            ChartHost.Visible = show
            Toolbar.Visible = show
            UpdateCommands()
            Render()
            Owner.UpdateChartViewButton()
        End Sub

        Private Sub UpdateCommands()
            Dim showing As Boolean = ShowingChart
            For Each command As Control In {UpButton, OverviewButton, ResetZoomButton, Totals, Choices, DrillButton}
                command.Visible = showing
                command.Enabled = Connected
            Next
            UpButton.Enabled = Connected AndAlso Navigation.Count > 0
            OverviewButton.Enabled = Connected AndAlso Navigation.Count > 0
            DrillButton.Enabled = Connected AndAlso Choices.Properties.Items.Count > 0
            Totals.Enabled = Connected AndAlso Navigation.Count = 0
        End Sub

        Private Function CanDrill(rowHandle As Integer) As Boolean
            If Not View.IsGroupRow(rowHandle) Then Return False
            Dim value As Integer
            If View.GetRowLevel(rowHandle) = 0 Then
                If TryGetGroupSummaryInteger(View, rowHandle, 0, value) AndAlso value > 0 Then Return False
            Else
                If TryGetGroupSummaryInteger(View, rowHandle, 1, value) AndAlso value = 1 Then Return False
            End If
            Return View.GetChildRowCount(rowHandle) > 0
        End Function

        Private Function MakeNode(rowHandle As Integer) As ChartNode
            If View.IsGroupRow(rowHandle) Then
                Dim level As Integer = View.GetRowLevel(rowHandle)
                Dim caption As String = Convert.ToString(View.GetGroupRowValue(rowHandle), CultureInfo.CurrentCulture)
                If level < View.GroupedColumns.Count AndAlso
                   IsOrderingField(View.GroupedColumns(level).FieldName) AndAlso caption.Length >= 5 Then
                    caption = caption.Substring(5)
                End If
                Return New ChartNode With {.RowHandle = rowHandle,
                    .Caption = If(String.IsNullOrWhiteSpace(caption), "(Unclassified)", caption.Trim()),
                    .Path = GetGroupPath(View, rowHandle), .CanDrill = CanDrill(rowHandle)}
            End If
            Return New ChartNode With {.RowHandle = rowHandle,
                .Caption = View.GetRowCellDisplayText(rowHandle, "ItemDesc"), .CanDrill = False}
        End Function

        Private Function ResolvePath(path As String) As Integer
            Dim handle As Integer = -1
            While View.IsValidRowHandle(handle)
                If View.IsGroupRow(handle) AndAlso GetGroupPath(View, handle) = path Then Return handle
                handle -= 1
            End While
            Return GridControl.InvalidRowHandle
        End Function

        Private Sub PopulateNodes()
            Nodes.Clear()
            If IsBalanceSheet Then
                If BalanceDocument Is Nothing Then Return
                While Navigation.Count > 0 AndAlso Not BalanceDocument.Nodes.Any(Function(n) n.Id = Navigation.Last() AndAlso BalanceDocument.Children(n.Id).Count > 0)
                    Navigation.RemoveAt(Navigation.Count - 1)
                End While
                For Each node In BalanceDocument.Children(If(Navigation.Count = 0, "", Navigation.Last()))
                    If Navigation.Count = 0 AndAlso node.IsTotal AndAlso Not Totals.Checked Then Continue For
                    Nodes.Add(New ChartNode With {.Caption = node.Caption, .Path = node.Id, .BalanceNode = node,
                              .CanDrill = node.Diagnostic.Length = 0 AndAlso BalanceDocument.Children(node.Id).Count > 0})
                Next
                Return
            End If
            'A structural change can remove the selected branch. Fall back to its
            'nearest surviving, expandable ancestor, never an unrelated row handle.
            While Navigation.Count > 0
                Dim handle As Integer = ResolvePath(Navigation.Last())
                If View.IsValidRowHandle(handle) AndAlso CanDrill(handle) Then Exit While
                Navigation.RemoveAt(Navigation.Count - 1)
            End While
            If Navigation.Count = 0 Then
                Dim handle As Integer = -1
                While View.IsValidRowHandle(handle)
                    If View.IsGroupRow(handle) AndAlso View.GetRowLevel(handle) = 0 Then
                        Dim titleLevel As Integer
                        If Totals.Checked OrElse Not TryGetGroupSummaryInteger(View, handle, 0, titleLevel) OrElse titleLevel = 0 Then
                            Nodes.Add(MakeNode(handle))
                        End If
                    End If
                    handle -= 1
                End While
            Else
                Dim parentHandle As Integer = ResolvePath(Navigation.Last())
                For index As Integer = 0 To View.GetChildRowCount(parentHandle) - 1
                    Dim handle As Integer = View.GetChildRowHandle(parentHandle, index)
                    If View.IsValidRowHandle(handle) Then Nodes.Add(MakeNode(handle))
                Next
            End If
        End Sub

        Private Sub Render()
            If Updating OrElse Not Dirty OrElse Not ShowingChart OrElse Not Page.Visible OrElse
               Not Connected OrElse Owner.AmInactiveState OrElse Owner.IsDisposed OrElse
               (Not IsBalanceSheet AndAlso View.GridControl.DataSource Is Nothing) Then Return
            Updating = True
            Dim watch As Abovo.SummitDiagnostics.DiagnosticTimer = Abovo.SummitDiagnostics.DiagnosticTimer.StartNew()
            Try
                Dim periods As List(Of GridColumn) = Owner.GetPeriodColumns(View)
                If IsBalanceSheet Then
                    BalanceDocument = Owner.BalanceSheetView.EnsureDocument()
                    If BalanceDocument Is Nothing Then Throw New InvalidOperationException("See the Balance Sheet figures panel for the source diagnostic.")
                    periods = BalanceDocument.Periods.Select(Function(caption) New GridColumn With {.Caption = caption}).ToList()
                End If
                Dim summaries As New Dictionary(Of String, GridSummaryItem)(StringComparer.Ordinal)
                For Each item As GridSummaryItem In View.GroupSummary
                    If item.SummaryType = DevExpress.Data.SummaryItemType.Sum Then summaries(item.FieldName) = item
                Next
                PopulateNodes()
                Chart.Series.Clear()
                Choices.Properties.Items.Clear()
                Dim missing As Integer = 0
                For Each node In Nodes
                    Dim series As New Series(node.Caption, ViewType.Line) With {
                        .Tag = node, .ArgumentScaleType = ScaleType.Qualitative,
                        .LabelsVisibility = DefaultBoolean.False,
                        .CrosshairLabelPattern = "{S}: {V:n2}"}
                    Dim line As LineSeriesView = DirectCast(series.View, LineSeriesView)
                    line.LineStyle.Thickness = 5
                    line.MarkerVisibility = DefaultBoolean.True
                    line.LineMarkerOptions.Size = 8
                    Dim values As Hashtable = If(Not IsBalanceSheet AndAlso View.IsGroupRow(node.RowHandle), View.GetGroupSummaryValues(node.RowHandle), Nothing)
                    Dim periodIndex As Integer = 0
                    For Each period In periods
                        Dim raw As Object = Nothing
                        If IsBalanceSheet Then
                            raw = node.BalanceNode.Values(periodIndex)
                        ElseIf values IsNot Nothing Then
                            Dim summary As GridSummaryItem = Nothing
                            If summaries.TryGetValue(period.FieldName, summary) Then raw = values(summary)
                        Else
                            raw = View.GetRowCellValue(node.RowHandle, period)
                        End If
                        Dim amount As Double
                        Dim valid As Boolean = TryChartNumber(raw, amount)
                        If Not valid Then amount = 0
                        Dim point As New SeriesPoint(period.Caption.Replace(vbLf, " ").Replace(vbCr, ""), amount)
                        point.IsEmpty = Not valid
                        If Not valid Then missing += 1
                        series.Points.Add(point)
                        periodIndex += 1
                    Next
                    Chart.Series.Add(series)
                    If node.CanDrill Then Choices.Properties.Items.Add(node)
                Next
                If Choices.Properties.Items.Count > 0 Then Choices.SelectedIndex = 0
                Dim diagram As XYDiagram = TryCast(Chart.Diagram, XYDiagram)
                If diagram IsNot Nothing Then
                    diagram.AxisX.Label.Font = Page.Font
                    diagram.AxisY.Label.Font = Page.Font
                    diagram.AxisX.Title.Font = Page.Font
                    diagram.AxisY.Title.Font = Page.Font
                    diagram.AxisX.Title.Text = "Forecast period"
                    diagram.AxisX.Title.Visibility = DefaultBoolean.True
                    diagram.AxisX.GridLines.Visible = False
                    diagram.AxisY.Title.Text = "Value (same units as figures)"
                    diagram.AxisY.Title.Visibility = DefaultBoolean.True
                    diagram.AxisY.Label.TextPattern = "{V:n0}"
                    diagram.AxisY.WholeRange.AlwaysShowZeroLevel = True
                    diagram.EnableAxisXZooming = True
                    diagram.EnableAxisXScrolling = True
                    diagram.EnableAxisYZooming = True
                    diagram.EnableAxisYScrolling = True
                    diagram.ZoomingOptions.UseMouseWheel = True
                    diagram.ZoomingOptions.UseKeyboard = True
                    diagram.ZoomingOptions.UseKeyboardWithMouse = False
                    diagram.ScrollingOptions.UseMouse = True
                    diagram.ScrollingOptions.ScrollMouseAction.MouseButton = DevExpress.Portable.Input.PortableMouseButtons.Left
                    diagram.ScrollingOptions.ScrollMouseAction.ModifierKeys = ChartModifierKeys.Control
                End If
                Dim labels As New List(Of String) From {Page.Text, If(Owner.CurrentDataSourceMode = AnalyserDataSourceMode.Comparison, "Differences", Owner.CurrentDataSourceMode.ToString()), "Overview"}
                For Each path In Navigation
                    labels.Add(If(IsBalanceSheet, BalanceDocument.Nodes.First(Function(n) n.Id = path).Caption, MakeNode(ResolvePath(path)).Caption))
                Next
                Trail.Text = String.Join("  >  ", labels)
                Trail.ToolTip = Trail.Text
                Status.Text = If(Nodes.Count = 0, "No records match this statement/filter.",
                    "Wheel: zoom. Ctrl+drag: pan. Right-click: Drill down / Go Up; use the list for overlapping lines.")
                If missing > 0 Then Status.Text &= " " & missing.ToString() & " unavailable values shown as gaps."
                If IsBalanceSheet Then Status.Text &= " Balances include openings; transaction contributions are cumulative."
                If Nodes.Count > 0 AndAlso Nodes.All(Function(node) node.Caption = "(Unclassified)") Then
                    Status.Text = "This level has no group labels in the workbook/grid. Drill down if detail is available; no classifications have been invented."
                End If
                Dirty = False
                UpdateCommands()
                Abovo.SummitDiagnostics.WriteLine("[Analyser Chart] statement=" & Page.Text &
                    ", mode=" & Owner.CurrentDataSourceMode.ToString() & ", depth=" & Navigation.Count.ToString() &
                    ", series=" & Nodes.Count.ToString() & ", total=" & watch.ElapsedMilliseconds.ToString() & " ms")
            Catch ex As Exception
                Chart.Series.Clear()
                Nodes.Clear()
                Choices.Properties.Items.Clear()
                Status.Text = "Chart unavailable. Figures remain available. " & ex.Message
                UpdateCommands()
                Abovo.SummitDiagnostics.WriteLine("[Analyser Chart] " & ex.ToString())
            Finally
                Updating = False
            End Try
        End Sub

        Private Shared Function TryChartNumber(value As Object, ByRef number As Double) As Boolean
            number = 0
            If value Is Nothing OrElse value Is DBNull.Value Then Return False
            Return Double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture),
                NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.InvariantCulture, number) AndAlso
                Not Double.IsNaN(number) AndAlso Not Double.IsInfinity(number)
        End Function

        Private Sub ChartClicked(sender As Object, e As MouseEventArgs)
            If e.Button <> MouseButtons.Right OrElse Not Connected Then Return
            Dim hit As ChartHitInfo = Chart.CalcHitInfo(e.Location)
            Dim series As Series = TryCast(hit.Series, Series)
            BuildNavigationMenu(If(series Is Nothing, Nothing, TryCast(series.Tag, ChartNode)))
            NavigationMenu.ShowPopup(Chart, e.Location)
        End Sub

        Private Sub BuildNavigationMenu(node As ChartNode)
            NavigationMenu.Items.Clear()
            NavigationMenu.Items.Add(New DXMenuItem("Go Up", Sub(s, e) GoUp()) With {
                .Enabled = Connected AndAlso Navigation.Count > 0})
            'Do not select an arbitrary neighbouring series on empty chart space.
            Dim caption As String = If(node Is Nothing, "Drill down to (select a series)", "Drill down to " & node.Caption)
            NavigationMenu.Items.Add(New DXMenuItem(caption, Sub(s, e) DrillInto(node)) With {
                .Enabled = Connected AndAlso node IsNot Nothing AndAlso node.CanDrill})
        End Sub

        Private Sub DrillInto(node As ChartNode)
            If node Is Nothing OrElse Not Connected Then Return
            If Not node.CanDrill Then
                Status.Text = "No further detail for " & node.Caption & ". Use Up or Overview to return."
                Return
            End If
            Navigation.Add(node.Path)
            Dirty = True
            Render()
        End Sub

        Private Sub GoUp()
            If Not Connected OrElse Navigation.Count = 0 Then Return
            Navigation.RemoveAt(Navigation.Count - 1)
            Dirty = True
            Render()
        End Sub
    End Class
End Class
