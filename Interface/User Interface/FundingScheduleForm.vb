Imports System.Drawing
Imports System.Windows.Forms
Imports Abovo
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Grid

Public Class FundingScheduleForm
    Inherits XtraForm

    Private ReadOnly ModelID As Integer
    Private ReadOnly SectionChoice As New CheckedListBoxControl With {.Height = 135, .CheckOnClick = True}
    Private ReadOnly FacilityChoice As New ComboBoxEdit
    Private ReadOnly ColourChoice As New ColorEdit With {.Color = Color.CornflowerBlue}
    Private ReadOnly PopulateChoice As New CheckEdit With {.Text = "Populate target cells with a fixed figure"}
    Private ReadOnly FigureChoice As New SpinEdit
    Private ReadOnly FigureUnits As New LabelControl With {.AutoSizeMode = LabelAutoSizeMode.Vertical, .Dock = DockStyle.Fill}
    Private ReadOnly PreviewTimer As New Timer With {.Interval = 300}
    Private Revision As Integer
    Private ReadOnly StartChoice As New DateEdit
    Private ReadOnly IntervalChoice As New ComboBoxEdit
    Private ReadOnly EndDateChoice As New DateEdit
    Private ReadOnly EndChoice As New ComboBoxEdit
    Private ReadOnly CountChoice As New SpinEdit
    Private ReadOnly ApplyButton As New SimpleButton With {.Text = "OK — add schedule", .Enabled = False}
    Private ReadOnly Status As New LabelControl With {.AutoSizeMode = LabelAutoSizeMode.Vertical, .Dock = DockStyle.Fill}
    Private ReadOnly Grid As New GridControl With {.Dock = DockStyle.Fill}
    Private ReadOnly View As New GridView
    Private ReadOnly Settings As New TableLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True, .ColumnCount = 2, .Padding = New Padding(0, 0, 18, 0)}
    Private PreviewDates As List(Of FundingScheduleDate)
    Private Busy As Boolean

    Public ReadOnly Property SelectedTarget As FundingScheduleTarget
        Get
            Return SelectedTargets.FirstOrDefault()
        End Get
    End Property
    Public ReadOnly Property SelectedTargets As List(Of FundingScheduleTarget)
        Get
            Return SectionChoice.Items.Cast(Of DevExpress.XtraEditors.Controls.CheckedListBoxItem)().
                Where(Function(item) item.CheckState = CheckState.Checked).
                Select(Function(item) DirectCast(item.Value, FundingScheduleTarget)).ToList()
        End Get
    End Property
    Public ReadOnly Property SelectedFacility As FundingScheduleFacility
        Get
            Return TryCast(FacilityChoice.SelectedItem, FundingScheduleFacility)
        End Get
    End Property
    Public ReadOnly Property ScheduleColour As Color
        Get
            Return ColourChoice.Color
        End Get
    End Property
    Public ReadOnly Property FixedFigure As Decimal?
        Get
            Return If(PopulateChoice.Checked, CType(FigureChoice.Value, Decimal?), Nothing)
        End Get
    End Property
    Public ReadOnly Property ConfigurationSummary As String
        Get
            Return "Start=" & StartChoice.DateTime.ToString("yyyy-MM-dd") & "; Interval=" & IntervalChoice.Text & "; Rule=Same day as start date" &
                If(EndChoice.SelectedIndex = 0, "; End=" & EndDateChoice.DateTime.ToString("yyyy-MM-dd"), "; Occurrences=" & CountChoice.Value.ToString()) &
                "; FixedFigure=" & If(FixedFigure.HasValue, FixedFigure.Value.ToString(Globalization.CultureInfo.InvariantCulture), "blank") &
                "; FigureUnits=" & FigureUnits.Text
        End Get
    End Property
    Public ReadOnly Property DatesToApply As List(Of DateTime)
        Get
            Return PreviewDates.Select(Function(d) d.DateToAdd).ToList()
        End Get
    End Property

    Public Sub New(model As Integer, targets As IEnumerable(Of FundingScheduleTarget), initial As String, Optional facilityColumn As Integer = -1)
        ModelID = model
        Text = "Funding — create date schedule"
        Font = Abovo.FontManager.DefaultFont
        ShowIcon = False
        StartPosition = FormStartPosition.CenterParent
        MinimumSize = New Size(850, 650)
        ClientSize = New Size(1040, 850)
        ShowInTaskbar = False
        MinimizeBox = False
        Dim layout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .Padding = New Padding(20), .ColumnCount = 2, .RowCount = 4}
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 55))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Dim heading As New LabelControl With {.Text = "Create a facility schedule" & vbCrLf &
            "Dates unlock the amount cells beside them. Existing entries are preserved; duplicate dates remain separate records.",
            .AutoSizeMode = LabelAutoSizeMode.Vertical, .Dock = DockStyle.Fill, .Margin = New Padding(0, 0, 0, 18)}
        heading.Appearance.ForeColor = Abovo.GeneralFunctions.AbovoBlue
        layout.Controls.Add(heading, 0, 0) : layout.SetColumnSpan(heading, 2)
        Settings.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 38))
        Settings.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 62))
        Dim choices = targets.ToList()
        SectionChoice.Height = If(choices.Count = 1, 34, 120)
        For Each target In choices
            SectionChoice.Items.Add(target, target.DateRange = initial OrElse (String.IsNullOrEmpty(initial) AndAlso target Is choices.First()))
        Next
        For Each facility In FundingScheduleGroups.Facilities(FileManager.ExcelModels(ModelID).WB, FundingScheduleGroups.IsInvestment(choices.First()))
            FacilityChoice.Properties.Items.Add(facility)
        Next
        FacilityChoice.SelectedIndex = If(FacilityChoice.Properties.Items.Count > 0, 0, -1)
        For i = 0 To FacilityChoice.Properties.Items.Count - 1
            If DirectCast(FacilityChoice.Properties.Items(i), FundingScheduleFacility).ColumnIndex = facilityColumn Then FacilityChoice.SelectedIndex = i
        Next
        StartChoice.DateTime = Date.Today
        IntervalChoice.Properties.Items.AddRange(New String() {"Monthly", "Quarterly", "Semi-annually", "Annually"})
        For Each combo In {FacilityChoice, IntervalChoice}
            combo.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
            If combo IsNot FacilityChoice Then combo.SelectedIndex = 0
        Next
        EndDateChoice.DateTime = Date.Today.AddYears(1)
        EndChoice.Properties.Items.AddRange(New String() {"End date", "Number of occurrences"})
        EndChoice.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        EndChoice.SelectedIndex = 0
        CountChoice.Properties.IsFloatValue = False
        CountChoice.Properties.MinValue = 1
        CountChoice.Properties.MaxValue = 1000
        CountChoice.Properties.Mask.EditMask = "d"
        CountChoice.Value = 12
        CountChoice.ToolTip = "Number of dates, including the start date (1 to 1,000)."
        For Each picker In {StartChoice, EndDateChoice}
            picker.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
            picker.Properties.DisplayFormat.FormatString = "dd MMM yyyy"
            picker.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime
            picker.Properties.EditFormat.FormatString = "dd/MM/yyyy"
            picker.Properties.Mask.EditMask = "dd/MM/yyyy"
            picker.Properties.MinValue = New DateTime(1900, 1, 1)
        Next
        AddSetting("Facility / loan", FacilityChoice)
        AddSetting("Sections", SectionChoice)
        AddSetting("Schedule colour", ColourChoice)
        AddSetting("Target values", PopulateChoice)
        FigureChoice.Properties.IsFloatValue = True
        FigureChoice.Properties.MinValue = Decimal.MinValue
        FigureChoice.Properties.MaxValue = Decimal.MaxValue
        'Numeric masks limit integer input to the number of digit placeholders.
        FigureChoice.Properties.Mask.EditMask = "############################0.##########"
        AddSetting("Fixed figure", FigureChoice)
        AddSetting("Units", FigureUnits)
        AddSetting("Start date", StartChoice)
        AddSetting("Interval", IntervalChoice)
        AddSetting("Schedule ends by", EndChoice)
        AddSetting("End date (inclusive)", EndDateChoice)
        AddSetting("Number of occurrences", CountChoice)
        StartChoice.ToolTip = "Repeat this day at each interval. A shorter month uses its last valid day; later dates return to the original day. Weekends and holidays are not adjusted."
        ColourChoice.ToolTip = "A very subtle Summit-only tint. Excel fills and editing rules are unchanged."
        Dim scroll As New XtraScrollableControl With {.Dock = DockStyle.Fill}
        scroll.Controls.Add(Settings)
        layout.Controls.Add(scroll, 0, 1)
        Grid.MainView = View : Grid.ViewCollection.Add(View)
        Dim formatter As New ObjectFormatter
        formatter.FormatGridControl(Grid)
        formatter.FormatGridView(View, Grid)
        View.OptionsBehavior.Editable = False
        View.OptionsSelection.MultiSelect = True
        View.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect
        View.OptionsClipboard.CopyColumnHeaders = DevExpress.Utils.DefaultBoolean.True
        View.OptionsView.ShowGroupPanel = False
        View.OptionsView.ShowIndicator = False
        View.OptionsView.RowAutoHeight = True
        View.OptionsView.ColumnAutoWidth = True
        View.OptionsView.ShowHorizontalLines = DevExpress.Utils.DefaultBoolean.False
        View.OptionsView.ShowVerticalLines = DevExpress.Utils.DefaultBoolean.False
        View.OptionsBehavior.AutoPopulateColumns = False
        Dim dateColumn = View.Columns.AddVisible("DateToAdd", "Date")
        dateColumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        dateColumn.DisplayFormat.FormatString = "ddd dd MMM yyyy"
        View.Appearance.HeaderPanel.ForeColor = Abovo.GeneralFunctions.AbovoBlue
        layout.Controls.Add(Grid, 1, 1)
        Status.Margin = New Padding(0, 12, 0, 12)
        layout.Controls.Add(Status, 0, 2) : layout.SetColumnSpan(Status, 2)
        Dim buttons As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .AutoSize = True, .FlowDirection = FlowDirection.RightToLeft}
        Dim cancel As New SimpleButton With {.Text = "Cancel", .DialogResult = DialogResult.Cancel}
        For Each button In {cancel, ApplyButton}
            button.AutoSize = True : button.MinimumSize = New Size(125, 34) : buttons.Controls.Add(button)
        Next
        layout.Controls.Add(buttons, 0, 3) : layout.SetColumnSpan(buttons, 2)
        Controls.Add(layout)
        CancelButton = cancel
        AcceptButton = ApplyButton
        AddHandler ApplyButton.Click, Sub()
                                          If PreviewDates IsNot Nothing AndAlso ApplyButton.Enabled Then DialogResult = DialogResult.OK
                                      End Sub
        AddHandler PreviewTimer.Tick, Async Sub()
                                               PreviewTimer.Stop()
                                               Await RefreshPreviewAsync()
                                           End Sub
        AddHandler Shown, Async Sub() Await RefreshPreviewAsync()
        AddHandler FormClosed, Sub() PreviewTimer.Stop()
        AddHandler Disposed, Sub() PreviewTimer.Dispose()
        AddHandler SectionChoice.ItemCheck, Sub() InvalidatePreview()
        For Each editor As BaseEdit In {FacilityChoice, ColourChoice, PopulateChoice, StartChoice, IntervalChoice, EndDateChoice, EndChoice, CountChoice}
            AddHandler editor.EditValueChanged, Sub() InvalidatePreview()
        Next
        'Do not rebuild the preview or read SpinEdit.Value on every digit typed.
        InvalidatePreview()
    End Sub

    Private Sub AddSetting(caption As String, control As Control)
        Dim row = Settings.RowCount : Settings.RowCount += 1
        Settings.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Settings.Controls.Add(New LabelControl With {.Text = caption, .Dock = DockStyle.Fill, .AutoSizeMode = LabelAutoSizeMode.Vertical, .Margin = New Padding(0, 7, 8, 7)}, 0, row)
        control.Dock = DockStyle.Fill : control.Margin = New Padding(0, 5, 0, 5)
        Settings.Controls.Add(control, 1, row)
    End Sub

    Private Sub InvalidatePreview()
        Revision += 1
        PreviewTimer.Stop()
        ApplyButton.Enabled = False
        PreviewDates = Nothing : Grid.DataSource = Nothing
        FigureChoice.Enabled = PopulateChoice.Checked
        EndDateChoice.Enabled = EndChoice.SelectedIndex = 0
        CountChoice.Enabled = EndChoice.SelectedIndex = 1
        Status.Text = "Updating preview… No workbook changes are made until OK."
        If Not IsDisposed Then PreviewTimer.Start()
    End Sub

    Private Function RefreshPreviewAsync() As Threading.Tasks.Task
        RefreshPreview()
        Return Threading.Tasks.Task.CompletedTask
    End Function

    Private Sub RefreshPreview()
        If Busy Then Return
        PreviewTimer.Stop()
        Dim currentRevision = Revision
        Busy = True : ApplyButton.Enabled = False
        Try
            If SelectedFacility Is Nothing Then Throw New ArgumentException("Choose a facility or loan.")
            Dim sections = SelectedTargets
            If sections.Count = 0 Then Throw New ArgumentException("Select at least one section.")
            Dim units = sections.Select(Function(t) FundingScheduleWriter.IsPercentage(FileManager.ExcelModels(ModelID).WB, t)).Distinct().ToList()
            FigureUnits.Text = If(units.Count > 1, "Mixed amounts and percentages — use separate fixed-figure schedules.",
                If(units(0), "Percentage: enter 5 for 5%.", "Same units as the grid (for example, £'000)."))
            If PopulateChoice.Checked AndAlso units.Count > 1 Then Throw New ArgumentException(FigureUnits.Text)
            If StartChoice.EditValue Is Nothing OrElse (EndChoice.SelectedIndex = 0 AndAlso EndDateChoice.EditValue Is Nothing) Then Throw New ArgumentException("Choose a start date and an end date.")
            Dim request As New FundingScheduleRequest With {.StartDate = StartChoice.DateTime.Date,
                .IntervalMonths = New Integer() {1, 3, 6, 12}(IntervalChoice.SelectedIndex), .Rule = FundingDateRule.SameDay,
                .EndDate = If(EndChoice.SelectedIndex = 0, CType(EndDateChoice.DateTime.Date, DateTime?), Nothing),
                .Occurrences = CInt(CountChoice.Value)}
            PreviewDates = FundingScheduleGenerator.Generate(request)
            Grid.DataSource = PreviewDates
            Dim capacity As New List(Of String)
            For Each section In sections
                Dim plan = FundingScheduleWriter.Placement(FileManager.ExcelModels(ModelID).WB, section, PreviewDates.Count)
                capacity.Add(section.Caption & ": " & If(plan.RowsToAdd = 0, "one existing blank block", "add " & plan.RowsToAdd.ToString() & " rows (five spare)") & "; dates stay together")
            Next
            Status.Text = PreviewDates.Count.ToString() & " dates × " & sections.Count.ToString() & " sections. " & String.Join("; ", capacity) & "." & vbCrLf &
                If(PopulateChoice.Checked, "Fixed figures go into available cells only; locked or occupied cells are skipped and reported. ", "Amounts stay blank. ") &
                "Undo removes dates and any figures, not added rows."
            View.Appearance.Row.BackColor = FundingScheduleGroups.SoftColour(ScheduleColour)
            ApplyButton.Enabled = True
        Catch ex As Exception
            If IsDisposed OrElse Disposing OrElse currentRevision <> Revision Then Return
            PreviewDates = Nothing : Grid.DataSource = Nothing
            Status.Text = ex.Message
        Finally
            If Not IsDisposed Then
                Busy = False
                If currentRevision <> Revision Then PreviewTimer.Start()
            End If
        End Try
    End Sub
End Class
