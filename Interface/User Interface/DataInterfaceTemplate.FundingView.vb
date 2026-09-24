Imports Abovo
Imports Abovo.DataObject
Imports Abovo.FileManager
Imports System.Drawing
Imports System.Windows.Forms
Imports DevExpress.Utils
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.XtraVerticalGrid.Rows
Imports DevExpress.XtraVerticalGrid.Events
Imports DevExpress.XtraBars.Docking2010

Partial Public Class DataInterfaceTemplate
    Private Class FundingViewState
        Public Key As String
        Public Funder As String = ""
        Public Facility As String = ""
        Public Identities As New Dictionary(Of Integer, FundingScheduleFacility)
        Public Updating As Boolean
    End Class
    Private ReadOnly FundingViews As New Dictionary(Of VGridControl, FundingViewState)
    Private FundingFunderMenu As ToolStripMenuItem
    Private FundingFacilityMenu As ToolStripMenuItem
    Private FundingClearMenu As ToolStripMenuItem
    Private FundingFilterButton As WindowsUIButton
    Private FundingNotice As DevExpress.XtraEditors.LabelControl
    Private FundingContextGrid As VGridControl
    Private FundingContextRecord As Integer
    Private FundingContextRow As BaseRow

    Private Function ContextFundingDateRange() As String
        Dim grid = TryCast(LastClipboardTarget, VGridControl)
        Dim tag = If(grid Is Nothing, Nothing, GetVGridColumnTag(If(FundingContextGrid Is grid, FundingContextRow, grid.FocusedRow)))
        Return FundingScheduleWriter.Targets.FirstOrDefault(Function(t) String.Equals(t.Caption, tag?.BandID, StringComparison.OrdinalIgnoreCase))?.DateRange
    End Function

    Private Sub ConfigureFundingView(grid As VGridControl)
        grid.OptionsView.ShowRecordHeaders = True
        grid.OptionsView.ShowFilterPanelMode = ShowFilterPanelMode.Never
        grid.RecordHeaderHeight = Math.Max(grid.RecordHeaderHeight, CInt(grid.Font.GetHeight() * 1.8))
        If FundingViews.ContainsKey(grid) Then Return
        Dim state As New FundingViewState With {.Key = "FundingFilter/" & ExcelModels(ModelID).FileName.ToUpperInvariant() & "/" & CSID.ToString() & "/" & GridPreferenceIdentity(grid)}
        state.Funder = GridPresentation.ReadPreference(state.Key, "Funder", "")
        state.Facility = GridPresentation.ReadPreference(state.Key, "Facility", "")
        FundingViews.Add(grid, state)
        RefreshFundingIdentities(grid, state)
        AddHandler grid.CustomRecordHeaderDisplayText, Sub(sender, e)
            Dim identity As FundingScheduleFacility = Nothing
            If state.Identities.TryGetValue(grid.GetDataSourceRecordIndex(e.Record), identity) Then e.DisplayText = identity.LoanName
        End Sub
        AddHandler grid.CustomDrawRecordHeader, AddressOf DrawFundingRecordBoundary
        'ConfigureFundingView runs after the in-place row-header editors have
        'installed their painters. Overlay the fixed-pane edge without repainting them.
        AddHandler grid.CustomDrawRowHeaderCell, AddressOf DrawFundingRowHeaderBoundary
        AddHandler grid.CustomRecordFilter, Sub(sender, e)
            Dim identity As FundingScheduleFacility = Nothing
            state.Identities.TryGetValue(e.RecordIndex, identity)
            e.Visible = identity Is Nothing OrElse
                ((state.Funder = "" OrElse String.Equals(state.Funder, identity.FunderName, StringComparison.OrdinalIgnoreCase)) AndAlso
                 (state.Facility = "" OrElse String.Equals(state.Facility, identity.FacilityName, StringComparison.OrdinalIgnoreCase)))
            e.Handled = True
        End Sub
        Dim tips As New ToolTipController()
        grid.ToolTipController = tips
        AddHandler tips.GetActiveObjectInfo, Sub(sender, e)
            Dim hit = grid.CalcHitInfo(e.ControlMousePosition)
            Dim identity As FundingScheduleFacility = Nothing
            If hit.RecordIndex >= 0 AndAlso state.Identities.TryGetValue(grid.GetDataSourceRecordIndex(hit.RecordIndex), identity) Then
                e.Info = New ToolTipControlInfo(grid.Name & "/" & identity.ColumnIndex.ToString(), identity.Hint)
            End If
        End Sub
        AddHandler grid.CellValueChanged, Sub()
            If state.Updating OrElse grid.IsDisposed OrElse Not grid.IsHandleCreated Then Return
            grid.BeginInvoke(New MethodInvoker(Sub()
                If grid.IsDisposed Then Return
                RefreshFundingViewRecords(grid)
                UpdateFundingFilterIndicator()
            End Sub))
        End Sub
        AddHandler grid.Disposed, Sub()
            FundingViews.Remove(grid)
            tips.Dispose()
        End Sub
        RefilterFundingRecords(grid, state)
        UpdateFundingFilterIndicator()
    End Sub

    Private Shared Sub DrawFundingRecordBoundary(sender As Object, e As CustomDrawRecordHeaderEventArgs)
        e.DefaultDraw()
        'This is a screen boundary for the pinned loan headings, not a workbook
        'fill or a zoom-scaled metric. Keep it three device pixels thick.
        Dim height = Math.Min(3, e.Bounds.Height)
        If height > 0 Then
            e.Cache.FillRectangle(e.Cache.GetSolidBrush(Color.SteelBlue),
                                  New Rectangle(e.Bounds.Left, e.Bounds.Bottom - height, e.Bounds.Width, height))
        End If
        e.Handled = True
    End Sub

    Private Shared Sub DrawFundingRowHeaderBoundary(sender As Object, e As CustomDrawRowHeaderCellEventArgs)
        'Category captions span the scrolling records too; only data row headers
        'belong to the fixed date/title pane. Keep the separator inside its existing
        'padding so neither the in-place editor nor its hit-test bounds change.
        If Not TypeOf e.Row Is EditorRow Then Return
        If Not e.Handled Then e.DefaultDraw()
        Dim width = Math.Min(3, e.Bounds.Width)
        If width > 0 Then
            Dim left = If(e.IsRightToLeft, e.Bounds.Left, e.Bounds.Right - width)
            e.Cache.FillRectangle(e.Cache.GetSolidBrush(Color.SteelBlue),
                                  New Rectangle(left, e.Bounds.Top, width, e.Bounds.Height))
        End If
        e.Handled = True
    End Sub

    Private Sub RefreshFundingIdentities(grid As VGridControl, state As FundingViewState)
        Dim source = TryCast(grid.DataSource, AbovoUnboundSource)
        If source Is Nothing Then Return
        Dim data = DataPres.DataSets(source.UBSTag.DSIndex)
        Dim wb = ExcelModels(ModelID).WB
        Dim identities = FundingScheduleGroups.Facilities(wb, False).Concat(FundingScheduleGroups.Facilities(wb, True)).GroupBy(Function(f) f.ColumnIndex).ToDictionary(Function(g) g.Key, Function(g) g.First())
        state.Identities.Clear()
        For i = 0 To data.DataRows.Count - 1
            For column = 0 To data.DataColumns.Count - 1
                If data.DataColumns(column).ColumnTag.IsDummyColumn Then Continue For
                Dim cell = InputSourceCell(source, i, column), identity As FundingScheduleFacility = Nothing
                If cell IsNot Nothing AndAlso cell.Worksheet.Name = "Funding Assumptions" AndAlso identities.TryGetValue(cell.ColumnIndex, identity) Then state.Identities(i) = identity
                Exit For
            Next
        Next
    End Sub

    Private Sub UpdateFundingFilterMenu(identity As FundingScheduleFacility)
        If FundingFunderMenu Is Nothing Then
            FundingFunderMenu = New ToolStripMenuItem()
            FundingFacilityMenu = New ToolStripMenuItem()
            FundingClearMenu = New ToolStripMenuItem("Remove filters", Nothing, Sub() ClearFundingFilters())
            ClipboardContextMenu.Items.AddRange({FundingFunderMenu, FundingFacilityMenu, FundingClearMenu})
            AddHandler FundingFunderMenu.Click, Sub() SetContextFundingFilter(True)
            AddHandler FundingFacilityMenu.Click, Sub() SetContextFundingFilter(False)
        End If
        'Visible is False while the popup itself is closed, even immediately
        'after setting it True. Do not use that effective value for its sibling.
        Dim showLoanFilters = identity IsNot Nothing AndAlso Not identity.IsInvestment
        FundingFunderMenu.Visible = showLoanFilters
        FundingFacilityMenu.Visible = showLoanFilters
        FundingFunderMenu.Text = "Filter to funder " & If(identity?.FunderName, "")
        FundingFacilityMenu.Text = "Filter to facility " & If(identity?.FacilityName, "")
        FundingFunderMenu.Enabled = Not String.IsNullOrWhiteSpace(identity?.FunderName)
        FundingFacilityMenu.Enabled = Not String.IsNullOrWhiteSpace(identity?.FacilityName)
        FundingClearMenu.Visible = FundingViews.Values.Any(Function(s) s.Funder <> "" OrElse s.Facility <> "")
        ClipboardCopyWithHeadersMenuItem.Text = If(identity Is Nothing, "Copy with headings", "Copy with Header and Row titles")
    End Sub

    Private Sub SetContextFundingFilter(byFunder As Boolean)
        Dim investment As Boolean, identity = ContextFundingFacility(investment)
        Dim grid = TryCast(LastClipboardTarget, VGridControl), state As FundingViewState = Nothing
        If identity Is Nothing OrElse grid Is Nothing OrElse Not FundingViews.TryGetValue(grid, state) OrElse Not CommitEditorsForSave() Then Return
        state.Funder = If(byFunder, identity.FunderName, "")
        state.Facility = If(byFunder, "", identity.FacilityName)
        ApplyFundingFilter(grid, state)
    End Sub

    Private Sub ApplyFundingFilter(grid As VGridControl, state As FundingViewState)
        state.Updating = True
        Try
            grid.HideEditor() : grid.ClearSelection()
            RefreshFundingIdentities(grid, state)
            'Changing a custom predicate does not change ActiveFilter.Criteria.
            'Toggle the native filter to invalidate its cached record mapping.
            RefilterFundingRecords(grid, state)
            GridPresentation.WritePreference(state.Key, "Funder", state.Funder)
            GridPresentation.WritePreference(state.Key, "Facility", state.Facility)
        Finally
            state.Updating = False
        End Try
        UpdateFundingFilterIndicator()
    End Sub

    Private Sub RefreshFundingViewRecords(grid As VGridControl)
        Dim state As FundingViewState = Nothing
        If Not FundingViews.TryGetValue(grid, state) OrElse state.Updating Then Return
        RefreshFundingIdentities(grid, state)
        RefilterFundingRecords(grid, state)
        grid.InvalidateRecordHeaders()
    End Sub

    Private Sub RefilterFundingRecords(grid As VGridControl, state As FundingViewState)
        '25.2's unbound controller bypasses custom filtering when there is no
        'active criterion. This non-restrictive native criterion enables the
        'predicate; it never changes or rebinds the original source records.
        Dim row = grid.Rows.Cast(Of BaseRow)().SelectMany(Function(r) r.ChildRows.Cast(Of BaseRow)()).OfType(Of EditorRow)().FirstOrDefault()
        grid.ActiveFilterEnabled = False
        grid.ActiveFilterString = If(row IsNot Nothing,
                                    "[" & row.Properties.FieldName & "] Is Null Or [" & row.Properties.FieldName & "] Is Not Null", "")
        grid.ActiveFilterEnabled = True
        grid.FilterRecords()
        grid.LayoutChanged()
    End Sub

    Private Sub ClearFundingFilters()
        If Not CommitEditorsForSave() Then Return
        For Each pair In FundingViews.ToArray()
            pair.Value.Funder = "" : pair.Value.Facility = ""
            ApplyFundingFilter(pair.Key, pair.Value)
        Next
    End Sub

    Private Sub UpdateFundingFilterIndicator()
        Dim active = FundingViews.Where(Function(p) p.Value.Funder <> "" OrElse p.Value.Facility <> "").ToList()
        If FundingFilterButton Is Nothing Then
            Dim svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><circle cx='16' cy='16' r='14' fill='none' stroke='#de8500' stroke-width='1.5'/><path d='M7 9h18l-7 8v7l-4-2v-5z' fill='#de8500'/></svg>"
            Using stream As New IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(svg))
                Dim options As New WindowsUIButtonImageOptions With {.SvgImage = DevExpress.Utils.Svg.SvgImage.FromStream(stream), .SvgImageColorizationMode = SvgImageColorizationMode.None}
                FundingFilterButton = New WindowsUIButton("Filter applied", False, options, ButtonStyle.PushButton,
                    "Some loan columns are hidden. Click to remove Funding filters.", -1, True, Nothing, True, False, True, "FundingFilter", -1, False)
            End Using
            FundingFilterButton.Appearance.ForeColor = Color.DarkOrange
            FundingFilterButton.Appearance.Options.UseForeColor = True
            FundingFilterButton.UseCaption = True
            'WindowsUI always skins glyphs; its supported HTML caption keeps
            'the active-filter warning orange independently of that skin.
            WindowsUIButtonPanelActions.AllowHtmlDraw = True
            FundingFilterButton.Caption = "<color=#DE8500>Filter applied</color>"
            WindowsUIButtonPanelActions.Buttons.Add(FundingFilterButton)
            AddHandler WindowsUIButtonPanelActions.ButtonClick, Sub(sender, e)
                If e.Button Is FundingFilterButton Then ClearFundingFilters()
            End Sub
        End If
        FundingFilterButton.Visible = active.Count > 0
        FundingFilterButton.ToolTip = If(active.Count = 0, "", "FILTER APPLIED — " & String.Join("; ", active.Select(Function(p) If(p.Value.Funder <> "", "Funder: " & p.Value.Funder, "Facility: " & p.Value.Facility) & " (" & p.Key.RecordCount.ToString() & " shown)")) & ". Click to remove filters.")
        WindowsUIButtonPanelActions.Invalidate()
        PresentationLayout.ApplyButtonPanel(WindowsUIButtonPanelActions, Me)
    End Sub

    Private Sub ShowFundingScheduleNotice(skipped As IList(Of String))
        Dim message = "Schedule added. " & skipped.Count.ToString() & " fixed amount cell(s) were skipped because they are locked, contain a formula or already have a value. Dates and other amounts were kept."
        SystemMessageManager.Publish(ModelID, message & " " & String.Join("; ", skipped), SystemMessageSeverity.Warning, "Funding schedule")
        If FundingNotice Is Nothing Then
            FundingNotice = New DevExpress.XtraEditors.LabelControl With {.Dock = DockStyle.Top, .AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Vertical, .Padding = New Padding(10)}
            FundingNotice.Appearance.BackColor = Color.LightGoldenrodYellow
            FundingNotice.Appearance.ForeColor = Color.DarkGoldenrod
            FundingNotice.Appearance.Options.UseBackColor = True
            Controls.Add(FundingNotice)
            FundingNotice.BringToFront()
            AddHandler FundingNotice.Click, Sub() FundingNotice.Visible = False
        End If
        FundingNotice.Text = message & " (Click to dismiss.)"
        FundingNotice.ToolTip = String.Join(vbCrLf, skipped.Take(12))
        FundingNotice.Visible = True
    End Sub

    Private Sub CopyFundingWithTitles(grid As VGridControl)
        Dim selected = grid.GetSelectedCells().ToList()
        If selected.Count = 0 AndAlso grid.FocusedRow IsNot Nothing AndAlso grid.FocusedRecord >= 0 Then grid.SelectCell(grid.FocusedRecord, grid.FocusedRow, 0) : selected = grid.GetSelectedCells().ToList()
        If selected.Count = 0 Then Return
        Dim records = selected.Select(Function(c) c.RecordIndex).Distinct().OrderBy(Function(i) i).ToList()
        Dim rows = selected.Select(Function(c) c.Row).Distinct().OrderBy(Function(r) r.VisibleIndex).ToList()
        Dim clean As Func(Of String, String) = Function(s) If(s, "").Replace(vbTab, " ").Replace(vbCr, " ").Replace(vbLf, " ")
        Dim lines As New List(Of String) From {"Section / date" & vbTab & String.Join(vbTab, records.Select(Function(i) clean(grid.GetRecordHeaderText(i))))}
        For Each row In rows
            Dim tag = GetVGridColumnTag(row)
            Dim title = row.Properties.Caption
            If tag IsNot Nothing Then
                Dim source = TryCast(grid.DataSource, AbovoUnboundSource)
                Dim sourceCell = InputSourceCell(source, grid.GetDataSourceRecordIndex(records(0)), GetVGridColumnIndex(row))
                Dim target = FundingScheduleWriter.Targets.FirstOrDefault(Function(t) t.Caption = tag.BandID)
                If target IsNot Nothing AndAlso sourceCell IsNot Nothing Then
                    Dim dateRange = FundingScheduleWriter.Ranges(ExcelModels(ModelID).WB, target).Item1
                    title = target.Caption & " — " & sourceCell.Worksheet.Cells(sourceCell.RowIndex, dateRange.LeftColumnIndex).DisplayText
                End If
            End If
            lines.Add(clean(title) & vbTab & String.Join(vbTab, records.Select(Function(i) If(selected.Any(Function(c) c.Row Is row AndAlso c.RecordIndex = i), clean(grid.GetCellDisplayText(row, i)), ""))))
        Next
        SetClipboardTextWithHeaders(String.Join(vbCrLf, lines), 1, 1)
    End Sub
End Class
