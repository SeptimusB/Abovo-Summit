Imports Abovo
Imports Abovo.DataObject
Imports DevExpress.XtraEditors
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.XtraVerticalGrid.Rows
Imports System.Drawing

Partial Public Class DataInterfaceTemplate
    Private FundingTints As New Dictionary(Of String, FundingScheduleTint)

    Private Function FundingTint(cell As DevExpress.Spreadsheet.Cell) As Color?
        If cell Is Nothing OrElse cell.Worksheet.Name <> "Funding Assumptions" Then Return Nothing
        Dim tint As FundingScheduleTint = Nothing
        If FundingTints.TryGetValue(cell.RowIndex.ToString() & ":" & cell.ColumnIndex.ToString(), tint) AndAlso tint.Active Then Return tint.Colour
        Return Nothing
    End Function

    Private Sub DrawFundingScheduleMarker(grid As VGridControl, e As DevExpress.XtraVerticalGrid.Events.CustomDrawRowValueCellEventArgs)
        If FundingTints.Count = 0 OrElse e.Row Is Nothing Then Return
        Dim cell = InputSourceCell(TryCast(grid.DataSource, AbovoUnboundSource), grid.GetDataSourceRecordIndex(e.RecordIndex), GetVGridColumnIndex(e.Row, e.CellIndex))
        If cell Is Nothing OrElse cell.Worksheet.Name <> "Funding Assumptions" Then Return
        Dim tint As FundingScheduleTint = Nothing
        If Not FundingTints.TryGetValue(cell.RowIndex.ToString() & ":" & cell.ColumnIndex.ToString(), tint) OrElse Not tint.Active Then Return
        'Keep native selection/unavailable shading, with a narrow identity marker
        'even when the selected facility is currently unavailable for entry.
        e.DefaultDraw()
        e.Cache.FillRectangle(tint.MarkerColour, New Rectangle(e.Bounds.Left, e.Bounds.Top + 1, Math.Min(PresentationScaleManager.Scale(3), e.Bounds.Width), Math.Max(1, e.Bounds.Height - 2)))
        e.Handled = True
    End Sub

    Private Sub ConfigureFundingScheduleGrid(grid As VGridControl)
        If GSID <> 0 OrElse (CSID <> 33 AndAlso CSID <> 138) Then Return
        Try
            FundingTints = FundingScheduleGroups.TintIndex(Abovo.FileManager.ExcelModels(ModelID).WB)
        Catch ex As Exception
            FundingTints.Clear()
            'Bad optional presentation metadata must not break workbook access.
            Abovo.SummitDiagnostics.WriteLine("Funding schedule appearance: " & ex.Message)
        End Try
        ConfigureFundingView(grid)
        grid.OptionsBehavior.RecordsMouseWheel = False
    End Sub

    Private FundingScheduleMenuItem As ToolStripMenuItem
    Private FundingScheduleMenuSeparator As ToolStripSeparator

    Private Sub InitialiseFundingScheduleMenu()
        FundingScheduleMenuSeparator = New ToolStripSeparator With {.Visible = False}
        FundingScheduleMenuItem = New ToolStripMenuItem("Add schedule…") With {.Visible = False}
        ClipboardContextMenu.Items.Add(FundingScheduleMenuSeparator)
        ClipboardContextMenu.Items.Add(FundingScheduleMenuItem)
        AddHandler FundingScheduleMenuItem.Click,
            Sub()
                Dim investment As Boolean
                Dim facility = ContextFundingFacility(investment)
                If facility Is Nothing Then Return
                Dim targets = FundingScheduleWriter.Targets.Where(Function(t) FundingScheduleGroups.IsInvestment(t) = investment).ToArray()
                ShowFundingSchedule(targets, ContextFundingDateRange(), facility.ColumnIndex)
            End Sub
    End Sub

    Private Function ContextFundingFacility(ByRef investment As Boolean) As FundingScheduleFacility
        If GSID <> 0 OrElse (CSID <> 33 AndAlso CSID <> 138) Then Return Nothing
        Dim grid = TryCast(LastClipboardTarget, VGridControl)
        If grid Is Nothing OrElse grid.IsDisposed Then Return Nothing
        Dim row = If(FundingContextGrid Is grid, FundingContextRow, grid.FocusedRow)
        Dim record = If(FundingContextGrid Is grid, FundingContextRecord, grid.FocusedRecord)
        If record < 0 OrElse record >= grid.RecordCount Then Return Nothing
        If row Is Nothing OrElse GetVGridColumnTag(row) Is Nothing Then
            row = grid.Rows.Cast(Of BaseRow)().SelectMany(Function(r) r.ChildRows.Cast(Of BaseRow)()).
                FirstOrDefault(Function(r) GetVGridColumnTag(r) IsNot Nothing)
        End If
        Dim tag = GetVGridColumnTag(row)
        If tag Is Nothing OrElse tag.IsDummyColumn Then Return Nothing
        Dim cell = InputSourceCell(TryCast(grid.DataSource, AbovoUnboundSource), grid.GetDataSourceRecordIndex(record), GetVGridColumnIndex(row))
        If cell Is Nothing OrElse cell.Worksheet.Name <> "Funding Assumptions" Then Return Nothing
        investment = tag.BandID IsNot Nothing AndAlso tag.BandID.StartsWith("Investment", StringComparison.Ordinal)
        Return FundingScheduleGroups.Facilities(Abovo.FileManager.ExcelModels(ModelID).WB, investment).
            FirstOrDefault(Function(f) f.ColumnIndex = cell.ColumnIndex)
    End Function

    Private Sub UpdateFundingScheduleMenu()
        If FundingScheduleMenuItem Is Nothing Then Return
        Dim investment As Boolean
        Dim facility = ContextFundingFacility(investment)
        FundingScheduleMenuItem.Visible = facility IsNot Nothing
        FundingScheduleMenuSeparator.Visible = facility IsNot Nothing
        FundingScheduleMenuItem.ToolTipText = If(facility Is Nothing, "", "Create a schedule for " & facility.Caption)
        UpdateFundingFilterMenu(facility)
    End Sub

    Private Function GetFundingScheduleAction(dateRange As String) As Action
        If GSID <> 0 OrElse (CSID <> 33 AndAlso CSID <> 138) Then Return Nothing
        Dim target = FundingScheduleWriter.Targets.FirstOrDefault(Function(t) t.DateRange = dateRange)
        If target Is Nothing Then Return Nothing
        Return Sub() ShowFundingSchedule({target}, dateRange, -1)
    End Function

    Private Sub ShowFundingSchedule(targets As IEnumerable(Of FundingScheduleTarget), dateRange As String, facilityColumn As Integer)
        If Not CommitEditorsForSave() Then Return
        Using dialog As New FundingScheduleForm(ModelID, targets, dateRange, facilityColumn)
            If dialog.ShowDialog(FindForm()) <> DialogResult.OK Then Return
            Try
                Dim target = dialog.SelectedTarget
                Dim selectedColumn = dialog.SelectedFacility.ColumnIndex
                Dim skipped As New List(Of String)
                Dim applied As FundingScheduleApplication
                Using splash As New FormSplashScreen(FindForm(), "Creating Funding schedule", "Adding date rows and refreshing the available inputs…")
                    applied = FundingScheduleWriter.ApplyAvailableAmountsWithLocations(ModelID, dialog.SelectedTargets, dialog.DatesToApply, dialog.SelectedFacility, dialog.ScheduleColour, dialog.ConfigurationSummary, dialog.FixedFigure, skipped)
                    UpdateAllRules()
                            RebuildAllSections()
                        End Using
                        If skipped.Count > 0 Then
                            ShowFundingScheduleNotice(skipped)
                        End If
                        Dim firstDateRow = applied.FirstDateRows(target.DateRange)
                        Dim targetTab = If(FundingScheduleGroups.IsInvestment(target), "Investments",
                            If({"IR_Fund_Var_Int", "IR_Fund_Min_Bal", "IR_Fund_Int_Rec", "IR_Fund_Int_Pay"}.Contains(target.DateRange), "Variable and Cash Rates", "Facilities and Loans"))
                        For Each page As DevExpress.XtraTab.XtraTabPage In XtraTabControlNewGIT.TabPages
                            If page.Text.Trim() = targetTab Then XtraTabControlNewGIT.SelectedTabPage = page : Exit For
                        Next
                        EnsureSectionBuilt(XtraTabControlNewGIT.SelectedTabPageIndex)
                        BeginInvoke(New MethodInvoker(Sub() FocusFundingScheduleCell(target.Caption, firstDateRow, selectedColumn)))
                    Catch ex As Exception
                        UpdateAllRules()
                        RebuildAllSections()
                        XtraMessageBox.Show(FindForm(), ex.Message, "Funding schedule", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End Try
                End Using
    End Sub

    Private Sub FocusFundingScheduleCell(section As String, worksheetRow As Integer, Optional facilityColumn As Integer = -1)
        If IsDisposed OrElse Disposing OrElse VertGridControls Is Nothing Then Return
        For Each grid In VertGridControls
            If grid Is Nothing OrElse grid.IsDisposed OrElse Not grid.Visible Then Continue For
            For Each category In grid.Rows.OfType(Of CategoryRow)()
                If category.Properties.Caption.Trim() = section Then category.Expanded = True
            Next
        Next
        Dim candidates = NavigationPositions().Where(Function(p) p.VRow IsNot Nothing).
            Where(Function(p)
                      Dim tag = GetVGridColumnTag(p.VRow)
                      If tag Is Nothing OrElse tag.BandID <> section Then Return False
                      Dim source = TryCast(DirectCast(p.Host, VGridControl).DataSource, AbovoUnboundSource)
                      Dim grid = DirectCast(p.Host, VGridControl)
                      Dim cell = InputSourceCell(source, grid.GetDataSourceRecordIndex(Math.Max(0, p.Record)), GetVGridColumnIndex(p.VRow))
                      Return cell IsNot Nothing AndAlso cell.Worksheet.Name = "Funding Assumptions" AndAlso cell.RowIndex = worksheetRow
                  End Function).ToList()
        Dim destination = candidates.Where(Function(p) p.Record >= 0 AndAlso (facilityColumn < 0 OrElse
            InputSourceCell(TryCast(DirectCast(p.Host, VGridControl).DataSource, AbovoUnboundSource), DirectCast(p.Host, VGridControl).GetDataSourceRecordIndex(p.Record), GetVGridColumnIndex(p.VRow)).ColumnIndex = facilityColumn)).OrderBy(Function(p) p.Record).FirstOrDefault()
        'A blank model may not yet define any usable loan columns. Do not
        'override its locks; leave the new date visible and focused instead.
        If destination Is Nothing Then destination = candidates.FirstOrDefault(Function(p) p.VHeader IsNot Nothing)
        If destination IsNot Nothing Then ActivateEditorPosition(destination)
    End Sub
End Class
