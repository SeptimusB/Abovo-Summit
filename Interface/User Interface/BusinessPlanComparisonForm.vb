Imports System.ComponentModel
Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraTreeList
Imports DevExpress.XtraTreeList.Columns
Imports DevExpress.XtraTab
Imports DevExpress.Spreadsheet

Public Class BusinessPlanComparisonForm
    Inherits XtraForm

    Private NotInheritable Class ModelChoice
        Public Property Participant As BusinessPlanComparisonParticipant
        Public ReadOnly Property Key As String
            Get
                Return Participant.Key
            End Get
        End Property
        Public ReadOnly Property DisplayName As String
            Get
                Return Participant.DisplayName
            End Get
        End Property
        Public Overrides Function ToString() As String
            Return Participant.DisplayName
        End Function
    End Class

    Private ReadOnly MainScreen As FormMainScreen
    Private ReadOnly BaseModelEdit As New LookUpEdit()
    Private ReadOnly ComparedModelsList As New CheckedListBoxControl()
    Private ReadOnly ResultsTree As New TreeList()
    Private ReadOnly StatusLabel As New LabelControl()
    Private ReadOnly RunButton As New SimpleButton()
    Private ReadOnly AssessMigrationButton As New SimpleButton()
    Private ReadOnly PopulateButton As New SimpleButton()
    Private ReadOnly UpgradeSourceEdit As New LookUpEdit()
    Private ReadOnly UpgradeTargetEdit As New LookUpEdit()
    Private ReadOnly UpgradeResultsTree As New TreeList()
    Private ReadOnly UpgradeStatusLabel As New LabelControl()
    Private Models As New BindingList(Of ModelChoice)()
    Private LastResult As BusinessPlanComparisonResult
    Private LastUpgradeResult As BusinessPlanMigrationResult
    Private ReadOnly ExternalWorkbooks As New Dictionary(Of String, Workbook)(StringComparer.OrdinalIgnoreCase)
    Private AssumptionWorksheetNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Public Sub New(ByVal ownerScreen As FormMainScreen)
        MainScreen = ownerScreen
        Text = "Compare Business Plans"
        StartPosition = FormStartPosition.CenterParent
        MinimumSize = New Size(1000, 650)
        Size = New Size(1450, 850)
        Icon = ownerScreen.Icon
        BuildInterface()
        RefreshModels()
    End Sub

    Public Sub RefreshModels()
        Dim previousBase As String = SelectedBaseKey()
        Dim checkedModels As New HashSet(Of String)(SelectedComparisonKeys(), StringComparer.OrdinalIgnoreCase)
        Dim previousUpgradeSource As String = SelectedLookupKey(UpgradeSourceEdit)
        Dim previousUpgradeTarget As String = SelectedLookupKey(UpgradeTargetEdit)
        Dim openPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Models = New BindingList(Of ModelChoice)()
        AssumptionWorksheetNames = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If ExcelModels IsNot Nothing Then
            For modelID As Integer = 0 To ExcelModels.Length - 1
                Dim model As ExcelModel = ExcelModels(modelID)
                If model Is Nothing OrElse model.Profile Is Nothing OrElse
                   Not String.Equals(model.Profile.ModelType, "AbovoBP", StringComparison.OrdinalIgnoreCase) Then Continue For
                Dim company As String = If(model.WBStructure Is Nothing, String.Empty, model.WBStructure.CompanyName)
                Dim participant As New BusinessPlanComparisonParticipant With {
                    .Key = "model:" & modelID.ToString(),
                    .ModelID = modelID,
                    .DisplayName = "(" & (modelID + 1).ToString() & ") " & company & " — " & IO.Path.GetFileName(model.FileName) & " [Open]",
                    .FilePath = model.FileName,
                    .Workbook = model.WB}
                Models.Add(New ModelChoice With {.Participant = participant})
                openPaths.Add(model.FileName)
                AssumptionWorksheetNames.UnionWith(BusinessPlanComparisonService.GetAssumptionWorksheetNames(model))
            Next
        End If

        For Each pair As KeyValuePair(Of String, Workbook) In ExternalWorkbooks
            If openPaths.Contains(pair.Key) Then Continue For
            Dim participant As New BusinessPlanComparisonParticipant With {
                .Key = "file:" & pair.Key,
                .DisplayName = IO.Path.GetFileName(pair.Key) & " [Read-only source]",
                .FilePath = pair.Key,
                .Workbook = pair.Value}
            Models.Add(New ModelChoice With {.Participant = participant})
        Next

        ConfigureLookupData(BaseModelEdit, Models)
        ComparedModelsList.Items.Clear()
        For Each choice As ModelChoice In Models
            ComparedModelsList.Items.Add(New CheckedListBoxItem(choice, checkedModels.Contains(choice.Key)))
        Next
        If Models.Any(Function(item) String.Equals(item.Key, previousBase, StringComparison.OrdinalIgnoreCase)) Then
            BaseModelEdit.EditValue = previousBase
        ElseIf Models.Count > 0 Then
            BaseModelEdit.EditValue = Models(0).Key
        Else
            BaseModelEdit.EditValue = Nothing
        End If

        ConfigureLookupData(UpgradeSourceEdit, Models)
        Dim openTargets As New BindingList(Of ModelChoice)(Models.Where(Function(item) item.Participant.IsOpenSummitModel).ToList())
        ConfigureLookupData(UpgradeTargetEdit, openTargets)
        If Models.Any(Function(item) String.Equals(item.Key, previousUpgradeSource, StringComparison.OrdinalIgnoreCase)) Then
            UpgradeSourceEdit.EditValue = previousUpgradeSource
        Else
            Dim defaultSource As ModelChoice = Models.FirstOrDefault(Function(item) Not item.Participant.IsOpenSummitModel)
            If defaultSource Is Nothing Then defaultSource = Models.FirstOrDefault()
            UpgradeSourceEdit.EditValue = If(defaultSource Is Nothing, Nothing, defaultSource.Key)
        End If
        If openTargets.Any(Function(item) String.Equals(item.Key, previousUpgradeTarget, StringComparison.OrdinalIgnoreCase)) Then
            UpgradeTargetEdit.EditValue = previousUpgradeTarget
        Else
            Dim defaultTarget As ModelChoice = openTargets.FirstOrDefault(
                Function(item) Not String.Equals(item.Key, SelectedLookupKey(UpgradeSourceEdit), StringComparison.OrdinalIgnoreCase))
            UpgradeTargetEdit.EditValue = If(defaultTarget Is Nothing, Nothing, defaultTarget.Key)
        End If
        UpdateActionState()
    End Sub
    Private Sub BuildInterface()
        Text = "Compare and Upgrade Business Plans"
        Dim tabs As New XtraTabControl With {.Dock = DockStyle.Fill}
        Dim comparePage As New XtraTabPage With {.Text = "Compare"}
        Dim upgradePage As New XtraTabPage With {.Text = "Upgrade"}
        Dim managerPage As New XtraTabPage With {.Text = "Structure Manager"}
        Dim managerButton As New SimpleButton With {.Text = "Open Structure Manager...", .Dock = DockStyle.Top, .Height = 42}
        Dim explanation As New LabelControl With {.Dock = DockStyle.Top, .Height = 80, .AutoSizeMode = LabelAutoSizeMode.None,
            .Text = "Author structure XML beside a read-only workbook and a real DIT preview, in an isolated window." & Environment.NewLine &
                    "Includes existing generic/client definitions and bespoke-rule review. Drafts only; no automatic migration or publication."}
        managerPage.Controls.Add(explanation)
        managerPage.Controls.Add(managerButton)
        AddHandler managerButton.Click, Sub(sender, e)
                                            StructureManagerForm.Launch(Me)
                                        End Sub
        tabs.TabPages.AddRange(New XtraTabPage() {comparePage, upgradePage, managerPage})
        comparePage.Controls.Add(BuildComparePage())
        upgradePage.Controls.Add(BuildUpgradePage())
        Controls.Add(tabs)
    End Sub

    Private Function BuildComparePage() As Control
        Dim root As TableLayoutPanel = CreatePageLayout(170.0F)
        Dim toolbar As FlowLayoutPanel = CreateToolbar()
        Dim addButton As New SimpleButton With {.Text = "Add business plans...", .AutoSize = True}
        Dim refreshButton As New SimpleButton With {.Text = "Refresh list", .AutoSize = True}
        RunButton.Text = "Compare selected"
        RunButton.AutoSize = True
        Dim exportButton As New SimpleButton With {.Text = "Export report...", .AutoSize = True}
        toolbar.Controls.AddRange(New Control() {addButton, refreshButton, RunButton, exportButton, StatusLabel})
        AddHandler addButton.Click, AddressOf AddModels
        AddHandler refreshButton.Click, Sub(sender, e) RefreshModels()
        AddHandler RunButton.Click, AddressOf RunComparison
        AddHandler exportButton.Click, AddressOf ExportReport
        AddHandler BaseModelEdit.EditValueChanged, Sub(sender, e) UpdateActionState()
        AddHandler ComparedModelsList.ItemCheck, Sub(sender, e) BeginInvoke(New Action(AddressOf UpdateActionState))
        root.Controls.Add(toolbar, 0, 0)

        Dim selectors As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 2}
        selectors.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180.0F))
        selectors.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        selectors.RowStyles.Add(New RowStyle(SizeType.Absolute, 42.0F))
        selectors.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        BaseModelEdit.Dock = DockStyle.Fill
        BaseModelEdit.Properties.ShowHeader = False
        BaseModelEdit.Properties.NullText = "Select the base business plan"
        ComparedModelsList.Dock = DockStyle.Fill
        ComparedModelsList.CheckOnClick = True
        selectors.Controls.Add(New LabelControl With {.Text = "Base business plan", .Dock = DockStyle.Fill}, 0, 0)
        selectors.Controls.Add(BaseModelEdit, 1, 0)
        selectors.Controls.Add(New LabelControl With {.Text = "Compare against", .Dock = DockStyle.Fill}, 0, 1)
        selectors.Controls.Add(ComparedModelsList, 1, 1)
        root.Controls.Add(selectors, 0, 1)

        ConfigureResultsTree(ResultsTree)
        root.Controls.Add(ResultsTree, 0, 2)
        Return root
    End Function

    Private Function BuildUpgradePage() As Control
        Dim root As TableLayoutPanel = CreatePageLayout(100.0F)
        Dim toolbar As FlowLayoutPanel = CreateToolbar()
        Dim addButton As New SimpleButton With {.Text = "Add source plans...", .AutoSize = True}
        Dim refreshButton As New SimpleButton With {.Text = "Refresh list", .AutoSize = True}
        AssessMigrationButton.Text = "Assess upgrade"
        AssessMigrationButton.AutoSize = True
        PopulateButton.Text = "Upgrade and save as..."
        PopulateButton.AutoSize = True
        Dim exportButton As New SimpleButton With {.Text = "Export report...", .AutoSize = True}
        toolbar.Controls.AddRange(New Control() {addButton, refreshButton, AssessMigrationButton, PopulateButton, exportButton, UpgradeStatusLabel})
        AddHandler addButton.Click, AddressOf AddModels
        AddHandler refreshButton.Click, Sub(sender, e) RefreshModels()
        AddHandler AssessMigrationButton.Click, AddressOf AssessMigration
        AddHandler PopulateButton.Click, AddressOf PopulateTarget
        AddHandler exportButton.Click, AddressOf ExportUpgradeReport
        AddHandler UpgradeSourceEdit.EditValueChanged, Sub(sender, e) UpdateActionState()
        AddHandler UpgradeTargetEdit.EditValueChanged, Sub(sender, e) UpdateActionState()
        root.Controls.Add(toolbar, 0, 0)

        Dim selectors As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 2}
        selectors.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180.0F))
        selectors.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        selectors.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))
        selectors.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))
        UpgradeSourceEdit.Dock = DockStyle.Fill
        UpgradeSourceEdit.Properties.ShowHeader = False
        UpgradeSourceEdit.Properties.NullText = "Select the populated source plan"
        UpgradeTargetEdit.Dock = DockStyle.Fill
        UpgradeTargetEdit.Properties.ShowHeader = False
        UpgradeTargetEdit.Properties.NullText = "Select one open blank target plan"
        selectors.Controls.Add(New LabelControl With {.Text = "Source business plan", .Dock = DockStyle.Fill}, 0, 0)
        selectors.Controls.Add(UpgradeSourceEdit, 1, 0)
        selectors.Controls.Add(New LabelControl With {.Text = "Blank target plan", .Dock = DockStyle.Fill}, 0, 1)
        selectors.Controls.Add(UpgradeTargetEdit, 1, 1)
        root.Controls.Add(selectors, 0, 1)

        ConfigureResultsTree(UpgradeResultsTree)
        root.Controls.Add(UpgradeResultsTree, 0, 2)
        Return root
    End Function

    Private Shared Function CreatePageLayout(ByVal selectorHeight As Single) As TableLayoutPanel
        Dim root As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 3,
            .Padding = New Padding(12)}
        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, selectorHeight))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        Return root
    End Function

    Private Shared Function CreateToolbar() As FlowLayoutPanel
        Return New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .AutoSize = True,
            .WrapContents = False,
            .Padding = New Padding(0, 0, 0, 8)}
    End Function

    Private Shared Sub ConfigureResultsTree(ByVal tree As TreeList)
        tree.Dock = DockStyle.Fill
        tree.OptionsBehavior.Editable = False
        tree.OptionsSelection.MultiSelect = True
        tree.OptionsView.ShowAutoFilterRow = False
        tree.OptionsView.ShowIndicator = False
        tree.OptionsClipboard.AllowCopy = DefaultBoolean.True
        tree.OptionsClipboard.CopyColumnHeaders = DefaultBoolean.True
    End Sub

    Private Shared Sub ConfigureLookupData(ByVal lookup As LookUpEdit,
                                           ByVal choices As Object)
        lookup.Properties.DataSource = choices
        lookup.Properties.DisplayMember = NameOf(ModelChoice.DisplayName)
        lookup.Properties.ValueMember = NameOf(ModelChoice.Key)
    End Sub
    Private Sub AddModels(ByVal sender As Object, ByVal e As EventArgs)
        Using dialog As New OpenFileDialog With {
            .Title = "Add read-only business plans to the comparison",
            .Filter = "Abovo Business Plans (*.xlsb;*.abp)|*.xlsb;*.abp",
            .Multiselect = True,
            .CheckFileExists = True}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            Cursor = Cursors.WaitCursor
            For Each fileName As String In dialog.FileNames
                Dim fullPath As String = IO.Path.GetFullPath(fileName)
                If FileManager.IsFileOpen(fullPath) OrElse ExternalWorkbooks.ContainsKey(fullPath) Then Continue For
                Dim workbook As New Workbook()
                Try
                    workbook.Options.CalculationMode = WorkbookCalculationMode.Manual
                    workbook.LoadDocument(fullPath)
                    ExternalWorkbooks.Add(fullPath, workbook)
                    workbook = Nothing
                Catch ex As Exception
                    MessageBox.Show(Me, "Could not load '" & IO.Path.GetFileName(fullPath) & "'." &
                                    Environment.NewLine & ex.Message, "Add business plan",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Finally
                    If workbook IsNot Nothing Then workbook.Dispose()
                End Try
            Next
        End Using
        Cursor = Cursors.Default
        RefreshModels()
    End Sub

    Private Sub RunComparison(ByVal sender As Object, ByVal e As EventArgs)
        Dim basePlan As BusinessPlanComparisonParticipant = SelectedBaseParticipant()
        Dim comparisonPlans As List(Of BusinessPlanComparisonParticipant) = SelectedComparisonParticipants().
            Where(Function(item) basePlan Is Nothing OrElse Not String.Equals(item.Key, basePlan.Key, StringComparison.OrdinalIgnoreCase)).ToList()
        If basePlan Is Nothing OrElse comparisonPlans.Count = 0 Then Return
        Cursor = Cursors.WaitCursor
        RunButton.Enabled = False
        Try
            LastResult = BusinessPlanComparisonService.Compare(basePlan, comparisonPlans, AssumptionWorksheetNames)
            ResultsTree.DataSource = LastResult.Items
            ResultsTree.KeyFieldName = NameOf(BusinessPlanComparisonItem.ID)
            ResultsTree.ParentFieldName = NameOf(BusinessPlanComparisonItem.ParentID)
            ResultsTree.PopulateColumns()
            HideColumn(NameOf(BusinessPlanComparisonItem.ID))
            HideColumn(NameOf(BusinessPlanComparisonItem.ParentID))
            HideColumn(NameOf(BusinessPlanComparisonItem.TargetModel))
            ResultsTree.ExpandToLevel(2)
            ResultsTree.BestFitColumns()
            StatusLabel.Text = LastResult.DifferenceCount.ToString("N0") & " differences; " &
                               LastResult.StructuralIssueCount.ToString("N0") & " structural issues; " &
                               LastResult.CheckIssueCount.ToString("N0") & " check issues" &
                               If(LastResult.Truncated, " (detail limit reached)", String.Empty)
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "Compare Business Plans", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
            UpdateActionState()
        End Try
    End Sub

    Private Sub AssessMigration(ByVal sender As Object, ByVal e As EventArgs)
        Dim sourcePlan As BusinessPlanComparisonParticipant = SelectedUpgradeSourceParticipant()
        Dim targetPlan As BusinessPlanComparisonParticipant = SelectedUpgradeTargetParticipant()
        If sourcePlan Is Nothing OrElse targetPlan Is Nothing OrElse
           String.Equals(sourcePlan.Key, targetPlan.Key, StringComparison.OrdinalIgnoreCase) Then
            MessageBox.Show(Me, "Select one source file and one different open blank target file.",
                            "Assess upgrade", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Cursor = Cursors.WaitCursor
        Try
            Dim assessment As BusinessPlanMigrationAssessment =
                BusinessPlanComparisonService.AssessMigration(sourcePlan, targetPlan, AssumptionWorksheetNames)
            Dim text As String = "Common writable assumption cells: " & assessment.CommonInputCellCount.ToString("N0") & Environment.NewLine &
                                 "Structural mismatches: " & assessment.StructuralMismatchCount.ToString("N0") & Environment.NewLine &
                                 "Missing ranges: " & assessment.MissingRangeCount.ToString("N0") & Environment.NewLine & Environment.NewLine &
                                 String.Join(Environment.NewLine, assessment.Lines.Take(30))
            If assessment.Lines.Count > 30 Then text &= Environment.NewLine & "...and " & (assessment.Lines.Count - 30).ToString() & " more."
            text &= Environment.NewLine & Environment.NewLine &
                    "Upgrade will copy all compatible data and common range geometry. Any structural or type problems will be retained in the saved upgrade report."
            MessageBox.Show(Me, text, "Upgrade assessment", MessageBoxButtons.OK,
                            If(assessment.CanPopulate, MessageBoxIcon.Information, MessageBoxIcon.Warning))
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "Assess upgrade", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub PopulateTarget(ByVal sender As Object, ByVal e As EventArgs)
        Dim sourcePlan As BusinessPlanComparisonParticipant = SelectedUpgradeSourceParticipant()
        Dim targetPlan As BusinessPlanComparisonParticipant = SelectedUpgradeTargetParticipant()
        If sourcePlan Is Nothing OrElse targetPlan Is Nothing OrElse
           Not targetPlan.IsOpenSummitModel OrElse
           String.Equals(sourcePlan.Key, targetPlan.Key, StringComparison.OrdinalIgnoreCase) Then Return
        Dim targetModel As ExcelModel = ExcelModels(targetPlan.ModelID)

        Dim destinationPath As String
        Using dialog As New SaveFileDialog With {
            .Title = "Save upgraded Business Plan as",
            .Filter = "Excel Binary Workbook (*.xlsb)|*.xlsb",
            .DefaultExt = "xlsb",
            .AddExtension = True,
            .OverwritePrompt = True,
            .CheckPathExists = True,
            .RestoreDirectory = True,
            .InitialDirectory = IO.Path.GetDirectoryName(targetModel.FileName),
            .FileName = IO.Path.GetFileNameWithoutExtension(targetModel.FileName) &
                        " - upgraded from " & IO.Path.GetFileNameWithoutExtension(sourcePlan.FilePath) & ".xlsb"}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            destinationPath = IO.Path.GetFullPath(dialog.FileName)
        End Using
        If String.Equals(destinationPath, IO.Path.GetFullPath(targetModel.FileName), StringComparison.OrdinalIgnoreCase) Then
            MessageBox.Show(Me, "Choose a different file name. The blank target will not be overwritten.",
                            "Upgrade Business Plan", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim reportPath As String = IO.Path.Combine(
            IO.Path.GetDirectoryName(destinationPath),
            IO.Path.GetFileNameWithoutExtension(destinationPath) & "_report.xlsx")

        Dim confirmation As DialogResult = MessageBox.Show(
            Me,
            "Source:" & Environment.NewLine & sourcePlan.FilePath & Environment.NewLine & Environment.NewLine &
            "Blank target:" & Environment.NewLine & targetModel.FileName & Environment.NewLine & Environment.NewLine &
            "Upgraded copy:" & Environment.NewLine & destinationPath & Environment.NewLine & Environment.NewLine &
            "Upgrade report:" & Environment.NewLine & reportPath & Environment.NewLine & Environment.NewLine &
            "Summit will copy all compatible assumption data, report individual structural or type errors, calculate both files, compare the detailed SOCI and Check Sheet, then save the upgraded target and report.",
            "Upgrade Business Plan",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2)
        If confirmation <> DialogResult.Yes Then Return

        Cursor = Cursors.WaitCursor
        PopulateButton.Enabled = False
        Dim activity As New FormSplashScreen(Me, "Upgrading Business Plan", "Checking assumption structures...")
        Try
            SystemMessageManager.Publish(
                targetPlan.ModelID,
                "Upgrade started from '" & IO.Path.GetFileName(sourcePlan.FilePath) & "'.",
                SystemMessageSeverity.Information,
                "Business Plan Upgrade",
                targetModel.FileName)
            Dim migration As BusinessPlanMigrationResult = BusinessPlanComparisonService.PopulateAssumptions(
                sourcePlan,
                targetPlan,
                AssumptionWorksheetNames,
                Sub(message)
                    activity.Update(message)
                    Application.DoEvents()
                End Sub)

            LastUpgradeResult = migration
            BindResultTree(UpgradeResultsTree, migration.ReportItems)
            UpgradeStatusLabel.Text = migration.ChangedCellCount.ToString("N0") & " cells copied; " &
                                      migration.AddedRecordCount.ToString("N0") & " records added; " &
                                      migration.ErrorCount.ToString("N0") & " upgrade errors; " &
                                      migration.DetailedSOCIDifferenceCount.ToString("N0") & " SOCI differences; " &
                                      migration.CheckIssueCount.ToString("N0") & " check issues"

            targetModel.RecoverySaveAsRequired = True
            activity.Update("Saving the upgraded Business Plan...")
            If Not targetModel.SaveFileAsTo(destinationPath, RequireDifferentPath:=True) Then
                Throw New InvalidOperationException(
                    "Upgrade completed in memory, but the upgraded model was not saved. Use Save As before closing it.")
            End If
            activity.Update("Saving the upgrade report...")
            UpgradeResultsTree.ExportToXlsx(reportPath)

            SystemMessageManager.Publish(
                targetPlan.ModelID,
                "Upgrade and report saved. " & UpgradeStatusLabel.Text & ".",
                If(migration.ErrorCount = 0 AndAlso migration.CheckIssueCount = 0,
                   SystemMessageSeverity.Success,
                   SystemMessageSeverity.Warning),
                "Business Plan Upgrade",
                targetModel.FileName)
            activity.Complete("Upgraded model and report saved.")
            RefreshModels()
            Dim messageLines As IEnumerable(Of String) = migration.Errors.Take(12)
            Dim reportMessage As String = UpgradeStatusLabel.Text & Environment.NewLine & Environment.NewLine &
                                    If(migration.ErrorCount = 0,
                                       "No upgrade errors were reported.",
                                       String.Join(Environment.NewLine, messageLines)) &
                                    Environment.NewLine & Environment.NewLine &
                                    "Saved as:" & Environment.NewLine & targetModel.FileName &
                                    Environment.NewLine & Environment.NewLine &
                                    "Report:" & Environment.NewLine & reportPath
            MessageBox.Show(Me, reportMessage, "Business Plan upgrade report",
                            MessageBoxButtons.OK,
                            If(migration.ErrorCount = 0 AndAlso migration.CheckIssueCount = 0,
                               MessageBoxIcon.Information,
                               MessageBoxIcon.Warning))
        Catch ex As Exception
            activity.Dispose()
            activity = Nothing
            MessageBox.Show(Me, ex.Message, "Upgrade Business Plan", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If activity IsNot Nothing Then activity.Dispose()
            Cursor = Cursors.Default
            UpdateActionState()
        End Try
    End Sub
    Private Sub ExportReport(ByVal sender As Object, ByVal e As EventArgs)
        If LastResult Is Nothing OrElse LastResult.Items.Count = 0 Then Return
        Using dialog As New SaveFileDialog With {
            .Title = "Export business-plan comparison",
            .Filter = "Excel workbook (*.xlsx)|*.xlsx",
            .DefaultExt = "xlsx",
            .AddExtension = True,
            .FileName = "Business Plan Comparison.xlsx"}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            ResultsTree.ExportToXlsx(dialog.FileName)
        End Using
    End Sub

    Private Sub ExportUpgradeReport(ByVal sender As Object, ByVal e As EventArgs)
        If LastUpgradeResult Is Nothing OrElse LastUpgradeResult.ReportItems.Count = 0 Then Return
        Using dialog As New SaveFileDialog With {
            .Title = "Export business-plan upgrade report",
            .Filter = "Excel workbook (*.xlsx)|*.xlsx",
            .DefaultExt = "xlsx",
            .AddExtension = True,
            .FileName = "Business Plan Upgrade_report.xlsx"}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            UpgradeResultsTree.ExportToXlsx(dialog.FileName)
        End Using
    End Sub

    Private Shared Function SelectedLookupKey(ByVal lookup As LookUpEdit) As String
        If lookup Is Nothing OrElse lookup.EditValue Is Nothing OrElse Convert.IsDBNull(lookup.EditValue) Then Return String.Empty
        Return Convert.ToString(lookup.EditValue)
    End Function

    Private Function SelectedLookupParticipant(ByVal lookup As LookUpEdit) As BusinessPlanComparisonParticipant
        Dim key As String = SelectedLookupKey(lookup)
        Dim choice As ModelChoice = Models.FirstOrDefault(
            Function(item) String.Equals(item.Key, key, StringComparison.OrdinalIgnoreCase))
        Return If(choice Is Nothing, Nothing, choice.Participant)
    End Function

    Private Function SelectedUpgradeSourceParticipant() As BusinessPlanComparisonParticipant
        Return SelectedLookupParticipant(UpgradeSourceEdit)
    End Function

    Private Function SelectedUpgradeTargetParticipant() As BusinessPlanComparisonParticipant
        Return SelectedLookupParticipant(UpgradeTargetEdit)
    End Function
    Private Function SelectedBaseKey() As String
        If BaseModelEdit.EditValue Is Nothing OrElse Convert.IsDBNull(BaseModelEdit.EditValue) Then Return String.Empty
        Return Convert.ToString(BaseModelEdit.EditValue)
    End Function

    Private Function SelectedBaseParticipant() As BusinessPlanComparisonParticipant
        Dim key As String = SelectedBaseKey()
        Dim choice As ModelChoice = Models.FirstOrDefault(Function(item) String.Equals(item.Key, key, StringComparison.OrdinalIgnoreCase))
        Return If(choice Is Nothing, Nothing, choice.Participant)
    End Function

    Private Function SelectedComparisonKeys() As List(Of String)
        Return SelectedComparisonParticipants().Select(Function(item) item.Key).ToList()
    End Function

    Private Function SelectedComparisonParticipants() As List(Of BusinessPlanComparisonParticipant)
        Dim result As New List(Of BusinessPlanComparisonParticipant)()
        For Each item As CheckedListBoxItem In ComparedModelsList.Items
            If item.CheckState <> CheckState.Checked Then Continue For
            Dim choice As ModelChoice = TryCast(item.Value, ModelChoice)
            If choice IsNot Nothing AndAlso choice.Participant IsNot Nothing Then result.Add(choice.Participant)
        Next
        Return result
    End Function

    Private Sub UpdateActionState()
        Dim basePlan As BusinessPlanComparisonParticipant = SelectedBaseParticipant()
        Dim selected As List(Of BusinessPlanComparisonParticipant) = SelectedComparisonParticipants().
            Where(Function(item) basePlan Is Nothing OrElse Not String.Equals(item.Key, basePlan.Key, StringComparison.OrdinalIgnoreCase)).ToList()
        RunButton.Enabled = basePlan IsNot Nothing AndAlso selected.Count > 0

        Dim upgradeSource As BusinessPlanComparisonParticipant = SelectedUpgradeSourceParticipant()
        Dim upgradeTarget As BusinessPlanComparisonParticipant = SelectedUpgradeTargetParticipant()
        Dim canUpgrade As Boolean = upgradeSource IsNot Nothing AndAlso
                                    upgradeTarget IsNot Nothing AndAlso
                                    upgradeTarget.IsOpenSummitModel AndAlso
                                    Not String.Equals(upgradeSource.Key, upgradeTarget.Key, StringComparison.OrdinalIgnoreCase)
        AssessMigrationButton.Enabled = canUpgrade
        PopulateButton.Enabled = canUpgrade
    End Sub
    Protected Overrides Sub OnFormClosed(ByVal e As FormClosedEventArgs)
        For Each workbook As Workbook In ExternalWorkbooks.Values
            workbook.Dispose()
        Next
        ExternalWorkbooks.Clear()
        MyBase.OnFormClosed(e)
    End Sub

    Private Sub BindResultTree(ByVal tree As TreeList,
                               ByVal items As IEnumerable(Of BusinessPlanComparisonItem))
        tree.DataSource = items.ToList()
        tree.KeyFieldName = NameOf(BusinessPlanComparisonItem.ID)
        tree.ParentFieldName = NameOf(BusinessPlanComparisonItem.ParentID)
        tree.PopulateColumns()
        HideColumn(tree, NameOf(BusinessPlanComparisonItem.ID))
        HideColumn(tree, NameOf(BusinessPlanComparisonItem.ParentID))
        HideColumn(tree, NameOf(BusinessPlanComparisonItem.TargetModel))
        tree.ExpandToLevel(2)
        tree.BestFitColumns()
    End Sub

    Private Sub HideColumn(ByVal fieldName As String)
        HideColumn(ResultsTree, fieldName)
    End Sub

    Private Shared Sub HideColumn(ByVal tree As TreeList,
                                  ByVal fieldName As String)
        Dim column As TreeListColumn = tree.Columns.ColumnByFieldName(fieldName)
        If column IsNot Nothing Then column.Visible = False
    End Sub
End Class
