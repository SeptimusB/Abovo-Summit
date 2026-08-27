Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing.Printing
Imports System.IO
Imports System.Linq
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Grid

''' <summary>
''' Separate, model-scoped Excel export workspace. Each selected DIT grid is
''' frozen to XLSX when added; publishing combines snapshots as worksheets.
''' </summary>
Public NotInheritable Class DITExcelExportManager
    Private ReadOnly ModelID As Integer
    Private Workspace As DITExcelExportWorkspace

    Public Sub New(ByVal setModelID As Integer)
        ModelID = setModelID
    End Sub

    Public Sub ShowForDIT(ByVal owner As DataInterfaceTemplate)
        If owner Is Nothing OrElse owner.IsDisposed Then Return
        Dim candidates As List(Of DITExportCandidate) = owner.GetExportCandidates()
        If candidates.Count = 0 Then
            XtraMessageBox.Show(owner,
                "This interface does not currently contain a grid that can be exported to Excel.",
                "Excel export", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        If Workspace Is Nothing OrElse Workspace.IsDisposed Then
            Workspace = New DITExcelExportWorkspace(ModelID)
        End If
        Workspace.SetSource(owner, candidates)
        Workspace.ShowForUser(owner)
    End Sub

    Public Sub CloseForModel()
        If Workspace IsNot Nothing Then Workspace.CloseForModel()
        Workspace = Nothing
    End Sub
End Class

Friend NotInheritable Class DITExcelExportWorkspace
    Inherits XtraForm

    Private ReadOnly TemporaryFolder As String
    Private ReadOnly StagedItems As New BindingList(Of DITExcelStagedItem)()
    Private ReadOnly Root As New TableLayoutPanel()
    Private ReadOnly CandidateLabel As New LabelControl()
    Private ReadOnly CandidateList As New CheckedListBoxControl()
    Private ReadOnly AddButton As New SimpleButton()
    Private ReadOnly QueueGrid As New GridControl()
    Private ReadOnly QueueView As New GridView()
    Private ReadOnly Toolbar As New FlowLayoutPanel()
    Private ReadOnly RemoveButton As New SimpleButton()
    Private ReadOnly MoveUpButton As New SimpleButton()
    Private ReadOnly MoveDownButton As New SimpleButton()
    Private ReadOnly ClearButton As New SimpleButton()
    Private ReadOnly PublishButton As New SimpleButton()
    Private ReadOnly CloseButton As New SimpleButton()
    Private ReadOnly StatusLabel As New LabelControl()
    Private ClosingForModel As Boolean

    Public Sub New(ByVal setModelID As Integer)
        TemporaryFolder = Path.Combine(Path.GetTempPath(), "AbovoSummit", "DITExcel",
            setModelID.ToString() & "_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(TemporaryFolder)
        Text = "Excel Export Workspace"
        StartPosition = FormStartPosition.CenterParent
        MinimumSize = New Size(800, 560)
        Size = New Size(1060, 720)
        BuildSurface()
    End Sub

    Private Sub BuildSurface()
        Root.Dock = DockStyle.Fill
        Root.Padding = New Padding(12)
        Root.ColumnCount = 1
        Root.RowCount = 7
        Root.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0F))
        Root.RowStyles.Add(New RowStyle(SizeType.Percent, 34.0F))
        Root.RowStyles.Add(New RowStyle(SizeType.Absolute, 42.0F))
        Root.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0F))
        Root.RowStyles.Add(New RowStyle(SizeType.Percent, 66.0F))
        Root.RowStyles.Add(New RowStyle(SizeType.Absolute, 50.0F))
        Root.RowStyles.Add(New RowStyle(SizeType.Absolute, 28.0F))
        Controls.Add(Root)

        CandidateLabel.Text = "Available grids in this interface"
        CandidateLabel.Appearance.Font = New Font(CandidateLabel.Font, FontStyle.Bold)
        CandidateLabel.Dock = DockStyle.Fill
        Root.Controls.Add(CandidateLabel, 0, 0)
        CandidateList.Dock = DockStyle.Fill
        CandidateList.CheckOnClick = True
        Root.Controls.Add(CandidateList, 0, 1)

        AddButton.Text = "Add selected snapshots to Excel"
        AddButton.Width = 230
        AddButton.Height = 30
        AddHandler AddButton.Click, AddressOf AddButton_Click
        Dim addPanel As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill, .FlowDirection = FlowDirection.LeftToRight,
            .Padding = New Padding(0, 5, 0, 3)}
        addPanel.Controls.Add(AddButton)
        Root.Controls.Add(addPanel, 0, 2)

        Dim queueLabel As New LabelControl With {
            .Text = "In-process workbook contents (one or more worksheets per snapshot)",
            .Dock = DockStyle.Fill}
        queueLabel.Appearance.Font = New Font(queueLabel.Font, FontStyle.Bold)
        Root.Controls.Add(queueLabel, 0, 3)

        QueueGrid.Dock = DockStyle.Fill
        QueueGrid.MainView = QueueView
        QueueGrid.ViewCollection.Add(QueueView)
        QueueGrid.DataSource = StagedItems
        QueueView.OptionsBehavior.Editable = False
        QueueView.OptionsSelection.MultiSelect = True
        QueueView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.RowSelect
        QueueView.OptionsView.ShowGroupPanel = False
        QueueView.OptionsView.ColumnAutoWidth = True
        QueueView.OptionsClipboard.CopyColumnHeaders = DefaultBoolean.True
        QueueView.PopulateColumns()
        Root.Controls.Add(QueueGrid, 0, 4)

        Toolbar.Dock = DockStyle.Fill
        Toolbar.FlowDirection = FlowDirection.LeftToRight
        Toolbar.WrapContents = False
        ConfigureButton(RemoveButton, "Remove", AddressOf RemoveButton_Click)
        ConfigureButton(MoveUpButton, "Move up", AddressOf MoveUpButton_Click)
        ConfigureButton(MoveDownButton, "Move down", AddressOf MoveDownButton_Click)
        ConfigureButton(ClearButton, "Clear", AddressOf ClearButton_Click)
        ConfigureButton(PublishButton, "Publish Excel...", AddressOf PublishButton_Click)
        ConfigureButton(CloseButton, "Close", AddressOf CloseButton_Click)
        Toolbar.Controls.AddRange(New Control() {
            RemoveButton, MoveUpButton, MoveDownButton, ClearButton, PublishButton, CloseButton})
        Root.Controls.Add(Toolbar, 0, 5)
        StatusLabel.Dock = DockStyle.Fill
        Root.Controls.Add(StatusLabel, 0, 6)
        RefreshStatus()
        AddHandler FormClosing, AddressOf Workspace_FormClosing
    End Sub

    Private Shared Sub ConfigureButton(ByVal button As SimpleButton,
                                       ByVal caption As String,
                                       ByVal handler As EventHandler)
        button.Text = caption
        button.AutoSize = True
        button.MinimumSize = New Size(88, 30)
        AddHandler button.Click, handler
    End Sub

    Public Sub SetSource(ByVal owner As DataInterfaceTemplate,
                         ByVal candidates As IEnumerable(Of DITExportCandidate))
        CandidateList.BeginUpdate()
        Try
            CandidateList.Items.Clear()
            For Each candidate As DITExportCandidate In candidates
                Dim state As CheckState = If(candidate.IsCurrentSection,
                    CheckState.Checked, CheckState.Unchecked)
                CandidateList.Items.Add(New CheckedListBoxItem(candidate, candidate.Title, state))
            Next
        Finally
            CandidateList.EndUpdate()
        End Try
        CandidateLabel.Text = "Available grids in " & owner.Text
        RefreshStatus()
    End Sub

    Public Sub ShowForUser(ByVal owner As IWin32Window)
        If Visible Then
            WindowState = FormWindowState.Normal
            BringToFront()
            Activate()
        Else
            Show(owner)
        End If
    End Sub

    Private Sub AddButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim selected As New List(Of DITExportCandidate)()
        For index As Integer = 0 To CandidateList.ItemCount - 1
            If CandidateList.GetItemChecked(index) Then
                Dim candidate As DITExportCandidate =
                    TryCast(CandidateList.Items(index).Value, DITExportCandidate)
                If candidate IsNot Nothing Then selected.Add(candidate)
            End If
        Next

        If selected.Count = 0 Then
            XtraMessageBox.Show(Me, "Select at least one grid to add.", "Excel export")
            Return
        End If

        Cursor = Cursors.WaitCursor
        Try
            For Each candidate As DITExportCandidate In selected
                Dim snapshotPath As String = Path.Combine(TemporaryFolder,
                    Guid.NewGuid().ToString("N") & ".xlsx")
                CreateSnapshot(candidate, snapshotPath)
                StagedItems.Add(New DITExcelStagedItem With {
                    .Sequence = StagedItems.Count + 1,
                    .Title = candidate.Title,
                    .InterfaceName = candidate.InterfaceName,
                    .SectionName = candidate.SectionName,
                    .Captured = DateTime.Now,
                    .SnapshotPath = snapshotPath})
            Next
            RefreshQueue()
        Catch ex As Exception
            XtraMessageBox.Show(Me,
                "The selected grid could not be added to the Excel workspace." &
                Environment.NewLine & Environment.NewLine & ex.Message,
                "Excel export", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Shared Sub CreateSnapshot(ByVal candidate As DITExportCandidate,
                                      ByVal snapshotPath As String)
        If candidate.PrintableComponent Is Nothing Then
            Throw New InvalidOperationException("The source control is no longer available.")
        End If
        Using printingSystem As New DevExpress.XtraPrinting.PrintingSystem()
            Using link As New DevExpress.XtraPrinting.PrintableComponentLink(printingSystem)
                link.Component = candidate.PrintableComponent
                link.PaperKind = PaperKind.A4
                link.Landscape = True
                link.Margins = New System.Drawing.Printing.Margins(35, 35, 45, 40)
                Dim title As String = candidate.Title
                AddHandler link.CreateReportHeaderArea,
                    Sub(sender As Object, args As DevExpress.XtraPrinting.CreateAreaEventArgs)
                        args.Graph.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
                        args.Graph.StringFormat = New DevExpress.XtraPrinting.BrickStringFormat(
                            StringAlignment.Near, StringAlignment.Center)
                        args.Graph.DrawString(title, Color.FromArgb(0, 79, 158),
                            New RectangleF(0.0F, 0.0F, args.Graph.ClientPageSize.Width, 34.0F),
                            DevExpress.XtraPrinting.BorderSide.Bottom)
                    End Sub
                link.CreateDocument()
                link.ExportToXlsx(snapshotPath)
            End Using
        End Using
    End Sub

    Private Sub RemoveButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim removeItems As New List(Of DITExcelStagedItem)()
        For Each rowHandle As Integer In QueueView.GetSelectedRows()
            Dim item As DITExcelStagedItem = TryCast(QueueView.GetRow(rowHandle), DITExcelStagedItem)
            If item IsNot Nothing Then removeItems.Add(item)
        Next
        For Each item As DITExcelStagedItem In removeItems
            DeleteSnapshot(item)
            StagedItems.Remove(item)
        Next
        RefreshQueue()
    End Sub

    Private Sub MoveUpButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        MoveFocusedItem(-1)
    End Sub

    Private Sub MoveDownButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        MoveFocusedItem(1)
    End Sub

    Private Sub MoveFocusedItem(ByVal offset As Integer)
        Dim item As DITExcelStagedItem = TryCast(QueueView.GetFocusedRow(), DITExcelStagedItem)
        If item Is Nothing Then Return
        Dim oldIndex As Integer = StagedItems.IndexOf(item)
        Dim newIndex As Integer = oldIndex + offset
        If newIndex < 0 OrElse newIndex >= StagedItems.Count Then Return
        StagedItems.RaiseListChangedEvents = False
        Try
            StagedItems.RemoveAt(oldIndex)
            StagedItems.Insert(newIndex, item)
        Finally
            StagedItems.RaiseListChangedEvents = True
        End Try
        RefreshQueue()
        QueueView.FocusedRowHandle = newIndex
    End Sub

    Private Sub ClearButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        If StagedItems.Count = 0 Then Return
        If XtraMessageBox.Show(Me, "Remove all staged Excel snapshots?", "Excel export",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        ClearStagedItems()
    End Sub

    Private Sub PublishButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        If StagedItems.Count = 0 Then
            XtraMessageBox.Show(Me, "Add at least one snapshot before publishing.", "Excel export")
            Return
        End If
        Using dialog As New SaveFileDialog With {
            .AddExtension = True, .DefaultExt = "xlsx",
            .Filter = "Excel workbooks (*.xlsx)|*.xlsx",
            .OverwritePrompt = True, .Title = "Publish staged Excel workbook"}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            Cursor = Cursors.WaitCursor
            Try
                PublishWorkbook(dialog.FileName)
                If XtraMessageBox.Show(Me,
                        "The Excel workbook was published successfully." &
                        Environment.NewLine & "Open it now?", "Excel export",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                    Process.Start(New ProcessStartInfo(dialog.FileName) With {.UseShellExecute = True})
                End If
            Catch ex As Exception
                XtraMessageBox.Show(Me,
                    "The Excel workbook could not be published." & Environment.NewLine &
                    Environment.NewLine & ex.Message, "Excel export",
                    MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Cursor = Cursors.Default
            End Try
        End Using
    End Sub

    Private Sub PublishWorkbook(ByVal outputPath As String)
        Using targetWorkbook As New Workbook()
            targetWorkbook.CreateNewDocument()
            Dim firstTargetSheet As Boolean = True
            Dim usedNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each item As DITExcelStagedItem In StagedItems
                Using sourceWorkbook As New Workbook()
                    sourceWorkbook.LoadDocument(item.SnapshotPath, DocumentFormat.Xlsx)
                    For sourceIndex As Integer = 0 To sourceWorkbook.Worksheets.Count - 1
                        Dim targetSheet As Worksheet
                        If firstTargetSheet Then
                            targetSheet = targetWorkbook.Worksheets(0)
                            firstTargetSheet = False
                        Else
                            targetSheet = targetWorkbook.Worksheets.Add()
                        End If
                        targetSheet.CopyFrom(sourceWorkbook.Worksheets(sourceIndex))
                        Dim requestedName As String = item.Title
                        If sourceWorkbook.Worksheets.Count > 1 Then
                            requestedName &= " " & (sourceIndex + 1).ToString()
                        End If
                        targetSheet.Name = MakeUniqueWorksheetName(requestedName, usedNames)
                    Next
                End Using
            Next
            targetWorkbook.SaveDocument(outputPath, DocumentFormat.Xlsx)
        End Using
    End Sub

    Private Shared Function MakeUniqueWorksheetName(
        ByVal requestedName As String,
        ByVal usedNames As HashSet(Of String)) As String
        Dim candidate As String = If(requestedName, String.Empty).Trim()
        For Each invalidCharacter As Char In New Char() {
            ":"c, "\"c, "/"c, "?"c, "*"c, "["c, "]"c}
            candidate = candidate.Replace(invalidCharacter, "-"c)
        Next
        candidate = candidate.Trim("'"c, " "c)
        If candidate.Length = 0 Then candidate = "Export"
        If candidate.Length > 31 Then candidate = candidate.Substring(0, 31).Trim()

        Dim baseName As String = candidate
        Dim suffixIndex As Integer = 2
        While usedNames.Contains(candidate)
            Dim suffix As String = " (" & suffixIndex.ToString() & ")"
            candidate = baseName.Substring(
                0, Math.Min(baseName.Length, 31 - suffix.Length)).Trim() & suffix
            suffixIndex += 1
        End While
        usedNames.Add(candidate)
        Return candidate
    End Function

    Private Sub CloseButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Hide()
    End Sub

    Private Sub Workspace_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs)
        If Not ClosingForModel Then
            e.Cancel = True
            Hide()
        End If
    End Sub

    Public Sub CloseForModel()
        ClosingForModel = True
        ClearStagedItems()
        Try
            If Directory.Exists(TemporaryFolder) Then Directory.Delete(TemporaryFolder, True)
        Catch
            'Temporary snapshots are best-effort cleanup only.
        End Try
        Close()
        Dispose()
    End Sub

    Private Sub ClearStagedItems()
        For Each item As DITExcelStagedItem In StagedItems.ToList()
            DeleteSnapshot(item)
        Next
        StagedItems.Clear()
        RefreshQueue()
    End Sub

    Private Shared Sub DeleteSnapshot(ByVal item As DITExcelStagedItem)
        Try
            If item IsNot Nothing AndAlso File.Exists(item.SnapshotPath) Then
                File.Delete(item.SnapshotPath)
            End If
        Catch
            'Temporary snapshots are best-effort cleanup only.
        End Try
    End Sub

    Private Sub RefreshQueue()
        For index As Integer = 0 To StagedItems.Count - 1
            StagedItems(index).Sequence = index + 1
        Next
        StagedItems.ResetBindings()
        QueueView.RefreshData()
        RefreshStatus()
    End Sub

    Private Sub RefreshStatus()
        StatusLabel.Text = StagedItems.Count.ToString() &
            If(StagedItems.Count = 1, " snapshot staged", " snapshots staged") &
            " | Excel and PDF workspaces are independent."
        PublishButton.Enabled = StagedItems.Count > 0
        RemoveButton.Enabled = StagedItems.Count > 0
        ClearButton.Enabled = StagedItems.Count > 0
        MoveUpButton.Enabled = StagedItems.Count > 1
        MoveDownButton.Enabled = StagedItems.Count > 1
    End Sub
End Class

Friend NotInheritable Class DITExcelStagedItem
    Public Property Sequence As Integer
    Public Property Title As String
    Public Property InterfaceName As String
    Public Property SectionName As String
    Public Property Captured As DateTime

    <Browsable(False)>
    Public Property SnapshotPath As String
End Class
