Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing.Printing
Imports System.IO
Imports System.Linq
Imports DevExpress.Pdf
Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Grid

''' <summary>A printable control exposed by a DataInterfaceTemplate section.</summary>
Public NotInheritable Class DITPdfExportCandidate

    Public Sub New(ByVal title As String,
                   ByVal interfaceName As String,
                   ByVal sectionName As String,
                   ByVal printableComponent As DevExpress.XtraPrinting.IPrintable,
                   ByVal isCurrentSection As Boolean)
        Me.Title = title
        Me.InterfaceName = interfaceName
        Me.SectionName = sectionName
        Me.PrintableComponent = printableComponent
        Me.IsCurrentSection = isCurrentSection
    End Sub

    Public ReadOnly Property Title As String
    Public ReadOnly Property InterfaceName As String
    Public ReadOnly Property SectionName As String
    Public ReadOnly Property PrintableComponent As DevExpress.XtraPrinting.IPrintable
    Public ReadOnly Property IsCurrentSection As Boolean
End Class

''' <summary>
''' Model-scoped PDF workspace. Items are rendered to temporary PDF snapshots
''' when added, so subsequent calculations cannot alter an item already staged.
''' </summary>
Public NotInheritable Class DITPdfExportManager

    Private ReadOnly ModelID As Integer
    Private Workspace As DITPdfExportWorkspace

    Public Sub New(ByVal setModelID As Integer)
        ModelID = setModelID
    End Sub

    Public Sub ShowForDIT(ByVal owner As DataInterfaceTemplate)
        If owner Is Nothing OrElse owner.IsDisposed Then Return

        Dim candidates As List(Of DITPdfExportCandidate) = owner.GetPdfExportCandidates()
        If candidates.Count = 0 Then
            XtraMessageBox.Show(owner,
                "This interface does not currently contain a grid that can be exported to PDF.",
                "PDF export", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If Workspace Is Nothing OrElse Workspace.IsDisposed Then
            Workspace = New DITPdfExportWorkspace(ModelID)
        End If

        Workspace.SetSource(owner, candidates)
        Workspace.ShowForUser(owner)
    End Sub

    Public Sub CloseForModel()
        If Workspace IsNot Nothing Then Workspace.CloseForModel()
        Workspace = Nothing
    End Sub
End Class

Friend NotInheritable Class DITPdfExportWorkspace
    Inherits XtraForm

    Private ReadOnly TemporaryFolder As String
    Private ReadOnly StagedItems As New BindingList(Of DITPdfStagedItem)()
    Private ReadOnly Root As New TableLayoutPanel()
    Private ReadOnly CandidateLabel As New LabelControl()
    Private ReadOnly CandidateList As New CheckedListBoxControl()
    Private ReadOnly AddButton As New SimpleButton()
    Private ReadOnly QueueLabel As New LabelControl()
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
        TemporaryFolder = Path.Combine(Path.GetTempPath(), "AbovoSummit", "DITPdf",
            setModelID.ToString() & "_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(TemporaryFolder)

        Text = "PDF Export Workspace"
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

        AddButton.Text = "Add selected snapshots to PDF"
        AddButton.Width = 230
        AddButton.Height = 30
        AddHandler AddButton.Click, AddressOf AddButton_Click
        Dim addPanel As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.LeftToRight,
            .Padding = New Padding(0, 5, 0, 3)}
        addPanel.Controls.Add(AddButton)
        Root.Controls.Add(addPanel, 0, 2)

        QueueLabel.Text = "In-process PDF contents (snapshots are retained in this order)"
        QueueLabel.Appearance.Font = New Font(QueueLabel.Font, FontStyle.Bold)
        QueueLabel.Dock = DockStyle.Fill
        Root.Controls.Add(QueueLabel, 0, 3)

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
        Toolbar.Padding = New Padding(0, 8, 0, 4)
        ConfigureButton(RemoveButton, "Remove", AddressOf RemoveButton_Click)
        ConfigureButton(MoveUpButton, "Move up", AddressOf MoveUpButton_Click)
        ConfigureButton(MoveDownButton, "Move down", AddressOf MoveDownButton_Click)
        ConfigureButton(ClearButton, "Clear", AddressOf ClearButton_Click)
        ConfigureButton(PublishButton, "Publish PDF...", AddressOf PublishButton_Click)
        ConfigureButton(CloseButton, "Close", AddressOf CloseButton_Click)
        Toolbar.Controls.AddRange(New Control() {
            RemoveButton, MoveUpButton, MoveDownButton, ClearButton, PublishButton, CloseButton})
        Root.Controls.Add(Toolbar, 0, 5)

        StatusLabel.Dock = DockStyle.Fill
        StatusLabel.Appearance.ForeColor = Color.FromArgb(32, 58, 89)
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
                         ByVal candidates As IEnumerable(Of DITPdfExportCandidate))
        CandidateList.BeginUpdate()
        Try
            CandidateList.Items.Clear()
            For Each candidate As DITPdfExportCandidate In candidates
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
        Dim selected As New List(Of DITPdfExportCandidate)()
        For index As Integer = 0 To CandidateList.ItemCount - 1
            If CandidateList.GetItemChecked(index) Then
                Dim candidate As DITPdfExportCandidate =
                    TryCast(CandidateList.Items(index).Value, DITPdfExportCandidate)
                If candidate IsNot Nothing Then selected.Add(candidate)
            End If
        Next

        If selected.Count = 0 Then
            XtraMessageBox.Show(Me, "Select at least one grid to add.", "PDF export")
            Return
        End If

        Cursor = Cursors.WaitCursor
        Try
            For Each candidate As DITPdfExportCandidate In selected
                Dim snapshotPath As String = Path.Combine(TemporaryFolder,
                    Guid.NewGuid().ToString("N") & ".pdf")
                CreateSnapshot(candidate, snapshotPath)
                StagedItems.Add(New DITPdfStagedItem With {
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
                "The selected grid could not be added to the PDF workspace." &
                Environment.NewLine & Environment.NewLine & ex.Message,
                "PDF export", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Shared Sub CreateSnapshot(ByVal candidate As DITPdfExportCandidate,
                                      ByVal snapshotPath As String)
        If candidate.PrintableComponent Is Nothing Then
            Throw New InvalidOperationException("The source control is no longer available.")
        End If

        Using printingSystem As New DevExpress.XtraPrinting.PrintingSystem()
            Using link As New DevExpress.XtraPrinting.PrintableComponentLink(printingSystem)
                link.Component = candidate.PrintableComponent
                link.PaperKind = PaperKind.A4
                link.Landscape = True
                link.Margins = New Margins(35, 35, 45, 40)

                Dim title As String = candidate.Title
                AddHandler link.CreateReportHeaderArea,
                    Sub(sender As Object, args As DevExpress.XtraPrinting.CreateAreaEventArgs)
                        args.Graph.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
                        args.Graph.StringFormat =
                            New DevExpress.XtraPrinting.BrickStringFormat(
                                StringAlignment.Near, StringAlignment.Center)
                        args.Graph.DrawString(title, Color.FromArgb(0, 79, 158),
                            New RectangleF(0.0F, 0.0F, args.Graph.ClientPageSize.Width, 34.0F),
                            DevExpress.XtraPrinting.BorderSide.Bottom)
                    End Sub

                link.CreateDocument()
                link.ExportToPdf(snapshotPath)
            End Using
        End Using
    End Sub

    Private Sub RemoveButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim removeItems As New List(Of DITPdfStagedItem)()
        For Each rowHandle As Integer In QueueView.GetSelectedRows()
            Dim item As DITPdfStagedItem = TryCast(QueueView.GetRow(rowHandle), DITPdfStagedItem)
            If item IsNot Nothing Then removeItems.Add(item)
        Next
        For Each item As DITPdfStagedItem In removeItems
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
        Dim item As DITPdfStagedItem = TryCast(QueueView.GetFocusedRow(), DITPdfStagedItem)
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
        If XtraMessageBox.Show(Me, "Remove all staged PDF snapshots?", "PDF export",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        ClearStagedItems()
    End Sub

    Private Sub PublishButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        If StagedItems.Count = 0 Then
            XtraMessageBox.Show(Me, "Add at least one snapshot before publishing.", "PDF export")
            Return
        End If

        Using dialog As New SaveFileDialog With {
            .AddExtension = True,
            .DefaultExt = "pdf",
            .Filter = "PDF documents (*.pdf)|*.pdf",
            .OverwritePrompt = True,
            .Title = "Publish staged PDF"}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return

            Cursor = Cursors.WaitCursor
            Try
                If StagedItems.Count = 1 Then
                    File.Copy(StagedItems(0).SnapshotPath, dialog.FileName, True)
                Else
                    Using processor As New PdfDocumentProcessor()
                        processor.CreateEmptyDocument()
                        For Each item As DITPdfStagedItem In StagedItems
                            processor.AppendDocument(item.SnapshotPath)
                        Next
                        processor.SaveDocument(dialog.FileName)
                    End Using
                End If

                If XtraMessageBox.Show(Me,
                        "The PDF was published successfully." & Environment.NewLine & "Open it now?",
                        "PDF export", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information) = DialogResult.Yes Then
                    Process.Start(New ProcessStartInfo(dialog.FileName) With {.UseShellExecute = True})
                End If
            Catch ex As Exception
                XtraMessageBox.Show(Me,
                    "The PDF could not be published." & Environment.NewLine &
                    Environment.NewLine & ex.Message,
                    "PDF export", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Cursor = Cursors.Default
            End Try
        End Using
    End Sub

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
        For Each item As DITPdfStagedItem In StagedItems.ToList()
            DeleteSnapshot(item)
        Next
        StagedItems.Clear()
        RefreshQueue()
    End Sub

    Private Shared Sub DeleteSnapshot(ByVal item As DITPdfStagedItem)
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
            " | Items are frozen when added and retained until cleared or the workbook closes."
        PublishButton.Enabled = StagedItems.Count > 0
        RemoveButton.Enabled = StagedItems.Count > 0
        ClearButton.Enabled = StagedItems.Count > 0
        MoveUpButton.Enabled = StagedItems.Count > 1
        MoveDownButton.Enabled = StagedItems.Count > 1
    End Sub
End Class

Friend NotInheritable Class DITPdfStagedItem
    Public Property Sequence As Integer
    Public Property Title As String
    Public Property InterfaceName As String
    Public Property SectionName As String
    Public Property Captured As DateTime

    <Browsable(False)>
    Public Property SnapshotPath As String
End Class
