Imports System.ComponentModel
Imports System.Configuration
Imports System.IO
Imports Abovo
Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraTab

Public Class ModelManagerForm
    Inherits XtraForm

    Private Definition As ModelManagerDefinition
    Private ReadOnly NameEdit As New TextEdit()
    Private ReadOnly ClientEdit As New TextEdit()
    Private ReadOnly FamilyEdit As New TextEdit()
    Private ReadOnly VersionEdit As New TextEdit()
    Private ReadOnly GenericEdit As New CheckEdit With {.Text = "Generic (not client-specific)", .Dock = DockStyle.Fill}
    Private ReadOnly RoleEdit As New ComboBoxEdit()
    Private ReadOnly NotesEdit As New MemoEdit()
    Private ReadOnly Status As New LabelControl With {.AutoSizeMode = LabelAutoSizeMode.Vertical, .Dock = DockStyle.Bottom, .Padding = New Padding(8)}
    Private ReadOnly RulesGrid As New GridControl With {.Dock = DockStyle.Fill}
    Private ReadOnly RulesView As New GridView()
    Private ReadOnly EvidenceGrid As New GridControl With {.Dock = DockStyle.Fill}
    Private ReadOnly EvidenceView As New GridView()
    Private ReadOnly EvidencePaths As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    Private IsChanged As Boolean
    Private LoadingDefinition As Boolean

    Public Sub New()
        Text = "Model Manager — definition trial"
        Size = New Size(1300, 800)
        MinimumSize = New Size(950, 600)
        StartPosition = FormStartPosition.CenterParent
        BuildInterface()
        LoadDefinition(New ModelManagerDefinition())
        AddHandler FormClosing, Sub(sender, e)
                                    If IsChanged AndAlso Not ConfirmDiscard() Then e.Cancel = True
                                End Sub
    End Sub

    Public Shared Function Authorize(owner As IWin32Window) As Boolean
        If Not ModelManagerAccess.PasswordRequired Then Return True
        Dim verifier = ConfigurationManager.AppSettings("ModelManager.PasswordVerifier")
        If String.IsNullOrWhiteSpace(verifier) Then
            XtraMessageBox.Show(owner, "Model Manager is locked. An administrator must configure its password verifier before live use.", "Model Manager")
            Return False
        End If
        Using dialog As New XtraForm With {.Text = "Model Manager access", .Size = New Size(390, 150), .StartPosition = FormStartPosition.CenterParent, .FormBorderStyle = FormBorderStyle.FixedDialog, .MaximizeBox = False, .MinimizeBox = False}
            Dim password As New TextEdit With {.Dock = DockStyle.Top}
            password.Properties.UseSystemPasswordChar = True
            password.Properties.MaxLength = 160
            Dim ok As New SimpleButton With {.Text = "Unlock", .Dock = DockStyle.Bottom, .DialogResult = DialogResult.OK}
            dialog.Controls.Add(password)
            dialog.Controls.Add(ok)
            dialog.AcceptButton = ok
            If dialog.ShowDialog(owner) <> DialogResult.OK Then Return False
            Dim accepted = ModelManagerAccess.Verify(password.Text, verifier)
            password.Text = ""
            If Not accepted Then XtraMessageBox.Show(owner, "Password not accepted.", "Model Manager")
            Return accepted
        End Using
    End Function

    Private Sub BuildInterface()
        Dim toolbar As New FlowLayoutPanel With {.Dock = DockStyle.Top, .Height = 78, .AutoScroll = True}
        AddButton(toolbar, "New definition", Sub()
                                                 If ConfirmDiscard() Then LoadDefinition(New ModelManagerDefinition())
                                             End Sub)
        AddButton(toolbar, "Saved definitions...", AddressOf OpenSaved)
        AddButton(toolbar, "Import XML...", AddressOf OpenXml)
        AddButton(toolbar, "Read from XLSB...", AddressOf OpenWorkbook)
        AddButton(toolbar, "Save revision", AddressOf SaveRevision)
        AddButton(toolbar, "Export XML...", AddressOf ExportXml)
        AddButton(toolbar, "Save XLSB copy with definition...", AddressOf SaveWorkbook)
        Dim metadata As New TableLayoutPanel With {.Dock = DockStyle.Top, .Height = 200, .ColumnCount = 4, .RowCount = 4, .Padding = New Padding(8)}
        metadata.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
        metadata.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
        metadata.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
        metadata.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        metadata.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
        metadata.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        metadata.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
        metadata.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        AddField(metadata, "Definition", NameEdit, 0, 0)
        AddField(metadata, "Client", ClientEdit, 2, 0)
        AddField(metadata, "Model family", FamilyEdit, 0, 1)
        AddField(metadata, "Base version", VersionEdit, 2, 1)
        RoleEdit.Properties.Items.AddRange([Enum].GetNames(GetType(ManagedModelRole)))
        RoleEdit.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        AddField(metadata, "File role", RoleEdit, 0, 2)
        metadata.Controls.Add(GenericEdit, 3, 2)
        AddField(metadata, "Notes", NotesEdit, 0, 3)
        metadata.SetColumnSpan(NotesEdit, 3)
        For Each editor As BaseEdit In {NameEdit, ClientEdit, FamilyEdit, VersionEdit, RoleEdit, GenericEdit, NotesEdit}
            AddHandler editor.EditValueChanged, Sub(sender, args) MarkChanged()
        Next
        Dim tabs As New XtraTabControl With {.Dock = DockStyle.Fill}
        Dim rulesPage As New XtraTabPage With {.Text = "Reviewed differences / range rules"}
        Dim evidencePage As New XtraTabPage With {.Text = "Three-file evidence"}
        tabs.TabPages.AddRange({rulesPage, evidencePage})
        ConfigureGrid(RulesGrid, RulesView, True)
        ConfigureGrid(EvidenceGrid, EvidenceView, False)
        Dim rulesToolbar As New FlowLayoutPanel With {.Dock = DockStyle.Top, .Height = 38}
        AddButton(rulesToolbar, "Add rule", Sub()
                                               Definition.Rules.Add(New ModelManagerRule With {.Description = "New review item"})
                                               BindRules()
                                               MarkChanged()
                                           End Sub)
        AddButton(rulesToolbar, "Remove focused rule", Sub()
                                                           Dim rule = TryCast(RulesView.GetFocusedRow(), ModelManagerRule)
                                                           If rule Is Nothing Then Return
                                                           If XtraMessageBox.Show(Me, "Remove this rule from the draft?", "Model Manager", MessageBoxButtons.YesNo) <> DialogResult.Yes Then Return
                                                           Definition.Rules.Remove(rule)
                                                           BindRules()
                                                           MarkChanged()
                                                       End Sub)
        rulesPage.Controls.Add(RulesGrid)
        rulesPage.Controls.Add(rulesToolbar)
        Dim evidenceToolbar As New FlowLayoutPanel With {.Dock = DockStyle.Top, .Height = 70}
        For Each role In {"Source", "Original template", "Latest template"}
            Dim selectedRole = role
            AddButton(evidenceToolbar, "Choose " & role.ToLowerInvariant() & "...", Sub() SelectEvidence(selectedRole))
        Next
        AddButton(evidenceToolbar, "Verify selected evidence", AddressOf VerifyEvidence)
        evidencePage.Controls.Add(EvidenceGrid)
        evidencePage.Controls.Add(evidenceToolbar)
        AddHandler RulesView.CellValueChanged, Sub(sender, e) MarkChanged()
        Controls.Add(tabs)
        Controls.Add(metadata)
        Controls.Add(toolbar)
        Controls.Add(Status)
    End Sub

    Private Shared Sub AddField(panel As TableLayoutPanel, caption As String, editor As Control, column As Integer, row As Integer)
        panel.Controls.Add(New LabelControl With {.Text = caption, .Dock = DockStyle.Fill}, column, row)
        editor.Dock = DockStyle.Fill
        panel.Controls.Add(editor, column + 1, row)
    End Sub

    Private Sub AddButton(panel As FlowLayoutPanel, caption As String, action As Action)
        Dim button As New SimpleButton With {.Text = caption, .AutoSize = True}
        AddHandler button.Click, Sub(sender, e)
                                     Try
                                         action()
                                     Catch ex As Exception
                                         XtraMessageBox.Show(Me, ex.Message, "Model Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                     End Try
                                 End Sub
        panel.Controls.Add(button)
    End Sub

    Private Shared Sub ConfigureGrid(grid As GridControl, view As GridView, editable As Boolean)
        grid.MainView = view
        grid.ViewCollection.Add(view)
        view.OptionsBehavior.Editable = editable
        view.OptionsSelection.MultiSelect = True
        view.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect
        view.OptionsClipboard.CopyColumnHeaders = DefaultBoolean.True
        view.OptionsView.ShowGroupPanel = False
        view.OptionsView.ColumnAutoWidth = False
    End Sub

    Private Sub MarkChanged()
        If Not LoadingDefinition Then IsChanged = True
    End Sub

    Private Function ConfirmDiscard() As Boolean
        RulesView.CloseEditor()
        RulesView.UpdateCurrentRow()
        Return Not IsChanged OrElse XtraMessageBox.Show(Me, "Discard unsaved definition changes?", "Model Manager", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes
    End Function

    Private Sub LoadDefinition(value As ModelManagerDefinition)
        LoadingDefinition = True
        Try
            Definition = value
            EvidencePaths.Clear()
            NameEdit.Text = value.DisplayName
            ClientEdit.Text = value.ClientName
            FamilyEdit.Text = value.ModelFamily
            VersionEdit.Text = value.BaseVersion
            GenericEdit.Checked = value.IsGeneric
            RoleEdit.Text = value.ModelRole.ToString()
            NotesEdit.Text = value.Notes
            BindRules()
            BindEvidence()
            IsChanged = False
            Status.Text = "Draft revision " & value.Revision & ". Metadata only: rules are not executed. Sources and templates remain read-only; outputs always use a new filename." & Environment.NewLine &
                          If(ModelManagerAccess.PasswordRequired, "Password access enabled.", "Trial: password access disabled.")
        Finally
            LoadingDefinition = False
        End Try
    End Sub

    Private Sub BindRules()
        RulesGrid.DataSource = New BindingList(Of ModelManagerRule)(Definition.Rules)
        RulesView.PopulateColumns()
        RulesView.BestFitColumns()
    End Sub

    Private Sub BindEvidence()
        EvidenceGrid.DataSource = New BindingList(Of ModelManagerEvidence)(Definition.Evidence)
        EvidenceView.PopulateColumns()
        EvidenceView.BestFitColumns()
    End Sub

    Private Sub ReadEditors()
        RulesView.CloseEditor()
        If Not RulesView.UpdateCurrentRow() Then Throw New InvalidDataException("Please finish editing the current rule.")
        Definition.DisplayName = NameEdit.Text.Trim()
        Definition.ClientName = ClientEdit.Text.Trim()
        Definition.ModelFamily = FamilyEdit.Text.Trim()
        Definition.BaseVersion = VersionEdit.Text.Trim()
        Definition.IsGeneric = GenericEdit.Checked
        Definition.ModelRole = CType([Enum].Parse(GetType(ManagedModelRole), RoleEdit.Text), ManagedModelRole)
        Definition.Notes = NotesEdit.Text
        ModelManagerStore.Validate(Definition)
    End Sub

    Private Sub OpenXml()
        If Not ConfirmDiscard() Then Return
        Using dialog As New OpenFileDialog With {.Filter = "Model definitions (*.xml)|*.xml", .InitialDirectory = ModelManagerStore.LibraryDirectory}
            If dialog.ShowDialog(Me) = DialogResult.OK Then LoadDefinition(ModelManagerStore.LoadDefinition(dialog.FileName))
        End Using
    End Sub

    Private Sub OpenSaved()
        If Not ConfirmDiscard() Then Return
        Dim table As New System.Data.DataTable()
        For Each column In {"Definition", "Client", "Generic", "Role", "Revision", "File"}
            table.Columns.Add(column)
        Next
        If Directory.Exists(ModelManagerStore.LibraryDirectory) Then
            For Each folder In Directory.GetDirectories(ModelManagerStore.LibraryDirectory)
                Dim latest = Directory.GetFiles(folder, "revision-*.xml").OrderByDescending(Function(item) item, StringComparer.OrdinalIgnoreCase).FirstOrDefault()
                If latest Is Nothing Then Continue For
                Try
                    Dim saved = ModelManagerStore.LoadDefinition(latest)
                    table.Rows.Add(saved.DisplayName, saved.ClientName, saved.IsGeneric.ToString(), saved.ModelRole.ToString(), saved.Revision.ToString(), latest)
                Catch ex As Exception
                    table.Rows.Add("Unreadable definition: " & ex.Message, "", "", "", "", latest)
                End Try
            Next
        End If
        If table.Rows.Count = 0 Then
            XtraMessageBox.Show(Me, "No local definitions yet. Use Save revision to retain the current definition.", "Model Manager")
            Return
        End If
        Using dialog As New XtraForm With {.Text = "Saved model definitions — latest revision", .Size = New Size(950, 500), .StartPosition = FormStartPosition.CenterParent}
            Dim grid As New GridControl With {.Dock = DockStyle.Fill}
            Dim view As New GridView()
            ConfigureGrid(grid, view, False)
            grid.DataSource = table
            view.PopulateColumns()
            view.Columns("File").Visible = False
            view.BestFitColumns()
            Dim openButton As New SimpleButton With {.Text = "Open selected", .Dock = DockStyle.Bottom, .DialogResult = DialogResult.OK}
            dialog.Controls.Add(grid)
            dialog.Controls.Add(openButton)
            AddHandler view.DoubleClick, Sub(sender, e)
                                             Dim hit = view.CalcHitInfo(grid.PointToClient(Control.MousePosition))
                                             If hit.InRow AndAlso hit.RowHandle >= 0 Then dialog.DialogResult = DialogResult.OK
                                         End Sub
            If dialog.ShowDialog(Me) = DialogResult.OK AndAlso view.FocusedRowHandle >= 0 Then
                LoadDefinition(ModelManagerStore.LoadDefinition(CStr(view.GetFocusedRowCellValue("File"))))
            End If
        End Using
    End Sub

    Private Sub OpenWorkbook()
        If Not ConfirmDiscard() Then Return
        Using dialog As New OpenFileDialog With {.Filter = "Business plans (*.xlsb)|*.xlsb", .Title = "Read embedded definition (saved file only)"}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim value = ModelManagerStore.ReadEmbedded(dialog.FileName)
            If value Is Nothing Then Throw New InvalidDataException("No Model Manager definition is embedded in this file. Create a new definition; this file has not been changed.")
            LoadDefinition(value)
        End Using
    End Sub

    Private Sub SaveRevision()
        ReadEditors()
        Dim fileName = ModelManagerStore.SaveRevision(Definition)
        IsChanged = False
        Status.Text = "Draft revision saved: " & fileName
    End Sub

    Private Sub ExportXml()
        ReadEditors()
        Using dialog As New SaveFileDialog With {.Filter = "Model definition (*.xml)|*.xml", .FileName = "Model-definition.xml", .OverwritePrompt = True}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            ModelManagerStore.ExportDefinition(Definition, dialog.FileName)
            Status.Text = "Exported draft XML: " & dialog.FileName & ". Local revision unchanged."
        End Using
    End Sub

    Private Sub SelectEvidence(role As String)
        Using dialog As New OpenFileDialog With {.Filter = "Business plans (*.xlsb)|*.xlsb", .Title = "Select " & role & " — saved file fingerprint"}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            RejectUnsavedModel(dialog.FileName)
            Dim evidence As New ModelManagerEvidence With {.Role = role, .FileName = Path.GetFileName(dialog.FileName), .SHA256 = ModelManagerStore.Fingerprint(dialog.FileName)}
            Dim previous = Definition.Evidence.FirstOrDefault(Function(item) item.Role = role)
            If previous IsNot Nothing AndAlso Not String.Equals(previous.SHA256, evidence.SHA256, StringComparison.OrdinalIgnoreCase) Then
                If XtraMessageBox.Show(Me, "This file differs from the recorded " & role.ToLowerInvariant() & ". Replace the evidence and reset every rule to Unreviewed?", "Evidence changed", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then Return
                For Each rule In Definition.Rules
                    rule.Disposition = ModelRuleDisposition.Unreviewed
                Next
                BindRules()
            End If
            Definition.Evidence.RemoveAll(Function(item) item.Role = role)
            Definition.Evidence.Add(evidence)
            EvidencePaths(role) = dialog.FileName
            BindEvidence()
            MarkChanged()
        End Using
    End Sub

    Private Sub VerifyEvidence()
        Dim issues As New List(Of String)
        For Each evidence In Definition.Evidence
            Dim localPath As String = Nothing
            If Not EvidencePaths.TryGetValue(evidence.Role, localPath) Then
                issues.Add(evidence.Role & ": reselect the file to verify this saved fingerprint.")
            ElseIf Not File.Exists(localPath) OrElse ModelManagerStore.Fingerprint(localPath) <> evidence.SHA256 Then
                issues.Add(evidence.Role & ": the file differs from the recorded evidence. Review the rules again.")
            End If
        Next
        Status.Text = If(Definition.Evidence.Count = 0, "No evidence selected.", If(issues.Count = 0, "Selected evidence fingerprints match. This does not approve rules or validate a migration.", String.Join(Environment.NewLine, issues)))
    End Sub

    Private Shared Sub RejectUnsavedModel(fileName As String)
        If FileManager.ExcelModels Is Nothing Then Return
        For Each model In FileManager.ExcelModels
            If model IsNot Nothing AndAlso model.IsDirty AndAlso String.Equals(Path.GetFullPath(model.FileName), Path.GetFullPath(fileName), StringComparison.OrdinalIgnoreCase) Then
                Throw New InvalidOperationException("This model has unsaved Summit edits. Save them first or select a different saved file.")
            End If
        Next
    End Sub

    Private Sub SaveWorkbook()
        ReadEditors()
        Using source As New OpenFileDialog With {.Filter = "Business plans (*.xlsb)|*.xlsb", .Title = "Choose the saved workbook to COPY (not modified)"}
            If source.ShowDialog(Me) <> DialogResult.OK Then Return
            RejectUnsavedModel(source.FileName)
            Using output As New SaveFileDialog With {.Filter = "Business plans (*.xlsb)|*.xlsb", .InitialDirectory = Path.GetDirectoryName(source.FileName), .FileName = Path.GetFileNameWithoutExtension(source.FileName) & " - managed.xlsb", .OverwritePrompt = True}
                If output.ShowDialog(Me) <> DialogResult.OK Then Return
                If XtraMessageBox.Show(Me, "This adds the current DRAFT definition to a new copy. It does not apply rules, populate assumptions or run an upgrade. Continue?", "Save managed copy", MessageBoxButtons.YesNo, MessageBoxIcon.Information) <> DialogResult.Yes Then Return
                Cursor = Cursors.WaitCursor
                Try
                    ModelManagerStore.SaveWorkbookCopy(source.FileName, output.FileName, Definition)
                    Status.Text = "Managed copy saved: " & output.FileName & ". Worksheet/VBA parts preserved; no upgrade was executed."
                Finally
                    Cursor = Cursors.Default
                End Try
            End Using
        End Using
    End Sub
End Class
