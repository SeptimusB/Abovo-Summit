Imports System.Data
Imports System.IO
Imports System.Xml.Linq
Imports Abovo
Imports Abovo.StructureAuthoringDraft
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraTreeList
Imports DevExpress.XtraTreeList.Nodes

Public Class StructureManagerForm
    Inherits XtraForm

    Private ReadOnly tree As New TreeList With {.Dock = DockStyle.Fill}
    Private ReadOnly propertiesGrid As New GridControl With {.Dock = DockStyle.Fill}
    Private ReadOnly propertiesView As New GridView()
    Private ReadOnly workbookHost As New Panel With {.Dock = DockStyle.Fill}
    Private ReadOnly previewHost As New Panel With {.Dock = DockStyle.Fill}
    Private ReadOnly outerSplit As New SplitContainerControl With {.Dock = DockStyle.Fill}
    Private ReadOnly workSplit As New SplitContainerControl With {.Dock = DockStyle.Fill}
    Private ReadOnly propertySplit As New SplitContainerControl With {.Dock = DockStyle.Fill, .Horizontal = False}
    Private ReadOnly status As New LabelControl With {.Dock = DockStyle.Bottom, .Height = 46, .AutoSizeMode = LabelAutoSizeMode.None, .Padding = New Padding(6)}
    Private ReadOnly selectionInfo As New LabelControl With {.Dock = DockStyle.Top, .Height = 28, .AutoSizeMode = LabelAutoSizeMode.None}
    Private ReadOnly identity As New LabelControl With {.Dock = DockStyle.Top, .Height = 44, .AutoSizeMode = LabelAutoSizeMode.None, .Padding = New Padding(6)}
    Private ReadOnly nodes As New Dictionary(Of XElement, TreeListNode)
    Private session As StructureAuthoringSession
    Private draft As StructureAuthoringDraft
    Private preview As DataInterfaceTemplate
    Private previewChild As XElement
    Private selected As XElement
    Private propertyTable As DataTable
    Private suppressSelection As Boolean
    Private linking As Boolean
    Private pendingProperties As Boolean

    Public Shared Sub Launch(owner As IWin32Window)
        Try
            'An interactive, separate process owns all authoring globals and temporary models.
            Diagnostics.Process.Start(New Diagnostics.ProcessStartInfo(Application.ExecutablePath, "--structure-manager") With {.UseShellExecute = False})
        Catch ex As Exception
            XtraMessageBox.Show(owner, "Structure Manager could not be opened: " & ex.Message, "Structure Manager")
        End Try
    End Sub

    Public Sub New()
        Text = "Structure Manager — isolated draft trial"
        StartPosition = FormStartPosition.CenterScreen
        Size = New Size(1600, 950)
        MinimumSize = New Size(1100, 650)
        Dim commands As New FlowLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True, .WrapContents = True, .Padding = New Padding(4)}
        AddCommand(commands, "Open XLSB…", AddressOf ChooseWorkbook)
        AddCommand(commands, "Load structure XML…", AddressOf ChooseStructure)
        AddCommand(commands, "Save draft as…", AddressOf SaveDraft)
        AddCommand(commands, "Refresh preview", Sub() RefreshPreview(True))
        AddCommand(commands, "Definitions / bespoke rules…", Sub()
                                                                    Using manager As New ModelManagerForm()
                                                                        manager.ShowDialog(Me)
                                                                    End Using
                                                                End Sub)
        Dim search As New TextEdit With {.Width = 240}
        search.Properties.NullValuePrompt = "Find an interface, range or field"
        commands.Controls.Add(search)
        AddCommand(commands, "Find next", Sub() FindNext(search.Text))
        Dim bindingCommands As New FlowLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True}
        AddCommand(bindingCommands, "Apply properties", AddressOf ApplyProperties)
        AddCommand(bindingCommands, "Use selected cells…", AddressOf BindSelection)
        AddCommand(bindingCommands, "Discard property edits", AddressOf ReadProperties)
        AddCommand(bindingCommands, "Locate binding", AddressOf LocateBinding)
        propertiesGrid.MainView = propertiesView
        propertiesGrid.ViewCollection.Add(propertiesView)
        propertiesView.OptionsView.ShowGroupPanel = False
        propertiesView.OptionsSelection.MultiSelect = True
        propertiesView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect
        propertiesView.OptionsClipboard.CopyColumnHeaders = DevExpress.Utils.DefaultBoolean.True
        AddHandler propertiesView.ShowingEditor, Sub(sender, e)
                                                      e.Cancel = propertiesView.FocusedColumn Is Nothing OrElse propertiesView.FocusedColumn.FieldName <> "Value" OrElse Not CBool(propertiesView.GetFocusedRowCellValue("Editable"))
                                                  End Sub
        AddHandler propertiesView.CellValueChanged, Sub() pendingProperties = True
        propertySplit.Panel1.Controls.Add(workbookHost)
        propertySplit.Panel1.Controls.Add(selectionInfo)
        propertySplit.Panel2.Controls.Add(propertiesGrid)
        propertySplit.Panel2.Controls.Add(bindingCommands)
        workSplit.Panel1.Controls.Add(propertySplit)
        workSplit.Panel2.Controls.Add(previewHost)
        outerSplit.Panel1.Controls.Add(tree)
        outerSplit.Panel2.Controls.Add(workSplit)
        tree.Columns.AddVisible("Structure").Caption = "Structure / bindings"
        tree.OptionsBehavior.Editable = False
        tree.OptionsView.ShowIndicator = False
        tree.OptionsView.AutoWidth = False
        tree.Columns(0).Width = 550
        AddHandler tree.FocusedNodeChanged, Sub(sender, e) SelectNode(e.Node, e.OldNode)
        Controls.Add(outerSplit)
        Controls.Add(identity)
        Controls.Add(commands)
        Controls.Add(status)
        identity.Text = "Open an XLSB to inspect its structure. The source is never opened for writing."
        status.Text = "Read-only workbook and real DIT preview. Edit bindings in the properties below the spreadsheet; save to a NEW XML file."
        AddHandler Shown, Sub()
                              outerSplit.SplitterPosition = CInt(ClientSize.Width * 0.2)
                              workSplit.SplitterPosition = CInt(outerSplit.Panel2.Width * 0.53)
                              propertySplit.SplitterPosition = CInt(propertySplit.Height * 0.65)
                          End Sub
        AddHandler FormClosing, Sub(sender, e)
                                    If Not ConfirmDiscard() Then e.Cancel = True
                                End Sub
    End Sub

    Private Sub AddCommand(parent As Control, caption As String, action As Action)
        Dim button As New SimpleButton With {.Text = caption, .AutoSize = True, .Padding = New Padding(5)}
        AddHandler button.Click, Sub() RunSafe(action)
        parent.Controls.Add(button)
    End Sub

    Private Sub RunSafe(action As Action)
        Try
            Cursor = Cursors.WaitCursor
            action()
        Catch ex As Exception
            status.Text = "Not applied: " & ex.Message
            XtraMessageBox.Show(Me, ex.Message, "Structure Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Function ConfirmDiscard() As Boolean
        propertiesView.CloseEditor()
        propertiesView.UpdateCurrentRow()
        Return Not (pendingProperties OrElse (draft IsNot Nothing AndAlso draft.IsDirty)) OrElse XtraMessageBox.Show(Me, "Discard unsaved draft/property changes? The workbook source has not been changed.", "Structure Manager", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes
    End Function

    Private Sub ChooseWorkbook()
        If Not ConfirmDiscard() Then Return
        Using dialog As New OpenFileDialog With {.Filter = "Business plan (*.xlsb)|*.xlsb", .CheckFileExists = True}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            LoadWorkbook(dialog.FileName)
        End Using
    End Sub

    Friend Sub LoadWorkbook(path As String)
        ReleaseSession()
        session = New StructureAuthoringSession(path)
        draft = session.Draft
        workbookHost.Controls.Add(session.Model.ModelSpreadsheetControl)
        AddHandler session.Model.ModelSpreadsheetControl.SelectionChanged, AddressOf WorkbookSelectionChanged
        identity.Text = "Source: " & session.SourceWorkbook & Environment.NewLine & "Structure: " & draft.Source & " — draft only; neither source is overwritten."
        BuildTree()
        status.Text = "Select an interface or binding on the left. Workbook and DIT previews are read-only; specialist class forms are inspection-only in this first trial."
    End Sub

    Private Sub ChooseStructure()
        If session Is Nothing Then Throw New InvalidOperationException("Open a workbook first.")
        If Not ConfirmDiscard() Then Return
        Using dialog As New OpenFileDialog With {.Filter = "Summit structure (*.xml)|*.xml", .CheckFileExists = True}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim replacement As New StructureAuthoringDraft(dialog.FileName)
            ClearPreview()
            draft = replacement
            identity.Text = "Source: " & session.SourceWorkbook & Environment.NewLine & "Structure: " & draft.Source & " — draft only."
            BuildTree()
        End Using
    End Sub

    Private Sub SaveDraft()
        If draft Is Nothing Then Return
        propertiesView.CloseEditor()
        propertiesView.UpdateCurrentRow()
        If pendingProperties Then Throw New InvalidOperationException("Apply or discard the pending property edits before saving.")
        Using dialog As New SaveFileDialog With {.Filter = "Summit structure (*.xml)|*.xml", .FileName = "Structure-draft-" & DateTime.Now.ToString("yyyyMMdd-HHmmss") & ".xml"}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            draft.SaveNew(dialog.FileName)
            status.Text = "Draft saved: " & dialog.FileName & ". Not published or embedded; the installed structure and XLSB remain unchanged."
        End Using
    End Sub

    Private Shared Function Caption(element As XElement) As String
        Dim title = {"GSName", "CSName", "ISName", "ISDName", "FieldName", "NRDSName", "Type"}.Select(Function(n) Value(element, n)).FirstOrDefault(Function(v) v <> "")
        title = If(title, "").Replace(vbCr, " ").Replace(vbLf, " ")
        Select Case element.Name.LocalName
            Case "GroupStructure"
                Return title & " [group " & Value(element, "GSID") & "]"
            Case "ChildStructure"
                Return title & " [" & Value(element, "CSID") & "]"
            Case "CSInterfaceSection", "CSHeader"
                Return "Section · " & title
            Case "CellRangeDataSource"
                Return "Range · " & If(title = "" OrElse title = "CR", Value(element, "DataRange"), title)
            Case "DataFieldDefinition"
                Return "Field · " & title
            Case "ISDatasource"
                Return "Data · " & title
            Case Else
                Return element.Name.LocalName & If(title = "", "", " · " & title)
        End Select
    End Function

    Private Sub BuildTree()
        suppressSelection = True
        tree.BeginUnboundLoad()
        Try
            tree.Nodes.Clear()
            nodes.Clear()
            AddNode(draft.Document.Root, Nothing)
            tree.Nodes(0).Expanded = True
            For Each child As TreeListNode In tree.Nodes(0).Nodes
                If DirectCast(child.Tag, XElement).Name.LocalName = "GroupStructure" Then child.Expanded = True
            Next
            selected = Nothing
            pendingProperties = False
            propertiesGrid.DataSource = Nothing
        Finally
            tree.EndUnboundLoad()
            suppressSelection = False
        End Try
    End Sub

    Private Sub AddNode(element As XElement, parent As TreeListNode)
        Dim node = tree.AppendNode(New Object() {Caption(element)}, parent, element)
        nodes.Add(element, node)
        For Each child In element.Elements().Where(Function(x) x.HasElements)
            AddNode(child, node)
        Next
    End Sub

    Private Sub SelectNode(node As TreeListNode, oldNode As TreeListNode)
        If suppressSelection OrElse node Is Nothing Then Return
        propertiesView.CloseEditor()
        propertiesView.UpdateCurrentRow()
        If pendingProperties AndAlso XtraMessageBox.Show(Me, "Discard the unapplied property edits before changing selection?", "Structure Manager", MessageBoxButtons.YesNo) <> DialogResult.Yes Then
            suppressSelection = True
            tree.FocusedNode = oldNode
            suppressSelection = False
            Return
        End If
        selected = DirectCast(node.Tag, XElement)
        ReadProperties()
        RunSafe(Sub()
                    RefreshPreview(False)
                    If Not linking Then LocateBinding()
                End Sub)
    End Sub

    Private Sub ReadProperties()
        If selected Is Nothing Then Return
        propertyTable = New DataTable()
        propertyTable.Columns.Add("Property", GetType(String))
        propertyTable.Columns.Add("Value", GetType(String))
        propertyTable.Columns.Add("Editable", GetType(Boolean))
        Dim editable = EditableNames(selected)
        Dim names = selected.Elements().Where(Function(x) Not x.HasElements AndAlso x.Name.LocalName <> "RejData").Select(Function(x) x.Name.LocalName).Concat(editable).Distinct()
        For Each propertyName As String In names
            propertyTable.Rows.Add(propertyName, Value(selected, propertyName), editable.Contains(propertyName))
        Next
        propertiesGrid.DataSource = propertyTable
        propertiesView.PopulateColumns()
        propertiesView.Columns("Property").Width = 140
        propertiesView.Columns("Value").Width = 260
        propertiesView.Columns("Editable").Visible = False
        pendingProperties = False
    End Sub

    Private Sub ApplyProperties()
        If selected Is Nothing Then Return
        propertiesView.CloseEditor()
        propertiesView.UpdateCurrentRow()
        Dim changes As New Dictionary(Of String, String)
        For Each row As DataRow In propertyTable.Rows
            If CBool(row("Editable")) AndAlso CStr(row("Value")) <> Value(selected, CStr(row("Property"))) Then changes.Add(CStr(row("Property")), CStr(row("Value")))
        Next
        If changes.Count = 0 Then Return
        draft.Apply(selected, changes, session.Model.WB)
        tree.FocusedNode.SetValue(0, Caption(selected))
        ReadProperties()
        RefreshPreview(True)
        LocateBinding()
        status.Text = "Draft properties applied. Preview rebuilt from the draft; no workbook cells or production XML were saved."
    End Sub

    Private Sub BindSelection()
        If selected Is Nothing OrElse selected.Name.LocalName <> "CellRangeDataSource" Then Throw New InvalidOperationException("Select a CellRangeDataSource node first. The trial does not change merged/period axes implicitly.")
        If pendingProperties Then Throw New InvalidOperationException("Apply or discard property edits first.")
        Dim cells = session.Model.ModelSpreadsheetControl.Selection
        Dim existing = ResolveRange(selected, session.Model.WB)
        If cells Is Nothing OrElse existing Is Nothing Then Return
        If cells.Areas.Count <> 1 OrElse existing.Areas.Count <> 1 OrElse cells.RowCount <> existing.RowCount OrElse cells.ColumnCount <> existing.ColumnCount Then Throw New InvalidOperationException("This first trial only rebinds a single contiguous range with the same geometry. Expansion and mixed month/year axes require explicit structural rules.")
        Dim changes As New Dictionary(Of String, String) From {{"Worksheet", cells.Worksheet.Name}}
        Dim name = Value(selected, "NRDSName")
        If name <> "" AndAlso name <> "CR" Then
            Dim matches = session.Model.WB.DefinedNames.Where(Function(n) n.Range IsNot Nothing AndAlso n.Range.GetReferenceA1() = cells.GetReferenceA1() AndAlso n.Range.Worksheet.Name = cells.Worksheet.Name).ToList()
            If matches.Count <> 1 Then Throw New InvalidOperationException("Select cells matching exactly one existing workbook named range, or enter the intended NRDSName in properties. Names are never created or resized implicitly.")
            changes.Add("NRDSName", matches(0).Name)
        Else
            changes.Add("DataRange", cells.GetReferenceA1())
        End If
        If XtraMessageBox.Show(Me, "Rebind this draft datasource to " & cells.Worksheet.Name & "!" & cells.GetReferenceA1() & "?" & Environment.NewLine & "Workbook values and named ranges will not change.", "Confirm draft binding", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        draft.Apply(selected, changes, session.Model.WB)
        ReadProperties()
        RefreshPreview(True)
    End Sub

    Private Sub WorkbookSelectionChanged(sender As Object, e As EventArgs)
        Dim cells = session?.Model.ModelSpreadsheetControl.Selection
        If cells IsNot Nothing Then selectionInfo.Text = "Selected: " & cells.Worksheet.Name & "!" & cells.GetReferenceA1() & " (read-only)"
    End Sub

    Private Sub LocateBinding()
        If selected Is Nothing OrElse session Is Nothing Then Return
        Dim cells = ResolveRange(selected, session.Model.WB)
        If cells IsNot Nothing Then
            ShowCells(cells)
        Else
            Dim sheetName = Value(Ancestor(selected, "ChildStructure"), "DefaultWorksheet")
            If session.Model.WB.Worksheets.Contains(sheetName) Then session.Model.WB.Worksheets.ActiveWorksheet = session.Model.WB.Worksheets(sheetName)
        End If
    End Sub

    Private Sub ShowCells(cells As CellRange)
        session.Model.WB.Worksheets.ActiveWorksheet = cells.Worksheet
        session.Model.ModelSpreadsheetControl.Selection = cells.Areas(0)
        WorkbookSelectionChanged(Me, EventArgs.Empty)
    End Sub

    Private Sub RefreshPreview(force As Boolean)
        If selected Is Nothing OrElse session Is Nothing Then Return
        If force AndAlso pendingProperties Then Throw New InvalidOperationException("Apply or discard pending property edits before refreshing.")
        Dim child = Ancestor(selected, "ChildStructure")
        If child Is Nothing Then
            ClearPreview()
            Return
        End If
        Dim section = Ancestor(selected, "CSInterfaceSection")
        Dim sectionIndex = If(section Is Nothing, 0, child.Elements("CSInterfaceSection").ToList().IndexOf(section))
        If force OrElse child IsNot previewChild Then
            ClearPreview()
            If Value(child, "SpecialElement") <> "" OrElse Not child.Elements("CSInterfaceSection").Any() Then
                previewHost.Controls.Add(New LabelControl With {.Dock = DockStyle.Top, .AutoSizeMode = LabelAutoSizeMode.None, .Height = 80, .Text = "This is a specialist interface, not a DIT. Its XML can be inspected here; native preview is not available in this trial."})
                Return
            End If
            session.PreparePreview(draft.Document)
            Dim groupId = Integer.Parse(Value(Ancestor(child, "GroupStructure"), "GSID"))
            Dim childId = Integer.Parse(Value(child, "CSID"))
            Try
                preview = New DataInterfaceTemplate(session.ModelID, groupId, childId, Nothing, "StructurePreview") With {.TopLevel = False, .FormBorderStyle = FormBorderStyle.None, .Dock = DockStyle.Fill}
                previewHost.Controls.Add(preview)
                previewChild = child
                AddHandler preview.AuthoringSourceSelected, AddressOf PreviewSourceSelected
                preview.Show()
            Catch
                ClearPreview()
                Throw
            End Try
        End If
        If preview IsNot Nothing Then preview.SelectAuthoringSection(sectionIndex)
    End Sub

    Private Sub PreviewSourceSelected(sheet As String, address As String)
        If linking OrElse session Is Nothing OrElse previewChild Is Nothing Then Return
        linking = True
        Try
            Dim cell = session.Model.WB.Worksheets(sheet).Cells(address)
            'Use runtime-resolved cell coordinates, not assumptions about pivot/period axes.
            For Each binding In previewChild.Descendants("CellRangeDataSource")
                Dim range As CellRange = Nothing
                Try
                    range = ResolveRange(binding, session.Model.WB)
                Catch
                    Continue For
                End Try
                If range IsNot Nothing AndAlso range.Worksheet.Name = sheet AndAlso range.Contains(cell) Then
                    tree.FocusedNode = nodes(binding)
                    tree.MakeNodeVisible(tree.FocusedNode)
                    Exit For
                End If
            Next
            ShowCells(cell)
        Finally
            linking = False
        End Try
    End Sub

    Private Sub FindNext(text As String)
        If String.IsNullOrWhiteSpace(text) Then Return
        Dim all = nodes.Values.ToList()
        Dim start = all.IndexOf(tree.FocusedNode)
        For offset = 1 To all.Count
            Dim node = all((start + offset) Mod all.Count)
            If CStr(node.GetValue(0)).IndexOf(text, StringComparison.OrdinalIgnoreCase) < 0 Then Continue For
            tree.FocusedNode = node
            tree.MakeNodeVisible(node)
            Return
        Next
        status.Text = "No matching structure node."
    End Sub

    Private Sub ClearPreview()
        If preview IsNot Nothing Then
            RemoveHandler preview.AuthoringSourceSelected, AddressOf PreviewSourceSelected
            preview.Dispose()
            preview = Nothing
        End If
        For Each item As Control In previewHost.Controls.Cast(Of Control)().ToArray()
            item.Dispose()
        Next
        previewHost.Controls.Clear()
        previewChild = Nothing
    End Sub

    Private Sub ReleaseSession()
        ClearPreview()
        If session IsNot Nothing Then
            RemoveHandler session.Model.ModelSpreadsheetControl.SelectionChanged, AddressOf WorkbookSelectionChanged
            workbookHost.Controls.Remove(session.Model.ModelSpreadsheetControl)
            session.Dispose()
            session = Nothing
        End If
        draft = Nothing
        selected = Nothing
        suppressSelection = True
        tree.Nodes.Clear()
        nodes.Clear()
        suppressSelection = False
        propertiesGrid.DataSource = Nothing
        pendingProperties = False
        identity.Text = "No workbook open — select an XLSB source."
        selectionInfo.Text = String.Empty
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso Not IsDisposed Then ReleaseSession()
        MyBase.Dispose(disposing)
    End Sub
End Class
