Option Infer On
Imports System.Drawing
Imports System.Windows.Forms
Imports Abovo
Imports DevExpress.XtraTreeList
Imports DevExpress.XtraTreeList.Nodes
Imports DevExpress.XtraTreeList.Columns
Imports DevExpress.Spreadsheet

Partial Public Class BPIncomeExpenditureAnalyserV2
    Private BalanceSheetView As StatementTree

    Private Sub InitialiseBalanceSheet()
        BalanceSheetView = New StatementTree(Function() ReadBalanceSheet(), WrapCG_SOCI.WrappedGridView)
        WrapCG_BS.WrappedCGC.Visible = False
        WrapCG_BS.Controls.Add(BalanceSheetView)
        BalanceSheetView.BringToFront()
        AddHandler XtraTabPageBSWrapped.VisibleChanged, Sub(s, e)
                                                          If XtraTabPageBSWrapped.Visible Then BalanceSheetView.EnsureDocument()
                                                      End Sub
    End Sub

    Private Function ReadBalanceSheet() As BalanceSheetDocument
        Dim workbook = FileManager.ExcelModels(ModelID).WB
        Dim live = BalanceSheetStatement.Read(workbook)
        If CurrentDataSourceMode = AnalyserDataSourceMode.Live Then Return live
        Dim snapshot = BalanceSheetSnapshot.Read(workbook, live)
        Return If(CurrentDataSourceMode = AnalyserDataSourceMode.Snapshot, snapshot, BalanceSheetStatement.Difference(live, snapshot))
    End Function

    'Independent workbook headlines plus their explanation: never sum totals and
    'children into another total, and never use the incomplete UseInBS filter.
    Private NotInheritable Class StatementTree
        Inherits UserControl
        Private ReadOnly Loader As Func(Of BalanceSheetDocument)
        Private ReadOnly Tree As New TreeList With {.Dock = DockStyle.Fill}
        Private ReadOnly Notice As New Label With {.Dock = DockStyle.Top, .AutoSize = False, .Height = 42}
        Private ReadOnly Detail As New TextBox With {.Dock = DockStyle.Bottom, .ReadOnly = True, .Multiline = True, .Height = 48, .ScrollBars = ScrollBars.Vertical}
        Private ReadOnly Menu As New ContextMenuStrip
        Private ReadOnly NodeMap As New Dictionary(Of String, TreeListNode)(StringComparer.Ordinal)
        Private ReadOnly Expanded As New HashSet(Of String)(StringComparer.Ordinal)
        Private ReadOnly DisplayValues As New Dictionary(Of String, String())(StringComparer.Ordinal)
        Private Document As BalanceSheetDocument
        Private Dirty As Boolean = True
        Private Connected As Boolean = True
        Private Building As Boolean
        Private HasState As Boolean
        Private FocusId As String
        Private FocusColumn As Integer
        Private TopIndex As Integer
        Private HotNode As TreeListNode
        Private ReadOnly DataFont As Font
        Private ReadOnly HeadingFont As Font

        Public Sub New(load As Func(Of BalanceSheetDocument), referenceView As DevExpress.XtraGrid.Views.Grid.GridView)
            Loader = load
            DataFont = DirectCast(referenceView.Appearance.Row.Font.Clone(), Font)
            HeadingFont = New Font(DataFont, FontStyle.Bold)
            Tree.Font = DataFont
            Tree.Appearance.Row.Font = DataFont
            Tree.Appearance.HeaderPanel.Assign(referenceView.Appearance.HeaderPanel)
            Tree.Appearance.HorzLine.Assign(referenceView.Appearance.HorzLine)
            Tree.Appearance.VertLine.Assign(referenceView.Appearance.VertLine)
            Tree.Appearance.SelectedRow.BackColor = Color.Wheat
            Tree.Appearance.SelectedRow.ForeColor = Color.Black
            Tree.Appearance.HideSelectionRow.Assign(Tree.Appearance.SelectedRow)
            Tree.OptionsView.ShowIndicator = False
            Tree.OptionsView.ShowHorzLines = True
            Tree.OptionsView.ShowVertLines = True
            Tree.OptionsView.ColumnHeaderAutoHeight = DevExpress.Utils.DefaultBoolean.True
            Notice.Font = DataFont : Detail.Font = DataFont
            Notice.Height = CInt(Math.Ceiling(DataFont.GetHeight() * 2.5))
            Detail.Height = CInt(Math.Ceiling(DataFont.GetHeight() * 3))
            Dock = DockStyle.Fill
            Controls.Add(Tree) : Controls.Add(Notice) : Controls.Add(Detail)
            Notice.Text = "Balances include opening amounts; drill-down movements are cumulative."
            Tree.OptionsBehavior.Editable = False
            Tree.OptionsBehavior.ReadOnly = True
            Tree.OptionsView.AutoWidth = False
            Tree.OptionsSelection.MultiSelect = True
            Tree.OptionsSelection.MultiSelectMode = TreeListMultiSelectMode.CellSelect
            Tree.OptionsSelection.EnableAppearanceFocusedCell = False
            Tree.OptionsSelection.EnableAppearanceFocusedRow = False
            Tree.OptionsClipboard.CopyColumnHeaders = DevExpress.Utils.DefaultBoolean.False
            Tree.OptionsCustomization.AllowSort = False
            Tree.OptionsCustomization.AllowFilter = False
            Tree.ContextMenuStrip = Menu
            AddHandler Tree.SizeChanged, Sub() LimitDescriptionWidth()
            GridPresentation.SetResetLayout(Tree, Sub() Tree.RowHeight = -1)
            Menu.Items.Add("Copy", Nothing, Sub(s, e) Copy(False))
            Menu.Items.Add("Copy with headings", Nothing, Sub(s, e) Copy(True))
            AddHandler Tree.FocusedNodeChanged, Sub(s, e) ShowSource()
            AddHandler Tree.NodeCellStyle, AddressOf StyleCell
            AddHandler Tree.MouseMove, Sub(s, e) SetHotNode(Tree.CalcHitInfo(e.Location).Node)
            AddHandler Tree.MouseLeave, Sub(s, e) SetHotNode(Nothing)
            AddHandler Tree.CustomDrawNodeCell, Sub(s, e)
                                                   Dim model = TryCast(e.Node.Tag, BalanceSheetNode)
                                                   Dim values As String() = Nothing
                                                   If model IsNot Nothing AndAlso e.Column.AbsoluteIndex > 0 AndAlso DisplayValues.TryGetValue(model.Id, values) Then e.CellText = values(e.Column.AbsoluteIndex - 1)
                                               End Sub
            AddHandler VisibleChanged, Sub(s, e)
                                           If Visible Then EnsureDocument()
                                       End Sub
        End Sub

        Private Sub SetHotNode(node As TreeListNode)
            If node Is HotNode Then Return
            Dim previous = HotNode
            HotNode = node
            If previous IsNot Nothing Then Tree.InvalidateNode(previous)
            If node IsNot Nothing Then Tree.InvalidateNode(node)
        End Sub

        Private Sub LimitDescriptionWidth()
            If Tree.Columns.Count = 0 OrElse Tree.ClientSize.Width <= 0 Then Return
            Tree.Columns(0).MaxWidth = Math.Max(Tree.Columns(0).MinWidth, CInt(Tree.ClientSize.Width * 0.4))
        End Sub

        Private Sub Copy(headings As Boolean)
            Tree.OptionsClipboard.CopyColumnHeaders = If(headings, DevExpress.Utils.DefaultBoolean.True, DevExpress.Utils.DefaultBoolean.False)
            Tree.CopyToClipboard()
            Tree.OptionsClipboard.CopyColumnHeaders = DevExpress.Utils.DefaultBoolean.False
        End Sub

        Public Sub SourceChanged(disconnected As Boolean)
            If Not Dirty AndAlso Document IsNot Nothing Then
                Expanded.Clear()
                For Each pair In NodeMap
                    If pair.Value.Expanded Then Expanded.Add(pair.Key)
                Next
                FocusId = TryCast(Tree.FocusedNode?.Tag, BalanceSheetNode)?.Id
                FocusColumn = If(Tree.FocusedColumn Is Nothing, 0, Tree.FocusedColumn.AbsoluteIndex)
                TopIndex = Tree.TopVisibleNodeIndex
                HasState = True
            End If
            Connected = Not disconnected
            Dirty = True
            Document = Nothing
            HotNode = Nothing
            Tree.Nodes.Clear() : NodeMap.Clear() : DisplayValues.Clear()
            Detail.Clear()
            Notice.Text = If(disconnected, "Analysis disconnected; refresh analysis before using these figures.", "Balances include opening amounts; movements are cumulative.")
            If Connected AndAlso Visible Then EnsureDocument()
        End Sub

        Public Function EnsureDocument() As BalanceSheetDocument
            If Not Connected OrElse Building Then Return Nothing
            If Not Dirty Then Return Document
            Building = True
            Tree.BeginUpdate()
            Try
                Document = Loader()
                Tree.Nodes.Clear() : NodeMap.Clear() : DisplayValues.Clear()
                If Tree.Columns.Count = 0 Then
                    Tree.Columns.Add(New TreeListColumn With {.FieldName = "Description", .Caption = "Item / Description", .VisibleIndex = 0, .Width = 360, .Fixed = FixedStyle.Left})
                    For p = 0 To 40
                        Tree.Columns.Add(New TreeListColumn With {.FieldName = "Period" & p.ToString(), .VisibleIndex = p + 1, .Width = 105})
                    Next
                End If
                For p = 0 To 40
                    Tree.Columns(p + 1).Caption = Document.Periods(p)
                Next
                LimitDescriptionWidth()
                'The native spreadsheet formatter handles Excel formats (including
                'zero/negative sections); no calculation or model write in paint.
                Using formatting As New Workbook()
                    Dim cell = formatting.Worksheets(0).Cells(0, 0)
                    Dim strings As New Dictionary(Of String, String)(StringComparer.Ordinal)
                    For Each model In Document.Nodes
                        Dim display(40) As String
                        For p = 0 To 40
                            Dim format = Document.Styles(model.Styles(p + 1)).Format
                            Dim key = format & "|" & model.Values(p).ToString("R", Globalization.CultureInfo.InvariantCulture)
                            If Not strings.TryGetValue(key, display(p)) Then
                                If Double.IsNaN(model.Values(p)) OrElse Double.IsInfinity(model.Values(p)) Then
                                    display(p) = "Unavailable"
                                Else
                                    cell.NumberFormat = format : cell.Value = model.Values(p)
                                    display(p) = cell.DisplayText
                                End If
                                strings.Add(key, display(p))
                            End If
                        Next
                        DisplayValues.Add(model.Id, display)
                    Next
                End Using
                'Difference-only branches may follow their children in the union.
                AppendChildren("", Nothing)
                For Each pair In NodeMap
                    pair.Value.Expanded = If(HasState, Expanded.Contains(pair.Key), CType(pair.Value.Tag, BalanceSheetNode).ParentId = "")
                Next
                If FocusId IsNot Nothing AndAlso NodeMap.ContainsKey(FocusId) Then Tree.FocusedNode = NodeMap(FocusId)
                Tree.FocusedColumn = Tree.Columns(Math.Min(FocusColumn, Tree.Columns.Count - 1))
                Tree.TopVisibleNodeIndex = TopIndex
                Notice.Text = If(Document.Diagnostic.Length = 0, "Balances include opening amounts; drill-down movements are cumulative.", "Some drill-down mappings are unavailable. Select the affected line for details.")
                ShowSource()
                Dirty = False
                Return Document
            Catch ex As Exception
                Document = Nothing
                Tree.Nodes.Clear() : NodeMap.Clear()
                Notice.Text = "Balance Sheet unavailable: " & ex.Message
                Detail.Text = ex.Message
                Abovo.SummitDiagnostics.WriteLine("[Balance Sheet view] " & ex.ToString())
                Dirty = False
                Return Nothing
            Finally
                Tree.EndUpdate()
                Building = False
            End Try
        End Function

        Private Sub AppendChildren(parentId As String, parent As TreeListNode)
            For Each model In Document.Children(parentId)
                Dim values(41) As Object
                values(0) = model.Caption
                For p = 0 To 40
                    values(p + 1) = DisplayValues(model.Id)(p)
                Next
                Dim node = Tree.AppendNode(values, parent)
                node.Tag = model
                NodeMap.Add(model.Id, node)
                AppendChildren(model.Id, node)
            Next
        End Sub

        Private Sub StyleCell(sender As Object, e As GetCustomNodeCellStyleEventArgs)
            Dim model = TryCast(e.Node.Tag, BalanceSheetNode)
            If model Is Nothing OrElse Document Is Nothing Then Return
            Dim index = model.Styles(e.Column.AbsoluteIndex)
            Dim style = Document.Styles(index)
            e.Appearance.Font = If(model.IsHeadline OrElse e.Node.HasChildren, HeadingFont, DataFont)
            e.Appearance.FontSizeDelta = CInt(Math.Round(e.Appearance.Font.SizeInPoints * (GridPresentation.ZoomPercent(Tree) / 100.0 - 1.0)))
            e.Appearance.ForeColor = If(e.Column.AbsoluteIndex > 0 AndAlso model.Values(e.Column.AbsoluteIndex - 1) < 0, Color.Red, Color.Black)
            e.Appearance.BackColor = If(model.IsTotal, Color.LightSteelBlue,
                If(e.Node.Level = 0, Color.FromArgb(235, 235, 250),
                If(e.Node.Level = 1, Color.FromArgb(240, 240, 250), Color.FromArgb(250, 250, 255))))
            If Tree.IsCellSelected(e.Node, e.Column) Then
                e.Appearance.BackColor = Color.Wheat
                e.Appearance.ForeColor = Color.Black
            ElseIf e.Node Is HotNode Then
                e.Appearance.BackColor = GeneralFunctions.AbovoBlue
                e.Appearance.ForeColor = Color.White
            End If
            Select Case CType(style.Alignment, SpreadsheetHorizontalAlignment)
                Case SpreadsheetHorizontalAlignment.Left
                    e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                Case SpreadsheetHorizontalAlignment.Center
                    e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                Case SpreadsheetHorizontalAlignment.Right
                    e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                Case Else
                    e.Appearance.TextOptions.HAlignment = If(e.Column.AbsoluteIndex = 0, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.HorzAlignment.Far)
            End Select
        End Sub

        Private Sub ShowSource()
            Dim model = TryCast(Tree.FocusedNode?.Tag, BalanceSheetNode)
            Detail.Text = If(model Is Nothing, "", model.Diagnostic & If(model.Diagnostic.Length > 0, Environment.NewLine, "") & model.Source & Environment.NewLine & model.Rule)
        End Sub

        Public Sub SetExpanded(expand As Boolean)
            EnsureDocument()
            If expand Then Tree.ExpandAll() Else Tree.CollapseAll()
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                Menu.Dispose()
            End If
            MyBase.Dispose(disposing)
            If disposing Then
                DataFont.Dispose()
                HeadingFont.Dispose()
            End If
        End Sub
    End Class
End Class
