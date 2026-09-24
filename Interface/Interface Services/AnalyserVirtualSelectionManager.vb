Imports System.Collections
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms
Imports Abovo.CustomGrid
Imports DevExpress.Utils.Drawing
Imports DevExpress.Utils.Menu
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Drawing
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo

Namespace Abovo

    Friend Enum AnalyserVirtualRowKind
        DataRow
        GroupHeading
        GroupFooter
    End Enum

    Friend NotInheritable Class AnalyserVirtualSelectionManager

        Private NotInheritable Class VisualRowSelection
            Public View As CustomGridView
            Public RowHandle As Integer
            Public Kind As AnalyserVirtualRowKind
            Public WholeRow As Boolean
            Public OrderY As Integer
            Public ReadOnly Fields As New HashSet(Of String)(StringComparer.Ordinal)
        End Class

        Private ReadOnly AttachedViews As New HashSet(Of CustomGridView)()
        Private ReadOnly Selections As New List(Of VisualRowSelection)()
        Private ReadOnly ExpandBranchActions As New Dictionary(
            Of CustomGridView, Action(Of CustomGridView, Integer))()
        Private ReadOnly CollapseBranchActions As New Dictionary(
            Of CustomGridView, Action(Of CustomGridView, Integer))()
        Private ReadOnly ContextMenus As New Dictionary(Of CustomGridView, ContextMenuStrip)()
        Private ReadOnly GroupClickTimer As New Timer()
        Private DragView As CustomGridView
        Private DragStart As Point
        Private DragCurrent As Point
        Private DragCandidate As Boolean
        Private Dragging As Boolean
        Private DragIsAdditive As Boolean
        Private DragAnchorRowHandle As Integer
        Private DragAnchorKind As AnalyserVirtualRowKind
        Private DragAnchorValid As Boolean
        Private SuppressNextClick As Boolean
        Private HoverView As CustomGridView
        Private HoverRowHandle As Integer
        Private HoverKind As AnalyserVirtualRowKind
        Private HasHover As Boolean
        Private PendingGroupView As CustomGridView
        Private PendingGroupRowHandle As Integer
        Private HasPendingGroupClick As Boolean
        Private ApplyingGroupAction As Boolean
        Private SuppressNativeGroupActionValue As Boolean

        Public ReadOnly Property IsApplyingGroupAction As Boolean
            Get
                Return ApplyingGroupAction
            End Get
        End Property

        Public ReadOnly Property ShouldSuppressNativeGroupAction As Boolean
            Get
                Return SuppressNativeGroupActionValue AndAlso Not ApplyingGroupAction
            End Get
        End Property

        Public Sub New()
            GroupClickTimer.Interval = Math.Max(100, SystemInformation.DoubleClickTime)
            AddHandler GroupClickTimer.Tick, AddressOf GroupClickTimer_Tick
        End Sub

        Public Sub Attach(
            ByVal view As CustomGridView,
            Optional ByVal expandBranch As Action(Of CustomGridView, Integer) = Nothing,
            Optional ByVal collapseBranch As Action(Of CustomGridView, Integer) = Nothing)
            If view Is Nothing OrElse Not AttachedViews.Add(view) Then Return
            If expandBranch IsNot Nothing Then ExpandBranchActions(view) = expandBranch
            If collapseBranch IsNot Nothing Then CollapseBranchActions(view) = collapseBranch
            AddHandler view.KeyDown, AddressOf View_KeyDown
            AddHandler view.Click, AddressOf View_Click
            AddHandler view.DoubleClick, AddressOf View_DoubleClick
            AddHandler view.MouseDown, AddressOf View_MouseDown
            AddHandler view.MouseMove, AddressOf View_MouseMove
            AddHandler view.MouseUp, AddressOf View_MouseUp
            AddHandler view.MouseLeave, AddressOf View_MouseLeave

            If view.GridControl IsNot Nothing Then
                AddHandler view.GridControl.SizeChanged, Sub()
                    Dim description = view.Columns.ColumnByFieldName("ItemDesc")
                    If description IsNot Nothing AndAlso view.GridControl.ClientSize.Width > 0 Then
                        description.MaxWidth = Math.Max(description.MinWidth, CInt(view.GridControl.ClientSize.Width * 0.4))
                    End If
                End Sub
                GridPresentation.SetResetLayout(view.GridControl, Sub()
                    view.BeginUpdate()
                    Try
                        'Undo an accidental row-edge drag as well as restoring text zoom.
                        view.RowHeight = -1
                        view.GroupRowHeight = -1
                    Finally
                        view.EndUpdate()
                    End Try
                End Sub)
                Dim analyserMenu As New ContextMenuStrip()
                ContextMenus(view) = analyserMenu
                view.GridControl.ContextMenuStrip = analyserMenu
                AddHandler analyserMenu.Opening, AddressOf ContextMenu_Opening
            End If
        End Sub

        Public Function ProcessClick(ByVal view As CustomGridView,
                                     ByVal e As MouseEventArgs) As Boolean
            If view Is Nothing OrElse e Is Nothing OrElse e.Button <> MouseButtons.Left Then Return False
            If SuppressNextClick Then
                SuppressNextClick = False
                Return True
            End If

            Dim hitInfo As GridHitInfo = view.CalcHitInfo(e.Location)
            If hitInfo Is Nothing Then Return False
            Dim rowHandle As Integer = hitInfo.RowHandle
            Dim kind As AnalyserVirtualRowKind
            If Not TryGetVirtualRow(view, hitInfo, rowHandle, kind) Then
                If (Control.ModifierKeys And (Keys.Control Or Keys.Shift)) = Keys.None Then ClearVirtualSelections(view)
                Return False
            End If

            Dim modifiers As Keys = Control.ModifierKeys
            Dim isControlClick As Boolean = (modifiers And Keys.Control) = Keys.Control
            Dim isShiftClick As Boolean = (modifiers And Keys.Shift) = Keys.Shift
            If Not isControlClick AndAlso Not isShiftClick Then
                Clear()
                view.ClearSelection()
            End If

            Dim selectedColumn As GridColumn = GetColumnAtX(view, e.X)
            If selectedColumn Is Nothing Then
                selectedColumn = GetSelectableVisibleColumns(view).FirstOrDefault()
            End If
            If selectedColumn Is Nothing Then Return False

            Dim existing As VisualRowSelection = FindSelection(view, rowHandle, kind)
            If existing IsNot Nothing AndAlso isControlClick AndAlso
               existing.Fields.Contains(selectedColumn.FieldName) Then
                existing.Fields.Remove(selectedColumn.FieldName)
                If existing.Fields.Count = 0 Then Selections.Remove(existing)
            Else
                If existing Is Nothing Then
                    existing = New VisualRowSelection With {
                        .View = view, .RowHandle = rowHandle, .Kind = kind, .OrderY = e.Y
                    }
                    Selections.Add(existing)
                End If
                existing.WholeRow = False
                existing.Fields.Add(selectedColumn.FieldName)
            End If
            InvalidateAttachedViews()
            Return kind = AnalyserVirtualRowKind.GroupFooter OrElse isControlClick OrElse isShiftClick
        End Function

        Public Function IsSelected(ByVal view As CustomGridView,
                                   ByVal rowHandle As Integer,
                                   ByVal kind As AnalyserVirtualRowKind) As Boolean
            Dim selection As VisualRowSelection = FindSelection(view, rowHandle, kind)
            Return selection IsNot Nothing AndAlso selection.WholeRow
        End Function

        Public Function IsCellSelected(ByVal view As CustomGridView,
                                       ByVal rowHandle As Integer,
                                       ByVal kind As AnalyserVirtualRowKind,
                                       ByVal column As GridColumn) As Boolean
            If column Is Nothing Then Return False
            Dim selection As VisualRowSelection = FindSelection(view, rowHandle, kind)
            Return selection IsNot Nothing AndAlso selection.Fields.Contains(column.FieldName)
        End Function

        Public Function HasRowSelection(ByVal view As CustomGridView,
                                        ByVal rowHandle As Integer,
                                        ByVal kind As AnalyserVirtualRowKind) As Boolean
            Return FindSelection(view, rowHandle, kind) IsNot Nothing
        End Function

        Public Function IsHotTracked(ByVal view As CustomGridView,
                                     ByVal rowHandle As Integer,
                                     ByVal kind As AnalyserVirtualRowKind) As Boolean
            Return HasHover AndAlso Object.ReferenceEquals(HoverView, view) AndAlso
                   HoverRowHandle = rowHandle AndAlso HoverKind = kind
        End Function

        Public Sub DrawSelectedCells(ByVal cache As GraphicsCache,
                                     ByVal view As CustomGridView,
                                     ByVal rowHandle As Integer,
                                     ByVal kind As AnalyserVirtualRowKind,
                                     ByVal rowBounds As Rectangle)
            Dim selection As VisualRowSelection = FindSelection(view, rowHandle, kind)
            If selection Is Nothing OrElse selection.WholeRow Then Return
            Dim info As GridViewInfo = TryCast(view.GetViewInfo(), GridViewInfo)
            If info Is Nothing Then Return
            Using brush As New SolidBrush(Color.FromArgb(165, Color.Wheat))
                For Each column As GridColumn In GetSelectableVisibleColumns(view)
                    If Not selection.Fields.Contains(column.FieldName) Then Continue For
                    Dim columnInfo As GridColumnInfoArgs = info.ColumnsInfo(column)
                    If columnInfo Is Nothing Then Continue For
                    Dim selectedBounds As Rectangle = Rectangle.Intersect(
                        rowBounds, New Rectangle(columnInfo.Bounds.X, rowBounds.Y,
                                                 columnInfo.Bounds.Width, rowBounds.Height))
                    If selectedBounds.Width > 0 AndAlso selectedBounds.Height > 0 Then
                        cache.FillRectangle(brush, selectedBounds)
                    End If
                Next
            End Using
        End Sub

        Public Sub Clear()
            Dim hadSelection As Boolean = Selections.Count > 0
            Selections.Clear()
            DragCandidate = False
            Dragging = False
            DragView = Nothing
            DragAnchorValid = False
            If hadSelection Then InvalidateAttachedViews()
        End Sub

        Private Sub View_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
            Dim view As CustomGridView = TryCast(sender, CustomGridView)
            If view Is Nothing OrElse e.Button <> MouseButtons.Left Then Return
            Dim hitInfo As GridHitInfo = view.CalcHitInfo(e.Location)
            SuppressNativeGroupActionValue =
                hitInfo IsNot Nothing AndAlso hitInfo.HitTest = GridHitTest.Row AndAlso
                view.IsGroupRow(hitInfo.RowHandle)
            Dim anchorRowHandle As Integer
            Dim anchorKind As AnalyserVirtualRowKind
            If Not TryGetVisualRow(view, hitInfo, anchorRowHandle, anchorKind) Then Return
            DragView = view
            DragStart = e.Location
            DragCurrent = e.Location
            DragCandidate = True
            Dragging = False
            DragIsAdditive = (Control.ModifierKeys And Keys.Control) = Keys.Control
            DragAnchorRowHandle = anchorRowHandle
            DragAnchorKind = anchorKind
            DragAnchorValid = True
        End Sub

        Private Sub View_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
            Dim view As CustomGridView = TryCast(sender, CustomGridView)
            If view Is Nothing Then Return
            UpdateHover(view, e.Location)
            If Not DragCandidate OrElse Not Object.ReferenceEquals(view, DragView) OrElse
               (e.Button And MouseButtons.Left) <> MouseButtons.Left Then Return
            If Not Dragging Then
                Dim dragSize As Size = SystemInformation.DragSize
                Dim dragThreshold As New Rectangle(
                    DragStart.X - (dragSize.Width \ 2), DragStart.Y - (dragSize.Height \ 2),
                    dragSize.Width, dragSize.Height)
                If dragThreshold.Contains(e.Location) Then Return
                Dragging = True
                SuppressNextClick = True
                If view.GridControl IsNot Nothing Then view.GridControl.Capture = True
            End If
            DragCurrent = ClampToView(view, e.Location)
            ApplyDragRectangle(view)
        End Sub

        Private Sub View_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
            Dim view As CustomGridView = TryCast(sender, CustomGridView)
            If view Is Nothing OrElse Not Object.ReferenceEquals(view, DragView) Then Return
            If Dragging Then
                DragCurrent = ClampToView(view, e.Location)
                ApplyDragRectangle(view)
                SuppressNextClick = True
            End If
            If view.GridControl IsNot Nothing Then view.GridControl.Capture = False
            DragCandidate = False
            Dragging = False
            DragView = Nothing
            DragAnchorValid = False
        End Sub

        Private Sub View_MouseLeave(ByVal sender As Object, ByVal e As EventArgs)
            Dim view As CustomGridView = TryCast(sender, CustomGridView)
            If view Is Nothing OrElse Not HasHover OrElse
               Not Object.ReferenceEquals(HoverView, view) Then Return
            HasHover = False
            HoverView = Nothing
            InvalidateAttachedViews()
        End Sub

        Private Sub UpdateHover(ByVal view As CustomGridView, ByVal location As Point)
            Dim hitInfo As GridHitInfo = view.CalcHitInfo(location)
            Dim rowHandle As Integer
            Dim kind As AnalyserVirtualRowKind
            Dim found As Boolean = TryGetVisualRow(view, hitInfo, rowHandle, kind)
            If found AndAlso HasHover AndAlso Object.ReferenceEquals(HoverView, view) AndAlso
               HoverRowHandle = rowHandle AndAlso HoverKind = kind Then Return
            If Not found AndAlso (Not HasHover OrElse Not Object.ReferenceEquals(HoverView, view)) Then Return

            HasHover = found
            HoverView = If(found, view, Nothing)
            If found Then
                HoverRowHandle = rowHandle
                HoverKind = kind
            End If
            InvalidateAttachedViews()
        End Sub

        Private Sub ApplyDragRectangle(ByVal view As CustomGridView)
            Dim selectionRectangle As Rectangle = NormaliseRectangle(DragStart, DragCurrent)
            If selectionRectangle.Width < 1 OrElse selectionRectangle.Height < 1 Then Return
            If Not DragIsAdditive Then
                ClearVirtualSelections(view)
                view.ClearSelection()
            End If

            Dim columns As List(Of GridColumn) = GetColumnsIntersecting(view, selectionRectangle)
            If columns.Count = 0 Then Return
            Dim lastKey As String = Nothing
            For y As Integer = selectionRectangle.Top To selectionRectangle.Bottom
                Dim sampleX As Integer = Math.Max(selectionRectangle.Left,
                    Math.Min(selectionRectangle.Right - 1, GetColumnCentre(view, columns(0))))
                Dim hitInfo As GridHitInfo = view.CalcHitInfo(New Point(sampleX, y))
                Dim rowHandle As Integer
                Dim kind As AnalyserVirtualRowKind
                If Not TryGetVisualRow(view, hitInfo, rowHandle, kind) Then Continue For
                Dim key As String = rowHandle.ToString() & ":" & CInt(kind).ToString()
                If key = lastKey Then Continue For
                lastKey = key
                AddDraggedRow(view, rowHandle, kind, columns, y)
            Next
            If DragAnchorValid Then
                AddDraggedRow(view, DragAnchorRowHandle, DragAnchorKind, columns, DragStart.Y)
            End If
            InvalidateAttachedViews()
        End Sub

        Private Sub AddDraggedRow(ByVal view As CustomGridView,
                                  ByVal rowHandle As Integer,
                                  ByVal kind As AnalyserVirtualRowKind,
                                  ByVal columns As IEnumerable(Of GridColumn),
                                  ByVal orderY As Integer)
            If kind = AnalyserVirtualRowKind.DataRow Then
                For Each column As GridColumn In columns
                    view.SelectCell(rowHandle, column)
                Next
            End If
            Dim selection As VisualRowSelection = FindSelection(view, rowHandle, kind)
            If selection Is Nothing Then
                selection = New VisualRowSelection With {
                    .View = view, .RowHandle = rowHandle, .Kind = kind, .OrderY = orderY
                }
                Selections.Add(selection)
            End If
            selection.WholeRow = False
            For Each column As GridColumn In columns
                selection.Fields.Add(column.FieldName)
            Next
        End Sub

        Private Shared Function NormaliseRectangle(ByVal startPoint As Point,
                                                   ByVal endPoint As Point) As Rectangle
            Return Rectangle.FromLTRB(Math.Min(startPoint.X, endPoint.X),
                Math.Min(startPoint.Y, endPoint.Y), Math.Max(startPoint.X, endPoint.X) + 1,
                Math.Max(startPoint.Y, endPoint.Y) + 1)
        End Function

        Private Shared Function ClampToView(ByVal view As CustomGridView,
                                            ByVal point As Point) As Point
            If view.GridControl Is Nothing Then Return point
            Return New Point(Math.Max(0, Math.Min(view.GridControl.ClientSize.Width - 1, point.X)),
                             Math.Max(0, Math.Min(view.GridControl.ClientSize.Height - 1, point.Y)))
        End Function

        Private Shared Function GetColumnsIntersecting(ByVal view As CustomGridView,
                                                       ByVal bounds As Rectangle) As List(Of GridColumn)
            Dim result As New List(Of GridColumn)()
            Dim info As GridViewInfo = TryCast(view.GetViewInfo(), GridViewInfo)
            If info Is Nothing Then Return result
            For Each column As GridColumn In GetSelectableVisibleColumns(view)
                Dim columnInfo As GridColumnInfoArgs = info.ColumnsInfo(column)
                If columnInfo IsNot Nothing AndAlso columnInfo.Bounds.Right > bounds.Left AndAlso
                   columnInfo.Bounds.Left < bounds.Right Then result.Add(column)
            Next
            Return result
        End Function

        Private Shared Function GetColumnCentre(ByVal view As CustomGridView,
                                                ByVal column As GridColumn) As Integer
            Dim info As GridViewInfo = TryCast(view.GetViewInfo(), GridViewInfo)
            If info Is Nothing Then Return 0
            Dim columnInfo As GridColumnInfoArgs = info.ColumnsInfo(column)
            If columnInfo Is Nothing Then Return 0
            Return columnInfo.Bounds.Left + (columnInfo.Bounds.Width \ 2)
        End Function

        Private Shared Function GetColumnAtX(ByVal view As CustomGridView,
                                             ByVal x As Integer) As GridColumn
            Dim info As GridViewInfo = TryCast(view.GetViewInfo(), GridViewInfo)
            If info Is Nothing Then Return Nothing
            For Each column As GridColumn In GetSelectableVisibleColumns(view)
                Dim columnInfo As GridColumnInfoArgs = info.ColumnsInfo(column)
                If columnInfo IsNot Nothing AndAlso x >= columnInfo.Bounds.Left AndAlso
                   x < columnInfo.Bounds.Right Then Return column
            Next
            Return Nothing
        End Function

        Private Shared Function GetSelectableVisibleColumns(ByVal view As CustomGridView) As List(Of GridColumn)
            Return view.VisibleColumns.Cast(Of GridColumn)().
                Where(Function(column) column.Visible AndAlso Not IsInternalColumn(column.FieldName)).
                OrderBy(Function(column) column.VisibleIndex).ToList()
        End Function

        Private Shared Function IsInternalColumn(ByVal fieldName As String) As Boolean
            Return String.Equals(fieldName, "Level 1", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(fieldName, "Level 2", StringComparison.OrdinalIgnoreCase) OrElse
                   fieldName.StartsWith("Ordered", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsSelectableHit(ByVal view As CustomGridView,
                                                ByVal hitInfo As GridHitInfo) As Boolean
            Dim rowHandle As Integer
            Dim kind As AnalyserVirtualRowKind
            Return TryGetVisualRow(view, hitInfo, rowHandle, kind)
        End Function

        Private Shared Function TryGetVisualRow(ByVal view As CustomGridView,
                                                ByVal hitInfo As GridHitInfo,
                                                ByRef rowHandle As Integer,
                                                ByRef kind As AnalyserVirtualRowKind) As Boolean
            If hitInfo Is Nothing Then Return False
            If TryGetVirtualRow(view, hitInfo, rowHandle, kind) Then Return True
            rowHandle = hitInfo.RowHandle
            If rowHandle >= 0 AndAlso Not view.IsGroupRow(rowHandle) Then
                kind = AnalyserVirtualRowKind.DataRow
                Return True
            End If
            Return False
        End Function

        Private Shared Function TryGetVirtualRow(ByVal view As CustomGridView,
                                                 ByVal hitInfo As GridHitInfo,
                                                 ByRef rowHandle As Integer,
                                                 ByRef kind As AnalyserVirtualRowKind) As Boolean
            If hitInfo Is Nothing Then Return False
            rowHandle = hitInfo.RowHandle
            If hitInfo.HitTest = GridHitTest.RowFooter Then
                If Not view.IsGroupRow(rowHandle) Then rowHandle = view.GetParentRowHandle(rowHandle)
                If view.IsGroupRow(rowHandle) Then
                    kind = AnalyserVirtualRowKind.GroupFooter
                    Return True
                End If
            ElseIf view.IsGroupRow(rowHandle) Then
                kind = AnalyserVirtualRowKind.GroupHeading
                Return True
            End If
            Return False
        End Function

        Private Function FindSelection(ByVal view As CustomGridView,
                                       ByVal rowHandle As Integer,
                                       ByVal kind As AnalyserVirtualRowKind) As VisualRowSelection
            Return Selections.FirstOrDefault(
                Function(selection) Object.ReferenceEquals(selection.View, view) AndAlso
                    selection.RowHandle = rowHandle AndAlso selection.Kind = kind)
        End Function

        Private Sub ClearVirtualSelections(ByVal view As CustomGridView)
            Selections.RemoveAll(Function(item) Object.ReferenceEquals(item.View, view))
        End Sub

        Private Sub View_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            Dim view As CustomGridView = TryCast(sender, CustomGridView)
            If view Is Nothing Then Return
            If e.Control AndAlso e.KeyCode = Keys.C AndAlso HasSelection(view) Then
                CopySelections(view)
                e.Handled = True
                e.SuppressKeyPress = True
            ElseIf e.KeyCode = Keys.Escape AndAlso HasSelection(view) Then
                Clear()
                view.ClearSelection()
                e.Handled = True
                e.SuppressKeyPress = True
            End If
        End Sub

        Private Sub View_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim view As CustomGridView = TryCast(sender, CustomGridView)
            Dim mouseEvent As MouseEventArgs = TryCast(e, MouseEventArgs)
            If view Is Nothing OrElse mouseEvent Is Nothing Then Return

            ProcessClick(view, mouseEvent)
            If mouseEvent.Button <> MouseButtons.Left Then Return

            Dim hitInfo As GridHitInfo = view.CalcHitInfo(mouseEvent.Location)
            If hitInfo Is Nothing OrElse hitInfo.HitTest <> GridHitTest.Row OrElse
               Not view.IsGroupRow(hitInfo.RowHandle) Then Return

            PendingGroupView = view
            PendingGroupRowHandle = hitInfo.RowHandle
            HasPendingGroupClick = True
            GroupClickTimer.Stop()
            GroupClickTimer.Start()
        End Sub

        Private Sub View_DoubleClick(ByVal sender As Object, ByVal e As EventArgs)
            Dim view As CustomGridView = TryCast(sender, CustomGridView)
            If view Is Nothing OrElse Not HasPendingGroupClick OrElse
               Not Object.ReferenceEquals(view, PendingGroupView) Then Return

            GroupClickTimer.Stop()
            Dim rowHandle As Integer = PendingGroupRowHandle
            HasPendingGroupClick = False
            PendingGroupView = Nothing
            Try
                If Not view.IsGroupRow(rowHandle) Then Return

                If view.GetRowExpanded(rowHandle) Then
                    Dim collapseAction As Action(Of CustomGridView, Integer) = Nothing
                    If CollapseBranchActions.TryGetValue(view, collapseAction) Then
                        RunManagedGroupAction(Sub() collapseAction(view, rowHandle))
                    End If
                Else
                    Dim expandAction As Action(Of CustomGridView, Integer) = Nothing
                    If ExpandBranchActions.TryGetValue(view, expandAction) Then
                        RunManagedGroupAction(Sub() expandAction(view, rowHandle))
                    End If
                End If
            Finally
                SuppressNativeGroupActionValue = False
            End Try
        End Sub

        Private Sub GroupClickTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
            GroupClickTimer.Stop()
            If Not HasPendingGroupClick Then Return

            Dim view As CustomGridView = PendingGroupView
            Dim rowHandle As Integer = PendingGroupRowHandle
            HasPendingGroupClick = False
            PendingGroupView = Nothing
            Try
                If view Is Nothing OrElse Not view.IsGroupRow(rowHandle) Then Return
                RunManagedGroupAction(
                    Sub() view.SetRowExpanded(
                        rowHandle, Not view.GetRowExpanded(rowHandle), False))
            Finally
                SuppressNativeGroupActionValue = False
            End Try
        End Sub

        Private Sub RunManagedGroupAction(ByVal action As Action)
            If action Is Nothing Then Return
            ApplyingGroupAction = True
            Try
                action()
            Finally
                ApplyingGroupAction = False
            End Try
        End Sub

        Private Sub View_PopupMenuShowing(ByVal sender As Object,
                                          ByVal e As PopupMenuShowingEventArgs)
            Dim view As CustomGridView = TryCast(sender, CustomGridView)
            If view Is Nothing OrElse e Is Nothing OrElse e.Menu Is Nothing Then Return

            'The analyser deliberately exposes only its rectangular clipboard
            'commands; filtering and other standard grid actions are not relevant.
            e.Menu.Items.Clear()
            Dim hasSelectionValue As Boolean = HasSelection(view)
            If hasSelectionValue Then
                Dim copyItem As New DXMenuItem("Copy selected analyser cells")
                AddHandler copyItem.Click, Sub() CopySelections(view)
                e.Menu.Items.Add(copyItem)

                Dim copyWithPeriodItem As New DXMenuItem("Copy with year/period")
                AddHandler copyWithPeriodItem.Click, Sub() CopySelections(view, True)
                e.Menu.Items.Add(copyWithPeriodItem)
            End If

            Dim hitInfo As GridHitInfo = e.HitInfo
            If hitInfo Is Nothing OrElse Not view.IsGroupRow(hitInfo.RowHandle) Then Return

            Dim rowHandle As Integer = hitInfo.RowHandle
            If view.GetRowExpanded(rowHandle) Then
                Dim collapseAction As Action(Of CustomGridView, Integer) = Nothing
                If CollapseBranchActions.TryGetValue(view, collapseAction) Then
                    Dim collapseItem As New DXMenuItem("Collapse branch")
                    collapseItem.BeginGroup = hasSelectionValue
                    AddHandler collapseItem.Click, Sub() collapseAction(view, rowHandle)
                    e.Menu.Items.Add(collapseItem)
                End If
            Else
                Dim expandAction As Action(Of CustomGridView, Integer) = Nothing
                If ExpandBranchActions.TryGetValue(view, expandAction) Then
                    Dim expandItem As New DXMenuItem("Expand branch")
                    expandItem.BeginGroup = hasSelectionValue
                    AddHandler expandItem.Click, Sub() expandAction(view, rowHandle)
                    e.Menu.Items.Add(expandItem)
                End If
            End If
        End Sub

        Private Sub ContextMenu_Opening(ByVal sender As Object,
                                        ByVal e As CancelEventArgs)
            Dim analyserMenu As ContextMenuStrip = TryCast(sender, ContextMenuStrip)
            If analyserMenu Is Nothing Then Return
            Dim view As CustomGridView =
                ContextMenus.FirstOrDefault(
                    Function(item) Object.ReferenceEquals(item.Value, analyserMenu)).Key
            If view Is Nothing OrElse view.GridControl Is Nothing Then
                e.Cancel = True
                Return
            End If

            analyserMenu.Items.Clear()
            Dim hasSelectionValue As Boolean = HasSelection(view)
            If hasSelectionValue Then
                Dim copyItem As New ToolStripMenuItem("Copy selected analyser cells")
                AddHandler copyItem.Click, Sub() CopySelections(view)
                analyserMenu.Items.Add(copyItem)

                Dim copyWithPeriodItem As New ToolStripMenuItem("Copy with year/period")
                AddHandler copyWithPeriodItem.Click, Sub() CopySelections(view, True)
                analyserMenu.Items.Add(copyWithPeriodItem)
            End If

            Dim clientPoint As Point =
                view.GridControl.PointToClient(Control.MousePosition)
            Dim hitInfo As GridHitInfo = view.CalcHitInfo(clientPoint)
            If hitInfo IsNot Nothing AndAlso view.IsGroupRow(hitInfo.RowHandle) Then
                If hasSelectionValue Then analyserMenu.Items.Add(New ToolStripSeparator())
                Dim rowHandle As Integer = hitInfo.RowHandle
                If view.GetRowExpanded(rowHandle) Then
                    Dim collapseAction As Action(Of CustomGridView, Integer) = Nothing
                    If CollapseBranchActions.TryGetValue(view, collapseAction) Then
                        Dim collapseItem As New ToolStripMenuItem("Collapse branch")
                        AddHandler collapseItem.Click,
                            Sub() RunManagedGroupAction(
                                Sub() collapseAction(view, rowHandle))
                        analyserMenu.Items.Add(collapseItem)
                    End If
                Else
                    Dim expandAction As Action(Of CustomGridView, Integer) = Nothing
                    If ExpandBranchActions.TryGetValue(view, expandAction) Then
                        Dim expandItem As New ToolStripMenuItem("Expand branch")
                        AddHandler expandItem.Click,
                            Sub() RunManagedGroupAction(
                                Sub() expandAction(view, rowHandle))
                        analyserMenu.Items.Add(expandItem)
                    End If
                End If
            End If
            'This menu is rebuilt on every open; restore shared presentation actions
            'after clearing/recreating the analyser's selection/branch commands.
            GridPresentation.SetMenuTarget(analyserMenu, view.GridControl)
            GridPresentation.AddMenu(view.GridControl)
            e.Cancel = analyserMenu.Items.Count = 0
        End Sub

        Private Function HasSelection(ByVal view As CustomGridView) As Boolean
            SynchroniseNativeCellSelections(view)
            Return Selections.Any(Function(item) Object.ReferenceEquals(item.View, view))
        End Function

        Private Sub SynchroniseNativeCellSelections(ByVal view As CustomGridView)
            If view Is Nothing Then Return

            For Each selectedCell In view.GetSelectedCells()
                If selectedCell.RowHandle < 0 OrElse selectedCell.Column Is Nothing OrElse
                   IsInternalColumn(selectedCell.Column.FieldName) Then Continue For

                Dim selection As VisualRowSelection =
                    FindSelection(view, selectedCell.RowHandle, AnalyserVirtualRowKind.DataRow)
                If selection Is Nothing Then
                    selection = New VisualRowSelection With {
                        .View = view,
                        .RowHandle = selectedCell.RowHandle,
                        .Kind = AnalyserVirtualRowKind.DataRow,
                        .OrderY = view.GetVisibleIndex(selectedCell.RowHandle)
                    }
                    Selections.Add(selection)
                End If
                selection.WholeRow = False
                selection.Fields.Add(selectedCell.Column.FieldName)
            Next
        End Sub

        Private Sub CopySelections(ByVal view As CustomGridView,
                                   Optional ByVal includeYearAndPeriod As Boolean = False)
            Dim selectedRows As List(Of VisualRowSelection) =
                Selections.Where(Function(item) Object.ReferenceEquals(item.View, view)).
                    OrderBy(Function(item) item.OrderY).ThenBy(Function(item) CInt(item.Kind)).ToList()
            If selectedRows.Count = 0 Then Return
            Dim selectedFields As New HashSet(Of String)(StringComparer.Ordinal)
            For Each row As VisualRowSelection In selectedRows
                selectedFields.UnionWith(row.Fields)
            Next
            Dim columns As List(Of GridColumn) = GetSelectableVisibleColumns(view).
                Where(Function(column) selectedFields.Contains(column.FieldName)).ToList()
            If columns.Count = 0 Then Return
            Dim clipboardRows As New List(Of String)()
            If includeYearAndPeriod Then
                clipboardRows.AddRange(BuildColumnHeadingRows(columns))
            End If
            For Each selection As VisualRowSelection In selectedRows
                clipboardRows.Add(BuildClipboardRow(selection, columns))
            Next
            Dim copiedText As String = String.Join(ControlChars.CrLf, clipboardRows)
            If includeYearAndPeriod Then
                SetClipboardTextWithHeaders(copiedText,
                                            clipboardRows.Count - selectedRows.Count, 0)
            Else
                Clipboard.SetText(copiedText, TextDataFormat.UnicodeText)
            End If
        End Sub

        Private Shared Function BuildColumnHeadingRows(
            ByVal columns As List(Of GridColumn)) As IEnumerable(Of String)

            Dim yearHeadings(columns.Count - 1) As String
            Dim periodHeadings(columns.Count - 1) As String
            Dim hasPeriodHeading As Boolean

            For index As Integer = 0 To columns.Count - 1
                Dim caption As String = If(columns(index).Caption, String.Empty).
                    Replace(ControlChars.CrLf, ControlChars.Lf).
                    Replace(ControlChars.Cr, ControlChars.Lf)
                Dim captionParts As String() = caption.Split(New Char() {ControlChars.Lf},
                                                             StringSplitOptions.None)
                yearHeadings(index) = captionParts(0).Trim()
                If captionParts.Length > 1 Then
                    periodHeadings(index) = String.Join(" ", captionParts.Skip(1)).Trim()
                    hasPeriodHeading = hasPeriodHeading OrElse periodHeadings(index).Length > 0
                End If
            Next

            Dim result As New List(Of String) From {
                String.Join(ControlChars.Tab, yearHeadings)
            }
            If hasPeriodHeading Then result.Add(String.Join(ControlChars.Tab, periodHeadings))
            Return result
        End Function

        Private Shared Function BuildClipboardRow(ByVal selection As VisualRowSelection,
                                                  ByVal columns As List(Of GridColumn)) As String
            Dim cells(columns.Count - 1) As String
            If selection.Kind = AnalyserVirtualRowKind.DataRow Then
                For index As Integer = 0 To columns.Count - 1
                    cells(index) = selection.View.GetRowCellDisplayText(selection.RowHandle, columns(index))
                Next
                Return String.Join(ControlChars.Tab, cells)
            End If

            Dim firstVisibleColumn As GridColumn = GetSelectableVisibleColumns(selection.View).FirstOrDefault()
            If firstVisibleColumn IsNot Nothing Then
                Dim captionIndex As Integer = columns.FindIndex(
                    Function(column) Object.ReferenceEquals(column, firstVisibleColumn))
                If captionIndex >= 0 Then
                    cells(captionIndex) = GetGroupCaption(selection.View, selection.RowHandle)
                    If selection.Kind = AnalyserVirtualRowKind.GroupFooter Then cells(captionIndex) &= " Total"
                End If
            End If
            Dim values As Hashtable = selection.View.GetGroupSummaryValues(selection.RowHandle)
            If values IsNot Nothing Then
                For Each summary As Object In selection.View.GroupSummary
                    Dim groupSummary As DevExpress.XtraGrid.GridGroupSummaryItem =
                        TryCast(summary, DevExpress.XtraGrid.GridGroupSummaryItem)
                    If groupSummary Is Nothing OrElse groupSummary.ShowInGroupColumnFooter Is Nothing Then Continue For
                    Dim columnIndex As Integer = columns.FindIndex(
                        Function(column) Object.ReferenceEquals(column, groupSummary.ShowInGroupColumnFooter))
                    If columnIndex >= 0 AndAlso values.ContainsKey(groupSummary) Then
                        cells(columnIndex) = groupSummary.GetDisplayText(values(groupSummary), False)
                    End If
                Next
            End If
            Return String.Join(ControlChars.Tab, cells)
        End Function

        Private Shared Function GetGroupCaption(ByVal view As CustomGridView,
                                                ByVal rowHandle As Integer) As String
            Dim level As Integer = view.GetRowLevel(rowHandle)
            Dim caption As String = view.GetGroupRowDisplayText(rowHandle)
            If level >= 0 AndAlso level < view.GroupedColumns.Count Then
                Dim dataRowHandle As Integer = view.GetDataRowHandleByGroupRowHandle(rowHandle)
                caption = view.GetRowCellDisplayText(dataRowHandle, view.GroupedColumns(level))
            End If
            caption = If(caption, String.Empty).Trim()
            If caption.Length > 5 Then
                Dim prefix As Integer
                If Integer.TryParse(caption.Substring(0, 2), prefix) Then caption = caption.Substring(5)
            End If
            Return caption
        End Function

        Private Sub InvalidateAttachedViews()
            For Each view As CustomGridView In AttachedViews
                If view IsNot Nothing AndAlso view.GridControl IsNot Nothing AndAlso
                   Not view.GridControl.IsDisposed Then view.GridControl.Invalidate(True)
            Next
        End Sub

    End Class

End Namespace
