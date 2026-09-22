Option Infer On
Imports System.Drawing
Imports System.Windows.Forms
Imports Abovo
Imports Abovo.CustomGrid
Imports Abovo.DataObject
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.XtraVerticalGrid.Rows

Partial Public Class DataInterfaceTemplate
    Private NavigationPending As Boolean
    Private EnterNavigationVertical As Boolean
    Private ReadOnly StandaloneNavigationEditors As New HashSet(Of BaseEdit)
    Private Sub ConfigureEditorNavigation(editor As BaseEdit)
        editor.EnterMoveNextControl = False
        RemoveHandler editor.PreviewKeyDown, AddressOf NavigationEditorPreviewKeyDown
        AddHandler editor.PreviewKeyDown, AddressOf NavigationEditorPreviewKeyDown
        RemoveHandler editor.KeyDown, AddressOf NavigationEditorKeyDown
        AddHandler editor.KeyDown, AddressOf NavigationEditorKeyDown
    End Sub

    Private Sub NavigationEditorPreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs)
        Dim editor = TryCast(sender, BaseEdit)
        If editor IsNot Nothing AndAlso StandaloneNavigationEditors.Contains(editor) AndAlso e.KeyCode <> Keys.Enter Then
            Dim popup = TryCast(editor, PopupBaseEdit)
            If IsNavigationKey(e.KeyData) AndAlso (popup Is Nothing OrElse Not popup.IsPopupOpen) Then RememberEnterDirection(e.KeyData)
            'Do not turn native standalone Tab/dialog navigation into an input key.
            Return
        End If
        If IsNavigationKey(e.KeyData) Then e.IsInputKey = True
    End Sub

    Private Sub NavigationEditorKeyDown(sender As Object, e As KeyEventArgs)
        If e.Handled Then Return
        Dim standalone = TryCast(sender, BaseEdit)
        If standalone IsNot Nothing AndAlso StandaloneNavigationEditors.Contains(standalone) Then
            If NavigateStandalone(standalone, e.KeyData) Then
                e.Handled = True : e.SuppressKeyPress = True
            End If
            Return
        End If
        Dim host = DirectCast(sender, Control).Parent
        While host IsNot Nothing AndAlso Not TypeOf host Is GridControl AndAlso Not TypeOf host Is VGridControl
            host = host.Parent
        End While
        If NavigateGrid(host, e.KeyData) Then
            e.Handled = True : e.SuppressKeyPress = True
        End If
    End Sub

    Private Function CommitHeaderWorkbookChange(sender As Object, change As DataChangeEvent, previous As Object) As Boolean
        WorkbookPostingDepth += 1
        Try
            Dim result = ChangeMan.ProcessChangeByNRAddressing(change)
            If result.BError Then Throw New InvalidOperationException(result.StrResponseMessage)
            Return True
        Catch ex As Exception
            sender.EditValue = previous
            Diagnostics.Trace.WriteLine("[Header editor] " & ex.ToString())
            XtraMessageBox.Show(Me, "The value could not be accepted. " & ex.Message, "Funding / assumptions entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        Finally
            WorkbookPostingDepth = Math.Max(0, WorkbookPostingDepth - 1)
            QueueWorkbookRefresh()
        End Try
    End Function

    'Refresh is not navigation. Preserve logical cell and both scroll axes;
    'never call Focus here (which would pull focus from another open interface).
    Private Function CaptureNavigationPosition() As Action
        Dim restore As New List(Of Action)
        If GridControls IsNot Nothing Then
            For Each grid In GridControls
                If grid Is Nothing OrElse grid.IsDisposed Then Continue For
                Dim view = TryCast(grid.MainView, GridView)
                If view Is Nothing Then Continue For
                Dim column = view.FocusedColumn, row = view.FocusedRowHandle
                Dim top = view.TopRowIndex, left = view.LeftCoord
                restore.Add(Sub()
                                If grid.IsDisposed OrElse view.GridControl IsNot grid Then Return
                                If column IsNot Nothing AndAlso column.View Is view Then view.FocusedColumn = column
                                If view.IsValidRowHandle(row) Then view.FocusedRowHandle = row
                                view.TopRowIndex = top : view.LeftCoord = left
                            End Sub)
            Next
        End If
        If VertGridControls IsNot Nothing Then
            For Each grid In VertGridControls
                If grid Is Nothing OrElse grid.IsDisposed Then Continue For
                Dim row = grid.FocusedRow, record = grid.FocusedRecord
                Dim top = grid.TopVisibleRowIndex, left = grid.LeftVisibleRecord
                restore.Add(Sub()
                                If grid.IsDisposed Then Return
                                If row IsNot Nothing AndAlso row.Grid Is grid Then grid.FocusedRow = row
                                If record >= 0 AndAlso record < grid.RecordCount Then grid.FocusedRecord = record
                                grid.TopVisibleRowIndex = top : grid.LeftVisibleRecord = left
                            End Sub)
            Next
        End If
        For Each panel In FindChildControls(Of ScrollableControl)(Me)
            If TypeOf panel Is XtraScrollableControl OrElse Not panel.Visible OrElse Not panel.AutoScroll Then Continue For
            Dim position = panel.AutoScrollPosition
            restore.Add(Sub()
                            If Not panel.IsDisposed Then panel.AutoScrollPosition = New Point(-position.X, -position.Y)
                        End Sub)
        Next
        For Each panel In FindChildControls(Of DevExpress.XtraEditors.XtraScrollableControl)(Me)
            If Not panel.Visible OrElse Not panel.AutoScroll Then Continue For
            Dim position = panel.AutoScrollPosition
            restore.Add(Sub()
                            If Not panel.IsDisposed Then panel.AutoScrollPosition = New Point(-position.X, -position.Y)
                        End Sub)
        Next
        Return Sub()
                   For Each action In restore
                       action()
                   Next
               End Sub
    End Function

    Private Class EditorPosition
        Public Host As Control
        Public Standalone As BaseEdit
        Public View As GridView
        Public Column As GridColumn
        Public RowHandle As Integer
        Public VRow As EditorRow
        Public Record As Integer
        Public Header As ColumnInplaceEditorHelper
        Public VHeader As VGridRowInplaceEditorHelper
        Public X As Integer
        Public Y As Integer
    End Class

    Private Shared Function IsNavigationKey(keys As Keys) As Boolean
        If (keys And (Keys.Control Or Keys.Alt)) <> Keys.None Then Return False
        Select Case keys And Keys.KeyCode
            Case Keys.Tab, Keys.Enter, Keys.Right, Keys.Left, Keys.Up, Keys.Down
                Return (keys And Keys.Shift) = Keys.None OrElse {Keys.Tab, Keys.Enter}.Contains(keys And Keys.KeyCode)
        End Select
        Return False
    End Function

    Private Function NavigateGrid(host As Control, keys As Keys, Optional header As ColumnInplaceEditorHelper = Nothing,
                                  Optional vHeader As VGridRowInplaceEditorHelper = Nothing) As Boolean
        If InterfaceResourcesReleased OrElse IsDisposed OrElse Disposing OrElse Not IsHandleCreated Then Return False
        If Not IsNavigationKey(keys) OrElse host Is Nothing OrElse host.IsDisposed OrElse IsAuthoringPreview Then Return False
        If NavigationPending Then Return True
        Dim grid = TryCast(host, GridControl)
        Dim view = If(grid Is Nothing, Nothing, TryCast(grid.FocusedView, GridView))
        Dim vertical = TryCast(host, VGridControl)
        If vertical IsNot Nothing AndAlso TypeOf vertical.FocusedRow Is MultiEditorRow Then Return False
        Dim editor = If(view IsNot Nothing, view.ActiveEditor, If(vertical IsNot Nothing, vertical.ActiveEditor, Nothing))
        Dim popup = TryCast(editor, PopupBaseEdit)
        If popup IsNot Nothing AndAlso popup.IsPopupOpen Then Return False
        Dim positions = NavigationPositions()
        Dim origin = positions.FirstOrDefault(Function(p) p.Host Is host AndAlso
            If(header IsNot Nothing, p.Header Is header,
               If(vHeader IsNot Nothing, p.VHeader Is vHeader,
                  If(view IsNot Nothing, p.Header Is Nothing AndAlso p.Column Is view.FocusedColumn AndAlso p.RowHandle = view.FocusedRowHandle,
                     vertical IsNot Nothing AndAlso p.VHeader Is Nothing AndAlso p.VRow Is vertical.FocusedRow AndAlso p.Record = vertical.FocusedRecord))))
        If origin Is Nothing Then Return False
        If editor IsNot Nothing AndAlso Not editor.DoValidate(PopupCloseMode.Normal) Then Return True
        If view IsNot Nothing Then
            If editor IsNot Nothing AndAlso (Not view.PostEditor() OrElse Not view.UpdateCurrentRow()) Then Return True
            view.CloseEditor()
        ElseIf vertical IsNot Nothing Then
            If editor IsNot Nothing AndAlso Not vertical.PostEditor() Then Return True
            vertical.CloseEditor()
        End If
        RememberEnterDirection(keys)
        QueueEditorNavigation(origin, keys)
        Return True
    End Function

    Private Sub RememberEnterDirection(keys As Keys)
        Select Case keys And Keys.KeyCode
            Case Keys.Up, Keys.Down : EnterNavigationVertical = True
            Case Keys.Left, Keys.Right, Keys.Tab : EnterNavigationVertical = False
        End Select
    End Sub

    Private Sub QueueEditorNavigation(origin As EditorPosition, keys As Keys)
        Dim host = origin.Host
        Dim isEnter = (keys And Keys.KeyCode) = Keys.Enter
        Dim vertical = EnterNavigationVertical
        NavigationPending = True
        BeginInvoke(New MethodInvoker(Sub()
            NavigationPending = False
            If InterfaceResourcesReleased OrElse IsDisposed OrElse Disposing OrElse host.IsDisposed OrElse Not host.Visible Then Return
            'Re-evaluate locks after workbook calculation/rules and queued refresh.
            Dim candidates = If(isEnter, EnterNavigationPositions(), NavigationPositions())
            Dim target = If(isEnter, FindEnterTarget(origin, candidates, vertical, (keys And Keys.Shift) <> Keys.None),
                            FindNavigationTarget(origin, candidates, keys))
            If target Is Nothing Then target = candidates.FirstOrDefault(Function(p) p.Host Is host AndAlso p.X = origin.X AndAlso p.Y = origin.Y)
            If target IsNot Nothing Then ActivateEditorPosition(target)
        End Sub))
    End Sub

    Private Sub RegisterStandaloneNavigation(editor As BaseEdit)
        If editor Is Nothing OrElse Not TypeOf editor.Tag Is SingleCellDataTag Then Return
        ConfigureEditorNavigation(editor)
        If StandaloneNavigationEditors.Add(editor) Then AddHandler editor.Disposed, AddressOf StandaloneNavigationDisposed
    End Sub

    Private Sub StandaloneNavigationDisposed(sender As Object, e As EventArgs)
        StandaloneNavigationEditors.Remove(DirectCast(sender, BaseEdit))
    End Sub

    Private Function CanNavigateStandalone(editor As BaseEdit) As Boolean
        If editor Is Nothing OrElse editor.IsDisposed OrElse Not editor.Visible OrElse Not editor.Enabled OrElse
           Not editor.TabStop OrElse editor.Properties.ReadOnly Then Return False
        Dim tag = TryCast(editor.Tag, SingleCellDataTag)
        Return tag IsNot Nothing AndAlso Not tag.IsCalculated AndAlso tag.TargetWorksheet IsNot Nothing AndAlso
            Not tag.TargetWorksheet.Cells(tag.TargetCell).Protection.Locked
    End Function

    Private Function NavigateStandalone(editor As BaseEdit, keys As Keys) As Boolean
        If InterfaceResourcesReleased OrElse IsDisposed OrElse Disposing OrElse Not IsHandleCreated OrElse IsAuthoringPreview Then Return False
        If Not IsNavigationKey(keys) OrElse Not CanNavigateStandalone(editor) Then Return False
        Dim popup = TryCast(editor, PopupBaseEdit)
        If popup IsNot Nothing AndAlso popup.IsPopupOpen Then Return False
        'Retain native caret/spin/Tab handling for standalone inputs; only Enter
        'uses the shared traversal. Arrows still select its remembered axis.
        If (keys And Keys.KeyCode) <> Keys.Enter Then
            RememberEnterDirection(keys)
            Return False
        End If
        If NavigationPending Then Return True
        If Not editor.DoValidate(PopupCloseMode.Normal) Then Return True
        Dim requested = editor.EditValue
        Dim tag = DirectCast(editor.Tag, SingleCellDataTag)
        Dim prior = EditorValueFromCell(tag.TargetWorksheet.Cells(tag.TargetCell), tag.DataType)
        If Not InplaceEditorFormatting.SameEditorValue(requested, prior) Then
            'Flush a pending buffered text/spin edit through its normal typed
            'ChangeManager route before moving. A rejected/reverted edit stays put.
            SingleCellDirtyMarker(editor, EventArgs.Empty)
            SingleCell_Value_Push(editor, EventArgs.Empty)
            If Not InplaceEditorFormatting.SameEditorValue(requested, editor.EditValue) Then Return True
        End If
        QueueEditorNavigation(New EditorPosition With {.Host = editor, .Standalone = editor}, keys)
        Return True
    End Function

    Private Function EnterNavigationPositions() As List(Of EditorPosition)
        Dim result = NavigationPositions()
        For Each editor In StandaloneNavigationEditors.ToArray()
            If CanNavigateStandalone(editor) Then result.Add(New EditorPosition With {.Host = editor, .Standalone = editor})
        Next
        Return result
    End Function

    Private Shared Function FindEnterTarget(origin As EditorPosition, candidates As List(Of EditorPosition),
                                           vertical As Boolean, previous As Boolean) As EditorPosition
        If candidates.Count = 0 Then Return Nothing
        Dim major As Func(Of EditorPosition, Integer) = Function(p) If(vertical, p.X, p.Y)
        Dim minor As Func(Of EditorPosition, Integer) = Function(p) If(vertical, p.Y, p.X)
        Dim ordered = candidates.Where(Function(p) p.Host Is origin.Host).OrderBy(major).ThenBy(minor)
        'Compare coordinates rather than list indices: calculation can lock or
        'remove the origin, which must not reset the next destination to the top.
        Dim later As Func(Of EditorPosition, Boolean) = Function(p) major(p) > major(origin) OrElse
            (major(p) = major(origin) AndAlso minor(p) > minor(origin))
        Dim earlier As Func(Of EditorPosition, Boolean) = Function(p) major(p) < major(origin) OrElse
            (major(p) = major(origin) AndAlso minor(p) < minor(origin))
        Dim target = If(previous, ordered.LastOrDefault(earlier), ordered.FirstOrDefault(later))
        If target IsNot Nothing Then Return target
        Dim hosts = candidates.Select(Function(p) p.Host).Append(origin.Host).Distinct().
            OrderBy(Function(c) c.PointToScreen(Point.Empty).Y).ThenBy(Function(c) c.PointToScreen(Point.Empty).X).ToList()
        Dim index = hosts.IndexOf(origin.Host)
        For offset = 1 To hosts.Count
            Dim nextIndex = (index + If(previous, -offset, offset) + hosts.Count) Mod hosts.Count
            Dim adjacent = candidates.Where(Function(p) p.Host Is hosts(nextIndex)).OrderBy(major).ThenBy(minor)
            target = If(previous, adjacent.LastOrDefault(), adjacent.FirstOrDefault())
            If target IsNot Nothing Then Return target
        Next
        Return Nothing
    End Function

    Private Function NavigationPositions() As List(Of EditorPosition)
        Dim result As New List(Of EditorPosition)
        If GridControls IsNot Nothing Then
            For Each grid In GridControls
                If grid Is Nothing OrElse grid.IsDisposed OrElse Not grid.Visible OrElse Not grid.Enabled Then Continue For
                Dim view = TryCast(grid.MainView, GridView)
                Dim source = TryCast(grid.DataSource, AbovoUnboundSource)
                If view Is Nothing OrElse source Is Nothing Then Continue For
                Dim index = source.UBSTag.DSIndex
                If index < 0 OrElse index >= DataPres.DataSets.Count Then Continue For
                Dim columns = view.VisibleColumns.Cast(Of GridColumn)().OrderBy(Function(c) c.VisibleIndex).ToList()
                For x = 0 To columns.Count - 1
                    Dim column = columns(x)
                    Dim tag = TryCast(column.Tag, Abovo.DataObject.DataColumnTag)
                    Dim helper As ColumnInplaceEditorHelper = Nothing
                    If tag IsNot Nothing AndAlso tag.HasIncolumnEditor Then helper = If(tag.InColumnEditorCombo?.InPlaceColumnHelper, tag.InColumnEditorDate?.InPlaceColumnHelper)
                    If helper IsNot Nothing Then result.Add(New EditorPosition With {.Host = grid, .View = view, .Header = helper, .Column = column, .X = x, .Y = -1})
                    For row = 0 To view.DataRowCount - 1
                        If view.GetVisibleIndex(row) < 0 Then Continue For
                        Dim dataRow = view.GetDataSourceRowIndex(row)
                        If Not CanPasteToDataPoint(DataPres.DataSets(index), dataRow, GetGridColumnIndex(column)) Then Continue For
                        result.Add(New EditorPosition With {.Host = grid, .View = view, .Column = column, .RowHandle = row, .X = x, .Y = view.GetVisibleIndex(row)})
                    Next
                Next
            Next
        End If
        If VertGridControls IsNot Nothing Then
            For Each grid In VertGridControls
                If grid Is Nothing OrElse grid.IsDisposed OrElse Not grid.Visible OrElse Not grid.Enabled Then Continue For
                Dim source = TryCast(grid.DataSource, AbovoUnboundSource)
                If source Is Nothing Then Continue For
                Dim index = source.UBSTag.DSIndex
                If index < 0 OrElse index >= DataPres.DataSets.Count Then Continue For
                Dim rows As New List(Of EditorRow)
                VisibleEditorRows(grid.Rows.Cast(Of BaseRow)(), rows)
                For y = 0 To rows.Count - 1
                    Dim row = rows(y), tag = GetVGridColumnTag(row)
                    Dim helper As VGridRowInplaceEditorHelper = Nothing
                    If tag IsNot Nothing AndAlso tag.HasIncolumnEditor Then helper = If(TryCast(tag.InColumnEditorCombo?.Tag, InColumnEditorTagCombo)?.InPlaceVGridRowHelper, TryCast(tag.InColumnEditorDate?.Tag, InColumnEditorTagDateEdit)?.InPlaceVGridRowHelper)
                    If helper IsNot Nothing Then result.Add(New EditorPosition With {.Host = grid, .VHeader = helper, .VRow = row, .Record = -1, .X = -1, .Y = y})
                    For record = 0 To grid.RecordCount - 1
                        If Not CanPasteToDataPoint(DataPres.DataSets(index), record, GetVGridColumnIndex(row)) Then Continue For
                        result.Add(New EditorPosition With {.Host = grid, .VRow = row, .Record = record, .X = record, .Y = y})
                    Next
                Next
            Next
        End If
        Return result
    End Function

    Private Shared Sub VisibleEditorRows(rows As IEnumerable(Of BaseRow), result As List(Of EditorRow))
        For Each row In rows.OrderBy(Function(r) r.VisibleIndex)
            If Not row.Visible Then Continue For
            If TypeOf row Is EditorRow Then result.Add(DirectCast(row, EditorRow))
            If row.Expanded Then VisibleEditorRows(row.ChildRows.Cast(Of BaseRow)(), result)
        Next
    End Sub

    Private Shared Function FindNavigationTarget(origin As EditorPosition, candidates As List(Of EditorPosition), keys As Keys) As EditorPosition
        Dim key = keys And Keys.KeyCode
        Dim previous = key = Keys.Left OrElse key = Keys.Up OrElse (key = Keys.Tab AndAlso (keys And Keys.Shift) <> Keys.None)
        Dim sameGrid = candidates.Where(Function(p) p.Host Is origin.Host).ToList()
        Dim horizontal = key = Keys.Tab OrElse key = Keys.Left OrElse key = Keys.Right
        If horizontal Then
            Dim sameRow = sameGrid.Where(Function(p) p.Y = origin.Y AndAlso If(previous, p.X < origin.X, p.X > origin.X))
            Dim target = If(previous, sameRow.OrderByDescending(Function(p) p.X).FirstOrDefault(), sameRow.OrderBy(Function(p) p.X).FirstOrDefault())
            If target IsNot Nothing OrElse key <> Keys.Tab Then Return target
        End If
        Dim nextRows = sameGrid.Where(Function(p) If(previous, p.Y < origin.Y, p.Y > origin.Y))
        If Not horizontal Then nextRows = nextRows.Where(Function(p) p.X = origin.X)
        Dim nextTarget = If(previous, nextRows.OrderByDescending(Function(p) p.Y).ThenByDescending(Function(p) p.X).FirstOrDefault(), nextRows.OrderBy(Function(p) p.Y).ThenBy(Function(p) p.X).FirstOrDefault())
        If nextTarget IsNot Nothing Then Return nextTarget
        Dim hosts = candidates.Select(Function(p) p.Host).Append(origin.Host).Distinct().OrderBy(Function(c) c.PointToScreen(Point.Empty).Y).ThenBy(Function(c) c.PointToScreen(Point.Empty).X).ToList()
        Dim position = hosts.IndexOf(origin.Host) + If(previous, -1, 1)
        If position < 0 OrElse position >= hosts.Count Then Return Nothing
        Dim adjacent = candidates.Where(Function(p) p.Host Is hosts(position))
        If horizontal Then Return If(previous, adjacent.OrderByDescending(Function(p) p.Y).ThenByDescending(Function(p) p.X).FirstOrDefault(), adjacent.OrderBy(Function(p) p.Y).ThenBy(Function(p) p.X).FirstOrDefault())
        Return If(previous, adjacent.OrderByDescending(Function(p) p.Y).ThenBy(Function(p) Math.Abs(p.X - origin.X)).FirstOrDefault(), adjacent.OrderBy(Function(p) p.Y).ThenBy(Function(p) Math.Abs(p.X - origin.X)).FirstOrDefault())
    End Function

    Private Sub ActivateEditorPosition(target As EditorPosition)
        If target.Standalone IsNot Nothing Then
            target.Standalone.Focus()
            ScrollEditorIntoPage(target.Standalone)
        ElseIf target.Header IsNot Nothing Then
            target.Header.ShowEditorFromKeyboard()
        ElseIf target.VHeader IsNot Nothing Then
            target.VHeader.ShowEditorFromKeyboard()
        ElseIf target.View IsNot Nothing Then
            If Not target.Host.ContainsFocus Then target.Host.Focus()
            target.View.FocusedRowHandle = target.RowHandle
            target.View.FocusedColumn = target.Column
            target.View.ShowEditor()
            ScrollEditorIntoPage(target.View.ActiveEditor)
        Else
            Dim grid = DirectCast(target.Host, VGridControl)
            If Not grid.ContainsFocus Then grid.Focus()
            grid.FocusedRow = target.VRow : grid.FocusedRecord = target.Record
            grid.MakeRowVisible(target.VRow)
            grid.ShowEditor()
            ScrollEditorIntoPage(grid.ActiveEditor)
        End If
    End Sub

    Private Shared Sub ScrollEditorIntoPage(editor As BaseEdit)
        If editor Is Nothing OrElse editor.IsDisposed Then Return
        Dim parent = editor.Parent
        While parent IsNot Nothing AndAlso Not TypeOf parent Is DevExpress.XtraTab.XtraTabPage
            parent = parent.Parent
        End While
        Dim page = TryCast(parent, DevExpress.XtraTab.XtraTabPage)
        If page Is Nothing OrElse Not page.AutoScroll Then Return
        Dim bounds = page.RectangleToClient(editor.RectangleToScreen(editor.ClientRectangle))
        Dim dx = If(bounds.Left < 0, bounds.Left, Math.Max(0, bounds.Right - page.ClientSize.Width))
        Dim dy = If(bounds.Top < 0, bounds.Top, Math.Max(0, bounds.Bottom - page.ClientSize.Height))
        If dx <> 0 OrElse dy <> 0 Then page.AutoScrollPosition = New Point(-page.AutoScrollPosition.X + dx, -page.AutoScrollPosition.Y + dy)
    End Sub

    Private Sub ConfigureHeaderNavigation(helper As ColumnInplaceEditorHelper)
        helper.Navigate = Function(keys) NavigateGrid(helper.GridControl, keys, header:=helper)
    End Sub
    Private Sub ConfigureHeaderNavigation(helper As VGridRowInplaceEditorHelper)
        helper.Navigate = Function(keys) NavigateGrid(helper.GridControl, keys, vHeader:=helper)
    End Sub
End Class

'A Funding VGrid may be taller than the page. WinForms' queued focus scrolling
'tries to show the WHOLE grid, taking the page back to its top after an edit.
'The grid handles its own cells; explicit keyboard navigation reveals only
'the destination editor. Other controls retain standard auto-scroll behaviour.
Public Class DITTabPage
    Inherits DevExpress.XtraTab.XtraTabPage
    Protected Overrides Function CanScrollControlIntoView(activeControl As Control) As Boolean
        Dim control = activeControl
        While control IsNot Nothing AndAlso control IsNot Me
            If TypeOf control Is GridControl OrElse TypeOf control Is VGridControl Then Return False
            control = control.Parent
        End While
        Return MyBase.CanScrollControlIntoView(activeControl)
    End Function
End Class
