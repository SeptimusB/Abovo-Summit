Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.CompilerServices
Imports DevExpress.Utils
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.BandedGrid
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.XtraVerticalGrid.Rows
Imports DevExpress.XtraEditors
Imports DevExpress.XtraTreeList

Namespace Abovo
    'Local presentation preferences only: never changes a workbook or its dirty state.
    Friend NotInheritable Class GridPresentation
        Private Class State
            Public Percent As Integer = 100
            Public Key As String
            Public Ready As Boolean
            Public Applying As Boolean
            Public PendingZoom As Integer?
            Public PendingPoint As Point?
            Public ZoomQueued As Boolean
            Public ResetLayout As Action
            Public VerticalWheelRemainder As Integer
            Public Fonts As New List(Of Font)
            Public EditorFonts As New Dictionary(Of String, Font)
            Public Metrics As New ConditionalWeakTable(Of Object, Dictionary(Of String, ZoomMetric))
        End Class
        Private Class ZoomMetric
            Public Baseline As Double
            Public LastValue As Double
        End Class
        Private Shared ReadOnly States As New ConditionalWeakTable(Of Control, State)
        Private Shared ReadOnly MenuTargets As New ConditionalWeakTable(Of ContextMenuStrip, MenuTarget)
        Private Class MenuTarget
            Public Control As Control
            Public OpeningAttached As Boolean
            Public Point As Point?
        End Class
        Private Const MinimumZoom As Integer = 60
        Private Const MinimumFontPoints As Single = 8.0F
        Private Class ZoomAnchor
            Public Point As Point
            Public FractionX As Double
            Public FractionY As Double
            Public RowHandle As Integer = -1
            '0=data cell, 1=group heading, 2=group footer (native virtual row).
            Public GridRowKind As Integer
            Public FooterOwnerRowHandle As Integer
            Public Column As DevExpress.XtraGrid.Columns.GridColumn
            Public VRow As BaseRow
            Public Record As Integer = -1
            Public CellIndex As Integer
            Public Node As DevExpress.XtraTreeList.Nodes.TreeListNode
            Public TreeColumn As DevExpress.XtraTreeList.Columns.TreeListColumn
        End Class
        Private Shared ReadOnly Wheel As New GridWheelFilter
        Private Shared Installed As Boolean
        Friend Shared Event ZoomChanged As EventHandler(Of PresentationScaleChangedEventArgs)

        Friend Shared Function ReadPreference(key As String, name As String, fallback As String) As String
            Try
                Using entry = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(PreferencePath(key))
                    Return CStr(If(entry?.GetValue(name, fallback), fallback))
                End Using
            Catch ex As Exception When TypeOf ex Is Security.SecurityException OrElse TypeOf ex Is UnauthorizedAccessException OrElse TypeOf ex Is IO.IOException
                Return fallback
            End Try
        End Function

        Friend Shared Sub WritePreference(key As String, name As String, value As String)
            Try
                Using entry = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(PreferencePath(key))
                    entry.SetValue(name, value)
                End Using
            Catch ex As Exception When TypeOf ex Is Security.SecurityException OrElse TypeOf ex Is UnauthorizedAccessException OrElse TypeOf ex Is IO.IOException
                SummitDiagnostics.WriteLine("Grid preference could not be saved: " & ex.Message)
            End Try
        End Sub

        Private Shared Function PreferencePath(key As String) As String
            Using hash = Security.Cryptography.SHA256.Create()
                Return "Software\Abovo\Summit\GridPresentation\" & BitConverter.ToString(hash.ComputeHash(System.Text.Encoding.UTF8.GetBytes(key))).Replace("-", "")
            End Using
        End Function

        Friend Shared Sub Configure(control As Control, Optional key As String = Nothing)
            If Not IsGrid(control) OrElse control.IsDisposed OrElse control.Disposing Then Return
            Dim existing As State = Nothing
            If States.TryGetValue(control, existing) Then
                If key IsNot Nothing Then existing.Key = key
                Return
            End If
            Dim path As New List(Of String)
            Dim parent = control
            While parent IsNot Nothing
                path.Insert(0, If(String.IsNullOrEmpty(parent.Name), parent.GetType().Name & "/" & If(parent.Parent Is Nothing, 0, parent.Parent.Controls.GetChildIndex(parent)).ToString(), parent.Name))
                parent = parent.Parent
            End While
            Dim state As New State With {.Key = If(key, String.Join("/", path))}
            States.Add(control, state)
            AddHandler control.Disposed, Sub()
                For Each font In state.Fonts
                    font.Dispose()
                Next
            End Sub
            If Not Installed Then
                Application.AddMessageFilter(Wheel)
                Installed = True
            End If
            Dim vertical = TryCast(control, VGridControl)
            If vertical IsNot Nothing Then
                vertical.OptionsView.FixedLineWidth = 3
                vertical.Appearance.FixedLine.BackColor = Color.SteelBlue
                vertical.Appearance.FixedLine.Options.UseBackColor = True
                vertical.OptionsBehavior.RecordsMouseWheel = False
                AddHandler vertical.ShownEditor, Sub() ScaleActiveEditor(control, vertical.ActiveEditor)
            ElseIf TypeOf control Is TreeList Then
                Dim tree = DirectCast(control, TreeList)
                tree.FixedLineWidth = 3
                tree.Appearance.FixedLine.BackColor = Color.SteelBlue
                tree.Appearance.FixedLine.Options.UseBackColor = True
                AddHandler tree.ShownEditor, Sub() ScaleActiveEditor(control, tree.ActiveEditor)
            Else
                Dim grid = DirectCast(control, GridControl)
                AddHandler grid.ViewRegistered, Sub(sender, args) AttachView(grid, TryCast(args.View, GridView))
                For Each view As GridView In grid.ViewCollection.OfType(Of GridView)()
                    AttachView(grid, view)
                Next
            End If
            AddHandler control.VisibleChanged, Sub() QueueInitialZoom(control, state)
            AddHandler control.HandleCreated, Sub() QueueInitialZoom(control, state)
            AddHandler control.ContextMenuStripChanged, Sub() AddMenu(control)
            QueueInitialZoom(control, state)
        End Sub

        Private Shared Function IsGrid(control As Control) As Boolean
            Return TypeOf control Is GridControl OrElse TypeOf control Is VGridControl OrElse TypeOf control Is TreeList
        End Function

        Friend Shared Sub SetMenuTarget(menu As ContextMenuStrip, control As Control)
            If menu IsNot Nothing AndAlso control IsNot Nothing Then MenuTargets.GetOrCreateValue(menu).Control = control
        End Sub

        Private Shared Function MenuGrid(menu As ContextMenuStrip, fallback As Control) As Control
            Dim target = menu.SourceControl
            While target IsNot Nothing AndAlso Not IsGrid(target)
                target = target.Parent
            End While
            Dim saved As MenuTarget = Nothing
            If target Is Nothing AndAlso MenuTargets.TryGetValue(menu, saved) Then target = saved.Control
            Return If(target, fallback)
        End Function

        Private Shared Sub AttachView(grid As GridControl, view As GridView)
            If view Is Nothing Then Return
            view.FixedLineWidth = 3
            view.Appearance.FixedLine.BackColor = Color.SteelBlue
            view.Appearance.FixedLine.Options.UseBackColor = True
            AddHandler view.ShownEditor, Sub() ScaleActiveEditor(grid, view.ActiveEditor)
        End Sub

        Private Shared Sub QueueInitialZoom(control As Control, state As State)
            If state.Ready OrElse control.IsDisposed OrElse Not control.IsHandleCreated OrElse Not control.Visible Then Return
            state.Ready = True
            control.BeginInvoke(New MethodInvoker(Sub()
                If control.IsDisposed Then Return
                AddMenu(control)
                Dim percent As Integer = 100
                If Not Integer.TryParse(ReadPreference(state.Key, "Zoom", "100"), percent) Then percent = 100
                SetZoom(control, percent, False)
            End Sub))
        End Sub

        Friend Shared Function ZoomPercent(control As Control) As Integer
            Dim state As State = Nothing
            Return If(control IsNot Nothing AndAlso States.TryGetValue(control, state), state.Percent, 100)
        End Function

        Friend Shared Function IsApplyingZoom(control As Control) As Boolean
            Dim state As State = Nothing
            Return control IsNot Nothing AndAlso States.TryGetValue(control, state) AndAlso state.Applying
        End Function

        Friend Shared Sub SetResetLayout(control As Control, reset As Action)
            Configure(control)
            States.GetOrCreateValue(control).ResetLayout = reset
        End Sub

        Friend Shared Sub ResetZoom(control As Control)
            SetZoom(control, 100)
            Dim state As State = Nothing
            If States.TryGetValue(control, state) Then state.ResetLayout?.Invoke()
        End Sub

        Private Shared Function ZoomValue(state As State, owner As Object, key As String, current As Double,
                                          oldPercent As Integer, newPercent As Integer) As Double
            Dim metrics = state.Metrics.GetOrCreateValue(owner)
            Dim metric As ZoomMetric = Nothing
            If Not metrics.TryGetValue(key, metric) Then
                metric = New ZoomMetric With {.Baseline = current * 100.0 / oldPercent, .LastValue = current}
                metrics.Add(key, metric)
            ElseIf Math.Abs(metric.LastValue - current) > 0.01 Then
                'Honour an explicit resize/global scale without accumulating rounded zoom steps.
                metric.Baseline = current * 100.0 / oldPercent
            End If
            metric.LastValue = metric.Baseline * newPercent / 100.0
            Return metric.LastValue
        End Function

        Friend Shared Sub AcceptRenderedMetric(control As Control, owner As Object, key As String, current As Double)
            Dim state As State = Nothing
            Dim metrics As Dictionary(Of String, ZoomMetric) = Nothing
            Dim metric As ZoomMetric = Nothing
            If control Is Nothing OrElse owner Is Nothing OrElse Not States.TryGetValue(control, state) Then Return
            If state.Metrics.TryGetValue(owner, metrics) AndAlso metrics.TryGetValue(key, metric) Then
                'Native editor chrome has minimum dimensions. Keep the requested zoom baseline,
                'but do not mistake this rendering clamp for a user resize on the next zoom.
                metric.LastValue = current
            End If
        End Sub

        Private Shared Function ZoomSize(state As State, owner As Object, key As String, current As Integer,
                                         oldPercent As Integer, newPercent As Integer, minimum As Integer,
                                         Optional maximum As Integer = Integer.MaxValue) As Integer
            Dim value = Math.Max(minimum, Math.Min(maximum, CInt(Math.Round(ZoomValue(state, owner, key, current, oldPercent, newPercent)))))
            state.Metrics.GetOrCreateValue(owner)(key).LastValue = value
            Return value
        End Function

        Private Shared Function BaselineFontSize(state As State, control As Control,
                                                 appearance As AppearanceObject, oldPercent As Integer) As Double
            Dim owner As Object = If(appearance.Options.UseFont OrElse appearance.FontSizeDelta <> 0, DirectCast(appearance, Object), control)
            Dim current = If(owner Is control, CDbl(control.Font.SizeInPoints), CDbl(appearance.Font.SizeInPoints + appearance.FontSizeDelta))
            Dim metrics As Dictionary(Of String, ZoomMetric) = Nothing
            Dim metric As ZoomMetric = Nothing
            If state.Metrics.TryGetValue(owner, metrics) AndAlso metrics.TryGetValue("Font", metric) AndAlso Math.Abs(metric.LastValue - current) <= 0.01 Then
                Return metric.Baseline
            End If
            Return current * 100.0 / oldPercent
        End Function

        Private Shared Function ZoomWidth(state As State, owner As Object, key As String, current As Integer,
                                          oldPercent As Integer, newPercent As Integer, baselineFont As Double,
                                          minimum As Integer, Optional maximum As Integer = Integer.MaxValue) As Integer
            Dim requested = ZoomValue(state, owner, key, current, oldPercent, newPercent)
            Dim metric = state.Metrics.GetOrCreateValue(owner)(key)
            'Do not shrink columns faster than their readable text. When the font
            'hits its eight-point floor, keep the same baseline text-to-width ratio.
            'Retain the unrounded requested baseline so reset never grows columns.
            Dim readable = metric.Baseline * MinimumFontPoints / Math.Max(0.1, baselineFont)
            Dim value = Math.Max(minimum, Math.Min(maximum, CInt(Math.Max(Math.Round(requested), Math.Ceiling(readable)))))
            metric.LastValue = value
            Return value
        End Function

        Private Shared Function SizedFont(state As State, original As Font, size As Single) As Font
            Dim key = original.FontFamily.Name & "/" & CInt(original.Style).ToString() & "/" & size.ToString(Globalization.CultureInfo.InvariantCulture)
            Dim font As Font = Nothing
            If Not state.EditorFonts.TryGetValue(key, font) Then
                font = New Font(original.FontFamily, size, original.Style, GraphicsUnit.Point)
                state.EditorFonts.Add(key, font) : state.Fonts.Add(font)
            End If
            Return font
        End Function

        Private Shared Sub QueueWheelZoom(control As Control, delta As Integer)
            Configure(control)
            Dim state As State = Nothing
            If Not States.TryGetValue(control, state) OrElse control.IsDisposed Then Return
            state.PendingZoom = Math.Max(MinimumZoom, Math.Min(200, If(state.PendingZoom, state.Percent) + Math.Sign(delta) * 10))
            state.PendingPoint = control.PointToClient(Control.MousePosition)
            If state.ZoomQueued Then Return
            If Not control.IsHandleCreated Then
                SetZoomAt(control, state.PendingZoom.Value, state.PendingPoint)
                Return
            End If
            state.ZoomQueued = True
            control.BeginInvoke(New MethodInvoker(Sub()
                state.ZoomQueued = False
                If control.IsDisposed OrElse Not state.PendingZoom.HasValue Then Return
                SetZoomAt(control, state.PendingZoom.Value, state.PendingPoint)
            End Sub))
        End Sub

        Friend Shared Sub SetZoom(control As Control, percent As Integer, Optional persist As Boolean = True)
            ApplyZoom(control, percent, persist, True, True)
        End Sub

        'Keep the data under the pointer in the same part of the viewport. Native
        'record/row scrollbars are discrete, so use the closest available position.
        Friend Shared Sub SetZoomAt(control As Control, percent As Integer, point As Point?, Optional persist As Boolean = True)
            If control Is Nothing OrElse control.IsDisposed OrElse control.Disposing Then Return
            Dim anchor = If(point.HasValue AndAlso control.ClientRectangle.Contains(point.Value), CaptureAnchor(control, point.Value), Nothing)
            Dim previous = ZoomPercent(control)
            ApplyZoom(control, percent, persist, True, True)
            If anchor IsNot Nothing AndAlso ZoomPercent(control) <> previous Then
                Dim state = States.GetOrCreateValue(control)
                Dim wasApplying = state.Applying
                state.Applying = True
                Try
                    'The viewport adjustment is still part of the zoom transaction:
                    'header helpers must not interpret its LayoutChanged as a commit.
                    RestoreAnchor(control, anchor)
                Finally
                    state.Applying = wasApplying
                End Try
            End If
        End Sub

        Private Shared Function CaptureAnchor(control As Control, point As Point) As ZoomAnchor
            Dim anchor As New ZoomAnchor With {.Point = point}
            Dim vertical = TryCast(control, VGridControl)
            Dim view = TryCast(TryCast(control, GridControl)?.MainView, GridView)
            Dim tree = TryCast(control, TreeList)
            If vertical IsNot Nothing Then
                Dim hit = vertical.CalcHitInfo(point)
                anchor.VRow = hit.Row : anchor.Record = hit.RecordIndex : anchor.CellIndex = Math.Max(0, hit.CellIndex)
                If hit.HitInfoType <> HitInfoTypeEnum.ValueCell OrElse anchor.VRow Is Nothing OrElse anchor.Record < 0 Then Return Nothing
            ElseIf view IsNot Nothing Then
                Dim hit = view.CalcHitInfo(point)
                anchor.RowHandle = hit.RowHandle : anchor.Column = hit.Column
                If view.IsGroupRow(hit.RowHandle) OrElse hit.HitTest = DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitTest.RowFooter OrElse
                   Not hit.InRowCell OrElse anchor.Column Is Nothing Then
                    If Not CaptureVirtualGridAnchor(view, point, anchor) Then Return Nothing
                End If
            ElseIf tree IsNot Nothing Then
                Dim hit = tree.CalcHitInfo(point)
                anchor.Node = hit.Node : anchor.TreeColumn = hit.Column
                If anchor.Node Is Nothing OrElse anchor.TreeColumn Is Nothing Then Return Nothing
            Else
                Return Nothing
            End If
            Dim bounds = AnchorBounds(control, anchor)
            If bounds.Width <= 0 OrElse bounds.Height <= 0 Then Return Nothing
            anchor.FractionX = Math.Max(0.0, Math.Min(1.0, CDbl(point.X - bounds.Left) / bounds.Width))
            anchor.FractionY = Math.Max(0.0, Math.Min(1.0, CDbl(point.Y - bounds.Top) / bounds.Height))
            Return anchor
        End Function

        Private Shared Function CaptureVirtualGridAnchor(view As GridView, point As Point, anchor As ZoomAnchor) As Boolean
            Dim info = TryCast(view.GetViewInfo(), DevExpress.XtraGrid.Views.Grid.ViewInfo.GridViewInfo)
            If info Is Nothing Then Return False
            'Only inspect materialised visible rows, never walk the workbook.
            For Each row As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridRowInfo In info.RowsInfo
                For Each footer As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridRowFooterInfo In row.RowFooters
                    If Not footer.Bounds.Contains(point) Then Continue For
                    anchor.GridRowKind = 2
                    anchor.RowHandle = footer.RowHandle
                    anchor.FooterOwnerRowHandle = row.RowHandle
                    Exit For
                Next
                If anchor.GridRowKind = 2 Then Exit For
                If row.IsGroupRow AndAlso row.Bounds.Contains(point) Then
                    anchor.GridRowKind = 1
                    anchor.RowHandle = row.RowHandle
                    Exit For
                End If
            Next
            If anchor.GridRowKind = 0 Then Return False
            'Fixed headers take priority over scrolled headers lying beneath them.
            For Each column In view.VisibleColumns.Cast(Of DevExpress.XtraGrid.Columns.GridColumn)().
                OrderBy(Function(c) If(c.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.None, 1, 0))
                Dim header = info.ColumnsInfo(column)
                If header Is Nothing OrElse point.X < header.Bounds.Left OrElse point.X >= header.Bounds.Right Then Continue For
                anchor.Column = column
                Return True
            Next
            Return False
        End Function

        Private Shared Function AnchorBounds(control As Control, anchor As ZoomAnchor) As Rectangle
            Dim vertical = TryCast(control, VGridControl)
            If vertical IsNot Nothing Then
                Dim info = vertical.ViewInfo.GetRowValueInfo(anchor.VRow, anchor.Record, anchor.CellIndex)
                Return If(info Is Nothing, Rectangle.Empty, info.Bounds)
            End If
            Dim view = TryCast(TryCast(control, GridControl)?.MainView, GridView)
            If view IsNot Nothing Then
                Dim info = TryCast(view.GetViewInfo(), DevExpress.XtraGrid.Views.Grid.ViewInfo.GridViewInfo)
                If info Is Nothing Then Return Rectangle.Empty
                If anchor.GridRowKind = 0 Then
                    Dim cell = info.GetGridCellInfo(anchor.RowHandle, anchor.Column)
                    Return If(cell Is Nothing, Rectangle.Empty, cell.Bounds)
                End If
                Dim bounds = Rectangle.Empty
                If anchor.GridRowKind = 1 Then
                    Dim row = info.GetGridRowInfo(anchor.RowHandle)
                    If row IsNot Nothing Then bounds = row.Bounds
                Else
                    For Each row As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridRowInfo In info.RowsInfo
                        For Each footer As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridRowFooterInfo In row.RowFooters
                            If footer.RowHandle <> anchor.RowHandle Then Continue For
                            bounds = footer.Bounds
                            Exit For
                        Next
                        If Not bounds.IsEmpty Then Exit For
                    Next
                End If
                Dim header = info.ColumnsInfo(anchor.Column)
                If bounds.IsEmpty OrElse header Is Nothing Then Return Rectangle.Empty
                Return New Rectangle(header.Bounds.Left, bounds.Top, header.Bounds.Width, bounds.Height)
            End If
            Dim tree = TryCast(control, TreeList)
            If tree IsNot Nothing Then
                Dim row = tree.ViewInfo.RowsInfo(anchor.Node)
                Dim cell = row?.Cells.FirstOrDefault(Function(c) c.Column Is anchor.TreeColumn)
                Return If(cell Is Nothing, Rectangle.Empty, cell.Bounds)
            End If
            Return Rectangle.Empty
        End Function

        Private Shared Sub RestoreAnchor(control As Control, anchor As ZoomAnchor)
            If control.IsDisposed OrElse control.Disposing Then Return
            Dim vertical = TryCast(control, VGridControl)
            Dim view = TryCast(TryCast(control, GridControl)?.MainView, GridView)
            Dim tree = TryCast(control, TreeList)
            If vertical IsNot Nothing Then
                vertical.MakeRowVisible(anchor.VRow)
                vertical.MakeRecordVisible(anchor.Record)
            ElseIf view IsNot Nothing Then
                view.MakeRowVisible(If(anchor.GridRowKind = 2, anchor.FooterOwnerRowHandle, anchor.RowHandle))
                view.MakeColumnVisible(anchor.Column)
            ElseIf tree IsNot Nothing Then
                tree.MakeNodeVisible(anchor.Node)
                tree.MakeColumnVisible(anchor.TreeColumn)
            End If
            For pass = 0 To 2
                control.Update()
                Dim bounds = AnchorBounds(control, anchor)
                If bounds.Width <= 0 OrElse bounds.Height <= 0 Then Exit For
                Dim deltaX = CInt(Math.Round(bounds.Left + bounds.Width * anchor.FractionX - anchor.Point.X))
                Dim deltaY = CInt(Math.Round(bounds.Top + bounds.Height * anchor.FractionY - anchor.Point.Y))
                If vertical IsNot Nothing Then
                    Dim beforeX = vertical.LeftVisibleRecord, beforePixelX = vertical.LeftVisibleRecordPixel
                    Dim beforePixelY = vertical.TopVisibleRowIndexPixel
                    If vertical.LayoutStyle = LayoutViewStyle.MultiRecordView Then
                        vertical.LeftVisibleRecordPixel = Math.Max(0, beforePixelX + deltaX)
                    Else
                        vertical.LeftVisibleRecord = Math.Max(0, beforeX + CInt(Math.Round(CDbl(deltaX) / Math.Max(1, vertical.RecordWidth))))
                    End If
                    Dim fixed = vertical.FixedTopRows.Cast(Of BaseRow)().Any(Function(r) r Is anchor.VRow) OrElse vertical.FixedBottomRows.Cast(Of BaseRow)().Any(Function(r) r Is anchor.VRow)
                    'Funding mixes category and date rows of different heights.
                    'The native cumulative pixel offset handles those boundaries;
                    'dividing by the hovered row height can scroll past that row.
                    If Not fixed Then vertical.TopVisibleRowIndexPixel = Math.Max(0, beforePixelY + deltaY)
                    vertical.Update()
                    If beforeX = vertical.LeftVisibleRecord AndAlso beforePixelX = vertical.LeftVisibleRecordPixel AndAlso
                       beforePixelY = vertical.TopVisibleRowIndexPixel Then Exit For
                ElseIf view IsNot Nothing Then
                    Dim beforeX = view.LeftCoord, beforeY = view.TopRowIndex
                    If anchor.Column.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.None Then view.LeftCoord = Math.Max(0, beforeX + deltaX)
                    view.TopRowIndex = Math.Max(0, beforeY + CInt(Math.Round(CDbl(deltaY) / bounds.Height)))
                    view.LayoutChanged()
                    If beforeX = view.LeftCoord AndAlso beforeY = view.TopRowIndex Then Exit For
                ElseIf tree IsNot Nothing Then
                    Dim beforeX = tree.LeftCoord, beforeY = tree.TopVisibleNodeIndex, beforePixel = tree.TopVisibleNodePixel
                    If anchor.TreeColumn.Fixed = DevExpress.XtraTreeList.Columns.FixedStyle.None Then tree.LeftCoord = Math.Max(0, beforeX + deltaX)
                    If tree.OptionsBehavior.AllowPixelScrolling = DefaultBoolean.True Then
                        tree.TopVisibleNodePixel = Math.Max(0, beforePixel + deltaY)
                    Else
                        tree.TopVisibleNodeIndex = Math.Max(0, beforeY + CInt(Math.Round(CDbl(deltaY) / bounds.Height)))
                    End If
                    tree.LayoutChanged()
                    If beforeX = tree.LeftCoord AndAlso beforeY = tree.TopVisibleNodeIndex AndAlso beforePixel = tree.TopVisibleNodePixel Then Exit For
                End If
            Next
            If vertical IsNot Nothing AndAlso AnchorBounds(control, anchor).IsEmpty Then
                'At a native scroll limit keep the source cell visible, even when
                'its original pointer position cannot be represented exactly.
                vertical.MakeRowVisible(anchor.VRow)
                vertical.MakeRecordVisible(anchor.Record)
                vertical.Update()
            End If
        End Sub

        'Absolute layout (window density / monitor change) must start from an
        'unzoomed grid, then restore the user's local zoom as one native update.
        'This is presentation only: do not post or close a pending editor.
        Friend Shared Function BeginBaseLayout(control As Control) As IDisposable
            Configure(control)
            Dim state As State = Nothing
            If control Is Nothing OrElse control.IsDisposed OrElse
               Not States.TryGetValue(control, state) OrElse state.Applying Then Return Nothing
            Return New BaseLayoutScope(control, state)
        End Function

        Private NotInheritable Class BaseLayoutScope
            Implements IDisposable

            Private ReadOnly Target As Control
            Private ReadOnly TargetState As State
            Private ReadOnly SavedPercent As Integer
            Private ReadOnly SavedPending As Integer?
            Private ReadOnly Vertical As VGridControl
            Private ReadOnly Tree As TreeList
            Private ReadOnly View As GridView
            Private ReadOnly VerticalAnchor As BaseRow
            Private Finished As Boolean

            Friend Sub New(control As Control, state As State)
                Target = control : TargetState = state
                SavedPercent = state.Percent : SavedPending = state.PendingZoom
                Vertical = TryCast(control, VGridControl)
                VerticalAnchor = CaptureVerticalViewportRow(Vertical)
                Tree = TryCast(control, TreeList)
                View = TryCast(TryCast(control, GridControl)?.MainView, GridView)
                control.SuspendLayout()
                Vertical?.BeginUpdate() : Tree?.BeginUpdate() : View?.BeginUpdate()
                Try
                    ClearVerticalRecordHeaderCache(Vertical)
                    ApplyZoom(control, 100, False, False, False)
                    state.Applying = True
                Catch
                    Dispose()
                    Throw
                End Try
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                If Finished Then Return
                Finished = True
                Try
                    TargetState.Applying = False
                    If Not Target.IsDisposed AndAlso Not Target.Disposing Then
                        TargetState.Metrics = New ConditionalWeakTable(Of Object, Dictionary(Of String, ZoomMetric))()
                        ApplyZoom(Target, SavedPercent, False, False, True, True)
                        TargetState.Applying = True
                    End If
                Finally
                    TargetState.PendingZoom = SavedPending
                    Try
                        If Not Target.IsDisposed Then
                            Try
                                Vertical?.EndUpdate() : Tree?.EndUpdate() : View?.EndUpdate()
                            Finally
                                Target.ResumeLayout(True)
                                RestoreEmptyVerticalViewport(Vertical, VerticalAnchor)
                                RefreshVerticalRecordHeaders(Vertical)
                            End Try
                        End If
                    Finally
                        TargetState.Applying = False
                    End Try
                    If Not Target.IsDisposed Then Target.Invalidate(True)
                End Try
            End Sub
        End Class

        Private Shared Function CaptureVerticalViewportRow(vertical As VGridControl) As BaseRow
            If vertical Is Nothing OrElse vertical.IsDisposed OrElse vertical.VisibleRows.Count = 0 Then Return Nothing
            Return vertical.VisibleRows(Math.Max(0, Math.Min(vertical.TopVisibleRowIndex, vertical.VisibleRows.Count - 1)))
        End Function

        Private Shared Sub RestoreEmptyVerticalViewport(vertical As VGridControl, anchor As BaseRow)
            If vertical Is Nothing OrElse vertical.IsDisposed OrElse anchor Is Nothing OrElse
               vertical.RecordCount = 0 OrElse vertical.ViewInfo.RowsViewInfo.Count > 0 Then Return
            'A large density change can leave the native scroller pointing beyond
            'its newly calculated bands. Restore the already-visible row through
            'the native scroller, without changing focus or rebuilding data.
            vertical.MakeRowVisible(anchor)
        End Sub

        Private Shared Sub ClearVerticalRecordHeaderCache(vertical As VGridControl)
            If vertical Is Nothing OrElse vertical.IsDisposed OrElse Not vertical.IsHandleCreated OrElse
               Not vertical.OptionsView.ShowRecordHeaders Then Return
            'DevExpress 25.2 retains these lines when a resize temporarily leaves
            'no visible rows. Font/cache updates dispose their borrowed brushes,
            'so discard the old drawing geometry before that update, not during paint.
            Dim headers = vertical.ViewInfo.RecordHeaderPanelViewInfo
            headers.HeadersInfo.Clear()
            headers.LinesInfo.Clear()
        End Sub

        Private Shared Sub RefreshVerticalRecordHeaders(vertical As VGridControl)
            If vertical Is Nothing OrElse vertical.IsDisposed OrElse Not vertical.IsHandleCreated OrElse
               Not vertical.OptionsView.ShowRecordHeaders Then Return
            'Recreate only header drawing information from the final native layout.
            'Do not refresh the datasource or hide/commit the active input editor.
            'An outer EndUpdate may have replaced brushes after an inner layout;
            'native Calc leaves the old lines intact when no rows are visible.
            ClearVerticalRecordHeaderCache(vertical)
            vertical.InvalidateRecordHeaders()
        End Sub

        Private Shared Sub NotifyZoomEditors(control As Control, oldPercent As Integer, newPercent As Integer)
            RaiseEvent ZoomChanged(control, New PresentationScaleChangedEventArgs(oldPercent / 100.0F, newPercent / 100.0F))
        End Sub

        Private Shared Function MinimumTextHeight(control As Control, font As Font) As Integer
            Using graphics = control.CreateGraphics()
                Dim line = Math.Max(font.GetHeight(graphics), TextRenderer.MeasureText(graphics, "Ag", font,
                    Size.Empty, TextFormatFlags.NoPadding Or TextFormatFlags.SingleLine).Height)
                Return CInt(Math.Ceiling(line)) + Math.Max(4, CInt(Math.Ceiling(graphics.DpiY / 24.0F)))
            End Using
        End Function

        Private Shared Sub ApplyZoom(control As Control, percent As Integer, persist As Boolean,
                                     commitEditor As Boolean, notifyEditors As Boolean, Optional forceReapply As Boolean = False)
            Configure(control)
            Dim state As State = Nothing
            If Not States.TryGetValue(control, state) OrElse state.Applying OrElse control.IsDisposed Then Return
            state.PendingZoom = Nothing
            percent = Math.Max(MinimumZoom, Math.Min(200, percent))
            If percent = state.Percent AndAlso Not forceReapply Then Return
            Dim vertical = TryCast(control, VGridControl)
            Dim grid = TryCast(control, GridControl)
            Dim tree = TryCast(control, TreeList)
            Dim view = If(grid Is Nothing, Nothing, TryCast(grid.MainView, GridView))
            If commitEditor AndAlso vertical IsNot Nothing AndAlso vertical.ActiveEditor IsNot Nothing Then
                If Not vertical.PostEditor() Then Return
                vertical.HideEditor()
            ElseIf commitEditor AndAlso view IsNot Nothing AndAlso view.ActiveEditor IsNot Nothing Then
                If Not view.PostEditor() OrElse Not view.UpdateCurrentRow() Then Return
                view.HideEditor()
            ElseIf commitEditor AndAlso tree IsNot Nothing Then
                If tree.ActiveEditor IsNot Nothing AndAlso Not tree.ActiveEditor.DoValidate() Then Return
                tree.CloseEditor()
            End If
            Dim old = state.Percent, ratio = CSng(percent) / old
            Dim bodyAppearance = If(vertical IsNot Nothing, vertical.Appearance.RecordValue,
                If(tree IsNot Nothing, tree.Appearance.Row, If(view IsNot Nothing, view.Appearance.Row, Nothing)))
            Dim baselineBodyFont = If(bodyAppearance Is Nothing, CDbl(control.Font.SizeInPoints) * 100.0 / old,
                BaselineFontSize(state, control, bodyAppearance, old))
            'Capture dimensions before Font changes trigger native autosizing.
            'Otherwise inherited-font layout can scale a width before we scale it again.
            Dim gridWidths = If(view Is Nothing, Nothing, view.Columns.Cast(Of DevExpress.XtraGrid.Columns.GridColumn)().ToDictionary(Function(c) c, Function(c) c.Width))
            Dim treeWidths = If(tree Is Nothing, Nothing, tree.Columns.Cast(Of DevExpress.XtraTreeList.Columns.TreeListColumn)().ToDictionary(Function(c) c, Function(c) c.Width))
            Dim verticalRecordWidth = If(vertical Is Nothing, 0, vertical.RecordWidth)
            Dim verticalHeaderWidth = If(vertical Is Nothing, 0, vertical.RowHeaderWidth)
            Dim verticalAnchor = CaptureVerticalViewportRow(vertical)
            Dim leftOffset = If(view IsNot Nothing, view.LeftCoord, If(tree IsNot Nothing, tree.LeftCoord, 0))
            Dim appearances As New HashSet(Of AppearanceObject)
            Dim scaleAppearance As Action(Of AppearanceObject) = Sub(a)
                If a Is Nothing OrElse Not appearances.Add(a) Then Return
                'Unset row/column appearances inherit from the view. Giving
                'them an explicit font would replace that with the skin's tiny
                'default font, producing different sizes in focused records.
                If Not a.Options.UseFont AndAlso a.FontSizeDelta = 0 Then Return
                'Use a stable point-size baseline, not rounded incremental deltas.
                'The analyser's custom painter reads Font directly (not FontSizeDelta).
                Dim size = CSng(Math.Max(MinimumFontPoints, ZoomValue(state, a, "Font", a.Font.SizeInPoints + a.FontSizeDelta, old, percent)))
                a.Font = SizedFont(state, a.Font, size)
                a.FontSizeDelta = 0
                state.Metrics.GetOrCreateValue(a)("Font").LastValue = a.Font.SizeInPoints
                a.Options.UseFont = True
            End Sub
            Dim layoutSuspended As Boolean, verticalUpdating As Boolean, treeUpdating As Boolean, viewUpdating As Boolean
            state.Applying = True
            Try
                control.SuspendLayout() : layoutSuspended = True
                If vertical IsNot Nothing Then
                    vertical.BeginUpdate() : verticalUpdating = True
                    ClearVerticalRecordHeaderCache(vertical)
                End If
                If tree IsNot Nothing Then
                    tree.BeginUpdate() : treeUpdating = True
                End If
                If view IsNot Nothing Then
                    view.BeginUpdate() : viewUpdating = True
                End If
                control.Font = SizedFont(state, control.Font, CSng(Math.Max(MinimumFontPoints, ZoomValue(state, control, "Font", control.Font.SizeInPoints, old, percent))))
                state.Metrics.GetOrCreateValue(control)("Font").LastValue = control.Font.SizeInPoints
                If vertical IsNot Nothing Then
                    vertical.BeginUpdate()
                    Try
                        For Each a As AppearanceObject In vertical.Appearance
                            scaleAppearance(a)
                        Next
                        Dim rowMinimum = MinimumTextHeight(control, vertical.Appearance.RecordValue.GetFont())
                        Dim captionMinimum = MinimumTextHeight(control, vertical.Appearance.RowHeaderPanel.GetFont())
                        Dim headerMinimum = MinimumTextHeight(control, vertical.Appearance.RecordHeader.GetFont())
                        Dim captionLine = Math.Max(1, captionMinimum - Math.Max(4, CInt(Math.Ceiling(control.DeviceDpi / 24.0F))))
                        For Each row In AllRows(vertical.Rows.Cast(Of BaseRow)())
                            scaleAppearance(row.Appearance)
                            scaleAppearance(row.AppearanceHeader)
                            If row.Height > 0 Then row.Height = ZoomSize(state, row, "Height", row.Height, old, percent,
                                Math.Max(rowMinimum, captionMinimum + captionLine * Math.Max(0, row.MaxCaptionLineCount - 1)))
                        Next
                        vertical.RowHeaderWidth = ZoomWidth(state, vertical, "RowHeaderWidth", verticalHeaderWidth, old, percent, baselineBodyFont, 30)
                        vertical.RecordWidth = ZoomWidth(state, vertical, "RecordWidth", verticalRecordWidth, old, percent, baselineBodyFont, 30)
                        vertical.RecordHeaderHeight = ZoomSize(state, vertical, "RecordHeaderHeight", vertical.RecordHeaderHeight, old, percent, Math.Max(18, headerMinimum))
                    Finally
                        vertical.EndUpdate()
                    End Try
                ElseIf tree IsNot Nothing Then
                    tree.BeginUpdate()
                    Try
                        For Each a As AppearanceObject In tree.Appearance
                            scaleAppearance(a)
                        Next
                        Dim rowMinimum = MinimumTextHeight(control, tree.Appearance.Row.GetFont())
                        Dim headerMinimum = MinimumTextHeight(control, tree.Appearance.HeaderPanel.GetFont())
                        For Each column As DevExpress.XtraTreeList.Columns.TreeListColumn In tree.Columns
                            scaleAppearance(column.AppearanceCell) : scaleAppearance(column.AppearanceHeader)
                            column.Width = ZoomWidth(state, column, "Width", treeWidths(column), old, percent, baselineBodyFont, column.MinWidth, If(column.MaxWidth > 0, column.MaxWidth, Integer.MaxValue))
                        Next
                        If tree.RowHeight > 0 Then tree.RowHeight = ZoomSize(state, tree, "RowHeight", tree.RowHeight, old, percent, rowMinimum)
                        If tree.ColumnPanelRowHeight > 0 Then tree.ColumnPanelRowHeight = ZoomSize(state, tree, "ColumnPanelRowHeight", tree.ColumnPanelRowHeight, old, percent, Math.Max(18, headerMinimum))
                    Finally
                        tree.EndUpdate()
                    End Try
                ElseIf view IsNot Nothing Then
                    view.BeginUpdate()
                    Try
                        For Each a As AppearanceObject In view.Appearance
                            scaleAppearance(a)
                        Next
                        Dim rowMinimum = MinimumTextHeight(control, view.Appearance.Row.GetFont())
                        Dim groupMinimum = MinimumTextHeight(control, view.Appearance.GroupRow.GetFont())
                        Dim headerMinimum = MinimumTextHeight(control, view.Appearance.HeaderPanel.GetFont())
                        For Each column In view.Columns.Cast(Of DevExpress.XtraGrid.Columns.GridColumn)()
                            scaleAppearance(column.AppearanceCell) : scaleAppearance(column.AppearanceHeader)
                            column.Width = ZoomWidth(state, column, "Width", gridWidths(column), old, percent, baselineBodyFont, column.MinWidth, If(column.MaxWidth > 0, column.MaxWidth, Integer.MaxValue))
                        Next
                        If view.RowHeight > 0 Then view.RowHeight = ZoomSize(state, view, "RowHeight", view.RowHeight, old, percent, rowMinimum)
                        If view.GroupRowHeight > 0 Then view.GroupRowHeight = ZoomSize(state, view, "GroupRowHeight", view.GroupRowHeight, old, percent, groupMinimum)
                        If view.FooterPanelHeight > 0 Then view.FooterPanelHeight = ZoomSize(state, view, "FooterPanelHeight", view.FooterPanelHeight, old, percent, rowMinimum)
                        If view.ColumnPanelRowHeight > 0 Then view.ColumnPanelRowHeight = ZoomSize(state, view, "ColumnPanelRowHeight", view.ColumnPanelRowHeight, old, percent, Math.Max(18, headerMinimum))
                        Dim banded = TryCast(view, BandedGridView)
                        If banded IsNot Nothing AndAlso banded.BandPanelRowHeight > 0 Then banded.BandPanelRowHeight = ZoomSize(state, banded, "BandPanelRowHeight", banded.BandPanelRowHeight, old, percent, Math.Max(18, headerMinimum))
                    Finally
                        view.EndUpdate()
                    End Try
                End If
                state.Percent = percent
                If view IsNot Nothing Then view.LeftCoord = CInt(leftOffset * ratio)
                If tree IsNot Nothing Then tree.LeftCoord = CInt(leftOffset * ratio)
                If notifyEditors Then NotifyZoomEditors(control, old, percent)
                If persist Then WritePreference(state.Key, "Zoom", percent.ToString())
            Finally
                'Include every repeating header editor in one native layout pass.
                Try
                    If verticalUpdating Then vertical.EndUpdate()
                    If treeUpdating Then tree.EndUpdate()
                    If viewUpdating Then view.EndUpdate()
                    If layoutSuspended Then control.ResumeLayout(True)
                    If verticalUpdating Then RestoreEmptyVerticalViewport(vertical, verticalAnchor)
                    If verticalUpdating Then RefreshVerticalRecordHeaders(vertical)
                Finally
                    state.Applying = False
                End Try
                control.Invalidate(True)
            End Try
        End Sub

        Private Shared Iterator Function AllRows(rows As IEnumerable(Of BaseRow)) As IEnumerable(Of BaseRow)
            For Each row In rows
                Yield row
                For Each child In AllRows(row.ChildRows.Cast(Of BaseRow)())
                    Yield child
                Next
            Next
        End Function

        Private Shared Sub ScaleActiveEditor(control As Control, editor As BaseEdit)
            If editor Is Nothing Then Return
            Dim state As State = Nothing
            If Not States.TryGetValue(control, state) OrElse state.Percent = 100 Then Return
            'Active editor properties are a private copy, not a repository shared
            'with a different grid. The drop-down follows the text size as well.
            'Keep the editor's font family/style (including workbook bold/italic).
            Dim original = editor.Properties.Appearance.GetFont()
            Dim vertical = TryCast(control, VGridControl)
            Dim view = TryCast(TryCast(control, GridControl)?.MainView, GridView)
            Dim tree = TryCast(control, TreeList)
            Dim size = If(vertical IsNot Nothing, vertical.Appearance.RecordValue.GetFont().SizeInPoints,
                          If(view IsNot Nothing, view.Appearance.Row.GetFont().SizeInPoints,
                             If(tree IsNot Nothing, tree.Appearance.Row.GetFont().SizeInPoints, control.Font.SizeInPoints)))
            Dim key = original.FontFamily.Name & "/" & CInt(original.Style).ToString() & "/" & size.ToString(Globalization.CultureInfo.InvariantCulture)
            Dim font As Font = Nothing
            If Not state.EditorFonts.TryGetValue(key, font) Then
                font = New Font(original.FontFamily, size, original.Style, GraphicsUnit.Point)
                state.EditorFonts.Add(key, font) : state.Fonts.Add(font)
            End If
            editor.Properties.Appearance.Font = font
            editor.Properties.Appearance.FontSizeDelta = 0
            editor.Properties.Appearance.Options.UseFont = True
            Dim combo = TryCast(editor.Properties, DevExpress.XtraEditors.Repository.RepositoryItemComboBox)
            If combo IsNot Nothing Then
                combo.AppearanceDropDown.Font = font
                combo.AppearanceDropDown.FontSizeDelta = 0
                combo.AppearanceDropDown.Options.UseFont = True
            End If
        End Sub

        Friend Shared Sub AddMenu(control As Control)
            If control Is Nothing OrElse control.IsDisposed OrElse control.Disposing Then Return
            Dim menu = control.ContextMenuStrip
            If menu IsNot Nothing AndAlso menu.IsDisposed Then Return
            If menu Is Nothing Then
                menu = New ContextMenuStrip()
                control.ContextMenuStrip = menu
            End If
            If menu.Items.Find("SummitGridZoom", False).Length > 0 Then Return
            Dim root As New ToolStripMenuItem("Zoom — " & ZoomPercent(MenuGrid(menu, control)).ToString() & "% (Ctrl + wheel)") With {.Name = "SummitGridZoom"}
            For Each stepValue In New Integer() {10, -10, 0}
                Dim amount = stepValue
                root.DropDownItems.Add(New ToolStripMenuItem(If(amount = 0, "Reset zoom (100%)", If(amount > 0, "Zoom in", "Zoom out")), Nothing,
                    Sub()
                        Dim target = MenuGrid(menu, control)
                        If amount = 0 Then
                            ResetZoom(target)
                        Else
                            SetZoomAt(target, ZoomPercent(target) + amount, MenuTargets.GetOrCreateValue(menu).Point)
                        End If
                    End Sub))
            Next
            Dim presets As New ToolStripMenuItem("Set zoom")
            For Each value In New Integer() {60, 75, 100, 125, 150, 200}
                Dim amount = value
                presets.DropDownItems.Add(New ToolStripMenuItem(amount.ToString() & "%", Nothing,
                    Sub() SetZoomAt(MenuGrid(menu, control), amount, MenuTargets.GetOrCreateValue(menu).Point)))
            Next
            root.DropDownItems.Add(presets)
            If menu.Items.Count > 0 Then menu.Items.Add(New ToolStripSeparator())
            menu.Items.Add(root)
            Dim menuState = MenuTargets.GetOrCreateValue(menu)
            If Not menuState.OpeningAttached Then
                menuState.OpeningAttached = True
                AddHandler menu.Opening, Sub()
                    Dim target = MenuGrid(menu, control)
                    If target IsNot Nothing AndAlso Not target.IsDisposed Then menuState.Point = target.PointToClient(Control.MousePosition)
                    Dim item = menu.Items.Find("SummitGridZoom", False).FirstOrDefault()
                    If item IsNot Nothing Then item.Text = "Zoom — " & ZoomPercent(MenuGrid(menu, control)).ToString() & "% (Ctrl + wheel)"
                End Sub
            End If
        End Sub

        Friend Shared Function ModifiedWheel(control As Control, delta As Integer, modifiers As Integer) As Boolean
            If control Is Nothing OrElse delta = 0 Then Return False
            If (modifiers And 8) <> 0 Then
                QueueWheelZoom(control, delta)
                Return True
            End If
            If (modifiers And 4) <> 0 Then
                Dim vertical = TryCast(control, VGridControl)
                If vertical IsNot Nothing Then
                    vertical.LeftVisibleRecord = Math.Max(0, vertical.LeftVisibleRecord - Math.Sign(delta))
                ElseIf TypeOf control Is TreeList Then
                    Dim tree = DirectCast(control, TreeList)
                    tree.LeftCoord = Math.Max(0, tree.LeftCoord - Math.Sign(delta) * 80)
                Else
                    Dim view = TryCast(TryCast(control, GridControl)?.MainView, GridView)
                    If view Is Nothing Then Return False
                    view.LeftCoord = Math.Max(0, view.LeftCoord - Math.Sign(delta) * 80)
                End If
                Return True
            End If
            Return False
        End Function

        Friend Shared Sub ScrollVerticalRows(grid As VGridControl, delta As Integer)
            Configure(grid)
            Dim state = States.GetOrCreateValue(grid)
            state.VerticalWheelRemainder += delta
            Dim notches = state.VerticalWheelRemainder \ 120
            state.VerticalWheelRemainder -= notches * 120
            If notches = 0 Then Return
            Dim lines = SystemInformation.MouseWheelScrollLines
            If lines < 0 Then lines = Math.Max(1, grid.ViewInfo.RowsViewInfo.Count - 1)
            If lines = 0 Then Return
            'Use the supported native row viewport. Forwarded WM_MOUSEWHEEL can
            'be swallowed by a rebuilt VGrid's editor/controller after filtering.
            'Keep focus, selection and the sticky record header unchanged.
            grid.TopVisibleRowIndex = Math.Max(0, grid.TopVisibleRowIndex - notches * lines)
        End Sub

        Private Class GridWheelFilter
            Implements IMessageFilter
            Public Function PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
                If m.Msg <> &H20A Then Return False
                Dim target = Control.FromChildHandle(m.HWnd)
                While target IsNot Nothing AndAlso Not IsGrid(target)
                    'An open editor owns wheel input (particularly a popup).
                    Dim popup = TryCast(target, PopupBaseEdit)
                    If popup IsNot Nothing AndAlso popup.IsPopupOpen Then Return False
                    target = target.Parent
                End While
                If target Is Nothing OrElse Not target.Enabled OrElse Not target.Visible Then Return False
                Dim form = TryCast(target.TopLevelControl, Form)
                If form IsNot Nothing AndAlso (Not form.Enabled OrElse (Form.ActiveForm IsNot Nothing AndAlso Form.ActiveForm IsNot form)) Then Return False
                If Not target.RectangleToScreen(target.ClientRectangle).Contains(Control.MousePosition) Then Return False
                Dim delta = CInt((m.WParam.ToInt64() >> 16) And &HFFFF)
                If delta >= &H8000 Then delta -= &H10000
                Return ModifiedWheel(target, delta, CInt(m.WParam.ToInt64() And &HFFFF))
            End Function
        End Class
    End Class
End Namespace
