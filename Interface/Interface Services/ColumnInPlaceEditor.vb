
Imports Abovo.GeneralFunctions
Imports Abovo.AbovoExtendedDEControls
Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports Abovo.DataObject

Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraEditors.ViewInfo
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Drawing
Imports DevExpress.XtraGrid.Views.BandedGrid
Imports DevExpress.XtraGrid.Views.BandedGrid.ViewInfo
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports Microsoft.VisualBasic

Namespace Abovo

    Friend Module InplaceEditorFormatting

        Private NotInheritable Class PendingLayout
            Public Queued As Boolean
        End Class

        Private ReadOnly PendingLayouts As New System.Runtime.CompilerServices.ConditionalWeakTable(Of Control, PendingLayout)
        Private ReadOnly HeaderMeasurements As New System.Runtime.CompilerServices.ConditionalWeakTable(Of Control, Dictionary(Of String, Single))

        Friend Sub QueuePresentationLayout(control As Control)
            If control Is Nothing OrElse control.IsDisposed OrElse Not control.IsHandleCreated Then Return
            Dim pending = PendingLayouts.GetValue(control, Function(key) New PendingLayout())
            If pending.Queued Then Return
            pending.Queued = True
            control.BeginInvoke(New MethodInvoker(Sub()
                pending.Queued = False
                If control.IsDisposed OrElse control.Disposing Then Return
                Dim vertical = TryCast(control, DevExpress.XtraVerticalGrid.VGridControl)
                If vertical IsNot Nothing Then
                    vertical.LayoutChanged()
                Else
                    Dim grid = TryCast(control, DevExpress.XtraGrid.GridControl)
                    If grid IsNot Nothing Then
                        For Each view As GridView In grid.ViewCollection.OfType(Of GridView)().ToArray()
                            view.LayoutChanged()
                        Next
                    End If
                End If
                control.Invalidate()
            End Sub))
        End Sub

        Friend Function OwnerFontSize(owner As Control) As Single
            Dim vertical = TryCast(owner, DevExpress.XtraVerticalGrid.VGridControl)
            If vertical IsNot Nothing Then Return vertical.Appearance.RecordValue.GetFont().SizeInPoints
            Dim view = TryCast(TryCast(owner, DevExpress.XtraGrid.GridControl)?.MainView, GridView)
            Return If(view Is Nothing, owner.Font.SizeInPoints, view.Appearance.Row.GetFont().SizeInPoints)
        End Function

        Friend Sub SetEditorFontSize(item As RepositoryItem, editor As BaseEdit, size As Single)
            If item Is Nothing OrElse item.Appearance.Font Is Nothing OrElse size <= 0.0F Then Return
            If HasEditorFontSize(item, size) AndAlso
               (editor Is Nothing OrElse editor.IsDisposed OrElse HasEditorFontSize(editor.Properties, size)) Then Return
            item.BeginUpdate()
            Try
                For Each appearance In New AppearanceObject() {item.Appearance, item.AppearanceFocused, item.AppearanceReadOnly, item.AppearanceDisabled}
                    If appearance IsNot item.Appearance AndAlso Not appearance.Options.UseFont Then Continue For
                    Dim current = appearance.GetFont()
                    If Math.Abs(current.SizeInPoints - size) < 0.01F AndAlso appearance.FontSizeDelta = 0 Then Continue For
                    appearance.Font = New Font(current.FontFamily, Math.Max(1.0F, size), current.Style, GraphicsUnit.Point)
                    appearance.FontSizeDelta = 0
                    appearance.Options.UseFont = True
                Next
                Dim combo = TryCast(item, RepositoryItemComboBox)
                If combo IsNot Nothing Then CopyEditorFont(item.Appearance, combo.AppearanceDropDown)
            Finally
                item.EndUpdate()
            End Try
            'A header editor is a separate control, not the grid's ActiveEditor.
            'Resize its presentation without committing, recreating or losing its input.
            If editor Is Nothing OrElse editor.IsDisposed Then Return
            Dim target = editor.Properties
            target.BeginUpdate()
            Try
                CopyEditorFont(item.Appearance, target.Appearance)
                CopyEditorFont(item.AppearanceFocused, target.AppearanceFocused)
                CopyEditorFont(item.AppearanceReadOnly, target.AppearanceReadOnly)
                CopyEditorFont(item.AppearanceDisabled, target.AppearanceDisabled)
                Dim combo = TryCast(target, RepositoryItemComboBox)
                If combo IsNot Nothing Then CopyEditorFont(item.Appearance, combo.AppearanceDropDown)
            Finally
                target.EndUpdate()
            End Try
        End Sub

        Private Function HasEditorFontSize(item As RepositoryItem, size As Single) As Boolean
            For Each appearance In New AppearanceObject() {item.Appearance, item.AppearanceFocused, item.AppearanceReadOnly, item.AppearanceDisabled}
                If appearance IsNot item.Appearance AndAlso Not appearance.Options.UseFont Then Continue For
                If Not appearance.Options.UseFont OrElse appearance.FontSizeDelta <> 0 OrElse Math.Abs(appearance.GetFont().SizeInPoints - size) >= 0.01F Then Return False
            Next
            Dim combo = TryCast(item, RepositoryItemComboBox)
            Return combo Is Nothing OrElse (combo.AppearanceDropDown.Options.UseFont AndAlso Math.Abs(combo.AppearanceDropDown.GetFont().SizeInPoints - size) < 0.01F)
        End Function

        Private Sub CopyEditorFont(source As AppearanceObject, target As AppearanceObject)
            target.Font = source.Font
            target.FontSizeDelta = source.FontSizeDelta
            target.Options.UseFont = source.Options.UseFont
        End Sub

        Private Function MeasurementKey(item As RepositoryItem, graphics As Graphics, size As Single) As String
            Dim font = item.Appearance.GetFont()
            Return item.GetType().FullName & "/" & font.Name & "/" &
                size.ToString("R", Globalization.CultureInfo.InvariantCulture) & "/" &
                CInt(font.Style).ToString() & "/" & CInt(item.BorderStyle).ToString() & "/" &
                graphics.DpiX.ToString("R", Globalization.CultureInfo.InvariantCulture) & "/" &
                graphics.DpiY.ToString("R", Globalization.CultureInfo.InvariantCulture) & "/" & item.LookAndFeel.ActiveSkinName
        End Function

        Private Function MeasurementItem(item As RepositoryItem, font As Font) As RepositoryItem
            'A view-info Appearance can be the repository's own Appearance.
            'Measure a disposable clone: assigning/discarding a temporary Font
            'through the live view-info would corrupt the real editor font.
            Dim measured = DirectCast(item.Clone(), RepositoryItem)
            measured.BeginUpdate()
            Try
                measured.AutoHeight = False
                For Each appearance In New AppearanceObject() {measured.Appearance, measured.AppearanceFocused, measured.AppearanceReadOnly, measured.AppearanceDisabled}
                    appearance.Font = font
                    appearance.FontSizeDelta = 0
                    appearance.Options.UseFont = True
                Next
            Finally
                measured.EndUpdate()
            End Try
            Return measured
        End Function

        Friend Function DesiredEditorHeight(item As RepositoryItem, owner As Control, Optional graphics As Graphics = Nothing) As Integer
            If graphics Is Nothing Then
                Using ownerGraphics = owner.CreateGraphics()
                    Return DesiredEditorHeight(item, owner, ownerGraphics)
                End Using
            End If
            Dim size = OwnerFontSize(owner)
            Dim key = MeasurementKey(item, graphics, size) & "/height"
            Dim cache = HeaderMeasurements.GetValue(owner, Function(control) New Dictionary(Of String, Single)(StringComparer.Ordinal))
            Dim height As Single
            If Not cache.TryGetValue(key, height) Then
                Using font As New Font(item.Appearance.Font.FontFamily, size, item.Appearance.GetFont().Style, GraphicsUnit.Point)
                    Using measured = MeasurementItem(item, font)
                        Dim info = measured.CreateViewInfo()
                        info.PaintAppearance = measured.Appearance
                        'Measure on the real owner's device context, not a newly
                        'created unattached editor with a different DPI/density.
                        height = Math.Max(info.CalcBestFit(graphics).Height, CInt(Math.Ceiling(font.GetHeight(graphics))) + 4)
                        Dim textHeight = Math.Max(font.GetHeight(graphics), TextRenderer.MeasureText(graphics, "Ag", font, System.Drawing.Size.Empty, TextFormatFlags.NoPadding Or TextFormatFlags.SingleLine).Height)
                        For attempt As Integer = 0 To 2
                            info.Bounds = New Rectangle(0, 0, 500, CInt(height))
                            info.CalcViewInfo(graphics)
                            Dim missing = CInt(Math.Ceiling(textHeight - info.GetTextBounds().Height))
                            If missing <= 0 Then Exit For
                            height += missing
                        Next
                    End Using
                End Using
                cache.Add(key, height)
            End If
            Return CInt(height)
        End Function

        Friend Function DesiredEditorWidth(item As RepositoryItem, owner As Control, graphics As Graphics, value As Object) As Integer
            Dim size = OwnerFontSize(owner)
            Dim sample = If(TypeOf item Is RepositoryItemDateEdit, DirectCast(New DateTime(2051, 9, 30), Object), value)
            Dim text = item.GetDisplayText(sample)
            If String.IsNullOrWhiteSpace(text) Then text = "Year 999"
            Dim key = MeasurementKey(item, graphics, size) & "/width/" & text
            Dim cache = HeaderMeasurements.GetValue(owner, Function(control) New Dictionary(Of String, Single)(StringComparer.Ordinal))
            Dim width As Single
            If Not cache.TryGetValue(key, width) Then
                Using font As New Font(item.Appearance.Font.FontFamily, size, item.Appearance.GetFont().Style, GraphicsUnit.Point)
                    Using measured = MeasurementItem(item, font)
                        Dim info = measured.CreateViewInfo()
                        info.PaintAppearance = measured.Appearance
                        info.EditValue = sample
                        info.Bounds = New Rectangle(0, 0, 1000, DesiredEditorHeight(item, owner, graphics))
                        info.CalcViewInfo(graphics)
                        Dim chrome = Math.Max(0, info.Bounds.Width - info.GetTextBounds().Width)
                        width = chrome + TextRenderer.MeasureText(graphics, text, font, System.Drawing.Size.Empty, TextFormatFlags.NoPadding Or TextFormatFlags.SingleLine).Width + 2
                    End Using
                End Using
                cache.Add(key, width)
            End If
            Return CInt(width)
        End Function

        Friend Sub FitEditorToBounds(item As RepositoryItem, editor As BaseEdit, owner As Control,
                                     graphics As Graphics, bounds As Rectangle)
            If bounds.Width <= 0 OrElse bounds.Height <= 0 Then Return
            Dim wanted = OwnerFontSize(owner)
            Dim key = MeasurementKey(item, graphics, wanted) & "/fit/" & bounds.Width.ToString() & "/" & bounds.Height.ToString()
            Dim cache = HeaderMeasurements.GetValue(owner, Function(control) New Dictionary(Of String, Single)(StringComparer.Ordinal))
            Dim fitted As Single
            If Not cache.TryGetValue(key, fitted) Then
                fitted = wanted
                Using font As New Font(item.Appearance.Font.FontFamily, wanted, item.Appearance.GetFont().Style, GraphicsUnit.Point)
                    Using measured = MeasurementItem(item, font)
                        Dim info = measured.CreateViewInfo()
                        info.PaintAppearance = measured.Appearance
                        info.Bounds = New Rectangle(Point.Empty, bounds.Size)
                        info.CalcViewInfo(graphics)
                        Dim available = Math.Max(1, Math.Min(bounds.Height - 2, info.GetTextBounds().Height))
                        Dim textHeight = Math.Max(font.GetHeight(graphics), TextRenderer.MeasureText(graphics, "Ag", font, Size.Empty, TextFormatFlags.NoPadding Or TextFormatFlags.SingleLine).Height)
                        'GDI text rectangles round to integer pixels; preserve
                        'fractional point sizes when only that rounding differs.
                        If textHeight > available + 1.0F Then fitted = Math.Max(Math.Min(8.0F, wanted), CSng(Math.Floor(wanted * available / textHeight * 20.0F) / 20.0F))
                    End Using
                End Using
                cache.Add(key, fitted)
            End If
            'Density and local zoom are already combined in the owner's font.
            'Never multiply the last header font: it may belong to another
            'window size, or may have been fitted into a shorter row.
            SetEditorFontSize(item, editor, fitted)
        End Sub

        Public Function IsPopupShortcut(keyData As Keys) As Boolean
            Dim modifiers = keyData And Keys.Modifiers
            Dim key = keyData And Keys.KeyCode
            Return (key = Keys.F4 AndAlso modifiers = Keys.None) OrElse
                (key = Keys.Down AndAlso (modifiers = Keys.Shift OrElse modifiers = Keys.Alt))
        End Function

        Public Function OpenPopupForShortcut(editor As BaseEdit, e As KeyEventArgs) As Boolean
            If Not IsPopupShortcut(e.KeyData) Then Return False
            Dim popup = TryCast(editor, PopupBaseEdit)
            If popup Is Nothing OrElse popup.IsDisposed OrElse popup.Properties.ReadOnly OrElse Not popup.Enabled Then Return False
            e.Handled = True
            e.SuppressKeyPress = True
            If Not popup.IsPopupOpen Then popup.ShowPopup()
            Return True
        End Function

        Public Function SameEditorValue(left As Object, right As Object) As Boolean
            Dim normalise As Func(Of Object, String) = Function(value)
                If TypeOf value Is OrdinalYearComboItem Then value = DirectCast(value, OrdinalYearComboItem).StoredValue
                Dim text = Convert.ToString(value, Globalization.CultureInfo.InvariantCulture)
                Return If(text.Trim().Equals("<Blank>", StringComparison.OrdinalIgnoreCase), "", text)
            End Function
            Return normalise(left) = normalise(right)
        End Function

        Public Sub ApplyStandardDateFormat(ByVal Item As RepositoryItem)

            Dim DateItem As RepositoryItemDateEdit = TryCast(Item, RepositoryItemDateEdit)
            If DateItem Is Nothing Then Exit Sub

            'Keep the established long display text, but use the client's compact
            'numeric format while the temporary editor has focus.
            With DateItem
                .DisplayFormat.FormatType = FormatType.DateTime
                .DisplayFormat.FormatString = "dd-MMM-yyyy"
                .EditFormat.FormatType = FormatType.DateTime
                .EditFormat.FormatString = "dd/MM/yy"
                .Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTimeAdvancingCaret
                .Mask.EditMask = "dd/MM/yy"
                .UseMaskAsDisplayFormat = False
            End With

        End Sub

    End Module

    Public Class ColumnInplaceEditorHelper
        Public Property Navigate As Func(Of Keys, Boolean)
        Public ReadOnly Property GridControl As Control
            Get
                Return bgview?.GridControl
            End Get
        End Property

        Private _Item As RepositoryItem
        Private _Column As BandedGridColumn
        Private bgview As GridView
        Private _EditorHeight As Integer = -1
        Private _AdvanceScheduled As Boolean = False
        Private _ValueChangedHandler As EventHandler
        Private _DoubleClickHandler As EventHandler
        Private _CommitInProgress As Boolean = False
        Private _ClosingEditor As Boolean = False
        Private _DoubleClickScheduled As Boolean = False

        'Exact editor rectangle used by the most recent custom header paint.
        'Mouse interaction must use the same rectangle rather than trying to
        'reconstruct it later from BandedGridViewInfo.ColumnsInfo.
        Private _LastPaintedEditorBounds As Rectangle = Rectangle.Empty

        Public LinkedComboBoxEdit As AbovoDEHeaderComboBox
        Public LinkedDateEdit As AbovoDEHeaderDateBox
        Public Tag As Object

        Public Sub New(ByVal column As BandedGridColumn,
                       ByVal inplaceEditor As RepositoryItem,
                       Optional ByVal valueChangedHandler As EventHandler = Nothing,
                       Optional ByVal doubleClickHandler As EventHandler = Nothing)
            If column Is Nothing Then Throw New ArgumentNullException(NameOf(column))
            If inplaceEditor Is Nothing Then Throw New ArgumentNullException(NameOf(inplaceEditor))
            _Column = column
            _Item = inplaceEditor
            _ValueChangedHandler = valueChangedHandler
            _DoubleClickHandler = doubleClickHandler
            bgview = TryCast(column.View, BandedGridView)
            If bgview Is Nothing Then Throw New ArgumentException("The header column must belong to a banded view.", NameOf(column))

            InplaceEditorFormatting.ApplyStandardDateFormat(_Item)

            'Lazy sections build their columns before attaching the view to a
            'GridControl. Measure against the real owner on the first paint or
            'zoom notification; construction must not assume it exists yet.
            _EditorHeight = If(bgview.GridControl Is Nothing,
                               Math.Max(1, _Item.Appearance.GetFont().Height + 8),
                               InplaceEditorFormatting.DesiredEditorHeight(_Item, bgview.GridControl))
            '_ActiveEditor = New BaseEdit
            '_ActiveEditor.ForeColor = Color.Red
            AddHandler bgview.CustomDrawColumnHeader, AddressOf view_CustomDrawColumnHeader
            AddHandler bgview.MouseDown, AddressOf view_MouseDown
            AddHandler bgview.Layout, AddressOf view_Layout
            AddHandler PresentationScaleManager.ScaleChanged, AddressOf PresentationScaleChanged
            AddHandler GridPresentation.ZoomChanged, AddressOf GridZoomChanged
            AddHandler bgview.Disposed, Sub()
                RemoveHandler PresentationScaleManager.ScaleChanged, AddressOf PresentationScaleChanged
                RemoveHandler GridPresentation.ZoomChanged, AddressOf GridZoomChanged
            End Sub
        End Sub

        Private Sub GridZoomChanged(sender As Object, e As PresentationScaleChangedEventArgs)
            If sender Is bgview?.GridControl Then PresentationScaleChanged(sender, e)
        End Sub

        Private Sub PresentationScaleChanged(
            ByVal sender As Object,
            ByVal e As PresentationScaleChangedEventArgs)

            'An unattached lazy view is temporary, not disposed: keep the
            'subscription so it receives later scale changes after attachment.
            If bgview Is Nothing OrElse bgview.GridControl Is Nothing Then Return
            If bgview.GridControl.IsDisposed Then
                RemoveHandler PresentationScaleManager.ScaleChanged, AddressOf PresentationScaleChanged
                RemoveHandler GridPresentation.ZoomChanged, AddressOf GridZoomChanged
                Return
            End If

            Dim isGridZoom = sender Is bgview.GridControl
            InplaceEditorFormatting.SetEditorFontSize(_Item, ActiveEditor, InplaceEditorFormatting.OwnerFontSize(bgview.GridControl))
            _EditorHeight = InplaceEditorFormatting.DesiredEditorHeight(_Item, bgview.GridControl)
            If bgview.ColumnPanelRowHeight > 0 AndAlso bgview.ColumnPanelRowHeight < _EditorHeight + 4 Then
                bgview.ColumnPanelRowHeight = _EditorHeight + 4
                GridPresentation.AcceptRenderedMetric(bgview.GridControl, bgview, "ColumnPanelRowHeight", bgview.ColumnPanelRowHeight)
            End If
            Using graphics = bgview.GridControl.CreateGraphics()
                Dim probe = New Rectangle(0, 0, 1000, _EditorHeight + 4)
                Dim reserved = probe.Width - DrawEditorHelper.GetEditorBounds(probe, GetRightIndent(), _EditorHeight).Width
                Dim value = If(ActiveEditor IsNot Nothing AndAlso Not ActiveEditor.IsDisposed, ActiveEditor.EditValue, EditValue)
                Dim minimumWidth = reserved + InplaceEditorFormatting.DesiredEditorWidth(_Item, bgview.GridControl, graphics, value)
                If _Column.Width < minimumWidth Then
                    _Column.Width = minimumWidth
                    GridPresentation.AcceptRenderedMetric(bgview.GridControl, _Column, "Width", _Column.Width)
                End If
            End Using
            _LastPaintedEditorBounds = Rectangle.Empty
            'Local zoom has one outer BeginUpdate/EndUpdate for the whole grid.
            'Global scaling also needs only one layout, not one per header editor.
            If Not isGridZoom Then InplaceEditorFormatting.QueuePresentationLayout(bgview.GridControl)
        End Sub

        Private Sub view_Layout(ByVal sender As Object, ByVal e As EventArgs)
            If Not GridPresentation.IsApplyingZoom(bgview?.GridControl) Then CommitAndCloseEditor()

            'The next CustomDrawColumnHeader will replace this with the new
            'authoritative painted rectangle.
            _LastPaintedEditorBounds = Rectangle.Empty

        End Sub

        Private _EditValue As Object
        Public Property EditValue() As Object
            Get
                Return _EditValue
            End Get
            Set(ByVal value As Object)
                _EditValue = value
                If _ActiveEditor IsNot Nothing AndAlso Not _ActiveEditor.IsDisposed Then
                    If _CommitInProgress OrElse Not _ActiveEditor.ContainsFocus Then _ActiveEditor.EditValue = value
                End If
                If bgview IsNot Nothing AndAlso bgview.GridControl IsNot Nothing AndAlso
                   Not bgview.GridControl.IsDisposed Then
                    bgview.GridControl.Invalidate()
                End If
            End Set
        End Property

        Private _ActiveEditor As BaseEdit
        Public Property ActiveEditor() As BaseEdit
            Get
                Return _ActiveEditor
            End Get
            Set(ByVal value As BaseEdit)
                _ActiveEditor = value
            End Set
        End Property

        Private Sub view_CustomDrawColumnHeader(ByVal sender As Object, ByVal e As ColumnHeaderCustomDrawEventArgs)
            If e.Column Is Nothing Then Return
            Dim ColTag As DataColumnTag = TryCast(e.Column.Tag, DataColumnTag)
            If ColTag IsNot Nothing AndAlso Not ColTag.HasIncolumnEditor Then
                e.Handled = True
                Return
            End If

            If e.Column.AbsoluteIndex = 0 Then Return
            If e.Column Is _Column Then
                'e.Appearance.Options.UseBackColor = True
                'e.Appearance.BackColor = AbovoComboBGC
                'e.Info.Caption = String.Empty
                'Dim br As New SolidBrush(AbovoComboBGC)
                'e.Cache.FillRectangle(br, e.Bounds)
                'e.DefaultDraw()
                'Paint the normal header over the full multi-line bounds, but suppress
                'the placeholder caption. The editor itself is then painted only in its
                'fixed single-line rectangle at the bottom of the header.
                e.Info.Caption = String.Empty
                e.Appearance.Options.UseBackColor = True
                e.Appearance.BackColor = Color.White
                e.Appearance.Options.UseForeColor = True
                e.Appearance.ForeColor = AbovoBlue
                e.Painter.DrawObject(e.Info)

                'Cache the EXACT rectangle that is being painted. Complex banded
                'headers can have several logical column rectangles at the same X
                'position, and ColumnsInfo may not describe the custom-painted
                'editor rectangle in the same way as CustomDrawColumnHeader.
                _EditorHeight = InplaceEditorFormatting.DesiredEditorHeight(_Item, bgview.GridControl, e.Graphics)
                _LastPaintedEditorBounds =
                    DrawEditorHelper.GetEditorBounds(
                        e.Bounds,
                        GetRightIndent(),
                        _EditorHeight)

                InplaceEditorFormatting.FitEditorToBounds(_Item, ActiveEditor, bgview.GridControl, e.Graphics, _LastPaintedEditorBounds)

                If ActiveEditor IsNot Nothing AndAlso Not ActiveEditor.IsDisposed AndAlso
                   ActiveEditor.Bounds <> _LastPaintedEditorBounds Then
                    ActiveEditor.Bounds = _LastPaintedEditorBounds
                End If

                DrawEditorHelper.DrawColumnInplaceEditor(e, _Item, EditValue, GetRightIndent(), _EditorHeight, useRepositoryAppearance:=True)
                e.Handled = True

            End If
        End Sub

        Public Function GetRightIndent() As Integer
            If _Column.OptionsColumn.AllowSort <> DevExpress.Utils.DefaultBoolean.False OrElse _Column.OptionsFilter.AllowFilter Then
                Return PresentationScaleManager.Scale(25)
            Else
                Return 0
            End If
        End Function
        Private Sub view_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
            CommitAndCloseEditor()
            Dim editorBounds As Rectangle
            If ClickInEditor(e, editorBounds) Then
                ShowEditor(editorBounds)
                If e.Button = MouseButtons.Left AndAlso e.Clicks >= 2 Then
                    RaiseEditorDoubleClick(ActiveEditor)
                End If
                DXMouseEventArgs.GetMouseArgs(e).Handled = True
            End If
        End Sub
        Private Function ClickInEditor(ByVal e As MouseEventArgs,
                                       <System.Runtime.InteropServices.Out()> ByRef editorBounds As Rectangle) As Boolean

            editorBounds = Rectangle.Empty

            'The editor's hot state is painted from CustomDrawColumnHeader, so the
            'most reliable click test is the exact rectangle used by that paint.
            '
            'Do NOT derive the rectangle again from ColumnsInfo. In a complex
            'BandedGridView the same horizontal position may represent several
            'columns/header rows, and the view-info collection can return a
            'different logical header than the one whose custom editor is visible.
            If Not _LastPaintedEditorBounds.IsEmpty AndAlso
               _LastPaintedEditorBounds.Contains(e.Location) Then

                editorBounds = _LastPaintedEditorBounds
                Return True

            End If

            'Fallback for the very unusual case where MouseDown occurs before the
            'first custom paint after a layout change.
            Dim vi As BandedGridViewInfo =
                TryCast(bgview.GetViewInfo(), BandedGridViewInfo)

            If vi Is Nothing Then Return False

            Dim columnInfo As GridColumnInfoArgs =
                CalcColumnHitInfo(
                    e.Location,
                    vi.ColumnsInfo)

            If columnInfo IsNot Nothing AndAlso
               columnInfo.Column Is _Column Then

                editorBounds =
                    DrawEditorHelper.GetEditorBounds(
                        columnInfo.Bounds,
                        GetRightIndent(),
                        _EditorHeight)

                Return editorBounds.Contains(e.Location)

            End If

            Return False

        End Function

        Private Function GetCurrentEditorBounds() As Rectangle

            If Not _LastPaintedEditorBounds.IsEmpty Then
                Return _LastPaintedEditorBounds
            End If

            Dim vi As BandedGridViewInfo =
                TryCast(bgview.GetViewInfo(), BandedGridViewInfo)

            If vi Is Nothing Then Return Rectangle.Empty

            For Each columnInfo As GridColumnInfoArgs In vi.ColumnsInfo

                If columnInfo IsNot Nothing AndAlso
                   columnInfo.Column Is _Column AndAlso
                   Not columnInfo.Bounds.IsEmpty Then

                    Return DrawEditorHelper.GetEditorBounds(
                        columnInfo.Bounds,
                        GetRightIndent(),
                        _EditorHeight)

                End If

            Next

            Return Rectangle.Empty

        End Function

        Private Shared Function GetHelper(ByVal column As BandedGridColumn) As ColumnInplaceEditorHelper

            If column Is Nothing Then Return Nothing

            Dim columnTag As DataColumnTag = TryCast(column.Tag, DataColumnTag)

            If columnTag Is Nothing OrElse Not columnTag.HasIncolumnEditor Then
                Return Nothing
            End If

            If columnTag.InColumnEditorCombo IsNot Nothing Then
                Return columnTag.InColumnEditorCombo.InPlaceColumnHelper
            End If

            If columnTag.InColumnEditorDate IsNot Nothing Then
                Return columnTag.InColumnEditorDate.InPlaceColumnHelper
            End If

            Return Nothing

        End Function

        Private Function GetAdjacentVisibleHelper(ByVal movePrevious As Boolean) As ColumnInplaceEditorHelper

            'BandedGridColumn.VisibleIndex is not a stable traversal key when
            'columns occupy band rows; AbsoluteIndex follows the defining-row
            'source order used when these header editors are constructed.
            Dim currentColumnIndex As Integer = _Column.AbsoluteIndex
            Dim adjacentColumnIndex As Integer =
                If(movePrevious, Integer.MinValue, Integer.MaxValue)
            Dim adjacentHelper As ColumnInplaceEditorHelper = Nothing

            For Each gridColumn As GridColumn In bgview.Columns

                Dim candidate As BandedGridColumn = TryCast(gridColumn, BandedGridColumn)

                If candidate Is Nothing OrElse Not candidate.Visible Then Continue For

                If movePrevious Then

                    If candidate.AbsoluteIndex >= currentColumnIndex OrElse
                       candidate.AbsoluteIndex <= adjacentColumnIndex Then

                        Continue For
                    End If

                Else

                    If candidate.AbsoluteIndex <= currentColumnIndex OrElse
                       candidate.AbsoluteIndex >= adjacentColumnIndex Then

                        Continue For
                    End If

                End If

                Dim candidateHelper As ColumnInplaceEditorHelper = GetHelper(candidate)

                If candidateHelper IsNot Nothing Then
                    adjacentColumnIndex = candidate.AbsoluteIndex
                    adjacentHelper = candidateHelper
                End If

            Next

            Return adjacentHelper

        End Function

        Protected Overridable Function CalcColumnHitInfo(ByVal pt As Point, ByVal cols As GridColumnsInfo) As GridColumnInfoArgs

            'BandedGridView can contain more than one column at the same X
            'position because columns may occupy different vertical rows within
            'a band. The old implementation tested X only and therefore returned
            'the first overlapping column, even when the mouse was actually over
            'a different header row.
            '
            'That made in-header editors appear correctly but ignore clicks on
            'multi-row banded interfaces such as Rent Assumptions > Voids/Bad
            'Debt and Service Charge Assumptions. Simpler single-row layouts such
            'as Stock Assumptions > New Lettings happened to work.
            '
            'Use the complete header rectangle so both X and Y participate in
            'the hit test.
            For Each ci As GridColumnInfoArgs In cols

                If ci Is Nothing OrElse
                   ci.Bounds.IsEmpty OrElse
                   ci.Type = GridColumnInfoType.EmptyColumn Then

                    Continue For

                End If

                If ci.Bounds.Contains(pt) Then
                    Return ci
                End If

            Next

            Return Nothing

        End Function

        Protected Function IntInRange(ByVal x As Integer, ByVal left As Integer, ByVal right As Integer) As Boolean
            If right < left Then
                Dim temp As Integer = left
                left = right
                right = temp
            End If
            Return (x >= left AndAlso x < right)
        End Function

        Private Sub ShowEditor(ByVal bounds As Rectangle,
                               Optional ByVal activateWithMouse As Boolean = True)

            ActiveEditor = _Item.CreateEditor()
            ActiveEditor.Properties.LockEvents()

            'Assign the repository settings first, then force this temporary editor
            'to use exactly the same fixed rectangle as the custom-painted version.
            ActiveEditor.Properties.Assign(_Item)
            ActiveEditor.Properties.AutoHeight = False
            ActiveEditor.EnterMoveNextControl = False

            ActiveEditor.Properties.Appearance.Options.UseBackColor = True
            ActiveEditor.Properties.Appearance.BackColor = _Item.Appearance.BackColor
            ActiveEditor.Properties.Appearance.Options.UseForeColor = True
            ActiveEditor.Properties.Appearance.ForeColor = _Item.Appearance.ForeColor
            ActiveEditor.Properties.Appearance.Options.UseFont = True
            ActiveEditor.Properties.Appearance.Font = _Item.Appearance.Font

            ActiveEditor.Tag = Tag
            ActiveEditor.Parent = bgview.GridControl
            ActiveEditor.Location = bounds.Location
            ActiveEditor.Size = bounds.Size
            ActiveEditor.CreateControl()
            ActiveEditor.EditValue = EditValue

            AddHandler ActiveEditor.Leave, AddressOf editor_Leave
            AddHandler ActiveEditor.KeyDown, AddressOf editor_KeyDown
            AddHandler ActiveEditor.PreviewKeyDown, AddressOf editor_PreviewKeyDown
            AddHandler ActiveEditor.DoubleClick, AddressOf editor_DoubleClick

            ActiveEditor.Properties.UnLockEvents()

            If activateWithMouse Then
                ActiveEditor.SendMouse(ActiveEditor.PointToClient(Control.MousePosition), Control.MouseButtons)
            Else
                ActiveEditor.Focus()
            End If

        End Sub

        Private Sub CloseEditor()

            If ActiveEditor IsNot Nothing AndAlso Not _ClosingEditor AndAlso Not _CommitInProgress Then

                _ClosingEditor = True

                Try
                EditValue = ActiveEditor.EditValue
                RemoveHandler ActiveEditor.Leave, AddressOf editor_Leave
                RemoveHandler ActiveEditor.KeyDown, AddressOf editor_KeyDown
                RemoveHandler ActiveEditor.PreviewKeyDown, AddressOf editor_PreviewKeyDown
                RemoveHandler ActiveEditor.DoubleClick, AddressOf editor_DoubleClick
                ActiveEditor.Dispose()
                ActiveEditor = Nothing
                Finally
                    _ClosingEditor = False
                End Try

            End If

        End Sub

        Private Function CommitActiveEditorValue() As Boolean

            If ActiveEditor Is Nothing Then Return True
            If _CommitInProgress Then Return False

            Dim editor As BaseEdit = ActiveEditor
            _CommitInProgress = True

            Try
                Dim requested = editor.EditValue
                EditValue = requested

                If _ValueChangedHandler IsNot Nothing Then
                    _ValueChangedHandler.Invoke(editor, EventArgs.Empty)
                End If
                Return Not editor.IsDisposed AndAlso InplaceEditorFormatting.SameEditorValue(requested, editor.EditValue)
            Finally
                _CommitInProgress = False
            End Try

        End Function

        Private Sub CommitAndCloseEditor()

            If ActiveEditor Is Nothing OrElse _ClosingEditor Then Return

            If Not _CommitInProgress Then

                If Not ActiveEditor.DoValidate(PopupCloseMode.Normal) Then Return
                If Not CommitActiveEditorValue() Then Return

            End If

            CloseEditor()

        End Sub

        Friend Function CommitForSave() As Boolean
            If ActiveEditor Is Nothing Then Return True
            If _ClosingEditor OrElse _CommitInProgress Then Return False
            If Not ActiveEditor.DoValidate(PopupCloseMode.Normal) OrElse Not CommitActiveEditorValue() Then Return False
            CloseEditor()
            Return True
        End Function

        Private Sub editor_PreviewKeyDown(ByVal sender As Object, ByVal e As PreviewKeyDownEventArgs)

            If {Keys.Tab, Keys.Enter, Keys.Left, Keys.Right, Keys.Up, Keys.Down}.Contains(e.KeyCode) Then e.IsInputKey = True

        End Sub

        Private Sub editor_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)

            If InplaceEditorFormatting.OpenPopupForShortcut(TryCast(sender, BaseEdit), e) Then Return
            If e.Control OrElse e.Alt Then Return
            Dim popup = TryCast(sender, PopupBaseEdit)
            If popup IsNot Nothing AndAlso popup.IsPopupOpen Then Return
            If e.KeyCode <> Keys.Enter AndAlso Not {Keys.Tab, Keys.Left, Keys.Right, Keys.Up, Keys.Down}.Contains(e.KeyCode) Then Return

            e.Handled = True
            e.SuppressKeyPress = True

            Dim editor As BaseEdit = TryCast(sender, BaseEdit)

            If editor Is Nothing OrElse
               Not editor.DoValidate(PopupCloseMode.Normal) Then

                Return

            End If

            If Not CommitActiveEditorValue() Then Return
            If Navigate IsNot Nothing Then
                CloseEditor()
                If Navigate(e.KeyData) Then Return
            End If
            ScheduleAdvanceToEditor(
                GetAdjacentVisibleHelper((e.KeyCode = Keys.Tab OrElse e.KeyCode = Keys.Enter) AndAlso e.Shift))

        End Sub

        Private Sub ScheduleAdvanceToEditor(ByVal adjacentHelper As ColumnInplaceEditorHelper)

            If _AdvanceScheduled Then Return

            If bgview Is Nothing OrElse
               bgview.GridControl Is Nothing OrElse
               bgview.GridControl.IsDisposed Then

                Return

            End If

            _AdvanceScheduled = True

            'Repository-item change handlers calculate and relayout the grid
            'synchronously. Queue one navigation request that survives Layout
            'closing the current temporary editor, then reacquire the next
            'header bounds after UpdateAllRules has completed.
            bgview.GridControl.BeginInvoke(
                New MethodInvoker(
                    Sub()
                        _AdvanceScheduled = False
                        CloseEditor()

                        If adjacentHelper IsNot Nothing Then
                            adjacentHelper.ShowEditorFromKeyboard()
                        End If
                    End Sub))

        End Sub

        Public Sub ShowEditorFromKeyboard()

            If ActiveEditor IsNot Nothing AndAlso Not ActiveEditor.IsDisposed Then
                ActiveEditor.Focus()
                Return
            End If

            If bgview Is Nothing OrElse
               bgview.GridControl Is Nothing OrElse
               bgview.GridControl.IsDisposed OrElse
               Not _Column.Visible Then

                Return

            End If

            'A calculation/layout pass clears the cached custom-draw rectangle.
            'Force the pending paint now so the next helper can use the exact
            'same bounds as its visible header editor.
            If _LastPaintedEditorBounds.IsEmpty Then
                bgview.GridControl.Refresh()
            End If

            Dim bounds As Rectangle = GetCurrentEditorBounds()
            If bounds.IsEmpty Then Return

            ShowEditor(bounds, False)

        End Sub

        Private Sub editor_Leave(ByVal sender As Object, ByVal e As EventArgs)
            CommitAndCloseEditor()
        End Sub

        Private Sub editor_DoubleClick(ByVal sender As Object, ByVal e As EventArgs)
            RaiseEditorDoubleClick(sender)
        End Sub

        Private Sub RaiseEditorDoubleClick(ByVal sender As Object)

            If _DoubleClickHandler Is Nothing OrElse _DoubleClickScheduled Then Return
            If bgview Is Nothing OrElse bgview.GridControl Is Nothing OrElse bgview.GridControl.IsDisposed Then Return

            'The second click can be reported by both the painted header and the
            'temporary BaseEdit. Queue one action after that mouse sequence has
            'finished so a resulting calculation/layout may safely close the editor.
            _DoubleClickScheduled = True
            bgview.GridControl.BeginInvoke(
                New MethodInvoker(
                    Sub()
                        _DoubleClickScheduled = False
                        If _DoubleClickHandler IsNot Nothing Then
                            _DoubleClickHandler.Invoke(_Item, EventArgs.Empty)
                        End If
                    End Sub))

        End Sub
    End Class



End Namespace
'-----------------------------------------------------------------------------
' Vertical Grid equivalent of ColumnInplaceEditorHelper.
'
' A VGrid displays datasource fields as rows, so a repeating-column header from
' the normal XtraGrid maps naturally to the row-header cell (the left/first
' visual column) of the corresponding EditorRow.
'-----------------------------------------------------------------------------
Namespace Abovo

    Public Class VGridRowInplaceEditorHelper
        Public Property PresentationTint As Func(Of Color?)
        Public Property Navigate As Func(Of Keys, Boolean)
        Public ReadOnly Property GridControl As Control
            Get
                Return _VGrid
            End Get
        End Property
        Private _Committing As Boolean
        Private _Closing As Boolean

        Private _Item As RepositoryItem
        Private _Row As DevExpress.XtraVerticalGrid.Rows.EditorRow
        Private _VGrid As DevExpress.XtraVerticalGrid.VGridControl
        Private _EditorHeight As Integer = -1
        Private _LastHeaderBounds As Rectangle = Rectangle.Empty
        Private _EditValue As Object
        Private _ActiveEditor As BaseEdit
        Private _ValueChangedHandler As EventHandler

        Public LinkedComboBoxEdit As AbovoDEHeaderComboBox
        Public LinkedDateEdit As AbovoDEHeaderDateBox
        Public Tag As Object

        Public Sub New(ByVal VGrid As DevExpress.XtraVerticalGrid.VGridControl,
                       ByVal Row As DevExpress.XtraVerticalGrid.Rows.EditorRow,
                       ByVal InplaceEditor As RepositoryItem,
                       ByVal ValueChangedHandler As EventHandler)

            _VGrid = VGrid
            _Row = Row
            _Item = InplaceEditor
            _ValueChangedHandler = ValueChangedHandler

            InplaceEditorFormatting.ApplyStandardDateFormat(_Item)

            _EditorHeight = InplaceEditorFormatting.DesiredEditorHeight(_Item, _VGrid)

            AddHandler _VGrid.CustomDrawRowHeaderCell, AddressOf VGrid_CustomDrawRowHeaderCell
            AddHandler _VGrid.MouseDown, AddressOf VGrid_MouseDown
            AddHandler _VGrid.Layout, AddressOf VGrid_Layout
            AddHandler PresentationScaleManager.ScaleChanged, AddressOf PresentationScaleChanged
            AddHandler GridPresentation.ZoomChanged, AddressOf GridZoomChanged
            AddHandler _VGrid.Disposed, Sub() RemoveHandler GridPresentation.ZoomChanged, AddressOf GridZoomChanged

        End Sub

        Private Sub GridZoomChanged(sender As Object, e As PresentationScaleChangedEventArgs)
            If sender Is _VGrid Then PresentationScaleChanged(sender, e)
        End Sub

        Private Sub PresentationScaleChanged(
            ByVal sender As Object,
            ByVal e As PresentationScaleChangedEventArgs)

            If _VGrid Is Nothing OrElse _VGrid.IsDisposed Then
                RemoveHandler PresentationScaleManager.ScaleChanged, AddressOf PresentationScaleChanged
                Return
            End If

            Dim isGridZoom = sender Is _VGrid
            InplaceEditorFormatting.SetEditorFontSize(_Item, _ActiveEditor, InplaceEditorFormatting.OwnerFontSize(_VGrid))
            _EditorHeight = InplaceEditorFormatting.DesiredEditorHeight(_Item, _VGrid)
            EnsureMinimumHeaderGeometry()
            _LastHeaderBounds = Rectangle.Empty
            If Not isGridZoom Then InplaceEditorFormatting.QueuePresentationLayout(_VGrid)
        End Sub

        Private Sub EnsureMinimumHeaderGeometry()
            'Native tree indents and editor buttons do not shrink with a local
            'font zoom. Keep enough room for their chrome instead of fitting
            'the date down to an unreadable one-point font at 50%.
            Dim minimumHeight = _EditorHeight + 2 * PresentationScaleManager.Scale(2)
            If _Row.Height < minimumHeight Then
                _Row.Height = minimumHeight
                GridPresentation.AcceptRenderedMetric(_VGrid, _Row, "Height", _Row.Height)
            End If
            Dim header = _Row.HeaderInfo
            Dim indent = Math.Max(0, (_Row.Level + 1) * (_VGrid.ViewInfo.RowIndentWidth + 1) - 1)
            If header IsNot Nothing AndAlso Not header.HeaderRect.IsEmpty AndAlso Not header.HeaderCellsRect.IsEmpty Then
                indent = Math.Max(0, header.HeaderRect.Width - header.HeaderCellsRect.Width)
            End If
            Using graphics = _VGrid.CreateGraphics()
                Dim value = If(_ActiveEditor IsNot Nothing AndAlso Not _ActiveEditor.IsDisposed, _ActiveEditor.EditValue, EditValue)
                Dim minimumWidth = indent + 2 * PresentationScaleManager.Scale(4) + InplaceEditorFormatting.DesiredEditorWidth(_Item, _VGrid, graphics, value)
                If _VGrid.RowHeaderWidth < minimumWidth Then
                    _VGrid.RowHeaderWidth = minimumWidth
                    GridPresentation.AcceptRenderedMetric(_VGrid, _VGrid, "RowHeaderWidth", _VGrid.RowHeaderWidth)
                End If
            End Using
        End Sub

        Public Property EditValue As Object
            Get
                Return _EditValue
            End Get
            Set(ByVal value As Object)
                _EditValue = value
                If _ActiveEditor IsNot Nothing AndAlso Not _ActiveEditor.IsDisposed Then
                    If _Committing OrElse Not _ActiveEditor.ContainsFocus Then
                        _ActiveEditor.EditValue = value
                    End If
                End If
                If _VGrid IsNot Nothing AndAlso Not _VGrid.IsDisposed Then
                    _VGrid.InvalidateRow(_Row)
                End If
            End Set
        End Property

        Private Function GetEditorBounds(ByVal HeaderBounds As Rectangle) As Rectangle

            Dim HorizontalPadding As Integer = PresentationScaleManager.Scale(4)
            Dim VerticalPadding As Integer = PresentationScaleManager.Scale(2)
            Dim EditorHeight As Integer = _EditorHeight

            If EditorHeight <= 0 Then EditorHeight = PresentationScaleManager.Scale(22)
            EditorHeight = Math.Min(EditorHeight, Math.Max(1, HeaderBounds.Height - (2 * VerticalPadding)))

            Dim EditorWidth As Integer = Math.Max(1, HeaderBounds.Width - (2 * HorizontalPadding))
            Dim EditorTop As Integer = HeaderBounds.Top + Math.Max(VerticalPadding, (HeaderBounds.Height - EditorHeight) \ 2)

            Return New Rectangle(HeaderBounds.Left + HorizontalPadding,
                                 EditorTop,
                                 EditorWidth,
                                 EditorHeight)

        End Function

        Private Sub VGrid_CustomDrawRowHeaderCell(ByVal sender As Object,
                                                   ByVal e As DevExpress.XtraVerticalGrid.Events.CustomDrawRowHeaderCellEventArgs)

            If e.Row IsNot _Row Then Return
            If e.CellIndex <> 0 Then Return

            _LastHeaderBounds = e.Bounds
            _EditorHeight = InplaceEditorFormatting.DesiredEditorHeight(_Item, _VGrid, e.Graphics)
            InplaceEditorFormatting.FitEditorToBounds(_Item, _ActiveEditor, _VGrid, e.Graphics, GetEditorBounds(e.Bounds))
            If _ActiveEditor IsNot Nothing AndAlso Not _ActiveEditor.IsDisposed AndAlso _ActiveEditor.Parent IsNot Nothing Then
                Dim rectangle = GetEditorBounds(e.Bounds)
                _ActiveEditor.Bounds = New Rectangle(_ActiveEditor.Parent.PointToClient(_VGrid.PointToScreen(rectangle.Location)), rectangle.Size)
            End If

            'Draw the normal row header first, but without the generated field
            'caption.  The repository editor is then painted inside the same
            'header cell, giving the VGrid the same visual language as the
            'XtraGrid in-header editors.
            Dim SavedCaption As String = e.Caption
            e.Caption = String.Empty
            e.DefaultDraw()
            e.Caption = SavedCaption

            Dim original As New AppearanceObject()
            original.Assign(_Item.Appearance)
            Try
                Dim tint As Color? = If(PresentationTint Is Nothing, Nothing, PresentationTint.Invoke())
                If tint.HasValue Then
                    _Item.Appearance.BackColor = tint.Value
                    _Item.Appearance.ForeColor = AbovoBlue
                    _Item.Appearance.Options.UseBackColor = True
                End If
                DrawEditorHelper.DrawEdit(e.Graphics, _Item, GetEditorBounds(e.Bounds), EditValue, useRepositoryAppearance:=True)
            Finally
                _Item.Appearance.Assign(original)
            End Try

            e.Handled = True

        End Sub

        Private Sub VGrid_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)

            If Not CommitEditor() Then Return
            CloseEditor()

            If _LastHeaderBounds.IsEmpty Then Return

            Dim HitInfo As DevExpress.XtraVerticalGrid.VGridHitInfo = _VGrid.CalcHitInfo(e.Location)

            If HitInfo Is Nothing Then Return
            If HitInfo.Row IsNot _Row Then Return
            If HitInfo.HitInfoType <> DevExpress.XtraVerticalGrid.HitInfoTypeEnum.HeaderCell Then Return

            Dim EditorBounds As Rectangle = GetEditorBounds(_LastHeaderBounds)
            If Not EditorBounds.Contains(e.Location) Then Return

            ShowEditor(EditorBounds)
            DXMouseEventArgs.GetMouseArgs(e).Handled = True

        End Sub

        Private Sub ShowEditor(ByVal Bounds As Rectangle, Optional keyboard As Boolean = False)

            If _ActiveEditor IsNot Nothing Then CloseEditor()
            _ActiveEditor = _Item.CreateEditor()
            _ActiveEditor.Properties.LockEvents()
            _ActiveEditor.Properties.Assign(_Item)
            _ActiveEditor.Properties.AutoHeight = False
            _ActiveEditor.EnterMoveNextControl = False

            _ActiveEditor.Properties.Appearance.Options.UseBackColor = True
            _ActiveEditor.Properties.Appearance.BackColor = _Item.Appearance.BackColor
            _ActiveEditor.Properties.Appearance.Options.UseForeColor = True
            _ActiveEditor.Properties.Appearance.ForeColor = _Item.Appearance.ForeColor
            _ActiveEditor.Properties.Appearance.Options.UseFont = True
            _ActiveEditor.Properties.Appearance.Font = _Item.Appearance.Font

            'CreateEditor/Properties.Assign does not reliably copy RepositoryItem.Tag.
            'The EditValueChanged event may be raised with either the BaseEdit or its
            'RepositoryItem as sender, so tag both objects explicitly.
            _ActiveEditor.Tag = Tag
            _ActiveEditor.Properties.Tag = Tag

            'Do not parent the temporary editor directly to VGridControl.
            'The VGrid owns an internal view/control hierarchy and overlaying an
            'arbitrary BaseEdit directly inside it can leave that hierarchy in an
            'invalid state.  Instead, make the editor a sibling of the VGrid and
            'translate the row-header rectangle into the VGrid parent's coordinates.
            Dim EditorParent As Control = _VGrid.Parent
            While TypeOf EditorParent Is TableLayoutPanel OrElse TypeOf EditorParent Is DevExpress.Utils.Layout.TablePanel OrElse TypeOf EditorParent Is DevExpress.XtraLayout.LayoutControl
                EditorParent = EditorParent.Parent
            End While

            If EditorParent Is Nothing Then

                _ActiveEditor.Properties.UnLockEvents()
                _ActiveEditor.Dispose()
                _ActiveEditor = Nothing
                Return

            End If

            Dim ScreenLocation As Point = _VGrid.PointToScreen(Bounds.Location)
            Dim ParentLocation As Point = EditorParent.PointToClient(ScreenLocation)

            _ActiveEditor.Parent = EditorParent
            _ActiveEditor.Location = ParentLocation
            _ActiveEditor.Size = Bounds.Size
            _ActiveEditor.CreateControl()

            'Set the initial value before wiring EditValueChanged.  The repository
            'item is only a template in a VGrid header; opening/creating the editor
            'must never be interpreted as a user workbook edit.
            _ActiveEditor.EditValue = EditValue
            _ActiveEditor.BringToFront()

            AddHandler _ActiveEditor.Leave, AddressOf Editor_Leave
            AddHandler _ActiveEditor.PreviewKeyDown, AddressOf Editor_PreviewKeyDown
            AddHandler _ActiveEditor.KeyDown, AddressOf Editor_KeyDown

            If keyboard Then
                _ActiveEditor.Focus()
            Else
                _ActiveEditor.SendMouse(_ActiveEditor.PointToClient(Control.MousePosition), Control.MouseButtons)
            End If
            _ActiveEditor.Properties.UnLockEvents()

        End Sub

        Private Sub CloseEditor()

            If _ActiveEditor Is Nothing OrElse _Closing OrElse _Committing Then Return
            _Closing = True
            Try
                EditValue = _ActiveEditor.EditValue
                RemoveHandler _ActiveEditor.Leave, AddressOf Editor_Leave
                RemoveHandler _ActiveEditor.PreviewKeyDown, AddressOf Editor_PreviewKeyDown
                RemoveHandler _ActiveEditor.KeyDown, AddressOf Editor_KeyDown
                _ActiveEditor.Dispose()
                _ActiveEditor = Nothing
            Finally
                _Closing = False
            End Try

        End Sub

        Private Sub Editor_Leave(ByVal sender As Object, ByVal e As EventArgs)
            If CommitEditor() Then CloseEditor()
        End Sub

        Private Sub VGrid_Layout(ByVal sender As Object, ByVal e As EventArgs)
            'A calculation can relayout the grid while an editor is committing.
            'Disposing that editor here used to interrupt entry and reset focus.
            If Not GridPresentation.IsApplyingZoom(_VGrid) AndAlso
               _ActiveEditor IsNot Nothing AndAlso Not _ActiveEditor.IsDisposed Then
                _VGrid.InvalidateRow(_Row)
            End If
        End Sub

        Private Function CommitEditor() As Boolean
            If _ActiveEditor Is Nothing Then Return True
            If _Committing OrElse _Closing Then Return False
            If Not _ActiveEditor.DoValidate(PopupCloseMode.Normal) Then Return False
            _Committing = True
            Try
                Dim requested = _ActiveEditor.EditValue
                _EditValue = requested
                If _ValueChangedHandler IsNot Nothing Then _ValueChangedHandler.Invoke(_ActiveEditor, EventArgs.Empty)
                Return InplaceEditorFormatting.SameEditorValue(requested, _ActiveEditor.EditValue)
            Finally
                _Committing = False
            End Try
        End Function

        Friend Function CommitForSave() As Boolean
            If Not CommitEditor() Then Return False
            CloseEditor()
            Return True
        End Function

        Private Sub Editor_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs)
            If {Keys.Tab, Keys.Enter, Keys.Left, Keys.Right, Keys.Up, Keys.Down}.Contains(e.KeyCode) Then e.IsInputKey = True
        End Sub

        Private Sub Editor_KeyDown(sender As Object, e As KeyEventArgs)
            If InplaceEditorFormatting.OpenPopupForShortcut(TryCast(sender, BaseEdit), e) Then Return
            If e.Control OrElse e.Alt Then Return
            Dim popup = TryCast(sender, PopupBaseEdit)
            If popup IsNot Nothing AndAlso popup.IsPopupOpen Then Return
            If e.KeyCode = Keys.Escape Then
                _ActiveEditor.EditValue = _EditValue
                CloseEditor()
                e.Handled = True : e.SuppressKeyPress = True
                Return
            End If
            If Not {Keys.Tab, Keys.Enter, Keys.Left, Keys.Right, Keys.Up, Keys.Down}.Contains(e.KeyCode) Then Return
            e.Handled = True : e.SuppressKeyPress = True
            If Not CommitEditor() Then Return
            CloseEditor()
            If Navigate IsNot Nothing Then Navigate.Invoke(e.KeyData)
        End Sub

        Public Sub ShowEditorFromKeyboard()
            If _ActiveEditor IsNot Nothing AndAlso Not _ActiveEditor.IsDisposed Then
                _ActiveEditor.Focus()
                Return
            End If
            If _VGrid Is Nothing OrElse _VGrid.IsDisposed OrElse Not _Row.Visible Then Return
            _VGrid.MakeRowVisible(_Row)
            _VGrid.Refresh()
            'The row can be outside the outer tab's viewport and not painted yet.
            'Use native layout bounds, never a stale cached rectangle from before a scroll.
            Dim rowInfo = _VGrid.ViewInfo.RowsViewInfo.Cast(Of DevExpress.XtraVerticalGrid.ViewInfo.BaseRowViewInfo)().FirstOrDefault(Function(info) info.Row Is _Row)
            If rowInfo IsNot Nothing AndAlso rowInfo.HeaderInfo IsNot Nothing Then _LastHeaderBounds = rowInfo.HeaderInfo.HeaderCellsRect
            If Not _LastHeaderBounds.IsEmpty Then ShowEditor(GetEditorBounds(_LastHeaderBounds), True)
        End Sub

        Public Sub DetachForDisposal()

            CloseEditor()
            RemoveHandler GridPresentation.ZoomChanged, AddressOf GridZoomChanged

            If _VGrid IsNot Nothing Then
                RemoveHandler _VGrid.CustomDrawRowHeaderCell, AddressOf VGrid_CustomDrawRowHeaderCell
                RemoveHandler _VGrid.MouseDown, AddressOf VGrid_MouseDown
                RemoveHandler _VGrid.Layout, AddressOf VGrid_Layout
            End If
            RemoveHandler PresentationScaleManager.ScaleChanged, AddressOf PresentationScaleChanged
            Navigate = Nothing

            _Row = Nothing
            _Item = Nothing
            _VGrid = Nothing
            _ValueChangedHandler = Nothing
            Tag = Nothing
            LinkedComboBoxEdit = Nothing
            LinkedDateEdit = Nothing

        End Sub

    End Class

End Namespace
