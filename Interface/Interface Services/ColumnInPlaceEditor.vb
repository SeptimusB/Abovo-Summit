
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

        Public Sub ApplyStandardDateFormat(ByVal Item As RepositoryItem)

            Dim DateItem As RepositoryItemDateEdit = TryCast(Item, RepositoryItemDateEdit)
            If DateItem Is Nothing Then Exit Sub

            'Use one date format for both the custom-painted editor and the
            'temporary live editor.  In .NET custom date formats, lower-case
            'dd/yyyy are the day/year tokens.
            With DateItem
                .DisplayFormat.FormatType = FormatType.DateTime
                .DisplayFormat.FormatString = "dd-MMM-yyyy"
                .EditFormat.FormatType = FormatType.DateTime
                .EditFormat.FormatString = "dd-MMM-yyyy"
                .Mask.EditMask = "dd-MMM-yyyy"
                .UseMaskAsDisplayFormat = True
            End With

        End Sub

    End Module

    Public Class ColumnInplaceEditorHelper

        Private _Item As RepositoryItem
        Private _Column As BandedGridColumn
        Private bgview As GridView
        Private _EditorHeight As Integer = -1

        'Exact editor rectangle used by the most recent custom header paint.
        'Mouse interaction must use the same rectangle rather than trying to
        'reconstruct it later from BandedGridViewInfo.ColumnsInfo.
        Private _LastPaintedEditorBounds As Rectangle = Rectangle.Empty

        Public LinkedComboBoxEdit As AbovoDEHeaderComboBox
        Public LinkedDateEdit As AbovoDEHeaderDateBox
        Public Tag As Object

        Public Sub New(ByVal column As BandedGridColumn, ByVal inplaceEditor As RepositoryItem)
            _Column = column
            _Item = inplaceEditor
            bgview = TryCast(column.View, BandedGridView)

            InplaceEditorFormatting.ApplyStandardDateFormat(_Item)

            _EditorHeight = DrawEditorHelper.GetNaturalEditorHeight(_Item, Nothing)
            '_ActiveEditor = New BaseEdit
            '_ActiveEditor.ForeColor = Color.Red
            AddHandler bgview.CustomDrawColumnHeader, AddressOf view_CustomDrawColumnHeader
            AddHandler bgview.MouseDown, AddressOf view_MouseDown
            AddHandler bgview.Layout, AddressOf view_Layout
        End Sub

        Private Sub view_Layout(ByVal sender As Object, ByVal e As EventArgs)
            CloseEditor()

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

            Debug.Print("Inplace editor - custom draw - " & e.Column.AbsoluteIndex)
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
                _LastPaintedEditorBounds =
                    DrawEditorHelper.GetEditorBounds(
                        e.Bounds,
                        GetRightIndent(),
                        _EditorHeight)

                DrawEditorHelper.DrawColumnInplaceEditor(e, _Item, EditValue, GetRightIndent(), _EditorHeight)
                Debug.Print("Inplace editor - finished custom draw - " & e.Column.AbsoluteIndex)
                e.Handled = True

            End If
        End Sub

        Public Function GetRightIndent() As Integer
            If _Column.OptionsColumn.AllowSort <> DevExpress.Utils.DefaultBoolean.False OrElse _Column.OptionsFilter.AllowFilter Then
                Return 25
            Else
                Return 0
            End If
        End Function
        Private Sub view_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
            CloseEditor()
            Dim editorBounds As Rectangle
            If ClickInEditor(e, editorBounds) Then
                ShowEditor(editorBounds)
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

        Private Sub ShowEditor(ByVal bounds As Rectangle)

            ActiveEditor = _Item.CreateEditor()
            ActiveEditor.Properties.LockEvents()

            'Assign the repository settings first, then force this temporary editor
            'to use exactly the same fixed rectangle as the custom-painted version.
            ActiveEditor.Properties.Assign(_Item)
            ActiveEditor.Properties.AutoHeight = False

            ActiveEditor.Properties.Appearance.Options.UseBackColor = True
            ActiveEditor.Properties.Appearance.BackColor = _Item.Appearance.BackColor
            ActiveEditor.Properties.Appearance.Options.UseForeColor = True
            ActiveEditor.Properties.Appearance.ForeColor = _Item.Appearance.ForeColor

            ActiveEditor.Tag = Tag
            ActiveEditor.Parent = bgview.GridControl
            ActiveEditor.Location = bounds.Location
            ActiveEditor.Size = bounds.Size
            ActiveEditor.CreateControl()
            ActiveEditor.EditValue = EditValue

            AddHandler ActiveEditor.Leave, AddressOf editor_Leave

            ActiveEditor.SendMouse(ActiveEditor.PointToClient(Control.MousePosition), Control.MouseButtons)
            ActiveEditor.Properties.UnLockEvents()

        End Sub

        Private Sub CloseEditor()
            If ActiveEditor IsNot Nothing Then
                EditValue = ActiveEditor.EditValue
                RemoveHandler ActiveEditor.Leave, AddressOf editor_Leave
                ActiveEditor.Dispose()
                ActiveEditor = Nothing
            End If
        End Sub
        Private Sub editor_Leave(ByVal sender As Object, ByVal e As EventArgs)
            CloseEditor()
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

            _EditorHeight = DrawEditorHelper.GetNaturalEditorHeight(_Item, Nothing)

            AddHandler _VGrid.CustomDrawRowHeaderCell, AddressOf VGrid_CustomDrawRowHeaderCell
            AddHandler _VGrid.MouseDown, AddressOf VGrid_MouseDown
            AddHandler _VGrid.Layout, AddressOf VGrid_Layout

        End Sub

        Public Property EditValue As Object
            Get
                Return _EditValue
            End Get
            Set(ByVal value As Object)
                _EditValue = value
                If _VGrid IsNot Nothing AndAlso Not _VGrid.IsDisposed Then
                    _VGrid.InvalidateRow(_Row)
                End If
            End Set
        End Property

        Private Function GetEditorBounds(ByVal HeaderBounds As Rectangle) As Rectangle

            Dim HorizontalPadding As Integer = 4
            Dim VerticalPadding As Integer = 2
            Dim EditorHeight As Integer = _EditorHeight

            If EditorHeight <= 0 Then EditorHeight = 22
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

            'Draw the normal row header first, but without the generated field
            'caption.  The repository editor is then painted inside the same
            'header cell, giving the VGrid the same visual language as the
            'XtraGrid in-header editors.
            Dim SavedCaption As String = e.Caption
            e.Caption = String.Empty
            e.DefaultDraw()
            e.Caption = SavedCaption

            DrawEditorHelper.DrawEdit(e.Graphics,
                                      _Item,
                                      GetEditorBounds(e.Bounds),
                                      EditValue)

            e.Handled = True

        End Sub

        Private Sub VGrid_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)

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

        Private Sub ShowEditor(ByVal Bounds As Rectangle)

            _ActiveEditor = _Item.CreateEditor()
            _ActiveEditor.Properties.LockEvents()
            _ActiveEditor.Properties.Assign(_Item)
            _ActiveEditor.Properties.AutoHeight = False

            _ActiveEditor.Properties.Appearance.Options.UseBackColor = True
            _ActiveEditor.Properties.Appearance.BackColor = _Item.Appearance.BackColor
            _ActiveEditor.Properties.Appearance.Options.UseForeColor = True
            _ActiveEditor.Properties.Appearance.ForeColor = _Item.Appearance.ForeColor

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

            _ActiveEditor.SendMouse(_ActiveEditor.PointToClient(Control.MousePosition), Control.MouseButtons)
            _ActiveEditor.Properties.UnLockEvents()

            If _ValueChangedHandler IsNot Nothing Then
                AddHandler _ActiveEditor.EditValueChanged, _ValueChangedHandler
            End If

        End Sub

        Private Sub CloseEditor()

            If _ActiveEditor Is Nothing Then Return

            EditValue = _ActiveEditor.EditValue

            If _ValueChangedHandler IsNot Nothing Then
                RemoveHandler _ActiveEditor.EditValueChanged, _ValueChangedHandler
            End If

            RemoveHandler _ActiveEditor.Leave, AddressOf Editor_Leave
            _ActiveEditor.Dispose()
            _ActiveEditor = Nothing

        End Sub

        Private Sub Editor_Leave(ByVal sender As Object, ByVal e As EventArgs)
            CloseEditor()
        End Sub

        Private Sub VGrid_Layout(ByVal sender As Object, ByVal e As EventArgs)
            CloseEditor()
        End Sub

        Public Sub DetachForDisposal()

            CloseEditor()

            If _VGrid IsNot Nothing Then
                RemoveHandler _VGrid.CustomDrawRowHeaderCell, AddressOf VGrid_CustomDrawRowHeaderCell
                RemoveHandler _VGrid.MouseDown, AddressOf VGrid_MouseDown
                RemoveHandler _VGrid.Layout, AddressOf VGrid_Layout
            End If

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
