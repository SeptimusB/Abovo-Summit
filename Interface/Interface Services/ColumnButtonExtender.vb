Imports Abovo.DataObject
Imports Abovo.LogDebugDev
Imports Abovo.GeneralFunctions

Imports System.Drawing.Drawing2D
Imports DevExpress.Pdf.Drawing
Imports DevExpress.Pdf.Native.BouncyCastle.Asn1
Imports DevExpress.Utils
Imports DevExpress.Utils.Drawing
Imports DevExpress.Utils.Extensions
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraEditors.Drawing
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.BandedGrid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Drawing
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Grid.Drawing
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports DevExpress.XtraGrid.Views.BandedGrid.ViewInfo
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar
Imports DevExpress.XtraSpreadsheet
Imports DevExpress.XtraSpreadsheet.Model

Namespace Abovo
    Public Class ColumnButtonExtender

        Private view As GridView
        Private ActiveColumn As GridColumn
        Private customButtonPainter As SkinEditorButtonPainter
        Private argsColATopButton As EditorButtonObjectInfoArgs
        Private buttonSize As Size = New Size(28, 28)
        Private MyOwner As DataInterfaceTemplate
        Private BandedMode As Boolean = False
        Private TargetRange As String
        Private ActTok As Integer
        Private ButtonsTTC As ToolTipController
        Private _actionToken As ActionToken
        Private SectionID As Integer

        'Exact rectangle used by the most recent custom header paint.
        'Using the painted rectangle keeps mouse interaction independent of
        'CalcHitInfo/ColumnsInfo quirks in complex banded views.
        Private LastPaintedButtonBounds As Rectangle = Rectangle.Empty

        'Tracks a complete click gesture that started on this button.
        Private MouseDownOnButton As Boolean = False
        Public Sub New(ByVal view As GridView, Opener As Object, Method As String, SetTargetRange As String, Col As GridColumn, SetSectionID As Integer, ActTok As ActionToken)

            SectionID = SetSectionID
            ButtonsTTC = New ToolTipController()
            TargetRange = SetTargetRange
            Me.view = view
            ActiveColumn = Col
            MyOwner = Opener
            _actionToken = ActTok

        End Sub

        Public Sub AddCustomButton()

            CreateButtonPainter()
            CreateButtonInfoArgs()
            SubscribeToEvents()

        End Sub

        Private Sub CreateButtonInfoArgs()

            Dim btn As New EditorButton(ButtonPredefines.Plus)
            btn.ToolTip = "Add Rows"
            btn.IsLeft = True

            argsColATopButton = New EditorButtonObjectInfoArgs(btn, New DevExpress.Utils.AppearanceObject())

        End Sub

        Private Sub CreateButtonPainter()

            customButtonPainter = New SkinEditorButtonPainter(DevExpress.LookAndFeel.UserLookAndFeel.Default.ActiveLookAndFeel)

        End Sub

        Private Sub SubscribeToEvents()

            On Error Resume Next
            RemoveHandler view.CustomDrawColumnHeader, AddressOf DefaultHelpers.DefaultCustomDrawColumnHeader

            AddHandler view.CustomDrawColumnHeader, AddressOf OnCustomDrawColumnHeader

            AddHandler view.MouseDown, AddressOf OnMouseDown
            AddHandler view.MouseUp, AddressOf OnMouseUp
            AddHandler view.MouseMove, AddressOf OnMouseMove

        End Sub

        Private Sub OnMouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)

            Dim WasPressed As Boolean = MouseDownOnButton
            MouseDownOnButton = False

            If Not WasPressed Then Return

            If Not IsButtonRect(e.Location, ActiveColumn) Then

                SetButtonState(ActiveColumn, ObjectState.Normal)
                Return

            End If

            SetButtonState(ActiveColumn, ObjectState.Normal)

            'Run the action only after a complete down/up gesture on the same
            'painted button rectangle.
            MyOwner.RunAction(_actionToken)

            DXMouseEventArgs.GetMouseArgs(e).Handled = True

        End Sub

        Private Sub OnMouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)

            If IsButtonRect(e.Location, ActiveColumn) Then

                If MouseDownOnButton Then
                    SetButtonState(ActiveColumn, ObjectState.Pressed)
                Else
                    SetButtonState(ActiveColumn, ObjectState.Hot)
                End If

            Else

                SetButtonState(ActiveColumn, ObjectState.Normal)

            End If

        End Sub

        Private Sub OnMouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)

            MouseDownOnButton = False

            If e.Button <> MouseButtons.Left Then Return

            If Not IsButtonRect(e.Location, ActiveColumn) Then Return

            MouseDownOnButton = True
            SetButtonState(ActiveColumn, ObjectState.Pressed)

            'Do not consume MouseDown here. The action is dispatched on MouseUp
            'after the completed button gesture.

        End Sub

        Private Sub SetButtonState(ByVal column As GridColumn, ByVal state As ObjectState)

            column.Tag.ButtonObjectState = state
            view.InvalidateColumnHeader(column)

        End Sub

        Private Function IsButtonRect(ByVal point As Point, ByVal column As GridColumn) As Boolean

            'Primary path: use the exact rectangle used to paint the button.
            If Not LastPaintedButtonBounds.IsEmpty Then
                Return LastPaintedButtonBounds.Contains(point)
            End If

            'Fallback only if mouse interaction occurs before the first custom
            'draw following a layout change.
            Dim info As New GraphicsInfo()
            info.AddGraphics(Nothing)

            Try

                Dim viewInfo As GridViewInfo =
                    TryCast(view.GetViewInfo(), GridViewInfo)

                If viewInfo Is Nothing Then Return False

                Dim columnArgs As GridColumnInfoArgs =
                    viewInfo.ColumnsInfo(column)

                If columnArgs Is Nothing Then Return False

                Dim buttonRect As Rectangle =
                    CalcButtonRect(
                        columnArgs,
                        info.Graphics)

                Return buttonRect.Contains(point)

            Finally

                info.ReleaseGraphics()

            End Try

        End Function

        Private Function CalcButtonRect(ByVal columnArgs As GridColumnInfoArgs, ByVal gr As Graphics) As Rectangle

            Dim columnRect As Rectangle = columnArgs.Bounds
            Dim innerElementsWidth As Integer = CalcInnerElementsMinWidth(columnArgs, gr)
            Dim buttonRect As New Rectangle(columnRect.Left + 5, columnRect.Y + (columnRect.Height - buttonSize.Height) - 1, buttonSize.Width, buttonSize.Height)

            Return buttonRect

        End Function

        Private Function CalcInnerElementsMinWidth(ByVal columnArgs As GridColumnInfoArgs, ByVal gr As Graphics) As Integer

            Dim canDrawMode As Boolean = True
            Return columnArgs.InnerElements.CalcMinSize(gr, canDrawMode).Width

        End Function

        Private Sub OnCustomDrawColumnHeader(ByVal sender As Object, ByVal e As ColumnHeaderCustomDrawEventArgs)

            If e.Column Is Nothing Then
                e.Handled = True
                Return
            End If

            If e.Column IsNot ActiveColumn Then GoTo DefaultDraw

            Dim ColTag As DataColumnTag = e.Column.Tag

            WithButtonDrawColumnHeader(e)
            DrawCustomButton(e)

            e.Handled = True

            Return

            Exit Sub

DefaultDraw:

            DefaultDrawColumnHeader(e)

            e.Handled = True

            Return

        End Sub

        Private Sub DrawCustomButton(ByVal e As ColumnHeaderCustomDrawEventArgs)

            SetUpButtonInfoArgs(e)
            customButtonPainter.DrawObject(argsColATopButton)

        End Sub

        Private Sub SetUpButtonInfoArgs(ByVal e As ColumnHeaderCustomDrawEventArgs)

            argsColATopButton.Cache = e.Cache

            LastPaintedButtonBounds =
                CalcButtonRect(
                    e.Info,
                    e.Cache.Graphics)

            argsColATopButton.Bounds =
                LastPaintedButtonBounds

            Dim state As ObjectState = ObjectState.Normal
            If TypeOf e.Column.Tag.ButtonObjectState Is ObjectState Then
                state = DirectCast(e.Column.Tag.ButtonObjectState, ObjectState)
            End If
            argsColATopButton.State = state

        End Sub

        Private Sub DefaultDrawColumnHeader(ByVal e As ColumnHeaderCustomDrawEventArgs)

            Dim bounds As Rectangle = e.Bounds

            If e.Column.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Near Then

                bounds.X += DefaultGridCellPadding

            Else

                bounds.Width -= DefaultGridCellPadding

            End If

            If e.Column Is Nothing Then
                e.Handled = True
                Return
            End If

            bounds.X += buttonSize.Width + 3
            bounds.Width -= buttonSize.Width

            If e.Column.AbsoluteIndex = 0 Then SystemLog("defdrawDraw Index 0 - " & e.Column.AbsoluteIndex)
            e.Appearance.FillRectangle(e.Cache, e.Bounds)
            e.Cache.DrawString(Trim(e.Info.Caption), e.Appearance.GetFont(), e.Appearance.GetForeBrush(e.Cache), e.Bounds, e.Appearance.GetStringFormat())
            e.Handled = True

            Return

        End Sub
        Private Sub WithButtonDrawColumnHeader(ByVal e As ColumnHeaderCustomDrawEventArgs)

            Dim CellPadding As Integer = 28

            ' Fill column headers with the specified colors.

            Dim bounds As Rectangle = e.Bounds

            If e.Column.Tag.WidthSet = False Then

                e.Column.Tag.DefaultColumnWidth = e.Column.Width
                SystemLog("DefaultColumnWidth: " & e.Column.Tag.DefaultColumnWidth)
                e.Column.Tag.ExtendedColumnWidth = e.Column.Width + CellPadding + 15
                SystemLog("ExtendedColumnWidth: " & e.Column.Tag.ExtendedColumnWidth)
                e.Column.Tag.WidthSet = True
                view.InvalidateColumnHeader(e.Column)
                e.Handled = True
                Return

            End If

            e.Column.Width = e.Column.Tag.ExtendedColumnWidth

            e.Appearance.FillRectangle(e.Cache, e.Bounds)

            bounds.X += CellPadding
            bounds.Width -= (2 * CellPadding)

            e.Cache.DrawString(e.Info.Caption, e.Appearance.GetFont(), e.Appearance.GetForeBrush(e.Cache), bounds, e.Appearance.GetStringFormat())

            e.Handled = True

        End Sub

        Private Sub UnsubscribeFromEvents()

            RemoveHandler view.CustomDrawColumnHeader, AddressOf OnCustomDrawColumnHeader
            RemoveHandler view.MouseDown, AddressOf OnMouseDown
            RemoveHandler view.MouseUp, AddressOf OnMouseUp
            RemoveHandler view.MouseMove, AddressOf OnMouseMove

        End Sub

        Public Sub RemoveCustomButton()

            UnsubscribeFromEvents()

            MouseDownOnButton = False
            LastPaintedButtonBounds = Rectangle.Empty

        End Sub

    End Class

    Public Class ColumnFooterButtonExtender

        Private view As GridView
        Private ActiveColumn As GridColumn
        Private customButtonPainter As SkinEditorButtonPainter
        Private args As EditorButtonObjectInfoArgs
        Private buttonSize As Size = New Size(35, 35)
        Private MyOwner As DataInterfaceTemplate
        Private TargetRange As String
        Private ActTok As Integer
        Private _actionToken As ActionToken
        Private SectionID As Integer
        Public Sub New(ByVal view As GridView, Opener As Object, Method As String, SetTargetRange As String, Col As GridColumn, SetSectionID As Integer, ActTok As ActionToken)

            SectionID = SetSectionID

            TargetRange = SetTargetRange
            Me.view = view
            ActiveColumn = Col
            MyOwner = Opener
            _actionToken = ActTok

        End Sub

        Public Sub AddCustomButton()

            CreateButtonPainter()
            CreateButtonInfoArgs()
            SubscribeToEvents()

        End Sub

        Private Sub CreateButtonInfoArgs()

            Dim btn As New EditorButton(ButtonPredefines.Plus)
            btn.ToolTip = "Add Rows"
            btn.IsLeft = True
            args = New EditorButtonObjectInfoArgs(btn, New DevExpress.Utils.AppearanceObject())

        End Sub

        Private Sub CreateButtonPainter()

            customButtonPainter = New SkinEditorButtonPainter(DevExpress.LookAndFeel.UserLookAndFeel.Default.ActiveLookAndFeel)

        End Sub

        Private Sub SubscribeToEvents()

            On Error Resume Next

            AddHandler view.CustomDrawColumnHeader, AddressOf OnCustomDrawColumnHeader

            AddHandler view.MouseDown, AddressOf OnMouseDown
            AddHandler view.MouseUp, AddressOf OnMouseUp
            AddHandler view.MouseMove, AddressOf OnMouseMove

        End Sub

        Private Sub OnMouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
            Dim hitInfo As GridHitInfo = view.CalcHitInfo(e.Location)
            If hitInfo.HitTest <> GridHitTest.Column Then
                Return
            End If
            Dim column As GridColumn = hitInfo.Column
            If IsButtonRect(e.Location, column) Then
                SetButtonState(column, ObjectState.Normal)
                MyOwner.RunAction(_actionToken)
                DXMouseEventArgs.GetMouseArgs(e).Handled = True
            End If
        End Sub

        Private Sub OnMouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
            Dim hitInfo As GridHitInfo = view.CalcHitInfo(e.Location)
            If hitInfo.HitTest <> GridHitTest.Column Then
                Return
            End If
            Dim column As GridColumn = hitInfo.Column
            If IsButtonRect(e.Location, column) Then
                SetButtonState(column, ObjectState.Hot)
            Else
                SetButtonState(column, ObjectState.Normal)
            End If
        End Sub

        Private Sub OnMouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
            Dim hitInfo As GridHitInfo = view.CalcHitInfo(e.Location)
            If hitInfo.HitTest <> GridHitTest.Column Then
                Return
            End If
            Dim column As GridColumn = hitInfo.Column
            If IsButtonRect(e.Location, column) Then
                SetButtonState(column, ObjectState.Pressed)
                DXMouseEventArgs.GetMouseArgs(e).Handled = True
            End If
        End Sub

        Private Sub SetButtonState(ByVal column As GridColumn, ByVal state As ObjectState)
            column.Tag.ButtonObjectState = state
            view.InvalidateColumnHeader(column)
        End Sub

        Private Function IsButtonRect(ByVal point As Point, ByVal column As GridColumn) As Boolean
            Dim info As New GraphicsInfo()
            info.AddGraphics(Nothing)
            Dim viewInfo As GridViewInfo = TryCast(view.GetViewInfo(), GridViewInfo)
            Dim columnArgs As GridColumnInfoArgs = viewInfo.ColumnsInfo(column)
            Dim buttonRect As Rectangle = CalcButtonRect(columnArgs, info.Graphics)
            info.ReleaseGraphics()
            Return buttonRect.Contains(point)
        End Function

        Private Function CalcButtonRect(ByVal columnArgs As GridColumnInfoArgs, ByVal gr As Graphics) As Rectangle
            Dim columnRect As Rectangle = columnArgs.Bounds
            Dim innerElementsWidth As Integer = CalcInnerElementsMinWidth(columnArgs, gr)
            'Dim buttonRect As New Rectangle(columnRect.Left + (columnRect.Width / 2) - (buttonSize.Width / 2), columnRect.Y - columnRect.Height \ 2 + buttonSize.Height \ 2, buttonSize.Width, buttonSize.Height)
            Dim buttonRect As New Rectangle(columnRect.Right - innerElementsWidth - buttonSize.Width - 2, columnRect.Y + columnRect.Height \ 2 - buttonSize.Height \ 2, buttonSize.Width, buttonSize.Height)
            Return buttonRect
        End Function

        Private Function CalcInnerElementsMinWidth(ByVal columnArgs As GridColumnInfoArgs, ByVal gr As Graphics) As Integer
            Dim canDrawMode As Boolean = True
            Return columnArgs.InnerElements.CalcMinSize(gr, canDrawMode).Width
        End Function

        Private Sub OnCustomDrawColumnHeader(ByVal sender As Object, ByVal e As ColumnHeaderCustomDrawEventArgs)

            If e.Column Is Nothing Then
                e.Handled = True
                Return
            End If

            If e.Column IsNot ActiveColumn Then GoTo DefaultDraw


            Dim ColTag As DataColumnTag = e.Column.Tag



            SystemLog("WBDraw Index - " & e.Column.AbsoluteIndex)
            'SystemLog(e.Column.Caption)

            WithButtonDrawColumnHeader(e)
            DrawCustomButton(e)

            e.Handled = True

            Return
            Exit Sub

DefaultDraw:

            DefaultDrawColumnHeader(e)

            e.Handled = True

            Return

        End Sub

        Private Sub DrawCustomButton(ByVal e As ColumnHeaderCustomDrawEventArgs)
            SetUpButtonInfoArgs(e)
            customButtonPainter.DrawObject(args)
        End Sub

        Private Sub SetUpButtonInfoArgs(ByVal e As ColumnHeaderCustomDrawEventArgs)
            args.Cache = e.Cache
            args.Bounds = CalcButtonRect(e.Info, e.Cache.Graphics)
            Dim state As ObjectState = ObjectState.Normal
            If TypeOf e.Column.Tag.ButtonObjectState Is ObjectState Then
                state = DirectCast(e.Column.Tag.ButtonObjectState, ObjectState)
            End If
            args.State = state
        End Sub

        Private Sub DefaultDrawColumnHeader(ByVal e As ColumnHeaderCustomDrawEventArgs)



            Dim bounds As Rectangle = e.Bounds

            If e.Column.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Near Then

                bounds.X += DefaultGridCellPadding

            Else

                bounds.Width -= DefaultGridCellPadding

            End If
            'If e.Column.Tag.WidthSet = False Then

            '    e.Column.Tag.DefaultColumnWidth = e.Column.Width
            '    SystemLog("DefaultColumnWidth: " & e.Column.Tag.DefaultColumnWidth)
            '    e.Column.Tag.ExtendedColumnWidth = e.Column.Width + 15
            '    SystemLog("ExtendedColumnWidth: " & e.Column.Tag.ExtendedColumnWidth)
            '    e.Column.Tag.WidthSet = True
            '    
            '    e.Handled = False
            '    Return

            'End If
            'Dim brush As Brush = New Brush()
            If e.Column Is Nothing Then
                e.Handled = True
                Return
            End If

            If e.Column.AbsoluteIndex = 0 Then SystemLog("defdrawDraw Index 0 - " & e.Column.AbsoluteIndex)
            e.Appearance.FillRectangle(e.Cache, e.Bounds)
            e.Cache.DrawString(Trim(e.Info.Caption), e.Appearance.GetFont(), e.Appearance.GetForeBrush(e.Cache), e.Bounds, e.Appearance.GetStringFormat())
            e.Handled = True
            Return


        End Sub
        Private Sub WithButtonDrawColumnHeader(ByVal e As ColumnHeaderCustomDrawEventArgs)

            Dim CellPadding As Integer = 28

            ' Fill column headers with the specified colors.

            Dim bounds As Rectangle = e.Bounds

            If e.Column.Tag.WidthSet = False Then

                e.Column.Tag.DefaultColumnWidth = e.Column.Width
                e.Column.Tag.ExtendedColumnWidth = e.Column.Width + CellPadding + 15
                e.Column.Tag.WidthSet = True
                view.InvalidateColumnHeader(e.Column)

                e.Handled = True

                Return

            End If

            e.Column.Width = e.Column.Tag.ExtendedColumnWidth

            e.Appearance.FillRectangle(e.Cache, e.Bounds)

            bounds.X += CellPadding
            bounds.Width -= (2 * CellPadding)

            e.Cache.DrawString(e.Info.Caption, e.Appearance.GetFont(), e.Appearance.GetForeBrush(e.Cache), bounds, e.Appearance.GetStringFormat())

            e.Handled = True

        End Sub

        Private Sub UnsubscribeFromEvents()

            RemoveHandler view.CustomDrawColumnHeader, AddressOf OnCustomDrawColumnHeader
            RemoveHandler view.MouseDown, AddressOf OnMouseDown
            RemoveHandler view.MouseUp, AddressOf OnMouseUp
            RemoveHandler view.MouseMove, AddressOf OnMouseMove

        End Sub

        Public Sub RemoveCustomButton()

            UnsubscribeFromEvents()

        End Sub

    End Class

    Public Class BandButtonExtender

        Private view As BandedGridView
        Private customButtonPainter As SkinEditorButtonPainter
        Private args As EditorButtonObjectInfoArgs
        Private buttonSize As Size = New Size(28, 28)
        Private SectionID As Integer
        Private MyOwner As DataInterfaceTemplate
        Private MyDesc As String
        Private _ActionToken As ActionToken
        Public Sub New(ByVal view As BandedGridView, Opener As DataInterfaceTemplate, DefaultAction As String, SetSectionID As Integer, SetTitle As String, ActTok As ActionToken)

            MyDesc = SetTitle
            Me.view = view
            SectionID = SetSectionID
            MyOwner = Opener
            _ActionToken = ActTok

        End Sub

        Public Sub AddCustomButton()

            CreateButtonPainter()
            CreateButtonInfoArgs()
            SubscribeToEvents()

        End Sub

        Private Sub CreateButtonInfoArgs()

            Dim btn As New EditorButton(ButtonPredefines.Plus)
            btn.ToolTip = "Add Columns"
            btn.ToolTipAnchor = ToolTipAnchor.Cursor
            btn.IsLeft = True
            args = New EditorButtonObjectInfoArgs(btn, New DevExpress.Utils.AppearanceObject())

        End Sub

        Private Sub CreateButtonPainter()

            customButtonPainter = New SkinEditorButtonPainter(DevExpress.LookAndFeel.UserLookAndFeel.Default.ActiveLookAndFeel)

        End Sub

        Private Sub SubscribeToEvents()

            AddHandler view.CustomDrawBandHeader, AddressOf OnCustomDrawBandHeader
            AddHandler view.MouseDown, AddressOf OnMouseDown
            AddHandler view.MouseUp, AddressOf OnMouseUp
            AddHandler view.MouseMove, AddressOf OnMouseMove
            AddHandler view.CustomDrawColumnHeader, AddressOf Abovo.DefaultHelpers.DefaultCustomDrawColumnHeader

        End Sub

        Private Sub OnMouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)

            Dim hitInfo As BandedGridHitInfo = view.CalcHitInfo(e.Location)

            If hitInfo.HitTest <> BandedGridHitTest.Band Then

                Return

            End If

            Dim band As GridBand = hitInfo.Band
            Dim BandTag As BandTag = band.Tag

            If IsButtonRect(e.Location, band) Then

                SetButtonState(band, ObjectState.Normal)
                MyOwner.RunAction(BandTag.ActionToken)
                DXMouseEventArgs.GetMouseArgs(e).Handled = True

            End If

        End Sub

        Private Sub OnMouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)

            Dim hitInfo As BandedGridHitInfo = view.CalcHitInfo(e.Location)

            If hitInfo.HitTest <> BandedGridHitTest.Band Then

                Return

            End If

            Dim band As GridBand = hitInfo.Band

            If IsButtonRect(e.Location, band) Then

                SetButtonState(band, ObjectState.Hot)

            Else

                SetButtonState(band, ObjectState.Normal)

            End If

        End Sub

        Private Sub OnMouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)

            Dim hitInfo As BandedGridHitInfo = view.CalcHitInfo(e.Location)

            If hitInfo.HitTest <> BandedGridHitTest.Band Then

                Return

            End If

            Dim Band As GridBand = hitInfo.Band

            If IsButtonRect(e.Location, Band) Then

                SetButtonState(Band, ObjectState.Pressed)
                DXMouseEventArgs.GetMouseArgs(e).Handled = True

            End If

        End Sub

        Private Sub SetButtonState(ByVal band As GridBand, ByVal state As ObjectState)

            band.Tag.ButtonObjectState = state
            view.InvalidateBandHeader(band)

        End Sub

        Private Function IsButtonRect(ByVal point As Point, ByVal band As GridBand) As Boolean

            Dim info As New GraphicsInfo()
            info.AddGraphics(Nothing)
            Dim viewInfo As BandedGridViewInfo = TryCast(view.GetViewInfo(), BandedGridViewInfo)
            Dim bandArgs As GridBandInfoArgs = viewInfo.BandsInfo(band)
            Dim buttonRect As Rectangle = CalcButtonRect(bandArgs, info.Graphics)
            info.ReleaseGraphics()
            Return buttonRect.Contains(point)

        End Function

        Private Function CalcButtonRect(ByVal bandArgs As GridBandInfoArgs, ByVal gr As Graphics) As Rectangle

            Dim bandRect As Rectangle = bandArgs.Bounds
            Dim innerElementsWidth As Integer = CalcInnerElementsMinWidth(bandArgs, gr)
            'Dim buttonRect As New Rectangle(columnRect.Left + (columnRect.Width / 2) - (buttonSize.Width / 2), columnRect.Y - columnRect.Height \ 2 + buttonSize.Height \ 2, buttonSize.Width, buttonSize.Height)
            Dim buttonRect As New Rectangle(bandRect.Right - innerElementsWidth - buttonSize.Width - 2, bandRect.Y + bandRect.Height \ 2 - buttonSize.Height \ 2, buttonSize.Width, buttonSize.Height)
            Return buttonRect

        End Function

        Private Function CalcInnerElementsMinWidth(ByVal bandArgs As GridBandInfoArgs, ByVal gr As Graphics) As Integer

            Dim canDrawMode As Boolean = True
            Return bandArgs.InnerElements.CalcMinSize(gr, canDrawMode).Width

        End Function

        Private Sub OnCustomDrawBandHeader(ByVal sender As Object, ByVal e As BandHeaderCustomDrawEventArgs)

            If e.Band Is Nothing Then

                e.Handled = True

                Return

            End If

            Dim BandTag As BandTag = e.Band.Tag

            If Not BandTag.HasActions And Not BandTag.DoBorder Then

                DefaultDrawBand(e)
                e.Handled = True
                Return

            End If

            If BandTag.HasActions Then DrawCustomButton(e)

            WithButtonDrawBandHeader(e)

            e.Handled = True

            Return

        End Sub

        Private Sub DrawCustomButton(ByVal e As BandHeaderCustomDrawEventArgs)

            SetUpButtonInfoArgs(e)

            customButtonPainter.DrawObject(args)

        End Sub

        Private Sub SetUpButtonInfoArgs(ByVal e As BandHeaderCustomDrawEventArgs)

            args.Cache = e.Cache
            args.Bounds = CalcButtonRect(e.Info, e.Cache.Graphics)
            Dim state As ObjectState = ObjectState.Normal

            If TypeOf e.Band.Tag.ButtonObjectState Is ObjectState Then

                state = DirectCast(e.Band.Tag.ButtonObjectState, ObjectState)

            End If

            args.State = state

        End Sub
        Private Sub SetUpBandButtonInfoArgs(ByVal e As BandHeaderCustomDrawEventArgs)

            args.Cache = e.Cache
            args.Bounds = CalcButtonRect(e.Info, e.Cache.Graphics)
            Dim state As ObjectState = ObjectState.Normal

            If TypeOf e.Band.Tag.ButtonObjectState Is ObjectState Then

                state = DirectCast(e.Band.Tag.ButtonObjectState, ObjectState)

            End If

            args.State = state

        End Sub
        Private Sub OnCustomDrawColumnHeader(ByVal sender As Object, ByVal e As ColumnHeaderCustomDrawEventArgs)

            If e.Column Is Nothing Then

                e.Handled = True
                Return

            End If

            Dim bounds As Rectangle = e.Bounds

            If e.Column.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Near Then

                bounds.X += DefaultGridCellPadding

            Else

                bounds.Width -= DefaultGridCellPadding

            End If

            Dim OffCh As Integer = 0

            If e.Column.VisibleIndex = 0 Then OffCh += DefaultGridCellPadding

            e.Cache.DrawString(e.Info.Caption, e.Appearance.GetFont(), e.Appearance.GetForeBrush(e.Cache), bounds, e.Appearance.GetStringFormat())

            Dim pen As Pen = New Pen(AbovoBlue, 3)

            e.Cache.DrawLine(pen, New Point(e.Bounds.X + OffCh, e.Bounds.Bottom), New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Bottom))

            e.Handled = True

        End Sub
        Private Sub DefaultDrawBand(ByVal e As BandHeaderCustomDrawEventArgs)

            Dim bounds As Rectangle = e.Bounds

            'If e.Column.Tag.WidthSet = False Then

            '    e.Column.Tag.DefaultColumnWidth = e.Column.Width
            '    SystemLog("DefaultColumnWidth: " & e.Column.Tag.DefaultColumnWidth)
            '    e.Column.Tag.ExtendedColumnWidth = e.Column.Width + 15
            '    SystemLog("ExtendedColumnWidth: " & e.Column.Tag.ExtendedColumnWidth)
            '    e.Column.Tag.WidthSet = True
            '    
            '    e.Handled = False
            '    Return

            'End If
            'Dim brush As Brush = New Brush()
            'e.Appearance.BackColor = Color.WhiteSmoke
            e.Cache.DrawString(e.Info.Caption, e.Appearance.GetFont(), e.Appearance.GetForeBrush(e.Cache), e.Bounds, e.Appearance.GetStringFormat())

            'e.Painter.DrawObject(e.Info)

        End Sub
        Private Sub WithButtonDrawBandHeader(ByVal e As BandHeaderCustomDrawEventArgs)

            Dim CellPadding As Integer = 28
            Dim bounds As Rectangle = e.Bounds
            bounds.Y -= 5
            Dim BandTag As BandTag = e.Band.Tag
            Dim pen As Pen = New Pen(BandTag.HighLightColour, 4)

            e.Cache.DrawString(e.Info.Caption, e.Appearance.GetFont(), e.Appearance.GetForeBrush(e.Cache), bounds, e.Appearance.GetStringFormat())

            e.Cache.DrawLine(pen, New Point(e.Bounds.X + CellPadding, e.Bounds.Bottom), New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Bottom))
            'e.Cache.DrawLine(pen, New Point(e.Bounds.X + CellPadding, e.Bounds.Top), New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Top))
            pen.Dispose()

            e.Handled = True

            Return

        End Sub

        Private Sub UnsubscribeFromEvents()

            RemoveHandler view.CustomDrawBandHeader, AddressOf OnCustomDrawBandHeader
            RemoveHandler view.MouseDown, AddressOf OnMouseDown
            RemoveHandler view.MouseUp, AddressOf OnMouseUp
            RemoveHandler view.MouseMove, AddressOf OnMouseMove

        End Sub

        Public Sub RemoveCustomButton()

            UnsubscribeFromEvents()

        End Sub

    End Class

End Namespace
