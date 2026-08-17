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
Imports DevExpress.CodeParser

Namespace Abovo
    Public Class CombinedColumnBandButton_HeaderFooterDrawer

        Private view As GridView
        Private bgview As BandedGridView
        Private ViewTag As GridViewTag

        Private MustInvalidateFooter As Boolean = False

        Private ActiveColumns() As ButtonedColumn
        Private ActiveColumnCount As Integer = -1

        Private ActiveBands() As ButtonedBand
        Private ActiveBandCount As Integer = -1

        Private ActiveButtons() As ActionButton
        Private ActiveButtonCount As Integer = -1
        Private ActiveBandButtons() As ActionButton
        Private ActiveBandButtonCount As Integer = -1
        Private ActiveColAddRowButtons() As ActionButton
        Private ActiveColAddRowButtonCount As Integer = -1

        Private BandsActive As Boolean = False
        Private CurrFH As Single
        Private ColsActive As Boolean = False

        Private customButtonPainter As SkinEditorButtonPainter
        Private argsColATopButton As EditorButtonObjectInfoArgs
        Private buttonSize As Size = New Size(28, 28)
        Private MyOwner As DataInterfaceTemplate
        Private BandedMode As Boolean = False
        Private TargetRange As String
        Private ButtonsTTC As ToolTipController
        Private SectionID As Integer

        Public Sub New(ByVal SendingView As Object, Opener As DataInterfaceTemplate, SetSectionID As Integer)

            SectionID = SetSectionID
            ButtonsTTC = New ToolTipController()
            MyOwner = Opener

            If SendingView.GetType Is GetType(GridView) Then

                Me.view = DirectCast(SendingView, GridView)

                BandedMode = False

            ElseIf SendingView.GetType Is GetType(BandedGridView) Then

                Me.bgview = DirectCast(SendingView, BandedGridView)

                BandedMode = True

            End If

            SubscribeToEvents()

            customButtonPainter = New SkinEditorButtonPainter(DevExpress.LookAndFeel.UserLookAndFeel.Default.ActiveLookAndFeel)

        End Sub
        Class ActionButton

            Public Button As EditorButton
            Public ActionToken As ActionToken
            Public ButtonType As String
            Public Column As GridColumn
            Public Band As GridBand
            Public BandMode As Boolean = False
            Public ToolTip As String
            Public InfoArgs As EditorButtonObjectInfoArgs

            'Exact rectangles used by the most recent custom paint pass.
            Public HeaderRect As Rectangle = Rectangle.Empty
            Public FooterRect As Rectangle = Rectangle.Empty
            Public BandRect As Rectangle = Rectangle.Empty

        End Class

        Class ButtonedColumn

            Public ActionButton As ActionButton
            Public Column As GridColumn

        End Class
        Class ButtonedBand

            Public ActionButton As ActionButton
            Public Band As GridBand

        End Class
        Public Sub AddCustomAddRowsButton(Col As GridColumn, ActionToken As ActionToken)

            ColsActive = True

            ActiveButtonCount += 1
            ReDim Preserve ActiveButtons(ActiveButtonCount)
            ActiveButtons(ActiveButtonCount) = New ActionButton() With {.Button = New EditorButton(ButtonPredefines.Plus), .ActionToken = ActionToken, .Column = Col, .ToolTip = "Add Lines"}
            ActiveButtons(ActiveButtonCount).Button.IsLeft = True

            ActiveColAddRowButtonCount += 1
            ReDim Preserve ActiveColAddRowButtons(ActiveColAddRowButtonCount)
            ActiveColAddRowButtons(ActiveColAddRowButtonCount) = ActiveButtons(ActiveButtonCount)

            ActiveColumnCount += 1
            ReDim Preserve ActiveColumns(ActiveColumnCount)
            ActiveColumns(ActiveColumnCount) = New ButtonedColumn() With {.ActionButton = ActiveButtons(ActiveButtonCount), .Column = Col}

            CreateButtonInfoArgs(ActiveButtons(ActiveButtonCount))

        End Sub
        Public Sub AddCustomBandAddColumsButton(Band As GridBand, ActionToken As ActionToken)

            BandsActive = True
            ActiveButtonCount += 1
            ReDim Preserve ActiveButtons(ActiveButtonCount)
            ActiveButtons(ActiveButtonCount) = New ActionButton() With {.Button = New EditorButton(ButtonPredefines.Plus), .ActionToken = ActionToken, .Band = Band, .ToolTip = "Add Lines", .BandMode = True}
            ActiveButtons(ActiveButtonCount).Button.IsLeft = True

            ActiveBandButtonCount += 1
            ReDim Preserve ActiveBandButtons(ActiveBandButtonCount)
            ActiveBandButtons(ActiveBandButtonCount) = ActiveButtons(ActiveButtonCount)

            ActiveBandCount += 1
            ReDim Preserve ActiveBands(ActiveBandCount)
            ActiveBands(ActiveBandCount) = New ButtonedBand() With {.ActionButton = ActiveBandButtons(ActiveBandButtonCount), .Band = Band}

            CreateButtonInfoArgs(ActiveBandButtons(ActiveBandButtonCount))

        End Sub

        Private Sub CreateButtonInfoArgs(ActionButton As ActionButton)

            ActionButton.InfoArgs = New EditorButtonObjectInfoArgs(ActionButton.Button, New DevExpress.Utils.AppearanceObject())

        End Sub

        Private Sub SubscribeToEvents()

            On Error Resume Next

            If BandedMode Then

                RemoveHandler bgview.CustomDrawColumnHeader, AddressOf DefaultHelpers.DefaultCustomDrawColumnHeader
                AddHandler bgview.CustomDrawColumnHeader, AddressOf OnCustomDrawColumnHeader
                AddHandler bgview.CustomDrawBandHeader, AddressOf OnCustomDrawBandHeader
                AddHandler bgview.MouseDown, AddressOf OnMouseDown
                AddHandler bgview.MouseUp, AddressOf OnMouseUp
                AddHandler bgview.MouseMove, AddressOf OnMouseMove
                AddHandler bgview.CustomDrawFooterCell, AddressOf GVCustomDrawFooterCell
                AddHandler bgview.CustomDrawFooter, AddressOf GVCustomDrawFooter

            Else

                RemoveHandler view.CustomDrawColumnHeader, AddressOf DefaultHelpers.DefaultCustomDrawColumnHeader
                AddHandler view.CustomDrawColumnHeader, AddressOf OnCustomDrawColumnHeader
                AddHandler view.CustomDrawFooterCell, AddressOf GVCustomDrawFooterCell
                AddHandler view.CustomDrawFooter, AddressOf GVCustomDrawFooter

                AddHandler view.MouseDown, AddressOf OnMouseDown
                AddHandler view.MouseUp, AddressOf OnMouseUp
                AddHandler view.MouseMove, AddressOf OnMouseMove

            End If

        End Sub
        Private Function IsButtonRectBand(ByVal point As Point, ByVal band As GridBand) As ActionButton

            If ActiveBandButtons Is Nothing Then Return Nothing
            If ActiveBandButtons.Length = 0 Then Return Nothing

            For Each BandBut In ActiveBandButtons

                If BandBut.Band Is band Then

                    Dim info As New GraphicsInfo()
                    info.AddGraphics(Nothing)
                    Dim viewInfo As BandedGridViewInfo = TryCast(bgview.GetViewInfo(), BandedGridViewInfo)
                    Dim bandArgs As GridBandInfoArgs = viewInfo.BandsInfo(band)
                    Dim buttonRect As Rectangle = CalcButtonRectBand(bandArgs, info.Graphics)
                    If buttonRect.Contains(point) Then
                        info.ReleaseGraphics()
                        Return BandBut
                    End If
                    info.ReleaseGraphics()

                End If

            Next

            Return Nothing

        End Function
        Private Function IsButtonRectBandGCol(ByVal point As Point, ByVal column As BandedGridColumn) As ActionButton

            If ActiveColAddRowButtons Is Nothing Then Return Nothing
            If ActiveColAddRowButtons.Length = 0 Then Return Nothing

            For Each AddRowBut In ActiveColAddRowButtons

                If AddRowBut.Column Is column Then

                    Dim info As New GraphicsInfo()
                    info.AddGraphics(Nothing)
                    Dim viewInfo As BandedGridViewInfo = TryCast(bgview.GetViewInfo(), BandedGridViewInfo)
                    Dim columnArgs As GridColumnInfoArgs = viewInfo.ColumnsInfo(column)
                    Dim buttonRect As Rectangle = CalcButtonRect(columnArgs, info.Graphics)

                    If buttonRect.Contains(point) Then
                        info.ReleaseGraphics()
                        Return AddRowBut
                        Exit Function
                    End If

                    info.ReleaseGraphics()

                End If

            Next

            Return Nothing

        End Function

        Private Function IsButtonRectBandGColFooter(ByVal point As Point, ByVal FTCell As GridFooterCellInfoArgs) As ActionButton

            If ActiveColAddRowButtons Is Nothing Then Return Nothing
            If ActiveColAddRowButtons.Length = 0 Then Return Nothing

            For Each AddRowBut In ActiveColAddRowButtons

                If AddRowBut.Column Is FTCell.Column Then

                    Dim info As New GraphicsInfo()
                    info.AddGraphics(Nothing)
                    Dim buttonRect As Rectangle = CalcButtonRectFooter(FTCell, info.Graphics)

                    If buttonRect.Contains(point) Then
                        info.ReleaseGraphics()
                        Return AddRowBut
                        Exit Function

                    End If

                    info.ReleaseGraphics()

                End If

            Next

            Return Nothing

        End Function

        Private Function IsButtonRect(ByVal point As Point, ByVal column As GridColumn) As ActionButton

            If ActiveColAddRowButtons Is Nothing Then Return Nothing
            If ActiveColAddRowButtons.Length = 0 Then Return Nothing

            For Each AddRowBut In ActiveColAddRowButtons

                If AddRowBut.Column Is column Then

                    Dim info As New GraphicsInfo()
                    info.AddGraphics(Nothing)
                    Dim viewInfo As GridViewInfo = TryCast(view.GetViewInfo(), GridViewInfo)
                    Dim columnArgs As GridColumnInfoArgs = viewInfo.ColumnsInfo(column)
                    Dim buttonRect As Rectangle = CalcButtonRect(columnArgs, info.Graphics)

                    If buttonRect.Contains(point) Then
                        info.ReleaseGraphics()
                        Return AddRowBut
                        Exit Function
                    End If

                    info.ReleaseGraphics()

                End If

            Next

            Return Nothing

        End Function

        Private Function IsButtonRectFooter(ByVal point As Point, ByVal FooterCell As GridFooterCellInfoArgs) As ActionButton

            If ActiveColAddRowButtons Is Nothing Then Return Nothing
            If ActiveColAddRowButtons.Length = 0 Then Return Nothing

            For Each AddRowBut In ActiveColAddRowButtons

                If AddRowBut.Column Is FooterCell.Column Then

                    Dim info As New GraphicsInfo()
                    info.AddGraphics(Nothing)

                    Dim buttonRect As Rectangle = CalcButtonRectFooter(FooterCell, info.Graphics)

                    If buttonRect.Contains(point) Then
                        info.ReleaseGraphics()
                        Return AddRowBut
                        Exit Function
                    End If

                    info.ReleaseGraphics()

                End If

            Next

            Return Nothing

        End Function

        'Private Function IsButtonFooterRect(ByVal point As Point, ByVal column As GridFooterCellInfoArgs) As ActionButton

        '    If ActiveColAddRowButtons Is Nothing Then Return Nothing
        '    If ActiveColAddRowButtons.Length = 0 Then Return Nothing

        '    For Each AddRowBut In ActiveColAddRowButtons

        '        If AddRowBut.Column Is column Then

        '            Dim info As New GraphicsInfo()
        '            info.AddGraphics(Nothing)
        '            Dim viewInfo As GridViewInfo = TryCast(view.GetViewInfo(), GridViewInfo)
        '            Dim columnArgs As GridColumnInfoArgs = viewInfo.ColumnsInfo(column)
        '            Dim buttonRect As Rectangle = CalcButtonRect(columnArgs, info.Graphics)

        '            If buttonRect.Contains(point) Then
        '                info.ReleaseGraphics()
        '                Return AddRowBut
        '                Exit Function
        '            End If

        '            info.ReleaseGraphics()

        '        End If

        '    Next

        '    Return Nothing

        'End Function
        Private Function GetPaintedActionButtonAtPoint(
            ByVal PointToTest As Point) As ActionButton

            If ActiveButtons Is Nothing Then Return Nothing
            If ActiveButtons.Length = 0 Then Return Nothing

            For Each But As ActionButton In ActiveButtons

                If But Is Nothing Then Continue For

                If Not But.BandRect.IsEmpty AndAlso
                   But.BandRect.Contains(PointToTest) Then

                    Return But

                End If

                If Not But.HeaderRect.IsEmpty AndAlso
                   But.HeaderRect.Contains(PointToTest) Then

                    Return But

                End If

                If Not But.FooterRect.IsEmpty AndAlso
                   But.FooterRect.Contains(PointToTest) Then

                    Return But

                End If

            Next

            Return Nothing

        End Function

        Private Sub SetActionButtonState(
            ByVal But As ActionButton,
            ByVal State As ObjectState)

            If But Is Nothing Then Exit Sub

            If But.BandMode AndAlso But.Band IsNot Nothing Then

                SetButtonStateBand(
                    But.Band,
                    State)

            ElseIf But.Column IsNot Nothing Then

                SetButtonState(
                    But.Column,
                    State)

            End If

        End Sub

        Private Sub OnMouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)

            If ActiveButtons Is Nothing Then Return
            If ActiveButtons.Length = 0 Then Return

            'Primary path: use the exact rectangle that was custom-painted.
            Dim PaintedButton As ActionButton =
                GetPaintedActionButtonAtPoint(
                    e.Location)

            If PaintedButton IsNot Nothing Then

                SetActionButtonState(
                    PaintedButton,
                    ObjectState.Normal)

                MyOwner.RunAction(
                    PaintedButton.ActionToken)

                DXMouseEventArgs.GetMouseArgs(e).Handled = True

                Return

            End If

            'Fallback: retain the original DevExpress hit-test routing for the
            'unusual case where interaction occurs before the first custom paint.
            Dim hitInfo As GridHitInfo = Nothing
            Dim hitInfoBand As BandedGridHitInfo = Nothing

            If BandedMode Then

                hitInfoBand = bgview.CalcHitInfo(e.Location)

                If hitInfoBand.HitTest <> BandedGridHitTest.Band Then

                    'CheckForColumn

                    If hitInfoBand.HitTest <> BandedGridHitTest.Column Then

                        If hitInfoBand.HitTest <> BandedGridHitTest.Footer Then

                            DXMouseEventArgs.GetMouseArgs(e).Handled = True
                            Return
                            Exit Sub

                        Else

                            If hitInfoBand.FooterCell IsNot Nothing Then

                                GoTo BandFooter

                            Else

                                Return
                                Exit Sub

                            End If

                        End If

                    Else

                        GoTo ColProcessBGC

                    End If

                Else

                    GoTo BandProcess

                End If

            Else

                hitInfo = view.CalcHitInfo(e.Location)

                If hitInfo.HitTest <> GridHitTest.Column Then

                    If hitInfo.HitTest <> GridHitTest.Footer Then

                        DXMouseEventArgs.GetMouseArgs(e).Handled = True
                        Return
                        Exit Sub

                    Else

                        If hitInfo.FooterCell IsNot Nothing Then

                            GoTo NormalFooter

                        Else

                            Return
                            Exit Sub

                        End If



                    End If

                Else

                    GoTo NormalCol

                End If

            End If

NormalCol:

            Dim column As GridColumn = hitInfo.Column

            If IsColumnActive(column) Is Nothing Then Return

            Dim ActBut = IsButtonRect(e.Location, column)

            If ActBut Is Nothing Then

                ActBut = IsButtonRect(e.Location, column)

                If ActBut Is Nothing Then Return

            End If

            SetButtonState(column, ObjectState.Normal)
            MyOwner.RunAction(ActBut.ActionToken)
            DXMouseEventArgs.GetMouseArgs(e).Handled = True

            Exit Sub

NormalFooter:

            Dim FTcell As GridFooterCellInfoArgs = hitInfo.FooterCell

            If IsColumnActive(FTcell.Column) Is Nothing Then Return

            Dim ActButNFooter = IsButtonRectFooter(e.Location, FTcell)

            If ActButNFooter Is Nothing Then Return

            SetButtonState(ActButNFooter.Column, ObjectState.Normal)
            MyOwner.RunAction(ActButNFooter.ActionToken)
            DXMouseEventArgs.GetMouseArgs(e).Handled = True

            Return
            Exit Sub

BandFooter:

            Dim FTcellBand As GridFooterCellInfoArgs = hitInfoBand.FooterCell

            If IsColumnActive(FTcellBand.Column) Is Nothing Then Return

            Dim ActButBandooter = IsButtonRectFooter(e.Location, FTcellBand)

            If ActButBandooter Is Nothing Then Return

            SetButtonState(ActButBandooter.Column, ObjectState.Normal)
            MyOwner.RunAction(ActButBandooter.ActionToken)
            DXMouseEventArgs.GetMouseArgs(e).Handled = True

            Return
            Exit Sub

ColProcessBGC:

            Dim BGcolumn As BandedGridColumn = hitInfoBand.Column

            If IsColumnActive(BGcolumn) Is Nothing Then Return

            Dim ActButBGC = IsButtonRectBandGCol(e.Location, BGcolumn)

            If ActButBGC Is Nothing Then Return

            SetButtonState(BGcolumn, ObjectState.Normal)
            MyOwner.RunAction(ActButBGC.ActionToken)
            DXMouseEventArgs.GetMouseArgs(e).Handled = True

            Exit Sub

BandProcess:

            Dim band As GridBand = hitInfoBand.Band

            If IsBandActive(band) Is Nothing Then Return

            Dim BandBut = IsButtonRectBand(e.Location, band)

            If BandBut Is Nothing Then Return

            SetButtonStateBand(band, ObjectState.Normal)
            MyOwner.RunAction(BandBut.ActionToken)
            DXMouseEventArgs.GetMouseArgs(e).Handled = True


        End Sub

        Private Sub OnMouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)



            ''''''''''''''''''''''''''''''''''''''''''


            If ActiveButtons Is Nothing Then Return
            If ActiveButtons.Length = 0 Then Return

            Dim PaintedButton As ActionButton =
                GetPaintedActionButtonAtPoint(
                    e.Location)

            If PaintedButton IsNot Nothing Then

                SetActionButtonState(
                    PaintedButton,
                    ObjectState.Hot)

                DXMouseEventArgs.GetMouseArgs(e).Handled = True

                Return

            End If

            Dim hitInfo As GridHitInfo = Nothing
            Dim hitInfoBand As BandedGridHitInfo = Nothing


            If BandedMode Then

                hitInfoBand = bgview.CalcHitInfo(e.Location)

                If hitInfoBand.HitTest <> BandedGridHitTest.Band Then

                    'CheckForColumn

                    If hitInfoBand.HitTest <> BandedGridHitTest.Column Then

                        'check for footer

                        If hitInfoBand.HitTest <> BandedGridHitTest.Footer Then

                            Return

                        Else

                            If hitInfoBand.FooterCell IsNot Nothing Then

                                GoTo ColProcessBCGF

                            Else

                                Return

                            End If

                        End If


                    Else

                        GoTo ColProcessBCG

                    End If

                Else

                    GoTo BandProcess

                End If

            Else


                hitInfo = view.CalcHitInfo(e.Location)

                If hitInfo.HitTest <> GridHitTest.Column Then

                    If hitInfo.HitTest <> GridHitTest.Footer Then

                        Return
                        Exit Sub

                    Else

                        If hitInfo.FooterCell IsNot Nothing Then

                            GoTo NormalFooter

                        Else

                            Return
                            Exit Sub

                        End If


                    End If

                Else

                    GoTo NormalColumn

                End If

            End If

NormalColumn:

            Dim column As GridColumn = hitInfo.Column

            If IsColumnActive(column) Is Nothing Then Return

            Dim ActBut = IsButtonRect(e.Location, column)

            If ActBut Is Nothing Then

                SetButtonState(column, ObjectState.Normal)

            Else

                SetButtonState(column, ObjectState.Hot)

            End If

            'MyOwner.RunAction(ActBut.ActionToken)
            DXMouseEventArgs.GetMouseArgs(e).Handled = True

            Return
            Exit Sub

NormalFooter:

            Dim FTcell As GridFooterCellInfoArgs = hitInfo.FooterCell

            If IsColumnActive(FTcell.Column) Is Nothing Then Return

            Dim ActButNFooter = IsButtonRectFooter(e.Location, FTcell)

            If ActButNFooter Is Nothing Then

                SetButtonState(FTcell.Column, ObjectState.Normal)

            Else

                SetButtonState(FTcell.Column, ObjectState.Hot)

            End If

            'view.InvalidateFooter()
            DXMouseEventArgs.GetMouseArgs(e).Handled = True

            Return
            Exit Sub

ColProcessBCG:

            Dim BGcolumn As BandedGridColumn = hitInfoBand.Column

            If IsColumnActive(BGcolumn) Is Nothing Then Return

            Dim ActButBGC = IsButtonRectBandGCol(e.Location, BGcolumn)

            If ActButBGC Is Nothing Then

                SetButtonState(BGcolumn, ObjectState.Normal)

            Else

                SetButtonState(BGcolumn, ObjectState.Hot)

            End If

            DXMouseEventArgs.GetMouseArgs(e).Handled = True
            Return
            Exit Sub

BandProcess:

            Dim band As GridBand = hitInfoBand.Band

            If IsBandActive(band) Is Nothing Then Return

            Dim BandBut = IsButtonRectBand(e.Location, band)

            If BandBut Is Nothing Then

                SetButtonStateBand(band, ObjectState.Normal)

            Else

                SetButtonStateBand(band, ObjectState.Hot)

            End If

            Return
            DXMouseEventArgs.GetMouseArgs(e).Handled = True

ColProcessBCGF:

            Dim BGFTcell As GridFooterCellInfoArgs = hitInfoBand.FooterCell

            If IsColumnActive(BGFTcell.Column) Is Nothing Then Return

            Dim ActButBGCF = IsButtonRectFooter(e.Location, BGFTcell)

            If ActButBGCF Is Nothing Then

                SetButtonState(BGFTcell.Column, ObjectState.Normal)

            Else

                SetButtonState(BGFTcell.Column, ObjectState.Hot)

            End If

            'bgview.InvalidateFooter()

            DXMouseEventArgs.GetMouseArgs(e).Handled = True
            Return
            Exit Sub

            ''''''''''''''''''''''''''''''''''''''''


        End Sub

        Private Sub OnMouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)

            'Dim hitInfo As GridHitInfo = view.CalcHitInfo(e.Location)

            'If hitInfo.HitTest <> GridHitTest.Column Then
            '    Return
            'End If

            'Dim column As GridColumn = hitInfo.Column

            'If IsButtonRect(e.Location, column) Then

            '    SetButtonState(column, ObjectState.Pressed)
            '    DXMouseEventArgs.GetMouseArgs(e).Handled = True

            'End If

        End Sub

        Private Sub SetButtonState(ByVal column As GridColumn, ByVal state As ObjectState)


            For Each ButtonCol In ActiveColumns
                If ButtonCol.Column Is column Then
                    If ButtonCol.ActionButton.InfoArgs.State = state Then Return
                    ButtonCol.ActionButton.InfoArgs.State = state
                    Exit For
                End If
            Next



            If BandedMode Then
                bgview.InvalidateColumnHeader(column)
                bgview.InvalidateFooter()
            Else
                view.InvalidateColumnHeader(column)
                view.InvalidateFooter()
            End If


        End Sub
        Private Sub SetButtonStateBand(ByVal band As GridBand, ByVal state As ObjectState)

            For Each ButtonBand In ActiveBands
                If ButtonBand.Band Is band Then
                    ButtonBand.ActionButton.InfoArgs.State = state
                End If
            Next

            bgview.InvalidateBandHeader(band)

        End Sub

        Private Function CalcButtonRectBand(ByVal bandArgs As GridBandInfoArgs, ByVal gr As Graphics) As Rectangle

            Dim bandRect As Rectangle = bandArgs.Bounds
            Dim innerElementsWidth As Integer = CalcInnerElementsMinWidthBand(bandArgs, gr)
            'Dim buttonRect As New Rectangle(columnRect.Left + (columnRect.Width / 2) - (buttonSize.Width / 2), columnRect.Y - columnRect.Height \ 2 + buttonSize.Height \ 2, buttonSize.Width, buttonSize.Height)
            Dim buttonRect As New Rectangle(bandRect.Right - innerElementsWidth - buttonSize.Width - 2, bandRect.Y + bandRect.Height \ 2 - buttonSize.Height \ 2, buttonSize.Width, buttonSize.Height)
            Return buttonRect

        End Function
        Private Function CalcButtonRect(ByVal columnArgs As GridColumnInfoArgs, ByVal gr As Graphics, Optional ByVal SentFh As Single = -1) As Rectangle

            If SentFh = -1 Then SentFh = CurrFH

            Dim columnRect As Rectangle = columnArgs.Bounds
            Dim innerElementsWidth As Integer = CalcInnerElementsMinWidth(columnArgs, gr)
            Dim buttonRect As New Rectangle(columnRect.Left + 5, columnRect.Y + columnRect.Height - (buttonSize.Height / 2) - ((DefaultGridCellPadding / 2) + (SentFh / 2)), buttonSize.Width, buttonSize.Height)

            Return buttonRect

        End Function
        Private Function CalcButtonRectFooter(ByVal FooterCellArgs As GridFooterCellInfoArgs, ByVal gr As Graphics, Optional ByVal SentFh As Single = -1) As Rectangle

            If SentFh = -1 Then SentFh = CurrFH

            Dim columnRect As Rectangle = FooterCellArgs.Bounds

            Dim buttonRect As New Rectangle(columnRect.Left - 5, columnRect.Y - (buttonSize.Height / 2) + (DefaultGridCellPadding / 2) + (SentFh / 2), buttonSize.Width, buttonSize.Height)

            Return buttonRect

        End Function
        Private Function CalcInnerElementsMinWidthBand(ByVal bandArgs As GridBandInfoArgs, ByVal gr As Graphics) As Integer

            Dim canDrawMode As Boolean = True
            Return bandArgs.InnerElements.CalcMinSize(gr, canDrawMode).Width

        End Function
        Private Function CalcInnerElementsMinWidth(ByVal columnArgs As GridColumnInfoArgs, ByVal gr As Graphics) As Integer

            Dim canDrawMode As Boolean = True
            Return columnArgs.InnerElements.CalcMinSize(gr, canDrawMode).Width

        End Function
        Private Function CalcInnerElementsMinWidthFooterCell(ByVal FooterCellInfoArgs As GridFooterCellInfoArgs, ByVal gr As Graphics) As Integer

            Dim canDrawMode As Boolean = True

            Return FooterCellInfoArgs.Column.Width                 'InnerElements.CalcMinSize(gr, canDrawMode).Width

        End Function




        Private Function IsColumnActive(ByVal column As GridColumn) As ButtonedColumn
            If ActiveColumns Is Nothing Then Return Nothing
            If ActiveColumns.Length = 0 Then Return Nothing
            For Each col In ActiveColumns
                If col.Column Is column Then Return col
            Next
            Return Nothing
        End Function
        Private Function IsBandActive(ByVal band As GridBand) As ButtonedBand
            If ActiveBands Is Nothing Then Return Nothing
            If ActiveBands.Length = 0 Then Return Nothing

            For Each bnd In ActiveBands
                If bnd.Band Is band Then Return bnd
            Next
            Return Nothing
        End Function
        Private Sub DrawCustomButton(ByVal e As ColumnHeaderCustomDrawEventArgs, But As ActionButton)

            SetUpButtonInfoArgs(e, But)

            customButtonPainter.DrawObject(But.InfoArgs)

        End Sub

        Private Sub DrawCustomButtonFooter(ByVal e As FooterCellCustomDrawEventArgs, But As ActionButton)

            SetUpButtonInfoArgsFooter(e, But)

            customButtonPainter.DrawObject(But.InfoArgs)

        End Sub

        Private Sub DrawCustomButtonBand(ByVal e As BandHeaderCustomDrawEventArgs, But As ActionButton)

            SetUpBandButtonInfoArgs(e, But)

            customButtonPainter.DrawObject(But.InfoArgs)

        End Sub

        Private Sub SetUpButtonInfoArgs(ByVal e As ColumnHeaderCustomDrawEventArgs, But As ActionButton)

            Dim FH = e.Appearance.GetFont().GetHeight
            CurrFH = FH
            But.InfoArgs.Cache = e.Cache
            But.InfoArgs.Bounds = CalcButtonRect(e.Info, e.Cache.Graphics, FH)
            But.HeaderRect = But.InfoArgs.Bounds

        End Sub

        Private Sub SetUpButtonInfoArgsFooter(ByVal e As FooterCellCustomDrawEventArgs, But As ActionButton)

            But.InfoArgs.Cache = e.Cache
            But.InfoArgs.Bounds = CalcButtonRectFooter(e.Info, e.Cache.Graphics)
            But.FooterRect = But.InfoArgs.Bounds

        End Sub

        Private Sub SetUpBandButtonInfoArgs(ByVal e As BandHeaderCustomDrawEventArgs, But As ActionButton)

            But.InfoArgs.Cache = e.Cache
            But.InfoArgs.Bounds = CalcButtonRectBand(e.Info, e.Cache.Graphics)
            But.BandRect = But.InfoArgs.Bounds

        End Sub
        Private Sub OnCustomDrawColumnHeader(ByVal sender As Object, ByVal e As ColumnHeaderCustomDrawEventArgs)
            Debug.Print("XXXXXXXXX - CD Header - XXXXXXX")
            If e.Column Is Nothing Then

                e.Handled = True
                Return

            End If


            Dim ColCheck As ButtonedColumn = IsColumnActive(e.Column)

            If ColCheck Is Nothing Then GoTo DefaultDraw

            'If e.Column.Tag.WidthSet = False Then

            '    Dim CellPadding As Integer = 28
            '    e.Column.Tag.DefaultColumnWidth = e.Column.Width
            '    e.Column.Tag.ExtendedColumnWidth = e.Column.Width + CellPadding + 40
            '    e.Column.Tag.WidthSet = True
            '    e.Column.MinWidth = e.Column.Width + CellPadding + 40
            '    e.Column.Width = e.Column.Width + CellPadding + 40

            '    MyOwner.ResizeFonts()
            '    MustInvalidateFooter = True

            '    If BandedMode Then

            '        ' bgview.Invalidate()
            '        'MyOwner.ResizeFonts()
            '        bgview.GridControl.Invalidate()
            '        'bgview.InvalidateFooter()

            '    Else

            '        'view.Invalidate()
            '        'MyOwner.ResizeFonts()
            '        view.GridControl.Invalidate()
            '        'view.InvalidateFooter()

            '    End If

            '    e.Handled = True
            '    Return

            'End If

            Dim FH = e.Appearance.GetFont().GetHeight
            WithButtonDrawColumnHeader(e)
            DrawCustomButton(e, ColCheck.ActionButton)

            e.Handled = True

            Return

            Exit Sub

DefaultDraw:

            DefaultDrawColumnHeader(e)

            e.Handled = True

            Return

        End Sub
        Private Sub DefaultDrawColumnHeader(ByVal e As ColumnHeaderCustomDrawEventArgs)

            Dim ColTag As DataColumnTag = e.Column.Tag

            'Debug.Print("CombinedEditDrawer - DefaultDrawColumnHeader - " & e.Column.AbsoluteIndex)

            If ColTag.IsDummyColumn Then
                e.Handled = True
                Return
            End If

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

            e.DefaultDraw()

            e.Handled = True

            Return

        End Sub
        Private Sub WithButtonDrawColumnHeader(ByVal e As ColumnHeaderCustomDrawEventArgs)

            Dim bounds As Rectangle = e.Bounds

            Dim CCellSFormat As StringFormat = e.Appearance.GetStringFormat().Clone

            CCellSFormat.Alignment = StringAlignment.Near

            Dim FH = e.Appearance.GetFont().GetHeight

            Dim BottomLineY As Integer = e.Bounds.Height - (DefaultGridCellPadding / 2) - FH

            bounds.Y += BottomLineY
            bounds.X += 33
            bounds.Height = FH

            'Debug.Print("CombinedEditDrawer - WithButtonDrawColumnHeader - " & e.Column.AbsoluteIndex & " : " & e.Info.Caption)

            Dim ForeBrush = New SolidBrush(Color.Black)

            e.Cache.DrawString(e.Info.Caption, e.Appearance.GetFont(), ForeBrush, bounds, CCellSFormat)

            If MustInvalidateFooter Then

                MustInvalidateFooter = False

                'If BandedMode Then
                '    bgview.DataController.TotalSummary.SetDirty()
                '    bgview.Invalidate()
                'Else
                '    view.DataController.TotalSummary.SetDirty()
                '    view.Invalidate()
                'End If



                'If BandedMode Then

                '    bgview.Invalidate()

                'Else

                '    view.Invalidate()

                'End If

            End If

            e.Handled = True

        End Sub
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

            Dim BandCheck As ButtonedBand = IsBandActive(e.Band)

            If BandTag.HasActions Then DrawCustomButtonBand(e, BandCheck.ActionButton)

            WithButtonDrawBandHeader(e)

            e.Handled = True

            Return

        End Sub


        Private Sub DefaultDrawBand(ByVal e As BandHeaderCustomDrawEventArgs)

            Dim CellPadding As Integer = 28
            Dim pen As Pen = New Pen(Color.Black, 4)
            Dim bounds As Rectangle = e.Bounds

            e.Cache.DrawString(e.Info.Caption, e.Appearance.GetFont(), e.Appearance.GetForeBrush(e.Cache), e.Bounds, e.Appearance.GetStringFormat())

        End Sub
        Private Sub WithButtonDrawBandHeader(ByVal e As BandHeaderCustomDrawEventArgs)

            Dim CellPadding As Integer = 28
            Dim bounds As Rectangle = e.Bounds
            bounds.Y -= 5
            Dim BandTag As BandTag = e.Band.Tag

            Dim pen As Pen = New Pen(BandTag.HighLightColour, 4)

            e.Cache.DrawString(e.Info.Caption, e.Appearance.GetFont(), e.Appearance.GetForeBrush(e.Cache), bounds, e.Appearance.GetStringFormat())

            e.Cache.DrawLine(pen, New Point(e.Bounds.X + CellPadding, e.Bounds.Bottom), New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Bottom))

            pen.Dispose()

            e.Handled = True

            Return

        End Sub

        Private Sub GVCustomDrawFooterCell(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.FooterCellCustomDrawEventArgs)

            Debug.Print("In Custom FOOTER Cell Draw col " & e.Column.AbsoluteIndex)

            If e.Column Is Nothing Then
                Debug.Print("Col is nothing.....")
                e.Handled = True
                Return
            End If

            Dim r As New Rectangle With {
            .X = e.Bounds.X,
            .Height = e.Bounds.Height,
            .Y = e.Bounds.Y,
            .Width = e.Bounds.Width - DefaultGridCellPadding
            }

            ' e.Cache.DrawString("Add lines", e.Appearance.GetFont(), e.Appearance.GetForeBrush(e.Cache), r, e.Appearance.GetStringFormat())

            Dim ColCheck As ButtonedColumn = IsColumnActive(e.Column)

            If ColCheck Is Nothing Then
                Debug.Print("Colinactive, defult draw")
                GoTo DrawDefault

            Else
                Debug.Print("Col is active, with button draw")
                GoTo WithButtonDraw

            End If

            If e.Info.DisplayText Is Nothing Then

                e.DefaultDraw()
                Return

            End If

DrawDefault:

            Dim pen As Pen = New Pen(AbovoBlue, 1)
            Dim StrToWrite As String = ""

            e.Cache.DrawLine(pen, New Point(e.Bounds.X, e.Bounds.Top), New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Top))
            e.Cache.DrawLine(pen, New Point(e.Bounds.X, e.Bounds.Bottom), New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Bottom))

            e.Appearance.ForeColor = AbovoBlue
            StrToWrite = e.Info.DisplayText

            If IsNumeric(StrToWrite) Then
                If CDbl(e.Info.DisplayText) < 0 Then

                    If Microsoft.VisualBasic.Left(StrToWrite, 1) = "-" Then StrToWrite = Microsoft.VisualBasic.Right(StrToWrite, Len(StrToWrite) - 1)
                    e.Appearance.ForeColor = Color.Red
                    StrToWrite = "(" & StrToWrite & ")"

                End If
            End If

            e.Appearance.DrawString(e.Cache, StrToWrite, r)
            e.Handled = True
            Return

WithButtonDraw:

            'Dim pen2 As Pen = New Pen(Color.Pink, 4)
            'e.Cache.DrawLine(pen2, New Point(e.Bounds.X, e.Bounds.Top), New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Top))
            'e.Cache.DrawLine(pen2, New Point(e.Bounds.X, e.Bounds.Bottom), New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Bottom))
            'e.Cache.DrawLine(pen2, New Point(e.Bounds.X, e.Bounds.Top), New Point(e.Bounds.X, e.Bounds.Bottom))
            'e.Cache.DrawLine(pen2, New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Top), New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Bottom))

            DrawCustomButtonFooter(e, ColCheck.ActionButton)
            Dim BoldF As Font = New Font(e.Appearance.GetFont(), FontStyle.Bold)

            Dim AppSF As StringFormat = e.Appearance.GetStringFormat().Clone
            AppSF.Alignment = StringAlignment.Near

            Dim FH = e.Appearance.GetFont().GetHeight
            Dim BottomLineY As Integer = e.Bounds.Y + e.Bounds.Height - (DefaultGridCellPadding / 2) - FH

            r.Y += (DefaultGridCellPadding / 2)
            r.X += 28
            r.Height = FH

            e.Cache.DrawString("Add lines", BoldF, e.Appearance.GetForeBrush(e.Cache), r, AppSF)
            Debug.Print("With button draw footer cell" & e.Column.AbsoluteIndex)
            e.Handled = True
            Return

        End Sub

        Private Sub GVCustomDrawFooter(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Base.RowObjectCustomDrawEventArgs)
            Debug.Print("CustomDrawFooter")
            e.Handled = True
            Return

        End Sub

        Private Sub UnsubscribeFromEvents()

            If BandedMode Then

                If bgview Is Nothing Then Exit Sub

                RemoveHandler bgview.CustomDrawColumnHeader, AddressOf OnCustomDrawColumnHeader
                RemoveHandler bgview.CustomDrawBandHeader, AddressOf OnCustomDrawBandHeader
                RemoveHandler bgview.MouseDown, AddressOf OnMouseDown
                RemoveHandler bgview.MouseUp, AddressOf OnMouseUp
                RemoveHandler bgview.MouseMove, AddressOf OnMouseMove
                RemoveHandler bgview.CustomDrawFooterCell, AddressOf GVCustomDrawFooterCell
                RemoveHandler bgview.CustomDrawFooter, AddressOf GVCustomDrawFooter

            Else

                If view Is Nothing Then Exit Sub

                RemoveHandler view.CustomDrawColumnHeader, AddressOf OnCustomDrawColumnHeader
                RemoveHandler view.MouseDown, AddressOf OnMouseDown
                RemoveHandler view.MouseUp, AddressOf OnMouseUp
                RemoveHandler view.MouseMove, AddressOf OnMouseMove
                RemoveHandler view.CustomDrawFooterCell, AddressOf GVCustomDrawFooterCell
                RemoveHandler view.CustomDrawFooter, AddressOf GVCustomDrawFooter

            End If

        End Sub

        Public Sub RemoveCustomButton()

            UnsubscribeFromEvents()

            If ActiveButtons IsNot Nothing Then

                For Each But As ActionButton In ActiveButtons

                    If But Is Nothing Then Continue For

                    But.HeaderRect = Rectangle.Empty
                    But.FooterRect = Rectangle.Empty
                    But.BandRect = Rectangle.Empty

                Next

            End If

        End Sub


    End Class

End Namespace