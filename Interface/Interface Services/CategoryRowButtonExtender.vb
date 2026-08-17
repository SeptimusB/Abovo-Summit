Imports System.Drawing
Imports System.Windows.Forms
Imports Abovo.GeneralFunctions
Imports DevExpress.Utils
Imports DevExpress.Utils.Drawing
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraEditors.Drawing
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.XtraVerticalGrid.Events
Imports DevExpress.XtraVerticalGrid.Rows


Namespace Abovo

    ''' <summary>
    ''' Marker stored in the synthetic VGrid footer row Tag.
    ''' The footer row is deliberately not a DataColumnTag row, so the normal
    ''' VGrid edit/paste/custom-editor handlers ignore it.
    ''' </summary>
    Public Class VGridCategoryFooterTag

        Public Property ParentCategory As CategoryRow

    End Class


    ''' <summary>
    ''' Gives a VGrid CategoryRow the same visual/action concept used by the
    ''' XtraGrid band/footer code.
    '''
    ''' A VGrid has no native per-category footer.  This class therefore adds
    ''' one read-only, unbound EditorRow as the final child of an ACTIONABLE category
    ''' and custom paints that row as a footer. Categories without an ActionToken
    ''' do not receive a footer row. The footer contains the + button and "Add rows" text.
    '''
    ''' The category header itself retains its normal DevExpress painting and
    ''' only receives the Abovo highlight line; the action button is deliberately
    ''' in the footer rather than the category header.
    ''' </summary>
    Public Class VGridCategoryButtonExtender

        Private view As VGridControl
        Private ActiveCategoryRow As CategoryRow
        Private MyOwner As DataInterfaceTemplate

        Private customButtonPainter As SkinEditorButtonPainter
        Private args As EditorButtonObjectInfoArgs

        Private FooterRow As EditorRow
        Private ReadOnly buttonSize As Size = New Size(28, 28)
        Private ReadOnly footerHeight As Integer = 38

        Private CurrentFooterButtonRect As Rectangle = Rectangle.Empty
        Private EventsSubscribed As Boolean = False


        Public Sub New(ByVal view As VGridControl,
                       ByVal Category As CategoryRow,
                       ByVal Opener As DataInterfaceTemplate,
                       ByVal DefaultAction As String,
                       ByVal SetSectionID As Integer,
                       ByVal SetTitle As String,
                       ByVal ActTok As ActionToken)

            Me.view = view
            Me.ActiveCategoryRow = Category
            Me.MyOwner = Opener

            'Retain compatibility with the first VGrid implementation.  The
            'BandTag remains the single source of truth for the action.
            Dim RowTag As BandTag = GetCategoryTag()

            If RowTag IsNot Nothing AndAlso ActTok IsNot Nothing Then
                RowTag.ActionToken = ActTok
                RowTag.HasActions = True
            End If

        End Sub


        Public ReadOnly Property CategoryFooterRow As EditorRow
            Get
                Return FooterRow
            End Get
        End Property


        Public Sub AddCustomButton()

            If view Is Nothing OrElse ActiveCategoryRow Is Nothing Then Return

            'Every category keeps the custom category-header drawing, but a
            'synthetic footer row exists only when there is a real action/button.
            If CategoryHasAction() Then

                EnsureFooterRow()
                CreateButtonPainter()
                CreateButtonInfoArgs()

            End If

            SubscribeToEvents()

            view.InvalidateRow(ActiveCategoryRow)

            If FooterRow IsNot Nothing Then
                view.InvalidateRow(FooterRow)
            End If

        End Sub


        Private Function GetCategoryTag() As BandTag

            If ActiveCategoryRow Is Nothing Then Return Nothing

            Return TryCast(ActiveCategoryRow.Tag, BandTag)

        End Function


        Private Function CategoryHasAction() As Boolean

            Dim RowTag As BandTag = GetCategoryTag()

            Return RowTag IsNot Nothing AndAlso
                   RowTag.HasActions AndAlso
                   RowTag.ActionToken IsNot Nothing

        End Function


        Private Sub EnsureFooterRow()

            If FooterRow IsNot Nothing Then Return
            If Not CategoryHasAction() Then Return

            'Do not create a second footer if AddCustomButton is called twice.
            For Each Child As BaseRow In ActiveCategoryRow.ChildRows

                Dim FooterTag As VGridCategoryFooterTag =
                    TryCast(Child.Tag, VGridCategoryFooterTag)

                If FooterTag IsNot Nothing Then

                    FooterRow = TryCast(Child, EditorRow)
                    Return

                End If

            Next

            FooterRow = New EditorRow() With {
                .Name = "VGridFooter_" & ActiveCategoryRow.Name,
                .Height = Math.Max(footerHeight, ActiveCategoryRow.Height),
                .Tag = New VGridCategoryFooterTag With {
                    .ParentCategory = ActiveCategoryRow
                }
            }

            With FooterRow.Properties

                .Caption = ""
                .ReadOnly = True

                'Deliberately leave FieldName blank.  DevExpress permits an
                'EditorRow created with the parameterless constructor to remain
                'unbound; all visible content is painted by this extender.

            End With

            'Use the same appearance object that ObjectFormatter.FormatVertGrid
            'uses for category/header rows (VGrid.Appearance.Category).
            FooterRow.AppearanceHeader.Assign(view.Appearance.Category)
            FooterRow.AppearanceCell.Assign(view.Appearance.Category)

            ActiveCategoryRow.ChildRows.Add(FooterRow)

        End Sub


        Private Sub CreateButtonInfoArgs()

            Dim btn As New EditorButton(ButtonPredefines.Plus)

            btn.ToolTip = "Add Rows"
            btn.ToolTipAnchor = ToolTipAnchor.Cursor
            btn.IsLeft = True

            args = New EditorButtonObjectInfoArgs(
                btn,
                New DevExpress.Utils.AppearanceObject()
            )

        End Sub


        Private Sub CreateButtonPainter()

            customButtonPainter =
                New SkinEditorButtonPainter(
                    DevExpress.LookAndFeel.UserLookAndFeel.Default.ActiveLookAndFeel
                )

        End Sub


        Private Sub SubscribeToEvents()

            If EventsSubscribed Then Return

            AddHandler view.CustomDrawRowHeaderCell,
                       AddressOf OnCustomDrawRowHeaderCell

            AddHandler view.CustomDrawRowValueCell,
                       AddressOf OnCustomDrawRowValueCell

            AddHandler view.MouseDown,
                       AddressOf OnMouseDown

            AddHandler view.MouseUp,
                       AddressOf OnMouseUp

            AddHandler view.MouseMove,
                       AddressOf OnMouseMove

            AddHandler view.MouseLeave,
                       AddressOf OnMouseLeave

            AddHandler view.Disposed,
                       AddressOf OnViewDisposed

            EventsSubscribed = True

        End Sub


        Private Sub OnMouseDown(ByVal sender As Object,
                                ByVal e As MouseEventArgs)

            If e.Button <> MouseButtons.Left Then Return
            If Not CategoryHasAction() Then Return
            If FooterRow Is Nothing Then Return

            Dim hitInfo As VGridHitInfo = view.CalcHitInfo(e.Location)

            If hitInfo.Row IsNot FooterRow Then Return

            If IsButtonRect(e.Location) Then

                SetButtonState(ObjectState.Pressed)
                DXMouseEventArgs.GetMouseArgs(e).Handled = True

            End If

        End Sub


        Private Sub OnMouseUp(ByVal sender As Object,
                              ByVal e As MouseEventArgs)

            If e.Button <> MouseButtons.Left Then Return

            Dim WasPressed As Boolean =
                (GetButtonState() = ObjectState.Pressed)

            If FooterRow Is Nothing OrElse view Is Nothing Then Return

            Dim hitInfo As VGridHitInfo = view.CalcHitInfo(e.Location)

            If hitInfo.Row IsNot FooterRow OrElse
               Not CategoryHasAction() Then

                If WasPressed Then SetButtonState(ObjectState.Normal)
                Return

            End If

            If IsButtonRect(e.Location) Then

                SetButtonState(ObjectState.Hot)

                If WasPressed Then

                    Dim RowTag As BandTag = GetCategoryTag()

                    If RowTag IsNot Nothing AndAlso
                       RowTag.ActionToken IsNot Nothing AndAlso
                       MyOwner IsNot Nothing Then

                        'IMPORTANT:
                        'Do not execute RunAction synchronously from the VGrid MouseUp
                        'event. RunAction can rebuild/dispose this VGrid. DevExpress is
                        'still processing the current mouse message after this handler
                        'returns (including hit-testing/sorting preparation), so disposing
                        'the control here can leave its internal ViewInfo/Grid reference
                        'Nothing and cause a NullReferenceException inside the DevExpress DLL.
                        '
                        'Queue the action on the owning form instead. The current VGrid
                        'mouse event can then unwind completely before any rebuild starts.
                        Dim ActionToRun As ActionToken = RowTag.ActionToken
                        Dim OwnerToCall As DataInterfaceTemplate = MyOwner

                        If OwnerToCall.IsHandleCreated AndAlso
                           Not OwnerToCall.IsDisposed Then

                            OwnerToCall.BeginInvoke(
                                New MethodInvoker(
                                    Sub()
                                        If Not OwnerToCall.IsDisposed Then
                                            OwnerToCall.RunAction(ActionToRun)
                                        End If
                                    End Sub))

                        End If

                    End If

                End If

                DXMouseEventArgs.GetMouseArgs(e).Handled = True

            ElseIf WasPressed Then

                SetButtonState(ObjectState.Normal)

            End If

        End Sub


        Private Sub OnMouseMove(ByVal sender As Object,
                                ByVal e As MouseEventArgs)

            If FooterRow Is Nothing OrElse view Is Nothing Then Return

            If Not CategoryHasAction() Then

                If GetButtonState() <> ObjectState.Normal Then
                    SetButtonState(ObjectState.Normal)
                End If

                Return

            End If

            Dim hitInfo As VGridHitInfo = view.CalcHitInfo(e.Location)

            If hitInfo.Row Is FooterRow AndAlso
               IsButtonRect(e.Location) Then

                If GetButtonState() <> ObjectState.Pressed Then
                    SetButtonState(ObjectState.Hot)
                End If

            Else

                If GetButtonState() <> ObjectState.Normal Then
                    SetButtonState(ObjectState.Normal)
                End If

            End If

        End Sub


        Private Sub OnMouseLeave(ByVal sender As Object,
                                 ByVal e As EventArgs)

            If GetButtonState() <> ObjectState.Normal Then
                SetButtonState(ObjectState.Normal)
            End If

        End Sub


        Private Function GetButtonState() As ObjectState

            Dim RowTag As BandTag = GetCategoryTag()

            If RowTag Is Nothing Then Return ObjectState.Normal

            Return RowTag.ButtonObjectState

        End Function


        Private Sub SetButtonState(ByVal state As ObjectState)

            Dim RowTag As BandTag = GetCategoryTag()

            If RowTag Is Nothing Then Return
            If RowTag.ButtonObjectState = state Then Return

            RowTag.ButtonObjectState = state

            If view IsNot Nothing AndAlso FooterRow IsNot Nothing Then
                view.InvalidateRow(FooterRow)
            End If

        End Sub


        Private Function IsButtonRect(ByVal point As Point) As Boolean

            If CurrentFooterButtonRect.IsEmpty Then Return False

            Return CurrentFooterButtonRect.Contains(point)

        End Function


        Private Function CalcFooterButtonRect(ByVal FooterHeaderBounds As Rectangle) As Rectangle

            Dim EffectiveHeight As Integer =
                Math.Min(buttonSize.Height,
                         Math.Max(1, FooterHeaderBounds.Height - 6))

            Dim EffectiveWidth As Integer =
                Math.Min(buttonSize.Width,
                         Math.Max(1, FooterHeaderBounds.Width - 8))

            'Match the XtraGrid footer concept: button first, then the
            '"Add rows" caption alongside it.
            Dim ButtonLeft As Integer =
                FooterHeaderBounds.Left + 5

            Dim ButtonTop As Integer =
                FooterHeaderBounds.Top +
                ((FooterHeaderBounds.Height - EffectiveHeight) \ 2)

            Return New Rectangle(
                ButtonLeft,
                ButtonTop,
                EffectiveWidth,
                EffectiveHeight
            )

        End Function


        Private Sub OnCustomDrawRowHeaderCell(
            ByVal sender As Object,
            ByVal e As CustomDrawRowHeaderCellEventArgs)

            If e.Row Is Nothing Then Return

            If e.Row Is ActiveCategoryRow Then

                DrawCategoryHeader(e)
                Return

            End If

            If FooterRow IsNot Nothing AndAlso e.Row Is FooterRow Then

                DrawFooterHeader(e)
                Return

            End If

        End Sub


        Private Sub DrawCategoryHeader(
            ByVal e As CustomDrawRowHeaderCellEventArgs)

            Dim RowTag As BandTag = GetCategoryTag()

            If RowTag Is Nothing Then Return

            'Retain the native category skin/caption/tree glyph.
            e.DefaultDraw()

            If RowTag.DoBorder Then
                DrawCategoryHighlight(e, RowTag.HighLightColour)
            End If

            e.Handled = True

        End Sub


        Private Sub DrawFooterHeader(
            ByVal e As CustomDrawRowHeaderCellEventArgs)

            DrawFooterAsCategoryHeader(e.Cache, e.Bounds)

            CurrentFooterButtonRect = Rectangle.Empty

            If CategoryHasAction() Then

                CurrentFooterButtonRect = CalcFooterButtonRect(e.Bounds)
                DrawCustomButton(e)
                DrawFooterText(e)

            End If

            e.Handled = True

        End Sub


        Private Sub OnCustomDrawRowValueCell(
            ByVal sender As Object,
            ByVal e As CustomDrawRowValueCellEventArgs)

            If FooterRow Is Nothing Then Return
            If e.Row IsNot FooterRow Then Return

            'Paint every record cell with the same category/header appearance
            'so the synthetic footer reads as a full-width header-style bar.
            DrawFooterAsCategoryHeader(e.Cache, e.Bounds)

            e.Handled = True

        End Sub


        Private Sub DrawFooterAsCategoryHeader(
            ByVal Cache As DevExpress.Utils.Drawing.GraphicsCache,
            ByVal Bounds As Rectangle)

            Dim HeaderAppearance As New DevExpress.Utils.AppearanceObject()

            Try

                HeaderAppearance.Assign(view.Appearance.Category)
                HeaderAppearance.FillRectangle(Cache, Bounds)

                'Use the same category highlight colour at the bottom edge.
                'This mirrors DrawCategoryHighlight while allowing the footer
                'value cells to continue the visual bar across all records.
                Dim RowTag As BandTag = GetCategoryTag()
                Dim HighlightColour As Color = AbovoBlue

                If RowTag IsNot Nothing Then
                    HighlightColour = RowTag.HighLightColour
                End If

                Using pen As New Pen(HighlightColour, 4)

                    Cache.DrawLine(
                        pen,
                        New Point(Bounds.Left, Bounds.Bottom - 1),
                        New Point(Bounds.Right, Bounds.Bottom - 1)
                    )

                End Using

            Finally

                HeaderAppearance.Dispose()

            End Try

        End Sub


        Private Sub DrawFooterText(
            ByVal e As CustomDrawRowHeaderCellEventArgs)

            Dim TextBounds As Rectangle = e.Bounds

            TextBounds.X = CurrentFooterButtonRect.Right + 6
            TextBounds.Width =
                Math.Max(0,
                         e.Bounds.Right -
                         TextBounds.Left - 6)

            If TextBounds.Width <= 0 Then Return

            Dim HeaderAppearance As New DevExpress.Utils.AppearanceObject()
            HeaderAppearance.Assign(view.Appearance.Category)

            Using BoldFont As New Font(HeaderAppearance.GetFont(), FontStyle.Bold)

                Dim AppSF As StringFormat =
                    DirectCast(HeaderAppearance.GetStringFormat().Clone(), StringFormat)

                Try

                    AppSF.Alignment = StringAlignment.Near
                    AppSF.LineAlignment = StringAlignment.Center

                    e.Cache.DrawString(
                        "Add rows",
                        BoldFont,
                        HeaderAppearance.GetForeBrush(e.Cache),
                        TextBounds,
                        AppSF
                    )

                Finally

                    AppSF.Dispose()

                End Try

            End Using

            HeaderAppearance.Dispose()

        End Sub


        Private Sub DrawCustomButton(
            ByVal e As CustomDrawRowHeaderCellEventArgs)

            If args Is Nothing OrElse customButtonPainter Is Nothing Then Return

            args.Cache = e.Cache
            args.Bounds = CurrentFooterButtonRect
            args.State = GetButtonState()

            customButtonPainter.DrawObject(args)

        End Sub


        Private Sub DrawCategoryHighlight(
            ByVal e As CustomDrawRowHeaderCellEventArgs,
            ByVal HighlightColour As Color)

            Dim CellPadding As Integer = 28

            Using pen As New Pen(HighlightColour, 4)

                e.Cache.DrawLine(
                    pen,
                    New Point(e.Bounds.Left + CellPadding,
                              e.Bounds.Bottom - 1),
                    New Point(e.Bounds.Right,
                              e.Bounds.Bottom - 1)
                )

            End Using

        End Sub


        Private Sub OnViewDisposed(ByVal sender As Object,
                                   ByVal e As EventArgs)

            UnsubscribeFromEvents()

        End Sub


        Private Sub UnsubscribeFromEvents()

            If Not EventsSubscribed OrElse view Is Nothing Then Return

            RemoveHandler view.CustomDrawRowHeaderCell,
                          AddressOf OnCustomDrawRowHeaderCell

            RemoveHandler view.CustomDrawRowValueCell,
                          AddressOf OnCustomDrawRowValueCell

            RemoveHandler view.MouseDown,
                          AddressOf OnMouseDown

            RemoveHandler view.MouseUp,
                          AddressOf OnMouseUp

            RemoveHandler view.MouseMove,
                          AddressOf OnMouseMove

            RemoveHandler view.MouseLeave,
                          AddressOf OnMouseLeave

            RemoveHandler view.Disposed,
                          AddressOf OnViewDisposed

            EventsSubscribed = False

        End Sub


        Public Sub DetachForDisposal()

            'Used when the owning VGrid is about to be disposed.  Only detach
            'handlers/references; do not mutate ChildRows during teardown.
            'VGridControl.Dispose() will dispose its row hierarchy itself.
            UnsubscribeFromEvents()
            CurrentFooterButtonRect = Rectangle.Empty

            FooterRow = Nothing
            ActiveCategoryRow = Nothing
            view = Nothing
            MyOwner = Nothing

        End Sub


        Public Sub RemoveCustomButton()

            UnsubscribeFromEvents()

            CurrentFooterButtonRect = Rectangle.Empty

            If ActiveCategoryRow IsNot Nothing AndAlso FooterRow IsNot Nothing Then

                Try
                    ActiveCategoryRow.ChildRows.Remove(FooterRow)
                Catch
                    'The grid/category may already be disposing.
                End Try

            End If

            FooterRow = Nothing

        End Sub

    End Class

End Namespace
