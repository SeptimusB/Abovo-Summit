Imports Abovo.CustomGrid
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo

Namespace Abovo.CustomGrid

    Public Class CustomGridViewInfo

            Inherits GridViewInfo

            Public Sub New(ByVal gridView As CustomGridView)
                MyBase.New(gridView)
            End Sub

            Protected Overrides Function CalcGroupFooterHeight() As Integer
                Dim statement = TryCast(View, CustomGridView)
                If statement Is Nothing OrElse Not statement.CompactStatementRows Then Return MyBase.CalcGroupFooterHeight()
                Dim font = View.Appearance.GroupFooter.GetFont()
                Dim dpi = If(View.GridControl Is Nothing, 96, View.GridControl.DeviceDpi)
                Dim zoom = Abovo.GridPresentation.ZoomPercent(View.GridControl) / 100.0
                Dim padding = Math.Max(3, CInt(4 * dpi / 96.0 * zoom))
                Return CInt(Math.Ceiling(Math.Max(12, CInt(Math.Ceiling(font.GetHeight(CSng(dpi)))) + padding) * 1.1R))
            End Function

            Public Overrides ReadOnly Property GroupFooterCellHeight As Integer
                Get
                    Dim statement = TryCast(View, CustomGridView)
                    If statement Is Nothing OrElse Not statement.CompactStatementRows Then Return MyBase.GroupFooterCellHeight
                    'The native summary-cell bounds must shrink with their row;
                    'otherwise bottom-aligned figures are clipped by the next row.
                    Return Math.Max(10, CalcGroupFooterHeight() - 2)
                End Get
            End Property

            Protected Overrides Sub CalcRowFooterInfo(ByVal ri As GridRowInfo, ByVal row As GridRow, ByVal nextRow As GridRow)
                Dim height As Integer = ri.RowFooters.RowFootersHeight
                If height = 0 Then
                    Return
                End If

                Dim isShowCurrentFooter As Boolean = IsShowCurrentRowFooter(ri)
                Dim startLevel As Integer = ri.Level
                Dim footerRowHandle As Integer = ri.RowHandle

                If (Not ri.IsGroupRow) OrElse (Not isShowCurrentFooter) Then
                    footerRowHandle = View.GetParentRowHandle(footerRowHandle)
                End If

                If Not isShowCurrentFooter Then
                    startLevel -= 1
                End If

                Dim top As Integer = ri.TotalBounds.Bottom - height - ri.RowSeparatorBounds.Height
                Dim left As Integer = ri.IndentRect.Right - (If((Not isShowCurrentFooter), LevelIndent, 0))
                If IsRightToLeft Then
                    left = ri.TotalBounds.Left
                End If
                ri.RowFooters.Bounds = New Rectangle(left, top, ri.DataBounds.Right - left, height)

                Dim n As Integer = 0
                Do While n < ri.RowFooters.RowFooterCount
                    Dim args As New ShowGroupFooterEventArgs(startLevel)
                    RaiseShowGroupFooter(args)

                    If Not args.Visible Then
                        startLevel -= 1
                        left -= LevelIndent
                        ri.RowFooters.RowFooterCount += 1
                        footerRowHandle = View.GetParentRowHandle(footerRowHandle)

                        n += 1
                        Continue Do
                    End If

                    Dim fi As New GridRowFooterInfo()
                    ri.RowFooters.Add(fi)
                    fi.RowHandle = footerRowHandle
                    fi.Bounds = ri.Bounds
                    fi.Level = startLevel
                    fi.Bounds.Y = top
                    fi.Bounds.X = left
                    fi.Bounds.Width = ri.DataBounds.Right - fi.Bounds.Left
                    fi.Bounds.Height = GroupFooterHeight
                    top += fi.Bounds.Height

                    If Not ri.IndicatorRect.IsEmpty Then
                        fi.IndicatorRect = ri.IndicatorRect
                        fi.IndicatorRect.Y = fi.Bounds.Y
                        fi.IndicatorRect.Height = fi.Bounds.Height
                    End If

                    If View.OptionsView.ShowHorizontalLines <> DevExpress.Utils.DefaultBoolean.False Then
                        ri.AddRowLineInfo(fi.Bounds.Left, fi.Bounds.Bottom - 1, fi.Bounds.Width, 1, PaintAppearance.HorzLine)

                        fi.Bounds.Height -= 1
                    End If

                    CalcRowCellsFooterInfo(fi, ri)
                    If DirectCast(View, CustomGridView).CompactStatementRows Then
                        For Each cell As DevExpress.XtraGrid.Drawing.GridFooterCellInfoArgs In fi.Cells
                            Dim bounds = cell.Bounds
                            bounds.Y = fi.Bounds.Y + 1
                            bounds.Height = Math.Max(1, fi.Bounds.Height - 2)
                            Dim columnInfo = If(cell.Column Is Nothing, Nothing, ColumnsInfo(cell.Column))
                            If columnInfo IsNot Nothing Then
                                'Use the same physical geometry as the header.
                                'Scrolled period footers must not paint beneath
                                'the fixed description column.
                                Dim leftEdge = columnInfo.Bounds.Left + 1
                                Dim rightEdge = columnInfo.Bounds.Right - 1
                                If cell.Column.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.None Then
                                    If Not ViewRects.FixedLeft.IsEmpty Then leftEdge = Math.Max(leftEdge, ViewRects.FixedLeft.Right)
                                    If Not ViewRects.FixedRight.IsEmpty Then rightEdge = Math.Min(rightEdge, ViewRects.FixedRight.Left)
                                End If
                                If View.GridControl IsNot Nothing Then rightEdge = Math.Min(rightEdge, View.GridControl.ClientSize.Width)
                                bounds.X = leftEdge
                                bounds.Width = Math.Max(0, rightEdge - leftEdge)
                            End If
                            cell.Bounds = bounds
                        Next
                    End If
                    footerRowHandle = View.GetParentRowHandle(footerRowHandle)
                    startLevel -= 1
                    left -= LevelIndent
                    n += 1
                Loop
            End Sub

            Public Overrides Function GetRowFooterCount(ByVal rowHandle As Integer, ByVal rowVisibleIndex As Integer, ByVal isExpanded As Boolean) As Integer
                Dim initialVisibleFootersCount As Integer = MyBase.GetRowFooterCount(rowHandle, rowVisibleIndex, isExpanded)
                Dim visibleFootersCount As Integer = initialVisibleFootersCount

                Dim footerRowHandle As Integer = rowHandle
                For i As Integer = 0 To initialVisibleFootersCount - 1
                    Dim args As New ShowGroupFooterEventArgs(View.GetRowLevel(footerRowHandle))
                    RaiseShowGroupFooter(args)

                    If Not args.Visible Then
                        visibleFootersCount -= 1
                    End If

                    footerRowHandle = View.GetParentRowHandle(footerRowHandle)
                Next i

                Return visibleFootersCount
            End Function

            Private Sub RaiseShowGroupFooter(ByVal args As ShowGroupFooterEventArgs)
            Dim aView As CustomGridView = TryCast(View, CustomGridView)
            If aView IsNot Nothing Then
                    aView.RaiseViewInfoShowGroupFooter(aView, args)
                End If
            End Sub
        End Class

End Namespace
