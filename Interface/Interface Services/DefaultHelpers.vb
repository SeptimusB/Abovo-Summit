Imports Abovo
Imports Abovo.DataObject
Imports Abovo.GeneralFunctions
Imports DevExpress.Utils
Imports DevExpress.XtraGrid.Views.Grid

Namespace Abovo
    Public Class DefaultHelpers

        Public Shared Sub DefaultCustomDrawColumnHeader(ByVal sender As Object, ByVal e As ColumnHeaderCustomDrawEventArgs)

            If e.Column Is Nothing Then

                e.Handled = True
                Return

            End If

            Dim ColTag As DataColumnTag = e.Column.Tag

            If ColTag.DontDrawCellHeader Then
                e.Handled = True
                Return
            End If

            If ColTag.HasControls Then Return

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

    End Class

End Namespace
