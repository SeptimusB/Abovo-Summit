Imports System.Drawing
Imports DevExpress.XtraGrid.Views.Base

Public Module CellDrawHelper
    Public Sub DrawCellBorder(ByVal e As RowCellCustomDrawEventArgs,
                              ByVal colour As Color)
        Dim frameBounds As New Rectangle(e.Bounds.X, e.Bounds.Y,
                                         e.Bounds.Width + 1, e.Bounds.Height + 1)
        Dim penWidth As Integer = 3 * CInt(Math.Ceiling(e.Cache.ScaleDPI.ScaleFactorHorz - 0.5))
        Dim borderPen As Pen = e.Cache.GetPen(colour, penWidth)
        e.Cache.DrawRectangle(borderPen, frameBounds)
    End Sub
End Module
