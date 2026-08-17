Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Abovo.CustomGrid
Imports DevExpress.XtraExport.Helpers
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.BandedGrid
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Namespace Abovo.CustomGrid
    Public Class CustomGridView

        Inherits GridView
        Public Sub New(ByVal ownerGrid As GridControl)

            MyBase.New(ownerGrid)
        End Sub
        Public Sub New()
            MyBase.New()
        End Sub

        Friend Sub RaiseViewInfoShowGroupFooter(ByVal sender As Object, ByVal e As ShowGroupFooterEventArgs)
            RaiseEvent ShowGroupFooter(sender, e)
        End Sub
        Protected Overrides ReadOnly Property ViewName() As String
            Get
                Return "CustomGridView"
            End Get
        End Property
        Protected Overrides Function GetShowGroupedColumns() As Boolean

            Return AllowPartialGroups OrElse OptionsView.ShowGroupedColumns
        End Function
        Protected Overrides Sub PreProcessVisibleColumnsList(ByVal tempList As List(Of GridColumn))
            MyBase.PreProcessVisibleColumnsList(tempList)   '.PreProcessVisibleColumnsList(tempList)

            If tempList.Count = 0 Then Return

            For i As Integer = tempList.Count - 1 To 0 Step -1

                If (TryCast(tempList(i), GridColumn)).GroupIndex > -1 Then

                    tempList.RemoveAt(i)
                End If
            Next
        End Sub

        Public Event ShowGroupFooter As ShowGroupFooterEventHandler



    End Class

    Public Class CustomBandedGridView

        Inherits BandedGridView
        Public Sub New(ByVal ownerGrid As GridControl)

            MyBase.New(ownerGrid)

        End Sub
        Public Sub New()
            MyBase.New()
        End Sub

        Friend Sub RaiseViewInfoShowGroupFooter(ByVal sender As Object, ByVal e As ShowGroupFooterEventArgs)

            RaiseEvent ShowGroupFooter(sender, e)

        End Sub
        Protected Overrides ReadOnly Property ViewName() As String

            Get
                Return "CustomBandedGridView"

            End Get
        End Property
        Protected Overrides Function GetShowGroupedColumns() As Boolean

            Return AllowPartialGroups OrElse OptionsView.ShowGroupedColumns

        End Function
        Protected Overrides Sub PreProcessVisibleColumnsList(ByVal tempList As List(Of GridColumn))
            MyBase.PreProcessVisibleColumnsList(tempList)   '.PreProcessVisibleColumnsList(tempList)

            If tempList.Count = 0 Then Return

            For i As Integer = tempList.Count - 1 To 0 Step -1

                If (TryCast(tempList(i), GridColumn)).GroupIndex > -1 Then

                    tempList.RemoveAt(i)
                End If
            Next
        End Sub

        Public Event ShowGroupFooter As ShowGroupFooterEventHandler



    End Class

    Public Class MyGridViewInfo
        Inherits CustomGridViewInfo

        Public Sub New(ByVal gridView As CustomGridView)
            MyBase.New(gridView)
        End Sub

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