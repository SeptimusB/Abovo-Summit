Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo

Namespace Abovo.CustomGrid


    Public Delegate Sub ShowGroupFooterEventHandler(ByVal sender As Object, ByVal e As ShowGroupFooterEventArgs)

        Public Class ShowGroupFooterEventArgs
            Inherits EventArgs

        Private footerLevel_Renamed As Integer
            Private footerVisible As Boolean

            Public Sub New(ByVal aFooterLevel As Integer)
                MyBase.New()
                footerVisible = True
                footerLevel_Renamed = aFooterLevel
            End Sub

            Public ReadOnly Property FooterLevel() As Integer
                Get
                    Return footerLevel_Renamed
                End Get
            End Property

            Public Property Visible() As Boolean
                Get
                    Return footerVisible
                End Get
                Set(ByVal value As Boolean)
                    footerVisible = value
                End Set
            End Property
        End Class



End Namespace