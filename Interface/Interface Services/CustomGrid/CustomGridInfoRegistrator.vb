Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Registrator
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Base.ViewInfo
Namespace Abovo.CustomGrid
    Public Class CustomGridInfoRegistrator

        Inherits GridInfoRegistrator
            Public Overrides Function CreateView(ByVal grid As GridControl) As BaseView
            Return New CustomGridView(TryCast(grid, GridControl))

        End Function
            Public Overrides Function CreateViewInfo(ByVal view As BaseView) As BaseViewInfo
            Return New CustomGridViewInfo(CType(view, CustomGridView))

        End Function
        Public Overrides ReadOnly Property ViewName() As String

            Get
                Return "CustomGridView"
            End Get

        End Property

    End Class

End Namespace