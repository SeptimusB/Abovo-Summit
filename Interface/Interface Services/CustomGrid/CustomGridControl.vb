Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Registrator
Imports DevExpress.XtraGrid.Views.Base

Namespace Abovo.CustomGrid
    Public Class CustomGridControl



        Inherits GridControl

        Public Sub New()

            MyBase.New()

        End Sub
        Protected Overrides Function CreateDefaultView() As BaseView

            Return CreateView("CustomGridView")

        End Function
        Protected Overrides Sub RegisterAvailableViewsCore(ByVal collection As InfoCollection)

            MyBase.RegisterAvailableViewsCore(collection)

            collection.Add(New CustomGridInfoRegistrator())

        End Sub

    End Class

End Namespace