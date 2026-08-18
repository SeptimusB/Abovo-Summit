Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraEditors
Namespace Abovo
    Public Class GridFormatter

        Sub FormatGrid(PassedGrid As DevExpress.XtraGrid.GridControl)
            'PassedGrid.Font = getfon
        End Sub

        Public Shared Sub ClearBlankHeaders(PassedGrid As DevExpress.XtraGrid.GridControl)

            Dim GV As GridView = PassedGrid.MainView

            Dim GVCol As GridColumn
            Dim x As Integer = 0
            For Each GVCol In GV.Columns

                If Len(GVCol.Caption) > 4 Then

                    If Left(GVCol.Caption, 5) = "Blank" Then GVCol.Caption = " "

                End If

                x = x + 1

            Next



        End Sub

    End Class

End Namespace

