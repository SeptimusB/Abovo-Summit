Namespace Abovo
    Public Class ObjectMiddler
        Public Shared Sub MiddleObject(ByVal SourceObject As Control, ByVal ParentObject As Control, Optional ByVal DoHeight As Boolean = False)

            Dim ParWidth As Integer = ParentObject.Width
            SourceObject.Left = (ParWidth - SourceObject.Width) / 2

            If DoHeight Then

                Dim ParHeight As Integer = ParentObject.Height
                SourceObject.Top = (ParHeight - SourceObject.Height) / 2

            End If

        End Sub
    End Class

End Namespace
