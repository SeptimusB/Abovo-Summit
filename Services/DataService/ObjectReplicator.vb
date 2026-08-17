Namespace Abovo
    Public Class ObjectReplicator
        Public Function CloneCopy(ObjToClone As Object, Optional ByVal DoShallow As Boolean = False) As Object

            Dim NewObj As Object

            If ObjToClone Is Nothing Then
                Return Nothing
            End If

            Dim ObjType As Type = ObjToClone

            If DoShallow Then

                NewObj = MemberwiseClone(ObjToClone)

            Else

                NewObj = ObjToClone.deepcopy

            End If

            Return NewObj

        End Function

    End Class

End Namespace