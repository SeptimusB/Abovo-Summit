Namespace Abovo
    Public Class ActionManager

        Public ActionSets() As DataAction
        Sub New(SetModelID As Integer, SetGSID As Integer, SetCSID As Integer, MyParent As Form)

            Dim ModelID As Integer = SetModelID
            Dim GSID As Integer = SetGSID
            Dim CSID As Integer = SetCSID
            Dim ParentForm As Form = MyParent

        End Sub


        Class DataAction

        End Class

    End Class
    Class ActionTag

        Public ID As Integer
        Public Parent As Form
        Public ActionType As String
        Public ActionDataStr1 As String
        Public ActionDataStr2 As String
        Public ActionDataStr3 As String

    End Class
End Namespace