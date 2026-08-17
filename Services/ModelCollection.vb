Namespace Abovo

    Public Class ModelCollection

        Public Shared AllModels(-1) As Object
        Public Shared AbovoBusinessPlans(-1) As AbovoBP
        Public Shared AbovoBusinessPlanIDs As Integer = -1

        'Public Shared AbovoDSAs As AbovoDSA()
        'Public Shared AbovoFortresses As AbovoFortress()

        Public Shared ModelCount As Integer = -1

        Public Shared Sub Initialise()

            ' Initialize the collection of business plans
            AbovoBusinessPlans = New AbovoBP(0) {}

            ' Initialize the collection of data source adapters
            ' AbovoDSAs = New AbovoDSA(0) {}

            ' Initialize the collection of fortresses
            ' AbovoFortresses = New AbovoFortress(0) {}

            ModelCount = 0

        End Sub
        Public Shared Function AddModel(ModelType As String) As Integer

            ' Add the model to the collection
            Dim newIndex As Integer = ModelCount
            ReDim Preserve AllModels(newIndex)
            Dim model As Object = Nothing

            If ModelType = "AbovoBP" Then

                model = New AbovoBP()

                AbovoBusinessPlanIDs += 1

                If AbovoBusinessPlans Is Nothing Then

                    ReDim Preserve AbovoBusinessPlans(AbovoBusinessPlanIDs)

                End If

                AbovoBusinessPlans(AbovoBusinessPlanIDs) = model

            ElseIf ModelType = "AbovoDSA" Then
                'etc
            End If

            ModelCount += 1
            AllModels(newIndex) = model



            Return newIndex

        End Function
        Public Shared Sub CloseAll()

            If ModelCount = 0 Then
                ' No models to close, exit early
                Return
            End If

            ' Close all business plans
            If AbovoBusinessPlans IsNot Nothing Then
                For Each bp As AbovoBP In AbovoBusinessPlans
                    If bp IsNot Nothing Then
                        'bp.Close()
                    End If
                Next
            End If

            ' Close all data source adapters
            ' If AbovoDSAs IsNot Nothing Then
            '     For Each dsa As AbovoDSA In AbovoDSAs
            '         If dsa IsNot Nothing Then
            '             dsa.Close()
            '         End If
            '     Next
            ' End If

            ' Close all fortresses
            ' If AbovoFortresses IsNot Nothing Then
            '     For Each fortress As AbovoFortress In AbovoFortresses
            '         If fortress IsNot Nothing Then
            '             fortress.Close()
            '         End If
            '     Next
            ' End If

        End Sub

    End Class

End Namespace
