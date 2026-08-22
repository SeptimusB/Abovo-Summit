Imports DevExpress.Spreadsheet
Imports Abovo.FileManager
Imports DevExpress.CodeParser

Namespace Abovo
    Public Class CalcEngine

        Public ActiveWSs(-1) As DevExpress.Spreadsheet.Worksheet
        Public WBCalcDirty As Boolean = False
        Public WBCalcMinDirty As Boolean = False

        Private ActiveWSCount As Integer = -1
        Private ActiveObjects(-1) As ActiveObject
        Public ActiveObjectCount As Integer = 0
        Private ActiveObjIndex As Integer = -1

#Region "Calclulation and engine"
        Public ModelID As Integer
        Sub New(SetModelID As Integer)

            ModelID = SetModelID

        End Sub
        Class ActiveObject

            Public ObjectID As Integer
            Public Obj As Object
            Public TaggedWorksheets() As TaggedWS
            Public WSCount As Integer = -1

            Sub New(ByVal ActiveObject As Object)

                Obj = ActiveObject

            End Sub
            Sub RefreshObjData()

                Dim TryObj As DataInterfaceTemplate

                TryObj = TryCast(Obj, DataInterfaceTemplate)

                If TryObj IsNot Nothing Then

                    TryObj.RefreshData()

                    Return

                End If

                Dim TryDashboard As BPDashboard = TryCast(Obj, BPDashboard)
                If TryDashboard IsNot Nothing Then
                    TryDashboard.RefreshData()
                    Return
                End If

                Dim TryFundingDashboard As FundingDashboard = TryCast(Obj, FundingDashboard)
                If TryFundingDashboard IsNot Nothing Then TryFundingDashboard.RefreshData()

            End Sub
            Public Sub AddWorksheet(ByVal ws As DevExpress.Spreadsheet.Worksheet, SetID As Integer)

                If WSCount > -1 Then
                    For Each tws In TaggedWorksheets
                        If tws.WS.Name = ws.Name Then
                            'Worksheet already exists, return existing index
                            Return
                        End If
                    Next
                End If
                WSCount += 1
                ReDim Preserve TaggedWorksheets(WSCount)
                Dim NewTaggedWS As New TaggedWS With {
                    .WSID = SetID,
                    .WS = ws
                }

                TaggedWorksheets(WSCount) = NewTaggedWS

            End Sub

            Structure TaggedWS

                Public WSID As Integer
                Public WS As Worksheet

            End Structure

        End Class
        Sub RefreshObjsData()

            If ActiveObjIndex = -1 Then Exit Sub
            If ActiveObjectCount = 0 Then Exit Sub

            For Each ActiveObj In ActiveObjects

                If ActiveObj IsNot Nothing Then

                    ActiveObj.RefreshObjData()

                End If
            Next

        End Sub
        Sub CalculateWSs()

            If ActiveObjectCount > 1 Then

                CalcFile()

            End If

            If ActiveWSCount = -1 Then Exit Sub
                Dim CalcedWSsCount As Integer = -1
                Dim CalcedWSs(-1) As DevExpress.Spreadsheet.Worksheet

                For Each ws As DevExpress.Spreadsheet.Worksheet In ActiveWSs

                    If Not ws Is Nothing Then

                        If CalcedWSsCount > -1 Then

                            For Each cws In CalcedWSs

                                If cws.Name = ws.Name Then GoTo NextWS

                            Next

                            ws.Calculate()
                            CalcedWSsCount += 1
                            ReDim Preserve CalcedWSs(CalcedWSsCount)
                            CalcedWSs(CalcedWSsCount) = ws

                        Else

                            ws.Calculate()
                            CalcedWSsCount += 1
                            ReDim Preserve CalcedWSs(CalcedWSsCount)
                            CalcedWSs(CalcedWSsCount) = ws

                        End If

                    End If

NextWS:

                Next


            WBCalcMinDirty = False

            RefreshObjsData()

        End Sub

        Public Function AddActiveObject(ByVal Pusher As Object) As Integer

            If ActiveObjIndex > -1 Then

                For Each ActiveObj In ActiveObjects

                    If ActiveObj IsNot Nothing Then

                        If ActiveObj.Obj Is Pusher Then

                            'Object already exists, return existing index
                            Return ActiveObj.ObjectID

                        End If

                    End If

                Next

            End If

            'add the object
            ActiveObjIndex += 1

            ReDim Preserve ActiveObjects(ActiveObjIndex)

            ActiveObjects(ActiveObjIndex) = New ActiveObject(Pusher)
            ActiveObjects(ActiveObjIndex).ObjectID = ActiveObjIndex
            ActiveObjectCount += 1
            Return ActiveObjIndex

        End Function

        Public Sub RemoveActiveObject(ObjectID As Integer)

            If ObjectID < 0 OrElse ObjectID >= ActiveObjects.Length Then Return

            If ActiveObjects(ObjectID) IsNot Nothing Then

                For Each tws In ActiveObjects(ObjectID).TaggedWorksheets

                    ActiveWSs(tws.WSID) = Nothing

                Next

                ActiveObjects(ObjectID) = Nothing

                ActiveObjectCount -= 1

            End If

        End Sub

        Public Sub RemoveActiveObject(ByVal pusher As Object)

            If pusher Is Nothing OrElse ActiveObjIndex < 0 Then Return

            For objectIndex As Integer = 0 To ActiveObjects.Length - 1
                Dim activeObject As ActiveObject = ActiveObjects(objectIndex)

                If activeObject IsNot Nothing AndAlso
                   Object.ReferenceEquals(activeObject.Obj, pusher) Then

                    RemoveActiveObject(objectIndex)
                    Return
                End If
            Next

        End Sub
        Public Function AddActiveWorksheet(ObjectID As Integer,
                                           ws As DevExpress.Spreadsheet.Worksheet,
                                           Optional ByVal CalculateNow As Boolean = True) As Integer

            If ActiveObjects(ObjectID) IsNot Nothing Then

                ActiveWSCount += 1
                ReDim Preserve ActiveWSs(ActiveWSCount)
                ActiveWSs(ActiveWSCount) = ws
                ActiveObjects(ObjectID).AddWorksheet(ws, ActiveWSCount)

                If CalculateNow Then CalculateWSs()

                Return ActiveWSCount

            Else

                Return -1

            End If

        End Function
        Public Sub CalcFile(Optional ByVal CalMode As Byte = 1)

            'If FileManager.BIsSaving Then Exit Sub

            If CalMode = 1 Then ExcelModels(ModelID).WB.Calculate()
            If CalMode = 2 Then ExcelModels(ModelID).WB.CalculateFull()
            If CalMode = 3 Then ExcelModels(ModelID).WB.CalculateFullRebuild()

            WBCalcDirty = False
            WBCalcMinDirty = False
            RefreshObjsData()

        End Sub
        Public Sub CalcManual()

            ExcelModels(ModelID).WB.DocumentSettings.Calculation.Mode = CalculationMode.Manual

        End Sub
        Public Sub CalcAuto()

            ExcelModels(ModelID).WB.DocumentSettings.Calculation.Mode = CalculationMode.Automatic

        End Sub
        Public Sub ChainCalc()

            ExcelModels(ModelID).WB.Options.CalculationEngineType = CalculationEngineType.ChainBased

        End Sub
        Public Sub RecursCalc()

            ExcelModels(ModelID).WB.Options.CalculationEngineType = CalculationEngineType.Recursive

        End Sub
#End Region

    End Class

End Namespace
