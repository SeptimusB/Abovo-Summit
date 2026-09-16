Imports DevExpress.Spreadsheet
Imports Abovo.FileManager
Imports DevExpress.CodeParser

Namespace Abovo
    Public Class CalcEngine

        Public Event CalculationCompleted As EventHandler

        Public ActiveWSs(-1) As DevExpress.Spreadsheet.Worksheet
        Public WBCalcDirty As Boolean = False
        Public WBCalcMinDirty As Boolean = False

        Private ActiveWSCount As Integer = -1
        Private ActiveObjects(-1) As ActiveObject
        Public ActiveObjectCount As Integer = 0
        Private ActiveObjIndex As Integer = -1
        Private DependencySensitiveCalculationCurrent As Boolean = False

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
                If TryFundingDashboard IsNot Nothing Then
                    TryFundingDashboard.RefreshData()
                    Return
                End If

                Dim TryAnalyser As BPIncomeExpenditureAnalyser = TryCast(Obj, BPIncomeExpenditureAnalyser)
                If TryAnalyser IsNot Nothing Then
                    TryAnalyser.RefreshCalculatedData()
                    Return
                End If

                Dim TryAnalyserV2 As BPIncomeExpenditureAnalyserV2 = TryCast(Obj, BPIncomeExpenditureAnalyserV2)
                If TryAnalyserV2 IsNot Nothing Then
                    TryAnalyserV2.RefreshCalculatedData()
                    Return
                End If

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

                    Try
                        ActiveObj.RefreshObjData()
                    Catch ex As Exception
                        Dim ObjectName As String = "Active interface"
                        If ActiveObj.Obj IsNot Nothing Then ObjectName = ActiveObj.Obj.GetType().Name
                        SystemMessageManager.Publish(
                            ModelID,
                            "The workbook recalculated, but '" & ObjectName & "' could not refresh: " & ex.Message,
                            SystemMessageSeverity.Warning,
                            "Calculation refresh",
                            ObjectName)
                    End Try

                End If
            Next

        End Sub
        Sub CalculateWSs()

            'Any ordinary calculation request may follow a workbook edit.  Until
            'that request completes a full workbook pass, a subsequently-bound
            'analyser must assume that cross-sheet cached values may be stale.
            Dim HadDependencySensitiveCalculation As Boolean =
                DependencySensitiveCalculationCurrent
            DependencySensitiveCalculationCurrent = False

            If ActiveObjectCount > 1 Then

                CalcFile()
                'A staged pass refreshes a model that already had its initial
                'recursive calculation, but cannot certify a newly loaded XLSB.
                DependencySensitiveCalculationCurrent =
                    HadDependencySensitiveCalculation
                Exit Sub
            End If

            If ActiveWSCount = -1 Then
                Exit Sub
            End If
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
            RaiseEvent CalculationCompleted(Me, EventArgs.Empty)

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

            'Direct full-calculation callers must not certify the initial XLSB
            'cache. CalculateWSs restores this marker only if Recursive had
            'already established it before the ordinary calculation request.
            DependencySensitiveCalculationCurrent = False

            Dim Workbook As IWorkbook = ExcelModels(ModelID).WB

            If CalMode = 1 Then Workbook.Calculate()
            If CalMode = 2 Then Workbook.CalculateFull()
            If CalMode = 3 Then Workbook.CalculateFullRebuild()

            Dim CalculationService As CustomCalcEngine =
                ExcelModels(ModelID).WBCalculationService
            If CalculationService IsNot Nothing AndAlso
               CalculationService.DontCalcTDBS AndAlso
               Workbook.Options.CalculationEngineType = CalculationEngineType.ChainBased Then
                CalculationService.CalculateDeferredWorksheets(Workbook)
            End If

            'A staged pass alone cannot establish valid initial XLSB caches.
            'Only CalculateDependencySensitiveFile can set the initial marker.

            WBCalcDirty = False
            WBCalcMinDirty = False
            RefreshObjsData()
            RaiseEvent CalculationCompleted(Me, EventArgs.Empty)

        End Sub

        Public Sub CalculateDependencySensitiveFile(Optional ByVal Reason As String = "Unspecified",
                                                    Optional ByVal Force As Boolean = False)

            Dim Workbook As IWorkbook = ExcelModels(ModelID).WB
            Dim PreviousEngine As CalculationEngineType = Workbook.Options.CalculationEngineType

            If Not Force AndAlso DependencySensitiveCalculationCurrent Then
                Exit Sub
            End If

            DependencySensitiveCalculationCurrent = False

            Try
                'A staged ChainBased pass still leaves stale XLSB cached values
                'on first analyser binding.  The recursive whole-workbook pass
                'is required here; later bindings can reuse this result.
                Workbook.Options.CalculationEngineType = CalculationEngineType.Recursive
                Workbook.Calculate()
                WBCalcDirty = False
                WBCalcMinDirty = False
                DependencySensitiveCalculationCurrent = True
            Finally
                Workbook.Options.CalculationEngineType = PreviousEngine
            End Try

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
