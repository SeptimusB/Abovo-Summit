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
        Private DependencyGraphPrepared As Boolean = False
        Private NavigationMutationGeneration As Long = 0
        Private NavigationCalculatedGeneration As Long = -1

        Public ReadOnly Property NavigationCalculationCurrent As Boolean
            Get
                Return System.Threading.Interlocked.Read(NavigationMutationGeneration) =
                       System.Threading.Interlocked.Read(NavigationCalculatedGeneration)
            End Get
        End Property

        Public ReadOnly Property NavigationCalculationGeneration As Long
            Get
                Return System.Threading.Interlocked.Read(NavigationMutationGeneration)
            End Get
        End Property

        Public Sub MarkPotentialWorkbookChange()
            System.Threading.Interlocked.Increment(NavigationMutationGeneration)
        End Sub

        Private Sub MarkNavigationCalculationCurrentIfUnchanged(ByVal requestedGeneration As Long)
            If Not ModelSafetyManager.IsBulkWorkbookMutationInProgress(ModelID) AndAlso
               System.Threading.Interlocked.Read(NavigationMutationGeneration) = requestedGeneration Then
                System.Threading.Interlocked.Exchange(
                    NavigationCalculatedGeneration, requestedGeneration)
            End If
        End Sub

        Public ReadOnly Property ActiveWorksheetRegistrationCount As Integer
            Get
                Dim count As Integer = 0
                For Each worksheet As Worksheet In ActiveWSs
                    If worksheet IsNot Nothing Then count += 1
                Next
                Return count
            End Get
        End Property

        Public ReadOnly Property ActiveObjectSlotCount As Integer
            Get
                Return ActiveObjects.Length
            End Get
        End Property

        Public ReadOnly Property ActiveWorksheetSlotCount As Integer
            Get
                Return ActiveWSs.Length
            End Get
        End Property

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
        Sub CalculateWSs(Optional ByVal InvalidateNavigation As Boolean = True,
                         Optional ByVal MetricContext As String = Nothing)

            'This is also called for edit, undo and redo. A single active
            'worksheet pass does not certify cross-sheet/deferred results.
            'Interface registration also calls this method to populate its
            'initial worksheet, but registration is not a workbook mutation.
            Dim timer As System.Diagnostics.Stopwatch =
                If(String.IsNullOrEmpty(MetricContext), Nothing,
                   System.Diagnostics.Stopwatch.StartNew())
            If InvalidateNavigation Then MarkPotentialWorkbookChange()
            If ActiveObjectCount > 1 Then

                CalcFile(1, MetricContext)
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

            Dim worksheetMs As Long = If(timer Is Nothing, 0, timer.ElapsedMilliseconds)

            WBCalcMinDirty = False

            RefreshObjsData()
            RaiseEvent CalculationCompleted(Me, EventArgs.Empty)
            If timer IsNot Nothing Then
                System.Diagnostics.Trace.WriteLine(
                    "[Population Benchmark] CalculateWSs: " & MetricContext &
                    ", worksheetCalc=" & worksheetMs.ToString() & " ms" &
                    ", refreshAndEvents=" &
                    (timer.ElapsedMilliseconds - worksheetMs).ToString() & " ms" &
                    ", total=" & timer.ElapsedMilliseconds.ToString() & " ms" &
                    ", activeObjects=" & ActiveObjectCount.ToString() &
                    ", activeWorksheets=" & ActiveWorksheetRegistrationCount.ToString())
            End If

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

            Dim objectIndex As Integer = -1
            For index As Integer = 0 To ActiveObjIndex
                If ActiveObjects(index) Is Nothing Then
                    objectIndex = index
                    Exit For
                End If
            Next
            If objectIndex = -1 Then
                ActiveObjIndex += 1
                ReDim Preserve ActiveObjects(ActiveObjIndex)
                objectIndex = ActiveObjIndex
            End If

            ActiveObjects(objectIndex) = New ActiveObject(Pusher)
            ActiveObjects(objectIndex).ObjectID = objectIndex
            ActiveObjectCount += 1
            Return objectIndex

        End Function

        Public Sub RemoveActiveObject(ObjectID As Integer)

            If ObjectID < 0 OrElse ObjectID >= ActiveObjects.Length Then Return

            If ActiveObjects(ObjectID) IsNot Nothing Then

                If ActiveObjects(ObjectID).TaggedWorksheets IsNot Nothing Then
                    For Each tws In ActiveObjects(ObjectID).TaggedWorksheets
                        If tws.WSID >= 0 AndAlso tws.WSID < ActiveWSs.Length Then
                            ActiveWSs(tws.WSID) = Nothing
                        End If
                    Next
                End If

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
                End If
            Next

        End Sub
        Public Function AddActiveWorksheet(ObjectID As Integer,
                                           ws As DevExpress.Spreadsheet.Worksheet,
                                           Optional ByVal CalculateNow As Boolean = True) As Integer

            If ObjectID >= 0 AndAlso ObjectID < ActiveObjects.Length AndAlso
               ActiveObjects(ObjectID) IsNot Nothing AndAlso ws IsNot Nothing Then

                Dim activeObject As ActiveObject = ActiveObjects(ObjectID)
                If activeObject.TaggedWorksheets IsNot Nothing Then
                    For Each taggedWorksheet In activeObject.TaggedWorksheets
                        If Object.ReferenceEquals(taggedWorksheet.WS, ws) Then
                            Return taggedWorksheet.WSID
                        End If
                    Next
                End If

                Dim worksheetIndex As Integer = -1
                For index As Integer = 0 To ActiveWSCount
                    If ActiveWSs(index) Is Nothing Then
                        worksheetIndex = index
                        Exit For
                    End If
                Next
                If worksheetIndex = -1 Then
                    ActiveWSCount += 1
                    ReDim Preserve ActiveWSs(ActiveWSCount)
                    worksheetIndex = ActiveWSCount
                End If
                ActiveWSs(worksheetIndex) = ws
                activeObject.AddWorksheet(ws, worksheetIndex)

                If CalculateNow Then CalculateWSs(False)

                Return worksheetIndex

            Else

                Return -1

            End If

        End Function
        Public Sub CalcFile(Optional ByVal CalMode As Byte = 1,
                            Optional ByVal MetricContext As String = Nothing)

            'If FileManager.BIsSaving Then Exit Sub

            Dim Workbook As IWorkbook = ExcelModels(ModelID).WB
            Dim timer As System.Diagnostics.Stopwatch =
                If(String.IsNullOrEmpty(MetricContext), Nothing,
                   System.Diagnostics.Stopwatch.StartNew())
            Dim ordinaryMs As Long = 0
            Dim deferredMs As Long = 0
            Dim refreshMs As Long = 0
            Dim stage As String = "workbook"
            Dim succeeded As Boolean = False
            Dim requestedGeneration As Long =
                System.Threading.Interlocked.Read(NavigationMutationGeneration)

            Try
                If CalMode = 1 Then Workbook.Calculate()
                If CalMode = 2 Then Workbook.CalculateFull()
                If CalMode = 3 Then Workbook.CalculateFullRebuild()
                If timer IsNot Nothing Then ordinaryMs = timer.ElapsedMilliseconds

                stage = "deferred worksheets"
                Dim CalculationService As CustomCalcEngine =
                    ExcelModels(ModelID).WBCalculationService
                If CalculationService IsNot Nothing AndAlso
                   CalculationService.DontCalcTDBS AndAlso
                   Workbook.Options.CalculationEngineType = CalculationEngineType.ChainBased Then
                    CalculationService.CalculateDeferredWorksheets(Workbook)
                End If
                If timer IsNot Nothing Then deferredMs = timer.ElapsedMilliseconds - ordinaryMs

                stage = "interface refresh"
                WBCalcDirty = False
                WBCalcMinDirty = False
                RefreshObjsData()
                If timer IsNot Nothing Then
                    refreshMs = timer.ElapsedMilliseconds - ordinaryMs - deferredMs
                End If

                stage = "completion events"
                RaiseEvent CalculationCompleted(Me, EventArgs.Empty)
                If CalMode >= 1 AndAlso CalMode <= 3 Then
                    MarkNavigationCalculationCurrentIfUnchanged(requestedGeneration)
                End If
                succeeded = True
            Finally
                If Not succeeded Then MarkPotentialWorkbookChange()
                If timer IsNot Nothing Then
                    Dim benchmarkPrefix As String =
                        If(MetricContext.StartsWith("DIT ", StringComparison.Ordinal),
                           "[Navigation Benchmark]", "[Population Benchmark]")
                    System.Diagnostics.Trace.WriteLine(
                        benchmarkPrefix & " CalcFile: " & MetricContext &
                        ", ordinary=" & ordinaryMs.ToString() & " ms" &
                        ", deferred=" & deferredMs.ToString() & " ms" &
                        ", interfaces=" & refreshMs.ToString() & " ms" &
                        ", total=" & timer.ElapsedMilliseconds.ToString() & " ms" &
                        ", activeObjects=" & ActiveObjectCount.ToString() &
                        ", activeWorksheets=" & ActiveWorksheetRegistrationCount.ToString() &
                        ", generation=" & NavigationCalculationGeneration.ToString() &
                        ", navigationCurrent=" & NavigationCalculationCurrent.ToString() &
                        ", outcome=" & If(succeeded, "ok", "failed at " & stage))
                End If
            End Try

        End Sub

        Public Sub CalculateDependencySensitiveFile(Optional ByVal Reason As String = "Unspecified",
                                                    Optional ByVal Force As Boolean = False)

            Dim Workbook As IWorkbook = ExcelModels(ModelID).WB
            Dim requestedGeneration As Long =
                System.Threading.Interlocked.Read(NavigationMutationGeneration)
            If Not Force AndAlso DependencyGraphPrepared Then
                'The analyser may have been hidden while a single DIT worksheet
                'was calculated. Bring all dirty dependencies and the deferred
                'Transactional DB sheets current before binding it again.
                Workbook.Calculate()
                Dim DeferredService As CustomCalcEngine =
                    ExcelModels(ModelID).WBCalculationService
                If DeferredService IsNot Nothing AndAlso
                   DeferredService.DontCalcTDBS AndAlso
                   Workbook.Options.CalculationEngineType = CalculationEngineType.ChainBased Then
                    DeferredService.CalculateDeferredWorksheets(Workbook)
                End If
                WBCalcDirty = False
                WBCalcMinDirty = False
                MarkNavigationCalculationCurrentIfUnchanged(requestedGeneration)
                Return
            End If

            Dim PreviousEngine As CalculationEngineType = Workbook.Options.CalculationEngineType
            Dim CalculationService As CustomCalcEngine =
                ExcelModels(ModelID).WBCalculationService
            Dim PreviousSkipTransactionalDB As Boolean =
                If(CalculationService Is Nothing, False, CalculationService.DontCalcTDBS)

            DependencyGraphPrepared = False

            Try
                'The imported XLSB needs its calculation chain rebuilt once.
                'Afterward ordinary edits can recalculate their dependents
                'incrementally, including the separately deferred worksheets.
                If CalculationService IsNot Nothing Then
                    CalculationService.DontCalcTDBS = False
                End If
                Workbook.Options.CalculationEngineType = CalculationEngineType.ChainBased
                Workbook.CalculateFullRebuild()
                WBCalcDirty = False
                WBCalcMinDirty = False
                DependencyGraphPrepared = True
            Finally
                Try
                    Workbook.Options.CalculationEngineType = PreviousEngine
                Finally
                    If CalculationService IsNot Nothing Then
                        CalculationService.DontCalcTDBS = PreviousSkipTransactionalDB
                    End If
                End Try
            End Try

            MarkNavigationCalculationCurrentIfUnchanged(requestedGeneration)

        End Sub

        Public Sub InvalidateDependencyGraph()
            DependencyGraphPrepared = False
            MarkPotentialWorkbookChange()
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
