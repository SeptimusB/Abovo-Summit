Imports System.Runtime.InteropServices
Imports System.Threading
Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.CalcEngine
Imports DevExpress.CodeParser
Imports DevExpress.Data
Imports DevExpress.Office.Crypto
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraBars.Navigation
Imports DevExpress.XtraEditors.Filtering
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraSpreadsheet
Imports DevExpress.XtraSpreadsheet.Model

Namespace Abovo
    Public Class FileManager
        'Core Properties

        Public Shared Property FileNameFormat As DocumentFormat

        Public Shared ExcelModels() As ExcelModel

        Public Shared ExcelModelCount As Integer
        Public Shared OpenModelCount As Integer
        Public Shared Parent As Object
        'File Objects

        Public Shared ActiveWB As DevExpress.Spreadsheet.Workbook

        Public Shared Sub Initialise(Parent As Object)

            'Initialisethe shared File manager object
            ExcelModelCount = -1
            OpenModelCount = 0

        End Sub
        Public Shared Function GetWorkBook(SetModelID As Integer) As DevExpress.Spreadsheet.IWorkbook

            Return ExcelModels(SetModelID).WB

        End Function
        Public Shared Function IsFileOpen(Path As String) As Boolean

            If IsNothing(ExcelModels) Then Return False

            If ExcelModels.Length = 0 Then Return False

            Dim CheckMod As ExcelModel

            For Each CheckMod In ExcelModels

                If CheckMod Is Nothing Then Continue For

                If CheckMod.WB.Path = Path Then

                    Return True
                    Exit Function

                End If

            Next

            Return False

        End Function
        Public Shared Function GetGroupID(ModelID As Integer, GroupName As String) As Integer

            Dim IntID As Integer = -1

            For Each IntCheck As GroupStructure In ExcelModels(ModelID).WBStructure.GroupStructures

                If IntCheck.GSName = GroupName Then

                    IntID = IntCheck.GSID
                    Exit For

                End If

            Next

            Return IntID

        End Function
        Public Shared Function GetCSName(ModelID As Integer, SearchGSID As Integer, SearchCSID As Integer) As String

            Dim IntID As Integer = -1

            For Each GS As GroupStructure In ExcelModels(ModelID).WBStructure.GroupStructures

                If GS.GSID = SearchGSID Then

                    For Each CSCheck As ChildStructure In GS.ChildStructures
                        If CSCheck.CSID = SearchCSID Then
                            Return CSCheck.CSName
                            Exit For
                        End If
                    Next

                End If

            Next

            Return "Error"

        End Function
        Public Shared Function GetCSID(ModelID As Integer, GSID As Integer, InstanceName As String) As Integer
            Dim CSID As Integer = -1
            Dim IntCheck As GroupStructure = ExcelModels(ModelID).WBStructure.GroupStructures(GSID)
            If IntCheck IsNot Nothing Then
                For Each CSCheck As ChildStructure In IntCheck.ChildStructures
                    If CSCheck.CSName = InstanceName Then
                        CSID = CSCheck.CSID
                        Exit For
                    End If
                Next
            End If
            Return CSID

        End Function





        Public Shared Sub RegisterModelInterface(SetModelID As Integer, FileInterface As FileInstanceInterface)

            ExcelModels(SetModelID).InstanceInterface = FileInterface

        End Sub
        Public Class ExcelModel

            Public ModelSpreadsheetControl As SpreadsheetControl
            Public SSViewer As MainModelViewer
            Public WB As IWorkbook
            Public WBStructure As Abovo_Model_Def
            Public WBStructureManager As StructureManager
            Public WBData As DataManager
            Public WBInterface As InterfaceManager
            Public WBCalcEngine As CalcEngine
            Public WBCalculationService As CustomCalcEngine
            Public WBDataPres As PresentationManager
            Public EventCoordinator As EventManager
            Public ExpendAnalyser As BPIncomeExpenditureAnalyser
            Public IsDirty As Boolean
            Public ModelID As Integer
            Public ChangeManager As ModelChangeManager
            Public TransDBM As TransDBManager
            Public ColourSwatch As Color
            Public SSViewInitialised As Boolean = False
            Public FileName As String
            Public FileInfo As System.IO.FileInfo
            Public InstanceInterface As FileInstanceInterface
            Public RDSM As RDSManager
            Public InterfaceDependencies As InterfaceDependencyManager
            Public TransDBSync As TransactionalDBSynchroniser
            Public TransDBMaterialiser As TransactionDBMaterialiser
            Public WorkbookMigrations As WorkbookMigrationManager
            Public WorkbookStructureRules As WorkbookStructureRuleManager

            Sub New(SetModelID As Integer)

                ModelID = SetModelID
                'WB = New IWorkbook
                ModelSpreadsheetControl = New SpreadsheetControl

                WBStructureManager = New StructureManager(SetModelID)
                WBData = New DataManager(SetModelID)
                WBCalcEngine = New CalcEngine(SetModelID)
                WBInterface = New InterfaceManager(SetModelID)
                WBDataPres = New PresentationManager(SetModelID)
                EventCoordinator = New EventManager(SetModelID, "AbovoBP")
                TransDBM = New TransDBManager(SetModelID)
                RDSM = New RDSManager(SetModelID)
                InterfaceDependencies = New InterfaceDependencyManager(SetModelID)
                TransDBSync = New TransactionalDBSynchroniser(SetModelID)
                TransDBMaterialiser = New TransactionDBMaterialiser(SetModelID)
                WorkbookMigrations = New WorkbookMigrationManager(SetModelID)
                WorkbookStructureRules = New WorkbookStructureRuleManager(SetModelID)

                IsDirty = False
                WB = ModelSpreadsheetControl.Document
                ModelSpreadsheetControl.ActiveWorksheet.ActiveView.ShowGridlines = False

                InitiateWorkbook(SetModelID)

                AddHandler ModelSpreadsheetControl.DocumentSaved, AddressOf ClearDirtyFlag
                AddHandler ModelSpreadsheetControl.DocumentPropertiesChanged, AddressOf SetDirtyFlag
                AddHandler ModelSpreadsheetControl.UnhandledException, AddressOf SSCUnhandledEvent
                AddHandler ModelSpreadsheetControl.ActiveSheetChanged, AddressOf ProcessSheetChange

            End Sub
            Sub ClearDirtyFlag()

                IsDirty = False

            End Sub


            Sub SetDirtyFlag()

                IsDirty = True

            End Sub
            Sub ProcessSheetChange()

                ModelSpreadsheetControl.ActiveWorksheet.ActiveView.ShowGridlines = False

            End Sub
            Public Sub ShowSpreadsheet(Optional ByVal SetSpreadsheet As DevExpress.Spreadsheet.Worksheet = Nothing, Optional ByVal Parent As Object = Nothing)

                If Not SSViewInitialised Then

                    InitialiseSpreadviewer()

                End If

                If SetSpreadsheet IsNot Nothing Then

                    WB.Worksheets.ActiveWorksheet = SetSpreadsheet

                End If

                SSViewer.Show()
                SSViewer.BringToFront()

            End Sub
            Public Sub HideSpreadsheet()

                If Not SSViewInitialised Then

                    InitialiseSpreadviewer()

                End If

                SSViewer.Hide()
                SSViewer.SendToBack()

            End Sub



            Public Sub InitialiseSpreadviewer()

                SSViewer = New MainModelViewer(ModelID)
                SSViewInitialised = True

            End Sub
            Sub InitiateWorkbook(SetModelID As Integer)

                'SetCustomFunctions

                Dim customFunction As New Abovo.PMCostFunction()

                If Not WB.Functions.GlobalCustomFunctions.Contains(customFunction.Name) Then

                    WB.Functions.GlobalCustomFunctions.Add(customFunction)

                End If

                Dim customFunction2 As New Abovo.ResponsiveCostFunction()

                If Not WB.Functions.GlobalCustomFunctions.Contains(customFunction2.Name) Then

                    WB.Functions.GlobalCustomFunctions.Add(customFunction2)

                End If

                WB.Options.CalculationMode = WorkbookCalculationMode.Manual
                WB.Options.CalculationEngineType = CalculationEngineType.ChainBased
                WB.DocumentSettings.Calculation.EnableMultiThreading = True
                WB.DocumentSettings.Calculation.ThreadCount = Environment.ProcessorCount
                WB.DocumentSettings.Calculation.Iterative = False

                ColourSwatch = GetColour(SetModelID)

            End Sub
            Public Sub PostLoadActions()

                WBCalculationService = New CustomCalcEngine(ModelID)

                WBCalculationService.TransDBSheetID = GetSheetID(ModelID, "Transactional DB")

                WBCalculationService.DontCalcTDBS = True

                WB.AddService(GetType(DevExpress.XtraSpreadsheet.Services.ICustomCalculationService), WBCalculationService)

                ChangeManager = New ModelChangeManager(ModelID)

                HistoryManager.Show()
                HistoryManager.Hide()

                If WorkbookMigrations IsNot Nothing Then

                    WorkbookMigrations.ReportMigrationStatus()

#If DEBUG Then
                    'During development the migration is deliberately applied
                    'automatically in Debug builds.  It is idempotent and the
                    'workbook is marked dirty so the migrated XLSB is only
                    'persisted if the user subsequently saves it.
                    Dim MigrationResult As AbovoTransaction =
                        WorkbookMigrations.ApplyPendingMigrations()

                    If MigrationResult.BError Then
                        Debug.Print("Workbook migration ERROR: " &
                                    MigrationResult.StringReturn)
                    Else
                        Debug.Print("Workbook migration: " &
                                    MigrationResult.StringReturn)
                    End If
#End If

                End If

            End Sub
            Public Function GetSheetID(ModelID As Integer, SheetName As String) As Integer

                Dim WS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets(SheetName)

                If WS Is Nothing Then

                    Return -1

                End If

                Return WS.Index

            End Function



            Function GetColour(ColStep As Integer) As Color

                Dim Col As Color

                Select Case ColStep
                    Case 0
                        Col = Color.FromArgb(0, 71, 187)
                    Case 1
                        Col = Color.FromArgb(143, 214, 189)
                    Case 2
                        Col = Color.FromArgb(240, 236, 116)
                    Case 3
                        Col = Color.FromArgb(198, 188, 208)
                    Case 4
                        Col = Color.FromArgb(255, 0, 255)
                    Case 5
                        Col = Color.FromArgb(0, 71, 187)
                    Case 6
                        Col = Color.FromArgb(143, 214, 189)
                    Case 7
                        Col = Color.FromArgb(240, 236, 116)
                    Case 8
                        Col = Color.FromArgb(198, 188, 208)
                    Case 9
                        Col = Color.FromArgb(255, 0, 255)
                    Case 10
                        Col = Color.FromArgb(198, 188, 208)
                    Case 11
                        Col = Color.FromArgb(255, 0, 255)
                    Case Else
                        Col = Color.FromArgb(100 + (ColStep * 3), 100 + (ColStep * 3), 100 + (ColStep * 3))
                End Select

                Return Col

            End Function

            Public Function SetCellValue(WSName As String, Row As Integer, Col As Integer, Value As Object) As Boolean

                Dim WS As DevExpress.Spreadsheet.Worksheet = WB.Worksheets(WSName)

                If WS Is Nothing Then

                    Return False

                End If

                Dim Cell As DevExpress.Spreadsheet.Cell = WS.Cells(Row, Col)

                If Cell Is Nothing Then

                    Return False

                End If

                Cell.Value = Value

                IsDirty = True

                Return True

            End Function
            Public Sub SaveFileAs()

                'need to change this to managed coded dialog?

                ModelSpreadsheetControl.SaveDocumentAs()

            End Sub
            Public Sub SaveFile()

                If IsDirty Then

                    Try

                        ModelSpreadsheetControl.SaveDocument()

                    Catch ex As Exception

                        MsgBox("Sorry, an error occurred while saving. Please check the file is not open in another program and that you have write permissions to the location. Error details: " & ex.Message)

                        GoTo Exiter

                    End Try

                    IsDirty = False

                End If
Exiter:

            End Sub
            Public Function CommitToCloseModel() As AbovoTransaction

                Dim CloseTrans As New AbovoTransaction

                If IsDirty Then

                    Dim response As MsgBoxResult = MsgBox("Save changes to " & FileName & "?", vbYesNoCancel)

                    If response = MsgBoxResult.Cancel Then

                        CloseTrans.StringReturn = "Cancel"
                        Return CloseTrans
                        Exit Function

                    ElseIf response = MsgBoxResult.Yes Then

                        SaveFile()
                        IsDirty = False

                    End If

                End If

                CloseTrans.StringReturn = "Proceed"

                Return CloseTrans

            End Function
            Public Sub CloseModel()

                'WB = Nothing
                ModelSpreadsheetControl.CreateNewDocument()
                ModelSpreadsheetControl.Dispose()
                ModelSpreadsheetControl = Nothing
                WBInterface.CloseInterfaces()
                WBInterface = Nothing
                WBData = Nothing
                WBCalcEngine = Nothing
                WBStructure = Nothing
                WBStructureManager = Nothing
                WBDataPres = Nothing
                If InterfaceDependencies IsNot Nothing Then InterfaceDependencies.Clear()
                InterfaceDependencies = Nothing
                TransDBMaterialiser = Nothing
                WorkbookMigrations = Nothing
                WorkbookStructureRules = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()

            End Sub
            Public Function DoesWorksheetExist(WSName As String) As Boolean

                If ModelSpreadsheetControl.Document Is Nothing Then

                    Return False

                End If

                If WB.Worksheets.Contains(WSName) Then

                    Return True

                Else

                    Return False

                End If

            End Function

            Public Sub ProcessAsAbovoBP()

                Dim wsGA As DevExpress.Spreadsheet.Worksheet = WB.Worksheets("Global Assumptions")
                Dim clCell As DevExpress.Spreadsheet.Cell

                clCell = wsGA.Cells(5, 2)

                WBStructure.CompanyName = CStr(clCell.Value.TextValue)

                clCell = wsGA.Cells(7, 2)

                WBStructure.StartDate = CDate(clCell.Value.DateTimeValue)

            End Sub

        End Class

        Private Shared InternalUnlockPassword As String = "23_t4qhe"
        Public Shared Property UnlockPassword As String

            Get
                Return InternalUnlockPassword
                Exit Property
            End Get

            Set(ByVal NewUnlockPassword As String)
                InternalUnlockPassword = NewUnlockPassword
            End Set

        End Property

        Private Shared InternalBIsSaving As Boolean
        Public Shared ReadOnly Property BIsSaving As Boolean

            Get
                Return InternalBIsSaving
                Exit Property
            End Get

        End Property
        Public Shared Sub CloseModel(ModelID As Integer)

            ExcelModels(ModelID) = Nothing
            OpenModelCount -= 1

        End Sub
        Private InternalCompanyName As String = ""
        Public Property CompanyName As String
            Get
                Return InternalCompanyName
                Exit Property
            End Get
            Set(ByVal NewCompanyName As String)
                InternalCompanyName = NewCompanyName
            End Set
        End Property

        Private Shared internalFixedStockSize As Integer = 20
        Public Shared Property StockSize As Short
            Get
                Return internalFixedStockSize
                Exit Property
            End Get
            Set(ByVal NewStockSize As Short)
                internalFixedStockSize = NewStockSize
            End Set
        End Property

        Private Shared InternalFileState As Integer

        '0 - initialising
        '1 - ready, empty
        '2 - ready, loaded clean
        '3 - ready, dirty
        '4 - saved undirty
        Public Property FileState As Byte

            Get

                Return InternalFileState
                Exit Property

            End Get

            Set(ByVal NewFileState As Byte)

                InternalFileState = NewFileState

            End Set

        End Property

        Private BUnappliedDataInternal As Boolean
        Public Property BUnappliedData As Boolean
            Get
                Return BUnappliedDataInternal
                Exit Property
            End Get
            Set(ByVal NewDataState As Boolean)
                BUnappliedDataInternal = NewDataState
            End Set
        End Property


        Public FileDetails As FileStructure

        Structure FileStructure

            Public intIdentifier As Byte

            Public CompanyName As String
            Public StartDate As Date
            Sub New(intSetIdentifier As Integer)

                intIdentifier = intSetIdentifier

            End Sub

        End Structure





#Region "Open close save validate transactions"
        Public Sub InitialiseFile()

            InternalFileState = 0
            'InitiateWorkbook()
            InternalBIsSaving = False

        End Sub





        Public Sub SaveFile(ModelID)

            Dim TimeStart As Date = Now()
            WriteLog("Starting corethread save of " & ExcelModels(ModelID).FileName)

            InternalBIsSaving = True

            ExcelModels(ModelID).WB.SaveDocument(ExcelModels(ModelID).FileName, DocumentFormat.Xlsm)
            InternalFileState = 2

            InternalBIsSaving = False
            WriteLog("Complete. Time taken: " & (Now() - TimeStart).ToString)
            If IsDev Then MsgBox("Saved. Time taken: " & (Now() - TimeStart).ToString)

        End Sub

        Public Sub SaveFileAs(StrNewFileName As String, DFFFormat As DocumentFormat)

            'ActiveWB.SaveDocumentAsync(StrNewFileName, DFFFormat)
            'InternalFileState = 2
            'StrFileName = StrNewFileName

        End Sub
        Public Sub CloseModel(ModelID As Integer, CallingForm As Form)

            If InternalFileState = 3 Then
                Dim c As New AbovoMessageBox("Current file not saved", MsgBoxStyle.YesNoCancel, CallingForm, "File Not Saved")
                If c.GetResponse = DialogResult.Cancel Then Exit Sub
            End If
            ActiveWB.Dispose()
            ActiveWB = Nothing
            InternalFileState = 1

        End Sub
        Public Shared Function CloseAllModelsFromFMS(Source As FormMainScreen) As Boolean

            Dim SaveCheck As ExcelModel

            For Each SaveCheck In ExcelModels

                If Not IsNothing(SaveCheck) Then

                    If ExcelModels(SaveCheck.ModelID).CommitToCloseModel.StringReturn = "Proceed" Then

                        ExcelModels(SaveCheck.ModelID).CloseModel()
                        FileManager.CloseModel(SaveCheck.ModelID)
                        Source.RemoveModel(SaveCheck.ModelID)

                    Else

                        Return False
                        Exit Function

                    End If

                End If

            Next

            ReDim ExcelModels(-1)
            ExcelModelCount = -1
            OpenModelCount = 0

            ActiveWB = Nothing
            InternalFileState = 1
            Return True

        End Function
        Public Shared Function ValidateOpenFile(ModelToCheck As Integer) As AbovoTransaction

            Dim ObjResponse As New AbovoTransaction

            'Check if the file is a valid Abovo file
            If ExcelModels(ModelToCheck).DoesWorksheetExist("Global Assumptions") = False Then

                ObjResponse.BError = True
                ObjResponse.StrResponseMessage = "Workbook is not a valid Abovo file."
                Return ObjResponse

            Else

                If ExcelModels(ModelToCheck).WB.Worksheets("Global Assumptions").Cells("A8").DisplayText = "Business Plan Start Date" Then

                    ObjResponse.BError = False
                    ObjResponse.StringReturn = "AbovoBP"

                End If

            End If

            Return ObjResponse

        End Function

        Private Shared Async Function LoadBPAsync(FileName As String) As Task(Of Boolean)
            Dim TimeStart As Date = Now()
            WriteLog("Starting async load of " & FileName)
            Await ActiveWB.LoadDocumentAsync(FileName)
            WriteLog("Complete. Time taken: " & (Now() - TimeStart).ToString)
            If IsDev Then MsgBox("Complete. Time taken: " & (Now() - TimeStart).ToString)
            Return True
        End Function
        Private Shared Sub LoadBPCentral(FileName As String)

            Dim TimeStart As Date = Now()
            WriteLog("Starting corethread load of " & FileName)
            ActiveWB.LoadDocument(FileName)
            WriteLog("Complete. Time taken: " & (Now() - TimeStart).ToString)
            'If IsDev Then MsgBox("Complete. Time taken: " & (Now() - TimeStart).ToString)
            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()

        End Sub

        Private Shared Sub SSCUnhandledEvent(ByVal sender As Object, ByVal e As DevExpress.XtraSpreadsheet.SpreadsheetUnhandledExceptionEventArgs)

            Dim My_Exception As Exception
            My_Exception = e.Exception
            e.Handled = True
            MessageBox.Show(My_Exception.Message, "An error has occured with the file")

        End Sub
        Public Shared Function OpenModel(strPath As String, FileInfo As System.IO.FileInfo) As AbovoTransaction

            Dim ObjResponse As New AbovoTransaction

            'CheckNotOpenAlready
            If OpenModelCount > 0 Then

                Dim CheckMod As ExcelModel

                For Each CheckMod In ExcelModels

                    If CheckMod.WB.Path = strPath Then

                        Beep()
                        ObjResponse.BError = True
                        ObjResponse.StrResponseMessage = "Already open"
                        Return ObjResponse
                        Exit Function

                    End If

                Next

            End If

            ExcelModelCount += 1

            ReDim Preserve ExcelModels(ExcelModelCount)

            ExcelModels(ExcelModelCount) = New ExcelModel(ExcelModelCount)

            ExcelModels(ExcelModelCount).FileInfo = FileInfo

            ExcelModels(ExcelModelCount).FileName = strPath

            Dim TimeStart As Date = Now()

            WriteLog("Starting corethread load of " & strPath)

            ExcelModels(ExcelModelCount).ModelSpreadsheetControl.LoadDocument(strPath)

            ExcelModels(ExcelModelCount).PostLoadActions()

            WriteLog("Complete. Time taken: " & (Now() - TimeStart).ToString)

            'ExcelModels(ExcelModelCount).WB = ExcelModels(ExcelModelCount).ModelSpreadsheetControl.Document

            'InitiateWorkbook(ExcelModels(OpenModelCount).WB)
            'If IsDev Then MsgBox("Complete. Time taken: " & (Now() - TimeStart).ToString)

            ExcelModels(ExcelModelCount).ModelSpreadsheetControl.Dock = DockStyle.Fill
            'ExcelModels(ExcelModelCount).SSInterface.Controls.Add(ExcelModels(ExcelModelCount).ModelSpreadsheetControl)

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()

            Dim CheckFile As AbovoTransaction = ValidateOpenFile(ExcelModelCount)


            If CheckFile.BError = False Then

                If CheckFile.StringReturn = "AbovoBP" Then



                    ObjResponse.BError = False
                    ObjResponse.IntReturnCode = 0
                    ObjResponse.StrResponseMessage = "File loaded successfully."
                    ObjResponse.IntegerReturn = ExcelModelCount
                    ObjResponse.StringReturn = "AbovoBP"
                    InternalFileState = 2

                    If ExcelModels(ExcelModelCount).WBStructureManager.CreateStructureFromXML("NeedtoSetToFileXML").BError = False Then

                        ExcelModels(ExcelModelCount).ProcessAsAbovoBP()

                    Else



                    End If







                End If

            Else

                Beep()
                ObjResponse.BError = True
                ObjResponse.StrResponseMessage = "Sorry, this is not a valid Abovo file."
                GoTo Exiter

            End If

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()

            'XML Part(address??????)



            ApplicationConfiguration.ActiveModelID = ExcelModelCount
            ExcelModels(ExcelModelCount).WBCalcEngine.CalcManual()
            ExcelModels(ExcelModelCount).WBCalcEngine.ChainCalc()
            ExcelModels(ExcelModelCount).FileName = strPath
Exiter:

            Return ObjResponse
            Exit Function

Err_Handler_1:

Err_Clean:

            ObjResponse.BError = True
            ObjResponse.IntReturnCode = -1

        End Function

#End Region






    End Class

End Namespace
