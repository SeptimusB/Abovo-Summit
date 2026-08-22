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
        Public Enum WorkbookOpenMode
            FullModel
            ImportSource
        End Enum

        'Core Properties

        Public Shared Property FileNameFormat As DocumentFormat

        Public Shared ExcelModels() As ExcelModel

        Public Shared ExcelModelCount As Integer
        Public Shared OpenModelCount As Integer
        Public Shared Parent As Object
        'File Objects

        Public Shared ActiveWB As DevExpress.Spreadsheet.Workbook

        Public Shared Sub Initialise(ByVal SetParent As Object)

            Parent = SetParent
            ExcelModels = Nothing
            ExcelModelCount = -1
            OpenModelCount = 0
            InternalFileState = 1
            InternalBIsSaving = False

        End Sub
        Public Shared Function GetWorkBook(SetModelID As Integer) As DevExpress.Spreadsheet.IWorkbook

            If ExcelModels Is Nothing OrElse
               SetModelID < 0 OrElse
               SetModelID >= ExcelModels.Length OrElse
               ExcelModels(SetModelID) Is Nothing Then Return Nothing

            Return ExcelModels(SetModelID).WB

        End Function
        Public Shared Function IsFileOpen(ByVal Path As String) As Boolean

            If String.IsNullOrWhiteSpace(Path) OrElse ExcelModels Is Nothing Then
                Return False
            End If

            If ExcelModels.Length = 0 Then Return False

            Dim CandidatePath As String

            Try
                CandidatePath = System.IO.Path.GetFullPath(Path)
            Catch
                Return False
            End Try

            For Each CheckMod As ExcelModel In ExcelModels

                If CheckMod Is Nothing Then Continue For

                Dim OpenPath As String = CheckMod.FileName
                If String.IsNullOrWhiteSpace(OpenPath) AndAlso
                   CheckMod.WB IsNot Nothing Then OpenPath = CheckMod.WB.Path

                If Not String.IsNullOrWhiteSpace(OpenPath) AndAlso
                   String.Equals(System.IO.Path.GetFullPath(OpenPath),
                                 CandidatePath,
                                 StringComparison.OrdinalIgnoreCase) Then

                    Return True
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
            Public ExpendAnalyserV2 As BPIncomeExpenditureAnalyserV2
            Public ReadOnly ResourceRegistry As New ModelResourceRegistry
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

            Sub New(
                SetModelID As Integer,
                Optional OpenMode As WorkbookOpenMode = WorkbookOpenMode.FullModel)

                ModelID = SetModelID
                'WB = New IWorkbook
                ModelSpreadsheetControl = New SpreadsheetControl

                If OpenMode = WorkbookOpenMode.FullModel Then
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
                End If

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
            Public Function PostLoadActions() As AbovoTransaction

                Dim Result As New AbovoTransaction

                Try
                    Dim TransactionSheetID As Integer =
                        GetSheetID(ModelID, "Transactional DB")

                    If TransactionSheetID < 0 Then
                        Throw New InvalidOperationException(
                            "The workbook is missing the 'Transactional DB' worksheet.")
                    End If

                    WBCalculationService = New CustomCalcEngine(ModelID) With {
                        .TransDBSheetID = TransactionSheetID,
                        .DontCalcTDBS = True
                    }

                    WB.AddService(
                        GetType(DevExpress.XtraSpreadsheet.Services.ICustomCalculationService),
                        WBCalculationService)

                    ChangeManager = New ModelChangeManager(ModelID)

                    HistoryManager.Show()
                    HistoryManager.Hide()

                    If WorkbookMigrations IsNot Nothing Then
                        'Every full-model load reconciles the workbook to the
                        'current schema in memory. This must also run in Release
                        'so an XLSB can round-trip through Summit and Excel
                        'without leaving production users on an older schema.
                        'Persistence still requires an explicit user save.
                        Dim MigrationResult As AbovoTransaction =
                            WorkbookMigrations.ApplyPendingMigrations()

                        If MigrationResult.BError Then
                            Return MigrationResult
                        End If
                    End If

                    Result.BSuccess = True
                    Result.StringReturn = "Workbook services initialized."
                    Result.StrResponseMessage = Result.StringReturn

                Catch ex As Exception
                    Result.BError = True
                    Result.IntReturnCode = -1
                    Result.StringReturn =
                        "Workbook services could not be initialized: " & ex.Message
                    Result.StrResponseMessage = Result.StringReturn
                End Try

                Return Result

            End Function
            Public Function GetSheetID(ModelID As Integer, SheetName As String) As Integer

                Dim TargetWorkbook As IWorkbook = GetWorkBook(ModelID)
                If TargetWorkbook Is Nothing OrElse
                   String.IsNullOrWhiteSpace(SheetName) OrElse
                   Not TargetWorkbook.Worksheets.Contains(SheetName) Then Return -1

                Return TargetWorkbook.Worksheets(SheetName).Index

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
            Public Function SaveFileAs() As Boolean

                Dim OriginalPath As String = FileName

                Try
                    ModelSpreadsheetControl.SaveDocumentAs()

                    Dim SavedPath As String = WB.Path
                    If String.IsNullOrWhiteSpace(SavedPath) OrElse
                       String.Equals(SavedPath,
                                     OriginalPath,
                                     StringComparison.OrdinalIgnoreCase) Then Return False

                    FileName = System.IO.Path.GetFullPath(SavedPath)
                    FileInfo = New System.IO.FileInfo(FileName)
                    IsDirty = False
                    Return True

                Catch ex As Exception
                    MessageBox.Show(
                        "The model could not be saved to the selected location." &
                        Environment.NewLine & Environment.NewLine & ex.Message,
                        "Save Abovo Model As",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
                    Return False
                End Try

            End Function
            Public Function SaveFile() As Boolean

                If Not IsDirty Then Return True

                Try
                    ModelSpreadsheetControl.SaveDocument()
                    IsDirty = False
                    Return True

                Catch ex As Exception
                    MessageBox.Show(
                        "Sorry, an error occurred while saving. Please check the file is not open in another program and that you have write permissions to the location." &
                        Environment.NewLine & Environment.NewLine & ex.Message,
                        "Save Abovo Model",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
                    Return False
                End Try

            End Function
            Public Function CommitToCloseModel() As AbovoTransaction

                Dim CloseTrans As New AbovoTransaction

                If IsDirty Then

                    Dim response As MsgBoxResult = MsgBox("Save changes to " & FileName & "?", vbYesNoCancel)

                    If response = MsgBoxResult.Cancel Then

                        CloseTrans.StringReturn = "Cancel"
                        Return CloseTrans
                        Exit Function

                    ElseIf response = MsgBoxResult.Yes Then

                        If Not SaveFile() Then
                            CloseTrans.StringReturn = "Cancel"
                            Return CloseTrans
                        End If

                    End If

                End If

                CloseTrans.StringReturn = "Proceed"

                Return CloseTrans

            End Function
            Public Sub CloseModel()

                ResourceRegistry.ReleaseAll()

                If ModelSpreadsheetControl IsNot Nothing Then
                    RemoveHandler ModelSpreadsheetControl.UnhandledException,
                                  AddressOf SSCUnhandledEvent
                End If

                If WBInterface IsNot Nothing Then WBInterface.CloseInterfaces()
                If InterfaceDependencies IsNot Nothing Then InterfaceDependencies.Clear()
                If SSViewer IsNot Nothing Then SSViewer.Dispose()
                If ModelSpreadsheetControl IsNot Nothing Then
                    ModelSpreadsheetControl.Dispose()
                End If

                SSViewer = Nothing
                ModelSpreadsheetControl = Nothing
                WB = Nothing
                WBInterface = Nothing
                WBData = Nothing
                WBCalcEngine = Nothing
                WBCalculationService = Nothing
                WBStructure = Nothing
                WBStructureManager = Nothing
                WBDataPres = Nothing
                EventCoordinator = Nothing
                ChangeManager = Nothing
                TransDBM = Nothing
                RDSM = Nothing
                InterfaceDependencies = Nothing
                TransDBSync = Nothing
                TransDBMaterialiser = Nothing
                WorkbookMigrations = Nothing
                WorkbookStructureRules = Nothing
                InstanceInterface = Nothing

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

            Public Function ProcessAsAbovoBP() As AbovoTransaction

                Dim Result As New AbovoTransaction

                Try
                    If WBStructure Is Nothing Then
                        Throw New InvalidOperationException(
                            "The workbook interface definition has not been loaded.")
                    End If

                    If WB Is Nothing OrElse
                       Not WB.Worksheets.Contains("Global Assumptions") Then
                        Throw New InvalidOperationException(
                            "The workbook is missing the 'Global Assumptions' worksheet.")
                    End If

                    Dim GlobalAssumptions As DevExpress.Spreadsheet.Worksheet =
                        WB.Worksheets("Global Assumptions")

                    WBStructure.CompanyName =
                        GlobalAssumptions.Cells(5, 2).Value.TextValue
                    WBStructure.StartDate =
                        GlobalAssumptions.Cells(7, 2).Value.DateTimeValue.ToString("yyyy-MM-dd")

                    If Not String.IsNullOrWhiteSpace(WBStructure.RejData) Then
                        UnlockPassword = WBStructure.RejData
                    End If

                    Result.BSuccess = True
                    Result.StringReturn = "Workbook metadata loaded."
                    Result.StrResponseMessage = Result.StringReturn

                Catch ex As Exception
                    Result.BError = True
                    Result.IntReturnCode = -1
                    Result.StringReturn =
                        "Workbook metadata could not be loaded: " & ex.Message
                    Result.StrResponseMessage = Result.StringReturn
                End Try

                Return Result

            End Function

        End Class

        Private Shared InternalUnlockPassword As String = String.Empty
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

            If ExcelModels Is Nothing OrElse
               ModelID < 0 OrElse
               ModelID >= ExcelModels.Length OrElse
               ExcelModels(ModelID) Is Nothing Then Return

            Dim ModelToClose As ExcelModel = ExcelModels(ModelID)
            ExcelModels(ModelID) = Nothing

            Try
                ModelToClose.CloseModel()
            Catch ex As Exception
                WriteLog("Error while releasing model " & ModelID.ToString() & ": " & ex.Message)
            End Try

            OpenModelCount = Math.Max(0, OpenModelCount - 1)

            While ExcelModelCount >= 0 AndAlso
                  (ExcelModels Is Nothing OrElse
                   ExcelModelCount >= ExcelModels.Length OrElse
                   ExcelModels(ExcelModelCount) Is Nothing)
                ExcelModelCount -= 1
            End While

            If ExcelModelCount < 0 Then
                ExcelModels = Nothing
            ElseIf ExcelModels.Length > ExcelModelCount + 1 Then
                ReDim Preserve ExcelModels(ExcelModelCount)
            End If

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

            If ExcelModels Is Nothing OrElse
               ModelID < 0 OrElse
               ModelID >= ExcelModels.Length OrElse
               ExcelModels(ModelID) Is Nothing Then
                Throw New ArgumentOutOfRangeException(NameOf(ModelID))
            End If

            Dim TimeStart As Date = Now()
            WriteLog("Starting corethread save of " & ExcelModels(ModelID).FileName)

            InternalBIsSaving = True

            Try
                If ExcelModels(ModelID).SaveFile() Then
                    InternalFileState = 2
                    WriteLog("Complete. Time taken: " & (Now() - TimeStart).ToString)
                    If IsDev Then
                        MsgBox("Saved. Time taken: " & (Now() - TimeStart).ToString)
                    End If
                End If
            Finally
                InternalBIsSaving = False
            End Try

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

            If ExcelModels Is Nothing OrElse ExcelModels.Length = 0 Then
                OpenModelCount = 0
                ExcelModelCount = -1
                ActiveWB = Nothing
                InternalFileState = 1
                Return True
            End If

            Dim SaveCheck As ExcelModel

            For Each SaveCheck In ExcelModels

                If Not IsNothing(SaveCheck) Then

                    If ExcelModels(SaveCheck.ModelID).CommitToCloseModel.StringReturn = "Proceed" Then

                        FileManager.CloseModel(SaveCheck.ModelID)
                        Source.RemoveModel(SaveCheck.ModelID)

                    Else

                        Return False
                        Exit Function

                    End If

                End If

            Next

            ExcelModels = Nothing
            ExcelModelCount = -1
            OpenModelCount = 0

            ActiveWB = Nothing
            InternalFileState = 1
            Return True

        End Function
        Public Shared Function ValidateOpenFile(ModelToCheck As Integer) As AbovoTransaction

            Return WorkbookContractValidator.Validate(GetWorkBook(ModelToCheck))

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
        Public Shared Function OpenModel(
            ByVal strPath As String,
            ByVal FileInfo As System.IO.FileInfo,
            Optional ByVal OpenMode As WorkbookOpenMode = WorkbookOpenMode.FullModel) As AbovoTransaction

            Dim Result As New AbovoTransaction
            Dim PreviousModelCount As Integer = ExcelModelCount
            Dim NewModelID As Integer = PreviousModelCount + 1
            Dim NewModel As ExcelModel = Nothing
            Dim TimeStart As Date = Now()

            Try
                If String.IsNullOrWhiteSpace(strPath) Then
                    Throw New ArgumentException(
                        "A model path is required.",
                        NameOf(strPath))
                End If

                Dim FullPath As String = System.IO.Path.GetFullPath(strPath)

                If Not System.IO.File.Exists(FullPath) Then
                    Throw New System.IO.FileNotFoundException(
                        "The selected model does not exist.",
                        FullPath)
                End If

                If IsFileOpen(FullPath) Then
                    Throw New InvalidOperationException(
                        "The selected model is already open.")
                End If

                WriteLog("Starting core-thread load of " & FullPath)

                ReDim Preserve ExcelModels(NewModelID)
                NewModel = New ExcelModel(NewModelID, OpenMode)
                ExcelModels(NewModelID) = NewModel
                ExcelModelCount = NewModelID

                NewModel.FileInfo =
                    If(FileInfo, New System.IO.FileInfo(FullPath))
                NewModel.FileName = FullPath
                NewModel.ModelSpreadsheetControl.LoadDocument(FullPath)

                Dim LoadedModelType As String = "ImportSource"

                If OpenMode = WorkbookOpenMode.FullModel Then
                    Dim ContractResult As AbovoTransaction =
                        ValidateOpenFile(NewModelID)
                    If ContractResult.BError Then
                        Throw New System.IO.InvalidDataException(
                            ContractResult.StringReturn)
                    End If

                    Dim ServiceResult As AbovoTransaction =
                        NewModel.PostLoadActions()
                    If ServiceResult.BError Then
                        Throw New InvalidOperationException(ServiceResult.StringReturn)
                    End If

                    Dim StructureResult As AbovoTransaction =
                        NewModel.WBStructureManager.CreateStructureFromXML()
                    If StructureResult.BError Then
                        Throw New InvalidOperationException(StructureResult.StringReturn)
                    End If

                    Dim MetadataResult As AbovoTransaction =
                        NewModel.ProcessAsAbovoBP()
                    If MetadataResult.BError Then
                        Throw New InvalidOperationException(MetadataResult.StringReturn)
                    End If

                    NewModel.ModelSpreadsheetControl.Dock = DockStyle.Fill
                    NewModel.WBCalcEngine.CalcManual()
                    NewModel.WBCalcEngine.ChainCalc()
                    LoadedModelType = ContractResult.StringReturn
                End If

                OpenModelCount += 1
                If OpenMode = WorkbookOpenMode.FullModel Then
                    InternalFileState = 2
                    ApplicationConfiguration.ActiveModelID = NewModelID
                End If

                Result.BSuccess = True
                Result.IntReturnCode = 0
                Result.IntegerReturn = NewModelID
                Result.StringReturn = LoadedModelType
                Result.StrResponseMessage = "File loaded successfully."

                WriteLog(
                    "Completed model load in " &
                    (Now() - TimeStart).ToString() &
                    ": " & FullPath)

            Catch ex As Exception
                RollBackFailedOpen(NewModelID, PreviousModelCount, NewModel)

                Result.BError = True
                Result.BSuccess = False
                Result.IntReturnCode = -1
                Result.StringReturn = ex.Message
                Result.StrResponseMessage = ex.Message

                WriteLog("Model load failed: " & ex.Message, strPath)
            End Try

            Return Result

        End Function

        Private Shared Sub RollBackFailedOpen(
            ByVal ModelID As Integer,
            ByVal PreviousModelCount As Integer,
            ByVal Model As ExcelModel)

            Try
                If Model IsNot Nothing Then Model.CloseModel()
            Catch
                'Preserve the original load failure.
            End Try

            If ExcelModels IsNot Nothing AndAlso
               ModelID >= 0 AndAlso
               ModelID < ExcelModels.Length Then
                ExcelModels(ModelID) = Nothing
            End If

            ExcelModelCount = PreviousModelCount

            If PreviousModelCount < 0 Then
                ExcelModels = Nothing
            ElseIf ExcelModels IsNot Nothing AndAlso
                   ExcelModels.Length > PreviousModelCount + 1 Then
                ReDim Preserve ExcelModels(PreviousModelCount)
            End If

        End Sub


#End Region






    End Class

End Namespace
