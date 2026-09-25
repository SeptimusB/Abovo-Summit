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

            Private Const UseCustomCalculationService As Boolean = True

            Public ModelSpreadsheetControl As SpreadsheetControl
            Public SSViewer As MainModelViewer
            Public WB As IWorkbook
            Public WBStructure As Abovo_Model_Def
            Public Profile As WorkbookModelProfile
            Public ManagedDefinition As ModelManagerDefinition
            Public ManagedDefinitionWarning As String
            Public WBStructureManager As StructureManager
            Public WBData As DataManager
            Public WBInterface As InterfaceManager
            Public InterfaceHistory As InterfaceHistoryService
            Public WBCalcEngine As CalcEngine
            Public WBCalculationService As CustomCalcEngine
            Public WBDataPres As PresentationManager
            Public EventCoordinator As EventManager
            Public ExpendAnalyser As BPIncomeExpenditureAnalyser
            Public ExpendAnalyserV2 As BPIncomeExpenditureAnalyserV2
            Public ReadOnly ResourceRegistry As New ModelResourceRegistry
            Private _isDirty As Boolean
            Private _userChangeRevision As Long
            Private _savedUserChangeRevision As Long
            Private _calculationRevision As Long
            Private _calculatedRevision As Long = -1
            Private _needsFullRebuild As Boolean = True
            'An unopened/unverified dependency graph needs a rebuild before its
            'results are consumed, but does not imply an in-session structural edit.
            Private _saveRebuildRequired As Boolean
            Private _formulaStructureRevision As Long
            Private _checkedFormulaStructureRevision As Long = -1
            Private _knownCloseValidationFailure As Boolean
            Private _recoveryCheckSheetFailed As Boolean
            Private _rememberedCheckSheetWarning As Boolean
            Private _lastAcceptedCheckSheetRevision As Long = -1
            Private _checkSheetPublicationInProgress As Boolean
            Private _openingCheckSheetAcceptable As Boolean
            Private _manualSavedSinceOpen As Boolean
            Private _openingCheckSheetPath As String
            Private _openingCheckSheetHash As String
            Private _openingUserRevision As Long
            Private _saveAfterCheckSheetClear As Boolean
            Private _writingPreparedWorkbook As Boolean
            Private _writingRecoveryWorkbook As Boolean
            Public RecoverySourcePath As String
            Friend RecoveryHasVerifiedBinaryMetadata As Boolean
            Private Const ResultsPendingProperty As String = "Abovo.Summit.ResultsPending"
            Private _deferredSaveResultsPending As Boolean

            Public ReadOnly Property ResultsPending As Boolean
                Get
                    Return _needsFullRebuild OrElse _calculatedRevision <> _calculationRevision
                End Get
            End Property

            Public ReadOnly Property DeferredSaveResultsPending As Boolean
                Get
                    Return _deferredSaveResultsPending
                End Get
            End Property

            'Saved inputs and calculated outputs are separate states. Sheet-level
            'CalculateWSs/IsCalculated refreshes do not certify whole-model caches.
            Public Function EnsureDeferredSaveResultsCurrent(reason As String) As Boolean
                If Not _deferredSaveResultsPending Then Return False
                Using activity As New FormSplashScreen(Form.ActiveForm, "Updating business plan results", reason)
                    EnsureSaveCalculationCurrent(AddressOf activity.Update)
                    If WBCalcEngine IsNot Nothing Then WBCalcEngine.RefreshAfterDeferredCalculation()
                    If InstanceInterface IsNot Nothing AndAlso Not InstanceInterface.IsDisposed Then InstanceInterface.PopulateFileInfo()
                    activity.Complete("Results updated.")
                End Using
                Return True
            End Function

            Friend Sub RestoreDeferredSaveState()
                Dim properties = WB.DocumentProperties.Custom
                If Not properties.Names.Contains(ResultsPendingProperty) Then Return
                Dim state = properties(ResultsPendingProperty)
                'Unknown/malformed markers are conservative, never treated as current.
                _deferredSaveResultsPending = Not (state.IsBoolean AndAlso Not state.BooleanValue)
                EnsureDeferredSaveResultsCurrent("Updating results saved with calculation pending...")
            End Sub
            'Retain the exact delegates: relaxed AddressOf conversions otherwise
            'create new wrappers that cannot be removed during model shutdown.
            Private ReadOnly SavedHandler As EventHandler = Sub(sender, e)
                                                                If Not _writingRecoveryWorkbook Then
                                                                    _manualSavedSinceOpen = True
                                                                    ClearDirtyFlag()
                                                                End If
                                                            End Sub
            Private ReadOnly ContentHandler As EventHandler = Sub(sender, e) WorkbookContentChanged()
            Private ReadOnly PropertiesHandler As DocumentPropertiesChangedEventHandler = Sub(sender, e)
                                                                                              'Save updates author/timestamp metadata itself.
                                                                                              If Not _writingPreparedWorkbook Then SetDirtyFlag(userChange:=False)
                                                                                          End Sub
            Private ReadOnly NativeCellHandler As CellValueChangedEventHandler = Sub(sender, e) NativeWorkbookStructureChanged()
            Private ReadOnly NativeRowsInsertedHandler As RowsInsertedEventHandler = Sub(sender, e) NativeWorkbookStructureChanged()
            Private ReadOnly NativeRowsRemovedHandler As RowsRemovedEventHandler = Sub(sender, e) NativeWorkbookStructureChanged()
            Private ReadOnly NativeColumnsInsertedHandler As ColumnsInsertedEventHandler = Sub(sender, e) NativeWorkbookStructureChanged()
            Private ReadOnly NativeColumnsRemovedHandler As ColumnsRemovedEventHandler = Sub(sender, e) NativeWorkbookStructureChanged()
            Private ReadOnly NativeSheetInsertedHandler As SheetInsertedEventHandler = Sub(sender, e) NativeWorkbookStructureChanged()
            Private ReadOnly NativeSheetRemovedHandler As SheetRemovedEventHandler = Sub(sender, e) NativeWorkbookStructureChanged()
            Private ReadOnly NativeSheetRenamedHandler As SheetRenamedEventHandler = Sub(sender, e) NativeWorkbookStructureChanged()
            Private ReadOnly NativeNameEditedHandler As DefinedNameEditedEventHandler = Sub(sender, e) NativeWorkbookStructureChanged()
            Private ReadOnly NativeNameAddedHandler As DefinedNameAddedEventHandler = Sub(sender, e) NativeWorkbookStructureChanged()
            Private ReadOnly NativeNameDeletedHandler As DefinedNameDeletedEventHandler = Sub(sender, e) NativeWorkbookStructureChanged()

            Public Event DirtyStateChanged As EventHandler
            Friend Event MetadataChanged As EventHandler

            Private Sub CommittedMetadataChanged(sender As Object, e As ChangeHistoryChangedEventArgsV2)
                If Not e.WorksheetNames.Contains("Global Assumptions", StringComparer.OrdinalIgnoreCase) Then Return
                RefreshModelMetadata()
            End Sub

            Friend Sub RefreshModelMetadata()
                If WB Is Nothing OrElse WBStructure Is Nothing OrElse Profile Is Nothing Then Return
                Dim previousCompany = WBStructure.CompanyName
                Dim previousStart = WBStructure.StartDate
                Profile.ApplyMetadata(WB, WBStructure)
                If String.Equals(previousCompany, WBStructure.CompanyName, StringComparison.Ordinal) AndAlso
                   String.Equals(previousStart, WBStructure.StartDate, StringComparison.Ordinal) Then Return

                'Metadata is a read-only projection of the committed workbook. Do
                'not recalculate, dirty the file again, or rebuild any interface.
                Dim subscribers As EventHandler = MetadataChangedEvent
                If subscribers Is Nothing Then Return
                For Each subscriber As [Delegate] In subscribers.GetInvocationList()
                    Try
                        DirectCast(subscriber, EventHandler).Invoke(Me, EventArgs.Empty)
                    Catch ex As Exception
                        Abovo.SummitDiagnostics.WriteLine("[Model metadata] Interface refresh failed: " & ex.Message)
                    End Try
                Next
            End Sub
            Friend Event ManualSaveAvailabilityChanged As EventHandler

            Friend ReadOnly Property ManualSaveAvailable As Boolean
                Get
                    Return IsDirty OrElse _saveAfterCheckSheetClear
                End Get
            End Property

            Private Sub OfferSaveAfterCheckSheetClear(available As Boolean)
                If _saveAfterCheckSheetClear = available Then Return
                _saveAfterCheckSheetClear = available
                Try
                    RaiseEvent ManualSaveAvailabilityChanged(Me, EventArgs.Empty)
                Catch ex As Exception
                    Abovo.SummitDiagnostics.WriteLine("[Save availability] Interface refresh failed: " & ex.Message)
                End Try
            End Sub

            Public Property IsDirty As Boolean
                Get
                    Return _isDirty
                End Get
                Set(value As Boolean)
                    Dim changed = (_isDirty <> value)
                    _isDirty = value
                    If Not value Then _savedUserChangeRevision = _userChangeRevision
                    If value Then MarkCalculationPending()
                    If changed Then RaiseEvent DirtyStateChanged(Me, EventArgs.Empty)
                End Set
            End Property

            'Recovery eligibility is based on committed user work, not recalculation,
            'background reconciliation, document timestamps or delayed content events.
            Friend ReadOnly Property UserChangeRevision As Long
                Get
                    Return _userChangeRevision
                End Get
            End Property

            Friend ReadOnly Property HasUnsavedUserChanges As Boolean
                Get
                    Return IsDirty AndAlso _userChangeRevision <> _savedUserChangeRevision
                End Get
            End Property

            Friend ReadOnly Property RecoveryAutosaveSuspended As Boolean
                Get
                    Return _recoveryCheckSheetFailed AndAlso Not RecoveryBackupManager.ContinueOnCheckSheetError
                End Get
            End Property

            Friend ReadOnly Property CheckSheetWarningActive As Boolean
                Get
                    Return _recoveryCheckSheetFailed AndAlso Not _rememberedCheckSheetWarning
                End Get
            End Property

            Friend ReadOnly Property CheckSheetWarningNeedsRecheck As Boolean
                Get
                    Return _recoveryCheckSheetFailed AndAlso _rememberedCheckSheetWarning
                End Get
            End Property

            Friend ReadOnly Property CheckSheetWarningCaption As String
                Get
                    Return If(CheckSheetWarningNeedsRecheck, "(Check sheet: recheck required)", "(Check sheet)")
                End Get
            End Property

            Friend Event CheckSheetStatusChanged As EventHandler
            Friend Event CheckSheetValuesRefreshed As EventHandler

            Friend ReadOnly Property LastAcceptedCheckSheetRevision As Long
                Get
                    Return _lastAcceptedCheckSheetRevision
                End Get
            End Property

            Private Sub NotifyCheckSheetSubscribers(subscribers As EventHandler, context As String)
                If subscribers Is Nothing Then Return
                For Each subscriber As [Delegate] In subscribers.GetInvocationList()
                    Try
                        DirectCast(subscriber, EventHandler).Invoke(Me, EventArgs.Empty)
                    Catch ex As Exception
                        'One stale presentation must not prevent other views from
                        'observing the accepted result or escape into a transaction.
                        Abovo.SummitDiagnostics.WriteLine("[Check Sheet " & context & "] Interface refresh failed: " & ex.Message)
                    End Try
                Next
            End Sub

            Private Sub SetCheckSheetWarning(active As Boolean, Optional remembered As Boolean = False)
                remembered = active AndAlso remembered
                If _recoveryCheckSheetFailed = active AndAlso _rememberedCheckSheetWarning = remembered Then Return
                _recoveryCheckSheetFailed = active
                _rememberedCheckSheetWarning = remembered
                NotifyCheckSheetSubscribers(CheckSheetStatusChangedEvent, "warning")
            End Sub

            Friend Sub RestoreRecoveryAutosaveHold(paused As Boolean)
                If Not paused Then Return
                If Not CheckSheetWarningActive Then SetCheckSheetWarning(True, remembered:=True)
                _knownCloseValidationFailure = True
            End Sub

            Private Shared Function CheckSheetFileHash(stream As System.IO.Stream) As String
                Using hash = System.Security.Cryptography.SHA256.Create()
                    Return Convert.ToBase64String(hash.ComputeHash(stream))
                End Using
            End Function

            Friend Sub CaptureCheckSheetDiskBeforeLoad()
                _openingCheckSheetHash = Nothing
                Try
                    _openingCheckSheetPath = System.IO.Path.GetFullPath(FileName)
                    Using source As New System.IO.FileStream(_openingCheckSheetPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                        _openingCheckSheetHash = CheckSheetFileHash(source)
                    End Using
                Catch ex As Exception
                    Abovo.SummitDiagnostics.WriteLine("[Check Sheet warning] Opening fingerprint unavailable: " & ex.Message)
                End Try
            End Sub

            Friend Sub CaptureCheckSheetOpenState()
                'A discard baseline describes the saved file AS OPENED, not a new
                'integrity certification. Do not clear an existing hold from cached
                'values. It only lets us undo a later, wholly unsaved session hold.
                _openingCheckSheetAcceptable = False
                _manualSavedSinceOpen = False
                If _openingCheckSheetHash Is Nothing OrElse RecoverySaveAsRequired OrElse Not String.IsNullOrWhiteSpace(RecoverySourcePath) Then Return
                Try
                    Using source As New System.IO.FileStream(_openingCheckSheetPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                        If CheckSheetFileHash(source) <> _openingCheckSheetHash Then Throw New System.IO.IOException("The saved file changed while it was opening.")
                    End Using
                    _openingUserRevision = _userChangeRevision
                    _openingCheckSheetAcceptable = Not CheckSheetWarningActive AndAlso Not ReadCheckSheetValidation().HasFailures
                Catch ex As Exception
                    _openingCheckSheetHash = Nothing
                    Abovo.SummitDiagnostics.WriteLine("[Check Sheet warning] Discard baseline unavailable: " & ex.Message)
                End Try
            End Sub

            Friend ReadOnly Property CanDiscardSessionCheckSheetWarning As Boolean
                Get
                    Return CheckSheetWarningActive AndAlso _openingCheckSheetAcceptable AndAlso
                        Not _manualSavedSinceOpen AndAlso _openingCheckSheetHash IsNot Nothing AndAlso
                        Not RecoverySaveAsRequired AndAlso IntegrityState = ModelIntegrityState.Healthy AndAlso
                        String.Equals(_openingCheckSheetPath, FileName, StringComparison.OrdinalIgnoreCase)
                End Get
            End Property

            Private Sub ClearDiscardedSessionCheckSheetWarning()
                If Not CanDiscardSessionCheckSheetWarning Then Return
                Try
                    'Keep the original protected against concurrent writes/deletion
                    'until the settings update finishes. Timestamp equality is not
                    'enough: compare its actual bytes with the opening baseline.
                    Using source As New System.IO.FileStream(_openingCheckSheetPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                        If CheckSheetFileHash(source) <> _openingCheckSheetHash Then Return
                        SetCheckSheetWarning(False)
                        _knownCloseValidationFailure = False
                        SystemMessageManager.Publish(ModelID, "Discarded the unsaved session's Check Sheet warning. The saved business plan is unchanged from opening.",
                            SystemMessageSeverity.Information, "Integrity", FileName)
                    End Using
                Catch ex As Exception
                    Abovo.SummitDiagnostics.WriteLine("[Check Sheet warning] Discard could not clear the remembered hold: " & ex.Message)
                End Try
            End Sub

            Friend Sub MarkUserChange()
                If IsClosing OrElse _writingPreparedWorkbook Then Return
                _userChangeRevision += 1
                IsDirty = True
            End Sub

            Public ReadOnly Property NeedsFullRebuild As Boolean
                Get
                    Return _needsFullRebuild
                End Get
            End Property

            Public ReadOnly Property NeedsSaveRebuild As Boolean
                Get
                    Return _saveRebuildRequired
                End Get
            End Property

            Public ReadOnly Property CloseValidationRequired As Boolean
                Get
                    Return WB Is Nothing OrElse IsDirty OrElse RecoverySaveAsRequired OrElse
                        IntegrityState <> ModelIntegrityState.Healthy OrElse _knownCloseValidationFailure OrElse
                        ModelSafetyManager.IsBulkWorkbookMutationInProgress(ModelID)
                End Get
            End Property

            Public ReadOnly Property NeedsFormulaPreflight As Boolean
                Get
                    Return _checkedFormulaStructureRevision <> _formulaStructureRevision
                End Get
            End Property

            Friend ReadOnly Property CalculationRevision As Long
                Get
                    Return _calculationRevision
                End Get
            End Property

            Friend Sub MarkCalculationPending()
                _calculationRevision += 1
            End Sub

            Public Sub RequireFullRebuild()
                If Not _needsFullRebuild Then Abovo.SummitDiagnostics.WriteLine("[Rebuild State] model=" & ModelID.ToString() & ", pending=True")
                _needsFullRebuild = True
                _saveRebuildRequired = True
                'A calculation/rebuild does not prove XLSB export compatibility.
                'Keep this separate from ordinary value/calculation revisions.
                _formulaStructureRevision += 1
                MarkCalculationPending()
            End Sub

            Friend Sub MarkFullCalculationCurrent(revision As Long, rebuilt As Boolean)
                If revision <> _calculationRevision OrElse ModelSafetyManager.IsBulkWorkbookMutationInProgress(ModelID) Then
                    Abovo.SummitDiagnostics.WriteLine("[Rebuild State] completion withheld: revision=" & revision.ToString() & ", current=" & _calculationRevision.ToString())
                    Return
                End If
                If rebuilt Then
                    _needsFullRebuild = False
                    _saveRebuildRequired = False
                End If
                If Not _needsFullRebuild Then
                    _calculatedRevision = revision
                    _deferredSaveResultsPending = False
                End If
            End Sub
            Public IntegrityState As ModelIntegrityState = ModelIntegrityState.Healthy
            Public RecoverySaveAsRequired As Boolean = False
            Public IntegrityReason As String = String.Empty
            Public IntegrityOperation As String = String.Empty
            Public IsClosing As Boolean
            Public ModelID As Integer
            Public ChangeManager As ModelChangeManagerV2
            Public HistoryManager As HistoryManagerV2
            Public PdfExportManager As DITPdfExportManager
            Public ExcelExportManager As DITExcelExportManager
            Public TransDBM As TransDBManager
            Public ColourSwatch As Color
            Public SSViewInitialised As Boolean = False
            Public FileName As String
            Public FileInfo As System.IO.FileInfo
            Public PreviousFileAccessTime As DateTime
            Public InstanceInterface As FileInstanceInterface
            Public RDSM As RDSManager
            Public InterfaceDependencies As InterfaceDependencyManager
            Public TransDBSync As TransactionalDBSynchroniser
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
                    InterfaceHistory = New InterfaceHistoryService(SetModelID)
                    WBInterface = New InterfaceManager(SetModelID)
                    WBDataPres = New PresentationManager(SetModelID)
                    RDSM = New RDSManager(SetModelID)
                    InterfaceDependencies = New InterfaceDependencyManager(SetModelID)
                End If

                IsDirty = False
                WB = ModelSpreadsheetControl.Document
                ModelSpreadsheetControl.ActiveWorksheet.ActiveView.ShowGridlines = False

                InitiateWorkbook(SetModelID)

                AddHandler ModelSpreadsheetControl.DocumentSaved, SavedHandler
                AddHandler ModelSpreadsheetControl.DocumentPropertiesChanged, PropertiesHandler
                AddHandler ModelSpreadsheetControl.ContentChanged, ContentHandler
                AddHandler ModelSpreadsheetControl.CellValueChanged, NativeCellHandler
                AddHandler ModelSpreadsheetControl.RowsInserted, NativeRowsInsertedHandler
                AddHandler ModelSpreadsheetControl.RowsRemoved, NativeRowsRemovedHandler
                AddHandler ModelSpreadsheetControl.ColumnsInserted, NativeColumnsInsertedHandler
                AddHandler ModelSpreadsheetControl.ColumnsRemoved, NativeColumnsRemovedHandler
                AddHandler ModelSpreadsheetControl.SheetInserted, NativeSheetInsertedHandler
                AddHandler ModelSpreadsheetControl.SheetRemoved, NativeSheetRemovedHandler
                AddHandler ModelSpreadsheetControl.SheetRenamed, NativeSheetRenamedHandler
                AddHandler ModelSpreadsheetControl.DefinedNameEdited, NativeNameEditedHandler
                AddHandler ModelSpreadsheetControl.DefinedNameAdded, NativeNameAddedHandler
                AddHandler ModelSpreadsheetControl.DefinedNameDeleted, NativeNameDeletedHandler
                AddHandler ModelSpreadsheetControl.UnhandledException, AddressOf SSCUnhandledEvent
                AddHandler ModelSpreadsheetControl.ActiveSheetChanged, AddressOf ProcessSheetChange

            End Sub
            Sub ClearDirtyFlag()

                IsDirty = False

            End Sub


            Sub SetDirtyFlag(Optional userChange As Boolean = True)

                If userChange Then
                    MarkUserChange()
                Else
                    IsDirty = True
                End If
                RequireFullRebuild() 'Unknown/bulk import or document-property change.
                If WBCalcEngine IsNot Nothing Then
                    WBCalcEngine.MarkPotentialWorkbookChange()
                End If

            End Sub
            Private Sub WorkbookContentChanged()
                'ContentChanged can be delivered later, after load/save has set
                'the native save point. Do not revive those already-clean changes.
                If Not IsClosing AndAlso Not _writingPreparedWorkbook AndAlso ModelSpreadsheetControl.Modified Then IsDirty = True
            End Sub

            Private Sub NativeWorkbookStructureChanged()
                If IsClosing OrElse _writingPreparedWorkbook Then Return
                'UI-only events: a native formula edit, rename or insertion is not
                'necessarily routed through the typed-edit/structural services.
                MarkUserChange()
                RequireFullRebuild()
                If WBCalcEngine IsNot Nothing Then WBCalcEngine.InvalidateDependencyGraph()
            End Sub
            Sub ProcessSheetChange()

                ModelSpreadsheetControl.ActiveWorksheet.ActiveView.ShowGridlines = False

            End Sub
            Public Sub ShowSpreadsheet(Optional ByVal SetSpreadsheet As DevExpress.Spreadsheet.Worksheet = Nothing, Optional ByVal Parent As Object = Nothing)

                EnsureDeferredSaveResultsCurrent("Opening the spreadsheet...")

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
                    ChangeManager = New ModelChangeManagerV2(ModelID)
                    AddHandler ChangeManager.HistoryChanged, AddressOf CommittedMetadataChanged
                    HistoryManager = New HistoryManagerV2(ModelID)
                    PdfExportManager = New DITPdfExportManager(ModelID)
                    ExcelExportManager = New DITExcelExportManager(ModelID)
                    EventCoordinator = New EventManager(
                        ModelID,
                        If(Profile Is Nothing, String.Empty, Profile.ModelType))

                    If Profile IsNot Nothing AndAlso
                       Profile.UsesTransactionalDatabase Then

                        Dim TransactionSheetID As Integer =
                            GetSheetID(ModelID, "Transactional DB")

                        If TransactionSheetID < 0 Then
                            Throw New InvalidOperationException(
                                "The workbook is missing the 'Transactional DB' worksheet.")
                        End If

                        If UseCustomCalculationService Then
                            WBCalculationService = New CustomCalcEngine(ModelID) With {
                                .TransDBSheetID = TransactionSheetID,
                                .ComparisonSheetID = GetSheetID(ModelID, "TDB Comparison"),
                                .CheckSheetID = GetSheetID(ModelID, "Check Sheet"),
                                .DontCalcTDBS = True
                            }

                            WB.AddService(
                                GetType(DevExpress.XtraSpreadsheet.Services.ICustomCalculationService),
                                WBCalculationService)
                        Else
                            WBCalculationService = Nothing
                        End If

                        TransDBM = New TransDBManager(ModelID)
                        TransDBSync = New TransactionalDBSynchroniser(ModelID)
                        WorkbookStructureRules =
                            New WorkbookStructureRuleManager(ModelID)
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

                SetDirtyFlag()

                Return True

            End Function
            Public Function SaveFileAs(Optional ByVal RequireDifferentPath As Boolean = False) As Boolean

                Dim OriginalPath As String = FileName
                Dim restoringRecovery = Not String.IsNullOrWhiteSpace(RecoverySourcePath)
                RequireDifferentPath = RequireDifferentPath OrElse restoringRecovery

                Try

                    If RequireDifferentPath Then

                        Using SaveDialog As New SaveFileDialog()

                            SaveDialog.Title = If(restoringRecovery, "Save recovered business plan as XLSB", "Save validated model copy")
                            SaveDialog.Filter = "Excel Binary Workbook (*.xlsb)|*.xlsb"
                            SaveDialog.DefaultExt = "xlsb"
                            SaveDialog.AddExtension = True
                            SaveDialog.OverwritePrompt = True
                            SaveDialog.CheckPathExists = True
                            SaveDialog.RestoreDirectory = True

                            Dim OriginalDirectory As String =
                                System.IO.Path.GetDirectoryName(If(restoringRecovery, RecoverySourcePath, OriginalPath))

                            If Not String.IsNullOrWhiteSpace(OriginalDirectory) AndAlso
                               System.IO.Directory.Exists(OriginalDirectory) Then

                                SaveDialog.InitialDirectory = OriginalDirectory
                            End If

                            SaveDialog.FileName =
                                If(restoringRecovery, System.IO.Path.GetFileNameWithoutExtension(RecoverySourcePath) & ".xlsb",
                                   System.IO.Path.GetFileNameWithoutExtension(OriginalPath) & " - validation copy.xlsb")

                            Do

                                If SaveDialog.ShowDialog() <> DialogResult.OK Then Return False

                                Dim SelectedPath As String =
                                    System.IO.Path.GetFullPath(SaveDialog.FileName)

                                If String.Equals(
                                    SelectedPath,
                                    System.IO.Path.GetFullPath(OriginalPath),
                                    StringComparison.OrdinalIgnoreCase) Then

                                    MessageBox.Show(
                                        "The validation copy must be saved to a different file. " &
                                        "Please choose another name or location.",
                                        "Save validated model copy",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information)

                                    Continue Do
                                End If

                                If Not SavePreparedWorkbook(Sub() ModelSpreadsheetControl.SaveDocument(SelectedPath, DocumentFormat.Xlsb)) Then Return False
                                FileName = SelectedPath
                                FileInfo = New System.IO.FileInfo(FileName)
                                IsDirty = False
                                RecoverySaveAsRequired = False
                                RefreshSavedFilePresentation()
                                SystemMessageManager.Publish(
                                    ModelID,
                                    If(IntegrityState = ModelIntegrityState.RecoveryRequired,
                                       "Recovery copy saved successfully. The original workbook was not overwritten.",
                                       "Model saved as '" & System.IO.Path.GetFileName(FileName) & "'."),
                                    SystemMessageSeverity.Success,
                                    "File Manager",
                                    FileName)
                                Return True

                            Loop

                        End Using

                    End If

                    If Not SavePreparedWorkbook(Sub() ModelSpreadsheetControl.SaveDocumentAs(), ShowNativeSaveDialog:=True) Then Return False

                    Dim SavedPath As String = WB.Path
                    If String.IsNullOrWhiteSpace(SavedPath) Then Return False

                    FileName = System.IO.Path.GetFullPath(SavedPath)
                    FileInfo = New System.IO.FileInfo(FileName)
                    IsDirty = False
                    RecoverySaveAsRequired = False
                    RefreshSavedFilePresentation()
                    SystemMessageManager.Publish(
                        ModelID,
                        "Model saved as '" & System.IO.Path.GetFileName(FileName) & "'.",
                        SystemMessageSeverity.Success,
                        "File Manager",
                        FileName)
                    Return True

                Catch ex As Exception
                    SystemMessageManager.Publish(
                        ModelID,
                        "The model could not be saved to the selected location: " & ex.Message,
                        SystemMessageSeverity.Error,
                        "File Manager",
                        OriginalPath)
                    MessageBox.Show(
                        "The model could not be saved to the selected location." &
                        Environment.NewLine & Environment.NewLine & ex.Message,
                        "Save Abovo Model As",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
                    Return False
                End Try

            End Function
            Public Function SaveFileAsTo(ByVal SelectedPath As String,
                                         Optional ByVal RequireDifferentPath As Boolean = True) As Boolean
                Dim OriginalPath As String = FileName
                Try
                    If String.IsNullOrWhiteSpace(SelectedPath) Then Return False
                    Dim FullPath As String = System.IO.Path.GetFullPath(SelectedPath)
                    If Not String.Equals(System.IO.Path.GetExtension(FullPath), ".xlsb", StringComparison.OrdinalIgnoreCase) Then
                        FullPath = System.IO.Path.ChangeExtension(FullPath, ".xlsb")
                    End If
                    If RequireDifferentPath AndAlso
                       String.Equals(FullPath, System.IO.Path.GetFullPath(OriginalPath), StringComparison.OrdinalIgnoreCase) Then
                        Throw New InvalidOperationException("The populated model must be saved to a different file.")
                    End If
                    If Not SavePreparedWorkbook(Sub() ModelSpreadsheetControl.SaveDocument(FullPath, DocumentFormat.Xlsb)) Then Return False
                    FileName = FullPath
                    FileInfo = New System.IO.FileInfo(FileName)
                    IsDirty = False
                    RecoverySaveAsRequired = False
                    RefreshSavedFilePresentation()
                    SystemMessageManager.Publish(
                        ModelID,
                        "Populated model saved as '" & System.IO.Path.GetFileName(FileName) & "'.",
                        SystemMessageSeverity.Success,
                        "Business Plan Population",
                        FileName)
                    Return True
                Catch ex As Exception
                    SystemMessageManager.Publish(
                        ModelID,
                        "The populated model could not be saved: " & ex.Message,
                        SystemMessageSeverity.Error,
                        "Business Plan Population",
                        OriginalPath)
                    MessageBox.Show(
                        "The populated model could not be saved to the selected location." &
                        Environment.NewLine & Environment.NewLine & ex.Message,
                        "Save populated Business Plan",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
                    Return False
                End Try
            End Function

            Public Function SaveFile() As Boolean

                If RecoverySaveAsRequired Then
                    SystemMessageManager.Publish(
                        ModelID,
                        If(Not String.IsNullOrWhiteSpace(RecoverySourcePath),
                           "This is a recovery copy. Save As will suggest the original XLSB filename; confirm replacement or choose a new name.",
                           "Normal save was redirected to Save As because an earlier workbook operation could not be rolled back reliably."),
                        SystemMessageSeverity.Warning,
                        "File Manager",
                        FileName)
                    Return SaveFileAs(RequireDifferentPath:=True)
                End If

                If Not ManualSaveAvailable Then Return True

                Try
                    If Not SavePreparedWorkbook(Sub() ModelSpreadsheetControl.SaveDocument()) Then Return False
                    IsDirty = False
                    RefreshSavedFilePresentation()
                    SystemMessageManager.Publish(
                        ModelID,
                        "Model saved successfully.",
                        SystemMessageSeverity.Success,
                        "File Manager",
                        FileName)
                    Return True

                Catch ex As Exception
                    SystemMessageManager.Publish(
                        ModelID,
                        "The model could not be saved: " & ex.Message,
                        SystemMessageSeverity.Error,
                        "File Manager",
                        FileName)
                    MessageBox.Show(
                        "Sorry, an error occurred while saving. Please check the file is not open in another program and that you have write permissions to the location." &
                        Environment.NewLine & Environment.NewLine & ex.Message,
                        "Save Abovo Model",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
                    Return False
                End Try

            End Function
            'A stream export must not become the document's Save point or filename.
            'XLSM retains the unmodified formulas; the normal XLSB preflight remains
            'mandatory when the recovered plan is explicitly saved as XLSB.
            Friend Sub WriteRecoverySnapshot(output As System.IO.Stream)
                If ChangeManager IsNot Nothing AndAlso ChangeManager.HasEngineEditingTrial Then Throw New InvalidOperationException("The engine-owned model requires its verified engine recovery path.")
                If _writingPreparedWorkbook OrElse _writingRecoveryWorkbook Then Throw New InvalidOperationException("Another save is in progress.")
                Dim nativeModified = ModelSpreadsheetControl.Modified
                Dim dirty = _isDirty
                Dim revision = _calculationRevision
                Dim originalPath = WB.Path
                Dim mode = WB.Options.CalculationMode
                Dim properties = WB.DocumentProperties.Custom
                Dim names = {RecoveryBackupStore.SourceProperty, RecoveryBackupStore.VersionProperty,
                             RecoveryBackupStore.DateProperty, RecoveryBackupStore.PendingProperty}
                Dim originals = names.ToDictionary(Function(n) n, Function(n) properties(n))
                Dim modifiedAt = WB.DocumentProperties.Modified
                Dim modifiedBy = WB.DocumentProperties.LastModifiedBy
                _writingPreparedWorkbook = True
                _writingRecoveryWorkbook = True
                Try
                    WB.Options.CalculationMode = WorkbookCalculationMode.Manual
                    properties(RecoveryBackupStore.SourceProperty) = System.IO.Path.GetFullPath(FileName)
                    properties(RecoveryBackupStore.VersionProperty) = "1"
                    properties(RecoveryBackupStore.DateProperty) = DateTime.UtcNow.ToString("O")
                    properties(RecoveryBackupStore.PendingProperty) = True
                    WB.SaveDocument(output, DocumentFormat.Xlsm)
                    If WB.Path <> originalPath Then Throw New InvalidOperationException("Recovery export unexpectedly changed the document path.")
                Finally
                    Try
                        For Each item In originals
                            properties(item.Key) = item.Value
                        Next
                        WB.DocumentProperties.Modified = modifiedAt
                        WB.DocumentProperties.LastModifiedBy = modifiedBy
                        WB.Options.CalculationMode = mode
                        ModelSpreadsheetControl.Modified = nativeModified
                        _isDirty = dirty
                        _calculationRevision = revision
                    Finally
                        _writingRecoveryWorkbook = False
                        _writingPreparedWorkbook = False
                    End Try
                End Try
            End Sub

            Private Function SavePreparedWorkbook(saveAction As Action, Optional ShowNativeSaveDialog As Boolean = False) As Boolean
                If ChangeManager IsNot Nothing AndAlso ChangeManager.HasEngineEditingTrial Then Throw New InvalidOperationException("The engine-owned model requires its verified engine save path.")
                'Value edits retain their immediate sheet calculation but may save
                'pending whole-model caches. Structural/unknown changes still rebuild.
                Dim previousEngine = WB.Options.CalculationEngineType
                Dim previousSkip = If(WBCalculationService Is Nothing, False, WBCalculationService.DontCalcTDBS)
                Dim originalDirty = IsDirty
                Dim originalDeferred = _deferredSaveResultsPending
                Dim originalMarker As CellValue = WB.DocumentProperties.Custom(ResultsPendingProperty)
                Dim saved As Boolean = False
                Dim checkedFormulaRevision As Long = -1
                Dim saveActionMs As Long
                Dim savedHandler As EventHandler = Sub(sender, args) saved = True
                Dim clock = Abovo.SummitDiagnostics.DiagnosticTimer.StartNew()
                AddHandler ModelSpreadsheetControl.DocumentSaved, savedHandler
                Try
                    Using activity As New FormSplashScreen(Form.ActiveForm, "Saving business plan",
                        If(NeedsFormulaPreflight, "Checking XLSB formula compatibility...", "Preparing workbook results..."))
                        Dim checkFormulas = NeedsFormulaPreflight
                        WorkbookXlsbFormulaCompatibility.Save(WB,
                            Function()
                                If Not checkFormulas AndAlso NeedsFormulaPreflight Then
                                    Throw New InvalidOperationException("Workbook formulas changed during save preparation. Please retry the save.")
                                End If
                                checkedFormulaRevision = _formulaStructureRevision
                                If ModelSafetyManager.IsBulkWorkbookMutationInProgress(ModelID) Then
                                    Throw New InvalidOperationException("Wait for the current workbook operation to finish before saving.")
                                End If
                                If NeedsSaveRebuild OrElse (NeedsFullRebuild AndAlso
                                   (RecoverySaveAsRequired OrElse IntegrityState <> ModelIntegrityState.Healthy OrElse _knownCloseValidationFailure)) Then
                                    If WBCalculationService IsNot Nothing Then WBCalculationService.DontCalcTDBS = False
                                    WB.Options.CalculationEngineType = CalculationEngineType.Recursive
                                    EnsureSaveCalculationCurrent(AddressOf activity.Update)
                                Else
                                    Abovo.SummitDiagnostics.WriteLine("[Save Calculation Benchmark] model=" & ModelID.ToString() &
                                        ", mode=" & If(ResultsPending, "deferred", "current") & ", total=0 ms" &
                                        ", initialRebuildPending=" & NeedsFullRebuild.ToString())
                                End If
                                If checkedFormulaRevision <> _formulaStructureRevision Then
                                    Throw New InvalidOperationException("Workbook formulas changed during calculation. Please retry the save.")
                                End If
                                Abovo.SummitDiagnostics.WriteLine("[XLSB Save Benchmark] prepared=" & clock.ElapsedMilliseconds.ToString() & " ms")
                                'Do not cover a native Save As dialog with a wait form.
                                If ShowNativeSaveDialog Then activity.Dispose() Else activity.Update("Writing the workbook...")
                                _writingPreparedWorkbook = True
                                Dim writeStarted = clock.ElapsedMilliseconds
                                Try
                                    _deferredSaveResultsPending = ResultsPending
                                    WB.DocumentProperties.Custom(ResultsPendingProperty) = _deferredSaveResultsPending
                                    saveAction()
                                    If Not saved Then WB.DocumentProperties.Custom(ResultsPendingProperty) = originalMarker
                                Catch
                                    WB.DocumentProperties.Custom(ResultsPendingProperty) = originalMarker
                                    Throw
                                Finally
                                    saveActionMs = clock.ElapsedMilliseconds - writeStarted
                                    _writingPreparedWorkbook = False
                                End Try
                                If saved Then activity.Complete("Business plan saved.")
                                Return saved
                            End Function, AddressOf RequireFullRebuild, checkFormulas)
                        If saved AndAlso checkedFormulaRevision = _formulaStructureRevision AndAlso
                           Not ModelSafetyManager.IsBulkWorkbookMutationInProgress(ModelID) Then
                            _checkedFormulaStructureRevision = checkedFormulaRevision
                        End If
                    End Using
                Finally
                    RemoveHandler ModelSpreadsheetControl.DocumentSaved, savedHandler
                    Dim restoreStarted = clock.ElapsedMilliseconds
                    Try
                        If WB.Options.CalculationEngineType <> previousEngine Then WB.Options.CalculationEngineType = previousEngine
                    Finally
                        If WBCalculationService IsNot Nothing Then WBCalculationService.DontCalcTDBS = previousSkip
                        If Not saved Then
                            _deferredSaveResultsPending = originalDeferred
                            IsDirty = originalDirty
                            RequireFullRebuild() 'The compatibility guard may have rolled formulas back.
                        End If
                        Abovo.SummitDiagnostics.WriteLine("[XLSB Save Benchmark] total=" & clock.ElapsedMilliseconds.ToString() &
                            " ms, saveAction=" & saveActionMs.ToString() & " ms, restore=" & (clock.ElapsedMilliseconds - restoreStarted).ToString() &
                            " ms, nativeSaveDialog=" & ShowNativeSaveDialog.ToString() & ", saved=" & saved.ToString())
                    End Try
                End Try
                Return saved
            End Function

            Private Sub EnsureSaveCalculationCurrent(Optional report As Action(Of String) = Nothing)
                EnsureCalculationCurrent(report, "Save Calculation")
            End Sub

            Private Sub EnsureCalculationCurrent(report As Action(Of String), benchmark As String)
                If WB Is Nothing Then Throw New InvalidOperationException("The workbook is unavailable.")
                If ModelSafetyManager.IsBulkWorkbookMutationInProgress(ModelID) Then
                    Throw New InvalidOperationException("Wait for the current workbook operation to finish before saving or closing.")
                End If
                Dim revision = _calculationRevision
                Dim rebuild = NeedsFullRebuild
                Dim calculationMode = If(rebuild, "rebuild", If(_calculatedRevision <> revision, "full", "current"))
                Dim timer = Abovo.SummitDiagnostics.DiagnosticTimer.StartNew()
                If calculationMode = "current" Then
                    Abovo.SummitDiagnostics.WriteLine("[" & benchmark & " Benchmark] model=" & ModelID.ToString() & ", mode=current, total=0 ms")
                    Return
                End If
                Dim previousEngine = WB.Options.CalculationEngineType
                Dim previousMode = WB.Options.CalculationMode
                Dim previousSkip = If(WBCalculationService Is Nothing, False, WBCalculationService.DontCalcTDBS)
                Try
                    If report IsNot Nothing Then report(If(rebuild, "Rebuilding workbook dependencies...", "Updating workbook results..."))
                    WB.Options.CalculationMode = WorkbookCalculationMode.Manual
                    If WBCalculationService IsNot Nothing Then WBCalculationService.DontCalcTDBS = False
                    WB.Options.CalculationEngineType = CalculationEngineType.Recursive
                    If rebuild Then WB.CalculateFullRebuild() Else WB.CalculateFull()
                    MarkFullCalculationCurrent(revision, rebuild)
                    If _calculatedRevision <> _calculationRevision OrElse NeedsFullRebuild Then
                        Throw New InvalidOperationException("The workbook changed during calculation. Please retry after the current operation has finished.")
                    End If
                Catch
                    RequireFullRebuild()
                    Throw
                Finally
                    Try
                        WB.Options.CalculationEngineType = previousEngine
                    Finally
                        WB.Options.CalculationMode = previousMode
                        If WBCalculationService IsNot Nothing Then WBCalculationService.DontCalcTDBS = previousSkip
                        Abovo.SummitDiagnostics.WriteLine("[" & benchmark & " Benchmark] model=" & ModelID.ToString() & ", mode=" & calculationMode & ", total=" & timer.ElapsedMilliseconds.ToString() & " ms")
                    End Try
                End Try
            End Sub

            Friend Sub CalculateForIdleIntegrity(report As Action(Of String))
                'Atomic: do not pump input or cancel between engine selection and restoration.
                EnsureCalculationCurrent(report, "Idle Integrity Calculation")
            End Sub

            Friend Function RecordIdleCheckSheetResult(revision As Long, hasFailures As Boolean, Optional worksheetOnly As Boolean = False) As Boolean
                'Preserve the legacy return contract: True means the failure state
                'changed, not merely that an unchanged result was accepted.
                Dim statusChanged As Boolean
                If Not TryRecordCheckSheetResult(revision, hasFailures, statusChanged, worksheetOnly) Then Return False
                Return statusChanged
            End Function

            Friend Function TryRecordCheckSheetResult(revision As Long,
                                                      hasFailures As Boolean,
                                                      ByRef statusChanged As Boolean,
                                                      Optional worksheetOnly As Boolean = False) As Boolean
                statusChanged = False
                'A correction alone does not certify the Check Sheet. Retain the
                'hold through edits/saves until a fresh current check passes.
                'The caller owns calculation-scope and operation-completion proof;
                'this method only publishes that result and never calculates.
                If IsClosing OrElse WB Is Nothing OrElse _checkSheetPublicationInProgress OrElse
                   revision <> _calculationRevision OrElse (ResultsPending AndAlso Not worksheetOnly) Then Return False

                _checkSheetPublicationInProgress = True
                Try
                    _knownCloseValidationFailure = hasFailures
                    Dim wasPaused = _recoveryCheckSheetFailed
                    'Stamp before presentation events. Any later edit advances the
                    'model revision, making this accepted result stale naturally.
                    _lastAcceptedCheckSheetRevision = revision
                    'Checks describe this in-memory revision, not the last saved file.
                    'Persist only after a successful explicit save (or to a recovery's own path).
                    SetCheckSheetWarning(hasFailures)
                    If Not hasFailures AndAlso Not _recoveryCheckSheetFailed AndAlso _openingCheckSheetHash IsNot Nothing AndAlso
                       Not _manualSavedSinceOpen AndAlso _userChangeRevision = _openingUserRevision Then _openingCheckSheetAcceptable = True
                    'Offer an optional explicit save of refreshed results without
                    'inventing user edits, pending calculation or recovery eligibility.
                    If wasPaused AndAlso Not _recoveryCheckSheetFailed Then OfferSaveAfterCheckSheetClear(True)
                    statusChanged = wasPaused <> _recoveryCheckSheetFailed
                    'An unchanged balanced/unbalanced status can still contain new
                    'figures, colours and links. Notify after every accepted result.
                    NotifyCheckSheetSubscribers(CheckSheetValuesRefreshedEvent, "values")
                    Return True
                Finally
                    _checkSheetPublicationInProgress = False
                End Try
            End Function

            Private Sub RefreshSavedFilePresentation()
                'Only successful manual Save/Save As paths reach here. A cancelled
                'dialog, failed write or recovery snapshot must retain the offer.
                OfferSaveAfterCheckSheetClear(False)
                _manualSavedSinceOpen = True
                RecoveryBackupManager.PersistCheckSheetPause(Me, _recoveryCheckSheetFailed)
                RecoverySourcePath = Nothing
                'DocumentSaved fires before Save As updates this model's path.
                'Refresh only after that path and the successful-save state agree.
                Try
                    FileInfo = New System.IO.FileInfo(FileName)
                    FileInfo.Refresh()
                    If InstanceInterface IsNot Nothing AndAlso Not InstanceInterface.IsDisposed Then InstanceInterface.PopulateFileInfo()
                Catch ex As Exception
                    Abovo.SummitDiagnostics.WriteLine("[Save Summary] File summary refresh failed: " & ex.Message)
                End Try
                If WBInterface Is Nothing OrElse WBInterface.GroupInterfaces Is Nothing Then Return
                For Each group In WBInterface.GroupInterfaces
                    Try
                        If group IsNot Nothing AndAlso group.RenderedForm IsNot Nothing AndAlso Not group.RenderedForm.IsDisposed Then group.RenderedForm.RequestSidebarRefresh()
                    Catch ex As Exception
                        Abovo.SummitDiagnostics.WriteLine("[Save Summary] Sidebar refresh failed: " & ex.Message)
                    End Try
                Next
            End Sub

            Public Function CommitToCloseModel() As AbovoTransaction

                Dim CloseTrans As New AbovoTransaction

                Dim saveRequested As Boolean = False
                If IsDirty AndAlso Not RecoverySaveAsRequired AndAlso IntegrityState = ModelIntegrityState.Healthy Then
                    Dim discardResponse = MessageBox.Show(
                        "Save changes to " & System.IO.Path.GetFileName(FileName) & " before closing?",
                        "Close business plan", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button3)
                    If discardResponse = DialogResult.Cancel Then
                        CloseTrans.StringReturn = "Cancel"
                        Return CloseTrans
                    End If
                    If discardResponse = DialogResult.No Then
                        CloseTrans.StringReturn = "Proceed"
                        Return CloseTrans
                    End If
                    saveRequested = True
                End If

                Dim Validation As CloseModelValidationResult

                Try
                    Cursor.Current = Cursors.WaitCursor
                    Validation = CalculateAndValidateForClose()
                Catch ex As Exception
                    Validation = New CloseModelValidationResult With {
                        .ValidationError =
                            "The close validation could not be completed: " &
                            ex.Message
                    }
                Finally
                    Cursor.Current = Cursors.Default
                End Try

                If Not String.IsNullOrWhiteSpace(Validation.ValidationError) Then

                    SystemMessageManager.Publish(
                        ModelID,
                        "Close validation failed. The original workbook cannot be overwritten; a separate validation copy is required.",
                        SystemMessageSeverity.Warning,
                        "Model validation",
                        FileName)

                    Dim ValidationResponse As DialogResult =
                        MessageBox.Show(
                            BuildCloseValidationMessage(Validation),
                            "Model validation warning",
                            MessageBoxButtons.OKCancel,
                            MessageBoxIcon.Warning,
                            MessageBoxDefaultButton.Button2)

                    If ValidationResponse <> DialogResult.OK OrElse
                       Not SaveFileAs(RequireDifferentPath:=True) Then

                        CloseTrans.StringReturn = "Cancel"
                        Return CloseTrans
                    End If

                    CloseTrans.StringReturn = "Proceed"
                    Return CloseTrans
                End If

                If RecoverySaveAsRequired Then
                    Dim RecoveryResponse As DialogResult =
                        MessageBox.Show(
                            "An earlier workbook operation could not be rolled back reliably." &
                            Environment.NewLine & Environment.NewLine &
                            "To protect the original workbook, Summit must save this model as a separate XLSB copy before closing.",
                            "Recovery copy required",
                            MessageBoxButtons.OKCancel,
                            MessageBoxIcon.Warning,
                            MessageBoxDefaultButton.Button2)

                    If RecoveryResponse <> DialogResult.OK OrElse
                       Not SaveFileAs(RequireDifferentPath:=True) Then
                        CloseTrans.StringReturn = "Cancel"
                        Return CloseTrans
                    End If
                End If

                If IsDirty Then

                    Dim response As MsgBoxResult = If(saveRequested, MsgBoxResult.Yes, MsgBox("Save changes to " & FileName & "?", vbYesNoCancel))

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

            Private Function CalculateAndValidateForClose() As CloseModelValidationResult
                If Not CloseValidationRequired Then
                    Abovo.SummitDiagnostics.WriteLine("[Close Validation] model=" & ModelID.ToString() & ", skipped=clean")
                    Return New CloseModelValidationResult()
                End If
                Dim result = CalculateAndValidateChangedModel()
                _knownCloseValidationFailure = result.HasFailures
                If RecordIdleCheckSheetResult(CalculationRevision, result.HasFailures) Then
                    'Close validation already presents its own warning dialog.
                    RecoveryBackupManager.NotifyCheckSheetState(Me, result, Form.ActiveForm)
                End If
                Return result
            End Function

            Private Function CalculateAndValidateChangedModel() As CloseModelValidationResult

                Dim Result As New CloseModelValidationResult()

                If WB Is Nothing Then
                    Result.ValidationError = "The workbook is not available for validation."
                    Return Result
                End If

                Dim PreviousEngineType As CalculationEngineType =
                    WB.Options.CalculationEngineType

                Dim PreviousDontCalcTransactionalDB As Boolean = True

                If WBCalculationService IsNot Nothing Then
                    PreviousDontCalcTransactionalDB =
                        WBCalculationService.DontCalcTDBS
                End If

                Try
                    If WBCalculationService IsNot Nothing Then
                        WBCalculationService.DontCalcTDBS = False
                    End If

                    WB.Options.CalculationEngineType =
                        CalculationEngineType.Recursive

                    EnsureSaveCalculationCurrent()

                Catch ex As Exception
                    Result.ValidationError =
                        "The full workbook calculation failed: " & ex.Message
                    Return Result

                Finally
                    WB.Options.CalculationEngineType = PreviousEngineType

                    If WBCalculationService IsNot Nothing Then
                        WBCalculationService.DontCalcTDBS =
                            PreviousDontCalcTransactionalDB
                    End If
                End Try

                Return ReadCheckSheetValidation()
            End Function

            Friend Function ReadCheckSheetValidation() As CloseModelValidationResult
                'Read cached results only. Callers must first establish a current revision.
                Dim Result As New CloseModelValidationResult()
                If WB Is Nothing Then
                    Result.ValidationError = "The workbook is not available for validation."
                    Return Result
                End If
                If Profile IsNot Nothing AndAlso
                   String.Equals(
                    Profile.ModelType,
                    "AbovoDSA",
                    StringComparison.OrdinalIgnoreCase) Then

                    ValidateDSAClose(Result)
                    Return Result
                End If

                Try
                    Dim ValidationName As DevExpress.Spreadsheet.DefinedName =
                        WB.DefinedNames.GetDefinedName("Outputs_CheckSheet")

                    If ValidationName Is Nothing OrElse
                       ValidationName.Range Is Nothing Then

                        Result.ValidationError =
                            "The workbook does not contain the Outputs_CheckSheet validation range."
                        Return Result
                    End If

                    Dim ValidationRange As DevExpress.Spreadsheet.CellRange =
                        ValidationName.Range

                    If ValidationRange.ColumnCount < 8 Then
                        Result.ValidationError =
                            "The Outputs_CheckSheet validation range does not contain the expected eight columns."
                        Return Result
                    End If

                    Dim ValidationSheet As DevExpress.Spreadsheet.Worksheet =
                        ValidationRange.Worksheet

                    For RowOffset As Integer = 0 To ValidationRange.RowCount - 1

                        Dim SheetRow As Integer =
                            ValidationRange.TopRowIndex + RowOffset

                        Dim FirstColumn As Integer =
                            ValidationRange.LeftColumnIndex

                        'An Excel error is not a temporary financial imbalance.
                        'Accepted balance overrides must not conceal broken formulas.
                        For Each CheckColumn As Integer In New Integer() {1, 3, 4}
                            Dim CheckCell = ValidationSheet.Cells(SheetRow, FirstColumn + CheckColumn)
                            If CheckCell.ModelValue().IsError Then
                                Result.ValidationError = "Check Sheet contains a formula error at " &
                                    CheckCell.GetReferenceA1() & ": " & CheckCell.ModelDisplayText()
                                Return Result
                            End If
                        Next

                        Dim StatusText As String =
                            ValidationSheet.Cells(SheetRow, FirstColumn + 4).
                                ModelDisplayText().Trim()

                        If String.IsNullOrWhiteSpace(StatusText) OrElse
                           String.Equals(
                               StatusText,
                               "OK",
                               StringComparison.OrdinalIgnoreCase) Then

                            Continue For
                        End If

                        Dim issue = New CloseModelValidationIssue With {
                                .CheckRow = SheetRow + 1,
                                .Label = ValidationSheet.Cells(
                                    SheetRow,
                                    FirstColumn).ModelDisplayText().Trim(),
                                .Status = StatusText,
                                .Message = ValidationSheet.Cells(
                                    SheetRow,
                                    FirstColumn + 5).ModelDisplayText().Trim(),
                                .TargetWorksheet = ValidationSheet.Cells(
                                    SheetRow,
                                    FirstColumn + 7).ModelDisplayText().Trim()
                            }
                        If WorkbookIntegritySupport.HasAcceptedOverride(ValidationRange, RowOffset, StatusText) Then
                            Result.OverriddenIssues.Add(issue)
                        Else
                            Result.Issues.Add(issue)
                        End If

                    Next

                Catch ex As Exception
                    Result.ValidationError =
                        "The Check Sheet could not be examined: " & ex.Message
                End Try

                Return Result

            End Function

            Private Sub ValidateDSAClose(
                ByVal Result As CloseModelValidationResult)

                Try
                    Dim CheckTotalName As DevExpress.Spreadsheet.DefinedName =
                        WB.DefinedNames.GetDefinedName("CheckTotal")

                    If CheckTotalName Is Nothing OrElse
                       CheckTotalName.Range Is Nothing Then

                        Result.ValidationError =
                            "The DSA workbook does not contain the CheckTotal validation range."
                        Return
                    End If

                    Dim CheckTotalText As String =
                        CheckTotalName.Range(0, 0).ModelDisplayText().Trim()
                    Dim CheckTotal As Decimal

                    If Not Decimal.TryParse(
                        CheckTotalText,
                        Globalization.NumberStyles.Any,
                        Globalization.CultureInfo.CurrentCulture,
                        CheckTotal) AndAlso
                       Not Decimal.TryParse(
                        CheckTotalText,
                        Globalization.NumberStyles.Any,
                        Globalization.CultureInfo.InvariantCulture,
                        CheckTotal) Then

                        Result.ValidationError =
                            "The DSA CheckTotal value could not be interpreted: " &
                            CheckTotalText
                        Return
                    End If

                    If CheckTotal = 0D Then Return

                    If Not WB.Worksheets.Contains("Check Sheet") Then
                        Result.ValidationError =
                            "The DSA workbook is missing the 'Check Sheet' worksheet."
                        Return
                    End If

                    Dim CheckSheet As DevExpress.Spreadsheet.Worksheet =
                        WB.Worksheets("Check Sheet")

                    'The v6.1500 contract uses rows 8-32 for individual checks
                    'and row 33 for CheckTotal.
                    For SheetRow As Integer = 7 To 31
                        Dim StatusText As String =
                            CheckSheet.Cells(SheetRow, 2).ModelDisplayText().Trim()
                        Dim CheckValueText As String =
                            CheckSheet.Cells(SheetRow, 3).ModelDisplayText().Trim()
                        Dim CheckValue As Decimal
                        Dim HasNonZeroValue As Boolean =
                            Decimal.TryParse(
                                CheckValueText,
                                Globalization.NumberStyles.Any,
                                Globalization.CultureInfo.CurrentCulture,
                                CheckValue) AndAlso CheckValue <> 0D

                        If Not HasNonZeroValue AndAlso
                           (String.IsNullOrWhiteSpace(StatusText) OrElse
                            String.Equals(
                                StatusText,
                                "OK",
                                StringComparison.OrdinalIgnoreCase)) Then Continue For

                        Result.Issues.Add(
                            New CloseModelValidationIssue With {
                                .CheckRow = SheetRow + 1,
                                .Label = CheckSheet.Cells(SheetRow, 0).ModelDisplayText().Trim(),
                                .Status = If(
                                    String.IsNullOrWhiteSpace(StatusText),
                                    CheckValueText,
                                    StatusText),
                                .Message = CheckValueText,
                                .TargetWorksheet =
                                    CheckSheet.Cells(SheetRow, 4).ModelDisplayText().Trim()
                            })
                    Next

                    If Result.Issues.Count = 0 Then
                        Result.ValidationError =
                            "The DSA CheckTotal is non-zero, but no individual " &
                            "Check Sheet row could be identified."
                    End If

                Catch ex As Exception
                    Result.ValidationError =
                        "The DSA Check Sheet could not be examined: " & ex.Message
                End Try

            End Sub

            Private Shared Function BuildCloseValidationMessage(
                ByVal Validation As CloseModelValidationResult) As String

                Dim MessageLines As New List(Of String) From {
                    "The model did not pass its Check Sheet validation.",
                    "",
                    "To protect the original workbook, Summit must save this model " &
                    "as a separate XLSB copy before closing."
                }

                If Not String.IsNullOrWhiteSpace(Validation.ValidationError) Then
                    MessageLines.Add("")
                    MessageLines.Add(Validation.ValidationError)
                End If

                If Validation.Issues.Count > 0 Then
                    MessageLines.Add("")
                    MessageLines.Add("Checks requiring attention:")

                    Dim MaximumDisplayedIssues As Integer = 12

                    For IssueIndex As Integer =
                        0 To Math.Min(
                            Validation.Issues.Count,
                            MaximumDisplayedIssues) - 1

                        Dim Issue As CloseModelValidationIssue =
                            Validation.Issues(IssueIndex)

                        Dim IssueDescription As String = Issue.Label

                        If String.IsNullOrWhiteSpace(IssueDescription) Then
                            IssueDescription = "Check Sheet row " &
                                Issue.CheckRow.ToString()
                        End If

                        If Not String.IsNullOrWhiteSpace(Issue.Message) Then
                            IssueDescription &= ": " & Issue.Message
                        Else
                            IssueDescription &= ": " & Issue.Status
                        End If

                        If Not String.IsNullOrWhiteSpace(Issue.TargetWorksheet) Then
                            IssueDescription &= " [" & Issue.TargetWorksheet & "]"
                        End If

                        MessageLines.Add("- " & IssueDescription)
                    Next

                    If Validation.Issues.Count > MaximumDisplayedIssues Then
                        MessageLines.Add(
                            "- ...and " &
                            (Validation.Issues.Count - MaximumDisplayedIssues).
                                ToString() &
                            " more check(s).")
                    End If
                End If

                MessageLines.Add("")
                MessageLines.Add(
                    "Select OK to choose a new filename, or Cancel to return to the model.")

                Return String.Join(Environment.NewLine, MessageLines)

            End Function

            Friend NotInheritable Class CloseModelValidationResult

                Public ReadOnly Issues As New List(Of CloseModelValidationIssue)()
                Public ReadOnly OverriddenIssues As New List(Of CloseModelValidationIssue)()
                Public ValidationError As String

                Public ReadOnly Property HasFailures As Boolean
                    Get
                        Return Not String.IsNullOrWhiteSpace(ValidationError) OrElse
                               Issues.Count > 0
                    End Get
                End Property

            End Class

            Friend NotInheritable Class CloseModelValidationIssue

                Public CheckRow As Integer
                Public Label As String
                Public Status As String
                Public Message As String
                Public TargetWorksheet As String

            End Class
            Public Sub CloseModel()

                If IsClosing Then Return
                IsClosing = True
                ClearDiscardedSessionCheckSheetWarning()
                RecoveryBackupManager.Forget(Me)
                IdleIntegrityManager.Forget(Me)

                Try
                    ResourceRegistry.ReleaseAll()
                Catch ex As Exception
                    WriteLog("Error releasing registered model resources: " & ex.Message, FileName)
                End Try

                Try
                    If ModelSpreadsheetControl IsNot Nothing Then
                        RemoveHandler ModelSpreadsheetControl.UnhandledException,
                                      AddressOf SSCUnhandledEvent
                        RemoveHandler ModelSpreadsheetControl.ContentChanged, ContentHandler
                        RemoveHandler ModelSpreadsheetControl.DocumentSaved, SavedHandler
                        RemoveHandler ModelSpreadsheetControl.DocumentPropertiesChanged, PropertiesHandler
                        RemoveHandler ModelSpreadsheetControl.CellValueChanged, NativeCellHandler
                        RemoveHandler ModelSpreadsheetControl.RowsInserted, NativeRowsInsertedHandler
                        RemoveHandler ModelSpreadsheetControl.RowsRemoved, NativeRowsRemovedHandler
                        RemoveHandler ModelSpreadsheetControl.ColumnsInserted, NativeColumnsInsertedHandler
                        RemoveHandler ModelSpreadsheetControl.ColumnsRemoved, NativeColumnsRemovedHandler
                        RemoveHandler ModelSpreadsheetControl.SheetInserted, NativeSheetInsertedHandler
                        RemoveHandler ModelSpreadsheetControl.SheetRemoved, NativeSheetRemovedHandler
                        RemoveHandler ModelSpreadsheetControl.SheetRenamed, NativeSheetRenamedHandler
                        RemoveHandler ModelSpreadsheetControl.DefinedNameEdited, NativeNameEditedHandler
                        RemoveHandler ModelSpreadsheetControl.DefinedNameAdded, NativeNameAddedHandler
                        RemoveHandler ModelSpreadsheetControl.DefinedNameDeleted, NativeNameDeletedHandler
                    End If
                Catch ex As Exception
                    WriteLog("Error detaching spreadsheet handlers: " & ex.Message, FileName)
                End Try

                'Interfaces must be detached and disposed while this model and its
                'workbook are still available. DevExpress controls can request one
                'final unbound value while their bindings are being torn down.
                Try
                    If InstanceInterface IsNot Nothing Then InstanceInterface.CloseStandaloneInterfaces()
                Catch ex As Exception
                    WriteLog("Error closing standalone interfaces: " & ex.Message, FileName)
                End Try

                Try
                    If WBInterface IsNot Nothing Then WBInterface.CloseInterfaces()
                Catch ex As Exception
                    WriteLog("Error closing model interfaces: " & ex.Message, FileName)
                End Try

                Try
                    If InterfaceHistory IsNot Nothing Then InterfaceHistory.Dispose()
                Catch ex As Exception
                    WriteLog("Error disposing interface history: " & ex.Message, FileName)
                End Try

                Try
                    If InterfaceDependencies IsNot Nothing Then InterfaceDependencies.Clear()
                Catch ex As Exception
                    WriteLog("Error clearing interface dependencies: " & ex.Message, FileName)
                End Try

                Try
                    If HistoryManager IsNot Nothing Then HistoryManager.CloseForModel()
                Catch ex As Exception
                    WriteLog("Error disposing history manager: " & ex.Message, FileName)
                End Try

                Try
                    If PdfExportManager IsNot Nothing Then PdfExportManager.CloseForModel()
                Catch ex As Exception
                    WriteLog("Error disposing PDF export manager: " & ex.Message, FileName)
                End Try

                Try
                    If ExcelExportManager IsNot Nothing Then ExcelExportManager.CloseForModel()
                Catch ex As Exception
                    WriteLog("Error disposing Excel export manager: " & ex.Message, FileName)
                End Try

                Try
                    If SSViewer IsNot Nothing Then SSViewer.Dispose()
                Catch ex As Exception
                    WriteLog("Error disposing spreadsheet viewer: " & ex.Message, FileName)
                End Try

                Try
                    If ModelSpreadsheetControl IsNot Nothing Then
                        ModelSpreadsheetControl.Dispose()
                    End If
                Catch ex As Exception
                    WriteLog("Error disposing spreadsheet control: " & ex.Message, FileName)
                End Try

                SSViewer = Nothing
                ModelSpreadsheetControl = Nothing
                WB = Nothing
                ExpendAnalyser = Nothing
                ExpendAnalyserV2 = Nothing
                WBInterface = Nothing
                InterfaceHistory = Nothing
                WBData = Nothing
                WBCalcEngine = Nothing
                WBCalculationService = Nothing
                WBStructure = Nothing
                WBStructureManager = Nothing
                WBDataPres = Nothing
                EventCoordinator = Nothing
                If ChangeManager IsNot Nothing Then RemoveHandler ChangeManager.HistoryChanged, AddressOf CommittedMetadataChanged
                ChangeManager = Nothing
                HistoryManager = Nothing
                PdfExportManager = Nothing
                ExcelExportManager = Nothing
                TransDBM = Nothing
                RDSM = Nothing
                InterfaceDependencies = Nothing
                TransDBSync = Nothing
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

            Public Function ProcessModelMetadata() As AbovoTransaction

                Dim Result As New AbovoTransaction

                Try
                    If WBStructure Is Nothing Then
                        Throw New InvalidOperationException(
                            "The workbook interface definition has not been loaded.")
                    End If

                    If Profile Is Nothing Then
                        Throw New InvalidOperationException(
                            "The workbook model profile has not been resolved.")
                    End If

                    Profile.ApplyMetadata(WB, WBStructure)

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

            Public Function ProcessAsAbovoBP() As AbovoTransaction

                Return ProcessModelMetadata()

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

            Try
                ModelToClose.CloseModel()
            Catch ex As Exception
                WriteLog("Error while releasing model " & ModelID.ToString() & ": " & ex.Message)
            Finally
                'Only remove the model after every interface has detached. Keeping
                'the slot alive during disposal prevents late grid callbacks from
                'dereferencing a missing ExcelModels(ModelID).
                If ExcelModels IsNot Nothing AndAlso
                   ModelID >= 0 AndAlso ModelID < ExcelModels.Length Then

                    ExcelModels(ModelID) = Nothing
                End If
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

            InterfaceHistoryCoordinator.NotifyChanged()

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

            Dim Profile As WorkbookModelProfile = Nothing
            Dim Result As AbovoTransaction =
                WorkbookContractValidator.Validate(
                    GetWorkBook(ModelToCheck),
                    Profile)

            If Not Result.BError Then
                ExcelModels(ModelToCheck).Profile = Profile
            End If

            Return Result

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
            Dim modelID As Integer = ResolveSpreadsheetModelID(TryCast(sender, SpreadsheetControl))
            SystemMessageManager.Publish(
                modelID,
                "A spreadsheet command could not be completed: " & My_Exception.Message,
                SystemMessageSeverity.Error,
                "Spreadsheet",
                If(modelID >= 0 AndAlso ExcelModels IsNot Nothing AndAlso
                   modelID < ExcelModels.Length AndAlso ExcelModels(modelID) IsNot Nothing,
                   ExcelModels(modelID).FileName,
                   String.Empty))
            MessageBox.Show(My_Exception.Message, "An error has occured with the file")

        End Sub

        Private Shared Function ResolveSpreadsheetModelID(ByVal control As SpreadsheetControl) As Integer
            If control Is Nothing OrElse ExcelModels Is Nothing Then Return -1
            For index As Integer = 0 To ExcelModels.Length - 1
                If ExcelModels(index) IsNot Nothing AndAlso
                   Object.ReferenceEquals(ExcelModels(index).ModelSpreadsheetControl, control) Then
                    Return index
                End If
            Next
            Return -1
        End Function
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
                'FileInfo timestamps are lazy and loading the workbook can update
                'LastAccessTime. Capture the filesystem value before LoadDocument.
                NewModel.FileInfo.Refresh()
                NewModel.PreviousFileAccessTime = NewModel.FileInfo.LastAccessTime
                If OpenMode = WorkbookOpenMode.FullModel Then NewModel.CaptureCheckSheetDiskBeforeLoad()
                If Not NewModel.ModelSpreadsheetControl.LoadDocument(FullPath) Then
                    Throw New System.IO.InvalidDataException("The spreadsheet engine could not read this workbook. No model was opened.")
                End If
                'Read at open, not from a potentially externally replaced original
                'during a later recovery save. No source workbook changes.
                NewModel.RecoveryHasVerifiedBinaryMetadata = RecoveryXlsmCompatibility.HasVerifiedBinaryProfile(FullPath)

                'Read metadata only; never apply rules or mutate workbook schema on open.
                Try
                    NewModel.ManagedDefinition = ModelManagerStore.ReadEmbedded(FullPath)
                Catch ex As Exception When TypeOf ex Is System.IO.IOException OrElse TypeOf ex Is System.IO.InvalidDataException OrElse TypeOf ex Is System.Xml.XmlException OrElse TypeOf ex Is InvalidOperationException OrElse TypeOf ex Is UnauthorizedAccessException
                    NewModel.ManagedDefinitionWarning = "Model Manager definition was not loaded: " & ex.Message
                    SystemMessageManager.Publish(NewModelID, NewModel.ManagedDefinitionWarning, SystemMessageSeverity.Warning, "Model Manager")
                End Try

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

                    Dim StructureSource As String =
                        NewModel.Profile.ResolveStructureSource(FullPath)

                    Dim StructureResult As AbovoTransaction =
                        NewModel.WBStructureManager.CreateStructureFromXML(
                            StructureSource)
                    If StructureResult.BError Then
                        Throw New InvalidOperationException(StructureResult.StringReturn)
                    End If

                    Dim MetadataResult As AbovoTransaction =
                        NewModel.ProcessModelMetadata()
                    If MetadataResult.BError Then
                        Throw New InvalidOperationException(MetadataResult.StringReturn)
                    End If

                    NewModel.ModelSpreadsheetControl.Dock = DockStyle.Fill
                    NewModel.WBCalcEngine.CalcManual()
                    NewModel.WBCalcEngine.ChainCalc()
                    'Do not rely on Excel VBA being enabled to refresh saved caches.
                    NewModel.RestoreDeferredSaveState()
                    LoadedModelType = ContractResult.StringReturn
                End If

                OpenModelCount += 1
                NewModel.ModelSpreadsheetControl.Modified = False
                NewModel.IsDirty = False 'Loading/presentation setup is not a client edit.
                If OpenMode = WorkbookOpenMode.FullModel Then
                    NewModel.RecoverySourcePath = RecoveryBackupStore.RecoveryOrigin(FullPath)
                    If Not String.IsNullOrWhiteSpace(NewModel.RecoverySourcePath) Then
                        NewModel.RecoverySaveAsRequired = True
                        NewModel.IsDirty = True
                        Try
                            NewModel.ChangeManager.RestoreRecoveryHistory(RecoveryHistoryStore.Read(FullPath, NewModel.ChangeManager.GetHistoryTable()))
                        Catch historyError As Exception
                            SystemMessageManager.Publish(NewModelID, "The recovery workbook opened, but its prior-session history could not be read: " & historyError.Message, SystemMessageSeverity.Warning, "Recovery backup", FullPath)
                        End Try
                        SystemMessageManager.Publish(NewModelID, "RECOVERY COPY: use Save As to save an XLSB business plan. Suggested original: " & NewModel.RecoverySourcePath & ". The original has not been overwritten.", SystemMessageSeverity.Warning, "Recovery backup", FullPath)
                    End If
                    RecoveryBackupManager.Track(NewModel)
                    NewModel.CaptureCheckSheetOpenState()
                    IdleIntegrityManager.Track(NewModel)
                    InternalFileState = 2
                    ApplicationConfiguration.ActiveModelID = NewModelID
                End If

                InterfaceHistoryCoordinator.NotifyChanged()

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
