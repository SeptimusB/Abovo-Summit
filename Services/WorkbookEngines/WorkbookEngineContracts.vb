Option Strict On

Imports System.Configuration
Imports System.Collections.ObjectModel
Imports System.Threading

Namespace Abovo.WorkbookEngines
    'Diagnostic identity only, never permission to terminate a process. Start
    'time prevents a recycled Windows PID being mistaken for our former owner.
    Public NotInheritable Class WorkbookNativeProcessIdentity
        Public ReadOnly Property ProcessId As Integer
        Public ReadOnly Property StartedUtc As DateTime
        Friend Sub New(processId As Integer, startedUtc As DateTime)
            Me.ProcessId = processId : Me.StartedUtc = startedUtc
        End Sub
    End Class

    ' Production callers remain read-only. Isolated value-edit trials require
    ' explicit opt-in. Candidate exports never replace a source or mark it saved.
    Public Enum WorkbookEnginePreference
        Automatic = 0
        DevExpressOnly = 1
        ExcelRequired = 2
    End Enum

    Public Enum WorkbookCalculationKind
        Incremental = 0
        Full = 1
        Rebuild = 2
    End Enum

    Public NotInheritable Class WorkbookEngineOptions
        Public ReadOnly Property Preference As WorkbookEnginePreference
        Public ReadOnly Property AllowTrustedVba As Boolean
        Public ReadOnly Property RequireSummitFunctions As Boolean
        Public ReadOnly Property OperationTimeoutMilliseconds As Integer
        Public ReadOnly Property EnableValueEditTrial As Boolean
        Public ReadOnly Property EnableCandidateSaveTrial As Boolean
        Public ReadOnly Property EnablePublicationTrial As Boolean

        Public Sub New(Optional preference As WorkbookEnginePreference = WorkbookEnginePreference.Automatic,
                       Optional allowTrustedVba As Boolean = True,
                       Optional requireSummitFunctions As Boolean = True,
                       Optional operationTimeoutMilliseconds As Integer = 120000,
                       Optional enableValueEditTrial As Boolean = False,
                       Optional enableCandidateSaveTrial As Boolean = False,
                       Optional enablePublicationTrial As Boolean = False)
            If Not [Enum].IsDefined(GetType(WorkbookEnginePreference), preference) Then Throw New ArgumentOutOfRangeException(NameOf(preference))
            If operationTimeoutMilliseconds < 50 OrElse operationTimeoutMilliseconds > 1800000 Then Throw New ArgumentOutOfRangeException(NameOf(operationTimeoutMilliseconds))
            Me.Preference = preference
            Me.AllowTrustedVba = allowTrustedVba
            Me.RequireSummitFunctions = requireSummitFunctions
            Me.OperationTimeoutMilliseconds = operationTimeoutMilliseconds
            Me.EnableValueEditTrial = enableValueEditTrial
            Me.EnableCandidateSaveTrial = enableCandidateSaveTrial
            If enablePublicationTrial AndAlso Not enableCandidateSaveTrial Then Throw New ArgumentException("Publication requires candidate-save opt-in.")
            Me.EnablePublicationTrial = enablePublicationTrial
        End Sub
    End Class

    Friend NotInheritable Class WorkbookEngineSettings
        Inherits ApplicationSettingsBase

        <UserScopedSetting(), DefaultSettingValue("Automatic")>
        Public Property Preference As String
            Get
                Return CStr(Me(NameOf(Preference)))
            End Get
            Set(value As String)
                Me(NameOf(Preference)) = value
            End Set
        End Property

        Friend Function ReadOptions() As WorkbookEngineOptions
            Dim selected As WorkbookEnginePreference
            If Not [Enum].TryParse(Preference, True, selected) OrElse Not [Enum].IsDefined(GetType(WorkbookEnginePreference), selected) Then
                selected = WorkbookEnginePreference.Automatic
            End If
            Return New WorkbookEngineOptions(selected)
        End Function
    End Class

    Public NotInheritable Class WorkbookReadArea
        Public ReadOnly Property Worksheet As String
        Public ReadOnly Property Row As Integer
        Public ReadOnly Property Column As Integer
        Public ReadOnly Property Rows As Integer
        Public ReadOnly Property Columns As Integer

        Public Sub New(worksheet As String, row As Integer, column As Integer, rows As Integer, columns As Integer)
            If String.IsNullOrWhiteSpace(worksheet) Then Throw New ArgumentException("A worksheet is required.", NameOf(worksheet))
            If row < 0 OrElse column < 0 OrElse rows < 1 OrElse columns < 1 OrElse
               CLng(row) + rows > 1048576 OrElse CLng(column) + columns > 16384 OrElse
               CLng(rows) * columns > 100000 Then Throw New ArgumentOutOfRangeException(NameOf(rows), "Read exceeds the bounded worksheet transport limits.")
            Me.Worksheet = worksheet : Me.Row = row : Me.Column = column : Me.Rows = rows : Me.Columns = columns
        End Sub

        Public ReadOnly Property Address As String
            Get
                Return ColumnName(Column) & (Row + 1).ToString(Globalization.CultureInfo.InvariantCulture) & ":" &
                       ColumnName(Column + Columns - 1) & (Row + Rows).ToString(Globalization.CultureInfo.InvariantCulture)
            End Get
        End Property

        Private Shared Function ColumnName(index As Integer) As String
            Dim result As String = ""
            index += 1
            Do While index > 0
                index -= 1
                result = ChrW(AscW("A"c) + index Mod 26) & result
                index \= 26
            Loop
            Return result
        End Function
    End Class

    Public NotInheritable Class WorkbookCellError
        Public ReadOnly Property Text As String
        Public Sub New(text As String)
            If String.IsNullOrEmpty(text) OrElse Not text.StartsWith("#", StringComparison.Ordinal) Then Throw New ArgumentException("Expected an Excel error.", NameOf(text))
            Me.Text = text
        End Sub
        Public Overrides Function ToString() As String
            Return Text
        End Function
    End Class

    Public NotInheritable Class WorkbookValueBlock
        Private ReadOnly data As Object(,)
        Public ReadOnly Property Area As WorkbookReadArea
        Public Sub New(area As WorkbookReadArea, values As Object(,))
            If area Is Nothing Then Throw New ArgumentNullException(NameOf(area))
            If values Is Nothing OrElse values.GetLowerBound(0) <> 0 OrElse values.GetLowerBound(1) <> 0 OrElse
               values.GetLength(0) <> area.Rows OrElse values.GetLength(1) <> area.Columns Then Throw New ArgumentException("Value rectangle differs from the request.")
            Me.Area = area
            data = DirectCast(values.Clone(), Object(,))
            For Each value In data
                If value IsNot Nothing AndAlso Not TypeOf value Is String AndAlso Not TypeOf value Is Boolean AndAlso
                   Not TypeOf value Is Double AndAlso Not TypeOf value Is WorkbookCellError Then Throw New ArgumentException("Only typed, detached spreadsheet values may cross the engine boundary.")
            Next
        End Sub
        Public Function ValueAt(row As Integer, column As Integer) As Object
            Return data(row, column)
        End Function
    End Class

    Public NotInheritable Class WorkbookCalculationResult
        Public ReadOnly Property SessionId As Guid
        Public ReadOnly Property Revision As Long
        Public ReadOnly Property SourceHash As String
        Public ReadOnly Property Engine As String
        Public ReadOnly Property CalculationMilliseconds As Long
        Public ReadOnly Property TransferMilliseconds As Long
        Public ReadOnly Property Blocks As ReadOnlyCollection(Of WorkbookValueBlock)
        Public ReadOnly Property CalculationGeneration As Long
        Public ReadOnly Property Presentation As ReadOnlyCollection(Of WorkbookPresentationBlock)
        Friend Sub New(sessionId As Guid, revision As Long, sourceHash As String, engine As String,
                       calculationMs As Long, transferMs As Long, blocks As IList(Of WorkbookValueBlock),
                       generation As Long, Optional presentation As IList(Of WorkbookPresentationBlock) = Nothing)
            Me.SessionId = sessionId : Me.Revision = revision : Me.SourceHash = sourceHash : Me.Engine = engine
            CalculationMilliseconds = calculationMs : TransferMilliseconds = transferMs
            Me.Blocks = New List(Of WorkbookValueBlock)(blocks).AsReadOnly()
            CalculationGeneration = generation
            Me.Presentation = New List(Of WorkbookPresentationBlock)(If(presentation, New List(Of WorkbookPresentationBlock)())).AsReadOnly()
        End Sub
    End Class

    ' Implementations and their native objects are confined to one STA owner.
    ' Consumer code must use WorkbookCalculationSession, never retain a backend.
    Public Interface IWorkbookCalculationBackend
        Inherits IDisposable
        ReadOnly Property Name As String
        ReadOnly Property Version As String
        Sub OpenReadOnly(path As String, options As WorkbookEngineOptions)
        Sub Calculate(kind As WorkbookCalculationKind)
        Function Read(area As WorkbookReadArea) As WorkbookValueBlock
    End Interface
End Namespace
