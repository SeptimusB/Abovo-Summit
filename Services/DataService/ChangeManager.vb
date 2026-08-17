
Imports Abovo
Imports Abovo.FileManager
Imports Abovo.LogDebugDev
Imports Abovo.AbovoAppCls
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils.Extensions
Imports DevExpress.XtraEditors.Repository
Imports System.Globalization
Namespace Abovo
    Public Class MasterChangeLog

        Public Shared ChangeLogEventID As Integer = -1
        Public Shared ChangeLog As DataTable

        Public Shared Sub AddChangeLogEvent(SentChangeLogEvent As ChangeLogEvent)

            ChangeLogEventID += 1

            ChangeLog.Rows.Add(New Object() {ChangeLogEventID, SentChangeLogEvent.TimeStamp, SentChangeLogEvent.ModelID, SentChangeLogEvent.Description, SentChangeLogEvent.WSName, SentChangeLogEvent.CellAddress, SentChangeLogEvent.OriginalValue, SentChangeLogEvent.ChangedValue, SentChangeLogEvent.UserName, SentChangeLogEvent.Status, SentChangeLogEvent.DataType})

        End Sub

        Public Shared Sub Initialise()

            ChangeLog = New DataTable("MasterChangeLog")

            ChangeLog.Columns.AddRange(New DataColumn() {
                New DataColumn("EventID", GetType(Integer)),
                New DataColumn("TimeStamp", GetType(DateTime)),
                New DataColumn("ModelID", GetType(Integer)),
                New DataColumn("Description", GetType(String)),
                New DataColumn("WSName", GetType(String)),
                New DataColumn("CellAddress", GetType(String)),
                New DataColumn("OriginalValue", GetType(String)),
                New DataColumn("ChangedValue", GetType(String)),
                New DataColumn("UserName", GetType(String)),
                New DataColumn("Status", GetType(Integer)),
                New DataColumn("DataType", GetType(String))
                            })
        End Sub



    End Class

    Public Structure ChangeLogEvent

        Public EventID As Integer
        Public ModelID As Integer
        Public Description As String
        Public WSName As String
        Public CellAddress As String
        Public OriginalValue As String
        Public ChangedValue As String
        Public TimeStamp As DateTime
        Public UserName As String

        '0 = Unprocessed, 1 = Processed, 2 = Rejected, 3 = Error, 4 = undone, 5 = Redo, 6 = NoticeOnly
        Public Status As Integer
        Public DataType As String

    End Structure

    Public Class ModelChangeManager

        Private WB As DevExpress.Spreadsheet.IWorkbook
        Public ModelID As Integer
        Public Sub New(ByRef SetModelID As Integer)

            ModelID = SetModelID
            WB = ExcelModels(ModelID).WB

            MasterChangeLog.AddChangeLogEvent(New ChangeLogEvent With {
                .ModelID = ModelID,
                .Description = "File " & WB.Path & " opened",
                .WSName = "System Message",
                .CellAddress = "",
                .OriginalValue = "",
                .ChangedValue = "",
                .TimeStamp = Now(),
                .UserName = Environment.UserName,
                .Status = 6
            })

        End Sub


        Public Function ProcessChange(SentDataChangeEvent As DataChangeEvent) As AbovoTransaction

            Dim ChangeTransaction As New AbovoTransaction
            Dim TargetCell As DevExpress.Spreadsheet.Cell = Nothing
            Dim OriginalCellValue As DevExpress.Spreadsheet.CellValue = DevExpress.Spreadsheet.CellValue.Empty
            Dim HasOriginalCellValue As Boolean = False

            Dim CLE As New ChangeLogEvent With {
                .ModelID = ModelID,
                .Description = SentDataChangeEvent.Description,
                .WSName = NormalizeIdentifier(SentDataChangeEvent.WSName),
                .CellAddress = NormalizeIdentifier(SentDataChangeEvent.CellAddress),
                .OriginalValue = Convert.ToString(SentDataChangeEvent.OriginalValue, CultureInfo.CurrentCulture),
                .ChangedValue = Convert.ToString(SentDataChangeEvent.ChangedValue, CultureInfo.CurrentCulture),
                .TimeStamp = SentDataChangeEvent.TimeStamp,
                .UserName = SentDataChangeEvent.UserName,
                .DataType = SentDataChangeEvent.DataFormat
            }

            Try
                TargetCell = WB.Worksheets(CLE.WSName).Cells(CLE.CellAddress)
                OriginalCellValue = TargetCell.Value
                HasOriginalCellValue = True
                WriteTypedValue(TargetCell, SentDataChangeEvent.ChangedValue, SentDataChangeEvent.DataFormat)

            Catch ex As Exception

                If TargetCell IsNot Nothing AndAlso HasOriginalCellValue Then
                    Try
                        TargetCell.Value = OriginalCellValue
                    Catch
                        'Preserve the original write error in the transaction.
                    End Try
                End If

                MsgBox("Error processing change event for ModelID " & ModelID & " on cell " & CLE.WSName & "!" & CLE.CellAddress & " - " & ex.Message, MsgBoxStyle.Critical)
                SystemLog("Error processing change event for ModelID " & ModelID & " on cell " & CLE.WSName & "!" & CLE.CellAddress & " - " & ex.Message)
                CLE.Status = 3
                ChangeTransaction.BError = True
                ChangeTransaction.BSuccess = False
                ChangeTransaction.StrResponseMessage = "Error processing change event for ModelID " & ModelID & " on cell " & CLE.WSName & "!" & CLE.CellAddress & " - " & ex.Message
                MasterChangeLog.AddChangeLogEvent(CLE)
                Return ChangeTransaction

            End Try

            CLE.Status = 1
            ChangeTransaction.BError = False
            ChangeTransaction.BSuccess = True
            MasterChangeLog.AddChangeLogEvent(CLE)
            ExcelModels(ModelID).IsDirty = True
            ExcelModels(ModelID).WBCalcEngine.CalculateWSs()

            Return ChangeTransaction

        End Function

        Public Function ProcessChangeByNRAddressing(SentDataChangeEvent As DataChangeEvent) As AbovoTransaction

            Dim ChangeTransaction As New AbovoTransaction
            Dim TargetRange As DevExpress.Spreadsheet.CellRange = Nothing
            Dim TargetCell As DevExpress.Spreadsheet.Cell = Nothing
            Dim OriginalCellValue As DevExpress.Spreadsheet.CellValue = DevExpress.Spreadsheet.CellValue.Empty
            Dim HasOriginalCellValue As Boolean = False
            Dim CLE As New ChangeLogEvent With {
                .ModelID = ModelID,
                .Description = SentDataChangeEvent.Description,
                .OriginalValue = Convert.ToString(SentDataChangeEvent.OriginalValue, CultureInfo.CurrentCulture),
                .ChangedValue = Convert.ToString(SentDataChangeEvent.ChangedValue, CultureInfo.CurrentCulture),
                .TimeStamp = SentDataChangeEvent.TimeStamp,
                .UserName = SentDataChangeEvent.UserName,
                .DataType = SentDataChangeEvent.DataFormat
            }

            Try
                TargetRange = WB.Range(NormalizeIdentifier(SentDataChangeEvent.TargetNR))

                If SentDataChangeEvent.NROrientation = Orientation.Horizontal Then
                    TargetCell = TargetRange(0, SentDataChangeEvent.TargetNRIndex)
                Else
                    TargetCell = TargetRange(SentDataChangeEvent.TargetNRIndex, 0)
                End If

                CLE.WSName = TargetRange.Worksheet.Name
                CLE.CellAddress = TargetCell.GetReferenceA1
                OriginalCellValue = TargetCell.Value
                HasOriginalCellValue = True
                WriteTypedValue(TargetCell, SentDataChangeEvent.ChangedValue, SentDataChangeEvent.DataFormat)

            Catch ex As Exception

                If TargetCell IsNot Nothing AndAlso HasOriginalCellValue Then
                    Try
                        TargetCell.Value = OriginalCellValue
                    Catch
                        'Preserve the original write error in the transaction.
                    End Try
                End If

                Dim TargetDescription As String = If(String.IsNullOrEmpty(CLE.WSName), NormalizeIdentifier(SentDataChangeEvent.TargetNR), CLE.WSName & "!" & CLE.CellAddress)
                MsgBox("Error processing change event for ModelID " & ModelID & " on " & TargetDescription & " - " & ex.Message, MsgBoxStyle.Critical)
                SystemLog("Error processing change event for ModelID " & ModelID & " on " & TargetDescription & " - " & ex.Message)
                CLE.Status = 3
                ChangeTransaction.BError = True
                ChangeTransaction.BSuccess = False
                ChangeTransaction.StrResponseMessage = "Error processing change event for ModelID " & ModelID & " on " & TargetDescription & " - " & ex.Message
                MasterChangeLog.AddChangeLogEvent(CLE)
                Return ChangeTransaction

            End Try

            CLE.Status = 1
            ChangeTransaction.BError = False
            ChangeTransaction.BSuccess = True
            MasterChangeLog.AddChangeLogEvent(CLE)
            ExcelModels(ModelID).IsDirty = True
            ExcelModels(ModelID).WBCalcEngine.CalculateWSs()

            Return ChangeTransaction

        End Function

        Private Shared Function NormalizeIdentifier(Value As String) As String
            Return If(Value Is Nothing, Nothing, Value.Trim())
        End Function

        Private Shared Sub WriteTypedValue(TargetCell As DevExpress.Spreadsheet.Cell,
                                           ChangedValue As Object,
                                           DataFormat As String)

            If ChangedValue Is Nothing OrElse Convert.IsDBNull(ChangedValue) Then
                TargetCell.ClearContents()
                Return
            End If

            Dim FormatCode As String = If(DataFormat, String.Empty).Trim().ToUpperInvariant()

            Select Case FormatCode
                Case "S", "FL", "DUMMY", String.Empty
                    TargetCell.Value = Convert.ToString(ChangedValue, CultureInfo.CurrentCulture)
                Case "B"
                    TargetCell.Value = If(ConvertToBoolean(ChangedValue), 1, 0)
                Case "I", "Y"
                    TargetCell.Value = ConvertToInteger(ChangedValue)
                Case "D", "DM"
                    TargetCell.Value = DevExpress.Spreadsheet.CellValue.FromObject(ConvertToDateOrSerial(ChangedValue))
                Case "N", "P", "C", "M", "SM", "R"
                    TargetCell.Value = ConvertToDouble(ChangedValue)
                Case Else
                    'Unknown structure formats retain the caller's runtime type.
                    TargetCell.Value = DevExpress.Spreadsheet.CellValue.FromObject(ChangedValue)
            End Select
        End Sub

        Private Shared Function ConvertToBoolean(Value As Object) As Boolean
            If TypeOf Value Is Boolean Then Return DirectCast(Value, Boolean)

            Dim Text As String = Convert.ToString(Value, CultureInfo.CurrentCulture).Trim()
            Dim ParsedBoolean As Boolean
            If Boolean.TryParse(Text, ParsedBoolean) Then Return ParsedBoolean

            Select Case Text.ToUpperInvariant()
                Case "Y", "YES"
                    Return True
                Case "N", "NO"
                    Return False
            End Select

            Return ConvertToDouble(Value) <> 0
        End Function

        Private Shared Function ConvertToInteger(Value As Object) As Integer
            If Not TypeOf Value Is String AndAlso TypeOf Value Is IConvertible Then
                Return Convert.ToInt32(Value, CultureInfo.InvariantCulture)
            End If

            Dim ParsedValue As Integer
            Dim Text As String = Convert.ToString(Value, CultureInfo.CurrentCulture).Trim()
            If Integer.TryParse(Text, NumberStyles.Integer Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, ParsedValue) Then Return ParsedValue
            If Integer.TryParse(Text, NumberStyles.Integer Or NumberStyles.AllowThousands, CultureInfo.InvariantCulture, ParsedValue) Then Return ParsedValue

            Throw New FormatException("'" & Text & "' is not a valid whole number.")
        End Function

        Private Shared Function ConvertToDouble(Value As Object) As Double
            If Not TypeOf Value Is String AndAlso TypeOf Value Is IConvertible Then
                Return Convert.ToDouble(Value, CultureInfo.InvariantCulture)
            End If

            Dim ParsedValue As Double
            Dim Text As String = Convert.ToString(Value, CultureInfo.CurrentCulture).Trim()
            If Double.TryParse(Text, NumberStyles.Any, CultureInfo.CurrentCulture, ParsedValue) Then Return ParsedValue
            If Double.TryParse(Text, NumberStyles.Any, CultureInfo.InvariantCulture, ParsedValue) Then Return ParsedValue

            Throw New FormatException("'" & Text & "' is not a valid number.")
        End Function

        Private Shared Function ConvertToDateOrSerial(Value As Object) As Object
            If TypeOf Value Is DateTime Then Return DirectCast(Value, DateTime)

            If Not TypeOf Value Is String AndAlso TypeOf Value Is IConvertible Then
                Return Convert.ToDouble(Value, CultureInfo.InvariantCulture)
            End If

            Dim Text As String = Convert.ToString(Value, CultureInfo.CurrentCulture).Trim()
            Dim SerialValue As Double
            If Double.TryParse(Text, NumberStyles.Any, CultureInfo.CurrentCulture, SerialValue) Then Return SerialValue
            If Double.TryParse(Text, NumberStyles.Any, CultureInfo.InvariantCulture, SerialValue) Then Return SerialValue

            Dim ParsedDate As DateTime
            If DateTime.TryParse(Text, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, ParsedDate) Then Return ParsedDate
            If DateTime.TryParse(Text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, ParsedDate) Then Return ParsedDate

            Throw New FormatException("'" & Text & "' is not a valid date.")
        End Function

    End Class

    '        'Implements IEnumerable

    '        'Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
    '        '    Return ChangeEvents
    '        'End Function
    '        'Private ReadOnly _ModelID As Integer
    '        'Private ReadOnly _TimeStamp As DateTime
    '        'Private ReadOnly _EventID As Integer
    '        'Private ReadOnly _Description As String
    '        'Private ReadOnly _ModelEventID As Integer

    '        'Private _WS As DevExpress.Spreadsheet.Worksheet

    '        'Private _OriginalValue As CellValue
    '        'Private _ChangedValue As CellValue

    '        'Public Sub New(SetEventID)

    '        '    _EventID = SetEventID
    '        '    _TimeStamp = Now()
    '        '    _ModelEventID = ModelChangeLogs(_ModelID).GetModelEventID

    '        'End Sub

    'Public Class ModelChangeManager

    '    Public HasChanges As Boolean
    '    Public ChangeCount As Integer = -1

    '    Public CLog As ModelChangeLog
    '    Private ModelID As Integer
    '    Private MyWorkbook As DevExpress.Spreadsheet.Workbook
    '    Private ModelLogCount As Integer = -1
    '    Public IsDirty As Boolean = False
    '    Public Sub New(SetModelID As Integer)

    '        ModelID = SetModelID
    '        'MyWorkbook = FileManager.GetWorkBook(ModelID)
    '        CLog = New ModelChangeLog(SetModelID)

    '    End Sub

    '    Public Class ChangeLog

    '        Public ModelID As Integer
    '        Public Sub New(SetModelID As Integer)

    '            ModelID = SetModelID
    '            'MyWorkbook = FileManager.GetWorkBook(ModelID)
    '            CLog = New ModelChangeLog(SetModelID)

    '        End Sub

    '        'Implements IEnumerable
    '        'Public ChangeEvents As New List(Of ChangeEvent)
    '        'Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
    '        '    For Each CE As ChangeEvent In ChangeEvents

    '        '        CE.Yield

    '        '    Next
    '        'End Function

    '        'Private Iterator Function AnimalsForType(ByVal ModelID As Integer) As IEnumerable
    '        '    For Each CE As ChangeEvent In ChangeEvents
    '        '        If (CE.ModelID = ModelID) Then
    '        '            Yield CE
    '        '        End If
    '        '    Next
    '        'End Function

    '    End Class

    '    Public ChangeLogs As New Dictionary(Of Integer, ChangeLog)




    '    Public Sub ExecuteChange(ChangeEvent As ChangeEvent)

    '        CLog.AddChange(ChangeEvent)
    '        HasChanges = True
    '        IsDirty = True

    '    End Sub
    '    Public Function ModelSaved() As Boolean

    '        IsDirty = False
    '        Return True

    '    End Function
    '    Public Function UndoChange(ChangeEventID As Integer) As Boolean




    '        Return True
    '    End Function

    '    Public Class ModelChangeLog

    '        Private ReadOnly _ModelID
    '        Private _IsClosed As Boolean = False
    '        Public EventIndex As Integer = -1
    '        Public ChangeCount As Integer = 0
    '        Public Sub New(SetModelID As Integer)

    '            _ModelID = SetModelID

    '        End Sub
    '        Public Function AddChange(NewChangeEvent As ChangeEvent) As Integer
    '            If _IsClosed = False Then
    '                ChangeCount += 1
    '                'ModelChangeLogs(_ModelID).ChangeEvents.Add(NewChangeEvent)
    '                Return ChangeCount
    '            Else
    '                Return -1
    '            End If
    '        End Function
    '        Public Function GetModelEventID() As Integer

    '            EventIndex += 1
    '            Return EventIndex

    '        End Function

    '        Public Sub CloseLog()

    '            _IsClosed = True

    '        End Sub

    '    End Class

    '    Public Class ChangeEvent

    '        'Implements IEnumerable

    '        'Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
    '        '    Return ChangeEvents
    '        'End Function
    '        'Private ReadOnly _ModelID As Integer
    '        'Private ReadOnly _TimeStamp As DateTime
    '        'Private ReadOnly _EventID As Integer
    '        'Private ReadOnly _Description As String
    '        'Private ReadOnly _ModelEventID As Integer

    '        'Private _WS As DevExpress.Spreadsheet.Worksheet

    '        'Private _OriginalValue As CellValue
    '        'Private _ChangedValue As CellValue

    '        'Public Sub New(SetEventID)

    '        '    _EventID = SetEventID
    '        '    _TimeStamp = Now()
    '        '    _ModelEventID = ModelChangeLogs(_ModelID).GetModelEventID

    '        'End Sub
    '        'Public ReadOnly Property ModelID As String

    '        '    Get
    '        '        Return _ModelID
    '        '    End Get

    '        'End Property

    '        'Public ReadOnly Property Description As String

    '        '    Get
    '        '        Return _Description
    '        '    End Get

    '        'End Property

    '        'Public ReadOnly Property WS As DevExpress.Spreadsheet.Worksheet

    '        '    Get
    '        '        Return _WS
    '        '    End Get

    '        'End Property

    '        'Public ReadOnly Property EventID As Integer

    '        '    Get
    '        '        Return _EventID
    '        '    End Get

    '        'End Property
    '        'Public Property OriginalValue As CellValue

    '        '    Get
    '        '        Return _OriginalValue
    '        '    End Get

    '        '    Set(value As CellValue)
    '        '        _OriginalValue = value
    '        '    End Set

    '        'End Property

    '        'Public Property ChangedValue As CellValue

    '        '    Get
    '        '        Return _ChangedValue
    '        '    End Get

    '        '    Set(value As CellValue)
    '        '        _ChangedValue = value
    '        '    End Set

    '        'End Property

    '        'Public Sub PublishChangeEvent()

    '        '    'xxxxxxxxxxxxxx

    '        'End Sub

    '    End Class

    'End Class






End Namespace

