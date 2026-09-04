Imports Abovo.DSAImport
Imports Abovo.EventManager
Imports Abovo.FormulaGeneration
Imports Abovo.ImportModels
Imports Abovo.FileManager
Imports Abovo.AbovoAppCls
Imports Abovo.PresentationManager
Imports Abovo.SpecificRowColumnEvents

Namespace Abovo
    Public Class EventManager

        Private ModelType As String
        Private ModelID As Integer
        Private Params As Object()
        Private IsCachingEvents As Boolean = False
        Public Sub New(SetModelID As Integer, ModelType As String)

            ModelID = SetModelID
            ModelType = ModelType

        End Sub

        Public Function TriggerEvent(ByVal EventType As String, EventParams As Object, ActioningForm As Form) As AbovoTransaction

            Dim EventTransaction As New AbovoTransaction("TriggerEvent")

            If EventType = "Link" Then

                Dim LinkTag As ElementInterfaceLinkTag = TryCast(EventParams, ElementInterfaceLinkTag)

                If LinkTag Is Nothing Then

                    MsgBox("Sorry, this link is not properly configured")
                    EventTransaction.BError = True
                    EventTransaction.EventCancelled = True
                    Return EventTransaction

                End If

                Dim SetGSID As Integer = -1

                If LinkTag.LinkGroup Is Nothing Then

                    SetGSID = LinkTag.LinkGroupID

                Else

                    SetGSID = GetGroupID(ModelID, LinkTag.LinkGroup)

                End If

                Dim FileInstance As FileInstanceInterface = ExcelModels(ModelID).InstanceInterface

                FileManager.ExcelModels(ModelID).WBInterface.ShowGroupInterface(ModelID, SetGSID, "Maximised", LinkTag.LinkData, FileInstance, LinkTag)

            ElseIf EventType = "Code" Then

                Dim BPE As New BPCodeEvent(
                    ModelID,
                    EventParams,
                    EventTransaction,
                    ActioningForm)
                Return BPE.EventTransaction

            ElseIf EventType = "RowColEvent" Then

                Dim RCE As RowColEventTag = TryCast(EventParams, RowColEventTag)

                If RCE Is Nothing Then

                    EventTransaction.BError = True
                    EventTransaction.EventCancelled = True
                    MsgBox("Sorry, this event is not properly configured")
                    Return EventTransaction

                End If

                ProcessSpecificRowColumnEvents(ModelID, RCE)

            ElseIf EventType = "GridButton" Then

                Dim BPE As BPGridEvent = New BPGridEvent(ModelID, EventParams, EventTransaction, ActioningForm)

            End If

            Return EventTransaction

        End Function

        Private Class BPCodeEvent

            Private ModelID As Integer
            Private EventName As String
            Public EventTransaction As AbovoTransaction


            Public Sub New(SetModelID As Integer,
                           SetEventParams As Object,
                           SetTransaction As AbovoTransaction,
                           ActioningForm As Form)

                ModelID = SetModelID
                EventTransaction = SetTransaction
                ProcessEvent(SetEventParams, ActioningForm)

            End Sub
            Public Sub ProcessEvent(SetEventParams As Object, ActioningForm As Form)

                Select Case SetEventParams

                    Case "FormulaGeneration"
                        ExecuteFormulaGeneration(ModelID, True, EventTransaction)

                    Case "ImportSingleDSA_File"
                        EventTransaction = DSAImport.ImportSingleDSA_File(ModelID)

                    Case "ImportConsolDSA_File"
                        EventTransaction = DSAImport.ImportConsolDSA_File(ModelID)

                    Case "ImportMultiDSA_Files"
                        EventTransaction = DSAImport.DSA_Folder(ModelID)

                    Case "ImportDSA_Template"
                        EventTransaction = DSAImport.ImportDSA_Template(ModelID)

                    Case "ImportStockRentModel"
                        EventTransaction = ImportModels.ImportStockRentModel(ModelID)

                    Case "ImportStockConditionSurvey"
                        EventTransaction = ImportModels.ImportStockConditionSurvey(ModelID)

                    Case "ImportManagementServiceCosts"
                        EventTransaction = ImportModels.ImportManagementServiceCosts(ModelID)

                End Select

            End Sub

        End Class

        Private Class BPGridEvent

            Private ModelID As Integer
            Private GridTag As AttachedGridCommandButton
            Public GridTransaction As AbovoTransaction
            Public Sub New(SetModelID As Integer, SetGridTag As Object, SetGridTransaction As AbovoTransaction, ActioningForm As Form)

                ModelID = SetModelID
                GridTag = TryCast(SetGridTag, AttachedGridCommandButton)

                GridTransaction = SetGridTransaction

                If GridTag Is Nothing Then
                    GridTransaction.BError = True
                    GridTransaction.EventCancelled = True
                    GridTransaction.StringReturn = "The grid button event is not properly configured."
                    Exit Sub
                End If

                ProcessEvent(ActioningForm)

            End Sub
            Public Sub ProcessEvent(ActioningForm As Form)

                Select Case GridTag.CommandData

                    Case "ProcessAddOFARecords"
                        SpecificRowColumnEvents.InsertOFAColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessDeleteOFARecords"
                        SpecificRowColumnEvents.DeleteOFAColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessAddHARecords"
                        SpecificRowColumnEvents.InsertHAColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessDeleteHARecords"
                        SpecificRowColumnEvents.DeleteHAColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessAddCapExRecords"
                        SpecificRowColumnEvents.InsertCapExColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessDeleteCapExRecords"
                        SpecificRowColumnEvents.DeleteCapExColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessAddCapGrantRecords"
                        SpecificRowColumnEvents.InsertCapGrantColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessDeleteCapGrantRecords"
                        SpecificRowColumnEvents.DeleteCapGrantColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessAddRepairsRecords"
                        SpecificRowColumnEvents.InsertRepairsColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessDeleteRepairsRecords"
                        SpecificRowColumnEvents.DeleteRepairsColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessAddFundingRecords"
                        SpecificRowColumnEvents.InsertFundingColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessDeleteFundingRecords"
                        SpecificRowColumnEvents.DeleteFundingColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessAddDevelopmentIdentifiedRecords"
                        SpecificRowColumnEvents.InsertDevelopmentIdentifiedColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessDeleteDevelopmentIdentifiedRecords"
                        SpecificRowColumnEvents.DeleteDevelopmentIdentifiedColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessAddDevelopmentMultiYearRecords"
                        SpecificRowColumnEvents.InsertDevelopmentMultiYearColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessDeleteDevelopmentMultiYearRecords"
                        SpecificRowColumnEvents.DeleteDevelopmentMultiYearColumns(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessAddJournalRecords"
                        SpecificRowColumnEvents.InsertJournalRows(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessDeleteJournalRecords"
                        SpecificRowColumnEvents.DeleteJournalRows(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessAddStockConversionRecords"
                        SpecificRowColumnEvents.InsertStockConversionRows(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case "ProcessDeleteStockConversionRecords"
                        SpecificRowColumnEvents.DeleteStockConversionRows(ModelID, GridTag, GridTransaction, ActioningForm)

                    Case Else
                        GridTransaction.BError = True
                        GridTransaction.EventCancelled = True
                        GridTransaction.StringReturn =
                            "Unknown grid event command '" & GridTag.CommandData & "'."

                End Select


            End Sub

        End Class
        Public Structure EventParam

            Public Name As String
            Public Value As Object
            Public Sub New(ParamName As String, ParamValue As Object)

                Name = ParamName
                Value = ParamValue

            End Sub

        End Structure

        'Class EventCacher
        '    Private _CachedEvents As New List(Of BPEvent)
        '    Private _ModelID As Integer
        '    Public Sub CacheEvent(ByVal EventName As String, ByVal ParamArray EventParams() As Object)
        '        Dim BPE As New BPEvent(_ModelID, EventName, EventParams)
        '        _CachedEvents.Add(BPE)
        '    End Sub
        '    Public Sub ProcessCachedEvents()
        '        For Each BPE In _CachedEvents
        '            BPE.ProcessEvent()
        '        Next
        '        _CachedEvents.Clear()
        '    End Sub
        'End Class

    End Class

End Namespace
