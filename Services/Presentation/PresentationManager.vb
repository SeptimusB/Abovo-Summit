Imports System.Xml.Serialization
Imports Abovo.AbovoAppCls
Imports Abovo.DataObject
Imports Abovo.LogDebugDev

Imports DevExpress.Xpo.DB
Imports DevExpress.XtraSpreadsheet.Model

Namespace Abovo
    Public Class PresentationManager

        Public ModelID As Integer
        Public DataPresentations() As DataPresentation
        Public DPCount As Integer
        Public DPIndex As Integer
        Public Sub New(SetModelID As Integer)

            ModelID = SetModelID
            DPCount = -1
            DPIndex = -1
            ReDim DataPresentations(-1)

        End Sub
        Function ValidatePresentation(SetGSID As Integer, SetCSID As Integer) As AbovoTransaction

            Dim DataTrans As AbovoTransaction = New AbovoTransaction

            If DPCount = 0 Then GoTo CreatePresentation

            For Each dataPres In DataPresentations

                If dataPres.GSID = SetGSID And dataPres.CSID = SetCSID Then

                    DataTrans.IntegerReturn = dataPres.DSIndex
                    DataTrans.BError = False
                    Return DataTrans
                    Exit Function

                End If

            Next

CreatePresentation:

            DPCount += 1
            DPIndex += 1
            ReDim Preserve DataPresentations(DPIndex)
            DataPresentations(DPIndex) = New DataPresentation(ModelID, SetGSID, SetCSID)
            'MsgBox(DPIndex.ToString)
            DataTrans.IntegerReturn = DPIndex
            Return DataTrans

        End Function
        Public Function CreatePresentation(SetGSID As Integer, SetCSID As Integer) As AbovoTransaction

            Dim TransReturn As New AbovoTransaction


            DPCount += 1
            DPIndex += 1

            ReDim Preserve DataPresentations(DPIndex)

            DataPresentations(DPIndex) = New DataPresentation(ModelID, SetGSID, SetCSID)

            TransReturn.IntegerReturn = DPIndex

            If DataPresentations(DPIndex).ProcessStructure = True Then

                TransReturn.BError = False

            Else

                TransReturn.BError = True

            End If

            Return TransReturn

        End Function
        Public Class DataPresentation

            Public Sections(-1) As PresentationSection
            Public Header As PresentationSection
            Public HasHeader As Boolean
            Public DataSets(-1) As DataCellRange
            Public SectionCount As Integer
            Public DefaultWorksheet As String
            Public ModelID As Integer
            Public GSID As Integer
            Public CSID As Integer
            Public PresType As String
            Public PresName As String
            Public HasError As Boolean
            Public HTMLOutput As String
            Public DataSetCount As Integer
            Public DPCS As ChildStructure
            Public DSIndex As Integer
            Public Name As String
            Public SElementCount As Integer = -1


            Public Function ProcessStructure() As Boolean

                Return True

            End Function

            Public Sub New(SetModelID As Integer, SetGSID As Integer, SetCSID As Integer)

                HasError = False
                GSID = SetGSID
                CSID = SetCSID
                ModelID = SetModelID
                DataSetCount = -1
                SectionCount = -1
                DSIndex = -1


                DPCS = FileManager.ExcelModels(SetModelID).WBStructure.GroupStructures(GSID).ResolveChildStructure(CSID)

                PresName = DPCS.CSName

                DefaultWorksheet = DPCS.DefaultWorksheet

                If DPCS.Header Is Nothing Then

                    HasHeader = False

                Else

                    HasHeader = True

                End If

                ReDim Sections(DPCS.InterfaceSections.Count - 1)

                Dim InitialDataSetCount As Integer = 0

                For Each SectionDefinition As CSInterfaceSection In DPCS.InterfaceSections
                    For Each ElementDefinition In SectionDefinition.IElements
                        Select Case ElementDefinition.Type
                            Case "Grid", "VGrid", "LiveGrid", "LiveVGrid", "TextBox", "Label", "ComboBox", "DateBox"
                                InitialDataSetCount += 1
                        End Select
                    Next
                Next

                ReDim DataSets(InitialDataSetCount - 1)

                For Each ISection In DPCS.InterfaceSections

                    SectionCount += 1
                    Sections(SectionCount) = New PresentationSection With {
                        .Name = ISection.ISName,
                        .ID = SectionCount}

                    ProcessSection(SetModelID, Sections(SectionCount), ISection)


                Next

            End Sub

            Sub ProcessSection(SetModelID As Integer, ByVal Section As PresentationSection, ByVal ISection As CSInterfaceSection)

                SElementCount = -1
                Dim Model = FileManager.ExcelModels(SetModelID)
                Dim WBData = Model.WBData

                If ISection.IsExpanded = "FALSE" Then Section.Expanded = False

                If ISection.ISLABEL = "TRUE" Then

                    Dim StrName As String = ""

                    StrName += ISection.ISLABELPreText

                    Dim LabelText As String =
                        Model.WB.Worksheets(DefaultWorksheet).Cells(ISection.ISLABELSource).DisplayText


                    If LabelText = "" Then

                        StrName += "(Undefined)"

                    Else

                        StrName += LabelText

                    End If

                    StrName += ISection.ISLABELPostText
                    Sections(SectionCount).Name = StrName

                End If

                'ReDim as the section may be recreated independently later.
                ReDim Section.SectionElements(ISection.IElements.Count - 1)

                For Each ISElement In ISection.IElements

                    SElementCount += 1

                    Section.SectionElements(SElementCount) = New PresentationSectionElement With {.PSEIndex = SElementCount}

                    If ISElement.Type = "Grid" Or ISElement.Type = "VGrid" Then

                        Section.SectionElements(SElementCount).Type = ISElement.Type

                        DataSetCount += 1
                        EnsureDataSetSlot(DataSetCount)
                        DataSets(DataSetCount) = WBData.GetISEDataStructure(ModelID, GSID, CSID, SectionCount, CInt(ISElement.DataSource))
                        Section.SectionElements(SElementCount).ControlSourceIndex = DataSetCount

                        If ISElement.GridControls.Count > 0 Then
                            Section.SectionElements(SElementCount).GridControls = New List(Of AttachedGridCommandButton)
                            For i = 0 To ISElement.GridControls.Count - 1
                                Dim NewButton As AttachedGridCommandButton = New AttachedGridCommandButton With {
                                    .CommandType = ISElement.GridControls(i).ButtonType,
                                    .CommandData = ISElement.GridControls(i).ButtonData,
                                    .CommandText = ISElement.GridControls(i).ButtonTxt,
                                    .CommandTip = ISElement.GridControls(i).ButtonTip
                                }
                                Section.SectionElements(SElementCount).GridControls.Add(NewButton)
                            Next



                        End If

                    ElseIf ISElement.Type = "LiveGrid" OrElse ISElement.Type = "LiveVGrid" Then

                        DataSetCount += 1
                        EnsureDataSetSlot(DataSetCount)
                        Section.SectionElements(SElementCount).Type = ISElement.Type
                        DataSets(DataSetCount) = WBData.GetISEDataStructure(ModelID, GSID, CSID, SectionCount, CInt(ISElement.DataSource))
                        Section.SectionElements(SElementCount).ControlSourceIndex = DataSetCount
                        Section.SectionElements(SElementCount).Tag.Description = ISElement.Description

                    ElseIf ISElement.Type = "ControlGroup" Then

                        Section.SectionElements(SElementCount).Type = "ControlGroup"

                        If ISElement.InterfaceControls.Count > 0 Then

                            ReDim Section.SectionElements(SElementCount).InterfaceControls(ISElement.InterfaceControls.Count - 1)
                            For i = 0 To ISElement.InterfaceControls.Count - 1
                                Section.SectionElements(SElementCount).InterfaceControls(i) = New ElementInterfaceControlTag With {
                                    .ItemID = ISElement.InterfaceControls(i).ItemID,
                                    .ItemTxt = ISElement.InterfaceControls(i).ItemTxt,
                                    .ItemTip = ISElement.InterfaceControls(i).ItemTip,
                                    .ItemType = ISElement.InterfaceControls(i).ItemType,
                                    .ItemData = ISElement.InterfaceControls(i).ItemData}
                            Next

                        End If


                        'Section.SectionElements(SElementCount).Tag
                        Section.SectionElements(SElementCount).Tag.Description = ISElement.Description

                    ElseIf ISElement.Type = "Link" Then

                        Section.SectionElements(SElementCount).Type = "Link"

                        If ISElement.LinkObjects.Count > 0 Then

                            ReDim Section.SectionElements(SElementCount).LinkControls(ISElement.LinkObjects.Count - 1)
                            For i = 0 To ISElement.LinkObjects.Count - 1
                                Section.SectionElements(SElementCount).LinkControls(i) = New ElementInterfaceLinkTag With {
                                    .LinkType = ISElement.LinkObjects(i).LinkType,
                                    .LinkData = ISElement.LinkObjects(i).LinkData,
                                    .LinkTip = ISElement.LinkObjects(i).LinkTip,
                                    .LinkText = ISElement.LinkObjects(i).LinkText,
                                    .LinkGroup = ISElement.LinkObjects(i).LinkGroup,
                                    .LinkReturnID = CSID,
                                    .LinkReturnName = PresName,
                                    .LinkReturnGroup = GSID,
                                    .LinkTargetSection = ISElement.LinkObjects(i).LinkTargetSection
                                    }
                            Next

                        End If


                        'Section.SectionElements(SElementCount).Tag
                        Section.SectionElements(SElementCount).Tag.Description = ISElement.Description

                    ElseIf ISElement.Type = "TextBlock" Then

                        Section.SectionElements(SElementCount).Type = "TextBlock"
                        Section.SectionElements(SElementCount).Tag.Description = ISElement.Description

                    ElseIf ISElement.Type = "MappedTable" Then

                        Section.SectionElements(SElementCount).Type = "MappedTable"
                        Section.SectionElements(SElementCount).MappedTableSection = ISElement.MappedTable

                    ElseIf ISElement.Type = "CompoundMappedTable" Then

                        Section.SectionElements(SElementCount).Type = "CompoundMappedTable"
                        Section.SectionElements(SElementCount).CompoundTable = ISElement.CompoundTable

                    ElseIf ISElement.Type = "Spreadsheet" Then

                        Section.SectionElements(SElementCount).Type = "Spreadsheet"
                        Section.SectionElements(SElementCount).SSOptions = ISElement.SpreadsheetOptions

                    ElseIf ISElement.Type = "TextBox" Or ISElement.Type = "Label" Then

                        DataSetCount += 1
                        Section.SectionElements(SElementCount).Type = ISElement.Type
                        EnsureDataSetSlot(DataSetCount)
                        DataSets(DataSetCount) = WBData.GetISEDataStructure(ModelID, GSID, CSID, SectionCount, CInt(ISElement.DataSource))
                        Section.SectionElements(SElementCount).ControlSourceIndex = DataSetCount

                    ElseIf ISElement.Type = "ComboBox" Then

                        DataSetCount += 1
                        Section.SectionElements(SElementCount).Type = "ComboBox"
                        EnsureDataSetSlot(DataSetCount)
                        DataSets(DataSetCount) = WBData.GetISEDataStructure(ModelID, GSID, CSID, SectionCount, CInt(ISElement.DataSource))
                        Section.SectionElements(SElementCount).ControlSourceIndex = DataSetCount

                    ElseIf ISElement.Type = "DateBox" Then

                        DataSetCount += 1
                        Section.SectionElements(SElementCount).Type = "DateBox"
                        EnsureDataSetSlot(DataSetCount)
                        DataSets(DataSetCount) = WBData.GetISEDataStructure(ModelID, GSID, CSID, SectionCount, CInt(ISElement.DataSource))
                        Section.SectionElements(SElementCount).ControlSourceIndex = DataSetCount

                    ElseIf ISElement.Type = "HTMLRender" Then

                        Section.SectionElements(SElementCount).Type = "HTMLRender"
                        HTMLOutput = FileManager.ExcelModels(SetModelID).WBData.RenderIEHTMLCource(ModelID, GSID, CSID, SectionCount, CInt(ISElement.DataSource))

                    ElseIf ISElement.Type = "TextArea" Then

                        Section.SectionElements(SElementCount).Type = "TextArea"
                        Section.SectionElements(SElementCount).Description = ISElement.Description

                    ElseIf ISElement.Type = "Browser" Then

                        Section.SectionElements(SElementCount).Type = "Browser"
                        Section.SectionElements(SElementCount).Description = ISElement.Description

                    ElseIf ISElement.Type = "Spacer" Then

                        Section.SectionElements(SElementCount).Type = "Spacer"

                    End If

                Next



            End Sub

            Private Sub EnsureDataSetSlot(ByVal RequiredIndex As Integer)

                If DataSets IsNot Nothing AndAlso RequiredIndex < DataSets.Length Then Return
                ReDim Preserve DataSets(RequiredIndex)

            End Sub

            Public Sub RedefineInterfaceSection(SectionID)

                Sections(SectionID) = Nothing

                Dim ISection As CSInterfaceSection = DPCS.InterfaceSections(SectionID)

                SectionCount = SectionID

                Sections(SectionID) = New PresentationSection With {
                        .Name = ISection.ISName,
                        .ID = SectionCount}

                ProcessSection(ModelID, Sections(SectionCount), ISection)

            End Sub


        End Class

        Public Class PresentationSection

            Public Name As String
            Public ID As Integer
            Public SectionElements() As PresentationSectionElement
            Public Controls() As Control
            Public Expanded As Boolean = True
            Public HelpID As String

        End Class
        Public Class PresentationSectionElement

            Public PSEIndex As Integer

            Public Name As String
            Public Type As String
            Public Description As String
            Public GroupID As Integer
            Public Control As Control
            Public ControlSourceIndex As Integer
            Public Tag As PresSectionElementTag
            Public SSOptions As SpreadsheetOptions
            Public MappedTableSection As MappedTable
            Public InterfaceControls() As ElementInterfaceControlTag
            Public LinkControls() As ElementInterfaceLinkTag
            Public GridControls As List(Of AttachedGridCommandButton)
            Public CompoundTable As CompoundMappedTable

        End Class
        Structure PresSectionElementTag

            Public Description As String

        End Structure

        Class ElementInterfaceControlTag

            Public ItemID As String
            Public ItemTxt As String
            Public ItemTip As String
            Public ItemType As String
            Public ItemData As String

        End Class
        Class ElementInterfaceLinkTag

            Public LinkType As String
            Public LinkData As String
            Public LinkGroup As String
            Public LinkGroupID As Integer
            Public LinkTip As String
            Public LinkText As String
            Public LinkTargetSection As String
            Public LinkReturnID As Integer
            Public LinkReturnGroup As Integer
            Public LinkReturnName As String

        End Class

        Class AttachedGridCommandButton

            Public CommandType As String
            Public CommandText As String
            Public CommandData As String
            Public CommandTip As String
            Public RequestedRecordCount As Integer
            Public DeleteLastRecords As Boolean

            'Normal XtraGrid owner.
            Public AttachedGrid As DevExpress.XtraGrid.GridControl

            'Vertical-grid owner for the small number of interfaces where a
            'structural command belongs to the VGrid rather than to a grid row.
            Public AttachedVGrid As DevExpress.XtraVerticalGrid.VGridControl

        End Class
        Class RowColEventTag

            Public ModelID As Integer
            Public GroupID As Integer
            Public SectionID As Integer
            Public ElementID As Integer
            Public RowIndex As Integer
            Public ColIndex As Integer

        End Class
    End Class

End Namespace
