Imports System.Data.SqlTypes
Imports System.Xml
Imports System.Xml.Serialization
Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager

Namespace Abovo
    Public Class StructureManager

        Public IsInitiliased As Boolean
        Public DefinedStructure As Abovo_Model_Def
        Public ModelID As Integer

        Sub New(SetModelId As Integer)

            IsInitiliased = False
            ModelID = SetModelId

        End Sub

        Public Function CreateStructureFromXML(
            Optional ByVal StructureSource As String = Nothing) As AbovoTransaction

            Dim Result As New AbovoTransaction

            Try
                Dim Serializer As New XmlSerializer(GetType(Abovo_Model_Def))
                Dim ParsedStructure As Abovo_Model_Def

                If Not String.IsNullOrWhiteSpace(StructureSource) AndAlso
                   StructureSource.TrimStart().StartsWith("<", StringComparison.Ordinal) Then

                    Using SourceReader As New System.IO.StringReader(StructureSource)
                        ParsedStructure =
                            CType(Serializer.Deserialize(SourceReader), Abovo_Model_Def)
                    End Using

                Else
                    Dim StructurePath As String

                    If String.IsNullOrWhiteSpace(StructureSource) Then
                        StructurePath =
                            System.IO.Path.Combine(Application.StartupPath, "Structure.xml")
                    Else
                        StructurePath = System.IO.Path.GetFullPath(StructureSource)
                    End If

                    If Not System.IO.File.Exists(StructurePath) Then
                        Throw New System.IO.FileNotFoundException(
                            "The workbook interface definition could not be found.",
                            StructurePath)
                    End If

                    Using SourceReader As New System.IO.StreamReader(StructurePath)
                        ParsedStructure =
                            CType(Serializer.Deserialize(SourceReader), Abovo_Model_Def)
                    End Using
                End If

                If ParsedStructure Is Nothing Then
                    Throw New InvalidOperationException(
                        "The workbook interface definition was empty.")
                End If

                ExcelModels(ModelID).WBStructure = ParsedStructure
                DefinedStructure = ParsedStructure
                IsInitiliased = True

                Result.BSuccess = True
                Result.StringReturn = "Workbook structure loaded."
                Result.StrResponseMessage = Result.StringReturn

            Catch ex As Exception
                IsInitiliased = False
                Result.BError = True
                Result.IntReturnCode = -1
                Result.StringReturn =
                    "The workbook interface definition could not be loaded: " &
                    ex.Message
                Result.StrResponseMessage = Result.StringReturn
            End Try

            Return Result

        End Function

    End Class
    Public Class Abovo_Model_Def

        <XmlElement("Name")> Public Name As String
        <XmlElement("DefID")> Public DefID As String
        <XmlElement("FileID")> Public FileID As String
        <XmlElement("CompanyName")> Public CompanyName As String
        <XmlElement("StartDate")> Public StartDate As String
        <XmlElement("RejData")> Public RejData As String
        <XmlElement("GroupStructure")> Public Property GroupStructures() As List(Of GroupStructure)
        <XmlElement("DefiningNRs")> Public Property DefiningNRs() As List(Of DefiningNR)
        <XmlElement("StressTestDefinition")> Public Property StressTestDefinition() As Abovo_Model_Def_StressTest
        <XmlElement("FFRDefinition")> Public Property FFRDefinition() As Abovo_Model_Def_FFRDef

    End Class
    <Serializable()>
    Public Class Abovo_Model_Def_StressTest

        <XmlElement("StresstestStressesRange")> Public StresstestStressesRange As String
        <XmlElement("StresstestMitigationsRange")> Public StresstestMitigationsRange As String
        <XmlElement("StresstestMitigationsMoneyRange")> Public StresstestMitigationsMoneyRange As String
        <XmlElement("StresstestMitigationsDevRange")> Public StresstestMitigationsDevRange As String
        <XmlElement("StresstestCovenantGraphsDataRange")> Public StresstestCovenantGraphsDataRange As String
        <XmlElement("StresstestBreachOutputRange")> Public StresstestBreachOutputRange As String
        <XmlElement("StresstestOutputTextRange")> Public StresstestOutputTextRange As String

    End Class
    <Serializable()>
    Public Class Abovo_Model_Def_FFRDef

        <XmlElement("FFRKeyPrecursorRange")> Public FFRKeyPrecursorRange As String
        <XmlElement("FFRKeys")> Public FFRKeys As String
        <XmlElement("FFRKeyAfterRange")> Public FFRKeyAfterRange As String
        <XmlElement("FFRKeyValidationRange")> Public FFRKeyValidationRange As String
        <XmlElement("StresstestCovenantGraphsDataRange")> Public StresstestCovenantGraphsDataRange As String
        <XmlElement("StresstestBreachOutputRange")> Public StresstestBreachOutputRange As String
        <XmlElement("StresstestOutputTextRange")> Public StresstestOutputTextRange As String
        <XmlElement("FFRRegSubsRange")> Public FFRRegSubsRange As String
        <XmlElement("FFRUnRegSubsRange")> Public FFRUnRegSubsRange As String
        <XmlElement("FFRInpAdj1")> Public FFRInpAdj1 As String
        <XmlElement("FFRInpAdj2")> Public FFRInpAdj2 As String

    End Class
    <Serializable()>
    Public Class DefiningNR

        <XmlElement("NRName")> Public NRName As String
        <XmlElement("BaseOrientation")> Public BaseOrientation As String
        <XmlElement("LockFirstRows")> Public LockFirstRows As String
        <XmlElement("LockLastRecords")> Public LockLastRecords As String
        <XmlElement("EditMessage")> Public EditMessage As String
        <XmlElement("HelpID")> Public HelpID As String
        <XmlElement("DataFieldDefinition")> Property DataFieldDefinitions As New List(Of DataFieldDefinition)

    End Class
    <Serializable()>
    Public Class GroupStructure

        <XmlElement("GSName")> Public GSName As String
        <XmlElement("GSID")> Public GSID As String
        <XmlElement("FirstChild")> Public FirstChild As String
        <XmlElement("ChildStructure")> Public Property ChildStructures As New List(Of ChildStructure)

        Public Function ResolveChildStructure(ByVal SetCSID As Integer) As ChildStructure

            Dim PositionalCandidate As ChildStructure = Nothing

            If SetCSID >= 0 AndAlso SetCSID < ChildStructures.Count Then
                PositionalCandidate = ChildStructures(SetCSID)

                Dim PositionalID As Integer
                If PositionalCandidate IsNot Nothing AndAlso
                   Integer.TryParse(PositionalCandidate.CSID, PositionalID) AndAlso
                   PositionalID = SetCSID Then

                    Return PositionalCandidate

                End If
            End If

            For Each Candidate As ChildStructure In ChildStructures
                If Candidate Is Nothing Then Continue For

                Dim CandidateID As Integer
                If Integer.TryParse(Candidate.CSID, CandidateID) AndAlso
                   CandidateID = SetCSID Then

                    Return Candidate

                End If
            Next

            'Retain compatibility with old structures whose CSID was blank or
            'non-numeric and which therefore could only be addressed by position.
            If PositionalCandidate IsNot Nothing Then Return PositionalCandidate

            Throw New ArgumentOutOfRangeException(
                NameOf(SetCSID),
                SetCSID,
                "The requested child structure does not exist in group " &
                If(GSName, String.Empty) & ".")

        End Function

    End Class

    <Serializable()>
    Public Class ChildStructure

        <XmlElement("CSName")> Public CSName As String
        <XmlElement("CSID")> Public CSID As String
        <XmlElement("ParentID")> Public ParentID As String
        <XmlElement("IsMaster")> Public IsMaster As String
        <XmlElement("GroupName")> Public GroupName As String
        <XmlElement("Expanded")> Public Expanded As Boolean
        <XmlElement("SpecialElement")> Public SpecialElement As String
        <XmlElement("DefaultWorksheet")> Public DefaultWorksheet As String
        <XmlElement("CSInterfaceSection")> Property InterfaceSections As New List(Of CSInterfaceSection)
        <XmlElement("CSHeader")> Property Header As CSInterfaceSection

    End Class
    <Serializable()>
    Public Class CSInterfaceSection

        <XmlElement("ISName")> Public ISName As String
        <XmlElement("CSID")> Public CSID As String
        <XmlElement("ParentID")> Public ParentID As String
        <XmlElement("CustomCodeID")> Public CustomCodeID As String
        <XmlElement("ISLABEL")> Public ISLABEL As String
        <XmlElement("ISLABELPreText")> Public ISLABELPreText As String
        <XmlElement("ISLABELPostText")> Public ISLABELPostText As String
        <XmlElement("ISLABELSource")> Public ISLABELSource As String
        <XmlElement("IsExpanded")> Public IsExpanded As String
        <XmlElement("HelpID")> Public HelpID As String
        <XmlElement("IElement")> Property IElements As New List(Of CSSectionElement)

        <XmlElement("ISDatasource")> Property ISDatasources As New List(Of ISEDatasource)
        <XmlElement("InterfaceControl")> Property InterfaceControls As New List(Of InterfaceControl)

    End Class




    <Serializable()>
    Public Class ISEDatasource

        <XmlElement("ISDName")> Public ISDName As String
        <XmlElement("ISDID")> Public CSID As String
        <XmlElement("DSType")> Public DSType As String
        <XmlElement("DSSource")> Public DSSource As String
        <XmlElement("ParentID")> Public ParentID As Integer
        <XmlElement("MergeSources")> Public MergeSources As String
        <XmlElement("MergedHeader")> Public MergedHeader As String
        <XmlElement("SourceDataFormat")> Public SourceDataFormat As String
        <XmlElement("Pivot")> Public Pivot As String
        <XmlElement("RowExpandByNR")> Public RowExpandByNR As String
        <XmlElement("StructureRuleID")> Public StructureRuleID As String
        <XmlElement("StructureAddCommand")> Public StructureAddCommand As String
        <XmlElement("StructureDeleteCommand")> Public StructureDeleteCommand As String
        <XmlElement("RowsExpandModel")> Public RowsExpandModel As String
        <XmlElement("SkipLastRecords")> Public SkipLastRecords As String
        <XmlElement("RO")> Public RO As String
        <XmlElement("Action")> Property Actions As New List(Of ApplicationAction)
        <XmlElement("NamedRangeSource")> Property NamedRangeSources As New List(Of NamedRangeDataSource)
        <XmlElement("CellRangeDataSource")> Property CellRangeSources As New List(Of CellRangeDataSource)
        <XmlElement("CalculatedSource")> Property CalculatedSources As New List(Of CalculatedDataSource)

    End Class
    <Serializable()>
    Public Class NamedRangeDataSource

        <XmlElement("NRDSName")> Public CSName As String
        <XmlElement("PositID")> Public PositID As String
        <XmlElement("ParentID")> Public ParentID As String
        <XmlElement("DefinedBy")> Public DefinedBy As String
        <XmlElement("ExpandsBy")> Public ExpandsBy As String
        <XmlElement("NRName")> Public NRName As String
        <XmlElement("NRFormatString")> Public NRFormat As String
        <XmlElement("DataFieldDefinition")> Property DataFieldDefinitions As New List(Of DataFieldDefinition)

    End Class
    <Serializable()>
    Public Class CellRangeDataSource

        Private _WSName As String

        <XmlElement("Worksheet")>
        Public Property WSName As String
            Get
                Return _WSName
            End Get
            Set(value As String)
                'XmlSerializer preserves indentation inside string content.
                'Worksheet names are identifiers, so trim formatting whitespace.
                _WSName = If(value Is Nothing, Nothing, value.Trim())
            End Set
        End Property
        <XmlElement("NRDSName")> Public NRDSName As String
        <XmlElement("ColsDefinedBy")> Public ColsDefinedBy As String
        <XmlElement("ColsDefinedByData")> Public ColsDefinedByData As String
        <XmlElement("ColsDescription")> Public ColsDescription As String
        <XmlElement("RowsDefinedBy")> Public RowsDefinedBy As String
        <XmlElement("RowsDefinedByData")> Public RowsDefinedByData As String
        <XmlElement("RowsDescription")> Public RowsDescription As String
        <XmlElement("MergeRowHeadingFormat")> Public MergeRowHeadingFormat As String
        <XmlElement("DataRange")> Public DataRange As String
        <XmlElement("DataRangeExtensionData")> Public DataRangeExtensionData As String
        <XmlElement("LiveGridSourceName")> Public LiveGridSourceName As String
        <XmlElement("LiveGridSourceRanges")> Public LiveGridSourceRanges As String
        <XmlElement("LiveGridSourceAreas")> Public LiveGridSourceAreas As String
        <XmlElement("LiveGridLeadingColumns")> Public LiveGridLeadingColumns As String
        <XmlElement("LiveGridHeaderRows")> Public LiveGridHeaderRows As String
        <XmlElement("LiveVGridCategoryRow")> Public LiveVGridCategoryRow As String
        <XmlElement("LiveVGridRecordHeaderColumns")> Public LiveVGridRecordHeaderColumns As String
        <XmlElement("PositID")> Public PositID As String
        <XmlElement("ParentID")> Public ParentID As String
        <XmlElement("BandID")> Public BandID As String
        <XmlElement("BandTipText")> Public BandTipText As String
        <XmlElement("BandEditDescription")> Public BandEditDescription As String
        <XmlElement("RO")> Public RO As String
        <XmlElement("IsCalculated")> Public IsCalculated As String
        <XmlElement("IsOffSet")> Public IsOffSet As String
        <XmlElement("OffSetNR")> Public OffSetNR As String
        <XmlElement("OffSetBy")> Public OffSetBy As String
        <XmlElement("HasExtraRecords")> Public HasExtraRecords As String
        <XmlElement("IsMappedRightByNR")> Public IsMappedRightByNR As String

        Public HasBands As Boolean = False
        <XmlElement("DataFieldDefinition")> Property DataFieldDefinitions As New List(Of DataFieldDefinition)
        <XmlElement("Action")> Property Actions As New List(Of ApplicationAction)

    End Class
    <Serializable()>
    Public Class CalculatedDataSource

        <XmlElement("NRDSName")> Public CSName As String
        <XmlElement("PositID")> Public PositID As Integer
        <XmlElement("ParentID")> Public ParentID As Integer
        <XmlElement("DefinedBy")> Public DefinedBy As String
        <XmlElement("ExpandsBy")> Public ExpandsBy As String
        <XmlElement("NRName")> Public NRName As String
        <XmlElement("NRFormatString")> Public NRFormat As String

    End Class
    <Serializable()>
    Public Class CalculatedField

        <XmlElement("CalcType")> Public CSName As String
        <XmlElement("PositID")> Public PositID As Integer
        <XmlElement("DataFormat")> Public DataFormat As String

    End Class
    <Serializable()>
    Public Class DataFieldDefinition

        <XmlElement("Index")> Public Index As String
        <XmlElement("FieldName")> Public FieldName As String
        <XmlElement("IsDummy")> Public IsDummy As String
        <XmlElement("DataFormat")> Public DataFormat As String
        <XmlElement("RepositaryItemID")> Public RepositaryItemID As String
        <XmlElement("RO")> Public RO As String
        <XmlElement("ShowSummary")> Public ShowSummary As String
        <XmlElement("MinVal")> Public MinVal As String
        <XmlElement("MaxVal")> Public MaxVal As String
        <XmlElement("Fixed")> Public Fixed As String

        'Optional minimum display width for a generated grid column, expressed
        'in approximate character cells rather than pixels. This keeps Structure
        'XML portable across DPI/display sizes.
        <XmlElement("MinWidthChars")> Public MinWidthChars As String

        <XmlElement("FontFormat")> Public FontFormat As String
        <XmlElement("RepeatsByNR")> Public RepeatsByNR As String
        <XmlElement("RepeatingNR")> Public RepeatingNR As String
        <XmlElement("EditRepNRHere")> Public EditRepNRHere As String
        <XmlElement("EditRepNRHereDataFormat")> Public EditRepNRHereDataFormat As String
        <XmlElement("EditRepNRHereEditor")> Public EditRepNRHereEditor As String
        <XmlElement("EditRepNRHereComboRepository")> Public EditRepNRHereComboRepository As String
        <XmlElement("EditRepNRHereExpansionMethod")> Public EditRepNRHereExpansionMethod As String
        <XmlElement("EditRepNRHereROPreNRLines")> Public EditRepNRHereROPreNRLines As String
        <XmlElement("EditRepNRHereROInitialLines")> Public EditRepNRHereROInitialLines As String
        <XmlElement("EditRepNRHereRule")> Public EditRepNRHereRule As String
        <XmlElement("RepeatsByCR")> Public RepeatsByCR As String
        <XmlElement("RepeatingPreRows")> Public RepeatingPreRows As String
        <XmlElement("RepeatingPostRows")> Public RepeatingPostRows As String
        <XmlElement("RepeatsByCRData")> Public RepeatsByCRData As String
        <XmlElement("IsColCalculated")> Public IsColCalculated As String
        <XmlElement("HasControls")> Public HasControls As String
        <XmlElement("ExtraHeadingPreWord")> Public ExtraHeadingPreWord As String
        <XmlElement("RepeatingHeaderText")> Public RepeatingHeaderText As String
        <XmlElement("BandEditDescription")> Public BandEditDescription As String
        <XmlElement("NRTextPerItem")> Public NRTextPerItem As String
        <XmlElement("Units")> Public Units As String
        <XmlElement("TipText")> Public TipText As String
        <XmlElement("HasRule")> Public HasRule As String

    End Class
    <Serializable()>
    Public Class DFDRule

        <XmlElement("FieldName")> Public FieldName As String
        <XmlElement("RuleType")> Public RuleType As String
        <XmlElement("RuleValue")> Public RuleValue As String
        <XmlElement("RuleMessage")> Public RuleMessage As String

    End Class
    <Serializable()>
    Public Class CSSectionElement

        <XmlElement("CSName")> Public CSName As String
        <XmlElement("CSID")> Public CSID As String
        <XmlElement("ParentID")> Public ParentID As String
        <XmlElement("Type")> Public Type As String
        <XmlElement("DataSource")> Public DataSource As String
        <XmlElement("Description")> Public Description As String
        <XmlElement("GridShowTotals")> Public GridShowTotals As String

        <XmlElement("CompoundTable")> Public CompoundTable As CompoundMappedTable
        <XmlElement("InterfaceControlObject")> Property InterfaceControls As New List(Of InterfaceControlObject)
        <XmlElement("InterfaceLinkObject")> Property LinkObjects As New List(Of InterfaceLinkElement)
        <XmlElement("GridControlButton")> Property GridControls As New List(Of GridCommandButton)
        <XmlElement("SpreadsheetOptions")> Property SpreadsheetOptions As New SpreadsheetOptions
        <XmlElement("MappedTable")> Property MappedTable As New MappedTable

    End Class

    <Serializable()>
    Public Class InterfaceControlObject

        <XmlElement("ItemID")> Public ItemID As String
        <XmlElement("ItemTxt")> Public ItemTxt As String
        <XmlElement("ItemTip")> Public ItemTip As String
        <XmlElement("ItemType")> Public ItemType As String
        <XmlElement("ItemData")> Public ItemData As String

    End Class
    <Serializable()>
    Public Class GridCommandButton

        <XmlElement("ButtonID")> Public ButtonID As String
        <XmlElement("ButtonTxt")> Public ButtonTxt As String
        <XmlElement("ButtonTip")> Public ButtonTip As String
        <XmlElement("ButtonType")> Public ButtonType As String
        <XmlElement("ButtonData")> Public ButtonData As String

    End Class

    <Serializable()>
    Public Class RepositaryItem

        <XmlElement("RIName")> Public RIName As String
        <XmlElement("RIID")> Public RIID As String
        <XmlElement("Type")> Public Type As String
        <XmlElement("DataSourceType")> Public DataSource As String
        <XmlElement("DataSourceData")> Public Data As Object
        <XmlElement("ListData")> Property ListData As New List(Of String)

    End Class
    <Serializable()>
    Public Class SpreadsheetOptions

        <XmlElement("SkipFirstRows")> Public SkipFirstRows As String
        <XmlElement("SkipLastRows")> Public SkipLastRows As String

    End Class

    <Serializable()>
    Public Class CompoundMappedTable

        <XmlElement("NumRows")> Public NumRows As String
        <XmlElement("StartingCol")> Public StartingCol As String
        <XmlElement("StartingRow")> Public StartingRow As String
        <XmlElement("CMTColDef")> Property CMTColDefs As New List(Of CompoundMappedTableColumnDefinition)
        <XmlElement("MappedTableRow")> Property MappedTableRows As New List(Of MappedTableRow)

    End Class
    Public Class CompoundMappedTableColumnDefinition

        <XmlElement("ColSetID")> Public ColSetID As String
        <XmlElement("ColSetType")> Public ColSetType As String
        <XmlElement("RepeatsBy")> Public RepeatsBy As String
        <XmlElement("ElementRepeats")> Public ElementRepeats As Boolean
        <XmlElement("IsFixed")> Public IsFixed As String
        <XmlElement("ColCount")> Public ColCount As Integer
        <XmlElement("StartColIndex")> Public StartColIndex As Integer
        <XmlElement("HasRules")> Public HasRules As String

    End Class
    <Serializable()>
    Public Class MappedTable

        <XmlElement("NumRows")> Public NumRows As String
        <XmlElement("NumCols")> Public NumCols As String

        <XmlElement("MappedTableRow")> Property MappedTableRows As New List(Of MappedTableRow)

    End Class
    <Serializable()>
    Public Class MappedTableRow

        <XmlElement("RowIndex")> Public RowIndex As String
        <XmlElement("ColSetID")> Public ColSetID As String
        <XmlElement("NumColumns")> Public NumColumns As String
        <XmlElement("RepeatsBy")> Public RepeatsBy As String
        <XmlElement("HasRules")> Public HasRules As String
        <XmlElement("MappedTableElement")> Property MappedTableElements As New List(Of MappedTableElement)

    End Class
    <Serializable()>
    Public Class MappedTableElement

        <XmlElement("CIndex")> Public CIndex As String
        <XmlElement("RowSpan")> Public RowSpan As String
        <XmlElement("ColSpan")> Public ColSpan As String
        <XmlElement("Type")> Public Type As String
        <XmlElement("ToolTip")> Public ToolTip As String
        <XmlElement("RepeatsBy")> Public RepeatsByNR As String

        <XmlElement("TELink")> Property TELinks As New List(Of TELink)
        <XmlElement("TEComboBox")> Property TEComboBoxes As New List(Of TEComboBox)
        <XmlElement("TETextBox")> Property TETextBoxes As New List(Of TETextBox)
        <XmlElement("TEDateEdits")> Property TEDateEdits As New List(Of TEDateBox)
        <XmlElement("TELabel")> Property TELabels As New List(Of TELabel)

    End Class
    <Serializable()>
    Public Class TEComboBox

        <XmlElement("CRSource")> Public CSSource As String
        <XmlElement("NRSource")> Public NRSource As String
        <XmlElement("CBText")> Public CBText As String
        <XmlElement("DataFormat")> Public DataFormat As String
        <XmlElement("DataType")> Public DataType As String
        <XmlElement("RepID")> Public RepID As String

    End Class
    <Serializable()>
    Public Class TELink

        <XmlElement("CRSource")> Public CSSource As String
        <XmlElement("LinkTarget")> Public LinkTarget As String
        <XmlElement("LinkTargetSection")> Public LinkTargetSection As String
        <XmlElement("LinkGroup")> Public LinkGroup As String

    End Class
    <Serializable()>
    Public Class TETextBox

        <XmlElement("IsReadOnly")> Public IsReadOnly As String
        <XmlElement("CRSource")> Public CSSource As String
        <XmlElement("NRSource")> Public NRSource As String
        <XmlElement("Text")> Public CBText As String
        <XmlElement("LinkRange")> Public LinkRange As String
        <XmlElement("LinkTarget")> Public LinkTarget As String
        <XmlElement("DataType")> Public DataType As String
        <XmlElement("MinVal")> Public MinVal As String
        <XmlElement("MaxVal")> Public MaxVal As String

    End Class
    <Serializable()>
    Public Class TEDateBox

        <XmlElement("CRSource")> Public CSSource As String
        <XmlElement("NRSource")> Public NRSource As String
        <XmlElement("Text")> Public CBText As String
        <XmlElement("LinkRange")> Public LinkRange As String
        <XmlElement("LinkTarget")> Public LinkTarget As String
        <XmlElement("DataFormat")> Public DataFormat As String
        <XmlElement("MinDate")> Public MinVal As String
        <XmlElement("MaxDate")> Public MaxVal As String
        <XmlElement("IsStatic")> Public IsStatic As String

    End Class
    <Serializable()>
    Public Class TELabel

        <XmlElement("IsTitle")> Public IsTitle As String
        <XmlElement("IsBold")> Public IsBold As String
        <XmlElement("IsBordered")> Public IsBordered As String
        <XmlElement("IsUnderline")> Public IsUnderline As String
        <XmlElement("CRSource")> Public CSSource As String
        <XmlElement("NRSource")> Public NRSource As String
        <XmlElement("IsStatic")> Public IsStatic As String
        <XmlElement("Alignment")> Public Alignment As String

    End Class

    <Serializable()>
    Public Class InterfaceControl

        <XmlElement("Type")> Public Type As String
        <XmlElement("DataSource")> Public DataSource As String
        <XmlElement("SuperTip")> Public SuperTip As String
        <XmlElement("HelpID")> Public HelpID As String

    End Class
    <Serializable()>
    Public Class InterfaceLinkElement

        <XmlElement("LinkType")> Public LinkType As String
        <XmlElement("LinkData")> Public LinkData As String
        <XmlElement("LinkGroup")> Public LinkGroup As String
        <XmlElement("LinkTip")> Public LinkTip As String
        <XmlElement("LinkText")> Public LinkText As String
        <XmlElement("LinkTargetSection")> Public LinkTargetSection As String
        <XmlElement("LinkReturn")> Public LinkReturn As Integer

    End Class
    Public Class ApplicationAction

        <XmlElement("Type")> Public Type As String
        <XmlElement("ActionText")> Public ActionText As String
        <XmlElement("ActionSuperTip")> Public ActionSuperTip As String
        <XmlElement("HelpID")> Public HelpID As String
        <XmlElement("ActionData1")> Public ActionData1 As String
        <XmlElement("ActionData2")> Public ActionData2 As String
        <XmlElement("ActionData3")> Public ActionData3 As String

    End Class

End Namespace
