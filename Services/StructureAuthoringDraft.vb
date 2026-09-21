Imports System.IO
Imports System.Xml
Imports System.Xml.Linq
Imports System.Xml.Serialization
Imports DevExpress.Spreadsheet

Namespace Abovo
    'Edit the original XML, not a serialization of the narrower runtime model.
    'Unknown metadata, comments, attributes and ordering must survive authoring.
    Public NotInheritable Class StructureAuthoringDraft
        Public ReadOnly Property Document As XDocument
        Public ReadOnly Property Source As String
        Public Property IsDirty As Boolean
        Private ReadOnly knownFormats As HashSet(Of String)

        Public Sub New(xmlOrPath As String)
            Source = If(xmlOrPath.TrimStart().StartsWith("<"), "Embedded workbook structure", Path.GetFullPath(xmlOrPath))
            Dim xml = If(xmlOrPath.TrimStart().StartsWith("<"), xmlOrPath, File.ReadAllText(xmlOrPath))
            Dim settings As New XmlReaderSettings With {.DtdProcessing = DtdProcessing.Prohibit, .XmlResolver = Nothing, .MaxCharactersInDocument = 16000000}
            Using input As New StringReader(xml), reader = XmlReader.Create(input, settings)
                Document = XDocument.Load(reader, LoadOptions.PreserveWhitespace)
            End Using
            Validate(Document)
            knownFormats = New HashSet(Of String)(Document.Descendants("DataFormat").Select(Function(x) x.Value.Trim()), StringComparer.Ordinal)
        End Sub

        Public Shared Function Validate(doc As XDocument) As Abovo_Model_Def
            If doc.Root Is Nothing OrElse doc.Root.Name <> XName.Get("Abovo_Model_Def") Then Throw New InvalidDataException("An Abovo_Model_Def structure is required.")
            Dim parsed As Abovo_Model_Def
            Using reader = doc.CreateReader()
                parsed = DirectCast(New XmlSerializer(GetType(Abovo_Model_Def)).Deserialize(reader), Abovo_Model_Def)
            End Using
            Dim groups = doc.Root.Elements("GroupStructure").ToList()
            For i = 0 To groups.Count - 1
                Dim id As Integer
                If Not Integer.TryParse(Value(groups(i), "GSID"), id) OrElse id <> i Then Throw New InvalidDataException("Group IDs must retain their ordered, zero-based Summit routes.")
                Dim childIds As New HashSet(Of Integer)
                For Each child In groups(i).Elements("ChildStructure")
                    If Not Integer.TryParse(Value(child, "CSID"), id) OrElse id < 0 OrElse Not childIds.Add(id) Then Throw New InvalidDataException("Child IDs must be unique within each group.")
                Next
            Next
            Return parsed
        End Function

        Public Shared Function Value(node As XElement, name As String) As String
            Return If(node?.Element(name)?.Value, String.Empty).Trim()
        End Function

        Public Shared Function Ancestor(node As XElement, name As String) As XElement
            Return node?.AncestorsAndSelf(name).FirstOrDefault()
        End Function

        Public Shared Function EditableNames(node As XElement) As String()
            Select Case node.Name.LocalName
                Case "ChildStructure"
                    Return {"CSName", "NavigatorCaption", "NavigatorGroupCaption", "DefaultWorksheet"}
                Case "CSInterfaceSection", "CSHeader"
                    Return {"ISName"}
                Case "ISDatasource"
                    Return {"ISDName", "RO"}
                Case "CellRangeDataSource"
                    Return {"Worksheet", "NRDSName", "DataRange", "RO", "RowsDescription", "ColsDescription"}
                Case "DataFieldDefinition"
                    Return {"FieldName", "TipText", "DataFormat", "RO", "MinWidthChars", "Units"}
                Case Else
                    Return {}
            End Select
        End Function

        Public Sub Apply(node As XElement, changes As IDictionary(Of String, String), workbook As IWorkbook)
            If node.Document IsNot Document Then Throw New ArgumentException("The selection no longer belongs to this draft.")
            Dim oldNode As New XElement(node)
            Try
                For Each item In changes
                    If Not EditableNames(node).Contains(item.Key) Then Throw New InvalidOperationException("This property is inspection-only in the trial: " & item.Key)
                    Dim valueText = If(item.Value, String.Empty).Trim()
                    If item.Key = "RO" AndAlso valueText <> "" AndAlso Not {"TRUE", "FALSE"}.Contains(valueText.ToUpperInvariant()) Then Throw New InvalidDataException("RO must be TRUE or FALSE.")
                    If item.Key = "RO" Then valueText = valueText.ToUpperInvariant()
                    If item.Key = "DataFormat" AndAlso valueText <> "" AndAlso Not knownFormats.Contains(valueText) Then Throw New InvalidDataException("Unknown data format. Use a format already defined in this structure: " & String.Join(", ", knownFormats.OrderBy(Function(x) x)))
                    If item.Key = "MinWidthChars" AndAlso valueText <> "" Then
                        Dim width As Double
                        If Not Double.TryParse(valueText, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, width) OrElse Double.IsNaN(width) OrElse width < 1 OrElse width > 200 Then Throw New InvalidDataException("Minimum width must be between 1 and 200 characters.")
                    End If
                    If node.Elements(item.Key).Count() > 1 Then Throw New InvalidDataException("Duplicate property: " & item.Key)
                    If node.Element(item.Key) IsNot Nothing OrElse valueText <> "" Then node.SetElementValue(item.Key, valueText)
                Next
                Validate(Document)
                If node.Name.LocalName = "CellRangeDataSource" Then
                    Dim resolved = ResolveRange(node, workbook)
                    If resolved Is Nothing Then Throw New InvalidDataException("The range does not resolve in this workbook.")
                    Dim original As CellRange = Nothing
                    Try
                        original = ResolveRange(oldNode, workbook)
                    Catch ex As Exception When TypeOf ex Is InvalidDataException OrElse TypeOf ex Is ArgumentException
                        'A broken old binding can be repaired; no workbook geometry is changed.
                    End Try
                    If original IsNot Nothing AndAlso (resolved.Areas.Count <> original.Areas.Count OrElse resolved.RowCount <> original.RowCount OrElse resolved.ColumnCount <> original.ColumnCount) Then Throw New InvalidDataException("This trial requires the same binding geometry. Structural expansion and mixed month/year axes need explicit rules.")
                    Dim sheetName = Value(node, "Worksheet")
                    If sheetName <> "" AndAlso (Not workbook.Worksheets.Contains(sheetName) OrElse sheetName <> resolved.Worksheet.Name) Then Throw New InvalidDataException("Worksheet must match the resolved range's worksheet.")
                End If
                If changes.ContainsKey("DefaultWorksheet") AndAlso Not workbook.Worksheets.Contains(Value(node, "DefaultWorksheet")) Then Throw New InvalidDataException("The worksheet does not exist.")
                IsDirty = True
            Catch
                'Restore scalar values without replacing child objects held by the tree.
                For Each key In changes.Keys.Where(Function(k) EditableNames(node).Contains(k))
                    Dim original = oldNode.Element(key)
                    If original Is Nothing Then
                        node.Element(key)?.Remove()
                    Else
                        node.SetElementValue(key, original.Value)
                    End If
                Next
                Throw
            End Try
        End Sub

        Public Shared Function ResolveRange(node As XElement, workbook As IWorkbook) As CellRange
            Dim binding = Ancestor(node, "CellRangeDataSource")
            If binding Is Nothing OrElse workbook Is Nothing Then Return Nothing
            Dim rangeName = Value(binding, "NRDSName")
            If rangeName <> "" AndAlso Not rangeName.Equals("CR", StringComparison.OrdinalIgnoreCase) Then
                Dim named = workbook.DefinedNames.GetDefinedName(rangeName)
                If named Is Nothing Then Throw New InvalidDataException("Named range not found: " & rangeName)
                Return named.Range
            End If
            Dim sheet = Value(binding, "Worksheet")
            If sheet = "" Then sheet = Value(Ancestor(binding, "ChildStructure"), "DefaultWorksheet")
            Dim address = Value(binding, "DataRange")
            If address = "" Then Return Nothing
            If Not workbook.Worksheets.Contains(sheet) Then Throw New InvalidDataException("Worksheet not found: " & sheet)
            Return workbook.Worksheets(sheet).Range(address)
        End Function

        Public Sub SaveNew(path As String)
            Validate(Document)
            If Not IO.Path.GetExtension(path).Equals(".xml", StringComparison.OrdinalIgnoreCase) Then Throw New InvalidDataException("Save the draft as an XML file.")
            'CreateNew prevents replacement even through aliases or a dialog race.
            Using stream As New FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None)
                Document.Save(stream, SaveOptions.DisableFormatting)
            End Using
            IsDirty = False
        End Sub
    End Class

    Public NotInheritable Class StructureAuthoringSession
        Implements IDisposable
        Public ReadOnly Property ModelID As Integer = -1
        Public ReadOnly Property SourceWorkbook As String
        Public ReadOnly Property Draft As StructureAuthoringDraft
        Private ReadOnly scratchDirectory As String
        Private ReadOnly scratchFile As String
        Private disposed As Boolean

        Public Sub New(path As String)
            If FileManager.OpenModelCount <> 0 Then Throw New InvalidOperationException("Structure authoring requires its own process, without normal models open.")
            SourceWorkbook = IO.Path.GetFullPath(path)
            If Not IO.Path.GetExtension(SourceWorkbook).Equals(".xlsb", StringComparison.OrdinalIgnoreCase) Then Throw New InvalidDataException("Select an XLSB business plan.")
            scratchDirectory = IO.Path.Combine(IO.Path.GetTempPath(), "Summit-Structure-" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(scratchDirectory)
            scratchFile = IO.Path.Combine(scratchDirectory, "Preview.xlsb")
            Try
                Using input As New FileStream(SourceWorkbook, FileMode.Open, FileAccess.Read, FileShare.Read), output As New FileStream(scratchFile, FileMode.CreateNew, FileAccess.Write, FileShare.None)
                    input.CopyTo(output)
                End Using
                Dim opened = FileManager.OpenModel(scratchFile, New FileInfo(scratchFile))
                If opened.BError Then Throw New InvalidDataException(opened.StringReturn)
                ModelID = opened.IntegerReturn
                Model.ChangeManager.IsReadOnlyPreview = True
                Model.ModelSpreadsheetControl.ReadOnly = True
                With Model.ModelSpreadsheetControl.Options.Behavior
                    .Save = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled
                    .SaveAs = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled
                    .Open = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled
                    .CreateNew = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled
                    .Drop = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled
                End With
                Draft = New StructureAuthoringDraft(Model.Profile.ResolveStructureSource(scratchFile))
            Catch
                Dispose()
                Throw
            End Try
        End Sub

        Public ReadOnly Property Model As FileManager.ExcelModel
            Get
                Return FileManager.ExcelModels(ModelID)
            End Get
        End Property

        Public Sub PreparePreview(doc As XDocument)
            Model.WBStructure = StructureAuthoringDraft.Validate(doc)
            Model.WBStructureManager.DefinedStructure = Model.WBStructure
            Model.WBData = New DataManager(ModelID)
            Model.WBDataPres = New PresentationManager(ModelID)
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If disposed Then Return
            disposed = True
            If ModelID >= 0 Then FileManager.CloseModel(ModelID)
            'Remove only the exact temporary file and its empty directory, never a tree.
            Try
                If scratchFile IsNot Nothing AndAlso File.Exists(scratchFile) Then File.Delete(scratchFile)
                If scratchDirectory IsNot Nothing AndAlso Directory.Exists(scratchDirectory) Then Directory.Delete(scratchDirectory, False)
            Catch ex As IOException
                Diagnostics.Trace.WriteLine("Structure preview cleanup deferred: " & ex.Message)
            End Try
        End Sub
    End Class
End Namespace
