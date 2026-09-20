Imports System.IO
Imports System.IO.Compression
Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml
Imports System.Xml.Linq
Imports System.Xml.Serialization

Namespace Abovo
    Public Enum ManagedModelRole
        Template
        PopulatedModel
    End Enum

    Public Enum ModelRuleDisposition
        Unreviewed
        PreserveCustomisation
        AssumptionInput
        CapacityOnly
        DoNotCarryForward
    End Enum

    Public Class ModelManagerRule
        Public Property Description As String = ""
        Public Property Worksheet As String = ""
        Public Property SourceRange As String = ""
        Public Property TargetRange As String = ""
        Public Property MasterRange As String = ""
        Public Property RelatedRanges As String = ""
        Public Property DataType As String = ""
        Public Property Disposition As ModelRuleDisposition
        Public Property Notes As String = ""
    End Class

    Public Class ModelManagerEvidence
        Public Property Role As String = ""
        Public Property FileName As String = ""
        Public Property SHA256 As String = ""
    End Class

    'This is declarative review data, never executable code or automatic permission to migrate.
    <XmlRoot("Abovo_Model_Manager", Namespace:=ModelManagerStore.XmlNamespace)>
    Public Class ModelManagerDefinition
        <XmlAttribute> Public Property SchemaVersion As Integer = 1
        Public Property DefinitionId As String = Guid.NewGuid().ToString("D")
        Public Property Revision As Integer
        Public Property DisplayName As String = "New model definition"
        Public Property ModelFamily As String = "HA Business Plan"
        Public Property ClientName As String = ""
        Public Property IsGeneric As Boolean
        Public Property ModelRole As ManagedModelRole
        Public Property BaseVersion As String = ""
        Public Property Status As String = "Draft"
        Public Property SavedUtc As String = ""
        Public Property Notes As String = ""
        Public Property Evidence As New List(Of ModelManagerEvidence)
        Public Property Rules As New List(Of ModelManagerRule)
    End Class

    Public NotInheritable Class ModelManagerStore
        Public Const XmlNamespace As String = "urn:abovo:summit:model-manager:1"
        Private Const MaximumXmlSize As Long = 2 * 1024 * 1024
        Private Shared ReadOnly Serializer As New XmlSerializer(GetType(ModelManagerDefinition))
        Private Shared ReadOnly RelationshipNs As XNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
        Private Shared ReadOnly ContentTypeNs As XNamespace = "http://schemas.openxmlformats.org/package/2006/content-types"
        Private Const OfficeRel As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"

        Public Shared ReadOnly Property LibraryDirectory As String
            Get
                Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                    "Abovo", "Summit", "ModelDefinitions")
            End Get
        End Property

        Public Shared Sub Validate(definition As ModelManagerDefinition)
            If definition Is Nothing OrElse definition.SchemaVersion <> 1 Then Throw New InvalidDataException("Unsupported Model Manager schema.")
            Dim parsedId As Guid
            If Not Guid.TryParse(definition.DefinitionId, parsedId) Then Throw New InvalidDataException("Invalid definition identity.")
            If definition.Revision < 0 Then Throw New InvalidDataException("Invalid revision.")
            If String.IsNullOrWhiteSpace(definition.DisplayName) Then Throw New InvalidDataException("A definition name is required.")
            If Not definition.IsGeneric AndAlso String.IsNullOrWhiteSpace(definition.ClientName) Then Throw New InvalidDataException("Specify the client, or explicitly mark this as a generic model.")
            If definition.Status <> "Draft" Then Throw New InvalidDataException("This trial supports draft definitions only; it cannot approve or execute bespoke upgrades.")
            If Not [Enum].IsDefined(GetType(ManagedModelRole), definition.ModelRole) Then Throw New InvalidDataException("Invalid model role.")
            If definition.Rules Is Nothing OrElse definition.Evidence Is Nothing Then Throw New InvalidDataException("The definition is incomplete.")
            For Each rule In definition.Rules
                If rule Is Nothing OrElse String.IsNullOrWhiteSpace(rule.Description) Then Throw New InvalidDataException("Each review rule needs a description.")
                If Not [Enum].IsDefined(GetType(ModelRuleDisposition), rule.Disposition) Then Throw New InvalidDataException("Invalid rule disposition.")
            Next
            For Each evidence In definition.Evidence
                If evidence Is Nothing OrElse Not {"Source", "Original template", "Latest template"}.Contains(evidence.Role) Then Throw New InvalidDataException("Invalid evidence role.")
                If String.IsNullOrWhiteSpace(evidence.FileName) OrElse evidence.FileName <> Path.GetFileName(evidence.FileName) Then Throw New InvalidDataException("Evidence must contain a filename, not a machine-specific path.")
                If evidence.SHA256 Is Nothing OrElse Not System.Text.RegularExpressions.Regex.IsMatch(evidence.SHA256, "\A[0-9A-Fa-f]{64}\z") Then Throw New InvalidDataException("Invalid evidence fingerprint.")
            Next
            If definition.Evidence.GroupBy(Function(item) item.Role).Any(Function(group) group.Count() > 1) Then Throw New InvalidDataException("Duplicate evidence roles.")
        End Sub

        Private Shared Function ReadXml(stream As Stream) As XDocument
            Dim settings As New XmlReaderSettings With {.DtdProcessing = DtdProcessing.Prohibit, .XmlResolver = Nothing, .MaxCharactersInDocument = MaximumXmlSize}
            Using reader = XmlReader.Create(stream, settings)
                Return XDocument.Load(reader, LoadOptions.PreserveWhitespace)
            End Using
        End Function

        Public Shared Function Parse(document As XDocument) As ModelManagerDefinition
            If document.Root Is Nothing OrElse document.Root.Name <> XName.Get("Abovo_Model_Manager", XmlNamespace) Then Throw New InvalidDataException("This is not a supported Model Manager definition.")
            If document.Root.Attribute("SchemaVersion") Is Nothing Then Throw New InvalidDataException("Missing schema version.")
            For Each field In {"DefinitionId", "Revision", "DisplayName", "IsGeneric", "ModelRole", "Status", "Evidence", "Rules"}
                If document.Root.Elements(XName.Get(field, XmlNamespace)).Count() <> 1 Then Throw New InvalidDataException("Missing or repeated definition field: " & field)
            Next
            Dim definition As ModelManagerDefinition
            Dim events As New XmlDeserializationEvents With {
                .OnUnknownElement = Sub(sender, args) Throw New InvalidDataException("Unknown definition element: " & args.Element.LocalName),
                .OnUnknownAttribute = Sub(sender, args) Throw New InvalidDataException("Unknown definition attribute: " & args.Attr.LocalName)}
            Using reader = document.CreateReader()
                definition = DirectCast(Serializer.Deserialize(reader, events), ModelManagerDefinition)
            End Using
            Validate(definition)
            Return definition
        End Function

        Public Shared Function ToDocument(definition As ModelManagerDefinition) As XDocument
            Validate(definition)
            Dim document As New XDocument()
            Using writer = document.CreateWriter()
                Dim namespaces As New XmlSerializerNamespaces()
                namespaces.Add("", XmlNamespace)
                Serializer.Serialize(writer, definition, namespaces)
            End Using
            If Encoding.UTF8.GetByteCount(document.ToString()) > MaximumXmlSize Then Throw New InvalidDataException("Definition exceeds the 2 MB limit.")
            Return document
        End Function

        Public Shared Function LoadDefinition(fileName As String) As ModelManagerDefinition
            Using stream As New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read)
                Return Parse(ReadXml(stream))
            End Using
        End Function

        Public Shared Function SaveRevision(definition As ModelManagerDefinition) As String
            Validate(definition)
            Dim saved = Parse(ToDocument(definition))
            Dim folder = Path.Combine(LibraryDirectory, saved.DefinitionId)
            Directory.CreateDirectory(folder)
            Dim revisions = Directory.GetFiles(folder, "revision-*.xml")
            Dim highest = revisions.Select(Function(item) Path.GetFileNameWithoutExtension(item).Substring(9)).
                Select(Function(item) If(System.Text.RegularExpressions.Regex.IsMatch(item, "\A[0-9]{6}\z"), Integer.Parse(item), 0)).DefaultIfEmpty(0).Max()
            saved.Revision = Math.Max(highest, saved.Revision) + 1
            saved.SavedUtc = DateTime.UtcNow.ToString("o", Globalization.CultureInfo.InvariantCulture)
            Dim destination = Path.Combine(folder, "revision-" & saved.Revision.ToString("D6") & ".xml")
            WriteNewXml(destination, ToDocument(saved))
            definition.Revision = saved.Revision
            definition.SavedUtc = saved.SavedUtc
            Return destination
        End Function

        Public Shared Sub ExportDefinition(definition As ModelManagerDefinition, destination As String)
            WriteNewXml(destination, ToDocument(definition))
        End Sub

        Private Shared Sub WriteNewXml(destination As String, document As XDocument)
            'Never truncate an existing definition, including on a failed save.
            Dim stage = destination & "." & Guid.NewGuid().ToString("N") & ".tmp"
            Try
                Using stream As New FileStream(stage, FileMode.CreateNew, FileAccess.Write, FileShare.None)
                    document.Save(stream)
                End Using
                File.Move(stage, destination)
            Finally
                If File.Exists(stage) Then File.Delete(stage)
            End Try
        End Sub

        Public Shared Function Fingerprint(fileName As String) As String
            Using stream As New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read)
                Return HashStream(stream)
            End Using
        End Function

        Private Shared Function HashStream(stream As Stream) As String
            Using hash = SHA256.Create()
                Return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "")
            End Using
        End Function

        Private Shared Function FindDefinition(package As ZipArchive) As ZipArchiveEntry
            Dim found As ZipArchiveEntry = Nothing
            For Each entry In package.Entries
                If Not entry.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase) OrElse
                   Not entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) Then Continue For
                'Read only the root of unrelated parts; do not reinterpret SharePoint metadata.
                Using stream = entry.Open()
                    Dim settings As New XmlReaderSettings With {.DtdProcessing = DtdProcessing.Prohibit, .XmlResolver = Nothing, .MaxCharactersInDocument = MaximumXmlSize}
                    Using reader = XmlReader.Create(stream, settings)
                        reader.MoveToContent()
                        If reader.LocalName <> "Abovo_Model_Manager" Then Continue For
                        If reader.NamespaceURI <> XmlNamespace Then Throw New InvalidDataException("Unsupported Model Manager XML namespace.")
                        If found IsNot Nothing Then Throw New InvalidDataException("Multiple Model Manager definitions found; select one explicitly before use.")
                        found = entry
                    End Using
                End Using
            Next
            Return found
        End Function

        Public Shared Function ReadEmbedded(fileName As String) As ModelManagerDefinition
            Using stream As New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read)
                Using package As New ZipArchive(stream, ZipArchiveMode.Read)
                    Dim entry = FindDefinition(package)
                    If entry Is Nothing Then Return Nothing
                    If entry.Length > MaximumXmlSize Then Throw New InvalidDataException("Embedded definition exceeds the size limit.")
                    Using xmlStream = entry.Open()
                        Return Parse(ReadXml(xmlStream))
                    End Using
                End Using
            End Using
        End Function

        Public Shared Sub SaveWorkbookCopy(source As String, destination As String, definition As ModelManagerDefinition)
            Dim xml = ToDocument(definition)
            source = Path.GetFullPath(source)
            destination = Path.GetFullPath(destination)
            If Not String.Equals(Path.GetExtension(source), ".xlsb", StringComparison.OrdinalIgnoreCase) OrElse
               Not String.Equals(Path.GetExtension(destination), ".xlsb", StringComparison.OrdinalIgnoreCase) Then Throw New InvalidDataException("Select XLSB files for this trial.")
            If String.Equals(source, destination, StringComparison.OrdinalIgnoreCase) OrElse File.Exists(destination) Then Throw New IOException("Choose a new output filename. Existing files are never overwritten.")
            Dim stage = destination & "." & Guid.NewGuid().ToString("N") & ".tmp"
            Try
                'Hold a read-only lock throughout copying and verification: provenance cannot race a save.
                Using sourceStream As New FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read)
                    Using staging As New FileStream(stage, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)
                        sourceStream.CopyTo(staging)
                    End Using
                    Dim changedParts As New HashSet(Of String)(StringComparer.Ordinal)
                    Using package = ZipFile.Open(stage, ZipArchiveMode.Update)
                        If package.Entries.Any(Function(item) item.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)) Then Throw New InvalidDataException("Digitally signed packages require a separate signing workflow; this trial will not invalidate the package signature.")
                        If package.GetEntry("xl/workbook.bin") Is Nothing Then Throw New InvalidDataException("Not an XLSB package.")
                        Dim existing = FindDefinition(package)
                        If existing IsNot Nothing Then
                            Using partStream = existing.Open()
                                Parse(ReadXml(partStream)) 'Do not overwrite an unsupported/corrupt definition.
                            End Using
                            WritePart(package, existing.FullName, xml, changedParts)
                        Else
                            Dim suffix = Guid.NewGuid().ToString("N")
                            Dim partName = "customXml/summitManager-" & suffix & ".xml"
                            Dim propsName = "customXml/summitManagerProps-" & suffix & ".xml"
                            WritePart(package, partName, xml, changedParts)
                            Dim ds As XNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
                            WritePart(package, propsName, New XDocument(New XElement(ds + "datastoreItem",
                                New XAttribute(ds + "itemID", "{" & Guid.NewGuid().ToString().ToUpperInvariant() & "}"),
                                New XElement(ds + "schemaRefs", New XElement(ds + "schemaRef", New XAttribute(ds + "uri", XmlNamespace))))), changedParts)
                            WritePart(package, "customXml/_rels/" & Path.GetFileName(partName) & ".rels",
                                New XDocument(New XElement(RelationshipNs + "Relationships", New XElement(RelationshipNs + "Relationship",
                                    New XAttribute("Id", "rId1"), New XAttribute("Type", OfficeRel & "customXmlProps"), New XAttribute("Target", Path.GetFileName(propsName))))), changedParts)
                            Dim relPath = "xl/_rels/workbook.bin.rels"
                            Dim rels = ReadPart(package, relPath)
                            rels.Root.Add(New XElement(RelationshipNs + "Relationship", New XAttribute("Id", "rIdSummit" & suffix),
                                                     New XAttribute("Type", OfficeRel & "customXml"), New XAttribute("Target", "../" & partName)))
                            WritePart(package, relPath, rels, changedParts)
                            Dim types = ReadPart(package, "[Content_Types].xml")
                            types.Root.Add(New XElement(ContentTypeNs + "Override", New XAttribute("PartName", "/" & partName), New XAttribute("ContentType", "application/xml")))
                            types.Root.Add(New XElement(ContentTypeNs + "Override", New XAttribute("PartName", "/" & propsName), New XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml")))
                            WritePart(package, "[Content_Types].xml", types, changedParts)
                        End If
                    End Using
                    sourceStream.Position = 0
                    Using original As New ZipArchive(sourceStream, ZipArchiveMode.Read, True), output = ZipFile.OpenRead(stage)
                        For Each entry In original.Entries
                            If changedParts.Contains(entry.FullName) Then Continue For
                            Dim copy = output.GetEntry(entry.FullName)
                            If copy Is Nothing Then Throw New InvalidDataException("Package part lost: " & entry.FullName)
                            Using first = entry.Open(), second = copy.Open()
                                If HashStream(first) <> HashStream(second) Then Throw New InvalidDataException("Package part changed: " & entry.FullName)
                            End Using
                        Next
                    End Using
                    Dim check = ReadEmbedded(stage)
                    If check Is Nothing OrElse Not XNode.DeepEquals(ToDocument(check), xml) Then Throw New InvalidDataException("Embedded definition verification failed.")
                    File.Move(stage, destination)
                End Using
            Finally
                If File.Exists(stage) Then File.Delete(stage)
            End Try
        End Sub

        Private Shared Function ReadPart(package As ZipArchive, name As String) As XDocument
            Dim entry = package.GetEntry(name)
            If entry Is Nothing Then Throw New InvalidDataException("Required package part missing: " & name)
            Using stream = entry.Open()
                Return ReadXml(stream)
            End Using
        End Function

        Private Shared Sub WritePart(package As ZipArchive, name As String, document As XDocument, changed As HashSet(Of String))
            Dim old = package.GetEntry(name)
            If old IsNot Nothing Then old.Delete()
            Using stream = package.CreateEntry(name, CompressionLevel.Optimal).Open()
                document.Save(stream)
            End Using
            changed.Add(name)
        End Sub
    End Class
End Namespace
