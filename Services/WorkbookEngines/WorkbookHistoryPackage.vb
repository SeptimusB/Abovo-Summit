Option Strict On

Imports System.IO
Imports System.IO.Compression
Imports System.Security.Cryptography
Imports System.Xml
Imports System.Xml.Linq

Namespace Abovo.WorkbookEngines
    ' Only an already verified, private same-format candidate enters here.
    ' Audit the entire ZIP delta so an owned history update cannot silently
    ' approve a change to unrelated XML, worksheets, VBA or array metadata.
    Friend NotInheritable Class WorkbookHistoryPackage
        Private Shared ReadOnly Rel As XNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
        Private Shared ReadOnly Types As XNamespace = "http://schemas.openxmlformats.org/package/2006/content-types"
        Private Shared ReadOnly Ds As XNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
        Private Const OfficeRel As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"

        Private Shared Function Fingerprints(path As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.Ordinal)
            Dim total As Long
            Using zip = ZipFile.OpenRead(path)
                If zip.Entries.Count > 20000 Then Throw New InvalidDataException("History candidate package is too large.")
                For Each entry In zip.Entries
                    total += entry.Length
                    If total > 2147483648L Then Throw New InvalidDataException("History candidate exceeds the 2 GB expanded inspection limit.")
                    Using stream = entry.Open(), algorithm = SHA256.Create()
                        result.Add(entry.FullName, Convert.ToBase64String(algorithm.ComputeHash(stream)))
                    End Using
                Next
            End Using
            Return result
        End Function

        Private Shared Function Document(zip As ZipArchive, path As String) As XDocument
            Dim entry = zip.GetEntry(path)
            If entry Is Nothing OrElse entry.Length > 16000000 Then Throw New InvalidDataException("Missing or oversized history package part.")
            Using stream = entry.Open(), reader = XmlReader.Create(stream, New XmlReaderSettings With {
                .DtdProcessing = DtdProcessing.Prohibit, .XmlResolver = Nothing, .MaxCharactersInDocument = 16000000})
                Return XDocument.Load(reader)
            End Using
        End Function

        Private Shared Sub RequireAppend(original As XDocument, updated As XDocument, additions As IEnumerable(Of XElement))
            Dim expected = New XDocument(original)
            expected.Root.Add(additions)
            If Not Equivalent(expected, updated) Then Throw New InvalidDataException("History write changed unrelated package declarations.")
        End Sub

        Private Shared Function Equivalent(expected As XDocument, actual As XDocument) As Boolean
            Dim left = New XDocument(expected), right = New XDocument(actual)
            ' Prefixes/default declarations are serialization details; expanded
            ' element/attribute names and all values/nodes must still agree.
            left.Descendants().Attributes().Where(Function(a) a.IsNamespaceDeclaration).Remove()
            right.Descendants().Attributes().Where(Function(a) a.IsNamespaceDeclaration).Remove()
            Return XNode.DeepEquals(left, right)
        End Function

        Friend Shared Function Apply(path As String, snapshot As ModelHistorySnapshot) As WorkbookCandidatePackage
            Dim before = Fingerprints(path)
            Dim relPath = If(before.ContainsKey("xl/workbook.bin"), "xl/_rels/workbook.bin.rels", "xl/_rels/workbook.xml.rels")
            Dim originalRels As XDocument, originalTypes As XDocument
            Using zip = ZipFile.OpenRead(path)
                originalRels = Document(zip, relPath) : originalTypes = Document(zip, "[Content_Types].xml")
            End Using
            Dim expected = snapshot.Document()
            Dim changed = RecoveryHistoryStore.WriteDocument(path, expected)
            Dim after = Fingerprints(path)
            If before.Keys.Any(Function(k) Not after.ContainsKey(k)) OrElse after.Keys.Any(Function(k) Not before.ContainsKey(k) AndAlso Not changed.Contains(k)) Then
                Throw New InvalidDataException("History write removed or introduced unrelated package parts.")
            End If
            For Each pair In before
                If Not changed.Contains(pair.Key) AndAlso after(pair.Key) <> pair.Value Then Throw New InvalidDataException("History write changed " & pair.Key)
            Next
            Dim payload As String
            Using zip = ZipFile.OpenRead(path)
                If changed.Count = 1 Then
                    payload = changed.Single()
                    If Not before.ContainsKey(payload) OrElse Not payload.StartsWith("customXml/", StringComparison.Ordinal) Then Throw New InvalidDataException("Invalid history replacement.")
                ElseIf changed.Count = 5 Then
                    Dim added = changed.Where(Function(k) Not before.ContainsKey(k)).ToList()
                    If added.Count <> 3 OrElse Not changed.Contains(relPath) OrElse Not changed.Contains("[Content_Types].xml") Then Throw New InvalidDataException("Invalid history package additions.")
                    payload = added.Single(Function(k) k.StartsWith("customXml/summitRecovery-", StringComparison.Ordinal))
                    Dim props = added.Single(Function(k) k.StartsWith("customXml/summitRecoveryProps-", StringComparison.Ordinal))
                    Dim linksPath = "customXml/_rels/" & IO.Path.GetFileName(payload) & ".rels"
                    If Not added.Contains(linksPath) Then Throw New InvalidDataException("Missing history properties relationship.")
                    Dim suffix = IO.Path.GetFileNameWithoutExtension(payload).Substring("summitRecovery-".Length)
                    RequireAppend(originalRels, Document(zip, relPath), {New XElement(Rel + "Relationship", New XAttribute("Id", "rIdRecovery" & suffix),
                        New XAttribute("Type", OfficeRel & "customXml"), New XAttribute("Target", "../" & payload))})
                    RequireAppend(originalTypes, Document(zip, "[Content_Types].xml"), {
                        New XElement(Types + "Override", New XAttribute("PartName", "/" & payload), New XAttribute("ContentType", "application/xml")),
                        New XElement(Types + "Override", New XAttribute("PartName", "/" & props), New XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"))})
                    Dim expectedLinks = New XDocument(New XElement(Rel + "Relationships", New XElement(Rel + "Relationship",
                        New XAttribute("Id", "rId1"), New XAttribute("Type", OfficeRel & "customXmlProps"), New XAttribute("Target", IO.Path.GetFileName(props)))))
                    If Not Equivalent(expectedLinks, Document(zip, linksPath)) Then Throw New InvalidDataException("Invalid history properties link.")
                    Dim properties = Document(zip, props).Root, itemId As Guid
                    If Not Guid.TryParse(CStr(properties.Attribute(Ds + "itemID")), itemId) Then Throw New InvalidDataException("Invalid history properties ID.")
                    Dim expectedProperties = New XElement(Ds + "datastoreItem", New XAttribute(Ds + "itemID", CStr(properties.Attribute(Ds + "itemID"))),
                        New XElement(Ds + "schemaRefs", New XElement(Ds + "schemaRef", New XAttribute(Ds + "uri", RecoveryHistoryStore.HistoryNamespace))))
                    ' Ignore serializer namespace-prefix declarations, not data.
                    properties.Attributes().Where(Function(a) a.IsNamespaceDeclaration).Remove()
                    If Not XNode.DeepEquals(expectedProperties, properties) Then Throw New InvalidDataException("Invalid history properties.")
                Else
                    Throw New InvalidDataException("Unexpected history package mutation.")
                End If
                If Not Equivalent(expected, Document(zip, payload)) Then Throw New InvalidDataException("Written history differs from the captured snapshot.")
            End Using
            Return WorkbookCandidatePackage.Capture(path)
        End Function
    End Class
End Namespace
