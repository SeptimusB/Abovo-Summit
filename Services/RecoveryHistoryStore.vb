Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
Imports System.Xml
Imports System.Xml.Linq

Namespace Abovo
    'Display-only evidence. No cell snapshots, executable commands or undo stack
    'are deserialised from a previous process. Stored in standard custom XML.
    Friend NotInheritable Class RecoveryHistoryStore
        Friend Const HistoryNamespace As String = "urn:abovo:summit:recovery-history:1"
        Private Shared ReadOnly Ns As XNamespace = HistoryNamespace
        Private Shared ReadOnly Rel As XNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
        Private Shared ReadOnly Types As XNamespace = "http://schemas.openxmlformats.org/package/2006/content-types"
        Private Const OfficeRel As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
        Friend Const MaximumRows As Integer = 1000

        Private Shared Function ReadXml(stream As Stream) As XDocument
            Using reader = XmlReader.Create(stream, New XmlReaderSettings With {.DtdProcessing = DtdProcessing.Prohibit, .XmlResolver = Nothing, .MaxCharactersInDocument = 16000000})
                Return XDocument.Load(reader)
            End Using
        End Function

        Private Shared Function Find(package As ZipArchive) As ZipArchiveEntry
            Dim found As ZipArchiveEntry = Nothing
            For Each part In package.Entries
                If Not part.FullName.StartsWith("customXml/", StringComparison.Ordinal) OrElse Not part.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) Then Continue For
                Using content = part.Open(), reader = XmlReader.Create(content, New XmlReaderSettings With {.DtdProcessing = DtdProcessing.Prohibit, .XmlResolver = Nothing, .MaxCharactersInDocument = 16000000})
                    reader.MoveToContent()
                    If reader.LocalName <> "RecoveryHistory" OrElse reader.NamespaceURI <> HistoryNamespace Then Continue For
                    If found IsNot Nothing Then Throw New InvalidDataException("Multiple recovery history parts.")
                    found = part
                End Using
            Next
            Return found
        End Function

        Private Shared Sub Put(package As ZipArchive, name As String, xml As XDocument)
            Dim old = package.GetEntry(name)
            If old IsNot Nothing Then old.Delete()
            Using content = package.CreateEntry(name, CompressionLevel.Optimal).Open()
                xml.Save(content, SaveOptions.DisableFormatting)
            End Using
        End Sub

        Friend Shared Sub Write(filePath As String, manager As ModelChangeManagerV2)
            WriteDocument(filePath, ToDocument(manager.GetHistoryTable()))
        End Sub

        Friend Shared Function ToDocument(table As DataTable) As XDocument
            Dim document As New XDocument(New XElement(Ns + "RecoveryHistory", New XAttribute("version", "1"), New XAttribute("limit", MaximumRows)))
            For Each row As DataRow In table.Rows.Cast(Of DataRow)().OrderByDescending(Function(r) CDate(r("TimeStamp"))).Take(MaximumRows)
                Dim item As New XElement(Ns + "Edit")
                For Each column As DataColumn In table.Columns
                    If column.ColumnName = "Action" Then Continue For
                    Dim value = If(column.DataType Is GetType(DateTime), CDate(row(column)).ToString("O", CultureInfo.InvariantCulture), Convert.ToString(row(column), CultureInfo.InvariantCulture))
                    If value.Length > 2048 Then value = value.Substring(0, 2048) & " [truncated in recovery history]"
                    item.Add(New XElement(Ns + column.ColumnName, value))
                Next
                document.Root.Add(item)
            Next
            Return document
        End Function

        ' Only private save candidates/recovery files call this writer. Return
        ' the exact permitted package delta for independent preservation checks.
        Friend Shared Function WriteDocument(filePath As String, document As XDocument) As HashSet(Of String)
            If document Is Nothing OrElse document.Root Is Nothing OrElse document.Root.Name <> Ns + "RecoveryHistory" OrElse
                CStr(document.Root.Attribute("version")) <> "1" Then Throw New InvalidDataException("Unsupported history document.")
            Dim changed As New HashSet(Of String)(StringComparer.Ordinal)
            Using package = ZipFile.Open(filePath, ZipArchiveMode.Update)
                If package.Entries.Any(Function(e) e.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)) Then Throw New InvalidDataException("Signed package recovery history requires a separate signing workflow.")
                Dim existing = Find(package)
                If existing IsNot Nothing Then
                    Using content = existing.Open()
                        If CStr(ReadXml(content).Root.Attribute("version")) <> "1" Then Throw New InvalidDataException("Unsupported existing history version.")
                    End Using
                    Put(package, existing.FullName, document)
                    changed.Add(existing.FullName)
                    Return changed
                End If
                Dim binary = package.GetEntry("xl/workbook.bin") IsNot Nothing
                If binary = (package.GetEntry("xl/workbook.xml") IsNot Nothing) Then Throw New InvalidDataException("Ambiguous or missing workbook part.")
                Dim relPath = If(binary, "xl/_rels/workbook.bin.rels", "xl/_rels/workbook.xml.rels")
                If package.GetEntry(relPath) Is Nothing Then Throw New InvalidDataException("Workbook relationships are missing.")
                Dim suffix = Guid.NewGuid().ToString("N")
                Dim part = "customXml/summitRecovery-" & suffix & ".xml"
                Dim props = "customXml/summitRecoveryProps-" & suffix & ".xml"
                Put(package, part, document)
                Dim ds As XNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
                Put(package, props, New XDocument(New XElement(ds + "datastoreItem", New XAttribute(ds + "itemID", "{" & Guid.NewGuid().ToString().ToUpperInvariant() & "}"),
                    New XElement(ds + "schemaRefs", New XElement(ds + "schemaRef", New XAttribute(ds + "uri", HistoryNamespace))))))
                Dim partRels = "customXml/_rels/" & Path.GetFileName(part) & ".rels"
                Put(package, partRels, New XDocument(New XElement(Rel + "Relationships", New XElement(Rel + "Relationship",
                    New XAttribute("Id", "rId1"), New XAttribute("Type", OfficeRel & "customXmlProps"), New XAttribute("Target", Path.GetFileName(props))))))
                Dim relationships As XDocument, contentTypes As XDocument
                Using content = package.GetEntry(relPath).Open()
                    relationships = ReadXml(content)
                End Using
                relationships.Root.Add(New XElement(Rel + "Relationship", New XAttribute("Id", "rIdRecovery" & suffix), New XAttribute("Type", OfficeRel & "customXml"), New XAttribute("Target", "../" & part)))
                Put(package, relPath, relationships)
                Using content = package.GetEntry("[Content_Types].xml").Open()
                    contentTypes = ReadXml(content)
                End Using
                contentTypes.Root.Add(New XElement(Types + "Override", New XAttribute("PartName", "/" & part), New XAttribute("ContentType", "application/xml")))
                contentTypes.Root.Add(New XElement(Types + "Override", New XAttribute("PartName", "/" & props), New XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml")))
                Put(package, "[Content_Types].xml", contentTypes)
                changed.UnionWith({part, props, partRels, relPath, "[Content_Types].xml"})
            End Using
            Return changed
        End Function

        Friend Shared Function Read(filePath As String, schema As DataTable) As DataTable
            Dim result = schema.Clone()
            Using package = ZipFile.OpenRead(filePath)
                Dim part = Find(package)
                If part Is Nothing Then Return result
                Dim document As XDocument
                Using content = part.Open()
                    document = ReadXml(content)
                End Using
                If CStr(document.Root.Attribute("version")) <> "1" Then Throw New InvalidDataException("Unsupported recovery history version.")
                For Each item In document.Root.Elements(Ns + "Edit").Take(MaximumRows)
                    Dim row = result.NewRow()
                    For Each column As DataColumn In result.Columns
                        Dim value = CStr(item.Element(Ns + column.ColumnName))
                        If column.DataType Is GetType(DateTime) Then
                            row(column) = DateTime.Parse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind)
                        ElseIf column.DataType Is GetType(Integer) Then
                            row(column) = Integer.Parse(value, CultureInfo.InvariantCulture)
                        Else
                            row(column) = If(value, String.Empty)
                        End If
                    Next
                    row("GroupID") = -(result.Rows.Count + 1)
                    row("Action") = String.Empty
                    If Not Convert.ToString(row("State")).StartsWith("Recovered: ", StringComparison.Ordinal) Then row("State") = "Recovered: " & Convert.ToString(row("State"))
                    result.Rows.Add(row)
                Next
            End Using
            Return result
        End Function
    End Class
End Namespace
