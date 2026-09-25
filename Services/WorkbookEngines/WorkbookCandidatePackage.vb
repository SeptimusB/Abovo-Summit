Option Strict On

Imports System.IO
Imports System.IO.Compression
Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml
Imports System.Xml.Linq

Namespace Abovo.WorkbookEngines
    ' Conservative same-format value-edit gate. This does not transplant opaque
    ' metadata across structural changes or claim a whole-workbook integrity check.
    Friend NotInheritable Class WorkbookCandidatePackage
        Private ReadOnly parts As New SortedDictionary(Of String, String)(StringComparer.Ordinal)
        Private ReadOnly links As New SortedSet(Of String)(StringComparer.Ordinal)
        Private ReadOnly customParts As New Dictionary(Of String, Byte())(StringComparer.Ordinal)
        Private ReadOnly customPayloads As New Dictionary(Of String, String)(StringComparer.Ordinal)
        Private ReadOnly vbaParts As New Dictionary(Of String, Byte())(StringComparer.Ordinal)
        Private customTypes As List(Of XElement)

        Friend Shared Function Capture(path As String) As WorkbookCandidatePackage
            Dim result As New WorkbookCandidatePackage()
            Using zip = ZipFile.OpenRead(path)
                If zip.Entries.Count > 20000 OrElse zip.GetEntry("[Content_Types].xml") Is Nothing Then Throw New InvalidDataException("Unsupported workbook package.")
                Dim typesDocument = ReadXml(zip.GetEntry("[Content_Types].xml"))
                Dim extension = IO.Path.GetExtension(path).ToLowerInvariant()
                Dim mainPart = If(extension = ".xlsb", "/xl/workbook.bin", "/xl/workbook.xml")
                Dim mainType = If(extension = ".xlsb", "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
                    If(extension = ".xlsm", "application/vnd.ms-excel.sheet.macroEnabled.main+xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"))
                Dim declaration = typesDocument.Root.Elements().SingleOrDefault(Function(e) CStr(e.Attribute("PartName")) = mainPart)
                If declaration Is Nothing Then declaration = typesDocument.Root.Elements().SingleOrDefault(Function(e) CStr(e.Attribute("Extension")) = IO.Path.GetExtension(mainPart).TrimStart("."c))
                If Not {".xlsb", ".xlsm", ".xlsx"}.Contains(extension) OrElse zip.GetEntry(mainPart.Substring(1)) Is Nothing OrElse
                    declaration Is Nothing OrElse CStr(declaration.Attribute("ContentType")) <> mainType Then
                    Throw New InvalidDataException("Workbook package format does not match its file extension.")
                End If
                Dim names As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim customBytes As Long
                For Each part In zip.Entries
                    If Not names.Add(part.FullName) Then Throw New InvalidDataException("Duplicate workbook package part.")
                    If part.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase) Then Throw New InvalidDataException("Digitally signed workbook packages need a separate signed-save policy.")
                    If part.FullName.IndexOf("vbaProjectSignature", StringComparison.OrdinalIgnoreCase) >= 0 Then Throw New InvalidDataException("Signed VBA projects need a separate signed-save policy.")
                    If IsCritical(part.FullName) Then result.parts.Add(part.FullName, PartHash(part))
                    If part.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase) Then
                        customBytes += part.Length
                        If customBytes > 67108864 Then Throw New InvalidDataException("Custom XML exceeds the trial preservation limit.")
                        Using content = part.Open(), buffer As New MemoryStream()
                            content.CopyTo(buffer)
                            result.customParts.Add(part.FullName, buffer.ToArray())
                        End Using
                    End If
                    If String.Equals(part.FullName, "xl/vbaProject.bin", StringComparison.OrdinalIgnoreCase) Then
                        If part.Length > 67108864 Then Throw New InvalidDataException("VBA package exceeds the trial preservation limit.")
                        Using content = part.Open(), buffer As New MemoryStream()
                            content.CopyTo(buffer)
                            result.vbaParts.Add(part.FullName, buffer.ToArray())
                        End Using
                    End If
                    If part.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) Then
                        Dim xml = ReadXml(part)
                        For Each relation In xml.Root.Elements()
                            Dim type = CStr(relation.Attribute("Type")), target = CStr(relation.Attribute("Target"))
                            If type Is Nothing OrElse target Is Nothing Then Throw New InvalidDataException("Invalid package relationship.")
                            If type.EndsWith("/customXml", StringComparison.Ordinal) OrElse type.EndsWith("/customXmlProps", StringComparison.Ordinal) OrElse
                                type.EndsWith("/vbaProject", StringComparison.Ordinal) OrElse type.EndsWith("/sheetMetadata", StringComparison.Ordinal) Then
                                If CStr(relation.Attribute("TargetMode")) = "External" Then Throw New InvalidDataException("External preservation-critical relationship is unsupported.")
                                Dim root = RelationshipOwner(part.FullName)
                                Dim resolved = Resolve(root, target)
                                If zip.GetEntry(resolved) Is Nothing Then Throw New InvalidDataException("Missing preservation-critical package part: " & resolved)
                                result.links.Add(root & "|" & type & "|" & resolved)
                                If type.EndsWith("/customXml", StringComparison.Ordinal) Then
                                    If Not resolved.StartsWith("customXml/", StringComparison.Ordinal) Then Throw New InvalidDataException("Nonstandard custom XML location needs a separate preservation policy.")
                                    Dim key = PartHash(zip.GetEntry(resolved))
                                    If result.customPayloads.ContainsKey(key) Then Throw New InvalidDataException("Ambiguous duplicate custom XML payloads need a separate preservation policy.")
                                    result.customPayloads.Add(key, resolved)
                                End If
                            End If
                        Next
                    End If
                Next
                result.customTypes = typesDocument.Root.Elements().Where(
                    Function(e) CStr(e.Attribute("PartName")) IsNot Nothing AndAlso CStr(e.Attribute("PartName")).StartsWith("/customXml/", StringComparison.Ordinal)).Select(Function(e) New XElement(e)).ToList()
            End Using
            Return result
        End Function

        Friend Sub PrepareValueOnlyCandidate(path As String)
            PreserveUnchangedCustomXml(path)
            If vbaParts.Count = 0 Then Return
            ' No exposed command changes VBA, sheet structure or code names in
            ' this trial. Preserve the original opaque project, not a native
            ' exporter's recompressed/recompiled image. Never extract VBA source.
            Using zip = ZipFile.Open(path, ZipArchiveMode.Update)
                For Each pair In vbaParts
                    Dim part = zip.GetEntry(pair.Key)
                    If part Is Nothing Then Throw New InvalidDataException("Candidate omitted the VBA project.")
                    part.Delete()
                    Using output = zip.CreateEntry(pair.Key).Open()
                        output.Write(pair.Value, 0, pair.Value.Length)
                    End Using
                Next
            End Using
        End Sub

        Private Shared Function RelationshipOwner(name As String) As String
            Dim root = name.Replace("/_rels/", "/")
            If root.StartsWith("_rels/", StringComparison.Ordinal) Then root = root.Substring(6)
            Return root.Substring(0, root.Length - 5)
        End Function

        Private Shared Function Resolve(owner As String, target As String) As String
            Return Uri.UnescapeDataString(New Uri(New Uri("http://package/" & owner), target).AbsolutePath.TrimStart("/"c))
        End Function

        ' Value-edit sessions expose no XML mutation API. Only when every native
        ' payload is semantically unchanged may we retain its original opaque
        ' companions. No dynamic-array parts, cells or formulas are transplanted.
        Friend Sub PreserveUnchangedCustomXml(path As String)
            If customParts.Count = 0 Then Return
            Dim actual = Capture(path)
            If customPayloads.Count <> actual.customPayloads.Count OrElse customPayloads.Keys.Any(Function(k) Not actual.customPayloads.ContainsKey(k)) Then
                Throw New InvalidDataException("Native export changed custom XML content; preservation cannot discard those changes.")
            End If
            Dim mapped As New Dictionary(Of String, String)(StringComparer.Ordinal)
            For Each pair In customPayloads
                mapped.Add(actual.customPayloads(pair.Key), pair.Value)
            Next
            ' Reject unexpected/unreferenced parts, rather than throwing them
            ' away merely because the known payloads matched.
            If actual.customParts.Count <> customParts.Count Then Throw New InvalidDataException("Native export changed the custom XML bundle shape.")
            Dim changed As New Dictionary(Of String, XDocument)(StringComparer.Ordinal)
            Using zip = ZipFile.OpenRead(path)
                For Each part In zip.Entries.Where(Function(e) e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) AndAlso Not e.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase))
                    Dim document = ReadXml(part), altered As Boolean = False
                    Dim owner = RelationshipOwner(part.FullName)
                    For Each relation In document.Root.Elements()
                        If CStr(relation.Attribute("TargetMode")) = "External" Then Continue For
                        Dim target = CStr(relation.Attribute("Target"))
                        If target Is Nothing Then Throw New InvalidDataException("Invalid package relationship.")
                        Dim resolved = Resolve(owner, target)
                        If Not resolved.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase) Then Continue For
                        If Not CStr(relation.Attribute("Type")).EndsWith("/customXml", StringComparison.Ordinal) OrElse Not mapped.ContainsKey(resolved) Then Throw New InvalidDataException("Unsupported custom XML binding relationship.")
                        Dim relative = New Uri("http://package/" & owner).MakeRelativeUri(New Uri("http://package/" & mapped(resolved))).ToString()
                        relation.SetAttributeValue("Target", relative)
                        altered = True
                    Next
                    If altered Then changed.Add(part.FullName, document)
                Next
                Dim types = ReadXml(zip.GetEntry("[Content_Types].xml"))
                types.Root.Elements().Where(Function(e) CStr(e.Attribute("PartName")) IsNot Nothing AndAlso CStr(e.Attribute("PartName")).StartsWith("/customXml/", StringComparison.Ordinal)).Remove()
                For Each item In customTypes
                    types.Root.Add(New XElement(item))
                Next
                changed.Add("[Content_Types].xml", types)
            End Using
            Using zip = ZipFile.Open(path, ZipArchiveMode.Update)
                For Each part In zip.Entries.Where(Function(e) e.FullName.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase)).ToList()
                    part.Delete()
                Next
                For Each pair In customParts
                    Using output = zip.CreateEntry(pair.Key).Open()
                        output.Write(pair.Value, 0, pair.Value.Length)
                    End Using
                Next
                For Each pair In changed
                    zip.GetEntry(pair.Key).Delete()
                    Using output = zip.CreateEntry(pair.Key).Open()
                        pair.Value.Save(output)
                    End Using
                Next
            End Using
        End Sub

        Private Shared Function IsCritical(name As String) As Boolean
            Return (name.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase) AndAlso Not name.Contains("/_rels/")) OrElse
                name.StartsWith("xl/vbaProject", StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(name, "xl/metadata.xml", StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(name, "xl/metadata.bin", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function ReadXml(part As ZipArchiveEntry) As XDocument
            If part.Length > 67108864 Then Throw New InvalidDataException("Preservation part exceeds the trial inspection limit.")
            Using content = part.Open(), reader = XmlReader.Create(content, New XmlReaderSettings With {
                .DtdProcessing = DtdProcessing.Prohibit, .XmlResolver = Nothing, .MaxCharactersInDocument = 67108864})
                Return XDocument.Load(reader)
            End Using
        End Function

        Private Shared Function PartHash(part As ZipArchiveEntry) As String
            Using hash = SHA256.Create()
                If part.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) Then
                    Dim text As New StringBuilder()
                    Dim document = ReadXml(part)
                    For Each node In document.Nodes()
                        If TypeOf node Is XText AndAlso String.IsNullOrWhiteSpace(DirectCast(node, XText).Value) Then Continue For
                        If TypeOf node Is XElement Then
                            Canonical(DirectCast(node, XElement), text)
                        Else
                            Token(node.ToString(), text)
                        End If
                    Next
                    Return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(text.ToString())))
                End If
                Using content = part.Open()
                    Return Convert.ToBase64String(hash.ComputeHash(content))
                End Using
            End Using
        End Function

        Private Shared Sub Token(text As String, output As StringBuilder)
            output.Append(text.Length).Append(":"c).Append(text)
        End Sub

        Private Shared Sub Canonical(node As XElement, output As StringBuilder)
            Token(node.Name.ToString(), output)
            For Each attr In node.Attributes().Where(Function(a) Not a.IsNamespaceDeclaration).OrderBy(Function(a) a.Name.ToString(), StringComparer.Ordinal)
                output.Append("A"c) : Token(attr.Name.ToString(), output) : Token(attr.Value, output)
            Next
            For Each child In node.Nodes()
                If TypeOf child Is XElement Then
                    output.Append("E"c) : Canonical(DirectCast(child, XElement), output)
                ElseIf TypeOf child Is XText Then
                    Dim value = DirectCast(child, XText).Value
                    ' DevExpress removes element-only line indentation from
                    ' SharePoint schema parts. Do not ignore leaf/mixed text or
                    ' xml:space-preserved content when matching payloads.
                    Dim space = node.AncestorsAndSelf().Attributes(XNamespace.Xml + "space").FirstOrDefault()
                    If String.IsNullOrWhiteSpace(value) AndAlso (value.Contains(vbLf) OrElse value.Contains(vbCr)) AndAlso
                        node.Elements().Any() AndAlso Not node.Nodes().OfType(Of XText)().Any(Function(t) Not String.IsNullOrWhiteSpace(t.Value)) AndAlso
                        (space Is Nothing OrElse space.Value <> "preserve") Then Continue For
                    output.Append("T"c) : Token(value, output)
                Else
                    output.Append("X"c) : Token(child.ToString(), output)
                End If
            Next
            output.Append("/"c)
        End Sub

        Friend Sub Verify(path As String)
            Dim actual = Capture(path)
            If parts.Count <> actual.parts.Count OrElse Not links.SetEquals(actual.links) Then Throw New InvalidDataException("Candidate changed the custom XML, VBA or dynamic-array package structure.")
            For Each pair In parts
                Dim value As String = Nothing
                If Not actual.parts.TryGetValue(pair.Key, value) OrElse value <> pair.Value Then Throw New InvalidDataException("Candidate did not preserve " & pair.Key & ". The original file is unchanged.")
            Next
        End Sub
    End Class
End Namespace
