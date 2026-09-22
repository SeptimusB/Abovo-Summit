Option Infer On
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports DevExpress.Spreadsheet

Namespace Abovo
    Public NotInheritable Class BalanceSheetSnapshot
        Public Const BundleName As String = "TDB_BS_Bundle"
        Public Const ExtraRangeName As String = "TDB_BS_Extra_Inputs"
        Public Const Marker As String = "Summit SOFP snapshot 1"
        Private Const BundleColumn As Integer = 75 'BX: outside the 74-column transaction contract.
        Private Const ChunkSize As Integer = 30000

        'Before ANY clear: reject client-owned content outside previously-owned outputs.
        Public Shared Sub ValidateDestination(sheet As Worksheet, oldTransactionName As String)
            Dim owned As New List(Of CellRange)
            For Each name In {oldTransactionName, BundleName, ExtraRangeName}
                Dim definition = sheet.DefinedNames.GetDefinedName(name)
                If definition Is Nothing Then Continue For
                Dim range = definition.Range
                If range Is Nothing OrElse range.Worksheet IsNot sheet Then Throw New InvalidOperationException("Conflicting snapshot name: " & name)
                If name = BundleName AndAlso (range.LeftColumnIndex <> BundleColumn OrElse range.ColumnCount <> 1 OrElse range.TopRowIndex <> 0 OrElse range.RowCount > 2000) Then Throw New InvalidOperationException("Conflicting Balance Sheet snapshot region.")
                If name = ExtraRangeName AndAlso (range.LeftColumnIndex <> 0 OrElse range.ColumnCount <> 74) Then Throw New InvalidOperationException("Conflicting Balance Sheet input region.")
                owned.Add(range)
            Next
            Dim used = sheet.GetUsedRange()
            For r = used.TopRowIndex To used.BottomRowIndex
                For c = used.LeftColumnIndex To used.RightColumnIndex
                    Dim cell = sheet.Cells(r, c)
                    If cell.Value.IsEmpty AndAlso Not cell.HasFormula Then Continue For
                    Dim cellRow = r, cellColumn = c
                    If Not owned.Any(Function(range) cellRow >= range.TopRowIndex AndAlso cellRow <= range.BottomRowIndex AndAlso cellColumn >= range.LeftColumnIndex AndAlso cellColumn <= range.RightColumnIndex) Then
                        Throw New InvalidOperationException("Snapshot creation would overwrite unrecognised content at '" & sheet.Name & "'!" & cell.GetReferenceA1() & ". Nothing was cleared.")
                    End If
                Next
            Next
        End Sub

        Public Shared Sub Capture(workbook As IWorkbook, document As BalanceSheetDocument)
            Dim source = workbook.Worksheets(TransactionalDBSnapshotManager.SourceWorksheetName)
            Dim snapshot = workbook.Worksheets(TransactionalDBSnapshotManager.SnapshotWorksheetName)
            Dim comparison = workbook.Worksheets(TransactionalDBSnapshotManager.ComparisonWorksheetName)
            Dim transactionRange = workbook.DefinedNames.GetDefinedName(TransactionalDBSnapshotManager.SourceRangeName).Range
            'Includes the reserve's extra source rows, calculated SOFP and the
            'intervening original output cells, at their original Excel addresses.
            Dim extra = source.Range.FromLTRB(0, transactionRange.BottomRowIndex + 1, 73, document.OutputBottom)
            For Each sheet In {snapshot, comparison}
                Dim target = sheet.Range.FromLTRB(extra.LeftColumnIndex, extra.TopRowIndex, extra.RightColumnIndex, extra.BottomRowIndex)
                target.CopyFrom(extra, PasteSpecial.Values)
                target.CopyFrom(extra, PasteSpecial.Formats)
                SetName(sheet, ExtraRangeName, target)
            Next
            For Each node In document.Nodes.Where(Function(n) n.IsHeadline)
                Dim index = Integer.Parse(node.Id.Substring(3))
                Dim row = document.OutputTop + index
                For p = 0 To 40
                    Dim c = BalanceSheetStatement.PeriodColumn(p)
                    Dim address = source.Cells(row, c).GetReferenceA1(ReferenceElement.ColumnAbsolute Or ReferenceElement.RowAbsolute)
                    comparison.Cells(row, c).FormulaInvariant = "='Transactional DB'!" & address & "-'TDB Snapshot'!" & address
                Next
            Next
            Dim xml = Serialize(document)
            Dim count = CInt(Math.Ceiling(xml.Length / CDbl(ChunkSize)))
            If count > 1997 Then Throw New InvalidOperationException("Balance Sheet snapshot is too large.")
            Dim bundle = snapshot.Range.FromLTRB(BundleColumn, 0, BundleColumn, count + 2)
            SetName(snapshot, BundleName, bundle)
            bundle(1, 0).Value = BalanceSheetStatement.Digest(xml)
            bundle(2, 0).Value = count
            For i = 0 To count - 1
                bundle(i + 3, 0).Value = xml.Substring(i * ChunkSize, Math.Min(ChunkSize, xml.Length - i * ChunkSize))
            Next
            Dim comparisonBundle = comparison.Range.FromLTRB(BundleColumn, 0, BundleColumn, 2)
            SetName(comparison, BundleName, comparisonBundle)
            comparisonBundle(1, 0).Value = document.Fingerprint
            comparisonBundle(2, 0).Value = document.Geometry
            'Publish markers LAST. The outer snapshot transaction invalidates
            'both outputs on any later verification or EndUpdate failure.
            bundle(0, 0).Value = Marker
            comparisonBundle(0, 0).Value = Marker
        End Sub

        Public Shared Function Read(workbook As IWorkbook, live As BalanceSheetDocument) As BalanceSheetDocument
            If Not workbook.Worksheets.Contains(TransactionalDBSnapshotManager.SnapshotWorksheetName) OrElse Not workbook.Worksheets.Contains(TransactionalDBSnapshotManager.ComparisonWorksheetName) Then Throw New InvalidOperationException("This workbook has no dedicated snapshot worksheets. Use a compatible upgraded template to capture a snapshot.")
            Dim snapshot = workbook.Worksheets(TransactionalDBSnapshotManager.SnapshotWorksheetName)
            Dim comparison = workbook.Worksheets(TransactionalDBSnapshotManager.ComparisonWorksheetName)
            Dim bundle = snapshot.DefinedNames.GetDefinedName(BundleName)?.Range
            Dim other = comparison.DefinedNames.GetDefinedName(BundleName)?.Range
            If bundle Is Nothing OrElse other Is Nothing OrElse bundle(0, 0).Value.ToString() <> Marker OrElse other(0, 0).Value.ToString() <> Marker Then
                Throw New InvalidOperationException("Balance Sheet was not included in this snapshot; create a new snapshot.")
            End If
            If bundle.ColumnCount <> 1 OrElse bundle.LeftColumnIndex <> BundleColumn OrElse bundle.TopRowIndex <> 0 OrElse bundle.RowCount < 4 OrElse bundle.RowCount > 2000 Then Throw New InvalidOperationException("Invalid Balance Sheet snapshot region.")
            If bundle.Worksheet IsNot snapshot OrElse other.Worksheet IsNot comparison OrElse other.ColumnCount <> 1 OrElse other.LeftColumnIndex <> BundleColumn OrElse other.TopRowIndex <> 0 OrElse other.RowCount <> 3 Then Throw New InvalidOperationException("Invalid Balance Sheet comparison region.")
            Dim count As Integer
            If Not Integer.TryParse(bundle(2, 0).Value.ToString(), count) OrElse count <> bundle.RowCount - 3 Then Throw New InvalidOperationException("Incomplete Balance Sheet snapshot.")
            Dim text As New StringBuilder
            For i = 0 To count - 1
                Dim chunk = bundle(i + 3, 0).Value.ToString()
                If chunk.Length > ChunkSize Then Throw New InvalidOperationException("Invalid Balance Sheet snapshot chunk.")
                text.Append(chunk)
            Next
            If BalanceSheetStatement.Digest(text.ToString()) <> bundle(1, 0).Value.ToString() Then Throw New InvalidOperationException("Balance Sheet snapshot data has changed or is incomplete.")
            Dim document As BalanceSheetDocument
            Dim settings As New XmlReaderSettings With {.DtdProcessing = DtdProcessing.Prohibit, .XmlResolver = Nothing, .MaxCharactersInDocument = 60000000}
            Using reader = XmlReader.Create(New StringReader(text.ToString()), settings)
                document = DirectCast(New XmlSerializer(GetType(BalanceSheetDocument)).Deserialize(reader), BalanceSheetDocument)
            End Using
            If document.Version <> 1 OrElse document.Fingerprint <> live.Fingerprint OrElse document.Geometry <> live.Geometry OrElse
                other(1, 0).Value.ToString() <> document.Fingerprint OrElse other(2, 0).Value.ToString() <> document.Geometry Then Throw New InvalidOperationException("Balance Sheet structure has changed since the snapshot; create a new snapshot.")
            If document.Periods Is Nothing OrElse Not document.Periods.SequenceEqual(live.Periods) OrElse document.OutputTop <> live.OutputTop OrElse document.OutputBottom <> live.OutputBottom OrElse document.Nodes.Count > 50000 Then Throw New InvalidOperationException("Invalid Balance Sheet snapshot layout.")
            For Each sheet In {snapshot, comparison}
                Dim extra = sheet.DefinedNames.GetDefinedName(ExtraRangeName)?.Range
                Dim source = workbook.DefinedNames.GetDefinedName(TransactionalDBSnapshotManager.SourceRangeName).Range
                If extra Is Nothing OrElse extra.Worksheet IsNot sheet OrElse extra.LeftColumnIndex <> 0 OrElse extra.ColumnCount <> 74 OrElse extra.TopRowIndex <> source.BottomRowIndex + 1 OrElse extra.BottomRowIndex <> document.OutputBottom Then Throw New InvalidOperationException("Incomplete Balance Sheet snapshot inputs.")
            Next
            Dim ids As New HashSet(Of String)(StringComparer.Ordinal)
            For Each node In document.Nodes
                If Not ids.Add(node.Id) OrElse node.Values Is Nothing OrElse node.Values.Length <> 41 OrElse node.Styles Is Nothing OrElse node.Styles.Length <> 42 OrElse node.Styles.Any(Function(s) s < 0 OrElse s >= document.Styles.Count) Then Throw New InvalidOperationException("Malformed Balance Sheet snapshot nodes.")
                If node.IsHeadline Then
                    Dim row = document.OutputTop + Integer.Parse(node.Id.Substring(3))
                    For p = 0 To 40
                        Dim value = snapshot.Cells(row, BalanceSheetStatement.PeriodColumn(p)).Value
                        'Excel/XLSB persists 15 significant digits; tolerate only
                        'serialization rounding, far below the reconciliation unit.
                        Dim tolerance = Math.Max(0.000000001, Math.Abs(node.Values(p)) * 0.00000000000001)
                        If (value.IsNumeric AndAlso (Double.IsNaN(node.Values(p)) OrElse Math.Abs(value.NumericValue - node.Values(p)) > tolerance)) OrElse (Not value.IsNumeric AndAlso Not Double.IsNaN(node.Values(p))) Then Throw New InvalidOperationException("Captured Balance Sheet cells have been edited at " & snapshot.Cells(row, BalanceSheetStatement.PeriodColumn(p)).GetReferenceA1() & "; create a new snapshot.")
                    Next
                End If
            Next
            If document.Nodes.Any(Function(n) n.ParentId.Length > 0 AndAlso Not ids.Contains(n.ParentId)) Then Throw New InvalidOperationException("Incomplete Balance Sheet snapshot hierarchy.")
            Dim parents = document.Nodes.ToDictionary(Function(n) n.Id, Function(n) n.ParentId, StringComparer.Ordinal)
            For Each node In document.Nodes
                Dim parent = node.ParentId, depth = 0
                While parent.Length > 0
                    depth += 1
                    If depth > 10 Then Throw New InvalidOperationException("Cyclic or excessively deep Balance Sheet hierarchy.")
                    parent = parents(parent)
                End While
            Next
            Return document
        End Function

        Public Shared Sub Invalidate(workbook As IWorkbook)
            For Each sheetName In {TransactionalDBSnapshotManager.SnapshotWorksheetName, TransactionalDBSnapshotManager.ComparisonWorksheetName}
                Dim sheet = workbook.Worksheets(sheetName)
                Dim bundle = sheet.DefinedNames.GetDefinedName(BundleName)?.Range
                If bundle IsNot Nothing AndAlso bundle.Worksheet Is sheet Then bundle(0, 0).ClearContents()
            Next
        End Sub

        Private Shared Function Serialize(document As BalanceSheetDocument) As String
            Using writer As New StringWriter(Globalization.CultureInfo.InvariantCulture)
                Dim serializer As New XmlSerializer(GetType(BalanceSheetDocument))
                serializer.Serialize(writer, document)
                Return writer.ToString()
            End Using
        End Function
        Private Shared Sub SetName(sheet As Worksheet, name As String, range As CellRange)
            Dim existing = sheet.DefinedNames.GetDefinedName(name)
            If existing IsNot Nothing Then
                existing.Range = range
            Else
                sheet.DefinedNames.Add(name, range.GetReferenceA1(ReferenceElement.IncludeSheetName Or ReferenceElement.ColumnAbsolute Or ReferenceElement.RowAbsolute))
            End If
        End Sub
    End Class
End Namespace
