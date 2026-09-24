Imports System.Drawing
Imports System.Globalization
Imports System.Xml.Linq
Imports DevExpress.Spreadsheet

Namespace Abovo
    Public Class FundingScheduleFacility
        Public Property HeaderName As String
        Public Property ColumnIndex As Integer
        Public Property Caption As String
        Public Property LoanName As String
        Public Property FunderName As String
        Public Property FacilityName As String
        Public ReadOnly Property Hint As String
            Get
                Return LoanName & " (" & If(FunderName = "", "No funder", FunderName) & "/" & If(FacilityName = "", "No facility", FacilityName) & ")"
            End Get
        End Property
        Public ReadOnly Property IsInvestment As Boolean
            Get
                Return HeaderName = "Rep_Fund_29"
            End Get
        End Property
        Public Overrides Function ToString() As String
            Return Caption
        End Function
    End Class

    Public Class FundingScheduleTint
        Public Property DateCell As Cell
        Public Property ExpectedDate As DateTime
        Public Property Colour As Color
        Public Property MarkerColour As Color
        Public ReadOnly Property Active As Boolean
            Get
                Dim value = DateCell.Value
                If value.IsDateTime Then Return value.DateTimeValue.Date = ExpectedDate
                'OOXML may reload date values as serial numbers. The anchor is
                'already validated against an actual Funding date range.
                If value.IsNumeric AndAlso value.NumericValue > 0 AndAlso value.NumericValue < 2958466 Then
                    Return DateTime.FromOADate(value.NumericValue).Date = ExpectedDate
                End If
                Return False
            End Get
        End Property
    End Class

    'Summit-only presentation metadata. Native CustomXmlParts round-trip with
    'XLSB/XLSM; no part is added on open or preview. Hidden names let Excel and
    'Summit move the associations without guessing from non-unique captions/dates.
    Public NotInheritable Class FundingScheduleGroups
        Public Const NamespaceUri As String = "urn:abovo:summit:funding-schedules:1"
        Private Shared ReadOnly Ns As XNamespace = NamespaceUri

        Public Shared Function IsInvestment(target As FundingScheduleTarget) As Boolean
            Return target.DateRange.StartsWith("IR_Fund_Inv_", StringComparison.Ordinal)
        End Function

        Public Shared Function Facilities(wb As IWorkbook, investment As Boolean) As List(Of FundingScheduleFacility)
            Dim headerName = If(investment, "Rep_Fund_29", "LoanDescs")
            Dim header = wb.DefinedNames.GetDefinedName(headerName)?.Range
            Dim result As New List(Of FundingScheduleFacility)
            If header Is Nothing Then Return result
            Dim labels = wb.DefinedNames.GetDefinedName("FacilityNames")?.Range
            Dim funders = wb.DefinedNames.GetDefinedName("Rep_Fund_03")?.Range
            For i = 0 To header.ColumnCount - 1
                Dim cell = header.Worksheet.Cells(header.TopRowIndex - If(investment, 1, 0), header.LeftColumnIndex + i), caption = cell.DisplayText.Trim()
                Dim facility = If(Not investment AndAlso labels IsNot Nothing AndAlso i < labels.ColumnCount, labels(0, i).DisplayText.Trim(), "")
                If caption = "" Then caption = If(investment, "Investment", "Loan") & " " & (i + 1).ToString()
                Dim loanName = caption
                Dim funder = If(Not investment AndAlso funders IsNot Nothing AndAlso i < funders.ColumnCount, funders(0, i).DisplayText.Trim(), "")
                If facility <> "" Then caption = facility & " — " & caption
                result.Add(New FundingScheduleFacility With {.HeaderName = headerName, .ColumnIndex = cell.ColumnIndex,
                    .Caption = caption & " [" & cell.GetReferenceA1() & "]", .LoanName = loanName, .FunderName = funder, .FacilityName = facility})
            Next
            Return result
        End Function

        Public Shared Function SoftColour(colour As Color) As Color
            'Bound the tint even when a dark/custom colour is selected.
            Return Color.FromArgb(218 + colour.R * 37 \ 255, 218 + colour.G * 37 \ 255, 218 + colour.B * 37 \ 255)
        End Function

        Private Shared Function Part(wb As IWorkbook) As ICustomXmlPart
            Dim matches = wb.CustomXmlParts.Where(Function(p) p.CustomXmlPartDocument.DocumentElement?.NamespaceURI = NamespaceUri).ToList()
            If matches.Count > 1 Then Throw New InvalidOperationException("Duplicate Funding schedule XML parts; no schedule was added.")
            Return matches.FirstOrDefault()
        End Function

        Public Shared Function Read(wb As IWorkbook) As XElement
            Dim existing = Part(wb)
            If existing Is Nothing Then Return New XElement(Ns + "FundingSchedules", New XAttribute("version", "1"))
            Dim root = XElement.Parse(existing.CustomXmlPartDocument.OuterXml)
            If root.Name <> Ns + "FundingSchedules" OrElse CStr(root.Attribute("version")) <> "1" Then Throw New InvalidOperationException("Unsupported Funding schedule XML version.")
            Return root
        End Function

        Private Shared Function AddAnchor(wb As IWorkbook, name As String, cell As Cell) As String
            Dim added = wb.DefinedNames.Add(name, cell.GetReferenceA1(ReferenceElement.IncludeSheetName Or ReferenceElement.ColumnAbsolute Or ReferenceElement.RowAbsolute))
            added.Hidden = True
            Return name
        End Function

        Public Shared Function Add(wb As IWorkbook, facility As FundingScheduleFacility, colour As Color,
                                   rows As IList(Of Tuple(Of FundingScheduleTarget, Cell, DateTime)), configuration As String) As Action
            Dim root = Read(wb), oldPart = Part(wb), oldXml = If(oldPart Is Nothing, Nothing, oldPart.CustomXmlPartDocument.OuterXml)
            Dim names As New List(Of String), currentPart As ICustomXmlPart = oldPart
            Dim undo As Action = Sub()
                                    If currentPart IsNot Nothing Then
                                        If oldXml Is Nothing Then
                                            wb.CustomXmlParts.Remove(currentPart)
                                        Else
                                            currentPart.CustomXmlPartDocument.LoadXml(oldXml)
                                        End If
                                    End If
                                    For Each name In names
                                        wb.DefinedNames.Remove(name)
                                    Next
                                End Sub
            Try
                Dim id = Guid.NewGuid().ToString("N"), prefix = "_SummitSchedule_" & id & "_"
                Dim header = wb.DefinedNames.GetDefinedName(facility.HeaderName).Range
                Dim facilityCell = header.Worksheet.Cells(header.TopRowIndex - If(facility.IsInvestment, 1, 0), facility.ColumnIndex)
                names.Add(AddAnchor(wb, prefix & "Facility", facilityCell))
                Dim group As New XElement(Ns + "Schedule", New XAttribute("id", id), New XAttribute("facilityAnchor", names.Last()),
                    New XAttribute("facilityHeader", facility.HeaderName), New XAttribute("colour", colour.ToArgb().ToString(CultureInfo.InvariantCulture)),
                    New XAttribute("createdUtc", DateTime.UtcNow.ToString("o")), New XElement(Ns + "Configuration", configuration))
                For Each entry In rows
                    Dim name = AddAnchor(wb, prefix & "Date" & names.Count.ToString(), entry.Item2)
                    names.Add(name)
                    group.Add(New XElement(Ns + "Date", New XAttribute("anchor", name), New XAttribute("section", entry.Item1.DateRange),
                        New XAttribute("values", entry.Item1.ValueRange), New XAttribute("value", entry.Item3.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))))
                Next
                root.Add(group)
                If currentPart Is Nothing Then
                    currentPart = wb.CustomXmlParts.Add(root.ToString())
                Else
                    currentPart.CustomXmlPartDocument.LoadXml(root.ToString())
                End If
                Return undo
            Catch
                undo()
                Throw
            End Try
        End Function

        Public Shared Function IsPresentationAnchor(wb As IWorkbook, name As DefinedName) As Boolean
            If Not name.Hidden OrElse Not name.Name.StartsWith("_SummitSchedule_", StringComparison.Ordinal) Then Return False
            Try
                Return Read(wb).DescendantsAndSelf().Attributes().Any(Function(a) (a.Name = "anchor" OrElse a.Name = "facilityAnchor") AndAlso a.Value = name.Name)
            Catch
                Return False
            End Try
        End Function

        Public Shared Function TintIndex(wb As IWorkbook) As Dictionary(Of String, FundingScheduleTint)
            Dim result As New Dictionary(Of String, FundingScheduleTint)(StringComparer.Ordinal)
            For Each group In Read(wb).Elements(Ns + "Schedule")
                Try
                    Dim facility = ResolveAnchor(wb, CStr(group.Attribute("facilityAnchor")))
                    If facility Is Nothing OrElse facility.RowCount <> 1 OrElse facility.ColumnCount <> 1 OrElse facility.Worksheet.Name <> "Funding Assumptions" Then Continue For
                    Dim marker = Color.FromArgb(Integer.Parse(CStr(group.Attribute("colour")), CultureInfo.InvariantCulture))
                    Dim colour = SoftColour(marker)
                    For Each entry In group.Elements(Ns + "Date")
                        Dim dateRange = ResolveAnchor(wb, CStr(entry.Attribute("anchor")))
                        Dim target = FundingScheduleWriter.Targets.FirstOrDefault(Function(t) t.DateRange = CStr(entry.Attribute("section")) AndAlso t.ValueRange = CStr(entry.Attribute("values")))
                        If dateRange Is Nothing OrElse dateRange.RowCount <> 1 OrElse dateRange.ColumnCount <> 1 OrElse target Is Nothing Then Continue For
                        Dim pair = FundingScheduleWriter.Ranges(wb, target)
                        If dateRange.Worksheet IsNot pair.Item1.Worksheet OrElse dateRange.LeftColumnIndex <> pair.Item1.LeftColumnIndex OrElse
                            dateRange.TopRowIndex < pair.Item1.TopRowIndex OrElse dateRange.TopRowIndex > pair.Item1.BottomRowIndex OrElse
                            facility.LeftColumnIndex < pair.Item2.LeftColumnIndex OrElse facility.LeftColumnIndex > pair.Item2.RightColumnIndex Then Continue For
                        Dim tint As New FundingScheduleTint With {.DateCell = dateRange(0, 0), .ExpectedDate = DateTime.ParseExact(CStr(entry.Attribute("value")), "yyyy-MM-dd", CultureInfo.InvariantCulture), .Colour = colour, .MarkerColour = marker}
                        result(dateRange.TopRowIndex.ToString() & ":" & dateRange.LeftColumnIndex.ToString()) = tint
                        result(dateRange.TopRowIndex.ToString() & ":" & facility.LeftColumnIndex.ToString()) = tint
                    Next
                Catch ex As Exception When TypeOf ex Is ArgumentException OrElse TypeOf ex Is FormatException OrElse TypeOf ex Is InvalidOperationException
                    'Deleted/invalid anchors are inactive presentation, never remapped
                    'by matching a duplicate date or facility name.
                End Try
            Next
            Return result
        End Function

        Private Shared Function ResolveAnchor(wb As IWorkbook, name As String) As CellRange
            If String.IsNullOrEmpty(name) OrElse Not name.StartsWith("_SummitSchedule_", StringComparison.Ordinal) Then Return Nothing
            Dim defined = wb.DefinedNames.GetDefinedName(name)
            If defined Is Nothing OrElse defined.RefersTo.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0 Then Return Nothing
            Return defined.Range
        End Function
    End Class
End Namespace
