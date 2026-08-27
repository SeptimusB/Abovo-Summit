Imports DevExpress.Spreadsheet
Imports System.IO
Imports System.IO.Compression
Imports System.Xml.Linq

Namespace Abovo

    Public MustInherit Class WorkbookModelProfile

        Public MustOverride ReadOnly Property ModelType As String
        Public MustOverride ReadOnly Property DisplayName As String
        Public MustOverride ReadOnly Property FallbackStructureFileName As String
        Public MustOverride ReadOnly Property UsesTransactionalDatabase As Boolean

        Public MustOverride Function ValidateContract(
            ByVal Workbook As IWorkbook) As String

        Public MustOverride Sub ApplyMetadata(
            ByVal Workbook As IWorkbook,
            ByVal ModelDefinition As Abovo_Model_Def)

        Public Function ResolveStructureSource(ByVal WorkbookPath As String) As String

            Dim EmbeddedDefinition As String =
                EmbeddedWorkbookStructureReader.TryRead(
                    WorkbookPath,
                    ModelType)

            If Not String.IsNullOrWhiteSpace(EmbeddedDefinition) Then
                Return EmbeddedDefinition
            End If

            Return Path.Combine(
                Application.StartupPath,
                FallbackStructureFileName)

        End Function

        Protected Shared Function GetDefinedRange(
            ByVal Workbook As IWorkbook,
            ByVal Name As String) As CellRange

            Dim Definition As DefinedName =
                Workbook.DefinedNames.GetDefinedName(Name)

            If Definition Is Nothing Then Return Nothing
            Return Definition.Range

        End Function

        Protected Shared Function CellDisplayText(
            ByVal Worksheet As Worksheet,
            ByVal Address As String) As String

            If Worksheet Is Nothing Then Return String.Empty
            Return Worksheet.Cells(Address).DisplayText.Trim()

        End Function

    End Class

    Public NotInheritable Class AbovoBPWorkbookProfile
        Inherits WorkbookModelProfile

        Public Overrides ReadOnly Property ModelType As String
            Get
                Return "AbovoBP"
            End Get
        End Property

        Public Overrides ReadOnly Property DisplayName As String
            Get
                Return "HA Business Plan"
            End Get
        End Property

        Public Overrides ReadOnly Property FallbackStructureFileName As String
            Get
                Return "Structure.xml"
            End Get
        End Property

        Public Overrides ReadOnly Property UsesTransactionalDatabase As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides Function ValidateContract(
            ByVal Workbook As IWorkbook) As String

            If Not Workbook.Worksheets.Contains("Global Assumptions") Then
                Return "The workbook is missing the 'Global Assumptions' worksheet."
            End If

            If Not String.Equals(
                Workbook.Worksheets("Global Assumptions").Cells("A8").DisplayText,
                "Business Plan Start Date",
                StringComparison.Ordinal) Then

                Return "The workbook does not contain the expected Abovo " &
                       "validation marker at Global Assumptions!A8."
            End If

            If Not Workbook.Worksheets.Contains("Transactional DB") Then
                Return "The workbook is missing the 'Transactional DB' worksheet."
            End If

            Return String.Empty

        End Function

        Public Overrides Sub ApplyMetadata(
            ByVal Workbook As IWorkbook,
            ByVal ModelDefinition As Abovo_Model_Def)

            Dim GlobalAssumptions As Worksheet =
                Workbook.Worksheets("Global Assumptions")

            ModelDefinition.CompanyName =
                GlobalAssumptions.Cells(5, 2).DisplayText.Trim()

            Dim StartDateValue As CellValue =
                GlobalAssumptions.Cells(7, 2).Value

            If StartDateValue.IsDateTime Then
                ModelDefinition.StartDate =
                    StartDateValue.DateTimeValue.ToString("yyyy-MM-dd")
            Else
                ModelDefinition.StartDate =
                    GlobalAssumptions.Cells(7, 2).DisplayText.Trim()
            End If

        End Sub

    End Class

    Public NotInheritable Class AbovoDSAWorkbookProfile
        Inherits WorkbookModelProfile

        Private Shared ReadOnly RequiredSheets As String() = {
            "Global Assumptions",
            "Key Dates",
            "Unit Mix",
            "Check Sheet",
            "Hidden - Export Sheet",
            "Cashflow Selector"
        }

        Private Shared ReadOnly RequiredNames As String() = {
            "ModelVersion",
            "SchemeName",
            "CheckTotal",
            "BplanYear"
        }

        Public Overrides ReadOnly Property ModelType As String
            Get
                Return "AbovoDSA"
            End Get
        End Property

        Public Overrides ReadOnly Property DisplayName As String
            Get
                Return "Development Scheme Appraisal"
            End Get
        End Property

        Public Overrides ReadOnly Property FallbackStructureFileName As String
            Get
                Return "DSAStructure.xml"
            End Get
        End Property

        Public Overrides ReadOnly Property UsesTransactionalDatabase As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides Function ValidateContract(
            ByVal Workbook As IWorkbook) As String

            For Each SheetName As String In RequiredSheets
                If Not Workbook.Worksheets.Contains(SheetName) Then
                    Return "The DSA workbook is missing the '" &
                           SheetName & "' worksheet."
                End If
            Next

            For Each DefinedNameText As String In RequiredNames
                If GetDefinedRange(Workbook, DefinedNameText) Is Nothing Then
                    Return "The DSA workbook is missing the '" &
                           DefinedNameText & "' defined range."
                End If
            Next

            Dim VersionText As String =
                GetDefinedRange(Workbook, "ModelVersion")(0, 0).
                    DisplayText.Trim()

            Dim NormalisedVersion As Decimal
            If Not Decimal.TryParse(
                VersionText,
                Globalization.NumberStyles.Any,
                Globalization.CultureInfo.InvariantCulture,
                NormalisedVersion) Then

                Decimal.TryParse(VersionText, NormalisedVersion)
            End If

            If NormalisedVersion <> 6.15D Then
                Return "This Summit build currently supports DSA model " &
                       "version 6.1500. The selected workbook reports version " &
                       If(String.IsNullOrWhiteSpace(VersionText), "<blank>", VersionText) & "."
            End If

            Return String.Empty

        End Function

        Public Overrides Sub ApplyMetadata(
            ByVal Workbook As IWorkbook,
            ByVal ModelDefinition As Abovo_Model_Def)

            Dim GlobalAssumptions As Worksheet =
                Workbook.Worksheets("Global Assumptions")

            ModelDefinition.CompanyName = CellDisplayText(GlobalAssumptions, "C7")

            If String.IsNullOrWhiteSpace(ModelDefinition.CompanyName) Then
                ModelDefinition.CompanyName = "Development Scheme Appraisal"
            End If

            Dim StartDateRange As CellRange =
                GetDefinedRange(Workbook, "SchStDate")

            If StartDateRange IsNot Nothing AndAlso
               StartDateRange(0, 0).Value.IsDateTime Then

                ModelDefinition.StartDate =
                    StartDateRange(0, 0).Value.DateTimeValue.ToString("yyyy-MM-dd")
            ElseIf StartDateRange IsNot Nothing Then
                ModelDefinition.StartDate = StartDateRange(0, 0).DisplayText.Trim()
            Else
                ModelDefinition.StartDate = CellDisplayText(GlobalAssumptions, "C16")
            End If

        End Sub

    End Class

    Public NotInheritable Class WorkbookModelProfileRegistry

        Private Sub New()
        End Sub

        Public Shared Function Resolve(
            ByVal Workbook As IWorkbook,
            ByRef FailureMessage As String) As WorkbookModelProfile

            If Workbook Is Nothing Then
                FailureMessage = "The workbook could not be loaded."
                Return Nothing
            End If

            Dim Profiles As WorkbookModelProfile() = {
                New AbovoBPWorkbookProfile(),
                New AbovoDSAWorkbookProfile()
            }

            Dim ValidationMessages As New List(Of String)

            For Each Profile As WorkbookModelProfile In Profiles
                Dim ValidationMessage As String =
                    Profile.ValidateContract(Workbook)

                If String.IsNullOrWhiteSpace(ValidationMessage) Then
                    FailureMessage = String.Empty
                    Return Profile
                End If

                ValidationMessages.Add(
                    Profile.DisplayName & ": " & ValidationMessage)
            Next

            FailureMessage =
                "The workbook is not a supported Abovo model." &
                Environment.NewLine &
                String.Join(Environment.NewLine, ValidationMessages)

            Return Nothing

        End Function

    End Class

    Public NotInheritable Class EmbeddedWorkbookStructureReader

        Private Sub New()
        End Sub

        Public Shared Function TryRead(
            ByVal WorkbookPath As String,
            ByVal ExpectedModelType As String) As String

            If String.IsNullOrWhiteSpace(WorkbookPath) OrElse
               Not File.Exists(WorkbookPath) Then Return Nothing

            Try
                Using WorkbookStream As New FileStream(
                    WorkbookPath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite)

                    Using Package As New ZipArchive(
                        WorkbookStream,
                        ZipArchiveMode.Read,
                        leaveOpen:=False)

                        For Each Entry As ZipArchiveEntry In Package.Entries
                            If Not Entry.FullName.StartsWith(
                                "customXml/",
                                StringComparison.OrdinalIgnoreCase) OrElse
                               Not Entry.FullName.EndsWith(
                                ".xml",
                                StringComparison.OrdinalIgnoreCase) Then Continue For

                            Using EntryStream As Stream = Entry.Open()
                                Dim Definition As XDocument =
                                    XDocument.Load(EntryStream, LoadOptions.PreserveWhitespace)

                                If Definition.Root Is Nothing OrElse
                                   Not String.Equals(
                                    Definition.Root.Name.LocalName,
                                    "Abovo_Model_Def",
                                    StringComparison.OrdinalIgnoreCase) Then Continue For

                                Dim ModelTypeElement As XElement =
                                    Definition.Root.Elements().FirstOrDefault(
                                        Function(Element) String.Equals(
                                            Element.Name.LocalName,
                                            "ModelType",
                                            StringComparison.OrdinalIgnoreCase))

                                If ModelTypeElement IsNot Nothing AndAlso
                                   String.Equals(
                                    ModelTypeElement.Value.Trim(),
                                    ExpectedModelType,
                                    StringComparison.OrdinalIgnoreCase) Then

                                    Return Definition.ToString(SaveOptions.DisableFormatting)
                                End If
                            End Using
                        Next
                    End Using
                End Using
            Catch ex As InvalidDataException
                'A missing or unreadable custom XML part must not prevent the
                'packaged structure fallback from being used.
            Catch ex As IOException
            Catch ex As UnauthorizedAccessException
            End Try

            Return Nothing

        End Function

    End Class

End Namespace
