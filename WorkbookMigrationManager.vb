Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Globalization
Imports System.Security.Cryptography
Imports System.Text
Imports Abovo.FileManager
Imports Abovo.AbovoAppCls
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas
Imports DevExpress.Spreadsheet.Functions

Namespace Abovo

    Public Class WorkbookMigrationManager

        'Schema 2 establishes the standing workbook contract:
        '
        '  * the persisted XLSB must remain fully functional in Microsoft Excel;
        '  * Summit may store metadata/templates in very-hidden worksheets;
        '  * live Transactional DB formula blocks remain formula-backed in the XLSB;
        '  * value-only materialisation must never become the persisted canonical state.
        '
        'Schema 1 was an experimental materialiser migration.  Schema 2 repairs
        'that experimental state, if present, by recreating the live formulas
        'from the cached R1C1 templates.
        Public Const CurrentSchemaVersion As Integer = 10

        'Retained for source compatibility with the experimental implementation.
        'Capacity reservation is no longer used in Excel-compatible mode.
        Public Const DevelopmentIdentifiedCapacity As Integer = 50

        Private Const MetadataSheetName As String = "__Abovo_Metadata"
        Private Const TemplateSheetName As String = "__Abovo_Templates"
        Private Const DevelopmentShadowSheetName As String = "__Abovo_TransDB_Development"
        Private Const DevelopmentShadowCapability As String = "DevelopmentIdentifiedShadowComparison"
        Private Const DevelopmentShadowCapacity As Integer = 50
        Private Const DevelopmentShadowBlockStride As Integer = 53
        Private Const SchemaMarker As String = "AbovoSummitWorkbookSchema"
        Private Const DevelopmentIdentifiedCapability As String =
            "DevelopmentIdentifiedFormulaTemplates"
        Private Const FormulaTemplatePrefix As String = "ABOVO_R1C1_FORMULA|"

        Private ReadOnly ModelID As Integer

        Private Shared ReadOnly DevelopmentSingleNames() As String = {
            "TransCopy_DevptSingle_A", "TransCopy_DevptSingle_B", "TransCopy_DevptSingle_C",
            "TransCopy_DevptSingle_D", "TransCopy_DevptSingle_E", "TransCopy_DevptSingle_F",
            "TransCopy_DevptSingle_G", "TransCopy_DevptSingle_H", "TransCopy_DevptSingle_I",
            "TransCopy_DevptSingle_J", "TransCopy_DevptSingle_K", "TransCopy_DevptSingle_L",
            "TransCopy_DevptSingle_M", "TransCopy_DevptSingle_N"
        }

        Public Sub New(ByVal SetModelID As Integer)

            ModelID = SetModelID

        End Sub

        Private Function GetWorkbook() As IWorkbook

            If ExcelModels Is Nothing Then Return Nothing
            If ModelID < 0 OrElse ModelID >= ExcelModels.Length Then Return Nothing
            If ExcelModels(ModelID) Is Nothing Then Return Nothing

            Return ExcelModels(ModelID).WB

        End Function

        Public Function GetSchemaVersion() As Integer

            Dim WB As IWorkbook = GetWorkbook()
            If WB Is Nothing Then Return 0

            Dim WS As Worksheet = TryGetWorksheet(WB, MetadataSheetName)

            If WS Is Nothing Then Return 0
            If Not WS.Cells(0, 0).Value.IsText Then Return 0

            If Not String.Equals(WS.Cells(0, 0).Value.TextValue,
                                 SchemaMarker,
                                 StringComparison.Ordinal) Then Return 0

            Dim VersionCell As CellValue = WS.Cells(0, 1).Value

            If VersionCell.IsNumeric Then
                Return CInt(VersionCell.NumericValue)
            End If

            If VersionCell.IsText Then

                Dim Parsed As Integer

                If Integer.TryParse(VersionCell.TextValue,
                                    NumberStyles.Integer,
                                    CultureInfo.InvariantCulture,
                                    Parsed) Then
                    Return Parsed
                End If

            End If

            Return 0

        End Function

        'This method name is retained because existing Summit code already calls
        'it.  "Installed" now means the workbook contains the safe template/
        'metadata capability; it does NOT mean value-only runtime mode is active.
        Public Function IsDevelopmentIdentifiedMaterialiserInstalled() As Boolean

            Dim WB As IWorkbook = GetWorkbook()
            If WB Is Nothing Then Return False
            If GetSchemaVersion() < 1 Then Return False

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(WB, MetadataSheetName)

            Dim TemplateWS As Worksheet =
                TryGetWorksheet(WB, TemplateSheetName)

            If MetadataWS Is Nothing OrElse
               TemplateWS Is Nothing Then Return False

            If MetadataWS.Cells(1, 0).Value.IsText Then

                Dim Capability As String =
                    MetadataWS.Cells(1, 0).Value.TextValue

                If Not String.Equals(Capability,
                                     DevelopmentIdentifiedCapability,
                                     StringComparison.Ordinal) AndAlso
                   Not String.Equals(Capability,
                                     "DevelopmentIdentifiedMaterialiser",
                                     StringComparison.Ordinal) Then

                    Return False

                End If

            End If

            For BlockIndex As Integer =
                0 To DevelopmentSingleNames.Length - 1

                Dim BaseRow As Integer =
                    2 + (BlockIndex * 5)

                If Not TemplateWS.Cells(BaseRow, 0).Value.IsText OrElse
                   Not String.Equals(
                       TemplateWS.Cells(BaseRow, 0).Value.TextValue,
                       DevelopmentSingleNames(BlockIndex),
                       StringComparison.OrdinalIgnoreCase) Then

                    Return False

                End If

            Next

            Return True

        End Function

        Public Function IsExcelStandaloneCompatibleSchema() As Boolean

            Return GetSchemaVersion() >= 2

        End Function

        Public Function ApplyPendingMigrations() As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            Dim Version As Integer = GetSchemaVersion()

            If Version > CurrentSchemaVersion Then
                Result.BError = True
                Result.StringReturn =
                    "Workbook schema " & Version.ToString &
                    " is newer than Summit schema " &
                    CurrentSchemaVersion.ToString & "."
                Return Result
            End If

            If Version = 3 OrElse Version = 4 Then
                Result.BError = True
                Result.StringReturn =
                    "This workbook contains abandoned experimental TransactionDB schema " &
                    Version.ToString & ". Use an original/schema-0 or safe schema-2/5/6/7/8/9 workbook."
                Return Result
            End If

            If Version = 0 Then
                Dim M1 As AbovoTransaction =
                    ApplyMigration001CaptureDevelopmentIdentifiedTemplates()
                If M1.BError Then Return M1
                Version = GetSchemaVersion()
            End If

            If Version = 1 Then
                Dim M2 As AbovoTransaction =
                    ApplyMigration002RestoreExcelFormulaState()
                If M2.BError Then Return M2
                Version = GetSchemaVersion()
            End If

            If Version = 2 OrElse
               Version = 5 OrElse
               Version = 6 OrElse
               Version = 7 OrElse
               Version = 8 OrElse
               Version = 9 Then

                Dim M10 As AbovoTransaction =
                    ApplyMigration010ActivateDevelopmentProductionMirrors()

                If M10.BError Then Return M10
                Version = GetSchemaVersion()
            End If

            If Version = 10 Then

                Dim RoundTripCheck As AbovoTransaction =
                    ReconcileExternalDevelopmentTransactionalDBEdits()

                If RoundTripCheck.BError Then
                    Return RoundTripCheck
                End If

            End If

            If Version <> CurrentSchemaVersion Then
                Result.BError = True
                Result.StringReturn =
                    "Workbook migration stopped at schema " &
                    Version.ToString &
                    "; Summit expected schema " &
                    CurrentSchemaVersion.ToString & "."
                Return Result
            End If

            Result.StringReturn =
                "Workbook schema is current at version " &
                Version.ToString &
                ". Development Identified production mirrors are active."

            Return Result

        End Function

        '=====================================================================
        ' MIGRATION 001
        '
        'SAFE FOR STANDALONE EXCEL.
        '
        'Captures the workbook's own Development Identified TransactionDB
        'formula/constant templates in a very-hidden worksheet.  It does NOT:
        '   * insert rows;
        '   * create reserved capacity;
        '   * replace live formulas with values;
        '   * alter the TransCopy_* active ranges.
        '
        'A later/newer XLSB is therefore migrated using its own current formula
        'templates rather than formulas hard-coded into Summit.
        '=====================================================================
        Private Function ApplyMigration001CaptureDevelopmentIdentifiedTemplates() As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            Dim TransactionWS As Worksheet =
                TryGetWorksheet(WB, "Transactional DB")

            If TransactionWS Is Nothing Then
                Result.BError = True
                Result.StringReturn =
                    "Migration 001 precondition failed: Transactional DB worksheet was not found."
                Return Result
            End If

            Dim ValidationResult As AbovoTransaction =
                ValidateDevelopmentSingleRanges(WB, TransactionWS)

            If ValidationResult.BError Then Return ValidationResult

            Dim PreviousCalculationMode As WorkbookCalculationMode =
                WB.Options.CalculationMode

            Dim PreviousEngine As CalculationEngineType =
                WB.Options.CalculationEngineType

            Dim PreviousHistoryEnabled As Boolean =
                WB.History.IsEnabled

            Dim EngineChanged As Boolean = False
            Dim HistoryChanged As Boolean = False
            Dim UpdateStarted As Boolean = False
            Dim MigrationTimer As Stopwatch = Stopwatch.StartNew()

            Try


                If WB.History.IsEnabled Then
                    WB.History.IsEnabled = False
                    HistoryChanged = True
                End If

                If WB.Options.CalculationEngineType <> CalculationEngineType.Recursive Then
                    WB.Options.CalculationEngineType = CalculationEngineType.Recursive
                    EngineChanged = True
                End If

                WB.Options.CalculationMode = WorkbookCalculationMode.Manual

                WB.BeginUpdate()
                UpdateStarted = True

                Dim MetadataWS As Worksheet =
                    GetOrCreateVeryHiddenWorksheet(WB, MetadataSheetName)

                Dim TemplateWS As Worksheet =
                    GetOrCreateVeryHiddenWorksheet(WB, TemplateSheetName)

                MetadataWS.Cells(0, 0).Value = SchemaMarker
                MetadataWS.Cells(0, 1).Value = 0

                MetadataWS.Cells(1, 0).Value =
                    DevelopmentIdentifiedCapability

                MetadataWS.Cells(1, 1).Value =
                    "TEMPLATES_INSTALLED"

                MetadataWS.Cells(2, 0).Value =
                    "ExcelStandaloneCompatibility"

                MetadataWS.Cells(2, 1).Value =
                    "FORMULA_STATE_REQUIRED"

                MetadataWS.Cells(3, 0).Value =
                    "PersistValueOnlyTransactionDB"

                MetadataWS.Cells(3, 1).Value =
                    "FALSE"

                CaptureDevelopmentIdentifiedTemplates(WB,
                                                      TransactionWS,
                                                      TemplateWS)

                MetadataWS.Cells(0, 1).Value = 1

                MetadataWS.Cells(4, 0).Value =
                    "Migration001AppliedUtc"

                MetadataWS.Cells(4, 1).Value =
                    DateTime.UtcNow.ToString("o",
                                             CultureInfo.InvariantCulture)

                ExcelModels(ModelID).IsDirty = True

                Result.StringReturn =
                    "Migration 001 captured Development Identified formula templates without changing live TransactionDB formulas."



            Catch ex As Exception

                Result.BError = True
                Result.StringReturn =
                    "Migration 001 failed: " & ex.Message


            Finally

                If UpdateStarted Then

                    Try
                        WB.EndUpdate()
                    Catch
                    End Try

                End If

                Try
                    WB.Options.CalculationMode =
                        PreviousCalculationMode
                Catch
                End Try

                If EngineChanged Then

                    Try
                        WB.Options.CalculationEngineType =
                            PreviousEngine
                    Catch
                    End Try

                End If

                If HistoryChanged Then

                    Try
                        WB.History.IsEnabled =
                            PreviousHistoryEnabled
                    Catch
                    End Try

                End If

            End Try

            Return Result

        End Function

        '=====================================================================
        ' MIGRATION 002
        '
        'Repairs the experimental schema-1 migration if it was previously
        'applied/saved.  Formula templates captured in __Abovo_Templates are
        'used to recreate every active Development Identified TransactionDB
        'data-row formula and footer formula.
        '
        'For a schema-1 workbook produced by the revised safe Migration 001,
        'this is effectively an idempotent formula rewrite.
        '
        'Experimental capacity gaps, if present, are left physically in place.
        'They are harmless to the Excel formula model and avoiding structural
        'deletion here prevents another expensive/risky dependency rewrite.
        'The old AbovoCap_* names are no longer used by Summit.
        '=====================================================================
        Private Function ApplyMigration002RestoreExcelFormulaState() As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(WB, MetadataSheetName)

            Dim TemplateWS As Worksheet =
                TryGetWorksheet(WB, TemplateSheetName)

            If MetadataWS Is Nothing OrElse
               TemplateWS Is Nothing Then

                Result.BError = True
                Result.StringReturn =
                    "Migration 002 requires the metadata/template worksheets created by Migration 001."

                Return Result

            End If

            Dim IsExperimentalValueOnlyState As Boolean =
                DetectExperimentalSchema1State(WB, MetadataWS)


            If IsExperimentalValueOnlyState Then

                If ExcelModels(ModelID).TransDBMaterialiser Is Nothing Then

                    Result.BError = True
                    Result.StringReturn =
                        "Migration 002 cannot repair experimental schema 1 because TransactionDBMaterialiser is unavailable."

                    Return Result

                End If

                Dim RestoreResult As AbovoTransaction =
                    ExcelModels(ModelID).TransDBMaterialiser.
                        RestoreDevelopmentIdentifiedFormulaState(True)

                If RestoreResult.BError Then Return RestoreResult


            Else

                'Safe Migration 001 never altered the live TransactionDB formulas.
                'Do NOT rewrite thousands of cells unnecessarily.  Re-capture the
                'templates with schema-2 encoded storage so a leading apostrophe
                'in a quoted sheet name can never be consumed as an Excel text
                'prefix character.
                Dim TransactionWS As Worksheet =
                    TryGetWorksheet(WB, "Transactional DB")

                If TransactionWS Is Nothing Then

                    Result.BError = True
                    Result.StringReturn =
                        "Migration 002 could not recapture templates because Transactional DB was not found."

                    Return Result

                End If

                CaptureDevelopmentIdentifiedTemplates(WB,
                                                      TransactionWS,
                                                      TemplateWS)


            End If

            MetadataWS.Cells(0, 0).Value = SchemaMarker
            MetadataWS.Cells(0, 1).Value = 2

            MetadataWS.Cells(1, 0).Value =
                DevelopmentIdentifiedCapability

            MetadataWS.Cells(1, 1).Value =
                "TEMPLATES_INSTALLED"

            MetadataWS.Cells(2, 0).Value =
                "ExcelStandaloneCompatibility"

            MetadataWS.Cells(2, 1).Value =
                "FORMULA_STATE_REQUIRED"

            MetadataWS.Cells(3, 0).Value =
                "PersistValueOnlyTransactionDB"

            MetadataWS.Cells(3, 1).Value =
                "FALSE"

            MetadataWS.Cells(5, 0).Value =
                "Migration002AppliedUtc"

            MetadataWS.Cells(5, 1).Value =
                DateTime.UtcNow.ToString("o",
                                         CultureInfo.InvariantCulture)

            ExcelModels(ModelID).IsDirty = True


            Result.StringReturn =
                "Migration 002 established the Excel-compatible workbook contract."

            Return Result

        End Function

        Private Function DetectExperimentalSchema1State(ByVal WB As IWorkbook,
                                                        ByVal MetadataWS As Worksheet) As Boolean

            If MetadataWS Is Nothing Then Return False

            Dim Capability As String = String.Empty
            Dim Status As String = String.Empty

            If MetadataWS.Cells(1, 0).Value.IsText Then
                Capability = MetadataWS.Cells(1, 0).Value.TextValue
            End If

            If MetadataWS.Cells(1, 1).Value.IsText Then
                Status = MetadataWS.Cells(1, 1).Value.TextValue
            End If

            If String.Equals(Capability,
                             "DevelopmentIdentifiedMaterialiser",
                             StringComparison.Ordinal) AndAlso
               String.Equals(Status,
                             "INSTALLED",
                             StringComparison.OrdinalIgnoreCase) Then

                Return True

            End If

            'Secondary detection for a partially/older experimental workbook:
            'the old migration created AbovoCap_TransCopy_DevptSingle_* names.
            For Each DN As DefinedName In WB.DefinedNames

                If DN IsNot Nothing AndAlso
                   DN.Name IsNot Nothing AndAlso
                   DN.Name.StartsWith("AbovoCap_TransCopy_DevptSingle_",
                                      StringComparison.OrdinalIgnoreCase) Then

                    Return True

                End If

            Next

            Return False

        End Function

        '=====================================================================
        ' MIGRATION 005
        '
        'DEVELOPMENT IDENTIFIED SHADOW COMPARISON.
        '
        'This is the replacement for the abandoned in-place schema 3/4
        'experiment.  It makes NO structural change to Transactional DB and
        'does NOT repoint TransCopy_DevptSingle_A:N.
        '
        'A very-hidden Excel-native worksheet is created:
        '
        '   __Abovo_TransDB_Development
        '
        'For each current Development Identified A:N mirror:
        '   * every active data row is reconstructed cell-by-cell;
        '   * formulas are converted Source A1 -> Source R1C1 -> Shadow A1;
        '   * constants are copied exactly from the corresponding source row;
        '   * the current footer is reconstructed likewise;
        '   * an additional canonical footer is stored at the end of a
        '     50-record physical block;
        '   * unused future rows are deliberately left empty for now.
        '
        'Every reconstructed active cell is then compared with the original
        'Transactional DB cell using independent FormulaEngine evaluation and
        'normalised R1C1 formula signatures.
        '
        'The original TransCopy_* names remain authoritative throughout.
        '=====================================================================
        Private Function ApplyMigration005CreateDevelopmentShadowComparison() As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            If GetSchemaVersion() <> 2 Then
                Result.BError = True
                Result.StringReturn =
                    "Migration 005 requires safe workbook schema 2."
                Return Result
            End If

            Dim SourceWS As Worksheet =
                TryGetWorksheet(WB, "Transactional DB")

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(WB, MetadataSheetName)

            If SourceWS Is Nothing OrElse MetadataWS Is Nothing Then
                Result.BError = True
                Result.StringReturn =
                    "Migration 005 requires Transactional DB and __Abovo_Metadata."
                Return Result
            End If

            Dim Validation As AbovoTransaction =
                ValidateDevelopmentSingleRanges(WB, SourceWS)

            If Validation.BError Then Return Validation

            Dim PreviousCalculationMode As WorkbookCalculationMode =
                WB.Options.CalculationMode

            Dim PreviousEngine As CalculationEngineType =
                WB.Options.CalculationEngineType

            Dim PreviousHistoryEnabled As Boolean =
                WB.History.IsEnabled

            Dim EngineChanged As Boolean = False
            Dim HistoryChanged As Boolean = False
            Dim UpdateStarted As Boolean = False
            Dim MigrationTimer As Stopwatch = Stopwatch.StartNew()

            Try


                If WB.History.IsEnabled Then
                    WB.History.IsEnabled = False
                    HistoryChanged = True
                End If

                If WB.Options.CalculationEngineType <>
                   CalculationEngineType.Recursive Then

                    WB.Options.CalculationEngineType =
                        CalculationEngineType.Recursive

                    EngineChanged = True

                End If

                WB.Options.CalculationMode =
                    WorkbookCalculationMode.Manual

                WB.BeginUpdate()
                UpdateStarted = True

                Dim ShadowWS As Worksheet =
                    TryGetWorksheet(WB, DevelopmentShadowSheetName)

                If ShadowWS IsNot Nothing Then
                    WB.Worksheets.Remove(ShadowWS)
                End If

                ShadowWS =
                    WB.Worksheets.Add(DevelopmentShadowSheetName)

                ShadowWS.VisibilityType =
                    WorksheetVisibilityType.VeryHidden

                Dim Engine As FormulaEngine =
                    WB.FormulaEngine

                Dim Culture As CultureInfo =
                    CultureInfo.GetCultureInfo("en-US")

                Dim TotalCells As Integer = 0
                Dim FormulaCells As Integer = 0
                Dim ConstantCells As Integer = 0
                Dim ValueMismatches As Integer = 0
                Dim FormulaMismatches As Integer = 0
                Dim EvaluationErrors As Integer = 0
                Dim ExamplesPrinted As Integer = 0
                Const MaximumExamples As Integer = 100

                For BlockIndex As Integer =
                    0 To DevelopmentSingleNames.Length - 1

                    Dim TargetName As String =
                        DevelopmentSingleNames(BlockIndex)

                    Dim SourceDN As DefinedName =
                        TryGetDefinedName(WB, TargetName)

                    If SourceDN Is Nothing OrElse
                       SourceDN.Range Is Nothing Then

                        Throw New InvalidOperationException(
                            "Migration 005 could not resolve " &
                            TargetName & ".")

                    End If

                    Dim SourceRange As CellRange =
                        SourceDN.Range

                    Dim CurrentDataRows As Integer =
                        SourceRange.RowCount - 1

                    If CurrentDataRows <= 0 Then

                        Throw New InvalidOperationException(
                            TargetName &
                            " has no data rows.")

                    End If

                    If CurrentDataRows >
                       DevelopmentShadowCapacity Then

                        Throw New InvalidOperationException(
                            TargetName &
                            " has " &
                            CurrentDataRows.ToString &
                            " data rows, exceeding shadow capacity " &
                            DevelopmentShadowCapacity.ToString & ".")

                    End If

                    Dim ShadowTop As Integer =
                        BlockIndex *
                        DevelopmentShadowBlockStride

                    Dim ShadowFooterRow As Integer =
                        ShadowTop +
                        CurrentDataRows

                    Dim CanonicalFooterRow As Integer =
                        ShadowTop +
                        DevelopmentShadowCapacity

                    ShadowWS.Cells(ShadowTop, 75).Value =
                        TargetName

                    ShadowWS.Cells(ShadowTop, 76).Value =
                        "ACTIVE_ROWS=" &
                        CurrentDataRows.ToString

                    Dim BlockTimer As Stopwatch =
                        Stopwatch.StartNew()

                    'Reconstruct the current active data rows + current footer.
                    For RowOffset As Integer =
                        0 To SourceRange.RowCount - 1

                        Dim SourceRow As Integer =
                            SourceRange.TopRowIndex +
                            RowOffset

                        Dim ShadowRow As Integer =
                            ShadowTop +
                            RowOffset

                        For ColumnOffset As Integer =
                            0 To SourceRange.ColumnCount - 1

                            Dim SourceColumn As Integer =
                                SourceRange.LeftColumnIndex +
                                ColumnOffset

                            Dim ShadowColumn As Integer =
                                ColumnOffset

                            Dim SourceCell As Cell =
                                SourceWS.Cells(SourceRow,
                                               SourceColumn)

                            Dim ShadowCell As Cell =
                                ShadowWS.Cells(ShadowRow,
                                               ShadowColumn)

                            ReconstructCellInShadow(
                                Engine,
                                Culture,
                                SourceWS,
                                SourceCell,
                                ShadowWS,
                                ShadowCell)

                            TotalCells += 1

                            If String.IsNullOrWhiteSpace(SourceCell.Formula) Then
                                ConstantCells += 1
                            Else
                                FormulaCells += 1
                            End If

                            Dim ValueReason As String = Nothing

                            Try

                                If Not ShadowCellMatchesSource(
                                    Engine,
                                    Culture,
                                    SourceWS,
                                    SourceCell,
                                    ShadowWS,
                                    ShadowCell,
                                    ValueReason) Then

                                    ValueMismatches += 1

                                    If ExamplesPrinted <
                                       MaximumExamples Then


                                        ExamplesPrinted += 1

                                    End If

                                End If

                            Catch ex As Exception

                                EvaluationErrors += 1

                                If ExamplesPrinted <
                                   MaximumExamples Then


                                    ExamplesPrinted += 1

                                End If

                            End Try

                            If Not String.IsNullOrWhiteSpace(
                                SourceCell.Formula) Then

                                Dim SourceSignature As String =
                                    GetR1C1Signature(
                                        Engine,
                                        Culture,
                                        SourceWS,
                                        SourceCell)

                                Dim ShadowSignature As String =
                                    GetR1C1Signature(
                                        Engine,
                                        Culture,
                                        ShadowWS,
                                        ShadowCell)

                                If Not String.Equals(
                                    SourceSignature,
                                    ShadowSignature,
                                    StringComparison.Ordinal) Then

                                    FormulaMismatches += 1

                                    If ExamplesPrinted <
                                       MaximumExamples Then


                                        ExamplesPrinted += 1

                                    End If

                                End If

                            End If

                        Next

                    Next

                    'Store a canonical footer at the fixed end of the 50-record
                    'physical block.  This is not part of parity comparison and
                    'does not change any TransCopy_* name.
                    Dim SourceFooterRow As Integer =
                        SourceRange.BottomRowIndex

                    For ColumnOffset As Integer =
                        0 To SourceRange.ColumnCount - 1

                        Dim SourceColumn As Integer =
                            SourceRange.LeftColumnIndex +
                            ColumnOffset

                        ReconstructCellInShadow(
                            Engine,
                            Culture,
                            SourceWS,
                            SourceWS.Cells(SourceFooterRow,
                                           SourceColumn),
                            ShadowWS,
                            ShadowWS.Cells(CanonicalFooterRow,
                                           ColumnOffset))

                    Next


                Next

                MetadataWS.Cells(11, 0).Value =
                    "DevelopmentIdentifiedShadowComparison"

                MetadataWS.Cells(11, 1).Value =
                    "INSTALLED"

                MetadataWS.Cells(12, 0).Value =
                    "DevelopmentShadowSheet"

                MetadataWS.Cells(12, 1).Value =
                    DevelopmentShadowSheetName

                MetadataWS.Cells(13, 0).Value =
                    "DevelopmentShadowCapacity"

                MetadataWS.Cells(13, 1).Value =
                    DevelopmentShadowCapacity

                MetadataWS.Cells(14, 0).Value =
                    "DevelopmentShadowValueMismatches"

                MetadataWS.Cells(14, 1).Value =
                    ValueMismatches

                MetadataWS.Cells(15, 0).Value =
                    "DevelopmentShadowFormulaMismatches"

                MetadataWS.Cells(15, 1).Value =
                    FormulaMismatches

                MetadataWS.Cells(16, 0).Value =
                    "DevelopmentShadowEvaluationErrors"

                MetadataWS.Cells(16, 1).Value =
                    EvaluationErrors

                MetadataWS.Cells(17, 0).Value =
                    "Migration005AppliedUtc"

                MetadataWS.Cells(17, 1).Value =
                    DateTime.UtcNow.ToString(
                        "o",
                        CultureInfo.InvariantCulture)

                If ValueMismatches = 0 AndAlso
                   FormulaMismatches = 0 AndAlso
                   EvaluationErrors = 0 Then

                    MetadataWS.Cells(11, 1).Value =
                        "PARITY_PASS"

                Else

                    MetadataWS.Cells(11, 1).Value =
                        "PARITY_ANALYSE"

                End If

                MetadataWS.Cells(0, 1).Value = 5

                ExcelModels(ModelID).IsDirty = True


                If ValueMismatches = 0 AndAlso
                   FormulaMismatches = 0 AndAlso
                   EvaluationErrors = 0 Then


                    Result.StringReturn =
                        "Migration 005 created the Development shadow and " &
                        "achieved exact active-range parity. Original TransCopy_* " &
                        "names remain unchanged."

                Else


                    Result.StringReturn =
                        "Migration 005 created the Development shadow but parity " &
                        "requires analysis. Original TransCopy_* names remain unchanged."

                End If


            Catch ex As Exception

                Result.BError = True
                Result.StringReturn =
                    "Migration 005 failed: " &
                    ex.Message


            Finally

                If UpdateStarted Then

                    Try
                        WB.EndUpdate()
                    Catch
                    End Try

                End If

                Try
                    WB.Options.CalculationMode =
                        PreviousCalculationMode
                Catch
                End Try

                If EngineChanged Then

                    Try
                        WB.Options.CalculationEngineType =
                            PreviousEngine
                    Catch
                    End Try

                End If

                If HistoryChanged Then

                    Try
                        WB.History.IsEnabled =
                            PreviousHistoryEnabled
                    Catch
                    End Try

                End If

            End Try

            Return Result

        End Function

        Private Sub ReconstructCellInShadow(ByVal Engine As FormulaEngine,
                                            ByVal Culture As CultureInfo,
                                            ByVal SourceWS As Worksheet,
                                            ByVal SourceCell As Cell,
                                            ByVal ShadowWS As Worksheet,
                                            ByVal ShadowCell As Cell)

            If String.IsNullOrWhiteSpace(SourceCell.Formula) Then

                ShadowCell.Value =
                    SourceCell.Value

                Return

            End If

            Dim SourceA1Context As New ExpressionContext(
                SourceCell.ColumnIndex,
                SourceCell.RowIndex,
                SourceWS,
                Culture,
                ReferenceStyle.A1,
                ExpressionStyle.Normal)

            Dim SourceR1C1Context As New ExpressionContext(
                SourceCell.ColumnIndex,
                SourceCell.RowIndex,
                SourceWS,
                Culture,
                ReferenceStyle.R1C1,
                ExpressionStyle.Normal)

            Dim ParsedSource As ParsedExpression =
                Engine.Parse(
                    SourceCell.Formula,
                    SourceA1Context)

            Dim R1C1Formula As String =
                ParsedSource.ToString(
                    SourceR1C1Context)

            Dim ShadowR1C1Context As New ExpressionContext(
                ShadowCell.ColumnIndex,
                ShadowCell.RowIndex,
                ShadowWS,
                Culture,
                ReferenceStyle.R1C1,
                ExpressionStyle.Normal)

            Dim ShadowA1Context As New ExpressionContext(
                ShadowCell.ColumnIndex,
                ShadowCell.RowIndex,
                ShadowWS,
                Culture,
                ReferenceStyle.A1,
                ExpressionStyle.Normal)

            Dim ParsedShadow As ParsedExpression =
                Engine.Parse(
                    R1C1Formula,
                    ShadowR1C1Context)

            Dim ShadowFormula As String =
                ParsedShadow.ToString(
                    ShadowA1Context)

            If Not ShadowFormula.StartsWith(
                "=",
                StringComparison.Ordinal) Then

                ShadowFormula =
                    "=" & ShadowFormula

            End If

            ShadowCell.Formula =
                ShadowFormula

        End Sub

        Private Function ShadowCellMatchesSource(ByVal Engine As FormulaEngine,
                                                 ByVal Culture As CultureInfo,
                                                 ByVal SourceWS As Worksheet,
                                                 ByVal SourceCell As Cell,
                                                 ByVal ShadowWS As Worksheet,
                                                 ByVal ShadowCell As Cell,
                                                 ByRef Reason As String) As Boolean

            Reason = Nothing

            If String.IsNullOrWhiteSpace(SourceCell.Formula) Then

                Return CellValuesEquivalent(
                    SourceCell.Value,
                    ShadowCell.Value,
                    Reason)

            End If

            Dim SourceContext As New ExpressionContext(
                SourceCell.ColumnIndex,
                SourceCell.RowIndex,
                SourceWS,
                Culture,
                ReferenceStyle.A1,
                ExpressionStyle.Normal)

            Dim ShadowContext As New ExpressionContext(
                ShadowCell.ColumnIndex,
                ShadowCell.RowIndex,
                ShadowWS,
                Culture,
                ReferenceStyle.A1,
                ExpressionStyle.Normal)

            Dim SourceValue As ParameterValue =
                Engine.Evaluate(
                    SourceCell.Formula,
                    SourceContext)

            Dim ShadowValue As ParameterValue =
                Engine.Evaluate(
                    ShadowCell.Formula,
                    ShadowContext)

            Return ParameterValuesEquivalent(
                SourceValue,
                ShadowValue,
                Reason)

        End Function

        Private Function GetR1C1Signature(ByVal Engine As FormulaEngine,
                                          ByVal Culture As CultureInfo,
                                          ByVal WS As Worksheet,
                                          ByVal Cell As Cell) As String

            If Cell Is Nothing OrElse
               String.IsNullOrWhiteSpace(Cell.Formula) Then

                Return String.Empty

            End If

            Dim A1Context As New ExpressionContext(
                Cell.ColumnIndex,
                Cell.RowIndex,
                WS,
                Culture,
                ReferenceStyle.A1,
                ExpressionStyle.Normal)

            Dim R1C1Context As New ExpressionContext(
                Cell.ColumnIndex,
                Cell.RowIndex,
                WS,
                Culture,
                ReferenceStyle.R1C1,
                ExpressionStyle.Normal)

            Dim Parsed As ParsedExpression =
                Engine.Parse(
                    Cell.Formula,
                    A1Context)

            Return Parsed.ToString(
                R1C1Context)

        End Function

        Private Function ParameterValuesEquivalent(ByVal Expected As ParameterValue,
                                                   ByVal Actual As ParameterValue,
                                                   ByRef Reason As String) As Boolean

            Reason = Nothing

            If Expected Is Nothing OrElse Actual Is Nothing Then

                If Expected Is Nothing AndAlso Actual Is Nothing Then
                    Return True
                End If

                Reason = "one ParameterValue is Nothing"
                Return False

            End If

            If Expected.IsEmpty OrElse Actual.IsEmpty Then

                If Expected.IsEmpty AndAlso Actual.IsEmpty Then
                    Return True
                End If

                Reason = "empty/non-empty mismatch"
                Return False

            End If

            If Expected.IsError OrElse Actual.IsError Then

                If Expected.IsError AndAlso
                   Actual.IsError AndAlso
                   String.Equals(
                       Expected.ErrorValue.ToString(),
                       Actual.ErrorValue.ToString(),
                       StringComparison.OrdinalIgnoreCase) Then

                    Return True

                End If

                Reason =
                    "error mismatch expected=" &
                    If(Expected.IsError,
                       Expected.ErrorValue.ToString(),
                       "(not error)") &
                    ", actual=" &
                    If(Actual.IsError,
                       Actual.ErrorValue.ToString(),
                       "(not error)")

                Return False

            End If

            If Expected.IsNumeric AndAlso
               Actual.IsNumeric Then

                Dim Difference As Double =
                    Math.Abs(
                        Expected.NumericValue -
                        Actual.NumericValue)

                Dim Tolerance As Double =
                    Math.Max(
                        0.000000001,
                        Math.Max(
                            Math.Abs(Expected.NumericValue),
                            Math.Abs(Actual.NumericValue)) *
                        0.0000000001)

                If Difference <= Tolerance Then
                    Return True
                End If

                Reason =
                    "numeric mismatch expected=" &
                    Expected.NumericValue.ToString(
                        "R",
                        CultureInfo.InvariantCulture) &
                    ", actual=" &
                    Actual.NumericValue.ToString(
                        "R",
                        CultureInfo.InvariantCulture)

                Return False

            End If

            If Expected.IsText AndAlso
               Actual.IsText Then

                If String.Equals(
                    Expected.TextValue,
                    Actual.TextValue,
                    StringComparison.Ordinal) Then

                    Return True

                End If

                Reason =
                    "text mismatch expected='" &
                    Expected.TextValue &
                    "', actual='" &
                    Actual.TextValue & "'"

                Return False

            End If

            If Expected.IsBoolean AndAlso
               Actual.IsBoolean Then

                If Expected.BooleanValue =
                   Actual.BooleanValue Then

                    Return True

                End If

                Reason =
                    "boolean mismatch expected=" &
                    Expected.BooleanValue.ToString &
                    ", actual=" &
                    Actual.BooleanValue.ToString

                Return False

            End If

            Reason =
                "ParameterValue type mismatch"

            Return False

        End Function

        Private Function CellValuesEquivalent(ByVal Expected As CellValue,
                                              ByVal Actual As CellValue,
                                              ByRef Reason As String) As Boolean

            Reason = Nothing

            If Expected.IsEmpty OrElse Actual.IsEmpty Then

                If Expected.IsEmpty AndAlso Actual.IsEmpty Then
                    Return True
                End If

                Reason = "empty/non-empty constant mismatch"
                Return False

            End If

            If Expected.IsNumeric AndAlso Actual.IsNumeric Then

                Dim Difference As Double =
                    Math.Abs(
                        Expected.NumericValue -
                        Actual.NumericValue)

                Dim Tolerance As Double =
                    Math.Max(
                        0.000000001,
                        Math.Max(
                            Math.Abs(Expected.NumericValue),
                            Math.Abs(Actual.NumericValue)) *
                        0.0000000001)

                If Difference <= Tolerance Then
                    Return True
                End If

                Reason =
                    "numeric constant mismatch"
                Return False

            End If

            If Expected.IsText AndAlso Actual.IsText Then

                If String.Equals(
                    Expected.TextValue,
                    Actual.TextValue,
                    StringComparison.Ordinal) Then

                    Return True

                End If

                Reason =
                    "text constant mismatch"
                Return False

            End If

            If Expected.IsBoolean AndAlso Actual.IsBoolean Then

                If Expected.BooleanValue =
                   Actual.BooleanValue Then

                    Return True

                End If

                Reason =
                    "boolean constant mismatch"
                Return False

            End If

            If Expected.IsError AndAlso Actual.IsError Then

                If String.Equals(
                    Expected.ErrorValue.ToString(),
                    Actual.ErrorValue.ToString(),
                    StringComparison.OrdinalIgnoreCase) Then

                    Return True

                End If

                Reason =
                    "error constant mismatch"
                Return False

            End If

            Reason =
                "constant CellValue type mismatch"
            Return False

        End Function

        Public Function IsDevelopmentIdentifiedShadowComparisonInstalled() As Boolean

            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing OrElse
               GetSchemaVersion() <> 9 Then

                Return False

            End If

            Dim ShadowWS As Worksheet =
                TryGetWorksheet(
                    WB,
                    DevelopmentShadowSheetName)

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(
                    WB,
                    MetadataSheetName)

            If ShadowWS Is Nothing OrElse
               MetadataWS Is Nothing Then

                Return False

            End If

            If Not MetadataWS.Cells(11, 1).Value.IsText Then
                Return False
            End If

            Return String.Equals(
                MetadataWS.Cells(11, 1).Value.TextValue,
                "PARITY_PASS",
                StringComparison.OrdinalIgnoreCase)

        End Function

        Public Sub DiagnoseDevelopmentIdentifiedShadowComparison()

            Dim WB As IWorkbook = GetWorkbook()


            If WB Is Nothing Then
                Return
            End If

            Dim ShadowWS As Worksheet =
                TryGetWorksheet(
                    WB,
                    DevelopmentShadowSheetName)

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(
                    WB,
                    MetadataSheetName)


            If ShadowWS IsNot Nothing Then
            End If

            If MetadataWS IsNot Nothing AndAlso
               MetadataWS.Cells(11, 1).Value.IsText Then


            End If


        End Sub

        Private Const DevelopmentRoundTripMetadataStartRow As Integer = 30

        Private Function GetDevelopmentSourceAnchorName(ByVal TargetName As String) As String

            Dim Suffix As String =
                TargetName.Substring(TargetName.Length - 1, 1)

            Return "__AbovoSrc_DevID_" & Suffix

        End Function

        Private Function GetDevelopmentRoundTripMetadataRow(ByVal TargetName As String) As Integer

            For BlockIndex As Integer = 0 To DevelopmentSingleNames.Length - 1

                If String.Equals(
                    DevelopmentSingleNames(BlockIndex),
                    TargetName,
                    StringComparison.OrdinalIgnoreCase) Then

                    Return DevelopmentRoundTripMetadataStartRow + BlockIndex

                End If

            Next

            Return -1

        End Function

        Private Function GetRangeStructuralFingerprint(ByVal SourceRange As CellRange) As String

            If SourceRange Is Nothing Then Return String.Empty

            Dim SB As New StringBuilder()

            SB.Append(SourceRange.Worksheet.Name)
            SB.Append("|")
            SB.Append(SourceRange.TopRowIndex.ToString(CultureInfo.InvariantCulture))
            SB.Append("|")
            SB.Append(SourceRange.LeftColumnIndex.ToString(CultureInfo.InvariantCulture))
            SB.Append("|")
            SB.Append(SourceRange.BottomRowIndex.ToString(CultureInfo.InvariantCulture))
            SB.Append("|")
            SB.Append(SourceRange.RightColumnIndex.ToString(CultureInfo.InvariantCulture))
            SB.Append("|")

            For RowIndex As Integer = SourceRange.TopRowIndex To SourceRange.BottomRowIndex

                For ColumnIndex As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex

                    Dim C As Cell =
                        SourceRange.Worksheet.Cells(RowIndex, ColumnIndex)

                    If Not String.IsNullOrWhiteSpace(C.Formula) Then

                        SB.Append("F:")
                        SB.Append(C.Formula)

                    ElseIf C.Value.IsEmpty Then

                        SB.Append("E:")

                    Else

                        SB.Append("V:")
                        SB.Append(C.Value.ToString())

                    End If

                    SB.Append(ChrW(30))

                Next

                SB.Append(ChrW(29))

            Next

            Using SHA As SHA256 = SHA256.Create()

                Dim Bytes() As Byte =
                    Encoding.UTF8.GetBytes(SB.ToString())

                Dim Hash() As Byte =
                    SHA.ComputeHash(Bytes)

                Dim Result As New StringBuilder(Hash.Length * 2)

                For Each B As Byte In Hash
                    Result.Append(B.ToString("x2", CultureInfo.InvariantCulture))
                Next

                Return Result.ToString()

            End Using

        End Function

        Private Sub StoreDevelopmentRoundTripSignatures(ByVal TargetName As String,
                                                        ByVal SourceRange As CellRange,
                                                        ByVal MirrorRange As CellRange)

            Dim WB As IWorkbook = GetWorkbook()
            If WB Is Nothing Then Return

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(WB, MetadataSheetName)

            If MetadataWS Is Nothing Then Return

            Dim RowIndex As Integer =
                GetDevelopmentRoundTripMetadataRow(TargetName)

            If RowIndex < 0 Then Return

            MetadataWS.Cells(RowIndex, 0).Value =
                TargetName

            MetadataWS.Cells(RowIndex, 1).Value =
                GetDevelopmentSourceAnchorName(TargetName)

            MetadataWS.Cells(RowIndex, 2).Value =
                GetRangeStructuralFingerprint(SourceRange)

            MetadataWS.Cells(RowIndex, 3).Value =
                GetRangeStructuralFingerprint(MirrorRange)

            MetadataWS.Cells(RowIndex, 4).Value =
                DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)

        End Sub

        Public Sub UpdateDevelopmentRoundTripSignatures()

            Dim WB As IWorkbook = GetWorkbook()
            If WB Is Nothing Then Return
            If GetSchemaVersion() <> 10 Then Return

            For Each TargetName As String In DevelopmentSingleNames

                Dim SourceDN As DefinedName =
                    TryGetDefinedName(
                        WB,
                        GetDevelopmentSourceAnchorName(TargetName))

                Dim MirrorDN As DefinedName =
                    TryGetDefinedName(WB, TargetName)

                If SourceDN Is Nothing OrElse
                   SourceDN.Range Is Nothing OrElse
                   MirrorDN Is Nothing OrElse
                   MirrorDN.Range Is Nothing Then

                    Continue For

                End If

                StoreDevelopmentRoundTripSignatures(
                    TargetName,
                    SourceDN.Range,
                    MirrorDN.Range)

            Next

            ExcelModels(ModelID).IsDirty = True

        End Sub

        Public Function IsDevelopmentIdentifiedProductionMirrorInstalled() As Boolean

            If GetSchemaVersion() <> 10 Then Return False

            Dim WB As IWorkbook = GetWorkbook()
            If WB Is Nothing Then Return False

            For Each TargetName As String In DevelopmentSingleNames

                Dim MirrorDN As DefinedName =
                    TryGetDefinedName(WB, TargetName)

                Dim SourceDN As DefinedName =
                    TryGetDefinedName(
                        WB,
                        GetDevelopmentSourceAnchorName(TargetName))

                If MirrorDN Is Nothing OrElse
                   MirrorDN.Range Is Nothing OrElse
                   SourceDN Is Nothing OrElse
                   SourceDN.Range Is Nothing Then

                    Return False

                End If

                If String.Equals(
                    MirrorDN.Range.Worksheet.Name,
                    "Transactional DB",
                    StringComparison.OrdinalIgnoreCase) Then

                    Return False

                End If

                If Not MirrorDN.Range.Worksheet.Name.StartsWith(
                    "__Abovo_TransDB_DevID_",
                    StringComparison.OrdinalIgnoreCase) Then

                    Return False

                End If

                If Not String.Equals(
                    SourceDN.Range.Worksheet.Name,
                    "Transactional DB",
                    StringComparison.OrdinalIgnoreCase) Then

                    Return False

                End If

            Next

            Return True

        End Function

        '=====================================================================
        ' MIGRATION 010
        '
        'Activates the separately-isolated Development Identified mirror sheets.
        'The original Transactional DB ranges are preserved and anchored under
        '__AbovoSrc_DevID_A:N for Excel round-trip detection.
        '
        'The live TransCopy_DevptSingle_A:N names are repointed to the mirrors.
        'No structural edits are made to Transactional DB.
        '=====================================================================
        Private Function ApplyMigration010ActivateDevelopmentProductionMirrors() As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            Dim SourceWS As Worksheet =
                TryGetWorksheet(WB, "Transactional DB")

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(WB, MetadataSheetName)

            If SourceWS Is Nothing OrElse MetadataWS Is Nothing Then
                Result.BError = True
                Result.StringReturn =
                    "Migration 010 requires Transactional DB and __Abovo_Metadata."
                Return Result
            End If

            Dim Validation As AbovoTransaction =
                ValidateDevelopmentSingleRanges(WB, SourceWS)

            If Validation.BError Then Return Validation

            Dim PreviousCalculationMode As WorkbookCalculationMode =
                WB.Options.CalculationMode

            Dim PreviousEngine As CalculationEngineType =
                WB.Options.CalculationEngineType

            Dim PreviousHistoryEnabled As Boolean =
                WB.History.IsEnabled

            Dim EngineChanged As Boolean = False
            Dim HistoryChanged As Boolean = False
            Dim UpdateStarted As Boolean = False
            Dim MigrationTimer As Stopwatch = Stopwatch.StartNew()

            Try


                If WB.History.IsEnabled Then
                    WB.History.IsEnabled = False
                    HistoryChanged = True
                End If

                If WB.Options.CalculationEngineType <>
                   CalculationEngineType.Recursive Then

                    WB.Options.CalculationEngineType =
                        CalculationEngineType.Recursive
                    EngineChanged = True
                End If

                WB.Options.CalculationMode =
                    WorkbookCalculationMode.Manual

                Dim Engine As FormulaEngine =
                    WB.FormulaEngine

                Dim Culture As CultureInfo =
                    CultureInfo.GetCultureInfo("en-US")

                WB.BeginUpdate()
                UpdateStarted = True

                For Each TargetName As String In DevelopmentSingleNames

                    Dim TargetDN As DefinedName =
                        TryGetDefinedName(WB, TargetName)

                    If TargetDN Is Nothing OrElse
                       TargetDN.Range Is Nothing Then

                        Throw New InvalidOperationException(
                            "Migration 010 could not resolve " &
                            TargetName & ".")

                    End If

                    Dim OriginalRange As CellRange =
                        TargetDN.Range

                    If Not String.Equals(
                        OriginalRange.Worksheet.Name,
                        "Transactional DB",
                        StringComparison.OrdinalIgnoreCase) Then

                        Throw New InvalidOperationException(
                            TargetName &
                            " is not currently on Transactional DB; migration 010 will not guess its source.")

                    End If

                    Dim SourceAnchorName As String =
                        GetDevelopmentSourceAnchorName(TargetName)

                    Dim SourceAnchor As DefinedName =
                        TryGetDefinedName(WB, SourceAnchorName)

                    If SourceAnchor Is Nothing Then

                        WB.DefinedNames.Add(
                            SourceAnchorName,
                            "'" &
                            SourceWS.Name.Replace("'", "''") &
                            "'!" &
                            OriginalRange.GetReferenceA1())

                        SourceAnchor =
                            TryGetDefinedName(WB, SourceAnchorName)

                    Else

                        SourceAnchor.Range =
                            OriginalRange

                    End If

                    Dim MirrorSheetName As String =
                        GetDevelopmentMirrorSheetName(TargetName)

                    Dim MirrorWS As Worksheet =
                        TryGetWorksheet(WB, MirrorSheetName)

                    If MirrorWS IsNot Nothing Then
                        WB.Worksheets.Remove(MirrorWS)
                    End If

                    MirrorWS =
                        WB.Worksheets.Add(MirrorSheetName)

                    MirrorWS.VisibilityType =
                        WorksheetVisibilityType.VeryHidden

                    For SupportRow As Integer = 0 To 5

                        For SupportColumn As Integer = 0 To 73

                            ReconstructCellInShadow(
                                Engine,
                                Culture,
                                SourceWS,
                                SourceWS.Cells(SupportRow, SupportColumn),
                                MirrorWS,
                                MirrorWS.Cells(SupportRow, SupportColumn))

                        Next

                    Next

                    For RowIndex As Integer = OriginalRange.TopRowIndex To OriginalRange.BottomRowIndex

                        For ColumnIndex As Integer = OriginalRange.LeftColumnIndex To OriginalRange.RightColumnIndex

                            ReconstructCellInShadow(
                                Engine,
                                Culture,
                                SourceWS,
                                SourceWS.Cells(RowIndex, ColumnIndex),
                                MirrorWS,
                                MirrorWS.Cells(RowIndex, ColumnIndex))

                        Next

                    Next

                    Dim MirrorRange As CellRange =
                        MirrorWS.Range.FromLTRB(
                            OriginalRange.LeftColumnIndex,
                            OriginalRange.TopRowIndex,
                            OriginalRange.RightColumnIndex,
                            OriginalRange.BottomRowIndex)

                    TargetDN.Range =
                        MirrorRange

                    StoreDevelopmentRoundTripSignatures(
                        TargetName,
                        SourceAnchor.Range,
                        MirrorRange)


                    System.Windows.Forms.Application.DoEvents()

                Next

                MetadataWS.Cells(25, 0).Value =
                    "DevelopmentProductionMirror"

                MetadataWS.Cells(25, 1).Value =
                    "ACTIVE"

                MetadataWS.Cells(26, 0).Value =
                    "DevelopmentRoundTripGuard"

                MetadataWS.Cells(26, 1).Value =
                    "SOURCE_AND_MIRROR_STRUCTURAL_FINGERPRINTS"

                MetadataWS.Cells(27, 0).Value =
                    "Migration010AppliedUtc"

                MetadataWS.Cells(27, 1).Value =
                    DateTime.UtcNow.ToString(
                        "o",
                        CultureInfo.InvariantCulture)

                MetadataWS.Cells(0, 1).Value = 10
                ExcelModels(ModelID).IsDirty = True

                WB.EndUpdate()
                UpdateStarted = False

                Result.StringReturn =
                    "Migration 010 activated Development Identified production mirrors."


            Catch ex As Exception

                Result.BError = True
                Result.StringReturn =
                    "Migration 010 failed: " &
                    ex.Message


            Finally

                If UpdateStarted Then
                    Try
                        WB.EndUpdate()
                    Catch
                    End Try
                End If

                Try
                    WB.Options.CalculationMode =
                        PreviousCalculationMode
                Catch
                End Try

                If EngineChanged Then
                    Try
                        WB.Options.CalculationEngineType =
                            PreviousEngine
                    Catch
                    End Try
                End If

                If HistoryChanged Then
                    Try
                        WB.History.IsEnabled =
                            PreviousHistoryEnabled
                    Catch
                    End Try
                End If

            End Try

            Return Result

        End Function

        '=====================================================================
        ' EXCEL ROUND-TRIP GUARD
        '
        'Detects edits made directly to the preserved Transactional DB blocks
        'after migration. Formula cached results are deliberately excluded from
        'the fingerprint; formula text and constants are included.
        '
        'If only TransactionDB changed and geometry still matches the live
        'mirror, the external edit is imported automatically.
        '
        'If both TransactionDB and mirror changed, automatic reconciliation is
        'refused because this is a genuine two-sided conflict.
        '=====================================================================
        Public Function ReconcileExternalDevelopmentTransactionalDBEdits() As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}

            If GetSchemaVersion() <> 10 Then Return Result

            Dim WB As IWorkbook = GetWorkbook()
            If WB Is Nothing Then Return Result

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(WB, MetadataSheetName)

            If MetadataWS Is Nothing Then
                Result.BError = True
                Result.StringReturn =
                    "Development round-trip metadata is missing."
                Return Result
            End If

            Dim Engine As FormulaEngine =
                WB.FormulaEngine

            Dim Culture As CultureInfo =
                CultureInfo.GetCultureInfo("en-US")

            Dim ImportedBlocks As Integer = 0

            For Each TargetName As String In DevelopmentSingleNames

                Dim RowIndex As Integer =
                    GetDevelopmentRoundTripMetadataRow(TargetName)

                Dim SourceDN As DefinedName =
                    TryGetDefinedName(
                        WB,
                        GetDevelopmentSourceAnchorName(TargetName))

                Dim MirrorDN As DefinedName =
                    TryGetDefinedName(WB, TargetName)

                If RowIndex < 0 OrElse
                   SourceDN Is Nothing OrElse
                   SourceDN.Range Is Nothing OrElse
                   MirrorDN Is Nothing OrElse
                   MirrorDN.Range Is Nothing Then

                    Result.BError = True
                    Result.StringReturn =
                        "Development round-trip metadata/ranges are incomplete for " &
                        TargetName & "."
                    Return Result

                End If

                Dim StoredSourceFingerprint As String =
                    MetadataWS.Cells(RowIndex, 2).Value.ToString()

                Dim StoredMirrorFingerprint As String =
                    MetadataWS.Cells(RowIndex, 3).Value.ToString()

                Dim CurrentSourceFingerprint As String =
                    GetRangeStructuralFingerprint(SourceDN.Range)

                Dim CurrentMirrorFingerprint As String =
                    GetRangeStructuralFingerprint(MirrorDN.Range)

                Dim SourceChanged As Boolean =
                    Not String.Equals(
                        StoredSourceFingerprint,
                        CurrentSourceFingerprint,
                        StringComparison.Ordinal)

                Dim MirrorChanged As Boolean =
                    Not String.Equals(
                        StoredMirrorFingerprint,
                        CurrentMirrorFingerprint,
                        StringComparison.Ordinal)

                If Not SourceChanged AndAlso
                   Not MirrorChanged Then

                    Continue For

                End If

                If SourceChanged AndAlso MirrorChanged Then

                    Result.BError = True
                    Result.StringReturn =
                        "Development TransactionDB round-trip conflict detected for " &
                        TargetName &
                        ": both the preserved Transactional DB block and the live Summit mirror changed since the last trusted state. " &
                        "Summit has not overwritten either version."


                    Return Result

                End If

                If MirrorChanged AndAlso
                   Not SourceChanged Then

                    'The live mirror is authoritative. This can happen after a
                    'Summit-side operation if metadata was not persisted at the
                    'same instant; accept it and refresh the trusted fingerprint.
                    StoreDevelopmentRoundTripSignatures(
                        TargetName,
                        SourceDN.Range,
                        MirrorDN.Range)


                    Continue For

                End If

                'Source changed only: candidate Excel-side TransactionDB edit.
                If SourceDN.Range.RowCount <> MirrorDN.Range.RowCount OrElse
                   SourceDN.Range.ColumnCount <> MirrorDN.Range.ColumnCount OrElse
                   SourceDN.Range.TopRowIndex <> MirrorDN.Range.TopRowIndex OrElse
                   SourceDN.Range.LeftColumnIndex <> MirrorDN.Range.LeftColumnIndex Then

                    Result.BError = True
                    Result.StringReturn =
                        "External Transactional DB edit detected for " &
                        TargetName &
                        ", but its range geometry no longer matches the Summit mirror. " &
                        "Automatic import has been refused."


                    Return Result

                End If

                Dim MirrorWS As Worksheet =
                    MirrorDN.Range.Worksheet

                WB.BeginUpdate()

                Try

                    For R As Integer = SourceDN.Range.TopRowIndex To SourceDN.Range.BottomRowIndex

                        For C As Integer = SourceDN.Range.LeftColumnIndex To SourceDN.Range.RightColumnIndex

                            ReconstructCellInShadow(
                                Engine,
                                Culture,
                                SourceDN.Range.Worksheet,
                                SourceDN.Range.Worksheet.Cells(R, C),
                                MirrorWS,
                                MirrorWS.Cells(R, C))

                        Next

                    Next

                Finally

                    WB.EndUpdate()

                End Try

                MirrorWS.Calculate()

                StoreDevelopmentRoundTripSignatures(
                    TargetName,
                    SourceDN.Range,
                    MirrorDN.Range)

                ImportedBlocks += 1


                System.Windows.Forms.Application.DoEvents()

            Next

            If ImportedBlocks > 0 Then

                ExcelModels(ModelID).IsDirty = True

                Result.StringReturn =
                    ImportedBlocks.ToString &
                    " externally-edited Development TransactionDB block(s) were imported into the live Summit mirrors."

            Else

                Result.StringReturn =
                    "Development TransactionDB round-trip guard passed."

            End If

            Return Result

        End Function

        Private Function GetDevelopmentMirrorSheetName(ByVal TargetName As String) As String

            Dim Suffix As String =
                TargetName.Substring(TargetName.Length - 1, 1)

            Return "__Abovo_TransDB_DevID_" & Suffix

        End Function

        '=====================================================================
        ' MIGRATION 009
        '
        'DEVELOPMENT IDENTIFIED SPARSE MULTI-SHEET CAPACITY VALIDATION.
        '
        'Creates one VeryHidden worksheet per Development A:N block.
        'Each target block remains at the SAME worksheet coordinates as its
        'TransactionDB source, eliminating the collision problem that would
        'occur if all blocks shared one capacity sheet.
        '
        'For each block-specific sheet:
        '   - support/header rows 1:6 are reconstructed;
        '   - current records 1..N are reconstructed exactly;
        '   - unused future capacity rows remain blank;
        '   - the current footer remains at the current logical footer;
        '   - the sheet is recalculated;
        '   - current records are compared against TransactionDB;
        '   - one future row is dry-run syntactically without being written.
        '
        'NO TransactionDB structural edits.
        'NO TransCopy_* repointing.
        '=====================================================================
        Private Function ApplyMigration009ValidateDevelopmentMultiSheetCapacity() As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            Dim ExistingVersion As Integer =
                GetSchemaVersion()

            If ExistingVersion <> 2 AndAlso
               ExistingVersion <> 5 AndAlso
               ExistingVersion <> 6 AndAlso
               ExistingVersion <> 7 AndAlso
               ExistingVersion <> 8 Then

                Result.BError = True
                Result.StringReturn =
                    "Migration 009 requires safe schema 2, 5, 6, 7 or 8."
                Return Result

            End If

            Dim SourceWS As Worksheet =
                TryGetWorksheet(WB, "Transactional DB")

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(WB, MetadataSheetName)

            If SourceWS Is Nothing OrElse MetadataWS Is Nothing Then

                Result.BError = True
                Result.StringReturn =
                    "Migration 009 requires Transactional DB and __Abovo_Metadata."
                Return Result

            End If

            Dim Validation As AbovoTransaction =
                ValidateDevelopmentSingleRanges(WB, SourceWS)

            If Validation.BError Then Return Validation

            Dim PreviousCalculationMode As WorkbookCalculationMode =
                WB.Options.CalculationMode

            Dim PreviousEngine As CalculationEngineType =
                WB.Options.CalculationEngineType

            Dim PreviousHistoryEnabled As Boolean =
                WB.History.IsEnabled

            Dim EngineChanged As Boolean = False
            Dim HistoryChanged As Boolean = False
            Dim UpdateStarted As Boolean = False

            Dim MigrationTimer As Stopwatch =
                Stopwatch.StartNew()

            Try


                If WB.History.IsEnabled Then
                    WB.History.IsEnabled = False
                    HistoryChanged = True
                End If

                If WB.Options.CalculationEngineType <>
                   CalculationEngineType.Recursive Then

                    WB.Options.CalculationEngineType =
                        CalculationEngineType.Recursive

                    EngineChanged = True

                End If

                WB.Options.CalculationMode =
                    WorkbookCalculationMode.Manual

                Dim Engine As FormulaEngine =
                    WB.FormulaEngine

                Dim Culture As CultureInfo =
                    CultureInfo.GetCultureInfo("en-US")

                Dim TotalCurrentValueMismatches As Integer = 0
                Dim TotalCurrentFormulaMismatches As Integer = 0
                Dim TotalDryRunErrors As Integer = 0
                Dim TotalDryRunFormulaCells As Integer = 0
                Dim TotalDryRunConstantCells As Integer = 0
                Dim TotalDryRunBlankCells As Integer = 0
                Dim BlockFailures As Integer = 0

                For Each TargetName As String In DevelopmentSingleNames

                    Dim BlockTimer As Stopwatch =
                        Stopwatch.StartNew()

                    Dim SourceDN As DefinedName =
                        TryGetDefinedName(WB, TargetName)

                    If SourceDN Is Nothing OrElse
                       SourceDN.Range Is Nothing Then

                        Throw New InvalidOperationException(
                            "Migration 009 could not resolve " &
                            TargetName & ".")

                    End If

                    Dim SourceRange As CellRange =
                        SourceDN.Range

                    Dim DataTop As Integer =
                        SourceRange.TopRowIndex

                    Dim SourceFooterRow As Integer =
                        SourceRange.BottomRowIndex

                    Dim CurrentDataCount As Integer =
                        SourceRange.RowCount - 1

                    If CurrentDataCount > DevelopmentShadowCapacity Then

                        Throw New InvalidOperationException(
                            TargetName &
                            " already contains " &
                            CurrentDataCount.ToString &
                            " data rows, exceeding reserved capacity " &
                            DevelopmentShadowCapacity.ToString & ".")

                    End If

                    Dim MirrorSheetName As String =
                        GetDevelopmentMirrorSheetName(TargetName)

                    WB.BeginUpdate()
                    UpdateStarted = True

                    Dim MirrorWS As Worksheet =
                        TryGetWorksheet(WB, MirrorSheetName)

                    If MirrorWS IsNot Nothing Then
                        WB.Worksheets.Remove(MirrorWS)
                    End If

                    MirrorWS =
                        WB.Worksheets.Add(MirrorSheetName)

                    MirrorWS.VisibilityType =
                        WorksheetVisibilityType.VeryHidden

                    'Required support/header cells at identical coordinates.
                    For SupportRow As Integer = 0 To 5

                        For SupportColumn As Integer = 0 To 73

                            ReconstructCellInShadow(
                                Engine,
                                Culture,
                                SourceWS,
                                SourceWS.Cells(SupportRow, SupportColumn),
                                MirrorWS,
                                MirrorWS.Cells(SupportRow, SupportColumn))

                        Next

                    Next

                    'Copy ONLY currently-active data rows.
                    For RowIndex As Integer = DataTop To SourceFooterRow - 1

                        For ColumnIndex As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex

                            ReconstructCellInShadow(
                                Engine,
                                Culture,
                                SourceWS,
                                SourceWS.Cells(RowIndex, ColumnIndex),
                                MirrorWS,
                                MirrorWS.Cells(RowIndex, ColumnIndex))

                        Next

                    Next

                    'Current footer remains at the current logical footer row.
                    For ColumnIndex As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex

                        ReconstructCellInShadow(
                            Engine,
                            Culture,
                            SourceWS,
                            SourceWS.Cells(SourceFooterRow, ColumnIndex),
                            MirrorWS,
                            MirrorWS.Cells(SourceFooterRow, ColumnIndex))

                    Next

                    'Ensure all reserved future rows are physically blank.
                    If CurrentDataCount < DevelopmentShadowCapacity Then

                        For RowIndex As Integer = DataTop + CurrentDataCount + 1 To DataTop + DevelopmentShadowCapacity

                            MirrorWS.Range.FromLTRB(
                                SourceRange.LeftColumnIndex,
                                RowIndex,
                                SourceRange.RightColumnIndex,
                                RowIndex).ClearContents()

                        Next

                    End If

                    WB.EndUpdate()
                    UpdateStarted = False

                    System.Windows.Forms.Application.DoEvents()

                    Dim CalculationTimer As Stopwatch =
                        Stopwatch.StartNew()

                    MirrorWS.Calculate()

                    System.Windows.Forms.Application.DoEvents()

                    Dim BlockValueMismatches As Integer = 0
                    Dim BlockFormulaMismatches As Integer = 0
                    Dim BlockDryRunErrors As Integer = 0

                    'Validate current records only.
                    For RowIndex As Integer = DataTop To SourceFooterRow - 1

                        For ColumnIndex As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex

                            Dim SourceCell As Cell =
                                SourceWS.Cells(RowIndex, ColumnIndex)

                            Dim MirrorCell As Cell =
                                MirrorWS.Cells(RowIndex, ColumnIndex)

                            Dim Reason As String = Nothing

                            If Not CellValuesEquivalent(
                                SourceCell.Value,
                                MirrorCell.Value,
                                Reason) Then

                                BlockValueMismatches += 1

                            End If

                            If Not String.IsNullOrWhiteSpace(SourceCell.Formula) Then

                                Dim SourceSignature As String =
                                    GetR1C1Signature(
                                        Engine,
                                        Culture,
                                        SourceWS,
                                        SourceCell)

                                Dim MirrorSignature As String =
                                    GetR1C1Signature(
                                        Engine,
                                        Culture,
                                        MirrorWS,
                                        MirrorCell)

                                If Not String.Equals(
                                    SourceSignature,
                                    MirrorSignature,
                                    StringComparison.Ordinal) Then

                                    BlockFormulaMismatches += 1

                                End If

                            End If

                        Next

                    Next

                    'Infer one generation rule for each column.
                    Dim Rules As New Dictionary(Of Integer, ShadowGenerationRule)

                    For ColumnIndex As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex

                        Dim Rule As ShadowGenerationRule =
                            InferShadowGenerationRule(
                                Engine,
                                Culture,
                                SourceWS,
                                ColumnIndex,
                                DataTop,
                                SourceFooterRow - 1)

                        If Rule.RuleType =
                           ShadowGenerationRuleType.MixedUnsafe Then

                            Throw New InvalidOperationException(
                                TargetName &
                                " has an unsafe future-row rule in column " &
                                SourceWS.Cells(0, ColumnIndex).
                                    GetReferenceA1().
                                    Replace("1", "") &
                                ": " &
                                Rule.Detail)

                        End If

                        Rules.Add(ColumnIndex, Rule)

                    Next

                    'Dry-run ONE representative future row syntactically, without
                    'writing it to the workbook. This proves that each formula rule
                    'can be translated to the next record while leaving reserved
                    'rows completely blank.
                    If CurrentDataCount < DevelopmentShadowCapacity Then

                        Dim DryRunRow As Integer =
                            DataTop + CurrentDataCount

                        For ColumnIndex As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex

                            Dim Rule As ShadowGenerationRule =
                                Rules(ColumnIndex)

                            Select Case Rule.RuleType

                                Case ShadowGenerationRuleType.Blank

                                    TotalDryRunBlankCells += 1

                                Case ShadowGenerationRuleType.FormulaPattern

                                    Dim DryRunCell As Cell =
                                        MirrorWS.Cells(DryRunRow, ColumnIndex)

                                    Dim DryRunR1C1Context As New ExpressionContext(
                                        DryRunCell.ColumnIndex,
                                        DryRunCell.RowIndex,
                                        MirrorWS,
                                        Culture,
                                        ReferenceStyle.R1C1,
                                        ExpressionStyle.Normal)

                                    Dim DryRunA1Context As New ExpressionContext(
                                        DryRunCell.ColumnIndex,
                                        DryRunCell.RowIndex,
                                        MirrorWS,
                                        Culture,
                                        ReferenceStyle.A1,
                                        ExpressionStyle.Normal)

                                    Try

                                        Dim Parsed As ParsedExpression =
                                            Engine.Parse(
                                                Rule.FormulaR1C1,
                                                DryRunR1C1Context)

                                        Dim DryRunFormula As String =
                                            Parsed.ToString(
                                                DryRunA1Context)

                                        If String.IsNullOrWhiteSpace(
                                            DryRunFormula) Then

                                            BlockDryRunErrors += 1

                                        Else

                                            TotalDryRunFormulaCells += 1

                                        End If

                                    Catch

                                        BlockDryRunErrors += 1

                                    End Try

                                Case ShadowGenerationRuleType.InvariantConstant,
                                     ShadowGenerationRuleType.IntegerSequence,
                                     ShadowGenerationRuleType.NumericLinearSequence

                                    TotalDryRunConstantCells += 1

                                Case Else

                                    BlockDryRunErrors += 1

                            End Select

                        Next

                    End If

                    'Validate that all reserved rows after the current footer are blank.
                    If CurrentDataCount < DevelopmentShadowCapacity Then

                        For RowIndex As Integer = SourceFooterRow + 1 To DataTop + DevelopmentShadowCapacity

                            For ColumnIndex As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex

                                Dim ReservedCell As Cell =
                                    MirrorWS.Cells(RowIndex, ColumnIndex)

                                If Not ReservedCell.Value.IsEmpty OrElse
                                   Not String.IsNullOrWhiteSpace(ReservedCell.Formula) Then

                                    BlockDryRunErrors += 1

                                End If

                            Next

                        Next

                    End If

                    TotalCurrentValueMismatches +=
                        BlockValueMismatches

                    TotalCurrentFormulaMismatches +=
                        BlockFormulaMismatches

                    TotalDryRunErrors +=
                        BlockDryRunErrors

                    If BlockValueMismatches <> 0 OrElse
                       BlockFormulaMismatches <> 0 OrElse
                       BlockDryRunErrors <> 0 Then

                        BlockFailures += 1

                    End If


                    System.Windows.Forms.Application.DoEvents()

                Next


                Dim ValidationPassed As Boolean =
                    TotalCurrentValueMismatches = 0 AndAlso
                    TotalCurrentFormulaMismatches = 0 AndAlso
                    TotalDryRunErrors = 0 AndAlso
                    BlockFailures = 0

                If ValidationPassed Then


                Else


                End If

                WB.BeginUpdate()
                UpdateStarted = True

                MetadataWS.Cells(21, 0).Value =
                    "DevelopmentSparseMultiSheetCapacityValidation"

                MetadataWS.Cells(21, 1).Value =
                    If(ValidationPassed,
                       "PASS",
                       "ANALYSE")

                MetadataWS.Cells(22, 0).Value =
                    "DevelopmentMultiSheetCapacity"

                MetadataWS.Cells(22, 1).Value =
                    DevelopmentShadowCapacity

                MetadataWS.Cells(23, 0).Value =
                    "DevelopmentSparseMultiSheetBlockFailures"

                MetadataWS.Cells(23, 1).Value =
                    BlockFailures

                MetadataWS.Cells(24, 0).Value =
                    "Migration009AppliedUtc"

                MetadataWS.Cells(24, 1).Value =
                    DateTime.UtcNow.ToString(
                        "o",
                        CultureInfo.InvariantCulture)

                MetadataWS.Cells(0, 1).Value = 9
                ExcelModels(ModelID).IsDirty = True

                Result.StringReturn =
                    "Migration 009 completed sparse Development multi-sheet capacity validation. " &
                    "Original TransCopy_* names remain unchanged."


            Catch ex As Exception

                Result.BError = True
                Result.StringReturn =
                    "Migration 009 failed: " &
                    ex.Message


            Finally

                If UpdateStarted Then
                    Try
                        WB.EndUpdate()
                    Catch
                    End Try
                End If

                Try
                    WB.Options.CalculationMode =
                        PreviousCalculationMode
                Catch
                End Try

                If EngineChanged Then
                    Try
                        WB.Options.CalculationEngineType =
                            PreviousEngine
                    Catch
                    End Try
                End If

                If HistoryChanged Then
                    Try
                        WB.History.IsEnabled =
                            PreviousHistoryEnabled
                    Catch
                    End Try
                End If

            End Try

            Return Result

        End Function

        Private Sub ReconstructCellAtDifferentCoordinate(ByVal Engine As FormulaEngine,
                                                         ByVal Culture As CultureInfo,
                                                         ByVal SourceWS As Worksheet,
                                                         ByVal SourceCell As Cell,
                                                         ByVal TargetWS As Worksheet,
                                                         ByVal TargetCell As Cell)

            If String.IsNullOrWhiteSpace(SourceCell.Formula) Then

                TargetCell.Value =
                    SourceCell.Value

                Return

            End If

            Dim SourceA1Context As New ExpressionContext(
                SourceCell.ColumnIndex,
                SourceCell.RowIndex,
                SourceWS,
                Culture,
                ReferenceStyle.A1,
                ExpressionStyle.Normal)

            Dim ParsedSource As ParsedExpression =
                Engine.Parse(
                    SourceCell.Formula,
                    SourceA1Context)

            Dim SourceR1C1Context As New ExpressionContext(
                SourceCell.ColumnIndex,
                SourceCell.RowIndex,
                SourceWS,
                Culture,
                ReferenceStyle.R1C1,
                ExpressionStyle.Normal)

            Dim R1C1Formula As String =
                ParsedSource.ToString(
                    SourceR1C1Context)

            Dim TargetR1C1Context As New ExpressionContext(
                TargetCell.ColumnIndex,
                TargetCell.RowIndex,
                TargetWS,
                Culture,
                ReferenceStyle.R1C1,
                ExpressionStyle.Normal)

            Dim ParsedTarget As ParsedExpression =
                Engine.Parse(
                    R1C1Formula,
                    TargetR1C1Context)

            Dim TargetA1Context As New ExpressionContext(
                TargetCell.ColumnIndex,
                TargetCell.RowIndex,
                TargetWS,
                Culture,
                ReferenceStyle.A1,
                ExpressionStyle.Normal)

            Dim TargetFormula As String =
                ParsedTarget.ToString(
                    TargetA1Context)

            If Not TargetFormula.StartsWith(
                "=",
                StringComparison.Ordinal) Then

                TargetFormula =
                    "=" & TargetFormula

            End If

            TargetCell.Formula =
                TargetFormula

        End Sub

        Private Enum ShadowGenerationRuleType

            Blank
            FormulaPattern
            InvariantConstant
            IntegerSequence
            NumericLinearSequence
            MixedUnsafe

        End Enum

        Private Class ShadowGenerationRule

            Public RuleType As ShadowGenerationRuleType
            Public FormulaR1C1 As String
            Public ConstantValue As CellValue
            Public FirstNumericValue As Double
            Public NumericStep As Double
            Public Detail As String

        End Class

        '=====================================================================
        ' MIGRATION 008
        '
        'DEVELOPMENT SHADOW GENERATION-RULE AUDIT.
        '
        'Schema 7 proved exact parity for all 12,432 currently-active cells.
        'Schema 8 keeps the same safe coordinate-preserving shadow construction,
        'recalculates it, requires exact parity, then classifies how each data
        'column can be generated for future records.
        '
        'No TransactionDB structural edits.
        'No TransCopy_* repointing.
        'No future rows are populated yet.
        '=====================================================================
        Private Function ApplyMigration008AuditDevelopmentShadowGenerationRules() As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            Dim ExistingVersion As Integer = GetSchemaVersion()

            If ExistingVersion <> 2 AndAlso
               ExistingVersion <> 5 AndAlso
               ExistingVersion <> 6 AndAlso
               ExistingVersion <> 7 Then

                Result.BError = True
                Result.StringReturn =
                    "Migration 008 requires safe schema 2, 5, 6 or 7."
                Return Result

            End If

            Dim SourceWS As Worksheet =
                TryGetWorksheet(WB, "Transactional DB")

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(WB, MetadataSheetName)

            If SourceWS Is Nothing OrElse MetadataWS Is Nothing Then

                Result.BError = True
                Result.StringReturn =
                    "Migration 008 requires Transactional DB and __Abovo_Metadata."
                Return Result

            End If

            Dim Validation As AbovoTransaction =
                ValidateDevelopmentSingleRanges(WB, SourceWS)

            If Validation.BError Then Return Validation

            Dim PreviousCalculationMode As WorkbookCalculationMode =
                WB.Options.CalculationMode

            Dim PreviousEngine As CalculationEngineType =
                WB.Options.CalculationEngineType

            Dim PreviousHistoryEnabled As Boolean =
                WB.History.IsEnabled

            Dim EngineChanged As Boolean = False
            Dim HistoryChanged As Boolean = False
            Dim UpdateStarted As Boolean = False
            Dim MigrationTimer As Stopwatch = Stopwatch.StartNew()

            Try


                If WB.History.IsEnabled Then
                    WB.History.IsEnabled = False
                    HistoryChanged = True
                End If

                If WB.Options.CalculationEngineType <>
                   CalculationEngineType.Recursive Then

                    WB.Options.CalculationEngineType =
                        CalculationEngineType.Recursive

                    EngineChanged = True

                End If

                WB.Options.CalculationMode =
                    WorkbookCalculationMode.Manual

                WB.BeginUpdate()
                UpdateStarted = True

                Dim ShadowWS As Worksheet =
                    TryGetWorksheet(WB, DevelopmentShadowSheetName)

                If ShadowWS IsNot Nothing Then
                    WB.Worksheets.Remove(ShadowWS)
                End If

                ShadowWS =
                    WB.Worksheets.Add(DevelopmentShadowSheetName)

                ShadowWS.VisibilityType =
                    WorksheetVisibilityType.VeryHidden

                Dim Engine As FormulaEngine =
                    WB.FormulaEngine

                Dim Culture As CultureInfo =
                    CultureInfo.GetCultureInfo("en-US")

                'Support/header cells at identical coordinates.
                For SupportRow As Integer = 0 To 5

                    For SupportColumn As Integer = 0 To 73

                        ReconstructCellInShadow(
                            Engine,
                            Culture,
                            SourceWS,
                            SourceWS.Cells(SupportRow, SupportColumn),
                            ShadowWS,
                            ShadowWS.Cells(SupportRow, SupportColumn))

                    Next

                Next

                'Build every active block before calculating.
                For Each TargetName As String In DevelopmentSingleNames

                    Dim SourceDN As DefinedName =
                        TryGetDefinedName(WB, TargetName)

                    If SourceDN Is Nothing OrElse
                       SourceDN.Range Is Nothing Then

                        Throw New InvalidOperationException(
                            "Migration 008 could not resolve " &
                            TargetName & ".")

                    End If

                    Dim SourceRange As CellRange =
                        SourceDN.Range

                    For SourceRow As Integer = SourceRange.TopRowIndex To SourceRange.BottomRowIndex

                        For SourceColumn As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex

                            ReconstructCellInShadow(
                                Engine,
                                Culture,
                                SourceWS,
                                SourceWS.Cells(SourceRow, SourceColumn),
                                ShadowWS,
                                ShadowWS.Cells(SourceRow, SourceColumn))

                        Next

                    Next

                    'Yield between blocks so the STA/UI thread can pump Windows
                    'messages during the long migration.
                    System.Windows.Forms.Application.DoEvents()

                Next

                If UpdateStarted Then
                    WB.EndUpdate()
                    UpdateStarted = False
                End If

                System.Windows.Forms.Application.DoEvents()

                Dim CalculationTimer As Stopwatch =
                    Stopwatch.StartNew()

                ShadowWS.Calculate()

                System.Windows.Forms.Application.DoEvents()


                'Exact parity gate before any generation-rule inference.
                Dim ValueMismatches As Integer = 0
                Dim FormulaMismatches As Integer = 0
                Dim ComparisonErrors As Integer = 0

                For Each TargetName As String In DevelopmentSingleNames

                    Dim SourceRange As CellRange =
                        TryGetDefinedName(WB, TargetName).Range

                    For SourceRow As Integer = SourceRange.TopRowIndex To SourceRange.BottomRowIndex

                        For SourceColumn As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex

                            Dim SourceCell As Cell =
                                SourceWS.Cells(SourceRow, SourceColumn)

                            Dim ShadowCell As Cell =
                                ShadowWS.Cells(SourceRow, SourceColumn)

                            Dim Reason As String = Nothing

                            Try

                                If Not CellValuesEquivalent(
                                    SourceCell.Value,
                                    ShadowCell.Value,
                                    Reason) Then

                                    ValueMismatches += 1

                                End If

                                If Not String.IsNullOrWhiteSpace(SourceCell.Formula) Then

                                    Dim SourceSignature As String =
                                        GetR1C1Signature(
                                            Engine,
                                            Culture,
                                            SourceWS,
                                            SourceCell)

                                    Dim ShadowSignature As String =
                                        GetR1C1Signature(
                                            Engine,
                                            Culture,
                                            ShadowWS,
                                            ShadowCell)

                                    If Not String.Equals(
                                        SourceSignature,
                                        ShadowSignature,
                                        StringComparison.Ordinal) Then

                                        FormulaMismatches += 1

                                    End If

                                End If

                            Catch

                                ComparisonErrors += 1

                            End Try

                        Next

                    Next

                    'Keep the UI/COM apartment responsive during the full
                    'parity scan.
                    System.Windows.Forms.Application.DoEvents()

                Next


                If ValueMismatches <> 0 OrElse
                   FormulaMismatches <> 0 OrElse
                   ComparisonErrors <> 0 Then

                    Throw New InvalidOperationException(
                        "Schema 8 parity gate failed; generation rules will not be inferred.")

                End If


                Dim UnsafeCount As Integer = 0
                Dim FormulaRuleCount As Integer = 0
                Dim BlankRuleCount As Integer = 0
                Dim InvariantRuleCount As Integer = 0
                Dim IntegerSequenceCount As Integer = 0
                Dim NumericLinearCount As Integer = 0

                Dim CrossBlockRules As New Dictionary(Of Integer, String)

                For Each TargetName As String In DevelopmentSingleNames

                    Dim SourceRange As CellRange =
                        TryGetDefinedName(WB, TargetName).Range

                    Dim DataTop As Integer =
                        SourceRange.TopRowIndex

                    Dim DataBottom As Integer =
                        SourceRange.BottomRowIndex - 1


                    For ColumnIndex As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex

                        Dim Rule As ShadowGenerationRule =
                            InferShadowGenerationRule(
                                Engine,
                                Culture,
                                SourceWS,
                                ColumnIndex,
                                DataTop,
                                DataBottom)

                        Dim RelativeColumn As Integer =
                            ColumnIndex -
                            SourceRange.LeftColumnIndex

                        Dim ColumnName As String =
                            SourceWS.Cells(0, ColumnIndex).
                            GetReferenceA1().
                            Replace("1", "")

                        Select Case Rule.RuleType

                            Case ShadowGenerationRuleType.Blank
                                BlankRuleCount += 1

                            Case ShadowGenerationRuleType.FormulaPattern
                                FormulaRuleCount += 1

                            Case ShadowGenerationRuleType.InvariantConstant
                                InvariantRuleCount += 1

                            Case ShadowGenerationRuleType.IntegerSequence
                                IntegerSequenceCount += 1

                            Case ShadowGenerationRuleType.NumericLinearSequence
                                NumericLinearCount += 1

                            Case ShadowGenerationRuleType.MixedUnsafe
                                UnsafeCount += 1

                        End Select

                        Dim RuleSignature As String =
                            Rule.RuleType.ToString &
                            "|" &
                            If(Rule.FormulaR1C1, String.Empty) &
                            "|" &
                            Rule.NumericStep.ToString(
                                "R",
                                CultureInfo.InvariantCulture)

                        If Not CrossBlockRules.ContainsKey(RelativeColumn) Then

                            CrossBlockRules.Add(
                                RelativeColumn,
                                RuleSignature)

                        ElseIf Not String.Equals(
                            CrossBlockRules(RelativeColumn),
                            RuleSignature,
                            StringComparison.Ordinal) Then


                        End If

                        If (RelativeColumn Mod 4) = 0 Then
                            System.Windows.Forms.Application.DoEvents()
                        End If

                        If Rule.RuleType =
                           ShadowGenerationRuleType.MixedUnsafe Then


                        ElseIf Rule.RuleType =
                               ShadowGenerationRuleType.IntegerSequence OrElse
                               Rule.RuleType =
                               ShadowGenerationRuleType.NumericLinearSequence Then


                        End If

                    Next

                    'The rule audit can be CPU-heavy because formula patterns
                    'are normalised.  Yield once per Development block.
                    System.Windows.Forms.Application.DoEvents()

                Next


                If UnsafeCount = 0 Then


                Else


                End If

                WB.BeginUpdate()
                UpdateStarted = True

                MetadataWS.Cells(11, 0).Value =
                    "DevelopmentIdentifiedShadowComparison"

                MetadataWS.Cells(11, 1).Value =
                    "PARITY_PASS"

                MetadataWS.Cells(13, 0).Value =
                    "DevelopmentShadowCoordinateMode"

                MetadataWS.Cells(13, 1).Value =
                    "SOURCE_COORDINATES_RECALCULATED"

                MetadataWS.Cells(18, 0).Value =
                    "DevelopmentGenerationRuleAudit"

                MetadataWS.Cells(18, 1).Value =
                    If(UnsafeCount = 0,
                       "PASS",
                       "ANALYSE")

                MetadataWS.Cells(19, 0).Value =
                    "DevelopmentGenerationUnsafeCount"

                MetadataWS.Cells(19, 1).Value =
                    UnsafeCount

                MetadataWS.Cells(20, 0).Value =
                    "Migration008AppliedUtc"

                MetadataWS.Cells(20, 1).Value =
                    DateTime.UtcNow.ToString(
                        "o",
                        CultureInfo.InvariantCulture)

                MetadataWS.Cells(0, 1).Value = 8
                ExcelModels(ModelID).IsDirty = True

                Result.StringReturn =
                    "Migration 008 completed the Development future-row " &
                    "generation-rule audit. Original TransCopy_* names remain unchanged."


            Catch ex As Exception

                Result.BError = True
                Result.StringReturn =
                    "Migration 008 failed: " &
                    ex.Message


            Finally

                If UpdateStarted Then
                    Try
                        WB.EndUpdate()
                    Catch
                    End Try
                End If

                Try
                    WB.Options.CalculationMode =
                        PreviousCalculationMode
                Catch
                End Try

                If EngineChanged Then
                    Try
                        WB.Options.CalculationEngineType =
                            PreviousEngine
                    Catch
                    End Try
                End If

                If HistoryChanged Then
                    Try
                        WB.History.IsEnabled =
                            PreviousHistoryEnabled
                    Catch
                    End Try
                End If

            End Try

            Return Result

        End Function

        Private Function InferShadowGenerationRule(ByVal Engine As FormulaEngine,
                                                   ByVal Culture As CultureInfo,
                                                   ByVal WS As Worksheet,
                                                   ByVal ColumnIndex As Integer,
                                                   ByVal DataTop As Integer,
                                                   ByVal DataBottom As Integer) As ShadowGenerationRule

            Dim Rule As New ShadowGenerationRule

            Dim FormulaCount As Integer = 0
            Dim BlankCount As Integer = 0
            Dim ConstantCount As Integer = 0
            Dim FirstFormulaSignature As String = Nothing
            Dim FormulaPatternConsistent As Boolean = True

            For RowIndex As Integer = DataTop To DataBottom

                Dim CurrentCell As Cell =
                    WS.Cells(RowIndex, ColumnIndex)

                If Not String.IsNullOrWhiteSpace(CurrentCell.Formula) Then

                    FormulaCount += 1

                    Dim Signature As String =
                        GetR1C1Signature(
                            Engine,
                            Culture,
                            WS,
                            CurrentCell)

                    If FirstFormulaSignature Is Nothing Then

                        FirstFormulaSignature =
                            Signature

                    ElseIf Not String.Equals(
                        FirstFormulaSignature,
                        Signature,
                        StringComparison.Ordinal) Then

                        FormulaPatternConsistent = False

                    End If

                ElseIf CurrentCell.Value.IsEmpty Then

                    BlankCount += 1

                Else

                    ConstantCount += 1

                End If

            Next

            Dim RowCount As Integer =
                DataBottom - DataTop + 1

            If BlankCount = RowCount Then

                Rule.RuleType =
                    ShadowGenerationRuleType.Blank

                Rule.Detail =
                    "all current data rows blank"

                Return Rule

            End If

            If FormulaCount = RowCount Then

                If FormulaPatternConsistent Then

                    Rule.RuleType =
                        ShadowGenerationRuleType.FormulaPattern

                    Rule.FormulaR1C1 =
                        FirstFormulaSignature

                    Rule.Detail =
                        "single normalised R1C1 formula pattern"

                Else

                    Rule.RuleType =
                        ShadowGenerationRuleType.MixedUnsafe

                    Rule.Detail =
                        "formula column contains multiple R1C1 patterns"

                End If

                Return Rule

            End If

            If ConstantCount = RowCount Then

                Dim FirstCell As Cell =
                    WS.Cells(DataTop, ColumnIndex)

                Dim Invariant As Boolean = True

                For RowIndex As Integer =
                    DataTop + 1 To DataBottom

                    Dim Reason As String = Nothing

                    If Not CellValuesEquivalent(
                        FirstCell.Value,
                        WS.Cells(RowIndex, ColumnIndex).Value,
                        Reason) Then

                        Invariant = False
                        Exit For

                    End If

                Next

                If Invariant Then

                    Rule.RuleType =
                        ShadowGenerationRuleType.InvariantConstant

                    Rule.ConstantValue =
                        FirstCell.Value

                    Rule.Detail =
                        "same constant in every current data row"

                    Return Rule

                End If

                Dim AllNumeric As Boolean = True

                For RowIndex As Integer = DataTop To DataBottom

                    If Not WS.Cells(RowIndex, ColumnIndex).Value.IsNumeric Then
                        AllNumeric = False
                        Exit For
                    End If

                Next

                If AllNumeric AndAlso RowCount >= 2 Then

                    Dim FirstValue As Double =
                        WS.Cells(DataTop, ColumnIndex).Value.NumericValue

                    Dim SecondValue As Double =
                        WS.Cells(DataTop + 1, ColumnIndex).Value.NumericValue

                    Dim StepValue As Double =
                        SecondValue - FirstValue

                    Dim Linear As Boolean = True

                    For RowIndex As Integer =
                        DataTop + 2 To DataBottom

                        Dim Expected As Double =
                            FirstValue +
                            ((RowIndex - DataTop) * StepValue)

                        Dim Actual As Double =
                            WS.Cells(RowIndex, ColumnIndex).Value.NumericValue

                        Dim Tolerance As Double =
                            Math.Max(
                                0.000000001,
                                Math.Max(
                                    Math.Abs(Expected),
                                    Math.Abs(Actual)) *
                                0.0000000001)

                        If Math.Abs(Expected - Actual) >
                           Tolerance Then

                            Linear = False
                            Exit For

                        End If

                    Next

                    If Linear Then

                        Rule.FirstNumericValue =
                            FirstValue

                        Rule.NumericStep =
                            StepValue

                        Dim IntegerLike As Boolean =
                            Math.Abs(
                                FirstValue -
                                Math.Round(FirstValue)) <
                            0.000000001 AndAlso
                            Math.Abs(
                                StepValue -
                                Math.Round(StepValue)) <
                            0.000000001

                        If IntegerLike Then

                            Rule.RuleType =
                                ShadowGenerationRuleType.IntegerSequence

                            Rule.Detail =
                                "deterministic integer sequence"

                        Else

                            Rule.RuleType =
                                ShadowGenerationRuleType.NumericLinearSequence

                            Rule.Detail =
                                "deterministic numeric linear sequence"

                        End If

                        Return Rule

                    End If

                End If

                Rule.RuleType =
                    ShadowGenerationRuleType.MixedUnsafe

                Rule.Detail =
                    "non-formula constants vary and are not invariant/linear"

                Return Rule

            End If

            Rule.RuleType =
                ShadowGenerationRuleType.MixedUnsafe

            Rule.Detail =
                "mixed formula/blank/constant population: formulas=" &
                FormulaCount.ToString &
                ", constants=" &
                ConstantCount.ToString &
                ", blanks=" &
                BlankCount.ToString

            Return Rule

        End Function

        '=====================================================================
        ' MIGRATION 007
        '
        'RECALCULATED COORDINATE-PRESERVING DEVELOPMENT SHADOW.
        '
        'Schemas 5/6 performed formula comparison while the workbook was in
        'Manual calculation mode immediately after creating the shadow cells.
        'That can leave newly-created dependent shadow formulas with empty or
        'stale cached values.
        '
        'Schema 7:
        '   1. rebuilds the coordinate-preserving shadow;
        '   2. reconstructs support rows 1:6 at identical coordinates;
        '   3. reconstructs ALL current A:N blocks;
        '   4. ends the DevExpress update batch;
        '   5. explicitly calculates ONLY the shadow worksheet;
        '   6. compares values/formulas after that calculation.
        '
        'Transactional DB is never structurally edited.
        'TransCopy_* names are never repointed.
        '=====================================================================
        Private Function ApplyMigration007CreateRecalculatedCoordinateShadow() As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            Dim ExistingVersion As Integer =
                GetSchemaVersion()

            If ExistingVersion <> 2 AndAlso
               ExistingVersion <> 5 AndAlso
               ExistingVersion <> 6 Then

                Result.BError = True
                Result.StringReturn =
                    "Migration 007 requires safe schema 2, schema 5 or schema 6."
                Return Result

            End If

            Dim SourceWS As Worksheet =
                TryGetWorksheet(WB, "Transactional DB")

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(WB, MetadataSheetName)

            If SourceWS Is Nothing OrElse MetadataWS Is Nothing Then

                Result.BError = True
                Result.StringReturn =
                    "Migration 007 requires Transactional DB and __Abovo_Metadata."
                Return Result

            End If

            Dim Validation As AbovoTransaction =
                ValidateDevelopmentSingleRanges(WB,
                                                SourceWS)

            If Validation.BError Then Return Validation

            Dim PreviousCalculationMode As WorkbookCalculationMode =
                WB.Options.CalculationMode

            Dim PreviousEngine As CalculationEngineType =
                WB.Options.CalculationEngineType

            Dim PreviousHistoryEnabled As Boolean =
                WB.History.IsEnabled

            Dim EngineChanged As Boolean = False
            Dim HistoryChanged As Boolean = False
            Dim UpdateStarted As Boolean = False
            Dim MigrationTimer As Stopwatch =
                Stopwatch.StartNew()

            Try


                If WB.History.IsEnabled Then
                    WB.History.IsEnabled = False
                    HistoryChanged = True
                End If

                If WB.Options.CalculationEngineType <>
                   CalculationEngineType.Recursive Then

                    WB.Options.CalculationEngineType =
                        CalculationEngineType.Recursive

                    EngineChanged = True

                End If

                WB.Options.CalculationMode =
                    WorkbookCalculationMode.Manual

                WB.BeginUpdate()
                UpdateStarted = True

                Dim ShadowWS As Worksheet =
                    TryGetWorksheet(WB,
                                    DevelopmentShadowSheetName)

                If ShadowWS IsNot Nothing Then
                    WB.Worksheets.Remove(ShadowWS)
                End If

                ShadowWS =
                    WB.Worksheets.Add(
                        DevelopmentShadowSheetName)

                ShadowWS.VisibilityType =
                    WorksheetVisibilityType.VeryHidden

                Dim Engine As FormulaEngine =
                    WB.FormulaEngine

                Dim Culture As CultureInfo =
                    CultureInfo.GetCultureInfo("en-US")

                'Same-position support/header cells referenced by unqualified
                'TransactionDB formulas.
                For SupportRow As Integer = 0 To 5

                    For SupportColumn As Integer = 0 To 73

                        ReconstructCellInShadow(
                            Engine,
                            Culture,
                            SourceWS,
                            SourceWS.Cells(SupportRow,
                                           SupportColumn),
                            ShadowWS,
                            ShadowWS.Cells(SupportRow,
                                           SupportColumn))

                    Next

                Next

                'FIRST PASS: construct every Development block.  Do not perform
                'value parity checks yet because dependent shadow cells have not
                'been recalculated.
                For Each TargetName As String In DevelopmentSingleNames

                    Dim SourceDN As DefinedName =
                        TryGetDefinedName(WB,
                                          TargetName)

                    If SourceDN Is Nothing OrElse
                       SourceDN.Range Is Nothing Then

                        Throw New InvalidOperationException(
                            "Migration 007 could not resolve " &
                            TargetName & ".")

                    End If

                    Dim SourceRange As CellRange =
                        SourceDN.Range

                    Dim BuildTimer As Stopwatch =
                        Stopwatch.StartNew()

                    For SourceRow As Integer = SourceRange.TopRowIndex To SourceRange.BottomRowIndex

                        For SourceColumn As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex

                            ReconstructCellInShadow(
                                Engine,
                                Culture,
                                SourceWS,
                                SourceWS.Cells(SourceRow,
                                               SourceColumn),
                                ShadowWS,
                                ShadowWS.Cells(SourceRow,
                                               SourceColumn))

                        Next

                    Next


                Next

                'Commit all formula assignments before calculation.
                If UpdateStarted Then
                    WB.EndUpdate()
                    UpdateStarted = False
                End If

                Dim CalculationTimer As Stopwatch =
                    Stopwatch.StartNew()

                ShadowWS.Calculate()


                'SECOND PASS: compare only after the shadow dependency graph has
                'been recalculated.
                Dim TotalCells As Integer = 0
                Dim FormulaCells As Integer = 0
                Dim ConstantCells As Integer = 0
                Dim ValueMismatches As Integer = 0
                Dim FormulaMismatches As Integer = 0
                Dim EvaluationErrors As Integer = 0
                Dim ExamplesPrinted As Integer = 0
                Const MaximumExamples As Integer = 100

                For Each TargetName As String In DevelopmentSingleNames

                    Dim SourceDN As DefinedName =
                        TryGetDefinedName(WB,
                                          TargetName)

                    Dim SourceRange As CellRange =
                        SourceDN.Range

                    Dim CompareTimer As Stopwatch =
                        Stopwatch.StartNew()

                    For SourceRow As Integer = SourceRange.TopRowIndex To SourceRange.BottomRowIndex

                        For SourceColumn As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex

                            Dim SourceCell As Cell =
                                SourceWS.Cells(SourceRow,
                                               SourceColumn)

                            Dim ShadowCell As Cell =
                                ShadowWS.Cells(SourceRow,
                                               SourceColumn)

                            TotalCells += 1

                            If String.IsNullOrWhiteSpace(SourceCell.Formula) Then
                                ConstantCells += 1
                            Else
                                FormulaCells += 1
                            End If

                            Dim ValueReason As String = Nothing

                            Try

                                'After ShadowWS.Calculate(), compare the actual
                                'calculated cell values.  This avoids using a
                                'one-off evaluator against stale dependent-cell
                                'caches.
                                If Not CellValuesEquivalent(
                                    SourceCell.Value,
                                    ShadowCell.Value,
                                    ValueReason) Then

                                    ValueMismatches += 1

                                    If ExamplesPrinted <
                                       MaximumExamples Then


                                        ExamplesPrinted += 1

                                    End If

                                End If

                            Catch ex As Exception

                                EvaluationErrors += 1

                                If ExamplesPrinted <
                                   MaximumExamples Then


                                    ExamplesPrinted += 1

                                End If

                            End Try

                            If Not String.IsNullOrWhiteSpace(
                                SourceCell.Formula) Then

                                Dim SourceSignature As String =
                                    GetR1C1Signature(
                                        Engine,
                                        Culture,
                                        SourceWS,
                                        SourceCell)

                                Dim ShadowSignature As String =
                                    GetR1C1Signature(
                                        Engine,
                                        Culture,
                                        ShadowWS,
                                        ShadowCell)

                                If Not String.Equals(
                                    SourceSignature,
                                    ShadowSignature,
                                    StringComparison.Ordinal) Then

                                    FormulaMismatches += 1

                                    If ExamplesPrinted <
                                       MaximumExamples Then


                                        ExamplesPrinted += 1

                                    End If

                                End If

                            End If

                        Next

                    Next


                Next

                WB.BeginUpdate()
                UpdateStarted = True

                MetadataWS.Cells(11, 0).Value =
                    "DevelopmentIdentifiedShadowComparison"

                MetadataWS.Cells(12, 0).Value =
                    "DevelopmentShadowSheet"

                MetadataWS.Cells(12, 1).Value =
                    DevelopmentShadowSheetName

                MetadataWS.Cells(13, 0).Value =
                    "DevelopmentShadowCoordinateMode"

                MetadataWS.Cells(13, 1).Value =
                    "SOURCE_COORDINATES_RECALCULATED"

                MetadataWS.Cells(14, 0).Value =
                    "DevelopmentShadowValueMismatches"

                MetadataWS.Cells(14, 1).Value =
                    ValueMismatches

                MetadataWS.Cells(15, 0).Value =
                    "DevelopmentShadowFormulaMismatches"

                MetadataWS.Cells(15, 1).Value =
                    FormulaMismatches

                MetadataWS.Cells(16, 0).Value =
                    "DevelopmentShadowEvaluationErrors"

                MetadataWS.Cells(16, 1).Value =
                    EvaluationErrors

                MetadataWS.Cells(17, 0).Value =
                    "Migration007AppliedUtc"

                MetadataWS.Cells(17, 1).Value =
                    DateTime.UtcNow.ToString(
                        "o",
                        CultureInfo.InvariantCulture)

                If ValueMismatches = 0 AndAlso
                   FormulaMismatches = 0 AndAlso
                   EvaluationErrors = 0 Then

                    MetadataWS.Cells(11, 1).Value =
                        "PARITY_PASS"

                Else

                    MetadataWS.Cells(11, 1).Value =
                        "PARITY_ANALYSE"

                End If

                MetadataWS.Cells(0, 1).Value = 7
                ExcelModels(ModelID).IsDirty = True


                If ValueMismatches = 0 AndAlso
                   FormulaMismatches = 0 AndAlso
                   EvaluationErrors = 0 Then


                    Result.StringReturn =
                        "Migration 007 achieved exact recalculated coordinate-shadow parity. " &
                        "Original TransCopy_* names remain unchanged."

                Else


                    Result.StringReturn =
                        "Migration 007 created/recalculated the coordinate shadow but " &
                        "parity still requires analysis. Original TransCopy_* names remain unchanged."

                End If


            Catch ex As Exception

                Result.BError = True
                Result.StringReturn =
                    "Migration 007 failed: " &
                    ex.Message


            Finally

                If UpdateStarted Then
                    Try
                        WB.EndUpdate()
                    Catch
                    End Try
                End If

                Try
                    WB.Options.CalculationMode =
                        PreviousCalculationMode
                Catch
                End Try

                If EngineChanged Then
                    Try
                        WB.Options.CalculationEngineType =
                            PreviousEngine
                    Catch
                    End Try
                End If

                If HistoryChanged Then
                    Try
                        WB.History.IsEnabled =
                            PreviousHistoryEnabled
                    Catch
                    End Try
                End If

            End Try

            Return Result

        End Function

        '=====================================================================
        ' MIGRATION 006
        '
        'COORDINATE-PRESERVING DEVELOPMENT SHADOW COMPARISON.
        '
        'Schema 5 proved that the formulas can be translated perfectly
        '(0 normalised formula mismatches), but compact relocation changed
        'formula semantics because some formulas depend on their worksheet
        'row/column context or on same-sheet support cells such as row 5.
        '
        'Schema 6 therefore:
        '   * recreates the shadow worksheet from scratch;
        '   * copies TransactionDB support/header rows 1:6 into the SAME
        '     coordinates on the shadow worksheet;
        '   * reconstructs every active A:N block at the SAME row and column
        '     coordinates as Transactional DB;
        '   * compares formula values and R1C1 signatures again;
        '   * makes NO structural edit to Transactional DB;
        '   * does NOT repoint TransCopy_* names.
        '
        'No spare/future records are generated in this comparison stage.
        '=====================================================================
        Private Function ApplyMigration006CreateCoordinatePreservingDevelopmentShadow() As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            Dim ExistingVersion As Integer = GetSchemaVersion()

            If ExistingVersion <> 2 AndAlso
               ExistingVersion <> 5 Then

                Result.BError = True
                Result.StringReturn =
                    "Migration 006 requires safe schema 2 or comparison schema 5."
                Return Result

            End If

            Dim SourceWS As Worksheet =
                TryGetWorksheet(WB, "Transactional DB")

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(WB, MetadataSheetName)

            If SourceWS Is Nothing OrElse MetadataWS Is Nothing Then
                Result.BError = True
                Result.StringReturn =
                    "Migration 006 requires Transactional DB and __Abovo_Metadata."
                Return Result
            End If

            Dim Validation As AbovoTransaction =
                ValidateDevelopmentSingleRanges(WB, SourceWS)

            If Validation.BError Then Return Validation

            Dim PreviousCalculationMode As WorkbookCalculationMode =
                WB.Options.CalculationMode

            Dim PreviousEngine As CalculationEngineType =
                WB.Options.CalculationEngineType

            Dim PreviousHistoryEnabled As Boolean =
                WB.History.IsEnabled

            Dim EngineChanged As Boolean = False
            Dim HistoryChanged As Boolean = False
            Dim UpdateStarted As Boolean = False
            Dim MigrationTimer As Stopwatch = Stopwatch.StartNew()

            Try


                If WB.History.IsEnabled Then
                    WB.History.IsEnabled = False
                    HistoryChanged = True
                End If

                If WB.Options.CalculationEngineType <>
                   CalculationEngineType.Recursive Then

                    WB.Options.CalculationEngineType =
                        CalculationEngineType.Recursive

                    EngineChanged = True

                End If

                WB.Options.CalculationMode =
                    WorkbookCalculationMode.Manual

                WB.BeginUpdate()
                UpdateStarted = True

                Dim ShadowWS As Worksheet =
                    TryGetWorksheet(WB, DevelopmentShadowSheetName)

                If ShadowWS IsNot Nothing Then
                    WB.Worksheets.Remove(ShadowWS)
                End If

                ShadowWS =
                    WB.Worksheets.Add(DevelopmentShadowSheetName)

                ShadowWS.VisibilityType =
                    WorksheetVisibilityType.VeryHidden

                Dim Engine As FormulaEngine =
                    WB.FormulaEngine

                Dim Culture As CultureInfo =
                    CultureInfo.GetCultureInfo("en-US")

                'Copy the TransactionDB support/header rows which are known to
                'be referenced by Development formulas (for example Q$5).
                'They stay at the identical coordinates.
                For SupportRow As Integer = 0 To 5

                    For SupportColumn As Integer = 0 To 73

                        ReconstructCellInShadow(
                            Engine,
                            Culture,
                            SourceWS,
                            SourceWS.Cells(SupportRow,
                                           SupportColumn),
                            ShadowWS,
                            ShadowWS.Cells(SupportRow,
                                           SupportColumn))

                    Next

                Next

                Dim TotalCells As Integer = 0
                Dim FormulaCells As Integer = 0
                Dim ConstantCells As Integer = 0
                Dim ValueMismatches As Integer = 0
                Dim FormulaMismatches As Integer = 0
                Dim EvaluationErrors As Integer = 0
                Dim ExamplesPrinted As Integer = 0
                Const MaximumExamples As Integer = 100

                For Each TargetName As String In DevelopmentSingleNames

                    Dim SourceDN As DefinedName =
                        TryGetDefinedName(WB, TargetName)

                    If SourceDN Is Nothing OrElse
                       SourceDN.Range Is Nothing Then

                        Throw New InvalidOperationException(
                            "Migration 006 could not resolve " &
                            TargetName & ".")

                    End If

                    Dim SourceRange As CellRange =
                        SourceDN.Range

                    Dim BlockTimer As Stopwatch =
                        Stopwatch.StartNew()

                    For SourceRow As Integer = SourceRange.TopRowIndex To SourceRange.BottomRowIndex

                        For SourceColumn As Integer = SourceRange.LeftColumnIndex To SourceRange.RightColumnIndex

                            Dim SourceCell As Cell =
                                SourceWS.Cells(SourceRow,
                                               SourceColumn)

                            Dim ShadowCell As Cell =
                                ShadowWS.Cells(SourceRow,
                                               SourceColumn)

                            ReconstructCellInShadow(
                                Engine,
                                Culture,
                                SourceWS,
                                SourceCell,
                                ShadowWS,
                                ShadowCell)

                            TotalCells += 1

                            If String.IsNullOrWhiteSpace(SourceCell.Formula) Then
                                ConstantCells += 1
                            Else
                                FormulaCells += 1
                            End If

                            Dim ValueReason As String = Nothing

                            Try

                                If Not ShadowCellMatchesSource(
                                    Engine,
                                    Culture,
                                    SourceWS,
                                    SourceCell,
                                    ShadowWS,
                                    ShadowCell,
                                    ValueReason) Then

                                    ValueMismatches += 1

                                    If ExamplesPrinted <
                                       MaximumExamples Then


                                        ExamplesPrinted += 1

                                    End If

                                End If

                            Catch ex As Exception

                                EvaluationErrors += 1

                                If ExamplesPrinted <
                                   MaximumExamples Then


                                    ExamplesPrinted += 1

                                End If

                            End Try

                            If Not String.IsNullOrWhiteSpace(
                                SourceCell.Formula) Then

                                Dim SourceSignature As String =
                                    GetR1C1Signature(
                                        Engine,
                                        Culture,
                                        SourceWS,
                                        SourceCell)

                                Dim ShadowSignature As String =
                                    GetR1C1Signature(
                                        Engine,
                                        Culture,
                                        ShadowWS,
                                        ShadowCell)

                                If Not String.Equals(
                                    SourceSignature,
                                    ShadowSignature,
                                    StringComparison.Ordinal) Then

                                    FormulaMismatches += 1

                                    If ExamplesPrinted <
                                       MaximumExamples Then


                                        ExamplesPrinted += 1

                                    End If

                                End If

                            End If

                        Next

                    Next


                Next

                MetadataWS.Cells(11, 0).Value =
                    "DevelopmentIdentifiedShadowComparison"

                MetadataWS.Cells(12, 0).Value =
                    "DevelopmentShadowSheet"

                MetadataWS.Cells(12, 1).Value =
                    DevelopmentShadowSheetName

                MetadataWS.Cells(13, 0).Value =
                    "DevelopmentShadowCoordinateMode"

                MetadataWS.Cells(13, 1).Value =
                    "SOURCE_COORDINATES"

                MetadataWS.Cells(14, 0).Value =
                    "DevelopmentShadowValueMismatches"

                MetadataWS.Cells(14, 1).Value =
                    ValueMismatches

                MetadataWS.Cells(15, 0).Value =
                    "DevelopmentShadowFormulaMismatches"

                MetadataWS.Cells(15, 1).Value =
                    FormulaMismatches

                MetadataWS.Cells(16, 0).Value =
                    "DevelopmentShadowEvaluationErrors"

                MetadataWS.Cells(16, 1).Value =
                    EvaluationErrors

                MetadataWS.Cells(17, 0).Value =
                    "Migration006AppliedUtc"

                MetadataWS.Cells(17, 1).Value =
                    DateTime.UtcNow.ToString(
                        "o",
                        CultureInfo.InvariantCulture)

                If ValueMismatches = 0 AndAlso
                   FormulaMismatches = 0 AndAlso
                   EvaluationErrors = 0 Then

                    MetadataWS.Cells(11, 1).Value =
                        "PARITY_PASS"

                Else

                    MetadataWS.Cells(11, 1).Value =
                        "PARITY_ANALYSE"

                End If

                MetadataWS.Cells(0, 1).Value = 6
                ExcelModels(ModelID).IsDirty = True


                If ValueMismatches = 0 AndAlso
                   FormulaMismatches = 0 AndAlso
                   EvaluationErrors = 0 Then


                    Result.StringReturn =
                        "Migration 006 achieved exact coordinate-preserving " &
                        "Development shadow parity. Original TransCopy_* names remain unchanged."

                Else


                    Result.StringReturn =
                        "Migration 006 created the coordinate-preserving shadow " &
                        "but parity requires analysis. Original TransCopy_* names remain unchanged."

                End If


            Catch ex As Exception

                Result.BError = True
                Result.StringReturn =
                    "Migration 006 failed: " &
                    ex.Message


            Finally

                If UpdateStarted Then
                    Try
                        WB.EndUpdate()
                    Catch
                    End Try
                End If

                Try
                    WB.Options.CalculationMode =
                        PreviousCalculationMode
                Catch
                End Try

                If EngineChanged Then
                    Try
                        WB.Options.CalculationEngineType =
                            PreviousEngine
                    Catch
                    End Try
                End If

                If HistoryChanged Then
                    Try
                        WB.History.IsEnabled =
                            PreviousHistoryEnabled
                    Catch
                    End Try
                End If

            End Try

            Return Result

        End Function

        '=====================================================================
        ' MIGRATION 003 - ABANDONED EXPERIMENTAL IN-PLACE CAPACITY MIGRATION
        '
        'FORMULA-BACKED RESERVED CAPACITY FOR DEVELOPMENT IDENTIFIED.
        '
        'This migration is deliberately Excel-native:
        '
        '   * all reserved record rows contain genuine Excel formulas/constants;
        '   * the fixed capacity footer contains the workbook's genuine footer
        '     formulas/constants;
        '   * the original TransCopy_DevptSingle_* names remain the active,
        '     contiguous Excel-facing ranges;
        '   * AbovoCap_TransCopy_DevptSingle_* names describe the physical
        '     formula-backed capacity available to Summit.
        '
        'Normal Summit Add/Delete operations can then move the logical footer and
        'resize the original name without structurally inserting/deleting rows.
        '
        'The workbook remains fully usable in Excel.  Excel/VBA can continue to
        'use the original TransCopy_* names exactly as before.
        '=====================================================================
        Private Function ApplyMigration003InstallFormulaBackedDevelopmentCapacity() As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            If GetSchemaVersion() <> 2 Then
                Result.BError = True
                Result.StringReturn =
                    "Migration 003 requires workbook schema 2."
                Return Result
            End If

            Dim TransactionWS As Worksheet =
                TryGetWorksheet(WB, "Transactional DB")

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(WB, MetadataSheetName)

            If TransactionWS Is Nothing OrElse MetadataWS Is Nothing Then
                Result.BError = True
                Result.StringReturn =
                    "Migration 003 requires Transactional DB and __Abovo_Metadata."
                Return Result
            End If

            Dim ValidationResult As AbovoTransaction =
                ValidateDevelopmentSingleRanges(WB, TransactionWS)

            If ValidationResult.BError Then Return ValidationResult

            Dim PreviousCalculationMode As WorkbookCalculationMode =
                WB.Options.CalculationMode

            Dim PreviousEngine As CalculationEngineType =
                WB.Options.CalculationEngineType

            Dim PreviousHistoryEnabled As Boolean =
                WB.History.IsEnabled

            Dim EngineChanged As Boolean = False
            Dim HistoryChanged As Boolean = False
            Dim UpdateStarted As Boolean = False
            Dim MigrationTimer As Stopwatch = Stopwatch.StartNew()

            Try


                If WB.History.IsEnabled Then
                    WB.History.IsEnabled = False
                    HistoryChanged = True
                End If

                If WB.Options.CalculationEngineType <>
                   CalculationEngineType.Recursive Then

                    WB.Options.CalculationEngineType =
                        CalculationEngineType.Recursive

                    EngineChanged = True

                End If

                WB.Options.CalculationMode =
                    WorkbookCalculationMode.Manual

                WB.BeginUpdate()
                UpdateStarted = True

                'Resolve and sort all active blocks BEFORE changing the sheet.
                'Processing bottom-to-top prevents one reserve insertion from
                'invalidating the row coordinates of blocks still to be handled.
                Dim Blocks As New List(Of DefinedName)

                For Each TargetName As String In DevelopmentSingleNames

                    Dim DN As DefinedName =
                        TryGetDefinedName(WB, TargetName)

                    If DN Is Nothing OrElse DN.Range Is Nothing Then
                        Throw New InvalidOperationException(
                            "Migration 003 could not resolve " &
                            TargetName & ".")
                    End If

                    Blocks.Add(DN)

                Next

                Blocks.Sort(
                    Function(A As DefinedName,
                             B As DefinedName) As Integer

                        Return B.Range.TopRowIndex.
                            CompareTo(A.Range.TopRowIndex)

                    End Function)

                For Each DN As DefinedName In Blocks

                    Dim TargetName As String = DN.Name
                    Dim ActiveRange As CellRange = DN.Range
                    Dim CurrentDataRows As Integer =
                        ActiveRange.RowCount - 1

                    If CurrentDataRows <= 0 Then

                        Throw New InvalidOperationException(
                            TargetName &
                            " does not contain a data row and footer.")

                    End If

                    If CurrentDataRows >
                       DevelopmentIdentifiedCapacity Then

                        Throw New InvalidOperationException(
                            TargetName &
                            " already contains " &
                            CurrentDataRows.ToString &
                            " data rows, exceeding migration capacity " &
                            DevelopmentIdentifiedCapacity.ToString & ".")

                    End If

                    Dim CapacityName As String =
                        GetCapacityName(TargetName)

                    Dim ExistingCapacity As DefinedName =
                        TryGetDefinedName(WB, CapacityName)

                    If ExistingCapacity IsNot Nothing AndAlso
                       ExistingCapacity.Range IsNot Nothing Then


                        Continue For

                    End If

                    Dim OriginalTop As Integer =
                        ActiveRange.TopRowIndex

                    Dim OriginalBottom As Integer =
                        ActiveRange.BottomRowIndex

                    Dim OriginalLeft As Integer =
                        ActiveRange.LeftColumnIndex

                    Dim OriginalRight As Integer =
                        ActiveRange.RightColumnIndex

                    Dim LastDataRow As Integer =
                        OriginalBottom - 1

                    Dim RowsToReserve As Integer =
                        DevelopmentIdentifiedCapacity -
                        CurrentDataRows

                    Dim CapacityBottom As Integer =
                        OriginalBottom + RowsToReserve

                    Dim BlockTimer As Stopwatch =
                        Stopwatch.StartNew()

                    If RowsToReserve > 0 Then

                        Dim UsedRange As CellRange =
                            TransactionWS.GetUsedRange()

                        Dim ShiftLeft As Integer =
                            UsedRange.LeftColumnIndex

                        Dim ShiftRight As Integer =
                            UsedRange.RightColumnIndex

                        Dim InsertRange As CellRange =
                            TransactionWS.Range.FromLTRB(
                                ShiftLeft,
                                OriginalBottom,
                                ShiftRight,
                                OriginalBottom + RowsToReserve - 1)

                        Dim InsertTimer As Stopwatch =
                            Stopwatch.StartNew()

                        TransactionWS.InsertCells(
                            InsertRange,
                            InsertCellsMode.ShiftCellsDown)

                        Dim InsertElapsed As Long =
                            InsertTimer.ElapsedMilliseconds

                        'The original footer has now moved to CapacityBottom.
                        'Populate every inserted reserve row as a genuine
                        'data-template row.  CopyFrom preserves Excel formulas,
                        'formatting and relative references.
                        Dim SourceDataTemplate As CellRange =
                            TransactionWS.Range.FromLTRB(
                                OriginalLeft,
                                LastDataRow,
                                OriginalRight,
                                LastDataRow)

                        Dim ReserveDataRows As CellRange =
                            TransactionWS.Range.FromLTRB(
                                OriginalLeft,
                                OriginalBottom,
                                OriginalRight,
                                CapacityBottom - 1)

                        Dim FillTimer As Stopwatch =
                            Stopwatch.StartNew()

                        ReserveDataRows.CopyFrom(
                            SourceDataTemplate,
                            PasteSpecial.All)

                        Dim FillElapsed As Long =
                            FillTimer.ElapsedMilliseconds

                        'The shifted original footer is our canonical footer at
                        'the bottom of the capacity area.  Copy it back to the
                        'current active end so the original TransCopy_* range
                        'retains exactly its pre-migration semantics.
                        Dim CanonicalFooter As CellRange =
                            TransactionWS.Range.FromLTRB(
                                OriginalLeft,
                                CapacityBottom,
                                OriginalRight,
                                CapacityBottom)

                        Dim ActiveFooter As CellRange =
                            TransactionWS.Range.FromLTRB(
                                OriginalLeft,
                                OriginalBottom,
                                OriginalRight,
                                OriginalBottom)

                        Dim FooterTimer As Stopwatch =
                            Stopwatch.StartNew()

                        ActiveFooter.CopyFrom(
                            CanonicalFooter,
                            PasteSpecial.All)


                    Else


                    End If

                    'Structural insertion may have expanded the original name.
                    'Reset it to precisely its old logical data rows + footer.
                    DN.Range =
                        TransactionWS.Range.FromLTRB(
                            OriginalLeft,
                            OriginalTop,
                            OriginalRight,
                            OriginalBottom)

                    'Do NOT create the capacity defined name here.
                    'Later insertions higher up Transactional DB will shift this
                    'physical block.  The original TransCopy_* name follows those
                    'shifts correctly, but a newly-created capacity name may not.
                    '
                    'All AbovoCap_* names are therefore created in one final pass
                    'after every structural insertion has completed.

                    'Migration 003 is intentionally a one-time structural
                    'operation and can take more than sixty seconds across all
                    'fourteen blocks.  Visual Studio's ContextSwitchDeadlock
                    'MDA fires when an STA/COM-owning thread does not pump
                    'messages for that long.
                    '
                    'Yield only at a completely committed block boundary:
                    '   - calculation remains Manual;
                    '   - calculation engine remains Recursive;
                    '   - history remains disabled;
                    '   - the current block has its active/capacity names;
                    '   - no workbook structural edit is half-complete.
                    '
                    'EndUpdate before pumping so UI/re-entrant code can never
                    'observe a half-applied DevExpress update transaction.
                    YieldMigrationContext(WB,
                                          UpdateStarted)

                Next

                'All structural insertions are now complete.  Create/update
                'capacity names from the FINAL shifted positions of the original
                'TransCopy_* names.  This is the key invariant: active names are
                'the authoritative location; capacity metadata is derived from
                'them, never from pre-shift row coordinates.

                For Each TargetName As String In DevelopmentSingleNames

                    Dim ActiveDN As DefinedName =
                        TryGetDefinedName(WB, TargetName)

                    If ActiveDN Is Nothing OrElse
                       ActiveDN.Range Is Nothing Then

                        Throw New InvalidOperationException(
                            "Migration 003 finalisation could not resolve " &
                            TargetName & ".")

                    End If

                    Dim ActiveRange As CellRange =
                        ActiveDN.Range

                    Dim FinalCapacityRange As CellRange =
                        TransactionWS.Range.FromLTRB(
                            ActiveRange.LeftColumnIndex,
                            ActiveRange.TopRowIndex,
                            ActiveRange.RightColumnIndex,
                            ActiveRange.TopRowIndex +
                            DevelopmentIdentifiedCapacity)

                    Dim CapacityName As String =
                        GetCapacityName(TargetName)

                    Dim CapacityDN As DefinedName =
                        TryGetDefinedName(WB, CapacityName)

                    If CapacityDN Is Nothing Then

                        WB.DefinedNames.Add(
                            CapacityName,
                            "'" &
                            TransactionWS.Name.Replace("'", "''") &
                            "'!" &
                            FinalCapacityRange.GetReferenceA1())

                    Else

                        CapacityDN.Range =
                            FinalCapacityRange

                    End If


                Next

                MetadataWS.Cells(6, 0).Value =
                    "DevelopmentIdentifiedFormulaCapacity"

                MetadataWS.Cells(6, 1).Value =
                    DevelopmentIdentifiedCapacity

                MetadataWS.Cells(7, 0).Value =
                    "DevelopmentIdentifiedCapacityMode"

                MetadataWS.Cells(7, 1).Value =
                    "FORMULA_BACKED"

                MetadataWS.Cells(8, 0).Value =
                    "Migration003AppliedUtc"

                MetadataWS.Cells(8, 1).Value =
                    DateTime.UtcNow.ToString(
                        "o",
                        CultureInfo.InvariantCulture)

                MetadataWS.Cells(0, 1).Value = 3

                ExcelModels(ModelID).IsDirty = True

                Result.StringReturn =
                    "Migration 003 installed formula-backed Development Identified TransactionDB capacity."


            Catch ex As Exception

                Result.BError = True
                Result.StringReturn =
                    "Migration 003 failed: " &
                    ex.Message


            Finally

                If UpdateStarted Then

                    Try
                        WB.EndUpdate()
                    Catch
                    End Try

                End If

                Try
                    WB.Options.CalculationMode =
                        PreviousCalculationMode
                Catch
                End Try

                If EngineChanged Then

                    Try
                        WB.Options.CalculationEngineType =
                            PreviousEngine
                    Catch
                    End Try

                End If

                If HistoryChanged Then

                    Try
                        WB.History.IsEnabled =
                            PreviousHistoryEnabled
                    Catch
                    End Try

                End If

            End Try

            Return Result

        End Function

        '=====================================================================
        ' MIGRATION 004 - ABANDONED EXPERIMENTAL CAPACITY-NAME REPAIR
        '
        'NON-STRUCTURAL REPAIR FOR EARLY SCHEMA-3 CAPACITY NAMES.
        '
        'The first schema-3 implementation created AbovoCap_* names while
        'processing TransactionDB blocks bottom-to-top.  Later insertions above
        'a completed block shifted the physical cells and the authoritative
        'TransCopy_* name, but the new AbovoCap_* name did not follow that shift.
        '
        'The physical formula-backed capacity is intact.  Migration 004 simply
        're-derives every capacity range from the FINAL active TransCopy_* range:
        '
        '   capacity top    = active top
        '   capacity left   = active left
        '   capacity right  = active right
        '   capacity bottom = active top + configured data capacity
        '
        'No cells, formulas, rows or columns are changed.
        '=====================================================================
        Private Function ApplyMigration004RepairDevelopmentCapacityNames() As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            Dim TransactionWS As Worksheet =
                TryGetWorksheet(WB, "Transactional DB")

            Dim MetadataWS As Worksheet =
                TryGetWorksheet(WB, MetadataSheetName)

            If TransactionWS Is Nothing OrElse
               MetadataWS Is Nothing Then

                Result.BError = True
                Result.StringReturn =
                    "Migration 004 requires Transactional DB and __Abovo_Metadata."
                Return Result

            End If


            Dim RepairTimer As Stopwatch = Stopwatch.StartNew()

            Try

                Dim CapacityValue As Integer =
                    DevelopmentIdentifiedCapacity

                If MetadataWS.Cells(6, 1).Value.IsNumeric Then

                    CapacityValue =
                        CInt(MetadataWS.Cells(6, 1).Value.NumericValue)

                End If

                If CapacityValue <= 0 Then
                    CapacityValue =
                        DevelopmentIdentifiedCapacity
                End If

                For Each TargetName As String In DevelopmentSingleNames

                    Dim ActiveDN As DefinedName =
                        TryGetDefinedName(WB, TargetName)

                    If ActiveDN Is Nothing OrElse
                       ActiveDN.Range Is Nothing Then

                        Throw New InvalidOperationException(
                            "Migration 004 could not resolve active range " &
                            TargetName & ".")

                    End If

                    Dim ActiveRange As CellRange =
                        ActiveDN.Range

                    If ActiveRange.Worksheet Is Nothing OrElse
                       Not String.Equals(
                           ActiveRange.Worksheet.Name,
                           TransactionWS.Name,
                           StringComparison.OrdinalIgnoreCase) Then

                        Throw New InvalidOperationException(
                            TargetName &
                            " is not located on Transactional DB.")

                    End If

                    Dim CorrectCapacity As CellRange =
                        TransactionWS.Range.FromLTRB(
                            ActiveRange.LeftColumnIndex,
                            ActiveRange.TopRowIndex,
                            ActiveRange.RightColumnIndex,
                            ActiveRange.TopRowIndex +
                            CapacityValue)

                    Dim CapacityName As String =
                        GetCapacityName(TargetName)

                    Dim CapacityDN As DefinedName =
                        TryGetDefinedName(WB, CapacityName)

                    Dim PreviousReference As String =
                        "(missing)"

                    If CapacityDN IsNot Nothing AndAlso
                       CapacityDN.Range IsNot Nothing Then

                        PreviousReference =
                            CapacityDN.Range.GetReferenceA1()

                        CapacityDN.Range =
                            CorrectCapacity

                    ElseIf CapacityDN Is Nothing Then

                        WB.DefinedNames.Add(
                            CapacityName,
                            "'" &
                            TransactionWS.Name.Replace("'", "''") &
                            "'!" &
                            CorrectCapacity.GetReferenceA1())

                    Else

                        Throw New InvalidOperationException(
                            CapacityName &
                            " exists but does not resolve to a range.")

                    End If


                Next

                MetadataWS.Cells(10, 0).Value =
                    "Migration004AppliedUtc"

                MetadataWS.Cells(10, 1).Value =
                    DateTime.UtcNow.ToString(
                        "o",
                        CultureInfo.InvariantCulture)

                MetadataWS.Cells(0, 1).Value = 4

                ExcelModels(ModelID).IsDirty = True

                Dim FailureReason As String = Nothing
                Dim Valid As Boolean =
                    ValidateDevelopmentIdentifiedFormulaCapacity(
                        FailureReason,
                        True)

                If Not Valid Then

                    Throw New InvalidOperationException(
                        "Capacity names were repaired but validation still failed: " &
                        FailureReason)

                End If

                Result.StringReturn =
                    "Migration 004 repaired Development Identified capacity names without changing workbook cells."


            Catch ex As Exception

                Result.BError = True
                Result.StringReturn =
                    "Migration 004 failed: " &
                    ex.Message


            End Try

            Return Result

        End Function

        Public Function IsDevelopmentIdentifiedFormulaCapacityInstalled() As Boolean

            Dim FailureReason As String = Nothing

            Return ValidateDevelopmentIdentifiedFormulaCapacity(
                FailureReason,
                False)

        End Function

        Private Function ValidateDevelopmentIdentifiedFormulaCapacity(ByRef FailureReason As String,
                                                                      ByVal PrintDetail As Boolean) As Boolean

            FailureReason = Nothing

            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then

                FailureReason =
                    "Workbook is not available."

                If PrintDetail Then
                End If

                Return False

            End If

            Dim SchemaVersion As Integer =
                GetSchemaVersion()

            If SchemaVersion < 3 Then

                FailureReason =
                    "Workbook schema is " &
                    SchemaVersion.ToString &
                    "; schema 3 or later is required."

                If PrintDetail Then
                End If

                Return False

            End If

            Dim ExpectedCapacity As Integer =
                DevelopmentIdentifiedCapacity

            Dim AllValid As Boolean = True

            For Each TargetName As String In DevelopmentSingleNames

                Dim CapacityName As String =
                    GetCapacityName(TargetName)

                Dim ActiveDN As DefinedName =
                    TryGetDefinedName(WB, TargetName)

                Dim CapacityDN As DefinedName =
                    TryGetDefinedName(WB, CapacityName)

                If ActiveDN Is Nothing Then

                    AllValid = False
                    FailureReason =
                        "Active defined name '" &
                        TargetName &
                        "' is missing."

                    If PrintDetail Then
                    End If

                    Continue For

                End If

                If ActiveDN.Range Is Nothing Then

                    AllValid = False
                    FailureReason =
                        "Active defined name '" &
                        TargetName &
                        "' does not resolve to a range."

                    If PrintDetail Then
                    End If

                    Continue For

                End If

                If CapacityDN Is Nothing Then

                    AllValid = False
                    FailureReason =
                        "Capacity defined name '" &
                        CapacityName &
                        "' is missing."

                    If PrintDetail Then
                    End If

                    Continue For

                End If

                If CapacityDN.Range Is Nothing Then

                    AllValid = False
                    FailureReason =
                        "Capacity defined name '" &
                        CapacityName &
                        "' does not resolve to a range."

                    If PrintDetail Then
                    End If

                    Continue For

                End If

                Dim ActiveRange As CellRange =
                    ActiveDN.Range

                Dim CapacityRange As CellRange =
                    CapacityDN.Range

                Dim SameWorksheet As Boolean =
                    ActiveRange.Worksheet IsNot Nothing AndAlso
                    CapacityRange.Worksheet IsNot Nothing AndAlso
                    String.Equals(
                        ActiveRange.Worksheet.Name,
                        CapacityRange.Worksheet.Name,
                        StringComparison.OrdinalIgnoreCase) AndAlso
                    String.Equals(
                        ActiveRange.Worksheet.Name,
                        "Transactional DB",
                        StringComparison.OrdinalIgnoreCase)

                Dim SameTop As Boolean =
                    ActiveRange.TopRowIndex =
                    CapacityRange.TopRowIndex

                Dim SameLeft As Boolean =
                    ActiveRange.LeftColumnIndex =
                    CapacityRange.LeftColumnIndex

                Dim SameRight As Boolean =
                    ActiveRange.RightColumnIndex =
                    CapacityRange.RightColumnIndex

                Dim CapacityDataRows As Integer =
                    CapacityRange.RowCount - 1

                Dim CapacityLargeEnough As Boolean =
                    CapacityDataRows >= ExpectedCapacity

                Dim ActiveWithinCapacity As Boolean =
                    ActiveRange.BottomRowIndex <=
                    CapacityRange.BottomRowIndex

                Dim BlockValid As Boolean =
                    SameWorksheet AndAlso
                    SameTop AndAlso
                    SameLeft AndAlso
                    SameRight AndAlso
                    CapacityLargeEnough AndAlso
                    ActiveWithinCapacity

                If PrintDetail Then


                End If

                If Not BlockValid Then

                    AllValid = False

                    If Not SameWorksheet Then
                        FailureReason =
                            TargetName &
                            " active/capacity worksheet mismatch."
                    ElseIf Not SameTop Then
                        FailureReason =
                            TargetName &
                            " active/capacity top-row mismatch."
                    ElseIf Not SameLeft OrElse Not SameRight Then
                        FailureReason =
                            TargetName &
                            " active/capacity column mismatch."
                    ElseIf Not CapacityLargeEnough Then
                        FailureReason =
                            TargetName &
                            " capacity is " &
                            CapacityDataRows.ToString &
                            " data rows; expected at least " &
                            ExpectedCapacity.ToString & "."
                    ElseIf Not ActiveWithinCapacity Then
                        FailureReason =
                            TargetName &
                            " active range extends beyond capacity."
                    End If

                End If

            Next

            If PrintDetail Then

                If AllValid Then
                Else
                End If

            End If

            Return AllValid

        End Function

        Public Sub DiagnoseDevelopmentIdentifiedFormulaCapacity()


            Dim FailureReason As String = Nothing
            Dim Valid As Boolean =
                ValidateDevelopmentIdentifiedFormulaCapacity(
                    FailureReason,
                    True)


            If Not Valid Then
            End If


        End Sub

        Public Function EnsureDevelopmentIdentifiedCapacity(ByVal RequiredDataRows As Integer) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            If Not IsDevelopmentIdentifiedFormulaCapacityInstalled() Then
                Result.BError = True
                Result.StringReturn =
                    "Formula-backed Development Identified capacity is not installed."
                Return Result
            End If

            Dim CurrentCapacity As Integer = Integer.MaxValue

            For Each TargetName As String In DevelopmentSingleNames

                Dim CapacityDN As DefinedName =
                    TryGetDefinedName(
                        WB,
                        GetCapacityName(TargetName))

                If CapacityDN Is Nothing OrElse
                   CapacityDN.Range Is Nothing Then

                    Result.BError = True
                    Result.StringReturn =
                        "Capacity metadata is missing for " &
                        TargetName & "."
                    Return Result

                End If

                CurrentCapacity =
                    Math.Min(CurrentCapacity,
                             CapacityDN.Range.RowCount - 1)

            Next

            If RequiredDataRows <= CurrentCapacity Then

                Result.StringReturn =
                    "Existing Development Identified capacity " &
                    CurrentCapacity.ToString &
                    " is sufficient."

                Return Result

            End If

            Dim NewCapacity As Integer =
                CInt(Math.Ceiling(
                    RequiredDataRows / 50.0R) * 50.0R)

            Dim TransactionWS As Worksheet =
                TryGetWorksheet(WB, "Transactional DB")

            If TransactionWS Is Nothing Then
                Result.BError = True
                Result.StringReturn =
                    "Transactional DB worksheet is unavailable."
                Return Result
            End If

            Dim PreviousCalculationMode As WorkbookCalculationMode =
                WB.Options.CalculationMode

            Dim PreviousEngine As CalculationEngineType =
                WB.Options.CalculationEngineType

            Dim PreviousHistoryEnabled As Boolean =
                WB.History.IsEnabled

            Dim EngineChanged As Boolean = False
            Dim HistoryChanged As Boolean = False
            Dim UpdateStarted As Boolean = False

            Try


                If WB.History.IsEnabled Then
                    WB.History.IsEnabled = False
                    HistoryChanged = True
                End If

                If WB.Options.CalculationEngineType <>
                   CalculationEngineType.Recursive Then

                    WB.Options.CalculationEngineType =
                        CalculationEngineType.Recursive

                    EngineChanged = True

                End If

                WB.Options.CalculationMode =
                    WorkbookCalculationMode.Manual

                WB.BeginUpdate()
                UpdateStarted = True

                Dim CapacityNames As New List(Of DefinedName)

                For Each TargetName As String In DevelopmentSingleNames

                    CapacityNames.Add(
                        TryGetDefinedName(
                            WB,
                            GetCapacityName(TargetName)))

                Next

                CapacityNames.Sort(
                    Function(A As DefinedName,
                             B As DefinedName) As Integer

                        Return B.Range.TopRowIndex.
                            CompareTo(A.Range.TopRowIndex)

                    End Function)

                For Each CapacityDN As DefinedName In CapacityNames

                    Dim CapacityRange As CellRange =
                        CapacityDN.Range

                    Dim ExistingCapacity As Integer =
                        CapacityRange.RowCount - 1

                    Dim RowsToAdd As Integer =
                        NewCapacity -
                        ExistingCapacity

                    If RowsToAdd <= 0 Then Continue For

                    Dim CanonicalDataRow As CellRange =
                        TransactionWS.Range.FromLTRB(
                            CapacityRange.LeftColumnIndex,
                            CapacityRange.TopRowIndex,
                            CapacityRange.RightColumnIndex,
                            CapacityRange.TopRowIndex)

                    Dim OldCanonicalFooterRow As Integer =
                        CapacityRange.BottomRowIndex

                    Dim UsedRange As CellRange =
                        TransactionWS.GetUsedRange()

                    Dim InsertRange As CellRange =
                        TransactionWS.Range.FromLTRB(
                            UsedRange.LeftColumnIndex,
                            OldCanonicalFooterRow,
                            UsedRange.RightColumnIndex,
                            OldCanonicalFooterRow + RowsToAdd - 1)

                    Dim T As Stopwatch =
                        Stopwatch.StartNew()

                    TransactionWS.InsertCells(
                        InsertRange,
                        InsertCellsMode.ShiftCellsDown)

                    Dim NewCanonicalFooterRow As Integer =
                        OldCanonicalFooterRow +
                        RowsToAdd

                    Dim NewReserveRows As CellRange =
                        TransactionWS.Range.FromLTRB(
                            CapacityRange.LeftColumnIndex,
                            OldCanonicalFooterRow,
                            CapacityRange.RightColumnIndex,
                            NewCanonicalFooterRow - 1)

                    NewReserveRows.CopyFrom(
                        CanonicalDataRow,
                        PasteSpecial.All)

                    CapacityDN.Range =
                        TransactionWS.Range.FromLTRB(
                            CapacityRange.LeftColumnIndex,
                            CapacityRange.TopRowIndex,
                            CapacityRange.RightColumnIndex,
                            NewCanonicalFooterRow)


                Next

                Dim MetadataWS As Worksheet =
                    TryGetWorksheet(WB, MetadataSheetName)

                If MetadataWS IsNot Nothing Then
                    MetadataWS.Cells(6, 1).Value = NewCapacity
                    MetadataWS.Cells(9, 0).Value =
                        "CapacityExpandedUtc"
                    MetadataWS.Cells(9, 1).Value =
                        DateTime.UtcNow.ToString(
                            "o",
                            CultureInfo.InvariantCulture)
                End If

                ExcelModels(ModelID).IsDirty = True

                Result.StringReturn =
                    "Development Identified formula capacity expanded to " &
                    NewCapacity.ToString & "."

            Catch ex As Exception

                Result.BError = True
                Result.StringReturn =
                    "Development Identified formula capacity expansion failed: " &
                    ex.Message

            Finally

                If UpdateStarted Then

                    Try
                        WB.EndUpdate()
                    Catch
                    End Try

                End If

                Try
                    WB.Options.CalculationMode =
                        PreviousCalculationMode
                Catch
                End Try

                If EngineChanged Then

                    Try
                        WB.Options.CalculationEngineType =
                            PreviousEngine
                    Catch
                    End Try

                End If

                If HistoryChanged Then

                    Try
                        WB.History.IsEnabled =
                            PreviousHistoryEnabled
                    Catch
                    End Try

                End If

            End Try

            Return Result

        End Function

        Private Function ValidateDevelopmentSingleRanges(ByVal WB As IWorkbook,
                                                         ByVal TransactionWS As Worksheet) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}

            For Each TargetName As String In DevelopmentSingleNames

                Dim DN As DefinedName = TryGetDefinedName(WB, TargetName)

                If DN Is Nothing OrElse
                   DN.Range Is Nothing OrElse
                   DN.Range.Worksheet Is Nothing OrElse
                   Not String.Equals(DN.Range.Worksheet.Name,
                                     TransactionWS.Name,
                                     StringComparison.OrdinalIgnoreCase) Then

                    Result.BError = True

                    Result.StringReturn =
                        "Development Identified TransactionDB precondition failed: " &
                        TargetName &
                        " is missing or is not on Transactional DB."

                    Return Result

                End If

                If DN.Range.RowCount < 2 Then

                    Result.BError = True

                    Result.StringReturn =
                        "Development Identified TransactionDB precondition failed: " &
                        TargetName &
                        " does not contain at least one data row and a footer."

                    Return Result

                End If

            Next

            Return Result

        End Function

        Private Sub CaptureDevelopmentIdentifiedTemplates(ByVal WB As IWorkbook,
                                                          ByVal TransactionWS As Worksheet,
                                                          ByVal TemplateWS As Worksheet)

            Dim Engine As FormulaEngine = WB.FormulaEngine
            Dim Culture As CultureInfo =
                CultureInfo.GetCultureInfo("en-US")

            TemplateWS.Cells(0, 0).Value =
                "Abovo Development Identified TransactionDB Templates"

            TemplateWS.Cells(0, 1).Value =
                "R1C1 formula text + typed constant values"

            TemplateWS.Cells(1, 0).Value =
                "Captured from workbook; do not hand-edit."

            For BlockIndex As Integer = 0 To DevelopmentSingleNames.Length - 1

                Dim TargetName As String =
                    DevelopmentSingleNames(BlockIndex)

                Dim DN As DefinedName =
                    WB.DefinedNames.GetDefinedName(TargetName)

                Dim Block As CellRange = DN.Range

                Dim DataTemplateRow As Integer =
                    Block.BottomRowIndex - 1

                Dim FooterTemplateRow As Integer =
                    Block.BottomRowIndex

                Dim BaseRow As Integer =
                    2 + (BlockIndex * 5)

                Dim ValueRow As Integer =
                    BaseRow + 1

                Dim FormulaRow As Integer =
                    BaseRow + 2

                Dim FooterValueRow As Integer =
                    BaseRow + 3

                Dim FooterFormulaRow As Integer =
                    BaseRow + 4

                TemplateWS.Cells(BaseRow, 0).Value =
                    TargetName

                TemplateWS.Cells(BaseRow, 1).Value =
                    Block.ColumnCount

                For Offset As Integer = 0 To Block.ColumnCount - 1

                    Dim ColumnIndex As Integer =
                        Block.LeftColumnIndex + Offset

                    Dim DataCell As Cell =
                        TransactionWS.Cells(DataTemplateRow,
                                            ColumnIndex)

                    Dim FooterCell As Cell =
                        TransactionWS.Cells(FooterTemplateRow,
                                            ColumnIndex)

                    TemplateWS.Cells(ValueRow, Offset).Value =
                        DataCell.Value

                    TemplateWS.Cells(FooterValueRow, Offset).Value =
                        FooterCell.Value

                    TemplateWS.Cells(FormulaRow, Offset).Value =
                        EncodeFormulaTemplate(
                            ConvertFormulaToR1C1Text(Engine,
                                                    DataCell,
                                                    TransactionWS,
                                                    Culture))

                    TemplateWS.Cells(FooterFormulaRow, Offset).Value =
                        EncodeFormulaTemplate(
                            ConvertFormulaToR1C1Text(Engine,
                                                    FooterCell,
                                                    TransactionWS,
                                                    Culture))

                Next

            Next

        End Sub

        Private Function EncodeFormulaTemplate(ByVal R1C1Formula As String) As String

            If String.IsNullOrWhiteSpace(R1C1Formula) Then
                Return String.Empty
            End If

            Return FormulaTemplatePrefix & R1C1Formula

        End Function

        Private Function ConvertFormulaToR1C1Text(ByVal Engine As FormulaEngine,
                                                   ByVal SourceCell As Cell,
                                                   ByVal WS As Worksheet,
                                                   ByVal Culture As CultureInfo) As String

            If SourceCell Is Nothing OrElse
               String.IsNullOrWhiteSpace(SourceCell.Formula) Then

                Return String.Empty

            End If

            Dim A1Context As New ExpressionContext(SourceCell.ColumnIndex,
                                                   SourceCell.RowIndex,
                                                   WS,
                                                   Culture,
                                                   ReferenceStyle.A1,
                                                   ExpressionStyle.Normal)

            Dim R1C1Context As New ExpressionContext(SourceCell.ColumnIndex,
                                                     SourceCell.RowIndex,
                                                     WS,
                                                     Culture,
                                                     ReferenceStyle.R1C1,
                                                     ExpressionStyle.Normal)

            Dim Parsed As ParsedExpression =
                Engine.Parse(SourceCell.Formula,
                             A1Context)

            Return Parsed.ToString(R1C1Context)

        End Function

        Private Sub YieldMigrationContext(ByVal WB As IWorkbook,
                                          ByRef UpdateStarted As Boolean)

            If WB Is Nothing Then Return

            Try

                If UpdateStarted Then
                    WB.EndUpdate()
                    UpdateStarted = False
                End If

                'Fully qualify this call because the Summit project also has
                'an ApplicationConfiguration type.  This is deliberately used
                'only by the one-time workbook migration, never by normal
                'structural Add/Delete processing.
                System.Windows.Forms.Application.DoEvents()

                WB.BeginUpdate()
                UpdateStarted = True

            Catch ex As Exception


                'If EndUpdate succeeded but BeginUpdate did not, leave
                'UpdateStarted False so the outer Finally does not issue an
                'unmatched EndUpdate.  Treat inability to resume the update
                'transaction as fatal rather than continuing structural work.
                If Not UpdateStarted Then
                    Throw
                End If

            End Try

        End Sub

        Private Function TryGetWorksheet(ByVal WB As IWorkbook,
                                         ByVal SheetName As String) As Worksheet

            If WB Is Nothing OrElse
               String.IsNullOrWhiteSpace(SheetName) Then Return Nothing

            For Each WS As Worksheet In WB.Worksheets

                If String.Equals(WS.Name,
                                 SheetName,
                                 StringComparison.OrdinalIgnoreCase) Then

                    Return WS

                End If

            Next

            Return Nothing

        End Function

        Private Function TryGetDefinedName(ByVal WB As IWorkbook,
                                           ByVal Name As String) As DefinedName

            If WB Is Nothing OrElse
               String.IsNullOrWhiteSpace(Name) Then Return Nothing

            For Each DN As DefinedName In WB.DefinedNames

                If DN IsNot Nothing AndAlso
                   String.Equals(DN.Name,
                                 Name,
                                 StringComparison.OrdinalIgnoreCase) Then

                    Return DN

                End If

            Next

            Return Nothing

        End Function

        Private Function GetOrCreateVeryHiddenWorksheet(ByVal WB As IWorkbook,
                                                         ByVal SheetName As String) As Worksheet

            Dim WS As Worksheet =
                TryGetWorksheet(WB, SheetName)

            If WS Is Nothing Then
                WS = WB.Worksheets.Add(SheetName)
            End If

            WS.VisibilityType =
                WorksheetVisibilityType.VeryHidden

            Return WS

        End Function

        'Retained for source compatibility with the schema-1 experimental code.
        Public Shared Function GetCapacityName(ByVal TargetNamedRange As String) As String

            Return "AbovoCap_" & TargetNamedRange

        End Function

        Public Shared Function GetMetadataSheetName() As String

            Return MetadataSheetName

        End Function

        Public Shared Function GetTemplateSheetName() As String

            Return TemplateSheetName

        End Function

        Public Shared Function GetDevelopmentShadowSheetName() As String

            Return DevelopmentShadowSheetName

        End Function

    End Class

End Namespace
