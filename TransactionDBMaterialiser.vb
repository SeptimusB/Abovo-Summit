Imports System
Imports System.Diagnostics
Imports System.Globalization
Imports Abovo.FileManager
Imports Abovo.AbovoAppCls
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas

Namespace Abovo

    Public Class TransactionDBMaterialiser

        Private ReadOnly ModelID As Integer
        Private Const FormulaTemplatePrefix As String = "ABOVO_R1C1_FORMULA|"

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

        'IMPORTANT:
        '
        'The value-only materialiser is intentionally NOT a live synchronisation
        'path in schema 2.  Persisted XLSB files must remain independently usable
        'in Microsoft Excel, therefore the active TransactionDB ranges must retain
        'their Excel formulas.
        '
        'TransactionalDBSynchroniser already checks this property before choosing
        'the value-only path.  Returning False keeps the proven generic structural
        'synchroniser active while still allowing us to retain/use template metadata
        'for diagnostics and future session-only acceleration.
        Public Function IsDevelopmentIdentifiedActive() As Boolean

            If ExcelModels Is Nothing OrElse
               ModelID < 0 OrElse
               ModelID >= ExcelModels.Length OrElse
               ExcelModels(ModelID) Is Nothing OrElse
               ExcelModels(ModelID).WorkbookMigrations Is Nothing Then

                Return False

            End If

            Return ExcelModels(ModelID).WorkbookMigrations.
                IsDevelopmentIdentifiedProductionMirrorInstalled()

        End Function

        Public Function IsDevelopmentIdentifiedTemplateCacheAvailable() As Boolean

            If ExcelModels Is Nothing OrElse
               ModelID < 0 OrElse
               ModelID >= ExcelModels.Length OrElse
               ExcelModels(ModelID) Is Nothing OrElse
               ExcelModels(ModelID).WorkbookMigrations Is Nothing Then

                Return False

            End If

            Return ExcelModels(ModelID).WorkbookMigrations.
                IsDevelopmentIdentifiedMaterialiserInstalled()

        End Function

        'Retained for API compatibility with TransactionalDBSynchroniser.
        '
        'Live value-only replacement is disabled.  A future accelerated session
        'mode may use the cached templates, but only after Summit has a guaranteed
        'restore-formulas boundary before every XLSB save/export.
        '=====================================================================
        ' FORMULA-BACKED DEVELOPMENT IDENTIFIED CAPACITY REFRESH
        '
        'No TransactionDB structural insertion/deletion is performed here.
        '
        'For each block:
        '   1. convert the old active footer back into a normal data row;
        '   2. convert the new active end row into a footer;
        '   3. resize the original TransCopy_* name.
        '
        'Canonical source rows live inside the reserved capacity:
        '   capacity first row  = ordinary data-row template
        '   capacity last row   = footer template
        '
        'CopyFrom is used so DevExpress adjusts relative Excel references in the
        'same way Excel would.  The persisted workbook therefore remains fully
        'formula-backed and independently usable in Excel.
        '=====================================================================
        Public Function RefreshDevelopmentIdentified(Optional ByVal Force As Boolean = False) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            If Not Force AndAlso
               Not IsDevelopmentIdentifiedActive() Then

                Result.BError = True
                Result.StringReturn =
                    "Development Identified production mirrors are not active."
                Return Result

            End If

            Dim SourceDN As DefinedName =
                TryGetDefinedName(WB, "HouseTypeInID")

            If SourceDN Is Nothing OrElse
               SourceDN.Range Is Nothing Then

                Result.BError = True
                Result.StringReturn =
                    "HouseTypeInID could not be resolved."
                Return Result

            End If

            Dim RequiredRows As Integer =
                SourceDN.Range.ColumnCount

            Dim RequiredDataRows As Integer =
                RequiredRows - 1

            If RequiredDataRows <= 0 Then

                Result.BError = True
                Result.StringReturn =
                    "HouseTypeInID resolved to an invalid Development Identified record count."
                Return Result

            End If

            Dim RefreshTimer As Stopwatch =
                Stopwatch.StartNew()

            Dim PreviousCalculationMode As WorkbookCalculationMode =
                WB.Options.CalculationMode

            Dim PreviousHistoryEnabled As Boolean =
                WB.History.IsEnabled

            Dim UpdateStarted As Boolean = False

            Try

                WB.Options.CalculationMode =
                    WorkbookCalculationMode.Manual

                If WB.History.IsEnabled Then
                    WB.History.IsEnabled = False
                End If

                WB.BeginUpdate()
                UpdateStarted = True

                For Each TargetName As String In DevelopmentSingleNames

                    Dim TargetDN As DefinedName =
                        TryGetDefinedName(WB, TargetName)

                    If TargetDN Is Nothing OrElse
                       TargetDN.Range Is Nothing Then

                        Throw New InvalidOperationException(
                            "Development mirror '" &
                            TargetName &
                            "' could not be resolved.")

                    End If

                    Dim ActiveRange As CellRange =
                        TargetDN.Range

                    Dim MirrorWS As Worksheet =
                        ActiveRange.Worksheet

                    Dim CurrentDataRows As Integer =
                        ActiveRange.RowCount - 1

                    Dim DataTop As Integer =
                        ActiveRange.TopRowIndex

                    Dim OldFooterRow As Integer =
                        ActiveRange.BottomRowIndex

                    Dim NewFooterRow As Integer =
                        DataTop + RequiredDataRows

                    Dim BlockTimer As Stopwatch =
                        Stopwatch.StartNew()

                    'DevExpress CellRange.CopyFrom automatically adjusts relative
                    'formula references for the destination.  Moving the footer as
                    'a whole row avoids per-cell FormulaEngine parsing.
                    If NewFooterRow <> OldFooterRow Then

                        Dim OldFooterRange As CellRange =
                            MirrorWS.Range.FromLTRB(
                                ActiveRange.LeftColumnIndex,
                                OldFooterRow,
                                ActiveRange.RightColumnIndex,
                                OldFooterRow)

                        Dim NewFooterRange As CellRange =
                            MirrorWS.Range.FromLTRB(
                                ActiveRange.LeftColumnIndex,
                                NewFooterRow,
                                ActiveRange.RightColumnIndex,
                                NewFooterRow)

                        NewFooterRange.CopyFrom(
                            OldFooterRange,
                            PasteSpecial.All)

                    End If

                    If RequiredDataRows > CurrentDataRows Then

                        'Schema 8 proved every Development A:N data column is one
                        'of FormulaPattern, InvariantConstant or Blank.  There are
                        'no varying numeric constants.  Therefore the final current
                        'data row is a complete safe template for future records.
                        '
                        'A one-row source copied into a taller destination is
                        'repeated by DevExpress to fill the destination, with
                        'relative formula references adjusted for every row.
                        Dim TemplateDataRow As Integer =
                            OldFooterRow - 1

                        Dim TemplateRange As CellRange =
                            MirrorWS.Range.FromLTRB(
                                ActiveRange.LeftColumnIndex,
                                TemplateDataRow,
                                ActiveRange.RightColumnIndex,
                                TemplateDataRow)

                        Dim NewDataRange As CellRange =
                            MirrorWS.Range.FromLTRB(
                                ActiveRange.LeftColumnIndex,
                                OldFooterRow,
                                ActiveRange.RightColumnIndex,
                                NewFooterRow - 1)

                        NewDataRange.CopyFrom(
                            TemplateRange,
                            PasteSpecial.All)

                    ElseIf RequiredDataRows < CurrentDataRows Then

                        'The new footer has already been copied above.  Clear all
                        'cells that are no longer part of the active named range in
                        'one range operation.
                        If NewFooterRow < OldFooterRow Then

                            Dim InactiveRange As CellRange =
                                MirrorWS.Range.FromLTRB(
                                    ActiveRange.LeftColumnIndex,
                                    NewFooterRow + 1,
                                    ActiveRange.RightColumnIndex,
                                    OldFooterRow)

                            InactiveRange.ClearContents()

                        End If

                    End If

                    TargetDN.Range =
                        MirrorWS.Range.FromLTRB(
                            ActiveRange.LeftColumnIndex,
                            DataTop,
                            ActiveRange.RightColumnIndex,
                            NewFooterRow)


                Next

                WB.EndUpdate()
                UpdateStarted = False

                If ExcelModels(ModelID).WorkbookMigrations IsNot Nothing Then

                    ExcelModels(ModelID).WorkbookMigrations.
                        UpdateDevelopmentRoundTripSignatures()

                End If

                ExcelModels(ModelID).IsDirty = True

                Result.StringReturn =
                    "Development Identified production mirrors refreshed without TransactionDB structural changes."


            Catch ex As Exception

                Result.BError = True
                Result.StringReturn =
                    "Development Identified production mirror refresh failed: " &
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

                Try
                    WB.History.IsEnabled =
                        PreviousHistoryEnabled
                Catch
                End Try

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

        '=====================================================================
        ' RESTORE DEVELOPMENT IDENTIFIED FORMULA STATE
        '
        'Recreates the live TransCopy_DevptSingle_A:N formulas from the R1C1
        'templates captured from the workbook itself.
        '
        'This is used by workbook migration 002 to repair the experimental
        'schema-1 value-only migration, and is also the method that a future
        'Summit-only acceleration mode must call before any XLSB save/export.
        '=====================================================================
        Public Function RestoreDevelopmentIdentifiedFormulaState(Optional ByVal Force As Boolean = False) As AbovoTransaction

            Dim Result As New AbovoTransaction With {.BError = False}
            Dim WB As IWorkbook = GetWorkbook()

            If WB Is Nothing Then
                Result.BError = True
                Result.StringReturn = "Workbook is not available."
                Return Result
            End If

            If Not Force AndAlso
               Not IsDevelopmentIdentifiedTemplateCacheAvailable() Then

                Result.BError = True
                Result.StringReturn =
                    "Development Identified formula template cache is not installed."

                Return Result

            End If

            Dim TransactionWS As Worksheet =
                TryGetWorksheet(WB, "Transactional DB")

            Dim TemplateWS As Worksheet =
                TryGetWorksheet(
                    WB,
                    WorkbookMigrationManager.GetTemplateSheetName())

            If TransactionWS Is Nothing OrElse TemplateWS Is Nothing Then

                Result.BError = True
                Result.StringReturn =
                    "Transactional DB or Abovo template worksheet is missing."

                Return Result

            End If

            Dim PreviousCalculationMode As WorkbookCalculationMode =
                WB.Options.CalculationMode

            Dim PreviousCalculationEngineType As CalculationEngineType =
                WB.Options.CalculationEngineType

            Dim PreviousHistoryEnabled As Boolean =
                WB.History.IsEnabled

            Dim CalculationEngineChanged As Boolean = False
            Dim HistoryChanged As Boolean = False
            Dim UpdateStarted As Boolean = False

            Dim RestoreTimer As Stopwatch = Stopwatch.StartNew()
            Dim FormulaCellsRestored As Integer = 0
            Dim ConstantCellsRestored As Integer = 0

            Try

                If WB.History.IsEnabled Then
                    WB.History.IsEnabled = False
                    HistoryChanged = True
                End If

                If WB.Options.CalculationEngineType <>
                   CalculationEngineType.Recursive Then

                    WB.Options.CalculationEngineType =
                        CalculationEngineType.Recursive

                    CalculationEngineChanged = True

                End If

                WB.Options.CalculationMode =
                    WorkbookCalculationMode.Manual

                WB.BeginUpdate()
                UpdateStarted = True

                Dim Engine As FormulaEngine =
                    WB.FormulaEngine

                Dim Culture As CultureInfo =
                    CultureInfo.GetCultureInfo("en-US")

                For BlockIndex As Integer = 0 To DevelopmentSingleNames.Length - 1

                    Dim TargetName As String =
                        DevelopmentSingleNames(BlockIndex)

                    Dim TargetDN As DefinedName =
                        TryGetDefinedName(WB, TargetName)

                    If TargetDN Is Nothing OrElse
                       TargetDN.Range Is Nothing Then

                        Throw New InvalidOperationException(
                            "Development Identified range '" &
                            TargetName &
                            "' could not be resolved.")

                    End If

                    Dim ActiveRange As CellRange =
                        TargetDN.Range

                    If ActiveRange.RowCount < 2 Then

                        Throw New InvalidOperationException(
                            "Development Identified range '" &
                            TargetName &
                            "' does not contain a data row and footer.")

                    End If

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

                    Dim LastDataRow As Integer =
                        ActiveRange.BottomRowIndex - 1

                    Dim BlockTimer As Stopwatch =
                        Stopwatch.StartNew()

                    For TargetRow As Integer =
                        ActiveRange.TopRowIndex To LastDataRow

                        For ColumnOffset As Integer =
                            0 To ActiveRange.ColumnCount - 1

                            Dim TargetColumn As Integer =
                                ActiveRange.LeftColumnIndex +
                                ColumnOffset

                            Dim TargetCell As Cell =
                                TransactionWS.Cells(TargetRow,
                                                    TargetColumn)

                            Dim R1C1Formula As String =
                                GetTemplateFormula(TemplateWS,
                                                   FormulaRow,
                                                   ColumnOffset)

                            If String.IsNullOrWhiteSpace(R1C1Formula) Then

                                TargetCell.Value =
                                    TemplateWS.Cells(ValueRow,
                                                     ColumnOffset).Value

                                ConstantCellsRestored += 1

                            Else

                                TargetCell.Formula =
                                    ConvertR1C1TemplateToA1Formula(
                                        Engine,
                                        R1C1Formula,
                                        TargetColumn,
                                        TargetRow,
                                        TransactionWS,
                                        Culture)

                                FormulaCellsRestored += 1

                            End If

                        Next

                    Next

                    Dim FooterRow As Integer =
                        ActiveRange.BottomRowIndex

                    For ColumnOffset As Integer =
                        0 To ActiveRange.ColumnCount - 1

                        Dim TargetColumn As Integer =
                            ActiveRange.LeftColumnIndex +
                            ColumnOffset

                        Dim FooterCell As Cell =
                            TransactionWS.Cells(FooterRow,
                                                TargetColumn)

                        Dim R1C1Formula As String =
                            GetTemplateFormula(TemplateWS,
                                               FooterFormulaRow,
                                               ColumnOffset)

                        If String.IsNullOrWhiteSpace(R1C1Formula) Then

                            FooterCell.Value =
                                TemplateWS.Cells(FooterValueRow,
                                                 ColumnOffset).Value

                            ConstantCellsRestored += 1

                        Else

                            FooterCell.Formula =
                                ConvertR1C1TemplateToA1Formula(
                                    Engine,
                                    R1C1Formula,
                                    TargetColumn,
                                    FooterRow,
                                    TransactionWS,
                                    Culture)

                            FormulaCellsRestored += 1

                        End If

                    Next


                Next

                Result.StringReturn =
                    "Development Identified TransactionDB formula state restored successfully."


            Catch ex As Exception

                Result.BError = True
                Result.StringReturn =
                    "Development Identified formula restoration failed: " &
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

                If CalculationEngineChanged Then

                    Try
                        WB.Options.CalculationEngineType =
                            PreviousCalculationEngineType
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

        'Persistence guard retained for compatibility.  In schema 3 the
        'runtime itself remains formula-backed, so this is normally a no-op
        'repair/verification boundary rather than a value-to-formula conversion.
        Public Function PrepareWorkbookForExcelPersistence() As AbovoTransaction

            If Not IsDevelopmentIdentifiedTemplateCacheAvailable() Then

                Dim NothingToDo As New AbovoTransaction With {
                    .BError = False,
                    .StringReturn =
                        "No Development Identified template cache is installed; no formula restoration was required."
                }

                Return NothingToDo

            End If

            Return RestoreDevelopmentIdentifiedFormulaState(True)

        End Function

        Private Function ConvertR1C1TemplateToA1Formula(ByVal Engine As FormulaEngine,
                                                        ByVal R1C1Formula As String,
                                                        ByVal ColumnIndex As Integer,
                                                        ByVal RowIndex As Integer,
                                                        ByVal WS As Worksheet,
                                                        ByVal Culture As CultureInfo) As String

            Dim R1C1Context As New ExpressionContext(ColumnIndex,
                                                     RowIndex,
                                                     WS,
                                                     Culture,
                                                     ReferenceStyle.R1C1,
                                                     ExpressionStyle.Normal)

            Dim A1Context As New ExpressionContext(ColumnIndex,
                                                   RowIndex,
                                                   WS,
                                                   Culture,
                                                   ReferenceStyle.A1,
                                                   ExpressionStyle.Normal)

            Dim Parsed As ParsedExpression =
                Engine.Parse(R1C1Formula,
                             R1C1Context)

            Dim A1Formula As String =
                Parsed.ToString(A1Context)

            If Not A1Formula.StartsWith("=", StringComparison.Ordinal) Then
                A1Formula = "=" & A1Formula
            End If

            Return A1Formula

        End Function

        Private Function GetTemplateFormula(ByVal TemplateWS As Worksheet,
                                            ByVal RowIndex As Integer,
                                            ByVal ColumnIndex As Integer) As String

            Dim Value As CellValue =
                TemplateWS.Cells(RowIndex,
                                 ColumnIndex).Value

            If Not Value.IsText Then Return String.Empty

            Dim Stored As String = Value.TextValue

            If String.IsNullOrWhiteSpace(Stored) Then
                Return String.Empty
            End If

            If Stored.StartsWith(FormulaTemplatePrefix,
                                 StringComparison.Ordinal) Then

                Return Stored.Substring(FormulaTemplatePrefix.Length)

            End If

            'Compatibility with the experimental schema-1 cache.  When a raw
            'R1C1 formula began with a quoted worksheet name, storing it as a
            'plain cell string caused Excel/DevExpress to consume the opening
            'apostrophe as the text-prefix character:
            '
            '   'FFR Key Defn'!R128C2
            ' became
            '   FFR Key Defn'!R128C2
            '
            'Only repair this very specific legacy shape.
            If Stored.IndexOf("'!",
                              StringComparison.Ordinal) > 0 AndAlso
               Not Stored.StartsWith("'", StringComparison.Ordinal) AndAlso
               Not Stored.StartsWith("=", StringComparison.Ordinal) AndAlso
               Stored.IndexOf("("c) < 0 Then

                Stored = "'" & Stored

            End If

            Return Stored

        End Function

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

    End Class

End Namespace
