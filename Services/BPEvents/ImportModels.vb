Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.LogDebugDev
Imports Abovo.WSSecurity
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraRichEdit.Import.OpenXml
Imports DevExpress.XtraRichEdit.Model

Namespace Abovo

    Public Class ImportModels



        Dim BPFile As String
        Dim MSFile As String

        Public Shared Function ImportStockRentModel(ModelID As Integer) As AbovoTransaction

            Const AllYears As Integer = 30
            Dim Result As New AbovoTransaction("ImportStockRentModel")
            Dim RentModelID As Integer = -1
            Dim Journal As New WorkbookCellJournal
            Dim MutationStarted As Boolean = False
            Dim StructuralMutation As Boolean = False

            Try
                Dim BusinessPlan As IWorkbook = FileManager.GetWorkBook(ModelID)
                If BusinessPlan Is Nothing Then
                    Throw New InvalidOperationException("The active business-plan workbook is not available.")
                End If

                Using FileDialog As New DevExpress.XtraEditors.XtraOpenFileDialog()
                    FileDialog.Filter =
                        "Rent model files (*.xls;*.xlsx;*.xlsm;*.xlsb;*.arm)|*.xls;*.xlsx;*.xlsm;*.xlsb;*.arm"
                    FileDialog.Title = "Select Rent Restructuring file"

                    If FileDialog.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                        Result.EventCancelled = True
                        Result.StrResponseMessage = "Rent-data import cancelled."
                        Return Result
                    End If

                    Dim RentModelPath As String = FileDialog.FileName
                    Dim FileOpenResult As AbovoTransaction =
                        FileManager.OpenModel(
                            RentModelPath,
                            New IO.FileInfo(RentModelPath),
                            FileManager.WorkbookOpenMode.ImportSource)

                    If FileOpenResult.BError Then
                        Throw New IO.InvalidDataException(FileOpenResult.StrResponseMessage)
                    End If

                    RentModelID = FileOpenResult.IntegerReturn
                End Using

                Dim RentModel As IWorkbook = FileManager.GetWorkBook(RentModelID)
                If RentModel Is Nothing Then
                    Throw New InvalidOperationException("The selected rent model could not be loaded.")
                End If

                'Resolve the complete transfer contract before changing the business
                'plan. This prevents a partially imported model when a required name
                'or worksheet is missing from the selected rent file.
                Dim TransferDate As CellRange = GetRequiredRange(BusinessPlan, "TransferDate", "business plan")
                Dim TargetBPCats As CellRange = GetRequiredRange(BusinessPlan, "BPCats", "business plan")
                Dim TargetStock As CellRange = GetRequiredRange(BusinessPlan, "StTrans", "business plan")
                Dim TargetRents As CellRange = GetRequiredRange(BusinessPlan, "BPRents", "business plan")
                Dim TargetReletRents As CellRange = GetRequiredRange(BusinessPlan, "BPTarRents", "business plan")
                Dim TargetReletRentsFV5 As CellRange = GetRequiredRange(BusinessPlan, "BPTarRentsFV5", "business plan")
                Dim TargetRentWeeks As CellRange = GetRequiredRange(BusinessPlan, "RentWks", "business plan")
                Dim TargetRealIncreases As CellRange = GetRequiredRange(BusinessPlan, "TransRealIncr", "business plan")
                Dim TargetRentFile As CellRange = GetRequiredRange(BusinessPlan, "Rentsfile", "business plan")
                Dim TargetTimestamp As CellRange = GetRequiredRange(BusinessPlan, "Datestamprent", "business plan")
                Dim TargetYearRows As CellRange = GetRequiredRange(BusinessPlan, "IR_ExistRentReal", "business plan")

                Dim SourceStartDate As CellRange = GetRequiredRange(RentModel, "BPStartDate", "rent model")
                Dim SourceBPCats As CellRange = GetRequiredRange(RentModel, "BPCats", "rent model")
                Dim SourceStock As CellRange = GetRequiredRange(RentModel, "StTrans", "rent model")
                Dim SourceRents As CellRange = GetRequiredRange(RentModel, "BPRents", "rent model")
                Dim SourceReletRents As CellRange = GetRequiredRange(RentModel, "BPTarRents", "rent model")
                Dim SourceRentWeeks As CellRange = GetRequiredRange(RentModel, "RentWks", "rent model")
                Dim SourceRealIncreases As CellRange = GetRequiredRange(RentModel, "TransRealIncr", "rent model")

                Journal.Capture(TargetBPCats)
                Journal.Capture(TargetStock)
                Journal.Capture(TargetRents)
                Journal.Capture(TargetReletRents)
                Journal.Capture(TargetReletRentsFV5)
                Journal.Capture(TargetRentWeeks)
                Journal.Capture(TargetRealIncreases)
                Journal.Capture(TargetRentFile)
                Journal.Capture(TargetTimestamp)
                Journal.Capture(TargetYearRows)

                If Not RentModel.Worksheets.Contains("Contents") Then
                    Throw New IO.InvalidDataException(
                        "The selected rent model is missing the 'Contents' worksheet.")
                End If

                'VBA Get_Rents uses xlPasteFormulas, which also carries a constant
                'TransferDate. Copy both cases explicitly in the DevExpress model.
                SourceStartDate.CopyFrom(
                    TransferDate,
                    PasteSpecial.Formulas Or PasteSpecial.Values)
                RentModel.Calculate()

                If SourceRealIncreases.RowCount < AllYears - 1 OrElse
                   SourceRealIncreases(AllYears - 2, 0).Value.IsError Then
                    Throw New IO.InvalidDataException(
                        "Please choose a later version of the rent model.")
                End If

                MutationStarted = True
                TargetBPCats.CopyFrom(SourceBPCats, PasteSpecial.Values)
                TargetStock.CopyFrom(SourceStock, PasteSpecial.Values)
                TargetRents.CopyFrom(SourceRents, PasteSpecial.Values)
                TargetReletRents.CopyFrom(SourceReletRents, PasteSpecial.Values)
                TargetReletRentsFV5.CopyFrom(SourceReletRents, PasteSpecial.Values)
                TargetRentWeeks.CopyFrom(SourceRentWeeks, PasteSpecial.Values)

                If TargetYearRows.RowCount < AllYears Then
                    StructuralMutation = True
                    Dim ResizeResult As AbovoTransaction =
                        WorkbookManager.SetRangeRowsToSize(
                            ModelID,
                            "IR_ExistRentReal",
                            AllYears)

                    If ResizeResult.BError Then
                        Throw New InvalidOperationException(ResizeResult.StringReturn)
                    End If

                    TargetYearRows =
                        GetRequiredRange(BusinessPlan, "IR_ExistRentReal", "business plan")
                    TargetRealIncreases =
                        GetRequiredRange(BusinessPlan, "TransRealIncr", "business plan")
                End If

                For YearNumber As Integer = 3 To AllYears
                    TargetYearRows(YearNumber - 1, 0).Value = YearNumber
                Next

                TargetRealIncreases.CopyFrom(SourceRealIncreases, PasteSpecial.Values)

                Dim Contents As Worksheet = RentModel.Worksheets("Contents")
                Dim RentFileLabel As Cell = Contents.Cells("B5")
                If String.IsNullOrWhiteSpace(RentFileLabel.DisplayText) Then
                    RentFileLabel = Contents.Cells("C5")
                End If

                TargetRentFile.Value = RentFileLabel.Value
                TargetTimestamp.Value = DateTime.Now

                FileManager.ExcelModels(ModelID).SetDirtyFlag()
                FileManager.ExcelModels(ModelID).WBCalcEngine.CalcFile()

                Result.BSuccess = True
                Result.BError = False
                Result.StringReturn = "Rent data imported successfully."
                Result.StrResponseMessage = Result.StringReturn

            Catch ex As Exception
                Result.BError = True
                Result.BSuccess = False
                Result.StringReturn = ex.Message
                Result.StrResponseMessage =
                    "The rent data could not be imported." &
                    Environment.NewLine & Environment.NewLine & ex.Message

                If MutationStarted Then
                    Dim RollbackFailure As String = String.Empty
                    Dim Restored As Boolean = Journal.Restore(RollbackFailure)
                    Try
                        Dim Model As IWorkbook = FileManager.GetWorkBook(ModelID)
                        If Model IsNot Nothing Then Model.Calculate()
                    Catch rollbackCalculationError As Exception
                        Restored = False
                        RollbackFailure &=
                            If(String.IsNullOrWhiteSpace(RollbackFailure), String.Empty, Environment.NewLine) &
                            "Rollback calculation: " & rollbackCalculationError.Message
                    End Try

                    If Restored AndAlso Not StructuralMutation Then
                        Result.StrResponseMessage &=
                            Environment.NewLine & Environment.NewLine &
                            "The previous business-plan values were restored."
                    Else
                        ModelSafetyManager.MarkRecoveryRequired(
                            ModelID,
                            "Import rent model",
                            If(String.IsNullOrWhiteSpace(RollbackFailure), ex.Message, RollbackFailure),
                            "Import Models")
                    End If
                End If

            Finally
                Try
                    If RentModelID >= 0 Then FileManager.CloseModel(RentModelID)
                Finally
                    Journal.Dispose()
                End Try
            End Try

            Return Result

        End Function

        Private Shared Function GetRequiredRange(Workbook As IWorkbook,
                                                 RangeName As String,
                                                 WorkbookDescription As String) As CellRange

            Try
                Dim RequiredRange As CellRange = Workbook.Range(RangeName)
                If RequiredRange IsNot Nothing Then Return RequiredRange
            Catch ex As Exception
                'Converted below to a consistent import-contract error.
            End Try

            Throw New IO.InvalidDataException(
                "The " & WorkbookDescription &
                " is missing the required named range '" & RangeName & "'.")

        End Function
        Public Shared Function ImportManagementServiceCosts(ModelID As Integer) As AbovoTransaction

            Const ImportYears As Integer = 5
            Const TargetSheetName As String = "Management Costs Assumptions"
            Const SourceSummarySheetName As String = "Summary Costs"
            Const SourceGlobalSheetName As String = "Global Assumptions"
            Dim Result As New AbovoTransaction("ImportManagementServiceCosts")
            Dim SourceModelID As Integer = -1
            Dim SourcePath As String = String.Empty
            Dim TargetWasProtected As Boolean = False
            Dim Journal As New WorkbookCellJournal
            Dim MutationStarted As Boolean = False

            Try
                Dim BusinessPlan As IWorkbook = FileManager.GetWorkBook(ModelID)
                If BusinessPlan Is Nothing Then
                    Throw New InvalidOperationException("The active business-plan workbook is not available.")
                End If

                Using FileDialog As New DevExpress.XtraEditors.XtraOpenFileDialog()
                    FileDialog.Title = "Select file which contains Management & Service Costs"
                    FileDialog.Filter = "Management and service cost files (*.xls;*.xlsx;*.xlsm;*.xlsb;*.amc)|*.xls;*.xlsx;*.xlsm;*.xlsb;*.amc"
                    If FileDialog.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                        Result.EventCancelled = True
                        Result.StrResponseMessage = "Management-cost import cancelled."
                        Return Result
                    End If

                    SourcePath = FileDialog.FileName
                    Dim OpenResult As AbovoTransaction =
                        FileManager.OpenModel(SourcePath,
                                              New IO.FileInfo(SourcePath),
                                              FileManager.WorkbookOpenMode.ImportSource)
                    If OpenResult.BError Then
                        Throw New IO.InvalidDataException(OpenResult.StrResponseMessage)
                    End If
                    SourceModelID = OpenResult.IntegerReturn
                End Using

                Dim SourceModel As IWorkbook = FileManager.GetWorkBook(SourceModelID)
                If SourceModel Is Nothing Then
                    Throw New InvalidOperationException("The selected management-cost model could not be loaded.")
                End If
                RequireWorksheet(BusinessPlan, TargetSheetName, "business plan")
                RequireWorksheet(SourceModel, SourceSummarySheetName, "management-cost model")
                RequireWorksheet(SourceModel, SourceGlobalSheetName, "management-cost model")
                Dim TargetSheet As Worksheet = BusinessPlan.Worksheets(TargetSheetName)

                'Resolve and size-check the entire contract before changing the
                'business plan, preventing a partial import from an invalid file.
                Dim SelectedCompany As CellRange = GetRequiredRange(BusinessPlan, "SelectTrust", "business plan")
                Dim TargetServices As CellRange = GetRequiredRange(BusinessPlan, "Services", "business plan")
                Dim TargetMCFile As CellRange = GetRequiredRange(BusinessPlan, "MCfile", "business plan")
                Dim TargetTimestamp As CellRange = GetRequiredRange(BusinessPlan, "Datestampmc", "business plan")
                Dim SourceCompany As CellRange = GetRequiredRange(SourceModel, "SelCompany", "management-cost model")
                Dim SourceServices As CellRange = GetRequiredRange(SourceModel, "Services", "management-cost model")
                Dim SourceInflation As CellRange = GetRequiredRange(SourceModel, "InflatIndic", "management-cost model")
                Dim SourceYear As CellRange = GetRequiredRange(SourceModel, "YearNo", "management-cost model")
                Dim SourceStaff As CellRange = GetRequiredRange(SourceModel, "StaffCosts", "management-cost model")
                Dim SourceOther As CellRange = GetRequiredRange(SourceModel, "OtherCosts", "management-cost model")
                Dim TargetStaff(ImportYears - 1) As CellRange
                Dim TargetOther(ImportYears - 1) As CellRange

                For YearIndex As Integer = 0 To ImportYears - 1
                    Dim YearNumber As String = (YearIndex + 1).ToString()
                    Dim StaffAnchor As CellRange = GetRequiredRange(BusinessPlan, "StaffCosts" & YearNumber, "business plan")
                    Dim OtherAnchor As CellRange = GetRequiredRange(BusinessPlan, "OtherCosts" & YearNumber, "business plan")
                    TargetStaff(YearIndex) = RangeFromAnchor(TargetSheet, StaffAnchor, SourceStaff)
                    TargetOther(YearIndex) = RangeFromAnchor(TargetSheet, OtherAnchor, SourceOther)
                Next

                Journal.Capture(TargetServices)
                Journal.Capture(TargetMCFile)
                Journal.Capture(TargetTimestamp)
                For YearIndex As Integer = 0 To ImportYears - 1
                    Journal.Capture(TargetStaff(YearIndex))
                    Journal.Capture(TargetOther(YearIndex))
                Next

                If TargetServices.RowCount <> SourceServices.ColumnCount OrElse
                   TargetServices.ColumnCount <> SourceServices.RowCount Then
                    Throw New IO.InvalidDataException("The Services range in the selected file does not fit the transposed Services range in the business plan.")
                End If

                Dim CompanyName As String = SelectedCompany(0, 0).DisplayText.Trim()
                If String.IsNullOrWhiteSpace(CompanyName) Then
                    Throw New IO.InvalidDataException("Select a company in the business plan before importing management costs.")
                End If

                Dim SummarySheet As Worksheet = SourceModel.Worksheets(SourceSummarySheetName)
                Dim CompanyValidation As DataValidation = SummarySheet.DataValidations.GetDataValidation(SourceCompany(0, 0))
                If CompanyValidation Is Nothing OrElse CompanyValidation.ValidationType <> DataValidationType.List Then
                    Throw New IO.InvalidDataException("The SelCompany list is missing from the selected management-cost model.")
                End If
                If Not SummarySheet.DataValidations.Validate(SourceCompany(0, 0), CellValue.FromObject(CompanyName)) Then
                    Throw New IO.InvalidDataException("Import cannot continue." & Environment.NewLine & "Please ensure the Company name is consistent in both files.")
                End If

                SourceCompany(0, 0).Value = CellValue.FromObject(CompanyName)
                SourceInflation(0, 0).Value = CellValue.FromObject("N")
                SourceModel.Calculate()

                TargetWasProtected = TargetSheet.IsProtected
                If TargetWasProtected Then UNProtectWS(ModelID, TargetSheetName)

                MutationStarted = True
                TargetServices.CopyFrom(SourceServices, PasteSpecial.Values, True)
                For YearIndex As Integer = 0 To ImportYears - 1
                    SourceYear(0, 0).Value = YearIndex + 1
                    SourceModel.Calculate()
                    TargetStaff(YearIndex).CopyFrom(SourceStaff, PasteSpecial.Values)
                    TargetOther(YearIndex).CopyFrom(SourceOther, PasteSpecial.Values)
                Next

                TargetMCFile(0, 0).Value = CellValue.FromObject(
                    IO.Path.Combine(IO.Path.GetDirectoryName(SourcePath),
                                    "[" & IO.Path.GetFileName(SourcePath) & "]" & SourceGlobalSheetName))
                TargetTimestamp(0, 0).Value = CellValue.FromObject(DateTime.Now)

                FileManager.ExcelModels(ModelID).SetDirtyFlag()
                FileManager.ExcelModels(ModelID).WBCalcEngine.CalcFile()
                Result.BSuccess = True
                Result.BError = False
                Result.StringReturn = "Management and service costs imported successfully."
                Result.StrResponseMessage = Result.StringReturn

            Catch ex As Exception
                Result.BError = True
                Result.BSuccess = False
                Result.StringReturn = ex.Message
                Result.StrResponseMessage = "The management and service costs could not be imported." & Environment.NewLine & Environment.NewLine & ex.Message

                If MutationStarted Then
                    Dim RollbackFailure As String = String.Empty
                    Dim Restored As Boolean = Journal.Restore(RollbackFailure)
                    Try
                        Dim Model As IWorkbook = FileManager.GetWorkBook(ModelID)
                        If Model IsNot Nothing Then Model.Calculate()
                    Catch rollbackCalculationError As Exception
                        Restored = False
                        RollbackFailure &=
                            If(String.IsNullOrWhiteSpace(RollbackFailure), String.Empty, Environment.NewLine) &
                            "Rollback calculation: " & rollbackCalculationError.Message
                    End Try

                    If Restored Then
                        Result.StrResponseMessage &=
                            Environment.NewLine & Environment.NewLine &
                            "The previous business-plan values were restored."
                    Else
                        ModelSafetyManager.MarkRecoveryRequired(
                            ModelID,
                            "Import management and service costs",
                            RollbackFailure,
                            "Import Models")
                    End If
                End If
            Finally
                If TargetWasProtected Then
                    Try
                        ProtectWS(ModelID, TargetSheetName)
                    Catch protectionError As Exception
                        Result.BError = True
                        Result.BSuccess = False
                        Result.StrResponseMessage &=
                            Environment.NewLine & Environment.NewLine &
                            "Worksheet protection could not be restored: " & protectionError.Message
                        ModelSafetyManager.MarkRecoveryRequired(
                            ModelID,
                            "Import management and service costs",
                            protectionError.Message,
                            "Import Models",
                            TargetSheetName)
                    End Try
                End If
                Try
                    If SourceModelID >= 0 Then FileManager.CloseModel(SourceModelID)
                Finally
                    Journal.Dispose()
                End Try
            End Try

            Return Result

        End Function

        Private Shared Sub RequireWorksheet(Workbook As IWorkbook,
                                            WorksheetName As String,
                                            WorkbookDescription As String)
            If Workbook.Worksheets.Contains(WorksheetName) Then Return
            Throw New IO.InvalidDataException("The " & WorkbookDescription & " is missing the '" & WorksheetName & "' worksheet.")
        End Sub

        Private Shared Function RangeFromAnchor(Worksheet As Worksheet,
                                                Anchor As CellRange,
                                                SourceRange As CellRange) As CellRange
            Return Worksheet.Range.FromLTRB(
                Anchor.LeftColumnIndex,
                Anchor.TopRowIndex,
                Anchor.LeftColumnIndex + SourceRange.ColumnCount - 1,
                Anchor.TopRowIndex + SourceRange.RowCount - 1)
        End Function

        Public Shared Function ImportStockConditionSurvey(SetModelID As Integer) As AbovoTransaction
            Dim Result As New AbovoTransaction("ImportStockConditionSurvey")
            Dim ActiveRMModelID As Integer = -1
            Dim Journal As New WorkbookCellJournal
            Dim MutationStarted As Boolean = False

            Try
                Dim BusPlanFile As IWorkbook = FileManager.GetWorkBook(SetModelID)
                If BusPlanFile Is Nothing Then Throw New InvalidOperationException("The active business-plan workbook is not available.")

                Dim FileToOpen As String
                Using FileDialog As New DevExpress.XtraEditors.XtraOpenFileDialog With {
                    .Title = "Select file which contains Stock Condition Survey data",
                    .Filter = "Abovo Repairs and Maint Models |*.xlsm;*.xlsb;*.armc"}
                    If FileDialog.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                        Result.EventCancelled = True
                        Result.StrResponseMessage = "Stock-condition import cancelled."
                        Return Result
                    End If
                    FileToOpen = FileDialog.FileName
                End Using

                Dim FileOpenResult As AbovoTransaction =
                    FileManager.OpenModel(FileToOpen, New IO.FileInfo(FileToOpen), FileManager.WorkbookOpenMode.ImportSource)
                If FileOpenResult.BError Then Throw New IO.InvalidDataException(FileOpenResult.StrResponseMessage)
                ActiveRMModelID = FileOpenResult.IntegerReturn

                Dim ActiveRMModel As IWorkbook = FileManager.GetWorkBook(ActiveRMModelID)
                If ActiveRMModel Is Nothing Then Throw New InvalidOperationException("The selected stock-condition model could not be loaded.")
                RequireWorksheet(ActiveRMModel, "Totals", "stock-condition model")
                Dim TargetFile As CellRange = GetRequiredRange(BusPlanFile, "SCSfile", "business plan")
                Dim TargetTimestamp As CellRange = GetRequiredRange(BusPlanFile, "Datestampscs", "business plan")
                Journal.Capture(TargetFile)
                Journal.Capture(TargetTimestamp)

                Dim SourceFileCell As Cell = ActiveRMModel.Worksheets("Totals").Cells("Z1")
                SourceFileCell.Formula = "=CELL(""FILENAME"")"
                ActiveRMModel.Calculate()

                MutationStarted = True
                TargetFile(0, 0).SetValueFromText(SourceFileCell.DisplayText)
                TargetTimestamp(0, 0).SetValue(Now())
                FileManager.ExcelModels(SetModelID).SetDirtyFlag()
                FileManager.ExcelModels(SetModelID).WBCalcEngine.CalcFile()

                Result.BSuccess = True
                Result.BError = False
                Result.StringReturn = "Stock-condition survey details imported successfully."
                Result.StrResponseMessage = Result.StringReturn
            Catch ex As Exception
                Result.BError = True
                Result.BSuccess = False
                Result.StringReturn = ex.Message
                Result.StrResponseMessage = "The stock-condition survey could not be imported." &
                    Environment.NewLine & Environment.NewLine & ex.Message
                If MutationStarted Then
                    Dim RollbackFailure As String = String.Empty
                    If Journal.Restore(RollbackFailure) Then
                        Result.StrResponseMessage &= Environment.NewLine & Environment.NewLine &
                            "The previous business-plan values were restored."
                    Else
                        ModelSafetyManager.MarkRecoveryRequired(
                            SetModelID,
                            "Import stock-condition survey",
                            RollbackFailure,
                            "Import Models")
                    End If
                End If
            Finally
                Try
                    If ActiveRMModelID >= 0 Then FileManager.CloseModel(ActiveRMModelID)
                Catch closeError As Exception
                    Result.BError = True
                    Result.BSuccess = False
                    Result.StrResponseMessage &= Environment.NewLine & Environment.NewLine &
                        "The source stock-condition model could not be closed: " & closeError.Message
                    SystemMessageManager.Publish(SetModelID, Result.StrResponseMessage,
                        SystemMessageSeverity.Error, "Import Models", "Stock Condition Survey")
                Finally
                    Journal.Dispose()
                End Try
            End Try

            Return Result
        End Function






    End Class
End Namespace
