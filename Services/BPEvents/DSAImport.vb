Imports Abovo.FileManager
Imports Abovo.AbovoAppCls
Imports Abovo.WSSecurity
Imports DevExpress.Spreadsheet
Imports System.IO
Imports System.Windows.Forms

Namespace Abovo
    Public Class DSAImport
        Private Const EarliestVersion As Double = 4.0988
        Private Const EarliestConsol As Double = 5.01

        Public Shared Function ImportSingleDSA_File(ModelID As Integer) As AbovoTransaction
            Const Action As String = "ImportSingleDSA_File"
            Try
                Using Dialog As New DevExpress.XtraEditors.XtraOpenFileDialog()
                    Dialog.Filter = "Abovo DSA models (*.xls;*.xlsx;*.xlsm;*.xlsb;*.adsa)|*.xls;*.xlsx;*.xlsm;*.xlsb;*.adsa"
                    Dialog.Title = "Select scheme"
                    If Dialog.ShowDialog() <> DialogResult.OK Then Return Cancelled(Action, "Scheme import cancelled.")
                    Dim Name As String = ImportScheme(ModelID, Dialog.FileName, "Single")
                    Return Succeeded(Action, "Scheme '" & Name & "' imported successfully.")
                End Using
            Catch ex As Exception
                Return Failed(Action, "The scheme could not be imported.", ex)
            End Try
        End Function

        Public Shared Function ImportConsolDSA_File(ModelID As Integer) As AbovoTransaction
            Const Action As String = "ImportConsolDSA_File"
            Try
                Using Dialog As New DevExpress.XtraEditors.XtraOpenFileDialog()
                    Dialog.Filter = "Consolidated DSA models (*.xls;*.xlsx;*.xlsm;*.xlsb;*.adsa)|*.xls;*.xlsx;*.xlsm;*.xlsb;*.adsa"
                    Dialog.Title = "Select consolidated model"
                    If Dialog.ShowDialog() <> DialogResult.OK Then Return Cancelled(Action, "Consolidated-scheme import cancelled.")
                    If MessageBox.Show("Schemes must either be all committed or all uncommitted." & Environment.NewLine & "Does this model contain only one type?", "Committed or uncommitted", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                        Return Cancelled(Action, "Consolidated-scheme import abandoned.")
                    End If
                    Dim Name As String = ImportScheme(ModelID, Dialog.FileName, "Consol")
                    Return Succeeded(Action, "Consolidated scheme '" & Name & "' imported successfully.")
                End Using
            Catch ex As Exception
                Return Failed(Action, "The consolidated schemes could not be imported.", ex)
            End Try
        End Function

        Public Shared Function DSA_Folder(ModelID As Integer) As AbovoTransaction
            Const Action As String = "ImportMultiDSA_Files"
            Try
                Using Dialog As New FolderBrowserDialog()
                    Dialog.Description = "Please select the FOLDER TO IMPORT"
                    If Dialog.ShowDialog() <> DialogResult.OK Then Return Cancelled(Action, "Folder import cancelled.")
                    Dim Files As New List(Of String)()
                    For Each Candidate As String In Directory.GetFiles(Dialog.SelectedPath)
                        Select Case Path.GetExtension(Candidate).ToLowerInvariant()
                            Case ".xls", ".xlsx", ".xlsm", ".xlsb", ".adsa"
                                Files.Add(Candidate)
                        End Select
                    Next
                    Files.Sort(StringComparer.OrdinalIgnoreCase)
                    If Files.Count = 0 Then Throw New InvalidDataException("No Excel or Abovo DSA files were found in the selected folder.")
                    If MessageBox.Show("This will import all compatible files in:" & Environment.NewLine & Dialog.SelectedPath & Environment.NewLine & Environment.NewLine & "Are you sure?", "Confirm Import", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) <> DialogResult.OK Then
                        Return Cancelled(Action, "Folder import cancelled.")
                    End If
                    ResetRejects(ModelID)
                    Dim Imported As Integer = 0
                    Dim Rejected As Integer = 0
                    For Each FilePath As String In Files
                        Try
                            ImportScheme(ModelID, FilePath, "Folder")
                            Imported += 1
                        Catch ex As Exception
                            Rejected += 1
                            AppendReject(ModelID, FilePath, ex.Message)
                        End Try
                    Next
                    If Imported = 0 Then Throw New InvalidDataException("No files were successfully imported. " & Rejected.ToString() & " file(s) were rejected.")
                    Dim Message As String = Imported.ToString() & " scheme file(s) imported."
                    If Rejected > 0 Then Message &= Environment.NewLine & Rejected.ToString() & " file(s) were rejected; details are recorded in DSARejects."
                    Return Succeeded(Action, Message)
                End Using
            Catch ex As Exception
                Return Failed(Action, "The folder of schemes could not be imported.", ex)
            End Try
        End Function

        Public Shared Function ImportDSA_Template(ModelID As Integer) As AbovoTransaction
            Const Action As String = "ImportDSA_Template"
            Dim SourceID As Integer = -1
            Try
                Dim SourcePath As String
                Using Dialog As New DevExpress.XtraEditors.XtraOpenFileDialog()
                    Dialog.Filter = "Development templates (*.xls;*.xlsx;*.xlsm;*.xlsb)|*.xls;*.xlsx;*.xlsm;*.xlsb"
                    Dialog.Title = "Select template file"
                    If Dialog.ShowDialog() <> DialogResult.OK Then Return Cancelled(Action, "Development-template import cancelled.")
                    SourcePath = Dialog.FileName
                End Using
                Dim OpenResult As AbovoTransaction = OpenSource(SourcePath)
                If OpenResult.BError Then Throw New InvalidDataException(OpenResult.StrResponseMessage)
                SourceID = OpenResult.IntegerReturn
                Dim BP As IWorkbook = Workbook(ModelID)
                Dim Source As IWorkbook = Workbook(SourceID)
                RequireSheet(BP, "Development BP Assumptions", "business plan")
                RequireSheet(Source, "Hidden - Workings", "development template")
                Dim ExportSheet As Worksheet = RequireSheet(Source, "Dvpt Export", "development template")
                Dim Count As Integer = PositiveInteger(RequiredRange(Source, "SchemeCount", "development template"), "SchemeCount")
                Dim Workings As CellRange = RequiredRange(Source, "TemplateWorkings", "development template")
                Dim ClearSource As CellRange = RequiredRange(Source, "RangesToClear", "development template")
                Dim SchemeRange As CellRange = RequiredRange(BP, "HouseTypeInID", "business plan")
                Dim Mappings As List(Of TemplateMapping) = ReadMappings(BP, ExportSheet, Workings, Count)
                Dim ClearNames As List(Of String) = ReadClearNames(BP, ClearSource)
                Dim ClearExisting As Boolean = MessageBox.Show("Do you wish to clear Development BP Assumptions before import?" & Environment.NewLine & Environment.NewLine & "Otherwise new columns will be added.", "Clear or add columns", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes
                Dim ExistingCount As Integer = Math.Max(0, SchemeRange.ColumnCount - 1)
                Dim RequiredCount As Integer = If(ClearExisting, Math.Max(Count, 10) + 1, SchemeRange.ColumnCount + Count)
                Dim Offset As Integer = If(ClearExisting, 0, ExistingCount)
                ResizeDevelopmentColumns(ModelID, RequiredCount)
                If ClearExisting Then
                    Dim ColumnsToClear As Integer = RequiredCount - 1
                    For Each RangeName As String In ClearNames
                        Dim Target As CellRange = RequiredRange(BP, RangeName, "business plan")
                        If Target.ColumnCount < ColumnsToClear Then Throw New InvalidDataException("The target range '" & RangeName & "' is too narrow for the template schemes.")
                        Target.Worksheet.Range.FromLTRB(Target.LeftColumnIndex, Target.TopRowIndex, Target.LeftColumnIndex + ColumnsToClear - 1, Target.BottomRowIndex).ClearContents()
                    Next
                End If
                For Each Mapping As TemplateMapping In Mappings
                    Dim Target As CellRange = RequiredRange(BP, Mapping.TargetName, "business plan")
                    If Target.RowCount < Mapping.RowCount OrElse Target.ColumnCount < Offset + Count Then Throw New InvalidDataException("The target range '" & Mapping.TargetName & "' does not fit the template data.")
                    Dim Destination As CellRange = Target.Worksheet.Range.FromLTRB(Target.LeftColumnIndex + Offset, Target.TopRowIndex, Target.LeftColumnIndex + Offset + Count - 1, Target.TopRowIndex + Mapping.RowCount - 1)
                    Destination.CopyFrom(Mapping.Source, PasteSpecial.Values, True)
                Next
                RequiredRange(BP, "DvptTemplFile", "business plan")(0, 0).Value = CellValue.FromObject(SourcePath)
                RequiredRange(BP, "DateStampDvptTempl", "business plan")(0, 0).Value = CellValue.FromObject(DateTime.Now)
                ExcelModels(ModelID).SetDirtyFlag()
                ExcelModels(ModelID).WBCalcEngine.CalcFile()
                Return Succeeded(Action, Count.ToString() & " development scheme(s) imported from the template.")
            Catch ex As Exception
                Return Failed(Action, "The development template could not be imported.", ex)
            Finally
                CloseModel(SourceID)
            End Try
        End Function

        Private Shared Function ImportScheme(ModelID As Integer, FilePath As String, ImportType As String) As String
            Dim SourceID As Integer = -1
            Dim SchemeMade As Boolean = False
            Dim TotalsMade As Boolean = False
            Dim Name As String = String.Empty
            Dim BP As IWorkbook = Workbook(ModelID)
            Dim ListSheet As Worksheet = RequireSheet(BP, "List Imported", "business plan")
            Dim TotalsTemplate As Worksheet = RequireSheet(BP, "Hidden - Tenure Totals Start", "business plan")
            Dim ImportsStart As Worksheet = RequireSheet(BP, "Hidden - Imports Start", "business plan")
            Dim TotalsEnd As Worksheet = RequireSheet(BP, "Hidden - Tenure Totals End", "business plan")
            Dim ListProtected As Boolean = ListSheet.IsProtected
            Dim TemplateProtected As Boolean = TotalsTemplate.IsProtected
            Dim ImportsVisible As Boolean = ImportsStart.Visible
            Dim TemplateVisible As Boolean = TotalsTemplate.Visible
            Dim EndVisible As Boolean = TotalsEnd.Visible
            Try
                Dim OpenResult As AbovoTransaction = OpenSource(FilePath)
                If OpenResult.BError Then Throw New InvalidDataException(OpenResult.StrResponseMessage)
                SourceID = OpenResult.IntegerReturn
                Dim Source As IWorkbook = Workbook(SourceID)
                Dim GlobalSheet As Worksheet = RequireSheet(Source, "Global Assumptions", "selected DSA model")
                RequireSheet(Source, "Hidden - Export Sheet", "selected DSA model")
                Dim Version As Double = Number(RequiredRange(Source, "ModelVersion", "selected DSA model"), "ModelVersion")
                Dim IsConsol As Boolean = String.Equals(ImportType, "Consol", StringComparison.Ordinal)
                Dim InitialChecks As Double = 0
                If IsConsol Then
                    If Version < EarliestConsol Then Throw New InvalidDataException("Please use a consolidated DSA model no earlier than Version " & EarliestConsol.ToString())
                    If Not GlobalSheet.Cells("A1").DisplayText.StartsWith("Consolidated DSA", StringComparison.OrdinalIgnoreCase) Then Throw New InvalidDataException("This does not appear to be a consolidated DSA model.")
                Else
                    RequireSheet(Source, "Unit Handovers & Sales", "selected DSA model")
                    RequireSheet(Source, "Cashflow Selector", "selected DSA model")
                    If Version < EarliestVersion Then Throw New InvalidDataException("Please use a DSA model no earlier than Version " & EarliestVersion.ToString())
                    InitialChecks = Number(RequiredRange(Source, "CheckTotal", "selected DSA model"), "CheckTotal")
                End If
                RequiredRange(Source, "BplanYear", "selected DSA model")(0, 0).Value = RequiredRange(BP, "TransferDate", "business plan")(0, 0).Value
                Source.CalculateFull()
                Name = SanitiseName(RequiredRange(Source, "SchemeName", "selected DSA model")(0, 0).DisplayText)
                If SchemeExists(BP, Name) Then Throw New InvalidDataException("Scheme '" & Name & "' has already been imported.")
                If Not IsConsol Then
                    Dim FinalChecks As Double = Number(RequiredRange(Source, "CheckTotal", "selected DSA model"), "CheckTotal")
                    If FinalChecks > 0 AndAlso InitialChecks > 0 Then Throw New InvalidDataException("There are errors shown on the Check Sheet of the DSA model.")
                    If FinalChecks > 0 Then Throw New InvalidDataException("Dates are too early in the DSA model; please contact an Abovo representative.")
                End If
                Dim Commitment As String
                If IsConsol Then
                    Commitment = If(MessageBox.Show("Does this model contain only Committed schemes?" & Environment.NewLine & "Select No if they are uncommitted.", "Committed or uncommitted", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes, "Committed", "Uncommitted")
                Else
                    Commitment = RequiredRange(Source, "Rep_Global_01", "selected DSA model")(0, 0).DisplayText
                End If
                ImportsStart.Visible = True
                TotalsTemplate.Visible = True
                TotalsEnd.Visible = True
                If ListProtected Then UNProtectWS(ModelID, ListSheet.Name)
                If TemplateProtected Then UNProtectWS(ModelID, TotalsTemplate.Name)
                InsertSchemeSheet(BP, Source, Name, IsConsol)
                SchemeMade = True
                InsertTotalsSheet(BP, Name)
                TotalsMade = True
                UpdateList(BP, Name, FilePath, Commitment)
                ExcelModels(ModelID).SetDirtyFlag()
                ExcelModels(ModelID).WBCalcEngine.CalcFile()
                Return Name
            Catch
                If TotalsMade AndAlso BP.Worksheets.Contains(Name & " Total") Then BP.Worksheets.Remove(BP.Worksheets(Name & " Total"))
                If SchemeMade AndAlso BP.Worksheets.Contains(Name) Then BP.Worksheets.Remove(BP.Worksheets(Name))
                Throw
            Finally
                ImportsStart.Visible = ImportsVisible
                TotalsTemplate.Visible = TemplateVisible
                TotalsEnd.Visible = EndVisible
                If ListProtected Then ProtectWS(ModelID, ListSheet.Name)
                If TemplateProtected Then ProtectWS(ModelID, TotalsTemplate.Name)
                CloseModel(SourceID)
            End Try
        End Function

        Private Shared Sub InsertSchemeSheet(BP As IWorkbook, Source As IWorkbook, Name As String, IsConsol As Boolean)
            If BP.Worksheets.Contains(Name) Then BP.Worksheets.Remove(BP.Worksheets(Name))
            BP.Worksheets.Insert(BP.Worksheets("Hidden - Tenure Totals Start").Index, Name)
            If Not IsConsol Then
                Source.Worksheets("Cashflow Selector").Cells("B6").Value = "All"
                Source.Worksheets("Cashflow Selector").Cells("B7").Value = "All"
                Source.Calculate()
            End If
            Dim Data As CellRange = Source.Worksheets("Hidden - Export Sheet").GetUsedRange()
            BP.Worksheets(Name).Cells("A1").CopyFrom(Data, PasteSpecial.Values)
            BP.Worksheets(Name).Cells("A1").CopyFrom(Data, PasteSpecial.Formats)
        End Sub

        Private Shared Sub InsertTotalsSheet(BP As IWorkbook, SchemeName As String)
            Dim Name As String = SchemeName & " Total"
            If BP.Worksheets.Contains(Name) Then BP.Worksheets.Remove(BP.Worksheets(Name))
            BP.Worksheets.Insert(BP.Worksheets("Hidden - Tenure Totals End").Index, Name)
            Dim Target As Worksheet = BP.Worksheets(Name)
            Dim Data As CellRange = BP.Worksheets("Hidden - Tenure Totals Start").GetUsedRange()
            Target.Cells("A1").CopyFrom(Data, PasteSpecial.ColumnWidths)
            Target.Cells("A1").CopyFrom(Data, PasteSpecial.Formats)
            Target.Cells("A1").CopyFrom(Data, PasteSpecial.Formulas Or PasteSpecial.Values Or PasteSpecial.NumberFormats)
            Target.Cells("A1").Value = SchemeName
            ReplaceFormulas(Target, "Hidden - Imports Start", SchemeName)
            ReplaceFormulas(Target, "Hidden - Tenure Totals Start", Name)
        End Sub

        Private Shared Sub ReplaceFormulas(Sheet As Worksheet, SearchText As String, Replacement As String)
            Dim Options As New SearchOptions With {.SearchBy = SearchBy.Rows, .SearchIn = SearchIn.Formulas, .MatchEntireCellContents = False}
            For Each FormulaCell As Cell In Sheet.Search(SearchText, Options)
                FormulaCell.Formula = FormulaCell.Formula.Replace(SearchText, Replacement)
            Next
        End Sub

        Private Shared Sub UpdateList(BP As IWorkbook, SchemeName As String, FilePath As String, Commitment As String)
            Dim Imported As CellRange = RequiredRange(BP, "ImportedSchemes", "business plan")
            Dim Offset As Integer = Imported.RowCount
            Dim Row As Integer = Imported.TopRowIndex + Offset
            Dim ExcelRow As Integer = Row + 1
            Dim Sheet As Worksheet = Imported.Worksheet
            Sheet.Cells(Row, 0).Value = SchemeName
            Sheet.Cells(Row, 1).Value = DateTime.Now
            Sheet.Cells(Row, 2).Value = FilePath
            Sheet.Cells(Row, 3).Value = Commitment
            Sheet.Cells(Row, 4).Value = "Yes"
            Sheet.Cells(Row, 5).Formula = "=IF(ConsolOption=1,""No"",IF(AND(D" & ExcelRow.ToString() & "<>""Committed"",ExclUnCommitDvpt=""TRUE""),""No"",IF(E" & ExcelRow.ToString() & "=""No"",""No"",""Yes"")))"
            Sheet.DataValidations.Add(Sheet.Cells(Row, 4), DataValidationType.List, "Yes,No")
            Sheet.Range.FromLTRB(10, Row, 23, Row).CopyFrom(
                Sheet.Range.FromLTRB(10, Imported.TopRowIndex, 23, Imported.TopRowIndex),
                PasteSpecial.Formats)
            For Index As Integer = 0 To 10
                Dim SourceColumn As Char = ChrW(AscW("D"c) + Index)
                Sheet.Cells(Row, 10 + Index).Formula = "=SUM('" & SchemeName & "'!" & SourceColumn & "$161:" & SourceColumn & "$201)"
            Next
            Sheet.Cells(Row, 21).Formula = "=SUM(K" & ExcelRow.ToString() & ":U" & ExcelRow.ToString() & ")"
            Sheet.Cells(Row, 22).Value = RequiredRange(BP, "TransferDate", "business plan")(0, 0).Value
            Sheet.Cells(Row, 23).Formula = "=IF(W" & ExcelRow.ToString() & "<>TransferDate,""WARNING: Date does not match current Business Plan Start Date"","""")"
            ResizeRows(BP, "ImportedSchemes", Offset + 1)
            ResizeRows(BP, "Appraisal1", Offset + 1)
            ResizeRows(BP, "Appraisal2", Offset + 1)
            RequiredRange(BP, "NumDSASchemes", "business plan")(0, 0).Value = Offset
            RequiredRange(BP, "SchemeLink", "business plan")(0, 0).Value = 1
        End Sub

        Private Shared Function SchemeExists(BP As IWorkbook, Name As String) As Boolean
            Dim Imported As CellRange = RequiredRange(BP, "ImportedSchemes", "business plan")
            For Row As Integer = 0 To Imported.RowCount - 1
                If String.Equals(Imported(Row, 0).DisplayText.Trim(), Name, StringComparison.OrdinalIgnoreCase) Then Return True
            Next
            Return False
        End Function

        Private Shared Sub ResetRejects(ModelID As Integer)
            Dim BP As IWorkbook = Workbook(ModelID)
            RequiredRange(BP, "DSARejects", "business plan").ClearContents()
            ResizeRows(BP, "DSARejects", 1)
        End Sub

        Private Shared Sub AppendReject(ModelID As Integer, FilePath As String, Reason As String)
            Dim BP As IWorkbook = Workbook(ModelID)
            Dim Rejects As CellRange = RequiredRange(BP, "DSARejects", "business plan")
            Dim Offset As Integer = Math.Max(0, Rejects.RowCount - 1)
            Rejects.Worksheet.Cells(Rejects.TopRowIndex + Offset, Rejects.LeftColumnIndex).Value = Path.GetFileName(FilePath) & " - " & Reason
            ResizeRows(BP, "DSARejects", Rejects.RowCount + 1)
        End Sub

        Private Shared Sub ResizeDevelopmentColumns(ModelID As Integer, RequiredCount As Integer)
            Dim BP As IWorkbook = Workbook(ModelID)
            Dim HouseTypes As CellRange = RequiredRange(BP, "HouseTypeInID", "business plan")
            Dim LastIdentified As CellRange = RequiredRange(BP, "LastIDColNum", "business plan")
            Dim Difference As Integer = RequiredCount - HouseTypes.ColumnCount
            If Difference = 0 Then Return

            Dim SheetNames As String() = {
                "Development BP Assumptions",
                "Development Stock",
                "Development Capital",
                "Development Revenue",
                "Development Expenditure",
                "Dvpt NonCash",
                "Dvpt Component Depn"
            }
            Dim Sheets As New List(Of Worksheet)()
            Dim WasProtected As New List(Of Boolean)()
            Dim WasVisible As New List(Of Boolean)()
            For Each SheetName As String In SheetNames
                Dim Sheet As Worksheet = RequireSheet(BP, SheetName, "business plan")
                Sheets.Add(Sheet)
                WasProtected.Add(Sheet.IsProtected)
                WasVisible.Add(Sheet.Visible)
            Next

            Dim InsertColumn As Integer = LastIdentified.LeftColumnIndex
            BP.BeginUpdate()
            Try
                For Index As Integer = 0 To Sheets.Count - 1
                    Dim Sheet As Worksheet = Sheets(Index)
                    Sheet.Visible = True
                    If WasProtected(Index) Then UNProtectWS(ModelID, Sheet.Name)
                    If Difference > 0 Then
                        Dim Used As CellRange = Sheet.GetUsedRange()
                        Dim Width As Single = Sheet.Columns(InsertColumn - 1).Width
                        Sheet.Columns.Insert(InsertColumn, Difference)
                        Dim Template As CellRange = Sheet.Range.FromLTRB(InsertColumn - 1, Used.TopRowIndex, InsertColumn - 1, Used.BottomRowIndex)
                        Dim Destination As CellRange = Sheet.Range.FromLTRB(InsertColumn, Used.TopRowIndex, InsertColumn + Difference - 1, Used.BottomRowIndex)
                        Destination.CopyFrom(Template, PasteSpecial.All)
                        For Column As Integer = InsertColumn To InsertColumn + Difference - 1
                            Sheet.Columns(Column).Width = Width
                        Next
                    Else
                        Dim DeleteCount As Integer = -Difference
                        Sheet.Columns.Remove(InsertColumn - DeleteCount, DeleteCount)
                    End If
                Next

                Dim Defined As DefinedName = BP.DefinedNames.GetDefinedName("HouseTypeInID")
                Defined.Range = BP.Worksheets("Development BP Assumptions").Range.FromLTRB(
                    HouseTypes.LeftColumnIndex,
                    HouseTypes.TopRowIndex,
                    HouseTypes.LeftColumnIndex + RequiredCount - 1,
                    HouseTypes.BottomRowIndex)
            Finally
                For Index As Integer = 0 To Sheets.Count - 1
                    If WasProtected(Index) Then ProtectWS(ModelID, Sheets(Index).Name)
                    Sheets(Index).Visible = WasVisible(Index)
                Next
                BP.EndUpdate()
            End Try

            If ExcelModels(ModelID).TransDBSync IsNot Nothing Then
                ExcelModels(ModelID).TransDBSync.SynchroniseForNamedRange("HouseTypeInID")
            End If
        End Sub

        Private Shared Sub ResizeRows(BP As IWorkbook, Name As String, Count As Integer)
            Dim Defined As DefinedName = BP.DefinedNames.GetDefinedName(Name)
            If Defined Is Nothing OrElse Defined.Range Is Nothing Then Throw New InvalidDataException("The business plan is missing the required named range '" & Name & "'.")
            Dim Current As CellRange = Defined.Range
            Defined.Range = Current.Worksheet.Range.FromLTRB(Current.LeftColumnIndex, Current.TopRowIndex, Current.RightColumnIndex, Current.TopRowIndex + Count - 1)
        End Sub

        Private Shared Function ReadMappings(BP As IWorkbook, ExportSheet As Worksheet, Workings As CellRange, SchemeCount As Integer) As List(Of TemplateMapping)
            Dim Result As New List(Of TemplateMapping)()
            Dim Anchor As Cell = ExportSheet.Cells("A6")
            For Row As Integer = 0 To Workings.RowCount - 1
                Dim TargetName As String = Workings(Row, 1).DisplayText.Trim()
                If String.IsNullOrWhiteSpace(TargetName) Then Continue For
                Dim SourceColumn As Integer = PositiveInteger(Workings(Row, 3), "TemplateWorkings source column")
                Dim TargetRows As Integer = PositiveInteger(Workings(Row, 4), "TemplateWorkings row count")
                If RequiredRange(BP, TargetName, "business plan").RowCount < TargetRows Then Throw New InvalidDataException("The target range '" & TargetName & "' is too short for the template data.")
                Dim Source As CellRange = ExportSheet.Range.FromLTRB(Anchor.ColumnIndex + SourceColumn - 1, Anchor.RowIndex, Anchor.ColumnIndex + SourceColumn + TargetRows - 2, Anchor.RowIndex + SchemeCount - 1)
                Result.Add(New TemplateMapping(TargetName, TargetRows, Source))
            Next
            If Result.Count = 0 Then Throw New InvalidDataException("TemplateWorkings contains no import mappings.")
            Return Result
        End Function

        Private Shared Function ReadClearNames(BP As IWorkbook, Source As CellRange) As List(Of String)
            Dim Result As New List(Of String)()
            For Each SourceCell As Cell In Source
                Dim Name As String = SourceCell.DisplayText.Trim()
                If String.IsNullOrWhiteSpace(Name) Then Continue For
                RequiredRange(BP, Name, "business plan")
                Result.Add(Name)
            Next
            Return Result
        End Function

        Private Shared Function OpenSource(Path As String) As AbovoTransaction
            Return OpenModel(Path, New FileInfo(Path), WorkbookOpenMode.ImportSource)
        End Function

        Private Shared Function Workbook(ID As Integer) As IWorkbook
            Dim Result As IWorkbook = GetWorkBook(ID)
            If Result Is Nothing Then Throw New InvalidOperationException("The requested workbook is not available.")
            Return Result
        End Function

        Private Shared Function RequireSheet(Book As IWorkbook, Name As String, Description As String) As Worksheet
            If Not Book.Worksheets.Contains(Name) Then Throw New InvalidDataException("The " & Description & " is missing the '" & Name & "' worksheet.")
            Return Book.Worksheets(Name)
        End Function

        Private Shared Function RequiredRange(Book As IWorkbook, Name As String, Description As String) As CellRange
            Try
                Dim Result As CellRange = Book.Range(Name)
                If Result IsNot Nothing Then Return Result
            Catch
            End Try
            Throw New InvalidDataException("The " & Description & " is missing the required named range '" & Name & "'.")
        End Function

        Private Shared Function Number(Range As CellRange, Description As String) As Double
            If Range(0, 0).Value.IsNumeric Then Return Range(0, 0).Value.NumericValue
            Dim Result As Double
            If Double.TryParse(Range(0, 0).DisplayText, Result) Then Return Result
            Throw New InvalidDataException(Description & " is not numeric.")
        End Function

        Private Shared Function PositiveInteger(Range As CellRange, Description As String) As Integer
            Return PositiveInteger(Range(0, 0), Description)
        End Function

        Private Shared Function PositiveInteger(ValueCell As Cell, Description As String) As Integer
            Dim Value As Double
            If ValueCell.Value.IsNumeric Then
                Value = ValueCell.Value.NumericValue
            ElseIf Not Double.TryParse(ValueCell.DisplayText, Value) Then
                Throw New InvalidDataException(Description & " is not numeric.")
            End If
            If Value < 1 OrElse Value <> Math.Truncate(Value) Then Throw New InvalidDataException(Description & " must be a positive whole number.")
            Return Convert.ToInt32(Value)
        End Function

        Private Shared Function SanitiseName(Value As String) As String
            Dim Result As String = Value.Trim()
            For Each Character As Char In New Char() {"/"c, "\"c, "*"c, "?"c, "'"c, "["c, "]"c, ":"c}
                Result = Result.Replace(Character, "_"c)
            Next
            If String.IsNullOrWhiteSpace(Result) Then Throw New InvalidDataException("The DSA model has no scheme name.")
            If Result.Length > 25 Then Throw New InvalidDataException("The scheme name is too long to create both scheme and totals worksheets: " & Result)
            Return Result
        End Function

        Private Shared Function Succeeded(Action As String, Message As String) As AbovoTransaction
            Return New AbovoTransaction(Action) With {.BSuccess = True, .BError = False, .StringReturn = Message, .StrResponseMessage = Message}
        End Function

        Private Shared Function Cancelled(Action As String, Message As String) As AbovoTransaction
            Return New AbovoTransaction(Action) With {.EventCancelled = True, .BError = False, .StringReturn = Message, .StrResponseMessage = Message}
        End Function

        Private Shared Function Failed(Action As String, Prefix As String, ErrorValue As Exception) As AbovoTransaction
            Dim Message As String = Prefix & Environment.NewLine & Environment.NewLine & ErrorValue.Message
            Return New AbovoTransaction(Action) With {.BSuccess = False, .BError = True, .StringReturn = ErrorValue.Message, .StrResponseMessage = Message}
        End Function

        Private Class TemplateMapping
            Public ReadOnly TargetName As String
            Public ReadOnly RowCount As Integer
            Public ReadOnly Source As CellRange
            Public Sub New(Name As String, Rows As Integer, SourceRange As CellRange)
                TargetName = Name
                RowCount = Rows
                Source = SourceRange
            End Sub
        End Class
    End Class
End Namespace
