Imports System.IO
Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.LogDebugDev
Imports Abovo.WSSecurity
Imports DevExpress.CodeParser
Imports DevExpress.DataAccess.Native.EntityFramework
Imports DevExpress.Pdf.Native.BouncyCastle.Asn1.X509
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraRichEdit.Export.Doc
Imports DevExpress.XtraRichEdit.Import.OpenXml
Imports DevExpress.XtraRichEdit.Model
Imports Microsoft.Office.Interop.Excel

Namespace Abovo
    Public Class DSAImport

        Private Shared CheckStatus As String
        Private Shared ReadOnly EarliestVersion As Double = 4.0988
        Private Shared ReadOnly EarliestConsol As Double = 5.01
        Private Shared ReadOnly HeadingColumns = 6

        Public Shared Sub ImportSingleDSA_File(SetModelID As Integer)

            Dim FileToOpen As String
            Dim BPModelID As Integer = SetModelID
            Dim XtraOpenFileDialogMainScreen As New DevExpress.XtraEditors.XtraOpenFileDialog()

            XtraOpenFileDialogMainScreen.Filter = "Abovo DSA Models |*.xlsm;*.xlsb;*.adsa"
            XtraOpenFileDialogMainScreen.Title = "Select DSA model to import"

            If XtraOpenFileDialogMainScreen.ShowDialog = DialogResult.Cancel Then

                Exit Sub

            End If

            FileToOpen = XtraOpenFileDialogMainScreen.FileName

            UNProtectWS(BPModelID, "List Imported")

            DSA_Import(BPModelID, FileToOpen, "Single")

            ProtectWS(BPModelID, "List Imported")

            GetWorkBook(BPModelID).CalculateFull()

            MsgBox("Import complete")

        End Sub
        Public Shared Sub ImportConsolDSA_File(SetModelID As Integer)

            Dim FileToOpen As String

            Dim BPModelID As Integer = SetModelID

            Dim XtraOpenFileDialogMainScreen As New DevExpress.XtraEditors.XtraOpenFileDialog()
            XtraOpenFileDialogMainScreen.Filter = "Abovo DSA Models |*.xlsm;*.xlsb;*.adsa"
            XtraOpenFileDialogMainScreen.Title = "Select DSA model to import"

            If XtraOpenFileDialogMainScreen.ShowDialog = DialogResult.Cancel Then

                Exit Sub

            End If

            FileToOpen = XtraOpenFileDialogMainScreen.FileName

            If MsgBox("Schemes must either be all committed or all uncommitted." & vbLf & "Does this model contain only one type?", MsgBoxStyle.YesNo, "Committed or uncommitted") = 6 Then                                 ' "Yes" response

                UNProtectWS(BPModelID, "List Imported")

                DSA_Import(BPModelID, FileToOpen, "Consol")

                ProtectWS(BPModelID, "List Imported")

                Dim BusPlanFile As IWorkbook = GetWorkBook(BPModelID)

                BusPlanFile.CalculateFull()

                MsgBox("Import complete")

            Else

                MsgBox("Import abandoned")
                Exit Sub

            End If

        End Sub

        Public Shared Sub DSA_Folder(SetModelID As Integer)

            Dim BusPlanFile As IWorkbook
            Dim strFileName As String
            Dim FolderPath As String
            Dim FileToOpen As String = Nothing

            Dim CurrentPosition As Integer
            Dim LastPosition As Integer


            Dim BPModelID As Integer = SetModelID
            Dim XtraOpenfolderDialogMainScreen As New FolderBrowserDialog()

            XtraOpenfolderDialogMainScreen.Description = "Select folder of DSA models to import"

            If XtraOpenfolderDialogMainScreen.ShowDialog = DialogResult.Cancel Then

                Exit Sub

            End If

            FolderPath = XtraOpenfolderDialogMainScreen.SelectedPath

            ClearRejectReport(BPModelID)

            ChDir(FolderPath)

            Dim FoundsFiles As String

            FoundsFiles = Dir(FolderPath, "*.xlsm;*.xlsb")

            If FoundsFiles = "" Then

                MsgBox("No files found in selected folder.")
                Exit Sub

            End If

            BusPlanFile = FileManager.GetWorkBook(SetModelID)

            Dim ErrorCount As Integer = 0
            Dim FileImportCount As Integer = 0

            If FoundsFiles <> "" Then

                If MsgBox("This will import all files in the folder " & FolderPath & Chr(13) & Chr(10) & "Are you sure?", vbOKCancel, "Confirm Import") <> vbCancel Then

                    UNProtectWS(BPModelID, "List Imported")

                    CurrentPosition = 0
                    ' Use the positions of "\" characters to distinguish the path from the filename

                    Do While FoundsFiles <> ""

                        LastPosition = CurrentPosition
                        CurrentPosition = InStr(LastPosition + 1, FileToOpen, "\")

                    Loop

                    strFileName = Dir("*.xls*", 7)

                    Do While strFileName <> ""

                        If DSA_Import(BPModelID, strFileName, "Folder") Then

                            FileImportCount += 1

                        Else

                            ErrorCount += 1

                        End If

                        'set name of next file
                        strFileName = Dir()

                    Loop


                    If FileImportCount > 0 Then

                        BusPlanFile.CalculateFull()

                        MsgBox("Imports Completed. & " & FileImportCount & " files were imported.", vbOKOnly, "DSA Import")

                    Else

                        MsgBox("No files were sucessfully imported.", vbOKOnly, "DSA Import")

                    End If

                    If ErrorCount > 0 Then

                        MsgBox("The following errors occurred", vbOKOnly, "DSA Import")

                        ShowDSAErrors(BPModelID)

                    End If

                End If

            Else

                MsgBox("No files were imported.")

            End If

            ProtectWS(BPModelID, "List Imported")
            ProtectWS(BPModelID, "Hidden - Lookup Lists")



        End Sub
        Public Shared Sub ShowDSAErrors(SetModelID As Integer)

            'GetWorkBook(SetModelID).Activate
            'DSARejectsForm.Show

        End Sub
        Public Shared Sub ClearRejectReport(ModelID As Integer)

            GetWorkBook(ModelID).Range("DSARejects").ClearContents

        End Sub

        Private Shared Function DSA_Import(SetModelID As Integer, FileToOpen As String, ImportType As String) As Boolean

            Dim BusPlanFile As IWorkbook
            Dim DSAFile As IWorkbook
            Dim ResponseMessage As Integer
            Dim InitialCheckCount As Integer
            Dim wsCheck As DevExpress.Spreadsheet.Worksheet
            Dim VarDate As Integer

            BusPlanFile = GetWorkBook(SetModelID)
            Dim MyFileInfos As IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FileToOpen)
            Dim FileTrans As AbovoTransaction = FileManager.OpenModel(
                FileToOpen,
                MyFileInfos,
                FileManager.WorkbookOpenMode.ImportSource)

            If FileTrans.BError Then
                MsgBox(FileTrans.StrResponseMessage)
                Return False
            End If

            Dim DSAId As Integer = FileTrans.IntegerReturn
            Try
                DSAFile = GetWorkBook(DSAId)

            If Not DSAFile.Worksheets.Contains("Unit Handovers & Sales") Then

                MsgBox("File " & FileToOpen & " does not appear to be an Abovo DSA model.")
                Return False

            End If

            wsCheck = DSAFile.Worksheets("Global Assumptions")

            If DSAFile.Range("ModelVersion")(0, 0).Value.NumericValue < EarliestVersion Then

                MsgBox("DSA  " & FileToOpen & " is earlier than the earliest possible import version of " & EarliestVersion.ToString)
                Return False

            End If

            ' If is a DSA model, and not a Consol model, of a late enough version,
            ' then record the Check Sheet's error count for later comparison.

            If ImportType <> "Consol" Then

                InitialCheckCount = DSAFile.Range("CheckTotal")(0, 0).Value.NumericValue

            Else

                InitialCheckCount = 0

            End If

            Dim NewDSAName As String

            VarDate = BusPlanFile.Range("TransferDate").Value.NumericValue

            DSAFile.Range("BplanYear").Value = VarDate

            DSAFile.CalculateFull()

            NewDSAName = DSAFile.Range("SchemeName").Value.TextValue

            NewDSAName = NewDSAName.Replace("/", "_")
            NewDSAName = NewDSAName.Replace("*", "_")
            NewDSAName = NewDSAName.Replace("?", "_")
            NewDSAName = NewDSAName.Replace("'/'", "_")

            'NewFileName = ActiveWorkbook.Name

            CheckStatus = "Stop"

            Dim CommitmentStatus As String

            If CheckFile(BusPlanFile, DSAFile, ImportType, InitialCheckCount, NewDSAName) Then

                If ImportType = "Consol" Then

                    ResponseMessage = MsgBox("Does this model contain only Committed schemes?" & Chr(13) & Chr(10) & "Please select No if they are uncommitted.", vbYesNo, Title:="Committed or uncommitted")

                    If ResponseMessage = 7 Then ' "No" response

                        CommitmentStatus = "Uncommitted"

                    Else

                        CommitmentStatus = "Committed"

                    End If

                Else

                    CommitmentStatus = DSAFile.Range("Rep_Global_01").Value.TextValue

                End If

                ' reveal hidden sheets
                BusPlanFile.Worksheets("Hidden - Imports Start").Visible = True
                BusPlanFile.Worksheets("Hidden - Tenure Totals Start").Visible = True
                BusPlanFile.Worksheets("Hidden - Tenure Totals End").Visible = True

                If InsertNewSheet(BusPlanFile, DSAFile, ImportType, NewDSAName) Then

                    If InsertNewTotals(BusPlanFile, SetModelID, DSAFile, ImportType, NewDSAName) Then

                        UpdateList(BusPlanFile, SetModelID, DSAFile, ImportType, NewDSAName, FileToOpen, CommitmentStatus)
                        DSA_Import = True

                    Else

                        ''Error in total sheet

                    End If

                Else

                    ''Error in insert new sheet

                End If

                ' conceal hidden sheets
                BusPlanFile.Worksheets("Hidden - Imports Start").Visible = False
                BusPlanFile.Worksheets("Hidden - Tenure Totals Start").Visible = False
                BusPlanFile.Worksheets("Hidden - Tenure Totals End").Visible = False

            End If

            Finally
                CloseModel(DSAId)
            End Try

            Return DSA_Import

        End Function


        '        Private Sub ShowDSARejects()

        '            DSARejectsForm.Show

        '        End Sub

        Private Shared Function CheckFile(BusPlanFile As IWorkbook, DSAFile As IWorkbook, ModelType As String, InitialCheckCount As Integer, NewDSAName As String) As Boolean

            On Error Resume Next

            If ModelType = "Consol" Then

                If BusPlanFile.Range("ModelVersion").Value.NumericValue < EarliestConsol Then

                    MsgBox("Please use a consolidated DSA model no earlier than Version " & EarliestConsol)

                ElseIf Left(DSAFile.Worksheets("Global Assumptions").Range("A1").Value.TextValue, 16) <> "Consolidated DSA" Then

                    MsgBox("This does not appear to be a consolidated DSA model.")

                    'ElseIf IsError(WorksheetFunction.Match(NewDSAName, BusPlanFile.Worksheets("List Imported").Range("ImportedSchemes"), 0)) Then

                    '    CheckStatus = "Go"

                Else

                    MsgBox("Consolidated scheme " & NewDSAName & " has already been imported")

                End If

            Else    ' ModelType is Single or Folder

                If DSAFile.Range("CheckTotal").Value.NumericValue > 0 Then   ' check for outstanding warnings on the "Check sheet"

                    If InitialCheckCount > 0 Then   ' errors were there at the outset

                        If ModelType = "Single" Then

                            MsgBox("There are errors shown on the Check Sheet of the DSA model.")

                        ElseIf ModelType = "Folder" Then

                            Call RejectReport(BusPlanFile, "Errors in DSA model", NewDSAName)

                        End If

                    Else ' a check error has been created by using the BP's start date

                        If ModelType = "Single" Then

                            MsgBox("Dates too early in DSA model - please contact Abovo rep.")

                        ElseIf ModelType = "Folder" Then

                            Call RejectReport(BusPlanFile, "Dates too early, please contact Abovo rep", NewDSAName)

                        End If

                    End If

                    'ElseIf IsError(WorksheetFunction.Match(NewDSAName, ThisWorkbook.Sheets("List Imported").Range("ImportedSchemes"), 0)) Then

                    '    CheckStatus = "Go"

                Else

                    If ModelType = "Single" Then

                        MsgBox("Scheme '" & NewDSAName & "' has already been imported")

                    ElseIf ModelType = "Folder" Then

                        Call RejectReport(BusPlanFile, "Scheme already imported", NewDSAName)

                    End If


                End If

            End If



        End Function
        Private Shared Sub RejectReport(BPFile As IWorkbook, RejectType As String, NewDSAName As String)

            Dim RangeRows As Integer

            On Error Resume Next

            RangeRows = BPFile.Range("DSARejects").RowCount
            BPFile.Range("DSARejects")(RangeRows, 0).Value = NewDSAName & "  -  " & RejectType & "  -  Filename: " & NewDSAName
            BPFile.Range("DSARejects").Resize(RangeRows + 1, BPFile.Range("DSARejects").ColumnCount).Name = "DSARejects"

        End Sub


        Private Shared Function InsertNewSheet(BusPlanFile As IWorkbook, DSAFile As IWorkbook, ModelType As String, NewDSAName As String) As Boolean

            On Error GoTo Err_Handler

            If BusPlanFile.Worksheets.Contains(NewDSAName) Then

                On Error Resume Next
                BusPlanFile.Worksheets.Remove(BusPlanFile.Worksheets(NewDSAName))

            End If

            Dim NewWS As DevExpress.Spreadsheet.Worksheet

            Dim HTTSIndex As Integer = BusPlanFile.Worksheets("Hidden - Tenure Totals Start").Index

            BusPlanFile.Worksheets.Insert(HTTSIndex, NewDSAName)

            NewWS = BusPlanFile.Worksheets(NewDSAName)

            If ModelType <> "Consol" Then

                DSAFile.Worksheets("Cashflow Selector").Range("B6").Value = "All"
                DSAFile.Worksheets("Cashflow Selector").Range("B7").Value = "All"
                DSAFile.Calculate()

            End If

            Dim Sourcerange As DevExpress.Spreadsheet.CellRange = DSAFile.Worksheets("Hidden - Export Sheet").GetUsedRange

            NewWS.Range("A1").CopyFrom(Sourcerange, PasteSpecial.Values)
            NewWS.Range("A1").CopyFrom(Sourcerange, PasteSpecial.Formats)

            BusPlanFile.CalculateFull()

            Return True

            Exit Function

Exiter:

            Exit Function

Err_Handler:

            On Error Resume Next
            BusPlanFile.Worksheets.Remove(BusPlanFile.Worksheets(NewDSAName))
            MsgBox("Could not add DSA sheet " & NewDSAName & ", sorry. Error: " & Err.Description)

            Resume Exiter

        End Function
        Private Shared Function InsertNewTotals(BusPlanFile As IWorkbook, BPID As Integer, DSAFile As IWorkbook, ModelType As String, NewDSAName As String) As Boolean

            On Error GoTo Err_Handler
            If bIsDevelopment Then On Error GoTo 0

            Dim TotalsWSCreated As Worksheet
            Dim rngNewsCells As DevExpress.Spreadsheet.CellRange
            Dim NewTotalsSheetName As String

            UNProtectWS(BPID, "Hidden - Tenure Totals Start")

            BusPlanFile.Worksheets("Hidden - Tenure Totals End").Visible = True
            BusPlanFile.Worksheets("Hidden - Tenure Totals Start").Visible = True
            BusPlanFile.Worksheets("Hidden - Imports Start").Visible = True

            NewTotalsSheetName = NewDSAName & " Total"

            If BusPlanFile.Worksheets.Contains(NewTotalsSheetName) Then

                On Error Resume Next
                BusPlanFile.Worksheets.Remove(BusPlanFile.Worksheets(NewTotalsSheetName))

            End If

            Dim HTTSIndex As Integer = BusPlanFile.Worksheets("Hidden - Tenure Totals End").Index

            BusPlanFile.Worksheets.Insert(HTTSIndex, NewTotalsSheetName)


            TotalsWSCreated = BusPlanFile.Worksheets(NewTotalsSheetName)

            Dim SourceRange As DevExpress.Spreadsheet.CellRange

            SourceRange = BusPlanFile.Worksheets("Hidden - Tenure Totals Start").GetUsedRange


            TotalsWSCreated.Range("A:1").CopyFrom(SourceRange, PasteSpecial.ColumnWidths)
            TotalsWSCreated.Range("A:1").CopyFrom(SourceRange, PasteSpecial.Formats)
            TotalsWSCreated.Range("A:1").CopyFrom(SourceRange, PasteSpecial.NumberFormats)
            TotalsWSCreated.Range("A:1").CopyFrom(SourceRange, PasteSpecial.Formulas)

            TotalsWSCreated.Range("A10").Select


            TotalsWSCreated.Range("A10").Value = NewDSAName


            BusPlanFile.CalculateFullRebuild()

            Dim SearchTerm As String = "Hidden - Imports Start"
            Dim ForumlaText As String
            Dim options As DevExpress.Spreadsheet.SearchOptions = New DevExpress.Spreadsheet.SearchOptions()

            options.SearchBy = SearchBy.Rows
            options.SearchIn = SearchIn.Formulas
            options.MatchEntireCellContents = False

            Dim searchResult As IEnumerable(Of Cell) = TotalsWSCreated.Search(SearchTerm, options)

            For Each cell As Cell In searchResult

                ForumlaText = cell.Formula.ToString()
                ForumlaText = ForumlaText.Replace("Hidden - Imports Start", NewDSAName)
                cell.Formula = ForumlaText

            Next cell

            SearchTerm = "Hidden - Tenure Totals Start"

            Dim searchResult2 As IEnumerable(Of Cell) = TotalsWSCreated.Search(SearchTerm, options)

            For Each cell As Cell In searchResult2

                ForumlaText = cell.Formula.ToString()
                ForumlaText = ForumlaText.Replace("Tenure Totals Start", NewTotalsSheetName)
                cell.Formula = ForumlaText

            Next cell

            ProtectWS(BPID, "Hidden - Tenure Totals Start")
            ProtectWS(BPID, NewTotalsSheetName)

            BusPlanFile.CalculateFull()

            Return True

            Exit Function

Exiter:

            'If Application.Calculation = xlCalculationManual Then Application.Calculation = xlCalculationAutomatic

            Exit Function

Err_Handler:

            On Error Resume Next

            MsgBox("Could not add DSA total sheet, sorry. Error: " & Err.Description)

            Resume Exiter

        End Function



        Private Shared Sub UpdateList(BusPlanFile As IWorkbook, BPID As Integer, DSAFile As IWorkbook, ImportType As String, NewDSAName As String, FileToOpen As String, Commitment As String)

            On Error GoTo Err_Handler

            Dim NewRowIndex As Integer

            UNProtectWS(BPID, "List Imported")

            NewRowIndex = BusPlanFile.Range("ImportedSchemes").RowCount

            BusPlanFile.Range("ImportedSchemes")(NewRowIndex, 0).Value = NewDSAName
            BusPlanFile.Range("ImportedSchemes")(NewRowIndex, 1).Value = Now()
            BusPlanFile.Range("ImportedSchemes")(NewRowIndex, 2).Value = FileToOpen
            BusPlanFile.Range("ImportedSchemes")(NewRowIndex, 3).Value = Commitment
            BusPlanFile.Range("ImportedSchemes")(NewRowIndex, 5).Formula = "=IF(ConsolOption=1,""No"",IF(AND(RC[-2]<>""Committed"",ExclUnCommitDvpt=""TRUE""),""No"",IF(RC[-1]=""No"",""No"",""Yes"")))"
            BusPlanFile.Range("ImportedSchemes")(NewRowIndex, 4).Value = "Yes"

            Dim CurrCell As DevExpress.Spreadsheet.Cell = BusPlanFile.Range("ImportedSchemes")(NewRowIndex, 4)

            Dim validation As DataValidation = BusPlanFile.Range("ImportedSchemes").Worksheet.DataValidations.Add(BusPlanFile.Range("ImportedSchemes").Worksheet(CurrCell.GetReferenceA1), DataValidationType.List, "Yes, No")

            BusPlanFile.Range("ImportedSchemes").Resize(NewRowIndex, BusPlanFile.Range("ImportedSchemes").ColumnCount).Name = "ImportedSchemes"
            BusPlanFile.Range("Appraisal1").Resize(NewRowIndex, BusPlanFile.Range("Appraisal1").ColumnCount).Name = "Appraisal1"
            BusPlanFile.Range("Appraisal2").Resize(NewRowIndex, BusPlanFile.Range("Appraisal2").ColumnCount).Name = "Appraisal2"

            BusPlanFile.Range("NumDSASchemes").Value = NewRowIndex + 1

            CopySummaryFormulae(BusPlanFile, NewRowIndex, NewDSAName)

            BusPlanFile.Worksheets("Development Consol Options").Range("SchemeLink").Value = 1  ' set drop-down used for deletion to top, empty value

Exiter:

            Exit Sub

Err_Handler:

            MsgBox("Could not add DSA to imported list, sorry. Error: " & Err.Description)
            Resume Exiter

        End Sub
        Private Shared Sub CopySummaryFormulae(BPFile As IWorkbook, NewRowIndex As Integer, NewDSAName As String)

            BPFile.CalculateFull()

            Dim RowA1Address As String = (NewRowIndex + 1).ToString

            BPFile.Range("ImportedSchemes")(NewRowIndex, 12).Formula = "=SUM('" & NewDSAName & "'!D$161:D$201)"
            BPFile.Range("ImportedSchemes")(NewRowIndex, 23).Formula = "=SUM(K" & RowA1Address & ":U" & RowA1Address & ")"

            Dim TargetWS As DevExpress.Spreadsheet.Worksheet = BPFile.Range("ImportedSchemes").Worksheet

            Dim TargetRange As DevExpress.Spreadsheet.CellRange = TargetWS.Range.FromLTRB(12, NewRowIndex, 22, NewRowIndex)

            TargetRange.CopyFrom(BPFile.Range("ImportedSchemes")(NewRowIndex, 12), PasteSpecial.Formulas)
            BPFile.Calculate()
            TargetRange.CopyFrom(TargetRange, PasteSpecial.Values)

            BPFile.Range("ImportedSchemes")(NewRowIndex, 24).Value = BPFile.Range("TransferDate").Value
            BPFile.Range("ImportedSchemes")(NewRowIndex, 25).Formula = "=IF(RC[-1]<>TransferDate,""WARNING: Date does not match current Business Plan Start Date"","""")"

        End Sub

        'Sub GoToDSA()

        '    CodeSafe JW 26/4/22

        '            Dim strSchemeNm As String
        '    Dim lnk As Range
        '    Dim ws As Worksheet

        '    InitiateCodeRun "GoToDSA"

        '    Set lnk = Range("SchemeLink")
        '    strSchemeNm = Sheets("List Imported").Range("ImportedSchemes").Cells(lnk.Value).Value

        '    Loop thru sheets to delete each one beginning with strSchemeNm
        '            For Each ws In Worksheets

        '        If ws.Name = strSchemeNm Then

        '            ws.Select

        '        End If

        '    Next ws

        '    Range("SchemeLink").Value = "0"

        '    Call fncPrtSht("List Imported")

        '    ResumeCodeRun "GoToDSA"

        '    EndCodeRun


        'End Sub
        '        Function RngNameExists(strName As String) As Boolean

        '            ' check for existence of a named range
        '            Dim strRngName As Variant

        '            For Each strRngName In ThisWorkbook.Names

        '                If strRngName.NameLocal = strName Then

        '                    RngNameExists = True
        '                    Exit Function

        '                Else

        '                    RngNameExists = False

        '                End If

        '            Next strRngName

        '        End Function
        '        Function GetFolder(Optional Message) As String

        '            ' select a folder for copying from or saving to
        '            Dim PathDialogue As FileDialog
        '            Dim SelectedPath As String

        '    Set PathDialogue = Application.FileDialog(msoFileDialogFolderPicker)

        '    With PathDialogue

        '                .Title = Message

        '                .AllowMultiSelect = False
        '                '        .InitialFileName = Application.DefaultFilePath

        '                If .Show <> -1 Then GoTo NextCode

        '                SelectedPath = .SelectedItems(1)

        '            End With

        'NextCode:

        '            GetFolder = SelectedPath
        '    Set PathDialogue = Nothing

        'End Function
        '        '
        '        ' **************************
        '        ' TEMPLATE IMPORT PROCEDURES
        '        ' **************************
        '        '
        '        Sub DSA_Template()

        '            ' import data from a development template file
        '            Dim FileToOpen
        '            Dim NumSchemes As Integer
        '            Dim x As Integer, y As Integer, FirstColumn As Integer, LastColumnToClear As Integer
        '            Dim CalcMethod As Integer

        '            InitiateCodeRun "DSA_Template", True

        '    Sheets("Development BP Assumptions").Activate

        '            ' Open and assign Source Data File
        '            FileToOpen = Application.GetOpenFilename(fileFilter:="Excel Files (*.xls*), *.xls*", Title:="Select template file")

        '            If FileToOpen <> False Then

        '                Workbooks.Open(FileToOpen)
        '        Set NewDSAFile = ActiveWorkbook

        '        With Sheets("Hidden - Workings")

        '                    .Visible = True
        '                    .Activate

        '                End With

        '                NumSchemes = Range("SchemeCount").Value

        '                Sheets("Dvpt Export").Activate

        '                ThisWorkbook.Activate

        '                '   Assign column number of starting point for insertion
        '                FirstColumn = Range("LastIDColNum").Column
        '                x = FirstColumn
        '                '   Assign number of columns to be inserted
        '                y = NumSchemes  ' - Range("HouseTypeInID").Columns.Count

        '                If MsgBox("Do you wish to clear Development BP Assumptions before import?" & Chr(13) & Chr(13) & Chr(10) & "Otherwise new columns will be added.", vbYesNo, Title:="Clear or add columns") = 6 Then ' "Yes" response

        '                    If NumSchemes > (x - HeadingColumns - 1) Then   ' replacement scheme details need more columns

        '                        Application.StatusBar = "Adding columns"
        '                        Call InsertDvptColumns(NumSchemes - (x - HeadingColumns) + 1, "Null", "Auto")
        '                        ResumeCodeRun "DSA_Template"

        '                Application.StatusBar = "Clearing rows"

        '                        LastColumnToClear = Range("LastIDColNum").Column - HeadingColumns - 1
        '                        Call ClearDvptColumns(LastColumnToClear)
        '                        ResumeCodeRun "DSA_Template"


        '            ElseIf NumSchemes < (x - HeadingColumns - 1) Then   ' replacement scheme details need fewer columns

        '                        Application.StatusBar = "Removing columns"
        '                        Call DeleteDvptColumns(x - 1, (x - HeadingColumns - NumSchemes - 1), "Auto")
        '                        ResumeCodeRun "DSA_Template"

        '                Application.StatusBar = "Clearing rows"
        '                        LastColumnToClear = Range("LastIDColNum").Column - HeadingColumns - 1

        '                        Call ClearDvptColumns(LastColumnToClear)
        '                        ResumeCodeRun "DSA_Template"


        '            Else  ' replacement scheme details need the same number of columns

        '                        Application.StatusBar = "Clearing rows"
        '                        LastColumnToClear = Range("LastIDColNum").Column - HeadingColumns - 1
        '                        Call ClearDvptColumns(LastColumnToClear)
        '                        ResumeCodeRun "DSA_Template"


        '            End If

        '                    ' now paste in data
        '                    x = HeadingColumns + 1
        '                    Application.StatusBar = "Pasting data"

        '                    Call TemplateData(x, NumSchemes)
        '                    ResumeCodeRun "DSA_Template"


        '        Else    ' retain data and add columns

        '                    Application.StatusBar = "Adding columns"
        '                    Call InsertDvptColumns(y, "Null", "Auto")
        '                    ResumeCodeRun "DSA_Template"

        '            Application.StatusBar = "Pasting data"

        '                    Call TemplateData(FirstColumn, NumSchemes)
        '                    ResumeCodeRun "DSA_Template"


        '        End If

        '                Call DvptTemplDateStamp()
        '                ResumeCodeRun "DSA_Template"

        '        NewDSAFile.Close SaveChanges:=False

        '        MsgBox "Import complete"

        '    End If

        '            Call fncPrtSht("Development BP Assumptions")
        '            ResumeCodeRun "DSA_Template"

        '    EndCodeRun True

        'End Sub
        '        Sub TemplateData(x As Integer, y As Integer)

        '            'CodeSafe JW 26/4/22

        '            ' copy and paste data from the template file
        '            Dim DataRows As Integer, FirstColumn As Integer, NextColumn As Integer
        '            Dim i As Integer, j As Integer, NumSchemes As Integer, RangeRows As Integer
        '            Dim FirstDataCell As String, NextRange As String

        '            InitiateCodeRun "TemplateData"

        '    FirstDataCell = "A6"
        '            FirstColumn = x
        '            NumSchemes = y

        '            NewDSAFile.Activate
        '            DataRows = Range("TemplateWorkings").Rows.Count

        '            For i = 1 To DataRows

        '                NewDSAFile.Activate
        '                Sheets("Hidden - Workings").Activate
        '                NextRange = Range("TemplateWorkings").Cells(i, 2).Value
        '                NextColumn = Range("TemplateWorkings").Cells(i, 4).Value
        '                RangeRows = Range("TemplateWorkings").Cells(i, 5).Value

        '                Sheets("Dvpt Export").Activate
        '                Range(Range(FirstDataCell).Cells(1, NextColumn), Range(FirstDataCell).Cells(NumSchemes, NextColumn + RangeRows - 1)).Copy

        '                ThisWorkbook.Activate
        '                Range(NextRange).Cells(1, FirstColumn - HeadingColumns).Select
        '                Selection.PasteSpecial Paste:=xlValues, Transpose:=True

        '        Application.CutCopyMode = False

        '            Next i

        '            EndCodeRun


        '        End Sub
        '        Sub ClearDvptColumns(y As Integer)

        '            'CodeSafe JW 26/4/22

        '            ' clear scheme-specific ranges for identified schemes on "Development BP Assumptions" sheet
        '            InitiateCodeRun "ClearDvptColumns"

        '    Dim i As Integer, NumberToClear As Integer, RowsToClear As Integer
        '            Dim RangeToClear As String

        '            NewDSAFile.Activate
        '            NumberToClear = Range("RangesToClear").Rows.Count

        '            ThisWorkbook.Activate

        '            With Sheets("Development BP Assumptions")

        '                .Activate
        '                .Unprotect Password:=PW

        '    End With

        '            For i = 1 To NumberToClear

        '                NewDSAFile.Activate
        '                RangeToClear = Range("RangesToClear").Cells(i).Value
        '                ThisWorkbook.Activate
        '                RowsToClear = Range(RangeToClear).Rows.Count
        '                Range(Range(RangeToClear).Cells(1, 1), Range(RangeToClear).Cells(RowsToClear, y)).Select
        '                Selection.ClearContents

        '            Next i

        '            EndCodeRun

        '        End Sub
        '        Sub DvptTemplDateStamp()

        '            'CodeSafe JW 26/4/22

        '            ' record filename of template and time of import

        '            On Error Resume Next

        '            InitiateCodeRun "DvptTemplDateStamp"

        '    ThisWorkbook.Activate
        '            Sheets("Global Assumptions").Activate
        '            Range("DvptTemplFile") = NewDSAFile.Path & "\" & NewDSAFile.Name
        '            Range("DateStampDvptTempl").Select
        '            ActiveCell.Value = Now()

        '            Selection.Copy
        '            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        '    EndCodeRun

        '        End Sub




        Sub InsertNewDSASheetLegMethod(SetModelID As Integer, SchemeName As String)

            On Error GoTo Err_Handler

            Dim NewName As String = SchemeName & " Scheme"

            Dim BusPlanFile As DevExpress.Spreadsheet.IWorkbook = GetWorkBook(SetModelID)

            If BusPlanFile.Worksheets.Contains(SchemeName) Then

                MsgBox("A Worksheet named " & NewName & " already exists.")
                GoTo Exiter
            Else

                'Worksheets(ActiveSheet.Name).Copy After:=Worksheets(ActiveSheet.Index)
                'ActiveSheet.Name = NewName

            End If



            BusPlanFile.Range("ClearSch1").ClearContents
            BusPlanFile.Range("ClearSch2").ClearContents
            BusPlanFile.Range("ClearSch3").ClearContents
            BusPlanFile.Range("ClearSch4").ClearContents
            BusPlanFile.Range("ClearSch5").ClearContents
            BusPlanFile.Range("ClearSch6").ClearContents
            BusPlanFile.Range("ClearSch7").ClearContents
            BusPlanFile.Range("ClearSch8").ClearContents
            BusPlanFile.Range("ClearSch9").ClearContents
            BusPlanFile.Range("ClearSch10").ClearContents
            BusPlanFile.Range("ClearSch11").ClearContents
            BusPlanFile.Range("ClearSch12").ClearContents
            BusPlanFile.Range("ClearSch13").ClearContents
            BusPlanFile.Range("ClearSch14").ClearContents
            BusPlanFile.Range("ClearSch15").ClearContents
            BusPlanFile.Range("ClearSch16").ClearContents
            BusPlanFile.Range("ClearSch17").ClearContents
            BusPlanFile.Range("ClearSch18").ClearContents
            BusPlanFile.Range("ClearSch19").ClearContents
            BusPlanFile.Range("ClearSch20").ClearContents

Exiter:

            Exit Sub

Err_Handler:


            MsgBox("Sorry, an error occured during procedure. The error has been logged.  The error is " & Err.Description)


            Err.Clear()

            Resume Exiter

        End Sub


    End Class

End Namespace
