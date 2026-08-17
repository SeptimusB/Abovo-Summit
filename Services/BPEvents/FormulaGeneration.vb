Imports DevExpress.XtraRichEdit.Model
Imports Microsoft.Office.Interop.Excel
Imports Abovo.WSSecurity
Imports Abovo.AbovoAppCls

Namespace Abovo
    Public Class FormulaGeneration


        Public Shared Sub ExecuteFormulaGeneration(SetModelID As Integer, ProtectOrNot As Boolean, SetTransaction As AbovoTransaction)
            Dim EventTransaction As AbovoTransaction
            EventTransaction = SetTransaction

            Dim FormulaGenTransaction As New AbovoTransaction("ExecuteFormulaGeneration")
            Dim BPModelID As Integer = SetModelID

            'Main procedure to copy formulae only to columns where category headings are in place

            On Error GoTo Err_Handler
            Dim ActiveSheet As DevExpress.Spreadsheet.Worksheet = Nothing

            Dim ArgDetails(0 To 5) As String
            Dim DownArgsUsed As Integer
            Dim AcrossArgs As Integer
            Dim DownArgs As Integer
            Dim CheckRange
            Dim CallingSheet
            Dim MaxBlocks As Integer



            '            'Note current sheet at the time the procedure was called
            '            CallingSheet = ActiveSheet.Name

            '            'Go to, unhide and unprotect sheet listing sheet names and locations

            '            If Sheets("Hidden - Sheet Lists").Visible = False Then

            '                Sheets("Hidden - Sheet Lists").Visible = True

            '            End If

            '            Sheets("Hidden - Sheet Lists").Select
            '            ActiveSheet.Unprotect Password:=PW

            '    'Maximum number of blocks on each sheet

            '            MaxBlocks = ActiveCell.SpecialCells(xlLastCell).Column - Range("TopStockSheets").Column

            '            'Count and assign the number of categorical series (eg, stock types) to be used
            '            DownArgsUsed = WorksheetFunction.CountA(Range("ColumnList"))

            '            For DownArgs = 1 To DownArgsUsed    'For each categorical series

            '                For AcrossArgs = 0 To 5     'Read and assign the details for each categorical series

            '                    ArgDetails(AcrossArgs) = Range("TopColumnList").Offset(DownArgs, AcrossArgs)

            '                Next AcrossArgs

            '                'Call the main procedure (separately for each series) supplying the above details as arguments
            '                ' only if there has been a change for that series, but calling for Repairs if Archetypes have changed.
            '                If Range("TopColumnList").Offset(DownArgs, 6) <> 0 Or (DownArgs = 3 And Range("TopColumnList").Offset(8, 6) <> 0) Then

            '                    Call FillColumns(ArgDetails(0), ArgDetails(1), ArgDetails(2), ArgDetails(3), ArgDetails(4), ArgDetails(5), MaxBlocks, ProtectOrNot)
            '                    ResumeCodeRun "AllColumns"

            '            Range("TopColumnList").Offset(DownArgs, 5) = ArgDetails(5)
            '                    Range("TopColumnList").Offset(DownArgs, 6) = 0

            '                End If

            '            Next DownArgs


            '            'Call the procedure to remove formulae for unused stock-condition archetypes
            '            CheckRange = Range("TopColumnList").Offset(DownArgsUsed + 1, 5).Value

            '            Call CutArchetypes(CheckRange)
            '            ResumeCodeRun "AllColumns"

            '    Range("TopColumnList").Offset(DownArgsUsed + 1, 5) = CheckRange
            '            Range("TopColumnList").Offset(DownArgsUsed + 1, 6) = 0

            '            Call DeleteFormGenMenu
            '            ResumeCodeRun "AllColumns"

            '    'Hide and protect sheet listing sheet names and locations
            '            Sheets("Hidden - Sheet Lists").Visible = False

            If ProtectOrNot Then

                ProtectWS(BPModelID, ActiveSheet.Name)
                '                Call fncPrtSht()
                '                ResumeCodeRun "AllColumns"

            End If

            '            'Return to original sheet
            '            Sheets(CallingSheet).Select

            '            DoEvents

            '            Application.CalculateFullRebuild()

            '            DoEvents

            '            EndCodeRun True
Exiter:

            Exit Sub
Err_Handler:
            '            Application.Calculation = xlCalculationAutomatic
            '            MsgBox "There was an error in Forumla Generation - please save your file with a different name and contact your Abovo representative."
        End Sub
        '        Sub FillColumns(CatRange, SheetRange, FirstCat, CatSheet, CatOrientation, CheckRange, MaxBlocks, ProtectOrNot)
        '            'Carry out the copying and pasting of formulae on the sheets selected by the "AllColumns" procedure

        '            Dim CatsUsed As Integer
        '            Dim SheetsUsed As Integer
        '            Dim CatsCount As Integer
        '            Dim CatsCode As String
        '            Dim ColCount As Integer
        '            Dim SheetCount As Integer
        '            Dim BlockCount As Integer
        '            Dim BlocksUsed As Integer
        '            Dim SheetName As String
        '            Dim CellAddress As String
        '            Dim HiddenSheet As Boolean

        '            InitiateCodeRun "FillColumns"
        '    'Count and assign number of used categories in series and the number of listed sheets for this series
        '            CatsUsed = WorksheetFunction.CountA(Range(CatRange))
        '            SheetsUsed = WorksheetFunction.CountA(Range(SheetRange))

        '            For SheetCount = 1 To SheetsUsed    'for each listed sheet

        '                Sheets("Hidden - Sheet Lists").Select

        '                'Assign the value of the sheet name, offset from the cell immediately above the list of sheet names
        '                SheetName = Range("Top" & SheetRange).Offset(SheetCount, 0).Value
        '                'Count and assign the number of starting-cell addresses listed to the right of the sheet name
        '                BlocksUsed = WorksheetFunction.CountA(Range(Range("Top" & SheetRange).Offset(SheetCount, 1), Range("Top" & SheetRange).Offset(SheetCount, MaxBlocks)))

        '                'Unhide hidden sheet and note if it was previously hidden
        '                If Sheets(SheetName).Visible = False Then
        '                    Sheets(SheetName).Visible = True
        '                    HiddenSheet = True
        '                Else
        '                    HiddenSheet = False
        '                End If

        '                'Select and unprotect sheet
        '                Sheets(SheetName).Select
        '                ActiveSheet.Unprotect Password:=PW

        '        For BlockCount = 1 To BlocksUsed    'for each of the cell addresses

        '                    'Initialise counts of categories and columns
        '                    If WorksheetFunction.CountA(Range(FirstCat)) = 0 And CatsUsed >= 1 Then
        '                        CatsCount = 0
        '                        CatsCode = "o"
        '                    Else
        '                        CatsCount = 1
        '                        CatsCode = "i"
        '                    End If
        '                    ColCount = 1

        '                    'Assign next starting-cell address
        '                    CellAddress = Range("Top" & SheetRange).Offset(SheetCount, BlockCount).Value

        '                    'Call clearing procedure, passing series details as arguments
        '                    Call ClearColumns(CatSheet, CatRange, SheetName, CellAddress, CatOrientation)
        '                    ResumeCodeRun "FillColumns"

        '            'Go to current sheet and cell address, drop down column and copy
        '                    Sheets(SheetName).Select
        '                    Range(CellAddress).Select
        '                    Range(Selection, Selection.End(xlDown)).Select
        '                    Selection.Copy
        '                    'Paste column to as many new columns as required
        '                    Do While CatsCount < CatsUsed
        '                        If (CatOrientation = "Horizontal" And WorksheetFunction.CountA(Range(FirstCat).Offset(0, ColCount))) _
        '                Or (CatOrientation = "Vertical" And WorksheetFunction.CountA(Range(FirstCat).Offset(ColCount, 0))) Then
        '                            Range(CellAddress).Offset(0, ColCount).Select
        '                            ActiveSheet.Paste
        '                            CatsCount = CatsCount + 1
        '                            CatsCode = CatsCode & "i"
        '                        Else
        '                            CatsCode = CatsCode & "o"
        '                        End If
        '                        ColCount = ColCount + 1
        '                    Loop
        '                Next BlockCount
        '                'Protect sheet and hide if previously hidden
        '                If ProtectOrNot = "Protect" Then

        '                    Call fncPrtSht(ActiveSheet.Name)

        '                    If HiddenSheet = True Then
        '                        Sheets(SheetName).Visible = False
        '                    End If

        '                    ResumeCodeRun "FillColumns"

        '        End If
        '            Next SheetCount

        '            'Record code of categories used
        '            CheckRange = CatsCode
        '            EndCodeRun

        '        End Sub
        '        Sub ClearColumns(CatSheet, CatRange, SheetName, CellAddress, CatOrientation)
        '            ' Clear redundant columns of formulae
        '            Dim ColsUsed As Integer
        '            '
        '            InitiateCodeRun "ClearColumns"
        '    'Go to category range
        '            Sheets(CatSheet).Select

        '            'Count all columns available for category headings
        '            If CatOrientation = "Horizontal" Then   'if category headings are arranged horizontally
        '                ColsUsed = Range(CatRange).Columns.Count - 1
        '            Else    'category heading must be arranged vertically
        '                ColsUsed = Range(CatRange).Rows.Count - 1
        '            End If

        '            'Go to starting cell on current sheet and clear as many columns to the right as have been counted above
        '            Sheets(SheetName).Select

        '            Range(Range(CellAddress).Offset(0, 1), Range(CellAddress).End(xlDown).Offset(0, ColsUsed)).Select

        '            Selection.Clear

        '            EndCodeRun

        '        End Sub
        '        Sub CutArchetypes(CheckRange)

        '            'CodeSafe JW 26/4/22

        '            ' Clear formulae for redundant stock-condition archetypes
        '            ' called by All_Columns procedure

        '            Dim CatsUsed As Integer
        '            Dim CatsCount As Integer
        '            Dim ColCount As Integer
        '            Dim SheetsUsed As Integer
        '            Dim SheetCount As Integer
        '            Dim BlockCount As Integer
        '            Dim BlocksUsed As Integer
        '            Dim SheetName As String
        '            Dim CellAddress As String
        '            Dim CatsCode As String
        '            Dim HiddenSheet As Boolean

        '            InitiateCodeRun "CutArchetypes"

        '    'Count and assign number of archetypes used and the number of stock condition sheets
        '            ' excluding the last sheet (VAT Shelter Calculation)
        '            CatsUsed = WorksheetFunction.CountA(Range("Archetype"))

        '            SheetsUsed = WorksheetFunction.CountA(Range("StockCondSheets")) - WorksheetFunction.CountA(Range("NonArchetype"))

        '            For SheetCount = 1 To SheetsUsed    'for each listed sheet

        '                Sheets("Hidden - Sheet Lists").Select
        '                'Assign the value of the sheet name, offset from the cell immediately above the list of sheet names
        '                SheetName = Range("TopStockCondSheets").Offset(SheetCount, 0).Value
        '                'Count and assign the number of starting-cell addresses listed to the right of the sheet name
        '                BlocksUsed = Range("Archetype").Rows.Count

        '                'Unhide hidden sheet and note if it was previously hidden
        '                If Sheets(SheetName).Visible = False Then

        '                    Sheets(SheetName).Visible = True
        '                    HiddenSheet = True

        '                Else

        '                    HiddenSheet = False

        '                End If

        '                'Select and unprotect sheet
        '                Sheets(SheetName).Select
        '                ActiveSheet.Unprotect Password:=PW

        '        For BlockCount = 1 To BlocksUsed 'for each of the cell addresses

        '                    If WorksheetFunction.CountA(Range("Archetype1").Offset(BlockCount - 1, 0)) = 0 Then

        '                        'Assign next starting-cell address
        '                        CellAddress = Range("TopStockCondSheets").Offset(SheetCount, BlockCount).Value
        '                        'Call clearing procedure, passing series details as arguments
        '                        Call ClearArchetypes(SheetName, CellAddress)
        '                        ResumeCodeRun "CutArchetypes"



        '            End If

        '                Next BlockCount

        '                'Protect sheet and hide if previously hidden

        '                Call fncPrtSht(ActiveSheet.Name)
        '                ResumeCodeRun "CutArchetypes"

        '        If HiddenSheet = True Then

        '                    Sheets(SheetName).Visible = False

        '                End If

        '            Next SheetCount

        '            CatsCode = ""
        '            ColCount = 0

        '            Do While CatsCount < CatsUsed

        '                If WorksheetFunction.CountA(Range("Archetype1").Offset(ColCount, 0)) Then

        '                    CatsCount = CatsCount + 1
        '                    CatsCode = CatsCode & "i"
        '                Else

        '                    CatsCode = CatsCode & "o"

        '                End If

        '                ColCount = ColCount + 1
        '            Loop

        '            'Record code of categories used
        '            CheckRange = CatsCode

        '            EndCodeRun False

        'End Sub
        '        Sub ClearArchetypes(SheetName, CellAddress)

        '            ' Macro recorded 26/08/2005 by Alan Lewis

        '            InitiateCodeRun "ClearArchetypes"

        '    Dim ColsUsed As Integer

        '            Sheets("Repairs & Maint. Assumptions").Select
        '            ColsUsed = Range("StockCondCats").Columns.Count - 1

        '            On Error Resume Next

        '            Sheets(SheetName).Select
        '            ActiveSheet.Unprotect Password:=PW
        '    Range(Range(CellAddress).Offset(0, 1), Range(CellAddress).End(xlDown).Offset(0, ColsUsed)).Select
        '            Selection.Clear


        '        End Sub
        '        Sub Launch_AllColumns()

        '            'CodeSafe JW 23/4/22

        '            ' macro to call AllColumns procedure with a "Protect" argument
        '            ' for use by launch button on "Check Sheet"


        '            Call AllColumns("Protect")


        '        End Sub






    End Class

End Namespace
