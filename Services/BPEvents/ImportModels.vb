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

        Public Shared Sub ImportStockRentModel(ModelID As Integer)

            Dim L As Integer, i As Integer, O, NewThings As Integer
            Dim AllYears As Integer = 30
            Dim iLength As Long, iPosition As Long
            Dim TargRng As String

            'On Error GoTo Err_Handler

            Dim BusPlanFile As IWorkbook = FileManager.GetWorkBook(ModelID)
            Dim ActiveRentModel As DevExpress.Spreadsheet.Workbook = Nothing

            Dim FileToOpen As String

            SystemLog("Starting perf import")
            Dim XtraOpenFileDialogMainScreen As New DevExpress.XtraEditors.XtraOpenFileDialog()
            XtraOpenFileDialogMainScreen.Filter = "Abovo Rent Models |*.xlsm;*.xlsb;*.arm"
            XtraOpenFileDialogMainScreen.Title = "Select Rent Model"

            If XtraOpenFileDialogMainScreen.ShowDialog = Windows.Forms.DialogResult.Cancel Then

                BusPlanFile = Nothing
                Exit Sub

            End If

            FileToOpen = XtraOpenFileDialogMainScreen.FileName

            Dim FileOpenResult As AbovoTransaction
            Dim MyFileInfos As IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FileToOpen)
            Dim ActiveRentModelID As Integer

            FileOpenResult = FileManager.OpenModel(FileToOpen, MyFileInfos)

            If FileOpenResult.BError = False Then

                ActiveRentModelID = FileOpenResult.IntegerReturn

                If Not FileOpenResult.StringReturn = "AbovoRM" Then

                    ActiveRentModel = FileManager.GetWorkBook(ActiveRentModelID)

                Else

                    MsgBox("The selected file is not a valid Rent Restructuring model. Please select another file.")
                    GoTo Exiter

                End If

            End If

            Dim SourceRange As CellRange

            SourceRange = BusPlanFile.Range("TransferDate")

            ActiveRentModel.Range("BPStartDate").CopyFrom(SourceRange, PasteSpecial.Values)

            ActiveRentModel.Calculate()




            If ActiveRentModel.Range("TransRealIncr").Value = CellValueType.Error Then

                MsgBox("Please choose more recent" & Chr(13) & "version of the rent model.")
                CloseModel(ActiveRentModelID)

                GoTo Exiter

            End If


            SourceRange = ActiveRentModel.Range("BPCats")

            BusPlanFile.Range("BPCats").CopyFrom(SourceRange, PasteSpecial.Values)

            SourceRange = ActiveRentModel.Range("StTrans")

            BusPlanFile.Range("StTrans").CopyFrom(SourceRange, PasteSpecial.Values)

            ActiveRentModel.Worksheets("Hidden - Business Plan Workings").Visible = True
            SourceRange = ActiveRentModel.Range("BPRents")
            BusPlanFile.Range("BPRents").CopyFrom(SourceRange, PasteSpecial.Values)

            SourceRange = ActiveRentModel.Range("BPTarRents")
            BusPlanFile.Range("BPTarRents").CopyFrom(SourceRange, PasteSpecial.Values)

            BusPlanFile.Range("BPTarRentsFV5").CopyFrom(SourceRange, PasteSpecial.Values)

            SourceRange = ActiveRentModel.Range("RentWks")
            BusPlanFile.Range("RentWks").CopyFrom(SourceRange, PasteSpecial.Values)




            '******************************************
            ' Copy and paste real rent increases,
            ' extending destination table if necessary
            ' and numbering years

            ' activate Economic Assumptions sheet in BP

            SourceRange = BusPlanFile.Range("IR_ExistRentReal")

            Dim TargetSize As Integer = SourceRange.RowCount

            If TargetSize < AllYears Then WorkbookManager.SetRangeRowsToSize(ModelID, "IR_ExistRentReal", AllYears)

            SourceRange = ActiveRentModel.Range("IR_ExistRentReal")

            Dim YearCell As Cell

            For i = 3 To AllYears

                YearCell = SourceRange(i - 1, 0)
                YearCell.Value = i

            Next i

            SourceRange = ActiveRentModel.Range("TransRealIncr")
            BusPlanFile.Range("TransRealIncr").CopyFrom(SourceRange, PasteSpecial.Values)

            If ActiveRentModel.Worksheets("Contents").Cells("B5").DisplayText = "" Then

                SourceRange = ActiveRentModel.Worksheets("Contents").Range("B5")

            Else

                SourceRange = ActiveRentModel.Worksheets("Contents").Range("B5")

            End If

            BusPlanFile.Range("Rentsfile").CopyFrom(SourceRange, PasteSpecial.Values)

            BusPlanFile.Range("Datestamprent").Value = Now()

Exiter:

            On Error Resume Next

            CloseModel(ActiveRentModelID)

            Exit Sub

Err_Handler:

            MsgBox("Sorry, an error occured during procedure. Logging the error failed too.  The error is " & Err.Description)

            Err.Clear()

            Resume Exiter

        End Sub
        Public Shared Sub ImportManagementServiceCosts(ModelID As Integer)

            '**************************************************************************
            '   Opens Management & Service Costs Model, assigns module level filenames
            '   KC 28/11/03
            '
            '   Calls: Chk_Co_Name_Exists
            '**************************************************************************

            'Declare variables

            Dim FileToOpen As String
            Dim BusPlanFile As IWorkbook = FileManager.GetWorkBook(ModelID)
            Dim ActiveMCModel As IWorkbook = Nothing

            On Error GoTo Err_Handler

            Dim XtraOpenFileDialogMainScreen As New DevExpress.XtraEditors.XtraOpenFileDialog()
            XtraOpenFileDialogMainScreen.Title = "Select file which contains Management & Service Costs"
            XtraOpenFileDialogMainScreen.Filter = "Abovo Management Cost Models |*.xlsm;*.xlsb;*.amc"

            If XtraOpenFileDialogMainScreen.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

            FileToOpen = XtraOpenFileDialogMainScreen.FileName

            Dim FileOpenResult As AbovoTransaction
            Dim MyFileInfos As IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FileToOpen)
            Dim ActiveMCModelID As Integer

            FileOpenResult = FileManager.OpenModel(FileToOpen, MyFileInfos)

            If FileOpenResult.BError = False Then

                ActiveMCModelID = FileOpenResult.IntegerReturn

                If Not FileOpenResult.StringReturn = "AbovoMCM" Then

                    ActiveMCModel = FileManager.GetWorkBook(ActiveMCModelID)

                Else

                    MsgBox("The selected file is not a valid Management & Service Costs model. Please select another file.")
                    GoTo Exiter

                End If

            End If



            If Chk_Co_Name_Exists(BusPlanFile, ActiveMCModel, ModelID) Then

                If Copy_Man_Ser_Costs(BusPlanFile, ActiveMCModel, ModelID) Then

                    MsgBox("Import Complete")

                End If
            Else

                MsgBox("Import cannot continue. " & Chr(10) & "Please ensure the Company name is consistent in both files")

                'Close M&S File




            End If




Exiter:


            On Error Resume Next
            FileManager.CloseModel(ActiveMCModelID)

            Exit Sub

Err_Handler:



            MsgBox("Sorry, an error occured: " & Err.Description)


            Err.Clear()

            Resume Exiter

        End Sub
        Shared Function Chk_Co_Name_Exists(BPFile As DevExpress.Spreadsheet.Workbook, MSFile As DevExpress.Spreadsheet.Workbook, BPModelID As Integer) As Boolean

            Return True

            '************************************************************************
            ' KC 15/12/03
            ' This Sub checks whether the Company name selected in the BP File in the
            ' SelectTrust dropdown exists in the Management Costs File. If it does it
            ' imports all the relevant data, if it doesn't it informs the user and ends
            '
            '  Called by :Get_Man_Ser_Costs
            '  Calls:     Copy_Man_Ser_Costs
            '*************************************************************************

            'declare variables
            Dim strCompanyName As String = ""
            Dim strCompanyList As String = ""
            Dim TargWS As Worksheet

            Dim c As CellRange
            Dim ValueExists As Boolean

            On Error GoTo Err_Handler


            'initialise variables

            ValueExists = False



            TargWS = BPFile.Worksheets("Global Assumptions")

            '    .Select
            '    .Range("SelectTrust").Select

            'End With

            '   Assign chosen company name in BP to var
            strCompanyName = BPFile.Range("SelectTrust").Value.TextValue

            Dim ValidaionColl As DevExpress.Spreadsheet.DataValidation



            For Each ValidaionColl In TargWS.DataValidations



            Next


            'Windows(MSFile).Activate
            'Sheets("Summary Costs").Select
            'Range("SelCompany").Select

            '   Extracts name of list populating dropdown SelCompany,
            '   without = sign and puts into var

            'strCompanyList = Mid(Range("SelCompany").Validation.Formula1, 2, Len(Range("SelCompany").Validation.Formula1) - 1)

            '   looks through the list to see if company name exists in msfile list
            '   if found boolean set to true


            'For Each c In Sheets("Global Assumptions").Range(strCompanyList)

            '    If c.Value = strCompanyName Then

            '        ValueExists = True
            '        Exit For

            '    End If

            'Next c


            'If ValueExists = True Then

            '    'put value in BP SelectTrust range into msfile SelCompany range
            '    ' and carries on to copy routine
            '    Range("SelCompany").Value = strCompanyName

            '    Return True

            'Else

            '    Return False

            'End If

Exiter:

            On Error Resume Next



            Exit Function

Err_Handler:



            MsgBox("Sorry, an error occured during procedure.  The error is " & Err.Description)



            Err.Clear()

            Resume Exiter

        End Function


        Public Shared Function Copy_Man_Ser_Costs(BPFile As DevExpress.Spreadsheet.Workbook, MSFile As DevExpress.Spreadsheet.Workbook, BPModelID As Integer) As Boolean

            '**************************************************************************
            '   Imports Management & Service Costs,
            '   TL
            '
            '   Called by : Chk_Co_Name_Exists
            '   Calls: Home
            '**************************************************************************


            Dim MgtCostCtr As Integer

            On Error GoTo Err_Handler

            Dim SourceRange As CellRange
            Dim DestRange As CellRange

            If bIsDevelopment Then On Error GoTo 0




            '   Unprotects key business plan worksheets
            UNProtectWS(BPModelID, "Management Costs Assumptions")


            SourceRange = MSFile.Worksheets("Global Assumptions").Range("Services")


            BPFile.Worksheets("Management Costs Assumptions").Range("Services").CopyFrom(SourceRange, PasteSpecial.Values)

            Dim CurrCell As DevExpress.Spreadsheet.Cell



            '   Set Inflation Indicator to N so that all cashflows brought through are real
            MSFile.Worksheets("Inflation & VAT Assumptions").Range("InflatIndic")(0, 0).SetValueFromText("N")
            MSFile.Calculate()

            For MgtCostCtr = 1 To 5

                CurrCell = MSFile.Range("YearNo")(0, 0)

                CurrCell.SetValue(MgtCostCtr)

                MSFile.Calculate()

                SourceRange = MSFile.Range("StaffCosts")

                DestRange = BPFile.Range("StaffCosts" & MgtCostCtr.ToString)

                DestRange.CopyFrom(SourceRange, PasteSpecial.Values)

                SourceRange = MSFile.Range("OtherCosts")

                DestRange = BPFile.Range("OtherCosts" & MgtCostCtr.ToString)

                DestRange.CopyFrom(SourceRange, PasteSpecial.Values)

            Next MgtCostCtr




            CurrCell = MSFile.Range("CompanyName")(0, 0)
            CurrCell.Formula = "=CELL(""FILENAME"")"
            MSFile.Calculate()

            Dim tempString As String
            tempString = CurrCell.DisplayText

            CurrCell = BPFile.Range("MCfile")(0, 0)
            CurrCell.SetValueFromText(tempString)

            CurrCell = BPFile.Range("Datestampmc")(0, 0)
            CurrCell.SetValueFromText(Now().ToString)

            MSFile = Nothing
            'JW Note 84233



Exiter:

            On Error Resume Next
            Exit Function


Err_Handler:


            MsgBox("Sorry, an error occured. The error has been logged.  The error is " & Err.Description)


            Err.Clear()

            Resume Exiter

        End Function
        Public Shared Sub ImportStockConditionSurvey(SetModelID As Integer)


            'CodeSafe JW 25/4/22

            '
            '   Imports Costs from SCS spreadsheet - needs to be set up for each organisation

            Dim FileToOpen
            Dim rmFile As IWorkbook
            Dim BusPlanFile As IWorkbook = FileManager.GetWorkBook(SetModelID)
            Dim ActiveRMModel As IWorkbook = Nothing

            On Error GoTo Err_Handler

            '   assign BP file

            Dim XtraOpenFileDialogMainScreen As New DevExpress.XtraEditors.XtraOpenFileDialog()
            XtraOpenFileDialogMainScreen.Title = "Select file which contains Stock Condition Survey data"
            XtraOpenFileDialogMainScreen.Filter = "Abovo Repairs and Maint Models |*.xlsm;*.xlsb;*.armc"

            If XtraOpenFileDialogMainScreen.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

            FileToOpen = XtraOpenFileDialogMainScreen.FileName

            Dim FileOpenResult As AbovoTransaction
            Dim MyFileInfos As IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FileToOpen)
            Dim ActiveRMModelID As Integer

            FileOpenResult = FileManager.OpenModel(FileToOpen, MyFileInfos)

            If FileOpenResult.BError = False Then

                ActiveRMModelID = FileOpenResult.IntegerReturn

                If Not FileOpenResult.StringReturn = "AbovoMCM" Then

                    ActiveRMModel = FileManager.GetWorkBook(ActiveRMModelID)

                Else

                    MsgBox("The selected file is not a valid Repairs and Maint Costs model. Please select another file.")
                    GoTo Exiter

                End If

            End If

            Dim SourceRange As CellRange
            Dim DestRange As CellRange
            Dim CurrCell As DevExpress.Spreadsheet.Cell
            Dim TempString As String

            '   Copy update file name and date stamp time of transfer
            CurrCell = ActiveRMModel.Worksheets("Totals").Range("z1")(0, 0)

            CurrCell.Formula = "=CELL(""FILENAME"")"
            ActiveRMModel.Calculate()
            TempString = CurrCell.DisplayText

            CurrCell = BusPlanFile.Range("SCSfile")(0, 0)

            CurrCell.SetValueFromText(TempString)

            CurrCell = BusPlanFile.Range("Datestampscs")(0, 0)

            CurrCell.SetValue(Now())



Exiter:

            On Error Resume Next



            Exit Sub

Err_Handler:

            MsgBox("Sorry, an error occured. The error has been logged.  The error is " & Err.Description)

            Err.Clear()

            Resume Exiter

        End Sub






    End Class
End Namespace