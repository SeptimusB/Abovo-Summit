Imports System.IO
Imports System.Linq
Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.CustomGrid
Imports Abovo.ExportServices
Imports Abovo.FileManager
Imports Abovo.GeneralFunctions
Imports Abovo.LogDebugDev
Imports Abovo.RepositaryItems
Imports DevExpress.CodeParser
Imports DevExpress.Skins
Imports DevExpress.Skins.XtraForm
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils
Imports DevExpress.Utils.Drawing
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Extensions
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraLayout.Customization.Templates
Imports DevExpress.XtraPrinting
Imports DevExpress.XtraSpreadsheet
Imports DevExpress.XtraSpreadsheet.Model
Imports DevExpress.XtraSpreadsheet.PrintLayoutEngine
Imports DevExpress.XtraTreeList.Features.OfficeScrolling
Imports StressTest


Public Class FFRForm

    Inherits DevExpress.XtraEditors.XtraForm

    Private KeyDefRDS As RangeDataSource
    Private KeyValidationRDS As RangeDataSource
    Private RegSubsRDS As RangeDataSource
    Private UnRegSubsRDS As RangeDataSource
    Private KeyDefAfterTextRDS As RangeDataSource
    Private FFRInpuAdjStatements1 As RangeDataSource
    Private FFRInpuAdjStatements2 As RangeDataSource

    Private TabFFRKeyInitialise As Boolean = False
    Private TabFFRValInitialise As Boolean = False
    Private TabFFRFronSInitialise As Boolean = False
    Private TabFFRIASInitialise As Boolean = False
    Private TabFFRValidInitialise As Boolean = False

    Private TabFFRWorkInitialise As Boolean = False

    Private Formatter As New ObjectFormatter

    Private ModelID As Integer

    Private ExportPackageCount As Integer = 0
    Private ExportPackagesIndex As Integer = -1
    Private ExportPackages(-1) As GridExportPackage

    Private ExWorkbook As IWorkbook
    Private ScaleUnits As Integer
    Private ExportMode As String
    Private MyColourSwatch As Color

    Public Sub SetMode(ExMode As String)

        ExportMode = ExMode

    End Sub

    Public Sub Initialise()


    End Sub
    Public Sub New(SetModelID As Integer)
        SystemLog("CounterReset")
        MyColourSwatch = ExcelModels(SetModelID).ColourSwatch
        InitializeComponent()

        ModelID = SetModelID
        SystemLog("FFR1")
        Me.Width = Screen.PrimaryScreen.Bounds.Width * 0.8
        Me.Height = Screen.PrimaryScreen.Bounds.Height * 0.8

        AddHandler Me.WindowsUIButtonPanelSave.ButtonClick, AddressOf WindowsUIButtonPanelActions_ButtonClick
        Me.Text = "Complete the NROSH+ Financial Forecast Return for " + ExcelModels(SetModelID).WBStructure.CompanyName
        ScaleUnits = Me.Width * 0.007
        SystemLog("FFR2")
        Form_InitilisationProcess_SetDataSources()
        SystemLog("FFR6")
        BuildTab2()
        SystemLog("FFR7")
        SystemLog("Counter")
    End Sub
    Sub Form_InitilisationProcess_SetDataSources()

        SystemLog("FFR3")

        'Tab #1 - FFR Key Definitions

        Dim worksheet As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets("FFR Key Defn")

        Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range(ExcelModels(ModelID).WBStructure.FFRDefinition.FFRKeys)

        Dim RDSOptions As New RangeDataSourceOptions With {
            .UseFirstRowAsHeader = False,
            .PreserveFormulas = False,
            .SkipHiddenRows = True,
            .SkipHiddenColumns = True,
            .EditingOptions = DataSourceEditingOptions.AllowEdit
        }

        Dim ColList As New List(Of String) From {
            "Description",
            "Category",
            "SelectCode ",
            "Activity ",
            "Housing Basis",
            "Detail"
        }

        Dim ColMap As String = "TTTTTT"

        Dim ColNames As New SourceColumnDetector(ColList, ColMap)

        RDSOptions.DataSourceColumnTypeDetector = ColNames

        KeyDefRDS = range.GetDataSource(RDSOptions)

        ColList = New List(Of String) From {
            "Entry1",
            "Entry2",
            "Entry3",
            "Entry4",
            "Entry5",
            "Entry6",
            "Entry7",
            "Entry8",
            "Entry9",
            "Entry10"
        }

        ColMap = "TTTTTTTTTT"

        ColNames = New SourceColumnDetector(ColList, ColMap)

        RDSOptions.DataSourceColumnTypeDetector = ColNames

        range = worksheet.Range(ExcelModels(ModelID).WBStructure.FFRDefinition.FFRKeyAfterRange)

        KeyDefAfterTextRDS = range.GetDataSource(RDSOptions)

        '''''''''''''''''''''''''''''''''''''''''''''

        'Tab #2 - FFR  Validation

        worksheet = ExcelModels(ModelID).WB.Worksheets("Front Sheet")

        range = worksheet.Range(ExcelModels(ModelID).WBStructure.FFRDefinition.FFRRegSubsRange)
        ColList = New List(Of String) From {
            "Entry",
            "RP Code and name"
        }

        ColMap = "IT"

        ColNames = New SourceColumnDetector(ColList, ColMap)

        RDSOptions.DataSourceColumnTypeDetector = ColNames

        RegSubsRDS = range.GetDataSource(RDSOptions)

        SystemLog("FFR4")

        range = worksheet.Range(ExcelModels(ModelID).WBStructure.FFRDefinition.FFRUnRegSubsRange)

        ColList = New List(Of String) From {
            "Entry",
            "Name of unregistered entity or joint venture"
        }

        ColMap = "IT"

        ColNames = New SourceColumnDetector(ColList, ColMap)

        RDSOptions.DataSourceColumnTypeDetector = ColNames

        UnRegSubsRDS = range.GetDataSource(RDSOptions)
        SystemLog("FFR5")




    End Sub
    Sub Form_InitilisationProcess_BuildTabs()

    End Sub

    Sub BuildTab1()

        'Header HML
        Dim worksheet As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets("FFR Key Defn")

        Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range(ExcelModels(ModelID).WBStructure.FFRDefinition.FFRKeyPrecursorRange)


        Dim RangeList As New List(Of DevExpress.Spreadsheet.CellRange)

        RangeList.Add(range)

        WebBrowserKeyDefHeader.DocumentText = RenderRangeCells(RangeList)

        'GC FFR Key Definitions


        Me.GridControlKeyDef.DataSource = KeyDefRDS

        Me.GridControlKeyDef.MainView.PopulateColumns()
        Dim GV As GridView = CType(Me.GridControlKeyDef.MainView, GridView)

        GV.Columns(0).OptionsColumn.ReadOnly = True
        GV.Columns(1).OptionsColumn.ReadOnly = True
        GV.Columns(3).OptionsColumn.ReadOnly = True
        GV.Columns(4).OptionsColumn.ReadOnly = True
        GV.Columns(5).OptionsColumn.ReadOnly = True
        GV.BestFitColumns()

        Formatter.FormatGridView(GV, GridControlKeyDef)





        GridControAfterText.DataSource = KeyDefAfterTextRDS

        Me.GridControAfterText.MainView.PopulateColumns()

        GV = CType(Me.GridControAfterText.MainView, GridView)
        Formatter.FormatGridView(GV, GridControAfterText)

    End Sub
    Sub BuildTab2()

        Dim worksheet As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets("FFR Validation Summary")

        Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range(ExcelModels(ModelID).WBStructure.FFRDefinition.FFRKeyValidationRange)
        Dim RangeList As New List(Of DevExpress.Spreadsheet.CellRange)
        RangeList = New List(Of DevExpress.Spreadsheet.CellRange)

        RangeList.Add(range)
        WebBrowserFFRValidation.DocumentText = RenderRangeCells(RangeList)

    End Sub
    Sub BuildTab3()

        ModelPostingTextBox1.Initialise(ModelID, "Front Sheet", "B5")
        ModelPostingDateBox1.Initialise(ModelID, "Front Sheet", "B6")

        ModelPostingComboBoxAgreeInclusion.SetModelID = ModelID
        ModelPostingComboBoxAgreeInclusion.SetTargetWorksheet = "Front Sheet"
        ModelPostingComboBoxAgreeInclusion.SetTargetCell = "B7"
        ModelPostingComboBoxAgreeInclusion.InitialiseStandard("Rep_YesNo")
        ModelPostingComboBoxAgreeInclusion.SetLimitToList = True
        ModelPostingComboBoxAgreeInclusion.ProcesDefValue()

        ModelPostingComboBoxConfirmListed.SetModelID = ModelID
        ModelPostingComboBoxConfirmListed.SetTargetWorksheet = "Front Sheet"
        ModelPostingComboBoxConfirmListed.SetTargetCell = "B36"
        ModelPostingComboBoxConfirmListed.InitialiseStandard("Rep_YesNo")
        ModelPostingComboBoxConfirmListed.SetLimitToList = True
        ModelPostingComboBoxConfirmListed.ProcesDefValue()


        Dim EditControl As AbovoRespositaryItem

        EditControl = RepositaryItems.GetNumericDropEditor(1, 25)






        GridControlRegSubs.DataSource = RegSubsRDS

        Dim GV As GridView = CType(Me.GridControlRegSubs.MainView, GridView)

        GridControlRegSubs.RepositoryItems.Add(EditControl.RetCombo)
        GV.Columns(0).ColumnEdit = EditControl.RetCombo

        GV.BestFitColumns()

        Formatter.FormatGridView(GV, GridControlRegSubs)

        GridControlNonRegEnts.DataSource = RegSubsRDS

        GV = CType(Me.GridControlNonRegEnts.MainView, GridView)
        GV.Columns(0).ColumnEdit = EditControl.RetCombo
        GV.BestFitColumns()

        Formatter.FormatGridView(GV, GridControlNonRegEnts)

    End Sub
    Sub BuildTab4()

        Dim FFRAdjInputs As New DataInterfaceTemplate(ModelID, 3, 0)

        FFRAdjInputs.Dock = DockStyle.Fill

        Me.TablePanelFFRInputs.Controls.Add(FFRAdjInputs)

    End Sub
    Sub BuildTab5()

        Dim FFRWorkings As New DataInterfaceTemplate(ModelID, 3, 1)

        FFRWorkings.Dock = DockStyle.Fill

        Me.XtraTabPageFFRWorkings.Controls.Add(FFRWorkings)

    End Sub
    Public Property FormBorderColor() As Color
        Get
            Return MyColourSwatch
        End Get
        Set(ByVal value As Color)
            MyColourSwatch = value
        End Set
    End Property
    Protected Overrides Function CreateFormBorderPainter() As DevExpress.Skins.XtraForm.FormPainter
        Return New CustomFormPainterFFR(Me, LookAndFeel)
    End Function


    Sub FFR_New_Extraction()


        Dim BusPlanFile As IWorkbook = FileManager.GetWorkBook(ModelID)
        Dim NumRanges As Integer



        NumRanges = BusPlanFile.Range("FFRRangeNames").RowCount


        Dim openFileDialog As New OpenFileDialog() With {
                                                        .Filter = "Excel Files|*.xlsm;*.xlsb",
                                                        .Title = "Select FFR Template"
                                                        }

        If Not openFileDialog.ShowDialog() = DialogResult.OK Then Exit Sub

        Dim FFRFile As New Workbook
        FFRFile.Options.CalculationMode = WorkbookCalculationMode.Manual
        FFRFile.Options.CalculationEngineType = CalculationEngineType.ChainBased

        FFRFile.DocumentSettings.Calculation.EnableMultiThreading = False
        FFRFile.LoadDocument(openFileDialog.FileName, DocumentFormat.Xlsm)

        If FFRFile.Worksheets("Cover Sheet").Range("$B$4").Value.TextValue <> "Spreadsheet Import Template - Financial Forecast Return (FFR)" Then

            MsgBox("Sorry, the file selected does not appear to be a correct FFR template which must be downloaded as a provider-specific template from NROSH+." + vbCr + "See https://nroshplus.regulatorofsocialhousing.org.uk/")

            FFRFile.Dispose()
            FFRFile = Nothing

            Return

        End If

        FFRFile.BeginUpdate()

        Dim AddressCell As DevExpress.Spreadsheet.Cell
        Dim SourceRange As DevExpress.Spreadsheet.CellRange
        Dim DestRange As DevExpress.Spreadsheet.CellRange

        For i = 1 To NumRanges - 1

            AddressCell = BusPlanFile.Range("FFRListHeading")(i, 8)
            SourceRange = BusPlanFile.Range(AddressCell.Value.TextValue)
            AddressCell = BusPlanFile.Range("FFRListHeading")(i, 3)
            DestRange = FFRFile.Range(AddressCell.Value.TextValue)
            DestRange.CopyFrom(SourceRange, PasteSpecial.Values)

        Next i

        FFRFile.EndUpdate()

        'FFRFile.CalculateFull()


        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Excel Files|*.xlsm|All Files|*.*"
        saveFileDialog.Title = "Save FFR Document"
        Dim FilePath As String

        If saveFileDialog.ShowDialog() = DialogResult.OK Then FFRFile.SaveDocument(saveFileDialog.FileName, DocumentFormat.Xlsm)
        FilePath = saveFileDialog.FileName
        'Dim stream As New MemoryStream
        'FFRFile.SaveDocument(stream, DevExpress.Spreadsheet.DocumentFormat.Xlsm)
        'stream.Position = 0
        'SpreadsheetControlExport.LoadDocument(stream, DevExpress.Spreadsheet.DocumentFormat.Xlsm)
        'SpreadsheetControlExport.Enabled = True
        'Me.XtraTabControlExport.SelectedTabPage = Me.XtraTabPageExportXLS

        FFRFile.Dispose()
        FFRFile = Nothing

        If MsgBox("FFR File saved as " & FilePath & ". Do you want to open it in Excel?", MsgBoxStyle.YesNo, "Open FFR File") = MsgBoxResult.Yes Then
            OpenFileInExcel(FilePath)
        End If



    End Sub


    Public Sub SaveNewWorkbook()


        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Excel Files|*.xlsm|All Files|*.*"
        saveFileDialog.Title = "Save Spreadsheet Document"

        ' Show the SaveFileDialog and check if the user clicked the Save button
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            ' Check if the file already exists
            If File.Exists(saveFileDialog.FileName) Then
                ' Optionally, prompt the user to confirm overwriting the file
                Dim result As DialogResult = MessageBox.Show("The file already exists. Do you want to overwrite it?", "Confirm Overwrite", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                If result = DialogResult.No Then
                    ' If the user chooses not to overwrite, exit the method
                    Return
                End If
            End If

            ' Save the document to the specified file path
            SpreadsheetControlExport.SaveDocument(saveFileDialog.FileName, DocumentFormat.Xlsm)
            OpenFileInExcel(saveFileDialog.FileName)

        End If


    End Sub


    Sub ClearWorkbook()

        ExWorkbook = SpreadsheetControlExport.Document

        ExWorkbook.BeginUpdate()

        Dim worksheetsToRemove As New List(Of DevExpress.Spreadsheet.Worksheet)
        For Each ws As DevExpress.Spreadsheet.Worksheet In ExWorkbook.Worksheets
            worksheetsToRemove.Add(ws)
        Next

        For Each ws As DevExpress.Spreadsheet.Worksheet In worksheetsToRemove
            ExWorkbook.Worksheets.Remove(ws)
        Next

        ExWorkbook.EndUpdate()

    End Sub




    Private Sub WindowsUIButtonPanelActions_ButtonClick(sender As Object, e As DevExpress.XtraBars.Docking2010.ButtonEventArgs)

        Dim tag As String = CType(e.Button, WindowsUIButton).Tag.ToString()
        Select Case tag

            Case "Clear"

                ClearWorkbook()


            Case "SaveNew"

                SaveNewWorkbook()

            Case "Preview"

                If Me.XtraTabControlFFR.SelectedTabPage.Name = "XtraTabPageExportXLS" Then



                Else

                    'ProcessPDFExports()

                End If

            Case "Close"

                Me.Hide()

            Case "ExpandAll"

                'GridView_Process_ExpandAll(ActiveGridView)

            Case "CollapseAll"

                ' ActiveGridView.CollapseAllGroups()
                'ActiveGridView.ExpandGroupLevel(0)
                ' Gr 'idView_Process_SetExpandedLevels(ActiveGridView)

        End Select

    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        FFR_New_Extraction()
    End Sub

    Friend Class SourceColumnDetector

        Implements IDataSourceColumnTypeDetector

        Private ColNames As List(Of String)
        Private ColMap As String

        Public Sub New(SentColNames As List(Of String), SentColMap As String)

            ColNames = SentColNames
            ColMap = SentColMap

        End Sub
        Public Function GetColumnName(ByVal index As Integer, ByVal offset As Integer, ByVal range As DevExpress.Spreadsheet.CellRange) As String Implements IDataSourceColumnTypeDetector.GetColumnName
            If index < 0 OrElse index >= ColNames.Count Then
                Throw New ArgumentOutOfRangeException("index", "Index is out of range.")
            End If
            If ColNames(index) = "Blank" Then
                Return String.Empty
            Else
                Return ColNames(index)
            End If

        End Function

        Public Function GetColumnType(ByVal index As Integer, ByVal offset As Integer, ByVal range As DevExpress.Spreadsheet.CellRange) As Type Implements IDataSourceColumnTypeDetector.GetColumnType

            Dim defaultType As Type = GetType(String)
            Select Case Mid(ColMap, index + 1, 1)
                Case "T"
                    Return GetType(String)
                Case "B"
                    Return GetType(Boolean)
                Case "D"
                    Return GetType(Date)
                Case "N"
                    Return GetType(Double)
                Case "I"
                    Return GetType(Integer)
                Case Else
                    Return defaultType
            End Select

        End Function
    End Class
    Public Sub ManualDispose()
        If Not IsNothing(KeyDefRDS) Then
            KeyDefRDS.Dispose()
            KeyDefRDS = Nothing
        End If
        If Not IsNothing(KeyValidationRDS) Then
            KeyValidationRDS.Dispose()
            KeyValidationRDS = Nothing
        End If
        If Not IsNothing(ExWorkbook) Then
            ExWorkbook.Dispose()
            ExWorkbook = Nothing
        End If
        Me.Dispose()
    End Sub
    Private Sub FFRForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        e.Cancel = True
        Me.Hide()

    End Sub


    Private Sub ModelPostingDateBox1_EditValueChanged(sender As Object, e As EventArgs) Handles ModelPostingDateBox1.EditValueChanged

    End Sub
    Private Sub ProcessTabChange(sender As Object, e As DevExpress.XtraTab.TabPageChangedEventArgs) Handles XtraTabControlFFR.SelectedPageChanged

        If e.Page.Name = "XtraTabPageFrontSheet" Then
            If Not TabFFRFronSInitialise Then

                SystemLog("Form_InitilisationProcess_BuildTabs3")
                BuildTab3()
                TabFFRFronSInitialise = True

            End If
        ElseIf e.Page.Name = "XtraTabPageInputAdjustments" Then
            If Not TabFFRIASInitialise Then
                SystemLog("Form_InitilisationProcess_BuildTabs4")
                BuildTab4()
                TabFFRIASInitialise = True
            End If
        ElseIf e.Page.Name = "XtraTabPageFFRWorkings" Then
            If Not TabFFRWorkInitialise Then
                SystemLog("Form_InitilisationProcess_BuildTabs5")
                BuildTab5()
                TabFFRWorkInitialise = True
            End If
        ElseIf e.Page.Name = "XtraTabPageFFRKey" Then
            If Not TabFFRKeyInitialise Then
                SystemLog("Form_InitilisationProcess_BuildTabs1")
                BuildTab1()
                TabFFRKeyInitialise = True
            End If
        End If

    End Sub

End Class

Public Class CustomFormPainterFFR
    Inherits FormPainter
    Public Sub New(ByVal owner As System.Windows.Forms.Control, ByVal provider As DevExpress.Skins.ISkinProvider)
        MyBase.New(owner, provider)
    End Sub
    Private Function GetFormBorderColor() As Color
        Dim formBorderColor = (TryCast(Owner, FFRForm)).FormBorderColor
        Return formBorderColor
    End Function
    Protected Overrides Sub DrawBackground(ByVal cache As GraphicsCache)
        Dim info = GetCaptionInfo()
        Dim ee = TryCast(info, ObjectInfoArgs)
        Dim formBorderColor = GetFormBorderColor()
        cache.FillRectangle(New SolidBrush(formBorderColor), ee.Bounds)
    End Sub
    Protected Overrides Sub DrawFrameCore(ByVal cache As GraphicsCache, ByVal info As SkinElementInfo, ByVal kind As FrameKind)
        Dim formBorderColor = GetFormBorderColor()
        cache.FillRectangle(formBorderColor, info.Bounds)
    End Sub
End Class