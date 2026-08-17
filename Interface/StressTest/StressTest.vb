Imports System.ComponentModel
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.Remoting.Contexts
Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.CustomGrid
Imports Abovo.DataObject
Imports Abovo.FileManager
Imports Abovo.GeneralFunctions
Imports Abovo.LogDebugDev
Imports Abovo.ObjectMiddler
Imports Abovo.PresentationManager
Imports Abovo.WorkbookManager
Imports Abovo.WSSecurity
Imports DevExpress.CodeParser
Imports DevExpress.Drawing
Imports DevExpress.Pdf.Native.BouncyCastle.Asn1.X509
Imports DevExpress.Skins
Imports DevExpress.Skins.XtraForm
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Functions
Imports DevExpress.Utils
Imports DevExpress.Utils.Drawing
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.XtraCharts
Imports DevExpress.XtraEditors.Mask
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Scrolling
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports DevExpress.XtraLayout.Customization.Templates
Imports DevExpress.XtraRichEdit.Layout
Imports DevExpress.XtraRichEdit.Model
Imports DevExpress.XtraSpreadsheet
Imports DevExpress.XtraSpreadsheet.Model
Imports DevExpress.XtraSpreadsheet.PrintLayoutEngine
Imports Microsoft.Office.Interop.Excel


Public Class StressTest

    Inherits DevExpress.XtraEditors.XtraForm

    Private STMode As String
    Private StandardPercentSpinEdit As RepositoryItemSpinEdit
    Private StandardYesEmptyEdit As RepositoryItemComboBox
    Private StandardIntegerTextBoxEdit As RepositoryItemTextEdit
    Private StandardPercentageTextBoxEdit As RepositoryItemTextEdit
    Private Standard2digitnumberTextBoxEdit As RepositoryItemTextEdit
    Private StandardStringTextBoxEdit As RepositoryItemTextEdit
    Private DSMitDataRange As RangeDataSource
    Private DSMitDevDataRange As RangeDataSource
    Private DSMitMoneyDataRange As RangeDataSource
    Private CovGraphsDataRange As RangeDataSource
    'Public DSBreachInitialised As Boolean
    Private DSStressesDataRange As RangeDataSource
    Private DSOutputsDataRange As RangeDataSource
    Private DSTextOutputsDataRange As RangeDataSource
    Private DSBreachOutputsDataRange As RangeDataSource
    Private ActiveWorkbook As IWorkbook


    'Tab #2
    Private DSStressSenitivityLiveDataRange As RangeDataSource
    Private DSStressSenitivityCaputresDataRange As RangeDataSource

    Private ModelID As Integer
    Private ExportPackageCount As Integer = 0
    Private ExportPackagesIndex As Integer = -1
    Private ExportPackages(-1) As GridExportPackage

    Private PresentedDS As Abovo.DataObject.DataCellRange
    Private ScaleUnits As Integer
    Private ExportMode As String
    Private MyColourSwatch As Color
    Private Formatter As ObjectFormatter
    Private WrapCG_Mits As CustomGridWrapper
    Private View_WrapCG_Mits As CustomGridView
    Private View_WrapTextGrid As GridView
    Private WrapCG_MitsDev As CustomGridWrapper
    Private WrapCG_MitsMoney As CustomGridWrapper
    Private View_WrapCG_Stresses As CustomGridView


    Public Sub SetMode(ExMode As String)

        ExportMode = ExMode

    End Sub

    Public Sub Initialise()

    End Sub
    Public Sub New(SetModelID As Integer)


        Me.BackColor = System.Drawing.Color.White
        MyColourSwatch = ExcelModels(SetModelID).ColourSwatch

        InitializeComponent()

        XtraTabControlStressTest.SelectedTabPage = XtraTabPageLMVP

        Formatter = New ObjectFormatter
        ModelID = SetModelID

        Me.WindowState = FormWindowState.Maximized

        ActiveWorkbook = ExcelModels(ModelID).WB

        ScaleUnits = Me.Width * 0.007

        CreateCustomEditors()

        If ActiveWorkbook.DefinedNames.GetDefinedName("StressTestMode").Range(0, 0).Value.TextValue = "Y" Then

            STMode = "Y"
            ToggleModeSwitch.IsOn = True
            TextEditMultivariableName.EditValue = ActiveWorkbook.DefinedNames.GetDefinedName("NewStressName").Range(0, 0).Value.TextValue

        Else

            STMode = "N"
            ToggleModeSwitch.IsOn = False
            TextEditMultivariableName.EditValue = "Base"

        End If

        Form_InitilisationProcess_SetDataSources()

        BuildCovCharts()

        BuildHTMLRenders()

        XtraTabControlStressTest.LookAndFeel.UseDefaultLookAndFeel = False
        XtraTabControlStressTest.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        XtraTabControlStressTest.Appearance.BackColor = Color.White


        PanelControl1.Dock = DockStyle.Fill
        PanelControl1.LookAndFeel.UseDefaultLookAndFeel = False
        PanelControl1.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        PanelControl1.Appearance.BackColor = Color.White

        PanelControl2.Dock = DockStyle.Fill
        PanelControl2.LookAndFeel.UseDefaultLookAndFeel = False
        PanelControl2.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        PanelControl2.Appearance.BackColor = Color.White

        PanelControl3.Dock = DockStyle.Fill
        PanelControl3.LookAndFeel.UseDefaultLookAndFeel = False
        PanelControl3.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        PanelControl3.Appearance.BackColor = Color.White

        PanelControl4.Dock = DockStyle.Fill
        PanelControl4.LookAndFeel.UseDefaultLookAndFeel = False
        PanelControl4.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        PanelControl4.Appearance.BackColor = Color.White

        PanelControlCovSel.Dock = DockStyle.Fill
        PanelControlCovSel.LookAndFeel.UseDefaultLookAndFeel = False
        PanelControlCovSel.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        PanelControlCovSel.Appearance.BackColor = Color.White

        WindowsUIButtonPanelStressNavigator.ForeColor = AbovoBlue



        AddHandlers()

    End Sub

    Sub CreateCustomEditors()

        StandardPercentSpinEdit = New RepositoryItemSpinEdit With {
                                    .MinValue = -200,
                                    .Increment = CDec(0.0025),
                                    .MaxValue = 200,
                                    .EditMask = "p2",
                                    .UseMaskAsDisplayFormat = True
                                    }

        StandardYesEmptyEdit = New RepositoryItemComboBox
        StandardYesEmptyEdit.Appearance.ForeColor = Color.White
        StandardYesEmptyEdit.Appearance.Options.UseForeColor = True
        StandardYesEmptyEdit.Appearance.BackColor = AbovoBlue
        StandardYesEmptyEdit.Appearance.Options.UseBackColor = True
        StandardYesEmptyEdit.Items.Add("Yes")
        StandardYesEmptyEdit.Items.Add("")

        StandardIntegerTextBoxEdit = New RepositoryItemTextEdit
        StandardIntegerTextBoxEdit.MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
        StandardIntegerTextBoxEdit.MaskSettings.Set("mask", "n0")
        StandardIntegerTextBoxEdit.MaskSettings.Set("culture", "en-GB")
        StandardIntegerTextBoxEdit.UseMaskAsDisplayFormat = True

        Standard2digitnumberTextBoxEdit = New RepositoryItemTextEdit
        Standard2digitnumberTextBoxEdit.MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
        Standard2digitnumberTextBoxEdit.MaskSettings.Set("mask", "n2")
        Standard2digitnumberTextBoxEdit.MaskSettings.Set("culture", "en-GB")
        Standard2digitnumberTextBoxEdit.UseMaskAsDisplayFormat = True

        StandardPercentageTextBoxEdit = New RepositoryItemTextEdit
        StandardPercentageTextBoxEdit.MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
        StandardPercentageTextBoxEdit.MaskSettings.Set("mask", "p2")
        StandardPercentageTextBoxEdit.MaskSettings.Set("culture", "en-GB")
        StandardPercentageTextBoxEdit.UseMaskAsDisplayFormat = True

        StandardStringTextBoxEdit = New RepositoryItemTextEdit

    End Sub
    Private Sub WindowsUIButtonPanelSaveClose_ButtonClick(sender As Object, e As ButtonEventArgs) Handles WindowsUIButtonPanelStressNavigator.ButtonClick

        Dim ButSender As WindowsUIButton = TryCast(e.Button, DevExpress.XtraBars.Docking2010.WindowsUIButton)

        If ButSender Is Nothing Then

            Return

        End If

        Dim tag As String = ButSender.Tag.ToString()

        Select Case tag

            Case "Home"

                SetDeactivated()
                Me.Hide()
                FormMainScreen.BringToFront()

            Case "LiMVP"

                XtraTabControlStressTest.SelectedTabPage = XtraTabPageLMVP

            Case "SSList"

                XtraTabControlStressTest.SelectedTabPage = XtraTabPageSSL

            Case "Comp1"

                XtraTabControlStressTest.SelectedTabPage = XtraTabPageCompA

            Case "Comp2"

                XtraTabControlStressTest.SelectedTabPage = XtraTabPageCompB

            Case "MVPlan"

                XtraTabControlStressTest.SelectedTabPage = XtraTabPageMVP

            Case "MVDash"

                XtraTabControlStressTest.SelectedTabPage = XtraTabPageDashboard

        End Select

    End Sub
    Public Sub SetActive()

        'ExcelModels(ModelID).WBCalcEngine.CalcAuto()

    End Sub

    Public Sub SetDeactivated()

        'ExcelModels(ModelID).WBCalcEngine.CalcManual()

    End Sub

    Private Sub DeactivateST(sender As Object, e As EventArgs)


    End Sub

    Private Sub AddHandlers()

        AddHandler View_WrapCG_Stresses.CellValueChanged, AddressOf StressesGrid_EditingValueModified
        AddHandler TextEditMultivariableName.EditValueChanged, AddressOf TextEditMultivariableName_EditValueChanged
        AddHandler View_WrapCG_Stresses.CustomRowCellEdit, AddressOf GVStressesCustEditor
        AddHandler View_WrapCG_Stresses.ShowingEditor, AddressOf GVStressesShowingEditor
        AddHandler View_WrapCG_Stresses.CustomDrawCell, AddressOf CustomDrawStressesGrid
        AddHandler View_WrapTextGrid.CustomRowCellEdit, AddressOf GVTextGridCustEditor


    End Sub

    Sub GVTextGridCustEditor(sender As Object, e As CustomRowCellEditEventArgs)

        Dim view As GridView = TryCast(sender, GridView)
        Dim FN As Integer = e.Column.AbsoluteIndex


        If FN = 0 Or FN = 5 Then

            e.RepositoryItem = StandardStringTextBoxEdit

        End If

        If FN = 2 Or FN = 3 Then

            If e.RowHandle = 0 Or e.RowHandle = 2 Or e.RowHandle = 4 Then

                e.RepositoryItem = StandardPercentageTextBoxEdit

            End If

            If e.RowHandle = 1 Or e.RowHandle = 3 Or e.RowHandle = 5 Then

                e.RepositoryItem = StandardIntegerTextBoxEdit

            End If

        End If

        If FN = 7 Or FN = 8 Then

            If e.RowHandle = 0 Or e.RowHandle = 2 Then

                e.RepositoryItem = Standard2digitnumberTextBoxEdit

            End If

            If e.RowHandle = 1 Or e.RowHandle = 3 Or e.RowHandle = 4 Then

                e.RepositoryItem = StandardIntegerTextBoxEdit

            End If

        End If
    End Sub

    Sub CustomDrawStressesGrid(sender As Object, e As RowCellCustomDrawEventArgs)

        Dim view As CustomGridView = TryCast(sender, CustomGridView)
        Dim TestVal As String = ""

        Dim PadRect As New Rectangle With {
            .X = e.Bounds.X - DefaultGridCellPadding - 2,
            .Width = e.Bounds.Width + (2 * DefaultGridCellPadding) + 4,
            .Height = e.Bounds.Height + 4 + (2 * DefaultGridCellPadding),
            .Y = e.Bounds.Y - (DefaultGridCellPadding + 2)
            }

        Dim FN As String = e.Column.FieldName

        If FN = "Column 4" Or FN = "Column 5" Or FN = "Column 6" Or FN = "Column 7" Then

            TestVal = view.GetRowCellValue(e.RowHandle, "Column 1")

            If TestVal = "Dvpt Property Sales Income" Or TestVal = "Initial New Lettings Rate" Or TestVal = "Base Rents" Or TestVal = "Base Build Costs" Or TestVal = "Base Sales Value" Or TestVal = "Base Grant" Then

                e.Appearance.BackColor = Color.White
                e.Appearance.ForeColor = Color.White
                e.DefaultDraw()
                e.Handled = True
                Exit Sub

            End If

        End If

        If FN = "Column 4" Or FN = "Column 5" Or FN = "Column 7" Then

            TestVal = view.GetRowCellValue(e.RowHandle, "Column 3").ToString

            If Not IsNumeric(TestVal) Then

                e.Appearance.BackColor = Color.Lavender
                e.Appearance.ForeColor = Color.WhiteSmoke
                e.DefaultDraw()
                e.Handled = True
                Exit Sub
            End If

            If CDbl(TestVal) = 0 Then

                e.Appearance.BackColor = Color.Lavender
                e.Appearance.ForeColor = Color.WhiteSmoke
                e.DefaultDraw()
                e.Handled = True
                Exit Sub
            End If

        End If

        If FN = "Column 6" Then

            TestVal = view.GetRowCellValue(e.RowHandle, "Column 4").ToString

            If Not IsNumeric(TestVal) Then

                e.Appearance.BackColor = Color.Lavender
                e.Appearance.ForeColor = Color.WhiteSmoke
                e.DefaultDraw()
                e.Handled = True
                Exit Sub
            End If

            If CDbl(TestVal) = 0 Then

                e.Appearance.BackColor = Color.Lavender
                e.Appearance.ForeColor = Color.WhiteSmoke
                e.DefaultDraw()
                e.Handled = True
                Exit Sub
            End If

        End If

        e.DefaultDraw()

    End Sub


    Sub GVStressesCustEditor(sender As Object, e As CustomRowCellEditEventArgs)

        Dim view As CustomGridView = TryCast(sender, CustomGridView)

        If e.Column.FieldName = "Column 3" Then

            Dim TestVal As String = view.GetRowCellValue(e.RowHandle, "Column 2")
            If TestVal = "delay by 1 year" Then
                e.RepositoryItem = StandardYesEmptyEdit
            Else
                e.RepositoryItem = StandardPercentSpinEdit
            End If

        ElseIf e.Column.FieldName = "Column 4" Then

            Dim TestVal As String = view.GetRowCellValue(e.RowHandle, "Column 1")

            If TestVal = "Dvpt Property Sales Income" Or TestVal = "Initial New Lettings Rate" Or TestVal = "Base Rents" Or TestVal = "Base Build Costs" Or TestVal = "Base Sales Value" Or TestVal = "Base Grant" Then

                e.RepositoryItem = Nothing

            Else

                e.RepositoryItem = StandardPercentSpinEdit

            End If

        ElseIf e.Column.FieldName = "Column 5" Or e.Column.FieldName = "Column 6" Or e.Column.FieldName = "Column 7" Then

            Dim TestVal As String = view.GetRowCellValue(e.RowHandle, "Column 1")

            If TestVal = "Dvpt Property Sales Income" Or TestVal = "Initial New Lettings Rate" Or TestVal = "Base Rents" Or TestVal = "Base Build Costs" Or TestVal = "Base Sales Value" Or TestVal = "Base Grant" Then

                e.RepositoryItem = Nothing

            End If

        End If

    End Sub

    Sub GVStressesShowingEditor(sender As Object, e As CancelEventArgs)

        Dim view As GridView = DirectCast(sender, GridView)

        Dim CurrField As String = view.FocusedColumn.FieldName

        Dim TestVal As String = view.GetFocusedRowCellValue("Column 1")

        If CurrField = "Column 4" Or CurrField = "Column 5" Or CurrField = "Column 6" Or CurrField = "Column 7" Then

            If TestVal = "Dvpt Property Sales Income" Or TestVal = "Initial New Lettings Rate" Or TestVal = "Base Rents" Or TestVal = "Base Build Costs" Or TestVal = "Base Sales Value" Or TestVal = "Base Grant" Then
                e.Cancel = True
                Return
            End If

        End If

        If CurrField = "Column 4" Or CurrField = "Column 5" Or CurrField = "Column 7" Then

            TestVal = view.GetFocusedRowCellValue("Column 3").ToString

            If Not IsNumeric(TestVal) Then
                e.Cancel = True
                Return
            End If

            If CDbl(TestVal) = 0 Then
                e.Cancel = False
                Return
            End If

        End If

        If CurrField = "Column 6" Then

            TestVal = view.GetFocusedRowCellValue("Column 4").ToString

            If Not IsNumeric(TestVal) Then
                e.Cancel = True
                Return
            End If

            If CDbl(TestVal) = 0 Then
                e.Cancel = False
                Return
            End If

        End If
    End Sub
    Private Sub StressesGrid_EditingValueModified(sender As Object, e As EventArgs)

        CalcChanges()

    End Sub

    Private Sub CalcChanges()

        Me.Cursor = Cursors.WaitCursor

        ActiveWorkbook.Calculate()

        Me.Cursor = Cursors.Default

    End Sub

    Sub RefreshDataSources()

        If STMode = "Y" Then

            ProcessBreachesGrid(True)

        Else

            ProcessBreachesGrid(False)

        End If

        'DSMitDataRange.Refresh()
        'DSMitDevDataRange.Refresh()
        'DSMitMoneyDataRange.Refresh()
        'DSStressesDataRange.Refresh()
        'DSTextOutputsDataRange.Refresh()
        'DSBreachOutputsDataRange.Refresh()
    End Sub

    Sub BuildHTMLRenders()

        RenderStressHeaderHTMLData()

    End Sub
    Sub Form_InitilisationProcess_SetDataSources()


        'Tab #1

        Dim worksheet As DevExpress.Spreadsheet.Worksheet = ActiveWorkbook.Worksheets("Live Multivariable Planner")

        Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range(ExcelModels(ModelID).WBStructure.StressTestDefinition.StresstestMitigationsRange)

        Dim RDSOptions As New RangeDataSourceOptions With {
            .UseFirstRowAsHeader = False,
            .PreserveFormulas = False,
            .SkipHiddenRows = True,
            .SkipHiddenColumns = True,
            .EditingOptions = DataSourceEditingOptions.AllowEdit
        }

        DSMitDataRange = range.GetDataSource(RDSOptions)

        range = worksheet.Range(ExcelModels(ModelID).WBStructure.StressTestDefinition.StresstestMitigationsDevRange)
        DSMitDevDataRange = range.GetDataSource(RDSOptions)

        range = worksheet.Range(ExcelModels(ModelID).WBStructure.StressTestDefinition.StresstestMitigationsMoneyRange)
        DSMitMoneyDataRange = range.GetDataSource(RDSOptions)

        range = worksheet.Range(ExcelModels(ModelID).WBStructure.StressTestDefinition.StresstestStressesRange)
        DSStressesDataRange = range.GetDataSource(RDSOptions)

        Dim ColList As New List(Of String) From {
            "Blank A",
            "Blank B",
            "Target ",
            "Current ",
            "Blank C",
            "Blank D",
            "Blank E",
            "Target",
            "Current"
        }

        Dim ColMap As String = "TTTTTTTTT"

        Dim ColNames As New SourceColumnDetector(ColList, ColMap)

        RDSOptions.DataSourceColumnTypeDetector = ColNames

        range = worksheet.Range(ExcelModels(ModelID).WBStructure.StressTestDefinition.StresstestOutputTextRange)
        DSTextOutputsDataRange = range.GetDataSource(RDSOptions)

        RDSOptions.DataSourceColumnTypeDetector = Nothing

        worksheet = ActiveWorkbook.Worksheets("OW - Live Covenant Calculation")

        range = worksheet.Range(ExcelModels(ModelID).WBStructure.StressTestDefinition.StresstestCovenantGraphsDataRange)
        CovGraphsDataRange = range.GetDataSource(RDSOptions)

        AddHandler CovGraphsDataRange.ListChanged, AddressOf CovGraphsDataRange_DataSourceChanged

        ModelPostingComboBoxSelectCovenant.SetModelID = ModelID
        ModelPostingComboBoxSelectCovenant.SetTargetWorksheet = "Live Multivariable Planner"
        ModelPostingComboBoxSelectCovenant.SetTargetCell = "AD3"
        ModelPostingComboBoxSelectCovenant.InitialiseFromNRP("CovenantSelect")
        ModelPostingComboBoxSelectCovenant.SetLimitToList = True
        ModelPostingComboBoxSelectCovenant.ProcesDefValue()

        ColList = New List(Of String) From {
            "Year",
            "Mvt",
            "Base",
            "Live",
            "Target",
            "Met/Breach"
        }

        ColMap = "TTTTTT"
        ColNames = New SourceColumnDetector(ColList, ColMap)
        RDSOptions.DataSourceColumnTypeDetector = ColNames

        worksheet = ActiveWorkbook.Worksheets("Live Multivariable Planner")
        range = worksheet.Range(ExcelModels(ModelID).WBStructure.StressTestDefinition.StresstestBreachOutputRange)

        DSBreachOutputsDataRange = range.GetDataSource(RDSOptions)

        RDSOptions.DataSourceColumnTypeDetector = Nothing

        worksheet = ActiveWorkbook.Worksheets("Stress Sensitivity List")
        range = worksheet.Range("StressSensitivity")

        DSStressSenitivityLiveDataRange = range.GetDataSource(RDSOptions)

        'AddHandler DSStressSenitivityLiveDataRange.ListChanged, AddressOf RenderStressHeaderHTMLData

        UnhideColumnsCommand()
        ProcessMitigationsGrid()
        ProcessStressesGrid()
        ProcesstextOutputGrid()

        If STMode = "Y" Then

            ProcessBreachesGrid(True)

        Else

            ProcessBreachesGrid(False)

        End If

        'Tab #2




    End Sub

    Sub UnhideColumnsCommand()

        Dim WSTarget As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets("Live Multivariable Planner")
        ' Unhide the columns in the range
        Try
            ExcelModels(ModelID).WB.Worksheets("Live Multivariable Planner").Columns.Unhide(25, 32)
        Catch ex As Exception

        End Try

    End Sub
    Sub CovGraphsDataRange_DataSourceChanged(sender As Object, e As EventArgs)

        BuildCovCharts()

    End Sub


    Sub ProcessBreachesGrid(OnOff As Boolean)

        Exit Sub

        If OnOff = False Then

            GridControlBreaches.DataSource = Nothing
            GridControlBreaches.Enabled = False
            GridControlBreaches.Visible = False

            Exit Sub

        Else

            GridControlBreaches.Enabled = True
            GridControlBreaches.Visible = True
            GridControlBreaches.DataSource = DSBreachOutputsDataRange

            Dim GV As GridView = GridControlBreaches.MainView

            GV.Columns(0).Caption = "Year"
            GV.Columns(1).Caption = "Mvt"
            GV.Columns(2).Caption = "Base"
            GV.Columns(3).Caption = "Live"
            GV.Columns(4).Caption = "Target"
            GV.Columns(5).Caption = "Met/Breach"

            GV.Columns(0).OptionsColumn.ReadOnly = True
            GV.Columns(1).OptionsColumn.ReadOnly = True
            GV.Columns(2).OptionsColumn.ReadOnly = True
            GV.Columns(3).OptionsColumn.ReadOnly = True
            GV.Columns(4).OptionsColumn.ReadOnly = True
            GV.Columns(5).OptionsColumn.ReadOnly = True

            GV.Columns(2).AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            GV.Columns(2).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            GV.Columns(2).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric

            GV.Columns(2).DisplayFormat.FormatString = "p2"
            GV.Columns(3).AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            GV.Columns(3).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            GV.Columns(3).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
            GV.Columns(3).DisplayFormat.FormatString = "p2"
            GV.Columns(4).AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            GV.Columns(4).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            GV.Columns(4).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
            GV.Columns(4).DisplayFormat.FormatString = "p2"

            GV.Columns(5).AppearanceCell.ForeColor = Color.Red

            GV.Columns(1).AppearanceCell.ForeColor = Color.Red
            GV.Columns(1).AppearanceCell.Font = New System.Drawing.Font("Wingdings", 13, FontStyle.Regular)
            GV.BestFitColumns()

            UpdateGridSizeNonWr(GV, GridControlTextOut)
            Formatter.FormatGridView(GV, GridControlBreaches)


        End If









    End Sub
    Public Sub ProcessMitigationsGrid()


        WrapCG_Mits = New CustomGridWrapper

        WrapCG_Mits.WrappedCGC.DataSource = DSMitDataRange
        WrapCG_Mits.Height = XtraTabPageMitigations.Height * 0.6
        View_WrapCG_Mits = WrapCG_Mits.WrappedGridView

        WrapCG_Mits.Dock = DockStyle.None
        WrapCG_Mits.Width = XtraTabPageMitigations.Width

        'GridView_InitialisationProcess_AddHandlers(View_WrapCG_BS)




        Me.XtraTabPageMitigations.Controls.Add(WrapCG_Mits)
        View_WrapCG_Mits.Columns(0).Caption = "Class"
        View_WrapCG_Mits.Columns(0).OptionsColumn.ReadOnly = True
        View_WrapCG_Mits.Columns(1).Caption = "Mitigation"
        View_WrapCG_Mits.Columns(1).OptionsColumn.ReadOnly = True
        View_WrapCG_Mits.Columns(2).Caption = "Type"
        View_WrapCG_Mits.Columns(2).OptionsColumn.ReadOnly = True
        View_WrapCG_Mits.Columns(3).Caption = "Change 1 %"

        View_WrapCG_Mits.Columns(3).ColumnEdit = New RepositoryItemSpinEdit With {
                                    .MinValue = -100,
                                    .Increment = CDec(0.0025),
                                    .MaxValue = 500,
                                    .EditMask = "p2",
                                    .UseMaskAsDisplayFormat = True
                                    }

        View_WrapCG_Mits.Columns(3).AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        View_WrapCG_Mits.Columns(3).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        View_WrapCG_Mits.Columns(3).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        View_WrapCG_Mits.Columns(3).DisplayFormat.FormatString = "p2"

        View_WrapCG_Mits.Columns(4).Caption = "Change 2 %"

        View_WrapCG_Mits.Columns(4).ColumnEdit = New RepositoryItemSpinEdit With {
                                    .MinValue = -100,
                                    .Increment = CDec(0.0025),
                                    .MaxValue = 500,
                                    .EditMask = "p2",
                                    .UseMaskAsDisplayFormat = True
                                    }

        View_WrapCG_Mits.Columns(4).AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        View_WrapCG_Mits.Columns(4).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        View_WrapCG_Mits.Columns(4).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        View_WrapCG_Mits.Columns(4).DisplayFormat.FormatString = "p2"

        View_WrapCG_Mits.Columns(5).Caption = "Change 1" & vbLf & "from year"


        Dim EditControl As RepositaryItems.AbovoRespositaryItem

        EditControl = RepositaryItems.GetEditor("Rep_OrdinalYears", ModelID)
        WrapCG_Mits.WrappedCGC.RepositoryItems.Add(EditControl.RetCombo)
        View_WrapCG_Mits.Columns(5).ColumnEdit = EditControl.RetCombo

        View_WrapCG_Mits.Columns(6).Caption = "Change 2" & vbLf & "from year"
        View_WrapCG_Mits.Columns(6).ColumnEdit = EditControl.RetCombo

        View_WrapCG_Mits.Columns(7).Caption = "to year"
        View_WrapCG_Mits.Columns(7).ColumnEdit = EditControl.RetCombo

        Formatter.FormatGridView(View_WrapCG_Mits, WrapCG_Mits.WrappedCGC)

        View_WrapCG_Mits.BestFitColumns()

        UpdateGridSize(View_WrapCG_Mits, WrapCG_Mits)

        ' The Development bit

        WrapCG_MitsDev = New CustomGridWrapper

        WrapCG_MitsDev.WrappedCGC.DataSource = DSMitDevDataRange

        Dim View_WrapCG_MitsDev As CustomGridView = WrapCG_MitsDev.WrappedGridView

        'Formatter.FormatGridView(View_WrapCG_MitsDev, WrapCG_MitsDev.WrappedCGC)

        View_WrapCG_MitsDev.OptionsView.ShowColumnHeaders = False

        WrapCG_MitsDev.Dock = DockStyle.None

        WrapCG_MitsDev.Width = XtraTabPageMitigations.Width

        WrapCG_MitsDev.Top = WrapCG_Mits.Bottom
        WrapCG_MitsDev.Height = XtraTabPageMitigations.Height * 0.2
        UpdateGridSize(View_WrapCG_MitsDev, WrapCG_MitsDev)


        Me.XtraTabPageMitigations.Controls.Add(WrapCG_MitsDev)



        View_WrapCG_MitsDev.Columns(0).OptionsColumn.ReadOnly = True
        View_WrapCG_MitsDev.Columns(1).OptionsColumn.ReadOnly = True
        View_WrapCG_MitsDev.Columns(2).OptionsColumn.ReadOnly = True
        Formatter.FormatGridView(View_WrapCG_MitsDev, WrapCG_MitsDev.WrappedCGC)

        EditControl = RepositaryItems.GetEditor("Rep_YesNo", ModelID)
        WrapCG_MitsDev.WrappedCGC.RepositoryItems.Add(EditControl.RetCombo)
        View_WrapCG_MitsDev.Columns(3).ColumnEdit = EditControl.RetCombo

        View_WrapCG_MitsDev.Columns(4).OptionsColumn.ReadOnly = True
        View_WrapCG_MitsDev.Columns(5).OptionsColumn.ReadOnly = True
        View_WrapCG_MitsDev.Columns(6).OptionsColumn.ReadOnly = True
        View_WrapCG_MitsDev.Columns(7).OptionsColumn.ReadOnly = True

        View_WrapCG_MitsDev.Columns(0).Width = View_WrapCG_Mits.Columns(0).Width
        View_WrapCG_MitsDev.Columns(1).Width = View_WrapCG_Mits.Columns(1).Width
        View_WrapCG_MitsDev.Columns(2).Width = View_WrapCG_Mits.Columns(2).Width
        View_WrapCG_MitsDev.Columns(3).Width = View_WrapCG_Mits.Columns(3).Width
        View_WrapCG_MitsDev.Columns(4).Width = View_WrapCG_Mits.Columns(4).Width
        View_WrapCG_MitsDev.Columns(5).Width = View_WrapCG_Mits.Columns(5).Width
        View_WrapCG_MitsDev.Columns(6).Width = View_WrapCG_Mits.Columns(6).Width
        View_WrapCG_MitsDev.Columns(7).Width = View_WrapCG_Mits.Columns(7).Width

        UpdateGridSize(View_WrapCG_Mits, WrapCG_Mits)
        ' The Moneyelopment bit

        WrapCG_MitsMoney = New CustomGridWrapper

        WrapCG_MitsMoney.WrappedCGC.DataSource = DSMitMoneyDataRange

        Dim View_WrapCG_MitsMoney As CustomGridView = WrapCG_MitsMoney.WrappedGridView

        'Formatter.FormatGridView(View_WrapCG_MitsMoney, WrapCG_MitsMoney.WrappedCGC)

        View_WrapCG_MitsMoney.OptionsView.ShowColumnHeaders = False

        WrapCG_MitsMoney.Dock = DockStyle.None

        WrapCG_MitsMoney.Width = WrapCG_MitsDev.Width
        WrapCG_MitsMoney.WrappedCGC.Width = WrapCG_MitsDev.Width

        WrapCG_MitsMoney.Height = XtraTabPageMitigations.Height * 0.4
        WrapCG_MitsMoney.Top = WrapCG_MitsDev.Bottom

        Me.XtraTabPageMitigations.Controls.Add(WrapCG_MitsMoney)

        WrapCG_MitsMoney.Width = WrapCG_MitsDev.Width
        WrapCG_MitsMoney.WrappedCGC.Width = WrapCG_MitsDev.Width
        WrapCG_MitsMoney.Height = XtraTabPageMitigations.Height * 0.4
        WrapCG_MitsMoney.Top = WrapCG_MitsDev.Bottom


        Formatter.FormatGridView(View_WrapCG_MitsMoney, WrapCG_MitsMoney.WrappedCGC)
        UpdateGridSize(View_WrapCG_MitsMoney, WrapCG_MitsMoney)

        View_WrapCG_MitsMoney.Columns(0).OptionsColumn.ReadOnly = True
        View_WrapCG_MitsMoney.Columns(1).OptionsColumn.ReadOnly = True
        View_WrapCG_MitsMoney.Columns(2).OptionsColumn.ReadOnly = True
        View_WrapCG_MitsMoney.Columns(3).OptionsColumn.ReadOnly = True
        View_WrapCG_MitsMoney.Columns(4).OptionsColumn.ReadOnly = False
        View_WrapCG_MitsMoney.Columns(5).OptionsColumn.ReadOnly = False
        View_WrapCG_MitsMoney.Columns(6).OptionsColumn.ReadOnly = False
        View_WrapCG_MitsMoney.Columns(7).OptionsColumn.ReadOnly = False

        View_WrapCG_MitsMoney.Columns(0).Width = View_WrapCG_Mits.Columns(0).Width
        View_WrapCG_MitsMoney.Columns(1).Width = View_WrapCG_Mits.Columns(1).Width
        View_WrapCG_MitsMoney.Columns(2).Width = View_WrapCG_Mits.Columns(2).Width
        View_WrapCG_MitsMoney.Columns(3).Width = View_WrapCG_Mits.Columns(3).Width
        View_WrapCG_MitsMoney.Columns(4).Width = View_WrapCG_Mits.Columns(4).Width
        View_WrapCG_MitsMoney.Columns(5).Width = View_WrapCG_Mits.Columns(5).Width
        View_WrapCG_MitsMoney.Columns(6).Width = View_WrapCG_Mits.Columns(6).Width
        View_WrapCG_MitsMoney.Columns(7).Width = View_WrapCG_Mits.Columns(7).Width

    End Sub
    Sub ProcessStressesGrid()

        Dim WrapCG_Stresses = New CustomGridWrapper

        WrapCG_Stresses.WrappedCGC.DataSource = DSStressesDataRange

        View_WrapCG_Stresses = WrapCG_Stresses.WrappedGridView



        WrapCG_Stresses.Dock = DockStyle.Fill

        'GridView_InitialisationProcess_AddHandlers(View_WrapCG_BS)




        Me.XtraTabPageStresses.Controls.Add(WrapCG_Stresses)
        View_WrapCG_Stresses.Columns(0).Caption = "Class"
        View_WrapCG_Stresses.Columns(0).OptionsColumn.ReadOnly = True
        View_WrapCG_Stresses.Columns(1).Caption = "Stress"
        View_WrapCG_Stresses.Columns(1).OptionsColumn.ReadOnly = True
        View_WrapCG_Stresses.Columns(2).Caption = "Type"
        View_WrapCG_Stresses.Columns(2).OptionsColumn.ReadOnly = True
        View_WrapCG_Stresses.Columns(3).Caption = "Change 1 %"

        View_WrapCG_Stresses.Columns(3).ColumnEdit = StandardYesEmptyEdit


        View_WrapCG_Stresses.Columns(3).AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        View_WrapCG_Stresses.Columns(3).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        View_WrapCG_Stresses.Columns(3).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        View_WrapCG_Stresses.Columns(3).DisplayFormat.FormatString = "p2"

        View_WrapCG_Stresses.Columns(4).Caption = "Change 2 %"

        View_WrapCG_Stresses.Columns(4).ColumnEdit = StandardYesEmptyEdit

        View_WrapCG_Stresses.Columns(4).AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        View_WrapCG_Stresses.Columns(4).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        View_WrapCG_Stresses.Columns(4).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        View_WrapCG_Stresses.Columns(4).DisplayFormat.FormatString = "p2"

        View_WrapCG_Stresses.Columns(5).Caption = "Change 1" & vbLf & "from year"


        Dim EditControl2 As RepositaryItems.AbovoRespositaryItem

        EditControl2 = RepositaryItems.GetEditor("Rep_OrdinalYears", ModelID)
        WrapCG_Stresses.WrappedCGC.RepositoryItems.Add(EditControl2.RetCombo)
        View_WrapCG_Stresses.Columns(5).ColumnEdit = EditControl2.RetCombo

        View_WrapCG_Stresses.Columns(6).Caption = "Change 2" & vbLf & "from year"
        View_WrapCG_Stresses.Columns(6).ColumnEdit = EditControl2.RetCombo

        View_WrapCG_Stresses.Columns(7).Caption = "to year"
        View_WrapCG_Stresses.Columns(7).ColumnEdit = EditControl2.RetCombo

        Formatter.FormatGridView(View_WrapCG_Stresses, WrapCG_Stresses.WrappedCGC)
        View_WrapCG_Stresses.BestFitColumns()





    End Sub
    Sub ProcesstextOutputGrid()

        GridControlTextOut.DataSource = DSTextOutputsDataRange
        View_WrapTextGrid = GridControlTextOut.MainView
        Formatter.FormatGridView(View_WrapTextGrid, GridControlTextOut)
        View_WrapTextGrid.Columns(0).Caption = " "
        View_WrapTextGrid.Columns(1).Caption = " "
        View_WrapTextGrid.Columns(2).Caption = "Target"
        View_WrapTextGrid.Columns(3).Caption = "Current"
        View_WrapTextGrid.Columns(4).Caption = " "
        View_WrapTextGrid.Columns(5).Caption = " "
        View_WrapTextGrid.Columns(6).Caption = " "
        View_WrapTextGrid.Columns(7).Caption = "Target"
        View_WrapTextGrid.Columns(8).Caption = "Current"
        View_WrapTextGrid.BestFitColumns()

        UpdateGridSizeNonWr(View_WrapTextGrid, GridControlTextOut)

        GridFormatter.ClearBlankHeaders(GridControlTextOut)

    End Sub
    Public Sub Clearup()

        On Error Resume Next

        RemoveHandler CovGraphsDataRange.ListChanged, AddressOf CovGraphsDataRange_DataSourceChanged 'RemoveHandler CovGraphsDataRange.ListChanged, AddressOf CovGraphsDataRange_DataSourceChanged

        If Not IsNothing(WrapCG_Mits) Then
            WrapCG_Mits.Dispose()
            WrapCG_Mits = Nothing
        End If

        If Not IsNothing(WrapCG_MitsDev) Then
            WrapCG_MitsDev.Dispose()
            WrapCG_MitsDev = Nothing
        End If

        If Not IsNothing(WrapCG_MitsMoney) Then
            WrapCG_MitsMoney.Dispose()
            WrapCG_MitsMoney = Nothing
        End If

        If Not IsNothing(CovGraphsDataRange) Then
            CovGraphsDataRange.Dispose()
            CovGraphsDataRange = Nothing
        End If

    End Sub


    Sub ManualDispose()

        If STMode = "N" Then

            Try
                ActiveWorkbook.Worksheets("Live Multivariable Planner").Columns.Hide(25, 32)
            Catch ex As Exception

            End Try

        End If

        Me.CovGraphsDataRange = Nothing
        Me.DSBreachOutputsDataRange = Nothing

    End Sub

    Private Sub UpdateGridSize(GV As CustomGridView, VGC As CustomGridWrapper)

        Dim viewInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridViewInfo = CType(GV.GetViewInfo(), GridViewInfo)
        Dim fi As FieldInfo = GetType(GridView).GetField("scrollInfo", BindingFlags.Instance Or BindingFlags.NonPublic)
        Dim scrollInfo As ScrollInfo = DirectCast(fi.GetValue(GV), ScrollInfo)

        Dim _height As Integer = viewInfo.CalcRealViewHeight(New System.Drawing.Rectangle(0, 0, Int32.MaxValue, Int32.MaxValue))

        If scrollInfo.HScrollVisible Then

            _height += scrollInfo.HScrollRect.Height + 50

        End If

        VGC.WrappedCGC.Height = _height
        VGC.Height = _height
        GV.LayoutChanged()

    End Sub
    Public Sub BuildCovCharts()

        Dim CovChart As ChartControl = Nothing

        Dim CoveWS As DevExpress.Spreadsheet.Worksheet = ActiveWorkbook.Worksheets("OW - Live Covenant Calculation")
        Dim DataRange As DevExpress.Spreadsheet.CellRange = Nothing
        Dim AddTitle As String = ""

        For ChartX = 1 To 5

            Select Case ChartX

                Case 1

                    CovChart = CvntChart1
                    DataRange = CoveWS.Range("I7: I46")
                    AddTitle = "Gearing Covenant Compliance"

                Case 2
                    CovChart = CvntChart2
                    DataRange = CoveWS.Range("M7:M46")
                    AddTitle = "Op Margin Covenant Compliance"

                Case 3
                    CovChart = CvntChart3
                    DataRange = CoveWS.Range("Q7:Q46")
                    AddTitle = "EBITDA MRI Covenant Compliance"

                Case 4
                    CovChart = CvntChart4
                    DataRange = CoveWS.Range("U7:U46")
                    AddTitle = "Debt / Unit Covenant Compliance"

                Case 5
                    CovChart = CvntChart5
                    DataRange = CoveWS.Range("Y7:Y46")
                    AddTitle = "Debt Covenant Compliance"

            End Select

            CovChart.Series.Clear()

            ' Create a line series.
            Dim series1 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Okay", ViewType.ScatterLine)
            Dim series2 As DevExpress.XtraCharts.Series = New DevExpress.XtraCharts.Series("Breach", ViewType.ScatterLine)

            Dim Datapoints As New List(Of DataPoint)
            Dim CellExamine As DevExpress.Spreadsheet.Cell


            Dim NeDp As DataPoint

            For x = 0 To 39
                NeDp = New DataPoint
                CellExamine = DataRange(x, 0)
                If CellExamine.DisplayText = "#N/A" Then
                    series1.Points.Add(New SeriesPoint(x + 1, 0.5))
                Else
                    series2.Points.Add(New SeriesPoint(x + 1, 0.5))
                End If

            Next

            ' Add the series to the chart.
            CovChart.Series.Add(series1)
            CovChart.Series.Add(series2)

            ' Set the numerical argument scale types for the series,
            ' as it is qualitative, by default.
            series1.ArgumentScaleType = ScaleType.Auto
            ' Access the view-type-specific options of the series.
            CType(series1.View, LineSeriesView).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True
            CType(series2.View, LineSeriesView).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True
            CType(series1.View, LineSeriesView).LineMarkerOptions.Size = 20
            CType(series1.View, LineSeriesView).LineMarkerOptions.Color = Color.Green
            CType(series1.View, LineSeriesView).LineMarkerOptions.Kind = MarkerKind.Circle
            CType(series2.View, LineSeriesView).LineMarkerOptions.Size = 20
            CType(series2.View, LineSeriesView).LineMarkerOptions.Color = Color.Red
            CType(series2.View, LineSeriesView).LineMarkerOptions.Kind = MarkerKind.Cross
            'CType(series1.View, LineSeriesView).LineStyle.DashStyle = DevExpress.XtraCharts.DashStyle.Empty
            CType(series1.View, LineSeriesView).LineStyle.Thickness = 1
            ' Access the view-type-specific options of the series.
            CType(CovChart.Diagram, XYDiagram).AxisY.WholeRange.SetMinMaxValues(0, 1)
            CType(CovChart.Diagram, XYDiagram).AxisY.Visibility = DevExpress.Utils.DefaultBoolean.False
            CType(CovChart.Diagram, XYDiagram).AxisX.Visibility = DevExpress.Utils.DefaultBoolean.False
            CType(CovChart.Diagram, XYDiagram).AxisX.Tickmarks.Visible = False
            CType(CovChart.Diagram, XYDiagram).AxisY.Tickmarks.Visible = False
            CType(CovChart.Diagram, XYDiagram).AxisX.GridLines.Visible = False
            CType(CovChart.Diagram, XYDiagram).AxisY.GridLines.Visible = False
            CType(CovChart.Diagram, XYDiagram).AxisY.InterlacedColor = Color.FromArgb(20, 60, 60, 60)
            CType(CovChart.Diagram, XYDiagram).AxisX.NumericScaleOptions.AutoGrid = False
            CType(CovChart.Diagram, XYDiagram).AxisX.Interlaced = False
            CType(CovChart.Diagram, XYDiagram).AxisY.Interlaced = False
            CType(CovChart.Diagram, XYDiagram).AxisX.NumericScaleOptions.GridSpacing = 1

            ' Hide the legend (if necessary).
            CovChart.Legend.Visibility = DevExpress.Utils.DefaultBoolean.False
            ' Add a title to the chart (if necessary).
            CovChart.Titles.Clear()
            CovChart.Titles.Add(New DevExpress.XtraCharts.ChartTitle())
            CovChart.Titles(0).Text = AddTitle
            CovChart.Titles(0).EnableAntialiasing = DevExpress.Utils.DefaultBoolean.True
            CovChart.Titles(0).DXFont = New DXFont("Tahoma", 10, DXFontStyle.Bold)
            CovChart.Height = 223
            ' Add the chart to the form.

            CovChart.Dock = DockStyle.Fill


        Next ChartX



    End Sub
    Private Class DataPoint
        Public Property Argument() As String
        Public Property Value() As Double

    End Class
    Private Sub UpdateGridSizeNonWr(GV As GridView, VGC As GridControl)

        Dim viewInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridViewInfo = CType(GV.GetViewInfo(), GridViewInfo)
        Dim fi As FieldInfo = GetType(GridView).GetField("scrollInfo", BindingFlags.Instance Or BindingFlags.NonPublic)
        Dim scrollInfo As ScrollInfo = DirectCast(fi.GetValue(GV), ScrollInfo)

        Dim _height As Integer = viewInfo.CalcRealViewHeight(New System.Drawing.Rectangle(0, 0, Int32.MaxValue, Int32.MaxValue))

        If scrollInfo.HScrollVisible Then

            _height += scrollInfo.HScrollRect.Height + 50

        End If

        VGC.Height = _height
        GV.LayoutChanged()

    End Sub

    Private Class SourceColumnDetector

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

    Protected Overrides Function CreateFormBorderPainter() As DevExpress.Skins.XtraForm.FormPainter

        Return New CustomFormPainterST(Me, LookAndFeel)

    End Function



    Public Property FormBorderColor() As Color

        Get
            Return MyColourSwatch
        End Get

        Set(ByVal value As Color)
            MyColourSwatch = value
        End Set

    End Property

    Sub ToggleStressTestMode()

        ' Toggles the stress test mode on or off
        ' and updates the UI accordingly

        If STMode = "Y" Then

            STMode = "N"

            StressTestModeSwitch(True)
            MsgBox("Stress Test Mode Deactivated")

        Else

            STMode = "Y"

            StressTestModeSwitch(True)
            MsgBox("Stress Test Mode Activated")

        End If

    End Sub
    Sub StressTestModeSwitch(DoChanges As Boolean)

        Me.Cursor = Cursors.WaitCursor

        'Switches the mode between Stress Test and Business Plan
        'and makes the necessary adjustments to the spreadsheet

        Dim ModeCell As DevExpress.Spreadsheet.CellRange = ActiveWorkbook.DefinedNames.GetDefinedName("StressTestMode").Range

        If STMode = "Y" Then

            If DoChanges Then

                ActiveWorkbook.DefinedNames.GetDefinedName("Mode").RefersTo = """Stress Test"""
                ModeCell(0, 0).Value = "Y"
                StressTestModeAdjustments()

            End If

            TextEditMultivariableName.EditValue = ActiveWorkbook.DefinedNames.GetDefinedName("NewStressName").Range(0, 0).Value.TextValue

            'SimpleButtonModeSwitch.Text = "Deactivate Stress Test Mode"
            PanelControlCovSel.Visible = True
            GridControlBreaches.Visible = True

        Else

            If DoChanges Then

                ActiveWorkbook.DefinedNames.GetDefinedName("Mode").RefersTo = """Business Plan"""

                ModeCell(0, 0).Value = "N"
                DeStressAdjustments()

            End If

            'SimpleButtonModeSwitch.Text = "Activate Stress Test Mode"
            PanelControlCovSel.Visible = False
            GridControlBreaches.Visible = False

        End If

        ActiveWorkbook.Calculate()
        BuildCovCharts()

        Me.Cursor = Cursors.Default

    End Sub

    Sub StressTestModeAdjustments()

        UNProtectWS(ModelID, "OW - Live Stress Reporting")

        Dim SourceRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.Range("StressSwitch")
        Dim DestRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.Range("StressBase")
        DestRange.CopyFrom(SourceRange, PasteSpecial.Values)

        ProtectWS(ModelID, "OW - Live Stress Reporting")


    End Sub

    Sub Stress_Sensitivity_Capture()

        Debug.Print("1")
        UNProtectWS(ModelID, "Stress Sensitivity List")

        Dim WB As IWorkbook = ExcelModels(ModelID).WB
        Debug.Print("2")
        WB.DefinedNames.GetDefinedName("StressSensitivityDate").Range(0.0).Value = Now()

RestartFrom:

        Dim TargetDefinedName As DevExpress.Spreadsheet.DefinedName = WB.DefinedNames.GetDefinedName("StressSensitivityData")
        Dim CRTarget As DevExpress.Spreadsheet.CellRange
        Dim WSTarget As DevExpress.Spreadsheet.Worksheet = WB.Worksheets("Stress Sensitivity List")

        Dim Lref As Integer = TargetDefinedName.Range.LeftColumnIndex
        Dim Rref As Integer = TargetDefinedName.Range.RightColumnIndex
        Dim Tref As Integer = TargetDefinedName.Range.TopRowIndex
        Dim Bref As Integer = TargetDefinedName.Range.BottomRowIndex

        TargetDefinedName = Nothing
        Debug.Print("3")
        CRTarget = WSTarget.Range.FromLTRB(Lref, Tref, Rref, Bref)
        Debug.Print("Target range from " & CRTarget.TopRowIndex & " to " & CRTarget.BottomRowIndex)
        'ExcelModels(ModelID).WB.BeginUpdate()

        Debug.Print("4")
        ' Find Empty Row
        Dim bottomRow As DevExpress.Spreadsheet.CellRange = Nothing

        Debug.Print("5")
        Debug.Print("Iterating from " & CRTarget.TopRowIndex & " to " & CRTarget.TopRowIndex + CRTarget.RowCount - 1)

        For i As Integer = CRTarget.TopRowIndex To CRTarget.TopRowIndex + CRTarget.RowCount - 1

            Dim rowCell As DevExpress.Spreadsheet.CellRange = WSTarget.Range.FromLTRB(CRTarget.LeftColumnIndex, i, CRTarget.RightColumnIndex, i)

            If rowCell.ToArray().All(Function(cell) cell.Value.IsEmpty) Then

                bottomRow = rowCell
                Exit For

            End If

            rowCell = Nothing

        Next
        Debug.Print("6")
        If (bottomRow Is Nothing) Then

            ' Extend
            Debug.Print("Bottom Row Is Nothing, extending the range")
            CRTarget = Nothing
            WSTarget = Nothing


            InsertRows(ModelID, "StressSensitivityData", 1, True)

            Debug.Print("rows inserted")
            Debug.Print("6a")
            GoTo RestartFrom

        End If

        Dim SourceRow As DevExpress.Spreadsheet.CellRange = WB.DefinedNames.GetDefinedName("StressSensitivity").Range
        Debug.Print("Copying from " & SourceRow.BottomRowIndex & " To " & bottomRow.BottomRowIndex)
        Debug.Print("7")
        bottomRow.CopyFrom(SourceRow, PasteSpecial.Values)

        bottomRow = Nothing
        SourceRow = Nothing

        CRTarget = Nothing
        WSTarget = Nothing
        WB = Nothing
        Debug.Print("8")
        ProtectWS(ModelID, "Stress Sensitivity List")

    End Sub


    Sub DeStressAdjustments()

        ExcelModels(ModelID).WB.Worksheets("Live Multivariable Planner").Columns.Hide(25, 32)
        ProcessBreachesGrid(False)

    End Sub

    Private Sub SimpleButtonModeSwitch_Click(sender As Object, e As EventArgs)


        ToggleStressTestMode()

    End Sub

    Private Sub StressTest_ResizeEnd(sender As Object, e As EventArgs) Handles MyBase.ResizeEnd

        MiddleObject(SimpleButtonQC, PanelControl2)
        'MiddleObject(SimpleButtonModeSwitch, PanelControl1)

        MiddleObject(SimpleButtonCapture, PanelControl3)
        MiddleObject(LabelControl1, PanelControl3)
        MiddleObject(SimpleButtonCapture, PanelControl3)

        MiddleObject(ModelPostingComboBoxSelectCovenant, PanelControlCovSel)
        MiddleObject(LabelControl2, PanelControlCovSel)

    End Sub

    Private Sub SimpleButtonQC_Click(sender As Object, e As EventArgs) Handles SimpleButtonQC.Click

        Me.Cursor = Cursors.WaitCursor

        Stress_Sensitivity_Capture()

        Me.Cursor = Cursors.Default

    End Sub

    Sub RenderStressHeaderHTMLData()
        ' This sub renders the header data for the stress test HTML output   

        Dim worksheet As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets("Stress Sensitivity List")
        Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range("A5:BI8")

        Dim RangeList As New List(Of DevExpress.Spreadsheet.CellRange)

        RangeList.Add(range)

        range = worksheet("StressSensitivity")

        RangeList.Add(range)

        Dim range2 As DevExpress.Spreadsheet.CellRange = worksheet("StressSensitivityData")

        Dim range3 As DevExpress.Spreadsheet.CellRange = worksheet.Range.FromLTRB(range.LeftColumnIndex, range2.TopRowIndex, range.RightColumnIndex, range2.BottomRowIndex)

        RangeList.Add(range3)

        WebBrowserStressCaptureOutput.DocumentText = RenderRangeCells(RangeList)



    End Sub


    'Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click

    '    'Debug.Print(WrapCG_MitsMoney.Width.ToString & " " & WrapCG_MitsMoney.Height.ToString)
    '    'Debug.Print(WrapCG_MitsDev.Width.ToString & " " & WrapCG_MitsDev.Height.ToString)
    '    'Debug.Print(WrapCG_Mits.Width.ToString & " " & WrapCG_Mits.Height.ToString)
    '    'Debug.Print(WrapCG_MitsMoney.WrappedCGC.Width.ToString & " " & WrapCG_MitsMoney.WrappedCGC.Height.ToString)
    '    'Debug.Print(WrapCG_MitsDev.WrappedCGC.Width.ToString & " " & WrapCG_MitsDev.WrappedCGC.Height.ToString)
    '    'Debug.Print(WrapCG_Mits.WrappedCGC.Width.ToString & " " & WrapCG_Mits.WrappedCGC.Height.ToString)
    'End Sub
    '    Sub StressTestDataCapture()

    '        'CodeSafe JW 26/4/22

    '        ' add Stress Test data
    '        ' launched by button on "Live Multivariable Planner" sheet
    '        Dim TestNumber As Integer

    '        On Error Resume Next




    '        If Range("StressTestNumber").Value = "" Then

    '            MsgBox("Please first Select a test number")
    '            Range("StressTestNumber").Select

    '        ElseIf Range("NewStressName").Value = "" Then

    '            MsgBox("Please first Select a Multivariable Name")
    '            Range("NewStressName").Select

    '        Else

    '            Sheets("Multivariable Planner").Unprotect Password:=PW
    '        Sheets("OW - Captured Data").Unprotect Password:=PW
    '        Sheets("Live Multivariable Planner").Unprotect Password:=PW

    '        TestNumber = CInt(Right(Range("StressTestNumber").Value, 2))

    '            Range("Assumptions" & TestNumber) = Range("LiveAssumptions").Value
    '            Range("AssumptionsA" & TestNumber) = Range("LiveAssumptionsA").Value
    '            Sheets("Multivariable Planner").Activate
    '            Range("Assumptions" & TestNumber).Cells(-1, 1) = Range("NewStressName").Value
    '            Range("A5").Select



    '            Range("StressLiveInfo").Copy
    '            Range("S" & TestNumber & "Data").PasteSpecial xlValues

    '        Range("NewStressName").ClearContents
    '            Range("LiveAssumptions").ClearContents
    '            Range("LiveAssumptionsA").ClearContents



    '            '        Range("StressLiveInfo").Copy
    '            '        Range("S0Data").PasteSpecial xlValues

    '            Call fncPrtSht("Multivariable Planner")


    '            Call fncPrtSht("OW - Captured Data")


    '            Call fncPrtSht("Live Multivariable Planner")


    '            MsgBox("Assumptions captured")

    '        End If




    '    End Sub
    '    Sub StressTestAssumptionsCapture()

    '        'codesafe JW 26/4/22

    '        ' add Stress Test data
    '        ' launched by button on "Multivariable Planner" sheet

    '        Dim TestNumber As Integer
    '        Dim Scenario As String, x

    '        Dim bPrepareBase As Boolean

    '        bPrepareBase = True


    '        On Error Resume Next






    '        Sheets("Multivariable Planner").Unprotect Password:=PW
    '    Sheets("OW - Captured Data").Unprotect Password:=PW
    '    Sheets("Live Multivariable Planner").Unprotect Password:=PW

    '    For TestNumber = 1 To 10




    '            If Range("ImportMode" & TestNumber) <> "Use assumptions below" Then

    '                Sheets("Multivariable Planner").Activate
    '                Range("Assumptions" & TestNumber).ClearContents
    '                Range("StressTestAssumptions").Select

    '                Sheets("OW - Captured Data").Activate

    '                Scenario = "S" & TestNumber

    '                Call Import_Separate_Plans(Scenario)



    '            Else

    '                x = Len(Range("R_StressTests").Cells(1, 5 * TestNumber + 1).Value)

    '                If (Range("R_StressTests").Cells(1, 5 * TestNumber + 1).Value) <> "" Then

    '                    ' Application.Calculation = xlCalculationManual
    '                    If bPrepareBase = True Then

    '                        Range("LiveAssumptions").ClearContents
    '                        Range("LiveAssumptionsA").ClearContents



    '                        Range("StressLiveInfo").Copy
    '                        Range("S0Data").PasteSpecial xlValues
    '                    bPrepareBase = False

    '                    End If

    '                    Range("StressTestMode") = "Y"
    '                    Range("Assumptions" & TestNumber).Copy
    '                    Range("LiveAssumptions").PasteSpecial xlValues
    '                Range("AssumptionsA" & TestNumber).Copy
    '                    Range("LiveAssumptionsA").PasteSpecial xlValues
    '                Range("NewStressName").Copy




    '                    Range("StressLiveInfo").Copy
    '                    Range("S" & TestNumber & "Data").PasteSpecial xlValues


    '                    Range("StressTestMode") = "N"
    '                    ActiveWorkbook.Names.Add Name:="Mode", RefersToR1C1:="=""Business Plan"""


    '                Else  ' no stress test name

    '                    Range("S" & TestNumber & "Data") = 0

    '                End If

    '            End If

    '        Next TestNumber

    '        Range("LiveAssumptions").ClearContents
    '        Range("LiveAssumptionsA").ClearContents

    '        Call fncPrtSht("Multivariable Planner")


    '        Call fncPrtSht("OW - Captured Data")


    '        Call fncPrtSht("Live Multivariable Planner")


    '        Sheets("Multivariable Dashboard").Activate



    '        MsgBox("Data captured")


    '    End Sub
    '    Sub Import_Base()


    '        Dim Scenario As String

    '        Scenario = "S0"

    '        Call Import_Separate_Plans(Scenario)



    '    End Sub
    '    Sub Import_S1()


    '        Dim Scenario As String

    '        Scenario = "S1"

    '        Call Import_Separate_Plans(Scenario)


    '    End Sub
    '    Sub Import_S2()


    '        Dim Scenario As String

    '        Scenario = "S2"
    '        Call Import_Separate_Plans(Scenario)



    '    End Sub
    '    Sub Import_S3()



    '        Dim Scenario As String

    '        Scenario = "S3"
    '        Call Import_Separate_Plans(Scenario)


    '    End Sub
    '    Sub Import_S4()


    '        Dim Scenario As String

    '        Scenario = "S4"
    '        Call Import_Separate_Plans(Scenario)


    '    End Sub
    '    Sub Import_S5()



    '        Dim Scenario As String

    '        Scenario = "S5"
    '        Call Import_Separate_Plans(Scenario)


    '    End Sub
    '    Sub Import_S6()



    '        Dim Scenario As String

    '        Scenario = "S6"
    '        Call Import_Separate_Plans(Scenario)


    '    End Sub
    '    Sub Import_S7()



    '        Dim Scenario As String

    '        Scenario = "S7"
    '        Call Import_Separate_Plans(Scenario)


    '    End Sub
    '    Sub Import_S8()



    '        Dim Scenario As String

    '        Scenario = "S8"
    '        Call Import_Separate_Plans(Scenario)


    '    End Sub
    '    Sub Import_S9()


    '        Dim Scenario As String

    '        Scenario = "S9"
    '        Call Import_Separate_Plans(Scenario)

    '    End Sub
    '    Sub Import_S10()


    '        Dim Scenario As String

    '        Scenario = "S10"
    '        Call Import_Separate_Plans(Scenario)

    '    End Sub
    '    Sub Import_Separate_Plans(Scenario As String)

    '        ' import data from the Detailed Sensitivities sheet of a business plan model
    '        Dim BPFile, FileToOpen
    '        Dim ScenarioName As String
    '        Dim ScenarioNumber As Integer



    '        ScenarioNumber = Right(Scenario, 2)

    '        If Scenario <> "S0" Then

    '            ScenarioName = Range("R_StressTestNames").Cells(ScenarioNumber + 1)

    '        Else

    '            ScenarioName = "Base Case"
    '        End If

    '        ' Open and assign Business Plan File
    '        FileToOpen = Application.GetOpenFilename(fileFilter:="Excel Files (*.xls*), *.xls*", Title:="Select BP file For " & ScenarioName)

    '        If FileToOpen <> False Then

    '            ' Set up BP file ready for data extraction


    '            Workbooks.Open(FileToOpen)
    '        Set BPFile = ActiveWorkbook

    '        ' Call procedures to clear old data and copy new data from the OW - Captured Data sheet.
    '        Call Clear_Data("S" & ScenarioNumber)
    '            ResumeCodeRun "Import_Separate_Plans"

    '        Call Copy_Scenario(BPFile, ScenarioNumber, ScenarioName)
    '            ResumeCodeRun "Import_Separate_Plans"

    '    Else

    '            MsgBox("Import cancelled")

    '        End If



    '    End Sub
    '    Sub Copy_Scenario(BPFile, ScenarioNumber As Integer, ScenarioName As String)

    '        'Codesafe JW 26/4/22

    '        ' Copy data from BP model's Detailed Sensitivities sheet to this model's OW - Captured Data sheet
    '        ' called by Import_Separate_Plans
    '        Dim Covenant As String
    '        Dim i As Integer, j As Integer, k As Integer
    '        Dim DSOffset As Integer
    '        Dim IDMultiplier As Integer
    '        Dim IncludeDev As String
    '        Dim InclMRBP As String, InclOSBP As String, InclComBP As String
    '        Dim InclMRImp As String, InclOSImp As String, InclComImp As String

    '        On Error Resume Next




    '        ' Assign multipliers used in creating offsets
    '        IDMultiplier = 6        ' for OW - Captured Data sheet

    '        ' modify BP formulae for build costs, sales and grant
    '        BPFile.Activate

    '        ' Copy data for "Base" position
    '        Application.StatusBar = "Importing data For " & ScenarioName

    '        ThisWorkbook.Activate
    '        Sheets("OW - Captured Data").Select

    '        BPFile.Activate
    '        Sheets("OW - Captured Data").Select


    '        Range("StressLiveInfo").Copy

    '        ThisWorkbook.Activate
    '        Range("R_DataOffset").Offset(1, 1 + ScenarioNumber * IDMultiplier).Select
    '        Selection.PasteSpecial xlValues

    '    BPFile.Activate



    '    End Sub
    '    Sub Modify_Formulae()

    '        On Error Resume Next



    '        Sheets("All Schemes Cashflow").Select
    '        ActiveSheet.Unprotect Password:=PWOld

    '    ' BUILD COSTS
    '        Range("AJ114").Formula = "= SUM('All Schemes BP Cashflow:All Schemes Imported Cashflow'!AJ114)*(1+$AJ$1)"
    '        Range("AJ114").Copy
    '        Range("AJ114:AK1048").Select
    '        Selection.PasteSpecial xlFormulas
    '    Range("BJ114:BJ1048").Select
    '        Selection.PasteSpecial xlFormulas
    '    Range("BL114:BL1048").Select
    '        Selection.PasteSpecial xlFormulas

    '    ' GRANT
    '        Range("AP114").Formula = "=SUM('All Schemes BP Cashflow:All Schemes Imported Cashflow'!AP114)*(1+$AJ$3)"
    '        Range("AP114").Copy
    '        Range("AP114:AP1048").Select
    '        Selection.PasteSpecial xlFormulas
    '    Range("BK114:BK1048").Select
    '        Selection.PasteSpecial xlFormulas
    '    Range("BM114:BM1048").Select
    '        Selection.PasteSpecial xlFormulas

    '    ' SALES
    '        Range("AL114").Formula = "=SUM('All Schemes BP Cashflow:All Schemes Imported Cashflow'!AL114)*(1+$AJ$2)"
    '        Range("AL114").Copy
    '        Range("AL114:AN1048").Select
    '        Selection.PasteSpecial xlFormulas

    '    Sheets("Existing Cashflows").Select
    '        ActiveSheet.Unprotect Password:=PWOld
    '    Range("J8").Formula = "=IF(Mode=""Valuation"",0,'RTB Receipts'!AU10-'RTB Receipts'!AV10)*(1+'All Schemes Cashflow'!$AJ$2)"
    '        Range("K8").Formula = "=IF(Mode=""Valuation"",0,'RTB Receipts'!AS10)*(1+'All Schemes Cashflow'!$AJ$2)"
    '        Range("L8").Formula = "=+'Other Disposal Income'!$D8*(1+'All Schemes Cashflow'!$AJ$2)"
    '        Range("J8:L8").Copy
    '        Range("J8:L47").Select
    '        Selection.PasteSpecial xlFormulas

    '    ' Additional income
    '        Sheets("Council, Other, Interco Income").Select
    '        ActiveSheet.Unprotect Password:=PWOld
    '    Range("M10").Formula = "=((SUMIF(OFFSET(Rep_Oinc_03c,0,2),R6C,Rep_Oinc_03c)+SUMIF(OFFSET(Rep_Oinc_03e,0,2),R6C&RC1,Rep_Oinc_03e)/1000+R1C)*'Target Rent Letting Numbers'!R[-2]C3)*'Other Income Factors'!R[-3]C10"
    '        Range("M10").Copy
    '        Range("M10:M49").Select
    '        Selection.PasteSpecial xlFormulas



    '    End Sub
    '    Sub Clear_Data(Scenario As String)



    '        ' Clear data from the OW - Captured Data sheet

    '        ThisWorkbook.Activate
    '        Sheets("OW - Captured Data").Select
    '        Range(Scenario & "Data").ClearContents
    '        Range("OWCapturedData").Select




    '    End Sub



    '    Sub Clear_Stress_Sensitivity_Capture()

    '        'CdoeSafe JW 26/4/22

    '        'Called by Live Multivariable Planner - Button 3

    '        ' clear all details



    '        Dim Answer

    '        Answer = MsgBox(Prompt:="Are you sure that you want" & Chr(13) & "to delete all captured data?", Buttons:=vbYesNo + vbQuestion)

    '        If Answer = vbYes Then


    '            With Sheets("Stress Sensitivity List")

    '                .Activate
    '                .Unprotect Password:=PW

    '        End With

    '            Range("StressSensitivityData").EntireRow.ClearContents

    '            Call fncPrtSht("Stress Sensitivity List")


    '        End If




    '    End Sub
    '    Sub Delete_Stress_Sensitivity_Capture()

    '        'CodeSafe JW 26/4/22

    '        'Called by Stress Sensitivity List - Button 2

    '        ' clear one row of details
    '        Dim y As Integer, LastRowNum As Integer
    '        Dim FirstRowNum As Integer
    '        Dim ChosenCell As Range, DefaultAddr As String
    '        Dim CalcMethod As Integer
    '        Dim ResponseMessage, Answer, Title As String
    '        Title = "Delete selected row"



    '        '   Assign row to be deleted
    '        DefaultAddr = ActiveCell.Address

    '    Set ChosenCell = Application.InputBox(Prompt:="Please select one cell in the row you wish to delete", Title:=Title, Default:=DefaultAddr, Type:=8)                    '   Address type

    '    y = ChosenCell.Row

    '        If y <> 0 Then



    '            With Sheets("Stress Sensitivity List")

    '                .Activate
    '                .Unprotect Password:=PW

    '        End With

    '            FirstRowNum = Range("StressSensitivityData").Cells(1).Row
    '            LastRowNum = Range("StressSensitivityData").Cells(Range("StressSensitivityData").Rows.Count).Row

    '            Rows(y).Select

    '            If (y >= FirstRowNum And y < LastRowNum) Then

    '                Answer = MsgBox(Prompt:="Are you sure that you want" & Chr(13) & "to delete Row " & y & "?", Buttons:=vbYesNo + vbQuestion)

    '                If Answer = vbYes Then

    '                    If LastRowNum - FirstRowNum = 1 Then    ' are there only two rows left

    '                        Rows(y).EntireRow.ClearContents

    '                    Else

    '                        Rows(y).Delete

    '                    End If

    '                End If

    '            Else

    '                MsgBox("This row cannot be deleted")

    '            End If

    '            Call fncPrtSht("Stress Sensitivity List")



    '        End If

    'Cancelled:



    '    End Sub

    Private Class CustomFormPainterST

        Inherits FormPainter
        Public Sub New(ByVal owner As System.Windows.Forms.Control, ByVal provider As DevExpress.Skins.ISkinProvider)

            MyBase.New(owner, provider)

        End Sub
        Private Function GetFormBorderColor() As Color

            Dim formBorderColor = (TryCast(Owner, StressTest)).FormBorderColor
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

    Private Sub StressTest_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        e.Cancel = True
        SetDeactivated()
        Me.Hide()

    End Sub

    Private Sub TextEditMultivariableName_EditValueChanged(sender As Object, e As EventArgs)

        ActiveWorkbook.DefinedNames.GetDefinedName("NewStressName").Range(0, 0).SetValueFromText(TextEditMultivariableName.EditValue)

    End Sub

    Private Sub ToggleModeSwitch_Toggled(sender As Object, e As EventArgs) Handles ToggleModeSwitch.Toggled

        If STMode = "N" Then

            If ToggleModeSwitch.IsOn = True Then

                STMode = "Y"
                StressTestModeSwitch(True)

            End If

        ElseIf STMode = "Y" Then

            If ToggleModeSwitch.IsOn = False Then

                STMode = "N"
                StressTestModeSwitch(True)

            End If

        End If

    End Sub
End Class

