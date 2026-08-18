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

    Private Const BreachOutputFirstColumnIndex As Integer = 26 'AA
    Private Const BreachOutputColumnCount As Integer = 6       'AA:AF

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

    Private NativePlannerScenario As DevExpress.XtraEditors.ComboBoxEdit
    Private NativePlannerName As DevExpress.XtraEditors.TextEdit
    Private NativePlannerImportMode As DevExpress.XtraEditors.ComboBoxEdit
    Private NativePlannerInclude As DevExpress.XtraEditors.CheckEdit
    Private NativePlannerGrid As GridControl
    Private NativePlannerView As GridView
    Private NativePlannerData As System.Data.DataTable
    Private NativeYearEditor As RepositoryItemComboBox
    Private NativeDashboardScenario As DevExpress.XtraEditors.ComboBoxEdit
    Private NativeDashboardCharts As TableLayoutPanel
    Private NativeDashboardSummary As GridControl
    Private NativeSensitivityGrid As GridControl
    Private NativeSensitivityView As GridView
    Private CovenantSummaryPanel As TableLayoutPanel
    Private NativeComparativeSelectors As New List(Of DevExpress.XtraEditors.ComboBoxEdit)
    Private NativeComparativeChartsA As TableLayoutPanel
    Private NativeComparativeChartsB As TableLayoutPanel
    Private NativeComparativeSummaryA As GridControl
    Private NativeComparativeSummaryB As GridControl
    Private LoadingNativeViews As Boolean

    Private ModelID As Integer
    Private ExportPackageCount As Integer = 0
    Private ExportPackagesIndex As Integer = -1
    Private ExportPackages(-1) As GridExportPackage

    Private PresentedDS As Abovo.DataObject.DataCellRange
    Private ScaleUnits As Integer
    Private FirstTabReferenceSize As Size = Size.Empty
    Private FirstTabBaseFontSize As Single
    Private UpdatingCovenantSelection As Boolean
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

        InitialiseNativeStressViews()

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

        ConfigureResponsiveFirstTab()
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
                RenderStressHeaderHTMLData()

            Case "Comp1"

                XtraTabControlStressTest.SelectedTabPage = XtraTabPageCompA
                RefreshNativeComparativeViews()

            Case "Comp2"

                XtraTabControlStressTest.SelectedTabPage = XtraTabPageCompB
                RefreshNativeComparativeViews()

            Case "MVPlan"

                XtraTabControlStressTest.SelectedTabPage = XtraTabPageMVP
                RefreshNativePlanner()

            Case "MVDash"

                XtraTabControlStressTest.SelectedTabPage = XtraTabPageDashboard
                RefreshNativeDashboard()

        End Select

    End Sub

    Public Sub SetActive()

        If XtraTabControlStressTest.SelectedTabPage Is XtraTabPageMVP Then
            RefreshNativePlanner()
        ElseIf XtraTabControlStressTest.SelectedTabPage Is XtraTabPageDashboard Then
            RefreshNativeDashboard()
        ElseIf XtraTabControlStressTest.SelectedTabPage Is XtraTabPageCompA Then
            RefreshNativeComparativeViews()
        ElseIf XtraTabControlStressTest.SelectedTabPage Is XtraTabPageCompB Then
            RefreshNativeComparativeViews()
        End If

    End Sub

    Public Sub SetDeactivated()

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
        'AA:AF may be hidden while Business Plan mode is active.  The interface
        'still needs all six fields available when Multivariable mode is enabled.
        RDSOptions.SkipHiddenColumns = False

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
        ModelPostingComboBoxSelectCovenant.Properties.TextEditStyle =
            DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        ModelPostingComboBoxSelectCovenant.ProcesDefValue()
        AddHandler ModelPostingComboBoxSelectCovenant.EditValueChanged,
            AddressOf CovenantSelectionChanged

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

        If STMode = "Y" Then UnhideColumnsCommand()
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
            ExcelModels(ModelID).WB.Worksheets("Live Multivariable Planner").Columns.Unhide(
                BreachOutputFirstColumnIndex, BreachOutputColumnCount)
        Catch ex As Exception

        End Try

    End Sub
    Sub CovGraphsDataRange_DataSourceChanged(sender As Object, e As EventArgs)

        BuildCovCharts()

    End Sub


    Sub ProcessBreachesGrid(OnOff As Boolean)

        If OnOff = False Then

            GridControlBreaches.DataSource = Nothing
            GridControlBreaches.Enabled = False
            GridControlBreaches.Visible = False

            Exit Sub

        Else

            GridControlBreaches.Enabled = True
            GridControlBreaches.Visible = True
            GridControlBreaches.DataSource = BuildBreachOutputTable()

            Dim GV As GridView = GridControlBreaches.MainView
            GridControlBreaches.ForceInitialize()
            GV.PopulateColumns()
            If GV.Columns.Count < 6 Then
                GridControlBreaches.RefreshDataSource()
                Return
            End If
            GV.OptionsView.ShowGroupPanel = False

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
            GV.Columns(3).AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            GV.Columns(3).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            GV.Columns(4).AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            GV.Columns(4).AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far

            GV.Columns(5).AppearanceCell.ForeColor = Color.Red

            GV.Columns(1).AppearanceCell.ForeColor = Color.Red
            GV.Columns(1).AppearanceCell.Font = New System.Drawing.Font("Wingdings", 13, FontStyle.Regular)
            GV.OptionsView.ColumnAutoWidth = True
            GV.BestFitColumns()

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
                ActiveWorkbook.Worksheets("Live Multivariable Planner").Columns.Hide(
                    BreachOutputFirstColumnIndex, BreachOutputColumnCount)
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
            CovChart.Titles(0).Visibility = DevExpress.Utils.DefaultBoolean.True
            CovChart.Titles(0).Dock = ChartTitleDockStyle.Top
            CovChart.Titles(0).Alignment = StringAlignment.Center
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
        ProcessBreachesGrid(STMode = "Y")
        RefreshCovenantSummary()
        BuildCovCharts()

        Me.Cursor = Cursors.Default

    End Sub

    Sub StressTestModeAdjustments()

        UnhideColumnsCommand()
        UNProtectWS(ModelID, "OW - Live Stress Reporting")

        Dim SourceRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.Range("StressSwitch")
        Dim DestRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.Range("StressBase")
        DestRange.CopyFrom(SourceRange, PasteSpecial.Values)

        ProtectWS(ModelID, "OW - Live Stress Reporting")


    End Sub

    Sub Stress_Sensitivity_Capture()

        UNProtectWS(ModelID, "Stress Sensitivity List")

        Dim WB As IWorkbook = ExcelModels(ModelID).WB
        Try
            WB.DefinedNames.GetDefinedName("StressSensitivityDate").Range(0, 0).Value = Now()

            Dim SourceRow As DevExpress.Spreadsheet.CellRange =
                WB.DefinedNames.GetDefinedName("StressSensitivity").Range
            Dim WSTarget As DevExpress.Spreadsheet.Worksheet =
                WB.Worksheets("Stress Sensitivity List")
            Dim BottomRow As DevExpress.Spreadsheet.CellRange = Nothing

            For Attempt As Integer = 0 To 1
                Dim DataRows As DevExpress.Spreadsheet.CellRange =
                    WB.DefinedNames.GetDefinedName("StressSensitivityData").Range
                Dim CaptureArea As DevExpress.Spreadsheet.CellRange =
                    WSTarget.Range.FromLTRB(
                        SourceRow.LeftColumnIndex,
                        DataRows.TopRowIndex,
                        SourceRow.RightColumnIndex,
                        DataRows.BottomRowIndex)

                For RowIndex As Integer = CaptureArea.TopRowIndex To CaptureArea.BottomRowIndex
                    Dim CandidateRow As DevExpress.Spreadsheet.CellRange =
                        WSTarget.Range.FromLTRB(
                            CaptureArea.LeftColumnIndex,
                            RowIndex,
                            CaptureArea.RightColumnIndex,
                            RowIndex)

                    If CandidateRow.ToArray().All(Function(Cell) Cell.Value.IsEmpty) Then
                        BottomRow = CandidateRow
                        Exit For
                    End If
                Next

                If BottomRow IsNot Nothing Then
                    Exit For
                End If

                InsertRows(ModelID, "StressSensitivityData", 1, True)
            Next

            If BottomRow Is Nothing Then
                Throw New InvalidOperationException("A new stress sensitivity row could not be created.")
            End If

            BottomRow.CopyFrom(SourceRow, PasteSpecial.Values)

            'The worksheet source formula deliberately reports "Base" while the
            'model is in Business Plan mode.  A quick capture made through this
            'interface should retain the description the user entered instead.
            Dim CaptureName As String = TextEditMultivariableName.Text.Trim()
            If Not String.IsNullOrWhiteSpace(CaptureName) AndAlso
               Not String.Equals(CaptureName, "Base", StringComparison.OrdinalIgnoreCase) Then
                BottomRow(0, 0).Value = CellValue.FromObject(CaptureName)
            End If

            WB.Calculate()
        Finally
            ProtectWS(ModelID, "Stress Sensitivity List")
        End Try

    End Sub


    Sub DeStressAdjustments()

        ExcelModels(ModelID).WB.Worksheets("Live Multivariable Planner").Columns.Hide(
            BreachOutputFirstColumnIndex, BreachOutputColumnCount)
        ProcessBreachesGrid(False)

    End Sub

    Private Function BuildBreachOutputTable() As System.Data.DataTable

        Dim Data As New System.Data.DataTable
        For Each ColumnName As String In {"Year", "Mvt", "Base", "Live", "Target", "Met/Breach"}
            Data.Columns.Add(ColumnName, GetType(String))
        Next

        Dim Sheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Live Multivariable Planner")
        Dim SourceRange As DevExpress.Spreadsheet.CellRange =
            Sheet.Range(ExcelModels(ModelID).WBStructure.StressTestDefinition.StresstestBreachOutputRange)

        For RowIndex As Integer = SourceRange.TopRowIndex To SourceRange.BottomRowIndex
            Dim Row As System.Data.DataRow = Data.NewRow()
            For ColumnOffset As Integer = 0 To 5
                Row(ColumnOffset) =
                    Sheet.Cells(RowIndex, SourceRange.LeftColumnIndex + ColumnOffset).DisplayText
            Next
            Data.Rows.Add(Row)
        Next

        Return Data

    End Function

    Private Sub ConfigureResponsiveFirstTab()

        TablePanelStressInputs.AutoSize = False
        TablePanelStressInputs.Dock = DockStyle.None
        TablePanelStressInputs.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom
        TablePanelStressInputs.UseSkinIndents = False
        TablePanelStressInputs.Padding = New Padding(8)
        TablePanelStressInputs.Columns.Clear()
        TablePanelStressInputs.Columns.AddRange(
            New DevExpress.Utils.Layout.TablePanelColumn() {
                New DevExpress.Utils.Layout.TablePanelColumn(
                    DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 18.0!),
                New DevExpress.Utils.Layout.TablePanelColumn(
                    DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 27.0!),
                New DevExpress.Utils.Layout.TablePanelColumn(
                    DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 34.0!),
                New DevExpress.Utils.Layout.TablePanelColumn(
                    DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 21.0!)
            })
        TablePanelStressInputs.Rows.Clear()
        TablePanelStressInputs.Rows.AddRange(
            New DevExpress.Utils.Layout.TablePanelRow() {
                New DevExpress.Utils.Layout.TablePanelRow(
                    DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 116.0!),
                New DevExpress.Utils.Layout.TablePanelRow(
                    DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 100.0!)
            })

        ConfigureModePanel()
        ConfigureCapturePanel()
        ConfigureQuickCapturePanel()
        ConfigureCovenantPanel()
        ConfigureCovenantOutputRows()
        AddHandler XtraTabPageLMVP.Resize, AddressOf FirstTabPageResized

        PanelControlCovSel.Visible = STMode = "Y"
        GridControlBreaches.Visible = STMode = "Y"

    End Sub

    Private Function NewFirstTabLayout(ColumnCount As Integer) As TableLayoutPanel

        Dim Layout As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .ColumnCount = ColumnCount,
            .Margin = New Padding(0),
            .Padding = New Padding(10, 7, 10, 7)
        }
        Return Layout

    End Function

    Private Sub ConfigureModePanel()

        PanelControl1.Controls.Clear()
        Dim Layout As TableLayoutPanel = NewFirstTabLayout(1)
        Layout.RowCount = 2
        Layout.RowStyles.Add(New RowStyle(SizeType.Percent, 45.0!))
        Layout.RowStyles.Add(New RowStyle(SizeType.Percent, 55.0!))
        LabelControl3.Dock = DockStyle.Fill
        LabelControl3.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        LabelControl3.Appearance.TextOptions.VAlignment = VertAlignment.Center
        ToggleModeSwitch.Dock = DockStyle.Left
        Layout.Controls.Add(LabelControl3, 0, 0)
        Layout.Controls.Add(ToggleModeSwitch, 0, 1)
        PanelControl1.Controls.Add(Layout)

    End Sub

    Private Sub ConfigureCapturePanel()

        PanelControl3.Controls.Clear()
        Dim Layout As TableLayoutPanel = NewFirstTabLayout(2)
        Layout.RowCount = 1
        Layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 42.0!))
        Layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 58.0!))
        SimpleButtonCapture.Text = "Capture scenario"
        SimpleButtonCapture.Dock = DockStyle.None
        SimpleButtonCapture.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        SimpleButtonCapture.Height = 42
        SimpleButtonCapture.Margin = New Padding(4, 12, 10, 12)

        Dim SelectionLayout As New TableLayoutPanel With {
            .Dock = DockStyle.Fill, .BackColor = Color.White,
            .ColumnCount = 1, .RowCount = 2, .Margin = New Padding(0)
        }
        SelectionLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 32.0!))
        SelectionLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        LabelControl1.Dock = DockStyle.Fill
        LabelControl1.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        LabelControl1.Appearance.TextOptions.VAlignment = VertAlignment.Center
        ComboBoxBreachMode.Dock = DockStyle.Fill
        ComboBoxBreachMode.Margin = New Padding(0, 0, 0, 4)
        SelectionLayout.Controls.Add(LabelControl1, 0, 0)
        SelectionLayout.Controls.Add(ComboBoxBreachMode, 0, 1)

        Layout.Controls.Add(SimpleButtonCapture, 0, 0)
        Layout.Controls.Add(SelectionLayout, 1, 0)
        PanelControl3.Controls.Add(Layout)

    End Sub

    Private Sub ConfigureQuickCapturePanel()

        PanelControl2.Controls.Clear()
        Dim Layout As TableLayoutPanel = NewFirstTabLayout(3)
        Layout.RowCount = 1
        Layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 34.0!))
        Layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 15.0!))
        Layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 51.0!))
        SimpleButtonQC.Dock = DockStyle.Fill
        SimpleButtonQC.Margin = New Padding(0, 2, 10, 2)
        LabelControl4.Text = "Capture as:"
        LabelControl4.Dock = DockStyle.Fill
        LabelControl4.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        LabelControl4.Appearance.TextOptions.HAlignment = HorzAlignment.Far
        LabelControl4.Appearance.TextOptions.VAlignment = VertAlignment.Center
        LabelControl4.Margin = New Padding(0, 0, 8, 0)
        TextEditMultivariableName.Properties.UseAdvancedMode = DefaultBoolean.False
        TextEditMultivariableName.Properties.NullValuePrompt = "Scenario name"
        TextEditMultivariableName.Dock = DockStyle.Fill
        TextEditMultivariableName.Margin = New Padding(0, 10, 0, 10)
        Layout.Controls.Add(SimpleButtonQC, 0, 0)
        Layout.Controls.Add(LabelControl4, 1, 0)
        Layout.Controls.Add(TextEditMultivariableName, 2, 0)
        PanelControl2.Controls.Add(Layout)

    End Sub

    Private Sub ConfigureCovenantPanel()

        PanelControlCovSel.Controls.Clear()
        Dim Layout As TableLayoutPanel = NewFirstTabLayout(1)
        Layout.RowCount = 2
        Layout.RowStyles.Add(New RowStyle(SizeType.Percent, 45.0!))
        Layout.RowStyles.Add(New RowStyle(SizeType.Percent, 55.0!))
        LabelControl2.Text = "Selected covenant"
        LabelControl2.Dock = DockStyle.Fill
        LabelControl2.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        LabelControl2.Appearance.TextOptions.VAlignment = VertAlignment.Center
        ModelPostingComboBoxSelectCovenant.Dock = DockStyle.Fill
        Layout.Controls.Add(LabelControl2, 0, 0)
        Layout.Controls.Add(ModelPostingComboBoxSelectCovenant, 0, 1)
        PanelControlCovSel.Controls.Add(Layout)

    End Sub

    Private Sub ConfigureCovenantOutputRows()

        TablePanelOutputs.UseSkinIndents = False
        TablePanelOutputs.Padding = New Padding(6)
        TablePanelOutputs.Rows.Clear()
        TablePanelOutputs.Rows.AddRange(
            New DevExpress.Utils.Layout.TablePanelRow() {
                New DevExpress.Utils.Layout.TablePanelRow(
                    DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 198.0!),
                New DevExpress.Utils.Layout.TablePanelRow(
                    DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 1.0!),
                New DevExpress.Utils.Layout.TablePanelRow(
                    DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 1.0!),
                New DevExpress.Utils.Layout.TablePanelRow(
                    DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 1.0!),
                New DevExpress.Utils.Layout.TablePanelRow(
                    DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 1.0!),
                New DevExpress.Utils.Layout.TablePanelRow(
                    DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 1.0!)
            })
        GridControlTextOut.Visible = False

        If CovenantSummaryPanel Is Nothing Then
            CovenantSummaryPanel = New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .BackColor = Color.White,
                .ColumnCount = 2,
                .RowCount = 1,
                .Margin = New Padding(0),
                .Padding = New Padding(0)
            }
            CovenantSummaryPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
            CovenantSummaryPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
            TablePanelOutputs.Controls.Add(CovenantSummaryPanel)
        End If

        TablePanelOutputs.SetColumn(CovenantSummaryPanel, 0)
        TablePanelOutputs.SetRow(CovenantSummaryPanel, 0)
        CovenantSummaryPanel.BringToFront()
        RefreshCovenantSummary()
        TablePanelOutputs.SetRow(CvntChart1, 1)
        TablePanelOutputs.SetRow(CvntChart2, 2)
        TablePanelOutputs.SetRow(CvntChart3, 3)
        TablePanelOutputs.SetRow(CvntChart4, 4)
        TablePanelOutputs.SetRow(CvntChart5, 5)
        For Each Chart As ChartControl In {
                CvntChart1, CvntChart2, CvntChart3, CvntChart4, CvntChart5}
            Chart.Dock = DockStyle.Fill
            Chart.Margin = New Padding(4, 5, 4, 5)
        Next

    End Sub

    Private Sub RefreshCovenantSummary()

        If CovenantSummaryPanel Is Nothing Then Return

        Dim Sheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Live Multivariable Planner")

        CovenantSummaryPanel.SuspendLayout()
        Try
            CovenantSummaryPanel.Controls.Clear()
            CovenantSummaryPanel.Controls.Add(
                CreateCovenantSummaryBlock(Sheet, 16, 18, 19), 0, 0) 'Q, S, T
            CovenantSummaryPanel.Controls.Add(
                CreateCovenantSummaryBlock(Sheet, 21, 23, 24), 1, 0) 'V, X, Y
        Finally
            CovenantSummaryPanel.ResumeLayout(True)
        End Try

    End Sub

    Private Function CreateCovenantSummaryBlock(
        Sheet As DevExpress.Spreadsheet.Worksheet,
        LabelColumn As Integer,
        TargetColumn As Integer,
        CurrentColumn As Integer) As TableLayoutPanel

        Dim Block As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .ColumnCount = 3,
            .RowCount = 7,
            .Margin = New Padding(4, 2, 4, 4),
            .Padding = New Padding(4, 0, 4, 0)
        }
        Block.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 56.0!))
        Block.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 22.0!))
        Block.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 22.0!))
        Block.RowStyles.Add(New RowStyle(SizeType.Absolute, 28.0!))
        For RowIndex As Integer = 1 To 6
            Block.RowStyles.Add(New RowStyle(SizeType.Percent, CSng(100.0 / 6.0)))
        Next

        Block.Controls.Add(CreateCovenantSummaryLabel("", Nothing, HorzAlignment.Near, True), 0, 0)
        Block.Controls.Add(CreateCovenantSummaryLabel("Target", Nothing, HorzAlignment.Far, True), 1, 0)
        Block.Controls.Add(CreateCovenantSummaryLabel("Current", Nothing, HorzAlignment.Far, True), 2, 0)

        For RowOffset As Integer = 0 To 5
            Dim WorksheetRow As Integer = 7 + RowOffset
            Block.Controls.Add(
                CreateCovenantSummaryLabel(
                    Sheet.Cells(WorksheetRow, LabelColumn).DisplayText,
                    Sheet.Cells(WorksheetRow, LabelColumn), HorzAlignment.Near, False),
                0, RowOffset + 1)
            Block.Controls.Add(
                CreateCovenantSummaryLabel(
                    Sheet.Cells(WorksheetRow, TargetColumn).DisplayText,
                    Sheet.Cells(WorksheetRow, TargetColumn), HorzAlignment.Far, False),
                1, RowOffset + 1)
            Block.Controls.Add(
                CreateCovenantSummaryLabel(
                    Sheet.Cells(WorksheetRow, CurrentColumn).DisplayText,
                    Sheet.Cells(WorksheetRow, CurrentColumn), HorzAlignment.Far, False),
                2, RowOffset + 1)
        Next

        Return Block

    End Function

    Private Function CreateCovenantSummaryLabel(
        Text As String,
        SourceCell As DevExpress.Spreadsheet.Cell,
        Alignment As HorzAlignment,
        IsHeader As Boolean) As DevExpress.XtraEditors.LabelControl

        Dim Label As New DevExpress.XtraEditors.LabelControl With {
            .Text = Text,
            .Dock = DockStyle.Fill,
            .AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None,
            .Margin = New Padding(2, 0, 2, 0),
            .Padding = New Padding(5, 0, 5, 0)
        }
        Label.Appearance.TextOptions.HAlignment = Alignment
        Label.Appearance.TextOptions.VAlignment = VertAlignment.Center

        If SourceCell IsNot Nothing Then
            ApplyWorkbookCellAppearance(Label.Appearance, SourceCell)
        Else
            Label.Appearance.BackColor = Color.WhiteSmoke
            Label.Appearance.ForeColor = AbovoBlue
            Label.Appearance.Options.UseBackColor = True
            Label.Appearance.Options.UseForeColor = True
        End If

        If IsHeader OrElse (SourceCell IsNot Nothing AndAlso SourceCell.Font.Bold) Then
            Label.Appearance.Font = New Font(Label.Font, FontStyle.Bold)
            Label.Appearance.Options.UseFont = True
        End If

        Return Label

    End Function

    Private Sub CovenantSelectionChanged(sender As Object, e As EventArgs)

        If UpdatingCovenantSelection Then Return
        Dim SelectedCovenant As String =
            Convert.ToString(ModelPostingComboBoxSelectCovenant.EditValue).Trim()
        If String.IsNullOrWhiteSpace(SelectedCovenant) Then Return

        UpdatingCovenantSelection = True
        Me.Cursor = Cursors.WaitCursor
        Try
            ActiveWorkbook.Worksheets("Live Multivariable Planner").
                Range("AD3")(0, 0).Value = CellValue.FromObject(SelectedCovenant)
            ActiveWorkbook.Calculate()
            RefreshCovenantSummary()
            BuildCovCharts()
            ProcessBreachesGrid(STMode = "Y")
        Finally
            Me.Cursor = Cursors.Default
            UpdatingCovenantSelection = False
        End Try

    End Sub

    Private Sub SimpleButtonModeSwitch_Click(sender As Object, e As EventArgs)


        ToggleStressTestMode()

    End Sub

    Private Sub StressTest_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        FirstTabReferenceSize = Me.ClientSize
        FirstTabBaseFontSize = Me.Font.Size
        ApplyResponsiveFirstTabScale()

    End Sub

    Private Sub FirstTabPageResized(sender As Object, e As EventArgs)

        ApplyResponsiveFirstTabScale()

    End Sub

    Private Sub StressTest_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize

        ApplyResponsiveFirstTabScale()

    End Sub

    Private Sub ApplyResponsiveFirstTabScale()

        If FirstTabBaseFontSize <= 0 OrElse
           Me.WindowState = FormWindowState.Minimized Then Return

        Dim WorkspaceWidth As Integer = Math.Min(2200, XtraTabPageLMVP.ClientSize.Width)
        Dim WorkspaceLeft As Integer =
            Math.Max(0, (XtraTabPageLMVP.ClientSize.Width - WorkspaceWidth) \ 2)
        TablePanelStressInputs.Bounds =
            New Rectangle(WorkspaceLeft, 0, WorkspaceWidth, XtraTabPageLMVP.ClientSize.Height)

        Dim WidthScale As Double = WorkspaceWidth / 1920.0
        Dim HeightScale As Double = XtraTabPageLMVP.ClientSize.Height / 980.0
        Dim Scale As Double = Math.Max(0.88, Math.Min(1.25, Math.Min(WidthScale, HeightScale)))
        Dim FontSize As Single = CSng(Math.Max(9.5, FirstTabBaseFontSize * Scale))
        ApplyControlFont(XtraTabPageLMVP, FontSize)
        Dim HeaderHeight As Single = CSng(Math.Max(112.0, 120.0 * Scale))
        TablePanelStressInputs.Rows(0).Height = HeaderHeight

        Dim OutputHeight As Double = Math.Min(
            1180.0, Math.Max(620.0, XtraTabPageLMVP.ClientSize.Height - HeaderHeight - 12.0))
        TablePanelOutputs.Rows(0).Height = CSng(Math.Max(190.0, OutputHeight * 0.22))
        For RowIndex As Integer = 1 To 5
            TablePanelOutputs.Rows(RowIndex).Height = 1.0!
        Next

        For Each Grid As GridControl In FindControls(Of GridControl)(XtraTabPageLMVP)
            Dim View As GridView = TryCast(Grid.MainView, GridView)
            If View Is Nothing Then Continue For
            View.Appearance.Row.Font = New Font(View.Appearance.Row.Font.FontFamily, FontSize)
            View.Appearance.HeaderPanel.Font =
                New Font(View.Appearance.HeaderPanel.Font.FontFamily, FontSize, FontStyle.Bold)
            View.RowHeight = CInt(Math.Max(22, 27 * Scale))
            View.ColumnPanelRowHeight = CInt(Math.Max(28, 34 * Scale))
        Next

        If GridView2.Columns.Count > 1 Then
            GridView2.Columns(1).AppearanceCell.Font =
                New Font("Wingdings", CSng(FontSize * 1.15), FontStyle.Regular)
        End If
        For Each Chart As ChartControl In {
                CvntChart1, CvntChart2, CvntChart3, CvntChart4, CvntChart5}
            If Chart.Titles.Count > 0 Then
                Chart.Titles(0).DXFont =
                    New DXFont("Tahoma", Math.Min(FontSize, 9.5F), DXFontStyle.Bold)
            End If
        Next

    End Sub

    Private Sub ApplyControlFont(Parent As Control, FontSize As Single)

        For Each Child As Control In Parent.Controls
            Child.Font = New Font(Child.Font.FontFamily, FontSize, Child.Font.Style)
            If Child.HasChildren Then ApplyControlFont(Child, FontSize)
        Next

    End Sub

    Private Iterator Function FindControls(Of T As Control)(
        Parent As Control) As IEnumerable(Of T)

        For Each Child As Control In Parent.Controls
            If TypeOf Child Is T Then Yield DirectCast(Child, T)
            If Child.HasChildren Then
                For Each Descendant As T In FindControls(Of T)(Child)
                    Yield Descendant
                Next
            End If
        Next

    End Function

    Private Sub SimpleButtonQC_Click(sender As Object, e As EventArgs) Handles SimpleButtonQC.Click

        RunStressSensitivityCapture()

    End Sub

    Private Sub RunStressSensitivityCapture()

        Me.Cursor = Cursors.WaitCursor
        Try
            Stress_Sensitivity_Capture()
            RenderStressHeaderHTMLData()
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub SimpleButtonExtraQC_Click(sender As Object, e As EventArgs) Handles SimpleButtonExtraQC.Click

        RunStressSensitivityCapture()

    End Sub

    Private Sub SimpleButtonClearAllQCData_Click(sender As Object, e As EventArgs) Handles SimpleButtonClearAllQCData.Click

        If DevExpress.XtraEditors.XtraMessageBox.Show(
                "Clear every captured stress sensitivity record? This cannot be undone within this screen.",
                "Clear stress sensitivity captures",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning) <> DialogResult.Yes Then
            Return
        End If

        Me.Cursor = Cursors.WaitCursor
        UNProtectWS(ModelID, "Stress Sensitivity List")

        Try
            Dim SourceRow As DevExpress.Spreadsheet.CellRange =
                ActiveWorkbook.DefinedNames.GetDefinedName("StressSensitivity").Range
            Dim DataRows As DevExpress.Spreadsheet.CellRange =
                ActiveWorkbook.DefinedNames.GetDefinedName("StressSensitivityData").Range
            Dim Worksheet As DevExpress.Spreadsheet.Worksheet =
                ActiveWorkbook.Worksheets("Stress Sensitivity List")
            Dim CaptureArea As DevExpress.Spreadsheet.CellRange =
                Worksheet.Range.FromLTRB(
                    SourceRow.LeftColumnIndex,
                    DataRows.TopRowIndex,
                    SourceRow.RightColumnIndex,
                    DataRows.BottomRowIndex)

            CaptureArea.ClearContents()
            ActiveWorkbook.Calculate()
            RenderStressHeaderHTMLData()
        Finally
            ProtectWS(ModelID, "Stress Sensitivity List")
            Me.Cursor = Cursors.Default
        End Try

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
        RefreshNativeSensitivityList()



    End Sub

    Private Sub InitialiseNativeStressViews()

        BuildNativeSensitivityPage()
        BuildNativePlannerPage()
        BuildNativeDashboardPage()
        BuildNativeComparativePage(XtraTabPageCompA, False)
        BuildNativeComparativePage(XtraTabPageCompB, True)

        If ComboBoxBreachMode.SelectedIndex < 0 Then ComboBoxBreachMode.SelectedIndex = 0
        AddHandler SimpleButtonCapture.Click, AddressOf CaptureCurrentLiveScenario_Click
        AddHandler ComboBoxBreachMode.SelectedIndexChanged, AddressOf LiveScenarioNumberChanged

        RefreshNativePlanner()
        RefreshNativeSensitivityList()
        RefreshNativeDashboard()
        RefreshNativeComparativeViews()

    End Sub

    Private Sub BuildNativeSensitivityPage()

        WebBrowserStressCaptureOutput.Visible = False
        NativeSensitivityGrid = CreateReadOnlyGrid()
        NativeSensitivityView = CType(NativeSensitivityGrid.MainView, GridView)
        NativeSensitivityView.OptionsSelection.MultiSelect = True
        NativeSensitivityView.OptionsSelection.MultiSelectMode =
            DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.RowSelect
        AddHandler NativeSensitivityView.CustomDrawCell,
            AddressOf NativeSensitivityCustomDrawCell
        AddHandler NativeSensitivityView.CustomDrawColumnHeader,
            AddressOf NativeSensitivityCustomDrawColumnHeader
        TablePanelSSList.SetColumn(NativeSensitivityGrid, 0)
        TablePanelSSList.SetRow(NativeSensitivityGrid, 1)
        TablePanelSSList.Controls.Add(NativeSensitivityGrid)
        NativeSensitivityGrid.BringToFront()
        SimpleButtonClearQCSelectedData.Visible = True
        AddHandler SimpleButtonClearQCSelectedData.Click,
            AddressOf DeleteSelectedStressCaptures_Click

    End Sub

    Private Sub RefreshNativeSensitivityList()

        If NativeSensitivityGrid Is Nothing Then Return
        Dim Sheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Stress Sensitivity List")
        Dim DataRows As DevExpress.Spreadsheet.CellRange =
            ActiveWorkbook.DefinedNames.GetDefinedName("StressSensitivityData").Range
        Dim Data As New System.Data.DataTable
        Data.Columns.Add("SourceRow", GetType(Integer))
        Data.Columns.Add("Description", GetType(String))
        Data.Columns.Add("Peak Debt", GetType(String))
        Data.Columns.Add("Peak Debt Year", GetType(String))
        Data.Columns.Add("Repayment Year", GetType(String))
        For ColumnIndex As Integer = 4 To 8
            Dim Header As String = Sheet.Cells(4, ColumnIndex).DisplayText
            If String.IsNullOrWhiteSpace(Header) Then Header = "Breach " & (ColumnIndex - 3).ToString()
            Data.Columns.Add(Header.Replace(Environment.NewLine, " "), GetType(String))
        Next
        Data.Columns.Add("Captured", GetType(String))
        Data.Columns.Add("File", GetType(String))

        For RowIndex As Integer = DataRows.TopRowIndex To DataRows.BottomRowIndex
            If String.IsNullOrWhiteSpace(Sheet.Cells(RowIndex, 0).DisplayText) Then Continue For
            Dim Row As System.Data.DataRow = Data.NewRow()
            Row("SourceRow") = RowIndex
            Row("Description") = Sheet.Cells(RowIndex, 0).DisplayText
            Row("Peak Debt") = Sheet.Cells(RowIndex, 1).DisplayText
            Row("Peak Debt Year") = Sheet.Cells(RowIndex, 2).DisplayText
            Row("Repayment Year") = Sheet.Cells(RowIndex, 3).DisplayText
            For ColumnIndex As Integer = 4 To 8
                Row(ColumnIndex + 1) = Sheet.Cells(RowIndex, ColumnIndex).DisplayText
            Next
            Row("Captured") = Sheet.Cells(RowIndex, 59).DisplayText
            Row("File") = Sheet.Cells(RowIndex, 60).DisplayText
            Data.Rows.Add(Row)
        Next

        NativeSensitivityGrid.DataSource = Data
        If NativeSensitivityView.Columns("SourceRow") IsNot Nothing Then
            NativeSensitivityView.Columns("SourceRow").Visible = False
        End If
        SetSensitivitySourceColumn("Description", 0)
        SetSensitivitySourceColumn("Peak Debt", 1)
        SetSensitivitySourceColumn("Peak Debt Year", 2)
        SetSensitivitySourceColumn("Repayment Year", 3)
        For ColumnIndex As Integer = 4 To 8
            NativeSensitivityView.Columns(ColumnIndex + 1).Tag = ColumnIndex
        Next
        SetSensitivitySourceColumn("Captured", 59)
        SetSensitivitySourceColumn("File", 60)
        NativeSensitivityView.BestFitColumns()

    End Sub

    Private Sub SetSensitivitySourceColumn(FieldName As String, SourceColumn As Integer)

        If NativeSensitivityView.Columns(FieldName) IsNot Nothing Then
            NativeSensitivityView.Columns(FieldName).Tag = SourceColumn
        End If

    End Sub

    Private Sub NativeSensitivityCustomDrawCell(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs)

        If e.RowHandle < 0 OrElse e.Column Is Nothing OrElse
           Not TypeOf e.Column.Tag Is Integer Then Return

        Dim SourceRowValue As Object =
            NativeSensitivityView.GetRowCellValue(e.RowHandle, "SourceRow")
        If SourceRowValue Is Nothing OrElse SourceRowValue Is DBNull.Value Then Return

        Dim SourceCell As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.Worksheets("Stress Sensitivity List").Cells(
                Convert.ToInt32(SourceRowValue), CInt(e.Column.Tag))

        ApplyWorkbookCellAppearance(e.Appearance, SourceCell)

        If NativeSensitivityView.IsCellSelected(e.RowHandle, e.Column) Then
            e.Appearance.BackColor = Color.Beige
            e.Appearance.ForeColor = Color.Black
        End If

        e.DefaultDraw()
        e.Handled = True

    End Sub

    Private Sub NativeSensitivityCustomDrawColumnHeader(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Grid.ColumnHeaderCustomDrawEventArgs)

        If e.Column Is Nothing OrElse Not TypeOf e.Column.Tag Is Integer Then Return

        Dim HeaderCell As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.Worksheets("Stress Sensitivity List").Cells(4, CInt(e.Column.Tag))
        ApplyWorkbookCellAppearance(e.Appearance, HeaderCell)
        e.DefaultDraw()
        e.Handled = True

    End Sub

    Private Sub DeleteSelectedStressCaptures_Click(sender As Object, e As EventArgs)

        Dim SelectedHandles As Integer() = NativeSensitivityView.GetSelectedRows()
        If SelectedHandles.Length = 0 Then
            DevExpress.XtraEditors.XtraMessageBox.Show(
                "Select one or more captured scenarios to delete.", "Delete captures")
            Return
        End If
        If DevExpress.XtraEditors.XtraMessageBox.Show(
                "Delete the selected captured stress-test records?",
                "Delete captures", MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning) <> DialogResult.Yes Then Return

        UNProtectWS(ModelID, "Stress Sensitivity List")
        Try
            Dim Sheet As DevExpress.Spreadsheet.Worksheet =
                ActiveWorkbook.Worksheets("Stress Sensitivity List")
            For Each RowHandle As Integer In SelectedHandles
                Dim SourceRow As Integer =
                    Convert.ToInt32(NativeSensitivityView.GetRowCellValue(RowHandle, "SourceRow"))
                Sheet.Range.FromLTRB(0, SourceRow, 60, SourceRow).ClearContents()
            Next
            ActiveWorkbook.Calculate()
            RefreshNativeSensitivityList()
        Finally
            ProtectWS(ModelID, "Stress Sensitivity List")
        End Try

    End Sub

    Private Sub BuildNativePlannerPage()

        XtraTabPageMVP.Controls.Clear()
        Dim Root As New TableLayoutPanel With {
            .Dock = DockStyle.Fill, .BackColor = Color.White, .ColumnCount = 1, .RowCount = 2
        }
        Root.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
        Root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        Dim Toolbar As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill, .BackColor = Color.White,
            .Padding = New Padding(12, 14, 12, 8), .WrapContents = False
        }

        NativePlannerScenario = CreateNativeCombo(160)
        NativePlannerScenario.Properties.Items.AddRange(DefaultScenarioNames().Cast(Of Object).ToArray())
        NativePlannerScenario.SelectedIndex = 0
        NativePlannerName = New DevExpress.XtraEditors.TextEdit With {.Width = 230}
        NativePlannerImportMode = CreateNativeCombo(230)
        NativePlannerImportMode.Properties.Items.AddRange(New Object() {
            "Use assumptions below", "Import data from another business plan model"
        })
        NativePlannerInclude = New DevExpress.XtraEditors.CheckEdit With {
            .Text = "Include base in data import", .AutoSizeInLayoutControl = True
        }
        Dim CalculateButton As New DevExpress.XtraEditors.SimpleButton With {
            .Text = "Apply and recalculate", .Width = 155, .Height = 36
        }
        Dim GenerateButton As New DevExpress.XtraEditors.SimpleButton With {
            .Text = "Generate dashboard", .Width = 155, .Height = 36
        }
        Dim ClearButton As New DevExpress.XtraEditors.SimpleButton With {
            .Text = "Clear scenario", .Width = 120, .Height = 36
        }
        Toolbar.Controls.Add(CreateNativeLabel("Scenario"))
        Toolbar.Controls.Add(NativePlannerScenario)
        Toolbar.Controls.Add(CreateNativeLabel("Name"))
        Toolbar.Controls.Add(NativePlannerName)
        Toolbar.Controls.Add(NativePlannerImportMode)
        Toolbar.Controls.Add(NativePlannerInclude)
        Toolbar.Controls.Add(CalculateButton)
        Toolbar.Controls.Add(GenerateButton)
        Toolbar.Controls.Add(ClearButton)

        NativePlannerGrid = New GridControl With {.Dock = DockStyle.Fill}
        NativePlannerView = New GridView(NativePlannerGrid)
        NativePlannerGrid.MainView = NativePlannerView
        NativePlannerGrid.ViewCollection.Add(NativePlannerView)
        NativePlannerView.OptionsView.ShowGroupPanel = False
        NativePlannerView.OptionsView.ShowAutoFilterRow = True
        NativePlannerView.OptionsView.ColumnAutoWidth = False
        NativePlannerView.OptionsBehavior.EditorShowMode = DevExpress.Utils.EditorShowMode.MouseDownFocused

        NativeYearEditor = New RepositoryItemComboBox
        NativeYearEditor.Items.Add("")
        For YearNumber As Integer = 1 To 40
            NativeYearEditor.Items.Add(YearNumber)
        Next
        NativeYearEditor.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        NativePlannerGrid.RepositoryItems.Add(NativeYearEditor)
        NativePlannerGrid.RepositoryItems.Add(StandardPercentSpinEdit)
        NativePlannerGrid.RepositoryItems.Add(StandardYesEmptyEdit)
        NativePlannerGrid.RepositoryItems.Add(Standard2digitnumberTextBoxEdit)

        Root.Controls.Add(Toolbar, 0, 0)
        Root.Controls.Add(NativePlannerGrid, 0, 1)
        XtraTabPageMVP.Controls.Add(Root)

        AddHandler NativePlannerScenario.SelectedIndexChanged, AddressOf NativePlannerScenarioChanged
        AddHandler NativePlannerName.EditValueChanged, AddressOf NativePlannerNameChanged
        AddHandler NativePlannerImportMode.SelectedIndexChanged, AddressOf NativePlannerImportModeChanged
        AddHandler NativePlannerInclude.CheckedChanged, AddressOf NativePlannerIncludeChanged
        AddHandler NativePlannerView.CellValueChanged, AddressOf NativePlannerCellValueChanged
        AddHandler NativePlannerView.CustomRowCellEdit, AddressOf NativePlannerCustomRowCellEdit
        AddHandler NativePlannerView.CustomDrawCell, AddressOf NativePlannerCustomDrawCell
        AddHandler NativePlannerView.ShowingEditor, AddressOf NativePlannerShowingEditor
        AddHandler NativePlannerView.CustomColumnDisplayText, AddressOf NativePlannerCustomColumnDisplayText
        AddHandler CalculateButton.Click, Sub() RecalculateAndRefreshNativeViews()
        AddHandler GenerateButton.Click, AddressOf GenerateMultivariableDashboard_Click
        AddHandler ClearButton.Click, AddressOf ClearNativeScenario_Click

    End Sub

    Private Sub BuildNativeDashboardPage()

        XtraTabPageDashboard.Controls.Clear()
        Dim Root As New TableLayoutPanel With {
            .Dock = DockStyle.Fill, .BackColor = Color.White, .ColumnCount = 1, .RowCount = 2
        }
        Root.RowStyles.Add(New RowStyle(SizeType.Absolute, 62))
        Root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        Dim Toolbar As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill, .BackColor = Color.White,
            .Padding = New Padding(12, 12, 12, 6), .WrapContents = False
        }
        NativeDashboardScenario = CreateNativeCombo(240)
        Toolbar.Controls.Add(CreateNativeLabel("Captured scenario"))
        Toolbar.Controls.Add(NativeDashboardScenario)

        Dim Body As New TableLayoutPanel With {
            .Dock = DockStyle.Fill, .BackColor = Color.White, .ColumnCount = 2, .RowCount = 1
        }
        Body.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 72))
        Body.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 28))
        NativeDashboardCharts = New TableLayoutPanel With {
            .Dock = DockStyle.Fill, .BackColor = Color.White,
            .ColumnCount = 3, .RowCount = 2, .Padding = New Padding(6)
        }
        For Index As Integer = 1 To 3
            NativeDashboardCharts.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.333F))
        Next
        NativeDashboardCharts.RowStyles.Add(New RowStyle(SizeType.Percent, 50))
        NativeDashboardCharts.RowStyles.Add(New RowStyle(SizeType.Percent, 50))

        NativeDashboardSummary = CreateReadOnlyGrid()
        Body.Controls.Add(NativeDashboardCharts, 0, 0)
        Body.Controls.Add(NativeDashboardSummary, 1, 0)
        Root.Controls.Add(Toolbar, 0, 0)
        Root.Controls.Add(Body, 0, 1)
        XtraTabPageDashboard.Controls.Add(Root)
        AddHandler NativeDashboardScenario.SelectedIndexChanged, AddressOf NativeDashboardScenarioChanged

    End Sub

    Private Sub BuildNativeComparativePage(Page As DevExpress.XtraTab.XtraTabPage, IsSecondPage As Boolean)

        Page.Controls.Clear()
        Dim Root As New TableLayoutPanel With {
            .Dock = DockStyle.Fill, .BackColor = Color.White, .ColumnCount = 1, .RowCount = 3
        }
        Root.RowStyles.Add(New RowStyle(SizeType.Absolute, 62))
        Root.RowStyles.Add(New RowStyle(SizeType.Percent, 68))
        Root.RowStyles.Add(New RowStyle(SizeType.Percent, 32))
        Dim Toolbar As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill, .BackColor = Color.White,
            .Padding = New Padding(10, 10, 10, 4), .WrapContents = False
        }
        For ScenarioSlot As Integer = 0 To 3
            Dim Selector As DevExpress.XtraEditors.ComboBoxEdit = CreateNativeCombo(145)
            Selector.Tag = ScenarioSlot
            NativeComparativeSelectors.Add(Selector)
            Toolbar.Controls.Add(CreateNativeLabel("Comparison " & (ScenarioSlot + 1).ToString()))
            Toolbar.Controls.Add(Selector)
            AddHandler Selector.SelectedIndexChanged, AddressOf NativeComparativeScenarioChanged
        Next
        If IsSecondPage Then
            AddComparisonYearSelector(Toolbar, "Gearing start", 12)
            AddComparisonYearSelector(Toolbar, "Op margin start", 14)
            AddComparisonYearSelector(Toolbar, "Debt/unit start", 13)
        Else
            AddComparisonYearSelector(Toolbar, "Debt start", 10)
            AddComparisonYearSelector(Toolbar, "EBITDA start", 11)
        End If

        Dim Charts As New TableLayoutPanel With {
            .Dock = DockStyle.Fill, .BackColor = Color.White,
            .ColumnCount = If(IsSecondPage, 3, 2), .RowCount = 1, .Padding = New Padding(6)
        }
        For Index As Integer = 1 To Charts.ColumnCount
            Charts.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, CSng(100.0 / Charts.ColumnCount)))
        Next
        Dim Summary As GridControl = CreateReadOnlyGrid()
        If IsSecondPage Then
            NativeComparativeChartsB = Charts
            NativeComparativeSummaryB = Summary
        Else
            NativeComparativeChartsA = Charts
            NativeComparativeSummaryA = Summary
        End If
        Root.Controls.Add(Toolbar, 0, 0)
        Root.Controls.Add(Charts, 0, 1)
        Root.Controls.Add(Summary, 0, 2)
        Page.Controls.Add(Root)

    End Sub

    Private Function CreateReadOnlyGrid() As GridControl

        Dim Grid As New GridControl With {.Dock = DockStyle.Fill}
        Dim View As New GridView(Grid)
        Grid.MainView = View
        Grid.ViewCollection.Add(View)
        View.OptionsBehavior.Editable = False
        View.OptionsView.ShowGroupPanel = False
        View.OptionsView.ColumnAutoWidth = True
        Return Grid

    End Function

    Private Function CreateNativeLabel(Caption As String) As DevExpress.XtraEditors.LabelControl

        Return New DevExpress.XtraEditors.LabelControl With {
            .Text = Caption, .AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Default,
            .Margin = New Padding(8, 8, 4, 0)
        }

    End Function

    Private Function CreateNativeCombo(Width As Integer) As DevExpress.XtraEditors.ComboBoxEdit

        Dim Editor As New DevExpress.XtraEditors.ComboBoxEdit With {.Width = Width}
        Editor.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Return Editor

    End Function

    Private Sub AddComparisonYearSelector(Toolbar As FlowLayoutPanel, Caption As String, WorkingRow As Integer)

        Dim Editor As New DevExpress.XtraEditors.SpinEdit With {.Width = 70, .Tag = WorkingRow}
        Editor.Properties.MinValue = 0
        Editor.Properties.MaxValue = 20
        Editor.Properties.IsFloatValue = False
        Editor.EditValue = GetNumericValue(
            ActiveWorkbook.Worksheets("OW - Covenant Calculation").Cells(WorkingRow - 1, 2))
        Toolbar.Controls.Add(CreateNativeLabel(Caption))
        Toolbar.Controls.Add(Editor)
        AddHandler Editor.EditValueChanged, AddressOf NativeComparisonYearChanged

    End Sub

    Private Function DefaultScenarioNames() As List(Of String)

        Dim Names As New List(Of String) From {"Base Case"}
        For Index As Integer = 1 To 10
            Names.Add("Scenario " & Index.ToString())
        Next
        Return Names

    End Function

    Private Function WorkbookScenarioNames() As List(Of String)

        Dim Result As New List(Of String)
        Dim Sheet As DevExpress.Spreadsheet.Worksheet = ActiveWorkbook.Worksheets("Multivariable Planner")
        Dim Defaults As List(Of String) = DefaultScenarioNames()
        For ScenarioIndex As Integer = 0 To 10
            Dim NameValue As String = Sheet.Cells(7, ScenarioStartColumn(ScenarioIndex)).DisplayText.Trim()
            If String.IsNullOrWhiteSpace(NameValue) Then NameValue = Defaults(ScenarioIndex)
            Result.Add(NameValue)
        Next
        Return Result

    End Function

    Private Function ScenarioStartColumn(ScenarioIndex As Integer) As Integer

        Return 3 + (ScenarioIndex * 5)

    End Function

    Private Sub RefreshNativePlanner()

        If NativePlannerScenario Is Nothing Then Return
        LoadingNativeViews = True
        Try
            Dim ScenarioIndex As Integer = Math.Max(0, NativePlannerScenario.SelectedIndex)
            Dim StartColumn As Integer = ScenarioStartColumn(ScenarioIndex)
            Dim Sheet As DevExpress.Spreadsheet.Worksheet = ActiveWorkbook.Worksheets("Multivariable Planner")
            NativePlannerName.EditValue = Sheet.Cells(7, StartColumn).DisplayText
            NativePlannerInclude.Visible = ScenarioIndex = 0
            NativePlannerImportMode.Visible = ScenarioIndex > 0
            If ScenarioIndex = 0 Then
                NativePlannerInclude.Checked =
                    String.Equals(Sheet.Cells(6, StartColumn).DisplayText, "Yes", StringComparison.OrdinalIgnoreCase)
            Else
                NativePlannerImportMode.EditValue = Sheet.Cells(6, StartColumn).DisplayText
                If String.IsNullOrWhiteSpace(Convert.ToString(NativePlannerImportMode.EditValue)) Then
                    NativePlannerImportMode.EditValue = "Use assumptions below"
                End If
            End If

            Dim Data As New System.Data.DataTable
            Data.Columns.Add("SourceRow", GetType(Integer))
            Data.Columns.Add("Section", GetType(String))
            Data.Columns.Add("Assumption", GetType(String))
            Data.Columns.Add("ShortName", GetType(String))
            Data.Columns.Add("Change1", GetType(Object))
            Data.Columns.Add("Change2", GetType(Object))
            Data.Columns.Add("Change1FromYear", GetType(Object))
            Data.Columns.Add("Change2FromYear", GetType(Object))
            Data.Columns.Add("ToYear", GetType(Object))
            Data.Columns.Add("ValueFormat", GetType(String))
            AddPlannerRows(Data, Sheet, StartColumn, 9, 47, "Stress")
            AddPlannerRows(Data, Sheet, StartColumn, 51, 75, "Mitigation")
            NativePlannerData = Data
            NativePlannerGrid.DataSource = Data
            ConfigureNativePlannerColumns()
        Finally
            LoadingNativeViews = False
        End Try

    End Sub

    Private Sub AddPlannerRows(Data As System.Data.DataTable,
                               Sheet As DevExpress.Spreadsheet.Worksheet,
                               StartColumn As Integer,
                               FirstRow As Integer,
                               LastRow As Integer,
                               SectionName As String)

        For RowIndex As Integer = FirstRow To LastRow
            Dim Row As System.Data.DataRow = Data.NewRow()
            Row("SourceRow") = RowIndex
            Row("Section") = SectionName
            Row("Assumption") = Sheet.Cells(RowIndex, 0).DisplayText
            Row("ShortName") = Sheet.Cells(RowIndex, 1).DisplayText
            For ValueIndex As Integer = 0 To 4
                Row(4 + ValueIndex) = CellToObject(Sheet.Cells(RowIndex, StartColumn + ValueIndex))
            Next
            Row("ValueFormat") = Sheet.Cells(RowIndex, StartColumn).NumberFormat
            Data.Rows.Add(Row)
        Next

    End Sub

    Private Sub ConfigureNativePlannerColumns()

        If NativePlannerView.Columns.Count = 0 Then Return
        NativePlannerView.Columns("SourceRow").Visible = False
        NativePlannerView.Columns("ValueFormat").Visible = False
        NativePlannerView.Columns("Section").OptionsColumn.AllowEdit = False
        NativePlannerView.Columns("Assumption").OptionsColumn.AllowEdit = False
        NativePlannerView.Columns("ShortName").OptionsColumn.AllowEdit = True
        NativePlannerView.Columns("Section").GroupIndex = 0
        NativePlannerView.Columns("Section").SortOrder = DevExpress.Data.ColumnSortOrder.Descending
        NativePlannerView.Columns("Assumption").Width = 330
        NativePlannerView.Columns("ShortName").Width = 180
        NativePlannerView.Columns("Change1").Caption = "Change 1"
        NativePlannerView.Columns("Change2").Caption = "Change 2"
        NativePlannerView.Columns("Change1FromYear").Caption = "Change 1 from year"
        NativePlannerView.Columns("Change2FromYear").Caption = "Change 2 from year"
        NativePlannerView.Columns("ToYear").Caption = "To year"
        For Each ColumnName As String In {"Change1FromYear", "Change2FromYear", "ToYear"}
            NativePlannerView.Columns(ColumnName).ColumnEdit = NativeYearEditor
            NativePlannerView.Columns(ColumnName).Width = 125
        Next
        NativePlannerView.Columns("Change1").Width = 115
        NativePlannerView.Columns("Change2").Width = 115
        NativePlannerView.ExpandAllGroups()

    End Sub

    Private Function CellToObject(Cell As DevExpress.Spreadsheet.Cell) As Object

        If Cell.Value.IsEmpty Then Return DBNull.Value
        If Cell.Value.IsNumeric Then Return Cell.Value.NumericValue
        If Cell.Value.IsBoolean Then Return Cell.Value.BooleanValue
        If Cell.Value.IsDateTime Then Return Cell.Value.DateTimeValue
        Return Cell.Value.TextValue

    End Function

    Private Sub NativePlannerScenarioChanged(sender As Object, e As EventArgs)

        If Not LoadingNativeViews Then RefreshNativePlanner()

    End Sub

    Private Sub NativePlannerNameChanged(sender As Object, e As EventArgs)

        If LoadingNativeViews Then Return
        Dim Sheet As DevExpress.Spreadsheet.Worksheet = ActiveWorkbook.Worksheets("Multivariable Planner")
        Sheet.Cells(7, ScenarioStartColumn(Math.Max(0, NativePlannerScenario.SelectedIndex))).Value =
            CellValue.FromObject(Convert.ToString(NativePlannerName.EditValue))

    End Sub

    Private Sub NativePlannerImportModeChanged(sender As Object, e As EventArgs)

        If LoadingNativeViews OrElse NativePlannerScenario.SelectedIndex <= 0 Then Return
        ActiveWorkbook.Worksheets("Multivariable Planner").
            Cells(6, ScenarioStartColumn(NativePlannerScenario.SelectedIndex)).Value =
            CellValue.FromObject(Convert.ToString(NativePlannerImportMode.EditValue))

    End Sub

    Private Sub NativePlannerIncludeChanged(sender As Object, e As EventArgs)

        If LoadingNativeViews OrElse NativePlannerScenario.SelectedIndex <> 0 Then Return
        ActiveWorkbook.Worksheets("Multivariable Planner").Cells(6, 3).Value =
            CellValue.FromObject(If(NativePlannerInclude.Checked, "Yes", ""))

    End Sub

    Private Sub NativePlannerCellValueChanged(sender As Object,
                                              e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs)

        If LoadingNativeViews OrElse e.RowHandle < 0 Then Return

        Dim SourceRow As Integer =
            Convert.ToInt32(NativePlannerView.GetRowCellValue(e.RowHandle, "SourceRow"))

        If e.Column.FieldName = "ShortName" Then
            Dim ShortNameCell As DevExpress.Spreadsheet.Cell =
                ActiveWorkbook.Worksheets("Multivariable Planner").Cells(SourceRow, 1)

            If e.Value Is Nothing OrElse e.Value Is DBNull.Value OrElse
               String.IsNullOrWhiteSpace(Convert.ToString(e.Value)) Then
                ShortNameCell.ClearContents()
            Else
                ShortNameCell.Value = CellValue.FromObject(Convert.ToString(e.Value))
            End If
            Return
        End If

        Dim ValueOffset As Integer
        Select Case e.Column.FieldName
            Case "Change1" : ValueOffset = 0
            Case "Change2" : ValueOffset = 1
            Case "Change1FromYear" : ValueOffset = 2
            Case "Change2FromYear" : ValueOffset = 3
            Case "ToYear" : ValueOffset = 4
            Case Else : Return
        End Select
        Dim Target As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.Worksheets("Multivariable Planner").
                Cells(SourceRow,
                      ScenarioStartColumn(Math.Max(0, NativePlannerScenario.SelectedIndex)) + ValueOffset)
        If e.Value Is Nothing OrElse e.Value Is DBNull.Value OrElse
           String.IsNullOrWhiteSpace(Convert.ToString(e.Value)) Then
            Target.ClearContents()
        Else
            Target.Value = CellValue.FromObject(e.Value)
        End If

    End Sub

    Private Sub NativePlannerCustomDrawCell(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs)

        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If Not TryGetNativePlannerSourceCell(e.RowHandle, e.Column, SourceCell) Then Return

        ApplyWorkbookCellAppearance(e.Appearance, SourceCell)

        If NativePlannerView.IsCellSelected(e.RowHandle, e.Column) Then
            e.Appearance.BackColor = Color.Beige
            e.Appearance.ForeColor = Color.Black
        End If

        e.DefaultDraw()
        e.Handled = True

    End Sub

    Private Sub NativePlannerShowingEditor(sender As Object, e As CancelEventArgs)

        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If Not TryGetNativePlannerSourceCell(
                NativePlannerView.FocusedRowHandle,
                NativePlannerView.FocusedColumn,
                SourceCell) OrElse SourceCell.Protection.Locked Then
            e.Cancel = True
        End If

    End Sub

    Private Function TryGetNativePlannerSourceCell(
        RowHandle As Integer,
        Column As DevExpress.XtraGrid.Columns.GridColumn,
        ByRef SourceCell As DevExpress.Spreadsheet.Cell) As Boolean

        If RowHandle < 0 OrElse Column Is Nothing Then Return False

        Dim SourceRowValue As Object =
            NativePlannerView.GetRowCellValue(RowHandle, "SourceRow")
        If SourceRowValue Is Nothing OrElse SourceRowValue Is DBNull.Value Then Return False

        Dim SourceColumn As Integer
        Select Case Column.FieldName
            Case "Assumption" : SourceColumn = 0
            Case "ShortName" : SourceColumn = 1
            Case "Change1" : SourceColumn = ScenarioStartColumn(Math.Max(0, NativePlannerScenario.SelectedIndex))
            Case "Change2" : SourceColumn = ScenarioStartColumn(Math.Max(0, NativePlannerScenario.SelectedIndex)) + 1
            Case "Change1FromYear" : SourceColumn = ScenarioStartColumn(Math.Max(0, NativePlannerScenario.SelectedIndex)) + 2
            Case "Change2FromYear" : SourceColumn = ScenarioStartColumn(Math.Max(0, NativePlannerScenario.SelectedIndex)) + 3
            Case "ToYear" : SourceColumn = ScenarioStartColumn(Math.Max(0, NativePlannerScenario.SelectedIndex)) + 4
            Case Else : Return False
        End Select

        SourceCell = ActiveWorkbook.Worksheets("Multivariable Planner").Cells(
            Convert.ToInt32(SourceRowValue), SourceColumn)
        Return SourceCell IsNot Nothing

    End Function

    Private Sub ApplyWorkbookCellAppearance(
        Appearance As DevExpress.Utils.AppearanceObject,
        SourceCell As DevExpress.Spreadsheet.Cell)

        If SourceCell Is Nothing Then Return

        Dim Background As Color = SourceCell.Fill.BackgroundColor
        If Background.IsEmpty OrElse Background.A = 0 Then Background = Color.White

        Dim Foreground As Color = SourceCell.Font.Color
        If Foreground.IsEmpty OrElse Foreground.A = 0 Then Foreground = AbovoBlue

        Appearance.BackColor = Background
        Appearance.ForeColor = Foreground
        Appearance.Options.UseBackColor = True
        Appearance.Options.UseForeColor = True

    End Sub

    Private Sub NativePlannerCustomRowCellEdit(sender As Object, e As CustomRowCellEditEventArgs)

        If e.RowHandle < 0 Then Return
        If e.Column.FieldName = "Change1FromYear" OrElse
           e.Column.FieldName = "Change2FromYear" OrElse e.Column.FieldName = "ToYear" Then
            e.RepositoryItem = NativeYearEditor
            Return
        End If
        If e.Column.FieldName <> "Change1" AndAlso e.Column.FieldName <> "Change2" Then Return
        Dim SourceRow As Integer =
            Convert.ToInt32(NativePlannerView.GetRowCellValue(e.RowHandle, "SourceRow"))
        Dim Format As String =
            Convert.ToString(NativePlannerView.GetRowCellValue(e.RowHandle, "ValueFormat"))
        If SourceRow = 34 OrElse SourceRow = 64 OrElse SourceRow = 65 Then
            e.RepositoryItem = StandardYesEmptyEdit
        ElseIf Format.Contains("%") Then
            e.RepositoryItem = StandardPercentSpinEdit
        Else
            e.RepositoryItem = Standard2digitnumberTextBoxEdit
        End If

    End Sub

    Private Sub NativePlannerCustomColumnDisplayText(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs)

        If e.ListSourceRowIndex < 0 OrElse
           (e.Column.FieldName <> "Change1" AndAlso e.Column.FieldName <> "Change2") OrElse
           e.Value Is Nothing OrElse e.Value Is DBNull.Value Then Return
        Dim RowHandle As Integer = NativePlannerView.GetRowHandle(e.ListSourceRowIndex)
        Dim Format As String =
            Convert.ToString(NativePlannerView.GetRowCellValue(RowHandle, "ValueFormat"))
        Dim NumericValue As Double
        If Not Double.TryParse(Convert.ToString(e.Value), NumericValue) Then Return
        If Format.Contains("%") Then
            e.DisplayText = NumericValue.ToString("P2")
        ElseIf Format.Contains(ChrW(&HA3)) OrElse Format.Contains("$") Then
            e.DisplayText = NumericValue.ToString(
                "C0", Globalization.CultureInfo.GetCultureInfo("en-GB"))
        Else
            e.DisplayText = NumericValue.ToString("N2")
        End If

    End Sub

    Private Sub ClearNativeScenario_Click(sender As Object, e As EventArgs)

        Dim ScenarioIndex As Integer = Math.Max(0, NativePlannerScenario.SelectedIndex)
        If DevExpress.XtraEditors.XtraMessageBox.Show(
                "Clear the selected scenario definition?", "Clear scenario",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        Dim Sheet As DevExpress.Spreadsheet.Worksheet = ActiveWorkbook.Worksheets("Multivariable Planner")
        Dim StartColumn As Integer = ScenarioStartColumn(ScenarioIndex)
        Sheet.Range.FromLTRB(StartColumn, 9, StartColumn + 4, 47).ClearContents()
        Sheet.Range.FromLTRB(StartColumn, 51, StartColumn + 4, 75).ClearContents()
        If ScenarioIndex > 0 Then
            ActiveWorkbook.DefinedNames.GetDefinedName(
                "S" & ScenarioIndex.ToString() & "Data").Range.ClearContents()
        End If
        ActiveWorkbook.Calculate()
        RefreshNativePlanner()

    End Sub

    Private Sub RecalculateAndRefreshNativeViews()

        Me.Cursor = Cursors.WaitCursor
        Try
            ActiveWorkbook.Calculate()
            RefreshNativeDashboard()
            RefreshNativeComparativeViews()
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub LiveScenarioNumberChanged(sender As Object, e As EventArgs)

        If ComboBoxBreachMode.SelectedIndex >= 0 Then
            ActiveWorkbook.DefinedNames.GetDefinedName("StressTestNumber").Range(0, 0).Value =
                CellValue.FromObject(ComboBoxBreachMode.SelectedItem.ToString())
        End If

    End Sub

    Private Sub CaptureCurrentLiveScenario_Click(sender As Object, e As EventArgs)

        If ComboBoxBreachMode.SelectedIndex < 0 Then
            DevExpress.XtraEditors.XtraMessageBox.Show(
                "Select a test number first.", "Capture multivariable assumptions")
            Return
        End If
        Dim ScenarioName As String = Convert.ToString(TextEditMultivariableName.EditValue).Trim()
        If String.IsNullOrWhiteSpace(ScenarioName) OrElse
           String.Equals(ScenarioName, "Base", StringComparison.OrdinalIgnoreCase) Then
            DevExpress.XtraEditors.XtraMessageBox.Show(
                "Enter a multivariable scenario name first.", "Capture multivariable assumptions")
            Return
        End If

        Dim ScenarioIndex As Integer = ComboBoxBreachMode.SelectedIndex + 1
        Me.Cursor = Cursors.WaitCursor
        UNProtectWS(ModelID, "Multivariable Planner")
        UNProtectWS(ModelID, "OW - Captured Data")
        Try
            CopyRangeValues(
                ActiveWorkbook.DefinedNames.GetDefinedName("LiveAssumptions").Range,
                ActiveWorkbook.DefinedNames.GetDefinedName(
                    "Assumptions" & ScenarioIndex.ToString()).Range)
            CopyRangeValues(
                ActiveWorkbook.DefinedNames.GetDefinedName("LiveAssumptionsA").Range,
                ActiveWorkbook.DefinedNames.GetDefinedName(
                    "AssumptionsA" & ScenarioIndex.ToString()).Range)
            ActiveWorkbook.Worksheets("Multivariable Planner").
                Cells(7, ScenarioStartColumn(ScenarioIndex)).Value =
                CellValue.FromObject(ScenarioName)
            ActiveWorkbook.DefinedNames.GetDefinedName(
                "ImportMode" & ScenarioIndex.ToString()).Range(0, 0).Value =
                CellValue.FromObject("Use assumptions below")
            CopyRangeValues(
                ActiveWorkbook.DefinedNames.GetDefinedName("StressLiveInfo").Range,
                ActiveWorkbook.DefinedNames.GetDefinedName(
                    "S" & ScenarioIndex.ToString() & "Data").Range)
            ActiveWorkbook.Calculate()
            RefreshAllNativeScenarioSelectors()
            RefreshNativePlanner()
            RefreshNativeDashboard()
            RefreshNativeComparativeViews()
        Finally
            ProtectWS(ModelID, "Multivariable Planner")
            ProtectWS(ModelID, "OW - Captured Data")
            Me.Cursor = Cursors.Default
        End Try
        DevExpress.XtraEditors.XtraMessageBox.Show(
            "The current live assumptions and results were captured as " &
            ScenarioName & ".", "Capture complete")

    End Sub

    Private Sub GenerateMultivariableDashboard_Click(sender As Object, e As EventArgs)

        If DevExpress.XtraEditors.XtraMessageBox.Show(
                "Run and capture every configured multivariable scenario? This may take several minutes.",
                "Generate multivariable dashboard",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) <> DialogResult.Yes Then Return

        Dim LiveAssumptions As DevExpress.Spreadsheet.CellRange =
            ActiveWorkbook.DefinedNames.GetDefinedName("LiveAssumptions").Range
        Dim LiveAssumptionsA As DevExpress.Spreadsheet.CellRange =
            ActiveWorkbook.DefinedNames.GetDefinedName("LiveAssumptionsA").Range
        Dim SavedLive As CellValue(,) = SnapshotRange(LiveAssumptions)
        Dim SavedLiveA As CellValue(,) = SnapshotRange(LiveAssumptionsA)
        Dim SavedMode As String =
            ActiveWorkbook.DefinedNames.GetDefinedName("StressTestMode").Range(0, 0).DisplayText
        Dim SavedModeReference As String =
            ActiveWorkbook.DefinedNames.GetDefinedName("Mode").RefersTo
        Dim OriginalTitle As String = Me.Text

        Me.Cursor = Cursors.WaitCursor
        UNProtectWS(ModelID, "Multivariable Planner")
        UNProtectWS(ModelID, "OW - Captured Data")
        UNProtectWS(ModelID, "Live Multivariable Planner")
        Try
            SetWorkbookStressMode(False)
            LiveAssumptions.ClearContents()
            LiveAssumptionsA.ClearContents()
            ActiveWorkbook.Calculate()
            CopyRangeValues(
                ActiveWorkbook.DefinedNames.GetDefinedName("StressLiveInfo").Range,
                ActiveWorkbook.DefinedNames.GetDefinedName("S0Data").Range)

            Dim Names As List(Of String) = WorkbookScenarioNames()
            For ScenarioIndex As Integer = 1 To 10
                Me.Text = "Generating dashboard - " & Names(ScenarioIndex)
                Windows.Forms.Application.DoEvents()
                Dim TargetData As DevExpress.Spreadsheet.CellRange =
                    ActiveWorkbook.DefinedNames.GetDefinedName(
                        "S" & ScenarioIndex.ToString() & "Data").Range
                If String.IsNullOrWhiteSpace(Names(ScenarioIndex)) Then
                    TargetData.ClearContents()
                    Continue For
                End If

                Dim ImportMode As String =
                    ActiveWorkbook.DefinedNames.GetDefinedName(
                        "ImportMode" & ScenarioIndex.ToString()).Range(0, 0).DisplayText
                If Not String.Equals(
                        ImportMode, "Use assumptions below",
                        StringComparison.OrdinalIgnoreCase) Then
                    If Not ImportScenarioResults(Names(ScenarioIndex), TargetData) Then Continue For
                Else
                    SetWorkbookStressMode(True)
                    CopyRangeValues(
                        ActiveWorkbook.DefinedNames.GetDefinedName(
                            "Assumptions" & ScenarioIndex.ToString()).Range,
                        LiveAssumptions)
                    CopyRangeValues(
                        ActiveWorkbook.DefinedNames.GetDefinedName(
                            "AssumptionsA" & ScenarioIndex.ToString()).Range,
                        LiveAssumptionsA)
                    ActiveWorkbook.Calculate()
                    CopyRangeValues(
                        ActiveWorkbook.DefinedNames.GetDefinedName("StressLiveInfo").Range,
                        TargetData)
                    SetWorkbookStressMode(False)
                End If
            Next
        Catch ex As Exception
            DevExpress.XtraEditors.XtraMessageBox.Show(
                "Dashboard generation stopped: " & ex.Message,
                "Multivariable dashboard", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            RestoreRange(LiveAssumptions, SavedLive)
            RestoreRange(LiveAssumptionsA, SavedLiveA)
            ActiveWorkbook.DefinedNames.GetDefinedName(
                "StressTestMode").Range(0, 0).Value = CellValue.FromObject(SavedMode)
            ActiveWorkbook.DefinedNames.GetDefinedName("Mode").RefersTo = SavedModeReference
            ActiveWorkbook.Calculate()
            ProtectWS(ModelID, "Multivariable Planner")
            ProtectWS(ModelID, "OW - Captured Data")
            ProtectWS(ModelID, "Live Multivariable Planner")
            Me.Text = OriginalTitle
            Me.Cursor = Cursors.Default
        End Try

        RefreshAllNativeScenarioSelectors()
        RefreshNativeDashboard()
        RefreshNativeComparativeViews()
        XtraTabControlStressTest.SelectedTabPage = XtraTabPageDashboard

    End Sub

    Private Sub SetWorkbookStressMode(Enabled As Boolean)

        ActiveWorkbook.DefinedNames.GetDefinedName("StressTestMode").Range(0, 0).Value =
            CellValue.FromObject(If(Enabled, "Y", "N"))
        ActiveWorkbook.DefinedNames.GetDefinedName("Mode").RefersTo =
            If(Enabled, """Stress Test""", """Business Plan""")

    End Sub

    Private Function ImportScenarioResults(
        ScenarioName As String,
        TargetData As DevExpress.Spreadsheet.CellRange) As Boolean

        Using Dialog As New DevExpress.XtraEditors.XtraOpenFileDialog()
            Dialog.Title = "Select business plan results for " & ScenarioName
            Dialog.Filter = "Abovo business plan models|*.xlsb;*.abp;*.xlsm;*.xlsx"
            If Dialog.ShowDialog() <> DialogResult.OK Then Return False
            Using SourceWorkbook As New DevExpress.Spreadsheet.Workbook()
                SourceWorkbook.LoadDocument(Dialog.FileName)
                Dim SourceName As DevExpress.Spreadsheet.DefinedName =
                    SourceWorkbook.DefinedNames.GetDefinedName("StressLiveInfo")
                If SourceName Is Nothing OrElse SourceName.Range Is Nothing Then
                    DevExpress.XtraEditors.XtraMessageBox.Show(
                        "The selected model does not contain the StressLiveInfo results range.",
                        "Import scenario")
                    Return False
                End If
                CopyRangeValues(SourceName.Range, TargetData)
            End Using
        End Using
        Return True

    End Function

    Private Sub CopyRangeValues(Source As DevExpress.Spreadsheet.CellRange,
                                Target As DevExpress.Spreadsheet.CellRange)

        If Source.RowCount <> Target.RowCount OrElse
           Source.ColumnCount <> Target.ColumnCount Then
            Throw New InvalidOperationException(
                "Stress-test source and target ranges do not have matching dimensions.")
        End If
        For RowIndex As Integer = 0 To Source.RowCount - 1
            For ColumnIndex As Integer = 0 To Source.ColumnCount - 1
                Target(RowIndex, ColumnIndex).Value = Source(RowIndex, ColumnIndex).Value
            Next
        Next

    End Sub

    Private Function SnapshotRange(
        Source As DevExpress.Spreadsheet.CellRange) As CellValue(,)

        Dim Values(Source.RowCount - 1, Source.ColumnCount - 1) As CellValue
        For RowIndex As Integer = 0 To Source.RowCount - 1
            For ColumnIndex As Integer = 0 To Source.ColumnCount - 1
                Values(RowIndex, ColumnIndex) = Source(RowIndex, ColumnIndex).Value
            Next
        Next
        Return Values

    End Function

    Private Sub RestoreRange(
        Target As DevExpress.Spreadsheet.CellRange, Values As CellValue(,))

        For RowIndex As Integer = 0 To Target.RowCount - 1
            For ColumnIndex As Integer = 0 To Target.ColumnCount - 1
                Target(RowIndex, ColumnIndex).Value = Values(RowIndex, ColumnIndex)
            Next
        Next

    End Sub

    Private Sub RefreshAllNativeScenarioSelectors()

        Dim Names As List(Of String) = WorkbookScenarioNames()
        LoadingNativeViews = True
        Try
            If NativePlannerScenario IsNot Nothing Then
                Dim Selected As Integer = Math.Max(0, NativePlannerScenario.SelectedIndex)
                NativePlannerScenario.Properties.Items.Clear()
                NativePlannerScenario.Properties.Items.AddRange(Names.Cast(Of Object).ToArray())
                NativePlannerScenario.SelectedIndex = Selected
            End If
            If NativeDashboardScenario IsNot Nothing Then
                Dim Selected As Integer = Math.Max(0, NativeDashboardScenario.SelectedIndex)
                NativeDashboardScenario.Properties.Items.Clear()
                NativeDashboardScenario.Properties.Items.AddRange(Names.Cast(Of Object).ToArray())
                NativeDashboardScenario.SelectedIndex = Math.Min(Selected, Names.Count - 1)
            End If
            For Each Selector As DevExpress.XtraEditors.ComboBoxEdit In NativeComparativeSelectors
                Dim SelectedText As String = Convert.ToString(Selector.EditValue)
                Selector.Properties.Items.Clear()
                Selector.Properties.Items.AddRange(Names.Cast(Of Object).ToArray())
                Dim Match As Integer = Names.FindIndex(
                    Function(Item) String.Equals(
                        Item, SelectedText, StringComparison.OrdinalIgnoreCase))
                Selector.SelectedIndex =
                    If(Match >= 0, Match,
                       Math.Min(Convert.ToInt32(Selector.Tag) + 1, Names.Count - 1))
            Next
        Finally
            LoadingNativeViews = False
        End Try

    End Sub

    Private Sub NativeDashboardScenarioChanged(sender As Object, e As EventArgs)

        If Not LoadingNativeViews Then RefreshNativeDashboard()

    End Sub

    Private Sub RefreshNativeDashboard()

        If NativeDashboardScenario Is Nothing Then Return
        Dim Names As List(Of String) = WorkbookScenarioNames()
        LoadingNativeViews = True
        Try
            Dim PreviousIndex As Integer = Math.Max(0, NativeDashboardScenario.SelectedIndex)
            NativeDashboardScenario.Properties.Items.Clear()
            NativeDashboardScenario.Properties.Items.AddRange(Names.Cast(Of Object).ToArray())
            NativeDashboardScenario.SelectedIndex = Math.Min(PreviousIndex, Names.Count - 1)
            Dim ScenarioIndex As Integer = Math.Max(0, NativeDashboardScenario.SelectedIndex)
            ActiveWorkbook.Worksheets("Multivariable Dashboard").Range("E6")(0, 0).Value =
                CellValue.FromObject(Names(ScenarioIndex))
            ActiveWorkbook.Calculate()

            Dim BaseData As DevExpress.Spreadsheet.CellRange =
                ActiveWorkbook.DefinedNames.GetDefinedName("S0Data").Range
            Dim ScenarioData As DevExpress.Spreadsheet.CellRange =
                ActiveWorkbook.DefinedNames.GetDefinedName(
                    "S" & ScenarioIndex.ToString() & "Data").Range
            Dim TargetData As DevExpress.Spreadsheet.CellRange =
                ActiveWorkbook.Worksheets("Multivariable Planner").Range("BI10:BM49")
            Dim YearData As DevExpress.Spreadsheet.CellRange =
                ActiveWorkbook.Worksheets("OW - Live Stress Reporting").Range("B8:B47")
            Dim DirectionData As DevExpress.Spreadsheet.CellRange =
                ActiveWorkbook.Worksheets("Multivariable Planner").Range("BO8:BS8")
            Dim MetricNames As String() = {
                "Gearing", "Operating Margin", "EBITDA MRI", "Debt / Unit", "Debt"
            }

            NativeDashboardCharts.SuspendLayout()
            NativeDashboardCharts.Controls.Clear()
            Dim Summary As New System.Data.DataTable
            Summary.Columns.Add("Metric", GetType(String))
            Summary.Columns.Add("Rule", GetType(String))
            Summary.Columns.Add("Breaches", GetType(Integer))
            Summary.Columns.Add("WorstValue", GetType(String))
            Summary.Columns.Add("WorstYear", GetType(String))

            For MetricIndex As Integer = 0 To 4
                Dim Chart As ChartControl = BuildMetricChart(
                    MetricNames(MetricIndex), YearData, TargetData, BaseData,
                    ScenarioData, MetricIndex, Names(ScenarioIndex))
                NativeDashboardCharts.Controls.Add(
                    Chart, MetricIndex Mod 3, MetricIndex \ 3)

                Dim Direction As String = DirectionData(0, MetricIndex).DisplayText
                Dim Breaches As Integer = 0
                Dim IsMinimum As Boolean =
                    String.Equals(Direction, "Greater", StringComparison.OrdinalIgnoreCase)
                Dim WorstValue As Double = If(IsMinimum, Double.MaxValue, Double.MinValue)
                Dim WorstYear As String = ""
                For RowIndex As Integer = 0 To 39
                    Dim Actual As Double = GetNumericValue(ScenarioData(RowIndex, MetricIndex))
                    Dim Target As Double = GetNumericValue(TargetData(RowIndex, MetricIndex))
                    If If(IsMinimum, Actual < Target, Actual > Target) Then Breaches += 1
                    If (IsMinimum AndAlso Actual < WorstValue) OrElse
                       (Not IsMinimum AndAlso Actual > WorstValue) Then
                        WorstValue = Actual
                        WorstYear = YearData(RowIndex, 0).DisplayText
                    End If
                Next
                Summary.Rows.Add(
                    MetricNames(MetricIndex), If(IsMinimum, "Minimum", "Maximum"),
                    Breaches, FormatMetricValue(MetricIndex, WorstValue), WorstYear)
            Next
            NativeDashboardCharts.ResumeLayout()
            NativeDashboardSummary.DataSource = Summary
            CType(NativeDashboardSummary.MainView, GridView).BestFitColumns()
        Finally
            LoadingNativeViews = False
        End Try

    End Sub

    Private Function BuildMetricChart(
        Title As String,
        Years As DevExpress.Spreadsheet.CellRange,
        Targets As DevExpress.Spreadsheet.CellRange,
        BaseData As DevExpress.Spreadsheet.CellRange,
        ScenarioData As DevExpress.Spreadsheet.CellRange,
        MetricIndex As Integer,
        ScenarioName As String) As ChartControl

        Dim Chart As New ChartControl With {.Dock = DockStyle.Fill, .BackColor = Color.White}
        Dim TargetSeries As New DevExpress.XtraCharts.Series("Target", ViewType.Line)
        Dim BaseSeries As New DevExpress.XtraCharts.Series("Base Case", ViewType.Line)
        Dim ScenarioSeries As New DevExpress.XtraCharts.Series(ScenarioName, ViewType.Line)
        For RowIndex As Integer = 0 To 39
            Dim YearLabel As String = Years(RowIndex, 0).DisplayText
            AddSeriesPoint(TargetSeries, YearLabel, Targets(RowIndex, MetricIndex))
            AddSeriesPoint(BaseSeries, YearLabel, BaseData(RowIndex, MetricIndex))
            AddSeriesPoint(ScenarioSeries, YearLabel, ScenarioData(RowIndex, MetricIndex))
        Next
        Chart.Series.Add(TargetSeries)
        Chart.Series.Add(BaseSeries)
        If Not String.Equals(
                ScenarioName, "Base Case", StringComparison.OrdinalIgnoreCase) Then
            Chart.Series.Add(ScenarioSeries)
        End If
        Chart.Titles.Add(New DevExpress.XtraCharts.ChartTitle With {.Text = Title})
        Chart.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True
        Dim Diagram As XYDiagram = TryCast(Chart.Diagram, XYDiagram)
        If Diagram IsNot Nothing Then
            Diagram.AxisX.Label.Angle = -45
            Diagram.AxisX.Label.ResolveOverlappingOptions.AllowRotate = True
        End If
        Return Chart

    End Function

    Private Sub AddSeriesPoint(
        Series As DevExpress.XtraCharts.Series,
        Argument As String,
        Cell As DevExpress.Spreadsheet.Cell)

        If Cell.Value.IsNumeric Then
            Series.Points.Add(New SeriesPoint(Argument, Cell.Value.NumericValue))
        End If

    End Sub

    Private Function GetNumericValue(Cell As DevExpress.Spreadsheet.Cell) As Double

        If Cell IsNot Nothing AndAlso Cell.Value.IsNumeric Then Return Cell.Value.NumericValue
        Return 0

    End Function

    Private Function FormatMetricValue(MetricIndex As Integer, Value As Double) As String

        If MetricIndex <= 2 Then Return Value.ToString("P1")
        Return Value.ToString("N1")

    End Function

    Private Sub NativeComparativeScenarioChanged(sender As Object, e As EventArgs)

        If LoadingNativeViews Then Return
        Dim Selector As DevExpress.XtraEditors.ComboBoxEdit =
            TryCast(sender, DevExpress.XtraEditors.ComboBoxEdit)
        If Selector Is Nothing Then Return
        Dim Slot As Integer = Convert.ToInt32(Selector.Tag)
        ActiveWorkbook.Worksheets("Comparative").Cells(11 + (Slot * 2), 2).Value =
            CellValue.FromObject(Convert.ToString(Selector.EditValue))
        ActiveWorkbook.Calculate()
        RefreshNativeComparativeViews()

    End Sub

    Private Sub NativeComparisonYearChanged(sender As Object, e As EventArgs)

        If LoadingNativeViews Then Return
        Dim Editor As DevExpress.XtraEditors.SpinEdit =
            TryCast(sender, DevExpress.XtraEditors.SpinEdit)
        If Editor Is Nothing Then Return
        Dim WorkingRow As Integer = Convert.ToInt32(Editor.Tag)
        ActiveWorkbook.Worksheets("OW - Covenant Calculation").
            Cells(WorkingRow - 1, 2).Value =
            CellValue.FromObject(Convert.ToInt32(Editor.Value))
        ActiveWorkbook.Calculate()
        RefreshNativeComparativeViews()

    End Sub

    Private Sub RefreshNativeComparativeViews()

        If NativeComparativeChartsA Is Nothing OrElse
           NativeComparativeChartsB Is Nothing Then Return
        Dim Names As List(Of String) = WorkbookScenarioNames()
        LoadingNativeViews = True
        Try
            Dim Sheet As DevExpress.Spreadsheet.Worksheet =
                ActiveWorkbook.Worksheets("Comparative")
            For Each Selector As DevExpress.XtraEditors.ComboBoxEdit In NativeComparativeSelectors
                Dim Slot As Integer = Convert.ToInt32(Selector.Tag)
                Dim SelectedName As String =
                    Sheet.Cells(11 + (Slot * 2), 2).DisplayText
                Selector.Properties.Items.Clear()
                Selector.Properties.Items.AddRange(Names.Cast(Of Object).ToArray())
                Dim Match As Integer = Names.FindIndex(
                    Function(Item) String.Equals(
                        Item, SelectedName, StringComparison.OrdinalIgnoreCase))
                Selector.SelectedIndex =
                    If(Match >= 0, Match, Math.Min(Slot + 1, Names.Count - 1))
            Next

            ActiveWorkbook.Calculate()
            Dim Working As DevExpress.Spreadsheet.Worksheet =
                ActiveWorkbook.Worksheets("OW - Covenant Calculation")
            NativeComparativeChartsA.SuspendLayout()
            NativeComparativeChartsA.Controls.Clear()
            NativeComparativeChartsA.Controls.Add(
                BuildComparisonChart("Debt", Working, 62, 63, 68), 0, 0)
            NativeComparativeChartsA.Controls.Add(
                BuildComparisonChart("EBITDA MRI", Working, 70, 71, 76), 1, 0)
            NativeComparativeChartsA.ResumeLayout()

            NativeComparativeChartsB.SuspendLayout()
            NativeComparativeChartsB.Controls.Clear()
            NativeComparativeChartsB.Controls.Add(
                BuildComparisonChart("Gearing", Working, 82, 83, 88), 0, 0)
            NativeComparativeChartsB.Controls.Add(
                BuildComparisonChart("Operating Margin", Working, 98, 99, 104), 1, 0)
            NativeComparativeChartsB.Controls.Add(
                BuildComparisonChart("Debt / Unit", Working, 90, 91, 96), 2, 0)
            NativeComparativeChartsB.ResumeLayout()

            NativeComparativeSummaryA.DataSource = BuildComparativeSummaryA()
            CType(NativeComparativeSummaryA.MainView, GridView).BestFitColumns()
            NativeComparativeSummaryB.DataSource = BuildComparativeSummaryB()
            CType(NativeComparativeSummaryB.MainView, GridView).BestFitColumns()
        Finally
            LoadingNativeViews = False
        End Try

    End Sub

    Private Function BuildComparisonChart(
        Title As String,
        Sheet As DevExpress.Spreadsheet.Worksheet,
        ArgumentColumn As Integer,
        FirstSeriesColumn As Integer,
        LastSeriesColumn As Integer) As ChartControl

        Dim Chart As New ChartControl With {.Dock = DockStyle.Fill, .BackColor = Color.White}
        For SeriesColumn As Integer = FirstSeriesColumn To LastSeriesColumn
            Dim SeriesName As String = Sheet.Cells(17, SeriesColumn).DisplayText
            If String.IsNullOrWhiteSpace(SeriesName) Then
                SeriesName = "Series " & (SeriesColumn - FirstSeriesColumn + 1).ToString()
            End If
            Dim NewSeries As New DevExpress.XtraCharts.Series(SeriesName, ViewType.Line)
            For RowIndex As Integer = 18 To 37
                AddSeriesPoint(
                    NewSeries, Sheet.Cells(RowIndex, ArgumentColumn).DisplayText,
                    Sheet.Cells(RowIndex, SeriesColumn))
            Next
            Chart.Series.Add(NewSeries)
        Next
        Chart.Titles.Add(New DevExpress.XtraCharts.ChartTitle With {.Text = Title})
        Chart.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True
        Dim Diagram As XYDiagram = TryCast(Chart.Diagram, XYDiagram)
        If Diagram IsNot Nothing Then
            Diagram.AxisX.Label.Angle = -45
            Diagram.AxisX.Label.ResolveOverlappingOptions.AllowRotate = True
        End If
        Return Chart

    End Function

    Private Function BuildComparativeSummaryA() As System.Data.DataTable

        Dim Table As New System.Data.DataTable
        Table.Columns.Add("Scenario", GetType(String))
        Table.Columns.Add("Peak Debt", GetType(String))
        Table.Columns.Add("Peak Debt Year", GetType(String))
        Table.Columns.Add("Repayment Year", GetType(String))
        Table.Columns.Add("Max Debt / Unit", GetType(String))
        Table.Columns.Add("Min EBITDA MRI", GetType(String))
        Table.Columns.Add("EBITDA Year", GetType(String))
        Dim Sheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Comparative")
        For ResultIndex As Integer = 0 To 5
            Dim SourceRow As Integer = 7 + (ResultIndex * 2)
            Dim EbitdaRow As Integer = 35 + (ResultIndex * 2)
            Table.Rows.Add(
                Sheet.Cells(SourceRow, 6).DisplayText,
                Sheet.Cells(SourceRow, 8).DisplayText,
                Sheet.Cells(SourceRow, 9).DisplayText,
                Sheet.Cells(SourceRow, 10).DisplayText,
                Sheet.Cells(SourceRow, 11).DisplayText,
                Sheet.Cells(EbitdaRow, 15).DisplayText,
                Sheet.Cells(EbitdaRow, 17).DisplayText)
        Next
        Return Table

    End Function

    Private Function BuildComparativeSummaryB() As System.Data.DataTable

        Dim Table As New System.Data.DataTable
        Table.Columns.Add("Scenario", GetType(String))
        Table.Columns.Add("Max Gearing", GetType(String))
        Table.Columns.Add("Gearing Year", GetType(String))
        Table.Columns.Add("Min Op Margin", GetType(String))
        Table.Columns.Add("Op Margin Year", GetType(String))
        Table.Columns.Add("Max Debt / Unit", GetType(String))
        Table.Columns.Add("Debt / Unit Year", GetType(String))
        Dim Sheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Comparative 2")
        For ResultIndex As Integer = 0 To 5
            Dim SourceRow As Integer = 34 + (ResultIndex * 2)
            Table.Rows.Add(
                Sheet.Cells(SourceRow, 15).DisplayText,
                Sheet.Cells(SourceRow, 17).DisplayText,
                Sheet.Cells(SourceRow, 18).DisplayText,
                Sheet.Cells(SourceRow, 20).DisplayText,
                Sheet.Cells(SourceRow, 21).DisplayText,
                Sheet.Cells(SourceRow, 23).DisplayText,
                Sheet.Cells(SourceRow, 24).DisplayText)
        Next
        Return Table

    End Function


    'Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click

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

