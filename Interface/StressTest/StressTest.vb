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
Imports DevExpress.XtraGrid.Views.BandedGrid
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
    Private StandardYEmptyEdit As RepositoryItemComboBox
    Private FirstTabOrdinalYearsEdit As RepositoryItemComboBox
    Private FirstTabOrdinalYearsLess1Edit As RepositoryItemComboBox
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
    Private NativePlannerView As BandedGridView
    Private NativePlannerData As System.Data.DataTable
    Private NativeYearEditor As RepositoryItemComboBox
    Private ReadOnly NativePlannerBandEditors As New List(Of NativePlannerBandEditorState)
    Private NativePlannerBandActiveEditor As DevExpress.XtraEditors.BaseEdit
    Private NativePlannerBandActiveState As NativePlannerBandEditorState
    Private NativePlannerBandActiveKind As NativePlannerBandEditorKind
    Private ClosingNativePlannerBandEditor As Boolean
    Private NativePlannerTabs As DevExpress.XtraTab.XtraTabControl
    Private NativeTargetsGrid As GridControl
    Private NativeTargetsView As BandedGridView
    Private NativeTargetsData As System.Data.DataTable
    Private ReadOnly NativeTargetRowEditors As New Dictionary(Of Integer, RepositoryItem)
    Private NativeDashboardHost As Panel
    Private NativeDashboardSpreadsheet As SpreadsheetControl
    Private NativeDashboardSpreadsheetOriginalParent As Control
    Private NativeDashboardSpreadsheetOriginalDock As DockStyle
    Private NativeDashboardSpreadsheetOriginalIndex As Integer = -1
    Private NativeDashboardSpreadsheetChangeHandlerAttached As Boolean
    Private ReplayingNativeDashboardCellChange As Boolean
    Private NativeSensitivityGrid As GridControl
    Private NativeSensitivityView As BandedGridView
    Private CovenantSummaryPanel As TableLayoutPanel
    Private NativeComparativeSelectors As New List(Of DevExpress.XtraEditors.ComboBoxEdit)
    Private NativeComparativeChartsA As TableLayoutPanel
    Private NativeComparativeChartsB As TableLayoutPanel
    Private NativeComparativeSummaryA As GridControl
    Private NativeComparativeSummaryB As GridControl
    Private LoadingNativeViews As Boolean

    Private Enum NativePlannerBandEditorKind
        ImportMode
        ScenarioName
        CopySource
    End Enum

    Private Class NativePlannerCopySource
        Public ScenarioIndex As Integer
        Public ScenarioName As String

        Public Overrides Function ToString() As String
            Return "Test " & ScenarioIndex.ToString() & " - " & ScenarioName
        End Function
    End Class

    Private Class NativePlannerBandEditorState
        Public Band As GridBand
        Public ScenarioIndex As Integer
        Public ImportModeEditor As RepositoryItemComboBox
        Public ScenarioNameEditor As RepositoryItemTextEdit
        Public CopySourceEditor As RepositoryItemComboBox
        Public GoButtonEditor As RepositoryItemButtonEdit
        Public ImportModeValue As Object
        Public ScenarioNameValue As Object
        Public CopySourceValue As NativePlannerCopySource
        Public ImportModeBounds As Rectangle = Rectangle.Empty
        Public ScenarioNameBounds As Rectangle = Rectangle.Empty
        Public CopySourceLabelBounds As Rectangle = Rectangle.Empty
        Public CopySourceBounds As Rectangle = Rectangle.Empty
        Public GoButtonBounds As Rectangle = Rectangle.Empty
    End Class

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
    Private ChangeMan As ModelChangeManager
    Private WrapCG_Mits As CustomGridWrapper
    Private View_WrapCG_Mits As CustomGridView
    Private ReadOnly FirstTabGridSources As New Dictionary(Of GridView, DevExpress.Spreadsheet.CellRange)
    Private ReadOnly FirstTabGridData As New Dictionary(Of GridView, System.Data.DataTable)
    Private FirstTabChangeInProgress As Boolean
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
        ChangeMan = ExcelModels(SetModelID).ChangeManager

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
        For Each Button As DevExpress.XtraEditors.ButtonPanel.IBaseButton In
            WindowsUIButtonPanelStressNavigator.Buttons
            If TypeOf Button Is DevExpress.XtraBars.Docking2010.WindowsUIButton Then
                DirectCast(
                    Button,
                    DevExpress.XtraBars.Docking2010.WindowsUIButton).IsLeft = False
            ElseIf TypeOf Button Is
                DevExpress.XtraBars.Docking2010.WindowsUISeparator Then
                DirectCast(
                    Button,
                    DevExpress.XtraBars.Docking2010.WindowsUISeparator).IsLeft = False
            End If
        Next
        Dim CentredButtons As DevExpress.XtraEditors.ButtonPanel.IBaseButton() =
            WindowsUIButtonPanelStressNavigator.Buttons.
                Cast(Of DevExpress.XtraEditors.ButtonPanel.IBaseButton)().
                Reverse().
                ToArray()
        WindowsUIButtonPanelStressNavigator.Buttons.Clear()
        WindowsUIButtonPanelStressNavigator.Buttons.AddRange(CentredButtons)
        WindowsUIButtonPanelStressNavigator.ContentAlignment =
            ContentAlignment.MiddleCenter

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

        StandardYEmptyEdit = New RepositoryItemComboBox
        StandardYEmptyEdit.Appearance.ForeColor = Color.White
        StandardYEmptyEdit.Appearance.Options.UseForeColor = True
        StandardYEmptyEdit.Appearance.BackColor = AbovoBlue
        StandardYEmptyEdit.Appearance.Options.UseBackColor = True
        StandardYEmptyEdit.Items.Add("Y")
        StandardYEmptyEdit.Items.Add("")
        StandardYEmptyEdit.TextEditStyle =
            DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor

        FirstTabOrdinalYearsEdit =
            RepositaryItems.GetEditor("Rep_OrdinalYears", ModelID).RetCombo
        FirstTabOrdinalYearsLess1Edit =
            RepositaryItems.GetEditor("Rep_OrdinalYearsLess1", ModelID).RetCombo

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

        RestoreNativeDashboardSpreadsheet()

    End Sub

    Private Sub DeactivateST(sender As Object, e As EventArgs)


    End Sub

    Private Sub AddHandlers()

        AddHandler TextEditMultivariableName.Validated, AddressOf TextEditMultivariableName_EditValueChanged
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

    Private Sub ObsoleteCustomDrawStressesGrid(sender As Object, e As RowCellCustomDrawEventArgs)

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


    Private Sub ObsoleteGVStressesCustEditor(sender As Object, e As CustomRowCellEditEventArgs)

        Dim view As CustomGridView = TryCast(sender, CustomGridView)

        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If TryGetFirstTabSourceCell(view, e.RowHandle, e.Column, SourceCell) AndAlso
           SourceCell.RowIndex = 34 AndAlso SourceCell.ColumnIndex = 3 Then
            e.RepositoryItem = StandardYEmptyEdit
            Return
        End If

        If e.Column.FieldName = "Column 3" Then

            Dim TestVal As String = view.GetRowCellValue(e.RowHandle, "Column 2")
            If TestVal = "delay by 1 year" Then
                e.RepositoryItem = StandardYEmptyEdit
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

    Private Sub ObsoleteGVStressesShowingEditor(sender As Object, e As CancelEventArgs)

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
                e.Cancel = True
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
                e.Cancel = True
                Return
            End If

        End If
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

        Dim MitigationStart As DevExpress.Spreadsheet.CellRange =
            worksheet.Range(
                ExcelModels(ModelID).WBStructure.StressTestDefinition.StresstestMitigationsRange)
        Dim MitigationEnd As DevExpress.Spreadsheet.CellRange =
            worksheet.Range(
                ExcelModels(ModelID).WBStructure.StressTestDefinition.StresstestMitigationsMoneyRange)
        Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range.FromLTRB(
            MitigationStart.LeftColumnIndex,
            MitigationStart.TopRowIndex,
            MitigationEnd.RightColumnIndex,
            MitigationEnd.BottomRowIndex)

        Dim RDSOptions As New RangeDataSourceOptions With {
            .UseFirstRowAsHeader = False,
            .PreserveFormulas = False,
            .SkipHiddenRows = True,
            .SkipHiddenColumns = True,
            .EditingOptions = DataSourceEditingOptions.AllowEdit
        }

        DSMitDataRange = range.GetDataSource(RDSOptions)

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
        ModelPostingComboBoxSelectCovenant.SuppressAutomaticPosting = True
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

        XtraTabPageMitigations.Controls.Clear()
        WrapCG_Mits = New CustomGridWrapper
        View_WrapCG_Mits = WrapCG_Mits.WrappedGridView

        Dim Sheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Live Multivariable Planner")
        Dim MitigationStart As DevExpress.Spreadsheet.CellRange =
            Sheet.Range(
                ExcelModels(ModelID).WBStructure.StressTestDefinition.StresstestMitigationsRange)
        Dim MitigationEnd As DevExpress.Spreadsheet.CellRange =
            Sheet.Range(
                ExcelModels(ModelID).WBStructure.StressTestDefinition.StresstestMitigationsMoneyRange)
        Dim SourceRange As DevExpress.Spreadsheet.CellRange = Sheet.Range.FromLTRB(
            MitigationStart.LeftColumnIndex,
            MitigationStart.TopRowIndex,
            MitigationEnd.RightColumnIndex,
            MitigationEnd.BottomRowIndex)
        Dim GridData As System.Data.DataTable = BuildFirstTabGridData(SourceRange)
        FirstTabGridData(View_WrapCG_Mits) = GridData
        WrapCG_Mits.WrappedCGC.DataSource = GridData
        RegisterFirstTabSourceGrid(
            View_WrapCG_Mits, SourceRange)

        WrapCG_Mits.Dock = DockStyle.Fill
        Me.XtraTabPageMitigations.Controls.Add(WrapCG_Mits)

        View_WrapCG_Mits.Columns(0).Caption = "Class"
        View_WrapCG_Mits.Columns(0).OptionsColumn.ReadOnly = True
        View_WrapCG_Mits.Columns(1).Caption = "Mitigation"
        View_WrapCG_Mits.Columns(1).OptionsColumn.ReadOnly = True
        View_WrapCG_Mits.Columns(2).Caption = "Type"
        View_WrapCG_Mits.Columns(2).OptionsColumn.ReadOnly = True
        View_WrapCG_Mits.Columns(3).Caption = "Change 1"
        View_WrapCG_Mits.Columns(4).Caption = "Change 2"
        View_WrapCG_Mits.Columns(5).Caption = "Change 1" & vbLf & "from year"
        View_WrapCG_Mits.Columns(6).Caption = "Change 2" & vbLf & "from year"
        View_WrapCG_Mits.Columns(7).Caption = "To year"
        For ColumnIndex As Integer = 3 To 7
            View_WrapCG_Mits.Columns(ColumnIndex).OptionsColumn.ReadOnly = False
            View_WrapCG_Mits.Columns(ColumnIndex).OptionsColumn.AllowEdit = True
            View_WrapCG_Mits.Columns(ColumnIndex).AppearanceHeader.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far
            View_WrapCG_Mits.Columns(ColumnIndex).AppearanceCell.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far
        Next

        WrapCG_Mits.WrappedCGC.RepositoryItems.Add(StandardYEmptyEdit)
        WrapCG_Mits.WrappedCGC.RepositoryItems.Add(FirstTabOrdinalYearsEdit)
        WrapCG_Mits.WrappedCGC.RepositoryItems.Add(FirstTabOrdinalYearsLess1Edit)
        WrapCG_Mits.WrappedCGC.RepositoryItems.Add(StandardPercentSpinEdit)
        WrapCG_Mits.WrappedCGC.RepositoryItems.Add(StandardIntegerTextBoxEdit)
        WrapCG_Mits.WrappedCGC.RepositoryItems.Add(Standard2digitnumberTextBoxEdit)
        Formatter.FormatGridView(View_WrapCG_Mits, WrapCG_Mits.WrappedCGC)
        View_WrapCG_Mits.BestFitColumns()

    End Sub
    Sub ProcessStressesGrid()

        Dim WrapCG_Stresses = New CustomGridWrapper

        View_WrapCG_Stresses = WrapCG_Stresses.WrappedGridView
        Dim SourceRange As DevExpress.Spreadsheet.CellRange =
            ActiveWorkbook.Worksheets("Live Multivariable Planner").Range(
                ExcelModels(ModelID).WBStructure.StressTestDefinition.StresstestStressesRange)
        Dim GridData As System.Data.DataTable = BuildFirstTabGridData(SourceRange)
        FirstTabGridData(View_WrapCG_Stresses) = GridData
        WrapCG_Stresses.WrappedCGC.DataSource = GridData
        RegisterFirstTabSourceGrid(
            View_WrapCG_Stresses,
            SourceRange)



        WrapCG_Stresses.Dock = DockStyle.Fill

        'GridView_InitialisationProcess_AddHandlers(View_WrapCG_BS)




        Me.XtraTabPageStresses.Controls.Add(WrapCG_Stresses)
        View_WrapCG_Stresses.Columns(0).Caption = "Class"
        View_WrapCG_Stresses.Columns(0).OptionsColumn.ReadOnly = True
        View_WrapCG_Stresses.Columns(1).Caption = "Stress"
        View_WrapCG_Stresses.Columns(1).OptionsColumn.ReadOnly = True
        View_WrapCG_Stresses.Columns(2).Caption = "Type"
        View_WrapCG_Stresses.Columns(2).OptionsColumn.ReadOnly = True
        View_WrapCG_Stresses.Columns(3).Caption = "Change 1"
        View_WrapCG_Stresses.Columns(4).Caption = "Change 2"
        View_WrapCG_Stresses.Columns(5).Caption = "Change 1" & vbLf & "from year"
        View_WrapCG_Stresses.Columns(6).Caption = "Change 2" & vbLf & "from year"
        View_WrapCG_Stresses.Columns(7).Caption = "To year"
        For ColumnIndex As Integer = 3 To 7
            View_WrapCG_Stresses.Columns(ColumnIndex).OptionsColumn.ReadOnly = False
            View_WrapCG_Stresses.Columns(ColumnIndex).OptionsColumn.AllowEdit = True
            View_WrapCG_Stresses.Columns(ColumnIndex).AppearanceHeader.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far
            View_WrapCG_Stresses.Columns(ColumnIndex).AppearanceCell.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far
        Next

        WrapCG_Stresses.WrappedCGC.RepositoryItems.Add(StandardYEmptyEdit)
        WrapCG_Stresses.WrappedCGC.RepositoryItems.Add(FirstTabOrdinalYearsEdit)
        WrapCG_Stresses.WrappedCGC.RepositoryItems.Add(FirstTabOrdinalYearsLess1Edit)
        WrapCG_Stresses.WrappedCGC.RepositoryItems.Add(StandardPercentSpinEdit)
        WrapCG_Stresses.WrappedCGC.RepositoryItems.Add(StandardIntegerTextBoxEdit)
        WrapCG_Stresses.WrappedCGC.RepositoryItems.Add(Standard2digitnumberTextBoxEdit)

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
                ExcelModels(ModelID).IsDirty = True

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
                ExcelModels(ModelID).IsDirty = True

            End If

            'SimpleButtonModeSwitch.Text = "Activate Stress Test Mode"
            PanelControlCovSel.Visible = False
            GridControlBreaches.Visible = False

        End If

        CalculateStressWorkbook()
        ProcessBreachesGrid(STMode = "Y")
        RefreshCovenantSummary()
        BuildCovCharts()

        Me.Cursor = Cursors.Default

    End Sub

    Sub StressTestModeAdjustments()

        UnhideColumnsCommand()
        Dim ReportingSheet As DevExpress.Spreadsheet.Worksheet =
            ExcelModels(ModelID).WB.Worksheets("OW - Live Stress Reporting")
        Dim WasProtected As Boolean = ReportingSheet.IsProtected
        UNProtectWS(ModelID, ReportingSheet.Name)

        Try
            Dim SourceRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.Range("StressSwitch")
            Dim DestRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.Range("StressBase")
            DestRange.CopyFrom(SourceRange, PasteSpecial.Values)
        Finally
            If WasProtected Then ProtectWS(ModelID, ReportingSheet.Name)
        End Try


    End Sub

    Sub Stress_Sensitivity_Capture()

        Dim WB As IWorkbook = ExcelModels(ModelID).WB
        Dim SensitivitySheet As DevExpress.Spreadsheet.Worksheet =
            WB.Worksheets("Stress Sensitivity List")
        Dim WasProtected As Boolean = SensitivitySheet.IsProtected
        UNProtectWS(ModelID, SensitivitySheet.Name)
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

            ExcelModels(ModelID).IsDirty = True
            CalculateStressWorkbook()
        Finally
            If WasProtected Then ProtectWS(ModelID, SensitivitySheet.Name)
        End Try

    End Sub


    Sub DeStressAdjustments()

        ExcelModels(ModelID).WB.Worksheets("Live Multivariable Planner").Columns.Hide(
            BreachOutputFirstColumnIndex, BreachOutputColumnCount)
        ProcessBreachesGrid(False)

    End Sub

    Private Function BuildFirstTabGridData(
        SourceRange As DevExpress.Spreadsheet.CellRange) As System.Data.DataTable

        Dim Data As New System.Data.DataTable
        For ColumnIndex As Integer = 0 To SourceRange.ColumnCount - 1
            Data.Columns.Add(
                "Column " & ColumnIndex.ToString(),
                If(ColumnIndex <= 2, GetType(String), GetType(Object)))
        Next

        For RowIndex As Integer = 0 To SourceRange.RowCount - 1
            Dim Row As System.Data.DataRow = Data.NewRow()
            For ColumnIndex As Integer = 0 To SourceRange.ColumnCount - 1
                Dim Value As Object = CellToObject(SourceRange(RowIndex, ColumnIndex))
                Row(ColumnIndex) = If(Value Is Nothing OrElse Value Is DBNull.Value,
                                      DBNull.Value, Value)
            Next
            Data.Rows.Add(Row)
        Next
        Data.AcceptChanges()
        Return Data

    End Function

    Private Sub RefreshFirstTabGridData(View As GridView)

        Dim SourceRange As DevExpress.Spreadsheet.CellRange = Nothing
        Dim Data As System.Data.DataTable = Nothing
        If View Is Nothing OrElse
           Not FirstTabGridSources.TryGetValue(View, SourceRange) OrElse
           Not FirstTabGridData.TryGetValue(View, Data) Then Return

        Data.BeginLoadData()
        Try
            For RowIndex As Integer = 0 To Math.Min(
                    SourceRange.RowCount, Data.Rows.Count) - 1
                For ColumnIndex As Integer = 0 To Math.Min(
                        SourceRange.ColumnCount, Data.Columns.Count) - 1
                    Dim Value As Object =
                        CellToObject(SourceRange(RowIndex, ColumnIndex))
                    Data.Rows(RowIndex)(ColumnIndex) =
                        If(Value Is Nothing OrElse Value Is DBNull.Value,
                           DBNull.Value, Value)
                Next
            Next
            Data.AcceptChanges()
        Finally
            Data.EndLoadData()
        End Try
        View.RefreshData()

    End Sub

    Private Sub RegisterFirstTabSourceGrid(
        View As GridView,
        SourceRange As DevExpress.Spreadsheet.CellRange)

        If View Is Nothing OrElse SourceRange Is Nothing Then Return

        FirstTabGridSources(View) = SourceRange
        RemoveHandler View.ShowingEditor, AddressOf FirstTabGridShowingEditor
        RemoveHandler View.ShownEditor, AddressOf FirstTabGridShownEditor
        RemoveHandler View.CellValueChanged, AddressOf FirstTabGridCellValueChanged
        RemoveHandler View.CustomRowCellEdit, AddressOf FirstTabGridCustomRowCellEdit
        RemoveHandler View.CustomDrawCell, AddressOf FirstTabGridCustomDrawCell
        RemoveHandler View.CustomColumnDisplayText, AddressOf FirstTabGridCustomColumnDisplayText
        AddHandler View.ShowingEditor, AddressOf FirstTabGridShowingEditor
        AddHandler View.ShownEditor, AddressOf FirstTabGridShownEditor
        AddHandler View.CellValueChanged, AddressOf FirstTabGridCellValueChanged
        AddHandler View.CustomRowCellEdit, AddressOf FirstTabGridCustomRowCellEdit
        AddHandler View.CustomDrawCell, AddressOf FirstTabGridCustomDrawCell
        AddHandler View.CustomColumnDisplayText, AddressOf FirstTabGridCustomColumnDisplayText

    End Sub

    Private Function TryGetFirstTabSourceCell(
        View As GridView,
        RowHandle As Integer,
        Column As DevExpress.XtraGrid.Columns.GridColumn,
        ByRef SourceCell As DevExpress.Spreadsheet.Cell) As Boolean

        SourceCell = Nothing
        If View Is Nothing OrElse Column Is Nothing OrElse RowHandle < 0 Then Return False

        Dim SourceRange As DevExpress.Spreadsheet.CellRange = Nothing
        If Not FirstTabGridSources.TryGetValue(View, SourceRange) Then Return False

        Dim SourceRowOffset As Integer = View.GetDataSourceRowIndex(RowHandle)
        Dim SourceColumnOffset As Integer = Column.AbsoluteIndex
        If SourceRowOffset < 0 OrElse SourceRowOffset >= SourceRange.RowCount OrElse
           SourceColumnOffset < 0 OrElse SourceColumnOffset >= SourceRange.ColumnCount Then Return False

        SourceCell = SourceRange(SourceRowOffset, SourceColumnOffset)
        Return SourceCell IsNot Nothing

    End Function

    Private Sub FirstTabGridCustomRowCellEdit(
        sender As Object,
        e As CustomRowCellEditEventArgs)

        Dim View As GridView = TryCast(sender, GridView)
        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If View Is Nothing OrElse
           Not TryGetFirstTabSourceCell(View, e.RowHandle, e.Column, SourceCell) Then Return

        Select Case SourceCell.ColumnIndex
            Case 3, 4
                If SourceCell.ColumnIndex = 3 AndAlso
                   (SourceCell.RowIndex = 34 OrElse
                    SourceCell.RowIndex = 64 OrElse
                    SourceCell.RowIndex = 65) Then
                    e.RepositoryItem = StandardYEmptyEdit
                Else
                    Dim Format As String = If(SourceCell.NumberFormat, String.Empty)
                    If Format.Contains("%") Then
                        e.RepositoryItem = StandardPercentSpinEdit
                    ElseIf Format.Contains("#,##0") Then
                        e.RepositoryItem = StandardIntegerTextBoxEdit
                    Else
                        e.RepositoryItem = Standard2digitnumberTextBoxEdit
                    End If
                End If
            Case 5, 6, 7
                If (SourceCell.RowIndex >= 9 AndAlso SourceCell.RowIndex <= 17) OrElse
                   (SourceCell.RowIndex >= 51 AndAlso SourceCell.RowIndex <= 54) Then
                    e.RepositoryItem = FirstTabOrdinalYearsLess1Edit
                Else
                    e.RepositoryItem = FirstTabOrdinalYearsEdit
                End If
        End Select

    End Sub

    Private Sub FirstTabGridCustomDrawCell(
        sender As Object,
        e As RowCellCustomDrawEventArgs)

        Dim View As GridView = TryCast(sender, GridView)
        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If View Is Nothing OrElse
           Not TryGetFirstTabSourceCell(View, e.RowHandle, e.Column, SourceCell) Then Return

        If e.Column.AbsoluteIndex >= 3 AndAlso
           Not IsWorkbookLinkedGridCellEditable(SourceCell) Then
            e.Appearance.BackColor = Color.Lavender
            e.Appearance.ForeColor = Color.WhiteSmoke
            e.Appearance.Options.UseBackColor = True
            e.Appearance.Options.UseForeColor = True
        Else
            ApplyWorkbookCellAppearance(e.Appearance, SourceCell)
        End If

        If View.IsCellSelected(e.RowHandle, e.Column) Then
            e.Appearance.BackColor = Color.Beige
            e.Appearance.ForeColor = Color.Black
        End If

        e.DefaultDraw()
        e.Handled = True

    End Sub

    Private Sub FirstTabGridCustomColumnDisplayText(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs)

        If e.ListSourceRowIndex < 0 OrElse
           e.Column Is Nothing OrElse
           e.Column.AbsoluteIndex < 3 OrElse e.Column.AbsoluteIndex > 4 OrElse
           e.Value Is Nothing OrElse e.Value Is DBNull.Value OrElse
           TypeOf e.Value Is String Then Return

        Dim View As GridView = TryCast(sender, GridView)
        If View Is Nothing Then Return
        Dim RowHandle As Integer = View.GetRowHandle(e.ListSourceRowIndex)
        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If Not TryGetFirstTabSourceCell(View, RowHandle, e.Column, SourceCell) Then Return

        Dim NumericValue As Double
        If Not Double.TryParse(
                Convert.ToString(e.Value, Globalization.CultureInfo.CurrentCulture),
                Globalization.NumberStyles.Any,
                Globalization.CultureInfo.CurrentCulture,
                NumericValue) Then Return

        Dim Format As String = If(SourceCell.NumberFormat, String.Empty)
        If Format.Contains("%") Then
            e.DisplayText = NumericValue.ToString("P2")
        ElseIf Format.Contains("#,##0") Then
            e.DisplayText = NumericValue.ToString("N0")
        Else
            e.DisplayText = NumericValue.ToString("N2")
        End If

    End Sub

    Private Sub FirstTabGridShowingEditor(sender As Object, e As CancelEventArgs)

        Dim View As GridView = TryCast(sender, GridView)
        If View Is Nothing Then Return

        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If Not TryGetFirstTabSourceCell(
                View, View.FocusedRowHandle, View.FocusedColumn, SourceCell) OrElse
           Not IsWorkbookLinkedGridCellEditable(SourceCell) Then
            e.Cancel = True
        End If

    End Sub

    Private Sub FirstTabGridShownEditor(sender As Object, e As EventArgs)

        Dim View As GridView = TryCast(sender, GridView)
        If View Is Nothing Then Return

        Dim ComboEditor As DevExpress.XtraEditors.ComboBoxEdit =
            TryCast(View.ActiveEditor, DevExpress.XtraEditors.ComboBoxEdit)
        If ComboEditor IsNot Nothing Then ComboEditor.ShowPopup()

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
            Dim Target As DevExpress.Spreadsheet.Cell =
                ActiveWorkbook.Worksheets("Live Multivariable Planner").Range("AD3")(0, 0)
            If Not ProcessStressTestCellChange(
                    Target, SelectedCovenant, "S",
                    "Stress-test covenant selection updated") Then
                ModelPostingComboBoxSelectCovenant.EditValue = Target.DisplayText
                Return
            End If
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

    Private Sub DisableStressTestGridFilteringAndSorting()

        For Each Grid As GridControl In FindControls(Of GridControl)(Me)
            Dim View As GridView = TryCast(Grid.MainView, GridView)
            If View Is Nothing Then Continue For

            View.OptionsCustomization.AllowFilter = False
            View.OptionsCustomization.AllowSort = False
            View.OptionsCustomization.AllowGroup = False
            View.OptionsMenu.EnableColumnMenu = False
            View.OptionsView.ShowAutoFilterRow = False
            View.OptionsView.ShowButtonMode =
                DevExpress.XtraGrid.Views.Base.ShowButtonModeEnum.ShowForFocusedRow
            View.OptionsView.ShowFilterPanelMode =
                DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never
            View.ClearColumnsFilter()

            For Each Column As DevExpress.XtraGrid.Columns.GridColumn In
                View.Columns
                Column.OptionsColumn.AllowSort =
                    DevExpress.Utils.DefaultBoolean.False
                Column.OptionsFilter.AllowFilter = False
            Next
        Next

    End Sub

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
        Dim SensitivitySheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Stress Sensitivity List")
        Dim WasProtected As Boolean = SensitivitySheet.IsProtected
        UNProtectWS(ModelID, SensitivitySheet.Name)

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
            ExcelModels(ModelID).IsDirty = True
            CalculateStressWorkbook()
            RenderStressHeaderHTMLData()
        Finally
            If WasProtected Then ProtectWS(ModelID, SensitivitySheet.Name)
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
        RefreshNativeTargets()
        RefreshNativeSensitivityList()
        RefreshNativeDashboard()
        RefreshNativeComparativeViews()
        DisableStressTestGridFilteringAndSorting()

    End Sub

    Private Sub BuildNativeSensitivityPage()

        WebBrowserStressCaptureOutput.Visible = False
        NativeSensitivityGrid = New GridControl With {.Dock = DockStyle.Fill}
        NativeSensitivityView = New BandedGridView(NativeSensitivityGrid)
        NativeSensitivityGrid.MainView = NativeSensitivityView
        NativeSensitivityGrid.ViewCollection.Add(NativeSensitivityView)
        NativeSensitivityView.OptionsBehavior.Editable = False
        NativeSensitivityView.OptionsView.ShowGroupPanel = False
        NativeSensitivityView.OptionsView.ColumnAutoWidth = False
        NativeSensitivityView.OptionsView.ShowBands = True
        NativeSensitivityView.Appearance.HeaderPanel.TextOptions.WordWrap =
            WordWrap.Wrap
        NativeSensitivityView.Appearance.HeaderPanel.Options.UseTextOptions = True
        NativeSensitivityView.OptionsView.ColumnHeaderAutoHeight =
            DevExpress.Utils.DefaultBoolean.True
        NativeSensitivityView.ColumnPanelRowHeight = 46
        NativeSensitivityView.OptionsSelection.MultiSelect = True
        NativeSensitivityView.OptionsSelection.MultiSelectMode =
            DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.RowSelect
        AddHandler NativeSensitivityView.CustomDrawCell,
            AddressOf NativeSensitivityCustomDrawCell
        AddHandler NativeSensitivityView.CustomDrawColumnHeader,
            AddressOf NativeSensitivityCustomDrawColumnHeader
        AddHandler NativeSensitivityView.CalcRowHeight,
            AddressOf NativeSensitivityCalcRowHeight
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
        Data.Columns.Add("RowKind", GetType(String))
        Data.Columns.Add("IsReferenceRow", GetType(Boolean))
        For ColumnIndex As Integer = 0 To 60
            Data.Columns.Add("C" & ColumnIndex.ToString(), GetType(String))
        Next

        AddNativeSensitivityRow(Data, Sheet, 6, "Target", True)
        AddNativeSensitivityRow(Data, Sheet, 7, "GoldenRule", True)
        AddNativeSensitivityRow(Data, Sheet, -1, "Spacer", True)
        AddNativeSensitivityRow(Data, Sheet, 9, "Live", True)

        For RowIndex As Integer = DataRows.TopRowIndex To DataRows.BottomRowIndex
            If String.IsNullOrWhiteSpace(Sheet.Cells(RowIndex, 0).DisplayText) Then Continue For
            AddNativeSensitivityRow(Data, Sheet, RowIndex, "Capture", False)
        Next

        NativeSensitivityGrid.DataSource = Data
        NativeSensitivityGrid.ForceInitialize()
        NativeSensitivityView.PopulateColumns()
        If NativeSensitivityView.Columns("SourceRow") IsNot Nothing Then
            NativeSensitivityView.Columns("SourceRow").Visible = False
        End If
        If NativeSensitivityView.Columns("RowKind") IsNot Nothing Then
            NativeSensitivityView.Columns("RowKind").Visible = False
        End If
        If NativeSensitivityView.Columns("IsReferenceRow") IsNot Nothing Then
            NativeSensitivityView.Columns("IsReferenceRow").Visible = False
        End If

        ConfigureNativeSensitivityBands(Sheet)

    End Sub

    Private Sub AddNativeSensitivityRow(
        Data As System.Data.DataTable,
        Sheet As DevExpress.Spreadsheet.Worksheet,
        RowIndex As Integer,
        RowKind As String,
        IsReferenceRow As Boolean)

        Dim Row As System.Data.DataRow = Data.NewRow()
        Row("SourceRow") = RowIndex
        Row("RowKind") = RowKind
        Row("IsReferenceRow") = IsReferenceRow
        If RowIndex >= 0 Then
            For ColumnIndex As Integer = 0 To 60
                Row("C" & ColumnIndex.ToString()) =
                    Sheet.Cells(RowIndex, ColumnIndex).DisplayText
            Next
        End If
        Data.Rows.Add(Row)

    End Sub

    Private Sub ConfigureNativeSensitivityBands(
        Sheet As DevExpress.Spreadsheet.Worksheet)

        NativeSensitivityView.Bands.Clear()

        Dim SummaryBand As New GridBand With {.Caption = "Summary"}
        NativeSensitivityView.Bands.Add(SummaryBand)
        ApplyWorkbookCellAppearance(SummaryBand.AppearanceHeader, Sheet.Cells(4, 0))
        For ColumnIndex As Integer = 0 To 8
            ConfigureNativeSensitivityColumn(
                SummaryBand, Sheet, ColumnIndex,
                Sheet.Cells(4, ColumnIndex).DisplayText,
                If(ColumnIndex = 0, 200,
                   If(ColumnIndex = 1, 105,
                      If(ColumnIndex <= 3, 95, 115))),
                ColumnIndex <= 3)
        Next

        For GroupIndex As Integer = 0 To 4
            Dim FirstColumn As Integer = 9 + (GroupIndex * 10)
            Dim MetricBand As New GridBand With {
                .Caption = Sheet.Cells(4, FirstColumn).DisplayText
            }
            NativeSensitivityView.Bands.Add(MetricBand)
            ApplyWorkbookCellAppearance(
                MetricBand.AppearanceHeader, Sheet.Cells(4, FirstColumn))
            For ColumnIndex As Integer = FirstColumn To FirstColumn + 9
                Dim Caption As String = Sheet.Cells(5, ColumnIndex).DisplayText
                If String.IsNullOrWhiteSpace(Caption) Then
                    Caption = (ColumnIndex - FirstColumn + 1).ToString()
                End If
                ConfigureNativeSensitivityColumn(
                    MetricBand, Sheet, ColumnIndex, Caption, 64, True)
            Next
        Next

        Dim CaptureBand As New GridBand With {.Caption = "Capture details"}
        NativeSensitivityView.Bands.Add(CaptureBand)
        ApplyWorkbookCellAppearance(CaptureBand.AppearanceHeader, Sheet.Cells(4, 59))
        ConfigureNativeSensitivityColumn(
            CaptureBand, Sheet, 59, Sheet.Cells(4, 59).DisplayText, 130, False)
        ConfigureNativeSensitivityColumn(
            CaptureBand, Sheet, 60, Sheet.Cells(4, 60).DisplayText, 280, False)

    End Sub

    Private Sub ConfigureNativeSensitivityColumn(
        Band As GridBand,
        Sheet As DevExpress.Spreadsheet.Worksheet,
        SourceColumn As Integer,
        Caption As String,
        Width As Integer,
        RightAlign As Boolean)

        Dim Column As BandedGridColumn =
            TryCast(NativeSensitivityView.Columns("C" & SourceColumn.ToString()),
                    BandedGridColumn)
        If Column Is Nothing Then Return
        Column.Caption = Caption.Trim()
        Column.Tag = SourceColumn
        Column.Width = Width
        Column.MinWidth = Math.Min(Width, 50)
        Column.OptionsColumn.ReadOnly = True
        Column.OptionsColumn.AllowEdit = False
        If RightAlign Then
            Column.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Far
            Column.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center
        End If
        Column.AppearanceHeader.TextOptions.WordWrap = WordWrap.Wrap
        Column.AppearanceHeader.TextOptions.VAlignment = VertAlignment.Bottom
        Column.AppearanceHeader.Options.UseTextOptions = True
        Band.Columns.Add(Column)

    End Sub

    Private Sub NativeSensitivityCustomDrawCell(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs)

        If e.RowHandle < 0 OrElse e.Column Is Nothing OrElse
           Not TypeOf e.Column.Tag Is Integer Then Return

        Dim RowKindValue As Object =
            NativeSensitivityView.GetRowCellValue(e.RowHandle, "RowKind")
        Dim RowKind As String =
            If(RowKindValue Is Nothing OrElse RowKindValue Is DBNull.Value,
               String.Empty, RowKindValue.ToString())
        If String.Equals(RowKind, "Spacer", StringComparison.Ordinal) Then
            e.Appearance.BackColor = Color.White
            e.Appearance.ForeColor = Color.White
            e.Appearance.Options.UseBackColor = True
            e.Appearance.Options.UseForeColor = True
            e.DisplayText = String.Empty
            e.DefaultDraw()
            e.Handled = True
            Return
        End If

        Dim SourceRowValue As Object =
            NativeSensitivityView.GetRowCellValue(e.RowHandle, "SourceRow")
        If SourceRowValue Is Nothing OrElse SourceRowValue Is DBNull.Value Then Return

        Dim SourceCell As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.Worksheets("Stress Sensitivity List").Cells(
                Convert.ToInt32(SourceRowValue), CInt(e.Column.Tag))

        ApplySensitivityCellAppearance(e.Appearance, SourceCell)

        e.DefaultDraw()
        e.Handled = True

    End Sub

    Private Sub ApplySensitivityCellAppearance(
        Appearance As DevExpress.Utils.AppearanceObject,
        SourceCell As DevExpress.Spreadsheet.Cell)

        If SourceCell Is Nothing Then Return

        'The sensitivity sheet's conditional formats resolve red/amber/green
        'against the Target and Golden Rule rows. FillColor is the resolved
        'workbook colour used by the original HTML renderer; BackgroundColor
        'would return only the unconditioned white base fill.
        Dim Background As Color = SourceCell.FillColor
        If Background.IsEmpty OrElse Background.A = 0 Then Background = Color.White
        Dim Foreground As Color = SourceCell.Font.Color
        If Foreground.IsEmpty OrElse Foreground.A = 0 Then Foreground = AbovoBlue

        Appearance.BackColor = Background
        Appearance.ForeColor = Foreground
        Appearance.Options.UseBackColor = True
        Appearance.Options.UseForeColor = True
        If SourceCell.Font.Bold Then
            Appearance.Font = New Font(
                Appearance.Font, Appearance.Font.Style Or FontStyle.Bold)
            Appearance.Options.UseFont = True
        End If

    End Sub

    Private Sub NativeSensitivityCalcRowHeight(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Grid.RowHeightEventArgs)

        If e.RowHandle < 0 Then Return
        Dim RowKindValue As Object =
            NativeSensitivityView.GetRowCellValue(e.RowHandle, "RowKind")
        Dim RowKind As String =
            If(RowKindValue Is Nothing OrElse RowKindValue Is DBNull.Value,
               String.Empty, RowKindValue.ToString())
        If String.Equals(RowKind, "Spacer", StringComparison.Ordinal) Then
            e.RowHeight = 12
        ElseIf String.Equals(RowKind, "Target", StringComparison.Ordinal) OrElse
               String.Equals(RowKind, "GoldenRule", StringComparison.Ordinal) Then
            e.RowHeight = 26
        Else
            e.RowHeight = 22
        End If

    End Sub

    Private Sub NativeSensitivityCustomDrawColumnHeader(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Grid.ColumnHeaderCustomDrawEventArgs)

        If e.Column Is Nothing OrElse Not TypeOf e.Column.Tag Is Integer Then Return

        Dim SourceColumn As Integer = CInt(e.Column.Tag)
        Dim HeaderRow As Integer = If(SourceColumn >= 9 AndAlso SourceColumn <= 58, 5, 4)
        Dim HeaderCell As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.Worksheets("Stress Sensitivity List").Cells(
                HeaderRow, SourceColumn)
        ApplySensitivityCellAppearance(e.Appearance, HeaderCell)
        'Workbook year headings resolve to a very pale foreground. Use a stable
        'high-contrast presentation for column headers; data-cell colours remain
        'entirely workbook-driven.
        e.Appearance.BackColor = Color.FromArgb(235, 235, 235)
        e.Appearance.BackColor2 = e.Appearance.BackColor
        e.Appearance.Options.UseBackColor = True
        e.Appearance.ForeColor = AbovoBlue
        e.Appearance.Options.UseForeColor = True
        e.Appearance.Font = New Font(
            e.Appearance.Font, e.Appearance.Font.Style Or FontStyle.Bold)
        e.Appearance.Options.UseFont = True
        e.Appearance.TextOptions.HAlignment = HorzAlignment.Center
        e.Appearance.TextOptions.VAlignment = VertAlignment.Bottom
        e.Appearance.TextOptions.WordWrap = WordWrap.Wrap
        e.Appearance.Options.UseTextOptions = True
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

        SelectedHandles = SelectedHandles.Where(
            Function(RowHandle)
                Dim IsReference As Object =
                    NativeSensitivityView.GetRowCellValue(RowHandle, "IsReferenceRow")
                Return IsReference Is Nothing OrElse IsReference Is DBNull.Value OrElse
                    Not Convert.ToBoolean(IsReference)
            End Function).ToArray()
        If SelectedHandles.Length = 0 Then
            DevExpress.XtraEditors.XtraMessageBox.Show(
                "Rows 7 and 8 are workbook reference rows and cannot be deleted.",
                "Delete captures")
            Return
        End If
        If DevExpress.XtraEditors.XtraMessageBox.Show(
                "Delete the selected captured stress-test records?",
                "Delete captures", MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning) <> DialogResult.Yes Then Return

        Dim SensitivitySheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Stress Sensitivity List")
        Dim WasProtected As Boolean = SensitivitySheet.IsProtected
        UNProtectWS(ModelID, SensitivitySheet.Name)
        Try
            For Each RowHandle As Integer In SelectedHandles
                Dim SourceRow As Integer =
                    Convert.ToInt32(NativeSensitivityView.GetRowCellValue(RowHandle, "SourceRow"))
                SensitivitySheet.Range.FromLTRB(0, SourceRow, 60, SourceRow).ClearContents()
            Next
            ExcelModels(ModelID).IsDirty = True
            CalculateStressWorkbook()
            RefreshNativeSensitivityList()
        Finally
            If WasProtected Then ProtectWS(ModelID, SensitivitySheet.Name)
        End Try

    End Sub

    Private Sub BuildNativePlannerPage()

        XtraTabPageMVP.Controls.Clear()
        NativePlannerTabs = New DevExpress.XtraTab.XtraTabControl With {
            .Dock = DockStyle.Fill
        }
        Dim TestsPage As New DevExpress.XtraTab.XtraTabPage With {
            .Text = "Tests"
        }
        Dim TargetsPage As New DevExpress.XtraTab.XtraTabPage With {
            .Text = "Targets and Golden Rules"
        }
        NativePlannerTabs.TabPages.AddRange(
            New DevExpress.XtraTab.XtraTabPage() {TestsPage, TargetsPage})
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
        NativePlannerScenario.Properties.Items.AddRange(
            DefaultScenarioNames().Skip(1).Cast(Of Object).ToArray())
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
            .Text = "Clear selected scenario", .Width = 165, .Height = 36
        }
        Toolbar.Controls.Add(CalculateButton)
        Toolbar.Controls.Add(GenerateButton)
        Toolbar.Controls.Add(ClearButton)

        NativePlannerGrid = New GridControl With {.Dock = DockStyle.Fill}
        NativePlannerView = New BandedGridView(NativePlannerGrid)
        NativePlannerGrid.MainView = NativePlannerView
        NativePlannerGrid.ViewCollection.Add(NativePlannerView)
        NativePlannerView.OptionsView.ShowGroupPanel = False
        NativePlannerView.OptionsView.ShowAutoFilterRow = False
        NativePlannerView.OptionsView.ColumnAutoWidth = False
        NativePlannerView.OptionsView.ShowBands = True
        NativePlannerView.OptionsBehavior.EditorShowMode = DevExpress.Utils.EditorShowMode.MouseDownFocused

        NativeYearEditor = New RepositoryItemComboBox
        NativeYearEditor.Items.Add("")
        For YearNumber As Integer = 1 To 40
            NativeYearEditor.Items.Add(YearNumber)
        Next
        NativeYearEditor.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        NativePlannerGrid.RepositoryItems.Add(NativeYearEditor)
        NativePlannerGrid.RepositoryItems.Add(StandardPercentSpinEdit)
        NativePlannerGrid.RepositoryItems.Add(StandardYEmptyEdit)
        NativePlannerGrid.RepositoryItems.Add(Standard2digitnumberTextBoxEdit)

        Root.Controls.Add(Toolbar, 0, 0)
        Root.Controls.Add(NativePlannerGrid, 0, 1)
        TestsPage.Controls.Add(Root)
        BuildNativeTargetsPage(TargetsPage)
        XtraTabPageMVP.Controls.Add(NativePlannerTabs)

        AddHandler NativePlannerScenario.SelectedIndexChanged, AddressOf NativePlannerScenarioChanged
        AddHandler NativePlannerName.Validated, AddressOf NativePlannerNameChanged
        AddHandler NativePlannerImportMode.SelectedIndexChanged, AddressOf NativePlannerImportModeChanged
        AddHandler NativePlannerInclude.CheckedChanged, AddressOf NativePlannerIncludeChanged
        AddHandler NativePlannerView.CellValueChanged, AddressOf NativePlannerCellValueChanged
        AddHandler NativePlannerView.CustomRowCellEdit, AddressOf NativePlannerCustomRowCellEdit
        AddHandler NativePlannerView.CustomDrawCell, AddressOf NativePlannerCustomDrawCell
        AddHandler NativePlannerView.CustomDrawBandHeader,
            AddressOf NativePlannerCustomDrawBandHeader
        AddHandler NativePlannerView.ShowingEditor, AddressOf NativePlannerShowingEditor
        AddHandler NativePlannerView.CustomColumnDisplayText, AddressOf NativePlannerCustomColumnDisplayText
        AddHandler NativePlannerView.FocusedColumnChanged, AddressOf NativePlannerFocusedColumnChanged
        AddHandler NativePlannerView.MouseDown, AddressOf NativePlannerBandMouseDown
        AddHandler NativePlannerView.Layout, AddressOf NativePlannerBandLayout
        AddHandler CalculateButton.Click, Sub() RecalculateAndRefreshNativeViews()
        AddHandler GenerateButton.Click, AddressOf GenerateMultivariableDashboard_Click
        AddHandler ClearButton.Click, AddressOf ClearNativeScenario_Click

    End Sub

    Private Sub BuildNativeTargetsPage(
        Page As DevExpress.XtraTab.XtraTabPage)

        NativeTargetsGrid = New GridControl With {.Dock = DockStyle.Fill}
        NativeTargetsView = New BandedGridView(NativeTargetsGrid)
        NativeTargetsGrid.MainView = NativeTargetsView
        NativeTargetsGrid.ViewCollection.Add(NativeTargetsView)
        NativeTargetsView.OptionsView.ShowBands = True
        NativeTargetsView.OptionsView.ShowColumnHeaders = False
        NativeTargetsView.OptionsView.ShowGroupPanel = False
        NativeTargetsView.OptionsView.ShowAutoFilterRow = False
        NativeTargetsView.OptionsView.ColumnAutoWidth = False
        NativeTargetsView.OptionsView.ColumnHeaderAutoHeight =
            DevExpress.Utils.DefaultBoolean.False
        'Rows 6/7/9 form the two band levels; workbook row 8 is represented by
        'the first real grid row so its text and list editors remain interactive.
        NativeTargetsView.BandPanelRowHeight = 40
        NativeTargetsView.RowHeight = 30
        NativeTargetsView.OptionsView.ShowButtonMode =
            DevExpress.XtraGrid.Views.Base.ShowButtonModeEnum.ShowForFocusedRow
        NativeTargetsView.OptionsBehavior.EditorShowMode =
            DevExpress.Utils.EditorShowMode.MouseDownFocused
        NativeTargetsView.OptionsCustomization.AllowFilter = False
        NativeTargetsView.OptionsCustomization.AllowSort = False
        NativeTargetsView.OptionsCustomization.AllowGroup = False
        NativeTargetsView.OptionsMenu.EnableColumnMenu = False

        NativeTargetsGrid.RepositoryItems.Add(StandardPercentSpinEdit)
        NativeTargetsGrid.RepositoryItems.Add(Standard2digitnumberTextBoxEdit)
        Page.Controls.Add(NativeTargetsGrid)

        AddHandler NativeTargetsView.CellValueChanged,
            AddressOf NativeTargetsCellValueChanged
        AddHandler NativeTargetsView.CustomRowCellEdit,
            AddressOf NativeTargetsCustomRowCellEdit
        AddHandler NativeTargetsView.CustomColumnDisplayText,
            AddressOf NativeTargetsCustomColumnDisplayText
        AddHandler NativeTargetsView.RowCellStyle,
            AddressOf NativeTargetsRowCellStyle
        AddHandler NativeTargetsView.ShowingEditor,
            AddressOf NativeTargetsShowingEditor

    End Sub

    Private Sub RefreshNativeTargets()

        If NativeTargetsGrid Is Nothing Then Return

        LoadingNativeViews = True
        Try
            Dim Sheet As DevExpress.Spreadsheet.Worksheet =
                ActiveWorkbook.Worksheets("Multivariable Planner")
            Dim Data As New System.Data.DataTable
            Data.Columns.Add("SourceRow", GetType(Integer))
            Data.Columns.Add("RowKind", GetType(String))
            Data.Columns.Add("Year", GetType(Object))
            For Offset As Integer = 0 To 4
                Data.Columns.Add("Target" & Offset.ToString(), GetType(Object))
            Next
            Data.Columns.Add("Spacer", GetType(Object))
            For Offset As Integer = 0 To 4
                Data.Columns.Add("Golden" & Offset.ToString(), GetType(Object))
            Next

            For Each RowIndex As Integer In
                Enumerable.Repeat(7, 1).Concat(Enumerable.Range(9, 39))
                Dim Row As System.Data.DataRow = Data.NewRow()
                Row("SourceRow") = RowIndex
                Row("RowKind") = If(RowIndex = 7, "Editor", "Data")
                Row("Year") = CellToObject(Sheet.Cells(RowIndex, 59))
                For Offset As Integer = 0 To 4
                    Row("Target" & Offset.ToString()) =
                        CellToObject(Sheet.Cells(RowIndex, 60 + Offset))
                    Row("Golden" & Offset.ToString()) =
                        CellToObject(Sheet.Cells(RowIndex, 66 + Offset))
                Next
                Row("Spacer") = DBNull.Value
                Data.Rows.Add(Row)
            Next
            Data.AcceptChanges()

            NativeTargetsData = Data
            NativeTargetsGrid.DataSource = Data
            NativeTargetsGrid.ForceInitialize()
            ConfigureNativeTargetColumns(Sheet)
        Finally
            LoadingNativeViews = False
        End Try

    End Sub

    Private Sub ConfigureNativeTargetColumns(
        Sheet As DevExpress.Spreadsheet.Worksheet)

        ResetNativeTargetRowEditors()
        NativeTargetsView.PopulateColumns()
        NativeTargetsView.Bands.Clear()
        NativeTargetsView.Columns("SourceRow").Visible = False
        NativeTargetsView.Columns("RowKind").Visible = False

        Dim YearColumn As BandedGridColumn =
            TryCast(NativeTargetsView.Columns("Year"), BandedGridColumn)
        YearColumn.Visible = True
        YearColumn.Caption = Sheet.Cells(8, 59).DisplayText
        YearColumn.Width = 72
        YearColumn.OptionsColumn.AllowEdit = False
        Dim YearBand As New GridBand With {
            .Caption = Sheet.Cells(8, 59).DisplayText,
            .Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left
        }
        ApplyWorkbookCellAppearance(YearBand.AppearanceHeader, Sheet.Cells(8, 59))
        YearBand.Columns.Add(YearColumn)
        NativeTargetsView.Bands.Add(YearBand)

        Dim TargetBand As New GridBand With {
            .Caption = Sheet.Cells(5, 60).DisplayText
        }
        ApplyWorkbookCellAppearance(TargetBand.AppearanceHeader, Sheet.Cells(5, 60))
        TargetBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center
        NativeTargetsView.Bands.Add(TargetBand)

        For Offset As Integer = 0 To 4
            ConfigureNativeTargetMetric(
                Sheet, TargetBand, "Target" & Offset.ToString(),
                60 + Offset, False)
        Next

        Dim SpacerColumn As BandedGridColumn =
            TryCast(NativeTargetsView.Columns("Spacer"), BandedGridColumn)
        SpacerColumn.Visible = True
        SpacerColumn.Caption = String.Empty
        SpacerColumn.Width = 26
        SpacerColumn.OptionsColumn.AllowEdit = False
        Dim SpacerBand As New GridBand With {
            .Caption = String.Empty, .Width = 26, .MinWidth = 26
        }
        SpacerBand.Columns.Add(SpacerColumn)
        NativeTargetsView.Bands.Add(SpacerBand)

        Dim GoldenBand As New GridBand With {
            .Caption = Sheet.Cells(5, 66).DisplayText
        }
        ApplyWorkbookCellAppearance(GoldenBand.AppearanceHeader, Sheet.Cells(5, 66))
        GoldenBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center
        NativeTargetsView.Bands.Add(GoldenBand)

        For Offset As Integer = 0 To 4
            ConfigureNativeTargetMetric(
                Sheet, GoldenBand, "Golden" & Offset.ToString(),
                66 + Offset, True)
        Next

        DisableStressTestGridFilteringAndSorting()

    End Sub

    Private Sub ConfigureNativeTargetMetric(
        Sheet As DevExpress.Spreadsheet.Worksheet,
        ParentBand As GridBand,
        FieldName As String,
        SourceColumn As Integer,
        IsGoldenRule As Boolean)

        Dim Column As BandedGridColumn =
            TryCast(NativeTargetsView.Columns(FieldName), BandedGridColumn)
        Column.Visible = True
        Column.Caption = String.Empty
        Column.Width = 126
        'Row 8 is editable on both sides; the locked Target result rows are
        'rejected individually by NativeTargetsShowingEditor.
        Column.OptionsColumn.AllowEdit = True
        Column.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.False
        Column.OptionsFilter.AllowFilter = False

        Dim MetricCaption As String = Sheet.Cells(6, SourceColumn).DisplayText
        Dim UnitCaption As String = Sheet.Cells(8, SourceColumn).DisplayText
        If Not String.IsNullOrWhiteSpace(UnitCaption) Then
            MetricCaption &= Environment.NewLine & UnitCaption
        End If
        Dim MetricBand As New GridBand With {
            .Caption = MetricCaption
        }
        ApplyWorkbookCellAppearance(
            MetricBand.AppearanceHeader, Sheet.Cells(6, SourceColumn))
        MetricBand.AppearanceHeader.Options.UseTextOptions = True
        MetricBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center
        MetricBand.AppearanceHeader.TextOptions.VAlignment = VertAlignment.Bottom
        MetricBand.AppearanceHeader.TextOptions.WordWrap = WordWrap.Wrap
        MetricBand.Columns.Add(Column)
        ParentBand.Children.Add(MetricBand)

        Dim SourceCell As DevExpress.Spreadsheet.Cell =
            Sheet.Cells(7, SourceColumn)
        Dim Editor As RepositoryItem
        If IsGoldenRule Then
            Dim Combo As New RepositoryItemComboBox
            Combo.TextEditStyle =
                DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
            For Each Item As Object In WorkbookValidationListItems(SourceCell)
                Combo.Items.Add(Item)
            Next
            Editor = Combo
        Else
            Editor = New RepositoryItemTextEdit
        End If
        ConfigureNativePlannerBandEditorAppearance(Editor, SourceCell)
        NativeTargetsGrid.RepositoryItems.Add(Editor)

        NativeTargetRowEditors(SourceColumn) = Editor

    End Sub

    Private Sub ResetNativeTargetRowEditors()

        For Each Editor As RepositoryItem In NativeTargetRowEditors.Values
            NativeTargetsGrid.RepositoryItems.Remove(Editor)
            Editor.Dispose()
        Next
        NativeTargetRowEditors.Clear()

    End Sub

    Private Function TryGetNativeTargetSourceColumn(
        Column As DevExpress.XtraGrid.Columns.GridColumn,
        ByRef SourceColumn As Integer) As Boolean

        If Column Is Nothing Then Return False
        If Column.FieldName = "Year" Then
            SourceColumn = 59
            Return True
        End If
        If Column.FieldName.StartsWith("Target", StringComparison.Ordinal) Then
            SourceColumn = 60 + Integer.Parse(Column.FieldName.Substring(6))
            Return True
        End If
        If Column.FieldName.StartsWith("Golden", StringComparison.Ordinal) Then
            SourceColumn = 66 + Integer.Parse(Column.FieldName.Substring(6))
            Return True
        End If
        Return False

    End Function

    Private Function TryGetNativeTargetSourceCell(
        RowHandle As Integer,
        Column As DevExpress.XtraGrid.Columns.GridColumn,
        ByRef SourceCell As DevExpress.Spreadsheet.Cell) As Boolean

        If RowHandle < 0 Then Return False
        Dim SourceColumn As Integer
        If Not TryGetNativeTargetSourceColumn(Column, SourceColumn) Then Return False
        Dim SourceRowValue As Object =
            NativeTargetsView.GetRowCellValue(RowHandle, "SourceRow")
        If SourceRowValue Is Nothing OrElse SourceRowValue Is DBNull.Value Then Return False
        SourceCell = ActiveWorkbook.Worksheets("Multivariable Planner").Cells(
            Convert.ToInt32(SourceRowValue), SourceColumn)
        Return True

    End Function

    Private Sub NativeTargetsShowingEditor(
        sender As Object,
        e As CancelEventArgs)

        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If Not TryGetNativeTargetSourceCell(
                NativeTargetsView.FocusedRowHandle,
                NativeTargetsView.FocusedColumn,
                SourceCell) OrElse
           Not IsWorkbookLinkedGridCellEditable(SourceCell) Then
            e.Cancel = True
        End If

    End Sub

    Private Sub NativeTargetsCustomRowCellEdit(
        sender As Object,
        e As CustomRowCellEditEventArgs)

        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If Not TryGetNativeTargetSourceCell(
                e.RowHandle, e.Column, SourceCell) OrElse
           Not IsWorkbookLinkedGridCellEditable(SourceCell) Then Return
        If SourceCell.RowIndex = 7 Then
            Dim SourceColumn As Integer
            If TryGetNativeTargetSourceColumn(e.Column, SourceColumn) AndAlso
               NativeTargetRowEditors.ContainsKey(SourceColumn) Then
                e.RepositoryItem = NativeTargetRowEditors(SourceColumn)
            End If
            Return
        End If
        If SourceCell.NumberFormat.Contains("%") Then
            e.RepositoryItem = StandardPercentSpinEdit
        Else
            e.RepositoryItem = Standard2digitnumberTextBoxEdit
        End If

    End Sub

    Private Sub NativeTargetsCustomColumnDisplayText(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs)

        If e.ListSourceRowIndex < 0 Then Return
        Dim RowHandle As Integer =
            NativeTargetsView.GetRowHandle(e.ListSourceRowIndex)
        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If TryGetNativeTargetSourceCell(RowHandle, e.Column, SourceCell) Then
            e.DisplayText = SourceCell.DisplayText
        End If

    End Sub

    Private Sub NativeTargetsRowCellStyle(
        sender As Object,
        e As RowCellStyleEventArgs)

        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If TryGetNativeTargetSourceCell(e.RowHandle, e.Column, SourceCell) Then
            ApplyWorkbookResolvedCellAppearance(e.Appearance, SourceCell)
        End If

    End Sub

    Private Sub NativeTargetsCellValueChanged(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs)

        If LoadingNativeViews Then Return
        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If Not TryGetNativeTargetSourceCell(
                e.RowHandle, e.Column, SourceCell) OrElse
           Not IsWorkbookLinkedGridCellEditable(SourceCell) Then
            RefreshNativeTargets()
            Return
        End If

        Dim IsHeaderEditor As Boolean = SourceCell.RowIndex = 7
        Dim DataFormat As String =
            If(IsHeaderEditor, "S",
               If(SourceCell.NumberFormat.Contains("%"), "P", "N"))
        ProcessStressTestCellChange(
            SourceCell,
            NormalizeStressTestEditValue(e.Value),
            DataFormat,
            If(IsHeaderEditor,
               "Stress-test covenant header setting updated",
               "Stress-test golden rule updated"),
            True)
        RefreshNativeTargets()

    End Sub

    Private Sub BuildNativeDashboardPage()

        XtraTabPageDashboard.Controls.Clear()
        NativeDashboardHost = New Panel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .Padding = New Padding(8)
        }
        XtraTabPageDashboard.Controls.Add(NativeDashboardHost)

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
        Editor.Properties.ReadOnly =
            ActiveWorkbook.Worksheets("OW - Covenant Calculation").
                Cells(WorkingRow - 1, 2).Protection.Locked
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

    Private Function SelectedPlannerScenarioIndex() As Integer

        If NativePlannerScenario Is Nothing OrElse
           NativePlannerScenario.SelectedIndex < 0 Then Return 1

        'The editable planner contains Test 1 through Test 10. The hidden D:H
        'block is the dashboard base result (S0), not planner scenario zero.
        Return Math.Min(10, NativePlannerScenario.SelectedIndex + 1)

    End Function

    Private Sub RefreshNativePlanner()

        If NativePlannerScenario Is Nothing Then Return
        LoadingNativeViews = True
        Try
            Dim ScenarioIndex As Integer = SelectedPlannerScenarioIndex()
            Dim Sheet As DevExpress.Spreadsheet.Worksheet = ActiveWorkbook.Worksheets("Multivariable Planner")
            RefreshNativePlannerScenarioControls(Sheet, ScenarioIndex)

            Dim Data As New System.Data.DataTable
            Data.Columns.Add("SourceRow", GetType(Integer))
            Data.Columns.Add("Section", GetType(String))
            Data.Columns.Add("Assumption", GetType(String))
            Data.Columns.Add("ShortName", GetType(String))
            For PlannerScenarioIndex As Integer = 1 To 10
                For ValueOffset As Integer = 0 To 4
                    Data.Columns.Add(
                        PlannerScenarioFieldName(PlannerScenarioIndex, ValueOffset),
                        GetType(Object))
                Next
                Data.Columns.Add(
                    PlannerScenarioFormatFieldName(PlannerScenarioIndex),
                    GetType(String))
            Next
            AddPlannerRows(Data, Sheet, 9, 47, "Stresses")
            AddPlannerRows(Data, Sheet, 51, 75, "Mitigations")
            NativePlannerData = Data
            NativePlannerGrid.DataSource = Data
            ConfigureNativePlannerColumns()
        Finally
            LoadingNativeViews = False
        End Try

    End Sub

    Private Sub AddPlannerRows(Data As System.Data.DataTable,
                               Sheet As DevExpress.Spreadsheet.Worksheet,
                               FirstRow As Integer,
                               LastRow As Integer,
                               SectionName As String)

        For RowIndex As Integer = FirstRow To LastRow
            Dim Row As System.Data.DataRow = Data.NewRow()
            Row("SourceRow") = RowIndex
            Row("Section") = SectionName
            Row("Assumption") = Sheet.Cells(RowIndex, 0).DisplayText
            Row("ShortName") = Sheet.Cells(RowIndex, 1).DisplayText
            For ScenarioIndex As Integer = 1 To 10
                Dim StartColumn As Integer = ScenarioStartColumn(ScenarioIndex)
                For ValueOffset As Integer = 0 To 4
                    Row(PlannerScenarioFieldName(ScenarioIndex, ValueOffset)) =
                        CellToObject(Sheet.Cells(RowIndex, StartColumn + ValueOffset))
                Next
                Row(PlannerScenarioFormatFieldName(ScenarioIndex)) =
                    Sheet.Cells(RowIndex, StartColumn).NumberFormat
            Next
            Data.Rows.Add(Row)
        Next

    End Sub

    Private Sub RefreshNativePlannerScenarioControls(
        Sheet As DevExpress.Spreadsheet.Worksheet,
        ScenarioIndex As Integer)

        Dim StartColumn As Integer = ScenarioStartColumn(ScenarioIndex)
        NativePlannerName.Properties.ReadOnly = Sheet.Cells(7, StartColumn).Protection.Locked
        NativePlannerInclude.Properties.ReadOnly = Sheet.Cells(6, 3).Protection.Locked
        NativePlannerImportMode.Properties.ReadOnly = Sheet.Cells(6, StartColumn).Protection.Locked
        NativePlannerName.EditValue = Sheet.Cells(7, StartColumn).DisplayText
        NativePlannerInclude.Visible = True
        NativePlannerImportMode.Visible = True
        NativePlannerInclude.Checked =
            String.Equals(Sheet.Cells(6, 3).DisplayText, "Yes", StringComparison.OrdinalIgnoreCase)
        NativePlannerImportMode.EditValue = Sheet.Cells(6, StartColumn).DisplayText
        If String.IsNullOrWhiteSpace(Convert.ToString(NativePlannerImportMode.EditValue)) Then
            NativePlannerImportMode.EditValue = "Use assumptions below"
        End If

    End Sub

    Private Function PlannerScenarioFieldName(
        ScenarioIndex As Integer,
        ValueOffset As Integer) As String

        Dim Suffix As String
        Select Case ValueOffset
            Case 0 : Suffix = "Change1"
            Case 1 : Suffix = "Change2"
            Case 2 : Suffix = "Change1FromYear"
            Case 3 : Suffix = "Change2FromYear"
            Case 4 : Suffix = "ToYear"
            Case Else
                Throw New ArgumentOutOfRangeException(NameOf(ValueOffset))
        End Select
        Return "S" & ScenarioIndex.ToString() & Suffix

    End Function

    Private Function PlannerScenarioFormatFieldName(ScenarioIndex As Integer) As String

        Return "S" & ScenarioIndex.ToString() & "ValueFormat"

    End Function

    Private Function TryGetPlannerScenarioColumn(
        Column As DevExpress.XtraGrid.Columns.GridColumn,
        ByRef ScenarioIndex As Integer,
        ByRef ValueOffset As Integer) As Boolean

        If Column Is Nothing Then Return False
        For CandidateScenario As Integer = 1 To 10
            For CandidateOffset As Integer = 0 To 4
                If String.Equals(
                        Column.FieldName,
                        PlannerScenarioFieldName(CandidateScenario, CandidateOffset),
                        StringComparison.Ordinal) Then
                    ScenarioIndex = CandidateScenario
                    ValueOffset = CandidateOffset
                    Return True
                End If
            Next
        Next
        Return False

    End Function

    Private Sub ConfigureNativePlannerColumns()

        NativePlannerView.PopulateColumns()
        If NativePlannerView.Columns.Count = 0 Then Return
        ResetNativePlannerBandEditors()
        NativePlannerView.Bands.Clear()
        NativePlannerView.Columns("SourceRow").Visible = False
        NativePlannerView.Columns("Section").OptionsColumn.AllowEdit = False
        NativePlannerView.Columns("Assumption").OptionsColumn.AllowEdit = False
        NativePlannerView.Columns("ShortName").OptionsColumn.AllowEdit = True
        NativePlannerView.Columns("Section").GroupIndex = 0
        NativePlannerView.Columns("Section").SortOrder = DevExpress.Data.ColumnSortOrder.Descending
        NativePlannerView.Columns("Assumption").Width = 330
        NativePlannerView.Columns("ShortName").Width = 180

        Dim DefinitionBand As New GridBand With {
            .Caption = "Stress and mitigation definition",
            .Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left,
            .RowCount = 5
        }
        DefinitionBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center
        DefinitionBand.Columns.Add(NativePlannerView.Columns("Assumption"))
        DefinitionBand.Columns.Add(NativePlannerView.Columns("ShortName"))
        NativePlannerView.Bands.Add(DefinitionBand)

        Dim Sheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Multivariable Planner")
        For ScenarioIndex As Integer = 1 To 10
            Dim StartColumn As Integer = ScenarioStartColumn(ScenarioIndex)
            Dim ScenarioName As String = Sheet.Cells(7, StartColumn).DisplayText.Trim()
            If String.IsNullOrWhiteSpace(ScenarioName) Then
                ScenarioName = "Scenario " & ScenarioIndex.ToString()
            End If
            Dim ImportMode As String = Sheet.Cells(6, StartColumn).DisplayText.Trim()
            If String.IsNullOrWhiteSpace(ImportMode) Then ImportMode = "Use assumptions below"

            Dim ScenarioBand As New GridBand With {
                .Caption = "Test " & ScenarioIndex.ToString(),
                .RowCount = 5
            }
            ScenarioBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center
            ScenarioBand.AppearanceHeader.TextOptions.WordWrap = WordWrap.Wrap
            NativePlannerView.Bands.Add(ScenarioBand)

            For ValueOffset As Integer = 0 To 4
                Dim Column = NativePlannerView.Columns(
                    PlannerScenarioFieldName(ScenarioIndex, ValueOffset))
                Select Case ValueOffset
                    Case 0
                        Column.Caption = "Change 1"
                        Column.Width = 105
                    Case 1
                        Column.Caption = "Change 2"
                        Column.Width = 105
                    Case 2
                        Column.Caption = "Change 1 from year"
                        Column.Width = 120
                        Column.ColumnEdit = NativeYearEditor
                    Case 3
                        Column.Caption = "Change 2 from year"
                        Column.Width = 120
                        Column.ColumnEdit = NativeYearEditor
                    Case 4
                        Column.Caption = "To year"
                        Column.Width = 100
                        Column.ColumnEdit = NativeYearEditor
                End Select
                Column.OptionsColumn.AllowEdit = True
                ScenarioBand.Columns.Add(Column)
            Next
            NativePlannerView.Columns(
                PlannerScenarioFormatFieldName(ScenarioIndex)).Visible = False
            AddNativePlannerBandEditors(
                ScenarioBand, Sheet, ScenarioIndex, ImportMode, ScenarioName)
        Next
        NativePlannerView.MinBandPanelRowCount = 5
        NativePlannerView.BandPanelRowHeight = 28
        NativePlannerView.ExpandAllGroups()

    End Sub

    Private Sub AddNativePlannerBandEditors(
        Band As GridBand,
        Sheet As DevExpress.Spreadsheet.Worksheet,
        ScenarioIndex As Integer,
        ImportMode As String,
        ScenarioName As String)

        Dim StartColumn As Integer = ScenarioStartColumn(ScenarioIndex)
        Dim ImportCell As DevExpress.Spreadsheet.Cell =
            Sheet.Cells(6, StartColumn)
        Dim NameCell As DevExpress.Spreadsheet.Cell =
            Sheet.Cells(7, StartColumn)

        Dim ImportEditor As New RepositoryItemComboBox
        ImportEditor.TextEditStyle =
            DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        For Each Item As Object In WorkbookValidationListItems(ImportCell)
            ImportEditor.Items.Add(Item)
        Next
        If ImportEditor.Items.Count = 0 AndAlso
           Not String.IsNullOrWhiteSpace(ImportMode) Then
            ImportEditor.Items.Add(ImportMode)
        End If
        ConfigureNativePlannerBandEditorAppearance(
            ImportEditor, ImportCell)

        Dim NameEditor As New RepositoryItemTextEdit
        ConfigureNativePlannerBandEditorAppearance(
            NameEditor, NameCell)

        Dim CopyEditor As New RepositoryItemComboBox
        CopyEditor.TextEditStyle =
            DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        For CopyScenarioIndex As Integer = 1 To 10
            If CopyScenarioIndex = ScenarioIndex Then Continue For
            Dim CopyName As String =
                Sheet.Cells(
                    7,
                    ScenarioStartColumn(CopyScenarioIndex)).DisplayText.Trim()
            If String.IsNullOrWhiteSpace(CopyName) Then
                CopyName = "Scenario " & CopyScenarioIndex.ToString()
            End If
            CopyEditor.Items.Add(
                New NativePlannerCopySource With {
                    .ScenarioIndex = CopyScenarioIndex,
                    .ScenarioName = CopyName
                })
        Next
        CopyEditor.Appearance.Options.UseTextOptions = True
        CopyEditor.Appearance.TextOptions.HAlignment = HorzAlignment.Center
        ConfigureStandaloneNativePlannerBandEditorAppearance(CopyEditor)

        Dim GoEditor As New RepositoryItemButtonEdit
        GoEditor.TextEditStyle =
            DevExpress.XtraEditors.Controls.TextEditStyles.HideTextEditor
        GoEditor.Buttons.Clear()
        GoEditor.Buttons.Add(
            New DevExpress.XtraEditors.Controls.EditorButton(
                DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph) With {
                    .Caption = "Go"
                })

        NativePlannerBandEditors.Add(
            New NativePlannerBandEditorState With {
                .Band = Band,
                .ScenarioIndex = ScenarioIndex,
                .ImportModeEditor = ImportEditor,
                .ScenarioNameEditor = NameEditor,
                .CopySourceEditor = CopyEditor,
                .GoButtonEditor = GoEditor,
                .ImportModeValue = ImportMode,
                .ScenarioNameValue = ScenarioName
            })

    End Sub

    Private Function WorkbookValidationListItems(
        SourceCell As DevExpress.Spreadsheet.Cell) As List(Of Object)

        Dim Result As New List(Of Object)
        If SourceCell Is Nothing Then Return Result

        Dim Validation As DevExpress.Spreadsheet.DataValidation =
            SourceCell.Worksheet.DataValidations.GetDataValidation(SourceCell)
        If Validation Is Nothing Then Return Result

        Dim Criteria As DevExpress.Spreadsheet.ValueObject =
            Validation.Criteria
        If Criteria Is Nothing Then Return Result

        If Criteria.IsText Then
            For Each Item As String In Criteria.TextValue.Split(
                    New Char() {","c, ";"c},
                    StringSplitOptions.RemoveEmptyEntries)
                Dim CleanItem As String = Item.Trim()
                If Not String.IsNullOrWhiteSpace(CleanItem) Then
                    Result.Add(CleanItem)
                End If
            Next
        ElseIf Criteria.IsRange Then
            Dim ValidationRange As DevExpress.Spreadsheet.CellRange =
                Criteria.RangeValue
            For RowIndex As Integer =
                    ValidationRange.TopRowIndex To ValidationRange.BottomRowIndex
                For ColumnIndex As Integer =
                        ValidationRange.LeftColumnIndex To ValidationRange.RightColumnIndex
                    Dim ItemText As String =
                        ValidationRange.Worksheet.Cells(
                            RowIndex, ColumnIndex).DisplayText.Trim()
                    If Not String.IsNullOrWhiteSpace(ItemText) Then
                        Result.Add(ItemText)
                    End If
                Next
            Next
        End If

        Return Result

    End Function

    Private Sub ConfigureNativePlannerBandEditorAppearance(
        Editor As RepositoryItem,
        SourceCell As DevExpress.Spreadsheet.Cell)

        ApplyWorkbookCellAppearance(Editor.Appearance, SourceCell)
        Dim ResolvedFill As Color = SourceCell.FillColor
        If Not ResolvedFill.IsEmpty AndAlso ResolvedFill.A > 0 Then
            Editor.Appearance.BackColor = ResolvedFill
            Editor.Appearance.Options.UseBackColor = True
        End If
        Editor.Appearance.Options.UseTextOptions = True
        Editor.Appearance.TextOptions.HAlignment = HorzAlignment.Center
        Editor.Appearance.TextOptions.VAlignment = VertAlignment.Center
        Editor.AppearanceDisabled.Assign(Editor.Appearance)
        Editor.AppearanceFocused.Assign(Editor.Appearance)
        Editor.AppearanceReadOnly.Assign(Editor.Appearance)

        Dim ComboEditor As RepositoryItemComboBox =
            TryCast(Editor, RepositoryItemComboBox)
        If ComboEditor IsNot Nothing Then
            ComboEditor.AppearanceDropDown.Assign(Editor.Appearance)
        End If

    End Sub

    Private Sub ConfigureStandaloneNativePlannerBandEditorAppearance(
        Editor As RepositoryItem)

        Editor.Appearance.BackColor = AbovoBlue
        Editor.Appearance.ForeColor = Color.White
        Editor.Appearance.Options.UseBackColor = True
        Editor.Appearance.Options.UseForeColor = True
        Editor.AppearanceDisabled.Assign(Editor.Appearance)
        Editor.AppearanceFocused.Assign(Editor.Appearance)
        Editor.AppearanceReadOnly.Assign(Editor.Appearance)

        Dim ComboEditor As RepositoryItemComboBox =
            TryCast(Editor, RepositoryItemComboBox)
        If ComboEditor IsNot Nothing Then
            ComboEditor.AppearanceDropDown.Assign(Editor.Appearance)
        End If

    End Sub

    Private Sub NativePlannerCustomDrawBandHeader(
        sender As Object,
        e As BandHeaderCustomDrawEventArgs)

        If e.Band Is Nothing Then Return
        Dim State As NativePlannerBandEditorState =
            NativePlannerBandEditors.FirstOrDefault(
                Function(Item) Item.Band Is e.Band)
        If State Is Nothing Then Return

        Dim SavedCaption As String = e.Info.Caption
        e.Info.Caption = String.Empty
        e.DefaultDraw()
        e.Info.Caption = SavedCaption

        Dim CaptionBounds As New Rectangle(
            e.Bounds.Left + 4, e.Bounds.Top + 2,
            Math.Max(1, e.Bounds.Width - 8), 20)
        Using CaptionFormat As New StringFormat With {
            .Alignment = StringAlignment.Center,
            .LineAlignment = StringAlignment.Center
        }
            e.Cache.DrawString(
                "Test " & State.ScenarioIndex.ToString(),
                e.Appearance.GetFont(),
                e.Appearance.GetForeBrush(e.Cache),
                CaptionBounds,
                CaptionFormat)
        End Using

        CalculateNativePlannerBandEditorBounds(State, e.Bounds)
        DrawEditorHelper.DrawEdit(
            e.Graphics, State.ImportModeEditor,
            State.ImportModeBounds, State.ImportModeValue, True)
        DrawEditorHelper.DrawEdit(
            e.Graphics, State.ScenarioNameEditor,
            State.ScenarioNameBounds, State.ScenarioNameValue, True)
        Using CopyLabelFormat As New StringFormat With {
            .Alignment = StringAlignment.Near,
            .LineAlignment = StringAlignment.Center
        }
            e.Cache.DrawString(
                "Copy From:",
                e.Appearance.GetFont(),
                e.Appearance.GetForeBrush(e.Cache),
                State.CopySourceLabelBounds,
                CopyLabelFormat)
        End Using
        DrawEditorHelper.DrawEdit(
            e.Graphics, State.CopySourceEditor,
            State.CopySourceBounds, State.CopySourceValue, True)
        DrawEditorHelper.DrawEdit(
            e.Graphics, State.GoButtonEditor,
            State.GoButtonBounds, Nothing)
        e.Handled = True

    End Sub

    Private Sub CalculateNativePlannerBandEditorBounds(
        State As NativePlannerBandEditorState,
        BandBounds As Rectangle)

        Const HorizontalPadding As Integer = 5
        Const CaptionHeight As Integer = 22
        Const EditorGap As Integer = 3
        Const EditorHeight As Integer = 27
        Const CopyTopGap As Integer = 8
        Const CopyLabelWidth As Integer = 74
        Const GoButtonWidth As Integer = 48
        Dim EditorWidth As Integer =
            Math.Max(1, BandBounds.Width - (2 * HorizontalPadding))
        Dim FirstTop As Integer = BandBounds.Top + CaptionHeight

        State.ImportModeBounds = New Rectangle(
            BandBounds.Left + HorizontalPadding,
            FirstTop,
            EditorWidth,
            EditorHeight)
        State.ScenarioNameBounds = New Rectangle(
            BandBounds.Left + HorizontalPadding,
            FirstTop + EditorHeight + EditorGap,
            EditorWidth,
            EditorHeight)
        Dim CopyTop As Integer =
            State.ScenarioNameBounds.Bottom + CopyTopGap
        State.CopySourceLabelBounds = New Rectangle(
            BandBounds.Left + HorizontalPadding,
            CopyTop,
            CopyLabelWidth,
            EditorHeight)
        State.GoButtonBounds = New Rectangle(
            BandBounds.Right - HorizontalPadding - GoButtonWidth,
            CopyTop,
            GoButtonWidth,
            EditorHeight)
        State.CopySourceBounds = New Rectangle(
            State.CopySourceLabelBounds.Right + EditorGap,
            CopyTop,
            Math.Max(
                1,
                State.GoButtonBounds.Left -
                    State.CopySourceLabelBounds.Right -
                    (2 * EditorGap)),
            EditorHeight)

    End Sub

    Private Sub NativePlannerBandMouseDown(
        sender As Object,
        e As MouseEventArgs)

        CloseNativePlannerBandEditor(True)

        For Each State As NativePlannerBandEditorState In
            NativePlannerBandEditors
            If State.ImportModeBounds.Contains(e.Location) Then
                ShowNativePlannerBandEditor(
                    State,
                    NativePlannerBandEditorKind.ImportMode,
                    State.ImportModeBounds)
                DXMouseEventArgs.GetMouseArgs(e).Handled = True
                Return
            End If
            If State.ScenarioNameBounds.Contains(e.Location) Then
                ShowNativePlannerBandEditor(
                    State,
                    NativePlannerBandEditorKind.ScenarioName,
                    State.ScenarioNameBounds)
                DXMouseEventArgs.GetMouseArgs(e).Handled = True
                Return
            End If
            If State.CopySourceBounds.Contains(e.Location) Then
                ShowNativePlannerBandEditor(
                    State,
                    NativePlannerBandEditorKind.CopySource,
                    State.CopySourceBounds)
                DXMouseEventArgs.GetMouseArgs(e).Handled = True
                Return
            End If
            If State.GoButtonBounds.Contains(e.Location) Then
                CopyNativePlannerScenario(State)
                DXMouseEventArgs.GetMouseArgs(e).Handled = True
                Return
            End If
        Next

    End Sub

    Private Sub ShowNativePlannerBandEditor(
        State As NativePlannerBandEditorState,
        Kind As NativePlannerBandEditorKind,
        Bounds As Rectangle)

        Dim Template As RepositoryItem
        Dim InitialValue As Object
        Select Case Kind
            Case NativePlannerBandEditorKind.ImportMode
                Template = State.ImportModeEditor
                InitialValue = State.ImportModeValue
            Case NativePlannerBandEditorKind.ScenarioName
                Template = State.ScenarioNameEditor
                InitialValue = State.ScenarioNameValue
            Case NativePlannerBandEditorKind.CopySource
                Template = State.CopySourceEditor
                InitialValue = State.CopySourceValue
            Case Else
                Return
        End Select

        NativePlannerBandActiveState = State
        NativePlannerBandActiveKind = Kind
        NativePlannerBandActiveEditor = Template.CreateEditor()
        NativePlannerBandActiveEditor.Properties.LockEvents()
        NativePlannerBandActiveEditor.Properties.Assign(Template)
        NativePlannerBandActiveEditor.BackColor = Template.Appearance.BackColor
        NativePlannerBandActiveEditor.ForeColor = Template.Appearance.ForeColor
        NativePlannerBandActiveEditor.Properties.AutoHeight = False
        NativePlannerBandActiveEditor.Parent = NativePlannerGrid
        NativePlannerBandActiveEditor.Bounds = Bounds
        NativePlannerBandActiveEditor.CreateControl()
        NativePlannerBandActiveEditor.EditValue = InitialValue
        NativePlannerBandActiveEditor.BringToFront()
        NativePlannerBandActiveEditor.Properties.UnLockEvents()

        AddHandler NativePlannerBandActiveEditor.Leave,
            AddressOf NativePlannerBandEditorLeave
        AddHandler NativePlannerBandActiveEditor.KeyDown,
            AddressOf NativePlannerBandEditorKeyDown
        If Kind <> NativePlannerBandEditorKind.ScenarioName Then
            AddHandler NativePlannerBandActiveEditor.EditValueChanged,
                AddressOf NativePlannerBandImportModeChanged
        End If

        NativePlannerBandActiveEditor.Focus()
        Dim ComboEditor As DevExpress.XtraEditors.ComboBoxEdit =
            TryCast(NativePlannerBandActiveEditor,
                    DevExpress.XtraEditors.ComboBoxEdit)
        If ComboEditor IsNot Nothing Then ComboEditor.ShowPopup()

    End Sub

    Private Sub NativePlannerBandImportModeChanged(
        sender As Object,
        e As EventArgs)

        Dim ChangedEditor As DevExpress.XtraEditors.BaseEdit =
            TryCast(sender, DevExpress.XtraEditors.BaseEdit)
        If ChangedEditor Is Nothing OrElse
           Not ReferenceEquals(ChangedEditor, NativePlannerBandActiveEditor) Then
            Return
        End If

        BeginInvoke(
            New MethodInvoker(
                Sub()
                    If ReferenceEquals(
                            ChangedEditor,
                            NativePlannerBandActiveEditor) Then
                        CloseNativePlannerBandEditor(True)
                    End If
                End Sub))

    End Sub

    Private Sub NativePlannerBandEditorLeave(
        sender As Object,
        e As EventArgs)

        CloseNativePlannerBandEditor(True)

    End Sub

    Private Sub NativePlannerBandEditorKeyDown(
        sender As Object,
        e As KeyEventArgs)

        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            e.SuppressKeyPress = True
            CloseNativePlannerBandEditor(True)
        ElseIf e.KeyCode = Keys.Escape Then
            e.Handled = True
            e.SuppressKeyPress = True
            CloseNativePlannerBandEditor(False)
        End If

    End Sub

    Private Sub NativePlannerBandLayout(
        sender As Object,
        e As EventArgs)

        CloseNativePlannerBandEditor(True)
        For Each State As NativePlannerBandEditorState In
            NativePlannerBandEditors
            State.ImportModeBounds = Rectangle.Empty
            State.ScenarioNameBounds = Rectangle.Empty
            State.CopySourceLabelBounds = Rectangle.Empty
            State.CopySourceBounds = Rectangle.Empty
            State.GoButtonBounds = Rectangle.Empty
        Next

    End Sub

    Private Sub CloseNativePlannerBandEditor(CommitValue As Boolean)

        If ClosingNativePlannerBandEditor OrElse
           NativePlannerBandActiveEditor Is Nothing Then Return

        ClosingNativePlannerBandEditor = True
        Try
            Dim Editor As DevExpress.XtraEditors.BaseEdit =
                NativePlannerBandActiveEditor
            Dim State As NativePlannerBandEditorState =
                NativePlannerBandActiveState
            Dim Kind As NativePlannerBandEditorKind =
                NativePlannerBandActiveKind
            Dim ChangedValue As Object = Editor.EditValue
            Dim OriginalValue As Object = Nothing
            Select Case Kind
                Case NativePlannerBandEditorKind.ImportMode
                    OriginalValue = State.ImportModeValue
                Case NativePlannerBandEditorKind.ScenarioName
                    OriginalValue = State.ScenarioNameValue
                Case NativePlannerBandEditorKind.CopySource
                    OriginalValue = State.CopySourceValue
            End Select

            RemoveHandler Editor.Leave,
                AddressOf NativePlannerBandEditorLeave
            RemoveHandler Editor.KeyDown,
                AddressOf NativePlannerBandEditorKeyDown
            RemoveHandler Editor.EditValueChanged,
                AddressOf NativePlannerBandImportModeChanged

            NativePlannerBandActiveEditor = Nothing
            NativePlannerBandActiveState = Nothing
            Editor.Dispose()

            If CommitValue AndAlso
               Not String.Equals(
                   Convert.ToString(ChangedValue),
                   Convert.ToString(OriginalValue),
                   StringComparison.Ordinal) Then
                CommitNativePlannerBandEditorValue(
                    State, Kind, ChangedValue)
            End If
        Finally
            ClosingNativePlannerBandEditor = False
        End Try

    End Sub

    Private Sub CommitNativePlannerBandEditorValue(
        State As NativePlannerBandEditorState,
        Kind As NativePlannerBandEditorKind,
        ChangedValue As Object)

        If State Is Nothing Then Return
        If Kind = NativePlannerBandEditorKind.CopySource Then
            State.CopySourceValue =
                TryCast(ChangedValue, NativePlannerCopySource)
            NativePlannerView.InvalidateBandHeader(State.Band)
            Return
        End If

        Dim Sheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Multivariable Planner")
        Dim TargetRow As Integer =
            If(Kind = NativePlannerBandEditorKind.ImportMode, 6, 7)
        Dim Description As String =
            If(Kind = NativePlannerBandEditorKind.ImportMode,
               "Stress-test scenario import mode updated",
               "Stress-test scenario name updated")
        Dim Target As DevExpress.Spreadsheet.Cell =
            Sheet.Cells(
                TargetRow,
                ScenarioStartColumn(State.ScenarioIndex))

        ProcessStressTestCellChange(
            Target,
            NormalizeStressTestEditValue(ChangedValue),
            "S",
            Description)

        LoadingNativeViews = True
        Try
            NativePlannerScenario.SelectedIndex =
                State.ScenarioIndex - 1
        Finally
            LoadingNativeViews = False
        End Try
        RefreshNativePlanner()

    End Sub

    Private Sub CopyNativePlannerScenario(
        State As NativePlannerBandEditorState)

        If State Is Nothing OrElse State.CopySourceValue Is Nothing Then
            DevExpress.XtraEditors.XtraMessageBox.Show(
                "Select a source test first.", "Copy test")
            Return
        End If

        Dim SourceIndex As Integer = State.CopySourceValue.ScenarioIndex
        Dim TargetIndex As Integer = State.ScenarioIndex
        If SourceIndex = TargetIndex Then Return

        If DevExpress.XtraEditors.XtraMessageBox.Show(
                "Copy Test " & SourceIndex.ToString() & " assumptions into Test " &
                TargetIndex.ToString() & "? This replaces the current test's " &
                "stress and mitigation inputs.",
                "Copy test", MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) <> DialogResult.Yes Then Return

        Dim PlannerSheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Multivariable Planner")
        Dim WasProtected As Boolean = PlannerSheet.IsProtected

        Me.Cursor = Cursors.WaitCursor
        UNProtectWS(ModelID, PlannerSheet.Name)
        Try
            CopyRangeValues(
                ActiveWorkbook.DefinedNames.GetDefinedName(
                    "Assumptions" & SourceIndex.ToString()).Range,
                ActiveWorkbook.DefinedNames.GetDefinedName(
                    "Assumptions" & TargetIndex.ToString()).Range)
            CopyRangeValues(
                ActiveWorkbook.DefinedNames.GetDefinedName(
                    "AssumptionsA" & SourceIndex.ToString()).Range,
                ActiveWorkbook.DefinedNames.GetDefinedName(
                    "AssumptionsA" & TargetIndex.ToString()).Range)

            ExcelModels(ModelID).IsDirty = True
            CalculateStressWorkbook(True)

            LoadingNativeViews = True
            Try
                NativePlannerScenario.SelectedIndex = TargetIndex - 1
            Finally
                LoadingNativeViews = False
            End Try
            RefreshNativePlanner()
            RefreshNativeDashboard()
            RefreshNativeComparativeViews()
        Catch ex As Exception
            DevExpress.XtraEditors.XtraMessageBox.Show(
                "The test could not be copied: " & ex.Message,
                "Copy test", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If WasProtected Then ProtectWS(ModelID, PlannerSheet.Name)
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub ResetNativePlannerBandEditors()

        CloseNativePlannerBandEditor(False)
        For Each State As NativePlannerBandEditorState In
            NativePlannerBandEditors
            State.ImportModeEditor.Dispose()
            State.ScenarioNameEditor.Dispose()
            State.CopySourceEditor.Dispose()
            State.GoButtonEditor.Dispose()
        Next
        NativePlannerBandEditors.Clear()

    End Sub

    Private Function CellToObject(Cell As DevExpress.Spreadsheet.Cell) As Object

        If Cell.Value.IsEmpty Then Return DBNull.Value
        If Cell.Value.IsNumeric Then Return Cell.Value.NumericValue
        If Cell.Value.IsBoolean Then Return Cell.Value.BooleanValue
        If Cell.Value.IsDateTime Then Return Cell.Value.DateTimeValue
        Return Cell.Value.TextValue

    End Function

    Private Sub NativePlannerScenarioChanged(sender As Object, e As EventArgs)

        If LoadingNativeViews Then Return
        LoadingNativeViews = True
        Try
            RefreshNativePlannerScenarioControls(
                ActiveWorkbook.Worksheets("Multivariable Planner"),
                SelectedPlannerScenarioIndex())
        Finally
            LoadingNativeViews = False
        End Try

    End Sub

    Private Sub NativePlannerFocusedColumnChanged(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Base.FocusedColumnChangedEventArgs)

        If LoadingNativeViews Then Return
        Dim ScenarioIndex As Integer
        Dim ValueOffset As Integer
        If Not TryGetPlannerScenarioColumn(
                e.FocusedColumn, ScenarioIndex, ValueOffset) OrElse
           NativePlannerScenario.SelectedIndex = ScenarioIndex - 1 Then Return

        LoadingNativeViews = True
        Try
            NativePlannerScenario.SelectedIndex = ScenarioIndex - 1
            RefreshNativePlannerScenarioControls(
                ActiveWorkbook.Worksheets("Multivariable Planner"), ScenarioIndex)
        Finally
            LoadingNativeViews = False
        End Try

    End Sub

    Private Sub NativePlannerNameChanged(sender As Object, e As EventArgs)

        If LoadingNativeViews Then Return
        Dim Sheet As DevExpress.Spreadsheet.Worksheet = ActiveWorkbook.Worksheets("Multivariable Planner")
        Dim Target As DevExpress.Spreadsheet.Cell =
            Sheet.Cells(7, ScenarioStartColumn(SelectedPlannerScenarioIndex()))
        ProcessStressTestCellChange(
            Target, NativePlannerName.EditValue, "S",
            "Stress-test scenario name updated")
        RefreshNativePlanner()

    End Sub

    Private Sub FirstTabGridCellValueChanged(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs)

        If FirstTabChangeInProgress Then Return

        Dim View As GridView = TryCast(sender, GridView)
        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If View Is Nothing OrElse
           Not TryGetFirstTabSourceCell(View, e.RowHandle, e.Column, SourceCell) OrElse
           Not IsWorkbookLinkedGridCellEditable(SourceCell) Then
            RefreshFirstTabGridData(View)
            Return
        End If

        FirstTabChangeInProgress = True
        Try
            Dim DataFormat As String = GetFirstTabDataFormat(SourceCell, e.Column, e.Value)
            ProcessStressTestCellChange(
                SourceCell, NormalizeStressTestEditValue(e.Value), DataFormat,
                "Live stress-test assumption updated", True)
            RefreshFirstTabGridData(View)
        Finally
            FirstTabChangeInProgress = False
        End Try

    End Sub

    Private Function GetFirstTabDataFormat(
        SourceCell As DevExpress.Spreadsheet.Cell,
        Column As DevExpress.XtraGrid.Columns.GridColumn,
        Value As Object) As String

        If Column IsNot Nothing AndAlso Column.AbsoluteIndex >= 5 Then Return "I"
        If Value Is Nothing OrElse Value Is DBNull.Value Then Return "S"
        If TypeOf Value Is String Then Return "S"
        If SourceCell IsNot Nothing AndAlso
           SourceCell.NumberFormat IsNot Nothing AndAlso
           SourceCell.NumberFormat.Contains("%") Then Return "P"
        Return "N"

    End Function

    Private Sub NativePlannerImportModeChanged(sender As Object, e As EventArgs)

        If LoadingNativeViews OrElse NativePlannerScenario.SelectedIndex < 0 Then Return
        Dim Target As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.Worksheets("Multivariable Planner").
                Cells(6, ScenarioStartColumn(SelectedPlannerScenarioIndex()))
        ProcessStressTestCellChange(
            Target, NativePlannerImportMode.EditValue, "S",
            "Stress-test scenario import mode updated")
        RefreshNativePlanner()

    End Sub

    Private Sub NativePlannerIncludeChanged(sender As Object, e As EventArgs)

        If LoadingNativeViews Then Return
        Dim Target As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.Worksheets("Multivariable Planner").Cells(6, 3)
        ProcessStressTestCellChange(
            Target, If(NativePlannerInclude.Checked, "Yes", Nothing), "S",
            "Base scenario import setting updated")
        RefreshNativePlanner()

    End Sub

    Private Sub NativePlannerCellValueChanged(sender As Object,
                                              e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs)

        If LoadingNativeViews OrElse e.RowHandle < 0 Then Return

        Dim SourceRow As Integer =
            Convert.ToInt32(NativePlannerView.GetRowCellValue(e.RowHandle, "SourceRow"))

        If e.Column.FieldName = "ShortName" Then
            Dim ShortNameCell As DevExpress.Spreadsheet.Cell =
                ActiveWorkbook.Worksheets("Multivariable Planner").Cells(SourceRow, 1)

            If Not ProcessStressTestCellChange(
                    ShortNameCell, NormalizeStressTestEditValue(e.Value), "S",
                    "Stress-test assumption short name updated", True) Then
                RefreshNativePlanner()
            End If
            Return
        End If

        Dim ScenarioIndex As Integer
        Dim ValueOffset As Integer
        If Not TryGetPlannerScenarioColumn(
                e.Column, ScenarioIndex, ValueOffset) Then Return
        Dim Target As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.Worksheets("Multivariable Planner").
                Cells(SourceRow,
                      ScenarioStartColumn(ScenarioIndex) + ValueOffset)
        Dim DataFormat As String = "N"
        If ValueOffset >= 2 Then
            DataFormat = "I"
        ElseIf SourceRow = 34 OrElse SourceRow = 64 OrElse SourceRow = 65 Then
            DataFormat = "S"
        ElseIf Convert.ToString(
                NativePlannerView.GetRowCellValue(
                    e.RowHandle,
                    PlannerScenarioFormatFieldName(ScenarioIndex))).Contains("%") Then
            DataFormat = "P"
        End If
        If Not ProcessStressTestCellChange(
                Target, NormalizeStressTestEditValue(e.Value), DataFormat,
                "Stress-test scenario " & ScenarioIndex.ToString() &
                " planner assumption updated", True) Then
            RefreshNativePlanner()
        End If

    End Sub

    Private Function NormalizeStressTestEditValue(Value As Object) As Object

        If Value Is Nothing OrElse Value Is DBNull.Value OrElse
           String.IsNullOrWhiteSpace(Convert.ToString(Value)) Then Return Nothing
        Return Value

    End Function

    Private Function ProcessStressTestCellChange(
        Target As DevExpress.Spreadsheet.Cell,
        ChangedValue As Object,
        DataFormat As String,
        Description As String,
        Optional RequireGridPattern As Boolean = False) As Boolean

        If Target Is Nothing OrElse Target.Protection.Locked OrElse ChangeMan Is Nothing OrElse
           (RequireGridPattern AndAlso Not IsWorkbookLinkedGridCellEditable(Target)) Then
            Return False
        End If

        Dim ChangeEvent As New DataChangeEvent With {
            .ModelID = ModelID,
            .Description = Description,
            .WSName = Target.Worksheet.Name,
            .CellAddress = Target.GetReferenceA1(),
            .OriginalValue = CellToObject(Target),
            .ChangedValue = ChangedValue,
            .DataFormat = DataFormat,
            .TimeStamp = Now(),
            .UserName = Environment.UserName
        }
        Return Not ChangeMan.ProcessChange(ChangeEvent).BError

    End Function

    Private Sub NativePlannerCustomDrawCell(
        sender As Object,
        e As DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs)

        Dim SourceCell As DevExpress.Spreadsheet.Cell = Nothing
        If Not TryGetNativePlannerSourceCell(e.RowHandle, e.Column, SourceCell) Then Return

        If e.Column.OptionsColumn.AllowEdit AndAlso
           Not IsWorkbookLinkedGridCellEditable(SourceCell) Then
            'Mirror DataInterfaceTemplate: the pattern decides whether a
            'rule-controlled cell is locked; locked cells are then drawn using
            'the standard disabled-rule appearance rather than the pattern colour.
            e.Appearance.BackColor = Color.Lavender
            e.Appearance.ForeColor = Color.WhiteSmoke
            e.Appearance.Options.UseBackColor = True
            e.Appearance.Options.UseForeColor = True
        Else
            ApplyWorkbookCellAppearance(e.Appearance, SourceCell)
        End If

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
                SourceCell) OrElse Not IsWorkbookLinkedGridCellEditable(SourceCell) Then
            e.Cancel = True
        End If

    End Sub

    Private Function IsWorkbookLinkedGridCellEditable(
        SourceCell As DevExpress.Spreadsheet.Cell) As Boolean

        If SourceCell Is Nothing OrElse SourceCell.Protection.Locked Then Return False

        'Keep the StressTest grids aligned with DataInterfaceTemplate. Workbook
        'conditional-format rules change the linked cell's fill pattern; Solid
        'means the input is active, while any other pattern makes it read-only.
        Return SourceCell.Fill.PatternType = PatternType.Solid

    End Function

    Private Function TryGetNativePlannerSourceCell(
        RowHandle As Integer,
        Column As DevExpress.XtraGrid.Columns.GridColumn,
        ByRef SourceCell As DevExpress.Spreadsheet.Cell) As Boolean

        If RowHandle < 0 OrElse Column Is Nothing Then Return False

        Dim SourceRowValue As Object =
            NativePlannerView.GetRowCellValue(RowHandle, "SourceRow")
        If SourceRowValue Is Nothing OrElse SourceRowValue Is DBNull.Value Then Return False

        Dim SourceColumn As Integer
        If Column.FieldName = "Assumption" Then
            SourceColumn = 0
        ElseIf Column.FieldName = "ShortName" Then
            SourceColumn = 1
        Else
            Dim ScenarioIndex As Integer
            Dim ValueOffset As Integer
            If Not TryGetPlannerScenarioColumn(
                    Column, ScenarioIndex, ValueOffset) Then Return False
            SourceColumn = ScenarioStartColumn(ScenarioIndex) + ValueOffset
        End If

        SourceCell = ActiveWorkbook.Worksheets("Multivariable Planner").Cells(
            Convert.ToInt32(SourceRowValue), SourceColumn)
        Return SourceCell IsNot Nothing

    End Function

    Private Sub ApplyWorkbookCellAppearance(
        Appearance As DevExpress.Utils.AppearanceObject,
        SourceCell As DevExpress.Spreadsheet.Cell)

        If SourceCell Is Nothing Then Return

        'DataInterfaceTemplate stores the workbook background separately from
        'its pattern-based rule lock. Keep those responsibilities separate here.
        Dim Background As Color = SourceCell.Fill.BackgroundColor
        If Background.IsEmpty OrElse Background.A = 0 Then Background = Color.White

        Dim Foreground As Color = SourceCell.Font.Color
        If Foreground.IsEmpty OrElse Foreground.A = 0 Then Foreground = AbovoBlue

        Appearance.BackColor = Background
        Appearance.ForeColor = Foreground
        Appearance.Options.UseBackColor = True
        Appearance.Options.UseForeColor = True

    End Sub

    Private Sub ApplyWorkbookResolvedCellAppearance(
        Appearance As DevExpress.Utils.AppearanceObject,
        SourceCell As DevExpress.Spreadsheet.Cell)

        If SourceCell Is Nothing Then Return

        'This range is a direct visual reproduction rather than a DIT rule
        'surface, so use the resolved fill (including conditional formatting).
        Dim Background As Color = SourceCell.FillColor
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
        Dim ScenarioIndex As Integer
        Dim ValueOffset As Integer
        If Not TryGetPlannerScenarioColumn(
                e.Column, ScenarioIndex, ValueOffset) Then Return
        If ValueOffset >= 2 Then
            e.RepositoryItem = NativeYearEditor
            Return
        End If
        Dim SourceRow As Integer =
            Convert.ToInt32(NativePlannerView.GetRowCellValue(e.RowHandle, "SourceRow"))
        Dim Format As String =
            Convert.ToString(
                NativePlannerView.GetRowCellValue(
                    e.RowHandle,
                    PlannerScenarioFormatFieldName(ScenarioIndex)))
        If SourceRow = 34 OrElse SourceRow = 64 OrElse SourceRow = 65 Then
            e.RepositoryItem = StandardYEmptyEdit
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
           e.Value Is Nothing OrElse e.Value Is DBNull.Value Then Return
        Dim ScenarioIndex As Integer
        Dim ValueOffset As Integer
        If Not TryGetPlannerScenarioColumn(
                e.Column, ScenarioIndex, ValueOffset) OrElse ValueOffset > 1 Then Return
        Dim RowHandle As Integer = NativePlannerView.GetRowHandle(e.ListSourceRowIndex)
        Dim Format As String =
            Convert.ToString(
                NativePlannerView.GetRowCellValue(
                    RowHandle,
                    PlannerScenarioFormatFieldName(ScenarioIndex)))
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

        Dim ScenarioIndex As Integer = SelectedPlannerScenarioIndex()
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
        ExcelModels(ModelID).IsDirty = True
        CalculateStressWorkbook()
        RefreshNativePlanner()

    End Sub

    Private Sub RecalculateAndRefreshNativeViews()

        Me.Cursor = Cursors.WaitCursor
        Try
            CalculateStressWorkbook()
            RefreshNativeTargets()
            RefreshNativeDashboard()
            RefreshNativeComparativeViews()
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub CalculateStressWorkbook(Optional UseRecursiveEngine As Boolean = False)

        Dim PreviousEngine As DevExpress.Spreadsheet.CalculationEngineType =
            ActiveWorkbook.Options.CalculationEngineType

        Try
            If UseRecursiveEngine AndAlso
               PreviousEngine <> DevExpress.Spreadsheet.CalculationEngineType.Recursive Then
                ActiveWorkbook.Options.CalculationEngineType =
                    DevExpress.Spreadsheet.CalculationEngineType.Recursive
            End If

            If ExcelModels(ModelID).WBCalcEngine IsNot Nothing Then
                ExcelModels(ModelID).WBCalcEngine.CalcFile()
            Else
                ActiveWorkbook.Calculate()
            End If
        Finally
            If ActiveWorkbook.Options.CalculationEngineType <> PreviousEngine Then
                ActiveWorkbook.Options.CalculationEngineType = PreviousEngine
            End If
        End Try

    End Sub

    Private Sub LiveScenarioNumberChanged(sender As Object, e As EventArgs)

        If LoadingNativeViews OrElse ComboBoxBreachMode.SelectedIndex < 0 Then Return

        Dim Target As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.DefinedNames.GetDefinedName("StressTestNumber").Range(0, 0)
        If Not ProcessStressTestCellChange(
                Target, ComboBoxBreachMode.SelectedItem.ToString(), "S",
                "Live stress-test capture slot updated") Then
            LoadingNativeViews = True
            Try
                ComboBoxBreachMode.EditValue = Target.DisplayText
            Finally
                LoadingNativeViews = False
            End Try
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
        Dim PlannerSheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Multivariable Planner")
        Dim CapturedDataSheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("OW - Captured Data")
        Dim PlannerWasProtected As Boolean = PlannerSheet.IsProtected
        Dim CapturedDataWasProtected As Boolean = CapturedDataSheet.IsProtected
        UNProtectWS(ModelID, PlannerSheet.Name)
        UNProtectWS(ModelID, CapturedDataSheet.Name)
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
            CalculateStressWorkbook(True)
            CopyRangeValues(
                ActiveWorkbook.DefinedNames.GetDefinedName("StressLiveInfo").Range,
                ActiveWorkbook.DefinedNames.GetDefinedName(
                    "S" & ScenarioIndex.ToString() & "Data").Range)
            ExcelModels(ModelID).IsDirty = True
            RefreshAllNativeScenarioSelectors()
            RefreshNativePlanner()
            RefreshNativeDashboard()
            RefreshNativeComparativeViews()
        Finally
            If PlannerWasProtected Then ProtectWS(ModelID, PlannerSheet.Name)
            If CapturedDataWasProtected Then ProtectWS(ModelID, CapturedDataSheet.Name)
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
        Dim PlannerSheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Multivariable Planner")
        Dim CapturedDataSheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("OW - Captured Data")
        Dim LivePlannerSheet As DevExpress.Spreadsheet.Worksheet =
            ActiveWorkbook.Worksheets("Live Multivariable Planner")
        Dim PlannerWasProtected As Boolean = PlannerSheet.IsProtected
        Dim CapturedDataWasProtected As Boolean = CapturedDataSheet.IsProtected
        Dim LivePlannerWasProtected As Boolean = LivePlannerSheet.IsProtected
        UNProtectWS(ModelID, PlannerSheet.Name)
        UNProtectWS(ModelID, CapturedDataSheet.Name)
        UNProtectWS(ModelID, LivePlannerSheet.Name)
        Try
            SetWorkbookStressMode(False)
            LiveAssumptions.ClearContents()
            LiveAssumptionsA.ClearContents()
            CalculateStressWorkbook(True)
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
                    CalculateStressWorkbook(True)
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
            Try
                RestoreRange(LiveAssumptions, SavedLive)
                RestoreRange(LiveAssumptionsA, SavedLiveA)
                ActiveWorkbook.DefinedNames.GetDefinedName(
                    "StressTestMode").Range(0, 0).Value = CellValue.FromObject(SavedMode)
                ActiveWorkbook.DefinedNames.GetDefinedName("Mode").RefersTo = SavedModeReference
                ExcelModels(ModelID).IsDirty = True
                CalculateStressWorkbook(True)
            Finally
                If PlannerWasProtected Then ProtectWS(ModelID, PlannerSheet.Name)
                If CapturedDataWasProtected Then ProtectWS(ModelID, CapturedDataSheet.Name)
                If LivePlannerWasProtected Then ProtectWS(ModelID, LivePlannerSheet.Name)
                Me.Text = OriginalTitle
                Me.Cursor = Cursors.Default
            End Try
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
                NativePlannerScenario.Properties.Items.AddRange(
                    Names.Skip(1).Cast(Of Object).ToArray())
                NativePlannerScenario.SelectedIndex =
                    Math.Min(Selected, NativePlannerScenario.Properties.Items.Count - 1)
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

    Private Sub RefreshNativeDashboard()

        If NativeDashboardHost Is Nothing OrElse ActiveWorkbook Is Nothing OrElse
           Not ActiveWorkbook.Worksheets.Contains("Multivariable Dashboard") Then Return

        AttachNativeDashboardSpreadsheet()
        ActiveWorkbook.Worksheets.ActiveWorksheet =
            ActiveWorkbook.Worksheets("Multivariable Dashboard")
        ActiveWorkbook.Worksheets.ActiveWorksheet.ActiveView.ShowGridlines = False
        ActiveWorkbook.Worksheets.ActiveWorksheet.ActiveView.ShowHeadings = False
        ActiveWorkbook.Worksheets.ActiveWorksheet.ActiveView.Zoom = 150
        CalculateStressWorkbook()
        If NativeDashboardSpreadsheet IsNot Nothing Then
            NativeDashboardSpreadsheet.Refresh()
            NativeDashboardSpreadsheet.Focus()
        End If

    End Sub

    Private Sub AttachNativeDashboardSpreadsheet()

        NativeDashboardSpreadsheet = ExcelModels(ModelID).ModelSpreadsheetControl
        If NativeDashboardSpreadsheet Is Nothing OrElse NativeDashboardHost Is Nothing Then Return

        If NativeDashboardSpreadsheet.Parent IsNot NativeDashboardHost Then
            If NativeDashboardSpreadsheetOriginalParent Is Nothing Then
                NativeDashboardSpreadsheetOriginalParent = NativeDashboardSpreadsheet.Parent
                NativeDashboardSpreadsheetOriginalDock = NativeDashboardSpreadsheet.Dock
                If NativeDashboardSpreadsheetOriginalParent IsNot Nothing Then
                    NativeDashboardSpreadsheetOriginalIndex =
                        NativeDashboardSpreadsheetOriginalParent.Controls.GetChildIndex(
                            NativeDashboardSpreadsheet)
                End If
            End If
            If NativeDashboardSpreadsheet.Parent IsNot Nothing Then
                NativeDashboardSpreadsheet.Parent.Controls.Remove(NativeDashboardSpreadsheet)
            End If
            NativeDashboardHost.Controls.Add(NativeDashboardSpreadsheet)
            NativeDashboardSpreadsheet.Dock = DockStyle.Fill
        End If

        If Not NativeDashboardSpreadsheetChangeHandlerAttached Then
            AddHandler NativeDashboardSpreadsheet.CellValueChanged,
                AddressOf NativeDashboardSpreadsheetCellValueChanged
            NativeDashboardSpreadsheetChangeHandlerAttached = True
        End If

    End Sub

    Private Sub RestoreNativeDashboardSpreadsheet()

        If NativeDashboardSpreadsheet Is Nothing Then Return
        If NativeDashboardSpreadsheetChangeHandlerAttached Then
            RemoveHandler NativeDashboardSpreadsheet.CellValueChanged,
                AddressOf NativeDashboardSpreadsheetCellValueChanged
            NativeDashboardSpreadsheetChangeHandlerAttached = False
        End If
        If NativeDashboardSpreadsheet.Parent IsNot Nothing Then
            NativeDashboardSpreadsheet.Parent.Controls.Remove(NativeDashboardSpreadsheet)
        End If
        If NativeDashboardSpreadsheetOriginalParent IsNot Nothing Then
            NativeDashboardSpreadsheetOriginalParent.Controls.Add(NativeDashboardSpreadsheet)
            NativeDashboardSpreadsheet.Dock = NativeDashboardSpreadsheetOriginalDock
            If NativeDashboardSpreadsheetOriginalIndex >= 0 Then
                NativeDashboardSpreadsheetOriginalParent.Controls.SetChildIndex(
                    NativeDashboardSpreadsheet,
                    Math.Min(NativeDashboardSpreadsheetOriginalIndex,
                             NativeDashboardSpreadsheetOriginalParent.Controls.Count - 1))
            End If
        End If
        NativeDashboardSpreadsheetOriginalParent = Nothing
        NativeDashboardSpreadsheetOriginalIndex = -1

    End Sub

    Private Sub NativeDashboardSpreadsheetCellValueChanged(
        sender As Object,
        e As SpreadsheetCellEventArgs)

        If ReplayingNativeDashboardCellChange OrElse e.Cell Is Nothing OrElse
           Not String.Equals(
               e.Cell.Worksheet.Name,
               "Multivariable Dashboard",
               StringComparison.OrdinalIgnoreCase) Then Return

        Dim ChangedValue As Object
        If e.Value.IsNumeric Then
            ChangedValue = e.Value.NumericValue
        ElseIf e.Value.IsBoolean Then
            ChangedValue = e.Value.BooleanValue
        ElseIf e.Value.IsDateTime Then
            ChangedValue = e.Value.DateTimeValue
        Else
            ChangedValue = e.Value.TextValue
        End If
        Dim DataFormat As String =
            If(e.Cell.NumberFormat IsNot Nothing AndAlso
               e.Cell.NumberFormat.Contains("%"), "P",
               If(e.Value.IsNumeric, "N", "S"))

        ReplayingNativeDashboardCellChange = True
        Try
            e.Cell.Value = e.OldValue
            If ProcessStressTestCellChange(
                    e.Cell,
                    ChangedValue,
                    DataFormat,
                    "Multivariable dashboard cell updated") Then
                CalculateStressWorkbook()
            End If
        Finally
            ReplayingNativeDashboardCellChange = False
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
        Dim Target As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.Worksheets("Comparative").Cells(11 + (Slot * 2), 2)
        If Not ProcessStressTestCellChange(
                Target, Selector.EditValue, "S",
                "Stress-test comparison scenario updated") Then
            RefreshNativeComparativeViews()
            Return
        End If
        RefreshNativeComparativeViews()

    End Sub

    Private Sub NativeComparisonYearChanged(sender As Object, e As EventArgs)

        If LoadingNativeViews Then Return
        Dim Editor As DevExpress.XtraEditors.SpinEdit =
            TryCast(sender, DevExpress.XtraEditors.SpinEdit)
        If Editor Is Nothing Then Return
        Dim WorkingRow As Integer = Convert.ToInt32(Editor.Tag)
        Dim Target As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.Worksheets("OW - Covenant Calculation").Cells(WorkingRow - 1, 2)
        If Not ProcessStressTestCellChange(
                Target, Convert.ToInt32(Editor.Value), "I",
                "Stress-test comparison start year updated") Then
            LoadingNativeViews = True
            Try
                Editor.EditValue = GetNumericValue(Target)
            Finally
                LoadingNativeViews = False
            End Try
            Return
        End If
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
                Dim SourceCell As DevExpress.Spreadsheet.Cell =
                    Sheet.Cells(11 + (Slot * 2), 2)
                Selector.Properties.ReadOnly = SourceCell.Protection.Locked
                Dim SelectedName As String =
                    SourceCell.DisplayText
                Selector.Properties.Items.Clear()
                Selector.Properties.Items.AddRange(Names.Cast(Of Object).ToArray())
                Dim Match As Integer = Names.FindIndex(
                    Function(Item) String.Equals(
                        Item, SelectedName, StringComparison.OrdinalIgnoreCase))
                Selector.SelectedIndex =
                    If(Match >= 0, Match, Math.Min(Slot + 1, Names.Count - 1))
            Next

            CalculateStressWorkbook()
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

        Dim Target As DevExpress.Spreadsheet.Cell =
            ActiveWorkbook.DefinedNames.GetDefinedName("NewStressName").Range(0, 0)
        If Not ProcessStressTestCellChange(
                Target, TextEditMultivariableName.EditValue, "S",
                "Live stress-test scenario name updated") Then
            RemoveHandler TextEditMultivariableName.Validated,
                AddressOf TextEditMultivariableName_EditValueChanged
            Try
                TextEditMultivariableName.EditValue = Target.DisplayText
            Finally
                AddHandler TextEditMultivariableName.Validated,
                    AddressOf TextEditMultivariableName_EditValueChanged
            End Try
        End If

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

