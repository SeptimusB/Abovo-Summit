Imports System.ComponentModel
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox
Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.AbovoExtendedDEControls
Imports Abovo.AbovoRangeDataSource
Imports Abovo.AbovoUnboundSource
Imports Abovo.ChangeLogManager
Imports Abovo.CustomGrid
Imports Abovo.DataObject
Imports Abovo.DefaultHelpers
Imports Abovo.FileManager
Imports Abovo.GeneralFunctions
Imports Abovo.LogDebugDev
Imports Abovo.MasterChangeLog
Imports Abovo.PresentationManager
Imports Abovo.RepositaryItems
Imports DataInterfaceTemplate.AbovoGridRespoitaryCombo
Imports DevExpress
Imports DevExpress.Accessibility
Imports DevExpress.CodeParser
Imports DevExpress.Data
Imports DevExpress.Pdf.Native
Imports DevExpress.Pdf.Native.BouncyCastle.Asn1.X509.Qualified
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils
Imports DevExpress.Utils.Behaviors.Common
Imports DevExpress.Utils.Drawing.Helpers
Imports DevExpress.Utils.Extensions
Imports DevExpress.Utils.Gesture
Imports DevExpress.Utils.Layout
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.XtraBars.Navigation
Imports DevExpress.XtraCharts
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraExport.Helpers
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Tab
Imports DevExpress.XtraGrid.Views
Imports DevExpress.XtraGrid.Views.BandedGrid
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports DevExpress.XtraLayout
Imports DevExpress.XtraLayout.Customization
Imports DevExpress.XtraLayout.Utils
Imports DevExpress.XtraPrinting.Export
Imports DevExpress.XtraPrinting.Native
Imports DevExpress.XtraRichEdit.Commands
Imports DevExpress.XtraRichEdit.Import.Rtf
Imports DevExpress.XtraRichEdit.Model
Imports DevExpress.XtraSpreadsheet
Imports DevExpress.XtraSpreadsheet.Commands.Internal
Imports DevExpress.XtraSpreadsheet.Model
Imports DevExpress.XtraTab
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.XtraVerticalGrid.Rows
Imports Microsoft.Win32
Public Class DataInterfaceTemplate

    Inherits DevExpress.XtraEditors.XtraForm



    'Public CollectionVariables
    Public ReliantNRs As List(Of String)
    Public AmActivated As Boolean = True
    Public DITName As String
    'Higher Level Classes/Objects
    Private DataPM As PresentationManager
    Private ActionMan As ActionManager
    Private DataPres As PresentationManager.DataPresentation
    Private Formatter As ObjectFormatter
    Private rs As New Resizer
    'Private TTController As ToolTipController
    Private ParentGroupForm As GroupInterfaceTemplate
    Private ActiveSpreadsheet As DevExpress.Spreadsheet.Worksheet
    Private ActiveWorkbook As IWorkbook
    'InterfaceTag
    Private InterfaceTag As AbovoInterfaceTag
    Private CalcEngID As Integer

    'Private Data Variables
    Private ModelID As Integer
    Private PresID As Integer
    Private GSID As Integer
    Private CSID As Integer
    Private CustCalcEdit As RepositoryItemCalcEdit
    Private UnboundDataSources() As AbovoUnboundSource
    Private RangeDataSources() As RangeDataSource
    Private UBSDataSourceCount As Integer = -1
    Private RangeDataSourceCount As Integer = -1
    Private DblClickCell As Boolean = False
    Private MyData As DataObject
    Private PresentedDS As Abovo.DataObject.DataCellRange
    Private PresentedColumn As Abovo.DataObject.SheetDataColumn
    Private ChangeMan As ModelChangeManager

    'Layout Variables
    Private ScaleUnits As Single
    Private Scalefactor As Single
    Private IdealGridRowHeight As Integer = 54
    Private MinGridRowHeight As Integer = 20

    'State Variables
    Private BIsDirty As Boolean

    Private ControlsInitialised As Boolean = False
    Private FooterOn As Boolean
    Private FooterDone As Boolean
    Private DIInitialised As Boolean = False
    Private SheetIInitialised As Boolean = False
    Private InterfaceMode As String
    Private RightClick As Boolean = False

    'Placeholder variables for constructors
    Private PropertiesCount As Integer
    Private PropertyList As IEnumerable(Of UnboundSourceProperty)
    Private PropType As System.Type
    Private ColList As List(Of String)
    Private TPs() As TablePanel
    Private XtraTabPages() As XtraTabPage
    Private XtraTabPageCount As Integer = 0
    Private ColCount As Integer = -1
    Private ColName As String
    'Private ColExtenders() As ColumnButtonExtender
    'Private ColExtenderCount As Integer = -1
    Private BandExtenders() As BandButtonExtender
    Private InHeaderCombos() As AbovoRespositaryItem
    Private InHeaderCombosCount As Integer = -1
    Private BandExtenderCount As Integer = -1
    Private FixedLeftIndent = 200

    'Static Controls
    Private AcControl As AccordionControl

    'Dynamic Arrays
    Private PropertyArray() As UnboundSourceProperty

    'Arrays for Dynamic Contols
    Private GridControls() As GridControl
    Private VertGridControls() As VGridControl
    Private ExpanableGridControls() As GridControl
    Private ExpanableGridControlsCount As Integer = -1
    Private TextBoxes() As AbovoDETextEdit
    Private DateBoxes() As AbovoDEDateEdit
    Private Labels() As DevExpress.XtraEditors.LabelControl
    Private SpinEdits() As AbovoDESpinEdit
    Private Combos() As DevExpress.XtraEditors.ComboBoxEdit
    Private AcElements() As AccordionControlElement
    Private AcElementlist As List(Of AccordionControlElement)
    Private UsedGridVIEWS(-1) As GridView
    Private UsedBANDedGridVIEWS(-1) As BandedGridView
    Private UsedBANDedGridViewBANDS(-1) As GridBand
    Private UsedBANDedGridViewCOLS(-1) As BandedGridColumn
    Private AcContainers() As AccordionContentContainer
    Private hyperlinkLabelControls() As HyperlinkLabelControl
    Private DontAddCDHDef As Boolean = False
    Private RefreshableControls() As Object
    Private HaveCalledData As Boolean = False
    Private ComboReposClasses() As AbovoGridRespoitaryCombo
    Private AbovoTabPages() As AbovoTabPage
    Private AbovoTabPagesCount As Integer = -1
    Private CurrentAbovoTabPage As AbovoTabPage

    'Layout Control variables



    'MappedTable Variables

    Private ADELabels() As AbovoDELabel
    Private ADELabelsCount As Integer = -1
    Private ADEDateEdits() As AbovoDEDateEdit
    Private ADEDateEditsCount As Integer = -1
    Private ADEHyperlinks() As AbovoDEHyperlinkLabel
    Private ADEHyperlinksCount As Integer = -1
    Private ADETextEdits() As AbovoDETextEdit
    Private ADETextEditsCount As Integer = -1
    Private ADEComboBoxes() As AbovoDEComboBox
    Private ADEComboBoxesCount As Integer = -1
    Private ComboReposClassesCount As Integer = -1
    'Initialisation Integers for Arrays
    Private RefreshableControlsCount As Integer = -1
    Private TextEditCount As Integer = -1
    Private BandExtendersCount As Integer = -1
    Private DateEditCount As Integer = -1
    Private CombosCount As Integer = -1
    Private GridCombosCount As Integer = -1
    Private SpinEditCount As Integer = -1
    Private ComboEditCount As Integer = -1
    Private LabelsCount As Integer = -1
    Private GridCount As Integer = -1
    Private VertGridCount As Integer = -1
    Private AcControlCount As Integer = -1
    Private AcElementCount As Integer = -1
    Private GridViewCount As Integer = -1
    Private AcContainersCount As Integer = -1
    Private BandGridViewsCount As Integer = -1
    Private BandGridViewBandsCount As Integer = -1
    Private BandGridViewColsCount As Integer = -1

    Private ActiveLinkElement As ElementInterfaceLinkTag

    'Interface controls
    Private WindowsUIButtonPanelSaveClose As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel = New DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel()

    'Development Variables
    Private DataCallCount As Integer = 0


    Public Sub New(SetModelID As Integer, Optional ByVal SetGSID As Integer = -1, Optional ByVal SetCSID As Integer = -1, Optional ByVal MyParent As GroupInterfaceTemplate = Nothing, Optional ByVal SetInterfaceMode As String = "Normal", Optional ByVal Interfacelink As ElementInterfaceLinkTag = Nothing)

        ParentGroupForm = MyParent

        InitializeComponent()

        If Not IsNothing(MyParent) Then

            Me.Width = MyParent.Width

        End If

        ActiveWorkbook = ExcelModels(SetModelID).WB

        ActiveLinkElement = Interfacelink

        AddMenuInterface(InterfaceMode)

        Formatter = New ObjectFormatter

        InterfaceMode = SetInterfaceMode
        GSID = SetGSID
        CSID = SetCSID
        ModelID = SetModelID
        FooterOn = False

        DITName = GetCSName(SetModelID, SetGSID, SetCSID)

        ChangeMan = ExcelModels(SetModelID).ChangeManager
        DataPM = ExcelModels(SetModelID).WBDataPres
        Me.ForeColor = ExcelModels(SetModelID).ColourSwatch

        Dim CheckTrans As AbovoTransaction
        CheckTrans = DataPM.ValidatePresentation(GSID, CSID)

        If CheckTrans.BError = False Then

            PresID = CheckTrans.IntegerReturn
            DataPres = DataPM.DataPresentations(PresID)

        End If

        CalcEngID = ExcelModels(ModelID).WBCalcEngine.AddActiveObject(Me)

        If DataPres.DefaultWorksheet IsNot Nothing Then

            ActiveSpreadsheet = ExcelModels(SetModelID).WB.Worksheets(DataPres.DefaultWorksheet)

            ExcelModels(ModelID).WBCalcEngine.AddActiveWorksheet(CalcEngID, ActiveSpreadsheet)

        End If

        'XtraTabControlDI = New XtraTabControl With {
        '    .Dock = DockStyle.Fill
        '    }

        'TablePanelDIT.Controls.Add(XtraTabControlDI)
        'TablePanelDIT.SetCell(XtraTabControlDI, 1, 0)
        'TTController = New ToolTipController With {
        '    .ToolTipType = ToolTipType.Default,
        '    .Rounded = True}

        If WorkMode = "INTERFACE" Then


            PopulateDataInterface()

        ElseIf WorkMode = "SPREADSHEET" Then


            PopulateSpreadhsheetInterface()

        End If

        CustCalcEdit = New RepositoryItemCalcEdit With {
            .TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard,
            .UseMaskAsDisplayFormat = True,
            .UseAdvancedMode = DefaultBoolean.True,
            .Precision = 0}

        AddHandler CustCalcEdit.Click, AddressOf CalcEditPopup

        If ActiveLinkElement IsNot Nothing Then ProcessLinkElement()

        ControlsInitialised = True

        Debug.Print("XXXXXXXXXX - Resize call -XXXXXXXXXXXX")

        ResizeFonts()

        Debug.Print("XXXXXXXXXX - Resize call complete - XXXXXXXXXXXX")

        Exit Sub

    End Sub
    'This is added to the recently created CalcEdit for single click show
    Sub CalcEditPopup(Sender As Object, e As EventArgs)

        Dim a As CalcEdit = Sender
        'a.ShowPopup()

    End Sub

#Region "Public Methods / Properties"

    Public Shadows Sub Deactivate()

        ExcelModels(ModelID).WBCalcEngine.RemoveActiveObject(CalcEngID)

        AmActivated = False

    End Sub
    Public Sub Reactivate()

        ExcelModels(ModelID).WBCalcEngine.AddActiveObject(Me)

        ExcelModels(ModelID).WBCalcEngine.CalcFile()

        RefreshData()

        UpdateTabPage()

        AmActivated = True

    End Sub

    Public Sub RefreshData()

        If Me.GridCount < 1 Then GoTo SkipGridRefresh

        If Not IsNothing(DataPres) Then

            For Each gridControl In GridControls

                If Not IsNothing(gridControl) Then

                    gridControl.RefreshDataSource()

                End If

            Next

        End If

SkipGridRefresh:

        If RefreshableControls Is Nothing Then GoTo SkipRefresh

        For Each RefControl As Object In RefreshableControls
            Try
                If Not IsNothing(RefControl) Then
                    RefControl.RefreshData()
                End If
            Catch ex As Exception
                ' Handle exception
                MsgBox("Error refreshing control: " & ex.Message)
            End Try
        Next

SkipRefresh:

        UpdateAllRules()

    End Sub



#End Region

#Region "Interface Methods"
    Public Sub RebuildInterface()

        If WorkMode = "INTERFACE" Then

            ClearDataInterface()

            PopulateDataInterface()

        ElseIf WorkMode = "SPREADSHEET" Then

            PopulateSpreadhsheetInterface()

        End If

    End Sub
    Public Sub ClearDataInterface()

        If Not IsNothing(XtraTabControlDI) Then

            XtraTabControlDI.TabPages.Clear()

            GridControls = Nothing
            UsedGridVIEWS = Nothing
            'UnboundDataSources = Nothing
            RangeDataSources = Nothing
            UBSDataSourceCount = -1
            RangeDataSourceCount = -1
            GridCount = -1
            GridViewCount = -1

        End If

        ControlsInitialised = False

    End Sub
    Sub PopulateSpreadhsheetInterface()

    End Sub
    Private Sub PopulateDataInterface()

        ResetTimer(Me.Name & " Populate Data Interface")

        If DataPres.Sections.Length < 0 Then Exit Sub

        'Format Tab Control?

        WindowsFormsSettings.SmartMouseWheelProcessing = True

        Dim Section As PresentationSection

        Dim hyperlinkLabelControl1 As New HyperlinkLabelControl()

        AcElementlist = New List(Of AccordionControlElement)

        Dim TPCount As Integer = -1
        Dim TPRowCount As Integer = -1
        XtraTabPageCount = 0
        Dim TPColumnCount As Integer = -1
        Dim SettingSectionID As Integer = -1

        Dim LastBottom As Integer = 0

        Dim SectionControlsCumlHeight As Integer = 0

        SystemLog("Presentation object: " & DataPres.Name & " with section count " & DataPres.Sections.Length.ToString)

        XtraTabControlNewGIT.BeginUpdate()

        For Each Section In DataPres.Sections

            XtraTabPageCount += 1
            ReDim Preserve XtraTabPages(XtraTabPageCount)
            Dim CtlName As String = "XtraTabPage" & XtraTabPageCount.ToString
            Dim TPName As String = "TablePanel" & XtraTabPageCount.ToString


            Dim CurrTabPage As XtraTabPage = XtraTabControlNewGIT.Controls(CtlName)
            AbovoTabPagesCount += 1
            ReDim Preserve AbovoTabPages(AbovoTabPagesCount)
            AbovoTabPages(AbovoTabPagesCount) = New AbovoTabPage With {.TabPage = CurrTabPage}
            CurrentAbovoTabPage = AbovoTabPages(AbovoTabPagesCount)

            Dim CurrDummyTP As TablePanel = CurrTabPage.Controls(TPName)

            Application.AddMessageFilter(New ScrollRediverter(CurrTabPage))

            CurrTabPage.PageVisible = True
            CurrTabPage.AutoScroll = True
            CurrTabPage.Text = " " & Section.Name & " "

            SectionControlsCumlHeight = 0

            SettingSectionID += 1

            TPRowCount = -1

            SystemLog("Adding presentation section: " & Section.Name)

            TPCount += 1

            ReDim Preserve TPs(TPCount)

            TPs(TPCount) = New DevExpress.Utils.Layout.TablePanel With {.Padding = DefaultTablePanelPadding, .AutoScroll = True, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowOnly}
            TPs(TPCount).Dock = DockStyle.Fill
            'TPs(TPCount).be()
            TPs(TPCount).Columns.Clear()
            TPs(TPCount).Columns.Add(New TablePanelColumn(TablePanelEntityStyle.Relative, 2))
            TPs(TPCount).Columns.Add(New TablePanelColumn(TablePanelEntityStyle.Relative, 1))
            TPs(TPCount).Columns.Add(New TablePanelColumn(TablePanelEntityStyle.Relative, 5))
            TPs(TPCount).Columns.Add(New TablePanelColumn(TablePanelEntityStyle.Absolute, 100))

            GetSectionControlCollection(Section, Me, SettingSectionID, TPs(TPCount))


            TPs(TPCount).PerformLayout()
            'TPs(TPCount).EndInit()


            CurrDummyTP.Height = TPs(TPCount).Height
            CurrDummyTP.Width = 1

            SystemLog("Immediate Rturned Height =" & TPs(TPCount).Height.ToString)

            CurrTabPage.Controls.Add(TPs(TPCount))
            TPs(TPCount).Location = New Point With {.X = 0, .Y = 0}

            CurrTabPage.AutoScroll = True

            SystemLog("After insertion height TPHeight =" & TPs(TPCount).Height.ToString)

        Next

        ' If XtraTabPageCount > 6 Then XtraTabControlNewGIT.MultiLine = 1

        TPs(TPCount).Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 100))
        Debug.Print("XXXXXXXXX - Process of section complete - XXXXXXX")
        XtraTabControlNewGIT.EndUpdate()
        Debug.Print("XXXXXXXXX - EndUpdate Page complete - XXXXXXX")

        'XtraTabControlNewGIT.Invalidate()

        Debug.Print("XXXXXXXXX - Invalidate Page complete - XXXXXXX")
        EndTimer()







        '        If GridViewCount > -1 Then

        '            For Each GV In UsedGridVIEWS

        '                'GV.BeginInit()
        '                'GV.BeginUpdate()
        '                'Dim GVTag As GridViewTag = GV.Tag
        '                Dim GC As GridControl = GV.GridControl
        '                If GC.FocusedView IsNot GV Then GoTo NextGV
        '                'Debug.Print("XXXXXXXXX - checkloaded - XXXXXXX")
        '                'GV.CheckLoaded()
        '                'Debug.Print("XXXXXXXXX - column widths - XXXXXXX")
        '                'Formatter.ProcessGVColumWidths(GV, Me)
        '                'Debug.Print("XXXXXXXXX - best fit - XXXXXXX")
        '                'GV.BestFitColumns()

        '                Dim maxSize As New Size(Me.Width, Me.Height)
        '                'Debug.Print("XXXXXXXXX - calc best - XXXXXXX")
        '                GC.Size = GC.CalcBestSize(maxSize, False)
        '                GC.Height += CInt(GC.Height * 0.13)
        '                Debug.Print("XXXXXXXXX - layout - XXXXXXX")
        '                'GV.LayoutChanged()
        '                'Debug.Print("XXXXXXXXX - Invalidate - XXXXXXX")
        '                'GV.Invalidate()
        '                'GV.EndUpdate()
        'NextGV:

        '            Next

        '        End If
        '        If BandGridViewsCount > -1 Then

        '            For Each BGV In UsedBANDedGridVIEWS

        '                'GV.BeginInit()
        '                'GV.BeginUpdate()
        '                'Dim GVTag As GridViewTag = GV.Tag
        '                Dim GC As GridControl = BGV.GridControl
        '                If GC.FocusedView IsNot BGV Then GoTo NextBGV
        '                'Debug.Print("XXXXXXXXX - checkloaded - XXXXXXX")
        '                'GV.CheckLoaded()
        '                'Debug.Print("XXXXXXXXX - column widths - XXXXXXX")
        '                'Formatter.ProcessGVColumWidths(GV, Me)
        '                'Debug.Print("XXXXXXXXX - best fit - XXXXXXX")
        '                'GV.BestFitColumns()

        '                Dim maxSize As New Size(Me.Width, Me.Height)
        '                'Debug.Print("XXXXXXXXX - calc best - XXXXXXX")
        '                GC.Size = GC.CalcBestSize(maxSize, False)
        '                GC.Height += CInt(GC.Height * 0.13)
        '                Debug.Print("XXXXXXXXX - layout - XXXXXXX")
        '                'GV.LayoutChanged()
        '                'Debug.Print("XXXXXXXXX - Invalidate - XXXXXXX")
        '                'GV.Invalidate()
        '                'GV.EndUpdate()


        '            Next

        '        End If



    End Sub
    Sub GetSectionControlCollection(Section As PresentationSection, Parent As Object, SetSectionID As Integer, TP As TablePanel)

        Dim LastBottom As Integer = 0
        Dim ActiveDataSet As DataCellRange
        Dim TPRowCount As Integer = -1

        Dim SectionControlsCumlHeight As Integer = 0
        TP.BeginInit()
        TP.Dock = DockStyle.Fill
        TP.AutoSizeMode = AutoSizeMode.GrowAndShrink

        Dim DefPaddin As New Padding

        DefPaddin.Left = CInt(Me.Width * 0.01)

        TP.Padding = DefaultTablePanelPadding
        TP.Width = ParentGroupForm.Width - 40
        TP.BackColor = Color.White
        TP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.AutoSize, TP.Width))

        'TP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.Absolute, CInt(Me.Width * 0.25)))
        'TP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.AutoSize, (Me.Width * 0.75) - 40))

        'TODO AcContainers(AcContainersCount).Controls.Add(TP)

        For Each SectionElement In Section.SectionElements

            TPRowCount += 1

            TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 200))

            TP.Rows(TPRowCount).Tag = TPRowCount.ToString
#End Region
#Region "LiveGrid"

            If SectionElement.Type = "LiveGrid" Then

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                Dim ColList As List(Of String) = New List(Of String)
                Dim TypeList As List(Of String) = New List(Of String)

                For Each PresentedColumn In ActiveDataSet.DataColumns

                    TypeList.Add(PresentedColumn.ColumnTag.DataType)
                    ColList.Add(PresentedColumn.ColumnTag.ColumnHeading)

                Next

                Dim SetTag As New AbovoRangeDataSourceTag With {.GSID = GSID, .CSID = CSID, .ColList = ColList, .TypeList = TypeList, .RO = ActiveDataSet.RO, .DSIndex = SectionElement.ControlSourceIndex, .Worksheet = ActiveDataSet.SourceWorksheet, .DataRange = ActiveDataSet.DataRange}
                Dim AbovoRDS As New AbovoRangeDataSource()

                RangeDataSourceCount += 1
                ReDim Preserve RangeDataSources(RangeDataSourceCount)
                RangeDataSources(RangeDataSourceCount) = AbovoRDS.GetRangeDS(SetTag)

                SystemLog("Adding data from: " & ActiveDataSet.Name)
                SystemLog("Col count: " & ActiveDataSet.ColCount)
                SystemLog("Row count: " & ActiveDataSet.RowCount)

                GridCount += 1
                ReDim Preserve GridControls(GridCount)

                GridControls(GridCount) = New GridControl() With {
                    .Name = "GridControl_" & GridCount.ToString,
                    .Parent = Me,
                    .Dock = DockStyle.None,
                    .DataSource = RangeDataSources(RangeDataSourceCount)
                }

                CurrentAbovoTabPage.AddGrid(GridControls(GridCount))

                TP.Controls.Add(GridControls(GridCount))

                TP.SetColumnSpan(GridControls(GridCount), 3)
                TP.SetCell(GridControls(GridCount), TPRowCount, 0)

                GridControls(GridCount).ForceInitialize()

                ' GridControls(GridCount).BeginUpdate()

                Formatter.FormatGridControl(GridControls(GridCount))

                GridViewCount += 1
                ReDim Preserve UsedGridVIEWS(GridViewCount)
                UsedGridVIEWS(GridViewCount) = New DevExpress.XtraGrid.Views.Grid.GridView

                UsedGridVIEWS(GridViewCount).Tag = New GridViewTag With {.ModelID = ModelID, .DSID = RangeDataSourceCount, .DataSet = ActiveDataSet}

                UsedGridVIEWS(GridViewCount).BeginUpdate()

                SystemLog("Formatting Grid")

                UsedGridVIEWS(GridViewCount).ColumnPanelRowHeight += 2 * DefaultGridCellPadding

                UsedGridVIEWS(GridViewCount).UserCellPadding = New System.Windows.Forms.Padding(DefaultGridCellPadding)

                GridControls(GridCount).ViewCollection.Add(UsedGridVIEWS(GridViewCount))
                GridControls(GridCount).MainView = UsedGridVIEWS(GridViewCount)
                ColCount = -1

                Dim IdealGridHeight As Integer = 2 * (((UsedGridVIEWS(GridViewCount).RowHeight + (2 * DefaultGridCellPadding)) * RangeDataSources(RangeDataSourceCount).Count + 1))
                GridControls(GridCount).Height = IdealGridHeight

                SectionControlsCumlHeight += IdealGridHeight + (2 * DefaultTablePanelPadding.Left)

                TP.Controls.Add(GridControls(GridCount))

                TP.SetCell(GridControls(GridCount), TPRowCount, 0)

#End Region

#Region "Grid"

            ElseIf SectionElement.Type = "Grid" Then

                Dim ApplyBands As Boolean = DataPres.DataSets(SectionElement.ControlSourceIndex).HasBands
                Dim HasActions As Boolean = False
                Dim HasLockRules As Boolean = False
                Dim BandColumButtonsAndDrawing As CombinedColumnBandButton_HeaderFooterDrawer = Nothing
                Dim CustSummariesAddedToGrid As Boolean = False
                Dim BandedFooterDone As Boolean = False

                ColCount = -1

                FooterOn = False
                FooterDone = False

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                SystemLog("Adding data from: " & ActiveDataSet.Name)

                ColList = New List(Of String)

                PropertiesCount = -1 'reset
                UBSDataSourceCount += 1
                ReDim Preserve UnboundDataSources(UBSDataSourceCount)
                Dim SetTag As New AbovoUnboundSourceTag With {.GSID = GSID, .CSID = CSID, .RO = ActiveDataSet.RO, .DSIndex = SectionElement.ControlSourceIndex}

                UnboundDataSources(UBSDataSourceCount) = New AbovoUnboundSource(UBSDataSourceCount, SetTag)

                SystemLog("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
                SystemLog("Adding New ubs with index " & UBSDataSourceCount.ToString & ", SectionElement.ControlSourceIndex: " & SectionElement.ControlSourceIndex)
                SystemLog("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")

                Dim IdealGridWidth As Double = 0

                For Each PresentedColumn In ActiveDataSet.DataColumns

                    If PresentedColumn.ColumnTag.IsCalculated Then UnboundDataSources(UBSDataSourceCount).UBSTag.HasCalcs = True

                    ColCount += 1
                    ColName = "Col_" & ColCount.ToString
                    PresentedColumn.ColumnTag.ActiveColumnName = ColName
                    PropertiesCount += 1
                    ReDim Preserve PropertyArray(PropertiesCount)

                    Select Case PresentedColumn.ColumnTag.DataType

                        Case "S"
                            PropType = GetType(String)

                        Case "I", "Y", "D"
                            PropType = GetType(Integer)

                        Case "N", "P", "C", "M", "SM", "R"
                            PropType = GetType(Double)

                        Case "B"
                            PropType = GetType(Integer)

                        Case Else
                            PropType = GetType(String)

                    End Select

                    SystemLog("Adding column: " & PresentedColumn.ColumnTag.ColumnHeading)

                    PropertyArray(PropertiesCount) = New UnboundSourceProperty With {
                        .UserTag = PresentedColumn.ColumnTag,
                        .DisplayName = " " & PresentedColumn.ColumnTag.ColumnHeading & " ",
                        .Name = ColName,
                        .PropertyType = PropType
                    }

                    ColList.Add(ColName)

                Next

                PropertyList = PropertyArray

                UnboundDataSources(UBSDataSourceCount).Properties.AddRange(PropertyList)

                SystemLog("Adding handlers")
                'Add dataaccess/push handlers
                AddHandler UnboundDataSources(UBSDataSourceCount).ValueNeeded, AddressOf UnboundDS_ValueNeeded
                AddHandler UnboundDataSources(UBSDataSourceCount).ValuePushed, AddressOf UnboundDS_ValuePushed

                'Create new grid for DS

                GridCount += 1
                ReDim Preserve GridControls(GridCount)

                SystemLog("Adding grid control with datasource index " & UBSDataSourceCount.ToString)

                GridControls(GridCount) = New GridControl() With {
                    .Name = "GridControl_" & GridCount.ToString,
                    .Parent = Me,
                    .Dock = DockStyle.Top,
                    .DataSource = UnboundDataSources(UBSDataSourceCount)
                }

                GridControls(GridCount).ForceInitialize()

                CurrentAbovoTabPage.AddGrid(GridControls(GridCount))

                UnboundDataSources(UBSDataSourceCount).AttachedGrid = GridControls(GridCount)

                GridControls(GridCount).BeginUpdate()

                SystemLog("Initialising Grid")

                GridViewCount += 1
                ReDim Preserve UsedGridVIEWS(GridViewCount)

                UsedGridVIEWS(GridViewCount) = New DevExpress.XtraGrid.Views.Grid.GridView

                UsedGridVIEWS(GridViewCount).BeginUpdate()

                UsedGridVIEWS(GridViewCount).Tag = New GridViewTag With {.ModelID = ModelID, .DSID = UBSDataSourceCount, .DataSet = ActiveDataSet}

                If ActiveDataSet.RO Then

                    UsedGridVIEWS(GridViewCount).OptionsBehavior.Editable = False
                    UsedGridVIEWS(GridViewCount).OptionsClipboard.PasteMode = DevExpress.Export.PasteMode.None

                Else

                    UsedGridVIEWS(GridViewCount).OptionsClipboard.PasteMode = DevExpress.Export.PasteMode.Update

                End If

                GridControls(GridCount).RepositoryItems.Add(CustCalcEdit)

                If ActiveDataSet.HasValidations Then

                    SystemLog("Adding validations")

                    For Each ValList In ActiveDataSet.ValidationLists

                        GridControls(GridCount).RepositoryItems.Add(RepositaryItems.GetEditorFromList(ValList).RetCombo)

                    Next

                End If

                UsedGridVIEWS(GridViewCount).AccessibleName = "GridView_" & GridViewCount.ToString
                UsedGridVIEWS(GridViewCount).Name = "GridView_" & GridViewCount.ToString
                GridControls(GridCount).ViewCollection.Add(UsedGridVIEWS(GridViewCount))
                GridControls(GridCount).MainView = UsedGridVIEWS(GridViewCount)
                UsedGridVIEWS(GridViewCount).GridControl = GridControls(GridCount)

                SystemLog("Adding grid control with datasource index " & UBSDataSourceCount.ToString)

                Dim FillSize As Integer = ActiveDataSet.RowCount
                UnboundDataSources(UBSDataSourceCount).SetRowCount(FillSize)

                SystemLog("Assigning columns grid control with datasource " & UBSDataSourceCount.ToString)

                For Each ColCheck As UnboundSourceProperty In PropertyArray

                    With UsedGridVIEWS(GridViewCount).Columns(ColCheck.UserTag.ActiveColumnName)

                        .Tag = ColCheck
                        .AppearanceCell.options.usebackcolor = True
                        .AppearanceCell.Options.UseBorderColor = True
                        .AppearanceCell.BorderColor = System.Drawing.Color.Silver
                        .AppearanceCell.backcolor = System.Drawing.Color.Silver
                        .AppearanceCell.forecolor = System.Drawing.Color.Black

                    End With

                    If ColCheck.UserTag.IsReadOnly OrElse ColCheck.UserTag.IsCalculated Then

                        With UsedGridVIEWS(GridViewCount).Columns(ColCheck.UserTag.ActiveColumnName)

                            .OptionsColumn.TabStop = False
                            .OptionsColumn.ReadOnly = True
                            .AppearanceCell.BorderColor = System.Drawing.Color.White
                            .appearanceCell.backColor = System.Drawing.Color.White
                            .AppearanceCell.ForeColor = System.Drawing.Color.DarkGray

                        End With

                    End If

                    If ColCheck.UserTag.IsFixed Then

                        UsedGridVIEWS(GridViewCount).OptionsView.ColumnAutoWidth = False
                        UsedGridVIEWS(GridViewCount).Columns(ColCheck.UserTag.ActiveColumnName).Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left

                    End If

                Next



                GridControls(GridCount).Top = LastBottom + 20

                Dim ThisCol As Integer

                For ThisCol = 0 To UsedGridVIEWS(GridViewCount).Columns.Count - 1

                    UsedGridVIEWS(GridViewCount).Columns(ThisCol).Tag = ActiveDataSet.DataColumns(ThisCol).ColumnTag

                Next

                Dim ColTag As DataColumnTag
                Dim GVcolumn As GridColumn
                Dim siTotal As GridColumnSummaryItem

                Dim DefWidth As Integer

                'Add the "Insert Rows" button


                For Each GVcolumn In UsedGridVIEWS(GridViewCount).Columns

                    GVcolumn.Width += 2 * DefaultGridCellPadding
                    ColTag = GVcolumn.Tag
                    DefWidth = GVcolumn.Width

                    With GVcolumn

                        .Caption = " " & ColTag.ColumnHeading & " "
                        .ShowButtonMode = DevExpress.XtraGrid.Views.Base.ShowButtonModeEnum.ShowForFocusedCell
                        .AppearanceHeader.BackColor = System.Drawing.Color.White
                        .AppearanceHeader.Options.UseBackColor = True
                        .AppearanceHeader.ForeColor = AbovoBlue
                        .AppearanceHeader.Options.UseForeColor = True
                        .AppearanceHeader.TextOptions.VAlignment = VertAlignment.Bottom
                        .AppearanceHeader.BorderColor = System.Drawing.Color.White
                        .AppearanceHeader.Options.UseBorderColor = True

                    End With

                    Dim EditControl As RepositaryItems.AbovoRespositaryItem

                    ColTag.DefaultColumnWidth = DefWidth

                    If ColTag.IsDummyColumn Then

                        GVcolumn.OptionsColumn.ReadOnly = True

                        GVcolumn.MinWidth = 50
                        GVcolumn.MaxWidth = 50
                        GVcolumn.Width = 50
                        GVcolumn.AppearanceCell.BackColor = Color.White
                        GVcolumn.AppearanceCell.ForeColor = Color.White
                        GVcolumn.AppearanceHeader.ForeColor = Color.White
                        GVcolumn.AppearanceHeader.ForeColor = Color.White

                    End If

                    If ColTag.IsReadOnly OrElse ColTag.IsCalculated Then

                        GVcolumn.AppearanceCell.BackColor = Color.WhiteSmoke
                        GVcolumn.AppearanceCell.ForeColor = AbovoBlue

                    End If

                    Dim UseCombo As Boolean = False

                    If Not IsNothing(ColTag.RepositaryID) Then

                        ColTag.HasComboEdit = True
                        UseCombo = True
                        EditControl = RepositaryItems.GetEditor(ColTag.RepositaryID, ModelID)

                        If EditControl.RepType = "CMB" Then

                            GridCombosCount += 1
                            EditControl.RetCombo.Name = "Combo_" & GridCombosCount.ToString
                            AddHandler EditControl.RetCombo.Enter, AddressOf ComboOpen
                            GridControls(GridCount).RepositoryItems.Add(EditControl.RetCombo)
                            ComboReposClassesCount += 1

                            ReDim Preserve ComboReposClasses(ComboReposClassesCount)

                            ComboReposClasses(ComboReposClassesCount) = New AbovoGridRespoitaryCombo With {
                                .ID = ComboReposClassesCount,
                                .RepoistaryID = ColTag.RepositaryID,
                                .Combo = EditControl.RetCombo,
                                .ModelID = ModelID,
                                .GridID = GridCount}
                            CurrentAbovoTabPage.AddRepCombo(ComboReposClasses(ComboReposClassesCount))
                            GVcolumn.ColumnEdit = EditControl.RetCombo

                        End If

                    End If

                    If Not IsNothing(ColTag.TipText) Then

                        GVcolumn.ToolTip = TrimSpaces(ColTag.TipText) 'HelpManager.CreateSuperTooltip(ColTag.ColumnHeading, ColTag.TipText, "1")

                    End If

                    Select Case ColTag.DataType

                        Case "S"

                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.None
                            ColTag.ShowDefaultmask = -1

                        Case "M"

                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                            GVcolumn.DisplayFormat.FormatString = "c0"
                            Dim edit As New RepositoryItemTextEdit()
                            edit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
                            edit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
                            edit.Mask.EditMask = "c5"
                            GVcolumn.ColumnEdit = edit
                            ColTag.DefaultTextEditor = edit
                            ColTag.ShowDefaultmask = 0

                        Case "SM"

                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far

                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                            GVcolumn.DisplayFormat.FormatString = "c2"
                            Dim edit As New RepositoryItemTextEdit()
                            edit.UseMaskAsDisplayFormat = True
                            edit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
                            edit.Mask.EditMask = "c2"
                            GVcolumn.ColumnEdit = edit
                            ColTag.DefaultTextEditor = edit
                            ColTag.ShowDefaultmask = 1

                        Case "I"

                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                            GVcolumn.DisplayFormat.FormatString = "#,##0"
                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                            ColTag.ShowDefaultmask = -1

                            If Not UseCombo Then

                                Dim edit As New RepositoryItemTextEdit()
                                edit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
                                edit.Mask.EditMask = "n0"
                                GVcolumn.ColumnEdit = edit
                                ColTag.DefaultTextEditor = edit

                            End If

                        Case "C"

                            'GVcolumn.ColumnEdit = New RepositoryItemTextEdit With {
                            '        .AllowNullInput = True
                            '        }

                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                            GVcolumn.DisplayFormat.FormatString = "#,###,##0" '{#,###,##0;(#,###,##0)}"
                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far

                        Case "R"

                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                            GVcolumn.DisplayFormat.FormatString = "#,##0.00"
                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far

                            Dim edit As New RepositoryItemTextEdit()
                            edit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
                            edit.Mask.EditMask = "n5"
                            GVcolumn.ColumnEdit = edit
                            ColTag.DefaultTextEditor = edit

                        Case "P"

                            GVcolumn.Width += 2 * DefaultGridCellPadding
                            GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                            GVcolumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                            GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                            GVcolumn.DisplayFormat.FormatString = "p2"

                            ColTag.ShowDefaultmask = 3
                            Dim SetMinVal As Object = IIf(ColTag.MinVal Is Nothing, 0, IIf(ColTag.MinVal = "NOMIN", Nothing, CDec(ColTag.MinVal)))
                            Dim SetMaxVal As Object = IIf(ColTag.MaxVal Is Nothing, 1, IIf(ColTag.MaxVal = "NOMAX", Nothing, CDec(ColTag.MaxVal)))

                            If Not ColTag.IsReadOnly Then

                                Dim edit As New RepositoryItemSpinEdit()

                                edit.MinValue = IIf(SetMinVal = Nothing, Nothing, CDec(SetMinVal))
                                edit.Increment = IIf(ColTag.DefIncrement Is Nothing, CDec(0.0025), CDec(ColTag.DefIncrement))

                                edit.MaxValue = IIf(SetMaxVal = Nothing, Nothing, CDec(SetMaxVal))
                                edit.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
                                edit.Mask.EditMask = "p2"
                                edit.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                                edit.DisplayFormat.FormatString = "p2"
                                GVcolumn.ColumnEdit = edit
                                ColTag.DefaultTextEditor = edit

                            End If

                        Case "B"

                        Case Else



                    End Select

                    If ColTag.ShowSummary = "Sum" Then

                        If Not FooterDone Then SetFooterOn(UsedGridVIEWS(GridViewCount))

                        siTotal = New GridColumnSummaryItem
                        siTotal.SummaryType = DevExpress.Data.SummaryItemType.Sum
                        siTotal.DisplayFormat = "{0:n0}"

                        GVcolumn.Summary.Add(siTotal)

                    ElseIf ColTag.ShowSummary = "Count" Then

                        If Not FooterDone Then SetFooterOn(UsedGridVIEWS(GridViewCount))

                        siTotal = New GridColumnSummaryItem
                        siTotal.SummaryType = DevExpress.Data.SummaryItemType.Count
                        siTotal.DisplayFormat = "{0:n0}"
                        GVcolumn.Summary.Add(siTotal)

                    End If

                    If ColTag.HasActions Then

                        If HasActions = False Then

                            HasActions = True

                        End If

                        UsedGridVIEWS(GridViewCount).InvalidateColumnHeader(GVcolumn)

                    End If

                    IdealGridWidth += GVcolumn.Width

                Next

                With UsedGridVIEWS(GridViewCount)

                    .OptionsSelection.MultiSelect = True
                    .OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect

                    .OptionsBehavior.FocusLeaveOnTab = True
                    .OptionsBehavior.EditorShowMode = DevExpress.Utils.EditorShowMode.MouseDownFocused

                    .OptionsView.RowAutoHeight = True
                    .OptionsView.BestFitMaxRowCount = -1
                    '.OptionsView.ColumnAutoWidth = True
                    .OptionsView.BestFitMaxRowCount = ActiveDataSet.UsedRows

                    .ColumnPanelRowHeight += 4 * DefaultGridCellPadding
                    .UserCellPadding = New System.Windows.Forms.Padding(DefaultGridCellPadding)

                End With

                If UnboundDataSources(UBSDataSourceCount).UBSTag.RO Then UsedGridVIEWS(GridViewCount).OptionsBehavior.Editable = False

                Formatter.FormatGridView(UsedGridVIEWS(GridViewCount), GridControls(GridCount))

                LastBottom = GridControls(GridCount).Bottom
                SystemLog("Ending grid update")
                UsedGridVIEWS(GridViewCount).EndUpdate()
                GridControls(GridCount).Height += 100

#Region "Bands"

                If ApplyBands Then



                    SystemLog("Starting bands")

                    Dim LastBand As String = ""
                    Dim LastBandTitle As String = "StartingBandDummyName"
                    Dim OriginColTag As DataColumnTag = Nothing
                    Dim LastColour As Color = Color.Red
                    BandGridViewsCount += 1
                    ReDim Preserve UsedBANDedGridVIEWS(BandGridViewsCount)

                    UsedBANDedGridVIEWS(BandGridViewsCount) = New BandedGridView With {
                        .Tag = UsedGridVIEWS(GridViewCount).Tag,
                        .UserCellPadding = New System.Windows.Forms.Padding(DefaultGridCellPadding),
                        .HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Auto,
                        .VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Auto
                    }

                    BandColumButtonsAndDrawing = New CombinedColumnBandButton_HeaderFooterDrawer(UsedBANDedGridVIEWS(BandGridViewsCount), Me, SetSectionID)

                    UsedBANDedGridVIEWS(BandGridViewsCount).BeginUpdate()



                    With UsedBANDedGridVIEWS(BandGridViewsCount)

                        .OptionsSelection.MultiSelect = True
                        .OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect

                        .OptionsBehavior.FocusLeaveOnTab = True
                        .OptionsBehavior.EditorShowMode = DevExpress.Utils.EditorShowMode.MouseDownFocused

                        .OptionsView.RowAutoHeight = True
                        '.OptionsView.ColumnAutoWidth = True
                        .OptionsView.BestFitMaxRowCount = ActiveDataSet.UsedRows

                        .ColumnPanelRowHeight += 4 * DefaultGridCellPadding
                        .UserCellPadding = New System.Windows.Forms.Padding(DefaultGridCellPadding)

                        .OptionsCustomization.AllowChangeColumnParent = False
                        .BandPanelRowHeight += (4 * DefaultGridCellPadding)

                    End With

                    If ActiveDataSet.RO Then

                        UsedBANDedGridVIEWS(BandGridViewsCount).OptionsBehavior.Editable = False
                        UsedBANDedGridVIEWS(BandGridViewsCount).OptionsClipboard.PasteMode = DevExpress.Export.PasteMode.None

                    Else

                        UsedBANDedGridVIEWS(BandGridViewsCount).OptionsClipboard.PasteMode = DevExpress.Export.PasteMode.Update

                    End If

                    IdealGridWidth = 0

                    'Dim CurrBand As New GridBand

                    For Each GVcolumn In UsedGridVIEWS(GridViewCount).Columns

                        OriginColTag = GVcolumn.Tag

                        If OriginColTag.BandID <> LastBandTitle Then

                            BandGridViewBandsCount += 1
                            ReDim Preserve UsedBANDedGridViewBANDS(BandGridViewBandsCount)

                            Dim BandCaption As String = IIf(OriginColTag.BandID Is Nothing, " ", OriginColTag.BandID)

                            BandCaption = BandCaption.Replace("vblf", "<br>")

                            If OriginColTag.BandID Is Nothing Then

                                BandCaption = " "

                            End If

                            UsedBANDedGridViewBANDS(BandGridViewBandsCount) = New GridBand With {.Caption = BandCaption}

                            If Not IsNothing(OriginColTag.BandTipText) Then UsedBANDedGridViewBANDS(BandGridViewBandsCount).ToolTip = TrimSpaces(OriginColTag.BandTipText)

                            If OriginColTag.BandID = "FixedLeft" Then

                                UsedBANDedGridViewBANDS(BandGridViewBandsCount).Caption = " "
                                UsedBANDedGridViewBANDS(BandGridViewBandsCount).Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left

                            End If


                            With UsedBANDedGridViewBANDS(BandGridViewBandsCount)

                                .AppearanceHeader.ForeColor = AbovoBlue
                                .AppearanceHeader.Options.UseForeColor = True
                                .AppearanceHeader.BackColor = Color.White
                                .AppearanceHeader.Options.UseBackColor = True
                                .AppearanceHeader.Options.UseTextOptions = True
                                .AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center
                                .AppearanceHeader.TextOptions.VAlignment = VertAlignment.Bottom

                            End With

                            If LastColour = Color.Red Then

                                LastColour = Color.White

                            ElseIf LastColour = Color.White Then

                                LastColour = Color.Wheat

                            ElseIf LastColour = Color.Wheat Then

                                LastColour = AbovoBlue

                            ElseIf LastColour = AbovoBlue Then

                                LastColour = Color.Wheat

                            End If

                            Dim EditText As String


                            If OriginColTag.BandEditDescription IsNot Nothing Then

                                EditText = OriginColTag.BandEditDescription

                            Else

                                EditText = "Edit " & GVcolumn.FieldName & " of " & OriginColTag.BandID

                            End If

                            Dim ActToken As New ActionToken With {
                                    .ActionType = OriginColTag.EditRepNRHereExpansionMethod,
                                    .ActionStrData1 = OriginColTag.ActionNR,
                                    .ActionNR = OriginColTag.ActionNR,
                                    .ActionNumber1 = SetSectionID,
                                    .ActionStrData2 = EditText} ', .ActionData3 = ActiveDataSet.RepeatingHeaderText}

                            BandColumButtonsAndDrawing.AddCustomBandAddColumsButton(UsedBANDedGridViewBANDS(BandGridViewBandsCount), ActToken)

                            '           .HasActions = OriginColTag.HasActions,

                            Dim NewBandTag As New BandTag With {
                                    .ID = BandGridViewsCount,
                                    .ActionToken = ActToken,
                                    .ActionNR = OriginColTag.ActionNR,
                                    .ActionDescription = OriginColTag.BandEditDescription,
                                    .HighLightColour = LastColour}

                            If OriginColTag.BandID IsNot Nothing Then NewBandTag.DoBorder = True

                            UsedBANDedGridViewBANDS(BandGridViewBandsCount).Tag = NewBandTag

                            UsedBANDedGridVIEWS(BandGridViewsCount).Bands.Add(UsedBANDedGridViewBANDS(BandGridViewBandsCount))

                            LastBandTitle = OriginColTag.BandID

                        End If

                        BandGridViewColsCount += 1
                        ReDim Preserve UsedBANDedGridViewCOLS(BandGridViewColsCount)

                        If OriginColTag.HasActions Then

                            UsedBANDedGridViewBANDS(BandGridViewBandsCount).Tag.HasActions = True
                            UsedBANDedGridViewBANDS(BandGridViewBandsCount).Tag.ActionNR = OriginColTag.ActionNR

                        End If

                        UsedBANDedGridViewCOLS(BandGridViewColsCount) = New BandedGridColumn With {
                             .Tag = OriginColTag,
                            .FieldName = GVcolumn.FieldName,
                            .Width = GVcolumn.Width,
                            .Visible = True,
                            .ColumnEdit = GVcolumn.ColumnEdit,
                            .Fixed = GVcolumn.Fixed,
                            .Caption = GVcolumn.Caption,
                            .SortIndex = GVcolumn.SortIndex
                            }

                        If GVcolumn.Fixed Then

                            UsedBANDedGridViewCOLS(BandGridViewColsCount).Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left
                            Debug.Print("'''''''''''''FixdCol_" & GVcolumn.AbsoluteIndex)

                        End If
                        UsedBANDedGridViewCOLS(BandGridViewColsCount).AppearanceCell.TextOptions.HAlignment = GVcolumn.AppearanceCell.TextOptions.HAlignment

                        With UsedBANDedGridViewCOLS(BandGridViewColsCount)

                            .AppearanceCell.Options.UseBackColor = True
                            .AppearanceCell.Options.UseBorderColor = True
                            .AppearanceCell.BorderColor = GVcolumn.AppearanceCell.BorderColor
                            .AppearanceCell.BackColor = GVcolumn.AppearanceCell.BackColor

                        End With

                        With UsedBANDedGridViewCOLS(BandGridViewColsCount)

                            .AppearanceHeader.ForeColor = AbovoBlue
                            .AppearanceHeader.Options.UseForeColor = True
                            .AppearanceHeader.BackColor = Color.White
                            .AppearanceHeader.Options.UseBackColor = True

                            .AppearanceHeader.TextOptions.HAlignment = GVcolumn.AppearanceHeader.TextOptions.HAlignment
                            .AppearanceHeader.TextOptions.VAlignment = VertAlignment.Bottom
                            .AppearanceHeader.Options.UseTextOptions = True

                            .DisplayFormat.FormatType = GVcolumn.DisplayFormat.FormatType
                            .DisplayFormat.FormatString = GVcolumn.DisplayFormat.FormatString

                            .ToolTip = GVcolumn.ToolTip
                            .OptionsColumn.ReadOnly = GVcolumn.ReadOnly

                        End With

                        If OriginColTag.ShowSummary = "Sum" Then

                            If Not BandedFooterDone Then

                                BandedFooterDone = True
                                UsedBANDedGridVIEWS(BandGridViewsCount).OptionsView.ShowFooter = True
                                UsedBANDedGridVIEWS(BandGridViewsCount).FooterPanelHeight += (2 * DefaultGridCellPadding)

                            End If

                            siTotal = New GridColumnSummaryItem
                            siTotal.SummaryType = DevExpress.Data.SummaryItemType.Sum
                            siTotal.DisplayFormat = "{0:n0}"

                            UsedBANDedGridViewCOLS(BandGridViewColsCount).Summary.Add(siTotal)

                        ElseIf OriginColTag.ShowSummary = "Count" Then

                            If Not FooterDone Then SetFooterOn(UsedGridVIEWS(GridViewCount))

                            siTotal = New GridColumnSummaryItem
                            siTotal.SummaryType = DevExpress.Data.SummaryItemType.Count
                            siTotal.DisplayFormat = "{0:n0}"
                            UsedBANDedGridViewCOLS(BandGridViewColsCount).Summary.Add(siTotal)

                            If Not BandedFooterDone Then

                                BandedFooterDone = True
                                UsedBANDedGridVIEWS(BandGridViewsCount).OptionsView.ShowFooter = True
                                UsedBANDedGridVIEWS(BandGridViewsCount).FooterPanelHeight += (2 * DefaultGridCellPadding)

                            End If

                        End If

                        UsedBANDedGridVIEWS(BandGridViewsCount).Columns.Add(UsedBANDedGridViewCOLS(BandGridViewColsCount))
                        UsedBANDedGridViewBANDS(BandGridViewBandsCount).Columns.Add(UsedBANDedGridViewCOLS(BandGridViewColsCount))

                        IdealGridWidth += GVcolumn.Width

                    Next

                    LastBandTitle = "LastBandDummyName"

                    Dim EditControlInitialised As Boolean = False
                    Dim InEditorAddingMode As Boolean = False
                    Dim DoneFirstItem As Boolean = False
                    Dim MaxItemLength As Integer = 5
                    Dim SetIndex As Integer = -1
                    Dim SetColIndex As Integer = -1
                    Dim CurrBaseCaption As String = ""

                    Formatter.FormatGridView(UsedBANDedGridVIEWS(BandGridViewsCount), GridControls(GridCount))

                    For Each BGC As BandedGridColumn In UsedBANDedGridVIEWS(BandGridViewsCount).Columns

                        OriginColTag = BGC.Tag

                        If OriginColTag.BandID <> LastBandTitle Then

                            SetIndex += 1
                            SetColIndex = 0

                            EditControlInitialised = False
                            InEditorAddingMode = False

                            If OriginColTag.EditRepNRHere Then
                                InEditorAddingMode = True
                            Else
                                InEditorAddingMode = False
                            End If

                            DoneFirstItem = False
                            MaxItemLength = 5

                        End If

                        If InEditorAddingMode Then

                            If OriginColTag.EditRepNRHereROColumn Then

                                'With BGC

                                '    .AppearanceHeader.Options.UseBackColor = True
                                '    .AppearanceHeader.BackColor = AbovoComboBGC
                                '    .AppearanceHeader.Options.UseForeColor = True
                                '    .AppearanceHeader.ForeColor = Color.White 'AbovoComboBGC

                                'End With

                            ElseIf OriginColTag.BandID Is Nothing Then

                                'do nothing - read only columns cannot have in-header editors

                            Else

                                If UCase(OriginColTag.EditRepNRHereEditor) = "COMBO" Then

                                    InHeaderCombosCount += 1
                                    ReDim Preserve InHeaderCombos(InHeaderCombosCount)

                                    Dim EditControl As New AbovoDEHeaderComboBox
                                    EditControl.AddBlankFirstItem = True
                                    EditControl.InitialiseStandard(OriginColTag.EditRepNRHereComboRepository)
                                    EditControl.EditValue = OriginColTag.EditRepNRHereInitialValue

                                    Dim RetComb As RepositoryItemComboBox = EditControl.Properties

                                    Dim InColumnEditorTag As New InColumnEditorTagCombo With {
                                        .EditingNRIndexPosition = OriginColTag.EditNRIndexPosition,
                                        .EditingNRName = OriginColTag.RepeatingNR,
                                        .NROrientation = IIf(OriginColTag.EditRepNRNROrientation = "PORT", Orientation.Vertical, Orientation.Horizontal),
                                        .EditorType = "COMBO",
                                        .InitialValue = OriginColTag.EditRepNRHereInitialValue,
                                        .LastEditorValue = OriginColTag.EditRepNRHereInitialValue,
                                        .EditorFormat = OriginColTag.EditRepNRHereDataFormat,
                                        .LinkedComboBoxEdit = EditControl,
                                        .ParentBandedGridColumn = BGC
                                        }

                                    If Not DoneFirstItem Then

                                        MaxItemLength = GetMaxListItemLength(EditControl.RenderedListItems)
                                        CurrBaseCaption = New String("_", MaxItemLength + 2)
                                        DoneFirstItem = True

                                    End If



                                    RetComb.AccessibleName = "InheaderCombo_" & InHeaderCombosCount.ToString

                                    If OriginColTag.EditRepNRHereDataFormat = "S" Then

                                        RetComb.AppearanceDropDown.TextOptions.HAlignment = HorzAlignment.Center
                                        RetComb.Appearance.TextOptions.HAlignment = HorzAlignment.Center
                                        BGC.Width = BGC.Width * 2

                                    Else

                                        RetComb.AppearanceDropDown.TextOptions.HAlignment = HorzAlignment.Far
                                        RetComb.AppearanceDropDown.Font = EditControl.Font
                                        RetComb.Appearance.TextOptions.HAlignment = HorzAlignment.Far
                                        OriginColTag.ColWidthMultiplier = 1.33

                                    End If

                                    With BGC

                                        .AppearanceHeader.Options.UseBackColor = True
                                        .AppearanceHeader.BackColor = AbovoComboBGC
                                        .AppearanceHeader.Options.UseForeColor = True
                                        .AppearanceHeader.ForeColor = Color.White 'AbovoComboBGC
                                        .Caption = CurrBaseCaption

                                    End With

                                    With RetComb

                                        .Appearance.Options.UseBorderColor = True
                                        .Appearance.BorderColor = AbovoComboBGC
                                        .Appearance.Options.UseTextOptions = True
                                        .Appearance.Options.UseBackColor = True
                                        .Appearance.BackColor = AbovoComboBGC
                                        .Appearance.ForeColor = Color.White

                                        .Appearance.Options.UseFont = True
                                        .Appearance.Font = New Font("Segoe UI", 12, FontStyle.Regular)
                                        '.Appearance.FontStyleDelta = FontStyle.Bold
                                        .Tag = InColumnEditorTag

                                    End With

                                    Dim DefAppearance As AppearanceObject = RetComb.Appearance

                                    RetComb.AppearanceDisabled.Assign(DefAppearance)
                                    RetComb.AppearanceFocused.Assign(DefAppearance)
                                    RetComb.AppearanceReadOnly.Assign(DefAppearance)

                                    'GridControls(GridCount).RepositoryItems.Add(RetComb)

                                    AddHandler RetComb.EditValueChanged, AddressOf ColumnHeaderEmbededComboChanged
                                    OriginColTag.InColumnEditorCombo = EditControl
                                    OriginColTag.HasIncolumnEditor = True
                                    Debug.Print("Adding ICH to column " & BGC.AbsoluteIndex)
                                    Dim helper As ColumnInplaceEditorHelper = New ColumnInplaceEditorHelper(BGC, RetComb)
                                    InColumnEditorTag.InPlaceColumnHelper = helper
                                    EditControl.Tag = InColumnEditorTag
                                    helper.Tag = InColumnEditorTag
                                    helper.EditValue = OriginColTag.EditRepNRHereInitialValue
                                    helper.LinkedComboBoxEdit = EditControl
                                    EditControl.InPlaceColumnHelper = helper


                                ElseIf UCase(OriginColTag.EditRepNRHereEditor) = "DATE" Then

                                    InHeaderCombosCount += 1
                                    ReDim Preserve InHeaderCombos(InHeaderCombosCount)

                                    Dim EditControl As New AbovoDEHeaderDateBox
                                    EditControl.AddBlankFirstItem = True
                                    EditControl.InitialiseStandard(OriginColTag.EditRepNRHereComboRepository)
                                    EditControl.EditValue = OriginColTag.EditRepNRHereInitialValue

                                    Dim RetComb As RepositoryItemDateEdit = EditControl.Properties

                                    Dim InColumnEditorTag As New InColumnEditorTagDateEdit With {
                                        .EditingNRIndexPosition = OriginColTag.EditNRIndexPosition,
                                        .EditingNRName = OriginColTag.RepeatingNR,
                                        .NROrientation = IIf(OriginColTag.EditRepNRNROrientation = "PORT", Orientation.Vertical, Orientation.Horizontal),
                                        .EditorType = "DATE",
                                        .InitialValue = OriginColTag.EditRepNRHereInitialValue,
                                        .LastEditorValue = OriginColTag.EditRepNRHereInitialValue,
                                        .EditorFormat = OriginColTag.EditRepNRHereDataFormat,
                                        .LinkedDateBoxEdit = EditControl,
                                        .ParentBandedGridColumn = BGC
                                        }





                                    RetComb.AccessibleName = "InheaderCombo_" & InHeaderCombosCount.ToString

                                    If OriginColTag.EditRepNRHereDataFormat = "S" Then

                                        RetComb.AppearanceDropDown.TextOptions.HAlignment = HorzAlignment.Center
                                        RetComb.Appearance.TextOptions.HAlignment = HorzAlignment.Center
                                        BGC.Width = BGC.Width * 2

                                    Else

                                        RetComb.AppearanceDropDown.TextOptions.HAlignment = HorzAlignment.Far
                                        RetComb.AppearanceDropDown.Font = EditControl.Font
                                        RetComb.Appearance.TextOptions.HAlignment = HorzAlignment.Far
                                        OriginColTag.ColWidthMultiplier = 1.33

                                    End If

                                    With BGC

                                        .AppearanceHeader.Options.UseBackColor = True
                                        .AppearanceHeader.BackColor = AbovoComboBGC
                                        .AppearanceHeader.Options.UseForeColor = True
                                        .AppearanceHeader.ForeColor = Color.White 'AbovoComboBGC
                                        .Caption = CurrBaseCaption

                                    End With

                                    With RetComb

                                        .Appearance.Options.UseBorderColor = True
                                        .Appearance.BorderColor = AbovoComboBGC
                                        .Appearance.Options.UseTextOptions = True
                                        .Appearance.Options.UseBackColor = True
                                        .Appearance.BackColor = AbovoComboBGC
                                        .Appearance.ForeColor = Color.White

                                        .Appearance.Options.UseFont = True
                                        .Appearance.Font = New Font("Segoe UI", 12, FontStyle.Regular)
                                        '.Appearance.FontStyleDelta = FontStyle.Bold
                                        .Tag = InColumnEditorTag

                                    End With

                                    Dim DefAppearance As AppearanceObject = RetComb.Appearance

                                    RetComb.AppearanceDisabled.Assign(DefAppearance)
                                    RetComb.AppearanceFocused.Assign(DefAppearance)
                                    RetComb.AppearanceReadOnly.Assign(DefAppearance)

                                    'GridControls(GridCount).RepositoryItems.Add(RetComb)

                                    AddHandler RetComb.EditValueChanged, AddressOf ColumnHeaderEmbededDateEChanged
                                    OriginColTag.InColumnEditorDate = EditControl
                                    OriginColTag.HasIncolumnEditor = True
                                    Debug.Print("Adding ICH to column " & BGC.AbsoluteIndex)
                                    Dim helper As ColumnInplaceEditorHelper = New ColumnInplaceEditorHelper(BGC, RetComb)
                                    InColumnEditorTag.InPlaceColumnHelper = helper
                                    EditControl.Tag = InColumnEditorTag
                                    helper.Tag = InColumnEditorTag
                                    helper.EditValue = OriginColTag.EditRepNRHereInitialValue
                                    helper.LinkedDateEdit = EditControl
                                    EditControl.InPlaceColumnHelper = helper

                                End If
                            End If

                        End If

                    Next

                    BandExtendersCount += 1

                    'Dim Actor As New ActionToken With {.ActionType = "NREDIT", .ActionStrData1 = OriginColTag.ActionNR, .ActionNumber1 = SetSectionID, .ActionStrData2 = OriginColTag.BandEditDescription} ', .ActionData3 = ActiveDataSet.RepeatingHeaderText}
                    'If Not DontAddCDHDef Then AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).CustomDrawColumnHeader, AddressOf DefaultCustomDrawColumnHeader
                    'AddHandler UsedGridVIEWS(GridViewCount).CustomDrawColumnHeader, AddressOf DefaultCustomDrawColumnHeader
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).CustomDrawCell, AddressOf GridView_CustomDrawCell
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).ValidatingEditor, AddressOf GridView_ValidatingEditor
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).ShowingEditor, AddressOf GridView_ShowingEditor
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).Click, AddressOf GridView_Event_SingleClick
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).CustomRowCellEdit, AddressOf GridView_CustomCellEditor
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).CustomRowCellEditForEditing, AddressOf GridView_RowCellEditForEditing
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).CalcRowHeight, AddressOf GridView_CalcBandedGridRowHeight
                    AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).ClipboardRowPasting, AddressOf GVPasting
                    UsedBANDedGridVIEWS(BandGridViewsCount).VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Auto

                    UsedBANDedGridVIEWS(BandGridViewsCount).EndUpdate()

                    UsedBANDedGridVIEWS(BandGridViewsCount).Name = "BandGridView_" & BandGridViewsCount.ToString
                    UsedBANDedGridVIEWS(BandGridViewsCount).AccessibleName = "BandGridView_" & BandGridViewsCount.ToString

                    GridControls(GridCount).ViewCollection.Add(UsedBANDedGridVIEWS(BandGridViewsCount))
                    GridControls(GridCount).MainView = UsedBANDedGridVIEWS(BandGridViewsCount)
                    UsedBANDedGridVIEWS(BandGridViewsCount).GridControl = GridControls(GridCount)



                    UsedBANDedGridVIEWS(BandGridViewsCount).OptionsView.BestFitMaxRowCount = ActiveDataSet.UsedRows
                    UnboundDataSources(UBSDataSourceCount).ActiveGridBandedView = UsedBANDedGridVIEWS(BandGridViewsCount)
                    UnboundDataSources(UBSDataSourceCount).InBandedMode = True

                Else

                    UnboundDataSources(UBSDataSourceCount).ActiveGridView = UsedGridVIEWS(GridViewCount)

                    'If Not DontAddCDHDef Then AddHandler UsedGridVIEWS(GridViewCount).CustomDrawColumnHeader, AddressOf DefaultCustomDrawColumnHeader

                    AddHandler UsedGridVIEWS(GridViewCount).CustomDrawCell, AddressOf GridView_CustomDrawCell
                    AddHandler UsedGridVIEWS(GridViewCount).ValidatingEditor, AddressOf GridView_ValidatingEditor
                    AddHandler UsedGridVIEWS(GridViewCount).ShowingEditor, AddressOf GridView_ShowingEditor
                    AddHandler UsedGridVIEWS(GridViewCount).ShownEditor, AddressOf GridView_ShownEditor
                    AddHandler UsedGridVIEWS(GridViewCount).CustomRowCellEdit, AddressOf GridView_CustomCellEditor
                    AddHandler UsedGridVIEWS(GridViewCount).Click, AddressOf GridView_Event_SingleClick
                    AddHandler UsedGridVIEWS(GridViewCount).CustomRowCellEditForEditing, AddressOf GridView_RowCellEditForEditing
                    ' AddHandler UsedGridVIEWS(GridViewCount).ColumnWidthChanged, AddressOf GridView_ColumnWidthChanged
                    AddHandler UsedGridVIEWS(GridViewCount).ClipboardRowPasting, AddressOf GVPasting
                    ExpanableGridControlsCount += 1
                    ReDim Preserve ExpanableGridControls(ExpanableGridControlsCount)
                    ExpanableGridControls(ExpanableGridControlsCount) = GridControls(GridCount)

                    UsedGridVIEWS(GridViewCount).VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Auto

                End If
#End Region



                If Not ApplyBands Then

                    BandColumButtonsAndDrawing = New CombinedColumnBandButton_HeaderFooterDrawer(UsedGridVIEWS(GridViewCount), Me, SetSectionID)

                End If



                If ActiveDataSet.RowExpandsByModel = "NRRI" OrElse ActiveDataSet.RowExpandsByModel = "NRCI" Then

                    FooterOn = True

                    Dim Actor As New ActionToken With {.ActionType = ActiveDataSet.RowExpandsByModel, .ActionStrData1 = ActiveDataSet.RowExpandByNR, .ActionNumber1 = SetSectionID, .ActionDescription = "Edit " & ActiveDataSet.RowExpandByNR}

                    Dim itemCust As New GridColumnSummaryItem
                    itemCust.SummaryType = DevExpress.Data.SummaryItemType.Custom

                    If ApplyBands Then

                        itemCust.FieldName = UsedBANDedGridVIEWS(BandGridViewsCount).Columns(0).FieldName
                        UsedBANDedGridVIEWS(BandGridViewsCount).Columns(0).Summary.Add(itemCust)
                        BandColumButtonsAndDrawing.AddCustomAddRowsButton(UsedBANDedGridVIEWS(BandGridViewsCount).Columns(0), Actor)
                        Dim Col1Tag As DataColumnTag = UsedBANDedGridVIEWS(BandGridViewsCount).Columns(0).Tag
                        Col1Tag.HasIncolumnButton = True
                        If Not CustSummariesAddedToGrid Then
                            'AddHandler UsedBANDedGridVIEWS(BandGridViewsCount).CustomSummaryCalculate, AddressOf GridView_CustomSummaryCalculate
                            CustSummariesAddedToGrid = True
                        End If

                        If Not BandedFooterDone Then

                            BandedFooterDone = True
                            UsedBANDedGridVIEWS(BandGridViewsCount).OptionsView.ShowFooter = True
                            UsedBANDedGridVIEWS(BandGridViewsCount).FooterPanelHeight += (2 * DefaultGridCellPadding)

                        End If

                    Else


                        itemCust.FieldName = UsedGridVIEWS(GridViewCount).Columns(0).FieldName
                        UsedGridVIEWS(GridViewCount).Columns(0).Summary.Add(itemCust)
                        BandColumButtonsAndDrawing.AddCustomAddRowsButton(UsedGridVIEWS(GridViewCount).Columns(0), Actor)
                        Dim Col1Tag As DataColumnTag = UsedGridVIEWS(GridViewCount).Columns(0).Tag
                        Col1Tag.HasIncolumnButton = True
                        If Not CustSummariesAddedToGrid Then
                            'AddHandler UsedGridVIEWS(GridViewCount).CustomSummaryCalculate, AddressOf GridView_CustomSummaryCalculate
                            CustSummariesAddedToGrid = True
                        End If

                        If Not FooterDone Then
                            FooterOn = True
                            SetFooterOn(UsedGridVIEWS(GridViewCount), False)
                        End If

                    End If

                    DontAddCDHDef = True

                End If

                AddHandler GridControls(GridCount).DoubleClick, AddressOf VerifyDoubleClick
                SystemLog("GridHeight: " & GridControls(GridCount).Height.ToString)

                Dim IdealGridHeight As Integer

                If ApplyBands Then

                    Dim GVI As GridViewInfo = UsedBANDedGridVIEWS(BandGridViewsCount).GetViewInfo()
                    Dim HeaderHeight As Integer = GVI.ColumnRowHeight
                    Dim FooterHeight As Integer = GVI.FooterCellHeight
                    IdealGridHeight = (IdealGridRowHeight + (2 * DefaultGridCellPadding)) * (ActiveDataSet.RowCount) + HeaderHeight + FooterHeight

                Else

                    Dim GVI As GridViewInfo = UsedGridVIEWS(GridViewCount).GetViewInfo()
                    Dim HeaderHeight As Integer = GVI.ColumnRowHeight
                    Dim FooterHeight As Integer = GVI.FooterCellHeight
                    IdealGridHeight = (IdealGridRowHeight + (2 * DefaultGridCellPadding)) * (ActiveDataSet.RowCount) + HeaderHeight + FooterHeight

                End If

                If IdealGridHeight > MaxGridHeight Then

                    IdealGridHeight = MaxGridHeight

                End If

                If SectionElement.GridControls IsNot Nothing Then

                    Dim LastRight As Integer = 0
                    Dim ButtText As String = ""
                    Dim NewButton As DevExpress.XtraEditors.SimpleButton = Nothing
                    Dim Layout As LayoutControl = New LayoutControl()
                    Layout.BackColor = Color.White

                    Layout.OptionsView.EnableTransparentBackColor = False
                    Layout.Padding = DefaultTablePanelPadding

                    Layout.Dock = DockStyle.Fill

                    Layout.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
                    Layout.LookAndFeel.UseDefaultLookAndFeel = False
                    Layout.OptionsView.EnableTransparentBackColor = True
                    Layout.Root.AppearanceGroup.BackColor = System.Drawing.Color.Transparent
                    Layout.Root.AppearanceGroup.Options.UseBackColor = True

                    TP.Controls.Add(Layout)
                    TP.SetCell(Layout, TPRowCount, 0)

                    Dim groupButtons As LayoutControlGroup = Layout.Root.AddGroup()
                    groupButtons.Name = "GroupButtons"
                    groupButtons.GroupBordersVisible = False
                    groupButtons.AppearanceGroup.BackColor = Color.White
                    groupButtons.AppearanceGroup.Options.UseBackColor = True

                    Dim lastItem As LayoutControlItem = Nothing
                    Dim DoneFirst As Boolean = False

                    For Each GridButtControl In SectionElement.GridControls

                        NewButton = New DevExpress.XtraEditors.SimpleButton()

                        Dim itemButton As LayoutControlItem = Nothing

                        If Not DoneFirst Then

                            itemButton = Layout.AddItem(groupButtons, InsertType.Left)

                            DoneFirst = True
                            lastItem = itemButton

                        Else

                            itemButton = Layout.AddItem(lastItem, InsertType.Right)
                            DoneFirst = True
                            lastItem = itemButton

                        End If

                        ButtText = GridButtControl.CommandText.Replace("vblf", "<br>")

                        NewButton.AllowHtmlDraw = DefaultBoolean.True
                        NewButton.Appearance.TextOptions.WordWrap = WordWrap.Wrap
                        NewButton.Text = ButtText
                        NewButton.ToolTip = GridButtControl.CommandTip

                        GridButtControl.AttachedGrid = GridControls(GridCount)

                        NewButton.Tag = GridButtControl

                        Dim bestSize As Size = NewButton.CalcBestSize()
                        NewButton.Width = bestSize.Width
                        NewButton.Height = bestSize.Height

                        AddHandler NewButton.Click, AddressOf Grid_ButtonClick

                        itemButton.Control = NewButton

                    Next

                    Dim emptySpace As New EmptySpaceItem
                    emptySpace.Parent = groupButtons

                    SectionControlsCumlHeight += NewButton.Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                    TPRowCount += 1

                    TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))

                    TP.Rows(TPRowCount).Tag = TPRowCount.ToString

                End If

                GridControls(GridCount).Height = IdealGridHeight

                If IdealGridWidth > TP.Width - 50 Then IdealGridWidth = TP.Width - 50

                TP.Controls.Add(GridControls(GridCount))
                TP.SetCell(GridControls(GridCount), TPRowCount, 0)
                TP.SetColumnSpan(GridControls(GridCount), 4)

                SectionControlsCumlHeight += IdealGridHeight + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)
                TP.AutoSize = True
                TP.AutoSizeMode = AutoSizeMode.GrowAndShrink

                Debug.Print("XXXXXXXXXXXXXX Grid process complete XXXXXXXXXXXXXXXXXX")
                'GridControls(GridCount).Refresh()


#End Region

#Region "VGrid"

            ElseIf SectionElement.Type = "VGrid" Then

                Dim ApplyBands As Boolean = DataPres.DataSets(SectionElement.ControlSourceIndex).HasBands
                Dim HasActions As Boolean = False
                Dim HasLockRules As Boolean = False
                Dim BandColumButtonsAndDrawing As CombinedColumnBandButton_HeaderFooterDrawer = Nothing
                Dim CustSummariesAddedToGrid As Boolean = False
                Dim BandedFooterDone As Boolean = False

                ColCount = -1

                FooterOn = False
                FooterDone = False

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                SystemLog("Adding data from: " & ActiveDataSet.Name)

                ColList = New List(Of String)

                PropertiesCount = -1 'reset
                UBSDataSourceCount += 1
                ReDim Preserve UnboundDataSources(UBSDataSourceCount)
                Dim SetTag As New AbovoUnboundSourceTag With {.GSID = GSID, .CSID = CSID, .RO = ActiveDataSet.RO, .DSIndex = SectionElement.ControlSourceIndex}

                UnboundDataSources(UBSDataSourceCount) = New AbovoUnboundSource(UBSDataSourceCount, SetTag)

                SystemLog("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
                SystemLog("Adding New ubs with index " & UBSDataSourceCount.ToString & ", SectionElement.ControlSourceIndex: " & SectionElement.ControlSourceIndex)
                SystemLog("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")

                Dim IdealGridWidth As Double = 0

                Dim ColTag As DataColumnTag = Nothing


                For Each PresentedColumn In ActiveDataSet.DataColumns

                    ColTag = PresentedColumn.ColumnTag

                    If ColTag.IsCalculated Then UnboundDataSources(UBSDataSourceCount).UBSTag.HasCalcs = True

                    ColCount += 1
                    ColName = "Col_" & ColCount.ToString
                    ColTag.ActiveColumnName = ColName
                    PropertiesCount += 1
                    ReDim Preserve PropertyArray(PropertiesCount)

                    Select Case ColTag.DataType

                        Case "S"
                            PropType = GetType(String)

                        Case "I", "Y", "D"
                            PropType = GetType(Integer)

                        Case "N", "P", "C", "M", "SM", "R"
                            PropType = GetType(Double)

                        Case "B"
                            PropType = GetType(Integer)

                        Case Else
                            PropType = GetType(String)

                    End Select

                    SystemLog("Adding column: " & ColTag.ColumnHeading)

                    PropertyArray(PropertiesCount) = New UnboundSourceProperty With {
                        .UserTag = PresentedColumn.ColumnTag,
                        .DisplayName = " " & PresentedColumn.ColumnTag.ColumnHeading & " ",
                        .Name = ColName,
                        .PropertyType = PropType
                    }

                    ColList.Add(ColName)

                Next

                PropertyList = PropertyArray

                UnboundDataSources(UBSDataSourceCount).Properties.AddRange(PropertyList)
                UnboundDataSources(UBSDataSourceCount).InVertMode = True

                SystemLog("Adding handlers")
                'Add dataaccess/push handlers
                AddHandler UnboundDataSources(UBSDataSourceCount).ValueNeeded, AddressOf UnboundDS_ValueNeeded
                AddHandler UnboundDataSources(UBSDataSourceCount).ValuePushed, AddressOf UnboundDS_ValuePushed

                'Create new grid for DS

                VertGridCount += 1
                ReDim Preserve VertGridControls(VertGridCount)


                SystemLog("Adding grid control with datasource index " & UBSDataSourceCount.ToString)

                VertGridControls(VertGridCount) = New VGridControl() With {
                    .Name = "VGridControl_" & GridCount.ToString,
                    .Parent = Me,
                    .Dock = DockStyle.Top
                }

                Dim VertGrid As VGridControl = VertGridControls(VertGridCount)

                VertGrid.DataSource = UnboundDataSources(UBSDataSourceCount)

                Dim FillSize As Integer = ActiveDataSet.RowCount
                UnboundDataSources(UBSDataSourceCount).SetRowCount(FillSize)
                VertGrid.ForceInitialize()


                Dim LastBandID As String = Nothing
                Dim CatRow As CategoryRow = Nothing
                Dim IndexPos As Integer = 0

                'The rows have already been generated by ForceInitialize().
                'Find each row by its bound field name rather than relying upon Rows(0).
                For Each PresentedColumn In ActiveDataSet.DataColumns

                    ColTag = PresentedColumn.ColumnTag

                    '-------------------------------------------------------------
                    ' Create category/band
                    '-------------------------------------------------------------
                    If Not String.Equals(ColTag.BandID, LastBandID) Then

                        Dim BandCaption As String =
            If(String.IsNullOrEmpty(ColTag.BandID), " ", ColTag.BandID)

                        CatRow = New CategoryRow("Category_" & IndexPos.ToString) With {
            .Height = 40
        }

                        CatRow.Properties.Caption = BandCaption

                        VertGrid.Rows.Add(CatRow)

                        LastBandID = ColTag.BandID

                    End If

                    '-------------------------------------------------------------
                    ' Find the auto-created row belonging to this data property
                    '-------------------------------------------------------------
                    Dim VRow As EditorRow =
        TryCast(
            VertGrid.Rows.GetRowByFieldName(
                ColTag.ActiveColumnName,
                True),
            EditorRow)

                    If VRow Is Nothing Then

                        'Shouldn't normally happen after ForceInitialize(), but this makes
                        'the code safe if auto row generation changes.
                        VRow = New EditorRow(ColTag.ActiveColumnName)

                        VertGrid.Rows.Add(VRow)

                    End If

                    With VRow

                        .Tag = ColTag
                        .Height = IdealGridRowHeight

                        With .Properties

                            .FieldName = ColTag.ActiveColumnName
                            .Caption = " " & ColTag.ColumnHeading & " "
                            .ReadOnly = ColTag.IsReadOnly OrElse ColTag.IsCalculated

                        End With

                    End With

                    '-------------------------------------------------------------
                    ' Appearance - equivalent to Grid column configuration
                    '-------------------------------------------------------------
                    With VRow.Appearance

                        .Options.UseBackColor = True
                        .Options.UseForeColor = True
                        .Options.UseBorderColor = True

                        .BackColor = Color.Silver
                        .ForeColor = Color.Black
                        .BorderColor = Color.Silver

                    End With

                    If ColTag.IsReadOnly OrElse ColTag.IsCalculated Then

                        With VRow.Appearance

                            .BackColor = Color.WhiteSmoke
                            .ForeColor = AbovoBlue
                            .BorderColor = Color.White

                        End With

                    End If

                    '-------------------------------------------------------------
                    ' Tool tip
                    '-------------------------------------------------------------
                    If Not IsNothing(ColTag.TipText) Then

                        VRow.Properties.Caption = " " & ColTag.ColumnHeading & " "

                        'Depending upon your DevExpress version BaseRow may expose
                        'ToolTip directly. If yours does, uncomment:
                        '
                        'VRow.ToolTip = TrimSpaces(ColTag.TipText)

                    End If

                    '-------------------------------------------------------------
                    ' Repository/combo editor
                    ' Same logic as GridColumn.ColumnEdit
                    '-------------------------------------------------------------
                    Dim UseCombo As Boolean = False

                    If Not IsNothing(ColTag.RepositaryID) Then

                        ColTag.HasComboEdit = True
                        UseCombo = True

                        Dim EditControl As RepositaryItems.AbovoRespositaryItem =
            RepositaryItems.GetEditor(
                ColTag.RepositaryID,
                ModelID)

                        If EditControl.RepType = "CMB" Then

                            GridCombosCount += 1

                            EditControl.RetCombo.Name =
                "VGridCombo_" & GridCombosCount.ToString

                            VertGrid.RepositoryItems.Add(EditControl.RetCombo)

                            'VGrid equivalent of GVcolumn.ColumnEdit
                            VRow.Properties.RowEdit = EditControl.RetCombo

                        End If

                    End If

                    '-------------------------------------------------------------
                    ' Data-type formatting/editors
                    ' Mirrors the Grid Select Case
                    '-------------------------------------------------------------
                    Select Case ColTag.DataType

                        Case "S"

                            VRow.Appearance.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Near

                            VRow.Properties.Format.FormatType =
                DevExpress.Utils.FormatType.None

                            ColTag.ShowDefaultmask = -1


                        Case "M"

                            VRow.Appearance.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far

                            VRow.Properties.Format.FormatType =
                DevExpress.Utils.FormatType.Numeric

                            VRow.Properties.Format.FormatString = "c0"

                            Dim edit As New RepositoryItemTextEdit()

                            edit.Mask.MaskType =
                DevExpress.XtraEditors.Mask.MaskType.Numeric

                            edit.Mask.EditMask = "c5"

                            VertGrid.RepositoryItems.Add(edit)
                            VRow.Properties.RowEdit = edit

                            ColTag.DefaultTextEditor = edit
                            ColTag.ShowDefaultmask = 0


                        Case "SM"

                            VRow.Appearance.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far

                            VRow.Properties.Format.FormatType =
                DevExpress.Utils.FormatType.Numeric

                            VRow.Properties.Format.FormatString = "c2"

                            Dim edit As New RepositoryItemTextEdit()

                            edit.UseMaskAsDisplayFormat = True

                            edit.Mask.MaskType =
                DevExpress.XtraEditors.Mask.MaskType.Numeric

                            edit.Mask.EditMask = "c2"

                            VertGrid.RepositoryItems.Add(edit)
                            VRow.Properties.RowEdit = edit

                            ColTag.DefaultTextEditor = edit
                            ColTag.ShowDefaultmask = 1


                        Case "I"

                            VRow.Appearance.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far

                            VRow.Properties.Format.FormatType =
                DevExpress.Utils.FormatType.Numeric

                            VRow.Properties.Format.FormatString = "#,##0"

                            ColTag.ShowDefaultmask = -1

                            If Not UseCombo Then

                                Dim edit As New RepositoryItemTextEdit()

                                edit.Mask.MaskType =
                    DevExpress.XtraEditors.Mask.MaskType.Numeric

                                edit.Mask.EditMask = "n0"

                                VertGrid.RepositoryItems.Add(edit)
                                VRow.Properties.RowEdit = edit

                                ColTag.DefaultTextEditor = edit

                            End If


                        Case "C"

                            VRow.Appearance.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far

                            VRow.Properties.Format.FormatType =
                DevExpress.Utils.FormatType.Numeric

                            VRow.Properties.Format.FormatString = "#,###,##0"


                        Case "R"

                            VRow.Appearance.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far

                            VRow.Properties.Format.FormatType =
                DevExpress.Utils.FormatType.Numeric

                            VRow.Properties.Format.FormatString = "#,##0.00"

                            Dim edit As New RepositoryItemTextEdit()

                            edit.Mask.MaskType =
                DevExpress.XtraEditors.Mask.MaskType.Numeric

                            edit.Mask.EditMask = "n5"

                            VertGrid.RepositoryItems.Add(edit)
                            VRow.Properties.RowEdit = edit

                            ColTag.DefaultTextEditor = edit


                        Case "P"

                            VRow.Appearance.TextOptions.HAlignment =
                DevExpress.Utils.HorzAlignment.Far

                            VRow.Properties.Format.FormatType =
                DevExpress.Utils.FormatType.Numeric

                            VRow.Properties.Format.FormatString = "p2"

                            ColTag.ShowDefaultmask = 3

                            Dim SetMinVal As Object =
                If(ColTag.MinVal Is Nothing,
                   0,
                   If(ColTag.MinVal = "NOMIN",
                      Nothing,
                      CDec(ColTag.MinVal)))

                            Dim SetMaxVal As Object =
                If(ColTag.MaxVal Is Nothing,
                   1,
                   If(ColTag.MaxVal = "NOMAX",
                      Nothing,
                      CDec(ColTag.MaxVal)))

                            If Not ColTag.IsReadOnly Then

                                Dim edit As New RepositoryItemSpinEdit()

                                If SetMinVal IsNot Nothing Then
                                    edit.MinValue = CDec(SetMinVal)
                                End If

                                edit.Increment =
                    If(ColTag.DefIncrement Is Nothing,
                       CDec(0.0025),
                       CDec(ColTag.DefIncrement))

                                If SetMaxVal IsNot Nothing Then
                                    edit.MaxValue = CDec(SetMaxVal)
                                End If

                                edit.Mask.MaskType =
                    DevExpress.XtraEditors.Mask.MaskType.Numeric

                                edit.Mask.EditMask = "p2"

                                edit.DisplayFormat.FormatType =
                    DevExpress.Utils.FormatType.Numeric

                                edit.DisplayFormat.FormatString = "p2"

                                VertGrid.RepositoryItems.Add(edit)
                                VRow.Properties.RowEdit = edit

                                ColTag.DefaultTextEditor = edit

                            End If


                        Case "B"

                            'At present your XtraGrid block doesn't explicitly
                            'assign a Boolean editor either, so leave the VGrid
                            'using the DevExpress default.


                        Case Else

                            'Default editor generated from the underlying type.

                    End Select

                    '-------------------------------------------------------------
                    ' Move the configured editor row beneath its category
                    '-------------------------------------------------------------
                    If CatRow IsNot Nothing Then

                        VertGrid.MoveRow(VRow, CatRow, False)

                    End If

                    IndexPos += 1

                Next

                If ActiveDataSet.HasValidations Then

                    SystemLog("Adding validations")

                    For Each ValList In ActiveDataSet.ValidationLists

                        VertGrid.RepositoryItems.Add(
            RepositaryItems.GetEditorFromList(ValList).RetCombo)

                    Next

                End If

                If ActiveDataSet.RO Then VertGrid.OptionsBehavior.Editable = False

                VertGrid.RepositoryItems.Add(CustCalcEdit)

                Dim RowIndex As Integer = 0

                'For Each ColCheck As UnboundSourceProperty In PropertyArray

                '    'With VertGrid.Rows(RowIndex)

                '    '    .Tag = ColCheck
                '    '    .AppearanceCell.Options.UseBackColor = True
                '    '    .AppearanceCell.Options.UseBorderColor = True
                '    '    .AppearanceCell.BorderColor = System.Drawing.Color.Silver
                '    '    .AppearanceCell.BackColor = System.Drawing.Color.Silver
                '    '    .AppearanceCell.ForeColor = System.Drawing.Color.Black

                '    'End With



                '    'If ColCheck.UserTag.IsReadOnly OrElse ColCheck.UserTag.IsCalculated Then

                '    '    With VertGrid.Rows(RowIndex)

                '    '        '.OptionsRow.TabStop = False
                '    '        '.OptionsColumn.ReadOnly = True
                '    '        .AppearanceCell.BorderColor = System.Drawing.Color.White
                '    '        .AppearanceCell.BackColor = System.Drawing.Color.White
                '    '        .AppearanceCell.ForeColor = System.Drawing.Color.DarkGray

                '    '    End With

                '    'End If

                '    'If ColCheck.UserTag.IsFixed Then

                '    '    UsedGridVIEWS(GridViewCount).OptionsView.ColumnAutoWidth = False
                '    '    UsedGridVIEWS(GridViewCount).Columns(ColCheck.UserTag.ActiveColumnName).Fixed = FixedStyle.Left

                '    'End If
                '    RowIndex += 1

                'Next

                Formatter.FormatVertGrid(VertGrid)

                Dim IdealGridHeight As Integer = (IdealGridRowHeight + (2 * DefaultGridCellPadding)) * (FillSize) + 50
                VertGrid.Height = IdealGridHeight
                TP.Controls.Add(VertGrid)
                TP.SetCell(VertGrid, TPRowCount, 0)
                TP.SetColumnSpan(VertGrid, 4)

                SectionControlsCumlHeight += IdealGridHeight + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)
                TP.AutoSize = True
                TP.AutoSizeMode = AutoSizeMode.GrowAndShrink
                VertGrid.BestFit()
                UnboundDataSources(UBSDataSourceCount).AttachedVertGrid = VertGrid

                AddHandler VertGrid.CustomDrawRowValueCell, AddressOf VGrid_CustomDrawCell
                AddHandler VertGrid.ValidatingEditor, AddressOf VGrid_ValidatingEditor
                AddHandler VertGrid.ShowingEditor, AddressOf VGrid_ShowingEditor
                AddHandler VertGrid.ShownEditor, AddressOf VGrid_ShownEditor
                AddHandler VertGrid.CustomRecordCellEdit, AddressOf VGrid_CustomCellEditor
                AddHandler VertGrid.CustomRecordCellEditForEditing, AddressOf VGrid_CellEditorForEditing


#End Region
#Region "Spreadsheet"
            ElseIf SectionElement.Type = "Spreadsheet" Then

                ExcelModels(ModelID).ShowSpreadsheet(Me.ActiveSpreadsheet, Me)
#End Region

#Region "Label"
            ElseIf SectionElement.Type = "Label" Then

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                LabelsCount += 1

                ReDim Preserve Labels(LabelsCount)

                Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl With {.Text = ActiveDataSet.DataRows(0).DataCells(0).StringValue}
                Labels(LabelsCount).Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                Labels(LabelsCount).Appearance.TextOptions.WordWrap = WordWrap.Wrap
                Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical
                'Labels(LabelsCount).ToolTipController = TTController


                TP.Controls.Add(Labels(LabelsCount))
                TP.SetCell(Labels(LabelsCount), TPRowCount, 0)
                'TP.SetColumnSpan((Labels(LabelsCount), 3)

                'TPRowCount += 1
#End Region

#Region "TextBox and Spin"


            ElseIf SectionElement.Type = "TextBox" Then

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                Dim SCDT As New SingleCellDataTag With {
                .TargetWorksheet = Me.ActiveSpreadsheet,
                .TargetCell = ActiveDataSet.DataRows(0).DataCells(0).SourceAddress,
                .DataType = ActiveDataSet.DataColumns(0).ColumnTag.DataType,
                .IsCalculated = ActiveDataSet.DataColumns(0).ColumnTag.IsCalculated,
                .Label = ActiveDataSet.DataColumns(0).ColumnTag.ColumnHeading
            }


                If ActiveDataSet.DataColumns(0).ColumnTag.MaxVal IsNot Nothing Then

                    SCDT.MaxValSet = True
                    SCDT.MaxVal = CDbl(ActiveDataSet.DataColumns(0).ColumnTag.MaxVal)

                End If

                If ActiveDataSet.DataColumns(0).ColumnTag.MinVal IsNot Nothing Then

                    SCDT.MinValSet = True
                    SCDT.MinVal = CDbl(ActiveDataSet.DataColumns(0).ColumnTag.MinVal)

                End If



                ' ReDim Preserve TextBoxes(TextEditCount)

                If SCDT.DataType = "P" And Not ActiveDataSet.RO And Not SCDT.IsCalculated Then

                    LabelsCount += 1

                    ReDim Preserve Labels(LabelsCount)

                    Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl With {
                    .Tag = SCDT,
                    .Text = SCDT.Label}
                    Labels(LabelsCount).Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                    Labels(LabelsCount).Appearance.TextOptions.WordWrap = WordWrap.Wrap
                    Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical
                    'Labels(LabelsCount).ToolTipController = TTController


                    TP.Controls.Add(Labels(LabelsCount))
                    TP.SetCell(Labels(LabelsCount), TPRowCount, 0)

                    'TPRowCount += 1
                    'TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))
                    'TP.Rows(TPRowCount).Tag = TPRowCount.ToString

                    SpinEditCount += 1
                    ReDim Preserve SpinEdits(SpinEditCount)
                    SpinEdits(SpinEditCount) = New AbovoDESpinEdit With {
                                                .Name = "SpinEdit_" & SpinEditCount.ToString,
                                                        .Tag = SCDT
                                            }

                    'SpinEdTE.Properties.Increment
                    SpinEdits(SpinEditCount).Properties.Increment = 0.0025
                    SpinEdits(SpinEditCount).EnterMoveNextControl = True
                    SpinEdits(SpinEditCount).EditValue = ActiveDataSet.DataRows(0).DataCells(0).RealValue
                    SpinEdits(SpinEditCount).Properties.MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
                    SpinEdits(SpinEditCount).Properties.MaskSettings.Set("mask", "p5")
                    SpinEdits(SpinEditCount).Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                    SpinEdits(SpinEditCount).Properties.DisplayFormat.FormatString = "p2"


                    TP.Controls.Add(SpinEdits(SpinEditCount))
                    TP.SetCell(SpinEdits(SpinEditCount), TPRowCount, 1)

                    SpinEdits(SpinEditCount).Properties.Appearance.BackColor = Me.ActiveSpreadsheet.Cells(ActiveDataSet.DataRows(0).DataCells(0).SourceAddress).Fill.BackgroundColor

                    SpinEdits(SpinEditCount).Properties.Appearance.ForeColor = Me.ActiveSpreadsheet.Cells(ActiveDataSet.DataRows(0).DataCells(0).SourceAddress).Font.Color

                    SectionControlsCumlHeight += SpinEdits(SpinEditCount).Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                    If ActiveDataSet.RO Or SCDT.IsCalculated = True Then

                        RefreshableControlsCount += 1

                        ReDim Preserve RefreshableControls(RefreshableControlsCount)
                        RefreshableControls(RefreshableControlsCount) = SpinEdits(SpinEditCount)

                    Else

                        AddHandler SpinEdits(SpinEditCount).Validating, AddressOf SingleCellControlValidatingEditor
                        AddHandler SpinEdits(SpinEditCount).EditValueChanged, AddressOf SingleCellDirtyMarker
                        AddHandler SpinEdits(SpinEditCount).Leave, AddressOf SingleCell_Value_Push

                    End If

                    LastBottom = SpinEdits(SpinEditCount).Bottom

                Else
                    LabelsCount += 1

                    ReDim Preserve Labels(LabelsCount)

                    Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl With {
                    .Tag = SCDT,
                    .Text = SCDT.Label}
                    Labels(LabelsCount).Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                    Labels(LabelsCount).Appearance.TextOptions.WordWrap = WordWrap.Wrap
                    Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical
                    'Labels(LabelsCount).ToolTipController = TTController


                    TP.Controls.Add(Labels(LabelsCount))
                    TP.SetCell(Labels(LabelsCount), TPRowCount, 0)


                    TextEditCount += 1
                    ReDim Preserve TextBoxes(TextEditCount)

                    TextBoxes(TextEditCount) = New AbovoDETextEdit With {
                                    .Top = LastBottom + 20,
                                    .Width = 400,
                                    .Name = "TextBox_" & TextEditCount.ToString,
                                    .ModelID = ModelID,
                                    .Tag = SCDT,
                                    .EnterMoveNextControl = True
                                }

                    TextBoxes(TextEditCount).Initialise()



                    If SCDT.TipText IsNot Nothing Then

                        'NewTE.ToolTipController = TTController
                        'Labels(LabelsCount).ToolTipController = TTController

                        Labels(LabelsCount).ToolTip = TrimSpaces(SCDT.TipText)
                        TextBoxes(TextEditCount).ToolTip = TrimSpaces(SCDT.TipText)

                    End If

                    TP.Controls.Add(TextBoxes(TextEditCount))
                    TP.SetCell(TextBoxes(TextEditCount), TPRowCount, 1)

                    SectionControlsCumlHeight += TextBoxes(TextEditCount).Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                    If ActiveDataSet.RO Or SCDT.IsCalculated = True Then

                        RefreshableControlsCount += 1

                        ReDim Preserve RefreshableControls(RefreshableControlsCount)
                        RefreshableControls(RefreshableControlsCount) = TextBoxes(TextEditCount)

                    Else

                        AddHandler TextBoxes(TextEditCount).Validating, AddressOf SingleCellControlValidatingEditor
                        AddHandler TextBoxes(TextEditCount).EditValueChanged, AddressOf SingleCellDirtyMarker
                        AddHandler TextBoxes(TextEditCount).Leave, AddressOf SingleCell_Value_Push

                    End If

                    LastBottom = TextBoxes(TextEditCount).Bottom

                End If


#End Region

#Region "ComboBox"


            ElseIf SectionElement.Type = "ComboBox" Then

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                Dim SCDT As New SingleCellDataTag With {
                .TargetWorksheet = Me.ActiveSpreadsheet,
                .TargetCell = ActiveDataSet.DataRows(0).DataCells(0).SourceAddress,
                .DataType = ActiveDataSet.DataColumns(0).ColumnTag.DataType,
                .IsCalculated = ActiveDataSet.DataColumns(0).ColumnTag.IsCalculated,
                .Label = ActiveDataSet.DataColumns(0).ColumnTag.ColumnHeading
            }

                LabelsCount += 1
                ReDim Preserve Labels(LabelsCount)
                Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl With {.Text = ActiveDataSet.DataColumns(0).ColumnTag.ColumnHeading}

                Labels(LabelsCount).Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                Labels(LabelsCount).Appearance.TextOptions.WordWrap = WordWrap.Wrap
                Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical
                Labels(LabelsCount).AutoSizeInLayoutControl = True
                Labels(LabelsCount).AllowHtmlString = True
                Labels(LabelsCount).Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
                Labels(LabelsCount).Appearance.Options.UseTextOptions = True

                TP.Controls.Add(Labels(LabelsCount))
                TP.SetCell(Labels(LabelsCount), TPRowCount, 0)
                SectionControlsCumlHeight += Labels(LabelsCount).Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                'TPRowCount += 1
                'TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))
                'TP.Rows(TPRowCount).Tag = TPRowCount.ToString

                If Not IsNothing(ActiveDataSet.DataColumns(0).ColumnTag.RepositaryID) Then

                    'Dim ResCollection As New RepositaryItems

                    Dim EditControl As RepositaryItems.AbovoRespositaryItem
                    EditControl = RepositaryItems.GetEditor(ActiveDataSet.DataColumns(0).ColumnTag.RepositaryID, ModelID)

                    If EditControl.RepType = "CMB" Then

                        CombosCount += 1
                        ReDim Preserve Combos(CombosCount)

                        Combos(CombosCount) = EditControl.RetCombo.CreateEditor

                        Combos(CombosCount).EditValue = ActiveDataSet.DataRows(0).DataCells(0).StringValue
                        Combos(CombosCount).Properties.Appearance.Options.UseBackColor = True
                        Combos(CombosCount).Properties.Appearance.Options.UseForeColor = True

                        Combos(CombosCount).Tag = SCDT
                        Combos(CombosCount).Properties.Items.AddRange(EditControl.ListItems)
                        Combos(CombosCount).Name = "Combo_" & CombosCount.ToString
                        Combos(CombosCount).Properties.Appearance.ForeColor = Color.White
                        Combos(CombosCount).Properties.Appearance.BackColor = Me.ActiveSpreadsheet.Cells(ActiveDataSet.DataRows(0).DataCells(0).SourceAddress).Fill.BackgroundColor

                        Combos(CombosCount).EnterMoveNextControl = True

                        SectionControlsCumlHeight += Combos(CombosCount).Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                        TP.Controls.Add(Combos(CombosCount))
                        TP.SetCell(Combos(CombosCount), TPRowCount, 1)

                        If ActiveDataSet.DataColumns(0).ColumnTag.TipText IsNot Nothing Then

                            'Combos(CombosCount).ToolTipController = TTController
                            'Labels(LabelsCount).ToolTipController = TTController

                            Labels(LabelsCount).ToolTip = TrimSpaces(SCDT.TipText)
                            Combos(CombosCount).ToolTip = TrimSpaces(SCDT.TipText)

                        End If

                        AddHandler Combos(CombosCount).Enter, AddressOf ComboOpen
                        AddHandler Combos(CombosCount).EditValueChanged, AddressOf SingleCell_Value_Push

                    End If

                End If


#End Region

#Region "DateBox"


            ElseIf SectionElement.Type = "DateBox" Then

                ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                LabelsCount += 1
                ReDim Preserve Labels(LabelsCount)

                Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl With {.Text = ActiveDataSet.DataColumns(0).ColumnTag.ColumnHeading}
                Labels(LabelsCount).Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                Labels(LabelsCount).Appearance.TextOptions.WordWrap = WordWrap.Wrap
                Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical

                TP.Controls.Add(Labels(LabelsCount))
                TP.SetCell(Labels(LabelsCount), TPRowCount, 0)

                'TPRowCount += 1
                'TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))
                'TP.Rows(TPRowCount).Tag = TPRowCount.ToString

                DateEditCount += 1
                ReDim Preserve DateBoxes(DateEditCount)

                Dim SCDT As New SingleCellDataTag With {
                .TargetWorksheet = Me.ActiveSpreadsheet,
                .TargetCell = ActiveDataSet.DataRows(0).DataCells(0).SourceAddress,
                .DataType = ActiveDataSet.DataColumns(0).ColumnTag.DataType,
                .IsCalculated = ActiveDataSet.DataColumns(0).ColumnTag.IsCalculated,
                .Label = ActiveDataSet.DataColumns(0).ColumnTag.ColumnHeading
            }

                DateBoxes(DateEditCount) = New AbovoDEDateEdit With {
                    .EnterMoveNextControl = True,
                    .TargetCell = SCDT.TargetCell,
                    .ModelID = ModelID,
                    .TargetWorksheet = Me.ActiveSpreadsheet,
                    .Tag = SCDT
                }

                DateBoxes(DateEditCount).Initialise()

                'DateBoxes(DateEditCount).DateTime = DateTime.FromOADate(ActiveDataSet.DataRows(0).DataCells(0).IntValue)

                DateBoxes(DateEditCount).BackColor = AbovoBlue
                DateBoxes(DateEditCount).ForeColor = Color.White

                If SCDT.TipText IsNot Nothing Then

                    'DateBoxes(DateEditCount).ToolTipController = TTController
                    'Labels(LabelsCount).ToolTipController = TTController

                    Labels(LabelsCount).ToolTip = TrimSpaces(SCDT.TipText)
                    DateBoxes(DateEditCount).ToolTip = TrimSpaces(SCDT.TipText)

                End If

                If SCDT.DataType = "DM" Then

                    DateBoxes(DateEditCount).Properties.DisplayFormat.FormatString = "dd-MMM"

                End If

                AddHandler DateBoxes(DateEditCount).EditValueChanged, AddressOf SingleCell_Value_Push

                SectionControlsCumlHeight += DateBoxes(DateEditCount).Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                TP.Controls.Add(DateBoxes(DateEditCount))
                TP.SetCell(DateBoxes(DateEditCount), TPRowCount, 1)
#End Region

#Region "Spacer"

            ElseIf SectionElement.Type = "Spacer" Then

                'LabelsCount += 1
                'ReDim Preserve Labels(LabelsCount)
                'Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl
                'Labels(LabelsCount).Visible = False
                'Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical
                'TP.Controls.Add(Labels(LabelsCount))
                'TP.SetCell(Labels(LabelsCount), TPRowCount, 0)
                'SectionControlsCumlHeight += Labels(LabelsCount).Height + (2 * DefaultTablePanelPadding)
#End Region

#Region "Link"

            ElseIf SectionElement.Type = "Link" Then

                For Each LControl In SectionElement.LinkControls

                    If LControl.LinkType = "GOTO" Then

                        Dim ButtText As String = LControl.LinkText

                        ButtText = ButtText.Replace("vblf", "<br>")
                        Dim NewButton As New DevExpress.XtraEditors.SimpleButton

                        NewButton.AllowHtmlDraw = DefaultBoolean.True
                        NewButton.Appearance.TextOptions.WordWrap = WordWrap.Wrap
                        NewButton.Text = ButtText
                        NewButton.ToolTip = LControl.LinkTip

                        LControl.LinkReturnID = GSID

                        NewButton.Tag = LControl

                        NewButton.Padding = New System.Windows.Forms.Padding(10)
                        Dim bestSize As Size = NewButton.CalcBestSize()
                        NewButton.Width = bestSize.Width
                        NewButton.Height = bestSize.Height

                        AddHandler NewButton.Click, AddressOf Link_ButtonClick
                        TP.Controls.Add(NewButton)
                        TP.SetCell(NewButton, TPRowCount, 0)
                        SectionControlsCumlHeight += NewButton.Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                        TPRowCount += 1
                        TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))
                        TP.Rows(TPRowCount).Tag = TPRowCount.ToString

                    End If

                Next



#End Region

#Region "TextArea"

            ElseIf SectionElement.Type = "TextArea" Then

                Dim NewRTF As New DevExpress.XtraEditors.LabelControl

                NewRTF.AllowHtmlString = True
                NewRTF.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
                NewRTF.Appearance.Options.UseTextOptions = True
                NewRTF.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Vertical
                'NewRTF.HyperlinkClick += LabelControl1_HyperlinkClick

                NewRTF.Text = SectionElement.Description
                TP.Controls.Add(NewRTF)
                TP.SetCell(NewRTF, TPRowCount, 0)
                SectionControlsCumlHeight += NewRTF.Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

#End Region

#Region "Browser"

            ElseIf SectionElement.Type = "Browser" Then

                Dim NewWB As New System.Windows.Forms.WebBrowser

                NewWB.DocumentText = SectionElement.Description

                TP.Controls.Add(NewWB)
                TP.SetCell(NewWB, TPRowCount, 0)
                TP.SetColumnSpan(NewWB, 3)
                SectionControlsCumlHeight += NewWB.Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

#End Region

#Region "CompundMappedTable"

            ElseIf SectionElement.Type = "CompoundMappedTable" Then


                Dim DockPosition As DockStyle = DockStyle.Right
                Dim x As Integer = 0
                Dim RowCount As Integer = 0
                Dim ColCount As Integer = 0
                Dim LastRowIndex As Integer = 0
                Dim ConPadding As Padding
                ConPadding.All = 10
                ConPadding.Top = 5
                ConPadding.Bottom = 5
                TP.Rows.Clear()
                TP.Columns.Clear()

                TP.Width = Me.Width

                Dim Newpadding As New Padding()
                Newpadding.Right += 20
                TP.Padding = DefaultTablePanelPadding

                Dim CMTab As CompoundMappedTable = SectionElement.CompoundTable

                'Define the CMT starting position and size

                Dim StartIndexCol As Integer = 0
                Dim StartIndexRow As Integer = 0
                Dim TotalColumnCount As Integer = 0

                Dim TestRange As DevExpress.Spreadsheet.CellRange = Nothing

                If Microsoft.VisualBasic.Left(CMTab.StartingCol, 4) = "Col-" Then

                    StartIndexCol = CInt(Mid(CMTab.StartingCol, 5))

                Else

                    Dim StartRnge As DevExpress.Spreadsheet.CellRange = Me.ActiveWorkbook.Range(Mid(CMTab.StartingCol, 5))
                    StartIndexCol = StartRnge.LeftColumnIndex

                End If

                If Microsoft.VisualBasic.Left(CMTab.StartingRow, 4) = "Row-" Then

                    StartIndexRow = CInt(Mid(CMTab.StartingRow, 5))

                Else

                    TestRange = Me.ActiveWorkbook.Range(Mid(CMTab.StartingRow, 4))
                    StartIndexRow = TestRange.TopRowIndex

                End If

                For Each CMTColDef As CompoundMappedTableColumnDefinition In CMTab.CMTColDefs

                    If Microsoft.VisualBasic.Left(CMTColDef.ColSetType, 6) = "Fixed-" Then

                        CMTColDef.ColCount = CInt(Mid(CMTColDef.ColSetType, 7))
                        TotalColumnCount += CInt(Mid(CMTColDef.ColSetType, 7))

                    Else

                        TestRange = Me.ActiveWorkbook.Range(Mid(CMTColDef.RepeatsBy, 6))
                        CMTColDef.ColCount = TestRange.ColumnCount
                        CMTColDef.StartColIndex = TestRange.LeftColumnIndex

                        For x = LastRowIndex To (LastRowIndex + CMTColDef.ColCount - 1)

                            TP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.AutoSize, 50))

                            TP.Columns(x).Tag = New MappedTableColumnTag With {
                            .ColDef = CMTColDef.ColSetType,
                            .ColSetID = CMTColDef.ColSetID,
                            .Index = x
                        }

                        Next

                        TotalColumnCount += CMTColDef.ColCount

                    End If

                    If Not IsNothing(CMTColDef.RepeatsBy) Then CMTColDef.ElementRepeats = True


                    Debug.Print(TotalColumnCount.ToString)

                Next

                Dim RowRepeats As Integer = 0
                Dim ElementRepeats As Integer = 0
                Dim OverallRowIndex As Integer = 0

                For Each MTRow As MappedTableRow In CMTab.MappedTableRows

                    RowRepeats = 0

                    If MTRow.RepeatsBy IsNot Nothing Then

                        If Microsoft.VisualBasic.Left(MTRow.RepeatsBy, 4) = "Row-" Then

                            RowRepeats = CInt(Mid(MTRow.RepeatsBy, 5))

                        Else

                            TestRange = Me.ActiveWorkbook.Range(Mid(MTRow.RepeatsBy, 5))
                            RowRepeats = CInt(MTRow.RepeatsBy)

                        End If

                    End If

                    Dim CurrentRowIndex As Integer = 0
                    Dim CurrColDef As Integer = 0

                    For x = 0 To RowRepeats

                        If CMTab.CMTColDefs(CurrColDef).ElementRepeats = True Then



                        End If

                        CurrColDef += 1

                        MTRow.RowIndex = OverallRowIndex

                        OverallRowIndex += 1

                    Next


                Next


#End Region

#Region "MappedTable"

            ElseIf SectionElement.Type = "MappedTable" Then

                Dim MTab As MappedTable = SectionElement.MappedTableSection
                Dim DockPosition As DockStyle = DockStyle.Right
                Dim x As Integer = 0
                Dim RowCount As Integer = CInt(MTab.NumRows)
                Dim ColCount As Integer = CInt(MTab.NumCols)
                Dim ConPadding = DefaultTablePanelPadding
                ConPadding.All = 10
                ConPadding.Top = 5
                ConPadding.Bottom = 5
                TP.Rows.Clear()
                TP.Columns.Clear()

                TP.Width = Me.Width

                'Dim Newpadding = DefaultTablePanelPadding
                'Newpadding.Right += 20
                TP.Padding = DefaultTablePanelPadding

                Dim MaxHeight As Integer = 0

                'Dim lastItem As LayoutControlItem = Nothing
                Dim DoneFirst As Boolean = False
                Dim CurrentCol As Integer = 0
                Dim CurrentRow As Integer = -1

                For x = 1 To ColCount + 1

                    TP.Columns.Add(New TablePanelColumn(TablePanelEntityStyle.AutoSize, TP.Width / CInt(ColCount)))
                    TP.Columns(CurrentCol).Tag = x.ToString

                    CurrentCol += 1

                Next



                For Each MappedTableRow In MTab.MappedTableRows

                    TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))
                    CurrentRow += 1
                    CurrentCol = 0

                    Dim LastRight As Integer = 0
                    Dim ColSpan As Integer = 1
                    Dim RowSpan As Integer = 1

                    Dim CurrControl As Control = Nothing

                    For Each MTE As MappedTableElement In MappedTableRow.MappedTableElements

                        ColSpan = 1
                        RowSpan = 1

                        DockPosition = DockStyle.Right

                        If MTE.Type = "Skip" Then

                            GoTo NextCell

                        ElseIf MTE.Type = "Link" Then

                            ADEHyperlinksCount += 1
                            ReDim Preserve ADEHyperlinks(ADEHyperlinksCount)

                            Dim TargetInterface As String = Me.ActiveSpreadsheet.Cells(MTE.TELinks(0).CSSource).DisplayText
                            Dim LinkTag As New ElementInterfaceLinkTag With {
                            .LinkReturnID = CSID,
                            .LinkReturnName = DataPres.PresName,
                            .LinkData = TargetInterface,
                            .LinkReturnGroup = GSID,
                            .LinkTargetSection = MTE.TELinks(0).LinkTargetSection
                        }

                            ADEHyperlinks(ADEHyperlinksCount) = New AbovoDEHyperlinkLabel With {
                            .TargetWorksheet = Me.ActiveSpreadsheet,
                            .AutoSizeHorizontal = True,
                            .TargetCell = MTE.TELinks(0).CSSource,
                            .Dock = DockStyle.Fill,
                            .Tag = LinkTag
                        }

                            DockPosition = DockStyle.Left

                            ADEHyperlinks(ADEHyperlinksCount).Initialise()
                            CurrControl = ADEHyperlinks(ADEHyperlinksCount)
                            AddHandler ADEHyperlinks(ADEHyperlinksCount).HyperlinkClick, AddressOf DEHyperLinkLabel_linkClick

                        ElseIf MTE.Type = "ComboBox" Then


                            Dim TECB As TEComboBox = MTE.TEComboBoxes(0)
                            Dim SCDT As New SingleCellDataTag With {
                            .TargetWorksheet = Me.ActiveSpreadsheet,
                            .DataType = TECB.DataType,
                            .TargetCell = TECB.CSSource
                        }

                            ADEComboBoxesCount += 1
                            ReDim Preserve ADEComboBoxes(ADEComboBoxesCount)
                            ADEComboBoxes(ADEComboBoxesCount) = New AbovoDEComboBox With {
                            .Tag = SCDT,
                            .RepID = TECB.RepID,
                            .Dock = DockStyle.Fill,
                            .AutoSizeHorizontal = True
                        }

                            RefreshableControlsCount += 1
                            ReDim Preserve RefreshableControls(RefreshableControlsCount)
                            RefreshableControls(RefreshableControlsCount) = ADEComboBoxes(ADEComboBoxesCount)

                            ADEComboBoxes(ADEComboBoxesCount).InitialiseStandard(TECB.RepID)
                            AddHandler ADEComboBoxes(ADEComboBoxesCount).EditValueChanged, AddressOf SingleCell_Value_Push
                            CurrControl = ADEComboBoxes(ADEComboBoxesCount)

                        ElseIf MTE.Type = "Label" Then

                            Dim TEL As TELabel = MTE.TELabels(0)

                            Dim SCDT As New SingleCellDataTag With {
                            .TargetWorksheet = Me.ActiveSpreadsheet,
                            .TargetCell = TEL.CSSource
                        }

                            Dim DoTitle As Integer = 0

                            If TEL.IsTitle Is Nothing OrElse TEL.IsTitle = "False" Then

                                DoTitle = 0

                            ElseIf TEL.IsTitle = "True" Then

                                DoTitle = 1

                            Else

                                DoTitle = CInt(TEL.IsTitle)

                            End If

                            ADELabelsCount += 1
                            ReDim Preserve ADELabels(ADELabelsCount)

                            If Not IsNothing(MTE.ColSpan) Then

                                ColSpan = CInt(MTE.ColSpan)

                            End If

                            If Not IsNothing(MTE.RowSpan) Then

                                RowSpan = CInt(MTE.RowSpan)

                            End If

                            ADELabels(ADELabelsCount) = New AbovoDELabel With {
                            .Tag = SCDT,
                            .AutoSizeHorizontal = True,
                            .Dock = DockStyle.Fill,
                            .IsBold = IIf(TEL.IsBold = "True", True, False),
                            .IsUnderline = IIf(TEL.IsUnderline = "True", True, False),
                            .IsTitle = DoTitle
                        }
                            ADELabels(ADELabelsCount).AutoSizeInLayoutControl = True
                            'ADELabels(ADELabelsCount).RealAutoSizeMode = AutoSizeMode.Vertical
                            If TEL.Alignment IsNot Nothing Then

                                Select Case TEL.Alignment.ToUpper
                                    Case "NEAR", "LEFT"
                                        ADELabels(ADELabelsCount).TxtHlignment = DevExpress.Utils.HorzAlignment.Near
                                    Case "FAR", "RIGHT"
                                        ADELabels(ADELabelsCount).TxtHlignment = DevExpress.Utils.HorzAlignment.Far
                                    Case "CENTER", "CENTRE", "MIDDLE"
                                        ADELabels(ADELabelsCount).TxtHlignment = DevExpress.Utils.HorzAlignment.Center
                                End Select

                            Else

                                ADELabels(ADELabelsCount).TxtHlignment = DevExpress.Utils.HorzAlignment.Far

                            End If

                            If TEL.IsBordered = "True" Then

                                ADELabels(ADELabelsCount).DoBorder = True

                            End If

                            If TEL.IsStatic IsNot "True" Then

                                RefreshableControlsCount += 1
                                ReDim Preserve RefreshableControls(RefreshableControlsCount)
                                RefreshableControls(RefreshableControlsCount) = ADELabels(ADELabelsCount)

                            Else

                                ADELabels(ADELabelsCount).IsStatic = True

                            End If

                            ADELabels(ADELabelsCount).Initialise()

                            If ADELabels(ADELabelsCount).TxtHlignment = HorzAlignment.Far Then
                                DockPosition = DockStyle.Right
                            Else
                                DockPosition = DockStyle.Left
                            End If

                            CurrControl = ADELabels(ADELabelsCount)

                        ElseIf MTE.Type = "DateEdit" Then

                            ADEDateEditsCount += 1
                            ReDim Preserve ADEDateEdits(ADEDateEditsCount)

                            If Not IsNothing(MTE.ColSpan) Then

                                ColSpan = CInt(MTE.ColSpan)

                            End If

                            If Not IsNothing(MTE.RowSpan) Then

                                RowSpan = CInt(MTE.RowSpan)

                            End If

                            ADEDateEdits(ADEDateEditsCount) = New AbovoDEDateEdit With {
                            .TargetWorksheet = Me.ActiveSpreadsheet,
                            .Dock = DockStyle.Fill,
                            .TargetCell = MTE.TELabels(0).CSSource
                        }

                            If MTE.TEDateEdits(0).IsStatic = "True" Then

                                ADEDateEdits(ADEDateEditsCount).IsStatic = True

                            Else

                                ADEDateEdits(ADEDateEditsCount).IsStatic = False
                                RefreshableControlsCount += 1
                                ReDim Preserve RefreshableControls(RefreshableControlsCount)
                                RefreshableControls(RefreshableControlsCount) = ADEDateEdits(ADEDateEditsCount)

                            End If

                        ElseIf MTE.Type = "TextEdit" Then

                            ADETextEditsCount += 1
                            ReDim Preserve ADETextEdits(ADETextEditsCount)

                            If Not IsNothing(MTE.ColSpan) Then

                                ColSpan = CInt(MTE.ColSpan)

                            End If

                            If Not IsNothing(MTE.RowSpan) Then

                                RowSpan = CInt(MTE.RowSpan)

                            End If

                            Dim TETE As TETextBox = MTE.TETextBoxes(0)

                            Dim SCDataDag As New SingleCellDataTag With {
                            .DataType = TETE.DataType,
                            .TargetWorksheet = Me.ActiveSpreadsheet,
                            .TargetCell = TETE.CSSource
                        }

                            If TETE.MaxVal IsNot Nothing Then

                                SCDataDag.MaxValSet = True
                                SCDataDag.MaxVal = CDbl(TETE.MaxVal)

                            End If

                            If TETE.MinVal IsNot Nothing Then

                                SCDataDag.MinValSet = True
                                SCDataDag.MinVal = CDbl(TETE.MinVal)

                            End If

                            If MTE.ToolTip IsNot Nothing Then

                                ADETextEdits(ADETextEditsCount).ToolTip = TrimSpaces(MTE.ToolTip)

                            End If

                            ADETextEdits(ADETextEditsCount) = New AbovoDETextEdit With {
                                .ModelID = ModelID,
                                .EnterMoveNextControl = True,
                                .Dock = DockStyle.Fill,
                            .Tag = SCDataDag
                        }

                            If MTE.TETextBoxes(0).IsReadOnly = "True" Then

                                ADETextEdits(ADETextEditsCount).IsReadOnly = True

                            Else

                                ADETextEdits(ADETextEditsCount).IsReadOnly = False
                                RefreshableControlsCount += 1
                                ReDim Preserve RefreshableControls(RefreshableControlsCount)
                                RefreshableControls(RefreshableControlsCount) = ADETextEdits(ADETextEditsCount)

                            End If

                            ADETextEdits(ADETextEditsCount).Initialise()

                            If ADETextEdits(ADETextEditsCount).Properties.Appearance.TextOptions.HAlignment = HorzAlignment.Far Then
                                DockPosition = DockStyle.Right
                            Else
                                DockPosition = DockStyle.Left
                            End If


                            CurrControl = ADETextEdits(ADETextEditsCount)
                            ADETextEdits(ADETextEditsCount).Properties.EditValueChangedFiringMode = EditValueChangedFiringMode.Buffered
                            ADETextEdits(ADETextEditsCount).Properties.EditValueChangedDelay = 400

                            AddHandler ADETextEdits(ADETextEditsCount).Validating, AddressOf SingleCellControlValidatingEditor
                            AddHandler ADETextEdits(ADETextEditsCount).EditValueChanged, AddressOf SingleCellDirtyMarker
                            AddHandler ADETextEdits(ADETextEditsCount).Leave, AddressOf SingleCell_Value_Push

                        ElseIf MTE.Type = "Hyperlink" Then

                            ADETextEditsCount += 1
                            ReDim Preserve ADETextEdits(ADETextEditsCount)

                            If Not IsNothing(MTE.ColSpan) Then

                                ColSpan = CInt(MTE.ColSpan)

                            End If

                            If Not IsNothing(MTE.RowSpan) Then

                                RowSpan = CInt(MTE.RowSpan)

                            End If

                            Dim TETE As TETextBox = MTE.TETextBoxes(0)

                            Dim SCDataDag As New SingleCellDataTag With {
                            .DataType = TETE.DataType,
                            .TargetWorksheet = Me.ActiveSpreadsheet,
                            .TargetCell = TETE.CSSource
                        }

                            If TETE.MaxVal IsNot Nothing Then

                                SCDataDag.MaxValSet = True
                                SCDataDag.MaxVal = CDbl(TETE.MaxVal)

                            End If

                            If TETE.MinVal IsNot Nothing Then

                                SCDataDag.MinValSet = True
                                SCDataDag.MinVal = CDbl(TETE.MinVal)

                            End If

                            If MTE.ToolTip IsNot Nothing Then

                                ADETextEdits(ADETextEditsCount).ToolTip = TrimSpaces(MTE.ToolTip)

                            End If

                            ADETextEdits(ADETextEditsCount) = New AbovoDETextEdit With {
                                .ModelID = ModelID,
                                .EnterMoveNextControl = True,
                                .Dock = DockStyle.Fill,
                            .Tag = SCDataDag
                        }

                            If MTE.TETextBoxes(0).IsReadOnly = "True" Then

                                ADETextEdits(ADETextEditsCount).IsReadOnly = True

                            Else

                                ADETextEdits(ADETextEditsCount).IsReadOnly = False
                                RefreshableControlsCount += 1
                                ReDim Preserve RefreshableControls(RefreshableControlsCount)
                                RefreshableControls(RefreshableControlsCount) = ADETextEdits(ADETextEditsCount)

                            End If

                            ADETextEdits(ADETextEditsCount).Initialise()

                            If ADETextEdits(ADETextEditsCount).Properties.Appearance.TextOptions.HAlignment = HorzAlignment.Far Then
                                DockPosition = DockStyle.Right
                            Else
                                DockPosition = DockStyle.Left
                            End If


                            CurrControl = ADETextEdits(ADETextEditsCount)
                            ADETextEdits(ADETextEditsCount).Properties.EditValueChangedFiringMode = EditValueChangedFiringMode.Buffered
                            ADETextEdits(ADETextEditsCount).Properties.EditValueChangedDelay = 400

                            AddHandler ADETextEdits(ADETextEditsCount).Validating, AddressOf SingleCellControlValidatingEditor
                            AddHandler ADETextEdits(ADETextEditsCount).EditValueChanged, AddressOf SingleCellDirtyMarker
                            AddHandler ADETextEdits(ADETextEditsCount).Leave, AddressOf SingleCell_Value_Push

                        End If


                        If ColSpan > 1 Then

                            TP.SetColumnSpan(CurrControl, ColSpan)

                        End If

                        If RowSpan = 3 Then

                            TP.SetRowSpan(CurrControl, RowSpan)

                        End If

                        CurrControl.Margin = ConPadding

                        If MaxHeight < CurrControl.Height Then MaxHeight = CurrControl.Height

                        CurrControl.Dock = DockPosition

                        TP.Controls.Add(CurrControl)
                        TP.SetCell(CurrControl, CurrentRow, CurrentCol)
                        'TP.SetColumnSpan(CurrControlGridControls(GridCount), 3)
NextCell:


                        CurrentCol += 1
                        CurrentCol += ColSpan - 1

                        CurrControl = Nothing

                    Next

                    SectionControlsCumlHeight += MaxHeight + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                    MaxHeight = 0

                Next

#End Region

#Region "Control Group"

            ElseIf SectionElement.Type = "ControlGroup" Then

                Dim LastRight As Integer = 0

                'For Each GridButtControl In SectionElement.GridControls

                '    ButtText = GridButtControl.CommandText.Replace("vblf", "<br>")

                '    NewButton.AllowHtmlDraw = DefaultBoolean.True
                '    NewButton.Appearance.TextOptions.WordWrap = WordWrap.Wrap
                '    NewButton.Text = ButtText
                '    NewButton.ToolTip = GridButtControl.CommandTip

                '    GridButtControl.AttachedGrid = GridControls(GridCount)

                '    NewButton.Tag = GridButtControl



                '    'NewButton.Padding = New Padding(10)
                '    Dim bestSize As Size = NewButton.CalcBestSize()
                '    NewButton.Width = bestSize.Width
                '    NewButton.Height = bestSize.Height

                '    AddHandler NewButton.Click, AddressOf Grid_ButtonClick
                '    'TP.Controls.Add(NewButton)
                '    'TP.SetCell(NewButton, TPRowCount, 0)
                '    'NewButton.Left = LastRight + 10
                '    'LastRight = NewButton.Right
                '    itemButton.Control = NewButton


                'Next

                Dim NewButton As DevExpress.XtraEditors.SimpleButton = Nothing
                Dim ButtText As String = ""

                Dim Layout As LayoutControl = New LayoutControl()
                Layout.BackColor = Color.White

                Layout.OptionsView.EnableTransparentBackColor = False
                Layout.Padding = DefaultTablePanelPadding

                Layout.Dock = DockStyle.Fill

                Layout.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
                Layout.LookAndFeel.UseDefaultLookAndFeel = False
                Layout.OptionsView.EnableTransparentBackColor = True
                Layout.Root.AppearanceGroup.BackColor = System.Drawing.Color.Transparent
                Layout.Root.AppearanceGroup.Options.UseBackColor = True

                TP.Controls.Add(Layout)
                TP.SetCell(Layout, TPRowCount, 0)

                Dim groupButtons As LayoutControlGroup = Layout.Root.AddGroup()
                groupButtons.Name = "GroupButtons"
                groupButtons.GroupBordersVisible = False

                Dim lastItem As LayoutControlItem = Nothing
                Dim DoneFirst As Boolean = False

                For Each IControl In SectionElement.InterfaceControls

                    NewButton = New DevExpress.XtraEditors.SimpleButton()


                    Dim itemButton As LayoutControlItem = Nothing

                    If Not DoneFirst Then

                        itemButton = Layout.AddItem(groupButtons, InsertType.Left)

                        DoneFirst = True
                        lastItem = itemButton

                    Else

                        itemButton = Layout.AddItem(lastItem, InsertType.Right)
                        DoneFirst = True
                        lastItem = itemButton

                    End If

                    ButtText = IControl.ItemTxt

                    ButtText = ButtText.Replace("vblf", "<br>")

                    NewButton.AllowHtmlDraw = DefaultBoolean.True
                    NewButton.Appearance.TextOptions.WordWrap = WordWrap.Wrap
                    NewButton.Text = ButtText
                    NewButton.ToolTip = IControl.ItemTip
                    NewButton.Tag = IControl.ItemData

                    NewButton.Left = LastRight + 10

                    NewButton.Padding = New System.Windows.Forms.Padding(10)
                    Dim bestSize As Size = NewButton.CalcBestSize()
                    NewButton.Width = bestSize.Width
                    NewButton.Height = bestSize.Height
                    itemButton.Control = NewButton
                    AddHandler NewButton.Click, AddressOf Interface_ButtonClick

                Next

                Dim emptySpace As New EmptySpaceItem

                emptySpace.Parent = groupButtons


                SectionControlsCumlHeight += NewButton.Height + (DefaultTablePanelPadding.Top + DefaultTablePanelPadding.Bottom)

                TPRowCount += 1

                TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.AutoSize, 50))

                TP.Rows(TPRowCount).Tag = TPRowCount.ToString

#End Region

            End If

            SystemLog("CumulativeControlHeight: " & SectionControlsCumlHeight)
            'AcContainers(AcContainersCount).Height = SectionControlsCumlHeight
            'AcContainers(AcContainersCount).Width = Me.Width
            'AcContainers(AcContainersCount).Appearance.BackColor = AbovoBlue

            'Add a small spacer
            TPRowCount += 1
            TP.Rows.Add(New TablePanelRow(TablePanelEntityStyle.Absolute, 50))
            TP.Rows(TPRowCount).Tag = TPRowCount.ToString
            'LabelsCount += 1
            'ReDim Preserve Labels(LabelsCount)
            'Labels(LabelsCount) = New DevExpress.XtraEditors.LabelControl
            'Labels(LabelsCount).Visible = False
            'Labels(LabelsCount).AutoSizeMode = LabelAutoSizeMode.Vertical
            'TP.Controls.Add(Labels(LabelsCount))
            'TP.SetCell(Labels(LabelsCount), TPRowCount, 0)
            'SectionControlsCumlHeight += Labels(LabelsCount).Height + (2 * DefaultTablePanelPadding)

        Next

        TP.Height = SectionControlsCumlHeight 'account for padding and the final spacer
        TP.EndInit()

        Formatter.FormatTablePanel(TP)
        Debug.Print("Final TP Height: " & TP.Height.ToString & " SCUMH: " & SectionControlsCumlHeight.ToString)

    End Sub

#Region "Public Methods"
    Sub RebuildSection(SectionIndex As Integer)

        DataPres.RedefineInterfaceSection(SectionIndex)

        Dim XTP As XtraTabPage = XtraTabControlNewGIT.TabPages(SectionIndex)

        Dim ThisSection As PresentationSection = DataPres.Sections(SectionIndex)

        TPs(SectionIndex).Dispose()
        TPs(SectionIndex) = Nothing
        TPs(SectionIndex) = New TablePanel
        GetSectionControlCollection(ThisSection, Me, SectionIndex, TPs(SectionIndex))
        XTP.Controls.Add(TPs(SectionIndex))

        'er?
        'AcContainers(AcContainersCount).Height = TPs(SectionIndex).Height
        'SystemLog("ReturnedTPHeight =" & TPs(SectionIndex).Height.ToString)

    End Sub

    Sub RebuildAllSections()

        ClearAllGrids()
        'ClearAllTPs()

        For Each sec In DataPres.Sections
            RebuildSection(sec.ID)
        Next


    End Sub

    Public Sub ClearAllTPs()
        On Error Resume Next
        For Each TP In TPs
            For Each C As Control In TP.Controls
                C.Dispose()
            Next
            TP.Controls.Clear()
            TP.Dispose()
            TP = Nothing
        Next

    End Sub

    Public Sub ClearAllGrids()

        On Error Resume Next


        For Each gc In GridControls

            Dim gcView As GridView = TryCast(gc.FocusedView, GridView)
            If gcView IsNot Nothing Then
                gcView.Columns.Clear()
            Else
                Dim bgvView As BandedGridView = TryCast(gc.FocusedView, BandedGridView)
                If bgvView IsNot Nothing Then
                    bgvView.Columns.Clear()
                    bgvView.Bands.Clear()
                End If
            End If

            gc.DataSource = Nothing

            gc.Dispose()
            gc = Nothing

        Next

        GridCount = -1
        GridControls = Nothing
        ReDim GridControls(-1)

        For Each UBS In UnboundDataSources

            UBS.Dispose()
            UBS = Nothing

        Next


        UnboundDataSources = Nothing
        UBSDataSourceCount = -1
        ReDim UnboundDataSources(-1)

        For Each GV In UsedGridVIEWS

            GV.Dispose()
            GV = Nothing

        Next

        UsedGridVIEWS = Nothing
        ReDim UsedGridVIEWS(-1)
        GridViewCount = -1

        For Each BGV In UsedBANDedGridVIEWS

            BGV.Dispose()
            BGV = Nothing

        Next

        UsedBANDedGridVIEWS = Nothing
        ReDim UsedBANDedGridVIEWS(-1)
        BandGridViewsCount = -1

    End Sub

    Public Sub RefreshROControls()

        For Each RefreshableObject In RefreshableControls

            Try

                RefreshableObject.RefreshData()

            Catch ex As Exception

                SystemLog("Error refreshing control: " & ex.Message)

            End Try

        Next


    End Sub
    Sub SetFooterOn(SetGridView As GridView, Optional ByVal ShowTotalsWord As Boolean = True)

        FooterOn = True
        UsedGridVIEWS(GridViewCount).OptionsView.ShowFooter = True


        'If ShowTotalsWord Then

        '    Dim itemCust As New GridGroupSummaryItem
        '    itemCust.FieldName = UsedGridVIEWS(GridViewCount).Columns(0).FieldName
        '    itemCust.SummaryType = DevExpress.Data.SummaryItemType.Custom
        '    itemCust.DisplayFormat = "Totals"
        '    itemCust.ShowInGroupColumnFooter = UsedGridVIEWS(GridViewCount).Columns(0)
        '    UsedGridVIEWS(GridViewCount).GroupSummary.Add(itemCust)

        'End If

        FooterDone = True
        SetGridView.FooterPanelHeight += (2 * DefaultGridCellPadding)

    End Sub

    Sub SetBandedFooterOn(SetGridView As GridView)


    End Sub


#End Region

#Region "Menu Button Actions"
    Sub AddMenuInterface(InterfaceMode As String)

        WindowsUIButtonPanelActions.AutoSizeInLayoutControl = False
        WindowsUIButtonPanelActions.ForeColor = AbovoBlue
        WindowsUIButtonPanelActions.AppearanceButton.Normal.ForeColor = AbovoBlue
        WindowsUIButtonPanelActions.AppearanceButton.Normal.ForeColor = AbovoBlue

    End Sub
    Private Sub ProcessLinkElement()

        For Each Button In Me.WindowsUIButtonPanelActions.Buttons

            If TryCast(Button, DevExpress.XtraBars.Docking2010.WindowsUIButton) IsNot Nothing Then

                Dim ValidButton As DevExpress.XtraBars.Docking2010.WindowsUIButton = Button

                If ValidButton.Tag = "Return" Then

                    ValidButton.ToolTip = "Return to " & ActiveLinkElement.LinkReturnName
                    ValidButton.Visible = True

                    GoTo SectionSelect

                End If

            End If

        Next

SectionSelect:

        If ActiveLinkElement.LinkTargetSection IsNot Nothing Then

            For Each TabPage In XtraTabControlNewGIT.TabPages

                If Trim(TabPage.Text) = Trim(ActiveLinkElement.LinkTargetSection) Then

                    XtraTabControlNewGIT.SelectedTabPage = TabPage
                    Exit Sub

                End If

            Next

        End If

    End Sub
    Public Sub ClearLinks()

        For Each Button In Me.WindowsUIButtonPanelActions.Buttons
            If TryCast(Button, DevExpress.XtraBars.Docking2010.WindowsUIButton) IsNot Nothing Then

                Dim ValidButton As DevExpress.XtraBars.Docking2010.WindowsUIButton = Button

                If ValidButton.Tag = "Return" Then

                    ValidButton.ToolTip = ""
                    ValidButton.Visible = False
                    Exit Sub

                End If
            End If
        Next
    End Sub
    Public Sub AddLink(LinkTag As ElementInterfaceLinkTag)

        Me.ActiveLinkElement = LinkTag
        ProcessLinkElement()

    End Sub
    Private Sub WindowsUIButtonPanelSaveClose_ButtonClick(sender As Object, e As ButtonEventArgs) Handles WindowsUIButtonPanelActions.ButtonClick

        Dim ButSender As WindowsUIButton = TryCast(e.Button, DevExpress.XtraBars.Docking2010.WindowsUIButton)

        If ButSender Is Nothing Then

            Return

        End If

        Dim tag As String = ButSender.Tag.ToString()

        Select Case tag

            Case "History"

                ' Close the model and dispose of the interface
                HistoryManager.Visible = True
                HistoryManager.Show()
                HistoryManager.BringToFront()

            Case "Spreadsheet"

                ' Close the model and dispose of the interface
                ExcelModels(ModelID).ShowSpreadsheet(Me.ActiveSpreadsheet, Me)

            Case "MainMenu"

                ' Close the model and dispose of the interface
                If FormMainScreen.WindowState = FormWindowState.Minimized Then
                    FormMainScreen.WindowState = FormWindowState.Normal
                End If

                FormMainScreen.BringToFront()

            Case "Refresh"

                ' Close the model and dispose of the interface
                RefreshData()

            Case "Return"

                ' Open the parent interface, passing the return ID to allow it to navigate to the correct place
                Dim LinkTag As New ElementInterfaceLinkTag With {
                    .LinkType = "Link",
                    .LinkData = ActiveLinkElement.LinkReturnName,
                    .LinkGroupID = ActiveLinkElement.LinkReturnGroup,
                    .LinkReturnID = CSID,
                    .LinkReturnGroup = GSID,
                    .LinkTargetSection = Nothing,
                    .LinkReturnName = DITName
                }
                ExcelModels(ModelID).EventCoordinator.TriggerEvent("Link", LinkTag, ParentGroupForm)


        End Select

    End Sub

#End Region

#Region "Interface Handlers"
    Private Sub ComboOpen(sender, e)

        Dim CB As ComboBoxEdit = sender
        CB.ShowPopup()

    End Sub
    Sub GVPasting(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.ClipboardRowPastingEventArgs)

        Debug.Print("RC" & e.DataRowCount)
        Dim okays = e.GetValidValues
        Dim Errs = e.GetInvalidValues

    End Sub
    Private Sub DEHyperLinkLabel_linkClick(sender As Object, e As DevExpress.Utils.HyperlinkClickEventArgs)

        Dim LinkTag As ElementInterfaceLinkTag = TryCast(sender.Tag, ElementInterfaceLinkTag)

        If LinkTag Is Nothing Then

            MsgBox("Sorry, this link Is Not properly configured")
            Return

        End If

        Dim HyperlinkLabelControl As AbovoDEHyperlinkLabel = sender
        HyperlinkLabelControl.LinkVisited = True

        ExcelModels(ModelID).EventCoordinator.TriggerEvent("Link", LinkTag, ParentGroupForm)

    End Sub
    Private Sub Grid_ButtonClick(sender As Object, e As EventArgs)

        Dim GridTag As AttachedGridCommandButton = TryCast(sender.Tag, AttachedGridCommandButton)

        If GridTag Is Nothing Then

            MsgBox("Sorry, this button Is Not properly configured")
            Return

        End If

        Dim Trans As AbovoTransaction = ExcelModels(ModelID).EventCoordinator.TriggerEvent("GridButton", GridTag, ParentGroupForm)

        If Trans.BError = False Then RebuildAllSections()

    End Sub
    Private Sub Interface_ButtonClick(sender As Object, e As EventArgs)


        ExcelModels(ModelID).EventCoordinator.TriggerEvent("Code", sender.Tag, ParentGroupForm)

    End Sub

    Private Sub Link_ButtonClick(sender As Object, e As EventArgs)

        Dim LinkTag As ElementInterfaceLinkTag = TryCast(sender.Tag, ElementInterfaceLinkTag)

        If LinkTag Is Nothing Then

            MsgBox("Sorry, this link Is Not properly configured")
            Return

        End If

        ExcelModels(ModelID).EventCoordinator.TriggerEvent("Link", sender.Tag, ParentGroupForm)

    End Sub


    Private Sub GridView_MouseWheel(sender As Object, e As MouseEventArgs)

        'AccordionControlM.Top += e.Delta
        DirectCast(e, DXMouseEventArgs).Handled = True

    End Sub
    Private Sub TooltipController_HyperlinkClick(sender As Object, e As DevExpress.Utils.HyperlinkClickEventArgs)

        Dim process As New Process()

        process.StartInfo.FileName = (e.Link)
        process.StartInfo.Verb = "open"
        process.StartInfo.WindowStyle = ProcessWindowStyle.Normal

        Try
            process.Start()
        Catch
            MsgBox("Sorry, the browser could Not be opened")
        End Try

    End Sub
    Sub VerifyDoubleClick(sender As Object, e As MouseEventArgs)



    End Sub
    Private Sub SetGridDisplayText(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs)

        'Dim ColTag As DataColumnTag = e.Column.Tag

        ''Select Case ColTag.DataType

        ''    Case "C"
        ''        Dim ciGB As CultureInfo = New CultureInfo("en-GB")
        ''        e.DisplayText = String.Format(ciGB, "{0:c0}", (e.Value))

        ''End Select

    End Sub

    Private Sub UndoLastAction()

        'ExcelModels(ModelID).UndoLastAction()
        RebuildAllSections()

    End Sub


    Private Sub KeydownListener(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Select Case e.KeyCode

            '  Ctrl + Shift + D
            Case Keys.P And (e.Shift And e.Control And Not e.Alt)
                MsgBox("Print")

                '  Ctrl + Z
            Case Keys.Z And (e.Control And Not e.Shift And Not e.Alt)

                MsgBox("Undo DIT")

                '  Shift and F1
            Case Keys.F1 And (e.Shift And Not e.Control And Not e.Alt)
                MsgBox("Help called")

            Case Keys.V And (e.Control And Not e.Shift And Not e.Alt)

                MsgBox("Paste command")

            Case Keys.C And (e.Control And Not e.Shift And Not e.Alt)

                MsgBox("Copy command")

        End Select

    End Sub

    Private Function GetVGridColumnIndex(ByVal Row As BaseRow) As Integer

        If Row Is Nothing Then Return -1

        Dim ER As EditorRow = TryCast(Row, EditorRow)

        If ER Is Nothing Then Return -1

        Dim FieldName As String = ER.Properties.FieldName

        If String.IsNullOrEmpty(FieldName) Then Return -1

        If FieldName.StartsWith("Col_", StringComparison.OrdinalIgnoreCase) Then

            Dim Result As Integer

            If Integer.TryParse(FieldName.Substring(4), Result) Then

                Return Result

            End If

        End If

        Return -1

    End Function

    Private Sub VGrid_CustomDrawCell(
    ByVal sender As Object,
    ByVal e As DevExpress.XtraVerticalGrid.Events.CustomDrawRowValueCellEventArgs)

        Dim VG As VGridControl = TryCast(sender, VGridControl)

        If VG Is Nothing Then Return
        If e.Row Is Nothing Then Return

        Dim ColTag As DataColumnTag = TryCast(e.Row.Tag, DataColumnTag)

        If ColTag Is Nothing Then Return

        Dim ColIndex As Integer = GetVGridColumnIndex(e.Row)

        If ColIndex < 0 Then Return
        If e.RecordIndex < 0 Then Return

        Dim UBS As AbovoUnboundSource = TryCast(VG.DataSource, AbovoUnboundSource)

        If UBS Is Nothing Then Return

        Dim DSIndex As Integer = UBS.UBSTag.DSIndex

        If DSIndex < 0 Then Return

        If e.RecordIndex >= DataPres.DataSets(DSIndex).DataRows.Count Then Return
        If ColIndex >= DataPres.DataSets(DSIndex).DataColumns.Count Then Return

        Dim CellHandle As CellDataPoint =
        DataPres.DataSets(DSIndex).DataRows(e.RecordIndex).DataCells(ColIndex)

        If CellHandle Is Nothing Then Return


        '----------------------------------------------------------
        ' Dummy column
        '----------------------------------------------------------
        If ColTag.IsDummyColumn Then

            e.Appearance.BackColor = Color.White
            e.Appearance.ForeColor = Color.White

            Return

        End If


        '----------------------------------------------------------
        ' Calculated / read-only cells
        '----------------------------------------------------------
        If ColTag.IsCalculated OrElse ColTag.IsReadOnly Then

            e.Appearance.BackColor = Color.WhiteSmoke
            e.Appearance.ForeColor = AbovoBlue

            Return

        End If


        '----------------------------------------------------------
        ' Rules / locked cells
        '----------------------------------------------------------
        If ColTag.HasRules AndAlso CellHandle.IsLocked Then

            e.Appearance.BackColor = Color.Lavender
            e.Appearance.ForeColor = AbovoBlue

            Return

        End If


        '----------------------------------------------------------
        ' Spreadsheet-derived appearance
        '----------------------------------------------------------
        e.Appearance.BackColor = CellHandle.BGColor
        e.Appearance.ForeColor = CellHandle.FoColor


        '----------------------------------------------------------
        ' Negative values in red
        '
        ' Unlike the XtraGrid version, we can simply alter the
        ' appearance and text and allow VGrid to perform its own
        ' normal painting.
        '----------------------------------------------------------
        If e.CellValue IsNot Nothing AndAlso IsNumeric(e.CellValue) Then

            If CDbl(e.CellValue) < 0 Then

                e.Appearance.ForeColor = Color.Red

                Dim StrToWrite As String = e.CellText

                If Not String.IsNullOrEmpty(StrToWrite) Then

                    If StrToWrite.StartsWith("-") Then

                        StrToWrite = StrToWrite.Substring(1)

                    End If

                    If Not StrToWrite.StartsWith("(") Then

                        StrToWrite = "(" & StrToWrite & ")"

                    End If

                    e.CellText = StrToWrite

                End If

            End If

        End If


        '----------------------------------------------------------
        ' Empty-value pseudo masks
        '----------------------------------------------------------
        If String.IsNullOrEmpty(e.CellText) Then

            Select Case ColTag.ShowDefaultmask

                Case 0
                    e.CellText = "£ ,   "

                Case 1
                    e.CellText = "£ .  "

                Case 3
                    e.CellText = " .  %"

            End Select

        End If

        'IMPORTANT:
        'Don't set e.Handled = True here.
        '
        'We've changed Appearance/CellText and allow DevExpress to
        'paint the cell normally.

    End Sub

    Private Sub VGrid_ValidatingEditor(
    ByVal sender As Object,
    ByVal e As DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventArgs)

        If e.Value Is Nothing Then

            e.Valid = True
            Return

        End If

        Dim VG As VGridControl = TryCast(sender, VGridControl)

        If VG Is Nothing Then Return
        If VG.FocusedRow Is Nothing Then Return

        Dim ColTag As DataColumnTag =
        TryCast(VG.FocusedRow.Tag, DataColumnTag)

        If ColTag Is Nothing Then Return


        Select Case ColTag.DataType

            Case "S"

                e.Valid = True
                Return


            Case Else

                If Not IsNumeric(e.Value) Then

                    e.ErrorText =
                    ColTag.ColumnHeading & " must be numeric."

                    e.Valid = False
                    Return

                End If


                Dim NewValue As Double = CDbl(e.Value)


                If ColTag.MaxVal IsNot Nothing Then

                    If NewValue > CDbl(ColTag.MaxVal) Then

                        e.ErrorText =
                        ColTag.ColumnHeading &
                        " cannot exceed " &
                        ColTag.MaxVal.ToString &
                        "."

                        e.Valid = False
                        Return

                    End If

                End If


                If ColTag.MinVal IsNot Nothing Then

                    If NewValue < CDbl(ColTag.MinVal) Then

                        e.ErrorText =
                        ColTag.ColumnHeading &
                        " must be " &
                        ColTag.MinVal.ToString &
                        " or greater."

                        e.Valid = False
                        Return

                    End If

                End If


                e.Valid = True

        End Select

    End Sub

    Private Sub VGrid_ShowingEditor(
    ByVal sender As Object,
    ByVal e As CancelEventArgs)

        Dim VG As VGridControl = TryCast(sender, VGridControl)

        If VG Is Nothing Then Return
        If VG.FocusedRow Is Nothing Then Return


        Dim ColTag As DataColumnTag =
        TryCast(VG.FocusedRow.Tag, DataColumnTag)

        If ColTag Is Nothing Then

            e.Cancel = True
            Return

        End If


        '----------------------------------------------------------
        ' Calculated/read-only rows
        '----------------------------------------------------------
        If ColTag.IsCalculated OrElse ColTag.IsReadOnly Then

            e.Cancel = True
            Return

        End If


        Dim UBS As AbovoUnboundSource =
        TryCast(VG.DataSource, AbovoUnboundSource)

        If UBS Is Nothing Then Return


        Dim DSIndex As Integer = UBS.UBSTag.DSIndex

        Dim RowIndex As Integer = VG.FocusedRecord

        Dim ColIndex As Integer =
        GetVGridColumnIndex(VG.FocusedRow)


        If RowIndex < 0 OrElse ColIndex < 0 Then

            e.Cancel = True
            Return

        End If


        Dim ThisDataSet As DataCellRange =
        DataPres.DataSets(DSIndex)


        If RowIndex >= ThisDataSet.DataRows.Count Then

            e.Cancel = True
            Return

        End If


        Dim SourceDataRow = ThisDataSet.DataRows(RowIndex)


        '----------------------------------------------------------
        ' Control/spacer rows
        '
        'These are DATA records in your model, despite rows being
        'displayed vertically.
        '----------------------------------------------------------
        If SourceDataRow.IsControlRow Then

            e.Cancel = True
            Return

        End If


        If SourceDataRow.IsSpacerRow Then

            e.Cancel = True
            Return

        End If


        '----------------------------------------------------------
        ' No rules means normal editing
        '----------------------------------------------------------
        If Not ColTag.HasRules Then Return


        Dim SourceDataPoint As CellDataPoint =
        SourceDataRow.DataCells(ColIndex)


        If SourceDataPoint Is Nothing Then

            e.Cancel = True
            Return

        End If


        '----------------------------------------------------------
        ' Spreadsheet fill rule
        '----------------------------------------------------------
        If Me.ActiveSpreadsheet.Range(
        SourceDataPoint.SourceAddress).Fill.PatternType <>
        PatternType.Solid Then

            e.Cancel = True
            Return

        End If


        '----------------------------------------------------------
        ' Locked by rule
        '----------------------------------------------------------
        If SourceDataPoint.IsLocked Then

            e.Cancel = True
            Return

        End If

    End Sub

    Private Sub VGrid_ShownEditor(
    ByVal sender As Object,
    ByVal e As EventArgs)

        Dim VG As VGridControl = TryCast(sender, VGridControl)

        If VG Is Nothing Then Return
        If VG.ActiveEditor Is Nothing Then Return


        If TypeOf VG.ActiveEditor Is DevExpress.XtraEditors.CalcEdit Then

            If DblClickCell Then

                Dim CalcEditor As CalcEdit =
                DirectCast(VG.ActiveEditor, CalcEdit)

                CalcEditor.ShowPopup()

                DblClickCell = False

            End If

        End If

    End Sub

    Private Sub VGrid_CustomCellEditor(
    ByVal sender As Object,
    ByVal e As DevExpress.XtraVerticalGrid.Events.GetCustomRowCellEditEventArgs)

        Dim VG As VGridControl = TryCast(sender, VGridControl)

        If VG Is Nothing Then Return
        If e.Row Is Nothing Then Return


        Dim ColTag As DataColumnTag =
        TryCast(e.Row.Tag, DataColumnTag)

        If ColTag Is Nothing Then Return


        Dim UBS As AbovoUnboundSource =
        TryCast(VG.DataSource, AbovoUnboundSource)

        If UBS Is Nothing Then Return


        Dim DSIndex As Integer = UBS.UBSTag.DSIndex

        Dim DataRowIndex As Integer = e.RecordIndex


        If DataRowIndex < 0 Then Return

        If DataRowIndex >=
        DataPres.DataSets(DSIndex).DataRows.Count Then

            Return

        End If


        Dim SourceRow =
        DataPres.DataSets(DSIndex).DataRows(DataRowIndex)


        'Equivalent of:
        '
        'If ViewTag.DataSet.DataRows(e.RowHandle).IsSpacerRow Then
        '    e.RepositoryItem = Nothing
        'End If

        If SourceRow.IsSpacerRow Then

            e.RepositoryItem = Nothing
            Return

        End If


        'Normally do nothing else here.
        '
        'The standard editor has already been assigned with:
        '
        '    EditorRow.Properties.RowEdit
        '
        'during VGrid construction.

    End Sub

    Private Sub VGrid_CellEditorForEditing(
    ByVal sender As Object,
    ByVal e As DevExpress.XtraVerticalGrid.Events.GetCustomRowCellEditEventArgs)

        Dim VG As VGridControl = TryCast(sender, VGridControl)

        If VG Is Nothing Then Return
        If e.Row Is Nothing Then Return


        Dim ColTag As DataColumnTag =
        TryCast(e.Row.Tag, DataColumnTag)

        If ColTag Is Nothing Then Return


        '----------------------------------------------------------
        ' Existing combo editor wins
        '----------------------------------------------------------
        If ColTag.HasComboEdit Then Return


        Dim UBS As AbovoUnboundSource =
        TryCast(VG.DataSource, AbovoUnboundSource)

        If UBS Is Nothing Then Return


        Dim DSIndex As Integer = UBS.UBSTag.DSIndex


        If e.RecordIndex < 0 OrElse
       e.RecordIndex >= DataPres.DataSets(DSIndex).DataRows.Count Then

            Return

        End If


        Dim SourceRow =
        DataPres.DataSets(DSIndex).DataRows(e.RecordIndex)


        If SourceRow.IsSpacerRow Then

            e.RepositoryItem = Nothing
            Return

        End If


        'Same exclusions as XtraGrid
        If ColTag.DataType = "S" OrElse
       ColTag.DataType = "P" Then

            Return

        End If


        'CustCalcEdit must already be in VertGrid.RepositoryItems.
        e.RepositoryItem = CustCalcEdit


        Select Case ColTag.DataType

            Case "I", "Y"

                CustCalcEdit.Precision = 0


            Case Else

                CustCalcEdit.Precision = 5

        End Select

    End Sub

#End Region

#Region "Draw Handlers"
    Private Sub acCustomDrawAcElement(ByVal sender As Object, ByVal e As CustomDrawElementEventArgs)

        If e.Element.Style = ElementStyle.Item Then

            Dim BBounds As Rectangle = e.ObjectInfo.HeaderBounds
            BBounds.X += (4 * DefaultGridCellPadding)
            e.DrawImage()
            e.DrawContextButtons()
            e.DrawExpandCollapseButton()
            e.DrawHeaderBackground()

            e.Cache.DrawString(e.ObjectInfo.Text, e.Appearance.GetFont(), e.Appearance.GetForeBrush(e.Cache), BBounds, e.Appearance.GetStringFormat())

            Dim pen As Pen = New Pen(AbovoBlue, 2)
            e.Cache.DrawLine(pen, New Point(e.ObjectInfo.HeaderBounds.Left + DefaultGridCellPadding - 2, e.ObjectInfo.HeaderBounds.Top), New Point(e.ObjectInfo.HeaderBounds.Right - DefaultGridCellPadding + 2, e.ObjectInfo.HeaderBounds.Top))
            e.Cache.DrawLine(pen, New Point(e.ObjectInfo.HeaderBounds.Left + DefaultGridCellPadding - 2, e.ObjectInfo.HeaderBounds.Bottom), New Point(e.ObjectInfo.HeaderBounds.Right - DefaultGridCellPadding + 2, e.ObjectInfo.HeaderBounds.Bottom))

            e.Handled = True

        End If

    End Sub

    Private Sub GVCustomDrawFooterCell(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.FooterCellCustomDrawEventArgs)

        If e.Info.DisplayText Is Nothing Then

            e.DefaultDraw()
            Return

        End If

        Dim pen As Pen = New Pen(AbovoBlue, 1)
        Dim StrToWrite As String = ""

        Dim r As New Rectangle With {
            .X = e.Bounds.X,
            .Height = e.Bounds.Height,
            .Y = e.Bounds.Y,
            .Width = e.Bounds.Width - DefaultGridCellPadding
            }

        e.Cache.DrawLine(pen, New Point(e.Bounds.X, e.Bounds.Top), New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Top))
        e.Cache.DrawLine(pen, New Point(e.Bounds.X, e.Bounds.Bottom), New Point(e.Bounds.X + e.Bounds.Width, e.Bounds.Bottom))

        e.Appearance.ForeColor = AbovoBlue
        StrToWrite = e.Info.DisplayText

        If IsNumeric(StrToWrite) Then
            If CDbl(e.Info.DisplayText) < 0 Then

                If Microsoft.VisualBasic.Left(StrToWrite, 1) = "-" Then StrToWrite = Microsoft.VisualBasic.Right(StrToWrite, Len(StrToWrite) - 1)
                e.Appearance.ForeColor = Color.Red
                StrToWrite = "(" & StrToWrite & ")"

            End If
        End If


        e.Appearance.DrawString(e.Cache, StrToWrite, r)
        e.Handled = True

    End Sub
    Private Sub GridView_CustomDrawCell(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs)

        Dim GV As Object = sender

        Dim ViewTag As GridViewTag = GV.Tag

        Dim PadRect As New Rectangle With {
            .X = e.Bounds.X - DefaultGridCellPadding - 2,
            .Width = e.Bounds.Width + (2 * DefaultGridCellPadding) + 4,
            .Height = e.Bounds.Height + 4 + (2 * DefaultGridCellPadding),
            .Y = e.Bounds.Y - (DefaultGridCellPadding + 2)
            }


        'If ViewTag.DataSet.DataRows(e.RowHandle).IsControlRow Then

        '    If e.Column.Tag.IsControlColumn Then
        '        Dim pen As Pen = New Pen(Color.White, IdealGridRowHeight / 2)
        '        e.Appearance.BackColor = Color.Silver
        '        e.DefaultDraw()
        '        e.Cache.DrawLine(pen, New Point(PadRect.Right, PadRect.Top), New Point(PadRect.Left, PadRect.Top))
        '        e.Cache.DrawLine(pen, New Point(PadRect.Right, PadRect.Bottom), New Point(PadRect.Left, PadRect.Bottom))
        '    Else
        '        e.Appearance.BackColor = Color.White
        '    End If


        '    Return

        'End If



        Dim CellHandle As CellDataPoint = ViewTag.DataSet.DataRows(e.RowHandle).DataCells(e.Column.AbsoluteIndex)

        Dim ColTag As DataColumnTag = e.Column.Tag

        If ColTag.IsDummyColumn Then
            e.Appearance.BackColor = Color.White
            e.DefaultDraw()
            e.Handled = True
            Return
        End If

        If sender.IsCellSelected(e.RowHandle, e.Column) Then

            'If CellHandle.IsLocked Then

            '    e.Appearance.ForeColor = Color.White
            '    Dim br As New SolidBrush(AbovoBlue)
            '    e.Cache.FillRectangle(br, PadRect)
            '    e.Handled = True

            '    Return

            'End If

            e.Appearance.BackColor = Color.Beige
            e.Appearance.ForeColor = Color.Black
            e.DefaultDraw()
            e.Handled = True
            Return

        End If

        If e.Appearance.BackColor = AbovoBlueL3 Then

            e.Appearance.ForeColor = Color.White

            If ColTag.HasRules Then

                If CellHandle.IsLocked Then


                    Dim br As New SolidBrush(AbovoBlueL3)
                    e.Cache.FillRectangle(br, PadRect)
                    e.Handled = True

                    Return

                End If

            End If

            e.DefaultDraw()
            Return

        End If

        If ColTag.HasRules Then

            If CellHandle.IsLocked Then

                Dim pen As Pen = New Pen(Color.Red, 2)
                e.Cache.FillRectangle(Brushes.Lavender, PadRect)
                e.Appearance.ForeColor = Color.White
                e.Appearance.DrawString(e.Cache, e.DisplayText, e.Bounds)

                e.Handled = True

                Return

            End If

        End If

        If ColTag.IsCalculated OrElse ColTag.IsReadOnly Then

            If ColTag.IsReadOnly Then e.Appearance.BackColor = Color.WhiteSmoke
            If ColTag.IsReadOnly Then e.Appearance.ForeColor = AbovoBlue
            e.DefaultDraw()
            e.Handled = True
            Return

        End If

        'Dim Deffont As New Font(e.Appearance.GetFont, FontStyle.Bold)

        If GV.FocusedRowHandle = e.RowHandle Then

            Dim pen As Pen = New Pen(AbovoBlue, 3)
            e.Cache.DrawLine(pen, New Point(PadRect.Right, PadRect.Top), New Point(PadRect.Left, PadRect.Top))
            e.Cache.DrawLine(pen, New Point(PadRect.Right, PadRect.Bottom), New Point(PadRect.Left, PadRect.Bottom))

        End If

        If GV.FocusedColumn.AbsoluteIndex = e.Column.AbsoluteIndex And GV.FocusedRowHandle = e.RowHandle Then

            e.Appearance.BackColor = Color.WhiteSmoke
            e.Appearance.ForeColor = Color.Black

        End If

        If ColTag.IsCalculated Then

            e.Appearance.ForeColor = AbovoBlue

        End If

        Dim strToWrite As String = e.DisplayText

        e.Appearance.BorderColor = CellHandle.BGColor 'e.Appearance.BackColor
        e.Appearance.BackColor = CellHandle.BGColor
        e.Appearance.DrawBackground(e.Cache, PadRect)
        e.Appearance.ForeColor = CellHandle.FoColor

        If e.DisplayText = "" Then
            If ColTag.ShowDefaultmask = 0 Then
                e.Appearance.DrawString(e.Cache, "£ ,   ", PadRect)
            ElseIf ColTag.ShowDefaultmask = 1 Then
                e.Appearance.DrawString(e.Cache, "£ .  ", PadRect)
            ElseIf ColTag.ShowDefaultmask = 3 Then
                e.Appearance.DrawString(e.Cache, " .  %", PadRect)
            End If
            e.Handled = True
            Return
        End If

        If IsNumeric(e.CellValue) Then

            If e.CellValue < 0 Then

                If Microsoft.VisualBasic.Left(strToWrite, 1) = "-" Then strToWrite = Microsoft.VisualBasic.Right(strToWrite, Len(strToWrite) - 1)

                e.Appearance.ForeColor = Color.Red

                If Microsoft.VisualBasic.Left(strToWrite, 1) <> "(" Then strToWrite = "(" & strToWrite & ")"

            End If

        End If

        DrawCellBorder(e, CellHandle.BGColor)

        e.Appearance.DrawString(e.Cache, strToWrite, e.Bounds)

        e.Handled = True

    End Sub


#End Region

#Region "Data and Date Event Handlers"
    Sub SingleCellDirtyMarker(ByVal sender As Object, ByVal e As Object)

        Try

            sender.MarkDirty()

        Catch ex As Exception

        End Try

    End Sub
    Sub SingleCellControlValidatingEditor(sender As Object, e As DevExpress.XtraEditors.Controls.BaseEditValidatingEventArgs)

        If sender.editValue = Nothing Then

            Return

        End If

        Dim DataTag As SingleCellDataTag = sender.tag

        Select Case DataTag.DataType

            Case "S"

                Return

            Case Else

                Dim Val As Double = sender.editValue

                If Not IsNumeric(Val) Then

                    e.ErrorText = DataTag.Label & " must be numeric."
                    e.Cancel = True
                    Return

                End If

                If DataTag.MaxValSet Then

                    If Val > CDbl(DataTag.MaxVal) Then

                        e.ErrorText = DataTag.Label & " cannot exceed " & DataTag.MaxVal.ToString & "."
                        e.Cancel = True
                        Return

                    End If

                End If

                If DataTag.MinValSet Then

                    If Val < CDbl(DataTag.MinVal) Then

                        e.ErrorText = DataTag.Label & " must be " & DataTag.MinVal.ToString & " or greater."
                        e.Cancel = True
                        Return

                    End If

                End If


        End Select

    End Sub
    Private Sub UnboundDS_ValueNeeded(ByVal sender As Object, ByVal e As DevExpress.Data.UnboundSourceValueNeededEventArgs)
        DataCallCount += 1
        Debug.Print("XXXXXXXXX - Val Needed - count : " & DataCallCount & "XXXXXXX")
        Dim UDSSender As AbovoUnboundSource = sender
        e.Value = GetDSData(UDSSender.UBSIndex, UDSSender.UBSTag.DSIndex, e.RowIndex, e.PropertyIndex)

    End Sub

    Private Sub vGridControl_CustomUnboundData(ByVal sender As Object, ByVal e As DevExpress.XtraVerticalGrid.Events.CustomDataEventArgs)

        'Dim ColIndex As Integer = e.Row.VisibleIndex

        'If e.IsGetData Then


        '        Dim DP As CellDataPoint = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(PropertyIndex)

        '        If DP Is Nothing Then Return Nothing

        '        Dim DPC As DevExpress.Spreadsheet.Cell = ExcelModels(ModelID).WB.Worksheets(DP.SourceSheet).Cells(DP.SourceAddress)

        '        If DPC.DisplayText = "" Then

        '            If DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag.DataType = "S" Then

        '                Return ""

        '            Else

        '                Return Nothing

        '            End If

        '        End If

        '        Dim ColTag As DataColumnTag = DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag

        '        If ColTag.IsDummyColumn Then Return Nothing

        '        Select Case ColTag.DataType

        '            Case "S"

        '                DP.StringValue = DPC.DisplayText
        '                'SystemLog("DSIndex " & SetDSIndex.ToString & " returning string " & DPC.Value.TextValue & " From " & DP.SourceAddress & " of " & DP.SourceSheet)
        '                Return DPC.DisplayText
        '                Exit Function

        '            Case "B"

        '                DP.BoolValue = DPC.Value.NumericValue
        '                Return DPC.Value.NumericValue
        '                Exit Function

        '            Case "N", "P", "C", "M", "R", "SM"
        '                'SystemLog("UBS Index " & UBSIndex.ToString & " with DSIndex " & SetDSIndex.ToString & " returning " & DPC.Value.NumericValue.ToString & " From " & DP.SourceAddress & " of " & DP.SourceSheet)
        '                DP.RealValue = DPC.Value.NumericValue
        '                Return DPC.Value.NumericValue
        '                Exit Function

        '            Case "I", "Y"

        '                'If DoDataLog Then SystemLog("Returning Integer " & DP.IntValue)
        '                'SystemLog("UBS Index " & UBSIndex.ToString & " with DSIndex " & SetDSIndex.ToString & " returning " & DPC.Value.NumericValue & " From " & DP.SourceAddress & " of " & DP.SourceSheet)
        '                DP.IntValue = DPC.Value.NumericValue
        '                Return CInt(DPC.Value.NumericValue)
        '                Exit Function

        '            Case Else

        '                Return Nothing

        '        End Select


        '        'If Storage.ContainsKey(e.ListSourceRowIndex) Then
        '        '    e.Value = Storage(e.ListSourceRowIndex)
        '        'Else
        '        '    Storage(e.ListSourceRowIndex) = String.Format("Unbound value {0}", e.ListSourceRowIndex)
        '        '    e.Value = Storage(e.ListSourceRowIndex)
        '        'End If



        '    End If

        '    If e.IsSetData Then
        '    'Storage(e.ListSourceRowIndex) = e.Value.ToString()
        'End If


    End Sub
    Private Function GetDSData(ByVal UBSIndex As Integer, SetDSIndex As Integer, ByVal rowIndex As Integer, ByVal PropertyIndex As Integer) As Object

        'If DoDataLog Then SystemLog("Value requested from dataset: " & SetDSIndex.ToString & " Row: " & rowIndex.ToString & " Column: " & PropertyIndex.ToString)
        'If DataPres.DataSets(SetDSIndex).DataRows(rowIndex).IsControlRow = True Then Return Nothing
        'If DataPres.DataSets(SetDSIndex).DataRows(rowIndex).IsSpacerRow = True Then Return Nothing

        Dim DP As CellDataPoint = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(PropertyIndex)

        If DP Is Nothing Then Return Nothing

        Dim DPC As DevExpress.Spreadsheet.Cell = ExcelModels(ModelID).WB.Worksheets(DP.SourceSheet).Cells(DP.SourceAddress)

        If DPC.DisplayText = "" Then

            If DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag.DataType = "S" Then

                Return ""

            Else

                Return Nothing

            End If

        End If

        Dim ColTag As DataColumnTag = DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag

        If ColTag.IsDummyColumn Then Return Nothing

        Select Case ColTag.DataType

            Case "S"

                DP.StringValue = DPC.DisplayText
                'SystemLog("DSIndex " & SetDSIndex.ToString & " returning string " & DPC.Value.TextValue & " From " & DP.SourceAddress & " of " & DP.SourceSheet)
                Return DPC.DisplayText
                Exit Function

            Case "B"

                DP.BoolValue = DPC.Value.NumericValue
                Return DPC.Value.NumericValue
                Exit Function

            Case "N", "P", "C", "M", "R", "SM"
                'SystemLog("UBS Index " & UBSIndex.ToString & " with DSIndex " & SetDSIndex.ToString & " returning " & DPC.Value.NumericValue.ToString & " From " & DP.SourceAddress & " of " & DP.SourceSheet)
                DP.RealValue = DPC.Value.NumericValue
                Return DPC.Value.NumericValue
                Exit Function

            Case "I", "Y"

                'If DoDataLog Then SystemLog("Returning Integer " & DP.IntValue)
                'SystemLog("UBS Index " & UBSIndex.ToString & " with DSIndex " & SetDSIndex.ToString & " returning " & DPC.Value.NumericValue & " From " & DP.SourceAddress & " of " & DP.SourceSheet)
                DP.IntValue = DPC.Value.NumericValue
                Return CInt(DPC.Value.NumericValue)
                Exit Function

            Case Else

                Return Nothing

        End Select

    End Function
    Private Sub UnboundDS_ValuePushed(ByVal sender As Object, ByVal e As DevExpress.Data.UnboundSourceValuePushedEventArgs)

        'SystemLog("Data push - " & Me.Text, Me, "Start")

        Me.Cursor = Cursors.WaitCursor

        Dim UDSSender As AbovoUnboundSource = sender
        Dim ColSent As Integer = Microsoft.VisualBasic.Right(e.PropertyName, Len(e.PropertyName) - 4)

        PushDSData(UDSSender.UBSTag.DSIndex, e.RowIndex, ColSent, e.Value)

        'On Error Resume Next

        If UDSSender.InBandedMode Then

            UDSSender.ActiveGridBandedView.Columns(ColSent).BestFit()
            UDSSender.ActiveGridBandedView.Columns(ColSent).Width = UDSSender.ActiveGridBandedView.Columns(ColSent).Width * 1.15

        ElseIf UDSSender.InVertMode Then

        Else

            Try
                UDSSender.ActiveGridView.Columns(ColSent).BestFit()
                UDSSender.ActiveGridView.Columns(ColSent).Width = UDSSender.ActiveGridView.Columns(ColSent).Width * 1.15
            Catch ex As Exception

            End Try


        End If

        UpdateRules(UDSSender.UBSTag.DSIndex)

        Me.Cursor = Cursors.Default

    End Sub
    Private Sub PushDSData(ByVal SetDSIndex As Integer, ByVal rowIndex As Integer, ByVal ColSent As Integer, Value As Object)


        Dim SentRSDT As String = DataPres.DataSets(SetDSIndex).DataColumns(ColSent).ColumnTag.DataType
        Dim SourceDataPoint As CellDataPoint = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent)

        If Value Is Nothing Then

            'Handle Null Push
            DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).IsEmpty = True


        Else

            SystemLog("Value sent to dataset index: " & SetDSIndex.ToString & " Row: " & rowIndex.ToString & " Column: " & ColSent.ToString)

            DataPres.DataSets(SetDSIndex).IsDirty = True

            DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).IsEmpty = False

            Dim DCE As New DataChangeEvent With {
                        .ModelID = ModelID,
                        .Description = "Cell Updated",
                        .WSName = GetWorkBook(ModelID).Worksheets(SourceDataPoint.SourceSheet).Name,
                        .CellAddress = SourceDataPoint.SourceAddress,
                        .ChangedValue = Value,
                        .DataFormat = SentRSDT,
                        .TimeStamp = Now(),
                        .UserName = Environment.UserName
                    }

            Select Case SentRSDT

                Case "S"

                    DCE.OriginalValue = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).StringValue
                    DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).StringValue = Value

                Case "B"

                    DCE.OriginalValue = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).BoolValue.ToString
                    DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).BoolValue = Value

                Case "N", "P", "C", "D", "M", "R", "SM"

                    DCE.OriginalValue = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).RealValue.ToString
                    DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).RealValue = Value

                Case "I", "Y"

                    DCE.OriginalValue = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).IntValue.ToString
                    DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).IntValue = Value

                Case Else
                    'DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(ColSent).StringValue = Value

            End Select

            ChangeMan.ProcessChange(DCE)

        End If

        RefreshData()

    End Sub
    Sub SingleCell_Value_Push(ByVal sender As Object, ByVal e As Object)

        If sender.editvalue = Nothing Then Exit Sub

        Me.Cursor = Cursors.WaitCursor

        'SystemLog("SDP Update called")
        Dim DataTag As SingleCellDataTag = sender.tag
        Dim OldValue As DevExpress.Spreadsheet.CellValue = DataTag.TargetWorksheet.Cells(DataTag.TargetCell).Value
        Dim OldValueString As String = OldValue.ToString

        If sender.editvalue.ToString = OldValueString Then

            Try

                sender.ClearDirtyFlag()
                Me.Cursor = Cursors.Default
                Return

            Catch ex As Exception

            End Try

        End If

        If TryCast(sender, AbovoDESpinEdit) IsNot Nothing Then

            If sender.IsDirty = False Then
                Me.Cursor = Cursors.Default
                Return
            End If

        End If

        If TryCast(sender, AbovoDETextEdit) IsNot Nothing Then

            If sender.IsDirty = False Then
                Me.Cursor = Cursors.Default
                Return
            End If

        End If

        'SystemLog("Fill=" & GetWorkBook(ModelID).Worksheets(SourceDataPoint.SourceSheet).Cells(SourceDataPoint.SourceAddress).Fill.ToString)

        Dim DCM As New DataChangeEvent With {
                        .ModelID = ModelID,
                        .Description = DataTag.Label & " updated from " & OldValueString & " to " & sender.editvalue.ToString,
                        .WSName = DataTag.TargetWorksheet.Name,
                        .CellAddress = DataTag.TargetCell,
                        .ChangedValue = sender.editvalue.ToString,
                        .OriginalValue = OldValueString,
                        .DataFormat = DataTag.DataType,
                        .TimeStamp = Now(),
                        .UserName = Environment.UserName
                    }

        If ChangeMan.ProcessChange(DCM).BError = True Then

            Try

                Select Case DataTag.DataType

                    Case "S"

                        sender.editvalue = OldValue.TextValue
                        Me.Cursor = Cursors.Default
                        Return

                    Case "D"

                        sender.editvalue = DateTime.FromOADate(OldValue.NumericValue)
                        Me.Cursor = Cursors.Default
                        Return

                    Case Else

                        sender.editvalue = OldValue.NumericValue
                        Me.Cursor = Cursors.Default
                        Return

                End Select

            Catch ex As Exception

            End Try

        End If

        Try
            sender.ClearDirtyFlag()

        Catch ex As Exception

        End Try

        Me.Cursor = Cursors.Default

    End Sub

#Region "Pasting"


#End Region

#Region "GridEvents"
    Private Sub GridView_ValidatingEditor(sender As Object, e As DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventArgs)

        If e.Value = Nothing Then

            e.Valid = True
            Return

        End If

        Dim view As ColumnView = sender
        Dim column As GridColumn = If(TryCast(e, EditFormValidateEditorEventArgs)?.Column, view.FocusedColumn)
        Dim ColTag As DataColumnTag = column.Tag

        Select Case ColTag.DataType

            Case "S"
                e.Valid = True
                Return

            Case Else

                If Not IsNumeric(e.Value) Then

                    e.ErrorText = ColTag.ColumnHeading & " must be numeric."
                    e.Valid = False
                    Return

                End If

                If Not ColTag.MaxVal Is Nothing Then

                    If e.Value > CDbl(ColTag.MaxVal) Then

                        e.ErrorText = ColTag.ColumnHeading & " cannot exceed " & ColTag.MaxVal.ToString & "."
                        e.Valid = False
                        Return
                    End If

                End If

                If ColTag.MinVal IsNot Nothing Then

                    If e.Value < CDbl(ColTag.MinVal) Then

                        e.ErrorText = ColTag.ColumnHeading & " must  be " & ColTag.MinVal.ToString & " or greater."
                        e.Valid = False
                        Return

                    End If

                End If

                e.Valid = True

        End Select

    End Sub

    Private Sub GridView_Event_DoubleClick(ByVal CGVSender As GridView, ByVal e As MouseEventArgs)

        'Dim hitInfo As GridHitInfo = CGVSender.CalcHitInfo(e.Location)
        'Dim ColTag As DataColumnTag = hitInfo.Column.Tag

        DblClickCell = True





    End Sub
    Private Sub GridView_Event_SingleClick(ByVal CGVSender As GridView, ByVal e As MouseEventArgs)

        Dim hitInfo As GridHitInfo = CGVSender.CalcHitInfo(e.Location)

        If e.Button = MouseButtons.Left Then

            CGVSender.ShowEditor()
            Return

        End If

    End Sub


    Sub GridView_FocusedColumnChange(ByVal CGVSender As GridView, ByVal e As FocusedColumnChangedEventArgs)

        CGVSender.ShowEditor()

    End Sub

    Sub GridView_CalcBandedGridRowHeight(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowHeightEventArgs)

        Dim GV As Object = sender

        Dim ViewTag As GridViewTag = GV.Tag

        If ViewTag.DataSet.DataRows(e.RowHandle).IsControlRow Then

            e.RowHeight = CInt(1.5 * IdealGridRowHeight)

        Else

            e.RowHeight = IdealGridRowHeight

        End If

    End Sub
    Sub GridView_ColumnWidthChanged(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Base.ColumnEventArgs)

        Dim view As GridView = TryCast(sender, GridView)
        Dim ViewTag As GridViewTag = view.Tag
        ViewTag.InManualReizeMode = True
        ' e.Column.MinWidth = e.Column.Width

        'Dim ColTag As DataColumnTag = e.Column.Tag
        'If ColTag Is Nothing Then Return
        'If ColTag.IsControlColumn Then
        '    For Each Column In view.Columns
        '        Dim ColT As DataColumnTag = Column.Tag
        '        If ColT IsNot Nothing Then
        '            If Not ColT.IsControlColumn Then
        '                Column.Width = e.Column.Width - (2 * DefaultGridCellPadding)
        '                Exit For
        '            End If
        '        End If
        '    Next
        'End If

    End Sub
    Sub GridView_RowCellEditForEditing(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Grid.CustomRowCellEditEventArgs)

        'If Not DblClickCell Then Return

        Dim view As GridView = TryCast(sender, GridView)
        Dim ViewTag As GridViewTag = view.Tag
        Dim ColTag As DataColumnTag = e.Column.Tag

        If ColTag.HasComboEdit Then Return

        If ViewTag.DataSet.DataRows(e.RowHandle).IsSpacerRow Then e.RepositoryItem = Nothing : Return

        If ColTag.DataType = "S" OrElse ColTag.DataType = "P" Then Return

        e.RepositoryItem = CustCalcEdit

        Select Case ColTag.DataType

            Case "I", "Y"


                CustCalcEdit.Precision = 0
                'RepositoryItem = CalcEdit
                DblClickCell = False


            Case Else

                CustCalcEdit.Precision = 5
                'e.RepositoryItem = CalcEdit

                Return

        End Select

    End Sub
    Private DummyCount As Integer = 0
    Sub GridView_CustomSummaryCalculate(ByVal sender As Object, ByVal e As CustomSummaryEventArgs)

        Dim view As GridView = TryCast(sender, GridView)
        Debug.Print("Custom Summary Event Triggered for " & view.Name)
        If e.IsTotalSummary Then
            Select Case e.SummaryProcess
                    ' Start calculation
                Case CustomSummaryProcess.Start
                    DummyCount = 0
                    ' Consequent calculations
                Case CustomSummaryProcess.Calculate

                    DummyCount += 1

                Case CustomSummaryProcess.Finalize
                    e.TotalValue = DummyCount
            End Select
        End If


    End Sub

    Sub GridView_CustomCellEditor(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Grid.CustomRowCellEditEventArgs)

        Dim view As GridView = TryCast(sender, GridView)
        Dim ViewTag As GridViewTag = view.Tag
        Dim ColTag As DataColumnTag = e.Column.Tag

        If ViewTag.DataSet.DataRows(e.RowHandle).IsSpacerRow Then
            e.RepositoryItem = Nothing
            Return
        End If

        'Dim view As GridView = DirectCast(sender, GridView)
        'Dim ViewTag As GridViewTag = view.Tag
        'Dim ColTag As DataColumnTag = view.FocusedColumn.Tag
        'If ColTag.DataType = "S" Then Return
        'If ColTag.IsCalculated Or ColTag.IsReadOnly Then Return
        'If DblClickCell Then
        '    If ColTag.DataType = "I" Then
        '        view.ActiveEditor.Properties = CustCalcEdit.Properties
        '    End If
        '    If ColTag.DataType = "N" Then
        '        view.ActiveEditor.Properties = CalcEdit.Properties
        '    End If
        '    DblClickCell = False
        '    view.ShowEditor()
        'End If
    End Sub

    'Private Sub GridView_CustomRowCellEditForEditing(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Grid.CustomRowCellEditEventArgs)

    '    Dim view As GridView = DirectCast(sender, GridView)
    '    Dim ViewTag As GridViewTag = view.Tag

    '    If DblClickCell Then

    '        Dim ColTag As DataColumnTag = e.Column.Tag
    '        If ColTag.DataType = "S" Then
    '            DblClickCell = False
    '            Return
    '        End If
    '        If ColTag.IsCalculated Or ColTag.IsReadOnly Then
    '            DblClickCell = False
    '            Return
    '        End If
    '        If ColTag.DataType = "I" Then
    '            e.Column.ColumnEdit = CalcEdit

    '        End If



    '        DblClickCell = False

    '        view.ShowEditor()

    '    End If

    'End Sub
    Private Sub GridView_ShownEditor(ByVal sender As Object, ByVal e As EventArgs)

        Dim gv As GridView = sender

        If TypeOf gv.ActiveEditor Is DevExpress.XtraEditors.CalcEdit Then

            If DblClickCell = True Then

                Dim Clcer As CalcEdit = gv.ActiveEditor
                Clcer.ShowPopup()
                DblClickCell = False


            End If

        End If



    End Sub
    Private Sub GridView_ShowingEditor(ByVal sender As Object, ByVal e As CancelEventArgs)




        Dim view As GridView = DirectCast(sender, GridView)

        Dim ViewTag As GridViewTag = view.Tag

        Dim ColTag As DataColumnTag = view.FocusedColumn.Tag

        If ColTag.IsCalculated Or ColTag.IsReadOnly Then

            e.Cancel = True
            Return

        End If




        If ViewTag.DataSet.DataRows(view.FocusedRowHandle).IsControlRow Then

            e.Cancel = True
            Return

        End If

        If Not ColTag.HasRules Then Return

        Dim SourceDataPoint As CellDataPoint = ViewTag.DataSet.DataRows(view.FocusedRowHandle).DataCells(view.FocusedColumn.AbsoluteIndex)

        If ViewTag.DataSet.DataRows(view.FocusedRowHandle).IsSpacerRow Then


            e.Cancel = True
            Return

        End If
        If Me.ActiveSpreadsheet.Range(SourceDataPoint.SourceAddress).Fill.PatternType <> PatternType.Solid Then
            e.Cancel = True
            Return

        End If

        If SourceDataPoint.IsLocked Then e.Cancel = True

        'Debug.Print(view.FocusedColumn.ColumnEdit.GetType.ToString)

        'Select Case ColTag.DataType

        '        Case "I", "Y"

        '            DblClickCell = False
        '            CustCalcEdit.Precision = 0
        '            'RepositoryItem = CalcEdit
        '            DblClickCell = False
        '            SendingView.ShowEditor()

        '        Case Else

        '            DblClickCell = False
        '            CustCalcEdit.Precision = 5
        '            'e.RepositoryItem = CalcEdit
        '            DblClickCell = False
        '            SendingView.ShowEditor()
        '            Return

        '    End Select
        'End If

    End Sub

#End Region

    Sub ShowCalcEdit(ByVal SendingView As GridView)

        Dim ViewTag As GridViewTag = SendingView.Tag
        Dim ColTag As DataColumnTag = SendingView.FocusedColumn.Tag

        If DblClickCell Then

            Select Case ColTag.DataType

                Case "I", "Y"

                    DblClickCell = False
                    CustCalcEdit.Precision = 0
                    'RepositoryItem = CalcEdit
                    DblClickCell = False
                    SendingView.ShowEditor()

                Case Else

                    DblClickCell = False
                    CustCalcEdit.Precision = 5
                    'e.RepositoryItem = CalcEdit
                    DblClickCell = False
                    SendingView.ShowEditor()
                    Return

            End Select

        Else

            Select Case ColTag.DataType

                Case "I", "Y"

                    'e.RepositoryItem = ColTag.DefaultTextEditor

                Case "N", "P", "C", "M"

                    'e.RepositoryItem = ColTag.DefaultTextEditor

            End Select



        End If

    End Sub

    Public Sub RunAction(ActToken As ActionToken)

        Select Case ActToken.ActionType

            Case "NREDIT"

            Case "NRRIbyCOL"

                Dim args As New XtraInputBoxArgs()

                args.Caption = "Abovo Summit - Add Columns"
                args.Prompt = "Add how many columns?"
                args.DefaultButtonIndex = 0

                Dim editor As New DevExpress.XtraEditors.TextEdit()

                With editor.Properties

                    .UseAdvancedMode = DevExpress.Utils.DefaultBoolean.True
                    .MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
                    .MaskSettings.Set("mask", "n0")

                End With

                args.Editor = editor
                args.DefaultResponse = "5"

                Dim result As Object = XtraInputBox.Show(args)

                If result = Nothing Then Return

                Debug.Print("Result: " & result.ToString)

                Dim NewRows As Integer = CInt(result)

                If NewRows > 0 Then

                    Me.Cursor = Cursors.WaitCursor

                    WorkbookManager.InsertRows(ModelID, ActToken.ActionStrData1, NewRows)

                    TransDBManager.CheckTransDBActions(ModelID, ActToken.ActionStrData1)

                    RebuildAllSections()
                    ResizeFonts()
                    UpdateAllRules()

                    Me.Cursor = Cursors.Default

                End If

                editor.Dispose()
                editor = Nothing

            Case "NRRI"

                Dim args As New XtraInputBoxArgs()

                args.Caption = "Abovo Summit - add lines"
                args.Prompt = "Add how many lines?"
                args.DefaultButtonIndex = 0

                Dim editor As New DevExpress.XtraEditors.TextEdit()

                With editor.Properties

                    .UseAdvancedMode = DevExpress.Utils.DefaultBoolean.True
                    .MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
                    .MaskSettings.Set("mask", "n0")

                End With

                args.Editor = editor
                args.DefaultResponse = "5"

                Dim result As Object = XtraInputBox.Show(args)

                If result = Nothing Then Return

                Debug.Print("Result: " & result.ToString)

                Dim NewRows As Integer = CInt(result)

                If NewRows > 0 Then

                    Me.Cursor = Cursors.WaitCursor

                    WorkbookManager.InsertRows(ModelID, ActToken.ActionStrData1, NewRows)

                    TransDBManager.CheckTransDBActions(ModelID, ActToken.ActionStrData1)

                    RebuildAllSections()
                    ResizeFonts()
                    UpdateAllRules()

                    Me.Cursor = Cursors.Default

                End If



                editor.Dispose()
                editor = Nothing

            Case "NRCI"

                Dim args As New XtraInputBoxArgs()

                args.Caption = "Abovo Summit - Add Columns"
                args.Prompt = "Add how many columns?"
                args.DefaultButtonIndex = 0

                Dim editor As New DevExpress.XtraEditors.TextEdit()

                With editor.Properties

                    .UseAdvancedMode = DevExpress.Utils.DefaultBoolean.True
                    .MaskSettings.Set("MaskManagerType", GetType(DevExpress.Data.Mask.NumericMaskManager))
                    .MaskSettings.Set("mask", "n0")

                End With

                args.Editor = editor
                args.DefaultResponse = "5"

                Dim result As Object = XtraInputBox.Show(args)

                If result = Nothing Then Return

                Debug.Print("Result: " & result.ToString)

                Dim NewCols As Integer = CInt(result)

                If NewCols > 0 Then

                    Me.Cursor = Cursors.WaitCursor

                    WorkbookManager.InsertColumns(ModelID, ActToken.ActionStrData1, NewCols)

                    TransDBManager.CheckTransDBActions(ModelID, ActToken.ActionStrData1)

                    RebuildAllSections()
                    ResizeFonts()
                    UpdateAllRules()

                    Me.Cursor = Cursors.Default

                End If

                editor.Dispose()
                editor = Nothing

        End Select

    End Sub

    Sub UpdateAllRules()

        If DataPres.DataSets.Count = 0 Then Exit Sub

        Dim x As Integer
        For x = 0 To DataPres.DataSets.Count - 1
            If DataPres.DataSets(x).HasRules Then
                DataPres.DataSets(x).UpdateLocks()
            End If
        Next
    End Sub
    Sub UpdateCalcs(DS As Integer)

        If DataPres.DataSets(DS).HasCalcs Then
            DataPres.DataSets(DS).UpdateCalcs()
        End If

    End Sub
    Sub UpdateRules(DS As Integer)

        If DataPres.DataSets(DS).HasRules Then
            DataPres.DataSets(DS).UpdateLocks()
        End If

    End Sub

#End Region

#Region "Interface Extensions"
    Public Class ScrollRediverter : Implements IMessageFilter

        Dim RecipientControl As System.Windows.Forms.Control
        Public Sub New(Recipient As System.Windows.Forms.Control)

            RecipientControl = Recipient

        End Sub
        Public Function PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage

            If m.Msg = DevExpress.Utils.Drawing.Helpers.MSG.WM_MOUSEWHEEL Then

                'SystemLog("TargetHandle: " & RecipientControl.Handle.ToString & " (" & RecipientControl.Name & ") SourceHandle: " & m.HWnd.ToString & ", Msg: " & m.Msg.ToString & ", WParam: " & m.WParam.ToString & ", LParam: " & m.LParam.ToString)

                If RecipientControl.FindForm().RectangleToScreen(RecipientControl.Bounds).Contains(Form.MousePosition) Then

                    NativeMethods.SendMessage(RecipientControl.Handle, m.Msg, m.WParam, m.LParam)

                    Return False

                Else

                    Return False

                End If

            Else

                Return False

            End If

        End Function

    End Class
#End Region
    Private Sub ColumnHeaderEmbededComboChanged(ByVal sender As Object, ByVal e As EventArgs)

        Dim SourceTag As InColumnEditorTagCombo = sender.tag

        Dim NewVal As Object = sender.editvalue

        If sender.editvalue = SourceTag.LastEditorValue Then Return



        Dim DCM As New DataChangeEvent(6, True) With {
                        .ModelID = ModelID,
                        .Description = "Column " & SourceTag.EditingNRIndexPosition.ToString & " updated from " & SourceTag.LastEditorValue.ToString & " to " & NewVal.ToString,
                        .TargetNR = SourceTag.EditingNRName,
                        .TargetNRIndex = SourceTag.EditingNRIndexPosition,
                        .ChangedValue = sender.editvalue.ToString,
                        .NROrientation = SourceTag.NROrientation,
                        .OriginalValue = SourceTag.LastEditorValue,
                        .DataFormat = SourceTag.EditorFormat,
                        .TimeStamp = Now(),
                        .UserName = Environment.UserName
                    }
        Try

            If ChangeMan.ProcessChangeByNRAddressing(DCM).BError = True Then

                'Select Case SourceTag.EditorFormat

                '    Case "S"

                '        sender.editvalue = OldValue.TextValue
                '        Me.Cursor = Cursors.Default
                '        Return

                '    Case "D"

                '        sender.editvalue = DateTime.FromOADate(OldValue.NumericValue)
                '        Me.Cursor = Cursors.Default
                '        Return

                '    Case Else

                '        sender.editvalue = OldValue.NumericValue
                '        Me.Cursor = Cursors.Default
                '        Return

                'End Select

            Else

                SourceTag.LastEditorValue = NewVal
                SourceTag.LinkedComboBoxEdit.EditValue = NewVal
                SourceTag.InPlaceColumnHelper.EditValue = NewVal

                UpdateAllRules()

            End If

        Catch ex As Exception

        End Try



    End Sub
    Private Sub ColumnHeaderEmbededDateEChanged(ByVal sender As Object, ByVal e As EventArgs)

        Dim SourceTag As InColumnEditorTagDateEdit = sender.tag

        Dim NewVal As Object = sender.editvalue

        Dim RevColumn As Integer = -1


        If SourceTag.LastEditorValue.ToString = "" Then

            If sender.editvalue.ToString = "" Then
                Return
            Else
                RevColumn = 1
            End If

        Else

            If sender.editvalue = SourceTag.LastEditorValue Then
                Return
            Else
                RevColumn = 2
            End If

        End If

        Dim DCM As New DataChangeEvent(6, True) With {
                        .ModelID = ModelID,
                        .Description = "Column " & SourceTag.EditingNRIndexPosition.ToString & " updated from " & SourceTag.LastEditorValue.ToString & " to " & NewVal.ToString,
                        .TargetNR = SourceTag.EditingNRName,
                        .TargetNRIndex = SourceTag.EditingNRIndexPosition,
                        .ChangedValue = sender.editvalue.ToString,
                        .NROrientation = SourceTag.NROrientation,
                        .OriginalValue = SourceTag.LastEditorValue,
                        .DataFormat = SourceTag.EditorFormat,
                        .TimeStamp = Now(),
                        .UserName = Environment.UserName
                    }
        Try

            If ChangeMan.ProcessChangeByNRAddressing(DCM).BError = True Then

                'Select Case SourceTag.EditorFormat

                '    Case "S"

                '        sender.editvalue = OldValue.TextValue
                '        Me.Cursor = Cursors.Default
                '        Return

                '    Case "D"

                '        sender.editvalue = DateTime.FromOADate(OldValue.NumericValue)
                '        Me.Cursor = Cursors.Default
                '        Return

                '    Case Else

                '        sender.editvalue = OldValue.NumericValue
                '        Me.Cursor = Cursors.Default
                '        Return

                'End Select

            Else

                SourceTag.LastEditorValue = NewVal
                SourceTag.LinkedDateBoxEdit.EditValue = NewVal
                SourceTag.InPlaceColumnHelper.EditValue = NewVal

                UpdateAllRules()

            End If

        Catch ex As Exception

        End Try



    End Sub


    Sub ResizeFonts()

        'Exit Sub


        Dim HaveDoneGrid As Boolean = False

        If ParentGroupForm Is Nothing Then
            Scalefactor = Me.Width / 1700
        Else
            Scalefactor = ParentGroupForm.Width / 1700
        End If


        Dim NewFont As Font = GetFont("Small", Scalefactor)

        For Each control In Me.Controls

            control.Font = NewFont

        Next

        If GridViewCount < 0 Then GoTo BandedGridViews

        If Me.UsedGridVIEWS.Length > 0 Then

            For Each GV In Me.UsedGridVIEWS

                GV.CheckLoaded()

                If GV.State = GridState.Editing Then
                    GV.CloseEditor()
                End If

                GV.BeginUpdate()

                For Each ap As AppearanceObject In GV.Appearance

                    ap.Font = NewFont

                Next

                Dim ViewTag As GridViewTag = GV.Tag

                If Not ViewTag.HaveProcessedColumns Then Formatter.ProcessGVColumWidths(GV, Me)

                'If Not HaveDoneGrid Then

                GV.BestFitColumns()
                HaveDoneGrid = True
                Dim GVI As GridViewInfo = GV.GetViewInfo()
                IdealGridRowHeight = GVI.MinRowHeight

                'End If

                GV.EndUpdate()

            Next

        End If

BandedGridViews:

        If BandGridViewsCount < 0 Then GoTo TextBoxes

        If Me.UsedBANDedGridVIEWS.Length > 0 Then

            For Each BGV In Me.UsedBANDedGridVIEWS


                BGV.CheckLoaded()

                BGV.BeginUpdate()

                For Each ap As AppearanceObject In BGV.Appearance

                    ap.Font = NewFont

                Next

                Dim ViewTag As GridViewTag = BGV.Tag

                If Not ViewTag.HaveProcessedColumns Then Formatter.ProcessGVColumWidths(BGV, Me)


                'Formatter.ProcessGVColumWidths(BGV, Me)

                'If Not HaveDoneGrid Then

                BGV.BestFitColumns()
                HaveDoneGrid = True
                Dim GVI As GridViewInfo = BGV.GetViewInfo()
                IdealGridRowHeight = GVI.MinRowHeight

                'End If

                BGV.EndUpdate()

            Next

        End If

TextBoxes:

        If TextEditCount < 0 Then GoTo ADETextEdits

        If Me.TextBoxes.Length > 0 Then

            For Each TB In Me.TextBoxes


                TB.Font = NewFont


            Next

        End If

ADETextEdits:

        If ADETextEditsCount < 0 Then GoTo ADICombos

        If Me.ADETextEdits.Length > 0 Then

            For Each ADETB In Me.ADETextEdits

                ADETB.Font = NewFont

            Next

        End If

ADICombos:

        If ADEComboBoxesCount < 0 Then GoTo Combos



        If Me.ADEComboBoxes.Length > 0 Then
            For Each adicb In Me.ADEComboBoxes
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                adicb.Properties.Appearance.Font = NewFont

                ' Next
            Next
        End If




Combos:
        If CombosCount < 0 Then GoTo ADELabels



        If Me.Combos.Length > 0 Then
            For Each cb In Me.Combos
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                cb.Properties.Appearance.Font = NewFont

                ' Next
            Next
        End If

        'Private ADELabels() As AbovoDELabel
        'Private ADELabelsCount As Integer = -1
ADELabels:

        If ADELabelsCount < 0 Then GoTo Labels

        If Me.ADELabels.Length > 0 Then



            For Each adelb In Me.ADELabels
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                adelb.Font = NewFont

                ' Next
            Next

        End If
Labels:

        If LabelsCount < 0 Then GoTo ADEDateBoxes



        If Me.Labels.Length > 0 Then

            For Each lb In Me.Labels
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                lb.Font = NewFont

                ' Next
            Next

        End If

ADEDateBoxes:

        If ADEDateEditsCount < 0 Then GoTo DateBoxes

        If Me.ADEDateEdits.Length > 0 Then


            For Each ADDEE In Me.ADEDateEdits
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                ADDEE.Properties.Appearance.Font = NewFont

                ' Next
            Next

        End If
DateBoxes:

        If DateEditCount < 0 Then GoTo HLinks



        If Me.DateBoxes.Length > 0 Then

            For Each ADbE In Me.DateBoxes
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                ADbE.Properties.Appearance.Font = NewFont

                ' Next
            Next

        End If
        'Private ADEHyper           links() As AbovoDEHyperlinkLabel
        'Private             As Integer = -1
HLinks:

        If ADEHyperlinksCount < 0 Then Exit Sub


        If Me.ADEHyperlinks.Length > 0 Then



            For Each ADhl In Me.ADEHyperlinks
                'For Each ap As AppearanceObject In cb.Properties.Appearance

                ADhl.Font = NewFont

                ' Next
            Next

        End If

TPans:

        If Me.TPs IsNot Nothing Then Exit Sub

        If Me.TPs.Length > 0 Then

            For Each tp In Me.TPs

                tp.Width = ParentGroupForm.Width - 40
                tp.Refresh()

            Next

        End If



        'Me.GridControlStockGrid.Font = GetFont("Small", Me.Scalefactor)

        'Me.GridViewStockNumbers.Appearance.OddRow.Font = GetFont("Small", Me.Scalefactor)
        'Me.GridViewStockNumbers.Appearance.ViewCaption.Font = GetFont("Small", Me.Scalefactor)

        'Dim x As Integer

        'For x = 0 To GridViewStockNumbers.Columns.Count - 1

        '    GridViewStockNumbers.Columns(x).AppearanceCell.Font = GetFont("Small", Me.Scalefactor)
        '    GridViewStockNumbers.Columns(x).AppearanceHeader.Font = GetFont("Small", Me.Scalefactor, True)

        'Next

        'RepositoryItemComboBoxOwnedManaged.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        'RepositoryItemLookUpEditSOCIStockType.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        'RepositoryItemLookUpEditSOCIRentType.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        'RepositoryItemIntegerEdit.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        'RepositoryItemComboBoxSOCIStockType.Appearance.Font = GetFont("Small", Me.Scalefactor, True)

        'SystemLog("SF=" & Scalefactor)

        'colPropertyStockDescription1.Width = GridControlStockGrid.Width * 0.25
        'colPropertyOwnedManaged1.Width = GridControlStockGrid.Width * 0.125
        'colPropertySOCIStockType1.Width = GridControlStockGrid.Width * 0.125
        'colPropertySOCIRentType1.Width = GridControlStockGrid.Width * 0.125
        'colPropertyCurrentStockNumbers1.Width = GridControlStockGrid.Width * 0.12
        'colPropertyNewLettings1.Width = GridControlStockGrid.Width * 0.12
        'colPropertyTotalOpeningStockCalc1.Width = GridControlStockGrid.Width * 0.12


        'colPropertyPreBPlanStartDateNewBuild1.Width = GridControlStockGrid.Width * 0.1
        'colPropertyPreBPlanStartDateDemolitions1.Width = GridControlStockGrid.Width * 0.1
        'colPropertyPreBPlanStartDateRTBs1.Width = GridControlStockGrid.Width * 0.1
        'colPropertyPreBPlanStartDateOtherDisposals1.Width = GridControlStockGrid.Width * 0.1
        'colPropertyExistingStocksCalc1.Width = GridControlStockGrid.Width * 0.1

        'colPropertyNewLettings1.Width = GridControlStockGrid.Width * 0.1


        'Me.hideContainerRight.Font = GetFont("Small", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.ActiveTab.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDo                 cking.HidePanelButton.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.HidePanelButtonActive.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaption.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaptionActive.Font = GetFont("Medium", Me.ScaleFactor)


        ''XtraTabControlMainNavigator.Appearance.FontSizeDelta = MediumFontSize - XtraTabControlMainNavigator.Appearance.Font.Size
        'Me.BarAndDockingControllerMainScreen.AppearancesBar.Dock.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.HidePanelButton.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.BarAndDockingControllerMainScreen.AppearancesDocking.PanelCaption.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.XtraTabControlMainNavigator.AppearancePage.HeaderActive.Font = GetFont("Medium", Me.ScaleFactor, False, True)
        'Me.XtraTabControlMainNavigator.AppearancePage.Header.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.XtraTabControlMainNavigator.Appearance.Font = GetFont("Medium", Me.ScaleFactor)
        'Me.XtraTabControlMainNavigator.AppearancePage.HeaderHotTracked.Font = GetFont("Medium", Me.ScaleFactor)

        'WindowsUIButtonPanelExitHelp.Font = GetFont("Small", Me.ScaleFactor)
        'Me.WindowsUIButtonPanelBPActions.Font = GetFont("Small", Me.ScaleFactor)

        'WindowsUIButtonPanelOpenCompare.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelOpenCompare.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelOpenCompare.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)

        'WindowsUIButtonPanelBPActions.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelBPActions.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelBPActions.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)

        'WindowsUIButtonPanelExitHelp.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelExitHelp.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelExitHelp.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)

        'WindowsUIButtonPanelSaveClose.AppearanceButton.Normal.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelSaveClose.AppearanceButton.Hovered.Font = GetFont("Small", Me.ScaleFactor)
        'WindowsUIButtonPanelSaveClose.AppearanceButton.Pressed.Font = GetFont("Small", Me.ScaleFactor)
        'Me.GroupBoxFileActions.Font = GetFont("Small", Me.ScaleFactor)
        'Me.WindowsUIButtonPanelBPActions.ButtonBackgroundImages
        'SystemLog("Small font size:" & Me.XtraTabControlMainNavigator.AppearancePage.HeaderHotTracked.Font.SizeInPoints.ToString)

    End Sub
    Sub ResizeControls()

        Dim SetWidth As Integer = Me.Width * 0.17
        ScaleUnits = Me.Width * 0.007

        'PictureBoxAbovoLogo.Top = ScaleUnits
        'PictureBoxAbovoLogo.Left = ScaleUnits
        'PictureBoxAbovoLogo.Width = SetWidth
        'PictureBoxAbovoLogo.Height = CInt(PictureBoxAbovoLogo.Width * 0.483)

        'DockPanelSettings.Width = SetWidth
        'SystemLog("GBPDHe:" & GroupBoxProgramDetails.Height)
        'SystemLog("WUIBTop:" & WindowsUIButtonPanelExitHelp.Top)
        'SystemLog("ABLBot:" & PictureBoxAbovoLogo.Bottom)
        'WindowsUIButtonPanelExitHelp.Left = ScaleUnits
        'GroupBoxProgramDetails.Width = SetWidth
        'XtraTabControlMainNavigator.Top = ScaleUnits
        'XtraTabControlMainNavigator.Left = PictureBoxAbovoLogo.Right + ScaleUnits
        'XtraTabControlMainNavigator.Width = Me.Width - SetWidth - (5 * ScaleUnits) - hideContainerRight.Width
        'XtraTabControlMainNavigator.Height = Me.Height - (6 * ScaleUnits)
        'XtraTabPageMainHABP.Height = XtraTabControlMainNavigator.PageClientBounds.Height
        'WindowsUIButtonPanelOpenCompare.Left = ScaleUnits
        'WindowsUIButtonPanelOpenCompare.Top = 3 * ScaleUnits
        'GroupBoxFileActions.Left = ScaleUnits
        'GroupBoxFileActions.Top = WindowsUIButtonPanelOpenCompare.Bottom + ScaleUnits
        'GroupBoxFileActions.Width = XtraTabControlMainNavigator.Width - (2 * ScaleUnits)
        'GroupBoxFileActions.Height = XtraTabPageMainHABP.Height - WindowsUIButtonPanelExitHelp.Height - (4 * ScaleUnits)
        'WindowsUIButtonPanelExitHelp.Top = XtraTabControlMainNavigator.Bottom - WindowsUIButtonPanelExitHelp.Height
        'WindowsUIButtonPanelExitHelp.Width = SetWidth
        'WindowsUIButtonPanelExitHelp.Left = ScaleUnits
        'WebBrowserBPInfo.Top = (2 * ScaleUnits)
        'WebBrowserBPInfo.Width = GroupBoxFileActions.Width - WindowsUIButtonPanelSaveClose.Width - (3 * ScaleUnits)
        'WebBrowserBPInfo.Height = GroupBoxFileActions.Height - WindowsUIButtonPanelBPActions.Height - (4 * ScaleUnits)
        'WindowsUIButtonPanelSaveClose.Left = WebBrowserBPInfo.Right + ScaleUnits
        'WindowsUIButtonPanelSaveClose.Top = WebBrowserBPInfo.Top
        'WindowsUIButtonPanelSaveClose.Height = GroupBoxFileActions.Height
        'WindowsUIButtonPanelOpenCompare.Width = XtraTabControlMainNavigator.PageClientBounds.Width - (2 * ScaleUnits)
        'WindowsUIButtonPanelBPActions.Top = WebBrowserBPInfo.Bottom + ScaleUnits
        'WindowsUIButtonPanelBPActions.Width = WebBrowserBPInfo.Width
        'GroupBoxProgramDetails.Top = PictureBoxAbovoLogo.Bottom + ScaleUnits
        'GroupBoxProgramDetails.Left = ScaleUnits
        'GroupBoxProgramDetails.Height = WindowsUIButtonPanelExitHelp.Top - PictureBoxAbovoLogo.Bottom - (2 * ScaleUnits)
        'SetBrowserText()

    End Sub

    Public Sub ResizeControlsCommand()

        ResizeFonts()

    End Sub
    Sub AnalyseGrids()



    End Sub


    Class AbovoGridRespoitaryCombo

        Public ID As Integer
        Public RepoistaryID As String
        Public Combo As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
        Public GridID As Integer
        Public ModelID As Integer

        Public Sub RefreshDataSource()

            Combo.Items.Clear()
            Dim TempCombo As DevExpress.XtraEditors.Repository.RepositoryItemComboBox = RepositaryItems.GetEditor(RepoistaryID, ModelID).RetCombo

            For Each it In TempCombo.Items

                Combo.Items.Add(it)

            Next

            TempCombo.Dispose()
            TempCombo = Nothing

        End Sub

        Class AbovoTabPage

            Public Index As Integer
            Public TabPage As XtraTabPage
            Public GridAndReposss() As GridAndReposs
            Public GridControls() As GridControl
            Public GridCountrolRepossIndex As Integer = -1
            Public x As Integer = 0

            Public Sub AddGrid(Grid As GridControl)

                GridCountrolRepossIndex += 1

                ReDim Preserve GridAndReposss(GridCountrolRepossIndex)
                GridAndReposss(GridCountrolRepossIndex) = New GridAndReposs With {.Grid = Grid}


            End Sub

            Public Sub AddRepCombo(AGR As AbovoGridRespoitaryCombo)

                GridAndReposss(GridCountrolRepossIndex).AddRepCombo(AGR)

            End Sub

            Public Sub UpdateReposes()

                If GridAndReposss Is Nothing Then Exit Sub

                For Each GR In GridAndReposss

                    If GR.HasRepos Then

                        GR.Grid.BeginUpdate()
                        GR.Grid.FocusedView.BeginUpdate()

                        For Each RepCombo In GR.Reposs
                            RepCombo.RefreshDataSource()
                        Next

                        GR.Grid.FocusedView.EndUpdate()
                        GR.Grid.EndUpdate()

                    End If

                Next

            End Sub
            Class GridAndReposs

                Public Grid As GridControl
                Public Reposs() As AbovoGridRespoitaryCombo
                Private RepossIndex As Integer = -1
                Public HasRepos As Boolean = False

                Public Sub AddRepCombo(AGR As AbovoGridRespoitaryCombo)

                    RepossIndex += 1
                    ReDim Preserve Reposs(RepossIndex)
                    Reposs(RepossIndex) = AGR
                    HasRepos = True

                End Sub

            End Class

        End Class

    End Class

    Private Sub XtraTabControlNewGIT_SelectedPageChanged(sender As Object, e As TabPageChangedEventArgs) Handles XtraTabControlNewGIT.SelectedPageChanged

        If AbovoTabPages Is Nothing Then Exit Sub
        If AbovoTabPages.Length = 0 Then Exit Sub

        For Each TP As AbovoGridRespoitaryCombo.AbovoTabPage In AbovoTabPages
            If TP.TabPage Is e.Page Then
                TP.UpdateReposes()
                Exit Sub
            End If

        Next

    End Sub
    Public Sub UpdateTabPage()
        Dim CurrTab As XtraTabPage = XtraTabControlNewGIT.SelectedTabPage

        If AbovoTabPages Is Nothing Then Exit Sub
        If AbovoTabPages.Length = 0 Then Exit Sub
        For Each TP As AbovoGridRespoitaryCombo.AbovoTabPage In AbovoTabPages
            If TP.TabPage Is CurrTab Then
                TP.UpdateReposes()
                Exit Sub
            End If
        Next
    End Sub

    Private Sub DataInterfaceTemplate_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        Debug.Print("--------------")

        If GridControls Is Nothing Then GoTo DoLabels
        If GridControls.Length = 0 Then GoTo DoLabels

        Dim maxSize As New Size(Me.Width, Me.Height - 200)

        For Each GC In GridControls

            GC.Size = GC.CalcBestSize(maxSize, False)

        Next

DoLabels:

        If Labels Is Nothing Then GoTo DoTPs
        If Labels.Length = 0 Then GoTo DoTPs

        For Each LB In Labels

            LB.PerformLayout()

        Next

DoTPs:

        For Each tp In TPs

            tp.PerformLayout()
            Dim x As Integer = 0
            For Each TPR In tp.Rows

                Debug.Print("Row " & x.ToString & " - " & TPR.Height.ToString)
                x += 1

            Next
            For Each control In tp.Controls
                Debug.Print("Control " & control.Name & " - " & control.Height.ToString)
            Next

        Next

    End Sub

End Class


