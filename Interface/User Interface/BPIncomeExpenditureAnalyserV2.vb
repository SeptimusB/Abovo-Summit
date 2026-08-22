Option Infer On

#Region "Imports"


Imports System.Globalization
Imports System.Text.RegularExpressions
Imports Abovo
Imports Abovo.CustomGrid
Imports Abovo.ExportServices
Imports Abovo.FileManager
Imports Abovo.GeneralFunctions
Imports Abovo.LogDebugDev
Imports DevExpress.ClipboardSource.SpreadsheetML
Imports DevExpress.Data
Imports DevExpress.DataAccess.Native.Sql.ConnectionProviders
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils
Imports DevExpress.Utils.Text
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.XtraBars.Navigation
Imports DevExpress.XtraEditors
Imports DevExpress.XtraExport.Helpers
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Drawing
Imports DevExpress.XtraGrid.Localization
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports DevExpress.XtraPrinting
Imports DevExpress.XtraSpreadsheet.Model

#End Region
Public Class BPIncomeExpenditureAnalyserV2

    Private DataPM As PresentationManager
    Private DataPres As PresentationManager.DataPresentation
    Public ModelID As Integer
    Private PresID As Integer
    Private GSID As Integer
    Private CSID As Integer
    Private rs As New Resizer
    Private BIsDirty As Boolean
    Private Grid1ExpandedView As Boolean
    Private CalcEngID As Integer = -1
    Private ParentGIT As GroupInterfaceTemplate
    Private ScaleUnits As Single
    Private Scalefactor As Single
    Private DataSources() As AbovoUnboundSource
    Private DataSourceCount As Integer
    Private PresentedDS As Abovo.DataObject.DataCellRange
    Private PresentedColumn As Abovo.DataObject.SheetDataColumn
    Private PropertyArray() As UnboundSourceProperty
    Private PropertiesCount As Integer
    Private PropertyList As IEnumerable(Of UnboundSourceProperty)
    Private PropType As System.Type
    Private ColList As List(Of String)
    Private ColCount As Integer = 0
    Private ColName As String
    Private GridControls() As GridControl
    Private GridCount As Integer = -1
    Private AmInactiveState As Boolean = False

    Private TransDBDataRange As DevExpress.Spreadsheet.CellRange

    Private LastRowWasRT As Boolean = False
    Private AcControl As AccordionControl
    Private AcElements() As AccordionControlElement
    Private AcElementlist As List(Of AccordionControlElement)
    Private AcControlCount As Integer = -1
    Private AcElementCount As Integer = -1
    Private AcContainers() As AccordionContentContainer
    Private hyperlinkLabelControls() As HyperlinkLabelControl
    Private AcContainersCount As Integer = -1
    Private UsedGridViews(-1) As GridView
    Private GridViewCount As Integer = -1
    Private Formatter As ObjectFormatter
    Private CurrChartWS As DevExpress.Spreadsheet.Worksheet
    Private DSAnalDataRange As RangeDataSource
    Private HasSnapshots As Boolean = False
    Private gridInfo As GridViewInfo = Nothing
    Private ActiveGridView As CustomGridView
    Private ActiveGridWrapper As CustomGridWrapper
    Private ActiveGridControl As GridControl

    'Things for hyperlinkLabelControls#group LastPaintedFooterColour

    Private LastFont1 As Font
    Private LastPaintedFooterColour As Color
    Private LastGroupHeading1 As String = ""
    Private LastGroupHeading2 As String = ""
    Private LastGroupHeading3 As String = ""
    Private LastGroupHeading4 As String = ""
    Private LastGroupHeading0 As String = ""
    Private AvoidEvents As Boolean = True

    Private WrapCG_CF As CustomGridWrapper
    Private WrapCG_SOCI As CustomGridWrapper
    Private WrapCG_BS As CustomGridWrapper


    Public Sub New(SetModelID As Integer, MyParent As GroupInterfaceTemplate)

        InitializeComponent()

        Formatter = New ObjectFormatter

        ModelID = SetModelID

        Form_InitilisationProcess_SetDataSource()

        'UnC for menus
        'GridLocalizer.Active = New GroupRowContextMenuLocalizerV2()

        Dim DSAnalDataRangeA As RangeDataSource = DSAnalDataRange
        Dim DSAnalDataRangeB As RangeDataSource = DSAnalDataRange
        Dim DSAnalDataRangeC As RangeDataSource = DSAnalDataRange

        Dim FilterString1 As String = "[UseInSOCI] > 0"
        Dim FilterString2 As String = "[UseInCF] > 0"
        Dim FilterString3 As String = "[UseInBS] > 0"

        'Instance a new wrapper for SOCI data

        WrapCG_SOCI = New CustomGridWrapper

        WrapCG_SOCI.WrappedCGC.DataSource = DSAnalDataRangeA

        Dim View_WrapCG_SOCI As CustomGridView = WrapCG_SOCI.WrappedGridView

        WrapCG_SOCI.Dock = DockStyle.Fill

        View_WrapCG_SOCI.ActiveFilterString = FilterString1

        GridView_InitialisationProcess_AddHandlers(View_WrapCG_SOCI)

        ParentGIT = MyParent
        Me.XtraTabPageSOCIWrapped.Controls.Add(WrapCG_SOCI)


        'Instance a new wrapper for CF data

        WrapCG_CF = New CustomGridWrapper

        WrapCG_CF.WrappedCGC.DataSource = DSAnalDataRangeB

        Dim View_WrapCG_CF As CustomGridView = WrapCG_CF.WrappedGridView

        WrapCG_CF.Dock = DockStyle.Fill

        View_WrapCG_CF.ActiveFilterString = FilterString2

        GridView_InitialisationProcess_AddHandlers(View_WrapCG_CF)

        Me.XtraTabPageCFWrapped.Controls.Add(WrapCG_CF)





        'Instance a new wrapper for BS data

        WrapCG_BS = New CustomGridWrapper

        WrapCG_BS.WrappedCGC.DataSource = DSAnalDataRangeC

        Dim View_WrapCG_BS As CustomGridView = WrapCG_BS.WrappedGridView

        WrapCG_BS.Dock = DockStyle.Fill

        View_WrapCG_BS.ActiveFilterString = FilterString3

        GridView_InitialisationProcess_AddHandlers(View_WrapCG_BS)

        Me.XtraTabPageBSWrapped.Controls.Add(WrapCG_BS)


        'Left in for GBP references if needed
        Dim ciGB As CultureInfo = New CultureInfo("en-GB")

        'Add title and summary items to the grids

        Dim item As GridGroupSummaryItem
        Dim item2 As GridGroupSummaryItem
        Dim item3 As GridGroupSummaryItem
        Dim item4 As GridGroupSummaryItem
        Dim item5 As GridGroupSummaryItem
        Dim item6 As GridGroupSummaryItem

        item = New GridGroupSummaryItem
        item.FieldName = "TitleLevel"
        item.SummaryType = DevExpress.Data.SummaryItemType.Max
        item.ShowInGroupColumnFooter = RequireColumn(View_WrapCG_SOCI, "OrderedSOCIHeading")
        View_WrapCG_SOCI.GroupSummary.Add(item)

        item2 = New GridGroupSummaryItem
        item2.FieldName = "TitleLevel"
        item2.SummaryType = DevExpress.Data.SummaryItemType.Max
        item2.ShowInGroupColumnFooter = RequireColumn(View_WrapCG_CF, "OrderedCFHeading")
        View_WrapCG_CF.GroupSummary.Add(item2)

        item3 = New GridGroupSummaryItem
        item3.FieldName = "AmDummy"
        item3.SummaryType = DevExpress.Data.SummaryItemType.Min
        item3.ShowInGroupColumnFooter = RequireColumn(View_WrapCG_SOCI, "Level 2 Copy")
        View_WrapCG_SOCI.GroupSummary.Add(item3)

        item4 = New GridGroupSummaryItem
        item4.FieldName = "AmDummy"
        item4.SummaryType = DevExpress.Data.SummaryItemType.Min
        item4.ShowInGroupColumnFooter = RequireColumn(View_WrapCG_CF, "Level 2 Copy")
        View_WrapCG_CF.GroupSummary.Add(item4)

        item5 = New GridGroupSummaryItem
        item5.FieldName = "TitleLevel"
        item5.SummaryType = DevExpress.Data.SummaryItemType.Max
        item5.ShowInGroupColumnFooter = RequireColumn(View_WrapCG_BS, "OrderedBSHeading")
        View_WrapCG_BS.GroupSummary.Add(item5)

        item6 = New GridGroupSummaryItem
        item6.FieldName = "AmDummy"
        item6.SummaryType = DevExpress.Data.SummaryItemType.Min
        item6.ShowInGroupColumnFooter = RequireColumn(View_WrapCG_BS, "Level 2 Copy")
        View_WrapCG_BS.GroupSummary.Add(item6)

        Dim unboundColumn As New GridColumn() With {
                .Caption = "Item\Description",
                .FieldName = "ItemDesc",
                .UnboundDataType = GetType(String),
                .UnboundExpression = "IIf([Level 3] = '', '', [Level 3] + ' - ') + IIf([Level 4] = '', '', [Level 4] + ' - ') + [Description]",
                .Visible = True
            }
        View_WrapCG_SOCI.Columns.Add(unboundColumn)

        Dim unboundColumn2 As New GridColumn() With {
                .Caption = "Item\Description",
                .FieldName = "ItemDesc",
                .UnboundDataType = GetType(String),
                .UnboundExpression = "IIf([Level 3] = '', '', [Level 3] + ' - ') + IIf([Level 4] = '', '', [Level 4] + ' - ') + [Description]",
                .Visible = True
            }
        View_WrapCG_CF.Columns.Add(unboundColumn2)
        unboundColumn2.VisibleIndex = 12

        Dim unboundColumn3 As New GridColumn() With {
                .Caption = "Item\Description",
                .FieldName = "ItemDesc",
                .UnboundDataType = GetType(String),
                .UnboundExpression = "IIf([Level 3] = '', '', [Level 3] + ' - ') + IIf([Level 4] = '', '', [Level 4] + ' - ') + [Description]",
                .Visible = True
            }
        View_WrapCG_BS.Columns.Add(unboundColumn3)
        unboundColumn3.VisibleIndex = 12

        GridView_InitialisationProcess_AddSummaries(View_WrapCG_SOCI)
        GridView_InitialisationProcess_AddSummaries(View_WrapCG_CF)
        GridView_InitialisationProcess_AddSummariesBS(View_WrapCG_BS)

        GridView_InitialisationProcess_LocalFormat(View_WrapCG_SOCI)
        GridView_InitialisationProcess_LocalFormat(View_WrapCG_CF)
        GridView_InitialisationProcess_LocalFormat(View_WrapCG_BS)

        Dim OpeningBalanceColumn As GridColumn = GetOpeningBalanceColumn(View_WrapCG_BS)
        If OpeningBalanceColumn IsNot Nothing Then OpeningBalanceColumn.Visible = True

        'Set groupings for grids

        With View_WrapCG_SOCI

            RequireColumn(View_WrapCG_SOCI, "OrderedSOCIGroup").Group()
            RequireColumn(View_WrapCG_SOCI, "OrderedSOCIHeading").Group()
            RequireColumn(View_WrapCG_SOCI, "Level 1 Copy").Group()
            RequireColumn(View_WrapCG_SOCI, "Level 2 Copy").Group()
            .CollapseAllGroups()
            .ExpandGroupLevel(0)

        End With

        With View_WrapCG_CF

            RequireColumn(View_WrapCG_CF, "OrderedCFGroup").Group()
            RequireColumn(View_WrapCG_CF, "OrderedCFHeading").Group()
            RequireColumn(View_WrapCG_CF, "Level 1 Copy").Group()
            RequireColumn(View_WrapCG_CF, "Level 2 Copy").Group()
            .CollapseAllGroups()
            .ExpandGroupLevel(0)

        End With

        With View_WrapCG_BS

            RequireColumn(View_WrapCG_BS, "OrderedBSGroup").Group()
            RequireColumn(View_WrapCG_BS, "OrderedBSHeading").Group()
            RequireColumn(View_WrapCG_BS, "Level 1 Copy").Group()
            RequireColumn(View_WrapCG_BS, "Level 2 Copy").Group()
            .CollapseAllGroups()
            .ExpandGroupLevel(0)

        End With

        'Format the grids and set the active grid

        Formatter.FormatGridView(View_WrapCG_SOCI, WrapCG_SOCI.WrappedCGC, "Smaller")
        Formatter.FormatGridView(View_WrapCG_CF, WrapCG_CF.WrappedCGC, "Smaller")
        Formatter.FormatGridView(View_WrapCG_BS, WrapCG_BS.WrappedCGC, "Smaller")

        GridView_Process_SetExpandedLevels(View_WrapCG_SOCI)
        GridView_Process_SetExpandedLevels(View_WrapCG_CF)
        GridView_Process_SetExpandedLevels(View_WrapCG_BS)

        XtraTabControlAnalyser.SelectedTabPage = XtraTabPageSOCIWrapped

        ActiveGridView = View_WrapCG_SOCI

        AvoidEvents = False

        CalcEngID = ExcelModels(ModelID).WBCalcEngine.AddActiveObject(Me)
        Dim ActiveSpreadsheet As DevExpress.Spreadsheet.Worksheet
        ActiveSpreadsheet = ExcelModels(SetModelID).WB.Worksheets("Transactional DB")

        ExcelModels(ModelID).WBCalcEngine.AddActiveWorksheet(CalcEngID, ActiveSpreadsheet)

        Exit Sub

    End Sub

    Public Sub ResizeControlsCommand()



    End Sub

#Region "Form initialisation, non-grid events and data"

    Sub Form_InitilisationProcess_SetDataSource()

        Dim worksheet As DevExpress.Spreadsheet.Worksheet = FileManager.ExcelModels(ModelID).WB.Worksheets("Transactional DB")
        Dim ColList As New List(Of String)
        Dim typelist As String = ""
        Dim CTD As New BPIEAColumnDetectorV2()
        TransDBDataRange = worksheet.Range("Transactional_Records")

        Dim RDSOptions As New RangeDataSourceOptions With {
            .UseFirstRowAsHeader = True,
            .PreserveFormulas = False,
            .SkipHiddenRows = True,
            .SkipHiddenColumns = True,
            .DataSourceColumnTypeDetector = CTD,
            .EditingOptions = DataSourceEditingOptions.ReadOnly
        }

        DSAnalDataRange = TransDBDataRange.GetDataSource(RDSOptions)

    End Sub

    Public Sub DisconnectRDS()

        If DSAnalDataRange IsNot Nothing Then

            AmInactiveState = True

            WrapCG_SOCI.WrappedCGC.DataSource = Nothing
            WrapCG_CF.WrappedCGC.DataSource = Nothing
            WrapCG_BS.WrappedCGC.DataSource = Nothing

            DSAnalDataRange.Dispose()

            DSAnalDataRange = Nothing

            TransDBDataRange = Nothing

        End If

    End Sub

    <Obsolete("Use DisconnectRDS.")>
    Public Sub DisconectRDS()
        DisconnectRDS()
    End Sub

    Public Sub ReconnectRDS()

        If DSAnalDataRange Is Nothing Then

            Form_InitilisationProcess_SetDataSource()
            WrapCG_SOCI.WrappedCGC.DataSource = DSAnalDataRange
            WrapCG_CF.WrappedCGC.DataSource = DSAnalDataRange
            WrapCG_BS.WrappedCGC.DataSource = DSAnalDataRange

            AmInactiveState = False

        End If

    End Sub

    'Sub AddIdentities()
    '    Dim ColList As New List(Of String)
    '    Dim TypeList As New List(Of String)

    '    Dim CD As New BPIEAColumnDetectorV2(ColList, TypeList)
    '    Dim RDSOptions As New RangeDataSourceOptions With {
    '            .UseFirstRowAsHeader = False,
    '            .PreserveFormulas = False,
    '            .SkipHiddenRows = True,
    '            .SkipHiddenColumns = True,
    '            .DataSourceColumnTypeDetector = CD,
    '            .EditingOptions = DataSourceEditingOptions.ReadOnly
    '            }
    'End Sub






    Private Sub WindowsUIButtonPanelBPActions_ButtonClick(sender As Object, e As DevExpress.XtraBars.Docking2010.ButtonEventArgs) Handles WindowsUIButtonPanelAnalyser.ButtonClick
        Dim ButSender As WindowsUIButton = TryCast(e.Button, DevExpress.XtraBars.Docking2010.WindowsUIButton)
        If ButSender Is Nothing OrElse ButSender.Tag Is Nothing Then
            Return
        End If
        Dim tag As String = ButSender.Tag.ToString()
        Select Case tag

            Case "ExportXL"

                Gridview_Process_ExportGridsToExcel()

            Case "OpenHome"

                If FormMainScreen.WindowState = FormWindowState.Minimized Then
                    FormMainScreen.WindowState = FormWindowState.Normal
                End If

                FormMainScreen.BringToFront()

            Case "Snapshot"

            Case "SideBySide"

                RunSideBySide(ModelID, "Assumptions", ParentGIT, True)

            Case "ExpandAll"

                GridView_Process_ExpandAll(ActiveGridView)

            Case "CollapseAll"

                ActiveGridView.CollapseAllGroups()
                ActiveGridView.ExpandGroupLevel(0)
                GridView_Process_SetExpandedLevels(ActiveGridView)

        End Select

    End Sub
    Private Sub XtraTabControlAnalyser_Selected(sender As Object, e As DevExpress.XtraTab.TabPageEventArgs) Handles XtraTabControlAnalyser.Selected

        If AvoidEvents Then Return

        If XtraTabControlAnalyser.SelectedTabPage Is XtraTabPageSOCIWrapped Then

            ActiveGridView = WrapCG_SOCI.WrappedGridView

        ElseIf XtraTabControlAnalyser.SelectedTabPage Is XtraTabPageCFWrapped Then

            ActiveGridView = WrapCG_CF.WrappedGridView

        ElseIf XtraTabControlAnalyser.SelectedTabPage Is XtraTabPageBSWrapped Then

            ActiveGridView = WrapCG_BS.WrappedGridView

        End If

    End Sub

#End Region

#Region "Grid Initialisation"

    Sub GridView_InitialisationProcess_AddHandlers(CGV As CustomGridView)

        AddHandler CGV.CustomDrawCell, AddressOf GridView_CustomDraw_GridCell
        AddHandler CGV.CustomDrawGroupRow, AddressOf GridView_CustomDraw_GroupRow
        AddHandler CGV.RowStyle, AddressOf GridView_CustomStyle_RowStyle
        AddHandler CGV.CustomDrawRowFooter, AddressOf GridView_CustomDraw_RowFooter
        AddHandler CGV.CustomDrawRowFooterCell, AddressOf GridView_CustomDraw_GroupFooterCells
        AddHandler CGV.CustomDrawGroupRowCell, AddressOf GridView_CustomDraw_GroupRowCell
        AddHandler CGV.CustomColumnDisplayText, AddressOf GridView_CustomColumnDisplayText
        AddHandler CGV.Click, AddressOf GridView_Event_SingleClick
        AddHandler CGV.GroupRowExpanding, AddressOf GridView_Event_GroupRowExpanding

    End Sub
    Sub GridView_InitialisationProcess_AddSummaries(CGV As CustomGridView)

        Dim item As GridGroupSummaryItem

        Dim periodColumns As List(Of GridColumn) = GetPeriodColumns(CGV)
        If periodColumns.Count = 0 Then
            Throw New InvalidOperationException(
                "Transactional_Records does not contain any period columns.")
        End If

        For periodIndex As Integer = 0 To periodColumns.Count - 1

            Dim column As GridColumn = periodColumns(periodIndex)

            column.DisplayFormat.FormatType = FormatType.Numeric
            column.DisplayFormat.FormatString = "n0"
            column.Caption = "Year " & CStr(periodIndex + 1) & vbLf &
                             column.CustomizationSearchCaption
            column.Width = 120

            item = New GridGroupSummaryItem
            item.FieldName = column.FieldName
            column.SummaryItem.DisplayFormat = "{0:n0;(n0)}"
            column.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            item.SummaryType = DevExpress.Data.SummaryItemType.Sum
            item.DisplayFormat = "{0:n0}"
            item.ShowInGroupColumnFooter = column

            CGV.GroupSummary.Add(item)

        Next

    End Sub
    Sub GridView_InitialisationProcess_AddSummariesBS(CGV As CustomGridView)

        Dim item As GridGroupSummaryItem

        Dim openingColumn As GridColumn = GetOpeningBalanceColumn(CGV)
        If openingColumn Is Nothing Then
            Throw New InvalidOperationException(
                "Transactional_Records does not contain an opening-balance column.")
        End If

        openingColumn.DisplayFormat.FormatType = FormatType.Numeric
        openingColumn.DisplayFormat.FormatString = "n0"
        openingColumn.Caption = "Opening Balance"
        openingColumn.Width = 120

        item = New GridGroupSummaryItem
        item.FieldName = openingColumn.FieldName
        openingColumn.SummaryItem.DisplayFormat = "{0:n0;(n0)}"
        openingColumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        item.SummaryType = DevExpress.Data.SummaryItemType.Sum
        item.DisplayFormat = "{0:n0}"
        item.ShowInGroupColumnFooter = openingColumn

        CGV.GroupSummary.Add(item)

        Dim periodColumns As List(Of GridColumn) = GetPeriodColumns(CGV)
        For periodIndex As Integer = 0 To periodColumns.Count - 1

            Dim column As GridColumn = periodColumns(periodIndex)
            column.DisplayFormat.FormatType = FormatType.Numeric
            column.DisplayFormat.FormatString = "n0"
            column.Caption = "Year " & CStr(periodIndex + 1) & vbLf &
                             column.CustomizationSearchCaption
            column.Width = 120

            item = New GridGroupSummaryItem
            item.FieldName = column.FieldName
            column.SummaryItem.DisplayFormat = "{0:n0;(n0)}"
            column.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            item.SummaryType = DevExpress.Data.SummaryItemType.Sum
            item.DisplayFormat = "{0:n0}"
            item.ShowInGroupColumnFooter = column

            CGV.GroupSummary.Add(item)

        Next

    End Sub
    Sub GridView_InitialisationProcess_LocalFormat(GV As CustomGridView)

        With GV

            For Each column As GridColumn In .Columns
                column.Visible = False
            Next

            Dim descriptionColumn As GridColumn = RequireColumn(GV, "ItemDesc")
            descriptionColumn.Visible = True
            descriptionColumn.VisibleIndex = 0
            descriptionColumn.BestFit()
            descriptionColumn.Fixed = FixedStyle.Left

            Dim periodColumns As List(Of GridColumn) = GetPeriodColumns(GV)

            For periodIndex As Integer = 0 To periodColumns.Count - 1
                Dim periodColumn As GridColumn = periodColumns(periodIndex)
                periodColumn.Visible = True
                periodColumn.VisibleIndex = periodIndex + 1
            Next

            '.Columns(12).BestFit()
            '.Columns(13).BestFit()
            '.Columns(14).BestFit()
            '.Columns(14).DisplayFormat.FormatType = DevExpress.Utils.FormatType.None

            '.Columns(12).Fixed = FixedStyle.Left
            '.Columns(13).Fixed = FixedStyle.Left
            '.Columns(14).Fixed = FixedStyle.Left

            .OptionsView.ShowGroupedColumns = False
            .OptionsBehavior.AlignGroupSummaryInGroupRow = DevExpress.Utils.DefaultBoolean.True
            .OptionsView.GroupFooterShowMode = GroupFooterShowMode.Hidden
            .OptionsView.EnableAppearanceEvenRow = False
            .OptionsView.EnableAppearanceOddRow = False
            .OptionsView.ShowGroupPanel = False
            .OptionsView.ShowGroupedColumns = False
            .OptionsView.ShowFooter = False
            .OptionsView.ShowColumnHeaders = True
            .OptionsView.ShowIndicator = False
            .OptionsView.ShowAutoFilterRow = False
            .OptionsView.ShowFilterPanelMode = ShowFilterPanelMode.Never
            .OptionsView.ShowViewCaption = False
            .OptionsView.ShowDetailButtons = False
            .OptionsView.ShowPreview = False
            .OptionsView.ShowHorizontalLines = DevExpress.Utils.DefaultBoolean.True
            .OptionsView.ShowVerticalLines = DevExpress.Utils.DefaultBoolean.True
            .OptionsView.ColumnAutoWidth = False
            .OptionsBehavior.AlignGroupSummaryInGroupRow = DefaultBoolean.False
            .OptionsView.GroupFooterShowMode = GroupFooterShowMode.VisibleIfExpanded
            .OptionsSelection.EnableAppearanceFocusedCell = False
            .OptionsSelection.EnableAppearanceHotTrackedRow = False
            .OptionsSelection.EnableAppearanceFocusedRow = False
            .OptionsSelection.MultiSelect = True
            .OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect
            .OptionsClipboard.CopyColumnHeaders = DefaultBoolean.False
            .OptionsClipboard.PasteMode = DevExpress.Export.PasteMode.None
            .OptionsBehavior.Editable = False
            .OptionsBehavior.ReadOnly = True
            .OptionsView.AllowHtmlDrawGroups = True
            .OptionsView.AllowHtmlDrawHeaders = True

            For Each periodColumn As GridColumn In GetPeriodColumns(GV)
                periodColumn.BestFit()
            Next

        End With

    End Sub

#End Region

#Region "Grid Processing and events"

    Sub GridView_Process_SetExpandedLevels(GV As CustomGridView)

        ' Obtain the number of data rows. 
        Dim DataRowCount As Integer = GV.RowCount

        ' Traverse data rows and change the expansion according to summaries. 



        Dim GRC As Integer = -1

        Do While GV.IsValidRowHandle(GRC)

            If GV.IsGroupRow(GRC) Then

                    Dim GRExpanded As Boolean = GV.GetRowExpanded(GRC)

                    Dim childRecordCount As Integer
                    If Not TryGetGroupSummaryInteger(
                        GV, GRC, 0, childRecordCount) Then

                        GRC -= 1
                        Continue Do
                    End If

                    If GV.GetRowLevel(GRC) = 0 Then

                        If childRecordCount > 0 Then

                            If GRExpanded Then GV.SetRowExpanded(GRC, False, True)

                        End If
                    End If

                    Dim dummyCount As Integer
                    If TryGetGroupSummaryInteger(GV, GRC, 1, dummyCount) AndAlso
                       dummyCount = 1 Then

                        If GRExpanded Then GV.SetRowExpanded(GRC, False, True)

                    End If

            End If

            GRC -= 1

        Loop



    End Sub
    Sub GridView_Process_ExpandChildRows(GV As CustomGridView, MasterRowHandle As Integer)

        If Not GV.IsGroupRow(MasterRowHandle) Then Return

        Dim RLev As Integer = GV.GetRowLevel(MasterRowHandle)
        Dim CRC As Integer = GV.GetChildRowCount(MasterRowHandle)
        Dim ChildRowHandle As Integer

        Select Case RLev

            Case 0

                If GridView_Function_CanExpand(GV, MasterRowHandle, 0) Then

                    GV.SetRowExpanded(MasterRowHandle, True, False)

                Else

                    Return

                End If

                For x = 0 To CRC - 1

                    ChildRowHandle = GV.GetChildRowHandle(MasterRowHandle, x)

                    GV.SetRowExpanded(ChildRowHandle, True, False)

                    If GV.IsGroupRow(ChildRowHandle) Then

                        For y = 0 To GV.GetChildRowCount(ChildRowHandle) - 1

                            Dim ChildChildRowHandle As Integer = GV.GetChildRowHandle(ChildRowHandle, y)

                            If GridView_Function_CanExpand(GV, ChildChildRowHandle, 2) Then GV.SetRowExpanded(ChildChildRowHandle, True, True)

                        Next

                    End If

                Next

            Case 1

                GV.SetRowExpanded(MasterRowHandle, True, False)

                For x = 0 To CRC - 1

                    ChildRowHandle = GV.GetChildRowHandle(MasterRowHandle, x)

                    If GridView_Function_CanExpand(GV, ChildRowHandle, 2) Then GV.SetRowExpanded(ChildRowHandle, True, True)

                Next

            Case 2

                If GridView_Function_CanExpand(GV, MasterRowHandle, 2) Then GV.SetRowExpanded(MasterRowHandle, True, True)

            Case Else

                GV.SetRowExpanded(MasterRowHandle, True, True)

        End Select

    End Sub
    Function GridView_Function_CanExpand(GV As CustomGridView, RowHandle As Integer, RLev As Integer) As Boolean

        Dim Expandable As Boolean = False

        Dim summaryValue As Integer

        If RLev = 0 AndAlso
           TryGetGroupSummaryInteger(GV, RowHandle, 0, summaryValue) Then

            Expandable = (summaryValue = 0)

        ElseIf RLev = 2 AndAlso
               TryGetGroupSummaryInteger(GV, RowHandle, 1, summaryValue) Then

            Expandable = (summaryValue = 0)
        End If

        Return Expandable

    End Function
    Sub Gridview_Process_ExportGridsToExcel()

        SetExportMode("XLSAnalysis")

        Dim CFExportPackage As New GridExportPackage With {
            .GridView = WrapCG_CF.WrappedGridView,
            .Description = "Cashflow Data",
            .GroupA = RequireColumn(WrapCG_CF.WrappedGridView, "OrderedCFGroup").AbsoluteIndex,
            .GroupB = RequireColumn(WrapCG_CF.WrappedGridView, "OrderedCFHeading").AbsoluteIndex,
            .GroupC = RequireColumn(WrapCG_CF.WrappedGridView, "Level 1 Copy").AbsoluteIndex,
            .GroupD = RequireColumn(WrapCG_CF.WrappedGridView, "Level 2 Copy").AbsoluteIndex,
            .IDStart = GetPeriodColumns(WrapCG_CF.WrappedGridView)(0).AbsoluteIndex,
            .IDCount = GetPeriodColumns(WrapCG_CF.WrappedGridView).Count - 1}

        Dim SOCIExportPackage As New GridExportPackage With {
            .GridView = WrapCG_SOCI.WrappedGridView,
            .Description = "SOCI Data",
            .GroupA = RequireColumn(WrapCG_SOCI.WrappedGridView, "OrderedSOCIGroup").AbsoluteIndex,
            .GroupB = RequireColumn(WrapCG_SOCI.WrappedGridView, "OrderedSOCIHeading").AbsoluteIndex,
            .GroupC = RequireColumn(WrapCG_SOCI.WrappedGridView, "Level 1 Copy").AbsoluteIndex,
            .GroupD = RequireColumn(WrapCG_SOCI.WrappedGridView, "Level 2 Copy").AbsoluteIndex,
            .IDStart = GetPeriodColumns(WrapCG_SOCI.WrappedGridView)(0).AbsoluteIndex,
            .IDCount = GetPeriodColumns(WrapCG_SOCI.WrappedGridView).Count - 1}

        Dim BSExportPackage As New GridExportPackage With {
            .GridView = WrapCG_BS.WrappedGridView,
            .Description = "Balance Sheet Data",
            .GroupA = RequireColumn(WrapCG_BS.WrappedGridView, "OrderedBSGroup").AbsoluteIndex,
            .GroupB = RequireColumn(WrapCG_BS.WrappedGridView, "OrderedBSHeading").AbsoluteIndex,
            .GroupC = RequireColumn(WrapCG_BS.WrappedGridView, "Level 1 Copy").AbsoluteIndex,
            .GroupD = RequireColumn(WrapCG_BS.WrappedGridView, "Level 2 Copy").AbsoluteIndex,
            .IDStart = GetPeriodColumns(WrapCG_BS.WrappedGridView)(0).AbsoluteIndex,
            .IDCount = GetPeriodColumns(WrapCG_BS.WrappedGridView).Count - 1}

        Exporter.ClearExportPackages()

        If XtraTabControlAnalyser.SelectedTabPage Is XtraTabPageSOCIWrapped Then

            Exporter.AddExportPackage(SOCIExportPackage, True)
            Exporter.AddExportPackage(CFExportPackage, False)
            Exporter.AddExportPackage(BSExportPackage, False)

        ElseIf XtraTabControlAnalyser.SelectedTabPage Is XtraTabPageCFWrapped Then

            Exporter.AddExportPackage(CFExportPackage, True)
            Exporter.AddExportPackage(SOCIExportPackage, False)
            Exporter.AddExportPackage(BSExportPackage, False)

        Else

            Exporter.AddExportPackage(BSExportPackage, True)
            Exporter.AddExportPackage(SOCIExportPackage, False)
            Exporter.AddExportPackage(CFExportPackage, False)

        End If

        Exporter.ProcessAdditions()

        Try

            Exporter.Show()
            Exporter.BringToFront()

        Catch ex As Exception

            XtraMessageBox.Show(
                Me,
                "The analysis export could not be opened." & vbCrLf & ex.Message,
                "Analysis export",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)

        End Try



    End Sub
    Sub GridView_Process_ExpandAll(GV As CustomGridView)

        ' Obtain the number of data rows. 
        Dim DataRowCount As Integer = GV.RowCount

        ' Traverse data rows and change the expansion according to summaries. 


        Dim GRC As Integer = -1

        Do While GV.IsValidRowHandle(GRC)

            If GV.IsGroupRow(GRC) Then

                    If GV.GetRowLevel(GRC) = 0 Then


                        Dim MaxTitleLevel As Integer
                        If Not TryGetGroupSummaryInteger(
                            GV, GRC, 0, MaxTitleLevel) Then

                            GRC -= 1
                            Continue Do
                        End If

                        If MaxTitleLevel = 0 Then GridView_Process_ExpandChildRows(GV, GRC)

                    End If

            End If

            GRC -= 1

        Loop

    End Sub
    Private Sub GridView_Process_CollapseMasterGroups(sender As CustomGridView)

        sender.BeginUpdate()
        Dim rowHandle As Integer = 0

        Try
            Do While rowHandle < sender.DataRowCount

                sender.SetRowCellValue(rowHandle, ColName, String.Empty)
                rowHandle += 1

            Loop

        Finally
            sender.EndUpdate()
        End Try

    End Sub

#End Region

#Region "Grid Draw and Style Handlers"

    Private Sub GridView_CustomStyle_RowStyle(GV As CustomGridView, e As RowStyleEventArgs)

        If AmInactiveState Then Return

        If GV Is Nothing Then Return

        If Not GV.IsGroupRow(e.RowHandle) Then

            e.Appearance.FontStyleDelta = FontStyle.Regular

        Else

            e.Appearance.FontStyleDelta = FontStyle.Bold


            If GV.GroupCount = 0 Then Return

            Dim childRecordCount As Integer
            If Not TryGetGroupSummaryInteger(
                GV, e.RowHandle, 0, childRecordCount) Then Return

            If childRecordCount = 1 Then

                e.Appearance.BackColor = Color.Pink

            ElseIf childRecordCount = 2 Then

                e.Appearance.BackColor = Color.Lavender
                e.Appearance.FontStyleDelta = FontStyle.Bold

            Else

                Select Case GV.GetRowLevel(e.RowHandle)
                    Case 0

                        e.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(235, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(250, Byte), Integer))

                    Case 1

                        e.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(250, Byte), Integer))

                    Case 2

                        e.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(255, Byte), Integer))

                    Case 3

                        e.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(250, Byte), Integer), CType(CType(255, Byte), Integer))

                    Case 4

                        e.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(254, Byte), Integer), CType(CType(254, Byte), Integer), CType(CType(255, Byte), Integer))

                End Select

            End If

        End If

    End Sub
    Private Sub GridView_GroupRowCollapsing(ByVal sender As Object, e As DevExpress.XtraGrid.Views.Base.RowAllowEventArgs)


        'If e.RowHandle Then
        '    e.Allow = False
        'End If

    End Sub
    Private Sub GridView_Event_GroupRowExpanding(ByVal sender As CustomGridView, e As DevExpress.XtraGrid.Views.Base.RowAllowEventArgs)


        Dim RLev As Integer = sender.GetRowLevel(e.RowHandle)

        Dim summaryValue As Integer

        If TryGetGroupSummaryInteger(
            sender, e.RowHandle, 0, summaryValue) Then

            If summaryValue > 0 Then

                e.Allow = False
                Return

            End If

        End If

        If TryGetGroupSummaryInteger(
            sender, e.RowHandle, 1, summaryValue) Then

            If summaryValue > 0 Then

                e.Allow = False
                Return

            End If

        End If

    End Sub
    Private Sub GridView_CustomDraw_GroupRowCell(sender As Object, e As RowGroupRowCellEventArgs)

        e.Appearance.FontStyleDelta = FontStyle.Bold

        Dim FN As String = e.SummaryItem.FieldName

        'If GridViewAnalysis.GetRowLevel(e.RowHandle) = 0 Then

        '    If Microsoft.VisualBasic.Left(FN, 4) = "Year" Then

        '        Dim CurrRef As Integer = CInt(Microsoft.VisualBasic.Right(FN, Len(FN) - 5))

        '        RunningTotals(CurrRef) += CDbl(e.DisplayText)

        '    End If

        'End If

        'Dim Deffont As New Font(e.Appearance.GetFont, FontStyle.Bold)

        Dim strToWrite As String = e.DisplayText
        'Dim BoldDeffont As New Font(e.Appearance.GetFont, FontStyle.Bold)

        'If GridViewAnalysis.GetRowLevel(e.RowHandle) < 1 Then

        '    Dim ht As Hashtable = GridViewAnalysis.GetGroupSummaryValues(e.RowHandle)

        '    Dim groupSummaryItem As GridSummaryItem = GridViewAnalysis.GroupSummary(0)

        '    Dim MasterTitleIndicator As Integer = Convert.ToInt32(ht(groupSummaryItem))

        '    If MasterTitleIndicator = 1 Then

        '        e.Appearance.ForeColor = Color.Linen
        '        'TryCast(e.Cell, GridCellInfo).CellButtonRect = Rectangle.Empty

        '    Else

        '        e.Appearance.Font = DefaultFontLarger

        '    End If

        '    e.Appearance.BackColor = Color.AntiqueWhite
        '    e.Appearance.Options.UseBackColor = True

        'End If

        strToWrite = e.DisplayText

        If IsNumeric(e.DisplayText) Then

            If CDbl(e.DisplayText) < 0 Then

                If Microsoft.VisualBasic.Left(strToWrite, 1) = "-" Then strToWrite = Microsoft.VisualBasic.Right(strToWrite, Len(strToWrite) - 1)

                e.Appearance.ForeColor = Color.Red
                strToWrite = "(" & strToWrite & ")"

            End If

        End If

        e.Appearance.DrawString(e.Cache, strToWrite, e.CaptionBounds)

        e.Handled = True

    End Sub
    Private Sub GridView_CustomDraw_GroupFooterCells(ByVal sender As CustomGridView, ByVal e As FooterCellCustomDrawEventArgs)

        Dim RLev As Integer = sender.GetRowLevel(e.RowHandle)
        Dim penColor As Color
        Dim penWidth As Single
        Select Case RLev


            Case 0


                penColor = Color.Silver
                penWidth = 3

            Case 1


                penColor = Color.LightGray
                penWidth = 2

            Case 2

                penColor = Color.Gainsboro
                penWidth = 2

            Case 3


                penColor = Color.Gainsboro
                penWidth = 1

            Case Else

                penColor = Color.WhiteSmoke
                penWidth = 1



        End Select
        e.Appearance.FontStyleDelta = FontStyle.Bold
        Dim DisplayText As String = e.Info.DisplayText
        If IsNumeric(DisplayText) Then

            If CDbl(DisplayText) < 0 Then

                If Microsoft.VisualBasic.Left(DisplayText, 1) = "-" Then DisplayText = Microsoft.VisualBasic.Right(DisplayText, Len(DisplayText) - 1)

                e.Appearance.ForeColor = Color.Red
                DisplayText = "(" & DisplayText & ")"

            End If

        End If


        Dim FPoint1 As New PointF(e.Bounds.X, e.Bounds.Y)
        Dim FPoint2 As New PointF(e.Bounds.X + e.Bounds.Width, e.Bounds.Y)
        Using pen As New Pen(penColor, penWidth)
            e.Cache.DrawLine(pen, FPoint1, FPoint2)
        End Using
        e.Appearance.DrawString(e.Cache, DisplayText, e.Bounds)

        e.Handled = True

    End Sub
    Sub GridView_CustomDraw_RowFooter(ByVal sender As CustomGridView, ByVal e As RowObjectCustomDrawEventArgs)

        e.Appearance.FontStyleDelta = FontStyle.Bold

        Dim RLev As Integer = sender.GetRowLevel(e.RowHandle)
        Dim Rect As New Rectangle With {
            .X = e.Bounds.X,
            .Height = e.Bounds.Height,
            .Y = e.Bounds.Y,
            .Width = e.Bounds.Width - DefaultGridCellPadding
            }
        Dim Caption As String = ""
        Dim ParentGroupRowHandle As Integer = sender.GetParentRowHandle(e.RowHandle)
        Dim CRC As Integer = sender.GetChildRowCount(ParentGroupRowHandle)

        Caption = sender.GetGroupRowDisplayText(e.RowHandle)
        Caption = Trim(Caption)

        Select Case RLev


            Case 0

                e.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(235, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(250, Byte), Integer))

                If Len(Caption) < 25 Then

                    '  Caption = ""

                Else

                    Caption = Microsoft.VisualBasic.Right(Caption, Len(Caption) - 25)

                End If

                Rect.Width -= 25
                Rect.X += 25


            Case 1

                If Len(Caption) < 25 Then

                    Caption = "     "

                Else

                    Caption = Microsoft.VisualBasic.Right(Caption, Len(Caption) - 26)

                End If

                Rect.X += 25
                Rect.Width -= 25

                e.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(250, Byte), Integer))

            Case 2

                Rect.X += 25
                Rect.Width -= 25

                If Len(Caption) < 14 Then

                    ' Caption = ""

                Else

                    Caption = Microsoft.VisualBasic.Right(Caption, Len(Caption) - 14)

                End If

                e.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(255, Byte), Integer))

            Case 3

                Rect.X += 25
                Rect.Width -= 25

                If Len(Caption) < 14 Then

                    Caption = ""

                Else

                    Caption = Microsoft.VisualBasic.Right(Caption, Len(Caption) - 14)

                End If

                e.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(250, Byte), Integer), CType(CType(255, Byte), Integer))

            Case 4

                Rect.X += 25
                Rect.Width -= 25

                If Len(Caption) < 25 Then

                    Caption = ""

                Else

                    Caption = Microsoft.VisualBasic.Right(Caption, Len(Caption) - 20)

                End If

                e.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(254, Byte), Integer), CType(CType(254, Byte), Integer), CType(CType(255, Byte), Integer))

            Case 5

                Rect.X += 25
                Rect.Width -= 25

                If Len(Caption) < 25 Then

                    Caption = ""

                Else

                    Caption = Microsoft.VisualBasic.Right(Caption, Len(Caption) - 20)

                End If

                e.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(254, Byte), Integer), CType(CType(254, Byte), Integer), CType(CType(255, Byte), Integer))

        End Select

        If Microsoft.VisualBasic.Right(Caption, 7) = "(MIN=0)" Then Caption = Microsoft.VisualBasic.Left(Caption, Len(Caption) - 17)

        Caption = Caption & " Total"

        e.Cache.FillRectangle(e.Cache.GetSolidBrush(e.Appearance.BackColor), e.Bounds)

        Dim FPoint As New PointF(Rect.X, Rect.Y + (DefaultGridCellPadding / 2))
        Dim BlackBrush As Brush = e.Cache.GetSolidBrush(Color.Black)
        e.Cache.DrawString(Caption, e.Appearance.GetFont(), BlackBrush, FPoint)

        e.Handled = True
        LastPaintedFooterColour = e.Appearance.BackColor

    End Sub
    Private Sub GridView_CustomColumnDisplayText(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs)

        'Dim ciGB As CultureInfo = New CultureInfo("en-GB")

        'If IsNumeric(e.Column.FieldName) Then
        '    ' e.DisplayText = String.Format(ciGB, "{0:c0}", (e.Value * 1000))
        '    'e.DisplayText = String.Format(ciGB, "{0:c0}", (e.Value))
        'End If

    End Sub
    Private Sub GridView_CustomDraw_GroupRow(sender As CustomGridView, e As RowObjectCustomDrawEventArgs)

        Dim DontIndentTitle As Boolean = False
        Dim BackColor As Color = Color.White
        Dim RLev As Integer = sender.GetRowLevel(e.RowHandle)
        Dim IsGRExpanded As Boolean = sender.GetRowExpanded(e.RowHandle)
        Dim DontDrawButton As Boolean = False
        Dim CurrentRow As GridGroupRowInfo = TryCast(e.Info, GridGroupRowInfo)

        e.Appearance.FontStyleDelta = FontStyle.Bold

        Dim info As GridGroupRowInfo = TryCast(e.Info, GridGroupRowInfo)
        If info Is Nothing OrElse info.Column Is Nothing Then Return

        Dim view As CustomGridView = TryCast(sender, CustomGridView)

        If IsOrderingField(info.Column.FieldName) Then

            Dim GroupingItem As String = Convert.ToString(view.GetGroupRowValue(e.RowHandle, info.Column))
            If Len(GroupingItem) < 5 Then

                info.GroupText = "UndefinedHeadingOrGroup"

            Else

                info.GroupText = " " & Microsoft.VisualBasic.Right(GroupingItem, Len(GroupingItem) - 5)

            End If

        End If

        Dim ht As Hashtable = sender.GetGroupSummaryValues(e.RowHandle)
        If ht Is Nothing OrElse sender.GroupSummary.Count < 2 Then Return

        If ht IsNot Nothing Then

            If RLev < 1 Then

                Dim MaxLevel As Integer

                If TryGetGroupSummaryInteger(
                    sender, e.RowHandle, 0, MaxLevel) AndAlso
                   MaxLevel = 1 Then

                    BackColor = Color.LightSteelBlue
                    DontDrawButton = True
                    DontIndentTitle = True
                    'info.AppearanceGroupButton.ForeColor = AbovoBlue
                    info.ButtonBounds = Rectangle.Empty

                    'e.Info.Paint.FillRectangle(e.Graphics, e.Appearance.GetBackBrush(e.Cache), e.Bounds)
                    'e.Painter.DrawObject(info)

                ElseIf MaxLevel = 2 Then


                    BackColor = System.Drawing.Color.FromArgb(CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(250, Byte), Integer))
                    DontDrawButton = True
                    DontIndentTitle = True
                    'info.AppearanceGroupButton.ForeColor = AbovoBlue
                    info.ButtonBounds = Rectangle.Empty

                End If

            End If

        End If






        Dim childRecordCount2 As Integer

        If TryGetGroupSummaryInteger(
            sender, e.RowHandle, 1, childRecordCount2) AndAlso
           childRecordCount2 = 1 Then

            DontDrawButton = True

        End If


        If BackColor = Color.White Then

            Select Case RLev

                Case 0

                    BackColor = System.Drawing.Color.FromArgb(CType(CType(235, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(250, Byte), Integer))

                Case 1

                    BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(250, Byte), Integer))

                Case 2

                    BackColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(255, Byte), Integer))

                Case 3

                    BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(250, Byte), Integer), CType(CType(255, Byte), Integer))

                Case 4

                    BackColor = System.Drawing.Color.FromArgb(CType(CType(254, Byte), Integer), CType(CType(254, Byte), Integer), CType(CType(255, Byte), Integer))

                Case 5

                    BackColor = System.Drawing.Color.FromArgb(CType(CType(254, Byte), Integer), CType(CType(254, Byte), Integer), CType(CType(255, Byte), Integer))

            End Select

        End If

        'GroupHeadings(e.RowHandle) = info.GroupText


        If Not IsGRExpanded Then

            Dim items As ArrayList = ExtractSummaryItems(view)
            If items.Count = 0 Then
                Return
            End If
            DrawBackground(e, view, DontDrawButton, BackColor, DontIndentTitle)
            DrawSummaryValues(e, view, items)
            e.Handled = True

        Else

            DrawBackground(e, view, DontDrawButton, BackColor, DontIndentTitle)
            e.Handled = True

        End If

    End Sub
    Private Sub GridView_CustomDraw_GridCell(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs)

        'LastRowWasRT
        If IsOrderingField(e.Column.FieldName) Then

            e.Handled = True
            Return

        End If
        If e.Column.FieldName = "Level 1" Then

            e.Handled = True
            Return

        End If
        If e.Column.FieldName = "Level 2" Then

            e.Handled = True
            Return

        End If
        e.Appearance.FontStyleDelta = FontStyle.Regular
        'If GridViewAnalysis.GetRowLevel(e.RowHandle) < 1 Then

        '    e.Appearance.Font = BoldDeffont
        '    e.Appearance.BackColor = Color.AntiqueWhite

        'End If

        'If GridViewAnalysis.GetRowLevel(e.RowHandle) < 1 Then

        '    Dim ht As Hashtable = GridViewAnalysis.GetGroupSummaryValues(e.RowHandle)

        '    Dim groupSummaryItem As GridSummaryItem = GridViewAnalysis.GroupSummary(0)

        '    Dim childRecordCount As Integer = Convert.ToInt32(ht(groupSummaryItem))

        '    If childRecordCount = 1 Then

        '        e.Appearance.Font = DefaultFontLargest
        '        e.Appearance.ForeColor = Color.Linen
        '        TryCast(e.Cell, GridCellInfo).CellButtonRect = Rectangle.Empty

        '    End If



        'End If



        Dim strToWrite As String = e.DisplayText
        Dim FN As String = e.Column.Caption

        If IsNumeric(e.CellValue) Then

            If e.CellValue < 0 Then

                If Microsoft.VisualBasic.Left(strToWrite, 1) = "-" Then strToWrite = Microsoft.VisualBasic.Right(strToWrite, Len(strToWrite) - 1)

                e.Appearance.ForeColor = Color.Red
                strToWrite = "(" & strToWrite & ")"

            End If

        Else

            If IsOrderingField(e.Column.FieldName) Then

                If Len(strToWrite) > 5 Then

                    strToWrite = Microsoft.VisualBasic.Right(e.DisplayText, Len(strToWrite) - 5)

                End If

            End If
        End If


        e.Appearance.DrawString(e.Cache, strToWrite, e.Bounds)
        e.Handled = True

    End Sub

#End Region

#Region "Functions for Custom Draw Totals"

    Private Function ExtractSummaryItems(ByVal view As CustomGridView) As ArrayList
        Dim items As New ArrayList()
        For Each si As GridSummaryItem In view.GroupSummary
            If TypeOf si Is GridGroupSummaryItem AndAlso si.SummaryType <> DevExpress.Data.SummaryItemType.None Then
                items.Add(si)
            End If
        Next si
        Return items
    End Function
    Private Sub DrawBackground(ByVal e As RowObjectCustomDrawEventArgs, ByVal view As GridView, ByVal DontDrawExpandButton As Boolean, BackColor As Color, DontIndentTitle As Boolean)

        Dim painter As GridGroupRowPainter

        Dim info As GridGroupRowInfo
        painter = TryCast(e.Painter, GridGroupRowPainter)
        info = TryCast(e.Info, GridGroupRowInfo)
        If painter Is Nothing OrElse info Is Nothing Then Return
        Dim buttonRect As Rectangle = info.ButtonBounds
        Dim level As Integer = view.GetRowLevel(e.RowHandle)
        If level < 0 OrElse level >= view.GroupedColumns.Count Then Return
        Dim row As Integer = view.GetDataRowHandleByGroupRowHandle(e.RowHandle)


        Dim Caption As String = view.GetRowCellDisplayText(row, view.GroupedColumns(level))

        If Len(Caption) > 5 Then
            If IsNumeric(Microsoft.VisualBasic.Left(Caption, 2)) Then Caption = Microsoft.VisualBasic.Right(Caption, Len(Caption) - 5)
        End If



        If DontDrawExpandButton Then
            info.ButtonBounds = Rectangle.Empty
            If Not DontIndentTitle Then
                Caption = "    " & Caption
            Else
                Caption = " " & Caption
            End If
        End If

        info.GroupText = Caption

        e.Appearance.BackColor = BackColor

        e.Appearance.DrawBackground(e.Cache, info.Bounds)
        painter.ElementsPainter.GroupRow.DrawObject(info)

    End Sub
    Private Sub DrawSummaryValues(ByVal e As RowObjectCustomDrawEventArgs, ByVal view As CustomGridView, ByVal items As ArrayList)

        Dim values As Hashtable = view.GetGroupSummaryValues(e.RowHandle)
        If values Is Nothing Then Return

        Dim RedBrush As Brush = e.Cache.GetSolidBrush(Color.Red)
        Dim BlackBrush As Brush = e.Cache.GetSolidBrush(Color.Black)

        For Each item As GridGroupSummaryItem In items

            Dim rect As Rectangle = GetColumnBounds(view, item)

            If rect.IsEmpty Then

                Continue For

            End If

            Dim text As String = item.GetDisplayText(values(item), False)

            If IsOrderingField(item.FieldName) Then

                text = Microsoft.VisualBasic.Right(text, Len(text) - 5)

            End If

            If IsNumeric(text) Then

                If CDbl(text) < 0 Then

                    If Microsoft.VisualBasic.Left(text, 1) = "-" Then text = Microsoft.VisualBasic.Right(text, Len(text) - 1)

                    text = "(" & text & ")"

                    rect = CalcRowSummaryRect(text, e, view.Columns(item.FieldName))

                    e.Appearance.DrawString(e.Cache, text, rect, RedBrush)


                Else

                    rect = CalcRowSummaryRect(text, e, view.Columns(item.FieldName))

                    e.Appearance.DrawString(e.Cache, text, rect, BlackBrush)

                End If
            Else

                rect = CalcRowSummaryRect(text, e, view.Columns(item.FieldName))

                e.Appearance.DrawString(e.Cache, text, rect, BlackBrush)

            End If

        Next item
    End Sub
    Private Function CalcRowSummaryRect(ByVal text As String, ByVal e As RowObjectCustomDrawEventArgs, ByVal column As GridColumn) As Rectangle

        Dim result As Rectangle = GetColumnBounds(column)
        Dim sz As SizeF = TextUtils.GetStringSize(e.Graphics, text, e.Appearance.Font)
        Dim width As Integer = Convert.ToInt32(sz.Width) + 1
        If (Not gridInfo.ViewRects.FixedLeft.IsEmpty) Then
            Dim fixedLeftRight As Integer = gridInfo.ViewRects.FixedLeft.Right
            Dim marginLeft As Integer = result.Right - width - fixedLeftRight
            If marginLeft < 0 AndAlso column.Fixed = FixedStyle.None Then
                Return Rectangle.Empty
            End If
        End If
        If (Not gridInfo.ViewRects.FixedRight.IsEmpty) Then
            Dim fixedRightLeft As Integer = gridInfo.ViewRects.FixedRight.Left
            If fixedRightLeft <= result.Right AndAlso column.Fixed = FixedStyle.None Then
                Return Rectangle.Empty
            End If
        End If
        result = FixLeftEdge(width, result)
        result.Width = result.Width
        result.Y = e.Bounds.Y
        result.Height = e.Bounds.Height - 2

        Return PreventSummaryTextOverlapping(e, result)

    End Function
    Private Function GetColumnBounds(ByVal view As GridView, ByVal item As GridGroupSummaryItem) As Rectangle
        Dim column As GridColumn = view.Columns(item.FieldName)
        Return GetColumnBounds(column)
    End Function
    Private Function GetColumnBounds(ByVal column As GridColumn) As Rectangle
        gridInfo = CType(column.View.GetViewInfo(), GridViewInfo)
        Dim colInfo As GridColumnInfoArgs = gridInfo.ColumnsInfo(column)

        If colInfo IsNot Nothing Then
            Return colInfo.Bounds
        Else
            Return Rectangle.Empty
        End If
    End Function
    Private Function FixLeftEdge(ByVal width As Integer, ByVal result As Rectangle) As Rectangle
        Dim delta As Integer = result.Width - width - 2
        If delta > 0 Then
            result.X += delta
            result.Width -= delta
        End If
        Return result
    End Function
    Private Function PreventSummaryTextOverlapping(ByVal e As RowObjectCustomDrawEventArgs, ByVal rect As Rectangle) As Rectangle
        Dim gInfo As GridGroupRowInfo = CType(e.Info, GridGroupRowInfo)
        Dim groupTextLocation As Integer = gInfo.ButtonBounds.Right + 10
        Dim groupTextWidth As Integer = TextUtils.GetStringSize(e.Graphics, gInfo.GroupText, e.Appearance.Font).Width
        Dim fixedLeft As Integer = gInfo.ViewInfo.ViewRects.FixedLeft.Left
        Dim r As New Rectangle(groupTextLocation, 0, groupTextWidth, e.Info.Bounds.Height)
        If r.Right > rect.X Then
            If r.Right > rect.Right Then
                rect.Width = 0
            Else
                rect.Width -= r.Right - rect.X
                rect.X = r.Right
            End If
        End If
        Return rect
    End Function

#End Region

#Region "Event Handlers"

    Private Sub GridView_Event_SingleClick(ByVal CGVSender As CustomGridView, ByVal e As MouseEventArgs)

        Dim hitInfo As GridHitInfo = CGVSender.CalcHitInfo(e.Location)

        If hitInfo.HitTest = GridHitTest.Row AndAlso CGVSender.IsGroupRow(hitInfo.RowHandle) Then

            If e.Button = MouseButtons.Left Then

                CGVSender.SetRowExpanded(hitInfo.RowHandle, Not CGVSender.GetRowExpanded(hitInfo.RowHandle), False)

            ElseIf e.Button = MouseButtons.Right Then

                If Not CGVSender.GetRowExpanded(hitInfo.RowHandle) Then

                    GridView_Process_ExpandChildRows(CGVSender, hitInfo.RowHandle)

                Else

                    CGVSender.SetRowExpanded(hitInfo.RowHandle, False, True)

                End If

            End If

        End If





    End Sub
    Private Sub GridView_Event_DoubleClick(ByVal sender As Object, ByVal e As MouseEventArgs)

        Dim hitInfo As GridHitInfo = sender.CalcHitInfo(e.Location)
        If hitInfo.HitTest = GridHitTest.Row AndAlso sender.IsGroupRow(hitInfo.RowHandle) Then

            sender.SetRowExpanded(hitInfo.RowHandle, Not sender.GetRowExpanded(hitInfo.RowHandle), True)

        End If
    End Sub
    Sub GridView_Event_PopupMenuShowing(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.PopupMenuShowingEventArgs)

        'Dim hitInfo As GridHitInfo = e.HitInfo

        'If TypeOf e.Menu Is DevExpress.XtraGrid.Menu.GridViewGroupRowMenu Then


        '    Dim Item0 As DevExpress.Utils.Menu.DXMenuItem()
        '    Dim Item1 As New DevExpress.Utils.Menu.DXMenuItem()
        '    Dim Item2 As New DevExpress.Utils.Menu.DXMenuItem()
        '    Dim DontA As Boolean = False
        '    Dim DontB As Boolean = False

        '    For x = 0 To e.Menu.Items.Count
        '        Try
        '            e.Menu.Remove(e.Menu.Items(x).Tag)
        '        Catch ex As Exception

        '        End Try

        '    Next





        '    Item1.Caption = "Expand this row and children"


        '    AddHandler Item1.Click, Sub(ss, ee)
        '                                GridViewAnalysis.SetRowExpanded(hitInfo.RowHandle, True, True)
        '                            End Sub
        '    e.Menu.Items.Add(Item1)





        '    Item2.Caption = "Collapse this row and children"


        '    AddHandler Item2.Click, Sub(ss, ee)
        '                                GridViewAnalysis.SetRowExpanded(hitInfo.RowHandle, False, True)
        '                            End Sub
        '    e.Menu.Items.Add(Item2)




        '    PopMenuIntialised = True



        'End If

    End Sub
    Sub GridView_OverrideEvent_ShowGroupFooter(ByVal sender As Object, ByVal e As ShowGroupFooterEventArgs)

        'Handles myGridView1.ShowGroupFooter

        If e.FooterLevel = 1 Then

            e.Visible = False

        End If

    End Sub
    Private Sub GridView_CalcRowHeight(ByVal sender As Object, ByVal e As RowHeightEventArgs)

        If Not sender.IsGroupRow(e.RowHandle) Then Return

        If sender.GetRowLevel(e.RowHandle) = 0 Then

            e.RowHeight += 20

        ElseIf sender.GetRowLevel(e.RowHandle) = 1 Then

            e.RowHeight += 10

        End If

    End Sub

#End Region

#Region "Calculation Handlers"

    Sub CalcUnboundTotals(Sender As CustomGridView, e As CustomColumnDataEventArgs)

        'Dim ThisField As String = e.Column.FieldName

        'If Microsoft.VisualBasic.Right(ThisField, 3) <> "MRT" Then Return

        'Dim FieldSource As String = Microsoft.VisualBasic.Left(ThisField, Len(ThisField) - 4)

        'Dim rowIndex = e.ListSourceRowIndex

        'Dim FunctionalAmount As Decimal = Convert.ToDecimal(GridViewAnalysis.GetListSourceRowCellValue(rowIndex, FieldSource))

        'If e.IsGetData Then
        '    If rowIndex = 0 Then
        '        e.Value = FunctionalAmount
        '    Else
        '        Dim PreviousRowBal As Decimal = Convert.ToDecimal(GridViewAnalysis.GetListSourceRowCellValue(rowIndex - 1, ThisField))
        '        e.Value = PreviousRowBal + FunctionalAmount
        '    End If
        'End If

    End Sub
    Sub CalcCustomSummary(Sender As CustomGridView, e As CustomSummaryEventArgs)

        Dim GSItem As GridSummaryItem = TryCast(e.Item, GridSummaryItem)

        If GSItem Is Nothing Then Return

        Dim FN As String = GSItem.FieldName
        Dim GroupLevel As Integer = Sender.GetRowLevel(e.RowHandle)
        Dim RH As Integer = e.RowHandle
        Dim view As CustomGridView = TryCast(Sender, CustomGridView)
        Dim IsMasterGT As String = view.GetRowCellValue(e.RowHandle, "IsMasterTotal").ToString
        Dim CurrRef As Integer = CInt(Microsoft.VisualBasic.Right(FN, Len(FN) - 5))

        Select Case e.SummaryProcess

            Case CustomSummaryProcess.Start

                'If FN = "Year 1" And GroupLevel = 5 And RH = 0 Then

                '    RunningTotalStore.Clear()
                '    ReDim RunningGrandTotals(40)

                'End If

                'GridTotals(CurrRef) = 0

            Case CustomSummaryProcess.Calculate

                'GridTotals(CurrRef) += CDbl(e.FieldValue)
                'If GroupLevel = 1 Then RunningGrandTotals(CurrRef) += CDbl(e.FieldValue)

            Case CustomSummaryProcess.Finalize

                If IsMasterGT = "1" Then

                    e.TotalValue = view.GetRowCellValue(view.GetChildRowHandle(e.GroupRowHandle - 1, view.GetChildRowCount(e.GroupRowHandle - 1) - 1), FN & " MRT")

                    'If GroupLevel = 1 Then

                    '    RunningTotalStore.AddRowTotal(RH, CurrRef, RunningGrandTotals(CurrRef))
                    '    e.TotalValue = RunningGrandTotals(CurrRef)

                    'Else

                    '    e.TotalValue = RunningTotalStore.GetRowTotal(RH, CurrRef)

                    'End If

                Else


                    ' e.TotalValue = GridTotals(CurrRef)


                End If

                ' LastProcessedGroupLevel = GroupLevel

        End Select

    End Sub

#End Region

#Region "Sizing Handlers"

    Sub GridView_Sizing_Fonts()

        Scalefactor = Me.Width / 1700


    End Sub
    Sub GridView_Sizing_ResizeControls()

        Dim SetWidth As Integer = Me.Width * 0.17
        ScaleUnits = Me.Width * 0.007

        'PictureBoxAbovoLogo.Top = ScaleUnits
        'PictureBoxAbovoLogo.Left = ScaleUnits
        'PictureBoxAbovoLogo.Width = SetWidth
        'PictureBoxAbovoLogo.Height = CInt(PictureBoxAbovoLogo.Width * 0.483)

        'DockPanelSettings.Width = SetWidth
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

    Private Shared Function RequireColumn(ByVal view As CustomGridView,
                                          ByVal fieldName As String) As GridColumn

        Dim column As GridColumn = view.Columns(fieldName)
        If column Is Nothing Then
            Throw New InvalidOperationException(
                "Transactional_Records is missing required column '" &
                fieldName & "'.")
        End If

        Return column

    End Function

    Private Shared Function TryGetGroupSummaryInteger(
        ByVal view As CustomGridView,
        ByVal rowHandle As Integer,
        ByVal summaryIndex As Integer,
        ByRef result As Integer) As Boolean

        result = 0

        If view Is Nothing OrElse
           summaryIndex < 0 OrElse
           summaryIndex >= view.GroupSummary.Count Then Return False

        Dim summaryValues As Hashtable = view.GetGroupSummaryValues(rowHandle)
        If summaryValues Is Nothing Then Return False

        Dim summaryItem As GridSummaryItem = view.GroupSummary(summaryIndex)
        Return TryConvertSummaryInteger(summaryValues(summaryItem), result)

    End Function

    Private Shared Function TryConvertSummaryInteger(
        ByVal rawValue As Object,
        ByRef result As Integer) As Boolean

        result = 0
        If rawValue Is Nothing OrElse rawValue Is DBNull.Value Then Return False

        Dim numericValue As Double

        If TypeOf rawValue Is String Then
            Dim textValue As String = DirectCast(rawValue, String).Trim()
            If textValue.Length = 0 Then Return False

            Dim isParenthesisedNegative As Boolean =
                textValue.Length > 2 AndAlso
                textValue.StartsWith("(", StringComparison.Ordinal) AndAlso
                textValue.EndsWith(")", StringComparison.Ordinal)

            If isParenthesisedNegative Then
                textValue = textValue.Substring(1, textValue.Length - 2).Trim()
            End If

            If Not Double.TryParse(
                textValue,
                NumberStyles.Any,
                CultureInfo.CurrentCulture,
                numericValue) AndAlso
               Not Double.TryParse(
                   textValue,
                   NumberStyles.Any,
                   CultureInfo.InvariantCulture,
                   numericValue) Then Return False

            If isParenthesisedNegative Then numericValue = -numericValue
        Else
            Try
                numericValue = Convert.ToDouble(rawValue, CultureInfo.InvariantCulture)
            Catch ex As Exception When TypeOf ex Is FormatException OrElse
                                       TypeOf ex Is InvalidCastException OrElse
                                       TypeOf ex Is OverflowException
                Return False
            End Try
        End If

        If Double.IsNaN(numericValue) OrElse
           Double.IsInfinity(numericValue) OrElse
           numericValue < Integer.MinValue OrElse
           numericValue > Integer.MaxValue Then Return False

        result = CInt(Math.Truncate(numericValue))
        Return True

    End Function

    Private Function GetPeriodColumns(ByVal view As CustomGridView) As List(Of GridColumn)

        Dim result As New List(Of GridColumn)

        For Each column As GridColumn In view.Columns
            If IsPeriodHeading(column.FieldName) OrElse
               IsPeriodHeading(column.CustomizationSearchCaption) OrElse
               IsPeriodHeading(column.Caption) Then

                result.Add(column)
            End If
        Next

        If result.Count = 0 AndAlso TransDBDataRange IsNot Nothing Then
            Dim sourceColumnCount As Integer =
                Math.Min(TransDBDataRange.ColumnCount, view.Columns.Count)

            For sourceColumnIndex As Integer = 0 To sourceColumnCount - 1
                If IsPeriodHeading(
                    TransDBDataRange(0, sourceColumnIndex).DisplayText) Then

                    result.Add(view.Columns(sourceColumnIndex))
                End If
            Next
        End If

        result.Sort(
            Function(leftColumn As GridColumn, rightColumn As GridColumn) _
                leftColumn.AbsoluteIndex.CompareTo(rightColumn.AbsoluteIndex))

        Return result

    End Function

    Private Shared Function IsPeriodHeading(ByVal candidate As String) As Boolean

        If String.IsNullOrWhiteSpace(candidate) Then Return False

        Dim lines As String() =
            candidate.Replace(vbCrLf, vbLf).
                      Replace(vbCr, vbLf).
                      Split(New String() {vbLf}, StringSplitOptions.RemoveEmptyEntries)

        For Each line As String In lines
            If Regex.IsMatch(
                line.Trim(),
                "^[0-9]{4}/[0-9]{2}$",
                RegexOptions.CultureInvariant) Then

                Return True
            End If
        Next

        Return False

    End Function

    Private Function GetOpeningBalanceColumn(ByVal view As CustomGridView) As GridColumn

        Dim periodColumns As List(Of GridColumn) = GetPeriodColumns(view)
        If periodColumns.Count = 0 Then Return Nothing

        Dim openingIndex As Integer = periodColumns(0).AbsoluteIndex - 1
        If openingIndex < 0 OrElse openingIndex >= view.Columns.Count Then Return Nothing

        Return view.Columns(openingIndex)

    End Function

    Private Shared Function IsOrderingField(ByVal fieldName As String) As Boolean

        Return String.Equals(fieldName, "OrderedSOCIGroup", StringComparison.Ordinal) OrElse
               String.Equals(fieldName, "OrderedSOCIHeading", StringComparison.Ordinal) OrElse
               String.Equals(fieldName, "OrderedCFGroup", StringComparison.Ordinal) OrElse
               String.Equals(fieldName, "OrderedCFHeading", StringComparison.Ordinal) OrElse
               String.Equals(fieldName, "OrderedBSGroup", StringComparison.Ordinal) OrElse
               String.Equals(fieldName, "OrderedBSHeading", StringComparison.Ordinal)

    End Function

    Friend Sub ReleaseAnalyserResources()

        DisconnectRDS()

        If ExcelModels Is Nothing OrElse
           ModelID < 0 OrElse ModelID >= ExcelModels.Length OrElse
           ExcelModels(ModelID) Is Nothing Then Return

        ExcelModels(ModelID).ResourceRegistry.Release(
            ModelResourceKeys.TransactionalRecordsRangeDataSource,
            Me)

        If CalcEngID >= 0 AndAlso ExcelModels(ModelID).WBCalcEngine IsNot Nothing Then
            ExcelModels(ModelID).WBCalcEngine.RemoveActiveObject(CalcEngID)
            CalcEngID = -1
        End If

        If Object.ReferenceEquals(ExcelModels(ModelID).ExpendAnalyserV2, Me) Then
            ExcelModels(ModelID).ExpendAnalyserV2 = Nothing
        End If

    End Sub


#End Region

End Class

#Region "Helper Classes"

Class GroupRowContextMenuLocalizerV2

    Inherits GridLocalizer

    Public Overrides ReadOnly Property Language() As String
        Get
            Return "English"
        End Get
    End Property

    Public Overrides Function GetLocalizedString(ByVal id As GridStringId) As String
        Select Case id
            Case GridStringId.MenuGroupRowCollapse
                Return "Collapse this row"
            Case GridStringId.MenuGroupRowExpand
                Return "Expand this row"
            Case GridStringId.MenuGroupPanelFullExpand
                Return "Expand all rows and children"
            Case GridStringId.MenuGroupPanelFullCollapse
                Return "Collapse all rows and children"
            Case Else
                Return MyBase.GetLocalizedString(id)
        End Select

    End Function

End Class
Class ArrayTotalStorageContainerV2

        Public RowCustomTotals(-1) As RowTotalArray
        Public IndexCount As Integer = -1

        Public Sub AddRowTotal(RowHandle As Integer, TotalIndex As Integer, NewTotal As Double)

            'Deprecated

            Dim UseRecordIndex As Integer = 0
            Dim Found As Boolean = False

            If RowCustomTotals.Count > 0 Then

                Dim i As Integer

                For i = 0 To RowCustomTotals.Count - 1

                    If RowCustomTotals(i).RowHandle = RowHandle Then
                        UseRecordIndex = i
                        Found = True
                    End If

                Next i

            End If

            If Not Found Then

                IndexCount += 1
                ReDim Preserve RowCustomTotals(IndexCount)
                RowCustomTotals(IndexCount) = New RowTotalArray
                UseRecordIndex = IndexCount

            End If

            RowCustomTotals(UseRecordIndex).RowHandle = RowHandle
            RowCustomTotals(UseRecordIndex).RowTotals(TotalIndex) = NewTotal

        End Sub

        Public Function GetRowTotal(RowHandle As Integer, TotalIndex As Integer) As Double

            'Deprecated
            Dim Found As Boolean = False

            Dim i As Integer

            For i = 0 To RowCustomTotals.Count - 1

                If RowCustomTotals(i).RowHandle = RowHandle Then

                    Found = True
                    Exit For

                End If

            Next i


            If Not Found Then
                Return -1
            Else
                Return RowCustomTotals(i).RowTotals(TotalIndex)
            End If

        End Function
        Public Sub Clear()

            IndexCount = -1
            ReDim RowCustomTotals(-1)

        End Sub

        Class RowTotalArray

            Public RowHandle As Integer
            Public RowTotals(40) As Double

        End Class


    End Class
Class BPIEAColumnDetectorV2
    Implements IDataSourceColumnTypeDetector

    Private ColList As List(Of String)
    Private TyList As List(Of String)
    Public Sub New()



    End Sub
    Public Function GetColumnName(ByVal index As Integer, ByVal offset As Integer, ByVal range As DevExpress.Spreadsheet.CellRange) As String Implements IDataSourceColumnTypeDetector.GetColumnName

        Return range(-1, offset).DisplayText

    End Function

    Public Function GetColumnType(ByVal index As Integer, ByVal offset As Integer, ByVal range As DevExpress.Spreadsheet.CellRange) As Type Implements IDataSourceColumnTypeDetector.GetColumnType

        Dim sawNumeric As Boolean = False
        Dim sawText As Boolean = False
        Dim sampleCount As Integer = Math.Min(range.RowCount, 250)

        For rowOffset As Integer = 0 To sampleCount - 1
            Dim cell As DevExpress.Spreadsheet.Cell = range(rowOffset, offset)
            If cell.Value.IsEmpty Then Continue For

            If cell.Value.IsNumeric Then
                sawNumeric = True
            Else
                sawText = True
            End If

            If sawNumeric AndAlso sawText Then Return GetType(String)
        Next

        If sawNumeric Then Return GetType(Double)
        Return GetType(String)


    End Function

End Class
#End Region
