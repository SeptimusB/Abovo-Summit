

Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports DevExpress.XtraCharts
Imports Abovo
Imports Abovo.GeneralFunctions
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.LogDebugDev

Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.Utils.Drawing
Imports DevExpress.Data
Imports System.Drawing.Drawing2D
Imports Abovo.DataObject
Imports DevExpress.XtraBars.Navigation
Imports DevExpress.XtraPrinting
Imports Abovo.PresentationManager
Imports Abovo.AbovoUnboundSource
Imports DevExpress.XtraLayout.Customization
Imports DevExpress.DataAccess.DataFederation
Imports DevExpress.UnitConversion
Imports DevExpress.Charts.Model
Imports DevExpress.Drawing
Imports DevExpress.Spreadsheet.Charts
Imports DevExpress.Snap.Core.Native
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils
Imports DevExpress.XtraSpreadsheet.Model
Imports System.Globalization
Imports DevExpress.XtraExport.Helpers
Imports DevExpress.XtraGrid.Views
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports Abovo.CustomGrid
Imports DevExpress.CodeParser



Public Class TransactionAnalyser

    Public DataPM As PresentationManager
    Public DataPres As PresentationManager.DataPresentation
    Public ModelID As Integer
    Public PresID As Integer
    Public GSID As Integer
    Public CSID As Integer
    Public rs As New Resizer
    Private BIsDirty As Boolean
    Private Grid1ExpandedView As Boolean
    Public ScaleUnits As Single
    Public Scalefactor As Single
    Public DataSources() As AbovoUnboundSource
    Public DataSourceCount As Integer
    Public PresentedDS As Abovo.DataObject.DataCellRange
    Public PresentedColumn As Abovo.DataObject.SheetDataColumn
    Public PropertyArray() As UnboundSourceProperty
    Public PropertiesCount As Integer
    Public PropertyList As IEnumerable(Of UnboundSourceProperty)
    Public PropType As System.Type
    Public ColList As List(Of String)
    Public ColCount As Integer = 0
    Public ColName As String
    Public GridControls() As GridControl
    Public GridCount As Integer = -1

    Public AcControl As AccordionControl
    Public AcElements() As AccordionControlElement
    Public AcElementlist As List(Of AccordionControlElement)
    Public AcControlCount As Integer = -1
    Public AcElementCount As Integer = -1
    Public AcContainers() As AccordionContentContainer
    Public hyperlinkLabelControls() As HyperlinkLabelControl
    Public AcContainersCount As Integer = -1
    Public UsedGridViews(-1) As GridView
    Public GridViewCount As Integer = -1
    Public Formatter As ObjectFormatter
    Public CurrChartWS As DevExpress.Spreadsheet.Worksheet
    Public DSAnalDataRange As RangeDataSource
    Public HasSnapshots As Boolean = False

    Public Sub New(SetModelID As Integer)


        ' This call is required by the designer.
        InitializeComponent()
        Formatter = New ObjectFormatter

        ModelID = SetModelID

        Dim RunningTotal1 As Integer = 0
        Dim UseRunningTotal1 As Boolean = False
        Dim RunningTotal2 As Integer = 0
        Dim UseRunningTotal2 As Boolean = False


        SetDataSource()

        Dim i As Integer



        GridControlSOCIData.DataSource = DSAnalDataRange

        'Remove non-SOCI data
        Dim filterString As String = "Len([SOCIHeading]) > 5"
        AddHandler GridViewAnalysis.CustomDrawCell, AddressOf GridView_CustomDrawCell

        GridViewAnalysis.Columns(0).FilterInfo = New ColumnFilterInfo(filterString)
        GridViewAnalysis.OptionsView.ShowGroupedColumns = False

        Dim strTemp As String
        Dim ciGB As CultureInfo = New CultureInfo("en-GB")


        For i = 25 To 64

            'If IsNumeric(e.Column.FieldName) Then
            '    ' e.DisplayText = String.Format(ciGB, "{0:c0}", (e.Value * 1000))
            '    e.DisplayText = String.Format(ciGB, "{0:c0}", (e.Value))
            'End If

            GridViewAnalysis.Columns(i).DisplayFormat.FormatType = FormatType.Numeric
            GridViewAnalysis.Columns(i).DisplayFormat.FormatString = "n0"
            strTemp = "Year " & CStr(i - 24)
            GridViewAnalysis.Columns(i).Caption = strTemp
            GridViewAnalysis.Columns(i).Width = 120
            'GridViewAnalysis.Columns(i).DisplayText = CInt(Math.Truncate(GridViewAnalysis.Columns(i).Value)).ToString()

        Next i

        GridViewAnalysis.OptionsBehavior.AlignGroupSummaryInGroupRow = DevExpress.Utils.DefaultBoolean.True
        GridViewAnalysis.OptionsView.GroupFooterShowMode = GroupFooterShowMode.VisibleIfExpanded

        'Dim itemCust As New GridGroupSummaryItem
        'itemCust.FieldName = "Description"
        'itemCust.SummaryType = DevExpress.Data.SummaryItemType.Custom
        'itemCust.DisplayFormat = "Total:"
        'itemCust.ShowInGroupColumnFooter = GridViewAnalysis.Columns(24)
        'GridViewAnalysis.GroupSummary.Add(itemCust)

        Dim item As GridGroupSummaryItem

        For i = 25 To 64

            item = New GridGroupSummaryItem
            item.FieldName = GridViewAnalysis.Columns(i).FieldName
            GridViewAnalysis.Columns(i).SummaryItem.DisplayFormat = "{0:n0;(n0)}"
            GridViewAnalysis.Columns(i).AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            item.SummaryType = DevExpress.Data.SummaryItemType.Sum
            item.DisplayFormat = "{0:n0}"
            item.ShowInGroupColumnFooter = GridViewAnalysis.Columns(i)

            GridViewAnalysis.GroupSummary.Add(item)

        Next

        GridViewAnalysis.Columns(2).Visible = False

        GridViewAnalysis.Columns(3).Visible = False
        GridViewAnalysis.Columns(4).Visible = False
        GridViewAnalysis.Columns(5).Visible = False
        GridViewAnalysis.Columns(6).Visible = False
        GridViewAnalysis.Columns(7).Visible = False
        GridViewAnalysis.Columns(8).Visible = False
        GridViewAnalysis.Columns(9).Visible = False

        GridViewAnalysis.Columns(10).Visible = False
        GridViewAnalysis.Columns(11).Visible = False
        GridViewAnalysis.Columns(12).Visible = False
        GridViewAnalysis.Columns(13).Visible = False
        GridViewAnalysis.Columns(14).Visible = False
        GridViewAnalysis.Columns(15).Visible = False
        GridViewAnalysis.Columns(16).Visible = False
        GridViewAnalysis.Columns(17).Visible = False
        GridViewAnalysis.Columns(18).Visible = False
        GridViewAnalysis.Columns(19).Visible = False
        GridViewAnalysis.Columns(65).Visible = False
        GridViewAnalysis.Columns(65).Visible = False

        GridViewAnalysis.Columns(0).Group()
        GridViewAnalysis.Columns(1).Group()
        GridViewAnalysis.Columns(19).Group()
        GridViewAnalysis.Columns(20).Group()
        GridViewAnalysis.Columns(21).Group()

        GridViewAnalysis.ExpandAllGroups()
        GridViewAnalysis.CollapseGroupLevel(1)

#Region "Old definitions"


        'GridViewAnalysis.Columns(1).Group()
        'GridViewAnalysis.Columns(2).Group()
        'GridViewAnalysis.Columns(3).Group()
        'GridViewAnalysis.Columns(4).Group()
        'GridViewAnalysis.Columns(5).Group()
        'GridViewAnalysis.Columns(6).Group()
        'GridViewAnalysis.Columns(7).Group()
        'GridViewAnalysis.Columns(8).Group()
        'GridViewAnalysis.Columns(9).Group()
        'GridViewAnalysis.Columns(16).Group()
        'GridViewAnalysis.Columns(20).Group()
        'GridViewAnalysis.Columns(21).Group()
        'GridViewAnalysis.Columns(22).Group()
        'GridViewAnalysis.Columns(23).Group()

        'GridViewAnalysis.BestFitColumns()
        'GridViewAnalysis.ExpandAllGroups()

        '    GridViewAnalysis.OptionsView.GroupFooterShowMode = GroupFooterShowMode.VisibleAlways
        '    Dim sortInfo As GridColumnSortInfo() = {
        '        New GridColumnSortInfo(GridViewAnalysis.Columns(0), ColumnSortOrder.Ascending),
        'New GridMergedColumnSortInfo({GridViewAnalysis.Columns(1), GridViewAnalysis.Columns(2), GridViewAnalysis.Columns(3), GridViewAnalysis.Columns(4), GridViewAnalysis.Columns(5), GridViewAnalysis.Columns(6), GridViewAnalysis.Columns(7)}, {ColumnSortOrder.None, ColumnSortOrder.None, ColumnSortOrder.None, ColumnSortOrder.None, ColumnSortOrder.None, ColumnSortOrder.None, ColumnSortOrder.None})
        '}

        'GridViewAnalysis.SortInfo.ClearAndAddRange(sortInfo, 2)

        'GridViewAnalysis.OptionsView.ShowGroupPanel = False
        'GridViewAnalysis.OptionsBehavior.AlignGroupSummaryInGroupRow = DevExpress.Utils.DefaultBoolean.True
        'GridViewAnalysis.OptionsView.ShowGroupedColumns = False

        ' Expand group rows.

        'GridViewAnalysis.SetGroupLevelExpanded = GroupFooterShowMode.VisibleAlways
        'GridViewAnalysis.Columns(0).Width = 160
        'GridViewAnalysis.Columns(3).Width = 150
        'GridViewAnalysis.Columns(4).Width = 150

        'GridViewAnalysis.Columns(5).Width = 125
        'GridViewAnalysis.Columns(10).Width = 140

        'GridViewAnalysis.Columns(0).Caption = "SOCI Heading"


        'Dim columnTotal As GridColumn = GridViewAnalysis.Columns(51)
        'columnTotal.FilterInfo = New ColumnFilterInfo("[Total] > 0")

        'GridViewAnalysis.Columns(0).SortOrder = DevExpress.Data.ColumnSortOrder.Ascending


        'GridControlSOCIData.MainView?.LayoutChanged()

        '' GridViewAnalysis.OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never
        'GridViewAnalysis.ShowCustomFilterDialog = True
#End Region



        GridViewAnalysis.OptionsView.AllowHtmlDrawGroups = True
        Formatter.FormatGridView(GridViewAnalysis, GridControlSOCIData, "Smaller")

        Exit Sub


    End Sub

    Private Sub gridView_GroupRowCollapsing(ByVal sender As Object, e As DevExpress.XtraGrid.Views.Base.RowAllowEventArgs) Handles GridViewAnalysis.GroupRowCollapsing

        SystemLog("collapsing " & e.RowHandle.ToString)

        'If e.RowHandle Then
        '    e.Allow = False
        'End If

    End Sub
    Private Sub GridView_CustomDrawCell(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs)

        Dim BoldDeffont As New Font(e.Appearance.GetFont, FontStyle.Bold)

        If GridViewAnalysis.GetRowLevel(e.RowHandle) < 1 Then

            e.Appearance.Font = BoldDeffont
            e.Appearance.BackColor = Color.AntiqueWhite

        End If

        Dim strToWrite As String = e.DisplayText


        If IsNumeric(e.CellValue) Then

            If e.CellValue < 0 Then

                If Microsoft.VisualBasic.Left(strToWrite, 1) = "-" Then strToWrite = Microsoft.VisualBasic.Right(strToWrite, Len(strToWrite) - 1)

                e.Appearance.ForeColor = Color.Red
                strToWrite = "(" & strToWrite & ")"

            End If

        End If


        e.Appearance.DrawString(e.Cache, strToWrite, e.Bounds)
        e.Handled = True

    End Sub
    Sub gridViewOverride1_ShowGroupFooter(ByVal sender As Object, ByVal e As ShowGroupFooterEventArgs)

        'Handles myGridView1.ShowGroupFooter

        If e.FooterLevel = 1 Then

            e.Visible = False

        End If

    End Sub
    Private Sub CollapseMasterGroups()

        GridViewAnalysis.BeginUpdate()
        Dim rowHandle As Integer = 0

        Do While rowHandle < GridViewAnalysis.DataRowCount

            GridViewAnalysis.SetRowCellValue(rowHandle, ColName, String.Empty)
            rowHandle += 1

        Loop

        GridViewAnalysis.EndUpdate()

    End Sub

    Private Sub GridViewAnalysis_CustomDrawGroupRowCell(sender As Object, e As RowGroupRowCellEventArgs) Handles GridViewAnalysis.CustomDrawGroupRowCell



        Dim Deffont As New Font(e.Appearance.GetFont, FontStyle.Bold)

        Dim strToWrite As String = e.DisplayText
        Dim BoldDeffont As New Font(e.Appearance.GetFont, FontStyle.Bold)

        If GridViewAnalysis.GetRowLevel(e.RowHandle) < 1 Then

            e.Appearance.Font = BoldDeffont
            e.Appearance.BackColor = Color.AntiqueWhite
            e.Appearance.Options.UseBackColor = True

        End If

        strToWrite = e.DisplayText

        If IsNumeric(e.DisplayText) Then

            If CDbl(e.DisplayText) < 0 Then

                If Microsoft.VisualBasic.Left(strToWrite, 1) = "-" Then strToWrite = Microsoft.VisualBasic.Right(strToWrite, Len(strToWrite) - 1)

                e.Appearance.ForeColor = Color.Red
                strToWrite = "(" & strToWrite & ")"

            End If

        End If
        ' e.Appearance.DrawString(e.Cache, strToWrite, e.Bounds)

        'e.Appearance.BackColor = Color.BlanchedAlmond
        'e.Appearance.FillRectangle(e.Cache, e.Bounds)
        'e.Appearance.ForeColor = Color.DimGray

        e.Appearance.DrawString(e.Cache, strToWrite, e.CaptionBounds)

        e.Handled = True

    End Sub
    Private Sub grid_CustomGroupDisplayText(ByVal sender As Object, ByVal e As RowObjectCustomDrawEventArgs) Handles GridViewAnalysis.CustomDrawGroupRow

        Dim info As GridGroupRowInfo = TryCast(e.Info, GridGroupRowInfo)

        Dim view As GridView = TryCast(sender, GridView)

        If info.Column.FieldName = "SOCIHeading" Or info.Column.FieldName = "SOCIGroup" Then

            Dim GroupingItem As String = Convert.ToString(view.GetGroupRowValue(e.RowHandle, info.Column))

            info.GroupText = Microsoft.VisualBasic.Right(GroupingItem, Len(GroupingItem) - 5)
            'info.GroupText &= "<color=LightSteelBlue>" & view.GetGroupSummaryText(e.RowHandle) & "</color> "
        End If

    End Sub

    Private Sub GridViewAnalysis_CustomColumnDisplayText(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs) Handles GridViewAnalysis.CustomColumnDisplayText

        Dim ciGB As CultureInfo = New CultureInfo("en-GB")

        If IsNumeric(e.Column.FieldName) Then
            ' e.DisplayText = String.Format(ciGB, "{0:c0}", (e.Value * 1000))
            'e.DisplayText = String.Format(ciGB, "{0:c0}", (e.Value))
        End If

    End Sub

    Sub SetDataSource()

        Dim worksheet As DevExpress.Spreadsheet.Worksheet = FileManager.ExcelModels(ModelID).WB.Worksheets("Transactional DB")
        'Dim worksheet As DevExpress.Spreadsheet.Worksheet = WBCoreBP.Worksheets("SOCILiveData")

        Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range("A5:BN1362")
        'Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range("BC1:DA200")

        Dim RDSOptions As New RangeDataSourceOptions With {
            .UseFirstRowAsHeader = True,
            .PreserveFormulas = False,
            .SkipHiddenRows = True,
            .SkipHiddenColumns = True,
            .EditingOptions = DataSourceEditingOptions.ReadOnly
        }

        DSAnalDataRange = range.GetDataSource(RDSOptions)

        AddHandler DSAnalDataRange.ListChanged, AddressOf DataSourceListChangedDelegate

    End Sub

    Private Sub DataSourceListChangedDelegate(ByVal sender As Object, e As System.ComponentModel.ListChangedEventArgs)

        SystemLog("List changed")

        If Not HasSnapshots Then Return


    End Sub





    Sub ResizeFonts()

        Scalefactor = Me.Width / 1700



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
    Private Sub StockAssumptionsInterface_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize

        ResizeFonts()
        ResizeControls()

    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click

        'DSAnalDataRange.

    End Sub


End Class





