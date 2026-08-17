


Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.PresentationManager
Imports Abovo.AbovoUnboundSource
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
Imports Abovo.DataObject
Imports DevExpress.XtraBars.Navigation
Imports DevExpress.XtraPrinting

Imports System.Globalization
Imports DevExpress.Charts.Native
Imports DevExpress.Snap.Core.Commands
Imports DevExpress.XtraRichEdit.Fields
Imports DevExpress.XtraLayout
Imports DevExpress.Utils



Public Class WebInterfaceTemplate

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
    Public MyData As DataObject
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
    Private FooterOn As Boolean
    Private FooterDone As Boolean

    Public Sub New(SetModelID As Integer, SetGSID As Integer, SetCSID As Integer)

        DataSourceCount = -1
        ' This call is required by the designer.
        InitializeComponent()
        Formatter = New ObjectFormatter
        GSID = SetGSID
        CSID = SetCSID
        ModelID = SetModelID
        FooterOn = False

        DataPM = ExcelModels(SetModelID).WBDataPres
        Dim CheckTrans As AbovoTransaction

        CheckTrans = DataPM.ValidatePresentation(GSID, CSID)

        If CheckTrans.BError = False Then

            PresID = CheckTrans.IntegerReturn
            DataPres = DataPM.DataPresentations(PresID)

        End If

        'ManagementCosts.ManagementCostData.GetStatus()
        PopulateWebInterface()

        Exit Sub

    End Sub
    Async Sub PopulateWebInterface()

        Await WebView2Main.EnsureCoreWebView2Async()
        WebView2Main.NavigateToString(DataPres.HTMLOutput)

        Exit Sub

        If DataPres.Sections.Length < 0 Then Exit Sub

        AcControl = New AccordionControl() With {
            .Dock = DockStyle.Fill,
            .Parent = Me,
            .Width = Me.Width
            }

        Formatter.FormatAccordianControl(AcControl)

        SystemLog("Starting")
        AcControl.BeginUpdate()

        Dim Section As PresentationSection
        Dim ActiveDataSet As DataCellRange
        Dim hyperlinkLabelControl1 As New HyperlinkLabelControl()
        AcElementlist = New List(Of AccordionControlElement)

        SystemLog("Presentation object: " & DataPres.Name & " with section count " & DataPres.Sections.Length.ToString)

        For Each Section In DataPres.Sections

            SystemLog("Adding presentation section: " & Section.Name)

            DataSourceCount += 1
            AcElementCount += 1

            ReDim Preserve AcElements(AcElementCount)

            AcElements(AcElementCount) = AcControl.AddItem

            With AcElements(AcElementCount)

                .Text = Section.Name
                '.Style = ElementStyle.Group
                .Name = "Element" & AcElementCount.ToString
                .Expanded = True

            End With

            Formatter.FormatAccordianControlElement(AcElements(AcElementCount))

            SystemLog("Adding element: " & AcElements(AcElementCount).Name)

            AcContainersCount += 1

            ReDim Preserve AcContainers(AcContainersCount)
            AcContainers(AcContainersCount) = New AccordionContentContainer()

            AcControl.Controls.Add(AcContainers(AcContainersCount))
            AcElements(AcElementCount).ContentContainer = AcContainers(AcContainersCount)

            For Each SectionElement In Section.SectionElements


                If SectionElement.Type = "Grid" Then


                    FooterOn = False
                    FooterDone = False

                    ActiveDataSet = DataPres.DataSets(SectionElement.ControlSourceIndex)

                    SystemLog("Adding data from: " & ActiveDataSet.Name)
                    SystemLog("Col count: " & ActiveDataSet.ColCount)
                    SystemLog("Row count: " & ActiveDataSet.RowCount)

                    ColList = New List(Of String)

                    PropertiesCount = -1 'reset

                    ReDim Preserve DataSources(DataSourceCount)
                    Dim SetTag As New AbovoUnboundSourceTag With {.GSID = GSID, .CSID = CSID, .DSIndex = SectionElement.ControlSourceIndex}


                    DataSources(DataSourceCount) = New AbovoUnboundSource(DataSourceCount, SetTag)

                    For Each PresentedColumn In ActiveDataSet.DataColumns

                        ColCount += 1
                        ColName = "Col_" & ColCount.ToString

                        PropertiesCount += 1
                        ReDim Preserve PropertyArray(PropertiesCount)

                        Select Case PresentedColumn.ColumnTag.DataType
                            Case "S"
                                PropType = GetType(String)
                            Case "I", "Y"
                                PropType = GetType(Integer)
                            Case "N", "P", "C"
                                PropType = GetType(Double)
                            Case "B"
                                PropType = GetType(Integer)
                            Case Else
                                PropType = GetType(String)
                        End Select
                        SystemLog("Adding column: " & PresentedColumn.ColumnTag.ColumnHeading)
                        PropertyArray(PropertiesCount) = New UnboundSourceProperty With {
                            .UserTag = ColCount,
                            .Name = ColName,
                            .PropertyType = PropType,
                            .DisplayName = PresentedColumn.ColumnTag.ColumnHeading
                        }

                        ColList.Add(ColName)

                    Next

                    PropertyList = PropertyArray

                    DataSources(DataSourceCount).Properties.AddRange(PropertyList)

                    AddHandler DataSources(DataSourceCount).ValueNeeded, AddressOf UnboundDS_ValueNeeded
                    AddHandler DataSources(DataSourceCount).ValuePushed, AddressOf UnboundDS_ValuePushed

                    GridCount += 1
                    ReDim Preserve GridControls(GridCount)

                    GridControls(GridCount) = New GridControl() With {
                        .Name = "GridControl_" & GridCount.ToString,
                        .Parent = Me,
                        .Dock = DockStyle.Fill,
                        .DataSource = DataSources(DataSourceCount)
                    }

                    GridViewCount += 1
                    ReDim Preserve UsedGridViews(GridViewCount)

                    UsedGridViews(GridViewCount) = New DevExpress.XtraGrid.Views.Grid.GridView
                    GridControls(GridCount).ViewCollection.Add(UsedGridViews(GridViewCount))
                    GridControls(GridCount).MainView = UsedGridViews(GridViewCount)
                    UsedGridViews(GridViewCount).PopulateColumns()

                    SystemLog("Calling records " & ActiveDataSet.RowCount)

                    Dim testUBS As AbovoUnboundSource = TryCast(GridControls(GridCount).DataSource, AbovoUnboundSource)
                    SystemLog("Adding grid control with datasource ")

                    DataSources(DataSourceCount).SetRowCount(ActiveDataSet.RowCount)



                    AcContainers(AcContainersCount).Controls.Add(GridControls(GridCount))

                    AcContainers(AcContainersCount).Height = 900
                    AcContainers(AcContainersCount).Width = Me.Width
                    AcContainers(AcContainersCount).Appearance.BackColor = Color.Aquamarine
                    GridControls(GridCount).Dock = DockStyle.Fill
                    GridControls(GridCount).ForceInitialize()




                    AddHandler UsedGridViews(GridViewCount).CustomColumnDisplayText, AddressOf SetGridDisplayText
                    Dim ThisCol As Integer

                    For ThisCol = 0 To UsedGridViews(GridViewCount).Columns.Count - 1
                        UsedGridViews(GridViewCount).Columns(ThisCol).Tag = ActiveDataSet.DataColumns(ThisCol).ColumnTag
                    Next

                    Dim ColTag As DataColumnTag
                    Dim GVcolumn As GridColumn
                    Dim siTotal As GridColumnSummaryItem
                    'siTotal.SummaryType = SummaryItemType.Sum
                    'siTotal.DisplayFormat = "{0} records"
                    'colOrderID.Summary.Add(siTotal)

                    For Each GVcolumn In UsedGridViews(GridViewCount).Columns

                        ColTag = GVcolumn.Tag

                        Select Case ColTag.DataType
                            Case "S"
                                GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
                                GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.None
                            Case "I"
                                GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                                GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                                GVcolumn.DisplayFormat.FormatString = "#,###,##0"
                            Case "P"
                                GVcolumn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
                                GVcolumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                                GVcolumn.DisplayFormat.FormatString = "p2"
                            Case "B"

                            Case Else

                        End Select

                        If ColTag.ShowSummary = "Sum" Then
                            If Not FooterDone Then SetFooterOn(UsedGridViews(GridViewCount))
                            FooterOn = True
                            siTotal = New GridColumnSummaryItem
                            siTotal.SummaryType = SummaryItemType.Sum
                            siTotal.DisplayFormat = "{0:n0}"

                            GVcolumn.Summary.Add(siTotal)
                        ElseIf ColTag.ShowSummary = "Count" Then
                            If Not FooterDone Then SetFooterOn(UsedGridViews(GridViewCount))
                            siTotal = New GridColumnSummaryItem
                            siTotal.SummaryType = SummaryItemType.Count
                            siTotal.DisplayFormat = "{0:n0}"
                            GVcolumn.Summary.Add(siTotal)
                        End If

                    Next

                    Formatter.FormatGridControl(GridControls(GridCount))
                    Formatter.FormatGridView(UsedGridViews(GridViewCount), GridControls(GridCount))

                    UsedGridViews(GridViewCount).BestFitColumns()

                    UsedGridViews(GridViewCount).OptionsView.BestFitMaxRowCount = ActiveDataSet.RowCount
                Else



                End If

            Next







        Next
        If FooterOn = True Then

            'UsedGridViews(GridViewCount).FooterPanelHeight = 70
        End If
        AcControl.ExpandElementMode = ExpandElementMode.Multiple
        AcControl.ExpandAll()
        AcControl.EndUpdate()



    End Sub
    Sub SetFooterOn(SetGridView As GridView)
        FooterOn = True
        UsedGridViews(GridViewCount).OptionsView.ShowFooter = True
        Dim itemCust As New GridGroupSummaryItem
        itemCust.FieldName = UsedGridViews(GridViewCount).Columns(0).FieldName
        itemCust.SummaryType = DevExpress.Data.SummaryItemType.Custom
        itemCust.DisplayFormat = "Total:"
        itemCust.ShowInGroupColumnFooter = UsedGridViews(GridViewCount).Columns(0)
        UsedGridViews(GridViewCount).GroupSummary.Add(itemCust)
        AddHandler UsedGridViews(GridViewCount).CustomDrawFooter, AddressOf GVCustomDrawFooter
        FooterDone = True
    End Sub
    Private Sub SetGridDisplayText(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs)
        Dim ColTag As DataColumnTag = e.Column.Tag
        Select Case ColTag.DataType
            Case "C"
                Dim ciGB As CultureInfo = New CultureInfo("en-GB")
                e.DisplayText = String.Format(ciGB, "{0:c0}", (e.Value))
        End Select
    End Sub
    Private Sub GVCustomDrawFooter(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Base.RowObjectCustomDrawEventArgs)
        If Not FooterOn Then Exit Sub
        SystemLog("CDF:" & sender.name)
        'Dim stringFormat As StringFormat = New StringFormat()
        'stringFormat.Alignment = StringAlignment.Near
        'stringFormat.LineAlignment = StringAlignment.Center
        'Dim rect = e.Bounds
        'rect.X += 10
        'e.DefaultDraw()
        'e.Cache.DrawString("Total:", e.Appearance.GetFont(), e.Appearance.GetForeBrush(e.Cache), rect, stringFormat)
        'e.Handled = True
    End Sub

    Private Sub InitAccordionControl()
        AcControl.BeginUpdate()
        Dim acRootGroupHome As New AccordionControlElement()
        Dim acItemActivity As New AccordionControlElement()
        Dim acItemNews As New AccordionControlElement()
        Dim acRootItemSettings As New AccordionControlElement()

        AddHandler AcControl.ElementClick, AddressOf AccordionControl1_ElementClick

        ' 
        ' Root Group 'Home'
        ' 
        acRootGroupHome.Elements.AddRange(New AccordionControlElement() {acItemActivity, acItemNews})
        acRootGroupHome.Expanded = True
        acRootGroupHome.ImageOptions.ImageUri.Uri = "Home;Office2013"
        acRootGroupHome.Name = "acRootGroupHome"
        acRootGroupHome.Text = "Home"
        ' 
        ' Child Item 'Activity'
        ' 
        acItemActivity.Name = "acItemActivity"
        acItemActivity.Style = ElementStyle.Item
        acItemActivity.Tag = "idActivity"
        acItemActivity.Text = "Activity"
        ' 
        ' Child Item 'News'
        ' 
        acItemNews.Name = "acItemNews"
        acItemNews.Style = ElementStyle.Item
        acItemNews.Tag = "idNews"
        acItemNews.Text = "News"
        ' 
        ' Root Item 'Settings' with ContentContainer
        ' 
        acRootItemSettings.ImageOptions.ImageUri.Uri = "Customization;Office2013"
        acRootItemSettings.Name = "acRootItemSettings"
        acRootItemSettings.Style = ElementStyle.Item
        acRootItemSettings.Text = "Settings"
        ' 
        ' itemSettingsControlContainer
        ' 
        Dim itemSettingsControlContainer As New AccordionContentContainer()
        Dim hyperlinkLabelControl1 As New HyperlinkLabelControl()
        Dim toggleSwitch1 As New ToggleSwitch()
        AcControl.Controls.Add(itemSettingsControlContainer)
        acRootItemSettings.ContentContainer = itemSettingsControlContainer
        itemSettingsControlContainer.Controls.Add(hyperlinkLabelControl1)
        itemSettingsControlContainer.Controls.Add(toggleSwitch1)
        itemSettingsControlContainer.Appearance.BackColor = System.Drawing.SystemColors.Control
        itemSettingsControlContainer.Appearance.Options.UseBackColor = True
        itemSettingsControlContainer.Height = 60
        ' 
        ' hyperlinkLabelControl1
        ' 
        hyperlinkLabelControl1.Location = New System.Drawing.Point(26, 33)
        hyperlinkLabelControl1.Size = New System.Drawing.Size(107, 13)
        hyperlinkLabelControl1.Text = "www.devexpress.com"
        AddHandler hyperlinkLabelControl1.HyperlinkClick, AddressOf HyperlinkLabelControl1_HyperlinkClick
        ' 
        ' toggleSwitch1
        ' 
        toggleSwitch1.EditValue = True
        toggleSwitch1.Location = New System.Drawing.Point(24, 3)
        toggleSwitch1.Properties.AllowFocused = False
        toggleSwitch1.Properties.AutoWidth = True
        toggleSwitch1.Properties.OffText = "Offline Mode"
        toggleSwitch1.Properties.OnText = "Onlne Mode"
        toggleSwitch1.Size = New System.Drawing.Size(134, 24)
        AddHandler toggleSwitch1.Toggled, AddressOf toggleSwitch1_Toggled

        AcControl.Elements.AddRange(New DevExpress.XtraBars.Navigation.AccordionControlElement() {acRootGroupHome, acRootItemSettings})

        acRootItemSettings.Expanded = True

        AcControl.EndUpdate()
    End Sub

    Private Sub AccordionControl1_ElementClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.Navigation.ElementClickEventArgs)
        If e.Element.Style = DevExpress.XtraBars.Navigation.ElementStyle.Group Then
            Return
        End If
        If e.Element.Tag Is Nothing Then
            Return
        End If
        Dim itemID As String = e.Element.Tag.ToString()
        If itemID = "idNews" Then
            '...
        End If
        'listBoxControl1.Items.Add(itemID & " clicked")
    End Sub

    Private Sub toggleSwitch1_Toggled(ByVal sender As Object, ByVal e As EventArgs)
        '...
    End Sub

    Private Sub HyperlinkLabelControl1_HyperlinkClick(ByVal sender As Object, ByVal e As DevExpress.Utils.HyperlinkClickEventArgs)
        Process.Start(e.Text)
    End Sub
    Private Sub UnboundDS_ValueNeeded(ByVal sender As Object, ByVal e As DevExpress.Data.UnboundSourceValueNeededEventArgs)

        Dim UDSSender As AbovoUnboundSource = sender
        e.Value = GetDSData(UDSSender.UBSTag.DSIndex, e.RowIndex, e.PropertyIndex)
    End Sub
    Private Sub UnboundDS_ValuePushed(ByVal sender As Object, ByVal e As DevExpress.Data.UnboundSourceValuePushedEventArgs)
        'something = e.Value ' Propagate the value into the storage.
    End Sub
    Private Function GetDSData(ByVal SetDSIndex As Integer, ByVal rowIndex As Integer, ByVal PropertyIndex As Integer) As Object

        'SystemLog("Value requested from dataset: " & SetDSIndex.ToString & " Row: " & rowIndex.ToString & " Column: " & PropertyIndex.ToString)

        Dim DP As CellDataPoint = DataPres.DataSets(SetDSIndex).DataRows(rowIndex).DataCells(PropertyIndex)

        'SystemLog("Returning value of type: " & DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag.DataType)

        Select Case DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag.DataType

            Case "S"

                Return DP.StringValue

            Case "B"

                Return DP.BoolValue

            Case "N", "P", "C"

                Return DP.RealValue

            Case "I", "Y"

                Return DP.IntValue

            Case Else
                Return Nothing

        End Select

    End Function

    'Public Shared Sub CustomDrawColumnHeader(ByVal gridControl As GridControl, ByVal gridView As GridView)
    '    ' Handle this event to paint columns headers manually
    '    AddHandler gridView.CustomDrawColumnHeader, Sub(s, e)
    '                                                    If e.Column Is Nothing OrElse e.Column.FieldName <> "Category_Name" Then
    '                                                        Return
    '                                                    End If
    '                                                    ' Fill column headers with the specified colors.
    '                                                    e.Cache.FillRectangle(Color.Coral, e.Bounds)
    '                                                    e.Appearance.DrawString(e.Cache, e.Info.Caption, e.Info.CaptionRect)
    '                                                    ' Draw the filter and sort buttons.
    '                                                    For Each info As DrawElementInfo In e.Info.InnerElements
    '                                                        If Not info.Visible Then
    '                                                            Continue For
    '                                                        End If
    '                                                        ObjectPainter.DrawObject(e.Cache, info.ElementPainter, info.ElementInfo)
    '                                                    Next info
    '                                                    e.Handled = True
    '                                                End Sub
    'End Sub
    'Private Sub gridView1_CustomDrawColumnHeader(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.ColumnHeaderCustomDrawEventArgs)
    '    e.Graphics.FillRectangle(Brushes.Green, e.Bounds)
    '    Using pen As New Pen(Color.Black, 3)
    '        e.Graphics.DrawRectangle(pen, e.Bounds)
    '    End Using
    '    e.Info.InnerElements.DrawObjects(e.Info, e.Cache, Point.Empty)
    '    e.Handled = True
    'End Sub
#Region "Interface events"

    Sub ResizeMe()
        Dim ScaleFactor As Single = Me.Width / 1920

    End Sub
    Private Sub WindowsUIButtonPanelItemEdit_ButtonChecked(sender As Object, e As ButtonEventArgs)
        Dim tag As String = CType(e.Button, WindowsUIButton).Tag.ToString()
        Select Case tag
            Case "AddDispo"


                'GridViewStockNumbers.Columns(11).Visible = False
                'GridViewStockNumbers.Columns(9).Visible = False
                'GridViewStockNumbers.Columns(8).Visible = False
                'GridViewStockNumbers.Columns(7).Visible = False
                'GridViewStockNumbers.Columns(6).Visible = False
                'GridViewStockNumbers.Columns(5).Visible = False
                'Grid1ExpandedView = False
                'SortGridColums()
                'CustomDrawCell(GridControlStockGrid, GridViewStockNumbers)
                'GridViewStockNumbers.LeftCoord = 0
        End Select
    End Sub
    Private Sub WindowsUIButtonPanelItemEdit_ButtonCheck(sender As Object, e As DevExpress.XtraBars.Docking2010.ButtonEventArgs)
        Dim tag As String = CType(e.Button, WindowsUIButton).Tag.ToString()
        Select Case tag
            Case "AddDispo"

                'GridViewStockNumbers.Columns(10).Visible = False

                'GridViewStockNumbers.Columns(5).Visible = True

                'GridViewStockNumbers.Columns(6).Visible = True

                'GridViewStockNumbers.Columns(7).Visible = True

                'GridViewStockNumbers.Columns(8).Visible = True

                'GridViewStockNumbers.Columns(9).Visible = True

                'GridViewStockNumbers.Columns(10).Visible = True

                'GridViewStockNumbers.Columns(11).Visible = True

                'Grid1ExpandedView = True
                'SortGridColums()
                'CustomDrawCell(GridControlStockGrid, GridViewStockNumbers)
                'Dim Column As DevExpress.XtraGrid.Columns.GridColumn = GridViewStockNumbers.Columns(5)
                'Dim info As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridViewInfo = GridViewStockNumbers.GetViewInfo()
                'GridViewStockNumbers.LeftCoord = info.GetColumnLeftCoord(Column) + GridViewStockNumbers.Columns(0).Width
        End Select
    End Sub

    Private Sub WindowsUIButtonPanelBPActions_ButtonClick(sender As Object, e As DevExpress.XtraBars.Docking2010.ButtonEventArgs)
        Dim tag As String = CType(e.Button, WindowsUIButton).Tag.ToString()
        Select Case tag
            Case "ApplyAndSave"
                ' OpenAssumptionsInterface
                WriteStockToBPAndSave()
            Case "ApplyToFile"
                ' Navigate to page B 
                WriteStockToBP()
            Case "Ad3"
                    ' Navigate to page C
            Case "Ad4"
                    ' Navigate to page D 
            Case "Ad5"
                ' Navigate to page E 
        End Select
    End Sub
    Private Sub RepositoryItemComboBoxSOCIRent_QueryPopUp(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim lookUpEdit As LookUpEdit = TryCast(sender, LookUpEdit)
        lookUpEdit.Properties.PopulateColumns()
        lookUpEdit.Properties.Columns(0).Visible = False

    End Sub
    Private Sub RepositoryItemComboBoxSOCIStocktype_QueryPopUp(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim lookUpEdit As LookUpEdit = TryCast(sender, LookUpEdit)
        lookUpEdit.Properties.PopulateColumns()
        lookUpEdit.Properties.Columns(0).Visible = False
        lookUpEdit.Properties.Columns(1).Visible = False
    End Sub
#End Region
#Region "Data events"

    Private Sub WriteStockToBP()

        ' AbovoBP.WriteStock()
        BIsDirty = False

    End Sub
    Private Sub WriteStockToBPAndSave()

        ' AbovoBP.WriteStock()
        BIsDirty = False
        ' AbovoBP.SaveBP()

    End Sub




    Private Sub SetArrayData(ByVal rowIndex As Integer, ByVal propertyName As String, ByVal value As Object)

        Select Case propertyName

            Case "PropertyStockDescription"
                AbovoBP.Stock.StockItems(rowIndex).StockDescription = value
            Case "PropertyOwnedManaged"
                AbovoBP.Stock.StockItems(rowIndex).OwnedManaged = value
            Case "PropertySOCIStockType"
                AbovoBP.Stock.StockItems(rowIndex).SOCIStockType = value
            Case "PropertySOCIRentType"
                AbovoBP.Stock.StockItems(rowIndex).SOCIRentType = value
            Case "PropertyCurrentStockNumbers"
                AbovoBP.Stock.StockItems(rowIndex).CurrentStockNumbers = value
            Case "PropertyInitialRateNewLettings"
                AbovoBP.Stock.StockItems(rowIndex).NewLetInitialRate = value
            Case "PropertyPreBPlanStartDateNewBuild"
                AbovoBP.Stock.StockItems(rowIndex).PreBPlanStartDateNewBuild = value
            Case "PropertyPreBPlanStartDateDemolitions"
                AbovoBP.Stock.StockItems(rowIndex).PreBPlanStartDateDemolitions = value
            Case "PropertyPreBPlanStartDateRTBs"
                AbovoBP.Stock.StockItems(rowIndex).PreBPlanStartDateRTBs = value
            Case "PropertyPreBPlanStartDateOtherDisposals"
                AbovoBP.Stock.StockItems(rowIndex).PreBPlanStartDateOtherDisposals = value
            Case "PropertyNewLettings"
                AbovoBP.Stock.StockItems(rowIndex).NewLettings = value
            Case Else

        End Select

        AbovoBP.Stock.StockItems(rowIndex).FUpdateStockTotals()

    End Sub
#End Region



    Private Sub GridViewStockNumbers_ValidatingEditor(sender As Object, e As DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventArgs)

        Dim view As ColumnView = sender
        Dim column As GridColumn = If(TryCast(e, EditFormValidateEditorEventArgs)?.Column, view.FocusedColumn)



        If column.Name = "colPropertyInitialRateNewLettings1" Then

            If (Convert.ToDecimal(e.Value) < 0) Or (Convert.ToDecimal(e.Value) > 1) Then
                MsgBox("Sorry, The value of initial New Lettings Rate must be more than 0 and less than 100", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If

            Exit Sub

        ElseIf column.Name = "colPropertyCurrentStockNumbers1" Then

            If Convert.ToInt32(e.Value) < 0 Then
                MsgBox("Sorry, the value of Current Stock Numbers must be more than 0", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If
            Exit Sub
        ElseIf column.Name = "colPropertyPreBPlanStartDateNewBuild1" Then

            If Convert.ToInt32(e.Value) < 0 Then
                MsgBox("Sorry, the value of New Build Numbers must be more than 0", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If
            Exit Sub
        ElseIf column.Name = "colPropertyPreBPlanStartDateDemolitions1" Then

            If Convert.ToInt32(e.Value) < 0 Then
                MsgBox("Sorry, the value of Demolitions must be more than 0", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If
            Exit Sub
        ElseIf column.Name = "colPropertyPreBPlanStartDateRTBs1" Then

            If Convert.ToInt32(e.Value) < 0 Then
                MsgBox("Sorry, the value of Right To Buys must be more than 0", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If
            Exit Sub
        ElseIf column.Name = "colPropertyPreBPlanStartDateOtherDisposals1" Then

            If Convert.ToInt32(e.Value) < 0 Then
                MsgBox("Sorry, the value of Other Disposals must be more than 0", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If
            Exit Sub
        ElseIf column.Name = "colPropertyNewLettings1" Then

            If Convert.ToInt32(e.Value) < 0 Then
                MsgBox("Sorry, the value of New Lettings must be more than 0", vbOKOnly, "Assumption Validation Error")
                e.Valid = False
            End If
            Exit Sub
        End If

    End Sub

    Private Sub GridViewStockNumbers_InvalidValueException(sender As Object, e As InvalidValueExceptionEventArgs)
        Dim view As ColumnView = sender

        view.HideEditor()
        Exit Sub
        'Dim view As ColumnView9 = sender
        'If view Is Nothing Then
        '    Return
        'End If
        'Dim column As GridColumn = If(TryCast(e, InvalidValueExceptionEventArgs)?.Column, view.FocusedColumn)
        'e.ExceptionMode = ExceptionMode.DisplayError
        'e.WindowCaption = "Input Error"
        'e.ErrorText = "The value should be greater than 0 and less than 100"
        '' Destroy the editor and discard the changes made within the edited cell
        'view.HideEditor()
    End Sub
    Sub SortGridColums()

        Dim np As Integer = 0
        'For i = 0 To GridViewStockNumbers.Columns.Count - 1
        '    If GridViewStockNumbers.Columns(i).Visible Then
        '        GridViewStockNumbers.Columns(i).VisibleIndex = np
        '        np += 1
        '    End If


        'Next i
    End Sub


    Private Sub RepositoryItemLookUpEditSOCIStockType_EditValueChanged(sender As Object, e As EventArgs)
        'Dim Editor As LookUpEdit = CType(sender, LookUpEdit)

        'Dim StrChosenStock As String = Convert.ToString(Editor.EditValue)
        'Dim IntFoundCatID As Integer = SOCIStock.GetSOCICategoryByName(StrChosenStock)
        'If IntFoundCatID = 0 Then GridViewStockNumbers.SetFocusedRowCellValue("PropertySOCIRentType", Convert.ToString("N/A"))
    End Sub

    Private Sub GridViewStockNumbers_CustomDrawColumnHeader(sender As Object, e As ColumnHeaderCustomDrawEventArgs)
        e.Cache.FillRectangle(Color.White, e.Bounds)
        e.Appearance.DrawString(e.Cache, e.Info.Caption, e.Info.CaptionRect)
        Using pen As New Pen(Color.Silver, 4)
            e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom, e.Bounds.Right + 4, e.Bounds.Bottom)
            'e.Graphics.DrawRectangle(pen, e.Bounds)
        End Using
        ' Draw the filter and sort buttons.
        For Each info As DrawElementInfo In e.Info.InnerElements
            If Not info.Visible Then
                Continue For
            End If
            ObjectPainter.DrawObject(e.Cache, info.ElementPainter, info.ElementInfo)
        Next info
        e.Handled = True
    End Sub



    Sub CustomDrawCell(ByVal gridControl As GridControl, ByVal gridView As GridView)
        ' Handle this event to paint cells manually
        Dim BDo As Boolean = False
        AddHandler gridView.CustomDrawCell,
            Sub(s, e)
                If Grid1ExpandedView Then
                    If e.Column.VisibleIndex = 0 Then
                        Using pen As New Pen(Color.Silver, 4)
                            e.Graphics.DrawLine(pen, e.Bounds.Right, e.Bounds.Top - 4, e.Bounds.Right, e.Bounds.Bottom + 15)
                            'e.Graphics.DrawRectangle(pen, e.Bounds)
                        End Using

                    End If
                    If e.Column.VisibleIndex = 12 Then
                        Using pen As New Pen(Color.Silver, 4)
                            e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Top - 4, e.Bounds.Left, e.Bounds.Bottom + 15)
                            'e.Graphics.DrawRectangle(pen, e.Bounds)
                        End Using

                    End If
                Else
                    If e.Column.VisibleIndex = 0 Then
                        Using pen As New Pen(Color.White, 4)
                            e.Graphics.DrawLine(pen, e.Bounds.Right, e.Bounds.Top, e.Bounds.Right, e.Bounds.Bottom)
                            'e.Graphics.DrawRectangle(pen, e.Bounds)
                        End Using

                    End If
                    If e.Column.VisibleIndex = 12 Then
                        Using pen As New Pen(Color.White, 4)
                            e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Top - 4, e.Bounds.Left, e.Bounds.Bottom + 15)
                            'e.Graphics.DrawRectangle(pen, e.Bounds)
                        End Using

                    End If
                End If
                If BDo Then
                    For Each info As DrawElementInfo In e.Cell.InnerElements
                        If Not info.Visible Then
                            Continue For
                        End If
                        ObjectPainter.DrawObject(e.Cache, info.ElementPainter, info.ElementInfo)
                    Next info
                    e.Handled = True
                End If

            End Sub
    End Sub



    Private Sub Button1_Click_1(sender As Object, e As EventArgs)


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

End Class





