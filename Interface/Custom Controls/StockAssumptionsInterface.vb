


Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.PresentationManager
Imports Abovo.AbovoUnboundSource
Imports Abovo.FileManager
Imports Abovo.LogDebugDev
Imports Abovo.DataObject

Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraEditors
Imports DevExpress.XtraRichEdit.Import.Html
Imports DevExpress.XtraSpreadsheet.Model
Imports DevExpress.Snap.Core.API
Imports DevExpress.Snap.Core.Native
Imports System
Imports System.Data
Imports System.Linq
Imports System.Collections.Generic
Imports System.ComponentModel.DataAnnotations
Imports DevExpress.CodeParser
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraGrid
Imports System.Runtime.InteropServices
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.XtraGrid.Extensions
Imports DevExpress.Skins
Imports DevExpress.XtraEditors.Mask.Design.MaskSettingsForm.DesignInfo.MaskManagerInfo
Imports DevExpress.LookAndFeel
Imports DevExpress.Utils.Drawing
Imports DevExpress.XtraExport.Helpers
Imports System.Globalization
Imports DevExpress.Utils
Imports DevExpress.ClipboardSource.SpreadsheetML
Imports DevExpress.Data
Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms
Imports DevExpress.Pdf.Native.BouncyCastle.Ocsp
Imports System.Runtime.Serialization
Imports DevExpress.XtraBars.Navigation






Public Class StockAssumptionsInterface
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

    Public Sub New(SetModelID As Integer, SetGSID As Integer, SetCSID As Integer)


        ' This call is required by the designer.
        InitializeComponent()
        'CustomDrawColumnHeader(GridControlStockGrid, GridViewStockNumbers)

        rs.FindAllControls(Me)
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
        PopulateDataInterface()

        GridViewStockNumbers.OptionsView.EnableAppearanceEvenRow = True
        GridViewStockNumbers.OptionsView.EnableAppearanceOddRow = True

        'XtraTabControlStockViews.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        'GridControlStockGrid.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        'GridControlStockGrid.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        'Me.TablePanelStockAssumptions.LookAndFeel.UseDefaultLookAndFeel = False
        'Me.TablePanelStockAssumptions.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat



        'Me.WindowsUIButtonPanelItemEdit.ForeColor = Color.SteelBlue
        'Me.GridViewStockNumbers.Appearance.EvenRow.BackColor = System.Drawing.Color.AliceBlue
        'Me.GridViewStockNumbers.Appearance.EvenRow.BorderColor = System.Drawing.Color.AliceBlue
        'Me.GridViewStockNumbers.Appearance.EvenRow.Font = New System.Drawing.Font("Segoe UI Variable Display", 9.857143!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'Me.GridViewStockNumbers.Appearance.EvenRow.Options.UseBackColor = True
        'Me.GridViewStockNumbers.Appearance.EvenRow.Options.UseBorderColor = True
        'Me.GridViewStockNumbers.Appearance.EvenRow.Options.UseFont = True
        Me.GridViewStockNumbers.Appearance.HorzLine.BackColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.HorzLine.BorderColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.HorzLine.ForeColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.HorzLine.Options.UseBackColor = True
        Me.GridViewStockNumbers.Appearance.HorzLine.Options.UseBorderColor = True
        Me.GridViewStockNumbers.Appearance.HorzLine.Options.UseForeColor = True
        'Me.GridViewStockNumbers.Appearance.OddRow.BackColor = System.Drawing.Color.White
        'Me.GridViewStockNumbers.Appearance.OddRow.BorderColor = System.Drawing.Color.White
        'Me.GridViewStockNumbers.Appearance.OddRow.Font = New System.Drawing.Font("Segoe UI Variable Display", 9.857143!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'Me.GridViewStockNumbers.Appearance.OddRow.Options.UseBackColor = True
        'Me.GridViewStockNumbers.Appearance.OddRow.Options.UseBorderColor = True
        'Me.GridViewStockNumbers.Appearance.OddRow.Options.UseFont = True
        GridViewStockNumbers.Appearance.VertLine.BackColor = System.Drawing.Color.White
        GridViewStockNumbers.Appearance.VertLine.BorderColor = System.Drawing.Color.White
        GridViewStockNumbers.Appearance.VertLine.ForeColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.VertLine.Options.UseBackColor = True
        Me.GridViewStockNumbers.Appearance.VertLine.Options.UseBorderColor = True
        Me.GridViewStockNumbers.Appearance.VertLine.Options.UseForeColor = True

        RepositoryItemLookUpEditSOCIRentType.DataSource = SOCIRentType.Init()
        RepositoryItemLookUpEditSOCIRentType.DisplayMember = "SOCIRentName"
        RepositoryItemLookUpEditSOCIRentType.ValueMember = "SOCIRentName"
        RepositoryItemLookUpEditSOCIStockType.DataSource = SOCIStock.Init
        RepositoryItemLookUpEditSOCIStockType.DisplayMember = "SOCIStockName"
        RepositoryItemLookUpEditSOCIStockType.ValueMember = "SOCIStockName"

        GridControlStockGrid.ForceInitialize()

        AddHandler GridViewStockNumbers.ShownEditor, AddressOf GridViewStockNumbers_ShownEditor
        'Dim ciGB As CultureInfo = New CultureInfo("en-GB")

        'If IsNumeric(e.Column.FieldName) Then
        '    e.DisplayText = String.Format(ciGB, "{0:c0}", e.Value)
        'End If


        GridViewStockNumbers.Columns(4).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        GridViewStockNumbers.Columns(4).DisplayFormat.FormatString = "#,###,##0"
        GridViewStockNumbers.Columns(5).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        GridViewStockNumbers.Columns(5).DisplayFormat.FormatString = "#,###,##0"
        GridViewStockNumbers.Columns(6).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        GridViewStockNumbers.Columns(6).DisplayFormat.FormatString = "#,###,##0"
        GridViewStockNumbers.Columns(7).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        GridViewStockNumbers.Columns(7).DisplayFormat.FormatString = "#,###,##0"
        GridViewStockNumbers.Columns(8).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        GridViewStockNumbers.Columns(8).DisplayFormat.FormatString = "#,###,##0"
        GridViewStockNumbers.Columns(9).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        GridViewStockNumbers.Columns(9).DisplayFormat.FormatString = "#,###,##0"
        GridViewStockNumbers.Columns(10).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        GridViewStockNumbers.Columns(10).DisplayFormat.FormatString = "#,###,##0"
        GridViewStockNumbers.Columns(12).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        GridViewStockNumbers.Columns(12).DisplayFormat.FormatString = "#,###,##0"

        'GridViewStockNumbers.Columns(5).Visible = False
        'GridViewStockNumbers.Columns(6).Visible = False
        'GridViewStockNumbers.Columns(7).Visible = False
        'GridViewStockNumbers.Columns(8).Visible = False
        'GridViewStockNumbers.Columns(9).Visible = False
        'GridViewStockNumbers.Columns(11).Visible = False
        'MsgBox(SkinManager.DefaultSkinName)

        Dim element As SkinElement = SkinManager.GetSkinElement(SkinProductId.Grid, DevExpress.LookAndFeel.UserLookAndFeel.Default, "Header")
        element.Border.Thin.Bottom = -1
        element.Border.Thin.Right = -1
        element.Border.Thin.Left = -1
        element.Border.Thin.Top = -1

        element = SkinManager.GetSkinElement(SkinProductId.Grid, DevExpress.LookAndFeel.UserLookAndFeel.Default, "HeaderRight")
        element.Border.Thin.Bottom = -1
        element.Border.Thin.Right = -1
        element.Border.Thin.Left = -1
        element.Border.Thin.Top = -1

        element = SkinManager.GetSkinElement(SkinProductId.Grid, DevExpress.LookAndFeel.UserLookAndFeel.Default, "HeaderLeft")
        element.Border.Thin.Bottom = -1
        element.Border.Thin.Right = -1
        element.Border.Thin.Left = -1
        element.Border.Thin.Top = -1

        element = SkinManager.GetSkinElement(SkinProductId.Grid, DevExpress.LookAndFeel.UserLookAndFeel.Default, "SingleRowHeader")
        If element IsNot Nothing Then
            element.Border.Thin.Bottom = -1
            element.Border.Thin.Right = -1
            element.Border.Thin.Left = -1
            element.Border.Thin.Top = -1
        End If

        LookAndFeelHelper.ForceDefaultLookAndFeelChanged()
        Grid1ExpandedView = False
        BIsDirty = False
        ResizeMe()
        CustomDrawCell(GridControlStockGrid, GridViewStockNumbers)

        'CustomDrawColumnHeader(GridControlStockGrid, GridViewStockNumbers)
    End Sub
    Private Sub PopulateDataInterface()

        If DataPres.Sections.Length < 0 Then Exit Sub

        AcControl = New AccordionControl() With {
            .Dock = DockStyle.Fill,
            .Parent = Me,
            .Width = Me.Width
            }

        Formatter.FormatAccordianControl(AcControl)

        AcControl.BeginUpdate()

        Dim Section As PresentationSection
        Dim ActiveDataSet As DataCellRange
        Dim hyperlinkLabelControl1 As New HyperlinkLabelControl()
        AcElementlist = New List(Of AccordionControlElement)


        For Each Section In DataPres.Sections


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

                    'AddHandler DataSources(DataSourceCount).ValueNeeded, AddressOf UnboundDS_ValueNeeded
                    'AddHandler DataSources(DataSourceCount).ValuePushed, AddressOf UnboundDS_ValuePushed

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


                    Dim testUBS As AbovoUnboundSource = TryCast(GridControls(GridCount).DataSource, AbovoUnboundSource)

                    DataSources(DataSourceCount).SetRowCount(ActiveDataSet.RowCount)



                    AcContainers(AcContainersCount).Controls.Add(GridControls(GridCount))

                    AcContainers(AcContainersCount).Height = 900
                    AcContainers(AcContainersCount).Width = Me.Width
                    AcContainers(AcContainersCount).Appearance.BackColor = Color.Aquamarine
                    GridControls(GridCount).Dock = DockStyle.Fill
                    GridControls(GridCount).ForceInitialize()




                    'AddHandler UsedGridViews(GridViewCount).CustomColumnDisplayText, AddressOf SetGridDisplayText
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
                            'If Not FooterDone Then SetFooterOn(UsedGridViews(GridViewCount))
                            FooterOn = True
                            siTotal = New GridColumnSummaryItem
                            siTotal.SummaryType = SummaryItemType.Sum
                            siTotal.DisplayFormat = "{0:n0}"

                            GVcolumn.Summary.Add(siTotal)
                        ElseIf ColTag.ShowSummary = "Count" Then
                            'If Not FooterDone Then SetFooterOn(UsedGridViews(GridViewCount))
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
    Private Sub WindowsUIButtonPanelItemEdit_ButtonChecked(sender As Object, e As ButtonEventArgs) Handles WindowsUIButtonPanelItemEdit.ButtonUnchecked
        Dim tag As String = CType(e.Button, WindowsUIButton).Tag.ToString()
        Select Case tag
            Case "AddDispo"


                GridViewStockNumbers.Columns(11).Visible = False
                GridViewStockNumbers.Columns(9).Visible = False
                GridViewStockNumbers.Columns(8).Visible = False
                GridViewStockNumbers.Columns(7).Visible = False
                GridViewStockNumbers.Columns(6).Visible = False
                GridViewStockNumbers.Columns(5).Visible = False
                Grid1ExpandedView = False
                SortGridColums()
                CustomDrawCell(GridControlStockGrid, GridViewStockNumbers)
                GridViewStockNumbers.LeftCoord = 0
        End Select
    End Sub
    Private Sub WindowsUIButtonPanelItemEdit_ButtonCheck(sender As Object, e As DevExpress.XtraBars.Docking2010.ButtonEventArgs) Handles WindowsUIButtonPanelItemEdit.ButtonChecked
        Dim tag As String = CType(e.Button, WindowsUIButton).Tag.ToString()
        Select Case tag
            Case "AddDispo"

                GridViewStockNumbers.Columns(10).Visible = False

                GridViewStockNumbers.Columns(5).Visible = True

                GridViewStockNumbers.Columns(6).Visible = True

                GridViewStockNumbers.Columns(7).Visible = True

                GridViewStockNumbers.Columns(8).Visible = True

                GridViewStockNumbers.Columns(9).Visible = True

                GridViewStockNumbers.Columns(10).Visible = True

                GridViewStockNumbers.Columns(11).Visible = True

                Grid1ExpandedView = True
                SortGridColums()
                CustomDrawCell(GridControlStockGrid, GridViewStockNumbers)
                Dim Column As DevExpress.XtraGrid.Columns.GridColumn = GridViewStockNumbers.Columns(5)
                Dim info As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridViewInfo = GridViewStockNumbers.GetViewInfo()
                GridViewStockNumbers.LeftCoord = info.GetColumnLeftCoord(Column) + GridViewStockNumbers.Columns(0).Width
        End Select
    End Sub

    Private Sub WindowsUIButtonPanelBPActions_ButtonClick(sender As Object, e As DevExpress.XtraBars.Docking2010.ButtonEventArgs) Handles WindowsUIButtonPanelItemEdit.ButtonClick
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
    Private Sub RepositoryItemComboBoxSOCIRent_QueryPopUp(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles RepositoryItemComboBoxSOCIRent.QueryPopUp
        Dim lookUpEdit As LookUpEdit = TryCast(sender, LookUpEdit)
        lookUpEdit.Properties.PopulateColumns()
        lookUpEdit.Properties.Columns(0).Visible = False

    End Sub
    Private Sub RepositoryItemComboBoxSOCIStocktype_QueryPopUp(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles RepositoryItemLookUpEditSOCIStockType.QueryPopUp
        Dim lookUpEdit As LookUpEdit = TryCast(sender, LookUpEdit)
        lookUpEdit.Properties.PopulateColumns()
        lookUpEdit.Properties.Columns(0).Visible = False
        lookUpEdit.Properties.Columns(1).Visible = False
    End Sub
#End Region
#Region "Data events"

    Private Sub WriteStockToBP()

        'AbovoBP.WriteStock()
        BIsDirty = False

    End Sub
    Private Sub WriteStockToBPAndSave()

        'AbovoBP.WriteStock()
        BIsDirty = False
        'AbovoBP.SaveBP()

    End Sub

    Private Sub UnboundSourceStocks_ValueNeeded(sender As Object, e As DevExpress.Data.UnboundSourceValueNeededEventArgs) Handles UnboundSourceStocks.ValueNeeded

        e.Value = GetArrayData(e.RowIndex, e.PropertyName)

    End Sub

    Private Sub UnboundSourceStocks_ValuePushed(sender As Object, e As DevExpress.Data.UnboundSourceValuePushedEventArgs) Handles UnboundSourceStocks.ValuePushed
        BIsDirty = True
        SetArrayData(e.RowIndex, e.PropertyName, e.Value)

    End Sub
    Private Function GetArrayData(ByVal rowIndex As Integer, ByVal propertyName As String) As Object

        Dim nully As Int32 = 0

        Select Case propertyName

            Case "PropertyStockDescription"
                Return AbovoBP.Stock.StockItems(rowIndex).StockDescription
            Case "PropertyOwnedManaged"
                Return AbovoBP.Stock.StockItems(rowIndex).OwnedManaged
            Case "PropertySOCIStockType"
                Return AbovoBP.Stock.StockItems(rowIndex).SOCIStockType
            Case "PropertySOCIRentType"
                Return AbovoBP.Stock.StockItems(rowIndex).SOCIRentType
            Case "PropertyCurrentStockNumbers"
                Return AbovoBP.Stock.StockItems(rowIndex).CurrentStockNumbers
            Case "PropertyInitialRateNewLettings"
                Return AbovoBP.Stock.StockItems(rowIndex).NewLetInitialRate
            Case "PropertyNewLettings"
                Return AbovoBP.Stock.StockItems(rowIndex).NewLettings
            Case "PropertyPreBPlanStartDateNewBuild"
                Return AbovoBP.Stock.StockItems(rowIndex).PreBPlanStartDateNewBuild
            Case "PropertyPreBPlanStartDateDemolitions"
                Return AbovoBP.Stock.StockItems(rowIndex).PreBPlanStartDateDemolitions
            Case "PropertyPreBPlanStartDateRTBs"
                Return AbovoBP.Stock.StockItems(rowIndex).PreBPlanStartDateRTBs
            Case "PropertyPreBPlanStartDateOtherDisposals"
                Return AbovoBP.Stock.StockItems(rowIndex).PreBPlanStartDateOtherDisposals
            Case "PropertyExistingStocksCalc"
                Return AbovoBP.Stock.StockItems(rowIndex).ExistingStocksCalc
            Case "PropertyTotalOpeningStockCalc"
                Return AbovoBP.Stock.StockItems(rowIndex).TotalOpeningStockCalc
            Case Else
                Return nully

        End Select

    End Function
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


    Private Sub StockTypeChanged()

    End Sub
    Private Sub GridViewStockNumbers_ValidatingEditor(sender As Object, e As DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventArgs) Handles GridViewStockNumbers.ValidatingEditor

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

    Private Sub GridViewStockNumbers_InvalidValueException(sender As Object, e As InvalidValueExceptionEventArgs) Handles GridViewStockNumbers.InvalidValueException
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
        Dim i As Integer
        Dim np As Integer = 0
        For i = 0 To GridViewStockNumbers.Columns.Count - 1
            If GridViewStockNumbers.Columns(i).Visible Then
                GridViewStockNumbers.Columns(i).VisibleIndex = np
                np += 1
            End If


        Next i
    End Sub
    Private Sub GridViewStockNumbers_ShownEditor(ByVal sender As Object, ByVal e As EventArgs)
        If GridViewStockNumbers.FocusedColumn.FieldName = "PropertySOCIRentType" Then
            Dim lookup As LookUpEdit = TryCast(GridViewStockNumbers.ActiveEditor, LookUpEdit)
            Dim StrCurrentStock As String = GridViewStockNumbers.GetFocusedRowCellValue("PropertySOCIStockType")
            Dim IntFoundCatID As Integer = SOCIStock.GetSOCICategoryByName(StrCurrentStock)
            If IntFoundCatID = 0 Then GridViewStockNumbers.SetFocusedRowCellValue("PropertySOCIRentType", Convert.ToString("N/A"))
            lookup.Properties.DataSource = SOCIRentType.GetSOCIRentTypeByCategory(IntFoundCatID)
        End If


    End Sub

    Private Sub RepositoryItemLookUpEditSOCIStockType_EditValueChanged(sender As Object, e As EventArgs) Handles RepositoryItemLookUpEditSOCIStockType.EditValueChanged
        Dim Editor As LookUpEdit = CType(sender, LookUpEdit)

        Dim StrChosenStock As String = Convert.ToString(Editor.EditValue)
        Dim IntFoundCatID As Integer = SOCIStock.GetSOCICategoryByName(StrChosenStock)
        If IntFoundCatID = 0 Then GridViewStockNumbers.SetFocusedRowCellValue("PropertySOCIRentType", Convert.ToString("N/A"))
    End Sub

    Private Sub GridViewStockNumbers_CustomDrawColumnHeader(sender As Object, e As ColumnHeaderCustomDrawEventArgs) Handles GridViewStockNumbers.CustomDrawColumnHeader
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



    Shared Sub CustomDrawCell(ByVal gridControl As GridControl, ByVal gridView As GridView)
        ' Handle this event to paint cells manually
        'Dim BDo As Boolean = False
        'AddHandler gridView.CustomDrawCell,
        '    Sub(s, e)
        '        If Grid1ExpandedView Then
        '            If e.Column.VisibleIndex = 0 Then
        '                Using pen As New Pen(Color.Silver, 4)
        '                    e.Graphics.DrawLine(pen, e.Bounds.Right, e.Bounds.Top - 4, e.Bounds.Right, e.Bounds.Bottom + 15)
        '                    'e.Graphics.DrawRectangle(pen, e.Bounds)
        '                End Using

        '            End If
        '            If e.Column.VisibleIndex = 12 Then
        '                Using pen As New Pen(Color.Silver, 4)
        '                    e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Top - 4, e.Bounds.Left, e.Bounds.Bottom + 15)
        '                    'e.Graphics.DrawRectangle(pen, e.Bounds)
        '                End Using

        '            End If
        '        Else
        '            If e.Column.VisibleIndex = 0 Then
        '                Using pen As New Pen(Color.White, 4)
        '                    e.Graphics.DrawLine(pen, e.Bounds.Right, e.Bounds.Top, e.Bounds.Right, e.Bounds.Bottom)
        '                    'e.Graphics.DrawRectangle(pen, e.Bounds)
        '                End Using

        '            End If
        '            If e.Column.VisibleIndex = 12 Then
        '                Using pen As New Pen(Color.White, 4)
        '                    e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Top - 4, e.Bounds.Left, e.Bounds.Bottom + 15)
        '                    'e.Graphics.DrawRectangle(pen, e.Bounds)
        '                End Using

        '            End If
        '        End If
        '        If BDo Then
        '            For Each info As DrawElementInfo In e.Cell.InnerElements
        '                If Not info.Visible Then
        '                    Continue For
        '                End If
        '                ObjectPainter.DrawObject(e.Cache, info.ElementPainter, info.ElementInfo)
        '            Next info
        '            e.Handled = True
        '        End If

        '    End Sub
    End Sub



    Private Sub Button1_Click_1(sender As Object, e As EventArgs)


    End Sub
    Sub ResizeFonts()

        Scalefactor = Me.Width / 1700



        Me.GridControlStockGrid.Font = GetFont("Small", Me.Scalefactor)

        Me.GridViewStockNumbers.Appearance.OddRow.Font = GetFont("Small", Me.Scalefactor)
        Me.GridViewStockNumbers.Appearance.ViewCaption.Font = GetFont("Small", Me.Scalefactor)

        Dim x As Integer

        For x = 0 To GridViewStockNumbers.Columns.Count - 1

            GridViewStockNumbers.Columns(x).AppearanceCell.Font = GetFont("Small", Me.Scalefactor)
            GridViewStockNumbers.Columns(x).AppearanceHeader.Font = GetFont("Small", Me.Scalefactor, True)

        Next

        RepositoryItemComboBoxOwnedManaged.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        RepositoryItemLookUpEditSOCIStockType.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        RepositoryItemLookUpEditSOCIRentType.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        RepositoryItemIntegerEdit.Appearance.Font = GetFont("Small", Me.Scalefactor, True)
        RepositoryItemComboBoxSOCIStockType.Appearance.Font = GetFont("Small", Me.Scalefactor, True)


        colPropertyStockDescription1.Width = GridControlStockGrid.Width * 0.25
        colPropertyOwnedManaged1.Width = GridControlStockGrid.Width * 0.125
        colPropertySOCIStockType1.Width = GridControlStockGrid.Width * 0.125
        colPropertySOCIRentType1.Width = GridControlStockGrid.Width * 0.125
        colPropertyCurrentStockNumbers1.Width = GridControlStockGrid.Width * 0.12
        colPropertyNewLettings1.Width = GridControlStockGrid.Width * 0.12
        colPropertyTotalOpeningStockCalc1.Width = GridControlStockGrid.Width * 0.12


        colPropertyPreBPlanStartDateNewBuild1.Width = GridControlStockGrid.Width * 0.1
        colPropertyPreBPlanStartDateDemolitions1.Width = GridControlStockGrid.Width * 0.1
        colPropertyPreBPlanStartDateRTBs1.Width = GridControlStockGrid.Width * 0.1
        colPropertyPreBPlanStartDateOtherDisposals1.Width = GridControlStockGrid.Width * 0.1
        colPropertyExistingStocksCalc1.Width = GridControlStockGrid.Width * 0.1

        colPropertyNewLettings1.Width = GridControlStockGrid.Width * 0.1


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

    End Sub
    Sub ResizeControls()

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
    Private Sub StockAssumptionsInterface_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        ResizeFonts()
        ResizeControls()
    End Sub

End Class


Public Class SOCIStock
    <Display(Order:=-1)>
    Public Property SOCIStockID() As Integer
    Public Property CategoryID() As Integer
    Public Property SOCIStockName() As String
    Public Shared Function Init() As List(Of SOCIStock)
        Return New List(Of SOCIStock)() From {
                New SOCIStock() With {.SOCIStockID = 0, .CategoryID = 1, .SOCIStockName = "Gen Needs"},
                New SOCIStock() With {.SOCIStockID = 1, .CategoryID = 0, .SOCIStockName = "LCHO"},
                New SOCIStock() With {.SOCIStockID = 2, .CategoryID = 1, .SOCIStockName = "Supported"},
                New SOCIStock() With {.SOCIStockID = 3, .CategoryID = 0, .SOCIStockName = "Other"},
                New SOCIStock() With {.SOCIStockID = 4, .CategoryID = 0, .SOCIStockName = "N/A"},
                New SOCIStock() With {.SOCIStockID = 5, .CategoryID = 0, .SOCIStockName = "Supported"},
                New SOCIStock() With {.SOCIStockID = 5, .CategoryID = 0, .SOCIStockName = "Non-social"}
            }
    End Function
    Public Shared Function GetSOCICategoryByName(ByVal MyStockName As String) As Integer
        Dim selectedValue As SOCIStock
        selectedValue = Init.Find(Function(p) p.SOCIStockName = MyStockName)
        Return selectedValue.CategoryID
    End Function

End Class
Public Class SOCIRentType
    <Display(Order:=-1)>
    Public Property SOCIRentTypetID() As Integer
    Public Property SOCIRentName() As String
    <Display(Order:=-1)>
    Public Property CategoryID() As Integer


    Public Shared Function Init() As List(Of SOCIRentType)
        Return New List(Of SOCIRentType)() From {
                New SOCIRentType() With {.SOCIRentTypetID = 0, .SOCIRentName = "N/A", .CategoryID = 0},
                New SOCIRentType() With {.SOCIRentTypetID = 1, .SOCIRentName = "Social Rent", .CategoryID = 1},
                New SOCIRentType() With {.SOCIRentTypetID = 2, .SOCIRentName = "Aff Rent", .CategoryID = 1}
            }
    End Function
    Public Shared Function GetSOCIRentTypeByCategory(ByVal categoryId As Integer) As List(Of SOCIRentType)
        Return Init().Where(Function(p) p.CategoryID = categoryId).ToList()
    End Function
End Class
Public Class OMType
    Public Property OMName() As String
    Public Shared Function Init() As List(Of OMType)
        Return New List(Of OMType)() From {
             New OMType() With {.OMName = "Owned"},
             New OMType() With {.OMName = "Managed"}
           }
    End Function
End Class

