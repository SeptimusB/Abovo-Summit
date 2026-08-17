Imports Abovo
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.Skins
Imports DevExpress.LookAndFeel
Imports DevExpress.Utils.Drawing
Imports DevExpress.CodeParser

Public Class InitialRateVaritationsInterface

    Private Shared BIsDirty As Boolean
    Private Shared Grid1ExpandedView As Boolean
    Public Sub New()

        If Not AbovoBP.IRVs.IsInitialised Then AbovoBP.IRVs.Initialise()

        If Not AbovoBP.IRVs.IsError Then
            MsgBox("Error") 'Dim Msg As New AbovoMessageBox("Sorry, Can't display the data due to error : ", 1, Me.ParentForm, "Assumptions error")
            Exit Sub
        End If

        ' This call is required by the designer.
        InitializeComponent()

        'CustomDrawColumnHeader(GridControlStockGrid, GridViewStockNumbers)
        UnboundSourceIRVs.SetRowCount(AbovoBP.StockSize)
        UnboundSourceIRVYears.SetRowCount(1)

        GridViewStockNumbers.OptionsView.EnableAppearanceEvenRow = True
        GridViewStockNumbers.OptionsView.EnableAppearanceOddRow = True


        Me.GridViewStockNumbers.Appearance.HorzLine.BackColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.HorzLine.BorderColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.HorzLine.ForeColor = System.Drawing.Color.White
        Me.GridViewStockNumbers.Appearance.HorzLine.Options.UseBackColor = True
        Me.GridViewStockNumbers.Appearance.HorzLine.Options.UseBorderColor = True
        Me.GridViewStockNumbers.Appearance.HorzLine.Options.UseForeColor = True

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

        GridControlIRVItems.ForceInitialize()

        AddHandler GridViewStockNumbers.ShownEditor, AddressOf GridViewStockNumbers_ShownEditor

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

        GridViewStockNumbers.Columns(5).Visible = False
        GridViewStockNumbers.Columns(6).Visible = False
        GridViewStockNumbers.Columns(7).Visible = False
        GridViewStockNumbers.Columns(8).Visible = False
        GridViewStockNumbers.Columns(9).Visible = False
        GridViewStockNumbers.Columns(11).Visible = False

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

        CustomDrawCell(GridControlIRVItems, GridViewStockNumbers)

        'CustomDrawColumnHeader(GridControlStockGrid, GridViewStockNumbers)
    End Sub
    Sub ResizeControls()

        GridControlIRVYears.Width = GridControlIRVItems.Width / 2

    End Sub


#Region "Interface events"
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
                CustomDrawCell(GridControlIRVItems, GridViewStockNumbers)
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
                CustomDrawCell(GridControlIRVItems, GridViewStockNumbers)
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


    End Sub
    Private Sub WriteStockToBPAndSave()



    End Sub

    Private Sub UnboundSourceStocks_ValueNeeded(sender As Object, e As DevExpress.Data.UnboundSourceValueNeededEventArgs) Handles UnboundSourceIRVs.ValueNeeded

        e.Value = GetArrayData(e.RowIndex, e.PropertyName)

    End Sub

    Private Sub UnboundSourceStocks_ValuePushed(sender As Object, e As DevExpress.Data.UnboundSourceValuePushedEventArgs) Handles UnboundSourceIRVs.ValuePushed
        BIsDirty = True
        SetArrayData(e.RowIndex, e.PropertyName, e.Value)

    End Sub
    Private Function GetArrayData(ByVal rowIndex As Integer, ByVal propertyName As String) As Object






        Return AbovoBP.Stock.InitialRateVariations.GetValue(rowIndex, propertyName)



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



End Class





