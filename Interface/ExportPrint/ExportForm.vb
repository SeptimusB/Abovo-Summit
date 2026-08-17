Imports Abovo
Imports Abovo.AbovoAppCls
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraGrid.Views.Grid
Imports Abovo.LogDebugDev
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.XtraLayout.Customization.Templates
Imports Abovo.CustomGrid
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraSpreadsheet
Imports System.IO
Imports System.Linq
Imports DevExpress.CodeParser
Imports DevExpress.Utils



Public Class ExportForm

    Inherits DevExpress.XtraEditors.XtraForm

    Public ExportPackageCount As Integer = 0
    Public ExportPackagesIndex As Integer = -1
    Public ExportPackages(-1) As GridExportPackage
    Public ExWorkbook As IWorkbook
    Private ScaleUnits As Integer
    Private ExportMode As String
    Public Sub SetMode(ExMode As String)

        ExportMode = ExMode

    End Sub

    Public Sub Initialise()


    End Sub
    Public Sub New()

        InitializeComponent()



        Me.Width = Screen.PrimaryScreen.Bounds.Width * 0.8
        Me.Height = Screen.PrimaryScreen.Bounds.Height * 0.8

        AddHandler Me.WindowsUIButtonPanelClose.ButtonClick, AddressOf WindowsUIButtonPanelActions_ButtonClick
        AddHandler Me.WindowsUIButtonPanelPreview.ButtonClick, AddressOf WindowsUIButtonPanelActions_ButtonClick
        AddHandler Me.WindowsUIButtonPanelSave.ButtonClick, AddressOf WindowsUIButtonPanelActions_ButtonClick

        ScaleUnits = Me.Width * 0.007

        Me.WebBrowserMessage.DocumentText = "<html><body><B>" +
                                               "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits * 1.2) & "px'>Export Data <br/></b>" +
                                                "</p>" +
                                                "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits) & "px'>
                                                Select the elements you wish to export from the below box.&nbsp; Further down are options to exclude or include certain features.&nbsp; You can rename the sheet like in Excel if you want to explore different options.
                                                </body>" +
                                               "</p><p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits) & "px'><a href='https://www.abovo-consult.co.uk/help/exportdata.html' target='new'>Export Help</a><br>" +
                                                "</body></html>"

    End Sub
    Public Sub ClearExportPackages()

        ExportPackagesIndex = -1
        ExportPackageCount = 0
        ReDim ExportPackages(-1)
        CheckedListBoxThingsToExport.Items.Clear()

    End Sub
    Public Sub ProcessAdditions()

        If ExportPackageCount > 0 Then

            For i = 0 To ExportPackageCount - 1

                CheckedListBoxThingsToExport.Items.Add(ExportPackages(i).Description)
                CheckedListBoxThingsToExport.Items(i).CheckState = IIf(ExportPackages(i).IsDefault, CheckState.Checked, CheckState.Unchecked)

            Next


        End If

    End Sub
    Public Sub ProcessExcelExports()



        For Each item In CheckedListBoxThingsToExport.Items

            Dim Index As Integer = CheckedListBoxThingsToExport.Items.IndexOf(item)
            If CheckedListBoxThingsToExport.GetItemCheckState(Index) = CheckState.Checked Then

                If ExportMode = "XLSAnalysis" Then ExportAnalysisToExcel(ExportPackages(Index))

            End If
        Next
        'Dim Exporter As New ExcelDataExport
        'Dim WsOptions As New ExcelDataExport.ExcelExportAdditions

    End Sub
    Sub ExportAnalysisToExcel(ExportPackage As GridExportPackage)

        Dim NumberFormatString As String = "#,###;[red](#,###);0"
        Dim ExportedGrid As GridView = ExportPackage.GridView
        Dim ExportedGridName As String = ExportPackage.Description
        Dim RowCount As Integer = ExportedGrid.RowCount
        Dim CurrWBRow As Integer = 0
        Dim DoSubTotals As Boolean = True
        Dim DoFormats As Boolean = True
        Dim OpenExcel As Boolean = False

        If CheckedListBoxElementsToExport.Items(0).CheckState = CheckState.Unchecked Then DoSubTotals = False
        If CheckedListBoxElementsToExport.Items(1).CheckState = CheckState.Checked Then DoFormats = False


        SpreadsheetControlExport.Enabled = True

        ExWorkbook = SpreadsheetControlExport.Document

        ExWorkbook.BeginUpdate()

        Dim NewSheetName As String = ExportPackage.Description & " Export"

        If ExWorkbook.Worksheets.Contains(NewSheetName) Then

            ExWorkbook.Worksheets.Remove(ExWorkbook.Worksheets(NewSheetName))

        End If

        ExWorkbook.Worksheets.Insert(0, NewSheetName)

        Dim ws As DevExpress.Spreadsheet.Worksheet = ExWorkbook.Worksheets(0)


        Dim i As Integer = -1
        Dim j As Integer = 0
        Dim Row As Integer = i
        Dim Col As Integer = 0

        Dim Cell As Cell

        Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 0)
        Cell.Borders.BottomBorder.LineStyle = BorderLineStyle.Thin
        Cell.SetValueFromText("Section / Heading")
        Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 1)
        Cell.Borders.BottomBorder.LineStyle = BorderLineStyle.Thin
        Cell.SetValueFromText("Subheading")
        Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 2)
        Cell.Borders.BottomBorder.LineStyle = BorderLineStyle.Thin
        Cell.SetValueFromText("Description")

        For i = 3 To (ExportPackage.IDCount + 3)

            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, i)
            Cell.Borders.BottomBorder.LineStyle = BorderLineStyle.Thin
            Cell.SetValueFromText(ExportedGrid.Columns(ExportPackage.IDStart + i - 3).Caption)

        Next


        Dim LastGroupA As String = ""
        Dim LastGroupB As String = ""
        Dim LastGroupC As String = ""
        Dim LastGroupD As String = ""
        Dim TempString As String = ""
        Dim StartPos As Integer = 3

        Dim CurrFillColour As Color

        Dim SumItems As ArrayList = ExtractSummaryItems(ExportedGrid)

        For i = -1 To -100000 Step -1

            If Not ExportedGrid.IsValidRowHandle(i) Then Exit For

            If ExportedGrid.IsGroupRow(i) Then

                Dim ColPos As Integer = StartPos

                Dim GroupValues As Hashtable = ExportedGrid.GetGroupSummaryValues(i)

                If ExportedGrid.GetRowLevel(i) = 0 Then

                    If Convert.ToInt32(GroupValues(ExportedGrid.GroupSummary(0))) > 0 Then ' it is a title row

                        CurrWBRow += 1
                        Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 0)

                        If Convert.ToInt32(GroupValues(ExportedGrid.GroupSummary(0))) = 1 Then

                            CurrFillColour = Color.LightSteelBlue
                            Cell.Alignment.Indent = 0

                        Else

                            CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(250, Byte), Integer))
                            Cell.Alignment.Indent = 1

                        End If

                        TempString = ExportedGrid.GetGroupRowDisplayText(i)
                        TempString = ProcessHeader(TempString)
                        If DoFormats Then Cell.FillColor = CurrFillColour
                        Cell.SetValueFromText(TempString)
                        Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                        LastGroupA = TempString
                        LastGroupB = ""
                        LastGroupC = ""
                        LastGroupD = ""
                        For z = 1 To 3

                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, z)
                            If DoFormats Then Cell.FillColor = CurrFillColour

                        Next
                        For Each item As GridGroupSummaryItem In SumItems

                            If Not item.FieldName = "TitleLevel" And Not item.FieldName = "AmDummy" Then

                                Dim text As String = item.GetDisplayText(GroupValues(item), False)

                                If IsNumeric(text) Then

                                    Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, ColPos)
                                    Cell.NumberFormat = NumberFormatString
                                    Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                    Cell.SetValueFromText(text)
                                    If DoFormats Then Cell.FillColor = CurrFillColour
                                    ColPos += 1

                                End If

                            End If

                        Next item

                    Else 'Not a title row

                        CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(235, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(250, Byte), Integer))

                        TempString = ExportedGrid.GetGroupRowDisplayText(i)

                        TempString = ProcessHeader(TempString)

                        If LastGroupA <> TempString Then

                            CurrWBRow += 1
                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 0)
                            Cell.SetValueFromText(TempString)
                            If DoFormats Then Cell.FillColor = CurrFillColour
                            Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                            Cell.Alignment.Indent = 2

                            LastGroupA = TempString
                            LastGroupB = ""
                            LastGroupC = ""
                            LastGroupD = ""

                            For z = 1 To 42

                                Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, z)
                                If DoFormats Then Cell.FillColor = CurrFillColour

                            Next

                        End If

                        If ExportedGrid.GetRowExpanded(i) Then ' The Level 0 Row is expanded

                            Dim CRC As Integer
                            Dim ChildRowHandle As Integer
                            CRC = ExportedGrid.GetChildRowCount(i)
                            Dim x As Integer

                            For x = 0 To CRC - 1

                                CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(250, Byte), Integer))
                                ChildRowHandle = ExportedGrid.GetChildRowHandle(i, x)

                                If ExportedGrid.IsGroupRow(ChildRowHandle) Then

                                    If Not ExportedGrid.GetRowExpanded(ChildRowHandle) Then 'The Level1 Row is not expanded, paint totals 


                                        TempString = ExportedGrid.GetGroupRowDisplayText(ChildRowHandle)
                                        TempString = ProcessHeader(TempString)


                                        CurrWBRow += 1
                                        Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 0)
                                        Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                        Cell.SetValueFromText(TempString)
                                        If DoFormats Then Cell.FillColor = CurrFillColour
                                        Cell.Alignment.Indent = 3

                                        For z = 1 To 3

                                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, z)
                                            If DoFormats Then Cell.FillColor = CurrFillColour

                                        Next

                                        ColPos = StartPos

                                        Dim Child1Values As Hashtable = ExportedGrid.GetGroupSummaryValues(ChildRowHandle)

                                        For Each item As GridGroupSummaryItem In SumItems

                                            If Not item.FieldName = "TitleLevel" And Not item.FieldName = "AmDummy" Then

                                                Dim text As String = item.GetDisplayText(Child1Values(item), False)

                                                If IsNumeric(text) Then

                                                    Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, ColPos)
                                                    Cell.NumberFormat = NumberFormatString
                                                    Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                                    If DoFormats Then Cell.FillColor = CurrFillColour
                                                    Cell.SetValueFromText(ProcessHeader(text))
                                                    ColPos += 1

                                                End If

                                            End If

                                        Next item

                                    Else 'The Level1 Row is expanded, iterate

                                        'Paint the header row

                                        TempString = ProcessHeader(ExportedGrid.GetGroupRowDisplayText(ChildRowHandle))

                                        CurrWBRow += 1
                                        Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 0)
                                        Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                        Cell.SetValueFromText(TempString)
                                        LastGroupB = TempString
                                        If DoFormats Then Cell.FillColor = CurrFillColour
                                        Cell.Alignment.Indent = 3

                                        For z = 1 To 42

                                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, z)
                                            If DoFormats Then Cell.FillColor = CurrFillColour

                                        Next

                                        CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(250, Byte), Integer))
                                        Dim ChildChildRowHandle As Integer
                                        Dim CRCrowC As Integer = ExportedGrid.GetChildRowCount(ChildRowHandle)
                                        Dim y As Integer

                                        For y = 0 To CRCrowC - 1

                                            ';'CurrWBRow += 1

                                            ChildChildRowHandle = ExportedGrid.GetChildRowHandle(ChildRowHandle, y)

                                            If ExportedGrid.GetRowExpanded(ChildChildRowHandle) Then 'The Level 2 Row is expanded, iterate

                                                TempString = ProcessHeader(ExportedGrid.GetGroupRowDisplayText(ChildChildRowHandle))
                                                LastGroupC = TempString
                                                CurrWBRow += 1
                                                Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 0)
                                                Cell.SetValueFromText(TempString)
                                                Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                                If DoFormats Then Cell.FillColor = CurrFillColour
                                                Cell.Alignment.Indent = 4

                                                For z = 1 To 42

                                                    Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, z)
                                                    If DoFormats Then Cell.FillColor = CurrFillColour

                                                Next

                                                CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(250, Byte), Integer), CType(CType(255, Byte), Integer))

                                                Dim ChildChildChildRowHandle As Integer
                                                Dim CRChildCrowC As Integer = ExportedGrid.GetChildRowCount(ChildChildRowHandle)
                                                Dim a As Integer

                                                'CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))

                                                For a = 0 To CRChildCrowC - 1

                                                    CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(254, Byte), Integer), CType(CType(254, Byte), Integer), CType(CType(255, Byte), Integer))


                                                    ChildChildChildRowHandle = ExportedGrid.GetChildRowHandle(ChildChildRowHandle, a)
                                                    TempString = ExportedGrid.GetGroupRowDisplayText(ChildChildChildRowHandle)
                                                    TempString = ProcessHeader(TempString)
                                                    LastGroupD = TempString
                                                    CurrWBRow += 1

                                                    Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 0)
                                                    Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                                    Cell.SetValueFromText(TempString)
                                                    If DoFormats Then Cell.FillColor = CurrFillColour
                                                    Cell.Alignment.Indent = 5

                                                    If ExportedGrid.GetRowExpanded(ChildChildChildRowHandle) Then 'The Level 3 Row is expanded, iterate

                                                        For z = 1 To 39 + StartPos

                                                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, z)
                                                            If DoFormats Then Cell.FillColor = CurrFillColour

                                                        Next z

                                                        'DataLines
                                                        CurrFillColour = Color.White

                                                        Dim DataRowHandle As Integer
                                                        Dim dataRowCount As Integer = ExportedGrid.GetChildRowCount(ChildChildChildRowHandle)
                                                        Dim b As Integer


                                                        For b = 0 To dataRowCount - 1

                                                            DataRowHandle = ExportedGrid.GetChildRowHandle(ChildChildChildRowHandle, b)

                                                            CurrWBRow += 1
                                                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 0)
                                                            If DoFormats Then Cell.FillColor = CurrFillColour
                                                            Cell.Alignment.Indent = 6
                                                            Cell.SetValueFromText(ExportedGrid.GetRowCellValue(DataRowHandle, ExportedGrid.Columns(ExportPackage.GroupC)))
                                                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 1)
                                                            If DoFormats Then Cell.FillColor = CurrFillColour
                                                            Cell.SetValueFromText(ExportedGrid.GetRowCellValue(DataRowHandle, ExportedGrid.Columns(13)))
                                                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 2)
                                                            If DoFormats Then Cell.FillColor = CurrFillColour
                                                            Cell.SetValueFromText(ExportedGrid.GetRowCellValue(DataRowHandle, ExportedGrid.Columns(14)))

                                                            For z = StartPos To 39 + StartPos

                                                                Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, z)
                                                                If DoFormats Then Cell.FillColor = CurrFillColour
                                                                Cell.NumberFormat = NumberFormatString
                                                                Cell.SetValueFromText(ExportedGrid.GetRowCellValue(DataRowHandle, ExportedGrid.Columns(ExportPackage.IDStart + z - 3)))

                                                            Next z

                                                        Next b

                                                        If DoSubTotals Then

                                                            CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(254, Byte), Integer), CType(CType(254, Byte), Integer), CType(CType(255, Byte), Integer))

                                                            CurrWBRow += 1
                                                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 0)
                                                            Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                                            Cell.SetValueFromText(LastGroupD & "Total")
                                                            If DoFormats Then Cell.FillColor = CurrFillColour
                                                            Cell.Alignment.Indent = 5

                                                            ColPos = StartPos

                                                            Dim SubTotalValues As Hashtable = ExportedGrid.GetGroupSummaryValues(ChildChildChildRowHandle)

                                                            For Each item As GridGroupSummaryItem In SumItems

                                                                For z = 1 To 3

                                                                    Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, z)
                                                                    If DoFormats Then Cell.FillColor = CurrFillColour

                                                                Next

                                                                If Not item.FieldName = "TitleLevel" And Not item.FieldName = "AmDummy" Then

                                                                    Dim text As String = item.GetDisplayText(SubTotalValues(item), False)

                                                                    If IsNumeric(text) Then

                                                                        Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, ColPos)
                                                                        Cell.NumberFormat = NumberFormatString
                                                                        Cell.Borders.TopBorder.LineStyle = BorderLineStyle.Hair
                                                                        Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                                                        If DoFormats Then Cell.FillColor = CurrFillColour
                                                                        Cell.SetValueFromText(ProcessHeader(text))
                                                                        ColPos += 1

                                                                    End If

                                                                End If

                                                            Next item

                                                        End If

                                                    Else


                                                        CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(254, Byte), Integer), CType(CType(254, Byte), Integer), CType(CType(255, Byte), Integer))

                                                        ColPos = StartPos

                                                        For z = 1 To 2

                                                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, z)
                                                            If DoFormats Then Cell.FillColor = CurrFillColour

                                                        Next z

                                                        Dim Child3Values As Hashtable = ExportedGrid.GetGroupSummaryValues(ChildChildChildRowHandle)

                                                        For Each item As GridGroupSummaryItem In SumItems

                                                            If Not item.FieldName = "TitleLevel" And Not item.FieldName = "AmDummy" Then

                                                                Dim text As String = item.GetDisplayText(Child3Values(item), False)

                                                                If IsNumeric(text) Then

                                                                    Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, ColPos)
                                                                    Cell.NumberFormat = NumberFormatString
                                                                    If DoFormats Then Cell.FillColor = CurrFillColour
                                                                    Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                                                    Cell.SetValueFromText(ProcessHeader(text))
                                                                    ColPos += 1

                                                                End If

                                                            End If

                                                        Next item

                                                    End If

                                                Next

                                                If DoSubTotals Then

                                                    CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(250, Byte), Integer), CType(CType(255, Byte), Integer))

                                                    CurrWBRow += 1
                                                    Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 0)
                                                    Cell.SetValueFromText(LastGroupC & "Total")
                                                    Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                                    If DoFormats Then Cell.FillColor = CurrFillColour
                                                    Cell.Alignment.Indent = 4

                                                    ColPos = StartPos

                                                    Dim SubTotalValues As Hashtable = ExportedGrid.GetGroupSummaryValues(ChildChildRowHandle)

                                                    For Each item As GridGroupSummaryItem In SumItems

                                                        For z = 1 To 3

                                                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, z)
                                                            If DoFormats Then Cell.FillColor = CurrFillColour

                                                        Next

                                                        If Not item.FieldName = "TitleLevel" And Not item.FieldName = "AmDummy" Then

                                                            Dim text As String = item.GetDisplayText(SubTotalValues(item), False)

                                                            If IsNumeric(text) Then


                                                                'Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                                                Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, ColPos)
                                                                Cell.Borders.TopBorder.LineStyle = BorderLineStyle.Thin
                                                                Cell.NumberFormat = NumberFormatString
                                                                If DoFormats Then Cell.FillColor = CurrFillColour
                                                                Cell.SetValueFromText(ProcessHeader(text))
                                                                Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                                                ColPos += 1

                                                            End If

                                                        End If

                                                    Next item

                                                End If

                                            Else 'The Level 2 Row is not expanded paint totals

                                                CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(255, Byte), Integer))

                                                CurrWBRow += 1
                                                TempString = ExportedGrid.GetGroupRowDisplayText(ChildChildRowHandle)

                                                Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 0)
                                                Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                                Cell.SetValueFromText(ProcessHeader(TempString))
                                                If DoFormats Then Cell.FillColor = CurrFillColour
                                                Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                                Cell.Alignment.Indent = 4

                                                Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 1)
                                                If DoFormats Then Cell.FillColor = CurrFillColour
                                                Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 2)
                                                If DoFormats Then Cell.FillColor = CurrFillColour
                                                Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 3)
                                                If DoFormats Then Cell.FillColor = CurrFillColour

                                                ColPos = StartPos

                                                Dim Child2Values As Hashtable = ExportedGrid.GetGroupSummaryValues(ChildChildRowHandle)

                                                For Each item As GridGroupSummaryItem In SumItems

                                                    If Not item.FieldName = "TitleLevel" And Not item.FieldName = "AmDummy" Then

                                                        Dim text As String = item.GetDisplayText(Child2Values(item), False)

                                                        If IsNumeric(text) Then

                                                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, ColPos)
                                                            Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                                            Cell.NumberFormat = NumberFormatString
                                                            If DoFormats Then Cell.FillColor = CurrFillColour
                                                            Cell.SetValueFromText(ProcessHeader(text))
                                                            ColPos += 1

                                                        End If

                                                    End If

                                                Next item

                                            End If

                                            'If GridView_Function_CanExpand(ExportedGrid, ChildChildRowHandle, 2) Then ExportedGrid.SetRowExpanded(ChildChildRowHandle, True, True)

                                        Next

                                        If DoSubTotals Then ' paint the level 1 sub total
                                            'CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(235, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(250, Byte), Integer))
                                            CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(250, Byte), Integer))

                                            CurrWBRow += 1
                                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 0)
                                            Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                            Cell.SetValueFromText(LastGroupB & "Total")
                                            If DoFormats Then Cell.FillColor = CurrFillColour
                                            Cell.Alignment.Indent = 3

                                            ColPos = StartPos

                                            Dim SubTotalValues0 As Hashtable = ExportedGrid.GetGroupSummaryValues(ChildRowHandle)

                                            For z = 1 To 3

                                                Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, z)
                                                If DoFormats Then Cell.FillColor = CurrFillColour

                                            Next

                                            For Each item As GridGroupSummaryItem In SumItems



                                                If Not item.FieldName = "TitleLevel" And Not item.FieldName = "AmDummy" Then

                                                    Dim text As String = item.GetDisplayText(SubTotalValues0(item), False)

                                                    If IsNumeric(text) Then

                                                        Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, ColPos)
                                                        Cell.NumberFormat = NumberFormatString
                                                        If DoFormats Then Cell.FillColor = CurrFillColour
                                                        Cell.Borders.TopBorder.LineStyle = BorderLineStyle.Medium
                                                        Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                                        Cell.SetValueFromText(ProcessHeader(text))
                                                        ColPos += 1
                                                        Cell.Borders.TopBorder.LineStyle = BorderLineStyle.Medium

                                                    End If

                                                End If

                                            Next item

                                        End If

                                    End If


                                End If

                            Next

                            If DoSubTotals Then ' paint the level 0 sub total

                                'CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(235, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(250, Byte), Integer))
                                CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(235, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(250, Byte), Integer))

                                CurrWBRow += 1
                                Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, 0)
                                Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                Cell.SetValueFromText(LastGroupA & "Total")
                                If DoFormats Then Cell.FillColor = CurrFillColour
                                Cell.Alignment.Indent = 2

                                ColPos = StartPos

                                Dim SubTotalValues As Hashtable = ExportedGrid.GetGroupSummaryValues(i)

                                For Each item As GridGroupSummaryItem In SumItems

                                    For z = 1 To 3

                                        Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, z)
                                        If DoFormats Then Cell.FillColor = CurrFillColour

                                    Next

                                    If Not item.FieldName = "TitleLevel" And Not item.FieldName = "AmDummy" Then

                                        Dim text As String = item.GetDisplayText(SubTotalValues(item), False)

                                        If IsNumeric(text) Then

                                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, ColPos)
                                            Cell.NumberFormat = NumberFormatString
                                            Cell.Borders.TopBorder.LineStyle = BorderLineStyle.Thick
                                            Cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                                            If DoFormats Then Cell.FillColor = CurrFillColour
                                            Cell.SetValueFromText(ProcessHeader(text))
                                            ColPos += 1

                                        End If

                                    End If

                                Next item

                            End If

                        Else 'Not Expanded, calc the sum totals

                            CurrFillColour = System.Drawing.Color.FromArgb(CType(CType(235, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(250, Byte), Integer))

                            If SumItems.Count > 0 Then

                                ColPos = StartPos

                                For Each item As GridGroupSummaryItem In SumItems

                                    If Not item.FieldName = "TitleLevel" And Not item.FieldName = "AmDummy" Then

                                        Dim text As String = item.GetDisplayText(GroupValues(item), False)

                                        If IsNumeric(text) Then

                                            Cell = ExWorkbook.Worksheets(0).Cells(CurrWBRow, ColPos)
                                            Cell.NumberFormat = NumberFormatString
                                            If DoFormats Then Cell.FillColor = CurrFillColour

                                            Cell.SetValueFromText(text)
                                            ColPos += 1

                                        End If

                                    End If

                                Next item

                            End If 'If SumItems.Count > 0 Then

                        End If ' If ExportedGrid.GetRowExpanded(i) Then

                    End If 'If Convert.ToInt32(GroupValues(ExportedGrid.GroupSummary(0))) > 0 Then ' it is a title row

                End If 'If ExportedGrid.GetRowLevel(i) = 0 Then

            End If 'If ExportedGrid.IsGroupRow(i) Then

        Next

        ExWorkbook.Worksheets(0).Columns.AutoFit(0, ExWorkbook.Worksheets(0).Columns.LastUsedIndex)
        ExWorkbook.EndUpdate()

    End Sub
    Sub ExportToPDF(ExportPackage As GridExportPackage)





    End Sub
    Sub ProcessPDFExports(ExportPackage As GridExportPackage)

        For Each item In CheckedListBoxThingsToExport.Items
            Dim Index As Integer = CheckedListBoxThingsToExport.Items.IndexOf(item)
            If CheckedListBoxThingsToExport.GetItemCheckState(Index) = CheckState.Checked Then

                ExportToPDF(ExportPackages(Index))

            End If
        Next

    End Sub
    Public Sub AddExportPackage(Package As GridExportPackage, Optional ByVal isDefault As Boolean = False)

        Package.IsDefault = isDefault
        ExportPackagesIndex += 1
        ExportPackageCount += 1
        ReDim Preserve ExportPackages(ExportPackagesIndex)
        ExportPackages(ExportPackagesIndex) = Package

    End Sub
    Public Sub SaveNewWorkbook()


        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Excel Files|*.xlsx|All Files|*.*"
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
            SpreadsheetControlExport.SaveDocument(saveFileDialog.FileName, DocumentFormat.Xlsx)
            If CheckedListBoxElementsToExport.Items(2).CheckState Then OpenFileInExcel(saveFileDialog.FileName)

        End If


    End Sub
    Sub SaveExistingWorkbook()

        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Excel Files|*.xlsx;*.xls"
        openFileDialog.Title = "Select an Existing Workbook"

        ' Show the dialog and get the selected file path

        If openFileDialog.ShowDialog() = DialogResult.OK Then

            Dim targetWorkbookPath As String = openFileDialog.FileName
            Dim targetWorkbook As Workbook = New Workbook()
            targetWorkbook.LoadDocument(targetWorkbookPath)

            ' Address the in-memory workbook
            Dim inMemoryWorkbook As IWorkbook = SpreadsheetControlExport.Document



            Dim copyOptions As New WorksheetCopyOptions()
            copyOptions.SheetMappings = inMemoryWorkbook.Sheets.ToDictionary(
                          Function(sheet) sheet.Name, Function(sheet) sheet.Name)

            ' Iterate through the worksheets in the in-memory workbook

            For Each sheet As Worksheet In inMemoryWorkbook.Worksheets

                Dim sheetName As String = sheet.Name

                ' Check if the worksheet already exists in the existing workbook
                If targetWorkbook.Worksheets.Contains(sheetName) Then


                    ' Prompt the user to overwrite
                    Dim result As DialogResult = MessageBox.Show($"Worksheet '{sheetName}' already exists. Do you want to overwrite it?", "Overwrite Worksheet", MessageBoxButtons.YesNo)

                    If result = DialogResult.Yes Then


                        targetWorkbook.Worksheets.Remove(targetWorkbook.Worksheets(sheetName))

                        If Not targetWorkbook.Worksheets.Contains(sheetName) Then

                            targetWorkbook.Worksheets.Add(sheetName)

                        End If

                        targetWorkbook.Worksheets(sheetName).CopyFrom(inMemoryWorkbook.Worksheets(sheetName))

                    End If

                Else

                    targetWorkbook.Worksheets.Add(sheetName)
                    targetWorkbook.Worksheets(sheetName).CopyFrom(inMemoryWorkbook.Worksheets(sheetName))

                End If

            Next

            ' Save the existing workbook
            targetWorkbook.SaveDocument(targetWorkbookPath)

            targetWorkbook.Dispose()
            targetWorkbook = Nothing
            If CheckedListBoxElementsToExport.Items(2).CheckState Then OpenFileInExcel(targetWorkbookPath)

        End If

    End Sub

    Sub ClearWorkbook()

        ExWorkbook = SpreadsheetControlExport.Document

        ExWorkbook.BeginUpdate()

        Dim worksheetsToRemove As New List(Of Worksheet)
        For Each ws As Worksheet In ExWorkbook.Worksheets
            worksheetsToRemove.Add(ws)
        Next

        For Each ws As Worksheet In worksheetsToRemove
            ExWorkbook.Worksheets.Remove(ws)
        Next

        ExWorkbook.EndUpdate()

    End Sub

    Sub OpenFileInExcel(Path As String)

        Try

            Process.Start("EXCEL.EXE", Path)

        Catch ex As Exception

            MessageBox.Show("Error opening Excel: " & ex.Message)

        End Try


    End Sub
    'Sub TryOut()
    '    Using sourceWorkbook As New Workbook()
    '        Using targetWorkbook As New Workbook()
    '            targetWorkbook.LoadDocument("Documents\Book1.xlsx")
    '            sourceWorkbook.LoadDocument("Documents\Book2.xlsx")
    '            Dim copyOptions As New WorksheetCopyOptions()
    '            ' Specify mappings between worksheets in the source
    '            ' and destination workbooks. Sheets are copied 
    '            ' with their original names.
    '            copyOptions.SheetMappings = sourceWorkbook.Sheets.ToDictionary(
    '              Function(sheet) sheet.Name, Function(sheet) sheet.Name)
    '            ' Copy all worksheets from the source workbook
    '            ' to the target workbook.
    '            CopySheetsToWorkbook(sourceWorkbook, targetWorkbook, copyOptions)
    '            targetWorkbook.Calculate()
    '            targetWorkbook.SaveDocument("MergedDocument.xlsx")
    '        End Using
    '    End Using
    'End Sub

    Private Sub WindowsUIButtonPanelActions_ButtonClick(sender As Object, e As DevExpress.XtraBars.Docking2010.ButtonEventArgs)

        Dim ButSender As WindowsUIButton = TryCast(e.Button, DevExpress.XtraBars.Docking2010.WindowsUIButton)
        If ButSender Is Nothing Then
            Return
        End If
        Dim tag As String = ButSender.Tag.ToString()

        Select Case tag

            Case "Clear"

                ClearWorkbook()

            Case "SaveExisting"

                SaveExistingWorkbook()

            Case "SaveNew"

                SaveNewWorkbook()

            Case "Preview"

                If Me.XtraTabControlExport.SelectedTabPage.Name = "XtraTabPageExportXLS" Then

                    ProcessExcelExports()

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
    Private Function ExtractSummaryItems(ByVal view As CustomGridView) As ArrayList
        Dim items As New ArrayList()
        For Each si As GridSummaryItem In view.GroupSummary
            If TypeOf si Is GridGroupSummaryItem AndAlso si.SummaryType <> DevExpress.Data.SummaryItemType.None Then
                items.Add(si)
            End If
        Next si
        Return items
    End Function
    Function GridView_Function_CanExpand(GV As CustomGridView, RowHandle As Integer, RLev As Integer) As Boolean

        Dim Expandable As Boolean = False

        Dim ht As Hashtable = GV.GetGroupSummaryValues(RowHandle)

        If ht IsNot Nothing Then

            If RLev = 0 Then

                If Convert.ToInt32(ht(GV.GroupSummary(0))) = 0 Then

                    Expandable = True

                End If

            ElseIf RLev = 2 Then

                If Convert.ToInt32(ht(GV.GroupSummary(1))) = 0 Then

                    Expandable = True

                End If

            End If

        End If

        Return Expandable

    End Function
    Function ProcessHeader(headerString As String) As String
        If Len(headerString) > 26 Then
            If Microsoft.VisualBasic.Left(headerString, 18) = "Ordered SOCI Group" Or Microsoft.VisualBasic.Left(headerString, 18) = "Ordered CFlo Group" Then

                headerString = Microsoft.VisualBasic.Right(headerString, Len(headerString) - 25)

            End If
        End If
        If Len(headerString) > 20 Then
            If Microsoft.VisualBasic.Left(headerString, 20) = "Ordered SOCI Heading" Or Microsoft.VisualBasic.Left(headerString, 20) = "Ordered CFlo Heading" Then

                headerString = Microsoft.VisualBasic.Right(headerString, Len(headerString) - 27)

            End If
        End If
        If Len(headerString) > 14 Then
            If Microsoft.VisualBasic.Left(headerString, 13) = "Level 1 Copy:" Or Microsoft.VisualBasic.Left(headerString, 18) = "Ordered CFlo Heading" Then

                headerString = Microsoft.VisualBasic.Right(headerString, Len(headerString) - 14)

            End If
        End If
        If Len(headerString) > 14 Then
            If Microsoft.VisualBasic.Left(headerString, 13) = "Level 2 Copy:" Or Microsoft.VisualBasic.Left(headerString, 18) = "Ordered CFlo Heading" Then

                headerString = Microsoft.VisualBasic.Right(headerString, Len(headerString) - 14)

            End If
        End If
        Return headerString
    End Function

    Private Sub TablePanelLeftSidePanel_Paint(sender As Object, e As PaintEventArgs) Handles TablePanelLeftSidePanel.Paint

    End Sub

    Private Sub ExportForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        e.Cancel = True
        Me.Hide()

    End Sub
End Class

