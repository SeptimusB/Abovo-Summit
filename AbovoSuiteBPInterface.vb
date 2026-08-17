Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.IO
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports DevExpress.XtraEditors
Imports DevExpress.Xpo
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraGrid
Imports DevExpress.XtraEditors.Repository
Imports System.Globalization
Imports DevExpress.Utils
Imports DevExpress.Pdf.Native.BouncyCastle.Asn1.X509
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Columns
Imports Abovo

Public Class AbovoSuiteBPInterface

    Inherits DevExpress.XtraEditors.XtraForm

    Public AbovoBusinessPlan As Abovo.AbovoBP
    Public AbvApp As AbovoAppCls
    Private internalAppTitle As String

    Private internalCurrFile As String
    Public Property CurrFile As String
        Get
            Return internalCurrFile
            Exit Property
        End Get
        Set(ByVal NewCurFile As String)
            internalCurrFile = NewCurFile
        End Set
    End Property

    Public Sub New(fromApplication As AbovoAppCls)

        AbvApp = fromApplication
        InitializeComponent()


        AbovoBusinessPlan = New Abovo.AbovoBP
        Me.Text = AbvApp.AppTitle & " " & AbvApp.decVersionNumber & " - No file"

    End Sub


    Private Sub butImportbpFile(ByVal sender As Object, ByVal e As EventArgs)

    End Sub



    Private Sub ImportbpFile()

        Debug.Print("Starting perf import")

        If OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            ProgressPanel("Loading..", "Abovo BP", 0)
            AbovoBusinessPlan.LoadBP(OpenFileDialog1.FileName)
            ProgressPanel("Populating..", "Abovo BP", 1)
            PopulateControls()
            ProgressPanel("Finished..", "Abovo BP", 2)
        Else

            Exit Sub

        End If

    End Sub



    Sub PopulateControls()

        Dim i As Integer




        GridControlSOCIData.DataSource = AbovoBusinessPlan.DSSOCIDataRange

        GridViewSOCIData.Columns(0).Caption = "SOCI Heading"

        GridViewSOCIData.Columns(1).Visible = False
        GridViewSOCIData.Columns(2).Visible = False

        GridViewSOCIData.Columns(6).Visible = False
        GridViewSOCIData.Columns(7).Visible = False
        GridViewSOCIData.Columns(8).Visible = False
        GridViewSOCIData.Columns(9).Visible = False

        GridViewSOCIData.Columns(0).Width = 160
        GridViewSOCIData.Columns(3).Width = 150
        GridViewSOCIData.Columns(4).Width = 150

        GridViewSOCIData.Columns(5).Width = 125
        GridViewSOCIData.Columns(10).Width = 140


        Dim strTemp As String

        For i = 11 To 50
            GridViewSOCIData.Columns(i).DisplayFormat.FormatType = FormatType.Numeric
            'GridViewSOCIData.Columns(i).DisplayFormat.FormatString = "c0"
            strTemp = "Year " & CStr(i - 10)
            GridViewSOCIData.Columns(i).Caption = strTemp
            GridViewSOCIData.Columns(i).Width = 100
            'GridViewSOCIData.Columns(i).DisplayText = CInt(Math.Truncate(GridViewSOCIData.Columns(i).Value)).ToString()
        Next i

        GridViewSOCIData.OptionsBehavior.AlignGroupSummaryInGroupRow = DevExpress.Utils.DefaultBoolean.True
        GridViewSOCIData.OptionsView.GroupFooterShowMode = GroupFooterShowMode.VisibleAlways

        Dim itemCust As New GridGroupSummaryItem
        itemCust.FieldName = "Description"
        itemCust.SummaryType = DevExpress.Data.SummaryItemType.Custom
        itemCust.DisplayFormat = "Total:"
        itemCust.ShowInGroupColumnFooter = GridViewSOCIData.Columns(10)
        GridViewSOCIData.GroupSummary.Add(itemCust)

        Dim item As GridGroupSummaryItem

        For i = 11 To 50

            item = New GridGroupSummaryItem
            item.FieldName = GridViewSOCIData.Columns(i).FieldName
            item.SummaryType = DevExpress.Data.SummaryItemType.Sum
            item.DisplayFormat = "{0:c0}"
            item.ShowInGroupColumnFooter = GridViewSOCIData.Columns(i)

            GridViewSOCIData.GroupSummary.Add(item)

        Next

        'Dim columnTotal As GridColumn = GridViewSOCIData.Columns(51)
        'columnTotal.FilterInfo = New ColumnFilterInfo("[Total] > 0")

        GridViewSOCIData.ActiveFilter.NonColumnFilter = "[Total] > 0"
        GridViewSOCIData.Columns(51).Visible = False
        GridViewSOCIData.OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never

    End Sub

    Private Sub GridViewSOCIData_CustomColumnDisplayText(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs) Handles GridViewSOCIData.CustomColumnDisplayText

        Dim ciGB As CultureInfo = New CultureInfo("en-GB")

        If IsNumeric(e.Column.FieldName) Then
            e.DisplayText = String.Format(ciGB, "{0:c0}", e.Value)
        End If

    End Sub
    Private Sub PerformImport()

        Dim strFileToOpen As String
        Debug.Print("Starting perf import")

        If OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then

            strFileToOpen = OpenFileDialog1.FileName

        Else

            Exit Sub

        End If

        ProgressPanel("Opening file", "Data import", 0)

        Dim wb As New Workbook()

        wb.LoadDocument(strFileToOpen)

        Dim customFunction As New Abovo.PMCostFunction()

        If Not wb.Functions.GlobalCustomFunctions.Contains(customFunction.Name) Then

            wb.Functions.GlobalCustomFunctions.Add(customFunction)

        End If

        Dim customFunction2 As New Abovo.ResponsiveCostFunction()

        If Not wb.Functions.GlobalCustomFunctions.Contains(customFunction2.Name) Then

            wb.Functions.GlobalCustomFunctions.Add(customFunction2)

        End If

        Dim options As New DevExpress.Spreadsheet.RangeDataSourceOptions()

        options.UseFirstRowAsHeader = False
        options.PreserveFormulas = True
        options.SkipHiddenRows = True

        Dim stream As New MemoryStream
        wb.SaveDocument(stream, DevExpress.Spreadsheet.DocumentFormat.Xlsm)
        stream.Position = 0

        ProgressPanel("Calculating", "Updating")

        'dbResult = Abovo.AbovoMCM.ExecProcessImportedFile(Session1)

        ProgressPanel("Refreshing data", "Updating")

        ProgressPanel("Refreshing data", "Finished", 2)

    End Sub

    Private Sub ProgressPanel(strDisplayText As String, strCaption As String, Optional ByVal intStage As Integer = 1)

        Select Case intStage

            Case 0

                SplashScreenManager1.ShowWaitForm()
                SplashScreenManager1.SetWaitFormCaption(strCaption)
                SplashScreenManager1.SetWaitFormDescription(strDisplayText)

            Case 1

                SplashScreenManager1.SetWaitFormDescription(strDisplayText)

            Case Else

                SplashScreenManager1.SetWaitFormCaption("Finished")
                SplashScreenManager1.SetWaitFormDescription(strDisplayText)
                Threading.Thread.Sleep(500)
                SplashScreenManager1.CloseWaitForm()

        End Select

    End Sub










End Class
