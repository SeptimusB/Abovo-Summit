Imports Abovo
Imports Abovo.FileManager

Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraPrinting.BarCode
Imports DevExpress.XtraSpreadsheet

Public Class MainModelViewer

    Private ModelID As Integer
    Private MyColourSwatch As Color
    Public SSC As SpreadsheetControl
    Public Sub New(SetModelID)

        ModelID = SetModelID

        MyColourSwatch = ExcelModels(ModelID).ColourSwatch
        ' This call is required by the designer.
        InitializeComponent()

        SSC = ExcelModels(ModelID).ModelSpreadsheetControl
        Me.TablePanelSSV.Controls.Add(SSC)
        TablePanelSSV.SetCell(SSC, 1, 0)
        Me.SpreadsheetFormulaBarMMV.SpreadsheetControl = SSC
        SSC.Dock = DockStyle.Fill
        AddHandler SSC.CellValueChanged, AddressOf ProcessCVC
        AddHandler SSC.ActiveSheetChanged, AddressOf RecalcSSC

    End Sub
    Public Sub RecalcSSC()

        SSC.ActiveWorksheet.Calculate()
        ExcelModels(ModelID).WBCalcEngine.CalculateWSs()

    End Sub
    Public Sub ProcessCVC(ByVal sender As Object, ByVal e As SpreadsheetCellEventArgs)

        'Structural and multi-range workbook services publish one transaction-level
        'result after their work completes.  SpreadsheetControl also raises this
        'interactive event for their programmatic range copies; recalculating the
        'entire model for every such event turns one structural command into dozens
        'of redundant full calculations.
        If ModelSafetyManager.IsBulkWorkbookMutationInProgress(ModelID) Then Return

        MasterChangeLog.AddChangeLogEvent(New ChangeLogEvent With {
                .ModelID = ModelID,
                .Description = "Sheet " & SSC.ActiveWorksheet.Name & " cell " & e.Cell.GetReferenceA1() & " changed from " & e.OldValue.ToString & " to " & e.Value.ToString,
                .WSName = SSC.ActiveWorksheet.Name,
                .CellAddress = e.Cell.GetReferenceA1(),
                .OriginalValue = e.OldValue.ToString,
                .ChangedValue = e.Value.ToString,
                .TimeStamp = Now(),
                .UserName = Environment.UserName,
                .Status = 1
            })

        SSC.ActiveWorksheet.Calculate()
        ExcelModels(ModelID).WBCalcEngine.CalculateWSs()

    End Sub
    Public Property FormBorderColor() As Color
        Get
            Return MyColourSwatch
        End Get
        Set(ByVal value As Color)
            MyColourSwatch = value
        End Set
    End Property
    Public Class CustomFormPainterMMV
        Inherits DevExpress.Skins.XtraForm.FormPainter
        Public Sub New(ByVal owner As System.Windows.Forms.Control, ByVal provider As DevExpress.Skins.ISkinProvider)
            MyBase.New(owner, provider)
        End Sub
        Private Function GetFormBorderColor() As Color
            Dim formBorderColor = (TryCast(Owner, MainModelViewer)).FormBorderColor
            Return formBorderColor
        End Function
        Protected Overrides Sub DrawBackground(ByVal cache As DevExpress.Utils.Drawing.GraphicsCache)
            Dim info = GetCaptionInfo()
            Dim ee = TryCast(info, DevExpress.Utils.Drawing.ObjectInfoArgs)
            Dim formBorderColor = GetFormBorderColor()
            cache.FillRectangle(New SolidBrush(formBorderColor), ee.Bounds)
        End Sub
    End Class
    Protected Overrides Function CreateFormBorderPainter() As DevExpress.Skins.XtraForm.FormPainter
        Return New CustomFormPainterMMV(Me, LookAndFeel)
    End Function
    Private Sub MainModelViewer_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        e.Cancel = True
        Me.Hide()

    End Sub

    Private Sub MainModelViewer_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
