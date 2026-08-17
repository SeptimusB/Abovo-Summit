Imports DevExpress.Skins
Imports DevExpress.Skins.XtraForm
Imports DevExpress.Utils.Drawing
Imports Abovo.FileManager

Public Class SpreadsheetViewer

    'Private MyColourSwatch As Color

    Public WithEvents SpreadsheetControlViewer As DevExpress.XtraSpreadsheet.SpreadsheetControl
    Public Sub New()
        'MyColourSwatch = ExcelModels(SetModelID).ColourSwatch
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    'Protected Overrides Function CreateFormBorderPainter() As DevExpress.Skins.XtraForm.FormPainter
    '    Return New CustomFormPainterSSV(Me, LookAndFeel)
    'End Function


    Sub SaveDocument()

        Dim SaveDialog As New SaveFileDialog
        SaveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        SaveDialog.Title = "Save the Excel File"
        SaveDialog.ShowDialog()
        If SaveDialog.FileName <> "" Then
            SpreadsheetControlViewer.Document.SaveDocument(SaveDialog.FileName)
        End If


    End Sub

    Private Sub SpreadsheetControlViewer_CellValueChanged(sender As Object, e As DevExpress.XtraSpreadsheet.SpreadsheetCellEventArgs) Handles SpreadsheetControlViewer.CellValueChanged



    End Sub

    'Public Property FormBorderColor() As Color
    '    Get
    '        Return MyColourSwatch
    '    End Get
    '    Set(ByVal value As Color)
    '        MyColourSwatch = value
    '    End Set
    'End Property

    'Public Class CustomFormPainterSSV
    '    Inherits FormPainter
    '    Public Sub New(ByVal owner As System.Windows.Forms.Control, ByVal provider As DevExpress.Skins.ISkinProvider)
    '        MyBase.New(owner, provider)
    '    End Sub
    '    Private Function GetFormBorderColor() As Color
    '        Dim formBorderColor = (TryCast(Owner, SpreadsheetViewer)).FormBorderColor
    '        Return formBorderColor
    '    End Function
    '    Protected Overrides Sub DrawBackground(ByVal cache As GraphicsCache)
    '        Dim info = GetCaptionInfo()
    '        Dim ee = TryCast(info, ObjectInfoArgs)
    '        Dim formBorderColor = GetFormBorderColor()
    '        cache.FillRectangle(New SolidBrush(formBorderColor), ee.Bounds)
    '    End Sub
    '    Protected Overrides Sub DrawFrameCore(ByVal cache As GraphicsCache, ByVal info As SkinElementInfo, ByVal kind As FrameKind)
    '        Dim formBorderColor = GetFormBorderColor()
    '        cache.FillRectangle(formBorderColor, info.Bounds)
    '    End Sub
    'End Class
End Class