
Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.PresentationManager
Imports Abovo.AbovoUnboundSource
Imports Abovo.GeneralFunctions
Imports Abovo.DataObject
Imports Abovo.DefaultHelpers
Imports Abovo.LogDebugDev
Imports Abovo.FontManager
Imports Abovo.ChangeLogManager

Imports DevExpress.Spreadsheet

Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid

Imports DevExpress.Data
Imports DevExpress.XtraBars.Navigation

Imports DevExpress.XtraGrid.Views.BandedGrid
Imports DevExpress.Utils
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.Utils.Layout

Imports System.ComponentModel
Imports System.Drawing.Drawing2D

Imports DevExpress.XtraSpreadsheet
Imports System.IO
Imports DevExpress.Utils.Drawing.Helpers
Imports DevExpress.XtraSpreadsheet.Model
Public Class FFRProcessing
    Sub FFR_New_Extraction()

        'Revised procedure to extract Financial Forecast Return data to FCA template






        Dim BusPlanFile As IWorkbook

        Dim NumRanges As Integer
        Dim i As Integer

        Dim DestCells() As String
        Dim SourceData() As DevExpress.Spreadsheet.CellRange

        Dim HeaderRows As Integer



        HeaderRows = 4  ' extra rows in model and not in template



        BusPlanFile = FileManager.ActiveWB '   assign Business Plan file

        NumRanges = BusPlanFile.Range("FFRRangeNames").RowCount

        ReDim DestCells(0 To NumRanges - 1)
        ReDim SourceData(0 To NumRanges - 1)



        'Open and assign FFR File

        'added by JW to filter files to FFR 2022

        Dim openFileDialog As New OpenFileDialog()

        openFileDialog.Filter = "Excel Files|*.xlsx;*.xls"
        openFileDialog.Title = "Select an Existing Workbook"

        ' Show the dialog and get the selected file path

        If Not openFileDialog.ShowDialog() = DialogResult.OK Then Return

        Dim FFRFile As New Workbook
        FFRFile.LoadDocument(openFileDialog.FileName)




        If FFRFile.Worksheets("Cover Sheet").Range("$B$4").Value <> "Spreadsheet Import Template - Financial Forecast Return (FFR)" Then

            MsgBox("Sorry, the file selected does not appear to be a correct FFR template which must be downloaded as a provider-specific template from NROSH+." + vbCr + "See https://nroshplus.regulatorofsocialhousing.org.uk/")

            FFRFile.Dispose()
            FFRFile = Nothing

            Return

        End If

        FFRFile.BeginUpdate()

        Dim AddressCell As DevExpress.Spreadsheet.Cell
        Dim SourceRange As DevExpress.Spreadsheet.CellRange
        Dim DestRange As DevExpress.Spreadsheet.CellRange

        For i = 0 To NumRanges - 1

            AddressCell = BusPlanFile.Range("FFRListHeading")(i, 8)
            SourceRange = BusPlanFile.Range(AddressCell.Value.TextValue)
            AddressCell = BusPlanFile.Range("FFRListHeading")(i, 3)
            DestRange = FFRFile.Range(AddressCell.Value.TextValue)
            DestRange.CopyFrom(SourceRange, PasteSpecial.Values)

        Next i

        FFRFile.EndUpdate()

        FFRFile.CalculateFull()


        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Excel Files|*.xlsm|All Files|*.*"
        saveFileDialog.Title = "Save FFR Document"

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
            FFRFile.SaveDocument(saveFileDialog.FileName, DocumentFormat.Xlsm)


        End If



        FFRFile.Dispose()
        FFRFile = Nothing





    End Sub
End Class
