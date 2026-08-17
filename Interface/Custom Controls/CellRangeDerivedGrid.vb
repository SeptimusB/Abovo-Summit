'Imports System.IO
'Imports System.Linq
'Imports System.Reflection
'Imports System.Runtime.Remoting.Contexts
'Imports Abovo
'Imports Abovo.AbovoAppCls
'Imports Abovo.CustomGrid
'Imports Abovo.DataObject
'Imports Abovo.FileManager
'Imports Abovo.GeneralFunctions
'Imports Abovo.LogDebugDev
'Imports Abovo.ObjectMiddler
'Imports Abovo.WSSecurity
'Imports DevExpress.CodeParser
'Imports DevExpress.Data
'Imports DevExpress.Drawing
'Imports DevExpress.Pdf.Native.BouncyCastle.Asn1.X509
'Imports DevExpress.Skins
'Imports DevExpress.Skins.XtraForm
'Imports DevExpress.Spreadsheet
'Imports DevExpress.Spreadsheet.Functions
'Imports DevExpress.Utils
'Imports DevExpress.Utils.Drawing
'Imports DevExpress.XtraBars.Docking2010
'Imports DevExpress.XtraCharts
'Imports DevExpress.XtraEditors.Repository
'Imports DevExpress.XtraGrid
'Imports DevExpress.XtraGrid.Scrolling
'Imports DevExpress.XtraGrid.Views.Base
'Imports DevExpress.XtraGrid.Views.Grid
'Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
'Imports DevExpress.XtraLayout.Customization.Templates
'Imports DevExpress.XtraRichEdit.Layout
'Imports DevExpress.XtraRichEdit.Model
'Imports DevExpress.XtraSpreadsheet
'Imports DevExpress.XtraSpreadsheet.PrintLayoutEngine
'Imports Microsoft.Office.Interop.Excel

'Namespace Abovo
'    Public Class CellRangeDerivedGrid

'        Private ModelID As Integer
'        Private ActiveWorkbook As IWorkbook
'        Private ActiveWorksheet As Worksheet
'        Private DataColumns() As SheetDataColumn
'        Private ActiveRange As DevExpress.Spreadsheet.CellRange
'        Private AddressArray As CellAddressArray

'        Public Sub New(SetModelID As Integer)
'            ModelID = SetModelID
'        End Sub
'        Public Sub InitialiseFromCellRange(CellRange As DevExpress.Spreadsheet.CellRange)
'            ActiveWorkbook = ExcelModels(ModelID).WB
'            ActiveWorksheet = CellRange.Worksheet
'        End Sub
'        Public Function RenderGridOfRangeCells(Ranges As List(Of DevExpress.Spreadsheet.CellRange), Optional ByVal WidthsList As List(Of Integer) = Nothing, Optional ByVal HeightsList As List(Of Integer) = Nothing) As GridControl

'            Dim iR As Integer, jC As Integer
'            Dim CellExamine As DevExpress.Spreadsheet.Cell

'            Dim ResultGrid As New GridControl
'            Dim View As New DevExpress.XtraGrid.Views.Grid.GridView(ResultGrid)
'            ResultGrid.MainView = View
'            ResultGrid.ViewCollection.Add(View)

'            ' Create the grid structure based on the ranges
'            ResultGrid.DataSource = FormUnboundSource()

'            Return ResultGrid


'        End Function
'        Function FormUnboundSource() As AbovoUnboundSource

'            Dim UBSTag As New AbovoUnboundSource.AbovoUnboundSourceTag
'            Dim UBS As New AbovoUnboundSource(0, UBSTag)
'            Dim PropType As System.Type = GetType(Object)

'            For Each PresentedColumn In ActiveRange.ColumnCount

'                If PresentedColumn.ColumnTag.IsCalculated Then UnboundDataSources(UBSDataSourceCount).UBSTag.HasCalcs = True

'                ColCount += 1
'                ColName = "Col_" & ColCount.ToString
'                PresentedColumn.ColumnTag.ActiveColumnName = ColName
'                PropertiesCount += 1
'                ReDim Preserve PropertyArray(PropertiesCount)

'                Select Case PresentedColumn.ColumnTag.DataType

'                    Case "S"
'                        PropType = GetType(String)

'                    Case "I", "Y"
'                        PropType = GetType(Integer)

'                    Case "N", "P", "C", "M"
'                        PropType = GetType(Double)

'                    Case "B"
'                        PropType = GetType(Integer)

'                    Case Else
'                        PropType = GetType(String)

'                End Select

'                SystemLog("Adding column: " & PresentedColumn.ColumnTag.ColumnHeading)

'                PropertyArray(PropertiesCount) = New UnboundSourceProperty With {
'                        .UserTag = PresentedColumn.ColumnTag,
'                        .Name = ColName,
'                        .PropertyType = PropType,
'                        .DisplayName = " " & PresentedColumn.ColumnTag.ColumnHeading & " "
'                    }

'                ColList.Add(ColName)

'            Next

'            AddHandler UBS.ValueNeeded, AddressOf UnboundDS_ValueNeeded
'            'UBS.RowCount = ActiveRange.RowCount
'            'UBS.ColumnCount = ActiveRange.ColumnCount

'            Return UBS
'        End Function

'        Sub RenderAddressArray()

'            AddressArray = New CellAddressArray
'            AddressArray.AddressRows = New CellAddressRow(ActiveRange.RowCount - 1) {}
'            For i As Integer = 0 To ActiveRange.RowCount - 1
'                AddressArray.AddressRows(i) = New CellAddressRow
'                AddressArray.AddressRows(i).AddressPoints = New CellAddressPoint(ActiveRange.ColumnCount - 1) {}
'                For j As Integer = 0 To ActiveRange.ColumnCount - 1
'                    AddressArray.AddressRows(i).AddressPoints(j) = New CellAddressPoint
'                Next
'            Next

'        End Sub
'        Private Sub UnboundDS_ValueNeeded(ByVal sender As Object, ByVal e As DevExpress.Data.UnboundSourceValueNeededEventArgs)

'            Dim UDSSender As AbovoUnboundSource = sender
'            e.Value = GetDSData(e.RowIndex, e.PropertyIndex)

'        End Sub
'        Private Function GetDSData(ByVal rowIndex As Integer, ByVal PropertyIndex As Integer) As Object

'            'If DoDataLog Then SystemLog("Value requested from dataset: " & SetDSIndex.ToString & " Row: " & rowIndex.ToString & " Column: " & PropertyIndex.ToString)

'            Dim DP As CellDataPoint = CellAddressArray.


'            DDataSets(SetDSIndex).DataRows(rowIndex).DataCells(PropertyIndex)
'            Dim DPC As DevExpress.Spreadsheet.Cell = ExcelModels(ModelID).WB.Worksheets(DP.SourceSheet).Cells(DP.SourceAddress)

'            If DPC.DisplayText = "" Then

'                If DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag.DataType = "S" Then
'                    Return ""
'                Else
'                    Return Nothing
'                End If

'            End If

'            Select Case DataPres.DataSets(SetDSIndex).DataColumns(PropertyIndex).ColumnTag.DataType

'                Case "S"

'                    DP.StringValue = DPC.DisplayText
'                    'SystemLog("DSIndex " & SetDSIndex.ToString & " returning string " & DPC.Value.TextValue & " From " & DP.SourceAddress & " of " & DP.SourceSheet)
'                    Return DPC.DisplayText
'                    Exit Function

'                Case "B"

'                    DP.BoolValue = DPC.Value.NumericValue
'                    Return DPC.Value.NumericValue
'                    Exit Function

'                Case "N", "P", "C", "M"
'                    SystemLog("UBS Index " & UBSIndex.ToString & " with DSIndex " & SetDSIndex.ToString & " returning " & DPC.Value.NumericValue.ToString & " From " & DP.SourceAddress & " of " & DP.SourceSheet)
'                    DP.RealValue = DPC.Value.NumericValue
'                    Return DPC.Value.NumericValue
'                    Exit Function

'                Case "I", "Y"

'                    'If DoDataLog Then SystemLog("Returning Integer " & DP.IntValue)
'                    'SystemLog("UBS Index " & UBSIndex.ToString & " with DSIndex " & SetDSIndex.ToString & " returning " & DPC.Value.NumericValue & " From " & DP.SourceAddress & " of " & DP.SourceSheet)
'                    DP.IntValue = DPC.Value.NumericValue
'                    Return CInt(DPC.Value.NumericValue)
'                    Exit Function

'                Case Else

'                    Return Nothing

'            End Select

'        End Function
'        Public Class EditorTemplateSelector
'            Inherits DataTemplateSelector

'            Public Overrides Function SelectTemplate(ByVal item As Object, ByVal container As DependencyObject) As DataTemplate
'                Dim data As EditGridCellData = CType(item, EditGridCellData)
'                Dim dataItem = TryCast(data.RowData.Row, TestData)
'                Return If(dataItem Is Nothing OrElse String.IsNullOrEmpty(dataItem.Editor), Nothing, CType(CType(container, FrameworkElement).FindResource(dataItem.Editor), DataTemplate))
'            End Function
'        End Class
'    End Class

'    Class CellPointTag

'        Public RRefRange As Integer
'        Public CRefRange As Integer
'        Public AddressRRef As Integer
'        Public AddressCRef As Integer
'        Public DataType As String

'        Public SourceSheet As Integer
'        Public SourceAddress As String
'        Public IsCalculated As Boolean
'        Public IsStatic As Boolean
'        Public IsLocked As Boolean
'        Public IsEditable As Boolean

'        Public BackColor As Color
'        Public ForeColor As Color
'        Public Bold As Boolean
'        Public Italic As Boolean

'        Public Sub RefreshData()



'        End Sub

'    End Class

'    Class CellAddressArray
'        Public AddressRows() As CellAddressRow
'        Public AddressColumns() As CellAddressColumns
'    End Class
'    Class CellAddressRow
'        Public AddressPoints() As CellAddressPoint

'    End Class
'End Namespace
