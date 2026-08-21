Imports System
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Runtime.Remoting.Messaging
Imports System.Text
Imports System.Windows.Forms
Imports System.Xml.Serialization
Imports Abovo.AbovoAppCls
Imports Abovo.AbovoExtendedDEControls
Imports Abovo.DataObject
Imports Abovo.FileManager
Imports Abovo.LogDebugDev
Imports DevExpress.Charts.Native
Imports DevExpress.CodeParser
Imports DevExpress.DataAccess.DataFederation
Imports DevExpress.DataAccess.Native.Data
Imports DevExpress.Spreadsheet
Imports DevExpress.Utils.Drawing
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.Utils.Extensions
Imports DevExpress.XtraCharts.Designer.Native
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraPrinting.Native
Imports DevExpress.XtraRichEdit.Model
Imports DevExpress.XtraSpreadsheet.Import.Xls
Imports DevExpress.XtraSpreadsheet.Model
Imports DevExpress.XtraTreeList.Data
Imports Microsoft.Office.Interop
Imports Microsoft.VisualBasic


Namespace Abovo
    Public Class DataManager

        Public ModelID As Integer
        Public DataSetIndex As Integer
        Public DataSets() As DataCellRange
        'Public DSCreators() As exceldataset

        Sub New(SetModelID As Integer)
            DataSetIndex = -1
            ModelID = SetModelID

        End Sub

        Public Function InitialiseDataSet(GSID, CSID) As AbovoTransaction

            Dim Trans As New AbovoTransaction

            Return Trans

        End Function

        Function GetDataSet(CSID As Integer) As AbovoTransaction

            Dim Trans As New AbovoTransaction
            Dim DS As New InstanceDataSet(ModelID, CSID)

            Trans.ObjectReturn = DS

            Return Trans

        End Function
        Public Function GetStatus() As String

            Dim Output As New StringBuilder(4096)
            Output.Append("'''''''''''''''''''''''''''''''''").Append(vbLf)

            If DataSets Is Nothing Then Return Output.ToString()

            For Each DS As DataCellRange In DataSets

                If DS Is Nothing Then Continue For

                Output.Append("DS(").Append(DS.Index).Append(") Name: ").Append(DS.Name).Append(vbLf)
                Output.Append("Dirty: ").Append(DS.IsDirty).Append(vbLf)
                Output.Append("RowsByParam: ").Append(DS.RowCount)
                Output.Append(" ColsByParam: ").Append(DS.ColCount).Append(vbLf)
                Output.Append("RowsByArray: ").Append(If(DS.DataRows Is Nothing, -1, DS.DataRows.Length - 1))
                Output.Append(" ColsByArray: ").Append(If(DS.DataColumns Is Nothing, -1, DS.DataColumns.Length - 1)).Append(vbLf)
                Output.Append(" DataFields: ")

                If DS.DataColumns IsNot Nothing Then
                    Dim ColumnCount As Integer = Math.Min(DS.ColCount, DS.DataColumns.Length)
                    For i As Integer = 0 To ColumnCount - 1
                        Dim Column As SheetDataColumn = DS.DataColumns(i)
                        If Column.ColumnTag Is Nothing Then Continue For
                        Output.Append(Column.ColumnTag.ColumnHeading)
                        Output.Append(" (").Append(Column.ColumnTag.DataType)
                        Output.Append(" DP:").Append(Column.DataAddress)
                        Output.Append(" SP:").Append(Column.SourceAddress).Append("), ")
                    Next
                End If

                Output.Append(vbLf)
                If DS.DataRows IsNot Nothing Then
                    For Each DSRow As SheetDataRow In DS.DataRows

                        If DSRow Is Nothing Then Continue For
                        Output.Append("Row ").Append(DSRow.Index).Append(vbLf)

                        If DSRow.DataCells IsNot Nothing Then
                            For Each DSRowCell As CellDataPoint In DSRow.DataCells
                                If DSRowCell Is Nothing Then Continue For
                                Output.Append("C").Append(DSRowCell.Index).Append(" (").Append(DSRowCell.DataType).Append("), ")
                            Next

                            Output.Append(vbLf)
                            For Each DSRowCell As CellDataPoint In DSRow.DataCells

                                If DSRowCell Is Nothing Then Continue For
                                Select Case DSRowCell.DataType
                                    Case "S"
                                        Output.Append(DSRowCell.StringValue)
                                    Case "I", "Y"
                                        Output.Append(DSRowCell.IntValue)
                                    Case "B"
                                        Output.Append(DSRowCell.BoolValue)
                                    Case "N", "P", "M", "SM", "R"
                                        Output.Append(DSRowCell.RealValue)
                                End Select

                                Output.Append(" (").Append(DSRowCell.SourceAddress).Append("), ")
                            Next
                        End If

                        Output.Append(vbLf)
                    Next
                End If

                Output.Append(vbLf).Append("'''").Append(vbLf)
            Next

            Return Output.ToString()

        End Function
        Public Function GetISEDataStructure(ModelID, SetGSID, SetCSID, SetIntSecID, SetISDID) As DataCellRange

            Dim Trans As New AbovoTransaction
            Dim g = ExcelModels(ModelID)
            Dim IEDSource As ISEDatasource = ExcelModels(ModelID).WBStructure.GroupStructures(SetGSID).ChildStructures(SetCSID).InterfaceSections(SetIntSecID).ISDatasources(SetISDID)
            Dim CellExamine As DevExpress.Spreadsheet.Cell
            Dim CellExamineRight As DevExpress.Spreadsheet.Cell

#Region "SimpleGrid"



            If IEDSource.DSType = "LiveGrid" Then

                Dim WSName As String = IEDSource.CellRangeSources(0).WSName
                Dim CRSource As CellRangeDataSource = IEDSource.CellRangeSources(0)
                Dim CurrWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets(WSName)

                DataSetIndex += 1
                ReDim Preserve DataSets(DataSetIndex)

                DataSets(DataSetIndex) = New DataCellRange(DataSetIndex, ModelID) With {
                    .Name = "Sheet" & WSName & "-IEDS-" & WSName,
                    .IsDirty = False,
                    .SourceWorksheet = WSName,
                    .DataRange = CRSource.DataRange,
                    .RO = IIf(IEDSource.RO = "TRUE", True, False),
                    .ColExpandByNR = If(CRSource.ColsDefinedBy = "NR", CRSource.ColsDefinedByData, Nothing),
                    .RowExpandByNR = If(CRSource.RowsDefinedBy = "NR", CRSource.RowsDefinedByData, Nothing),
                    .LiveGridSourceName = CRSource.LiveGridSourceName,
                    .LiveGridSourceRanges = CRSource.LiveGridSourceRanges,
                    .LiveGridHeaderRows = CRSource.LiveGridHeaderRows,
                    .FormatMap = CRSource.DataFieldDefinitions(0).DataFormat
                }

                DataSets(DataSetIndex).LiveGridSourceAreaReferences =
                    ResolveLiveGridSourceAreaReferences(
                        CurrWS,
                        ExcelModels(ModelID).WB,
                        CRSource)

                If DataSets(DataSetIndex).LiveGridSourceAreaReferences.Count > 0 Then

                    Dim ColIndex As Integer = -1

                    For Each AreaReference As String In
                        DataSets(DataSetIndex).LiveGridSourceAreaReferences

                        Dim SourceArea As DevExpress.Spreadsheet.CellRange =
                            CurrWS.Range(AreaReference)

                        For SourceColumnIndex As Integer =
                            SourceArea.LeftColumnIndex To SourceArea.RightColumnIndex

                            If Not CurrWS.Columns(SourceColumnIndex).Visible Then Continue For

                            ColIndex += 1
                            EnsureArrayCapacity(DataSets(DataSetIndex).DataColumns, ColIndex)
                            DataSets(DataSetIndex).DataColumns(ColIndex) =
                                New SheetDataColumn With {
                                    .ColumnTag = New DataColumnTag With {
                                        .ColumnHeading = String.Empty,
                                        .DataType = "S",
                                        .IsReadOnly = True,
                                        .IsCalculated = True
                                    },
                                    .Index = ColIndex
                                }
                        Next
                    Next

                Else

                    For Each CellRangeSource In IEDSource.CellRangeSources

                        Dim ColIndex As Integer = -1

                        For Each DataFieldDefinition In CellRangeSource.DataFieldDefinitions

                            'MultiRepeatingHeaders = False

                            'CurrDataType = DataFieldDefinition.DataFormat

                            'If CurrDataType = "DUMMY" Then GoTo NextDFD

                            Dim RepeatCount As Integer = 0

                            If DataFieldDefinition.RepeatsByNR = "TRUE" Or DataFieldDefinition.RepeatsByCR = "TRUE" Then

                                Dim DefiningRange As DevExpress.Spreadsheet.CellRange

                                If DataFieldDefinition.RepeatsByNR = "TRUE" Then

                                    DefiningRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(DataFieldDefinition.RepeatingNR).Range

                                Else

                                    DefiningRange = ExcelModels(ModelID).WB.Worksheets(CellRangeSource.WSName).Range(DataFieldDefinition.RepeatsByCRData)

                                End If

                                Dim ColHead As String
                                Dim CellExamineNRD As DevExpress.Spreadsheet.Cell

                                Dim RepeatingItemCount As Integer =
                                If(DefiningRange.RowCount = 1,
                                   DefiningRange.ColumnCount,
                                   If(DefiningRange.ColumnCount = 1,
                                      DefiningRange.RowCount,
                                      DefiningRange.ColumnCount))

                                For x = 0 To RepeatingItemCount - 1

                                    ColIndex += 1
                                    CellExamineNRD =
                                    If(DefiningRange.RowCount = 1,
                                       DefiningRange(0, x),
                                       DefiningRange(x, 0))
                                    ColHead = Replace(DataFieldDefinition.FieldName, "vblf", vbLf) & " vblf " & CellExamineNRD.DisplayText
                                    If Len(DataFieldDefinition.Units) > 0 Then ColHead += DataFieldDefinition.Units
                                    EnsureArrayCapacity(DataSets(DataSetIndex).DataColumns, ColIndex)

                                    DataSets(DataSetIndex).DataColumns(ColIndex) = New SheetDataColumn With {
                                        .ColumnTag = New DataColumnTag With {
                                        .ColumnHeading = ColHead,
                                        .DataType = DataFieldDefinition.DataFormat,
                                        .IsReadOnly = True,
                                        .IsCalculated = True,
                                        .RepeatingNR = DataFieldDefinition.RepeatingNR,
                                        .Units = DataFieldDefinition.Units,
                                        .ShowSummary = DataFieldDefinition.ShowSummary,
                                        .TipText = DataFieldDefinition.TipText
                                },
                                        .Index = ColIndex
                                        }

                                    DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsReadOnly = True

                                Next

                            Else

                                ColIndex += 1

                                EnsureArrayCapacity(DataSets(DataSetIndex).DataColumns, ColIndex)

                                DataSets(DataSetIndex).DataColumns(ColIndex) = New SheetDataColumn With {
                                        .ColumnTag = New DataColumnTag With {
                                        .ColumnHeading = DataFieldDefinition.FieldName,
                                        .DataType = DataFieldDefinition.DataFormat,
                                        .IsReadOnly = True,
                                        .IsCalculated = True,
                                        .Units = DataFieldDefinition.Units,
                                        .ShowSummary = DataFieldDefinition.ShowSummary,
                                        .TipText = DataFieldDefinition.TipText
                                        },
                                        .Index = ColIndex
                                        }

                                DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsReadOnly = True

                            End If

                        Next

                    Next

                End If

                'LiveGrid is a direct, read-only workbook display. Its configured
                'range provides the anchor and row extent, while the generated
                'column definitions provide the current structural width. This
                'keeps StockType-driven workings aligned after columns are added or
                'removed without changing the workbook or Structure.xml.
                Dim ConfiguredLiveRange As DevExpress.Spreadsheet.CellRange =
                    If(DataSets(DataSetIndex).LiveGridSourceAreaReferences.Count > 0,
                       CurrWS.Range(DataSets(DataSetIndex).LiveGridSourceAreaReferences(0)),
                       CurrWS.Range(CRSource.DataRange))

                'LiveGrid binds the worksheet range directly and therefore does
                'not materialise DataRows. Trim the chunk-grown column array before
                'using its length to resolve the current StockType-driven width.
                TrimDataCellCapacity(DataSets(DataSetIndex))

                Dim LiveColumnCount As Integer = DataSets(DataSetIndex).DataColumns.Length

                If LiveColumnCount > 0 Then
                    Dim ResolvedLiveRange As DevExpress.Spreadsheet.CellRange =
                        CurrWS.Range.FromLTRB(
                            ConfiguredLiveRange.LeftColumnIndex,
                            ConfiguredLiveRange.TopRowIndex,
                            ConfiguredLiveRange.LeftColumnIndex + LiveColumnCount - 1,
                            ConfiguredLiveRange.BottomRowIndex)

                    DataSets(DataSetIndex).DataRange = ResolvedLiveRange.GetReferenceA1()
                    DataSets(DataSetIndex).ColCount = LiveColumnCount
                    DataSets(DataSetIndex).RowCount = ResolvedLiveRange.RowCount
                    DataSets(DataSetIndex).UsedRows = ResolvedLiveRange.RowCount
                    DataSets(DataSetIndex).Capacity = 0
                End If

                'The common post-processing path below counts materialised
                'DataRows for Grid/VGrid datasets. A direct LiveGrid has no such
                'objects, so return the fully resolved range dataset here.
                Return DataSets(DataSetIndex)

                ''''''''''''''''''''
                '''''End LiveGrid
                ''''''''''''''''''''

            ElseIf IEDSource.DSType = "SimpleGrid" Then

                If IEDSource.DSSource = "CR" Then

                    Dim WSName As String = IEDSource.CellRangeSources(0).WSName
                    Dim CRSource As CellRangeDataSource = IEDSource.CellRangeSources(0)
                    Dim CurrWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets(WSName)

                    DataSetIndex += 1
                    ReDim Preserve DataSets(DataSetIndex)

                    DataSets(DataSetIndex) = New DataCellRange(DataSetIndex, ModelID) With {
                        .Name = "Sheet" & WSName & "-IEDS-" & WSName,
                        .IsDirty = False,
                        .SourceWorksheet = WSName,
                        .RO = IIf(IEDSource.RO = "TRUE", True, False),
                        .FormatMap = CRSource.DataFieldDefinitions(0).DataFormat
                    }

                    Trans.IntegerReturn = DataSetIndex

                    Dim ColHead As String
                    Dim ColHeaderRange As DevExpress.Spreadsheet.CellRange
                    Dim RowHeaderRange As DevExpress.Spreadsheet.CellRange

                    Dim DataRange As DevExpress.Spreadsheet.CellRange = CurrWS.Range(CRSource.DataRange)

                    If IEDSource.Pivot = "TRUE" Then

                        If CRSource.RowsDefinedBy = "NR" Then
                            ColHeaderRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.RowsDefinedByData).Range
                        Else
                            ColHeaderRange = CurrWS.Range(CRSource.RowsDefinedByData)
                        End If

                        If CRSource.ColsDefinedBy = "NR" Then
                            RowHeaderRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.ColsDefinedByData).Range
                        Else
                            RowHeaderRange = CurrWS.Range(CRSource.ColsDefinedByData)
                        End If

                        ReDim DataSets(DataSetIndex).DataColumns(ColHeaderRange.RowCount)

                        DataSets(DataSetIndex).DataColumns(0) = New SheetDataColumn With {
                                .ColumnTag = New DataColumnTag With {.ColumnHeading = CRSource.ColsDescription, .DataType = "S"},
                                .Index = 0
                                }

                        Dim x As Integer

                        For x = 1 To ColHeaderRange.RowCount

                            CellExamine = ColHeaderRange(x - 1, 0)
                            ColHead = CRSource.RowsDescription & " " & CellExamine.DisplayText

                            If ColHeaderRange.ColumnCount = 2 Then

                                CellExamineRight = ColHeaderRange(x - 1, 1)
                                ColHead = ColHead & vbLf & " (" & CellExamineRight.DisplayText & ")"

                            ElseIf ColHeaderRange.ColumnCount = 3 Then

                                CellExamineRight = ColHeaderRange(x - 1, 1)
                                ColHead = ColHead & vbLf & " (" & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(x - 1, 2)
                                ColHead = ColHead & CellExamineRight.DisplayText & ")"

                            ElseIf ColHeaderRange.ColumnCount = 4 Then

                                CellExamineRight = ColHeaderRange(x - 1, 1)
                                ColHead = ColHead & vbLf & " (" & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(x - 1, 2)
                                ColHead = ColHead & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(x - 1, 3)
                                ColHead = ColHead & CellExamineRight.DisplayText & ")"

                            End If

                            DataSets(DataSetIndex).DataColumns(x) = New SheetDataColumn With {
                                    .ColumnTag = New DataColumnTag With {.ColumnHeading = ColHead, .DataType = CRSource.DataFieldDefinitions(0).DataFormat},
                                    .Index = x
                                    }

                            If CRSource.DataFieldDefinitions(0).ShowSummary IsNot Nothing Then

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.ShowSummary = CRSource.DataFieldDefinitions(0).ShowSummary

                            Else

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.ShowSummary = "FALSE"

                            End If

                            If CRSource.DataFieldDefinitions(0).MinVal = Nothing Then

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MinVal = "FALSE"

                            Else

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MinVal = CRSource.DataFieldDefinitions(0).MinVal

                            End If

                            If CRSource.DataFieldDefinitions(0).MaxVal = Nothing Then

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MaxVal = "FALSE"

                            Else

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MaxVal = CRSource.DataFieldDefinitions(0).MaxVal

                            End If

                        Next

                        Dim RowHead As String

                        DataSets(DataSetIndex).ColCount = ColHeaderRange.RowCount + 1

                        ReDim DataSets(DataSetIndex).DataRows(RowHeaderRange.ColumnCount - 1)

                        For x = 0 To RowHeaderRange.ColumnCount - 1

                            CellExamine = RowHeaderRange(0, x)
                            RowHead = CellExamine.DisplayText

                            If RowHeaderRange.RowCount = 2 Then

                                CellExamineRight = RowHeaderRange(1, x)
                                RowHead = RowHead & vbLf & " (" & CellExamineRight.DisplayText & ")"

                            ElseIf RowHeaderRange.RowCount = 3 Then

                                CellExamineRight = RowHeaderRange(1, x)
                                RowHead = RowHead & vbLf & " (" & CellExamineRight.DisplayText & " "
                                CellExamineRight = RowHeaderRange(2, x)
                                RowHead = RowHead & CellExamineRight.DisplayText & ")"

                            End If

                            DataSets(DataSetIndex).DataRows(x) = New SheetDataRow With {
                                    .Index = x
                                    }

                            ReDim DataSets(DataSetIndex).DataRows(x).DataCells(ColHeaderRange.RowCount)
                            DataSets(DataSetIndex).DataRows(x).DataCells(0) = New CellDataPoint With {
                                .Index = 0,
                                .DataType = "S",
                                .StringValue = RowHead
                            }

                        Next

                        DataSets(DataSetIndex).RowCount = RowHeaderRange.ColumnCount

                        'Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.Worksheets(WSName).Range()
                        Dim clCell As DevExpress.Spreadsheet.Cell
                        Dim i As Integer, j As Integer

                        For i = 0 To RowHeaderRange.ColumnCount - 1

                            For j = 1 To ColHeaderRange.RowCount

                                clCell = DataRange(j - 1, i)

                                DataSets(DataSetIndex).DataRows(i).DataCells(j) = New CellDataPoint
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).SourceAddress = clCell.GetReferenceA1
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).Index = j
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).DataType = DataSets(DataSetIndex).DataColumns(j).ColumnTag.DataType
                                Select Case DataSets(DataSetIndex).DataColumns(j).ColumnTag.DataType

                                    Case "S"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).StringValue = clCell.DisplayText

                                    Case "B"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).BoolValue = CInt(clCell.Value.NumericValue)

                                    Case "D", "P", "C", "M", "SM", "R"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).RealValue = CDbl(clCell.Value.NumericValue)

                                    Case "I", "Y"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).IntValue = CInt(clCell.Value.NumericValue)


                                End Select

                            Next

                        Next


                    Else 'not pivoted

                        If CRSource.RowsDefinedBy = "NR" Then

                            ColHeaderRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.ColsDefinedByData).Range

                        Else

                            ColHeaderRange = CurrWS.Range(CRSource.ColsDefinedByData)

                        End If

                        If CRSource.ColsDefinedBy = "NR" Then

                            RowHeaderRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.RowsDefinedByData).Range

                        Else

                            RowHeaderRange = CurrWS.Range(CRSource.RowsDefinedByData)

                        End If

                        ReDim DataSets(DataSetIndex).DataColumns(ColHeaderRange.ColumnCount)

                        DataSets(DataSetIndex).DataColumns(0) = New SheetDataColumn With {
                                .ColumnTag = New DataColumnTag With {.ColumnHeading = CRSource.RowsDescription, .DataType = "S"},
                                .Index = 0
                                }

                        Dim x As Integer



                        For x = 1 To ColHeaderRange.ColumnCount

                            CellExamine = ColHeaderRange(0, x - 1)
                            ColHead = CellExamine.DisplayText

                            If ColHeaderRange.RowCount = 2 Then

                                CellExamineRight = ColHeaderRange(1, x - 1)
                                ColHead = ColHead & vbLf & " (" & CellExamineRight.DisplayText & ")"

                            ElseIf ColHeaderRange.RowCount = 3 Then

                                CellExamineRight = ColHeaderRange(1, x - 1)
                                ColHead = ColHead & vbLf & " (" & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(2, x - 1)
                                ColHead = ColHead & CellExamineRight.DisplayText & ")"

                            ElseIf ColHeaderRange.RowCount = 4 Then

                                CellExamineRight = ColHeaderRange(1, x - 1)
                                ColHead = ColHead & vbLf & " (" & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(2, x - 1)
                                ColHead = ColHead & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(3, x - 1)
                                ColHead = ColHead & CellExamineRight.DisplayText & ")"

                            End If

                            DataSets(DataSetIndex).DataColumns(x) = New SheetDataColumn With {
                                    .ColumnTag = New DataColumnTag With {.ColumnHeading = ColHead, .DataType = CRSource.DataFieldDefinitions(0).DataFormat},
                                    .Index = x
                                    }

                            If CRSource.DataFieldDefinitions(0).ShowSummary IsNot Nothing Then

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.ShowSummary = CRSource.DataFieldDefinitions(0).ShowSummary

                            Else

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.ShowSummary = "FALSE"

                            End If

                            If CRSource.DataFieldDefinitions(0).MinVal = Nothing Then

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MinVal = "FALSE"

                            Else

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MinVal = CRSource.DataFieldDefinitions(0).MinVal

                            End If

                            If CRSource.DataFieldDefinitions(0).MaxVal = Nothing Then

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MaxVal = "FALSE"

                            Else

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MaxVal = CRSource.DataFieldDefinitions(0).MaxVal

                            End If

                        Next

                        Dim RowHead As String

                        DataSets(DataSetIndex).RowCount = RowHeaderRange.RowCount

                        ReDim DataSets(DataSetIndex).DataRows(RowHeaderRange.RowCount - 1)

                        For x = 0 To RowHeaderRange.RowCount - 1

                            CellExamine = RowHeaderRange(x, 0)
                            RowHead = CellExamine.DisplayText

                            If RowHeaderRange.ColumnCount = 2 Then

                                CellExamineRight = RowHeaderRange(x, 1)
                                RowHead = RowHead & " (" & CellExamineRight.DisplayText & ")"

                            ElseIf RowHeaderRange.ColumnCount = 3 Then

                                CellExamineRight = RowHeaderRange(x, 1)
                                RowHead = RowHead & " (" & CellExamineRight.DisplayText & " "
                                CellExamineRight = RowHeaderRange(x, 2)
                                RowHead = RowHead & CellExamineRight.DisplayText & ")"

                            End If

                            DataSets(DataSetIndex).DataRows(x) = New SheetDataRow With {
                                    .Index = x
                                    }

                            ReDim DataSets(DataSetIndex).DataRows(x).DataCells(ColHeaderRange.ColumnCount)
                            DataSets(DataSetIndex).DataRows(x).DataCells(0) = New CellDataPoint With {
                                .Index = 0,
                                .DataType = "S",
                                .StringValue = RowHead
                            }
                        Next

                        'DataSets(DataSetIndex).RowCount = RowHeaderRange.ColumnCount

                        'Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.Worksheets(WSName).Range()
                        Dim clCell As DevExpress.Spreadsheet.Cell
                        Dim i As Integer, j As Integer

                        For i = 0 To RowHeaderRange.RowCount - 1

                            For j = 1 To ColHeaderRange.ColumnCount

                                clCell = DataRange(i, j - 1)

                                DataSets(DataSetIndex).DataRows(i).DataCells(j) = New CellDataPoint
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).SourceAddress = clCell.GetReferenceA1
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).Index = j
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).DataType = DataSets(DataSetIndex).DataColumns(j).ColumnTag.DataType

                                Select Case DataSets(DataSetIndex).DataColumns(j).ColumnTag.DataType

                                    Case "S"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).StringValue = clCell.DisplayText

                                    Case "B"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).BoolValue = CInt(clCell.Value.NumericValue)

                                    Case "D", "P", "C", "M", "SM", "R"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).RealValue = CDbl(clCell.Value.NumericValue)

                                    Case "I", "Y"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).IntValue = CInt(clCell.Value.NumericValue)


                                End Select

                            Next

                        Next

                    End If

                End If
#End Region

#Region "MergeAndPivot"

            ElseIf IEDSource.DSType = "MergeDownAndPivot" Then

                Dim SourceWS As String
                Dim WSName As String = IEDSource.CellRangeSources(0).WSName
                Dim CRSource As CellRangeDataSource = IEDSource.CellRangeSources(0)
                Dim CurrWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets(WSName)
                Dim ColIndex As Integer = -1
                Dim CurrDataType As String
                Dim RowOffset As Integer = 0
                Dim ForBandRowOffset As Integer = 0
                Dim BandRowsApplied As Boolean = False
                Dim DoBands As Boolean = False
                Dim ColOffset As Integer = 0
                Dim IsCalc As Boolean = False
                Dim i As Integer, j As Integer, k As Integer
                Dim MultiRepeatingHeaders As Boolean = False
                Dim ColHead As String
                Dim RecordSkipOffset As Integer = 0
                Dim IsMappedRightPosition As Integer = 0
                Dim RowCount As Integer = 0
                Dim ColCount As Integer = 0

                Dim RowExpandsByModel As String = "NONE"

                If IEDSource.RowsExpandModel = "NRRI" Then RowExpandsByModel = "NRRI"
                If IEDSource.RowsExpandModel = "NRCI" Then RowExpandsByModel = "NRCI"

                Dim ApplyRule As Boolean = False
                DataSetIndex += 1
                ReDim Preserve DataSets(DataSetIndex)

                DataSets(DataSetIndex) = New DataCellRange(DataSetIndex, ModelID) With {
                    .Name = "MergeDS" & WSName & "-IEDS-" & IEDSource.ISDName,
                    .IsDirty = False,
                    .RowExpandsByModel = RowExpandsByModel,
                    .HasBands = False,
                    .SkipLastRecords = IIf(IEDSource.SkipLastRecords Is Nothing, 0, CInt(IEDSource.SkipLastRecords)),
                    .SourceWorksheet = WSName,
                    .RO = IIf(IEDSource.RO = "TRUE", True, False)
                }

                RecordSkipOffset = DataSets(DataSetIndex).SkipLastRecords

                Trans.IntegerReturn = DataSetIndex

                Dim DataRange As DevExpress.Spreadsheet.CellRange = Nothing

                If CRSource.NRDSName = "CR" Then

                    DataRange = CurrWS.Range(CRSource.DataRange)

                Else

                    DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.NRDSName).Range

                End If

                ReDim Preserve DataSets(DataSetIndex).DataRows(DataRange.ColumnCount - (1 + RecordSkipOffset))

                DataSets(DataSetIndex).RowCount = DataRange.ColumnCount - RecordSkipOffset

                Dim conditionalFormattings As DevExpress.Spreadsheet.ConditionalFormattingCollection = CurrWS.ConditionalFormattings


                For i = 0 To DataRange.ColumnCount - (1 + RecordSkipOffset)

                    DataSets(DataSetIndex).DataRows(i) = New SheetDataRow With {.Index = i}

                Next

                k = DataSets(DataSetIndex).RowCount

                For Each CellRangeSource In IEDSource.CellRangeSources

                    If CellRangeSource.BandID IsNot Nothing Then
                        DoBands = True
                        ForBandRowOffset = 0
                        Exit For
                    End If

                Next

                For Each CellRangeSource In IEDSource.CellRangeSources

                    RowOffset = 0
                    ColOffset = 0

                    If CellRangeSource.IsCalculated = "TRUE" Then

                        If CellRangeSource.OffSetNR = "CR" Then

                            DataRange = CurrWS.Range(CellRangeSource.DataRange)

                        Else

                            DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CellRangeSource.OffSetNR).Range

                        End If

                        SourceWS = DataRange.Worksheet.Name
                        'CurrWS = ExcelModels(ModelID).WB.Worksheets(SourceWS)
                        IsCalc = True
                        DataSets(DataSetIndex).HasCalcs = True

                        Dim OffSet As String = CellRangeSource.OffSetBy
                        Dim Offsetby As Integer = CInt(Right(OffSet, Len(OffSet) - 1))

                        If Left(OffSet, 1) = "R" Then

                            RowOffset += Offsetby
                            If Offsetby > 0 Then RowOffset += DataRange.RowCount - 1

                        Else

                            ColOffset += Offsetby
                            If Offsetby > 0 Then ColOffset += DataRange.ColumnCount - 1

                        End If

                    Else

                        If CellRangeSource.NRDSName = "CR" Then

                            DataRange = CurrWS.Range(CellRangeSource.DataRange)

                        Else

                            DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CellRangeSource.NRDSName).Range

                        End If


                        SourceWS = DataRange.Worksheet.Name

                    End If

                    If Not IsNothing(CellRangeSource.IsMappedRightByNR) AndAlso CellRangeSource.IsMappedRightByNR <> "" Then

                        Dim MappedRightRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CellRangeSource.IsMappedRightByNR).Range
                        Dim MappedRightCellColCount As Integer = DataRange.ColumnCount - MappedRightRange.ColumnCount
                        ColOffset += MappedRightCellColCount

                    End If

                    i = 0
                    j = 0

                    For Each DataFieldDefinition In CellRangeSource.DataFieldDefinitions

                        MultiRepeatingHeaders = False

                        CurrDataType = DataFieldDefinition.DataFormat

                        If CurrDataType = "DUMMY" Then GoTo NextDFD

                        Dim RepeatCount As Integer = 0

                        '''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'SET UP FOR REPEATING HEADERS IF REQUIRED
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If DataFieldDefinition.IsDummy = "TRUE" Then

                            ColIndex += 1
                            EnsureArrayCapacity(DataSets(DataSetIndex).DataColumns, ColIndex)
                            DataSets(DataSetIndex).DataColumns(ColIndex) = New SheetDataColumn With {
                                    .ColumnTag = New DataColumnTag With {
                                    .ColumnHeading = "",
                                    .DataType = "DUMMY",
                                    .IsReadOnly = True,
                                    .IsCalculated = False,
                                    .IsDummyColumn = True
                                    },
                                    .Index = ColIndex
                                    }

                            GoTo NextDFD

                        End If

                        If CellRangeSource.BandID IsNot Nothing Then DataSets(DataSetIndex).HasBands = True

                        If DataFieldDefinition.RepeatsByNR = "TRUE" Or DataFieldDefinition.RepeatsByCR = "TRUE" Then

                            Dim RepeatMethod As String = "PORT"

                            Dim DefiningRange As DevExpress.Spreadsheet.CellRange

                            If DataFieldDefinition.RepeatsByNR = "TRUE" Then
                                DefiningRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(DataFieldDefinition.RepeatingNR).Range
                            Else
                                DefiningRange = ExcelModels(ModelID).WB.Worksheets(CellRangeSource.WSName).Range(DataFieldDefinition.RepeatsByCRData)
                            End If

                            Dim RecordPreOffset As Integer = IIf(IsNothing(DataFieldDefinition.RepeatingPreRows), 0, CInt(DataFieldDefinition.RepeatingPreRows))
                            Dim RecordPostOffset As Integer = IIf(IsNothing(DataFieldDefinition.RepeatingPostRows), 0, CInt(DataFieldDefinition.RepeatingPostRows))

                            'Create in column editors for the Repeating NR's column definition

                            If DefiningRange.RowCount > DefiningRange.ColumnCount Then

                                RepeatMethod = "PORT"
                                RepeatCount = DefiningRange.RowCount

                            Else

                                RepeatMethod = "LAND"
                                RepeatCount = DefiningRange.ColumnCount

                            End If

                            If IEDSource.MergedHeader Then

                                If DefiningRange.ColumnCount > 1 Then

                                    MultiRepeatingHeaders = True

                                End If

                            End If

                            RepeatCount = DefiningRange.RowCount

                            Dim CellExamineNRD As DevExpress.Spreadsheet.Cell

                            j = -1
                            j -= RecordPreOffset
                            Dim ActualRowIndex As Integer = -1

                            For DefNRIndex = (0 - RecordPreOffset) To RepeatCount - 1 + RecordPostOffset

                                j += 1
                                ActualRowIndex += 1

                                If RepeatMethod = "PORT" Then

                                    CellExamineNRD = DefiningRange(DefNRIndex, 0)

                                Else

                                    CellExamineNRD = DefiningRange(0, DefNRIndex)

                                End If

                                ColHead = DataFieldDefinition.FieldName & " " & CellExamineNRD.DisplayText

                                'If Not IsNothing(DataFieldDefinition.RepeatingHeaderText) Then ColHead = DataFieldDefinition.RepeatingHeaderText & " " & ColHead

                                If MultiRepeatingHeaders = True Then

                                    CellExamineRight = DefiningRange(DefNRIndex, 1)

                                    If Not IsNothing(DataFieldDefinition.ExtraHeadingPreWord) Then
                                        ColHead += vbLf & " (" & DataFieldDefinition.ExtraHeadingPreWord & " " & CellExamineRight.DisplayText & ")"
                                    Else
                                        ColHead += vbLf & " (" & CellExamineRight.DisplayText & ")"
                                    End If

                                End If

                                Dim ReadOnyHeaderEditColumns As Integer = 0

                                If Not IsNothing(DataFieldDefinition.EditRepNRHereROInitialLines) Then ReadOnyHeaderEditColumns = CInt(DataFieldDefinition.EditRepNRHereROInitialLines)

                                If Not DataFieldDefinition.Units Is Nothing Then ColHead += vbLf & DataFieldDefinition.Units

                                'If CellExamineNRD.DisplayText = "" Then GoTo Nextnrds

                                ColIndex += 1

                                EnsureArrayCapacity(DataSets(DataSetIndex).DataColumns, ColIndex)

                                If DataFieldDefinition.HasRule = "TRUE" Then

                                    DataSets(DataSetIndex).HasRules = True
                                    ApplyRule = True

                                Else

                                    ApplyRule = False

                                End If

                                Dim CalcCol As Boolean = False

                                If CellRangeSource.IsCalculated = "TRUE" OrElse DataFieldDefinition.IsColCalculated = "TRUE" Then CalcCol = True

                                DataSets(DataSetIndex).DataColumns(ColIndex) = New SheetDataColumn With {
                                .ColumnTag = New DataColumnTag With {
                                .ColumnHeading = Replace(ColHead, "vblf", vbLf),
                                .DataType = CurrDataType,
                                .IsReadOnly = IIf(DataFieldDefinition.RO = "TRUE", True, False),
                                .IsCalculated = CalcCol,
                                .HasActions = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", True, False),
                                .HasRules = ApplyRule,
                                .RepeatingNR = DataFieldDefinition.RepeatingNR,
                                .MinimumWidthChars = ParseMinimumWidthChars(DataFieldDefinition.MinWidthChars),
                                .ActionNR = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", DataFieldDefinition.RepeatingNR, Nothing),
                                .ActionData = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", "NRRI", Nothing),
                                .Units = DataFieldDefinition.Units,
                                .IsFixed = IIf(DataFieldDefinition.Fixed = "TRUE", True, False),
                                .ShowSummary = DataFieldDefinition.ShowSummary,
                                .BandEditDescription = DataFieldDefinition.BandEditDescription,
                                .MinVal = DataFieldDefinition.MinVal,
                                .MaxVal = DataFieldDefinition.MaxVal,
                                .BandID = CellRangeSource.BandID,
                                .BandTipText = CellRangeSource.BandTipText,
                                .EditRepNRHere = IIf(UCase(DataFieldDefinition.EditRepNRHere) = "TRUE", True, False),
                                .EditRepNRHereDataFormat = DataFieldDefinition.EditRepNRHereDataFormat,
                                .EditRepNRHereEditor = DataFieldDefinition.EditRepNRHereEditor,
                                .EditRepNRHereComboRepository = DataFieldDefinition.EditRepNRHereComboRepository,
                                .EditRepNRHereExpansionMethod = DataFieldDefinition.EditRepNRHereExpansionMethod,
                                .EditRepNRHereRule = DataFieldDefinition.EditRepNRHereRule,
                                .AllowEditRepNRHereBlanks = True,
                                .EditRepNRNROrientation = RepeatMethod,
                                .EditNRIndexPosition = DefNRIndex,
                                .EditRepNRHereInitialValue = CellExamineNRD.DisplayText,
                                .IsControlColumn = True,
                                .RepeatingHeaderText = DataFieldDefinition.RepeatingHeaderText,
                                .TipText = DataFieldDefinition.TipText,
                                .RepositaryID = DataFieldDefinition.RepositaryItemID
                                },
                                .Index = ColIndex
                                }


                                If DefNRIndex < ReadOnyHeaderEditColumns Then DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.EditRepNRHereROColumn = True
                                ' .DontDrawCellHeader = True,
                                'If DoBands And Not BandRowsApplied Then

                                '    Dim q As Integer = DataSets(DataSetIndex).DataRows.Count

                                '    ReDim Preserve DataSets(DataSetIndex).DataRows(q)

                                '    DataSets(DataSetIndex).DataRows(q) = New SheetDataRow With {.Index = q}
                                '    DataSets(DataSetIndex).DataRows(0).IsControlRow = True
                                '    DataSets(DataSetIndex).HasBands = True
                                '    BandRowsApplied = True

                                'End If


                                If CellRangeSource.BandID IsNot Nothing Then

                                    DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.BandID = CellRangeSource.BandID
                                    DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.BandTipText = CellRangeSource.BandTipText
                                    If CellRangeSource.BandEditDescription IsNot Nothing Then DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.BandEditDescription = CellRangeSource.BandEditDescription

                                End If

                                If CellRangeSource.RO = "TRUE" Then

                                    DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsReadOnly = True

                                End If

                                ''''TODO if DFD prop includes IsCal
                                '''''DataFieldDefinition.
                                '''''''''''

                                'If IEDSource.MergedHeader Then

                                '    If DefiningRange.ColumnCount > 1 Then



                                '    End If

                                'End If

                                For i = 0 To k - 1

                                    CellExamine = DataRange(ActualRowIndex + RowOffset, i + ColOffset)

                                    EnsureArrayCapacity(
                                        DataSets(DataSetIndex).DataRows(i + ForBandRowOffset).DataCells,
                                        ColIndex)

                                    Dim LockCell As Boolean = False

                                    If CellExamine.Protection.Locked Then LockCell = True

                                    DataSets(DataSetIndex).DataRows(i + ForBandRowOffset).DataCells(ColIndex) = New CellDataPoint With {
                                    .FoColor = CellExamine.Font.Color,
                                    .BGColor = CellExamine.Fill.BackgroundColor,
                                    .Index = ColIndex,
                                    .IsLocked = LockCell,
                                    .SourceAddress = CellExamine.GetReferenceA1,
                                    .SourceSheet = SourceWS}

                                    Dim CDP As CellDataPoint = DataSets(DataSetIndex).DataRows(i + ForBandRowOffset).DataCells(ColIndex)

                                    If ApplyRule Then CDP.IsLocked = IIf(CellExamine.Fill.PatternType = PatternType.Solid, False, True)

                                    If Not CellExamine.Value.IsEmpty Then

                                        If Not DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsCalculated Then CDP.IsEmpty = False

                                        Select Case CurrDataType

                                            Case "S"

                                                CDP.StringValue = CellExamine.DisplayText

                                            Case "B"

                                                CDP.BoolValue = CInt(CellExamine.Value.NumericValue)

                                            Case "D", "P", "C", "M", "SM", "R"

                                                CDP.RealValue = CDbl(CellExamine.Value.NumericValue)

                                            Case "I", "Y"

                                                CDP.IntValue = CInt(CellExamine.Value.NumericValue)

                                        End Select

                                        CDP.IsEmpty = False

                                    Else

                                        CDP.IsEmpty = True
                                        CDP.StringValue = Nothing
                                        CDP.BoolValue = Nothing
                                        CDP.RealValue = Nothing
                                        CDP.IntValue = Nothing

                                    End If

                                Next
Nextnrds:
                            Next

                        Else

                            '''''''''''''''''''''''''''''''''''''''''''''''''''''
                            'SET UP FOR NON REPEATING COLUMNS
                            '''''''''''''''''''''''''''''''''''''''''''''''''''''

                            If DataFieldDefinition.IsDummy = "TRUE" Then

                                ColIndex += 1
                                EnsureArrayCapacity(DataSets(DataSetIndex).DataColumns, ColIndex)
                                DataSets(DataSetIndex).DataColumns(ColIndex) = New SheetDataColumn With {
                                    .ColumnTag = New DataColumnTag With {
                                    .ColumnHeading = "",
                                    .DataType = "DUMMY",
                                    .IsReadOnly = True,
                                    .IsCalculated = False,
                                    .IsDummyColumn = True
                                    },
                                    .Index = ColIndex
                                    }

                                GoTo NextDFD

                            End If

                            ColIndex += 1

                            EnsureArrayCapacity(DataSets(DataSetIndex).DataColumns, ColIndex)

                            If DataFieldDefinition.HasRule = "TRUE" Then
                                DataSets(DataSetIndex).HasRules = True
                                ApplyRule = True
                            Else
                                ApplyRule = False
                            End If

                            Dim CalcCol As Boolean = False

                            If CellRangeSource.IsCalculated = "TRUE" OrElse DataFieldDefinition.IsColCalculated = "TRUE" Then CalcCol = True



                            DataSets(DataSetIndex).DataColumns(ColIndex) = New SheetDataColumn With {
                            .ColumnTag = New DataColumnTag With {
                            .ColumnHeading = Replace(DataFieldDefinition.FieldName, "vblf", vbLf),
                            .DataType = CurrDataType,
                            .IsReadOnly = IIf(DataFieldDefinition.RO = "TRUE", True, False),
                            .IsCalculated = CalcCol,
                            .HasRules = ApplyRule,
                            .ShowSummary = DataFieldDefinition.ShowSummary,
                            .BandEditDescription = DataFieldDefinition.BandEditDescription,
                            .MinVal = DataFieldDefinition.MinVal,
                            .Units = DataFieldDefinition.Units,
                            .MaxVal = DataFieldDefinition.MaxVal,
                            .IsFixed = IIf(DataFieldDefinition.Fixed = "TRUE", True, False),
                            .TipText = DataFieldDefinition.TipText,
                            .RepositaryID = DataFieldDefinition.RepositaryItemID
                            },
                            .Index = ColIndex
                            }

                            If CellRangeSource.RO = "TRUE" Then

                                DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsReadOnly = True

                            End If

                            'If DoBands And Not BandRowsApplied Then

                            '    Dim q As Integer = DataSets(DataSetIndex).DataRows.Count

                            '    ReDim Preserve DataSets(DataSetIndex).DataRows(q)

                            '    DataSets(DataSetIndex).DataRows(q) = New SheetDataRow With {.Index = q}
                            '    DataSets(DataSetIndex).DataRows(0).IsControlRow = True
                            '    DataSets(DataSetIndex).HasBands = True
                            '    BandRowsApplied = True

                            'End If


                            If CellRangeSource.BandID IsNot Nothing Then

                                DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.BandID = CellRangeSource.BandID
                                DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.BandTipText = CellRangeSource.BandTipText
                                If CellRangeSource.BandEditDescription IsNot Nothing Then DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.BandEditDescription = CellRangeSource.BandEditDescription

                            End If




                            Dim LockCell As Boolean = False

                            For i = 0 To k - 1

                                LockCell = False

                                CellExamine = DataRange(j + RowOffset, i + ColOffset)

                                EnsureArrayCapacity(
                                    DataSets(DataSetIndex).DataRows(i + ForBandRowOffset).DataCells,
                                    ColIndex)

                                DataSets(DataSetIndex).DataRows(i + ForBandRowOffset).DataCells(ColIndex) = New CellDataPoint With {
                                .FoColor = CellExamine.Font.Color,
                                .BGColor = CellExamine.Fill.BackgroundColor,
                                .Index = ColIndex,
                                .IsLocked = LockCell,
                                .SourceAddress = CellExamine.GetReferenceA1,
                                .SourceSheet = SourceWS}

                                Dim CDP As CellDataPoint = DataSets(DataSetIndex).DataRows(i + ForBandRowOffset).DataCells(ColIndex)

                                If ApplyRule Then

                                    CDP.IsLocked = IIf(CellExamine.Fill.PatternType = PatternType.Solid, False, True)

                                End If

                                If Len(CellExamine.DisplayText) > 0 Then

                                    If Not DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsCalculated Then

                                        DataSets(DataSetIndex).DataRows(i + ForBandRowOffset).IsEmpty = False

                                    End If

                                    Select Case CurrDataType

                                        Case "S"

                                            CDP.StringValue = CellExamine.DisplayText

                                        Case "B"

                                            CDP.BoolValue = CInt(CellExamine.Value.NumericValue)

                                        Case "D", "P", "C", "M", "SM", "R"

                                            CDP.RealValue = CDbl(CellExamine.Value.NumericValue)

                                        Case "I", "Y"

                                            CDP.IntValue = CInt(CellExamine.Value.NumericValue)

                                    End Select

                                    CDP.IsEmpty = False

                                Else

                                    CDP.IsEmpty = True

                                    CDP.StringValue = Nothing
                                    CDP.BoolValue = Nothing
                                    CDP.RealValue = Nothing
                                    CDP.IntValue = Nothing

                                End If

                            Next

                        End If
NextDFD:
                        j += 1
                    Next

                Next

#End Region

#Region "Single Cell"

            ElseIf IEDSource.DSType = "SingleCell" Then

                Dim CurrDataType As String
                Dim WSName As String = IEDSource.CellRangeSources(0).WSName

                DataSetIndex += 1

                ReDim Preserve DataSets(DataSetIndex)
                Dim CRSource As CellRangeDataSource = IEDSource.CellRangeSources(0)

                DataSets(DataSetIndex) = New DataCellRange(DataSetIndex, ModelID) With {
                .Name = "SingleCellDS" & WSName & "-IEDS-" & IEDSource.ISDName,
                .IsDirty = False,
                .SourceWorksheet = WSName,
                .RO = IIf(IEDSource.RO = "TRUE", True, False)
            }
                Dim DataRange As DevExpress.Spreadsheet.CellRange = Nothing

                If CRSource.NRDSName Is Nothing Then

                    DataRange = ExcelModels(ModelID).WB.Worksheets(CRSource.WSName).Range(CRSource.DataRange)

                Else

                    DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.NRDSName).Range

                End If


                Dim CellDF As DataFieldDefinition = CRSource.DataFieldDefinitions(0)

                ReDim Preserve DataSets(DataSetIndex).DataColumns(0)

                CurrDataType = CellDF.DataFormat

                DataSets(DataSetIndex).DataColumns(0) = New SheetDataColumn With {
                                    .ColumnTag = New DataColumnTag With {
                                    .ColumnHeading = Replace(CellDF.FieldName, "vblf", vbLf),
                                    .DataType = CurrDataType,
                                    .IsReadOnly = IIf(CellDF.RO = "TRUE", True, False),
                                    .IsCalculated = IIf(CRSource.IsCalculated = "TRUE", True, False),
                                    .MinVal = CellDF.MinVal,
                                    .MaxVal = CellDF.MaxVal,
                                    .RepositaryID = CellDF.RepositaryItemID
                                    },
                                    .Index = 0
                                    }
                ReDim Preserve DataSets(DataSetIndex).DataRows(0)

                DataSets(DataSetIndex).DataRows(0) = New SheetDataRow With {.Index = 0}

                ReDim Preserve DataSets(DataSetIndex).DataRows(0).DataCells(0)
                Dim LockCell As Boolean = False

                CellExamine = DataRange(0, 0)

                If CellExamine.Protection.Locked Then LockCell = True

                DataSets(DataSetIndex).DataRows(0).DataCells(0) = New CellDataPoint With {
                                        .FoColor = CellExamine.Font.Color,
                                        .BGColor = CellExamine.Fill.BackgroundColor,
                                        .Index = 0,
                                        .IsLocked = LockCell,
                                        .SourceAddress = CellExamine.GetReferenceA1,
                                        .SourceSheet = WSName}

                If Not CellExamine.Value.IsEmpty Then


                    Select Case CurrDataType

                        Case "S"

                            DataSets(DataSetIndex).DataRows(0).DataCells(0).StringValue = CellExamine.DisplayText

                        Case "B"

                            DataSets(DataSetIndex).DataRows(0).DataCells(0).BoolValue = CInt(CellExamine.Value.NumericValue)

                        Case "P", "C", "M", "SM", "R"

                            DataSets(DataSetIndex).DataRows(0).DataCells(0).RealValue = CDbl(CellExamine.Value.NumericValue)

                        Case "I", "Y", "D"

                            DataSets(DataSetIndex).DataRows(0).DataCells(0).IntValue = CInt(CellExamine.Value.NumericValue)

                    End Select

                    DataSets(DataSetIndex).DataRows(0).DataCells(0).IsEmpty = False

                Else

                    DataSets(DataSetIndex).DataRows(0).DataCells(0).IsEmpty = True

                End If
#End Region

#Region "MergeAcrossNoPivot"

            ElseIf IEDSource.DSType = "MergeAcross" Then

                'Dim SourceWS As String
                Dim WSName As String = IEDSource.CellRangeSources(0).WSName
                Dim CRSource As CellRangeDataSource = IEDSource.CellRangeSources(0)
                Dim CurrWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets(WSName)
                Dim ColIndex As Integer = -1
                Dim CurrDataType As String
                Dim RowOffset As Integer = 0
                Dim ColOffset As Integer = 0
                Dim ColHead As String
                Dim RecordEndSkipOffset As Integer = 0

                Dim ValidationsSet As New DataValidationsSet(CurrWS)

                Dim IsCalc As Boolean = False
                Dim i As Integer, j As Integer, k As Integer
                DataSetIndex += 1
                ReDim Preserve DataSets(DataSetIndex)
                Dim ApplyRule As Boolean = False

                Dim RowExpandsByModel As String = "NONE"

                If IEDSource.RowsExpandModel = "NRRI" Then RowExpandsByModel = "NRRI"
                If IEDSource.RowsExpandModel = "NRCI" Then RowExpandsByModel = "NRCI"

                Dim RowExpandsByNR As String = "NONE"

                If Not IEDSource.RowExpandByNR Is Nothing Then RowExpandsByNR = IEDSource.RowExpandByNR

                DataSets(DataSetIndex) = New DataCellRange(DataSetIndex, ModelID) With {
                    .Name = "MergeAcrossDS" & WSName & "-IEDS: " & IEDSource.ISDName,
                    .IsDirty = False,
                    .SourceWorksheet = WSName,
                    .HasBands = False,
                    .SkipLastRecords = IIf(IEDSource.SkipLastRecords Is Nothing, 0, CInt(IEDSource.SkipLastRecords)),
                    .RowExpandsByModel = RowExpandsByModel,
                    .RowExpandByNR = RowExpandsByNR,
                    .RO = IIf(IEDSource.RO = "TRUE", True, False)
                }

                RecordEndSkipOffset = DataSets(DataSetIndex).SkipLastRecords

                Trans.IntegerReturn = DataSetIndex

                Dim DataRange As DevExpress.Spreadsheet.CellRange

                If CRSource.IsCalculated = "TRUE" Then

                    DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.OffSetNR).Range
                    WSName = DataRange.Worksheet.Name

                Else

                    If CRSource.NRDSName = "CR" Then

                        DataRange = CurrWS.Range(CRSource.DataRange)

                    Else

                        DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.NRDSName).Range

                    End If

                End If

                k = DataRange.RowCount - RecordEndSkipOffset


                ReDim Preserve DataSets(DataSetIndex).DataRows(k - 1)

                DataSets(DataSetIndex).Capacity = k

                DataSets(DataSetIndex).RowCount = k

                For i = 0 To k - 1

                    DataSets(DataSetIndex).DataRows(i) = New SheetDataRow With {.Index = i}

                Next

                For Each CellRangeSource In IEDSource.CellRangeSources

                    RowOffset = 0
                    ColOffset = 0

                    If CellRangeSource.IsCalculated = "TRUE" Then

                        If CellRangeSource.NRDSName = "CR" Then

                            DataRange = CurrWS.Range(CellRangeSource.DataRange)

                        Else

                            DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CellRangeSource.OffSetNR).Range

                        End If

                        Dim RngWidth As Integer = DataRange.ColumnCount

                        IsCalc = True

                        Dim OffSet As String = CellRangeSource.OffSetBy
                        Dim Offsetby As Integer = CInt(Right(OffSet, Len(OffSet) - 1))

                        If Left(OffSet, 1) = "R" Then

                            RowOffset += Offsetby
                            'If Offsetby > 0 Then RowOffset += k - 1

                        Else

                            ColOffset += Offsetby + (RngWidth - 1)
                            'If Offsetby > 0 Then ColOffset += k - 1

                        End If

                    Else

                        If CellRangeSource.NRDSName = "CR" Then

                            DataRange = CurrWS.Range(CellRangeSource.DataRange)

                        Else

                            DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CellRangeSource.NRDSName).Range

                        End If

                    End If

                    WSName = DataRange.Worksheet.Name

                    i = 0
                    j = 0

                    Dim MultiRepeatingHeaders As Boolean = False

                    For Each DataFieldDefinition In CellRangeSource.DataFieldDefinitions

                        If DataFieldDefinition.IsDummy = "TRUE" Then

                            ColIndex += 1
                            EnsureArrayCapacity(DataSets(DataSetIndex).DataColumns, ColIndex)
                            DataSets(DataSetIndex).DataColumns(ColIndex) = New SheetDataColumn With {
                                        .ColumnTag = New DataColumnTag With {
                                        .ColumnHeading = "",
                                        .DataType = "DUMMY",
                                        .IsReadOnly = True,
                                        .IsFixed = False,
                                        .IsCalculated = False,
                                        .IsDummyColumn = True
                                        },
                                        .Index = ColIndex
                                        }

                            GoTo NextDFD2

                        End If

                        MultiRepeatingHeaders = False
                        CurrDataType = DataFieldDefinition.DataFormat

                        If CurrDataType = "DUMMY" Then GoTo NextDFD2

                        Dim RepeatCount As Integer = 0

                        If CellRangeSource.BandID IsNot Nothing Then DataSets(DataSetIndex).HasBands = True

                        If DataFieldDefinition.RepeatsByNR = "TRUE" OrElse DataFieldDefinition.RepeatsByCR = "TRUE" Then

                            Dim RepeatMethod As String = "PORT"

                            Dim DefiningRange As DevExpress.Spreadsheet.CellRange = Nothing

                            If DataFieldDefinition.RepeatsByNR = "TRUE" Then

                                DefiningRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(DataFieldDefinition.RepeatingNR).Range

                            Else

                                DefiningRange = CurrWS.Range(DataFieldDefinition.RepeatsByCRData)

                            End If

                            Dim RecordPreOffset As Integer = IIf(IsNothing(DataFieldDefinition.RepeatingPreRows), 0, CInt(DataFieldDefinition.RepeatingPreRows))
                            Dim RecordPostOffset As Integer = IIf(IsNothing(DataFieldDefinition.RepeatingPostRows), 0, CInt(DataFieldDefinition.RepeatingPostRows))
                            Dim NumHeaderCells As Integer = 0 + RecordPreOffset + RecordPostOffset

                            If DefiningRange.RowCount > DefiningRange.ColumnCount Then

                                RepeatMethod = "PORT"
                                RepeatCount = DefiningRange.RowCount
                                NumHeaderCells = DefiningRange.ColumnCount

                                If IEDSource.MergedHeader Then

                                    If DefiningRange.ColumnCount > 1 Then

                                        MultiRepeatingHeaders = True

                                    End If

                                End If

                            Else

                                RepeatMethod = "LAND"

                                RepeatCount = DefiningRange.ColumnCount

                                NumHeaderCells = DefiningRange.RowCount

                                If IEDSource.MergedHeader Then

                                    If DefiningRange.RowCount > 1 Then

                                        MultiRepeatingHeaders = True

                                    End If

                                End If

                            End If

                            Dim CellExamineNRD As DevExpress.Spreadsheet.Cell

                            Dim ROPreNRColummns As Integer = 0
                            If IsNumeric(DataFieldDefinition.EditRepNRHereROPreNRLines) Then ROPreNRColummns = CInt(DataFieldDefinition.EditRepNRHereROPreNRLines)

                            Dim DataRangeIndex As Integer = -1

                            For DefNRIndex = (0 - ROPreNRColummns) To RepeatCount - 1

                                DataRangeIndex += 1

                                If RepeatMethod = "PORT" Then

                                    CellExamineNRD = DefiningRange(DefNRIndex, 0)

                                Else

                                    CellExamineNRD = DefiningRange(0, DefNRIndex)

                                End If

                                Dim ReadOnyHeaderEditColumns As Integer = 0


                                'If Not IsNothing(DataFieldDefinition.EditRepNRHereROInitialLines) Then
                                If IsNumeric(DataFieldDefinition.EditRepNRHereROInitialLines) Then ReadOnyHeaderEditColumns = CInt(DataFieldDefinition.EditRepNRHereROInitialLines)
                                'End If


                                'If CellExamineNRD.DisplayText = "" Then GoTo Nextnrds2

                                ColIndex += 1

                                If DataFieldDefinition.HasRule = "TRUE" Then

                                    DataSets(DataSetIndex).HasRules = True
                                    ApplyRule = True

                                Else

                                    ApplyRule = False

                                End If

                                EnsureArrayCapacity(DataSets(DataSetIndex).DataColumns, ColIndex)

                                'Dim CurrColHeading As String
                                Dim HeaderText As String = ""
                                Dim CellExamineHeader As DevExpress.Spreadsheet.Cell
                                Dim HeadCellCount As Integer = 0

                                If NumHeaderCells = 1 Then

                                    HeaderText = CellExamineNRD.DisplayText

                                Else

                                    If RepeatMethod = "PORT" Then

                                        For HeadCellCount = 0 - ROPreNRColummns To NumHeaderCells - 1

                                            If HeadCellCount > 0 Then HeaderText += "vblf"
                                            CellExamineHeader = DefiningRange(DefNRIndex, HeadCellCount)
                                            HeaderText += CellExamineHeader.DisplayText

                                        Next

                                    Else

                                        For HeadCellCount = 0 - ROPreNRColummns To NumHeaderCells - 1

                                            If HeadCellCount > 0 Then HeaderText += "vblf"
                                            CellExamineHeader = DefiningRange(HeadCellCount, DefNRIndex)
                                            HeaderText += CellExamineHeader.DisplayText

                                        Next

                                    End If

                                End If

                                ColHead = DataFieldDefinition.FieldName & " " & CellExamineNRD.DisplayText
                                If Not DataFieldDefinition.Units Is Nothing Then ColHead += vbLf & DataFieldDefinition.Units

                                'CurrColHeading = If(Microsoft.VisualBasic.Right(DataFieldDefinition.FieldName, 4) = "NONE", "", DataFieldDefinition.FieldName & " ") & HeaderText
                                'CurrColHeading = Replace(CurrColHeading, "vblf", vbLf)

                                Dim CalcCol As Boolean = False

                                DataSets(DataSetIndex).DataColumns(ColIndex) = New SheetDataColumn With {
                                    .ColumnTag = New DataColumnTag With {
                                    .ColumnHeading = Replace(ColHead, "vblf", vbLf),
                                    .DataType = CurrDataType,
                                    .IsReadOnly = IIf(DataFieldDefinition.RO = "TRUE", True, False),
                                    .IsCalculated = CalcCol,
                                    .HasActions = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", True, False),
                                    .HasRules = ApplyRule,
                                    .RepeatingNR = DataFieldDefinition.RepeatingNR,
                                    .MinimumWidthChars = ParseMinimumWidthChars(DataFieldDefinition.MinWidthChars),
                                    .ActionNR = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", DataFieldDefinition.RepeatingNR, Nothing),
                                    .ActionData = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", "NRRI", Nothing),
                                    .Units = DataFieldDefinition.Units,
                                    .IsFixed = IIf(DataFieldDefinition.Fixed = "TRUE", True, False),
                                    .ShowSummary = DataFieldDefinition.ShowSummary,
                                    .BandEditDescription = DataFieldDefinition.BandEditDescription,
                                    .MinVal = DataFieldDefinition.MinVal,
                                    .MaxVal = DataFieldDefinition.MaxVal,
                                    .BandID = CellRangeSource.BandID,
                                    .BandTipText = CellRangeSource.BandTipText,
                                    .EditRepNRHere = IIf(UCase(DataFieldDefinition.EditRepNRHere) = "TRUE", True, False),
                                    .EditRepNRHereDataFormat = DataFieldDefinition.EditRepNRHereDataFormat,
                                    .EditRepNRHereEditor = DataFieldDefinition.EditRepNRHereEditor,
                                    .EditRepNRHereComboRepository = DataFieldDefinition.EditRepNRHereComboRepository,
                                    .EditRepNRHereExpansionMethod = DataFieldDefinition.EditRepNRHereExpansionMethod,
                                    .EditRepNRHereRule = DataFieldDefinition.EditRepNRHereRule,
                                    .AllowEditRepNRHereBlanks = True,
                                    .EditRepNRNROrientation = RepeatMethod,
                                    .EditNRIndexPosition = DefNRIndex,
                                    .EditRepNRHereInitialValue = CellExamineNRD.DisplayText,
                                    .IsControlColumn = True,
                                    .RepeatingHeaderText = DataFieldDefinition.RepeatingHeaderText,
                                    .TipText = DataFieldDefinition.TipText,
                                    .RepositaryID = DataFieldDefinition.RepositaryItemID
                                    },
                                    .Index = ColIndex
                                    }


                                If DefNRIndex < ReadOnyHeaderEditColumns Then DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.EditRepNRHereROColumn = True

                                If CellRangeSource.RO = "TRUE" Then DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsReadOnly = True

                                Dim HasControls As Boolean = False
                                If DataFieldDefinition.HasControls = "TRUE" Then HasControls = True

                                For j = 0 To k - 1

                                    CellExamine = DataRange(j + RowOffset, (DataRangeIndex + ColOffset))

                                    EnsureArrayCapacity(
                                        DataSets(DataSetIndex).DataRows(j).DataCells,
                                        ColIndex)

                                    Dim LockCell As Boolean = False

                                    If CellExamine.Protection.Locked Then LockCell = True

                                    DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex) = New CellDataPoint With {
                                        .FoColor = CellExamine.Font.Color,
                                        .BGColor = CellExamine.Fill.BackgroundColor,
                                        .Index = ColIndex,
                                        .IsLocked = LockCell,
                                        .SourceAddress = CellExamine.GetReferenceA1,
                                        .SourceSheet = WSName}

                                    If HasControls Then

                                        If ValidationsSet.HasValidations Then

                                            Dim ChkRange As DevExpress.Spreadsheet.CellRange =
                                                DataRange(j + RowOffset, DataRangeIndex + ColOffset)
                                            Dim TryList As List(Of String) = ValidationsSet.CheckValidation(ChkRange)

                                            If TryList IsNot Nothing Then

                                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).IsValidated = True
                                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).ValidationListID = DataSets(DataSetIndex).AddValList(TryList)
                                                DataSets(DataSetIndex).HasValidations = True

                                            End If

                                        End If

                                    End If

                                    If ApplyRule Then
                                        DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).IsLocked =
                                            IIf(CellExamine.Fill.PatternType = PatternType.Solid, False, True)
                                    End If

                                    If Not CellExamine.Value.IsEmpty Then

                                        If Not DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsCalculated Then 'And Not CellExamine.DisplayText = "0"

                                            DataSets(DataSetIndex).DataRows(j).IsEmpty = False

                                        End If

                                        Select Case CurrDataType



                                            Case "FL"

                                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).StringValue = Left(CellExamine.DisplayText, 10)
                                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).ExtraData = CellExamine.DisplayText

                                            Case "B"

                                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).BoolValue = CInt(CellExamine.Value.NumericValue)

                                            Case "D", "P", "C", "M", "SM", "R"

                                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).RealValue = CDbl(CellExamine.Value.NumericValue)

                                            Case "I", "Y"


                                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).IntValue = CInt(CellExamine.Value.NumericValue)
                                            Case Else

                                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).StringValue = CellExamine.DisplayText
                                        End Select

                                        DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).IsEmpty = False

                                    Else

                                        DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).IsEmpty = True
                                        DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).StringValue = Nothing
                                        DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).BoolValue = Nothing
                                        DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).RealValue = Nothing
                                        DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).IntValue = Nothing

                                    End If

                                Next
Nextnrds2:
                            Next

                        Else 'does not repeat br nr

                            ColIndex += 1

                            EnsureArrayCapacity(DataSets(DataSetIndex).DataColumns, ColIndex)

                            If DataFieldDefinition.HasRule = "TRUE" Then

                                DataSets(DataSetIndex).HasRules = True
                                ApplyRule = True

                            Else

                                ApplyRule = False

                            End If

                            DataSets(DataSetIndex).DataColumns(ColIndex) = New SheetDataColumn With {
                                    .ColumnTag = New DataColumnTag With {
                                    .ColumnHeading = Replace(DataFieldDefinition.FieldName, "vblf", vbLf),
                                    .DataType = CurrDataType,
                                    .Units = DataFieldDefinition.Units,
                                    .IsReadOnly = IIf(DataFieldDefinition.RO = "TRUE", True, False),
                                    .ShowSummary = DataFieldDefinition.ShowSummary,
                                    .IsFixed = IIf(DataFieldDefinition.Fixed = "TRUE", True, False),
                                    .HasRules = ApplyRule,
                                    .RepeatingNR = DataFieldDefinition.RepeatingNR,
                                    .MinimumWidthChars = ParseMinimumWidthChars(DataFieldDefinition.MinWidthChars),
                                    .HasActions = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", True, False),
                                    .IsCalculated = IIf(CellRangeSource.IsCalculated = "TRUE", True, False),
                                    .MinVal = DataFieldDefinition.MinVal,
                                    .MaxVal = DataFieldDefinition.MaxVal,
                                    .BandID = CellRangeSource.BandID,
                                    .BandTipText = CellRangeSource.BandTipText,
                                    .RepositaryID = DataFieldDefinition.RepositaryItemID
                                    },
                                    .Index = ColIndex
                                    }

                            If CellRangeSource.RO = "TRUE" Then

                                DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsReadOnly = True

                            End If

                            Dim HasControls As Boolean = False
                            If DataFieldDefinition.HasControls = "TRUE" Then HasControls = True

                            For j = 0 To k - 1


                                CellExamine = DataRange(j + RowOffset, i + ColOffset)

                                EnsureArrayCapacity(
                                    DataSets(DataSetIndex).DataRows(j).DataCells,
                                    ColIndex)

                                Dim LockCell As Boolean = False
                                If CellExamine.Protection.Locked Then LockCell = True

                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex) = New CellDataPoint With {
                                        .FoColor = CellExamine.Font.Color,
                                        .BGColor = CellExamine.Fill.BackgroundColor,
                                        .Index = ColIndex,
                                        .IsLocked = LockCell,
                                        .SourceAddress = CellExamine.GetReferenceA1,
                                        .SourceSheet = WSName}

                                If HasControls Then

                                    If ValidationsSet.HasValidations Then

                                        Dim ChkRange As DevExpress.Spreadsheet.CellRange = DataRange(j + RowOffset, i + ColOffset)
                                        Dim TryList As List(Of String) = ValidationsSet.CheckValidation(ChkRange)

                                        If TryList IsNot Nothing Then

                                            DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).IsValidated = True
                                            DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).ValidationListID = DataSets(DataSetIndex).AddValList(TryList)
                                            DataSets(DataSetIndex).HasValidations = True

                                        End If

                                    End If

                                End If

                                If ApplyRule Then DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).IsLocked = IIf(CellExamine.Fill.PatternType = PatternType.Solid, False, True)

                                If Not CellExamine.Value.IsEmpty Then

                                    If Not DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsCalculated Then

                                        DataSets(DataSetIndex).DataRows(j).IsEmpty = False

                                    End If

                                    Select Case CurrDataType

                                        Case "S"

                                            DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).StringValue = CellExamine.DisplayText

                                        Case "B"

                                            DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).BoolValue = CInt(CellExamine.Value.NumericValue)

                                        Case "D", "P", "C", "M", "SM", "R"

                                            DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).RealValue = CDbl(CellExamine.Value.NumericValue)

                                        Case "I", "Y"

                                            DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).IntValue = CInt(CellExamine.Value.NumericValue)

                                    End Select

                                    DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).IsEmpty = False

                                Else

                                    DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).IsEmpty = True
                                    DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).StringValue = Nothing
                                    DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).BoolValue = Nothing
                                    DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).RealValue = Nothing
                                    DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).IntValue = Nothing

                                End If

                            Next

                        End If
NextDFD2:
                        i += 1

                    Next

                Next
#End Region
            End If

            'The merge builders add cells column-by-column. Grow their arrays in
            'chunks while building, then restore the exact public array shape once.
            'This avoids copying every prior cell for every newly-created column.
            TrimDataCellCapacity(DataSets(DataSetIndex))

            Dim UsedRows As Integer = 0

            For i = 0 To DataSets(DataSetIndex).RowCount - 1

                If Not DataSets(DataSetIndex).DataRows(i).IsEmpty Then

                    UsedRows += 1

                End If

            Next

            DataSets(DataSetIndex).UsedRows = UsedRows
            DataSets(DataSetIndex).Capacity = DataSets(DataSetIndex).RowCount - UsedRows

            Return DataSets(DataSetIndex)

        End Function

        Private Shared Sub EnsureArrayCapacity(Of T)(ByRef Items() As T,
                                                     ByVal RequiredIndex As Integer)

            If RequiredIndex < 0 Then Return

            Dim CurrentLength As Integer = If(Items Is Nothing, 0, Items.Length)
            If RequiredIndex < CurrentLength Then Return

            Dim NewLength As Integer = If(CurrentLength = 0, 4, CurrentLength)

            While NewLength <= RequiredIndex
                NewLength *= 2
            End While

            ReDim Preserve Items(NewLength - 1)

        End Sub

        Private Shared Sub TrimDataCellCapacity(ByVal DataSet As DataCellRange)

            If DataSet Is Nothing OrElse DataSet.DataRows Is Nothing Then Return

            Dim LastColumnIndex As Integer = -1

            If DataSet.DataColumns IsNot Nothing Then
                For ColumnIndex As Integer = DataSet.DataColumns.Length - 1 To 0 Step -1
                    If DataSet.DataColumns(ColumnIndex).ColumnTag IsNot Nothing Then
                        LastColumnIndex = ColumnIndex
                        Exit For
                    End If
                Next
            End If

            If DataSet.DataColumns IsNot Nothing AndAlso
               DataSet.DataColumns.Length <> LastColumnIndex + 1 Then

                ReDim Preserve DataSet.DataColumns(LastColumnIndex)
            End If

            For Each DataRow As SheetDataRow In DataSet.DataRows
                If DataRow Is Nothing OrElse DataRow.DataCells Is Nothing Then Continue For
                If DataRow.DataCells.Length = LastColumnIndex + 1 Then Continue For

                ReDim Preserve DataRow.DataCells(LastColumnIndex)
            Next

        End Sub

        Private Shared Function ResolveLiveGridSourceAreaReferences(
            ByVal Worksheet As DevExpress.Spreadsheet.Worksheet,
            ByVal Workbook As DevExpress.Spreadsheet.IWorkbook,
            ByVal Source As CellRangeDataSource) As List(Of String)

            Dim Result As New List(Of String)

            If Worksheet Is Nothing OrElse
               Workbook Is Nothing OrElse
               Source Is Nothing Then Return Result

            If Not String.IsNullOrWhiteSpace(Source.LiveGridSourceRanges) Then
                For Each RangeToken As String In Source.LiveGridSourceRanges.Split(";"c)
                    If String.IsNullOrWhiteSpace(RangeToken) Then Continue For
                    Try
                        Result.Add(
                            Worksheet.Range(
                                RangeToken.Trim().Replace("$", String.Empty)).GetReferenceA1())
                    Catch
                        'Ignore invalid optional direct-range fragments.
                    End Try
                Next

                Return Result
            End If

            If String.IsNullOrWhiteSpace(Source.LiveGridSourceName) Then Return Result

            Dim Defined As DevExpress.Spreadsheet.DefinedName =
                Workbook.DefinedNames.GetDefinedName(Source.LiveGridSourceName.Trim())

            If Defined Is Nothing OrElse String.IsNullOrWhiteSpace(Defined.RefersTo) Then Return Result

            Dim CandidateRanges As New List(Of DevExpress.Spreadsheet.CellRange)
            Dim RefersTo As String = Defined.RefersTo.Trim()
            If RefersTo.StartsWith("=", StringComparison.Ordinal) Then
                RefersTo = RefersTo.Substring(1)
            End If

            For Each AreaPart As String In RefersTo.Split(","c)
                Dim BangIndex As Integer = AreaPart.LastIndexOf("!"c)
                If BangIndex < 0 OrElse BangIndex = AreaPart.Length - 1 Then Continue For

                Dim SheetToken As String =
                    AreaPart.Substring(0, BangIndex).Trim().Trim("'"c).Replace("''", "'")
                If Not String.Equals(SheetToken, Worksheet.Name, StringComparison.Ordinal) Then Continue For

                Dim Address As String =
                    AreaPart.Substring(BangIndex + 1).Trim().Replace("$", String.Empty)

                Try
                    CandidateRanges.Add(Worksheet.Range(Address))
                Catch
                    'Ignore non-range name fragments. Workings names use A1 areas.
                End Try
            Next

            If CandidateRanges.Count = 0 Then Return Result

            'Some Workings names retain an earlier block as well as the later block
            'identified by the name suffix, or include a title cell before the
            'actual table. A LiveGrid must project one aligned row block, so choose
            'the last block and then its modal row extent.
            Dim SelectedTopRow As Integer =
                CandidateRanges.Max(Function(Item) Item.TopRowIndex)
            Dim RangesAtSelectedTop As List(Of DevExpress.Spreadsheet.CellRange) =
                CandidateRanges.Where(
                    Function(Item) Item.TopRowIndex = SelectedTopRow).ToList()
            Dim SelectedRowCount As Integer =
                RangesAtSelectedTop.
                    GroupBy(Function(Item) Item.RowCount).
                    OrderByDescending(Function(Group) Group.Count()).
                    ThenByDescending(Function(Group) Group.Key).
                    First().Key
            Dim EligibleRanges As List(Of DevExpress.Spreadsheet.CellRange) =
                RangesAtSelectedTop.Where(
                    Function(Item) Item.RowCount = SelectedRowCount).ToList()

            Dim SelectedAreaIndexes As New List(Of Integer)

            If String.IsNullOrWhiteSpace(Source.LiveGridSourceAreas) Then
                For AreaIndex As Integer = 0 To EligibleRanges.Count - 1
                    SelectedAreaIndexes.Add(AreaIndex)
                Next
            Else
                For Each AreaToken As String In Source.LiveGridSourceAreas.Split(","c)
                    Dim OneBasedAreaIndex As Integer
                    If Integer.TryParse(AreaToken.Trim(), OneBasedAreaIndex) AndAlso
                       OneBasedAreaIndex > 0 AndAlso
                       OneBasedAreaIndex <= EligibleRanges.Count Then

                        SelectedAreaIndexes.Add(OneBasedAreaIndex - 1)
                    End If
                Next
            End If

            Dim LeadingColumnCount As Integer = 0
            Integer.TryParse(Source.LiveGridLeadingColumns, LeadingColumnCount)

            If LeadingColumnCount > 0 AndAlso
               EligibleRanges.Count > 0 AndAlso
               Not SelectedAreaIndexes.Contains(0) Then

                Dim LeadingSource As DevExpress.Spreadsheet.CellRange = EligibleRanges(0)
                LeadingColumnCount = Math.Min(LeadingColumnCount, LeadingSource.ColumnCount)

                Dim LeadingRange As DevExpress.Spreadsheet.CellRange =
                    Worksheet.Range.FromLTRB(
                        LeadingSource.LeftColumnIndex,
                        LeadingSource.TopRowIndex,
                        LeadingSource.LeftColumnIndex + LeadingColumnCount - 1,
                        LeadingSource.BottomRowIndex)

                Result.Add(LeadingRange.GetReferenceA1())
            End If

            For Each AreaIndex As Integer In SelectedAreaIndexes
                Result.Add(EligibleRanges(AreaIndex).GetReferenceA1())
            Next

            Return Result

        End Function

        Public Function RenderIEHTMLCource(ModelID, SetGSID, SetCSID, SetIntSecID, SetISDID) As String

            Dim StrOutput As New StringBuilder(My.Resources.StringTemplates.HTMLFinanceTableHeader)
            StrOutput.Append(My.Resources.StringTemplates.HTMLFinanceTablePrecursor)

            Dim Trans As New AbovoTransaction
            Dim g = ExcelModels(ModelID)

            Dim IEDSource As ISEDatasource = ExcelModels(ModelID).WBStructure.GroupStructures(SetGSID).ChildStructures(SetCSID).InterfaceSections(SetIntSecID).ISDatasources(SetISDID)
            Dim CellExamine As DevExpress.Spreadsheet.Cell

            If IEDSource.DSType = "SimpleGrid" Then

                If IEDSource.DSSource = "CR" Then

                    Dim WSName As String = IEDSource.CellRangeSources(0).WSName
                    Dim CRSource As CellRangeDataSource = IEDSource.CellRangeSources(0)
                    Dim CurrWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets(WSName)

                    Dim DataRange As DevExpress.Spreadsheet.CellRange = CurrWS.Range(CRSource.DataRange)

                    Dim i As Integer, j As Integer
                    Dim TestString As String
                    Dim NumExamine As Double
                    Dim HeightS As String = "80"


                    For i = 0 To DataRange.RowCount - 1

                        StrOutput.Append("<tr class=xl822235 height=").Append(HeightS).Append(" style='height:15.45pt'>")

                        For j = 0 To DataRange.ColumnCount - 1

                            CellExamine = DataRange(i, j)

                            StrOutput.Append("<td height = ").Append(HeightS).Append(" Class=xl882235 style='font-size:11.0pt;")

                            If CellExamine.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center Then

                                StrOutput.Append("text-align:general;")

                            ElseIf CellExamine.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left Then

                                StrOutput.Append("text-align:left;")

                            ElseIf CellExamine.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right Then

                                StrOutput.Append("text-align:right;")

                            End If

                            TestString = CellExamine.DisplayText

                            If IsNumeric(TestString) Then

                                NumExamine = CDbl(TestString)

                                If NumExamine < 0 Then

                                    StrOutput.Append("color:red;")

                                End If


                            End If

                            If CellExamine.Font.Bold = True Then

                                StrOutput.Append("font-weight:800;")

                            Else

                                StrOutput.Append("font-weight:400;")

                            End If

                            StrOutput.Append("background:").Append(CellExamine.Fill.BackgroundColor).Append(";")

                            StrOutput.Append("Text-decoration: none;text-underline-style:none;text-line-through:none;
                                              Font-family: Arial, sans - serif;mso-background-source: auto;mso-pattern:red thin - diag - stripe'>")


                            StrOutput.Append(CellExamine.DisplayText)
                            StrOutput.Append("</td>")
                        Next

                        StrOutput.Append("</tr>")
                        HeightS = "21"

                    Next

                    StrOutput.Append("

                        </table>

                        </div>


                        <!----------------------------->
                        <!--END OF OUTPUT FROM ABOVO SYSTEM-->
                        <!----------------------------->
                        </body>

                        </html>")

                End If

            End If

            Return StrOutput.ToString()

        End Function
        Public Function RenderIEHTMLCourceFromDR(DataRanges As List(Of DevExpress.Spreadsheet.CellRange)) As String

            Dim StrOutput As New StringBuilder(My.Resources.StringTemplates.HTMLFinanceTableHeader)
            StrOutput.Append(My.Resources.StringTemplates.HTMLFinanceTablePrecursor)

            Dim CellExamine As DevExpress.Spreadsheet.Cell
            Dim i As Integer, j As Integer
            Dim TestString As String
            Dim NumExamine As Double
            Dim HeightS As String = "80"

            For Each DataRange In DataRanges



                For i = 0 To DataRange.RowCount - 1

                    StrOutput.Append("<tr class=xl822235 height=").Append(HeightS).Append(" style='height:15.45pt'>")

                    For j = 0 To DataRange.ColumnCount - 1

                        CellExamine = DataRange(i, j)

                        StrOutput.Append("<td height = ").Append(HeightS).Append(" Class=xl882235 style='font-size:9.0pt;")

                        If CellExamine.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center Then

                            StrOutput.Append("text-align:general;")

                        ElseIf CellExamine.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left Then

                            StrOutput.Append("text-align:left;")

                        ElseIf CellExamine.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right Then

                            StrOutput.Append("text-align:right;")

                        End If

                        TestString = CellExamine.DisplayText

                        If IsNumeric(TestString) Then

                            NumExamine = CDbl(TestString)

                            If NumExamine < 0 Then

                                StrOutput.Append("color:red;")

                            End If


                        End If

                        If CellExamine.Font.Bold = True Then

                            StrOutput.Append("font-weight:800;")

                        Else

                            StrOutput.Append("font-weight:400;")

                        End If

                        StrOutput.Append("background:").Append(CellExamine.Fill.BackgroundColor).Append(";")

                        StrOutput.Append("Text-decoration: none;text-underline-style:none;text-line-through:none;
                                              Font-family: Arial, sans - serif;mso-background-source: auto;mso-pattern:red thin - diag - stripe'>")


                        StrOutput.Append(CellExamine.DisplayText)
                        StrOutput.Append("</td>")
                    Next

                    StrOutput.Append("</tr>")
                    HeightS = "21"

                Next

            Next

            StrOutput.Append("

                        </table>

                        </div>


                        <!----------------------------->
                        <!--END OF OUTPUT FROM ABOVO SYSTEM-->
                        <!----------------------------->
                        </body>

                        </html>")



            Return StrOutput.ToString()

        End Function

    End Class

    Public Structure DataChangeEvent

        Public EventID As Integer
        Public ModelID As Integer
        Public Description As String
        Public WSName As String
        Public CellAddress As String
        Public NRAddressing As Boolean
        Public NROrientation As Orientation
        Public TargetNR As String
        Public TargetNRIndex As Integer
        Public OriginalValue As Object
        Public ChangedValue As Object
        Public DataFormat As String
        Public TimeStamp As DateTime
        Public UserName As String

        '0 = Unprocessed, 1 = Processed, 2 = Rejected, 3 = Error, 4 = undone, 5 = Redo, 6 = NoticeOnly
        Public Status As Integer
        Public Sub New(Optional ByVal Status As Integer = 0, Optional ByVal SetNRAddressing As Boolean = False)

            TimeStamp = Now()
            Me.Status = Status
            Me.NRAddressing = SetNRAddressing

        End Sub
    End Structure
    Public Class InstanceDataSet

        Public Controls() As Control

        Sub New(SetModelID As Integer, CSID As Integer)

        End Sub

    End Class


    Partial Public Class AbovoBP

        'Data Structures for AbovoBP

#Region "Data structure creation"
        Public Sub CreateDataRanges()

            CreateCoreDataBindings()
            CreateExistingSOCIOutput()
            CreateExistingStocksOutput()
            InstanceDataObjects()


        End Sub
        Private Sub CreateCoreDataBindings()

            Dim wsGA As DevExpress.Spreadsheet.Worksheet = WBCoreBP.Worksheets("Global Assumptions")
            Dim clCell As DevExpress.Spreadsheet.Cell
            clCell = wsGA.Cells(5, 2)
            InternalCompanyName = CStr(clCell.Value.TextValue)
            BPDetails.CompanyName = InternalCompanyName
            clCell = wsGA.Cells(7, 2)
            BPDetails.StartDate = CDate(clCell.Value.DateTimeValue)

        End Sub

        Public Sub PopulateStock()

            ' On Error GoTo ErrorHandler

            WBCoreBP.DocumentSettings.R1C1ReferenceStyle = True

            Dim wsStock As DevExpress.Spreadsheet.Worksheet = WBCoreBP.Worksheets("Stock Assumptions")

            Dim clCell As DevExpress.Spreadsheet.Cell
            Dim options As New DevExpress.Spreadsheet.RangeDataSourceOptions With {
                .UseFirstRowAsHeader = False,
                .PreserveFormulas = True,
                .SkipHiddenRows = True
            }
            Dim i, sRef As Short
            Dim strCellRef, strCollRef As String

            For i = 0 To internalFixedStockSize - 1

                sRef = 4 + i

                clCell = wsStock(4, sRef)

                If Len(clCell.DisplayText) > 0 Then

                    Stock.StockItems(i).StockDescription = clCell.DisplayText


                    clCell = wsStock(5, sRef)
                    AbovoBP.Stock.StockItems(i).OwnedManaged = clCell.DisplayText


                    clCell = wsStock(6, sRef)
                    AbovoBP.Stock.StockItems(i).SOCIStockType = clCell.DisplayText


                    clCell = wsStock(7, sRef)
                    AbovoBP.Stock.StockItems(i).SOCIRentType = clCell.DisplayText


                    clCell = wsStock(16, sRef)
                    If Len(clCell.DisplayText) > 0 Then AbovoBP.Stock.StockItems(i).CurrentStockNumbers = CInt(clCell.DisplayText)


                    clCell = wsStock(19, sRef)
                    If Len(clCell.DisplayText) > 0 Then AbovoBP.Stock.StockItems(i).PreBPlanStartDateNewBuild = CInt(clCell.DisplayText)


                    clCell = wsStock(21, sRef)
                    If Len(clCell.DisplayText) > 0 Then AbovoBP.Stock.StockItems(i).PreBPlanStartDateDemolitions = CInt(clCell.DisplayText)


                    clCell = wsStock(22, sRef)
                    If Len(clCell.DisplayText) > 0 Then AbovoBP.Stock.StockItems(i).PreBPlanStartDateRTBs = CInt(clCell.DisplayText)

                    clCell = wsStock(27, sRef)
                    If Len(clCell.DisplayText) > 0 Then AbovoBP.Stock.StockItems(i).NewLettings = CInt(clCell.DisplayText)

                    clCell = wsStock(34, sRef)
                    If Len(clCell.DisplayText) > 0 Then AbovoBP.Stock.StockItems(i).NewLetInitialRate = CDbl(clCell.Value.NumericValue)
                    AbovoBP.Stock.StockItems(i).FUpdateStockTotals()

                End If

            Next i

Exiter:

            WBCoreBP.DocumentSettings.R1C1ReferenceStyle = False
            Exit Sub

ErrorHandler:

            WriteLog("Error in populate stock: " & Err.Description, "Populate Stock")
            Resume Exiter

        End Sub
        Public Shared Sub InstanceDataObjects()

            IRVs = New InitialRateVariations


        End Sub

        Public Sub CreateExistingStocksOutput()


            Dim worksheet As DevExpress.Spreadsheet.Worksheet = WBCoreBP.Worksheets("Existing Stock Numbers")
            ' Access the table on the worksheet. 

            Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range("B12:V52")

            ' Specify the data source settings.
            Dim options As New RangeDataSourceOptions With {
                .UseFirstRowAsHeader = False,
                .PreserveFormulas = False,
                .SkipHiddenRows = True,
                .SkipHiddenColumns = True,
                .EditingOptions = DataSourceEditingOptions.ReadOnly
            }

            DSExistingStocksRange = range.GetDataSource(options)

        End Sub

        Public Sub CreateExistingSOCIOutput()


            Dim worksheet As DevExpress.Spreadsheet.Worksheet = WBCoreBP.Worksheets("SOCI Data")
            'Dim worksheet As DevExpress.Spreadsheet.Worksheet = WBCoreBP.Worksheets("SOCILiveData")

            Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range("B5:BA485")
            'Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range("BC1:DA200")

            Dim options As New RangeDataSourceOptions With {
                .UseFirstRowAsHeader = True,
                .PreserveFormulas = False,
                .SkipHiddenRows = True,
                .SkipHiddenColumns = True,
                .EditingOptions = DataSourceEditingOptions.ReadOnly
            }

            DSAnalDataRange = range.GetDataSource(options)

        End Sub

#End Region


    End Class
    Public Class DataObject

        Public DataSets() As DataCellRange
        Public Name As String
        Public DataSetIndex As Integer
        Public DataSetCount As Integer
        Public IsError As Boolean
        Public ErrorText As String


        Public Sub New(SetName As String)


            Name = SetName
            DataSetIndex = -1
            DataSetCount = 0
            IsError = False
            ErrorText = ""

        End Sub
        Public Function AddLine(DataSetindex As Integer, RowAddress As Integer) As AbovoTransaction
            Dim Trans As New AbovoTransaction
            Trans.IntegerReturn = 1
            Return Trans
        End Function

        Public Function GetDataSet(ID As Integer) As DataCellRange
            Return DataSets(ID)
        End Function
        Public Function DeleteLine(DataSetindex As Integer, RowAddress As Integer) As AbovoTransaction

            Dim Trans As New AbovoTransaction
            Trans.IntegerReturn = 1
            Return Trans

        End Function
        Public Function DuplicateLine(DataSetindex As Integer, RowAddress As Integer) As AbovoTransaction

            Dim Trans As New AbovoTransaction
            Trans.IntegerReturn = 1
            Return Trans

        End Function
        Function CombineDataSetsModeA(DataSetSourceColumnsIndex As Integer, DataSetRowsIndex As Integer) As AbovoTransaction

            Dim Trans As New AbovoTransaction
            Trans.IntegerReturn = 1
            Return Trans

        End Function
        Function GetIndexByName(SearchName As String) As Integer

            Dim i As Integer

            For i = 0 To UBound(DataSets)
                If DataSets(i).Name = SearchName Then
                    Return i
                    Exit Function
                End If
            Next

            Return -1

        End Function
        'Function UnionDown(DataSetsToMerge As List(Of Integer), SetName As String) As AbovoTransaction


        '    DataSetIndex += 1
        '    ReDim Preserve DataSets(DataSetIndex)
        '    DataSetCount += 1

        '    DataSets(DataSetIndex) = New DataCellRange(DataSetIndex, ModelID) With {
        '        .Name = SetName,
        '        .IsDirty = False
        '    }

        '    Dim ColCount As Integer = 0

        '    Dim Trans As New AbovoTransaction

        '    Dim SetRows As Boolean = False

        '    Dim j As Integer = 0

        '    For Each DS As Integer In DataSetsToMerge


        '        'add columns
        '        For Each SourceColumn As SheetDataColumn In DataSets(DS).DataColumns

        '            ColCount += 1
        '            ReDim Preserve DataSets(DataSetIndex).DataColumns(ColCount - 1)
        '            DataSets(DataSetIndex).DataColumns(ColCount - 1) = New SheetDataColumn With {
        '                .ColumnTag = New DataColumnTag With {.ColumnHeading = SourceColumn.ColumnTag.ColumnHeading, .DataType = SourceColumn.ColumnTag.DataType},
        '                .Index = ColCount - 1
        '            }

        '        Next

        '        'add rows
        '        If Not SetRows Then

        '            ReDim Preserve DataSets(DataSetIndex).DataRows(DataSets(DS).RowCount - 1)

        '            For j = 0 To DataSets(DS).RowCount - 1

        '                DataSets(DataSetIndex).DataRows(j) = New SheetDataRow With {.Index = j}

        '            Next
        '            DataSets(DataSetIndex).RowCount = DataSets(DS).RowCount
        '            SetRows = True

        '        End If

        '    Next

        '    'add cells
        '    For j = 0 To DataSets(DataSetIndex).RowCount - 1

        '        ReDim DataSets(DataSetIndex).DataRows(j).DataCells(ColCount - 1)

        '    Next

        '    'add data

        '    Dim ColPosition As Integer = 0
        '    Dim LastPosition As Integer = 0


        '    For Each DS As Integer In DataSetsToMerge


        '        'add columns
        '        For Each SourceRow As SheetDataRow In DataSets(DS).DataRows


        '            For Each SourceCell As CellDataPoint In SourceRow.DataCells


        '                DataSets(DataSetIndex).DataRows(SourceRow.Index).DataCells(SourceCell.Index + ColPosition) = New CellDataPoint With {
        '                    .IsEmpty = SourceCell.IsEmpty,
        '                    .Index = SourceCell.Index + ColPosition,
        '                    .SourceAddress = SourceCell.SourceAddress,
        '                    .DataType = SourceCell.DataType
        '                }

        '                Select Case SourceCell.DataType

        '                    Case "S"

        '                        DataSets(DataSetIndex).DataRows(SourceRow.Index).DataCells(SourceCell.Index + ColPosition).StringValue = SourceCell.StringValue

        '                    Case "B"

        '                        DataSets(DataSetIndex).DataRows(SourceRow.Index).DataCells(SourceCell.Index + ColPosition).BoolValue = SourceCell.BoolValue

        '                    Case "N", "P", "M"

        '                        DataSets(DataSetIndex).DataRows(SourceRow.Index).DataCells(SourceCell.Index + ColPosition).RealValue = SourceCell.RealValue

        '                    Case "I", "Y"

        '                        DataSets(DataSetIndex).DataRows(SourceRow.Index).DataCells(SourceCell.Index + ColPosition).IntValue = SourceCell.IntValue


        '                End Select

        '                LastPosition = SourceCell.Index

        '            Next

        '        Next

        '        ColPosition += LastPosition + 1

        '    Next

        '    Trans.IntegerReturn = DataSetIndex
        '    Return Trans

        'End Function
        Function UnionAcross(DataSetsToMerge As List(Of Integer), SetName As String, Optional ByVal MakeFirstReadOnly As Boolean = False) As AbovoTransaction


            'DataSetIndex += 1
            'ReDim Preserve DataSets(DataSetIndex)
            'DataSetCount += 1

            'DataSets(DataSetIndex) = New DataCellRange(DataSetIndex, ModelID) With {
            '    .Name = SetName,
            '    .IsDirty = False
            '}

            'Dim ColCount As Integer = 0

            Dim Trans As New AbovoTransaction

            'Dim SetRows As Boolean = False

            'Dim j As Integer = 0


            'For Each DS As Integer In DataSetsToMerge


            '    'add columns
            '    For Each SourceColumn As SheetDataColumn In DataSets(DS).DataColumns

            '        ColCount += 1
            '        ReDim Preserve DataSets(DataSetIndex).DataColumns(ColCount - 1)

            '        DataSets(DataSetIndex).DataColumns(ColCount - 1) = New SheetDataColumn With {
            '            .ColumnTag = New DataColumnTag With {.ColumnHeading = SourceColumn.ColumnTag.ColumnHeading, .DataType = SourceColumn.ColumnTag.DataType, .IsReadOnly = MakeFirstReadOnly},
            '            .Index = ColCount - 1
            '        }

            '    Next

            '    'add rows
            '    If Not SetRows Then

            '        ReDim Preserve DataSets(DataSetIndex).DataRows(DataSets(DS).RowCount - 1)

            '        For j = 0 To DataSets(DS).RowCount - 1

            '            DataSets(DataSetIndex).DataRows(j) = New SheetDataRow With {.Index = j}

            '        Next

            '        DataSets(DataSetIndex).RowCount = DataSets(DS).RowCount

            '        SetRows = True

            '    End If

            '    MakeFirstReadOnly = False

            'Next

            'DataSets(DataSetIndex).ColCount = ColCount

            ''add cells

            'For j = 0 To DataSets(DataSetIndex).RowCount - 1

            '    ReDim DataSets(DataSetIndex).DataRows(j).DataCells(ColCount - 1)

            'Next

            ''add data

            'Dim ColPosition As Integer = 0
            'Dim LastPosition As Integer = 0


            'For Each DS As Integer In DataSetsToMerge


            '    'add columns
            '    For Each SourceRow As SheetDataRow In DataSets(DS).DataRows


            '        For Each SourceCell As CellDataPoint In SourceRow.DataCells


            '            DataSets(DataSetIndex).DataRows(SourceRow.Index).DataCells(SourceCell.Index + ColPosition) = New CellDataPoint With {
            '                .IsEmpty = SourceCell.IsEmpty,
            '                .Index = SourceCell.Index + ColPosition,
            '                .SourceAddress = SourceCell.SourceAddress,
            '                .DataType = SourceCell.DataType
            '            }

            '            Select Case SourceCell.DataType

            '                Case "S"

            '                    DataSets(DataSetIndex).DataRows(SourceRow.Index).DataCells(SourceCell.Index + ColPosition).StringValue = SourceCell.StringValue

            '                Case "B"

            '                    DataSets(DataSetIndex).DataRows(SourceRow.Index).DataCells(SourceCell.Index + ColPosition).BoolValue = SourceCell.BoolValue

            '                Case "N", "P", "M"

            '                    DataSets(DataSetIndex).DataRows(SourceRow.Index).DataCells(SourceCell.Index + ColPosition).RealValue = SourceCell.RealValue

            '                Case "I", "Y"

            '                    DataSets(DataSetIndex).DataRows(SourceRow.Index).DataCells(SourceCell.Index + ColPosition).IntValue = SourceCell.IntValue

            '            End Select

            '            LastPosition = SourceCell.Index

            '        Next

            '    Next

            '    ColPosition += LastPosition + 1

            'Next

            Trans.IntegerReturn = 1

            Return Trans

        End Function



        Public Function AddHorizontalHeaderDataRange(RangeName As String, FormatMap As String, Titles As List(Of String), Optional ByVal IgnoreLastRows As Integer = 0) As AbovoTransaction

            Dim Trans As New AbovoTransaction

            'Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = AbovoBP.WBCoreBP.DefinedNames.GetDefinedName(RangeName).Range
            'Dim CRTargetWorksheet As DevExpress.Spreadsheet.Worksheet = CRTargetRange.Worksheet
            'Dim clCell As DevExpress.Spreadsheet.Cell

            'DataSetIndex += 1
            'DataSetCount += 1

            'ReDim Preserve DataSets(DataSetIndex)

            'Trans.IntegerReturn = DataSetIndex

            'DataSets(DataSetIndex) = New DataCellRange(DataSetIndex, ModelID) With {
            '    .Name = RangeName,
            '    .RowCount = CRTargetRange.ColumnCount,
            '    .IsDirty = False,
            '    .SourceWorksheet = CRTargetRange.Worksheet.Name,
            '    .FormatMap = FormatMap
            '}

            'Dim i, j As Short
            'Dim c As String

            'Dim Activeposition As Integer = -1

            ''add columns
            'For i = 1 To Len(FormatMap)

            '    c = Mid(FormatMap, i, 1)

            '    If c <> "D" Then

            '        Activeposition += 1

            '        ReDim Preserve DataSets(DataSetIndex).DataColumns(Activeposition)

            '        DataSets(DataSetIndex).DataColumns(Activeposition) = New SheetDataColumn With {
            '            .ColumnTag = New DataColumnTag With {.ColumnHeading = Titles(Activeposition), .DataType = c},
            '            .Index = Activeposition,
            '            .SourceAddress = i - 1
            '        }

            '    End If

            'Next

            'DataSets(DataSetIndex).ColCount = Activeposition + 1

            ''add rows

            'ReDim Preserve DataSets(DataSetIndex).DataRows(DataSets(DataSetIndex).RowCount - 1)

            ''add cells
            'For i = 0 To DataSets(DataSetIndex).RowCount - 1

            '    DataSets(DataSetIndex).DataRows(i) = New SheetDataRow With {.Index = i}
            '    ReDim Preserve DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition)

            'Next i

            'Activeposition = -1

            'For i = 1 To Len(FormatMap)

            '    c = Mid(FormatMap, i, 1)

            '    If c <> "D" Then

            '        Activeposition += 1

            '        For j = 0 To DataSets(DataSetIndex).RowCount - 1

            '            clCell = CRTargetRange(i - 1, j)

            '            DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition) = New CellDataPoint With {
            '                .Index = Activeposition,
            '                .DataType = c,
            '                .SourceAddress = clCell.GetReferenceA1
            '            }

            '            If Len(clCell.DisplayText) > 0 Then

            '                DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).IsEmpty = False


            '                Select Case DataSets(DataSetIndex).DataColumns(Activeposition).ColumnTag.DataType

            '                    Case "S"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).StringValue = clCell.DisplayText

            '                    Case "B"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).BoolValue = CInt(clCell.Value.NumericValue)

            '                    Case "N", "P", "M"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).RealValue = CSng(clCell.Value.NumericValue)

            '                    Case "I", "Y"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).IntValue = CInt(clCell.Value.NumericValue)

            '                End Select

            '            Else

            '                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IsEmpty = True

            '                Select Case DataSets(DataSetIndex).DataColumns(Activeposition).ColumnTag.DataType

            '                    Case "S"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).StringValue = ""

            '                    Case "B"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).BoolValue = -1

            '                    Case "N", "P", "M"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).RealValue = 0

            '                    Case "I", "Y"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).IntValue = 0

            '                End Select

            '            End If

            '        Next

            '    End If

            'Next

            Return Trans

        End Function
        Public Function AddPivottedDataRangeByDefinedColums(RangeName As String, DefinedColumnsDS As Integer, SetDataType As String, Optional ByVal IgnoreLastRows As Integer = 0) As AbovoTransaction

            Dim Trans As New AbovoTransaction
            'Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = AbovoBP.WBCoreBP.DefinedNames.GetDefinedName(RangeName).Range
            'Dim CRTargetWorksheet As DevExpress.Spreadsheet.Worksheet = CRTargetRange.Worksheet
            'Dim clCell As DevExpress.Spreadsheet.Cell

            'DataSetIndex += 1
            'DataSetCount += 1

            'ReDim Preserve DataSets(DataSetIndex)

            'Trans.IntegerReturn = DataSetIndex

            'DataSets(DataSetIndex) = New DataCellRange(DataSetIndex) With {
            '    .Name = RangeName & "_Pivotted",
            '    .RowCount = CRTargetRange.ColumnCount,
            '    .IsDirty = False,
            '    .SourceWorksheet = CRTargetRange.Worksheet.Name,
            '    .FormatMap = SetDataType
            '}

            'Dim i, j As Short
            'Dim NewTitle As String
            'Dim Activeposition As Integer = -1

            ''add columns

            'For i = 0 To DataSets(DefinedColumnsDS).RowCount - 1

            '    If Not DataSets(DefinedColumnsDS).DataRows(i).DataCells(0).IsEmpty Then

            '        Activeposition += 1

            '        ReDim Preserve DataSets(DataSetIndex).DataColumns(Activeposition)

            '        NewTitle = DataSets(DefinedColumnsDS).DataColumns(0).ColumnTag.ColumnHeading & " "

            '        Select Case DataSets(DefinedColumnsDS).DataColumns(0).ColumnTag.DataType

            '            Case "S"

            '                NewTitle += DataSets(DefinedColumnsDS).DataRows(i).DataCells(0).StringValue

            '            Case "B"

            '                NewTitle += DataSets(DefinedColumnsDS).DataRows(i).DataCells(0).BoolValue.ToString

            '            Case "N", "P", "M"

            '                NewTitle += DataSets(DefinedColumnsDS).DataRows(i).DataCells(0).RealValue.ToString

            '            Case "I", "Y"

            '                NewTitle += DataSets(DefinedColumnsDS).DataRows(i).DataCells(0).IntValue.ToString


            '        End Select

            '        DataSets(DataSetIndex).DataColumns(i) = New SheetDataColumn With {
            '            .ColumnTag = New DataColumnTag With {.ColumnHeading = NewTitle, .IsReadOnly = False, .DataType = SetDataType},
            '                .Index = i,
            '                .SourceAddress = -1
            '             }

            '    End If

            'Next

            'DataSets(DataSetIndex).ColCount = Activeposition + 1

            ''add rows

            'ReDim Preserve DataSets(DataSetIndex).DataRows(CRTargetRange.ColumnCount - 1)


            'For i = 0 To DataSets(DataSetIndex).RowCount - 1

            '    DataSets(DataSetIndex).DataRows(i) = New SheetDataRow With {.Index = i}
            '    ReDim Preserve DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition)

            'Next i
            ''add cells

            'Activeposition = -1

            'For i = 0 To DataSets(DataSetIndex).ColCount - 1

            '    If Not DataSets(DefinedColumnsDS).DataRows(i).DataCells(0).IsEmpty Then

            '        Activeposition += 1

            '        For j = 0 To CRTargetRange.ColumnCount - 1

            '            clCell = CRTargetRange(i, j)

            '            DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition) = New CellDataPoint With {
            '                .Index = Activeposition,
            '                .DataType = SetDataType,
            '                .SourceAddress = clCell.GetReferenceA1
            '            }

            '            If Len(clCell.DisplayText) > 0 Then

            '                DataSets(DataSetIndex).DataRows(j).DataCells(i).IsEmpty = False


            '                Select Case SetDataType

            '                    Case "S"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(i).StringValue = clCell.DisplayText

            '                    Case "B"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(i).BoolValue = CInt(clCell.Value.NumericValue)

            '                    Case "N", "P", "M"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(i).RealValue = CSng(clCell.Value.NumericValue)

            '                    Case "I", "Y"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(i).IntValue = CInt(clCell.Value.NumericValue)

            '                End Select

            '            Else

            '                DataSets(DataSetIndex).DataRows(i).DataCells(i).IsEmpty = True

            '                Select Case SetDataType

            '                    Case "S"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(i).StringValue = ""

            '                    Case "B"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(i).BoolValue = 0

            '                    Case "N", "P", "M"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(i).RealValue = 0

            '                    Case "I", "Y"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(i).IntValue = 0



            '                End Select

            '            End If

            '        Next
            '    End If
            'Next

            Return Trans

        End Function
        Public Function AddPivottedDataRangeByTitleList(RangeName As String, FormatMap As String, Titles As List(Of String), Optional ByVal IgnoreLastRows As Integer = 0) As AbovoTransaction

            Dim Trans As New AbovoTransaction
            'Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = AbovoBP.WBCoreBP.DefinedNames.GetDefinedName(RangeName).Range
            'Dim CRTargetWorksheet As DevExpress.Spreadsheet.Worksheet = CRTargetRange.Worksheet
            'Dim clCell As DevExpress.Spreadsheet.Cell

            'DataSetIndex += 1
            'DataSetCount += 1

            'ReDim Preserve DataSets(DataSetIndex)

            'Trans.IntegerReturn = DataSetIndex

            'DataSets(DataSetIndex) = New DataCellRange(DataSetIndex) With {
            '    .Name = RangeName & "_Pivotted",
            '    .RowCount = CRTargetRange.ColumnCount,
            '    .IsDirty = False,
            '    .SourceWorksheet = CRTargetRange.Worksheet.Name,
            '    .FormatMap = FormatMap
            '}

            ''add columns
            'Dim i, j As Short
            'Dim c As String
            'Dim Activeposition As Integer = -1

            'For i = 1 To Len(FormatMap)

            '    c = Mid(FormatMap, i, 1)

            '    If c <> "D" Then

            '        Activeposition += 1

            '        ReDim Preserve DataSets(DataSetIndex).DataColumns(Activeposition)

            '        DataSets(DataSetIndex).DataColumns(Activeposition) = New SheetDataColumn With {
            '            .ColumnTag = New DataColumnTag With {.ColumnHeading = Titles(Activeposition), .DataType = c, .IsReadOnly = False},
            '            .Index = Activeposition,
            '            .SourceAddress = i - 1
            '        }

            '    End If

            'Next

            'DataSets(DataSetIndex).ColCount = Activeposition + 1

            ''add rows
            'ReDim Preserve DataSets(DataSetIndex).DataRows(DataSets(DataSetIndex).RowCount - 1)

            'For i = 0 To DataSets(DataSetIndex).RowCount - 1

            '    DataSets(DataSetIndex).DataRows(i) = New SheetDataRow With {
            '            .Index = i
            '        }
            '    ReDim Preserve DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition)

            'Next

            'Activeposition = -1

            'For i = 1 To Len(FormatMap)


            '    c = Mid(FormatMap, i, 1)

            '    If c <> "D" Then

            '        Activeposition += 1

            '        For j = 0 To CRTargetRange.ColumnCount - 1

            '            clCell = CRTargetRange(i - 1, j)
            '            DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition) = New CellDataPoint With {
            '                .Index = Activeposition,
            '                .DataType = c,
            '                .SourceAddress = clCell.GetReferenceA1
            '            }

            '            If Len(clCell.DisplayText) > 0 Then

            '                DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).IsEmpty = False

            '                Select Case c

            '                    Case "S"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).StringValue = clCell.DisplayText

            '                    Case "B"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).BoolValue = CInt(clCell.Value.NumericValue)

            '                    Case "N", "P", "M"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).RealValue = CSng(clCell.Value.NumericValue)

            '                    Case "I", "Y"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).IntValue = CInt(clCell.Value.NumericValue)

            '                End Select

            '            Else

            '                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IsEmpty = True

            '                Select Case c

            '                    Case "S"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).StringValue = ""

            '                    Case "B"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).BoolValue = 0

            '                    Case "N", "P", "M"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).RealValue = 0

            '                    Case "I", "Y"

            '                        DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition).IntValue = 0

            '                End Select

            '            End If

            '        Next

            '    End If

            'Next

            Return Trans

        End Function
        Public Function AddStandardNamedRange(RangeName As String, FormatMap As String, Titles As List(Of String), Optional ByVal IgnoreLastRows As Integer = 0) As AbovoTransaction

            Dim Trans As New AbovoTransaction

            '            Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = AbovoBP.WBCoreBP.DefinedNames.GetDefinedName(RangeName).Range
            '            Dim CRTargetWorksheet As DevExpress.Spreadsheet.Worksheet = CRTargetRange.Worksheet
            '            Dim clCell As DevExpress.Spreadsheet.Cell

            '            DataSetIndex += 1
            '            ReDim Preserve DataSets(DataSetIndex)
            '            DataSetCount += 1
            '            Trans.IntegerReturn = DataSetIndex

            '            Dim SetRowCount As Integer

            '            SetRowCount = CRTargetRange.RowCount - IgnoreLastRows

            '            DataSets(DataSetIndex) = New DataCellRange(DataSetIndex) With {
            '                .Name = RangeName,
            '                .ColCount = CRTargetRange.ColumnCount,
            '                .RowCount = SetRowCount,
            '                .IsDirty = False,
            '                .SourceWorksheet = CRTargetRange.Worksheet.Name,
            '                .FormatMap = FormatMap
            '            }

            '            ReDim Preserve DataSets(DataSetIndex).DataRows(SetRowCount - 1)

            '            Dim i, j As Short
            '            Dim c As String
            '            Dim Activeposition As Integer = 0

            '            For i = 1 To Len(FormatMap)

            '                c = Mid(FormatMap, i, 1)

            '                If c <> "D" Then

            '                    Activeposition += 1

            '                End If

            '            Next

            '            DataSets(DataSetIndex).ColCount = Activeposition
            '            ReDim Preserve DataSets(DataSetIndex).DataColumns(Activeposition - 1)

            '            Activeposition = -1

            '            For i = 1 To CRTargetRange.ColumnCount

            '                c = Mid(FormatMap, i, 1)

            '                If c <> "D" Then

            '                    Activeposition += 1

            '                    DataSets(DataSetIndex).DataColumns(Activeposition) = New SheetDataColumn With {
            '                        .ColumnTag = New DataColumnTag With {.ColumnHeading = Titles(Activeposition), .DataType = c, .IsReadOnly = False},
            '                        .Index = Activeposition,
            '                        .SourceAddress = i - 1
            '                    }

            '                End If

            '            Next

            '            For i = 0 To DataSets(DataSetIndex).RowCount - 1

            '                DataSets(DataSetIndex).DataRows(i) = New SheetDataRow With {.Index = i}

            '                Activeposition = -1

            '                For j = 0 To CRTargetRange.ColumnCount - 1

            '                    c = Mid(FormatMap, j + 1, 1)
            '                    If c = "D" Then GoTo NextCell

            '                    Activeposition += 1

            '                    ReDim Preserve DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition)

            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition) = New CellDataPoint

            '                    clCell = CRTargetRange(i, j)

            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).SourceAddress = clCell.GetReferenceA1
            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).Index = Activeposition
            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).DataType = DataSets(DataSetIndex).DataColumns(Activeposition).ColumnTag.DataType

            '                    If Len(clCell.DisplayText) > 0 Then

            '                        DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IsEmpty = False

            '                        Select Case DataSets(DataSetIndex).DataColumns(Activeposition).ColumnTag.DataType

            '                            Case "S"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).StringValue = clCell.DisplayText

            '                            Case "B"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).BoolValue = CInt(clCell.Value.NumericValue)

            '                            Case "N", "P", "M"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).RealValue = CSng(clCell.Value.NumericValue)

            '                            Case "I", "Y"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IntValue = CInt(clCell.Value.NumericValue)

            '                            Case "Y"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IntValue = CInt(clCell.Value.NumericValue)

            '                        End Select

            '                    Else

            '                        DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IsEmpty = True

            '                        Select Case DataSets(DataSetIndex).DataColumns(Activeposition).ColumnTag.DataType

            '                            Case "S"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).StringValue = ""

            '                            Case "B"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).BoolValue = -1

            '                            Case "N", "P", "M"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).RealValue = 0

            '                            Case "I", "Y"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IntValue = 0

            '                        End Select

            '                    End If
            'NextCell:
            '                Next

            '            Next

            Return Trans

        End Function

        Public Function AddStandardCellRange(SourceSheet As String, StartCell As String, FormatMap As String, Titles As List(Of String), Optional ByVal IgnoreLastRows As Integer = 0) As AbovoTransaction

            Dim Trans As New AbovoTransaction

            '            Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = AbovoBP.WBCoreBP.Worksheets(SourceSheet).Range()
            '            Dim CRTargetWorksheet As DevExpress.Spreadsheet.Worksheet = CRTargetRange.Worksheet
            '            Dim clCell As DevExpress.Spreadsheet.Cell

            '            DataSetIndex += 1
            '            ReDim Preserve DataSets(DataSetIndex)
            '            DataSetCount += 1
            '            Trans.IntegerReturn = DataSetIndex

            '            Dim SetRowCount As Integer

            '            SetRowCount = CRTargetRange.RowCount - IgnoreLastRows

            '            DataSets(DataSetIndex) = New DataCellRange(DataSetIndex, ModelID) With {
            '                .Name = "Sheet" & SourceSheet & "!" & StartCell,
            '                .ColCount = CRTargetRange.ColumnCount,
            '                .RowCount = SetRowCount,
            '                .IsDirty = False,
            '                .SourceWorksheet = CRTargetRange.Worksheet.Name,
            '                .FormatMap = FormatMap
            '            }

            '            ReDim Preserve DataSets(DataSetIndex).DataRows(SetRowCount - 1)

            '            Dim i, j As Short
            '            Dim c As String
            '            Dim Activeposition As Integer = 0

            '            For i = 1 To Len(FormatMap)

            '                c = Mid(FormatMap, i, 1)

            '                If c <> "D" Then

            '                    Activeposition += 1

            '                End If

            '            Next

            '            DataSets(DataSetIndex).ColCount = Activeposition
            '            ReDim Preserve DataSets(DataSetIndex).DataColumns(Activeposition - 1)

            '            Activeposition = -1

            '            For i = 1 To CRTargetRange.ColumnCount

            '                c = Mid(FormatMap, i, 1)

            '                If c <> "D" Then

            '                    Activeposition += 1

            '                    DataSets(DataSetIndex).DataColumns(Activeposition) = New SheetDataColumn With {
            '                        .Index = Activeposition,
            '                        .ColumnTag = New DataColumnTag With {.ColumnHeading = Titles(Activeposition), .IsReadOnly = False, .DataType = c},
            '                        .SourceAddress = i - 1
            '                    }

            '                End If

            '            Next

            '            Activeposition -= 1

            '            For i = 0 To DataSets(DataSetIndex).RowCount - 1

            '                DataSets(DataSetIndex).DataRows(i) = New SheetDataRow With {.Index = i}

            '                Activeposition = -1

            '                For j = 0 To CRTargetRange.ColumnCount - 1

            '                    c = Mid(FormatMap, j + 1, 1)
            '                    If c = "D" Then GoTo NextCell

            '                    Activeposition += 1

            '                    ReDim Preserve DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition)

            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition) = New CellDataPoint

            '                    clCell = CRTargetRange(i, j)

            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).SourceAddress = clCell.GetReferenceA1
            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).Index = Activeposition
            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).DataType = DataSets(DataSetIndex).DataColumns(Activeposition).ColumnTag.DataType

            '                    If Len(clCell.DisplayText) > 0 Then

            '                        DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IsEmpty = False

            '                        Select Case DataSets(DataSetIndex).DataColumns(Activeposition).ColumnTag.DataType

            '                            Case "S"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).StringValue = clCell.DisplayText

            '                            Case "B"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).BoolValue = CInt(clCell.Value.NumericValue)

            '                            Case "N", "P", "M"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).RealValue = CSng(clCell.Value.NumericValue)

            '                            Case "I", "Y"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IntValue = CInt(clCell.Value.NumericValue)

            '                            Case "Y"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IntValue = CInt(clCell.Value.NumericValue)

            '                        End Select

            '                    Else

            '                        DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IsEmpty = True

            '                        Select Case DataSets(DataSetIndex).DataColumns(Activeposition).ColumnTag.DataType

            '                            Case "S"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).StringValue = ""

            '                            Case "B"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).BoolValue = -1

            '                            Case "N", "P", "M"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).RealValue = 0

            '                            Case "I", "Y"

            '                                DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IntValue = 0

            '                        End Select

            '                    End If
            'NextCell:
            '                Next

            '            Next

            Return Trans

        End Function
        Public Function GetISEDataStructure(ModelID, SetGSID, SetCSID, SetIntSecID, SetISDID) As DataCellRange

            MsgBox("CallContext")

            Dim Trans As New AbovoTransaction
            Dim IEDSource As ISEDatasource = ExcelModels(ModelID).WBStructure.GroupStructures(SetGSID).ChildStructures(SetCSID).InterfaceSections(SetIntSecID).ISDatasources(SetISDID)
            Dim CellExamine As DevExpress.Spreadsheet.Cell
            Dim CellExamineRight As DevExpress.Spreadsheet.Cell

            If IEDSource.DSType = "SimpleGrid" Then

                If IEDSource.DSSource = "CR" Then

                    Dim WSName As String = IEDSource.CellRangeSources(0).WSName
                    Dim CRSource As CellRangeDataSource = IEDSource.CellRangeSources(0)
                    Dim CurrWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets(WSName)

                    DataSetIndex += 1
                    ReDim Preserve DataSets(DataSetIndex)

                    DataSets(DataSetIndex) = New DataCellRange(DataSetIndex, ModelID) With {
                        .Name = "Sheet" & WSName & "-IEDS-" & WSName,
                        .IsDirty = False,
                        .SourceWorksheet = WSName,
                        .RO = IIf(IEDSource.RO = "TRUE", True, False),
                        .FormatMap = CRSource.DataFieldDefinitions(0).DataFormat
                    }

                    Trans.IntegerReturn = DataSetIndex
                    Dim ColHead As String
                    Dim ColHeaderRange As DevExpress.Spreadsheet.CellRange
                    Dim RowHeaderRange As DevExpress.Spreadsheet.CellRange

                    Dim DataRange As DevExpress.Spreadsheet.CellRange = CurrWS.Range(CRSource.DataRange)

                    If IEDSource.Pivot = "TRUE" Then

                        If CRSource.RowsDefinedBy = "NR" Then

                            ColHeaderRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.RowsDefinedByData).Range

                        Else

                            ColHeaderRange = CurrWS.Range(CRSource.RowsDefinedByData)

                        End If

                        If CRSource.ColsDefinedBy = "NR" Then

                            RowHeaderRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.ColsDefinedByData).Range

                        Else

                            RowHeaderRange = CurrWS.Range(CRSource.ColsDefinedByData)

                        End If

                        DataSets(DataSetIndex).DataColumns(0) = New SheetDataColumn With {
                                .ColumnTag = New DataColumnTag With {.ColumnHeading = "Description", .DataType = "S"},
                                .Index = 0
                                }

                        Dim x As Integer

                        For x = 1 To ColHeaderRange.RowCount

                            CellExamine = ColHeaderRange(x, 0)
                            ReDim Preserve DataSets(DataSetIndex).DataColumns(x)
                            ColHead = CellExamine.DisplayText

                            If ColHeaderRange.ColumnCount = 2 Then

                                CellExamineRight = ColHeaderRange(x, 1)
                                ColHead = ColHead & " (" & CellExamineRight.DisplayText & ")"

                            ElseIf ColHeaderRange.ColumnCount = 3 Then

                                CellExamineRight = ColHeaderRange(x, 1)
                                ColHead = ColHead & " (" & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(x, 2)
                                ColHead = ColHead & CellExamineRight.DisplayText & ")"

                            End If

                            DataSets(DataSetIndex).DataColumns(x) = New SheetDataColumn With {
                                    .ColumnTag = New DataColumnTag With {.ColumnHeading = ColHead, .DataType = CRSource.DataFieldDefinitions(0).DataFormat},
                                    .Index = x
                                    }

                            If CRSource.DataFieldDefinitions(0).ShowSummary IsNot Nothing Then

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.ShowSummary = CRSource.DataFieldDefinitions(0).ShowSummary

                            Else

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.ShowSummary = "FALSE"

                            End If

                            If CRSource.DataFieldDefinitions(0).MinVal = Nothing Then

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MinVal = "FALSE"

                            Else

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MinVal = CRSource.DataFieldDefinitions(0).MinVal

                            End If

                            If CRSource.DataFieldDefinitions(0).MaxVal = Nothing Then

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MaxVal = "FALSE"

                            Else

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MaxVal = CRSource.DataFieldDefinitions(0).MinVal

                            End If
                        Next

                        DataSets(DataSetIndex).ColCount = x + 1
                        Dim RowHead As String

                        For x = 0 To RowHeaderRange.ColumnCount

                            CellExamine = ColHeaderRange(0, x)
                            ReDim Preserve DataSets(DataSetIndex).DataRows(x)
                            RowHead = "<B>" + CellExamine.DisplayText + "</B>"

                            If ColHeaderRange.RowCount = 2 Then

                                CellExamineRight = ColHeaderRange(1, x)
                                RowHead = RowHead & "<br>(" & CellExamineRight.DisplayText & ")"

                            ElseIf ColHeaderRange.ColumnCount = 3 Then

                                CellExamineRight = ColHeaderRange(1, x)
                                RowHead = RowHead & "<br>(" & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(2, x)
                                RowHead = RowHead & CellExamineRight.DisplayText & ")"

                            End If

                            DataSets(DataSetIndex).DataRows(x) = New SheetDataRow With {
                                    .Index = x
                                    }

                            ReDim DataSets(DataSetIndex).DataRows(x).DataCells(0)
                            DataSets(DataSetIndex).DataRows(x).DataCells(0) = New CellDataPoint With {
                                .Index = 0,
                                .DataType = "S",
                                .StringValue = RowHead
                            }

                        Next

                        DataSets(DataSetIndex).RowCount = RowHeaderRange.ColumnCount

                        Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.Worksheets(WSName).Range()
                        Dim clCell As DevExpress.Spreadsheet.Cell
                        Dim i As Integer, j As Integer

                        For i = 0 To RowHeaderRange.ColumnCount - 1

                            For j = 1 To ColHeaderRange.RowCount

                                clCell = DataRange(i, j)
                                ReDim Preserve DataSets(DataSetIndex).DataRows(i).DataCells(i)
                                DataSets(DataSetIndex).DataRows(i).DataCells(j) = New CellDataPoint
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).SourceAddress = clCell.GetReferenceA1
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).Index = j
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).DataType = DataSets(DataSetIndex).DataColumns(j).ColumnTag.DataType
                                Select Case DataSets(DataSetIndex).DataColumns(j).ColumnTag.DataType

                                    Case "S"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).StringValue = clCell.DisplayText

                                    Case "B"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).BoolValue = CInt(clCell.Value.NumericValue)

                                    Case "N", "P", "M", "SM", "R"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).RealValue = CDbl(clCell.Value.NumericValue)

                                    Case "I", "Y", "D"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).IntValue = CInt(clCell.Value.NumericValue)

                                    Case "Y"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).IntValue = CInt(clCell.Value.NumericValue)

                                End Select

                            Next

                        Next

                    Else 'not pivoted

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        If CRSource.RowsDefinedBy = "NR" Then

                            ColHeaderRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.ColsDefinedByData).Range

                        Else

                            ColHeaderRange = CurrWS.Range(CRSource.ColsDefinedByData)

                        End If

                        If CRSource.ColsDefinedBy = "NR" Then

                            RowHeaderRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.RowsDefinedByData).Range

                        Else

                            RowHeaderRange = CurrWS.Range(CRSource.RowsDefinedByData)

                        End If

                        DataSets(DataSetIndex).DataColumns(0) = New SheetDataColumn With {
                                .ColumnTag = New DataColumnTag With {.ColumnHeading = "Description", .DataType = "S"},
                                .Index = 0
                                }

                        Dim x As Integer

                        For x = 1 To ColHeaderRange.RowCount

                            CellExamine = ColHeaderRange(x, 0)
                            ReDim Preserve DataSets(DataSetIndex).DataColumns(x)
                            ColHead = CellExamine.DisplayText

                            If ColHeaderRange.ColumnCount = 2 Then

                                CellExamineRight = ColHeaderRange(x, 1)
                                ColHead = ColHead & " (" & CellExamineRight.DisplayText & ")"

                            ElseIf ColHeaderRange.ColumnCount = 3 Then

                                CellExamineRight = ColHeaderRange(x, 1)
                                ColHead = ColHead & " (" & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(x, 2)
                                ColHead = ColHead & CellExamineRight.DisplayText & ")"

                            End If

                            DataSets(DataSetIndex).DataColumns(x) = New SheetDataColumn With {
                                    .ColumnTag = New DataColumnTag With {.ColumnHeading = ColHead, .DataType = CRSource.DataFieldDefinitions(0).DataFormat},
                                    .Index = x
                                    }

                            If CRSource.DataFieldDefinitions(0).ShowSummary IsNot Nothing Then

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.ShowSummary = CRSource.DataFieldDefinitions(0).ShowSummary

                            Else

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.ShowSummary = "FALSE"

                            End If

                            If CRSource.DataFieldDefinitions(0).MinVal = Nothing Then

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MinVal = "FALSE"

                            Else

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MinVal = CRSource.DataFieldDefinitions(0).MinVal

                            End If

                            If CRSource.DataFieldDefinitions(0).MaxVal = Nothing Then

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MaxVal = "FALSE"

                            Else

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MaxVal = CRSource.DataFieldDefinitions(0).MinVal

                            End If

                        Next

                        DataSets(DataSetIndex).ColCount = x + 1
                        Dim RowHead As String

                        For x = 0 To RowHeaderRange.ColumnCount

                            CellExamine = ColHeaderRange(0, x)
                            ReDim Preserve DataSets(DataSetIndex).DataRows(x)
                            RowHead = CellExamine.DisplayText

                            If ColHeaderRange.RowCount = 2 Then

                                CellExamineRight = ColHeaderRange(1, x)
                                RowHead = RowHead & " (" & CellExamineRight.DisplayText & ")"

                            ElseIf ColHeaderRange.ColumnCount = 3 Then

                                CellExamineRight = ColHeaderRange(1, x)
                                RowHead = RowHead & " (" & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(2, x)
                                RowHead = RowHead & CellExamineRight.DisplayText & ")"

                            End If

                            DataSets(DataSetIndex).DataRows(x) = New SheetDataRow With {
                                    .Index = x
                                    }

                            ReDim DataSets(DataSetIndex).DataRows(x).DataCells(0)
                            DataSets(DataSetIndex).DataRows(x).DataCells(0) = New CellDataPoint With {
                                .Index = 0,
                                .DataType = "S",
                                .StringValue = RowHead
                            }

                        Next

                        DataSets(DataSetIndex).RowCount = RowHeaderRange.ColumnCount

                        Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.Worksheets(WSName).Range()
                        Dim clCell As DevExpress.Spreadsheet.Cell
                        Dim i As Integer, j As Integer

                        For i = 0 To RowHeaderRange.ColumnCount - 1

                            For j = 1 To ColHeaderRange.RowCount

                                clCell = DataRange(i, j)
                                ReDim Preserve DataSets(DataSetIndex).DataRows(i).DataCells(i)
                                DataSets(DataSetIndex).DataRows(i).DataCells(j) = New CellDataPoint
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).SourceAddress = clCell.GetReferenceA1
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).Index = j
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).DataType = DataSets(DataSetIndex).DataColumns(j).ColumnTag.DataType

                                Select Case DataSets(DataSetIndex).DataColumns(j).ColumnTag.DataType

                                    Case "S"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).StringValue = clCell.DisplayText

                                    Case "B"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).BoolValue = CInt(clCell.Value.NumericValue)

                                    Case "N", "P", "M", "SM", "R"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).RealValue = CDbl(clCell.Value.NumericValue)

                                    Case "I", "Y", "D"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).IntValue = CInt(clCell.Value.NumericValue)

                                    Case "Y"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).IntValue = CInt(clCell.Value.NumericValue)

                                End Select

                            Next

                        Next

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    End If




                End If

            End If

            Return DataSets(DataSetIndex)

        End Function
        Public Class GridViewTag

            Public ModelID As Integer
            Public DSID As Integer
            Public DataSet As DataCellRange
            Public InManualReizeMode As Boolean = False
            Public HaveProcessedColumns As Boolean = False
            Public IsLiveGrid As Boolean = False
            Public LiveGridWorksheet As String
            Public LiveGridRange As String
            Public LiveGridSourceRows As List(Of Integer)
            Public LiveGridSourceColumns As List(Of Integer)

        End Class


        Public Class ExcelDatasetCreator

            Public ModelID As Integer

            Public DSCreationID As Integer

            Public RowDefinitionType As String
            Public RowDefinitionNRSource As String
            Public RowsFixed As Boolean
            Public RowDefinitionStart As String
            Public RowDefinitionEnd As String
            Public MergeRowDefs As Boolean
            Public RowDefMergeFormat As String

            Public ColDefinitionType As String
            Public ColDefinitionNRSource As String
            Public ColFixed As Boolean
            Public ColDefinitionStart As String
            Public ColDefinitionEnd As String
            Public MergeColDefs As Boolean
            Public ColDefMergeFormat As String

            Public DataFormat As String
            Public DataStart As String

        End Class

        Public Class CellDataPoint

            Public Index As Integer
            Public BGColor As Color
            Public FoColor As Color
            Public FontBold As Boolean
            Public DataType As String
            Public StringValue As String
            Public BoolValue As Integer
            Public RealValue As Double
            Public IntValue As Integer
            Public IsEmpty As Boolean
            Public SourceAddress As String
            Public SourceSheet As String
            Public IsDirty As Boolean = False
            Public IsLocked As Boolean = False
            Public ExtraData As Object
            Public ValidationListID As Integer
            Public IsValidated As Boolean = False

        End Class

        Structure SheetDataColumn

            Public Index As Integer
            Public DataAddress As Integer
            Public SourceAddress As Integer
            Public ColumnTag As DataColumnTag

        End Structure

        Friend Shared Function ParseMinimumWidthChars(ByVal WidthChars As String) As Integer

            If String.IsNullOrWhiteSpace(WidthChars) Then Return 0

            Dim RetWidth As Integer = 0

            If Integer.TryParse(WidthChars.Trim(), RetWidth) Then
                Return Math.Max(0, RetWidth)
            End If

            Return 0

        End Function

        Class DataColumnTag

            Public DataType As String
            Public ColumnHeading As String
            Public IsReadOnly As Boolean
            Public IsCalculated As Boolean
            Public ShowSummary As String
            Public MinVal As String
            Public MaxVal As String
            Public RepositaryID As String
            Public TipText As String
            Public ActiveColumnName As String
            Public IsControlColumn As Boolean = False
            Public IsFixed As Boolean = False
            Public HasActions As Boolean = False
            Public HasRules As Boolean = False
            Public ActionNR As String
            Public ActionData As String
            Public RepeatingNR As String

            'Optional minimum display width from Structure.xml, expressed in
            'characters. A value <= 0 means "no configured minimum".
            Public MinimumWidthChars As Integer = 0

            Public ButtonObjectState As ObjectState
            Public DefaultColumnWidth As Integer
            Public DontDrawCellHeader As Boolean = False
            Public HasComboEdit As Boolean = False
            Public ShowDefaultmask As Integer

            Public IsDummyColumn As Boolean = False

            Public EditRepNRHere As Boolean
            Public EditRepNRHereDataFormat As String
            Public EditRepNRHereEditor As String
            Public EditRepNRHereComboRepository As String
            Public EditRepNRHereExpansionMethod As String
            Public EditRepNRHereRule As String
            Public EditRepNRNROrientation As String
            Public EditNRIndexPosition As Integer = -1
            Public EditRepNRHereROColumn As Boolean = False
            Public EditRepNRHereInitialValue As Object
            Public AllowEditRepNRHereBlanks As Boolean = True

            Public HasIncolumnButton As Boolean = False
            Public HasIncolumnEditor As Boolean = False
            Public InColumnEditorCombo As AbovoDEHeaderComboBox = Nothing
            Public InColumnEditorDate As AbovoDEHeaderDateBox = Nothing

            Public ExtendedColumnWidth As Integer
            Public WidthSet As Boolean = False
            Public BandID As String
            Public BandTipText As String
            Public BandEditDescription As String
            Public RepeatingHeaderText As String
            Public Units As String
            Public LabelText As String
            Public DefIncrement As String
            Public FormatString As String
            Public HasControls As Boolean = False
            Public DefaultTextEditor As RepositoryItemTextEdit

            Public ColWidthMultiplier As Double = 1
            Public ColumnWidthFixed As Boolean = False

        End Class
        Class SingleCellDataTag

            Public MinVal As Double
            Public MaxVal As Double
            Public MinValSet As Boolean
            Public MaxValSet As Boolean
            Public DataType As String
            Public TargetWorksheet As DevExpress.Spreadsheet.Worksheet
            Public TargetCell As String
            Public Label As String
            Public IsCalculated As Boolean
            Public TipText As String

        End Class

        Class SheetDataRow

            Public Index As Integer
            Public IsControlRow As Boolean = False
            Public IsEmpty As Boolean = True
            Public IsSpacerRow As Boolean = False
            Public DataCells() As CellDataPoint

        End Class

        Public Class DataCellRange

            Implements ICloneable

            Public RowCount As Integer
            Public ColCount As Integer
            Public Name As String
            Public IsDirty As Boolean
            Public Index As Integer
            Public ModelID As Integer
            Public FormatMap As String
            Public RO As Boolean
            Public NewDataModel As String
            Public RowExpandsByModel As String
            Public RowExpandByNR As String
            Public ColExpandByNR As String
            Public DefaultDataNR As String
            Public RepeatingNR As String
            Public DataRange As String
            Public LiveGridSourceName As String
            Public LiveGridSourceRanges As String
            Public LiveGridSourceAreaReferences As New List(Of String)
            Public LiveGridHeaderRows As String
            Public Capacity As Integer = 0
            Public UsedRows As Integer = 0
            Public DataRows() As SheetDataRow
            Public DataColumns() As SheetDataColumn
            Public SourceWorksheet As String
            Public HasBands As Boolean
            Public HasCalcs As Boolean = False
            Public HasRules As Boolean = False
            Public SkipLastRecords As Integer = 0
            Public HasValidations As Boolean = False
            Public ValidationLists() As List(Of String)
            Public ValListCount As Integer = -1
            Private ReadOnly ValidationListIndexes As New Dictionary(Of List(Of String), Integer)

            Sub New(SetIndex As Integer, SetModelID As Integer)

                ReDim DataRows(-1)
                ReDim DataColumns(-1)

                Index = SetIndex
                ModelID = SetModelID

            End Sub
            Public Function AddValList(AddedList As List(Of String)) As Integer

                If AddedList Is Nothing Then Return -1

                Dim ExistingIndex As Integer
                If ValidationListIndexes.TryGetValue(AddedList, ExistingIndex) Then
                    Return ExistingIndex
                End If

                ValListCount += 1
                ReDim Preserve ValidationLists(ValListCount)
                ValidationLists(ValListCount) = AddedList
                ValidationListIndexes.Add(AddedList, ValListCount)
                Return ValListCount

            End Function

            Sub UpdateCalcs()

                If Not HasCalcs Then Return

                Dim DP As CellDataPoint
                Dim SourceCell As DevExpress.Spreadsheet.Cell
                Dim WB As DevExpress.Spreadsheet.IWorkbook = GetWorkBook(ModelID)
                If WB Is Nothing Then Return

                Dim WorksheetCache As New Dictionary(Of String, DevExpress.Spreadsheet.Worksheet)(StringComparer.OrdinalIgnoreCase)

                For Each SheetDataColumn In DataColumns

                    If SheetDataColumn.ColumnTag.IsCalculated Then

                        For Each SheetDataRow In DataRows

                            DP = SheetDataRow.DataCells(SheetDataColumn.Index)
                            SourceCell = GetCachedWorksheet(WB, WorksheetCache, DP.SourceSheet).Cells(DP.SourceAddress)

                            Select Case SheetDataColumn.ColumnTag.DataType

                                Case "S"

                                    SheetDataRow.DataCells(SheetDataColumn.Index).StringValue = SourceCell.DisplayText

                                Case "B"
                                    SheetDataRow.DataCells(SheetDataColumn.Index).BoolValue = CInt(SourceCell.Value.NumericValue)


                                Case "D", "P", "C", "M"
                                    SheetDataRow.DataCells(SheetDataColumn.Index).RealValue = CDbl(SourceCell.Value.NumericValue)


                                Case "I", "Y"

                                    SheetDataRow.DataCells(SheetDataColumn.Index).IntValue = CInt(SourceCell.Value.NumericValue)

                            End Select

                            'DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).IsEmpty = False




                        Next


                    End If


                Next



            End Sub
            Sub UpdateLocks()

                If Not HasRules Then Return

                Dim DP As CellDataPoint
                Dim SourceCell As DevExpress.Spreadsheet.Cell
                Dim WB As DevExpress.Spreadsheet.IWorkbook = ExcelModels(ModelID).WB
                If WB Is Nothing Then Return

                Dim WorksheetCache As New Dictionary(Of String, DevExpress.Spreadsheet.Worksheet)(StringComparer.OrdinalIgnoreCase)

                For Each SheetDataColumn In DataColumns

                    If SheetDataColumn.ColumnTag.HasRules Then

                        For Each SheetDataRow In DataRows

                            If SheetDataRow.IsControlRow Or SheetDataRow.IsSpacerRow Then GoTo nextDP
                            DP = SheetDataRow.DataCells(SheetDataColumn.Index)
                            SourceCell = GetCachedWorksheet(WB, WorksheetCache, DP.SourceSheet).Cells(DP.SourceAddress)
                            SheetDataRow.DataCells(SheetDataColumn.Index).IsLocked =
                                SourceCell.Fill.PatternType <> PatternType.Solid
nextDP:
                        Next

                    End If

                Next

            End Sub

            Private Shared Function GetCachedWorksheet(
                ByVal WB As DevExpress.Spreadsheet.IWorkbook,
                ByVal WorksheetCache As Dictionary(Of String, DevExpress.Spreadsheet.Worksheet),
                ByVal WorksheetName As String) As DevExpress.Spreadsheet.Worksheet

                Dim WS As DevExpress.Spreadsheet.Worksheet = Nothing

                If Not WorksheetCache.TryGetValue(WorksheetName, WS) Then
                    WS = WB.Worksheets(WorksheetName)
                    WorksheetCache.Add(WorksheetName, WS)
                End If

                Return WS
            End Function

            Public Function Clone() As Object Implements ICloneable.Clone
                Return Me.MemberwiseClone
            End Function

        End Class



    End Class
    Public Class AbovoUnboundSource

        Inherits DevExpress.Data.UnboundSource
        Implements ICloneable

        Public UBSIndex As Integer
        Public UBSTag As AbovoUnboundSourceTag
        Public AttachedGrid As DevExpress.XtraGrid.GridControl
        Public AttachedVertGrid As VGridControl
        Public InBandedMode As Boolean = False
        Public InVertMode As Boolean = False
        Public ActiveGridView As DevExpress.XtraGrid.Views.Grid.GridView
        Public ActiveGridBandedView As DevExpress.XtraGrid.Views.BandedGrid.BandedGridView

        Public Sub New(SetUBSIndex As Integer, SetTag As AbovoUnboundSourceTag)
            UBSIndex = SetUBSIndex
            UBSTag = SetTag

        End Sub
        Public Function Clone() As Object Implements ICloneable.Clone

            Return Me.MemberwiseClone

        End Function
        Class AbovoUnboundSourceTag

            Public DSIndex As Integer
            Public ModelID As Integer
            Public GSID As Integer
            Public CSID As Integer
            Public RO As Boolean
            Public HasCalcs As Boolean = False
            Public IsLiveGrid As Boolean = False
            Public LiveGridWorksheet As String
            Public LiveGridSourceRows As List(Of Integer)
            Public LiveGridSourceColumns As List(Of Integer)
            Public GridTag As GridViewTag
            Public SectionTag As InterfaceSectionTag

        End Class

    End Class

    Public Class InterfaceSectionTag

        Public ModelID As Integer
        Public GSID As Integer
        Public CSID As Integer
        Public IntSecID As Integer
        Public Desription As String

    End Class
    Public Class AbovoRangeDataSource

        Implements ICloneable

        Public RangeDS As RangeDataSource
        Public DSTag As AbovoRangeDataSourceTag

        Private _ModelID As Integer
        Private _worksheet As String
        Private _range As String
        Public Property ModelID As Integer
            Get
                Return _ModelID
            End Get
            Set(value As Integer)
                _ModelID = value
            End Set
        End Property
        Public Property Worksheet As String
            Get
                Return _worksheet
            End Get
            Set(value As String)
                _worksheet = value
            End Set
        End Property
        Public Property Range As String
            Get
                Return _range
            End Get
            Set(value As String)
                _range = value
            End Set
        End Property
        Public Sub New()



        End Sub
        Public Function GetRangeDS(SetTag As AbovoRangeDataSourceTag) As RangeDataSource

            '''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''


            Dim worksheet As DevExpress.Spreadsheet.Worksheet = ExcelModels(SetTag.ModelID).WB.Worksheets(SetTag.Worksheet)

            Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range(SetTag.DataRange)
            Dim CD As New Abovo.ColumnDetector(SetTag.ColList, SetTag.TypeList)
            Dim RDSOptions As New RangeDataSourceOptions With {
            .UseFirstRowAsHeader = False,
            .PreserveFormulas = False,
            .SkipHiddenRows = True,
            .SkipHiddenColumns = True,
            .DataSourceColumnTypeDetector = CD,
            .EditingOptions = DataSourceEditingOptions.ReadOnly
        }

            RangeDS = range.GetDataSource(RDSOptions)

            Return RangeDS

        End Function
        Public Function Clone() As Object Implements ICloneable.Clone

            Return Me.MemberwiseClone

        End Function
        Class AbovoRangeDataSourceTag

            Public DSIndex As Integer
            Public ModelID As Integer
            Public GSID As Integer
            Public CSID As Integer
            Public Worksheet As String
            Public DataRange As String
            Public RO As Boolean
            Public HasCalcs As Boolean = False
            Public TypeList As List(Of String)
            Public ColList As List(Of String)

        End Class

    End Class
    Class DataValidationsSet


        Private ReadOnly ValidatedRanges As New List(Of ValidatedRange)
        Private LastMatchedRangeIndex As Integer = -1
        Private HasOverlappingRanges As Boolean = False
        Public HasValidations As Boolean = False

        Public Sub New(WS As DevExpress.Spreadsheet.Worksheet)

            Dim validations As DevExpress.Spreadsheet.DataValidationCollection = WS.DataValidations

            If validations IsNot Nothing Then

                If validations.Count > 0 Then

                    For Each validation As DevExpress.Spreadsheet.DataValidation In validations

                        If validation.ValidationType = DevExpress.Spreadsheet.DataValidationType.List Then
                            Dim CriList As List(Of String) = GetValidationChoices(validation, WS)

                            If CriList.Count > 0 Then
                                AddValidatedRange(validation.Range, CriList)
                                HasValidations = True
                            End If
                        End If

                    Next validation

                End If


            End If


        End Sub

        Private Shared Function GetValidationChoices(
            Validation As DevExpress.Spreadsheet.DataValidation,
            WS As DevExpress.Spreadsheet.Worksheet) As List(Of String)

            Dim Choices As New List(Of String)
            Dim Criteria As DevExpress.Spreadsheet.ValueObject = Validation.Criteria

            If Criteria Is Nothing OrElse Criteria.IsEmpty Then Return Choices

            If Criteria.IsRange Then
                AddRangeChoices(Criteria.RangeValue, Choices)
                Return Choices
            End If

            If Criteria.IsFormula Then
                Dim Formula As String = Criteria.FormulaInvariant
                Dim SourceRange As DevExpress.Spreadsheet.CellRange = ResolveValidationRange(WS, Formula)

                If SourceRange IsNot Nothing Then
                    AddRangeChoices(SourceRange, Choices)
                    Return Choices
                End If
            End If

            If Criteria.IsText AndAlso Not String.IsNullOrWhiteSpace(Criteria.TextValue) Then
                For Each Item As String In Criteria.TextValue.Split(","c)
                    Dim Choice As String = Item.Trim().Trim(""""c)
                    If Choice.Length > 0 Then Choices.Add(Choice)
                Next
            End If

            Return Choices
        End Function

        Private Shared Function ResolveValidationRange(
            WS As DevExpress.Spreadsheet.Worksheet,
            Formula As String) As DevExpress.Spreadsheet.CellRange

            If String.IsNullOrWhiteSpace(Formula) Then Return Nothing

            Dim Reference As String = Formula.Trim()
            If Reference.StartsWith("=", StringComparison.Ordinal) Then Reference = Reference.Substring(1)

            Try
                Dim SheetName As DevExpress.Spreadsheet.DefinedName = WS.DefinedNames.GetDefinedName(Reference)
                If SheetName IsNot Nothing Then Return SheetName.Range
            Catch
                'The expression is not a worksheet-scoped range name.
            End Try

            Try
                Dim WorkbookName As DevExpress.Spreadsheet.DefinedName = WS.Workbook.DefinedNames.GetDefinedName(Reference)
                If WorkbookName IsNot Nothing Then Return WorkbookName.Range
            Catch
                'The expression is not a workbook-scoped range name.
            End Try

            Try
                Return WS.Workbook.Range(Reference)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub AddRangeChoices(
            SourceRange As DevExpress.Spreadsheet.CellRange,
            Choices As List(Of String))

            If SourceRange Is Nothing Then Return

            For RowIndex As Integer = 0 To SourceRange.RowCount - 1
                For ColumnIndex As Integer = 0 To SourceRange.ColumnCount - 1
                    Dim Choice As String = SourceRange(RowIndex, ColumnIndex).DisplayText.Trim()
                    If Choice.Length > 0 Then Choices.Add(Choice)
                Next
            Next
        End Sub

        Sub AddValidatedRange(CellRange As DevExpress.Spreadsheet.CellRange, ChoiceList As List(Of String))

            For Each ExistingRange As ValidatedRange In ValidatedRanges
                If CellRange.IsIntersecting(ExistingRange.CellRange) Then
                    HasOverlappingRanges = True
                    Exit For
                End If
            Next

            ValidatedRanges.Add(New ValidatedRange With {
                .CellRange = CellRange,
                .ChoiceList = ChoiceList})

        End Sub

        Structure ValidatedRange

            Public CellRange As DevExpress.Spreadsheet.CellRange
            Public ChoiceList As List(Of String)

        End Structure

        Public Function CheckValidation(CellToExamine As DevExpress.Spreadsheet.CellRange) As List(Of String)

            If CellToExamine Is Nothing Then Return Nothing

            If Not HasOverlappingRanges AndAlso
               LastMatchedRangeIndex >= 0 AndAlso
               LastMatchedRangeIndex < ValidatedRanges.Count Then

                Dim LastMatch As ValidatedRange = ValidatedRanges(LastMatchedRangeIndex)
                If CellToExamine.IsIntersecting(LastMatch.CellRange) Then
                    Return LastMatch.ChoiceList
                End If
            End If

            For RangeIndex As Integer = 0 To ValidatedRanges.Count - 1

                If RangeIndex = LastMatchedRangeIndex Then Continue For

                Dim ValRange As ValidatedRange = ValidatedRanges(RangeIndex)

                If CellToExamine.IsIntersecting(ValRange.CellRange) Then

                    LastMatchedRangeIndex = RangeIndex
                    Return ValRange.ChoiceList

                End If

            Next

            Return Nothing

        End Function


    End Class

End Namespace
