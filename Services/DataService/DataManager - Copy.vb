Imports Abovo.LogDebugDev
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.DataObject

Imports System.ComponentModel
Imports System.Text
Imports System
Imports Microsoft.Office.Interop
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraSpreadsheet.Model
Imports System.Collections.ObjectModel
Imports DevExpress.CodeParser
Imports DevExpress.Utils.Extensions

Imports DevExpress.XtraTreeList.Data
Imports DevExpress.Charts.Native
Imports DevExpress.DataAccess.Native.Data
Imports DevExpress.XtraCharts.Designer.Native
Imports DevExpress.XtraPrinting.Native
Imports DevExpress.XtraSpreadsheet.Import.Xls

Imports System.Runtime.Remoting.Messaging
Imports DevExpress.DataAccess.DataFederation
Imports DevExpress.XtraRichEdit.Model
Imports DevExpress.Utils.Drawing


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

            SystemLog(vbLf + "XXXXXXXXXXXXXXX Full Object Status Request XXXXXXXXXXXXXXXX" + vbLf + vbLf)
            Dim DS As DataCellRange

            Dim StrOutput As String = "'''''''''''''''''''''''''''''''''" + vbLf
            Dim i As Integer



            For Each DS In DataSets

                StrOutput += "DS(" & DS.Index & ") Name: " & DS.Name + vbLf
                StrOutput += "Dirty: " & DS.IsDirty.ToString + vbLf
                StrOutput += "RowsByParam: " & DS.RowCount.ToString
                StrOutput += " ColsByParam: " & DS.ColCount.ToString + vbLf
                StrOutput += "RowsByArray: " & UBound(DS.DataRows).ToString
                StrOutput += " ColsByArray: " & UBound(DS.DataColumns).ToString + vbLf
                StrOutput += " DataFields: "

                For i = 0 To DS.ColCount - 1

                    StrOutput += DS.DataColumns(i).ColumnTag.ColumnHeading
                    StrOutput += " (" & DS.DataColumns(i).ColumnTag.DataType & " DP:" & DS.DataColumns(i).DataAddress & " SP:" & DS.DataColumns(i).SourceAddress & "), "

                Next

                StrOutput += vbLf
                For Each DSRow As SheetDataRow In DS.DataRows

                    StrOutput += "Row " & DSRow.Index.ToString + vbLf

                    For Each DSRowCell As CellDataPoint In DSRow.DataCells

                        StrOutput += "C" & DSRowCell.Index.ToString + " (" & DSRowCell.DataType & "), "

                        Select Case DSRowCell.DataType

                            Case "S"

                        End Select

                    Next

                    StrOutput += vbLf
                    For Each DSRowCell As CellDataPoint In DSRow.DataCells

                        Select Case DSRowCell.DataType

                            Case "S"

                                StrOutput += DSRowCell.StringValue

                            Case "I", "Y"

                                StrOutput += DSRowCell.IntValue.ToString

                            Case "B"

                                StrOutput += DSRowCell.BoolValue.ToString

                            Case "N", "P", "M"

                                StrOutput += DSRowCell.RealValue.ToString

                        End Select

                        StrOutput += " (" & DSRowCell.SourceAddress & "), "

                    Next

                    StrOutput += vbLf

                Next

                StrOutput += vbLf + "'''" + vbLf

            Next

            Return StrOutput

        End Function
        Public Function GetISEDataStructure(ModelID, SetGSID, SetCSID, SetIntSecID, SetISDID) As DataCellRange

            Dim Trans As New AbovoTransaction
            Dim g = ExcelModels(ModelID)
            SystemLog("'''''''''''''''''''''''''''''''''''")
            SystemLog("Data structure call.  ModelID: " & ModelID.ToString & " CS ID:" & SetCSID & " Interface Section ID:" & SetIntSecID & " Datasource ID:" & SetISDID)
            SystemLog("'''''''''''''''''''''''''''''''''''")
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
                    .RO = True,
                    .FormatMap = CRSource.DataFieldDefinitions(0).DataFormat
                }
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

                            For x = 0 To DefiningRange.ColumnCount - 1

                                ColIndex += 1
                                CellExamineNRD = DefiningRange(0, x)
                                ColHead = Replace(DataFieldDefinition.FieldName, "vblf", vbLf) & " vblf " & CellExamineNRD.DisplayText
                                If Len(DataFieldDefinition.Units) > 0 Then ColHead += DataFieldDefinition.Units
                                ReDim Preserve DataSets(DataSetIndex).DataColumns(ColIndex)

                                DataSets(DataSetIndex).DataColumns(ColIndex) = New SheetDataColumn With {
                                        .ColumnTag = New DataColumnTag With {
                                        .ColumnHeading = ColHead,
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

                            Next

                        Else

                            ColIndex += 1

                            ReDim Preserve DataSets(DataSetIndex).DataColumns(ColIndex)

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

                        ReDim Preserve DataSets(DataSetIndex).DataColumns(0)

                        DataSets(DataSetIndex).DataColumns(0) = New SheetDataColumn With {
                                .ColumnTag = New DataColumnTag With {.ColumnHeading = CRSource.ColsDescription, .DataType = "S"},
                                .Index = 0
                                }

                        Dim x As Integer

                        For x = 1 To ColHeaderRange.RowCount

                            CellExamine = ColHeaderRange(x - 1, 0)
                            ReDim Preserve DataSets(DataSetIndex).DataColumns(x)
                            ColHead = CRSource.RowsDescription & " " & CellExamine.DisplayText

                            If ColHeaderRange.ColumnCount = 2 Then

                                CellExamineRight = ColHeaderRange(x - 1, 1)
                                ColHead = ColHead & vbLf & " (" & CellExamineRight.DisplayText & ")"

                            ElseIf ColHeaderRange.ColumnCount = 3 Then

                                CellExamineRight = ColHeaderRange(x - 1, 1)
                                ColHead = ColHead & vbLf & " (" & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(x, 2)
                                ColHead = ColHead & CellExamineRight.DisplayText & ")"

                            ElseIf ColHeaderRange.ColumnCount = 4 Then

                                CellExamineRight = ColHeaderRange(x - 1, 1)
                                ColHead = ColHead & vbLf & " (" & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(x, 2)
                                ColHead = ColHead & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(x, 3)
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

                        Dim RowHead As String

                        DataSets(DataSetIndex).ColCount = ColHeaderRange.RowCount + 1

                        For x = 0 To RowHeaderRange.ColumnCount

                            CellExamine = RowHeaderRange(0, x)
                            ReDim Preserve DataSets(DataSetIndex).DataRows(x)
                            RowHead = CellExamine.DisplayText

                            If ColHeaderRange.RowCount = 2 Then

                                CellExamineRight = ColHeaderRange(1, x)
                                RowHead = RowHead & vbLf & " (" & CellExamineRight.DisplayText & ")"

                            ElseIf ColHeaderRange.ColumnCount = 3 Then

                                CellExamineRight = ColHeaderRange(1, x)
                                RowHead = RowHead & vbLf & " (" & CellExamineRight.DisplayText & " "
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

                        'Dim CRTargetRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.Worksheets(WSName).Range()
                        Dim clCell As DevExpress.Spreadsheet.Cell
                        Dim i As Integer, j As Integer

                        For i = 0 To RowHeaderRange.ColumnCount - 1

                            For j = 1 To ColHeaderRange.RowCount

                                clCell = DataRange(j - 1, i)

                                ReDim Preserve DataSets(DataSetIndex).DataRows(i).DataCells(j)

                                DataSets(DataSetIndex).DataRows(i).DataCells(j) = New CellDataPoint
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).SourceAddress = clCell.GetReferenceA1
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).Index = j
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).DataType = DataSets(DataSetIndex).DataColumns(j).ColumnTag.DataType
                                Select Case DataSets(DataSetIndex).DataColumns(j).ColumnTag.DataType

                                    Case "S"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).StringValue = clCell.DisplayText

                                    Case "B"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).BoolValue = CInt(clCell.Value.NumericValue)

                                    Case "D", "P", "C", "M"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).RealValue = CSng(clCell.Value.NumericValue)

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

                        ReDim Preserve DataSets(DataSetIndex).DataColumns(0)

                        DataSets(DataSetIndex).DataColumns(0) = New SheetDataColumn With {
                                .ColumnTag = New DataColumnTag With {.ColumnHeading = CRSource.RowsDescription, .DataType = "S"},
                                .Index = 0
                                }

                        Dim x As Integer



                        For x = 1 To ColHeaderRange.ColumnCount

                            CellExamine = ColHeaderRange(0, x - 1)
                            ReDim Preserve DataSets(DataSetIndex).DataColumns(x)
                            ColHead = CellExamine.DisplayText

                            If ColHeaderRange.ColumnCount = 2 Then

                                CellExamineRight = ColHeaderRange(1, x - 1)
                                ColHead = ColHead & vbLf & " (" & CellExamineRight.DisplayText & ")"

                            ElseIf ColHeaderRange.ColumnCount = 3 Then

                                CellExamineRight = ColHeaderRange(1, x - 1)
                                ColHead = ColHead & vbLf & " (" & CellExamineRight.DisplayText & " "
                                CellExamineRight = ColHeaderRange(2, x - 1)
                                ColHead = ColHead & CellExamineRight.DisplayText & ")"

                            ElseIf ColHeaderRange.ColumnCount = 4 Then

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

                                DataSets(DataSetIndex).DataColumns(x).ColumnTag.MaxVal = CRSource.DataFieldDefinitions(0).MinVal

                            End If

                        Next

                        Dim RowHead As String

                        DataSets(DataSetIndex).RowCount = RowHeaderRange.RowCount

                        For x = 0 To RowHeaderRange.RowCount

                            CellExamine = RowHeaderRange(x, 0)
                            ReDim Preserve DataSets(DataSetIndex).DataRows(x)
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

                            ReDim DataSets(DataSetIndex).DataRows(x).DataCells(0)
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

                                ReDim Preserve DataSets(DataSetIndex).DataRows(i).DataCells(j)

                                DataSets(DataSetIndex).DataRows(i).DataCells(j) = New CellDataPoint
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).SourceAddress = clCell.GetReferenceA1
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).Index = j
                                DataSets(DataSetIndex).DataRows(i).DataCells(j).DataType = DataSets(DataSetIndex).DataColumns(j).ColumnTag.DataType

                                Select Case DataSets(DataSetIndex).DataColumns(j).ColumnTag.DataType

                                    Case "S"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).StringValue = clCell.DisplayText

                                    Case "B"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).BoolValue = CInt(clCell.Value.NumericValue)

                                    Case "D", "P", "C", "M"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).RealValue = CSng(clCell.Value.NumericValue)

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

                Dim SourceWS As String = ""
                Dim WSName As String = IEDSource.CellRangeSources(0).WSName
                Dim CRSource As CellRangeDataSource = IEDSource.CellRangeSources(0)
                Dim CurrWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets(WSName)
                Dim ColIndex As Integer = -1
                Dim CurrDataType As String
                Dim RowOffset As Integer = 0
                Dim ColOffset As Integer = 0F
                Dim IsCalc As Boolean = False
                Dim i As Integer, j As Integer, k As Integer
                Dim MultiRepeatingHeaders As Boolean = False
                Dim ColHead As String
                Dim ApplyRule As Boolean = False
                Dim LastCR As String = ""
                DataSetIndex += 1
                ReDim Preserve DataSets(DataSetIndex)

                DataSets(DataSetIndex) = New DataCellRange(DataSetIndex, ModelID) With {
                    .Name = "MergeDS" & WSName & "-IEDS-" & IEDSource.ISDName,
                    .IsDirty = False,
                    .HasBands = False,
                    .SourceWorksheet = WSName,
                    .RO = IIf(IEDSource.RO = "TRUE", True, False)
                }

                Trans.IntegerReturn = DataSetIndex
                Dim DataRange As DevExpress.Spreadsheet.CellRange

                If CRSource.NRDSName = "CR" Then

                    DataRange = CurrWS.Range(CRSource.DataRange)
                    SystemLog("Reading from " & CRSource.DataRange)

                Else

                    DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.NRDSName).Range

                End If


                ReDim Preserve DataSets(DataSetIndex).DataRows(DataRange.ColumnCount - 1)

                DataSets(DataSetIndex).RowCount = DataRange.ColumnCount

                Dim conditionalFormattings As DevExpress.Spreadsheet.ConditionalFormattingCollection = CurrWS.ConditionalFormattings

                SystemLog("Adding rows: " & DataRange.ColumnCount.ToString)

                For i = 0 To DataRange.ColumnCount - 1

                    DataSets(DataSetIndex).DataRows(i) = New SheetDataRow With {.Index = i}

                Next

                For Each CellRangeSource In IEDSource.CellRangeSources

                    RowOffset = 0
                    ColOffset = 0

                    If CellRangeSource.IsCalculated = "TRUE" Then

                        DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CellRangeSource.OffSetNR).Range



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

                        If CRSource.NRDSName = "CR" Then

                            DataRange = CurrWS.Range(CRSource.DataRange)
                            SourceWS = CurrWS.Name


                        Else

                            DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.NRDSName).Range

                        End If



                    End If

                    i = 0
                    j = 0

                    For Each DataFieldDefinition In CellRangeSource.DataFieldDefinitions

                        MultiRepeatingHeaders = False

                        CurrDataType = DataFieldDefinition.DataFormat

                        If CurrDataType = "DUMMY" Then GoTo NextDFD

                        Dim RepeatCount As Integer = 0

                        If DataFieldDefinition.RepeatsByNR = "TRUE" Or DataFieldDefinition.RepeatsByCR = "TRUE" Then

                            Dim RepeatMethod As String = "PORT"

                            DataSets(DataSetIndex).HasBands = True

                            Dim DefiningRange As DevExpress.Spreadsheet.CellRange

                            If DataFieldDefinition.RepeatsByNR = "TRUE" Then

                                DefiningRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(DataFieldDefinition.RepeatingNR).Range

                            Else

                                DefiningRange = ExcelModels(ModelID).WB.Worksheets(CellRangeSource.WSName).Range(DataFieldDefinition.RepeatsByCRData)

                            End If

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

                            For AddCols = 0 To RepeatCount - 1

                                j += 1

                                If RepeatMethod = "PORT" Then

                                    CellExamineNRD = DefiningRange(AddCols, 0)

                                Else

                                    CellExamineNRD = DefiningRange(0, AddCols)

                                End If

                                ColHead = DataFieldDefinition.FieldName & " " & CellExamineNRD.DisplayText

                                'If Not IsNothing(DataFieldDefinition.RepeatingHeaderText) Then ColHead = DataFieldDefinition.RepeatingHeaderText & " " & ColHead

                                If MultiRepeatingHeaders = True Then

                                    CellExamineRight = DefiningRange(AddCols, 1)

                                    If Not IsNothing(DataFieldDefinition.ExtraHeadingPreWord) Then
                                        ColHead += vbLf & " (" & DataFieldDefinition.ExtraHeadingPreWord & " " & CellExamineRight.DisplayText & ")"
                                    Else
                                        ColHead += vbLf & " (" & CellExamineRight.DisplayText & ")"
                                    End If

                                End If

                                If Not DataFieldDefinition.Units Is Nothing Then ColHead += vbLf & DataFieldDefinition.Units

                                If CellExamineNRD.DisplayText = "" Then GoTo Nextnrds

                                ColIndex += 1

                                ReDim Preserve DataSets(DataSetIndex).DataColumns(ColIndex)

                                If DataFieldDefinition.HasRule = "TRUE" Then
                                    DataSets(DataSetIndex).HasRules = True
                                    ApplyRule = True
                                Else
                                    ApplyRule = False
                                End If

                                DataSets(DataSetIndex).DataColumns(ColIndex) = New SheetDataColumn With {
                                        .ColumnTag = New DataColumnTag With {
                                        .ColumnHeading = Replace(ColHead, "vblf", vbLf),
                                        .DataType = CurrDataType,
                                        .IsReadOnly = IIf(DataFieldDefinition.RO = "TRUE", True, False),
                                        .IsCalculated = IIf(CellRangeSource.IsCalculated = "TRUE", True, False),
                                        .HasActions = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", True, False),
                                        .HasRules = ApplyRule,
                                        .ActionNR = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", DataFieldDefinition.RepeatingNR, Nothing),
                                        .ActionData = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", "NRRI", Nothing),
                                        .Units = DataFieldDefinition.Units,
                                        .IsFixed = IIf(DataFieldDefinition.Fixed = "TRUE", True, False),
                                        .ShowSummary = DataFieldDefinition.ShowSummary,
                                        .MinVal = DataFieldDefinition.MinVal,
                                        .MaxVal = DataFieldDefinition.MaxVal,
                                        .BandID = DataFieldDefinition.RepeatingHeaderText,
                                        .TipText = DataFieldDefinition.TipText,
                                        .RepositaryID = DataFieldDefinition.RepositaryItemID
                                        },
                                        .Index = ColIndex
                                        }

                                If CellRangeSource.BandID IsNot Nothing Then

                                    DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.BandID = CellRangeSource.BandID
                                    DataSets(DataSetIndex).HasBands = True

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

                                '''''''
                                For i = 0 To k - 1

                                    CellExamine = DataRange(j + RowOffset, i + ColOffset)

                                    ReDim Preserve DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex)

                                    DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex) = New CellDataPoint With {
                                        .Index = ColIndex,
                                        .SourceAddress = CellExamine.GetReferenceA1,
                                        .SourceSheet = SourceWS}

                                    If ApplyRule Then DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).IsLocked = IIf(CellExamine.Fill.PatternType = PatternType.Solid, False, True)

                                    If Not CellExamine.Value.IsEmpty Then

                                        If Not DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsCalculated Then

                                            DataSets(DataSetIndex).DataRows(i).IsEmpty = False

                                        End If

                                        Select Case CurrDataType

                                            Case "S"

                                                DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).StringValue = CellExamine.DisplayText

                                            Case "B"

                                                DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).BoolValue = CInt(CellExamine.Value.NumericValue)

                                            Case "D", "P", "C", "M"

                                                DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).RealValue = CSng(CellExamine.Value.NumericValue)

                                            Case "I", "Y"

                                                DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).IntValue = CInt(CellExamine.Value.NumericValue)

                                        End Select

                                        DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).IsEmpty = False

                                    Else

                                        DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).IsEmpty = True
                                        DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).StringValue = Nothing
                                        DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).BoolValue = Nothing
                                        DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).RealValue = Nothing
                                        DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).IntValue = Nothing

                                    End If

                                Next
Nextnrds:
                            Next

                        Else

                            ColIndex += 1

                            ReDim Preserve DataSets(DataSetIndex).DataColumns(ColIndex)

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
                                    .IsReadOnly = IIf(DataFieldDefinition.RO = "TRUE", True, False),
                                    .IsCalculated = IIf(CellRangeSource.IsCalculated = "TRUE", True, False),
                                    .HasRules = ApplyRule,
                                    .ShowSummary = DataFieldDefinition.ShowSummary,
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

                            If CellRangeSource.BandID IsNot Nothing Then

                                DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.BandID = CellRangeSource.BandID
                                DataSets(DataSetIndex).HasBands = True

                            End If

                            If CRSource.NRDSName = "CR" Then

                                If LastCR <> CRSource.DataRange Then
                                    DataRange = CurrWS.Range(CRSource.DataRange)
                                    LastCR = CRSource.DataRange
                                    k = 0
                                End If

                                SystemLog("Reading from " & CRSource.DataRange)

                            End If



                            For i = 0 To k - 1

                                CellExamine = DataRange(j + RowOffset, i + ColOffset)

                                ReDim Preserve DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex)

                                DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex) = New CellDataPoint With {
                                    .Index = ColIndex,
                                    .SourceAddress = CellExamine.GetReferenceA1,
                                    .SourceSheet = SourceWS}
                                SystemLog("Adding cell " & CellExamine.GetReferenceA1 & " Source sheet: " & SourceWS)

                                If ApplyRule Then

                                    DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).IsLocked = IIf(CellExamine.Fill.PatternType = PatternType.Solid, False, True)

                                End If

                                If Len(CellExamine.DisplayText) > 0 Then

                                    If Not DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsCalculated Then

                                        DataSets(DataSetIndex).DataRows(i).IsEmpty = False

                                    End If

                                    SystemLog("Adding from: " & CellExamine.GetReferenceA1 & ". Display text " & CellExamine.DisplayText & " (Len:" & Len(CellExamine.DisplayText) & "). Numeric value " & CellExamine.Value.NumericValue.ToString)

                                    Select Case CurrDataType

                                        Case "S"

                                            DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).StringValue = CellExamine.DisplayText

                                        Case "B"

                                            DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).BoolValue = CInt(CellExamine.Value.NumericValue)

                                        Case "D", "P", "C", "M"

                                            DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).RealValue = CSng(CellExamine.Value.NumericValue)

                                        Case "I", "Y"

                                            DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).IntValue = CInt(CellExamine.Value.NumericValue)

                                    End Select

                                    DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).IsEmpty = False

                                Else

                                    DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).IsEmpty = True

                                    DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).StringValue = Nothing
                                    DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).BoolValue = Nothing
                                    DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).RealValue = Nothing
                                    DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).IntValue = Nothing

                                End If
                            Next

                        End If

                        j += 1

NextDFD:

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

                Dim DataRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.NRDSName).Range
                Dim CellDF As DataFieldDefinition = CRSource.DataFieldDefinitions(0)

                ReDim Preserve DataSets(DataSetIndex).DataColumns(0)

                CurrDataType = CellDF.DataFormat

                DataSets(DataSetIndex).DataColumns(0) = New SheetDataColumn With {
                                    .ColumnTag = New DataColumnTag With {
                                    .ColumnHeading = Replace(CellDF.FieldName, "vblf", vbLf),
                                    .DataType = CurrDataType,
                                    .IsReadOnly = IIf(CellDF.RO = "True", True, False),
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

                CellExamine = DataRange(0, 0)

                DataSets(DataSetIndex).DataRows(0).DataCells(0) = New CellDataPoint With {
                                    .Index = 0,
                                    .SourceAddress = CellExamine.GetReferenceA1,
                                    .SourceSheet = WSName
}
                If Not CellExamine.Value.IsEmpty Then


                    Select Case CurrDataType

                        Case "S"

                            DataSets(DataSetIndex).DataRows(0).DataCells(0).StringValue = CellExamine.DisplayText

                        Case "B"

                            DataSets(DataSetIndex).DataRows(0).DataCells(0).BoolValue = CInt(CellExamine.Value.NumericValue)

                        Case "P", "C", "M"

                            DataSets(DataSetIndex).DataRows(0).DataCells(0).RealValue = CSng(CellExamine.Value.NumericValue)

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

                Dim SourceWS As String
                Dim WSName As String = IEDSource.CellRangeSources(0).WSName
                Dim CRSource As CellRangeDataSource = IEDSource.CellRangeSources(0)
                Dim CurrWS As DevExpress.Spreadsheet.Worksheet = ExcelModels(ModelID).WB.Worksheets(WSName)
                Dim ColIndex As Integer = -1
                Dim CurrDataType As String
                Dim RowOffset As Integer = 0
                Dim ColOffset As Integer = 0

                Dim IsCalc As Boolean = False
                Dim i As Integer, j As Integer, k As Integer
                DataSetIndex += 1
                ReDim Preserve DataSets(DataSetIndex)
                Dim ApplyRule As Boolean = False

                DataSets(DataSetIndex) = New DataCellRange(DataSetIndex, ModelID) With {
                    .Name = "MergeAcrossDS" & WSName & "-IEDS: " & IEDSource.ISDName,
                    .IsDirty = False,
                    .SourceWorksheet = WSName,
                    .HasBands = False,
                    .RO = IIf(IEDSource.RO = "TRUE", True, False)
                }

                Trans.IntegerReturn = DataSetIndex

                Dim DataRange As DevExpress.Spreadsheet.CellRange

                If CRSource.IsCalculated = "TRUE" Then

                    DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.OffSetNR).Range

                Else

                    DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CRSource.NRDSName).Range

                End If




                ReDim Preserve DataSets(DataSetIndex).DataRows(DataRange.RowCount - 1)

                DataSets(DataSetIndex).Capacity = DataRange.RowCount
                DataSets(DataSetIndex).RowCount = DataRange.RowCount

                For i = 0 To DataRange.RowCount - 1

                    DataSets(DataSetIndex).DataRows(i) = New SheetDataRow With {.Index = i}

                Next

                k = DataSets(DataSetIndex).RowCount

                For Each CellRangeSource In IEDSource.CellRangeSources

                    RowOffset = 0
                    ColOffset = 0

                    If CellRangeSource.IsCalculated = "TRUE" Then

                        DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CellRangeSource.OffSetNR).Range
                        SourceWS = DataRange.Worksheet.Name
                        IsCalc = True

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

                        DataRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(CellRangeSource.NRDSName).Range
                        SourceWS = DataRange.Worksheet.Name

                    End If



                    SourceWS = DataRange.Worksheet.Name

                    i = 0
                    j = 0

                    For Each DataFieldDefinition In CellRangeSource.DataFieldDefinitions

                        CurrDataType = DataFieldDefinition.DataFormat

                        If CurrDataType = "DUMMY" Then GoTo NextDFD2

                        Dim RepeatCount As Integer = 0

                        If DataFieldDefinition.RepeatsByNR = "TRUE" Then

                            Dim RepeatMethod As String = "PORT"

                            DataSets(DataSetIndex).HasBands = True

                            Dim DefiningRange As DevExpress.Spreadsheet.CellRange = ExcelModels(ModelID).WB.DefinedNames.GetDefinedName(DataFieldDefinition.RepeatingNR).Range

                            If DefiningRange.RowCount > DefiningRange.ColumnCount Then

                                RepeatMethod = "PORT"
                                RepeatCount = DefiningRange.RowCount

                            Else

                                RepeatMethod = "LAND"
                                RepeatCount = DefiningRange.ColumnCount

                            End If

                            Dim CellExamineNRD As DevExpress.Spreadsheet.Cell

                            For AddCols = 0 To RepeatCount - 1

                                If RepeatMethod = "PORT" Then

                                    CellExamineNRD = DefiningRange(AddCols, 0)

                                Else

                                    CellExamineNRD = DefiningRange(0, AddCols)

                                End If

                                If CellExamineNRD.DisplayText = "" Then GoTo Nextnrds2

                                ColIndex += 1

                                If DataFieldDefinition.HasRule = "TRUE" Then

                                    DataSets(DataSetIndex).HasRules = True
                                    ApplyRule = True

                                Else

                                    ApplyRule = False

                                End If

                                ReDim Preserve DataSets(DataSetIndex).DataColumns(ColIndex)
                                Dim CurrColHeading As String = Replace(If(Microsoft.VisualBasic.Right(DataFieldDefinition.FieldName, 4) = "NONE", "", DataFieldDefinition.FieldName & " ") & CellExamineNRD.DisplayText, "vblf", vbLf)
                                DataSets(DataSetIndex).DataColumns(ColIndex) = New SheetDataColumn With {
                                    .ColumnTag = New DataColumnTag With {
                                    .ColumnHeading = CurrColHeading,
                                    .DataType = CurrDataType,
                                    .Units = DataFieldDefinition.Units,
                                    .IsReadOnly = IIf(DataFieldDefinition.RO = "TRUE", True, False),
                                    .IsCalculated = IIf(CellRangeSource.IsCalculated = "TRUE", True, False),
                                    .IsFixed = IIf(DataFieldDefinition.Fixed = "TRUE", True, False),
                                    .HasActions = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", True, False),
                                    .HasRules = IIf(DataFieldDefinition.HasRule = "TRUE", True, False),
                                    .ActionNR = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", DataFieldDefinition.RepeatingNR, Nothing),
                                    .ActionData = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", "NRCI", Nothing),
                                    .ShowSummary = DataFieldDefinition.ShowSummary,
                                    .MinVal = DataFieldDefinition.MinVal,
                                    .MaxVal = DataFieldDefinition.MaxVal,
                                    .BandID = DataFieldDefinition.RepeatingHeaderText,
                                    .RepositaryID = DataFieldDefinition.RepositaryItemID
                                    },
                                    .Index = ColIndex
                                    }

                                If CellRangeSource.RO = "TRUE" Then DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsReadOnly = True

                                For j = 0 To k - 1

                                    CellExamine = DataRange(j + RowOffset, AddCols + ColOffset)

                                    ReDim Preserve DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex)

                                    DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex) = New CellDataPoint With {
                                        .Index = ColIndex,
                                        .SourceAddress = CellExamine.GetReferenceA1,
                                        .SourceSheet = SourceWS}

                                    If ApplyRule Then DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).IsLocked = IIf(CellExamine.Fill.PatternType = PatternType.Solid, False, True)

                                    If Not CellExamine.Value.IsEmpty Then

                                        If Not DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsCalculated Then 'And Not CellExamine.DisplayText = "0"

                                            DataSets(DataSetIndex).DataRows(j).IsEmpty = False

                                        End If

                                        Select Case CurrDataType

                                            Case "S"

                                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).StringValue = CellExamine.DisplayText

                                            Case "FL"

                                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).StringValue = Left(CellExamine.DisplayText, 10)
                                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).ExtraData = CellExamine.DisplayText

                                            Case "B"

                                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).BoolValue = CInt(CellExamine.Value.NumericValue)

                                            Case "D", "P", "C", "M"

                                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).RealValue = CSng(CellExamine.Value.NumericValue)

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
Nextnrds2:
                            Next

                        Else

                            ColIndex += 1

                            ReDim Preserve DataSets(DataSetIndex).DataColumns(ColIndex)

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
                                    .HasActions = IIf(DataFieldDefinition.EditRepNRHere = "TRUE", True, False),
                                    .IsCalculated = IIf(CellRangeSource.IsCalculated = "TRUE", True, False),
                                    .MinVal = DataFieldDefinition.MinVal,
                                    .MaxVal = DataFieldDefinition.MaxVal,
                                    .BandID = DataFieldDefinition.RepeatingHeaderText,
                                    .RepositaryID = DataFieldDefinition.RepositaryItemID
                                    },
                                    .Index = ColIndex
                                    }

                            If CellRangeSource.RO = "TRUE" Then

                                DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsReadOnly = True

                            End If

                            For j = 0 To k - 1


                                CellExamine = DataRange(j + RowOffset, i + ColOffset)

                                ReDim Preserve DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex)

                                DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex) = New CellDataPoint With {
                                    .Index = ColIndex,
                                    .SourceAddress = CellExamine.GetReferenceA1,
                                    .SourceSheet = SourceWS}

                                If ApplyRule Then DataSets(DataSetIndex).DataRows(i).DataCells(ColIndex).IsLocked = IIf(CellExamine.Fill.PatternType = PatternType.Solid, False, True)

                                If Not CellExamine.Value.IsEmpty Then

                                    If Not DataSets(DataSetIndex).DataColumns(ColIndex).ColumnTag.IsCalculated Then

                                        DataSets(DataSetIndex).DataRows(j).IsEmpty = False

                                    End If

                                    Select Case CurrDataType

                                        Case "S"

                                            DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).StringValue = CellExamine.DisplayText

                                        Case "B"

                                            DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).BoolValue = CInt(CellExamine.Value.NumericValue)

                                        Case "D", "P", "C", "M"

                                            DataSets(DataSetIndex).DataRows(j).DataCells(ColIndex).RealValue = CSng(CellExamine.Value.NumericValue)

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
        Public Function RenderIEHTMLCource(ModelID, SetGSID, SetCSID, SetIntSecID, SetISDID) As String

            Dim StrOutput As String = My.Resources.StringTemplates.HTMLFinanceTableHeader
            StrOutput += My.Resources.StringTemplates.HTMLFinanceTablePrecursor

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
                    Dim NumExamine As Single
                    Dim HeightS As String = "80"


                    For i = 0 To DataRange.RowCount - 1

                        StrOutput += "<tr class=xl822235 height=" & HeightS & " style='height:15.45pt'>"

                        For j = 0 To DataRange.ColumnCount - 1

                            CellExamine = DataRange(i, j)

                            StrOutput += "<td height = " & HeightS & " Class=xl882235 style='font-size:11.0pt;"

                            If CellExamine.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center Then

                                StrOutput += "text-align:general;"

                            ElseIf CellExamine.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left Then

                                StrOutput += "text-align:left;"

                            ElseIf CellExamine.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right Then

                                StrOutput += "text-align:right;"

                            End If

                            TestString = CellExamine.DisplayText

                            If IsNumeric(TestString) Then

                                NumExamine = CSng(TestString)

                                If NumExamine < 0 Then

                                    StrOutput += "color:red;"

                                End If


                            End If

                            If CellExamine.Font.Bold = True Then

                                StrOutput += "font-weight:800;"

                            Else

                                StrOutput += "font-weight:400;"

                            End If

                            StrOutput += "background:" & CellExamine.Fill.BackgroundColor.ToString & ";"

                            StrOutput += "Text-decoration: none;text-underline-style:none;text-line-through:none;
                                              Font-family: Arial, sans - serif;mso-background-source: auto;mso-pattern:red thin - diag - stripe'>"


                            StrOutput += CellExamine.DisplayText
                            StrOutput += "</td>"
                        Next

                        StrOutput += "</tr>"
                        HeightS = "21"

                    Next

                    StrOutput += "

                        </table>

                        </div>


                        <!----------------------------->
                        <!--END OF OUTPUT FROM ABOVO SYSTEM-->
                        <!----------------------------->
                        </body>

                        </html>"

                End If

            End If

            Return StrOutput

        End Function


    End Class

    Public Class InstanceDataSet

        Public Controls() As Control

        Sub New(SetModelID As Integer, CSID As Integer)

        End Sub

    End Class


    Partial Public Class AbovoBP

        'Data Strcutures for AbovoBP

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
                    If Len(clCell.DisplayText) > 0 Then AbovoBP.Stock.StockItems(i).NewLetInitialRate = CSng(clCell.Value.NumericValue)
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

        '        SystemLog("Merging " & DataSets(DS).Name)

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

        '    SystemLog("Merged array: " & DataSets(DataSetIndex).Name)
        '    SystemLog("Rows: " & UBound(DataSets(DataSetIndex).DataRows))
        '    SystemLog("Defined Columns: " & UBound(DataSets(DataSetIndex).DataColumns))
        '    SystemLog("Columns from row: " & UBound(DataSets(DataSetIndex).DataRows(0).DataCells))

        '    For Each DS As Integer In DataSetsToMerge

        '        SystemLog("Populating from: " & DataSets(DS).Name)

        '        'add columns
        '        For Each SourceRow As SheetDataRow In DataSets(DS).DataRows

        '            'SystemLog("Doing source row:" & SourceRow.Index)

        '            For Each SourceCell As CellDataPoint In SourceRow.DataCells

        '                'SystemLog("Doing source cell:" & SourceCell.Index)

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

            'SystemLog(vbLf + "--------------UNION ACROSS-----------------")
            'SystemLog("Merging to " & SetName)

            'For Each DS As Integer In DataSetsToMerge

            '    SystemLog("Sourcing from " & DataSets(DS).Name & " (" & DS.ToString & ")")

            '    'add columns
            '    For Each SourceColumn As SheetDataColumn In DataSets(DS).DataColumns

            '        ColCount += 1
            '        ReDim Preserve DataSets(DataSetIndex).DataColumns(ColCount - 1)

            '        DataSets(DataSetIndex).DataColumns(ColCount - 1) = New SheetDataColumn With {
            '            .ColumnTag = New DataColumnTag With {.ColumnHeading = SourceColumn.ColumnTag.ColumnHeading, .DataType = SourceColumn.ColumnTag.DataType, .IsReadOnly = MakeFirstReadOnly},
            '            .Index = ColCount - 1
            '        }

            '        SystemLog("Adding column " & SourceColumn.ColumnTag.ColumnHeading & " (" & SourceColumn.ColumnTag.DataType & ")")
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

            'SystemLog("Merged array: " & DataSets(DataSetIndex).Name)
            'SystemLog("Rows: " & UBound(DataSets(DataSetIndex).DataRows))
            'SystemLog("Defined Columns: " & UBound(DataSets(DataSetIndex).DataColumns))
            'SystemLog("Columns from row: " & UBound(DataSets(DataSetIndex).DataRows(0).DataCells))

            'For Each DS As Integer In DataSetsToMerge

            '    SystemLog("Populating from: " & DataSets(DS).Name)

            '    'add columns
            '    For Each SourceRow As SheetDataRow In DataSets(DS).DataRows

            '        'SystemLog("Doing source row:" & SourceRow.Index)

            '        For Each SourceCell As CellDataPoint In SourceRow.DataCells

            '            'SystemLog("Doing source cell:" & SourceCell.Index)

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

            ''SystemLog("Reading header range: " & RangeName)
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

            ''SystemLog("Reading header range: " & RangeName)
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

            'SystemLog("Setting rows to: " & Activeposition.ToString)

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

            ''SystemLog("Reading header range: " & RangeName)
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

            '    'SystemLog("Doing i: " & i.ToString)

            '    c = Mid(FormatMap, i, 1)

            '    If c <> "D" Then

            '        Activeposition += 1

            '        For j = 0 To CRTargetRange.ColumnCount - 1

            '            clCell = CRTargetRange(i - 1, j)
            '            'SystemLog("Readig cell: " & clCell.GetReferenceA1)
            '            'SystemLog("Cellnv: " & clCell.Value.NumericValue.ToString)
            '            'SystemLog("Cellnv: " & clCell.DisplayText)
            '            DataSets(DataSetIndex).DataRows(j).DataCells(Activeposition) = New CellDataPoint With {
            '                .Index = Activeposition,
            '                .DataType = c,
            '                .SourceAddress = clCell.GetReferenceA1
            '            }

            '            If Len(clCell.DisplayText) > 0 Then
            '                'SystemLog("About to cast row: " & j.ToString)
            '                'SystemLog("About to cast Col: " & Activeposition.ToString)

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

            SystemLog("Finish")
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
            '            SystemLog("Reading range: " & RangeName)
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

            '            SystemLog("Source Col count: " & CRTargetRange.ColumnCount.ToString)
            '            SystemLog("Source Row count: " & CRTargetRange.RowCount.ToString)
            '            SystemLog("Set Row count: " & SetRowCount.ToString)
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

            '                    SystemLog("Adding Cell data point to " & DataSets(DataSetIndex).Name & " Row: ")
            '                    SystemLog("Row: " & DataSets(DataSetIndex).DataRows(i).Index.ToString)
            '                    SystemLog("Cell: " & Activeposition.ToString)
            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition) = New CellDataPoint

            '                    clCell = CRTargetRange(i, j)

            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).SourceAddress = clCell.GetReferenceA1
            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).Index = Activeposition
            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).DataType = DataSets(DataSetIndex).DataColumns(Activeposition).ColumnTag.DataType

            '                    If Len(clCell.DisplayText) > 0 Then

            '                        DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IsEmpty = False
            '                        SystemLog(DataSets(DataSetIndex).DataColumns(Activeposition).ColumnTag.DataType)

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
            '            SystemLog("Reading worksheet: " & SourceSheet)
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

            '            SystemLog("Source Col count: " & CRTargetRange.ColumnCount.ToString)
            '            SystemLog("Source Row count: " & CRTargetRange.RowCount.ToString)
            '            SystemLog("Set Row count: " & SetRowCount.ToString)
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

            '                    SystemLog("Adding Cell data point to " & DataSets(DataSetIndex).Name & " Row: ")
            '                    SystemLog("Row: " & DataSets(DataSetIndex).DataRows(i).Index.ToString)
            '                    SystemLog("Cell: " & Activeposition.ToString)
            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition) = New CellDataPoint

            '                    clCell = CRTargetRange(i, j)

            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).SourceAddress = clCell.GetReferenceA1
            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).Index = Activeposition
            '                    DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).DataType = DataSets(DataSetIndex).DataColumns(Activeposition).ColumnTag.DataType

            '                    If Len(clCell.DisplayText) > 0 Then

            '                        DataSets(DataSetIndex).DataRows(i).DataCells(Activeposition).IsEmpty = False
            '                        SystemLog(DataSets(DataSetIndex).DataColumns(Activeposition).ColumnTag.DataType)

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

            SystemLog("CALL GetISEDataStructure")
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

                                    Case "N", "P", "M"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).RealValue = CSng(clCell.Value.NumericValue)

                                    Case "I", "Y"

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

                                    Case "N", "P", "M"

                                        DataSets(DataSetIndex).DataRows(i).DataCells(j).RealValue = CSng(clCell.Value.NumericValue)

                                    Case "I", "Y"

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
            Public DataType As String
            Public StringValue As String
            Public BoolValue As Integer
            Public RealValue As Single
            Public IntValue As Integer
            Public IsEmpty As Boolean
            Public SourceAddress As String
            Public SourceSheet As String
            Public IsDirty As Boolean = False
            Public IsLocked As Boolean = False
            Public ExtraData As Object

        End Class

        Structure SheetDataColumn

            Public Index As Integer
            Public DataAddress As Integer
            Public SourceAddress As Integer
            Public ColumnTag As DataColumnTag

        End Structure

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
            Public IsFixed As Boolean
            Public HasActions As Boolean = False
            Public HasRules As Boolean = False
            Public ActionNR As String
            Public ActionData As String
            Public ButtonObjectState As ObjectState
            Public DefaultColumnWidth As Integer
            Public ExtendedColumnWidth As Integer
            Public WidthSet As Boolean = False
            Public BandID As String
            Public Units As String
            Public LabelText As String
            Public DefIncrement As String

        End Class
        Class SingleCellDataTag

            Public ColumnTag As DataColumnTag
            Public CellDP As CellDataPoint

        End Class

        Class SheetDataRow

            Public Index As Integer
            Public IsEmpty As Boolean = True
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
            Public DefaultDataNR As String
            Public DataRange As String
            Public Capacity As Integer = 0
            Public UsedRows As Integer = 0
            Public DataRows() As SheetDataRow
            Public DataColumns() As SheetDataColumn
            Public SourceWorksheet As String
            Public HasBands As Boolean
            Public HasCalcs As Boolean = False
            Public HasRules As Boolean = False

            Sub New(SetIndex As Integer, SetModelID As Integer)

                ReDim DataRows(-1)
                ReDim DataColumns(-1)
                Index = SetIndex
                ModelID = SetModelID

            End Sub

            Sub UpdateCalcs()

                If Not HasCalcs Then Return

                Dim DP As CellDataPoint
                Dim SourceCell As DevExpress.Spreadsheet.Cell

                For Each SheetDataColumn In DataColumns

                    If SheetDataColumn.ColumnTag.IsCalculated Then

                        For Each SheetDataRow In DataRows

                            DP = SheetDataRow.DataCells(SheetDataColumn.Index)
                            SourceCell = GetWorkBook(ModelID).Worksheets(DP.SourceSheet).Cells(DP.SourceAddress)

                            Select Case SheetDataColumn.ColumnTag.DataType

                                Case "S"

                                    SheetDataRow.DataCells(SheetDataColumn.Index).StringValue = SourceCell.DisplayText

                                Case "B"
                                    SheetDataRow.DataCells(SheetDataColumn.Index).BoolValue = CInt(SourceCell.Value.NumericValue)


                                Case "D", "P", "C", "M"
                                    SheetDataRow.DataCells(SheetDataColumn.Index).RealValue = CSng(SourceCell.Value.NumericValue)


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

                For Each SheetDataColumn In DataColumns

                    If SheetDataColumn.ColumnTag.HasRules Then

                        For Each SheetDataRow In DataRows

                            DP = SheetDataRow.DataCells(SheetDataColumn.Index)
                            SourceCell = GetWorkBook(ModelID).Worksheets(DP.SourceSheet).Cells(DP.SourceAddress)
                            SheetDataRow.DataCells(SheetDataColumn.Index).IsLocked = IIf(SourceCell.Fill.PatternType = PatternType.Solid, False, True)

                        Next

                    End If

                Next

            End Sub

            Public Function Clone() As Object Implements ICloneable.Clone
                Return Me.MemberwiseClone
            End Function

        End Class



    End Class
    Public Class AbovoUnboundSource

        Inherits DevExpress.Data.UnboundSource
        Implements ICloneable

        Public UBSTag As AbovoUnboundSourceTag
        Public Sub New(SetTag As AbovoUnboundSourceTag)

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

        End Class

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
End Namespace
