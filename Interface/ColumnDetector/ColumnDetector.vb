Imports DevExpress.Spreadsheet
Imports DevExpress.Utils.Extensions

Namespace Abovo
    Friend Class ColumnDetector

        Implements IDataSourceColumnTypeDetector

        Private ColList As List(Of String)
        Private TyList As List(Of String)
        Public Sub New(ColumnList As List(Of String), TypeList As List(Of String))

            ColList = ColumnList
            TyList = TypeList

        End Sub
        Public Function GetColumnName(ByVal index As Integer, ByVal offset As Integer, ByVal range As CellRange) As String Implements IDataSourceColumnTypeDetector.GetColumnName

            If range(-3, offset).DisplayText = "" Then

                Return "(Empty Column " & offset.ToString & ")"

            Else

                Return ColList.ElementAt(offset)

            End If

        End Function

        Public Function GetColumnType(ByVal index As Integer, ByVal offset As Integer, ByVal range As CellRange) As Type Implements IDataSourceColumnTypeDetector.GetColumnType

            Dim defaultType As Type = GetType(String)

            Select Case TyList.ElementAt(offset)

                Case "S"

                    Return GetType(String)

                Case "I"

                    Return GetType(Integer)

                Case "D"

                    Return GetType(Double)

                Case "B"

                    Return GetType(Boolean)

                Case "P"

                    Return GetType(Decimal)

                Case Else

                    Return GetType(String)

            End Select

            Return defaultType

        End Function

    End Class

End Namespace
