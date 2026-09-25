Option Strict On

Imports System.Collections
Imports System.ComponentModel
Imports DevExpress.Spreadsheet
Imports RangeDataSource = DevExpress.XtraSpreadsheet.Model.RangeDataSource

Namespace Abovo
    ' Read-only analyser adapter. The small presentation document contains only
    ' detached values, never financial formulas. DevExpress retains ownership of
    ' its column-name normalisation and typed list conversions. It is not saved,
    ' calculated, or used as the model's calculation/editing owner.
    Public NotInheritable Class ModelRangeDataSource
        Inherits System.Windows.Forms.BindingSource
        Private nativeSource As RangeDataSource
        Private presentationBook As Workbook
        Private view As ModelEngineView
        Private accepted As WorkbookEngines.WorkbookCalculationResult
        Private disposedSource As Boolean

        Public Shared Function Create(range As CellRange, options As RangeDataSourceOptions) As ModelRangeDataSource
            If range Is Nothing Then Throw New ArgumentNullException(NameOf(range))
            If options Is Nothing Then Throw New ArgumentNullException(NameOf(options))
            Dim source As New ModelRangeDataSource()
            Try
                source.view = ModelEngineView.Find(range(0, 0))
                If source.view Is Nothing Then
                    source.nativeSource = DirectCast(range.GetDataSource(options), RangeDataSource)
                    source.DataSource = source.nativeSource
                    Return source
                End If
                If options.EditingOptions <> DataSourceEditingOptions.ReadOnly OrElse options.PreserveFormulas Then
                    Throw New NotSupportedException("Engine result ranges are read-only values. Use the change manager for inputs.")
                End If
                If Not options.UseFirstRowAsHeader OrElse options.DataSourceColumnTypeDetector Is Nothing OrElse range.RowCount < 2 Then
                    Throw New NotSupportedException("This engine range requires an explicit analyser schema and a header row.")
                End If
                source.accepted = source.view.CaptureRange(range)
                source.presentationBook = New Workbook()
                source.presentationBook.Options.CalculationMode = WorkbookCalculationMode.Manual
                source.presentationBook.DocumentSettings.Calculation.Use1904DateSystem = range.Worksheet.Workbook.DocumentSettings.Calculation.Use1904DateSystem
                Dim sheet = source.presentationBook.Worksheets(0)
                sheet.Name = range.Worksheet.Name
                Dim target = sheet.Range.FromLTRB(range.LeftColumnIndex, range.TopRowIndex, range.RightColumnIndex, range.BottomRowIndex)
                For Each block In source.accepted.Blocks
                    For row = 0 To block.Area.Rows - 1
                        For column = 0 To block.Area.Columns - 1
                            Dim raw = block.ValueAt(row, column)
                            If TypeOf raw Is WorkbookEngines.WorkbookCellError Then
                                raw = DirectCast(raw, WorkbookEngines.WorkbookCellError).Text
                            End If
                            ' Value assignment, not Formula, preserves literal '=...' input.
                            sheet.Cells(block.Area.Row + row, block.Area.Column + column).Value = CellValue.FromObject(raw)
                        Next
                    Next
                Next
                Dim dataRange = range.Worksheet.Range.FromLTRB(range.LeftColumnIndex, range.TopRowIndex + 1, range.RightColumnIndex, range.BottomRowIndex)
                Dim names As New Dictionary(Of Integer, String)()
                Dim types As New Dictionary(Of Integer, Type)()
                Dim visibleIndex = 0
                For column = 0 To range.ColumnCount - 1
                    Dim visible = range.Worksheet.Columns(range.LeftColumnIndex + column).Visible
                    sheet.Columns(range.LeftColumnIndex + column).Visible = visible
                    If options.SkipHiddenColumns AndAlso Not visible Then Continue For
                    names.Add(column, options.DataSourceColumnTypeDetector.GetColumnName(visibleIndex, column, dataRange))
                    Dim columnType = options.DataSourceColumnTypeDetector.GetColumnType(visibleIndex, column, dataRange)
                    If columnType IsNot GetType(String) AndAlso columnType IsNot GetType(Double) AndAlso columnType IsNot GetType(Integer) Then
                        Throw New NotSupportedException("This analyser schema supports text and numeric columns only.")
                    End If
                    types.Add(column, columnType)
                    visibleIndex += 1
                Next
                For row = 0 To range.RowCount - 1
                    sheet.Rows(range.TopRowIndex + row).Visible = range.Worksheet.Rows(range.TopRowIndex + row).Visible
                Next
                Dim detachedOptions As New RangeDataSourceOptions With {
                    .UseFirstRowAsHeader = True, .PreserveFormulas = False,
                    .SkipHiddenRows = options.SkipHiddenRows, .SkipHiddenColumns = options.SkipHiddenColumns,
                    .EditingOptions = DataSourceEditingOptions.ReadOnly,
                    .DataSourceColumnTypeDetector = New DetachedSchema(names, types)}
                source.nativeSource = DirectCast(target.GetDataSource(detachedOptions), RangeDataSource)
                source.DataSource = New GuardedRows(source, source.nativeSource)
                If Not source.HasCurrentValues Then Throw New InvalidOperationException("The analyser was superseded during binding.")
                Return source
            Catch
                source.Dispose()
                Throw
            End Try
        End Function

        Public ReadOnly Property HasCurrentValues As Boolean
            Get
                Return Not disposedSource AndAlso (view Is Nothing OrElse view.IsCurrent(accepted))
            End Get
        End Property

        Protected Overrides Sub Dispose(disposing As Boolean)
            disposedSource = True
            If disposing Then
                DataSource = Nothing
                nativeSource?.Dispose() : nativeSource = Nothing
                presentationBook?.Dispose() : presentationBook = Nothing
            End If
            MyBase.Dispose(disposing)
        End Sub

        Private NotInheritable Class DetachedSchema
            Implements IDataSourceColumnTypeDetector
            Private ReadOnly names As Dictionary(Of Integer, String)
            Private ReadOnly types As Dictionary(Of Integer, Type)
            Friend Sub New(names As Dictionary(Of Integer, String), types As Dictionary(Of Integer, Type))
                Me.names = names : Me.types = types
            End Sub
            Public Function GetColumnName(index As Integer, offset As Integer, range As CellRange) As String Implements IDataSourceColumnTypeDetector.GetColumnName
                Return names(offset)
            End Function
            Public Function GetColumnType(index As Integer, offset As Integer, range As CellRange) As Type Implements IDataSourceColumnTypeDetector.GetColumnType
                Return types(offset)
            End Function
        End Class

        Private NotInheritable Class GuardedRows
            Inherits BindingList(Of GuardedRow)
            Implements ITypedList
            Private ReadOnly properties As PropertyDescriptorCollection
            Friend Sub New(owner As ModelRangeDataSource, values As RangeDataSource)
                AllowNew = False : AllowEdit = False : AllowRemove = False
                Dim descriptors As New List(Of PropertyDescriptor)()
                For Each field As PropertyDescriptor In DirectCast(values, ITypedList).GetItemProperties(Nothing)
                    descriptors.Add(New GuardedField(owner, field))
                Next
                properties = New PropertyDescriptorCollection(descriptors.ToArray(), True)
                For Each row As Object In DirectCast(values, IList)
                    Items.Add(New GuardedRow(row))
                Next
            End Sub
            Public Function GetItemProperties(accessors As PropertyDescriptor()) As PropertyDescriptorCollection Implements ITypedList.GetItemProperties
                Return properties
            End Function
            Public Function GetListName(accessors As PropertyDescriptor()) As String Implements ITypedList.GetListName
                Return "Current model results"
            End Function
            Protected Overrides Sub InsertItem(index As Integer, item As GuardedRow)
                Throw New NotSupportedException("Calculated rows are read-only.")
            End Sub
            Protected Overrides Sub SetItem(index As Integer, item As GuardedRow)
                Throw New NotSupportedException("Calculated rows are read-only.")
            End Sub
            Protected Overrides Sub RemoveItem(index As Integer)
                Throw New NotSupportedException("Calculated rows are read-only.")
            End Sub
            Protected Overrides Sub ClearItems()
                Throw New NotSupportedException("Calculated rows are read-only.")
            End Sub
        End Class
        Private NotInheritable Class GuardedRow
            Friend ReadOnly Value As Object
            Friend Sub New(value As Object)
                Me.Value = value
            End Sub
        End Class
        Private NotInheritable Class GuardedField
            Inherits PropertyDescriptor
            Private ReadOnly owner As ModelRangeDataSource
            Private ReadOnly field As PropertyDescriptor
            Friend Sub New(owner As ModelRangeDataSource, field As PropertyDescriptor)
                MyBase.New(field.Name, Nothing)
                Me.owner = owner : Me.field = field
            End Sub
            Public Overrides ReadOnly Property ComponentType As Type
                Get
                    Return GetType(GuardedRow)
                End Get
            End Property
            Public Overrides ReadOnly Property PropertyType As Type
                Get
                    Return field.PropertyType
                End Get
            End Property
            Public Overrides ReadOnly Property IsReadOnly As Boolean
                Get
                    Return True
                End Get
            End Property
            Public Overrides Function GetValue(component As Object) As Object
                ' Null is a pending/closed value, never an old financial figure.
                If Not owner.HasCurrentValues Then Return Nothing
                Return field.GetValue(DirectCast(component, GuardedRow).Value)
            End Function
            Public Overrides Sub SetValue(component As Object, value As Object)
                Throw New NotSupportedException("Use the model change manager to edit inputs.")
            End Sub
            Public Overrides Function CanResetValue(component As Object) As Boolean
                Return False
            End Function
            Public Overrides Sub ResetValue(component As Object)
                Throw New NotSupportedException("Calculated values cannot be reset.")
            End Sub
            Public Overrides Function ShouldSerializeValue(component As Object) As Boolean
                Return False
            End Function
        End Class
    End Class
End Namespace
