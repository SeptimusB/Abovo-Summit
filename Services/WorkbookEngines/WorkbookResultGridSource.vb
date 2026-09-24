Option Strict On

Imports System.ComponentModel
Imports System.Threading

Namespace Abovo.WorkbookEngines
    ''' <summary>
    ''' Read-only native-grid binding over an accepted engine result. It never
    ''' writes a worksheet or invokes a second calculation engine. The controller
    ''' must marshal Publish/InvalidateDisplay to the grid's owner thread.
    ''' </summary>
    Public NotInheritable Class WorkbookResultGridSource
        Inherits BindingList(Of WorkbookResultGridRow)
        Implements ITypedList

        Private ReadOnly ownerThread As Integer = Thread.CurrentThread.ManagedThreadId
        Private ReadOnly session As WorkbookCalculationSession
        Private ReadOnly fields As PropertyDescriptorCollection
        Private result As WorkbookCalculationResult
        Private block As WorkbookValueBlock
        Private Shared ReadOnly unavailable As New WorkbookCellError("#N/A")
        Public ReadOnly Property Area As WorkbookReadArea

        Public Sub New(session As WorkbookCalculationSession, area As WorkbookReadArea)
            If session Is Nothing Then Throw New ArgumentNullException(NameOf(session))
            If area Is Nothing Then Throw New ArgumentNullException(NameOf(area))
            Me.session = session : Me.Area = area
            AllowEdit = False : AllowNew = False : AllowRemove = False
            Dim descriptors As New List(Of PropertyDescriptor)()
            For c As Integer = 0 To area.Columns - 1
                descriptors.Add(New ValueProperty(c))
            Next
            fields = New PropertyDescriptorCollection(descriptors.ToArray(), True)
            For r As Integer = 0 To area.Rows - 1
                Items.Add(New WorkbookResultGridRow(Me, r))
            Next
        End Sub

        Public ReadOnly Property HasCurrentValues As Boolean
            Get
                RequireOwner()
                Return result IsNot Nothing AndAlso session.IsCurrent(result)
            End Get
        End Property

        Public Function TryPublish(nextResult As WorkbookCalculationResult) As Boolean
            RequireOwner()
            If Not session.IsCurrent(nextResult) Then Return False
            Dim matched As WorkbookValueBlock = Nothing
            For Each candidate In nextResult.Blocks
                Dim actual = candidate.Area
                If String.Equals(actual.Worksheet, Area.Worksheet, StringComparison.OrdinalIgnoreCase) AndAlso
                   actual.Row = Area.Row AndAlso actual.Column = Area.Column AndAlso
                   actual.Rows = Area.Rows AndAlso actual.Columns = Area.Columns Then
                    If matched IsNot Nothing Then Throw New ArgumentException("Duplicate display rectangle in engine result.", NameOf(nextResult))
                    matched = candidate
                End If
            Next
            If matched Is Nothing Then Throw New ArgumentException("Engine result does not contain the display rectangle.", NameOf(nextResult))
            If Not session.IsCurrent(nextResult) Then Return False
            block = matched : result = nextResult
            ResetBindings()
            Return HasCurrentValues
        End Function

        Public Sub InvalidateDisplay()
            RequireOwner()
            block = Nothing : result = Nothing
            ResetBindings()
        End Sub

        Friend Function ReadValue(row As Integer, column As Integer) As Object
            RequireOwner()
            If column < 0 OrElse column >= Area.Columns Then Throw New ArgumentOutOfRangeException(NameOf(column))
            ' Even if the controller's refresh is still queued, never give the
            ' grid a value from an obsolete or failed calculation revision.
            If Not HasCurrentValues Then Return unavailable
            Return block.ValueAt(row, column)
        End Function

        Private Sub RequireOwner()
            If Thread.CurrentThread.ManagedThreadId <> ownerThread Then Throw New InvalidOperationException("Grid values must be accessed on their UI owner thread.")
        End Sub

        Protected Overrides Sub InsertItem(index As Integer, item As WorkbookResultGridRow)
            Throw New NotSupportedException("Engine-result rows are read-only.")
        End Sub
        Protected Overrides Sub SetItem(index As Integer, item As WorkbookResultGridRow)
            Throw New NotSupportedException("Engine-result rows are read-only.")
        End Sub
        Protected Overrides Sub RemoveItem(index As Integer)
            Throw New NotSupportedException("Engine-result rows are read-only.")
        End Sub
        Protected Overrides Sub ClearItems()
            Throw New NotSupportedException("Engine-result rows are read-only.")
        End Sub
        Public Function GetItemProperties(listAccessors As PropertyDescriptor()) As PropertyDescriptorCollection Implements ITypedList.GetItemProperties
            Return fields
        End Function
        Public Function GetListName(listAccessors As PropertyDescriptor()) As String Implements ITypedList.GetListName
            Return Area.Worksheet & "!" & Area.Address
        End Function

        Private NotInheritable Class ValueProperty
            Inherits PropertyDescriptor
            Private ReadOnly column As Integer
            Friend Sub New(column As Integer)
                MyBase.New("Column" & column.ToString(Globalization.CultureInfo.InvariantCulture), Nothing)
                Me.column = column
            End Sub
            Public Overrides ReadOnly Property ComponentType As Type
                Get
                    Return GetType(WorkbookResultGridRow)
                End Get
            End Property
            Public Overrides ReadOnly Property PropertyType As Type
                Get
                    Return GetType(Object)
                End Get
            End Property
            Public Overrides ReadOnly Property IsReadOnly As Boolean
                Get
                    Return True
                End Get
            End Property
            Public Overrides Function GetValue(component As Object) As Object
                Return DirectCast(component, WorkbookResultGridRow).GetValue(column)
            End Function
            Public Overrides Sub SetValue(component As Object, value As Object)
                Throw New NotSupportedException("Route edits through the model change manager, not result cells.")
            End Sub
            Public Overrides Function CanResetValue(component As Object) As Boolean
                Return False
            End Function
            Public Overrides Sub ResetValue(component As Object)
                Throw New NotSupportedException("Engine-result cells are read-only.")
            End Sub
            Public Overrides Function ShouldSerializeValue(component As Object) As Boolean
                Return False
            End Function
        End Class
    End Class

    Public NotInheritable Class WorkbookResultGridRow
        Private ReadOnly source As WorkbookResultGridSource
        Private ReadOnly offset As Integer
        Friend Sub New(source As WorkbookResultGridSource, offset As Integer)
            Me.source = source : Me.offset = offset
        End Sub
        Public ReadOnly Property WorksheetRow As Integer
            Get
                Return source.Area.Row + offset
            End Get
        End Property
        Public Function GetValue(columnOffset As Integer) As Object
            Return source.ReadValue(offset, columnOffset)
        End Function
    End Class
End Namespace
