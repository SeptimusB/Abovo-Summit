Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraVerticalGrid

Partial Public Class DataInterfaceTemplate
    Private ReadOnly authoringHooks As New HashSet(Of Control)
    Private authoringTabsHooked As Boolean
    Friend Event AuthoringSourceSelected(sheet As String, address As String)

    Friend ReadOnly Property IsAuthoringPreview As Boolean
        Get
            Return String.Equals(InterfaceMode, "StructurePreview", StringComparison.Ordinal)
        End Get
    End Property

    Friend Sub SelectAuthoringSection(index As Integer)
        If Not IsAuthoringPreview Then Return
        If index >= 0 AndAlso index < XtraTabControlNewGIT.TabPages.Count Then
            XtraTabControlNewGIT.SelectedTabPageIndex = index
            EnsureSectionBuilt(index)
        End If
        ConfigureAuthoringPreview()
    End Sub

    Private Sub ConfigureAuthoringPreview()
        If Not IsAuthoringPreview OrElse IsDisposed Then Return
        WindowsUIButtonPanelActions.Visible = False
        If Not authoringTabsHooked Then
            authoringTabsHooked = True
            AddHandler XtraTabControlNewGIT.SelectedPageChanged, Sub() ConfigureAuthoringPreview()
        End If
        For Each control As Control In FindChildControls(Of Control)(Me)
            Dim editor = TryCast(control, BaseEdit)
            If editor IsNot Nothing Then editor.Properties.ReadOnly = True
            If TypeOf control Is SimpleButton Then control.Enabled = False
            Dim grid = TryCast(control, GridControl)
            Dim mapped = TryCast(control, ReadOnlyMappedTableGrid)
            If mapped IsNot Nothing Then mapped.IsAuthoringPreview = True
            If grid IsNot Nothing Then
                For Each view As DevExpress.XtraGrid.Views.Base.BaseView In grid.ViewCollection
                    Dim rows = TryCast(view, GridView)
                    If rows IsNot Nothing Then
                        rows.OptionsBehavior.Editable = False
                        rows.OptionsBehavior.ReadOnly = True
                    End If
                Next
            End If
            Dim vertical = TryCast(control, VGridControl)
            If vertical IsNot Nothing Then vertical.OptionsBehavior.Editable = False
            If (grid IsNot Nothing OrElse vertical IsNot Nothing) AndAlso authoringHooks.Add(control) Then
                AddHandler control.MouseUp, AddressOf AuthoringMouseUp
            End If
        Next
    End Sub

    Private Sub AuthoringMouseUp(sender As Object, e As MouseEventArgs)
        Dim sheet As String = Nothing, address As String = Nothing
        If TryGetAuthoringSource(DirectCast(sender, Control), sheet, address) Then RaiseEvent AuthoringSourceSelected(sheet, address)
    End Sub

    Friend Function TryGetAuthoringSource(target As Control, ByRef sheet As String, ByRef address As String) As Boolean
        If Not IsAuthoringPreview OrElse DataPres Is Nothing Then Return False
        Dim grid = TryCast(target, GridControl)
        Dim mapped = TryCast(target, ReadOnlyMappedTableGrid)
        If mapped IsNot Nothing Then
            Dim source = mapped.AuthoringSelectedCell()
            If source Is Nothing Then Return False
            sheet = source.Worksheet.Name
            address = source.GetReferenceA1()
            Return True
        End If
        Dim view = TryCast(grid?.FocusedView, GridView)
        If view IsNot Nothing Then
            Dim liveCell As DevExpress.Spreadsheet.Cell = Nothing
            If TryGetLiveGridSourceCell(view, view.FocusedRowHandle, view.FocusedColumn, liveCell) Then
                sheet = liveCell.Worksheet.Name
                address = liveCell.GetReferenceA1()
                Return True
            End If
        End If
        Dim choices = GetSelectedClipboardDataCells(target)
        If choices.Count = 0 Then Return False
        Dim choice = choices(0)
        If choice.DataSetIndex < 0 OrElse choice.DataSetIndex >= DataPres.DataSets.Count Then Return False
        Dim data = DataPres.DataSets(choice.DataSetIndex)
        If data.DataRows Is Nothing OrElse choice.DataRowIndex >= data.DataRows.Length Then Return False
        Dim cells = data.DataRows(choice.DataRowIndex).DataCells
        If cells Is Nothing OrElse choice.DataColumnIndex >= cells.Length Then Return False
        sheet = cells(choice.DataColumnIndex).SourceSheet
        address = cells(choice.DataColumnIndex).SourceAddress
        Return Not String.IsNullOrWhiteSpace(sheet) AndAlso Not String.IsNullOrWhiteSpace(address)
    End Function
End Class
