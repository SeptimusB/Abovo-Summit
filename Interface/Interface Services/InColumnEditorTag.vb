Imports Abovo.AbovoExtendedDEControls
Imports DevExpress.XtraGrid.Views.BandedGrid
Namespace Abovo
    Public Class InColumnEditorTagCombo

        Public EditingNRName As String
        Public EditingNRIndexPosition As Integer = -1
        Public NROrientation As Orientation
        Public EditorType As String
        Public InitialValue As Object
        Public LastEditorValue As Object
        Public EditorFormat As String
        Public ParentBandedGridColumn As BandedGridColumn
        Public LinkedComboBoxEdit As AbovoDEHeaderComboBox
        Public InPlaceColumnHelper As ColumnInplaceEditorHelper
        Public InPlaceVGridRowHelper As VGridRowInplaceEditorHelper

    End Class

    Public Class InColumnEditorTagDateEdit

        Public EditingNRName As String
        Public EditingNRIndexPosition As Integer = -1
        Public NROrientation As Orientation
        Public EditorType As String
        Public InitialValue As Object
        Public LastEditorValue As Object
        Public EditorFormat As String
        Public ParentBandedGridColumn As BandedGridColumn
        Public LinkedDateBoxEdit As AbovoDEHeaderDateBox
        Public InPlaceColumnHelper As ColumnInplaceEditorHelper
        Public InPlaceVGridRowHelper As VGridRowInplaceEditorHelper

    End Class


    Public Class InColumnNRActionTag

        Public EditingNRName As String
        Public EditingMethod As String
        Public NROrientation As Orientation
        Public MinElements As String = 5
        Public MaxElements As String = 20
        Public EditingMsg As String
        Public PostEditAction As String
        Public PostEditAMsg As String

    End Class

End Namespace