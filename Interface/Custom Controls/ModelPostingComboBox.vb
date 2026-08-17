Imports System.ComponentModel
Imports Abovo
Imports Abovo.FileManager
Imports Abovo.RepositaryItems
Imports DevExpress.XtraEditors
Public Class ModelPostingComboBox

    Inherits DevExpress.XtraEditors.ComboBoxEdit

    Private ModelID As Integer
    Private TargetWorksheet As String
    Private TargetCell As String
    Private CurrList As List(Of String)
    Private LitmitToList As Boolean = True

    Public Property SetLimitToList As Boolean
        Get
            Return LitmitToList
        End Get
        Set(value As Boolean)
            LitmitToList = value
            If value Then
                'Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
            Else
                'Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard
            End If
        End Set

    End Property

    Public Sub InitialiseFromNRP(NRName As String)

        ClearList()

        Dim ListItems As List(Of String) = RepositaryItems.GetListFromNR(NRName, ModelID)
        Properties.Items.AddRange(ListItems)
        CurrList = ListItems
        AddHandler MyBase.EditValueChanged, AddressOf ProcessChange

    End Sub
    Public Sub InitialiseStandard(RepID As String)

        ClearList()

        Dim ListItems As List(Of String) = RepositaryItems.GetList(RepID, ModelID)
        Properties.Items.AddRange(ListItems)
        CurrList = ListItems
        AddHandler MyBase.EditValueChanged, AddressOf ProcessChange

    End Sub
    Sub ProcesDefValue()

        Try

            EditValue = ExcelModels(ModelID).WB.Worksheets(TargetWorksheet).Cells(TargetCell).DisplayText

        Catch ex As Exception

        End Try


    End Sub
    Public Property SetTargetCell As String
        Get
            Return TargetCell
        End Get
        Set(value As String)
            TargetCell = value
        End Set
    End Property
    Public Property SetTargetWorksheet As String
        Get
            Return TargetWorksheet
        End Get
        Set(value As String)
            TargetWorksheet = value
        End Set
    End Property

    Public Property SetModelID As Integer
        Get
            Return ModelID
        End Get
        Set(value As Integer)
            ModelID = value
        End Set
    End Property
    Public Sub ClearList()

        Properties.Items.Clear()

    End Sub
    Protected Sub ProcessChange(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim NewVal As String = SelectedText

        If Not String.IsNullOrEmpty(NewVal) Then

            ExcelModels(ModelID).WB.Worksheets(TargetWorksheet).Cells(TargetCell).Value = NewVal

        End If

    End Sub


End Class
