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
    Private SuppressPosting As Boolean
    Private HistoryBinding As ModelPostingHistoryBinding
    Public Property SuppressAutomaticPosting As Boolean = False

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
        ConfigurePosting()

    End Sub
    Public Sub InitialiseStandard(RepID As String)

        ClearList()

        Dim ListItems As List(Of String) = RepositaryItems.GetList(RepID, ModelID)
        Properties.Items.AddRange(ListItems)
        CurrList = ListItems
        ConfigurePosting()

    End Sub
    Sub ProcesDefValue()
        RefreshFromWorkbook()
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

        If SuppressAutomaticPosting OrElse SuppressPosting Then Return
        Dim result = PostModelCellValue(ModelID, TargetWorksheet, TargetCell,
                                        If(EditValue Is Nothing OrElse Convert.IsDBNull(EditValue), Nothing, EditValue),
                                        "S", "Selection updated")
        If result.BError Then RefreshFromWorkbook()

    End Sub

    Private Sub ConfigurePosting()
        RemoveHandler MyBase.EditValueChanged, AddressOf ProcessChange
        AddHandler MyBase.EditValueChanged, AddressOf ProcessChange
        If HistoryBinding Is Nothing Then
            HistoryBinding = New ModelPostingHistoryBinding(Me, ModelID, TargetWorksheet, AddressOf RefreshFromWorkbook)
        End If
    End Sub

    Private Sub RefreshFromWorkbook()
        If String.IsNullOrWhiteSpace(TargetWorksheet) OrElse String.IsNullOrWhiteSpace(TargetCell) Then Return
        SuppressPosting = True
        Try
            Dim cell = ExcelModels(ModelID).WB.Worksheets(TargetWorksheet).Cells(TargetCell)
            EditValue = EditorValueFromCell(cell, "S")
        Finally
            SuppressPosting = False
        End Try
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message,
                                               ByVal keyData As Keys) As Boolean
        If TryProcessModelHistoryShortcut(Me, ModelID, keyData) Then Return True
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function


End Class
