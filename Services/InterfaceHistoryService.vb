Imports System.ComponentModel
Imports System.Drawing
Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Grid

Namespace Abovo

    Public Enum InterfaceHistoryDestinationKind
        GroupInterface
        FinancialForecastReturn
        StressTest
    End Enum

    Public NotInheritable Class InterfaceHistoryEntry
        Public Property VisitedAt As DateTime
        Public Property InterfaceName As String
        Public Property AreaName As String

        Friend Property DestinationKind As InterfaceHistoryDestinationKind
        Friend Property ModelID As Integer
        Friend Property GSID As Integer
        Friend Property CSID As Integer
        Friend Property ShowSpecial As Boolean
        Friend Property SpecialData As String
        Friend Property DestinationKey As String
        Friend Property HostGroupInterface As GroupInterfaceTemplate

        Friend Function CloneEntry() As InterfaceHistoryEntry
            Return DirectCast(MemberwiseClone(), InterfaceHistoryEntry)
        End Function
    End Class

    Public NotInheritable Class InterfaceHistoryService
        Implements IDisposable

        Private Const MaximumEntries As Integer = 50
        Private ReadOnly ModelID As Integer
        Private ReadOnly ItemsLock As New Object()
        Private ReadOnly Items As New List(Of InterfaceHistoryEntry)()
        Private IsDisposed As Boolean

        Public Event HistoryChanged As EventHandler

        Public Sub New(ByVal setModelID As Integer)
            ModelID = setModelID
        End Sub

        Public Sub RecordGroupInterface(ByVal host As GroupInterfaceTemplate,
                                        ByVal gsid As Integer,
                                        ByVal csid As Integer,
                                        ByVal interfaceName As String,
                                        ByVal areaName As String,
                                        ByVal showSpecial As Boolean,
                                        ByVal specialData As String)
            RecordVisit(New InterfaceHistoryEntry With {
                .DestinationKind = InterfaceHistoryDestinationKind.GroupInterface,
                .ModelID = ModelID,
                .GSID = gsid,
                .CSID = csid,
                .InterfaceName = interfaceName,
                .AreaName = areaName,
                .ShowSpecial = showSpecial,
                .SpecialData = If(specialData, String.Empty),
                .DestinationKey = "Group|" & gsid.ToString() & "|" & csid.ToString(),
                .HostGroupInterface = host})
        End Sub

        Public Sub RecordStandalone(ByVal destinationKind As InterfaceHistoryDestinationKind,
                                    ByVal interfaceName As String,
                                    ByVal areaName As String)
            RecordVisit(New InterfaceHistoryEntry With {
                .DestinationKind = destinationKind,
                .ModelID = ModelID,
                .InterfaceName = interfaceName,
                .AreaName = areaName,
                .DestinationKey = "Standalone|" & destinationKind.ToString()})
        End Sub

        Public Function SnapshotItems() As List(Of InterfaceHistoryEntry)
            Dim result As New List(Of InterfaceHistoryEntry)()
            SyncLock ItemsLock
                For Each item As InterfaceHistoryEntry In Items
                    result.Add(item.CloneEntry())
                Next
            End SyncLock
            Return result
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If IsDisposed Then Return
            SyncLock ItemsLock
                Items.Clear()
                IsDisposed = True
            End SyncLock
        End Sub

        Private Sub RecordVisit(ByVal item As InterfaceHistoryEntry)
            If item Is Nothing OrElse IsDisposed Then Return
            item.VisitedAt = Now()

            SyncLock ItemsLock
                Dim existing As InterfaceHistoryEntry =
                    Items.FirstOrDefault(
                        Function(candidate As InterfaceHistoryEntry) As Boolean
                            Return String.Equals(candidate.DestinationKey,
                                                 item.DestinationKey,
                                                 StringComparison.Ordinal)
                        End Function)
                If existing IsNot Nothing Then Items.Remove(existing)
                Items.Insert(0, item)
                While Items.Count > MaximumEntries
                    Items.RemoveAt(Items.Count - 1)
                End While
            End SyncLock

            Try
                RaiseEvent HistoryChanged(Me, EventArgs.Empty)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(
                    "Interface history notification failed: " & ex.ToString())
            End Try
        End Sub
    End Class

    Public NotInheritable Class InterfaceHistoryView
        Inherits XtraUserControl

        Private ReadOnly ModelID As Integer
        Private ReadOnly HistoryService As InterfaceHistoryService
        Private ReadOnly ViewItems As New BindingList(Of InterfaceHistoryEntry)()
        Private ReadOnly HistoryGrid As New GridControl()
        Private ReadOnly HistoryGridView As New GridView()
        Private IsDisposedLocally As Boolean

        Public Sub New(ByVal setModelID As Integer,
                       ByVal service As InterfaceHistoryService)
            ModelID = setModelID
            HistoryService = service
            BuildLayout()
            AddHandler HistoryService.HistoryChanged, AddressOf HistoryChanged
            RefreshHistory()
        End Sub

        Public Sub RefreshHistory()
            If IsDisposedLocally Then Return
            If InvokeRequired Then
                Try
                    BeginInvoke(New MethodInvoker(AddressOf RefreshHistory))
                Catch ex As ObjectDisposedException
                    'The owning GIT is closing; there is no view left to refresh.
                Catch ex As InvalidOperationException
                    'The owning GIT is closing; there is no view left to refresh.
                End Try
                Return
            End If

            HistoryGridView.BeginDataUpdate()
            Try
                ViewItems.RaiseListChangedEvents = False
                ViewItems.Clear()
                For Each item As InterfaceHistoryEntry In HistoryService.SnapshotItems()
                    ViewItems.Add(item)
                Next
            Finally
                ViewItems.RaiseListChangedEvents = True
                ViewItems.ResetBindings()
                HistoryGridView.EndDataUpdate()
            End Try

            HistoryGridView.BestFitColumns()
            If HistoryGridView.RowCount > 0 Then HistoryGridView.MoveFirst()
        End Sub

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso Not IsDisposedLocally Then
                IsDisposedLocally = True
                RemoveHandler HistoryService.HistoryChanged, AddressOf HistoryChanged
                HistoryGridView.Dispose()
                HistoryGrid.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        Private Sub BuildLayout()
            Dock = DockStyle.Fill
            Padding = New Padding(4)

            HistoryGrid.MainView = HistoryGridView
            HistoryGrid.ViewCollection.Add(HistoryGridView)
            HistoryGrid.DataSource = ViewItems
            HistoryGrid.Dock = DockStyle.Fill

            HistoryGridView.OptionsBehavior.Editable = False
            HistoryGridView.OptionsBehavior.ReadOnly = True
            HistoryGridView.OptionsSelection.EnableAppearanceFocusedCell = False
            HistoryGridView.OptionsView.ShowGroupPanel = False
            HistoryGridView.OptionsView.ShowIndicator = False
            HistoryGridView.OptionsView.ShowAutoFilterRow = False
            HistoryGridView.OptionsView.ColumnAutoWidth = True
            HistoryGridView.OptionsView.EnableAppearanceEvenRow = True
            HistoryGridView.Appearance.EvenRow.BackColor = Color.FromArgb(248, 250, 252)

            Dim fontSize As Single = Math.Max(8.0F, 9.0F * PresentationScaleManager.UserScale)
            HistoryGridView.Appearance.Row.Font = New Font("Segoe UI", fontSize)
            HistoryGridView.Appearance.HeaderPanel.Font = New Font("Segoe UI", fontSize, FontStyle.Bold)

            HistoryGridView.Columns.AddVisible("VisitedAt", "Time")
            HistoryGridView.Columns.AddVisible("InterfaceName", "Interface")
            HistoryGridView.Columns.AddVisible("AreaName", "Area")

            Dim timeColumn = HistoryGridView.Columns.ColumnByFieldName("VisitedAt")
            timeColumn.DisplayFormat.FormatType = FormatType.DateTime
            timeColumn.DisplayFormat.FormatString = "HH:mm:ss"
            timeColumn.Width = 72
            HistoryGridView.Columns.ColumnByFieldName("InterfaceName").Width = 220
            HistoryGridView.Columns.ColumnByFieldName("AreaName").Width = 100

            AddHandler HistoryGridView.DoubleClick, AddressOf HistoryGridDoubleClick
            AddHandler HistoryGridView.KeyDown, AddressOf HistoryGridKeyDown
            Controls.Add(HistoryGrid)
        End Sub

        Private Sub HistoryChanged(ByVal sender As Object, ByVal e As EventArgs)
            RefreshHistory()
        End Sub

        Private Sub HistoryGridDoubleClick(ByVal sender As Object, ByVal e As EventArgs)
            Dim mouseArgs As MouseEventArgs = TryCast(e, MouseEventArgs)
            If mouseArgs Is Nothing Then Return
            Dim hitInfo = HistoryGridView.CalcHitInfo(mouseArgs.Location)
            If Not hitInfo.InRow AndAlso Not hitInfo.InRowCell Then Return
            ActivateFocusedEntry()
        End Sub

        Private Sub HistoryGridKeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            If e.KeyCode <> Keys.Enter Then Return
            ActivateFocusedEntry()
            e.Handled = True
            e.SuppressKeyPress = True
        End Sub

        Private Sub ActivateFocusedEntry()
            Dim entry As InterfaceHistoryEntry =
                TryCast(HistoryGridView.GetFocusedRow(), InterfaceHistoryEntry)
            If entry Is Nothing Then Return

            If FileManager.ExcelModels Is Nothing OrElse
               ModelID < 0 OrElse
               ModelID >= FileManager.ExcelModels.Length OrElse
               FileManager.ExcelModels(ModelID) Is Nothing OrElse
               FileManager.ExcelModels(ModelID).WBInterface Is Nothing Then Return

            FileManager.ExcelModels(ModelID).WBInterface.ActivateHistoryEntry(entry)
        End Sub
    End Class

End Namespace
