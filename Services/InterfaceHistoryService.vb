Imports System.ComponentModel
Imports System.Configuration
Imports System.Drawing
Imports System.IO
Imports System.Threading
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
        Public Property ModelName As String

        Friend Property DestinationKind As InterfaceHistoryDestinationKind
        Friend Property ModelID As Integer
        Friend Property ModelInstanceID As Guid
        Friend Property VisitSequence As Long
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
        Public ReadOnly Property InstanceID As Guid = Guid.NewGuid()
        Private ReadOnly ItemsLock As New Object()
        Private ReadOnly Items As New List(Of InterfaceHistoryEntry)()
        Private IsDisposed As Boolean

        Public Event HistoryChanged As EventHandler

        Public Sub New(ByVal setModelID As Integer)
            ModelID = setModelID
            InterfaceHistoryCoordinator.Register(Me)
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
                .ModelInstanceID = InstanceID,
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
                .ModelInstanceID = InstanceID,
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
            InterfaceHistoryCoordinator.Unregister(Me)
        End Sub

        Private Sub RecordVisit(ByVal item As InterfaceHistoryEntry)
            If item Is Nothing OrElse IsDisposed Then Return
            item.VisitedAt = Now()
            item.VisitSequence = InterfaceHistoryCoordinator.NextVisitSequence()

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
                Abovo.SummitDiagnostics.WriteLine(
                    "Interface history notification failed: " & ex.ToString())
            End Try
            InterfaceHistoryCoordinator.NotifyChanged()
        End Sub
    End Class

    Friend NotInheritable Class InterfaceHistorySettings
        Inherits ApplicationSettingsBase

        Private Shared ReadOnly InstanceValue As InterfaceHistorySettings =
            CType(Synchronized(New InterfaceHistorySettings()), InterfaceHistorySettings)

        Public Shared ReadOnly Property [Default] As InterfaceHistorySettings
            Get
                Return InstanceValue
            End Get
        End Property

        <UserScopedSetting(), DefaultSettingValue("False")>
        Public Property ShowAllModels As Boolean
            Get
                Return CBool(Me(NameOf(ShowAllModels)))
            End Get
            Set(value As Boolean)
                Me(NameOf(ShowAllModels)) = value
            End Set
        End Property
    End Class

    Public NotInheritable Class InterfaceHistoryCoordinator

        Private Shared ReadOnly RegistryLock As New Object()
        Private Shared ReadOnly Services As New List(Of InterfaceHistoryService)()
        Private Shared VisitCounter As Long
        Private Shared PreferredAllModels As Boolean = LoadPreference()

        Public Shared Event Changed As EventHandler

        Private Sub New()
        End Sub

        Friend Shared Sub Register(ByVal service As InterfaceHistoryService)
            SyncLock RegistryLock
                Services.Add(service)
            End SyncLock
            NotifyChanged()
        End Sub

        Friend Shared Sub Unregister(ByVal service As InterfaceHistoryService)
            SyncLock RegistryLock
                Services.Remove(service)
            End SyncLock
            NotifyChanged()
        End Sub

        Friend Shared Function NextVisitSequence() As Long
            Return Interlocked.Increment(VisitCounter)
        End Function

        Public Shared ReadOnly Property ShowAllModels As Boolean
            Get
                Return PreferredAllModels
            End Get
        End Property

        Public Shared Sub SetShowAllModels(ByVal value As Boolean)
            If PreferredAllModels = value Then Return
            PreferredAllModels = value
            Try
                InterfaceHistorySettings.Default.ShowAllModels = value
                InterfaceHistorySettings.Default.Save()
            Catch ex As Exception
                Abovo.SummitDiagnostics.WriteLine(
                    "Unable to save interface history scope: " & ex.Message)
            End Try
            NotifyChanged()
        End Sub

        Public Shared ReadOnly Property OpenModelCount As Integer
            Get
                Return ActiveServices().Count
            End Get
        End Property

        Public Shared Function SnapshotItems(ByVal requestingModelID As Integer,
                                             ByVal includeAllModels As Boolean) As List(Of InterfaceHistoryEntry)
            Dim result As New List(Of InterfaceHistoryEntry)()
            For Each service As InterfaceHistoryService In ActiveServices()
                Dim items As List(Of InterfaceHistoryEntry) = service.SnapshotItems()
                For Each item As InterfaceHistoryEntry In items
                    If Not includeAllModels AndAlso item.ModelID <> requestingModelID Then Continue For
                    Dim model As FileManager.ExcelModel =
                        FileManager.ExcelModels(item.ModelID)
                    item.ModelName = DescribeModel(model)
                    result.Add(item)
                Next
            Next
            result.Sort(
                Function(left As InterfaceHistoryEntry, right As InterfaceHistoryEntry) As Integer
                    Return right.VisitSequence.CompareTo(left.VisitSequence)
                End Function)
            Return result
        End Function

        Friend Shared Sub NotifyChanged()
            Try
                RaiseEvent Changed(Nothing, EventArgs.Empty)
            Catch ex As Exception
                Abovo.SummitDiagnostics.WriteLine(
                    "Interface history view notification failed: " & ex.ToString())
            End Try
        End Sub

        Private Shared Function ActiveServices() As List(Of InterfaceHistoryService)
            Dim registered As InterfaceHistoryService()
            SyncLock RegistryLock
                registered = Services.ToArray()
            End SyncLock

            Dim result As New List(Of InterfaceHistoryService)()
            Dim models() As FileManager.ExcelModel = FileManager.ExcelModels
            If models Is Nothing Then Return result

            For Each service As InterfaceHistoryService In registered
                For Each model As FileManager.ExcelModel In models
                    If model Is Nothing OrElse model.IsClosing OrElse
                       model.WBStructure Is Nothing Then Continue For
                    If Object.ReferenceEquals(model.InterfaceHistory, service) Then
                        result.Add(service)
                        Exit For
                    End If
                Next
            Next
            Return result
        End Function

        Private Shared Function DescribeModel(ByVal model As FileManager.ExcelModel) As String
            Dim modelType As String = If(model.Profile Is Nothing,
                                         "Model",
                                         model.Profile.DisplayName)
            Dim companyName As String = If(model.WBStructure.CompanyName, String.Empty).Trim()
            Dim fileName As String = If(String.IsNullOrWhiteSpace(model.FileName),
                                        String.Empty,
                                        Path.GetFileName(model.FileName))
            Dim description As String = If(String.IsNullOrWhiteSpace(companyName),
                                           fileName,
                                           companyName)
            If Not String.IsNullOrWhiteSpace(companyName) AndAlso
               Not String.IsNullOrWhiteSpace(fileName) Then
                description &= " (" & fileName & ")"
            End If
            Return modelType & ": " & description
        End Function

        Private Shared Function LoadPreference() As Boolean
            Try
                Return InterfaceHistorySettings.Default.ShowAllModels
            Catch
                Return False
            End Try
        End Function
    End Class

    Public NotInheritable Class InterfaceHistoryView
        Inherits XtraUserControl

        Private ReadOnly ModelID As Integer
        Private ReadOnly ViewItems As New BindingList(Of InterfaceHistoryEntry)()
        Private ReadOnly HistoryGrid As New GridControl()
        Private ReadOnly HistoryGridView As New GridView()
        Private ReadOnly ScopeSelector As New ComboBoxEdit()
        Private ChangingScopeSelector As Boolean
        Private IsDisposedLocally As Boolean

        Public Sub New(ByVal setModelID As Integer,
                       ByVal service As InterfaceHistoryService)
            ModelID = setModelID
            If service Is Nothing Then Throw New ArgumentNullException(NameOf(service))
            BuildLayout()
            AddHandler InterfaceHistoryCoordinator.Changed, AddressOf HistoryChanged
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

            Dim multipleModels As Boolean = InterfaceHistoryCoordinator.OpenModelCount > 1
            SuspendLayout()
            HistoryGrid.BeginUpdate()
            Try
                ScopeSelector.Visible = multipleModels
                Dim includeAllModels As Boolean =
                    multipleModels AndAlso InterfaceHistoryCoordinator.ShowAllModels
                ChangingScopeSelector = True
                Try
                    ScopeSelector.SelectedIndex = If(includeAllModels, 1, 0)
                Finally
                    ChangingScopeSelector = False
                End Try

                HistoryGridView.BeginDataUpdate()
                Try
                    ViewItems.RaiseListChangedEvents = False
                    ViewItems.Clear()
                    For Each item As InterfaceHistoryEntry In
                        InterfaceHistoryCoordinator.SnapshotItems(ModelID, includeAllModels)
                        ViewItems.Add(item)
                    Next
                Finally
                    ViewItems.RaiseListChangedEvents = True
                    ViewItems.ResetBindings()
                    HistoryGridView.EndDataUpdate()
                End Try

                Dim modelColumn = HistoryGridView.Columns.ColumnByFieldName("ModelName")
                modelColumn.Visible = includeAllModels
                modelColumn.VisibleIndex = If(includeAllModels, 3, -1)
                If HistoryGridView.RowCount > 0 Then HistoryGridView.MoveFirst()
                'Keep the established column proportions while records are replaced.
                'Only a genuine model-scope change should change column visibility.
            Finally
                HistoryGrid.EndUpdate()
                ResumeLayout(True)
            End Try
        End Sub

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso Not IsDisposedLocally Then
                IsDisposedLocally = True
                RemoveHandler InterfaceHistoryCoordinator.Changed, AddressOf HistoryChanged
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
            ObjectFormatter.FormatInformationalGrid(HistoryGrid, HistoryGridView)

            HistoryGridView.Columns.AddVisible("VisitedAt", "Time")
            HistoryGridView.Columns.AddVisible("InterfaceName", "Interface")
            HistoryGridView.Columns.AddVisible("AreaName", "Area")
            HistoryGridView.Columns.AddVisible("ModelName", "Model")

            Dim timeColumn = HistoryGridView.Columns.ColumnByFieldName("VisitedAt")
            timeColumn.DisplayFormat.FormatType = FormatType.DateTime
            timeColumn.DisplayFormat.FormatString = "HH:mm:ss"
            timeColumn.Width = 72
            HistoryGridView.Columns.ColumnByFieldName("InterfaceName").Width = 220
            HistoryGridView.Columns.ColumnByFieldName("AreaName").Width = 100
            HistoryGridView.Columns.ColumnByFieldName("ModelName").Width = 200

            ScopeSelector.Properties.Items.AddRange(
                New Object() {"This model", "All models"})
            ScopeSelector.Properties.TextEditStyle =
                DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
            ScopeSelector.SelectedIndex = 0
            ScopeSelector.Dock = DockStyle.Top
            ScopeSelector.Visible = False
            AddHandler ScopeSelector.SelectedIndexChanged, AddressOf ScopeSelectionChanged

            AddHandler HistoryGrid.MouseDoubleClick, AddressOf HistoryGridDoubleClick
            AddHandler HistoryGridView.KeyDown, AddressOf HistoryGridKeyDown
            Controls.Add(HistoryGrid)
            Controls.Add(ScopeSelector)
        End Sub

        Private Sub ScopeSelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
            If ChangingScopeSelector Then Return
            InterfaceHistoryCoordinator.SetShowAllModels(ScopeSelector.SelectedIndex = 1)
        End Sub

        Private Sub HistoryChanged(ByVal sender As Object, ByVal e As EventArgs)
            RefreshHistory()
        End Sub

        Private Sub HistoryGridDoubleClick(ByVal sender As Object, ByVal e As MouseEventArgs)
            If e.Button <> MouseButtons.Left Then Return
            Dim hitInfo = HistoryGridView.CalcHitInfo(e.Location)
            If Not hitInfo.InRow AndAlso Not hitInfo.InRowCell Then Return
            ActivateEntry(TryCast(HistoryGridView.GetRow(hitInfo.RowHandle),
                                  InterfaceHistoryEntry))
        End Sub

        Private Sub HistoryGridKeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            If e.KeyCode <> Keys.Enter Then Return
            ActivateFocusedEntry()
            e.Handled = True
            e.SuppressKeyPress = True
        End Sub

        Private Sub ActivateFocusedEntry()
            ActivateEntry(TryCast(HistoryGridView.GetFocusedRow(),
                                  InterfaceHistoryEntry))
        End Sub

        Private Sub ActivateEntry(ByVal entry As InterfaceHistoryEntry)
            If entry Is Nothing Then Return

            If FileManager.ExcelModels Is Nothing OrElse
               entry.ModelID < 0 OrElse
               entry.ModelID >= FileManager.ExcelModels.Length Then Return

            Dim targetModel As FileManager.ExcelModel =
                FileManager.ExcelModels(entry.ModelID)
            If targetModel Is Nothing OrElse targetModel.IsClosing OrElse
               targetModel.InterfaceHistory Is Nothing OrElse
               targetModel.InterfaceHistory.InstanceID <> entry.ModelInstanceID OrElse
               targetModel.WBInterface Is Nothing Then Return

            targetModel.WBInterface.ActivateHistoryEntry(entry)
        End Sub
    End Class

End Namespace
