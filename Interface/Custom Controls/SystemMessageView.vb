Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Grid

Namespace Abovo
    Public NotInheritable Class SystemMessageView
        Inherits XtraUserControl

        Private ReadOnly ModelID As Integer
        Private ReadOnly MessageManager As SystemMessageManager
        Private ReadOnly ViewItems As New BindingList(Of SystemMessageRecord)()
        Private ReadOnly MessageGrid As New GridControl()
        Private ReadOnly MessageGridView As New GridView()
        Private ReadOnly StatusLabel As New LabelControl()
        Private IsDisposedLocally As Boolean

        Public Sub New(ByVal setModelID As Integer)
            ModelID = setModelID
            MessageManager = SystemMessageManager.Acquire(ModelID)
            BuildLayout()
            AddHandler MessageManager.MessagesChanged, AddressOf MessagesChanged
            RefreshMessages()
        End Sub

        Public Sub RefreshMessages()
            If IsDisposedLocally Then Return
            If InvokeRequired Then
                BeginInvoke(New MethodInvoker(AddressOf RefreshMessages))
                Return
            End If

            MessageGridView.BeginDataUpdate()
            Try
                ViewItems.RaiseListChangedEvents = False
                ViewItems.Clear()
                For Each item As SystemMessageRecord In MessageManager.SnapshotItems()
                    ViewItems.Add(item)
                Next
            Finally
                ViewItems.RaiseListChangedEvents = True
                ViewItems.ResetBindings()
                MessageGridView.EndDataUpdate()
            End Try
            StatusLabel.Text = ViewItems.Count.ToString() & " message" & If(ViewItems.Count = 1, String.Empty, "s")
            If MessageGridView.RowCount > 0 Then MessageGridView.MoveLast()
        End Sub

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso Not IsDisposedLocally Then
                IsDisposedLocally = True
                RemoveHandler MessageManager.MessagesChanged, AddressOf MessagesChanged
                MessageManager.Dispose()
                MessageGridView.Dispose()
                MessageGrid.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        Private Sub BuildLayout()
            Dock = DockStyle.Fill
            Padding = New Padding(6)

            Dim commandPanel As New FlowLayoutPanel With {
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Dock = DockStyle.Top,
                .FlowDirection = FlowDirection.LeftToRight,
                .Padding = New Padding(0, 0, 0, 6),
                .WrapContents = True}

            commandPanel.Controls.Add(CreateButton("Refresh", AddressOf RefreshButtonClick))
            commandPanel.Controls.Add(CreateButton("Copy", AddressOf CopyButtonClick))
            commandPanel.Controls.Add(CreateButton("Save", AddressOf SaveButtonClick))
            commandPanel.Controls.Add(CreateButton("Email", AddressOf EmailButtonClick))
            StatusLabel.AutoSizeMode = LabelAutoSizeMode.Vertical
            StatusLabel.Padding = New Padding(8, 7, 0, 0)
            commandPanel.Controls.Add(StatusLabel)

            MessageGrid.MainView = MessageGridView
            MessageGrid.ViewCollection.Add(MessageGridView)
            MessageGrid.DataSource = ViewItems
            MessageGrid.Dock = DockStyle.Fill

            MessageGridView.OptionsBehavior.Editable = False
            MessageGridView.OptionsBehavior.ReadOnly = True
            MessageGridView.OptionsSelection.MultiSelect = True
            MessageGridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.RowSelect
            MessageGridView.OptionsView.ShowAutoFilterRow = True
            MessageGridView.OptionsView.ShowGroupPanel = False
            MessageGridView.OptionsView.ShowIndicator = False
            MessageGridView.OptionsView.RowAutoHeight = True
            MessageGridView.OptionsView.ColumnAutoWidth = False
            MessageGridView.OptionsView.EnableAppearanceEvenRow = True
            MessageGridView.Appearance.EvenRow.BackColor = Color.FromArgb(248, 250, 252)
            MessageGridView.Appearance.Row.Font = New Font("Segoe UI", 9.0F)
            MessageGridView.Appearance.HeaderPanel.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
            MessageGridView.PopulateColumns()
            ConfigureColumns()
            AddHandler MessageGridView.RowStyle, AddressOf MessageGridView_RowStyle

            Controls.Add(MessageGrid)
            Controls.Add(commandPanel)
        End Sub

        Private Sub ConfigureColumns()
            HideColumn("EventID")
            HideColumn("Severity")
            Dim timeColumn = MessageGridView.Columns.ColumnByFieldName("TimeStamp")
            If timeColumn IsNot Nothing Then
                timeColumn.Caption = "Time"
                timeColumn.DisplayFormat.FormatType = FormatType.DateTime
                timeColumn.DisplayFormat.FormatString = "g"
                timeColumn.Width = 125
                timeColumn.VisibleIndex = 0
            End If
            SetColumn("SeverityText", "Type", 75, 1)
            SetColumn("Message", "Message", 300, 2)
            SetColumn("Source", "Source", 120, 3)
            SetColumn("Location", "Location", 100, 4)
            SetColumn("UserName", "User", 90, 5)
        End Sub

        Private Sub HideColumn(ByVal fieldName As String)
            Dim column = MessageGridView.Columns.ColumnByFieldName(fieldName)
            If column IsNot Nothing Then column.Visible = False
        End Sub

        Private Sub SetColumn(ByVal fieldName As String,
                              ByVal caption As String,
                              ByVal width As Integer,
                              ByVal visibleIndex As Integer)
            Dim column = MessageGridView.Columns.ColumnByFieldName(fieldName)
            If column Is Nothing Then Return
            column.Caption = caption
            column.Width = width
            column.VisibleIndex = visibleIndex
        End Sub

        Private Function CreateButton(ByVal caption As String,
                                      ByVal clickHandler As EventHandler) As SimpleButton
            Dim button As New SimpleButton With {
                .Text = caption,
                .AutoSize = True,
                .Margin = New Padding(0, 0, 5, 0)}
            AddHandler button.Click, clickHandler
            Return button
        End Function

        Private Sub MessagesChanged(ByVal sender As Object, ByVal e As EventArgs)
            RefreshMessages()
        End Sub

        Private Sub MessageGridView_RowStyle(ByVal sender As Object,
                                             ByVal e As RowStyleEventArgs)
            If e.RowHandle < 0 Then Return
            Dim item As SystemMessageRecord = TryCast(MessageGridView.GetRow(e.RowHandle), SystemMessageRecord)
            If item Is Nothing Then Return
            Select Case item.Severity
                Case SystemMessageSeverity.Success
                    e.Appearance.ForeColor = Color.FromArgb(20, 122, 61)
                Case SystemMessageSeverity.Warning
                    e.Appearance.ForeColor = Color.FromArgb(154, 103, 0)
                Case SystemMessageSeverity.Error
                    e.Appearance.ForeColor = Color.FromArgb(198, 40, 40)
                    e.Appearance.FontStyleDelta = FontStyle.Bold
            End Select
        End Sub

        Private Sub RefreshButtonClick(ByVal sender As Object, ByVal e As EventArgs)
            RefreshMessages()
        End Sub

        Private Sub CopyButtonClick(ByVal sender As Object, ByVal e As EventArgs)
            Try
                Clipboard.SetText(MessageManager.CreateTextExport())
            Catch ex As Exception
                XtraMessageBox.Show(Me, ex.Message, "Copy system messages", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub

        Private Sub SaveButtonClick(ByVal sender As Object, ByVal e As EventArgs)
            Using dialog As New SaveFileDialog With {
                .Title = "Save system messages",
                .Filter = "HTML file (*.html)|*.html|Text file (*.txt)|*.txt",
                .DefaultExt = "html",
                .AddExtension = True,
                .FileName = "Summit system messages " & Now().ToString("yyyy-MM-dd HHmm")}
                If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
                If String.Equals(Path.GetExtension(dialog.FileName), ".txt", StringComparison.OrdinalIgnoreCase) Then
                    File.WriteAllText(dialog.FileName, MessageManager.CreateTextExport(), System.Text.Encoding.UTF8)
                Else
                    File.WriteAllText(dialog.FileName, MessageManager.CreateHtmlExport(), System.Text.Encoding.UTF8)
                End If
            End Using
        End Sub

        Private Sub EmailButtonClick(ByVal sender As Object, ByVal e As EventArgs)
            Dim exportDirectory As String = Path.Combine(Path.GetTempPath(), "Abovo Summit")
            Directory.CreateDirectory(exportDirectory)
            Dim attachmentPath As String = Path.Combine(exportDirectory,
                "System messages " & Now().ToString("yyyyMMdd HHmmss") & ".txt")
            File.WriteAllText(attachmentPath, MessageManager.CreateTextExport(), System.Text.Encoding.UTF8)

            Dim outlook As Object = Nothing
            Dim mail As Object = Nothing
            Try
                outlook = CreateObject("Outlook.Application")
                mail = outlook.CreateItem(0)
                mail.Subject = "Abovo Summit system messages"
                mail.Body = "The exported Abovo Summit system messages are attached."
                mail.Attachments.Add(attachmentPath)
                mail.Display()
            Catch
                Dim body As String = MessageManager.CreateTextExport()
                If body.Length > 1500 Then body = body.Substring(0, 1500) & Environment.NewLine & "[Export shortened for email.]"
                Dim mailUri As String = "mailto:?subject=" & System.Uri.EscapeDataString("Abovo Summit system messages") &
                                        "&body=" & System.Uri.EscapeDataString(body)
                Try
                    Process.Start(New ProcessStartInfo(mailUri) With {.UseShellExecute = True})
                Catch ex As Exception
                    XtraMessageBox.Show(Me,
                        "The default email application could not be opened." & Environment.NewLine &
                        "The complete message export was saved to:" & Environment.NewLine & attachmentPath &
                        Environment.NewLine & Environment.NewLine & ex.Message,
                        "Email system messages", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Try
            Finally
                If mail IsNot Nothing AndAlso Marshal.IsComObject(mail) Then Marshal.FinalReleaseComObject(mail)
                If outlook IsNot Nothing AndAlso Marshal.IsComObject(outlook) Then Marshal.FinalReleaseComObject(outlook)
            End Try
        End Sub
    End Class
End Namespace
