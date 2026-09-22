Imports DevExpress.XtraEditors

Namespace Abovo

    Public Class ApplicationOptionsForm
        Inherits XtraForm

        Private ReadOnly ScaleTrack As TrackBarControl
        Private ReadOnly ValueLabel As LabelControl
        Private ReadOnly PreviewTitle As LabelControl
        Private ReadOnly PreviewText As LabelControl
        Private ReadOnly ApplyButton As SimpleButton
        Private ReadOnly OkButton As SimpleButton
        Private ReadOnly CancelActionButton As SimpleButton
        Private ReadOnly ResetButton As SimpleButton
        Private ReadOnly BackupEnabled As CheckEdit
        Private ReadOnly BackupMinutes As SpinEdit
        Private ReadOnly IntegrityEnabled As CheckEdit
        Private ReadOnly IntegrityMinutes As SpinEdit

        Public Sub New()
            Text = "Abovo Summit options"
            StartPosition = FormStartPosition.CenterParent
            MinimizeBox = False
            MaximizeBox = False
            ShowIcon = False
            MinimumSize = New Size(620, 410)
            Size = New Size(720, 460)
            Dim shell As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2, .Padding = New Padding(8)}
            shell.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            shell.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Controls.Add(shell)
            Dim tabs As New DevExpress.XtraTab.XtraTabControl With {.Dock = DockStyle.Fill}
            Dim displayTab As New DevExpress.XtraTab.XtraTabPage With {.Text = "Display"}
            Dim backupTab As New DevExpress.XtraTab.XtraTabPage With {.Text = "Recovery backup"}
            Dim integrityTab As New DevExpress.XtraTab.XtraTabPage With {.Text = "Integrity"}
            tabs.TabPages.AddRange({displayTab, backupTab, integrityTab})
            shell.Controls.Add(tabs, 0, 0)

            Dim Layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 5,
                .Padding = New Padding(18)
            }
            Layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            Layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            displayTab.Controls.Add(Layout)

            Dim Heading As New LabelControl With {
                .Text = "Interface scale",
                .AutoSizeMode = LabelAutoSizeMode.Vertical,
                .Dock = DockStyle.Fill
            }
            Heading.Appearance.Font = New Font(Heading.Font, FontStyle.Bold)
            Layout.Controls.Add(Heading, 0, 0)

            Dim Explanation As New LabelControl With {
                .Text = "Adjust fonts, editors and interface spacing. Windows display scaling remains active.",
                .AutoSizeMode = LabelAutoSizeMode.Vertical,
                .Dock = DockStyle.Fill,
                .Padding = New Padding(0, 4, 0, 8)
            }
            Layout.Controls.Add(Explanation, 0, 1)

            Dim SliderPanel As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .AutoSize = True
            }
            SliderPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            SliderPanel.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

            ScaleTrack = New TrackBarControl With {.Dock = DockStyle.Fill}
            ScaleTrack.Properties.Minimum = PresentationScaleManager.MinimumPercent
            ScaleTrack.Properties.Maximum = PresentationScaleManager.MaximumPercent
            ScaleTrack.Properties.SmallChange = 5
            ScaleTrack.Properties.LargeChange = 25
            ScaleTrack.Properties.TickFrequency = 25
            ScaleTrack.EditValue = PresentationScaleManager.InterfaceScalePercent
            AddHandler ScaleTrack.EditValueChanged, AddressOf ScaleTrack_EditValueChanged
            SliderPanel.Controls.Add(ScaleTrack, 0, 0)

            ValueLabel = New LabelControl With {
                .AutoSizeMode = LabelAutoSizeMode.None,
                .Width = 70,
                .Dock = DockStyle.Fill
            }
            ValueLabel.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            SliderPanel.Controls.Add(ValueLabel, 1, 0)
            Layout.Controls.Add(SliderPanel, 0, 2)

            Dim Preview As New GroupControl With {
                .Text = "Preview",
                .Dock = DockStyle.Fill,
                .Padding = New Padding(12)
            }
            PreviewTitle = New LabelControl With {
                .Text = "Funding assumptions",
                .Dock = DockStyle.Top,
                .AutoSizeMode = LabelAutoSizeMode.Vertical
            }
            PreviewText = New LabelControl With {
                .Text = "Year 13 - 2038/39" & Environment.NewLine & "Example editor and navigation text",
                .Dock = DockStyle.Top,
                .AutoSizeMode = LabelAutoSizeMode.Vertical,
                .Padding = New Padding(0, 10, 0, 0)
            }
            Preview.Controls.Add(PreviewText)
            Preview.Controls.Add(PreviewTitle)
            Layout.Controls.Add(Preview, 0, 3)

            Dim RangeLabel As New LabelControl With {
                .Text = "75%     100%     125%     150%     175%     200%",
                .Dock = DockStyle.Fill,
                .Padding = New Padding(0, 6, 0, 8)
            }
            RangeLabel.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            Layout.Controls.Add(RangeLabel, 0, 4)

            Dim Buttons As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .AutoSize = True,
                .WrapContents = False
            }

            CancelActionButton = New SimpleButton With {.Text = "Cancel", .DialogResult = DialogResult.Cancel}
            OkButton = New SimpleButton With {.Text = "OK"}
            ApplyButton = New SimpleButton With {.Text = "Apply"}
            ResetButton = New SimpleButton With {.Text = "Reset to 100%", .AutoSize = True, .MinimumSize = New Size(110, 24)}
            AddHandler OkButton.Click, AddressOf OkButton_Click
            AddHandler ApplyButton.Click, AddressOf ApplyButton_Click
            AddHandler ResetButton.Click, AddressOf ResetButton_Click

            Buttons.Controls.Add(CancelActionButton)
            Buttons.Controls.Add(OkButton)
            Buttons.Controls.Add(ApplyButton)
            Buttons.Controls.Add(ResetButton)
            shell.Controls.Add(Buttons, 0, 1)

            RecoveryBackupManager.Initialise()
            Dim backupLayout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .AutoScroll = True, .ColumnCount = 1, .RowCount = 4, .Padding = New Padding(18)}
            backupLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            backupLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            backupLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            backupLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            backupTab.Controls.Add(backupLayout)
            BackupEnabled = New CheckEdit With {.Text = "Enable backup save", .Dock = DockStyle.Top, .Checked = RecoveryBackupManager.Enabled}
            backupLayout.Controls.Add(BackupEnabled, 0, 0)
            Dim interval As New FlowLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True, .WrapContents = False, .Padding = New Padding(0, 12, 0, 12)}
            interval.Controls.Add(New LabelControl With {.Text = "Every", .Padding = New Padding(0, 5, 6, 0)})
            BackupMinutes = New SpinEdit With {.Width = 80, .Enabled = BackupEnabled.Checked}
            BackupMinutes.Properties.IsFloatValue = False
            BackupMinutes.Properties.MinValue = 1
            BackupMinutes.Properties.MaxValue = 120
            BackupMinutes.Properties.Increment = 1
            BackupMinutes.EditValue = RecoveryBackupManager.Minutes
            interval.Controls.Add(BackupMinutes)
            interval.Controls.Add(New LabelControl With {.Text = "minutes", .Padding = New Padding(6, 5, 0, 0)})
            backupLayout.Controls.Add(interval, 0, 1)
            AddHandler BackupEnabled.CheckedChanged, Sub() BackupMinutes.Enabled = BackupEnabled.Checked
            Dim guidance As New LabelControl With {.Dock = DockStyle.Top, .AutoSizeMode = LabelAutoSizeMode.Vertical,
                .Text = "Recovery backups are separate XLSM files named <plan>_recovery.xlsm in the plan's folder. They contain committed inputs and preserve formulas and VBA; full calculated results are refreshed when recovered in Summit." & Environment.NewLine & Environment.NewLine &
                        "Normal Save still saves your XLSB business plan. A recovery backup does not clear unsaved changes and is not a replacement for Save." & Environment.NewLine & Environment.NewLine &
                        "Backups wait for a short idle period and for edits or long operations to finish. Newer recovery copies are offered when opening the original, even if backup saving is later disabled." & Environment.NewLine & Environment.NewLine &
                        "After recovery: use Save As, choose Excel Binary Workbook (.xlsb), and use the suggested original folder/filename or a new name. Replacing the original requires confirmation."}
            backupLayout.Controls.Add(guidance, 0, 2)

            IdleIntegrityManager.Initialise()
            Dim integrityLayout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .AutoScroll = True, .ColumnCount = 1, .RowCount = 3, .Padding = New Padding(18)}
            integrityLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            integrityLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            integrityLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            integrityTab.Controls.Add(integrityLayout)
            IntegrityEnabled = New CheckEdit With {.Text = "Check integrity every", .Dock = DockStyle.Top, .Checked = IdleIntegrityManager.Enabled}
            integrityLayout.Controls.Add(IntegrityEnabled, 0, 0)
            Dim integrityInterval As New FlowLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True, .WrapContents = False, .Padding = New Padding(0, 12, 0, 12)}
            IntegrityMinutes = New SpinEdit With {.Width = 80, .Enabled = IntegrityEnabled.Checked}
            IntegrityMinutes.Properties.IsFloatValue = False
            IntegrityMinutes.Properties.MinValue = 1
            IntegrityMinutes.Properties.MaxValue = 120
            IntegrityMinutes.EditValue = IdleIntegrityManager.Minutes
            integrityInterval.Controls.Add(IntegrityMinutes)
            integrityInterval.Controls.Add(New LabelControl With {.Text = "minutes if idle", .Padding = New Padding(6, 5, 0, 0)})
            integrityLayout.Controls.Add(integrityInterval, 0, 1)
            AddHandler IntegrityEnabled.CheckedChanged, Sub() IntegrityMinutes.Enabled = IntegrityEnabled.Checked
            integrityLayout.Controls.Add(New LabelControl With {.Dock = DockStyle.Top, .AutoSizeMode = LabelAutoSizeMode.Vertical,
                .Text = "Waits for at least two minutes without keyboard or mouse input in your Windows session, with no pending editor, dialog, save or workbook operation." & Environment.NewLine & Environment.NewLine &
                        "Updates pending calculations and rebuilds dependencies when required, then checks the Check Sheet, named references, supported Transactional DB mirror sizes and cached cell errors in stages." & Environment.NewLine & Environment.NewLine &
                        "Input pauses the next stage only. A calculation already running must finish safely and may temporarily delay interaction. Checks resume after two minutes idle; an edit restarts the pass." & Environment.NewLine & Environment.NewLine &
                        "Results and sampled problem locations appear in System Messages. Nothing is automatically repaired or saved, and external links/macros are not run. This is not a financial sign-off or a replacement for Excel/VBA testing."}, 0, 2)

            AcceptButton = OkButton
            CancelButton = CancelActionButton
            UpdatePreview()
        End Sub

        Private ReadOnly Property SelectedPercent As Integer
            Get
                Return Convert.ToInt32(ScaleTrack.EditValue)
            End Get
        End Property

        Private Sub ScaleTrack_EditValueChanged(ByVal sender As Object, ByVal e As EventArgs)
            UpdatePreview()
        End Sub

        Private Sub ApplyPresentationScale()
            UpdatePreview()
        End Sub

        Private Sub UpdatePreview()
            Dim PreviewScale As Single = CSng(SelectedPercent) / 100.0F
            ValueLabel.Text = SelectedPercent.ToString() & "%"
            PreviewTitle.Appearance.Font = New Font(Font.FontFamily, 11.0F * PreviewScale, FontStyle.Bold)
            PreviewText.Appearance.Font = New Font(Font.FontFamily, 9.0F * PreviewScale, FontStyle.Regular)
        End Sub

        Private Sub ApplyButton_Click(ByVal sender As Object, ByVal e As EventArgs)
            If Not ApplyBackupSettings() Then Return
            PresentationScaleManager.SetInterfaceScale(SelectedPercent)
        End Sub

        Private Sub OkButton_Click(ByVal sender As Object, ByVal e As EventArgs)
            If Not ApplyBackupSettings() Then Return
            PresentationScaleManager.SetInterfaceScale(SelectedPercent)
            DialogResult = DialogResult.OK
            Close()
        End Sub

        Private Function ApplyBackupSettings() As Boolean
            Try
                RecoveryBackupManager.Configure(BackupEnabled.Checked, Convert.ToInt32(BackupMinutes.EditValue))
                IdleIntegrityManager.Configure(IntegrityEnabled.Checked, Convert.ToInt32(IntegrityMinutes.EditValue))
                Return True
            Catch ex As Exception
                XtraMessageBox.Show(Me, "The options could not all be saved: " & ex.Message, "Options", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End Try
        End Function

        Private Sub ResetButton_Click(ByVal sender As Object, ByVal e As EventArgs)
            ScaleTrack.EditValue = PresentationScaleManager.DefaultPercent
        End Sub
    End Class

End Namespace
