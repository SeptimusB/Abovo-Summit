Imports DevExpress.XtraEditors

Namespace Abovo

    Public Class ApplicationOptionsForm
        Inherits XtraForm

        Private ReadOnly ScaleTrack As TrackBarControl
        Private ReadOnly ValueLabel As LabelControl
        Private ReadOnly PreviewTitle As LabelControl
        Private ReadOnly PreviewText As LabelControl
        Private ReadOnly PreviewPanel As GroupControl
        Private PreviewHeadingFont As Font
        Private PreviewBodyFont As Font
        Private ReadOnly OkButton As SimpleButton
        Private ReadOnly CancelActionButton As SimpleButton
        Private ReadOnly ResetButton As SimpleButton
        Private ReadOnly BackupEnabled As CheckEdit
        Private ReadOnly BackupMinutes As SpinEdit
        Private ReadOnly BackupWhenIdle As CheckEdit
        Private ReadOnly BackupIdleMinutes As SpinEdit
        Private ReadOnly BackupAlways As CheckEdit
        Private ReadOnly ContinueBackupOnError As CheckEdit
        Private ReadOnly IntegrityEnabled As CheckEdit
        Private ReadOnly IntegrityMinutes As SpinEdit
        Private ReadOnly IntegrityPlan As ComboBoxEdit
        Private ReadOnly IntegrityPlans As New List(Of FileManager.ExcelModel)
        Private ReadOnly WatchEnabled As CheckEdit
        Private ReadOnly WatchMinutes As SpinEdit
        Private ReadOnly WatchIdle As SpinEdit
        Private ReadOnly WatchTiming As CheckEdit
        Private ReadOnly EngineChoice As ComboBoxEdit

        Public Sub New()
            Me.New(Nothing)
        End Sub

        Public Sub New(preferredModelID As Integer?)
            Text = "Abovo Summit options"
            StartPosition = FormStartPosition.CenterParent
            MinimizeBox = False
            MaximizeBox = False
            ShowIcon = False
            MinimumSize = New Size(620, 410)
            Size = New Size(760, 570)
            Dim shell As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2, .Padding = New Padding(8)}
            shell.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            shell.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Controls.Add(shell)
            Dim tabs As New DevExpress.XtraTab.XtraTabControl With {.Dock = DockStyle.Fill}
            Dim displayTab As New DevExpress.XtraTab.XtraTabPage With {.Text = "Display", .AutoScroll = True}
            Dim backupTab As New DevExpress.XtraTab.XtraTabPage With {.Text = "Recovery backup"}
            Dim integrityTab As New DevExpress.XtraTab.XtraTabPage With {.Text = "Integrity"}
            tabs.TabPages.AddRange({displayTab, backupTab, integrityTab})
            Dim engineTab As New DevExpress.XtraTab.XtraTabPage With {.Text = "Calculation engine", .AutoScroll = True}
            tabs.TabPages.Add(engineTab)
            Dim engineLayout As New TableLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True, .ColumnCount = 1, .Padding = New Padding(18)}
            engineTab.Controls.Add(engineLayout)
            engineLayout.Controls.Add(New LabelControl With {.Text = "Engine for newly opened business plans", .Dock = DockStyle.Top})
            EngineChoice = New ComboBoxEdit With {.Name = "CalculationEngineChoice", .Dock = DockStyle.Top}
            EngineChoice.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
            EngineChoice.Properties.Items.AddRange({"Prefer compatible Microsoft Excel", "DevExpress — no Excel required"})
            EngineChoice.SelectedIndex = If(New WorkbookEngines.WorkbookEngineSettings().ReadOptions().Preference = WorkbookEngines.WorkbookEnginePreference.DevExpressOnly, 1, 0)
            engineLayout.Controls.Add(EngineChoice)
            engineLayout.Controls.Add(New LabelControl With {.Dock = DockStyle.Top, .AutoSizeMode = LabelAutoSizeMode.Vertical, .Padding = New Padding(0, 10, 0, 10),
                .Text = "Applies when a business plan is next opened. Open plans keep their current engine." & Environment.NewLine & Environment.NewLine &
                    "Excel must be installed, compatible and permitted to run the model's VBA functions under your existing Excel security settings. If it is unavailable, Summit uses DevExpress. No Excel security settings are changed." & Environment.NewLine & Environment.NewLine &
                    "This integration stage supports value edits, calculations, Undo/Redo, same-format saves and recovery backups. Use DevExpress for structural changes, schedules, imports, Stress Test or direct spreadsheet editing. Recovered files use DevExpress for format conversion. Excel Save As requires a new filename in the same format."})
            Dim engineStatus As New LabelControl With {.Name = "CalculationEngineStatus", .Dock = DockStyle.Top, .AutoSizeMode = LabelAutoSizeMode.Vertical,
                .Text = "Open plans:" & Environment.NewLine}
            Dim openPlans = If(FileManager.ExcelModels, New FileManager.ExcelModel() {}).Where(Function(model) model IsNot Nothing AndAlso Not model.IsClosing AndAlso model.WB IsNot Nothing).ToArray()
            engineStatus.Text &= If(openPlans.Length = 0, "None", String.Join(Environment.NewLine,
                openPlans.Select(Function(model) IO.Path.GetFileName(model.FileName) & ": " & WorkbookEngines.ModelEngineSelection.Status(model))))
            engineLayout.Controls.Add(engineStatus)
            Dim watchTab As New DevExpress.XtraTab.XtraTabPage With {.Text = "Check Sheet trial", .AutoScroll = True}
            tabs.TabPages.Add(watchTab)
            CheckSheetWatch.Initialise()
            Dim watchLayout As New FlowLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True, .FlowDirection = FlowDirection.TopDown, .WrapContents = False, .Padding = New Padding(18)}
            watchTab.Controls.Add(watchLayout)
            WatchEnabled = New CheckEdit With {.Text = "Automatically check the Check Sheet", .AutoSizeInLayoutControl = True, .Width = 420, .Checked = CheckSheetWatch.Enabled}
            watchLayout.Controls.Add(WatchEnabled)
            Dim watchInterval As New FlowLayoutPanel With {.AutoSize = True, .WrapContents = False}
            watchInterval.Controls.Add(New LabelControl With {.Text = "Every", .Padding = New Padding(0, 5, 4, 0)})
            WatchMinutes = New SpinEdit With {.Width = 70, .EditValue = CheckSheetWatch.Minutes}
            WatchIdle = New SpinEdit With {.Width = 70, .EditValue = CheckSheetWatch.IdleMinutes}
            For Each spinner In {WatchMinutes, WatchIdle}
                spinner.Properties.IsFloatValue = False
                spinner.Properties.MinValue = 1 : spinner.Properties.MaxValue = 120
            Next
            watchInterval.Controls.Add(WatchMinutes)
            watchInterval.Controls.Add(New LabelControl With {.Text = "minutes, when idle for", .Padding = New Padding(4, 5, 4, 0)})
            watchInterval.Controls.Add(WatchIdle)
            watchInterval.Controls.Add(New LabelControl With {.Text = "minutes", .Padding = New Padding(4, 5, 0, 0)})
            watchLayout.Controls.Add(watchInterval)
            WatchTiming = New CheckEdit With {.Text = "Temporary timings for Check Sheet and edits", .Width = 480, .Checked = CheckSheetWatch.Benchmark}
            watchLayout.Controls.Add(WatchTiming)
            watchLayout.Controls.Add(New LabelControl With {.AutoSizeMode = LabelAutoSizeMode.Vertical, .Width = 580,
                .Text = "Checks open plans after inputs change. Runs a full workbook calculation (CalculateFull), then updates the Check Sheet balance message. This may take several seconds. It does not rebuild dependencies or certify formula integrity." & Environment.NewLine & Environment.NewLine &
                    "Waits for edits, saves and other operations to finish. Once a calculation starts it finishes safely. The separate Integrity tab checks formulas and references." & Environment.NewLine & Environment.NewLine &
                    "Timings appear in debugger output. Turn off the trial timings after testing. They do not enable the other application traces."})
            AddHandler WatchEnabled.CheckedChanged, Sub()
                WatchMinutes.Enabled = WatchEnabled.Checked : WatchIdle.Enabled = WatchEnabled.Checked
            End Sub
            WatchMinutes.Enabled = WatchEnabled.Checked : WatchIdle.Enabled = WatchEnabled.Checked
            shell.Controls.Add(tabs, 0, 0)

            Dim Layout As New TableLayoutPanel With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .ColumnCount = 1,
                .RowCount = 5,
                .Padding = New Padding(18)
            }
            Layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
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

            PreviewPanel = New GroupControl With {
                .Name = "ScalePreview",
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
            PreviewPanel.Controls.Add(PreviewText)
            PreviewPanel.Controls.Add(PreviewTitle)
            Layout.Controls.Add(PreviewPanel, 0, 4)
            AddHandler PreviewPanel.SizeChanged, Sub() SizePreview()
            AddHandler PreviewTitle.SizeChanged, Sub() SizePreview()
            AddHandler PreviewText.SizeChanged, Sub() SizePreview()

            Dim RangeLabel As New LabelControl With {
                .Text = "75% – 200%",
                .AutoSizeMode = LabelAutoSizeMode.Default,
                .Padding = New Padding(0, 6, 12, 8)
            }
            RangeLabel.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            Dim scaleActions As New FlowLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True}
            scaleActions.Controls.Add(RangeLabel)
            Layout.Controls.Add(scaleActions, 0, 3)

            Dim Buttons As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .AutoSize = True,
                .WrapContents = False
            }

            CancelActionButton = New SimpleButton With {.Text = "Cancel", .DialogResult = DialogResult.Cancel}
            OkButton = New SimpleButton With {.Text = "OK"}
            ResetButton = New SimpleButton With {.Name = "ResetDisplayScale", .Text = "Reset to 100%", .AutoSize = True, .MinimumSize = New Size(110, 24)}
            AddHandler OkButton.Click, AddressOf OkButton_Click
            AddHandler ResetButton.Click, AddressOf ResetButton_Click

            Buttons.Controls.Add(CancelActionButton)
            Buttons.Controls.Add(OkButton)
            scaleActions.Controls.Add(ResetButton)
            shell.Controls.Add(Buttons, 0, 1)

            RecoveryBackupManager.Initialise()
            Dim backupLayout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .AutoScroll = True, .ColumnCount = 1, .RowCount = 5, .Padding = New Padding(18)}
            backupLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            backupLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            backupLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            backupLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            backupLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            backupLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            backupTab.Controls.Add(backupLayout)
            BackupEnabled = New CheckEdit With {.Text = "Enable backup save", .Dock = DockStyle.Top, .Checked = RecoveryBackupManager.Enabled}
            backupLayout.Controls.Add(BackupEnabled, 0, 0)
            Dim interval As New FlowLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True, .WrapContents = False, .Padding = New Padding(0, 12, 0, 12)}
            BackupAlways = New CheckEdit With {.Text = "Always every", .Width = 170, .Checked = RecoveryBackupManager.AlwaysEvery}
            interval.Controls.Add(BackupAlways)
            BackupMinutes = New SpinEdit With {.Width = 80, .Enabled = BackupEnabled.Checked}
            BackupMinutes.Properties.IsFloatValue = False
            BackupMinutes.Properties.MinValue = 1
            BackupMinutes.Properties.MaxValue = 120
            BackupMinutes.Properties.Increment = 1
            BackupMinutes.EditValue = RecoveryBackupManager.Minutes
            interval.Controls.Add(BackupMinutes)
            interval.Controls.Add(New LabelControl With {.Text = "minutes", .Padding = New Padding(6, 5, 0, 0)})
            backupLayout.Controls.Add(interval, 0, 2)
            Dim idleInterval As New FlowLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True, .WrapContents = False, .Padding = New Padding(0, 12, 0, 0)}
            BackupWhenIdle = New CheckEdit With {.Text = "When idle for", .Width = 170, .Checked = RecoveryBackupManager.WhenIdle}
            BackupIdleMinutes = New SpinEdit With {.Width = 80}
            BackupIdleMinutes.Properties.IsFloatValue = False
            BackupIdleMinutes.Properties.MinValue = 1
            BackupIdleMinutes.Properties.MaxValue = 120
            BackupIdleMinutes.Properties.Increment = 1
            BackupIdleMinutes.EditValue = RecoveryBackupManager.IdleMinutes
            idleInterval.Controls.Add(BackupWhenIdle)
            idleInterval.Controls.Add(BackupIdleMinutes)
            idleInterval.Controls.Add(New LabelControl With {.Text = "minutes", .Padding = New Padding(6, 5, 0, 0)})
            backupLayout.Controls.Add(idleInterval, 0, 1)
            Dim updateTiming As Action = Sub()
                                             BackupWhenIdle.Enabled = BackupEnabled.Checked
                                             BackupAlways.Enabled = BackupEnabled.Checked
                                             BackupIdleMinutes.Enabled = BackupEnabled.Checked AndAlso BackupWhenIdle.Checked
                                             BackupMinutes.Enabled = BackupEnabled.Checked AndAlso BackupAlways.Checked
                                         End Sub
            AddHandler BackupEnabled.CheckedChanged, Sub() updateTiming()
            AddHandler BackupWhenIdle.CheckedChanged, Sub() updateTiming()
            AddHandler BackupAlways.CheckedChanged, Sub() updateTiming()
            Dim sizeTiming As Action = Sub()
                                           'Measure in the native editor's DPI/font context, including its checkbox glyph.
                                           Dim width = Math.Max(BackupWhenIdle.CalcBestSize().Width, BackupAlways.CalcBestSize().Width)
                                           BackupWhenIdle.Width = width
                                           BackupAlways.Width = width
                                       End Sub
            AddHandler BackupWhenIdle.FontChanged, Sub() sizeTiming()
            AddHandler BackupAlways.FontChanged, Sub() sizeTiming()
            AddHandler Shown, Sub() sizeTiming()
            sizeTiming()
            updateTiming()
            ContinueBackupOnError = New CheckEdit With {.Text = "Continue autosave if Check Sheet error?", .Dock = DockStyle.Top, .Checked = RecoveryBackupManager.ContinueOnCheckSheetError}
            backupLayout.Controls.Add(ContinueBackupOnError, 0, 3)
            Dim guidance As New LabelControl With {.Dock = DockStyle.Top, .AutoSizeMode = LabelAutoSizeMode.Vertical, .Margin = New Padding(3, 6, 20, 3),
                .Text = "Recovery copies: ~<plan>_recovery.xlsm beside your plan. Unsaved user changes trigger backups; recalculation does not." & Environment.NewLine & Environment.NewLine &
                        "Check Sheet figures can be incomplete while you enter a plan. Backups continue by default; clear the option above if you prefer to pause them until the figures balance. Workbook overrides are respected. Formula and reference checks are separate." & Environment.NewLine & Environment.NewLine &
                        "When idle waits for no keyboard or mouse input. Always every is a maximum interval: it may interrupt work, but waits for active edits, calculations and dialogs to finish. Select either or both; each completed backup restarts the timers. Only new unsaved user changes are backed up." & Environment.NewLine & Environment.NewLine &
                        "A five-second notice lets you snooze until idle for one minute, even when Always every is due. Once writing starts it cannot stop midway. Recovery copies do not replace normal XLSB saves. A newer recovery is offered when opening the original, even if backups are disabled." & Environment.NewLine & Environment.NewLine &
                        "After recovery, results are refreshed. Use Save As with Excel Binary Workbook (.xlsb). Choose the suggested original filename or a new name; replacing the original requires confirmation."}
            backupLayout.Controls.Add(guidance, 0, 4)

            IdleIntegrityManager.Initialise()
            Dim integrityLayout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .AutoScroll = True, .ColumnCount = 1, .RowCount = 4, .Padding = New Padding(18)}
            integrityLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            integrityLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            integrityLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            integrityLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
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
            Dim runPanel As New TableLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True, .ColumnCount = 1, .RowCount = 4, .Padding = New Padding(0, 0, 0, 12)}
            For row = 0 To 3
                runPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Next
            runPanel.Controls.Add(New LabelControl With {.Text = "Business plan to check", .Dock = DockStyle.Top}, 0, 0)
            IntegrityPlan = New ComboBoxEdit With {.Dock = DockStyle.Top}
            IntegrityPlan.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
            IntegrityPlan.Properties.NullText = "Select the business plan to check"
            If FileManager.ExcelModels IsNot Nothing Then
                For Each model In FileManager.ExcelModels
                    If model Is Nothing OrElse model.IsClosing OrElse model.WB Is Nothing Then Continue For
                    IntegrityPlans.Add(model)
                    IntegrityPlan.Properties.Items.Add(IO.Path.GetFileName(model.FileName) & " — " & model.FileName)
                Next
            End If
            Dim targetLabel As New LabelControl With {.Name = "IntegrityTargetPath", .Dock = DockStyle.Top, .AutoSizeMode = LabelAutoSizeMode.Vertical, .Padding = New Padding(0, 6, 0, 6)}
            Dim runNow As New SimpleButton With {.Name = "RunIntegrityNow", .Text = "Run integrity check now", .AutoSize = True, .Anchor = AnchorStyles.Left, .Enabled = False}
            AddHandler runNow.Click, AddressOf RunIntegrityNow
            Dim updateTarget As Action = Sub()
                Dim selected = IntegrityPlan.SelectedIndex
                runNow.Enabled = selected >= 0 AndAlso selected < IntegrityPlans.Count
                targetLabel.Text = If(runNow.Enabled, "Checking: " & IntegrityPlans(selected).FileName,
                    "Select a plan above. No integrity check will run until a plan is selected.")
            End Sub
            AddHandler IntegrityPlan.SelectedIndexChanged, Sub() updateTarget()
            Dim preferredIndex = If(preferredModelID.HasValue,
                IntegrityPlans.FindIndex(Function(model) model.ModelID = preferredModelID.Value), -1)
            IntegrityPlan.SelectedIndex = If(preferredIndex >= 0, preferredIndex, If(IntegrityPlans.Count = 1, 0, -1))
            updateTarget()
            runPanel.Controls.Add(IntegrityPlan, 0, 1)
            runPanel.Controls.Add(targetLabel, 0, 2)
            runPanel.Controls.Add(runNow, 0, 3)
            integrityLayout.Controls.Add(runPanel, 0, 2)
            integrityLayout.Controls.Add(New LabelControl With {.Dock = DockStyle.Top, .AutoSizeMode = LabelAutoSizeMode.Vertical,
                .Text = "Waits for at least two minutes without keyboard or mouse input in your Windows session, with no pending editor, dialog, save or workbook operation." & Environment.NewLine & Environment.NewLine &
                        "Updates pending calculations and rebuilds dependencies when required, then checks the Check Sheet, named references, supported Transactional DB mirror sizes and cached cell errors in stages." & Environment.NewLine & Environment.NewLine &
                        "Calculation and the brief Check Sheet validation finish together, updating any Check Sheet warning before returning to input. The longer remaining checks pause on input and resume after two minutes idle; an edit restarts the pass." & Environment.NewLine & Environment.NewLine &
                        "Run now checks the selected plan after closing Options, even with scheduled checks disabled. It applies these options first. Results and sampled locations appear in System Messages. No repair, save or macros run; this is not financial sign-off."}, 0, 3)

            AcceptButton = OkButton
            CancelButton = CancelActionButton
            UpdatePreview()
        End Sub

        Private ReadOnly Property SelectedPercent As Integer
            Get
                Return Convert.ToInt32(ScaleTrack.EditValue)
            End Get
        End Property

        Private Sub RunIntegrityNow(sender As Object, e As EventArgs)
            If IntegrityPlan.SelectedIndex < 0 OrElse Not ApplyBackupSettings() Then Return
            Try
                IdleIntegrityManager.RequestNow(IntegrityPlans(IntegrityPlan.SelectedIndex))
                PresentationScaleManager.SetInterfaceScale(SelectedPercent)
                DialogResult = DialogResult.OK
                Close()
            Catch ex As Exception
                XtraMessageBox.Show(Me, "The integrity check could not start: " & ex.Message, "Integrity", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub

        Private Sub ScaleTrack_EditValueChanged(ByVal sender As Object, ByVal e As EventArgs)
            UpdatePreview()
        End Sub

        Private Sub ApplyPresentationScale()
            UpdatePreview()
        End Sub

        Private Sub UpdatePreview()
            If ValueLabel Is Nothing OrElse PreviewTitle Is Nothing OrElse PreviewText Is Nothing Then Return
            Dim PreviewScale As Single = CSng(SelectedPercent) / 100.0F
            ValueLabel.Text = SelectedPercent.ToString() & "%"
            Dim oldHeading = PreviewHeadingFont, oldBody = PreviewBodyFont
            PreviewHeadingFont = New Font(Font.FontFamily, 11.0F * PreviewScale, FontStyle.Bold)
            PreviewBodyFont = New Font(Font.FontFamily, 9.0F * PreviewScale, FontStyle.Regular)
            PreviewTitle.Appearance.Font = PreviewHeadingFont
            PreviewText.Appearance.Font = PreviewBodyFont
            SizePreview()
            If oldHeading IsNot Nothing Then oldHeading.Dispose()
            If oldBody IsNot Nothing Then oldBody.Dispose()
        End Sub

        Private Sub SizePreview()
            If PreviewPanel Is Nothing OrElse PreviewTitle Is Nothing OrElse PreviewText Is Nothing Then Return
            'A percentage row can shrink this to the caption at high DPI. Reserve
            'the measured content height; the Display tab scrolls in short windows.
            'LabelControl has already measured wrapped text in its own DPI context.
            'Screen-global TextRenderer measurements can double-scale this height.
            Dim height = Math.Max(CInt(90 * DeviceDpi / 96.0F), PreviewText.Bottom + PreviewPanel.Padding.Bottom + CInt(8 * DeviceDpi / 96.0F))
            If PreviewPanel.MinimumSize.Height <> height Then PreviewPanel.MinimumSize = New Size(0, height)
            ValueLabel.MinimumSize = New Size(TextRenderer.MeasureText("200%", ValueLabel.Font).Width + 12, 0)
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            MyBase.Dispose(disposing)
            If disposing Then
                If PreviewHeadingFont IsNot Nothing Then PreviewHeadingFont.Dispose()
                If PreviewBodyFont IsNot Nothing Then PreviewBodyFont.Dispose()
            End If
        End Sub

        Private Sub OkButton_Click(ByVal sender As Object, ByVal e As EventArgs)
            If Not ApplyBackupSettings() Then Return
            PresentationScaleManager.SetInterfaceScale(SelectedPercent)
            DialogResult = DialogResult.OK
            Close()
        End Sub

        Private Function ApplyBackupSettings() As Boolean
            Try
                Dim engines As New WorkbookEngines.WorkbookEngineSettings()
                engines.Preference = If(EngineChoice.SelectedIndex = 1, WorkbookEngines.WorkbookEnginePreference.DevExpressOnly, WorkbookEngines.WorkbookEnginePreference.Automatic).ToString()
                engines.Save()
                If BackupEnabled.Checked OrElse BackupWhenIdle.Checked OrElse BackupAlways.Checked Then
                    RecoveryBackupManager.ConfigureTiming(BackupWhenIdle.Checked, Convert.ToInt32(BackupIdleMinutes.EditValue),
                        BackupAlways.Checked, Convert.ToInt32(BackupMinutes.EditValue))
                End If
                RecoveryBackupManager.Configure(BackupEnabled.Checked, Convert.ToInt32(BackupMinutes.EditValue))
                RecoveryBackupManager.ConfigureCheckSheetPolicy(ContinueBackupOnError.Checked)
                IdleIntegrityManager.Configure(IntegrityEnabled.Checked, Convert.ToInt32(IntegrityMinutes.EditValue))
                CheckSheetWatch.Configure(WatchEnabled.Checked, Convert.ToInt32(WatchMinutes.EditValue), Convert.ToInt32(WatchIdle.EditValue), WatchTiming.Checked)
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
