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

        Public Sub New()
            Text = "Abovo Summit options"
            StartPosition = FormStartPosition.CenterParent
            MinimizeBox = False
            MaximizeBox = False
            ShowIcon = False
            MinimumSize = New Size(520, 315)
            Size = New Size(620, 370)

            Dim Layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 6,
                .Padding = New Padding(18)
            }
            Layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            Layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Controls.Add(Layout)

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
            ResetButton = New SimpleButton With {.Text = "Reset to 100%"}
            AddHandler OkButton.Click, AddressOf OkButton_Click
            AddHandler ApplyButton.Click, AddressOf ApplyButton_Click
            AddHandler ResetButton.Click, AddressOf ResetButton_Click

            Buttons.Controls.Add(CancelActionButton)
            Buttons.Controls.Add(OkButton)
            Buttons.Controls.Add(ApplyButton)
            Buttons.Controls.Add(ResetButton)
            Layout.Controls.Add(Buttons, 0, 5)

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
            PresentationScaleManager.SetInterfaceScale(SelectedPercent)
        End Sub

        Private Sub OkButton_Click(ByVal sender As Object, ByVal e As EventArgs)
            PresentationScaleManager.SetInterfaceScale(SelectedPercent)
            DialogResult = DialogResult.OK
            Close()
        End Sub

        Private Sub ResetButton_Click(ByVal sender As Object, ByVal e As EventArgs)
            ScaleTrack.EditValue = PresentationScaleManager.DefaultPercent
        End Sub
    End Class

End Namespace
