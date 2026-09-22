Public Class WaitFormA
    Private LayoutReady As Boolean
    Private SizingNotice As Boolean

    Public Sub New()
        InitializeComponent()
        'Native ProgressPanel measures wrapped text; very long notices can scroll.
        AutoSize = False
        tableLayoutPanel1.AutoSize = False
        tableLayoutPanel1.AutoScroll = True
        tableLayoutPanel1.RowStyles(0).SizeType = System.Windows.Forms.SizeType.AutoSize
        progressPanel1.Dock = System.Windows.Forms.DockStyle.Top
        progressPanel1.AutoWidth = False
        progressPanel1.AutoHeight = True
        progressPanel1.AppearanceCaption.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        progressPanel1.AppearanceCaption.Options.UseTextOptions = True
        progressPanel1.AppearanceDescription.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        progressPanel1.AppearanceDescription.Options.UseTextOptions = True
        LayoutReady = True
        FitContents()
    End Sub

    Public Overrides Sub SetCaption(ByVal caption As String)
        MyBase.SetCaption(caption)
        Me.progressPanel1.Caption = If(caption, String.Empty)
        AccessibleName = progressPanel1.Caption
        FitContents()
    End Sub

    Public Overrides Sub SetDescription(ByVal description As String)
        MyBase.SetDescription(description)
        Me.progressPanel1.Description = If(description, String.Empty)
        AccessibleDescription = progressPanel1.Description
        FitContents()
    End Sub

    Protected Overrides Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        FitContents() 'The splash manager has now positioned us on the owner's monitor.
    End Sub

    Protected Overrides Sub OnDpiChanged(e As System.Windows.Forms.DpiChangedEventArgs)
        MyBase.OnDpiChanged(e)
        FitContents()
    End Sub

    Private Sub FitContents()
        If Not LayoutReady OrElse SizingNotice OrElse IsDisposed Then Return
        SizingNotice = True
        Try
            Dim workArea = System.Windows.Forms.Screen.FromControl(Me).WorkingArea
            Dim scale As Double = Math.Max(1.0, DeviceDpi / 96.0)
            Dim gap = CInt(Math.Ceiling(18 * scale))
            Dim borderWidth = Width - ClientSize.Width
            Dim borderHeight = Height - ClientSize.Height
            Dim maxWidth = Math.Max(1, Math.Min(CInt(680 * scale), workArea.Width - 2 * gap - borderWidth))
            Dim maxHeight = Math.Max(1, Math.Min(CInt(660 * scale), workArea.Height - 2 * gap - borderHeight))
            Dim flags = System.Windows.Forms.TextFormatFlags.NoPrefix Or System.Windows.Forms.TextFormatFlags.NoPadding
            Dim captionWidth = System.Windows.Forms.TextRenderer.MeasureText(progressPanel1.Caption, progressPanel1.AppearanceCaption.Font,
                                                        New System.Drawing.Size(Integer.MaxValue, Integer.MaxValue), flags).Width
            Dim descriptionWidth = System.Windows.Forms.TextRenderer.MeasureText(progressPanel1.Description, progressPanel1.AppearanceDescription.Font,
                                                            New System.Drawing.Size(Integer.MaxValue, Integer.MaxValue), flags).Width
            Dim desiredWidth = Math.Min(maxWidth, Math.Max(CInt(340 * scale), Math.Max(captionWidth, descriptionWidth) + CInt(100 * scale)))
            Dim centre = New System.Drawing.Point(Left + Width \ 2, Top + Height \ 2)
            tableLayoutPanel1.Padding = New System.Windows.Forms.Padding(gap, gap, gap, gap)
            tableLayoutPanel1.AutoScrollPosition = System.Drawing.Point.Empty
            tableLayoutPanel1.AutoScroll = False
            ClientSize = New System.Drawing.Size(desiredWidth, ClientSize.Height)
            'Let docking settle before measuring. Removing a scrollbar gives text
            'more width, so settle that second layout too before final sizing.
            For pass = 0 To 2
                tableLayoutPanel1.PerformLayout()
                'Docking can change the available width without remeasuring the
                'ProgressPanel's auto-height, especially after overflow scrolling.
                progressPanel1.AutoHeight = False
                progressPanel1.AutoHeight = True
                Dim desiredHeight = progressPanel1.Height + tableLayoutPanel1.Padding.Vertical + progressPanel1.Margin.Vertical
                ClientSize = New System.Drawing.Size(desiredWidth, Math.Min(maxHeight, desiredHeight))
                tableLayoutPanel1.AutoScroll = desiredHeight > maxHeight
            Next
            'Keep the owner's initial centre as messages change, bounded to screen.
            If Visible Then
                Location = New System.Drawing.Point(Math.Max(workArea.Left, Math.Min(centre.X - Width \ 2, workArea.Right - Width)),
                                     Math.Max(workArea.Top, Math.Min(centre.Y - Height \ 2, workArea.Bottom - Height)))
            End If
        Finally
            SizingNotice = False
        End Try
    End Sub

    Public Overrides Sub ProcessCommand(ByVal cmd As System.Enum, ByVal arg As Object)
        MyBase.ProcessCommand(cmd, arg)
    End Sub

    Public Enum WaitFormCommand
        SomeCommandId
    End Enum
End Class
