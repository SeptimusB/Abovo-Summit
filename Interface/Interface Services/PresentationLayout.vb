Imports System.Drawing
Imports System.Windows.Forms
Imports DevExpress.Utils
Imports DevExpress.XtraBars.Docking2010
Imports DevExpress.XtraEditors

Namespace Abovo
    'Absolute measurements, never cumulative scaling. Windows supplies native DPI;
    'GetDisplayScale supplies the existing large-workspace and user preference scale.
    Friend Module PresentationLayout
        Friend Const BrowserOwnsScale As String = "SummitOwnsFontScale"
        Private ReadOnly WatchedPanels As New Runtime.CompilerServices.ConditionalWeakTable(Of WindowsUIButtonPanel, PanelImages)
        Private ReadOnly BitmapSources As New Runtime.CompilerServices.ConditionalWeakTable(Of WindowsUIButton, ButtonBitmap)
        Private ReadOnly CompactPanelsIcon As DevExpress.Utils.Svg.SvgImage = CreatePanelsIcon(True)
        Private ReadOnly RestorePanelsIcon As DevExpress.Utils.Svg.SvgImage = CreatePanelsIcon(False)

        Private Class ButtonBitmap
            Public Source As Image
            Public Scaled As Image
        End Class

        Private Class PanelImages
            Inherits List(Of WindowsUIButton)
            Public Backgrounds As ImageCollection
            Public BackgroundBitmaps As New List(Of Image)
        End Class

        Friend Sub FitCommandButton(button As SimpleButton)
            If button Is Nothing OrElse button.IsDisposed Then Return
            Dim best As Size = button.CalcBestSize()
            If button.Dock = DockStyle.None Then
                button.AutoSize = True
                button.Size = best
            End If
            'DIT VGrid command hosts have an explicit height, unlike auto-size rows.
            Dim host As Panel = TryCast(button.Parent, Panel)
            If host IsNot Nothing AndAlso TypeOf host.Tag Is DevExpress.XtraVerticalGrid.VGridControl Then
                host.Height = host.Controls.OfType(Of Control)().Max(Function(c) c.Bottom) + host.Padding.Bottom
            End If
        End Sub

        Friend Sub ApplyButtonPanel(panel As WindowsUIButtonPanel, reference As Control,
                                   Optional displayScale As Single = 0.0F)
            If panel Is Nothing OrElse panel.IsDisposed Then Return
            If displayScale <= 0 Then displayScale = AbovoAppCls.GetDisplayScale(reference)
            Dim font As Font
            Using baseFont As Font = AbovoAppCls.GetFont("Small", 1.0F)
                font = New Font(baseFont.FontFamily, baseFont.SizeInPoints * displayScale, baseFont.Style)
            End Using
            panel.Font = font
            For Each appearance In {panel.AppearanceButton.Normal, panel.AppearanceButton.Hovered,
                                    panel.AppearanceButton.Pressed}
                appearance.Font = font
                appearance.Options.UseFont = True
            Next
            Dim iconSize As Integer = CInt(Math.Round(28 * displayScale))
            For Each button As WindowsUIButton In panel.Buttons.OfType(Of WindowsUIButton)()
                If button.ImageOptions.HasSvgImage OrElse button.ImageOptions.ImageUri.HasSvgImage Then
                    button.ImageOptions.SvgImageSize = New Size(iconSize, iconSize)
                ElseIf button.ImageOptions.Image IsNot Nothing Then
                    Dim source As ButtonBitmap = BitmapSources.GetValue(button, Function(b) New ButtonBitmap With {.Source = b.ImageOptions.Image})
                    If Not Object.ReferenceEquals(button.ImageOptions.Image, source.Scaled) AndAlso
                       Not Object.ReferenceEquals(button.ImageOptions.Image, source.Source) Then
                        source.Source = button.ImageOptions.Image
                        If source.Scaled IsNot Nothing Then source.Scaled.Dispose()
                        source.Scaled = Nothing
                    End If
                    If source.Scaled Is Nothing OrElse source.Scaled.Width <> iconSize Then
                        Dim scaled As New Bitmap(iconSize, iconSize)
                        Using graphics = System.Drawing.Graphics.FromImage(scaled)
                            graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
                            Dim factor As Single = Math.Min(CSng(iconSize) / source.Source.Width, CSng(iconSize) / source.Source.Height)
                            Dim width As Single = source.Source.Width * factor
                            Dim height As Single = source.Source.Height * factor
                            graphics.DrawImage(source.Source, (iconSize - width) / 2, (iconSize - height) / 2, width, height)
                        End Using
                        Dim previous As Image = source.Scaled
                        source.Scaled = scaled
                        button.ImageOptions.Image = scaled
                        If previous IsNot Nothing Then previous.Dispose()
                    End If
                End If
            Next
            If panel.Orientation = Orientation.Horizontal Then panel.WrapButtons = True
            Dim marker As PanelImages = Nothing
            If Not WatchedPanels.TryGetValue(panel, marker) Then
                marker = New PanelImages
                WatchedPanels.Add(panel, marker)
                AddHandler panel.SizeChanged, Sub(sender, e) FitButtonPanelHeight(panel)
                AddHandler panel.Disposed,
                    Sub(sender, e)
                        'The native Buttons collection has already been released at Disposed.
                        For Each button As WindowsUIButton In marker
                            Dim bitmap As ButtonBitmap = Nothing
                            If BitmapSources.TryGetValue(button, bitmap) Then
                                If bitmap.Scaled IsNot Nothing Then bitmap.Scaled.Dispose()
                                BitmapSources.Remove(button)
                            End If
                        Next
                        If marker.Backgrounds IsNot Nothing Then marker.Backgrounds.Dispose()
                        For Each bitmap As Image In marker.BackgroundBitmaps
                            bitmap.Dispose()
                        Next
                    End Sub
            End If
            For Each button As WindowsUIButton In panel.Buttons.OfType(Of WindowsUIButton)()
                If Not marker.Contains(button) Then marker.Add(button)
            Next
            'WindowsUI's default circle is fixed at 42 logical pixels. Scale its
            'supported background collection together with the glyph, not the glyph alone.
            Dim diameter As Integer = CInt(Math.Round(42 * displayScale))
            If panel.ButtonBackgroundImages Is Nothing OrElse Object.ReferenceEquals(panel.ButtonBackgroundImages, marker.Backgrounds) Then
                If marker.Backgrounds Is Nothing OrElse marker.Backgrounds.ImageSize.Width <> diameter Then
                    Dim backgrounds As New ImageCollection With {.ImageSize = New Size(diameter, diameter)}
                    Dim owned As New List(Of Image)
                    For state As Integer = 0 To 2
                        'ImageCollection retains this image; it must stay alive until
                        'the collection is replaced or its owning panel is disposed.
                        Dim circle As New Bitmap(diameter, diameter)
                        Using graphics = System.Drawing.Graphics.FromImage(circle)
                            graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
                            Dim border As Single = CSng(Math.Max(1.5, diameter / 28.0))
                            Dim bounds As New RectangleF(border, border, diameter - 2 * border, diameter - 2 * border)
                            If state > 0 Then
                                Using fill As New SolidBrush(Color.FromArgb(If(state = 1, 25, 50), 0, 91, 170))
                                    graphics.FillEllipse(fill, bounds)
                                End Using
                            End If
                            Using pen As New Pen(Color.FromArgb(0, 91, 170), border)
                                graphics.DrawEllipse(pen, bounds)
                            End Using
                        End Using
                        backgrounds.AddImage(circle)
                        owned.Add(circle)
                    Next
                    Dim previous = marker.Backgrounds
                    marker.Backgrounds = backgrounds
                    panel.ButtonBackgroundImages = backgrounds
                    If previous IsNot Nothing Then previous.Dispose()
                    For Each bitmap As Image In marker.BackgroundBitmaps
                        bitmap.Dispose()
                    Next
                    marker.BackgroundBitmaps = owned
                End If
            End If
            FitButtonPanelHeight(panel)
        End Sub

        Private Sub FitButtonPanelHeight(panel As WindowsUIButtonPanel)
            If panel.IsDisposed OrElse panel.Orientation <> Orientation.Horizontal Then Return
            Dim height As Integer = ButtonPanelHeight(panel, panel.ClientSize.Width)
            Dim table = TryCast(panel.Parent, DevExpress.Utils.Layout.TablePanel)
            If table IsNot Nothing Then
                Dim row As Integer = table.GetRow(panel)
                If row >= 0 AndAlso row < table.Rows.Count Then
                    If table.Rows(row).Style = DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize Then
                        'A designer-scaled TablePanel applies its own factor to absolute
                        'row heights. An AutoSize row uses these actual control bounds.
                        panel.MinimumSize = New Size(0, height)
                        panel.Height = height
                    Else
                        table.Rows(row).Height = height + panel.Margin.Vertical
                    End If
                End If
            ElseIf panel.Dock = DockStyle.Top OrElse panel.Dock = DockStyle.Bottom Then
                If panel.Height <> height Then panel.Height = height
            End If
        End Sub

        Friend Function ButtonPanelHeight(panel As WindowsUIButtonPanel, availableWidth As Integer) As Integer
            Dim dpi As Single = CSng(panel.DeviceDpi) / 96.0F
            Dim rows As Integer = 1
            Dim rowWidth As Integer = 0
            Dim iconHeight As Integer = CInt(Math.Ceiling(50 * dpi))
            Dim backgrounds = TryCast(panel.ButtonBackgroundImages, ImageCollection)
            If backgrounds IsNot Nothing Then iconHeight = Math.Max(iconHeight, CInt((backgrounds.ImageSize.Height + 8) * dpi))
            Dim font As Font = panel.AppearanceButton.Normal.Font
            For Each button As WindowsUIButton In panel.Buttons.OfType(Of WindowsUIButton)().Where(Function(b) b.Visible)
                Dim icon As Size = button.ImageOptions.SvgImageSize
                If button.ImageOptions.Image IsNot Nothing Then icon = button.ImageOptions.Image.Size
                iconHeight = Math.Max(iconHeight, CInt(Math.Ceiling((icon.Height + 16) * dpi)))
                Dim caption As Size = If(button.UseCaption, TextRenderer.MeasureText(button.Caption, font), Size.Empty)
                Dim slot As Integer = Math.Max(iconHeight, caption.Width) + 2 * panel.ButtonInterval + CInt(8 * dpi)
                If panel.WrapButtons AndAlso rowWidth > 0 AndAlso rowWidth + slot > Math.Max(1, availableWidth) Then
                    rows += 1
                    rowWidth = 0
                End If
                rowWidth += slot
            Next
            Return rows * (iconHeight + font.Height) + CInt(8 * dpi)
        End Function

        Friend Function CreatePanelsButton() As WindowsUIButton
            Dim button As New WindowsUIButton With {.Tag = "TogglePanels"}
            UpdatePanelsButton(button, False)
            Return button
        End Function

        Friend Sub UpdatePanelsButton(button As WindowsUIButton, panelsHidden As Boolean)
            button.Caption = If(panelsHidden, "Restore", "Compact")
            button.UseCaption = False
            button.ToolTip = If(panelsHidden,
                "Restore — return navigators and sidebars to their previous state",
                "Compact — hide navigators and sidebars")
            'Shared vectors retain the native button's size and colour treatment.
            'Compact uses outward arrows; Restore shows the three-panel layout.
            button.ImageOptions.SvgImage = If(panelsHidden, RestorePanelsIcon, CompactPanelsIcon)
        End Sub

        Private Function CreatePanelsIcon(outwardArrows As Boolean) As DevExpress.Utils.Svg.SvgImage
            Dim detail As String = If(outwardArrows,
                "<path d='M12 14H6m3-3l-3 3 3 3M16 14H22m-3-3l3 3-3 3' stroke-linejoin='round'/>",
                "<path d='M9 5v18M19 5v18'/>")
            Dim svg As String = "<svg xmlns='http://www.w3.org/2000/svg' width='28' height='28' viewBox='0 0 28 28'><g fill='none' stroke='#005baa' stroke-width='2'><rect x='3' y='5' width='22' height='18' rx='1'/>" & detail & "</g></svg>"
            Using stream As New IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(svg))
                Return DevExpress.Utils.Svg.SvgImage.FromStream(stream)
            End Using
        End Function
    End Module
End Namespace
