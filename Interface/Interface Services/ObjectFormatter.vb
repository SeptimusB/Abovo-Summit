Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar
Imports Abovo.DataManager
Imports Abovo.DataObject
Imports Abovo.FontManager
Imports Abovo.GeneralFunctions
Imports Abovo.LogDebugDev
Imports DevExpress.LookAndFeel
Imports DevExpress.Skins
Imports DevExpress.Utils
Imports DevExpress.XtraBars.Navigation
Imports DevExpress.XtraEditors
Imports DevExpress.XtraExport.Helpers
Imports DevExpress.XtraVerticalGrid
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.BandedGrid
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraSpreadsheet.PrintLayoutEngine

Namespace Abovo


    Public Class ObjectFormatter

        Public DefaultLargeFont As Font = New System.Drawing.Font("Segoe UI Variable Display", 16.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Public DefaultFormatterFont As Font = New System.Drawing.Font("Segoe UI Variable Display", 14.85714!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Public DefaultHeaderFont As Font = New System.Drawing.Font("Segoe UI Variable Display", 14.85714!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

        Public Sub FormatGridControl(ObjGrid As DevExpress.XtraGrid.GridControl)

            ObjGrid.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
            ObjGrid.LookAndFeel.UseDefaultLookAndFeel = False
            ObjGrid.BackColor = Color.White

        End Sub

        Sub FormatTablePanel(ObjTablePanel As DevExpress.Utils.Layout.TablePanel)

            With ObjTablePanel

                .Appearance.BackColor = Color.White
                .Appearance.Options.UseBackColor = True

            End With
        End Sub

        Public Sub FormatVertGrid(ObjVGrid As VGridControl, Optional ByVal SetFontSize As String = "Normal")

            Dim FontToUse As Font
            Dim currentSize As Single



            Select Case SetFontSize

                Case "Massive"
                    currentSize = DefaultFont.SizeInPoints
                    currentSize += 6
                    FontToUse = New Font(DefaultFont.Name, currentSize, DefaultFont.Style)
                Case "Larger"
                    currentSize = DefaultFont.SizeInPoints
                    currentSize += 4
                    FontToUse = New Font(DefaultFont.Name, currentSize, DefaultFont.Style)
                Case "Larger"
                    currentSize = DefaultFont.SizeInPoints
                    currentSize += 2
                    FontToUse = New Font(DefaultFont.Name, currentSize, DefaultFont.Style)
                Case "Smaller"
                    currentSize = DefaultFont.SizeInPoints
                    currentSize -= 2
                    FontToUse = New Font(DefaultFont.Name, currentSize, DefaultFont.Style)
                Case Else
                    FontToUse = DefaultFont
            End Select

            With ObjVGrid

                .LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
                .LookAndFeel.UseDefaultLookAndFeel = False
                .BackColor = Color.White

                .Appearance.Category.BackColor = Color.WhiteSmoke

                .Appearance.Category.Options.UseBackColor = True
                .Appearance.Category.ForeColor = AbovoBlue
                .Appearance.Category.Options.UseForeColor = True



                .Appearance.Category.Font = FontToUse
                .Appearance.Category.FontStyleDelta = System.Drawing.FontStyle.Bold

                .Appearance.Category.Options.UseFont = True



                .Appearance.Caption.Font = FontToUse

                'VGrid field captions are drawn in the row-header panel rather
                'than in the value-cell CustomDrawRowValueCell event.  Keep the
                'caption formatting here so every VGrid uses the same interface
                'visual language and the cell custom-draw handler can remain
                'concerned with spreadsheet/value formatting.
                .Appearance.RowHeaderPanel.Font = FontToUse
                .Appearance.RowHeaderPanel.FontStyleDelta = System.Drawing.FontStyle.Bold
                .Appearance.RowHeaderPanel.Options.UseFont = True
                .Appearance.RowHeaderPanel.BackColor = System.Drawing.Color.White
                .Appearance.RowHeaderPanel.Options.UseBackColor = True
                .Appearance.RowHeaderPanel.ForeColor = AbovoBlue
                .Appearance.RowHeaderPanel.Options.UseForeColor = True
                .Appearance.RowHeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
                .Appearance.RowHeaderPanel.Options.UseTextOptions = True

                .Appearance.FocusedCell.BackColor = Color.WhiteSmoke
                .Appearance.Caption.BackColor = System.Drawing.Color.White
                .Appearance.Caption.Options.UseBackColor = True
                .Appearance.Caption.Font = FontToUse
                .Appearance.Caption.FontStyleDelta = System.Drawing.FontStyle.Bold

                .OptionsBehavior.ResizeRowHeaders = True
                .OptionsBehavior.AutoFocusNewRecord = True



            End With

        End Sub
        Public Sub FormatGridView(ObjGridView As GridView, ParentGrid As DevExpress.XtraGrid.GridControl, Optional ByVal SetFontSize As String = "Normal")


            Dim DoBandedExtras As Boolean

            Dim TryCheck As BandedGridView = TryCast(ObjGridView, BandedGridView)

            If TryCheck Is Nothing Then

                DoBandedExtras = False

            Else

                DoBandedExtras = True

            End If

            Dim FontToUse As Font
            Dim currentSize As Single

            Select Case SetFontSize

                Case "Massive"
                    currentSize = DefaultFont.SizeInPoints
                    currentSize += 6
                    FontToUse = New Font(DefaultFont.Name, currentSize, DefaultFont.Style)
                Case "Larger"
                    currentSize = DefaultFont.SizeInPoints
                    currentSize += 4
                    FontToUse = New Font(DefaultFont.Name, currentSize, DefaultFont.Style)
                Case "Larger"
                    currentSize = DefaultFont.SizeInPoints
                    currentSize += 2
                    FontToUse = New Font(DefaultFont.Name, currentSize, DefaultFont.Style)
                Case "Smaller"
                    currentSize = DefaultFont.SizeInPoints
                    currentSize -= 2
                    FontToUse = New Font(DefaultFont.Name, currentSize, DefaultFont.Style)
                Case Else
                    FontToUse = DefaultFont
            End Select


            With ObjGridView

                'Global settings

                .OptionsBehavior.AllowPartialGroups = True
                .OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never
                .OptionsView.AllowHtmlDrawGroups = True
                .OptionsView.AllowHtmlDrawHeaders = True
                .OptionsView.AllowHtmlDrawDetailTabs = True
                .OptionsView.ColumnAutoWidth = False
                '.OptionsView.ShowButtonMode = DevExpress.XtraGrid.Views.Base.ShowButtonModeEnum.ShowForFocusedRow
                .OptionsView.ShowGroupPanel = False
                .OptionsView.ShowIndicator = False
                .OptionsView.RowAutoHeight = True
                .OptionsSelection.MultiSelect = True
                .OptionsCustomization.AllowColumnMoving = False
                .OptionsCustomization.AllowFilter = False
                .OptionsCustomization.AllowGroup = False
                .OptionsCustomization.AllowRowSizing = True
                .OptionsCustomization.AllowSort = False
                .OptionsMenu.EnableColumnMenu = False
                .OptionsSelection.EnableAppearanceFocusedCell = False
                .OptionsSelection.EnableAppearanceHotTrackedRow = False
                .OptionsSelection.EnableAppearanceFocusedRow = False
                .OptionsClipboard.CopyColumnHeaders = DefaultBoolean.False
                .OptionsClipboard.PasteMode = DevExpress.Export.PasteMode.Update
                .FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus

                'Lines
                .FixedLineWidth = 1


                .Appearance.HorzLine.BorderColor = System.Drawing.Color.White
                .Appearance.HorzLine.ForeColor = System.Drawing.Color.White
                .Appearance.HorzLine.Options.UseBackColor = True
                .Appearance.HorzLine.Options.UseBorderColor = True
                .Appearance.HorzLine.Options.UseForeColor = True
                .Appearance.VertLine.BackColor = System.Drawing.Color.White
                .Appearance.VertLine.BorderColor = System.Drawing.Color.White
                .Appearance.VertLine.ForeColor = System.Drawing.Color.White
                .Appearance.VertLine.Options.UseBackColor = True
                .Appearance.VertLine.Options.UseBorderColor = True
                .Appearance.VertLine.Options.UseForeColor = True
                .Appearance.FixedLine.BackColor = System.Drawing.Color.White
                .Appearance.FixedLine.BorderColor = System.Drawing.Color.White
                .Appearance.FixedLine.Options.UseBackColor = True
                .Appearance.FixedLine.Options.UseBorderColor = True

                .Appearance.GroupFooter.BackColor = System.Drawing.Color.White

                .Appearance.HeaderPanel.BackColor = System.Drawing.Color.White
                .Appearance.HeaderPanel.Options.UseBackColor = True
                .Appearance.HeaderPanel.Font = FontToUse
                .Appearance.HeaderPanel.FontStyleDelta = System.Drawing.FontStyle.Bold

                .Appearance.GroupFooter.BorderColor = System.Drawing.Color.White
                .Appearance.GroupFooter.Options.UseBackColor = True
                .Appearance.GroupFooter.Options.UseBorderColor = True
                .Appearance.GroupFooter.Options.UseBackColor = False

                'Empty elements
                '.Appearance.Empty.BackColor = System.Drawing.Color.White
                .Appearance.Empty.BackColor = System.Drawing.Color.White
                .Appearance.Empty.Options.UseBackColor = True
                .Appearance.Empty.BorderColor = System.Drawing.Color.White
                .Appearance.Empty.Options.UseBorderColor = True

                'Grid Cells
                .Appearance.Row.Font = FontToUse
                .Appearance.Row.Options.UseFont = True
                .Appearance.HorzLine.BackColor = System.Drawing.Color.WhiteSmoke
                .Appearance.FocusedCell.BackColor = System.Drawing.Color.WhiteSmoke
                .Appearance.FocusedCell.Options.UseBackColor = True
                .Appearance.GroupRow.FontStyleDelta = System.Drawing.FontStyle.Bold
                .Appearance.GroupRow.Options.UseFont = True


                .Appearance.HotTrackedRow.Options.UseBackColor = True
                .Appearance.HotTrackedRow.Options.UseBorderColor = True
                .Appearance.HotTrackedRow.Options.UseForeColor = True

                .Appearance.HotTrackedRow.BackColor = AbovoBlueL3 'Color.FromArgb(70, AbovoBlue)
                .Appearance.HotTrackedRow.ForeColor = System.Drawing.Color.White
                .Appearance.HotTrackedRow.BorderColor = Color.WhiteSmoke

                .Appearance.HotTrackedRow.Options.UseFont = True
                .Appearance.HotTrackedRow.Font = DefaultFormatterFont

                'ColumnHeader

                '.Appearance.Cell.Font = FontToUse


                .Appearance.GroupFooter.Font = FontToUse
                .Appearance.GroupRow.FontStyleDelta = System.Drawing.FontStyle.Bold
                .Appearance.GroupFooter.Options.UseFont = True
                '.Appearance.grouprow.

                .Appearance.HeaderPanel.BackColor = System.Drawing.Color.White
                .Appearance.HeaderPanel.BorderColor = System.Drawing.Color.White
                .Appearance.HeaderPanel.ForeColor = AbovoBlue

                .Appearance.HeaderPanel.Options.UseBackColor = True
                .Appearance.HeaderPanel.Options.UseBorderColor = True
                .Appearance.HeaderPanel.Options.UseFont = True
                .Appearance.HeaderPanel.Options.UseForeColor = True
                .Appearance.HeaderPanel.Options.UseTextOptions = True

                .Appearance.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
                .Appearance.HeaderPanel.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Bottom
                .Appearance.HeaderPanel.ForeColor = AbovoBlue


                '.Appearance.FooterPanel.Font = FontToUse
                .Appearance.FooterPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Default
                .Appearance.HeaderPanel.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Bottom


                .Appearance.HideSelectionRow.BackColor = System.Drawing.Color.White
                .Appearance.HideSelectionRow.Options.UseBackColor = True

                .Appearance.ViewCaption.Options.UseFont = True
                .BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder

                '.RowHeight = CInt(0.5 * DefaultGridCellPadding) + CInt(FontToUse.GetHeight * 1.5) ' + 6)

            End With


            Dim element As SkinElement = SkinManager.GetSkinElement(SkinProductId.Grid, DevExpress.LookAndFeel.UserLookAndFeel.Default, "Header")
            element.Border.Thin.Bottom = -1
            element.Border.Thin.Right = -1
            element.Border.Thin.Left = -1
            element.Border.Thin.Top = -1

            element = SkinManager.GetSkinElement(SkinProductId.Grid, DevExpress.LookAndFeel.UserLookAndFeel.Default, "HeaderRight")
            element.Border.Thin.Bottom = -1
            element.Border.Thin.Right = -1
            element.Border.Thin.Left = -1
            element.Border.Thin.Top = -1

            element = SkinManager.GetSkinElement(SkinProductId.Grid, DevExpress.LookAndFeel.UserLookAndFeel.Default, "HeaderLeft")
            element.Border.Thin.Bottom = -1
            element.Border.Thin.Right = -1
            element.Border.Thin.Left = -1
            element.Border.Thin.Top = -1

            element = SkinManager.GetSkinElement(SkinProductId.Grid, DevExpress.LookAndFeel.UserLookAndFeel.Default, "SingleRowHeader")

            If element IsNot Nothing Then

                element.Border.Thin.Bottom = -1
                element.Border.Thin.Right = -1
                element.Border.Thin.Left = -1
                element.Border.Thin.Top = -1

            End If

            If DoBandedExtras Then


                Dim GVBandedGridView As BandedGridView = ObjGridView
                Dim g As Graphics = ParentGrid.CreateGraphics()
                Using g

                    For Each gridBand As DevExpress.XtraGrid.Views.BandedGrid.GridBand In GVBandedGridView.Bands

                        Dim CaptionSize As SizeF = g.MeasureString(gridBand.Caption, gridBand.AppearanceHeader.Font)

                        gridBand.MinWidth = CaptionSize.Width + 100

                    Next

                End Using

                With GVBandedGridView

                    .Appearance.BandPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
                    .Appearance.BandPanelBackground.BackColor = Color.White
                    .Appearance.BandPanelBackground.Options.UseBackColor = True
                    .Appearance.BandPanelBackground.ForeColor = AbovoBlue
                    .Appearance.BandPanelBackground.Options.UseForeColor = True

                    .Appearance.HeaderPanel.ForeColor = AbovoBlue
                    .Appearance.HeaderPanel.BackColor = Color.White
                    .Appearance.HeaderPanel.Options.UseForeColor = True
                    .Appearance.HeaderPanel.Options.UseBackColor = True
                    .Appearance.BandPanel.Font = FontToUse
                    .Appearance.BandPanel.FontStyleDelta = System.Drawing.FontStyle.Bold

                    .Appearance.BandPanel.ForeColor = AbovoBlue
                    .Appearance.BandPanel.BackColor = Color.White
                    .Appearance.BandPanel.Options.UseForeColor = True
                    .Appearance.BandPanel.Options.UseBackColor = True

                    .Appearance.HeaderPanelBackground.ForeColor = AbovoBlue
                    .Appearance.HeaderPanelBackground.BackColor = Color.White
                    .Appearance.HeaderPanelBackground.Options.UseForeColor = True
                    .Appearance.HeaderPanelBackground.Options.UseBackColor = True

                End With

            End If

            Try
                'ObjGridView.BestFitColumns()
            Catch ex As Exception

            End Try



        End Sub

        Public Sub ProcessGVColumWidths(ObjGridView As GridView, sender As Form)


            Dim ColTag As DataColumnTag
            Dim GC As GridControl = ObjGridView.GridControl

            Dim ViewTag As GridViewTag = ObjGridView.Tag

            ViewTag.HaveProcessedColumns = True

            If ViewTag.InManualReizeMode Then GoTo SkipCols

            'Dim CummWidth As Integer

            With ObjGridView

                For Each column As DevExpress.XtraGrid.Columns.GridColumn In .Columns

                    ColTag = column.Tag

                    If Not ColTag.ColumnWidthFixed Then

                        If ColTag.HasIncolumnButton Then

                            Dim g As Graphics = GC.CreateGraphics()
                            Using g

                                Dim CaptionSize As SizeF = g.MeasureString(column.Caption, ObjGridView.Appearance.HeaderPanel.Font)

                                column.MinWidth = CaptionSize.Width + 100
                                column.Width = column.MinWidth

                            End Using

                        End If

                        If ColTag.ColWidthMultiplier <> 1 Then

                            column.Width = CInt(column.Width * ColTag.ColWidthMultiplier)
                            column.MinWidth = column.Width

                        End If

                        ColTag.ColumnWidthFixed = True

                    End If

                Next

            End With

            'ObjGridView.LayoutChanged()

SkipCols:

            'Dim maxSize As New Size(sender.Width, sender.Height)
            'GC.Size = GC.CalcBestSize(maxSize, False)
            'GC.MainView?.LayoutChanged()

            'GC.Invalidate()

            'GC.Size = GC.CalcBestSize(maxSize, True)
            'GC.MainView?.LayoutChanged()
            'GC.BeginInvoke(New MethodInvoker(Sub()
            '                                     ObjGridView.InvalidateFooter()
            '                                     ObjGridView.Invalidate()
            '                                 End Sub))

        End Sub

        Public Sub FormatAccordianControl(ObjAC As AccordionControl)



        End Sub

        Public Sub FormatAccordianControlElement(ObjACE As AccordionControlElement)

            With ObjACE

                '.Appearance.BackColor = Color.White
                '.Appearance.Options.UseBackColor = True

            End With

        End Sub
        Public Sub FormatAccordianControlContainer(ObjACC As AccordionContentContainer)

            With ObjACC

                .Appearance.BackColor = Color.White
                .Appearance.Options.UseBackColor = True

            End With

        End Sub

    End Class

End Namespace
