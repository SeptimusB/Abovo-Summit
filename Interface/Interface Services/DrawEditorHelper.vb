Imports DevExpress.Utils.Drawing
Imports DevExpress.XtraEditors.Drawing
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Repository
Imports Abovo.GeneralFunctions
Imports DevExpress.XtraEditors.ViewInfo
Imports DevExpress.XtraGrid.Views.Grid

Namespace Abovo



	Public NotInheritable Class DrawEditorHelper
		Private Sub New()
		End Sub
		Public Shared Sub DrawEdit(ByVal g As Graphics, ByVal edit As RepositoryItem, ByVal r As Rectangle, ByVal value As Object)

			Dim info As BaseEditViewInfo = edit.CreateViewInfo()
			'info.Appearance.Options.UseBackColor = True
			'edit.Appearance.Options.UseBackColor = True
			'info.Appearance.BackColor = AbovoComboBGC
			'edit.Appearance.BackColor = AbovoComboBGC
			'info.Appearance.ForeColor = Color.White
			'edit.Appearance.ForeColor = Color.White
			'info.DefaultAppearance = edit.Appearance
			'info.AppearanceDisabled = edit.Appearance

			info.EditValue = value

			info.Bounds = r
			info.CalcViewInfo(g)
			'Dim D As Rectangle = info.GetTextBounds
			'r.Y += (r.Height - D.Height - 4)
			'info.Bounds = r
			'info.CalcViewInfo(g)
			Dim args As New ControlGraphicsInfoArgs(info, New GraphicsCache(g), r)

			edit.CreatePainter().Draw(args)
			args.Cache.Dispose()

		End Sub

		Public Shared Function GetNaturalEditorHeight(ByVal item As RepositoryItem, ByVal value As Object) As Integer

			If item Is Nothing Then Return 22

			Dim editor As BaseEdit = Nothing

			Try

				editor = item.CreateEditor()
				editor.Properties.Assign(item)
				editor.EditValue = value
				editor.CreateControl()

				Dim bestSize As Size = editor.CalcBestSize()

				If bestSize.Height > 0 Then Return bestSize.Height

			Catch

				'Fall through to the font-based/default estimate below.

			Finally

				If editor IsNot Nothing Then editor.Dispose()

			End Try

			If item.Appearance IsNot Nothing AndAlso item.Appearance.Font IsNot Nothing Then
				Return item.Appearance.Font.Height + 8
			End If

			Return 22

		End Function

		Public Shared Function GetEditorBounds(ByVal totalBounds As Rectangle,
											  ByVal rightIndent As Integer,
											  Optional ByVal preferredHeight As Integer = -1) As Rectangle

			Dim leftIndent As Integer = 20
			Dim rightPadding As Integer = 1 + Math.Max(0, rightIndent)
			Dim verticalPadding As Integer = 2

			Dim editorHeight As Integer = preferredHeight

			If editorHeight <= 0 Then
				editorHeight = totalBounds.Height - (2 * verticalPadding)
			End If

			'Never allow the editor to exceed the available header height.
			editorHeight = Math.Min(editorHeight,
									Math.Max(1, totalBounds.Height - (2 * verticalPadding)))

			Dim editorWidth As Integer = totalBounds.Width - leftIndent - rightPadding
			If editorWidth < 1 Then editorWidth = 1

			'Anchor the editor at the bottom of the header. This makes a multi-line
			'header behave exactly like a single-line header: the editor retains its
			'natural one-line height and the extra header height remains above it.
			Dim editorTop As Integer = totalBounds.Bottom - verticalPadding - editorHeight

			Return New Rectangle(totalBounds.Left + leftIndent,
								 editorTop,
								 editorWidth,
								 editorHeight)

		End Function

		Public Shared Sub DrawColumnInplaceEditor(ByVal e As ColumnHeaderCustomDrawEventArgs,
														ByVal item As RepositoryItem,
														ByVal value As Object,
														ByVal rightIndent As Integer,
														Optional ByVal preferredHeight As Integer = -1)

			Dim targetRect As Rectangle = GetEditorBounds(e.Bounds, rightIndent, preferredHeight)
			DrawEdit(e.Graphics, item, targetRect, value)

		End Sub
	End Class


End Namespace
