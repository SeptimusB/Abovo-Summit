Imports DevExpress.XtraEditors

Namespace Abovo
    Public Class FontManager

        Public Shared DefaultFont As New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Public Shared DefaultFontBold As New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Public Shared DefaultFontSmaller As New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Public Shared DefaultFontLarger As New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Public Shared DefaultFontLargest As New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

        Public Shared Sub Initialise()

            Dim CurrdefFont As Font = WindowsFormsSettings.DefaultFont
            Dim NewFont As Font
            Dim ScreenWidth As Integer = Screen.PrimaryScreen.Bounds.Width

            If ScreenWidth <= 800 Then
                NewFont = New Font(CurrdefFont.FontFamily, 7.0!, CurrdefFont.Style, CurrdefFont.Unit, CurrdefFont.GdiCharSet, CurrdefFont.GdiVerticalFont)
            ElseIf ScreenWidth <= 1024 Then
                NewFont = New Font(CurrdefFont.FontFamily, 7.0!, CurrdefFont.Style, CurrdefFont.Unit, CurrdefFont.GdiCharSet, CurrdefFont.GdiVerticalFont)
            ElseIf ScreenWidth <= 1280 Then
                NewFont = New Font(CurrdefFont.FontFamily, 7.0!, CurrdefFont.Style, CurrdefFont.Unit, CurrdefFont.GdiCharSet, CurrdefFont.GdiVerticalFont)
            ElseIf ScreenWidth <= 1600 Then
                NewFont = New Font(CurrdefFont.FontFamily, 7.0!, CurrdefFont.Style, CurrdefFont.Unit, CurrdefFont.GdiCharSet, CurrdefFont.GdiVerticalFont)
            ElseIf ScreenWidth <= 1920 Then
                NewFont = New Font(CurrdefFont.FontFamily, 8.0!, CurrdefFont.Style, CurrdefFont.Unit, CurrdefFont.GdiCharSet, CurrdefFont.GdiVerticalFont)
            Else
                NewFont = New Font(CurrdefFont.FontFamily, 11.0!, CurrdefFont.Style, CurrdefFont.Unit, CurrdefFont.GdiCharSet, CurrdefFont.GdiVerticalFont)
            End If

            DefaultFont = NewFont
            WindowsFormsSettings.DefaultFont = NewFont

        End Sub
        Public Function GetFont(ObjectType As String) As System.Drawing.Font

            Dim RetFont As New System.Drawing.Font(DefaultFont, DefaultFont.Style)


            Select Case ObjectType

                Case "GridHeader"

                Case Else

            End Select

            Return RetFont

        End Function

    End Class

End Namespace
