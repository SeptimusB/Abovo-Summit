Imports DevExpress.CodeParser
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraCharts.Design
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Namespace Abovo
    Public Class GeneralFunctions

        'Colours
        Public Shared AbovoBlue As Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))
        Public Shared AbovoBlueL1 As Color = System.Drawing.Color.FromArgb(CType(CType(108, Byte), Integer), CType(CType(159, Byte), Integer), CType(CType(248, Byte), Integer))
        Public Shared AbovoBlueL2 As Color = System.Drawing.Color.FromArgb(CType(CType(138, Byte), Integer), CType(CType(186, Byte), Integer), CType(CType(255, Byte), Integer))
        Public Shared AbovoBlueL3 As Color = System.Drawing.Color.FromArgb(CType(CType(157, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(255, Byte), Integer))
        Public Shared AbovoBlueL4 As Color = System.Drawing.Color.FromArgb(CType(CType(181, Byte), Integer), CType(CType(228, Byte), Integer), CType(CType(255, Byte), Integer))
        Public Shared AbovoBlueL5 As Color = System.Drawing.Color.FromArgb(CType(CType(177, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(255, Byte), Integer))
        Public Shared AbovoBlueL6 As Color = System.Drawing.Color.FromArgb(CType(CType(193, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(255, Byte), Integer))
        Public Shared AbovoBlueL7 As Color = System.Drawing.Color.FromArgb(CType(CType(199, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(255, Byte), Integer))
        Public Shared AbovoBlueL8 As Color = System.Drawing.Color.FromArgb(CType(CType(205, Byte), Integer), CType(CType(251, Byte), Integer), CType(CType(255, Byte), Integer))
        Public Shared AbovoBlueL9 As Color = System.Drawing.Color.FromArgb(CType(CType(208, Byte), Integer), CType(CType(254, Byte), Integer), CType(CType(255, Byte), Integer))
        Public Shared AbovoBlueL10 As Color = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Public Shared AbovoBlueL11 As Color = System.Drawing.Color.FromArgb(CType(CType(218, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Public Shared AbovoBlueL12 As Color = System.Drawing.Color.FromArgb(CType(CType(227, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Public Shared AbovoBlueL13 As Color = System.Drawing.Color.FromArgb(CType(CType(234, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Public Shared AbovoComboBGC As Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(170, Byte), Integer))

        Public Shared DefaultGridCellPadding As Integer = 10
        Public Shared DefaultGridCellPaddingObject As New Padding With {.Left = 20, .Right = 20, .Top = 10, .Bottom = 10}
        Public Shared DefaultTablePanelPadding As System.Windows.Forms.Padding = New Padding(20, 10, 20, 40)
        Public Shared IsDebugRun As Boolean = True

        Public Shared Function GetNumericalIntegerInput(ByRef Prompt As String, Title As String, Optional ByVal DefaultValue As Integer = 3, Optional ByVal minval As Integer = -1000, Optional ByVal maxval As Integer = 1000) As Integer

            Dim TextEd As New TextEdit
            TextEd.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric
            TextEd.Properties.MaskSettings.Set("mask", "n0")

            Dim args As New XtraInputBoxArgs() With {
                .Caption = Title,
                .Prompt = Prompt,
                .DefaultButtonIndex = 0,
                .Editor = TextEd,
                .DefaultResponse = DefaultValue
            }

            Dim result = XtraInputBox.Show(args)

            If result IsNot Nothing AndAlso Integer.TryParse(result.ToString(), Nothing) Then

                Dim inputValue As Integer = Convert.ToInt32(result)

                If inputValue < minval OrElse inputValue > maxval Then

                    XtraMessageBox.Show($"Please enter a value between {minval} and {maxval}.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return GetNumericalIntegerInput(Prompt, Title, DefaultValue, minval, maxval)

                End If

                Return inputValue

            Else

                Return 0

            End If

        End Function
        Public Shared Sub InstaMsg(msg As String)

            Dim MsgFrm As New InsantMsgForm(msg)
            MsgFrm.ShowDialog()
            MsgFrm = Nothing

        End Sub

        Public Shared Function TrimSpaces(InString As String) As String

            Dim TempString As String = InString.Trim

            If Len(TempString) = 0 Or TempString Is Nothing Then Return ""

            Dim Outstring As String = ""

            Dim LastChar As String, ThisChar As String

            LastChar = Left(TempString, 1)

            If Len(TempString) = 1 Then

                Return TempString

            Else

                Outstring += LastChar

                For x = 2 To Len(TempString)

                    ThisChar = Mid(TempString, x, 1)

                    If LastChar = " " And ThisChar = " " Then

                    Else

                        Outstring += ThisChar
                        LastChar = ThisChar

                    End If

                Next

                Return Outstring

            End If

        End Function

        Public Shared Function GetMaxListItemLength(List As List(Of String)) As Integer

            Dim MaxLength As Integer = 0
            For Each Item As String In List
                If Len(Item) > MaxLength Then MaxLength = Len(Item)
            Next
            Return MaxLength

        End Function
        Public Shared Function RemoveWhitespace(fullString As String) As String

            Return New String(fullString.Where(Function(x) Not Char.IsWhiteSpace(x)).ToArray())

        End Function
        Public Shared Function ToHex(col As System.Drawing.Color) As String
            If col = Color.Empty Then Return "#FFFFFF"
            Return "#" & col.R.ToString("X2") & col.G.ToString("X2") & col.B.ToString("X2")

        End Function

        Public Shared Sub RunSideBySide(ModelID As Integer, TargetInterface As String, Caller As Form, CallGoesRight As Boolean)

            Select Case TargetInterface

                Case "Assumptions"

                    If CallGoesRight Then

                        FileManager.ExcelModels(ModelID).WBInterface.ShowGroupInterface(ModelID, 0, "Left", "Assumptions", FileManager.ExcelModels(ModelID).InstanceInterface)

                        FileManager.ExcelModels(ModelID).WBInterface.SetFormToHalfScreen(Caller, "Right")

                    Else

                        FileManager.ExcelModels(ModelID).WBInterface.ShowGroupInterface(ModelID, 0, "Right", "Assumptions", FileManager.ExcelModels(ModelID).InstanceInterface)

                        FileManager.ExcelModels(ModelID).WBInterface.SetFormToHalfScreen(Caller, "Left")

                    End If

            End Select

        End Sub
        Public Shared Function DoesFileExist(FilePath As String) As Boolean
            If IO.File.Exists(FilePath) Then
                Return True
            Else
                Return False
            End If
        End Function



        Public Shared Function RenderRangeCells(Ranges As List(Of DevExpress.Spreadsheet.CellRange), Optional ByVal WidthsList As List(Of Integer) = Nothing, Optional ByVal HeightsList As List(Of Integer) = Nothing) As String

            Dim i As Integer, j As Integer
            Dim CellExamine As DevExpress.Spreadsheet.Cell


            Dim HeightS As String = "21"
            Dim TestString As String = ""
            Dim NumExamine As Double
            Dim MaxColCount As Integer = 0

            Dim Range As DevExpress.Spreadsheet.CellRange

            For Each Range In Ranges

                If Range.ColumnCount > MaxColCount Then MaxColCount = Range.ColumnCount

            Next

            Dim StrOutput As String = My.Resources.StringTemplates.HTMLFinanceTableHeader
            StrOutput += My.Resources.StringTemplates.HTMLFinanceTablePrecursor

            For Each Range In Ranges


                For i = 0 To Range.RowCount - 1

                    StrOutput += "<tr class=xl822235 height=" & HeightS & " style='height:13.45pt'>"

                    For j = 0 To MaxColCount - 1 'Range.ColumnCount - 1




                        If Range.ColumnCount - 1 < j Then
                            'Debug.Print("Skipping " & i & " " & j)
                            StrOutput += "<td height = " & HeightS & " Class=xl882235 style='font-size:10.0pt;'>@nbsp;</td>"
                            Continue For
                        End If

                        CellExamine = Range(i, j)

                        'Debug.Print(CellExamine.FillColor.ToString & " " & CellExamine.DisplayText)

                        StrOutput += "<td height = " & HeightS & " bgcolor='" & ToHex(CellExamine.FillColor) & "' Class=xl882235 style='font-size:10.0pt;"

                        If CellExamine.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center Then

                            StrOutput += "text-align:general;"

                        ElseIf CellExamine.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left OrElse CellExamine.Alignment.Horizontal = SpreadsheetHorizontalAlignment.General Then

                            StrOutput += "text-align:left;"

                        ElseIf CellExamine.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right Then

                            StrOutput += "text-align:right;"

                        End If

                        TestString = CellExamine.DisplayText

                        StrOutput += "color:" & ToHex(CellExamine.Font.Color) & ";"

                        If IsNumeric(TestString) Then

                            NumExamine = CDbl(TestString)

                            If NumExamine < 0 Then

                                StrOutput += "color:" & ToHex(CellExamine.Font.Color) & ";"

                            End If


                        End If



                        If CellExamine.Font.Bold = True Then

                            StrOutput += "font-weight:600;"

                        Else

                            StrOutput += "font-weight:300;"

                        End If

                        StrOutput += "background:" & ToHex(CellExamine.FillColor) & ";"

                        StrOutput += "Text-decoration: none;text-underline-style:none;text-line-through:none;
                                              Font-family: Arial, sans - serif;mso-background-source: auto;mso-pattern:red thin - diag - stripe'>"

                        If TestString = "" Then StrOutput += "&nbsp;"
                        StrOutput += CellExamine.DisplayText
                        StrOutput += "</td>"
                    Next

                    StrOutput += "</tr>"
                    HeightS = "21"

                Next

            Next

            StrOutput += "

                        </table>

                        </div>


                        <!----------------------------->
                        <!--END OF OUTPUT FROM ABOVO SYSTEM-->
                        <!----------------------------->
                        </body>

                        </html>"
            Return StrOutput

        End Function
        'Case "GoAssumpt"
        ' OpenAssumptionsInterface
        'Dim x As Integer = ApplicationConfiguration.ActiveModelID
        '        ExcelModels(x).WBInterface.ShowGroupInterface(0)

        '    Case "GoWorkings"
        'Dim x As Integer = ApplicationConfiguration.ActiveModelID
        '        ExcelModels(x).WBInterface.ShowGroupInterface(1)

        '    Case "GoOutputs"
        'Dim x As Integer = ApplicationConfiguration.ActiveModelID
        '        ExcelModels(x).WBInterface.ShowGroupInterface(2)
    End Class

End Namespace
