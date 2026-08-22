Imports Abovo
Imports Abovo.GeneralFunctions
Imports Abovo.LogDebugDev
Imports DevExpress.XtraEditors
Imports System.IO
Imports DevExpress.XtraBars
Imports System.Runtime.CompilerServices
Imports DevExpress.XtraRichEdit.Model
Imports Abovo.DataObject

Namespace Abovo
    Public Class AbovoAppCls



        Public Shared WorkMode As String = "INTERFACE"
        Public Shared ReadOnly Property IsDev As Boolean = True
        Public Shared ReadOnly Property MaxGridHeight As Integer = CInt(Screen.PrimaryScreen.Bounds.Height * 0.7)
        Public Shared ReadOnly Property DecVersionNumber As Decimal = 7.25D
        Public Shared ReadOnly Property AppTitle As String = "abovo summit"
        Public Shared Property DefaultLrgFontSize As Integer = 12
        Public Shared Property DefaultMediumFontSize As Integer = 10
        Public Shared Property DefaultSmallFontSize As Integer = 9
        Private Shared Property SystemDefaultFont As Font = New Font("Segoe UI", 9.857143!, System.Drawing.FontStyle.Regular)

        Private Shared internalAppState As Integer

        Public Shared StandardFontSize As Single
        Public Shared Property AppState As Integer

            '0 initilising
            '1 menu mode
            '2 BP Mode
            '3 DSA Mode
            '4 MCM mode
            '5 Portal mode

            Get
                Return internalAppState
                Exit Property
            End Get

            Set(ByVal NewAppState As Integer)
                internalAppState = NewAppState
            End Set

        End Property

        Private Shared SystemLogText As String
        Public ReadOnly Property GetSystemLog As String
            Get
                Return SystemLogText
                Exit Property
            End Get
        End Property

        Private Shared internalCurrFile As String
        Public Shared Property CurrFile As String
            Get
                Return internalCurrFile
                Exit Property
            End Get
            Set(ByVal NewCurFile As String)
                internalCurrFile = NewCurFile
            End Set
        End Property
        Public Shared Sub Initialise()

            WindowsFormsSettings.LoadApplicationSettings()

            FontManager.Initialise()

            internalAppState = 0
            SystemLogText = ""
            SetDefaults()
            StandardFontSize = 10
            DefaultLrgFontSize = CInt(Screen.PrimaryScreen.Bounds.Width / 200)
            DefaultMediumFontSize = CInt(DefaultLrgFontSize * 0.75)
            DefaultSmallFontSize = CInt(DefaultLrgFontSize * 0.55)
            MasterChangeLog.Initialise()
            MasterChangeLog.AddChangeLogEvent(New ChangeLogEvent With {
                .Description = "Abovo Summit opened",
                .WSName = "System Message",
                .CellAddress = "",
                .OriginalValue = "",
                .ChangedValue = "",
                .TimeStamp = Now(),
                .UserName = Environment.UserName,
                .Status = 6
            })

        End Sub
        Public Shared Function GetFont(FontClass As String, Scale As Single, Optional ByVal Bold As Boolean = False, Optional ByVal Underline As Boolean = False, Optional ByVal Italic As Boolean = False) As Font

            Dim FontSize As Single

            Select Case FontClass

                Case "Large"

                    FontSize = DefaultLrgFontSize * Scale
                    If FontSize < 9 Then FontSize = 9
                    If FontSize > 16 Then FontSize = 16

                Case "Medium"

                    FontSize = DefaultMediumFontSize * Scale
                    If FontSize < 7 Then FontSize = 7
                    If FontSize > 14 Then FontSize = 14

                Case "Small"

                    FontSize = DefaultSmallFontSize * Scale
                    If FontSize < 6 Then FontSize = 6
                    If FontSize > 12 Then FontSize = 12

            End Select
            If FontSize < 6 Then FontSize = 6
            Dim ReturnFontStyle As FontStyle = FontStyle.Regular

            If (Bold) Then

                ReturnFontStyle = ReturnFontStyle Or FontStyle.Bold

            End If

            If (Underline) Then

                ReturnFontStyle = ReturnFontStyle Or FontStyle.Underline

            End If

            If (Italic) Then

                ReturnFontStyle = ReturnFontStyle Or FontStyle.Italic

            End If

            Dim ReturnFont As New Font(SystemDefaultFont.FontFamily, FontSize, ReturnFontStyle)

            Return ReturnFont

        End Function

        Public Sub New()


        End Sub
        Private Shared Sub SetDefaults()

            DevExpress.XtraEditors.WindowsFormsSettings.AllowRoundedWindowCorners = True
            'WindowsFormsSettings.DisableFormSkins()
            'DevExpress.XtraEditors.WindowsFormsSettings.SetAccentColor(Color.FromArgb(0, 91, 170))

        End Sub
        Public Shared Sub WriteLog(strEntry As String, Optional ByVal strSource As String = "")

            SystemLogText += strEntry & ", " & strSource & ", " & Now().ToString & vbLf

        End Sub
        Public Shared Function ConvertToStringNum(stPassed As String) As String

            Dim strTemp As String = ""
            Dim strTest As String
            For x = 0 To Len(stPassed) - 1
                strTest = stPassed.Substring(x, 1)

                If strTest = "." Or "-" Or IsNumeric(strTest) Then
                    Dim v As String = strTemp & strTest
                    strTemp = v
                End If
            Next x
            Return strTemp

        End Function

        Public Shared Function ConvertToNum(stPassed As String) As Double

            Dim strTemp As String = ""
            Dim strTest As String

            For x = 0 To Len(stPassed) - 1

                strTest = stPassed.Substring(x, 1)

                If strTest = "." Or "-" Or IsNumeric(strTest) Then

                    Dim v As String = strTemp & strTest
                    strTemp = v

                End If

            Next x

            Return CDbl(strTemp)

        End Function
        Public Shared Function ConvertToNumZeroed(stPassed As String) As Double

            If Len(stPassed) = 0 Then GoTo ConvertToNumZeroedReturn

            Dim strTemp As String = ""
            Dim strTest As String

            For x = 0 To Len(stPassed) - 1

                strTest = stPassed.Substring(x, 1)

                If strTest = "." Or strTest = "-" Or IsNumeric(strTest) Then

                    Dim v As String = strTemp & strTest
                    strTemp = v

                End If
            Next x

            If Len(strTemp) > 0 And Not strTemp = "-" Then

                Return CDbl(strTemp)

            Else
ConvertToNumZeroedReturn:
                Return 0

            End If

        End Function

        Public Class AbovoTransaction

            Public StrResponseMessage As String
            Public IntReturnCode As Integer
            Public BError As Boolean
            Public BSuccess As Boolean
            Public InternalMessage As String
            Public StrCallingProcedure As String
            Public EventCancelled As Boolean = False
            Public IntegerReturn As Integer
            Public StringReturn As String
            Public ObjectReturn As Object

            Sub New(Optional ByVal SetCallingProcedure As String = Nothing)

                StrCallingProcedure = SetCallingProcedure
                StrResponseMessage = "Initiated" '
                IntReturnCode = 0
                BError = False
                BSuccess = False
                InternalMessage = StrCallingProcedure & " transaction "

            End Sub

        End Class

    End Class
    Class AbovoMessageBox

        Private ReadOnly Message As String
        Private ReadOnly Header As String
        Private ReadOnly MessType As String
        Private ReadOnly CallingForm As Form
        Private ReadOnly CallingFormTitle As String

        Public Sub New(SetStrMessage As String, SetIntMessType As String, SetCallingForm As Form, Optional ByVal SetHeader As String = "Abovo")

            Message = SetStrMessage
            MessType = SetIntMessType
            Header = SetHeader
            CallingForm = SetCallingForm
            'CallingFormTitle = CallingForm.Text

        End Sub
        Function GetResponse() As Windows.Forms.DialogResult

            Dim Answer As Windows.Forms.DialogResult

            CallingForm.Text = "Awaiting user input - " & CallingFormTitle

            Using FrmMsg As New MessageForm(Message, MessType)

                FrmMsg.Text = Header
                Answer = FrmMsg.ShowDialog(Me)
                Return Answer
                CallingForm.Text = CallingFormTitle

            End Using

        End Function

    End Class

End Namespace

