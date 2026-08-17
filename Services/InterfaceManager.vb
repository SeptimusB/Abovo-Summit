Imports System.Drawing
Imports Abovo.PresentationManager
Imports Abovo.FileManager
Imports DevExpress.Utils.Drawing
Imports DevExpress.Utils.Extensions
Imports DevExpress.XtraSpreadsheet.API.Native.Implementation

Namespace Abovo
    Public Class InterfaceManager

        Public Shared GroupInterfaces() As GroupInterfaceObject
        Private Shared GroupInterfaceCount As Integer
        Public ModelID As Integer
        Sub New(SetModelID As Integer)

            GroupInterfaceCount = -1
            ModelID = SetModelID

        End Sub
        Public Sub ShowGroupInterface(ModelID As Integer, GSID As Integer, ShowStyle As String, InterfaceType As String, ParentLoadingForm As FileInstanceInterface, Optional ByVal LinkTag As ElementInterfaceLinkTag = Nothing)

            Dim IntCheck As GroupInterfaceObject

            If GroupInterfaceCount > -1 Then

                For Each IntCheck In GroupInterfaces

                    If IntCheck.GSID = GSID And IntCheck.ModelId = ModelID Then

                        IntCheck.RenderedForm.Visible = True
                        IntCheck.RenderedForm.Activate()
                        IntCheck.RenderedForm.BringToFront()

                        If LinkTag IsNot Nothing Then

                            Dim FindCSID As Integer = GetCSID(ModelID, GSID, LinkTag.LinkData)
                            IntCheck.RenderedForm.ShowInterface(ModelID, FindCSID,, , LinkTag)

                        End If

                        If ShowStyle = "Normal" Then

                            If IntCheck.RenderedForm.IsFormMinimized Then IntCheck.RenderedForm.WindowState = FormWindowState.Normal

                            Exit Sub

                        ElseIf ShowStyle = "Maximised" Then

                            If IntCheck.RenderedForm.IsFormMinimized Then IntCheck.RenderedForm.WindowState = FormWindowState.Maximized

                            Exit Sub

                        Else

                            SetFormToHalfScreen(IntCheck.RenderedForm, ShowStyle)

                            Exit Sub

                        End If



                    End If

                Next

            End If

            GroupInterfaceCount += 1

            ReDim Preserve GroupInterfaces(GroupInterfaceCount)
            GroupInterfaces(GroupInterfaceCount) = New GroupInterfaceObject With {.ModelId = ModelID, .GSID = GSID, .ID = GroupInterfaceCount}
            GroupInterfaces(GroupInterfaceCount).Initialise(ModelID, GSID, InterfaceType)

            If LinkTag IsNot Nothing Then

                Dim FindCSID As Integer = GetCSID(ModelID, GSID, LinkTag.LinkData)
                GroupInterfaces(GroupInterfaceCount).RenderedForm.ShowInterface(ModelID, FindCSID,, , LinkTag)

            End If

            If ShowStyle = "Normal" Then

                GroupInterfaces(GroupInterfaceCount).RenderedForm.WindowState = FormWindowState.Normal

                Exit Sub

            ElseIf ShowStyle = "Maximised" Then

                GroupInterfaces(GroupInterfaceCount).RenderedForm.WindowState = FormWindowState.Maximized

                Exit Sub

            Else

                SetFormToHalfScreen(GroupInterfaces(GroupInterfaceCount).RenderedForm, ShowStyle)

                Exit Sub

            End If

        End Sub
        Public Sub SetFormToHalfScreen(SentForm As Form, WhichSide As String)

            SentForm.WindowState = FormWindowState.Normal

            If WhichSide = "Left" Then

                SentForm.Left = 0
                SentForm.Width = Screen.PrimaryScreen.Bounds.Width / 2
                SentForm.Height = Screen.PrimaryScreen.Bounds.Height
                SentForm.Top = 0

            Else

                SentForm.Left = Screen.PrimaryScreen.Bounds.Width / 2
                SentForm.Width = Screen.PrimaryScreen.Bounds.Width / 2
                SentForm.Height = Screen.PrimaryScreen.Bounds.Height
                SentForm.Top = 0

            End If



        End Sub
        Sub CloseInterfaces()

            Dim IntCheck As GroupInterfaceObject

            If GroupInterfaceCount > -1 Then

                For Each IntCheck In GroupInterfaces

                    If IntCheck.RenderedForm IsNot Nothing Then

                        IntCheck.RenderedForm.Close()
                        IntCheck.RenderedForm.Dispose()
                        IntCheck.RenderedForm = Nothing

                    End If

                Next

            End If

            GroupInterfaceCount = -1
            ReDim GroupInterfaces(-1)

        End Sub

    End Class

    Public Class GroupInterfaceObject

        Public ModelId As Integer
        Public ID As Integer
        Public GSID As Integer
        Public RenderedForm As GroupInterfaceTemplate
        Public Sub Initialise(SetModelID As Integer, GSID As Integer, SetInterfaceMode As String)

            ModelId = SetModelID
            RenderedForm = New GroupInterfaceTemplate(SetModelID, GSID, SetInterfaceMode)
            RenderedForm.Show()

        End Sub

    End Class
    Public Class AbovoInterfaceTag

        Public TargetID As Integer
        Public SpecialItem As Boolean
        Public SpecialItemData As String

    End Class
    Public Class ActionToken

        Public ActionType As String
        Public ActionDescription As String
        Public ActionNR As String
        Public ActionStrData1 As String
        Public ActionStrData2 As String
        Public ActionStrData3 As String
        Public ActionNumber1 As Integer
        Public ActionNumber2 As Integer
        Public ActionNumber3 As Integer

    End Class

    Public Class MappedTableColumnTag

        Public ColDef As String
        Public Index As Integer
        Public ColSetID As Integer
        Public IsVisible As Boolean = True

    End Class
    Public Class BandTag

        Public ID As Integer
        Public HasActions As Boolean = False
        Public HighLightColour As Color
        Public ButtonObjectState As ObjectState
        Public BandEditDescription As String
        Public ActionToken As ActionToken
        Public ActionMethod As String
        Public ActionNR As String
        Public ActionDescription As String
        Public DoBorder As Boolean = False

    End Class

End Namespace
