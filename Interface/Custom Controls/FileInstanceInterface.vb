Imports Abovo
Imports Abovo.AbovoAppCls
Imports Abovo.FileManager
Imports Abovo.PresentationManager
Imports DevExpress
Imports DevExpress.Utils
Imports DevExpress.XtraBars.Docking2010
Public Class FileInstanceInterface

    Inherits System.Windows.Forms.UserControl

    Public BPModelID As Integer

    Private ScaleUnits As Single
    Private MyFileInfos As System.IO.FileInfo
    Public MyFilePath As String
    Public MyCompanyName As String

    Public StressTester As StressTest
    Public FFRer As FFRForm
    Private STInit As Boolean
    Private FFRInit As Boolean
    Private MyChildInterfaces() As GroupInterfaceTemplate
    Public Property BPModelInstance As Integer

        Get

            Return BPModelID

        End Get

        Set(value As Integer)

            BPModelID = value

        End Set

    End Property

    Public ReadOnly Property ModelID As Integer
        Get
            Return BPModelID
        End Get
    End Property
    Public Sub New(ModelID As Integer)

        ' This call is required by the designer.
        InitializeComponent()
        BPModelID = ModelID
        ScaleUnits = 5
        'PopulateFileInfo()

        ' Add any initialization after the InitializeComponent() call.
        STInit = False
        FFRInit = False
        FileManager.RegisterModelInterface(BPModelID, Me)
        ConfigureModelActions()

    End Sub
    Private Sub ConfigureModelActions()

        If ExcelModels Is Nothing OrElse
           BPModelID < 0 OrElse
           BPModelID >= ExcelModels.Length OrElse
           ExcelModels(BPModelID) Is Nothing Then Return

        Dim IsBusinessPlan As Boolean =
            ExcelModels(BPModelID).Profile IsNot Nothing AndAlso
            String.Equals(
                ExcelModels(BPModelID).Profile.ModelType,
                "AbovoBP",
                StringComparison.OrdinalIgnoreCase)

        For Each Item As DevExpress.XtraEditors.ButtonPanel.IBaseButton In
            WindowsUIButtonPanelBPActions.Buttons

            Dim ActionButton As WindowsUIButton = TryCast(Item, WindowsUIButton)
            If ActionButton Is Nothing OrElse ActionButton.Tag Is Nothing Then Continue For

            Dim ActionTag As String = ActionButton.Tag.ToString()

            Select Case ActionTag
                Case "GoAssumpt"
                    ActionButton.Visible =
                        GetGroupID(BPModelID, "Assumptions") >= 0
                Case "GoWorkings"
                    ActionButton.Visible =
                        GetGroupID(BPModelID, "Workings") >= 0
                Case "GoOutputs"
                    ActionButton.Visible =
                        GetGroupID(BPModelID, "Outputs") >= 0
                Case "GoFFR", "GoData", "StressTest"
                    ActionButton.Visible = IsBusinessPlan
            End Select
        Next

    End Sub
    Public Sub ProcessBPInstance(ByVal BPModelID As Integer)

        Me.BPModelID = BPModelID

        ' Set the title of the interface


    End Sub
    Public Sub ShowInterface(InterfaceName As String, Optional ByVal LinkTag As ElementInterfaceLinkTag = Nothing)

        Dim GroupId As Integer = GetGroupID(BPModelID, InterfaceName)

        If GroupId < 0 Then
            MessageBox.Show(
                Me,
                "This model does not define an '" & InterfaceName & "' interface.",
                "Interface unavailable",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
            Return
        End If

        ExcelModels(BPModelID).WBInterface.ShowGroupInterface(BPModelID, GroupId, "Maximised", InterfaceName, Me)

    End Sub


    Private Sub WindowsUIButtonPanelBPActions_ButtonClick(sender As Object, e As DevExpress.XtraBars.Docking2010.ButtonEventArgs) Handles WindowsUIButtonPanelBPActions.ButtonClick

        Dim ButSender As WindowsUIButton = TryCast(e.Button, DevExpress.XtraBars.Docking2010.WindowsUIButton)

        If ButSender Is Nothing Then

            Return

        End If

        Dim tag As String = ButSender.Tag.ToString()

        Select Case tag

            Case "GoAssumpt"
                ' OpenAssumptionsInterface

                ShowInterface("Assumptions")

            Case "GoWorkings"

                ShowInterface("Workings")

            Case "GoOutputs"

                ShowInterface("Outputs")

            Case "GoOther"



            Case "GoFFR"

                If Not FFRInit Then
                    Me.Cursor = Cursors.WaitCursor
                    FFRer = New FFRForm(BPModelID)
                    FFRer.Show()
                    FFRInit = True
                    Me.Cursor = Cursors.Default
                Else

                    FFRer.Show()

                End If




            Case "StressTest"

                If Not STInit Then

                    StressTester = New StressTest(BPModelID)
                    StressTester.SetActive()
                    StressTester.ShowDialog()

                    STInit = True

                Else

                    StressTester.SetActive()
                    StressTester.ShowDialog()

                End If

        End Select

    End Sub
    Public Sub PopulateFileInfo()

        ScaleUnits = CInt(Screen.PrimaryScreen.Bounds.Width / 300)

        Dim StrFileDescription As String

        MyFilePath = ExcelModels(BPModelID).FileName
        MyCompanyName = ExcelModels(BPModelID).WBStructure.CompanyName

        Dim ModelDescription As String =
            If(ExcelModels(BPModelID).Profile Is Nothing,
               "Abovo model",
               ExcelModels(BPModelID).Profile.DisplayName)

        StrFileDescription = "<html><body><p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits * 1.5) & "px'>Model type: " & ModelDescription & "<br/>"
        StrFileDescription += "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits * 1.5) & "px'>Model name: " & ExcelModels(BPModelID).WBStructure.CompanyName & "<br/>"
        StrFileDescription += "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits * 1.5) & "px'>Start Date: " & ExcelModels(BPModelID).WBStructure.StartDate & " (<a href='editbpdate'>edit</a>)<br/>"
        StrFileDescription += "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits * 1.3) & "px'>File Name: " & ExcelModels(BPModelID).FileName & "<br/>"
        StrFileDescription += "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits * 1.3) & "px'>Opened: " & Now().ToString & "<br/>"

        StrFileDescription += "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits * 1.3) & "px'>Created: " & ExcelModels(BPModelID).FileInfo.CreationTime & "<br/>"
        StrFileDescription += "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits * 1.3) & "px'>Last Previous Access: " & ExcelModels(BPModelID).FileInfo.LastAccessTime & "<br/>"
        StrFileDescription += "<p style ='font-family:verdana' style='font-size:" & CInt(ScaleUnits * 1.3) & "px'>Size: " & Format((ExcelModels(BPModelID).FileInfo.Length / 1000000), "###.##") & "Mb<br/>"
        StrFileDescription += "</body></html>"

        WebBrowserBPInfo.DocumentText = StrFileDescription

    End Sub
    Private Sub WindowsUIButtonPanelSaveClose_ButtonClick(sender As Object, e As ButtonEventArgs) Handles WindowsUIButtonPanelSaveClose.ButtonClick

        Dim ButSender As WindowsUIButton = TryCast(e.Button, DevExpress.XtraBars.Docking2010.WindowsUIButton)
        If ButSender Is Nothing Then
            Return
        End If
        Dim tag As String = ButSender.Tag.ToString()

        Select Case tag

            Case "SaveBP"

                Me.Cursor = Cursors.WaitCursor
                Try
                    ExcelModels(BPModelID).SaveFile()
                Finally
                    Me.Cursor = Cursors.Default
                End Try

            Case "SaveBPAs"

                'Me.Cursor = Cursors.WaitCursor
                ExcelModels(BPModelID).SaveFileAs()
                'Me.Cursor = Cursors.Default

            Case "CloseBP"

                ' Close the model and dispose of the interface
                If ExcelModels(BPModelID).CommitToCloseModel.StringReturn = "Proceed" Then

                    FileManager.CloseModel(BPModelID)
                    FormMainScreen.RemoveModel(BPModelID)
                    Me.Dispose()

                End If

            Case "Spreadsheet"

                ' Close the model and dispose of the interface
                ExcelModels(BPModelID).ShowSpreadsheet()


        End Select

    End Sub
    Public Function SaveFileAs() As Boolean
        Me.Cursor = Cursors.WaitCursor
        Try
            Return ExcelModels(BPModelID).SaveFileAs()
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Function
    Sub SetScale()
        Dim ScaleFactor As Single
        ScaleFactor = Me.Width / 700


        Me.WindowsUIButtonPanelBPActions.Font = GetFont("Small", ScaleFactor)


        WindowsUIButtonPanelBPActions.AppearanceButton.Normal.Font = GetFont("Small", ScaleFactor)
        WindowsUIButtonPanelBPActions.AppearanceButton.Hovered.Font = GetFont("Small", ScaleFactor)
        WindowsUIButtonPanelBPActions.AppearanceButton.Pressed.Font = GetFont("Small", ScaleFactor)


        WindowsUIButtonPanelSaveClose.AppearanceButton.Normal.Font = GetFont("Small", ScaleFactor)
        WindowsUIButtonPanelSaveClose.AppearanceButton.Hovered.Font = GetFont("Small", ScaleFactor)
        WindowsUIButtonPanelSaveClose.AppearanceButton.Pressed.Font = GetFont("Small", ScaleFactor)
        Me.GroupBoxFileActions.Font = GetFont("Small", ScaleFactor)
    End Sub
End Class
