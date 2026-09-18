Imports DevExpress.XtraEditors

Imports System.Configuration
Imports System.Collections.Generic
Imports System.Reflection
Imports System.Runtime.CompilerServices

Namespace Abovo
    Public Class FontManager

        Private Shared ApplicationSettingsLoaded As Boolean
        Private Shared BaseApplicationFont As Font
        Public Shared DefaultFont As New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Public Shared DefaultFontBold As New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Public Shared DefaultFontSmaller As New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Public Shared DefaultFontLarger As New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Public Shared DefaultFontLargest As New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

        Public Shared Sub InitialiseApplication()

            If ApplicationSettingsLoaded Then Return

            PresentationScaleManager.Initialise()

            'This must run before the first form is constructed. It loads the
            'DevExpress Per-Monitor V2 and default-font settings from App.config.
            WindowsFormsSettings.LoadApplicationSettings()

            Dim ApplicationFont As Font = WindowsFormsSettings.DefaultFont
            If ApplicationFont Is Nothing Then ApplicationFont = SystemFonts.MessageBoxFont
            BaseApplicationFont = New Font(ApplicationFont, FontStyle.Regular)

            Dim UserScale As Single = PresentationScaleManager.UserScale
            DefaultFont = New Font(ApplicationFont.FontFamily, ApplicationFont.SizeInPoints * UserScale, FontStyle.Regular)
            DefaultFontBold = New Font(ApplicationFont.FontFamily, ApplicationFont.SizeInPoints * UserScale, FontStyle.Bold)
            DefaultFontSmaller = New Font(
                ApplicationFont.FontFamily,
                Math.Max(7.0F, ApplicationFont.SizeInPoints - 1.0F) * UserScale,
                FontStyle.Regular)
            DefaultFontLarger = New Font(
                ApplicationFont.FontFamily,
                (ApplicationFont.SizeInPoints + 1.0F) * UserScale,
                FontStyle.Bold)
            DefaultFontLargest = New Font(
                ApplicationFont.FontFamily,
                (ApplicationFont.SizeInPoints + 2.0F) * UserScale,
                FontStyle.Bold)

            WindowsFormsSettings.DefaultFont = DefaultFont
            WindowsFormsSettings.DefaultMenuFont = DefaultFont
            ApplicationSettingsLoaded = True

        End Sub

        Friend Shared Sub ApplyUserScaleDefaults()
            If BaseApplicationFont Is Nothing Then Return

            Dim UserScale As Single = PresentationScaleManager.UserScale
            DefaultFont = New Font(BaseApplicationFont.FontFamily, BaseApplicationFont.SizeInPoints * UserScale, FontStyle.Regular)
            DefaultFontBold = New Font(BaseApplicationFont.FontFamily, BaseApplicationFont.SizeInPoints * UserScale, FontStyle.Bold)
            DefaultFontSmaller = New Font(
                BaseApplicationFont.FontFamily,
                Math.Max(7.0F, BaseApplicationFont.SizeInPoints - 1.0F) * UserScale,
                FontStyle.Regular)
            DefaultFontLarger = New Font(
                BaseApplicationFont.FontFamily,
                (BaseApplicationFont.SizeInPoints + 1.0F) * UserScale,
                FontStyle.Bold)
            DefaultFontLargest = New Font(
                BaseApplicationFont.FontFamily,
                (BaseApplicationFont.SizeInPoints + 2.0F) * UserScale,
                FontStyle.Bold)

            WindowsFormsSettings.DefaultFont = DefaultFont
            WindowsFormsSettings.DefaultMenuFont = DefaultFont
        End Sub

        Public Shared Sub Initialise()

            InitialiseApplication()

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

    Friend NotInheritable Class PresentationScaleSettings
        Inherits ApplicationSettingsBase

        Private Shared ReadOnly InstanceValue As PresentationScaleSettings =
            CType(Synchronized(New PresentationScaleSettings()), PresentationScaleSettings)

        Public Shared ReadOnly Property [Default] As PresentationScaleSettings
            Get
                Return InstanceValue
            End Get
        End Property

        <UserScopedSetting(), DefaultSettingValue("100")>
        Public Property InterfaceScalePercent As Integer
            Get
                Return CInt(Me(NameOf(InterfaceScalePercent)))
            End Get
            Set(value As Integer)
                Me(NameOf(InterfaceScalePercent)) = value
            End Set
        End Property
    End Class

    Public NotInheritable Class PresentationScaleChangedEventArgs
        Inherits EventArgs

        Public Sub New(ByVal OldScale As Single, ByVal NewScale As Single)
            Me.OldScale = OldScale
            Me.NewScale = NewScale
        End Sub

        Public ReadOnly Property OldScale As Single
        Public ReadOnly Property NewScale As Single

        Public ReadOnly Property Ratio As Single
            Get
                If OldScale <= 0.0F Then Return 1.0F
                Return NewScale / OldScale
            End Get
        End Property
    End Class

    Public NotInheritable Class PresentationScaleManager

        Public Const MinimumPercent As Integer = 75
        Public Const MaximumPercent As Integer = 200
        Public Const DefaultPercent As Integer = 100

        Private Shared Initialised As Boolean
        Private Shared CurrentPercent As Integer = DefaultPercent
        Private Shared ReadOnly InitialisedForms As New ConditionalWeakTable(Of Form, Object)()
        Private Shared ReadOnly InitialisedControls As New ConditionalWeakTable(Of Control, Object)()

        Public Shared Event ScaleChanged As EventHandler(Of PresentationScaleChangedEventArgs)

        Private Sub New()
        End Sub

        Public Shared Sub Initialise()
            If Initialised Then Return

            Dim StoredPercent As Integer = DefaultPercent
            Try
                StoredPercent = PresentationScaleSettings.Default.InterfaceScalePercent
            Catch
                StoredPercent = DefaultPercent
            End Try

            CurrentPercent = ClampPercent(StoredPercent)
            AddHandler Application.Idle, AddressOf ApplyScaleToNewForms
            Initialised = True
        End Sub

        Public Shared ReadOnly Property UserScale As Single
            Get
                Initialise()
                Return CSng(CurrentPercent) / 100.0F
            End Get
        End Property

        Public Shared ReadOnly Property InterfaceScalePercent As Integer
            Get
                Initialise()
                Return CurrentPercent
            End Get
        End Property

        Public Shared Sub SetInterfaceScale(ByVal Percent As Integer,
                                            Optional ByVal Persist As Boolean = True)
            Initialise()

            Dim NewPercent As Integer = ClampPercent(Percent)
            If NewPercent = CurrentPercent Then Return

            Dim OldScale As Single = UserScale
            CurrentPercent = NewPercent

            If Persist Then
                Try
                    PresentationScaleSettings.Default.InterfaceScalePercent = NewPercent
                    PresentationScaleSettings.Default.Save()
                Catch ex As Exception
                    Debug.WriteLine("Unable to save interface scale: " & ex.Message)
                End Try
            End If

            Dim Arguments As New PresentationScaleChangedEventArgs(OldScale, UserScale)
            FontManager.ApplyUserScaleDefaults()
            RaiseEvent ScaleChanged(Nothing, Arguments)
            ApplyToOpenForms(Arguments)
        End Sub

        Public Shared Function Scale(ByVal Value As Integer) As Integer
            Return Math.Max(1, CInt(Math.Round(Value * UserScale)))
        End Function

        Public Shared Function Scale(ByVal Value As Single) As Single
            Return Value * UserScale
        End Function

        Public Shared Sub ShowOptions(ByVal Owner As IWin32Window)
            Using Options As New ApplicationOptionsForm()
                Options.ShowDialog(Owner)
            End Using
        End Sub

        Private Shared Function ClampPercent(ByVal Value As Integer) As Integer
            Return Math.Max(MinimumPercent, Math.Min(MaximumPercent, Value))
        End Function

        Private Shared Sub ApplyToOpenForms(ByVal Arguments As PresentationScaleChangedEventArgs)
            Dim OpenForms As New List(Of Form)()
            For Each OpenForm As Form In Application.OpenForms
                OpenForms.Add(OpenForm)
            Next

            For Each OpenForm As Form In OpenForms
                Dim TargetForm As Form = OpenForm
                RunOnFormThread(
                    TargetForm,
                    Sub()
                        ScaleControlTree(TargetForm, Arguments.Ratio, True)
                        InvokeScaleHook(TargetForm)
                        TargetForm.PerformLayout()
                        TargetForm.Invalidate(True)
                    End Sub)
            Next
        End Sub

        Private Shared Sub ApplyScaleToNewForms(ByVal Sender As Object, ByVal e As EventArgs)
            Dim OpenForms As New List(Of Form)()
            For Each OpenForm As Form In Application.OpenForms
                OpenForms.Add(OpenForm)
            Next

            For Each OpenForm As Form In OpenForms
                Dim TargetForm As Form = OpenForm
                RunOnFormThread(
                    TargetForm,
                    Sub()
                        Dim Marker As Object = Nothing
                        If InitialisedForms.TryGetValue(TargetForm, Marker) Then Return
                        InitialisedForms.Add(TargetForm, New Object())

                        ApplyInitialScaleToControlTree(TargetForm)
                        InvokeScaleHook(TargetForm)
                        TargetForm.PerformLayout()
                        TargetForm.Invalidate(True)
                    End Sub)
            Next
        End Sub

        Private Shared Sub RunOnFormThread(ByVal TargetForm As Form,
                                           ByVal Action As MethodInvoker)
            If TargetForm Is Nothing OrElse Action Is Nothing Then Return

            Try
                If TargetForm.IsDisposed Then Return
                If TargetForm.InvokeRequired Then
                    If Not TargetForm.IsHandleCreated Then Return
                    TargetForm.BeginInvoke(Action)
                Else
                    Action.Invoke()
                End If
            Catch ex As ObjectDisposedException
                'The form closed while a scale refresh was being queued.
            Catch ex As InvalidOperationException
                Debug.WriteLine(
                    "Unable to marshal interface scale to " & TargetForm.GetType().Name &
                    ": " & ex.Message)
            End Try
        End Sub

        Private Shared Sub ApplyInitialScaleToControlTree(ByVal Target As Control)
            If Target Is Nothing OrElse Target.IsDisposed Then Return

            Dim Marker As Object = Nothing
            If Not InitialisedControls.TryGetValue(Target, Marker) Then
                InitialisedControls.Add(Target, New Object())
                AddHandler Target.ControlAdded, AddressOf ScaledControlAdded
                ScaleKnownAppearances(Target, UserScale)

                Dim Accordion As DevExpress.XtraBars.Navigation.AccordionControl =
                    TryCast(Target, DevExpress.XtraBars.Navigation.AccordionControl)
                If Accordion IsNot Nothing Then
                    Accordion.BeginUpdate()
                    Try
                        For Each Element As DevExpress.XtraBars.Navigation.AccordionControlElement In Accordion.Elements
                            ScaleAccordionElement(Element, UserScale)
                        Next
                    Finally
                        Accordion.EndUpdate()
                    End Try
                End If

                Dim Browser As WebBrowser = TryCast(Target, WebBrowser)
                If Browser IsNot Nothing Then
                    AddHandler Browser.DocumentCompleted,
                        Sub(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
                            ApplyBrowserScale(TryCast(sender, WebBrowser))
                        End Sub
                    ApplyBrowserScale(Browser)
                End If
            End If

            For Each Child As Control In Target.Controls
                ApplyInitialScaleToControlTree(Child)
            Next
        End Sub

        Private Shared Sub ScaledControlAdded(ByVal Sender As Object,
                                              ByVal e As ControlEventArgs)
            If e Is Nothing OrElse e.Control Is Nothing Then Return
            ApplyInitialScaleToControlTree(e.Control)
            e.Control.PerformLayout()
            e.Control.Invalidate(True)
        End Sub

        Private Shared Sub ScaleControlTree(ByVal Parent As Control,
                                            ByVal Ratio As Single,
                                            ByVal ScaleControlFont As Boolean)
            If Parent Is Nothing OrElse Parent.IsDisposed Then Return

            If ScaleControlFont AndAlso Parent.Font IsNot Nothing AndAlso Ratio > 0.0F Then
                Parent.Font = New Font(
                    Parent.Font.FontFamily,
                    Math.Max(5.0F, Parent.Font.SizeInPoints * Ratio),
                    Parent.Font.Style,
                    GraphicsUnit.Point)
            End If

            ScaleKnownAppearances(Parent, Ratio)

            Dim Browser As WebBrowser = TryCast(Parent, WebBrowser)
            If Browser IsNot Nothing Then ApplyBrowserScale(Browser)

            Dim Accordion As DevExpress.XtraBars.Navigation.AccordionControl =
                TryCast(Parent, DevExpress.XtraBars.Navigation.AccordionControl)
            If Accordion IsNot Nothing Then
                For Each Element As DevExpress.XtraBars.Navigation.AccordionControlElement In Accordion.Elements
                    ScaleAccordionElement(Element, Ratio)
                Next
            End If

            For Each Child As Control In Parent.Controls
                ScaleControlTree(Child, Ratio, ScaleControlFont)
            Next
        End Sub

        Private Shared Sub ApplyBrowserScale(ByVal Browser As WebBrowser)
            If Browser Is Nothing OrElse Browser.IsDisposed OrElse Browser.Document Is Nothing OrElse
               Browser.Document.Body Is Nothing Then Return
            Browser.Document.Body.Style =
                "zoom:" & InterfaceScalePercent.ToString(Globalization.CultureInfo.InvariantCulture) & "%;"
        End Sub

        Private Shared Sub ScaleKnownAppearances(ByVal Target As Control, ByVal Ratio As Single)
            Dim Label As LabelControl = TryCast(Target, LabelControl)
            If Label IsNot Nothing Then ScaleAppearanceFont(Label.Appearance, Ratio)

            Dim Group As GroupControl = TryCast(Target, GroupControl)
            If Group IsNot Nothing Then ScaleAppearanceFont(Group.AppearanceCaption, Ratio)

            Dim Tabs As DevExpress.XtraTab.XtraTabControl =
                TryCast(Target, DevExpress.XtraTab.XtraTabControl)
            If Tabs IsNot Nothing Then
                ScaleAppearanceFont(Tabs.Appearance, Ratio)
                ScaleAppearanceFont(Tabs.AppearancePage.Header, Ratio)
                ScaleAppearanceFont(Tabs.AppearancePage.HeaderActive, Ratio)
                ScaleAppearanceFont(Tabs.AppearancePage.HeaderHotTracked, Ratio)
            End If

            Dim ButtonPanel As DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel =
                TryCast(Target, DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel)
            If ButtonPanel IsNot Nothing Then
                ScaleAppearanceFont(ButtonPanel.AppearanceButton.Normal, Ratio)
                ScaleAppearanceFont(ButtonPanel.AppearanceButton.Hovered, Ratio)
                ScaleAppearanceFont(ButtonPanel.AppearanceButton.Pressed, Ratio)
            End If

            Dim Grid As DevExpress.XtraGrid.GridControl =
                TryCast(Target, DevExpress.XtraGrid.GridControl)
            If Grid IsNot Nothing Then
                For Each View As DevExpress.XtraGrid.Views.Base.BaseView In Grid.ViewCollection
                    Dim TableView As DevExpress.XtraGrid.Views.Grid.GridView =
                        TryCast(View, DevExpress.XtraGrid.Views.Grid.GridView)
                    If TableView Is Nothing Then Continue For

                    ScaleAppearanceFont(TableView.Appearance.Row, Ratio)
                    ScaleAppearanceFont(TableView.Appearance.HeaderPanel, Ratio)
                    ScaleAppearanceFont(TableView.Appearance.GroupRow, Ratio)
                    ScaleAppearanceFont(TableView.Appearance.GroupFooter, Ratio)
                    ScaleAppearanceFont(TableView.Appearance.FooterPanel, Ratio)
                    If TableView.RowHeight > 0 Then TableView.RowHeight = ScaleMetric(TableView.RowHeight, Ratio)
                    If TableView.ColumnPanelRowHeight > 0 Then
                        TableView.ColumnPanelRowHeight = ScaleMetric(TableView.ColumnPanelRowHeight, Ratio)
                    End If

                    Dim BandedView As DevExpress.XtraGrid.Views.BandedGrid.BandedGridView =
                        TryCast(TableView, DevExpress.XtraGrid.Views.BandedGrid.BandedGridView)
                    If BandedView IsNot Nothing Then
                        ScaleAppearanceFont(BandedView.Appearance.BandPanel, Ratio)
                        If BandedView.BandPanelRowHeight > 0 Then
                            BandedView.BandPanelRowHeight = ScaleMetric(BandedView.BandPanelRowHeight, Ratio)
                        End If
                    End If
                Next
            End If

            Dim VerticalGrid As DevExpress.XtraVerticalGrid.VGridControl =
                TryCast(Target, DevExpress.XtraVerticalGrid.VGridControl)
            If VerticalGrid IsNot Nothing Then
                ScaleAppearanceFont(VerticalGrid.Appearance.RecordValue, Ratio)
                ScaleAppearanceFont(VerticalGrid.Appearance.RowHeaderPanel, Ratio)
                ScaleAppearanceFont(VerticalGrid.Appearance.Category, Ratio)
                ScaleAppearanceFont(VerticalGrid.Appearance.FocusedCell, Ratio)
                ScaleAppearanceFont(VerticalGrid.Appearance.FocusedRecord, Ratio)
            End If

            Dim Chart As DevExpress.XtraCharts.ChartControl =
                TryCast(Target, DevExpress.XtraCharts.ChartControl)
            If Chart IsNot Nothing Then
                For Each Title As DevExpress.XtraCharts.ChartTitle In Chart.Titles
                    If Title.DXFont Is Nothing Then Continue For
                    Title.DXFont = New DevExpress.Drawing.DXFont(
                        Title.DXFont.Name,
                        Math.Max(5.0F, CSng(Title.DXFont.Size * Ratio)),
                        Title.DXFont.Style)
                Next
            End If
        End Sub

        Private Shared Function ScaleMetric(ByVal Value As Integer, ByVal Ratio As Single) As Integer
            Return Math.Max(1, CInt(Math.Round(Value * Ratio)))
        End Function

        Private Shared Sub ScaleAccordionElement(
            ByVal Element As DevExpress.XtraBars.Navigation.AccordionControlElement,
            ByVal Ratio As Single)

            ScaleAppearanceFont(Element.Appearance.Default, Ratio)
            ScaleAppearanceFont(Element.Appearance.Normal, Ratio)
            ScaleAppearanceFont(Element.Appearance.Hovered, Ratio)
            ScaleAppearanceFont(Element.Appearance.Pressed, Ratio)
            ScaleAppearanceFont(Element.Appearance.Disabled, Ratio)

            For Each Child As DevExpress.XtraBars.Navigation.AccordionControlElement In Element.Elements
                ScaleAccordionElement(Child, Ratio)
            Next
        End Sub

        Private Shared Sub ScaleAppearanceFont(ByVal Appearance As DevExpress.Utils.AppearanceObject,
                                               ByVal Ratio As Single)
            If Appearance Is Nothing OrElse Appearance.Font Is Nothing OrElse
               Not Appearance.Options.UseFont Then Return
            Appearance.Font = New Font(
                Appearance.Font.FontFamily,
                Math.Max(5.0F, Appearance.Font.SizeInPoints * Ratio),
                Appearance.Font.Style,
                GraphicsUnit.Point)
            Appearance.Options.UseFont = True
        End Sub

        Private Shared Function InvokeScaleHook(ByVal Target As Object) As Boolean
            Dim TargetType As Type = Target.GetType()
            Dim Hook As MethodInfo = TargetType.GetMethod(
                "ApplyPresentationScale",
                BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                Nothing,
                Type.EmptyTypes,
                Nothing)

            If Hook Is Nothing Then
                Hook = TargetType.GetMethod(
                    "ResizeFonts",
                    BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                    Nothing,
                    Type.EmptyTypes,
                    Nothing)
            End If

            If Hook Is Nothing Then Return False

            Try
                Hook.Invoke(Target, Nothing)
                Return True
            Catch ex As Exception
                Debug.WriteLine("Unable to apply interface scale to " & TargetType.Name & ": " & ex.Message)
                Return False
            End Try
        End Function
    End Class

End Namespace
